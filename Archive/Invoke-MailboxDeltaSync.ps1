<#
.SYNOPSIS
    Incremental delta sync — keeps the Mailboxes table aligned with EXO.

.DESCRIPTION
    Scheduled every 15 minutes. Uses a WhenChangedUTC timestamp filter
    against Exchange Online — there is no Graph delta for mailboxes.

    Steps:
      1. Reads mailboxes_delta_timestamp from DeltaTokens
      2. Subtracts DeltaOverlapMinutes to set the query window start
      3. Fetches active + soft-deleted mailboxes changed in that window
      4. Soft-deletes removed mailboxes, inserts new, updates changed
      5. Saves new timestamp ONLY after all processing succeeds
         (crash-safe — next run reprocesses rather than skips)

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Scheduled Task:
      Trigger : Daily, repeat every 15 minutes indefinitely
      Program : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-MailboxDeltaSync.ps1"
      Settings: Do not start a new instance if already running.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-MailboxDeltaSync"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ──────────────────────────────────────────────────────────────
# Bootstrap
# ──────────────────────────────────────────────────────────────

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")  -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")     -Force
Import-Module (Join-Path $sharedPath "ExoHelpers.psm1")     -Force

try {
    $Config = Import-SyncConfig -Path $ConfigPath
    Initialize-Logging -Config $Config -ProcessName $scriptName
}
catch {
    Write-Host "FATAL BOOTSTRAP ERROR in $scriptName" -ForegroundColor Red
    Write-Host "Message : $($_.Exception.Message)"    -ForegroundColor White
    exit 1
}

Write-LogSection "Invoke-MailboxDeltaSync"
Write-LogInfo "Config  : $ConfigPath"
Write-LogInfo "EXO Org : $($Config.ExchangeOnline.Organisation)"

# ──────────────────────────────────────────────────────────────
# Authentication
# ──────────────────────────────────────────────────────────────

Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Initialize-SqlContext -Config $Config
    Initialize-ExoContext -Config $Config
    Connect-ExoSession    -Config $Config
    Write-LogInfo "Azure and Exchange Online connected"
}
catch {
    Write-LogError "Authentication failed" -ErrorRecord $_
    Disconnect-ExoSession
    Close-Logging -Status "Failed"
    exit 1
}

# ──────────────────────────────────────────────────────────────
# Run tracking
# ──────────────────────────────────────────────────────────────

$runId = [guid]::NewGuid()
Invoke-SqlNonQuery -Query @"
INSERT INTO dbo.SyncLog (RunId, FunctionName, StartedAt, Status)
VALUES (@RunId, @Function, SYSUTCDATETIME(), 'Running')
"@ -Parameters @{ '@RunId' = $runId; '@Function' = $scriptName } | Out-Null

$inserted     = 0
$updated      = 0
$softDeleted  = 0
$processed    = 0
$errors       = 0
$newTimestamp = $null
$overallSw    = [System.Diagnostics.Stopwatch]::StartNew()

Write-LogInfo "Run ID: $runId"

# ──────────────────────────────────────────────────────────────
# EXO mailbox properties
# ──────────────────────────────────────────────────────────────

$mailboxProps = @(
    'ExchangeGuid', 'ExternalDirectoryObjectId', 'PrimarySmtpAddress',
    'UserPrincipalName', 'DisplayName', 'Alias', 'RecipientTypeDetails',
    'HiddenFromAddressListsEnabled', 'LitigationHoldEnabled', 'ArchiveStatus',
    'ForwardingAddress', 'ForwardingSmtpAddress', 'GrantSendOnBehalfTo',
    'IsDirSynced', 'WhenMailboxCreated', 'WhenChangedUTC'
)

$mailboxTypes = $Config.ExchangeOnline.MailboxTypes -split ',' |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ }

# ──────────────────────────────────────────────────────────────
# Phase 1 — Determine sync window
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 1 — Determine sync window"

$storedTimestamp = Get-DeltaToken -TokenName "mailboxes_delta_timestamp"

if (-not $storedTimestamp) {
    $msg = "No active mailboxes_delta_timestamp. Run Invoke-MailboxBaselineLoad first."
    Write-LogError $msg
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = $msg } | Out-Null
    Disconnect-ExoSession
    Close-Logging -Status "Failed"
    exit 1
}

$overlapMinutes = [int]($Config.Sync.DeltaOverlapMinutes ?? 30)
$sinceUtc       = [datetime]::Parse($storedTimestamp).ToUniversalTime().AddMinutes(-$overlapMinutes)
$filterValue    = $sinceUtc.ToString('MM/dd/yyyy HH:mm:ss')

Write-LogInfo "Stored timestamp : $storedTimestamp"
Write-LogInfo "Overlap minutes  : $overlapMinutes"
Write-LogInfo "Query window from: $($sinceUtc.ToString('o')) UTC"

# ──────────────────────────────────────────────────────────────
# Phase 2 — Fetch changed mailboxes from EXO
# Checks both active and soft-deleted mailboxes in the window.
# Soft-deleted mailboxes appear in Exchange for 30 days before
# permanent removal — catching them here avoids waiting for the
# weekly baseline reconciliation pass.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 2 — EXO delta fetch"

$changedMailboxes = [System.Collections.Generic.List[PSObject]]::new()

try {
    $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()
    $filter  = "WhenChangedUTC -ge '$filterValue'"

    # Active mailboxes that changed in the window
    $active = Get-EXOMailbox `
        -Filter     $filter `
        -ResultSize Unlimited `
        -Properties $mailboxProps `
        -ErrorAction Stop |
        Where-Object { $_.RecipientTypeDetails -in $mailboxTypes }

    foreach ($m in $active) { $m | Add-Member -NotePropertyName '_IsDeleted' -NotePropertyValue $false -Force; $changedMailboxes.Add($m) }

    # Soft-deleted mailboxes that changed in the window
    $softDeletedFromEXO = Get-EXOMailbox `
        -SoftDeletedMailbox `
        -Filter     $filter `
        -ResultSize Unlimited `
        -Properties $mailboxProps `
        -ErrorAction SilentlyContinue |
        Where-Object { $_ -and $_.RecipientTypeDetails -in $mailboxTypes }

    foreach ($m in $softDeletedFromEXO) { $m | Add-Member -NotePropertyName '_IsDeleted' -NotePropertyValue $true -Force; $changedMailboxes.Add($m) }

    $phaseSw.Stop()
    Write-LogInfo "EXO returned $($active.Count) changed + $($softDeletedFromEXO.Count) soft-deleted in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s"
}
catch {
    Write-LogError "Phase 2 (EXO delta fetch) failed" -ErrorRecord $_
    Disconnect-ExoSession
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = "EXO fetch failed: $($_.Exception.Message)" } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

# ── No changes ────────────────────────────────────────────────
if ($changedMailboxes.Count -eq 0) {
    Write-LogInfo "No changes in window — advancing timestamp"
    $newTimestamp = (Get-Date).ToUniversalTime().ToString('o')
    Save-DeltaToken -TokenName "mailboxes_delta_timestamp" -TokenValue $newTimestamp
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Success',
    MailboxesProcessed=0, TokenAdvancedTo=@Token WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Token' = $newTimestamp } | Out-Null
    Write-LogSummary @{ "Status" = "Success (no changes)"; "Duration" = "$([Math]::Round($overallSw.Elapsed.TotalSeconds,1))s" }
    Disconnect-ExoSession
    Remove-OldLogFiles
    Close-Logging -Status "Success"
    exit 0
}

# ──────────────────────────────────────────────────────────────
# Phase 3 — Apply changes
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 3 — Apply changes"

$softDeleteSql = @"
UPDATE dbo.Mailboxes SET
    IsDeleted = 1, DeletedAt = SYSUTCDATETIME(),
    LastSyncedAt = SYSUTCDATETIME(), LastModifiedAt = SYSUTCDATETIME(),
    SyncSource = 'Delta', LastSyncRunId = @RunId
WHERE ExchangeGuid = @ExchangeGuid AND IsDeleted = 0
"@

$upsertSql = @"
MERGE dbo.Mailboxes AS target
USING (SELECT @ExchangeGuid AS ExchangeGuid) AS source
ON target.ExchangeGuid = source.ExchangeGuid
WHEN MATCHED THEN UPDATE SET
    UserId                  = COALESCE(@UserId,                  UserId),
    PrimarySmtpAddress      = COALESCE(@PrimarySmtp,             PrimarySmtpAddress),
    UserPrincipalName       = COALESCE(@UPN,                     UserPrincipalName),
    DisplayName             = COALESCE(@DisplayName,             DisplayName),
    Alias                   = COALESCE(@Alias,                   Alias),
    RecipientTypeDetails    = COALESCE(@RecipientTypeDetails,    RecipientTypeDetails),
    MailboxType             = COALESCE(@MailboxType,             MailboxType),
    HiddenFromAddressLists  = COALESCE(@HiddenFromAddressLists,  HiddenFromAddressLists),
    LitigationHoldEnabled   = COALESCE(@LitigationHoldEnabled,   LitigationHoldEnabled),
    ArchiveStatus           = COALESCE(@ArchiveStatus,           ArchiveStatus),
    ForwardingAddress       = COALESCE(@ForwardingAddress,       ForwardingAddress),
    ForwardingSmtpAddress   = COALESCE(@ForwardingSmtpAddress,   ForwardingSmtpAddress),
    GrantSendOnBehalfTo     = COALESCE(@GrantSendOnBehalfTo,     GrantSendOnBehalfTo),
    IsDirSynced             = COALESCE(@IsDirSynced,             IsDirSynced),
    WhenMailboxCreated      = COALESCE(@WhenMailboxCreated,      WhenMailboxCreated),
    WhenChangedUTC          = COALESCE(@WhenChangedUTC,          WhenChangedUTC),
    LastSyncedAt            = SYSUTCDATETIME(),
    LastModifiedAt          = SYSUTCDATETIME(),
    IsDeleted               = 0,
    DeletedAt               = NULL,
    SyncSource              = 'Delta',
    LastSyncRunId           = @RunId
WHEN NOT MATCHED THEN INSERT (
    ExchangeGuid, UserId, PrimarySmtpAddress, UserPrincipalName,
    DisplayName, Alias, RecipientTypeDetails, MailboxType,
    HiddenFromAddressLists, LitigationHoldEnabled, ArchiveStatus,
    ForwardingAddress, ForwardingSmtpAddress, GrantSendOnBehalfTo,
    IsDirSynced, WhenMailboxCreated, WhenChangedUTC,
    SyncSource, LastSyncRunId
) VALUES (
    @ExchangeGuid, @UserId, @PrimarySmtp, @UPN,
    @DisplayName, @Alias, @RecipientTypeDetails, @MailboxType,
    @HiddenFromAddressLists, @LitigationHoldEnabled, @ArchiveStatus,
    @ForwardingAddress, @ForwardingSmtpAddress, @GrantSendOnBehalfTo,
    @IsDirSynced, @WhenMailboxCreated, @WhenChangedUTC,
    'Delta', @RunId
)
"@

$existsSql = "SELECT COUNT(1) FROM dbo.Mailboxes WHERE ExchangeGuid = @ExchangeGuid"
$phaseSw   = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($m in $changedMailboxes) {
    try {
        if ($m._IsDeleted) {
            $rows = Invoke-SqlNonQuery -Query $softDeleteSql -Parameters @{
                '@ExchangeGuid' = [guid]$m.ExchangeGuid
                '@RunId'        = $runId
            }
            if ($rows -gt 0) {
                $softDeleted++
                Write-LogInfo "Soft-deleted: $($m.PrimarySmtpAddress)"
            }
            $processed++
            continue
        }

        $mailboxType = switch ($m.RecipientTypeDetails) {
            'UserMailbox'      { 'User'      }
            'SharedMailbox'    { 'Shared'    }
            'RoomMailbox'      { 'Room'      }
            'EquipmentMailbox' { 'Equipment' }
            default            { $m.RecipientTypeDetails }
        }

        $userId = if (-not [string]::IsNullOrWhiteSpace($m.ExternalDirectoryObjectId)) {
            try { [guid]$m.ExternalDirectoryObjectId } catch { $null }
        } else { $null }

        $grantSoB = if ($m.GrantSendOnBehalfTo -and $m.GrantSendOnBehalfTo.Count -gt 0) {
            ($m.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() }) -join ';'
        } else { $null }

        $exists = Invoke-SqlScalar -Query $existsSql `
            -Parameters @{ '@ExchangeGuid' = [guid]$m.ExchangeGuid }

        Invoke-SqlNonQuery -Query $upsertSql -Parameters @{
            '@ExchangeGuid'          = [guid]$m.ExchangeGuid
            '@UserId'                = $userId
            '@PrimarySmtp'           = $m.PrimarySmtpAddress
            '@UPN'                   = $m.UserPrincipalName
            '@DisplayName'           = $m.DisplayName
            '@Alias'                 = $m.Alias
            '@RecipientTypeDetails'  = $m.RecipientTypeDetails
            '@MailboxType'           = $mailboxType
            '@HiddenFromAddressLists'= if ($null -ne $m.HiddenFromAddressListsEnabled) { [bool]$m.HiddenFromAddressListsEnabled } else { $null }
            '@LitigationHoldEnabled' = if ($null -ne $m.LitigationHoldEnabled)         { [bool]$m.LitigationHoldEnabled }         else { $null }
            '@ArchiveStatus'         = if ($m.ArchiveStatus) { $m.ArchiveStatus.ToString() } else { $null }
            '@ForwardingAddress'     = $m.ForwardingAddress
            '@ForwardingSmtpAddress' = $m.ForwardingSmtpAddress
            '@GrantSendOnBehalfTo'   = $grantSoB
            '@IsDirSynced'           = if ($null -ne $m.IsDirSynced) { [bool]$m.IsDirSynced } else { $null }
            '@WhenMailboxCreated'    = $m.WhenMailboxCreated
            '@WhenChangedUTC'        = $m.WhenChangedUTC
            '@RunId'                 = $runId
        } | Out-Null

        if ([int]$exists -eq 0) { $inserted++ } else { $updated++ }
        $processed++
    }
    catch {
        $errors++
        Write-LogWarning "Failed to process $($m.PrimarySmtpAddress): $($_.Exception.Message)"
    }
}

$phaseSw.Stop()
Write-LogInfo "Changes applied in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $inserted inserted, $updated updated, $softDeleted soft-deleted, $errors errors"

# ──────────────────────────────────────────────────────────────
# Phase 4 — Advance timestamp
# Saved ONLY after processing completes.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 4 — Advance timestamp"

$newTimestamp = (Get-Date).ToUniversalTime().ToString('o')
Save-DeltaToken -TokenName "mailboxes_delta_timestamp" -TokenValue $newTimestamp
Write-LogInfo "Timestamp advanced to: $newTimestamp"

# ──────────────────────────────────────────────────────────────
# Finalise
# ──────────────────────────────────────────────────────────────

$overallSw.Stop()
$finalStatus = if ($errors -gt 0) { "PartialFailure" } else { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt          = SYSUTCDATETIME(),
    Status               = @Status,
    MailboxesInserted    = @Inserted,
    MailboxesUpdated     = @Updated,
    MailboxesSoftDeleted = @Deleted,
    MailboxesProcessed   = @Processed,
    ErrorCount           = @Errors,
    TokenAdvancedTo      = @Token
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'     = $runId
    '@Status'    = $finalStatus
    '@Inserted'  = $inserted
    '@Updated'   = $updated
    '@Deleted'   = $softDeleted
    '@Processed' = $processed
    '@Errors'    = $errors
    '@Token'     = $newTimestamp
} | Out-Null

Write-LogSummary @{
    "Status"       = $finalStatus
    "Inserted"     = $inserted
    "Updated"      = $updated
    "Soft-deleted" = $softDeleted
    "Errors"       = $errors
    "Processed"    = $processed
    "Duration"     = "$([Math]::Round($overallSw.Elapsed.TotalSeconds,1))s"
    "Run ID"       = $runId
}

Disconnect-ExoSession
Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
