<#
.SYNOPSIS
    Full baseline load — fetches all Exchange Online mailboxes into
    the Mailboxes table.

.DESCRIPTION
    Run once on initial deployment, and weekly as a reconciliation pass.

    Steps:
      1. Fetches all mailboxes via Get-EXOMailbox (ResultSize Unlimited)
      2. MERGEs every mailbox into the Mailboxes table (idempotent)
         Uses a single open SQL connection for the whole loop.
      3. Reconciles deletions — soft-deletes any DB row whose ExchangeGuid
         is no longer in EXO (catches permanent deletions that delta misses)
      4. Saves the current UTC timestamp so Invoke-MailboxDeltaSync can run

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Typical duration: 15–45 minutes for a 40,000 mailbox tenant.
    Schedule weekly (e.g. Sunday 03:00) after Invoke-UserBaselineLoad.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-MailboxBaselineLoad"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ──────────────────────────────────────────────────────────────
# Bootstrap
# ──────────────────────────────────────────────────────────────

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")     -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1")    -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")        -Force
Import-Module (Join-Path $sharedPath "ExoHelpers.psm1")        -Force
Import-Module (Join-Path $sharedPath "DependencyHelpers.psm1") -Force

try {
    $Config = Import-SyncConfig -Path $ConfigPath
    Initialize-Logging -Config $Config -ProcessName $scriptName
}
catch {
    Write-Host "FATAL BOOTSTRAP ERROR in $scriptName" -ForegroundColor Red
    Write-Host "Message : $($_.Exception.Message)"    -ForegroundColor White
    exit 1
}

Write-LogSection "Invoke-MailboxBaselineLoad"
Write-LogInfo "Config       : $ConfigPath"
Write-LogInfo "Server       : $($Config.Database.Server)"
Write-LogInfo "Database     : $($Config.Database.Name)"
Write-LogInfo "EXO Org      : $($Config.ExchangeOnline.Organisation)"
Write-LogInfo "Mailbox types: $($Config.ExchangeOnline.MailboxTypes)"

Write-LogSection "Dependency Validation"
Initialize-ModuleDependencies -Config $Config

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

$inserted       = 0
$updated        = 0
$softDeleted    = 0
$processed      = 0
$errors         = 0
$savedTimestamp = $null
$overallSw      = [System.Diagnostics.Stopwatch]::StartNew()

Write-LogInfo "Run ID: $runId"

# Pre-load known UserIds to prevent FK violations when a mailbox
# arrives before its user (race condition between User and Mailbox delta syncs)
$knownUserIds = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
try {
    Invoke-SqlQuery -Query "SELECT CAST(UserId AS NVARCHAR(36)) AS UserId FROM dbo.Users WHERE IsDeleted=0" |
        ForEach-Object { [void]$knownUserIds.Add($_.UserId) }
    Write-LogInfo "Known UserIds loaded: $($knownUserIds.Count)"
}
catch {
    Write-LogWarning "Could not load UserIds — UserId will be set to NULL for unknown users: $($_.Exception.Message)"
}

# ──────────────────────────────────────────────────────────────
# EXO mailbox properties — fetched on every Get-EXOMailbox call
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
# Phase 1 — Fetch all mailboxes from EXO
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 1 — EXO mailbox fetch"

$rawMailboxes = $null
try {
    $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()

    $rawMailboxes = Get-EXOMailbox `
        -ResultSize Unlimited `
        -Properties $mailboxProps `
        -ErrorAction Stop |
        Where-Object { $_.RecipientTypeDetails -in $mailboxTypes }

    $phaseSw.Stop()
    Write-LogInfo "EXO returned $($rawMailboxes.Count) mailboxes in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s"
}
catch {
    Write-LogError "Phase 1 (EXO fetch) failed — aborting" -ErrorRecord $_
    Disconnect-ExoSession
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = "EXO fetch failed: $($_.Exception.Message)" } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

# Build a set of ExchangeGuids from EXO for deletion reconciliation in Phase 3
$exoGuids = [System.Collections.Generic.HashSet[string]]::new(
    ($rawMailboxes | ForEach-Object { $_.ExchangeGuid.ToString() }),
    [System.StringComparer]::OrdinalIgnoreCase
)

# ──────────────────────────────────────────────────────────────
# Phase 2 — MERGE each mailbox into the Mailboxes table
#
# One connection held open for the entire loop — efficient at scale.
# Per-mailbox error handling: a single failure does not stop the batch.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 2 — SQL upsert"

$mergeSql = @"
MERGE dbo.Mailboxes AS target
USING (SELECT @ExchangeGuid AS ExchangeGuid) AS source
ON target.ExchangeGuid = source.ExchangeGuid
WHEN MATCHED THEN UPDATE SET
    UserId                  = @UserId,
    PrimarySmtpAddress      = @PrimarySmtp,
    UserPrincipalName       = @UPN,
    DisplayName             = @DisplayName,
    Alias                   = @Alias,
    RecipientTypeDetails    = @RecipientTypeDetails,
    MailboxType             = @MailboxType,
    HiddenFromAddressLists  = @HiddenFromAddressLists,
    LitigationHoldEnabled   = @LitigationHoldEnabled,
    ArchiveStatus           = @ArchiveStatus,
    ForwardingAddress       = @ForwardingAddress,
    ForwardingSmtpAddress   = @ForwardingSmtpAddress,
    GrantSendOnBehalfTo     = @GrantSendOnBehalfTo,
    IsDirSynced             = @IsDirSynced,
    WhenMailboxCreated      = @WhenMailboxCreated,
    WhenChangedUTC          = @WhenChangedUTC,
    LastSyncedAt            = SYSUTCDATETIME(),
    LastModifiedAt          = SYSUTCDATETIME(),
    IsDeleted               = 0,
    DeletedAt               = NULL,
    SyncSource              = 'Baseline',
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
    'Baseline', @RunId
);
"@

$existsSql        = "SELECT COUNT(1) FROM dbo.Mailboxes WHERE ExchangeGuid = @ExchangeGuid"
$progressInterval = 2000
$phaseSw          = [System.Diagnostics.Stopwatch]::StartNew()

$conn = Open-SqlConnection
try {
    foreach ($m in $rawMailboxes) {
        try {
            # Normalise RecipientTypeDetails to a simple label
            $mailboxType = switch ($m.RecipientTypeDetails) {
                'UserMailbox'      { 'User'      }
                'SharedMailbox'    { 'Shared'    }
                'RoomMailbox'      { 'Room'      }
                'EquipmentMailbox' { 'Equipment' }
                default            { $m.RecipientTypeDetails }
            }

            # Entra Object ID — FK to Users.UserId (NULL if user not yet in DB)
            $userId = if (-not [string]::IsNullOrWhiteSpace($m.ExternalDirectoryObjectId)) {
                try {
                    $parsed = [guid]$m.ExternalDirectoryObjectId
                    if ($knownUserIds.Contains($parsed.ToString())) { $parsed } else { $null }
                } catch { $null }
            } else { $null }

            # GrantSendOnBehalfTo is a multi-value — flatten to semicolon string
            $grantSoB = if ($m.GrantSendOnBehalfTo -and $m.GrantSendOnBehalfTo.Count -gt 0) {
                ($m.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() }) -join ';'
            } else { $null }

            $exists = Invoke-SqlScalar -Connection $conn -Query $existsSql `
                -Parameters @{ '@ExchangeGuid' = [guid]$m.ExchangeGuid }

            Invoke-SqlNonQuery -Connection $conn -Query $mergeSql -Parameters @{
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
        }
        catch {
            $errors++
            Write-LogWarning "MERGE failed for $($m.PrimarySmtpAddress): $($_.Exception.Message)"
        }
        $processed++

        if ($processed % $progressInterval -eq 0) {
            $rate = [Math]::Round($processed / $phaseSw.Elapsed.TotalMinutes, 0)
            Write-LogInfo "Progress: $processed / $($rawMailboxes.Count) | Errors: $errors | $rate/min"
        }
    }
}
finally {
    $conn.Dispose()
}

$phaseSw.Stop()
Write-LogInfo "SQL upsert complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $inserted inserted, $updated updated, $errors errors"

# ──────────────────────────────────────────────────────────────
# Phase 3 — Reconcile deletions
# Soft-delete any DB row whose ExchangeGuid is no longer in EXO.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 3 — Deletion reconciliation"

try {
    $dbGuids = Invoke-SqlQuery -Query @"
SELECT CAST(ExchangeGuid AS NVARCHAR(36)) AS ExchangeGuid
FROM dbo.Mailboxes WHERE IsDeleted = 0
"@ | ForEach-Object { $_.ExchangeGuid }

    Write-LogInfo "DB active: $($dbGuids.Count) | EXO returned: $($exoGuids.Count)"

    $softDeleteSql = @"
UPDATE dbo.Mailboxes SET
    IsDeleted = 1, DeletedAt = SYSUTCDATETIME(),
    LastSyncedAt = SYSUTCDATETIME(), SyncSource = 'Baseline', LastSyncRunId = @RunId
WHERE ExchangeGuid = @ExchangeGuid AND IsDeleted = 0
"@
    foreach ($guid in $dbGuids) {
        if (-not $exoGuids.Contains($guid)) {
            try {
                $rows = Invoke-SqlNonQuery -Query $softDeleteSql `
                    -Parameters @{ '@ExchangeGuid' = [guid]$guid; '@RunId' = $runId }
                if ($rows -gt 0) {
                    $softDeleted++
                    Write-LogInfo "Soft-deleted: $guid (no longer in EXO)"
                }
            }
            catch {
                $errors++
                Write-LogWarning "Failed to soft-delete $guid : $($_.Exception.Message)"
            }
        }
    }

    Write-LogInfo "Reconciliation complete — $softDeleted soft-deleted"
}
catch {
    $errors++
    Write-LogError "Phase 3 (reconciliation) error" -ErrorRecord $_
}

# ──────────────────────────────────────────────────────────────
# Phase 4 — Save delta timestamp
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 4 — Save delta timestamp"

try {
    $savedTimestamp = (Get-Date).ToUniversalTime().ToString('o')
    Save-DeltaToken -TokenName "mailboxes_delta_timestamp" -TokenValue $savedTimestamp
    Write-LogInfo "Delta timestamp saved: $savedTimestamp"
}
catch {
    $errors++
    Write-LogError "Failed to save delta timestamp" -ErrorRecord $_
}

# ──────────────────────────────────────────────────────────────
# Finalise
# ──────────────────────────────────────────────────────────────

$overallSw.Stop()
$finalStatus = if ($errors -gt 0 -and $inserted -eq 0 -and $updated -eq 0) { "Failed" }
               elseif ($errors -gt 0) { "PartialFailure" }
               else                   { "Success" }

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
    '@Token'     = $savedTimestamp
} | Out-Null

Write-LogSummary @{
    "Status"             = $finalStatus
    "Mailboxes inserted" = $inserted
    "Mailboxes updated"  = $updated
    "Soft-deleted"       = $softDeleted
    "Errors"             = $errors
    "Total processed"    = $processed
    "Duration"           = "$([Math]::Round($overallSw.Elapsed.TotalMinutes, 2)) min"
    "Run ID"             = $runId
}

Disconnect-ExoSession
Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
