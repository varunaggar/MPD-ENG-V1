<#
.SYNOPSIS
    Incremental permission sync — keeps all permission tables current.

.DESCRIPTION
    Scheduled every 15 minutes. Two distinct mechanisms:

    PHASE 0 — New mailbox mini-baseline
        Queries Mailboxes WHERE PermissionsBaselinedAt IS NULL.
        For any new mailbox, collects the full initial permission
        state (Full Access, Send-As, Send-on-Behalf, Folder permissions)
        then marks PermissionsBaselinedAt. This fires on the 15-minute
        cycle immediately after a new mailbox is detected — no waiting
        for the monthly full baseline.

    PHASES 3–6 — CloudAppEvents delta (Microsoft Defender for Cloud Apps)
        Queries the CloudAppEvents table via the Defender XDR Advanced
        Hunting API for permission-change audit events since the last run.
        Processes them per type:
          Full Access + Send-As  : event-driven re-fetch from EXO per
                                   affected mailbox (handles Add/Remove/Set
                                   without format-mismatch risk)
          Send-on-Behalf         : surgical replace from Set-Mailbox event
                                   payload (GrantSendOnBehalfTo parameter)
          Folder permissions     : surgical INSERT / soft-delete per event
                                   (Add/Set/Remove-MailboxFolderPermission)

    MFC group membership changes are handled separately by
    Invoke-MfcGroupDeltaSync and are NOT processed here.

    PREREQUISITES:
      - Microsoft Defender for Cloud Apps deployed, M365 activities connected
      - App registration: ThreatHunting.Read.All (Graph application permission)
      - Invoke-PermissionBaselineLoad completed at least once
        (sets permissions_delta_timestamp and PermissionsBaselinedAt)

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Scheduled Task:
      Trigger  : Daily, repeat every 15 minutes indefinitely
      Program  : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-PermissionDeltaSync.ps1"
      Settings : Do not start a new instance if already running
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-PermissionDeltaSync"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ══════════════════════════════════════════════════════════════
# Private helpers
# ══════════════════════════════════════════════════════════════

# Extracts a named parameter value from a RawEventData.Parameters array.
# Returns $null if the parameter is not present.
function Get-AuditParameter {
    param($Parameters, [string]$Name)
    if (-not $Parameters) { return $null }
    ($Parameters | Where-Object { $_.Name -eq $Name } | Select-Object -First 1).Value
}

# Resolves a mailbox identity string (SMTP or UPN) to ExchangeGuid.
# Returns $null if the mailbox is not in the Mailboxes table.
function Resolve-MailboxGuid {
    param([string]$Identity)
    if ([string]::IsNullOrWhiteSpace($Identity)) { return $null }
    return $mailboxLookup[$Identity.ToLower().Trim()]
}

# ══════════════════════════════════════════════════════════════
# Bootstrap
# ══════════════════════════════════════════════════════════════

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")     -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1")    -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")        -Force
Import-Module (Join-Path $sharedPath "ExoHelpers.psm1")        -Force
Import-Module (Join-Path $sharedPath "GraphHelpers.psm1")      -Force
Import-Module (Join-Path $sharedPath "AdminApiHelpers.psm1")   -Force
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

Write-LogSection "Invoke-PermissionDeltaSync"
Write-LogInfo "Config  : $ConfigPath"
Write-LogInfo "Server  : $($Config.Database.Server)"
Write-LogInfo "EXO Org : $($Config.ExchangeOnline.Organisation)"

Write-LogSection "Dependency Validation"
Initialize-ModuleDependencies -Config $Config

# ══════════════════════════════════════════════════════════════
# Authentication
# ══════════════════════════════════════════════════════════════

Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Initialize-SqlContext      -Config $Config
    Initialize-ExoContext      -Config $Config
    Connect-ExoSession         -Config $Config
    Initialize-GraphContext    -Config $Config
    Initialize-AdminApiContext -Config $Config
    Write-LogInfo "All connections established"
}
catch {
    Write-LogError "Authentication failed" -ErrorRecord $_
    Disconnect-ExoSession
    Close-Logging -Status "Failed"
    exit 1
}

# ══════════════════════════════════════════════════════════════
# Run tracking
# ══════════════════════════════════════════════════════════════

$runId     = [guid]::NewGuid()
$overallSw = [System.Diagnostics.Stopwatch]::StartNew()

Invoke-SqlNonQuery -Query @"
INSERT INTO dbo.SyncLog (RunId, FunctionName, StartedAt, Status)
VALUES (@RunId, @Function, SYSUTCDATETIME(), 'Running')
"@ -Parameters @{ '@RunId' = $runId; '@Function' = $scriptName } | Out-Null

$counters = @{
    NewMailboxes   = 0   # mailboxes mini-baselined in Phase 0
    FullAccess     = 0   # permission rows inserted / replaced
    SendAs         = 0
    SendOnBehalf   = 0
    FolderPerms    = 0
    EventsProcessed= 0   # CloudAppEvents rows processed
    Skipped        = 0   # events for unknown mailboxes
    Errors         = 0
}
$newTimestamp = $null

Write-LogInfo "Run ID: $runId"

# ══════════════════════════════════════════════════════════════
# SQL used across multiple phases — defined once, referenced below
# ══════════════════════════════════════════════════════════════

$softDeleteFullAccessSql = @"
UPDATE dbo.FullAccessPermissions SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId
WHERE ExchangeGuid=@ExchangeGuid AND IsDeleted=0
"@
$insertFullAccessSql = @"
INSERT INTO dbo.FullAccessPermissions
    (ExchangeGuid, TrusteeRawIdentity, IsInherited, AutoMapping, SyncSource, LastSyncRunId)
VALUES (@ExchangeGuid, @TrusteeRaw, @IsInherited, @AutoMapping, 'Delta', @RunId)
"@

$softDeleteSendAsSql = @"
UPDATE dbo.SendAsPermissions SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId
WHERE ExchangeGuid=@ExchangeGuid AND IsDeleted=0
"@
$insertSendAsSql = @"
INSERT INTO dbo.SendAsPermissions
    (ExchangeGuid, TrusteeRawIdentity, AccessControlType, SyncSource, LastSyncRunId)
VALUES (@ExchangeGuid, @TrusteeRaw, @AccessControlType, 'Delta', @RunId)
"@

$softDeleteSoBSql = @"
UPDATE dbo.SendOnBehalfPermissions SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId
WHERE ExchangeGuid=@ExchangeGuid AND IsDeleted=0
"@
$insertSoBSql = @"
INSERT INTO dbo.SendOnBehalfPermissions
    (ExchangeGuid, TrusteeRawIdentity, SyncSource, LastSyncRunId)
VALUES (@ExchangeGuid, @TrusteeRaw, 'Delta', @RunId)
"@

$softDeleteFolderByMailboxSql = @"
UPDATE dbo.FolderPermissions SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId
WHERE ExchangeGuid=@ExchangeGuid AND IsDeleted=0
"@
$softDeleteFolderByKeysSql = @"
UPDATE dbo.FolderPermissions SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId
WHERE ExchangeGuid=@ExchangeGuid AND FolderPath=@FolderPath AND TrusteeRawIdentity=@TrusteeRaw AND IsDeleted=0
"@
$insertFolderPermSql = @"
INSERT INTO dbo.FolderPermissions
    (ExchangeGuid, FolderPath, FolderName, TrusteeRawIdentity,
     AccessRights, SharingPermissionFlags, IsDefaultTrustee, SyncSource, LastSyncRunId)
VALUES
    (@ExchangeGuid, @FolderPath, @FolderName, @TrusteeRaw,
     @AccessRights, @SharingFlags, @IsDefault, 'Delta', @RunId)
"@

# ══════════════════════════════════════════════════════════════
# Shared: Replace-MailboxPermissions
# Used by both Phase 0 (mini-baseline) and Phase 3 (re-fetch).
# Replaces Full Access and Send-As for one mailbox using an
# already-open SQL connection.
# ══════════════════════════════════════════════════════════════

function Replace-MailboxPermissions {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [guid]$ExchangeGuid,
        [string]$Upn,
        [string]$MailboxType  # 'User' or 'Shared'
    )

    # Full Access — user mailboxes only (shared mailboxes use MFC groups)
    if ($MailboxType -eq 'User') {
        try {
            $faPerms = Get-EXOMailboxPermission -Identity $Upn -ErrorAction Stop |
                Where-Object {
                    -not $_.IsInherited -and
                    $_.Deny -ne $true -and
                    $_.User -notlike 'NT AUTHORITY\*' -and
                    $_.User -ne 'S-1-5-10'
                }

            Invoke-SqlNonQuery -Connection $Conn -Query $softDeleteFullAccessSql `
                -Parameters @{ '@ExchangeGuid' = $ExchangeGuid; '@RunId' = $runId } | Out-Null

            foreach ($p in $faPerms) {
                Invoke-SqlNonQuery -Connection $Conn -Query $insertFullAccessSql -Parameters @{
                    '@ExchangeGuid' = $ExchangeGuid
                    '@TrusteeRaw'   = [string]$p.User
                    '@IsInherited'  = [bool]$p.IsInherited
                    '@AutoMapping'  = if ($null -ne $p.AutoMapping) { [bool]$p.AutoMapping } else { $null }
                    '@RunId'        = $runId
                } | Out-Null
                $counters.FullAccess++
            }
        }
        catch {
            $counters.Errors++
            Write-LogWarning "  Full Access re-fetch failed for $Upn : $($_.Exception.Message)"
        }
    }

    # Send-As — all mailbox types
    try {
        $saPerms = Get-EXORecipientPermission -Identity $Upn -ErrorAction Stop |
            Where-Object {
                $_.Trustee -notlike 'NT AUTHORITY\*' -and
                $_.Trustee -ne 'S-1-5-10'
            }

        Invoke-SqlNonQuery -Connection $Conn -Query $softDeleteSendAsSql `
            -Parameters @{ '@ExchangeGuid' = $ExchangeGuid; '@RunId' = $runId } | Out-Null

        foreach ($p in $saPerms) {
            Invoke-SqlNonQuery -Connection $Conn -Query $insertSendAsSql -Parameters @{
                '@ExchangeGuid'      = $ExchangeGuid
                '@TrusteeRaw'        = [string]$p.Trustee
                '@AccessControlType' = [string]$p.AccessControlType
                '@RunId'             = $runId
            } | Out-Null
            $counters.SendAs++
        }
    }
    catch {
        $counters.Errors++
        Write-LogWarning "  Send-As re-fetch failed for $Upn : $($_.Exception.Message)"
    }
}

# ══════════════════════════════════════════════════════════════
# PHASE 0 — New mailbox mini-baseline
#
# Any mailbox with PermissionsBaselinedAt IS NULL has never had
# permissions collected. Run a full permission collection for each
# one and mark it as baselined. This handles new mailboxes without
# waiting for the monthly Invoke-PermissionBaselineLoad run.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 0 — New mailbox mini-baseline"

$newMailboxes = Invoke-SqlQuery -Query @"
SELECT
    CAST(ExchangeGuid AS NVARCHAR(36)) AS ExchangeGuid,
    PrimarySmtpAddress,
    UserPrincipalName,
    Alias,
    MailboxType,
    GrantSendOnBehalfTo
FROM dbo.Mailboxes
WHERE IsDeleted = 0 AND PermissionsBaselinedAt IS NULL
ORDER BY MailboxType, UserPrincipalName
"@

if ($newMailboxes.Count -eq 0) {
    Write-LogInfo "No new mailboxes pending baseline — skipping Phase 0"
}
else {
    Write-LogInfo "$($newMailboxes.Count) new mailbox(es) need permission collection"
    $phaseSw  = [System.Diagnostics.Stopwatch]::StartNew()
    $appAnchor = Get-AnchorMailboxHeader -Mode AppOnly

    $conn = Open-SqlConnection
    try {
        foreach ($mailbox in $newMailboxes) {
            $upn  = $mailbox.UserPrincipalName
            $guid = [guid]$mailbox.ExchangeGuid
            Write-LogInfo "  Mini-baseline: $upn ($($mailbox.MailboxType))"

            try {
                # Full Access + Send-As
                Replace-MailboxPermissions -Conn $conn -ExchangeGuid $guid `
                    -Upn $upn -MailboxType $mailbox.MailboxType

                # Send-on-Behalf from DB (already stored in Mailboxes.GrantSendOnBehalfTo)
                if (-not [string]::IsNullOrWhiteSpace($mailbox.GrantSendOnBehalfTo)) {
                    Invoke-SqlNonQuery -Connection $conn -Query $softDeleteSoBSql `
                        -Parameters @{ '@ExchangeGuid' = $guid; '@RunId' = $runId } | Out-Null

                    foreach ($trustee in ($mailbox.GrantSendOnBehalfTo -split ';' | Where-Object { $_ })) {
                        Invoke-SqlNonQuery -Connection $conn -Query $insertSoBSql -Parameters @{
                            '@ExchangeGuid' = $guid
                            '@TrusteeRaw'   = $trustee.Trim()
                            '@RunId'        = $runId
                        } | Out-Null
                        $counters.SendOnBehalf++
                    }
                }

                # Folder permissions — Graph folder list + Admin API per folder
                $folderPaths = Get-GraphMailFolderPaths -UserId $upn
                $anchor      = Get-AnchorMailboxHeader -Mode Mailbox -MailboxUpn $upn

                $tx = $conn.BeginTransaction()
                try {
                    Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                        -Query $softDeleteFolderByMailboxSql `
                        -Parameters @{ '@ExchangeGuid' = $guid; '@RunId' = $runId } | Out-Null

                    foreach ($folderPath in $folderPaths) {
                        try {
                            $normalPath = $folderPath.TrimStart('\')
                            $raw = Invoke-AdminApiPagedRequest `
                                -Endpoint      'MailboxFolderPermission' `
                                -CmdletName    'Get-MailboxFolderPermission' `
                                -Parameters    @{ Identity = "${upn}:\${normalPath}"; ResultSize = 'Unlimited' } `
                                -AnchorMailbox $anchor `
                                -Select        'Identity,FolderName,User,AccessRights,SharingPermissionFlags,IsValid'

                            foreach ($entry in $raw) {
                                if ($entry.IsValid -eq $false) { continue }
                                if ($entry.User -in @('Default','Anonymous') -and
                                    ([string]::IsNullOrWhiteSpace($entry.AccessRights) -or $entry.AccessRights -eq 'None')) { continue }

                                $rights = if ($entry.AccessRights -is [array]) { $entry.AccessRights -join ',' } else { [string]$entry.AccessRights }
                                $flags  = $null
                                if ($entry.SharingPermissionFlags) {
                                    $f = if ($entry.SharingPermissionFlags -is [array]) { $entry.SharingPermissionFlags -join ',' } else { [string]$entry.SharingPermissionFlags }
                                    if (-not [string]::IsNullOrWhiteSpace($f)) { $flags = $f }
                                }

                                Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                                    -Query $insertFolderPermSql -Parameters @{
                                        '@ExchangeGuid' = $guid
                                        '@FolderPath'   = "\$normalPath"
                                        '@FolderName'   = $entry.FolderName
                                        '@TrusteeRaw'   = $entry.User
                                        '@AccessRights' = $rights
                                        '@SharingFlags' = $flags
                                        '@IsDefault'    = ($entry.User -in @('Default','Anonymous'))
                                        '@RunId'        = $runId
                                    } | Out-Null
                                $counters.FolderPerms++
                            }
                        }
                        catch {
                            $counters.Errors++
                            Write-LogWarning "    $upn [$folderPath]: $($_.Exception.Message)"
                        }
                    }
                    $tx.Commit()
                }
                catch {
                    try { $tx.Rollback() } catch {}
                    throw
                }

                # Mark as baselined
                Invoke-SqlNonQuery -Connection $conn -Query @"
UPDATE dbo.Mailboxes SET PermissionsBaselinedAt=SYSUTCDATETIME() WHERE ExchangeGuid=@ExchangeGuid
"@ -Parameters @{ '@ExchangeGuid' = $guid } | Out-Null

                $counters.NewMailboxes++
            }
            catch {
                $counters.Errors++
                Write-LogWarning "  Mini-baseline failed for $upn : $($_.Exception.Message)"
            }
        }
    }
    finally {
        $conn.Dispose()
    }

    $phaseSw.Stop()
    Write-LogInfo "Phase 0 complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $($counters.NewMailboxes) mailbox(es) baselined"
}

# ══════════════════════════════════════════════════════════════
# PHASE 1 — Load delta timestamp
# If no timestamp exists this is the first run after baseline.
# Save the current time and exit — next run will have a valid window.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 1 — Load delta timestamp"

$storedTimestamp = Get-DeltaToken -TokenName "permissions_delta_timestamp"

if (-not $storedTimestamp) {
    Write-LogInfo "No permissions_delta_timestamp found — saving initial timestamp and exiting CloudAppEvents phase"
    $newTimestamp = (Get-Date).ToUniversalTime().ToString('o')
    Save-DeltaToken -TokenName "permissions_delta_timestamp" -TokenValue $newTimestamp

    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Success',
    MfcMembersProcessed=@New, ErrorCount=@Err, TokenAdvancedTo=@Token WHERE RunId=@RunId
"@ -Parameters @{
        '@RunId' = $runId; '@New' = $counters.NewMailboxes
        '@Err'   = $counters.Errors; '@Token' = $newTimestamp
    } | Out-Null

    Write-LogSummary @{
        "Status"        = "Success (initial run — timestamp saved)"
        "New mailboxes" = $counters.NewMailboxes
        "Errors"        = $counters.Errors
    }
    Disconnect-ExoSession
    Remove-OldLogFiles
    Close-Logging -Status "Success"
    exit 0
}

Write-LogInfo "Stored timestamp: $storedTimestamp"

# ══════════════════════════════════════════════════════════════
# PHASE 2 — Load mailbox lookup from DB
# Maps SMTP/UPN (lower) → { ExchangeGuid, MailboxType }
# Used throughout Phases 3–6 to resolve audit event identities.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 2 — Load mailbox lookup"

$mailboxLookup = @{}   # identity string (lower) → [PSCustomObject]{ExchangeGuid, MailboxType}

try {
    Invoke-SqlQuery -Query @"
SELECT
    CAST(ExchangeGuid AS NVARCHAR(36)) AS ExchangeGuid,
    PrimarySmtpAddress,
    UserPrincipalName,
    MailboxType
FROM dbo.Mailboxes WHERE IsDeleted = 0
"@ | ForEach-Object {
        $entry = $_
        if (-not [string]::IsNullOrWhiteSpace($entry.PrimarySmtpAddress)) {
            $mailboxLookup[$entry.PrimarySmtpAddress.ToLower()] = $entry
        }
        if (-not [string]::IsNullOrWhiteSpace($entry.UserPrincipalName)) {
            $mailboxLookup[$entry.UserPrincipalName.ToLower()] = $entry
        }
    }
    Write-LogInfo "Mailbox lookup loaded: $($mailboxLookup.Count) entries"
}
catch {
    Write-LogError "Phase 2: failed to load mailbox lookup — aborting CloudAppEvents phase" -ErrorRecord $_
    $counters.Errors++
}

# ══════════════════════════════════════════════════════════════
# PHASE 3 — Query CloudAppEvents
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 3 — CloudAppEvents query"

$overlapMinutes = [int]($Config.Sync.DeltaOverlapMinutes ?? 15)
$sinceUtc       = [datetime]::Parse($storedTimestamp).ToUniversalTime().AddMinutes(-$overlapMinutes)
$sinceKql       = $sinceUtc.ToString('yyyy-MM-ddTHH:mm:ssZ')

Write-LogInfo "Query window from: $sinceKql (overlap: ${overlapMinutes}min)"

$kqlQuery = @"
CloudAppEvents
| where Timestamp > datetime($sinceKql)
| where Application == 'Microsoft Exchange Online'
| where ActionType in (
    'Add-MailboxPermission', 'Remove-MailboxPermission', 'Set-MailboxPermission',
    'Add-RecipientPermission', 'Remove-RecipientPermission',
    'Set-Mailbox',
    'Add-MailboxFolderPermission', 'Set-MailboxFolderPermission', 'Remove-MailboxFolderPermission'
  )
| project Timestamp, ActionType, ObjectName, RawEventData
| take 100000
| order by Timestamp asc
"@

$huntingResults = @()

try {
    $phaseSw       = [System.Diagnostics.Stopwatch]::StartNew()
    $huntingResults = Invoke-GraphHuntingQuery -Query $kqlQuery
    $phaseSw.Stop()
    Write-LogInfo "CloudAppEvents returned $($huntingResults.Count) event(s) in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s"
}
catch {
    $counters.Errors++
    Write-LogError "Phase 3: CloudAppEvents query failed — skipping Phases 4–6" -ErrorRecord $_
}

if ($huntingResults.Count -eq 0 -and $counters.Errors -eq 0) {
    Write-LogInfo "No permission change events in window"
}

# ══════════════════════════════════════════════════════════════
# PHASE 4 — Full Access + Send-As (event-driven re-fetch)
#
# Extract unique affected mailboxes from Full Access and Send-As
# events, then re-fetch the FULL current permission state from EXO
# for each. This avoids trustee identity format mismatch issues
# and handles Add/Remove/Set with a single code path.
#
# One SQL connection held open for the phase.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 4 — Full Access and Send-As"

$faEvents  = $huntingResults | Where-Object {
    $_.ActionType -in @('Add-MailboxPermission','Remove-MailboxPermission','Set-MailboxPermission')
}
$saEvents  = $huntingResults | Where-Object {
    $_.ActionType -in @('Add-RecipientPermission','Remove-RecipientPermission')
}

# Collect unique affected ExchangeGuids across both event types
$affectedGuids = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

foreach ($e in ($faEvents + $saEvents)) {
    $params   = $e.RawEventData?.Parameters
    $identity = Get-AuditParameter $params 'Identity'
    if (-not $identity) { $identity = $e.ObjectName }
    $mb = Resolve-MailboxGuid $identity
    if ($mb) { [void]$affectedGuids.Add($mb.ExchangeGuid) }
    else {
        $counters.Skipped++
        Write-LogInfo "  Phase 4: unknown mailbox '$identity' — skipped"
    }
}

if ($affectedGuids.Count -gt 0) {
    Write-LogInfo "Re-fetching Full Access + Send-As for $($affectedGuids.Count) affected mailbox(es)"
    $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()

    # Build a quick GUID → mailbox record lookup for the affected set
    $guidToMailbox = @{}
    foreach ($e in ($faEvents + $saEvents)) {
        $params   = $e.RawEventData?.Parameters
        $identity = Get-AuditParameter $params 'Identity'
        if (-not $identity) { $identity = $e.ObjectName }
        $mb = Resolve-MailboxGuid $identity
        if ($mb) { $guidToMailbox[$mb.ExchangeGuid.ToString().ToLower()] = $mb }
    }

    $conn = Open-SqlConnection
    try {
        foreach ($guidStr in $affectedGuids) {
            $mb   = $guidToMailbox[$guidStr.ToLower()]
            if (-not $mb) { continue }
            $guid = [guid]$mb.ExchangeGuid
            $upn  = $mb.UserPrincipalName

            Write-LogInfo "  Re-fetching permissions: $upn"
            Replace-MailboxPermissions -Conn $conn -ExchangeGuid $guid `
                -Upn $upn -MailboxType $mb.MailboxType
        }
    }
    finally {
        $conn.Dispose()
    }

    $phaseSw.Stop()
    Write-LogInfo "Phase 4 complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — FullAccess: $($counters.FullAccess), SendAs: $($counters.SendAs)"
}
else {
    Write-LogInfo "No Full Access or Send-As events in window"
}

# ══════════════════════════════════════════════════════════════
# PHASE 5 — Send-on-Behalf (surgical from Set-Mailbox events)
#
# Set-Mailbox with GrantSendOnBehalfTo replaces the entire list.
# The event payload contains the new value — we replace all SoB
# for that mailbox directly from the event, no EXO call needed.
# Multiple Set-Mailbox events for the same mailbox are processed
# in timestamp order so the last one (final state) wins.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 5 — Send-on-Behalf"

$sobEvents = $huntingResults | Where-Object { $_.ActionType -eq 'Set-Mailbox' }

if ($sobEvents.Count -gt 0) {
    # Filter to only Set-Mailbox events that include GrantSendOnBehalfTo
    $sobEvents = $sobEvents | Where-Object {
        $null -ne (Get-AuditParameter $_.RawEventData?.Parameters 'GrantSendOnBehalfTo')
    }
}

if ($sobEvents.Count -eq 0) {
    Write-LogInfo "No Send-on-Behalf events in window"
}
else {
    Write-LogInfo "$($sobEvents.Count) Send-on-Behalf event(s) to process"
    $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()

    $conn = Open-SqlConnection
    try {
        foreach ($e in $sobEvents) {   # already in timestamp order from query
            $params   = $e.RawEventData?.Parameters
            $identity = Get-AuditParameter $params 'Identity'
            if (-not $identity) { $identity = $e.ObjectName }
            $mb = Resolve-MailboxGuid $identity

            if (-not $mb) {
                $counters.Skipped++
                Write-LogInfo "  Phase 5: unknown mailbox '$identity' — skipped"
                continue
            }

            $guid      = [guid]$mb.ExchangeGuid
            $rawSoBVal = Get-AuditParameter $params 'GrantSendOnBehalfTo'

            # Parse trustee list — may be comma or semicolon separated, may be empty
            $trustees = if (-not [string]::IsNullOrWhiteSpace($rawSoBVal)) {
                $rawSoBVal -split '[,;]' |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            } else { @() }

            Write-LogInfo "  SoB replace: $($mb.UserPrincipalName) — $($trustees.Count) trustee(s)"

            try {
                Invoke-SqlNonQuery -Connection $conn -Query $softDeleteSoBSql `
                    -Parameters @{ '@ExchangeGuid' = $guid; '@RunId' = $runId } | Out-Null

                foreach ($trustee in $trustees) {
                    Invoke-SqlNonQuery -Connection $conn -Query $insertSoBSql -Parameters @{
                        '@ExchangeGuid' = $guid
                        '@TrusteeRaw'   = $trustee
                        '@RunId'        = $runId
                    } | Out-Null
                    $counters.SendOnBehalf++
                }
                $counters.EventsProcessed++
            }
            catch {
                $counters.Errors++
                Write-LogWarning "  SoB replace failed for $($mb.UserPrincipalName): $($_.Exception.Message)"
            }
        }
    }
    finally {
        $conn.Dispose()
    }

    $phaseSw.Stop()
    Write-LogInfo "Phase 5 complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $($counters.SendOnBehalf) SoB entries written"
}

# ══════════════════════════════════════════════════════════════
# PHASE 6 — Folder permissions (surgical from audit events)
#
# Add-MailboxFolderPermission → INSERT row
# Set-MailboxFolderPermission → soft-delete existing for this
#                               folder+trustee, INSERT with new rights
# Remove-MailboxFolderPermission → soft-delete existing
#
# Events are already in timestamp order so sequential processing
# produces the correct final state even with multiple operations
# on the same folder in one window.
#
# Folder identity format in events: "mailbox@domain.com:\Calendar"
# Split on ':\ ' to get mailbox part and folder path.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 6 — Folder permissions"

$folderEvents = $huntingResults | Where-Object {
    $_.ActionType -in @(
        'Add-MailboxFolderPermission',
        'Set-MailboxFolderPermission',
        'Remove-MailboxFolderPermission'
    )
}

if ($folderEvents.Count -eq 0) {
    Write-LogInfo "No folder permission events in window"
}
else {
    Write-LogInfo "$($folderEvents.Count) folder permission event(s) to process"
    $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()

    $conn = Open-SqlConnection
    try {
        foreach ($e in $folderEvents) {
            $params   = $e.RawEventData?.Parameters
            $identity = Get-AuditParameter $params 'Identity'
            if (-not $identity) { $identity = $e.ObjectName }

            # Parse folder identity: "mailbox@domain.com:\Calendar" or "alias:\Inbox\Reports"
            $colonBackslashIdx = $identity.IndexOf(':\')
            if ($colonBackslashIdx -lt 0) {
                $counters.Skipped++
                Write-LogWarning "  Phase 6: cannot parse folder identity '$identity' — skipped"
                continue
            }

            $mailboxPart = $identity.Substring(0, $colonBackslashIdx)
            $folderPath  = '\' + $identity.Substring($colonBackslashIdx + 2)
            $mb          = Resolve-MailboxGuid $mailboxPart

            if (-not $mb) {
                $counters.Skipped++
                Write-LogInfo "  Phase 6: unknown mailbox '$mailboxPart' — skipped"
                continue
            }

            $guid      = [guid]$mb.ExchangeGuid
            $trustee   = Get-AuditParameter $params 'User'
            if (-not $trustee) { $trustee = Get-AuditParameter $params 'Trustee' }
            if (-not $trustee) {
                $counters.Skipped++
                Write-LogWarning "  Phase 6: no trustee in event for '$identity' ($($e.ActionType)) — skipped"
                continue
            }

            try {
                switch ($e.ActionType) {
                    'Add-MailboxFolderPermission' {
                        $rights = Get-AuditParameter $params 'AccessRights'
                        if ($rights -is [array]) { $rights = $rights -join ',' }

                        Invoke-SqlNonQuery -Connection $conn -Query $insertFolderPermSql -Parameters @{
                            '@ExchangeGuid' = $guid
                            '@FolderPath'   = $folderPath
                            '@FolderName'   = Split-Path $folderPath -Leaf
                            '@TrusteeRaw'   = $trustee
                            '@AccessRights' = [string]$rights
                            '@SharingFlags' = $null
                            '@IsDefault'    = ($trustee -in @('Default','Anonymous'))
                            '@RunId'        = $runId
                        } | Out-Null
                        $counters.FolderPerms++
                    }

                    'Set-MailboxFolderPermission' {
                        $rights = Get-AuditParameter $params 'AccessRights'
                        if ($rights -is [array]) { $rights = $rights -join ',' }

                        # Soft-delete existing then insert with updated rights
                        Invoke-SqlNonQuery -Connection $conn -Query $softDeleteFolderByKeysSql -Parameters @{
                            '@ExchangeGuid' = $guid; '@FolderPath' = $folderPath
                            '@TrusteeRaw'   = $trustee; '@RunId' = $runId
                        } | Out-Null

                        Invoke-SqlNonQuery -Connection $conn -Query $insertFolderPermSql -Parameters @{
                            '@ExchangeGuid' = $guid
                            '@FolderPath'   = $folderPath
                            '@FolderName'   = Split-Path $folderPath -Leaf
                            '@TrusteeRaw'   = $trustee
                            '@AccessRights' = [string]$rights
                            '@SharingFlags' = $null
                            '@IsDefault'    = ($trustee -in @('Default','Anonymous'))
                            '@RunId'        = $runId
                        } | Out-Null
                        $counters.FolderPerms++
                    }

                    'Remove-MailboxFolderPermission' {
                        Invoke-SqlNonQuery -Connection $conn -Query $softDeleteFolderByKeysSql -Parameters @{
                            '@ExchangeGuid' = $guid; '@FolderPath' = $folderPath
                            '@TrusteeRaw'   = $trustee; '@RunId' = $runId
                        } | Out-Null
                    }
                }

                $counters.EventsProcessed++
            }
            catch {
                $counters.Errors++
                Write-LogWarning "  Phase 6: $($e.ActionType) failed for '$identity': $($_.Exception.Message)"
            }
        }
    }
    finally {
        $conn.Dispose()
    }

    $phaseSw.Stop()
    Write-LogInfo "Phase 6 complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $($counters.FolderPerms) folder permission entries written"
}

# ══════════════════════════════════════════════════════════════
# PHASE 7 — Advance delta timestamp
# Saved ONLY after all phases complete successfully.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 7 — Advance delta timestamp"

try {
    $newTimestamp = (Get-Date).ToUniversalTime().ToString('o')
    Save-DeltaToken -TokenName "permissions_delta_timestamp" -TokenValue $newTimestamp
    Write-LogInfo "Timestamp advanced to: $newTimestamp"
}
catch {
    $counters.Errors++
    Write-LogWarning "Failed to advance timestamp: $($_.Exception.Message)"
}

# ══════════════════════════════════════════════════════════════
# Finalise
# ══════════════════════════════════════════════════════════════

$overallSw.Stop()
$finalStatus = if ($counters.Errors -gt 0) { "PartialFailure" } else { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt                = SYSUTCDATETIME(),
    Status                     = @Status,
    MfcMembersProcessed        = @NewMailboxes,
    FullAccessProcessed        = @FullAccess,
    SendAsProcessed            = @SendAs,
    SendOnBehalfProcessed      = @SoB,
    FolderPermissionsProcessed = @FolderPerms,
    MailboxesSkipped           = @Skipped,
    ErrorCount                 = @Errors,
    TokenAdvancedTo            = @Token
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'       = $runId;          '@Status'      = $finalStatus
    '@NewMailboxes'= $counters.NewMailboxes
    '@FullAccess'  = $counters.FullAccess;  '@SendAs'      = $counters.SendAs
    '@SoB'         = $counters.SendOnBehalf; '@FolderPerms' = $counters.FolderPerms
    '@Skipped'     = $counters.Skipped;      '@Errors'      = $counters.Errors
    '@Token'       = $newTimestamp
} | Out-Null

Write-LogSummary @{
    "Status"               = $finalStatus
    "New mailboxes"        = $counters.NewMailboxes
    "CloudApp events"      = $counters.EventsProcessed
    "Full Access"          = $counters.FullAccess
    "Send-As"              = $counters.SendAs
    "Send-on-Behalf"       = $counters.SendOnBehalf
    "Folder permissions"   = $counters.FolderPerms
    "Skipped (unknown mbx)"= $counters.Skipped
    "Errors"               = $counters.Errors
    "Duration"             = "$([Math]::Round($overallSw.Elapsed.TotalSeconds, 1))s"
    "Run ID"               = $runId
}

Disconnect-ExoSession
Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
