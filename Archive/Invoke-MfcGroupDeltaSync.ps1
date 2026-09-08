<#
.SYNOPSIS
    Incremental delta sync — keeps MfcGroups and MfcGroupMembers aligned
    with changes to mfc-SHXXXXX distribution groups.

.DESCRIPTION
    Scheduled every 15 minutes. Uses the Graph /groups/delta token
    stored by Invoke-PermissionBaselineLoad.

    Azure AD updates a group's lastModifiedDateTime when members are
    added or removed, so membership changes cause the group to appear
    in the delta response. For any changed MFC group, the full current
    membership is re-fetched from the Admin API and replaced.

    Steps:
      1. Read mfc_groups_delta_token from DeltaTokens
      2. Call Graph /groups/delta with that token
      3. For each changed group in the delta:
           - Skip non-MFC groups (displayName does not start with 'mfc-SH'
             and GroupObjectId is not in our MfcGroups table)
           - @removed → soft-delete MfcGroups + delete MfcGroupMembers
           - New/changed → MERGE MfcGroups + replace MfcGroupMembers
             (DELETE then INSERT in a transaction)
      4. Save new delta token ONLY after all processing succeeds
      5. On HTTP 410 Gone: deactivate token and exit — re-run baseline to recover

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Scheduled Task:
      Trigger  : Daily, repeat every 15 minutes indefinitely
      Program  : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-MfcGroupDeltaSync.ps1"
      Settings : Do not start a new instance if already running
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-MfcGroupDeltaSync"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ══════════════════════════════════════════════════════════════
# Bootstrap
# ══════════════════════════════════════════════════════════════

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")     -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1")    -Force
Import-Module (Join-Path $sharedPath "GraphHelpers.psm1")      -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")        -Force
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

Write-LogSection "Invoke-MfcGroupDeltaSync"
Write-LogInfo "Config   : $ConfigPath"
Write-LogInfo "Server   : $($Config.Database.Server)"
Write-LogInfo "Database : $($Config.Database.Name)"

Write-LogSection "Dependency Validation"
Initialize-ModuleDependencies -Config $Config

# ══════════════════════════════════════════════════════════════
# Authentication
# ══════════════════════════════════════════════════════════════

Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Initialize-SqlContext      -Config $Config
    Initialize-GraphContext    -Config $Config
    Initialize-AdminApiContext -Config $Config
    Write-LogInfo "Service principal, SQL, Graph and Admin API connected"
}
catch {
    Write-LogError "Authentication failed" -ErrorRecord $_
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
    GroupsProcessed = 0   # MFC groups changed (new + changed + removed)
    MembersWritten  = 0   # MfcGroupMembers rows inserted
    SoftDeleted     = 0   # MFC groups soft-deleted (removed from tenant)
    Skipped         = 0   # Non-MFC groups in delta (expected — silently skipped)
    Errors          = 0
}
$newToken = $null

Write-LogInfo "Run ID: $runId"

# ══════════════════════════════════════════════════════════════
# PHASE 1 — Load delta token
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 1 — Fetch delta token"

$storedToken = Get-DeltaToken -TokenName "mfc_groups_delta_token"

if (-not $storedToken) {
    $msg = "No active mfc_groups_delta_token. Run Invoke-PermissionBaselineLoad first."
    Write-LogError $msg
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = $msg } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

Write-LogInfo "Token found (length: $($storedToken.Length))"

# ══════════════════════════════════════════════════════════════
# PHASE 2 — Load known MFC groups and shared mailbox aliases from DB
# Used to classify delta objects and resolve new group → mailbox links.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 2 — Load reference data"

$knownGroupIds = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
$aliasToGuid = @{}   # shared mailbox Alias (lower) → ExchangeGuid

try {
    Invoke-SqlQuery -Query @"
SELECT CAST(GroupObjectId AS NVARCHAR(36)) AS GroupObjectId
FROM dbo.MfcGroups WHERE IsDeleted = 0
"@ | ForEach-Object { [void]$knownGroupIds.Add($_.GroupObjectId) }

    Write-LogInfo "Known MFC groups: $($knownGroupIds.Count)"

    Invoke-SqlQuery -Query @"
SELECT Alias, CAST(ExchangeGuid AS NVARCHAR(36)) AS ExchangeGuid
FROM dbo.Mailboxes WHERE MailboxType='Shared' AND IsDeleted=0
"@ | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace($_.Alias)) {
            $aliasToGuid[$_.Alias.ToLower()] = $_.ExchangeGuid
        }
    }

    Write-LogInfo "Shared mailbox alias lookup: $($aliasToGuid.Count) entries"
}
catch {
    Write-LogError "Phase 2: failed to load reference data — aborting" -ErrorRecord $_
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = "Phase 2 failed: $($_.Exception.Message)" } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

# ══════════════════════════════════════════════════════════════
# PHASE 3 — Graph groups delta fetch
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 3 — Graph groups delta fetch"

$deltaUrl    = "$($Config.Graph.BaseUrl)/groups/delta" +
               "?`$select=id,displayName,mail" +
               "&`$deltatoken=$storedToken"
$deltaResult = $null

try {
    $phaseSw     = [System.Diagnostics.Stopwatch]::StartNew()
    $deltaResult = Invoke-GraphDeltaQuery -Uri $deltaUrl
    $phaseSw.Stop()
    Write-LogInfo "Delta fetch complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $($deltaResult.Objects.Count) changed group(s)"
}
catch {
    Write-LogError "Graph delta fetch failed" -ErrorRecord $_
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = "Graph delta failed: $($_.Exception.Message)" } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

# HTTP 410 — token expired
if ($deltaResult.TokenExpired) {
    $msg = "Delta token expired (HTTP 410). Re-run Invoke-PermissionBaselineLoad to recover."
    Write-LogWarning $msg
    Disable-DeltaToken -TokenName "mfc_groups_delta_token" -Reason $msg
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = $msg } | Out-Null
    Close-Logging -Status "Failed — token expired"
    exit 1
}

# No changes
if ($deltaResult.Objects.Count -eq 0) {
    Write-LogInfo "No group changes since last run"
    if ($deltaResult.DeltaToken) {
        Save-DeltaToken -TokenName "mfc_groups_delta_token" -TokenValue $deltaResult.DeltaToken
        $newToken = $deltaResult.DeltaToken
    }
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Success',
    MfcGroupsProcessed=0, ErrorCount=0, TokenAdvancedTo=@Token WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Token' = $newToken } | Out-Null
    Write-LogSummary @{ "Status" = "Success (no changes)"; "Duration" = "$([Math]::Round($overallSw.Elapsed.TotalSeconds,1))s" }
    Remove-OldLogFiles
    Close-Logging -Status "Success"
    exit 0
}

# ══════════════════════════════════════════════════════════════
# PHASE 4 — Apply group changes
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 4 — Apply group changes"

$mergeMfcGroupSql = @"
MERGE dbo.MfcGroups AS target
USING (SELECT @GroupObjectId AS GroupObjectId) AS source
ON target.GroupObjectId = source.GroupObjectId
WHEN MATCHED THEN UPDATE SET
    DisplayName        = @DisplayName,
    Mail               = @Mail,
    ExchangeGuid       = COALESCE(@ExchangeGuid, ExchangeGuid),
    SharedMailboxAlias = COALESCE(@SharedMailboxAlias, SharedMailboxAlias),
    IsDeleted          = 0,
    DeletedAt          = NULL,
    LastSyncedAt       = SYSUTCDATETIME(),
    SyncSource         = 'Delta',
    LastSyncRunId      = @RunId
WHEN NOT MATCHED THEN INSERT
    (GroupObjectId, DisplayName, Mail, ExchangeGuid, SharedMailboxAlias, SyncSource, LastSyncRunId)
VALUES
    (@GroupObjectId, @DisplayName, @Mail, @ExchangeGuid, @SharedMailboxAlias, 'Delta', @RunId);
"@

$softDeleteGroupSql = @"
UPDATE dbo.MfcGroups SET
    IsDeleted=1, DeletedAt=SYSUTCDATETIME(),
    LastSyncedAt=SYSUTCDATETIME(), SyncSource='Delta', LastSyncRunId=@RunId
WHERE GroupObjectId=@GroupObjectId AND IsDeleted=0
"@

$deleteMembersSql = "DELETE FROM dbo.MfcGroupMembers WHERE GroupObjectId = @GroupObjectId"

$insertMemberSql  = @"
INSERT INTO dbo.MfcGroupMembers
    (GroupObjectId, PermissionType, TrusteeRawIdentity, TrusteeDisplayName,
     TrusteeRecipientType, LastSyncRunId, LastSyncedAt)
VALUES
    (@GroupObjectId, @PermType, @TrusteeRaw, @TrusteeDisplay,
     @TrusteeRecType, @RunId, SYSUTCDATETIME())
"@

$appOnlyAnchor = Get-AnchorMailboxHeader -Mode AppOnly
$phaseSw       = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($group in $deltaResult.Objects) {

    # Classify: is this a group we track or should track?
    $isKnown  = $knownGroupIds.Contains($group.id)
    $isNewMfc = (-not $isKnown) -and
                (-not [string]::IsNullOrWhiteSpace($group.displayName)) -and
                ($group.displayName -like 'mfc-SH*')

    if (-not $isKnown -and -not $isNewMfc) {
        $counters.Skipped++
        continue
    }

    $counters.GroupsProcessed++

    # Removed group
    if ($group.'@removed') {
        if (-not $isKnown) { continue }   # removed before we ever tracked it

        Write-LogInfo "Group removed: $($group.id)"
        try {
            $conn = Open-SqlConnection
            try {
                $tx = $conn.BeginTransaction()
                Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                    -Query $softDeleteGroupSql `
                    -Parameters @{ '@GroupObjectId' = [guid]$group.id; '@RunId' = $runId } | Out-Null
                Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                    -Query $deleteMembersSql `
                    -Parameters @{ '@GroupObjectId' = [guid]$group.id } | Out-Null
                $tx.Commit()
            }
            catch {
                try { $tx.Rollback() } catch {}
                throw
            }
            finally {
                $conn.Dispose()
            }
            $counters.SoftDeleted++
            [void]$knownGroupIds.Remove($group.id)
        }
        catch {
            $counters.Errors++
            Write-LogWarning "Failed to remove group $($group.id): $($_.Exception.Message)"
        }
        continue
    }

    # New or changed MFC group
    $displayName  = $group.displayName
    $sharedAlias  = if ($displayName -match '^mfc-(.+)$') { $Matches[1] } else { $null }
    $exchangeGuid = if ($sharedAlias) { $aliasToGuid[$sharedAlias.ToLower()] } else { $null }

    if ($isNewMfc) {
        Write-LogInfo "New MFC group: $displayName ($($group.id))"
        if (-not $exchangeGuid) {
            Write-LogWarning "  No matching shared mailbox for alias '$sharedAlias'"
        }
    }
    else {
        Write-LogInfo "Changed MFC group: $displayName ($($group.id))"
    }

    try {
        # MERGE group record — COALESCE preserves FK/alias if delta returns partial object
        Invoke-SqlNonQuery -Query $mergeMfcGroupSql -Parameters @{
            '@GroupObjectId'      = [guid]$group.id
            '@DisplayName'        = $displayName
            '@Mail'               = $group.mail
            '@ExchangeGuid'       = if ($exchangeGuid) { [guid]$exchangeGuid } else { $null }
            '@SharedMailboxAlias' = $sharedAlias
            '@RunId'              = $runId
        } | Out-Null

        # Re-fetch full current membership from Admin API
        $groupIdentity = if (-not [string]::IsNullOrWhiteSpace($displayName)) { $displayName }
                         elseif ($group.mail) { $group.mail }
                         else                 { $group.id }

        $members = Invoke-AdminApiPagedRequest `
            -Endpoint      'DistributionGroupMember' `
            -CmdletName    'Get-DistributionGroupMember' `
            -Parameters    @{ Identity = $groupIdentity; ResultSize = 'Unlimited' } `
            -AnchorMailbox $appOnlyAnchor `
            -Select        'PrimarySmtpAddress,DisplayName,RecipientTypeDetails'

        # Replace membership atomically
        $conn = Open-SqlConnection
        try {
            $tx = $conn.BeginTransaction()

            Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                -Query $deleteMembersSql `
                -Parameters @{ '@GroupObjectId' = [guid]$group.id } | Out-Null

            foreach ($m in $members) {
                foreach ($permType in @('FullAccess', 'SendOnBehalf')) {
                    Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                        -Query $insertMemberSql `
                        -Parameters @{
                            '@GroupObjectId' = [guid]$group.id
                            '@PermType'      = $permType
                            '@TrusteeRaw'    = $m.PrimarySmtpAddress
                            '@TrusteeDisplay'= $m.DisplayName
                            '@TrusteeRecType'= $m.RecipientTypeDetails
                            '@RunId'         = $runId
                        } | Out-Null
                }
            }

            $tx.Commit()
        }
        catch {
            try { $tx.Rollback() } catch {}
            throw
        }
        finally {
            $conn.Dispose()
        }

        $counters.MembersWritten += $members.Count
        if ($isNewMfc) { [void]$knownGroupIds.Add($group.id) }
        Write-LogInfo "  $displayName — $($members.Count) member(s) written"
    }
    catch {
        $counters.Errors++
        Write-LogWarning "Failed to process group '$displayName' ($($group.id)): $($_.Exception.Message)"
    }
}

$phaseSw.Stop()
Write-LogInfo ("Phase 4 complete in {0:F1}s — {1} processed, {2} members written, {3} removed, {4} skipped, {5} errors" -f
    $phaseSw.Elapsed.TotalSeconds,
    $counters.GroupsProcessed, $counters.MembersWritten,
    $counters.SoftDeleted, $counters.Skipped, $counters.Errors)

# ══════════════════════════════════════════════════════════════
# PHASE 5 — Advance delta token
# Saved ONLY after all processing completes.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 5 — Advance delta token"

if ($deltaResult.DeltaToken) {
    try {
        Save-DeltaToken -TokenName "mfc_groups_delta_token" -TokenValue $deltaResult.DeltaToken
        $newToken = $deltaResult.DeltaToken
        Write-LogInfo "Delta token advanced"
    }
    catch {
        $counters.Errors++
        Write-LogWarning "Failed to save delta token — next run reprocesses this window: $($_.Exception.Message)"
    }
}
else {
    Write-LogWarning "No delta token returned — token NOT advanced"
}

# ══════════════════════════════════════════════════════════════
# Finalise
# ══════════════════════════════════════════════════════════════

$overallSw.Stop()
$finalStatus = if ($counters.Errors -gt 0) { "PartialFailure" } else { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt         = SYSUTCDATETIME(),
    Status              = @Status,
    MfcGroupsProcessed  = @Groups,
    MfcMembersProcessed = @Members,
    MailboxesSkipped    = @Skipped,
    ErrorCount          = @Errors,
    TokenAdvancedTo     = @Token
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'   = $runId;    '@Status'  = $finalStatus
    '@Groups'  = $counters.GroupsProcessed
    '@Members' = $counters.MembersWritten
    '@Skipped' = $counters.Skipped
    '@Errors'  = $counters.Errors
    '@Token'   = $newToken
} | Out-Null

Write-LogSummary @{
    "Status"           = $finalStatus
    "Groups processed" = $counters.GroupsProcessed
    "Members written"  = $counters.MembersWritten
    "Groups removed"   = $counters.SoftDeleted
    "Non-MFC skipped"  = $counters.Skipped
    "Errors"           = $counters.Errors
    "Duration"         = "$([Math]::Round($overallSw.Elapsed.TotalSeconds, 1))s"
    "Run ID"           = $runId
}

Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
