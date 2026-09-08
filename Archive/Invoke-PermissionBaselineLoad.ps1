<#
.SYNOPSIS
    Full baseline load — captures all Exchange Online permissions.

.DESCRIPTION
    Reads the active mailbox list from SQL (from Invoke-MailboxBaselineLoad)
    and collects all permission types.

    PERMISSION SOURCES:
      MFC group membership (Full Access + Send-on-Behalf on shared mailboxes)
        → Graph /groups to discover mfc-SH* groups
        → Admin API Get-DistributionGroupMember for each group
      Full Access on user mailboxes
        → Get-EXOMailboxPermission per mailbox
      Send-As on all mailboxes
        → Get-EXORecipientPermission per mailbox
      Send-on-Behalf on user mailboxes
        → Already stored in Mailboxes.GrantSendOnBehalfTo (no EXO call)
      Folder permissions on all mailboxes
        → Graph mailFolders API to enumerate folder paths
        → Admin API Get-MailboxFolderPermission per folder

    SQL WRITE PATTERN:
      MFC group members   — DELETE then INSERT per group in a transaction
      Full Access/Send-As — Soft-delete ALL at phase start, INSERT via
                            one open connection held for the whole loop
      Send-on-Behalf      — Soft-delete ALL then INSERT from DB
      Folder permissions  — Soft-delete + INSERT per mailbox in a transaction
                            (one connection held open, transaction per mailbox)

    PREREQUISITE: app registration needs Mail.ReadBasic.All (for Graph
    mailFolders) and Exchange.ManageAsAppV2 (for Admin API).

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    EXPECTED DURATION (40,000 mailbox tenant, sequential):
      Phase 2 MFC groups    : 30-90 min
      Phase 3 Full Access   : 4-8 hours
      Phase 4 Send-on-Behalf: 5-15 min
      Phase 5 Folder perms  : 12-24 hours
      Total                 : 18-34 hours

    Schedule as a monthly or on-demand task. Use delta sync for daily changes.

    Scheduled Task:
      Program  : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-PermissionBaselineLoad.ps1"
      Settings : ExecutionTimeLimit=48h, MultipleInstances=IgnoreNew
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-PermissionBaselineLoad"
$sharedPath            = Join-Path $PSScriptRoot "shared"

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

Write-LogSection "Invoke-PermissionBaselineLoad"
Write-LogInfo "Config   : $ConfigPath"
Write-LogInfo "Server   : $($Config.Database.Server)"
Write-LogInfo "Database : $($Config.Database.Name)"
Write-LogInfo "EXO Org  : $($Config.ExchangeOnline.Organisation)"

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
    Initialize-AdminApiContext -Config $Config
    Initialize-GraphContext    -Config $Config
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
    MfcGroups    = 0
    MfcMembers   = 0
    FullAccess   = 0
    SendAs       = 0
    SendOnBehalf = 0
    FolderPerms  = 0
    Skipped      = 0
    Errors       = 0
}

Write-LogInfo "Run ID: $runId"

# ══════════════════════════════════════════════════════════════
# PHASE 1 — Load active mailboxes from SQL
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 1 — Load mailboxes from DB"

$allMailboxes = $null
$aliasToGuid  = @{}   # shared mailbox Alias (lower) → ExchangeGuid

try {
    $allMailboxes = Invoke-SqlQuery -Query @"
SELECT
    CAST(ExchangeGuid AS NVARCHAR(36)) AS ExchangeGuid,
    PrimarySmtpAddress,
    UserPrincipalName,
    Alias,
    MailboxType
FROM dbo.Mailboxes
WHERE IsDeleted = 0
ORDER BY MailboxType, PrimarySmtpAddress
"@

    if (-not $allMailboxes -or $allMailboxes.Count -eq 0) {
        throw "No active mailboxes in DB. Run Invoke-MailboxBaselineLoad first."
    }

    $userMailboxes   = @($allMailboxes | Where-Object { $_.MailboxType -eq 'User'   })
    $sharedMailboxes = @($allMailboxes | Where-Object { $_.MailboxType -eq 'Shared' })

    foreach ($m in $sharedMailboxes) {
        if (-not [string]::IsNullOrWhiteSpace($m.Alias)) {
            $aliasToGuid[$m.Alias.ToLower()] = $m.ExchangeGuid
        }
    }

    Write-LogInfo "Loaded $($allMailboxes.Count) mailboxes — User: $($userMailboxes.Count), Shared: $($sharedMailboxes.Count)"
}
catch {
    Write-LogError "Phase 1 failed — aborting" -ErrorRecord $_
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = "Phase 1 failed: $($_.Exception.Message)" } | Out-Null
    Disconnect-ExoSession
    Close-Logging -Status "Failed"
    exit 1
}

# ══════════════════════════════════════════════════════════════
# PHASE 2 — MFC Groups: discovery (Graph) + membership (Admin API)
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 2 — MFC Groups"

$mergeMfcGroupSql = @"
MERGE dbo.MfcGroups AS target
USING (SELECT @GroupObjectId AS GroupObjectId) AS source
ON target.GroupObjectId = source.GroupObjectId
WHEN MATCHED THEN UPDATE SET
    DisplayName        = @DisplayName,
    Mail               = @Mail,
    ExchangeGuid       = @ExchangeGuid,
    SharedMailboxAlias = @SharedMailboxAlias,
    IsDeleted          = 0,
    DeletedAt          = NULL,
    LastSyncedAt       = SYSUTCDATETIME(),
    SyncSource         = 'Baseline',
    LastSyncRunId      = @RunId
WHEN NOT MATCHED THEN INSERT
    (GroupObjectId, DisplayName, Mail, ExchangeGuid, SharedMailboxAlias, SyncSource, LastSyncRunId)
VALUES
    (@GroupObjectId, @DisplayName, @Mail, @ExchangeGuid, @SharedMailboxAlias, 'Baseline', @RunId);
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

$phaseSw = [System.Diagnostics.Stopwatch]::StartNew()

try {
    $groupsUri = "$($Config.Graph.BaseUrl)/groups" +
                 "?`$filter=startsWith(displayName,'mfc-SH')" +
                 "&`$select=id,displayName,mail" +
                 "&`$count=true"

    Write-LogInfo "Querying Graph for mfc-SH* groups..."
    $mfcGroups = Invoke-GraphPagedRequest -Uri $groupsUri
    Write-LogInfo "Found $($mfcGroups.Count) mfc-SH* group(s)"

    $appOnlyAnchor = Get-AnchorMailboxHeader -Mode AppOnly

    foreach ($group in $mfcGroups) {
        try {
            $sharedAlias  = if ($group.displayName -match '^mfc-(.+)$') { $Matches[1] } else { $null }
            $exchangeGuid = if ($sharedAlias) { $aliasToGuid[$sharedAlias.ToLower()] } else { $null }

            if (-not $sharedAlias) {
                Write-LogWarning "Cannot extract alias from '$($group.displayName)' — skipping"
                $counters.Errors++
                continue
            }
            if (-not $exchangeGuid) {
                Write-LogWarning "No matching shared mailbox for alias '$sharedAlias'"
            }

            Invoke-SqlNonQuery -Query $mergeMfcGroupSql -Parameters @{
                '@GroupObjectId'      = [guid]$group.id
                '@DisplayName'        = $group.displayName
                '@Mail'               = $group.mail
                '@ExchangeGuid'       = if ($exchangeGuid) { [guid]$exchangeGuid } else { $null }
                '@SharedMailboxAlias' = $sharedAlias
                '@RunId'              = $runId
            } | Out-Null

            $counters.MfcGroups++

            # Fetch members directly from Admin API
            $members = Invoke-AdminApiPagedRequest `
                -Endpoint      'DistributionGroupMember' `
                -CmdletName    'Get-DistributionGroupMember' `
                -Parameters    @{ Identity = $group.displayName; ResultSize = 'Unlimited' } `
                -AnchorMailbox $appOnlyAnchor `
                -Select        'PrimarySmtpAddress,DisplayName,RecipientTypeDetails'

            # Replace membership atomically: DELETE + INSERT in a transaction
            $conn = Open-SqlConnection
            try {
                $tx = $conn.BeginTransaction()

                Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                    -Query      $deleteMembersSql `
                    -Parameters @{ '@GroupObjectId' = [guid]$group.id } | Out-Null

                foreach ($m in $members) {
                    foreach ($permType in @('FullAccess', 'SendOnBehalf')) {
                        Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                            -Query      $insertMemberSql `
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

            $counters.MfcMembers += $members.Count
            Write-LogInfo "  $($group.displayName): $($members.Count) member(s) written"
        }
        catch {
            $counters.Errors++
            Write-LogWarning "MFC group '$($group.displayName)' failed: $($_.Exception.Message)"
        }
    }

    $phaseSw.Stop()
    Write-LogInfo "Phase 2 complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $($counters.MfcGroups) groups, $($counters.MfcMembers) members"

    # Save initial groups delta token so Invoke-MfcGroupDeltaSync can run immediately
    Write-LogInfo "Acquiring initial groups delta token..."
    try {
        $initDelta = Invoke-GraphDeltaQuery -Uri "$($Config.Graph.BaseUrl)/groups/delta?`$select=id&`$top=999"
        if ($initDelta.DeltaToken) {
            Save-DeltaToken -TokenName 'mfc_groups_delta_token' -TokenValue $initDelta.DeltaToken
            Write-LogInfo "Groups delta token saved"
        }
    }
    catch {
        Write-LogWarning "Could not save groups delta token: $($_.Exception.Message)"
        $counters.Errors++
    }
}
catch {
    $counters.Errors++
    Write-LogError "Phase 2 error — continuing" -ErrorRecord $_
}

# ══════════════════════════════════════════════════════════════
# PHASE 3 — Full Access (user mailboxes) + Send-As (all mailboxes)
#
# Soft-delete all existing rows at phase start, then one connection
# is held open for the entire mailbox loop for efficiency.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 3 — Full Access and Send-As permissions"

$insertFullAccessSql = @"
INSERT INTO dbo.FullAccessPermissions
    (ExchangeGuid, TrusteeRawIdentity, IsInherited, AutoMapping, SyncSource, LastSyncRunId)
VALUES
    (@ExchangeGuid, @TrusteeRaw, @IsInherited, @AutoMapping, 'Baseline', @RunId)
"@

$insertSendAsSql = @"
INSERT INTO dbo.SendAsPermissions
    (ExchangeGuid, TrusteeRawIdentity, AccessControlType, SyncSource, LastSyncRunId)
VALUES
    (@ExchangeGuid, @TrusteeRaw, @AccessControlType, 'Baseline', @RunId)
"@

try {
    Write-LogInfo "Soft-deleting all existing FullAccessPermissions..."
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.FullAccessPermissions SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId WHERE IsDeleted=0
"@ -Parameters @{ '@RunId' = $runId } | Out-Null

    Write-LogInfo "Soft-deleting all existing SendAsPermissions..."
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SendAsPermissions SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId WHERE IsDeleted=0
"@ -Parameters @{ '@RunId' = $runId } | Out-Null
}
catch {
    $counters.Errors++
    Write-LogError "Phase 3: soft-delete failed" -ErrorRecord $_
}

$phaseSw         = [System.Diagnostics.Stopwatch]::StartNew()
$phase3Processed = 0
$progressEvery   = 500

$conn = Open-SqlConnection
try {
    foreach ($mailbox in $allMailboxes) {
        $phase3Processed++
        $upn  = $mailbox.UserPrincipalName
        $guid = [guid]$mailbox.ExchangeGuid

        try {
            if ($mailbox.MailboxType -eq 'User') {
                $faPerms = Get-EXOMailboxPermission -Identity $upn -ErrorAction Stop |
                    Where-Object {
                        -not $_.IsInherited -and
                        $_.Deny -ne $true -and
                        $_.User -notlike 'NT AUTHORITY\*' -and
                        $_.User -ne 'S-1-5-10'
                    }
                foreach ($perm in $faPerms) {
                    Invoke-SqlNonQuery -Connection $conn -Query $insertFullAccessSql -Parameters @{
                        '@ExchangeGuid' = $guid
                        '@TrusteeRaw'   = [string]$perm.User
                        '@IsInherited'  = [bool]$perm.IsInherited
                        '@AutoMapping'  = if ($null -ne $perm.AutoMapping) { [bool]$perm.AutoMapping } else { $null }
                        '@RunId'        = $runId
                    } | Out-Null
                    $counters.FullAccess++
                }
            }

            $saPerms = Get-EXORecipientPermission -Identity $upn -ErrorAction Stop |
                Where-Object {
                    $_.Trustee -notlike 'NT AUTHORITY\*' -and
                    $_.Trustee -ne 'S-1-5-10'
                }
            foreach ($perm in $saPerms) {
                Invoke-SqlNonQuery -Connection $conn -Query $insertSendAsSql -Parameters @{
                    '@ExchangeGuid'      = $guid
                    '@TrusteeRaw'        = [string]$perm.Trustee
                    '@AccessControlType' = [string]$perm.AccessControlType
                    '@RunId'             = $runId
                } | Out-Null
                $counters.SendAs++
            }
        }
        catch {
            $counters.Errors++
            $counters.Skipped++
            Write-LogWarning "Phase 3: $upn — $($_.Exception.Message)"
        }

        if ($phase3Processed % $progressEvery -eq 0) {
            $elapsed = [Math]::Round($phaseSw.Elapsed.TotalMinutes, 1)
            $rate    = if ($elapsed -gt 0) { [Math]::Round($phase3Processed / $elapsed, 0) } else { 0 }
            Write-LogInfo "Phase 3: $phase3Processed/$($allMailboxes.Count) | FullAccess: $($counters.FullAccess) | SendAs: $($counters.SendAs) | ${rate}/min | Errors: $($counters.Errors)"
        }
    }
}
finally {
    $conn.Dispose()
}

$phaseSw.Stop()
Write-LogInfo "Phase 3 complete in $([Math]::Round($phaseSw.Elapsed.TotalMinutes,2))min — FullAccess: $($counters.FullAccess), SendAs: $($counters.SendAs)"

# ══════════════════════════════════════════════════════════════
# PHASE 4 — Send-on-Behalf (user mailboxes, from DB)
#
# GrantSendOnBehalfTo already stored in Mailboxes from
# Invoke-MailboxBaselineLoad. Promoted here to normalised table.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 4 — Send-on-Behalf permissions"

$insertSoBSql = @"
INSERT INTO dbo.SendOnBehalfPermissions
    (ExchangeGuid, TrusteeRawIdentity, SyncSource, LastSyncRunId)
VALUES
    (@ExchangeGuid, @TrusteeRaw, 'Baseline', @RunId)
"@

try {
    Write-LogInfo "Soft-deleting all existing SendOnBehalfPermissions..."
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SendOnBehalfPermissions SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId WHERE IsDeleted=0
"@ -Parameters @{ '@RunId' = $runId } | Out-Null

    $sobRows = Invoke-SqlQuery -Query @"
SELECT CAST(ExchangeGuid AS NVARCHAR(36)) AS ExchangeGuid, GrantSendOnBehalfTo
FROM dbo.Mailboxes
WHERE MailboxType='User' AND IsDeleted=0
  AND GrantSendOnBehalfTo IS NOT NULL AND GrantSendOnBehalfTo <> ''
"@

    Write-LogInfo "$($sobRows.Count) user mailbox(es) with GrantSendOnBehalfTo set"

    $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()
    $conn    = Open-SqlConnection
    try {
        foreach ($row in $sobRows) {
            foreach ($trustee in ($row.GrantSendOnBehalfTo -split ';' | Where-Object { $_ })) {
                try {
                    Invoke-SqlNonQuery -Connection $conn -Query $insertSoBSql -Parameters @{
                        '@ExchangeGuid' = [guid]$row.ExchangeGuid
                        '@TrusteeRaw'   = $trustee.Trim()
                        '@RunId'        = $runId
                    } | Out-Null
                    $counters.SendOnBehalf++
                }
                catch {
                    $counters.Errors++
                    Write-LogWarning "SoB insert failed for $($row.ExchangeGuid): $($_.Exception.Message)"
                }
            }
        }
    }
    finally {
        $conn.Dispose()
    }

    $phaseSw.Stop()
    Write-LogInfo "Phase 4 complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $($counters.SendOnBehalf) entries written"
}
catch {
    $counters.Errors++
    Write-LogError "Phase 4 failed — continuing" -ErrorRecord $_
}

# ══════════════════════════════════════════════════════════════
# PHASE 5 — Folder Permissions
#
# Per mailbox:
#   1. Graph mailFolders → enumerate visible folder paths
#   2. Admin API Get-MailboxFolderPermission per folder
#   3. Atomic replace: soft-delete + INSERT per mailbox in a transaction
#
# One SQL connection held open for the whole phase.
# Per-mailbox transactions ensure consistent state per mailbox.
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 5 — Folder Permissions"

$softDeleteFolderSql = @"
UPDATE dbo.FolderPermissions
SET IsDeleted=1, DeletedAt=SYSUTCDATETIME(), LastSyncRunId=@RunId
WHERE ExchangeGuid=@ExchangeGuid AND IsDeleted=0
"@

$insertFolderPermSql = @"
INSERT INTO dbo.FolderPermissions
    (ExchangeGuid, FolderPath, FolderName, TrusteeRawIdentity,
     AccessRights, SharingPermissionFlags, IsDefaultTrustee, SyncSource, LastSyncRunId)
VALUES
    (@ExchangeGuid, @FolderPath, @FolderName, @TrusteeRaw,
     @AccessRights, @SharingFlags, @IsDefault, 'Baseline', @RunId)
"@

$phaseSw         = [System.Diagnostics.Stopwatch]::StartNew()
$phase5Processed = 0
$progressEvery   = 100

$conn = Open-SqlConnection
try {
    foreach ($mailbox in $allMailboxes) {
        $phase5Processed++
        $upn    = $mailbox.UserPrincipalName
        $guid   = [guid]$mailbox.ExchangeGuid
        $anchor = Get-AnchorMailboxHeader -Mode Mailbox -MailboxUpn $upn

        try {
            $folderPaths = Get-GraphMailFolderPaths -UserId $upn

            if (-not $folderPaths -or $folderPaths.Count -eq 0) {
                Write-LogInfo "  $upn — no folders returned by Graph, skipping"
                $counters.Skipped++
                continue
            }

            # Collect all folder permissions for this mailbox
            $permRows = [System.Collections.Generic.List[hashtable]]::new()

            foreach ($folderPath in $folderPaths) {
                try {
                    $normalPath = $folderPath.TrimStart('\')
                    $identity   = "${upn}:\${normalPath}"

                    $raw = Invoke-AdminApiPagedRequest `
                        -Endpoint      'MailboxFolderPermission' `
                        -CmdletName    'Get-MailboxFolderPermission' `
                        -Parameters    @{ Identity = $identity; ResultSize = 'Unlimited' } `
                        -AnchorMailbox $anchor `
                        -Select        'Identity,FolderName,User,AccessRights,SharingPermissionFlags,IsValid'

                    foreach ($entry in $raw) {
                        if ($entry.IsValid -eq $false) { continue }
                        if ($entry.User -in @('Default','Anonymous') -and
                            ([string]::IsNullOrWhiteSpace($entry.AccessRights) -or $entry.AccessRights -eq 'None')) {
                            continue
                        }

                        $rights = if ($entry.AccessRights -is [array]) { $entry.AccessRights -join ',' } else { [string]$entry.AccessRights }
                        $flags  = $null
                        if ($entry.SharingPermissionFlags) {
                            $f = if ($entry.SharingPermissionFlags -is [array]) { $entry.SharingPermissionFlags -join ',' } else { [string]$entry.SharingPermissionFlags }
                            if (-not [string]::IsNullOrWhiteSpace($f)) { $flags = $f }
                        }

                        $permRows.Add(@{
                            FolderPath   = "\$normalPath"
                            FolderName   = $entry.FolderName
                            TrusteeRaw   = $entry.User
                            AccessRights = $rights
                            SharingFlags = $flags
                            IsDefault    = ($entry.User -in @('Default','Anonymous'))
                        })
                        $counters.FolderPerms++
                    }
                }
                catch {
                    $counters.Errors++
                    Write-LogWarning "  $upn [$folderPath]: $($_.Exception.Message)"
                }
            }

            # Atomic replace: soft-delete existing + insert current — in one transaction
            $tx = $conn.BeginTransaction()
            try {
                Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                    -Query $softDeleteFolderSql `
                    -Parameters @{ '@ExchangeGuid' = $guid; '@RunId' = $runId } | Out-Null

                foreach ($r in $permRows) {
                    Invoke-SqlNonQuery -Connection $conn -Transaction $tx `
                        -Query $insertFolderPermSql `
                        -Parameters @{
                            '@ExchangeGuid' = $guid
                            '@FolderPath'   = $r.FolderPath
                            '@FolderName'   = $r.FolderName
                            '@TrusteeRaw'   = $r.TrusteeRaw
                            '@AccessRights' = $r.AccessRights
                            '@SharingFlags' = $r.SharingFlags
                            '@IsDefault'    = [bool]$r.IsDefault
                            '@RunId'        = $runId
                        } | Out-Null
                }

                $tx.Commit()
            }
            catch {
                try { $tx.Rollback() } catch {}
                throw
            }
        }
        catch {
            $counters.Errors++
            $counters.Skipped++
            Write-LogWarning "Phase 5: $upn — $($_.Exception.Message)"
        }

        if ($phase5Processed % $progressEvery -eq 0) {
            $elapsed   = [Math]::Round($phaseSw.Elapsed.TotalMinutes, 1)
            $rate      = if ($elapsed -gt 0) { [Math]::Round($phase5Processed / $elapsed, 1) } else { 0 }
            $remaining = if ($rate -gt 0) { [Math]::Round(($allMailboxes.Count - $phase5Processed) / $rate, 0) } else { '?' }
            Write-LogInfo "Phase 5: $phase5Processed/$($allMailboxes.Count) | FolderPerms: $($counters.FolderPerms) | ~${remaining}min remaining | Errors: $($counters.Errors)"
        }
    }
}
finally {
    $conn.Dispose()
}

$phaseSw.Stop()
Write-LogInfo "Phase 5 complete in $([Math]::Round($phaseSw.Elapsed.TotalMinutes,2))min — $($counters.FolderPerms) entries written"

# Mark all active mailboxes as baselined so Invoke-PermissionDeltaSync
# Phase 0 does not attempt to re-collect them on the next run.
Write-LogInfo "Marking all active mailboxes as baselined..."
try {
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.Mailboxes SET PermissionsBaselinedAt = SYSUTCDATETIME() WHERE IsDeleted = 0
"@ | Out-Null
    Write-LogInfo "PermissionsBaselinedAt set on all active mailboxes"
}
catch {
    $counters.Errors++
    Write-LogWarning "Failed to set PermissionsBaselinedAt: $($_.Exception.Message)"
}

# ══════════════════════════════════════════════════════════════
# PHASE 6 — Save timestamp
# ══════════════════════════════════════════════════════════════

Write-LogSection "Phase 6 — Save timestamp"

$savedTimestamp = $null
try {
    $savedTimestamp = (Get-Date).ToUniversalTime().ToString('o')
    Save-DeltaToken -TokenName "permissions_baseline_timestamp" -TokenValue $savedTimestamp
    Write-LogInfo "Baseline timestamp saved: $savedTimestamp"
}
catch {
    $counters.Errors++
    Write-LogError "Failed to save baseline timestamp" -ErrorRecord $_
}

# ══════════════════════════════════════════════════════════════
# Finalise
# ══════════════════════════════════════════════════════════════

$overallSw.Stop()

$finalStatus = if ($counters.Errors -gt 0 -and $counters.FullAccess -eq 0 -and
                   $counters.FolderPerms -eq 0 -and $counters.MfcGroups -eq 0) {
    "Failed"
} elseif ($counters.Errors -gt 0) { "PartialFailure" }
else                               { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt                = SYSUTCDATETIME(),
    Status                     = @Status,
    MfcGroupsProcessed         = @MfcGroups,
    MfcMembersProcessed        = @MfcMembers,
    FullAccessProcessed        = @FullAccess,
    SendAsProcessed            = @SendAs,
    SendOnBehalfProcessed      = @SendOnBehalf,
    FolderPermissionsProcessed = @FolderPerms,
    MailboxesSkipped           = @Skipped,
    ErrorCount                 = @Errors,
    ErrorMessage               = @ErrorMsg,
    TokenAdvancedTo            = @Token
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'       = $runId;       '@Status'      = $finalStatus
    '@MfcGroups'   = $counters.MfcGroups;    '@MfcMembers'  = $counters.MfcMembers
    '@FullAccess'  = $counters.FullAccess;   '@SendAs'      = $counters.SendAs
    '@SendOnBehalf'= $counters.SendOnBehalf; '@FolderPerms' = $counters.FolderPerms
    '@Skipped'     = $counters.Skipped;      '@Errors'      = $counters.Errors
    '@ErrorMsg'    = if ($counters.Errors -gt 0) { "$($counters.Errors) errors — check log" } else { $null }
    '@Token'       = $savedTimestamp
} | Out-Null

Write-LogSummary @{
    "Status"             = $finalStatus
    "MFC Groups"         = $counters.MfcGroups
    "MFC Members"        = $counters.MfcMembers
    "Full Access"        = $counters.FullAccess
    "Send-As"            = $counters.SendAs
    "Send-on-Behalf"     = $counters.SendOnBehalf
    "Folder Permissions" = $counters.FolderPerms
    "Mailboxes skipped"  = $counters.Skipped
    "Errors"             = $counters.Errors
    "Duration"           = "$([Math]::Round($overallSw.Elapsed.TotalMinutes, 2)) min"
    "Run ID"             = $runId
}

Disconnect-ExoSession
Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
