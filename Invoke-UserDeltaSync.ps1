<#
.SYNOPSIS
    Incremental delta sync — keeps the Users table aligned with Entra ID.

.DESCRIPTION
    Scheduled every 15 minutes. Uses Graph /users/delta token stored in
    DeltaTokens. Processes only users that changed since the last run.

    Steps:
      1. Reads users_delta token from DeltaTokens
      2. Calls /users/delta with that token — gets only changes since last run
      3. Soft-deletes removed users, inserts new users, updates changed users
         COALESCE on UPDATE prevents overwriting unchanged fields — Graph
         returns only the properties that changed, not the full object.
      4. Saves the new delta token ONLY after all processing succeeds
         (crash-safe — next run reprocesses rather than skips)
      5. On HTTP 410 Gone: deactivates token, exits. Re-run baseline to recover.

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Scheduled Task:
      Trigger : Daily, repeat every 15 minutes indefinitely
      Program : pwsh.exe
      Arguments: -NonInteractive -File "C:\M365PermSync\Invoke-UserDeltaSync.ps1"
      Settings: Do not start a new instance if already running.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-UserDeltaSync"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ──────────────────────────────────────────────────────────────
# Bootstrap
# ──────────────────────────────────────────────────────────────

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")  -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1") -Force
Import-Module (Join-Path $sharedPath "GraphHelpers.psm1")   -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")     -Force

try {
    $Config = Import-SyncConfig -Path $ConfigPath
    Initialize-Logging -Config $Config -ProcessName $scriptName
}
catch {
    Write-Host "FATAL BOOTSTRAP ERROR in $scriptName" -ForegroundColor Red
    Write-Host "Message : $($_.Exception.Message)"    -ForegroundColor White
    exit 1
}

Write-LogSection "Invoke-UserDeltaSync"
Write-LogInfo "Config   : $ConfigPath"
Write-LogInfo "Server   : $($Config.Database.Server)"
Write-LogInfo "Database : $($Config.Database.Name)"

# ──────────────────────────────────────────────────────────────
# Authentication
# ──────────────────────────────────────────────────────────────

Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Initialize-GraphContext -Config $Config
    Initialize-SqlContext   -Config $Config
    Write-LogInfo "Service principal connected"
}
catch {
    Write-LogError "Authentication failed" -ErrorRecord $_
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

$inserted    = 0
$updated     = 0
$softDeleted = 0
$processed   = 0
$errors      = 0
$newToken    = $null
$overallSw   = [System.Diagnostics.Stopwatch]::StartNew()

Write-LogInfo "Run ID: $runId"

# ──────────────────────────────────────────────────────────────
# Phase 1 — Fetch stored delta token
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 1 — Fetch delta token"

$storedToken = Get-DeltaToken -TokenName "users_delta"

if (-not $storedToken) {
    $msg = "No active users_delta token. Run Invoke-UserBaselineLoad first."
    Write-LogError $msg
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = $msg } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

Write-LogInfo "Token found (length: $($storedToken.Length))"

# ──────────────────────────────────────────────────────────────
# Phase 2 — Graph delta fetch
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 2 — Graph delta fetch"

$select      = "id,userPrincipalName,displayName,mail,accountEnabled," +
               "userType,onPremisesSyncEnabled,department,jobTitle,createdDateTime"
$url         = "$($Config.Graph.BaseUrl)/users/delta?`$select=$select&`$deltatoken=$storedToken"
$deltaResult = $null

try {
    $phaseSw     = [System.Diagnostics.Stopwatch]::StartNew()
    $deltaResult = Invoke-GraphDeltaQuery -Uri $url
    $phaseSw.Stop()
    Write-LogInfo "Delta fetch complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $($deltaResult.Objects.Count) changed object(s)"
}
catch {
    Write-LogError "Graph delta fetch failed" -ErrorRecord $_
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = "Graph delta fetch failed: $($_.Exception.Message)" } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

# ── HTTP 410 — token expired ──────────────────────────────────
if ($deltaResult.TokenExpired) {
    $msg = "Delta token expired (HTTP 410). Re-run Invoke-UserBaselineLoad to recover."
    Write-LogWarning $msg
    Disable-DeltaToken -TokenName "users_delta" -Reason $msg
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = $msg } | Out-Null
    Close-Logging -Status "Failed — token expired"
    exit 1
}

# ── No changes ────────────────────────────────────────────────
if ($deltaResult.Objects.Count -eq 0) {
    Write-LogInfo "No changes since last run"
    if ($deltaResult.DeltaToken) {
        Save-DeltaToken -TokenName "users_delta" -TokenValue $deltaResult.DeltaToken
        $newToken = $deltaResult.DeltaToken
    }
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Success',
    UsersProcessed=0, TokenAdvancedTo=@Token WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Token' = $newToken } | Out-Null
    Write-LogSummary @{ "Status" = "Success (no changes)"; "Duration" = "$([Math]::Round($overallSw.Elapsed.TotalSeconds,1))s" }
    Remove-OldLogFiles
    Close-Logging -Status "Success"
    exit 0
}

# ──────────────────────────────────────────────────────────────
# Phase 3 — Apply changes
# Deletions are processed first to avoid FK conflicts.
# COALESCE on UPDATE preserves existing values when Graph returns
# only the changed properties (partial delta objects).
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 3 — Apply changes"

$softDeleteSql = @"
UPDATE dbo.Users SET
    IsDeleted      = 1,
    DeletedAt      = SYSUTCDATETIME(),
    LastSyncedAt   = SYSUTCDATETIME(),
    LastModifiedAt = SYSUTCDATETIME(),
    SyncSource     = 'Delta',
    LastSyncRunId  = @RunId
WHERE UserId = @UserId AND IsDeleted = 0
"@

$upsertSql = @"
MERGE dbo.Users AS target
USING (SELECT @UserId AS UserId) AS source
ON target.UserId = source.UserId
WHEN MATCHED THEN UPDATE SET
    UserPrincipalName     = COALESCE(@UPN,                  UserPrincipalName),
    DisplayName           = COALESCE(@DisplayName,           DisplayName),
    Mail                  = COALESCE(@Mail,                  Mail),
    AccountEnabled        = COALESCE(@AccountEnabled,        AccountEnabled),
    UserType              = COALESCE(@UserType,              UserType),
    OnPremisesSyncEnabled = COALESCE(@OnPremisesSyncEnabled, OnPremisesSyncEnabled),
    Department            = COALESCE(@Department,            Department),
    JobTitle              = COALESCE(@JobTitle,              JobTitle),
    EntraCreatedDateTime  = COALESCE(@CreatedDateTime,       EntraCreatedDateTime),
    LastSyncedAt          = SYSUTCDATETIME(),
    LastModifiedAt        = SYSUTCDATETIME(),
    SyncSource            = 'Delta',
    LastSyncRunId         = @RunId,
    IsDeleted             = 0,
    DeletedAt             = NULL
WHEN NOT MATCHED THEN INSERT (
    UserId, UserPrincipalName, DisplayName, Mail, AccountEnabled,
    UserType, OnPremisesSyncEnabled, Department, JobTitle,
    EntraCreatedDateTime, SyncSource, LastSyncRunId
) VALUES (
    @UserId, @UPN, @DisplayName, @Mail, @AccountEnabled,
    @UserType, @OnPremisesSyncEnabled, @Department, @JobTitle,
    @CreatedDateTime, 'Delta', @RunId
)
"@

$existsSql = "SELECT COUNT(1) FROM dbo.Users WHERE UserId = @UserId"
$phaseSw   = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($u in $deltaResult.Objects) {
    try {
        if ($u.'@removed') {
            $rows = Invoke-SqlNonQuery -Query $softDeleteSql -Parameters @{
                '@UserId' = [guid]$u.id
                '@RunId'  = $runId
            }
            if ($rows -gt 0) {
                $softDeleted++
                Write-LogInfo "Soft-deleted: $($u.id)"
            }
            $processed++
            continue
        }

        $exists = Invoke-SqlScalar -Query $existsSql -Parameters @{ '@UserId' = [guid]$u.id }

        Invoke-SqlNonQuery -Query $upsertSql -Parameters @{
            '@UserId'                = [guid]$u.id
            '@UPN'                   = $u.userPrincipalName
            '@DisplayName'           = $u.displayName
            '@Mail'                  = $u.mail
            '@AccountEnabled'        = if ($null -eq $u.accountEnabled)        { $null } else { [bool]$u.accountEnabled }
            '@UserType'              = $u.userType
            '@OnPremisesSyncEnabled' = if ($null -eq $u.onPremisesSyncEnabled) { $null } else { [bool]$u.onPremisesSyncEnabled }
            '@Department'            = $u.department
            '@JobTitle'              = $u.jobTitle
            '@CreatedDateTime'       = if ($u.createdDateTime) { [datetime]$u.createdDateTime } else { $null }
            '@RunId'                 = $runId
        } | Out-Null

        if ([int]$exists -eq 0) { $inserted++ } else { $updated++ }
        $processed++
    }
    catch {
        $errors++
        Write-LogWarning "Failed to process user $($u.id): $($_.Exception.Message)"
    }
}

$phaseSw.Stop()
Write-LogInfo "Changes applied in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s"

# ──────────────────────────────────────────────────────────────
# Phase 4 — Advance delta token
# Saved ONLY after all processing completes.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 4 — Advance delta token"

if ($deltaResult.DeltaToken) {
    Save-DeltaToken -TokenName "users_delta" -TokenValue $deltaResult.DeltaToken
    $newToken = $deltaResult.DeltaToken
    Write-LogInfo "Delta token advanced"
}
else {
    Write-LogWarning "No new delta token returned — token NOT advanced"
}

# ──────────────────────────────────────────────────────────────
# Finalise
# ──────────────────────────────────────────────────────────────

$overallSw.Stop()
$finalStatus = if ($errors -gt 0) { "PartialFailure" } else { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt      = SYSUTCDATETIME(),
    Status           = @Status,
    UsersInserted    = @Inserted,
    UsersUpdated     = @Updated,
    UsersSoftDeleted = @Deleted,
    UsersProcessed   = @Processed,
    ErrorCount       = @Errors,
    TokenAdvancedTo  = @Token
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'     = $runId
    '@Status'    = $finalStatus
    '@Inserted'  = $inserted
    '@Updated'   = $updated
    '@Deleted'   = $softDeleted
    '@Processed' = $processed
    '@Errors'    = $errors
    '@Token'     = $newToken
} | Out-Null

Write-LogSummary @{
    "Status"      = $finalStatus
    "Inserted"    = $inserted
    "Updated"     = $updated
    "Soft-deleted"= $softDeleted
    "Errors"      = $errors
    "Processed"   = $processed
    "Duration"    = "$([Math]::Round($overallSw.Elapsed.TotalSeconds,1))s"
    "Run ID"      = $runId
}

Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
