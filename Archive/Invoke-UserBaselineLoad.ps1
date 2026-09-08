<#
.SYNOPSIS
    Full baseline load — pages all users from Graph into the Users table.

.DESCRIPTION
    Run once on initial deployment and whenever the delta token expires.

    Steps:
      1. Pages through Graph /users — fetches all user objects
      2. MERGEs every user into the Users table (idempotent)
         Uses a single open SQL connection for the whole loop — efficient
         at 40k+ users.
      3. Calls /users/delta to capture a starting delta token
      4. Saves the token so Invoke-UserDeltaSync can run from this point

.PARAMETER ConfigPath
    Path to config.xml. Defaults to config.xml in the same folder.

.NOTES
    Typical duration: 10–30 minutes for a 40,000 user tenant.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml")
)

$ErrorActionPreference = "Stop"
$scriptName            = "Invoke-UserBaselineLoad"
$sharedPath            = Join-Path $PSScriptRoot "shared"

# ──────────────────────────────────────────────────────────────
# Bootstrap
# ──────────────────────────────────────────────────────────────

Import-Module (Join-Path $sharedPath "ConfigHelpers.psm1")     -Force
Import-Module (Join-Path $sharedPath "LoggingHelpers.psm1")    -Force
Import-Module (Join-Path $sharedPath "GraphHelpers.psm1")      -Force
Import-Module (Join-Path $sharedPath "SqlHelpers.psm1")        -Force
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

Write-LogSection "Invoke-UserBaselineLoad"
Write-LogInfo "Config   : $ConfigPath"
Write-LogInfo "Server   : $($Config.Database.Server)"
Write-LogInfo "Database : $($Config.Database.Name)"
Write-LogInfo "Tenant   : $($Config.Authentication.TenantId)"

Write-LogSection "Dependency Validation"
Initialize-ModuleDependencies -Config $Config

# ──────────────────────────────────────────────────────────────
# Authentication
# ──────────────────────────────────────────────────────────────

Write-LogSection "Authentication"

try {
    Connect-SyncServicePrincipal -Config $Config
    Write-LogInfo "Service principal connected"
    Initialize-GraphContext -Config $Config
    Initialize-SqlContext   -Config $Config
}
catch {
    Write-LogError "Authentication failed" -ErrorRecord $_
    Close-Logging -Status "Failed"
    exit 1
}

# ──────────────────────────────────────────────────────────────
# Start run tracking
# ──────────────────────────────────────────────────────────────

$runId = [guid]::NewGuid()
Invoke-SqlNonQuery -Query @"
INSERT INTO dbo.SyncLog (RunId, FunctionName, StartedAt, Status)
VALUES (@RunId, @Function, SYSUTCDATETIME(), 'Running')
"@ -Parameters @{ '@RunId' = $runId; '@Function' = $scriptName } | Out-Null

$inserted      = 0
$errors        = 0
$processed     = 0
$capturedToken = $null
$overallSw     = [System.Diagnostics.Stopwatch]::StartNew()

Write-LogInfo "Run ID: $runId"

# ──────────────────────────────────────────────────────────────
# Phase 1 — Page through Graph /users
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 1 — Graph user fetch"

$allUsers = $null
try {
    $pageSize = [int]($Config.Graph.PageSize ?? 999)
    $select   = "id,userPrincipalName,displayName,mail,accountEnabled," +
                "userType,onPremisesSyncEnabled,department,jobTitle,createdDateTime"
    $url      = "$($Config.Graph.BaseUrl)/users?`$select=$select&`$top=$pageSize"

    $phaseSw = [System.Diagnostics.Stopwatch]::StartNew()
    $allUsers = Invoke-GraphPagedRequest -Uri $url
    $phaseSw.Stop()

    Write-LogInfo "Graph returned $($allUsers.Count) users in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s"
}
catch {
    Write-LogError "Phase 1 (Graph fetch) failed — aborting" -ErrorRecord $_
    Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET CompletedAt=SYSUTCDATETIME(), Status='Failed',
    ErrorCount=1, ErrorMessage=@Msg WHERE RunId=@RunId
"@ -Parameters @{ '@RunId' = $runId; '@Msg' = "Phase 1 failed: $($_.Exception.Message)" } | Out-Null
    Close-Logging -Status "Failed"
    exit 1
}

# ──────────────────────────────────────────────────────────────
# Phase 2 — MERGE all users into the Users table
#
# One connection held open for the entire loop — efficient at scale.
# Each user is processed individually so a single failure does not
# stop the rest of the batch.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 2 — SQL upsert"

$mergeSql = @"
MERGE dbo.Users AS target
USING (SELECT @UserId AS UserId) AS source
ON target.UserId = source.UserId
WHEN MATCHED THEN UPDATE SET
    UserPrincipalName     = @UPN,
    DisplayName           = @DisplayName,
    Mail                  = @Mail,
    AccountEnabled        = @AccountEnabled,
    UserType              = @UserType,
    OnPremisesSyncEnabled = @OnPremisesSyncEnabled,
    Department            = @Department,
    JobTitle              = @JobTitle,
    EntraCreatedDateTime  = @CreatedDateTime,
    LastSyncedAt          = SYSUTCDATETIME(),
    LastModifiedAt        = SYSUTCDATETIME(),
    SyncSource            = 'Baseline',
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
    @CreatedDateTime, 'Baseline', @RunId
);
"@

$progressInterval = 5000
$phaseSw          = [System.Diagnostics.Stopwatch]::StartNew()

$conn = Open-SqlConnection
try {
    foreach ($u in $allUsers) {
        try {
            Invoke-SqlNonQuery -Connection $conn -Query $mergeSql -Parameters @{
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
            $inserted++
        }
        catch {
            $errors++
            Write-LogWarning "MERGE failed for $($u.userPrincipalName): $($_.Exception.Message)"
        }
        $processed++

        if ($processed % $progressInterval -eq 0) {
            $rate = [Math]::Round($processed / $phaseSw.Elapsed.TotalMinutes, 0)
            Write-LogInfo "Progress: $processed / $($allUsers.Count) | Errors: $errors | $rate/min"
        }
    }
}
finally {
    $conn.Dispose()
}

$phaseSw.Stop()
Write-LogInfo "SQL upsert complete in $([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s — $inserted upserted, $errors errors"

# ──────────────────────────────────────────────────────────────
# Phase 3 — Capture starting delta token
# Uses $select=id only — we only need the deltaLink, not the payload.
# ──────────────────────────────────────────────────────────────

Write-LogSection "Phase 3 — Capture delta token"

try {
    $phaseSw     = [System.Diagnostics.Stopwatch]::StartNew()
    $deltaResult = Invoke-GraphDeltaQuery -Uri "$($Config.Graph.BaseUrl)/users/delta?`$select=id&`$top=999"
    $phaseSw.Stop()

    if ($deltaResult.TokenExpired -or -not $deltaResult.DeltaToken) {
        throw "Delta initialisation did not return a valid token"
    }

    Save-DeltaToken -TokenName "users_delta" -TokenValue $deltaResult.DeltaToken
    $capturedToken = $deltaResult.DeltaToken
    Write-LogInfo "Delta token saved ($([Math]::Round($phaseSw.Elapsed.TotalSeconds,1))s, $($deltaResult.Objects.Count) objects consumed)"
}
catch {
    $errors++
    Write-LogError "Phase 3 (delta token capture) failed — data loaded but delta sync will not work" -ErrorRecord $_
}

# ──────────────────────────────────────────────────────────────
# Finalise
# ──────────────────────────────────────────────────────────────

$overallSw.Stop()
$finalStatus = if ($errors -gt 0 -and -not $capturedToken) { "Failed" }
               elseif ($errors -gt 0)                      { "PartialFailure" }
               else                                         { "Success" }

Invoke-SqlNonQuery -Query @"
UPDATE dbo.SyncLog SET
    CompletedAt      = SYSUTCDATETIME(),
    Status           = @Status,
    UsersInserted    = @Inserted,
    UsersProcessed   = @Processed,
    ErrorCount       = @Errors,
    TokenAdvancedTo  = @Token
WHERE RunId = @RunId
"@ -Parameters @{
    '@RunId'     = $runId
    '@Status'    = $finalStatus
    '@Inserted'  = $inserted
    '@Processed' = $processed
    '@Errors'    = $errors
    '@Token'     = $capturedToken
} | Out-Null

Write-LogSummary @{
    "Status"          = $finalStatus
    "Users processed" = $processed
    "Users upserted"  = $inserted
    "Errors"          = $errors
    "Token captured"  = ($null -ne $capturedToken)
    "Duration"        = "$([Math]::Round($overallSw.Elapsed.TotalMinutes, 2)) min"
    "Run ID"          = $runId
}

Remove-OldLogFiles
Close-Logging -Status $finalStatus

if ($finalStatus -eq "Failed") { exit 1 }
