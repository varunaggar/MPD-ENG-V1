<#
.SYNOPSIS
    Azure SQL Database helpers — certificate-based authentication.

.DESCRIPTION
    Infrastructure only: token acquisition, connection management,
    and parameterised query execution.

    REMOVED from previous version:
      - Invoke-SqlBatch         → scripts use Open-SqlConnection directly
      - Start-SyncLogEntry      → scripts write SyncLog inline
      - Complete-SyncLogEntry   → scripts write SyncLog inline

    ADDED:
      - Open-SqlConnection      → public; returns a live SqlConnection.
                                   Caller is responsible for Dispose().

    UPDATED:
      - Invoke-SqlNonQuery      → accepts optional -Connection / -Transaction
      - Invoke-SqlScalar        → accepts optional -Connection
      - Invoke-SqlQuery         → accepts optional -Connection

    BATCHING PATTERN FOR SCRIPTS:
        $conn = Open-SqlConnection
        try {
            $tx = $conn.BeginTransaction()
            Invoke-SqlNonQuery -Connection $conn -Transaction $tx -Query $q1 -Parameters $p1
            Invoke-SqlNonQuery -Connection $conn -Transaction $tx -Query $q2 -Parameters $p2
            $tx.Commit()
        }
        catch { $tx.Rollback(); throw }
        finally { $conn.Dispose() }

    For bulk inserts without transaction (each row commits independently):
        $conn = Open-SqlConnection
        try {
            foreach ($item in $items) {
                Invoke-SqlNonQuery -Connection $conn -Query $insertSql -Parameters @{...}
            }
        }
        finally { $conn.Dispose() }
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

$script:SqlServer         = $null
$script:SqlDatabase       = $null
$script:SqlTargetTenantId = $null
$script:ConnectionTimeout = 30
$script:CommandTimeout    = 120

$script:SqlTokenCache = @{
    Token     = $null
    ExpiresAt = [datetime]::MinValue
}

# ──────────────────────────────────────────────────────────────
# Public: Initialize-SqlContext
# ──────────────────────────────────────────────────────────────

function Initialize-SqlContext {
    <#
    .SYNOPSIS
        Stores SQL connection settings from config. Call once per script.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    $script:SqlServer         = $Config.Database.Server
    $script:SqlDatabase       = $Config.Database.Name
    $script:SqlTargetTenantId = $Config.Database.TargetTenantId
    $script:ConnectionTimeout = if ($Config.Database.ConnectionTimeoutSec) { [int]$Config.Database.ConnectionTimeoutSec } else { 30 }
    $script:CommandTimeout    = if ($Config.Database.CommandTimeoutSec)    { [int]$Config.Database.CommandTimeoutSec }    else { 120 }

    $script:SqlTokenCache.Token     = $null
    $script:SqlTokenCache.ExpiresAt = [datetime]::MinValue

    Write-LogInfo "SQL context initialised (Server=$($script:SqlServer), DB=$($script:SqlDatabase))"
}

# ──────────────────────────────────────────────────────────────
# Private: Get-SqlAccessToken
# ──────────────────────────────────────────────────────────────

function Get-SqlAccessToken {
    param([switch]$ForceRefresh)

    $now = [datetime]::UtcNow

    if (-not $ForceRefresh -and
        $script:SqlTokenCache.Token -and
        $script:SqlTokenCache.ExpiresAt -gt $now.AddMinutes(2)) {
        return $script:SqlTokenCache.Token
    }

    try {
        $params = @{
            ResourceUrl = 'https://database.windows.net/'
            ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($script:SqlTargetTenantId)) {
            $params.TenantId = $script:SqlTargetTenantId
        }

        Write-LogInfo "Acquiring SQL access token..."
        $tokenInfo = Get-AzAccessToken @params

        $plain = if ($tokenInfo.Token -is [System.Security.SecureString]) {
            $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenInfo.Token)
            try   { [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
            finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
        } else {
            $tokenInfo.Token
        }

        $script:SqlTokenCache.Token     = $plain
        $script:SqlTokenCache.ExpiresAt = $tokenInfo.ExpiresOn.UtcDateTime.AddMinutes(-5)

        Write-LogInfo "SQL token acquired. Valid until $($script:SqlTokenCache.ExpiresAt) UTC"
        return $plain
    }
    catch {
        throw "Failed to acquire SQL access token: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Private: New-SqlConnection
# Opens and returns an authenticated SqlConnection.
# ──────────────────────────────────────────────────────────────

function New-SqlConnection {
    $connStr = "Server=$($script:SqlServer);Database=$($script:SqlDatabase);" +
               "Encrypt=True;TrustServerCertificate=False;" +
               "Connection Timeout=$($script:ConnectionTimeout)"

    $conn                  = [System.Data.SqlClient.SqlConnection]::new()
    $conn.ConnectionString = $connStr
    $token                 = Get-SqlAccessToken
    $conn.AccessToken      = $token

    try {
        $jwt      = ConvertFrom-JwtToken -Token $token
        $identity = $jwt.upn ?? $jwt.unique_name ?? $jwt.appid ?? $jwt.oid
        Write-LogInfo "SQL login to [$($script:SqlDatabase)] as: $identity"
    } catch {}

    $conn.Open()
    return $conn
}

# ──────────────────────────────────────────────────────────────
# Private: Set-SqlParameters
# Adds parameters to a SqlCommand; converts $null → DBNull.
# ──────────────────────────────────────────────────────────────

function Set-SqlParameters {
    param(
        [System.Data.SqlClient.SqlCommand]$Command,
        [hashtable]$Parameters
    )
    if (-not $Parameters) { return }
    foreach ($p in $Parameters.GetEnumerator()) {
        $value = if ($null -eq $p.Value) { [System.DBNull]::Value } else { $p.Value }
        [void]$Command.Parameters.AddWithValue($p.Key, $value)
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Open-SqlConnection
# Returns an open, authenticated SqlConnection.
# Caller is responsible for calling .Dispose() when done.
# Use this to hold a connection open across multiple SQL calls.
# ──────────────────────────────────────────────────────────────

function Open-SqlConnection {
    <#
    .SYNOPSIS
        Opens and returns an authenticated Azure SQL connection.
        Caller must call .Dispose() when done.
    .EXAMPLE
        $conn = Open-SqlConnection
        try {
            Invoke-SqlNonQuery -Connection $conn -Query $q -Parameters $p
        }
        finally { $conn.Dispose() }
    #>
    [CmdletBinding()]
    param()
    return New-SqlConnection
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-SqlNonQuery
# Executes INSERT / UPDATE / DELETE / MERGE.
# Returns rows affected.
#
# If -Connection is provided, uses it without closing it.
# If -Connection is not provided, opens and closes its own.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlNonQuery {
    <#
    .SYNOPSIS
        Executes a non-query SQL statement. Returns rows affected.
    .PARAMETER Query
        The SQL to execute.
    .PARAMETER Parameters
        Hashtable of named parameters, e.g. @{ '@UserId' = $id }.
    .PARAMETER Connection
        Optional open SqlConnection. If provided, the connection is NOT
        closed after the call — the caller manages its lifecycle.
    .PARAMETER Transaction
        Optional SqlTransaction. Must belong to the provided Connection.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters  = @{},
        [System.Data.SqlClient.SqlConnection]$Connection   = $null,
        [System.Data.SqlClient.SqlTransaction]$Transaction = $null,
        [int]$TimeoutSec = 0
    )

    if ($TimeoutSec -le 0) { $TimeoutSec = $script:CommandTimeout }

    $ownConn = ($null -eq $Connection)
    $conn    = if ($ownConn) { New-SqlConnection } else { $Connection }

    try {
        $cmd                = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        if ($Transaction) { $cmd.Transaction = $Transaction }
        Set-SqlParameters -Command $cmd -Parameters $Parameters
        return $cmd.ExecuteNonQuery()
    }
    catch {
        $snippet = $Query.Substring(0, [Math]::Min(150, $Query.Length)).Trim() -replace '\s+', ' '
        throw "Invoke-SqlNonQuery failed: $($_.Exception.Message) | Query: $snippet"
    }
    finally {
        if ($ownConn -and $conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-SqlScalar
# Returns a single scalar value, or $null.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlScalar {
    <#
    .SYNOPSIS
        Executes a query and returns the first column of the first row.
    .PARAMETER Connection
        Optional open SqlConnection.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters = @{},
        [System.Data.SqlClient.SqlConnection]$Connection = $null,
        [int]$TimeoutSec = 30
    )

    $ownConn = ($null -eq $Connection)
    $conn    = if ($ownConn) { New-SqlConnection } else { $Connection }

    try {
        $cmd                = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        Set-SqlParameters -Command $cmd -Parameters $Parameters

        $result = $cmd.ExecuteScalar()
        return if ($null -eq $result -or $result -is [System.DBNull]) { $null } else { $result }
    }
    catch {
        $snippet = $Query.Substring(0, [Math]::Min(150, $Query.Length)).Trim() -replace '\s+', ' '
        throw "Invoke-SqlScalar failed: $($_.Exception.Message) | Query: $snippet"
    }
    finally {
        if ($ownConn -and $conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-SqlQuery
# Executes a SELECT. Returns an array of row objects.
# ──────────────────────────────────────────────────────────────

function Invoke-SqlQuery {
    <#
    .SYNOPSIS
        Executes a SELECT query. Returns an array of PSCustomObjects.
    .PARAMETER Connection
        Optional open SqlConnection.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query,
        [hashtable]$Parameters = @{},
        [System.Data.SqlClient.SqlConnection]$Connection = $null,
        [int]$TimeoutSec = 0
    )

    if ($TimeoutSec -le 0) { $TimeoutSec = $script:CommandTimeout }

    $ownConn = ($null -eq $Connection)
    $conn    = if ($ownConn) { New-SqlConnection } else { $Connection }

    try {
        $cmd                = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $TimeoutSec
        Set-SqlParameters -Command $cmd -Parameters $Parameters

        $dt      = [System.Data.DataTable]::new()
        $adapter = [System.Data.SqlClient.SqlDataAdapter]::new($cmd)
        [void]$adapter.Fill($dt)

        return $dt.Rows | ForEach-Object {
            $row = [ordered]@{}
            foreach ($col in $dt.Columns) { $row[$col.ColumnName] = $_[$col.ColumnName] }
            [PSCustomObject]$row
        }
    }
    catch {
        $snippet = $Query.Substring(0, [Math]::Min(150, $Query.Length)).Trim() -replace '\s+', ' '
        throw "Invoke-SqlQuery failed: $($_.Exception.Message) | Query: $snippet"
    }
    finally {
        if ($ownConn -and $conn) { $conn.Close(); $conn.Dispose() }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Get-DeltaToken
# Returns the active token value for a named token, or $null.
# ──────────────────────────────────────────────────────────────

function Get-DeltaToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)] [string]$TokenName)

    return Invoke-SqlScalar -Query @"
SELECT TokenValue FROM dbo.DeltaTokens
WHERE TokenName = @Name AND IsActive = 1
"@ -Parameters @{ '@Name' = $TokenName }
}

# ──────────────────────────────────────────────────────────────
# Public: Save-DeltaToken
# Upserts (insert or update) a named delta token.
# ──────────────────────────────────────────────────────────────

function Save-DeltaToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$TokenName,
        [Parameter(Mandatory)] [string]$TokenValue
    )

    Invoke-SqlNonQuery -Query @"
MERGE dbo.DeltaTokens AS target
USING (SELECT @Name AS TokenName) AS source
ON target.TokenName = source.TokenName
WHEN MATCHED THEN UPDATE SET
    TokenValue         = @Value,
    UpdatedAt          = SYSUTCDATETIME(),
    IsActive           = 1,
    DeactivatedAt      = NULL,
    DeactivationReason = NULL
WHEN NOT MATCHED THEN INSERT
    (TokenName, TokenValue, CreatedAt, UpdatedAt, IsActive)
VALUES
    (@Name, @Value, SYSUTCDATETIME(), SYSUTCDATETIME(), 1);
"@ -Parameters @{ '@Name' = $TokenName; '@Value' = $TokenValue } | Out-Null
}

# ──────────────────────────────────────────────────────────────
# Public: Disable-DeltaToken
# Marks an active token as inactive with a reason.
# ──────────────────────────────────────────────────────────────

function Disable-DeltaToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$TokenName,
        [string]$Reason = 'Deactivated'
    )

    Invoke-SqlNonQuery -Query @"
UPDATE dbo.DeltaTokens SET
    IsActive           = 0,
    DeactivatedAt      = SYSUTCDATETIME(),
    DeactivationReason = @Reason
WHERE TokenName = @Name
"@ -Parameters @{ '@Name' = $TokenName; '@Reason' = $Reason } | Out-Null
}

# ──────────────────────────────────────────────────────────────
# Public: ConvertFrom-JwtToken
# Decodes a JWT payload without signature verification.
# ──────────────────────────────────────────────────────────────

function ConvertFrom-JwtToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Token)

    $plain = if ($Token -is [System.Security.SecureString]) {
        [System.Net.NetworkCredential]::new('', $Token).Password
    } else { $Token }

    if ([string]::IsNullOrWhiteSpace($plain)) { return $null }

    $parts = $plain.Split('.')
    if ($parts.Count -ne 3) { throw "Invalid JWT: expected 3 parts." }

    $b64 = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($b64.Length % 4) {
        2 { $b64 += '==' }
        3 { $b64 += '='  }
    }

    return [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64)) |
           ConvertFrom-Json
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-SqlContext',
    'Open-SqlConnection',
    'Invoke-SqlNonQuery',
    'Invoke-SqlScalar',
    'Invoke-SqlQuery',
    'Get-DeltaToken',
    'Save-DeltaToken',
    'Disable-DeltaToken',
    'ConvertFrom-JwtToken'
)
