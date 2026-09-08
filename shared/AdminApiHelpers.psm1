<#
.SYNOPSIS
    Exchange Online Admin API helpers — REST-based cmdlet execution.

.DESCRIPTION
    Provides token acquisition and POST-based request execution
    for the EXO Admin API v2.0.

    Infrastructure only: token management, single POST with retry,
    and paged POST following @odata.nextLink.

    REMOVED from previous version:
      - Get-MfcGroupMembers          → scripts call Invoke-AdminApiPagedRequest directly
      - Get-MailboxFolderPermissions → scripts call Invoke-AdminApiPagedRequest directly

    Scripts build CmdletInput request bodies and process response.value
    themselves. Use Get-AnchorMailboxHeader to build the X-AnchorMailbox value.

.NOTES
    Dependencies: Az.Accounts (established by Connect-SyncServicePrincipal)
    Config consumed from config.xml:
        Authentication.TenantId
        ExchangeOnline.Organisation         (for X-AnchorMailbox app-only routing)
        AdminApi.BaseUrl
        AdminApi.MaxRetries
        AdminApi.ThrottleBackoffMaxSec

    Base URL pattern:
        {AdminApi.BaseUrl}/{TenantId}/{EndpointName}[?$select=...]
    Example:
        https://outlook.office365.com/adminapi/v2.0/{guid}/DistributionGroupMember
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

$script:AdminApiConfig  = $null
$script:AdminApiBaseUrl = $null   # Resolved once in Initialize-AdminApiContext

$script:TokenCache = @{
    Token     = $null
    ExpiresAt = [datetime]::MinValue
}

# Well-known system mailbox GUID — identical across all M365 tenants.
# Used for X-AnchorMailbox routing in app-only flows that have no
# specific mailbox target (e.g. DistributionGroupMember queries).
# Ref: https://learn.microsoft.com/en-us/exchange/reference/admin-api-get-started
$script:SystemMailboxGuid = 'bb558c35-97f1-4cb9-8ff7-d53741dc928c'

# ──────────────────────────────────────────────────────────────
# Public: Initialize-AdminApiContext
# Called once per script, after Connect-SyncServicePrincipal.
# ──────────────────────────────────────────────────────────────

function Initialize-AdminApiContext {
    <#
    .SYNOPSIS
        Initialises module-scoped Admin API configuration.
    .PARAMETER Config
        The config object from Import-SyncConfig.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    $script:AdminApiConfig = $Config

    # Reset token cache — never reuse tokens across script runs
    $script:TokenCache.Token     = $null
    $script:TokenCache.ExpiresAt = [datetime]::MinValue

    # Validate required config
    $tenantId = $Config.Authentication.TenantId
    $baseUrl  = $Config.AdminApi.BaseUrl
    $org      = $Config.ExchangeOnline.Organisation

    foreach ($check in @(
        @{ Name = 'Authentication.TenantId';      Value = $tenantId }
        @{ Name = 'AdminApi.BaseUrl';             Value = $baseUrl  }
        @{ Name = 'ExchangeOnline.Organisation';  Value = $org      }
    )) {
        if ([string]::IsNullOrWhiteSpace($check.Value)) {
            throw "AdminApiHelpers: config.xml '$($check.Name)' is not set."
        }
    }

    # Construct and cache the tenant-scoped base URL
    # e.g. https://outlook.office365.com/adminapi/v2.0/{tenantId}
    $script:AdminApiBaseUrl = "$($baseUrl.TrimEnd('/'))/$tenantId"

    Write-LogInfo "Admin API context initialised (BaseUrl=$($script:AdminApiBaseUrl))"
}

# ──────────────────────────────────────────────────────────────
# Public: Get-AdminApiToken
# Returns a valid EXO Admin API access token.
# Reuses a cached token until 2 minutes before expiry.
# Az context must already be established via Connect-SyncServicePrincipal.
# ──────────────────────────────────────────────────────────────

function Get-AdminApiToken {
    <#
    .SYNOPSIS
        Returns a valid EXO Admin API bearer token, refreshing if needed.
    .PARAMETER ForceRefresh
        Bypass the cache and acquire a new token immediately.
    #>
    [CmdletBinding()]
    param(
        [switch]$ForceRefresh
    )

    $now = [datetime]::UtcNow

    if (-not $ForceRefresh -and
        $script:TokenCache.Token -and
        $script:TokenCache.ExpiresAt -gt $now.AddMinutes(2)) {
        return $script:TokenCache.Token
    }

    Write-LogInfo "Acquiring EXO Admin API access token..."

    try {
        # Same Az context as Graph/SQL — just a different resource URL
        $tokenInfo = Get-AzAccessToken `
            -ResourceUrl "https://outlook.office365.com/" `
            -ErrorAction Stop

        # Handle SecureString return from modern Az.Accounts versions
        $plainToken = if ($tokenInfo.Token -is [System.Security.SecureString]) {
            $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenInfo.Token)
            try   { [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr) }
            finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) }
        }
        else {
            $tokenInfo.Token
        }

        $script:TokenCache.Token     = $plainToken
        $script:TokenCache.ExpiresAt = $tokenInfo.ExpiresOn.UtcDateTime.AddMinutes(-10)

        Write-LogInfo "Admin API token acquired. Valid until $($script:TokenCache.ExpiresAt) UTC"
        return $plainToken
    }
    catch {
        throw "Failed to acquire EXO Admin API token: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Private: Get-AnchorMailboxHeader
# Builds the X-AnchorMailbox value for a given routing context.
#
# Two modes:
#   Mailbox  — targets a specific mailbox (folder permission calls)
#              Value: UPN:<mailboxUpn>
#   AppOnly  — no specific mailbox target (group membership calls)
#              Value: APP:SystemMailbox{<guid>}@<org>
# ──────────────────────────────────────────────────────────────

function Get-AnchorMailboxHeader {
    param(
        [ValidateSet('Mailbox', 'AppOnly')]
        [string]$Mode,

        # Required when Mode = Mailbox
        [string]$MailboxUpn = $null
    )

    switch ($Mode) {
        'Mailbox' {
            if ([string]::IsNullOrWhiteSpace($MailboxUpn)) {
                throw "Get-AnchorMailboxHeader: MailboxUpn is required when Mode = Mailbox"
            }
            return "UPN:$MailboxUpn"
        }
        'AppOnly' {
            $org = $script:AdminApiConfig.ExchangeOnline.Organisation
            return "APP:SystemMailbox{$script:SystemMailboxGuid}@$org"
        }
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-AdminApiRequest
# Single POST to an Admin API endpoint with retry logic.
# Returns the raw response object from Invoke-RestMethod.
#
# Retry behaviour mirrors GraphHelpers:
#   429 → Retry-After header backoff (capped at ThrottleBackoffMaxSec)
#   401 → token refresh + one retry
#   5xx → exponential backoff
#   Other → throw immediately
# ──────────────────────────────────────────────────────────────

function Invoke-AdminApiRequest {
    <#
    .SYNOPSIS
        Single POST to an EXO Admin API endpoint with retry.
    .PARAMETER Endpoint
        Endpoint name, e.g. 'DistributionGroupMember', 'MailboxFolderPermission'.
    .PARAMETER CmdletName
        The EXO cmdlet to invoke, e.g. 'Get-DistributionGroupMember'.
    .PARAMETER Parameters
        Hashtable of cmdlet parameters.
    .PARAMETER AnchorMailbox
        Pre-built X-AnchorMailbox header value from Get-AnchorMailboxHeader.
    .PARAMETER Select
        Optional comma-separated list of properties for ?$select= query param.
    .PARAMETER OverrideUrl
        Full URL override — used for @odata.nextLink pagination continuation.
        When set, Endpoint/CmdletName/Parameters/Select are ignored for the URL
        but CmdletName and Parameters are still used for the POST body.
    .PARAMETER TimeoutSec
        HTTP timeout in seconds. Default 120 — Admin API can be slow under load.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Endpoint,
        [Parameter(Mandatory)] [string]$CmdletName,
        [Parameter(Mandatory)] [hashtable]$Parameters,
        [Parameter(Mandatory)] [string]$AnchorMailbox,
        [string]$Select        = $null,
        [string]$OverrideUrl   = $null,
        [int]$TimeoutSec       = 120
    )

    $maxRetries  = if ($script:AdminApiConfig.AdminApi.MaxRetries) {
                       [int]$script:AdminApiConfig.AdminApi.MaxRetries } else { 5 }
    $throttleMax = if ($script:AdminApiConfig.AdminApi.ThrottleBackoffMaxSec) {
                       [int]$script:AdminApiConfig.AdminApi.ThrottleBackoffMaxSec } else { 300 }

    # Build URL — use override for nextLink pagination, build fresh otherwise
    $url = if ($OverrideUrl) {
        $OverrideUrl
    }
    else {
        $built = "$script:AdminApiBaseUrl/$Endpoint"
        if (-not [string]::IsNullOrWhiteSpace($Select)) {
            $built += "?`$select=$Select"
        }
        $built
    }

    # Build CmdletInput body — same body used for both initial request and nextLink pages
    $body = @{
        CmdletInput = @{
            CmdletName = $CmdletName
            Parameters = $Parameters
        }
    } | ConvertTo-Json -Depth 5

    $attempt   = 0
    $lastError = $null

    while ($attempt -lt $maxRetries) {
        $attempt++

        try {
            $token   = Get-AdminApiToken
            $headers = @{
                Authorization   = "Bearer $token"
                'Content-Type'  = 'application/json'
                'X-AnchorMailbox' = $AnchorMailbox
            }

            return Invoke-RestMethod `
                -Uri         $url `
                -Method      POST `
                -Headers     $headers `
                -Body        $body `
                -TimeoutSec  $TimeoutSec `
                -ErrorAction Stop
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

            # Extract API error message from the response body
            $apiMessage = Get-AdminApiErrorMessage -ErrorRecord $_

            # 429 — throttled: honour Retry-After, cap at max
            if ($statusCode -eq 429) {
                $wait = 30
                try {
                    $h = $_.Exception.Response.Headers['Retry-After']
                    if ($h) { $wait = [int]$h }
                }
                catch {}
                $wait = [Math]::Min($wait, $throttleMax)
                Write-LogWarning "Admin API throttled (429). Waiting ${wait}s before retry $attempt/$maxRetries. Endpoint: $Endpoint$apiMessage"
                Start-Sleep -Seconds $wait
                $lastError = $_
                continue
            }

            # 401 — force token refresh, one retry only
            if ($statusCode -eq 401 -and $attempt -eq 1) {
                Write-LogWarning "Admin API returned 401. Refreshing token and retrying. Endpoint: $Endpoint$apiMessage"
                Get-AdminApiToken -ForceRefresh | Out-Null
                $lastError = $_
                continue
            }

            # 5xx — transient: exponential backoff
            if ($statusCode -ge 500 -and $statusCode -le 599) {
                $wait = [Math]::Min([Math]::Pow(2, $attempt), 60)
                Write-LogWarning "Admin API returned $statusCode. Waiting ${wait}s before retry $attempt/$maxRetries. Endpoint: $Endpoint$apiMessage"
                Start-Sleep -Seconds $wait
                $lastError = $_
                continue
            }

            # Anything else — throw immediately with context
            throw "Admin API error (HTTP $statusCode) on $Endpoint [$CmdletName]: $($_.Exception.Message)$apiMessage"
        }
    }

    throw "Admin API [$CmdletName] on $Endpoint failed after $maxRetries attempts. Last error: $($lastError.Exception.Message)"
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-AdminApiPagedRequest
# Fetches all pages from an Admin API endpoint, following
# @odata.nextLink automatically.
#
# Returns a flat array of all result objects.
#
# IMPORTANT — nextLink expiry:
#   Admin API nextLink URLs are valid for only 5-10 minutes.
#   This function fetches all pages immediately with no processing
#   between pages. Do not interleave SQL writes between page calls.
#   If a page takes more than 4 minutes, a warning is emitted.
# ──────────────────────────────────────────────────────────────

function Invoke-AdminApiPagedRequest {
    <#
    .SYNOPSIS
        Fetches all pages from an EXO Admin API endpoint.
    .PARAMETER Endpoint
        Endpoint name, e.g. 'DistributionGroupMember'.
    .PARAMETER CmdletName
        The EXO cmdlet to invoke.
    .PARAMETER Parameters
        Hashtable of cmdlet parameters. Do not include ResultSize here —
        pass it via the Parameters hashtable directly.
    .PARAMETER AnchorMailbox
        Pre-built X-AnchorMailbox header value.
    .PARAMETER Select
        Optional comma-separated property list for ?$select=.
    .OUTPUTS
        Array of result objects from the value collection.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Endpoint,
        [Parameter(Mandatory)] [string]$CmdletName,
        [Parameter(Mandatory)] [hashtable]$Parameters,
        [Parameter(Mandatory)] [string]$AnchorMailbox,
        [string]$Select = $null
    )

    $allObjects  = [System.Collections.Generic.List[object]]::new()
    $nextLinkUrl = $null
    $page        = 0

    do {
        $page++
        $pageStart = [datetime]::UtcNow

        $invokeParams = @{
            Endpoint       = $Endpoint
            CmdletName     = $CmdletName
            Parameters     = $Parameters
            AnchorMailbox  = $AnchorMailbox
            Select         = $Select
        }

        # From page 2 onward, POST to the nextLink URL with the same body
        if ($nextLinkUrl) {
            $invokeParams['OverrideUrl'] = $nextLinkUrl
        }

        Write-LogInfo "Admin API [$CmdletName] — fetching page $page..."

        $response = Invoke-AdminApiRequest @invokeParams

        if ($response.value) {
            $allObjects.AddRange([object[]]$response.value)
        }

        $nextLinkUrl = $response.'@odata.nextLink'

        # Warn if this page took long enough that the nextLink may be close to expiry
        $elapsed = ([datetime]::UtcNow - $pageStart).TotalSeconds
        if ($nextLinkUrl -and $elapsed -gt 240) {
            Write-LogWarning ("Admin API [$CmdletName] page $page took {0:F0}s. " +
                "nextLink expires in 5-10 min — subsequent pages may fail if processing is slow.") -f $elapsed
        }

    } while ($nextLinkUrl)

    Write-LogInfo "Admin API [$CmdletName] complete: $page page(s), $($allObjects.Count) total objects"
    return $allObjects.ToArray()
}

# ──────────────────────────────────────────────────────────────
# Private: Get-AdminApiErrorMessage
# ──────────────────────────────────────────────────────────────

function Get-AdminApiErrorMessage {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    try {
        $stream = $ErrorRecord.Exception.Response.GetResponseStream()
        if ($null -eq $stream) { return '' }
        $body = [System.IO.StreamReader]::new($stream).ReadToEnd()
        if ($body -match '"message"\s*:\s*"([^"]+)"') { return " | API: $($Matches[1])" }
        if ($body.Length -gt 0) { return " | API: $($body.Substring(0, [Math]::Min(200, $body.Length)))" }
    }
    catch {}
    return ''
}

# ──────────────────────────────────────────────────────────────
# Exports
# Infrastructure only: token, context, HTTP POST with retry/paging.
# Business logic (GetMfcGroupMembers, GetMailboxFolderPermissions)
# removed — scripts build their own request bodies and process responses.
#
# Scripts call Invoke-AdminApiRequest (single POST) or
# Invoke-AdminApiPagedRequest (POST + follow nextLink) directly,
# building the CmdletInput body and processing response.value themselves.
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-AdminApiContext',
    'Get-AdminApiToken',
    'Get-AnchorMailboxHeader',
    'Invoke-AdminApiRequest',
    'Invoke-AdminApiPagedRequest'
)
