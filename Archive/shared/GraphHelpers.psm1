<#
.SYNOPSIS
    Microsoft Graph REST API helpers — certificate-based authentication.

.DESCRIPTION
    Provides token acquisition, paged requests, and delta query support.
    Assumes Connect-SyncServicePrincipal has already been called
    (which establishes an Az context using the certificate).

    Token is cached for its lifetime. All retry/throttle/paging
    logic is encapsulated here — calling scripts just call the
    top-level functions and get results.

.NOTES
    Dependencies: Az.Accounts (must be installed on the server)
    Config is passed in via Initialize-GraphContext.
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

$script:GraphConfig = $null

$script:TokenCache = @{
    Token     = $null
    ExpiresAt = [datetime]::MinValue
}

# ──────────────────────────────────────────────────────────────
# Public: Initialize-GraphContext
# Called once per script after loading config.
# ──────────────────────────────────────────────────────────────

function Initialize-GraphContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    $script:GraphConfig = $Config
    # Reset token cache so any previous session token is not reused
    $script:TokenCache.Token     = $null
    $script:TokenCache.ExpiresAt = [datetime]::MinValue
    
    Write-LogInfo "Graph context initialised (BaseUrl=$($Config.Graph.BaseUrl))"
}

# ──────────────────────────────────────────────────────────────
# Public: Get-GraphToken
# Returns a valid Graph access token, refreshing if needed.
# The Az context established by Connect-SyncServicePrincipal
# allows Get-AzAccessToken to work without re-authenticating.
# ──────────────────────────────────────────────────────────────

function Get-GraphToken {
    [CmdletBinding()]
    param([switch]$ForceRefresh)

    $now = [datetime]::UtcNow

    if (-not $ForceRefresh -and
        $script:TokenCache.Token -and
        $script:TokenCache.ExpiresAt -gt $now.AddMinutes(2)) {
        # Silence frequent cache hits to keep logs clean
        return $script:TokenCache.Token
    }

    Write-LogInfo "Acquiring Graph access token..."

    try {
        $tokenInfo = Get-AzAccessToken `
            -ResourceUrl "https://graph.microsoft.com/" `
            -ErrorAction Stop

        # Ensure the token is a plain string. Modern Az modules return SecureString.
        $plainToken = if ($tokenInfo.Token -is [System.Security.SecureString]) {
            $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenInfo.Token)
            try {
                [Runtime.InteropServices.Marshal]::PtrToStringUni($ptr)
            }
            finally {
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
            }
        } else {
            $tokenInfo.Token
        }

        $script:TokenCache.Token     = $plainToken
        # Expire 10 minutes early as a safety margin
        $script:TokenCache.ExpiresAt = $tokenInfo.ExpiresOn.UtcDateTime.AddMinutes(-10)

        Write-LogInfo "Graph token acquired. Valid until $($script:TokenCache.ExpiresAt) UTC"
        return $plainToken
    }
    catch {
        throw "Failed to acquire Graph access token: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphRequest
# Single REST call with retry on 429, 401 and 5xx.
# ──────────────────────────────────────────────────────────────

function Invoke-GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Uri,
        [ValidateSet("GET","POST","PATCH","PUT","DELETE")] [string]$Method = "GET",
        [object]$Body = $null,
        [int]$MaxRetries = $null,
        [int]$TimeoutSec = 100
    )

    # Use config value if not explicitly passed
    if (-not $MaxRetries) {
        $MaxRetries = if ($script:GraphConfig -and $script:GraphConfig.Graph.MaxRetries) { [int]$script:GraphConfig.Graph.MaxRetries } else { 5 }
    }
    $throttleMax = if ($script:GraphConfig -and $script:GraphConfig.Graph.ThrottleBackoffMaxSec) { [int]$script:GraphConfig.Graph.ThrottleBackoffMaxSec } else { 300 }

    $attempt   = 0
    $lastError = $null

    while ($attempt -lt $MaxRetries) {
        $attempt++

        try {
            $token   = Get-GraphToken
            $headers = @{
                Authorization    = "Bearer $token"
                "Content-Type"   = "application/json"
                ConsistencyLevel = "eventual"
            }

            $params = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $headers
                TimeoutSec  = $TimeoutSec
                ErrorAction = "Stop"
            }

            if ($Body -and $Method -in @("POST","PATCH","PUT")) {
                $params.Body = if ($Body -isnot [string]) { $Body | ConvertTo-Json -Depth 10 } else { $Body }
            }

            return Invoke-RestMethod @params
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
            
            # Attempt to extract the actual Graph error message from the response body
            $apiErrorMessage = ""
            try {
                $responseStream = $_.Exception.Response.GetResponseStream()
                if ($null -ne $responseStream) {
                    $reader = New-Object System.IO.StreamReader($responseStream)
                    $responseBody = $reader.ReadToEnd()
                    # Basic regex to pull the 'message' field from Graph's error JSON
                    if ($responseBody -match '"message":"([^"]+)"') {
                        $apiErrorMessage = " | API Message: $($Matches[1])"
                    }
                }
            } catch {}

            # 429 — throttled
            if ($statusCode -eq 429) {
                $wait = 30
                try {
                    $h = $_.Exception.Response.Headers["Retry-After"]
                    if ($h) { $wait = [int]$h }
                } catch {}
                $wait = [Math]::Min($wait, $throttleMax)
                
                $logMsg = "Graph throttled (429). Waiting ${wait}s before retry $attempt/$MaxRetries$apiErrorMessage"
                Write-LogWarning $logMsg
                
                Start-Sleep -Seconds $wait
                $lastError = $_
                continue
            }

            # 401 — force token refresh once
            if ($statusCode -eq 401 -and $attempt -eq 1) {
                $logMsg = "Graph returned 401. Refreshing token and retrying...$apiErrorMessage"
                Write-LogWarning $logMsg
                
                Get-GraphToken -ForceRefresh | Out-Null
                $lastError = $_
                continue
            }

            # 5xx — transient server error, exponential backoff
            if ($statusCode -ge 500 -and $statusCode -le 599) {
                $wait = [Math]::Min([Math]::Pow(2, $attempt), 60)
                
                $logMsg = "Graph returned $statusCode. Waiting ${wait}s before retry $attempt/$MaxRetries$apiErrorMessage"
                Write-LogWarning $logMsg
                
                Start-Sleep -Seconds $wait
                $lastError = $_
                continue
            }

            # 410 Gone — delta token expired; re-throw immediately for caller to handle
            if ($statusCode -eq 410) { throw }

            # All other errors — do not retry
            throw
        }
    }

    $finalError = "Invoke-GraphRequest ($Method) failed after $MaxRetries attempts. URI: $Uri | Error: $($lastError.Exception.Message)"
    # Include API message in the final throw if we found one
    if ($apiErrorMessage) { $finalError += $apiErrorMessage }
    
    throw $finalError
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphPagedRequest
# Follows @odata.nextLink automatically.
# Returns all result objects as a single array.
# ──────────────────────────────────────────────────────────────

function Invoke-GraphPagedRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Uri
    )

    $allObjects = [System.Collections.Generic.List[object]]::new()
    $currentUri = $Uri
    $page       = 0

    do {
        $page++
        Write-LogInfo "Fetching page $page from Graph..."

        $response = Invoke-GraphRequest -Uri $currentUri

        if ($response.value) {
            $allObjects.AddRange([object[]]$response.value)
        }

        $currentUri = $response.'@odata.nextLink'
    } while ($currentUri)

    Write-LogInfo "Paged request complete: $page pages, $($allObjects.Count) total objects"
    return $allObjects.ToArray()
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphDeltaQuery
# Handles the full delta pattern:
#   - pages through all results following @odata.nextLink
#   - captures @odata.deltaLink at the end of the page chain
#   - handles HTTP 410 Gone (expired token) gracefully
# Returns a hashtable with:
#   Objects      — changed objects
#   DeltaLink    — full delta link URL
#   DeltaToken   — extracted token value
#   TokenExpired — $true if HTTP 410 was received
# ──────────────────────────────────────────────────────────────

function Invoke-GraphDeltaQuery {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Uri
    )

    $allObjects = [System.Collections.Generic.List[object]]::new()
    $currentUri = $Uri
    $deltaLink  = $null
    $page       = 0

    try {
        do {
            $page++
            Write-LogInfo "Fetching delta page $page..."

            $response = Invoke-GraphRequest -Uri $currentUri

            if ($response.value) {
                $allObjects.AddRange([object[]]$response.value)
            }

            if ($response.'@odata.deltaLink') {
                $deltaLink  = $response.'@odata.deltaLink'
                $currentUri = $null
            }
            elseif ($response.'@odata.nextLink') {
                $currentUri = $response.'@odata.nextLink'
            }
            else {
                $currentUri = $null
            }

        } while ($currentUri)

        if (-not $deltaLink) {
            throw "Delta query completed $page pages but no @odata.deltaLink was returned"
        }

        Write-LogInfo "Delta query complete: $page pages, $($allObjects.Count) changed objects" 

        return @{
            Objects      = $allObjects.ToArray()
            DeltaLink    = $deltaLink
            DeltaToken   = Get-DeltaTokenFromUrl -Url $deltaLink
            TokenExpired = $false
        }
    }
    catch {
        $statusCode = $null
        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

        if ($statusCode -eq 410) {
            $msg = "Delta token expired (HTTP 410 Gone). Re-initialisation required."
            Write-LogWarning $msg
            
            return @{
                Objects      = @()
                DeltaLink    = $null
                DeltaToken   = $null
                TokenExpired = $true
            }
        }
        throw
    }
}

# ──────────────────────────────────────────────────────────────
# Private: Get-DeltaTokenFromUrl
# Extracts the $deltatoken query parameter from a deltaLink URL.
# ──────────────────────────────────────────────────────────────

function Get-DeltaTokenFromUrl {
    param([string]$Url)

    if ([string]::IsNullOrEmpty($Url)) { return $null }

    if ($Url -match '\$deltatoken=([^&]+)') { return [uri]::UnescapeDataString($Matches[1]) }
    if ($Url -match '\$skiptoken=([^&]+)')  { return [uri]::UnescapeDataString($Matches[1]) }

    return $null
}

# ──────────────────────────────────────────────────────────────
# Public: Get-GraphMailFolderPaths
# Returns a flat array of all visible folder paths for a mailbox,
# by traversing the mailFolders hierarchy via the Graph API.
#
# Uses GET /users/{userId}/mailFolders for root-level folders and
# GET /users/{userId}/mailFolders/{id}/childFolders recursively
# for any folder where childFolderCount > 0.
#
# Hidden folders (RecoverableItems, SubstrateHolds, etc.) are
# excluded automatically — Graph omits them when
# includeHiddenFolders is not set (the default).
#
# Returns paths in the format the Admin API MailboxFolderPermission
# endpoint expects:
#     \Calendar
#     \Inbox
#     \Inbox\Reports
#
# PREREQUISITE: the app registration must have the Graph application
# permission Mail.ReadBasic.All (least privileged for folder metadata).
# ──────────────────────────────────────────────────────────────

function Get-GraphMailFolderPaths {
    <#
    .SYNOPSIS
        Returns all visible folder paths for a mailbox via the Graph mailFolders API.
    .DESCRIPTION
        Calls GET /users/{userId}/mailFolders for root-level folders, then
        recursively traverses childFolders for any folder whose childFolderCount
        is greater than zero. Pagination is handled automatically via
        Invoke-GraphPagedRequest.

        Hidden/system folders (RecoverableItems, SubstrateHolds, etc.) are
        excluded by the Graph API by default and do not appear in the output.

        Required app permission: Mail.ReadBasic.All (application).
    .PARAMETER UserId
        UPN or Entra Object ID of the mailbox owner. UPN is preferred for
        consistency with the Admin API X-AnchorMailbox header.
    .OUTPUTS
        String[] of folder paths, e.g. '\Calendar', '\Inbox', '\Inbox\Reports'.
        Returns an empty array if the mailbox has no folders or the API returns
        nothing (rather than throwing).
    .EXAMPLE
        $paths = Get-GraphMailFolderPaths -UserId 'alex@contoso.com'
        # Returns: @('\Calendar','\Inbox','\Inbox\Reports','\Sent Items', ...)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$UserId
    )

    if (-not $script:GraphConfig) {
        throw "Get-GraphMailFolderPaths: Graph context not initialised. Call Initialize-GraphContext first."
    }

    $allPaths = [System.Collections.Generic.List[string]]::new()
    $baseUrl  = $script:GraphConfig.Graph.BaseUrl

    # Root-level folders — hidden folders excluded by default (Graph behaviour)
    $uri = "$baseUrl/users/$UserId/mailFolders" +
           "?`$select=id,displayName,childFolderCount&`$top=999"

    try {
        $rootFolders = Invoke-GraphPagedRequest -Uri $uri
    }
    catch {
        throw "Get-GraphMailFolderPaths: failed to list root folders for '$UserId': $($_.Exception.Message)"
    }

    foreach ($folder in $rootFolders) {
        $folderPath = "\$($folder.displayName)"
        $allPaths.Add($folderPath)

        # Recurse into any folder that has children
        if ([int]$folder.childFolderCount -gt 0) {
            try {
                $childPaths = Get-GraphChildFolderPaths `
                    -UserId     $UserId `
                    -FolderId   $folder.id `
                    -ParentPath $folderPath

                if ($childPaths -and $childPaths.Count -gt 0) {
                    $allPaths.AddRange([string[]]$childPaths)
                }
            }
            catch {
                # A failure on a single subtree is non-fatal.
                # Log and continue — the parent folder path is still captured.
                Write-LogWarning "Get-GraphMailFolderPaths: '$UserId' failed to enumerate children of '$folderPath': $($_.Exception.Message)"
            }
        }
    }

    return $allPaths.ToArray()
}

# ──────────────────────────────────────────────────────────────
# Private: Get-GraphChildFolderPaths
# Recursive helper for Get-GraphMailFolderPaths.
# Fetches child folders under a given folder ID and recurses
# for any child whose childFolderCount > 0.
# Builds folder paths incrementally: ParentPath + '\' + displayName.
# Not exported — called only by Get-GraphMailFolderPaths.
# ──────────────────────────────────────────────────────────────

function Get-GraphChildFolderPaths {
    param(
        [string]$UserId,
        [string]$FolderId,
        [string]$ParentPath
    )

    $paths   = [System.Collections.Generic.List[string]]::new()
    $baseUrl = $script:GraphConfig.Graph.BaseUrl

    $uri = "$baseUrl/users/$UserId/mailFolders/$FolderId/childFolders" +
           "?`$select=id,displayName,childFolderCount&`$top=999"

    $children = Invoke-GraphPagedRequest -Uri $uri

    foreach ($child in $children) {
        $childPath = "$ParentPath\$($child.displayName)"
        $paths.Add($childPath)

        if ([int]$child.childFolderCount -gt 0) {
            $deepPaths = Get-GraphChildFolderPaths `
                -UserId     $UserId `
                -FolderId   $child.id `
                -ParentPath $childPath

            if ($deepPaths -and $deepPaths.Count -gt 0) {
                $paths.AddRange([string[]]$deepPaths)
            }
        }
    }

    return $paths
}

# ──────────────────────────────────────────────────────────────
# Public: Invoke-GraphHuntingQuery
# Runs a KQL query against the Microsoft Defender XDR Advanced
# Hunting schema via POST /security/runHuntingQuery.
#
# Requires: ThreatHunting.Read.All application permission.
# Prerequisite: Microsoft Defender for Cloud Apps must be deployed
#               with Microsoft 365 activities connected.
#
# Returns the results array (PSCustomObject[]). An empty result
# set returns an empty array, not $null.
#
# Row limit: Defender caps results at 10,000 rows by default.
# Add '| take 100000' to the KQL to raise to the maximum.
# If the cap is reached a warning is logged — the caller should
# narrow the time window or add more specific filters.
# ──────────────────────────────────────────────────────────────

function Invoke-GraphHuntingQuery {
    <#
    .SYNOPSIS
        Runs a KQL query against the Defender XDR Advanced Hunting schema.
    .PARAMETER Query
        A valid KQL query string targeting CloudAppEvents or other tables.
    .OUTPUTS
        PSCustomObject[] — each element is one row from the result set.
        Returns an empty array when the query returns no results.
    .EXAMPLE
        $rows = Invoke-GraphHuntingQuery -Query @"
        CloudAppEvents
        | where Timestamp > datetime(2025-06-01T00:00:00Z)
        | where ActionType == 'Add-MailboxPermission'
        | project Timestamp, ActionType, ObjectName, RawEventData
        "@
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Query
    )

    if (-not $script:GraphConfig) {
        throw "Invoke-GraphHuntingQuery: Graph context not initialised. Call Initialize-GraphContext first."
    }

    $maxRetries  = if ($script:GraphConfig.Graph.MaxRetries)          { [int]$script:GraphConfig.Graph.MaxRetries }          else { 5 }
    $throttleMax = if ($script:GraphConfig.Graph.ThrottleBackoffMaxSec){ [int]$script:GraphConfig.Graph.ThrottleBackoffMaxSec } else { 300 }

    $uri  = "$($script:GraphConfig.Graph.BaseUrl)/security/runHuntingQuery"
    $body = @{ Query = $Query } | ConvertTo-Json -Depth 2

    $attempt = 0
    while ($attempt -lt $maxRetries) {
        $attempt++
        try {
            $token    = Get-GraphToken
            $headers  = @{
                Authorization  = "Bearer $token"
                'Content-Type' = 'application/json'
            }

            $response = Invoke-RestMethod `
                -Uri         $uri `
                -Method      POST `
                -Headers     $headers `
                -Body        $body `
                -ErrorAction Stop

            $results = if ($response.results) { $response.results } else { @() }

            # Warn if the result cap may have been hit
            if ($results.Count -eq 10000) {
                Write-LogWarning "Invoke-GraphHuntingQuery: result count hit 10,000 row cap. Add '| take 100000' to your query or narrow the time window."
            }

            Write-LogInfo "Invoke-GraphHuntingQuery: $($results.Count) row(s) returned"
            return $results
        }
        catch {
            $statusCode = $null
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

            if ($statusCode -eq 429) {
                $wait = 60
                try { $h = $_.Exception.Response.Headers['Retry-After']; if ($h) { $wait = [int]$h } } catch {}
                $wait = [Math]::Min($wait, $throttleMax)
                Write-LogWarning "Hunting query throttled (429). Waiting ${wait}s. Attempt $attempt/$maxRetries"
                Start-Sleep -Seconds $wait
                continue
            }

            if ($statusCode -eq 401 -and $attempt -eq 1) {
                Write-LogWarning "Hunting query returned 401. Refreshing token and retrying."
                Get-GraphToken -ForceRefresh | Out-Null
                continue
            }

            if ($statusCode -ge 500 -and $statusCode -le 599) {
                $wait = [Math]::Min([Math]::Pow(2, $attempt), 60)
                Write-LogWarning "Hunting query returned $statusCode. Waiting ${wait}s. Attempt $attempt/$maxRetries"
                Start-Sleep -Seconds $wait
                continue
            }

            throw "Invoke-GraphHuntingQuery failed (HTTP $statusCode): $($_.Exception.Message)"
        }
    }

    throw "Invoke-GraphHuntingQuery failed after $maxRetries attempts."
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-GraphContext',
    'Get-GraphToken',
    'Invoke-GraphRequest',
    'Invoke-GraphPagedRequest',
    'Invoke-GraphDeltaQuery',
    'Get-GraphMailFolderPaths',
    'Invoke-GraphHuntingQuery'
)
