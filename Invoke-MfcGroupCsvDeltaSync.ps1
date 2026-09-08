<#
.SYNOPSIS
    Synchronizes mfc-sh* groups from Microsoft Graph to a CSV file.

.DESCRIPTION
    The first run reads /groups/delta and creates a delta token. Later runs
    use the saved delta link to fetch new, updated, and deleted groups.

    This script intentionally uses no modules and defines no functions.

.PARAMETER ConfigPath
    Path to the existing M365 permissions config.xml file.

.PARAMETER CsvPath
    Destination CSV file. Defaults to mfc-groups.csv beside this script.

.PARAMETER TokenPath
    JSON file containing the Graph delta link. Defaults to
    mfc-groups-delta-token.json beside this script.
#>

[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.xml"),
    [string]$CsvPath = (Join-Path $PSScriptRoot "mfc-groups.csv"),
    [string]$TokenPath = (Join-Path $PSScriptRoot "mfc-groups-delta-token.json")
)

$ErrorActionPreference = "Stop"

try {
    # Load runtime settings and acquire an application-only Graph access token.
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        throw "Configuration file was not found: $ConfigPath"
    }

    [xml]$config = Get-Content -LiteralPath $ConfigPath -Raw
    $tenantId = "086b944f-0278-4374-8dfe-305c4b5d1d70"
    $clientId = "f2a1282c-a035-4e4b-b16d-789f1d5bc2c7"
    $clientSecret = "U4r8Q~emCNI64M9pjjucn_yvpjfDx1290PfD3awc"
    $graphBaseUrl = "https://graph.microsoft.com/v1.0"
    $pageSize = '5'
    $maxRetries = '5'

    if ([string]::IsNullOrWhiteSpace($tenantId) -or
        [string]::IsNullOrWhiteSpace($clientId) -or
        [string]::IsNullOrWhiteSpace($clientSecret)) {
        throw "TenantId, AppId, and ClientSecret are required in config.xml."
    }

    if ([string]::IsNullOrWhiteSpace($graphBaseUrl)) {
        $graphBaseUrl = "https://graph.microsoft.com/v1.0"
    }
    if ($pageSize -le 0) { $pageSize = 999 }
    if ($maxRetries -le 0) { $maxRetries = 5 }

    $tokenRequest = @{
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }
    $authResponse = Invoke-RestMethod `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
        -Method Post `
        -Body $tokenRequest `
        -ContentType "application/x-www-form-urlencoded"

    Write-Host "[GRAPH AUTH] Access token acquired for Microsoft Graph." -ForegroundColor DarkCyan

    if ([string]::IsNullOrWhiteSpace([string]$authResponse.access_token)) {
        throw "The token endpoint returned no access token."
    }

    $headers = @{
        Authorization    = "Bearer $($authResponse.access_token)"
        Accept           = "application/json"
        ConsistencyLevel = "eventual"
    }

    # Load the existing flat CSV into memory so delta results can be merged.
    $rowsByGroupId = @{}
    if (Test-Path -LiteralPath $CsvPath) {
        Import-Csv -LiteralPath $CsvPath | ForEach-Object {
            $existingGroupId = [string]$_.GroupId
            if ([string]::IsNullOrWhiteSpace($existingGroupId)) {
                $existingGroupId = [string]$_.Id
            }
            if (-not [string]::IsNullOrWhiteSpace($existingGroupId)) {
                if (-not $rowsByGroupId.ContainsKey($existingGroupId)) {
                    $rowsByGroupId[$existingGroupId] = [System.Collections.Generic.List[object]]::new()
                }
                [void]$rowsByGroupId[$existingGroupId].Add($_)
            }
        }
    }

    $storedDeltaLink = $null
    $tokenCreatedUtc = $null

    # Read the opaque delta link and inspect its age before making Graph calls.
    if (Test-Path -LiteralPath $TokenPath) {
        $tokenDocument = Get-Content -LiteralPath $TokenPath -Raw | ConvertFrom-Json
        $storedDeltaLink = [string]$tokenDocument.deltaLink
        $tokenCreatedUtc = $tokenDocument.createdUtc
        if ($null -eq $tokenCreatedUtc) {
            $tokenCreatedUtc = $tokenDocument.updatedUtc
        }
    }

    if ($null -ne $tokenCreatedUtc) {
        $parsedTokenCreatedUtc = [datetime]::MinValue
        $tokenDateParsed = if ($tokenCreatedUtc -is [datetime]) {
            $parsedTokenCreatedUtc = ([datetime]$tokenCreatedUtc).ToUniversalTime()
            $true
        } else {
            [datetime]::TryParse(
                [string]$tokenCreatedUtc,
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::RoundtripKind,
                [ref]$parsedTokenCreatedUtc
            )
        }
        if ($tokenDateParsed) {
            $tokenAgeDays = ([datetime]::UtcNow - $parsedTokenCreatedUtc.ToUniversalTime()).TotalDays
            Write-Host "[GRAPH] Stored delta token age: $([math]::Round($tokenAgeDays, 2)) day(s)." -ForegroundColor DarkCyan
            if ($tokenAgeDays -ge 5) {
                Write-Warning "The stored delta token is $([math]::Round($tokenAgeDays, 2)) day(s) old and may expire soon. A full baseline may be required after seven days."
            }
        } else {
            Write-Warning "The token metadata contains an invalid createdUtc/updatedUtc value: $tokenCreatedUtc"
        }
    }

    if ([string]::IsNullOrWhiteSpace($storedDeltaLink)) {
        # No token means this is a baseline query; later runs reuse the full delta link.
        $select = "id,displayName,mail,members"
        $requestUri = "$($graphBaseUrl.TrimEnd('/'))/groups/delta?`$select=$select&`$top=$pageSize"
        Write-Host "[GRAPH] Starting initial mfc-sh group delta query: $($requestUri.Split('?')[0])" -ForegroundColor Cyan
    } else {
        $requestUri = $storedDeltaLink
        Write-Host "[GRAPH] Using stored group delta link: $($requestUri.Split('?')[0])" -ForegroundColor Cyan
    }

    $pendingChanges = @{}
    $groupAttributeChanges = 0
    $membershipChanges = 0
    $nextDeltaLink = $null
    $pageNumber = 0

    # Walk every delta page. The final page supplies the next checkpoint link.
    do {
        $pageNumber++
        $response = $null
        $attempt = 0

        do {
            $attempt++
            Write-Host "[GRAPH CALL] GET group delta page $pageNumber, attempt ${attempt}: $($requestUri.Split('?')[0])" -ForegroundColor DarkCyan
            try {
                $response = Invoke-RestMethod `
                    -Uri $requestUri `
                    -Method Get `
                    -Headers $headers `
                    -ErrorAction Stop
            } catch {
                $statusCode = 0
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

                if ($statusCode -eq 401 -and $attempt -lt $maxRetries) {
                    throw "Graph returned HTTP 401. The access token may have expired; rerun the script."
                }
                if (($statusCode -eq 429 -or $statusCode -ge 500) -and $attempt -lt $maxRetries) {
                    $retryAfter = 5
                    try {
                        $retryHeader = $_.Exception.Response.Headers["Retry-After"]
                        if ($retryHeader) { $retryAfter = [int]$retryHeader }
                    } catch {}
                    $retryAfter = [Math]::Min([Math]::Max($retryAfter, 1), 300)
                    Write-Warning "Graph returned HTTP $statusCode. Retrying in $retryAfter second(s)."
                    Start-Sleep -Seconds $retryAfter
                } else {
                    throw "Graph request failed with HTTP $statusCode. URI: $requestUri. $($_.Exception.Message)"
                }
            }
        } while ($null -eq $response -and $attempt -lt $maxRetries)

        if ($null -eq $response) {
            throw "Graph request failed after $maxRetries attempts."
        }

        $groupResponseCount = @($response.value).Count
        Write-Host "[GRAPH RESPONSE] Group delta page $pageNumber returned $groupResponseCount object(s); next page: $([bool]$response.'@odata.nextLink'); delta link: $([bool]$response.'@odata.deltaLink')" -ForegroundColor Cyan

        # Classify each result as deleted, group-attribute, membership, or combined change.
        foreach ($group in @($response.value)) {
            $groupId = [string]$group.id
            if ([string]::IsNullOrWhiteSpace($groupId)) { continue }

            if ($null -ne $group.'@removed') {
                $pendingChanges[$groupId] = $null
                Write-Host "[CHANGE: DELETE] Group $groupId ($($group.displayName)) was deleted in Graph." -ForegroundColor Red
                continue
            }

            $existingRowForFilter = if ($rowsByGroupId.ContainsKey($groupId)) { @($rowsByGroupId[$groupId])[0] } else { $null }
            $groupDisplayNameForFilter = [string]$group.displayName
            if ([string]::IsNullOrWhiteSpace($groupDisplayNameForFilter) -and $null -ne $existingRowForFilter) {
                $groupDisplayNameForFilter = [string]$existingRowForFilter.GroupDisplayName
            }

            if ($groupDisplayNameForFilter -notlike "mfc-sh*") {
                if ($rowsByGroupId.ContainsKey($groupId)) {
                    $pendingChanges[$groupId] = $null
                    Write-Host "[CHANGE: DELETE] Group $groupId no longer matches mfc-sh* and will be removed from CSV." -ForegroundColor Yellow
                }
                continue
            }

            $existingRow = if ($rowsByGroupId.ContainsKey($groupId)) { @($rowsByGroupId[$groupId])[0] } else { $null }
            $displayName = [string]$group.displayName
            $mail = [string]$group.mail
            if ([string]::IsNullOrWhiteSpace($displayName) -and $null -ne $existingRow) {
                $displayName = [string]$existingRow.GroupDisplayName
            }
            if ([string]::IsNullOrWhiteSpace($mail) -and $null -ne $existingRow) {
                $mail = [string]$existingRow.GroupMail
            }

            $hasGroupAttributeChange = $null -eq $existingRow -or
                $displayName -ne [string]$existingRow.GroupDisplayName -or
                $mail -ne [string]$existingRow.GroupMail
            $hasMembershipChange = $null -ne $group.'members@delta' -or $null -eq $existingRow

            $pendingChanges[$groupId] = [pscustomobject]@{
                Id                    = $groupId
                DisplayName           = $displayName
                Mail                  = $mail
                GroupAttributeChanged = $hasGroupAttributeChange
                MembershipChanged     = $hasMembershipChange
            }

            $changeColor = if ($null -eq $existingRow) { "Green" } else { "Yellow" }
            if ($hasGroupAttributeChange -and $hasMembershipChange) {
                Write-Host "[CHANGE: GROUP + MEMBERSHIP] Group $groupId | displayName='$displayName' | mail='$mail'" -ForegroundColor $changeColor
            } elseif ($hasGroupAttributeChange) {
                Write-Host "[CHANGE: GROUP ATTRIBUTES] Group $groupId | displayName='$displayName' | mail='$mail'" -ForegroundColor $changeColor
            } elseif ($hasMembershipChange) {
                Write-Host "[CHANGE: MEMBERSHIP] Group $groupId | displayName='$displayName' | mail='$mail'" -ForegroundColor DarkCyan
            }
        }

        $nextPageLink = [string]$response.'@odata.nextLink'
        $nextDeltaLink = [string]$response.'@odata.deltaLink'
        $requestUri = $nextPageLink
    } while (-not [string]::IsNullOrWhiteSpace($requestUri))

    if ([string]::IsNullOrWhiteSpace($nextDeltaLink)) {
        throw "Graph returned no @odata.deltaLink. The CSV and token were not changed."
    }

    $groupAttributeChanges = @($pendingChanges.Values | Where-Object { $null -ne $_ -and $_.GroupAttributeChanged }).Count
    $membershipChanges = @($pendingChanges.Values | Where-Object { $null -ne $_ -and $_.MembershipChanged }).Count
    Write-Host "[CHANGE SUMMARY] Group attribute changes: $groupAttributeChanges; membership changes: $membershipChanges; deleted groups: $(@($pendingChanges.Values | Where-Object { $null -eq $_ }).Count)" -ForegroundColor Cyan

    # Rebuild member rows for every changed group so the CSV remains one row per member.
    foreach ($groupId in @($pendingChanges.Keys)) {
        if ($null -eq $pendingChanges[$groupId]) {
            [void]$rowsByGroupId.Remove($groupId)
            continue
        }

        $group = $pendingChanges[$groupId]
        $memberRows = [System.Collections.Generic.List[object]]::new()
        $membersUri = "$($graphBaseUrl.TrimEnd('/'))/groups/$groupId/members?`$select=id,displayName,mail&`$top=$pageSize"
        $memberPageNumber = 0

        do {
            $memberPageNumber++
            $memberResponse = $null
            $attempt = 0
            do {
                $attempt++
                Write-Host "[GRAPH CALL] GET members for group $groupId, page $memberPageNumber, attempt ${attempt}: $($membersUri.Split('?')[0])" -ForegroundColor DarkCyan
                try {
                    $memberResponse = Invoke-RestMethod `
                        -Uri $membersUri `
                        -Method Get `
                        -Headers $headers `
                        -ErrorAction Stop
                } catch {
                    $statusCode = 0
                    try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                    if (($statusCode -eq 429 -or $statusCode -ge 500) -and $attempt -lt $maxRetries) {
                        $retryAfter = 5
                        try {
                            $retryHeader = $_.Exception.Response.Headers["Retry-After"]
                            if ($retryHeader) { $retryAfter = [int]$retryHeader }
                        } catch {}
                        $retryAfter = [Math]::Min([Math]::Max($retryAfter, 1), 300)
                        Write-Warning "Graph member request returned HTTP $statusCode. Retrying in $retryAfter second(s)."
                        Start-Sleep -Seconds $retryAfter
                    } else {
                        throw "Graph member request failed with HTTP $statusCode. URI: $membersUri. $($_.Exception.Message)"
                    }
                }
            } while ($null -eq $memberResponse -and $attempt -lt $maxRetries)

            Write-Host "[GRAPH RESPONSE] Members page $memberPageNumber for group $groupId returned $(@($memberResponse.value).Count) object(s); next page: $([bool]$memberResponse.'@odata.nextLink')" -ForegroundColor Cyan

            foreach ($member in @($memberResponse.value)) {
                [void]$memberRows.Add([pscustomobject]@{
                    GroupId          = $group.Id
                    GroupDisplayName = $group.DisplayName
                    GroupMail        = $group.Mail
                    MemberId         = [string]$member.id
                    MemberDisplayName = [string]$member.displayName
                    MemberMail       = [string]$member.mail
                })
                $memberLogLabel = if ($group.MembershipChanged) { "MEMBERSHIP CHANGE" } else { "MEMBER SNAPSHOT" }
                Write-Host "[$memberLogLabel] Group $groupId <- member $($member.id) | displayName='$($member.displayName)' | mail='$($member.mail)'" -ForegroundColor DarkCyan
            }
            $membersUri = [string]$memberResponse.'@odata.nextLink'
        } while (-not [string]::IsNullOrWhiteSpace($membersUri))

        if ($memberRows.Count -eq 0) {
            [void]$memberRows.Add([pscustomobject]@{
                GroupId           = $group.Id
                GroupDisplayName  = $group.DisplayName
                GroupMail         = $group.Mail
                MemberId          = ""
                MemberDisplayName = ""
                MemberMail        = ""
            })
            $memberLogLabel = if ($group.MembershipChanged) { "MEMBERSHIP CHANGE" } else { "MEMBER SNAPSHOT" }
            Write-Host "[$memberLogLabel] Group $groupId has no members; retaining one empty-member row." -ForegroundColor DarkGray
        }
        $rowsByGroupId[$groupId] = $memberRows
        Write-Host "[CSV] Prepared $($memberRows.Count) row(s) for group $groupId." -ForegroundColor Gray
    }

    $csvDirectory = Split-Path -Parent $CsvPath
    $tokenDirectory = Split-Path -Parent $TokenPath
    if (-not [string]::IsNullOrWhiteSpace($csvDirectory)) {
        [void](New-Item -ItemType Directory -Path $csvDirectory -Force)
    }
    if (-not [string]::IsNullOrWhiteSpace($tokenDirectory)) {
        [void](New-Item -ItemType Directory -Path $tokenDirectory -Force)
    }

    # Write the complete CSV snapshot only when Graph reported changes.
    $csvRows = @($rowsByGroupId.Values | ForEach-Object { $_ })
    if ($pendingChanges.Count -gt 0) {
        $csvRows |
            Sort-Object GroupDisplayName, GroupId, MemberDisplayName, MemberId |
            Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Host "[CSV] Wrote $($csvRows.Count) row(s) to $CsvPath after $($pendingChanges.Count) group change(s)." -ForegroundColor Gray
    } else {
        Write-Host "[CSV] No group changes detected; existing CSV was not rewritten ($($csvRows.Count) row(s))." -ForegroundColor DarkGray
    }

    # Save the new delta checkpoint only after Graph and CSV processing succeed.
    $tokenDocument = [pscustomobject]@{
        deltaLink             = $nextDeltaLink
        createdUtc            = [datetime]::UtcNow.ToString("o")
        lastUsedUtc           = [datetime]::UtcNow.ToString("o")
        lastSuccessfulSyncUtc = [datetime]::UtcNow.ToString("o")
        updatedUtc            = [datetime]::UtcNow.ToString("o")
    }
    $tokenDocument | ConvertTo-Json | Set-Content -LiteralPath $TokenPath -Encoding UTF8
    Write-Host "[GRAPH] Delta token saved to $TokenPath." -ForegroundColor DarkCyan

    Write-Host "Sync complete. Pages: $pageNumber; changed groups: $($pendingChanges.Count); group attribute changes: $groupAttributeChanges; membership changes: $membershipChanges; groups in CSV: $($rowsByGroupId.Count)."
    Write-Host "CSV: $CsvPath"
    Write-Host "Delta token JSON: $TokenPath"
} catch {
    # Leave the previous CSV and delta link intact so the next run can retry safely.
    Write-Error "MFC group CSV delta sync failed: $($_.Exception.Message)"
    exit 1
}