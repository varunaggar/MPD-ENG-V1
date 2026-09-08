<#
.SYNOPSIS
    Configuration loader and Azure authentication helpers.

.DESCRIPTION
    Reads config.xml and exposes it as a typed object.
    Handles one-time Connect-AzAccount using the certificate
    stored in the Windows certificate store.

.NOTES
    Every script imports this module first, loads config,
    then calls Connect-SyncServicePrincipal before doing
    any Graph or SQL work.

    Dependencies: Az.Accounts (must be installed on the server)
    Install: Install-Module Az.Accounts -Scope AllUsers -Force
#>

# ──────────────────────────────────────────────────────────────
# Public: Import-SyncConfig
# Reads config.xml and returns the configuration object.
# ──────────────────────────────────────────────────────────────

function Import-SyncConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string]$Path
    )

    try {
        [xml]$raw = Get-Content -Path $Path -Encoding UTF8 -ErrorAction Stop
        $cfg = $raw.M365PermsSyncConfig

        if ($null -eq $cfg) {
            throw "Invalid configuration: Root element <M365PermsSyncConfig> not found in '$Path'."
        }

        $validAuthTypes = @('Certificate', 'Secret', 'ManagedIdentity', 'User')
        $authType = $cfg.Authentication.AuthType

        # Validate mandatory placeholders have been replaced
        $requiredFields = @(
            @{ Path = 'Authentication.AuthType';            Value = $authType }
            @{ Path = 'Authentication.TenantId';            Value = $cfg.Authentication.TenantId }
            @{ Path = 'Database.Server';                    Value = $cfg.Database.Server }
            @{ Path = 'Database.Name';                      Value = $cfg.Database.Name }
            #@{ Path = 'Logging.Directory';                  Value = $cfg.Logging.Directory }
        )

        foreach ($field in $requiredFields) {
            if ([string]::IsNullOrWhiteSpace($field.Value) -or $field.Value -match '^REPLACE-') {
                throw "Configuration error: The field '$($field.Path)' is missing or contains a placeholder. Please update config.xml."
            }
        }

        if ($validAuthTypes -notcontains $authType) {
            throw "Configuration error: Authentication.AuthType must be one of: $($validAuthTypes -join ', ')."
        }

        $modeSpecificFields = @()
        switch ($authType) {
            'Certificate' {
                $modeSpecificFields = @(
                    @{ Path = 'Authentication.AppId';                Value = $cfg.Authentication.AppId }
                    @{ Path = 'Authentication.CertificateThumbprint'; Value = $cfg.Authentication.CertificateThumbprint }
                    @{ Path = 'Authentication.CertificateStoreLocation'; Value = $cfg.Authentication.CertificateStoreLocation }
                    @{ Path = 'Authentication.CertificateStoreName';   Value = $cfg.Authentication.CertificateStoreName }
                )
            }
            'Secret' {
                $modeSpecificFields = @(
                    @{ Path = 'Authentication.AppId';      Value = $cfg.Authentication.AppId }
                    @{ Path = 'Authentication.ClientSecret'; Value = $cfg.Authentication.ClientSecret }
                )
            }
            'ManagedIdentity' {
                $modeSpecificFields = @() # no extra required fields for system-assigned MI
            }
            'User' {
                $modeSpecificFields = @() # optional UserPrincipalName for interactive auth
            }
        }

        foreach ($field in $modeSpecificFields) {
            if ([string]::IsNullOrWhiteSpace($field.Value) -or $field.Value -match '^REPLACE-') {
                throw "Configuration error: The field '$($field.Path)' is missing or contains a placeholder. Please update config.xml."
            }
        }

        return $cfg
    }
    catch {
        throw "Failed to load configuration from '$Path': $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Connect-SyncServicePrincipal
# Authenticates to Azure using the configured authentication method.
# Called once per script; Get-AzAccessToken then works for
# both Graph and SQL tokens within the same session.
# ──────────────────────────────────────────────────────────────

function Connect-SyncServicePrincipal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Config
    )

    $authType              = $Config.Authentication.AuthType
    $appId                 = $Config.Authentication.AppId
    $tenantId              = $Config.Authentication.TenantId
    $thumbprint            = $Config.Authentication.CertificateThumbprint
    $storeLoc              = $Config.Authentication.CertificateStoreLocation
    $storeName             = $Config.Authentication.CertificateStoreName
    $clientSecret          = $Config.Authentication.ClientSecret
    $managedIdentityClientId = $Config.Authentication.ManagedIdentityClientId
    $userPrincipalName     = $Config.Authentication.UserPrincipalName

    Write-LogInfo "Establishing Azure connection (AuthType: $authType, Tenant: $tenantId, AppId: $appId)..."

    try {
        switch ($authType) {
            'Certificate' {
                $thumbprint = $thumbprint.Replace(' ', '').Trim()
                $certPath = "Cert:\$storeLoc\$storeName\$thumbprint"

                if (-not (Test-Path $certPath)) {
                    $errMsg = "Authentication Certificate not found in store: $certPath"
                    Write-LogError $errMsg
                    throw $errMsg
                }

                Connect-AzAccount `
                    -ServicePrincipal `
                    -ApplicationId $appId `
                    -Tenant $tenantId `
                    -CertificateThumbprint $thumbprint `
                    -ErrorAction Stop | Out-Null
                Write-LogInfo "Successfully authenticated to Azure using certificate auth."
            }
            'Secret' {
                $secureSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential($appId, $secureSecret)

                Connect-AzAccount `
                    -ServicePrincipal `
                    -Tenant $tenantId `
                    -Credential $credential `
                    -ErrorAction Stop | Out-Null
                Write-LogInfo "Successfully authenticated to Azure using client secret auth."
            }
            'ManagedIdentity' {
                if ([string]::IsNullOrWhiteSpace($managedIdentityClientId)) {
                    Connect-AzAccount -Identity -ErrorAction Stop | Out-Null
                }
                else {
                    Connect-AzAccount -Identity -AccountId $managedIdentityClientId -ErrorAction Stop | Out-Null
                }
                Write-LogInfo "Successfully authenticated to Azure using managed identity auth."
            }
            'User' {
                if ([string]::IsNullOrWhiteSpace($userPrincipalName)) {
                    Connect-AzAccount -Tenant $tenantId -ErrorAction Stop | Out-Null
                }
                else {
                    Write-LogInfo "UserPrincipalName is configured, but interactive auth will use device/browser login because Connect-AzAccount does not accept a Username parameter in this module version."
                    Connect-AzAccount -Tenant $tenantId -UseDeviceAuthentication -ErrorAction Stop | Out-Null
                }
                Write-LogInfo "Successfully authenticated to Azure using interactive user auth."
            }
            default {
                throw "Unsupported Authentication.AuthType value: $authType"
            }
        }
    }
    catch {
        $errMsg = "Azure authentication failed: $($_.Exception.Message)"
        Write-LogError $errMsg -ErrorRecord $_
        throw $errMsg
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Get-ConfigValue
# Safe helper to read a string value from the config with a default.
# ──────────────────────────────────────────────────────────────

function Get-ConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config,
        [Parameter(Mandatory)] [string]$Section,
        [Parameter(Mandatory)] [string]$Key,
        [string]$Default = $null
    )

    try {
        $value = $Config.$Section.$Key
        if ([string]::IsNullOrWhiteSpace($value)) { return $Default }
        return $value
    }
    catch {
        return $Default
    }
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Import-SyncConfig',
    'Connect-SyncServicePrincipal',
    'Get-ConfigValue'
)
