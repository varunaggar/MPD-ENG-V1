<#
.SYNOPSIS
    Exchange Online PowerShell helpers — connection management only.

.DESCRIPTION
    Provides certificate-based Connect/Disconnect for Exchange Online.

    REMOVED from previous version:
      - Get-AllExoMailboxes      → scripts call Get-EXOMailbox directly
      - Get-ChangedExoMailboxes  → scripts call Get-EXOMailbox directly
      - ConvertTo-MailboxObject  → scripts normalise mailbox data inline

    The app registration needs Exchange.ManageAsApp application permission
    and the service principal must be assigned the Exchange Online
    management role:
      New-ManagementRoleAssignment -App "<app-name>" -Role "Mail Recipients"
#>

$script:ExoConfig   = $null
$script:IsConnected = $false

# ──────────────────────────────────────────────────────────────
# Public: Initialize-ExoContext
# ──────────────────────────────────────────────────────────────

function Initialize-ExoContext {
    <#
    .SYNOPSIS
        Validates EXO config and checks module availability.
        Call once per script before Connect-ExoSession.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Config)

    $script:ExoConfig   = $Config
    $script:IsConnected = $false

    $org = $Config.ExchangeOnline.Organisation
    if ([string]::IsNullOrWhiteSpace($org) -or $org -like 'REPLACE-*') {
        throw "config.xml ExchangeOnline.Organisation is not set."
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module not installed. " +
              "Run: Install-Module ExchangeOnlineManagement -Scope AllUsers -Force"
    }

    Write-LogInfo "EXO context initialised (Organisation=$org)"
}

# ──────────────────────────────────────────────────────────────
# Public: Connect-ExoSession
# ──────────────────────────────────────────────────────────────

function Connect-ExoSession {
    <#
    .SYNOPSIS
        Connects to Exchange Online using certificate-based auth.
        Safe to call multiple times — no-op if already connected.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)] [object]$Config)

    if ($script:IsConnected) { return }

    $appId      = $Config.Authentication.AppId
    $thumbprint = $Config.Authentication.CertificateThumbprint.Replace(' ', '').Trim()
    $storeLoc   = $Config.Authentication.CertificateStoreLocation
    $org        = $Config.ExchangeOnline.Organisation

    $certPath = "Cert:\$storeLoc\My\$thumbprint"
    if (-not (Test-Path $certPath)) {
        throw "Certificate '$thumbprint' not found at $certPath"
    }

    Write-LogInfo "Connecting to Exchange Online (Org: $org)..."

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction Stop

        Connect-ExchangeOnline `
            -AppId                 $appId `
            -CertificateThumbprint $thumbprint `
            -Organization          $org `
            -ShowBanner:           $false `
            -ErrorAction           Stop

        $script:IsConnected = $true
        Write-LogInfo "Exchange Online connected"
    }
    catch {
        $script:IsConnected = $false
        throw "Failed to connect to Exchange Online: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Disconnect-ExoSession
# ──────────────────────────────────────────────────────────────

function Disconnect-ExoSession {
    <#
    .SYNOPSIS
        Cleanly disconnects the EXO session. Call at end of each script.
    #>
    [CmdletBinding()]
    param()

    if (-not $script:IsConnected) { return }

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-LogInfo "Exchange Online disconnected"
    }
    catch {
        Write-LogWarning "EXO disconnect error: $($_.Exception.Message)"
    }
    finally {
        $script:IsConnected = $false
    }
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-ExoContext',
    'Connect-ExoSession',
    'Disconnect-ExoSession'
)
