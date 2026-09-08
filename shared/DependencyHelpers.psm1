<#
.SYNOPSIS
    Dependency validation and local module loading.

.DESCRIPTION
    Ensures all required PowerShell modules are present in a local 'Modules' 
    folder and meet version requirements before execution continues.
#>

function Initialize-ModuleDependencies {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config
    )

    # 1. Extract modules from Config (handles single or multiple XML nodes)
    $requiredModules = @()
    if ($null -ne $Config.Dependencies -and $null -ne $Config.Dependencies.Module) {
        $requiredModules = $Config.Dependencies.Module
    }

    # 2. Check if there are any required modules. If not, exit early.
    if (@($requiredModules).Count -eq 0) {
        Write-LogInfo "No dependencies defined in config.xml. Skipping local module check."
        return
    }

    # 3. Since dependencies are required, check for the local Modules folder
    $localModulesPath = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\Modules"))

    if (-not (Test-Path $localModulesPath)) {
        $errMsg = "Fatal: Dependencies are required in config.xml, but the local 'Modules' folder was not found at '$localModulesPath'."
        Write-LogError $errMsg
        exit 1
    }

    Write-LogInfo "Validating required dependencies in local folder: $localModulesPath"

    # 3. Validate and Import each module
    foreach ($modReq in $requiredModules) {
        $name    = $modReq.Name
        $minVer  = [version]$modReq.MinimumVersion

        Write-LogInfo "Checking dependency: $name (>= $minVer)"

        # Find module folder(s) matching the required module name
        $moduleFolders = Get-ChildItem -Path $localModulesPath -Directory -ErrorAction SilentlyContinue | 
                         Where-Object { $_.Name -like "$name*" }

        if ($moduleFolders.Count -eq 0) {
            $errMsg = "Fatal: Module '$name' not found in '$localModulesPath'."
            Write-LogError $errMsg
            exit 1
        }

        # Find all .psd1 manifests and validate versions
        $validModules = @()
        foreach ($folder in $moduleFolders) {
            Get-ChildItem -Path $folder -Filter "*.psd1" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $manifest = Test-ModuleManifest -Path $_.FullName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                    if ($null -ne $manifest -and $manifest.Version -ge $minVer) {
                        $validModules += @{
                            Path    = $_.FullName
                            Version = $manifest.Version
                        }
                    }
                }
                catch {
                    # Skip manifests with errors
                }
            }
        }

        if ($validModules.Count -eq 0) {
            $errMsg = "Fatal: Module '$name' with version >= '$minVer' was not found in '$localModulesPath'."
            Write-LogError $errMsg
            exit 1
        }

        # Select the highest version
        $moduleToLoad = $validModules | Sort-Object -Property Version -Descending | Select-Object -First 1

        try {
            # Import the module directly by its full path into global scope so its commands
            # are available to the calling script.
            Import-Module -FullyQualifiedName $moduleToLoad.Path -Scope Global -ErrorAction Stop
            Write-LogInfo "Successfully imported $name version $($moduleToLoad.Version) from local path"
        }
        catch {
            $errMsg = "Fatal: Failed to import module from '$($moduleToLoad.Path)': $($_.Exception.Message)"
            Write-LogError $errMsg
            exit 1
        }
    }

    Write-LogInfo "All dependencies validated and loaded successfully."
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-ModuleDependencies'
)