<#
.SYNOPSIS
    File-based structured logging for the M365 permissions sync solution.

.DESCRIPTION
    Each process that calls Initialize-Logging gets its own log file.
    Log file name format: {ProcessName}_{yyyy-MM-dd_HH-mm-ss}.log

    Every log line is written to:
      1. The process-specific log file on disk
      2. The console (host) for interactive/task-scheduler visibility

    Log rotation: Remove-OldLogFiles deletes files older than RetentionDays.
    Call this at the end of each script execution.

.NOTES
    No external dependencies.
    Imported by every script as the first module after ConfigHelpers.
#>

# ──────────────────────────────────────────────────────────────
# Module-scoped state
# ──────────────────────────────────────────────────────────────

$script:LogFilePath    = $null
$script:LogDirectory   = $null
$script:ProcessName    = $null
$script:RetentionDays  = 30
$script:LogLevel       = 'Info'
$script:LogToEventLog  = $false
$script:EventSource    = 'M365PermSync'

# ──────────────────────────────────────────────────────────────
# Public: Initialize-Logging
# Must be called once at the start of each script.
# Creates the log directory if it doesn't exist.
# Opens the process-specific log file.
# ──────────────────────────────────────────────────────────────

function Initialize-Logging {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [object]$Config,
        [Parameter(Mandatory)] [string]$ProcessName
    )

    $script:ProcessName   = $ProcessName

    if ($null -ne $Config.Logging) {
        # Correctly identify empty or whitespace tags to trigger fallback
        $script:LogDirectory  = if (-not [string]::IsNullOrWhiteSpace([string]$Config.Logging.Directory)) { $Config.Logging.Directory } else { Join-Path $PSScriptRoot "..\logs" }
        $script:RetentionDays = if ($null -ne $Config.Logging.RetentionDays) { [int]$Config.Logging.RetentionDays } else { 30 }
        $script:LogLevel      = if ($null -ne $Config.Logging.Level) { $Config.Logging.Level } else { 'Info' }
        $script:LogToEventLog = if ($null -ne $Config.Logging.LogToEventLog) { $Config.Logging.LogToEventLog -eq 'true' } else { $false }
        $script:EventSource   = if ($null -ne $Config.Logging.EventSource) { $Config.Logging.EventSource } else { 'M365PermSync' }
    } else {
        # Fallback if the entire Logging section is missing
        $script:LogDirectory  = Join-Path $PSScriptRoot "..\logs"
        $script:RetentionDays = 30
        $script:LogLevel      = 'Info'
        $script:LogToEventLog = $false
        $script:EventSource   = 'M365PermSync'
    }

    # Create log directory if it doesn't exist
    try {
        if (-not (Test-Path $script:LogDirectory)) {
            New-Item -ItemType Directory -Path $script:LogDirectory -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        throw "LoggingHelpers: Failed to create or access log directory at '$($script:LogDirectory)'. Error: $($_.Exception.Message)"
    }

    # Build the log file name: ProcessName_yyyy-MM-dd_HH-mm-ss.log
    $timestamp          = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $logFileName        = "${ProcessName}_${timestamp}.log"
    $script:LogFilePath = Join-Path $script:LogDirectory $logFileName

    # Write header to log file
    $header = @"
================================================================================
  M365 Permissions Sync — Log File
  Process  : $ProcessName
  Started  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') UTC
  Host     : $($env:COMPUTERNAME)
  User     : $($env:USERNAME)
  Log File : $($script:LogFilePath)
================================================================================

"@
    Add-Content -Path $script:LogFilePath -Value $header -Encoding UTF8

    Write-LogInfo "Logging initialised — log file: $($script:LogFilePath)"
}

# ──────────────────────────────────────────────────────────────
# Public: Write-LogInfo
# ──────────────────────────────────────────────────────────────

function Write-LogInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position=0)] [string]$Message)

    Write-LogEntry -Level 'INFO ' -Message $Message -ConsoleColor Cyan
}

# ──────────────────────────────────────────────────────────────
# Public: Write-LogWarning
# ──────────────────────────────────────────────────────────────

function Write-LogWarning {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0)] [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $logMsg = $Message
    if ($ErrorRecord) {
        $logMsg += " | Detail: $($ErrorRecord.Exception.Message)"
        if ($ErrorRecord.InvocationInfo) {
            $logMsg += " [At line $($ErrorRecord.InvocationInfo.ScriptLineNumber)]"
        }
    }

    Write-LogEntry -Level 'WARN ' -Message $logMsg -ConsoleColor Yellow
}

# ──────────────────────────────────────────────────────────────
# Public: Write-LogError
# ──────────────────────────────────────────────────────────────

function Write-LogError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position=0)] [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $logMsg = $Message
    if ($ErrorRecord) {
        $invoc = $ErrorRecord.InvocationInfo
        $location = if ($invoc) { " [Source: $($invoc.ScriptName) Line: $($invoc.ScriptLineNumber)]" } else { "" }
        $logMsg += " | Exception: $($ErrorRecord.Exception.Message)$location"
        
        # For deep debugging, we can log the stack trace to the file only
        if ($script:LogFilePath -and $ErrorRecord.ScriptStackTrace) {
            Add-Content -Path $script:LogFilePath -Value "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) [STACK] $($ErrorRecord.ScriptStackTrace)" -Encoding UTF8
        }
    }

    Write-LogEntry -Level 'ERROR' -Message $logMsg -ConsoleColor Red
}

# ──────────────────────────────────────────────────────────────
# Public: Write-LogSection
# Writes a visible section separator to the log — useful for
# marking the start of a major phase within a script.
# ──────────────────────────────────────────────────────────────

function Write-LogSection {
    [CmdletBinding()]
    param([Parameter(Mandatory, Position=0)] [string]$Title)

    $line = "─── $Title " + ("─" * [Math]::Max(1, 60 - $Title.Length))
    Write-LogEntry -Level 'INFO ' -Message $line -ConsoleColor Cyan
}

# ──────────────────────────────────────────────────────────────
# Public: Write-LogSummary
# Writes a structured results summary block at the end of a run.
# ──────────────────────────────────────────────────────────────

function Write-LogSummary {
    [CmdletBinding()]
    param([hashtable]$Metrics)

    $lines = @("", "── Run Summary " + ("─" * 47))
    foreach ($k in $Metrics.Keys) {
        $lines += "  {0,-30} {1}" -f $k, $Metrics[$k]
    }
    $lines += "─" * 61
    $lines += ""

    foreach ($line in $lines) {
        Write-LogEntry -Level 'INFO ' -Message $line -ConsoleColor Green
    }
}

# ──────────────────────────────────────────────────────────────
# Public: Close-Logging
# Writes the closing footer to the log file.
# Call at the very end of each script (success or failure).
# ──────────────────────────────────────────────────────────────

function Close-Logging {
    [CmdletBinding()]
    param([string]$Status = "Completed")

    $footer = @"

================================================================================
  Status   : $Status
  Finished : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') UTC
================================================================================
"@

    if ($script:LogFilePath -and (Test-Path (Split-Path $script:LogFilePath))) {
        Add-Content -Path $script:LogFilePath -Value $footer -Encoding UTF8
    }
    Write-Host $footer -ForegroundColor Cyan
}

# ──────────────────────────────────────────────────────────────
# Public: Remove-OldLogFiles
# Deletes log files older than RetentionDays in the log directory.
# Call once per script execution, typically at the end.
# ──────────────────────────────────────────────────────────────

function Remove-OldLogFiles {
    [CmdletBinding()]
    param()

    if (-not $script:LogDirectory -or -not (Test-Path $script:LogDirectory)) {
        return
    }

    $cutoff  = (Get-Date).AddDays(-$script:RetentionDays)
    $deleted = 0

    Get-ChildItem -Path $script:LogDirectory -Filter '*.log' | Where-Object {
        $_.LastWriteTime -lt $cutoff
    } | ForEach-Object {
        try {
            Remove-Item $_.FullName -Force
            $deleted++
        }
        catch {
            Write-LogWarning "Could not delete old log file $($_.Name): $($_.Exception.Message)"
        }
    }

    if ($deleted -gt 0) {
        Write-LogInfo "Log rotation: removed $deleted file(s) older than $($script:RetentionDays) days"
    }
}

# ──────────────────────────────────────────────────────────────
# Private: Write-LogEntry
# Core write function — sends to log file and console.
# ──────────────────────────────────────────────────────────────

function Write-LogEntry {
    param(
        [string]$Level,
        [string]$Message,
        [System.ConsoleColor]$ConsoleColor = [System.ConsoleColor]::White
    )

    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$ts [$Level] $Message"

    # Write to log file
    if ($script:LogFilePath) {
        try {
            Add-Content -Path $script:LogFilePath -Value $line -Encoding UTF8
        }
        catch {
            # If file write fails, don't crash the script — just warn on console
            Write-Host "[LOG WRITE FAILED] $line" -ForegroundColor Red
        }
    }

    # Write to Windows Event Log
    if ($script:LogToEventLog) {
        $entryType = switch ($Level.Trim()) {
            'ERROR' { 'Error' }
            'WARN'  { 'Warning' }
            default { 'Information' }
        }
        try {
            Write-EventLog -LogName 'Application' -Source $script:EventSource -EntryType $entryType -EventId 1000 -Message $Message -ErrorAction SilentlyContinue
        } catch {}
    }

    # Write to console
    Write-Host $line -ForegroundColor $ConsoleColor
}

# ──────────────────────────────────────────────────────────────
# Exports
# ──────────────────────────────────────────────────────────────

Export-ModuleMember -Function @(
    'Initialize-Logging',
    'Write-LogInfo',
    'Write-LogWarning',
    'Write-LogError',
    'Write-LogSection',
    'Write-LogSummary',
    'Close-Logging',
    'Remove-OldLogFiles'
)
