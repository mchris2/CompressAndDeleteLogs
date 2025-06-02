<#
.SYNOPSIS
    Compresses and deletes log files older than a specified number of days, preserving directory structure.

.DESCRIPTION
    This script searches for log files older than a configurable retention period, compresses them into .zip archives,
    and deletes the originals (unless -ArchiveOnly is specified). It logs all actions and errors to a specified log file.
    By default, archived files are stored in an 'Archive' subfolder within the same directory as the source file.

.PARAMETER SourcePath
    The root directory to search for log files.

.PARAMETER DestinationPath
    The root directory where compressed files will be stored. If not specified, uses 'Archive' subfolder in each source directory.

.PARAMETER LogFilePath
    The path to the log file.

.PARAMETER RetentionDays
    The number of days to retain log files before compressing and deleting.

.PARAMETER ArchiveOnly
    If set, archives logs but does not delete the originals.

.EXAMPLE
    .\CompressAndDeleteLogs.ps1 -SourcePath "C:\inetpub\logs\LogFiles" -RetentionDays 30
    
.EXAMPLE
    .\CompressAndDeleteLogs.ps1 -SourcePath "C:\inetpub\logs\LogFiles" -DestinationPath "E:\Logs" -LogFilePath "D:\Scripts\CompressAndDeleteLogs.log" -RetentionDays 30
#>

param (
    [string]$SourcePath = "C:\inetpub\logs\LogFiles",
    [string]$DestinationPath = "",
    [string]$LogFilePath = "$(Split-Path -Parent $MyInvocation.MyCommand.Path)\CompressAndDeleteLogs.log",
    [int]$RetentionDays = 30,
    [switch]$ArchiveOnly
)

# --- Functions ---

function Write-Log {
    param (
        [string]$Message,
        [string]$LogLevel = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$LogLevel] $Message"
    Write-Host $logMessage
    try {
        Add-Content -Path $LogFilePath -Value $logMessage -Encoding UTF8
    } catch {
        Write-Host "Failed to write to log file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Write-EventLogSummary {
    param (
        [string]$Summary
    )
    $source = "CompressAndDeleteLogs"
    try {
        if (-not [System.Diagnostics.EventLog]::SourceExists($source)) {
            New-EventLog -LogName Application -Source $source
        }
        Write-EventLog -LogName Application -Source $source -EntryType Error -EventId 1001 -Message $Summary
    } catch {
        Write-Log "Could not write to Windows Event Log: $($_.Exception.Message)" "ERROR"
    }
}

function Test-Permissions {
    param (
        [string]$Path,
        [string]$Type = "ReadWrite"
    )
    try {
        if ($Type -eq "ReadWrite") {
            $testFile = Join-Path $Path ".__permtest"
            New-Item -Path $testFile -ItemType File -Force | Out-Null
            Remove-Item -Path $testFile -Force
        } elseif ($Type -eq "Read") {
            Get-ChildItem -Path $Path | Out-Null
        }
        return $true
    } catch {
        return $false
    }
}

function Invoke-LogRotation {
    param (
        [string]$LogFilePath,
        [int]$MaxLogSizeMB = 10
    )
    if (Test-Path $LogFilePath) {
        $logSizeMB = (Get-Item $LogFilePath).Length / 1MB
        if ($logSizeMB -ge $MaxLogSizeMB) {
            $archiveLog = "$LogFilePath.$(Get-Date -Format 'yyyyMMddHHmmss').bak"
            Move-Item -Path $LogFilePath -Destination $archiveLog -Force
        }
    }
}

function Test-OrCreateDirectory {
    param (
        [string]$Path
    )
    if (-not (Test-Path -Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-OldLogFiles {
    param (
        [string]$SourcePath,
        [int]$RetentionDays
    )
    return Get-ChildItem -Path $SourcePath -Recurse -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays) }
}

function Get-FreeSpace {
    param([string]$Path)
    try {
        $drive = (Get-Item $Path).PSDrive
        if ($drive -and $null -ne $drive.Free) {
            return $drive.Free
        }
        # For UNC, try WMI
        $root = [System.IO.Path]::GetPathRoot((Resolve-Path $Path).Path)
        $wmi = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $root.TrimEnd('\') }
        if ($wmi) { return $wmi.FreeSpace }
    } catch { }
    return $null
}

function Get-DestinationPath {
    param (
        [string]$SourceFilePath,
        [string]$GlobalDestinationPath
    )
    if ([string]::IsNullOrWhiteSpace($GlobalDestinationPath)) {
        # Use Archive subfolder in the same directory as the source file
        $sourceDir = Split-Path -Parent $SourceFilePath
        return Join-Path $sourceDir "Archive"
    } else {
        # Use the global destination path with preserved directory structure
        $relativePath = $SourceFilePath.Substring($SourcePath.Length).TrimStart("\")
        $relativeDir = Split-Path -Parent $relativePath
        if ([string]::IsNullOrWhiteSpace($relativeDir)) {
            return $GlobalDestinationPath
        } else {
            return Join-Path $GlobalDestinationPath $relativeDir
        }
    }
}

function Write-LogArchiveSummary {
    param (
        $filesToArchive,
        $filesSkipped,
        $totalOriginalSize,
        $totalArchivedSize,
        $freeSpaceBefore,
        $freeSpaceAfter,
        [int]$FailureCount,
        $JobErrors,
        [switch]$ArchiveOnly
    )
    $totalOriginalSizeMB = [math]::Round($totalOriginalSize / 1MB, 2)
    $totalOriginalSizeGB = [math]::Round($totalOriginalSize / 1GB, 2)
    $totalArchivedSizeMB = [math]::Round($totalArchivedSize / 1MB, 2)
    $totalArchivedSizeGB = [math]::Round($totalArchivedSize / 1GB, 2)
    $spaceSavedMB = [math]::Round(($totalOriginalSize - $totalArchivedSize) / 1MB, 2)
    $spaceSavedGB = [math]::Round(($totalOriginalSize - $totalArchivedSize) / 1GB, 2)

    # Output archived files list to log only
    if ($filesToArchive.Count -gt 0) {
        $filesToArchive | ForEach-Object {
            Write-Log "Archived file: $($_.File) [$($_.SizeMB) MB]"
        }
    }
    if ($filesSkipped.Count -gt 0) {
        $filesSkipped | ForEach-Object {
            Write-Log "Skipped file: $($_.File) [$($_.SizeMB) MB] Reason: $($_.Reason)"
        }
    }
    if ($JobErrors -and $JobErrors.Count -gt 0) {
        foreach ($err in $JobErrors) {
            Write-Log $err "ERROR"
        }
    }

    $summaryLines = @()
    $summaryLines += ""
    $summaryLines += "===== Archive Summary ====="
    $summaryLines += "Files to archive: $($filesToArchive.Count)"
    $summaryLines += "Files skipped: $($filesSkipped.Count)"
    $summaryLines += "Total original size: $totalOriginalSizeMB MB ($totalOriginalSizeGB GB)"
    $summaryLines += "Total archived size: $totalArchivedSizeMB MB ($totalArchivedSizeGB GB)"
    $summaryLines += "Estimated space saved: $spaceSavedMB MB ($spaceSavedGB GB)"
    if ($null -ne $freeSpaceBefore) {
        $summaryLines += "Free space before: $([math]::Round($freeSpaceBefore / 1GB, 2)) GB"
    }
    if ($null -ne $freeSpaceAfter) {
        $summaryLines += "Free space after: $([math]::Round($freeSpaceAfter / 1GB, 2)) GB"
    }
    if ($FailureCount -gt 0) {
        $summaryLines += "Failures: $FailureCount"
    }

    $summaryLines | ForEach-Object {
        Write-Host $_
        Write-Log $_
    }
}

# --- Main Script ---

Invoke-LogRotation -LogFilePath $LogFilePath

# Permissions checks
if (-not (Test-Permissions -Path $SourcePath -Type "Read")) {
    Write-Host "ERROR: No read permission on $SourcePath"
    exit 1
}

# Only check destination permissions if a global destination is specified
if (-not [string]::IsNullOrWhiteSpace($DestinationPath) -and -not (Test-Permissions -Path $DestinationPath -Type "ReadWrite")) {
    Write-Host "ERROR: No write permission on $DestinationPath"
    exit 1
}

if (-not (Test-Permissions -Path (Split-Path $LogFilePath) -Type "ReadWrite")) {
    Write-Host "ERROR: No write permission for log file at $LogFilePath"
    exit 1
}

if (-not (Get-Command Compress-Archive -ErrorAction SilentlyContinue)) {
    Write-Log "Compress-Archive cmdlet not found. PowerShell 5.0 or later is required." "ERROR"
    Write-EventLogSummary "Compress-Archive cmdlet not found. Script failed."
    exit 1
}

$destinationMode = if ([string]::IsNullOrWhiteSpace($DestinationPath)) { "Local Archive folders" } else { "Global destination: $DestinationPath" }
Write-Log "Script started. Source: $SourcePath, Destination mode: $destinationMode, Retention: $RetentionDays days"

# Only create global destination if specified
if (-not [string]::IsNullOrWhiteSpace($DestinationPath)) {
    Test-OrCreateDirectory $DestinationPath
}

# Only write to the event log if the entire script fails (e.g., cannot enumerate files)
try {
    $oldFiles = Get-OldLogFiles -SourcePath $SourcePath -RetentionDays $RetentionDays
    Write-Log "Found $($oldFiles.Count) files older than $RetentionDays days."
} catch {
    Write-Log "Failed to enumerate files: $($_.Exception.Message)" "ERROR"
    # Only here do we write to the event log
    Write-EventLogSummary "Failed to enumerate files: $($_.Exception.Message)"
    exit 1
}

# --- Create destination directories as needed ---
$uniqueDirs = $oldFiles | ForEach-Object {
    Get-DestinationPath -SourceFilePath $_.FullName -GlobalDestinationPath $DestinationPath
} | Sort-Object -Unique

foreach ($dir in $uniqueDirs) {
    Test-OrCreateDirectory $dir
}

# --- Archive files ---
$filesToArchive = @()
$filesSkipped = @()
$jobErrors = @()
$totalOriginalSize = 0
$totalArchivedSize = 0
$failureCount = 0

$freeSpaceBefore = Get-FreeSpace $SourcePath

foreach ($file in $oldFiles) {
    $destinationDir = Get-DestinationPath -SourceFilePath $file.FullName -GlobalDestinationPath $DestinationPath
    $zipFile = Join-Path $destinationDir ($file.BaseName + ".zip")
    
    if (Test-Path $zipFile) {
        Write-Log "Zip file already exists, skipping: $zipFile" "WARNING"
        $filesSkipped += [PSCustomObject]@{
            File = $file.FullName
            Reason = "Zip exists"
            SizeMB = [math]::Round($file.Length / 1MB, 2)
        }
        continue
    }
    
    $totalOriginalSize += $file.Length
    $filesToArchive += [PSCustomObject]@{
        File = $file.FullName
        SizeMB = [math]::Round($file.Length / 1MB, 2)
    }
    
    try {
        Compress-Archive -Path $file.FullName -DestinationPath $zipFile -Force -ErrorAction Stop
        $zipSize = (Get-Item $zipFile).Length
        $totalArchivedSize += $zipSize
        Write-Log "Archived: $($file.FullName) -> $zipFile"
        
        if (-not $ArchiveOnly) {
            Remove-Item -Path $file.FullName -Force
            Write-Log "Deleted original: $($file.FullName)"
        }
    } catch {
        $failureCount++
        $errMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to process $($file.FullName): $($_.Exception.Message)"
        $jobErrors += $errMsg
        Write-Log $errMsg "ERROR"
    }
}

$freeSpaceAfter = Get-FreeSpace $SourcePath

Write-LogArchiveSummary -filesToArchive $filesToArchive `
    -filesSkipped $filesSkipped `
    -totalOriginalSize $totalOriginalSize `
    -totalArchivedSize $totalArchivedSize `
    -freeSpaceBefore $freeSpaceBefore `
    -freeSpaceAfter $freeSpaceAfter `
    -FailureCount $failureCount `
    -JobErrors $jobErrors `
    -ArchiveOnly:$ArchiveOnly

Write-Log "Script completed."