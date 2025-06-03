<#
.SYNOPSIS
    Compress and manage log files by archiving and optionally deleting them after a retention period.

.DESCRIPTION
    This script searches for log files older than a specified retention period, compresses them into .zip archives,
    and deletes the originals (unless -ArchiveOnly is specified). It supports NTFS decompression before deletion,
    configurable archive retention, and robust logging to both file and Windows Event Log.
    By default, archives are stored in an 'Archive' subfolder within each source folder, but a global destination can be specified.

.AUTHOR
    Chris McCorrin

.VERSION
    1.1.0

.LICENCE
    MIT Licence

.LASTUPDATED
    2025-06-03

.PARAMETER SourcePath
    The root folder to search for log files.

.PARAMETER DestinationPath
    The root folder where compressed files will be stored. If not specified, uses an 'Archive' subfolder in each source folder.

.PARAMETER LogFilePath
    The path to the log file. Defaults to 'CompressAndDeleteLogs.log' in the script folder.

.PARAMETER RetentionDays
    The number of days to retain log files before compressing and deleting.

.PARAMETER ArchiveRetentionDays
    The number of days to retain archived ZIP files before deleting them from the archive folder.

.PARAMETER ArchiveOnly
    If set, archives logs but does not delete the originals.

.PARAMETER DecompressBeforeDelete
    If set, decompresses NTFS-compressed files before deletion to free up the full logical size.

.EXAMPLE
    .\CompressAndDeleteLogs.ps1 -SourcePath "C:\inetpub\logs\LogFiles" -RetentionDays 30 -DecompressBeforeDelete

.EXAMPLE
    .\CompressAndDeleteLogs.ps1 -SourcePath "C:\inetpub\logs\LogFiles" -DestinationPath "E:\Logs" -LogFilePath "D:\Scripts\CompressAndDeleteLogs.log" -RetentionDays 30

.NOTES
    v1.1.0 (2025-06-03): Major refactor with modular functions, improved logging, NTFS decompression support, archive retention, and robust error handling.
#>

param (
    [string]$SourcePath = "C:\inetpub\logs\LogFiles",
    [string]$DestinationPath = "",
    [string]$LogFilePath = "$(Split-Path -Parent $MyInvocation.MyCommand.Path)\CompressAndDeleteLogs.log",
    [int]$RetentionDays = 30,
    [int]$ArchiveRetentionDays = 90,
    [switch]$ArchiveOnly,
    [switch]$DecompressBeforeDelete
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
    # Common log file extensions
    $logExtensions = @(".log", ".txt", ".out", ".err", ".trace")
    
    return Get-ChildItem -Path $SourcePath -Recurse -File | Where-Object { 
        # Only include files older than retention period
        $_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays) -and
        # Only include common log file extensions
        $logExtensions -contains $_.Extension -and
        # Exclude files in Archive folders
        $_.DirectoryName -notlike "*\Archive" -and
        $_.DirectoryName -notlike "*\Archive\*"
    }
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
        [string]$GlobalDestinationPath,
        [string]$SourceRootPath  # Add this parameter
    )
    if ([string]::IsNullOrWhiteSpace($GlobalDestinationPath)) {
        # Use Archive subfolder in the same directory as the source file
        $sourceDir = Split-Path -Parent $SourceFilePath
        return Join-Path $sourceDir "Archive"
    } else {
        # Use the global destination path with preserved directory structure
        $relativePath = $SourceFilePath.Substring($SourceRootPath.Length).TrimStart("\")
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
        $totalOriginalLogicalSize,
        $totalOriginalCompressedSize,
        $totalArchivedSize,
        $freeSpaceBefore,
        $freeSpaceAfter,
        [int]$FailureCount,
        $JobErrors,
        [long]$DeletedZipSize = 0,
        [switch]$ArchiveOnly

    )
    $totalOriginalLogicalSizeMB = [math]::Round($totalOriginalLogicalSize / 1MB, 2)
    $totalOriginalLogicalSizeGB = [math]::Round($totalOriginalLogicalSize / 1GB, 2)
    $totalOriginalCompressedSizeMB = [math]::Round($totalOriginalCompressedSize / 1MB, 2)
    $totalOriginalCompressedSizeGB = [math]::Round($totalOriginalCompressedSize / 1GB, 2)
    $totalArchivedSizeMB = [math]::Round($totalArchivedSize / 1MB, 2)
    $totalArchivedSizeGB = [math]::Round($totalArchivedSize / 1GB, 2)
    $deletedZipSizeMB = [math]::Round($DeletedZipSize / 1MB, 2)

    $spaceSavedFromLogicalMB = [math]::Round(($totalOriginalLogicalSize - $totalArchivedSize) / 1MB, 2)
    $spaceSavedFromLogicalGB = [math]::Round(($totalOriginalLogicalSize - $totalArchivedSize) / 1GB, 2)
    $spaceSavedFromCompressedMB = [math]::Round(($totalOriginalCompressedSize - $totalArchivedSize) / 1MB, 2)
    $spaceSavedFromCompressedGB = [math]::Round(($totalOriginalCompressedSize - $totalArchivedSize) / 1GB, 2)

    # Output detailed file information
    if ($filesToArchive.Count -gt 0) {
        $filesToArchive | ForEach-Object {
            $compressionText = if ($_.IsNTFSCompressed) { " (was NTFS compressed: $($_.CompressedSizeMB) MB)" } else { "" }
            Write-Log "Archived file: $($_.File) [Logical: $($_.LogicalSizeMB) MB$compressionText -> ZIP: $($_.ZipSizeMB) MB, Ratio: $($_.CompressionRatio)%]"
        }
    }
    if ($filesSkipped.Count -gt 0) {
        $filesSkipped | ForEach-Object {
            $compressionText = if ($_.IsNTFSCompressed) { " (NTFS compressed: $($_.CompressedSizeMB) MB)" } else { "" }
            Write-Log "Skipped file: $($_.File) [Logical: $($_.LogicalSizeMB) MB$compressionText] Reason: $($_.Reason)"
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
    $summaryLines += "Files archived: $($filesToArchive.Count)"
    $summaryLines += "Files skipped: $($filesSkipped.Count)"
    $summaryLines += "Original logical size: $totalOriginalLogicalSizeMB MB ($totalOriginalLogicalSizeGB GB)"
    $summaryLines += "Original size on disk: $totalOriginalCompressedSizeMB MB ($totalOriginalCompressedSizeGB GB)"
    $summaryLines += "Final ZIP archive size: $totalArchivedSizeMB MB ($totalArchivedSizeGB GB)"
    $summaryLines += "Space saved vs logical: $spaceSavedFromLogicalMB MB ($spaceSavedFromLogicalGB GB)"
    $summaryLines += "Space saved vs disk usage: $spaceSavedFromCompressedMB MB ($spaceSavedFromCompressedGB GB)"
    $summaryLines += "Total size of deleted archived ZIP files: $deletedZipSizeMB MB"
    
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

function Set-ArchiveFolderNotCompressed {
    param (
        [string]$ArchivePath
    )
    try {
        # Create the directory if it doesn't exist
        Test-OrCreateDirectory $ArchivePath
        
        # Check if the Archive folder has NTFS compression enabled
        $archiveDir = Get-Item $ArchivePath
        if ($archiveDir.Attributes -band [System.IO.FileAttributes]::Compressed) {
            Write-Log "Archive folder is NTFS compressed, removing compression: $ArchivePath" "INFO"
            
            # Remove compression from the Archive folder
            $compactResult = & compact.exe /u "`"$ArchivePath`"" 2>&1
            $compactExitCode = $LASTEXITCODE
            
            if ($compactExitCode -eq 0) {
                Write-Log "Successfully removed NTFS compression from Archive folder: $ArchivePath" "INFO"
            } else {
                Write-Log "Warning: Failed to remove NTFS compression from Archive folder: $($compactResult -join '; ')" "WARNING"
            }
        }
    } catch {
        Write-Log "Error checking/fixing Archive folder compression: $($_.Exception.Message)" "WARNING"
    }
}

function Optimize-FileForArchiving {
    param (
        [string]$FilePath
    )
    try {
        # Check if compact.exe is available
        if (-not (Get-Command compact.exe -ErrorAction SilentlyContinue)) {
            Write-Log "compact.exe not found in PATH" "WARNING"
            $sizeInfo = Get-FileActualSize -FilePath $FilePath
            return @{
                Success = $true
                OriginalLogicalSize = $sizeInfo.LogicalSize
                OriginalCompressedSize = $sizeInfo.CompressedSize
                FinalSize = $sizeInfo.LogicalSize
                WasDecompressed = $false
                SpaceIncrease = 0
                Error = "compact.exe not available"
            }
        }

        $sizeInfo = Get-FileActualSize -FilePath $FilePath
        
        if ($sizeInfo.IsCompressed) {
            Write-Log "Decompressing NTFS compressed file in place: $FilePath" "INFO"
            
            # Decompress the file in place using compact.exe
            $compactResult = & compact.exe /u "`"$FilePath`"" 2>&1
            $compactExitCode = $LASTEXITCODE
            
            if ($compactExitCode -eq 0) {
                # Verify decompression
                Start-Sleep -Milliseconds 200
                $newFile = Get-Item $FilePath
                $stillCompressed = $newFile.Attributes -band [System.IO.FileAttributes]::Compressed
                
                if (-not $stillCompressed) {
                    Write-Log "Successfully decompressed in place: $FilePath (New size: $([math]::Round($newFile.Length / 1MB, 2)) MB)" "INFO"
                    return @{
                        Success = $true
                        OriginalLogicalSize = $sizeInfo.LogicalSize
                        OriginalCompressedSize = $sizeInfo.CompressedSize
                        FinalSize = $newFile.Length
                        WasDecompressed = $true
                        SpaceIncrease = $newFile.Length - $sizeInfo.CompressedSize
                    }
                } else {
                    return @{
                        Success = $false
                        Error = "File still appears compressed after compact.exe"
                        OriginalLogicalSize = $sizeInfo.LogicalSize
                        OriginalCompressedSize = $sizeInfo.CompressedSize
                        FinalSize = $sizeInfo.LogicalSize
                        WasDecompressed = $false
                        SpaceIncrease = 0
                    }
                }
            } else {
                return @{
                    Success = $false
                    Error = "compact.exe failed with exit code $compactExitCode`: $($compactResult -join '; ')"
                    OriginalLogicalSize = $sizeInfo.LogicalSize
                    OriginalCompressedSize = $sizeInfo.CompressedSize
                    FinalSize = $sizeInfo.LogicalSize
                    WasDecompressed = $false
                    SpaceIncrease = 0
                }
            }
        } else {
            # File is not compressed, use as-is
            return @{
                Success = $true
                OriginalLogicalSize = $sizeInfo.LogicalSize
                OriginalCompressedSize = $sizeInfo.CompressedSize
                FinalSize = $sizeInfo.LogicalSize
                WasDecompressed = $false
                SpaceIncrease = 0
            }
        }
    } catch {
        return @{
            Success = $false
            Error = "Exception in Optimize-FileForArchiving: $($_.Exception.Message)"
            OriginalLogicalSize = 0
            OriginalCompressedSize = 0
            FinalSize = 0
            WasDecompressed = $false
            SpaceIncrease = 0
        }
    }
}

function Get-FileActualSize {
    param (
        [string]$FilePath
    )
    try {
        # Get both logical size and compressed size
        $file = Get-Item $FilePath
        $logicalSize = $file.Length
        
        # Use multiple methods to detect NTFS compression
        $isCompressed = $false
        $compressedSize = $logicalSize
        
        # Method 1: Check file attributes
        if ($file.Attributes -band [System.IO.FileAttributes]::Compressed) {
            $isCompressed = $true
            
            # Method 2: Use WMI to get actual compressed size
            try {
                $escapedPath = $FilePath.Replace('\','\\').Replace("'","''")
                $wmiFile = Get-WmiObject -Class CIM_DataFile -Filter "Name='$escapedPath'" -ErrorAction Stop
                if ($wmiFile -and $wmiFile.CompressedFileSize -gt 0) {
                    $compressedSize = $wmiFile.CompressedFileSize
                } else {
                    # Method 3: Use compact.exe to get size info
                    $compactResult = & compact.exe "$FilePath" 2>$null
                    if ($compactResult -and $compactResult.Count -gt 1) {
                        # Parse compact output for size information
                        $sizeLine = $compactResult | Where-Object { $_ -match '\d+:\d+' }
                        if ($sizeLine -and $sizeLine -match '(\d+):(\d+)') {
                            $compressedSize = [long]$matches[1]
                        }
                    }
                }
            } catch {
                # If WMI fails, estimate based on typical compression ratios
                Write-Log "WMI query failed for $FilePath, using file attributes only" "WARNING"
                $compressedSize = [math]::Round($logicalSize * 0.3) # Estimate 30% of original
            }
        }
        
        return @{
            LogicalSize = $logicalSize
            CompressedSize = $compressedSize
            IsCompressed = $isCompressed
        }
    } catch {
        Write-Log "Error getting file size for $FilePath`: $($_.Exception.Message)" "ERROR"
        return @{
            LogicalSize = if ($file) { $file.Length } else { 0 }
            CompressedSize = if ($file) { $file.Length } else { 0 }
            IsCompressed = $false
        }
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

# --- Create destination directories and ensure they're not compressed ---
$uniqueDirs = $oldFiles | ForEach-Object {
    Get-DestinationPath -SourceFilePath $_.FullName -GlobalDestinationPath $DestinationPath -SourceRootPath $SourcePath
} | Sort-Object -Unique

foreach ($dir in $uniqueDirs) {
    Set-ArchiveFolderNotCompressed $dir
}

# --- Archive files ---
$filesToArchive = @()
$filesSkipped = @()
$jobErrors = @()
$totalOriginalLogicalSize = 0
$totalOriginalCompressedSize = 0
$totalArchivedSize = 0
$failureCount = 0

$freeSpaceBefore = Get-FreeSpace $SourcePath

# Initialize progress tracking
$totalFiles = $oldFiles.Count
$currentFile = 0

foreach ($file in $oldFiles) {
    $currentFile++
    $progressPercent = [math]::Round(($currentFile / $totalFiles) * 100, 1)
    
    # Update progress with current file info
    $progressActivity = "Processing log files ($currentFile of $totalFiles)"
    $progressStatus = "Current: $($file.Name) - $progressPercent% Complete"
    Write-Progress -Activity $progressActivity -Status $progressStatus -PercentComplete $progressPercent
    
    $destinationDir = Get-DestinationPath -SourceFilePath $file.FullName -GlobalDestinationPath $DestinationPath -SourceRootPath $SourcePath
    $zipFile = Join-Path $destinationDir ($file.BaseName + ".zip")
    
    if (Test-Path $zipFile) {
        Write-Log "Zip file already exists, will be overwritten: $zipFile" "INFO"
    }

    try {
        # Before decompressing:
        $sizeInfo = Get-FileActualSize -FilePath $file.FullName
        $totalOriginalLogicalSize += $sizeInfo.LogicalSize
        $totalOriginalCompressedSize += $sizeInfo.CompressedSize
        
        # Step 1: Decompress file in place if needed (for better ZIP compression and to free up logical size)
        $optimizeResult = Optimize-FileForArchiving -FilePath $file.FullName
        
        if (-not $optimizeResult.Success) {
            Write-Log "Failed to optimize file for archiving: $($optimizeResult.Error)" "WARNING"
            # Continue with original file even if decompression failed
        }
        
        $totalOriginalLogicalSize += $optimizeResult.OriginalLogicalSize
        $totalOriginalCompressedSize += $optimizeResult.OriginalCompressedSize
        
        # Step 2: Compress to ZIP archive
        Compress-Archive -Path $file.FullName -DestinationPath $zipFile -Force -ErrorAction Stop
        
        # Preserve the original file's timestamps on the zip file
        $zipFileItem = Get-Item $zipFile
        $zipFileItem.CreationTime = $file.CreationTime
        $zipFileItem.LastWriteTime = $file.LastWriteTime
        $zipFileItem.LastAccessTime = $file.LastAccessTime
        
        $zipSize = $zipFileItem.Length
        $totalArchivedSize += $zipSize
        
        # Calculate compression ratio based on the actual file size at compression time
        $compressionRatio = [math]::Round(($zipSize / $optimizeResult.FinalSize) * 100, 1)
        
        # Archive files tracking
        $filesToArchive += [PSCustomObject]@{
            File = $file.FullName
            LogicalSizeMB = [math]::Round($optimizeResult.OriginalLogicalSize / 1MB, 2)
            CompressedSizeMB = [math]::Round($optimizeResult.OriginalCompressedSize / 1MB, 2)
            ZipSizeMB = [math]::Round($zipSize / 1MB, 2)
            IsNTFSCompressed = $optimizeResult.WasDecompressed
            CompressionRatio = $compressionRatio
        }
        
        $compressionRatio = if ($optimizeResult.FinalSize -gt 0) {
            [math]::Round($zipSize / 1MB, 2).ToString() + "MB (" + [math]::Round(($zipSize / $optimizeResult.FinalSize) * 100, 1) + "%)"
        } else {
            [math]::Round($zipSize / 1MB, 2).ToString() + "MB"
        }
        Write-Log ("Archived: {0} -> {1} - Ratio: {2}" -f $file.FullName, $zipFile, $compressionRatio)

        if (-not $ArchiveOnly) {
            $spaceFreed = $optimizeResult.FinalSize
            Remove-Item -Path $file.FullName -Force

            $decompText = if ($optimizeResult.WasDecompressed) { " (decompressed)" } else { "" }
            Write-Log ("Deleted original: {0}{1} - Space freed: {2} MB" -f $file.FullName, $decompText, [math]::Round($spaceFreed / 1MB, 2))
        }
        
    } catch {
        $errorMsg = "Error processing file $($file.FullName): $($_.Exception.Message)"
        Write-Log $errorMsg "ERROR"
        $jobErrors += $errorMsg
        $failureCount++
    }
}

# Complete the progress bar
Write-Progress -Activity "Processing log files" -Status "Complete" -PercentComplete 100 -Completed

$freeSpaceAfter = Get-FreeSpace $SourcePath

# --- Delete old ZIP files in Archive folders ---
Write-Log "Checking for archived ZIP files older than $ArchiveRetentionDays days to delete..."

$now = Get-Date
$deletedArchives = 0
$deletedZipSize = 0
$archiveFolders = $uniqueDirs | Where-Object { $_ -match '\\Archive($|\\)' }

foreach ($archiveFolder in $archiveFolders) {
    if (Test-Path $archiveFolder) {
        $oldZips = Get-ChildItem -Path $archiveFolder -Recurse -File -Filter *.zip | Where-Object {
            $_.LastWriteTime -lt $now.AddDays(-$ArchiveRetentionDays)
        }
        foreach ($zip in $oldZips) {
            $deletedZipSize += $zip.Length
            try {
                Remove-Item -Path $zip.FullName -Force
                Write-Log "Deleted archived ZIP: $($zip.FullName) (older than $ArchiveRetentionDays days)"
                $deletedArchives++
            } catch {
                Write-Log "Failed to delete archived ZIP: $($zip.FullName) - $($_.Exception.Message)" "WARNING"
            }
        }
    }
}

Write-Log "Deleted $deletedArchives archived ZIP files older than $ArchiveRetentionDays days."

Write-LogArchiveSummary -filesToArchive $filesToArchive `
    -filesSkipped $filesSkipped `
    -totalOriginalLogicalSize $totalOriginalLogicalSize `
    -totalOriginalCompressedSize $totalOriginalCompressedSize `
    -totalArchivedSize $totalArchivedSize `
    -freeSpaceBefore $freeSpaceBefore `
    -freeSpaceAfter $freeSpaceAfter `
    -FailureCount $failureCount `
    -JobErrors $jobErrors `
    -ArchiveOnly:$ArchiveOnly `
    -DeletedZipSize $deletedZipSize

Write-Log "Script completed $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log "=================================================================="
Write-Log ""