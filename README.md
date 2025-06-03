# CompressAndDeleteLogs

A PowerShell script to compress and manage log files by archiving and optionally deleting them after a retention period. Designed for Windows environments, this script is ideal for managing IIS and other application logs, helping you save disk space and maintain tidy log directories.

## Features

- **Compresses log files** older than a specified number of days into `.zip` archives.
- **Deletes original log files** after archiving (unless `-ArchiveOnly` is specified).
- **Deletes old archives**: Removes archived ZIP files older than a configurable retention period.
- **Supports NTFS decompression**: Optionally decompresses NTFS-compressed files before deletion to reclaim full logical space.
- **Flexible archive location**: Archives can be stored in an `Archive` subfolder within each source folder or in a global destination.
- **Robust logging**: Logs actions and errors to a log file and the Windows Event Log.
- **Progress reporting**: Displays progress and summary information during execution.
- **Safe and idempotent**: Handles errors gracefully and checks permissions before making changes.

## Requirements

- Windows PowerShell 5.0 or later (requires `Compress-Archive` cmdlet)
- Sufficient permissions to read, write, and delete files in the source and destination folders

## Usage

```powershell
.\CompressAndDeleteLogs.ps1 -SourcePath "C:\inetpub\logs\LogFiles" -RetentionDays 30 -DecompressBeforeDelete
```

**Parameters:**

- `-SourcePath`  
  The root folder to search for log files.

- `-DestinationPath`  
  (Optional) The root folder where compressed files will be stored. If not specified, uses an `Archive` subfolder in each source folder.

- `-LogFilePath`  
  (Optional) The path to the log file. Defaults to `CompressAndDeleteLogs.log` in the script folder.

- `-RetentionDays`  
  The number of days to retain log files before compressing and deleting.

- `-ArchiveRetentionDays`  
  The number of days to retain archived ZIP files before deleting them from the archive folder.

- `-ArchiveOnly`  
  If set, archives logs but does not delete the originals.

- `-DecompressBeforeDelete`  
  If set, decompresses NTFS-compressed files before deletion to free up the full logical size.

## Examples

Archive and delete logs older than 30 days, decompressing NTFS-compressed files before deletion:

```powershell
.\CompressAndDeleteLogs.ps1 -SourcePath "C:\inetpub\logs\LogFiles" -RetentionDays 30 -DecompressBeforeDelete
```

Archive logs to a global destination and log to a custom file:

```powershell
.\CompressAndDeleteLogs.ps1 -SourcePath "C:\inetpub\logs\LogFiles" -DestinationPath "E:\Logs" -LogFilePath "D:\Scripts\CompressAndDeleteLogs.log" -RetentionDays 30
```

## Logging

- Actions and errors are logged to the specified log file (default: `CompressAndDeleteLogs.log`).
- Critical errors are also written to the Windows Event Log under the source `CompressAndDeleteLogs`.

## Notes

- Ensure you have appropriate permissions to run the script and modify files in the specified directories.
- The script is safe to run multiple times; it will not re-archive or re-delete files unnecessarily.

## Version

- **1.1.0** (2025-06-03): Major refactor with modular functions, improved logging, NTFS decompression support, archive retention, and robust error handling.

## Licence

MIT Licence

---

**Author:** Chris McMorrin

For more information or to report issues, please use the repository's issue tracker.