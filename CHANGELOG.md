# Changelog

## [1.1.0] - 2025-06-03

### Added
- Comment-based help now includes `.AUTHOR`, `.VERSION`, `.LICENSE`, and `.LASTUPDATED` metadata.
- New parameters: `ArchiveRetentionDays` (for deleting old ZIPs), `ArchiveOnly`, and `DecompressBeforeDelete`.
- Functions for:
  - Logging (`Write-Log`)
  - Event log summary (`Write-EventLogSummary`)
  - Permission checks (`Test-Permissions`)
  - Log rotation (`Invoke-LogRotation`)
  - Directory creation (`Test-OrCreateDirectory`)
  - Old log file discovery (`Get-OldLogFiles`)
  - Free space calculation (`Get-FreeSpace`)
  - Destination path calculation with directory structure preservation (`Get-DestinationPath`)
  - Archive summary reporting (`Write-LogArchiveSummary`)
  - Ensuring archive folders are not NTFS compressed (`Set-ArchiveFolderNotCompressed`)
  - NTFS decompression before archiving (`Optimize-FileForArchiving`)
  - Accurate file size detection (`Get-FileActualSize`)

### Changed
- Major refactor for modularity, error handling, and robustness.
- Progress reporting now includes file counts and percentage complete.
- Improved logging with detailed archive summaries and error reporting.
- Archive folders are now checked and decompressed if NTFS compression is detected.
- Space calculations now include logical, compressed, and archived sizes.
- Log rotation is performed if the log file exceeds a set size.
- Old ZIP archives in archive folders are deleted based on `ArchiveRetentionDays`.

### Fixed
- Directory structure is preserved when using a global destination path.
- Improved detection and handling of NTFS-compressed files.
- More reliable error handling and reporting throughout the script.

---

## [1.0.0] - Initial release

- Compresses and deletes log files older than a specified number of days.
- Archives files to an 'Archive' subfolder by default.
- Basic logging and error handling.
