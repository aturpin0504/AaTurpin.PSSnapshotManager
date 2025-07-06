# AaTurpin.PSSnapshotManager

A comprehensive PowerShell module for creating, comparing, and managing network share file snapshots with BITS-based file operations. This module provides enterprise-grade snapshot capture with exclusion pattern support, snapshot comparison for change detection, and efficient file staging operations using Background Intelligent Transfer Service (BITS).

## Features

- **Snapshot Capture**: Create comprehensive snapshots of network shares with pre-compiled exclusion pattern support for optimal performance
- **Change Detection**: Compare snapshots to identify added, modified, and deleted files with detailed reporting
- **BITS Integration**: Leverage Background Intelligent Transfer Service for reliable file transfers with progress tracking and retry logic
- **Staging Operations**: Efficiently copy changed files to staging areas and move them back to original locations
- **Exclusion Support**: Use regex-based exclusion patterns to filter out unwanted directories during snapshot operations
- **Comprehensive Logging**: Detailed logging integration with AaTurpin.PSLogger for audit trails and troubleshooting
- **Error Handling**: Robust error tracking and reporting for access issues and transfer failures
- **Performance Optimized**: Batch processing and pre-compiled patterns for handling large file sets efficiently

## Installation

First, register the NuGet repository if you haven't already:
```powershell
Register-PSRepository -Name "NuGet" -SourceLocation "https://api.nuget.org/v3/index.json" -PublishLocation "https://www.nuget.org/api/v2/package/" -InstallationPolicy Trusted
```
Then install the module:
```powershell
Install-Module -Name AaTurpin.PSSnapshotManager -Repository NuGet -Scope CurrentUser
```
Or for all users (requires administrator privileges):
```powershell
Install-Module -Name AaTurpin.PSSnapshotManager -Repository NuGet -Scope AllUsers
```

## Prerequisites

This module requires the following dependencies (automatically installed):
- **AaTurpin.PSLogger** (v1.0.2+) - Thread-safe logging capabilities
- **AaTurpin.PSConfig** (v1.2.0+) - Configuration management with pre-compiled exclusion patterns
- **BitsTransfer** (v1.0.0.0+) - Background Intelligent Transfer Service support
- **PowerShell 5.1** or later

## Quick Start

### 1. Configure Your Environment

First, set up your configuration file with monitored directories:

```powershell
# Import required modules
Import-Module AaTurpin.PSConfig
Import-Module AaTurpin.PSSnapshotManager

# Create or read configuration
$config = Read-SettingsFile -SettingsPath "settings.json" -LogPath "app.log"

# Add a monitored directory with exclusions
Add-MonitoredDirectory -Path "V:\aeapps\fc_tools" -Exclusions @("temp", "cache", "logs") -LogPath "app.log"
```

### 2. Create a Snapshot

```powershell
# Get the monitored directory configuration
$config = Read-SettingsFile -SettingsPath "settings.json" -LogPath "app.log"
$monitoredDir = $config.monitoredDirectories[0]

# Create snapshot
Get-NetworkShareSnapshot -MonitoredDirectory $monitoredDir -JsonOutputPath "snapshot_$(Get-Date -Format 'yyyyMMdd_HHmmss').json" -AccessErrorsJsonPath "errors_$(Get-Date -Format 'yyyyMMdd_HHmmss').json" -LogPath "app.log"
```

### 3. Compare Snapshots

```powershell
# Compare two snapshots
Compare-NetworkShareSnapshots -Snapshot1Path "snapshot_20250706_100000.json" -Snapshot2Path "snapshot_20250706_140000.json" -OutputPath "comparison_results.json" -LogPath "app.log"
```

### 4. Stage Changed Files

```powershell
# Copy changed files to staging area
Copy-SnapshotChangesToStaging -ComparisonResultsPath "comparison_results.json" -StagingArea "C:\StagingArea" -LogPath "app.log"
```

### 5. Deploy Staged Files

```powershell
# Move files from staging back to network shares
Move-StagingFilesToNetwork -StagingArea "C:\StagingArea" -SuccessLogPath "move_success.log" -FailureLogPath "move_failures.log" -LogPath "app.log"
```

## Core Functions

### Get-NetworkShareSnapshot

Captures a comprehensive snapshot of files in a monitored directory with exclusion pattern support.

```powershell
Get-NetworkShareSnapshot -MonitoredDirectory $monitoredDir -JsonOutputPath "snapshot.json" -AccessErrorsJsonPath "errors.json" -LogPath "app.log"
```

**Key Features:**
- Uses pre-compiled regex patterns for optimal performance
- Handles access errors gracefully
- Exports detailed metadata including file attributes, sizes, and timestamps
- Batch processing for large file sets

### Compare-NetworkShareSnapshots

Compares two snapshots to identify changes with detailed analysis.

```powershell
Compare-NetworkShareSnapshots -Snapshot1Path "before.json" -Snapshot2Path "after.json" -OutputPath "changes.json" -Snapshot1AccessErrorsPath "before_errors.json" -Snapshot2AccessErrorsPath "after_errors.json" -LogPath "app.log"
```

**Detected Changes:**
- **Added Files**: New files in the second snapshot
- **Modified Files**: Files with changed size, timestamp, or attributes
- **Deleted Files**: Files present in first snapshot but missing in second
- **Unchanged Files**: Automatically excluded from results

### Copy-SnapshotChangesToStaging

Copies modified and added files to a staging area using BITS for reliable transfers.

```powershell
Copy-SnapshotChangesToStaging -ComparisonResultsPath "changes.json" -StagingArea "C:\Staging" -LogPath "app.log" -Priority "High" -MaxConcurrentJobs 5
```

**Features:**
- BITS-based transfers with retry logic
- Preserves original directory structure
- Configurable transfer priority and concurrency
- Progress tracking and detailed reporting

### Move-StagingFilesToNetwork

Moves files from staging area back to their original network locations.

```powershell
Move-StagingFilesToNetwork -StagingArea "C:\Staging" -SuccessLogPath "success.log" -FailureLogPath "failures.log" -LogPath "app.log"
```

**Capabilities:**
- Validates destination drive accessibility
- Handles file collisions automatically
- Separate logging for successful and failed operations
- Atomic move operations (copy + delete source)

## Configuration Integration

This module integrates seamlessly with **AaTurpin.PSConfig** for managing monitored directories and exclusion patterns:

```json
{
  "stagingArea": "C:\\StagingArea",
  "driveMappings": [
    {
      "letter": "V",
      "path": "\\\\server\\share"
    }
  ],
  "monitoredDirectories": [
    {
      "path": "V:\\aeapps\\fc_tools",
      "exclusions": ["temp", "cache", "logs", "backup"]
    }
  ]
}
```

The configuration system automatically pre-compiles exclusion patterns into regex objects for optimal performance during file filtering operations.

## Advanced Usage

### Batch Snapshot Processing

```powershell
# Process multiple monitored directories
$config = Read-SettingsFile -LogPath "app.log"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

foreach ($dir in $config.monitoredDirectories) {
    $dirName = Split-Path $dir.path -Leaf
    $snapshotPath = "snapshots\${dirName}_${timestamp}.json"
    $errorsPath = "snapshots\${dirName}_errors_${timestamp}.json"
    
    Get-NetworkShareSnapshot -MonitoredDirectory $dir -JsonOutputPath $snapshotPath -AccessErrorsJsonPath $errorsPath -LogPath "app.log"
}
```

### Custom BITS Configuration

```powershell
# High-priority transfer with custom settings
Copy-SnapshotChangesToStaging -ComparisonResultsPath "changes.json" -StagingArea "D:\FastStaging" -LogPath "app.log" -Priority "Foreground" -RetryInterval 30 -RetryTimeout 600 -MaxConcurrentJobs 1
```

### Error Analysis

```powershell
# Analyze access errors from snapshot
$errors = Get-Content "errors.json" | ConvertFrom-Json
$errors.AccessErrors | Group-Object ErrorType | Select-Object Name, Count
```

## Performance Considerations

- **Exclusion Patterns**: Pre-compiled regex patterns provide significant performance improvements for large directory structures
- **Batch Processing**: Files are processed in configurable batches (default: 50,000) to optimize memory usage
- **BITS Transfers**: Concurrent job limits prevent resource exhaustion while maximizing throughput
- **Memory Management**: Automatic cleanup of compiled patterns and job objects

## Error Handling

The module provides comprehensive error handling:

- **Access Errors**: Tracked separately and excluded from comparisons
- **Transfer Failures**: Detailed logging with retry mechanisms
- **Partial Failures**: Copy succeeded but source deletion failed scenarios
- **Drive Accessibility**: Validation before attempting move operations

## Logging Integration

All operations integrate with **AaTurpin.PSLogger** for consistent, thread-safe logging:

```powershell
# Example log output
[2025-07-06 14:30:15.123] [Information] Starting network share snapshot for: V:\aeapps\fc_tools
[2025-07-06 14:30:15.456] [Information] Using 4 pre-compiled exclusion patterns
[2025-07-06 14:30:45.789] [Information] File enumeration completed: 15,234 files, 2 errors
[2025-07-06 14:31:20.012] [Information] Snapshot export completed successfully: snapshot.json (2.3 MB)
```

## Output Formats

### Snapshot JSON Structure

```json
{
  "Metadata": {
    "SnapshotDate": "2025-07-06 14:30:15",
    "SnapshotDateUtc": "2025-07-06 18:30:15",
    "NetworkSharePath": "V:\\aeapps\\fc_tools",
    "TotalItems": 15234,
    "AccessErrorsCount": 2
  },
  "Files": [
    {
      "FullPath": "V:\\aeapps\\fc_tools\\example.txt",
      "Name": "example.txt",
      "Type": "File",
      "SizeBytes": 1024,
      "SizeKB": 1.0,
      "SizeMB": 0.0,
      "LastWriteTime": "2025-07-06 12:00:00",
      "CreationTime": "2025-07-06 10:00:00",
      "Attributes": "Archive",
      "Extension": ".txt",
      "Directory": "V:\\aeapps\\fc_tools"
    }
  ]
}
```

### Comparison Results Structure

```json
{
  "Metadata": {
    "ComparisonDate": "2025-07-06 15:00:00",
    "Snapshot1Date": "2025-07-06 14:00:00",
    "Snapshot2Date": "2025-07-06 14:30:00",
    "TotalComparisons": 156
  },
  "Results": [
    {
      "FullPath": "V:\\aeapps\\fc_tools\\modified.txt",
      "Status": "Modified",
      "OldSizeBytes": 1024,
      "NewSizeBytes": 2048,
      "SizeDifferenceBytes": 1024,
      "ModificationReason": "Size changed from 1024 to 2048 bytes"
    }
  ]
}
```

## Best Practices

1. **Regular Snapshots**: Create snapshots at regular intervals to track changes over time
2. **Exclusion Patterns**: Use specific exclusion patterns to avoid unnecessary processing of temporary files
3. **Staging Management**: Regularly clean up staging areas after successful deployments
4. **Log Monitoring**: Monitor access error logs to identify permission or connectivity issues
5. **BITS Configuration**: Adjust concurrency and priority based on network conditions and system resources

## Troubleshooting

### Common Issues

**Large File Sets**: For directories with millions of files, consider:
- Increasing batch sizes
- Using more specific exclusion patterns
- Running during off-peak hours

**Network Connectivity**: For unreliable networks:
- Increase retry timeouts
- Reduce concurrent job counts
- Use higher BITS priority

**Permission Issues**: For access denied errors:
- Check service account permissions
- Verify network drive mappings
- Review exclusion patterns for system directories

### Diagnostic Commands

```powershell
# Check BITS transfer status
Get-BitsTransfer | Where-Object { $_.DisplayName -like "*Snapshot*" }

# Analyze file size distribution
$snapshot = Get-Content "snapshot.json" | ConvertFrom-Json
$snapshot.Files | Measure-Object -Property SizeBytes -Sum -Average

# Review access errors
$errors = Get-Content "errors.json" | ConvertFrom-Json
$errors.AccessErrors | Group-Object ErrorType
```

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit issues, feature requests, or pull requests.

## Support

For support and questions, please visit the [project repository](https://github.com/aturpin0504/AaTurpin.PSSnapshotManager).