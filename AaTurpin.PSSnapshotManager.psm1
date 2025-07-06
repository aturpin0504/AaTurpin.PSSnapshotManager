function Get-FilesWithExclusions {
    <#
    .SYNOPSIS
        Gets all files from a monitored directory with exclusion support and error tracking.
        Optimized to use pre-compiled regex patterns for maximum performance.
    
    .DESCRIPTION
        This function efficiently retrieves files from a monitored directory while respecting exclusion patterns.
        Uses pre-compiled regex patterns for optimal performance when available.
        Returns a custom object containing both the file results and any access errors encountered.
    
    .PARAMETER MonitoredDirectory
        A monitored directory object from Read-SettingsFile containing path and compiledExclusionPatterns properties.
    
    .PARAMETER LogPath
        The path to the log file where operations will be logged using PSLogger.
    
    .OUTPUTS
        PSCustomObject with properties:
        - Files: System.Collections.Generic.List[System.IO.FileInfo] containing the filtered file list
        - AccessErrors: System.Collections.Generic.List[PSCustomObject] containing any access errors encountered
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if (-not $_.PSObject.Properties.Name -contains "path") {
                throw "MonitoredDirectory must have a 'path' property"
            }
            if (-not $_.PSObject.Properties.Name -contains "compiledExclusionPatterns") {
                throw "MonitoredDirectory must have a 'compiledExclusionPatterns' property. Use Read-SettingsFile to ensure proper initialization."
            }
            return $true
        })]
        $MonitoredDirectory,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )
    
    $path = $MonitoredDirectory.path
    $compiledExclusionPatterns = $MonitoredDirectory.compiledExclusionPatterns
    
    Write-LogDebug -LogPath $LogPath -Message "Starting file enumeration for path: $path"
    Write-LogDebug -LogPath $LogPath -Message "Using $($compiledExclusionPatterns.Count) compiled exclusion patterns"
    
    # Pre-allocate with estimated capacity for better performance
    $files = [System.Collections.Generic.List[object]]::new(10000)
    $errors = [System.Collections.Generic.List[object]]::new()
    
    try {
        $dirInfo = [System.IO.DirectoryInfo]::new($path)
        
        try {
            Write-LogDebug -LogPath $LogPath -Message "Enumerating all files in directory structure"
            $allFiles = $dirInfo.GetFiles("*", [System.IO.SearchOption]::AllDirectories)
            Write-LogInfo -LogPath $LogPath -Message "Found $($allFiles.Count) total files in $path"
            
            # If no exclusions, add all files directly - optimized path
            if ($compiledExclusionPatterns.Count -eq 0) {
                Write-LogDebug -LogPath $LogPath -Message "No exclusions configured, adding all files"
                $files.AddRange($allFiles)
            }
            else {
                Write-LogDebug -LogPath $LogPath -Message "Applying exclusion patterns to filter files"
                # Pre-calculate base path for relative path calculations
                $basePath = $dirInfo.FullName.TrimEnd('\')
                $baseLength = $basePath.Length + 1
                
                $excludedCount = 0
                
                # Use LINQ-style filtering for better performance with large file sets
                foreach ($file in $allFiles) {
                    $relativePath = if ($file.DirectoryName.Length -gt $baseLength) { 
                        $file.DirectoryName.Substring($baseLength).ToLowerInvariant()
                    } else { 
                        [string]::Empty 
                    }
                    
                    # Optimized exclusion check - exit early on first match
                    $excluded = $false
                    for ($i = 0; $i -lt $compiledExclusionPatterns.Count; $i++) {
                        try {
                            if ($compiledExclusionPatterns[$i].IsMatch($relativePath)) {
                                $excluded = $true
                                $excludedCount++
                                break
                            }
                        }
                        catch {
                            Write-LogWarning -LogPath $LogPath -Message "Error matching exclusion pattern $i against path '$relativePath'" -Exception $_.Exception
                            continue
                        }
                    }
                    
                    if (-not $excluded) {
                        $files.Add($file)
                    }
                }
                
                Write-LogInfo -LogPath $LogPath -Message "Filtered files: $($files.Count) included, $excludedCount excluded"
            }
        }
        catch {
            $errorMsg = "Failed to enumerate files in directory: $($_.Exception.Message)"
            Write-LogError -LogPath $LogPath -Message $errorMsg -Exception $_.Exception
            $errors.Add([PSCustomObject]@{
                Path = $path
                ErrorType = "FileEnumeration"
                ErrorMessage = $_.Exception.Message
                Timestamp = Get-Date
                ParentPath = Split-Path $path -Parent
            })
        }
    }
    catch {
        $errorMsg = "Failed to access directory path: $($_.Exception.Message)"
        Write-LogError -LogPath $LogPath -Message $errorMsg -Exception $_.Exception
        $errors.Add([PSCustomObject]@{
            Path = $path
            ErrorType = "PathAccess"
            ErrorMessage = $_.Exception.Message
            Timestamp = Get-Date
            ParentPath = Split-Path $path -Parent
        })
    }
    
    Write-LogInfo -LogPath $LogPath -Message "File enumeration completed: $($files.Count) files, $($errors.Count) errors"
    
    return [PSCustomObject]@{
        Files = $files
        AccessErrors = $errors
    }
}

function New-FileInfoObject {
    <#
    .SYNOPSIS
    Creates a standardized file information object for snapshots with optimized attribute processing.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo]$FileInfo
    )
    
    # Use script-scoped dictionary for better performance (PowerShell doesn't support static)
    if (-not $script:AttributeMap) {
        $script:AttributeMap = @{
            [System.IO.FileAttributes]::ReadOnly = "ReadOnly"
            [System.IO.FileAttributes]::Hidden = "Hidden"
            [System.IO.FileAttributes]::System = "System"
            [System.IO.FileAttributes]::Archive = "Archive"
            [System.IO.FileAttributes]::Compressed = "Compressed"
            [System.IO.FileAttributes]::Encrypted = "Encrypted"
            [System.IO.FileAttributes]::Normal = "Normal"
        }
    }
    
    # Pre-calculate all values to avoid repeated calculations
    $sizeBytes = $FileInfo.Length
    $sizeKB = [math]::Round($sizeBytes / 1KB, 2)
    $sizeMB = [math]::Round($sizeBytes / 1MB, 2)
    
    # Process attributes inline for better performance
    $attributeList = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $script:AttributeMap.Keys) {
        if ($FileInfo.Attributes -band $key) {
            $attributeList.Add($script:AttributeMap[$key])
        }
    }
    
    $attributesString = if ($attributeList.Count -eq 0) { "None" } else { ($attributeList -join ", ") }
    
    return [PSCustomObject]@{
        FullPath = $FileInfo.FullName
        Name = $FileInfo.Name
        Type = "File"
        SizeBytes = $sizeBytes
        SizeKB = $sizeKB
        SizeMB = $sizeMB
        LastWriteTime = $FileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
        LastWriteTimeUtc = $FileInfo.LastWriteTimeUtc.ToString("yyyy-MM-dd HH:mm:ss")
        CreationTime = $FileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
        LastAccessTime = $FileInfo.LastAccessTime.ToString("yyyy-MM-dd HH:mm:ss")
        Attributes = $attributesString
        AttributesRaw = $FileInfo.Attributes.ToString()
        Extension = $FileInfo.Extension
        Directory = $FileInfo.DirectoryName
    }
}

function Export-ToJson {
    <#
    .SYNOPSIS
    Universal function to export data to JSON with standardized metadata.
    
    .DESCRIPTION
    Unified export function that handles all JSON export operations including:
    - Files, access errors, or any other data type
    - Directory creation
    - Metadata generation
    - JSON conversion and file writing
    - File size reporting
    
    .PARAMETER Data
    The data to export (files list, access errors list, etc.)
    
    .PARAMETER OutputPath
    Path where the JSON file will be saved
    
    .PARAMETER NetworkSharePath
    Source network share path for metadata
    
    .PARAMETER DataType
    Type of data being exported (files, access_errors, snapshot, comparison, etc.)
    
    .PARAMETER AccessErrorsCount
    Number of access errors for metadata (optional, used for files export)
    
    .PARAMETER LogPath
    The path to the log file where operations will be logged using PSLogger.
    
    .OUTPUTS
    Boolean indicating success/failure
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        $Data,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$true)]
        [string]$NetworkSharePath,
        
        [Parameter(Mandatory=$true)]
        [ValidateSet("files", "access_errors", "snapshot", "comparison")]
        [string]$DataType,
        
        [Parameter(Mandatory=$false)]
        [int]$AccessErrorsCount = 0,
        
        [Parameter(Mandatory=$true)]
        [string]$LogPath
    )
    
    try {
        # Ensure Data is not null and convert to array if needed
        $dataArray = if ($null -eq $Data) {
            Write-LogDebug -LogPath $LogPath -Message "Data was null, using empty array"
            @()
        } elseif ($Data -is [System.Collections.Generic.List[object]]) {
            Write-LogDebug -LogPath $LogPath -Message "Converting List to array for JSON export"
            $Data.ToArray()
        } else {
            @($Data)
        }
        
        $itemCount = $dataArray.Count
        Write-Host "Exporting $itemCount $DataType to JSON..." -ForegroundColor Yellow
        Write-LogInfo -LogPath $LogPath -Message "Starting export of $itemCount $DataType to: $OutputPath"
        
        # Create output directory if it doesn't exist
        $outputDir = Split-Path $OutputPath -Parent
        if ($outputDir -and -not (Test-Path $outputDir -PathType Container)) {
            Write-LogDebug -LogPath $LogPath -Message "Creating output directory: $outputDir"
            [void](New-Item -ItemType Directory -Path $outputDir -Force)
        }
        
        # Create standardized metadata
        $now = Get-Date
        $metadata = @{
            SnapshotDate = $now.ToString("yyyy-MM-dd HH:mm:ss")
            SnapshotDateUtc = $now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
            NetworkSharePath = $NetworkSharePath
        }
        
        # Create export data structure based on type
        $exportData = switch ($DataType) {
            "files" {
                $metadata.TotalItems = $itemCount
                $metadata.AccessErrorsCount = $AccessErrorsCount
                @{
                    Metadata = $metadata
                    Files = $dataArray
                }
            }
            "access_errors" {
                $metadata.TotalAccessErrors = $itemCount
                @{
                    Metadata = $metadata
                    AccessErrors = $dataArray
                }
            }
            default {
                # Generic format for snapshot, comparison, or other data types
                $metadata."TotalItems" = $itemCount
                @{
                    Metadata = $metadata
                    Data = $dataArray
                }
            }
        }
        
        Write-LogDebug -LogPath $LogPath -Message "Converting data to JSON format"
        
        # Convert to JSON and save
        $jsonString = $exportData | ConvertTo-Json -Depth 10 -Compress
        [System.IO.File]::WriteAllText($OutputPath, $jsonString, [System.Text.Encoding]::UTF8)
        
        # Display file export information
        $fileSizeKB = [math]::Round((Get-Item $OutputPath).Length / 1KB, 2)
        Write-Host "$DataType exported to: $OutputPath ($fileSizeKB KB)" -ForegroundColor Green
        Write-LogInfo -LogPath $LogPath -Message "$DataType export completed successfully: $OutputPath ($fileSizeKB KB)"
        
        return $true
    }
    catch {
        $errorMsg = "Failed to export $DataType to JSON: $($_.Exception.Message)"
        Write-LogError -LogPath $LogPath -Message $errorMsg -Exception $_.Exception
        Write-Error $errorMsg
        return $false
    }
}

function Get-NetworkShareSnapshot {
    <#
    .SYNOPSIS
    Captures a snapshot of all files within a monitored directory with JSON export.
    
    .DESCRIPTION
    This function captures a comprehensive snapshot of files in a monitored directory,
    using the pre-compiled exclusion patterns from the configuration system.
    
    .PARAMETER MonitoredDirectory
    A monitored directory object from Read-SettingsFile containing path and compiledExclusionPatterns
    
    .PARAMETER JsonOutputPath
    Path where the JSON file will be saved.
    
    .PARAMETER AccessErrorsJsonPath
    Path where access errors will be exported to JSON.
    
    .PARAMETER LogPath
    The path to the log file where operations will be logged using PSLogger.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({
            if (-not $_.PSObject.Properties.Name -contains "path") {
                throw "MonitoredDirectory must have a 'path' property"
            }
            if (-not $_.PSObject.Properties.Name -contains "compiledExclusionPatterns") {
                throw "MonitoredDirectory must have a 'compiledExclusionPatterns' property. Use Read-SettingsFile to ensure proper initialization."
            }
            return $true
        })]
        $MonitoredDirectory,
        
        [Parameter(Mandatory=$true)]
        [string]$JsonOutputPath,
        
        [Parameter(Mandatory=$true)]
        [string]$AccessErrorsJsonPath,
        
        [Parameter(Mandatory=$true)]
        [string]$LogPath
    )
    
    $networkSharePath = $MonitoredDirectory.path
    $compiledExclusionPatterns = $MonitoredDirectory.compiledExclusionPatterns
    
    Write-LogInfo -LogPath $LogPath -Message "Starting network share snapshot for: $networkSharePath"
    
    # Validate path
    if (-not (Test-Path $networkSharePath)) {
        $errorMsg = "Network share path '$networkSharePath' does not exist or is not accessible"
        Write-LogError -LogPath $LogPath -Message $errorMsg
        throw $errorMsg
    }
    
    Write-LogInfo -LogPath $LogPath -Message "Path validation successful"
    Write-LogInfo -LogPath $LogPath -Message "Output files - Snapshot: $JsonOutputPath, Access Errors: $AccessErrorsJsonPath"
    
    Write-Host "Starting file snapshot of: $networkSharePath" -ForegroundColor Cyan
    if ($compiledExclusionPatterns.Count -gt 0) {
        Write-Host "Using $($compiledExclusionPatterns.Count) pre-compiled exclusion patterns" -ForegroundColor Yellow
        Write-LogInfo -LogPath $LogPath -Message "Using $($compiledExclusionPatterns.Count) pre-compiled exclusion patterns"
        
        # Optionally show the original exclusion names for debugging
        if ($MonitoredDirectory.PSObject.Properties.Name -contains "exclusions" -and $MonitoredDirectory.exclusions.Count -gt 0) {
            Write-Host "Exclusions: $($MonitoredDirectory.exclusions -join ', ')" -ForegroundColor Gray
            Write-LogDebug -LogPath $LogPath -Message "Exclusion patterns: $($MonitoredDirectory.exclusions -join ', ')"
        }
    }
    
    $startTime = Get-Date
    Write-LogInfo -LogPath $LogPath -Message "Snapshot operation started at: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    
    try {
        # Get all files with exclusions
        Write-Host "Scanning directory structure..." -ForegroundColor Yellow
        Write-LogInfo -LogPath $LogPath -Message "Beginning directory structure scan"
        
        $result = Get-FilesWithExclusions -MonitoredDirectory $MonitoredDirectory -LogPath $LogPath
        
        Write-Host "Processing $($result.Files.Count) files in batches..." -ForegroundColor Yellow
        Write-LogInfo -LogPath $LogPath -Message "File scan completed: $($result.Files.Count) files found"
        
        if ($result.AccessErrors.Count -gt 0) {
            Write-Host "Encountered $($result.AccessErrors.Count) access errors during scanning" -ForegroundColor Yellow
            Write-LogWarning -LogPath $LogPath -Message "Access errors encountered during scan: $($result.AccessErrors.Count)"
            
            # Log individual access errors for debugging
            foreach ($error in $result.AccessErrors) {
                Write-LogWarning -LogPath $LogPath -Message "Access error - Path: $($error.Path), Type: $($error.ErrorType), Error: $($error.ErrorMessage)"
            }
        }
        
        # Pre-allocate file list with known capacity
        $fileList = [System.Collections.Generic.List[object]]::new($result.Files.Count)
        
        # Process files in optimized batches for better memory management and performance
        $BatchSize = 50000
        $fileCount = $result.Files.Count
        
        Write-LogInfo -LogPath $LogPath -Message "Processing $fileCount files in batches of $BatchSize"
        
        for ($i = 0; $i -lt $fileCount; $i += $BatchSize) {
            $batchEnd = [Math]::Min($i + $BatchSize - 1, $fileCount - 1)
            $batch = $result.Files[$i..$batchEnd]
            $batchNumber = [Math]::Floor($i / $BatchSize) + 1
            $totalBatches = [Math]::Ceiling($fileCount / $BatchSize)
            
            Write-LogDebug -LogPath $LogPath -Message "Processing batch $batchNumber of $totalBatches ($($batch.Count) files)"
            
            # Process batch with minimal error handling overhead
            foreach ($file in $batch) {
                try {
                    $fileList.Add((New-FileInfoObject -FileInfo $file))
                }
                catch {
                    # Add processing errors to the access errors collection
                    Write-LogWarning -LogPath $LogPath -Message "Error processing file: $($file.FullName) - $($_.Exception.Message)" -Exception $_.Exception
                    $result.AccessErrors.Add([PSCustomObject]@{
                        Path = $file.FullName
                        ErrorType = "FileProcessing"
                        ErrorMessage = $_.Exception.Message
                        Timestamp = Get-Date
                        ParentPath = Split-Path $file.FullName -Parent
                    })
                }
            }
        }
        
        Write-LogInfo -LogPath $LogPath -Message "File processing completed: $($fileList.Count) files processed successfully"
        
        # Export results using unified export function
        Write-Host "Exporting results..." -ForegroundColor Yellow
        Write-LogInfo -LogPath $LogPath -Message "Beginning export operations"
        
        # Export access errors first (smaller file)
        Write-LogDebug -LogPath $LogPath -Message "Exporting access errors to: $AccessErrorsJsonPath"
        $accessErrorsExportResult = Export-ToJson -Data $result.AccessErrors -OutputPath $AccessErrorsJsonPath -NetworkSharePath $networkSharePath -DataType "access_errors" -LogPath $LogPath
        
        # Export main snapshot
        Write-LogDebug -LogPath $LogPath -Message "Exporting snapshot data to: $JsonOutputPath"
        $snapshotExportResult = Export-ToJson -Data $fileList -OutputPath $JsonOutputPath -NetworkSharePath $networkSharePath -DataType "files" -AccessErrorsCount $result.AccessErrors.Count -LogPath $LogPath
        
        # Check if exports were successful
        if (-not $accessErrorsExportResult -or -not $snapshotExportResult) {
            $errorMsg = "Failed to export snapshot data"
            Write-LogError -LogPath $LogPath -Message $errorMsg
            throw $errorMsg
        }
        
        Write-LogInfo -LogPath $LogPath -Message "Export operations completed successfully"
        
        # Show summary
        $endTime = Get-Date
        Show-SnapshotSummary -FileList $fileList -AccessErrorsCount $result.AccessErrors.Count -StartTime $startTime -EndTime $endTime -LogPath $LogPath
        
        Write-LogInfo -LogPath $LogPath -Message "Snapshot operation completed successfully at: $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    }
    catch {
        $errorMsg = "Snapshot failed: $($_.Exception.Message)"
        Write-LogError -LogPath $LogPath -Message $errorMsg -Exception $_.Exception
        Write-Error $errorMsg
        throw
    }
}

function Show-SnapshotSummary {
    <#
    .SYNOPSIS
    Displays a detailed summary of the network share snapshot results.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        $FileList,
        
        [Parameter(Mandatory=$true)]
        [int]$AccessErrorsCount,
        
        [Parameter(Mandatory=$true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory=$true)]
        [datetime]$EndTime,
        
        [Parameter(Mandatory=$true)]
        [string]$LogPath
    )
    
    # Convert to array and get count safely
    $fileArray = if ($FileList -is [System.Collections.Generic.List[object]]) {
        $FileList.ToArray()
    } else {
        @($FileList)
    }
    $fileCount = $fileArray.Count
    
    $duration = $EndTime - $StartTime
    
    Write-LogInfo -LogPath $LogPath -Message "Generating snapshot summary"
    Write-LogInfo -LogPath $LogPath -Message "Summary statistics - Files: $fileCount, Errors: $AccessErrorsCount, Duration: $([math]::Round($duration.TotalSeconds, 2)) seconds"
    
    # Display summary
    Write-Host "`nSnapshot completed successfully!" -ForegroundColor Green
    Write-Host "Total files processed: $fileCount" -ForegroundColor Cyan
    
    if ($AccessErrorsCount -gt 0) {
        Write-Host "Files skipped due to access errors: $AccessErrorsCount" -ForegroundColor Yellow
        Write-LogWarning -LogPath $LogPath -Message "Access errors resulted in $AccessErrorsCount files being skipped"
    }
    
    # Show detailed statistics
    Write-Host "`nDetailed Summary:" -ForegroundColor Cyan
    Write-Host "Files: $fileCount" -ForegroundColor White
    Write-Host "Total execution time: $([math]::Round($duration.TotalSeconds, 2)) seconds ($([math]::Round($duration.TotalMinutes, 2)) minutes)" -ForegroundColor Green
    
    if ($duration.TotalSeconds -gt 0) {
        $processingRate = [math]::Round($fileCount / $duration.TotalSeconds, 0)
        Write-Host "Processing rate: $processingRate files/second" -ForegroundColor Green
        Write-LogInfo -LogPath $LogPath -Message "Processing performance: $processingRate files/second"
    }
    
    if ($fileCount -gt 0) {
        Write-LogDebug -LogPath $LogPath -Message "Calculating file size statistics"
        
        # Optimize size calculations using LINQ-style operations
        $sizeStats = $fileArray | Measure-Object -Property SizeBytes -Sum -Average
        $totalSize = $sizeStats.Sum
        $averageSize = $sizeStats.Average
        
        Write-Host "Total size: $([math]::Round($totalSize / 1MB, 2)) MB ($([math]::Round($totalSize / 1GB, 2)) GB)" -ForegroundColor White
        Write-Host "Average file size: $([math]::Round($averageSize / 1KB, 2)) KB" -ForegroundColor White
        
        Write-LogInfo -LogPath $LogPath -Message "Size statistics - Total: $([math]::Round($totalSize / 1MB, 2)) MB, Average: $([math]::Round($averageSize / 1KB, 2)) KB"
        
        # Optimized file size distribution using single pass
        $sizeRanges = @{
            "0-1KB" = 0
            "1KB-1MB" = 0
            "1MB-100MB" = 0
            "100MB+" = 0
        }
        
        foreach ($file in $fileArray) {
            $size = $file.SizeBytes
            if ($size -le 1KB) { $sizeRanges["0-1KB"]++ }
            elseif ($size -le 1MB) { $sizeRanges["1KB-1MB"]++ }
            elseif ($size -le 100MB) { $sizeRanges["1MB-100MB"]++ }
            else { $sizeRanges["100MB+"]++ }
        }
        
        Write-Host "`nFile Size Distribution:" -ForegroundColor Cyan
        foreach ($range in $sizeRanges.GetEnumerator()) {
            $percentage = [math]::Round(($range.Value / $fileCount) * 100, 1)
            Write-Host "  $($range.Key): $($range.Value) files ($percentage%)" -ForegroundColor White
            Write-LogDebug -LogPath $LogPath -Message "Size distribution - $($range.Key): $($range.Value) files ($percentage%)"
        }
    }
    
    Write-LogInfo -LogPath $LogPath -Message "Snapshot summary completed"
}

function Compare-NetworkShareSnapshots {
    <#
    .SYNOPSIS
    Compares two network share snapshots to identify added, modified, or deleted files.
    
    .DESCRIPTION
    This function compares two snapshots created by Get-NetworkShareSnapshot and identifies:
    - New files (present in snapshot 2 but not in snapshot 1)
    - Deleted files (present in snapshot 1 but not in snapshot 2)
    - Modified files (present in both but with different size or last write time)
    - Files that were inaccessible in either snapshot are excluded from the comparison
    - Unchanged files are automatically excluded from the results
    - Results are always exported to JSON at the specified OutputPath
    
    .PARAMETER Snapshot1Path
    Path to the first snapshot JSON file
    
    .PARAMETER Snapshot2Path
    Path to the second snapshot JSON file
    
    .PARAMETER OutputPath
    Path where the comparison results will be saved as JSON (mandatory)
    
    .PARAMETER Snapshot1AccessErrorsPath
    Optional. Path to the first snapshot's access errors JSON file
    
    .PARAMETER Snapshot2AccessErrorsPath
    Optional. Path to the second snapshot's access errors JSON file
    
    .PARAMETER LogPath
    The path to the log file where operations will be logged using PSLogger.
        
    .EXAMPLE
    Compare-NetworkShareSnapshots -Snapshot1Path "C:\snap1.json" -Snapshot2Path "C:\snap2.json" -OutputPath "C:\comparison_results.json" -LogPath "app.log"
    
    .EXAMPLE
    Compare-NetworkShareSnapshots -Snapshot1Path "C:\snap1.json" -Snapshot2Path "C:\snap2.json" -OutputPath "C:\comparison.json" -Snapshot1AccessErrorsPath "C:\snap1_AccessErrors.json" -Snapshot2AccessErrorsPath "C:\snap2_AccessErrors.json" -LogPath "app.log"
    
    .OUTPUTS
    None. Results are always exported to the specified JSON file.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Snapshot1Path,
        
        [Parameter(Mandatory=$true)]
        [string]$Snapshot2Path,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$false)]
        [string]$Snapshot1AccessErrorsPath,
        
        [Parameter(Mandatory=$false)]
        [string]$Snapshot2AccessErrorsPath,
        
        [Parameter(Mandatory=$true)]
        [string]$LogPath
    )
    
    $startTime = Get-Date
    Write-Host "Starting snapshot comparison..."
    Write-LogInfo -LogPath $LogPath -Message "Starting snapshot comparison: $Snapshot1Path vs $Snapshot2Path -> $OutputPath"
    
    # Validate and load snapshots
    $snapshot1Data, $snapshot2Data = @($Snapshot1Path, $Snapshot2Path) | ForEach-Object {
        if (-not (Test-Path $_)) {
            throw "Snapshot file not found: $_"
        }
        [System.IO.File]::ReadAllText($_) | ConvertFrom-Json
    }
    
    $snapshot1 = $snapshot1Data.Files
    $snapshot2 = $snapshot2Data.Files
    
    Write-Host "Snapshot 1: $($snapshot1.Count) items, Snapshot 2: $($snapshot2.Count) items"
    Write-LogInfo -LogPath $LogPath -Message "Loaded snapshots: $($snapshot1.Count) and $($snapshot2.Count) files"
    
    # Load access errors and filter snapshots
    $exclusionPaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    
    @($Snapshot1AccessErrorsPath, $Snapshot2AccessErrorsPath) | Where-Object { $_ -and (Test-Path $_) } | ForEach-Object {
        try {
            $errorsData = [System.IO.File]::ReadAllText($_) | ConvertFrom-Json
            $errorsData.AccessErrors | ForEach-Object { [void]$exclusionPaths.Add($_.Path) }
        }
        catch {
            Write-LogWarning -LogPath $LogPath -Message "Failed to load access errors from $_" -Exception $_.Exception
        }
    }
    
    if ($exclusionPaths.Count -gt 0) {
        Write-Host "Excluding files under $($exclusionPaths.Count) inaccessible paths..."
        $filter = { 
            $filePath = $_.FullPath
            -not ($exclusionPaths | Where-Object { $filePath.StartsWith($_, [StringComparison]::OrdinalIgnoreCase) })
        }
        $snapshot1 = $snapshot1 | Where-Object $filter
        $snapshot2 = $snapshot2 | Where-Object $filter
        Write-LogInfo -LogPath $LogPath -Message "After filtering: $($snapshot1.Count) and $($snapshot2.Count) files"
    }
    
    # Create lookup hashtables
    $snapshot1Hash = @{}
    $snapshot2Hash = @{}
    $snapshot1 | ForEach-Object { $snapshot1Hash[$_.FullPath] = $_ }
    $snapshot2 | ForEach-Object { $snapshot2Hash[$_.FullPath] = $_ }
    
    Write-Host "Analyzing changes..."
    $comparisonResults = [System.Collections.Generic.List[object]]::new()
    $counts = @{ Added = 0; Modified = 0; Deleted = 0 }
    $totalSizeChange = 0
    
    # Process all files in snapshot2 (new and modified)
    foreach ($item2 in $snapshot2) {
        $path = $item2.FullPath
        $item1 = $snapshot1Hash[$path]
        
        if ($item1) {
            # Check for modifications
            $modifications = @()
            if ($item1.SizeBytes -ne $item2.SizeBytes) {
                $modifications += "Size changed from $($item1.SizeBytes) to $($item2.SizeBytes) bytes"
            }
            
            $time1 = [DateTime]::Parse($item1.LastWriteTime)
            $time2 = [DateTime]::Parse($item2.LastWriteTime)
            if ([Math]::Abs(($time2 - $time1).TotalSeconds) -gt 1) {
                $modifications += "Last write time changed from $($item1.LastWriteTime) to $($item2.LastWriteTime)"
            }
            
            if ($item1.Attributes -ne $item2.Attributes) {
                $modifications += "Attributes changed from '$($item1.Attributes)' to '$($item2.Attributes)'"
            }
            
            # Only add if modified
            if ($modifications.Count -gt 0) {
                $sizeDiff = [int64]$item2.SizeBytes - [int64]$item1.SizeBytes
                $totalSizeChange += $sizeDiff
                $counts.Modified++
                
                $comparisonResults.Add([PSCustomObject]@{
                    FullPath = $path
                    Name = $item2.Name
                    Status = "Modified"
                    Type = $item2.Type
                    OldSizeBytes = $item1.SizeBytes
                    NewSizeBytes = $item2.SizeBytes
                    SizeDifferenceBytes = $sizeDiff
                    OldLastWriteTime = $item1.LastWriteTime
                    NewLastWriteTime = $item2.LastWriteTime
                    OldAttributes = $item1.Attributes
                    NewAttributes = $item2.Attributes
                    ModificationReason = $modifications -join "; "
                    Directory = $item2.Directory
                })
            }
        } else {
            # New file
            $totalSizeChange += $item2.SizeBytes
            $counts.Added++
            
            $comparisonResults.Add([PSCustomObject]@{
                FullPath = $path
                Name = $item2.Name
                Status = "Added"
                Type = $item2.Type
                OldSizeBytes = $null
                NewSizeBytes = $item2.SizeBytes
                SizeDifferenceBytes = $item2.SizeBytes
                OldLastWriteTime = $null
                NewLastWriteTime = $item2.LastWriteTime
                OldAttributes = $null
                NewAttributes = $item2.Attributes
                ModificationReason = "File added"
                Directory = $item2.Directory
            })
        }
    }
    
    # Process deleted files (in snapshot1 but not in snapshot2)
    foreach ($item1 in $snapshot1) {
        if (-not $snapshot2Hash.ContainsKey($item1.FullPath)) {
            $sizeDiff = -[int64]$item1.SizeBytes
            $totalSizeChange += $sizeDiff
            $counts.Deleted++
            
            $comparisonResults.Add([PSCustomObject]@{
                FullPath = $item1.FullPath
                Name = $item1.Name
                Status = "Deleted"
                Type = $item1.Type
                OldSizeBytes = $item1.SizeBytes
                NewSizeBytes = $null
                SizeDifferenceBytes = $sizeDiff
                OldLastWriteTime = $item1.LastWriteTime
                NewLastWriteTime = $null
                OldAttributes = $item1.Attributes
                NewAttributes = $null
                ModificationReason = "File deleted"
                Directory = $item1.Directory
            })
        }
    }
    
    Write-LogInfo -LogPath $LogPath -Message "Analysis complete: $($counts.Added) added, $($counts.Modified) modified, $($counts.Deleted) deleted"
    
    # Export results
    Write-Host "Exporting comparison results..."
    $outputDir = Split-Path $OutputPath -Parent
    if ($outputDir -and -not (Test-Path $outputDir)) {
        [void](New-Item -ItemType Directory -Path $outputDir -Force)
    }
    
    $now = Get-Date
    $comparisonData = @{
        Metadata = @{
            ComparisonDate = $now.ToString("yyyy-MM-dd HH:mm:ss")
            ComparisonDateUtc = $now.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
            Snapshot1Path = $Snapshot1Path
            Snapshot2Path = $Snapshot2Path
            Snapshot1Date = $snapshot1Data.Metadata.SnapshotDate
            Snapshot2Date = $snapshot2Data.Metadata.SnapshotDate
            TotalComparisons = $comparisonResults.Count
            ExcludeUnchanged = $true
        }
        Results = $comparisonResults
    }
    
    $jsonString = $comparisonData | ConvertTo-Json -Depth 10 -Compress
    [System.IO.File]::WriteAllText($OutputPath, $jsonString, [System.Text.Encoding]::UTF8)
    
    # Display summary
    $duration = (Get-Date) - $startTime
    $fileSizeKB = [math]::Round((Get-Item $OutputPath).Length / 1KB, 2)
    
    Write-Host "`nComparison completed successfully!" -ForegroundColor Green
    Write-Host "Added: $($counts.Added), Modified: $($counts.Modified), Deleted: $($counts.Deleted)" -ForegroundColor Cyan
    Write-Host "Net size change: $([math]::Round($totalSizeChange / 1MB, 2)) MB" -ForegroundColor Cyan
    Write-Host "Results exported to: $OutputPath ($fileSizeKB KB)" -ForegroundColor Green
    Write-Host "Execution time: $([math]::Round($duration.TotalSeconds, 2)) seconds" -ForegroundColor Green
    
    Write-LogInfo -LogPath $LogPath -Message "Comparison completed in $([math]::Round($duration.TotalSeconds, 2))s - Results: $($comparisonResults.Count) changes, Size: $fileSizeKB KB"
}

function Copy-SnapshotChangesToStaging {
    <#
    .SYNOPSIS
        Copies modified and added files from a snapshot comparison to the staging area using BITS.
    
    .DESCRIPTION
        This cmdlet reads the comparison results from Compare-NetworkShareSnapshots and copies
        only the modified and added files to the staging area using Background Intelligent Transfer Service (BITS).
        The original directory structure is preserved within the staging area, organized by drive letter.
        Files are always overwritten in the staging area if they already exist.
    
    .PARAMETER ComparisonResultsPath
        Path to the JSON file containing comparison results from Compare-NetworkShareSnapshots.
    
    .PARAMETER StagingArea
        The staging area path where files will be copied.
    
    .PARAMETER LogPath
        The path to the log file where operations will be logged using PSLogger.
    
    .PARAMETER DisplayName
        Display name for the BITS transfer job. Defaults to "Snapshot Changes Copy".
    
    .PARAMETER Priority
        BITS transfer priority. Valid values: Foreground, High, Normal, Low. Default is Normal.
    
    .PARAMETER RetryInterval
        Retry interval in seconds for BITS transfer. Default is 60 seconds.
    
    .PARAMETER RetryTimeout
        Timeout in seconds for BITS transfer completion. Default is 300 seconds (5 minutes).
    
    .PARAMETER MaxConcurrentJobs
        Maximum number of concurrent BITS jobs to run. Default is 3.
        
    .EXAMPLE
        Copy-SnapshotChangesToStaging -ComparisonResultsPath "C:\Reports\comparison.json" -StagingArea "C:\StagingArea" -LogPath "C:\Logs\copy.log"
        
        Copies all modified and added files from the comparison results to the specified staging area.
    
    .OUTPUTS
        PSCustomObject containing copy operation results including success count, failure count, and detailed results.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "Comparison results file not found: $_"
            }
            return $true
        })]
        [string]$ComparisonResultsPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$StagingArea,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [string]$DisplayName = "Snapshot Changes Copy",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Foreground", "High", "Normal", "Low")]
        [string]$Priority = "Normal",
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(10, 3600)]
        [int]$RetryInterval = 60,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(60, 7200)]
        [int]$RetryTimeout = 300,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 10)]
        [int]$MaxConcurrentJobs = 3
    )
    
    $startTime = Get-Date
    Write-LogInfo -LogPath $LogPath -Message "Starting copy operation: $ComparisonResultsPath -> $StagingArea"
    
    # Load and filter comparison results to get files to copy
    try {
        $comparisonData = [System.IO.File]::ReadAllText($ComparisonResultsPath) | ConvertFrom-Json
        $filesToCopy = $comparisonData.Results | Where-Object { $_.Status -in @("Added", "Modified") }
        Write-LogInfo -LogPath $LogPath -Message "Found $($filesToCopy.Count) files to copy"
    }
    catch {
        Write-LogError -LogPath $LogPath -Message "Failed to load comparison results" -Exception $_.Exception
        throw
    }
    
    # Early exit if no files to copy
    if ($filesToCopy.Count -eq 0) {
        Write-Host "No files to copy." -ForegroundColor Yellow
        return New-OperationResult -OperationType "Copy" -TotalFiles 0 -Duration ((Get-Date) - $startTime)
    }
    
    # Prepare copy operations - transform comparison results to operation objects
    $copyOps = @()
    $skipped = 0
    $totalSize = 0
    
    foreach ($file in $filesToCopy) {
        if (-not (Test-Path $file.FullPath -PathType Leaf)) {
            Write-LogWarning -LogPath $LogPath -Message "Source not found: $($file.FullPath)"
            $skipped++
            continue
        }
        
        # Build destination path: StagingArea\DriveLetter\PathAfterDrive
        $driveLetter = $file.FullPath[0]
        $pathAfterDrive = $file.FullPath.Substring(2).TrimStart('\')
        $destPath = Join-Path $StagingArea $driveLetter $pathAfterDrive
        
        try {
            $fileSize = (Get-Item $file.FullPath).Length
            $totalSize += $fileSize
        }
        catch {
            $fileSize = 0
            Write-LogWarning -LogPath $LogPath -Message "Could not get size for: $($file.FullPath)"
        }
        
        $copyOps += [PSCustomObject]@{
            Source = $file.FullPath
            Destination = $destPath
            Status = $file.Status
            Size = $fileSize
            Name = $file.Name
        }
    }
    
    if ($copyOps.Count -eq 0) {
        Write-Host "No valid files to copy." -ForegroundColor Yellow
        return New-OperationResult -OperationType "Copy" -TotalFiles $filesToCopy.Count -SkippedFiles $skipped -Duration ((Get-Date) - $startTime)
    }
    
    # Execute the file operations using shared logic
    $params = @{
        Operations = $copyOps
        OperationType = "Copy"
        LogPath = $LogPath
        DisplayName = $DisplayName
        Priority = $Priority
        RetryInterval = $RetryInterval
        RetryTimeout = $RetryTimeout
        MaxConcurrentJobs = $MaxConcurrentJobs
        TotalInputFiles = $filesToCopy.Count
        TotalSize = $totalSize
        SkippedFiles = $skipped
        StartTime = $startTime
        ShouldProcess = { $true }
    }
    
    return Invoke-FileOperations @params
}

function Move-StagingFilesToNetwork {
    <#
    .SYNOPSIS
        Moves files from the staging area back to their original network locations using BITS.
    
    .DESCRIPTION
        This cmdlet moves files from the staging area back to their original network share locations.
        It reconstructs the original file paths by taking the drive letter from the staging folder structure,
        appending :\ and the remainder of the path. Files are moved (not copied) using Background Intelligent 
        Transfer Service (BITS) for reliability and progress tracking.
    
    .PARAMETER StagingArea
        The staging area path where files are currently stored.
    
    .PARAMETER SuccessLogPath
        Path to log file for successfully moved files.
    
    .PARAMETER FailureLogPath
        Path to log file for files that failed to move (file in use, access denied, etc.).
    
    .PARAMETER LogPath
        The path to the main log file where operations will be logged using PSLogger.
    
    .PARAMETER DisplayName
        Display name for the BITS transfer job. Defaults to "Staging to Network Move".
    
    .PARAMETER Priority
        BITS transfer priority. Valid values: Foreground, High, Normal, Low. Default is Normal.
    
    .PARAMETER RetryInterval
        Retry interval in seconds for BITS transfer. Default is 60 seconds.
    
    .PARAMETER RetryTimeout
        Timeout in seconds for BITS transfer completion. Default is 300 seconds (5 minutes).
    
    .PARAMETER MaxConcurrentJobs
        Maximum number of concurrent BITS jobs to run. Default is 3.
        
    .EXAMPLE
        Move-StagingFilesToNetwork -StagingArea "C:\StagingArea" -SuccessLogPath "C:\Logs\moved_success.log" -FailureLogPath "C:\Logs\moved_failures.log" -LogPath "C:\Logs\move.log"
        
        Moves all files from the staging area back to their original network locations.
    
    .OUTPUTS
        PSCustomObject containing move operation results including success count, failure count, and detailed results.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "Staging area not found: $_"
            }
            return $true
        })]
        [string]$StagingArea,
        
        [Parameter(Mandatory = $true)]
        [string]$SuccessLogPath,
        
        [Parameter(Mandatory = $true)]
        [string]$FailureLogPath,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [string]$DisplayName = "Staging to Network Move",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("Foreground", "High", "Normal", "Low")]
        [string]$Priority = "Normal",
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(10, 3600)]
        [int]$RetryInterval = 60,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(60, 7200)]
        [int]$RetryTimeout = 300,
        
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 10)]
        [int]$MaxConcurrentJobs = 3
    )
    
    $startTime = Get-Date
    Write-LogInfo -LogPath $LogPath -Message "Starting move operation from staging: $StagingArea"
    
    # Ensure log directories exist
    foreach ($logFile in @($SuccessLogPath, $FailureLogPath)) {
        $logDir = Split-Path $logFile -Parent
        if ($logDir -and -not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
    }
    
    # Discover all files in staging area
    try {
        Write-Host "Scanning staging area for files..." -ForegroundColor Yellow
        Write-LogInfo -LogPath $LogPath -Message "Discovering files in staging area"
        
        $stagingFiles = Get-ChildItem -Path $StagingArea -File -Recurse -ErrorAction SilentlyContinue
        Write-LogInfo -LogPath $LogPath -Message "Found $($stagingFiles.Count) files in staging area"
    }
    catch {
        Write-LogError -LogPath $LogPath -Message "Failed to scan staging area" -Exception $_.Exception
        throw
    }
    
    # Early exit if no files found
    if ($stagingFiles.Count -eq 0) {
        Write-Host "No files found in staging area." -ForegroundColor Yellow
        Write-LogInfo -LogPath $LogPath -Message "No files found in staging area"
        return New-OperationResult -OperationType "Move" -TotalFiles 0 -Duration ((Get-Date) - $startTime)
    }
    
    # Prepare move operations - transform staging files to operation objects
    Write-Host "Preparing move operations..." -ForegroundColor Yellow
    $moveOps = @()
    $skipped = 0
    $totalSize = 0
    
    foreach ($file in $stagingFiles) {
        try {
            # Reconstruct original path: StagingArea\DriveLetter\PathAfterDrive -> DriveLetter:\PathAfterDrive
            $relativePath = $file.FullName.Substring($StagingArea.Length).TrimStart('\')
            
            # Extract drive letter (first component of relative path)
            $pathComponents = $relativePath.Split('\')
            if ($pathComponents.Length -lt 2) {
                Write-LogWarning -LogPath $LogPath -Message "Invalid staging path structure: $($file.FullName)"
                $skipped++
                continue
            }
            
            $driveLetter = $pathComponents[0]
            $pathAfterDrive = ($pathComponents[1..($pathComponents.Length - 1)]) -join '\'
            $originalPath = "$driveLetter`:\$pathAfterDrive"
            
            # Validate drive letter
            if ($driveLetter.Length -ne 1 -or -not [char]::IsLetter($driveLetter)) {
                Write-LogWarning -LogPath $LogPath -Message "Invalid drive letter '$driveLetter' in path: $($file.FullName)"
                $skipped++
                continue
            }
            
            # Get file size
            try {
                $fileSize = $file.Length
                $totalSize += $fileSize
            }
            catch {
                $fileSize = 0
                Write-LogWarning -LogPath $LogPath -Message "Could not get size for: $($file.FullName)"
            }
            
            $moveOps += [PSCustomObject]@{
                Source = $file.FullName
                Destination = $originalPath
                DriveLetter = $driveLetter
                Size = $fileSize
                Name = $file.Name
                RelativePath = $relativePath
            }
        }
        catch {
            Write-LogError -LogPath $LogPath -Message "Error preparing move for file: $($file.FullName)" -Exception $_.Exception
            $skipped++
        }
    }
    
    if ($moveOps.Count -eq 0) {
        Write-Host "No valid files to move." -ForegroundColor Yellow
        return New-OperationResult -OperationType "Move" -TotalFiles $stagingFiles.Count -SkippedFiles $skipped -Duration ((Get-Date) - $startTime)
    }
    
    # Filter out operations targeting inaccessible drives
    $filteredOps, $inaccessibleCount = Filter-OperationsByDriveAccessibility -Operations $moveOps -LogPath $LogPath
    
    if ($filteredOps.Count -eq 0) {
        Write-Host "No destination drives are accessible." -ForegroundColor Red
        Write-LogError -LogPath $LogPath -Message "No destination drives are accessible"
        return New-OperationResult -OperationType "Move" -TotalFiles $stagingFiles.Count -SkippedFiles ($skipped + $inaccessibleCount) -Duration ((Get-Date) - $startTime)
    }
    
    # Execute the file operations using shared logic
    $params = @{
        Operations = $filteredOps
        OperationType = "Move"
        LogPath = $LogPath
        DisplayName = $DisplayName
        Priority = $Priority
        RetryInterval = $RetryInterval
        RetryTimeout = $RetryTimeout
        MaxConcurrentJobs = $MaxConcurrentJobs
        TotalInputFiles = $stagingFiles.Count
        TotalSize = $totalSize
        SkippedFiles = ($skipped + $inaccessibleCount)
        StartTime = $startTime
        SuccessLogPath = $SuccessLogPath
        FailureLogPath = $FailureLogPath
        ShouldProcess = $PSCmdlet.ShouldProcess
    }
    
    return Invoke-FileOperations @params
}

function Invoke-FileOperations {
    <#
    .SYNOPSIS
        Common function to handle both copy and move operations.
    
    .DESCRIPTION
        This function consolidates the shared logic between copy and move operations,
        including directory creation, BITS transfers, and result reporting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Operations,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("Copy", "Move")]
        [string]$OperationType,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $true)]
        [string]$Priority,
        
        [Parameter(Mandatory = $true)]
        [int]$RetryInterval,
        
        [Parameter(Mandatory = $true)]
        [int]$RetryTimeout,
        
        [Parameter(Mandatory = $true)]
        [int]$MaxConcurrentJobs,
        
        [Parameter(Mandatory = $true)]
        [int]$TotalInputFiles,
        
        [Parameter(Mandatory = $true)]
        [long]$TotalSize,
        
        [Parameter(Mandatory = $true)]
        [int]$SkippedFiles,
        
        [Parameter(Mandatory = $true)]
        [datetime]$StartTime,
        
        [Parameter(Mandatory = $false)]
        [string]$SuccessLogPath,
        
        [Parameter(Mandatory = $false)]
        [string]$FailureLogPath,
        
        [Parameter(Mandatory = $true)]
        [scriptblock]$ShouldProcess
    )
    
    Write-Host "Prepared $($Operations.Count) files for $($OperationType.ToLower()) ($(Format-FileSize $TotalSize))" -ForegroundColor Green
    Write-LogInfo -LogPath $LogPath -Message "Prepared $($Operations.Count) files, skipped $SkippedFiles"
    
    # Create destination directories
    $uniqueDirs = $Operations | ForEach-Object { Split-Path $_.Destination -Parent } | Sort-Object -Unique
    $createdDirs = 0
    
    foreach ($dir in $uniqueDirs) {
        if (-not (Test-Path $dir)) {
            try {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
                $createdDirs++
            }
            catch {
                Write-LogError -LogPath $LogPath -Message "Failed to create directory: $dir" -Exception $_.Exception
                if ($OperationType -eq "Copy") {
                    throw  # Copy operations should fail if directories can't be created
                }
                # Move operations continue with other directories
            }
        }
    }
    
    Write-Host "Created $createdDirs directories" -ForegroundColor Green
    Write-LogInfo -LogPath $LogPath -Message "Created $createdDirs directories"
    
    # Execute BITS transfers
    Write-Host "Starting BITS $($OperationType.ToLower()) operations..." -ForegroundColor Cyan
    $bitsParams = @{
        Operations = $Operations
        LogPath = $LogPath
        DisplayName = $DisplayName
        Priority = $Priority
        RetryInterval = $RetryInterval
        RetryTimeout = $RetryTimeout
        MaxConcurrentJobs = $MaxConcurrentJobs
        OperationType = $OperationType
        ShouldProcess = $ShouldProcess
    }
    
    # Add move-specific parameters if needed
    if ($OperationType -eq "Move") {
        $bitsParams.SuccessLogPath = $SuccessLogPath
        $bitsParams.FailureLogPath = $FailureLogPath
    }
    
    $results = Start-BitsOperations @bitsParams
    
    # Final summary and result
    $duration = (Get-Date) - $StartTime
    Show-OperationSummary -Results $results -Duration $duration -Skipped $SkippedFiles -LogPath $LogPath -OperationType $OperationType
    
    return New-OperationResult -OperationType $OperationType -TotalFiles $TotalInputFiles -SuccessfulOperations $results.Success -FailedOperations $results.Failed -SkippedFiles $SkippedFiles -TotalSizeBytes $TotalSize -Duration $duration -Results $results.Details
}

function Filter-OperationsByDriveAccessibility {
    <#
    .SYNOPSIS
        Filters move operations to only include those targeting accessible drives.
    
    .DESCRIPTION
        This function is specific to move operations where we need to validate
        that destination drives are accessible before attempting the move.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Operations,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath
    )
    
    Write-Host "Validating destination drives..." -ForegroundColor Yellow
    $driveGroups = $Operations | Group-Object DriveLetter
    $accessibleDrives = @()
    
    foreach ($group in $driveGroups) {
        $drive = "$($group.Name):\"
        if (Test-Path $drive) {
            $accessibleDrives += $group.Name
            Write-LogInfo -LogPath $LogPath -Message "Drive $drive is accessible ($($group.Count) files)"
        }
        else {
            Write-Host "Warning: Drive $drive is not accessible ($($group.Count) files will be skipped)" -ForegroundColor Yellow
            Write-LogWarning -LogPath $LogPath -Message "Drive $drive is not accessible, $($group.Count) files will be skipped"
        }
    }
    
    # Filter to only accessible drives
    $accessibleOps = $Operations | Where-Object { $_.DriveLetter -in $accessibleDrives }
    $inaccessibleCount = $Operations.Count - $accessibleOps.Count
    
    return $accessibleOps, $inaccessibleCount
}

function Start-BitsOperations {
    <#
    .SYNOPSIS
        Generic BITS transfer manager that handles both copy and move operations.
    
    .DESCRIPTION
        This function manages BITS transfers with configurable behavior for copy vs move operations.
        Handles job queuing, progress monitoring, error handling, and cleanup.
    
    .PARAMETER Operations
        Array of operation objects containing Source, Destination, Name, and Size properties.
    
    .PARAMETER LogPath
        Path to the main log file for operation logging.
    
    .PARAMETER DisplayName
        Display name prefix for BITS transfer jobs.
    
    .PARAMETER Priority
        BITS transfer priority (Foreground, High, Normal, Low).
    
    .PARAMETER RetryInterval
        Retry interval in seconds for BITS transfers.
    
    .PARAMETER RetryTimeout
        Timeout in seconds for BITS transfer completion.
    
    .PARAMETER MaxConcurrentJobs
        Maximum number of concurrent BITS jobs to run.
    
    .PARAMETER OperationType
        Type of operation: "Copy" or "Move".
    
    .PARAMETER SuccessLogPath
        Optional. Path to log successful operations (used for move operations).
    
    .PARAMETER FailureLogPath
        Optional. Path to log failed operations (used for move operations).
    
    .PARAMETER ShouldProcess
        WhatIf/Confirm support scriptblock.
    
    .OUTPUTS
        Hashtable with Success count, Failed count, and Details array.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Operations,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("Foreground", "High", "Normal", "Low")]
        [string]$Priority,
        
        [Parameter(Mandatory = $true)]
        [int]$RetryInterval,
        
        [Parameter(Mandatory = $true)]
        [int]$RetryTimeout,
        
        [Parameter(Mandatory = $true)]
        [int]$MaxConcurrentJobs,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("Copy", "Move")]
        [string]$OperationType,
        
        [Parameter(Mandatory = $false)]
        [string]$SuccessLogPath,
        
        [Parameter(Mandatory = $false)]
        [string]$FailureLogPath,
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$ShouldProcess
    )
    
    $activeJobs = @()
    $completed = @{ Success = 0; Failed = 0; Details = @() }
    $opIndex = 0
    $isMove = ($OperationType -eq "Move")
    
    Write-LogInfo -LogPath $LogPath -Message "Starting $OperationType operations with $($Operations.Count) items"
    
    while ($opIndex -lt $Operations.Count -or $activeJobs.Count -gt 0) {
        # Start new jobs if capacity available
        while ($activeJobs.Count -lt $MaxConcurrentJobs -and $opIndex -lt $Operations.Count) {
            $op = $Operations[$opIndex]
            $jobName = "$DisplayName - $($op.Name)"
            
            try {
                # Handle existing destination file
                if (Test-Path $op.Destination) {
                    if ($isMove) {
                        Write-LogInfo -LogPath $LogPath -Message "Destination exists, will overwrite: $($op.Destination)"
                    } else {
                        Remove-Item $op.Destination -Force
                    }
                }
                
                # Start BITS transfer
                $job = Start-BitsTransfer -Source $op.Source -Destination $op.Destination -DisplayName $jobName -Priority $Priority -RetryInterval $RetryInterval -RetryTimeout $RetryTimeout -Asynchronous
                $activeJobs += @{ Job = $job; Operation = $op; StartTime = Get-Date }
                Write-LogDebug -LogPath $LogPath -Message "Started BITS $OperationType job for: $($op.Name)"
            }
            catch {
                $errorMsg = $_.Exception.Message
                $completed.Failed++
                
                $completed.Details += New-OperationResultDetail -Operation $op -Result "Failed" -Error $errorMsg -OperationType $OperationType
                
                if ($isMove) {
                    Write-LogError -LogPath $FailureLogPath -Message "FAILED TO START: $($op.Source) -> $($op.Destination) | Error: $errorMsg"
                }
                
                Write-LogError -LogPath $LogPath -Message "Failed to start BITS $OperationType job for: $($op.Source)" -Exception $_.Exception
            }
            $opIndex++
        }
        
        # Check active jobs
        if ($activeJobs.Count -gt 0) {
            Start-Sleep -Seconds 2
            $newActiveJobs = @()
            
            foreach ($activeJob in $activeJobs) {
                try {
                    $job = Get-BitsTransfer -JobId $activeJob.Job.JobId -ErrorAction SilentlyContinue
                    
                    if (-not $job) {
                        Write-LogWarning -LogPath $LogPath -Message "BITS job disappeared: $($activeJob.Operation.Name)"
                        continue
                    }
                    
                    switch ($job.JobState) {
                        "Transferred" {
                            Complete-BitsTransfer -BitsJob $job
                            
                            if ($isMove) {
                                # Handle move operation (copy + delete source)
                                try {
                                    Remove-Item -Path $activeJob.Operation.Source -Force -ErrorAction Stop
                                    
                                    # Success - log to success file using PSLogger
                                    Write-LogInfo -LogPath $SuccessLogPath -Message "MOVED: $($activeJob.Operation.Source) -> $($activeJob.Operation.Destination)"
                                    
                                    $completed.Success++
                                    $completed.Details += New-OperationResultDetail -Operation $activeJob.Operation -Result "Success" -OperationType "Move"
                                    Write-Host "✓ $($activeJob.Operation.Name)" -ForegroundColor Green
                                    Write-LogInfo -LogPath $LogPath -Message "Successfully moved: $($activeJob.Operation.Source) -> $($activeJob.Operation.Destination)"
                                }
                                catch {
                                    # Copy succeeded but delete failed - this is a partial failure
                                    $errorMsg = "Copy succeeded but failed to delete source: $($_.Exception.Message)"
                                    
                                    Write-LogError -LogPath $FailureLogPath -Message "PARTIAL FAILURE: $($activeJob.Operation.Source) -> $($activeJob.Operation.Destination) | Error: $errorMsg"
                                    
                                    $completed.Failed++
                                    $completed.Details += New-OperationResultDetail -Operation $activeJob.Operation -Result "Partial Failure" -Error $errorMsg -OperationType "Move"
                                    Write-Host "⚠ $($activeJob.Operation.Name) - Copy succeeded, delete failed" -ForegroundColor Yellow
                                    Write-LogWarning -LogPath $LogPath -Message "Partial failure for: $($activeJob.Operation.Source) - $errorMsg"
                                }
                            } else {
                                # Handle copy operation
                                $completed.Success++
                                $completed.Details += New-OperationResultDetail -Operation $activeJob.Operation -Result "Success" -OperationType "Copy"
                                Write-Host "✓ $($activeJob.Operation.Name)" -ForegroundColor Green
                                Write-LogInfo -LogPath $LogPath -Message "Successfully copied: $($activeJob.Operation.Source)"
                            }
                        }
                        { $_ -in @("Error", "TransientError", "Fatal") } {
                            $errorMsg = $job.ErrorDescription -or "Unknown BITS error"
                            Remove-BitsTransfer -BitsJob $job -ErrorAction SilentlyContinue
                            
                            if ($isMove) {
                                Write-LogError -LogPath $FailureLogPath -Message "TRANSFER FAILED: $($activeJob.Operation.Source) -> $($activeJob.Operation.Destination) | Error: $ErrorMsg"
                            }
                            
                            $completed.Failed++
                            $completed.Details += New-OperationResultDetail -Operation $activeJob.Operation -Result "Failed" -Error $errorMsg -OperationType $OperationType
                            Write-Host "✗ $($activeJob.Operation.Name) - $errorMsg" -ForegroundColor Red
                            Write-LogError -LogPath $LogPath -Message "BITS $OperationType transfer failed: $($activeJob.Operation.Source) - $errorMsg"
                        }
                        default {
                            # Still in progress
                            $newActiveJobs += $activeJob
                        }
                    }
                }
                catch {
                    Write-LogWarning -LogPath $LogPath -Message "Error checking BITS job: $($_.Exception.Message)"
                }
            }
            
            $activeJobs = $newActiveJobs
            
            # Progress update
            $total = $completed.Success + $completed.Failed
            if ($total -gt 0 -and $total % 10 -eq 0) {
                Write-Host "Progress: $total/$($Operations.Count) (Active: $($activeJobs.Count))" -ForegroundColor Cyan
            }
        }
    }
    
    # Cleanup any remaining jobs
    foreach ($activeJob in $activeJobs) {
        try {
            $job = Get-BitsTransfer -JobId $activeJob.Job.JobId -ErrorAction SilentlyContinue
            if ($job) { Remove-BitsTransfer -BitsJob $job -ErrorAction SilentlyContinue }
        }
        catch { 
            Write-LogDebug -LogPath $LogPath -Message "Error cleaning up BITS job: $($_.Exception.Message)"
        }
    }
    
    Write-LogInfo -LogPath $LogPath -Message "$OperationType operations completed - Success: $($completed.Success), Failed: $($completed.Failed)"
    return $completed
}

function New-OperationResult {
    <#
    .SYNOPSIS
        Creates a standardized result object for both copy and move operations.
    
    .PARAMETER OperationType
        Type of operation: "Copy" or "Move".
    
    .PARAMETER TotalFiles
        Total number of files processed.
    
    .PARAMETER SuccessfulOperations
        Number of successful operations.
    
    .PARAMETER FailedOperations
        Number of failed operations.
    
    .PARAMETER SkippedFiles
        Number of skipped files.
    
    .PARAMETER TotalSizeBytes
        Total size in bytes of all files.
    
    .PARAMETER Duration
        Duration of the operation.
    
    .PARAMETER Results
        Array of detailed results.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Copy", "Move")]
        [string]$OperationType,
        
        [int]$TotalFiles = 0,
        [int]$SuccessfulOperations = 0,
        [int]$FailedOperations = 0,
        [int]$SkippedFiles = 0,
        [long]$TotalSizeBytes = 0,
        [timespan]$Duration = [timespan]::Zero,
        [array]$Results = @()
    )
    
    $resultObject = [PSCustomObject]@{
        TotalFiles = $TotalFiles
        SkippedFiles = $SkippedFiles
        TotalSizeBytes = $TotalSizeBytes
        ExecutionTime = $Duration
        Results = $Results
    }
    
    # Add operation-specific properties
    if ($OperationType -eq "Copy") {
        $resultObject | Add-Member -MemberType NoteProperty -Name "SuccessfulCopies" -Value $SuccessfulOperations
        $resultObject | Add-Member -MemberType NoteProperty -Name "FailedCopies" -Value $FailedOperations
    } else {
        $resultObject | Add-Member -MemberType NoteProperty -Name "SuccessfulMoves" -Value $SuccessfulOperations
        $resultObject | Add-Member -MemberType NoteProperty -Name "FailedMoves" -Value $FailedOperations
    }
    
    return $resultObject
}

function New-OperationResultDetail {
    <#
    .SYNOPSIS
        Creates a detailed result object for individual file operations.
    
    .PARAMETER Operation
        The operation object containing file details.
    
    .PARAMETER Result
        Result of the operation (Success, Failed, etc.).
    
    .PARAMETER Error
        Error message if operation failed.
    
    .PARAMETER OperationType
        Type of operation: "Copy" or "Move".
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Operation,
        
        [Parameter(Mandatory = $true)]
        [string]$Result,
        
        [Parameter(Mandatory = $false)]
        [string]$Error = $null,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("Copy", "Move")]
        [string]$OperationType
    )
    
    $detailObject = [PSCustomObject]@{
        SourcePath = $Operation.Source
        DestinationPath = $Operation.Destination
        Result = $Result
        Error = $Error
        SizeBytes = $Operation.Size
        StartTime = Get-Date
        EndTime = Get-Date
    }
    
    # Add operation-specific properties
    if ($OperationType -eq "Copy" -and $Operation.PSObject.Properties.Name -contains "Status") {
        $detailObject | Add-Member -MemberType NoteProperty -Name "Status" -Value $Operation.Status
    } elseif ($OperationType -eq "Move" -and $Operation.PSObject.Properties.Name -contains "DriveLetter") {
        $detailObject | Add-Member -MemberType NoteProperty -Name "DriveLetter" -Value $Operation.DriveLetter
    }
    
    return $detailObject
}

function Show-OperationSummary {
    <#
    .SYNOPSIS
        Displays a summary of operation results for both copy and move operations.
    
    .PARAMETER Results
        Results object from BITS operations.
    
    .PARAMETER Duration
        Duration of the operation.
    
    .PARAMETER Skipped
        Number of skipped files.
    
    .PARAMETER LogPath
        Path to the log file.
    
    .PARAMETER OperationType
        Type of operation: "Copy" or "Move".
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Results,
        
        [Parameter(Mandatory = $true)]
        [timespan]$Duration,
        
        [Parameter(Mandatory = $true)]
        [int]$Skipped,
        
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("Copy", "Move")]
        [string]$OperationType
    )
    
    $verb = if ($OperationType -eq "Copy") { "copied" } else { "moved" }
    $verbPastTense = if ($OperationType -eq "Copy") { "copy" } else { "move" }
    
    Write-Host "`n$OperationType operation completed!" -ForegroundColor Green
    Write-Host "Successfully $verb`: $($Results.Success)" -ForegroundColor Green
    Write-Host "Failed to $verbPastTense`: $($Results.Failed)" -ForegroundColor $(if ($Results.Failed -gt 0) { "Red" } else { "Green" })
    Write-Host "Skipped: $Skipped" -ForegroundColor Yellow
    Write-Host "Duration: $([math]::Round($Duration.TotalSeconds, 2)) seconds" -ForegroundColor Cyan
    
    $logMsg = "$OperationType completed - Success: $($Results.Success), Failed: $($Results.Failed), Skipped: $Skipped, Duration: $([math]::Round($Duration.TotalSeconds, 2))s"
    Write-LogInfo -LogPath $LogPath -Message $logMsg
}

function Format-FileSize {
    <#
    .SYNOPSIS
        Formats file size in human-readable format.
    #>
    [CmdletBinding()]
    param([long]$Bytes)
    
    if ($Bytes -ge 1GB) { return "$([math]::Round($Bytes / 1GB, 2)) GB" }
    if ($Bytes -ge 1MB) { return "$([math]::Round($Bytes / 1MB, 2)) MB" }
    if ($Bytes -ge 1KB) { return "$([math]::Round($Bytes / 1KB, 2)) KB" }
    return "$Bytes bytes"
}

Export-ModuleMember -Function Get-NetworkShareSnapshot, Compare-NetworkShareSnapshots, Copy-SnapshotChangesToStaging, Move-StagingFilesToNetwork