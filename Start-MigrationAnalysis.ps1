# Start-MigrationAnalysis.ps1
# Main orchestration script for SharePoint migration analysis

<#
.SYNOPSIS
    Orchestrates the SharePoint migration analysis process.

.DESCRIPTION
    Reads configuration, manages checkpoint/restart, coordinates scanning,
    and triggers Python-based classification and reporting.

.PARAMETER ConfigFile
    Path to configuration JSON file (default: config.json)

.PARAMETER Resume
    Resume from last checkpoint if available. Will skip WizTree and permissions collection if data is fresh (within 24 hours) and scan paths unchanged.

.PARAMETER SkipScan
    Skip the file system scan and go directly to classification/reporting

.PARAMETER ForceScan
    Force a fresh scan even if existing data is available and fresh
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = ".\config.json",
    
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipScan,
    
    [Parameter(Mandatory=$false)]
    [switch]$ForceScan
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Function definitions
function Test-ExistingDataFreshness {
    param(
        [hashtable]$Config,
        [object]$Checkpoint
    )
    
    # Check if required output files exist
    if (-not (Test-Path $Config.rawDataFile)) {
        Write-Host "  Raw data file missing: $($Config.rawDataFile)" -ForegroundColor Yellow
        return $false
    }
    
    if (-not (Test-Path $Config.permissionsFile)) {
        Write-Host "  Permissions file missing: $($Config.permissionsFile)" -ForegroundColor Yellow
        return $false
    }
    
    # Check if data is recent (within last 24 hours)
    $maxAgeHours = 24
    $rawDataAge = (Get-Date) - (Get-Item $Config.rawDataFile).LastWriteTime
    $permissionsAge = (Get-Date) - (Get-Item $Config.permissionsFile).LastWriteTime
    
    if ($rawDataAge.TotalHours -gt $maxAgeHours) {
        Write-Host "  Raw data is older than $maxAgeHours hours: $([math]::Round($rawDataAge.TotalHours, 1))h" -ForegroundColor Yellow
        return $false
    }
    
    if ($permissionsAge.TotalHours -gt $maxAgeHours) {
        Write-Host "  Permissions data is older than $maxAgeHours hours: $([math]::Round($permissionsAge.TotalHours, 1))h" -ForegroundColor Yellow
        return $false
    }
    
    # Check if scan paths have changed (only if checkpoint has scanPaths)
    if ($Checkpoint.scanPaths) {
        $currentPaths = @($Config.paths | Sort-Object)
        $checkpointPaths = @($Checkpoint.scanPaths | Sort-Object)
        
        if (Compare-Object $currentPaths $checkpointPaths) {
            Write-Host "  Scan paths have changed since last run" -ForegroundColor Yellow
            Write-Host "    Current: $($currentPaths -join ', ')" -ForegroundColor Gray
            Write-Host "    Checkpoint: $($checkpointPaths -join ', ')" -ForegroundColor Gray
            return $false
        }
    } else {
        # Old checkpoint format - assume paths might have changed
        Write-Host "  Old checkpoint format - cannot verify scan paths unchanged" -ForegroundColor Yellow
        return $false
    }
    
    # Check if any scan paths have been modified since last scan
    if ($Checkpoint.lastScanTime) {
        $lastScanTime = [DateTime]::Parse($Checkpoint.lastScanTime)
        foreach ($scanPath in $Config.paths) {
            if (Test-Path $scanPath) {
                $pathLastWrite = (Get-Item $scanPath).LastWriteTime
                if ($pathLastWrite -gt $lastScanTime) {
                    Write-Host "  Scan path modified since last scan: $scanPath" -ForegroundColor Yellow
                    return $false
                }
            }
        }
    } else {
        # Old checkpoint format - cannot verify timestamps
        Write-Host "  Old checkpoint format - cannot verify scan timestamps" -ForegroundColor Yellow
        return $false
    }
    
    Write-Host "  Data is fresh and scan paths unchanged" -ForegroundColor Green
    return $true
}

function Save-ScanCheckpoint {
    param(
        [hashtable]$Config,
        [array]$ScanPaths
    )
    
    try {
        $checkpoint = @{
            lastScanTime = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
            scanPaths = $ScanPaths
            rawDataFile = $Config.rawDataFile
            permissionsFile = $Config.permissionsFile
            lastCompletedPath = if ($ScanPaths.Count -gt 0) { $ScanPaths[-1] } else { "" }
            filesProcessed = 0  # Will be updated by actual scan logic
            foldersProcessed = 0  # Will be updated by actual scan logic
        }
        
        $checkpoint | ConvertTo-Json -Depth 3 | Set-Content -Path $Config.checkpointFile -Encoding UTF8
        Write-Host "Checkpoint saved: $($Config.checkpointFile)" -ForegroundColor Gray
    } catch {
        Write-Warning "Failed to save checkpoint: $_"
    }
}

# Display banner
Write-Host @"
╔═══════════════════════════════════════════════════════════════╗
║   SharePoint Migration Analysis Tool                         ║
║   File System Scanner & Department Classifier                ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

# Load configuration
Write-Host "`n[1/4] Loading configuration..." -ForegroundColor Yellow

if (-not (Test-Path $ConfigFile)) {
    Write-Error "Configuration file not found: $ConfigFile"
    exit 1
}

try {
    $configJson = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
    
    # Convert to hashtable for easier handling
    $config = @{
        paths = $configJson.paths
        outputDirectory = $configJson.outputDirectory
        rawDataFile = $configJson.rawDataFile
        permissionsFile = $configJson.permissionsFile
        checkpointFile = $configJson.checkpointFile
        excelOutputFile = $configJson.excelOutputFile
        thresholds = @{
            maxPathLength = $configJson.thresholds.maxPathLength
            maxFileSize = $configJson.thresholds.maxFileSize
            maxFilesPerFolder = $configJson.thresholds.maxFilesPerFolder
        }
        unsafeExtensions = $configJson.unsafeExtensions
        unsupportedCharacters = $configJson.unsupportedCharacters
    }
    
    # Generate timestamped Excel filename to avoid conflicts
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($config.excelOutputFile)
    $directory = [System.IO.Path]::GetDirectoryName($config.excelOutputFile)
    $config.excelOutputFile = Join-Path $directory "$baseFileName`_$timestamp.xlsx"
    
    # Optional: Include departmentKeywordsFile if specified (for custom keywords)
    if ($configJson.PSObject.Properties['departmentKeywordsFile'] -and $configJson.departmentKeywordsFile) {
        $config.departmentKeywordsFile = $configJson.departmentKeywordsFile
    }
    
    Write-Host "Configuration loaded successfully" -ForegroundColor Green
    
    # Auto-discover local SMB shares if no paths configured
    if (-not $config.paths -or $config.paths.Count -eq 0) {
        Write-Host "`n  No paths configured - discovering local SMB shares..." -ForegroundColor Yellow
        
        try {
            $discoveredPaths = @()
            
            Get-SmbShare | Where-Object { 
                -not $_.Special -and 
                $_.Name -ne 'CertEnroll' -and 
                $_.Name -ne 'print$' 
            } | ForEach-Object {
                $sharePath = $_.Path
                if ($sharePath -and (Test-Path $sharePath)) {
                    $discoveredPaths += $sharePath
                    Write-Host "    Found: $($_.Name) -> $sharePath" -ForegroundColor Gray
                }
            }
            
            if ($discoveredPaths.Count -gt 0) {
                $config.paths = $discoveredPaths
                Write-Host "  Discovered $($discoveredPaths.Count) local shares" -ForegroundColor Green
            } else {
                Write-Error "No paths configured and no local shares discovered"
                exit 1
            }
        } catch {
            Write-Error "Failed to discover local shares: $_"
            exit 1
        }
    } else {
        Write-Host "  Paths to scan: $($config.paths.Count)" -ForegroundColor Gray
    }
    
    Write-Host "  Output directory: $($config.outputDirectory)" -ForegroundColor Gray
    
} catch {
    Write-Error "Failed to load configuration: $_"
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path $config.outputDirectory)) {
    New-Item -Path $config.outputDirectory -ItemType Directory -Force | Out-Null
    Write-Host "Created output directory: $($config.outputDirectory)" -ForegroundColor Green
}

# Check for checkpoint and validate data freshness
$checkpoint = $null
$useExistingData = $false

if ($Resume -and (Test-Path $config.checkpointFile)) {
    Write-Host "`n[2/4] Loading checkpoint..." -ForegroundColor Yellow
    try {
        $checkpoint = Get-Content -Path $config.checkpointFile -Raw | ConvertFrom-Json
        Write-Host "Checkpoint loaded - Last completed: $($checkpoint.lastCompletedPath)" -ForegroundColor Green
        Write-Host "  Files processed: $($checkpoint.filesProcessed)" -ForegroundColor Gray
        Write-Host "  Folders processed: $($checkpoint.foldersProcessed)" -ForegroundColor Gray
        
        # Check if we can use existing data (unless ForceScan is specified)
        if ($ForceScan) {
            Write-Host "`n[2/4] ForceScan specified - will re-scan regardless of existing data..." -ForegroundColor Yellow
            $useExistingData = $false
        } else {
            $useExistingData = Test-ExistingDataFreshness -Config $config -Checkpoint $checkpoint
            if ($useExistingData) {
                Write-Host "`n[2/4] Using existing scan data (still fresh)..." -ForegroundColor Green
            } else {
                Write-Host "`n[2/4] Existing data is stale, will re-scan..." -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Warning "Failed to load checkpoint: $_"
        $checkpoint = $null
    }
}

function Invoke-WizTreeExport {
    param(
        [string]$TargetPath,
        [string]$OutputCsv,
        [string]$WizTreeExe
    )
    
    Write-Host "Starting WizTree export for: $TargetPath" -ForegroundColor Gray
    
    # Start the WizTree process and wait for it to complete
    $process = Start-Process -FilePath $WizTreeExe -ArgumentList @(
        $TargetPath,
        "/export=`"$OutputCsv`"",
        "/admin=1",
        "/exportUTCTime=1", 
        "/exportfolders=1",
        "/exportfiles=1"
    ) -Wait -PassThru -WindowStyle Hidden
    
    if ($process.ExitCode -eq 0) {
        Write-Host "WizTree export completed successfully" -ForegroundColor Green
    } else {
        Write-Warning "WizTree export completed with exit code: $($process.ExitCode)"
    }
}

function Convert-WizTreeCsvToRawSchema {
    param(
        [string[]]$InputCsvPaths,
        [string]$OutputCsv,
        [hashtable]$Config
    )

    Write-Host "Converting WizTree exports to raw schema..." -ForegroundColor Yellow

    $maxPathLength = [int]$Config.thresholds.maxPathLength
    $maxFileSize = [long]$Config.thresholds.maxFileSize
    $maxFilesPerFolder = [int]$Config.thresholds.maxFilesPerFolder
    $unsafeExt = @($Config.unsafeExtensions)
    $unsupportedChars = @($Config.unsupportedCharacters)

    # Pre-compile regex patterns for unsupported characters (major performance boost)
    $unsupportedPatterns = @()
    foreach ($ch in $unsupportedChars) {
        if ([string]::IsNullOrEmpty($ch)) { continue }
        $escapedChar = [regex]::Escape($ch)
        $unsupportedPatterns += [regex]::new($escapedChar, [System.Text.RegularExpressions.RegexOptions]::Compiled)
    }

    # Create output directory if needed
    $outputDir = Split-Path -Path $OutputCsv -Parent
    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
    }

    # Use StringBuilder for efficient CSV writing
    $csvBuilder = New-Object System.Text.StringBuilder
    $csvBuilder.AppendLine("Path,Name,Extension,SizeBytes,Created,LastModified,Type,PathLength,HasUnsupportedChars,IsUnsafeExtension,IsLargeFile,IsTooManyFiles,IsTooLongPath,HasExplicitPermissions,FileCountInFolder") | Out-Null

    $totalRows = 0
    $processedRows = 0

    foreach ($csvPath in $InputCsvPaths) {
        if (-not (Test-Path $csvPath)) { 
            Write-Warning "CSV file not found: $csvPath"
            continue 
        }
        
        Write-Host "Processing WizTree CSV: $csvPath" -ForegroundColor Gray
        
        # Stream processing - read file line by line instead of loading all into memory
        $reader = [System.IO.File]::OpenText($csvPath)
        $lineNumber = 0
        $headerLine = $null
        
        try {
            while (($line = $reader.ReadLine()) -ne $null) {
                $lineNumber++
                
                # Skip comment line (line 1)
                if ($lineNumber -eq 1) { continue }
                
                # Store header line (line 2)
                if ($lineNumber -eq 2) { 
                    $headerLine = $line
                    continue 
                }
                
                # Process data lines
                if ($lineNumber -gt 2) {
                    $totalRows++
                    
                    # Parse CSV line manually (much faster than Import-Csv)
                    $fields = Parse-CsvLine $line
                    if ($fields.Count -lt 6) { continue }  # Skip malformed lines
                    
                    $fullName = $fields[0]
                    if ([string]::IsNullOrEmpty($fullName)) { continue }

                    $isFolder = $fullName.EndsWith("\")
                    $path = if ($isFolder) { $fullName.TrimEnd('\') } else { $fullName }
                    $name = Split-Path -Path $path -Leaf
                    $extension = if ($isFolder) { "" } else { [System.IO.Path]::GetExtension($name) }
                    
                    # Parse size efficiently
                    $sizeBytes = 0
                    if (-not $isFolder -and $fields[1]) {
                        if ([long]::TryParse($fields[1], [ref]$sizeBytes)) {
                            # Success
                        } else {
                            # Try decimal parsing for large numbers
                            if ([decimal]::TryParse($fields[1], [ref]$sizeBytes)) {
                                $sizeBytes = [long]$sizeBytes
                            }
                        }
                    }
                    
                    $lastModified = $fields[3]
                    $created = $lastModified
                    $pathLength = $path.Length
                    
                    # Parse file count for folders
                    $fileCountInFolder = 0
                    if ($isFolder -and $fields.Count -gt 5 -and $fields[5]) {
                        [int]::TryParse($fields[5], [ref]$fileCountInFolder) | Out-Null
                    }

                    # Optimized unsupported character checking
                    $hasUnsupportedChars = $false
                    if ($unsupportedPatterns.Count -gt 0) {
                        foreach ($pattern in $unsupportedPatterns) {
                            if ($pattern.IsMatch($name)) {
                                $hasUnsupportedChars = $true
                                break
                            }
                        }
                    }

                    # Check unsafe extension
                    $isUnsafeExtension = $false
                    if (-not $isFolder -and $extension) {
                        $isUnsafeExtension = $unsafeExt -contains $extension.ToLower()
                    }

                    # Calculate flags
                    $isLargeFile = (-not $isFolder -and $sizeBytes -gt $maxFileSize)
                    $isTooManyFiles = ($isFolder -and $fileCountInFolder -gt $maxFilesPerFolder)
                    $isTooLongPath = ($pathLength -gt $maxPathLength)

                    # Build CSV line directly (much faster than PSCustomObject)
                    $csvLine = "$path,$name,$extension,$sizeBytes,$created,$lastModified,$(if ($isFolder) { 'Folder' } else { 'File' }),$pathLength,$(if ($hasUnsupportedChars) { 'True' } else { 'False' }),$(if ($isUnsafeExtension) { 'True' } else { 'False' }),$(if ($isLargeFile) { 'True' } else { 'False' }),$(if ($isTooManyFiles) { 'True' } else { 'False' }),$(if ($isTooLongPath) { 'True' } else { 'False' }),False,$fileCountInFolder"
                    $csvBuilder.AppendLine($csvLine) | Out-Null
                    
                    $processedRows++
                    
                    # Progress indicator for large files
                    if ($processedRows % 10000 -eq 0) {
                        Write-Host "  Processed $processedRows rows..." -ForegroundColor Gray
                    }
                }
            }
        } finally {
            $reader.Close()
        }
        
        Write-Host "  Found $totalRows rows in CSV" -ForegroundColor Gray
    }
    
    Write-Host "Total rows to export: $processedRows" -ForegroundColor Cyan
    if ($processedRows -eq 0) {
        Write-Warning "No rows to export - check WizTree CSV files"
        return
    }
    
    # Write all data at once (much faster than Export-Csv)
    [System.IO.File]::WriteAllText($OutputCsv, $csvBuilder.ToString(), [System.Text.Encoding]::UTF8)
    Write-Host "Raw data CSV written: $OutputCsv" -ForegroundColor Green
}

# Helper function to parse CSV line manually (much faster than Import-Csv)
function Parse-CsvLine {
    param([string]$line)
    
    $fields = @()
    $currentField = ""
    $inQuotes = $false
    $i = 0
    
    while ($i -lt $line.Length) {
        $char = $line[$i]
        
        if ($char -eq '"') {
            if ($inQuotes -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '"') {
                # Escaped quote
                $currentField += '"'
                $i += 2
            } else {
                # Toggle quote state
                $inQuotes = -not $inQuotes
                $i++
            }
        } elseif ($char -eq ',' -and -not $inQuotes) {
            # Field separator
            $fields += $currentField
            $currentField = ""
            $i++
        } else {
            $currentField += $char
            $i++
        }
    }
    
    # Add the last field
    $fields += $currentField
    
    return $fields
}

function ApplyExplicitPermissionsFlag {
    param(
        [string]$RawDataCsv,
        [string]$PermissionsCsv
    )

    if (-not (Test-Path $PermissionsCsv)) {
        Write-Warning "Permissions file not found: $PermissionsCsv. 'HasExplicitPermissions' will remain false."
        return
    }

    Write-Host "Applying explicit permissions flags from permissions CSV..." -ForegroundColor Yellow

    $rawRows = Import-Csv -Path $RawDataCsv
    $permRows = Import-Csv -Path $PermissionsCsv

    $explicitPaths = New-Object System.Collections.Generic.List[string]

    $hasIsInherited = $false
    if ($permRows.Count -gt 0) {
        $hasIsInherited = $permRows[0].PSObject.Properties.Name -contains 'IsInherited'
    }

    foreach ($pr in $permRows) {
        $pPath = $pr.Path
        if (-not $pPath) { continue }
        if ($hasIsInherited) {
            $isInheritedVal = $pr.IsInherited
            $isInherited = $false
            if ($isInheritedVal -is [bool]) { $isInherited = [bool]$isInheritedVal }
            else { $isInherited = ($isInheritedVal -eq 'True' -or $isInheritedVal -eq 'true') }
            if (-not $isInherited) { $explicitPaths.Add($pPath) | Out-Null }
        } else {
            $explicitPaths.Add($pPath) | Out-Null
        }
    }

    # Deduplicate and sort by length desc for prefix matching
    $explicitPaths = [System.Linq.Enumerable]::ToList([System.Linq.Enumerable]::Distinct($explicitPaths))
    $explicitPaths = $explicitPaths | Sort-Object Length -Descending

    foreach ($row in $rawRows) {
        $itemPath = $row.Path
        $flag = $false
        foreach ($base in $explicitPaths) {
            if ($itemPath.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) { $flag = $true; break }
        }
        $row.HasExplicitPermissions = $flag
    }

    $rawRows | Export-Csv -Path $RawDataCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Updated raw data with HasExplicitPermissions." -ForegroundColor Green
}

# File system scanning phase (replaced with WizTree export)
if (-not $SkipScan -and -not $useExistingData) {
    Write-Host "`n[2/4] Starting file system scan (WizTree)..." -ForegroundColor Yellow

    $wizTreeExe = Join-Path $PSScriptRoot 'WizTree64.exe'
    if (-not (Test-Path $wizTreeExe)) {
        $wizTreeExe = Join-Path $PSScriptRoot 'WizTree.exe'
    }
    if (-not (Test-Path $wizTreeExe)) {
        Write-Error "WizTree executable not found in script directory. Expected WizTree64.exe or WizTree.exe"
        exit 1
    }
    
    $totalPaths = $config.paths.Count
    $pathIndex = 0
    $tempDir = Join-Path $config.outputDirectory 'wiztree_tmp'
    if (-not (Test-Path $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }
    $exportParts = @()
    
    foreach ($scanPath in $config.paths) {
        $pathIndex++
        
        Write-Host "`n--- Scanning Path $pathIndex of $totalPaths ---" -ForegroundColor Cyan
        Write-Host "Path: $scanPath" -ForegroundColor White
        
        if (-not (Test-Path $scanPath)) {
            Write-Warning "Path does not exist, skipping: $scanPath"
            continue
        }
        
        # Skip path if we're using existing data (already handled by $useExistingData)
        if ($useExistingData) {
            Write-Host "Skipping WizTree export (using existing data)" -ForegroundColor Yellow
            continue
        }
        
        if ($checkpoint -and $checkpoint.lastCompletedPath -and $scanPath -eq $checkpoint.lastCompletedPath -and -not $useExistingData) {
            Write-Host "Skipping already completed path (from checkpoint)" -ForegroundColor Yellow
            continue
        }
        
        try {
            $out = Join-Path $tempDir ("wiztree_" + [Guid]::NewGuid().ToString() + ".csv")
            Write-Host "Exporting to: $out" -ForegroundColor Gray
            Invoke-WizTreeExport -TargetPath $scanPath -OutputCsv $out -WizTreeExe $wizTreeExe
            
            if (Test-Path $out) {
                $fileSize = (Get-Item $out).Length
                Write-Host "WizTree export created: $out (Size: $fileSize bytes)" -ForegroundColor Green
                $exportParts += $out
            } else {
                Write-Warning "WizTree export file not found: $out"
            }
            
            Write-Host "Completed WizTree export for: $scanPath" -ForegroundColor Green
        } catch {
            Write-Error "Failed to export via WizTree for $scanPath : $_"
        }
    }

    # Convert to raw schema expected by classifier
    Convert-WizTreeCsvToRawSchema -InputCsvPaths $exportParts -OutputCsv $config.rawDataFile -Config $config

    # Collect permissions for each path using Get-FileSystemAnalysis.ps1
    Write-Host "`nCollecting permissions data..." -ForegroundColor Yellow
    
    $permissionsCollected = $false
    foreach ($scanPath in $config.paths) {
        if (-not (Test-Path $scanPath)) {
            Write-Warning "Path does not exist, skipping permissions collection: $scanPath"
            continue
        }
        
        try {
            Write-Host "Collecting permissions for: $scanPath" -ForegroundColor Gray
            
            # Call Get-FileSystemAnalysis.ps1 for permissions only
            $permissionsParams = @{
                Path = $scanPath
                Config = $config
                RawDataFile = $config.rawDataFile
                PermissionsFile = $config.permissionsFile
                CheckpointFile = $config.checkpointFile
                Resume = $Resume
                PermissionsOnly = $true
            }
            
            # Run permissions collection (this will append to permissionsFile)
            & ".\Get-FileSystemAnalysis.ps1" @permissionsParams
            
            $permissionsCollected = $true
            Write-Host "Completed permissions collection for: $scanPath" -ForegroundColor Green
            
        } catch {
            Write-Error "Failed to collect permissions for $scanPath : $_"
        }
    }
    
    if ($permissionsCollected) {
        Write-Host "Permissions collection completed!" -ForegroundColor Green
    } else {
        Write-Warning "No permissions were collected - HasExplicitPermissions will remain false"
    }

    # Apply permissions flag if permissions CSV is present
    ApplyExplicitPermissionsFlag -RawDataCsv $config.rawDataFile -PermissionsCsv $config.permissionsFile
    
    # Save checkpoint with actual processed paths
    $processedPaths = @()
    foreach ($scanPath in $config.paths) {
        if (Test-Path $scanPath) {
            $processedPaths += $scanPath
        }
    }
    Save-ScanCheckpoint -Config $config -ScanPaths $processedPaths
    
    Write-Host "`n[3/4] File system scan completed!" -ForegroundColor Green
} elseif ($useExistingData) {
    Write-Host "`n[2/4] Using existing scan data (skipping WizTree and permissions collection)..." -ForegroundColor Green
    Write-Host "  Raw data file: $($config.rawDataFile)" -ForegroundColor Gray
    Write-Host "  Permissions file: $($config.permissionsFile)" -ForegroundColor Gray
} else {
    Write-Host "`n[2/4] Skipping file system scan (SkipScan flag set)" -ForegroundColor Yellow
}

# Classification and reporting phase
Write-Host "`n[3/4] Starting classification and report generation..." -ForegroundColor Yellow

# Run setup validation to ensure uv is available
Write-Host "Validating setup and dependencies..." -ForegroundColor Yellow

try {
    # Run Test-Setup.ps1 to check and install uv if needed
    & ".\Test-Setup.ps1" -ConfigFile $ConfigFile
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Setup validation failed. Please check the errors above and fix them before continuing."
        exit 1
    }
    
    Write-Host "Setup validation completed successfully" -ForegroundColor Green
} catch {
    Write-Error "Failed to run setup validation: $_"
    exit 1
}

# Run Python classification and reporting with uv
try {
    Write-Host "Running classification and report generation (uv will handle dependencies automatically)..." -ForegroundColor Yellow
    
    # Build uv run command - keywords are now embedded, but can be overridden with custom file
    $uvArgs = @(
        "run",
        ".\classify_and_report.py",
        "--config", $ConfigFile,
        "--raw-data", $config.rawDataFile,
        "--permissions", $config.permissionsFile,
        "--output", $config.excelOutputFile,
        "--use-ai"
    )
    
    
    & uv @uvArgs
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`n[4/4] Report generation completed!" -ForegroundColor Green
        Write-Host "`nExcel report saved to: $($config.excelOutputFile)" -ForegroundColor Cyan
    } else {
        Write-Error "Python script failed with exit code: $LASTEXITCODE"
    }
    
} catch {
    Write-Error "Failed to run classification and reporting: $_"
    exit 1
}

# Display summary
Write-Host @"

╔═══════════════════════════════════════════════════════════════╗
║   Analysis Complete!                                          ║
╚═══════════════════════════════════════════════════════════════╝

Output files:
  - Excel Report: $($config.excelOutputFile)
  - Raw Data: $($config.rawDataFile)
  - Permissions: $($config.permissionsFile)
  - Checkpoint: $($config.checkpointFile)

"@ -ForegroundColor Green

Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

