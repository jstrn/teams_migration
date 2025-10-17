# Test-CsvPerformance.ps1
# Performance testing script for optimized WizTree CSV processing

param(
    [string]$LargeCsvPath = ".\wiztree_large.csv",
    [string]$OutputPath = ".\output\performance_test.csv"
)

# =============================================================================
# OPTIMIZED FUNCTIONS (copied from Start-MigrationAnalysis.ps1)
# =============================================================================

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

# Load configuration
$configPath = ".\config.json"
if (-not (Test-Path $configPath)) {
    Write-Error "Configuration file not found: $configPath"
    exit 1
}

$configJson = Get-Content -Path $configPath -Raw | ConvertFrom-Json
$config = @{
    thresholds = @{
        maxPathLength = $configJson.thresholds.maxPathLength
        maxFileSize = $configJson.thresholds.maxFileSize
        maxFilesPerFolder = $configJson.thresholds.maxFilesPerFolder
    }
    unsafeExtensions = $configJson.unsafeExtensions
    unsupportedCharacters = $configJson.unsupportedCharacters
}

Write-Host "=== CSV Processing Performance Test ===" -ForegroundColor Cyan
Write-Host "Large CSV: $LargeCsvPath" -ForegroundColor Gray
Write-Host "Output: $OutputPath" -ForegroundColor Gray

# Check if large CSV exists
if (-not (Test-Path $LargeCsvPath)) {
    Write-Error "Large CSV file not found: $LargeCsvPath"
    exit 1
}

# Get file size for reference
$fileSize = (Get-Item $LargeCsvPath).Length
$fileSizeMB = [math]::Round($fileSize / 1MB, 2)
Write-Host "File size: $fileSizeMB MB" -ForegroundColor Gray

# Count lines for reference
$lineCount = (Get-Content $LargeCsvPath | Measure-Object -Line).Lines
Write-Host "Total lines: $lineCount" -ForegroundColor Gray

# Create output directory
$outputDir = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

Write-Host "`nStarting optimized CSV processing..." -ForegroundColor Yellow
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Test the optimized function
try {
    Convert-WizTreeCsvToRawSchema -InputCsvPaths @($LargeCsvPath) -OutputCsv $OutputPath -Config $config
    $stopwatch.Stop()
    
    $elapsedSeconds = $stopwatch.Elapsed.TotalSeconds
    $rowsPerSecond = [math]::Round($lineCount / $elapsedSeconds, 0)
    $mbPerSecond = [math]::Round($fileSizeMB / $elapsedSeconds, 2)
    
    Write-Host "`n=== Performance Results ===" -ForegroundColor Green
    Write-Host "Processing time: $([math]::Round($elapsedSeconds, 2)) seconds" -ForegroundColor Green
    Write-Host "Rows per second: $rowsPerSecond" -ForegroundColor Green
    Write-Host "MB per second: $mbPerSecond" -ForegroundColor Green
    
    # Check output file
    if (Test-Path $OutputPath) {
        $outputSize = (Get-Item $OutputPath).Length
        $outputSizeMB = [math]::Round($outputSize / 1MB, 2)
        $outputLines = (Get-Content $OutputPath | Measure-Object -Line).Lines
        
        Write-Host "`n=== Output Results ===" -ForegroundColor Cyan
        Write-Host "Output file size: $outputSizeMB MB" -ForegroundColor Gray
        Write-Host "Output lines: $outputLines" -ForegroundColor Gray
        Write-Host "Compression ratio: $([math]::Round($fileSizeMB / $outputSizeMB, 2))x" -ForegroundColor Gray
        
        # Show sample of output
        Write-Host "`n=== Sample Output (first 5 lines) ===" -ForegroundColor Cyan
        Get-Content $OutputPath | Select-Object -First 5 | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
    } else {
        Write-Warning "Output file was not created"
    }
    
} catch {
    $stopwatch.Stop()
    Write-Error "Processing failed: $_"
    Write-Host "Elapsed time: $([math]::Round($stopwatch.Elapsed.TotalSeconds, 2)) seconds" -ForegroundColor Red
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Cyan
