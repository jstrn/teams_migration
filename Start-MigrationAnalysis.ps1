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
    Resume from last checkpoint if available

.PARAMETER SkipScan
    Skip the file system scan and go directly to classification/reporting
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = ".\config.json",
    
    [Parameter(Mandatory=$false)]
    [switch]$Resume,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipScan
)

# Set error action preference
$ErrorActionPreference = "Stop"

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

# Check for checkpoint
$checkpoint = $null

if ($Resume -and (Test-Path $config.checkpointFile)) {
    Write-Host "`n[2/4] Loading checkpoint..." -ForegroundColor Yellow
    try {
        $checkpoint = Get-Content -Path $config.checkpointFile -Raw | ConvertFrom-Json
        Write-Host "Checkpoint loaded - Last completed: $($checkpoint.lastCompletedPath)" -ForegroundColor Green
        Write-Host "  Files processed: $($checkpoint.filesProcessed)" -ForegroundColor Gray
        Write-Host "  Folders processed: $($checkpoint.foldersProcessed)" -ForegroundColor Gray
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
    & $WizTreeExe $TargetPath /export="$OutputCsv" /admin=1 /exportUTCTime=1 /exportfolders=1 /exportfiles=1 | Out-Null
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

    $allRows = New-Object System.Collections.Generic.List[object]

    foreach ($csvPath in $InputCsvPaths) {
        if (-not (Test-Path $csvPath)) { continue }
        $rows = Import-Csv -Path $csvPath
        foreach ($r in $rows) {
            $fullName = [string]$r.'File Name'
            if (-not $fullName) { continue }

            $isFolder = $fullName.EndsWith("\")
            $path = if ($isFolder) { $fullName.TrimEnd('\') } else { $fullName }
            $name = Split-Path -Path $path -Leaf
            $extension = if ($isFolder) { "" } else { [System.IO.Path]::GetExtension($name) }
            $sizeBytes = if ($isFolder) { 0 } else { [long]($r.Size -as [decimal]) }
            $lastModified = [string]$r.Modified
            $created = $lastModified
            $pathLength = ($path).Length
            $fileCountInFolder = 0
            if ($isFolder -and $r.PSObject.Properties.Name -contains 'Files' -and $r.Files) {
                [int]::TryParse($r.Files, [ref]$fileCountInFolder) | Out-Null
            }

            $hasUnsupportedChars = $false
            foreach ($ch in $unsupportedChars) {
                if ([string]::IsNullOrEmpty($ch)) { continue }
                if ($path -like ("*" + $ch + "*")) { $hasUnsupportedChars = $true; break }
            }

            $isUnsafeExtension = $false
            if (-not $isFolder -and $extension) {
                $isUnsafeExtension = $unsafeExt -contains $extension.ToLower()
            }

            $isLargeFile = (-not $isFolder -and $sizeBytes -gt $maxFileSize)
            $isTooManyFiles = ($isFolder -and $fileCountInFolder -gt $maxFilesPerFolder)
            $isTooLongPath = ($pathLength -gt $maxPathLength)

            $obj = [PSCustomObject]@{
                'Path' = $path
                'Name' = $name
                'Extension' = $extension
                'SizeBytes' = $sizeBytes
                'Created' = $created
                'LastModified' = $lastModified
                'Type' = if ($isFolder) { 'Folder' } else { 'File' }
                'PathLength' = $pathLength
                'HasUnsupportedChars' = [bool]$hasUnsupportedChars
                'IsUnsafeExtension' = [bool]$isUnsafeExtension
                'IsLargeFile' = [bool]$isLargeFile
                'IsTooManyFiles' = [bool]$isTooManyFiles
                'IsTooLongPath' = [bool]$isTooLongPath
                'HasExplicitPermissions' = $false
                'FileCountInFolder' = $fileCountInFolder
            }
            $allRows.Add($obj) | Out-Null
        }
    }

    if (-not (Test-Path (Split-Path -Path $OutputCsv -Parent))) {
        New-Item -Path (Split-Path -Path $OutputCsv -Parent) -ItemType Directory -Force | Out-Null
    }
    $allRows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Raw data CSV written: $OutputCsv" -ForegroundColor Green
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
if (-not $SkipScan) {
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

        if ($checkpoint -and $checkpoint.lastCompletedPath -and $scanPath -eq $checkpoint.lastCompletedPath) {
            Write-Host "Skipping already completed path (from checkpoint)" -ForegroundColor Yellow
            continue
        }

        try {
            $out = Join-Path $tempDir ("wiztree_" + [Guid]::NewGuid().ToString() + ".csv")
            Invoke-WizTreeExport -TargetPath $scanPath -OutputCsv $out -WizTreeExe $wizTreeExe
            $exportParts += $out
            Write-Host "Completed WizTree export for: $scanPath" -ForegroundColor Green
        } catch {
            Write-Error "Failed to export via WizTree for $scanPath : $_"
        }
    }

    # Convert to raw schema expected by classifier
    Convert-WizTreeCsvToRawSchema -InputCsvPaths $exportParts -OutputCsv $config.rawDataFile -Config $config

    # Apply permissions flag if permissions CSV is present (kept as-is per plan)
    ApplyExplicitPermissionsFlag -RawDataCsv $config.rawDataFile -PermissionsCsv $config.permissionsFile

    Write-Host "`n[3/4] File system scan completed!" -ForegroundColor Green
} else {
    Write-Host "`n[2/4] Skipping file system scan (SkipScan flag set)" -ForegroundColor Yellow
}

# Classification and reporting phase
Write-Host "`n[3/4] Starting classification and report generation..." -ForegroundColor Yellow

# Check if uv is available
Write-Host "Checking for uv..." -ForegroundColor Yellow

try {
    $uvVersion = & uv --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Found uv: $uvVersion" -ForegroundColor Green
    } else {
        throw "uv not found"
    }
} catch {
    Write-Error @"
uv is not installed. Please install uv to continue:

  Windows (PowerShell):
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"

  Or visit: https://docs.astral.sh/uv/getting-started/installation/
"@
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
        "--output", $config.excelOutputFile
    )
    
    # If a custom keywords file is specified and exists, use it
    if ($config.PSObject.Properties['departmentKeywordsFile'] -and 
        $config.departmentKeywordsFile -and 
        (Test-Path $config.departmentKeywordsFile)) {
        Write-Host "Using custom keywords file: $($config.departmentKeywordsFile)" -ForegroundColor Cyan
        $uvArgs += "--keywords"
        $uvArgs += $config.departmentKeywordsFile
    }
    
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

