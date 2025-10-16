# Test-Setup.ps1
# Validation script to verify setup and requirements

<#
.SYNOPSIS
    Validates the SharePoint Migration Analysis Tool setup.

.DESCRIPTION
    Checks all prerequisites and dependencies before running the analysis.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = ".\config.json"
)

$ErrorCount = 0
$WarningCount = 0


Write-Host "`nValidating setup...`n" -ForegroundColor Yellow

# Test 1: PowerShell Version
Write-Host "[1/10] Checking PowerShell version..." -NoNewline
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -ge 5) {
    Write-Host " OK (v$psVersion)" -ForegroundColor Green
} else {
    Write-Host " FAIL" -ForegroundColor Red
    Write-Host "       PowerShell 5.1+ required, found v$psVersion" -ForegroundColor Red
    $ErrorCount++
}

# Test 2: Required Scripts
Write-Host "[2/10] Checking required scripts..." -NoNewline
$requiredScripts = @(
    ".\Get-FileSystemAnalysis.ps1",
    ".\Start-MigrationAnalysis.ps1"
)

$missingScripts = @()
foreach ($script in $requiredScripts) {
    if (-not (Test-Path $script)) {
        $missingScripts += $script
    }
}

if ($missingScripts.Count -eq 0) {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " FAIL" -ForegroundColor Red
    foreach ($script in $missingScripts) {
        Write-Host "       Missing: $script" -ForegroundColor Red
    }
    $ErrorCount++
}

# Test 3: uv Installation
Write-Host "[3/10] Checking uv installation..." -NoNewline

try {
    $uvVersion = & uv --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host " OK ($uvVersion)" -ForegroundColor Green
    } else {
        throw "uv not found"
    }
} catch {
    Write-Host " INSTALLING" -ForegroundColor Yellow
    Write-Host "       Installing uv..." -ForegroundColor Yellow
    
    try {
        # Set TLS 1.2 for secure downloads
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        # Install uv
        irm https://astral.sh/uv/install.ps1 | iex
        
        # Update PATH for current session
        $Env:PATH = "~\.local\bin;$Env:PATH"
        
        # Install Python 3.12
        Write-Host "       Installing Python 3.12..." -ForegroundColor Yellow
        uv python install 3.12
        
        # Verify installation
        $uvVersion = & uv --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "       Installation successful: $uvVersion" -ForegroundColor Green
        } else {
            throw "Installation verification failed"
        }
    } catch {
        Write-Host " FAIL" -ForegroundColor Red
        Write-Host "       Failed to install uv: $_" -ForegroundColor Red
        Write-Host "       Manual installation required:" -ForegroundColor Yellow
        Write-Host "       Visit: https://docs.astral.sh/uv/getting-started/installation/" -ForegroundColor Yellow
        $ErrorCount++
    }
}

# Test 4: Python Script
Write-Host "[4/10] Checking Python script..." -NoNewline
if (Test-Path ".\classify_and_report.py") {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " FAIL" -ForegroundColor Red
    Write-Host "       Missing: .\classify_and_report.py" -ForegroundColor Red
    $ErrorCount++
}

# Test 5: Python Dependencies
Write-Host "[5/10] Checking Python dependencies..." -NoNewline
Write-Host " OK (uv handles automatically)" -ForegroundColor Green

# Test 6: Configuration File
Write-Host "[6/10] Checking configuration file..." -NoNewline
if (Test-Path $ConfigFile) {
    try {
        $config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
        Write-Host " OK" -ForegroundColor Green
        
        # Validate config structure
        Write-Host "       Validating configuration structure..." -NoNewline
        $requiredFields = @("paths", "outputDirectory")
        $missingFields = @()
        
        foreach ($field in $requiredFields) {
            if (-not $config.PSObject.Properties[$field]) {
                $missingFields += $field
            }
        }
        
        if ($missingFields.Count -eq 0) {
            Write-Host " OK" -ForegroundColor Green
        } else {
            Write-Host " FAIL" -ForegroundColor Red
            Write-Host "       Missing fields: $($missingFields -join ', ')" -ForegroundColor Red
            $ErrorCount++
        }
    } catch {
        Write-Host " FAIL" -ForegroundColor Red
        Write-Host "       Invalid JSON format: $_" -ForegroundColor Red
        $ErrorCount++
        $config = $null
    }
} else {
    Write-Host " WARNING" -ForegroundColor Yellow
    Write-Host "       Configuration file not found: $ConfigFile" -ForegroundColor Yellow
    Write-Host "       Create from: config.sample.json" -ForegroundColor Yellow
    $WarningCount++
    $config = $null
}

# Test 7: Department Keywords File
Write-Host "[7/10] Checking department keywords..." -NoNewline

# Test 8: Output Directory
Write-Host "[8/10] Checking output directory..." -NoNewline
$outputDir = if ($config) { $config.outputDirectory } else { ".\output" }

if (Test-Path $outputDir) {
    Write-Host " OK" -ForegroundColor Green
} else {
    Write-Host " WARNING" -ForegroundColor Yellow
    Write-Host "       Output directory will be created: $outputDir" -ForegroundColor Yellow
    $WarningCount++
}

# Test 9: Scan Paths Accessibility
if ($config -and $config.paths) {
    Write-Host "[9/10] Checking scan paths accessibility..." -NoNewline
    $inaccessiblePaths = @()
    $validPaths = 0
    
    foreach ($path in $config.paths) {
        if (Test-Path $path) {
            $validPaths++
        } else {
            $inaccessiblePaths += $path
        }
    }
    
    if ($inaccessiblePaths.Count -eq 0) {
        Write-Host " OK ($validPaths paths)" -ForegroundColor Green
    } elseif ($validPaths -gt 0) {
        Write-Host " WARNING" -ForegroundColor Yellow
        Write-Host "       $($inaccessiblePaths.Count) path(s) not accessible:" -ForegroundColor Yellow
        foreach ($path in $inaccessiblePaths) {
            Write-Host "         - $path" -ForegroundColor Yellow
        }
        $WarningCount++
    } else {
        Write-Host " FAIL" -ForegroundColor Red
        Write-Host "       No paths are accessible" -ForegroundColor Red
        $ErrorCount++
    }
} else {
    Write-Host "[9/10] Checking scan paths accessibility..." -NoNewline
    Write-Host " SKIPPED (No config)" -ForegroundColor Gray
}

# Test 10: Disk Space
Write-Host "[10/10] Checking disk space..." -NoNewline
try {
    $drive = (Get-Item .).PSDrive
    $freeSpace = $drive.Free / 1GB
    
    if ($freeSpace -gt 5) {
        Write-Host " OK ($([math]::Round($freeSpace, 2)) GB free)" -ForegroundColor Green
    } elseif ($freeSpace -gt 1) {
        Write-Host " WARNING" -ForegroundColor Yellow
        Write-Host "       Low disk space: $([math]::Round($freeSpace, 2)) GB free" -ForegroundColor Yellow
        $WarningCount++
    } else {
        Write-Host " FAIL" -ForegroundColor Red
        Write-Host "       Insufficient disk space: $([math]::Round($freeSpace, 2)) GB free" -ForegroundColor Red
        $ErrorCount++
    }
} catch {
    Write-Host " WARNING (Could not check)" -ForegroundColor Yellow
    $WarningCount++
}

# Summary
Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
Write-Host "Validation Complete" -ForegroundColor Cyan
Write-Host ("=" * 70) -ForegroundColor Cyan

if ($ErrorCount -eq 0 -and $WarningCount -eq 0) {
    Write-Host "`n All checks passed! Ready to run analysis." -ForegroundColor Green
    Write-Host "`nTo start analysis, run:" -ForegroundColor White
    Write-Host "  .\Start-MigrationAnalysis.ps1" -ForegroundColor Cyan
    exit 0
} elseif ($ErrorCount -eq 0) {
    Write-Host "`n  $WarningCount warning(s) found. Review above and proceed with caution." -ForegroundColor Yellow
    Write-Host "`nYou can proceed, but some features may not work correctly." -ForegroundColor Yellow
    exit 0
} else {
    Write-Host "`n $ErrorCount error(s) found. Please fix the issues above before running." -ForegroundColor Red
    
    if ($WarningCount -gt 0) {
        Write-Host "  $WarningCount warning(s) also found." -ForegroundColor Yellow
    }
    

}

