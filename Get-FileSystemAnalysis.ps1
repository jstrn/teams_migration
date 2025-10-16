# Get-FileSystemAnalysis.ps1
# Core scanner module for SharePoint migration analysis

<#
.SYNOPSIS
    Scans file system paths and gathers detailed metadata for SharePoint migration analysis.

.DESCRIPTION
    Recursively scans file system paths, extracts NTFS/SMB permissions, identifies migration blockers,
    and streams results to CSV files for further processing.

.PARAMETER Path
    The root path to analyze

.PARAMETER Config
    Configuration object with thresholds and settings

.PARAMETER RawDataFile
    Output CSV file for raw scan data

.PARAMETER PermissionsFile
    Output CSV file for permissions data

.PARAMETER CheckpointFile
    Checkpoint file for restart capability

.PARAMETER Resume
    Resume from last checkpoint if available
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    
    [Parameter(Mandatory=$true)]
    [hashtable]$Config,
    
    [Parameter(Mandatory=$true)]
    [string]$RawDataFile,
    
    [Parameter(Mandatory=$true)]
    [string]$PermissionsFile,
    
    [Parameter(Mandatory=$true)]
    [string]$CheckpointFile,
    
    [Parameter(Mandatory=$false)]
    [switch]$Resume
)

# Initialize counters
$script:FilesProcessed = 0
$script:FoldersProcessed = 0
$script:StartTime = Get-Date
$script:LastCheckpointTime = Get-Date

# Function to check if an account should be excluded
function Test-ExcludedAccount {
    param([string]$AccountName)
    
    if ([string]::IsNullOrWhiteSpace($AccountName)) { return $true }
    
    # Exclude SYSTEM, Built-In accounts, and orphaned SIDs
    $excludePatterns = @(
        '^SYSTEM$',
        '^NT AUTHORITY\\',
        '^BUILTIN\\',
        '^S-1-5-\d+',
        '^NT SERVICE\\',
        '^APPLICATION PACKAGE AUTHORITY\\',
        '^Window Manager\\',
        '^Font Driver Host\\',
        '^ALL APPLICATION PACKAGES$',
        '^ALL RESTRICTED APPLICATION PACKAGES$'
    )
    
    foreach ($pattern in $excludePatterns) {
        if ($AccountName -match $pattern) {
            return $true
        }
    }
    
    return $false
}

# Function to convert FileSystemRights to Read or Read/Write
function ConvertTo-AccessLevel {
    param([System.Security.AccessControl.FileSystemRights]$Rights)
    
    $writeRights = [System.Security.AccessControl.FileSystemRights]::Write -bor 
                   [System.Security.AccessControl.FileSystemRights]::WriteData -bor
                   [System.Security.AccessControl.FileSystemRights]::AppendData -bor
                   [System.Security.AccessControl.FileSystemRights]::WriteExtendedAttributes -bor
                   [System.Security.AccessControl.FileSystemRights]::WriteAttributes -bor
                   [System.Security.AccessControl.FileSystemRights]::Delete -bor
                   [System.Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles -bor
                   [System.Security.AccessControl.FileSystemRights]::Modify -bor
                   [System.Security.AccessControl.FileSystemRights]::FullControl
    
    if (($Rights -band $writeRights) -ne 0) {
        return "Read/Write"
    } else {
        return "Read"
    }
}

# Function to get non-inherited permissions for a folder
function Get-FolderPermissions {
    param([string]$FolderPath)
    
    try {
        $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
        
        # Check if all permissions are inherited
        $hasExplicitPermissions = $false
        foreach ($access in $acl.Access) {
            if (-not $access.IsInherited) {
                $hasExplicitPermissions = $true
                break
            }
        }
        
        if (-not $hasExplicitPermissions) {
            return @()
        }
        
        # Extract non-inherited permissions
        $permissions = @()
        foreach ($access in $acl.Access) {
            if (-not $access.IsInherited) {
                $accountName = $access.IdentityReference.Value
                
                if (-not (Test-ExcludedAccount -AccountName $accountName)) {
                    $accessLevel = ConvertTo-AccessLevel -Rights $access.FileSystemRights
                    
                    $permissions += [PSCustomObject]@{
                        Path = $FolderPath
                        Account = $accountName
                        AccessLevel = $accessLevel
                        AccessControlType = $access.AccessControlType
                    }
                }
            }
        }
        
        return $permissions
    } catch {
        Write-Warning "Failed to get permissions for $FolderPath : $_"
        return @()
    }
}

# Function to check for unsupported characters
function Test-UnsupportedCharacters {
    param([string]$Name, [array]$UnsupportedChars)
    
    foreach ($char in $UnsupportedChars) {
        if ($Name.Contains($char)) {
            return $true
        }
    }
    return $false
}

# Function to count files in a folder
function Get-FolderFileCount {
    param([string]$FolderPath)
    
    try {
        $count = (Get-ChildItem -Path $FolderPath -File -Force -ErrorAction Stop).Count
        return $count
    } catch {
        return 0
    }
}

# Function to append to CSV file
function Export-ToCsv {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$Data,
        
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        
        [Parameter(Mandatory=$false)]
        [switch]$IsFirst
    )
    
    try {
        if ($IsFirst) {
            $Data | Export-Csv -Path $FilePath -NoTypeInformation -Force
        } else {
            $Data | Export-Csv -Path $FilePath -NoTypeInformation -Append
        }
    } catch {
        Write-Warning "Failed to write to CSV: $_"
    }
}

# Function to save checkpoint
function Save-Checkpoint {
    param(
        [string]$LastCompletedPath,
        [int]$FilesProcessed,
        [int]$FoldersProcessed
    )
    
    $checkpoint = @{
        lastCompletedPath = $LastCompletedPath
        filesProcessed = $FilesProcessed
        foldersProcessed = $FoldersProcessed
        startTime = $script:StartTime.ToString("o")
        lastUpdateTime = (Get-Date).ToString("o")
        resumeFlag = $true
    }
    
    $checkpoint | ConvertTo-Json | Set-Content -Path $CheckpointFile -Force
}

# Function to calculate ETA
function Get-ETA {
    param(
        [int]$ItemsProcessed,
        [int]$TotalItems
    )
    
    if ($ItemsProcessed -eq 0) { return "Calculating..." }
    
    $elapsed = (Get-Date) - $script:StartTime
    $rate = $ItemsProcessed / $elapsed.TotalSeconds
    
    if ($rate -eq 0) { return "Calculating..." }
    
    $remaining = $TotalItems - $ItemsProcessed
    $etaSeconds = $remaining / $rate
    $eta = [TimeSpan]::FromSeconds($etaSeconds)
    
    return "{0:hh\:mm\:ss}" -f $eta
}

# Main scanning function
function Invoke-FileSystemScan {
    Write-Host "Starting file system scan for: $Path" -ForegroundColor Cyan
    
    # Check if path exists
    if (-not (Test-Path -Path $Path)) {
        Write-Error "Path does not exist: $Path"
        return
    }
    
    # Initialize CSV files only if files don't exist
    if (-not (Test-Path $RawDataFile)) {
        $header = [PSCustomObject]@{
            Type = ""
            Path = ""
            Name = ""
            Extension = ""
            SizeBytes = ""
            Created = ""
            LastModified = ""
            PathLength = ""
            FileCountInFolder = ""
            HasUnsupportedChars = ""
            IsUnsafeExtension = ""
            IsLargeFile = ""
            IsTooManyFiles = ""
            IsTooLongPath = ""
            HasExplicitPermissions = ""
        }
        Export-ToCsv -Data $header -FilePath $RawDataFile -IsFirst
    }
    
    if (-not (Test-Path $PermissionsFile)) {
        $permHeader = [PSCustomObject]@{
            Path = ""
            Account = ""
            AccessLevel = ""
            AccessControlType = ""
        }
        Export-ToCsv -Data $permHeader -FilePath $PermissionsFile -IsFirst
    }
    
    # Get all items recursively
    Write-Host "Enumerating all items..." -ForegroundColor Yellow
    
    try {
        # Process folders first
        $folders = Get-ChildItem -Path $Path -Directory -Recurse -Force -ErrorAction SilentlyContinue
        $allFolders = @($Path) + $folders.FullName
        
        Write-Host "Found $($allFolders.Count) folders to process" -ForegroundColor Green
        
        $folderIndex = 0
        foreach ($folderPath in $allFolders) {
            $folderIndex++
            
            try {
                $folderItem = Get-Item -Path $folderPath -Force -ErrorAction Stop
                
                # Count files in this folder
                $fileCount = Get-FolderFileCount -Path $folderPath
                
                # Check for issues
                $pathLength = $folderPath.Length
                $hasUnsupportedChars = Test-UnsupportedCharacters -Name $folderItem.Name -UnsupportedChars $Config.unsupportedCharacters
                $isTooManyFiles = $fileCount -gt $Config.thresholds.maxFilesPerFolder
                $isTooLongPath = $pathLength -gt $Config.thresholds.maxPathLength
                
                # Get permissions
                $permissions = Get-FolderPermissions -FolderPath $folderPath
                $hasExplicitPermissions = $permissions.Count -gt 0
                
                # Write folder data
                $folderData = [PSCustomObject]@{
                    Type = "Folder"
                    Path = $folderPath
                    Name = $folderItem.Name
                    Extension = ""
                    SizeBytes = 0
                    Created = $folderItem.CreationTime.ToString("o")
                    LastModified = $folderItem.LastWriteTime.ToString("o")
                    PathLength = $pathLength
                    FileCountInFolder = $fileCount
                    HasUnsupportedChars = $hasUnsupportedChars
                    IsUnsafeExtension = $false
                    IsLargeFile = $false
                    IsTooManyFiles = $isTooManyFiles
                    IsTooLongPath = $isTooLongPath
                    HasExplicitPermissions = $hasExplicitPermissions
                }
                
                Export-ToCsv -Data $folderData -FilePath $RawDataFile
                
                # Write permissions
                foreach ($perm in $permissions) {
                    Export-ToCsv -Data $perm -FilePath $PermissionsFile
                }
                
                $script:FoldersProcessed++
                
                # Progress update every 100 folders
                if ($folderIndex % 100 -eq 0) {
                    $percentComplete = ($folderIndex / $allFolders.Count) * 100
                    $eta = Get-ETA -ItemsProcessed $folderIndex -TotalItems $allFolders.Count
                    Write-Progress -Activity "Processing Folders" `
                                   -Status "Folder $folderIndex of $($allFolders.Count) - ETA: $eta" `
                                   -PercentComplete $percentComplete
                }
                
                # Save checkpoint every 500 folders
                if ($folderIndex % 500 -eq 0) {
                    Save-Checkpoint -LastCompletedPath $folderPath `
                                   -FilesProcessed $script:FilesProcessed `
                                   -FoldersProcessed $script:FoldersProcessed
                }
                
            } catch {
                Write-Warning "Failed to process folder $folderPath : $_"
            }
        }
        
        Write-Progress -Activity "Processing Folders" -Completed
        
        # Process files (optional)
        $shouldScanFiles = $false
        if ($Config.PSObject.Properties['scanFiles']) { $shouldScanFiles = [bool]$Config.scanFiles }
        if ($shouldScanFiles) {
            Write-Host "Processing files..." -ForegroundColor Yellow
            $files = Get-ChildItem -Path $Path -File -Recurse -Force -ErrorAction SilentlyContinue
            
            Write-Host "Found $($files.Count) files to process" -ForegroundColor Green
            
            $fileIndex = 0
            foreach ($file in $files) {
                $fileIndex++
            
            try {
                $pathLength = $file.FullName.Length
                $hasUnsupportedChars = Test-UnsupportedCharacters -Name $file.Name -UnsupportedChars $Config.unsupportedCharacters
                $isUnsafeExtension = $Config.unsafeExtensions -contains $file.Extension.ToLower()
                $isLargeFile = $file.Length -gt $Config.thresholds.maxFileSize
                $isTooLongPath = $pathLength -gt $Config.thresholds.maxPathLength
                
                $fileData = [PSCustomObject]@{
                    Type = "File"
                    Path = $file.FullName
                    Name = $file.Name
                    Extension = $file.Extension
                    SizeBytes = $file.Length
                    Created = $file.CreationTime.ToString("o")
                    LastModified = $file.LastWriteTime.ToString("o")
                    PathLength = $pathLength
                    FileCountInFolder = 0
                    HasUnsupportedChars = $hasUnsupportedChars
                    IsUnsafeExtension = $isUnsafeExtension
                    IsLargeFile = $isLargeFile
                    IsTooManyFiles = $false
                    IsTooLongPath = $isTooLongPath
                    HasExplicitPermissions = $false
                }
                
                Export-ToCsv -Data $fileData -FilePath $RawDataFile
                
                $script:FilesProcessed++
                
                # Progress update every 1000 files
                if ($fileIndex % 1000 -eq 0) {
                    $percentComplete = ($fileIndex / $files.Count) * 100
                    $eta = Get-ETA -ItemsProcessed $fileIndex -TotalItems $files.Count
                    Write-Progress -Activity "Processing Files" `
                                   -Status "File $fileIndex of $($files.Count) - ETA: $eta" `
                                   -PercentComplete $percentComplete
                }
                
                # Save checkpoint every 5000 files
                if ($fileIndex % 5000 -eq 0) {
                    Save-Checkpoint -LastCompletedPath $file.FullName `
                                   -FilesProcessed $script:FilesProcessed `
                                   -FoldersProcessed $script:FoldersProcessed
                }
                
            } catch {
                Write-Warning "Failed to process file $($file.FullName) : $_"
            }
            }
        }
        else {
            Write-Host "Skipping file enumeration (scanFiles=false)" -ForegroundColor Yellow
        }
        
        Write-Progress -Activity "Processing Files" -Completed
        
    } catch {
        Write-Error "Error during scan: $_"
    }
    
    # Final checkpoint
    Save-Checkpoint -LastCompletedPath $Path `
                   -FilesProcessed $script:FilesProcessed `
                   -FoldersProcessed $script:FoldersProcessed
    
    Write-Host "`nScan completed!" -ForegroundColor Green
    Write-Host "Files processed: $($script:FilesProcessed)" -ForegroundColor Cyan
    Write-Host "Folders processed: $($script:FoldersProcessed)" -ForegroundColor Cyan
}

# Execute the scan
Invoke-FileSystemScan


