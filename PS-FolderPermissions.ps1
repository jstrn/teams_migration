# File Server Folder Permissions Analysis Script @onderakoz
# Analyzes who has permissions on top-level folders on a domain member file server

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "File Server Analysis"
$form.Size = New-Object System.Drawing.Size(900, 650)
$form.StartPosition = "CenterScreen"

# Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(870, 540)
$form.Controls.Add($tabControl)

# Tab 1: Folder Permissions
$tabFolderPerms = New-Object System.Windows.Forms.TabPage
$tabFolderPerms.Text = "Folder Permissions"
$tabControl.Controls.Add($tabFolderPerms)

$labelPath = New-Object System.Windows.Forms.Label
$labelPath.Location = New-Object System.Drawing.Point(10, 10)
$labelPath.Size = New-Object System.Drawing.Size(100, 20)
$labelPath.Text = "Folder Path:"
$tabFolderPerms.Controls.Add($labelPath)

$textBoxPath = New-Object System.Windows.Forms.TextBox
$textBoxPath.Location = New-Object System.Drawing.Point(120, 10)
$textBoxPath.Size = New-Object System.Drawing.Size(450, 20)
$textBoxPath.Text = "C:\"
$tabFolderPerms.Controls.Add($textBoxPath)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Location = New-Object System.Drawing.Point(580, 8)
$buttonBrowse.Size = New-Object System.Drawing.Size(75, 25)
$buttonBrowse.Text = "Browse"
$tabFolderPerms.Controls.Add($buttonBrowse)

$buttonAnalyze = New-Object System.Windows.Forms.Button
$buttonAnalyze.Location = New-Object System.Drawing.Point(665, 8)
$buttonAnalyze.Size = New-Object System.Drawing.Size(75, 25)
$buttonAnalyze.Text = "Analyze"
$tabFolderPerms.Controls.Add($buttonAnalyze)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 45)
$dataGridView.Size = New-Object System.Drawing.Size(840, 400)
$dataGridView.AutoSizeColumnsMode = "Fill"
$dataGridView.ReadOnly = $true
$dataGridView.AllowUserToAddRows = $false
$tabFolderPerms.Controls.Add($dataGridView)

$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Location = New-Object System.Drawing.Point(10, 455)
$buttonExport.Size = New-Object System.Drawing.Size(100, 30)
$buttonExport.Text = "Export CSV"
$buttonExport.Enabled = $false
$tabFolderPerms.Controls.Add($buttonExport)

$buttonExportExcel = New-Object System.Windows.Forms.Button
$buttonExportExcel.Location = New-Object System.Drawing.Point(120, 455)
$buttonExportExcel.Size = New-Object System.Drawing.Size(100, 30)
$buttonExportExcel.Text = "Export Excel"
$buttonExportExcel.Enabled = $false
$tabFolderPerms.Controls.Add($buttonExportExcel)

# Tab 2: SMB Share Access
$tabSMBShare = New-Object System.Windows.Forms.TabPage
$tabSMBShare.Text = "SMB Share Access"
$tabControl.Controls.Add($tabSMBShare)

$buttonAnalyzeSMB = New-Object System.Windows.Forms.Button
$buttonAnalyzeSMB.Location = New-Object System.Drawing.Point(10, 8)
$buttonAnalyzeSMB.Size = New-Object System.Drawing.Size(150, 25)
$buttonAnalyzeSMB.Text = "Get SMB Share Access"
$tabSMBShare.Controls.Add($buttonAnalyzeSMB)

$dataGridViewSMB = New-Object System.Windows.Forms.DataGridView
$dataGridViewSMB.Location = New-Object System.Drawing.Point(10, 45)
$dataGridViewSMB.Size = New-Object System.Drawing.Size(840, 400)
$dataGridViewSMB.AutoSizeColumnsMode = "Fill"
$dataGridViewSMB.ReadOnly = $true
$dataGridViewSMB.AllowUserToAddRows = $false
$tabSMBShare.Controls.Add($dataGridViewSMB)

$buttonExportSMB = New-Object System.Windows.Forms.Button
$buttonExportSMB.Location = New-Object System.Drawing.Point(10, 455)
$buttonExportSMB.Size = New-Object System.Drawing.Size(100, 30)
$buttonExportSMB.Text = "Export CSV"
$buttonExportSMB.Enabled = $false
$tabSMBShare.Controls.Add($buttonExportSMB)

$buttonExportExcelSMB = New-Object System.Windows.Forms.Button
$buttonExportExcelSMB.Location = New-Object System.Drawing.Point(120, 455)
$buttonExportExcelSMB.Size = New-Object System.Drawing.Size(100, 30)
$buttonExportExcelSMB.Text = "Export Excel"
$buttonExportExcelSMB.Enabled = $false
$tabSMBShare.Controls.Add($buttonExportExcelSMB)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 560)
$progressBar.Size = New-Object System.Drawing.Size(870, 20)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 585)
$statusLabel.Size = New-Object System.Drawing.Size(870, 20)
$statusLabel.Text = "Ready"
$form.Controls.Add($statusLabel)

$global:results = @()
$global:smbResults = @()

function Convert-FileSystemRights {
    param([System.Security.AccessControl.FileSystemRights]$Rights)
    $rightsList = @()
    if ($Rights -band [System.Security.AccessControl.FileSystemRights]::FullControl) {
        $rightsList += "Full Control"
    } else {
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::Read) { $rightsList += "Read" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::Write) { $rightsList += "Write" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::Execute) { $rightsList += "Execute" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::Delete) { $rightsList += "Delete" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::Modify) { $rightsList += "Modify" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) { $rightsList += "Read and Execute" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::ListDirectory) { $rightsList += "List Directory" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::CreateFiles) { $rightsList += "Create Files" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::CreateDirectories) { $rightsList += "Create Folders" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::ReadPermissions) { $rightsList += "Read Permissions" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::ChangePermissions) { $rightsList += "Change Permissions" }
        if ($Rights -band [System.Security.AccessControl.FileSystemRights]::TakeOwnership) { $rightsList += "Take Ownership" }
    }
    return ($rightsList -join ", ")
}

function Convert-InheritanceFlags {
    param([System.Security.AccessControl.InheritanceFlags]$Inheritance)
    switch ($Inheritance) {
        "None" { return "Does Not Inherit" }
        "ContainerInherit" { return "Inherits to Subfolders" }
        "ObjectInherit" { return "Inherits to Files" }
        "ContainerInherit, ObjectInherit" { return "Inherits to All Child Objects" }
        default { return $Inheritance.ToString() }
    }
}

function Convert-PropagationFlags {
    param([System.Security.AccessControl.PropagationFlags]$Propagation)
    switch ($Propagation) {
        "None" { return "Unrestricted" }
        "NoPropagateInherit" { return "One Level Only" }
        "InheritOnly" { return "Inheritance Only" }
        default { return $Propagation.ToString() }
    }
}

function Get-FolderPermissions {
    param(
        [string]$Path,
        [int]$MaxDepth = 3
    )
    $results = [System.Collections.ArrayList]@()
    $statusLabel.Text = "Starting analysis..."
    $progressBar.Value = 0

    function Get-PermissionsRecursive {
        param(
            [string]$CurrentPath,
            [int]$CurrentDepth
        )
        try {
            if ($CurrentDepth -gt $MaxDepth) { return }
            $acl = Get-Acl -Path $CurrentPath
            foreach ($access in $acl.Access) {
                if ($access.AccessControlType -eq "Allow") {
                    $userInfo = $access.IdentityReference.ToString()
                    $isDomainUser = $userInfo -match "^[^\\]+\\[^\\]+$" -and $userInfo -notmatch "^BUILTIN\\\\" -and $userInfo -notmatch "^NT AUTHORITY\\\\"
                    $result = [PSCustomObject]@{
                        'Folder' = $CurrentPath
                        'User/Group' = $userInfo
                        'Permission' = Convert-FileSystemRights -Rights $access.FileSystemRights
                        'Permission Code' = $access.FileSystemRights.value__
                        'Inheritance' = Convert-InheritanceFlags -Inheritance $access.InheritanceFlags
                        'Propagation' = Convert-PropagationFlags -Propagation $access.PropagationFlags
                        'Domain User' = if ($isDomainUser) { "Yes" } else { "No" }
                        'Depth' = $CurrentDepth
                    }
                    $results.Add($result) | Out-Null
                }
            }
            if ($CurrentDepth -lt $MaxDepth) {
                $subfolders = Get-ChildItem -Path $CurrentPath -Directory -ErrorAction SilentlyContinue
                foreach ($subfolder in $subfolders) {
                    Get-PermissionsRecursive -CurrentPath $subfolder.FullName -CurrentDepth ($CurrentDepth + 1)
                }
            }
        }
        catch {
            Write-Host "Error: $CurrentPath - $_" -ForegroundColor Red
        }
    }

    Get-PermissionsRecursive -CurrentPath $Path -CurrentDepth 0
    return $results.ToArray()
}

$buttonBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the folder to analyze"
    $folderBrowser.SelectedPath = $textBoxPath.Text
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $textBoxPath.Text = $folderBrowser.SelectedPath
    }
})

$buttonAnalyze.Add_Click({
    $path = $textBoxPath.Text
    if (-not (Test-Path $path)) {
        [System.Windows.Forms.MessageBox]::Show("The specified folder was not found!", "Error", "OK", "Error")
        return
    }
    $statusLabel.Text = "Analyzing..."
    $progressBar.Style = "Marquee"
    $buttonAnalyze.Enabled = $false
    $global:results = Get-FolderPermissions -Path $path
    $dataGridView.DataSource = $global:results
    $statusLabel.Text = "Analysis complete. Found " + $global:results.Count + " permission entries."
    $progressBar.Style = "Continuous"
    $progressBar.Value = 100
    $buttonAnalyze.Enabled = $true
    $buttonExport.Enabled = $true
    $buttonExportExcel.Enabled = $true
})

$buttonExport.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "FileServer_Permissions_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".csv"
    if ($saveDialog.ShowDialog() -eq "OK") {
        try {
            $global:results | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("CSV file saved successfully!", "Success", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("CSV export error: $_", "Error", "OK", "Error")
        }
    }
})

$buttonExportExcel.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $saveDialog.FileName = "FileServer_Permissions_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".xlsx"
    if ($saveDialog.ShowDialog() -eq "OK") {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = "Folder Permissions"
            $headers = @("Folder", "User/Group", "Permission", "Permission Code", "Inheritance", "Propagation", "Domain User", "Depth")
            for ($i = 0; $i -lt $headers.Length; $i++) {
                $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
            }
            for ($row = 0; $row -lt $global:results.Count; $row++) {
                $worksheet.Cells.Item($row + 2, 1) = $global:results[$row].'Folder'
                $worksheet.Cells.Item($row + 2, 2) = $global:results[$row].'User/Group'
                $worksheet.Cells.Item($row + 2, 3) = $global:results[$row].'Permission'
                $worksheet.Cells.Item($row + 2, 4) = $global:results[$row].'Permission Code'
                $worksheet.Cells.Item($row + 2, 5) = $global:results[$row].'Inheritance'
                $worksheet.Cells.Item($row + 2, 6) = $global:results[$row].'Propagation'
                $worksheet.Cells.Item($row + 2, 7) = $global:results[$row].'Domain User'
                $worksheet.Cells.Item($row + 2, 8) = $global:results[$row].'Depth'
            }
            $worksheet.UsedRange.EntireColumn.AutoFit()
            $workbook.SaveAs($saveDialog.FileName)
            $workbook.Close()
            $excel.Quit()
            [System.Windows.Forms.MessageBox]::Show("Excel file saved successfully!", "Success", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Excel export error: $_", "Error", "OK", "Error")
        }
        finally {
            if ($excel) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            [System.GC]::Collect()
        }
    }
})

# SMB Share Access Button Click
$buttonAnalyzeSMB.Add_Click({
    try {
        $statusLabel.Text = "Analyzing SMB Shares..."
        $progressBar.Style = "Marquee"
        $buttonAnalyzeSMB.Enabled = $false
        
        $global:smbResults = Get-SmbShare | Where-Object {$_.Name -notlike "*$*"} | ForEach-Object {
            Get-SmbShareAccess -Name $_.Name | Select-Object @{Name='ShareName';Expression={$_.Name}}, AccountName, AccessControlType, AccessRight
        }
        
        $dataGridViewSMB.DataSource = $global:smbResults
        $statusLabel.Text = "SMB Share analysis complete. Found " + $global:smbResults.Count + " access entries."
        $progressBar.Style = "Continuous"
        $progressBar.Value = 100
        $buttonAnalyzeSMB.Enabled = $true
        $buttonExportSMB.Enabled = $true
        $buttonExportExcelSMB.Enabled = $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("SMB Share analysis error: $_`n`nNote: This requires administrative privileges.", "Error", "OK", "Error")
        $statusLabel.Text = "Error during SMB Share analysis"
        $progressBar.Style = "Continuous"
        $progressBar.Value = 0
        $buttonAnalyzeSMB.Enabled = $true
    }
})

# SMB Export CSV Button Click
$buttonExportSMB.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "SMB_ShareAccess_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".csv"
    if ($saveDialog.ShowDialog() -eq "OK") {
        try {
            $global:smbResults | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("CSV file saved successfully!", "Success", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("CSV export error: $_", "Error", "OK", "Error")
        }
    }
})

# SMB Export Excel Button Click
$buttonExportExcelSMB.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $saveDialog.FileName = "SMB_ShareAccess_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".xlsx"
    if ($saveDialog.ShowDialog() -eq "OK") {
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            $worksheet.Name = "SMB Share Access"
            $headers = @("ShareName", "AccountName", "AccessControlType", "AccessRight")
            for ($i = 0; $i -lt $headers.Length; $i++) {
                $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
            }
            for ($row = 0; $row -lt $global:smbResults.Count; $row++) {
                $worksheet.Cells.Item($row + 2, 1) = $global:smbResults[$row].'ShareName'
                $worksheet.Cells.Item($row + 2, 2) = $global:smbResults[$row].'AccountName'
                $worksheet.Cells.Item($row + 2, 3) = $global:smbResults[$row].'AccessControlType'
                $worksheet.Cells.Item($row + 2, 4) = $global:smbResults[$row].'AccessRight'
            }
            $worksheet.UsedRange.EntireColumn.AutoFit()
            $workbook.SaveAs($saveDialog.FileName)
            $workbook.Close()
            $excel.Quit()
            [System.Windows.Forms.MessageBox]::Show("Excel file saved successfully!", "Success", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Excel export error: $_", "Error", "OK", "Error")
        }
        finally {
            if ($excel) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            [System.GC]::Collect()
        }
    }
})

$form.ShowDialog()
