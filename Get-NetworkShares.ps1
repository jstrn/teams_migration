# Get all SMB shares on the local machine, excluding administrative shares
$hostname = $env:COMPUTERNAME
$lines = @()

Get-SmbShare | Where-Object { -not $_.Special -and $_.Name -ne 'CertEnroll' -and $_.Name -ne 'print$' } | ForEach-Object {
	$uncPath = "\\$hostname\$($_.Name)".tolower()
	$lines += "$uncPath, [$($_.Path)]"
}

# Convert to plaintext (newline-separated: "UNC<TAB>Path")
$sharesText = ($lines -join "`n")

# Output the results
Write-Output $sharesText
Ninja-Property-Set networkShares $sharesText
