<# 
.SYNOPSIS
  Classify folders by department based on path keywords.

.PARAMETER RootPath
  Root folder to scan.

.PARAMETER KeywordsCsv
  CSV with columns: Department, Keywords
  Keywords is a comma-separated list per department.

.PARAMETER MaxDepth
  Maximum depth relative to RootPath to scan. Default 3.

.PARAMETER IncludeFiles
  If set, also classify files. Default is folders only.

.PARAMETER OutputCsv
  Where to save the results CSV. Default next to script with timestamp.

.OUTPUTS
  CSV with columns:
    Path, RelativeDepth, CandidateDepartment, Score, MatchedKeywords, Ties, ScoresJson
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$RootPath,

  [string]$KeywordsCsv = "$(Split-Path -Parent $PSCommandPath)\Department_Keywords.csv",

  [int]$MaxDepth = 3,

  [switch]$IncludeFiles,

  [string]$OutputCsv
)

# Validate root
if (-not (Test-Path -LiteralPath $RootPath)) {
  throw "RootPath not found: $RootPath"
}

# Load keywords
if (-not (Test-Path -LiteralPath $KeywordsCsv)) {
  throw "Keywords CSV not found: $KeywordsCsv"
}

$raw = Import-Csv -LiteralPath $KeywordsCsv
if (-not ($raw -and $raw[0].PSObject.Properties.Name -contains 'Department' -and $raw[0].PSObject.Properties.Name -contains 'Keywords')) {
  throw "Keywords CSV must have columns Department and Keywords"
}

# Build in-memory map: Department -> tokens[]
$DeptMap = @{}
foreach ($row in $raw) {
  $dept = ($row.Department).Trim()
  if ([string]::IsNullOrWhiteSpace($dept)) { continue }
  $tokens = @()
  foreach ($k in ($row.Keywords -split ',')) {
    $t = ($k).Trim()
    if (-not [string]::IsNullOrWhiteSpace($t)) { $tokens += $t }
  }
  if ($tokens.Count -gt 0) { $DeptMap[$dept] = $tokens }
}

if ($DeptMap.Count -eq 0) {
  throw "No department tokens loaded from $KeywordsCsv"
}

# Build regex patterns for each token
function New-TokenRegex {
  param([string]$Token)
  # Escape regex metachars
  $escaped = [Regex]::Escape($Token)
  # Allow separators where underscores appear
  $escaped = $escaped -replace '\\_', '[\\/_\-\.\s]+'
  # Make wordish boundaries using common path separators
  $sep = '[\\/_\-\.\s]'
  $pattern = "(?i)(^|$sep)$escaped($|$sep)"
  return $pattern
}

# Precompute token regex per department
$DeptRegex = @{}
foreach ($dept in $DeptMap.Keys) {
  $DeptRegex[$dept] = $DeptMap[$dept] | ForEach-Object { 
    [PSCustomObject]@{ Token = $_; Pattern = New-TokenRegex -Token $_ }
  }
}

# Helper to compute relative depth from RootPath
function Get-RelativeDepth {
  param([string]$Path, [string]$Root)
  $p = (Resolve-Path -LiteralPath $Path).Path
  $r = (Resolve-Path -LiteralPath $Root).Path
  $pSegs = $p.TrimEnd('\','/').Split('\','/')
  $rSegs = $r.TrimEnd('\','/').Split('\','/')
  return [Math]::Max(0, $pSegs.Count - $rSegs.Count)
}

# Enumerate items up to MaxDepth
$rootResolved = (Resolve-Path -LiteralPath $RootPath).Path
$items = @()

# Include the root itself as depth 0
$items += Get-Item -LiteralPath $rootResolved

# Recurse and prune by relative depth
$enum = Get-ChildItem -LiteralPath $rootResolved -Recurse -Force -ErrorAction SilentlyContinue
foreach ($it in $enum) {
  if (-not $IncludeFiles -and -not $it.PSIsContainer) { continue }
  $d = Get-RelativeDepth -Path $it.FullName -Root $rootResolved
  if ($d -le $MaxDepth) { $items += $it }
}

# Scoring function
function Score-Path {
  param([string]$PathLower)

  $scores = [ordered]@{}
  $matches = [ordered]@{}
  foreach ($dept in $DeptRegex.Keys) {
    $deptScore = 0
    $deptHits = @()
    foreach ($tr in $DeptRegex[$dept]) {
      $pat = $tr.Pattern
      # Exact segment hit gets 2 points, substring context 1 point
      if ($PathLower -match $pat) {
        $deptScore += 2
        $deptHits += $tr.Token
      } else {
        # secondary relaxed check: pure substring, case-insensitive
        if ($PathLower -like "*$($tr.Token.ToLower())*") {
          $deptScore += 1
          $deptHits += $tr.Token
        }
      }
    }
    $scores[$dept] = $deptScore
    $matches[$dept] = ($deptHits | Select-Object -Unique)
  }

  # Determine top scores
  $max = ($scores.Values | Measure-Object -Maximum).Maximum
  $top = @()
  foreach ($kv in $scores.GetEnumerator()) {
    if ($kv.Value -eq $max -and $max -gt 0) { $top += $kv.Key }
  }

  # Build output
  return [PSCustomObject]@{
    Scores = $scores
    TopDepartments = $top
    TopScore = $max
    Matched = $matches
  }
}

# Classify each item
$results = foreach ($it in $items | Sort-Object FullName -Unique) {
  try {
    $relDepth = Get-RelativeDepth -Path $it.FullName -Root $rootResolved
    $pathLower = $it.FullName.ToLowerInvariant()
    $sc = Score-Path -PathLower $pathLower

    $cand = if ($sc.TopDepartments.Count -gt 0) { $sc.TopDepartments[0] } else { "" }
    $ties = if ($sc.TopDepartments.Count -gt 1) { ($sc.TopDepartments -join '; ') } else { "" }

    # Build matched keywords string for candidate dept
    $matchedForCand = ""
    if ($cand -ne "") {
      $matchedForCand = ($sc.Matched[$cand] -join '; ')
    }

    # Serialize all scores as compact JSON
    $scoresJson = ($sc.Scores.GetEnumerator() | Sort-Object Name | ForEach-Object {
      @{ Department = $_.Key; Score = $_.Value }
    } | ConvertTo-Json -Compress)

    [PSCustomObject]@{
      Path = $it.FullName
      RelativeDepth = $relDepth
      CandidateDepartment = $cand
      Score = $sc.TopScore
      MatchedKeywords = $matchedForCand
      Ties = $ties
      ScoresJson = $scoresJson
      ItemType = if ($it.PSIsContainer) { "Folder" } else { "File" }
    }
  } catch {
    [PSCustomObject]@{
      Path = $it.FullName
      RelativeDepth = $null
      CandidateDepartment = ""
      Score = 0
      MatchedKeywords = ""
      Ties = ""
      ScoresJson = "{}"
      ItemType = if ($it.PSIsContainer) { "Folder" } else { "File" }
      Error = $_.Exception.Message
    }
  }
}

# Output file path
if (-not $OutputCsv) {
  $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
  $OutputCsv = Join-Path -Path (Split-Path -Parent $PSCommandPath) -ChildPath "Path_Department_Classification_$stamp.csv"
}

$results | Sort-Object Path | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "Classification complete. Output: $OutputCsv" -ForegroundColor Green
