# SharePoint Migration Analysis Tool

A comprehensive tool for analyzing file shares before migrating to SharePoint Online/Microsoft Teams. This tool scans network file shares, identifies migration blockers, analyzes permissions, and classifies content by department using an intelligent consensus algorithm.

## Features

- **Comprehensive File System Scanning**
  - Recursive scanning of multiple network paths
  - NTFS and SMB permission extraction
  - File and folder metadata collection
  - Progress tracking with ETA estimation

- **Migration Blocker Detection**
  - Path length validation (>255 characters)
  - Large file identification (>50MB)
  - Folder file count limits (>20,000 files)
  - Unsupported character detection
  - Unsafe file extension flagging

- **Intelligent Department Classification**
  - Bottom-up consensus algorithm
  - Keyword-based classification using configurable keywords
  - Confidence scoring
  - Hierarchical classification propagation

- **Detailed Excel Reporting**
  - Summary statistics and issue counts
  - Detailed file and folder listings
  - Permission analysis (Read vs Read/Write)
  - Consolidated issues view
  - Department classification details

- **Checkpoint & Resume Capability**
  - Automatic progress checkpointing
  - Resume from last completed path
  - Handles long-running scans gracefully

## Requirements

### PowerShell
- PowerShell 5.1 or later (PowerShell 7+ recommended)
- Windows operating system
- Read access to target file shares
- Permissions to query NTFS/SMB ACLs

### uv (Python Package Manager)
- **uv** - Modern, fast Python package manager that handles everything automatically
- No need to install Python or manage dependencies manually!
- Required packages (installed automatically by uv):
  - pandas >= 1.5.0
  - openpyxl >= 3.0.0

## Installation

1. Clone or download this repository to your local machine

2. Install uv (if not already installed):
   ```powershell
   # Windows PowerShell
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   
   # Or visit: https://docs.astral.sh/uv/getting-started/installation/
   ```
   
   **That's it!** uv will automatically install Python and all dependencies when you run the tool.

3. (Optional) Customize department keywords by creating a Department_Keywords.csv file
   - The tool includes 27 pre-configured departments with embedded keywords
   - Only needed if you want to override the default classification rules

## Configuration

Edit `config.json` to specify your scan parameters:

```json
{
  "paths": [],
  "outputDirectory": ".\\output",
  "thresholds": {
    "maxPathLength": 255,
    "maxFileSize": 52428800,
    "maxFilesPerFolder": 20000
  }
}
```

### Configuration Options

- **paths**: Array of UNC paths or local paths to scan
  - **Leave empty `[]` to auto-discover local SMB shares** (excludes administrative/hidden shares)
  - Or specify paths manually: `["\\\\server\\share", "C:\\SharedFiles"]`
- **outputDirectory**: Where to save output files
- **rawDataFile**: CSV file for raw scan results
- **permissionsFile**: CSV file for permissions data
- **checkpointFile**: JSON file for checkpoint/resume
- **excelOutputFile**: Final Excel report location
- **departmentKeywordsFile**: CSV file with department keywords
- **thresholds**: Validation thresholds for migration blockers
- **unsafeExtensions**: File extensions that cannot sync
- **unsupportedCharacters**: Characters not allowed in SharePoint

## Usage

### Basic Usage

Run the complete analysis:

```powershell
.\Start-MigrationAnalysis.ps1
```

### Resume from Checkpoint

If the scan was interrupted, resume from the last checkpoint:

```powershell
.\Start-MigrationAnalysis.ps1 -Resume
```

### Skip Scan and Regenerate Report

If you already have scan data and want to regenerate the report:

```powershell
.\Start-MigrationAnalysis.ps1 -SkipScan
```

### Custom Configuration File

Use a different configuration file:

```powershell
.\Start-MigrationAnalysis.ps1 -ConfigFile ".\my-config.json"
```

## Output Files

The tool generates several output files in the configured output directory:

### 1. migration_analysis.xlsx
Comprehensive Excel workbook with multiple sheets:

- **Summary**: High-level statistics, issue counts, department breakdown
- **Files**: Detailed file listing with metadata and flags
- **Folders**: Folder analysis with file counts and classifications
- **Permissions**: Non-inherited permissions (Read vs Read/Write)
- **Issues**: Consolidated view of all migration blockers
- **Classification**: Department assignments with confidence scores

### 2. raw_scan_data.csv
Streaming CSV output from the PowerShell scanner containing all raw data.

### 3. permissions_data.csv
Extracted permissions data filtered to relevant accounts and access levels.

### 4. checkpoint.json
Progress checkpoint for resume capability.

## Department Classification

The tool includes **27 pre-configured departments with embedded keywords**, making it ready to use out of the box. The tool uses a sophisticated bottom-up consensus algorithm to classify folders and files by department:

1. **Keyword Matching**: Each file and folder name is tokenized and matched against department keywords
2. **Scoring**: Matches are scored with weights:
   - Folder names: 5.0x weight
   - Subfolder classifications: 3.0x weight
   - File names: 1.0x weight
3. **Propagation**: Classifications propagate from deepest folders upward
4. **Consensus**: Parent folders inherit classification based on aggregate scores from children

### Customizing Department Keywords (Optional)

The tool includes embedded keywords for 27 common departments. To customize classifications for your organization, create a `Department_Keywords.csv` file and reference it in your config:

**Embedded Departments (27 total):**
Executive, Finance, Accounting, Human Resources, Payroll, Legal, Administration, Operations, Facilities, Information Technology, Security, Sales, Marketing, Customer Service, Product Management, Project Management, Procurement, Supply Chain, Inventory, Logistics, Engineering, Research and Development, Quality Assurance, Training, Business Development, Vendor Management, Compliance

**To customize, create Department_Keywords.csv:**
```csv
Department,Keywords
Finance,"finance, financials, budgets, forecasting"
Legal,"legal, contracts, agreements, nda, litigation"
```

Then add to your config.json:
```json
{
  "departmentKeywordsFile": ".\\Department_Keywords.csv",
  ...
}
```

**Keyword Guidelines:**
- Use comma-separated keywords
- Underscores in keywords are converted to spaces
- Multi-word keywords receive bonus scoring
- Keywords are case-insensitive

## Permission Filtering

The tool filters permissions to show only relevant information for migration planning:

### Excluded Accounts
- SYSTEM
- NT AUTHORITY\* accounts
- BUILTIN\* groups
- Orphaned SIDs (S-1-5-*)
- Service accounts

### Permission Simplification
- Only **Read** vs **Read/Write** access is reported
- Only **Allow** permissions are shown (Deny rules filtered out)
- Only folders with **explicit (non-inherited)** permissions are listed

## Performance Considerations

- **Large File Shares**: The tool handles millions of files through streaming and checkpointing
- **Progress Tracking**: Real-time progress with ETA estimation
- **Memory Efficiency**: Results streamed to CSV files during scanning
- **Checkpoint Frequency**: 
  - Every 500 folders processed
  - Every 5,000 files processed

## Troubleshooting

### Python Not Found
Ensure Python is installed and in your PATH. The script will search for `python`, `python3`, or `py` commands.

### Missing Python Packages
The script automatically attempts to install missing packages. If this fails, manually install:
```powershell
pip install pandas openpyxl
```

### Access Denied Errors
Ensure you have:
- Read access to all target paths
- Permission to query ACLs on folders
- Administrative privileges if scanning system folders

### Large Scans Taking Too Long
- Use the checkpoint/resume feature for multi-day scans
- Consider breaking very large shares into multiple config paths
- Run during off-hours to avoid network congestion

## Migration Blockers Reference

### Path Length (>255 characters)
SharePoint Online has a 400-character limit, but the Windows client has a 255-character limit. Files exceeding this may not sync properly.

**Remediation**: Shorten folder/file names or restructure hierarchy.

### Large Files (>50MB)
Files over 50MB may have slow sync performance. SharePoint supports up to 250GB per file, but large files require special handling.

**Remediation**: Consider alternative storage (OneDrive, Azure Files) or chunked uploads.

### Too Many Files (>20,000 per folder)
Folders with over 20,000 items can cause sync and performance issues.

**Remediation**: Restructure folders to distribute files more evenly.

### Unsupported Characters
SharePoint doesn't allow: ~ # % & * { } \ : < > ? / | "

**Remediation**: Rename files/folders to remove these characters.

### Unsafe Extensions
Database and application files may not sync properly or could cause conflicts.

**Remediation**: Exclude from sync or migrate to appropriate storage.

## License

This tool is provided as-is for migration planning purposes.

## Support

For issues or questions, please review the generated Excel report's Issues sheet for specific migration blockers and remediation guidance.


