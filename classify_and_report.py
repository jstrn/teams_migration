#!/usr/bin/env python3
"""
classify_and_report.py
Department classification and Excel report generation for SharePoint migration analysis.
"""
# /// script
# dependencies = ["pandas", "openpyxl", "requests", "python-dotenv"]
# ///

import argparse
import json
import csv
import re
import os
import requests
from pathlib import Path
from collections import defaultdict, Counter
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from dotenv import load_dotenv

load_dotenv()


# Department list for AI classification
DEPARTMENTS = [
    "Executive", "Finance", "Accounting", "Human Resources", "Payroll", "Legal", 
    "Administration", "Operations", "Facilities", "Information Technology", "Security",
    "Sales", "Marketing", "Customer Service", "Product Management", "Project Management",
    "Procurement", "Supply Chain", "Inventory", "Logistics", "Engineering", 
    "Research and Development", "Quality Assurance", "Training", "Business Development",
    "Vendor Management", "Compliance"
]


class DepartmentClassifier:
    """Implements pure AI-based classification for department folders."""
    
    def __init__(self, use_ai=True, openai_api_key=None):
        """
        Initialize classifier with AI-based classification.
        
        Args:
            use_ai: Whether to use OpenAI API for classification
            openai_api_key: OpenAI API key (if None, uses OPENAI_API_KEY environment variable)
        """
        self.use_ai = use_ai
        self.openai_api_key = openai_api_key or os.getenv('OPENAI_API_KEY')
        
        if self.use_ai and not self.openai_api_key:
            print("Error: OpenAI API key not found.")
            print("Please set OPENAI_API_KEY environment variable or use --openai-key parameter.")
            raise ValueError("OpenAI API key is required for AI-based classification.")
    
    
    def get_top_level_folders(self, df):
        """Extract top-level folders from the dataset."""
        folders_df = df[df['Type'] == 'Folder'].copy()
        top_level_folders = {}
        
        for idx, row in folders_df.iterrows():
            path = row['Path']
            path_parts = path.split(os.sep)
            
            # Get the first non-root part as top-level folder
            if len(path_parts) > 1:
                top_level_folder = path_parts[1]  # Skip the drive letter
                full_path = path_parts[0] + os.sep + top_level_folder
                
                if full_path not in top_level_folders:
                    top_level_folders[full_path] = {
                        'name': top_level_folder,
                        'path': full_path,
                        'samples': []
                    }
        
        return top_level_folders
    
    def get_folder_samples(self, df, top_level_path, sample_size=25):
        """Get sample file paths from a top-level folder."""
        # Get files and folders under this top-level path
        relevant_items = df[
            (df['Path'].str.startswith(top_level_path + os.sep)) & 
            (df['Path'] != top_level_path)  # Exclude the top-level folder itself
        ]
        
        # Sample files and folders
        files = relevant_items[relevant_items['Type'] == 'File']
        folders = relevant_items[relevant_items['Type'] == 'Folder']
        
        # Combine and sample
        all_items = pd.concat([files, folders])
        if len(all_items) > sample_size:
            all_items = all_items.sample(n=sample_size, random_state=42)
        
        return all_items['Path'].tolist()
    
    def classify_with_openai(self, folder_name, file_samples, model="gpt-4o-mini"):
        """Classify a folder using OpenAI API."""
        try:
            # Prepare the prompt
            samples_text = '\n'.join(file_samples[:25])  # Limit to 25 samples
            departments = ', '.join(DEPARTMENTS)
            
            prompt = f"""You are a file system analyst helping to classify department folders for a SharePoint migration.

Available departments: {departments}

Folder name: {folder_name}

Sample file paths under this folder:
{samples_text}

Based on the folder name and the file path patterns, classify this folder into the most appropriate department. Consider:
1. The folder name itself
2. The types of files and subfolders present
3. The naming patterns and content indicators
4. The overall context of the file structure

Respond with ONLY the department name (one of: {departments}) or "Unknown" if none clearly fit."""

            # Call OpenAI API
            headers = {
                "Authorization": f"Bearer {self.openai_api_key}",
                "Content-Type": "application/json"
            }
            
            data = {
                "model": model,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 50,
                "temperature": 0.1
            }
            
            response = requests.post(
                "https://api.openai.com/v1/chat/completions",
                headers=headers,
                json=data,
                timeout=30
            )
            response.raise_for_status()
            
            result = response.json()
            classification = result['choices'][0]['message']['content'].strip()
            
            # Validate the response
            if classification in DEPARTMENTS:
                return classification
            else:
                print(f"OpenAI returned invalid classification: {classification}")
                return "Unknown"
                
        except Exception as e:
            print(f"OpenAI API call failed: {e}")
            return "Unknown"
    
    def get_immediate_children(self, df, parent_path):
        """Get immediate children (files and folders) of a parent path."""
        children = []
        parent_path_normalized = parent_path.rstrip(os.sep) + os.sep
        
        for idx, row in df.iterrows():
            path = row['Path']
            if path.startswith(parent_path_normalized):
                # Get the relative path after the parent
                relative_path = path[len(parent_path_normalized):]
                
                # Check if this is an immediate child (only one path component)
                if relative_path and os.sep not in relative_path:
                    children.append(path)
        
        return children
    
    def classify_hierarchy(self, df):
        """
        Implement pure AI-based classification for top-level folders and their immediate children.
        
        Algorithm:
        1. Identify top-level folders and classify them using AI
        2. For each top-level folder, classify its immediate children using AI
        3. Propagate classifications downward to all files and subfolders
        """
        print("\nStarting AI-based hierarchical classification...")
        
        # Prepare data structures
        folders_df = df[df['Type'] == 'Folder'].copy()
        files_df = df[df['Type'] == 'File'].copy()
        
        # Dictionary to store folder classifications and scores
        folder_classifications = {}
        folder_confidence = {}
        
        if not self.use_ai:
            print("Error: AI classification is required but not available.")
            print("Please set OPENAI_API_KEY environment variable or use --openai-key parameter.")
            raise ValueError("AI classification is required but OpenAI API key is not available.")
        
        # Step 1: Classify top-level folders using AI
        print("Classifying top-level folders using AI...")
        top_level_folders = self.get_top_level_folders(df)
        
        for folder_path, folder_info in top_level_folders.items():
            folder_name = folder_info['name']
            print(f"  Classifying top-level: {folder_name}")
            
            # Get sample file paths from the entire folder tree
            samples = self.get_folder_samples(df, folder_path, sample_size=25)
            print(f"    Found {len(samples)} sample paths")
            
            # Classify using AI
            classification = self.classify_with_openai(folder_name, samples)
            print(f"    AI Classification: {classification}")
            
            folder_classifications[folder_path] = classification
            folder_confidence[folder_path] = 1.0  # High confidence for AI classification
        
        # Step 2: Classify immediate children of top-level folders using AI
        print("Classifying immediate children of top-level folders using AI...")
        
        for top_level_path in top_level_folders.keys():
            print(f"  Processing children of: {top_level_path}")
            
            # Get immediate children (files and folders)
            immediate_children = self.get_immediate_children(df, top_level_path)
            
            for child_path in immediate_children:
                # Skip if already classified
                if child_path in folder_classifications:
                    continue
                
                # Get child info
                child_row = df[df['Path'] == child_path].iloc[0]
                child_name = child_row['Name']
                child_type = child_row['Type']
                
                # Only classify folders, skip files
                if child_type != 'Folder':
                    continue
                
                print(f"    Classifying {child_type}: {child_name}")
                
                # Get samples from this child and its descendants
                child_samples = self.get_folder_samples(df, child_path, sample_size=15)
                print(f"      Found {len(child_samples)} sample paths")
                
                # Classify using AI
                classification = self.classify_with_openai(child_name, child_samples)
                print(f"      AI Classification: {classification}")
                
                folder_classifications[child_path] = classification
                folder_confidence[child_path] = 0.8  # Slightly lower confidence for children
        
        # Step 3: No inheritance - we only classify top-level and immediate children
        print("AI-based folder classification complete!")
        
        # Return empty file classifications since we only classify folders
        file_classifications = {}
        
        return folder_classifications, folder_confidence, file_classifications
    


class ExcelReportGenerator:
    """Generate formatted Excel workbook with multiple sheets."""
    
    def __init__(self, output_file):
        self.output_file = output_file
        self.workbook = Workbook()
        
        # Remove default sheet
        if 'Sheet' in self.workbook.sheetnames:
            del self.workbook['Sheet']
        
        # Define styles
        self.header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        self.header_font = Font(bold=True, color="FFFFFF", size=11)
        self.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    
    def apply_header_style(self, ws, row=1):
        """Apply header styling to the first row."""
        for cell in ws[row]:
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = self.border
    
    def auto_fit_columns(self, ws, min_width=10, max_width=50):
        """Auto-fit column widths based on content."""
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            
            adjusted_width = min(max(max_length + 2, min_width), max_width)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    def create_summary_sheet(self, df, folder_classifications, file_classifications):
        """Create summary sheet with high-level statistics."""
        ws = self.workbook.create_sheet("Summary", 0)
        
        # Calculate statistics
        total_files = len(df[df['Type'] == 'File'])
        total_folders = len(df[df['Type'] == 'Folder'])
        total_size = df[df['Type'] == 'File']['SizeBytes'].sum()
        
        # Convert size to human-readable format
        def format_size(bytes_val):
            for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
                if bytes_val < 1024:
                    return f"{bytes_val:.2f} {unit}"
                bytes_val /= 1024
            return f"{bytes_val:.2f} PB"
        
        # Count issues
        issues = {
            'Long Paths (>255 chars)': df['IsTooLongPath'].sum(),
            'Large Files (>50MB)': df['IsLargeFile'].sum(),
            'Folders with >20k files': df['IsTooManyFiles'].sum(),
            'Unsupported Characters': df['HasUnsupportedChars'].sum(),
            'Unsafe File Extensions': df['IsUnsafeExtension'].sum(),
            'Explicit Permissions': df['HasExplicitPermissions'].sum()
        }
        
        # Count by department (folders only)
        dept_counts = Counter(folder_classifications.values())
        
        # Write summary data
        row = 1
        ws.cell(row, 1, "SharePoint Migration Analysis - Summary Report")
        ws.cell(row, 1).font = Font(bold=True, size=14)
        row += 2
        
        ws.cell(row, 1, "Generated:")
        ws.cell(row, 2, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        row += 2
        
        # Overall statistics
        ws.cell(row, 1, "Overall Statistics")
        ws.cell(row, 1).font = Font(bold=True, size=12)
        row += 1
        
        ws.cell(row, 1, "Total Files:")
        ws.cell(row, 2, total_files)
        row += 1
        
        ws.cell(row, 1, "Total Folders:")
        ws.cell(row, 2, total_folders)
        row += 1
        
        ws.cell(row, 1, "Total Size:")
        ws.cell(row, 2, format_size(total_size))
        row += 2
        
        # Issues summary
        ws.cell(row, 1, "Migration Issues Found")
        ws.cell(row, 1).font = Font(bold=True, size=12)
        row += 1
        
        for issue_name, count in issues.items():
            ws.cell(row, 1, issue_name)
            ws.cell(row, 2, count)
            if count > 0:
                ws.cell(row, 2).font = Font(color="FF0000", bold=True)
            row += 1
        
        row += 1
        
        # Department breakdown
        ws.cell(row, 1, "Department Classification")
        ws.cell(row, 1).font = Font(bold=True, size=12)
        row += 1
        
        ws.cell(row, 1, "Department")
        ws.cell(row, 2, "File Count")
        ws.cell(row, 1).font = Font(bold=True)
        ws.cell(row, 2).font = Font(bold=True)
        row += 1
        
        for dept, count in sorted(dept_counts.items(), key=lambda x: x[1], reverse=True):
            ws.cell(row, 1, dept)
            ws.cell(row, 2, count)
            row += 1
        
        # Auto-fit columns
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        
        return ws
    
    def create_files_sheet(self, df, file_classifications):
        """Create detailed files sheet."""
        files_df = df[df['Type'] == 'File'].copy()
        
        # Files don't have direct classifications, but we can show parent folder info
        files_df['ParentFolder'] = files_df['Path'].apply(lambda x: str(Path(x).parent))
        
        # Select and order columns (removed Department since files aren't classified)
        columns = [
            'Path', 'Name', 'Extension', 'SizeBytes', 'Created', 'LastModified',
            'ParentFolder', 'PathLength', 'HasUnsupportedChars', 'IsUnsafeExtension',
            'IsLargeFile', 'IsTooLongPath'
        ]
        
        files_df = files_df[columns]
        
        # Create sheet
        ws = self.workbook.create_sheet("Files")
        
        # Write data
        for r_idx, row in enumerate(dataframe_to_rows(files_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(r_idx, c_idx, value)
        
        # Apply header style
        self.apply_header_style(ws)
        
        # Auto-fit columns
        self.auto_fit_columns(ws)
        
        # Freeze first row
        ws.freeze_panes = 'A2'
        
        return ws
    
    def create_folders_sheet(self, df, folder_classifications, folder_confidence):
        """Create folders sheet."""
        folders_df = df[df['Type'] == 'Folder'].copy()
        
        # Add classification and confidence
        folders_df['Department'] = folders_df['Path'].map(folder_classifications)
        folders_df['Confidence'] = folders_df['Path'].map(folder_confidence)
        folders_df['Confidence'] = folders_df['Confidence'].fillna(0).round(2)
        
        # Select and order columns
        columns = [
            'Path', 'Name', 'Department', 'Confidence', 'FileCountInFolder',
            'Created', 'LastModified', 'PathLength', 'HasUnsupportedChars',
            'IsTooManyFiles', 'IsTooLongPath', 'HasExplicitPermissions'
        ]
        
        folders_df = folders_df[columns]
        
        # Create sheet
        ws = self.workbook.create_sheet("Folders")
        
        # Write data
        for r_idx, row in enumerate(dataframe_to_rows(folders_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(r_idx, c_idx, value)
        
        # Apply header style
        self.apply_header_style(ws)
        
        # Auto-fit columns
        self.auto_fit_columns(ws)
        
        # Freeze first row
        ws.freeze_panes = 'A2'
        
        return ws
    
    def create_permissions_sheet(self, permissions_df):
        """Create permissions sheet."""
        ws = self.workbook.create_sheet("Permissions")
        
        # Filter to only Allow access
        permissions_df = permissions_df[permissions_df['AccessControlType'] == 'Allow'].copy()
        
        # Select columns
        columns = ['Path', 'Account', 'AccessLevel']
        permissions_df = permissions_df[columns]
        
        # Write data
        for r_idx, row in enumerate(dataframe_to_rows(permissions_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(r_idx, c_idx, value)
        
        # Apply header style
        self.apply_header_style(ws)
        
        # Auto-fit columns
        self.auto_fit_columns(ws)
        
        # Freeze first row
        ws.freeze_panes = 'A2'
        
        return ws
    
    def create_issues_sheet(self, df, file_classifications, folder_classifications):
        """Create consolidated issues sheet."""
        issues_list = []
        
        # Process each row
        for idx, row in df.iterrows():
            item_path = row['Path']
            item_type = row['Type']
            item_name = row['Name']
            
            # Get classification (only folders are classified)
            if item_type == 'File':
                # For files, get department from parent folder
                parent_folder = str(Path(item_path).parent)
                dept = "Unclassified"
                
                # Find the closest parent folder that has a classification
                current_path = parent_folder
                while current_path:
                    if current_path in folder_classifications:
                        dept = folder_classifications[current_path]
                        break
                    parent = str(Path(current_path).parent)
                    if parent == current_path:  # Reached root
                        break
                    current_path = parent
            else:
                dept = folder_classifications.get(item_path, 'Unclassified')
            
            # Check for issues
            if row['IsTooLongPath']:
                issues_list.append({
                    'Issue Type': 'Long Path',
                    'Path': item_path,
                    'Name': item_name,
                    'Type': item_type,
                    'Department': dept,
                    'Details': f"Path length: {row['PathLength']} chars (>255)"
                })
            
            if row['IsLargeFile']:
                size_mb = row['SizeBytes'] / (1024 * 1024)
                issues_list.append({
                    'Issue Type': 'Large File',
                    'Path': item_path,
                    'Name': item_name,
                    'Type': item_type,
                    'Department': dept,
                    'Details': f"File size: {size_mb:.2f} MB (>50MB)"
                })
            
            if row['IsTooManyFiles']:
                issues_list.append({
                    'Issue Type': 'Too Many Files',
                    'Path': item_path,
                    'Name': item_name,
                    'Type': item_type,
                    'Department': dept,
                    'Details': f"File count: {row['FileCountInFolder']} (>20,000)"
                })
            
            if row['HasUnsupportedChars']:
                issues_list.append({
                    'Issue Type': 'Unsupported Characters',
                    'Path': item_path,
                    'Name': item_name,
                    'Type': item_type,
                    'Department': dept,
                    'Details': 'Name contains unsupported characters'
                })
            
            if row['IsUnsafeExtension']:
                issues_list.append({
                    'Issue Type': 'Unsafe Extension',
                    'Path': item_path,
                    'Name': item_name,
                    'Type': item_type,
                    'Department': dept,
                    'Details': f"Extension: {row['Extension']}"
                })
        
        # Convert to DataFrame
        if issues_list:
            issues_df = pd.DataFrame(issues_list)
        else:
            # Create empty DataFrame with columns
            issues_df = pd.DataFrame(columns=[
                'Issue Type', 'Path', 'Name', 'Type', 'Department', 'Details'
            ])
        
        # Create sheet
        ws = self.workbook.create_sheet("Issues")
        
        # Write data
        for r_idx, row in enumerate(dataframe_to_rows(issues_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(r_idx, c_idx, value)
        
        # Apply header style
        self.apply_header_style(ws)
        
        # Auto-fit columns
        self.auto_fit_columns(ws)
        
        # Freeze first row
        ws.freeze_panes = 'A2'
        
        return ws
    
    def create_classification_sheet(self, df, folder_classifications, folder_confidence, 
                                    file_classifications):
        """Create department classification details sheet (folders only)."""
        classification_list = []
        
        # Only process folders that were actually classified by AI
        for path, department in folder_classifications.items():
            # Find the folder row in the dataframe
            folder_row = df[df['Path'] == path]
            if not folder_row.empty:
                row = folder_row.iloc[0]
            classification_list.append({
                'Type': 'Folder',
                'Path': path,
                'Name': row['Name'],
                    'Department': department,
                'Confidence Score': folder_confidence.get(path, 0.0)
            })
        
        classification_df = pd.DataFrame(classification_list)
        
        # Create sheet
        ws = self.workbook.create_sheet("Classification")
        
        # Write data
        for r_idx, row in enumerate(dataframe_to_rows(classification_df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(r_idx, c_idx, value)
        
        # Apply header style
        self.apply_header_style(ws)
        
        # Auto-fit columns
        self.auto_fit_columns(ws)
        
        # Freeze first row
        ws.freeze_panes = 'A2'
        
        return ws
    
    def save(self):
        """Save the workbook."""
        self.workbook.save(self.output_file)
        print(f"\nExcel report saved to: {self.output_file}")


def main():
    """Main execution function."""
    parser = argparse.ArgumentParser(
        description='Classify and generate Excel report for SharePoint migration analysis'
    )
    parser.add_argument('--config', required=True, help='Configuration JSON file')
    parser.add_argument('--raw-data', required=True, help='Raw scan data CSV file')
    parser.add_argument('--permissions', required=True, help='Permissions CSV file')
    parser.add_argument('--output', required=True, help='Output Excel file')
    parser.add_argument('--use-ai', action='store_true', help='Use OpenAI API for classification (required)')
    parser.add_argument('--openai-key', required=False, help='OpenAI API key (if not provided, uses OPENAI_API_KEY environment variable)')
    
    args = parser.parse_args()
    
    print("=" * 70)
    print("SharePoint Migration Analysis - Classification & Reporting")
    print("=" * 70)
    
    # Check if input files exist
    if not Path(args.raw_data).exists():
        print(f"Error: Raw data file not found: {args.raw_data}")
        return 1
    
    if not Path(args.permissions).exists():
        print(f"Error: Permissions file not found: {args.permissions}")
        return 1
    
    # Load data
    print("\nLoading scan data...")
    df = pd.read_csv(args.raw_data)
    print(f"Loaded {len(df)} items")
    
    # Convert boolean columns
    bool_columns = [
        'HasUnsupportedChars', 'IsUnsafeExtension', 'IsLargeFile',
        'IsTooManyFiles', 'IsTooLongPath', 'HasExplicitPermissions'
    ]
    for col in bool_columns:
        df[col] = df[col].astype(bool)
    
    # Load permissions
    print("\nLoading permissions data...")
    permissions_df = pd.read_csv(args.permissions)
    print(f"Loaded {len(permissions_df)} permission entries")
    
    # Initialize classifier with AI support
    classifier = DepartmentClassifier(
        use_ai=args.use_ai,
        openai_api_key=args.openai_key
    )
    
    # Perform classification
    folder_classifications, folder_confidence, file_classifications = \
        classifier.classify_hierarchy(df)
    
    # Generate Excel report
    print("\nGenerating Excel report...")
    report = ExcelReportGenerator(args.output)
    
    report.create_summary_sheet(df, folder_classifications, file_classifications)
    print("  - Summary sheet created")
    
    report.create_files_sheet(df, file_classifications)
    print("  - Files sheet created")
    
    report.create_folders_sheet(df, folder_classifications, folder_confidence)
    print("  - Folders sheet created")
    
    report.create_permissions_sheet(permissions_df)
    print("  - Permissions sheet created")
    
    report.create_issues_sheet(df, file_classifications, folder_classifications)
    print("  - Issues sheet created")
    
    report.create_classification_sheet(df, folder_classifications, folder_confidence,
                                       file_classifications)
    print("  - Classification sheet created")
    
    # Save workbook
    report.save()
    
    print("\n" + "=" * 70)
    print("Report generation complete!")
    print("=" * 70)
    
    return 0


if __name__ == '__main__':
    exit(main())


