#!/usr/bin/env python3
"""
Excel Converter Script
Converts XLSM files to XLSX and removes password protections
"""

import os
import sys
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

def list_xlsm_files(src_dir):
    """List all XLSM files in the source directory"""
    xlsm_files = []
    if src_dir.exists():
        xlsm_files = list(src_dir.glob("*.xlsm"))
    return xlsm_files

def select_file():
    """Select XLSM file from src directory via command line"""
    src_dir = Path("src")
    if not src_dir.exists():
        src_dir.mkdir(exist_ok=True)
        print(f"Created 'src' directory. Please place your XLSM files there and run again.")
        return None
    
    xlsm_files = list_xlsm_files(src_dir)
    
    if not xlsm_files:
        print(f"No XLSM files found in '{src_dir}' directory.")
        print("Please place your XLSM files in the 'src' folder and run again.")
        return None
    
    print(f"\nFound {len(xlsm_files)} XLSM file(s) in '{src_dir}':")
    print("-" * 50)
    
    for i, file_path in enumerate(xlsm_files, 1):
        file_size = file_path.stat().st_size / 1024  # KB
        modified_time = datetime.fromtimestamp(file_path.stat().st_mtime)
        print(f"{i:2d}. {file_path.name}")
        print(f"     Size: {file_size:.1f} KB, Modified: {modified_time.strftime('%Y-%m-%d %H:%M')}")
    
    print("-" * 50)
    
    while True:
        try:
            choice = input(f"Select file (1-{len(xlsm_files)}) or 'q' to quit: ").strip().lower()
            
            if choice == 'q':
                return None
            
            file_index = int(choice) - 1
            if 0 <= file_index < len(xlsm_files):
                return str(xlsm_files[file_index])
            else:
                print(f"Please enter a number between 1 and {len(xlsm_files)}")
                
        except ValueError:
            print("Please enter a valid number or 'q' to quit")
        except KeyboardInterrupt:
            print("\nOperation cancelled.")
            return None

def remove_protection(worksheet):
    """Remove password protection from a worksheet"""
    try:
        # Remove sheet protection
        if worksheet.protection.sheet:
            worksheet.protection.sheet = False
        
        # Clear password hash if it exists
        if hasattr(worksheet.protection, 'password'):
            worksheet.protection.password = None
            
        # Disable protection attributes
        worksheet.protection.formatCells = False
        worksheet.protection.formatColumns = False
        worksheet.protection.formatRows = False
        worksheet.protection.insertColumns = False
        worksheet.protection.insertRows = False
        worksheet.protection.insertHyperlinks = False
        worksheet.protection.deleteColumns = False
        worksheet.protection.deleteRows = False
        worksheet.protection.selectLockedCells = False
        worksheet.protection.sort = False
        worksheet.protection.autoFilter = False
        worksheet.protection.pivotTables = False
        worksheet.protection.selectUnlockedCells = False
        
        return True
    except Exception as e:
        print(f"Warning: Could not fully remove protection from worksheet '{worksheet.title}': {e}")
        return False

def convert_xlsm_to_xlsx(input_path):
    """Convert XLSM to XLSX and remove password protections"""
    try:
        print(f"\nLoading workbook: {Path(input_path).name}")
        
        # Load the workbook
        workbook = load_workbook(input_path, keep_vba=False)
        
        # Remove workbook protection
        if hasattr(workbook, 'security') and workbook.security:
            workbook.security = None
            print("Removed workbook security settings")
            
        # Remove protection from all worksheets
        protected_sheets = []
        for worksheet in workbook.worksheets:
            if worksheet.protection.sheet:
                if remove_protection(worksheet):
                    protected_sheets.append(worksheet.title)
        
        if protected_sheets:
            print(f"Removed protection from {len(protected_sheets)} worksheet(s): {', '.join(protected_sheets)}")
        else:
            print("No protected worksheets found")
        
        # Create output directory with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = Path("converted") / timestamp
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Generate output filename
        input_file = Path(input_path)
        output_filename = input_file.stem + ".xlsx"
        output_path = output_dir / output_filename
        
        # Save as XLSX (this automatically removes VBA macros)
        print(f"Saving converted file: {output_path}")
        workbook.save(output_path)
        
        return str(output_path)
        
    except Exception as e:
        raise Exception(f"Error converting file: {str(e)}")

def main():
    """Main function"""
    print("Excel XLSM to XLSX Converter")
    print("=" * 50)
    print("This tool converts XLSM files to XLSX and removes password protections")
    print("=" * 50)
    
    try:
        # Select input file
        input_file = select_file()
        
        if not input_file:
            print("No file selected or no files available. Exiting...")
            return
        
        print(f"\nSelected file: {Path(input_file).name}")
        
        # Confirm conversion
        confirm = input("Proceed with conversion? (y/N): ").strip().lower()
        if confirm not in ['y', 'yes']:
            print("Conversion cancelled.")
            return
        
        # Convert the file
        output_file = convert_xlsm_to_xlsx(input_file)
        
        # Show success message
        print("\n" + "=" * 50)
        print("✓ CONVERSION COMPLETED SUCCESSFULLY!")
        print("=" * 50)
        print(f"Input:  {input_file}")
        print(f"Output: {output_file}")
        print(f"Size:   {Path(output_file).stat().st_size / 1024:.1f} KB")
        
        # Ask if user wants to open the output directory
        open_dir = input("\nOpen output directory? (y/N): ").strip().lower()
        if open_dir in ['y', 'yes']:
            output_dir = os.path.dirname(output_file)
            try:
                if sys.platform == "darwin":  # macOS
                    os.system(f"open '{output_dir}'")
                elif sys.platform == "win32":  # Windows
                    os.startfile(output_dir)
                else:  # Linux
                    os.system(f"xdg-open '{output_dir}'")
                print(f"Opened: {output_dir}")
            except Exception as e:
                print(f"Could not open directory: {e}")
                print(f"Manual path: {output_dir}")
        
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        print("\nTroubleshooting tips:")
        print("- Ensure the XLSM file is not open in Excel")
        print("- Check file permissions")
        print("- Verify the file is not corrupted")
    
    finally:
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()