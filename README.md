# Excel XLSM to XLSX Converter

A Python script that converts Excel macro-enabled files (XLSM) to standard Excel format (XLSX) while removing all password protections and security restrictions.

## Features

- 🔄 Convert XLSM files to XLSX format
- 🔓 Remove all worksheet password protections
- 🛡️ Remove workbook security settings
- 🧹 Strip VBA macros automatically
- 📁 Organized output with timestamp directories
- 💻 Cross-platform support (Windows, macOS, Linux)
- 🎯 Interactive command-line interface

## Requirements

- Python 3.6 or higher
- openpyxl library

## Installation

1. Clone or download this repository
2. Install the required dependency:
```bash
pip install openpyxl
```

## Project Structure

```
xlsm_to_xlsx/
├── convert.py          # Main conversion script
├── src/               # Place your XLSM files here
├── converted/         # Output directory (created automatically)
│   └── YYYYMMDD_HHMMSS/  # Timestamped subdirectories
└── README.md
```

## Usage

### Quick Start

1. Place your XLSM files in the `src` directory
2. Run the converter:
```bash
python3 convert.py
```
3. Select a file from the numbered list
4. Confirm the conversion
5. Find your converted file in `converted/YYYYMMDD_HHMMSS/`

### Step-by-Step Example

```bash
$ python3 convert.py

Excel XLSM to XLSX Converter
==================================================
This tool converts XLSM files to XLSX and removes password protections
==================================================

Found 2 XLSM file(s) in 'src':
--------------------------------------------------
 1. financial_report.xlsm
     Size: 245.3 KB, Modified: 2025-05-31 14:22
 2. protected_workbook.xlsm
     Size: 156.7 KB, Modified: 2025-05-30 09:15
--------------------------------------------------
Select file (1-2) or 'q' to quit: 1

Selected file: financial_report.xlsm
Proceed with conversion? (y/N): y

Loading workbook: financial_report.xlsm
Removed protection from 3 worksheet(s): Sheet1, Data, Summary
Saving converted file: converted/20250531_142856/financial_report.xlsx

==================================================
✓ CONVERSION COMPLETED SUCCESSFULLY!
==================================================
Input:  src/financial_report.xlsm
Output: converted/20250531_142856/financial_report.xlsx
Size:   201.4 KB

Open output directory? (y/N): y
Opened: converted/20250531_142856

Press Enter to exit...
```

## What Gets Removed

### Security & Protection
- Sheet password protection
- Workbook security settings
- Cell formatting restrictions
- Row/column insertion/deletion restrictions
- Sorting and filtering restrictions
- Hyperlink insertion restrictions
- Pivot table restrictions

### Content Changes
- VBA macros (automatically removed during XLSM → XLSX conversion)
- Macro security settings
- Protected view restrictions

## Output Structure

Each conversion creates a timestamped directory:
```
converted/
├── 20250531_142856/
│   └── your_file.xlsx
├── 20250531_143012/
│   └── another_file.xlsx
└── ...
```

This prevents accidental overwrites and keeps a history of conversions.

## Troubleshooting

### Common Issues

**"No XLSM files found"**
- Ensure files are placed in the `src` directory
- Check file extensions are `.xlsm` (not `.xlsx` or `.xls`)

**"Permission denied" errors**
- Close the XLSM file in Excel before conversion
- Check file/folder permissions
- Run as administrator if needed (Windows)

**"File is corrupted" errors**
- Try opening the file in Excel first to verify it's valid
- Check if the file requires a password to open (not supported)
- Ensure the file isn't damaged

### File Requirements

- Files must be in XLSM format
- Files should not be password-protected for opening (worksheet protection is fine)
- Files must not be open in Excel during conversion

## Limitations

- Cannot handle files that require a password to open
- Does not preserve VBA macros (by design)
- Large files (>100MB) may take longer to process
- Some advanced Excel features may not be preserved

## Security Note

This tool removes security protections from Excel files. Only use it on files you own or have permission to modify. The converted files will have no password protection.

## License

This project is provided as-is for educational and legitimate business purposes. Use responsibly and in accordance with your organization's policies.

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

---

**Version:** 1.0  
**Last Updated:** May 2025