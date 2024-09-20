# Recon-Automation

# Recon Automation Script v1.0

## Overview

This script automates a series of tasks related to file management, data extraction, and Excel workbook manipulations. It utilizes libraries like `openpyxl` for Excel handling and `win32com` (Windows-only) for specific COM interactions. The script is primarily designed to manage reconciliations using a combination of ZIP file extraction and data processing in Excel.

## Features

- **Archive Extraction**: Unzip files and extract their contents.
- **Excel Workbook Manipulation**: Use `openpyxl` to manipulate Excel files, including formatting cells, creating tables, and applying formulas.
- **Automated Data Handling**: Uses `pandas` for data frame manipulations that are then written into Excel sheets.
- **Cross-Compatibility**: The script is designed for Windows environments, as it requires the `win32com` library for certain tasks.

## Requirements

- Python 3.x
- Libraries:
  - `openpyxl`: For reading and writing Excel files.
  - `pandas`: For data frame handling.
  - `win32com`: Windows-specific library for Excel COM automation.
  - `shutil`: For high-level file operations (included with Python).
  - `re`: For regular expressions (included with Python).
  - `subprocess`: For running system commands (included with Python).
  - `zipfile`: For handling ZIP archives (included with Python).

### Installation

To install the required Python libraries, run:

```bash
pip install openpyxl pandas pywin32
```

Note: `pywin32` (which includes `win32com`) is only available for Windows.

## Usage

1. **Extract Archives**:
   - The script can handle ZIP file extraction. Provide the input folder containing the ZIP files and an output folder where the extracted files will be stored.

   Example:
   ```python
   extract_archive('input.zip', 'output_folder')
   ```

2. **Excel Workbook Manipulations**:
   - The script processes Excel files by applying styles, creating tables, and adding formulas. Ensure your source Excel file is properly structured for this.

3. **Running the Script**:
   Run the Python script via the command line:
   ```bash
   python Recon_Automation_Code_V1.0.py
   ```

## Cross-Platform Considerations

While the script heavily relies on the `win32com` library for certain Windows-specific Excel automation tasks, you can modify the script to be cross-platform by replacing COM-based functionality with `openpyxl` or other platform-independent libraries if necessary.

### Example:

- Modify `win32com` Excel automation to use `openpyxl` for reading and writing Excel files on macOS or Linux.

## Troubleshooting

- If you encounter a `ModuleNotFoundError` for `win32com`, ensure that the script is running on a Windows system with the `pywin32` package installed.
- For ZIP extraction issues, check the validity of the archive file and ensure proper paths are provided.

---
