# MC Reconciliation Tool

## Repository Information

GitHub Repository: https://github.com/fujiwen/MC_RECON_WITH_ARTICLE_SUMMARY

## Project Description

This is a tool for processing receipt records and generating supplier reconciliation details. The tool is developed using Python and uses PyQt5 to build the graphical user interface.

## Requirements

- Python 3.10 or higher
- Dependencies: pandas, numpy, openpyxl, PyQt5

## Running Without Virtual Environment

### Method 1: Using Batch File

1. Double-click `run_without_venv.bat` file
2. The batch file will automatically check Python environment and dependencies, then start the program

### Method 2: Manual Run

1. Ensure Python 3.10 or higher is installed
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Run the program:
   ```
   python MC_Recon_UI.py
   ```

## Building Executable

To build a standalone executable, use the following command:

```
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --icon=favicon.ico --name="MC_Recon_Tool" MC_Recon_UI.py
```

After building, the executable will be located in the `dist` directory.

## Automatic Build

This project has configured GitHub Actions workflow. When code is pushed to the main branch, it will automatically build Windows executable. The build results can be downloaded from GitHub Actions artifacts.