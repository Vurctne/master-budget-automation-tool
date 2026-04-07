# Master Budget Automation Tool v1.0.2

This release packages the current desktop budgeting utility for GitHub distribution and Microsoft Store submission preparation.

## Included assets

- `MasterBudgetAutomationTool_Setup_1.0.2.exe`
  Recommended Windows installer built with Inno Setup.
- `Master Budget Automation Tool v1.0.2.exe`
  Standalone PyInstaller build for manual distribution and testing.

## Highlights

- Imports Compass Expense Sub-Program data into the Master Budget workbook
- Supports CSV, XLSX, and XLSM source files
- Prevents unsafe output selections that would overwrite the template or source workbook
- Detects duplicate account and sub-program codes in source data
- Improves CSV import compatibility with common Windows encodings
- Adds Store-ready packaging scripts and submission metadata

## Validation

- `python -m py_compile app.py budget_automation.py app_metadata.py`
- `powershell -ExecutionPolicy Bypass -File .\build_store_installer.ps1 -SkipDependencyInstall`

## Notes

- The installer is not code-signed yet.
- For Microsoft Store submission, the final installer must be code-signed and hosted at a public, versioned HTTPS URL.
- If this repository remains private, GitHub Release asset links are not suitable as the Microsoft Store package URL.
