# Master Budget Automation Tool v1.0.2

This app imports values from the Expense Sub-Program export into the Master Budget workbook without relying on OFFSET/MATCH formulas.

## How to download the Expense Sub-Program Mastersheet from Compass
1. Open Compass and go to **Financial Management**.
2. In the **Financial Management Dashboard**, select the budget year you want in **Selected Financial Period**.
3. Click **Financial Periods**.
4. On the **Financial Periods** page, find the same budget year.
5. Open the **Budget Reports** drop-down for that year and select **Expense Sub-Program Mastersheet**.
6. Compass will download the file. Save it, then use that downloaded file as the **Expense Sub-Program file** in this app.

The screenshots for these steps are available from the **Instructions** button inside the app.

## What it does
- imports data directly into the **Master** sheet
- refreshes **Compass** if that sheet exists
- preserves macros and button bindings on Windows when Microsoft Excel is installed
- keeps blank source values blank
- keeps the original template unchanged by saving to a new output workbook
- highlights mismatch items
- inserts source-only account codes and source-only sub-program codes into **Master** in numeric position

## Before you start
1. Close the template workbook and output workbook in Excel.
2. Keep the original template as a separate file.
3. Save the output workbook as a new file name.

## How to run
1. Double-click `run_budget_tool.bat`.
2. Click **Browse** next to **Expense Sub-Program file** and choose the source file.
3. Click **Browse** next to **Master Budget template** and choose the original template workbook.
4. Click **Browse** next to **Output workbook** and choose a new file name and folder.
5. Click **Generate budget workbook**.
6. Wait for the progress bar to finish.
7. Open the output workbook in Excel and review the imported data.

## App buttons
- **Generate budget workbook**: runs the import
- **Create suggested output name**: creates a suggested output file name
- **Open output folder**: opens the folder where the output workbook will be saved
- **Instructions**: opens the in-app user guide with screenshots
- **Clear**: clears the selected file paths and run summary

## Build and release
- Run `build_windows_exe.bat` to produce a Windows EXE using the PyInstaller spec file.
- Run `build_store_installer.ps1` to produce a silent-installable Windows installer for Microsoft Store EXE submission.
- Microsoft Store submission notes and listing templates are in the `store` folder.

## Suggestions
Please send suggestions to contact@vurctne.com
