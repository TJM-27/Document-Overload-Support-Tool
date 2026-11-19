# DOST Zip Document GUI

`DOST-ZipDocumentGui.ps1` is a Windows PowerShell 5.1 WinForms tool for reviewing and exporting documents from a zip file or an already-extracted folder. `DOST.exe` is a small launcher for the GUI.

## What It Does

- Prompts you to select a source `.zip` or source folder.
- Extracts the source zip, or copies the source folder, into `work\<timestamp>_<source-name>`.
- Recursively extracts any nested `.zip` files it finds.
- Groups all discovered files into tabs by type. `.xls` and `.xlsx` files are grouped together in the `EXCEL` tab.
- Shows the export parent folder and file name in each tab's scrollable table, sorted by parent folder.
- Exports a tab as copied files, preserving the parent folder grouping such as `TD`, `BMO`, etc.
- Exports the `EXCEL` tab as one combined `.xlsx` workbook per parent folder.
- Exports the `PDF` tab as one merged bookmarked `.pdf` per parent folder.
- Shows export progress with an estimated time remaining while exports are running.

## Launch

Run:

```powershell
.\DOST.exe
```

Or open Windows PowerShell 5.1 and run:

```powershell
Set-Location "C:\Users\Forensic Workstation\Documents\Cursor Workspaces\DOST"
powershell.exe -ExecutionPolicy Bypass -File ".\DOST-ZipDocumentGui.ps1"
```

PowerShell 5.1 is recommended because the app uses WinForms and Microsoft Excel COM automation.

## Portable Use

The tool can be copied to a USB stick and run on another Windows computer if PowerShell and Microsoft Excel are installed. Keep `DOST.exe` and `DOST-ZipDocumentGui.ps1` in the same folder.

PDF merging uses a local iTextSharp-compatible library under `lib\`. If `lib\` is not present, DOST will try to download that library from NuGet the first time `Export Merged PDF...` is used. On an offline or restricted computer, that download may fail.

For reliable USB/offline use, run a PDF merge once on this computer so `lib\` is populated, then copy the entire `DOST` folder to the USB stick, including `lib\`.

To rebuild `DOST.exe` after changing the launcher, run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Build-DOSTExecutable.ps1"
```

## Export Options

### Export Files

Available on every tab.

Choose a destination folder. The tool creates a type folder and then parent-folder subdirectories.

Example for the `TXT` tab:

```text
TXT\
  TD\
    notes.txt
  BMO\
    memo.txt
```

Duplicate output filenames are handled by adding a numeric suffix.

### Export Combined Excel

Available on the `EXCEL` tab.

Requirements:

- Microsoft Excel must be installed.
- Input files may be `.xls` or `.xlsx`.

For each parent folder, the tool creates one `.xlsx` workbook. Each visible worksheet from each source workbook is copied into the combined workbook. If a source workbook has multiple visible sheets, the new worksheet names use the source filename plus the original sheet name. Invalid Excel worksheet characters are removed and names are shortened to Excel's 31-character limit.

Example:

```text
EXCEL_Combined\
  TD.xlsx
  BMO.xlsx
```

### Export Merged PDF

Available on the `PDF` tab.

For each parent folder, the tool creates one merged PDF. Each source PDF starts at a bookmark named from the source filename.

Example:

```text
PDF_Merged\
  TD.pdf
  BMO.pdf
```

On first use, the script downloads a local iTextSharp-compatible NuGet package into `lib\` so it can merge PDFs and add bookmarks.

## Parent Folder Rule

The tool walks down through single wrapper folders until it finds the first folder that contains more than one child folder. Those child folders usually become the parent groups.

This handles zips or folders shaped like this:

```text
Luke\
  TD\
    statement.pdf
  BMO\
    statement.pdf
```

In that case, `TD` and `BMO` are used as the export parent folders, not `Luke`.

If the first folder with multiple child folders also directly contains `.pdf`, `.xls`, or `.xlsx` files, DOST treats that folder as its own parent group by stepping one level up for sorting. This handles folders shaped like this:

```text
Luke\
  RBC\
    Account Recap.xls
    Transactions.pdf
    etransfers\
      E-transfer Report.xlsx
    statements\
      Statement.pdf
```

In that case, `RBC` is used as the export parent folder for both the direct files and the files under `etransfers` and `statements`.

It also handles zips shaped like this:

```text
TD\statement.pdf
BMO\statement.pdf
```

Files outside the detected parent grouping root are grouped under `_Unsorted`.

## Logs

Every export writes an export log beside the exported files. Logs include copied files, generated outputs, skipped files, and errors from Excel or PDF processing.

## Work Folder Cleanup

The `work\` folder only stores temporary extracted zip contents or copied source-folder contents. Export logs are written beside exported files in the destination you choose, not in `work\`.

DOST automatically clears old `work\` contents before loading a new zip or folder and when the app closes.

## Zip Import Warnings

If the selected source zip is not a readable zip archive, the import stops and shows a clear error. This usually means the file is incomplete, corrupted, a split archive, or not actually a zip file.

If an internal `.zip` file cannot be read, the import continues and reports a warning. Successfully expanded internal zip files are not shown as separate files in the ZIP tab; only zip files that remain unexpanded are listed.

## Sample Zip

Use `New-DOSTSampleZip.ps1` to create a small test zip with `TD`, `BMO`, nested zip content, text files, PNG placeholders, sample PDFs, and CSV-based spreadsheet placeholders.

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\New-DOSTSampleZip.ps1"
```

The generated sample is written under `samples\`.

## Optional Smoke Tests

The GUI is the normal workflow, but the main script also includes non-GUI validation switches for testing.

Inventory only:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\DOST-ZipDocumentGui.ps1" -InventoryOnlyZip ".\samples\DOST_Sample_yyyyMMdd_HHmmss.zip"
```

Folder inventory only:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\DOST-ZipDocumentGui.ps1" -InventoryOnlyFolder ".\samples\DOST_Sample_yyyyMMdd_HHmmss\payload"
```

Inventory plus raw file export:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\DOST-ZipDocumentGui.ps1" -InventoryOnlyZip ".\samples\DOST_Sample_yyyyMMdd_HHmmss.zip" -SmokeExportRoot ".\samples\smoke-export"
```

Inventory, raw export, and merged PDF export:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\DOST-ZipDocumentGui.ps1" -InventoryOnlyZip ".\samples\DOST_Sample_yyyyMMdd_HHmmss.zip" -SmokeExportRoot ".\samples\smoke-export-pdf" -SmokeMergePdf
```

Inventory, raw export, and combined Excel export:

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\DOST-ZipDocumentGui.ps1" -InventoryOnlyZip ".\samples\DOST_Sample_yyyyMMdd_HHmmss.zip" -SmokeExportRoot ".\samples\smoke-export-excel" -SmokeCombineExcel
```
