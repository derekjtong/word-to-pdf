# Word to PDF

A PowerShell script that batch-converts Word documents (`.doc`, `.docx`, `.docm`) to PDF using Microsoft Word's COM automation. Includes an optional Windows Explorer right-click context menu integration.

## Requirements

- Windows with Microsoft Word installed
- PowerShell 5.1+ (Windows PowerShell)

## Usage

```powershell
.\ExportWordDocsToPdf.ps1 [-SourcePath <path>] [-Recurse] [-Overwrite] [-ShowWord] [-UpdateFields] [-NoPause]
```

### Parameters

| Parameter       | Description                                                          |
| --------------- | ---------------------------------------------------------------------|
| `-SourcePath`    | Folder to scan for Word documents. Defaults to the script's folder. |
| `-Recurse`       | Include subfolders when searching for documents.                    |
| `-Overwrite`     | Overwrite existing PDFs. Without this, existing PDFs are skipped.   |
| `-UpdateFields`  | Update fields and tables of contents before exporting.              |
| `-ShowWord`      | Show the Word application window during conversion.                 |
| `-NoPause`       | Don't wait for a keypress before closing the window.                |

### Example

Convert all Word documents in a folder and its subfolders, overwriting any existing PDFs:

```powershell
.\ExportWordDocsToPdf.ps1 -SourcePath "C:\Documents\Reports" -Recurse -Overwrite
```

## Explorer context menu (optional)

`ExportWordDocsToPdf.reg` adds an **"Export Word Docs to PDF"** option to the right-click menu when you right-click a folder or its background in File Explorer.

1. Copy `ExportWordDocsToPdf.ps1` to `C:\Scripts\ExportWordDocsToPdf.ps1` (or edit the `.reg` file to point at a different location).
2. Double-click `ExportWordDocsToPdf.reg` and confirm the registry import.
3. Right-click any folder (or the background of a folder) and choose **Export Word Docs to PDF**.

To remove the menu entry, delete these registry keys:

```
HKEY_CLASSES_ROOT\Directory\Background\shell\ExportWordPDF
HKEY_CLASSES_ROOT\Directory\shell\ExportWordPDF
```

## Notes

- Files that already have a non-empty PDF counterpart are skipped unless `-Overwrite` is passed.
- Temporary Word lock files (starting with `~$`) are ignored.
- A summary of converted, skipped, and failed files is printed when the script finishes.
