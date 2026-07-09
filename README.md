# Word to PDF

A PowerShell script that batch-converts Word documents (`.doc`, `.docx`, `.docm`) to PDF using Microsoft Word's COM automation. Includes an optional Windows Explorer right-click context menu integration.

## Requirements

- Windows with Microsoft Word installed
- PowerShell 5.1+ (Windows PowerShell)

## Usage

```powershell
.\ExportWordDocsToPdf.ps1 [-SourcePath <path>] [-Recurse] [-Overwrite] [-ShowWord] [-NoPause]
```

### Parameters

| Parameter       | Description                                                          |
| --------------- | ---------------------------------------------------------------------|
| `-SourcePath`    | Folder to scan for Word documents. Defaults to the script's folder. |
| `-Recurse`       | Include subfolders when searching for documents.                    |
| `-Overwrite`     | Replace existing PDFs without asking. Without this flag, the script prompts for each existing PDF. |
| `-ShowWord`      | Show the Word application window during conversion.                 |
| `-NoPause`       | Don't wait for a keypress before closing the window.                |

### Existing PDFs

If a document already has a matching PDF (e.g. `hello.docx` next to `hello.pdf`), the script asks whether to replace it:

```
PDF already exists: C:\Documents\Reports\hello.pdf
Replace it? [Y]es / [N]o / [A]ll / n[o]ne
```

- **Y** — replace this PDF
- **N** — skip this PDF
- **A** — replace this and all remaining PDFs without asking again
- **O** — skip this and all remaining existing PDFs without asking again

Pass `-Overwrite` to skip the prompt entirely and replace everything.

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
4. To update the script, simply replace `C:\Scripts\ExportWordDocsToPdf.ps1` with the latest file.

To remove the menu entry, delete these registry keys:

```
HKEY_CLASSES_ROOT\Directory\Background\shell\ExportWordPDF
HKEY_CLASSES_ROOT\Directory\shell\ExportWordPDF
```

## Notes

- Temporary Word lock files (starting with `~$`) are ignored.
- A summary of converted, skipped, and failed files is printed when the script finishes.
