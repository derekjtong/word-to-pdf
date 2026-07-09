# Word to PDF

A PowerShell script that batch-converts Word documents (`.doc`, `.docx`, `.docm`) to PDF using Microsoft Word's COM automation. Includes an optional Windows Explorer right-click context menu integration.

## Requirements

- Windows with Microsoft Word installed
- PowerShell 5.1+ (Windows PowerShell)

## Usage

```powershell
.\ExportWordDocsToPdf.ps1 [-SourcePath <path>] [-Recurse] [-Overwrite] [-ShowWord] [-NoPause] [-Quiet]
```

| Parameter     | Description                                                          |
| ------------- | ---------------------------------------------------------------------|
| `-SourcePath` | Folder to scan for Word documents. Defaults to the script's folder.  |
| `-Recurse`    | Include subfolders when searching for documents.                    |
| `-Overwrite`  | Replace existing PDFs without asking.                                |
| `-ShowWord`   | Show the Word application window during conversion.                 |
| `-NoPause`    | Don't wait for a keypress before closing the window.                |
| `-Quiet`      | Compact one-line-per-file progress, no prompts, closes automatically (implies `-NoPause`). Existing PDFs are skipped unless `-Overwrite` is also given. |

### Existing PDFs

Without `-Overwrite` or `-Quiet`, the script prompts per file:

```
PDF already exists: C:\Documents\Reports\hello.pdf
Replace it? [Y]es / [N]o / [A]ll / n[o]ne
```

`-Overwrite` replaces everything without prompting; `-Quiet` never prompts (replaces only if `-Overwrite` is also set).

### Example

```powershell
.\ExportWordDocsToPdf.ps1 -SourcePath "C:\Documents\Reports" -Recurse -Overwrite
```

## Explorer context menu (optional)

Two `.reg` files add an **"Export Word Docs to PDF"** right-click option (on folders and folder backgrounds) — they're alternatives, not both-at-once, since they register the same menu entry:

- `ExportWordDocsToPdf.reg` — normal mode (prompts, full output)
- `ExportWordDocsToPdf-quiet.reg` — runs with `-Overwrite -Quiet`

Setup:

1. Copy `ExportWordDocsToPdf.ps1` to `C:\Scripts\ExportWordDocsToPdf.ps1` (or edit the `.reg` file to point elsewhere).
2. Double-click the `.reg` file of your choice and confirm the import.
3. Right-click a folder (or its background) and choose **Export Word Docs to PDF**.

To remove the menu entry, delete these registry keys:

```
HKEY_CLASSES_ROOT\Directory\Background\shell\ExportWordPDF
HKEY_CLASSES_ROOT\Directory\shell\ExportWordPDF
```

## Notes

- Temporary Word lock files (starting with `~$`) are ignored.
- A summary of converted, skipped, and failed files is printed when the script finishes.
