param(
    [string]$SourcePath = $PSScriptRoot,
    [switch]$Recurse,
    [switch]$Overwrite,
    [switch]$ShowWord,
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"

function Pause-BeforeExit {
    param([int]$ExitCode)

    if (-not $NoPause) {
        Write-Host ""
        Read-Host "Press Enter to close this window"
    }

    exit $ExitCode
}

function Get-AbsolutePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return (Get-Location).Path
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

trap {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Pause-BeforeExit 1
}

$sourceRoot = Get-AbsolutePath $SourcePath

Write-Host "Word to PDF export"
Write-Host "Source folder: $sourceRoot"
Write-Host "Include subfolders: $Recurse"
Write-Host "Overwrite existing PDFs: $Overwrite"
Write-Host "Show Word window: $ShowWord"
Write-Host ""

if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
    Write-Error "SourcePath does not exist or is not a folder: $sourceRoot"
    Pause-BeforeExit 1
}

$searchOptions = @{
    LiteralPath = $sourceRoot
    File        = $true
}

if ($Recurse) {
    $searchOptions.Recurse = $true
}

$documents = Get-ChildItem @searchOptions |
    Where-Object { $_.Extension -in ".doc", ".docx", ".docm" } |
    Where-Object { -not $_.Name.StartsWith("~$") } |
    Sort-Object FullName

Write-Host "Documents found: $($documents.Count)"
Write-Host ""

if (-not $documents) {
    Write-Host "No Word documents found."
    Pause-BeforeExit 0
}

$word = $null
$converted = 0
$skipped = 0
$failed = 0

try {
    Write-Host "Starting Microsoft Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    $word.DisplayAlerts = 0
    Write-Host "Microsoft Word started."
    Write-Host ""

    foreach ($file in $documents) {
        $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")
        $pdfExists = Test-Path -LiteralPath $pdfPath -PathType Leaf
        $pdfIsEmpty = $pdfExists -and ((Get-Item -LiteralPath $pdfPath).Length -eq 0)

        if ($pdfExists -and -not $pdfIsEmpty -and -not $Overwrite) {
            Write-Host "Skipping existing PDF: $pdfPath"
            $skipped++
            continue
        }

        $document = $null

        try {
            Write-Host "Exporting: $($file.FullName)"
            Write-Host "Target PDF: $pdfPath"

            $document = $word.Documents.Open(
                $file.FullName,
                $false,
                $true,
                $false
            )

            $document.ExportAsFixedFormat($pdfPath, 17)

            Write-Host "Created: $pdfPath"
            Write-Host ""
            $converted++
        }
        catch {
            Write-Warning "Failed: $($file.FullName) - $($_.Exception.Message)"
            $failed++
        }
        finally {
            if ($document -ne $null) {
                $document.Close($false)
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
            }
        }
    }
}
finally {
    if ($word -ne $null) {
        Write-Host "Closing Microsoft Word..."
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}

Write-Host ""
Write-Host "Done. Converted: $converted; skipped: $skipped; failed: $failed"

if ($failed -gt 0) {
    Pause-BeforeExit 1
}

Pause-BeforeExit 0