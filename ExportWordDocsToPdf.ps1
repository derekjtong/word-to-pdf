param(
    [string]$SourcePath = $PSScriptRoot,
    [switch]$Recurse,
    [switch]$Overwrite,
    [switch]$ShowWord,
    [switch]$NoPause,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"

# Quiet mode: minimal output, no prompts, exit without waiting
if ($Quiet) {
    $NoPause = $true
}

function Write-Info {
    param([string]$Message = "")

    if (-not $Quiet) {
        Write-Host $Message
    }
}

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

Write-Info "Word to PDF export"
Write-Info "Source folder: $sourceRoot"
Write-Info "Include subfolders: $Recurse"
Write-Info "Overwrite existing PDFs: $Overwrite"
Write-Info "Show Word window: $ShowWord"
Write-Info

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
Write-Info

if (-not $documents) {
    Write-Info "No Word documents found."
    Pause-BeforeExit 0
}

$word = $null
$converted = 0
$skipped = 0
$failed = 0
$total = $documents.Count
$index = 0
$overwriteAll = [bool]$Overwrite
$skipAll = $false

function Confirm-ReplacePdf {
    param([string]$PdfPath)

    while ($true) {
        Write-Host "PDF already exists: $PdfPath"
        $answer = Read-Host "Replace it? [Y]es / [N]o / [A]ll / n[o]ne"

        switch ($answer.Trim().ToLowerInvariant()) {
            "y" { return "Yes" }
            "n" { return "No" }
            "a" { return "All" }
            "o" { return "None" }
            default { Write-Host "Please answer Y, N, A, or O." }
        }
    }
}

function Update-DocumentFields {
    param($Document)

    # Refresh fields in every story: body, headers, footers, text boxes, footnotes...
    foreach ($story in $Document.StoryRanges) {
        $range = $story
        while ($range -ne $null) {
            [void]$range.Fields.Update()
            $range = $range.NextStoryRange
        }
    }

    # TOC-style tables need their own Update() to rebuild entries and page numbers
    foreach ($toc in $Document.TablesOfContents) { [void]$toc.Update() }
    foreach ($tof in $Document.TablesOfFigures) { [void]$tof.Update() }
    foreach ($toa in $Document.TablesOfAuthorities) { [void]$toa.Update() }
}

try {
    Write-Info "Starting Microsoft Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    $word.DisplayAlerts = 0
    Write-Info "Microsoft Word started."
    Write-Info

    foreach ($file in $documents) {
        $index++
        $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")
        $pdfExists = Test-Path -LiteralPath $pdfPath -PathType Leaf
        $pdfIsEmpty = $pdfExists -and ((Get-Item -LiteralPath $pdfPath).Length -eq 0)

        if ($pdfExists -and -not $pdfIsEmpty -and -not $overwriteAll) {
            $replace = $false

            # Quiet mode never prompts; existing PDFs are skipped unless -Overwrite was given
            if (-not $skipAll -and -not $Quiet) {
                switch (Confirm-ReplacePdf $pdfPath) {
                    "Yes" { $replace = $true }
                    "All" { $replace = $true; $overwriteAll = $true }
                    "None" { $skipAll = $true }
                }
            }

            if (-not $replace) {
                if ($Quiet) {
                    Write-Host "[$index/$total] Skipped (PDF exists): $($file.Name)"
                } else {
                    Write-Host "Skipping existing PDF: $pdfPath"
                }
                $skipped++
                continue
            }
        }

        $document = $null

        try {
            if ($Quiet) {
                Write-Host "[$index/$total] $($file.Name)"
            }

            Write-Info "Exporting: $($file.FullName)"
            Write-Info "Target PDF: $pdfPath"

            $document = $word.Documents.Open(
                $file.FullName,
                $false,
                $true,
                $false
            )

            Update-DocumentFields $document

            $document.ExportAsFixedFormat($pdfPath, 17)

            Write-Info "Created: $pdfPath"
            Write-Info
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
        Write-Info "Closing Microsoft Word..."
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}

Write-Info
Write-Host "Done. Converted: $converted; skipped: $skipped; failed: $failed"

if ($failed -gt 0) {
    Pause-BeforeExit 1
}

Pause-BeforeExit 0