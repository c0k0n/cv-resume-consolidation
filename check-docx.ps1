# Check how many pages Word actually renders each DOCX as.
#
# python-docx cannot paginate - only Word's layout engine can. Chrome and Word
# measure text differently, so a DOCX can silently spill onto a second page even
# when the PDF is a clean single page. Run this after any content or styling
# change.
#
# Usage (PowerShell):
#   .\check-docx.ps1
#
# If script execution is blocked, allow it for this session first:
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
#
# Requires Microsoft Word. Writes results to docx-check.txt and prints them.
# Any document showing pages>1 is flagged with OVERFLOWS.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

# Filenames are derived from the source H1, so discover whatever was built
# rather than listing them.
$targets = @()
Get-ChildItem -Path . -Filter *.docx | ForEach-Object { $targets += $_.Name }

# Include anything built into applications/ or examples/
foreach ($dir in @("applications", "examples")) {
    if (Test-Path $dir) {
        Get-ChildItem -Path $dir -Filter *.docx |
            ForEach-Object { $targets += ("$dir\" + $_.Name) }
    }
}

$lines = @()
$word = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $lines += "Word $($word.Version)"
    $lines += ""

    foreach ($rel in $targets) {
        if (-not (Test-Path $rel)) {
            $lines += ("MISSING   {0}" -f $rel)
            continue
        }
        $abs = (Resolve-Path $rel).Path
        $doc = $word.Documents.Open($abs, $false, $true)
        $pages = $doc.ComputeStatistics(2)
        $flag = ""
        if ($pages -gt 1) { $flag = "   <-- OVERFLOWS" }
        $lines += ("{0,-52} pages={1}{2}" -f (Split-Path $abs -Leaf), $pages, $flag)
        $doc.Close($false) | Out-Null
    }
}
catch {
    $lines += "ERROR: $($_.Exception.Message)"
    $lines += "(Microsoft Word is required for this check.)"
}
finally {
    if ($null -ne $word) { $word.Quit() | Out-Null }
}

$lines | Set-Content -Path "docx-check.txt" -Encoding UTF8
$lines