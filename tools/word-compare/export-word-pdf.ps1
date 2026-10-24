# Exports docs\<name>.docx to out\<name>-word.pdf with Microsoft Word (read-only, nothing saved).
# Usage: powershell -File export-word-pdf.ps1 -WorkDir D:\some\folder
param([Parameter(Mandatory)][string]$WorkDir)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
New-Item -ItemType Directory -Force (Join-Path $WorkDir 'out') | Out-Null
try {
  Get-ChildItem (Join-Path $WorkDir 'docs') -Filter *.docx | ForEach-Object {
    $doc = $word.Documents.Open($_.FullName, $false, $true)
    $doc.ExportAsFixedFormat((Join-Path $WorkDir "out\$($_.BaseName)-word.pdf"), 17)
    "$($_.BaseName): $($doc.ComputeStatistics(2)) pages in Word"
    $doc.Close(0)
  }
} finally { $word.Quit() }
