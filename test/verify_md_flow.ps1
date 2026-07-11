# verify_md_flow.ps1 - Self-test of the md export WPS side (reproduces screenshot-2 failure)
# Goal: prove fix #3 - when the original is .doc, SaveAs2(...,12) must yield a valid .docx so mammoth can read it.
# Chain: build a .doc with text+image (mimicking the user's legacy binary original)
#        -> SaveAs2(intermediate.docx, 12) -> assert valid OOXML -> original .doc unchanged.
# Pair with verify_md_convert.js (Node side runs the real convertToMarkdown) for end-to-end (no GUI dialog).
$ErrorActionPreference = "Stop"

$root = "D:\Claude\WPS_WM_Tools\test"
New-Item -ItemType Directory -Force -Path $root | Out-Null

$imgPath  = Join-Path $root "mdpic.png"
$origDoc  = Join-Path $root "mdtest_orig.doc"
$interDocx= Join-Path $root "mdtest_inter.docx"
Remove-Item $origDoc,$interDocx,$imgPath -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap(30,30)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(255,30,158,99)); $g.Dispose()
$bmp.Save($imgPath,[System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

Write-Host "=== Launch WPS Writer (COM: kwps.Application) - headless ==="
$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = 0 } catch {}

    $doc = $app.Documents.Add(); $sel = $app.Selection
    $sel.TypeText("Body paragraph one."); $sel.TypeParagraph()
    $sel.InlineShapes.AddPicture($imgPath) | Out-Null
    $sel.TypeText(" "); $sel.TypeParagraph()
    $sel.TypeText("Second paragraph."); $sel.TypeParagraph()

    # 1) Save as .doc (format 0 = wdFormatDocument) - mimics the user's legacy .doc original
    $doc.SaveAs2($origDoc, 0) | Out-Null
    Write-Host ("[orig] .doc saved, exists: " + (Test-Path $origDoc))

    # 2) THE FIX: SaveAs2 to .docx with format=12 (wdFormatXMLDocument), converting .doc -> .docx
    #    (also switches the doc away from origDoc, releasing its file lock so we can hash it below)
    $doc.SaveAs2($interDocx, 12) | Out-Null
    Write-Host ("[fix]  .docx saved via SaveAs2(...,12), exists: " + (Test-Path $interDocx))
    $h1 = (Get-FileHash $origDoc -Algorithm SHA256).Hash

    # close doc so the intermediate .docx is released and readable
    try { $doc.Close($false) | Out-Null } catch {}

    # 3) Assert the intermediate is valid OOXML (PK zip header)
    $bytes = [System.IO.File]::ReadAllBytes($interDocx)
    $isPk = ($bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B)
    Write-Host ("[fix]  intermediate is ZIP/OOXML (PK header): " + $isPk)
    if (-not $isPk) {
        Write-Host "[FAIL] SaveAs2(...,12) did NOT produce a valid .docx - mammoth would still fail!" -ForegroundColor Red
    } else {
        Write-Host "[PASS] SaveAs2(...,12) produced a valid .docx (mammoth can read it)" -ForegroundColor Green
    }

    $h2 = (Get-FileHash $origDoc -Algorithm SHA256).Hash
    if ($h1 -eq $h2) { Write-Host "[PASS] original .doc unchanged ($h1)" -ForegroundColor Green }
    else { Write-Host "[FAIL] original .doc was modified!" -ForegroundColor Red }
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
