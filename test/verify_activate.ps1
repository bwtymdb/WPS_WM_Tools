# verify_activate.ps1 - Confirm the reopen+activate+close sequence leaves the ORIGINAL as ActiveDocument.
# Mirrors the new SaveFinalVersion ending: SaveAs2(intermediate) -> Open(original) -> Activate -> Close(intermediate).
$ErrorActionPreference = "Stop"
$root = "D:\Claude\WPS_WM_Tools\test"
$origPath = Join-Path $root "act_orig.docx"
$interPath = Join-Path $root "act_inter.docx"
Remove-Item $origPath,$interPath -ErrorAction SilentlyContinue

$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = 0 } catch {}
    $doc = $app.Documents.Add(); $sel = $app.Selection
    $sel.TypeText("Original content."); $sel.TypeParagraph()
    $doc.SaveAs2($origPath, 12) | Out-Null
    Write-Host ("original saved: " + $origPath)

    # snapshot to intermediate (doc object now points to intermediate; original closed)
    $doc.SaveAs2($interPath, 12) | Out-Null
    Write-Host ("after SaveAs2(inter): docs.Count=" + $app.Documents.Count + " Active=" + $app.ActiveDocument.Name)

    # reopen original (intermediate still open => 2 docs)
    $orig = $app.Documents.Open($origPath)
    Write-Host ("after Open(orig): docs.Count=" + $app.Documents.Count + " Active=" + $app.ActiveDocument.Name)

    $orig.Activate()
    Write-Host ("after orig.Activate: Active=" + $app.ActiveDocument.Name)

    # close the intermediate (the $doc object became the intermediate via SaveAs2)
    $doc.Close(0)
    Write-Host ("after Close(inter): docs.Count=" + $app.Documents.Count + " Active=" + $app.ActiveDocument.Name)

    if ($app.ActiveDocument.FullName -eq $origPath) {
        Write-Host "[PASS] final ActiveDocument is the ORIGINAL" -ForegroundColor Green
    } else {
        Write-Host ("[FAIL] final ActiveDocument is NOT the original: " + $app.ActiveDocument.FullName) -ForegroundColor Red
    }
    try { $app.ActiveDocument.Close($false) | Out-Null } catch {}
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
