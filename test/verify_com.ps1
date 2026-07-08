# COM engine-level verification (mirrors finalize.js core calls):
#  1) SaveAs2 to a new path does NOT modify the original file (hash equal) -- #1 hard constraint
#  2) AcceptAllRevisions / Comments / InlineShapes API work; order = turn off track changes first
$ErrorActionPreference = "Stop"

$root = "D:\Claude\wps_doc_removePicandNote\test"
New-Item -ItemType Directory -Force -Path $root | Out-Null
$imgPath  = Join-Path $root "pic.png"
$origPath = Join-Path $root "sample_orig.docx"
$finalPath= Join-Path $root "sample_final.docx"
Remove-Item $origPath,$finalPath,$imgPath -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap(40,40)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(255,30,158,99))
$g.Dispose(); $bmp.Save($imgPath,[System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

function Save-Doc($doc,$path){
    try { $doc.SaveAs2($path) | Out-Null }
    catch { $doc.SaveAs($path) | Out-Null }
}

Write-Host "=== Launch WPS Writer (COM: kwps.Application) ==="
$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = 0 } catch {}

    $doc = $app.Documents.Add()
    $sel = $app.Selection
    $sel.TypeText("Body paragraph one.")
    $sel.TypeParagraph()

    # picture as normal content (track changes still off)
    $shape = $sel.InlineShapes.AddPicture($imgPath)
    Write-Host ("InlineShape.Type = {0} (expect wdInlineShapePicture=3)" -f $shape.Type)
    $sel.TypeParagraph()

    # revision: turn on track changes, then insert text
    $doc.TrackRevisions = $true
    $sel.TypeText("This sentence is an inserted revision.")
    $sel.TypeParagraph()

    # comment anchored to paragraph 1
    try {
        $rng = $doc.Paragraphs.Item(1).Range
        $doc.Comments.Add($rng, "a test comment") | Out-Null
    } catch { Write-Host "Comments.Add failed: $_" }

    # save as "original"
    Save-Doc $doc $origPath
    Write-Host ("[orig] revisions={0} comments={1} inlineShapes={2}" -f $doc.Revisions.Count,$doc.Comments.Count,$doc.InlineShapes.Count)

    # ---- mirror finalize: save-as new file -> process -> save (orig untouched) ----
    Save-Doc $doc $finalPath                      # $doc now points to finalPath; orig is released
    $h1 = (Get-FileHash $origPath -Algorithm SHA256).Hash
    Write-Host "orig SHA256 (before process): $h1"

    $doc.TrackRevisions = $false                  # turn OFF first (else deletes become revisions)
    try { $doc.AcceptAllRevisions() } catch { Write-Host "AcceptAllRevisions error: $_" }
    Write-Host ("after AcceptAllRevisions: revisions={0}" -f $doc.Revisions.Count)

    try { $doc.DeleteAllComments() }
    catch {
        Write-Host "DeleteAllComments not available, iterating: $_"
        for ($i=$doc.Comments.Count; $i -ge 1; $i--) { $doc.Comments.Item($i).Delete() | Out-Null }
    }
    for ($i=$doc.InlineShapes.Count; $i -ge 1; $i--) { $doc.InlineShapes.Item($i).Delete() | Out-Null }
    $doc.Save() | Out-Null

    Write-Host ("[final] revisions={0} comments={1} inlineShapes={2}" -f $doc.Revisions.Count,$doc.Comments.Count,$doc.InlineShapes.Count)

    $h2 = (Get-FileHash $origPath -Algorithm SHA256).Hash
    Write-Host "orig SHA256 (after process):  $h2"
    if ($h1 -eq $h2) { Write-Host "[PASS] original file unchanged at byte level (#1 hard constraint OK)" }
    else { Write-Host "[FAIL] original file was modified!" }
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
