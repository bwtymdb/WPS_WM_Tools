# Extended engine verification: mirrors finalize.js deleteAllPictures recursively.
# Builds a doc with: body inline pic, body floating pic, header inline pic, textbox inline pic,
# plus a comment and a revision; then runs the full finalize flow and checks all pictures gone
# and the original file unchanged (hash).
$ErrorActionPreference = "Stop"

$root = "D:\Claude\wps_doc_removePicandNote\test"
New-Item -ItemType Directory -Force -Path $root | Out-Null
$imgPath  = Join-Path $root "pic.png"
$origPath = Join-Path $root "sample_rich.docx"
$finalPath= Join-Path $root "sample_rich_final.docx"
Remove-Item $origPath,$finalPath -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap(30,30)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(255,43,87,154)); $g.Dispose()
$bmp.Save($imgPath,[System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

function Save-Doc($doc,$path){ try { $doc.SaveAs2($path) | Out-Null } catch { $doc.SaveAs($path) | Out-Null } }

# picture-type filter sets (mirror finalize.js constants)
$PIC_INLINE = @{3=$true;4=$true;9=$true;10=$true;11=$true}
$PIC_SHAPE  = @{13=$true;11=$true}
$MSO_GROUP  = 6

function Clean-Inline($shapes){
    if (-not $shapes) { return }
    for ($i=$shapes.Count; $i -ge 1; $i--) {
        try { $sh=$shapes.Item($i); if ($PIC_INLINE[[int]$sh.Type]) { $sh.Delete() | Out-Null } } catch {}
    }
}
function Clean-OneShape($sh){
    try {
        $t=[int]$sh.Type
        if ($PIC_SHAPE[$t]) { $sh.Delete() | Out-Null; return }
        if ($t -eq $MSO_GROUP) {
            $gi=$sh.GroupItems
            for ($i=$gi.Count; $i -ge 1; $i--) { try { Clean-OneShape $gi.Item($i) } catch {} }
        }
        $tr=$sh.TextFrame.TextRange
        if ($tr) { Clean-Inline $tr.InlineShapes; Clean-Shapes $tr.Shapes }
    } catch {}
}
function Clean-Shapes($shapes){
    if (-not $shapes) { return }
    for ($i=$shapes.Count; $i -ge 1; $i--) { try { Clean-OneShape $shapes.Item($i) } catch {} }
}
function Delete-AllPictures($doc){
    Clean-Inline $doc.InlineShapes
    Clean-Shapes $doc.Shapes
    $secs=$doc.Sections
    for ($s=1; $s -le $secs.Count; $s++) {
        $sec=$secs.Item($s)
        foreach ($col in @($sec.Headers,$sec.Footers)) {
            foreach ($k in 1..3) {
                $hf=$null; try { $hf=$col.Item($k) } catch {}
                if ($hf) {
                    try { Clean-Inline $hf.Range.InlineShapes } catch {}
                    try { Clean-Shapes $hf.Shapes } catch {}
                }
            }
        }
    }
}

Write-Host "=== Launch WPS Writer (COM) for extended picture test ==="
$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible=$false } catch {}
    try { $app.DisplayAlerts=0 } catch {}
    $doc=$app.Documents.Add(); $sel=$app.Selection
    $sel.TypeText("Body text. "); $sel.TypeParagraph()

    # 1) body inline picture
    $ip=$sel.InlineShapes.AddPicture($imgPath); $sel.TypeText(" ")
    Write-Host ("body inline pic Type={0}" -f $ip.Type)

    # 2) body floating picture
    $fp=$doc.Shapes.AddPicture($imgPath)
    Write-Host ("body floating pic Type={0} (expect msoPicture=13)" -f $fp.Type)
    $sel.TypeParagraph()

    # 3) header inline picture
    $hdr=$doc.Sections.Item(1).Headers.Item(1)
    $hp=$hdr.Range.InlineShapes.AddPicture($imgPath)
    Write-Host ("header inline pic Type={0}" -f $hp.Type)

    # 4) textbox containing a picture
    $tb=$doc.Shapes.AddTextbox(1,100,100,200,80)
    $tp=$tb.TextFrame.TextRange.InlineShapes.AddPicture($imgPath)
    Write-Host ("textbox inline pic Type={0}" -f $tp.Type)

    # 5) revision
    $doc.TrackRevisions=$true; $sel.TypeText("revised insert.")

    # 6) comment
    try { $doc.Comments.Add($doc.Paragraphs.Item(1).Range,"note") | Out-Null } catch { Write-Host "comment add failed: $_" }

    Save-Doc $doc $origPath
    Write-Host ("[orig] revisions={0} comments={1} bodyInline={2} bodyShapes={3} headerInline={4}" -f $doc.Revisions.Count,$doc.Comments.Count,$doc.InlineShapes.Count,$doc.Shapes.Count,$hdr.Range.InlineShapes.Count)

    # mirror finalize
    Save-Doc $doc $finalPath
    $h1=(Get-FileHash $origPath -Algorithm SHA256).Hash
    $doc.TrackRevisions=$false
    try { $doc.AcceptAllRevisions() } catch { Write-Host "AcceptAllRevisions: $_" }
    try { $doc.DeleteAllComments() } catch { for($i=$doc.Comments.Count;$i -ge 1;$i--){ $doc.Comments.Item($i).Delete()|Out-Null } }
    Delete-AllPictures $doc
    $doc.Save() | Out-Null

    # refresh header ref from final doc
    $hdr2=$doc.Sections.Item(1).Headers.Item(1)
    Write-Host ("[final] revisions={0} comments={1} bodyInline={2} bodyShapes={3} headerInline={4}" -f $doc.Revisions.Count,$doc.Comments.Count,$doc.InlineShapes.Count,$doc.Shapes.Count,$hdr2.Range.InlineShapes.Count)
    $h2=(Get-FileHash $origPath -Algorithm SHA256).Hash
    if ($h1 -eq $h2) { Write-Host "[PASS] original unchanged ($h1)" } else { Write-Host "[FAIL] orig changed: $h1 -> $h2" }
    try { $doc.Close($false) | Out-Null } catch {}
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
