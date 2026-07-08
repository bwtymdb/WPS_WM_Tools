# Parameterized picture-deletion verification (mirrors refactored finalize.js).
# Builds a rich doc, then runs cleanup with a MIXED option set:
#   pic = { inline:$false, floating:$true, headerFooter:$false, inTextbox:$true }
# Expected: body inline pic KEPT, floating pic DELETED, header pic KEPT,
#           textbox inner pic DELETED (textbox shell kept); revisions accepted; comments gone; orig unchanged.
$ErrorActionPreference = "Stop"

$root = "D:\Claude\wps_doc_removePicandNote\test"
New-Item -ItemType Directory -Force -Path $root | Out-Null
$imgPath  = Join-Path $root "pic.png"
$origPath = Join-Path $root "sample_mix.docx"
$finalPath= Join-Path $root "sample_mix_final.docx"
Remove-Item $origPath,$finalPath -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Drawing
$bmp = New-Object System.Drawing.Bitmap(30,30)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::FromArgb(255,43,87,154)); $g.Dispose()
$bmp.Save($imgPath,[System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

function Save-Doc($doc,$path){ try { $doc.SaveAs2($path) | Out-Null } catch { $doc.SaveAs($path) | Out-Null } }

$PIC_INLINE = @{3=$true;4=$true;9=$true;10=$true;11=$true}
$PIC_SHAPE  = @{13=$true;11=$true}
$MSO_GROUP  = 6

function Clean-Inline($shapes){
    if(-not $shapes){return}
    for($i=$shapes.Count;$i -ge 1;$i--){ try{ $sh=$shapes.Item($i); if($PIC_INLINE[[int]$sh.Type]){ $sh.Delete()|Out-Null } }catch{} }
}
function Visit-Shape($sh,$floatingFlag,$p){
    try{ $t=[int]$sh.Type }catch{ return }
    if($PIC_SHAPE[$t]){ if($floatingFlag){ try{ $sh.Delete()|Out-Null }catch{} }; return }
    if($t -eq $MSO_GROUP){ try{ $gi=$sh.GroupItems; for($i=$gi.Count;$i -ge 1;$i--){ try{ Visit-Shape $gi.Item($i) $floatingFlag $p }catch{} } }catch{} }
    if($p.inTextbox){ try{ $tr=$sh.TextFrame.TextRange; if($tr){ Clean-Inline $tr.InlineShapes; Traverse-Shapes $tr.Shapes $p.inTextbox $p } }catch{} }
}
function Traverse-Shapes($shapes,$floatingFlag,$p){
    if(-not $shapes){return}
    for($i=$shapes.Count;$i -ge 1;$i--){ try{ Visit-Shape $shapes.Item($i) $floatingFlag $p }catch{} }
}
function Delete-AllPictures($doc,$p){
    if($p.inline){ Clean-Inline $doc.InlineShapes }
    Traverse-Shapes $doc.Shapes $p.floating $p
    $secs=$doc.Sections
    for($s=1;$s -le $secs.Count;$s++){
        $sec=$secs.Item($s)
        foreach($col in @($sec.Headers,$sec.Footers)){
            foreach($k in 1..3){ $hf=$null; try{ $hf=$col.Item($k) }catch{}; if($hf){
                if($p.headerFooter){ try{ Clean-Inline $hf.Range.InlineShapes }catch{} }
                try{ Traverse-Shapes $hf.Shapes $p.headerFooter $p }catch{}
            }}
        }
    }
}

Write-Host "=== Launch WPS Writer (COM) for parameterized test ==="
$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible=$false } catch {}
    try { $app.DisplayAlerts=0 } catch {}
    $doc=$app.Documents.Add(); $sel=$app.Selection
    $sel.TypeText("Body text. "); $sel.TypeParagraph()
    $ip=$sel.InlineShapes.AddPicture($imgPath); $sel.TypeText(" ")       # body inline
    $fp=$doc.Shapes.AddPicture($imgPath)                                  # body floating
    $sel.TypeParagraph()
    $hdr=$doc.Sections.Item(1).Headers.Item(1)
    $hp=$hdr.Range.InlineShapes.AddPicture($imgPath)                      # header inline
    $tb=$doc.Shapes.AddTextbox(1,100,100,200,80)                          # textbox
    $tp=$tb.TextFrame.TextRange.InlineShapes.AddPicture($imgPath)         # textbox inner
    $doc.TrackRevisions=$true; $sel.TypeText("revised insert.")           # revision
    try { $doc.Comments.Add($doc.Paragraphs.Item(1).Range,"note")|Out-Null } catch {}

    Save-Doc $doc $origPath
    Write-Host ("[orig] revisions={0} comments={1} bodyInline={2} bodyShapes={3} headerInline={4}" -f $doc.Revisions.Count,$doc.Comments.Count,$doc.InlineShapes.Count,$doc.Shapes.Count,$hdr.Range.InlineShapes.Count)

    Save-Doc $doc $finalPath
    $h1=(Get-FileHash $origPath -Algorithm SHA256).Hash
    $doc.TrackRevisions=$false
    try { $doc.AcceptAllRevisions() } catch {}
    try { $doc.DeleteAllComments() } catch { for($i=$doc.Comments.Count;$i -ge 1;$i--){ $doc.Comments.Item($i).Delete()|Out-Null } }
    # MIXED options: keep inline & header, delete floating & textbox-inner
    $p = @{ inline=$false; floating=$true; headerFooter=$false; inTextbox=$true }
    Delete-AllPictures $doc $p
    $doc.Save()|Out-Null

    $hdr2=$doc.Sections.Item(1).Headers.Item(1)
    Write-Host ("[final] revisions={0} comments={1} bodyInline={2} bodyShapes={3} headerInline={4}" -f $doc.Revisions.Count,$doc.Comments.Count,$doc.InlineShapes.Count,$doc.Shapes.Count,$hdr2.Range.InlineShapes.Count)
    Write-Host "expect: bodyInline=1(kept) bodyShapes=1(textbox kept) headerInline=1(kept)"
    $h2=(Get-FileHash $origPath -Algorithm SHA256).Hash
    if ($h1 -eq $h2) { Write-Host "[PASS] original unchanged" } else { Write-Host "[FAIL] orig changed" }
    try { $doc.Close($false)|Out-Null } catch {}
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
