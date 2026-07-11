# verify_html.ps1 - Check whether WPS SaveAs filtered HTML emits semantic tags (quality check for the HTML-based md approach)
# Creates a doc with: paragraph + bullet list + table + heading (via OutlineLevel), SaveAs2(html, 10), then inspects the HTML.
# Decision: if <p>/<ul>/<table>/<h1> are present -> HTML approach viable (htmlToMarkdown can recover structure).
$ErrorActionPreference = "Stop"

$root = "D:\Claude\WPS_WM_Tools\test"
New-Item -ItemType Directory -Force -Path $root | Out-Null
$htmlPath = Join-Path $root "htmltest.html"
Remove-Item $htmlPath -ErrorAction SilentlyContinue

$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = 0 } catch {}
    $doc = $app.Documents.Add(); $sel = $app.Selection

    $sel.TypeText("Plain paragraph one."); $sel.TypeParagraph()

    # bullet list
    try { $sel.Range.ListFormat.ApplyBulletDefault() | Out-Null } catch {}
    $sel.TypeText("Bullet A"); $sel.TypeParagraph()
    $sel.TypeText("Bullet B"); $sel.TypeParagraph()
    try { $sel.Range.ListFormat.RemoveNumbers() | Out-Null } catch {}

    # table 2x2
    $tbl = $doc.Tables.Add($sel.Range, 2, 2)
    $tbl.Cell(1,1).Range.Text = "H1"; $tbl.Cell(1,2).Range.Text = "H2"
    $tbl.Cell(2,1).Range.Text = "v1"; $tbl.Cell(2,2).Range.Text = "v2"
    $sel.EndKey(6) | Out-Null   # wdStory = 6 -> end of doc
    $sel.TypeParagraph()

    # heading via OutlineLevel (1 = Heading 1), avoids locale style-name issues
    $sel.TypeText("Heading One")
    try { $doc.Paragraphs.Item($doc.Paragraphs.Count).OutlineLevel = 1 } catch {}
    $sel.TypeParagraph()

    # save as filtered HTML (wdFormatFilteredHTML = 10)
    $doc.SaveAs2($htmlPath, 10) | Out-Null
    Write-Host ("HTML saved, exists: " + (Test-Path $htmlPath))
    try { $doc.Close($false) | Out-Null } catch {}
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}

# inspect HTML
$html = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
Write-Host "=== semantic tag presence ==="
foreach ($tag in @("p>","table","ul>","ol>","h1","li>","tr","td>","strong","b>","i>")) {
    $present = $html -like "*<$tag*"
    Write-Host ("  <" + $tag.PadRight(8) + " : " + $present)
}
Write-Host ("=== HTML size: " + $html.Length + " chars ===")
$n = [Math]::Min(900, $html.Length)
Write-Host "=== first $n chars ==="
Write-Host $html.Substring(0, $n)
