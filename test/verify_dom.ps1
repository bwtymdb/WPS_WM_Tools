# verify_dom.ps1 - Validate the WPS document-object-model API + DOM-walker logic for building Markdown directly.
# Creates heading+paragraph+list+table, then walks doc.Paragraphs printing OutlineLevel / ListString / Range.Tables.Count / Text,
# and dumps table cells. Confirms the API values the JS walker (docToMarkdown) will rely on.
$ErrorActionPreference = "Stop"
$root = "D:\Claude\WPS_WM_Tools\test"
New-Item -ItemType Directory -Force -Path $root | Out-Null

$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = 0 } catch {}
    $doc = $app.Documents.Add(); $sel = $app.Selection

    # body paragraph
    $sel.TypeText("Body paragraph one."); $sel.TypeParagraph()
    # heading via OutlineLevel = 1
    $sel.TypeText("The Heading")
    try { $doc.Paragraphs.Item($doc.Paragraphs.Count).OutlineLevel = 1 } catch {}
    $sel.TypeParagraph()
    # bullet list
    try { $sel.Range.ListFormat.ApplyBulletDefault() | Out-Null } catch {}
    $sel.TypeText("Bullet A"); $sel.TypeParagraph()
    $sel.TypeText("Bullet B"); $sel.TypeParagraph()
    try { $sel.Range.ListFormat.RemoveNumbers() | Out-Null } catch {}
    # table 2x2
    $r = $sel.Range
    $tbl = $doc.Tables.Add($r, 2, 2)
    $tbl.Cell(1,1).Range.Text = "H1"; $tbl.Cell(1,2).Range.Text = "H2"
    $tbl.Cell(2,1).Range.Text = "v1"; $tbl.Cell(2,2).Range.Text = "v2"
    $sel.EndKey(6) | Out-Null
    $sel.TypeParagraph()
    $sel.TypeText("Trailing paragraph.")

    Write-Host "=== walk doc.Paragraphs ==="
    $emitted = @{}
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $ol = 10; try { $ol = $p.OutlineLevel } catch {}
        $ls = ""; try { $ls = $p.Range.ListFormat.ListString } catch {}
        $tc = 0; $rt = $null; try { $rt = $p.Range.Tables; if ($rt) { $tc = $rt.Count } } catch {}
        $txt = ""; try { $txt = ($p.Range.Text -replace "[`r`n`u0007]","") } catch {}
        Write-Host ("[p$i] OutlineLevel=$ol ListStr='$ls' TablesInRange=$tc Text='$txt'")
        if ($tc -gt 0) {
            $t = $rt.Item(1)
            $key = $t.Range.Start
            if (-not $emitted.ContainsKey($key)) {
                $emitted[$key] = $true
                Write-Host "   -> TABLE rows=$($t.Rows.Count)"
                for ($rr = 1; $rr -le $t.Rows.Count; $rr++) {
                    $vals = @()
                    for ($cc = 1; $cc -le $t.Rows.Item($rr).Cells.Count; $cc++) {
                        $ct = ($t.Rows.Item($rr).Cells.Item($cc).Range.Text -replace "[`r`n`u0007]","")
                        $vals += $ct
                    }
                    Write-Host ("      row $rr : " + ($vals -join " | "))
                }
            }
        }
    }
    try { $doc.Close($false) | Out-Null } catch {}
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
