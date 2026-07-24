# COM engine-level verification for the text-mod tools (mirrors textmod.js / quotes.js / supsub.js core calls):
#  1) Quotes: paragraph scan + same-length Range.Text replacement pairs quotes correctly; Undo restores.
#  2) Supsub: wildcard Find + ReplaceAll with Replacement.Font.Superscript (ReplaceWith="^&"):
#     text unchanged, format applied to matches only, ONE Undo reverts everything (native one-step undo).
#  3) Preset wildcard matches [3-4] [1] [12] [1,3] but not [abc]; manual find loop (no-arg Execute) counts 4.
#  4) Report whether Application.UndoRecord is available in this WPS build.
$ErrorActionPreference = "Stop"

$pass = 0; $fail = 0
function Check($name, $cond) {
    if ($cond) { Write-Host "[PASS] $name"; $script:pass++ }
    else { Write-Host "[FAIL] $name"; $script:fail++ }
}

# supsub.js SUPSUB_PRESETS[1] —— 与 js/supsub.js 中的预设通配符保持一致
$presetBracket = '[\[［][-0-9,，、–— ]@[\]］]'

Write-Host "=== Launch WPS Writer (COM: kwps.Application) ==="
$app = New-Object -ComObject kwps.Application
try {
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = 0 } catch {}

    # ---------- 0) UndoRecord availability ----------
    $undoRecordOK = $false
    try { $app.UndoRecord.StartCustomRecord("probe"); $app.UndoRecord.EndCustomRecord(); $undoRecordOK = $true } catch {}
    Write-Host ("UndoRecord supported: {0}" -f $undoRecordOK)

    # ================= Part 1: quotes (mirrors quotes.js) =================
    $doc = $app.Documents.Add()
    $sel = $app.Selection
    $sel.TypeText('于"月经后期""闭经""崩漏"等')
    $sel.TypeParagraph()
    $sel.TypeText('已有“正确”引号和"直引号"并存')

    if ($undoRecordOK) { try { $app.UndoRecord.StartCustomRecord("quotes-test") } catch {} }

    $open  = [char]0x201C   # “
    $close = [char]0x201D   # ”
    $paras = $doc.Content.Paragraphs
    for ($pi = 1; $pi -le $paras.Count; $pi++) {
        $para = $paras.Item($pi)
        $text = $para.Range.Text
        if (-not $text.Contains('"')) { continue }
        $pStart = $para.Range.Start
        $inD = $false
        for ($i = 0; $i -lt $text.Length; $i++) {
            $ch = $text[$i]
            if ($ch -eq $open)  { $inD = $true;  continue }
            if ($ch -eq $close) { $inD = $false; continue }
            if ($ch -eq '"') {
                $r = $para.Range.Duplicate
                $r.Start = $pStart + $i
                $r.End   = $pStart + $i + 1
                if ($inD) { $r.Text = [string]$close; $inD = $false }
                else      { $r.Text = [string]$open;  $inD = $true  }
            }
        }
    }

    if ($undoRecordOK) { try { $app.UndoRecord.EndCustomRecord() } catch {} }

    $t = $doc.Content.Text
    Check "quotes: pairs converted (于“月经后期”“闭经”“崩漏”等)" $t.Contains('于“月经后期”“闭经”“崩漏”等')
    Check "quotes: state respects existing curly quotes (“直引号”)" $t.Contains('已有“正确”引号和“直引号”并存')
    Check "quotes: no straight double quote left" (-not $t.Contains('"'))

    # Undo must restore (1 step if UndoRecord works, otherwise several)
    $steps = 0
    while ($steps -lt 50 -and -not $doc.Content.Text.Contains('"')) { $doc.Undo() | Out-Null; $steps++ }
    Check ("quotes: Undo restores original text (steps=$steps)") $doc.Content.Text.Contains('于"月经后期"')
    try { $doc.Close(0) } catch {}

    # ================= Part 2: supsub (mirrors supsub.js) =================
    $doc2 = $app.Documents.Add()
    $sel2 = $app.Selection
    $sel2.TypeText('病位主要在肾、肝、脾三脏及冲任、胞宫[3-4]。文献[1]与[12]、混合[1,3]，而[abc]保留。')

    # 2a) manual find loop (mirrors tmForEachMatch: no-arg Execute + Wrap=0) should count 4
    $scope = $doc2.Content
    $scopeEnd = $scope.End
    $mr = $scope.Duplicate
    $mf = $mr.Find
    $mf.ClearFormatting()
    $mf.Text = $presetBracket
    $mf.Forward = $true
    $mf.Wrap = 0
    $mf.Format = $false
    $mf.MatchCase = $true
    $mf.MatchWildcards = $true
    $count = 0; $guard = 0
    while ($guard++ -lt 1000) {
        $found = $false
        try { $found = $mf.Execute() } catch { break }
        if (-not $found) { break }
        $mEnd = $mr.End
        if ($mEnd -gt $scopeEnd) { break }
        $count++
        $mr.Start = $mEnd
        if ($mEnd -ge $scopeEnd) { break }
        $mr.End = $scopeEnd
    }
    Check "supsub: manual loop counts 4 matches ([3-4],[1],[12],[1,3])" ($count -eq 4)

    # 2b) format-only ReplaceAll (mirrors formatReplaceAll)
    $fr = $doc2.Content
    $f = $fr.Find
    $f.ClearFormatting()
    $f.Replacement.ClearFormatting()
    $f.Replacement.Text = "^&"
    $f.Replacement.Font.Superscript = $true
    # (FindText, MatchCase, MatchWholeWord, MatchWildcards, SoundsLike, AllWordForms, Forward, Wrap, Format, ReplaceWith, Replace=2)
    $f.Execute($presetBracket, $true, $false, $true, $false, $false, $true, 0, $true, "^&", 2) | Out-Null

    $t2 = $doc2.Content.Text
    Check "supsub: text unchanged after format-only replace (still has [3-4])" $t2.Contains('[3-4]')
    Check "supsub: no literal ^& leaked into text" (-not $t2.Contains('^&'))

    # superscript applied to [3-4] (plain literal find, wildcards OFF)
    $chk = $doc2.Content
    $cf = $chk.Find
    $cf.ClearFormatting()
    $cf.MatchWildcards = $false
    $cf.Text = '[3-4]'
    $cf.Execute() | Out-Null
    $sup = $chk.Font.Superscript
    Check "supsub: [3-4] is superscript (Font.Superscript=$sup)" ($sup -eq -1 -or $sup -eq $true)

    # [abc] must NOT be superscript
    $chk2 = $doc2.Content
    $cf2 = $chk2.Find
    $cf2.ClearFormatting()
    $cf2.MatchWildcards = $false
    $cf2.Text = '[abc]'
    $cf2.Execute() | Out-Null
    $sup2 = $chk2.Font.Superscript
    Check "supsub: [abc] left untouched (Font.Superscript=$sup2)" ($sup2 -eq 0 -or $sup2 -eq $false)

    # ONE Undo must revert the whole ReplaceAll
    $doc2.Undo() | Out-Null
    $chk3 = $doc2.Content
    $cf3 = $chk3.Find
    $cf3.ClearFormatting()
    $cf3.MatchWildcards = $false
    $cf3.Text = '[3-4]'
    $cf3.Execute() | Out-Null
    $sup3 = $chk3.Font.Superscript
    Check "supsub: single Undo reverts ReplaceAll (Font.Superscript=$sup3)" ($sup3 -eq 0 -or $sup3 -eq $false)
    try { $doc2.Close(0) } catch {}

    Write-Host ("`n==== summary: PASS={0} FAIL={1} ====" -f $pass, $fail)
    if ($fail -gt 0) { exit 1 }
} finally {
    try { $app.Quit() } catch {}
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
}
