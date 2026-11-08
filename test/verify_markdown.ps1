# verify_markdown.ps1 —— Markdown 排版工具的 WPS COM 原语验证（引擎级）
# 在新建的临时文档上验证：内置标题样式 / 原生列表 / 表格 / 超链接 / 底纹 / 边框 /
# TrackRevisions 关闭 / UndoRecord 整批一步撤销。不碰用户已打开的文档。
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"
$script:PASS = 0; $script:FAIL = 0
function Check($name, $cond) {
    if ($cond) { $script:PASS++; Write-Output ("PASS " + $name) }
    else { $script:FAIL++; Write-Output ("FAIL " + $name) }
}

$app = $null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("KWps.Application") } catch {}
if (-not $app) { try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Wps.Application") } catch {} }
if (-not $app) {
    # 没有运行中的 WPS 才自己起一个（脚本结束会关掉临时文档）
    $app = New-Object -ComObject KWps.Application
}
$app.Visible = $true

$doc = $app.Documents.Add()
try {
    $lines = @(
        '# 一级标题',
        '## 二级标题',
        '- 甲项目',
        '- 乙项目',
        '1. 第一步',
        '1. 第二步',
        '1. 第三步',
        '> 引用文字',
        '---',
        '普通段落'
    )
    $doc.Content.Text = ($lines -join "`r")
    $original = $doc.Content.Text
    $doc.TrackRevisions = $false
    Check "TrackRevisions 可写" ($doc.TrackRevisions -eq $false)

    $began = $false
    try { $app.UndoRecord.StartCustomRecord("MD 原语验证"); $began = $true } catch {}

    # —— 1) 内置标题样式（wdStyleHeading1=-2, Heading2=-3）——
    $p1 = $doc.Paragraphs.Item(1).Range
    $styleOk = $true
    try { $p1.Style = $doc.Styles.Item(-2) } catch { $styleOk = $false }
    Check "Styles.Item(-2) 内置标题样式可赋值" $styleOk
    if ($styleOk) {
        $sn = ($doc.Paragraphs.Item(1).Range.Style).NameLocal
        Write-Output ("  样式名=" + $sn)
        Check "标题样式已生效(非正文)" ($sn -notmatch '正文|Normal')
    }

    # —— 2) 原生项目符号列表（段3-4）——
    $r = $doc.Range($doc.Paragraphs.Item(3).Range.Start, $doc.Paragraphs.Item(4).Range.End)
    $ok = $true; try { $r.ListFormat.ApplyBulletDefault() } catch { $ok = $false }
    Check "ApplyBulletDefault 可用" $ok
    Check "项目符号 ListType=2" ($doc.Paragraphs.Item(3).Range.ListFormat.ListType -eq 2)

    # —— 3) 原生编号列表（段5-7）连续编号 ——
    $r = $doc.Range($doc.Paragraphs.Item(5).Range.Start, $doc.Paragraphs.Item(7).Range.End)
    $ok = $true; try { $r.ListFormat.ApplyNumberDefault() } catch { $ok = $false }
    Check "ApplyNumberDefault 可用" $ok
    $lt = $doc.Paragraphs.Item(5).Range.ListFormat.ListType
    Check "编号 ListType 3/4" ($lt -eq 3 -or $lt -eq 4)
    # WPS 中文环境默认给「一、二、」中文编号，且直接改 ListTemplate.NumberStyle 不生效
    # → 用自定义 ListTemplate（阿拉伯数字 %1.）ApplyListTemplate 覆盖
    $ok = $true
    try {
        $ltpl = $doc.ListTemplates.Add($true)
        $lvl = $ltpl.ListLevels.Item(1)
        $lvl.NumberStyle = 0          # wdListNumberStyleArabic
        $lvl.NumberFormat = "%1."
        $lvl.TrailingCharacter = 1    # 编号后空格
        $r.ListFormat.ApplyListTemplate($ltpl)
    } catch { $ok = $false }
    Check "自定义 ListTemplate 阿拉伯编号" $ok
    $ls2 = $doc.Paragraphs.Item(6).Range.ListFormat.ListString
    Write-Output ("  第2项编号=" + $ls2)
    Check "编号连续(第2项=2.)" ($ls2 -match '^2')
    $ok = $true; try { $doc.Paragraphs.Item(7).Range.ListFormat.ListIndent() } catch { $ok = $false }
    Check "ListIndent 二级缩进可用" $ok

    # —— 4) 表格：Tables.Add / 填充 / 表头加粗 / 边框 / 对齐 ——
    $pos = $doc.Content.End - 1
    $ins = $doc.Range($pos, $pos)
    $tbl = $doc.Tables.Add($ins, 3, 2)
    $tbl.Cell(1,1).Range.Text = '名称'; $tbl.Cell(1,2).Range.Text = '数量'
    $tbl.Cell(2,1).Range.Text = '苹果'; $tbl.Cell(2,2).Range.Text = '3'
    $tbl.Cell(3,1).Range.Text = '香蕉'; $tbl.Cell(3,2).Range.Text = '5'
    $tbl.Rows.Item(1).Range.Font.Bold = 1
    $tbl.Borders.InsideLineStyle = 1; $tbl.Borders.OutsideLineStyle = 1
    $tbl.Cell(2,2).Range.ParagraphFormat.Alignment = 1
    Check "Tables.Add 插入成功" ($doc.Tables.Count -eq 1)
    Check "单元格写入" ($tbl.Cell(2,1).Range.Text -match '苹果')
    Check "表头加粗" ($tbl.Cell(1,1).Range.Font.Bold -ne 0)
    Check "表格边框" ($tbl.Borders.InsideLineStyle -eq 1)
    Check "单元格对齐" ($tbl.Cell(2,2).Range.ParagraphFormat.Alignment -eq 1)

    # —— 5) 超链接（保留锚点文字）——
    $last = $doc.Paragraphs.Item(10).Range   # 普通段落
    $anchor = $doc.Range($last.Start, $last.Start + 2)  # “普通”
    $ok = $true
    try { $doc.Hyperlinks.Add($anchor, "https://example.com") | Out-Null } catch { $ok = $false }
    Check "Hyperlinks.Add 可用" ($ok -and $doc.Hyperlinks.Count -ge 1)
    if ($doc.Hyperlinks.Count -ge 1) {
        Check "链接地址正确" ($doc.Hyperlinks.Item(1).Address -eq "https://example.com")
        Check "锚点文字保留" ($anchor.Text -eq '普通')
    }

    # —— 6) 段落底纹 + 左右/下边框 ——
    $q = $doc.Paragraphs.Item(8).Range
    $q.ParagraphFormat.Shading.BackgroundPatternColor = 15921906  # 0xF2F2F2
    Check "段落底纹" ($q.ParagraphFormat.Shading.BackgroundPatternColor -eq 15921906)
    $q.Borders.Item(-2).LineStyle = 1   # wdBorderLeft
    Check "段落左边框(引用竖线)" ($q.Borders.Item(-2).LineStyle -eq 1)
    $hr = $doc.Paragraphs.Item(9).Range
    $hr.Borders.Item(-3).LineStyle = 1  # wdBorderBottom
    Check "段落下边框(分割线)" ($hr.Borders.Item(-3).LineStyle -eq 1)

    # —— 7) 等宽字体 + 删除文字不产生修订 ——
    $q.Font.Name = 'Consolas'
    Write-Output ("  Name=" + $q.Font.Name + " NameAscii=" + $q.Font.NameAscii + " NameFarEast=" + $q.Font.NameFarEast)
    $fontOk = ($q.Font.Name -eq 'Consolas') -or ($q.Font.NameAscii -eq 'Consolas')
    Check "等宽字体设置" $fontOk
    $doc.Range($doc.Paragraphs.Item(9).Range.Start, $doc.Paragraphs.Item(9).Range.End - 1).Text = ''
    Check "删除不产生修订" ($doc.Revisions.Count -eq 0)

    # —— 8) 整批一步撤销 ——
    if ($began) { $app.UndoRecord.EndCustomRecord() }
    $doc.Undo(1)
    $after = $doc.Content.Text
    Check "一步 Ctrl+Z 全部还原(文字)" ($after -eq $original)
    Check "撤销后表格消失" ($doc.Tables.Count -eq 0)
    Check "撤销后超链接消失" ($doc.Hyperlinks.Count -eq 0)
}
finally {
    try { $doc.Close(0) } catch {}   # 0 = wdDoNotSaveChanges
}
Write-Output ("RESULT PASS=" + $script:PASS + " FAIL=" + $script:FAIL)
