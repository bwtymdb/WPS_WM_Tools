// markdown.js —— 「Markdown 排版」业务逻辑
// 把粘贴进 Word 的 Markdown 原始文本转换为真正的排版：
//   结构级：代码块围栏 / 管道表格 / 标题 / 列表 / 引用 / 分割线
//   行内级：加粗 **、斜体 *、删除线 ~~、行内代码 `、链接 [文字](url)
// 作用域：选中文字 → 选区；未选中 → 全文（tmGetScope）。
// 撤销：全部原生 Range/Tables/Hyperlinks 操作，UndoRecord 合并为一步 Ctrl+Z；
//       删除文字前关闭 TrackRevisions（否则删除会被记成修订而不生效）。
// 实现要点：
//   - 结构阶段（Phase A）按块「从文档尾部向头部」处理，删除/插入只影响后面的位置；
//   - 列表先做 ListFormat（纯格式、不动文字位置），再倒序删除行首标记；
//   - 行内阶段（Phase B）重新逐段读取，段内命中从右往左应用，先删标记再设格式；
//   - 代码段用 NameAscii=Consolas 标记，Phase B 据此跳过（代码内容保持字面）。

// ==================== 纯解析（不依赖 WPS，可被 Node 单测） ====================

var MD_RE = {
    fence:      /^\s{0,3}```(.*)$/,
    fenceClose: /^\s*```\s*$/,
    heading:    /^\s{0,3}(#{1,6})\s+(\S[\s\S]*)$/,
    ul:         /^(\s*)[-*+]\s+(\S[\s\S]*)$/,
    ol:         /^(\s*)\d{1,3}[.)]\s+(\S[\s\S]*)$/,
    quote:      /^\s{0,3}>\s?([\s\S]*)$/,
    hr:         /^\s{0,3}(-{3,}|\*{3,}|_{3,})\s*$/,
    blank:      /^\s*$/
};

// 疑似表格行：含 | 且除 | : - 空格外还有实际内容
function mdIsTableLine(t) {
    if (t.indexOf("|") < 0) return false;
    if (/^[\s|:-]*$/.test(t)) return true;   // 分隔行 |---|:---:| 也算表格行
    return /\|/.test(t) && /[^\s|:-]/.test(t) && t.split("|").length >= 2;
}

// 行首空白 → 列表层级：0~1 个空格为一级，≥2 空格或含 Tab 为二级（更深并入二级）
function mdListLevel(spaces) {
    return (spaces.length >= 2 || spaces.indexOf("\t") >= 0) ? 2 : 1;
}

// 切分表格行：去首尾 |，按 | 分列并 trim
function mdSplitTableRow(line) {
    var t = line.replace(/^\s+|\s+$/g, "");
    if (t.charAt(0) === "|") t = t.substring(1);
    if (t.charAt(t.length - 1) === "|") t = t.substring(0, t.length - 1);
    var cells = t.split("|");
    for (var i = 0; i < cells.length; i++) cells[i] = cells[i].replace(/^\s+|\s+$/g, "");
    return cells;
}

// 解析一组连续表格行；无法解析（无分隔行/列数不一致）返回 null
function mdParseTable(lines) {
    if (!lines || lines.length < 2) return null;
    var sep = lines[1];
    if (sep.indexOf("|") < 0 || sep.indexOf("-") < 0) return null;
    if (!/^[\s|:-]*$/.test(sep)) return null;
    var header = mdSplitTableRow(lines[0]);
    var sepCells = mdSplitTableRow(sep);
    if (sepCells.length !== header.length) return null;
    var aligns = [];
    for (var c = 0; c < sepCells.length; c++) {
        var s = sepCells[c];
        var left = s.charAt(0) === ":", right = s.charAt(s.length - 1) === ":";
        aligns.push(left && right ? 1 : (right ? 2 : 0));   // 0左 1居中 2右（同 Word Alignment 值）
    }
    var rows = [header];
    for (var i = 2; i < lines.length; i++) {
        var r = mdSplitTableRow(lines[i]);
        if (r.length !== header.length) return null;
        rows.push(r);
    }
    return { rows: rows, aligns: aligns };
}

// 逐段分类，返回块数组（块按文档顺序，start/end 为段落下标，含端点）
// kind: code / table / tableSkip / heading / ul / ol / quote / hr / para
function mdClassifyLines(lines) {
    var blocks = [];
    var n = lines.length, i = 0;
    while (i < n) {
        var t = lines[i], j, m;
        if (MD_RE.blank.test(t)) { blocks.push({ kind: "para", start: i, end: i }); i++; continue; }

        if (MD_RE.fence.test(t)) {
            j = i + 1;
            while (j < n && !MD_RE.fenceClose.test(lines[j])) j++;
            if (j < n) { blocks.push({ kind: "code", start: i, end: j }); i = j + 1; continue; }
            // 围栏未闭合：不当代码块，本行按普通段落处理
            blocks.push({ kind: "para", start: i, end: i }); i++; continue;
        }

        if (MD_RE.hr.test(t)) { blocks.push({ kind: "hr", start: i, end: i }); i++; continue; }

        m = MD_RE.heading.exec(t);
        if (m) {
            blocks.push({ kind: "heading", start: i, end: i, level: m[1].length,
                          prefixLen: m[0].length - m[2].length });
            i++; continue;
        }

        if (mdIsTableLine(t)) {
            j = i;
            var group = [];
            while (j < n && mdIsTableLine(lines[j])) { group.push(lines[j]); j++; }
            var parsed = mdParseTable(group);
            if (parsed) blocks.push({ kind: "table", start: i, end: j - 1, rows: parsed.rows, aligns: parsed.aligns });
            else blocks.push({ kind: "tableSkip", start: i, end: j - 1 });
            i = j; continue;
        }

        if (MD_RE.ul.test(t)) {
            var ulItems = [];
            j = i;
            while (j < n) {
                m = MD_RE.ul.exec(lines[j]);
                if (!m) break;
                ulItems.push({ idx: j, level: mdListLevel(m[1]), markerLen: m[0].length - m[2].length });
                j++;
            }
            blocks.push({ kind: "ul", start: i, end: j - 1, items: ulItems }); i = j; continue;
        }

        if (MD_RE.ol.test(t)) {
            var olItems = [];
            j = i;
            while (j < n) {
                m = MD_RE.ol.exec(lines[j]);
                if (!m) break;
                olItems.push({ idx: j, level: mdListLevel(m[1]), markerLen: m[0].length - m[2].length });
                j++;
            }
            blocks.push({ kind: "ol", start: i, end: j - 1, items: olItems }); i = j; continue;
        }

        if (MD_RE.quote.test(t)) {
            j = i;
            while (j < n && MD_RE.quote.test(lines[j])) j++;
            blocks.push({ kind: "quote", start: i, end: j - 1 }); i = j; continue;
        }

        blocks.push({ kind: "para", start: i, end: i }); i++;
    }
    return blocks;
}

// 行内命中：返回按 start 降序的命中数组（互不重叠，高优先级类型先占用区间）
// 每项 { type, start, end, contentStart, contentEnd, url? }，偏移均相对段文本开头
function mdInlineMatches(text, opts) {
    var matches = [], used = [];
    function isFree(s, e) {
        for (var k = 0; k < used.length; k++) {
            if (s < used[k][1] && e > used[k][0]) return false;
        }
        return true;
    }
    function collect(re, type, openLen, closeLen, getUrl) {
        re.lastIndex = 0;
        var m;
        while ((m = re.exec(text)) !== null) {
            var s = m.index, e = s + m[0].length;
            if (m[0].length === 0) { re.lastIndex++; continue; }
            if (!isFree(s, e)) continue;
            used.push([s, e]);
            var hit = { type: type, start: s, end: e,
                        contentStart: s + openLen, contentEnd: e - closeLen };
            if (getUrl) hit.url = getUrl(m);
            matches.push(hit);
        }
    }
    // 顺序即优先级：行内代码（保护内容字面）→ 加粗/斜体（等长星号对）→ 删除线 → 链接
    if (opts.inlineCode) collect(/`([^`]+)`/g, "inlineCode", 1, 1, null);
    if (opts.bold || opts.italic) {
        // 粗/斜统一匹配「等长星号对」：**粗** 或 *斜*。
        // 不能分开扫单星号：否则 **粗**和*斜* 中会产生伪匹配 *和*，吃掉真正的斜体起点。
        var sr = /(\*{1,2})([^*\s][^*]*)\1/g, sm;
        while ((sm = sr.exec(text)) !== null) {
            var stars = sm[1].length;
            var type = stars === 2 ? "bold" : "italic";
            if ((type === "bold" && !opts.bold) || (type === "italic" && !opts.italic)) continue;
            var ss = sm.index, se = ss + sm[0].length;
            if (ss > 0 && text.charAt(ss - 1) === "*") continue;   // *** 等更长序列保守跳过
            if (se < text.length && text.charAt(se) === "*") continue;
            if (!isFree(ss, se)) continue;
            used.push([ss, se]);
            matches.push({ type: type, start: ss, end: se, contentStart: ss + stars, contentEnd: se - stars });
        }
    }
    if (opts.strike)     collect(/~~([^~]+)~~/g, "strike", 2, 2, null);
    if (opts.link) {
        var lr = /\[([^\]]+)\]\(([^)\s]+)\)/g, lm;
        while ((lm = lr.exec(text)) !== null) {
            var ls = lm.index, le = ls + lm[0].length;
            if (ls > 0 && text.charAt(ls - 1) === "!") continue;    // 图片语法不处理
            if (!isFree(ls, le)) continue;
            used.push([ls, le]);
            matches.push({ type: "link", start: ls, end: le,
                           contentStart: ls + 1, contentEnd: ls + 1 + lm[1].length, url: lm[2] });
        }
    }
    matches.sort(function (a, b) { return b.start - a.start; });
    return matches;
}

// ==================== WPS COM 转换 ====================

var MD_HEADING_SIZES = [18, 16, 15, 14, 13, 12];   // 「直接格式化」模式下 1~6 级字号
var MD_CODE_COLOR = 15921906;                      // 0xF2F2F2 浅灰底纹

// opts 全量字段（面板收集）：
// { headings, headingMode:'style'|'format', bold, italic, strike, inlineCode,
//   codeBlock, table, ul, ol, quote, hr, link }

// 读取作用域段落：[{start,end,text}]（text 不含段尾 \r）
function mdReadParas(scopeRange) {
    var out = [];
    var paras, n = 0;
    try { paras = scopeRange.Paragraphs; n = paras.Count; } catch (e) { return out; }
    for (var i = 1; i <= n; i++) {
        try {
            var pr = paras.Item(i).Range;
            var t = pr.Text || "";
            if (t.charAt(t.length - 1) === "\r") t = t.substring(0, t.length - 1);
            out.push({ start: pr.Start, end: pr.End, text: t });
        } catch (e) {}
    }
    return out;
}

function mdDel(doc, start, end) {
    try { doc.Range(start, end).Text = ""; } catch (e) {}
}

// —— Phase A 各结构块（全部使用原始段落位置，自底向上调用） ——

function mdDoHeading(doc, p, block, opts, counts) {
    var contentLen = p.text.length - block.prefixLen;
    if (contentLen <= 0) return;
    mdDel(doc, p.start, p.start + block.prefixLen);          // 去 「## 」前缀
    var r = doc.Range(p.start, p.start + contentLen);
    var done = false;
    if (opts.headingMode !== "format") {
        try { r.Style = doc.Styles.Item(-1 - block.level); done = true; } catch (e) {}
    }
    if (!done) {                                             // 直接格式化：加粗+分级字号，保留字体
        try { r.Font.Bold = 1; r.Font.Size = MD_HEADING_SIZES[block.level - 1]; } catch (e) {}
    }
    counts.headings++;
}

function mdDoList(doc, paras, block, opts, counts) {
    var first = paras[block.start], last = paras[block.end];
    var r = doc.Range(first.start, last.end);
    try {
        if (block.kind === "ul") {
            r.ListFormat.ApplyBulletDefault();
        } else {
            // 中文 WPS 默认编号是「一、」→ 用自定义模板强制阿拉伯数字「1.」
            var ltpl = doc.ListTemplates.Add(true);
            var lvl = ltpl.ListLevels.Item(1);
            lvl.NumberStyle = 0;            // wdListNumberStyleArabic
            lvl.NumberFormat = "%1.";
            lvl.TrailingCharacter = 1;      // 编号后空格
            r.ListFormat.ApplyListTemplate(ltpl);
        }
    } catch (e) {}
    // 二级项缩进（纯格式，不动文字位置）
    for (var k = 0; k < block.items.length; k++) {
        var it = block.items[k];
        if (it.level !== 2) continue;
        try {
            var pr = paras[it.idx];
            doc.Range(pr.start, pr.end).ListFormat.ListIndent();
        } catch (e) {}
    }
    // 倒序删除行首标记（「- 」「1. 」及行首空白）
    for (var k2 = block.items.length - 1; k2 >= 0; k2--) {
        var it2 = block.items[k2], p = paras[it2.idx];
        mdDel(doc, p.start, p.start + it2.markerLen);
        if (block.kind === "ul") counts.ul++; else counts.ol++;
    }
}

function mdDoQuote(doc, paras, block, counts) {
    for (var k = block.end; k >= block.start; k--) {
        var p = paras[k];
        var m = MD_RE.quote.exec(p.text);
        if (!m) continue;
        var prefixLen = p.text.length - m[1].length;
        mdDel(doc, p.start, p.start + prefixLen);
        try {
            // 范围含段落标记（+\r）：部分 WPS 版本对纯文字范围设段落边框不生效
            var r = doc.Range(p.start, p.start + m[1].length + 1);
            r.ParagraphFormat.LeftIndent = 21;               // 左缩进 ~0.74cm
            r.Borders.Item(-2).LineStyle = 1;                // 左侧竖线
            counts.quote++;
        } catch (e) {}
    }
}

function mdDoHr(doc, p, counts) {
    mdDel(doc, p.start, p.end - 1);                          // 清空 「---」文字，保留段落
    try {
        doc.Range(p.start, p.start + 1).Borders.Item(-3).LineStyle = 1;  // 段落下框线
        counts.hr++;
    } catch (e) {}
}

function mdDoCode(doc, paras, block, counts) {
    for (var k = block.start + 1; k <= block.end - 1; k++) {
        var p = paras[k];
        try {
            var r = doc.Range(p.start, p.end);
            r.Font.NameAscii = "Consolas";
            r.ParagraphFormat.Shading.BackgroundPatternColor = MD_CODE_COLOR;
        } catch (e) {}
    }
    mdDel(doc, paras[block.end].start, paras[block.end].end);     // 删闭合围栏
    mdDel(doc, paras[block.start].start, paras[block.start].end); // 删开始围栏（含语言标记）
    counts.codeBlock++;
}

function mdDoTable(doc, paras, block, counts) {
    var tblStart = paras[block.start].start;
    for (var k = block.end; k >= block.start; k--) {               // 倒序删除原文段落
        mdDel(doc, paras[k].start, paras[k].end);
    }
    var rows = block.rows, nRows = rows.length, nCols = rows[0].length;
    var tbl;
    try { tbl = doc.Tables.Add(doc.Range(tblStart, tblStart), nRows, nCols); } catch (e) { return; }
    try {
        for (var r = 0; r < nRows; r++) {
            for (var c = 0; c < nCols; c++) {
                var cell = tbl.Cell(r + 1, c + 1);
                cell.Range.Text = rows[r][c];
                if (block.aligns[c]) cell.Range.ParagraphFormat.Alignment = block.aligns[c];
            }
        }
        tbl.Rows.Item(1).Range.Font.Bold = 1;                      // 表头加粗
        tbl.Borders.InsideLineStyle = 1;
        tbl.Borders.OutsideLineStyle = 1;
        counts.table++;
    } catch (e) {}
}

// —— Phase B：行内格式（重新逐段读取，段内从右往左） ——

function mdApplyInline(doc, pStart, m, counts) {
    var contentLen = m.contentEnd - m.contentStart;
    mdDel(doc, pStart + m.contentEnd, pStart + m.end);        // 尾部标记
    mdDel(doc, pStart + m.start, pStart + m.contentStart);    // 头部标记
    var r = doc.Range(pStart + m.start, pStart + m.start + contentLen);
    try {
        switch (m.type) {
            case "bold":       r.Font.Bold = 1; break;
            case "italic":     r.Font.Italic = 1; break;
            case "strike":     r.Font.StrikeThrough = 1; break;
            case "inlineCode": r.Font.NameAscii = "Consolas"; break;
            case "link":       doc.Hyperlinks.Add(r, m.url); break;
        }
        counts[m.type]++;
    } catch (e) {}
}

function mdInlinePhase(doc, scopeRange, opts, counts) {
    var paras, n = 0;
    try { paras = scopeRange.Paragraphs; n = paras.Count; } catch (e) { return; }
    for (var i = n; i >= 1; i--) {
        var pr, t;
        try { pr = paras.Item(i).Range; t = pr.Text || ""; } catch (e) { continue; }
        if (t.charAt(t.length - 1) === "\r") t = t.substring(0, t.length - 1);
        if (!t) continue;
        var nameAscii = "";
        try { nameAscii = pr.Font.NameAscii; } catch (e) {}
        if (nameAscii === "Consolas") continue;               // 代码段内容保持字面
        var ms = mdInlineMatches(t, opts);
        var pStart = pr.Start;
        for (var k = 0; k < ms.length; k++) mdApplyInline(doc, pStart, ms[k], counts);
    }
}

// ==================== 对外入口 ====================

function mdNewCounts() {
    return { headings: 0, ul: 0, ol: 0, quote: 0, hr: 0, codeBlock: 0,
             table: 0, tableSkip: 0, bold: 0, italic: 0, strike: 0, inlineCode: 0, link: 0 };
}

// 段落对象数组 → 文本数组（不用 Array.map，兼容老 JS 引擎）
function mdParaTexts(paras) {
    var out = [];
    for (var i = 0; i < paras.length; i++) out.push(paras[i].text);
    return out;
}

var MD_ALL_ON = { headings: true, bold: true, italic: true, strike: true, inlineCode: true,
                  codeBlock: true, table: true, ul: true, ol: true, quote: true, hr: true, link: true };

// 检测：返回 { ok, label, stats, total }；stats 与 counts 同结构（tableSkip=疑似但无法解析的表格）
function AnalyzeMarkdown() {
    var app = window.Application;
    var scope = tmGetScope(app);
    if (!scope) return { ok: false, error: "请先在 WPS 中打开一个 Word 文档。" };
    var paras = mdReadParas(scope.range);
    var blocks = mdClassifyLines(mdParaTexts(paras));
    var stats = mdNewCounts();
    for (var b = 0; b < blocks.length; b++) {
        var blk = blocks[b];
        switch (blk.kind) {
            case "heading":   stats.headings++; break;
            case "ul":        stats.ul += blk.items.length; break;
            case "ol":        stats.ol += blk.items.length; break;
            case "quote":     stats.quote += blk.end - blk.start + 1; break;
            case "hr":        stats.hr++; break;
            case "code":      stats.codeBlock++; break;
            case "table":     stats.table++; break;
            case "tableSkip": stats.tableSkip++; break;
        }
    }
    // 行内统计：跳过代码块内段落
    var inCode = {};
    for (var b2 = 0; b2 < blocks.length; b2++) {
        if (blocks[b2].kind !== "code") continue;
        for (var k = blocks[b2].start; k <= blocks[b2].end; k++) inCode[k] = true;
    }
    for (var i = 0; i < paras.length; i++) {
        if (inCode[i] || !paras[i].text) continue;
        var ms = mdInlineMatches(paras[i].text, MD_ALL_ON);
        for (var j = 0; j < ms.length; j++) stats[ms[j].type]++;
    }
    var total = 0;
    for (var key in stats) total += stats[key];
    return { ok: true, label: scope.label, isSelection: scope.isSelection, stats: stats, total: total };
}

// 转换：opts 见上；onDone(ok) 与 quotes.js 同约定（ok=true 可关对话框）
function ConvertMarkdown(opts, onDone) {
    var app = window.Application;
    var doc = app.ActiveDocument;
    if (!doc) { alert("请先打开一个 Word 文档。"); onDone(false); return; }
    var scope = tmGetScope(app);
    if (!scope) { alert("无法获取文档范围。"); onDone(false); return; }

    var counts = mdNewCounts();
    var began = tmBeginUndo(app, "WM_TOOLS Markdown 排版");
    var oldTrack = false, trackSaved = false;
    try { oldTrack = doc.TrackRevisions; doc.TrackRevisions = false; trackSaved = true; } catch (e) {}
    try { app.ScreenUpdating = false; } catch (e) {}
    try {
        var paras = mdReadParas(scope.range);
        var blocks = mdClassifyLines(mdParaTexts(paras));
        // Phase A：结构级，自底向上
        for (var b = blocks.length - 1; b >= 0; b--) {
            var blk = blocks[b];
            try {
                switch (blk.kind) {
                    case "heading":   if (opts.headings)  mdDoHeading(doc, paras[blk.start], blk, opts, counts); break;
                    case "ul":        if (opts.ul)        mdDoList(doc, paras, blk, opts, counts); break;
                    case "ol":        if (opts.ol)        mdDoList(doc, paras, blk, opts, counts); break;
                    case "quote":     if (opts.quote)     mdDoQuote(doc, paras, blk, counts); break;
                    case "hr":        if (opts.hr)        mdDoHr(doc, paras[blk.start], counts); break;
                    case "code":      if (opts.codeBlock) mdDoCode(doc, paras, blk, counts); break;
                    case "table":     if (opts.table)     mdDoTable(doc, paras, blk, counts); break;
                    case "tableSkip": if (opts.table)     counts.tableSkip++; break;
                }
            } catch (e) {}
        }
        // Phase B：行内级
        if (opts.bold || opts.italic || opts.strike || opts.inlineCode || opts.link) {
            mdInlinePhase(doc, scope.range, opts, counts);
        }
    } catch (e) {
        alert("转换过程出现异常：" + (e && e.message ? e.message : e) + "\n已执行的部分可用 Ctrl+Z 一步撤销。");
    } finally {
        try { app.ScreenUpdating = true; } catch (e) {}
        if (trackSaved) { try { doc.TrackRevisions = oldTrack; } catch (e) {} }
        if (began) tmEndUndo(app);
    }

    var total = 0;
    for (var key in counts) { if (key !== "tableSkip") total += counts[key]; }
    if (total === 0 && counts.tableSkip === 0) {
        alert(scope.label + "内没有检测到可转换的 Markdown 元素。");
    } else {
        var lines = ["已在" + scope.label + "内完成排版，可用 Ctrl+Z 一步撤销。", ""];
        var NAMES = { headings: "标题", ul: "项目符号列表项", ol: "编号列表项", quote: "引用段",
                      hr: "分割线", codeBlock: "代码块", table: "表格", bold: "加粗",
                      italic: "斜体", strike: "删除线", inlineCode: "行内代码", link: "超链接" };
        for (var k2 in NAMES) { if (counts[k2]) lines.push(NAMES[k2] + "：" + counts[k2] + " 处"); }
        if (counts.tableSkip) lines.push("跳过无法解析的疑似表格：" + counts.tableSkip + " 个（已保留原文）");
        alert(lines.join("\n"));
    }
    onDone(true);
}

// Node 单测导出（WPS 环境下 module 未定义，不影响）
if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        mdClassifyLines: mdClassifyLines,
        mdInlineMatches: mdInlineMatches,
        mdParseTable: mdParseTable,
        mdIsTableLine: mdIsTableLine
    };
}
