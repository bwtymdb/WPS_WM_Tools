// quotes.js —— 「英文引号转中文」业务逻辑
// 在作用域（选中文字 / 全文）内把英文直引号成对替换为中文弯引号。
// 判定：逐段落扫描，段内维护「是否在引号内」状态（同时参考已有的弯引号），
//       直引号按当前状态判开/闭；段首重置，段与段互不影响。
// 替换：直引号与弯引号都是 1 个字符，等长就地替换，不破坏格式、不移动后续偏移。
// 撤销：全程原生 Range 操作进撤销栈；WPS 支持 UndoRecord 时合并为一步 Ctrl+Z。

// opts: { double:bool, single:bool }
// onDone(ok)：ok=true 可关闭对话框；false 表示应留对话框（无文档等）。
function ConvertStraightQuotes(opts, onDone) {
    var app = window.Application;
    var doc = app.ActiveDocument;
    if (!doc) { alert("请先打开一个 Word 文档。"); onDone(false); return; }
    var scope = tmGetScope(app);
    if (!scope) { alert("无法获取文档范围。"); onDone(false); return; }

    var stats = { openD: 0, closeD: 0, openS: 0, closeS: 0 };
    var began = tmBeginUndo(app, "WM_TOOLS 英文引号转中文");
    try { app.ScreenUpdating = false; } catch (e) {}
    try {
        convertQuotesInScope(scope.range, opts, stats);
    } catch (e) {
        alert("转换过程出现异常：" + (e && e.message ? e.message : e));
    } finally {
        try { app.ScreenUpdating = true; } catch (e) {}
        if (began) tmEndUndo(app);
    }

    var total = stats.openD + stats.closeD + stats.openS + stats.closeS;
    if (total === 0) {
        alert(scope.label + "内没有找到需要转换的英文引号。");
    } else {
        var lines = ["已在" + scope.label + "内完成转换，可用 Ctrl+Z 撤销。", ""];
        if (opts.double) lines.push("双引号：左 “ " + stats.openD + " 处，右 ” " + stats.closeD + " 处");
        if (opts.single) lines.push("单引号：左 ‘ " + stats.openS + " 处，右 ’ " + stats.closeS + " 处");
        alert(lines.join("\n"));
    }
    onDone(true);
}

function convertQuotesInScope(range, opts, stats) {
    var paras;
    try { paras = range.Paragraphs; } catch (e) { return; }
    var n = 0;
    try { n = paras.Count; } catch (e) {}
    for (var i = 1; i <= n; i++) {
        var p;
        try { p = paras.Item(i); } catch (e) { continue; }
        convertQuotesInParagraph(p, opts, stats);
    }
}

function convertQuotesInParagraph(p, opts, stats) {
    var pr, text;
    try { pr = p.Range; text = pr.Text; } catch (e) { return; }
    if (!text) return;
    if (text.indexOf('"') < 0 && !(opts.single && text.indexOf("'") >= 0)) return;
    var pStart;
    try { pStart = pr.Start; } catch (e) { return; }

    var inD = false, inS = false;   // 段首重置：不成对引号的影响限制在本段内
    var n = text.length;
    for (var i = 0; i < n; i++) {
        var ch = text.charAt(i);
        if (ch === '“') { inD = true; continue; }
        if (ch === '”') { inD = false; continue; }
        if (ch === '‘') { inS = true; continue; }
        if (ch === '’') { inS = false; continue; }

        if (ch === '"' && opts.double) {
            if (inD) { replaceQuoteChar(pr, pStart + i, '”'); inD = false; stats.closeD++; }
            else     { replaceQuoteChar(pr, pStart + i, '“'); inD = true;  stats.openD++; }
            continue;
        }
        if (ch === "'" && opts.single) {
            // 跳过 ASCII 单词内部的撇号（don't、it's），避免误伤英文
            var prev = i > 0 ? text.charAt(i - 1) : "";
            var next = i + 1 < n ? text.charAt(i + 1) : "";
            if (isAsciiWordChar(prev) && isAsciiWordChar(next)) continue;
            if (inS) { replaceQuoteChar(pr, pStart + i, '’'); inS = false; stats.closeS++; }
            else     { replaceQuoteChar(pr, pStart + i, '‘'); inS = true;  stats.openS++; }
        }
    }
}

function isAsciiWordChar(c) { return !!c && /[A-Za-z0-9]/.test(c); }

// 等长替换（直引号 → 弯引号均为 1 字符）：后续字符偏移不变，可边扫描边替换
function replaceQuoteChar(paraRange, absPos, ch) {
    var r;
    try { r = paraRange.Duplicate; } catch (e) { return; }
    try {
        r.Start = absPos;
        r.End = absPos + 1;
        r.Text = ch;
    } catch (e) {}
}
