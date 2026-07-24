// supsub.js —— 「上下标批量设置」业务逻辑
// 把作用域（选中文字 / 全文）内匹配指定模式的内容批量设为上标 / 下标（或取消上下标）。
// 典型场景：论文引用标注 [3-4] → 上标。
//
// 两条实现路径：
//   主路径（设上/下标且不去首尾字符）：Find 通配符 + 仅格式替换
//     （Replacement.Text="^&" + Format=True + Replace=wdReplaceAll），
//     与 WPS 自带「替换 → 格式 → 全部替换」完全相同，原生一步撤销。
//   回退路径（取消上下标、去首尾字符，或主路径抛异常）：
//     手动查找循环逐个命中设置 Font.Superscript / Subscript。
//
// 撤销：全程原生 Find/Range 操作进撤销栈；WPS 支持 UndoRecord 时合并为一步 Ctrl+Z。

// 预设模式（面板与业务共用）。
// 通配符说明：[\[［] 左方括号（半角/全角）；[-0-9,，、–— ]@ 一个或多个「数字/逗号/顿号/连字符」；
//             [（(]、[）)] 半角或全角圆括号。
var SUPSUB_PRESETS = {
    1: { name: "方括号引用标注", example: "[1]、[2-5]、[1,3]",
         wildcard: "[\\[［][-0-9,，、–— ]@[\\]］]" },
    2: { name: "圆括号引用标注", example: "(1)、（1-3）",
         wildcard: "[（(][-0-9,，、–— ]@[）)]" }
};

// opts: { preset:1|2|3, customPattern:string, mode:"super"|"sub"|"clear", trimEnds:bool }
function resolveSupsubWildcard(opts) {
    if (opts.preset === 1 || opts.preset === 2) {
        return { ok: true, wildcard: SUPSUB_PRESETS[opts.preset].wildcard, warnings: [] };
    }
    return tmTranslatePattern(opts.customPattern || "");
}

// 只统计不修改（对话框「先数一下」用）。
// 返回 { ok, count, scopeLabel, wildcard, warnings } 或 { ok:false, error }
function CountSupsubMatches(opts) {
    var app = window.Application;
    var doc = app.ActiveDocument;
    if (!doc) return { ok: false, error: "请先打开一个 Word 文档。" };
    var w = resolveSupsubWildcard(opts);
    if (!w.ok) return { ok: false, error: w.warnings.join("；") || "模式无效" };
    var scope = tmGetScope(app);
    if (!scope) return { ok: false, error: "无法获取文档范围。" };
    var n = tmForEachMatch(scope.range, w.wildcard, null);
    return { ok: true, count: n, scopeLabel: scope.label, wildcard: w.wildcard, warnings: w.warnings };
}

// onDone(ok)：ok=true 可关闭对话框；false 表示应留对话框（无文档 / 模式无效 / 0 命中）。
function ApplyVertAlign(opts, onDone) {
    var app = window.Application;
    var doc = app.ActiveDocument;
    if (!doc) { alert("请先打开一个 Word 文档。"); onDone(false); return; }
    var w = resolveSupsubWildcard(opts);
    if (!w.ok) { alert("模式无效：" + (w.warnings.join("；") || "内容为空")); onDone(false); return; }
    var scope = tmGetScope(app);
    if (!scope) { alert("无法获取文档范围。"); onDone(false); return; }

    // 先统计命中数（用于报告与 0 命中拦截），再执行修改
    var total = tmForEachMatch(scope.range, w.wildcard, null);
    if (total === 0) {
        alert(scope.label + "内没有匹配的内容。\n\n实际使用的通配符：" + w.wildcard);
        onDone(false);
        return;
    }

    var began = tmBeginUndo(app, "WM_TOOLS 上下标批量设置");
    try { app.ScreenUpdating = false; } catch (e) {}
    var usedManual = (opts.mode === "clear" || opts.trimEnds);
    try {
        if (usedManual) {
            manualApplyVertAlign(scope.range, w.wildcard, opts);
        } else {
            formatReplaceAll(scope.range, w.wildcard, opts.mode);
        }
    } catch (e) {
        if (!usedManual) {          // 主路径异常 → 手动循环回退
            try { manualApplyVertAlign(scope.range, w.wildcard, opts); }
            catch (e2) { alert("设置过程出现异常：" + (e2 && e2.message ? e2.message : e2)); }
        } else {
            alert("设置过程出现异常：" + (e && e.message ? e.message : e));
        }
    } finally {
        try { app.ScreenUpdating = true; } catch (e) {}
        if (began) tmEndUndo(app);
    }

    var modeName = opts.mode === "super" ? "上标" : (opts.mode === "sub" ? "下标" : "取消上下标");
    alert("已在" + scope.label + "内完成「" + modeName + "」，共 " + total + " 处。\n可用 Ctrl+Z 撤销本次修改。");
    onDone(true);
}

// 主路径：仅格式替换（文字不变，命中片段整体设为指定格式），与原生「全部替换」同为一步撤销。
function formatReplaceAll(scope, wildcard, mode) {
    var rng = scope.Duplicate;
    var f = rng.Find;
    f.ClearFormatting();
    f.Replacement.ClearFormatting();
    f.Replacement.Text = "^&";                 // ^& = 命中的原文（仅改格式，文字不变）
    if (mode === "super") f.Replacement.Font.Superscript = true;
    if (mode === "sub")   f.Replacement.Font.Subscript = true;
    // Execute 参数全部显式传，避免「预设属性 vs 参数缺省」的引擎差异：
    // (FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms,
    //  Forward, Wrap, Format, ReplaceWith, Replace)  —— Wrap=0 wdFindStop，Replace=2 wdReplaceAll
    f.Execute(wildcard, true, false, true, false, false, true, 0, true, "^&", 2);
}

// 回退路径：手动循环逐个命中设置格式（支持「去首尾各 1 字符」与「取消上下标」）。
function manualApplyVertAlign(scope, wildcard, opts) {
    tmForEachMatch(scope, wildcard, function (hit) {
        var r;
        try { r = hit.Duplicate; } catch (e) { return; }
        if (opts.trimEnds) {
            var len = 0;
            try { len = r.End - r.Start; } catch (e) {}
            if (len > 2) { r.Start = r.Start + 1; r.End = r.End - 1; }
            else return;   // 命中太短（≤2 字符），去首尾后为空，跳过
        }
        try {
            if (opts.mode === "super") r.Font.Superscript = true;
            else if (opts.mode === "sub") r.Font.Subscript = true;
            else { r.Font.Superscript = false; r.Font.Subscript = false; }
        } catch (e) {}
    });
}
