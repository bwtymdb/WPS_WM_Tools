// markdown_panel.js —— 「Markdown 排版」配置对话框逻辑
// 打开时自动检测作用域内的 Markdown 元素并显示数量；勾选持久化到 PluginStorage。

var STORAGE_KEY = "WMTools.Markdown.defaults";

var DEFAULTS = {
    headings: true, headingMode: "style",
    ul: true, ol: true, table: true, codeBlock: true, quote: true, hr: true,
    bold: true, italic: true, strike: true, inlineCode: true, link: true
};

// 统计字段 → 面板显示名（与 markdown.js 的 stats/counts 键一致）
var STAT_NAMES = [
    ["headings", "标题"], ["bold", "加粗"], ["italic", "斜体"], ["strike", "删除线"],
    ["inlineCode", "行内代码"], ["codeBlock", "代码块"], ["table", "表格"],
    ["ul", "无序列表项"], ["ol", "有序列表项"], ["quote", "引用段"], ["hr", "分割线"], ["link", "超链接"]
];

var CHECK_IDS = {
    optHeadings: "headings", optUl: "ul", optOl: "ol", optTable: "table",
    optCodeBlock: "codeBlock", optQuote: "quote", optHr: "hr",
    optBold: "bold", optItalic: "italic", optStrike: "strike",
    optInlineCode: "inlineCode", optLink: "link"
};

function getDefaults() {
    try {
        var s = window.Application.PluginStorage.getItem(STORAGE_KEY);
        if (s) {
            var o = JSON.parse(s);
            var d = {};
            for (var k in DEFAULTS) d[k] = (o[k] !== undefined) ? o[k] : DEFAULTS[k];
            if (d.headingMode !== "format") d.headingMode = "style";
            return d;
        }
    } catch (e) {}
    return DEFAULTS;
}

function saveDefaults(o) {
    try { window.Application.PluginStorage.setItem(STORAGE_KEY, JSON.stringify(o)); } catch (e) {}
}

function isChecked(id) { var el = document.getElementById(id); return !!(el && el.checked); }

function gatherOpts() {
    var o = { headingMode: document.getElementById("modeFormat").checked ? "format" : "style" };
    for (var id in CHECK_IDS) o[CHECK_IDS[id]] = isChecked(id);
    return o;
}

function refreshAnalysis() {
    var scopeEl = document.getElementById("scopeLabel");
    var statsEl = document.getElementById("statsBox");
    try {
        var r = AnalyzeMarkdown();   // markdown.js 提供
        if (!r.ok) {
            scopeEl.textContent = "（未打开文档）";
            statsEl.textContent = r.error || "无法检测。";
            return;
        }
        scopeEl.textContent = r.label;
        var parts = [];
        for (var i = 0; i < STAT_NAMES.length; i++) {
            var key = STAT_NAMES[i][0];
            if (r.stats[key]) parts.push(STAT_NAMES[i][1] + " " + r.stats[key] + " 处");
        }
        var text = parts.length ? ("检测到：" + parts.join("、")) : "未检测到 Markdown 元素（直接执行也不会改动文档）。";
        if (r.stats.tableSkip) text += "\n另有 " + r.stats.tableSkip + " 个疑似表格无法解析，将保留原文。";
        statsEl.textContent = text;
    } catch (e) {
        statsEl.textContent = "检测失败：" + (e && e.message ? e.message : e);
    }
}

function onExecute() {
    if (!window.Application.ActiveDocument) {
        alert("请先在 WPS 中打开一个 Word 文档。");
        return;
    }
    var opts = gatherOpts();
    var any = false;
    for (var k in CHECK_IDS) { if (opts[CHECK_IDS[k]]) { any = true; break; } }
    if (!any) { alert("请至少勾选一种要转换的元素。"); return; }
    saveDefaults(opts);              // 记住本次选择，作为下次默认
    var btn = document.getElementById("btnExecute");
    var oldText = btn ? btn.textContent : "";
    if (btn) { btn.disabled = true; btn.textContent = "处理中…"; }
    try {
        ConvertMarkdown(opts, function (ok) {   // markdown.js 提供
            if (ok) { closeDialog(); return; }
            if (btn) { btn.disabled = false; btn.textContent = oldText; }
        });
    } catch (e) {
        alert("执行失败：" + (e && e.message ? e.message : e));
        if (btn) { btn.disabled = false; btn.textContent = oldText; }
    }
}

function closeDialog() { try { window.close(); } catch (e) {} }

window.onload = function () {
    var d = getDefaults();
    for (var id in CHECK_IDS) document.getElementById(id).checked = d[CHECK_IDS[id]];
    document.getElementById("modeStyle").checked = (d.headingMode !== "format");
    document.getElementById("modeFormat").checked = (d.headingMode === "format");
    refreshAnalysis();
    document.getElementById("btnExecute").onclick = onExecute;
    document.getElementById("btnCancel").onclick = closeDialog;
    document.getElementById("btnRescan").onclick = refreshAnalysis;
};
