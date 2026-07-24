// quotes_panel.js —— 「英文引号转中文」配置对话框逻辑
// 默认配置用 PluginStorage 持久化（key: WMTools.Quotes.defaults），下次自动套用。

var STORAGE_KEY = "WMTools.Quotes.defaults";

var DEFAULTS = { double: true, single: false };

function getDefaults() {
    try {
        var s = window.Application.PluginStorage.getItem(STORAGE_KEY);
        if (s) {
            var o = JSON.parse(s);
            return {
                double: o.double !== false,
                single: o.single === true
            };
        }
    } catch (e) {}
    return DEFAULTS;
}

function saveDefaults(o) {
    try { window.Application.PluginStorage.setItem(STORAGE_KEY, JSON.stringify(o)); } catch (e) {}
}

function isChecked(id) { var el = document.getElementById(id); return !!(el && el.checked); }

function refreshScopeLabel() {
    var el = document.getElementById("scopeLabel");
    if (!el) return;
    try {
        var s = tmGetScope(window.Application);
        el.textContent = s ? s.label : "（未打开文档）";
    } catch (e) { el.textContent = "（未知）"; }
}

function gatherOpts() {
    return { double: isChecked("optDouble"), single: isChecked("optSingle") };
}

function onExecute() {
    if (!window.Application.ActiveDocument) {
        alert("请先在 WPS 中打开一个 Word 文档。");
        return;
    }
    var opts = gatherOpts();
    if (!opts.double && !opts.single) {
        alert("请至少勾选一种引号。");
        return;
    }
    saveDefaults(opts);              // 记住本次选择，作为下次默认
    var btn = document.getElementById("btnExecute");
    var oldText = btn ? btn.textContent : "";
    if (btn) { btn.disabled = true; btn.textContent = "处理中…"; } // 防止处理期间重复点击
    try {
        ConvertStraightQuotes(opts, function (ok) {   // quotes.js 提供
            if (ok) { closeDialog(); return; }
            if (btn) { btn.disabled = false; btn.textContent = oldText; } // 失败则恢复供重试
        });
    } catch (e) {
        alert("执行失败：" + (e && e.message ? e.message : e));
        if (btn) { btn.disabled = false; btn.textContent = oldText; }
    }
}

function closeDialog() { try { window.close(); } catch (e) {} }

window.onload = function () {
    var d = getDefaults();
    document.getElementById("optDouble").checked = d.double;
    document.getElementById("optSingle").checked = d.single;
    refreshScopeLabel();
    document.getElementById("btnExecute").onclick = onExecute;
    document.getElementById("btnCancel").onclick = closeDialog;
};
