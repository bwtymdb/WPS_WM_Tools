// supsub_panel.js —— 「上下标批量设置」配置对话框逻辑
// 默认配置用 PluginStorage 持久化（key: WMTools.Supsub.defaults），下次自动套用。

var STORAGE_KEY = "WMTools.Supsub.defaults";

var DEFAULTS = {
    preset: 1,                       // 1=方括号引用 2=圆括号引用 3=自定义
    customPattern: "[{一串数字}]",
    mode: "super",                   // super=上标 sub=下标 clear=取消上下标
    trimEnds: false                  // 设置时不包含命中文字首尾各 1 个字符
};

function getDefaults() {
    try {
        var s = window.Application.PluginStorage.getItem(STORAGE_KEY);
        if (s) {
            var o = JSON.parse(s);
            return {
                preset: (o.preset === 2 || o.preset === 3) ? o.preset : 1,
                customPattern: typeof o.customPattern === "string" ? o.customPattern : DEFAULTS.customPattern,
                mode: (o.mode === "sub" || o.mode === "clear") ? o.mode : "super",
                trimEnds: o.trimEnds === true
            };
        }
    } catch (e) {}
    return DEFAULTS;
}

function saveDefaults(o) {
    try { window.Application.PluginStorage.setItem(STORAGE_KEY, JSON.stringify(o)); } catch (e) {}
}

function isChecked(id) { var el = document.getElementById(id); return !!(el && el.checked); }

function gatherOpts() {
    var preset = 1;
    if (isChecked("preset2")) preset = 2;
    else if (isChecked("preset3")) preset = 3;
    var mode = "super";
    if (isChecked("modeSub")) mode = "sub";
    else if (isChecked("modeClear")) mode = "clear";
    var ta = document.getElementById("customPattern");
    return {
        preset: preset,
        customPattern: ta ? ta.value : "",
        mode: mode,
        trimEnds: isChecked("trimEnds")
    };
}

// —— 界面联动 ——
function syncPresetState() {
    var custom = isChecked("preset3");
    var ta = document.getElementById("customPattern");
    if (ta) ta.disabled = !custom;
    var btns = document.getElementById("tokenBtns");
    if (btns) {
        var bs = btns.getElementsByTagName("button");
        for (var i = 0; i < bs.length; i++) bs[i].disabled = !custom;
    }
    updateWildcardPreview();
}

// 实时显示「实际使用的 WPS 通配符」与提示
function updateWildcardPreview() {
    var pv = document.getElementById("wildcardPreview");
    var warn = document.getElementById("patternWarn");
    if (!pv) return;
    var opts = gatherOpts();
    var w = resolveSupsubWildcard(opts);   // supsub.js 提供
    pv.textContent = w.ok ? w.wildcard : "（模式无效）";
    if (warn) warn.textContent = (w.warnings && w.warnings.length) ? ("提示：" + w.warnings.join("；")) : "";
}

function refreshScopeLabel() {
    var el = document.getElementById("scopeLabel");
    if (!el) return;
    try {
        var s = tmGetScope(window.Application);
        el.textContent = s ? s.label : "（未打开文档）";
    } catch (e) { el.textContent = "（未知）"; }
}

// 在自定义模式输入框的光标处插入占位符
function insertToken(token) {
    var ta = document.getElementById("customPattern");
    if (!ta || ta.disabled) return;
    var start = ta.selectionStart, end = ta.selectionEnd;
    if (typeof start !== "number") { ta.value += token; }
    else {
        ta.value = ta.value.substring(0, start) + token + ta.value.substring(end);
        ta.selectionStart = ta.selectionEnd = start + token.length;
    }
    ta.focus();
    updateWildcardPreview();
}

// 「先数一下」：只统计不修改
function onCount() {
    if (!window.Application.ActiveDocument) { alert("请先在 WPS 中打开一个 Word 文档。"); return; }
    var r = CountSupsubMatches(gatherOpts());   // supsub.js 提供
    var el = document.getElementById("countResult");
    if (!el) return;
    if (!r.ok) { el.textContent = r.error; return; }
    var tip = r.scopeLabel + "内将命中 " + r.count + " 处";
    if (r.warnings && r.warnings.length) tip += "（" + r.warnings.join("；") + "）";
    el.textContent = tip;
}

function onExecute() {
    if (!window.Application.ActiveDocument) {
        alert("请先在 WPS 中打开一个 Word 文档。");
        return;
    }
    var opts = gatherOpts();
    if (opts.preset === 3 && !opts.customPattern.replace(/\s+/g, "")) {
        alert("自定义模式为空，请输入要匹配的内容。");
        return;
    }
    saveDefaults(opts);              // 记住本次选择，作为下次默认
    var btn = document.getElementById("btnExecute");
    var oldText = btn ? btn.textContent : "";
    if (btn) { btn.disabled = true; btn.textContent = "处理中…"; } // 防止处理期间重复点击
    try {
        ApplyVertAlign(opts, function (ok) {   // supsub.js 提供
            if (ok) { closeDialog(); return; }
            if (btn) { btn.disabled = false; btn.textContent = oldText; } // 取消/失败则恢复供重试
        });
    } catch (e) {
        alert("执行失败：" + (e && e.message ? e.message : e));
        if (btn) { btn.disabled = false; btn.textContent = oldText; }
    }
}

function closeDialog() { try { window.close(); } catch (e) {} }

window.onload = function () {
    var d = getDefaults();
    document.getElementById("preset1").checked = (d.preset === 1);
    document.getElementById("preset2").checked = (d.preset === 2);
    document.getElementById("preset3").checked = (d.preset === 3);
    document.getElementById("customPattern").value = d.customPattern;
    document.getElementById("modeSuper").checked = (d.mode === "super");
    document.getElementById("modeSub").checked = (d.mode === "sub");
    document.getElementById("modeClear").checked = (d.mode === "clear");
    document.getElementById("trimEnds").checked = d.trimEnds;

    // 占位符插入按钮（数据来自 textmod.js 的 TM_TOKENS，与说明表同源）
    var box = document.getElementById("tokenBtns");
    if (box && typeof TM_TOKENS !== "undefined") {
        for (var i = 0; i < TM_TOKENS.length; i++) {
            (function (t) {
                var b = document.createElement("button");
                b.type = "button";
                b.className = "token";
                b.textContent = t.token;
                b.title = t.meaning;
                b.onclick = function () { insertToken(t.token); };
                box.appendChild(b);
            })(TM_TOKENS[i]);
        }
    }

    document.getElementById("preset1").onchange = syncPresetState;
    document.getElementById("preset2").onchange = syncPresetState;
    document.getElementById("preset3").onchange = syncPresetState;
    document.getElementById("customPattern").oninput = updateWildcardPreview;

    refreshScopeLabel();
    syncPresetState();
    document.getElementById("btnCount").onclick = onCount;
    document.getElementById("btnExecute").onclick = onExecute;
    document.getElementById("btnCancel").onclick = closeDialog;
};
