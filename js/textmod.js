// textmod.js —— 文本批量修改公共助手（引号转换 / 上下标批量设置共用）
// 只放与具体工具无关的通用能力：
//   1) 作用域获取（有选中 → 选区；未选中 → 全文主文档）
//   2) 自定义撤销记录（UndoRecord）探测与包裹
//   3) 友好占位符 → WPS 通配符翻译（含字面量自动转义）
//   4) 通配符手动查找循环（ReplaceAll 之外的回退路径）

// —— 1) 作用域 ——
// 返回 { doc, range, isSelection, label }；无文档或失败返回 null
function tmGetScope(app) {
    var doc = app.ActiveDocument;
    if (!doc) return null;
    var rng = null, isSel = false;
    try {
        var sr = app.Selection.Range;
        if (sr && sr.End > sr.Start) { rng = sr.Duplicate; isSel = true; }
    } catch (e) {}
    if (!rng) { try { rng = doc.Content; } catch (e) { rng = null; } }
    if (!rng) return null;
    return { doc: doc, range: rng, isSelection: isSel, label: isSel ? "选中文字" : "全文" };
}

// —— 2) 自定义撤销记录 ——
// WPS 支持 Application.UndoRecord 时，把整批操作合并为一步 Ctrl+Z；
// 不支持则静默退化（操作仍是原生 Range 修改，可逐条撤销）。
function tmBeginUndo(app, label) {
    try { app.UndoRecord.StartCustomRecord(label || "WM_TOOLS 批量修改"); return true; } catch (e) { return false; }
}
function tmEndUndo(app) {
    try { app.UndoRecord.EndCustomRecord(); } catch (e) {}
}

// —— 3) 友好占位符 → WPS 通配符 ——
// 占位符表（也是 ui/supsub.html 说明表的数据来源）
var TM_TOKENS = [
    { token: "{数字}",     wildcard: "[0-9]",     meaning: "任意一个数字",                example: "3、7" },
    { token: "{一串数字}", wildcard: "[0-9]@",    meaning: "一个或多个连续数字",          example: "12、2024" },
    { token: "{字母}",     wildcard: "[a-zA-Z]",  meaning: "任意一个英文字母",            example: "a、Z" },
    { token: "{一串字母}", wildcard: "[a-zA-Z]@", meaning: "一个或多个连续字母",          example: "abc、OK" },
    { token: "{汉字}",     wildcard: "[一-龥]",   meaning: "任意一个汉字",                example: "引、文" },
    { token: "{任意字符}", wildcard: "?",         meaning: "任意一个字符（含标点）",      example: "只匹配一个字" },
    { token: "{任意内容}", wildcard: "*",         meaning: "任意长度内容（可为空，慎用）", example: "容易匹配过多" }
];

function tmLookupToken(name) {
    for (var i = 0; i < TM_TOKENS.length; i++) {
        if (TM_TOKENS[i].token === name) return TM_TOKENS[i];
    }
    return null;
}

// 通配符模式下需要转义的字面字符（全部位于字符类 [] 之外：用户字面量中的 [ ] 会被转义，
// 因此不可能形成字符类，故 - 和 ! 无需转义）
var TM_ESCAPE_RE = /([\\?*\[\](){}@<>^])/g;

function tmEscapeLiterals(s) { return String(s).replace(TM_ESCAPE_RE, "\\$1"); }

// 把「友好写法」翻译成 WPS 通配符。
// 规则：{占位符} 查表替换；其余字符一律视为字面量并自动转义（普通符号原样书写即可）。
// 返回 { ok, wildcard, warnings:[..] }
function tmTranslatePattern(input) {
    var warnings = [];
    if (!input || !String(input).length) return { ok: false, wildcard: "", warnings: ["模式为空"] };
    var out = [];
    var i = 0, n = input.length;
    while (i < n) {
        var ch = input.charAt(i);
        if (ch === "{") {
            var j = input.indexOf("}", i + 1);
            if (j > i) {
                var name = input.substring(i, j + 1);
                var hit = tmLookupToken(name);
                if (hit) { out.push(hit.wildcard); i = j + 1; continue; }
                warnings.push("「" + name + "」不是内置占位符，已按普通文字处理");
                out.push(tmEscapeLiterals(name));
                i = j + 1;
                continue;
            }
            warnings.push("发现没有配对的「{」，已按普通文字处理");
            out.push(tmEscapeLiterals(ch));
            i++;
            continue;
        }
        out.push(tmEscapeLiterals(ch));
        i++;
    }
    var w = out.join("");
    if (!w) return { ok: false, wildcard: "", warnings: ["模式为空"] };
    return { ok: true, wildcard: w, warnings: warnings };
}

// —— 4) 通配符手动查找循环 ——
// 在 scope 内用通配符逐个命中；fn(hitRange) 对每个命中片段回调（只许改格式，不许改文字）。
// 返回命中次数。带安全上限防死循环。
function tmForEachMatch(scope, wildcard, fn) {
    var count = 0;
    var scopeEnd;
    try { scopeEnd = scope.End; } catch (e) { return 0; }
    var rng;
    try { rng = scope.Duplicate; } catch (e) { return 0; }
    var f;
    try { f = rng.Find; } catch (e) { return 0; }
    try {
        f.ClearFormatting();
        f.Text = wildcard;
        f.Forward = true;
        f.Wrap = 0;            // wdFindStop：不越出 scope
        f.Format = false;
        f.MatchCase = true;
        f.MatchWholeWord = false;
        f.MatchWildcards = true;
    } catch (e) { return 0; }
    var guard = 0;
    while (guard++ < 100000) {
        var found = false;
        try { found = f.Execute(); } catch (e) { break; }
        if (!found) break;
        // Execute 命中后 rng 已被重定义为命中片段
        var mEnd;
        try { mEnd = rng.End; } catch (e) { break; }
        if (mEnd > scopeEnd) break;   // 保险：不处理越界命中
        count++;
        if (fn) { try { fn(rng); } catch (e) {} }
        // 收缩到命中之后继续向后找
        try {
            rng.Start = mEnd;
            if (mEnd >= scopeEnd) break;
            rng.End = scopeEnd;
        } catch (e) { break; }
    }
    return count;
}
