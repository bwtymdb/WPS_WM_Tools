// md_e2e.js - Markdown end-to-end verification (run: cscript //E:JScript //NoLogo test/md_e2e.js)
// Loads the REAL js/textmod.js + js/markdown.js, runs ConvertMarkdown on a scratch WPS doc.
// Checks: heading style / bold+italic / native lists / table / code literal / hyperlink / single undo.
// IMPORTANT: this file must stay PURE ASCII (Chinese only as \u escapes) - cscript reads .js as ANSI,
// UTF-8 comments get misdecoded and silently corrupt array literals (hard-learned lesson).
// Never touches user documents; quits WPS only if no other documents remain open.

var PASS = 0, FAIL = 0;
function check(name, cond) {
    if (cond) { PASS++; WScript.Echo("PASS " + name); }
    else { FAIL++; WScript.Echo("FAIL " + name); }
}
function readFileUtf8(path) {
    var st = new ActiveXObject("ADODB.Stream");
    st.Type = 2; st.Charset = "utf-8"; st.Open(); st.LoadFromFile(path);
    var s = st.ReadText(-1); st.Close(); return s;
}
// escape non-ASCII so console output never turns into mojibake
function esc(s) {
    var out = "";
    for (var i = 0; i < s.length; i++) {
        var c = s.charCodeAt(i);
        if (c >= 32 && c < 127) out += s.charAt(i);
        else { var h = c.toString(16); while (h.length < 4) h = "0" + h; out += "\\u" + h; }
    }
    return out;
}

var app = null, doc = null;
function find(text) {
    var r = doc.Content.Duplicate;
    var f = r.Find;
    f.ClearFormatting(); f.Text = text; f.Forward = true; f.Wrap = 0;
    return f.Execute() ? r : null;
}

try {
    try { app = GetObject("", "KWps.Application"); } catch (e) {}
    if (!app) { try { app = GetObject("", "Wps.Application"); } catch (e) {} }
    if (!app) app = new ActiveXObject("KWps.Application");
    app.Visible = true;

    var base = "D:\\Claude\\WPS_WM_Tools\\";
    var src = readFileUtf8(base + "js\\textmod.js") + "\n" + readFileUtf8(base + "js\\markdown.js");
    var window = this;
    window.Application = app;
    var alerts = [];
    function alert(m) { alerts.push(String(m)); }
    eval(src);

    doc = app.Documents.Add();
    var lines = [
        "# \u9605\u8bfb\u8d44\u6599",
        "\u8fd9\u662f**\u52a0\u7c97**\u548c*\u659c\u4f53*\u6587\u5b57",
        "- \u7532\u9879\u76ee",
        "- \u4e59\u9879\u76ee",
        "1. \u7b2c\u4e00\u6b65",
        "1. \u7b2c\u4e8c\u6b65",
        "> \u5f15\u7528\u4e00\u53e5",
        "---",
        "| \u540d\u79f0 | \u6570\u91cf |",
        "|---|---:|",
        "| \u82f9\u679c | 3 |",
        "```js",
        "var a = **1**;",
        "```",
        "\u89c1[\u793a\u4f8b](https://example.com)\u5373\u53ef"
    ];
    doc.Content.Text = lines.join("\r");
    var original = doc.Content.Text;
    check("scratch doc has 15 paras", doc.Paragraphs.Count == 15);

    var opts = { headings: true, headingMode: "style", bold: true, italic: true, strike: true,
                 inlineCode: true, codeBlock: true, table: true, ul: true, ol: true,
                 quote: true, hr: true, link: true };
    var done = false;
    var _mdInlineMatches = mdInlineMatches;
    mdInlineMatches = function (t, o) {
        var r = _mdInlineMatches(t, o);
        if (t.indexOf("\u659c\u4f53") >= 0) {
            WScript.Echo("INLINE on: " + esc(t));
            for (var z = 0; z < r.length; z++) WScript.Echo("  M " + r[z].type + " [" + r[z].start + "," + r[z].end + "] cs=" + r[z].contentStart + " ce=" + r[z].contentEnd);
        }
        return r;
    };
    var _mdApplyInline = mdApplyInline;
    mdApplyInline = function (d, ps, m, c) {
        WScript.Echo("APPLY " + m.type + " pStart=" + ps + " [" + m.start + "," + m.end + "]");
        _mdApplyInline(d, ps, m, c);
    };
    ConvertMarkdown(opts, function (ok) { done = ok; });
    check("ConvertMarkdown completed", done === true);
    for (var a = 0; a < alerts.length; a++) WScript.Echo("ALERT: " + esc(alerts[a]));

    var txt = doc.Content.Text;
    WScript.Echo("TEXT_AFTER: " + esc(txt));

    check("heading # removed", txt.indexOf("# \u9605\u8bfb") < 0);
    check("heading style applied", String(doc.Paragraphs.Item(1).Range.Style.NameLocal).indexOf("\u6807\u9898") >= 0);

    var rb = find("\u52a0\u7c97");
    check("bold text found", !!rb);
    if (rb) check("bold format", rb.Font.Bold != 0);
    var ri = find("\u659c\u4f53");
    check("italic format", !!ri && ri.Font.Italic != 0);
    var p2 = doc.Paragraphs.Item(2).Range.Text;
    check("inline markers removed", p2.indexOf("*") < 0);

    var rl = find("\u7532\u9879\u76ee");
    check("bullet item1 ListType=2", !!rl && rl.ListFormat.ListType == 2);
    var rl2 = find("\u4e59\u9879\u76ee");
    check("bullet item2 ListType=2", !!rl2 && rl2.ListFormat.ListType == 2);
    check("bullet markers stripped", !!rl && rl.Paragraphs.Item(1).Range.Text.indexOf("- ") < 0);
    var n1 = find("\u7b2c\u4e00\u6b65");
    var n2 = find("\u7b2c\u4e8c\u6b65");
    check("ol starts at 1.", !!n1 && n1.ListFormat.ListString == "1.");
    check("ol continues 2.", !!n2 && n2.ListFormat.ListString == "2.");
    check("ol markers stripped", !!n1 && n1.Paragraphs.Item(1).Range.Text.indexOf("1. ") < 0);

    var rq = find("\u5f15\u7528\u4e00\u53e5");
    check("quote left border", !!rq && rq.Paragraphs.Item(1).Borders.Item(-2).LineStyle == 1);
    check("quote prefix stripped", !!rq && rq.Paragraphs.Item(1).Range.Text.indexOf("> ") < 0);
    check("hr text removed", txt.indexOf("---") < 0);

    check("table inserted", doc.Tables.Count == 1);
    if (doc.Tables.Count == 1) {
        var tbl = doc.Tables.Item(1);
        check("table cell filled", tbl.Cell(2, 1).Range.Text.indexOf("\u82f9\u679c") >= 0);
        check("table header bold", tbl.Cell(1, 1).Range.Font.Bold != 0);
        check("table right align", tbl.Cell(2, 2).Range.ParagraphFormat.Alignment == 2);
        check("table borders", tbl.Borders.InsideLineStyle == 1);
    }

    var rc = find("var a");
    check("code font Consolas", !!rc && rc.Font.NameAscii == "Consolas");
    check("code literal ** kept", !!rc && rc.Paragraphs.Item(1).Range.Text.indexOf("**1**") >= 0);
    check("code fence removed", txt.indexOf("```") < 0);

    check("hyperlink added", doc.Hyperlinks.Count == 1);
    if (doc.Hyperlinks.Count == 1) {
        check("hyperlink address", doc.Hyperlinks.Item(1).Address == "https://example.com");
        check("link text kept", txt.indexOf("\u793a\u4f8b") >= 0);
        check("link syntax removed", txt.indexOf("](") < 0);
    }

    doc.Undo(1);
    check("single undo restores all", doc.Content.Text == original);
    check("undo removes table", doc.Tables.Count == 0);
    check("undo removes hyperlink", doc.Hyperlinks.Count == 0);

    // selection scope: select first two paragraphs only (p0 [0,7], p1 [7,23])
    doc.Range(0, 23).Select();
    alerts = [];
    ConvertMarkdown(opts, function (ok) {});
    check("sel: heading converted", String(doc.Paragraphs.Item(1).Range.Style.NameLocal).indexOf("\u6807\u9898") >= 0);
    check("sel: inline converted", doc.Paragraphs.Item(2).Range.Text.indexOf("*") < 0);
    check("sel: bullet outside untouched", doc.Content.Text.indexOf("- \u7532\u9879\u76ee") >= 0);
    check("sel: no table created", doc.Tables.Count == 0);
    check("sel: no hyperlink created", doc.Hyperlinks.Count == 0);
    doc.Undo(1);
    check("sel: undo restores", doc.Content.Text == original);
} catch (e) {
    FAIL++; WScript.Echo("FAIL exception: " + (e && e.message ? e.message : e) + " @" + (e && e.number));
} finally {
    try { if (doc) doc.Close(0); } catch (e) {}
    try { if (app && app.Documents.Count == 0) app.Quit(); } catch (e) {}
}
WScript.Echo("RESULT PASS=" + PASS + " FAIL=" + FAIL);
WScript.Quit(FAIL ? 1 : 0);
