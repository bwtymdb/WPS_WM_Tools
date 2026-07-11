// finalize.js —— 「清理并另存为」核心逻辑
// 把当前 Word 快照另存为新文件，按用户选项：接受修订 / 删批注 / 删图片（可选类型）；
// 且原文件从不被写入、其修订状态保持不变。

// —— 形状类型常量（已在 WPS 12.x COM 验证：内嵌图片=3，浮动图片=13）——
var PIC_INLINE = { 3: 1, 4: 1, 9: 1, 10: 1, 11: 1 }; // wdInlineShape 图片类
var PIC_SHAPE  = { 13: 1, 11: 1 };                     // msoPicture / msoLinkedPicture
var MSO_GROUP  = 6;
var HDR_PRIMARY = 1, HDR_FIRST = 2, HDR_EVEN = 3;

// opts: { acceptRevisions, deleteComments, deletePictures,
//         pic:{ inline, floating, headerFooter, inTextbox }, toMarkdown }
// onDone(ok)：流程结束或用户取消后回调。ok=true 可关闭对话框；false 表示用户取消（留对话框供重试）。
function SaveFinalVersion(opts, onDone) {
    var app = window.Application;
    var doc = app.ActiveDocument;
    if (!doc) { alert("请先打开一个 Word 文档。"); onDone(false); return; }

    var originalPath = doc.FullName || "";
    var toMd = !!opts.toMarkdown;

    var docxPath, mdPath;
    if (toMd) {
        // md 模式：原文件必须已存盘（用于派生输出路径与快照）。
        // 不弹 Word「另存为」对话框——它只能选 Word 格式（.doc/.docx），无法选 .md，会误导用户。
        if (!originalPath || !/[\\\/]/.test(originalPath)) {
            alert("请先将文档保存到磁盘，再导出 Markdown。");
            onDone(false); return;
        }
        var stem = cleanedStem(originalPath);            // 去掉已存在的 _清理版 后缀，避免叠加
        var dir = dirOf(originalPath);
        mdPath = dir + "\\" + stem + "_清理版.md";
        docxPath = dir + "\\" + stem + "_清理版.docx";   // 中间 .docx（强制 docx 格式，mammoth 才能读）
    } else {
        // 非 md：弹「另存为」选位置，按所选格式保存
        var outPath = pickOutputPath(app, originalPath, doc.Name);
        if (!outPath) { onDone(false); return; }
        docxPath = outPath;
        mdPath = null;
    }

    // 快照到 docxPath —— 原文件从不被写入。md 模式强制 .docx（12=wdFormatXMLDocument），
    // 即便原文是 .doc 也转成 .docx，否则 mammoth 报“Could not find main document part”。
    if (toMd) { doc.SaveAs2(docxPath, 12); } else { doc.SaveAs2(docxPath); }

    // 按选项清理：先关修订记录，确保删除是真实删除
    try {
        doc.TrackRevisions = false;
        if (opts.acceptRevisions) acceptAllRevisions(doc);
        if (opts.deleteComments)  deleteAllComments(doc);
        if (opts.deletePictures)  deleteAllPictures(doc, opts.pic || {});
    } catch (e) {
        alert("清理过程出现异常（新版本仍已另存，可能未完全清理）：" + (e && e.message ? e.message : e));
    }
    try { doc.Save(); } catch (e) {}

    // 非 md：重开并激活原文件，再关掉清理版窗口——界面停在原文件，不自动打开清理版
    if (!toMd) {
        restoreOriginal(app, originalPath, doc);
        alert("已生成清理后的版本：\n" + docxPath + "\n（未自动打开，界面停在原文件）");
        onDone(true);
        return;
    }

    // md：遍历文档对象模型生成 Markdown → 写 .md → 关闭并删除中间 docx，只留 .md。
    // 直接读 doc 对象（不读文件字节、不用 mammoth，避免在 WPS 网页引擎里卡死）；
    // 不 Documents.Open 重开原文件（0 文档时从对话框重开会卡死）——原文件未改动，可自行重开。
    var mdText = "";
    try { mdText = docToMarkdown(doc); }
    catch (e) {
        restoreOriginal(app, originalPath, doc);
        alert("生成 Markdown 失败：" + (e && e.message ? e.message : e) + "\n\n已保留清理版 docx：\n" + docxPath);
        onDone(true); return;
    }

    try { writeMd(app, mdPath, mdText); }
    catch (e) {
        restoreOriginal(app, originalPath, doc);
        alert("写入 .md 失败：" + (e && e.message ? e.message : e) + "\n\n已保留清理版 docx：\n" + docxPath);
        onDone(true); return;
    }

    // .md 已写好：重开并激活原文件（中间件仍开着，2 文档，不卡），再关中间件、删中间件
    restoreOriginal(app, originalPath, doc);
    try { removeFile(docxPath); } catch (e) {}

    alert("已生成 Markdown：\n" + mdPath + "\n（未自动打开，界面停在原文件）");
    onDone(true);
}

// —— 选择输出路径：循环直到与原文件不同，或用户取消 ——
function pickOutputPath(app, originalPath, docName) {
    while (true) {
        var fd = app.FileDialog(2); // msoFileDialogSaveAs
        try { fd.InitialFileName = buildDefaultName(originalPath, docName); } catch (e) {}
        var shown = -1;
        try { shown = fd.Show(); } catch (e) { return null; }
        if (!shown) return null;
        var out;
        try { out = fd.SelectedItems.Item(1); } catch (e) { return null; }
        if (!out) return null;
        if (samePath(out, originalPath)) {
            alert("不能覆盖原文件，请选择其它路径或文件名。");
            continue;
        }
        return out;
    }
}

function buildDefaultName(originalPath, docName) {
    var dir = "", base = docName || "文档";
    if (originalPath && /[\\\/]/.test(originalPath)) {
        var slash = Math.max(originalPath.lastIndexOf("\\"), originalPath.lastIndexOf("/"));
        dir = originalPath.substring(0, slash);
        base = originalPath.substring(slash + 1);
    }
    var dot = base.lastIndexOf(".");
    var stem = dot > 0 ? base.substring(0, dot) : base;
    var ext = dot > 0 ? base.substring(dot) : ".docx";
    var name = stem + "_清理版" + ext;
    return dir ? (dir + "\\" + name) : name;
}

// —— 接受所有修订 ——
function acceptAllRevisions(doc) {
    try { doc.AcceptAllRevisions(); } catch (e) {}
}

// —— 删除所有批注 ——
function deleteAllComments(doc) {
    if (typeof doc.DeleteAllComments === "function") {
        try { doc.DeleteAllComments(); return; } catch (e) {}
    }
    var cs = doc.Comments;
    for (var i = cs.Count; i >= 1; i--) {
        try { cs.Item(i).Delete(); } catch (e) {}
    }
}

// —— 删除图片（按类型）：主文档 + 各节页眉/页脚 + 文本框/组合内（递归）——
// p: { inline, floating, headerFooter, inTextbox }
function deleteAllPictures(doc, p) {
    p = p || {};
    if (p.inline) cleanInline(doc.InlineShapes);
    traverseShapes(doc.Shapes, p.floating, p);

    var secs = doc.Sections;
    for (var s = 1; s <= secs.Count; s++) {
        var sec = secs.Item(s);
        eachHeaderFooter(sec.Headers, function (hf) {
            if (p.headerFooter) { try { cleanInline(hf.Range.InlineShapes); } catch (e) {} }
            try { traverseShapes(hf.Shapes, p.headerFooter, p); } catch (e) {}
        });
        eachHeaderFooter(sec.Footers, function (hf) {
            if (p.headerFooter) { try { cleanInline(hf.Range.InlineShapes); } catch (e) {} }
            try { traverseShapes(hf.Shapes, p.headerFooter, p); } catch (e) {}
        });
    }
}

function eachHeaderFooter(col, fn) {
    var idxs = [HDR_PRIMARY, HDR_FIRST, HDR_EVEN];
    for (var k = 0; k < idxs.length; k++) {
        var hf;
        try { hf = col.Item(idxs[k]); } catch (e) { continue; }
        if (hf) fn(hf);
    }
}

//删除集合中类型为图片的 InlineShape（保留图表/OLE/SmartArt）
function cleanInline(shapes) {
    if (!shapes) return;
    for (var i = shapes.Count; i >= 1; i--) {
        var sh, t;
        try { sh = shapes.Item(i); t = sh.Type; } catch (e) { continue; }
        if (PIC_INLINE[t]) { try { sh.Delete(); } catch (e) {} }
    }
}

//floatingFlag: 当前上下文中是否删除"直接图片"（正文=floating，页眉页脚=headerFooter，文本框内=inTextbox）
function traverseShapes(shapes, floatingFlag, p) {
    if (!shapes) return;
    for (var i = shapes.Count; i >= 1; i--) {
        var sh;
        try { sh = shapes.Item(i); } catch (e) { continue; }
        visitShape(sh, floatingFlag, p);
    }
}

function visitShape(sh, floatingFlag, p) {
    var t;
    try { t = sh.Type; } catch (e) { return; }
    if (PIC_SHAPE[t]) { if (floatingFlag) { try { sh.Delete(); } catch (e) {} } return; }
    if (t === MSO_GROUP) {                                   // 组合：按同一上下文递归子形状
        try {
            var gi = sh.GroupItems;
            for (var i = gi.Count; i >= 1; i--) { try { visitShape(gi.Item(i), floatingFlag, p); } catch (e) {} }
        } catch (e) {}
    }
    if (p.inTextbox) {                                       // 文本框等带文字形状：递归删除其文字范围里的图片
        try {
            var tr = sh.TextFrame.TextRange;
            if (tr) { cleanInline(tr.InlineShapes); traverseShapes(tr.Shapes, p.inTextbox, p); }
        } catch (e) {}
    }
}

// —— 工具 ——
function samePath(a, b) {
    if (!a || !b) return false;
    function norm(p) { return String(p).toLowerCase().replace(/\//g, "\\"); }
    return norm(a) === norm(b);
}

function activateByPath(app, path) {
    try {
        var docs = app.Documents;
        for (var i = 1; i <= docs.Count; i++) {
            if (samePath(docs.Item(i).FullName, path)) { docs.Item(i).Activate(); return; }
        }
    } catch (e) {}
}

// 导出后把界面停回原文件：先 Open 原文件（中间文档仍开着，2 文档，不会卡死），激活它，再关掉中间文档。
// 顺序关键——必须「Open 在前、Close 在后」，不能 0 文档时从对话框 Open（会卡死）。
function restoreOriginal(app, originalPath, docToClose) {
    var orig = null;
    if (originalPath && /[\\\/]/.test(originalPath)) {
        try { orig = app.Documents.Open(originalPath); } catch (e) {}
    }
    if (orig) { try { orig.Activate(); } catch (e) {} }
    if (docToClose) { try { docToClose.Close(0); } catch (e) {} }
}

// ============== Markdown 导出（纯 JS，不依赖 Python / mammoth） ==============
// 直接遍历 WPS 文档对象模型（Paragraphs 的 OutlineLevel 判标题、Range.Tables 判表格）拼出 Markdown，
// 再用 window.Application.FileSystem 写 .md。不读文件字节、不用 mammoth（在 WPS 网页引擎里会卡死）。
// v1 暂不还原列表与行内加粗/斜体（WPS 里检测不稳），后续可加。

// 取路径所在目录（不含末尾分隔符）
function dirOf(p) {
    var slash = Math.max(p.lastIndexOf("\\"), p.lastIndexOf("/"));
    return slash >= 0 ? p.substring(0, slash) : "";
}

// 取文件名主干（去扩展名；并去掉已存在的 _清理版 后缀，避免重复运行时叠加成 _清理版_清理版）
function cleanedStem(p) {
    var slash = Math.max(p.lastIndexOf("\\"), p.lastIndexOf("/"));
    var base = slash >= 0 ? p.substring(slash + 1) : p;
    var dot = base.lastIndexOf(".");
    var stem = dot > 0 ? base.substring(0, dot) : base;
    var suffix = "_清理版";
    if (stem.length >= suffix.length && stem.substring(stem.length - suffix.length) === suffix) {
        stem = stem.substring(0, stem.length - suffix.length);
    }
    return stem;
}

// 遍历 WPS 文档对象模型 → Markdown 字符串。
// 直接读 doc 对象（不读文件字节、不用 mammoth）：OutlineLevel 判标题、Range.Tables 判表格、其余为正文段落。
// （v1 暂不还原列表与行内加粗/斜体——这些在 WPS 里检测不稳，留作后续。）
function docToMarkdown(doc) {
    var out = [];
    var paras = doc.Paragraphs;
    var n = 0;
    try { n = paras.Count; } catch (e) {}
    var emitted = {};   // 已输出的表格，按 Range.Start 去重，避免单元格段落重复
    for (var i = 1; i <= n; i++) {
        var p;
        try { p = paras.Item(i); } catch (e) { continue; }

        // 表格内的段落：整张表只输出一次，跳过各单元格段落
        var rt = null;
        try { rt = p.Range.Tables; } catch (e) {}
        if (rt && rt.Count > 0) {
            var key = -1, t = null;
            try { t = rt.Item(1); key = t.Range.Start; } catch (e) {}
            if (key >= 0 && !emitted[key]) {
                emitted[key] = true;
                var tm = "";
                try { tm = tableToMd(t); } catch (e) {}
                if (tm) out.push(tm);
            }
            continue;
        }

        // 标题：OutlineLevel 1..6
        var ol = 10;
        try { ol = p.OutlineLevel; } catch (e) {}
        if (ol >= 1 && ol <= 6) {
            var ht = paraText(p);
            if (ht) out.push("\n\n" + repeat("#", ol) + " " + ht + "\n\n");
            continue;
        }

        // 正文段落
        var tx = paraText(p);
        if (tx) out.push("\n\n" + tx + "\n\n");
    }
    var md = out.join("");
    return md.replace(/\n{3,}/g, "\n\n").replace(/^\s+/, "").replace(/\s+$/, "") + "\n";
}

// 段落纯文本（去段尾回车/制表符/单元格结束符，压缩空白）
function paraText(p) {
    var t = "";
    try { t = p.Range.Text; } catch (e) { return ""; }
    return String(t).replace(/[\r\n\t]/g, "").replace(/ /g, " ").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
}

// 表格 → GFM 表格
function tableToMd(t) {
    var rows = t.Rows;
    var rc = 0;
    try { rc = rows.Count; } catch (e) {}
    var lines = [];
    var headerDone = false;
    for (var r = 1; r <= rc; r++) {
        var row;
        try { row = rows.Item(r); } catch (e) { continue; }
        var cells = row.Cells;
        var cc = 0;
        try { cc = cells.Count; } catch (e) {}
        var vals = [];
        for (var c = 1; c <= cc; c++) {
            var ct = "";
            try { ct = cells.Item(c).Range.Text; } catch (e) {}
            vals.push(String(ct).replace(/[\r\n\t]/g, "").replace(/ /g, " ").replace(/\|/g, "\\|").replace(/^\s+|\s+$/g, ""));
        }
        if (!vals.length) continue;
        lines.push("| " + vals.join(" | ") + " |");
        if (!headerDone) {
            var sep = [];
            for (var k = 0; k < vals.length; k++) sep.push("---");
            lines.push("| " + sep.join(" | ") + " |");
            headerDone = true;
        }
    }
    return lines.length ? "\n\n" + lines.join("\n") + "\n\n" : "";
}

// 写 .md 到 OS 路径。FileSystem 是沙箱的（拒含 / \ 的路径），故：
// 1) 首选 ActiveX ADODB.Stream（写 UTF-8 无 BOM，干净、不闪新文档）——引擎支持 ActiveX 时走这条；
// 2) 回退 WPS COM SaveAs2（已验证能写 OS 路径）：新文档 → Content.Text → 另存为 UTF-8 文本。
function writeMd(app, path, text) {
    try { writeUtf8ViaAdodb(path, text); return; } catch (e) {}
    writeMdViaCom(app, path, text);
}

// ActiveX ADODB.Stream：写 UTF-8 无 BOM（文本流写出后，跳过 3 字节 BOM 拷贝到二进制流再落盘）
function writeUtf8ViaAdodb(path, text) {
    var s = new ActiveXObject("ADODB.Stream");
    s.Type = 2;            // adTypeText
    s.Charset = "utf-8";
    s.Open();
    s.WriteText(text);
    var b = new ActiveXObject("ADODB.Stream");
    b.Type = 1;            // adTypeBinary
    b.Open();
    s.Position = 3;        // 跳过 UTF-8 BOM（EF BB BF）
    s.CopyTo(b);
    b.SaveToFile(path, 2); // adSaveCreateOverWrite
    b.Close();
    s.Close();
}

// COM 回退：用 WPS 新建文档承载 md 文本，SaveAs2 为 UTF-8 纯文本（Content.Text 赋值不触发自动套用格式）
function writeMdViaCom(app, path, text) {
    var newDoc = null;
    try { app.ScreenUpdating = false; } catch (e) {}
    try {
        newDoc = app.Documents.Add();
        newDoc.Content.Text = text;
        newDoc.SaveAs2(path, 2, false, "", false, "", false, false, false, false, false, 65001); // 2=wdFormatText, 65001=UTF-8
    } finally {
        try { if (newDoc) newDoc.Close(0); } catch (e) {}
        try { app.ScreenUpdating = true; } catch (e) {}
    }
}

// 删除中间文件（OS 路径）。FileSystem 多半也是沙箱的，故再回退 ActiveX FSO；都失败则保留。
function removeFile(path) {
    var fs;
    try { fs = window.Application.FileSystem; } catch (e) { fs = null; }
    if (fs) {
        try { if (typeof fs.unlinkSync === "function") { fs.unlinkSync(path); return; } } catch (e) {}
        try { if (typeof fs.Remove === "function") { fs.Remove(path); return; } } catch (e) {}
    }
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        if (fso.FileExists(path)) fso.DeleteFile(path, true);
    } catch (e) {}
}

function repeat(str, n) { var r = ""; for (var i = 0; i < n; i++) r += str; return r; }
