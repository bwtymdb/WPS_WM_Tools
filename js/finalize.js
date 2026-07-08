// finalize.js —— 「清理并另存为」核心逻辑
// 把当前 Word 快照另存为新文件，按用户选项：接受修订 / 删批注 / 删图片（可选类型）；
// 且原文件从不被写入、其修订状态保持不变。

// —— 形状类型常量（已在 WPS 12.x COM 验证：内嵌图片=3，浮动图片=13）——
var PIC_INLINE = { 3: 1, 4: 1, 9: 1, 10: 1, 11: 1 }; // wdInlineShape 图片类
var PIC_SHAPE  = { 13: 1, 11: 1 };                     // msoPicture / msoLinkedPicture
var MSO_GROUP  = 6;
var HDR_PRIMARY = 1, HDR_FIRST = 2, HDR_EVEN = 3;

// opts: { acceptRevisions, deleteComments, deletePictures,
//         pic:{ inline, floating, headerFooter, inTextbox } }
// 返回 true=已生成；false=未执行（无文档/用户取消）
function SaveFinalVersion(opts) {
    var app = window.Application;
    var doc = app.ActiveDocument;
    if (!doc) { alert("请先打开一个 Word 文档。"); return false; }

    var originalPath = doc.FullName || "";

    // 1) 让用户确认输出路径（默认 原目录\原名_清理版.ext）
    var outPath = pickOutputPath(app, originalPath, doc.Name);
    if (!outPath) return false;

    // 2) 把当前文档快照到 outPath —— 原文件从不被写入
    doc.SaveAs2(outPath);

    // 3) 按选项清理：先关修订记录，确保删除是真实删除
    try {
        doc.TrackRevisions = false;
        if (opts.acceptRevisions) acceptAllRevisions(doc);
        if (opts.deleteComments)  deleteAllComments(doc);
        if (opts.deletePictures)  deleteAllPictures(doc, opts.pic || {});
    } catch (e) {
        alert("清理过程出现异常（新版本仍已另存，可能未完全清理）：" + (e && e.message ? e.message : e));
    }

    // 4) 保存新版本
    try { doc.Save(); } catch (e) {}

    // 5) 重新打开原文件，保持其修订状态不变（原文件从未被写入）
    if (originalPath && /[\\\/]/.test(originalPath)) {
        try { app.Documents.Open(originalPath); } catch (e) {}
    }

    // 6) 把新版本窗口置前
    activateByPath(app, outPath);
    alert("已生成清理后的版本：\n" + outPath);
    return true;
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
