// finalize.js —— 核心：把当前 Word 另存为「最终版」
// 新版本中：接受所有修订、删除全部批注、删除全部图片；
// 且原文件从不被写入、其修订状态保持不变。

// —— 形状类型常量（参考 Office 对象模型；WPS 实际值在调试期核对）——
var PIC_INLINE = { 3: 1, 4: 1, 9: 1, 10: 1, 11: 1 }; // wdInlineShape 图片类：图片/链接图片/水平线图片/图片项目符号
var PIC_SHAPE = { 13: 1, 11: 1 };                     // msoPicture / msoLinkedPicture
var MSO_GROUP = 6;                                    // msoGroup（组合，需递归）
var HDR_PRIMARY = 1, HDR_FIRST = 2, HDR_EVEN = 3;     // 页眉页脚类型

function SaveFinalVersion() {
    var app = window.Application;
    var doc = app.ActiveDocument;
    if (!doc) { alert("请先打开一个 Word 文档。"); return; }

    var originalPath = doc.FullName || "";   // 原文件全路径（从未保存则为空）

    // 1) 让用户确认输出路径（默认 原目录\原名_最终版.ext）
    var outPath = pickOutputPath(app, originalPath, doc.Name);
    if (!outPath) return;                    // 用户取消

    // 2) 把当前文档快照到 outPath —— 原文件从不被写入
    doc.SaveAs2(outPath);

    // 3) 处理新版本：先关闭修订记录（否则后续删除会变成"修订"而非真实删除）→ 接受修订 → 删批注 → 删图片
    try {
        doc.TrackRevisions = false;
        acceptAllRevisions(doc);
        deleteAllComments(doc);
        deleteAllPictures(doc);
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

    alert("已生成最终版：\n" + outPath);
}

// —— 选择输出路径：循环直到与原文件不同，或用户取消 ——
function pickOutputPath(app, originalPath, docName) {
    while (true) {
        var fd = app.FileDialog(2); // msoFileDialogSaveAs
        try { fd.InitialFileName = buildDefaultName(originalPath, docName); } catch (e) {}
        var shown = -1;
        try { shown = fd.Show(); } catch (e) { return null; }
        if (!shown) return null;    // 用户取消
        var out;
        try { out = fd.SelectedItems.Item(1); } catch (e) { return null; }
        if (!out) return null;
        if (samePath(out, originalPath)) {
            alert("不能覆盖原文件，请选择其它路径或文件名。");
            continue;               // 重新选择
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
    var name = stem + "_最终版" + ext;
    return dir ? (dir + "\\" + name) : name;
}

// —— 接受所有修订 ——
function acceptAllRevisions(doc) {
    try { doc.AcceptAllRevisions(); } catch (e) {}  // 无修订时为空操作
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

// —— 删除所有图片：主文档 + 各节页眉/页脚 + 文本框/组合内（递归）——
function deleteAllPictures(doc) {
    cleanInline(doc.InlineShapes);
    cleanShapes(doc.Shapes);

    var secs = doc.Sections;
    for (var s = 1; s <= secs.Count; s++) {
        var sec = secs.Item(s);
        eachHeaderFooter(sec.Headers, function (hf) {
            try { cleanInline(hf.Range.InlineShapes); } catch (e) {}
            try { cleanShapes(hf.Shapes); } catch (e) {}
        });
        eachHeaderFooter(sec.Footers, function (hf) {
            try { cleanInline(hf.Range.InlineShapes); } catch (e) {}
            try { cleanShapes(hf.Shapes); } catch (e) {}
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

//删除 InlineShapes 中类型为图片者（保留图表/OLE/SmartArt）
function cleanInline(shapes) {
    if (!shapes) return;
    for (var i = shapes.Count; i >= 1; i--) {
        var sh, t;
        try { sh = shapes.Item(i); t = sh.Type; } catch (e) { continue; }
        if (PIC_INLINE[t]) { try { sh.Delete(); } catch (e) {} }
    }
}

//遍历 Shapes：删除图片，递归组合与文本框
function cleanShapes(shapes) {
    if (!shapes) return;
    for (var i = shapes.Count; i >= 1; i--) {
        var sh;
        try { sh = shapes.Item(i); } catch (e) { continue; }
        cleanOneShape(sh);
    }
}

function cleanOneShape(sh) {
    var t;
    try { t = sh.Type; } catch (e) { return; }
    if (PIC_SHAPE[t]) { try { sh.Delete(); } catch (e) {} return; }
    //组合：递归子形状
    if (t === MSO_GROUP) {
        try {
            var gi = sh.GroupItems;
            for (var i = gi.Count; i >= 1; i--) {
                try { cleanOneShape(gi.Item(i)); } catch (e) {}
            }
        } catch (e) {}
    }
    //文本框等带文字的形状：递归删除其文字范围里的图片
    try {
        var tr = sh.TextFrame.TextRange;
        if (tr) { cleanInline(tr.InlineShapes); cleanShapes(tr.Shapes); }
    } catch (e) {}
}

// —— 工具：路径比较（大小写、斜杠无关）——
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
