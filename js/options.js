// options.js —— 「清理并另存为」配置对话框逻辑
// 默认配置用 PluginStorage 持久化（key: WMTools.Cleanup.defaults），下次自动套用。

var STORAGE_KEY = "WMTools.Cleanup.defaults";

var DEFAULTS = {
    acceptRevisions: true,
    deleteComments: true,
    deletePictures: true,
    pic: { inline: true, floating: true, headerFooter: true, inTextbox: true }
};

function getDefaults() {
    try {
        var s = window.Application.PluginStorage.getItem(STORAGE_KEY);
        if (s) {
            var o = JSON.parse(s);
            o.pic = o.pic || {};
            return {
                acceptRevisions: o.acceptRevisions !== false,
                deleteComments: o.deleteComments !== false,
                deletePictures: o.deletePictures !== false,
                pic: {
                    inline: o.pic.inline !== false,
                    floating: o.pic.floating !== false,
                    headerFooter: o.pic.headerFooter !== false,
                    inTextbox: o.pic.inTextbox !== false
                }
            };
        }
    } catch (e) {}
    return DEFAULTS;
}

function saveDefaults(o) {
    try { window.Application.PluginStorage.setItem(STORAGE_KEY, JSON.stringify(o)); } catch (e) {}
}

function setChecked(id, v) { var el = document.getElementById(id); if (el) el.checked = !!v; }
function isChecked(id) { var el = document.getElementById(id); return !!(el && el.checked); }

function applyDefaultsToUI(o) {
    setChecked("acceptRevisions", o.acceptRevisions);
    setChecked("deleteComments", o.deleteComments);
    setChecked("deletePictures", o.deletePictures);
    setChecked("picInline", o.pic.inline);
    setChecked("picFloating", o.pic.floating);
    setChecked("picHeaderFooter", o.pic.headerFooter);
    setChecked("picInTextbox", o.pic.inTextbox);
    syncPicSubState();
}

//「删除图片」未勾选时，禁用四个子项
function syncPicSubState() {
    var master = isChecked("deletePictures");
    ["picInline", "picFloating", "picHeaderFooter", "picInTextbox"].forEach(function (id) {
        var el = document.getElementById(id);
        if (el) el.disabled = !master;
    });
}

function gatherOpts() {
    return {
        acceptRevisions: isChecked("acceptRevisions"),
        deleteComments: isChecked("deleteComments"),
        deletePictures: isChecked("deletePictures"),
        pic: {
            inline: isChecked("picInline"),
            floating: isChecked("picFloating"),
            headerFooter: isChecked("picHeaderFooter"),
            inTextbox: isChecked("picInTextbox")
        }
    };
}

function onExecute() {
    if (!window.Application.ActiveDocument) {
        alert("请先在 WPS 中打开一个 Word 文档。");
        return;
    }
    var opts = gatherOpts();
    saveDefaults(opts);              // 记住本次选择，作为下次默认
    var ok = SaveFinalVersion(opts); // finalize.js 提供
    if (ok) closeDialog();
}

function closeDialog() { try { window.close(); } catch (e) {} }

window.onload = function () {
    applyDefaultsToUI(getDefaults());
    document.getElementById("btnExecute").onclick = onExecute;
    document.getElementById("btnCancel").onclick = closeDialog;
    document.getElementById("deletePictures").onchange = syncPicSubState;
};
