// ribbon.js —— WM_TOOLS 功能区入口

//加载项第一个执行的函数：保存 ribbonUI 与枚举对象
function OnAddinLoad(ribbonUI) {
    if (typeof (window.Application.ribbonUI) != "object") {
        window.Application.ribbonUI = ribbonUI;
    }
    if (typeof (window.Application.Enum) != "object") {
        window.Application.Enum = WPS_Enum;
    }
    return true;
}

//按钮点击分发：打开对应工具的配置对话框
function OnAction(control) {
    if (control.Id === "btnCleanup") {
        try {
            var dpr = window.devicePixelRatio || 1;
            window.Application.ShowDialog(
                GetUrlPath() + "/ui/options.html",
                "清理并另存为",
                440 * dpr, 500 * dpr,
                true // 模态
            );
        } catch (e) {
            alert("打开配置对话框失败：" + (e && e.message ? e.message : e));
        }
    }
    return true;
}

//按钮图标
function GetImage(control) {
    return "images/cleanup.svg";
}

//鼠标悬停说明（简短标题 / 详细说明）
function GetScreentip(control) {
    return "清理并另存为（不影响原文件）";
}

function GetSupertip(control) {
    return "把当前 Word 另存为一份“清理后的版本”：按勾选项接受修订、删除批注、删除各类图片（内嵌/浮动/页眉页脚/文本框内）。原文件与修订状态保持不变。\n点击后弹出选项，可保存为默认配置。";
}
