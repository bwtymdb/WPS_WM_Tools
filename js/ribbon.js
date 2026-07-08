// ribbon.js —— WPS 功能区按钮入口

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

//按钮点击分发
function OnAction(control) {
    if (control.Id === "btnSaveFinal") {
        try {
            SaveFinalVersion();
        } catch (e) {
            alert("生成最终版时发生错误：" + (e && e.message ? e.message : e));
        }
    }
    return true;
}

//按钮图标
function GetImage(control) {
    return "images/final.svg";
}
