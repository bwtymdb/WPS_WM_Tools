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

//各工具的对话框配置：control.Id → { url, title, w, h }（宽高为 1x DPR 基准像素）
var TOOL_DIALOGS = {
    btnCleanup: { url: "/ui/options.html", title: "清理并另存为", w: 440, h: 560 },
    btnQuotes:  { url: "/ui/quotes.html",  title: "英文引号转中文", w: 460, h: 420 },
    btnSupsub:  { url: "/ui/supsub.html",  title: "上下标批量设置", w: 640, h: 780 },
    btnMarkdown:{ url: "/ui/markdown.html",title: "Markdown 排版", w: 560, h: 720 }
};

//按钮点击分发：打开对应工具的配置对话框
function OnAction(control) {
    var cfg = TOOL_DIALOGS[control.Id];
    if (!cfg) return true;
    try {
        var dpr = window.devicePixelRatio || 1;
        window.Application.ShowDialog(
            GetUrlPath() + cfg.url,
            cfg.title,
            cfg.w * dpr, cfg.h * dpr,
            true // 模态
        );
    } catch (e) {
        alert("打开配置对话框失败：" + (e && e.message ? e.message : e));
    }
    return true;
}

//按钮图标
function GetImage(control) {
    switch (control.Id) {
        case "btnQuotes": return "images/quotes.svg";
        case "btnSupsub": return "images/supsub.svg";
        case "btnMarkdown": return "images/markdown.svg";
        default: return "images/cleanup.svg";
    }
}

//鼠标悬停说明（简短标题 / 详细说明）
function GetScreentip(control) {
    switch (control.Id) {
        case "btnQuotes": return "直引号成对转中文引号（可撤销）";
        case "btnSupsub": return "引用标注等批量设为上标/下标（可撤销）";
        case "btnMarkdown": return "Markdown 原始文字一键转正常排版（可撤销）";
        default: return "清理并另存为（不影响原文件）";
    }
}

function GetSupertip(control) {
    switch (control.Id) {
        case "btnQuotes":
            return "把选中文字（未选中则全文）里的英文直引号 \" 成对换成中文弯引号 “ ”，可选单引号。\n直接修改当前文档，可用 Ctrl+Z 撤销。\n点击后弹出选项，可保存为默认配置。";
        case "btnSupsub":
            return "把选中文字（未选中则全文）里匹配指定模式的内容（如论文引用 [1]、[2-5]）批量设为上标 / 下标，也可取消。\n内置引用标注预设，自定义模式用易懂占位符（无需懂通配符），可先数命中数再执行。\n直接修改当前文档，可用 Ctrl+Z 撤销。";
        case "btnMarkdown":
            return "AI 生成的 Markdown 文字粘贴到 Word 后全是 #、**、- 等原始符号？一键转成正常排版：标题样式、加粗/斜体、原生列表、真正的表格、代码块、超链接等。\n选中文字则只处理选区，否则处理全文；面板自动检测各类元素数量，可按类勾选。\n直接修改当前文档，Ctrl+Z 一步撤销；配置自动保存。";
        default:
            return "把当前 Word 另存为一份“清理后的版本”：按勾选项接受修订、删除批注、删除各类图片（内嵌/浮动/页眉页脚/文本框内）。原文件与修订状态保持不变。\n还可勾选「导出为 Markdown」：清理后转为 .md，中间 docx 自动删除（纯 JS，无需装 Python）。\n点击后弹出选项，可保存为默认配置。";
    }
}
