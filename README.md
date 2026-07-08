# WM_TOOLS —— WPS 文字工具栏

一个 WPS 文字加载项工具栏 **WM_TOOLS**，集合若干实用小工具。每个工具都**便于分享、一看就懂**：鼠标停在按钮上有说明，点击弹出选项面板，常用配置可保存为默认，原文件默认不受影响。

> 本文件随工具增补**持续更新**。开发规范见 [CLAUDE.md](CLAUDE.md)。

## 当前工具

### 清理并另存为… （`btnCleanup`）

把当前 Word 另存为一份「清理后的版本」。点击后弹出选项，可勾选：

- **接受所有修订**（增/删/改落地为正文）
- **删除所有批注**
- **删除图片**，可细分：
  - 内嵌图片
  - 浮动图片
  - 页眉页脚图片
  - 文本框内图片
  （保留文本框、图表、SmartArt、自选图形等非图片对象）

选项会**自动记住**作为下次默认（存在 WPS `PluginStorage`，键 `WMTools.Cleanup.defaults`）。执行时再弹「另存为」选保存位置（默认原目录、文件名加 `_清理版`）。

**原文件与修订状态字节级不变**：通过 `SaveAs2` 把当前文档快照到新路径（原文件从不被写入），所有清理只作用于新文件；完成后重新打开原文件。关键点：先关 `TrackRevisions` 再删除，否则删除会被记成修订而不生效。

## 开发与调试

环境：Node.js + WPS（支持 JSAPI 加载项的版本，如 WPS 2019+/现行 WPS Office）。

```bash
npm install        # 安装 vite / wps-jsapi-declare / wpsjs
npx wpsjs debug    # 启动本地服务并拉起 WPS 加载本工具栏
```

调试时 WPS 功能区出现「WM_TOOLS」标签。改 `js/`、`ui/` 会热更新；改 `manifest.xml`/`package.json` 的 name 需重启 debug。

## 目录结构

```text
manifest.xml        加载项清单（Name: WMTools）
ribbon.xml          功能区：WM_TOOLS 标签 + 各工具按钮（含 getSupertip 悬停说明）
main.js             入口，按序加载 util/ribbon
js/ribbon.js        按钮入口：OnAction 按 control.Id 分发，打开对应工具的配置对话框
js/finalize.js      「清理并另存为」业务逻辑（参数化 SaveFinalVersion(opts)）
js/options.js       「清理并另存为」配置对话框逻辑（默认配置 PluginStorage 持久化）
js/util.js          模板自带工具（WPS_Enum / GetUrlPath）
ui/options.html     「清理并另存为」配置对话框页面
images/cleanup.svg  按钮图标
vite.config.js      构建配置
CLAUDE.md           项目开发规范（新增工具必读）
test/verify_com*.ps1 引擎级验证脚本（COM 复刻核心调用 + 原文件哈希比对）
```

## 新增一个工具（统一模式）

1. `js/<tool>.js`：写业务逻辑，导出全局函数（接收 options 对象）。
2. `ui/<tool>.html` + `js/<tool>.js`：配置对话框，`window.Application` 可直接用；用 `PluginStorage`（键 `WMTools.<tool>.<config>`）存默认配置；执行后 `window.close()`。
3. `images/<tool>.svg`：贴近功能的图标。
4. `ribbon.xml`：在 `WM_TOOLS` 标签下加按钮，配 `onAction`、`getImage`、`getScreentip`、`getSupertip`。
5. `js/ribbon.js`：`OnAction` 增加按 `control.Id` 分发，`ShowDialog` 打开该工具对话框。
6. 更新本 README「当前工具」与 [CLAUDE.md](CLAUDE.md)「当前工具」。

## 验证

### 引擎级（自动化）

```powershell
powershell -ExecutionPolicy Bypass -File test\verify_com.ps1
powershell -ExecutionPolicy Bypass -File test\verify_com2.ps1
```

应输出 `[final] revisions=0 comments=0 ...` 与 `[PASS] original file unchanged`。

### 端到端（手动）

1. WPS 打开一份含修订/批注/图片的 `.docx`（可用 `test/sample_rich.docx`）。
2. 备份原文件并记录 SHA256。
3. 点「WM_TOOLS → 清理并另存为…」，勾选清理项 → 选保存位置。
4. 核对：新版本按勾选项清理；原文件 SHA256 与备份一致、修订/批注/图片原样存在。

## 范围外

- 批量处理、命令行、独立 exe 形态（每个工具只处理当前活动文档）。
- 嵌在图片上的批注锚点等极端边界。
