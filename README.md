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
- **导出为 Markdown 格式（.md）**：与上述清理项并列、默认不勾。不勾则仍另存为 Word；勾选则在清理后把文档转为 Markdown 文件，中间的 `.docx` 自动删除，最终只留一个 `.md`。只勾这一项=原样转 md（图片以 base64 内联）；全勾=先清理图片等再转 md=纯文本 md。**纯 JS 实现**（随插件打包的 `mammoth.browser.js`，即 markitdown 的 docx 引擎），**对方无需安装 Python**。

选项会**自动记住**作为下次默认（存在 WPS `PluginStorage`，键 `WMTools.Cleanup.defaults`）。执行时再弹「另存为」选保存位置（默认原目录、文件名加 `_清理版`）。

**原文件与修订状态字节级不变**：通过 `SaveAs2` 把当前文档快照到新路径（原文件从不被写入），所有清理只作用于新文件；完成后重新打开原文件。关键点：先关 `TrackRevisions` 再删除，否则删除会被记成修订而不生效。

### 英文引号转中文… （`btnQuotes`）

把英文直引号 `"` **成对**换成中文弯引号 `“ ”`（解决 `于"月经后期""闭经"` 这类问题）。可勾选双引号（默认开）/ 单引号（默认关，自动跳过 `don't` 这类英文撇号）。

- **作用范围**：文档里选中了文字 → 只处理选区；未选中 → 处理全文。对话框里会显示本次作用于哪里。
- **判定方式**：逐段按出现顺序成对判定开/闭引号（同时参考已有的弯引号），段与段互不影响；等长单字符就地替换，不破坏格式；只处理正文。
- **可撤销**：全程原生 Range 操作进撤销栈；WPS 支持 `UndoRecord` 时整批合并为**一步 Ctrl+Z**（已实测支持）。

选项自动记住（键 `WMTools.Quotes.defaults`）。逻辑见 `js/quotes.js`，配置面板见 `ui/quotes.html`。

### 上下标批量设置… （`btnSupsub`）

把匹配指定模式的内容批量设为**上标 / 下标 / 取消上下标**，典型场景：论文引用 `[3-4]` → 上标。

- **预设开箱即用**：方括号引用 `[1] [2-5] [1,3]`（默认）、圆括号引用 `(1) （1-3）`，均含括号一起设置（符合 GB/T 7714）。
- **自定义模式（无需懂通配符）**：普通符号原样输入自动转义；用易懂占位符 `{数字}` `{一串数字}` `{字母}` `{一串字母}` `{汉字}` `{任意字符}` `{任意内容}` 表达变化部分；面板实时显示翻译后的 WPS 通配符；「先数一下」按钮只统计命中数不修改，防止误操作。
- **选项**：可勾选「不包含命中文字首尾各 1 个字符」（如不想上标方括号本身）。
- **作用范围 / 可撤销**：同引号工具。主路径用 Find 通配符**仅格式替换**（`Replacement.Text="^&"` + 全部替换），与 WPS 自带替换完全相同、原生一步撤销；「去首尾字符 / 取消格式」走手动查找循环（`UndoRecord` 合并一步撤销）。
- 只处理正文；格式化整个命中片段（「只上标片段中间的字符」如 H₂O 暂不支持，面板内有说明）。

选项自动记住（键 `WMTools.Supsub.defaults`）。逻辑见 `js/supsub.js`，配置面板见 `ui/supsub.html`。

## 安装到 WPS（日常使用 / 分享给别人）

加载项装一次就**常驻**在 WPS 里，以后每次开 WPS 都自动出现「WM_TOOLS」标签，**不需要任何后台服务、不联网**。

### 自己用 / 换电脑 / 重装 WPS 后

双击项目里的 **`安装到WPS.bat`** 即可。它会：

1. 把加载项文件复制到 `%APPDATA%\kingsoft\wps\jsaddons\wm-tools_1.0.0\`；
2. 在 WPS 的 `publish.xml` 里登记一条**离线**记录（`<jsplugin url="wm-tools_1.0.0">`）。

装完**完全关闭所有 WPS 窗口**再重新打开，功能区就有「WM_TOOLS」标签了。

> 命令行等价：`powershell -ExecutionPolicy Bypass -File install.ps1`
> 卸载：`powershell -ExecutionPolicy Bypass -File install.ps1 -Uninstall`

### 分享给别人

把**整个项目文件夹**拷给对方即可——**不需要** `node_modules`，对方也**不用装 Node.js**。对方收到后双击 `安装到WPS.bat`，再重启 WPS 就能用。打包时直接压缩项目根目录（可先删掉 `node_modules/` 减小体积）。

### 为什么之前会“消失”？

用 `npx wpsjs debug` 调试时，加载项是**在线开发模式**：WPS 从一个本地调试服务器（`http://127.0.0.1:3890`）实时拉取代码——**那个终端一关、或电脑一重启，插件就不见了**。`安装到WPS.bat` 走的是**离线模式**，代码直接从本地文件加载，装一次就一直在，不会消失。

## 开发与调试（开发者改代码时用；日常使用请用上面的「安装到 WPS」）

环境：Node.js + WPS（支持 JSAPI 加载项的版本，如 WPS 2019+/现行 WPS Office）。

```bash
npm install        # 安装 vite / wps-jsapi-declare / wpsjs
npx wpsjs debug    # 在线开发模式：起本地服务器 + 拉起 WPS 加载本工具栏
```

`wpsjs debug` 是**在线开发模式**，依赖本地服务器常驻（关掉终端插件就消失，原因见上）。它的好处是：改 `js/`、`ui/` 会**热更新**，无需重装。改 `manifest.xml` / `package.json` 的 `name` 需重启 debug。开发完毕、想稳定常驻，跑一次 `安装到WPS.bat` 即可切到离线模式。

## 目录结构

```text
manifest.xml        加载项清单（Name: WMTools）
ribbon.xml          功能区：WM_TOOLS 标签 + 各工具按钮（含 getSupertip 悬停说明）
main.js             入口，按序加载 util/ribbon
js/ribbon.js        按钮入口：OnAction 按 control.Id 分发，打开对应工具的配置对话框
js/finalize.js      「清理并另存为」业务逻辑（参数化 SaveFinalVersion(opts)）
js/options.js       「清理并另存为」配置对话框逻辑（默认配置 PluginStorage 持久化）
js/textmod.js       文本批量修改公共助手（作用域/UndoRecord/占位符→通配符/手动查找循环）
js/quotes.js        「英文引号转中文」业务逻辑（ConvertStraightQuotes(opts)）
js/quotes_panel.js  「英文引号转中文」配置对话框逻辑
js/supsub.js        「上下标批量设置」业务逻辑（ApplyVertAlign(opts) / CountSupsubMatches(opts)）
js/supsub_panel.js  「上下标批量设置」配置对话框逻辑
js/util.js          模板自带工具（WPS_Enum / GetUrlPath）
js/vendor/          Markdown 导出依赖（mammoth.browser.js + es6-promise，随插件打包，无需对方装 Python）
ui/options.html     「清理并另存为」配置对话框页面
ui/quotes.html      「英文引号转中文」配置对话框页面
ui/supsub.html      「上下标批量设置」配置对话框页面
images/*.svg        按钮图标（cleanup / quotes / supsub）
vite.config.js      构建配置
install.ps1         离线安装脚本（核心逻辑；读取 package.json 的 name/version）
安装到WPS.bat       双击即装（调用 install.ps1）
CLAUDE.md           项目开发规范（新增工具必读）
test/verify_com*.ps1 引擎级验证脚本（COM 复刻核心调用 + 原文件哈希比对）
test/verify_textmod.ps1 引擎级验证脚本（引号成对转换/仅格式替换/撤销）
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
powershell -ExecutionPolicy Bypass -File test\verify_textmod.ps1
```

应输出 `[final] revisions=0 comments=0 ...` 与 `[PASS] original file unchanged`；
`verify_textmod.ps1` 应输出 `summary: PASS=10 FAIL=0`（引号成对转换、仅格式替换、单步撤销）。

### 端到端（手动）

1. WPS 打开一份含修订/批注/图片的 `.docx`（可用 `test/sample_rich.docx`）。
2. 备份原文件并记录 SHA256。
3. 点「WM_TOOLS → 清理并另存为…」，勾选清理项 → 选保存位置。
4. 核对：新版本按勾选项清理；原文件 SHA256 与备份一致、修订/批注/图片原样存在。

## 范围外

- 批量处理、命令行、独立 exe 形态（每个工具只处理当前活动文档）。
- 嵌在图片上的批注锚点等极端边界。
