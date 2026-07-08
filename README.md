# Word 另存最终版（WPS 加载项）

一个 WPS 文字加载项：把当前 Word 文档一键**另存为一份「最终版」**——
在新版本中接受所有修订、删除全部批注、删除全部图片，且**原文件与修订状态原封不动**。

## 功能

点击功能区「**另存最终版**」标签下的按钮后，对当前活动文档：

1. 弹出「另存为」对话框（默认名 `原名_最终版.docx`，默认目录为原文件所在目录，可改）。
2. 把当前文档**快照**到所选新路径（原文件从不被写入）。
3. 在新版本中：
   - 关闭修订记录 → 接受**全部修订**（增/删/改落地为正文）；
   - 删除**全部批注**；
   - 删除**全部图片**（正文内嵌图片 + 浮动图片 + 页眉/页脚/文本框内的图片；保留文本框、图表、SmartArt、自选图形）。
4. 保存新版本并置于前台；重新打开原文件，保持其修订状态不变。

> 硬约束：原文件字节级不变。已用 WPS COM 自动化对 `SaveAs2` 做过哈希校验通过（见 `test/verify_com.ps1`）。

## 开发与调试

环境：Node.js + WPS（支持 JSAPI 加载项的版本，如 WPS 2019+/现行 WPS Office）。

```bash
npm install        # 安装 vite / wps-jsapi-declare / wpsjs
npx wpsjs debug    # 启动本地服务并拉起 WPS 加载该加载项
```

调试时 WPS 会打开，功能区出现「另存最终版」标签。修改 `js/` 下代码后，WPS 会热更新。

## 目录结构

```
manifest.xml        加载项清单（名称 SaveFinalDoc）
ribbon.xml          功能区定义（一个按钮 btnSaveFinal）
main.js             入口，按序加载 util/ribbon/finalize
js/ribbon.js        按钮入口：OnAction → SaveFinalVersion()
js/finalize.js      核心逻辑：快照另存 + 接受修订 + 删批注 + 递归删图片 + 重开原文件
js/util.js          模板自带工具（WPS_Enum 等）
images/final.svg    按钮图标
vite.config.js      构建配置
test/verify_com.ps1 引擎级验证脚本（COM 复刻核心调用 + 原文件哈希比对）
```

## 核心机制：为何不影响原文件

通过 `doc.SaveAs2(outPath)` 把**当前文档对象**另存到新路径——这是对内存中文档的另存操作，
**不会写回原文件**。随后所有处理（接受修订/删批注/删图片）都作用于这个已另存的新文档对象。
完成后 `Documents.Open(originalPath)` 重新打开原文件，确保用户原文件仍以原始修订状态存在。

关键细节：**必须先 `TrackRevisions = false` 再删除内容**，否则在修订模式下删除会被记录为
"修订删除"而非真实删除（图片不会真正消失，还会新增修订）。此顺序已由 `test/verify_com.ps1` 验证。

## 验证

### 引擎级（自动化）
```powershell
powershell -ExecutionPolicy Bypass -File test\verify_com.ps1
```
应输出：`[final] revisions=0 comments=0 inlineShapes=0` 与 `[PASS] original file unchanged`。

### 端到端（手动）
1. 用一份含修订、批注、图片的 `.docx`（可先用 `test/sample_orig.docx`）在 WPS 中打开。
2. 复制一份原文件备份，记录其 SHA256。
3. 点击「另存最终版」，选择输出路径。
4. 核对：
   - 新版本：修订全部接受、批注为 0、图片消失、文本框/图表保留、修订记录已关闭。
   - 原文件：SHA256 与备份一致；重开后修订/批注/图片原样存在。

## 范围外

- 批量处理、命令行、独立 exe 形态（本加载项只处理当前活动文档）。
- 极端边界（如批注锚定在图片上的特殊情况）。
