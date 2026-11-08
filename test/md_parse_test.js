// md_parse_test.js —— js/markdown.js 纯解析函数的 Node 单测（无 WPS 依赖）
// 运行：node test/md_parse_test.js
const assert = require("assert");
const md = require("../js/markdown.js");

let pass = 0, fail = 0;
function t(name, fn) {
    try { fn(); pass++; console.log("PASS " + name); }
    catch (e) { fail++; console.log("FAIL " + name + " :: " + e.message); }
}
const OPTS = { bold: true, italic: true, strike: true, inlineCode: true, link: true };
function kinds(lines) { return md.mdClassifyLines(lines).map(b => b.kind); }

// —— 标题 ——
t("ATX 标题 1~6 级", () => {
    const b = md.mdClassifyLines(["# 阅读资料：叶片里的“阳光工厂”", "## 二级", "###### 六级"]);
    assert.deepStrictEqual(b.map(x => x.kind), ["heading", "heading", "heading"]);
    assert.strictEqual(b[0].level, 1);
    assert.strictEqual(b[0].prefixLen, 2);           // 「# 」
    assert.strictEqual(b[2].level, 6);
});
t("七个 # 不是标题", () => assert.deepStrictEqual(kinds(["####### x"]), ["para"]));
t("# 后无空格不是标题", () => assert.deepStrictEqual(kinds(["#标签"]), ["para"]));

// —— 列表 ——
t("无序列表成组 + 二级缩进", () => {
    const b = md.mdClassifyLines(["- 小麦、水稻", "- 玉米", "  - 子项", "普通段"]);
    assert.strictEqual(b.length, 2);
    assert.strictEqual(b[0].kind, "ul");
    assert.strictEqual(b[0].items.length, 3);
    assert.strictEqual(b[0].items[0].markerLen, 2);  // 「- 」
    assert.strictEqual(b[0].items[2].level, 2);
    assert.strictEqual(b[1].kind, "para");
});
t("有序列表 1. 与 1)", () => {
    const b = md.mdClassifyLines(["1. 光合作用是什么？", "2) 第二步"]);
    assert.strictEqual(b[0].kind, "ol");
    assert.strictEqual(b[0].items.length, 2);
});
t("--- 是分割线不是列表", () => assert.deepStrictEqual(kinds(["---"]), ["hr"]));
t("*** 与 ___ 也是分割线", () => assert.deepStrictEqual(kinds(["***", "___"]), ["hr", "hr"]));

// —— 代码块 ——
t("代码块含语言标记", () => {
    const b = md.mdClassifyLines(["```js", "var a = **1**;", "```", "后文"]);
    assert.deepStrictEqual(b.map(x => x.kind), ["code", "para"]);
    assert.strictEqual(b[0].start, 0);
    assert.strictEqual(b[0].end, 2);
});
t("围栏未闭合保持原文", () => {
    assert.deepStrictEqual(kinds(["```js", "var a = 1;"]), ["para", "para"]);
});

// —— 表格 ——
t("标准管道表格 + 对齐", () => {
    const b = md.mdClassifyLines(["| 名称 | 数量 | 价格 |", "|---|:---:|---:|", "| 苹果 | 3 | 5 |"]);
    assert.strictEqual(b.length, 1);
    assert.strictEqual(b[0].kind, "table");
    assert.deepStrictEqual(b[0].rows, [["名称", "数量", "价格"], ["苹果", "3", "5"]]);
    assert.deepStrictEqual(b[0].aligns, [0, 1, 2]);
});
t("列数不一致 → tableSkip", () => {
    const b = md.mdClassifyLines(["| a | b |", "|---|---|", "| 只有一列 |"]);
    assert.strictEqual(b[0].kind, "tableSkip");
});
t("无分隔行 → tableSkip", () => {
    const b = md.mdClassifyLines(["| a | b |", "| c | d |"]);
    assert.strictEqual(b[0].kind, "tableSkip");
});

// —— 引用 ——
t("连续引用成组", () => {
    const b = md.mdClassifyLines(["> 引用一", "> 引用二", "正文"]);
    assert.deepStrictEqual(b.map(x => x.kind), ["quote", "para"]);
    assert.strictEqual(b[0].end, 1);
});

// —— 行内 ——
t("加粗命中与偏移", () => {
    const ms = md.mdInlineMatches("这是**加粗**文字", OPTS);
    assert.strictEqual(ms.length, 1);
    const m = ms[0];
    assert.strictEqual(m.type, "bold");
    assert.strictEqual(m.start, 2);
    assert.strictEqual(m.end, 8);
    assert.strictEqual("这是**加粗**文字".substring(m.contentStart, m.contentEnd), "加粗");
});
t("斜体与加粗共存不重叠", () => {
    const ms = md.mdInlineMatches("这是*斜体*和**加粗**", OPTS);
    const types = ms.map(m => m.type).sort();
    assert.deepStrictEqual(types, ["bold", "italic"]);
});
t("加粗在前斜体在后（伪匹配回归）", () => {
    // 旧实现会在 **加粗** 和 *斜体* 之间造出伪匹配 *和*，吃掉真正的斜体起点
    const ms = md.mdInlineMatches("这是**加粗**和*斜体*文字", OPTS);
    const types = ms.map(m => m.type).sort();
    assert.deepStrictEqual(types, ["bold", "italic"]);
    const it = ms.filter(m => m.type === "italic")[0];
    assert.strictEqual("这是**加粗**和*斜体*文字".substring(it.contentStart, it.contentEnd), "斜体");
});
t("连续多组粗斜混合", () => {
    const ms = md.mdInlineMatches("**甲**、*乙*、**丙**、*丁*", OPTS);
    assert.deepStrictEqual(ms.map(m => m.type), ["italic", "bold", "italic", "bold"]); // 降序
});
t("行内代码保护内部星号", () => {
    const ms = md.mdInlineMatches("代码 `a*b*c` 和**粗**", OPTS);
    const types = ms.map(m => m.type).sort();
    assert.deepStrictEqual(types, ["bold", "inlineCode"]);
});
t("删除线", () => {
    const ms = md.mdInlineMatches("~~删除~~", OPTS);
    assert.strictEqual(ms[0].type, "strike");
});
t("链接提取 url", () => {
    const ms = md.mdInlineMatches("见[示例](https://a.com)即可", OPTS);
    assert.strictEqual(ms.length, 1);
    assert.strictEqual(ms[0].type, "link");
    assert.strictEqual(ms[0].url, "https://a.com");
    assert.strictEqual("见[示例](https://a.com)即可".substring(ms[0].contentStart, ms[0].contentEnd), "示例");
});
t("图片语法不当链接", () => {
    assert.strictEqual(md.mdInlineMatches("![alt](https://a.com)", OPTS).length, 0);
});
t("多处加粗从右往左应用顺序", () => {
    const ms = md.mdInlineMatches("**甲**和**乙**", OPTS);
    assert.strictEqual(ms.length, 2);
    assert.ok(ms[0].start > ms[1].start);   // 降序
});
t("开关关闭则不命中", () => {
    assert.strictEqual(md.mdInlineMatches("**粗**", { bold: false }).length, 0);
});

console.log("RESULT PASS=" + pass + " FAIL=" + fail);
process.exit(fail ? 1 : 0);
