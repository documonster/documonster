# Word 模块

[English](README.md)

零依赖的 TypeScript 库，用于读取、写入和操作 DOCX 文件。
支持 Node.js 22+ 和现代浏览器。

## 快速开始

```typescript
import { Document, toBuffer, readDocx } from "@cj-tech-master/excelts/word";

// 创建文档
const doc = Document.create();
Document.addHeading(doc, "Hello World", 1);
Document.addParagraph(doc, "This is a paragraph.");
Document.addTable(doc, [
  ["A", "B"],
  ["1", "2"]
]);

const buffer = await toBuffer(Document.build(doc));
// 写入文件、作为响应返回等

// 读取文档
const parsed = await readDocx(buffer);
console.log(parsed.body.length, "elements");
```

## 核心 API

### 文档构建器

```typescript
import { Document, paragraph, text, bold, italic, heading } from "@cj-tech-master/excelts/word";

const doc = Document.create();

// 文本
Document.addParagraph(doc, "Simple text");
Document.addParagraphElement(doc, paragraph([text("Normal "), bold("bold "), italic("italic")]));
Document.addHeading(doc, "Title", 1);

// 表格
Document.addTable(
  doc,
  [
    ["Header1", "Header2"],
    ["cell1", "cell2"]
  ],
  { headerRow: true }
);

// 图片
Document.addImage(doc, pngBytes, "png", width, height);
Document.addFloatingImage(doc, jpgBytes, "jpeg", width, height, { wrap: { style: "square" } });

// 列表
Document.addBulletList(doc, ["Item 1", "Item 2", "Item 3"]);
Document.addNumberedList(doc, ["First", "Second", "Third"]);

// 页面布局
Document.setSectionProperties(doc, {
  pageSize: { width: 11906, height: 16838 }, // A4
  margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
});

// 页眉/页脚
Document.setHeader(doc, "default", { children: [textParagraph("Page Header")] });

// 构建与导出
const model = Document.build(doc);
const bytes = await toBuffer(model);
```

### 读取文档

```typescript
import { readDocx, extractText, searchText } from "@cj-tech-master/excelts/word";

const doc = await readDocx(fileBuffer);

// 提取文本内容
const text = extractText(doc);

// 搜索
const results = searchText(doc, /pattern/g);
```

### 修改文档

```typescript
import { readDocx, replaceText, toBuffer } from "@cj-tech-master/excelts/word";

const doc = await readDocx(buffer);
const modified = replaceText(doc, "OLD_TEXT", "NEW_TEXT");
const output = await toBuffer(modified);
```

## 高级功能

### 模板引擎

```typescript
import { fillTemplate } from "@cj-tech-master/excelts/word";

const filled = fillTemplate(doc, {
  name: "John",
  showDetails: true,
  items: ["A", "B", "C"]
});
// 支持：{{variable}}、{{#if cond}}...{{/if}}、{{#each arr}}...{{/each}}
```

### 表单字段

```typescript
import {
  extractFormFields,
  fillFormFields,
  formTextField,
  formCheckboxField
} from "@cj-tech-master/excelts/word";

// 提取表单数据
const fields = extractFormFields(doc);
// → [{ name: "FullName", type: "text", value: "..." }, ...]

// 填充表单数据
const filled = fillFormFields(
  doc,
  new Map([
    ["FullName", "Jane Doe"],
    ["AgreeTerms", true],
    ["Country", 2]
  ])
);
```

### 数据绑定（OpenDoPE）

```typescript
import { resolveDataBindings } from "@cj-tech-master/excelts/word";

// 根据 CustomXML 部件解析 SDT 数据绑定
const resolved = resolveDataBindings(doc);

// 或使用覆盖数据
const resolved2 = resolveDataBindings(
  doc,
  new Map([["{GUID}", "<root><field>value</field></root>"]])
);
```

### 带特效的绘图形状

```typescript
import { createShape, createRect, createEllipse } from "@cj-tech-master/excelts/word";

const shape = createShape({
  shapeType: "roundRect",
  width: 3000000, // EMU
  height: 2000000,
  fill: {
    type: "gradient",
    stops: [
      { position: 0, color: "FF0000" },
      { position: 100000, color: "0000FF" }
    ]
  },
  effects: {
    shadow: {
      type: "outer",
      color: "000000",
      blurRadius: 50800,
      distance: 38100,
      direction: 2700000
    },
    glow: { color: "FFFF00", radius: 101600 },
    reflection: { startOpacity: 50, endOpacity: 0, distance: 25400 },
    softEdges: 63500,
    effect3d: {
      camera: "perspectiveFront",
      bevelTop: { width: 127000, height: 63500, preset: "circle" }
    }
  }
});
```

### 字体嵌入与子集化

```typescript
import { embedFont, addEmbeddedFonts, subsetFont } from "@cj-tech-master/excelts/word";

// 嵌入并自动子集化（只嵌入文档中用到的字形）
const result = embedFont({
  name: "CustomFont",
  data: fontFileBytes,
  style: "regular",
  usedCharacters: "Hello World" // 只嵌入这些字形
});

// 添加到文档
const docWithFonts = addEmbeddedFonts(doc, [result]);
```

### 修订追踪

```typescript
import { acceptAllRevisions, rejectAllRevisions } from "@cj-tech-master/excelts/word";

const accepted = acceptAllRevisions(doc);
const rejected = rejectAllRevisions(doc);
```

### 文档比对

```typescript
import { diffDocuments } from "@cj-tech-master/excelts/word";

const diff = diffDocuments(docA, docB);
// → { changes: [{ type: "added"|"removed"|"modified", ... }] }
```

### 文档合并

```typescript
import { mergeDocuments } from "@cj-tech-master/excelts/word";

const merged = mergeDocuments([doc1, doc2, doc3], { sectionBreak: "nextPage" });
```

### 流式写入器

```typescript
import { createDocxStream } from "@cj-tech-master/excelts/word";

const stream = createDocxStream();
stream.addParagraph("Title", { style: "Heading1" });
for (const item of largeDataset) {
  stream.addParagraph(item.text);
}
const buffer = await stream.finalize();
```

### 文档保护

```typescript
import {
  protectDocument,
  isDocumentProtected,
  verifyProtectionPassword
} from "@cj-tech-master/excelts/word";

const protected = protectDocument(doc, { type: "readOnly", password: "secret" });
const isProtected = isDocumentProtected(protected); // true
const valid = verifyProtectionPassword(protected, "secret"); // true
```

### 校验

```typescript
import { validateDocument } from "@cj-tech-master/excelts/word";

const result = validateDocument(doc);
if (!result.valid) {
  console.log(result.issues); // [{ severity, message, path }]
}
```

### HTML/Markdown 转换

```typescript
// DOCX → Markdown（GFM：标题、粗体/斜体/删除线、行内代码、
// 代码块、引用块、有序/无序列表、带对齐的表格、
// 链接、图片、脚注）
import { renderToMarkdown } from "@cj-tech-master/excelts/word/markdown";
const md = renderToMarkdown(doc);
const mdSetext = renderToMarkdown(doc, { headingStyle: "setext" });

// Markdown → DOCX（完整文档或正文片段）
import { markdownToDocx, markdownToDocxBody } from "@cj-tech-master/excelts/word/markdown";
const doc = markdownToDocx("# Title\n\nHello **world**");
const bodyItems = markdownToDocxBody("- a\n- b");

// DOCX → HTML
import { renderToHtml } from "@cj-tech-master/excelts/word/html";
const html = renderToHtml(doc);

// HTML → DOCX 正文内容
import { htmlToDocxBody } from "@cj-tech-master/excelts/word/html";
const body = htmlToDocxBody("<h1>Hello</h1><p>World</p>");
```

### Flat OPC 格式

```typescript
import { parseFlatOpc, toFlatOpc, isFlatOpc } from "@cj-tech-master/excelts/word";

// DOCX 的单一 XML 表示形式
const flatXml = toFlatOpc(doc);
const doc = parseFlatOpc(flatXmlString);
```

### Excel → Word

将 Excel `Workbook` 转换为 `DocxDocument`，把工作表映射为
带单元格格式（字体、颜色、对齐、填充、边框）、列宽、富文本 run
以及可选标题页的 Word 表格。隐藏的工作表会被跳过；行/列数量可以设上限。

```typescript
import { excelToDocx, extractTablesToExcel } from "@cj-tech-master/excelts/word/excel";
import { packageDocx } from "@cj-tech-master/excelts/word";

// Workbook → DocxDocument（所有可见工作表，保留格式）
const doc = excelToDocx(workbook);
const docxBytes = await packageDocx(doc);

// 带选项：标题页、只取部分工作表、限制行列数
const doc2 = excelToDocx(workbook, {
  titlePage: { title: "Q3 Report", subtitle: "Sales" },
  sheets: ["Summary", 2], // 按名称或从 0 开始的索引
  maxRows: 100,
  maxColumns: 12,
  preserveFormatting: true
});

// 反向：把 DOCX 中的表格抽取为 Workbook
const extracted = extractTablesToExcel(doc);
```

嵌入在 Word 文档中的图表也会桥接到 Excel 图表引擎
（27 个图表系列，经典图表与现代 ChartEx）。PDF 渲染侧参见
`createWordChartPdfRenderer` / `createWordLayoutChartPdfRenderer`。

### Word → PDF

将 `DocxDocument` 转换为 PDF 字节。该桥接是共享 Word 布局引擎之上的薄层，
因此换行、分页、表格、行内图片、页眉/页脚和浮动元素的渲染结果
与 SVG 路径完全一致。

```typescript
import { readDocx } from "@cj-tech-master/excelts/word";
import { docxToPdf } from "@cj-tech-master/excelts/pdf";

const doc = await readDocx(docxBytes);
const pdfBytes = await docxToPdf(doc);

// 覆盖页面几何（单位：磅）。省略的字段会回退到
// 文档的 section properties，再回退到引擎默认值（US Letter，1 英寸）。
const pdf2 = await docxToPdf(doc, {
  pageWidth: 595, // A4 宽度
  pageHeight: 842, // A4 高度
  marginTop: 72,
  marginBottom: 72,
  marginLeft: 72,
  marginRight: 72,
  headerMargin: 36, // 页眉带相对顶边的偏移
  footerMargin: 36, // 页脚带相对底边的偏移
  defaultFont: "Helvetica",
  defaultFontSize: 11
});
```

**图表。** 调用过 `installChartSupport()` 后，经典图表
（`<c:chart>`）和现代 ChartEx（`<cx:chartSpace>` —— 旭日图、矩形树图、
瀑布图、漏斗图、箱线图、直方图、帕累托图、地图）都会自动渲染为
完整矢量 PDF。若未安装图表支持，图表会降级为带标题的占位框
（不抛错、不留空白页）。要提供你自己的经典图表渲染器：

```typescript
import { installChartSupport } from "@cj-tech-master/excelts/chart";
import { docxToPdf, createWordChartPdfRenderer } from "@cj-tech-master/excelts/pdf";

installChartSupport();
const pdf = await docxToPdf(doc, {
  chartRenderer: createWordChartPdfRenderer()
});
```

`chartRenderer` 可以返回 `false` 来拒绝渲染某个图表；此时桥接会
回退到内置矢量渲染器，再回退到行内 SVG / 占位框。

## OOXML Strict 格式

本模块会自动处理以 ISO 29500 Strict 一致性级别保存的文档。
读取 Strict 格式的 .docx 时，命名空间 URI 和关系类型会被透明地
归一化为对应的 Transitional 形式 —— 无需任何用户操作。

## 兼容性

| 特性             | 支持情况                                             |
| ---------------- | ---------------------------------------------------- |
| .docx（读取）    | ✅ 广泛（常见元素结构化；未知部件保留）              |
| .docx（写入）    | ✅ 广泛（常见元素；未知部件以不透明形式写出）        |
| .dotx（模板）    | ✅                                                   |
| .docm（宏）      | ✅ 往返（保留 VBA，不执行/不编辑）                   |
| .dotm（宏模板）  | ✅ 往返                                              |
| Flat OPC（.xml） | ✅                                                   |
| ISO 29500 Strict | ✅ 自动归一化                                        |
| 加密的 .docx     | ✅ 用密码解密（Agile Encryption）                    |
| 数字签名         | 🔍 检测与元数据提取（不签名/不验证）                 |
| 浏览器           | ✅（从 "@cj-tech-master/excelts/word/browser" 导入） |
| Node.js 22+      | ✅                                                   |

## 从 `docx`（npm）迁移

| docx（npm）                         | excelts/word                                             |
| ----------------------------------- | -------------------------------------------------------- |
| `new Document()`                    | `Document.create()`                                      |
| `new Paragraph({ text })`           | `textParagraph(text)`                                    |
| `new TextRun({ bold: true, text })` | `bold(text)`                                             |
| `new Table({ rows })`               | `table(rows)`                                            |
| `Packer.toBuffer(doc)`              | `toBuffer(doc)`                                          |
| ❌ 无读取功能                       | `readDocx(buffer)`                                       |
| ❌ 无修改功能                       | `replaceText(doc, old, new)`                             |
| ❌ 无模板功能                       | `fillTemplate(doc, data)`                                |
| ❌ 无表单功能                       | `extractFormFields(doc)` / `fillFormFields(doc, values)` |

## 从 `mammoth.js` 迁移

| mammoth.js                         | excelts/word                                         |
| ---------------------------------- | ---------------------------------------------------- |
| `mammoth.convertToHtml(input)`     | `readDocx(buf)` → `renderToHtml(doc)`                |
| `mammoth.convertToMarkdown(input)` | `readDocx(buf)` → `renderToMarkdown(doc)`            |
| 样式映射                           | `parseStyleMap(rules)` / `matchStyleMap(style, map)` |
| ❌ 无写入功能                      | 完整读/写/修改                                       |
