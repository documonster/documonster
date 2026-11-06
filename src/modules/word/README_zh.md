# Word 模块

[English](README.md)

零依赖的 TypeScript 库，用于读取、写入和操作 DOCX 文件。
支持 Node.js 22+ 和现代浏览器。

## 功能

- **创建、读取和修改 DOCX 文件** — 完整的 WordprocessingML 支持
- **文档构建器** — 段落、标题、富文本 run（粗体/斜体/下划线/颜色/高亮）
- **表格** — 样式、边框、单元格合并（水平 + 垂直）、嵌套表格
- **图片** — 内联和浮动（JPEG、PNG、GIF、BMP、TIFF，SVG 带光栅 fallback）
- **列表** — 无序、有序、多级
- **页面布局** — 分节、页面尺寸/方向/边距、分栏、分隔符
- **页眉和页脚** — 默认、首页、奇偶页
- **超链接和书签** — 外部链接、内部链接、交叉引用
- **目录** — 基于域，带缓存条目
- **脚注、尾注和批注** — 包括线程/扩展批注
- **域（field）** — PAGE/NUMPAGES、TOC、INDEX/XE、REF、SEQ、STYLEREF、公式等
- **数学公式** — OMML 公式（分数、根式、矩阵、n 元运算符……）
- **绘图形状** — 预设几何，支持填充、线条、渐变、阴影/发光/3D 特效
- **图表** — 不透明保留 + 从零构建；桥接到 Excel 图表引擎
- **修订追踪** — 接受/拒绝插入、删除、移动和属性变更
- **模板** — `{{variable}}`、`{{#if}}`、`{{#each}}`、邮件合并、模板修补
- **表单字段和内容控件（SDT）** — 提取和填充；OpenDoPE 数据绑定
- **字体嵌入** — 自动字形子集化
- **文档保护** — 只读/批注/表单，带密码
- **加密** — 读写受密码保护的 DOCX（Agile Encryption，与 Word 写出的一致）
- **数字签名** — 检测和元数据提取
- **转换** — DOCX ↔ HTML、DOCX ↔ Markdown、Excel → Word、Word → PDF
- **Flat OPC** — 单一 XML `.xml` 表示形式往返
- **比对和合并** — 比较两个文档、合并多个文档
- **流式写入器** — `new Stream.DocumentWriter()` 处理大文档
- **校验** — 结构检查，带严重级别标记
- **OOXML Strict** — 读取时透明归一化为 Transitional
- **浏览器支持** — 同一个 `documonster/word` 导入在 Node.js 和浏览器中均可用

## 快速开始

```typescript
import { Document } from "documonster/word";

// 创建文档
const doc = Document.create();
Document.addHeading(doc, "Hello World", 1);
Document.addParagraph(doc, "This is a paragraph.");
Document.addTable(doc, [
  ["A", "B"],
  ["1", "2"]
]);

const buffer = await Document.toBuffer(doc);
// 写入文件、作为响应返回等

// 读取文档
const parsed = await Document.read(buffer);
console.log(parsed.body.length, "elements");
```

## 核心 API

### 文档构建器

```typescript
import { Document, Build, PageSizes, Units } from "documonster/word";

const doc = Document.create();

// 文本
Document.addParagraph(doc, "Simple text");
Document.addParagraph(
  doc,
  Build.paragraph([Build.text("Normal "), Build.bold("bold "), Build.italic("italic")])
);
Document.addHeading(doc, "Title", 1);

// 表格
Document.addTable(
  doc,
  [
    ["Header1", "Header2"],
    ["cell1", "cell2"]
  ],
  // headerRow 让首行在每页重复；headerRunProperties 决定它的样式
  { headerRow: true, headerRunProperties: { bold: true } }
);

// 图片
// width / height 单位是 EMU，用 Units 换算
Document.addImage(doc, pngBytes, {
  type: "png",
  width: Units.cmToEmu(5),
  height: Units.cmToEmu(3)
});
Document.addFloatingImage(doc, jpgBytes, {
  type: "jpeg",
  width: Units.cmToEmu(4),
  height: Units.cmToEmu(4),
  wrap: { style: "square" }
});

// 列表
Document.addBulletList(doc, ["Item 1", "Item 2", "Item 3"]);
Document.addNumberedList(doc, ["First", "Second", "Third"]);

// 页面布局
Document.setSectionProperties(doc, {
  pageSize: PageSizes.A4,
  margins: {
    top: Units.cmToTwips(2.5),
    bottom: Units.cmToTwips(2.5),
    left: Units.cmToTwips(2),
    right: Units.cmToTwips(2)
  }
});

// 页眉/页脚
Document.setHeader(doc, "default", { children: [Build.textParagraph("Page Header")] });

// 构建与导出
const bytes = await Document.toBuffer(doc);
```

### 读取文档

```typescript
import { Document, Query } from "documonster/word";

const doc = await Document.read(fileBuffer);

// 提取文本内容
const text = Query.getText(doc);

// 搜索
const results = Query.searchText(doc, /pattern/g);
```

在 Node 中，`Document.readFile(path)` 与 `Document.writeFile(doc, path)` 直接读写文件；浏览器构建不提供这两个成员——请自行读取字节后交给
`Document.read`。系统不允许读取或写入的文件报 `DocxFileError`（含 `path`、`operation`），以系统错误为 `cause`（`error.cause.code`
为 `ENOENT`、`EACCES` 等），与 Excel 的 `ExcelFileError` 一致；读到了但不是有效文档的文件报 `DocxParseError`。

### 唯一的文档对象

`Document.create()` 与 `Document.read()` 返回同一种对象——`DocxDocument`，即文档的内容树
——所有命名空间都接收它。因此
`Document.*` 辅助函数既能用于新建的文档，也能用于读取的文档。其字段是公开的，也可以与
辅助函数交替直接编辑这棵树。只要两次辅助函数调用之间的每次编辑属于以下之一，新增的脚注、
尾注、批注、图片和列表就会获得与文档当时已有 id 不冲突的编号：

- 为某个集合赋予一个新数组——在文档本身上，或在副本上（如 `{ ...doc, footnotes: [...] }`，
  这正是 `readonly` 模型类型表达编辑的方式）；
- 在集合末尾追加条目（`push`）。

这两种编辑都会被精确识别，且每次辅助函数调用的开销为 O(1)，因此循环添加 100,000 项仍是
线性的。要替换、删除或重排条目，请赋予新数组：原地修改条目（`doc.footnotes[0] = …`、
`splice`、`sort`、修改条目对象）不会被下一次辅助函数调用看到。但这种编辑绝不会被静默写入
文件——若它使两个脚注、尾注、批注或编号定义拥有相同 id，或使两张不同的图片共用同一个键，
`Document.toBuffer` 会以 `DocxDuplicateIdError` 拒绝，并指明集合与 id：正文按 id 引用这些条目，
写出器无法判断某个引用指的是哪一条。书签与绘图 id 发生冲突时由写出器重新编号，因此对它们的
任何编辑都是安全的。

`Document.*` 与 `Build.*` 会在修改任何内容之前检查传入的参数。文件中无法表示的值——`NaN` 或
非整数的字号、1–9 以外的标题级别、写成 `"#FF0000"` 而非 `"FF0000"` 的颜色、Word 未定义的下划线
样式、本应是文本处的 `null`、超出末尾的 `insertContent` 下标——会抛出 `DocxArgumentError`，
其 `field` 指明出错的参数（`properties.run.size`），消息指明调用：
`Document.addHeading: level must be an integer 1–9, got 12`。绘图的尺寸或位置以其子类
`DocxInvalidDrawingError` 报告。手工构建的内容直到写出时才会检查。

图片的键——`ImageDef.rId` 或其 `aliasRIds` 之一——在每个部件中都指向这张图片：正文、页眉、页脚、
脚注、批注和图片项目符号都按键引用 `doc.images`。而包中的关系 id 只在各自部件内有效（两个页眉
可以各自把自己的标志称为 `rId1`），因此 `Document.read` 会为本地 id 已被另一张图片占用的图片分配新键，
写出器也会为每个显示该图片的部件单独登记关系。`Transform.merge` 会重命名追加文档中与基础文档
冲突的键，无论它们在何处被引用。

`Document.create()` 创建的文档带有 `deriveEastAsianDefaults: true`：写出、排版或转换时，
若未显式指定，会根据文本推导东亚校对语言与字体。`Document.read()` 读取的文档不带该字段，因此
往返读写不会凭空获得作者未选择的语言。它是普通字段——副本会保留它，也可以在任何文档上设置或清除。

同一字段还会在未指定处补上当前 Word 新建文档的默认值——Aptos 12 磅、`en-US`、段后 8 磅、
1.16 倍行距（`Document.setDocDefaults` 与 `Document.useDefaultStyles` 设定的值保持不变）。
文档按这些值排版并写出，因此无论 Word 版本或 Normal 模板如何，打开后的分页都与本库一致。
读取的文件若没有 `w:docDefaults` 则保持原样；本库按 Calibri 11 磅单倍行距排版，这只是猜测——
Word 会从用户的 Normal 模板取这类文件的默认值。

```typescript
import { Document } from "documonster/word";

const doc = await Document.read(fileBuffer);
Document.addParagraph(doc, "追加的段落");
Document.addFootnote(doc, "新的脚注");
const bytes = await Document.toBuffer(doc);
```

#### 谁修改什么

一条规则决定函数是否修改传入的文档：

- **`Document.*` 原地修改传入的文档**——包括各构建函数，以及整文档编辑
  `Document.replaceText`、`Document.mailMerge`、`Document.acceptAllRevisions` /
  `rejectAllRevisions` / `acceptRevision` / `rejectRevision` 与
  `Document.setCompatibilityMode`。
- **其他命名空间中返回 `DocxDocument` 的函数都不修改输入**——`Template.fill` / `fillFromSource`、
  `Transform.*`、`Security.*`、`Ole.add`、`Vba.add` / `remove`、`Font.addEmbedded`。
  结果可能与输入共享未改动的节点（无需改动时甚至就是输入本身）。`Document.*` 在修改前会先复制两份文档
  共享的部分，因此原地编辑任何一方都不会影响另一方——无需克隆。需要快照、或要直接手改节点时再用
  `Document.clone(...)`。

`Query.*` 只读。

```typescript
import { Document, Template } from "documonster/word";

const template = await Document.read(templateBytes);
for (const row of rows) {
  // `fill` 不修改 `template`，编辑结果也不会影响它。
  const letter = Template.fill(template, row);
  Document.replaceText(letter, "DRAFT", "FINAL");
  await Document.writeFile(letter, `out/${row.id}.docx`);
}
```

### 修改文档

```typescript
import { Document } from "documonster/word";

const doc = await Document.read(buffer);
const count = Document.replaceText(doc, "OLD_TEXT", "NEW_TEXT"); // 原地修改 `doc`，返回替换次数
const output = await Document.toBuffer(doc);
```

## 高级功能

### 模板引擎

```typescript
import { Template } from "documonster/word";

const filled = Template.fill(doc, {
  name: "John",
  showDetails: true,
  items: ["A", "B", "C"]
});
// 支持：{{variable}}、{{#if cond}}...{{/if}}、{{#each arr}}...{{/each}}
```

`Template.fill`（以及在 `.docx` 字节上使用同一引擎的 `Package.fillTemplate`）根据一个数据对象
解析变量、条件与循环，包括批注正文。`Package.patch` / `Package.patchTemplate` 是另一个引擎：
把每个命名占位符精确替换为文字、段落、表格或图片——没有数据对象、没有条件与循环，也不处理批注。
模板含逻辑时用 `fill`，只需把占位符换成现成内容时用 `patch`。

### 字段

```typescript
import { Document, Query, Transform } from "documonster/word";

// 每个字段的解读结果——整体字段，或跨 run、跨段落分片保存的字段
const fields = Query.listFields(doc);
// → [{ type: "PAGE", instruction: " PAGE ", result: "3", split: false, part: "footer", … }, …]

// 用数据替换合并域；重新计算页码、SEQ、REF、公式与目录——
// 替换字段 separate 与 end 之间的结果，字段本身保留
Document.mailMerge(doc, { Name: "Alice" }); // 原地修改
const updated = Transform.updateFields(doc);
```

两者都像 Word 一样应用字段开关：`\* Upper` / `Lower` / `Caps` / `FirstCap`
、数字格式（`roman`、`ALPHABETIC`、`Ordinal` 等）与英文单词（`CardText`、`OrdText`、`DollarText`）、`\#` 数字图片
（`\# "#,##0.00"`）和 `\@` 日期图片（`\@ "d MMMM yyyy"`）；`MERGEFIELD` 的
`\b` / `\f` 文字仅在值非空时添加。`updateFields` 也会计算 `QUOTE`（其文字，再应用开关）。合并后，测试合并域的 `IF` 会用合并后的值重新求值。
`listFields` 中，嵌套在另一字段指令里的字段按 Word“切换域代码”视图显示
（`IF { MERGEFIELD Vip } = "Y" …`）；`split` 表示模型如何保存该字段
（按 `w:fldChar` 分片，或作为一个 `FieldContent`）。

此处计算的页码（`PAGE`、`NUMPAGES`、`SECTIONPAGES`、`PAGEREF`、`INDEX`、目录条目）
来自本库自己的分页布局——即 `Layout.compute` 与 PDF 导出生成的页面——按各字体自身的度量、
Word 的段落间距、孤行控制与段落保持规则计算。目录先以新条目参与分页再编号，与 Word 相同；
`PAGEREF` 与目录条目显示的是该页所在节给出的页码（如 `ii`，或重新编号后的 `1`）。
不会设置 `w:updateFields`：Word 打开文档时不会询问是否更新域。调用方仍可自行设置
`settings.updateFieldsOnOpen`。

### 表单字段

```typescript
import { Query, Transform } from "documonster/word";

// 提取表单数据
const fields = Query.listFormFields(doc);
// → [{ name: "FullName", type: "text", value: "..." }, ...]

// 填充表单数据（返回新文档，`doc` 不变）。文本域取字符串，复选框取布尔值，
// 下拉框取选项序号或选项文字；文档中没有的域名会被拒绝。
const filled = Transform.fillFormFields(
  doc,
  new Map<string, string | boolean | number>([
    ["FullName", "Jane Doe"],
    ["AgreeTerms", true],
    ["Country", 2]
  ])
);
```

### 数据绑定（OpenDoPE）

```typescript
import { Transform } from "documonster/word";

// 根据 CustomXML 部件解析 SDT 数据绑定（返回新文档）
const resolved = Transform.resolveDataBindings(doc);

// 或使用覆盖数据
const resolved2 = Transform.resolveDataBindings(
  doc,
  new Map([["{GUID}", "<root><field>value</field></root>"]])
);
```

### 带特效的绘图形状

```typescript
import { Build } from "documonster/word";

const shape = Build.shape({
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
import { Font } from "documonster/word";

// 嵌入并自动子集化（只嵌入文档中用到的字形）
const result = Font.embed({
  name: "CustomFont",
  data: fontFileBytes,
  style: "regular",
  usedCharacters: "Hello World" // 只嵌入这些字形
});

// 添加到文档
const docWithFonts = Font.addEmbedded(doc, [result]);
```

### 修订追踪

```typescript
import { Document } from "documonster/word";

// 两者都原地修改文档并返回处理的修订数；要保留原文档请先克隆。
const accepted = Document.clone(doc);
Document.acceptAllRevisions(accepted);
const rejected = Document.clone(doc);
Document.rejectAllRevisions(rejected);
```

### 文档比对

```typescript
import { Diff } from "documonster/word";

const diff = Diff.compare(docA, docB);
// → { changes: [{ type: "added"|"removed"|"modified", ... }] }
```

### 文档合并

```typescript
import { Transform } from "documonster/word";

const merged = Transform.merge([doc1, doc2, doc3], { sectionBreak: "nextPage" });
```

### 流式写入器

```typescript
import { Stream } from "documonster/word";

const stream = new Stream.DocumentWriter();
stream.addText("Title", { style: "Heading1" });
for (const item of largeDataset) {
  stream.addText(item.text);
}
const buffer = await stream.toBuffer();
```

传入 `new Stream.DocumentWriter({ sink })` 时，字节会边生成边写入 sink：用 `await stream.commit()` 收尾，它会结束该 sink（此时 `toBuffer()` 会被拒绝，因为写入器不持有字节）。

### 文档保护

```typescript
import { Security } from "documonster/word";

// 返回新文档，`doc` 本身不变。
const protectedDoc = await Security.protect(doc, { edit: "readOnly", password: "secret" });
const isProtected = Security.isProtected(protectedDoc); // true
const valid = await Security.verifyPassword(protectedDoc, "secret"); // true
```

哈希算法在这里和 `WordCrypto.encrypt` 中都写作 `"SHA-256"`（与 `Security.getState` 报告的一致）；两者也接受加密格式自己的写法 `"SHA256"`。

文档保护只限制在 Word 中编辑，并不加密文件。要写出带密码的文件，请给
`Document.toBuffer` 传入 `password`（见[加密与签名](#加密与签名)）。

### 校验

```typescript
import { Validation } from "documonster/word";

const result = Validation.validate(doc);
if (!result.valid) {
  console.log(result.issues); // [{ severity, message, path }]
}
```

读入的文档可能给某个字段写了它不接受的取值，例如 `<w:jc w:val="bogus"/>`。这时字段读作没有，原文保留下来，只要字段仍是空的，
写出时就原样写回；`Validation.validate` 会以警告报告它，规则为 `"unrecognized-value"`，并给出字段路径（如 `body[3].properties.alignment`）。
设置字段后以设置的值为准；但把字段再清空，保留的原文又会生效，所以要去掉这些值，请调用 `Document.discardUnrecognizedValues(doc)`；传 `{ filter }` 可以按路径和原文只去掉其中一部分。
`Document.listUnrecognizedValues(doc)` 列出当前仍保留的值，每项是 `DocxUnrecognizedValue`（`{ path, value }`），与 `filter` 收到的对象相同。

### HTML/Markdown 转换

```typescript
// DOCX → Markdown（GFM：标题、粗体/斜体/删除线、行内代码、
// 代码块、引用块、有序/无序列表、带对齐的表格、
// 链接、图片、脚注）
import { WordMarkdown } from "documonster/word/markdown";
const md = WordMarkdown.fromDocx(doc);
const mdSetext = WordMarkdown.fromDocx(doc, { headingStyle: "setext" });

// Markdown → DOCX（完整文档或正文片段）
const fromMarkdown = await WordMarkdown.toDocx("# Title\n\nHello **world**");
const { body: bodyItems } = await WordMarkdown.toDocxBody("- a\n- b");

// 表格列宽按内容测量，代码块字号收缩到最长行不折行为止 —— 这两件事只能由生产方决定，
// 因为 .docx 里存的是列网格和字号，而不是"稍后再算"的指令。正文若不是投放到 Letter
// 页面，就把目标版心宽度说明清楚：
const wide = await WordMarkdown.toDocx(markdown, {
  contentWidth: 15840 - 2 * 720, // 横向 Letter，半英寸页边距
  codeBlockFit: "shrink" // 默认值；"wrap" 保持字号、改为折行
});

// DOCX → HTML
import { WordHtml } from "documonster/word/html";
const { html } = WordHtml.fromDocx(doc);

// HTML → 正文内容，以及它引用的样式、列表编号和图片
const { body, styles, numberingInstances, images } = WordHtml.toDocxBody(
  "<h1>Hello</h1><p>World</p>"
);

// HTML → 完整文档（已包含标题样式与列表编号定义）
const htmlDoc = WordHtml.toDocx("<h1>Hello</h1><ul><li>World</li></ul>");
```

### Flat OPC 格式

```typescript
import { Convert, Document } from "documonster/word";

// DOCX 的单一 XML 表示形式
const flatXml = await Document.toFlatOpc(doc);
const again = await Document.read(flatXml); // Document.read 也能读取 Flat OPC 文本（或字节）
const parts = Convert.flatOpcToParts(flatXml); // Map<partPath, bytes>

// Document.toBase64 写出的 base64
const fromBase64 = await Document.read(await Document.toBase64(doc), { base64: true });
```

### 加密与签名

受密码保护的 DOCX——OLE 复合文件中的 Agile 加密，即 Word「用密码进行加密」写出的文件——用 `{ password }` 读取
（`Document.read`、`Document.readFile`、`Package` 系列函数、`new Stream.DocumentReader`），也用它写出（`Document.toBuffer`、`Document.toBase64`、
`Document.writeFile`）；`encryption` 用于调整参数（Office 默认值：AES-256、SHA-512、100,000 次迭代）。Excel 的
`Workbook.read` / `toBuffer` 接受相同的选项，两个模块共用一份实现。只有提供密码时才加载加密代码。未提供密码时，
加密文件报告为 `DocxEncryptedError`；密码错误时为 `DocxDecryptionError`，其 `reason` 为 `"password"`。只实现了
Agile 加密：Word 2007 的标准加密会以 `reason: "unsupported"` 拒绝。

```typescript
import { Document } from "documonster/word";

const encrypted = await Document.toBuffer(doc, { password: "password" });
const reread = await Document.read(encrypted, { password: "password" });
```

更底层的工具位于 `documonster/word/crypto` 子路径，因此不会被打进只读写普通 DOCX 的 bundle。

```typescript
import { Document } from "documonster/word";
import { WordCrypto } from "documonster/word/crypto";

// 把受密码保护的 DOCX 解密为数据包（Agile Encryption）
if (WordCrypto.isEncrypted(bytes)) {
  const plain = await WordCrypto.decrypt(bytes, "password");
}

// 加密已有的数据包字节
const encrypted = await WordCrypto.encrypt(await Document.toBuffer(doc), "password");

// 不解密数据包，仅校验密码
const ok = await WordCrypto.verifyPassword(bytes, "password");

// 检查嵌入的 XMLDSig 签名（只读；不做验证）
const signatures = WordCrypto.listSignatures(doc);
```

`decrypt` 背后的底层构件同样公开，便于自定义流程，位于两个底层子命名空间：
`WordCrypto.Cfb`（`read` / `write`，OLE2 复合文件）与 `WordCrypto.Agile`
（`parseEncryptionInfo`、配合 `BLOCK_KEYS` 使用的 `deriveKey`、针对已解析信息的
`verifyPassword` 以及 `decryptPackage`，后者和 `decrypt` 一样用 `{ maxDecryptedSize }`
限制大小）。ODTTF 字体混淆原语属于字体工具，位于
`documonster/word` 的 `Font`（`Font.obfuscate` / `Font.deobfuscate` /
`Font.generateObfuscationKey`）。

```typescript
import { WordCrypto } from "documonster/word/crypto";

const entries = WordCrypto.Cfb.read(bytes);
const infoXml = new TextDecoder().decode(
  entries.find(e => e.name === "EncryptionInfo")!.data.slice(8)
);
const info = WordCrypto.Agile.parseEncryptionInfo(infoXml);
const zip = await WordCrypto.Agile.decryptPackage(
  entries.find(e => e.name === "EncryptedPackage")!.data,
  info,
  "password"
);
```

### Excel → Word

将 Excel `Workbook` 转换为 `DocxDocument`，把工作表映射为
带单元格格式（字体、颜色、对齐、填充、边框）、列宽、富文本 run
以及可选标题页的 Word 表格。隐藏的工作表会被跳过；行/列数量可以设上限。

```typescript
import { WordExcel } from "documonster/word/excel";
import { Document } from "documonster/word";

// Workbook → DocxDocument（所有可见工作表，保留格式）
const doc = WordExcel.toDocx(workbook);
const docxBytes = await Document.toBuffer(doc);

// 带选项：标题页、只取部分工作表、限制行列数。
// 请用名称指定工作表：数字位置已弃用（1.0 移除）——它在这里从 0 开始，
// 在 `Pdf.fromExcel` 中从 1 开始，在 `Workbook.getWorksheet` 中则是工作表 id。
const doc2 = WordExcel.toDocx(workbook, {
  includeTitlePage: true,
  sheets: ["Summary", "Regions"],
  maxRows: 100,
  maxColumns: 12,
  preserveFormatting: true
});

// 反向：把 DOCX 中的表格抽取为 Workbook
const extracted = WordExcel.extractTables(doc);
```

嵌入在 Word 文档中的图表也会桥接到 Excel 图表引擎
（27 个图表系列，经典图表与现代 ChartEx）；`Pdf.fromDocx` 默认就用它绘制图表
（见下文“图表”）。桥接构件均已公开：`WordExcel.toChartSvg(chart)`
渲染图表，`WordExcel.toChartModel(chart)` 返回 Excel 图表模型，
`WordExcel.toChartXlsx(series)` 生成内嵌数据工作簿（`.xlsx` 字节），
`WordExcel.toChartExXml(options)` 生成 ChartEx 部件。图表的自动系列颜色与方案颜色取自其文档主题：
`Pdf.fromDocx` 使用 `doc.theme`，`WordExcel.toChartSvg(chart, { themeColors: doc.theme?.colorScheme.colors })`
接受这些颜色——`Layout.toSvg` 也以 `context.themeColors` 把同样的颜色交给其 `renderChart` 回调。
从文件读取的图表写回时不丢失任何内容：`Chart` 没有字段表达的部分——方案颜色及其变换、渐变、线型、网格线、
文本格式、布局、逐个数据标签、数据引用、扩展列表、部件的样式与颜色部件——按读取时的原样写回（并据此绘制）；
设置某个字段只改写它所对应的元素（`values`、坐标轴的 `min`、`title`、`color`、`pointColors[i]` 等）。
`pointColors` 按数据点索引：第 `i` 项为第 `i` 个数据点的颜色，没有自身颜色的数据点为空位。

### Word → PDF

将 `DocxDocument` 转换为 PDF 字节。该桥接是共享 Word 布局引擎之上的薄层，
因此换行、分页、表格、行内图片、页眉/页脚和浮动元素的渲染结果
与 SVG 路径完全一致。

```typescript
import { Document } from "documonster/word";
import { Pdf } from "documonster/pdf";

const doc = await Document.read(docxBytes);
const pdfBytes = await Pdf.fromDocx(doc);

// 覆盖页面几何（单位：磅）。省略的字段会回退到
// 文档的 section properties，再回退到引擎默认值（US Letter，1 英寸）。
const pdf2 = await Pdf.fromDocx(doc, {
  pageWidth: 595, // A4 宽度
  pageHeight: 842, // A4 高度
  margins: { top: 72, right: 72, bottom: 72, left: 72 },
  headerMargin: 36, // 页眉带相对顶边的偏移
  footerMargin: 36, // 页脚带相对底边的偏移
  defaultFontFamily: "Helvetica",
  defaultFontSize: 11
});

// 与所有转换函数相同的文档选项：元数据（未给出的字段取自文档核心属性，包括关键词）、
// 用于可复现字节的固定创建日期，以及密码。
const pdf3 = await Pdf.fromDocx(doc, {
  title: "Quarterly report",
  creationDate: new Date("2026-01-01T00:00:00Z"),
  encryption: { ownerPassword: "owner", userPassword: "open" }
});
```

**图表。** 经典图表（`<c:chart>`）和现代 ChartEx（`<cx:chartSpace>` —— 旭日图、矩形树图、
瀑布图、漏斗图、箱线图、直方图、帕累托图、地图）默认就由 Excel 图表引擎渲染为完整矢量 PDF，
无需传入任何选项。引擎画不了的图表回退到行内 SVG，再回退到带标题的占位框
（不抛错、不留空白页）。要自己绘制经典图表，传入 `chartRenderer`：

```typescript
import { Pdf } from "documonster/pdf";

const pdf = await Pdf.fromDocx(doc, {
  chartRenderer: (chart, page, rect) => {
    if (chart.type !== "pie") return false; // 拒绝：交给内置渲染器绘制
    page.drawRect({ ...rect, stroke: { r: 0, g: 0, b: 0 } });
  }
});
```

`chartRenderer` 可以返回 `false` 来拒绝渲染某个图表；此时桥接会
回退到内置矢量渲染器，再回退到行内 SVG / 占位框。

## OOXML Strict 格式

本模块会自动处理以 ISO 29500 Strict 一致性级别保存的文档。
读取 Strict 格式的 .docx 时，命名空间 URI 和关系类型会被透明地
归一化为对应的 Transitional 形式 —— 无需任何用户操作。

## 兼容性

| 特性             | 支持情况                                      |
| ---------------- | --------------------------------------------- |
| .docx（读取）    | ✅ 广泛（常见元素结构化；未知部件保留）       |
| .docx（写入）    | ✅ 广泛（常见元素；未知部件以不透明形式写出） |
| .dotx（模板）    | ✅                                            |
| .docm（宏）      | ✅ 往返（保留 VBA，不执行/不编辑）            |
| .dotm（宏模板）  | ✅ 往返                                       |
| Flat OPC（.xml） | ✅                                            |
| ISO 29500 Strict | ✅ 自动归一化                                 |
| 加密的 .docx     | ✅ 用密码读写（Agile Encryption）             |
| 数字签名         | 🔍 检测与元数据提取（不签名/不验证）          |
| 浏览器           | ✅（同一个 `documonster/word` 导入）          |
| Node.js 22+      | ✅                                            |
