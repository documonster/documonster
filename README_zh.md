# ExcelTS

[![Build Status](https://github.com/cjnoname/excelts/actions/workflows/ci.yml/badge.svg?branch=main&event=push)](https://github.com/cjnoname/excelts/actions/workflows/ci.yml)

现代化的 TypeScript Excel 工作簿管理器 - 读取、操作和写入电子表格数据和样式到 XLSX 和 JSON。

## 关于本项目

ExcelTS 是零依赖的 TypeScript 电子表格与文档工具包：

- 🚀 **零运行时依赖** — 纯 TypeScript 实现，无任何外部包依赖
- 📦 **七大模块** — Excel（XLSX/JSON）、PDF（独立引擎 + Excel 桥接）、CSV（RFC 4180）、Markdown（GFM 表格）、XML（SAX/DOM/Writer）、Archive（ZIP/TAR）、Stream（跨平台）
- 🤖 **AI 最佳伴侣** — 简洁一致的 API 设计，让 AI 编程助手更容易生成处理各种文件的代码——一个库搞定 Excel、PDF、CSV、Markdown、XML、ZIP 和流式处理，保持最高效率
- ✅ **跨平台** — Node.js 22+、Bun、Chrome 89+、Firefox 102+、Safari 14.1+
- ✅ **ESM 优先** — 原生 ES Modules，兼容 CommonJS，完整 tree-shaking 支持

## 模块文档

- [XML 模块](src/modules/xml/README.md) — 零依赖 SAX/DOM 解析器、查询引擎和双模式写入器
- [PDF 模块](src/modules/pdf/README.md) — 零依赖 PDF 引擎，支持加密和字体嵌入
- [CSV 模块](src/modules/csv/README.md) — RFC 4180 解析/格式化、流式处理、数据生成
- [Markdown 模块](src/modules/md/README.md) — GFM 表格解析/格式化，支持对齐方式往返
- [归档模块](src/modules/archive/README.md) — ZIP/TAR 创建/读取/编辑、压缩、加密
- [流模块](src/modules/stream/README.md) — 跨平台 Readable/Writable/Transform/Duplex

## 翻译

- [English Documentation](README.md)

## 安装

```bash
npm install @cj-tech-master/excelts
```

## 快速开始

### 创建工作簿

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("我的工作表");

// 添加数据
sheet.addRow(["姓名", "年龄", "邮箱"]);
sheet.addRow(["张三", 30, "zhang@example.com"]);
sheet.addRow(["李四", 25, "li@example.com"]);

// 保存文件
// 仅 Node.js：写入到文件路径
await workbook.xlsx.writeFile("output.xlsx");

// 浏览器：使用 `writeBuffer()` 并保存为 Blob（见「浏览器支持」章节）
```

### 读取工作簿

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
// 仅 Node.js：从文件路径读取
await workbook.xlsx.readFile("input.xlsx");

// 浏览器：使用 `xlsx.load(arrayBuffer)`（见「浏览器支持」章节）

const worksheet = workbook.getWorksheet(1);
worksheet.eachRow((row, rowNumber) => {
  console.log("第 " + rowNumber + " 行 = " + JSON.stringify(row.values));
});
```

### 单元格样式

```javascript
// 设置单元格值和样式
const cell = worksheet.getCell("A1");
cell.value = "你好";
cell.font = {
  name: "Arial",
  size: 16,
  bold: true,
  color: { argb: "FFFF0000" }
};
cell.fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFFF00" }
};
```

## 功能特性

- **Excel 操作**
  - 创建、读取和修改 XLSX 文件
  - 多工作表支持
  - 单元格样式（字体、颜色、边框、填充）
  - 单元格合并和格式化
  - 行和列属性
  - 冻结窗格和拆分视图

- **数据处理**
  - 富文本支持
  - 公式和计算值
  - 数据验证
  - 条件格式
  - 图片和图表
  - 超链接
  - 数据透视表

- **PDF 导出**
  - 零依赖 Excel 转 PDF
  - 完整的单元格样式支持（字体、颜色、边框、填充、对齐）
  - 自动分页与重复表头行
  - TrueType 字体嵌入，支持 Unicode/CJK 中文文本
  - JPEG 和 PNG 图片嵌入，支持透明度
  - 密码保护和加密
  - 每个工作表独立的页面设置（纸张大小、方向、页边距）
  - 支持 Tree-shaking（不导入 = 不打包）

- **高级功能**
  - 大文件流式处理
  - CSV 导入/导出
  - Markdown 表格导入/导出
  - 带自动筛选的表格
  - 页面设置和打印选项
  - 数据保护
  - 注释和批注

## 子路径导出

ExcelTS 提供独立模块的子路径导出：

```typescript
// 主入口 - Excel 核心（Workbook, Worksheet, Cell 等）
import { Workbook, WorkbookWriter } from "@cj-tech-master/excelts";

// XML 工具包（SAX 解析器、DOM 解析器、查询引擎、写入器）
import { SaxParser, parseXml, XmlWriter, query } from "@cj-tech-master/excelts/xml";

// ZIP/TAR 归档工具
import { zip, unzip, ZipArchive, compress } from "@cj-tech-master/excelts/zip";

// CSV 解析、格式化与流式处理
import { parseCsv, formatCsv, CsvParserStream } from "@cj-tech-master/excelts/csv";

// Markdown 表格解析与格式化
import { parseMd, formatMd, parseMdAll } from "@cj-tech-master/excelts/md";

// 跨平台流式原语
import { Readable, pipeline, createTransform } from "@cj-tech-master/excelts/stream";
```

每个子路径支持 `browser`、`import`（ESM）和 `require`（CJS）条件。

## PDF 导出

零依赖将任意工作簿导出为 PDF：

```javascript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("报告");
sheet.columns = [
  { header: "产品", key: "product", width: 20 },
  { header: "收入", key: "revenue", width: 15 }
];
sheet.addRow({ product: "组件A", revenue: 1000 });
sheet.getColumn("revenue").numFmt = "¥#,##0.00";

// 一行导出
const pdf = excelToPdf(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "销售报告"
});

// Node.js：写入文件
import { writeFileSync } from "fs";
writeFileSync("report.pdf", pdf);

// 浏览器：下载
const blob = new Blob([pdf], { type: "application/pdf" });
const url = URL.createObjectURL(blob);
window.open(url);
```

### 将现有 XLSX 转换为 PDF

```javascript
const workbook = new Workbook();
await workbook.xlsx.readFile("input.xlsx");
const pdf = excelToPdf(workbook);
```

### 加密

```javascript
const pdf = excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "reader",
    permissions: { print: true, copy: false }
  }
});
```

### Unicode / CJK 中文支持

```javascript
import { readFileSync } from "fs";

const pdf = excelToPdf(workbook, {
  font: readFileSync("NotoSansSC-Regular.ttf") // 嵌入 TrueType 字体以支持中文
});
```

### 独立 PDF 生成（无需 Excel）

无需工作簿，直接从数组数据生成 PDF：

```javascript
import { pdf } from "@cj-tech-master/excelts/pdf";

// 最简方式 — 传入二维数组
const bytes = pdf([
  ["产品", "收入"],
  ["组件A", 1000],
  ["组件B", 2500]
]);

// 带列宽和样式
const bytes = pdf(
  {
    name: "报告",
    columns: [
      { width: 25, header: "产品" },
      { width: 15, header: "收入" }
    ],
    data: [
      ["组件A", 1000],
      ["组件B", 2500]
    ]
  },
  { showGridLines: true }
);
```

完整 API 参考和所有选项请查看 [PDF 模块文档](src/modules/pdf/README.md)。

## 归档工具（ZIP/TAR）

ExcelTS 内置 ZIP/TAR 归档工具（用于 XLSX 管线）。若直接使用归档相关 API，
可通过 `ZipStringEncoding` 自定义 ZIP 字符串编码：

- 默认：`"utf-8"`
- 传统兼容：`"cp437"`
- 自定义：提供带 `encode`/`decode` 的 codec，并可设置可选标记

当使用非 UTF-8 编码时，可写入 Unicode Extra Field 以提升跨工具兼容性。

## XML 工具包

ExcelTS 内置一个独立的、零依赖的 XML 模块，支持流式和缓冲两种解析/写入模式。它为 XLSX 管线提供底层支持，也可作为独立子路径导出使用。

### 核心特性

- **SAX 解析器** — 事件驱动的高性能流式解析器
- **DOM 解析器** — 从字符串构建可查询的 XML 树
- **查询引擎** — 简化路径表达式（`a/b[@id='1']`、`a//c`、`a/b[0]`）
- **双模式写入** — 缓冲模式（`XmlWriter`）和流式模式（`XmlStreamWriter`）
- **完整命名空间支持** — 前缀解析、保留命名空间强制、未绑定前缀检测
- **安全加固** — 实体展开限制、嵌套深度限制、重复属性拒绝、BOM 处理

### 快速示例

```typescript
import { SaxParser, parseXml, XmlWriter, queryAll } from "@cj-tech-master/excelts/xml";

// SAX 流式解析
const parser = new SaxParser();
parser.on("opentag", tag => console.log(tag.name, tag.attributes));
parser.write('<root><item id="1">hello</item></root>');
parser.close();

// DOM 解析 + 查询
const doc = parseXml("<root><a><b>1</b><b>2</b></a></root>");
const items = queryAll(doc.root, "a/b"); // 所有 <b> 元素

// 写入 XML
const w = new XmlWriter();
w.openXml();
w.openNode("root");
w.leafNode("item", { id: "1" }, "hello");
w.closeNode();
console.log(w.xml);
```

完整 API 参考请见 [XML 模块文档](src/modules/xml/README.md)。

### 编辑现有 ZIP（ZipEditor）

ExcelTS 也提供 ZIP 编辑器：对已有压缩包做类似文件系统的编辑，然后输出一个新的 ZIP。

- 支持 `set()` / `delete()` / `rename()` / `deleteDirectory()` / `setComment()`
- 对未改动条目会尽量走高效 passthrough（可用时不重压缩）

```js
import { editZip } from "@cj-tech-master/excelts";

const editor = await editZip(existingZipBytes, {
  reproducible: true,

  // 未改动条目的保留策略：
  // - "strict"（默认）：必须能读取 raw passthrough，否则直接抛错
  // - "best-effort"：raw 不可用时退化为 extract + 重新写入（可能更耗 CPU/内存）
  preserve: "best-effort",
  onWarning: w => console.warn(w.code, w.entry, w.message)
});

editor.delete("old.txt");
editor.rename("a.txt", "renamed.txt");
editor.set("new.txt", "hello");

const out = await editor.bytes();
```

## 流式 API

处理大型 Excel 文件时无需将整个文件加载到内存中，ExcelTS 提供了流式读写 API。

- **Node.js**：`WorkbookReader` 支持从文件路径读取；`WorkbookWriter` 支持写入到 `filename`。
- **浏览器**：读取使用 `Uint8Array` / `ArrayBuffer` / Web `ReadableStream<Uint8Array>`；写入使用 Web `WritableStream<Uint8Array>`。
- 说明：ExcelTS 不再从主入口 re-export 内部那套 stream 工具类（如 `Readable`、`Writable`）。建议直接使用标准 Web Streams（浏览器/Node 22+）或 Node.js 原生 streams。

### 流式读取器

以最小内存占用读取大型 XLSX 文件：

```javascript
import { WorkbookReader } from "@cj-tech-master/excelts";

// Node.js：从文件路径读取
const reader = new WorkbookReader("large-file.xlsx", {
  worksheets: "emit", // 触发工作表事件
  sharedStrings: "cache", // 缓存共享字符串以获取单元格值
  hyperlinks: "ignore", // 忽略超链接
  styles: "ignore" // 忽略样式以加快解析
});

for await (const worksheet of reader) {
  console.log(`正在读取: ${worksheet.name}`);
  for await (const row of worksheet) {
    console.log(row.values);
  }
}
```

### Web Streams（Node.js 22+ 与浏览器）

`WorkbookWriter` 支持写入 Web `WritableStream<Uint8Array>`，`WorkbookReader` 支持从 Web `ReadableStream<Uint8Array>` 读取。

这里直接使用标准 Web Streams API，**不需要**从 ExcelTS 额外导入一堆 stream 工具类。

- 可运行完整示例: [src/modules/excel/examples/web-streams-reader-writer.ts](src/modules/excel/examples/web-streams-reader-writer.ts)

本地运行（Node.js 22+）:

```bash
npx tsx src/modules/excel/examples/web-streams-reader-writer.ts
```

最小端到端示例:

```javascript
import { WorkbookWriter, WorkbookReader } from "@cj-tech-master/excelts";

// 1) 写入工作簿 -> Web WritableStream
const chunks = [];
const writable = new WritableStream({
  write(chunk) {
    chunks.push(chunk);
  }
});

const writer = new WorkbookWriter({ stream: writable });
const sheet = writer.addWorksheet("Sheet1");
sheet.addRow(["Name", "Score"]).commit();
sheet.addRow(["Alice", 98]).commit();
await sheet.commit();
await writer.commit();

// 2) 读取工作簿 <- Web ReadableStream
const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
let offset = 0;
for (const c of chunks) {
  bytes.set(c, offset);
  offset += c.length;
}

const readable = new ReadableStream({
  start(controller) {
    controller.enqueue(bytes);
    controller.close();
  }
});

const reader = new WorkbookReader(readable, { worksheets: "emit" });
for await (const ws of reader) {
  for await (const row of ws) {
    console.log(row.values);
  }
}
```

### 流式写入器

逐行写入大型 XLSX 文件：

```javascript
import { WorkbookWriter } from "@cj-tech-master/excelts";

// Node.js：写入到 filename
const workbook = new WorkbookWriter({
  filename: "output.xlsx",
  useSharedStrings: true,
  useStyles: true
});

const sheet = workbook.addWorksheet("Data");

// 逐行写入
for (let i = 0; i < 1000000; i++) {
  sheet.addRow([`第 ${i} 行`, i, new Date()]).commit();
}

// 提交工作表并完成
sheet.commit();
await workbook.commit();
```

## CSV 支持

### Node.js（完整流式支持）

```javascript
import { Workbook } from "@cj-tech-master/excelts";
import fs from "fs";

const workbook = new Workbook();

// 从文件读取 CSV
await workbook.readCsvFile("data.csv");

// 从流读取 CSV
const stream = fs.createReadStream("data.csv");
await workbook.readCsv(stream, { sheetName: "Imported" });

// 写入 CSV 到文件
await workbook.writeCsvFile("output.csv");

// 写入 CSV 到流
const writeStream = fs.createWriteStream("output.csv");
await workbook.writeCsv(writeStream);

// 写入 CSV 为字符串 / 字节
const csvText = workbook.writeCsv();
const bytes = await workbook.writeCsvBuffer();
```

### 浏览器（内存中）

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// 从字符串读取 CSV
await workbook.readCsv(csvString);

// 从 ArrayBuffer 读取 CSV（例如从 fetch）
const response = await fetch("data.csv");
const arrayBuffer = await response.arrayBuffer();
await workbook.readCsv(arrayBuffer);

// 从 File 读取 CSV（例如 <input type="file">）
await workbook.readCsv(file);

// 写入 CSV 为字符串
const csvOutput = workbook.writeCsv();

// 写入 CSV 为 Uint8Array 字节
const bytes = await workbook.writeCsvBuffer();
```

## Markdown 支持

### Node.js

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// 从文件读取 Markdown 表格
await workbook.readMdFile("table.md");

// 从字符串读取 Markdown 表格
workbook.readMd("| Name | Age |\n| --- | --- |\n| Alice | 30 |");

// 使用选项读取
workbook.readMd(mdString, {
  sheetName: "Data",
  map: (v, col) => Number(v) || v
});

// 写入 Markdown 到文件
await workbook.writeMdFile("output.md");

// 写入 Markdown 为字符串
const mdText = workbook.writeMd();

// 写入 Markdown 为 Uint8Array 字节
const bytes = workbook.writeMdBuffer();
```

### 浏览器（内存中）

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// 从字符串读取 Markdown 表格
workbook.readMd(mdString);

// 写入 Markdown 为字符串
const mdOutput = workbook.writeMd();

// 写入 Markdown 为 Uint8Array 字节
const bytes = workbook.writeMdBuffer();
```

### 独立解析器 / 格式化器

```typescript
import { parseMd, formatMd, parseMdAll } from "@cj-tech-master/excelts/md";

// 解析 Markdown 表格
const result = parseMd("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
// result.headers = ["Name", "Age"]
// result.rows = [["Alice", "30"]]
// result.alignments = ["none", "none"]

// 解析文档中的所有表格
const tables = parseMdAll(markdownDoc);

// 格式化数据为 Markdown 表格
const md = formatMd(
  ["Name", "Age"],
  [
    ["Alice", "30"],
    ["Bob", "25"]
  ],
  { alignment: "left", padding: true }
);
```

## 浏览器支持

ExcelTS 原生支持浏览器环境，现代打包工具**无需任何配置**。

### 在打包工具中使用（Vite, Webpack, Rollup, esbuild）

直接导入 ExcelTS - 无需 polyfills 或额外配置：

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");
sheet.getCell("A1").value = "你好，浏览器！";

// 写入 buffer 并下载
const buffer = await workbook.xlsx.writeBuffer();
const blob = new Blob([buffer], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});
const url = URL.createObjectURL(blob);
// ... 触发下载
```

### 使用 Script 标签（无打包工具）

```html
<script src="https://unpkg.com/@cj-tech-master/excelts/dist/iife/excelts.iife.min.js"></script>
<script>
  const { Workbook } = ExcelTS;
  const wb = new Workbook();
  // ... 使用 workbook API
</script>
```

### 手动浏览器示例（本地）

如果你想在真实浏览器里做快速手动冒烟测试（创建/下载/读取 XLSX、工作表保护等），可以用这个页面：

- [src/modules/excel/examples/browser-smoke.html](src/modules/excel/examples/browser-smoke.html)

使用方式：

```bash
npm run build:browser:bundle
npx serve .
```

然后打开 `http://localhost:3000/src/modules/excel/examples/browser-smoke.html`。

### 浏览器版本注意事项

- **完全支持 PDF 导出**（浏览器端零配置即可使用）
- **支持 CSV 操作**（使用原生 RFC 4180 标准实现）
  - 使用 `await workbook.readCsv(input)` 读取 CSV
  - 使用 `workbook.writeCsv()` 或 `await workbook.writeCsvBuffer()` 写入 CSV
- **支持 Markdown 表格操作**（GFM 表格语法）
  - 使用 `workbook.readMd(input)` 读取 Markdown 表格
  - 使用 `workbook.writeMd()` 或 `workbook.writeMdBuffer()` 写入 Markdown
- 使用 `xlsx.load(arrayBuffer)` 代替 `xlsx.readFile()`
- 使用 `xlsx.writeBuffer()` 代替 `xlsx.writeFile()`
- 完全支持带密码的工作表保护（纯 JS SHA-512 实现）

## 工具导出

主入口还导出了常用工具函数：

```typescript
import {
  // Excel 日期转换
  dateToExcel, // JS Date -> Excel 序列号
  excelToDate, // Excel 序列号 -> JS Date

  // 日期解析/格式化（高性能，零依赖）
  DateParser, // 批量日期解析器，支持格式自动检测
  DateFormatter, // 批量日期格式化器

  // 二进制工具（跨平台）
  base64ToUint8Array,
  uint8ArrayToBase64,
  concatUint8Arrays,
  toUint8Array,
  stringToUint8Array,
  uint8ArrayToString,

  // XML 工具
  xmlEncode,
  xmlDecode,
  xmlEncodeAttr,
  validateXmlName,

  // PDF 导出
  pdf, // 最简单：pdf([["A", 1], ["B", 2]]) → Uint8Array
  excelToPdf, // Workbook -> Uint8Array（Excel 转 PDF）
  PageSizes, // 内置页面尺寸定义
  PdfError, // PDF 基础错误
  PdfRenderError, // 布局/渲染错误
  PdfFontError, // 字体解析/嵌入错误
  PdfStructureError, // PDF 结构组装错误
  isPdfError, // PDF 错误类型守卫

  // 错误基础设施
  BaseError, // 所有库错误的基类
  ExcelError, // Excel 基础错误（支持 instanceof 检查）
  toError, // 将 unknown 标准化为 Error
  errorToJSON, // 序列化错误（包含 cause 链）
  getErrorChain, // 获取完整的错误 cause 链数组
  getRootCause // 获取 cause 链中最深层的错误
} from "@cj-tech-master/excelts";
```

## 系统要求

### Node.js

- **Node.js >= 22.0.0**（原生支持 ES2020）

### 浏览器（无需 Polyfills）

- **Chrome >= 89**（2021年3月）
- **Edge >= 89**（2021年3月）
- **Firefox >= 102**（2022年6月）
- **Safari >= 14.1**（2021年4月）
- **Opera >= 75**（2021年3月）

对于不支持原生 `CompressionStream` API 的旧浏览器（Firefox < 113, Safari < 16.4），ExcelTS 会自动使用内置的纯 JavaScript DEFLATE 实现 - 无需任何配置或 polyfills。

浏览器端不要求必须支持 `crypto.randomUUID()`：ExcelTS 内置了 UUID v4 生成器，并会优先使用 `crypto.getRandomValues()` 作为 fallback。

## API 文档

详细的 API 文档，请参考以下综合文档部分：

- 工作簿管理
- 工作表
- 单元格和值
- 样式
- 公式
- 数据验证
- 条件格式
- 文件输入输出

## 许可证

MIT License

详见 LICENSE。

第三方软件的声明与归属信息请见 THIRD_PARTY_NOTICES.md。

## 链接

- [GitHub 仓库](https://github.com/cjnoname/excelts)
- [问题跟踪](https://github.com/cjnoname/excelts/issues)

## 更新日志

详细版本历史请查看 [CHANGELOG.md](CHANGELOG.md)。
