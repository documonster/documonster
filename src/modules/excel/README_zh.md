# Excel 模块

现代 TypeScript Excel 工作簿管理器 — 读取、操作和写入 XLSX 与 JSON 格式的电子表格，零运行时依赖。

## 功能特性

- **创建、读取和修改 XLSX 文件** — 完整 Open XML 支持
- **多工作表支持** — 添加、删除、重排序、复制
- **单元格样式** — 字体、颜色、边框、填充、对齐、数字格式
- **单元格合并和格式化** — 合并区域、富文本、超链接
- **行列属性** — 宽度、高度、隐藏、大纲级别、自动适应
- **冻结窗格和拆分视图** — 冻结行/列、在指定位置拆分
- **富文本支持** — 单个单元格内多种字体/样式
- **公式和计算值** — 共享公式、定义名称
- **数据验证** — 列表、整数、小数、日期、文本长度、自定义
- **条件格式** — 单元格值、色阶、数据条、图标集
- **图片** — JPEG、PNG、GIF，支持单单元格和双单元格锚点
- **超链接** — 内部链接、外部链接、邮件链接
- **数据透视表** — 读取和保留数据透视表定义
- **表格** — 自动筛选、汇总行、结构化引用
- **批注和备注** — 线程批注、旧版备注
- **复选框** — 表单控件和单元格级复选框
- **页面设置** — 打印区域、打印标题、页眉/页脚、分页符
- **数据保护** — 带密码的工作表保护（SHA-512）
- **流式处理** — `WorkbookReader` 和 `WorkbookWriter` 处理大文件
- **CSV 导入/导出** — `readCsv`、`writeCsv`、`readCsvFile`、`writeCsvFile`
- **Markdown 导入/导出** — `readMd`、`writeMd`、`readMdFile`、`writeMdFile`
- **PDF 导出** — `excelToPdf()`，完整支持样式、分页、字体、加密
- **浏览器支持** — `xlsx.load()`、`xlsx.writeBuffer()`，无需 polyfill

## 快速开始

### 创建工作簿

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("My Sheet");

// 添加数据
sheet.addRow(["姓名", "年龄", "邮箱"]);
sheet.addRow(["张三", 30, "zhang@example.com"]);
sheet.addRow(["李四", 25, "li@example.com"]);

// Node.js：写入文件
await workbook.xlsx.writeFile("output.xlsx");

// 浏览器：写入缓冲区
const buffer = await workbook.xlsx.writeBuffer();
```

### 读取工作簿

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// Node.js：从文件读取
await workbook.xlsx.readFile("input.xlsx");

// 浏览器：从 ArrayBuffer 读取
await workbook.xlsx.load(arrayBuffer);

const worksheet = workbook.getWorksheet(1);
worksheet.eachRow((row, rowNumber) => {
  console.log("行 " + rowNumber + " = " + JSON.stringify(row.values));
});
```

### 设置单元格样式

```typescript
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
cell.border = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" }
};
cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
cell.numFmt = "$#,##0.00";
```

### 数字格式

```typescript
// 货币
cell.numFmt = "$#,##0.00";

// 百分比
cell.numFmt = "0.00%";

// 日期
cell.numFmt = "yyyy-mm-dd";

// 自定义
cell.numFmt = '#,##0.00 "单位"';
```

### 富文本

```typescript
cell.value = {
  richText: [
    { text: "粗体 ", font: { bold: true } },
    { text: "和 ", font: {} },
    { text: "红色", font: { color: { argb: "FFFF0000" } } }
  ]
};
```

### 公式

```typescript
cell.value = { formula: "SUM(A1:A10)" };
cell.value = { formula: "A1+B1", result: 42 }; // 带缓存结果

// 共享公式
sheet.getCell("A1").value = { formula: "B1*2", shareType: "shared", ref: "A1:A10" };

// 定义名称
workbook.definedNames.add("MyRange", "Sheet1!$A$1:$B$10");
```

### 数据验证

```typescript
worksheet.getCell("A1").dataValidation = {
  type: "list",
  allowBlank: true,
  formulae: ['"选项1,选项2,选项3"']
};

worksheet.getCell("B1").dataValidation = {
  type: "whole",
  operator: "between",
  formulae: [1, 100],
  showErrorMessage: true,
  errorTitle: "无效",
  error: "请输入 1 到 100 之间的数字"
};
```

### 条件格式

```typescript
worksheet.addConditionalFormatting({
  ref: "A1:A100",
  rules: [
    {
      type: "cellIs",
      operator: "greaterThan",
      formulae: [90],
      style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF00FF00" } } },
      priority: 1
    }
  ]
});
```

### 图片

```typescript
import { readFileSync } from "fs";

const imageId = workbook.addImage({
  buffer: readFileSync("logo.png"),
  extension: "png"
});

worksheet.addImage(imageId, {
  tl: { col: 0, row: 0 },
  br: { col: 3, row: 5 }
});
```

### 表格

```typescript
worksheet.addTable({
  name: "SalesTable",
  ref: "A1",
  headerRow: true,
  totalsRow: true,
  columns: [
    { name: "产品", totalsRowLabel: "合计", filterButton: true },
    { name: "收入", totalsRowFunction: "sum", filterButton: true }
  ],
  rows: [
    ["小工具", 1000],
    ["大工具", 2500]
  ]
});
```

### 合并单元格

```typescript
worksheet.mergeCells("A1:D1");
worksheet.getCell("A1").value = "合并标题";
worksheet.getCell("A1").alignment = { horizontal: "center" };
```

### 冻结窗格

```typescript
// 冻结首行
worksheet.views = [{ state: "frozen", ySplit: 1 }];

// 冻结首列
worksheet.views = [{ state: "frozen", xSplit: 1 }];

// 同时冻结
worksheet.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
```

### 页面设置

```typescript
worksheet.pageSetup = {
  paperSize: 9, // A4
  orientation: "landscape",
  fitToPage: true,
  fitToWidth: 1,
  fitToHeight: 0,
  margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 }
};

// 打印区域
worksheet.pageSetup.printArea = "A1:G20";

// 打印标题（每页重复第 1-2 行）
worksheet.pageSetup.printTitlesRow = "1:2";
```

### 工作表保护

```typescript
await worksheet.protect("password123", {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  insertRows: false,
  deleteRows: false,
  sort: true,
  autoFilter: true
});
```

### 批注

```typescript
worksheet.getCell("A1").note = "简单批注";

worksheet.getCell("B1").note = {
  texts: [{ text: "作者：", font: { bold: true } }, { text: "这是一个富文本批注" }]
};
```

### 自动适应列宽

```typescript
worksheet.columns.forEach(column => {
  column.width = column.values
    ? Math.max(...column.values.map(v => String(v ?? "").length)) + 2
    : 10;
});
```

## PDF 导出

零外部依赖将任意工作簿导出为 PDF：

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("报告");
sheet.columns = [
  { header: "产品", key: "product", width: 20 },
  { header: "收入", key: "revenue", width: 15 }
];
sheet.addRow({ product: "小工具", revenue: 1000 });
sheet.getColumn("revenue").numFmt = "$#,##0.00";

const pdf = excelToPdf(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "销售报告"
});

// Node.js
import { writeFileSync } from "fs";
writeFileSync("report.pdf", pdf);

// 浏览器
const blob = new Blob([pdf], { type: "application/pdf" });
window.open(URL.createObjectURL(blob));
```

### XLSX 转 PDF

```typescript
const workbook = new Workbook();
await workbook.xlsx.readFile("input.xlsx");
const pdf = excelToPdf(workbook);
```

### PDF 加密

```typescript
const pdf = excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "reader",
    permissions: { print: true, copy: false }
  }
});
```

### Unicode / CJK 字体嵌入

```typescript
import { readFileSync } from "fs";
const pdf = excelToPdf(workbook, {
  font: readFileSync("NotoSansSC-Regular.ttf")
});
```

## CSV 导入/导出

```typescript
import { Workbook } from "@cj-tech-master/excelts";
import fs from "fs";

const workbook = new Workbook();

// Node.js：读写 CSV 文件
await workbook.readCsvFile("data.csv");
await workbook.writeCsvFile("output.csv");

// 从流读取 CSV
await workbook.readCsv(fs.createReadStream("data.csv"), { sheetName: "导入数据" });

// 写入 CSV 到流
await workbook.writeCsv(fs.createWriteStream("output.csv"));

// 写入 CSV 到字符串 / 字节
const csvText = workbook.writeCsv();
const bytes = await workbook.writeCsvBuffer();

// 浏览器：从字符串/ArrayBuffer/File 读取
await workbook.readCsv(csvString);
await workbook.readCsv(arrayBuffer);
```

## Markdown 导入/导出

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// 读取 Markdown 表格
workbook.readMd("| 姓名 | 年龄 |\n| --- | --- |\n| Alice | 30 |");
await workbook.readMdFile("table.md");

// 写入 Markdown
const mdText = workbook.writeMd();
await workbook.writeMdFile("output.md");
const bytes = workbook.writeMdBuffer();
```

## 流式 API

### 流式读取器

以最小内存使用量读取大型 XLSX 文件：

```typescript
import { WorkbookReader } from "@cj-tech-master/excelts";

const reader = new WorkbookReader("large-file.xlsx", {
  worksheets: "emit",
  sharedStrings: "cache",
  hyperlinks: "ignore",
  styles: "ignore"
});

for await (const worksheet of reader) {
  console.log(`正在读取：${worksheet.name}`);
  for await (const row of worksheet) {
    console.log(row.values);
  }
}
```

### 流式写入器

逐行写入大型 XLSX 文件：

```typescript
import { WorkbookWriter } from "@cj-tech-master/excelts";

const workbook = new WorkbookWriter({
  filename: "output.xlsx",
  useSharedStrings: true,
  useStyles: true
});

const sheet = workbook.addWorksheet("数据");
for (let i = 0; i < 1000000; i++) {
  sheet.addRow([`行 ${i}`, i, new Date()]).commit();
}

sheet.commit();
await workbook.commit();
```

### Web Streams（Node.js 22+ 和浏览器）

```typescript
import { WorkbookWriter, WorkbookReader } from "@cj-tech-master/excelts";

// 写入到 Web WritableStream
const chunks: Uint8Array[] = [];
const writable = new WritableStream({
  write(chunk) {
    chunks.push(chunk);
  }
});

const writer = new WorkbookWriter({ stream: writable });
const sheet = writer.addWorksheet("Sheet1");
sheet.addRow(["姓名", "分数"]).commit();
sheet.addRow(["Alice", 98]).commit();
await sheet.commit();
await writer.commit();

// 从 Web ReadableStream 读取
const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0));
let offset = 0;
for (const c of chunks) {
  bytes.set(c, offset);
  offset += c.length;
}

const readable = new ReadableStream({
  start(ctrl) {
    ctrl.enqueue(bytes);
    ctrl.close();
  }
});

const reader = new WorkbookReader(readable, { worksheets: "emit" });
for await (const ws of reader) {
  for await (const row of ws) {
    console.log(row.values);
  }
}
```

## 浏览器支持

### 使用打包工具（Vite、Webpack、Rollup、esbuild）

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");
sheet.getCell("A1").value = "你好，浏览器！";

const buffer = await workbook.xlsx.writeBuffer();
const blob = new Blob([buffer], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});
const url = URL.createObjectURL(blob);
```

### 使用 Script 标签

```html
<script src="https://unpkg.com/@cj-tech-master/excelts/dist/iife/excelts.iife.min.js"></script>
<script>
  const { Workbook } = ExcelTS;
  const wb = new Workbook();
</script>
```

### 浏览器注意事项

- 使用 `xlsx.load(arrayBuffer)` 代替 `xlsx.readFile()`
- 使用 `xlsx.writeBuffer()` 代替 `xlsx.writeFile()`
- PDF 导出完全支持
- CSV 和 Markdown 操作完全支持
- 工作表密码保护使用纯 JS SHA-512

## 工具导出

```typescript
import {
  // 日期转换
  dateToExcel,
  excelToDate,
  DateParser,
  DateFormatter,

  // 二进制工具
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
  pdf,
  excelToPdf,
  PageSizes,
  PdfError,
  isPdfError,

  // 错误处理
  BaseError,
  ExcelError,
  toError,
  errorToJSON,
  getErrorChain,
  getRootCause
} from "@cj-tech-master/excelts";
```

## 示例

查看 [examples 目录](examples/) 获取覆盖所有功能的可运行代码：

- 工作簿创建、读取和复制
- 单元格样式、字体、边框、填充
- 公式、数据验证、条件格式
- 图片（JPEG、PNG）、超链接、批注
- 带自动筛选和汇总的表格
- 合并单元格、冻结窗格、页面设置
- 流式读取器和写入器
- Web Streams 集成
- PDF 导出
- 更多...
