# PDF 模块

[English](README.md)

功能完整的零依赖 PDF 引擎，纯 TypeScript 从零构建。使用 `pdf()` 函数或 `excelToPdf()` 桥接 API **写入** PDF。使用 `readPdf()` **读取**任意 PDF — 从所有主流 PDF 版本中提取文本、图片和元数据。所有 API 均为异步，在页间让出事件循环以避免阻塞。

```typescript
// 写入 — 独立使用
import { pdf } from "@cj-tech-master/excelts/pdf";

// 写入 — 从 Excel 转换
import { excelToPdf } from "@cj-tech-master/excelts/pdf";

// 读取 — 提取文本、图片、元数据
import { readPdf } from "@cj-tech-master/excelts/pdf";
```

## 功能特性

### 写入

- **零依赖** — 纯 TypeScript PDF 生成，无外部库
- **PDF 2.0** — 生成现代 PDF 2.0 格式
- **独立引擎** — 使用 `pdf()` 函数配合普通数组和对象，无需 Excel 依赖
- **Excel 桥接** — `excelToPdf(workbook)` 一行代码完成 Excel 转 PDF
- **跨平台** — Node.js 和浏览器使用同一 API
- **完整样式** — 字体、颜色、边框、填充、对齐、合并单元格
- **富文本** — 单元格内混合格式，支持自动换行
- **自动分页** — 垂直/水平自动分页，支持重复标题行
- **图片** — JPEG 和 PNG 嵌入，支持 alpha 透明度
- **AES-256 加密** — 密码保护，AES-256 (V=5, R=5) 加密与权限控制
- **字体嵌入** — TrueType 字体子集化，支持 Unicode/CJK 文本
- **页面设置** — 每工作表独立纸张大小、方向、边距、打印区域
- **Tree-Shakeable** — 未导入则不打包
- **非阻塞** — 在页间让出事件循环，避免阻塞

### 读取

- **通用读取器** — 读取所有主流 PDF 版本 (1.0 至 2.0)
- **文本提取** — 完整文本及行重建、多栏检测、表格识别
- **多语言** — WinAnsi、MacRoman、CJK (ToUnicode CMap)、Identity-H/V、Symbol、ZapfDingbats
- **图片提取** — JPEG、JPEG2000、CCITT、JBIG2、raw/Flate，支持 SMask/alpha
- **标注提取** — 链接、批注、高亮、图章、自由文本等
- **表单字段** — AcroForm 提取：文本输入、复选框、单选按钮、下拉列表、签名
- **元数据** — Info 字典 + XMP（标题、作者、日期、页数、页面尺寸）
- **全加密格式** — RC4-40、RC4-128、AES-128、AES-256（读取所有版本）
- **容错** — 交叉引用表/流恢复，增量更新支持

---

## 快速开始

### 读取 PDF

```typescript
import { readPdf } from "@cj-tech-master/excelts/pdf";
import { readFileSync } from "fs";

const bytes = readFileSync("document.pdf");
const result = await readPdf(bytes);

// All text
console.log(result.text);

// Per-page text
for (const page of result.pages) {
  console.log(`Page ${page.pageNumber}: ${page.text.length} chars`);
}

// Metadata
console.log(result.metadata.title);
console.log(result.metadata.author);
console.log(result.metadata.pageCount);

// Images
for (const page of result.pages) {
  for (const img of page.images) {
    console.log(img.format, img.width, img.height);
  }
}

// Annotations (links, comments, highlights)
for (const page of result.pages) {
  for (const annot of page.annotations) {
    console.log(annot.subtype, annot.contents, annot.uri);
  }
}

// Form fields
for (const field of result.formFields) {
  console.log(field.name, field.type, field.value);
}
```

### 读取加密 PDF

```typescript
const result = await readPdf(bytes, { password: "secret" });
```

### 选择性提取

```typescript
// Only pages 1 and 3, text only (no images)
const result = await readPdf(bytes, {
  pages: [1, 3],
  extractImages: false
});
```

### Excel 转 PDF（桥接 API）

从 Excel 工作簿生成 PDF 的最简方式：

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sales");
sheet.columns = [
  { header: "Product", key: "product", width: 20 },
  { header: "Revenue", key: "revenue", width: 15 }
];
sheet.addRow({ product: "Widget", revenue: 1000 });
sheet.addRow({ product: "Gadget", revenue: 2500 });
sheet.getColumn("revenue").numFmt = "$#,##0.00";

const pdf = await excelToPdf(workbook);

// Node.js
import { writeFileSync } from "fs";
writeFileSync("output.pdf", pdf);

// Browser
const blob = new Blob([pdf], { type: "application/pdf" });
const url = URL.createObjectURL(blob);
window.open(url);
```

### 读取 XLSX，导出 PDF

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
await workbook.xlsx.readFile("report.xlsx");

const pdf = await excelToPdf(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "Monthly Report"
});
```

### 独立 PDF（无需 Excel）

从纯数据生成 PDF — 无需 Excel 模块，无需 Map 对象，无需样板代码：

```typescript
import { pdf } from "@cj-tech-master/excelts/pdf";

// Simplest — pass a 2D array
const bytes = await pdf([
  ["Product", "Revenue"],
  ["Widget", 1000],
  ["Gadget", 2500]
]);

// With options
const bytes = await pdf(
  [
    ["Name", "Score"],
    ["Alice", 95],
    ["Bob", 87]
  ],
  { showGridLines: true, title: "Scores" }
);

// Multiple sheets
const bytes = await pdf({
  sheets: [
    {
      name: "Sales",
      data: [
        ["Product", "Revenue"],
        ["Widget", 1000]
      ]
    },
    {
      name: "Costs",
      data: [
        ["Item", "Amount"],
        ["Rent", 500]
      ]
    }
  ]
});

// Column widths + styled cells
const bytes = await pdf({
  name: "Report",
  columns: [{ width: 25 }, { width: 15 }],
  data: [
    [
      { value: "Product", bold: true },
      { value: "Revenue", bold: true }
    ],
    ["Widget", 1000],
    ["Gadget", 2500]
  ]
});
```

---

## 架构

PDF 模块分为四个层次：

```
src/modules/pdf/
├── core/               # PDF primitives (objects, streams, writer, encryption, crypto)
├── font/               # TTF parsing, glyph metrics, font subsetting, embedding
├── render/             # Layout engine, page renderer, style converter
│   ├── layout-engine   — PdfSheetData → LayoutPage[] (zero @excel imports)
│   ├── page-renderer   — LayoutPage → PDF content stream (zero @excel imports)
│   ├── style-converter — PdfCellStyle → PDF rendering params (zero @excel imports)
│   ├── png-decoder     — PNG image decoding for PDF embedding (zero @excel imports)
│   └── pdf-exporter    — PdfWorkbook → Uint8Array (zero @excel imports)
├── reader/             # PDF reader — tokenizer, parser, decryption, text/image extraction
│   ├── pdf-tokenizer   — byte-level PDF tokenization
│   ├── pdf-parser      — objects, xref tables/streams, trailer
│   ├── pdf-document    — document structure, page tree, object resolution
│   ├── pdf-decrypt     — RC4/AES decryption for all PDF encryption versions
│   ├── stream-filters  — Flate, ASCII85, ASCIIHex, LZW, RunLength decoders
│   ├── cmap-parser     — ToUnicode CMap parsing with variable-length codespace
│   ├── font-decoder    — Type1, TrueType, Type0/CID, Symbol, ZapfDingbats
│   ├── content-interpreter — BT/ET, Tj/TJ, Tm/Td, Form XObject, inline images
│   ├── text-reconstruction — line building, table/multi-column detection, RTL
│   ├── image-extractor — JPEG, JPEG2000, CCITT, JBIG2, raw, SMask
│   ├── annotation-extractor — Link, Text, Highlight, FreeText, Stamp, etc.
│   ├── form-extractor — AcroForm: text, checkbox, radio, dropdown, listbox, signature
│   ├── metadata-reader — Info dict + XMP metadata
│   ├── reader-utils    — shared reader utility functions
│   └── pdf-reader      — public API: readPdf()
├── types.ts            # PdfWorkbook, PdfSheetData, PdfCellData, etc.
├── excel-bridge.ts     # Excel Workbook → PdfWorkbook conversion (ONLY @excel dependency)
└── index.ts
```

整个 PDF 引擎（core、font、render、reader）**零导入 Excel 模块**。`excel-bridge.ts` 是唯一与 Excel 交互的文件 — 它将 `Workbook` 转换为 `PdfWorkbook`。

**写入策略：** 仅写入 PDF 2.0（现代格式，AES-256）。
**读取策略：** 读取所有主流 PDF 版本（1.0 至 2.0，所有加密类型）。

---

## 写入选项

```typescript
interface PdfExportOptions {
  // Page layout
  pageSize?: PageSizeName | PdfPageSize; // "A4", "LETTER", "A3", etc. or { width, height }
  orientation?: "portrait" | "landscape";
  margins?: Partial<PdfMargins>; // { top, right, bottom, left } in points (72pt = 1in)
  fitToPage?: boolean; // Scale columns to fit page width (default: true)
  scale?: number; // Additional scale factor (default: 1.0)

  // Content
  showGridLines?: boolean; // Render cell grid lines
  gridLineColor?: string; // ARGB color for grid lines (e.g. "FF3366CC")
  repeatRows?: number | false; // Number of header rows to repeat on each page
  sheets?: (string | number)[]; // Select specific sheets by name or 1-based index

  // Headers & footers
  showSheetNames?: boolean; // Show sheet name at top of each page
  showPageNumbers?: boolean; // Show "Page X of Y" at bottom of each page

  // Metadata
  title?: string;
  author?: string;
  subject?: string;
  creator?: string; // PDF producer string (default: "excelts")

  // Font
  font?: Uint8Array; // TrueType font file bytes (for Unicode/CJK)
  defaultFontFamily?: string; // Fallback font family (default: "Helvetica")
  defaultFontSize?: number; // Fallback font size (default: 11)

  // Encryption (AES-256, PDF 2.0)
  encryption?: {
    ownerPassword: string; // Owner password (required)
    userPassword?: string; // User open password (optional)
    permissions?: {
      print?: boolean; // Allow printing
      modify?: boolean; // Allow modification
      copy?: boolean; // Allow copy/paste
      annotate?: boolean; // Allow annotations
      fillForms?: boolean; // Allow form filling
      accessibility?: boolean; // Allow accessibility extraction
      assemble?: boolean; // Allow document assembly
      printHighQuality?: boolean; // Allow high-quality printing
    };
  };
}
```

## 读取选项

```typescript
interface ReadPdfOptions {
  password?: string; // Password for encrypted PDFs (user or owner)
  pages?: number[]; // Which pages to extract (1-based). Omit for all pages
  extractText?: boolean; // Extract text (default: true)
  extractImages?: boolean; // Extract images (default: true)
  extractMetadata?: boolean; // Extract metadata (default: true)
  extractAnnotations?: boolean; // Extract annotations (default: true)
  extractFormFields?: boolean; // Extract form fields (default: true)
}
```

### 读取结果

```typescript
interface ReadPdfResult {
  text: string; // All text from all pages
  pages: ReadPdfPage[]; // Per-page results
  metadata: PdfMetadata; // Document metadata
  formFields: PdfFormField[]; // Form fields (document-level)
}

interface ReadPdfPage {
  pageNumber: number; // 1-based
  text: string; // Page text
  textLines: TextLine[]; // Structured lines with positions
  textFragments: TextFragment[]; // Raw fragments with exact coordinates
  images: ExtractedImage[]; // Extracted images
  annotations: PdfAnnotation[]; // Annotations (links, comments, highlights)
  width: number; // Page width in points
  height: number; // Page height in points
  warnings: string[]; // Non-fatal extraction warnings
}

interface PdfAnnotation {
  subtype: string; // "Link", "Text", "Highlight", "FreeText", "Stamp", etc.
  rect: PdfRect; // Bounding rectangle { x1, y1, x2, y2 }
  contents: string; // Text content
  author: string; // Author / title
  uri: string; // For Link: destination URI
  destination: string; // For Link: named destination
  color: number[]; // Color array [r, g, b] in [0,1]
  flags: number; // Annotation flags
}

interface PdfFormField {
  name: string; // Fully qualified name (e.g. "form.address.city")
  type: PdfFormFieldType; // "text" | "checkbox" | "radio" | "dropdown" | "listbox" | "button" | "signature"
  value: string; // Current value
  defaultValue: string; // Default value
  readOnly: boolean; // Read-only flag
  required: boolean; // Required flag
  options: string[]; // For choice fields: available options
  exportValue: string; // For checkboxes: export value when checked
}

interface PdfMetadata {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  creator?: string;
  producer?: string;
  creationDate?: Date;
  modificationDate?: Date;
  pdfVersion: string;
  pageCount: number;
  encrypted: boolean;
}
```

### 页面尺寸

通过 `PageSizes` 访问内置页面尺寸：

| 名称        | 尺寸 (pt)        | 毫米      |
| ----------- | ---------------- | --------- |
| `"LETTER"`  | 612 x 792        | 216 x 279 |
| `"LEGAL"`   | 612 x 1008       | 216 x 356 |
| `"TABLOID"` | 792 x 1224       | 279 x 432 |
| `"A3"`      | 841.89 x 1190.55 | 297 x 420 |
| `"A4"`      | 595.28 x 841.89  | 210 x 297 |
| `"A5"`      | 419.53 x 595.28  | 148 x 210 |

自定义尺寸：`{ width: 396, height: 612 }`（单位为点，72pt = 1 英寸）。

---

## 样式支持

PDF 写入器渲染所有标准单元格样式：

### 文本

- 字体族、大小、粗体、斜体
- 字体颜色（ARGB、主题颜色及色调）
- 下划线和删除线
- 富文本（单元格内混合格式）
- 通过 `numFmt` 进行数字/日期/货币格式化
- 超链接（可点击标注）

### 对齐

- 水平：左对齐、居中、右对齐
- 垂直：顶部、居中、底部
- 文本自动换行
- 文本缩进
- 文本旋转（倾斜和垂直堆叠）

### 单元格

- 背景填充（纯色，支持 alpha 透明度）
- 边框：细线、中等、粗线、虚线、点线（支持颜色）
- 合并单元格（水平和垂直跨越）

---

## 分页

### 自动分页

- 行超出页面高度：自动垂直分页
- 列超出页面宽度：自动水平分页
- `fitToPage: true`（默认）：缩放所有列以适应页面宽度

### 重复标题行

```typescript
await excelToPdf(workbook, { repeatRows: 2 }); // Repeat first 2 rows on every page
```

或通过工作表页面设置：

```typescript
worksheet.pageSetup.printTitlesRow = "1:2"; // Repeat rows 1-2
```

### 手动分页符

```typescript
worksheet.getRow(20).addPageBreak(); // Break after row 20
```

### 打印区域

```typescript
worksheet.pageSetup.printArea = "A1:F50"; // Export only this range
```

> **注意：** 如果设置了多区域打印范围（例如 `"A1:B5&&D1:E10"`），PDF 导出仅使用第一个区域。

---

## 图片

工作表包含图片时，JPEG 和 PNG 图片会被嵌入：

```typescript
const imageId = workbook.addImage({
  buffer: jpegBytes,
  extension: "jpeg"
});

worksheet.addImage(imageId, {
  tl: { col: 0, row: 0 },
  ext: { width: 200, height: 150 }
});

const pdf = await excelToPdf(workbook);
// Image appears in the PDF at the specified position
```

PNG 透明度（RGBA 和 tRNS）通过 PDF 软遮罩保留。

---

## 加密

写入器生成 **AES-256 加密 PDF**（PDF 2.0, V=5, R=5）。读取器可解密**所有主流加密格式**，包括旧版 RC4。

### 写入加密（AES-256）

#### 仅所有者密码（无打开密码）

```typescript
const pdf = await excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    permissions: { print: true, copy: false, modify: false }
  }
});
// Opens without password, but copy/modify is restricted
```

#### 需要打开密码

```typescript
const pdf = await excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "reader"
  }
});
// Requires "reader" to open
```

### 读取解密（所有格式）

读取器自动检测并解密：

| 格式    | 版本       | 支持 |
| ------- | ---------- | ---- |
| RC4-40  | V=1, R=2   | 读取 |
| RC4-128 | V=2, R=3   | 读取 |
| AES-128 | V=4, R=4   | 读取 |
| AES-256 | V=5, R=5/6 | 读取 |

```typescript
// Automatically detects encryption type
const result = await readPdf(encryptedBytes, { password: "secret" });
```

---

## Unicode / CJK 支持

标准 Type1 字体（Helvetica、Times、Courier）仅支持拉丁字符。对于 Unicode、CJK 或其他文字系统，需提供 TrueType 字体：

```typescript
import { readFileSync } from "fs";

const pdf = await excelToPdf(workbook, {
  font: readFileSync("NotoSansSC-Regular.ttf")
});
```

字体会自动子集化（仅嵌入使用到的字形）以最小化 PDF 文件大小。

---

## 每工作表页面设置

使用 Excel 桥接时，每个工作表的 `pageSetup` 都会被采用：

```typescript
const ws1 = workbook.addWorksheet("Summary");
ws1.pageSetup.paperSize = 9; // A4
ws1.pageSetup.orientation = "portrait";

const ws2 = workbook.addWorksheet("Data");
ws2.pageSetup.paperSize = 1; // Letter
ws2.pageSetup.orientation = "landscape";

// Each sheet renders with its own page size/orientation
const pdf = await excelToPdf(workbook);
```

工作表边距同样会被继承：

```typescript
ws.pageSetup.margins = {
  left: 0.5, // inches
  right: 0.5,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3
};
```

---

## Tree-Shaking

PDF 模块完全支持 tree-shaking。如果你不导入任何 PDF 导出项，该模块对你的打包体积**零影响**：

```typescript
// Only imports Excel core — PDF module is NOT included
import { Workbook } from "@cj-tech-master/excelts";

// Imports Excel + PDF bridge
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";
```

---

## 示例

可运行示例位于 `src/modules/pdf/examples/`：

| 文件                   | 演示内容                                             |
| ---------------------- | ---------------------------------------------------- |
| `pdf-basic.ts`         | 页面大小、边距、元数据、工作表选择、缩放             |
| `pdf-styled.ts`        | 字体、填充、边框、对齐、合并、旋转、富文本、数字格式 |
| `pdf-advanced.ts`      | 分页、分页符、加密、透明度、书签、隐藏行列           |
| `pdf-excel-to-pdf.ts`  | 读取 `.xlsx` 文件并转换为 PDF                        |
| `pdf-images.ts`        | 图片嵌入（JPEG、PNG 透明度）                         |
| `pdf-reader.ts`        | 文本提取、元数据、图片、加密 PDF、选择性提取         |
| `pdf-reader-stress.ts` | 大规模压力测试：数千单元格、加密往返、基准测试       |

运行任意示例：

```bash
npx tsx src/modules/pdf/examples/pdf-basic.ts
# Output: tmp/pdf-examples/*.pdf

npx tsx src/modules/pdf/examples/pdf-reader.ts
# Output: tmp/pdf-reader-examples/
```

---

## API 参考

### `readPdf(data, options?)`

读取 PDF 文件并提取文本、图片和元数据。返回 `Promise<ReadPdfResult>`。

```typescript
import { readPdf } from "@cj-tech-master/excelts/pdf";

// Basic
const result = await readPdf(pdfBytes);
console.log(result.text);
console.log(result.pages[0].images);
console.log(result.pages[0].annotations);
console.log(result.formFields);
console.log(result.metadata);

// Encrypted
const result = await readPdf(pdfBytes, { password: "secret" });

// Selective
const result = await readPdf(pdfBytes, {
  pages: [1, 3],
  extractImages: false,
  extractMetadata: false
});
```

### `pdf(input, options?)`

从纯数据生成 PDF。返回 `Promise<Uint8Array>`。

```typescript
// 2D array
await pdf([["Name", "Age"], ["Alice", 30]]);

// Single sheet with column widths
await pdf({ name: "Report", columns: [{ width: 25 }, 15], data: [["A", "B"]] });

// Multiple sheets
await pdf({ sheets: [{ name: "S1", data: [...] }, { name: "S2", data: [...] }] });

// With options
await pdf([["A", 1]], { showGridLines: true, pageSize: "A4" });
```

### `excelToPdf(workbook, options?)`

将 Excel `Workbook` 转换为 PDF。返回 `Promise<Uint8Array>`。

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
// ... build workbook ...
const bytes = await excelToPdf(workbook, { showGridLines: true });
```

### 错误类型

```typescript
import {
  PdfError, // Base class for all PDF errors
  PdfRenderError, // Layout/rendering failures
  PdfFontError, // Font parsing/embedding failures
  PdfStructureError, // PDF structure assembly failures
  isPdfError // Type guard: (err: unknown) => err is PdfError
} from "@cj-tech-master/excelts";
```

所有错误继承自 `BaseError`，支持 `cause` 链式追踪：

```typescript
try {
  await excelToPdf(workbook);
} catch (err) {
  if (isPdfError(err)) {
    console.error(err.message, err.cause);
  }
}
```
