# Excel Module

[中文](README_zh.md)

Modern TypeScript Excel Workbook Manager — read, manipulate, and write XLSX and JSON spreadsheets with zero runtime dependencies.

## Features

- **Create, read, and modify XLSX files** — full Open XML support
- **Multiple worksheet support** — add, remove, reorder, copy
- **Cell styling** — fonts, colors, borders, fills, alignment, number formats
- **Cell merging and formatting** — merge ranges, rich text, hyperlinks
- **Row and column properties** — width, height, hidden, outline level, auto-fit
- **Freeze panes and split views** — freeze rows/columns, split at position
- **Rich text support** — multiple fonts/styles within a single cell
- **Formulas and calculated values** — shared formulas, defined names
- **Data validation** — list, whole, decimal, date, textLength, custom
- **Conditional formatting** — cell value, color scale, data bar, icon set
- **Images** — JPEG, PNG, GIF with one-cell and two-cell anchors
- **Hyperlinks** — internal, external, email
- **Pivot tables** — read and preserve pivot table definitions
- **Tables** — auto-filters, totals row, structured references
- **Comments and notes** — threaded comments, legacy notes
- **Checkboxes** — form controls and cell-level checkboxes
- **Page setup** — print area, print titles, header/footer, page breaks
- **Data protection** — sheet protection with password (SHA-512)
- **Streaming** — `WorkbookReader` and `WorkbookWriter` for large files
- **CSV import/export** — `readCsv`, `writeCsv`, `readCsvFile`, `writeCsvFile`
- **Markdown import/export** — `readMarkdown`, `writeMarkdown`, `readMarkdownFile`, `writeMarkdownFile`
- **PDF export** — `excelToPdf()` with full styling, pagination, fonts, encryption
- **Browser support** — `xlsx.load()`, `xlsx.writeBuffer()`, no polyfills needed

## Quick Start

### Creating a Workbook

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("My Sheet");

// Add data
sheet.addRow(["Name", "Age", "Email"]);
sheet.addRow(["John Doe", 30, "john@example.com"]);
sheet.addRow(["Jane Smith", 25, "jane@example.com"]);

// Node.js: write to file
await workbook.xlsx.writeFile("output.xlsx");

// Browser: write to buffer
const buffer = await workbook.xlsx.writeBuffer();
```

### Reading a Workbook

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// Node.js: read from file
await workbook.xlsx.readFile("input.xlsx");

// Browser: read from ArrayBuffer
await workbook.xlsx.load(arrayBuffer);

const worksheet = workbook.getWorksheet(1);
worksheet.eachRow((row, rowNumber) => {
  console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
});
```

### Styling Cells

```typescript
const cell = worksheet.getCell("A1");
cell.value = "Hello";
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

### Number Formats

```typescript
// Currency
cell.numFmt = "$#,##0.00";

// Percentage
cell.numFmt = "0.00%";

// Date
cell.numFmt = "yyyy-mm-dd";

// Custom
cell.numFmt = '#,##0.00 "units"';
```

### Rich Text

```typescript
cell.value = {
  richText: [
    { text: "Bold ", font: { bold: true } },
    { text: "and ", font: {} },
    { text: "Red", font: { color: { argb: "FFFF0000" } } }
  ]
};
```

### Formulas

```typescript
cell.value = { formula: "SUM(A1:A10)" };
cell.value = { formula: "A1+B1", result: 42 }; // with cached result

// Shared formulas
sheet.getCell("A1").value = { formula: "B1*2", shareType: "shared", ref: "A1:A10" };

// Defined names
workbook.definedNames.add("MyRange", "Sheet1!$A$1:$B$10");
```

### Data Validation

```typescript
worksheet.getCell("A1").dataValidation = {
  type: "list",
  allowBlank: true,
  formulae: ['"Option1,Option2,Option3"']
};

worksheet.getCell("B1").dataValidation = {
  type: "whole",
  operator: "between",
  formulae: [1, 100],
  showErrorMessage: true,
  errorTitle: "Invalid",
  error: "Enter a number between 1 and 100"
};
```

### Conditional Formatting

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

### Images

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

### Tables

```typescript
worksheet.addTable({
  name: "SalesTable",
  ref: "A1",
  headerRow: true,
  totalsRow: true,
  columns: [
    { name: "Product", totalsRowLabel: "Total", filterButton: true },
    { name: "Revenue", totalsRowFunction: "sum", filterButton: true }
  ],
  rows: [
    ["Widget", 1000],
    ["Gadget", 2500]
  ]
});
```

### Merge Cells

```typescript
worksheet.mergeCells("A1:D1");
worksheet.getCell("A1").value = "Merged Header";
worksheet.getCell("A1").alignment = { horizontal: "center" };
```

### Freeze Panes

```typescript
// Freeze first row
worksheet.views = [{ state: "frozen", ySplit: 1 }];

// Freeze first column
worksheet.views = [{ state: "frozen", xSplit: 1 }];

// Freeze both
worksheet.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
```

### Page Setup

```typescript
worksheet.pageSetup = {
  paperSize: 9, // A4
  orientation: "landscape",
  fitToPage: true,
  fitToWidth: 1,
  fitToHeight: 0,
  margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 }
};

// Print area
worksheet.pageSetup.printArea = "A1:G20";

// Print titles (repeat rows 1-2 on every page)
worksheet.pageSetup.printTitlesRow = "1:2";
```

### Sheet Protection

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

### Comments

```typescript
worksheet.getCell("A1").note = "Simple comment";

worksheet.getCell("B1").note = {
  texts: [{ text: "Author: ", font: { bold: true } }, { text: "This is a rich text comment" }]
};
```

### Auto-Fit Column Width

```typescript
worksheet.columns.forEach(column => {
  column.width = column.values
    ? Math.max(...column.values.map(v => String(v ?? "").length)) + 2
    : 10;
});
```

## PDF Export

Export any workbook to PDF with zero external dependencies:

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Report");
sheet.columns = [
  { header: "Product", key: "product", width: 20 },
  { header: "Revenue", key: "revenue", width: 15 }
];
sheet.addRow({ product: "Widget", revenue: 1000 });
sheet.getColumn("revenue").numFmt = "$#,##0.00";

const pdf = excelToPdf(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "Sales Report"
});

// Node.js
import { writeFileSync } from "fs";
writeFileSync("report.pdf", pdf);

// Browser
const blob = new Blob([pdf], { type: "application/pdf" });
window.open(URL.createObjectURL(blob));
```

### XLSX to PDF Conversion

```typescript
const workbook = new Workbook();
await workbook.xlsx.readFile("input.xlsx");
const pdf = excelToPdf(workbook);
```

### PDF Encryption

```typescript
const pdf = excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "reader",
    permissions: { print: true, copy: false }
  }
});
```

### Unicode / CJK Font Embedding

```typescript
import { readFileSync } from "fs";
const pdf = excelToPdf(workbook, {
  font: readFileSync("NotoSansSC-Regular.ttf")
});
```

## CSV Import/Export

```typescript
import { Workbook } from "@cj-tech-master/excelts";
import fs from "fs";

const workbook = new Workbook();

// Node.js: read/write CSV files
await workbook.readCsvFile("data.csv");
await workbook.writeCsvFile("output.csv");

// Read CSV from stream
await workbook.readCsv(fs.createReadStream("data.csv"), { sheetName: "Imported" });

// Write CSV to stream
await workbook.writeCsv(fs.createWriteStream("output.csv"));

// Write CSV to string / bytes
const csvText = workbook.writeCsv();
const bytes = await workbook.writeCsvBuffer();

// Browser: read from string/ArrayBuffer/File
await workbook.readCsv(csvString);
await workbook.readCsv(arrayBuffer);
```

## Markdown Import/Export

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// Read Markdown table
workbook.readMarkdown("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
await workbook.readMarkdownFile("table.md");

// Write Markdown
const mdText = workbook.writeMarkdown();
await workbook.writeMarkdownFile("output.md");
const bytes = workbook.writeMarkdownBuffer();
```

## Streaming API

### Streaming Reader

Read large XLSX files with minimal memory usage:

```typescript
import { WorkbookReader } from "@cj-tech-master/excelts";

const reader = new WorkbookReader("large-file.xlsx", {
  worksheets: "emit",
  sharedStrings: "cache",
  hyperlinks: "ignore",
  styles: "ignore"
});

for await (const worksheet of reader) {
  console.log(`Reading: ${worksheet.name}`);
  for await (const row of worksheet) {
    console.log(row.values);
  }
}
```

### Streaming Writer

Write large XLSX files row by row:

```typescript
import { WorkbookWriter } from "@cj-tech-master/excelts";

const workbook = new WorkbookWriter({
  filename: "output.xlsx",
  useSharedStrings: true,
  useStyles: true
});

const sheet = workbook.addWorksheet("Data");
for (let i = 0; i < 1000000; i++) {
  sheet.addRow([`Row ${i}`, i, new Date()]).commit();
}

sheet.commit();
await workbook.commit();
```

### Web Streams (Node.js 22+ and Browsers)

```typescript
import { WorkbookWriter, WorkbookReader } from "@cj-tech-master/excelts";

// Write to Web WritableStream
const chunks: Uint8Array[] = [];
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

// Read from Web ReadableStream
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

## Browser Support

### Using with Bundlers (Vite, Webpack, Rollup, esbuild)

```typescript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");
sheet.getCell("A1").value = "Hello, Browser!";

const buffer = await workbook.xlsx.writeBuffer();
const blob = new Blob([buffer], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});
const url = URL.createObjectURL(blob);
```

### Using with Script Tags

```html
<script src="https://unpkg.com/@cj-tech-master/excelts/dist/iife/excelts.iife.min.js"></script>
<script>
  const { Workbook } = ExcelTS;
  const wb = new Workbook();
</script>
```

### Browser Notes

- Use `xlsx.load(arrayBuffer)` instead of `xlsx.readFile()`
- Use `xlsx.writeBuffer()` instead of `xlsx.writeFile()`
- PDF export is fully supported
- CSV and Markdown operations are supported
- Sheet protection with passwords uses pure JS SHA-512

## Utility Exports

```typescript
import {
  // Date conversion
  dateToExcel,
  excelToDate,
  DateParser,
  DateFormatter,

  // Binary utilities
  base64ToUint8Array,
  uint8ArrayToBase64,
  concatUint8Arrays,
  toUint8Array,
  stringToUint8Array,
  uint8ArrayToString,

  // XML utilities
  xmlEncode,
  xmlDecode,
  xmlEncodeAttr,
  validateXmlName,

  // PDF export
  pdf,
  excelToPdf,
  PageSizes,
  PdfError,
  isPdfError,

  // Errors
  BaseError,
  ExcelError,
  toError,
  errorToJSON,
  getErrorChain,
  getRootCause
} from "@cj-tech-master/excelts";
```

## Examples

See the [examples directory](examples/) for runnable code covering all features:

- Workbook creation, reading, and copying
- Cell styling, fonts, borders, fills
- Formulas, data validation, conditional formatting
- Images (JPEG, PNG), hyperlinks, comments
- Tables with auto-filters and totals
- Merge cells, freeze panes, page setup
- Streaming reader and writer
- Web Streams integration
- PDF export
- And more...
