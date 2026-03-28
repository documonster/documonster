# ExcelTS

[![Build Status](https://github.com/cjnoname/excelts/actions/workflows/ci.yml/badge.svg?branch=main&event=push)](https://github.com/cjnoname/excelts/actions/workflows/ci.yml)

Modern TypeScript Excel Workbook Manager - Read, manipulate and write spreadsheet data and styles to XLSX and JSON.

## About This Project

ExcelTS is a modern TypeScript Excel workbook manager with:

- 🚀 **Zero Runtime Dependencies** - Pure TypeScript implementation with no external packages
- ✅ **Broad Runtime Support** - LTS Node.js, Bun, and modern browsers (Chrome, Firefox, Safari, Edge)
- ✅ **Full TypeScript Support** - Complete type definitions and modern TypeScript patterns
- ✅ **Modern Build System** - Using Rolldown for faster builds
- ✅ **Enhanced Testing** - Migrated to Vitest with browser testing support
- ✅ **ESM First** - Native ES Module support with CommonJS compatibility
- ✅ **Named Exports** - All exports are named for better tree-shaking

## Translations

- [中文文档](README_zh.md)

## Installation

```bash
npm install @cj-tech-master/excelts
```

## Quick Start

### Creating a Workbook

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("My Sheet");

// Add data
sheet.addRow(["Name", "Age", "Email"]);
sheet.addRow(["John Doe", 30, "john@example.com"]);
sheet.addRow(["Jane Smith", 25, "jane@example.com"]);

// Save to file
// Node.js only: write to a file path
await workbook.xlsx.writeFile("output.xlsx");

// Browser: use `writeBuffer()` and save as a Blob (see the Browser Support section)
```

### Reading a Workbook

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
// Node.js only: read from a file path
await workbook.xlsx.readFile("input.xlsx");

// Browser: use `xlsx.load(arrayBuffer)` (see the Browser Support section)

const worksheet = workbook.getWorksheet(1);
worksheet.eachRow((row, rowNumber) => {
  console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
});
```

### Styling Cells

```javascript
// Set cell value and style
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
```

## Features

- **Excel Operations**
  - Create, read, and modify XLSX files
  - Multiple worksheet support
  - Cell styling (fonts, colors, borders, fills)
  - Cell merging and formatting
  - Row and column properties
  - Freeze panes and split views

- **Data Handling**
  - Rich text support
  - Formulas and calculated values
  - Data validation
  - Conditional formatting
  - Images and charts
  - Hyperlinks
  - Pivot tables

- **PDF Export**
  - Zero-dependency Excel-to-PDF conversion
  - Full cell styling (fonts, colors, borders, fills, alignment)
  - Automatic pagination with repeat header rows
  - TrueType font embedding for Unicode/CJK text
  - JPEG and PNG image embedding with transparency
  - Password protection and encryption
  - Per-worksheet page setup (size, orientation, margins)
  - Tree-shakeable (not imported = not bundled)

- **Advanced Features**
  - Streaming for large files
  - CSV import/export
  - Tables with auto-filters
  - Page setup and printing options
  - Data protection
  - Comments and notes

## Subpath Exports

ExcelTS provides focused subpath exports for standalone module usage:

```typescript
// Main entry - Excel core (Workbook, Worksheet, Cell, etc.)
import { Workbook, WorkbookWriter } from "@cj-tech-master/excelts";

// ZIP/TAR archive utilities
import { zip, unzip, ZipArchive, compress } from "@cj-tech-master/excelts/zip";

// CSV parsing, formatting, and streaming
import { parseCsv, formatCsv, CsvParserStream } from "@cj-tech-master/excelts/csv";

// Cross-platform stream primitives
import { Readable, pipeline, createTransform } from "@cj-tech-master/excelts/stream";
```

Each subpath supports `browser`, `import` (ESM), and `require` (CJS) conditions. See the module READMEs for details:

- [PDF Module](src/modules/pdf/README.md) - Zero-dependency Excel-to-PDF export with encryption and font embedding
- [CSV Module](src/modules/csv/README.md) - RFC 4180 parser/formatter, streaming, data generation
- [Archive Module](src/modules/archive/README.md) - ZIP/TAR create/read/edit, compression, encryption
- [Stream Module](src/modules/stream/README.md) - Cross-platform Readable/Writable/Transform/Duplex

## PDF Export

Export any workbook to PDF with zero external dependencies:

```javascript
import { Workbook, exportPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Report");
sheet.columns = [
  { header: "Product", key: "product", width: 20 },
  { header: "Revenue", key: "revenue", width: 15 }
];
sheet.addRow({ product: "Widget", revenue: 1000 });
sheet.getColumn("revenue").numFmt = "$#,##0.00";

// One-line export
const pdf = exportPdf(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "Sales Report"
});

// Node.js: write to file
import { writeFileSync } from "fs";
writeFileSync("report.pdf", pdf);

// Browser: download
const blob = new Blob([pdf], { type: "application/pdf" });
const url = URL.createObjectURL(blob);
window.open(url);
```

### Convert Existing XLSX to PDF

```javascript
const workbook = new Workbook();
await workbook.xlsx.readFile("input.xlsx");
const pdf = exportPdf(workbook);
```

### Encryption

```javascript
const pdf = exportPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "reader",
    permissions: { print: true, copy: false }
  }
});
```

### Unicode / CJK

```javascript
import { readFileSync } from "fs";

const pdf = exportPdf(workbook, {
  font: readFileSync("NotoSansSC-Regular.ttf") // TrueType font for CJK text
});
```

For the full API reference and all options, see the [PDF Module documentation](src/modules/pdf/README.md).

## Archive Utilities (ZIP/TAR)

ExcelTS includes internal ZIP/TAR utilities used by the XLSX pipeline. If you use the
archive APIs directly, ZIP string encoding can be customized via `ZipStringEncoding`:

- Default: `"utf-8"`
- Legacy: `"cp437"`
- Custom: provide a codec with `encode`/`decode` plus optional flags

When a non-UTF-8 encoding is used, Unicode extra fields can be emitted for better
cross-tool compatibility.

### Editing an existing ZIP (ZipEditor)

ExcelTS also includes a ZIP editor that can apply filesystem-like edits to an existing archive
and then output a new ZIP.

- Supports `set()`, `delete()`, `rename()`, `deleteDirectory()`, `setComment()`
- Unchanged entries are passed through efficiently when possible

```js
import { editZip } from "@cj-tech-master/excelts";

const editor = await editZip(existingZipBytes, {
  reproducible: true,

  // Passthrough behavior for unchanged entries:
  // - "strict" (default): raw passthrough must be available or it throws
  // - "best-effort": if raw passthrough is unavailable, fall back to extract+re-add
  preserve: "best-effort",
  onWarning: w => console.warn(w.code, w.entry, w.message)
});

editor.delete("old.txt");
editor.rename("a.txt", "renamed.txt");
editor.set("new.txt", "hello");

const out = await editor.bytes();
```

## Streaming API

For processing large Excel files without loading them entirely into memory, ExcelTS provides streaming reader and writer APIs.

- **Node.js**: `WorkbookReader` supports reading from a file path, and `WorkbookWriter` supports writing to a filename.
- **Browsers**: use `Uint8Array` / `ArrayBuffer` / Web `ReadableStream<Uint8Array>` for reading, and Web `WritableStream<Uint8Array>` for writing.
- Note: ExcelTS does not re-export the internal stream utility surface (e.g. `Readable`, `Writable`). Prefer standard Web Streams (browser/Node 22+) or Node.js streams.

### Streaming Reader

Read large XLSX files with minimal memory usage:

```javascript
import { WorkbookReader } from "@cj-tech-master/excelts";

// Node.js: read from file path
const reader = new WorkbookReader("large-file.xlsx", {
  worksheets: "emit", // emit worksheet events
  sharedStrings: "cache", // cache shared strings for cell values
  hyperlinks: "ignore", // ignore hyperlinks
  styles: "ignore" // ignore styles for faster parsing
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

```javascript
import { WorkbookWriter } from "@cj-tech-master/excelts";

// Node.js: write to filename
const workbook = new WorkbookWriter({
  filename: "output.xlsx",
  useSharedStrings: true,
  useStyles: true
});

const sheet = workbook.addWorksheet("Data");

// Write rows one at a time
for (let i = 0; i < 1000000; i++) {
  sheet.addRow([`Row ${i}`, i, new Date()]).commit();
}

// Commit worksheet and finalize
sheet.commit();
await workbook.commit();
```

### Web Streams (Node.js 22+ and Browsers)

`WorkbookWriter` can write to a Web `WritableStream<Uint8Array>`, and `WorkbookReader` can read from a Web `ReadableStream<Uint8Array>`.

This does **not** require importing any extra stream utility surface from ExcelTS; it uses the standard Web Streams API.

- Full runnable example: [src/modules/excel/examples/web-streams-reader-writer.ts](src/modules/excel/examples/web-streams-reader-writer.ts)

Run locally (Node.js 22+):

```bash
npx tsx src/modules/excel/examples/web-streams-reader-writer.ts
```

Minimal end-to-end snippet:

```javascript
import { WorkbookWriter, WorkbookReader } from "@cj-tech-master/excelts";

// 1) Write workbook -> Web WritableStream
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

// 2) Read workbook <- Web ReadableStream
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

## CSV Support

### Node.js (Full Streaming Support)

```javascript
import { Workbook } from "@cj-tech-master/excelts";
import fs from "fs";

const workbook = new Workbook();

// Read CSV from file
await workbook.readCsvFile("data.csv");

// Read CSV from stream
const stream = fs.createReadStream("data.csv");
await workbook.readCsv(stream, { sheetName: "Imported" });

// Write CSV to file
await workbook.writeCsvFile("output.csv");

// Write CSV to stream
const writeStream = fs.createWriteStream("output.csv");
await workbook.writeCsv(writeStream);

// Write CSV to string / bytes
const csvText = workbook.writeCsv();
const bytes = await workbook.writeCsvBuffer();
```

### Browser (In-Memory)

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// Read CSV from string
await workbook.readCsv(csvString);

// Read CSV from ArrayBuffer (e.g., from fetch)
const response = await fetch("data.csv");
const arrayBuffer = await response.arrayBuffer();
await workbook.readCsv(arrayBuffer);

// Read CSV from File (e.g., <input type="file">)
await workbook.readCsv(file);

// Write CSV to string
const csvOutput = workbook.writeCsv();

// Write CSV to Uint8Array bytes
const bytes = await workbook.writeCsvBuffer();
```

## Browser Support

ExcelTS has native browser support with **zero configuration** required for modern bundlers.

### Using with Bundlers (Vite, Webpack, Rollup, esbuild)

Simply import ExcelTS - no polyfills or configuration needed:

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");
sheet.getCell("A1").value = "Hello, Browser!";

// Write to buffer and download
const buffer = await workbook.xlsx.writeBuffer();
const blob = new Blob([buffer], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});
const url = URL.createObjectURL(blob);
// ... trigger download
```

### Using with Script Tags (No Bundler)

```html
<script src="https://unpkg.com/@cj-tech-master/excelts/dist/iife/excelts.iife.min.js"></script>
<script>
  const { Workbook } = ExcelTS;
  const wb = new Workbook();
  // ... use workbook API
</script>
```

### Manual Browser Example (Local)

For a quick manual smoke test in a real browser (create/download/read XLSX, worksheet protection, etc.), use:

- [src/modules/excel/examples/browser-smoke.html](src/modules/excel/examples/browser-smoke.html)

Steps:

```bash
npm run build:browser:bundle
npx serve .
```

Then open `http://localhost:3000/src/modules/excel/examples/browser-smoke.html`.

### Browser-Specific Notes

- **PDF export is fully supported** in browsers with zero configuration
- **CSV operations are supported** using native RFC 4180 implementation
  - Use `await workbook.readCsv(input)` to read CSV
  - Use `workbook.writeCsv()` or `await workbook.writeCsvBuffer()` to write CSV
- Use `xlsx.load(arrayBuffer)` instead of `xlsx.readFile()`
- Use `xlsx.writeBuffer()` instead of `xlsx.writeFile()`
- Worksheet protection with passwords is fully supported (pure JS SHA-512)

## Utility Exports

The main entry also exports commonly useful utilities:

```typescript
import {
  // Excel date conversion
  dateToExcel, // JS Date -> Excel serial number
  excelToDate, // Excel serial number -> JS Date

  // Date parsing/formatting (high-performance, zero-dep)
  DateParser, // Batch date parser with format auto-detection
  DateFormatter, // Batch date formatter

  // Binary utilities (cross-platform)
  base64ToUint8Array,
  uint8ArrayToBase64,
  concatUint8Arrays,
  toUint8Array,
  stringToUint8Array,
  uint8ArrayToString,

  // XML utilities
  xmlEncode,
  xmlDecode,

  // PDF export
  exportPdf, // Workbook -> Uint8Array (PDF)
  PdfExporter, // Class-based PDF export
  PageSizes, // Built-in page size definitions
  PdfError, // Base PDF error
  PdfRenderError, // Layout/rendering failures
  PdfFontError, // Font parsing/embedding failures
  PdfStructureError, // PDF structure assembly failures
  isPdfError, // Type guard for PDF errors

  // Error infrastructure
  BaseError, // Base class for all library errors
  ExcelError, // Base Excel error (instanceof checks)
  toError, // Normalize unknown -> Error
  errorToJSON, // Serialize error (with cause chain)
  getErrorChain, // Get full error cause chain as array
  getRootCause // Get deepest error in cause chain
} from "@cj-tech-master/excelts";
```

## Requirements

### Node.js

- **Node.js >= 22.0.0** (ES2020 native support)

### Browsers (No Polyfills Required)

- **Chrome >= 89** (March 2021)
- **Edge >= 89** (March 2021)
- **Firefox >= 102** (June 2022)
- **Safari >= 14.1** (April 2021)
- **Opera >= 75** (March 2021)

For older browsers without native `CompressionStream` API (Firefox < 113, Safari < 16.4), ExcelTS automatically uses a built-in pure JavaScript DEFLATE implementation - no configuration or polyfills needed.

ExcelTS does **not** require `crypto.randomUUID()` in browsers; it uses an internal UUID v4 generator with a `crypto.getRandomValues()` fallback.

## Maintainer

This project is actively maintained by [CJ (@cjnoname)](https://github.com/cjnoname).

### Maintenance Status

**Active Maintenance** - This project is actively maintained with a focus on:

- 🔒 **Security Updates** - Timely security patches and dependency updates
- 🐛 **Bug Fixes** - Critical bug fixes and stability improvements
- 📦 **Dependency Management** - Keeping dependencies up-to-date and secure
- 🔍 **Code Review** - Reviewing and merging community contributions

### Contributing

While I may not have the bandwidth to develop new features regularly, **community contributions are highly valued and encouraged!**

- 💡 **Pull Requests Welcome** - I will review and merge quality PRs promptly
- 🚀 **Feature Proposals** - Open an issue to discuss new features before implementing
- 🐛 **Bug Reports** - Please report bugs with reproducible examples
- 📖 **Documentation** - Improvements to documentation are always appreciated

## API Documentation

For detailed API documentation, please refer to the comprehensive documentation sections:

- Workbook Management
- Worksheets
- Cells and Values
- Styling
- Formulas
- Data Validation
- Conditional Formatting
- File I/O
- [PDF Export](src/modules/pdf/README.md)

## Contributing Guidelines

Contributions are welcome! Please feel free to submit a Pull Request.

### Before Submitting a PR

1. **Bug Fixes**: Add a unit-test or integration-test (in `src/**/__tests__`) that reproduces the issue
2. **New Features**: Open an issue first to discuss the feature and implementation approach
3. **Documentation**: Update relevant documentation and type definitions
4. **Code Style**: Follow the existing code style and pass all linters (`npm run lint`)
5. **Tests**: Ensure all tests pass (`npm test`) and add tests for new functionality

### Important Notes

- **Version Numbers**: Please do not modify package version in PRs. Versions are managed through releases.
- **License**: All contributions will be included under the project's MIT license
- **Commit Messages**: Write clear, descriptive commit messages

### Getting Help

If you need help or have questions:

- 📖 Check existing [issues](https://github.com/cjnoname/excelts/issues) and [documentation](https://github.com/cjnoname/excelts)
- 💬 Open a [new issue](https://github.com/cjnoname/excelts/issues/new) for discussion
- 🐛 Use issue templates for bug reports

## License

MIT License

See LICENSE.

Third-party software notices and attributions are provided in THIRD_PARTY_NOTICES.md.

## Links

- [GitHub Repository](https://github.com/cjnoname/excelts)
- [Issue Tracker](https://github.com/cjnoname/excelts/issues)

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for detailed version history.
