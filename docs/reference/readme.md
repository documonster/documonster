# README

# ExcelTS

[![Build Status](https://github.com/cjnoname/excelts/actions/workflows/ci.yml/badge.svg?branch=main&event=push)](https://github.com/cjnoname/excelts/actions/workflows/ci.yml)

Modern TypeScript Excel Workbook Manager - Read, manipulate and write spreadsheet data and styles to XLSX and JSON.

## About This Project

ExcelTS is a modern TypeScript Excel workbook manager with:

- 🚀 **Zero Runtime Dependencies** - Pure TypeScript implementation with no external packages
- ✅ **Full TypeScript Support** - Complete type definitions and modern TypeScript patterns
- ✅ **Modern Build System** - Using Rolldown for faster builds
- ✅ **Enhanced Testing** - Migrated to Vitest with browser testing support
- ✅ **ESM First** - Native ES Module support with CommonJS compatibility
- ✅ **Node 20+** - Optimized for modern Node.js versions
- ✅ **Named Exports** - All exports are named for better tree-shaking
- ✅ **Broad Browser Support** - Works in Chrome 89+, Firefox 102+, Safari 14.1+ (with built-in fallbacks for missing `CompressionStream`)

## Translations

- [中文文档](/reference/readme-zh)

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

- **Advanced Features**
  - Streaming for large files
  - CSV import/export
  - Tables with auto-filters
  - Page setup and printing options
  - Data protection
  - Comments and notes

## Streaming API

For processing large Excel files without loading them entirely into memory, ExcelTS provides streaming reader and writer APIs.

- **Node.js**: `WorkbookReader` supports reading from a file path, and `WorkbookWriter` supports writing to a filename.
- **Browsers**: use `Uint8Array` / `ArrayBuffer` / Web `ReadableStream<Uint8Array>` for reading, and Web `WritableStream<Uint8Array>` for writing.
- Note: ExcelTS does not re-export the internal stream utility surface (e.g. `Readable`, `Writable`). Prefer standard Web Streams (browser/Node 20+) or Node.js streams.

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

### Web Streams (Node.js 20+ and Browsers)

`WorkbookWriter` can write to a Web `WritableStream<Uint8Array>`, and `WorkbookReader` can read from a Web `ReadableStream<Uint8Array>`.

This does **not** require importing any extra stream utility surface from ExcelTS; it uses the standard Web Streams API.

- Full runnable example: [src/modules/excel/examples/web-streams-reader-writer.ts](https://github.com/cjnoname/excelts/blob/main/src/modules/excel/examples/web-streams-reader-writer.ts)

Run locally (Node.js 20+):

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

const workbook = new Workbook();

// Read CSV from file (streaming)
await workbook.csv.readFile("data.csv");

// Read CSV from stream
import fs from "fs";
const stream = fs.createReadStream("data.csv");
await workbook.csv.read(stream, { sheetName: "Imported" });

// Write CSV to file (streaming)
await workbook.csv.writeFile("output.csv");

// Write CSV to stream
const writeStream = fs.createWriteStream("output.csv");
await workbook.csv.write(writeStream);

// Write CSV to buffer
const buffer = await workbook.csv.writeBuffer();
```

### Browser (In-Memory)

```javascript
import { Workbook } from "@cj-tech-master/excelts";

const workbook = new Workbook();

// Load CSV from string
workbook.csv.load(csvString);

// Load CSV from ArrayBuffer (e.g., from fetch or file input)
const response = await fetch("data.csv");
const arrayBuffer = await response.arrayBuffer();
workbook.csv.load(arrayBuffer);

// Write CSV to string
const csvOutput = workbook.csv.writeString();

// Write CSV to Uint8Array buffer
const buffer = workbook.csv.writeBuffer();
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

- [src/modules/excel/examples/browser-smoke.html](https://github.com/cjnoname/excelts/blob/main/src/modules/excel/examples/browser-smoke.html)

Steps:

```bash
npm run build:browser:bundle
npx serve .
```

Then open `http://localhost:3000/src/modules/excel/examples/browser-smoke.html`.

### Browser-Specific Notes

- **CSV operations are supported** using native RFC 4180 implementation
  - Use `csv.load(stringOrArrayBuffer)` to read CSV
  - Use `csv.writeString()` or `csv.writeBuffer()` to write CSV
- Use `xlsx.load(arrayBuffer)` instead of `xlsx.readFile()`
- Use `xlsx.writeBuffer()` instead of `xlsx.writeFile()`
- Worksheet protection with passwords is fully supported (pure JS SHA-512)

## Requirements

### Node.js

- **Node.js >= 20.0.0** (ES2020 native support)

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

See [CHANGELOG.md](/reference/changelog) for detailed version history.
