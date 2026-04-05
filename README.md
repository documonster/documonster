# ExcelTS

[![Build Status](https://github.com/cjnoname/excelts/actions/workflows/ci.yml/badge.svg?branch=main&event=push)](https://github.com/cjnoname/excelts/actions/workflows/ci.yml) &nbsp; [中文](README_zh.md)

Zero-dependency, AI-friendly TypeScript toolkit — Excel, PDF, CSV, Markdown, XML, ZIP, and streaming in one library with seven standalone modules.

## About This Project

ExcelTS is a zero-dependency TypeScript toolkit for spreadsheets and documents:

- **AI-Friendly** — Clean, consistent API designed for AI coding agents. Every module has comprehensive documentation and runnable examples for AI to learn from
- **Zero Runtime Dependencies** — Pure TypeScript, no external packages
- **Seven Modules** — Excel, PDF, CSV, Markdown, XML, Archive, Stream
- **Cross-Platform** — Node.js 22+, Bun, Chrome 89+, Firefox 102+, Safari 14.1+
- **ESM First** — Native ES Modules with CommonJS compatibility and full tree-shaking

## Modules

ExcelTS is organized into seven standalone modules. Each module has its own documentation and runnable examples.

### Excel — XLSX/JSON Workbook Manager

Create, read, and modify Excel spreadsheets with full styling, formulas, images, and streaming support.

- [Documentation](src/modules/excel/README.md) | [中文](src/modules/excel/README_zh.md)
- [Examples](src/modules/excel/examples/)

### PDF — Zero-Dependency PDF Engine

Full-featured PDF generation with font embedding, encryption, images, and Excel-to-PDF conversion.

- [Documentation](src/modules/pdf/README.md) | [中文](src/modules/pdf/README_zh.md)
- [Examples](src/modules/pdf/examples/)

### CSV — RFC 4180 Parser/Formatter

High-performance CSV parsing and formatting with streaming, dynamic typing, data generation, and worker pool support.

- [Documentation](src/modules/csv/README.md) | [中文](src/modules/csv/README_zh.md)
- [Examples](src/modules/csv/examples/)

### Markdown — GFM Table Parser/Formatter

Parse and format GitHub Flavored Markdown tables with alignment round-trip and Workbook integration.

- [Documentation](src/modules/markdown/README.md) | [中文](src/modules/markdown/README_zh.md)
- [Examples](src/modules/markdown/examples/)

### XML — SAX/DOM Parser, Query Engine, Writer

Streaming and buffered XML processing with query engine, namespace support, and dual-mode writing.

- [Documentation](src/modules/xml/README.md) | [中文](src/modules/xml/README_zh.md)
- [Examples](src/modules/xml/examples/)

### Archive — ZIP/TAR Create/Read/Edit

ZIP and TAR archive creation, reading, editing, streaming, encryption, and compression utilities.

- [Documentation](src/modules/archive/README.md) | [中文](src/modules/archive/README_zh.md)
- [Examples](src/modules/archive/examples/)

### Stream — Cross-Platform Streaming

Node.js-compatible Readable/Writable/Transform/Duplex that works identically in Node.js and browsers.

- [Documentation](src/modules/stream/README.md) | [中文](src/modules/stream/README_zh.md)
- [Examples](src/modules/stream/examples/)

## Installation

```bash
npm install @cj-tech-master/excelts
```

## Quick Start

```typescript
import { Workbook } from "@cj-tech-master/excelts";

// Create
const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sheet1");
sheet.addRow(["Name", "Age"]);
sheet.addRow(["Alice", 30]);
await workbook.xlsx.writeFile("output.xlsx");

// Read
const wb = new Workbook();
await wb.xlsx.readFile("output.xlsx");
wb.getWorksheet(1).eachRow((row, n) => console.log(n, row.values));
```

```typescript
// PDF — generate from data, no Workbook needed
import { pdf } from "@cj-tech-master/excelts/pdf";
const pdfBytes = pdf([
  ["Product", "Revenue"],
  ["Widget", 1000]
]);

// CSV — parse and format
import { parseCsv, formatCsv } from "@cj-tech-master/excelts/csv";
const rows = parseCsv("name,age\nAlice,30", { headers: true });
const csv = formatCsv([{ name: "Bob", age: 25 }], { headers: true });

// XML — parse, query, write
import { parseXml, queryAll, XmlWriter } from "@cj-tech-master/excelts/xml";
const titles = queryAll(parseXml(xmlString).root, "book/title");

// ZIP — create and extract
import { zip, unzip } from "@cj-tech-master/excelts/zip";
const archive = await zip().add("hello.txt", "Hello!").bytes();

// Markdown — parse and format tables
import { parseMarkdown, formatMarkdown } from "@cj-tech-master/excelts/markdown";
const table = parseMarkdown("| A | B |\n|---|---|\n| 1 | 2 |");
```

## Subpath Exports

Each module is available as a standalone subpath export:

```typescript
import { Workbook, WorkbookWriter } from "@cj-tech-master/excelts";
import { SaxParser, parseXml, XmlWriter, query } from "@cj-tech-master/excelts/xml";
import { zip, unzip, ZipArchive, compress } from "@cj-tech-master/excelts/zip";
import { parseCsv, formatCsv, CsvParserStream } from "@cj-tech-master/excelts/csv";
import { parseMarkdown, formatMarkdown, parseMarkdownAll } from "@cj-tech-master/excelts/markdown";
import { Readable, pipeline, createTransform } from "@cj-tech-master/excelts/stream";
```

Each subpath supports `browser`, `import` (ESM), and `require` (CJS) conditions.

## Browser Support

ExcelTS has native browser support with **zero configuration** for modern bundlers.

```typescript
// Bundlers (Vite, Webpack, Rollup, esbuild) — just import
import { Workbook } from "@cj-tech-master/excelts";
const buffer = await new Workbook().addWorksheet("S1").workbook.xlsx.writeBuffer();
```

```html
<!-- Script tag (no bundler) -->
<script src="https://unpkg.com/@cj-tech-master/excelts/dist/iife/excelts.iife.min.js"></script>
```

For older browsers without native `CompressionStream` API, ExcelTS automatically uses a built-in pure JavaScript DEFLATE implementation — no polyfills needed.

## Requirements

- **Node.js >= 22.0.0**

| Browser | Minimum Version    |
| ------- | ------------------ |
| Chrome  | 89+ (March 2021)   |
| Edge    | 89+ (March 2021)   |
| Firefox | 102+ (June 2022)   |
| Safari  | 14.1+ (April 2021) |
| Opera   | 75+ (March 2021)   |

## Links

- 🏠 [GitHub Repository](https://github.com/cjnoname/excelts)
- 🐛 [Issue Tracker](https://github.com/cjnoname/excelts/issues)
- 📋 [Changelog](CHANGELOG.md)
- 🔄 [Migration Guide](MIGRATION.md)
- 🗺️ [Roadmap](ROADMAP.md)
- 📄 [License (MIT)](LICENSE)
- 📦 [Third-Party Notices](THIRD_PARTY_NOTICES.md)
