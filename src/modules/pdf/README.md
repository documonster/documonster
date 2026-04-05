# PDF Module

[中文](README_zh.md)

A full-featured, zero-dependency PDF engine built from scratch in pure TypeScript. Can be used **standalone** with the `pdf()` function, or as an **Excel-to-PDF converter** via the `excelToPdf()` bridge API.

```typescript
// Standalone PDF (simplest)
import { pdf } from "@cj-tech-master/excelts/pdf";

// Excel-to-PDF
import { excelToPdf } from "@cj-tech-master/excelts/pdf";
```

## Features

- **Zero Dependencies** — Pure TypeScript PDF generation, no external libraries
- **Standalone Engine** — Use `pdf()` with plain arrays and objects, no Excel dependency
- **Excel Bridge** — One-line `excelToPdf(workbook)` for Excel-to-PDF conversion
- **Cross-Platform** — Same API in Node.js and browsers
- **Full Styling** — Fonts, colors, borders, fills, alignment, merged cells
- **Rich Text** — Mixed formatting within a single cell, with word-wrap support
- **Pagination** — Automatic vertical/horizontal page splitting with repeat headers
- **Images** — JPEG and PNG embedding with alpha transparency
- **Encryption** — Password protection with 128-bit RC4 and permission controls
- **Font Embedding** — TrueType font subsetting for Unicode/CJK text
- **Page Setup** — Per-sheet paper size, orientation, margins, print area
- **Tree-Shakeable** — Not imported? Not in your bundle

---

## Quick Start

### Excel-to-PDF (Bridge API)

The simplest way to generate PDFs from Excel workbooks:

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

const pdf = excelToPdf(workbook);

// Node.js
import { writeFileSync } from "fs";
writeFileSync("output.pdf", pdf);

// Browser
const blob = new Blob([pdf], { type: "application/pdf" });
const url = URL.createObjectURL(blob);
window.open(url);
```

### Read XLSX, Export PDF

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
await workbook.xlsx.readFile("report.xlsx");

const pdf = excelToPdf(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "Monthly Report"
});
```

### Standalone PDF (No Excel)

Generate PDFs from plain data — no Excel module, no Map objects, no boilerplate:

```typescript
import { pdf } from "@cj-tech-master/excelts/pdf";

// Simplest — pass a 2D array
const bytes = pdf([
  ["Product", "Revenue"],
  ["Widget", 1000],
  ["Gadget", 2500]
]);

// With options
const bytes = pdf(
  [
    ["Name", "Score"],
    ["Alice", 95],
    ["Bob", 87]
  ],
  { showGridLines: true, title: "Scores" }
);

// Multiple sheets
const bytes = pdf({
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
const bytes = pdf({
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

## Architecture

The PDF module is split into three layers:

```
src/modules/pdf/
├── core/               # PDF primitives (objects, streams, writer, encryption)
├── font/               # TTF parsing, glyph metrics, font subsetting, embedding
├── render/             # Layout engine, page renderer, style converter
│   ├── layout-engine   — PdfSheetData → LayoutPage[] (zero @excel imports)
│   ├── page-renderer   — LayoutPage → PDF content stream (zero @excel imports)
│   ├── style-converter — PdfCellStyle → PDF rendering params (zero @excel imports)
│   └── pdf-exporter    — PdfWorkbook → Uint8Array (zero @excel imports)
├── types.ts            # PdfWorkbook, PdfSheetData, PdfCellData, etc.
├── excel-bridge.ts     # Excel Workbook → PdfWorkbook conversion (ONLY @excel dependency)
└── index.ts
```

The entire PDF engine (core, font, render) has **zero imports from the Excel module**. The `excel-bridge.ts` is the only file that knows about Excel — it converts `Workbook` to `PdfWorkbook`.

---

## Options

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

  // Encryption
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

### Page Sizes

Built-in page sizes accessible via `PageSizes`:

| Name        | Dimensions (pt)  | Millimeters |
| ----------- | ---------------- | ----------- |
| `"LETTER"`  | 612 x 792        | 216 x 279   |
| `"LEGAL"`   | 612 x 1008       | 216 x 356   |
| `"TABLOID"` | 792 x 1224       | 279 x 432   |
| `"A3"`      | 841.89 x 1190.55 | 297 x 420   |
| `"A4"`      | 595.28 x 841.89  | 210 x 297   |
| `"A5"`      | 419.53 x 595.28  | 148 x 210   |

Custom sizes: `{ width: 396, height: 612 }` (in points, 72pt = 1 inch).

---

## Styling Support

The PDF engine renders all standard cell styles:

### Text

- Font family, size, bold, italic
- Font color (ARGB, theme colors with tint)
- Underline and strikethrough
- Rich text (mixed formatting per cell)
- Number/date/currency formatting via `numFmt`
- Hyperlinks (clickable annotations)

### Alignment

- Horizontal: left, center, right
- Vertical: top, middle, bottom
- Text wrapping with word-break
- Text indentation
- Text rotation (angled and vertical stacked)

### Cells

- Background fills (solid colors with alpha transparency)
- Borders: thin, medium, thick, dashed, dotted (with colors)
- Merged cells (horizontal and vertical spans)

---

## Pagination

### Automatic

- Rows overflow the page height: automatic vertical page breaks
- Columns overflow the page width: automatic horizontal page splitting
- `fitToPage: true` (default): scales all columns to fit within page width

### Repeat Header Rows

```typescript
excelToPdf(workbook, { repeatRows: 2 }); // Repeat first 2 rows on every page
```

Or via worksheet page setup:

```typescript
worksheet.pageSetup.printTitlesRow = "1:2"; // Repeat rows 1-2
```

### Manual Page Breaks

```typescript
worksheet.getRow(20).addPageBreak(); // Break after row 20
```

### Print Area

```typescript
worksheet.pageSetup.printArea = "A1:F50"; // Export only this range
```

> **Note:** If a multi-range print area is set (e.g. `"A1:B5&&D1:E10"`), only the first range is used for PDF export.

---

## Images

JPEG and PNG images are embedded when sheets contain images:

```typescript
const imageId = workbook.addImage({
  buffer: jpegBytes,
  extension: "jpeg"
});

worksheet.addImage(imageId, {
  tl: { col: 0, row: 0 },
  ext: { width: 200, height: 150 }
});

const pdf = excelToPdf(workbook);
// Image appears in the PDF at the specified position
```

PNG transparency (RGBA and tRNS) is preserved via PDF soft masks.

---

## Encryption

### Owner-Only (No Open Password)

```typescript
const pdf = excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    permissions: { print: true, copy: false, modify: false }
  }
});
// Opens without password, but copy/modify is restricted
```

### Open Password Required

```typescript
const pdf = excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "reader"
  }
});
// Requires "reader" to open
```

---

## Unicode / CJK Support

Standard Type1 fonts (Helvetica, Times, Courier) only support Latin characters. For Unicode, CJK, or other scripts, provide a TrueType font:

```typescript
import { readFileSync } from "fs";

const pdf = excelToPdf(workbook, {
  font: readFileSync("NotoSansSC-Regular.ttf")
});
```

The font is automatically subsetted (only used glyphs are embedded) to minimize PDF file size.

---

## Per-Sheet Page Setup

Each sheet's `pageSetup` is respected when using the Excel bridge:

```typescript
const ws1 = workbook.addWorksheet("Summary");
ws1.pageSetup.paperSize = 9; // A4
ws1.pageSetup.orientation = "portrait";

const ws2 = workbook.addWorksheet("Data");
ws2.pageSetup.paperSize = 1; // Letter
ws2.pageSetup.orientation = "landscape";

// Each sheet renders with its own page size/orientation
const pdf = excelToPdf(workbook);
```

Worksheet margins are also inherited:

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

The PDF module is fully tree-shakeable. If you don't import any PDF exports, the module adds **zero bytes** to your bundle:

```typescript
// Only imports Excel core — PDF module is NOT included
import { Workbook } from "@cj-tech-master/excelts";

// Imports Excel + PDF bridge
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";
```

---

## Examples

Runnable examples are in `src/modules/pdf/examples/`:

| File                  | What it demonstrates                                                           |
| --------------------- | ------------------------------------------------------------------------------ |
| `pdf-basic.ts`        | Page sizes, margins, metadata, sheet selection, scale                          |
| `pdf-styled.ts`       | Fonts, fills, borders, alignment, merge, rotation, rich text, number formats   |
| `pdf-advanced.ts`     | Pagination, page breaks, encryption, transparency, bookmarks, hidden rows/cols |
| `pdf-excel-to-pdf.ts` | Reading real `.xlsx` files and converting to PDF                               |

Run any example:

```bash
npx tsx src/modules/pdf/examples/pdf-basic.ts
# Output: tmp/pdf-examples/*.pdf
```

---

## API Reference

### `pdf(input, options?)`

Generate a PDF from plain data. Returns `Uint8Array`.

```typescript
// 2D array
pdf([["Name", "Age"], ["Alice", 30]]);

// Single sheet with column widths
pdf({ name: "Report", columns: [{ width: 25 }, 15], data: [["A", "B"]] });

// Multiple sheets
pdf({ sheets: [{ name: "S1", data: [...] }, { name: "S2", data: [...] }] });

// With options
pdf([["A", 1]], { showGridLines: true, pageSize: "A4" });
```

### `excelToPdf(workbook, options?)`

Convert an Excel `Workbook` to PDF. Returns `Uint8Array`.

```typescript
import { Workbook, excelToPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
// ... build workbook ...
const bytes = excelToPdf(workbook, { showGridLines: true });
```

### Error Types

```typescript
import {
  PdfError, // Base class for all PDF errors
  PdfRenderError, // Layout/rendering failures
  PdfFontError, // Font parsing/embedding failures
  PdfStructureError, // PDF structure assembly failures
  isPdfError // Type guard: (err: unknown) => err is PdfError
} from "@cj-tech-master/excelts";
```

All errors extend `BaseError` with `cause` chain support:

```typescript
try {
  excelToPdf(workbook);
} catch (err) {
  if (isPdfError(err)) {
    console.error(err.message, err.cause);
  }
}
```
