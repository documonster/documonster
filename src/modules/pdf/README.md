# PDF Module

[中文](README_zh.md)

A full-featured, zero-dependency PDF engine built from scratch in pure TypeScript. **Write** PDFs with the `pdf()` function or `excelToPdf()` bridge. **Read** any PDF with `readPdf()` — extract text, images, and metadata from all major PDF versions.

```typescript
// Write — standalone
import { pdf } from "@cj-tech-master/excelts/pdf";

// Write — from Excel
import { excelToPdf } from "@cj-tech-master/excelts/pdf";

// Read — extract text, images, metadata
import { readPdf } from "@cj-tech-master/excelts/pdf";
```

## Features

### Writing

- **Zero Dependencies** — Pure TypeScript PDF generation, no external libraries
- **PDF 2.0** — Writes modern PDF 2.0 format
- **Standalone Engine** — Use `pdf()` with plain arrays and objects, no Excel dependency
- **Excel Bridge** — One-line `excelToPdf(workbook)` for Excel-to-PDF conversion
- **Cross-Platform** — Same API in Node.js and browsers
- **Full Styling** — Fonts, colors, borders, fills, alignment, merged cells
- **Rich Text** — Mixed formatting within a single cell, with word-wrap support
- **Pagination** — Automatic vertical/horizontal page splitting with repeat headers
- **Images** — JPEG and PNG embedding with alpha transparency
- **AES-256 Encryption** — Password protection with AES-256 (V=5, R=5) and permission controls
- **Font Embedding** — TrueType font subsetting for Unicode/CJK text
- **Page Setup** — Per-sheet paper size, orientation, margins, print area
- **Tree-Shakeable** — Not imported? Not in your bundle

### Reading

- **Universal Reader** — Reads all major PDF versions (1.0 through 2.0)
- **Text Extraction** — Full text with line reconstruction, multi-column, and table detection
- **Multilingual** — WinAnsi, MacRoman, CJK via ToUnicode CMap, Identity-H/V, Symbol, ZapfDingbats
- **Image Extraction** — JPEG, JPEG2000, CCITT, JBIG2, raw/Flate with SMask/alpha
- **Annotation Extraction** — Links, comments, highlights, stamps, free text, and more
- **Form Fields** — AcroForm extraction: text inputs, checkboxes, radio buttons, dropdowns, signatures
- **Metadata** — Info dictionary + XMP (title, author, dates, page count, page sizes)
- **All Encryption** — RC4-40, RC4-128, AES-128, AES-256 (reads all versions)
- **Fault-Tolerant** — Cross-reference table/stream recovery, incremental updates

---

## Quick Start

### Read a PDF

```typescript
import { readPdf } from "@cj-tech-master/excelts/pdf";
import { readFileSync } from "fs";

const bytes = readFileSync("document.pdf");
const result = readPdf(bytes);

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

### Read Encrypted PDF

```typescript
const result = readPdf(bytes, { password: "secret" });
```

### Selective Extraction

```typescript
// Only pages 1 and 3, text only (no images)
const result = readPdf(bytes, {
  pages: [1, 3],
  extractImages: false
});
```

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

The PDF module is split into four layers:

```
src/modules/pdf/
├── core/               # PDF primitives (objects, streams, writer, encryption, crypto)
├── font/               # TTF parsing, glyph metrics, font subsetting, embedding
├── render/             # Layout engine, page renderer, style converter
│   ├── layout-engine   — PdfSheetData → LayoutPage[] (zero @excel imports)
│   ├── page-renderer   — LayoutPage → PDF content stream (zero @excel imports)
│   ├── style-converter — PdfCellStyle → PDF rendering params (zero @excel imports)
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
│   └── pdf-reader      — public API: readPdf()
├── types.ts            # PdfWorkbook, PdfSheetData, PdfCellData, etc.
├── excel-bridge.ts     # Excel Workbook → PdfWorkbook conversion (ONLY @excel dependency)
└── index.ts
```

The entire PDF engine (core, font, render, reader) has **zero imports from the Excel module**. The `excel-bridge.ts` is the only file that knows about Excel — it converts `Workbook` to `PdfWorkbook`.

**Write strategy:** Write PDF 2.0 only (modern, AES-256).
**Read strategy:** Read all major PDF versions (1.0 through 2.0, all encryption types).

---

## Writer Options

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

## Reader Options

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

### Reader Result

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

The PDF writer renders all standard cell styles:

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

The writer produces **AES-256 encrypted PDFs** (PDF 2.0, V=5, R=5). The reader can decrypt **all major encryption formats** including legacy RC4.

### Writer Encryption (AES-256)

#### Owner-Only (No Open Password)

```typescript
const pdf = excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    permissions: { print: true, copy: false, modify: false }
  }
});
// Opens without password, but copy/modify is restricted
```

#### Open Password Required

```typescript
const pdf = excelToPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "reader"
  }
});
// Requires "reader" to open
```

### Reader Decryption (All Formats)

The reader automatically detects and decrypts:

| Format  | Version    | Support |
| ------- | ---------- | ------- |
| RC4-40  | V=1, R=2   | Read    |
| RC4-128 | V=2, R=3   | Read    |
| AES-128 | V=4, R=4   | Read    |
| AES-256 | V=5, R=5/6 | Read    |

```typescript
// Automatically detects encryption type
const result = readPdf(encryptedBytes, { password: "secret" });
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

| File                   | What it demonstrates                                                           |
| ---------------------- | ------------------------------------------------------------------------------ |
| `pdf-basic.ts`         | Page sizes, margins, metadata, sheet selection, scale                          |
| `pdf-styled.ts`        | Fonts, fills, borders, alignment, merge, rotation, rich text, number formats   |
| `pdf-advanced.ts`      | Pagination, page breaks, encryption, transparency, bookmarks, hidden rows/cols |
| `pdf-excel-to-pdf.ts`  | Reading real `.xlsx` files and converting to PDF                               |
| `pdf-images.ts`        | Image embedding (JPEG, PNG with transparency)                                  |
| `pdf-reader.ts`        | Text extraction, metadata, images, encrypted PDFs, selective extraction        |
| `pdf-reader-stress.ts` | Large-scale stress test: thousands of cells, encrypted roundtrip, benchmarks   |

Run any example:

```bash
npx tsx src/modules/pdf/examples/pdf-basic.ts
# Output: tmp/pdf-examples/*.pdf

npx tsx src/modules/pdf/examples/pdf-reader.ts
# Output: tmp/pdf-reader-examples/
```

---

## API Reference

### `readPdf(data, options?)`

Read a PDF file and extract text, images, and metadata. Returns `ReadPdfResult`.

```typescript
import { readPdf } from "@cj-tech-master/excelts/pdf";

// Basic
const result = readPdf(pdfBytes);
console.log(result.text);
console.log(result.pages[0].images);
console.log(result.pages[0].annotations);
console.log(result.formFields);
console.log(result.metadata);

// Encrypted
const result = readPdf(pdfBytes, { password: "secret" });

// Selective
const result = readPdf(pdfBytes, {
  pages: [1, 3],
  extractImages: false,
  extractMetadata: false
});
```

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
