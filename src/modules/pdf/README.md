# PDF Module

Zero-dependency Excel-to-PDF export with full styling, pagination, and encryption support.

```typescript
import { exportPdf, PdfExporter } from "@cj-tech-master/excelts";
```

## Features

- **Zero Dependencies** - Pure TypeScript PDF generation, no external libraries
- **Cross-Platform** - Same API in Node.js and browsers
- **Full Styling** - Fonts, colors, borders, fills, alignment, merged cells
- **Rich Text** - Mixed formatting within a single cell, with word-wrap support
- **Pagination** - Automatic vertical/horizontal page splitting with repeat headers
- **Images** - JPEG and PNG embedding with alpha transparency
- **Encryption** - Password protection with 128-bit RC4 and permission controls
- **Font Embedding** - TrueType font subsetting for Unicode/CJK text
- **Page Setup** - Per-worksheet paper size, orientation, margins, print area
- **Tree-Shakeable** - Not imported? Not in your bundle

---

## Quick Start

### One-Line Export

```typescript
import { Workbook, exportPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
const sheet = workbook.addWorksheet("Sales");
sheet.columns = [
  { header: "Product", key: "product", width: 20 },
  { header: "Revenue", key: "revenue", width: 15 }
];
sheet.addRow({ product: "Widget", revenue: 1000 });
sheet.addRow({ product: "Gadget", revenue: 2500 });
sheet.getColumn("revenue").numFmt = "$#,##0.00";

const pdf = exportPdf(workbook);

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
import { Workbook, exportPdf } from "@cj-tech-master/excelts";

const workbook = new Workbook();
await workbook.xlsx.readFile("report.xlsx");

const pdf = exportPdf(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "Monthly Report"
});
```

### PdfExporter Class

```typescript
import { Workbook, PdfExporter } from "@cj-tech-master/excelts";

const workbook = new Workbook();
// ... populate workbook ...

const exporter = new PdfExporter(workbook);
const pdf = exporter.export({
  pageSize: "A4",
  orientation: "landscape",
  showGridLines: true,
  showSheetNames: true,
  showPageNumbers: true
});
```

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

The PDF exporter renders all standard Excel cell styles:

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
exportPdf(workbook, { repeatRows: 2 }); // Repeat first 2 rows on every page
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

---

## Images

JPEG and PNG images are embedded when worksheets contain images:

```typescript
const imageId = workbook.addImage({
  buffer: jpegBytes,
  extension: "jpeg"
});

worksheet.addImage(imageId, {
  tl: { col: 0, row: 0 },
  ext: { width: 200, height: 150 }
});

const pdf = exportPdf(workbook);
// Image appears in the PDF at the specified position
```

PNG transparency (RGBA and tRNS) is preserved via PDF soft masks.

---

## Encryption

### Owner-Only (No Open Password)

```typescript
const pdf = exportPdf(workbook, {
  encryption: {
    ownerPassword: "admin",
    permissions: { print: true, copy: false, modify: false }
  }
});
// Opens without password, but copy/modify is restricted
```

### Open Password Required

```typescript
const pdf = exportPdf(workbook, {
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

const pdf = exportPdf(workbook, {
  font: readFileSync("NotoSansSC-Regular.ttf")
});
```

The font is automatically subsetted (only used glyphs are embedded) to minimize PDF file size.

---

## Per-Worksheet Page Setup

Each worksheet's `pageSetup` is respected:

```typescript
const ws1 = workbook.addWorksheet("Summary");
ws1.pageSetup.paperSize = 9; // A4
ws1.pageSetup.orientation = "portrait";

const ws2 = workbook.addWorksheet("Data");
ws2.pageSetup.paperSize = 1; // Letter
ws2.pageSetup.orientation = "landscape";

// Each sheet renders with its own page size/orientation
const pdf = exportPdf(workbook);
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

// Imports Excel + PDF module (~136 KB additional)
import { Workbook, exportPdf } from "@cj-tech-master/excelts";
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

### `exportPdf(workbook, options?)`

Convenience function. Returns `Uint8Array` containing the PDF file.

```typescript
function exportPdf(workbook: Workbook, options?: PdfExportOptions): Uint8Array;
```

### `PdfExporter`

Class-based API for advanced usage.

```typescript
class PdfExporter {
  constructor(workbook: Workbook);
  export(options?: PdfExportOptions): Uint8Array;
}
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
  exportPdf(workbook);
} catch (err) {
  if (isPdfError(err)) {
    console.error(err.message, err.cause);
  }
}
```
