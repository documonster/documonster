# PDF Module

[中文](README_zh.md)

A full-featured, zero-dependency PDF engine built from scratch in pure TypeScript. **Write** PDFs with `Pdf.fromTable()` or the `Pdf.fromExcel()` bridge. **Read** any PDF with `Pdf.extract()` — extract text, images, and metadata from all major PDF versions. **Build** free-form PDFs with `Pdf.Builder` — text, vector graphics, SVG paths, annotations, and form fields. **Edit** existing PDFs with `Pdf.Editor` — overlay content, fill forms, add/remove/rotate pages, merge documents, and sign digitally. All APIs are async and yield to the event loop between pages to avoid blocking.

```typescript
import { Pdf } from "documonster/pdf";

// Write — standalone:      Pdf.fromTable(rows)
// Write — from Excel:      await Pdf.fromExcel(workbook)
// Read — text/images/meta: Pdf.extract(bytes)
// Build — free-form:       new Pdf.Builder()
// Edit — overlay/forms:    Pdf.Editor.read(bytes)
// Sign — digital sigs:     builder.sign(...) / Pdf.verifySignatures(bytes)
```

## Features

### Writing

- **Zero Dependencies** — Pure TypeScript PDF generation, no external libraries
- **PDF 2.0** — Writes modern PDF 2.0 format
- **Standalone Engine** — Use `Pdf.fromTable()` with plain arrays and objects, no Excel dependency
- **Excel Bridge** — One-line `Pdf.fromExcel(workbook)` for Excel-to-PDF conversion
- **Cross-Platform** — Same API in Node.js and browsers
- **Full Styling** — Fonts, colors, borders, fills, alignment, merged cells
- **Rich Text** — Mixed formatting within a single cell, with word-wrap support
- **Pagination** — Automatic vertical/horizontal page splitting with repeat headers
- **Images** — JPEG and PNG embedding with alpha transparency
- **AES-256 Encryption** — Password protection with AES-256 (V=5, R=5) and permission controls
- **Font Embedding** — TrueType font subsetting for Unicode/CJK text
- **Watermarks** — Text and image watermarks with opacity, rotation, tiling, per-page/per-sheet filtering
- **Page Setup** — Full print-setting fidelity: paper size, orientation, margins, print area, scaling, page order, print titles, headings, center-on-page, black and white, draft, cell errors
- **Tree-Shakeable** — Not imported? Not in your bundle
- **Non-Blocking** — Yields to the event loop between pages to avoid blocking

### Reading

- **Universal Reader** — Reads all major PDF versions (1.0 through 2.0)
- **Text Extraction** — Full text with line reconstruction, multi-column, and table detection
- **Multilingual** — WinAnsi, MacRoman, CJK via ToUnicode CMap, Identity-H/V, Symbol, ZapfDingbats
- **Image Extraction** — JPEG, JPEG2000, CCITT, JBIG2, raw/Flate with SMask/alpha
- **Annotation Extraction** — Links, comments, highlights, stamps, free text, and more
- **Form Fields** — AcroForm extraction: text inputs, checkboxes, radio buttons, dropdowns, signatures
- **Bookmark Extraction** — Nested outline tree with named/action destinations
- **Table Extraction** — Heuristic table detection from text fragment positioning
- **Metadata** — Info dictionary + XMP (title, author, dates, page count, page sizes)
- **All Encryption** — RC4-40, RC4-128, AES-128, AES-256 (reads all versions)
- **Fault-Tolerant** — Cross-reference table/stream recovery, incremental updates

### Building (PdfDocumentBuilder)

- **Free Text Positioning** — Place text anywhere with font, size, color, bold/italic
- **Vector Graphics** — Rectangles, circles, ellipses, lines, arbitrary paths with fill/stroke
- **SVG Path Rendering** — Parse and render SVG `d` attributes (all commands including arcs)
- **Images** — JPEG and PNG embedding at any position
- **Annotations** — Create Highlight, Underline, StrikeOut, Squiggly, Text (sticky note), FreeText, and Stamp annotations
- **Form Field Creation** — Create TextField, Checkbox, Dropdown, and RadioGroup from scratch
- **Bookmarks** — Nested outline tree with page destinations
- **Table of Contents** — Auto-generated TOC with dot leaders, page numbers, and clickable links
- **PDF/A-1b** — Archival compliance with XMP metadata, OutputIntent, and sRGB ICC profile
- **AES-256 Encryption** — Password-protect builder-created PDFs
- **Font Embedding** — TrueType font subsetting for Unicode/CJK text

### Editing (PdfEditor)

- **Overlay Content** — Draw text, shapes, images on existing PDF pages
- **Annotations on Existing Pages** — Add Highlight, Text, FreeText, Stamp, etc. to existing PDFs
- **Form Fields on Existing Pages** — Add TextField, Checkbox, Dropdown to existing PDFs
- **SVG Paths on Existing Pages** — Draw SVG paths on existing PDF pages
- **Form Filling** — Set text field values and checkbox states
- **Add Pages** — Append new blank pages with content
- **Remove Pages** — Delete pages by index
- **Rotate Pages** — Rotate pages by 90/180/270 degrees
- **Split Pages** — Split a PDF into individual single-page PDFs
- **Merge / Copy Pages** — Copy pages from other PDFs (including encrypted sources)
- **Incremental Save** — Append-only updates preserving original bytes (safe for signed PDFs)
- **Full Save** — Rebuild entire PDF with all modifications
- **Metadata Preservation** — Retains XMP, page properties (Rotate, CropBox, etc.)

### Digital Signatures

- **Signature Verification** — `Pdf.verifySignatures(bytes)` finds every signature field and checks each one's integrity (digest, RSA, `/ByteRange`, whole-file coverage) and reports the certificate it carries — no trust evaluation; `Pdf.Signature.verify` checks one signature you located yourself
- **Signature Creation** — Create CMS SignedData signatures with ByteRange placeholder/backfill
- **ASN.1 DER Codec** — Parse and encode ASN.1 structures (shared by verify and sign)
- **X.509 Certificate** — Extract public keys from DER-encoded certificates
- **Platform-Native Crypto** — Uses `node:crypto` on Node.js, Web Crypto API in browsers

---

## Quick Start

### Read a PDF

```typescript
import { readFileSync } from "fs";

import { Pdf } from "documonster/pdf";

const bytes = readFileSync("document.pdf");
const result = await Pdf.extract(bytes);

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

// Bookmarks (document outline)
for (const bm of result.bookmarks) {
  console.log(bm.title, bm.pageIndex);
}
```

### Read Encrypted PDF

```typescript
const result = await Pdf.extract(bytes, { password: "secret" });
```

### Selective Extraction

```typescript
// Only pages 1 and 3, text only (no images)
const result = await Pdf.extract(bytes, {
  pages: [1, 3],
  extractImages: false
});

// Extract bookmarks (outline tree)
const result = await Pdf.extract(bytes, { extractBookmarks: true });
for (const bm of result.bookmarks) {
  console.log(bm.title, bm.pageIndex === undefined ? "(no page)" : `→ page ${bm.pageIndex + 1}`);
}

// Extract tables (heuristic text-position detection)
const result = await Pdf.extract(bytes, { extractTables: true });
for (const page of result.pages) {
  for (const table of page.tables) {
    for (const row of table.rows) {
      console.log(row.cells.map(c => c.text).join(" | "));
    }
  }
}
```

### Excel-to-PDF (Bridge API)

The simplest way to generate PDFs from Excel workbooks:

```typescript
import { Column, Row, Workbook } from "documonster/excel";
import { Pdf } from "documonster/pdf";

const workbook = Workbook.create();
const sheet = Workbook.addWorksheet(workbook, "Sales");
Column.setDefinitions(sheet, [
  { header: "Product", key: "product", width: 20 },
  { header: "Revenue", key: "revenue", width: 15 }
]);
Row.add(sheet, { product: "Widget", revenue: 1000 });
Row.add(sheet, { product: "Gadget", revenue: 2500 });
Column.setStyle(sheet, "revenue", { numFmt: "$#,##0.00" });

const pdf = await Pdf.fromExcel(workbook);

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
import { Workbook } from "documonster/excel";
import { Pdf } from "documonster/pdf";

const workbook = await Workbook.readFile("report.xlsx");

const pdf = await Pdf.fromExcel(workbook, {
  showGridLines: true,
  showPageNumbers: true,
  title: "Monthly Report"
});
```

### Standalone PDF (No Excel)

Generate PDFs from plain data — no Excel module, no Map objects, no boilerplate:

```typescript
import { Pdf } from "documonster/pdf";

// Simplest — pass a 2D array
const bytes = await Pdf.fromTable([
  ["Product", "Revenue"],
  ["Widget", 1000],
  ["Gadget", 2500]
]);

// With options
const bytes = await Pdf.fromTable(
  [
    ["Name", "Score"],
    ["Alice", 95],
    ["Bob", 87]
  ],
  { showGridLines: true, title: "Scores" }
);

// Multiple sheets
const bytes = await Pdf.fromTable({
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
const bytes = await Pdf.fromTable({
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

A `Date` cell is shown from its UTC fields, as `Pdf.fromExcel` shows a date cell with no
number format: `2024-01-15` at midnight UTC, `2024-01-15 13:05:09` otherwise — the same on
every machine (it used to be `toLocaleDateString()`, which varied by locale and, west of UTC,
by day). Build dates with `Date.UTC(…)`.

### Watermarks

Add text or image watermarks to any PDF generated via `Pdf.fromTable()` or `Pdf.fromExcel()`:

```text
// Text watermark — centered, semi-transparent, rotated
const bytes = await Pdf.fromTable(data, {
  watermark: {
    type: "text",
    text: "CONFIDENTIAL",
    fontSize: 48,
    color: { r: 0.8, g: 0.8, b: 0.8 },
    opacity: 0.3,
    rotation: -45,
    position: "center"
  }
});

// Image watermark — tiled across every page
const bytes = await Pdf.fromExcel(workbook, {
  watermark: {
    type: "image",
    data: logoPngBytes,
    format: "png",
    width: 100,
    height: 50,
    opacity: 0.1,
    repeat: true,
    repeatSpacingX: 150,
    repeatSpacingY: 100
  }
});

// Watermark on specific pages or sheets only
const bytes = await Pdf.fromTable(data, {
  watermark: {
    type: "text",
    text: "DRAFT",
    fontSize: 60,
    color: { r: 1, g: 0, b: 0 },
    opacity: 0.2,
    pages: [1], // Only first page
    sheets: ["Summary"] // Only "Summary" sheet
  }
});
```

A `watermark.sheets` name is compared as Excel compares names (ignoring case), and one that names
no sheet is refused with `PdfArgumentError` — as an unknown name in `sheets` is — rather than
leaving every page without the watermark. A `watermark.pages` number past the last page matches
no page and is skipped, as a page number past the end in `Pdf.extract({ pages })` is. `sheets: []`
selects every visible sheet, like leaving `sheets` out.

### Build Free-Form PDFs (PdfDocumentBuilder)

Create PDFs with precise control over text, shapes, and layout:

```text
const doc: Pdf.Builder;
doc.setMetadata({ title: "My Report", author: "documonster" });

const page = doc.addPage({ width: 595, height: 842 }); // A4

// Font configuration may follow addPage(), but must precede the first drawText().
doc.embedFont(fontFileBytes);

// Text
page.drawText("Hello, World!", { x: 72, y: 770, fontSize: 24, bold: true });

// Shapes
page.drawRect({ x: 72, y: 700, width: 200, height: 50, fill: { r: 0.2, g: 0.4, b: 0.8 } });
page.drawCircle({ cx: 400, cy: 725, r: 25, fill: { r: 1, g: 0, b: 0 } });

// SVG paths
page.drawSvgPath("M 100 600 C 150 500 250 500 300 600", {
  stroke: { r: 0, g: 0.5, b: 0 },
  lineWidth: 2
});

// Annotations
page.addAnnotation({
  type: "Highlight",
  rect: [72, 765, 250, 785],
  color: { r: 1, g: 1, b: 0 }
});

// Form fields
page.addFormField({
  type: "text",
  name: "email",
  rect: [72, 550, 300, 575]
});

// Encryption
doc.setEncryption({ ownerPassword: "admin", userPassword: "reader" });

const bytes = await doc.toBuffer();
```

### Edit Existing PDFs (PdfEditor)

Overlay content, fill forms, merge documents, and manipulate pages:

```typescript
import { Pdf } from "documonster/pdf";

const editor = Pdf.Editor.read(existingPdfBytes);

// Overlay text and shapes on page 1
const page = editor.getPage(0);
page.drawText("CONFIDENTIAL", { x: 200, y: 400, fontSize: 36, color: { r: 1, g: 0, b: 0 } });

// Add annotations to existing pages
page.addAnnotation({ type: "Highlight", rect: [72, 700, 300, 720] });

// Add form fields to existing pages
page.addFormField({ type: "text", name: "note", rect: [72, 650, 300, 675] });

// Draw SVG paths on existing pages
page.drawSvgPath("M 100 600 L 200 600 L 150 550 Z", { fill: { r: 0, g: 0.5, b: 1 } });

// Fill form fields
editor.setFormField("name", "Jane Doe");
editor.setFormField("agree", "Yes");

// Page manipulation
editor.removePage(2); // Remove page 3 — idempotent; indices name the pages as read and never shift
editor.rotatePage(0, 90); // Rotate page 1
editor.addPage(); // Add blank page

// Copy pages from another PDF — with their annotations and form fields
editor.copyPagesFrom(otherPdfBytes);

// Document properties: /Info and the XMP packet are updated together
editor.setMetadata({ title: "Q3 report", author: "Ann; Bob", keywords: "finance, Q3" });
editor.setMetadata({ subject: null }); // null removes a property (a JSON merge patch)

// Save (full rebuild or incremental append)
const result = await editor.toBuffer();
const incremental = await editor.toBuffer({ incremental: true }); // preserves original bytes
```

On a signed document (`editor.hasSignatures`), `toBuffer()`, `sign()` and a
`toBuffer({ incremental: true })` that must fall back to a full rewrite throw
`PdfSignatureInvalidationError` unless you pass `{ invalidateSignatures: true }`.
(The builder's `doc.setMetadata({ creationDate, modDate })` and the export options of the
same names fix the dates for reproducible output.)

`editor.setMetadata` takes a `PdfMetadataUpdate` — a JSON merge patch: a value sets a field,
`null` removes it (the `/Info` entry and the XMP property, in whatever form the packet holds
it), and a field left out is kept. It changes only the fields it is given, in `/Info` and in
the XMP packet alike; the editor never adds a `/Creator` or `/Producer` of its own, and no date changes
unless you pass it (pass `modDate` to record the edit). The XMP packet is edited as XML: only
the `x-default` entry of a title or description is replaced, `author` is split on `;` into
`dc:creator` entries, and `xmp:MetadataDate` is updated. A packet that cannot be decoded or is
not XMP is left untouched. A copied page keeps its MediaBox origin, inherited rotation and
crop box, and its annotations; its form fields join the document's form, and a field name the
document already uses is refused with `PdfArgumentError` when it is written.

A document counts as signed when a signature field (`/FT /Sig`, possibly
inherited from a parent field) has a value with `/ByteRange` and `/Contents`;
`/SigFlags` alone does not. `toBuffer({ incremental: true })` leaves the signed byte range
intact, but certification (DocMDP) and FieldMDP permissions are **not**
evaluated — an update a certifying signature forbids is still written.

Encrypted documents stay encrypted, with the file's own security handler (RC4, AES-128 or
AES-256 R5/R6): `toBuffer({ incremental: true })` encrypts the appended objects with it,
and a full rewrite — `toBuffer()`, `sign()`, `splitPages()`, or an update that falls back
— encrypts every object again under the original `/Encrypt` dictionary, key and `/ID`. The
result opens with the same user and owner passwords and grants the same permissions,
whichever password the editor was opened with. To write a **decrypted** copy without
`/Encrypt`, pass `{ encryption: "remove" }`; an update cannot drop encryption, so with
`incremental: true` that rewrites the file, and with `incremental: "required"` it throws
`PdfIncrementalUpdateError`. If the `/Encrypt` dictionary is unreadable, a rewrite throws
`PdfEncryptionError` rather than write the document unencrypted.

The deprecated `save()` writes a decrypted copy, as before — and so does the deprecated
`saveIncremental()` when it falls back to it; `toBuffer()` keeps the encryption by default.
Pass `{ encryption: "keep" }` to either to keep it.

A write reflects the edits made before it was called: an edit made while it is in flight,
or queued behind another write on the same editor, belongs to the next write.

Page indices are 0-based, and one index names each page for the editor's lifetime: `0` to
`pageCount - 1` are the pages the document was read with, and the pages added with `addPage()`
or `copyPagesFrom()` are numbered after them, in the order added. `getPage`, `removePage`,
`rotatePage`, `isPageRemoved` and `splitPages` all take the whole range; any index that names
no page — out of range, fractional, `NaN`, not a number — throws `PdfArgumentError` with
`field` naming the argument (`index`, or `options.pageIndices[i]` for `splitPages`). Removing
a page renumbers nothing. `pageCount` stays the number of pages the document was read with;
`outputPageCount` is the number a write produces now. `getPage` returns one handle per page:
for a read or copied page it draws over the page's content, for a page added with `addPage()`
onto that page (the builder `addPage()` returned draws onto it too). Removing an added page
leaves it out of the output; rotating one sets its `/Rotate`.
`splitPages({ pageIndices })` and `copyPagesFrom(source, { pageIndices })` name pages, so an
index that names none — out of range, removed, not an integer — is refused with
`PdfArgumentError` (field `options.pageIndices[i]`) before anything is written or copied;
`splitPages` returns one part per index, in order. (Their deprecated positional forms keep
skipping an index out of range.) A removed page cannot be edited:
`getPage` and `rotatePage` refuse it, and a page handle — or the builder `addPage()` returned —
held from before the removal refuses every drawing call, with `PdfArgumentError` — what was drawn on it used to vanish. A form field name
already in use (in the document, added, or brought by a copied page) is refused by the
`addFormField` or `copyPagesFrom` call that introduces it; to merge the same form twice (or
two forms sharing field names), copy with `{ formFields: "drop" }`, which copies the pages and
their other annotations without the form fields. The readers (`Pdf.extract`,
`Pdf.Editor.read`, `copyPagesFrom`, `Pdf.verifySignatures` and `Pdf.Signature.*`) take a
`Uint8Array`, any other `ArrayBufferView` (a Node `Buffer`, a `DataView`) or an `ArrayBuffer`,
viewed rather than copied.

Text drawn on an editor page is embedded as the builder embeds it: `Pdf.Editor.read` takes the
builder's font options — `fonts`, `preferSystemFonts`, `textLanguage`,
`disableFontAutoDiscovery` and `onWarning` — with the same meaning. Without them, characters
the standard faces cannot encode (Chinese, say) are drawn with a system face found on the host,
and `onWarning` reports that choice, and any character no face could draw:

```typescript
const editor = Pdf.Editor.read(existingPdfBytes, {
  fonts: { default: { regular: notoSansScBytes } },
  onWarning: message => console.warn(message)
});
editor.getPage(0).drawText("机密", { x: 72, y: 72 });
```

A full rewrite (`toBuffer()`, a removed page, each part of `splitPages()`) keeps the
document-level entries of the catalog — bookmarks, named destinations, page labels, the open
action, viewer preferences, optional content, output intents, attachments, the language — and
adjusts the ones that name pages to the pages written:

| Entry                                        | After a full rewrite                                                                                                                                       |
| -------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `/Outlines` (bookmarks)                      | A bookmark to a removed page loses its destination; it stays while a bookmark beneath it survives, and is dropped otherwise. Counts and links are rebuilt. |
| `/Dests`, `/Names /Dests`                    | Named destinations to removed pages are dropped.                                                                                                           |
| `/PageLabels`                                | Every page keeps its label; an added page continues the numbering before it.                                                                               |
| `/OpenAction`                                | Dropped when it goes to a removed page.                                                                                                                    |
| `/ViewerPreferences`                         | Kept; `/PrintPageRange` is renumbered, and dropped when none of its pages is left.                                                                         |
| `/Threads`                                   | Beads on removed pages are removed and the rest relinked; a thread with no bead left is dropped. Each page keeps its `/B`.                                 |
| `/StructTreeRoot`, `/MarkInfo`, `/DPartRoot` | Kept while no page is removed, dropped when one is.                                                                                                        |
| `/Perms`, `/Legal`                           | Dropped: they concern signatures, which a rewrite invalidates.                                                                                             |
| Everything else                              | Copied as it is.                                                                                                                                           |

A part of `splitPages()` keeps only the bookmarks that lead to its page, with their ancestors.

Cross-reference chains: the update's `/Prev` is the `startxref` offset the reader
actually parsed. Files whose latest section is an xref stream get an xref-stream
update; hybrid-reference files (classic trailer with `/XRefStm`) get a classic
section chained to that trailer, so objects listed only in the `/XRefStm` stream stay
reachable (ISO 32000-1 §7.5.8.4). If the cross-reference data had to be
reconstructed (`PdfDocument.xrefRecovered` — e.g. a wrong or damaged `startxref`, or an
entry that does not point at the object it names),
there is no reliable section to chain onto, so `toBuffer({ incremental: true })` falls back to a
full `toBuffer()` — subject to the same signature guard.

To make any fallback an error instead, pass `{ incremental: "required" }`: it appends
exactly as `true` does when it can, and otherwise throws `PdfIncrementalUpdateError`
with the reason, writing nothing.

Filling a field also sets the AcroForm's `/NeedAppearances true` (even where the file
said `false`). In a **hybrid** form — AcroForm fields plus an XFA packet — it **removes
`/XFA`**, so XFA-aware viewers show the values the AcroForm now holds rather than the
stale XFA data. A **dynamic** XFA form (`/NeedsRendering true`, or XFA with no AcroForm
field tree) is defined only by its XFA, so it is never stripped: saving edited values
into one throws `PdfXfaFormError` and writes nothing. A save with no edited values keeps
`/XFA` and `/NeedsRendering` unchanged.

`setFormField` and `setFormFields` take a field the document has — as `getFormFields()`
lists it — and throw `PdfArgumentError` for one it does not, or for a checkbox or radio
state its widgets cannot show (`"Off"` or the button's own value). A button is set by its
appearance state and a name value, so it is drawn checked, not only reported so. A field
added through the editor must have a name of its own: one the document already uses, or
two added fields with one name, make the write throw `PdfArgumentError`. The builder
applies the same rule across its pages.

### Digital Signatures

```typescript
import { Pdf } from "documonster/pdf";

// Verify every signature in a file
for (const sig of await Pdf.verifySignatures(pdfBytes)) {
  console.log(sig.fieldName, sig.claimedSigner, sig.intact, sig.coversWholeFile);
}

// Verify one signature whose /Contents and /ByteRange you already have
const result = await Pdf.Signature.verify(pdfBytes, signatureHex, byteRange);

// Complete a signature placeholder already in the file (DER certificate + PKCS#8 key)
const signed = await Pdf.Signature.fillPlaceholder(pdfWithPlaceholder, { certificate, privateKey });
```

---

## Architecture

The PDF module is split into four layers:

```
src/modules/pdf/
├── core/               # PDF primitives (objects, streams, writer, encryption, digital signatures)
├── font/               # TTF parsing, glyph metrics, font subsetting, embedding
├── render/             # Layout engine, page renderer, style converter
│   ├── layout-engine   — PdfSheetData → LayoutPage[] (zero @excel imports)
│   ├── page-renderer   — LayoutPage → PDF content stream (zero @excel imports)
│   ├── style-converter — PdfCellStyle → PDF rendering params (zero @excel imports)
│   ├── png-decoder     — PNG image decoding for PDF embedding (zero @excel imports)
│   └── pdf-exporter    — PdfWorkbook → Uint8Array (zero @excel imports)
├── builder/            # Free-form PDF creation and editing
│   ├── document-builder — PdfDocumentBuilder + PdfPageBuilder (text, shapes, SVG, annotations, forms)
│   ├── pdf-editor      — PdfEditor + PdfEditorPage (overlay, merge, split, sign)
│   ├── form-appearance — Form field appearance stream generation
│   ├── resource-merger — Resource dictionary merge for overlays
│   └── image-utils     — Shared image XObject writing
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
│   ├── form-extractor  — AcroForm: text, checkbox, radio, dropdown, listbox, signature
│   ├── bookmark-extractor — Nested outline tree extraction
│   ├── table-extractor — Heuristic table detection from text positions
│   ├── metadata-reader — Info dict + XMP metadata
│   ├── reader-utils    — shared reader utility functions
│   └── pdf-reader      — public API: Pdf.extract()
├── types.ts            # PdfWorkbook, PdfSheetData, PdfCellData, etc.
├── excel-bridge.ts     # Excel Workbook → PdfWorkbook conversion (ONLY @excel dependency)
└── index.ts
```

The entire PDF engine (core, font, render, reader) has **zero imports from the Excel module**. The `excel-bridge.ts` is the only file that knows about Excel — it converts `Workbook` to `PdfWorkbook`.

**Write strategy:** Write PDF 2.0 only (modern, AES-256).
**Read strategy:** Read all major PDF versions (1.0 through 2.0, all encryption types).

---

## Writer Options

`Pdf.fromTable()` accepts `PdfExportOptions`. `Pdf.fromExcel()` accepts
`PdfFromExcelOptions`, which extends the same writer options with an optional,
strongly typed formula recalculation hook:

```typescript
import { ExcelFormula } from "documonster/excel/formula";
import { Pdf, type PdfFromExcelOptions } from "documonster/pdf";

const options: PdfFromExcelOptions = {
  recalculate: ExcelFormula.calculate,
  fitToPage: true
};
const bytes = await Pdf.fromExcel(workbook, options);
```

Omit `recalculate` to render the workbook's existing cached formula results.
The separate callback keeps the formula engine out of bundles that do not opt
in.

```typescript
interface PdfExportOptions {
  // Page layout
  pageSize?: PdfPageSizeName | PdfPageSize; // "A4", "LETTER", "A3", etc. or { width, height }
  orientation?: "portrait" | "landscape";
  margins?: Partial<PdfMargins>; // { top, right, bottom, left } in points (72pt = 1in)
  horizontalCentered?: boolean; // Center the grid horizontally (default: sheet's pageSetup)
  verticalCentered?: boolean; // Center the grid vertically (default: sheet's pageSetup)
  pageOrder?: "downThenOver" | "overThenDown"; // Multi-page traversal (default: sheet's pageSetup, else "downThenOver")

  // Scaling — see "Scaling" below.
  fitToPage?: boolean; // Shrink to one page wide when nothing else sets a scale (default: true)
  scale?: number; // Explicit factor, 0.1–4.0 (matches Excel's 10–400%; may enlarge)
  fitToWidth?: number; // Shrink to at most N pages wide; 0 = unconstrained
  fitToHeight?: number; // Shrink to at most M pages tall; 0 = unconstrained

  // Content
  showGridLines?: boolean; // Render cell grid lines
  gridLineColor?: string; // ARGB color for grid lines (e.g. "FF3366CC")
  showRowColHeaders?: boolean; // Print row numbers and column letters (default: sheet's pageSetup)
  repeatRows?: number | false; // Header rows to repeat on each page (default: sheet's printTitlesRow)
  repeatCols?: number | false; // Leading columns to repeat on each page (default: sheet's printTitlesColumn)
  blackAndWhite?: boolean; // Grayscale vector content (default: sheet's pageSetup)
  draft?: boolean; // Omit images and charts (default: sheet's pageSetup)
  errors?: "displayed" | "blank" | "dash" | "NA"; // How error cells print (default: sheet's pageSetup)
  cellComments?: "none" | "atEnd" | "asDisplayed"; // Print comments (default: sheet's pageSetup)
  sheets?: (string | number)[]; // Sheet names; numeric positions are deprecated (see below)
  ignorePrintArea?: boolean; // Export the full used range, ignoring each sheet's print area (default: false)

  // Headers & footers
  showSheetNames?: boolean; // Show sheet name at top of each page
  showPageNumbers?: boolean; // Show "Page X of Y" at bottom of each page

  // Metadata — shared by every converter (PdfDocumentOptions, below)
  title?: string; // Default: the workbook's title
  author?: string; // Default: the workbook's creator
  subject?: string; // Default: the workbook's subject
  keywords?: string; // Default: the workbook's keywords
  creator?: string; // /Creator (default: "documonster")
  creationDate?: Date; // Fix it for byte-reproducible output (default: the time of export)
  modDate?: Date;

  // Font
  fonts?: PdfFontConfig; // Named families, styled faces, TTC selection, and fallback
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

### Options every converter shares

`Pdf.fromTable`, `Pdf.fromExcel`, `Pdf.fromDocx` and `Pdf.fromChart` all take
`PdfDocumentOptions` — `title`, `author`, `subject`, `keywords`, `creator`, `creationDate`,
`modDate`, `encryption`, `fonts`, `preferSystemFonts`, `disableFontAutoDiscovery`,
`textLanguage` and `onWarning` — with the same names, meanings and checks. So a Word
document or a single chart can be encrypted, carry metadata, and come out byte-for-byte
reproducible (fix `creationDate`) just like a workbook:

```typescript
import { Pdf, type PdfDocumentOptions } from "documonster/pdf";

const common: PdfDocumentOptions = {
  author: "Finance",
  keywords: "q3, report",
  creationDate: new Date("2026-01-01T00:00:00Z"),
  encryption: { ownerPassword: "owner", userPassword: "open" },
  disableFontAutoDiscovery: true
};
await Pdf.fromExcel(workbook, common);
await Pdf.fromDocx(doc, { ...common, margins: { top: 54 }, defaultFontFamily: "Times-Roman" });
await Pdf.fromChart(chart, { ...common, title: "Revenue" });
```

Each metadata field `Pdf.fromExcel` leaves out comes from the workbook's own properties, and
each `Pdf.fromDocx` leaves out from the document's core properties (`title`, `creator`,
`subject`, `keywords`). `Pdf.fromDocx` takes its margins as one `margins: { top, right,
bottom, left }` object and its default face as `defaultFontFamily`, as `Pdf.fromExcel` does.
`Pdf.fromChart`'s `title` is the PDF's title: the chart keeps drawing its own on every path.

Select sheets by **name** (`sheets: ["Summary", "Q1"]`, compared ignoring case; an unknown
name throws `PdfArgumentError`). Numeric positions are deprecated (removed in 1.0): they are
1-based here but 0-based in `WordExcel.toDocx` and sheet ids in `Workbook.getWorksheet`, so the
same number picks a different sheet in each. They still work exactly as before; a call
passing a number is struck through in an editor, while options typed as the published
`PdfFromExcelOptions` are not (they may hold only names).

## Reader Options

```typescript
interface PdfExtractOptions {
  password?: string; // Password for encrypted PDFs (user or owner)
  pages?: number[]; // Which pages to extract (1-based). Omit for all pages
  extractText?: boolean; // Extract text (default: true)
  extractImages?: boolean; // Extract images (default: true)
  extractMetadata?: boolean; // Extract metadata (default: true)
  extractAnnotations?: boolean; // Extract annotations (default: true)
  extractFormFields?: boolean; // Extract form fields (default: true)
  extractBookmarks?: boolean; // Extract bookmarks/outlines (default: true)
  extractTables?: boolean; // Extract tables via heuristics (default: false)
  maxDecodedBytes?: number; // Per-stream decode ceiling (default: 256 MiB)
}
```

Hitting `maxDecodedBytes`, the /Filter chain limit (16) or the object nesting limit throws
`PdfLimitExceededError` (a `PdfStructureError`) from `Pdf.extract`; it is never
downgraded to a page warning, so a result is never silently truncated. Its `limit` says
which one fired (`"maxDecodedBytes"`, `"filterChain"` or `"depth"`) and `allowed` its value.
`maxDecodedBytes` itself must be a number `>= 0` (or `Infinity`); anything else is refused
with `PdfArgumentError` (field `options.maxDecodedBytes`).

`Pdf.extract` and `Pdf.verifySignatures` read a `Blob` — a browser `File` — as well as bytes
(`Uint8Array`, another `ArrayBufferView`, an `ArrayBuffer`):

```typescript
const result = await Pdf.extract(input.files[0]);
```

### Reader Result

```typescript
interface PdfExtractResult {
  text: string; // All text from all pages
  pages: PdfExtractedPage[]; // Per-page results
  metadata: PdfMetadata; // Document metadata
  formFields: PdfFormField[]; // Form fields (document-level)
  bookmarks: PdfBookmark[]; // Document outline / TOC
}

interface PdfExtractedPage {
  pageNumber: number; // 1-based
  text: string; // Page text
  textLines: PdfTextLine[]; // Structured lines with positions
  textFragments: PdfTextFragment[]; // Raw fragments with exact coordinates
  images: PdfExtractedImage[]; // Extracted images
  annotations: PdfAnnotation[]; // Annotations (links, comments, highlights)
  width: number; // Page width in points
  height: number; // Page height in points
  warnings: string[]; // Non-fatal extraction warnings
}

interface PdfAnnotation {
  subtype: string; // "Link", "Text", "Highlight", "FreeText", "Stamp", etc.
  rect: PdfAnnotationRect; // Bounding rectangle { x1, y1, x2, y2 }
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
  // Text fields hold "" when the document does not give them.
  title: string;
  author: string;
  subject: string;
  keywords: string;
  creator: string;
  producer: string;
  // Absent when the document does not give them (or a date is not a date).
  creationDate?: Date;
  modDate?: Date;
  pageSize?: { width: number; height: number }; // First page, in points
  xmpXml?: string; // Raw XMP packet
  pdfVersion: string;
  pageCount: number;
  encrypted: boolean;
  custom: Record<string, string>; // Other Info-dictionary entries
}
```

### Page Sizes

Built-in page sizes accessible via `Pdf.PageSizes`:

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
- Content that is narrower or shorter than the page starts at the top-left margin — same as Excel

### Scaling

Excel offers two mutually exclusive scaling modes, and both are honored:

| Excel Page Setup                    | Sheet fields                                    | Export option                 |
| ----------------------------------- | ----------------------------------------------- | ----------------------------- |
| **Adjust to N %**                   | `scale`, `fitToPage: false`                     | `scale` (as a 0.1–4.0 factor) |
| **Fit to N page(s) wide by M tall** | `fitToWidth` / `fitToHeight`, `fitToPage: true` | `fitToWidth` / `fitToHeight`  |

```typescript
// Explicit
await Pdf.fromExcel(workbook, { fitToWidth: 1, fitToHeight: 0 });
await Pdf.fromExcel(workbook, { scale: 0.8 });

// Or let the sheet decide
worksheet.pageSetup.fitToPage = true;
worksheet.pageSetup.fitToWidth = 1;
```

**`fitToWidth` / `fitToHeight` only ever shrink**, like Excel: a grid smaller
than the target is printed at actual size. `scale` is different — it is a plain
multiplier and _can_ enlarge (Excel allows 10–400%).

`fitToPage` (default `true`) is documonster's own fallback — "shrink to one page
wide" — and applies only when neither the caller nor the sheet expresses a
scaling intent, so a sheet asking for 80% is never shrunk twice. Note that
`scale` does not disable it: `{ scale: 2 }` still fits one page wide, resolving
to `min(scale, contentWidth / tableWidth)`. Pass `fitToPage: false` for a plain
multiplier.

Both fit constraints are solved against the real paginator by bisection, so
indivisible columns/rows, manual page breaks and repeated title bands are all
accounted for — a total-area ratio alone would under-shrink (three columns at
60% of the page width occupy "1.8 pages" by area but still need three).

> **The page target is best-effort within Excel's scaling range.** Scaling stops
> at Excel's 10% floor, so a target that is unreachable there — manual page
> breaks alone can force more pages than requested — yields more than N pages
> rather than an illegible grid.
>
> Row heights are measured once unscaled and then scaled linearly while
> bisecting. That is exact for every cell, wrapped ones included: the wrap
> calculation scales the column width, padding and font size together, so the
> number of wrapped lines does not depend on the print scale.

### Row and Column Headings

Mirrors Excel's **Sheet → Row and column headings**. The bands keep a fixed size
regardless of the print scale so they stay legible on shrunken output.

```typescript
await Pdf.fromExcel(workbook, { showRowColHeaders: true });
worksheet.pageSetup.showRowColHeaders = true; // or via the sheet
```

### Black and White, Draft, Cell Errors

```typescript
await Pdf.fromExcel(workbook, {
  blackAndWhite: true, // grayscale vector content (see note below)
  draft: true, // omit images and charts
  errors: "dash" // print "--" instead of #DIV/0! etc.
});
```

`blackAndWhite` converts colors to their luminance-preserving grayscale
equivalent, so relative lightness — and therefore contrast — is kept rather than
flattening everything to black. Opacity is preserved.

Vector content — cell text, fills and borders, gridlines, the heading bands,
chart vectors, `&K`-colored header/footer runs, text watermarks — has its colors
converted up-front.

**Raster content is converted too**, per pixel rather than by painting over it:

- **PNG** is already decoded when it is embedded, so its RGB samples collapse to
  one `/DeviceGray` luma component. Alpha lives in a separate `SMask` and is
  untouched, so transparent regions stay transparent.
- **JPEG** keeps its `DCTDecode` data — no decoder required — and is
  reinterpreted through a `[/DeviceN [...] /DeviceGray <luma>]` color space whose
  tint transform is the Rec. 601 formula.

The tempting shortcut of a `/BM /Saturation` overlay is deliberately _not_ used:
its source is opaque black, so it paints transparent PNG regions solid black
instead of filtering them.

`draft` omits images and charts. A chartsheet still emits its page — blank — so
page numbering is unaffected, as in Excel.

`errors` accepts `"displayed"` (default), `"blank"`, `"dash"` and `"NA"`, and
applies both to plain error cells and to formulas whose result is an error.

### Cell Comments

Mirrors Excel's **Sheet → Comments and notes**. Off by default; both of Excel's
modes are supported:

```typescript
// A list after the sheet's pages, each entry addressed back to its cell
await Pdf.fromExcel(workbook, { cellComments: "atEnd" });

// Boxes drawn where they sit on the sheet, with Excel's red corner marker
await Pdf.fromExcel(workbook, { cellComments: "asDisplayed" });

worksheet.pageSetup.cellComments = "atEnd"; // or via the sheet
```

Classic notes and threaded comments are both included. `asDisplayed` decodes the
note's VML anchor to place the box, falling back to Excel's default offset from
the commented cell when a note has no anchor; a box whose anchor lies outside the
current page is skipped rather than sliced across the page seam.

### Center on Page

Mirrors Excel's **Page Setup → Margins → Center on page**. Off by default, so a
narrow sheet stays anchored to the left/top margin:

```typescript
await Pdf.fromExcel(workbook, { horizontalCentered: true, verticalCentered: true });
```

Or via worksheet page setup (`<printOptions horizontalCentered="1"/>` in the XLSX):

```typescript
worksheet.pageSetup.horizontalCentered = true;
worksheet.pageSetup.verticalCentered = true;
```

An explicit export option always wins over the worksheet setting, and options
are resolved per sheet, so sheets in one workbook can differ.

### Repeat Header Rows and Columns

```typescript
await Pdf.fromExcel(workbook, { repeatRows: 2, repeatCols: 1 });
```

Or via worksheet page setup:

```typescript
worksheet.pageSetup.printTitlesRow = "1:2"; // Repeat rows 1-2 on every page
worksheet.pageSetup.printTitlesColumn = "A"; // Repeat column A on every page
```

Pass `false` explicitly to suppress a sheet's print titles.

Print titles are **absolute** and independent of the print area, as in Excel: a
sheet printing `E1:T3` still repeats `A:B` down the left of every page. A band
that lies _inside_ the printed range still repeats on every page after the first,
but is not hoisted to the left edge — that would reorder the first page's grid.

### Page Order

Mirrors Excel's **Page Setup → Sheet → Page order** for sheets that paginate in
both directions. The default, like Excel's, is `downThenOver`: each column band
is printed top to bottom before moving right.

```typescript
await Pdf.fromExcel(workbook, { pageOrder: "overThenDown" });
```

### Manual Page Breaks

```typescript
import { Column, Row } from "documonster/excel";

Row.addPageBreak(worksheet, 20); // page 2 starts at row 21
Column.addPageBreak(worksheet, "F"); // the next page starts at column G
```

A break spans the full width or height of the sheet — the only kind Excel can
author, and the only kind this exporter renders.

### Print Area

```typescript
worksheet.pageSetup.printArea = "A1:F50"; // Export only this range
```

> **Note:** If a multi-range print area is set (e.g. `"A1:B5&&D1:E10"`), only the first range is used for PDF export.

To export the full used range and ignore any print area without modifying the workbook, pass `ignorePrintArea`:

```typescript
await Pdf.fromExcel(workbook, { ignorePrintArea: true });
```

---

## Images

JPEG and PNG images are embedded when sheets contain images:

```typescript
import { Workbook, Image } from "documonster/excel";
import { Pdf } from "documonster/pdf";

const imageId = Image.add(workbook, {
  buffer: jpegBytes,
  extension: "jpeg"
});

Image.place(worksheet, imageId, {
  tl: { col: 0, row: 0 },
  ext: { width: 200, height: 150 }
});

const pdf = await Pdf.fromExcel(workbook);
// Image appears in the PDF at the specified position
```

PNG transparency (RGBA and tRNS) is preserved via PDF soft masks.

---

## Encryption

The writer produces **AES-256 encrypted PDFs** (PDF 2.0, V=5, R=5). The reader can decrypt **all major encryption formats** including legacy RC4.

### Writer Encryption (AES-256)

#### Owner-Only (No Open Password)

```typescript
const pdf = await Pdf.fromExcel(workbook, {
  encryption: {
    ownerPassword: "admin",
    permissions: { print: true, copy: false, modify: false }
  }
});
// Opens without password, but copy/modify is restricted
```

#### Open Password Required

```typescript
const pdf = await Pdf.fromExcel(workbook, {
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
const result = await Pdf.extract(encryptedBytes, { password: "secret" });
```

---

## Font Configuration

Standard Type1 fonts (Helvetica, Times, Courier) cover a limited character set. Use `PdfFontConfig` to embed TrueType families for Unicode text, style faces, and deterministic fallback:

```typescript
import { readFileSync } from "node:fs";

import type { PdfFontConfig, PdfFontSource } from "documonster/pdf";

const notoCjk: PdfFontSource = {
  data: readFileSync("NotoSansCJK.ttc"),
  collectionIndex: 2 // zero-based face index; defaults to 0
};

const fonts: PdfFontConfig = {
  // Used when no requested family name or alias matches. regular is required;
  // omitted bold/italic/boldItalic slots fall back to this regular face.
  default: {
    regular: readFileSync("NotoSans-Regular.ttf"),
    bold: readFileSync("NotoSans-Bold.ttf"),
    italic: readFileSync("NotoSans-Italic.ttf"),
    boldItalic: readFileSync("NotoSans-BoldItalic.ttf")
  },
  families: [
    {
      name: "Noto Sans CJK SC",
      aliases: ["Microsoft YaHei", "SimSun"],
      faces: { regular: notoCjk }
    }
  ],
  // Ordered names or aliases from families. Only listed families participate
  // in missing-glyph fallback; default is tried last.
  fallbackFamilies: ["Noto Sans CJK SC"]
};
```

Family names and aliases are matched case-insensitively. Fallback chooses one face for each Unicode grapheme cluster, so a grapheme is never split across fonts. Fonts are subsetted to the sequences used in the output.

The same configuration is accepted by every writing bridge and the builder:

```typescript
// Standalone data or Excel workbook
const tablePdf = await Pdf.fromTable(rows, { fonts });
const workbookPdf = await Pdf.fromExcel(workbook, { fonts });

// Word document; the same fonts drive layout measurement and PDF rendering
const wordPdf = await Pdf.fromDocx(docxDocument, { fonts });

// Standalone Excel chart; fonts apply to selectable vector-chart text
const chartPdf = await Pdf.fromChart(chart, { fonts }); // chartToPdf bridge

// Builder: configure in the constructor or with embedFonts().
const builder = new Pdf.Builder({ fonts });
const page = builder.addPage();
page.drawText("Report", { x: 72, y: 760, fontFamily: "Noto Sans CJK SC" });
```

`Builder.embedFont(source)` is the single-face shortcut: it is `embedFonts({ default: { regular: source } })`, where `source` is the font bytes or `{ data, collectionIndex }` to pick a face of a `.ttc`, and it replaces any prior Builder font configuration. `embedFont()` and `embedFonts()` must run before the first text command across all pages. Calling `addPage()` first is allowed if no text has yet been drawn.

If no configured face covers a character, the PDF displays that face's `.notdef` glyph, but retains the original Unicode sequence in ToUnicode for copy, search, accessibility, and extraction. One embedded face can encode at most 65,535 distinct Unicode sequences per PDF.

Because that degradation is visual only, pass `onWarning` to detect it. Every
bridge and the builder accept it, and it reports the uncovered code points so a
fallback family can be added:

```typescript
await Pdf.fromExcel(workbook, {
  fonts,
  onWarning: message => console.warn(message)
});
```

### System font auto-discovery

On Node, a document that contains characters WinAnsi cannot encode — CJK, most
notably — and that supplies neither `font` nor `fonts` triggers a best-effort
scan of the host's font directories for a face to borrow glyphs for those
characters from. Latin text keeps its standard-14 face, so bold and italic
survive. Nothing is scanned in the browser, or when a font is supplied
explicitly.

The scan prefers the platform's own CJK faces and treats broad-coverage
catch-alls such as Arial Unicode MS as a last resort: it covers almost every
script in one file, but its CJK glyphs sit small on the em and read as visibly
worse than any system CJK face. Collections are searched face by face, since the
faces of one `.ttc` differ in both coverage and regional conventions — macOS
`Songti.ttc` holds 8,535 glyphs in face 0 and 43,033 in face 1.

`preferSystemFonts` names the families you would rather have, in order:

```typescript
// Simplified Chinese punctuation conventions; the Traditional face of the same
// .ttc would be "Heiti TC".
await Pdf.fromExcel(workbook, { preferSystemFonts: ["Heiti SC", "Songti SC"] });

new Pdf.Builder().preferSystemFonts(["Microsoft YaHei", "Noto Sans CJK SC"]);
```

Names match the font's family name case-insensitively, and are also how a
specific face inside a collection is reached. A family that is not installed,
cannot be parsed, or does not cover the document's text is skipped and the
built-in order applies — this steers a best-effort search rather than
constraining it. Use `fonts` when a face is a requirement.

### East Asian language and glyph forms

Unicode Han Unification gives Chinese, Japanese and Korean the same code points
for the characters they share, but not the same shapes — 「者」「骨」「今」「青」
「每」are each drawn differently by region. A font chosen purely by _coverage_ can
therefore be correct and still wrong: a Japanese face draws Chinese text that a
Chinese reader sees as malformed.

Selection is therefore by language first, coverage second:

```typescript
await Pdf.fromExcel(workbook, { textLanguage: "zh-Hans" }); // or zh-Hant, ja, ko
new Pdf.Builder().setTextLanguage("zh-Hant");
```

Left unset, the language is inferred from the content: a single kana settles
Japanese, Hangul settles Korean, and characters that exist in only one of
Simplified or Traditional Chinese settle those. Text made purely of forms common
to all of CJK carries no evidence, and Simplified Chinese families are preferred
in that case — a default rather than a detection, and the reason to state the
language when you know it.

`Pdf.fromChart` accepts both options too, so a standalone chart's labels get the
same regional control as a workbook or a document.

Where a collection holds several weights of one family — macOS `Songti.ttc` lists
`Songti SC` at Black before its Regular — the regular weight is chosen rather than
whichever face the file happens to list first.

The language also picks between the faces of one collection: macOS
`STHeiti Light.ttc` holds `Heiti TC` at face 0 and `Heiti SC` at face 1, and
`Songti.ttc` holds eight faces across both scripts. `preferSystemFonts` still
outranks the language when you name a family yourself.

Two limits are worth knowing. CFF-flavoured fonts (`.otf`, and `.ttc` files
whose faces are CFF) are rejected because the subsetting embedder needs `glyf`
outlines — which rules out macOS PingFang and Hiragino, and the official Noto
Sans CJK `.otf`/`.ttc` releases; use Noto Sans SC's `.ttf` build instead. And
because results depend on what the host has installed, `disableFontAutoDiscovery`
turns the scan off entirely when byte-stable output matters more than legibility.

`disableFontAutoDiscovery` is available on every entry point, not just the
builder — the same export is otherwise reproducible or not depending on which
function produced it:

```typescript
import { Pdf } from "documonster/pdf";

// Deterministic: either the supplied font draws the text or nothing does.
const bytes = await Pdf.fromExcel(workbook, {
  disableFontAutoDiscovery: true,
  fonts: { default: { regular: myFontBytes } }
});
```

It disables _discovery_, not embedding, so a font you supply is still used. With
no font supplied the characters keep their Unicode for copy and search and render
as `.notdef` boxes, and `onWarning` names the blocks that were lost — turning the
scan off does not turn the diagnostic off with it.

Font fallback is scalar rendering, not complex text layout. OpenType shaping (GSUB/GPOS), bidi reordering, and color emoji are **not supported** — configuring a font does not make shaping-dependent Arabic or Indic text render correctly.

These limits are detectable rather than silent: when the document contains a complex script, right-to-left text, or an embedded font carrying color glyph tables (`COLR`/`CBDT`/`sbix`/`SVG`), `onWarning` reports it once per feature, naming the scripts or font families involved. Pre-shape and pre-order such text (or render it as an image) before drawing it.

---

## Per-Sheet Page Setup

Each sheet's `pageSetup` is respected when using the Excel bridge:

```typescript
import { Workbook } from "documonster/excel";
import { Pdf } from "documonster/pdf";

const ws1 = Workbook.addWorksheet(workbook, "Summary");
ws1.pageSetup.paperSize = 9; // A4
ws1.pageSetup.orientation = "portrait";

const ws2 = Workbook.addWorksheet(workbook, "Data");
ws2.pageSetup.paperSize = 1; // Letter
ws2.pageSetup.orientation = "landscape";

// Each sheet renders with its own page size/orientation
const pdf = await Pdf.fromExcel(workbook);
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
import { Workbook } from "documonster/excel";
```

```typescript
// Imports Excel + PDF bridge
import { Workbook } from "documonster/excel";
import { Pdf } from "documonster/pdf";
```

---

## Examples

Runnable examples are in `src/modules/pdf/examples/`:

| File                   | What it demonstrates                                                                                         |
| ---------------------- | ------------------------------------------------------------------------------------------------------------ |
| `pdf-basic.ts`         | Page sizes, margins, metadata, sheet selection, scale                                                        |
| `pdf-styled.ts`        | Fonts, fills, borders, alignment, merge, rotation, rich text, number formats                                 |
| `pdf-advanced.ts`      | Pagination, page breaks, encryption, transparency, bookmarks, hidden rows/cols                               |
| `pdf-excel-to-pdf.ts`  | Reading real `.xlsx` files and converting to PDF                                                             |
| `pdf-images.ts`        | Image embedding (JPEG, PNG with transparency)                                                                |
| `pdf-reader.ts`        | Text extraction, metadata, images, encrypted PDFs, selective extraction                                      |
| `pdf-reader-stress.ts` | Large-scale stress test: thousands of cells, encrypted roundtrip, benchmarks                                 |
| `pdf-builder.ts`       | PdfDocumentBuilder, PdfEditor, annotations, forms, SVG paths, bookmarks, TOC, PDF/A, merge, incremental save |
| `pdf-signatures.ts`    | Digital signature placeholders, ASN.1 parsing, signature verification                                        |

Run any example:

```bash
pnpm example --filter pdf-basic
# Output: tmp/pdf-examples/*.pdf

pnpm example --filter pdf-builder
# Output: tmp/pdf-builder-examples/*.pdf

pnpm example --filter pdf-signatures
# Output: tmp/pdf-signature-examples/
```

---

## API Reference

### `Pdf.extract(data, options?)`

Read a PDF file and extract text, images, and metadata. Returns `Promise<PdfExtractResult>`.

```typescript
import { Pdf } from "documonster/pdf";

// Basic
const result = await Pdf.extract(pdfBytes);
console.log(result.text);
console.log(result.pages[0].images);
console.log(result.pages[0].annotations);
console.log(result.formFields);
console.log(result.metadata);

// Encrypted
const result = await Pdf.extract(pdfBytes, { password: "secret" });

// Selective
const result = await Pdf.extract(pdfBytes, {
  pages: [1, 3],
  extractImages: false,
  extractMetadata: false
});
```

### `Pdf.fromTable(input, options?)`

Generate a PDF from plain data. Returns `Promise<Uint8Array>`.

```typescript
// 2D array
await Pdf.fromTable([
  ["Name", "Age"],
  ["Alice", 30]
]);

// Single sheet with column widths
await Pdf.fromTable({ name: "Report", columns: [{ width: 25 }, 15], data: [["A", "B"]] });

// Multiple sheets
await Pdf.fromTable({
  sheets: [
    { name: "S1", data: [["A"]] },
    { name: "S2", data: [["B"]] }
  ]
});

// With options
await Pdf.fromTable([["A", 1]], { showGridLines: true, pageSize: "A4" });
```

### `Pdf.fromExcel(workbook, options?)`

Convert an Excel `Workbook` to PDF. Returns `Promise<Uint8Array>`.

```typescript
import { Workbook } from "documonster/excel";
import { Pdf } from "documonster/pdf";

const workbook = Workbook.create();
// ... build workbook ...
const bytes = await Pdf.fromExcel(workbook, { showGridLines: true });
```

### `Pdf.Builder`

Build free-form PDFs with text, vector graphics, annotations, and form fields.

```text
import { Pdf } from "documonster/pdf";

const doc = new Pdf.Builder();
doc.setMetadata({ title, author, subject, keywords, creator, producer });
doc.setEncryption({ ownerPassword, userPassword?, permissions? });
doc.setPdfACompliance();       // Enable PDF/A-1b
doc.embedFont(fontBytes);      // One default face: bytes, or { data, collectionIndex } for a .ttc
// Or: doc.embedFonts(fonts);  // Complete PdfFontConfig; the later call replaces earlier config

const page = doc.addPage({ width?, height? }); // Returns Pdf.PageBuilder

// PdfPageBuilder methods:
page.drawText(text, { x, y, fontSize?, fontFamily?, bold?, italic?, color? });
page.drawRect({ x, y, width, height, fill?, stroke?, lineWidth? });
page.drawCircle({ cx, cy, r, fill?, stroke? });
page.drawEllipse({ cx, cy, rx, ry, fill?, stroke? });
page.drawLine({ x1, y1, x2, y2, color?, lineWidth? });
page.drawPath(ops, { fill?, stroke?, lineWidth? });
page.drawSvgPath(d, { fill?, stroke?, lineWidth? });
page.drawImage({ x, y, width, height, data, format });
page.addAnnotation({ type, rect, ...options });
page.addFormField({ type, name, rect, ...options });

doc.addBookmark(title, pageIndex, parent?);
doc.generateTableOfContents({ title?, fontSize?, indent? });

const bytes = await doc.toBuffer(); // Returns Promise<Uint8Array>
```

### `Pdf.Editor`

Edit existing PDFs — overlay content, fill forms, merge, split, and sign.

```text
const editor = Pdf.Editor.read(pdfBytes, { password? });

// Page access
const page = editor.getPage(index);    // Returns PdfEditorPage
const count = editor.pageCount;         // pages as read (getter); outputPageCount: pages written

// PdfEditorPage methods (same drawing API as PdfPageBuilder):
page.drawText(text, options);
page.drawRect(options);
page.drawCircle(options);
page.drawLine(options);
page.drawImage(options);
page.drawSvgPath(d, options);
page.drawPath(ops, options);
page.addAnnotation(options);
page.addFormField(options);

// Page manipulation
editor.addPage(options?);              // Returns Pdf.PageBuilder
editor.removePage(index);             // any page, read or added; idempotent, others keep their index
editor.isPageRemoved(index);
editor.rotatePage(index, degrees);     // 90, 180, 270
editor.copyPagesFrom(otherPdfBytes, { pageIndices: [0, 2] });

// Form filling
editor.setFormField(name, value);
editor.setFormFields({ name: value, ... });

// Save
const full = await editor.toBuffer();             // Full rebuild
const incr = await editor.toBuffer({ incremental: true });  // Append-only
const pages = await editor.splitPages();      // One PDF per page of what toBuffer() would write
```

### `Pdf.verifySignatures(pdfData, options?)`

Find every signature in a PDF and verify each one. Returns
`Promise<PdfSignatureReport[]>` — one entry per signature field, `[]` for an unsigned
file. Signatures are located through the AcroForm field tree (a `/FT /Sig` field whose
value carries `/ByteRange` and `/Contents`), not by searching the bytes.

```typescript
import { Pdf } from "documonster/pdf";

const reports = await Pdf.verifySignatures(pdfBytes, { password: "optional" });
for (const r of reports) {
  // r.fieldName        — qualified field name, e.g. "Signature1"
  // r.intact           — digest matches and the signature verifies with the embedded
  //                      certificate the SignerInfo names (integrity, not trust)
  // r.byteRangeValid   — ranges start at 0, stay in the file, and enclose exactly /Contents
  // r.coversWholeFile  — false when bytes were appended after signing (an incremental update)
  // r.certificate      — the signer's certificate (DER), for your own trust evaluation
  // r.claimedSigner    — its common name — a claim until you trust the certificate
  // r.certificateValidity — { notBefore, notAfter } as the certificate declares them
  // r.name / r.reason / r.location / r.contactInfo / r.signingTime
  //                    — from the signature dictionary (unauthenticated)
  // r.failure          — why the signature is not intact (if !intact)
}
```

**No trust is evaluated.** `intact` does not build a certificate chain, check revocation,
honour a timestamp or judge the validity period — a self-signed certificate can carry any
name. Evaluate `r.certificate` against your own trust store before believing
`claimedSigner`.

A signature that is `intact` but not `coversWholeFile` signed an earlier revision: the
signed bytes are unchanged, and whatever follows them is unsigned. Whether that later
update was permitted by `/DocMDP` or `/FieldMDP` is not evaluated.

### `Pdf.Signature.verify(pdfData, signatureHex, byteRange)`

Verify a digital signature. Returns `Promise<PdfSignatureCheck>`, which uses the same
field names as a `Pdf.verifySignatures` report.

```typescript
import { Pdf } from "documonster/pdf";

const result = await Pdf.Signature.verify(pdfBytes, sigHex, [0, off1, off2, len2]);
// result.intact           — boolean
// result.coversWholeFile  — boolean (no unsigned gaps)
// result.digestAlgorithm  — OID string
// result.failure          — why it is not intact (if !intact)
```

### `Pdf.Signature.fillPlaceholder(pdfBytes, { certificate, privateKey })`

Complete the signature placeholder already in a PDF — the last step of signing only.
Returns `Promise<Uint8Array>`; throws `PdfSignaturePlaceholderMissingError` (a
`PdfStructureError`) when the file has no placeholder, and `PdfArgumentError` when the
certificate is not DER X.509 or the key is not DER PKCS#8 RSA — `Pdf.Builder#sign` and
`Pdf.Editor#sign` check the same when they are called. `Pdf.Builder#sign` and
`Pdf.Editor#sign` write and fill the placeholder in one call.

```typescript
import { Pdf } from "documonster/pdf";

// Step 1: Build placeholder
const { dictString, placeholder } = Pdf.Signature.buildDictPlaceholder({
  name: "Alice",
  reason: "Approval",
  location: "London",
  contactInfo: "alice@example.com"
});

// Step 2: Sign (certificate = DER X.509, privateKey = DER PKCS#8)
const signed = await Pdf.Signature.fillPlaceholder(pdfWithPlaceholder, { certificate, privateKey });
```

### `Pdf.createDrawSurface(page, rect, options?)`

The PDF backend for `documonster/draw`: draws a display list onto a page, its top-left
corner at `rect`'s top-left (PDF points; `y` is the bottom edge). `options.scale`
multiplies the list's coordinates. Nothing is clipped or fitted to `rect` — a content
stream has to balance its `q`/`Q` and a surface is never told the list has ended — so size
the rect as the list's size times the scale.

```typescript
import { Draw } from "documonster/draw";
import { Pdf } from "documonster/pdf";

const doc = new Pdf.Builder();
const page = doc.addPage();
Draw.render(list, Pdf.createDrawSurface(page, { x: 72, y: 500, width: 200, height: 100 }));
```

### `Pdf.parseSvgPath(d)`

Parse an SVG path `d` attribute into an array of `PdfPathOp` objects for `drawPath()`.

```typescript
import { Pdf } from "documonster/pdf";

const ops = Pdf.parseSvgPath("M10 10 L90 10 L50 80 Z");
page.drawPath(ops, { fill: { r: 1, g: 0, b: 0 } });
```

### Error Types

```typescript
import {
  PdfError, // Base class for all PDF errors
  PdfRenderError, // Layout/rendering failures
  PdfFontError, // Font parsing/embedding failures
  PdfStructureError, // PDF structure assembly failures
  PdfPasswordError, // A PdfStructureError: an encrypted file the password given does not open
  PdfLimitExceededError, // A PdfStructureError: a read hit a resource limit (.limit, .allowed)
  PdfSignaturePlaceholderMissingError, // A PdfStructureError: no signature placeholder to fill
  PdfArgumentError, // A builder/editor/converter argument with no valid form (.field names it)
  isPdfError // Type guard: (err: unknown) => err is PdfError
} from "documonster/pdf";
```

A `PdfArgumentError`'s message starts with the call as you wrote it (`Pdf.fromExcel: …`,
`Pdf.PageBuilder#drawText: …`, `Pdf.Editor#splitPages: …`), and its `field` is the path from
that call's parameter name: `options.fontSize`, `options.watermark.pages[0]`, `metadata.title`,
`data`.

All errors extend `BaseError` with `cause` chain support:

```typescript
import { isPdfError, Pdf } from "documonster/pdf";

try {
  await Pdf.fromExcel(workbook);
} catch (err) {
  if (isPdfError(err)) {
    console.error(err.message, err.cause);
  }
}
```
