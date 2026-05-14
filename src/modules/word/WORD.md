# excelts/word — DOCX Processing Module

Zero-dependency TypeScript library for reading, writing, and manipulating DOCX files.
Works in Node.js 22+ and modern browsers.

## Quick Start

```typescript
import { Document, toBuffer, readDocx } from "excelts/word";

// Create a document
const doc = Document.create();
Document.addHeading(doc, "Hello World", 1);
Document.addParagraph(doc, "This is a paragraph.");
Document.addTable(doc, [
  ["A", "B"],
  ["1", "2"]
]);

const buffer = await toBuffer(Document.build(doc));
// Write to file, send as response, etc.

// Read a document
const parsed = await readDocx(buffer);
console.log(parsed.body.length, "elements");
```

## Core API

### Document Builder

```typescript
import { Document, paragraph, text, bold, italic, heading } from "excelts/word";

const doc = Document.create();

// Text
Document.addParagraph(doc, "Simple text");
Document.addParagraphElement(doc, paragraph([text("Normal "), bold("bold "), italic("italic")]));
Document.addHeading(doc, "Title", 1);

// Tables
Document.addTable(
  doc,
  [
    ["Header1", "Header2"],
    ["cell1", "cell2"]
  ],
  { headerRow: true }
);

// Images
Document.addImage(doc, pngBytes, "png", width, height);
Document.addFloatingImage(doc, jpgBytes, "jpeg", width, height, { wrap: { style: "square" } });

// Lists
Document.addBulletList(doc, ["Item 1", "Item 2", "Item 3"]);
Document.addNumberedList(doc, ["First", "Second", "Third"]);

// Page layout
Document.setSectionProperties(doc, {
  pageSize: { width: 11906, height: 16838 }, // A4
  margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
});

// Headers/Footers
Document.setHeader(doc, "default", { children: [textParagraph("Page Header")] });

// Build & export
const model = Document.build(doc);
const bytes = await toBuffer(model);
```

### Reading Documents

```typescript
import { readDocx, extractText, searchText } from "excelts/word";

const doc = await readDocx(fileBuffer);

// Extract text content
const text = extractText(doc);

// Search
const results = searchText(doc, /pattern/g);
```

### Modifying Documents

```typescript
import { readDocx, replaceText, toBuffer } from "excelts/word";

const doc = await readDocx(buffer);
const modified = replaceText(doc, "OLD_TEXT", "NEW_TEXT");
const output = await toBuffer(modified);
```

## Advanced Features

### Template Engine

```typescript
import { fillTemplate } from "excelts/word";

const filled = fillTemplate(doc, {
  name: "John",
  showDetails: true,
  items: ["A", "B", "C"]
});
// Supports: {{variable}}, {{#if cond}}...{{/if}}, {{#each arr}}...{{/each}}
```

### Form Fields

```typescript
import { extractFormFields, fillFormFields, formTextField, formCheckboxField } from "excelts/word";

// Extract form data
const fields = extractFormFields(doc);
// → [{ name: "FullName", type: "text", value: "..." }, ...]

// Fill form data
const filled = fillFormFields(
  doc,
  new Map([
    ["FullName", "Jane Doe"],
    ["AgreeTerms", true],
    ["Country", 2]
  ])
);
```

### Data Binding (OpenDoPE)

```typescript
import { resolveDataBindings } from "excelts/word";

// Resolve SDT data bindings against CustomXML parts
const resolved = resolveDataBindings(doc);

// Or with override data
const resolved2 = resolveDataBindings(
  doc,
  new Map([["{GUID}", "<root><field>value</field></root>"]])
);
```

### Drawing Shapes with Effects

```typescript
import { createShape, createRect, createEllipse } from "excelts/word";

const shape = createShape({
  shapeType: "roundRect",
  width: 3000000, // EMU
  height: 2000000,
  fill: {
    type: "gradient",
    stops: [
      { position: 0, color: "FF0000" },
      { position: 100000, color: "0000FF" }
    ]
  },
  effects: {
    shadow: {
      type: "outer",
      color: "000000",
      blurRadius: 50800,
      distance: 38100,
      direction: 2700000
    },
    glow: { color: "FFFF00", radius: 101600 },
    reflection: { startOpacity: 50, endOpacity: 0, distance: 25400 },
    softEdges: 63500,
    effect3d: {
      camera: "perspectiveFront",
      bevelTop: { width: 127000, height: 63500, preset: "circle" }
    }
  }
});
```

### Font Embedding with Subsetting

```typescript
import { embedFont, addEmbeddedFonts, subsetFont } from "excelts/word";

// Embed with automatic subsetting (only glyphs used in document)
const result = embedFont({
  name: "CustomFont",
  data: fontFileBytes,
  style: "regular",
  usedCharacters: "Hello World" // Only these glyphs are embedded
});

// Add to document
const docWithFonts = addEmbeddedFonts(doc, [result]);
```

### Track Changes

```typescript
import { acceptAllRevisions, rejectAllRevisions } from "excelts/word";

const accepted = acceptAllRevisions(doc);
const rejected = rejectAllRevisions(doc);
```

### Document Diff

```typescript
import { diffDocuments } from "excelts/word";

const diff = diffDocuments(docA, docB);
// → { changes: [{ type: "added"|"removed"|"modified", ... }] }
```

### Document Merge

```typescript
import { mergeDocuments } from "excelts/word";

const merged = mergeDocuments([doc1, doc2, doc3], { sectionBreak: "nextPage" });
```

### Streaming Writer

```typescript
import { createDocxStream } from "excelts/word";

const stream = createDocxStream();
stream.addParagraph("Title", { style: "Heading1" });
for (const item of largeDataset) {
  stream.addParagraph(item.text);
}
const buffer = await stream.finalize();
```

### Document Protection

```typescript
import { protectDocument, isDocumentProtected, verifyProtectionPassword } from "excelts/word";

const protected = protectDocument(doc, { type: "readOnly", password: "secret" });
const isProtected = isDocumentProtected(protected); // true
const valid = verifyProtectionPassword(protected, "secret"); // true
```

### Validation

```typescript
import { validateDocument } from "excelts/word";

const result = validateDocument(doc);
if (!result.valid) {
  console.log(result.issues); // [{ severity, message, path }]
}
```

### HTML/Markdown Conversion

```typescript
// DOCX → Markdown
import { renderToMarkdown } from "excelts/word/markdown";
const md = renderToMarkdown(doc);

// DOCX → HTML
import { renderToHtml } from "excelts/word/html";
const html = renderToHtml(doc);

// HTML → DOCX body content
import { htmlToDocxBody } from "excelts/word/html";
const body = htmlToDocxBody("<h1>Hello</h1><p>World</p>");
```

### Flat OPC Format

```typescript
import { parseFlatOpc, toFlatOpc, isFlatOpc } from "excelts/word";

// Single-XML representation of a DOCX
const flatXml = toFlatOpc(doc);
const doc = parseFlatOpc(flatXmlString);
```

## OOXML Strict Format

The module automatically handles documents saved in ISO 29500 Strict conformance.
When reading a Strict-format .docx, namespace URIs and relationship types are
transparently normalized to their Transitional equivalents — no user action required.

## Compatibility

| Feature                | Support                                                        |
| ---------------------- | -------------------------------------------------------------- |
| .docx (read)           | ✅ Broad (common elements structured; unknown parts preserved) |
| .docx (write)          | ✅ Broad (common elements; unknown parts written as opaque)    |
| .dotx (template)       | ✅                                                             |
| .docm (macro)          | ✅ Round-trip (VBA preserved, not executed/edited)             |
| .dotm (macro template) | ✅ Round-trip                                                  |
| Flat OPC (.xml)        | ✅                                                             |
| ISO 29500 Strict       | ✅ Auto-normalized                                             |
| Encrypted .docx        | ✅ Decrypt with password (Agile Encryption)                    |
| Digital Signatures     | 🔍 Detection & metadata extraction (no signing/verification)   |
| Browser                | ✅ (import from "excelts/word/browser")                        |
| Node.js 22+            | ✅                                                             |

## Migration from `docx` (npm)

| docx (npm)                          | excelts/word                                             |
| ----------------------------------- | -------------------------------------------------------- |
| `new Document()`                    | `Document.create()`                                      |
| `new Paragraph({ text })`           | `textParagraph(text)`                                    |
| `new TextRun({ bold: true, text })` | `bold(text)`                                             |
| `new Table({ rows })`               | `table(rows)`                                            |
| `Packer.toBuffer(doc)`              | `toBuffer(doc)`                                          |
| ❌ No reader                        | `readDocx(buffer)`                                       |
| ❌ No modify                        | `replaceText(doc, old, new)`                             |
| ❌ No template                      | `fillTemplate(doc, data)`                                |
| ❌ No forms                         | `extractFormFields(doc)` / `fillFormFields(doc, values)` |

## Migration from `mammoth.js`

| mammoth.js                         | excelts/word                                         |
| ---------------------------------- | ---------------------------------------------------- |
| `mammoth.convertToHtml(input)`     | `readDocx(buf)` → `renderToHtml(doc)`                |
| `mammoth.convertToMarkdown(input)` | `readDocx(buf)` → `renderToMarkdown(doc)`            |
| Style maps                         | `parseStyleMap(rules)` / `matchStyleMap(style, map)` |
| ❌ No writing                      | Full read/write/modify                               |
