# Word Module

[中文](README_zh.md)

Zero-dependency TypeScript library for reading, writing, and manipulating DOCX files.
Works in Node.js 22+ and modern browsers.

## Features

- **Create, read, and modify DOCX files** — full WordprocessingML support
- **Document builder** — paragraphs, headings, rich text runs (bold/italic/underline/color/highlight)
- **Tables** — styles, borders, cell merges (horizontal + vertical), nested tables
- **Images** — inline and floating (JPEG, PNG, GIF, BMP, TIFF, SVG with raster fallback)
- **Lists** — bulleted, numbered, multi-level
- **Page layout** — sections, page size/orientation/margins, columns, breaks
- **Headers and footers** — default, first, even
- **Hyperlinks and bookmarks** — external, internal, cross-references
- **Table of contents** — field-driven, with cached entries
- **Footnotes, endnotes, and comments** — including threaded/extended comments
- **Fields** — PAGE/NUMPAGES, TOC, INDEX/XE, REF, SEQ, STYLEREF, formulas, and more
- **Math** — OMML equations (fractions, radicals, matrices, n-ary, …)
- **Drawing shapes** — preset geometries with fill, line, gradient, shadow/glow/3-D effects
- **Charts** — opaque preservation plus a from-scratch builder; bridge to the Excel chart engine
- **Track changes** — accept/reject insertions, deletions, moves, and property changes
- **Templates** — `{{variable}}`, `{{#if}}`, `{{#each}}`, mail merge, template patching
- **Form fields and content controls (SDT)** — extract and fill; OpenDoPE data binding
- **Font embedding** — with automatic glyph subsetting
- **Document protection** — read-only/comments/forms with password
- **Encryption** — read and write password-protected DOCX (Agile Encryption, as Word writes it)
- **Digital signatures** — detection and metadata extraction
- **Conversion** — DOCX ↔ HTML, DOCX ↔ Markdown, Excel → Word, Word → PDF
- **Flat OPC** — single-XML `.xml` representation round-trip
- **Diff and merge** — compare two documents, combine multiple documents
- **Streaming writer** — `new Stream.DocumentWriter()` for large documents
- **Validation** — structural checks with severity-tagged issues
- **OOXML Strict** — transparent normalization to Transitional on read
- **Browser support** — the same `documonster/word` import works in Node.js and browsers

## Quick Start

```typescript
import { Document } from "documonster/word";

// Create a document
const doc = Document.create();
Document.addHeading(doc, "Hello World", 1);
Document.addParagraph(doc, "This is a paragraph.");
Document.addTable(doc, [
  ["A", "B"],
  ["1", "2"]
]);

const buffer = await Document.toBuffer(doc);
// Write to file, send as response, etc.

// Read a document
const parsed = await Document.read(buffer);
console.log(parsed.body.length, "elements");
```

## Core API

### Document Builder

```typescript
import { Document, Build, PageSizes, Units } from "documonster/word";

const doc = Document.create();

// Text
Document.addParagraph(doc, "Simple text");
Document.addParagraph(
  doc,
  Build.paragraph([Build.text("Normal "), Build.bold("bold "), Build.italic("italic")])
);
Document.addHeading(doc, "Title", 1);

// Tables
Document.addTable(
  doc,
  [
    ["Header1", "Header2"],
    ["cell1", "cell2"]
  ],
  // headerRow repeats the row on every page; headerRunProperties makes it look like a header
  { headerRow: true, headerRunProperties: { bold: true } }
);

// Images
// width / height are EMU — convert with Units
Document.addImage(doc, pngBytes, {
  type: "png",
  width: Units.cmToEmu(5),
  height: Units.cmToEmu(3)
});
Document.addFloatingImage(doc, jpgBytes, {
  type: "jpeg",
  width: Units.cmToEmu(4),
  height: Units.cmToEmu(4),
  wrap: { style: "square" }
});

// Lists
Document.addBulletList(doc, ["Item 1", "Item 2", "Item 3"]);
Document.addNumberedList(doc, ["First", "Second", "Third"]);

// Page layout
Document.setSectionProperties(doc, {
  pageSize: PageSizes.A4,
  margins: {
    top: Units.cmToTwips(2.5),
    bottom: Units.cmToTwips(2.5),
    left: Units.cmToTwips(2),
    right: Units.cmToTwips(2)
  }
});

// Headers/Footers
Document.setHeader(doc, "default", {
  children: [Build.textParagraph("Page Header")]
});

// Build & export
const bytes = await Document.toBuffer(doc);
```

### Reading Documents

```typescript
import { Document, Query } from "documonster/word";

const doc = await Document.read(fileBuffer);

// Extract text content
const text = Query.getText(doc);

// Search
const results = Query.searchText(doc, /pattern/g);
```

On Node, `Document.readFile(path)` and `Document.writeFile(doc, path)` read and write a file
directly; the browser build publishes neither — read the bytes yourself and pass
them to `Document.read`. A file the system will not read or write is a `DocxFileError`
(`path`, `operation`) whose `cause` is the system's error (`error.cause.code` is
`ENOENT`, `EACCES`, …), as Excel's `ExcelFileError` is; a file that was read but is
not a valid document is a `DocxParseError`.

A part's XML is read as Word reads it: from its first `<` (white space a producer wrote before the
XML declaration is no error), and a byte that is not well-formed UTF-8 is read as U+FFFD rather than
refused, so a document with a damaged byte still opens. `Document.read` and `Stream.DocumentReader`
read alike, under the same `securityPolicy`; a limit either reaches is a `DocxLimitExceededError`.
Every part counts toward `maxPackageSize` — one the streaming reader never inflates, such as media,
by the size its entry declares — and a part is refused at `maxPartSize` before it is inflated, not
after. The policy is read once per call: each field is copied, checked, and the copy is what the read
uses. The main part is `w:document` in Word's namespace (Transitional or Strict, under any prefix),
with `w:body` as its child; markup the model does not represent is written back with the namespace
declarations it relied on from its ancestors.

### One Document Object

`Document.create()` and `Document.read()` return the same thing — a
`DocxDocument`, the document's content tree — and every namespace takes it.
`Document.*` helpers therefore work on a document you read as well as on one
you created. Its fields are public, so you may also edit the tree directly,
interleaved with helper calls. New footnotes, endnotes, comments, images and
lists get ids that do not collide with the ones the document uses at that
moment, provided each edit between helper calls either

- gives a collection a new array — on the document, or on a copy such as
  `{ ...doc, footnotes: [...] }`, which is how the `readonly` model types
  express an edit — or
- appends entries to the end of a collection (`push`).

Both are tracked exactly and cost the helpers O(1) per call, so a loop of
100,000 additions stays linear. To replace, remove or reorder entries, assign
a new array: an entry changed in place (`doc.footnotes[0] = …`, `splice`,
`sort`, editing an entry object) is not seen by the next helper call. Such an
edit never reaches the file silently, though — if it leaves two footnotes,
endnotes, comments or numbering definitions with the same id, or two different
images under one key, `Document.toBuffer` rejects with a `DocxDuplicateIdError` that names
the collection and the id, because content refers to those entries by id and
the writer cannot tell which entry a reference meant. Bookmark and drawing ids
are renumbered by the writer whenever they clash, so any edit is safe there.

`Document.*` and `Build.*` check what they are given before they change
anything. A value with no valid form in the file — a `NaN` or fractional font
size, a heading level outside 1–9, a color written `"#FF0000"` instead of
`"FF0000"`, an underline style Word does not define, `null` where text belongs,
an `insertContent` index past the end — throws a `DocxArgumentError`
whose `field` names the argument (`properties.run.size`) and whose message
names the call: `Document.addHeading: level must be an integer 1–9, got 12`.
A drawing's size or position is reported the same way, as its subclass
`DocxInvalidDrawingError`. Content you build by hand is not checked until it is
written.

An image's key — `ImageDef.rId`, or one of its `aliasRIds` — names the picture
in every part: the body, headers, footers, notes, comments and picture bullets
all refer to `doc.images` by key. A package's relationship ids, by contrast,
are local to each part (two headers may each call their own logo `rId1`), so
`Document.read` gives a picture whose local id another picture already holds a fresh
key, and the writer gives every part that shows a picture a relationship of
its own. `Transform.merge` renames the keys of an appended document that
collide with the base's, wherever they are used.

A document from `Document.create()` carries `deriveEastAsianDefaults: true`:
when it is written, laid out or converted, the East Asian proofing language
and typeface are derived from its text unless you stated them. A document from
`Document.read()` does not, so a round trip never acquires a language its author did
not choose. It is an ordinary field — copies keep it, and you may set or clear
it on any document.

The same flag states the document defaults current Word gives a new document —
Aptos 12pt, `en-US`, 8pt after each paragraph, 1.16 line spacing — wherever you
left them out (`Document.setDocDefaults` and `Document.useDefaultStyles` values
are kept). The document is laid out and written with them, so it opens in any
Word version, whatever its Normal template, exactly as the library paginated
it. A file read without `w:docDefaults` keeps lacking them; the library lays it
out in Calibri 11pt, single spaced, which is a guess — Word takes such a file's
defaults from the user's Normal template.

```typescript
import { Document } from "documonster/word";

const doc = await Document.read(fileBuffer);
Document.addParagraph(doc, "Appended paragraph");
Document.addFootnote(doc, "A new footnote");
const bytes = await Document.toBuffer(doc);
```

#### Who changes what

One rule decides whether a function edits the document you give it:

- **`Document.*` edits the document you pass, in place** — the builders, and
  the whole-document edits `Document.replaceText`, `Document.mailMerge`,
  `Document.acceptAllRevisions` / `rejectAllRevisions` / `acceptRevision` /
  `rejectRevision` and `Document.setCompatibilityMode`.
- **Every other namespace function that returns a `DocxDocument` leaves its
  input unchanged** — `Template.fill` / `fillFromSource`, `Transform.*`, `Security.*`, `Ole.add`,
  `Vba.add` / `remove`, `Font.addEmbedded`. The result may share unchanged
  nodes with the input (or be the input itself when there is nothing to
  change). `Document.*` copies whatever two documents share before it changes
  it, so editing either one in place never reaches the other — no clone needed.
  Take `Document.clone(...)` for a snapshot, or before editing nodes by hand.

`Query.*` only reads.

```typescript
import { Document, Template } from "documonster/word";

const template = await Document.read(templateBytes);
for (const row of rows) {
  // `fill` leaves `template` alone, and editing the letter never reaches it.
  const letter = Template.fill(template, row);
  Document.replaceText(letter, "DRAFT", "FINAL");
  await Document.writeFile(letter, `out/${row.id}.docx`);
}
```

### Modifying Documents

```typescript
import { Document } from "documonster/word";

const doc = await Document.read(buffer);
const count = Document.replaceText(doc, "OLD_TEXT", "NEW_TEXT"); // edits `doc`, returns the number replaced
const output = await Document.toBuffer(doc);
```

## Advanced Features

### Template Engine

```typescript
import { Template } from "documonster/word";

const filled = Template.fill(doc, {
  name: "John",
  showDetails: true,
  items: ["A", "B", "C"]
});
// Supports: {{variable}}, {{#if cond}}...{{/if}}, {{#each arr}}...{{/each}}
```

`Template.fill` (and `Package.fillTemplate`, the same engine on `.docx`
bytes) resolves variables, conditionals and loops from one data object,
comment bodies included. `Package.patch` / `Package.patchTemplate` are a
different engine: each named placeholder is replaced exactly with text, a
paragraph, a table or an image — no data object, no conditionals or loops,
comments left alone. Use `fill` when the template carries logic, `patch` when
you only swap placeholders for ready-made content.

A loop (`{{#each}}`, table-row loops and `{{#cols}}` included) treats a value
that is not an array as one item, and `undefined` or `null` as none. So one
template fills one line item as it fills several — XML cannot say "a list of
one", and `Template.createXmlDataSource` reads a single `<item>` as an object
where two read as an array.

XML attributes are `@name` keys: `<amount currency="EUR">10.00</amount>` reads
as `{ "@currency": "EUR", "#text": "10.00" }`, and the template writes and tests
a record with `#text` as that text — `{{amount}}` is `10.00`,
`{{amount.@currency}}` is `EUR`. An element with `xsi:nil="true"` is `null`;
names are read without their namespace prefix.

Inside a loop, `{{@index}}` counts from 0 and `{{@number}}` from 1, and
`{{@first}}` / `{{@last}}` mark the ends (`{{#if @last}}`). `{{.name}}` is the
current item's own `name` — the way to reach an item key such as an XML
attribute `@number` that a loop variable shares a name with.

A placeholder takes Word's field switches, as a `MERGEFIELD` does:

```typescript
import { Document, Template } from "documonster/word";

const invoice = Document.create();
Document.addParagraph(invoice, 'Total: {{total \\# "#,##0.00"}} due {{due \\@ "d MMMM yyyy"}}');
Document.addParagraph(invoice, '{{customer \\* Upper}}{{phone \\b " — Tel. "}}');
const filled = Template.fill(invoice, {
  total: "1234.5",
  due: "2026-11-07",
  customer: "Acme",
  phone: ""
});
// Total: 1,234.50 due 7 November 2026
// ACME

Template.fill(
  invoice,
  { total: "1234.5", due: "2026-11-07", customer: "Acme", phone: "" },
  {
    locale: "de-DE"
  }
); // a picture is then written in German symbols: "#.##0,00" → 1.234,50, "7. November 2026"
```

`\#` is a numeric picture, `\@` a date-time picture, `\*` a case or number
format (`Upper`, `Caps`, `roman`, `CardText`, …), and `\b` / `\f` text written
before / after a value that is not empty. The data is read as machines write it
— numbers with a decimal point, dates as ISO, XML Schema `dateTime` or SAP's
`YYYYMMDD`, a `Date` by its UTC time — whatever `locale` is; `locale` is what
the picture is written in, as Word's regional settings are, and it groups digits
as the locale does (`en-IN` writes `12,34,567.00`). A value a picture cannot
take throws in strict mode and is written as it is otherwise; a switch that is
not one always throws.

### Fields

```typescript
import { Document, Query, Transform } from "documonster/word";

// Every field as it reads — whole, or held in pieces across runs and paragraphs
const fields = Query.listFields(doc);
// → [{ type: "PAGE", instruction: " PAGE ", result: "3", split: false, part: "footer", … }, …]

// Merge fields replaced with values; page numbers, SEQ, REF, formulas and
// tables of contents recomputed — the result between a field's separate and
// end characters replaced, the field kept
Document.mailMerge(doc, { Name: "Alice" }); // in place
const updated = Transform.updateFields(doc);
```

Both apply a field's switches as Word does: `\* Upper` / `Lower` / `Caps` /
`FirstCap`, the number formats (`roman`, `ALPHABETIC`, `Ordinal`, …) and English
words (`CardText`, `OrdText`, `DollarText`), `\#`
numeric pictures (`\# "#,##0.00"`) and `\@` date pictures
(`\@ "d MMMM yyyy"`); a `MERGEFIELD`'s `\b` / `\f` text is added only around a
non-empty value. `updateFields` also evaluates `QUOTE` (its text, then its
switches). After merging, an `IF` that tests a merge field is evaluated
again with the merged value. In `listFields`, a field nested in another's
instruction appears as Word's field codes view shows it
(`IF { MERGEFIELD Vip } = "Y" …`), and `split` tells how the model stores the
field (as its `w:fldChar` pieces, or as one `FieldContent`).

Page numbers computed here (`PAGE`, `NUMPAGES`, `SECTIONPAGES`, `PAGEREF`,
`INDEX`, table of contents entries) come from the library's own page layout —
the pages `Layout.compute` and the PDF export produce — with each face's own
metrics, Word's paragraph spacing, widow control and keep rules. A table of
contents is laid out with its new entries before they are numbered, as Word
does, and a `PAGEREF` or table of contents entry shows a page's number as its
section numbers it (`ii`, or `1` after a restart). No `w:updateFields` is set:
Word opens the document without asking to update its fields. A caller can still
set `settings.updateFieldsOnOpen` themselves.

### Form Fields

```typescript
import { Query, Transform } from "documonster/word";

// `Build.formTextField` / `Build.formCheckboxField` build form-field runs.

// Extract form data
const fields = Query.listFormFields(doc);
// → [{ name: "FullName", type: "text", value: "..." }, ...]

// Fill form data (a new document; `doc` is unchanged). A text field takes a
// string, a checkbox a boolean, a drop-down an entry's index or its text; a
// name no field has is refused.
const filled = Transform.fillFormFields(
  doc,
  new Map<string, string | boolean | number>([
    ["FullName", "Jane Doe"],
    ["AgreeTerms", true],
    ["Country", 2]
  ])
);
```

### Data Binding (OpenDoPE)

```typescript
import { Transform } from "documonster/word";

// Resolve SDT data bindings against CustomXML parts (a new document)
const resolved = Transform.resolveDataBindings(doc);

// Or with override data
const resolved2 = Transform.resolveDataBindings(
  doc,
  new Map([["{GUID}", "<root><field>value</field></root>"]])
);
```

A binding's XPath prefixes mean what its `w:prefixMappings` binds them to: `/p:root/p:value` with
`xmlns:p="urn:x"` finds the elements of `urn:x`, whatever prefix the data gives that namespace. A
prefix the binding does not map is matched by local name; a store whose namespaces do not resolve
matches no mapped prefix.

The XPath read is an absolute location path of child steps (`/root/item`), as Word and OpenDoPE
write them, each with an optional 1-based position (`[2]`, `[last()]`), and an optional final
attribute step (`/@id`, an unprefixed one naming the attribute in no namespace). A binding using
anything else — `//`, another axis, a function, a union, another predicate — leaves its SDT as it
was.

### Drawing Shapes with Effects

```typescript
import { Build } from "documonster/word";

const shape = Build.shape({
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
import { Font } from "documonster/word";

// Embed with automatic subsetting (only glyphs used in document)
const result = Font.embed({
  name: "CustomFont",
  data: fontFileBytes,
  style: "regular",
  usedCharacters: "Hello World" // Only these glyphs are embedded
});

// Add to document
const docWithFonts = Font.addEmbedded(doc, [result]);
```

### Track Changes

```typescript
import { Document } from "documonster/word";

// Both edit the document in place and return how many revisions they resolved;
// clone first to keep the original.
const accepted = Document.clone(doc);
Document.acceptAllRevisions(accepted);
const rejected = Document.clone(doc);
Document.rejectAllRevisions(rejected);
```

### Document Diff

```typescript
import { Diff } from "documonster/word";

const diff = Diff.compare(docA, docB);
// → { changes: [{ type: "added"|"removed"|"modified", ... }] }
```

### Document Merge

```typescript
import { Transform } from "documonster/word";

const merged = Transform.merge([doc1, doc2, doc3], {
  sectionBreak: "nextPage"
});
```

### Streaming Writer

```typescript
import { Stream } from "documonster/word";

const stream = new Stream.DocumentWriter();
stream.addText("Title", { style: "Heading1" });
for (const item of largeDataset) {
  stream.addText(item.text);
}
const buffer = await stream.toBuffer();
```

With `new Stream.DocumentWriter({ sink })` the bytes flow into the sink as they are produced: finish with `await stream.commit()`, which ends the sink (`toBuffer()` is refused, since the writer holds no bytes).

### Document Protection

```typescript
import { Security } from "documonster/word";

// Returns a new document; `doc` is left as it was.
const protectedDoc = await Security.protect(doc, {
  edit: "readOnly",
  password: "secret"
});
const isProtected = Security.isProtected(protectedDoc); // true
const valid = await Security.verifyPassword(protectedDoc, "secret"); // true
```

A hash algorithm is written `"SHA-256"` (as `Security.getState` reports it) here and in `WordCrypto.encrypt`; the
encryption format's own spelling `"SHA256"` is accepted by both as well.

Protection restricts editing in Word; it does not encrypt the file. To write a
password-protected file, pass `password` to `Document.toBuffer` (see
[Encryption & Signatures](#encryption--signatures)).

### Validation

```typescript
import { Validation } from "documonster/word";

const result = Validation.validate(doc);
if (!result.valid) {
  console.log(result.issues); // [{ severity, message, path }]
}
```

A read document may state a value a field does not take — `<w:jc w:val="bogus"/>`. The field then reads as absent,
the text is kept and written back unchanged while the field is absent, and `Validation.validate` reports it as a
warning with rule `"unrecognized-value"` and the path to the field (`body[3].properties.alignment`). A value you set
takes over; clearing the field again brings the kept text back, so `Document.discardUnrecognizedValues(doc)` is how
to remove them — all, or with `{ filter }` the ones you pick by path and text. `Document.listUnrecognizedValues(doc)`
lists the ones kept now, each a `DocxUnrecognizedValue` (`{ path, value }`), the same objects the filter is given.

### HTML/Markdown Conversion

```typescript
// DOCX → Markdown (GFM: headings, bold/italic/strike, inline code,
// code blocks, blockquotes, ordered/unordered lists, tables with
// alignment, links, images, footnotes)
import { WordMarkdown } from "documonster/word/markdown";
const md = WordMarkdown.fromDocx(doc);
const mdSetext = WordMarkdown.fromDocx(doc, { headingStyle: "setext" });

// Markdown → DOCX (full document or body fragment)
const fromMarkdown = await WordMarkdown.toDocx("# Title\n\nHello **world**");
const { body: bodyItems } = await WordMarkdown.toDocxBody("- a\n- b");

// Table columns are sized from their content and a fenced code block is set
// small enough that its longest line does not wrap — both are decisions the
// producer has to make, because a `.docx` carries a column grid and a run size,
// not an instruction to work them out later. Say which measure they should be
// sized against if the body is going somewhere other than a Letter page:
const wide = await WordMarkdown.toDocx(markdown, {
  contentWidth: 15840 - 2 * 720, // landscape Letter, half-inch margins
  codeBlockFit: "shrink" // the default; "wrap" keeps the size and breaks lines
});

// DOCX → HTML
import { WordHtml } from "documonster/word/html";
const { html } = WordHtml.fromDocx(doc);

// HTML → body content, with the styles, list numbering and images it refers to
const { body, styles, numberingInstances, images } = WordHtml.toDocxBody(
  "<h1>Hello</h1><p>World</p>"
);

// HTML → a complete document (heading styles and list numbering included)
const htmlDoc = WordHtml.toDocx("<h1>Hello</h1><ul><li>World</li></ul>");
```

### Flat OPC Format

```typescript
import { Convert, Document } from "documonster/word";

// Single-XML representation of a DOCX
const flatXml = await Document.toFlatOpc(doc);
const again = await Document.read(flatXml); // Document.read takes Flat OPC text (or bytes) too
const parts = Convert.flatOpcToParts(flatXml); // Map<partPath, bytes>

// Base64, as Document.toBase64 writes it
const fromBase64 = await Document.read(await Document.toBase64(doc), {
  base64: true
});
```

`Convert.isFlatOpc` answers `true` only when the root element (after any declaration, comments,
processing instructions and DOCTYPE) is `package` in the package namespace; a ZIP is never Flat OPC.
Reading refuses, with `DocxParseError`, a part whose `xmlData` does not hold exactly one root element,
a part with more than one `xmlData`/`binaryData`, elements inside `binaryData`, and `xmlData` in a
document with a DOCTYPE — a part is copied as written, so the DOCTYPE's defaulted attributes and
entities could not come with it.

### Encryption & Signatures

A password-protected DOCX — Agile Encryption in an OLE compound file, what Word writes for "Encrypt with
Password" — is read with `{ password }` (`Document.read`, `Document.readFile`, the `Package` functions,
`new Stream.DocumentReader`) and written with it (`Document.toBuffer`, `Document.toBase64`, `Document.writeFile`); `encryption` tunes
it (Office's defaults: AES-256, SHA-512, 100,000 iterations). Excel's `Workbook.read` / `toBuffer` take the same
options, and both modules share one implementation. The encryption code is loaded only when a password is given.
Without a password an encrypted file is a `DocxEncryptedError`; a wrong one is a `DocxDecryptionError` whose
`reason` is `"password"`. Only Agile Encryption is implemented: Word 2007's Standard Encryption is refused with
`reason: "unsupported"`.

```typescript
import { Document } from "documonster/word";

const encrypted = await Document.toBuffer(doc, { password: "password" });
const reread = await Document.read(encrypted, { password: "password" });
const tuned = await Document.toBuffer(doc, {
  password: "password",
  encryption: { keyBits: 128, hashAlgorithm: "SHA-256", spinCount: 50_000 }
});
```

Lower-level helpers live on the `documonster/word/crypto` subpath, so they stay
out of bundles that only read/write plain DOCX.

```typescript
import { Document } from "documonster/word";
import { WordCrypto } from "documonster/word/crypto";

// Decrypt a password-protected DOCX to its package (Agile Encryption)
if (WordCrypto.isEncrypted(bytes)) {
  const plain = await WordCrypto.decrypt(bytes, "password");
}

// Encrypt package bytes you already have
const encrypted = await WordCrypto.encrypt(await Document.toBuffer(doc), "password");

// Check a password without decrypting the package
const ok = await WordCrypto.verifyPassword(bytes, "password");

// Inspect embedded XMLDSig signatures (read-only; no verification)
const signatures = WordCrypto.listSignatures(doc);
```

The building blocks behind `decrypt` are public too, for custom pipelines, in
two low-level sub-namespaces: `WordCrypto.Cfb` (`read` / `write` for OLE2
compound files) and `WordCrypto.Agile` (`parseEncryptionInfo`, `deriveKey` with
`BLOCK_KEYS`, `verifyPassword` against parsed info, and `decryptPackage`, which
takes the size cap as `{ maxDecryptedSize }` as `decrypt` does). The
ODTTF font obfuscation primitives are font tools and live in `Font`
(`Font.obfuscate` / `Font.deobfuscate` / `Font.generateObfuscationKey`, from
`documonster/word`).

```typescript
import { WordCrypto } from "documonster/word/crypto";

const entries = WordCrypto.Cfb.read(bytes);
const infoXml = new TextDecoder().decode(
  entries.find(e => e.name === "EncryptionInfo")!.data.slice(8)
);
const info = WordCrypto.Agile.parseEncryptionInfo(infoXml);
const zip = await WordCrypto.Agile.decryptPackage(
  entries.find(e => e.name === "EncryptedPackage")!.data,
  info,
  "password"
);
```

### Excel → Word

Convert an Excel `Workbook` into a `DocxDocument`, mapping worksheets to
Word tables with cell formatting (fonts, colours, alignment, fills,
borders), column widths, rich-text runs, and an optional title page.
Hidden sheets are skipped; rows/columns can be capped.

```typescript
import { Document } from "documonster/word";
import { WordExcel } from "documonster/word/excel";

// Workbook → DocxDocument (all visible sheets, formatting preserved)
const doc = WordExcel.toDocx(workbook);
const docxBytes = await Document.toBuffer(doc);

// With options: a title page, only some sheets, capped dimensions.
// Name the sheets: numeric positions are deprecated (removed in 1.0) — they
// are 0-based here but 1-based in `Pdf.fromExcel` and sheet ids in
// `Workbook.getWorksheet`.
const doc2 = WordExcel.toDocx(workbook, {
  includeTitlePage: true,
  sheets: ["Summary", "Regions"],
  maxRows: 100,
  maxColumns: 12,
  preserveFormatting: true
});

// Reverse direction: pull the tables out of a DOCX into a Workbook
const extracted = WordExcel.extractTables(doc);
```

Charts embedded in a Word document also bridge to the Excel chart engine
(27 chart families, classic and modern ChartEx); `Pdf.fromDocx` draws them with it by
default (see "Charts" below). The bridge pieces are public: `WordExcel.toChartSvg(chart)`
renders a chart, `WordExcel.toChartModel(chart)` returns the Excel chart model,
`WordExcel.toChartXlsx(series)` builds the embedded data workbook (`.xlsx` bytes) and
`WordExcel.toChartExXml(options)` builds a ChartEx part. A chart's automatic series colours
and scheme colours are its document theme's: `Pdf.fromDocx` uses `doc.theme`, and
`WordExcel.toChartSvg(chart, { themeColors: doc.theme?.colorScheme.colors })` takes them —
`Layout.toSvg` hands its `renderChart` callback the same colours as `context.themeColors`.
A chart read from a file and written back loses nothing: what `Chart` has no field for — scheme
colours and their transforms, gradients, line styles, gridlines, text formatting, layouts, per-label
data labels, data references, extension lists, the part's style and colour parts — is written back
as it was read (and drawn), and setting a field rewrites only what it governs (`values`, an axis'
`min`, `title`, `color`, `pointColors[i]`, …). `pointColors` is by point index: entry `i` colours
point `i`, and a point without a colour of its own is a hole.

### Word → PDF

Convert a `DocxDocument` to PDF bytes. The bridge is a thin layer over
the shared Word layout engine, so line wrapping, pagination, tables,
inline images, headers/footers, and floats all render identically to the
SVG path.

```typescript
import { Pdf } from "documonster/pdf";
import { Document } from "documonster/word";

const doc = await Document.read(docxBytes);
const pdfBytes = await Pdf.fromDocx(doc);

// Override page geometry (points). Any field omitted falls back to the
// document's section properties, then engine defaults (US Letter, 1").
const pdf2 = await Pdf.fromDocx(doc, {
  pageWidth: 595, // A4 width
  pageHeight: 842, // A4 height
  margins: { top: 72, right: 72, bottom: 72, left: 72 },
  headerMargin: 36, // header band offset from the top edge
  footerMargin: 36, // footer band offset from the bottom edge
  defaultFontFamily: "Helvetica",
  defaultFontSize: 11
});

// Document options, as every converter takes them: metadata (each field left out comes
// from the document's core properties, keywords included), a fixed creation date for
// reproducible bytes, and a password.
const pdf3 = await Pdf.fromDocx(doc, {
  title: "Quarterly report",
  creationDate: new Date("2026-01-01T00:00:00Z"),
  encryption: { ownerPassword: "owner", userPassword: "open" }
});
```

**Charts.** Both classic (`<c:chart>`) and modern ChartEx (`<cx:chartSpace>` — sunburst,
treemap, waterfall, funnel, boxWhisker, histogram, pareto, regionMap) charts render as
full vector PDF by default, drawn by the Excel chart engine; nothing needs to be passed.
A chart the engine cannot draw falls back to its inline SVG, then to a titled placeholder
box (no throw, no blank page). To draw classic charts yourself, supply a `chartRenderer`:

```typescript
import { Pdf } from "documonster/pdf";

const pdf = await Pdf.fromDocx(doc, {
  chartRenderer: (chart, page, rect) => {
    if (chart.type !== "pie") return false; // decline: the built-in renderer draws it
    page.drawRect({ ...rect, stroke: { r: 0, g: 0, b: 0 } });
  }
});
```

A `chartRenderer` may return `false` to decline a particular chart; the
bridge then falls back to the built-in vector renderer, then to the
inline SVG / placeholder.

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
| Encrypted .docx        | ✅ Read and write with a password (Agile Encryption)           |
| Digital Signatures     | 🔍 Detection & metadata extraction (no signing/verification)   |
| Browser                | ✅ (same `documonster/word` import)                            |
| Node.js 22+            | ✅                                                             |
