# Migration Guide: `@cj-tech-master/excelts` → `documonster`

> **Version.** This guide targets **documonster 0.22.0 or later**. Check what npm serves with
> `npm view documonster version`: until 0.22.0 is published, a bare `npm install documonster` installs 0.21.1,
> which this guide does not describe. Install `documonster@^0.22.0` (see [Install](#install)).

This guide describes every breaking change introduced when the package
`@cj-tech-master/excelts` was renamed and restructured into `documonster`. It is
written to be followed mechanically: an AI or a developer can rewrite an existing
`excelts` codebase to `documonster` by applying the rules below in order.

Every replacement named here is the **current, non-deprecated** form. Two generated companions cover what the
prose does not:

- the [name index (§12)](#12-name-index-generated) — every name excelts 10.2.0 exported that `documonster` publishes
  under another name, entry by entry;
- [docs/api-changes-from-excelts-10.2.0.md](docs/api-changes-from-excelts-10.2.0.md) — every name that kept its name
  but changed type.

A name `documonster` still accepts but has deprecated works until 1.0; your editor strikes it through and its TSDoc
names the replacement. (If you were already on an earlier `documonster` release rather than on excelts,
[MIGRATION.md](MIGRATION.md) is the guide for you.)

**Baseline: `@cj-tech-master/excelts` 10.2.0**, the last release. Every name that release publishes — each export of
each subpath, node and browser, and each member of the Excel classes — is either still available unchanged in
`documonster` or accounted for in this guide; `pnpm verify:excelts-guide` checks that against the published package,
and checks that every replacement this guide names exists and is not deprecated.

Coming from `exceljs` itself rather than excelts? Read
[MIGRATION_EXCELJS_TO_DOCUMONSTER.md](MIGRATION_EXCELJS_TO_DOCUMONSTER.md): excelts had already renamed or reshaped
part of the exceljs API, so the two guides differ in their left-hand columns.

> **TL;DR of the four biggest changes**
>
> 1. **Package renamed**: `@cj-tech-master/excelts` → `documonster`.
> 2. **No root entry**: the bare `import ... from "@cj-tech-master/excelts"`
>    (and `/browser`) subpath is **gone**. Excel now lives at
>    `documonster/excel`.
> 3. **Excel is no longer class-based.** The `Workbook` / `Worksheet` / `Cell`
>    / `Row` / `Column` / `Range` / `Table` / `Image` / `Note` / `FormCheckbox`
>    **classes were removed** and replaced by **namespaces of plain functions**
>    (`Workbook.create()`, `Cell.setValue(ws, "A1", 42)`, …). You never call
>    `new` and never hold instances with methods; you pass the workbook or
>    worksheet into functions and address cells, rows and columns by position.
> 4. **CSV, Markdown, formula and PDF integrations moved off `Workbook`.** They
>    now live at `documonster/excel/csv`, `documonster/excel/markdown`,
>    `documonster/excel/formula`, and `documonster/pdf` as free functions taking
>    the workbook as the first arg.

## Contents

- [Quick start: the ten most common calls](#quick-start-the-ten-most-common-calls)
- [Code that compiles but behaves differently](#code-that-compiles-but-behaves-differently)
- [1. Package & tooling changes](#1-package--tooling-changes)
- [2. Import subpath map](#2-import-subpath-map)
- [3. Excel: from classes to namespaces](#3-excel-from-classes-to-namespaces-the-big-one)
- [4. Excel API mapping tables](#4-excel-api-mapping-tables-class-member--namespace-function)
- [5. CSV & Markdown moved off `Workbook`](#5-csv--markdown-moved-off-workbook)
- [6. PDF: `Pdf.*` namespace](#6-pdf-pdf--exceltopdf--pdf-namespace)
- [7. Formula engine: no more install step](#7-formula-engine-no-more-install-step)
- [8. Other modules now namespaced](#8-other-modules-now-namespaced)
- [9. Step-by-step migration checklist](#9-step-by-step-migration-checklist)
- [10. Complete before/after example](#10-complete-beforeafter-example)
- [11. Quick reference: what was removed](#11-quick-reference-what-was-removed)
- [12. Name index (generated)](#12-name-index-generated)

### Quick start: the ten most common calls

| Task           | excelts 10.2.0                                           | documonster                                                                                           |
| -------------- | -------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| Create         | `new Workbook()`                                         | `Workbook.create()`                                                                                   |
| Read a file    | `await wb.xlsx.readFile("in.xlsx")`                      | `const wb = await Workbook.readFile("in.xlsx")` (Node; `Workbook.read(bytes)` anywhere)               |
| Write a file   | `await wb.xlsx.writeFile("out.xlsx")`                    | `await Workbook.writeFile(wb, "out.xlsx")` (Node; `Workbook.toBuffer(wb)` anywhere)                   |
| Get / set      | `ws.getCell("A1").value` / `ws.getCell("A1").value = 42` | `Cell.getValue(ws, "A1")` / `Cell.setValue(ws, "A1", 42)`                                             |
| Style          | `ws.getCell("A1").font = { bold: true }`                 | `Cell.setFont(ws, "A1", { bold: true })`                                                              |
| Add a row      | `ws.addRow(["a", 1])`                                    | `Row.add(ws, ["a", 1])` (returns the row number)                                                      |
| Merge          | `ws.mergeCells("A1:B2")`                                 | `Worksheet.merge(ws, "A1:B2")`                                                                        |
| Iterate rows   | `ws.eachRow((row, n) => row.values)`                     | `Row.each(ws, n => Row.getValues(ws, n, { base: 1 }))`                                                |
| Recalculate    | `installFormulaEngine(); wb.calculateFormulas()`         | `ExcelFormula.calculate(wb)` from `documonster/excel/formula`                                         |
| Export PDF/CSV | `await excelToPdf(wb)` / `wb.writeCsv()`                 | `await Pdf.fromExcel(wb)` from `documonster/pdf` / `ExcelCsv.format(wb)` from `documonster/excel/csv` |

```ts
import { Cell, Row, Workbook, Worksheet } from "documonster/excel";
import { ExcelCsv } from "documonster/excel/csv";
import { ExcelFormula } from "documonster/excel/formula";
import { Pdf } from "documonster/pdf";

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Sheet1");
Row.add(ws, ["Item", "Qty"]);
Row.add(ws, ["Apples", 3]);
Cell.setValue(ws, "B3", { formula: "B2*2" });
Cell.setFont(ws, "A1", { bold: true });
Worksheet.merge(ws, "D1:E1");
ExcelFormula.calculate(wb);
Row.each(ws, n => console.log(n, Row.getValues(ws, n, { base: 1 })));

await Workbook.writeFile(wb, "tmp/quick-start.xlsx");
const again = await Workbook.readFile("tmp/quick-start.xlsx");
console.log(Cell.getValue(Workbook.getWorksheet(again, "Sheet1")!, "A2"));
const pdf = await Pdf.fromExcel(wb);
const csv = ExcelCsv.format(wb);
```

### Code that compiles but behaves differently

These ports type-check and still change what your program does. Look for each one:

- **Facet getters return a copy.** `Cell.getFont(ws, "A1")!.bold = true` edits nothing; write it back with
  `Cell.setFont(ws, "A1", { ...Cell.getFont(ws, "A1"), bold: true })` (same for `Row.*` / `Column.*` facets).
  `getStyle` is the one live, editable record.
- **What you pass in is copied.** A value, style or note object handed to a setter (`Cell.setValue`,
  `Cell.setNote`, …) is copied; mutating your object afterwards does not change the cell — call the setter again.
- **`Row.getValues(ws, n)` is 0-based.** excelts' `row.values` was 1-based with a hole at index 0; use
  `Row.getValues(ws, n, { base: 1 })` (and `Column.getValues(ws, c, { base: 1 })`) where you index by number.
- **Reading returns a new workbook.** `await wb.xlsx.load(bytes)` filled `wb`; `Workbook.read(bytes)` leaves any
  workbook you hold untouched — use the one it returns: `const wb = await Workbook.read(bytes)`.
- **Formula results are not computed for you.** A formula you set has no result until
  `ExcelFormula.calculate(wb)` runs; `Cell.getResult(ws, "B3")` is `undefined` before that.
- **Committed streaming rows are closed.** On a streamed sheet, `Cell.getValue` / `Cell.setValue` on a row already
  committed throws `ExcelStreamStateError` instead of answering or silently writing nothing.
- **Bytes are `Uint8Array`, not `Buffer`.** `toBuffer` results have no `.toString("base64")`, and `.slice()` copies
  rather than viewing; use `subarray()` for a view, `Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)` where a Node
  API needs a `Buffer` (it shares the memory; `Buffer.from(bytes)` copies).
- **Absence is `undefined`, not `null`.** `=== null` against `Table.get`, `Watermark.get`, the archive readers, …
  compiles and never matches; test `=== undefined` (or `== null`).
- **Table names Excel rejects now throw.** `Table.add(ws, { name: "my table", … })` throws `ExcelTableError` (its
  message suggests `"my_table"`) instead of being renamed silently; a duplicate name (case-insensitive) throws too.
- **A six-digit colour reads back with alpha.** `{ argb: "f56600" }` is written as `FF` + rgb (the file format
  requires four bytes), so a written-then-read file gives `"FFf56600"`. Compare against the eight-digit form.

---

## 1. Package & tooling changes

| Aspect             | Old (`excelts`)           | New (`documonster`)                        |
| ------------------ | ------------------------- | ------------------------------------------ |
| Package name       | `@cj-tech-master/excelts` | `documonster`                              |
| License            | MIT                       | Apache-2.0                                 |
| Root `.` export    | present (Excel API)       | **removed**                                |
| `./browser` export | present                   | **removed** (per-module browser builds)    |
| Module format      | ESM + CommonJS builds     | **ESM only** (CommonJS via `require(esm)`) |
| Node.js            | `>=22.0.0`                | `>=22.13.0`                                |

### Install

```bash
# npm
npm remove @cj-tech-master/excelts
npm install documonster@^0.22.0

# pnpm
pnpm remove @cj-tech-master/excelts
pnpm add documonster@^0.22.0

# yarn
yarn remove @cj-tech-master/excelts
yarn add documonster@^0.22.0
```

### Node.js, Bun and CommonJS

`documonster` needs **Node.js 22.13 or later**, Bun (the latest release only — Bun has
no LTS line to promise a range against), or a modern browser. The package is
**ESM-only** and ships one entry per module.

CommonJS code keeps working through Node's `require(esm)`, with the call site
unchanged:

```js
const { Workbook, Cell } = require("documonster/excel");
```

`require(esm)` is unflagged from Node 22.12, so 22.12 already works — but on 22.12
every such `require()` prints an `ExperimentalWarning: … Support for loading ES Module
in require() is an experimental feature` to stderr. From 22.13 it is silent, which is
why the floor is 22.13 rather than 22.12. On 22.0–22.11 it needs
`--experimental-require-module`; upgrade instead. Do not rely on a top-level `await`
anywhere in the import chain of a CommonJS consumer: `require(esm)` refuses such a
graph.

### TypeScript

The library itself is compiled with TypeScript 7.0; that is a build detail, not a
requirement on you. The published declarations are plain `.d.ts` files, and the
consumer settings documented in the [README](README.md) were verified on tsc 5.9.
The one setting that matters is `moduleResolution`:

| `moduleResolution`                    | Result                                                                      |
| ------------------------------------- | --------------------------------------------------------------------------- |
| `"bundler"`                           | works                                                                       |
| `"nodenext"`                          | works, including from a CommonJS file (it knows about `require(esm)`)       |
| `"node10"` / `"node"` with `commonjs` | works (types resolve through `typesVersions`)                               |
| `"node16"`                            | **fails** with `TS1479` — that mode predates `require(esm)`; use `nodenext` |

To type-check against the **browser** surface (where the Node-only members such as
`Workbook.readFile` are absent), add the condition:

```json
{
  "compilerOptions": {
    "moduleResolution": "bundler",
    "customConditions": ["browser"]
  }
}
```

### Jest

Jest cannot `require()` an ES module, so it needs a transform — the same two files a
Jest project already needs for `chalk`, `uuid` or `nanoid`. If you have them, add
`documonster` to the pattern you already have:

```js
// babel.config.cjs
module.exports = { presets: [["@babel/preset-env", { targets: { node: "current" } }]] };
```

```json
// jest.config.json
{
  "transform": { "\\.[jt]sx?$": "babel-jest" },
  "transformIgnorePatterns": ["/node_modules/(?!documonster)"]
}
```

Install `babel-jest` and `@babel/preset-env` as dev dependencies. Vitest needs no
configuration.

### Browsers

The browser build is chosen by the `browser` export condition, so a bundler (Vite,
webpack, Rollup, esbuild) needs nothing beyond `import { Workbook } from "documonster/excel"`.
The browser entry publishes **no file-path functions** — `Workbook.readFile` and
`Workbook.writeFile` are absent there, and calling them is a type error, not a runtime
one. Read from bytes and hand bytes back:

```ts browser
import { Workbook } from "documonster/excel";

async function openUpload(file: File) {
  return Workbook.read(new Uint8Array(await file.arrayBuffer()));
}

async function download(wb: Workbook.Handle, name: string) {
  const bytes = await Workbook.toBuffer(wb);
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  URL.revokeObjectURL(a.href);
}
```

Without a bundler, load the IIFE build from `dist/iife`. excelts published one bundle
for everything; `documonster` publishes one per module, each extending the shared
`Documonster` global:

<!-- x-release-please-start-version -->

```html
<script src="https://unpkg.com/documonster@0.25.0/dist/iife/documonster.excel.iife.min.js"></script>
<script>
  const { Workbook, Cell } = Documonster.Excel;
  const wb = Workbook.create();
  Cell.setValue(Workbook.addWorksheet(wb, "S1"), "A1", "Hello");
  Workbook.toBuffer(wb).then(bytes => console.log(bytes.byteLength));
</script>
```

<!-- x-release-please-end -->

The URL pins the version; keep it pinned in production, and raise it deliberately when you upgrade. Swap `excel` for
`word`, `pdf`, `csv`, … and read `Documonster.Word`, `Documonster.Pdf`, …. The IIFE
bundles do not include the formula calculation engine; recalculate through ESM and
`documonster/excel/formula`.

---

## 2. Import subpath map

The root entry (`.`) and `./browser` were removed. Every module now has its own
subpath. Update all import specifiers:

| Old specifier                           | New specifier                | Notes                                                          |
| --------------------------------------- | ---------------------------- | -------------------------------------------------------------- |
| `@cj-tech-master/excelts`               | `documonster/excel`          | Excel API (see §3, now namespaces)                             |
| `@cj-tech-master/excelts/browser`       | `documonster/excel`          | Browser build auto-selected via the `browser` export condition |
| `@cj-tech-master/excelts/formula`       | `documonster/formula`        | Now `Formula.*` namespace (see §7)                             |
| `@cj-tech-master/excelts/chart`         | `documonster/chart`          | Unchanged shape (free functions)                               |
| `@cj-tech-master/excelts/pdf`           | `documonster/pdf`            | `pdf`/`excelToPdf` → `Pdf.*` (see §6)                          |
| `@cj-tech-master/excelts/word`          | `documonster/word`           | Namespaced already                                             |
| `@cj-tech-master/excelts/word/html`     | `documonster/word/html`      |                                                                |
| `@cj-tech-master/excelts/word/crypto`   | `documonster/word/crypto`    |                                                                |
| `@cj-tech-master/excelts/word/markdown` | `documonster/word/markdown`  |                                                                |
| `@cj-tech-master/excelts/word/excel`    | `documonster/word/excel`     |                                                                |
| `@cj-tech-master/excelts/csv`           | `documonster/csv`            | Standalone CSV parser/formatter                                |
| `@cj-tech-master/excelts/markdown`      | `documonster/markdown`       | Standalone Markdown table module                               |
| `@cj-tech-master/excelts/xml`           | `documonster/xml`            | Now `Xml.*` namespace                                          |
| `@cj-tech-master/excelts/stream`        | `documonster/stream`         | Mostly unchanged                                               |
| `@cj-tech-master/excelts/archive`       | `documonster/archive`        |                                                                |
| _(CSV on Workbook)_                     | `documonster/excel/csv`      | **New** — Workbook CSV I/O (see §5)                            |
| _(Markdown on Workbook)_                | `documonster/excel/markdown` | **New** — Workbook Markdown I/O (see §5)                       |
| _(`wb.calculateFormulas()`)_            | `documonster/excel/formula`  | **New** — Workbook formula recalculation (see §7)              |

Browser vs Node is chosen automatically by the bundler/runtime through the
package `exports` `"browser"` condition — you no longer import a `/browser`
subpath explicitly.

---

## 3. Excel: from classes to namespaces (the big one)

`documonster/excel` exports **namespaces of functions**, not classes. The
mental model changes:

- **Before:** you created objects with `new` and called methods on them
  (`wb.addWorksheet("S")`, `cell.value = 42`, `ws.getCell("A1").font = {...}`).
- **After:** you call namespace functions and pass **opaque handles** as the
  first argument (`Workbook.addWorksheet(wb, "S")`,
  `Cell.setValue(ws, "A1", 42)`, `Cell.setFont(ws, "A1", {...})`).

The workbook and its worksheets are **handles** (`Workbook.Handle`, and the
`Worksheet.Handle` that `Workbook.addWorksheet` returns): plain data with **no
methods**. Never call a method on them; always route through a namespace
function. Cells, rows and columns have no handle at all — they are **addressed by
position** (`"A1"` or `{ row, col }`, a row number, a column letter or number) on
the worksheet handle. Row and cell handles exist only in the streaming API
(§4.11), where a row is written once and committed.

### 3.1 Available namespaces

`import { Workbook, Worksheet, Cell, Row, Column, Range, Chart, Table, Image, Pivot, Sparkline, Form, Chartsheet, ConditionalFormatting, DataValidation, DefinedNames, Note, Address, Anchor, Watermark, HeaderFooterImage, Format, Stream } from "documonster/excel";`

- **Platform-independent:** `Worksheet`, `Cell`, `Row`, `Column`, `Range`,
  `Chart`, `Table`, `Image`, `Pivot`, `Sparkline`, `Form`, `Chartsheet`,
  `ConditionalFormatting`, `DataValidation`, `DefinedNames`, `Note`, `Address`,
  `Anchor`, `Watermark`, `HeaderFooterImage`, `Format`.
- **Platform-specific** (resolve to Node or browser variant automatically):
  `Workbook` (Node adds file-path `readFile`/`writeFile`) and `Stream`.

The integrations that used to be `Workbook` methods or root exports live on their own
entries:

| Namespace       | Import from                  | Replaces                                                 |
| --------------- | ---------------------------- | -------------------------------------------------------- |
| `ExcelFormula`  | `documonster/excel/formula`  | `installFormulaEngine()` + `wb.calculateFormulas()` (§7) |
| `ExcelCsv`      | `documonster/excel/csv`      | `wb.readCsv` / `wb.writeCsv` / … (§5.1)                  |
| `ExcelMarkdown` | `documonster/excel/markdown` | `wb.readMarkdown` / `wb.writeMarkdown` / … (§5.2)        |
| `Pdf`           | `documonster/pdf`            | `excelToPdf` / `pdf` / `PdfDocumentBuilder` / … (§6)     |
| `WordCrypto`    | `documonster/word/crypto`    | `encryptDocx` (§8.5)                                     |

Also exported, flat: the error classes `ExcelError`, `isExcelError`,
`ExcelArgumentError` (a wrong argument; it names the call and the parameter),
`ExcelWorksheetNameError`, `ExcelInvalidAddressError`, `ExcelColumnOutOfBoundsError`,
`ExcelRowOutOfBoundsError`, `ExcelMergeConflictError`, `ExcelInvalidValueTypeError`,
`ExcelUnsupportedFeatureError`, `ExcelEncryptedError`, `ExcelDecryptionError`,
`ExcelFileError`, `ExcelStreamStateError`, `ExcelDownloadError`, `ExcelPivotTableError`,
`ExcelChartOptionsError`, `ExcelTableError`, `ExcelDefinedNameError`, `ExcelImageError`,
`ExcelLimitExceededError`, `XlsxParseError`, `XlsbParseError`, `XlsbFormulaDecodeError`.

**Error classes excelts exported from its root**, and what to catch now (the old names
still work as deprecated aliases of the same classes, so `instanceof` keeps matching):

| excelts 10.2.0              | documonster (`documonster/excel`)                                                     |
| --------------------------- | ------------------------------------------------------------------------------------- |
| `ExcelError`                | `ExcelError` (unchanged)                                                              |
| `isExcelError`              | `isExcelError` (unchanged)                                                            |
| `ExcelFileError`            | `ExcelFileError` (unchanged)                                                          |
| `ExcelDownloadError`        | `ExcelDownloadError` (unchanged)                                                      |
| `ExcelStreamStateError`     | `ExcelStreamStateError` (unchanged)                                                   |
| `ExcelNotSupportedError`    | `ExcelUnsupportedFeatureError`                                                        |
| `InvalidAddressError`       | `ExcelInvalidAddressError`                                                            |
| `ColumnOutOfBoundsError`    | `ExcelColumnOutOfBoundsError`                                                         |
| `RowOutOfBoundsError`       | `ExcelRowOutOfBoundsError`                                                            |
| `MergeConflictError`        | `ExcelMergeConflictError`                                                             |
| `InvalidValueTypeError`     | `ExcelInvalidValueTypeError`                                                          |
| `WorksheetNameError`        | `ExcelWorksheetNameError`                                                             |
| `PivotTableError`           | `ExcelPivotTableError`                                                                |
| `TableError`                | `ExcelTableError`                                                                     |
| `ImageError`                | `ExcelImageError`                                                                     |
| `MaxItemsExceededError`     | `ExcelLimitExceededError`                                                             |
| `XmlParseError`             | `XlsxParseError` — unexpected content in a workbook part (an `ExcelError`, as before) |
| `PdfError`, `isPdfError`, … | the same names, from `documonster/pdf`                                                |
| `BaseError`                 | not public — catch the module's base class (`ExcelError`, `PdfError`, …)              |

Every error the Excel module throws is an `ExcelError`, including a bad argument
(`ExcelArgumentError`, which names the call and the parameter) and a refused file
operation (`ExcelFileError`, with the system error as `cause`):

```ts
import { Cell, ExcelWorksheetNameError, Workbook, isExcelError } from "documonster/excel";

const wb = Workbook.create();
try {
  Workbook.addWorksheet(wb, "bad/name");
} catch (error) {
  if (error instanceof ExcelWorksheetNameError) {
    console.log("rejected sheet name:", error.message);
  }
}
try {
  Cell.setValue(Workbook.addWorksheet(wb, "S"), "not an address", 1);
} catch (error) {
  if (isExcelError(error)) {
    console.log(error.name, error.message); // ExcelInvalidAddressError …
  }
}
```

> Note: the standalone `Column`, `Range`, `Table`, `Image`, `Note`,
> `FormCheckbox`, `DataValidations`, `DefinedNames`, `Chartsheet` **class
> exports have been removed** from the top level. Their functionality is now in
> the same-named namespaces.

### 3.2 Canonical before/after

**Before (excelts):**

```ts before
import { Workbook } from "@cj-tech-master/excelts";

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");
ws.getCell("A1").value = 42;
ws.getCell("A1").font = { bold: true };
ws.addRow(["a", "b", "c"]);
await wb.xlsx.writeFile("out.xlsx"); // Node
const buf = await wb.xlsx.writeBuffer(); // Browser
```

**After (documonster):**

```ts
import { Cell, Row, Workbook } from "documonster/excel";

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Sheet1");
Cell.setValue(ws, "A1", 42);
Cell.setFont(ws, "A1", { bold: true });
Row.add(ws, ["a", "b", "c"]);
await Workbook.writeFile(wb, "out.xlsx"); // Node
const buf = await Workbook.toBuffer(wb); // works in Node & browser
```

### 3.3 The `xlsx` I/O gateway is gone

The old `workbook.xlsx.*` accessor (`readFile`, `writeFile`, `load`,
`writeBuffer`) is removed. Use `Workbook` namespace functions directly:

| Old (`wb.xlsx.*`)             | New (`Workbook.*`)                         |
| ----------------------------- | ------------------------------------------ |
| `await wb.xlsx.writeFile(p)`  | `await Workbook.writeFile(wb, p)` _(Node)_ |
| `await wb.xlsx.readFile(p)`   | `wb = await Workbook.readFile(p)` _(Node)_ |
| `await wb.xlsx.writeBuffer()` | `await Workbook.toBuffer(wb)`              |
| `await wb.xlsx.load(bytes)`   | `wb = await Workbook.read(bytes)`          |
| `wb.xlsx.write(stream)`       | `await Workbook.writeStream(wb, stream)`   |
| `wb.xlsx.read(stream)`        | `wb = await Workbook.readStream(stream)`   |

> **`writeStream` respects backpressure — attach the consumer first.** Unlike
> exceljs, whose zip writer ignores `write()` returning `false` and lets the
> sink buffer without bound, `Workbook.writeStream` waits for the sink to
> `'drain'` at each zip-entry boundary. Awaiting it on an **unconsumed**
> intermediate stream therefore deadlocks once the stream's buffers fill
> (~128 KB of output on Node 22+, i.e. every non-trivial workbook), while small
> workbooks still succeed:
>
> ```ts
> // ❌ never resolves: the consumer is attached only after `writeStream` returns
> const passThrough = new PassThrough();
> await Workbook.writeStream(wb, passThrough);
> await upload(passThrough);
>
> // ✅ start the consumer, then await the producer
> const passThrough = new PassThrough();
> const uploading = upload(passThrough);
> await Workbook.writeStream(wb, passThrough);
> await uploading;
>
> // ✅ or don't stream at all
> await upload(await Workbook.toBuffer(wb));
> ```
>
> A terminal sink (`fs.createWriteStream`, an HTTP response, an SDK upload body
> that starts reading immediately) needs no special handling.

Reading returns a new workbook handle:

```ts
import { Workbook } from "documonster/excel";

const wb = await Workbook.read(bytes); // browser / any platform
// or
const fromFile = await Workbook.readFile("in.xlsx"); // Node only
```

---

## 4. Excel API mapping tables (class member → namespace function)

Apply these substitutions. In all rows, `wb` is a workbook handle, `ws` is a
worksheet handle. Addresses are `"A1"` strings **or** 1-based `{ row, col }` positions
where noted.

### 4.1 `Workbook`

| Old (`Workbook` class)                                    | New (`Workbook.*` function)                                                                                                                                                                                     |
| --------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `new Workbook()` / `new Workbook({ formulaSyntaxProbe })` | `Workbook.create()` / `Workbook.create({ formulaSyntaxProbe })`                                                                                                                                                 |
| `wb.addWorksheet(name, opts)`                             | `Workbook.addWorksheet(wb, name, opts)`                                                                                                                                                                         |
| `wb.getWorksheet(id?)`                                    | `Workbook.getWorksheet(wb, id?)`                                                                                                                                                                                |
| `wb.worksheets`                                           | `Workbook.listWorksheets(wb)`                                                                                                                                                                                   |
| `wb.removeWorksheet(id)`                                  | `Workbook.removeWorksheet(wb, id)`                                                                                                                                                                              |
| `wb.eachSheet(cb)`                                        | `Workbook.eachWorksheet(wb, cb)`                                                                                                                                                                                |
| `wb.importSheet(source, name?)`                           | `Workbook.importWorksheet(wb, source, name?)`                                                                                                                                                                   |
| `await wb.protect(pw?, opts?)`                            | `Workbook.protect(wb, { password, ...opts })`                                                                                                                                                                   |
| `wb.unprotect()`                                          | `Workbook.unprotect(wb)`                                                                                                                                                                                        |
| `wb.addChartsheet(name, opts)`                            | `Workbook.addChartsheet(wb, name, opts)`                                                                                                                                                                        |
| `wb.addPivotChartsheet(name, pt, opts)`                   | `Workbook.addPivotChartsheet(wb, name, pt, opts)`                                                                                                                                                               |
| `wb.chartsheets`                                          | `Workbook.listChartsheets(wb)`                                                                                                                                                                                  |
| `wb.getChartsheet(x)`                                     | `Workbook.getChartsheet(wb, x)`                                                                                                                                                                                 |
| `wb.removeChartsheet(x)`                                  | `Workbook.removeChartsheet(wb, x)`                                                                                                                                                                              |
| `wb.renameChartsheet(x, name)`                            | `const cs = Workbook.getChartsheet(wb, x); if (cs) Chartsheet.setName(cs, name);` — the old call answered `false` when there was no such chartsheet; a name over 31 characters is refused, where excelts cut it |
| `wb.copyChartsheet(x, name?)`                             | `Workbook.copyChartsheet(wb, x, name?)`                                                                                                                                                                         |
| `wb.replaceChartsheetChart(x, chart)`                     | `Workbook.replaceChartsheetChart(wb, x, chart)`                                                                                                                                                                 |
| `wb.definedNames`                                         | `Workbook.getDefinedNames(wb)`                                                                                                                                                                                  |
| `wb.registerPerson(name, userId?, provId?)`               | `Workbook.registerPerson(wb, name, { userId, providerId })`                                                                                                                                                     |
| `wb.registerFunction(name, fn, opts?)`                    | `Workbook.registerFunction(wb, name, fn, opts?)` — see §7 for the tightened callback type (`opts.volatile` is kept)                                                                                             |
| `wb.unregisterFunction(name)`                             | `Workbook.unregisterFunction(wb, name)`                                                                                                                                                                         |
| `wb.addExternalLink(input)`                               | `Workbook.addExternalLink(wb, input)`                                                                                                                                                                           |
| `wb.getExternalLink(x)`                                   | `Workbook.getExternalLink(wb, x)`                                                                                                                                                                               |
| `wb.model` (get)                                          | `Workbook.getModel(wb)`                                                                                                                                                                                         |
| `wb.model = m` (set)                                      | `Workbook.setModel(wb, m)`                                                                                                                                                                                      |
| `wb.addImage(image)`                                      | `Image.add(wb, image)` _(see §4.7)_                                                                                                                                                                             |
| `Workbook.createStreamWriter(opts)`                       | `new Stream.WorkbookWriter(opts)`                                                                                                                                                                               |
| `Workbook.createStreamReader(input, opts)`                | `new Stream.WorkbookReader(input, opts)`                                                                                                                                                                        |
| `wb.calculateFormulas()`                                  | _removed from Workbook_ — use `ExcelFormula.calculate(wb)` from `documonster/excel/formula` (see §7)                                                                                                            |
| `wb.defaultFont` / `wb.defaultFont = f`                   | `Workbook.getDefaultFont(wb)` / `Workbook.setDefaultFont(wb, f)`                                                                                                                                                |
| `wb.persons`                                              | `Workbook.listPersons(wb)`                                                                                                                                                                                      |
| `wb.externalLinks`                                        | `Workbook.listExternalLinks(wb)` (the field is still on the handle)                                                                                                                                             |
| `wb.userFunctions`                                        | `Workbook.registerFunction` / `Workbook.unregisterFunction` (the field is still on the handle)                                                                                                                  |
| `wb.removeWorksheetEx(ws)`                                | `Workbook.removeWorksheet(wb, ws)`                                                                                                                                                                              |
| `wb.media` / `wb.getImage(id)`                            | `wb.media` is still on the handle; `wb.getImage(id)` was `wb.media[Number(id)]`                                                                                                                                 |
| `wb.validateSheetName(name)`                              | no standalone check: `Workbook.addWorksheet` and `Worksheet.setName` throw `ExcelWorksheetNameError`                                                                                                            |

> **The streaming writers keep their methods, with the document model's names.** On a
> `Stream.WorkbookWriter`, `protect(pw?, opts?)` becomes `protect({ password, ...opts })` and
> `defineCellStyle(name, style)` becomes `addCellStyle(name, style)` (which returns the stored entry);
> `useBuiltInCellStyle(id)` is new. On its worksheet writer, `protect(pw?, opts?)` becomes
> `protect({ password, ...opts })`; `getWatermark()` is unchanged. The old forms still work,
> deprecated.

**Document-property fields** — `wb.title`, `wb.subject`, `wb.creator`, `wb.lastModifiedBy`, `wb.created`,
`wb.modified`, `wb.lastPrinted`, `wb.company`, `wb.manager`, `wb.category`, `wb.keywords`, `wb.description`,
`wb.language`, `wb.revision`, `wb.contentStatus` — and the settings `wb.properties`, `wb.views` and
`wb.calcProperties` are still plain fields on the handle — read and assign them directly:

```ts
wb.title = "Report";
wb.creator = "Me";
```

Protection, images, pivot tables and external links are managed through their
functions (`Workbook.protect`, `Image.add`, `Pivot.add`, `Workbook.addExternalLink`, …).

The workbook's bookkeeping members have no counterpart, because the operation that used them is the call to make:
`wb.nextId`, `wb.clearThemes()`, `wb._tableNames`, and the chart-part registry (`nextChartNumber`, `nextChartExNumber`,
`addChartEntry`, `getChartEntry`, `removeChartEntry`, `addChartExStructuredEntry`, `getChartExStructuredEntry`,
`removeChartExStructuredEntry`, `setChartStyle`, `setChartColors`, `copyChartSidecars`, `copyChartExSidecars`) —
`Chart.place`, `Chart.setStyle`, `Chart.copyTo` and `Chart.remove` do that work.

CSV / Markdown methods (`readCsv`, `writeCsv`, `readMarkdown`, …) — see §5.
PDF (`workbook`→PDF) — see §6.

### 4.2 `Worksheet`

Structure / lifecycle operations live on `Worksheet`; cell/row/column/chart/
table/image/pivot operations moved to their own namespaces.

| Old (`Worksheet` class)                                | New                                                                                                                    |
| ------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------- |
| `ws.name` (get) / `ws.name = n`                        | `Worksheet.getName(ws)` / `Worksheet.setName(ws, n)`                                                                   |
| `ws.getCell(addr)` / `ws.getCell(r, c)`                | Use `Cell.*` functions directly (see §4.3). There is no cell object to hold.                                           |
| `ws.getRow(r)`                                         | Use `Row.*` functions with the row number (see §4.4).                                                                  |
| `ws.getColumn(c)`                                      | Use `Column.*` functions with the column ref (see §4.5).                                                               |
| `ws.addRow(values, style?)`                            | `Row.add(ws, values, { copyStyle?, includeEmptyCells? })` → row number                                                 |
| `ws.addRows(values, style?)`                           | `Row.addMany(ws, values, opts?)` → row numbers                                                                         |
| `ws.getRows(start, len)`                               | Use `Row.*` / `Cell.*` with row numbers `start … start + len - 1`                                                      |
| `ws.findRow(r)` / `ws.findRows(s, l)`                  | `Row.has(ws, r)` to test, then `Row.getValues(ws, r, { base: 1 })` for what `row.values` gave (`[]` for a missing row) |
| `ws.eachRow(cb)` / `ws.eachRow(opt, cb)`               | `Row.each(ws, rowNumber => …, opt?)`                                                                                   |
| `ws.insertRow(pos, v, style?)`                         | `Row.insert(ws, pos, v, { copyStyle?, includeEmptyCells? })`                                                           |
| `ws.insertRows(pos, vs, style?)`                       | `Row.insertMany(ws, pos, vs, opts?)`                                                                                   |
| `ws.duplicateRow(n, count, insert?)`                   | `Row.duplicate(ws, n, count, { insert? })`                                                                             |
| `ws.spliceRows(start, count, ...inserts)`              | `Row.splice(ws, start, count, ...inserts)`                                                                             |
| `ws.spliceColumns(start, count, ...inserts)`           | `Column.splice(ws, start, count, ...inserts)`                                                                          |
| `ws.getSheetValues()`                                  | `Worksheet.getValues(ws, { base: 1 })` — 1-based, as excelts returned it                                               |
| `ws.mergeCells(...cells)`                              | `Worksheet.merge(ws, ...cells)`                                                                                        |
| `ws.mergeCellsWithoutStyle(...)`                       | `Worksheet.mergeWithoutStyle(ws, ...)`                                                                                 |
| `ws.unMergeCells(...)`                                 | `Worksheet.unmerge(ws, ...)`                                                                                           |
| `ws.fillFormula(range, formula, results?, shareType?)` | `Worksheet.fillFormula(ws, range, formula, { results?, shareType? })`                                                  |
| `await ws.protect(pw?, opts?)`                         | `Worksheet.protect(ws, { password, ...opts })`                                                                         |
| `ws.unprotect()`                                       | `Worksheet.unprotect(ws)`                                                                                              |
| `ws.destroy()`                                         | `Workbook.removeWorksheet(wb, ws)`                                                                                     |
| `ws.autoFitColumn(c)`                                  | `Column.autoFit(ws, c)`                                                                                                |
| `ws.autoFitColumns(s?, e?)`                            | `Column.autoFit(ws, { start?, end? })`                                                                                 |
| `ws.autoFitRow(n)`                                     | `Row.autoFit(ws, n)`                                                                                                   |
| `ws.autoFitRows(s?, e?)`                               | `Row.autoFit(ws, { start?, end? })`                                                                                    |
| `ws.addConditionalFormatting(cf)`                      | `ConditionalFormatting.add(ws, cf)`                                                                                    |
| `ws.removeConditionalFormatting(filter)` / `()`        | `ConditionalFormatting.remove(ws, filter)` / `ConditionalFormatting.clear(ws)`                                         |
| `ws.toJSON(opts?)`                                     | `Worksheet.toJson(ws, opts?)`                                                                                          |
| `ws.addJSON(data, opts?)`                              | `Worksheet.addJson(ws, data, opts?)`                                                                                   |
| `ws.toAOA()`                                           | `Worksheet.toAoa(ws)`                                                                                                  |
| `ws.addAOA(data, opts?)`                               | `Worksheet.addAoa(ws, data, opts?)`                                                                                    |
| `ws.dimensions`                                        | `Worksheet.getDimensions(ws)`                                                                                          |
| `ws.columnCount` / `ws.actualColumnCount`              | `Column.getCount(ws)` / `Column.getActualCount(ws)`                                                                    |
| `ws.rowCount` / `ws.actualRowCount`                    | `Row.getCount(ws)` / `Row.getActualCount(ws)`                                                                          |
| `ws.hasMerges`                                         | `Worksheet.hasMerges(ws)`                                                                                              |
| `ws.mergedRegions`                                     | `Worksheet.getMergedRegions(ws)`                                                                                       |
| `ws.columns` (get) / `ws.columns = defs`               | `Column.getDefinitions(ws)` / `Column.setDefinitions(ws, defs)`                                                        |
| `ws.lastColumn` / `ws.lastRow`                         | `Column.getLast(ws)` / `Row.getLast(ws)` — numbers, `undefined` when empty                                             |
| `ws.model` (get/set)                                   | `Worksheet.getModel(ws)` / `Worksheet.setModel(ws, m)`                                                                 |
| `ws.autoFilter` / `= v`                                | `Worksheet.getAutoFilter(ws)` / `Worksheet.setAutoFilter(ws, v)`                                                       |
| `ws.dataValidations`                                   | `DataValidation.*(ws, …)` (§4.10)                                                                                      |
| `ws.conditionalFormattings`                            | `ConditionalFormatting.list(ws)`                                                                                       |
| `ws.formControls`                                      | `Form.list(ws)`                                                                                                        |
| `ws.sheetProtection`                                   | `Worksheet.protect(ws, …)` / `Worksheet.unprotect(ws)` (the field is on the handle)                                    |
| `ws.rowBreaks` / `ws.colBreaks`                        | `Row.addPageBreak(ws, n)` / `Column.addPageBreak(ws, c)` (the fields are on the handle)                                |
| `ws.getColumnKey(k)` / `ws.eachColumnKey(cb)`          | `Column.getNumber(ws, k)` / `Column.getDefinitions(ws)` and each definition's `key`                                    |
| `ws.setColumnKey(k, col)`                              | `Column.setKey(ws, c, k)`                                                                                              |
| `ws.deleteColumnKey(k)`                                | no direct counterpart: give the column another key with `Column.setKey`                                                |

`ws.id`, `ws.orderNo`, `ws.state`, `ws.properties`, `ws.pageSetup`, `ws.headerFooter`, `ws.views`,
`ws.ignoredErrors` and `ws.threadedComments` are still plain fields on the handle.

> **Note the JSON/AOA casing change:** `toJSON`/`addJSON`/`toAOA`/`addAOA`
> became `toJson`/`addJson`/`toAoa`/`addAoa`.

Chart / Table / Image / Pivot / Sparkline / Form / Watermark **creation**
methods that were on `Worksheet` moved to their own namespaces (see §4.7–§4.10).

### 4.3 `Cell`

All cell operations take `(ws, addr)`. There is **no cell
object**. `addr` is `"A1"` or a 1-based `{ row, col }` position.

| Old (`cell` = `ws.getCell(...)`)                | New (`Cell.*`)                                                                                  |
| ----------------------------------------------- | ----------------------------------------------------------------------------------------------- |
| `cell.value` (get)                              | `Cell.getValue(ws, addr)`                                                                       |
| `cell.value = v`                                | `Cell.setValue(ws, addr, v)` / `Cell.setValue(ws, { row, col }, v)`                             |
| `cell.text`                                     | `Cell.getText(ws, addr)`                                                                        |
| `cell.html`                                     | `Cell.getText(ws, addr)` with `&`, `<`, `>` and `"` escaped as HTML — nothing writes it for you |
| `cell.worksheet` / `cell.workbook`              | no back-reference: keep the `ws` and `wb` you addressed the cell with                           |
| `cell.destroy()`                                | delete the call — there is no cell object to release                                            |
| `cell.displayText`                              | `Cell.getDisplayText(ws, addr)`                                                                 |
| `cell.type`                                     | `Cell.getType(ws, addr)`                                                                        |
| `cell.effectiveType`                            | `Cell.getEffectiveType(ws, addr)`                                                               |
| `cell.formula`                                  | `Cell.getFormula(ws, addr)`                                                                     |
| `cell.result` (get) / `cell.result = r`         | `Cell.getResult(ws, addr)` / `Cell.setResult(ws, addr, r)`                                      |
| `cell.style` (get)                              | `Cell.getStyle(ws, addr)`                                                                       |
| _(assign to `cell.style`)_                      | `Cell.setStyle(ws, addr, partialStyle)` / `Cell.setStyle(ws, { row, col }, style)`              |
| `cell.font` (get) / `= v`                       | `Cell.getFont(ws, addr)` / `Cell.setFont(ws, addr, v)`                                          |
| `cell.numFmt` / `= v`                           | `Cell.getNumFmt(ws, addr)` / `Cell.setNumFmt(ws, addr, v)`                                      |
| `cell.alignment` / `= v`                        | `Cell.getAlignment(ws, addr)` / `Cell.setAlignment(ws, addr, v)`                                |
| `cell.border` / `= v`                           | `Cell.getBorder(ws, addr)` / `Cell.setBorder(ws, addr, v)`                                      |
| `cell.fill` / `= v`                             | `Cell.getFill(ws, addr)` / `Cell.setFill(ws, addr, v)`                                          |
| `cell.protection` / `= v`                       | `Cell.getProtection(ws, addr)` / `Cell.setProtection(ws, addr, v)`                              |
| `cell.isMerged`                                 | `Cell.isMerged(ws, addr)`                                                                       |
| `cell.master`                                   | `Cell.getMergeMasterAddress(ws, addr)` — the master's address                                   |
| `cell.hyperlink`                                | `Cell.getHyperlink(ws, addr)`                                                                   |
| `cell.note` (get) / `= v`                       | `Cell.getNote(ws, addr)` / `Cell.setNote(ws, addr, v)`                                          |
| `cell.comment` (get) / `= v`                    | `Cell.getNote` + `Cell.getNoteAuthor` / `Cell.setNote(ws, addr, v, { author })`                 |
| `cell.dataValidation` (get) / `= v`             | `Cell.getValidation(ws, addr)` / `Cell.setValidation(ws, addr, v)`                              |
| `cell.name` / `cell.names`                      | `Cell.getNames(ws, addr)`                                                                       |
| `cell.addName(n)`                               | `Cell.addName(ws, addr, n)`                                                                     |
| `cell.removeName(n)`                            | `Cell.removeName(ws, addr, n)`                                                                  |
| `cell.removeAllNames()`                         | `Cell.clearNames(ws, addr)`                                                                     |
| _(set a single name)_                           | `Cell.setNames(ws, addr, [n])`                                                                  |
| _(set names)_                                   | `Cell.setNames(ws, addr, names)`                                                                |
| `cell.fullAddress`                              | `Cell.getFullAddress(ws, addr)`                                                                 |
| `cell.model` (get) / `= m`                      | `Cell.getModel(ws, addr)` / `Cell.setModel(ws, addr, m)`                                        |
| `cell.merge(master)` / `cell.unmerge()`         | Use `Worksheet.merge(ws, range)` / `Worksheet.unmerge(ws, range)`                               |
| `cell.isHyperlink`                              | `Cell.getHyperlink(ws, addr) !== undefined`                                                     |
| `cell.isMergedTo(master)`                       | `Cell.getMergeMasterAddress(ws, addr) === masterAddress`                                        |
| `cell.formulaType`                              | the value's shape: a `formula`, `sharedFormula` or `shareType` field                            |
| `cell.toCsvString()`                            | no counterpart; `ExcelCsv.format` writes a sheet (§5.1)                                         |
| `cell.$col$row`                                 | build `$A$1` from `Address.decodeCell(addr)` and `Address.encodeCol`                            |
| `cell.addMergeRef()` / `cell.releaseMergeRef()` | bookkeeping of `Worksheet.merge`; nothing to port                                               |

### 4.4 `Row`

All operations take `(ws, rowNumber)`.

| Old (`row` = `ws.getRow(n)`)                  | New (`Row.*`)                                                                                                  |
| --------------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| `row.height` (get) / `= h`                    | `Row.getHeight(ws, n)` / `Row.setHeight(ws, n, h)`                                                             |
| `row.hidden` (get) / `= b`                    | `Row.isHidden(ws, n)` / `Row.setHidden(ws, n, b)`                                                              |
| `row.outlineLevel` (get) / `= l`              | `Row.getOutlineLevel(ws, n)` / `Row.setOutlineLevel(ws, n, l)`                                                 |
| _(assign row style)_                          | `Row.getStyle(ws, n)` / `Row.setStyle(ws, n, style)`                                                           |
| `row.font = v`                                | `Row.setFont(ws, n, v)`                                                                                        |
| `row.alignment = v`                           | `Row.setAlignment(ws, n, v)`                                                                                   |
| `row.border = v`                              | `Row.setBorder(ws, n, v)`                                                                                      |
| `row.fill = v`                                | `Row.setFill(ws, n, v)`                                                                                        |
| `row.values` (get, 1-based)                   | `Row.getValues(ws, n, { base: 1 })` _(1-based, leading empty slot)_                                            |
| `row.getValues()` (0-based)                   | `Row.getValues(ws, n)` _(0-based dense array)_                                                                 |
| `row.values = v`                              | `Row.setValues(ws, n, v)`                                                                                      |
| `row.getCell(col)`                            | `Cell.*(ws, { row: n, col })` — `col` a number; a letter or key: `Column.getNumber(ws, col)`                   |
| `row.splice(start, count, ...v)`              | `Row.spliceCells(ws, n, start, count, ...v)` — moves values, styles and notes                                  |
| `row.eachCell(cb)` / `(opt, cb)`              | `Row.eachAddress(ws, n, (address, row, col) => …, opt?)`                                                       |
| `row.commit()`                                | delete it (edits apply at once); streamed: `Stream.commitRow(row)`                                             |
| `row.number`                                  | `n` — rows are addressed by number                                                                             |
| `row.hasValues`                               | `Row.hasValues(ws, n)`                                                                                         |
| `row.cellCount`                               | `Row.getCellCount(ws, n)`                                                                                      |
| `row.actualCellCount`                         | `Row.getActualCellCount(ws, n)`                                                                                |
| `row.dimensions`                              | `Row.getDimensions(ws, n)` _(`undefined`, not `null`)_                                                         |
| `row.collapsed`                               | `Row.isCollapsed(ws, n)`                                                                                       |
| `row.model`                                   | `Row.getModel(ws, n)` _(`undefined`, not `null`)_                                                              |
| `ws.findCell(addr)` + read                    | `Cell.getView(ws, addr)` — read-only, creates nothing                                                          |
| `row.customHeight` / `row.dyDescent`          | `Row.getModel(ws, n)` — fields of the row model                                                                |
| `row.valuesToString(sep)`                     | `Row.getValues(ws, n).join(sep)` — what it was defined as                                                      |
| `row.addPageBreak()`                          | `Row.addPageBreak(ws, n)`                                                                                      |
| `row.getCellEx(address)`                      | `Cell.*(ws, { row, col })`                                                                                     |
| `row.numFmt` / `row.protection` (get / `= v`) | `Row.getNumFmt(ws, n)` / `Row.setNumFmt(ws, n, v)`, `Row.getProtection(ws, n)` / `Row.setProtection(ws, n, v)` |
| `row.findCell(col)`                           | `Cell.has(ws, { row: n, col })`, then `Cell.getView(ws, { row: n, col })` — creates nothing                    |
| `row.worksheet`, `ws.workbook`                | no back-reference: keep the `ws` and `wb` handles you have                                                     |
| `row.destroy()`                               | delete the call — there is no row object to release                                                            |

> **`values` vs `getValues`:** `Row.getValues(ws, n, { base: 1 })` keeps Excel's **1-based** layout
> (index 0 is an empty slot; column A is index 1). `Row.getValues(ws, n)` returns a
> **0-based** dense array (column A is index 0). Old `row.values` was 1-based
> and old `row.getValues()` was 0-based — same semantics, now one function. `Row.setValues`
> takes the same `{ base }`, so a read and a write with one `base` are inverses.

**Is there still a row object?** Not in the document model. Every row function
takes the row number, and the ones that create rows — `Row.add`, `Row.addMany`,
`Row.insert`, `Row.insertMany` — answer with their numbers. Row handles exist
only in the streaming API (`Stream.RowHandle`). So a signature that used to be
`(ws: Worksheet, row: Row)` becomes:

```ts
import { Row } from "documonster/excel";
import type { Worksheet } from "documonster/excel";

const stripe = (ws: Worksheet.Handle, rowNumber: number) =>
  Row.setFill(ws, rowNumber, {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF7F7F7" }
  });
```

### 4.5 `Column`

All operations take `(ws, colRef)` where `colRef` is a key string, letter
(`"A"`), or 1-based number.

| Old (`col` = `ws.getColumn(c)`)                                                                     | New (`Column.*`)                                                                                                                                                                                          |
| --------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `col.width` (get) / `= w`                                                                           | `Column.getWidth(ws, c)` / `Column.setWidth(ws, c, w)`                                                                                                                                                    |
| `col.header` (get) / `= h`                                                                          | `Column.getHeader(ws, c)` / `Column.setHeader(ws, c, h)`                                                                                                                                                  |
| `col.key` (get) / `= k`                                                                             | `Column.getKey(ws, c)` / `Column.setKey(ws, c, k)`                                                                                                                                                        |
| `col.hidden` (get) / `= b`                                                                          | `Column.isHidden(ws, c)` / `Column.setHidden(ws, c, b)`                                                                                                                                                   |
| `col.outlineLevel` (get) / `= l`                                                                    | `Column.getOutlineLevel(ws, c)` / `Column.setOutlineLevel(ws, c, l)`                                                                                                                                      |
| `col.style` (get)                                                                                   | `Column.getStyle(ws, c)` / `Column.setStyle(ws, c, style)`                                                                                                                                                |
| `col.values` (get) / `= v`                                                                          | `Column.getValues(ws, c, { base: 1 })` / `Column.setValues(ws, c, v)`                                                                                                                                     |
| `col.eachCell(cb)`                                                                                  | `Column.eachAddress(ws, c, (address, row, col) => …, opt?)` — see below                                                                                                                                   |
| `col.bestFit`                                                                                       | `Column.getDefinition(ws, c).bestFit`                                                                                                                                                                     |
| `col.headers` / `col.headerCount`                                                                   | `Column.getHeader(ws, c)` — a string or an array; its length is the header row count                                                                                                                      |
| `col.isCustomWidth`                                                                                 | `Column.getWidth(ws, c) !== undefined`                                                                                                                                                                    |
| `col.number` / `col.letter`                                                                         | `Column.getNumber(ws, c)` / `Column.getLetter(ws, c)`                                                                                                                                                     |
| `col.addPageBreak()`                                                                                | `Column.addPageBreak(ws, c)`                                                                                                                                                                              |
| `col.font`, `col.fill`, `col.border`, `col.alignment`, `col.numFmt`, `col.protection` (get / `= v`) | `Column.getFont(ws, c)` / `Column.setFont(ws, c, v)`, and `getFill` / `setFill`, `getBorder` / `setBorder`, `getAlignment` / `setAlignment`, `getNumFmt` / `setNumFmt`, `getProtection` / `setProtection` |
| `col.collapsed`                                                                                     | `Column.isCollapsed(ws, c)`                                                                                                                                                                               |
| `col.toString()`                                                                                    | no counterpart: it printed the column for debugging; `Column.getDefinition(ws, c)` reads it                                                                                                               |
| `col.worksheet`                                                                                     | no back-reference: keep the `ws` you have                                                                                                                                                                 |
| `Column.toModel(cols)`, `col.isDefault`, `col.equivalentTo(o)`, `col.equivalentToModel(m)`          | the writer's own column-merging helpers; nothing to port — `Column.getDefinitions(ws)` reads every column                                                                                                 |

Per-facet setters and getters mirror `Cell` and `Row`: `Column.setFont` /
`getFont`, `setAlignment`, `setBorder`, `setFill`, `setNumFmt`, `setProtection`, … The
`Row` / `Column` facet getters are pure reads — they create nothing and return a copy;
`getStyle` is the live, editable object.

`Column.getValues(ws, c, { base: 1 })` is indexed by **row number**, exactly like `col.values` was, so
the index composes with the `Cell` namespace:

```ts
import { Cell, Column } from "documonster/excel";

const colNumber = Column.getNumber(ws, "total");
Column.getValues(ws, "total", { base: 1 }).forEach((value, rowNumber) => {
  if (typeof value === "number" && value < 0) {
    Cell.setFont(ws, { row: rowNumber, col: colNumber }, { color: { argb: "FFC00000" } });
  }
});
```

`Column.getValues` is the same read 0-based, and both leave the sheet untouched —
`col.eachCell` did not, because it asked every row for a cell in the column and
therefore created one.

To visit positions rather than values, `Column.eachAddress(ws, c, cb, { includeEmpty })`
is the direct counterpart of `col.eachCell`: it hands over `(address, row, col)` and
creates nothing. For the usual reason to reach for `eachCell` — applying a format
down a column — use `Column.setStyle` or `Column.setNumFmt`, which write the column
_and_ the cells already in it.

### 4.6 `Range`

The `Range` class became stateless geometry helpers. Handles are opaque
(`Range.Handle`).

| Old (`Range` class)                                     | New (`Range.*`)                                                                                                   |
| ------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| `new Range("A1:B2")` (and other ctors)                  | `Range.create("A1:B2")`                                                                                           |
| `range.contains(addr)`                                  | `Range.contains(r, addr)`                                                                                         |
| `range.containsEx(address)`                             | `Range.containsCell(r, address)`                                                                                  |
| `range.intersects(other)`                               | `Range.intersects(r, other)`                                                                                      |
| `range.forEachAddress(cb)`                              | `Range.eachAddress(r, cb)`                                                                                        |
| `range.expand(t, l, b, r)`                              | `Range.expand(r, t, l, b, right)`                                                                                 |
| `range.expandToAddress(addr)`                           | `Range.expandToAddress(r, addr)`                                                                                  |
| `range.count`                                           | `Range.getCount(r)`                                                                                               |
| `range.top` / `left` / `bottom` / `right` / `sheetName` | the same fields on the handle (`r.top`, …)                                                                        |
| `range.tl` / `range.br`                                 | `Address.encodeCell({ c: r.left - 1, r: r.top - 1 })` / `Address.encodeCell({ c: r.right - 1, r: r.bottom - 1 })` |
| `range.$t$l` / `range.$b$r`                             | `"$" + Address.encodeCol(r.left - 1) + "$" + r.top` / `"$" + Address.encodeCol(r.right - 1) + "$" + r.bottom`     |
| `range.range` / `range.toString()`                      | `Range.toString(r)` — `B2:D5`, after the sheet name it was made with (`Sheet1!B2:D5`)                             |
| `range.$range`                                          | `Range.toString(r, { absolute: true })` — `$B$2:$D$5`                                                             |
| `range.shortRange`                                      | `Range.toString(r, { compact: true })` — a one-cell range as `B2`                                                 |
| `range.$shortRange`                                     | `Range.toString(r, { absolute: true, compact: true })`                                                            |
| `range.setTLBR(…)` / `range.expandRow(row)`             | `Range.create(…)` for a new range / `Range.expand(r, …)`                                                          |

### 4.7 `Image`

| Old                                      | New (`Image.*`)                                                                         |
| ---------------------------------------- | --------------------------------------------------------------------------------------- |
| `wb.addImage(image)` → imageId           | `Image.add(wb, image)` → imageId                                                        |
| `ws.addImage(imageId, range)`            | `Image.place(ws, imageId, range)`                                                       |
| `ws.getImages()`                         | `Image.list(ws)`                                                                        |
| _(new)_                                  | `Image.remove(ws, image)`                                                               |
| `ws.addBackgroundImage(imageId)`         | `Image.setBackground(ws, imageId)`                                                      |
| `ws.getBackgroundImageId()`              | `Image.getBackground(ws)`                                                               |
| `ws.addShape(options)`                   | `Image.addShape(ws, options)`                                                           |
| `ws.getShapes()`                         | `Image.listShapes(ws)`                                                                  |
| `new Image(ws, model)` / `image.clone()` | `Image.place(ws, imageId, range)`; `Image.getModel(image)` reads one                    |
| `image.type` / `image.worksheet`         | `Image.getModel(image).type`; the sheet is the `ws` you placed it on (`Image.list(ws)`) |

### 4.8 `Chart`

Every chart-creation method on `Worksheet` became one function, `Chart.place(ws, range, options)`, which takes the
placement **before** the options and returns the chart's **handle** (the old methods returned the chart's part
number; `handle.chartNumber` — `chartExNumber` for the Office 2016 types — is that number). The per-type methods
become a `type` in the options, and presets, tables and rows have option builders. Chart-instance methods (formerly
on the `Chart` class) are also `Chart.*` functions taking a chart handle.

| Old (`ws.*` creation)                                                                                                                          | New                                                                                                                                                     |
| ---------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ws.addChart(opts, range)`                                                                                                                     | `Chart.place(ws, range, opts)`                                                                                                                          |
| `ws.addColumnChart(opts, range)`                                                                                                               | `Chart.place(ws, range, { ...opts, type: "bar", barDir: "col" })`                                                                                       |
| `ws.addBarChart(opts, range)`                                                                                                                  | `Chart.place(ws, range, { ...opts, type: "bar", barDir: "bar" })`                                                                                       |
| `ws.addLineChart(opts, range)`                                                                                                                 | `Chart.place(ws, range, { ...opts, type: "line" })`                                                                                                     |
| `ws.addAreaChart`, `addPieChart`, `addDoughnutChart`, `addScatterChart`, `addBubbleChart`, `addRadarChart`, `addStockChart`, `addSurfaceChart` | `Chart.place(ws, range, { ...opts, type })` with `type` `"area"`, `"pie"`, `"doughnut"`, `"scatter"`, `"bubble"`, `"radar"`, `"stock"`, `"surface"`     |
| `ws.addHistogramChart(opts, range)`                                                                                                            | `Chart.place(ws, range, { ...opts, type: "histogram" })`                                                                                                |
| `ws.addParetoChart`, `addWaterfallChart`, `addFunnelChart`, `addTreemapChart`, `addSunburstChart`, `addBoxWhiskerChart`, `addRegionMapChart`   | `Chart.place(ws, range, { ...opts, type })` with `type` `"pareto"`, `"waterfall"`, `"funnel"`, `"treemap"`, `"sunburst"`, `"boxWhisker"`, `"regionMap"` |
| `ws.addComboChart(opts, range)`                                                                                                                | `Chart.place(ws, range, opts)` (`opts.groups`)                                                                                                          |
| `ws.addChartEx(opts, range)`                                                                                                                   | `Chart.place(ws, range, opts)`                                                                                                                          |
| `ws.addPresetChart(preset, opts, range)`                                                                                                       | `Chart.place(ws, range, Chart.applyPreset(preset, opts))`                                                                                               |
| `ws.addPresetChartEx(...)`                                                                                                                     | `Chart.place(ws, range, Chart.applyPreset(preset, opts))`                                                                                               |
| `ws.addChartFromTable(t, opts, range)`                                                                                                         | `Chart.place(ws, range, Chart.createOptionsFromTable(ws, t, opts))`                                                                                     |
| `ws.addChartExFromTable(...)`                                                                                                                  | `Chart.place(ws, range, Chart.createOptionsFromTable(ws, t, opts))`                                                                                     |
| `ws.addChartFromRows(rows, opts, range)`                                                                                                       | `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, opts))`                                                                                   |
| `ws.addColumnChartFromRows(...)`                                                                                                               | the same, with `type: "bar", barDir: "col"` in `opts`                                                                                                   |
| `ws.addChartExFromRows(...)`                                                                                                                   | `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, opts))`                                                                                   |
| `ws.addPivotChart(pt, opts, range)`                                                                                                            | `Chart.place(ws, range, { ...opts, pivotTable: pt })`                                                                                                   |
| `ws.addPivotComboChart(pt, opts, range)`                                                                                                       | `Chart.place(ws, range, { ...opts, pivotTable: pt })` (`opts.groups`)                                                                                   |
| `ws.seriesFromColumns(opts)`                                                                                                                   | `Chart.createSeriesFromColumns(ws, opts)`                                                                                                               |
| `ws.getCharts()`                                                                                                                               | `Chart.list(ws)`                                                                                                                                        |
| `ws.removeChart(chart)`                                                                                                                        | `Chart.remove(ws, chart)`                                                                                                                               |

Chart-instance members (the old `Chart` class) map one to one:

| Old (`chart.*`)                                                                                                                                                                         | New                                                                                                                                 |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `chart.worksheet`, `chartNumber`, `chartExNumber`, `range`                                                                                                                              | the same fields on the handle                                                                                                       |
| `chart.isChartEx`                                                                                                                                                                       | `Chart.isEx(chart)`                                                                                                                 |
| `chart.model` / `chart.chartModel` / `chart.chartExModel`                                                                                                                               | `Chart.getAnchorModel(chart)` / `Chart.getModel(chart)` / `Chart.getExModel(chart)`                                                 |
| `chart.unknownElements` / `chart.userShapesXml`                                                                                                                                         | `Chart.getUnknownElements(chart)` / `Chart.getUserShapesXml(chart)`                                                                 |
| `chart.setUserShapesXml(x)` / `chart.removeUserShapes()`                                                                                                                                | `Chart.setUserShapesXml(chart, x)` / `Chart.removeUserShapes(chart)`                                                                |
| `chart.toSVG(opts)` / `chart.toPNG(opts)`                                                                                                                                               | `Chart.toSvg(chart, opts)` / `Chart.toPng(chart, opts)`                                                                             |
| `chart.title` / `chart.titleRichText` / `chart.setTitleRichText(rt)`                                                                                                                    | `Chart.getTitle(chart)` / `Chart.getTitleRichText(chart)` / `Chart.setTitle(chart, rt)`                                             |
| `chart.chartTypes` / `chart.axes` / `chart.legend` / `chart.spPr`                                                                                                                       | `Chart.getTypes(chart)` / `Chart.getAxes(chart)` / `Chart.getLegend(chart)` / `Chart.getSpPr(chart)`                                |
| `chart.categoryAxis` / `chart.valueAxis` / `chart.plotArea` / `chart.getAxis(axId)`                                                                                                     | `Chart.getCategoryAxis(chart)` / `Chart.getValueAxis(chart)` / `Chart.getPlotArea(chart)` / `Chart.getAxis(chart, axId)`            |
| `chart.mutate(fn)` / `chart.mutateChartEx(fn)`                                                                                                                                          | `Chart.mutate(chart, fn)` / `Chart.mutateEx(chart, fn)`                                                                             |
| `chart.setStyle(n)` / `chart.setBuiltInStyle(n)`                                                                                                                                        | `Chart.setStyle(chart, n)`                                                                                                          |
| `chart.addSeries`, `removeSeries`, `getSeries`, `updateSeries`, `addSeriesFromOptions`, `setSeriesValues`, `setSeriesCategories`, `setSeriesName`, `getSeriesCount`, `totalSeriesCount` | `Chart.addSeries(chart, …)` and the rest under the same names; `Chart.getTotalSeriesCount(chart)`                                   |
| `chart.copyTo(ws, range?)` / `chart.clone(range?)`                                                                                                                                      | `Chart.copyTo(chart, ws, range?)` / `Chart.copyTo(chart, chart.worksheet, range?)` — returns the copy's handle, not its part number |

Chart-handle operations (`chart` is the handle): `Chart.toSvg(chart)`,
`Chart.toPng(chart)`, `Chart.mutate(chart, fn)`, `Chart.setStyle(chart, n)`,
`Chart.addSeries(...)`, `Chart.removeSeries(...)`,
`Chart.getSeries(chart, i)`, `Chart.updateSeries(...)`, `Chart.setTitle(...)`,
`Chart.getTitle(chart)`, `Chart.getLegend(chart)`, `Chart.setLegend(...)`,
`Chart.getAxis(...)`, `Chart.getCategoryAxis(chart)`, `Chart.getValueAxis(chart)`,
`Chart.getPlotArea(chart)`, `Chart.getModel(chart)`, `Chart.getAnchorModel(chart)`,
`Chart.isEx(chart)`, `Chart.copyTo(chart, targetWs)`, etc.
(Full list: `getTitle`, `setTitle`, `getTitleRichText`, `mutate`, `mutateEx`, `setStyle`,
`addSeries`, `removeSeries`, `getSeries`, `updateSeries`, `addSeriesFromOptions`, `setSeriesValues`,
`setSeriesCategories`, `setSeriesName`, `getSeriesCount`, `getTotalSeriesCount`,
`getTypes`, `getAxes`, `getAxis`, `getCategoryAxis`, `getValueAxis`, `getPlotArea`,
`getLegend`, `setLegend`, `getSpPr`, `setSpPr`, `getUnknownElements`, `getUserShapesXml`,
`setUserShapesXml`, `removeUserShapes`, `copyTo`, `getModel`, `getExModel`, `isEx`,
`getAnchorModel`, `toSvg`, `toPng`.)

### 4.9 `Table`

| Old (`Table` class / `ws`)                                        | New (`Table.*`)                                                                                                                                                                |
| ----------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `ws.addTable(model)`                                              | `Table.add(ws, model)`                                                                                                                                                         |
| `ws.getTable(name)`                                               | `Table.get(ws, name)`                                                                                                                                                          |
| `ws.getTables()`                                                  | `Table.list(ws)`                                                                                                                                                               |
| `ws.removeTable(name)`                                            | `Table.remove(ws, name)`                                                                                                                                                       |
| `new Table(ws, model)`                                            | `Table.add(ws, model)`                                                                                                                                                         |
| `table.addRow(values, rowNumber?, opts?)`                         | `Table.addRow(t, values, rowNumber?)` — drop `{ commit }`                                                                                                                      |
| `table.removeRows(idx, count?, opts?)`                            | `Table.removeRows(t, idx, count?)` — drop `{ commit }`                                                                                                                         |
| `table.getColumn(idx)`                                            | `Table.getColumn(t, idx)` (`undefined` past the end)                                                                                                                           |
| `table.addColumn(col, values, idx?)`                              | `Table.addColumn(t, col, values, idx?)`                                                                                                                                        |
| `table.removeColumns(idx, count?)`                                | `Table.removeColumns(t, idx, count?)`                                                                                                                                          |
| `table.commit()`                                                  | delete the call — every edit applies as it is made                                                                                                                             |
| `table.model` (get/set)                                           | `Table.getModel(t)` / `Table.setModel(t, m)`                                                                                                                                   |
| `table.ref` / `= v`                                               | `Table.getRef(t)` / `Table.setRef(t, v)`                                                                                                                                       |
| `table.name` / `= v`                                              | `Table.getName(t)` / `Table.setName(t, v)`                                                                                                                                     |
| `table.displayName` / `= v`                                       | `Table.getDisplayName(t)` / `Table.setDisplayName(t, v)`                                                                                                                       |
| `table.headerRow` / `= v`                                         | `Table.hasHeaderRow(t)` / `Table.setHeaderRow(t, v)`                                                                                                                           |
| `table.totalsRow` / `= v`                                         | `Table.hasTotalsRow(t)` / `Table.setTotalsRow(t, v)`                                                                                                                           |
| `table.theme` / `= v`                                             | `Table.getTheme(t)` / `Table.setTheme(t, v)`                                                                                                                                   |
| `table.showFirstColumn` / `= v`                                   | `Table.hasFirstColumnStyle(t)` / `Table.setFirstColumnStyle(t, v)`                                                                                                             |
| `table.showLastColumn` / `= v`                                    | `Table.hasLastColumnStyle(t)` / `Table.setLastColumnStyle(t, v)`                                                                                                               |
| `table.showRowStripes` / `= v`                                    | `Table.hasRowStripes(t)` / `Table.setRowStripes(t, v)`                                                                                                                         |
| `table.showColumnStripes` / `= v`                                 | `Table.hasColumnStripes(t)` / `Table.setColumnStripes(t, v)`                                                                                                                   |
| `table.filterHeight` / `table.tableHeight` / `table.cacheState()` | the table writer's bookkeeping; `Table.getRef(t)` gives the extent                                                                                                             |
| `table.store()`                                                   | nothing to port: every `Table.*` edit writes the table's cells as it is made                                                                                                   |
| `table.width` / `table.height`                                    | `Table.getModel(t).columns.length` / the data rows' count: `Table.getModel(t).rows.length`                                                                                     |
| `table.getFormula(column)`                                        | `Table.getColumnTotalsRowFormula(Table.getColumn(t, idx)!)` — the same `SUBTOTAL(109,T[a])` for a totals function, or the custom formula; `undefined` where it answered `null` |
| `table.validate()` / `table.load(ws)` / `table.worksheet`         | nothing to port: `Table.add` / `Table.setModel` check the model, `Workbook.read` loads tables, and the sheet is the `ws` you passed                                            |

Table-column accessors (formerly on the table-column wrapper class) are on the
`Table` namespace prefixed with `column`, and take the column `Table.getColumn(t, idx)` returns:
`const col = Table.getColumn(t, idx)!`, then `Table.getColumnName(col)`,
`Table.setColumnName(col, v)`, `Table.hasColumnFilterButton`,
`Table.setColumnFilterButton`, `Table.getColumnStyle`, `Table.setColumnStyle`,
`Table.getColumnTotalsRowLabel`, `Table.setColumnTotalsRowLabel`,
`Table.getColumnTotalsRowFunction`, `Table.setColumnTotalsRowFunction`,
`Table.getColumnTotalsRowResult`, `Table.setColumnTotalsRowResult`,
`Table.getColumnTotalsRowFormula`, `Table.setColumnTotalsRowFormula`.

### 4.10 Pivot / Sparkline / Form / Watermark / DataValidation / DefinedNames / Note / Anchor / Chartsheet / Address

| Old                                                                                                                                                                    | New                                                                                                                                                                                                                                                               |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ws.addPivotTable(model)`                                                                                                                                              | `Pivot.add(ws, model)`                                                                                                                                                                                                                                            |
| `ws.pivotTables`                                                                                                                                                       | `Pivot.list(ws)`                                                                                                                                                                                                                                                  |
| `wb.pivotTables` (every sheet's)                                                                                                                                       | `Workbook.listWorksheets(wb).flatMap(ws => Pivot.list(ws))`                                                                                                                                                                                                       |
| `wb.protection`                                                                                                                                                        | `Workbook.getModel(wb).protection` reads it; `Workbook.protect` / `Workbook.unprotect` / `Workbook.verifyPassword` change and check it                                                                                                                            |
| `PivotTable` (class) / its fields                                                                                                                                      | `Pivot.Handle` — the same fields (`worksheetName`, `cacheId`, `cacheFields`, `dataFields`, …)                                                                                                                                                                     |
| `PivotTableModel`, `PivotTableValue`, `PivotTableSource`, `PivotTableChartFormat`, `PivotTableSubtotal` (types)                                                        | `Pivot.Model`, `Pivot.Value`, `Pivot.Source`, `Pivot.ChartFormat`, `Pivot.Subtotal`                                                                                                                                                                               |
| `CacheField`, `DataField`, `SharedItemValue`, `RecordValue`, `ParsedCacheDefinition` (types)                                                                           | `Pivot.CacheField`, `Pivot.DataField`, `Pivot.SharedItemValue`, `Pivot.RecordValue`, `Pivot.ParsedCacheDefinition`                                                                                                                                                |
| `ParsedCacheRecords` (type)                                                                                                                                            | no counterpart — the reader's internal shape; the records are on `Pivot.Handle`                                                                                                                                                                                   |
| _(new)_                                                                                                                                                                | `Pivot.remove(ws, pivot)`                                                                                                                                                                                                                                         |
| `ws.addSparklineGroup(opts)`                                                                                                                                           | `Sparkline.add(ws, opts)`                                                                                                                                                                                                                                         |
| `ws.getSparklineGroups()`                                                                                                                                              | `Sparkline.list(ws)`                                                                                                                                                                                                                                              |
| `ws.removeSparklineGroup(x)`                                                                                                                                           | `Sparkline.remove(ws, x)`                                                                                                                                                                                                                                         |
| `ws.addFormCheckbox(range, opts?)`                                                                                                                                     | `Form.addCheckbox(ws, range, opts?)`                                                                                                                                                                                                                              |
| `ws.getFormCheckboxes()`                                                                                                                                               | `Form.list(ws)` (a copy of the list)                                                                                                                                                                                                                              |
| _(new)_                                                                                                                                                                | `Form.remove(ws, cb)`                                                                                                                                                                                                                                             |
| `checkbox.checked` / `= b`                                                                                                                                             | `Form.isChecked(cb)` / `Form.setChecked(cb, b)`                                                                                                                                                                                                                   |
| `checkbox.link` / `= v`                                                                                                                                                | `Form.getLink(cb)` / `Form.setLink(cb, v)`                                                                                                                                                                                                                        |
| `checkbox.text` / `= v`                                                                                                                                                | `Form.getText(cb)` / `Form.setText(cb, v)`                                                                                                                                                                                                                        |
| `checkbox.getVmlAnchor()`, `getVmlStyle()`, `getVmlCheckedValue()` (and the statics)                                                                                   | the writer's own VML helpers; nothing to port                                                                                                                                                                                                                     |
| `checkbox.model` / `checkbox.worksheet`                                                                                                                                | the same fields on `Form`'s handle (`cb.model`, `cb.worksheet`)                                                                                                                                                                                                   |
| `ws.addWatermark(opts)`                                                                                                                                                | `Watermark.set(ws, opts)`                                                                                                                                                                                                                                         |
| `ws.getWatermark()`                                                                                                                                                    | `Watermark.get(ws)`                                                                                                                                                                                                                                               |
| `ws.removeWatermark()`                                                                                                                                                 | `Watermark.remove(ws)`                                                                                                                                                                                                                                            |
| `new DataValidations()`                                                                                                                                                | not needed — the functions take the sheet itself (a free-standing collection was never written)                                                                                                                                                                   |
| `dv.add(addr, rule)`                                                                                                                                                   | `DataValidation.add(ws, addr, rule)`                                                                                                                                                                                                                              |
| `dv.find(addr)` / `dv.remove(addr)`                                                                                                                                    | `DataValidation.find(ws, addr)` / `DataValidation.remove(ws, addr)` (returns `boolean`)                                                                                                                                                                           |
| `dv.model`                                                                                                                                                             | `DataValidation.list(ws)` — a copy, one `{ ref, rule }` per range                                                                                                                                                                                                 |
| `new DefinedNames()`                                                                                                                                                   | `Workbook.getDefinedNames(wb)` — the workbook's own registry (a free-standing one cannot be attached)                                                                                                                                                             |
| `dn.add(ref, name)`                                                                                                                                                    | `DefinedNames.add(dn, ref, name, { localSheetId? })` → the stored `DefinedNameModel`                                                                                                                                                                              |
| `dn.addFormula(name, formula)`                                                                                                                                         | `DefinedNames.addFormula(dn, name, formula, { localSheetId? })` → the stored `DefinedNameModel`                                                                                                                                                                   |
| `dn.remove(ref, name)` / `dn.getNames(address)`                                                                                                                        | `DefinedNames.remove(dn, ref, name)` (just those cells) or `DefinedNames.remove(dn, name, { localSheetId? })` (the whole name); both return `boolean` / `DefinedNames.getNames(dn, address)`                                                                      |
| `dn.getAllEntries()` / `dn.getAllNames()`                                                                                                                              | `DefinedNames.list(dn)` — one `DefinedNameModel` per name and scope                                                                                                                                                                                               |
| `dn.getRanges(name)` / `dn.getRangesScoped(name, localSheetId)`                                                                                                        | `DefinedNames.get(dn, name)` / `DefinedNames.get(dn, name, { localSheetId })`                                                                                                                                                                                     |
| `dn.addHidden(ref, name)`                                                                                                                                              | no counterpart — hidden names are the chart writer's (`_xlchart.*`) and are written for you                                                                                                                                                                       |
| `dn.model` / `= m`                                                                                                                                                     | `DefinedNames.getModel(dn)` / `DefinedNames.setModel(dn, m)`                                                                                                                                                                                                      |
| `dn.forEach((name, cell) => …)`                                                                                                                                        | `for (const { name, ranges } of DefinedNames.list(dn))`                                                                                                                                                                                                           |
| `dn.removeAllNames(location)`                                                                                                                                          | `Cell.clearNames(ws, addr)` for each cell of the location                                                                                                                                                                                                         |
| `dn.spliceRows(…)` / `dn.spliceColumns(…)`                                                                                                                             | nothing to call: `Row.insert` / `Row.splice` / `Column.splice` move the names they shift                                                                                                                                                                          |
| `dn.matrixMap`, `formulaMap`, `localSheetIdMap`, `hiddenMap`, `opaqueMap`, `nameForKey`, `getMatrix`, `addEx`, `removeEx`, `getNamesEx`, `normaliseMatrix`, `_explore` | the registry's internals; read through `DefinedNames.list` / `get` / `getNames`                                                                                                                                                                                   |
| `new Note(text, author)`                                                                                                                                               | `Note.create(text, { author })`                                                                                                                                                                                                                                   |
| `Note.fromModel(model)`                                                                                                                                                | `Note.fromModel(model)` _(function)_                                                                                                                                                                                                                              |
| `new Anchor(ws, "B2")` (via Image APIs)                                                                                                                                | `Anchor.create(ws, "B2")`, `Anchor.getCol(a)`, `Anchor.setRow(a, n)`, `Anchor.getModel(a)`, `Anchor.clone(a)`                                                                                                                                                     |
| `anchor.nativeCol`, `nativeRow`, `nativeColOff`, `nativeRowOff`, `colWidth`, `rowHeight`; `IAnchor` (type)                                                             | the same fields on `Anchor.Handle`; `anchor.col` / `anchor.row` are `Anchor.getCol(anchor)` / `Anchor.getRow(anchor)` — 0-based, with the offset into the cell as a fraction, as before                                                                           |
| `anchor.model` / `= m`; `anchor.worksheet`                                                                                                                             | `Anchor.getModel(anchor)` / `Anchor.setModel(anchor, m)`; `anchor.worksheet` is the same field on `Anchor.Handle`                                                                                                                                                 |
| Chartsheet handle methods                                                                                                                                              | `Chartsheet.getName(cs)`, `Chartsheet.getChart(cs)`, `Chartsheet.getModel(cs)`, `Chartsheet.setPageSetup(cs, …)`, … (management via `Workbook.addChartsheet`/`getChartsheet`/…)                                                                                   |
| `cs.sheetNo`, `chartModel`, `chartExModel`, `isChartEx`, `pageMargins`, `tabSelected`, `zoomScale`, `workbookViewId`, `zoomToFit`                                      | `Chartsheet.getSheetNo(cs)`, `getChartModel`, `getChartExModel`, `isChartEx`, `getPageMargins` / `setPageMargins`, `isTabSelected` / `setTabSelected`, `getZoomScale` / `setZoomScale`, `getWorkbookViewId` / `setWorkbookViewId`, `isZoomToFit` / `setZoomToFit` |
| `cs.rename(name)` / `cs.copy(name?)` / `cs.replaceChart(chart)`                                                                                                        | `Chartsheet.setName(cs, name)` / `Workbook.copyChartsheet(wb, cs, name?)` / `Workbook.replaceChartsheetChart(wb, cs, chart)`                                                                                                                                      |
| `cs.id`, `cs.state`, `cs.chartNumber`, `cs.chartExNumber`, `cs.model`, `cs.pageSetup`                                                                                  | `Chartsheet.getId(cs)`, `Chartsheet.getState(cs)` / `setState`, `Chartsheet.getChartNumber(cs)`, `Chartsheet.getChartExNumber(cs)`, `Chartsheet.getModel(cs)`, `Chartsheet.getPageSetup(cs)` / `setPageSetup`                                                     |
| `cs.remove()`                                                                                                                                                          | `Workbook.removeChartsheet(wb, Chartsheet.getName(cs))`                                                                                                                                                                                                           |
| `image.opacity`, `headerWidth`, `headerHeight`, `applyTo` (header/footer images)                                                                                       | the same fields on `Image.Handle`                                                                                                                                                                                                                                 |
| `decodeCell`, `encodeCol`, `decodeRange`… (old top-level named exports)                                                                                                | `Address.decodeCell`, `Address.encodeCol`, `Address.decodeRange`, `Address.encodeCell`, `Address.encodeRange`, `Address.decodeRow`, `Address.encodeRow`, `Address.quoteSheetName` — these are **no longer top-level named exports**; use the `Address` namespace  |

### 4.11 Streaming: `WorkbookWriter` and `WorkbookReader`

The streaming classes are still classes, now on the `Stream` namespace: `WorkbookWriter` →
`Stream.WorkbookWriter`, `WorkbookReader` → `Stream.WorkbookReader`, with the same options
(`filename`, `stream`, `useSharedStrings`, `useStyles`; `worksheets`, `sharedStrings`,
`hyperlinks`, `styles`). This is the one place rows and cells **are** handles: `addRow`
returns the row being written. What changed is that the row and cell handles have no methods
— you write through `Stream.*` functions and commit with `Stream.commitRow`.

| excelts 10.2.0                                | documonster                                                                          |
| --------------------------------------------- | ------------------------------------------------------------------------------------ |
| `new WorkbookWriter(opts)`                    | `new Stream.WorkbookWriter(opts)`                                                    |
| `wb.addWorksheet(name, opts)`                 | `Workbook.addWorksheet(writer, name, opts)`                                          |
| `ws.addRow(values)` → `row`                   | `sheet.addRow(values)` → the row handle (no methods)                                 |
| `row.getCell(n).value = v`                    | `Stream.setCellValue(Stream.getRowCell(row, n), v)`                                  |
| `row.font = f`                                | `Stream.setRowFont(row, f)`                                                          |
| `row.commit()`                                | `Stream.commitRow(row)`                                                              |
| `ws.commit()`                                 | `sheet.commit()` (optional — `writer.commit()` commits open sheets)                  |
| `await wb.commit()`                           | `await writer.commit()` — or, with no `filename`/`stream`, `await writer.toBuffer()` |
| `new WorkbookReader(input, opts)`             | `new Stream.WorkbookReader(input, opts)`                                             |
| `row.values` in `for await (const row of ws)` | `Stream.getRowValues(row, { base: 1 })` (1-based, as `row.values` was)               |

**Before (excelts):**

```ts before
import { WorkbookReader, WorkbookWriter } from "@cj-tech-master/excelts";

const wb = new WorkbookWriter({ filename: "big.xlsx", useSharedStrings: true });
const ws = wb.addWorksheet("Data");
for (let i = 1; i <= 3; i++) {
  const row = ws.addRow([`Row ${i}`, i]);
  row.getCell(3).value = i * 2;
  row.commit();
}
await wb.commit();

for await (const sheet of new WorkbookReader("big.xlsx", { worksheets: "emit" })) {
  for await (const row of sheet) {
    console.log(row.number, row.values);
  }
}
```

**After (documonster):**

```ts
import { Stream, Workbook } from "documonster/excel";

const writer = new Stream.WorkbookWriter({ filename: "tmp/big.xlsx", useSharedStrings: true });
const sheet = Workbook.addWorksheet(writer, "Data");
for (let i = 1; i <= 3; i++) {
  const row = sheet.addRow([`Row ${i}`, i]);
  Stream.setCellValue(Stream.getRowCell(row, 3), i * 2);
  Stream.commitRow(row);
}
await writer.commit();

const reader = new Stream.WorkbookReader("tmp/big.xlsx", { worksheets: "emit" });
for await (const ws of reader) {
  for await (const row of ws) {
    console.log(row.number, Stream.getRowValues(row, { base: 1 }));
  }
}
```

To build the file in memory instead (a browser, or an HTTP response), give the writer no
`filename` or `stream` and finish with `await writer.toBuffer()`. A streamed sheet also
accepts the positional functions while its rows are still held — `Cell.setValue(sheet, "B2", 1)`,
`Worksheet.merge(sheet, "A1:C1")`, `Column.setWidth(sheet, "A", 30)`, `DataValidation.add(sheet, "A2:A100", rule)` —
and refuses an edit to a row already written with `ExcelStreamStateError`.

---

## 5. CSV & Markdown moved off `Workbook`

The `readCsv`/`writeCsv`/`readMarkdown`/`writeMarkdown` **methods on the
`Workbook` class are gone.** They are now the `ExcelCsv` and `ExcelMarkdown`
namespaces in dedicated subpaths, taking the workbook handle as the first argument
(the reads also have a form that returns a new workbook — see below).

### 5.1 CSV — `documonster/excel/csv`

```ts
import fs from "node:fs"; // only for the file-path rows below

import { Workbook } from "documonster/excel";
import {
  ExcelCsv,
  type ExcelCsvInput,
  type ExcelCsvReadOptions,
  type ExcelCsvWriteOptions
} from "documonster/excel/csv";
```

| Old (`Workbook` method)                       | New                                                                         |
| --------------------------------------------- | --------------------------------------------------------------------------- |
| `await wb.readCsv(input, opts?)`              | `await ExcelCsv.addWorksheet(wb, input, opts?)`                             |
| `wb.writeCsv(opts?)` → string                 | `ExcelCsv.format(wb, opts?)` → string                                       |
| `await wb.writeCsv(stream, opts?)`            | `await ExcelCsv.writeStream(wb, stream, opts?)`                             |
| `await wb.writeCsvBuffer(opts?)`              | `await ExcelCsv.toBuffer(wb, opts?)`                                        |
| `wb.createCsvReadStream(opts?)`               | `ExcelCsv.toStream(wb, opts?)`                                              |
| `wb.createCsvWriteStream(opts?)`              | `ExcelCsv.createImportStream(wb, opts?)`                                    |
| `await wb.readCsvFile(path, opts?)` _(Node)_  | `await ExcelCsv.addWorksheet(wb, fs.createReadStream(path, "utf8"), opts?)` |
| `await wb.writeCsvFile(path, opts?)` _(Node)_ | `await ExcelCsv.writeFile(wb, path, opts?)`                                 |

The old methods added a worksheet to the workbook they were called on, which is what
`ExcelCsv.addWorksheet` does. If you created the workbook only to read into it, read it
directly instead: `const wb = await ExcelCsv.read(input, opts?)` and
`const wb = await ExcelCsv.readFile(path, opts?)` return a new workbook holding one worksheet.
(`ExcelCsv.read(wb, input)` and `ExcelCsv.readFile(wb, path)` still work as deprecated forms.)

The excelts `CsvInput` / `CsvOptions` types are `ExcelCsvInput` and, split by direction,
`ExcelCsvReadOptions` / `ExcelCsvWriteOptions` (`ExcelCsvOptions` extends both), all from
`documonster/excel/csv`. The unprefixed names still resolve there as deprecated aliases. The
option fields `info`, `raw` and `skipRecordsWithError` have no effect on a workbook and are
deprecated.

### 5.2 Markdown — `documonster/excel/markdown`

```ts
import { readFile } from "node:fs/promises"; // only for the file-path rows below

import { ExcelMarkdown } from "documonster/excel/markdown";
```

| Old (`Workbook` method)                     | New                                                                                    |
| ------------------------------------------- | -------------------------------------------------------------------------------------- |
| `wb.readMarkdown(input, opts?)`             | `ExcelMarkdown.addWorksheet(wb, input, opts?)`                                         |
| `wb.readMarkdownAll(input, opts?)`          | `ExcelMarkdown.addWorksheets(wb, input, opts?)`                                        |
| `wb.writeMarkdown(opts?)`                   | `ExcelMarkdown.format(wb, opts?)`                                                      |
| `wb.writeMarkdownBuffer(opts?)`             | `ExcelMarkdown.toBufferSync(wb, opts?)` (or `await ExcelMarkdown.toBuffer(wb, opts?)`) |
| `await wb.readMarkdownFile(path, opts?)`    | `ExcelMarkdown.addWorksheet(wb, await readFile(path, "utf8"), opts?)`                  |
| `await wb.readMarkdownAllFile(path, opts?)` | `ExcelMarkdown.addWorksheets(wb, await readFile(path, "utf8"), opts?)`                 |
| `await wb.writeMarkdownFile(path, opts?)`   | `await ExcelMarkdown.writeFile(wb, path, opts?)`                                       |

As with CSV, the old methods added to an existing workbook. To start from the Markdown instead,
`ExcelMarkdown.read(input, opts?)` returns a new workbook — the first table, or one worksheet per table with
`{ tables: "all" }` — and `await ExcelMarkdown.readFile(path, opts?)` does the same from a file. (The form of `read`
and `readFile` taking the workbook first still works, deprecated.)

> The worksheet options are `ExcelMarkdownReadOptions` / `ExcelMarkdownWriteOptions`, from
> `documonster/excel/markdown`. `MarkdownOptions` from `documonster/markdown` still resolves
> but is deprecated: worksheet options do not belong to the standalone Markdown module.

---

## 6. PDF: `pdf` / `excelToPdf` → `Pdf.*` namespace

All PDF value functions now live under the `Pdf` namespace. The old flat
named exports (`pdf`, `readPdf`, `excelToPdf`, `docxToPdf`,
`PdfDocumentBuilder`, `PdfPageBuilder`, `PdfEditor`, `PageSizes`) were
**removed** from the top level.

```ts
import { Pdf } from "documonster/pdf";
```

| Old (`…/pdf`)                                       | New                                                                |
| --------------------------------------------------- | ------------------------------------------------------------------ |
| `import { pdf } ...; pdf(rows)`                     | `Pdf.fromTable(rows)`                                              |
| `import { excelToPdf } ...; excelToPdf(wb, opts?)`  | `await Pdf.fromExcel(wb, opts?)`                                   |
| `import { docxToPdf } ...; docxToPdf(doc, opts?)`   | `await Pdf.fromDocx(doc, opts?)`                                   |
| `import { readPdf } ...; readPdf(bytes)`            | `Pdf.extract(bytes)`                                               |
| `new PdfDocumentBuilder()`                          | `new Pdf.Builder()`                                                |
| `PdfPageBuilder`                                    | `Pdf.PageBuilder`                                                  |
| `parseSvgPath`                                      | `Pdf.parseSvgPath`                                                 |
| `new PdfEditor()` / `PdfEditorPage`                 | `Pdf.Editor.read(bytes, { password? })` / `Pdf.EditorPage`         |
| `PageSizes`                                         | `Pdf.PageSizes`                                                    |
| `chartToPdf(chart, opts?)`                          | `await Pdf.fromChart(chart, opts?)`                                |
| `createWordChartPdfRenderer()`                      | delete it: `Pdf.fromDocx` draws Word charts by default             |
| Digital signatures (`signPdf`/`verifyPdfSignature`) | `Pdf.Signature.fillPlaceholder(...)` / `Pdf.Signature.verify(...)` |
| _(new)_ chart → PDF                                 | `await Pdf.fromChart(chartHandle, opts?)`                          |

Word charts need no renderer of their own: `Pdf.fromDocx` draws every chart, classic and
Office 2016, with the Excel chart engine by default. (`Pdf.wordChartRenderer`, which some
earlier documonster releases asked you to pass as `chartRenderer`, is deprecated — omit it.)

> **`excelToPdf` and `pdf` were removed as top-level named exports.** Use
> `Pdf.fromExcel(wb, options)` / `Pdf.fromTable(rows)`. The converters are lazily
> loaded, so a bundle that only builds PDFs (`Pdf.fromTable(...)`) never pulls in
> the excel/word object model.

`PdfError`, `PdfRenderError`, `PdfFontError`, `PdfStructureError`, `isPdfError`
remain named exports of `documonster/pdf`.

**Select sheets by name.** `Pdf.fromExcel(wb, { sheets: ["Summary", "Q1"] })`. Numeric
`sheets` entries still work but are deprecated: they are 1-based here while
`WordExcel.toDocx` counts from 0, so a number meant different sheets in the two calls.

### Recalculation with PDF export

The old `excelToPdf` auto-recalculated when the formula engine was installed.
Now recalculation is explicit — pass `ExcelFormula.calculate`:

```ts
import { Workbook } from "documonster/excel";
import { ExcelFormula } from "documonster/excel/formula";
import { Pdf } from "documonster/pdf";

const bytes = await Pdf.fromExcel(wb, { recalculate: ExcelFormula.calculate });
```

The Excel-specific callback is declared on `PdfFromExcelOptions`, not on the
generic `PdfExportOptions`. If you explicitly typed an options object, update
the annotation:

```ts
import { ExcelFormula } from "documonster/excel/formula";
import type { PdfFromExcelOptions } from "documonster/pdf";

const options: PdfFromExcelOptions = { recalculate: ExcelFormula.calculate };
```

The old `PdfExportOptions.recalculate` parameter used `never`, which accepted
incompatible callbacks and deferred the failure to runtime. That unsafe
declaration was removed rather than preserved as a
compatibility alias.

---

## 7. Formula engine: no more install step

The formula engine dropped its global install/registration mechanism. It is now
a pure functional namespace.

**Removed exports** (from `documonster/formula`):
`installFormulaEngine`, `uninstallFormulaEngine`, `createFormulaSyntaxProbe`,
and the `SyntaxProbe` type export. The `install.ts` / `host-registry.ts`
modules were deleted.

| Old (`…/formula`)                                                                      | New                                                                                              |
| -------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `import { installFormulaEngine } ...; installFormulaEngine(); wb.calculateFormulas();` | `import { ExcelFormula } from "documonster/excel/formula"; ExcelFormula.calculate(wb);`          |
| `import { calculateFormulas } ...`                                                     | `ExcelFormula.calculate` from `documonster/excel/formula`                                        |
| `import { tokenize } ...`                                                              | `Formula.tokenize`                                                                               |
| `import { parse } ...`                                                                 | `Formula.parse`                                                                                  |
| `new Workbook({ formulaSyntaxProbe })`                                                 | `Workbook.create({ formulaSyntaxProbe })` (probe factory removed; pass your own if you have one) |

`wb.calculateFormulas()` (the Workbook method) is removed. Call
`ExcelFormula.calculate(wb)` from `documonster/excel/formula` instead.
The workbook-shaped host SPI (`WorkbookLike` and friends) is gone too: the
calculation core now operates on immutable snapshots behind the Excel adapter,
rather than requiring user data to imitate a workbook object graph. Formula
errors (`FormulaError`, `FormulaParseError`, `isFormulaError`) remain exported,
and `documonster/formula` now also exports the syntax types `FormulaToken`,
`FormulaTokenType`, `FormulaAstNode` and `FormulaNodeType` (with `Formula…Token` /
`Formula…Node` for each variant). excelts' unprefixed `Token`, `TokenType`, `AstNode` and
`NodeType` still work as deprecated aliases.

The host-shaped structural types are gone with `Formula.calculate`:
`WorkbookLike`, `WorksheetLike`, `RowLike`, `CellLike`, `CellErrorValueLike`,
`FormulaResultLike`, `DefinedNameEntry`, `DefinedNamesLike`, `DimensionsLike`
and `SpillRegion` are no longer exported.

**Custom functions.** `Workbook.registerFunction` now provides typed
`FormulaValue[]` arguments. Callbacks may keep returning primitive numbers,
strings, booleans or nullish values for compatibility; tagged `FormulaValue`
results remain available when an error or array value is needed. Import the
types and tag map from `documonster/excel/formula`:

```ts
import { Workbook } from "documonster/excel";
import { FormulaValueKind } from "documonster/excel/formula";
import type { FormulaValue } from "documonster/excel/formula";

Workbook.registerFunction(wb, "ANSWER", (): FormulaValue => ({
  kind: FormulaValueKind.Number,
  value: 42
}));
```

The optional `volatile` flag remains supported. Set it when a custom function
must be evaluated afresh instead of being reused from the calculation cache.

**Before:**

```ts before
import { Workbook } from "@cj-tech-master/excelts";
import { installFormulaEngine } from "@cj-tech-master/excelts/formula";

installFormulaEngine();
const wb = new Workbook();
// ... build formulas ...
wb.calculateFormulas();
```

**After:**

```ts
import { Workbook } from "documonster/excel";
import { ExcelFormula } from "documonster/excel/formula";

const wb = Workbook.create();
// ... build formulas ...
ExcelFormula.calculate(wb);
```

---

## 8. Other modules now namespaced

These modules moved their value APIs behind a single namespace re-export
(`export * as X`). Types and error classes remain flat named exports. This
applies to the **standalone** `xml`, `markdown`, `formula`, `csv`, and `pdf`
modules — not only the Excel/Word integrations.

| Module                 | Namespace                                                 | Example                                           |
| ---------------------- | --------------------------------------------------------- | ------------------------------------------------- |
| `documonster/xml`      | `Xml`                                                     | `Xml.parse(...)`, `Xml.encode(...)`               |
| `documonster/markdown` | `Markdown`                                                | `Markdown.parse(...)`, `Markdown.format(...)`     |
| `documonster/csv`      | `Csv`                                                     | `Csv.parse(...)`, `Csv.format(...)`               |
| `documonster/formula`  | `Formula`                                                 | `Formula.tokenize(...)`, `Formula.parse(...)`     |
| `documonster/pdf`      | `Pdf`                                                     | `Pdf.fromTable(...)`, `Pdf.extract(...)`          |
| `documonster/word`     | `Document`, `Build`, `Query`, `Transform`, … (many, §8.5) | `Document.read(...)`, `Document.replaceText(...)` |

If you previously imported individual functions from these modules (e.g.
`import { parseCsv } from ".../csv"`), switch to the namespace form
(`Csv.parse`). A name not listed in this guide keeps its name inside the namespace
(`parseTar` → `Archive.parseTar`); §12 lists every such move and every deprecated forward. Members were also **renamed** (de-prefixed) inside the
namespace — see §8.2 and §8.3.

The remaining modules are namespaced too: `documonster/archive` publishes
`Archive`, `Compression` and (Node only) `Fs`; `documonster/chart` publishes
`ChartClassic`, `ChartEx` and `DrawingMl`; `documonster/draw` publishes `Draw`
and `Matrix`; `documonster/mermaid` publishes `Mermaid`. `documonster/stream`
keeps its `node:stream`-compatible names (`Readable`, `pipeline`, …) flat and
puts everything else in `Streams` (e.g.
`import { Compression } from "documonster/archive"; Compression.gzip(bytes)`).

### 8.1 Archive, stream, chart — the names that changed beyond the namespace

Most of `documonster/archive` and `documonster/stream` is the excelts API inside `Archive`, `Compression` and
`Streams`; §12 lists every such move, and every old name that still works deprecated with what replaces it. These are
the ones that changed shape:

| Old                                                                                                                                        | New                                                                                                                                                          |
| ------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `new BufferReader(bytes)` (`size`, `read`, `close`)                                                                                        | `Archive.createBufferReader(bytes)` — a `RandomAccessReader` with the same `size`, `read` and `close`                                                        |
| `new StringChunk(s)` / `new ByteChunk(bytes)` (`length`, `copy`, `toUint8Array`)                                                           | `Streams.createStringChunk(s)` / `Streams.createByteChunk(bytes)` — a `StreamDataChunk` with the same members                                                |
| `installChartSupport()` / `uninstallChartSupport()` (`…/chart`)                                                                            | delete the calls: charts work without an install step (as the formula engine does, §7)                                                                       |
| browser `Readable`, `Writable`, `Duplex`, `Transform`, `PassThrough`, `Collector`: `stream[Symbol.dispose]()`, `readable.asIndexedPairs()` | `await using` / `stream[Symbol.asyncDispose]()`; `readable.map((chunk, i) => …)` — the browser streams now have exactly Node's members, and Node has neither |

### 8.2 Standalone CSV — `documonster/csv`

The standalone CSV module is now the `Csv` namespace; **every** value export
was moved into it and most were de-prefixed. Types and errors (`CsvError`,
`CsvArgumentError`, `isCsvError`) remain flat named exports, and every flat type now
carries the `Csv` prefix: `CsvDecimalSeparator`, `CsvRow`, `CsvRowMap`, `CsvColumnConfig`,
`CsvRecordInfo`, `CsvChunkMeta`, … (the unprefixed `DecimalSeparator`, `Row`, `ColumnConfig`,
… are deprecated aliases). `CsvWorkerError` is no longer exported; catch `CsvError`.

| Old (`…/csv` named export)                       | New (`Csv.*`)                                                                    |
| ------------------------------------------------ | -------------------------------------------------------------------------------- |
| `parseCsv`                                       | `Csv.parse`                                                                      |
| `parseCsvAsync`                                  | `Csv.parseAsync`                                                                 |
| `parseCsvRows`                                   | `Csv.parseStream`                                                                |
| `parseCsvWithProgress`                           | `Csv.parseAsync(input, { onProgress })`                                          |
| `formatCsv`                                      | `Csv.format`                                                                     |
| `createCsvParserStream`                          | `new Csv.ParserStream()`                                                         |
| `createCsvFormatterStream`                       | `new Csv.FormatterStream()`                                                      |
| `csvGenerate`                                    | `Csv.generate`                                                                   |
| `csvGenerateRows`                                | `Csv.generateRows`                                                               |
| `csvGenerateAsync`                               | `Csv.generateRowsAsync`                                                          |
| `csvGenerateData`                                | `Csv.generateData`                                                               |
| `createCsvGenerator`                             | `Csv.createGenerator`                                                            |
| `formatNumberForCsv(n, sep)`                     | `Csv.formatNumber(n, { decimalSeparator })`                                      |
| `parseNumberFromCsv(s, sep)`                     | `Csv.parseNumber(s, { decimalSeparator })`                                       |
| `detectDelimiter(input, quote?, guesses?, …)`    | `Csv.detectDelimiter(input, { quote, delimitersToGuess, comment })`              |
| `detectLinebreak(input, quote?)` / `stripBom`    | `Csv.detectLinebreak(input, { quote })` / `Csv.stripBom`                         |
| `applyDynamicTyping` / `applyDynamicTypingToRow` | `Csv.applyDynamicTyping` / `Csv.applyDynamicTypingToRow(row, cfg, { castDate })` |
| `isFormattedValue` / `quoted` / `unquoted`       | `Csv.isFormattedValue` / `Csv.forceQuote` / `Csv.preventQuote`                   |

> This is the **standalone** CSV parser/formatter. It is separate from the
> Excel↔CSV bridge at `documonster/excel/csv` (§5.1), which reads/writes
> whole workbooks.

### 8.3 Standalone Markdown — `documonster/markdown`

| Old (`…/markdown` named export) | New (`Markdown.*`)  |
| ------------------------------- | ------------------- |
| `parseMarkdown`                 | `Markdown.parse`    |
| `parseMarkdownAll`              | `Markdown.parseAll` |
| `formatMarkdown`                | `Markdown.format`   |

Types (`MarkdownParseResult`, `MarkdownFormatOptions`, …)
and errors (`MarkdownError`, `MarkdownParseError`, `MarkdownArgumentError`, `isMarkdownError`) remain
flat named exports. This is separate from the Excel↔Markdown bridge at
`documonster/excel/markdown` (§5.2), whose worksheet options are `ExcelMarkdownReadOptions` /
`ExcelMarkdownWriteOptions` (the old `MarkdownOptions` here is deprecated).

### 8.4 Standalone XML — `documonster/xml`

The `Xml` namespace collects encode/decode, DOM + SAX parsing, writers, and
query. Members were de-prefixed (`xmlEncode`→`Xml.encode`,
`xmlDecode`→`Xml.decode`, `xmlEncodeAttr`→`Xml.encodeAttr`, `parseXml`→`Xml.parse`,
`parseXmlToObject`→`Xml.parseToObject`, plus `Xml.query`, `Xml.encodeCData`, …), and the writers are classes in the
namespace: `new XmlWriter()` → `new Xml.Writer()`, `new XmlStreamWriter(…)` → `new Xml.StreamWriter(…)` — with the
same methods, except `xml` → `toString()` and `save` → `snapshot()` (§12). Types and errors
(`XmlError`, `XmlParseError`, `XmlWriteError`, `XmlArgumentError`, `XmlLimitExceededError`,
`isXmlError`, `isXmlParseError`) remain flat named exports; the SAX and object-conversion
types carry the `Xml` prefix (`XmlSaxTag`, `XmlSaxHandlers`, `XmlSaxOptions`,
`XmlToPlainObjectOptions`, `XmlParseToObjectOptions`, …).

### 8.5 Word — flat functions grouped into namespaces (`documonster/word`)

Every namespace below is imported from `documonster/word` unless its row says otherwise
(`WordCrypto` is `documonster/word/crypto`). The Word module kept the same underlying functionality but reorganized its
**flat named function exports into namespaces**. Where a function's old name
already read well inside the namespace it kept its name; many were de-prefixed.
Types (`export type`), the document-model interfaces, and the error classes
stayed flat. Apply these mappings:

**Namespaces where members keep their old names** (just prefix with the
namespace): `Build.*` (all run/paragraph/table/math/shape/field builders:
`Build.text`, `Build.paragraph`, `Build.table`, `Build.heading`,
`Build.hyperlink`, `Build.field`, `Build.checkbox`, `Build.mathFraction`,
`Build.rect`, …), `Units.*` (`Units.inchesToTwips`, `Units.ptToEmu`, …),
`Query.*` (`Query.listTables`, `Query.searchText`, …). A builder's trailing positional
settings now go in an options object: `Build.field(instr, { cachedValue })` (and every
`Build.*Field`), `Build.table(rows, { ...properties, columnWidths })`,
`Build.mathFraction(num, den, { fractionType })`, `Build.mathNary(char, body, { sub, sup })`
(and `mathIntegral` / `mathSum` / `mathProduct`). The in-place edits moved to
`Document` (`Document.replaceText`, `Document.mailMerge`,
`Document.acceptAllRevisions`, …). Renamed inside `Query`: `extractText`→`getText`,
`resolveStyle`→`getParagraphStyle`, `resolveRunStyle`→`getRunStyle`,
`resolveTableStyle`→`getTableStyle`, `extractFormFields`→`listFormFields`,
`walkDocument`→`walk`, `getHeadings`→`listHeadings`, `findBookmark`→`getBookmark`,
`findComment`→`getComment`, `paragraphCount`→`countParagraphs`,
`tableCount`→`countTables`. `Convert.*`: `parseFlatOpc`→`Convert.flatOpcToParts` (or
`Document.read` for the document itself),
`toFlatOpc`→`Convert.partsToFlatOpc`, `readOdt`→`Convert.fromOdt`,
`writeOdt`→`Convert.toOdt`, `ommlToMathML`→`Convert.ommlToMathMl`,
`mathMLToOmml`→`Convert.mathMlToOmml`, `docxToSemantic`→`Convert.toSemantic`,
`mapDocument`→`Transform.map`.

**Renamed members** (name changed inside the namespace):

| Old flat export                                                    | New access path                                                                                             |
| ------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------- |
| `readDocx`                                                         | `Document.read`                                                                                             |
| `packageDocx`                                                      | `Document.toBuffer`                                                                                         |
| `mergeDocuments`                                                   | `Transform.merge`                                                                                           |
| `splitDocument`                                                    | `Transform.split`                                                                                           |
| `editDocxIncremental`                                              | `Package.edit`                                                                                              |
| `listDocxParts` / `readDocxPart`                                   | `Package.listParts` / `Package.readPart`                                                                    |
| `patchDocument` / `compileTemplate` / `patchTemplate`              | `Package.patch` / `Package.compileTemplate` / `Package.patchTemplate`                                       |
| `updateFields` / `updateTableOfContents`                           | `Transform.updateFields` / `Transform.updateTableOfContents`                                                |
| `toBuffer` / `toBase64`                                            | `Document.toBuffer` / `Document.toBase64`                                                                   |
| _(new)_ file-path IO                                               | `Document.readFile` / `Document.writeFile`                                                                  |
| `resolveThemeColor`                                                | `Theme.resolveColor`                                                                                        |
| `fillTemplate` / `fillTemplateEnhanced` / `listTemplateTags`       | `Template.fill` / `Template.fill(doc, data, { richPlaceholders: true })` / `Template.listTags`              |
| `fillTemplateFromSource` / `bindChartData`                         | `Template.fillFromSource` / `Template.bindChartData`                                                        |
| `JsonDataSource` (class)                                           | `Template.createJsonDataSource` (factory)                                                                   |
| `XmlDataSource` (class)                                            | `Template.createXmlDataSource` (factory)                                                                    |
| `CsvDataSource` (class)                                            | `Template.createCsvDataSource` (factory)                                                                    |
| `CompositeDataSource` (class)                                      | `Template.createCompositeDataSource` (factory)                                                              |
| `embedFont`                                                        | `Font.embed`                                                                                                |
| `embedFontFamily`                                                  | `Font.embedFamily`                                                                                          |
| `addEmbeddedFonts`                                                 | `Font.addEmbedded`                                                                                          |
| `subsetFont`                                                       | `Font.subset`                                                                                               |
| `shapeText` / `detectScript` / `detectDirection`                   | `Font.shapeText` / `Font.detectScript` / `Font.detectDirection`                                             |
| `createHyphenator` / `hyphenateWord` / `hyphenateText`             | `Font.createHyphenator` / `Font.hyphenateWord(…, { hyphenChar })` / `Font.hyphenateText(…, { hyphenChar })` |
| `layoutDocument`                                                   | `Layout.getPageMap`                                                                                         |
| `layoutDocumentFull`                                               | `Layout.compute`                                                                                            |
| `renderPageToSvg` / `renderDocumentToSvg` / `renderPageFromLayout` | `Layout.pageToSvg` / `Layout.toSvg` / `Layout.computedPageToSvg`                                            |
| `encryptDocx`                                                      | `WordCrypto.encrypt` — from `documonster/word/crypto`, not `documonster/word`                               |
| `protectDocument` / `unprotectDocument`                            | `Security.protect` / `Security.unprotect`                                                                   |
| `isDocumentProtected`                                              | `Security.isProtected`                                                                                      |
| `getProtectionState`                                               | `Security.getState`                                                                                         |
| `verifyProtectionPassword`                                         | `Security.verifyPassword`                                                                                   |
| `DEFAULT_SECURITY_POLICY` / `STRICT_SECURITY_POLICY`               | `Security.DEFAULT_POLICY` / `Security.STRICT_POLICY`                                                        |
| `parseStyleMap`                                                    | `Styles.parse`                                                                                              |
| `createStyleMap`                                                   | `Styles.create`                                                                                             |
| `mergeStyleMaps`                                                   | `Styles.merge`                                                                                              |
| `matchStyleMap`                                                    | `Styles.match`                                                                                              |
| `DEFAULT_STYLE_MAP`                                                | `Styles.DEFAULT`                                                                                            |
| `validateDocument`                                                 | `Validation.validate`                                                                                       |
| `diffDocuments`                                                    | `Diff.compare`                                                                                              |
| `hasVbaProject` / `getVbaProjectInfo` / `getVbaProjectData`        | `Vba.has` / `Vba.getInfo` / `Vba.getData`                                                                   |
| `addVbaProject` / `removeVbaProject` / `listVbaParts`              | `Vba.add` / `Vba.remove` / `Vba.listParts`                                                                  |
| `extractOleObjects` / `hasOleObjects` / `getOleObjectData`         | `Ole.extract` / `Ole.has` / `Ole.getData`                                                                   |
| `createOleEmbedding` / `addOleObject`                              | `Ole.createEmbedding` / `Ole.add`                                                                           |
| `createBuildingBlock` / `createGlossaryDocument`                   | `Glossary.createBlock` / `Glossary.createDocument`                                                          |
| `findBuildingBlock` / `listBuildingBlocks`                         | `Glossary.findBlock` / `Glossary.listBlocks`                                                                |
| `getAutoTextEntries` / `getQuickParts`                             | `Glossary.listAutoTextEntries` / `Glossary.listQuickParts`                                                  |
| `StreamingDocxWriter` / `createDocxStream`                         | `new Stream.DocumentWriter()`                                                                               |
| `writer.finalize()`                                                | `writer.toBuffer()` (bytes), or `writer.commit()` when writing to `options.sink`                            |
| _(new)_ streaming reader                                           | `new Stream.DocumentReader(bytes, options?)`                                                                |
| `createRenderContext` / `createIdGenerators`                       | no counterpart — the writer's own state; `Document.toBuffer` and `Stream.DocumentWriter` create it          |

**Unchanged (still flat):** the error classes (`DocxError`, `DocxParseError`,
`DocxWriteError`, `DocxMissingPartError`, `DocxInvalidStructureError`,
`DocxUnsupportedFeatureError`, `DocxEncryptedError`, `DocxDecryptionError`,
`DocxLimitExceededError`, `isDocxError`, `DocxTemplateError`) and every `export type`
(data-model types, options, layout/IR types). The integrations on their own
subpaths are namespaces: `WordHtml` (`documonster/word/html`), `WordMarkdown`
(`documonster/word/markdown`), `WordExcel` (`documonster/word/excel`) and `WordCrypto`
(`documonster/word/crypto`). PDF conversion is `Pdf.fromDocx` in `documonster/pdf`.

> **`Document`** is still imported the same way
> (`import { Document } from "documonster/word"`), but it is now a **namespace
> re-export** of the document-handle module rather than a single binding.
> Members are reached as `Document.<member>`. The document object itself is
> `DocxDocument`: `Document.create()` and `Document.read()` return the same type, so
> a document read from disk can be edited with `Document.*`, and there is no
> build step — pass it straight to `Document.toBuffer`.

### 8.6 Removed utility re-exports (previously on the root entry)

The old root `.` entry (`@cj-tech-master/excelts`) re-exported a grab-bag of
low-level helpers. These are **no longer exported from `documonster/excel`**; most are
not public anywhere, and the few that are live on another entry's namespace. Migrate as
follows:

| Old (root named export)                                                                                                                         | New                                                                                                                                                                                       |
| ----------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `decodeCell` / `encodeCol` / `decodeRow` / `encodeRow` / `encodeCell` / `decodeRange` / `encodeRange`                                           | `Address.*` from `documonster/excel` (§4.10)                                                                                                                                              |
| `xmlEncode`                                                                                                                                     | `Xml.encode` from `documonster/xml`                                                                                                                                                       |
| `xmlDecode`                                                                                                                                     | `Xml.decode` from `documonster/xml`                                                                                                                                                       |
| `xmlEncodeAttr` / `validateXmlName`                                                                                                             | `Xml.encodeAttr` / `Xml.validateName`                                                                                                                                                     |
| `createTextWatermarkImage` / `TextWatermarkImageOptions`                                                                                        | Use `Watermark.set(ws, opts)` on a worksheet (§4.10); the standalone image generator is no longer a public export                                                                         |
| `getCellDisplayText` / `formatCellValue` / `isDateDisplayFormat`                                                                                | Use `Cell.getDisplayText(ws, addr)` (§4.3); the standalone helpers are no longer public                                                                                                   |
| `dateToExcel` / `excelToDate`                                                                                                                   | Internal — no longer public. Compute Excel serial dates in your own code, or store JS `Date` values directly in cells.                                                                    |
| `base64ToUint8Array` / `uint8ArrayToBase64`                                                                                                     | Native: `Buffer.from(s, "base64")` / `Buffer.from(bytes).toString("base64")` on Node; `Uint8Array.fromBase64` / `bytes.toBase64()` where available (not Node 22)                          |
| `concatUint8Arrays` / `stringToUint8Array` / `uint8ArrayToString` / `toUint8Array`                                                              | Native: `new Uint8Array(total)` + `out.set(part, offset)` (or `Buffer.concat`), `new TextEncoder().encode`, `new TextDecoder(encoding).decode`; `toUint8Array` → `new Uint8Array(buffer)` |
| `DateParser` (`parse`, `parseAll`, `parseValid`, `iso`) / `DateFormatter` (`format`, `formatAll`, `iso`) / `getSupportedFormats` / `DateFormat` | Internal — no longer public.                                                                                                                                                              |
| `BaseError` / `toError` / `errorToJSON` / `getErrorChain` / `getRootCause`                                                                      | Internal — no longer public. Catch the module-specific error classes (`ExcelError`, `PdfError`, `XmlError`, …) instead.                                                                   |
| `CsvParserStream` / `CsvFormatterStream` (and `createCsvParserStream` / `createCsvFormatterStream`)                                             | `new Csv.ParserStream()` / `new Csv.FormatterStream()` from `documonster/csv` (§8.2)                                                                                                      |
| `MarkdownAlignment`, `MarkdownParseOptions`, `MarkdownColumnConfig` (types)                                                                     | the same names from `documonster/markdown`                                                                                                                                                |
| `PdfPageSize`, `PdfOrientation`, `PdfMargins`, `PdfColor` (types) / `PageSizeName`                                                              | the same names from `documonster/pdf` / `PdfPageSizeName` from `documonster/pdf`                                                                                                          |
| `Enums` (`Enums.ValueType`, …)                                                                                                                  | the constants themselves: `ValueType`, `FormulaType`, `ErrorValue`, `PaperSize`, … from `documonster/excel` (§8.8)                                                                        |

If you depended on any of these, replace them with the namespace equivalent
above or inline a small implementation — they were never intended as a stable
public surface and are not re-exported by the per-module entries.

### 8.7 Signature and handle changes on the `Cell` / `Column` / `Worksheet` namespaces

Things that used to need a workaround, or worked by accident:

| Old (excelts)                                                                        | Now                                                                                                                                                                                                                             |
| ------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ws.getColumn("total").number`, or `ws.columns.find(c => c.key === "total")!.number` | `Column.getNumber(ws, "total")` (and `Column.getLetter`)                                                                                                                                                                        |
| `[...ws.columns.map(c => c.defn), extra]`, or hand-copying 7 fields                  | `[...Column.getDefinitions(ws), extra]` (single column: `Column.getDefinition(ws, ref)`)                                                                                                                                        |
| `ws.eachRow` + `row.getCell` to walk one column                                      | `Column.getValues(ws, ref, { base: 1 })` (row-number indexed) / `Column.getValues(ws, ref)` (0-based) / `Column.setValues(ws, ref, v)`; neither read materialises a cell, and the writer takes everything `Cell.setValue` takes |
| `cell.value` on a handle from `row.eachCell`                                         | `Row.eachAddress(ws, n, address => Cell.getValue(ws, address))` — iterate positions, read and write through `Cell.*`                                                                                                            |
| `ws.getCell(3, 2).font = …`                                                          | `Cell.setFont(ws, { row: 3, col: 2 }, …)` — every cell function takes `"A1"` **or** `{ row, col }`                                                                                                                              |

Two things that change shape:

- **A cell is one argument.** excelts' `ws.getCell(r, c)` becomes `{ row: r, col: c }`.
  The positional `Cell.getValue(ws, r, c)` form still compiles, deprecated; an address
  mixed with a column (`Cell.getValue(ws, "A1", 99)`) is a type error.
- **Columns are not handed out.** `ws.columns[0].header = "X"` (which used to
  work, unadvertised, and bypassed the key registry and cell styles) has no
  counterpart. Declare columns with `Column.setDefinitions`, change them with
  `Column.set*`, and read them with `Column.getDefinitions`, which returns plain
  data.

`Column.getCount(ws)` is **not** `Column.getDefinitions(ws).length`: it counts
the cells actually present in rows, while `Column.getDefinitions` lists
the column _records_, which also exist for columns that were merely declared or
padded into being by a cell access. Both are documented as such now.

`Cell.getStyle`, `Row.getStyle` and `Column.getStyle` return the **live** style
record — editing it changes the cell, row or column, and reading it materialises
that cell, row or column. Read-modify-write
(`Cell.setStyle(ws, a, { ...Cell.getStyle(ws, b) })`) is a supported pattern, so
copy before editing if you do not mean to touch the source. The facet getters
(`Cell.getFont`, `getAlignment`, `getBorder`, `getFill`, `getProtection`,
`getNumFmt`, and the same on `Row` / `Column`) return a **copy** and create nothing:
change a facet with its setter (`Cell.setFont(ws, "A1", { ...Cell.getFont(ws, "A1"), bold: true })`).
`Cell.findStyle` (and `Row.findStyle` / `Column.findStyle`) returns a copy of the whole
record, or `undefined`, without materialising anything.

Cell and row **handles** belong to the streaming API (§4.11) — `Stream.setCellValue(cell, …)`
and friends take the row `addRow` returned, or a cell from `Stream.getRowCell(row, col)`.

### 8.8 Type imports (`Style`, `ColumnDefn`, `PageSetup`, …)

The old root entry did `export * from "@excel/types"`, so the whole style /
worksheet / table vocabulary was importable by name. `documonster/excel` exports
the same vocabulary, flat and under the same **declared** names — so in almost
all cases only the specifier changes:

```ts before
import type { Style, Alignment, Border, ColumnDefn, PageSetup } from "@cj-tech-master/excelts";
```

```ts
import type { Style, Alignment, Border, ColumnDefn, PageSetup } from "documonster/excel";
```

Three names changed, because the namespaces introduced in documonster took them,
or to fix a typo:

| Old type (excelts)                          | New type (documonster/excel) | Why                                                                              |
| ------------------------------------------- | ---------------------------- | -------------------------------------------------------------------------------- |
| `DataValidation` (the rule object)          | `DataValidationRule`         | `DataValidation` is now the namespace (`DataValidation.add(ws, "A1:A10", rule)`) |
| `Address` (decoded `{ address, row, col }`) | `DecodedAddress`             | `Address` is now the namespace (`Address.decodeCell(…)`)                         |
| `Cvfo`                                      | `Cfvo`                       | Spelling fix — OOXML calls it a _conditional format value object_ (`cfvo`)       |

`CellAddress` (the 0-indexed `{ c, r }` pair returned by `Address.decodeCell`),
`SheetRange` and `Origin` keep their names. `Cell.getFullAddress` now returns
`DecodedAddress` as well, so the three near-duplicate address types excelts
carried internally are down to two.

Removed, because they were unreferenced dead declarations:

| Old type                                   | Replacement                                                                                                                                                                                                                                                |
| ------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `StyleInput` / `StyleOutput`               | `Style`. The split was never wired up and `StyleOutput` was wrong: the reader resolves `numFmtId` to its **format-code string**, so a style read back from a workbook has `numFmt: string`, not `{ id, formatCode }`                                       |
| `Media`                                    | `WorkbookMedia` (what the workbook actually stores)                                                                                                                                                                                                        |
| `ImagePosition` / `ImageHyperlinkValue`    | Use `AddImageRange` — the range/position/hyperlink shape `Image.place` accepts. Note `documonster/excel` also exports `ImageAbsolutePosition` (`{ x, y }`, for absolute anchors); it is a different concept and deliberately **not** named `ImagePosition` |
| `TableColumn`                              | `Required<TableColumnProperties>`                                                                                                                                                                                                                          |
| `DefinedNamesRanges` / `DefinedNamesModel` | `DefinedNameModel[]` — one entry per defined name (`DefinedNames.getModel(dn)` returns the array)                                                                                                                                                          |
| `Buffer` (alias of `Uint8Array`)           | `Uint8Array`                                                                                                                                                                                                                                               |

Streaming types moved onto the `Stream` namespace, next to the streaming classes:

```ts before
import type {
  WorkbookWriterOptions,
  WorksheetReaderOptions,
  RowEvent
} from "@cj-tech-master/excelts";
```

```ts
import { Stream } from "documonster/excel";
type Options = Stream.WorkbookWriterOptions;
type SheetOptions = Stream.WorksheetReaderOptions;
type Row = Stream.RowEvent;
```

`NodeInput` (the streaming reader input) is now `Stream.WorkbookReaderInput` on
both entries — the Node variant still adds the file-path form.

`PaperSize` is still a value, and so are the enums — but all of them are now
plain constant objects rather than TS `enum`s, so they tree-shake away when
unused. Member access is unchanged:

```ts
import { Cell, ErrorValue, PaperSize, ValueType } from "documonster/excel";

if (Cell.getType(ws, "A1") === ValueType.Number) {
  /* … */
}
ws.pageSetup.paperSize = PaperSize.A4;
```

Two consequences of dropping `enum`:

- the **reverse mapping is gone** (`ValueType[2]` no longer yields `"Number"`);
  map the numbers yourself if you relied on it, e.g.
  `Object.entries(ValueType).find(([, v]) => v === t)?.[0]`.
- the types are literal unions now, so an arbitrary `number` is no longer
  silently assignable to `ValueType` — a cast is required where you genuinely
  have an unchecked number. `PageSetup.paperSize` stays `number`, so passing a
  code that `PaperSize` does not name (Letter = `1`, A3 = `8`, …) still works.

`ErrorValueType` is now simply `ErrorValue` (the constant object and its derived
union share one name), and it gained the two members the old type union had but
the old constant lacked (`Spill` = `#SPILL!`, `Calc` = `#CALC!`), so
`CellErrorValue.error` and `ErrorValue` are finally the same set of strings.

`RelationshipType` / `ReadingOrder` / `DocumentType` are no longer public — they
described internal OOXML plumbing.

---

### 8.9 Names that kept their name but changed type

Most of what kept its name kept its type. Every exception is listed, name by name, in
[docs/api-changes-from-excelts-10.2.0.md](docs/api-changes-from-excelts-10.2.0.md) (generated, like §12, by
`pnpm verify:excelts-guide --write`). These are the ones that change what your code does or whether it compiles:

| Change                                                                                                                                                                                                                                            | What to do                                                                                                                                                                    |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ArchiveFile#getEntry`, `readEntry` / `readEntrySync`, `readAsText` / `readAsTextSync`, `getZipEntryComment` and `sourcePath` answer `undefined` for a missing entry, not `null`                                                                  | replace `=== null` / `!== null` with `=== undefined` / `!== undefined` (or `== null`, which matches both). **A strict `=== null` check compiles and is silently never true.** |
| `HeaderFooter`'s six text fields are `string \| null`                                                                                                                                                                                             | handle `null` (no header or footer of that kind) where you read them                                                                                                          |
| `PageSetup.orientation`, `pageOrder`, `cellComments` and `errors` are the vocabularies `PageSetupOrientation`, `PageSetupOrder`, `PageSetupCellComments` and `PageSetupErrors` (`orientation` gained `"default"`, `cellComments` gained `"none"`) | an exhaustive `switch` needs the added members; a value from an untyped `string` needs checking before it is assigned                                                         |
| `Font.vertAlign` gained `"baseline"`, `Alignment.horizontal` `"general"`, `Alignment.readingOrder` `"context"`                                                                                                                                    | an exhaustive `switch` over them needs the new member                                                                                                                         |
| Byte results are `Uint8Array<ArrayBuffer>`, never a `Buffer`; byte inputs (`ImageData.buffer`, `WorkbookMedia.buffer`) take any `Uint8Array`                                                                                                      | drop `Buffer`-only calls on results (`.toString("base64")` → `Buffer.from(bytes).toString("base64")`)                                                                         |
| `documonster/stream`'s Node `Writable` is `node:stream`'s, which takes no type parameter; `Readable`, `Transform`, `PassThrough` and `Duplex` default their chunk type to `any`, as Node's do, and `IReadable<T>` to `unknown`                    | write `Writable`, not `Writable<T>`; give the chunk type where you relied on the old `Uint8Array` default                                                                     |
| `finished`, `streamToPromise` and `pipeline` called **with a callback** return what Node's do — a cleanup function, or `pipeline`'s last stream — not a promise; called without one they still return a promise                                   | do not `await` a call that passes a callback: await the callback-less form, or use the callback                                                                               |
| Option arrays are `readonly` (`sheets`, `pages`, `columns`, `headers`, `delimitersToGuess`, annotation `rect`s, …)                                                                                                                                | nothing — a mutable array is still accepted; only code that pushed into an options object it received changes                                                                 |
| `CsvParseOptions.encoding` is a `string` (was `BufferEncoding`); `TypeTransformMap.object` receives `Record<string, unknown>` (was `any`)                                                                                                         | narrow the object's values before using them                                                                                                                                  |
| Error constructors' options objects accept more fields than `BaseErrorOptions` (`field`, `where`, …)                                                                                                                                              | nothing; the renamed error classes are in §12                                                                                                                                 |
| Word document nodes (`Paragraph`, `Run`, `Table`, …) gained optional range markers; `LayoutOptions.measureText` is also called with `bold` and `italic`                                                                                           | nothing for code that builds nodes; a `measureText` you supply may use the two extra arguments                                                                                |
| Word `EncryptOptions.hashAlgorithm` / `DocumentProtectionOptions.hashAlgorithm` name the hash as Web Crypto does (`"SHA-256"`)                                                                                                                    | the compact spellings (`"SHA256"`) still work                                                                                                                                 |

## 9. Step-by-step migration checklist

Apply in order:

1. **Swap the dependency**: uninstall `@cj-tech-master/excelts`, install
   `documonster@^0.22.0`.
2. **Global find/replace import specifiers** using the §2 table. In particular:
   - `from "@cj-tech-master/excelts"` → `from "documonster/excel"`
   - `from "@cj-tech-master/excelts/browser"` → `from "documonster/excel"`
   - `@cj-tech-master/excelts/<sub>` → `documonster/<sub>`
3. **Rewrite Excel object construction**: `new Workbook()` → `Workbook.create()`.
   The other Excel domain types are created by the call that attaches them, so
   nothing is built detached and forgotten: `new Range(…)` → `Range.create(…)`,
   `new Note(text, author)` → `Note.create(text, { author })`, `new Table(ws, m)` →
   `Table.add(ws, m)`, `new Image(…)` → `Image.place(ws, imageId, range)`,
   `new DataValidations()` → pass the sheet to `DataValidation.add(ws, …)`.
   (`Table.create`, `Image.create` and `DataValidation.create` exist but are
   deprecated: what they build is never written.)
4. **Rewrite every Excel method/property access** using §4 tables. The pattern
   is always: `obj.method(args)` → `Namespace.method(obj, args)` and
   `obj.prop` / `obj.prop = v` → `Namespace.getProp(obj)` /
   `Namespace.setProp(obj, v)`. Watch for renamed members
   (`getSheetValues`→`getValues`, `mergeCells`→`merge`, `toJSON`→`toJson`, …).
   A cell is **one** argument, `"A1"` or `{ row, col }` — never `(row, col)` — and
   trailing settings (a password, an author, results, a share type) go in an
   options object.
5. **Replace `wb.xlsx.*` I/O** with `Workbook.read` / `Workbook.readFile` /
   `Workbook.toBuffer` / `Workbook.writeFile` (§3.3). Reading returns the
   handle: `const wb = await Workbook.read(bytes);`.
6. **Move CSV/Markdown calls** to `documonster/excel/csv` /
   `documonster/excel/markdown` (`ExcelCsv.*` / `ExcelMarkdown.*`), passing `wb` first (§5).
7. **Replace `excelToPdf(wb, opts)`** with `await Pdf.fromExcel(wb, opts)` and
   drop any `pdf`/`excelToPdf` import in favor of `Pdf.*` (§6).
8. **Remove `installFormulaEngine()`** and replace `wb.calculateFormulas()`
   with `ExcelFormula.calculate(wb)` from `documonster/excel/formula` (§7).
9. **Keep document-property assignments** (`wb.title = …`, `wb.creator = …`)
   as they are; other workbook-level state goes through its functions (§4.1).
10. **Update standalone-module calls** to their namespaces: `parseCsv`→`Csv.parse`,
    `formatMarkdown`→`Markdown.format`, `xmlEncode`→`Xml.encode`, and every Word
    flat function → its namespace (`readDocx`→`Document.read`, `replaceText`→
    `Document.replaceText`, `layoutDocument`→`Layout.getPageMap`, `encryptDocx`→
    `WordCrypto.encrypt`, …) per §8.2–§8.5.
11. **Type-check.** Because the API is now function-based, TypeScript will flag
    every remaining `.method`/`.prop` on a handle and every missing named export
    — use those errors as a to-do list until the project compiles.
12. **Clear the deprecations.** Turn on a deprecation lint (`@typescript-eslint/no-deprecated`
    or your editor's strike-through) and replace each hit with the name its TSDoc gives;
    every deprecated name is removed in 1.0. [MIGRATION.md](MIGRATION.md) lists them all.
13. **Re-check `null` tests.** Absence is `undefined` everywhere (`Watermark.get`,
    `Table.get`, the archive readers, …); TypeScript does not flag `=== null` against
    such a result, it just never matches.

**Done?** The migration is finished when all of these hold:

- [ ] `tsc --noEmit` is clean.
- [ ] No deprecated name remains: `@typescript-eslint/no-deprecated` reports nothing (no strike-through in the editor).
- [ ] `grep -rn "=== null\|!== null" src` shows no check against a documonster result (absence is `undefined`).
- [ ] `grep -rn "Buffer" src` shows no `Buffer`-only call on a documonster byte result (`.toString("base64")`, `.slice()` as a view).
- [ ] A real production file round-trips: `Workbook.readFile` → your code → `Workbook.writeFile`, then opens in Excel
      with the expected values, styles and formula results.

---

## 10. Complete before/after example

**Before (excelts):**

```ts before
import { Workbook } from "@cj-tech-master/excelts";
import { installFormulaEngine } from "@cj-tech-master/excelts/formula";
import { excelToPdf } from "@cj-tech-master/excelts/pdf";

installFormulaEngine();

const wb = new Workbook();
wb.creator = "Reports Bot";

const ws = wb.addWorksheet("Sales");
ws.columns = [
  { header: "Product", key: "product", width: 20 },
  { header: "Revenue", key: "revenue", width: 15 }
];
ws.addRow({ product: "Widget", revenue: 1000 });
ws.addRow({ product: "Gadget", revenue: 2500 });
ws.getCell("B4").value = 42;
ws.getCell("B4").font = { bold: true };
ws.getCell("C1").value = { formula: "SUM(B2:B3)" };
ws.mergeCells("A6:B6");

wb.calculateFormulas();

await wb.xlsx.writeFile("sales.xlsx");
const pdfBytes = await excelToPdf(wb);

const csv = wb.writeCsv();
```

**After (documonster):**

```ts
import { Cell, Column, Row, Workbook, Worksheet } from "documonster/excel";
import { ExcelCsv } from "documonster/excel/csv";
import { ExcelFormula } from "documonster/excel/formula";
import { Pdf } from "documonster/pdf";

const wb = Workbook.create();
wb.creator = "Reports Bot";

const ws = Workbook.addWorksheet(wb, "Sales");
Column.setDefinitions(ws, [
  { header: "Product", key: "product", width: 20 },
  { header: "Revenue", key: "revenue", width: 15 }
]);
Row.add(ws, { product: "Widget", revenue: 1000 });
Row.add(ws, { product: "Gadget", revenue: 2500 });
Cell.setValue(ws, "B4", 42);
Cell.setFont(ws, "B4", { bold: true });
Cell.setValue(ws, "C1", { formula: "SUM(B2:B3)" });
Worksheet.merge(ws, "A6:B6");

ExcelFormula.calculate(wb);

await Workbook.writeFile(wb, "sales.xlsx");
const pdfBytes = await Pdf.fromExcel(wb);

const csv = ExcelCsv.format(wb);
```

The workbook is recalculated once, before both exports. If you only export a PDF,
you can let the exporter do it instead — `Pdf.fromExcel(wb, { recalculate: ExcelFormula.calculate })`
— and drop the separate `ExcelFormula.calculate(wb)` call (§6).

---

## 11. Quick reference: what was removed

- Package `@cj-tech-master/excelts` (renamed to `documonster`).
- Root `.` export and `./browser` export.
- All Excel domain **classes**: `Workbook`, `Worksheet`, `Cell`, `Row`,
  `Column`, `Range`, `Table`, `Image`, `Note`, `FormCheckbox`,
  `DataValidations`, `DefinedNames`, `Chartsheet`, `Chart` (as classes).
- `Workbook.xlsx` I/O accessor.
- `Workbook` CSV methods (`readCsv`, `writeCsv`, `writeCsvBuffer`,
  `readCsvFile`, `writeCsvFile`, `createCsvReadStream`, `createCsvWriteStream`)
  and Markdown methods (`readMarkdown`, `readMarkdownAll`, `writeMarkdown`,
  `writeMarkdownBuffer`, `readMarkdownFile`, `readMarkdownAllFile`,
  `writeMarkdownFile`).
- `Workbook.calculateFormulas()`.
- Top-level `excelToPdf` / `pdf` / `readPdf` / `PageSizes` /
  `PdfDocumentBuilder` / `PdfEditor` named exports (→ `Pdf.*`).
- Formula `installFormulaEngine`, `uninstallFormulaEngine`,
  `createFormulaSyntaxProbe`, and the `SyntaxProbe` export.
- `Formula.calculate` (workbook recalculation now lives at
  `documonster/excel/formula`).
- The structural host types `WorkbookLike`, `WorksheetLike`, `RowLike`,
  `CellLike`, `CellErrorValueLike`, `FormulaResultLike`, `DefinedNameEntry`,
  `DefinedNamesLike`, `DimensionsLike`, `SpillRegion`.
- Root-entry utility re-exports: `decodeCell`/`encodeCol`/… (→ `Address.*`),
  `xmlEncode`/`xmlDecode`/… (→ `Xml.*`), `createTextWatermarkImage`,
  `getCellDisplayText`/`formatCellValue`, `dateToExcel`/`excelToDate`,
  `base64ToUint8Array`/`uint8ArrayToBase64`, `concatUint8Arrays`/`toUint8Array`/
  `stringToUint8Array`/`uint8ArrayToString`, `DateParser`/`DateFormatter`,
  `BaseError`/`toError`/`errorToJSON`/`getErrorChain`/`getRootCause` (§8.6).

Everything above has a functional replacement documented in this guide.

---

## 12. Name index (generated)

Every name `@cj-tech-master/excelts` 10.2.0 exports that `documonster` publishes under another name, generated by
`pnpm verify:excelts-guide --write` from the two packages' declarations. A name in neither this index nor the sections
above is unchanged: import it from the `documonster` entry in §2. Two kinds of row appear here:

- **the same name inside a namespace** — `WorksheetWriter` is `Stream.WorksheetWriter`, `parseTar` is
  `Archive.parseTar`;
- **a name `documonster` still exports but has deprecated** — it works until 1.0, and the row gives what replaces it,
  from the deprecation itself; `—` means no one name does, and the note says what to do instead.

<!-- excelts-index:start -->

**`@cj-tech-master/excelts`** → `documonster/excel`

| excelts                                       | documonster                             | Note                                                                                                                                                                        |
| --------------------------------------------- | --------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `AddAOAOptions`                               | `AddAoaOptions`                         | `AddAOAOptions` is deprecated: Use `AddAoaOptions`. Removed in 1.0.                                                                                                         |
| `AddJSONOptions`                              | `AddJsonOptions`                        | `AddJSONOptions` is deprecated: Use `AddJsonOptions`. Removed in 1.0.                                                                                                       |
| `CacheField`                                  | `Pivot.CacheField`                      |                                                                                                                                                                             |
| `ColumnOutOfBoundsError`                      | `ExcelColumnOutOfBoundsError`           | `ColumnOutOfBoundsError` is deprecated: Use `ExcelColumnOutOfBoundsError`. Removed in 1.0.                                                                                  |
| `DataField`                                   | `Pivot.DataField`                       |                                                                                                                                                                             |
| `decodeCell`                                  | `Address.decodeCell`                    |                                                                                                                                                                             |
| `decodeCol`                                   | `Address.decodeCol`                     |                                                                                                                                                                             |
| `decodeRange`                                 | `Address.decodeRange`                   |                                                                                                                                                                             |
| `decodeRow`                                   | `Address.decodeRow`                     |                                                                                                                                                                             |
| `encodeCell`                                  | `Address.encodeCell`                    |                                                                                                                                                                             |
| `encodeCol`                                   | `Address.encodeCol`                     |                                                                                                                                                                             |
| `encodeRange`                                 | `Address.encodeRange`                   |                                                                                                                                                                             |
| `encodeRow`                                   | `Address.encodeRow`                     |                                                                                                                                                                             |
| `ExcelNotSupportedError`                      | `ExcelUnsupportedFeatureError`          | `ExcelNotSupportedError` is deprecated: Use `ExcelUnsupportedFeatureError` — the same class, named as Word's `DocxUnsupportedFeatureError` is. Removed in 1.0.              |
| `HyperlinkEvent`                              | `Stream.HyperlinkEvent`                 |                                                                                                                                                                             |
| `HyperlinksEvent`                             | `Stream.HyperlinksEvent`                |                                                                                                                                                                             |
| `ImageError`                                  | `ExcelImageError`                       | `ImageError` is deprecated: Use `ExcelImageError`. Removed in 1.0.                                                                                                          |
| `InvalidAddressError`                         | `ExcelInvalidAddressError`              | `InvalidAddressError` is deprecated: Use `ExcelInvalidAddressError`. Removed in 1.0.                                                                                        |
| `InvalidValueTypeError`                       | `ExcelInvalidValueTypeError`            | `InvalidValueTypeError` is deprecated: Use `ExcelInvalidValueTypeError`. Removed in 1.0.                                                                                    |
| `MaxItemsExceededError`                       | `ExcelLimitExceededError`               | `MaxItemsExceededError` is deprecated: Use `ExcelLimitExceededError`. Removed in 1.0.                                                                                       |
| `MaxItemsExceededError#maxItems`              | `MaxItemsExceededError#allowed`         | deprecated: Use `allowed`. Removed in 1.0.                                                                                                                                  |
| `MergeConflictError`                          | `ExcelMergeConflictError`               | `MergeConflictError` is deprecated: Use `ExcelMergeConflictError`. Removed in 1.0.                                                                                          |
| `ParsedCacheDefinition`                       | `Pivot.ParsedCacheDefinition`           |                                                                                                                                                                             |
| `ParseEvent`                                  | `Stream.ParseEvent`                     |                                                                                                                                                                             |
| `PivotTableError`                             | `ExcelPivotTableError`                  | `PivotTableError` is deprecated: Use `ExcelPivotTableError`. Removed in 1.0.                                                                                                |
| `RecordValue`                                 | `Pivot.RecordValue`                     |                                                                                                                                                                             |
| `RowEvent`                                    | `Stream.RowEvent`                       |                                                                                                                                                                             |
| `RowOutOfBoundsError`                         | `ExcelRowOutOfBoundsError`              | `RowOutOfBoundsError` is deprecated: Use `ExcelRowOutOfBoundsError`. Removed in 1.0.                                                                                        |
| `SharedItemValue`                             | `Pivot.SharedItemValue`                 |                                                                                                                                                                             |
| `SharedStringEvent`                           | `Stream.SharedStringEvent`              |                                                                                                                                                                             |
| `SheetToJSONOptions`                          | `SheetToJsonOptions`                    | `SheetToJSONOptions` is deprecated: Use `SheetToJsonOptions`. Removed in 1.0.                                                                                               |
| `TableError`                                  | `ExcelTableError`                       | `TableError` is deprecated: Use `ExcelTableError`. Removed in 1.0.                                                                                                          |
| `WorkbookReader`                              | `Stream.WorkbookReader`                 |                                                                                                                                                                             |
| `WorkbookReaderOptions`                       | `Stream.WorkbookReaderOptions`          |                                                                                                                                                                             |
| `WorkbookWriter`                              | `Stream.WorkbookWriter`                 |                                                                                                                                                                             |
| `WorkbookWriter#addApp`                       | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addContentTypes`              | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addCore`                      | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addFeaturePropertyBag`        | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addImage`                     | `Stream.WorkbookWriter#add`             | deprecated: Use `Image.add`: `Image.add(wb, image)`, then `Image.place(ws, id, range)`. Removed in 1.0.                                                                     |
| `WorkbookWriter#addMedia`                     | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes the media a second time. Use `commit`. Removed in 1.0. |
| `WorkbookWriter#addMetadata`                  | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addOfficeRels`                | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addSharedStrings`             | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addStyles`                    | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addThemes`                    | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addWorkbook`                  | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addWorkbookRels`              | `Stream.WorkbookWriter#commit`          | deprecated: One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.  |
| `WorkbookWriter#addWorksheet`                 | `Stream.WorkbookWriter#addWorksheet`    | deprecated: Use `Workbook.addWorksheet`: `Workbook.addWorksheet(wb, name, options)`. Removed in 1.0.                                                                        |
| `WorkbookWriter#commentRefs`                  | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#compressionLevel`             | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#definedNames`                 | `Stream.WorkbookWriter#getDefinedNames` | deprecated: Use `Workbook.getDefinedNames`: `Workbook.getDefinedNames(wb)`. Removed in 1.0.                                                                                 |
| `WorkbookWriter#dynamicArrayCount`            | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#getWorksheet`                 | `Stream.WorkbookWriter#getWorksheet`    | deprecated: Use `Workbook.getWorksheet`: `Workbook.getWorksheet(wb, id)`. Removed in 1.0.                                                                                   |
| `WorkbookWriter#media`                        | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#nextId`                       | —                                       | deprecated: The id `addWorksheet` gives the next sheet; nothing outside the writer needs it. Removed in 1.0.                                                                |
| `WorkbookWriter#promise`                      | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#protect`                      | `Stream.WorkbookWriter#protect`         | deprecated: Use `Workbook.protect`: `Workbook.protect(wb, { password, ...flags })`. Removed in 1.0.                                                                         |
| `WorkbookWriter#sharedStrings`                | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#stream`                       | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#styles`                       | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#unprotect`                    | `Stream.WorkbookWriter#unprotect`       | deprecated: Use `Workbook.unprotect`. Removed in 1.0.                                                                                                                       |
| `WorkbookWriter#zip`                          | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriter#zipOptions`                   | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorkbookWriterOptions`                       | `Stream.WorkbookWriterOptions`          |                                                                                                                                                                             |
| `WorkbookZipOptions`                          | `Stream.WorkbookZipOptions`             |                                                                                                                                                                             |
| `WorksheetEvent`                              | `Stream.WorksheetEvent`                 |                                                                                                                                                                             |
| `WorksheetHyperlink`                          | `Stream.WorksheetHyperlink`             |                                                                                                                                                                             |
| `WorksheetNameError`                          | `ExcelWorksheetNameError`               | `WorksheetNameError` is deprecated: Use `ExcelWorksheetNameError`. Removed in 1.0.                                                                                          |
| `WorksheetReader`                             | `Stream.WorksheetReader`                |                                                                                                                                                                             |
| `WorksheetReaderOptions`                      | `Stream.WorksheetReaderOptions`         |                                                                                                                                                                             |
| `WorksheetReadyEvent`                         | `Stream.WorksheetReadyEvent`            |                                                                                                                                                                             |
| `WorksheetWriter`                             | `Stream.WorksheetWriter`                |                                                                                                                                                                             |
| `WorksheetWriter#addBackgroundImage`          | `Stream.WorksheetWriter#setBackground`  | deprecated: Use `Image.setBackground`: `Image.setBackground(ws, imageId)`. Removed in 1.0.                                                                                  |
| `WorksheetWriter#addConditionalFormatting`    | `Stream.WorksheetWriter#add`            | deprecated: Use `ConditionalFormatting.add`: `ConditionalFormatting.add(ws, cf)`. Removed in 1.0.                                                                           |
| `WorksheetWriter#addImage`                    | `Stream.WorksheetWriter#place`          | deprecated: Use `Image.place`: `Image.place(ws, imageId, range)`. Removed in 1.0.                                                                                           |
| `WorksheetWriter#addWatermark`                | `Stream.WorksheetWriter#set`            | deprecated: Use `Watermark.set`: `Watermark.set(ws, options)`. Removed in 1.0.                                                                                              |
| `WorksheetWriter#drawing`                     | —                                       | deprecated: The workbook writer's view of the sheet while it writes the package; nothing outside the writer needs it. `getImages()` lists what was added. Removed in 1.0.   |
| `WorksheetWriter#getBackgroundImageId`        | `Stream.WorksheetWriter#getBackground`  | deprecated: Use `Image.getBackground`, which answers with the id as a string (`"0"`), as it does for a document-model sheet. Removed in 1.0.                                |
| `WorksheetWriter#getImages`                   | `Stream.WorksheetWriter#list`           | deprecated: Use `Image.list`, which lists the placed pictures without the entry an overlay watermark adds. Removed in 1.0.                                                  |
| `WorksheetWriter#getWatermark`                | `Stream.WorksheetWriter#get`            | deprecated: Use `Watermark.get`, which answers with a copy. Removed in 1.0.                                                                                                 |
| `WorksheetWriter#hasComments`                 | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorksheetWriter#mergeCells`                  | `Stream.WorksheetWriter#merge`          | deprecated: Use `Worksheet.merge`: `Worksheet.merge(ws, "A1:B2")`. Removed in 1.0.                                                                                          |
| `WorksheetWriter#protect`                     | `Stream.WorksheetWriter#protect`        | deprecated: Use `Worksheet.protect`: `Worksheet.protect(ws, { password, ...flags })`. Removed in 1.0.                                                                       |
| `WorksheetWriter#removeConditionalFormatting` | `Stream.WorksheetWriter#remove`         | deprecated: Use `ConditionalFormatting.remove` (one block) or `ConditionalFormatting.clear` (all). Removed in 1.0.                                                          |
| `WorksheetWriter#removeWatermark`             | `Stream.WorksheetWriter#remove`         | deprecated: Use `Watermark.remove`. Removed in 1.0.                                                                                                                         |
| `WorksheetWriter#rId`                         | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorksheetWriter#startedData`                 | —                                       | deprecated: Internal; not part of the public API. Removed in 1.0.                                                                                                           |
| `WorksheetWriter#stream`                      | —                                       | deprecated: The sheet part's output while the writer encodes it; writing to it directly corrupts the part, and nothing outside the writer needs it. Removed in 1.0.         |
| `WorksheetWriter#unprotect`                   | `Stream.WorksheetWriter#unprotect`      | deprecated: Use `Worksheet.unprotect`. Removed in 1.0.                                                                                                                      |
| `ZlibOptions`                                 | `Stream.ZlibOptions`                    |                                                                                                                                                                             |

**`@cj-tech-master/excelts/archive`** → `documonster/archive`

| excelts                              | documonster                                           | Note                                                                                                                                                                                                                           |
| ------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `AES_AUTH_CODE_LENGTH`               | `Archive.Crypto.AES_AUTH_CODE_LENGTH`                 |                                                                                                                                                                                                                                |
| `AES_EXTRA_FIELD_ID`                 | `Archive.Crypto.AES_EXTRA_FIELD_ID`                   |                                                                                                                                                                                                                                |
| `AES_KEY_LENGTH`                     | `Archive.Crypto.AES_KEY_LENGTH`                       |                                                                                                                                                                                                                                |
| `AES_PASSWORD_VERIFY_LENGTH`         | `Archive.Crypto.AES_PASSWORD_VERIFY_LENGTH`           |                                                                                                                                                                                                                                |
| `AES_SALT_LENGTH`                    | `Archive.Crypto.AES_SALT_LENGTH`                      |                                                                                                                                                                                                                                |
| `AES_VENDOR_ID`                      | `Archive.Crypto.AES_VENDOR_ID`                        |                                                                                                                                                                                                                                |
| `AES_VERSION_AE1`                    | `Archive.Crypto.AES_VERSION_AE1`                      |                                                                                                                                                                                                                                |
| `AES_VERSION_AE2`                    | `Archive.Crypto.AES_VERSION_AE2`                      |                                                                                                                                                                                                                                |
| `aesDecrypt`                         | `Archive.Crypto.decryptAes`                           | `Archive.aesDecrypt` is deprecated: Use `Crypto.decryptAes`. Removed in 1.0.                                                                                                                                                   |
| `aesEncrypt`                         | `Archive.Crypto.encryptAes`                           | `Archive.aesEncrypt` is deprecated: Use `Crypto.encryptAes`. Removed in 1.0.                                                                                                                                                   |
| `aesEncryptedSize`                   | `Archive.Crypto.getAesEncryptedSize`                  | `Archive.aesEncryptedSize` is deprecated: Use `Crypto.getAesEncryptedSize`. Removed in 1.0.                                                                                                                                    |
| `ArchiveFile`                        | `Archive.ArchiveFile`                                 |                                                                                                                                                                                                                                |
| `ArchiveFile#addZipComment`          | `Archive.ArchiveFile#setZipComment`                   | deprecated: Use `setZipComment`: this sets the comment, it does not add one. Removed in 1.0.                                                                                                                                   |
| `ArchiveFile#bytes`                  | `Archive.ArchiveFile#toBuffer`                        | deprecated: Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                    |
| `ArchiveFile#bytesSync`              | `Archive.ArchiveFile#toBufferSync`                    | deprecated: Use `toBufferSync`. Removed in 1.0.                                                                                                                                                                                |
| `ArchiveFile#getEntryCountAsync`     | `Archive.ArchiveFile#entryCount`                      | deprecated: Use `entryCount`, which now counts the entries of an opened TAR too. Removed in 1.0.                                                                                                                               |
| `ArchiveFile#symlink`                | `Archive.ArchiveFile#addSymlink`                      | deprecated: Use `addSymlink`, which takes the mode as `{ mode }`. Removed in 1.0.                                                                                                                                              |
| `ArchiveFile#writeToFile`            | `Archive.ArchiveFile#writeFile`                       | deprecated: Use `writeFile`. Removed in 1.0.                                                                                                                                                                                   |
| `ArchiveFile#writeToFileSync`        | `Archive.ArchiveFile#writeFileSync`                   | deprecated: Use `writeFileSync`. Removed in 1.0.                                                                                                                                                                               |
| `base64ToUint8Array`                 | —                                                     | `Archive.base64ToUint8Array` is deprecated: Not an archive operation. Use `Bytes.fromBase64` from `documonster/utils`, which reads base64 the same way on every platform. Removed in 1.0.                                      |
| `buildAesExtraField`                 | `Archive.Crypto.createAesExtraField`                  | `Archive.buildAesExtraField` is deprecated: Use `Crypto.createAesExtraField`. Removed in 1.0.                                                                                                                                  |
| `compress`                           | `Compression.deflate`                                 | `compress` is deprecated: Use `Compression.deflate`. Removed in 1.0.                                                                                                                                                           |
| `COMPRESSION_METHOD_AES`             | `Archive.Crypto.COMPRESSION_METHOD_AES`               |                                                                                                                                                                                                                                |
| `compressSync`                       | `Compression.deflateSync`                             | `compressSync` is deprecated: Use `Compression.deflateSync`. Removed in 1.0.                                                                                                                                                   |
| `concatUint8Arrays`                  | —                                                     | `Archive.concatUint8Arrays` is deprecated: Not an archive operation. Use `Bytes.concat` from `documonster/utils`. Removed in 1.0.                                                                                              |
| `copyFile`                           | `Fs.copyFile`                                         |                                                                                                                                                                                                                                |
| `copyFileSync`                       | `Fs.copyFileSync`                                     |                                                                                                                                                                                                                                |
| `crc32`                              | `Archive.crc32`                                       |                                                                                                                                                                                                                                |
| `crc32Finalize`                      | `Archive.crc32Finalize`                               |                                                                                                                                                                                                                                |
| `Crc32MismatchError`                 | `ZipCrc32MismatchError`                               | `Crc32MismatchError` is deprecated: Use `ZipCrc32MismatchError`. Removed in 1.0.                                                                                                                                               |
| `crc32Update`                        | `Archive.crc32Update`                                 |                                                                                                                                                                                                                                |
| `createAbortError`                   | `new AbortError(reason)`                              | `Archive.createAbortError` is deprecated: Use `new AbortError(reason)` (a flat export of `documonster/archive` and `documonster/stream`); this one returns `reason` itself when it already is an `AbortError`. Removed in 1.0. |
| `createDeflateStream`                | `Compression.createDeflateStream`                     |                                                                                                                                                                                                                                |
| `createGunzipStream`                 | `Compression.createGunzipStream`                      |                                                                                                                                                                                                                                |
| `createGzipStream`                   | `Compression.createGzipStream`                        |                                                                                                                                                                                                                                |
| `createInflateStream`                | `Compression.createInflateStream`                     |                                                                                                                                                                                                                                |
| `createReadStream`                   | —                                                     | `createReadStream` is deprecated: Use createReadStream from node:fs; this only forwards to it. Removed in 1.0.                                                                                                                 |
| `createTarArchive`                   | `Archive.createTar`                                   | `Archive.createTarArchive` is deprecated: Use `createTar`. Removed in 1.0.                                                                                                                                                     |
| `createTarReader`                    | `Archive.openTar`                                     | `Archive.createTarReader` is deprecated: Use `openTar`. Removed in 1.0.                                                                                                                                                        |
| `createTempDir`                      | `Fs.createTempDir`                                    |                                                                                                                                                                                                                                |
| `createTempDirSync`                  | `Fs.createTempDirSync`                                |                                                                                                                                                                                                                                |
| `createUnzlibStream`                 | `Compression.createUnzlibStream`                      |                                                                                                                                                                                                                                |
| `createWriteStream`                  | —                                                     | `createWriteStream` is deprecated: Use createWriteStream from node:fs; this only forwards to it. Removed in 1.0.                                                                                                               |
| `createZlibStream`                   | `Compression.createZlibStream`                        |                                                                                                                                                                                                                                |
| `decompress`                         | `Compression.inflate`                                 | `decompress` is deprecated: Use `Compression.inflate`. Removed in 1.0.                                                                                                                                                         |
| `decompressAuto`                     | `Compression.decompressAuto`                          |                                                                                                                                                                                                                                |
| `decompressAutoSync`                 | `Compression.decompressAutoSync`                      |                                                                                                                                                                                                                                |
| `decompressSync`                     | `Compression.inflateSync`                             | `decompressSync` is deprecated: Use `Compression.inflateSync`. Removed in 1.0.                                                                                                                                                 |
| `DecryptionError`                    | `ZipDecryptionError`                                  | `DecryptionError` is deprecated: Use `ZipDecryptionError`. Removed in 1.0.                                                                                                                                                     |
| `deflateBatchWithPool`               | `Compression.Workers.deflateBatch`                    | `deflateBatchWithPool` is deprecated: Use `Compression.Workers.deflateBatch`. Removed in 1.0.                                                                                                                                  |
| `deflateWithPool`                    | `Compression.Workers.deflate`                         | `deflateWithPool` is deprecated: Use `Compression.Workers.deflate`. Removed in 1.0.                                                                                                                                            |
| `detectCompressionFormat`            | `Compression.detectFormat`                            | `detectCompressionFormat` is deprecated: Use `Compression.detectFormat`. Removed in 1.0.                                                                                                                                       |
| `editZip`                            | `Archive.ZipEditor.open`                              | `Archive.editZip` is deprecated: Use `ZipEditor.open`; this only forwards to it. Removed in 1.0.                                                                                                                               |
| `editZipUrl`                         | `Archive.ZipEditor.openUrl`                           | `Archive.editZipUrl` is deprecated: Use `ZipEditor.openUrl`; this only forwards to it. Removed in 1.0.                                                                                                                         |
| `encryptionMethodFromAesKeyStrength` | `Archive.Crypto.getEncryptionMethodForAesKeyStrength` | `Archive.encryptionMethodFromAesKeyStrength` is deprecated: Use `Crypto.getEncryptionMethodForAesKeyStrength`. Removed in 1.0.                                                                                                 |
| `ensureDir`                          | `Fs.ensureDir`                                        |                                                                                                                                                                                                                                |
| `ensureDirSync`                      | `Fs.ensureDirSync`                                    |                                                                                                                                                                                                                                |
| `EntrySizeMismatchError`             | `ZipEntrySizeMismatchError`                           | `EntrySizeMismatchError` is deprecated: Use `ZipEntrySizeMismatchError`. Removed in 1.0.                                                                                                                                       |
| `EntrySizeMismatchReason`            | `ZipEntrySizeMismatchReason`                          | `EntrySizeMismatchReason` is deprecated: Use `ZipEntrySizeMismatchReason`. Removed in 1.0.                                                                                                                                     |
| `EocdNotFoundError`                  | `ZipEocdNotFoundError`                                | `EocdNotFoundError` is deprecated: Use `ZipEocdNotFoundError`. Removed in 1.0.                                                                                                                                                 |
| `ExtractOptions`                     | `RemoteZipExtractOptions`                             | `ExtractOptions` is deprecated: Use `RemoteZipExtractOptions`. Removed in 1.0.                                                                                                                                                 |
| `FileEntry`                          | `FsFileEntry`                                         | `FileEntry` is deprecated: Use `FsFileEntry`. Removed in 1.0.                                                                                                                                                                  |
| `fileExists`                         | `Fs.fileExists`                                       |                                                                                                                                                                                                                                |
| `fileExistsSync`                     | `Fs.fileExistsSync`                                   |                                                                                                                                                                                                                                |
| `FileTooLargeError`                  | `ArchiveFileTooLargeError`                            | `FileTooLargeError` is deprecated: Use `ArchiveFileTooLargeError`. Removed in 1.0.                                                                                                                                             |
| `getAesKeyStrength`                  | `Archive.Crypto.getAesKeyStrength`                    |                                                                                                                                                                                                                                |
| `getDefaultWorkerPool`               | `Compression.Workers.getDefaultPool`                  | `getDefaultWorkerPool` is deprecated: Use `Compression.Workers.getDefaultPool`. Removed in 1.0.                                                                                                                                |
| `getEncryptionMethodName`            | `Archive.Crypto.getEncryptionMethodName`              |                                                                                                                                                                                                                                |
| `glob`                               | `Fs.glob`                                             |                                                                                                                                                                                                                                |
| `GlobOptions`                        | `FsGlobOptions`                                       | `GlobOptions` is deprecated: Use `FsGlobOptions`. Removed in 1.0.                                                                                                                                                              |
| `globSync`                           | `Fs.globSync`                                         |                                                                                                                                                                                                                                |
| `globToRegex`                        | `Fs.globToRegex`                                      |                                                                                                                                                                                                                                |
| `gunzip`                             | `Compression.gunzip`                                  |                                                                                                                                                                                                                                |
| `gunzipSync`                         | `Compression.gunzipSync`                              |                                                                                                                                                                                                                                |
| `gzip`                               | `Compression.gzip`                                    |                                                                                                                                                                                                                                |
| `gzipSync`                           | `Compression.gzipSync`                                |                                                                                                                                                                                                                                |
| `hasCompressionStream`               | `Compression.hasCompressionStream`                    |                                                                                                                                                                                                                                |
| `hasDeflateRaw`                      | `Compression.hasDeflateRaw`                           |                                                                                                                                                                                                                                |
| `hasGzipCompressionStream`           | `Compression.hasGzipCompressionStream`                |                                                                                                                                                                                                                                |
| `hasGzipDecompressionStream`         | `Compression.hasGzipDecompressionStream`              |                                                                                                                                                                                                                                |
| `hasWorkerSupport`                   | `Compression.hasWorkerSupport`                        |                                                                                                                                                                                                                                |
| `HttpRangeError`                     | `ArchiveHttpRangeError`                               | `HttpRangeError` is deprecated: Use `ArchiveHttpRangeError`. Removed in 1.0.                                                                                                                                                   |
| `HttpRangeReader`                    | `Archive.HttpRangeReader`                             |                                                                                                                                                                                                                                |
| `inflateBatchWithPool`               | `Compression.Workers.inflateBatch`                    | `inflateBatchWithPool` is deprecated: Use `Compression.Workers.inflateBatch`. Removed in 1.0.                                                                                                                                  |
| `inflateWithPool`                    | `Compression.Workers.inflate`                         | `inflateWithPool` is deprecated: Use `Compression.Workers.inflate`. Removed in 1.0.                                                                                                                                            |
| `InvalidZipSignatureError`           | `ZipInvalidSignatureError`                            | `InvalidZipSignatureError` is deprecated: Use `ZipInvalidSignatureError`. Removed in 1.0.                                                                                                                                      |
| `isAesEncryption`                    | `Archive.Crypto.isAesEncryption`                      |                                                                                                                                                                                                                                |
| `isGzipData`                         | `Compression.isGzipData`                              |                                                                                                                                                                                                                                |
| `isTarDataEntry`                     | `Archive.isTarDataEntry`                              |                                                                                                                                                                                                                                |
| `isTarDirectory`                     | `Archive.isTarDirectory`                              |                                                                                                                                                                                                                                |
| `isTarFile`                          | `Archive.isTarFile`                                   |                                                                                                                                                                                                                                |
| `isTarHardLink`                      | `Archive.isTarHardLink`                               |                                                                                                                                                                                                                                |
| `isTarSymlink`                       | `Archive.isTarSymlink`                                |                                                                                                                                                                                                                                |
| `isZlibData`                         | `Compression.isZlibData`                              |                                                                                                                                                                                                                                |
| `matchGlob`                          | `Fs.matchGlob`                                        |                                                                                                                                                                                                                                |
| `matchGlobAny`                       | `Fs.matchGlobAny`                                     |                                                                                                                                                                                                                                |
| `OpenArchiveOptions`                 | `ArchiveFileOpenOptions`                              | `OpenArchiveOptions` is deprecated: Use `ArchiveFileOpenOptions`. Removed in 1.0.                                                                                                                                              |
| `OpenArchiveOptionsTar`              | `ArchiveFileOpenOptionsTar`                           | `OpenArchiveOptionsTar` is deprecated: Use `ArchiveFileOpenOptionsTar`. Removed in 1.0.                                                                                                                                        |
| `OpenArchiveOptionsZip`              | `ArchiveFileOpenOptionsZip`                           | `OpenArchiveOptionsZip` is deprecated: Use `ArchiveFileOpenOptionsZip`. Removed in 1.0.                                                                                                                                        |
| `OpenTarOptions`                     | `ArchiveFileOpenOptionsTar`                           | `OpenTarOptions` is deprecated: Use `ArchiveFileOpenOptionsTar`. Removed in 1.0.                                                                                                                                               |
| `OpenZipOptions`                     | `ArchiveFileOpenOptionsZip`                           | `OpenZipOptions` is deprecated: Use `ArchiveFileOpenOptionsZip`. Removed in 1.0.                                                                                                                                               |
| `parseTar`                           | `Archive.parseTar`                                    |                                                                                                                                                                                                                                |
| `parseTarGz`                         | `Archive.parseTarGz`                                  |                                                                                                                                                                                                                                |
| `parseTarGzStream`                   | `Archive.parseTarGzStream`                            |                                                                                                                                                                                                                                |
| `parseTarStream`                     | `Archive.parseTarStream`                              |                                                                                                                                                                                                                                |
| `PasswordRequiredError`              | `ZipPasswordRequiredError`                            | `PasswordRequiredError` is deprecated: Use `ZipPasswordRequiredError`. Removed in 1.0.                                                                                                                                         |
| `randomBytes`                        | —                                                     | `Archive.randomBytes` is deprecated: Not an archive operation. Use `Bytes.random` from `documonster/utils`. Removed in 1.0.                                                                                                    |
| `RangeNotSupportedError`             | `ArchiveRangeNotSupportedError`                       | `RangeNotSupportedError` is deprecated: Use `ArchiveRangeNotSupportedError`. Removed in 1.0.                                                                                                                                   |
| `readFileBytes`                      | `Fs.readFileBytes`                                    |                                                                                                                                                                                                                                |
| `readFileBytesSync`                  | `Fs.readFileBytesSync`                                |                                                                                                                                                                                                                                |
| `readFileText`                       | `Fs.readFileText`                                     |                                                                                                                                                                                                                                |
| `readFileTextSync`                   | `Fs.readFileTextSync`                                 |                                                                                                                                                                                                                                |
| `ReadStreamOptions`                  | —                                                     | `ReadStreamOptions` is deprecated: Use `ReadStreamOptions` from `node:fs`, as `createReadStream` from `node:fs` replaces this module's. Removed in 1.0.                                                                        |
| `RemoteZipReader`                    | `Archive.RemoteZipReader`                             |                                                                                                                                                                                                                                |
| `remove`                             | `Fs.remove`                                           |                                                                                                                                                                                                                                |
| `removeSync`                         | `Fs.removeSync`                                       |                                                                                                                                                                                                                                |
| `safeStats`                          | `Fs.safeStats`                                        |                                                                                                                                                                                                                                |
| `safeStatsSync`                      | `Fs.safeStatsSync`                                    |                                                                                                                                                                                                                                |
| `setFileTime`                        | `Fs.setFileTime`                                      |                                                                                                                                                                                                                                |
| `setFileTimeSync`                    | `Fs.setFileTimeSync`                                  |                                                                                                                                                                                                                                |
| `stringToUint8Array`                 | —                                                     | `Archive.stringToUint8Array` is deprecated: Not an archive operation. Use `Bytes.fromText` from `documonster/utils`. Removed in 1.0.                                                                                           |
| `tar`                                | `Archive.packTar`                                     | `Archive.tar` is deprecated: Use `packTar`. Removed in 1.0.                                                                                                                                                                    |
| `TAR_BLOCK_SIZE`                     | `Archive.TAR_BLOCK_SIZE`                              |                                                                                                                                                                                                                                |
| `TAR_TYPE`                           | `Archive.TarType`                                     | `Archive.TAR_TYPE` is deprecated: Use `TarType`. Removed in 1.0.                                                                                                                                                               |
| `TarArchive`                         | `Archive.TarArchive`                                  |                                                                                                                                                                                                                                |
| `TarArchive#bytes`                   | `Archive.TarArchive#toBuffer`                         | deprecated: Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                    |
| `TarArchive#bytesSync`               | `Archive.TarArchive#toBufferSync`                     | deprecated: Use `toBufferSync`. Removed in 1.0.                                                                                                                                                                                |
| `targz`                              | `Archive.packTarGz`                                   | `targz` is deprecated: Use `Archive.packTarGz`. Removed in 1.0.                                                                                                                                                                |
| `TarGzArchive`                       | `Archive.TarGzArchive`                                |                                                                                                                                                                                                                                |
| `TarGzArchive#bytes`                 | `Archive.TarGzArchive#toBuffer`                       | deprecated: Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                    |
| `TarGzArchive#bytesSync`             | `Archive.TarGzArchive#toBufferSync`                   | deprecated: Use `toBufferSync`. Removed in 1.0.                                                                                                                                                                                |
| `TarReader`                          | `Archive.TarReader`                                   |                                                                                                                                                                                                                                |
| `TarReaderEntry`                     | `Archive.TarReaderEntry`                              |                                                                                                                                                                                                                                |
| `tarSync`                            | `Archive.packTarSync`                                 | `Archive.tarSync` is deprecated: Use `packTarSync`. Removed in 1.0.                                                                                                                                                            |
| `TaskOptions`                        | `WorkerTaskOptions`                                   | `TaskOptions` is deprecated: Use `WorkerTaskOptions`. Removed in 1.0.                                                                                                                                                          |
| `TaskResult`                         | `WorkerTaskResult`                                    | `TaskResult` is deprecated: Use `WorkerTaskResult`. Removed in 1.0.                                                                                                                                                            |
| `terminateDefaultWorkerPool`         | `Compression.Workers.terminateDefaultPool`            | `terminateDefaultWorkerPool` is deprecated: Use `Compression.Workers.terminateDefaultPool`. Removed in 1.0.                                                                                                                    |
| `throwIfAborted`                     | `Archive.throwIfAborted`                              |                                                                                                                                                                                                                                |
| `toAsyncIterable`                    | `Archive.toAsyncIterable`                             |                                                                                                                                                                                                                                |
| `toNodeReadable`                     | `Archive.toNodeReadable`                              |                                                                                                                                                                                                                                |
| `toReadableStream`                   | `Archive.toReadableStream`                            |                                                                                                                                                                                                                                |
| `TransformEntryData`                 | `ArchiveTransformEntryData`                           | `TransformEntryData` is deprecated: Use `ArchiveTransformEntryData`. Removed in 1.0.                                                                                                                                           |
| `TransformFunction`                  | `ArchiveTransformFunction`                            | `TransformFunction` is deprecated: Use `ArchiveTransformFunction`. Removed in 1.0.                                                                                                                                             |
| `TransformResult`                    | `ArchiveTransformResult`                              | `TransformResult` is deprecated: Use `ArchiveTransformResult`. Removed in 1.0.                                                                                                                                                 |
| `traverseDirectory`                  | `Fs.traverseDirectory`                                |                                                                                                                                                                                                                                |
| `traverseDirectorySync`              | `Fs.traverseDirectorySync`                            |                                                                                                                                                                                                                                |
| `TraverseOptions`                    | `FsTraverseOptions`                                   | `TraverseOptions` is deprecated: Use `FsTraverseOptions`. Removed in 1.0.                                                                                                                                                      |
| `uint8ArrayToBase64`                 | —                                                     | `Archive.uint8ArrayToBase64` is deprecated: Not an archive operation. Use `Bytes.toBase64` from `documonster/utils`. Removed in 1.0.                                                                                           |
| `uint8ArrayToString`                 | —                                                     | `Archive.uint8ArrayToString` is deprecated: Not an archive operation. Use `Bytes.toText(bytes, { encoding })` from `documonster/utils`. Removed in 1.0.                                                                        |
| `UnsupportedCompressionError`        | `ZipUnsupportedCompressionError`                      | `UnsupportedCompressionError` is deprecated: Use `ZipUnsupportedCompressionError`. Removed in 1.0.                                                                                                                             |
| `untar`                              | `Archive.unpackTar`                                   | `Archive.untar` is deprecated: Use `unpackTar`. Removed in 1.0.                                                                                                                                                                |
| `untargz`                            | `Archive.unpackTarGz`                                 | `untargz` is deprecated: Use `Archive.unpackTarGz`. Removed in 1.0.                                                                                                                                                            |
| `unzip`                              | `Archive.openZip`                                     | `Archive.unzip` is deprecated: Use `openZip`, or `openTar` for `{ format: "tar" }`. `unzip` named a reader while `untar` named a map. Removed in 1.0.                                                                          |
| `UnzipEntry`                         | —                                                     | `Archive.UnzipEntry` is deprecated: Entries come from a reader — `ZipReader.entries()` / `get()`, or iterating `Archive.openZip(…)` — and the shapes this takes are the reader's own. Removed in 1.0.                          |
| `unzlib`                             | `Compression.unzlib`                                  |                                                                                                                                                                                                                                |
| `unzlibSync`                         | `Compression.unzlibSync`                              |                                                                                                                                                                                                                                |
| `WorkerPool`                         | `Compression.Workers.Pool`                            | `WorkerPool` is deprecated: Use `Compression.Workers.Pool`. Removed in 1.0.                                                                                                                                                    |
| `writeFileBytes`                     | `Fs.writeFileBytes`                                   |                                                                                                                                                                                                                                |
| `writeFileBytesSync`                 | `Fs.writeFileBytesSync`                               |                                                                                                                                                                                                                                |
| `writeFileText`                      | `Fs.writeFileText`                                    |                                                                                                                                                                                                                                |
| `writeFileTextSync`                  | `Fs.writeFileTextSync`                                |                                                                                                                                                                                                                                |
| `WriteStreamOptions`                 | —                                                     | `WriteStreamOptions` is deprecated: Use `WriteStreamOptions` from `node:fs`, as `createWriteStream` from `node:fs` replaces this module's. Removed in 1.0.                                                                     |
| `zip`                                | `Archive.createZip`                                   | `Archive.zip` is deprecated: Use `createZip`, or `createTar` for `{ format: "tar" }`. `zip` named a builder while `tar` named bytes. Removed in 1.0.                                                                           |
| `ZIP_CRYPTO_HEADER_SIZE`             | `Archive.Crypto.ZIP_CRYPTO_HEADER_SIZE`               |                                                                                                                                                                                                                                |
| `ZipArchive`                         | `Archive.ZipArchive`                                  |                                                                                                                                                                                                                                |
| `ZipArchive#bytes`                   | `Archive.ZipArchive#toBuffer`                         | deprecated: Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                    |
| `ZipArchive#bytesSync`               | `Archive.ZipArchive#toBufferSync`                     | deprecated: Use `toBufferSync`. Removed in 1.0.                                                                                                                                                                                |
| `zipCryptoDecrypt`                   | `Archive.Crypto.decryptZipCrypto`                     | `Archive.zipCryptoDecrypt` is deprecated: Use `Crypto.decryptZipCrypto`, which throws an `ArchiveError` where this returns `null`. Removed in 1.0.                                                                             |
| `zipCryptoEncrypt`                   | `Archive.Crypto.encryptZipCrypto`                     | `Archive.zipCryptoEncrypt` is deprecated: Use `Crypto.encryptZipCrypto`. Removed in 1.0.                                                                                                                                       |
| `zipCryptoInitKeys`                  | `Archive.Crypto.createZipCryptoKeys`                  | `Archive.zipCryptoInitKeys` is deprecated: Use `Crypto.createZipCryptoKeys`. Removed in 1.0.                                                                                                                                   |
| `ZipEditor`                          | `Archive.ZipEditor`                                   |                                                                                                                                                                                                                                |
| `ZipEditor#bytes`                    | `Archive.ZipEditor#toBuffer`                          | deprecated: Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                    |
| `ZipEditor#getEntries`               | `Archive.ZipEditor#getSourceEntries`                  | deprecated: Use `getSourceEntries`. The name read as the archive the edits produce — which `ArchiveFile#getEntries` returns — while this has always returned the source's entries. Removed in 1.0.                             |
| `ZipEditPlan`                        | `Archive.ZipEditPlan`                                 |                                                                                                                                                                                                                                |
| `ZipEditPlan#toJSON`                 | `Archive.ZipEditPlan#toJson`                          | deprecated: Use `toJson`. Being named `toJSON`, this is also what `JSON.stringify(plan)` calls — and since it returns a Promise, that call serialises `{}`. Removed in 1.0.                                                    |
| `ZipReader`                          | `Archive.ZipReader`                                   |                                                                                                                                                                                                                                |
| `zlib`                               | `Compression.zlib`                                    |                                                                                                                                                                                                                                |
| `zlibSync`                           | `Compression.zlibSync`                                |                                                                                                                                                                                                                                |

**`@cj-tech-master/excelts/browser`** → `documonster/excel [browser]` — every row of `@cj-tech-master/excelts` above applies; only the differences are listed

**`@cj-tech-master/excelts/chart`** → `documonster/chart`

| excelts                          | documonster                                               | Note                                                                                                                                                                                                                                                                                                                                                 |
| -------------------------------- | --------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `applyAxisTransform`             | `ChartClassic.Backend.applyAxisTransform`                 |                                                                                                                                                                                                                                                                                                                                                      |
| `applyChartExPreset`             | `ChartEx.applyPreset`                                     | `applyChartExPreset` is deprecated: Use `ChartEx.applyPreset`. Removed in 1.0.                                                                                                                                                                                                                                                                       |
| `applyChartPreset`               | `ChartClassic.applyPreset`                                | `applyChartPreset` is deprecated: Use `ChartClassic.applyPreset`. Removed in 1.0.                                                                                                                                                                                                                                                                    |
| `buildChartColors`               | `ChartClassic.toColorsXml`                                | `buildChartColors` is deprecated: Use `ChartClassic.toColorsXml`. Removed in 1.0.                                                                                                                                                                                                                                                                    |
| `buildChartExModel`              | `ChartEx.createModel`                                     | `buildChartExModel` is deprecated: Use `ChartEx.createModel`. Removed in 1.0.                                                                                                                                                                                                                                                                        |
| `buildChartModel`                | `ChartClassic.createModel`                                | `buildChartModel` is deprecated: Use `ChartClassic.createModel`. Removed in 1.0.                                                                                                                                                                                                                                                                     |
| `buildChartScene`                | `ChartClassic.Backend.toScene`                            | `buildChartScene` is deprecated: Use `ChartClassic.Backend.toScene`. Removed in 1.0.                                                                                                                                                                                                                                                                 |
| `buildChartStyle`                | `ChartClassic.toStyleXml`                                 | `buildChartStyle` is deprecated: Use `ChartClassic.toStyleXml`. Removed in 1.0.                                                                                                                                                                                                                                                                      |
| `buildComboChartModel`           | `ChartClassic.createComboModel`                           | `buildComboChartModel` is deprecated: Use `ChartClassic.createComboModel`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `buildEffectFilter`              | `ChartClassic.Backend.toSvgFilter`                        | `buildEffectFilter` is deprecated: Use `ChartClassic.Backend.toSvgFilter`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `buildSpPr`                      | `DrawingMl.createSpPr`                                    | `buildSpPr` is deprecated: Use `DrawingMl.createSpPr`. Removed in 1.0.                                                                                                                                                                                                                                                                               |
| `buildTxPr`                      | `DrawingMl.createTxPr`                                    | `buildTxPr` is deprecated: Use `DrawingMl.createTxPr`. Removed in 1.0.                                                                                                                                                                                                                                                                               |
| `canRenderChartExAsVectorPdf`    | —                                                         | `canRenderChartExAsVectorPdf` is deprecated: Every Office 2016 chart has a vector picture: `ChartEx.toDrawList(model)` draws any model, and `Pdf.fromChart` picks its path itself. This still says whether each series' layout is one of `VECTOR_PDF_CHART_EX_LAYOUT_IDS`. Removed in 1.0.                                                           |
| `chartExOptionsFromRows`         | `ChartEx.createOptionsFromRows`                           | `chartExOptionsFromRows` is deprecated: Use `ChartEx.createOptionsFromRows`. Removed in 1.0.                                                                                                                                                                                                                                                         |
| `chartExOptionsFromTable`        | `ChartEx.createOptionsFromTable`                          | `chartExOptionsFromTable` is deprecated: Use `ChartEx.createOptionsFromTable`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `chartOptionsFromRows`           | `ChartClassic.createOptionsFromRows`                      | `chartOptionsFromRows` is deprecated: Use `ChartClassic.createOptionsFromRows`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `chartOptionsFromTable`          | `ChartClassic.createOptionsFromTable`                     | `chartOptionsFromTable` is deprecated: Use `ChartClassic.createOptionsFromTable`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `ChartPdfDrawingSurface`         | —                                                         | `ChartPdfDrawingSurface` is deprecated: A chart goes onto a PDF page as a display list, through `Pdf.createDrawSurface` (a `PdfDrawPage`). Removed in 1.0.                                                                                                                                                                                           |
| `ChartPdfPathOp`                 | —                                                         | `ChartPdfPathOp` is deprecated: A path operation of `ChartPdfDrawingSurface`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `drawChartExPdf`                 | —                                                         | `drawChartExPdf` is deprecated: Put the chart on the page as a display list, as every producer does: `Draw.render(ChartEx.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`. This draws that same display list — the chart and its chart area — onto a surface with `drawRect` / `drawLine` / `drawText`. Removed in 1.0.    |
| `drawChartPdf`                   | —                                                         | `drawChartPdf` is deprecated: Put the chart on the page as a display list, as every producer does: `Draw.render(ChartClassic.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`. This draws that same display list — the chart and its chart area — onto a surface with `drawRect` / `drawLine` / `drawText`. Removed in 1.0. |
| `EXCEL_CHART_EX_PRESETS`         | —                                                         | `EXCEL_CHART_EX_PRESETS` is deprecated: The preset names are the keys of the table: `Object.keys(CHART_EX_PRESETS)`. Removed in 1.0.                                                                                                                                                                                                                 |
| `EXCEL_CHART_PRESETS`            | —                                                         | `EXCEL_CHART_PRESETS` is deprecated: The preset names are the keys of the table: `Object.keys(CHART_PRESETS)`. Removed in 1.0.                                                                                                                                                                                                                       |
| `fillChartCaches`                | `ChartClassic.fillCaches(model, workbook, { worksheet })` | `fillChartCaches` is deprecated: Use `ChartClassic.fillCaches(model, workbook, { worksheet })`. Removed in 1.0.                                                                                                                                                                                                                                      |
| `fillChartExCaches`              | `ChartEx.fillCaches(model, workbook, { worksheet })`      | `fillChartExCaches` is deprecated: Use `ChartEx.fillCaches(model, workbook, { worksheet })`. Removed in 1.0.                                                                                                                                                                                                                                         |
| `fillNumRef`                     | `ChartClassic.fillNumRef`                                 |                                                                                                                                                                                                                                                                                                                                                      |
| `fillStrRef`                     | `ChartClassic.fillStrRef`                                 |                                                                                                                                                                                                                                                                                                                                                      |
| `getSpPrFillColor`               | `DrawingMl.getSpPrFillColor`                              |                                                                                                                                                                                                                                                                                                                                                      |
| `getSpPrGradient`                | `DrawingMl.getSpPrGradient`                               |                                                                                                                                                                                                                                                                                                                                                      |
| `getSpPrLine`                    | `DrawingMl.getSpPrLine`                                   |                                                                                                                                                                                                                                                                                                                                                      |
| `getSpPrPattern`                 | `DrawingMl.getSpPrPattern`                                |                                                                                                                                                                                                                                                                                                                                                      |
| `getTxPrColor`                   | `DrawingMl.getTxPrColor`                                  |                                                                                                                                                                                                                                                                                                                                                      |
| `getTxPrFontSize`                | `DrawingMl.getTxPrFontSize`                               |                                                                                                                                                                                                                                                                                                                                                      |
| `parseChartColors`               | `ChartClassic.fromColorsXml`                              | `parseChartColors` is deprecated: Use `ChartClassic.fromColorsXml`. Removed in 1.0.                                                                                                                                                                                                                                                                  |
| `parseChartEx`                   | `ChartEx.fromXml`                                         | `parseChartEx` is deprecated: Use `ChartEx.fromXml`. Removed in 1.0.                                                                                                                                                                                                                                                                                 |
| `parseChartStyle`                | `ChartClassic.fromStyleXml`                               | `parseChartStyle` is deprecated: Use `ChartClassic.fromStyleXml`. Removed in 1.0.                                                                                                                                                                                                                                                                    |
| `parseSpPr`                      | `DrawingMl.parseSpPr`                                     |                                                                                                                                                                                                                                                                                                                                                      |
| `parseTxPr`                      | `DrawingMl.parseTxPr`                                     |                                                                                                                                                                                                                                                                                                                                                      |
| `PdfChartRenderOptions`          | —                                                         | `PdfChartRenderOptions` is deprecated: Pass `ChartRenderOptions` to `ChartClassic.toDrawList` and the box to `Pdf.createDrawSurface`. Removed in 1.0.                                                                                                                                                                                                |
| `PdfColor`                       | —                                                         | `PdfColor` is deprecated: See `ChartPdfColor`. Removed in 1.0.                                                                                                                                                                                                                                                                                       |
| `renderChartEx`                  | `ChartEx.toXml`                                           | `renderChartEx` is deprecated: Use `ChartEx.toXml` — this returns the part's XML, not a picture. Removed in 1.0.                                                                                                                                                                                                                                     |
| `renderChartExPng`               | `ChartEx.toPng`                                           | `renderChartExPng` is deprecated: Use `ChartEx.toPng`. Removed in 1.0.                                                                                                                                                                                                                                                                               |
| `renderChartExSvg`               | `ChartEx.toSvg`                                           | `renderChartExSvg` is deprecated: Use `ChartEx.toSvg`. Removed in 1.0.                                                                                                                                                                                                                                                                               |
| `renderChartPng`                 | `ChartClassic.toPng`                                      | `renderChartPng` is deprecated: Use `ChartClassic.toPng`. Removed in 1.0.                                                                                                                                                                                                                                                                            |
| `renderChartSvg`                 | `ChartClassic.toSvg`                                      | `renderChartSvg` is deprecated: Use `ChartClassic.toSvg`. Removed in 1.0.                                                                                                                                                                                                                                                                            |
| `resolveTopologyObject`          | `ChartClassic.Backend.resolveTopologyObject`              |                                                                                                                                                                                                                                                                                                                                                      |
| `seriesFromColumns`              | `ChartClassic.createSeriesFromColumns`                    | `seriesFromColumns` is deprecated: Use `ChartClassic.createSeriesFromColumns`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `setSpPrFill`                    | `DrawingMl.withSpPrFill`                                  | `setSpPrFill` is deprecated: Use `DrawingMl.withSpPrFill`, which — as this always did — returns a copy rather than editing `spPr`. Removed in 1.0.                                                                                                                                                                                                   |
| `setSpPrLine`                    | `DrawingMl.withSpPrLine`                                  | `setSpPrLine` is deprecated: Use `DrawingMl.withSpPrLine`, which — as this always did — returns a copy rather than editing `spPr`. Removed in 1.0.                                                                                                                                                                                                   |
| `VECTOR_PDF_CHART_EX_LAYOUT_IDS` | —                                                         | `VECTOR_PDF_CHART_EX_LAYOUT_IDS` is deprecated: It only answers the deprecated `canRenderChartExAsVectorPdf`: every Office 2016 chart now has a vector picture through `ChartEx.toDrawList`. Removed in 1.0.                                                                                                                                         |

**`@cj-tech-master/excelts/csv`** → `documonster/csv`

| excelts                         | documonster                         | Note                                                                                       |
| ------------------------------- | ----------------------------------- | ------------------------------------------------------------------------------------------ |
| `applyDynamicTyping`            | `Csv.applyDynamicTyping`            |                                                                                            |
| `applyDynamicTypingToRow`       | `Csv.applyDynamicTypingToRow`       |                                                                                            |
| `BuiltinColumnType`             | `CsvBuiltinColumnType`              | `BuiltinColumnType` is deprecated: Use `CsvBuiltinColumnType`. Removed in 1.0.             |
| `CastDateConfig`                | `CsvCastDateConfig`                 | `CastDateConfig` is deprecated: Use `CsvCastDateConfig`. Removed in 1.0.                   |
| `ChunkMeta`                     | `CsvChunkMeta`                      | `ChunkMeta` is deprecated: Use `CsvChunkMeta`. Removed in 1.0.                             |
| `ColumnConfig`                  | `CsvColumnConfig`                   | `ColumnConfig` is deprecated: Use `CsvColumnConfig`. Removed in 1.0.                       |
| `ColumnDef`                     | `CsvColumnDef`                      | `ColumnDef` is deprecated: Use `CsvColumnDef`. Removed in 1.0.                             |
| `ColumnMismatchConfig`          | `CsvColumnMismatchConfig`           | `ColumnMismatchConfig` is deprecated: Use `CsvColumnMismatchConfig`. Removed in 1.0.       |
| `ColumnMismatchLess`            | `CsvColumnMismatchLess`             | `ColumnMismatchLess` is deprecated: Use `CsvColumnMismatchLess`. Removed in 1.0.           |
| `ColumnMismatchMore`            | `CsvColumnMismatchMore`             | `ColumnMismatchMore` is deprecated: Use `CsvColumnMismatchMore`. Removed in 1.0.           |
| `DecimalSeparator`              | `CsvDecimalSeparator`               | `DecimalSeparator` is deprecated: Use `CsvDecimalSeparator`. Removed in 1.0.               |
| `deduplicateHeaders`            | `Csv.deduplicateHeaders`            |                                                                                            |
| `deduplicateHeadersWithRenames` | `Csv.deduplicateHeadersWithRenames` |                                                                                            |
| `detectDelimiter`               | `Csv.detectDelimiter`               |                                                                                            |
| `detectLinebreak`               | `Csv.detectLinebreak`               |                                                                                            |
| `DynamicTypingConfig`           | `CsvDynamicTypingConfig`            | `DynamicTypingConfig` is deprecated: Use `CsvDynamicTypingConfig`. Removed in 1.0.         |
| `FormattedValue`                | `CsvFormattedValue`                 | `FormattedValue` is deprecated: Use `CsvFormattedValue`. Removed in 1.0.                   |
| `GeneratorColumnConfig`         | `CsvGeneratorColumnConfig`          | `GeneratorColumnConfig` is deprecated: Use `CsvGeneratorColumnConfig`. Removed in 1.0.     |
| `GeneratorContext`              | `CsvGeneratorContext`               | `GeneratorContext` is deprecated: Use `CsvGeneratorContext`. Removed in 1.0.               |
| `GeneratorFn`                   | `CsvGeneratorFn`                    | `GeneratorFn` is deprecated: Use `CsvGeneratorFn`. Removed in 1.0.                         |
| `HeaderArray`                   | `CsvHeaderArray`                    | `HeaderArray` is deprecated: Use `CsvHeaderArray`. Removed in 1.0.                         |
| `HeaderTransformFunction`       | `CsvHeaderTransformFunction`        | `HeaderTransformFunction` is deprecated: Use `CsvHeaderTransformFunction`. Removed in 1.0. |
| `isFormattedValue`              | `Csv.isFormattedValue`              |                                                                                            |
| `isRowHashArray`                | `Csv.isRowHashArray`                |                                                                                            |
| `OnSkipCallback`                | `CsvOnSkipCallback`                 | `OnSkipCallback` is deprecated: Use `CsvOnSkipCallback`. Removed in 1.0.                   |
| `processColumns`                | `Csv.processColumns`                |                                                                                            |
| `quoted`                        | `CsvRecordInfo.quoted`              |                                                                                            |
| `RecordInfo`                    | `CsvRecordInfo`                     | `RecordInfo` is deprecated: Use `CsvRecordInfo`. Removed in 1.0.                           |
| `RecordWithInfo`                | `CsvRecordWithInfo`                 | `RecordWithInfo` is deprecated: Use `CsvRecordWithInfo`. Removed in 1.0.                   |
| `Row`                           | `CsvRow`                            | `Row` is deprecated: Use `CsvRow`. Removed in 1.0.                                         |
| `RowArray`                      | `CsvRowArray`                       | `RowArray` is deprecated: Use `CsvRowArray`. Removed in 1.0.                               |
| `RowHashArray`                  | `CsvRowHashArray`                   | `RowHashArray` is deprecated: Use `CsvRowHashArray`. Removed in 1.0.                       |
| `rowHashArrayMapByHeaders`      | `Csv.getRowValuesByHeaders`         | `Csv.rowHashArrayMapByHeaders` is deprecated: Use `getRowValuesByHeaders`. Removed in 1.0. |
| `rowHashArrayToHeaders`         | `Csv.getRowHeaders`                 | `Csv.rowHashArrayToHeaders` is deprecated: Use `getRowHeaders`. Removed in 1.0.            |
| `rowHashArrayToValues`          | `Csv.getRowValues`                  | `Csv.rowHashArrayToValues` is deprecated: Use `getRowValues`. Removed in 1.0.              |
| `RowMap`                        | `CsvRowMap`                         | `RowMap` is deprecated: Use `CsvRowMap`. Removed in 1.0.                                   |
| `RowTransformCallback`          | `CsvRowTransformCallback`           | `RowTransformCallback` is deprecated: Use `CsvRowTransformCallback`. Removed in 1.0.       |
| `RowTransformFunction`          | `CsvRowTransformFunction`           | `RowTransformFunction` is deprecated: Use `CsvRowTransformFunction`. Removed in 1.0.       |
| `RowValidateCallback`           | `CsvRowValidateCallback`            | `RowValidateCallback` is deprecated: Use `CsvRowValidateCallback`. Removed in 1.0.         |
| `RowValidateFunction`           | `CsvRowValidateFunction`            | `RowValidateFunction` is deprecated: Use `CsvRowValidateFunction`. Removed in 1.0.         |
| `StopCondition`                 | `CsvStopCondition`                  | `StopCondition` is deprecated: Use `CsvStopCondition`. Removed in 1.0.                     |
| `StopContext`                   | `CsvStopContext`                    | `StopContext` is deprecated: Use `CsvStopContext`. Removed in 1.0.                         |
| `stripBom`                      | `Csv.stripBom`                      |                                                                                            |
| `TransformContext`              | `CsvTransformContext`               | `TransformContext` is deprecated: Use `CsvTransformContext`. Removed in 1.0.               |
| `TransformResult`               | `CsvTransformResult`                | `TransformResult` is deprecated: Use `CsvTransformResult`. Removed in 1.0.                 |
| `TypeTransformMap`              | `CsvTypeTransformMap`               | `TypeTransformMap` is deprecated: Use `CsvTypeTransformMap`. Removed in 1.0.               |
| `unquoted`                      | `Csv.preventQuote`                  | `Csv.unquoted` is deprecated: Use `preventQuote`. Removed in 1.0.                          |

**`@cj-tech-master/excelts/formula`** → `documonster/formula`

| excelts    | documonster        | Note |
| ---------- | ------------------ | ---- |
| `parse`    | `Formula.parse`    |      |
| `tokenize` | `Formula.tokenize` |      |

**`@cj-tech-master/excelts/markdown`** → `documonster/markdown`

| excelts           | documonster | Note                                                                                                                                                                                                                                                          |
| ----------------- | ----------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `MarkdownOptions` | —           | `MarkdownOptions` is deprecated: Worksheet options do not belong to the standalone Markdown module. Use `ExcelMarkdownReadOptions` / `ExcelMarkdownWriteOptions` from `documonster/excel/markdown`, which the `ExcelMarkdown` functions take. Removed in 1.0. |

**`@cj-tech-master/excelts/pdf`** → `documonster/pdf`

| excelts                         | documonster                                  | Note                                                                                                                                                                                             |
| ------------------------------- | -------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `AnnotationOptions`             | `PdfAnnotationOptions`                       | `AnnotationOptions` is deprecated: Use `PdfAnnotationOptions`. Removed in 1.0.                                                                                                                   |
| `AnnotationType`                | `PdfAnnotationType`                          | `AnnotationType` is deprecated: Use `PdfAnnotationType`. Removed in 1.0.                                                                                                                         |
| `Asn1Node`                      | `PdfAsn1Node`                                | `Asn1Node` is deprecated: Use `PdfAsn1Node`. Removed in 1.0.                                                                                                                                     |
| `asn1Parse`                     | `Pdf.Signature.parseAsn1`                    | `Pdf.asn1Parse` is deprecated: Use `Pdf.Signature.parseAsn1`. Removed in 1.0.                                                                                                                    |
| `buildSignatureDictPlaceholder` | `Pdf.Signature.buildDictPlaceholder`         | `Pdf.buildSignatureDictPlaceholder` is deprecated: Use `Pdf.Signature.buildDictPlaceholder`. Removed in 1.0.                                                                                     |
| `ChartToPdfOptions`             | `PdfFromChartOptions`                        | `ChartToPdfOptions` is deprecated: Use `PdfFromChartOptions`. Removed in 1.0.                                                                                                                    |
| `CheckboxOptions`               | `PdfCheckboxOptions`                         | `CheckboxOptions` is deprecated: Use `PdfCheckboxOptions`. Removed in 1.0.                                                                                                                       |
| `CmsSignedData`                 | —                                            | `CmsSignedData` is deprecated: No public function takes or returns this. Removed in 1.0.                                                                                                         |
| `DocumentMetadata`              | `PdfDocumentMetadata`                        | `DocumentMetadata` is deprecated: Use `PdfDocumentMetadata`. Removed in 1.0.                                                                                                                     |
| `DocxToPdfOptions`              | `PdfFromDocxOptions`                         | `DocxToPdfOptions` is deprecated: Use `PdfFromDocxOptions`. Removed in 1.0.                                                                                                                      |
| `DrawCircleOptions`             | `PdfDrawCircleOptions`                       | `DrawCircleOptions` is deprecated: Use `PdfDrawCircleOptions`. Removed in 1.0.                                                                                                                   |
| `DrawEllipseOptions`            | `PdfDrawEllipseOptions`                      | `DrawEllipseOptions` is deprecated: Use `PdfDrawEllipseOptions`. Removed in 1.0.                                                                                                                 |
| `DrawImageOptions`              | `PdfDrawImageOptions`                        | `DrawImageOptions` is deprecated: Use `PdfDrawImageOptions`. Removed in 1.0.                                                                                                                     |
| `DrawLineOptions`               | `PdfDrawLineOptions`                         | `DrawLineOptions` is deprecated: Use `PdfDrawLineOptions`. Removed in 1.0.                                                                                                                       |
| `DrawPathOptions`               | `PdfDrawPathOptions`                         | `DrawPathOptions` is deprecated: Use `PdfDrawPathOptions`. Removed in 1.0.                                                                                                                       |
| `DrawRectOptions`               | `PdfDrawRectOptions`                         | `DrawRectOptions` is deprecated: Use `PdfDrawRectOptions`. Removed in 1.0.                                                                                                                       |
| `DrawSvgOptions`                | `PdfDrawSvgOptions`                          | `DrawSvgOptions` is deprecated: Use `PdfDrawSvgOptions`. Removed in 1.0.                                                                                                                         |
| `DrawTextOptions`               | `PdfDrawTextOptions`                         | `DrawTextOptions` is deprecated: Use `PdfDrawTextOptions`. Removed in 1.0.                                                                                                                       |
| `DropdownOptions`               | `PdfDropdownOptions`                         | `DropdownOptions` is deprecated: Use `PdfDropdownOptions`. Removed in 1.0.                                                                                                                       |
| `ExtractedImage`                | `PdfExtractedImage`                          | `ExtractedImage` is deprecated: Use `PdfExtractedImage`. Removed in 1.0.                                                                                                                         |
| `FormFieldOptions`              | `PdfFormFieldOptions`                        | `FormFieldOptions` is deprecated: Use `PdfFormFieldOptions`. Removed in 1.0.                                                                                                                     |
| `FreeTextAnnotationOptions`     | `PdfFreeTextAnnotationOptions`               | `FreeTextAnnotationOptions` is deprecated: Use `PdfFreeTextAnnotationOptions`. Removed in 1.0.                                                                                                   |
| `LoadOptions`                   | `PdfEditorReadOptions`                       | `LoadOptions` is deprecated: Use `PdfEditorReadOptions`. Removed in 1.0.                                                                                                                         |
| `PageOptions`                   | `PdfPageOptions`                             | `PageOptions` is deprecated: Use `PdfPageOptions`. Removed in 1.0.                                                                                                                               |
| `PageSizeName`                  | `PdfPageSizeName`                            | `PageSizeName` is deprecated: Use `PdfPageSizeName`. Removed in 1.0.                                                                                                                             |
| `PageSizes`                     | `Pdf.PageSizes`                              |                                                                                                                                                                                                  |
| `parseSvgPath`                  | `Pdf.parseSvgPath`                           |                                                                                                                                                                                                  |
| `PathOp`                        | `PdfPathOp`                                  | `PathOp` is deprecated: Use `PdfPathOp`. Removed in 1.0.                                                                                                                                         |
| `PdfDocumentBuilder#build`      | `Pdf.Builder#toBuffer`                       | deprecated: Use `PdfDocumentBuilder.toBuffer`. Removed in 1.0.                                                                                                                                   |
| `PdfEditor#save`                | `Pdf.Editor#toBuffer`                        | deprecated: Use `PdfEditor.toBuffer` (which keeps the encryption by default; pass `{ encryption: "remove" }` for this method's decrypted copy). Removed in 1.0.                                  |
| `PdfEditor#saveIncremental`     | `Pdf.Editor#toBuffer({ incremental: true })` | deprecated: Use `toBuffer({ incremental: true })` (whose fallback keeps the encryption by default). Removed in 1.0.                                                                              |
| `PdfRect`                       | `PdfAnnotationRect`                          | `PdfRect` is deprecated: Use `PdfAnnotationRect`. A rectangle by its corners, as an annotation's `/Rect` gives it; the general name read as the `{ x, y, width, height }` shape. Removed in 1.0. |
| `RadioGroupOptions`             | `PdfRadioGroupOptions`                       | `RadioGroupOptions` is deprecated: Use `PdfRadioGroupOptions`. Removed in 1.0.                                                                                                                   |
| `ReadPdfOptions`                | `PdfExtractOptions`                          | `ReadPdfOptions` is deprecated: Use `PdfExtractOptions`. Removed in 1.0.                                                                                                                         |
| `ReadPdfPage`                   | `PdfExtractedPage`                           | `ReadPdfPage` is deprecated: Use `PdfExtractedPage`. Removed in 1.0.                                                                                                                             |
| `ReadPdfResult`                 | `PdfExtractResult`                           | `ReadPdfResult` is deprecated: Use `PdfExtractResult`. Removed in 1.0.                                                                                                                           |
| `SignatureVerificationResult`   | —                                            | `SignatureVerificationResult` is deprecated: What the deprecated `Pdf.verifySignature` returns; `Pdf.Signature.verify` returns `PdfSignatureCheck`. Removed in 1.0.                              |
| `SignOptions`                   | —                                            | `SignOptions` is deprecated: No public function takes this; signing options are `PdfSignOptions`(`Pdf.Signature.fillPlaceholder`) and `PdfSignatureOptions` (`sign`). Removed in 1.0.            |
| `StampAnnotationOptions`        | `PdfStampAnnotationOptions`                  | `StampAnnotationOptions` is deprecated: Use `PdfStampAnnotationOptions`. Removed in 1.0.                                                                                                         |
| `TextAnnotationOptions`         | `PdfTextAnnotationOptions`                   | `TextAnnotationOptions` is deprecated: Use `PdfTextAnnotationOptions`. Removed in 1.0.                                                                                                           |
| `TextFieldOptions`              | `PdfTextFieldOptions`                        | `TextFieldOptions` is deprecated: Use `PdfTextFieldOptions`. Removed in 1.0.                                                                                                                     |
| `TextLine`                      | `PdfTextLine`                                | `TextLine` is deprecated: Use `PdfTextLine`. Removed in 1.0.                                                                                                                                     |
| `TextMarkupAnnotationOptions`   | `PdfTextMarkupAnnotationOptions`             | `TextMarkupAnnotationOptions` is deprecated: Use `PdfTextMarkupAnnotationOptions`. Removed in 1.0.                                                                                               |
| `TocOptions`                    | `PdfTocOptions`                              | `TocOptions` is deprecated: Use `PdfTocOptions`. Removed in 1.0.                                                                                                                                 |

**`@cj-tech-master/excelts/stream`** → `documonster/stream`

| excelts                                   | documonster                                       | Note                                                                                                                                                               |
| ----------------------------------------- | ------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `BufferedStream`                          | `Streams.BufferedStream`                          |                                                                                                                                                                    |
| `BufferedStream#toUint8Array`             | `Streams.BufferedStream#readAll`                  | deprecated: Use `readAll`, which does the same: despite the `to` name this consumes the buffered data. Removed in 1.0.                                             |
| `bytes`                                   | `Streams.bytes`                                   |                                                                                                                                                                    |
| `ChunkedBuilder`                          | `Streams.ChunkedBuilder`                          |                                                                                                                                                                    |
| `collect`                                 | `Streams.collect`                                 |                                                                                                                                                                    |
| `Collector`                               | `Streams.Collector`                               |                                                                                                                                                                    |
| `copyStream`                              | `pipeline`                                        | `copyStream` is deprecated: Use `pipeline`, which this duplicated. Removed in 1.0.                                                                                 |
| `createBufferedStream`                    | `StreamFactories.createBufferedStream`            |                                                                                                                                                                    |
| `createCollector`                         | `new Streams.Collector(options)`                  | `createCollector` is deprecated: Use `new Streams.Collector(options)`. Removed in 1.0.                                                                             |
| `createDuplex`                            | `Streams.createDuplex`                            |                                                                                                                                                                    |
| `createEmptyReadable`                     | `Streams.createEmptyReadable`                     |                                                                                                                                                                    |
| `createNullWritable`                      | `Streams.createNullWritable`                      |                                                                                                                                                                    |
| `createPassThrough`                       | `Streams.createPassThrough`                       |                                                                                                                                                                    |
| `createPullStream`                        | `StreamFactories.createPullStream`                |                                                                                                                                                                    |
| `createReadable`                          | `Streams.createReadable`                          |                                                                                                                                                                    |
| `createReadableFromArray`                 | `Streams.createReadableFromArray`                 |                                                                                                                                                                    |
| `createReadableFromAsyncIterable`         | `Streams.createReadableFromAsyncIterable`         |                                                                                                                                                                    |
| `createReadableFromGenerator`             | `Streams.createReadableFromGenerator`             |                                                                                                                                                                    |
| `createReadableFromPromise`               | `Streams.createReadableFromPromise`               |                                                                                                                                                                    |
| `createTransform`                         | `Streams.createTransform`                         |                                                                                                                                                                    |
| `createWritable`                          | `Streams.createWritable`                          |                                                                                                                                                                    |
| `DataChunk`                               | `StreamDataChunk`                                 | `DataChunk` is deprecated: Use `StreamDataChunk`. Removed in 1.0.                                                                                                  |
| `drainStream`                             | `Streams.discard`                                 | `drainStream` is deprecated: Use `Streams.discard`. Removed in 1.0.                                                                                                |
| `eventedReadableToAsyncIterableNoDestroy` | `Streams.LowLevel.eventsToAsyncIterableNoDestroy` | `eventedReadableToAsyncIterableNoDestroy` is deprecated: Use `Streams.LowLevel.eventsToAsyncIterableNoDestroy`. Removed in 1.0.                                    |
| `EventListener`                           | `StreamEventListener`                             | `EventListener` is deprecated: Use `StreamEventListener`; the bare name shadows the DOM's global `EventListener`. Removed in 1.0.                                  |
| `filter`                                  | `Streams.filter`                                  |                                                                                                                                                                    |
| `finishedAll`                             | `Streams.waitForFinished`                         | `finishedAll` is deprecated: Use `Streams.waitForFinished`. Removed in 1.0.                                                                                        |
| `fromBytes`                               | `Streams.fromBytes`                               |                                                                                                                                                                    |
| `fromJSON`                                | `StreamFactories.fromJSON`                        |                                                                                                                                                                    |
| `fromString`                              | `Streams.fromString`                              |                                                                                                                                                                    |
| `isAsyncIterable`                         | `Streams.isAsyncIterable`                         |                                                                                                                                                                    |
| `isDuplex`                                | `Streams.isDuplex`                                |                                                                                                                                                                    |
| `isReadableStream`                        | `Streams.isReadableStream`                        |                                                                                                                                                                    |
| `isReadableStreamLike`                    | `Streams.isReadableStream`                        | `isReadableStreamLike` is deprecated: Use `Streams.isReadableStream`, which this duplicated. Removed in 1.0.                                                       |
| `isStream`                                | `Streams.isStream`                                |                                                                                                                                                                    |
| `isTransform`                             | `Streams.isTransform`                             |                                                                                                                                                                    |
| `isTransformStream`                       | `Streams.isTransformStream`                       |                                                                                                                                                                    |
| `isWritableStream`                        | `Streams.isWritableStream`                        |                                                                                                                                                                    |
| `json`                                    | `Streams.json`                                    |                                                                                                                                                                    |
| `onceEvent`                               | `Streams.LowLevel.onceEvent`                      |                                                                                                                                                                    |
| `promisify`                               | `Streams.LowLevel.fromCallback`                   | `promisify` is deprecated: Use `Streams.LowLevel.fromCallback`: `promisify` read as `node:util`'s, which wraps a function rather than running one. Removed in 1.0. |
| `PullStream`                              | `Streams.PullStream`                              |                                                                                                                                                                    |
| `readableStreamToAsyncIterable`           | `Streams.readableStreamToAsyncIterable`           |                                                                                                                                                                    |
| `SinkAdapter`                             | `Streams.LowLevel.SinkAdapter`                    |                                                                                                                                                                    |
| `streamToBuffer`                          | `Streams.bytes`                                   | `streamToBuffer` is deprecated: Use `Streams.bytes`, which this duplicated: both resolve to a `Uint8Array`, not a `Buffer`. Removed in 1.0.                        |
| `streamToPromise`                         | `finished`                                        | `streamToPromise` is deprecated: Use `finished`, which this duplicated. Removed in 1.0.                                                                            |
| `streamToString`                          | `Streams.text`                                    | `streamToString` is deprecated: Use `Streams.text`; pass the encoding as `{ encoding }`. Removed in 1.0.                                                           |
| `streamToUint8Array`                      | `Streams.bytes`                                   | `streamToUint8Array` is deprecated: Use `Streams.bytes`. Removed in 1.0.                                                                                           |
| `StreamTypeError`                         | `StreamArgumentError`                             | `StreamTypeError` is deprecated: Use `StreamArgumentError`. Removed in 1.0.                                                                                        |
| `text`                                    | `Streams.text`                                    |                                                                                                                                                                    |
| `toWritable`                              | `Streams.toWritable`                              |                                                                                                                                                                    |
| `TransactionalChunkedBuilder`             | `Streams.TransactionalChunkedBuilder`             |                                                                                                                                                                    |
| `transform`                               | `StreamFactories.transform`                       |                                                                                                                                                                    |
| `UnsupportedStreamTypeError`              | `StreamUnsupportedTypeError`                      | `UnsupportedStreamTypeError` is deprecated: Use `StreamUnsupportedTypeError`. Removed in 1.0.                                                                      |

**`@cj-tech-master/excelts/word`** → `documonster/word`

| excelts                          | documonster                          | Note                                                                                                                                                                                                                                                                                                                         |
| -------------------------------- | ------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `acceptAllRevisions`             | `Document.acceptAllRevisions`        |                                                                                                                                                                                                                                                                                                                              |
| `acceptRevision`                 | `Document.acceptRevision`            |                                                                                                                                                                                                                                                                                                                              |
| `AssetRef`                       | —                                    | `AssetRef` is deprecated: Reserved for a package-model API that was never added: no function takes or returns it. Removed in 1.0.                                                                                                                                                                                            |
| `authorField`                    | `Build.authorField`                  |                                                                                                                                                                                                                                                                                                                              |
| `BiDiDirection`                  | `DocxBiDiDirection`                  | `BiDiDirection` is deprecated: Use `DocxBiDiDirection`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `bindChartData`                  | `Template.bindChartData`             |                                                                                                                                                                                                                                                                                                                              |
| `bold`                           | `Build.bold`                         |                                                                                                                                                                                                                                                                                                                              |
| `bookmarkEnd`                    | `Build.bookmarkEnd`                  |                                                                                                                                                                                                                                                                                                                              |
| `bookmarkStart`                  | `Build.bookmarkStart`                |                                                                                                                                                                                                                                                                                                                              |
| `border`                         | `Build.border`                       |                                                                                                                                                                                                                                                                                                                              |
| `BuildingBlock`                  | `DocxBuildingBlock`                  | `BuildingBlock` is deprecated: Use `DocxBuildingBlock`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `BuildingBlockGallery`           | `DocxBuildingBlockGallery`           | `BuildingBlockGallery` is deprecated: Use `DocxBuildingBlockGallery`. Removed in 1.0.                                                                                                                                                                                                                                        |
| `carriageReturn`                 | `Build.carriageReturn`               |                                                                                                                                                                                                                                                                                                                              |
| `cell`                           | `Build.cell`                         |                                                                                                                                                                                                                                                                                                                              |
| `chart`                          | `ChartContent.chart`                 |                                                                                                                                                                                                                                                                                                                              |
| `ChartBinding`                   | `DocxChartBinding`                   | `ChartBinding` is deprecated: Use `DocxChartBinding`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `ChartSeriesData`                | `DocxChartSeriesData`                | `ChartSeriesData` is deprecated: Use `DocxChartSeriesData`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `ChartTemplateData`              | `DocxChartTemplateData`              | `ChartTemplateData` is deprecated: Use `DocxChartTemplateData`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `checkBox`                       | `Build.checkbox`                     | `Build.checkBox` is deprecated: Use `checkbox`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `cmToEmu`                        | `Units.cmToEmu`                      |                                                                                                                                                                                                                                                                                                                              |
| `cmToTwips`                      | `Units.cmToTwips`                    |                                                                                                                                                                                                                                                                                                                              |
| `collectParagraphs`              | `Query.listParagraphs`               | `Query.collectParagraphs` is deprecated: Use `listParagraphs`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `collectRuns`                    | `Query.listRuns`                     | `Query.collectRuns` is deprecated: Use `listRuns`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `collectTables`                  | `Query.listTables`                   | `Query.collectTables` is deprecated: Use `listTables`, passing the parts to include — `listTables` searches only the body unless asked. Removed in 1.0.                                                                                                                                                                      |
| `columnBreak`                    | `Build.columnBreak`                  |                                                                                                                                                                                                                                                                                                                              |
| `commentRangeEnd`                | `Build.commentRangeEnd`              |                                                                                                                                                                                                                                                                                                                              |
| `commentRangeStart`              | `Build.commentRangeStart`            |                                                                                                                                                                                                                                                                                                                              |
| `commentReference`               | `Build.commentReference`             |                                                                                                                                                                                                                                                                                                                              |
| `CompatibilityMode`              | `DocxCompatibilityMode`              | `CompatibilityMode` is deprecated: Use `DocxCompatibilityMode`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `CompiledTemplate`               | `DocxCompiledTemplate`               | `CompiledTemplate` is deprecated: Use `DocxCompiledTemplate`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `compileTemplate`                | `Package.compileTemplate`            |                                                                                                                                                                                                                                                                                                                              |
| `ConversionAsset`                | `DocxConversionAsset`                | `ConversionAsset` is deprecated: Use `DocxConversionAsset`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `ConversionContext`              | `DocxConversionContext`              | `ConversionContext` is deprecated: Use `DocxConversionContext`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `ConversionWarning`              | `DocxConversionWarning`              | `ConversionWarning` is deprecated: Use `DocxConversionWarning`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `countByFormat`                  | `Query.getCountByFormat`             | `Query.countByFormat` is deprecated: Use `getCountByFormat`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `countWords`                     | `Query.getWordCount`                 | `Query.countWords` is deprecated: Use `getWordCount`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `createArrow`                    | `Build.arrow`                        | `Build.createArrow` is deprecated: Use `arrow`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `createCallout`                  | `Build.callout`                      | `Build.createCallout` is deprecated: Use `callout`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `createConversionContext`        | `Convert.createConversionContext`    |                                                                                                                                                                                                                                                                                                                              |
| `createDocxStream`               | `new Stream.DocumentWriter(options)` | `Stream.createDocxStream` is deprecated: Use `new Stream.DocumentWriter(options)`. Removed in 1.0.                                                                                                                                                                                                                           |
| `createEllipse`                  | `Build.ellipse`                      | `Build.createEllipse` is deprecated: Use `ellipse`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `createFlowchartShape`           | `Build.flowchartShape`               | `Build.createFlowchartShape` is deprecated: Use `flowchartShape`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `createHyphenator`               | `Font.createHyphenator`              |                                                                                                                                                                                                                                                                                                                              |
| `createLine`                     | `Build.line`                         | `Build.createLine` is deprecated: Use `line`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `createRect`                     | `Build.rect`                         | `Build.createRect` is deprecated: Use `rect`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `createRoundRect`                | `Build.roundRect`                    | `Build.createRoundRect` is deprecated: Use `roundRect`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `createShape`                    | `Build.shape`                        | `Build.createShape` is deprecated: Use `shape`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `CreateShapeOptions`             | `DocxCreateShapeOptions`             | `CreateShapeOptions` is deprecated: Use `DocxCreateShapeOptions`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `createStar`                     | `Build.star`                         | `Build.createStar` is deprecated: Use `star`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `DataSource`                     | `DocxDataSource`                     | `DataSource` is deprecated: Use `DocxDataSource`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `dateField`                      | `Build.dateField`                    |                                                                                                                                                                                                                                                                                                                              |
| `DEFAULT_SECURITY_POLICY`        | `Security.DEFAULT_POLICY`            | `Security.DEFAULT_SECURITY_POLICY` is deprecated: Use `DEFAULT_POLICY`. Removed in 1.0.                                                                                                                                                                                                                                      |
| `deletedRun`                     | `Build.deletedRun`                   |                                                                                                                                                                                                                                                                                                                              |
| `detectDirection`                | `Font.detectDirection`               |                                                                                                                                                                                                                                                                                                                              |
| `detectScript`                   | `Font.detectScript`                  |                                                                                                                                                                                                                                                                                                                              |
| `DiffChangeType`                 | `DocxDiffChangeType`                 | `DiffChangeType` is deprecated: Use `DocxDiffChangeType`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `DiffEntry`                      | `DocxDiffEntry`                      | `DiffEntry` is deprecated: Use `DocxDiffEntry`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `DiffResult`                     | `DocxDiffResult`                     | `DiffResult` is deprecated: Use `DocxDiffResult`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `DiffSummary`                    | `DocxDiffSummary`                    | `DiffSummary` is deprecated: Use `DocxDiffSummary`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `DocumentHandle`                 | `DocxDocument`                       | `DocumentHandle` is deprecated: Use `DocxDocument`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `DocumentHeading`                | `DocxHeading`                        | `DocumentHeading` is deprecated: Use `DocxHeading`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `DocumentProtectionOptions`      | `DocxProtectionOptions`              | `DocumentProtectionOptions` is deprecated: Use `DocxProtectionOptions`. Removed in 1.0.                                                                                                                                                                                                                                      |
| `DocumentSection`                | `DocxSectionInfo`                    | `DocumentSection` is deprecated: Use `DocxSectionInfo`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `DocxLimitExceededError#maximum` | `DocxLimitExceededError#allowed`     | deprecated: Use `allowed`. Removed in 1.0.                                                                                                                                                                                                                                                                                   |
| `DocxOptions`                    | —                                    | `DocxOptions` is deprecated: No function takes this type; use `DocxWriteOptions` (the options of `Document.toBuffer`), which has `compressionLevel` too. Removed in 1.0.                                                                                                                                                     |
| `docxToSemantic`                 | `Convert.toSemantic`                 | `Convert.docxToSemantic` is deprecated: Use `toSemantic`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `drawingShape`                   | `Build.shape`                        | `Build.drawingShape` is deprecated: Use `shape`, which builds the same shape and more: `fillColor` → `fill: { type: "solid", color }`, `noFill` → `fill: { type: "none" }`, `outlineColor` / `outlineWidth` / `noOutline` → `outline: { color, width, noLine }`, `textContent` → `textBody: { paragraphs }`. Removed in 1.0. |
| `editDocxIncremental`            | `Package.edit`                       | `Io.editDocxIncremental` is deprecated: Use `Package.edit`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `Effect3D`                       | `DocxEffect3D`                       | `Effect3D` is deprecated: Use `DocxEffect3D`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `eighthPointToPt`                | `Units.eighthPointToPt`              |                                                                                                                                                                                                                                                                                                                              |
| `EmbedFontOptions`               | `DocxEmbedFontOptions`               | `EmbedFontOptions` is deprecated: Use `DocxEmbedFontOptions`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `EmbedFontResult`                | `DocxEmbedFontResult`                | `EmbedFontResult` is deprecated: Use `DocxEmbedFontResult`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `emuToCm`                        | `Units.emuToCm`                      |                                                                                                                                                                                                                                                                                                                              |
| `emuToInches`                    | `Units.emuToInches`                  |                                                                                                                                                                                                                                                                                                                              |
| `emuToPx`                        | `Units.emuToPx`                      |                                                                                                                                                                                                                                                                                                                              |
| `EncryptOptions`                 | `DocxEncryptOptions`                 | `EncryptOptions` is deprecated: Use `DocxEncryptOptions`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `ENGLISH_US_PATTERNS`            | `Font.ENGLISH_US_PATTERNS`           |                                                                                                                                                                                                                                                                                                                              |
| `extractFormFields`              | `Query.listFormFields`               | `Query.extractFormFields` is deprecated: Use `listFormFields`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `extractText`                    | `Query.getText`                      | `Query.extractText` is deprecated: Use `getText`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `field`                          | `Build.field`                        |                                                                                                                                                                                                                                                                                                                              |
| `FieldUpdateOptions`             | `DocxFieldUpdateOptions`             | `FieldUpdateOptions` is deprecated: Use `DocxFieldUpdateOptions`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `fileNameField`                  | `Build.fileNameField`                |                                                                                                                                                                                                                                                                                                                              |
| `fileSizeField`                  | `Build.fileSizeField`                |                                                                                                                                                                                                                                                                                                                              |
| `fillFormFields`                 | `Transform.fillFormFields`           |                                                                                                                                                                                                                                                                                                                              |
| `fillTemplate`                   | `Package.fillTemplate`               |                                                                                                                                                                                                                                                                                                                              |
| `fillTemplateEnhanced`           | `Template.fill`                      | `Template.fillTemplateEnhanced` is deprecated: Use `fill` with `{ richPlaceholders: true }`. Removed in 1.0.                                                                                                                                                                                                                 |
| `fillTemplateFromBuffer`         | `Package.fillTemplate`               | `Io.fillTemplateFromBuffer` is deprecated: Use `Package.fillTemplate`. Removed in 1.0.                                                                                                                                                                                                                                       |
| `fillTemplateFromSource`         | `Template.fillFromSource`            | `Template.fillTemplateFromSource` is deprecated: Use `fillFromSource`. Removed in 1.0.                                                                                                                                                                                                                                       |
| `findBookmark`                   | `Query.getBookmark`                  | `Query.findBookmark` is deprecated: Use `getBookmark`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `findComment`                    | `Query.getComment`                   | `Query.findComment` is deprecated: Use `getComment`. Removed in 1.0.                                                                                                                                                                                                                                                         |
| `FlatOpcPart`                    | `DocxFlatOpcPart`                    | `FlatOpcPart` is deprecated: Use `DocxFlatOpcPart`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `floatingImage`                  | `Build.floatingImage`                |                                                                                                                                                                                                                                                                                                                              |
| `FontEmbedStyle`                 | `DocxFontEmbedStyle`                 | `FontEmbedStyle` is deprecated: Use `DocxFontEmbedStyle`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `FormatCriteria`                 | `DocxFormatCriteria`                 | `FormatCriteria` is deprecated: Use `DocxFormatCriteria`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `FormatSearchResult`             | `DocxFormatSearchResult`             | `FormatSearchResult` is deprecated: Use `DocxFormatSearchResult`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `formCheckboxField`              | `Build.formCheckboxField`            |                                                                                                                                                                                                                                                                                                                              |
| `formDropdownField`              | `Build.formDropdownField`            |                                                                                                                                                                                                                                                                                                                              |
| `FormFieldEntry`                 | `DocxFormFieldEntry`                 | `FormFieldEntry` is deprecated: Use `DocxFormFieldEntry`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `formTextField`                  | `Build.formTextField`                |                                                                                                                                                                                                                                                                                                                              |
| `FullLayoutOptions`              | `DocxFullLayoutOptions`              | `FullLayoutOptions` is deprecated: Use `DocxFullLayoutOptions`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `getCompatibilityMode`           | `Query.getCompatibilityMode`         |                                                                                                                                                                                                                                                                                                                              |
| `getHeadings`                    | `Query.listHeadings`                 | `Query.getHeadings` is deprecated: Use `listHeadings`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `getUsedFormats`                 | `Query.listUsedFormats`              | `Query.getUsedFormats` is deprecated: Use `listUsedFormats`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `GlossaryDocument`               | `DocxGlossaryDocument`               | `GlossaryDocument` is deprecated: Use `DocxGlossaryDocument`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `GlowEffect`                     | `DocxGlowEffect`                     | `GlowEffect` is deprecated: Use `DocxGlowEffect`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `GradientFill`                   | `DocxGradientFill`                   | `GradientFill` is deprecated: Use `DocxGradientFill`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `GradientStop`                   | `DocxGradientStop`                   | `GradientStop` is deprecated: Use `DocxGradientStop`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `gridBorders`                    | `Build.gridBorders`                  |                                                                                                                                                                                                                                                                                                                              |
| `halfPointToPt`                  | `Units.halfPointToPt`                |                                                                                                                                                                                                                                                                                                                              |
| `heading`                        | `Build.heading`                      |                                                                                                                                                                                                                                                                                                                              |
| `hyperlink`                      | `Build.hyperlink`                    |                                                                                                                                                                                                                                                                                                                              |
| `hyperlinkField`                 | `Build.hyperlinkField`               |                                                                                                                                                                                                                                                                                                                              |
| `hyphenateText`                  | `Font.hyphenateText`                 |                                                                                                                                                                                                                                                                                                                              |
| `hyphenateWord`                  | `Font.hyphenateWord`                 |                                                                                                                                                                                                                                                                                                                              |
| `HyphenationOptions`             | `DocxHyphenationOptions`             | `HyphenationOptions` is deprecated: Use `DocxHyphenationOptions`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `HyphenationPatterns`            | `DocxHyphenationPatterns`            | `HyphenationPatterns` is deprecated: Use `DocxHyphenationPatterns`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `IdGenerators`                   | `DocxIdGenerators`                   | `IdGenerators` is deprecated: Use `DocxIdGenerators`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `ifField`                        | `Build.ifField`                      |                                                                                                                                                                                                                                                                                                                              |
| `inchesToEmu`                    | `Units.inchesToEmu`                  |                                                                                                                                                                                                                                                                                                                              |
| `inchesToTwips`                  | `Units.inchesToTwips`                |                                                                                                                                                                                                                                                                                                                              |
| `includePictureField`            | `Build.includePictureField`          |                                                                                                                                                                                                                                                                                                                              |
| `includeTextField`               | `Build.includeTextField`             |                                                                                                                                                                                                                                                                                                                              |
| `IncrementalEdit`                | `DocxIncrementalEdit`                | `IncrementalEdit` is deprecated: Use `DocxIncrementalEdit`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `IncrementalEditOptions`         | `DocxIncrementalEditOptions`         | `IncrementalEditOptions` is deprecated: Use `DocxIncrementalEditOptions`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `indexEntryField`                | `Build.indexEntryField`              |                                                                                                                                                                                                                                                                                                                              |
| `indexField`                     | `Build.indexField`                   |                                                                                                                                                                                                                                                                                                                              |
| `insertedRun`                    | `Build.insertedRun`                  |                                                                                                                                                                                                                                                                                                                              |
| `isFlatOpc`                      | `Convert.isFlatOpc`                  |                                                                                                                                                                                                                                                                                                                              |
| `isTemplateChart`                | `Template.isChart`                   | `Template.isTemplateChart` is deprecated: Use `isChart`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `italic`                         | `Build.italic`                       |                                                                                                                                                                                                                                                                                                                              |
| `keywordsField`                  | `Build.keywordsField`                |                                                                                                                                                                                                                                                                                                                              |
| `LayoutAltChunk`                 | `DocxLayoutAltChunk`                 | `LayoutAltChunk` is deprecated: Use `DocxLayoutAltChunk`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `LayoutChart`                    | `DocxLayoutChart`                    | `LayoutChart` is deprecated: Use `DocxLayoutChart`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `LayoutCheckBox`                 | `DocxLayoutCheckBox`                 | `LayoutCheckBox` is deprecated: Use `DocxLayoutCheckBox`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `LayoutDocument`                 | `DocxLayoutDocument`                 | `LayoutDocument` is deprecated: Use `DocxLayoutDocument`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `LayoutFloat`                    | `DocxLayoutFloat`                    | `LayoutFloat` is deprecated: Use `DocxLayoutFloat`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `LayoutFloatWrap`                | `DocxLayoutFloatWrap`                | `LayoutFloatWrap` is deprecated: Use `DocxLayoutFloatWrap`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `LayoutImage`                    | `DocxLayoutImage`                    | `LayoutImage` is deprecated: Use `DocxLayoutImage`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `LayoutMath`                     | `DocxLayoutMath`                     | `LayoutMath` is deprecated: Use `DocxLayoutMath`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `LayoutOpaqueDrawing`            | `DocxLayoutOpaqueDrawing`            | `LayoutOpaqueDrawing` is deprecated: Use `DocxLayoutOpaqueDrawing`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `LayoutOptions`                  | `DocxLayoutOptions`                  | `LayoutOptions` is deprecated: Use `DocxLayoutOptions`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `LayoutPage`                     | `DocxLayoutPage`                     | `LayoutPage` is deprecated: Use `DocxLayoutPage`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `LayoutParagraph`                | `DocxLayoutParagraph`                | `LayoutParagraph` is deprecated: Use `DocxLayoutParagraph`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `LayoutResult`                   | `DocxLayoutResult`                   | `LayoutResult` is deprecated: Use `DocxLayoutResult`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `LayoutSdt`                      | `DocxLayoutSdt`                      | `LayoutSdt` is deprecated: Use `DocxLayoutSdt`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `LayoutShape`                    | `DocxLayoutShape`                    | `LayoutShape` is deprecated: Use `DocxLayoutShape`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `LayoutTable`                    | `DocxLayoutTable`                    | `LayoutTable` is deprecated: Use `DocxLayoutTable`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `LayoutTableCell`                | `DocxLayoutTableCell`                | `LayoutTableCell` is deprecated: Use `DocxLayoutTableCell`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `LayoutTableOfContents`          | `DocxLayoutTableOfContents`          | `LayoutTableOfContents` is deprecated: Use `DocxLayoutTableOfContents`. Removed in 1.0.                                                                                                                                                                                                                                      |
| `LayoutTextBox`                  | `DocxLayoutTextBox`                  | `LayoutTextBox` is deprecated: Use `DocxLayoutTextBox`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `LineBox`                        | `DocxLineBox`                        | `LineBox` is deprecated: Use `DocxLineBox`. Removed in 1.0.                                                                                                                                                                                                                                                                  |
| `LineBoxItem`                    | `DocxLineBoxItem`                    | `LineBoxItem` is deprecated: Use `DocxLineBoxItem`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `lineBreak`                      | `Build.lineBreak`                    |                                                                                                                                                                                                                                                                                                                              |
| `LineDash`                       | `DocxLineDash`                       | `LineDash` is deprecated: Use `DocxLineDash`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `LineEndSize`                    | `DocxLineEndSize`                    | `LineEndSize` is deprecated: Use `DocxLineEndSize`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `LineEndType`                    | `DocxLineEndType`                    | `LineEndType` is deprecated: Use `DocxLineEndType`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `lineMultiplierToSpacing`        | `Units.lineMultiplierToSpacing`      |                                                                                                                                                                                                                                                                                                                              |
| `listDocxParts`                  | `Package.listParts`                  | `Io.listDocxParts` is deprecated: Use `Package.listParts`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `listHyperlinks`                 | `Query.listHyperlinks`               |                                                                                                                                                                                                                                                                                                                              |
| `listImages`                     | `Query.listImages`                   |                                                                                                                                                                                                                                                                                                                              |
| `listRevisions`                  | `Query.listRevisions`                |                                                                                                                                                                                                                                                                                                                              |
| `listSections`                   | `Query.listSections`                 |                                                                                                                                                                                                                                                                                                                              |
| `listTables`                     | `Query.listTables`                   |                                                                                                                                                                                                                                                                                                                              |
| `listTemplateTags`               | `Template.listTags`                  | `Template.listTemplateTags` is deprecated: Use `listTags`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `mailMerge`                      | `Document.mailMerge`                 |                                                                                                                                                                                                                                                                                                                              |
| `mapDocument`                    | `Transform.map`                      | `Convert.mapDocument` is deprecated: Use `Transform.map`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `MapOptions`                     | `DocxMapOptions`                     | `MapOptions` is deprecated: Use `DocxMapOptions`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `MappingCondition`               | `DocxMappingCondition`               | `MappingCondition` is deprecated: Use `DocxMappingCondition`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `MappingSourceType`              | `DocxMappingSourceType`              | `MappingSourceType` is deprecated: Use `DocxMappingSourceType`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `MappingTarget`                  | `DocxMappingTarget`                  | `MappingTarget` is deprecated: Use `DocxMappingTarget`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `mathAccent`                     | `Build.mathAccent`                   |                                                                                                                                                                                                                                                                                                                              |
| `mathBar`                        | `Build.mathBar`                      |                                                                                                                                                                                                                                                                                                                              |
| `mathBlock`                      | `Build.mathBlock`                    |                                                                                                                                                                                                                                                                                                                              |
| `mathBorderBox`                  | `Build.mathBorderBox`                |                                                                                                                                                                                                                                                                                                                              |
| `mathBox`                        | `Build.mathBox`                      |                                                                                                                                                                                                                                                                                                                              |
| `mathDelimiter`                  | `Build.mathDelimiter`                |                                                                                                                                                                                                                                                                                                                              |
| `mathEquationArray`              | `Build.mathEquationArray`            |                                                                                                                                                                                                                                                                                                                              |
| `mathFraction`                   | `Build.mathFraction`                 |                                                                                                                                                                                                                                                                                                                              |
| `mathFunction`                   | `Build.mathFunction`                 |                                                                                                                                                                                                                                                                                                                              |
| `mathGroupChar`                  | `Build.mathGroupChar`                |                                                                                                                                                                                                                                                                                                                              |
| `mathIntegral`                   | `Build.mathIntegral`                 |                                                                                                                                                                                                                                                                                                                              |
| `mathLimit`                      | `Build.mathLimit`                    |                                                                                                                                                                                                                                                                                                                              |
| `mathMatrix`                     | `Build.mathMatrix`                   |                                                                                                                                                                                                                                                                                                                              |
| `mathMLToOmml`                   | `Convert.mathMlToOmml`               | `Convert.mathMLToOmml` is deprecated: Use `mathMlToOmml`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `mathNary`                       | `Build.mathNary`                     |                                                                                                                                                                                                                                                                                                                              |
| `mathPhantom`                    | `Build.mathPhantom`                  |                                                                                                                                                                                                                                                                                                                              |
| `mathPreSubSuperScript`          | `Build.mathPreSubSuperscript`        | `Build.mathPreSubSuperScript` is deprecated: Use `mathPreSubSuperscript`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `mathProduct`                    | `Build.mathProduct`                  |                                                                                                                                                                                                                                                                                                                              |
| `mathRoot`                       | `Build.mathRoot`                     |                                                                                                                                                                                                                                                                                                                              |
| `mathRun`                        | `Build.mathRun`                      |                                                                                                                                                                                                                                                                                                                              |
| `mathSqrt`                       | `Build.mathSqrt`                     |                                                                                                                                                                                                                                                                                                                              |
| `mathSubScript`                  | `Build.mathSubscript`                | `Build.mathSubScript` is deprecated: Use `mathSubscript`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `mathSubSuperScript`             | `Build.mathSubSuperscript`           | `Build.mathSubSuperScript` is deprecated: Use `mathSubSuperscript`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `mathSum`                        | `Build.mathSum`                      |                                                                                                                                                                                                                                                                                                                              |
| `mathSuperScript`                | `Build.mathSuperscript`              | `Build.mathSuperScript` is deprecated: Use `mathSuperscript`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `MediaAsset`                     | —                                    | `MediaAsset` is deprecated: Reserved for a package-model API that was never added: no function takes or returns it; a document's images are `ImageDef`s. Removed in 1.0.                                                                                                                                                     |
| `MergeOptions`                   | `DocxMergeOptions`                   | `MergeOptions` is deprecated: Use `DocxMergeOptions`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `mmToTwips`                      | `Units.mmToTwips`                    |                                                                                                                                                                                                                                                                                                                              |
| `movedFromRun`                   | `Build.movedFromRun`                 |                                                                                                                                                                                                                                                                                                                              |
| `movedToRun`                     | `Build.movedToRun`                   |                                                                                                                                                                                                                                                                                                                              |
| `moveFromRangeEnd`               | `Build.moveFromRangeEnd`             |                                                                                                                                                                                                                                                                                                                              |
| `moveFromRangeStart`             | `Build.moveFromRangeStart`           |                                                                                                                                                                                                                                                                                                                              |
| `moveToRangeEnd`                 | `Build.moveToRangeEnd`               |                                                                                                                                                                                                                                                                                                                              |
| `moveToRangeStart`               | `Build.moveToRangeStart`             |                                                                                                                                                                                                                                                                                                                              |
| `noBreakHyphen`                  | `Build.noBreakHyphen`                |                                                                                                                                                                                                                                                                                                                              |
| `NoFill`                         | `DocxNoFill`                         | `NoFill` is deprecated: Use `DocxNoFill`. Removed in 1.0.                                                                                                                                                                                                                                                                    |
| `noteRefField`                   | `Build.noteRefField`                 |                                                                                                                                                                                                                                                                                                                              |
| `OleDisplayAs`                   | `DocxOleDisplayAs`                   | `OleDisplayAs` is deprecated: Use `DocxOleDisplayAs`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `OleEmbeddingResult`             | `DocxOleEmbeddingResult`             | `OleEmbeddingResult` is deprecated: Use `DocxOleEmbeddingResult`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `OleExtractionResult`            | `DocxOleExtractionResult`            | `OleExtractionResult` is deprecated: Use `DocxOleExtractionResult`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `OleObject`                      | `DocxOleObject`                      | `OleObject` is deprecated: Use `DocxOleObject`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `OleObjectType`                  | `DocxOleObjectType`                  | `OleObjectType` is deprecated: Use `DocxOleObjectType`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `ommlToMathML`                   | `Convert.ommlToMathMl`               | `Convert.ommlToMathML` is deprecated: Use `ommlToMathMl`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `pageBreak`                      | `Build.pageBreak`                    |                                                                                                                                                                                                                                                                                                                              |
| `PageContent`                    | `DocxPageContent`                    | `PageContent` is deprecated: Use `DocxPageContent`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `PageGeometry`                   | `DocxPageGeometry`                   | `PageGeometry` is deprecated: Use `DocxPageGeometry`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `PageGeometryOverride`           | `DocxPageGeometryOverride`           | `PageGeometryOverride` is deprecated: Use `DocxPageGeometryOverride`. Removed in 1.0.                                                                                                                                                                                                                                        |
| `pageNumberField`                | `Build.pageNumberField`              |                                                                                                                                                                                                                                                                                                                              |
| `pageRefField`                   | `Build.pageRefField`                 |                                                                                                                                                                                                                                                                                                                              |
| `paragraph`                      | `Build.paragraph`                    |                                                                                                                                                                                                                                                                                                                              |
| `paragraphCount`                 | `Query.getParagraphCount`            | `Query.paragraphCount` is deprecated: Use `getParagraphCount`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `parseFlatOpc`                   | `Convert.flatOpcToParts`             | `Convert.parseFlatOpc` is deprecated: Use `flatOpcToParts` for the part map, or `Document.read` for the document a Flat OPC file holds. Removed in 1.0.                                                                                                                                                                      |
| `PatchContent`                   | `DocxPatchContent`                   | `PatchContent` is deprecated: Use `DocxPatchContent`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `patchDocument`                  | `Package.patch`                      | `Io.patchDocument` is deprecated: Use `Package.patch`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `PatchOperation`                 | `DocxPatchOperation`                 | `PatchOperation` is deprecated: Use `DocxPatchOperation`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `PatchOptions`                   | `DocxPatchOptions`                   | `PatchOptions` is deprecated: Use `DocxPatchOptions`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `patchTemplate`                  | `Package.patchTemplate`              |                                                                                                                                                                                                                                                                                                                              |
| `PatternFill`                    | `DocxPatternFill`                    | `PatternFill` is deprecated: Use `DocxPatternFill`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `percentToTablePct`              | `Units.percentToTablePct`            |                                                                                                                                                                                                                                                                                                                              |
| `positionalTab`                  | `Build.positionalTab`                |                                                                                                                                                                                                                                                                                                                              |
| `PositionedInlineImage`          | `DocxPositionedInlineImage`          | `PositionedInlineImage` is deprecated: Use `DocxPositionedInlineImage`. Removed in 1.0.                                                                                                                                                                                                                                      |
| `PositionedRun`                  | `DocxPositionedRun`                  | `PositionedRun` is deprecated: Use `DocxPositionedRun`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `ProtectionEditType`             | `DocxProtectionEditType`             | `ProtectionEditType` is deprecated: Use `DocxProtectionEditType`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `ProtectionHashAlgorithm`        | `DocxProtectionHashAlgorithm`        | `ProtectionHashAlgorithm` is deprecated: Use `DocxProtectionHashAlgorithm`. Removed in 1.0.                                                                                                                                                                                                                                  |
| `ProtectionState`                | `DocxProtectionState`                | `ProtectionState` is deprecated: Use `DocxProtectionState`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `ptToEighthPoint`                | `Units.ptToEighthPoint`              |                                                                                                                                                                                                                                                                                                                              |
| `ptToEmu`                        | `Units.ptToEmu`                      |                                                                                                                                                                                                                                                                                                                              |
| `ptToHalfPoint`                  | `Units.ptToHalfPoint`                |                                                                                                                                                                                                                                                                                                                              |
| `ptToTwips`                      | `Units.ptToTwips`                    |                                                                                                                                                                                                                                                                                                                              |
| `pxToEmu`                        | `Units.pxToEmu`                      |                                                                                                                                                                                                                                                                                                                              |
| `quoteField`                     | `Build.quoteField`                   |                                                                                                                                                                                                                                                                                                                              |
| `ReadDocxOptions`                | `DocxReadOptions`                    | `ReadDocxOptions` is deprecated: Use `DocxReadOptions`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `readDocxPart`                   | `Package.readPart`                   | `Io.readDocxPart` is deprecated: Use `Package.readPart`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `readOdt`                        | `Convert.fromOdt`                    | `Convert.readOdt` is deprecated: Use `fromOdt`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `refField`                       | `Build.refField`                     |                                                                                                                                                                                                                                                                                                                              |
| `ReflectionEffect`               | `DocxReflectionEffect`               | `ReflectionEffect` is deprecated: Use `DocxReflectionEffect`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `rejectAllRevisions`             | `Document.rejectAllRevisions`        |                                                                                                                                                                                                                                                                                                                              |
| `rejectRevision`                 | `Document.rejectRevision`            |                                                                                                                                                                                                                                                                                                                              |
| `renderDocumentToSvg`            | `Layout.toSvg`                       | `Layout.renderDocumentToSvg` is deprecated: Use `toSvg`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `RenderOptions`                  | `DocxRenderOptions`                  | `RenderOptions` is deprecated: Use `DocxRenderOptions`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `renderPageFromLayout`           | `Layout.computedPageToSvg`           | `Layout.renderPageFromLayout` is deprecated: Use `computedPageToSvg`. Removed in 1.0.                                                                                                                                                                                                                                        |
| `renderPageToSvg`                | `Layout.pageToSvg`                   | `Layout.renderPageToSvg` is deprecated: Use `pageToSvg`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `replaceText`                    | `Document.replaceText`               |                                                                                                                                                                                                                                                                                                                              |
| `resolveDataBindings`            | `Transform.resolveDataBindings`      |                                                                                                                                                                                                                                                                                                                              |
| `ResolvedFormatting`             | `DocxResolvedFormatting`             | `ResolvedFormatting` is deprecated: Use `DocxResolvedFormatting`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `ResolvedNumberingLevel`         | `DocxResolvedNumberingLevel`         | `ResolvedNumberingLevel` is deprecated: Use `DocxResolvedNumberingLevel`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `ResolvedParagraphStyle`         | `DocxResolvedParagraphStyle`         | `ResolvedParagraphStyle` is deprecated: Use `DocxResolvedParagraphStyle`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `ResolvedRunStyle`               | `DocxResolvedRunStyle`               | `ResolvedRunStyle` is deprecated: Use `DocxResolvedRunStyle`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `resolveNumberingLevel`          | `Query.getNumberingLevel`            | `Query.resolveNumberingLevel` is deprecated: Use `getNumberingLevel`. Removed in 1.0.                                                                                                                                                                                                                                        |
| `resolveRunStyle`                | `Query.getRunStyle`                  | `Query.resolveRunStyle` is deprecated: Use `getRunStyle`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `resolveSecurityPolicy`          | `Security.resolvePolicy`             | `Security.resolveSecurityPolicy` is deprecated: Use `resolvePolicy`. Removed in 1.0.                                                                                                                                                                                                                                         |
| `resolveStyle`                   | `Query.getParagraphStyle`            | `Query.resolveStyle` is deprecated: Use `getParagraphStyle`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `resolveTableStyle`              | `Query.getTableStyle`                | `Query.resolveTableStyle` is deprecated: Use `getTableStyle`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `RevisionEntry`                  | `DocxRevisionEntry`                  | `RevisionEntry` is deprecated: Use `DocxRevisionEntry`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `row`                            | `Build.row`                          |                                                                                                                                                                                                                                                                                                                              |
| `ruby`                           | `Build.ruby`                         |                                                                                                                                                                                                                                                                                                                              |
| `ScriptType`                     | `DocxScriptType`                     | `ScriptType` is deprecated: Use `DocxScriptType`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `searchByFormat`                 | `Query.searchByFormat`               |                                                                                                                                                                                                                                                                                                                              |
| `SearchResult`                   | `DocxSearchResult`                   | `SearchResult` is deprecated: Use `DocxSearchResult`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `searchText`                     | `Query.searchText`                   |                                                                                                                                                                                                                                                                                                                              |
| `sectionField`                   | `Build.sectionField`                 |                                                                                                                                                                                                                                                                                                                              |
| `sectionPagesField`              | `Build.sectionPagesField`            |                                                                                                                                                                                                                                                                                                                              |
| `SemanticBlock`                  | `DocxSemanticBlock`                  | `SemanticBlock` is deprecated: Use `DocxSemanticBlock`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `SemanticDocument`               | `DocxSemanticDocument`               | `SemanticDocument` is deprecated: Use `DocxSemanticDocument`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `SemanticInline`                 | `DocxSemanticInline`                 | `SemanticInline` is deprecated: Use `DocxSemanticInline`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `sequenceField`                  | `Build.sequenceField`                |                                                                                                                                                                                                                                                                                                                              |
| `setCompatibilityMode`           | `Document.setCompatibilityMode`      |                                                                                                                                                                                                                                                                                                                              |
| `ShadowEffect`                   | `DocxShadowEffect`                   | `ShadowEffect` is deprecated: Use `DocxShadowEffect`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `ShapedCluster`                  | `DocxShapedCluster`                  | `ShapedCluster` is deprecated: Use `DocxShapedCluster`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `ShapeEffects`                   | `DocxShapeEffects`                   | `ShapeEffects` is deprecated: Use `DocxShapeEffects`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `ShapeFill`                      | `DocxShapeFill`                      | `ShapeFill` is deprecated: Use `DocxShapeFill`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `ShapeOutline`                   | `DocxShapeOutline`                   | `ShapeOutline` is deprecated: Use `DocxShapeOutline`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `shapeText`                      | `Font.shapeText`                     |                                                                                                                                                                                                                                                                                                                              |
| `ShapeTextBody`                  | `DocxShapeTextBody`                  | `ShapeTextBody` is deprecated: Use `DocxShapeTextBody`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `ShapingOptions`                 | `DocxShapingOptions`                 | `ShapingOptions` is deprecated: Use `DocxShapingOptions`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `simpleTable`                    | `Build.simpleTable`                  |                                                                                                                                                                                                                                                                                                                              |
| `softHyphen`                     | `Build.softHyphen`                   |                                                                                                                                                                                                                                                                                                                              |
| `SolidFill`                      | `DocxSolidFill`                      | `SolidFill` is deprecated: Use `DocxSolidFill`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `spacingToLineMultiplier`        | `Units.spacingToLineMultiplier`      |                                                                                                                                                                                                                                                                                                                              |
| `SplitOptions`                   | `DocxSplitOptions`                   | `SplitOptions` is deprecated: Use `DocxSplitOptions`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `StandardShapeType`              | `DocxStandardShapeType`              | `StandardShapeType` is deprecated: Use `DocxStandardShapeType`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `StreamingDocxOptions`           | `DocxStreamingWriteOptions`          | `StreamingDocxOptions` is deprecated: Use `DocxStreamingWriteOptions`. Removed in 1.0.                                                                                                                                                                                                                                       |
| `StreamingDocxWriter`            | `Stream.DocumentWriter`              | `Stream.StreamingDocxWriter` is deprecated: Use `DocumentWriter`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `StreamingDocxWriter#finalize`   | `Stream.DocumentWriter#commit`       | deprecated: Use `commit` to finish the document and `toBuffer` to obtain its bytes. Removed in 1.0.                                                                                                                                                                                                                          |
| `StreamingProgressCallback`      | `DocxStreamingProgressCallback`      | `StreamingProgressCallback` is deprecated: Use `DocxStreamingProgressCallback`. Removed in 1.0.                                                                                                                                                                                                                              |
| `STRICT_SECURITY_POLICY`         | `Security.STRICT_POLICY`             | `Security.STRICT_SECURITY_POLICY` is deprecated: Use `STRICT_POLICY`. Removed in 1.0.                                                                                                                                                                                                                                        |
| `strikethrough`                  | `Build.strikethrough`                |                                                                                                                                                                                                                                                                                                                              |
| `structuredDocumentTag`          | `Build.structuredDocumentTag`        |                                                                                                                                                                                                                                                                                                                              |
| `StyleMap`                       | `DocxStyleMap`                       | `StyleMap` is deprecated: Use `DocxStyleMap`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `StyleMapOptions`                | `DocxStyleMapOptions`                | `StyleMapOptions` is deprecated: Use `DocxStyleMapOptions`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `StyleMappingRule`               | `DocxStyleMappingRule`               | `StyleMappingRule` is deprecated: Use `DocxStyleMappingRule`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `styleRefField`                  | `Build.styleRefField`                |                                                                                                                                                                                                                                                                                                                              |
| `StyleResolveContext`            | `DocxStyleResolveContext`            | `StyleResolveContext` is deprecated: Use `DocxStyleResolveContext`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `subjectField`                   | `Build.subjectField`                 |                                                                                                                                                                                                                                                                                                                              |
| `symbol`                         | `Build.symbol`                       |                                                                                                                                                                                                                                                                                                                              |
| `tab`                            | `Build.tab`                          |                                                                                                                                                                                                                                                                                                                              |
| `table`                          | `Build.table`                        |                                                                                                                                                                                                                                                                                                                              |
| `tableCount`                     | `Query.getTableCount`                | `Query.tableCount` is deprecated: Use `getTableCount`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `tablePctToPercent`              | `Units.tablePctToPercent`            |                                                                                                                                                                                                                                                                                                                              |
| `tcField`                        | `Build.tcField`                      |                                                                                                                                                                                                                                                                                                                              |
| `TemplateChart`                  | `DocxTemplateChart`                  | `TemplateChart` is deprecated: Use `DocxTemplateChart`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `TemplateError`                  | `DocxTemplateError`                  | `TemplateError` is deprecated: Use `DocxTemplateError`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `TemplateHtmlChunk`              | `DocxTemplateHtmlChunk`              | `TemplateHtmlChunk` is deprecated: Use `DocxTemplateHtmlChunk`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `TemplateImage`                  | `DocxTemplateImage`                  | `TemplateImage` is deprecated: Use `DocxTemplateImage`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `TemplateOptions`                | `DocxTemplateOptions`                | `TemplateOptions` is deprecated: Use `DocxTemplateOptions`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `TemplateRichText`               | `DocxTemplateRichText`               | `TemplateRichText` is deprecated: Use `DocxTemplateRichText`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `TemplateSubDocument`            | `DocxTemplateSubDocument`            | `TemplateSubDocument` is deprecated: Use `DocxTemplateSubDocument`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `TemplateTag`                    | `DocxTemplateTag`                    | `TemplateTag` is deprecated: Use `DocxTemplateTag`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `text`                           | `TextContent.text`                   |                                                                                                                                                                                                                                                                                                                              |
| `textParagraph`                  | `Build.textParagraph`                |                                                                                                                                                                                                                                                                                                                              |
| `TextVerticalAnchor`             | `DocxTextVerticalAnchor`             | `TextVerticalAnchor` is deprecated: Use `DocxTextVerticalAnchor`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `TextWrap`                       | `DocxTextWrap`                       | `TextWrap` is deprecated: Use `DocxTextWrap`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `timeField`                      | `Build.timeField`                    |                                                                                                                                                                                                                                                                                                                              |
| `titleField`                     | `Build.titleField`                   |                                                                                                                                                                                                                                                                                                                              |
| `toBase64`                       | `Document.toBase64`                  |                                                                                                                                                                                                                                                                                                                              |
| `toBuffer`                       | `Document.toBuffer`                  |                                                                                                                                                                                                                                                                                                                              |
| `tocField`                       | `Build.tocField`                     |                                                                                                                                                                                                                                                                                                                              |
| `toFlatOpc`                      | `Document.toFlatOpc`                 |                                                                                                                                                                                                                                                                                                                              |
| `toFlatOpcFromDoc`               | —                                    | `Io.toFlatOpcFromDoc` is deprecated: Use `toFlatOpc`, which takes an options object. Removed in 1.0.                                                                                                                                                                                                                         |
| `totalPagesField`                | `Build.totalPagesField`              |                                                                                                                                                                                                                                                                                                                              |
| `twipsToCm`                      | `Units.twipsToCm`                    |                                                                                                                                                                                                                                                                                                                              |
| `twipsToInches`                  | `Units.twipsToInches`                |                                                                                                                                                                                                                                                                                                                              |
| `twipsToPt`                      | `Units.twipsToPt`                    |                                                                                                                                                                                                                                                                                                                              |
| `underline`                      | `Build.underline`                    |                                                                                                                                                                                                                                                                                                                              |
| `updateFields`                   | `Transform.updateFields`             |                                                                                                                                                                                                                                                                                                                              |
| `updateTableOfContents`          | `Transform.updateTableOfContents`    |                                                                                                                                                                                                                                                                                                                              |
| `ValidationIssue`                | `DocxValidationIssue`                | `ValidationIssue` is deprecated: Use `DocxValidationIssue`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `ValidationOptions`              | `DocxValidationOptions`              | `ValidationOptions` is deprecated: Use `DocxValidationOptions`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `ValidationResult`               | `DocxValidationResult`               | `ValidationResult` is deprecated: Use `DocxValidationResult`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `ValidationSeverity`             | `DocxValidationSeverity`             | `ValidationSeverity` is deprecated: Use `DocxValidationSeverity`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `VbaProjectInfo`                 | `DocxVbaProjectInfo`                 | `VbaProjectInfo` is deprecated: Use `DocxVbaProjectInfo`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `VisitAction`                    | `DocxVisitAction`                    | `VisitAction` is deprecated: Use `DocxVisitAction`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `walkBlocks`                     | `Query.eachBlock`                    | `Query.walkBlocks` is deprecated: Use `eachBlock`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `walkDocument`                   | `Query.each`                         | `Query.walkDocument` is deprecated: Use `each`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `WalkOptions`                    | `DocxWalkOptions`                    | `WalkOptions` is deprecated: Use `DocxWalkOptions`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `WalkPath`                       | `DocxWalkPath`                       | `WalkPath` is deprecated: Use `DocxWalkPath`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `WordRenderContext`              | `DocxRenderContext`                  | `WordRenderContext` is deprecated: Use `DocxRenderContext`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `WordSecurityPolicy`             | `DocxSecurityPolicy`                 | `WordSecurityPolicy` is deprecated: Use `DocxSecurityPolicy`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `writeOdt`                       | `Convert.toOdt`                      | `Convert.writeOdt` is deprecated: Use `toOdt`. Removed in 1.0.                                                                                                                                                                                                                                                               |

**`@cj-tech-master/excelts/word/crypto`** → `documonster/word/crypto`

| excelts                  | documonster                            | Note                                                                                                                                                                     |
| ------------------------ | -------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `AGILE_BLOCK_KEYS`       | `WordCrypto.Agile.BLOCK_KEYS`          | `AGILE_BLOCK_KEYS` is deprecated: Use `WordCrypto.Agile.BLOCK_KEYS`. Removed in 1.0.                                                                                     |
| `AgileEncryptionInfo`    | `DocxAgileEncryptionInfo`              | `AgileEncryptionInfo` is deprecated: Use `DocxAgileEncryptionInfo`. Removed in 1.0.                                                                                      |
| `CfbEntry`               | `DocxCfbEntry`                         | `CfbEntry` is deprecated: Use `DocxCfbEntry`. Removed in 1.0.                                                                                                            |
| `decryptDocx`            | `WordCrypto.decrypt`                   | `decryptDocx` is deprecated: Use `WordCrypto.decrypt`. Removed in 1.0.                                                                                                   |
| `decryptPackage`         | `WordCrypto.Agile.decryptPackage`      |                                                                                                                                                                          |
| `deobfuscateFont`        | —                                      | `deobfuscateFont` is deprecated: Use `Font.deobfuscate` from `documonster/word`. Removed in 1.0.                                                                         |
| `deriveEncryptionKey`    | `WordCrypto.Agile.deriveKey`           | `deriveEncryptionKey` is deprecated: Use `WordCrypto.Agile.deriveKey`, which takes `(info, password, blockKey)` instead of `(password, info, blockKey)`. Removed in 1.0. |
| `DigitalSignatureInfo`   | `DocxDigitalSignatureInfo`             | `DigitalSignatureInfo` is deprecated: Use `DocxDigitalSignatureInfo`. Removed in 1.0.                                                                                    |
| `encryptDocx`            | `WordCrypto.encrypt`                   | `encryptDocx` is deprecated: Use `WordCrypto.encrypt`. Removed in 1.0.                                                                                                   |
| `EncryptOptions`         | `DocxEncryptOptions`                   | `EncryptOptions` is deprecated: Use `DocxEncryptOptions`. Removed in 1.0.                                                                                                |
| `extractSignatures`      | `WordCrypto.listSignatures`            | `extractSignatures` is deprecated: Use `WordCrypto.listSignatures`. Removed in 1.0.                                                                                      |
| `generateFontKey`        | —                                      | `generateFontKey` is deprecated: Use `Font.generateObfuscationKey` from `documonster/word`. Removed in 1.0.                                                              |
| `hasDigitalSignatures`   | `WordCrypto.hasSignatures`             | `hasDigitalSignatures` is deprecated: Use `WordCrypto.hasSignatures`, which takes the document or the package's part names (`Package.listParts(bytes)`). Removed in 1.0. |
| `isEncryptedDocx`        | `WordCrypto.isEncrypted`               | `isEncryptedDocx` is deprecated: Use `WordCrypto.isEncrypted`. Removed in 1.0.                                                                                           |
| `isWellFormedSignature`  | `WordCrypto.isWellFormedSignature`     |                                                                                                                                                                          |
| `obfuscateFont`          | —                                      | `obfuscateFont` is deprecated: Use `Font.obfuscate` from `documonster/word`. Removed in 1.0.                                                                             |
| `parseEncryptionInfoXml` | `WordCrypto.Agile.parseEncryptionInfo` | `parseEncryptionInfoXml` is deprecated: Use `WordCrypto.Agile.parseEncryptionInfo`. Removed in 1.0.                                                                      |
| `parseSignatureXml`      | `WordCrypto.parseSignature`            | `parseSignatureXml` is deprecated: Use `WordCrypto.parseSignature`. Removed in 1.0.                                                                                      |
| `readCfb`                | `WordCrypto.Cfb.read`                  | `readCfb` is deprecated: Use `WordCrypto.Cfb.read`. Removed in 1.0.                                                                                                      |
| `verifyPassword`         | `WordCrypto.verifyPassword`            |                                                                                                                                                                          |
| `writeCfb`               | `WordCrypto.Cfb.write`                 | `writeCfb` is deprecated: Use `WordCrypto.Cfb.write`. Removed in 1.0.                                                                                                    |

**`@cj-tech-master/excelts/word/excel`** → `documonster/word/excel`

| excelts                     | documonster               | Note                                                                                    |
| --------------------------- | ------------------------- | --------------------------------------------------------------------------------------- |
| `buildWordChartExXml`       | `WordExcel.toChartExXml`  | `buildWordChartExXml` is deprecated: Use `WordExcel.toChartExXml`. Removed in 1.0.      |
| `ExcelToDocxOptions`        | `DocxFromExcelOptions`    | `ExcelToDocxOptions` is deprecated: Use `DocxFromExcelOptions`. Removed in 1.0.         |
| `extractTablesToExcel`      | `WordExcel.extractTables` | `extractTablesToExcel` is deprecated: Use `WordExcel.extractTables`. Removed in 1.0.    |
| `generateChartEmbeddedXlsx` | `WordExcel.toChartXlsx`   | `generateChartEmbeddedXlsx` is deprecated: Use `WordExcel.toChartXlsx`. Removed in 1.0. |
| `renderWordChartSvg`        | `WordExcel.toChartSvg`    | `renderWordChartSvg` is deprecated: Use `WordExcel.toChartSvg`. Removed in 1.0.         |
| `WordChartExOptions`        | `DocxChartExOptions`      | `WordChartExOptions` is deprecated: Use `DocxChartExOptions`. Removed in 1.0.           |
| `wordChartToChartModel`     | `WordExcel.toChartModel`  | `wordChartToChartModel` is deprecated: Use `WordExcel.toChartModel`. Removed in 1.0.    |

**`@cj-tech-master/excelts/word/html`** → `documonster/word/html`

| excelts             | documonster             | Note                                                                                                                                                                                      |
| ------------------- | ----------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `HtmlImportOptions` | `DocxHtmlImportOptions` | `HtmlImportOptions` is deprecated: Use `DocxHtmlImportOptions`. Removed in 1.0.                                                                                                           |
| `HtmlRenderOptions` | `DocxHtmlRenderOptions` | `HtmlRenderOptions` is deprecated: Use `DocxHtmlRenderOptions`. Removed in 1.0.                                                                                                           |
| `HtmlRenderResult`  | `DocxHtmlRenderResult`  | `HtmlRenderResult` is deprecated: Use `DocxHtmlRenderResult`. Removed in 1.0.                                                                                                             |
| `htmlToDocx`        | `WordHtml.toDocxBody`   | `htmlToDocx` is deprecated: Use `WordHtml.toDocxBody`, which also returns the styles and list numbering the body refers to, or `WordHtml.toDocx` for a complete document. Removed in 1.0. |
| `htmlToDocxBody`    | `WordHtml.toDocxBody`   | `htmlToDocxBody` is deprecated: Use `WordHtml.toDocxBody``(html).body`. Removed in 1.0.                                                                                                   |
| `HtmlToDocxResult`  | `WordHtml.toDocxBody`   | `HtmlToDocxResult` is deprecated: Use `WordHtml.toDocxBody`, whose result is `DocxHtmlBodyResult`. Removed in 1.0.                                                                        |
| `renderToHtml`      | `WordHtml.fromDocx`     | `renderToHtml` is deprecated: Use `WordHtml.fromDocx`. Removed in 1.0.                                                                                                                    |

**`@cj-tech-master/excelts/word/markdown`** → `documonster/word/markdown`

| excelts                 | documonster                 | Note                                                                                    |
| ----------------------- | --------------------------- | --------------------------------------------------------------------------------------- |
| `MarkdownBodyResult`    | `DocxMarkdownBodyResult`    | `MarkdownBodyResult` is deprecated: Use `DocxMarkdownBodyResult`. Removed in 1.0.       |
| `MarkdownImageData`     | `DocxMarkdownImageData`     | `MarkdownImageData` is deprecated: Use `DocxMarkdownImageData`. Removed in 1.0.         |
| `MarkdownImportOptions` | `DocxMarkdownImportOptions` | `MarkdownImportOptions` is deprecated: Use `DocxMarkdownImportOptions`. Removed in 1.0. |
| `MarkdownRenderOptions` | `DocxMarkdownRenderOptions` | `MarkdownRenderOptions` is deprecated: Use `DocxMarkdownRenderOptions`. Removed in 1.0. |
| `markdownToDocx`        | `WordMarkdown.toDocx`       | `markdownToDocx` is deprecated: Use `WordMarkdown.toDocx`. Removed in 1.0.              |
| `markdownToDocxBody`    | `WordMarkdown.toDocxBody`   | `markdownToDocxBody` is deprecated: Use `WordMarkdown.toDocxBody`. Removed in 1.0.      |
| `renderToMarkdown`      | `WordMarkdown.fromDocx`     | `renderToMarkdown` is deprecated: Use `WordMarkdown.fromDocx`. Removed in 1.0.          |

**`@cj-tech-master/excelts/xml`** → `documonster/xml`

| excelts                   | documonster               | Note                                                                                                                                                                    |
| ------------------------- | ------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `attr`                    | `Xml.getAttr`             | `Xml.attr` is deprecated: Use `getAttr`. Removed in 1.0.                                                                                                                |
| `encodeCData`             | `Xml.encodeCData`         |                                                                                                                                                                         |
| `findChild`               | `Xml.findChild`           |                                                                                                                                                                         |
| `findChildren`            | `Xml.listChildren`        | `Xml.findChildren` is deprecated: Use `listChildren`. Removed in 1.0.                                                                                                   |
| `InvalidCharHandling`     | `XmlInvalidCharHandling`  | `InvalidCharHandling` is deprecated: Use `XmlInvalidCharHandling`. Removed in 1.0.                                                                                      |
| `parseSax`                | `Xml.parseSaxStream`      | `Xml.parseSax` is deprecated: Use `parseSaxStream`. An incremental reader takes the `Stream` suffix, as `Csv.parseStream` and `Workbook.readStream` do. Removed in 1.0. |
| `ParseXmlToObjectOptions` | `XmlParseToObjectOptions` | `ParseXmlToObjectOptions` is deprecated: Use `XmlParseToObjectOptions`. Removed in 1.0.                                                                                 |
| `query`                   | `Xml.query`               |                                                                                                                                                                         |
| `queryAll`                | `Xml.queryAll`            |                                                                                                                                                                         |
| `SaxEvent`                | `XmlSaxEvent`             | `SaxEvent` is deprecated: Use `XmlSaxEvent`. Removed in 1.0.                                                                                                            |
| `SaxEventAny`             | `XmlSaxEventAny`          | `SaxEventAny` is deprecated: Use `XmlSaxEventAny`. Removed in 1.0.                                                                                                      |
| `SaxHandlers`             | `XmlSaxHandlers`          | `SaxHandlers` is deprecated: Use `XmlSaxHandlers`. Removed in 1.0.                                                                                                      |
| `SaxOptions`              | `XmlSaxOptions`           | `SaxOptions` is deprecated: Use `XmlSaxOptions`. Removed in 1.0.                                                                                                        |
| `SaxParser`               | `Xml.SaxParser`           |                                                                                                                                                                         |
| `saxStream`               | `Xml.feedSax`             | `Xml.saxStream` is deprecated: Use `feedSax`. Removed in 1.0.                                                                                                           |
| `SaxTag`                  | `XmlSaxTag`               | `SaxTag` is deprecated: Use `XmlSaxTag`. Removed in 1.0.                                                                                                                |
| `StdDocAttributes`        | `Xml.DEFAULT_DECLARATION` | `Xml.StdDocAttributes` is deprecated: Use `DEFAULT_DECLARATION`. Removed in 1.0.                                                                                        |
| `textContent`             | `Xml.getText`             | `Xml.textContent` is deprecated: Use `getText`. Removed in 1.0.                                                                                                         |
| `toPlainObject`           | `Xml.toPlainObject`       |                                                                                                                                                                         |
| `ToPlainObjectOptions`    | `XmlToPlainObjectOptions` | `ToPlainObjectOptions` is deprecated: Use `XmlToPlainObjectOptions`. Removed in 1.0.                                                                                    |
| `validateCommentText`     | `Xml.validateCommentText` |                                                                                                                                                                         |
| `validateXmlName`         | `Xml.validateName`        | `Xml.validateXmlName` is deprecated: Use `validateName`. Removed in 1.0.                                                                                                |
| `walk`                    | `Xml.eachElement`         | `Xml.walk` is deprecated: Use `eachElement`. Removed in 1.0.                                                                                                            |
| `WritableTarget`          | `XmlWritableTarget`       | `WritableTarget` is deprecated: Use `XmlWritableTarget`. Removed in 1.0.                                                                                                |
| `XmlWriter#save`          | `Xml.Writer#snapshot`     | deprecated: Use `snapshot`. Removed in 1.0.                                                                                                                             |
| `XmlWriter#xml`           | `Xml.Writer#toString`     | deprecated: Use `toString`. Removed in 1.0.                                                                                                                             |

<!-- excelts-index:end -->
