# Migration Guide: `exceljs` → `documonster`

> **Version.** This guide targets **documonster 0.22.0 or later**, the first release with the namespace API it
> describes. Check what npm serves with `npm view documonster version`: until 0.22.0 is published,
> `npm install documonster` installs 0.21.1, which this guide does not describe. Install with
> `npm install documonster@^0.22.0` (§1).

This guide maps the [`exceljs`](https://github.com/exceljs/exceljs) API (4.x) onto `documonster`, member by member. It is
written to be applied mechanically — by a person or by a code generator — and every replacement it names is the
current, non-deprecated form.

`documonster` began as a fork of exceljs (through `@cj-tech-master/excelts`), so the **file model is the same**: the
cell value shapes, the style objects, the page setup, the table and data-validation definitions you pass are the
objects you already build. What changed is **how you reach them**. If you are coming from `@cj-tech-master/excelts`
instead, read [MIGRATION_EXCELTS_TO_DOCUMONSTER.md](MIGRATION_EXCELTS_TO_DOCUMONSTER.md); if you are upgrading an
earlier `documonster`, read [MIGRATION.md](MIGRATION.md).

> **The five rules that cover most of the work**
>
> 1. **No methods on documents.** Every operation is a function in a namespace, and the document is its first
>    argument: `ws.getCell("A1").value = 42` → `Cell.setValue(ws, "A1", 42)`. Workbooks, worksheets and tables are
>    handles you pass around; rows, columns and cells are addressed by position and never handed out.
> 2. **A cell is one argument:** `"A1"` or `{ row, col }` (1-based) — never `(row, col)`. The one place loose numbers
>    remain is a **range** given by its four corners, `Worksheet.merge(ws, top, left, bottom, right)`, which mirrors
>    exceljs's `ws.mergeCells(t, l, b, r)`. Prefer the string form, `Worksheet.merge(ws, "A1:C3")`: it cannot be
>    misread, and it is what every other range-taking function accepts.
> 3. **A read returns a new workbook:** `await wb.xlsx.load(buf)` → `const wb = await Workbook.read(buf)`.
> 4. **Bytes are `Uint8Array`**, on Node too: `wb.xlsx.writeBuffer()` resolved to a `Buffer`, `Workbook.toBuffer`
>    resolves to a `Uint8Array`.
> 5. **Settings go in a trailing options object** (`Worksheet.protect(ws, { password, … })`), never in positional
>    arguments after the target.

## Contents

- [5-minute quick start](#5-minute-quick-start)
- [Code that compiles but behaves differently](#code-that-compiles-but-behaves-differently)
- [TypeScript types](#typescript-types)
- [In the browser](#in-the-browser)

1. [Install, imports and runtime](#1-install-imports-and-runtime)
2. [The model in one example](#2-the-model-in-one-example)
3. [Reading and writing files](#3-reading-and-writing-files)
4. [Workbook](#4-workbook)
5. [Worksheet](#5-worksheet)
6. [Rows](#6-rows)
7. [Columns](#7-columns)
8. [Cells](#8-cells)
9. [Cell values](#9-cell-values)
10. [Styles](#10-styles)
11. [Merged cells, names, notes, hyperlinks](#11-merged-cells-names-notes-hyperlinks)
12. [Data validation and conditional formatting](#12-data-validation-and-conditional-formatting)
13. [Tables](#13-tables)
14. [Images and backgrounds](#14-images-and-backgrounds)
15. [Page setup, views, freeze panes, auto filter, protection](#15-page-setup-views-freeze-panes-auto-filter-protection)
16. [Streaming](#16-streaming)
17. [CSV](#17-csv)
18. [Dates](#18-dates)
19. [Errors](#19-errors)
20. [What exceljs does not have, and what documonster does not have](#20-what-exceljs-does-not-have-and-what-documonster-does-not-have)
21. [Checklist](#21-checklist)
22. [Complete before/after example](#22-complete-beforeafter-example)

- [Appendix: exceljs internals](#appendix-exceljs-internals)

---

## 5-minute quick start

The everyday exceljs program, line for line. Every call on the right is a function that takes the workbook or the
worksheet first; there are no row, column or cell objects.

```ts before
// before (exceljs)
import ExcelJS from "exceljs";

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile("in.xlsx");
const ws = wb.getWorksheet("People")!;

// read and write cells
const name = ws.getCell("A2").value;
const age = ws.getCell(2, 2).value;
ws.getCell("D1").value = "Updated";

// style
ws.getRow(1).font = { bold: true };
ws.getCell("A1").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFFF00" }
};
ws.getCell("B2").numFmt = "#,##0";

// append rows
ws.addRow(["Linus", 28]);
ws.addRows([
  ["Ada", 36],
  ["Grace", 45]
]);

// every row and every cell
ws.eachRow((row, rowNumber) => {
  row.eachCell((cell, colNumber) => {
    console.log(rowNumber, colNumber, cell.address, cell.value);
  });
});

// write
await wb.xlsx.writeFile("out.xlsx");
```

```ts
// after (documonster)
import { Cell, Row, Workbook, Worksheet } from "documonster/excel";

const wb = await Workbook.readFile("in.xlsx"); // a read returns the workbook
const ws = Workbook.getWorksheet(wb, "People");
if (!ws) throw new Error("no People sheet");

// read and write cells
const name = Cell.getValue(ws, "A2");
const age = Cell.getValue(ws, { row: 2, col: 2 });
Cell.setValue(ws, "D1", "Updated");

// style
Row.setFont(ws, 1, { bold: true });
Cell.setFill(ws, "A1", {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFFF00" }
});
Cell.setNumFmt(ws, "B2", "#,##0");

// append rows
Row.add(ws, ["Linus", 28]); // returns the new row's number
Row.addMany(ws, [
  ["Ada", 36],
  ["Grace", 45]
]);

// every row and every cell: the callbacks get numbers and addresses, not objects
Row.each(ws, rowNumber => {
  Row.eachAddress(ws, rowNumber, (address, row, col) => {
    console.log(row, col, address, Cell.getValue(ws, address));
  });
});

// the sheet as records keyed by its header row
const people = Worksheet.toJson(ws); // [{ Name: "Ada", Age: 36 }, …]

// write
await Workbook.writeFile(wb, "out.xlsx");
```

`Worksheet.toJson` is what exceljs code usually built by hand from `ws.columns` keys and `row.values`. It keys each
record by the first row; `{ header: ["name", "age"] }` gives your own keys (then pass `range: 1` to skip the header
row), `{ header: 1 }` returns arrays, and `range: "A1:C20"` limits the read. Column **keys** (`Column.setDefinitions(ws,
[{ key, … }])`) are not stored in the file, as in exceljs: after a read, set them again before calling
`Row.add(ws, { key: value })` with an object.

In Node, `Workbook.readFile` and `Workbook.writeFile` take a path; everywhere, `Workbook.read(bytes)` and
`Workbook.toBuffer(wb)` work on bytes (§3, [In the browser](#in-the-browser)).

---

## Code that compiles but behaves differently

A port that type-checks can still rely on an exceljs behaviour that changed. Check these by hand:

- **A facet you read is a copy.** `Cell.getFont(ws, a)!.bold = true` changes nothing (likewise `getFill`, `getBorder`,
  `Row.getFont`, `Column.getFont`, …). Write with the setter: `Cell.setFont(ws, a, { ...Cell.getFont(ws, a), bold: true })`.
- **Values and notes are copied in and out.** `Cell.setValue(ws, a, obj)` / `Cell.setNote` store a copy, and what
  `Cell.getValue` / `Cell.getNote` return is yours — editing it later does not reach the cell. Set it again instead.
- **`Row.getValues` is 0-based by default.** exceljs's `row.values` had column A at index 1; write
  `Row.getValues(ws, n, { base: 1 })` (and `Column.getValues` / `Stream.getRowValues` likewise).
- **A read returns a new workbook.** `Workbook.read` / `Workbook.readFile` do not fill a workbook you created: use
  `const wb = await Workbook.read(bytes)`.
- **Formula results are what the file cached.** `Cell.getResult` is `undefined` for a formula nobody calculated; call
  `ExcelFormula.calculate(wb)` (`documonster/excel/formula`) first.
- **A committed streamed row is gone.** Reading or editing it (`Cell.getValue`, `Cell.setValue`, `Row.getValues`, …)
  on a `Stream.WorksheetWriter` throws `ExcelStreamStateError`; finish a row before `Stream.commitRow(row)`.
- **Bytes are `Uint8Array`, not `Buffer`.** `bytes.slice()` copies and `bytes.toString("base64")` does not exist; use
  `subarray` for a view, `Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)` for `Buffer` methods.
- **Absence is `undefined`, not `null`.** `Workbook.getWorksheet`, `Table.get`, `Cell.getNote`, `Row.getLast`, … return
  `undefined`: a `=== null` test is always false — write `=== undefined` or `== null`. (An empty cell's
  _value_ is still `null`.)
- **A six-digit colour reads back with alpha.** `{ argb: "f56600" }` is written as `FF` + rgb (the file format
  requires four bytes), so a written-then-read file gives `"FFf56600"`. Compare against the eight-digit form.

---

## TypeScript types

exceljs's interfaces were objects with methods. Here the document objects are opaque handles, and the rows, columns and
cells inside them have no type at all, because nothing hands them out. The data shapes — values, styles, page setup,
rules — keep their exceljs names.

| exceljs type                                                                                              | documonster type                                                                                                              |
| --------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| `ExcelJS.Workbook`                                                                                        | `Workbook.Handle`                                                                                                             |
| `ExcelJS.Worksheet`                                                                                       | `Worksheet.Handle`                                                                                                            |
| `ExcelJS.Row`                                                                                             | no counterpart: a row is its number (`number`)                                                                                |
| `ExcelJS.Column`                                                                                          | no counterpart: a column is its key, letter or number (`string \| number`); its settings are a `ColumnDefn`                   |
| `ExcelJS.Cell`                                                                                            | no counterpart: a cell is its address, `string \| CellPosition` (`{ row, col }`)                                              |
| `ExcelJS.Table`                                                                                           | `Table.Handle`; a definition is `TableProperties`, a column `TableColumnProperties`                                           |
| `ExcelJS.Image` / `ImageRange`                                                                            | `ImageData` (what `Image.add` takes) / `AddImageRange`                                                                        |
| `ExcelJS.CellValue`                                                                                       | `CellValue` (what a read gives) and `CellValueInput` (what a write accepts)                                                   |
| `CellFormulaValue`, `CellSharedFormulaValue`, `CellRichTextValue`, `CellHyperlinkValue`, `CellErrorValue` | the same names                                                                                                                |
| `RichText`                                                                                                | `RichText`                                                                                                                    |
| `Style`, `Font`, `Fill`, `FillPattern`, `FillGradientAngle`, `FillGradientPath`                           | the same names                                                                                                                |
| `Border`, `Borders`, `BorderStyle`, `Alignment`, `Protection`, `Color`                                    | the same names                                                                                                                |
| `PageSetup`, `HeaderFooter`, `Margins`, `WorksheetProperties`, `WorksheetView`, `AutoFilter`              | the same names                                                                                                                |
| `AddWorksheetOptions`, `WorksheetProtection`, `Location`                                                  | the same names                                                                                                                |
| `DataValidation` (the rule)                                                                               | `DataValidationRule`                                                                                                          |
| `ConditionalFormattingOptions`, `ConditionalFormattingRule`, `Cvfo`                                       | `ConditionalFormattingOptions`, `ConditionalFormattingRule`, `Cfvo`                                                           |
| `stream.xlsx.WorkbookWriter` / `WorkbookReader`                                                           | `Stream.WorkbookWriter` / `Stream.WorkbookReader` (classes); a streamed row is `Stream.RowHandle`, a cell `Stream.CellHandle` |
| `Buffer`                                                                                                  | `Uint8Array`                                                                                                                  |

A helper that took a worksheet keeps its shape; it takes the handle:

```ts
import { Cell } from "documonster/excel";
import type { Worksheet } from "documonster/excel";

function highlight(ws: Worksheet.Handle, address: string) {
  Cell.setFont(ws, address, { bold: true, color: { argb: "FFC00000" } });
}
```

---

## In the browser

The browser build of `documonster/excel` has everything except what needs a file system: there is no
`Workbook.readFile` or `Workbook.writeFile` (a call is a compile error, not a runtime one), and an image cannot be added
by `filename`. Bundlers (Vite, webpack, Rollup, esbuild) choose that build on their own through the `browser`
condition. For TypeScript to see the browser surface too, add the condition to your `tsconfig.json`:

```json
{ "compilerOptions": { "customConditions": ["browser"] } }
```

Reading a file the user picked and offering the result as a download:

```ts browser
import { Cell, Workbook } from "documonster/excel";

const input = document.querySelector<HTMLInputElement>("input[type=file]")!;
input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;
  // A File is a Blob, which Workbook.read accepts; `await file.arrayBuffer()` works as well.
  const wb = await Workbook.read(file);
  const ws = Workbook.listWorksheets(wb)[0]!;
  Cell.setValue(ws, "A1", "Edited in the browser");

  const bytes = await Workbook.toBuffer(wb);
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "edited.xlsx";
  link.click();
  URL.revokeObjectURL(link.href);
});
```

This replaces exceljs's `wb.xlsx.load(await file.arrayBuffer())` and `new Blob([await wb.xlsx.writeBuffer()])`. Without
a bundler, the place of `exceljs/dist/exceljs.min.js` is taken by a `<script>` tag for
`dist/iife/documonster.excel.iife.min.js`, which defines `Documonster.Excel` (`const { Workbook, Cell } =
Documonster.Excel`); the [README's browser section](README.md#browser-support) gives the pinned URL. That bundle does
not include the formula engine (`ExcelFormula`).

---

## 1. Install, imports and runtime

```bash
npm remove exceljs
npm install documonster@^0.22.0
# or: pnpm add documonster@^0.22.0
# or: yarn add documonster@^0.22.0
```

| Aspect        | exceljs                                        | documonster                                                                      |
| ------------- | ---------------------------------------------- | -------------------------------------------------------------------------------- |
| Package       | `exceljs`                                      | `documonster`                                                                    |
| Import        | `import ExcelJS from "exceljs"` (one default)  | named imports from `documonster/excel` (one per namespace)                       |
| CSV           | `wb.csv.*`, bundled                            | `documonster/excel/csv` (`ExcelCsv.*`), loaded only when you import it           |
| Module format | CommonJS (plus a browser bundle)               | ESM-only; CommonJS works through `require(esm)`                                  |
| Node          | ≥ 8                                            | ≥ 22.13 (also Bun and modern browsers)                                           |
| Browser       | `exceljs/dist/exceljs.min.js`                  | the same import; bundlers pick the browser build through the `browser` condition |
| Runtime deps  | `archiver`, `unzipper`, `fast-csv`, `jszip`, … | none                                                                             |

```ts before
import ExcelJS from "exceljs";
const wb = new ExcelJS.Workbook();
```

```ts
import { Workbook } from "documonster/excel";
const wb = Workbook.create();
```

CommonJS keeps its shape: `const { Workbook, Cell } = require("documonster/excel")` works on Node ≥ 22.13. Two setups
need a one-line change: TypeScript with `moduleResolution: "node16"` (use `"nodenext"` or `"bundler"`), and Jest, which
cannot `require()` an ES module (transpile it with `babel-jest` and
`"transformIgnorePatterns": ["/node_modules/(?!documonster)"]`).

The namespaces `documonster/excel` publishes: `Workbook`, `Worksheet`, `Cell`, `Row`, `Column`, `Range`, `Chart`,
`Table`, `Image`, `Pivot`, `Sparkline`, `Form`, `Chartsheet`, `ConditionalFormatting`, `DataValidation`,
`DefinedNames`, `Note`, `Address`, `Anchor`, `Watermark`, `HeaderFooterImage`, `Format`, `Stream`. The constants
(`ValueType`, `ErrorValue`, `PaperSize`, …), the style and value types (`Style`, `Font`, `Fill`, `CellValue`, …) and the
error classes are flat named exports beside them.

`ExcelJS.config.setValue("promise", …)` has no counterpart: every asynchronous call returns a native `Promise`.

---

## 2. The model in one example

```ts before
// before (exceljs)
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet("Sales");
ws.columns = [
  { header: "Region", key: "region", width: 15 },
  { header: "Total", key: "total", width: 12 }
];
ws.addRow({ region: "North", total: 120 });
ws.getCell("B2").numFmt = "#,##0";
ws.getRow(1).font = { bold: true };
await wb.xlsx.writeFile("sales.xlsx");
```

```ts
// after (documonster)
import { Cell, Column, Row, Workbook } from "documonster/excel";

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Sales");
Column.setDefinitions(ws, [
  { header: "Region", key: "region", width: 15 },
  { header: "Total", key: "total", width: 12 }
]);
Row.add(ws, { region: "North", total: 120 });
Cell.setNumFmt(ws, "B2", "#,##0");
Row.setFont(ws, 1, { bold: true });
await Workbook.writeFile(wb, "sales.xlsx");
```

The pattern is always the same:

| exceljs             | documonster                                                     |
| ------------------- | --------------------------------------------------------------- |
| `obj.method(args)`  | `Namespace.method(handle, args)`                                |
| `obj.prop` (read)   | `Namespace.getProp(handle)` (`isProp` / `hasProp` for booleans) |
| `obj.prop = v`      | `Namespace.setProp(handle, v)`                                  |
| `ws.getRow(n).x`    | `Row.*(ws, n)` — the row is its number                          |
| `ws.getColumn(c).x` | `Column.*(ws, c)` — the column is its key, letter or number     |
| `ws.getCell(a).x`   | `Cell.*(ws, a)` — the cell is `"A1"` or `{ row, col }`          |

Two kinds of objects still have plain fields you read and assign directly, as in exceljs: the **workbook's document
properties** (`wb.creator`, `wb.title`, `wb.properties.date1904`, `wb.calcProperties`, `wb.views`, …) and the
**worksheet's print and view settings** (`ws.pageSetup`, `ws.headerFooter`, `ws.views`, `ws.properties`, `ws.state`).
Everything else goes through a function.

---

## 3. Reading and writing files

| exceljs                                                                                                          | documonster                                                                                                                                          |
| ---------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- |
| `await wb.xlsx.readFile(path)`                                                                                   | `const wb = await Workbook.readFile(path)` _(Node)_                                                                                                  |
| `await wb.xlsx.load(buffer)`                                                                                     | `const wb = await Workbook.read(buffer)`                                                                                                             |
| `await wb.xlsx.read(readable)`                                                                                   | `const wb = await Workbook.readStream(readable)`                                                                                                     |
| `await wb.xlsx.writeFile(path)`                                                                                  | `await Workbook.writeFile(wb, path)` _(Node)_                                                                                                        |
| `await wb.xlsx.writeBuffer()`                                                                                    | `await Workbook.toBuffer(wb)` → `Uint8Array`                                                                                                         |
| `await wb.xlsx.write(writable)`                                                                                  | `await Workbook.writeStream(wb, writable)`                                                                                                           |
| _(none)_                                                                                                         | `Workbook.toStream(wb)` — a readable you pipe somewhere                                                                                              |
| `wb.xlsx.load(data, { base64: true })`                                                                           | `Workbook.read(base64String, { base64: true })`                                                                                                      |
| `wb.xlsx.load(data, { ignoreNodes })`                                                                            | `Workbook.read(data, { ignoreNodes })`                                                                                                               |
| `wb.xlsx` (the `Xlsx` object)                                                                                    | the functions above, which take the workbook                                                                                                         |
| `XlsxWriteOptions` (`zip`, `useSharedStrings`, `useStyles`)                                                      | `XlsxWriteOptions`, the same three; the `filename` and `stream` it inherited from the streaming writer's options were never read by `wb.xlsx.write*` |
| `writeBuffer({ zip: { compression: "STORE" } })` / `{ compressionOptions: { level } }` (`JSZipGeneratorOptions`) | `Workbook.toBuffer(wb, { zip: { level: 0 } })` / `{ zip: { level } }`                                                                                |

**A read returns a new workbook.** exceljs filled the `Workbook` you constructed; there is no object to fill here,
so drop the `new ExcelJS.Workbook()` that preceded a read:

```ts before
const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile("in.xlsx");
```

```ts
import { Workbook } from "documonster/excel";

const wb = await Workbook.readFile("in.xlsx");
```

`Workbook.read` accepts a `Uint8Array`, an `ArrayBuffer`, any `ArrayBufferView`, a `Blob`, or a base64 string with
`{ base64: true }`. `Workbook.readStream` accepts a Node readable, a Web `ReadableStream` or an async iterable of
bytes.

**`toBuffer` gives a `Uint8Array`.** `fs.writeFile`, `res.end`, `Readable.from` and `new Blob([…])` take it unchanged. If
some code needs a `Buffer` method, wrap without copying: `Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)`.
Note that `bytes.slice()` copies, as it does for every `Uint8Array`; `subarray` is the view.

**`writeStream` honours backpressure**, which exceljs did not: attach the consumer before awaiting it.

```ts
import { PassThrough } from "node:stream";

import { Workbook } from "documonster/excel";

declare function upload(body: PassThrough): Promise<void>; // your consumer

const pass = new PassThrough();
const uploading = upload(pass); // starts reading
await Workbook.writeStream(wb, pass);
await uploading;
```

Awaiting `writeStream` on a stream nobody is reading never resolves once its buffer fills — about 128 KB of output.
A file stream or an HTTP response reads on its own and needs no care.

**New options** with no exceljs counterpart: `password` (read and write encrypted workbooks), `format: "xlsb"` (read and
write the binary format), `maxRows` / `maxCols` / `zip` limits for untrusted input, `formulas` and `blankCells` on read.

---

## 4. Workbook

| exceljs                                                                                                                                                                                | documonster                                                                                         |
| -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `new ExcelJS.Workbook()`                                                                                                                                                               | `Workbook.create()`                                                                                 |
| `wb.addWorksheet(name, options)`                                                                                                                                                       | `Workbook.addWorksheet(wb, name, options)` — same `options` shape                                   |
| `wb.getWorksheet(name)` / `wb.getWorksheet(id)`                                                                                                                                        | `Workbook.getWorksheet(wb, name)` / `Workbook.getWorksheet(wb, id)` → `undefined` when absent       |
| `wb.worksheets`                                                                                                                                                                        | `Workbook.listWorksheets(wb)`                                                                       |
| `wb.eachSheet((ws, id) => …)`                                                                                                                                                          | `Workbook.eachWorksheet(wb, (ws, id) => …)`                                                         |
| `wb.removeWorksheet(id)` / `wb.removeWorksheetEx(ws)`                                                                                                                                  | `Workbook.removeWorksheet(wb, id \| name \| ws)` → `boolean`                                        |
| `wb.getImage(id)` / `wb.media`                                                                                                                                                         | `wb.media[Number(id)]` — the media list is still on the handle (`WorkbookMedia`, exceljs's `Media`) |
| `wb.addImage({ … })`                                                                                                                                                                   | `Image.add(wb, { … })` → image id (§14)                                                             |
| `wb.definedNames`                                                                                                                                                                      | `Workbook.getDefinedNames(wb)` (§11)                                                                |
| `wb.model` / `wb.model = m`                                                                                                                                                            | `Workbook.getModel(wb)` / `Workbook.setModel(wb, m)`                                                |
| `wb.creator`, `wb.lastModifiedBy`, `wb.created`, `wb.modified`, `wb.lastPrinted`, `wb.title`, `wb.subject`, `wb.keywords`, `wb.category`, `wb.description`, `wb.company`, `wb.manager` | unchanged — plain fields on the handle                                                              |
| `wb.properties.date1904 = true`                                                                                                                                                        | unchanged                                                                                           |
| `wb.calcProperties.fullCalcOnLoad = true`                                                                                                                                              | unchanged                                                                                           |
| `wb.views = [{ … }]`                                                                                                                                                                   | unchanged                                                                                           |
| `wb.xlsx.*` / `wb.csv.*`                                                                                                                                                               | `Workbook.*` (§3) / `ExcelCsv.*` (§17)                                                              |

Things exceljs did not have that you may want while you are here: `Workbook.protect(wb, { password, lockStructure })`,
`Workbook.addCellStyle` / `Cell.applyStyle` for named cell styles, `Workbook.importWorksheet` to copy a sheet from
another workbook, and `ExcelFormula.calculate(wb)` (`documonster/excel/formula`) to compute every formula's result.

---

## 5. Worksheet

| exceljs                                                                             | documonster                                                                                                      |
| ----------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| `ws.name` / `ws.name = n`                                                           | `Worksheet.getName(ws)` / `Worksheet.setName(ws, n)`                                                             |
| `ws.id`, `ws.state`, `ws.properties`, `ws.pageSetup`, `ws.headerFooter`, `ws.views` | unchanged — plain fields (§15)                                                                                   |
| `ws.getCell(a)` / `ws.getCell(r, c)`                                                | `Cell.*(ws, "A1")` / `Cell.*(ws, { row: r, col: c })` (§8) — a letter or key `c`: `Column.getNumber(ws, c)`      |
| `ws.getRow(n)`                                                                      | `Row.*(ws, n)` (§6)                                                                                              |
| `ws.getColumn(c)`                                                                   | `Column.*(ws, c)` (§7)                                                                                           |
| `ws.columns = defs`                                                                 | `Column.setDefinitions(ws, defs)` — same `{ header, key, width, … }`                                             |
| `ws.columns` (read)                                                                 | `Column.getDefinitions(ws)` — plain data                                                                         |
| `ws.addRow(values)`                                                                 | `Row.add(ws, values)` → **the row number**                                                                       |
| `ws.addRow(values, "i")` / `"i+"`                                                   | `Row.add(ws, values, { copyStyle: "above" })` / `{ copyStyle: "above", includeEmptyCells: true }`                |
| `ws.addRow(values, "o")` / `"n"`                                                    | `Row.add(ws, values)` — a new row has no style to keep                                                           |
| `ws.addRows(rows)`                                                                  | `Row.addMany(ws, rows)` → row numbers                                                                            |
| `ws.insertRow(pos, values, "i")` / `"o"`                                            | `Row.insert(ws, pos, values, { copyStyle: "above" })` / `{ copyStyle: "below" }` (the row it pushed down)        |
| `ws.insertRows(pos, rows)`                                                          | `Row.insertMany(ws, pos, rows)`                                                                                  |
| `ws.duplicateRow(n, count, insert)`                                                 | `Row.duplicate(ws, n, count, { insert })`                                                                        |
| `ws.spliceRows(start, count, ...rows)`                                              | `Row.splice(ws, start, count, ...rows)`                                                                          |
| `ws.spliceColumns(start, count, ...cols)`                                           | `Column.splice(ws, start, count, ...cols)`                                                                       |
| `ws.getRows(start, length)`                                                         | loop `start … start + length - 1` with `Row.*`                                                                   |
| `ws.findRow(n)` / `ws.findRows(start, length)`                                      | `Row.has(ws, n)` per row (and `Cell.getView` for a read that creates nothing)                                    |
| `ws.eachRow((row, n) => …)`                                                         | `Row.each(ws, n => …)` — the callback gets the **row number**                                                    |
| `ws.eachRow({ includeEmpty: true }, cb)`                                            | `Row.each(ws, cb, { includeEmpty: true })`                                                                       |
| `ws.rowCount` / `ws.actualRowCount`                                                 | `Row.getCount(ws)` / `Row.getActualCount(ws)`                                                                    |
| `ws.columnCount` / `ws.actualColumnCount`                                           | `Column.getCount(ws)` / `Column.getActualCount(ws)`                                                              |
| `ws.lastRow` / `ws.lastColumn`                                                      | `Row.getLast(ws)` / `Column.getLast(ws)` — numbers, `undefined` when empty                                       |
| `ws.dimensions`                                                                     | `Worksheet.getDimensions(ws)`                                                                                    |
| `ws.getSheetValues()`                                                               | `Worksheet.getValues(ws, { base: 1 })` — 1-based, as exceljs returns it                                          |
| `ws.mergeCells("A1:B2")` / `(t, l, b, r)`                                           | `Worksheet.merge(ws, "A1:B2")` / `Worksheet.merge(ws, t, l, b, r)`                                               |
| `ws.mergeCellsWithoutStyle(range)`                                                  | `Worksheet.mergeWithoutStyle(ws, range)`                                                                         |
| `ws.getColumnKey(k)` / `ws.eachColumnKey(cb)`                                       | `Column.getNumber(ws, k)` / `Column.getDefinitions(ws)` and each definition's `key`                              |
| `ws.setColumnKey(k, col)`                                                           | `Column.setKey(ws, c, k)`                                                                                        |
| `ws.deleteColumnKey(k)`                                                             | no direct counterpart: give the column another key with `Column.setKey`                                          |
| `ws.unMergeCells("A1:B2")`                                                          | `Worksheet.unmerge(ws, "A1:B2")`                                                                                 |
| `ws.fillFormula(range, f, results, shareType)`                                      | `Worksheet.fillFormula(ws, range, f, { results, shareType })`                                                    |
| `ws.protect(pw, options)` / `ws.unprotect()`                                        | `await Worksheet.protect(ws, { password: pw, ...options })` / `Worksheet.unprotect(ws)`                          |
| `ws.autoFilter = "A1:C1"`                                                           | `Worksheet.setAutoFilter(ws, "A1:C1")` / `Worksheet.getAutoFilter(ws)`                                           |
| `ws.addImage(id, range)`                                                            | `Image.place(ws, id, range)` (§14)                                                                               |
| `ws.addBackgroundImage(id)`                                                         | `Image.setBackground(ws, id)`                                                                                    |
| `ws.getBackgroundImageId()`                                                         | `Image.getBackground(ws)` — the id as a string (`Number(id)` is the id `Image.add` returned)                     |
| `ws.getImages()`                                                                    | `Image.list(ws)`                                                                                                 |
| `ws.addTable(def)` / `ws.getTable(n)` / `ws.getTables()` / `ws.removeTable(n)`      | `Table.add(ws, def)` / `Table.get(ws, n)` / `Table.list(ws)` / `Table.remove(ws, n)` (§13)                       |
| `ws.addConditionalFormatting(cf)`                                                   | `ConditionalFormatting.add(ws, cf)` (§12)                                                                        |
| `ws.removeConditionalFormatting(filter)`                                            | `ConditionalFormatting.remove(ws, ref)` / `ConditionalFormatting.clear(ws)`                                      |
| `ws.dataValidations.add(ref, rule)` (untyped in exceljs)                            | `DataValidation.add(ws, ref, rule)` (§12)                                                                        |
| `wb.removeWorksheet(ws.id)` / `ws.destroy()`                                        | `Workbook.removeWorksheet(wb, ws)`                                                                               |
| `ws.model` / `ws.model = m`                                                         | `Worksheet.getModel(ws)` / `Worksheet.setModel(ws, m)`                                                           |
| `ws.model.merges`                                                                   | `Worksheet.getMergedRegions(ws)`; a `WorksheetModel` holds them as `mergeCells`                                  |
| `ws.workbook`, `cell.workbook`, `cell.worksheet`, `col.worksheet`                   | no back-reference: keep the workbook and worksheet handles you have; `cell.sheetName` is `Worksheet.getName(ws)` |
| `ws.commit()` (streaming)                                                           | `ws.commit()` — streaming writers keep it (§16)                                                                  |

**The row and column functions never create what they read.** `ws.getRow(n).height` in exceljs created the row;
`Row.getHeight(ws, n)` does not, and neither do the facet getters (`Cell.getFont`, `Row.getFill`, …). Only
`getStyle` hands out the live style record, and creates its cell, row or column to do so.

New with no exceljs counterpart: `Worksheet.freeze` / `Worksheet.split` / `Worksheet.unfreeze` for panes (§15),
`Row.autoFit` / `Column.autoFit`, `Worksheet.toJson` / `Worksheet.addJson`, `Worksheet.toAoa` / `Worksheet.addAoa`, and
`Range.*` to read, write or style a block of cells at once.

---

## 6. Rows

Every function takes `(ws, rowNumber)`.

| exceljs (`row = ws.getRow(n)`)                                             | documonster                                                                                    |
| -------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- |
| `row.number`                                                               | `n` — a row is its number                                                                      |
| `row.values` (read)                                                        | `Row.getValues(ws, n, { base: 1 })` — **1-based**, index 0 empty, as exceljs                   |
| `row.values = [, "a", "b"]` (index 0 empty: column-numbered)               | `Row.setValues(ws, n, [, "a", "b"], { base: 1 })`                                              |
| `row.values = ["a", "b"]` (index 0 set: exceljs reads it from column A)    | `Row.setValues(ws, n, ["a", "b"], { base: 0 })`                                                |
| `row.values = { id: 1, name: "x" }`                                        | `Row.setValues(ws, n, { id: 1, name: "x" })` (by column key)                                   |
| `row.getCell(c)` — `c` a column number                                     | `Cell.*(ws, { row: n, col: c })`                                                               |
| `row.getCell(c)` — `c` a letter (`"B"`) or a column key                    | `Cell.*(ws, { row: n, col: Column.getNumber(ws, c) })` — `col` is always a number              |
| `row.findCell(c)`                                                          | `Cell.getView(ws, { row: n, col: c })` — `undefined` when absent; creates nothing              |
| `row.eachCell((cell, col) => …)`                                           | `Row.eachAddress(ws, n, (address, row, col) => …)`                                             |
| `row.eachCell({ includeEmpty: true }, cb)`                                 | `Row.eachAddress(ws, n, cb, { includeEmpty: true })`                                           |
| `row.height` / `= h`                                                       | `Row.getHeight(ws, n)` / `Row.setHeight(ws, n, h)`                                             |
| `row.hidden` / `= b`                                                       | `Row.isHidden(ws, n)` / `Row.setHidden(ws, n, b)`                                              |
| `row.outlineLevel` / `= l`                                                 | `Row.getOutlineLevel(ws, n)` / `Row.setOutlineLevel(ws, n, l)`                                 |
| `row.collapsed`                                                            | `Row.isCollapsed(ws, n)`                                                                       |
| `row.font = f` (and `fill`, `border`, `alignment`, `numFmt`, `protection`) | `Row.setFont(ws, n, f)` (`setFill`, `setBorder`, `setAlignment`, `setNumFmt`, `setProtection`) |
| `row.font` (read)                                                          | `Row.getFont(ws, n)` — a copy                                                                  |
| `row.style`                                                                | `Row.getStyle(ws, n)` (live) / `Row.setStyle(ws, n, style)`                                    |
| `row.hasValues`                                                            | `Row.hasValues(ws, n)`                                                                         |
| `row.cellCount` / `row.actualCellCount`                                    | `Row.getCellCount(ws, n)` / `Row.getActualCellCount(ws, n)`                                    |
| `row.dimensions`                                                           | `Row.getDimensions(ws, n)` → `undefined` for an empty row                                      |
| `row.model`                                                                | `Row.getModel(ws, n)`                                                                          |
| `row.addPageBreak()`                                                       | `Row.addPageBreak(ws, n)`                                                                      |
| `row.splice(start, count, ...values)`                                      | `Row.spliceCells(ws, n, start, count, ...values)` — moves values, styles and notes, as exceljs |
| `row.commit()`                                                             | delete it in the document model; streamed: `Stream.commitRow(row)` (§16)                       |

**Watch the index base.** exceljs's `row.values` is **1-based**: index 0 is empty and column A is index 1.
`Row.getValues(ws, n)` without options returns a **0-based** dense array (column A at index 0). When you port
`row.values`, write `{ base: 1 }` — or shift your indices by one. `Row.setValues` without `base` guesses the way
exceljs did (an array whose index 0 is empty is column-numbered; otherwise it starts at column A), so a column-numbered
array whose column A is genuinely empty is misread by both; pass `base` and nothing is guessed. Note also that
`Row.setValues` **replaces** the row — cells the values do not carry are cleared — as assigning `row.values` did.

A function that took a row now takes its number:

```ts before
function stripe(row: ExcelJS.Row) {
  row.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF2F2F2" }
  };
}
ws.eachRow((row, n) => n % 2 === 0 && stripe(row));
```

```ts
import { Row } from "documonster/excel";
import type { Worksheet } from "documonster/excel";

function stripe(ws: Worksheet.Handle, row: number) {
  Row.setFill(ws, row, {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFF2F2F2" }
  });
}
Row.each(ws, n => {
  if (n % 2 === 0) stripe(ws, n);
});
```

---

## 7. Columns

Every function takes `(ws, col)` where `col` is the column's key, its letter (`"B"`) or its 1-based number.

| exceljs (`col = ws.getColumn(c)`)                                                                    | documonster                                                                                                                                                                                               |
| ---------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `col.number` / `col.letter`                                                                          | `Column.getNumber(ws, c)` / `Column.getLetter(ws, c)`                                                                                                                                                     |
| `col.header` / `= h`                                                                                 | `Column.getHeader(ws, c)` / `Column.setHeader(ws, c, h)`                                                                                                                                                  |
| `col.key` / `= k`                                                                                    | `Column.getKey(ws, c)` / `Column.setKey(ws, c, k)`                                                                                                                                                        |
| `col.width` / `= w`                                                                                  | `Column.getWidth(ws, c)` / `Column.setWidth(ws, c, w)`                                                                                                                                                    |
| `col.hidden` / `= b`                                                                                 | `Column.isHidden(ws, c)` / `Column.setHidden(ws, c, b)`                                                                                                                                                   |
| `col.outlineLevel` / `= l`                                                                           | `Column.getOutlineLevel(ws, c)` / `Column.setOutlineLevel(ws, c, l)`                                                                                                                                      |
| `col.values` (read)                                                                                  | `Column.getValues(ws, c, { base: 1 })` — indexed by row number, as exceljs                                                                                                                                |
| `col.values = [ , "a", "b"]`                                                                         | `Column.setValues(ws, c, [, "a", "b"], { base: 1 })`                                                                                                                                                      |
| `col.eachCell((cell, row) => …)`                                                                     | `Column.eachAddress(ws, c, (address, row, col) => …)`                                                                                                                                                     |
| `col.font`, `col.fill`, `col.border`, `col.alignment`, `col.numFmt`, `col.protection` (read / `= v`) | `Column.getFont(ws, c)` / `Column.setFont(ws, c, v)`, and `getFill` / `setFill`, `getBorder` / `setBorder`, `getAlignment` / `setAlignment`, `getNumFmt` / `setNumFmt`, `getProtection` / `setProtection` |
| `col.style`                                                                                          | `Column.getStyle(ws, c)` (live) / `Column.setStyle(ws, c, style)`                                                                                                                                         |
| `col.defn`                                                                                           | `Column.getDefinition(ws, c)`                                                                                                                                                                             |
| `col.toString()`                                                                                     | no counterpart: it printed the column for debugging; `Column.getDefinition(ws, c)` reads it                                                                                                               |
| `col.headers` / `col.headerCount`                                                                    | `Column.getHeader(ws, c)` — a string or an array; its length is the header row count                                                                                                                      |
| `col.isCustomWidth`                                                                                  | `Column.getWidth(ws, c) !== undefined` (exceljs also treated the default width 9 as not custom)                                                                                                           |
| `ws.getColumn(c).collapsed`                                                                          | `Column.getDefinition(ws, c)` and the outline levels                                                                                                                                                      |

`Column.setStyle` / `Column.setNumFmt` write the column **and the cells already in it**, which is what you usually
reached for `col.eachCell` to do. `col.eachCell` created a cell in every row it visited; `Column.eachAddress` creates
nothing.

---

## 8. Cells

Every function takes `(ws, addr)` where `addr` is `"A1"` or `{ row, col }`. There is no cell object to hold.

| exceljs (`cell = ws.getCell(a)`)                                                  | documonster                                                                                                               |
| --------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `cell.value` / `= v`                                                              | `Cell.getValue(ws, a)` / `Cell.setValue(ws, a, v)` (§9)                                                                   |
| `cell.text`                                                                       | `Cell.getText(ws, a)` — the same text exceljs gave (a `Date` as `Date#toString`)                                          |
| _(the text Excel shows)_                                                          | `Cell.getDisplayText(ws, a)` — formatted with the number format                                                           |
| `cell.type`                                                                       | `Cell.getType(ws, a)` — compare with `ValueType.*`, as before                                                             |
| `cell.effectiveType`                                                              | `Cell.getEffectiveType(ws, a)`                                                                                            |
| `cell.formula` / `cell.result`                                                    | `Cell.getFormula(ws, a)` / `Cell.getResult(ws, a)`                                                                        |
| `cell.result = r`                                                                 | `Cell.setResult(ws, a, r)`                                                                                                |
| `cell.numFmt` / `= f`                                                             | `Cell.getNumFmt(ws, a)` / `Cell.setNumFmt(ws, a, f)`                                                                      |
| `cell.font` / `= f`                                                               | `Cell.getFont(ws, a)` (a copy) / `Cell.setFont(ws, a, f)`                                                                 |
| `cell.fill`, `cell.border`, `cell.alignment`, `cell.protection`                   | `Cell.getFill` / `setFill`, `getBorder` / `setBorder`, `getAlignment` / `setAlignment`, `getProtection` / `setProtection` |
| `cell.style` / `cell.style = s`                                                   | `Cell.getStyle(ws, a)` (live) / `Cell.setStyle(ws, a, s)`                                                                 |
| `cell.dataValidation = rule`                                                      | `Cell.setValidation(ws, a, rule)` / `Cell.getValidation(ws, a)`                                                           |
| `cell.note = "…"` / `{ texts, … }`                                                | `Cell.setNote(ws, a, "…")` / `Cell.setNote(ws, a, { texts, … })`                                                          |
| `cell.note` (read)                                                                | `Cell.getNote(ws, a)`; author `Cell.getNoteAuthor(ws, a)`                                                                 |
| `cell.hyperlink` / `cell.isHyperlink`                                             | `Cell.getHyperlink(ws, a)` / `Cell.getHyperlink(ws, a) !== undefined`                                                     |
| `cell.name = n`                                                                   | `Cell.setNames(ws, a, [n])`                                                                                               |
| `cell.names` / `cell.addName(n)` / `cell.removeName(n)` / `cell.removeAllNames()` | `Cell.getNames` / `Cell.addName` / `Cell.removeName` / `Cell.clearNames`                                                  |
| `cell.isMerged`                                                                   | `Cell.isMerged(ws, a)`                                                                                                    |
| `cell.master`                                                                     | `Cell.getMergeMasterAddress(ws, a)` — its address                                                                         |
| `cell.isMergedTo(master)`                                                         | `Cell.getMergeMasterAddress(ws, a) === masterAddress`                                                                     |
| `cell.address`, `cell.row`, `cell.col`                                            | you already have them: `a`, or `Address.decodeCell(a)`                                                                    |
| `cell.$col$row` (`$A$1`)                                                          | `"$" + Address.encodeCol(col - 1) + "$" + row`, with `{ row, col }` from `Address.decodeCell(a)`                          |
| `cell.fullAddress`                                                                | `Cell.getFullAddress(ws, a)`                                                                                              |
| `cell.model` / `= m`                                                              | `Cell.getModel(ws, a)` / `Cell.setModel(ws, a, m)`                                                                        |
| `cell.merge(master)` / `cell.unmerge()`                                           | `Worksheet.merge(ws, range)` / `Worksheet.unmerge(ws, range)`                                                             |
| `cell.toCsvString()`                                                              | no counterpart; `ExcelCsv.format` writes a sheet                                                                          |
| `ws.findCell(a)`                                                                  | `Cell.getView(ws, a)` — a read-only view, `undefined` when absent; creates nothing                                        |

**Many cells at once:** `Range.setValues(ws, "A2:C3", [[…], […]])`, `Range.getValues(ws, "A2:C3")`,
`Range.setStyle(ws, "A1:C1", { font: { bold: true } })` (and `setFont`, `setFill`, `setBorder`, `setAlignment`,
`setNumFmt`, `setProtection`). A loop of `Cell.*` calls works, but a range is one call and one argument check.

**Values you pass are copied.** `Cell.setValue(ws, a, obj)` stores a copy of `obj`; changing `obj` afterwards does not
change the cell, and the object `Cell.getValue` returns is yours to change. In exceljs both were the cell's own.

---

## 9. Cell values

The value shapes are exceljs's, unchanged:

| Kind                    | Value                                                          |
| ----------------------- | -------------------------------------------------------------- |
| empty                   | `null`                                                         |
| number, string, boolean | as is                                                          |
| date                    | `Date` (§18)                                                   |
| hyperlink               | `{ text, hyperlink, tooltip? }`                                |
| formula                 | `{ formula: "A1+B1", result? }` — no leading `=`               |
| shared formula          | `{ sharedFormula: "A1", result? }`                             |
| array formula           | `{ formula, shareType: "array", ref: "A1:B2", result? }`       |
| rich text               | `{ richText: [{ text, font? }, …] }`                           |
| error                   | `{ error: "#N/A" }` — or `{ error: ErrorValue.NotApplicable }` |

`ValueType` keeps its member names (`ValueType.Number`, `ValueType.Formula`, …) and adds `Checkbox`. It is a plain
constant object rather than a TypeScript `enum`: `ValueType[2]` no longer gives `"Number"`, and an arbitrary `number`
is not assignable to `ValueType` without a cast. `FormulaType` is still exported, but no public reader returns it: test `Cell.getFormula` or the value's
shape (`"sharedFormula" in value`, `value.shareType === "array"`). `ErrorValue` names every Excel error, `#SPILL!` and
`#CALC!` included.

A value Excel cannot store is refused when you set it — an `ExcelInvalidValueTypeError` or `ExcelArgumentError`, at
the call — where exceljs wrote it and produced a file Excel repairs.

The assignments port one for one:

```ts before
ws.getCell("B5").value = { formula: "SUM(B2:B4)", result: 109 };
ws.getCell("A6").value = {
  richText: [{ text: "Bold ", font: { bold: true } }, { text: "plain" }]
};
ws.getCell("A7").value = {
  text: "Docs",
  hyperlink: "https://example.com",
  tooltip: "Open"
};
```

```ts
import { Cell } from "documonster/excel";

Cell.setValue(ws, "B5", { formula: "SUM(B2:B4)", result: 109 });
Cell.setValue(ws, "A6", {
  richText: [{ text: "Bold ", font: { bold: true } }, { text: "plain" }]
});
Cell.setValue(ws, "A7", {
  text: "Docs",
  hyperlink: "https://example.com",
  tooltip: "Open"
});
```

**Formula results.** As in exceljs, a formula you write has no result unless you give one, and `Cell.getResult` reads
only the result cached in the file. Excel recalculates on open, but a program that reads the file back — or a file
written by a tool that caches nothing — sees `undefined`. exceljs could not compute it; here the formula engine can:

```ts
import { Cell } from "documonster/excel";
import { ExcelFormula } from "documonster/excel/formula";

Cell.setValue(ws, "B6", { formula: "B5*2" });
ExcelFormula.calculate(wb); // fills every formula's result in the workbook
console.log(Cell.getResult(ws, "B6")); // twice B5's computed value, not undefined
```

`ExcelFormula` is a separate entry so that a bundle which never calculates does not carry the engine.

---

## 10. Styles

The style objects are exceljs's: `font` (`name`, `size`, `bold`, `color: { argb }`, …), `fill` (`{ type: "pattern",
pattern: "solid", fgColor }` or a gradient), `border` (`{ top: { style, color }, … }`), `alignment` (`horizontal`,
`vertical`, `wrapText`, `indent`, `textRotation`, …), `numFmt` (a format string) and `protection` (`locked`, `hidden`).
The types are flat exports: `Style`, `Font`, `Fill`, `Borders`, `Alignment`, `Protection`, `NumFmt`.

| exceljs                                           | documonster                                                    |
| ------------------------------------------------- | -------------------------------------------------------------- |
| `ws.getCell(a).font = f`                          | `Cell.setFont(ws, a, f)`                                       |
| `ws.getRow(n).font = f`                           | `Row.setFont(ws, n, f)` — the row and its existing cells       |
| `ws.getColumn(c).font = f`                        | `Column.setFont(ws, c, f)` — the column and its existing cells |
| `ws.getCell(a).style = s`                         | `Cell.setStyle(ws, a, s)` — merged into the cell's style       |
| `cell.font.bold = true` (editing the read object) | `Cell.setFont(ws, a, { ...Cell.getFont(ws, a), bold: true })`  |
| _(a block of cells)_                              | `Range.setStyle(ws, "A1:D1", s)` / `Range.setFont(…)`, …       |

**A facet you read is a copy.** In exceljs, `cell.font.bold = true` sometimes changed the cell and sometimes did not
(it did not for a cell that shared its style object). Here `Cell.getFont` always returns a copy; change the cell with a
setter, or edit the live record `Cell.getStyle(ws, a)` returns.

Named cell styles ("Good", "Heading 1", your own) are new: `Workbook.addCellStyle(wb, name, style)`,
`Workbook.useBuiltInCellStyle(wb, "Heading1")` and `Cell.applyStyle(ws, a, name)`.

---

## 11. Merged cells, names, notes, hyperlinks

```ts before
ws.mergeCells("A1:C1");
ws.getCell("A2").name = "first";
wb.definedNames.add("Sheet1!$A$2:$A$9", "ids");
ws.getCell("B2").note = "Checked by QA";
ws.getCell("C2").value = { text: "Docs", hyperlink: "https://example.com" };
```

```ts
import { Cell, DefinedNames, Workbook, Worksheet } from "documonster/excel";

Worksheet.merge(ws, "A1:C1");
Cell.setNames(ws, "A2", ["first"]);
DefinedNames.add(Workbook.getDefinedNames(wb), "Sheet1!$A$2:$A$9", "ids");
Cell.setNote(ws, "B2", "Checked by QA");
Cell.setValue(ws, "C2", { text: "Docs", hyperlink: "https://example.com" });
```

| exceljs                                              | documonster                                                                                  |
| ---------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| `ws.mergeCells(range)` / `ws.unMergeCells(range)`    | `Worksheet.merge(ws, range)` / `Worksheet.unmerge(ws, range)`                                |
| `ws.mergeCellsWithoutStyle(range)`                   | `Worksheet.mergeWithoutStyle(ws, range)`                                                     |
| `ws.hasMerges` / `ws.model.merges`                   | `Worksheet.hasMerges(ws)` / `Worksheet.getMergedRegions(ws)`                                 |
| `wb.definedNames.add(ref, name)`                     | `DefinedNames.add(Workbook.getDefinedNames(wb), ref, name)`                                  |
| `wb.definedNames.remove(ref, name)`                  | `DefinedNames.remove(dn, ref, name)` — or `DefinedNames.remove(dn, name)` for the whole name |
| `wb.definedNames.getNames(ref)`                      | `DefinedNames.getNames(dn, ref)`                                                             |
| `wb.definedNames.getRanges(name)`                    | `DefinedNames.get(dn, name)?.ranges`                                                         |
| `wb.definedNames.forEach((name, cell) => …)`         | `DefinedNames.list(dn)` — one entry per name, with its `ranges`                              |
| `wb.definedNames.removeAllNames(location)`           | `Cell.clearNames(ws, address)` — the names of that cell                                      |
| `wb.definedNames.model` / `= m`                      | `DefinedNames.getModel(dn)` / `DefinedNames.setModel(dn, m)`                                 |
| _(a name that is a formula)_                         | `DefinedNames.addFormula(dn, name, "Sheet1!$A$1*2")`                                         |
| `cell.note = { texts, margins, protection, editAs }` | `Cell.setNote(ws, a, { texts, margins, protection, editAs })`                                |
| `cell.note = undefined`                              | `Cell.removeNote(ws, a)`                                                                     |

exceljs's internal `Range` objects and the types behind its defined names are mapped in the
[appendix](#appendix-exceljs-internals).

A note has an author in Excel; exceljs could not set it. `Cell.setNote(ws, a, text, { author: "QA" })` does.

---

## 12. Data validation and conditional formatting

The rule objects are exceljs's.

```ts before
ws.getCell("B2").dataValidation = {
  type: "list",
  allowBlank: true,
  formulae: ['"Yes,No"']
};
// exceljs has this at run time, but its index.d.ts does not declare it:
(ws as any).dataValidations.add("C2:C99", {
  type: "whole",
  operator: "between",
  formulae: [1, 10]
});
ws.addConditionalFormatting({
  ref: "D2:D99",
  rules: [
    {
      type: "cellIs",
      operator: "lessThan",
      formulae: [0],
      priority: 1,
      style: { font: { color: { argb: "FFC00000" } } }
    }
  ]
});
```

```ts
import { Cell, ConditionalFormatting, DataValidation } from "documonster/excel";

Cell.setValidation(ws, "B2", {
  type: "list",
  allowBlank: true,
  formulae: ['"Yes,No"']
});
DataValidation.add(ws, "C2:C99", {
  type: "whole",
  operator: "between",
  formulae: [1, 10]
});
ConditionalFormatting.add(ws, {
  ref: "D2:D99",
  rules: [
    {
      type: "cellIs",
      operator: "lessThan",
      formulae: [0],
      priority: 1,
      style: { font: { color: { argb: "FFC00000" } } }
    }
  ]
});
```

| exceljs                                       | documonster                                                                 |
| --------------------------------------------- | --------------------------------------------------------------------------- |
| `cell.dataValidation` (read)                  | `Cell.getValidation(ws, a)`                                                 |
| `ws.dataValidations.find(a)` / `.remove(ref)` | `DataValidation.find(ws, a)` / `DataValidation.remove(ws, ref)`             |
| _(every rule on the sheet)_                   | `DataValidation.list(ws)` / `DataValidation.clear(ws)`                      |
| `ws.addConditionalFormatting(cf)`             | `ConditionalFormatting.add(ws, cf)` → the stored rule set                   |
| `ws.removeConditionalFormatting(…)`           | `ConditionalFormatting.remove(ws, ref)` / `ConditionalFormatting.clear(ws)` |
| `ws.conditionalFormattings`                   | `ConditionalFormatting.list(ws)`                                            |

A rule is checked when you add it (an unknown `type` or `operator`, a missing `formulae`, a bad range) and refused
with `ExcelArgumentError`; exceljs wrote it and Excel dropped it on open.

The rule object is exceljs's — `type`, `operator`, `formulae`, `allowBlank`, `showInputMessage`, `promptTitle`,
`prompt`, `showErrorMessage`, `errorStyle`, `errorTitle`, `error` — and its type is `DataValidationRule` (`DataValidation`
is the namespace). A conditional format's `cfvo` entries are `Cfvo` (exceljs spelled the type `Cvfo`).

---

## 13. Tables

```ts before
ws.addTable({
  name: "Sales",
  ref: "A1",
  headerRow: true,
  totalsRow: true,
  style: { theme: "TableStyleMedium2", showRowStripes: true },
  columns: [
    { name: "Item", totalsRowLabel: "Total" },
    { name: "Qty", totalsRowFunction: "sum" }
  ],
  rows: [
    ["Apples", 3],
    ["Pears", 5]
  ]
});
const table = ws.getTable("Sales");
table.addRow(["Plums", 2]);
table.commit();
```

```ts
import { Table } from "documonster/excel";

Table.add(ws, {
  name: "Sales",
  ref: "A1",
  headerRow: true,
  totalsRow: true,
  style: { theme: "TableStyleMedium2", showRowStripes: true },
  columns: [
    { name: "Item", totalsRowLabel: "Total" },
    { name: "Qty", totalsRowFunction: "sum" }
  ],
  rows: [
    ["Apples", 3],
    ["Pears", 5]
  ]
});
const table = Table.get(ws, "Sales");
if (table) Table.addRow(table, ["Plums", 2]); // no commit: the edit is already in the sheet
```

| exceljs (`table = ws.getTable(n)`)                                               | documonster                                                                                                                                                                  |
| -------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ws.getTable(name)`                                                              | `Table.get(ws, name)` → `undefined` when absent                                                                                                                              |
| `ws.removeTable(name)`                                                           | `Table.remove(ws, name \| table)` → `boolean`                                                                                                                                |
| `table.addRow(values, rowNumber?)`                                               | `Table.addRow(table, values, rowNumber?)`                                                                                                                                    |
| `table.removeRows(index, count?)`                                                | `Table.removeRows(table, index, count?)`                                                                                                                                     |
| `table.addColumn(column, values, index?)`                                        | `Table.addColumn(table, column, values, index?)`                                                                                                                             |
| `TableColumn` (type)                                                             | `Required<TableColumnProperties>`                                                                                                                                            |
| `table.removeColumns(index, count?)`                                             | `Table.removeColumns(table, index, count?)`                                                                                                                                  |
| `table.getColumn(i)`                                                             | `Table.getColumn(table, i)` — then `Table.getColumnName` / `setColumnName`, `getColumnTotalsRowFunction` / `setColumnTotalsRowFunction`, …                                   |
| `table.commit()`                                                                 | delete it — edits apply as they are made                                                                                                                                     |
| `table.ref`, `name`, `displayName`, `theme`                                      | `Table.getRef` / `setRef`, `getName` / `setName`, `getDisplayName` / `setDisplayName`, `getTheme` / `setTheme`                                                               |
| `table.headerRow`, `totalsRow`                                                   | `Table.hasHeaderRow` / `setHeaderRow`, `Table.hasTotalsRow` / `setTotalsRow`                                                                                                 |
| `table.showRowStripes`, `showColumnStripes`, `showFirstColumn`, `showLastColumn` | `Table.hasRowStripes` / `setRowStripes`, `hasColumnStripes` / `setColumnStripes`, `hasFirstColumnStyle` / `setFirstColumnStyle`, `hasLastColumnStyle` / `setLastColumnStyle` |
| `table.model`                                                                    | `Table.getModel(table)` (a copy) / `Table.setModel(table, m)`                                                                                                                |

---

## 14. Images and backgrounds

```ts before
const id = wb.addImage({ filename: "logo.png", extension: "png" });
ws.addImage(id, "B2:D6");
ws.addImage(id, { tl: { col: 1, row: 1 }, ext: { width: 120, height: 40 } });
ws.addBackgroundImage(id);
```

```ts
import { Image } from "documonster/excel";

const id = Image.add(wb, { filename: "logo.png", extension: "png" });
Image.place(ws, id, "B2:D6");
Image.place(ws, id, {
  tl: { col: 1, row: 1 },
  ext: { width: 120, height: 40 }
});
Image.setBackground(ws, id);
```

`Image.add` takes exceljs's `{ filename | buffer | base64, extension }` (`buffer` is a `Uint8Array`; a `Buffer` is
one), plus an optional `svg` beside a raster fallback. `Image.place` takes the same ranges and anchors (`"B2:D6"`,
`{ tl, br }`, `{ tl, ext, editAs }`, `hyperlinks`) and returns the placed image's handle. `ws.getImages()` is
`Image.list(ws)`, and `Image.remove(ws, image)` / `Image.removeBackground(ws)` are new. A range is checked when you
place it (corners inside the grid, `br` not above or left of `tl`); exceljs wrote `NaN` for a malformed one.

**In a browser, give the bytes.** `filename` is read from disk when the workbook is written, so it works in Node only;
in a browser `Workbook.toBuffer` fails with `ExcelUnsupportedFeatureError`. Pass `buffer` (a `Uint8Array`, for example
`new Uint8Array(await (await fetch(url)).arrayBuffer())`) or `base64` instead.

The range types are `AddImageRange` (exceljs's `ImageRange` — the same `tl` / `br` — `ImagePosition` and
`ImageHyperlinkValue` together). An
anchor (`IAnchor`, the `Anchor` class) is `Anchor.Handle`, with exceljs's `nativeCol`, `nativeRow`, `nativeColOff` and
`nativeRowOff` fields. `anchor.col` / `anchor.row` are `Anchor.getCol(anchor)` / `Anchor.getRow(anchor)` — 0-based, with the
offset into the cell as a fraction, as exceljs reads them — and `Anchor.setCol` / `Anchor.setRow` set them;
`anchor.worksheet` is the sheet you placed the image on.

---

## 15. Page setup, views, freeze panes, auto filter, protection

The print and view settings stay plain fields, exactly as in exceljs:

```ts
import { PaperSize } from "documonster/excel";

ws.pageSetup.paperSize = PaperSize.A4;
ws.pageSetup.orientation = "landscape";
ws.pageSetup.printArea = "A1:G20";
ws.pageSetup.printTitlesRow = "1:1";
ws.headerFooter.oddFooter = "Page &P of &N";
ws.properties.tabColor = { argb: "FF00B050" };
ws.state = "hidden";
```

The same objects can be passed to `Workbook.addWorksheet(wb, name, { pageSetup, headerFooter, properties, views, state })`.

| exceljs                                                       | documonster                                                                          |
| ------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| `ws.views = [{ state: "frozen", xSplit: 1, ySplit: 1 }]`      | `Worksheet.freeze(ws, { columns: 1, rows: 1 })` (assigning `ws.views` still works)   |
| `ws.views = [{ state: "split", xSplit: 2000, ySplit: 3000 }]` | `Worksheet.split(ws, { x: 2000, y: 3000 })`                                          |
| _(remove frozen panes)_                                       | `Worksheet.unfreeze(ws)`                                                             |
| _(read the panes)_                                            | `Worksheet.getPanes(ws)`                                                             |
| `ws.autoFilter = "A1:C1"` / `{ from, to }`                    | `Worksheet.setAutoFilter(ws, "A1:C1")` / `Worksheet.setAutoFilter(ws, { from, to })` |
| `ws.protect(password, options)`                               | `await Worksheet.protect(ws, { password, ...options })`                              |
| `ws.unprotect()`                                              | `Worksheet.unprotect(ws)`                                                            |
| `row.addPageBreak()`                                          | `Row.addPageBreak(ws, n)` (and `Column.addPageBreak(ws, c)`)                         |

`PaperSize` is a constant object (`PaperSize.A4`); `ws.pageSetup.paperSize` is still a `number`, so codes it does not
name still work. `Worksheet.verifyPassword(ws, password)` and `Workbook.protect` / `Workbook.verifyPassword` are new.

---

## 16. Streaming

### Writing

```ts before
const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
  filename: "big.xlsx",
  useStyles: true,
  useSharedStrings: true
});
const ws = wb.addWorksheet("Data");
for (const r of rows) ws.addRow(r).commit();
ws.commit();
await wb.commit();
```

```ts
import { Stream, Workbook } from "documonster/excel";

declare const rows: (string | number)[][];

const wb = new Stream.WorkbookWriter({
  filename: "big.xlsx",
  useStyles: true,
  useSharedStrings: true
});
const ws = Workbook.addWorksheet(wb, "Data");
for (const r of rows) Stream.commitRow(ws.addRow(r));
ws.commit();
await wb.commit();
```

**What replaces `ws.addRow(…).commit()`:** the writer's worksheet keeps `addRow`, which returns a row handle
(`Stream.RowHandle`); commit it with `Stream.commitRow(row)` instead of `row.commit()`. `Row.add(ws, values)` also
accepts a streamed sheet and returns the row number, when you do not need the handle. `ws.commit()` and
`await wb.commit()` are unchanged: `wb.commit()` finishes the file and ends the output the writer was given.

Streaming straight into an HTTP response (or any Node writable, or a Web `WritableStream`) — exceljs's
`{ stream: res }` option, unchanged:

```ts
import { createServer } from "node:http";

import { Stream, Workbook } from "documonster/excel";

createServer(async (_req, res) => {
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", 'attachment; filename="report.xlsx"');
  const wb = new Stream.WorkbookWriter({ stream: res });
  const ws = Workbook.addWorksheet(wb, "Data");
  for (let i = 1; i <= 100_000; i++) {
    Stream.commitRow(ws.addRow([i, `row ${i}`]));
  }
  ws.commit();
  await wb.commit(); // ends `res`
}).listen(3000);
```

With a document-model workbook rather than a streaming writer, `await Workbook.writeStream(wb, res)` does the same (§3).

| exceljs                                        | documonster                                                                                                                                                                                                 |
| ---------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `new ExcelJS.stream.xlsx.WorkbookWriter(opts)` | `new Stream.WorkbookWriter(opts)` — `WorkbookStreamWriterOptions` are `Stream.WorkbookWriterOptions`, with the same `filename` or `stream` (a Node or Web writable), `useStyles`, `useSharedStrings`, `zip` |
| `wb.addWorksheet(name, opts)`                  | `Workbook.addWorksheet(wb, name, opts)`                                                                                                                                                                     |
| `ws.addRow(values)` → row                      | `ws.addRow(values)` → row handle                                                                                                                                                                            |
| `row.commit()`                                 | `Stream.commitRow(row)`                                                                                                                                                                                     |
| `row.getCell(c).value = v`                     | `Stream.setCellValue(Stream.getRowCell(row, c), v)`                                                                                                                                                         |
| `row.getCell(c).font = f`                      | `Stream.setCellFont(Stream.getRowCell(row, c), f)` (and `setCellFill`, `setCellBorder`, `setCellAlignment`, `setCellNumFmt`, `setCellProtection`, `setCellNote`)                                            |
| `row.height = h`, `row.font = f`, …            | `Stream.setRowHeight(row, h)`, `Stream.setRowFont(row, f)`, …                                                                                                                                               |
| `ws.columns = defs`                            | `Column.setDefinitions(ws, defs)` — before the first row is committed                                                                                                                                       |
| `ws.mergeCells(range)`                         | `Worksheet.merge(ws, range)`                                                                                                                                                                                |
| `ws.commit()` / `await wb.commit()`            | unchanged                                                                                                                                                                                                   |

Most document-model functions accept a streaming writer too — `Cell.*`, `Row.*`, `Column.*`, `Range.*`, `Image.*`,
`DataValidation.*`, `ConditionalFormatting.*`, `Workbook.addWorksheet`, … — and refuse with
`ExcelStreamStateError` what the writer has already written. The ones that need the whole document (`Row.insert`,
`Row.splice`, `Workbook.removeWorksheet`, …) do not accept it, and TypeScript says so at the call.

### Reading

```ts before
const reader = new ExcelJS.stream.xlsx.WorkbookReader("big.xlsx", {
  worksheets: "emit",
  sharedStrings: "cache"
});
for await (const worksheet of reader) {
  for await (const row of worksheet) {
    console.log(row.values);
  }
}
```

```ts
import { Stream } from "documonster/excel";

const reader = new Stream.WorkbookReader("big.xlsx", {
  worksheets: "emit",
  sharedStrings: "cache"
});
for await (const worksheet of reader) {
  for await (const row of worksheet) {
    console.log(Stream.getRowValues(row, { base: 1 })); // 1-based, as row.values was
  }
}
```

The reader options are exceljs's `WorkbookStreamReaderOptions` (`worksheets`, `sharedStrings`, `hyperlinks`, `styles`,
`entries`), as `Stream.WorkbookReaderOptions`, plus `password`
and size limits; `reader.parse()` and `worksheet.parse()` are still there. The writer's `zip` option is exceljs's
`ArchiverZipOptions` (`comment`, `forceLocalTime`, `forceZip64`, `store`, `zlib`), as `Stream.WorkbookZipOptions`, and
its `zlib` is `Stream.ZlibOptions`. The writer's package-part methods (`addApp`, `addCore`, `addContentTypes`,
`addOfficeRels`, `addSharedStrings`, `addWorkbookRels`, …), a reader's `getImage` and a worksheet reader's
`workbook` / `entry` options (`WorksheetReaderOptions`) are the writer's and reader's own plumbing — `commit()` writes those parts; images you
add are the ids `Image.add` returned. A worksheet must be read before the next one arrives; breaking out of a sheet's rows skips the rest
of it. The readers are also event emitters (`reader.read()` emits `"worksheet"`, a worksheet reader emits `"row"`), as
in exceljs, but the `for await` form is the one to use: it applies backpressure and propagates errors to your loop.

---

## 17. CSV

CSV moved to its own entry, `documonster/excel/csv`, so a bundle that does not use it does not carry the parser.

```ts
import { ExcelCsv } from "documonster/excel/csv";
```

| exceljs                                        | documonster                                                                                                                                         |
| ---------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| `const ws = await wb.csv.readFile(path, opts)` | `const ws = await ExcelCsv.addWorksheet(wb, createReadStream(path), opts)` — or `const wb = await ExcelCsv.readFile(path, opts)` for a new workbook |
| `const ws = await wb.csv.read(stream, opts)`   | `const ws = await ExcelCsv.addWorksheet(wb, stream, opts)` — or `const wb = await ExcelCsv.read(stream, opts)`                                      |
| `await wb.csv.writeFile(path, opts)`           | `await ExcelCsv.writeFile(wb, path, opts)`                                                                                                          |
| `await wb.csv.writeBuffer(opts)`               | `await ExcelCsv.toBuffer(wb, opts)` → `Uint8Array`                                                                                                  |
| `await wb.csv.write(stream, opts)`             | `await ExcelCsv.writeStream(wb, stream, opts)`                                                                                                      |
| _(a string)_                                   | `ExcelCsv.format(wb, opts)`                                                                                                                         |
| `wb.csv.createInputStream(opts)`               | `ExcelCsv.createImportStream(wb, opts)` — a writable you pipe CSV into                                                                              |

The options are flat rather than nested under `fast-csv`'s names:

| exceljs option                                                                            | documonster option                                                                                                                                        |
| ----------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `sheetName`, `sheetId`                                                                    | `sheetName`, `sheetId`                                                                                                                                    |
| `dateFormat`, `dateUTC` (write)                                                           | `dateFormat`, `dateUTC`                                                                                                                                   |
| `dateFormats` (read)                                                                      | `dateFormats`                                                                                                                                             |
| `map(value, index)`                                                                       | `map` — read: `(text, column) => value`; write: `(value, index) => value`                                                                                 |
| `parserOptions.delimiter`, `quote`, `escape`, `headers`, `trim`, …                        | `delimiter`, `quote`, `escape`, `headers`, `trim`, … — flat                                                                                               |
| `formatterOptions.delimiter`, `quote`, `escape`, `quoteColumns`, `quoteHeaders`           | `delimiter`, `quote`, `escape`, `quoteColumns`, `quoteHeaders`                                                                                            |
| `formatterOptions.rowDelimiter`                                                           | `lineEnding`                                                                                                                                              |
| `formatterOptions.writeBOM`                                                               | no option: prepend `"\uFEFF"` to `ExcelCsv.format`'s string, or write it to the stream first                                                              |
| `includeEmptyRows`                                                                        | `includeEmptyRows`                                                                                                                                        |
| `encoding` (write)                                                                        | `encoding` — for `writeFile`; the other writers produce UTF-8                                                                                             |
| `parserOptions.ignoreEmpty`                                                               | `skipEmptyLines`                                                                                                                                          |
| `parserOptions.comment`, `ltrim`, `rtrim`, `skipLines`, `skipRows`, `maxRows`, `encoding` | the same names, flat                                                                                                                                      |
| `parserOptions.strictColumnHandling` / `discardUnmappedColumns`                           | `columnMismatch: { less, more }` — with `headers`, a record with the wrong number of fields fails the read unless `columnMismatch` or `onSkip` handles it |
| `parserOptions.headers: [...]` with `renameHeaders: true`                                 | `headers: [...]` with `skipLines: 1`                                                                                                                      |
| `parserOptions.objectMode` / `formatterOptions.objectMode`                                | nothing: rows go into a sheet                                                                                                                             |
| `formatterOptions.alwaysWriteHeaders`                                                     | `writeHeaders`                                                                                                                                            |
| `formatterOptions.transform` (a whole row)                                                | `mapRow: (values, rowNumber) => row` — after `map`, for every row written; the callback form (`(row, done) => …`) becomes a plain return                  |
| `formatterOptions.includeEndRowDelimiter`                                                 | no option: append the `lineEnding` yourself                                                                                                               |

New options: `encoding` (read any `TextDecoder` label), `decimalSeparator`, `escapeFormulae` (CSV-injection
protection), `append`, `maxRows`. Like exceljs, a read turns numbers, dates, `true` / `false` and error literals into
values; pass `map` to keep text as text. **Two differences in dates:**

- exceljs read a date with no time zone (`2024-01-15`, `2024-01-15T10:00:00`) as that time in the machine's local
  time zone, so the same file gave a different instant on every server; documonster reads it as that calendar date
  and time in UTC, `2024-01-15T00:00:00Z`, as every other date in a workbook is read (§18).
- exceljs left a timestamp that states its zone — `2024-01-15T10:00:00Z`, `2024-01-15T10:00:00+02:00` — as text;
  documonster reads it as that instant. To keep such timestamps as text, pass exceljs's formats:
  `dateFormats: ["MM-DD-YYYY", "YYYY-MM-DD[T]HH:mm:ss", "YYYY-MM-DD"]`.

The option types are `ExcelCsvReadOptions` / `ExcelCsvWriteOptions` (exceljs's `CsvReadOptions` / `CsvWriteOptions`).
fast-csv's types that exceljs re-exported — `HeaderArray`, `HeaderTransformFunction`, `RowArray`, `RowHashArray`,
`RowMap`, `RowTransformCallback`, `RowTransformFunction` — are `CsvHeaderArray`, `CsvHeaderTransformFunction`, … in
`documonster/csv`; `Rows` is `CsvRow[]`, and `QuoteColumns` / `QuoteColumnMap` are the shape of the `quoteColumns`
option. `FastCsvParserOptionsArgs` / `FastCsvFormatterOptionsArgs` have no counterpart: the options are flat.

The standalone CSV parser/formatter (rows in, rows out, no workbook) is `Csv` in `documonster/csv`.

---

## 18. Dates

A cell holds a **calendar** value; a `Date` is an **instant**. Like exceljs, documonster stores a `Date` by its **UTC**
fields and reads a date cell back as a `Date` at UTC midnight (or the UTC time of day). So the exceljs advice still
holds: `new Date(2024, 0, 15)` is midnight _local_ time, which is a different UTC day in every timezone east of UTC.

```ts
import { Cell } from "documonster/excel";

Cell.setValue(ws, "A1", new Date(Date.UTC(2024, 0, 15))); // as in exceljs
Cell.setDateParts(ws, "A2", { year: 2024, month: 1, day: 15 }); // no Date at all
```

Where the runtime has `Temporal` (and your TypeScript `lib` declares it), a civil value needs no `Date` either:
`Cell.setValue(ws, "A3", Temporal.PlainDate.from("2024-01-15"))`.

Differences from exceljs:

- `Cell.getDateParts`, `Cell.setDateParts`, `Cell.getDateKind` and `Cell.getTemporal` read and write the calendar value
  without a `Date`.
- The 1904 date system (`wb.properties.date1904`) is honoured by every reader and writer, the streaming ones included.
- Dates before 1900-03-01 match Excel's serials (Excel counts a 29 February 1900 that never existed).

**When a read gives a `Date` and when a number.** A file stores every date as a serial number; the cell's number format
is what makes it a date. On read, a number cell — or a formula's numeric result — whose format is a date or time format
(`yyyy-mm-dd`, `m/d/yy h:mm`, the built-in date formats) comes back as a `Date`, as in exceljs. The same serial under
`General` or `#,##0` comes back as a `number` (`45306`), and so does a date written by a tool that left the format off;
`Cell.getDateKind(ws, a)` is `undefined` for it. If you know such a column holds dates, convert the serial yourself —
for the default 1900 date system and any date after February 1900,
`new Date(Date.UTC(1899, 11, 30) + serial * 86_400_000)`. For a cell that is a date, `Cell.getDateKind` tells you what
its format says (`"date"`, `"time"`, `"dateTime"`, `"duration"` or `"unknown"`). A `Date` you write without a number
format is written with a date format, so it reads back as a `Date`; set your own with `Cell.setNumFmt` as before.

---

## 19. Errors

exceljs threw plain `Error`s. documonster throws classes you can test for, all subclasses of `ExcelError`
(`isExcelError(e)`):

| Class                                                     | When                                                                  |
| --------------------------------------------------------- | --------------------------------------------------------------------- |
| `ExcelArgumentError`                                      | an argument is wrong; `field` names it and the message names the call |
| `ExcelInvalidAddressError`                                | an address that is not one                                            |
| `ExcelRowOutOfBoundsError`, `ExcelColumnOutOfBoundsError` | a row or column outside the sheet                                     |
| `ExcelWorksheetNameError`                                 | a sheet name Excel refuses (too long, a duplicate, a bad character)   |
| `ExcelMergeConflictError`                                 | a merge overlapping another                                           |
| `ExcelInvalidValueTypeError`                              | a cell value Excel cannot store                                       |
| `ExcelStreamStateError`                                   | editing what a streaming writer has already written                   |
| `ExcelFileError`                                          | the file system refused a read or write (the system error is `cause`) |
| `ExcelEncryptedError`, `ExcelDecryptionError`             | an encrypted file without, or with a wrong, password                  |
| `ExcelLimitExceededError`                                 | a read past `maxRows`, `maxCols` or a `zip` limit                     |
| `XlsxParseError`, `XlsbParseError`                        | a file that is not a valid workbook                                   |

The classes form a hierarchy — `ExcelInvalidAddressError` is an `ExcelArgumentError`, and every one is an
`ExcelError` — so a `catch` can be as broad or as narrow as it needs, and test classes instead of message text:

```ts before
try {
  ws.getCell("not an address").value = 1;
} catch (e) {
  if (e instanceof Error && e.message.includes("Invalid Address")) {
    /* … */
  }
}
```

```ts
import {
  Cell,
  ExcelArgumentError,
  ExcelEncryptedError,
  Workbook,
  isExcelError
} from "documonster/excel";

try {
  Cell.setValue(ws, "not an address", 1);
} catch (e) {
  if (e instanceof ExcelArgumentError) {
    console.error(`bad ${e.field}: ${e.message}`); // a programming error
  } else {
    throw e;
  }
}

try {
  const wb = await Workbook.read(bytes);
} catch (e) {
  if (e instanceof ExcelEncryptedError) {
    // ask for a password, then Workbook.read(bytes, { password })
  } else if (isExcelError(e)) {
    console.error("not a readable workbook:", e.message);
  } else {
    throw e;
  }
}
```

---

## 20. What exceljs does not have, and what documonster does not have

**Not in exceljs, available here:** encrypted workbooks (`password`), XLSB read and write, charts (`Chart.place`),
pivot tables (`Pivot.add`), sparklines, form checkboxes, watermarks and header/footer images, formula calculation
(`ExcelFormula.calculate`), named cell styles, PDF export (`Pdf.fromExcel` in `documonster/pdf`), Markdown
(`documonster/excel/markdown`), and references that follow inserted and deleted rows and renamed sheets.

**In exceljs, with no direct counterpart:**

| exceljs                                        | Instead                                                                        |
| ---------------------------------------------- | ------------------------------------------------------------------------------ |
| `ExcelJS.config.setValue("promise")`           | native promises only                                                           |
| `cell.toCsvString()`                           | `ExcelCsv.format`                                                              |
| `FormulaType` / `cell.formulaType`             | the value's shape (§9)                                                         |
| `wb.nextId`, `wb.clearThemes()`                | nothing: ids are assigned by `Workbook.addWorksheet`; themes are kept as read  |
| `ValueType[number]` reverse lookup             | `Object.entries(ValueType).find(([, v]) => v === t)?.[0]`                      |
| `ReadingOrder` (enum)                          | the strings `Alignment.readingOrder` takes: `"ltr"`, `"rtl"` (and `"context"`) |
| `RelationshipType`, `DocumentType` (enums)     | internal OOXML plumbing; nothing to port                                       |
| `Buffer` (type)                                | `Uint8Array`                                                                   |
| mutating a returned row, column or cell object | the setter for that facet                                                      |

---

## 21. Checklist

1. Replace the dependency and the import: `import ExcelJS from "exceljs"` → named imports from `documonster/excel`
   (and `documonster/excel/csv` for CSV).
2. `new ExcelJS.Workbook()` → `Workbook.create()`; before a read, delete it and take what the read returns.
3. `wb.xlsx.*` / `wb.csv.*` → §3 / §17.
4. Rewrite every `ws.getCell` / `ws.getRow` / `ws.getColumn` chain as a `Cell.*` / `Row.*` / `Column.*` call (§6–§8).
   A cell is `"A1"` or `{ row, col }`.
5. Port `row.values` / `col.values` / `getSheetValues()` with `{ base: 1 }`.
6. `ws.eachRow((row, n) => …)` → `Row.each(ws, n => …)`; `row.eachCell((cell, c) => …)` →
   `Row.eachAddress(ws, n, (address, row, col) => …)`.
7. Delete `table.commit()`, and `row.commit()` outside streaming.
8. Replace `Buffer` assumptions on `toBuffer`'s result (§3).
9. Replace `catch` blocks that matched on messages with the error classes (§19).
10. Type-check. Every remaining `.method` or `.prop` on a handle is a compile error — work through them until the
    project compiles.

### Migration done? checklist

- [ ] `tsc --noEmit` is clean.
- [ ] No deprecated name is left: no strike-through in the editor, and `@typescript-eslint/no-deprecated` reports
      nothing.
- [ ] `grep -rn "=== null\|!== null"` finds no test on a documonster return — absence is `undefined`
      ([Code that compiles but behaves differently](#code-that-compiles-but-behaves-differently)).
- [ ] `grep -rn "Buffer"` finds no `Buffer` method called on a `toBuffer` result (§3).
- [ ] A real file of yours round-trips: `Workbook.read` it, `Workbook.toBuffer` it, and open the result in Excel.

---

## 22. Complete before/after example

```ts before
// before (exceljs)
import ExcelJS from "exceljs";

const wb = new ExcelJS.Workbook();
wb.creator = "Reports";
const ws = wb.addWorksheet("Q1", { views: [{ state: "frozen", ySplit: 1 }] });
ws.columns = [
  { header: "Region", key: "region", width: 16 },
  { header: "Revenue", key: "revenue", width: 14 },
  { header: "Date", key: "date", width: 12 }
];
ws.addRow({
  region: "North",
  revenue: 1200,
  date: new Date(Date.UTC(2024, 0, 31))
});
ws.addRow({
  region: "South",
  revenue: 950,
  date: new Date(Date.UTC(2024, 0, 31))
});
ws.getRow(1).font = { bold: true };
ws.getColumn("revenue").numFmt = "#,##0";
ws.getColumn("date").numFmt = "yyyy-mm-dd";
ws.getCell("B4").value = { formula: "SUM(B2:B3)" };
ws.autoFilter = "A1:C1";
ws.eachRow((row, n) => {
  if (n > 1 && Number(row.getCell(2).value) < 1000) {
    row.getCell(2).font = { color: { argb: "FFC00000" } };
  }
});
await wb.xlsx.writeFile("q1.xlsx");

const again = new ExcelJS.Workbook();
await again.xlsx.readFile("q1.xlsx");
console.log(again.getWorksheet("Q1")?.getRow(2).values);
```

```ts
// after (documonster)
import { Cell, Column, Row, Workbook, Worksheet } from "documonster/excel";

const wb = Workbook.create();
wb.creator = "Reports";
const ws = Workbook.addWorksheet(wb, "Q1");
Worksheet.freeze(ws, { rows: 1 });
Column.setDefinitions(ws, [
  { header: "Region", key: "region", width: 16 },
  { header: "Revenue", key: "revenue", width: 14 },
  { header: "Date", key: "date", width: 12 }
]);
Row.add(ws, {
  region: "North",
  revenue: 1200,
  date: new Date(Date.UTC(2024, 0, 31))
});
Row.add(ws, {
  region: "South",
  revenue: 950,
  date: new Date(Date.UTC(2024, 0, 31))
});
Row.setFont(ws, 1, { bold: true });
Column.setNumFmt(ws, "revenue", "#,##0");
Column.setNumFmt(ws, "date", "yyyy-mm-dd");
Cell.setValue(ws, "B4", { formula: "SUM(B2:B3)" });
Worksheet.setAutoFilter(ws, "A1:C1");
Row.each(ws, n => {
  if (n > 1 && Number(Cell.getValue(ws, { row: n, col: 2 })) < 1000) {
    Cell.setFont(ws, { row: n, col: 2 }, { color: { argb: "FFC00000" } });
  }
});
await Workbook.writeFile(wb, "q1.xlsx");

const again = await Workbook.readFile("q1.xlsx");
const q1 = Workbook.getWorksheet(again, "Q1");
if (q1) console.log(Row.getValues(q1, 2, { base: 1 }));
```

---

## Appendix: exceljs internals

Members exceljs exposed for its own bookkeeping. Few programs touched them; they are here so a port of one that did has an answer.

### `ModelContainer`

| exceljs                                                                                     | documonster                                                                                                                                     |
| ------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| `new ExcelJS.ModelContainer(model).xlsx.writeBuffer()` (exported at run time, not declared) | `const wb = Workbook.create(); Workbook.setModel(wb, model); await Workbook.toBuffer(wb);` — `model` is a workbook model, as `wb.model` gave it |

### `Range`

exceljs's `Range` objects (from its internals: `DefinedNames`, merges) are `Range.*` functions over a handle, made with
`Range.create("A1:B2")` (or a `Location`'s four numbers):

| exceljs `Range` member                                                       | documonster                                                                                                                  |
| ---------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `top`, `left`, `bottom`, `right`, `sheetName`                                | the same fields on the handle; a `Location` (`{ top, left, bottom, right }`) is those four                                   |
| `range.tl` / `range.br`                                                      | `Address.encodeCell({ c: r.left - 1, r: r.top - 1 })` / `Address.encodeCell({ c: r.right - 1, r: r.bottom - 1 })`            |
| `range.$t$l` / `range.$b$r`                                                  | `"$" + Address.encodeCol(r.left - 1) + "$" + r.top` / `"$" + Address.encodeCol(r.right - 1) + "$" + r.bottom`                |
| `range.range` / `range.toString()`                                           | `Range.toString(r)` — `B2:D5`, after the sheet name it was made with (`Sheet1!B2:D5`)                                        |
| `range.$range`                                                               | `Range.toString(r, { absolute: true })` — `$B$2:$D$5`                                                                        |
| `range.shortRange`                                                           | `Range.toString(r, { compact: true })` — a one-cell range as `B2`                                                            |
| `range.$shortRange`                                                          | `Range.toString(r, { absolute: true, compact: true })`                                                                       |
| `count`                                                                      | `Range.getCount(r)`                                                                                                          |
| `decode(…)` / `expand(t, l, b, r)` / `expandRow(row)` / `expandToAddress(a)` | `Range.create(…)` / `Range.expand(r, t, l, b, right)` / `Range.expand` with the row's extent / `Range.expandToAddress(r, a)` |
| `intersects(other)` / `contains(a)` / `containsEx(address)`                  | `Range.intersects(r, other)` / `Range.contains(r, a)` / `Range.containsCell(r, address)`                                     |
| `forEachAddress(cb)`                                                         | `Range.eachAddress(r, cb)`                                                                                                   |

`DefinedNamesRanges` / `DefinedNamesModel` are `DefinedNameModel[]` (`DefinedNames.list(dn)`); `CellMatrix`, the index
behind exceljs's defined names, is internal.

### Bookkeeping with nothing to port

| exceljs                                                                                                                                                                                                                                                                                              | Instead                                                                                                     |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| Internals: `cell.addMergeRef()`, `releaseMergeRef()`, `release()`, `destroy()`, `row.getCellEx()`, `col.equivalentTo()`, `col.isDefault`, `definedNames.getMatrix()` / `addEx()` / `removeEx()` / `getNamesEx()`, the stream writer's `addStyles()` / `addThemes()` / `addWorkbook()` … part writers | bookkeeping exceljs exposed by accident; nothing to port — the operation that used them is the call to make |
