# Migration Guide

This document describes user-facing breaking changes and recommended migrations.

---

## Table of Contents

1. [CSV: `workbook.csv` accessor removed — use Workbook methods directly](#csv-workbookcsv-accessor-removed)
2. [CSV: type aliases removed](#csv-legacy-type-aliases-removed)
3. [CSV: standalone API restructured](#csv-standalone-api-restructured)
4. [CSV: parse options renamed/restructured](#csv-parse-options-renamedrestructured)
5. [CSV: format options renamed](#csv-format-options-renamed)
6. [Stream: `BufferChunk` renamed to `ByteChunk`](#stream-bufferchunk-renamed-to-bytechunk)
7. [Stream: `normalizeWritable` / `Writeable` replaced by `toWritable`](#stream-normalizewritable--writeable-replaced-by-towritable)
8. [Stream: `EventEmitter` moved out of stream module](#stream-eventemitter-moved-out-of-stream-module)
9. [Stream: binary utilities moved to `@utils/binary`](#stream-binary-utilities-moved-to-utilsbinary)
10. [Stream: `once` removed from stream exports](#stream-once-removed-from-stream-exports)
11. [Stream: `BufferedStream.toUint8Array()` now consumes the buffer](#stream-bufferedstreamtouint8array-now-consumes-the-buffer)
12. [Stream: `PullStreamOptions` simplified](#stream-pullstreamoptions-simplified)
13. [Stream: `ReadWriteBufferOptions` type removed](#stream-readwritebufferoptions-type-removed)
14. [Stream: `ITransform._flush` is now optional](#stream-itransform_flush-is-now-optional)
15. [Archive: `UnzipEntry.isDirectory` removed](#archive-unzipentryisdirectory-removed)
16. [Archive: `ZipReader.entries()` signature changed](#archive-zipreaderentries-signature-changed)
17. [Archive: browser entry no longer re-exports archive APIs](#archive-browser-entry-no-longer-re-exports-archive-apis)
18. [Excel: `Image` type renamed to `ImageData`](#excel-image-type-renamed-to-imagedata)
19. [Excel: `ZipOptions` renamed to `WorkbookZipOptions`](#excel-zipoptions-renamed-to-workbookzipoptions)
20. [Excel: error types changed to structured error classes](#excel-error-types-changed-to-structured-error-classes)
21. [Excel: sheet-utils removed — use native Worksheet/Workbook methods](#excel-sheet-utils-removed)
22. [EventEmitter: behavioral changes for Node.js parity](#eventemitter-behavioral-changes-for-nodejs-parity)
23. [Package: new subpath exports](#package-new-subpath-exports)
24. [PDF: `pdf()`, `readPdf()`, `excelToPdf()` are now async](#pdf-pdf-readpdf-exceltopdf-are-now-async)

### Since v9.2.1

25. [Excel: `CellModel.richText` union widened — branch on cell `type` before accessing](#excel-cellmodel-richtext-union-widened)
26. [Excel: `{ richText: [] }` no longer classifies as `RichText`](#excel-empty-richtext-no-longer-classifies-as-richtext)
27. [Excel: `HyperlinkValue.text` setter now drops `richText` to preserve invariants](#excel-hyperlinkvalue-text-setter-drops-richtext)
28. [Excel: `TableProperties.rows` tightened from `any[][]` to typed cell values](#excel-tableproperties-rows-tightened)
29. [Excel: empty `<border>` now parses to `undefined` instead of `{}`](#excel-empty-border-now-parses-to-undefined)
30. [Excel: table names must be unique across the workbook (case-insensitive)](#excel-table-names-must-be-unique-across-workbook)

---

## CSV: `workbook.csv` accessor removed

### What changed

The `workbook.csv` getter (which returned a lazy `CSV` instance) has been removed. CSV operations are now methods directly on the `Workbook` class.

### How to migrate

```ts
// Before
const ws = workbook.csv.load(csvText);
const ws = await workbook.csv.read(stream);
const text = workbook.csv.writeString();
const buf = await workbook.csv.writeBuffer();
await workbook.csv.readFile("data.csv");
await workbook.csv.writeFile("out.csv");

// After
const ws = await workbook.readCsv(csvText);
const ws = await workbook.readCsv(stream);
const text = workbook.writeCsv();
const buf = await workbook.writeCsvBuffer();
await workbook.readCsvFile("data.csv");
await workbook.writeCsvFile("out.csv");
```

New methods also support additional input types:

```ts
// Read from File/Blob (browser)
const ws = await workbook.readCsv(file);

// Stream factories
const readable = workbook.createCsvReadStream();
const writable = workbook.createCsvWriteStream();
```

### Notes

- `readCsv()` is always async; the old `load()` was sync.
- `writeCsv()` returns a `string` synchronously, or accepts a writable stream.
- `readCsvFile()` / `writeCsvFile()` throw in browser environments.

---

## CSV: Legacy type aliases removed

### What changed

The following type aliases have been removed from the main entry point:

- `CsvReadOptions`
- `CsvWriteOptions`
- `CsvStreamReadOptions`
- `CsvStreamWriteOptions`

### How to migrate

Use `CsvOptions` (for workbook-level CSV) or `CsvParseOptions` / `CsvFormatOptions` (for standalone parsing/formatting).

```ts
// Before
import type { CsvReadOptions, CsvWriteOptions } from "@cj-tech-master/excelts";

// After
import type { CsvOptions } from "@cj-tech-master/excelts";
```

---

## CSV: standalone API restructured

### What changed

The CSV module has been restructured from `csv-core.ts` / `csv-stream.ts` into a modular `csv/` directory. The standalone functions (`parseCsv`, `formatCsv`, `parseCsvStream`) were previously internal-only; they are now publicly exported via the `@cj-tech-master/excelts/csv` subpath. The old `parseCsvStream()` async generator is replaced by `parseCsvRows()`.

`CsvParserStream` and `CsvFormatterStream` remain available from the main entry point, and two new factory functions have been added.

### How to migrate

```ts
// Standalone CSV functions — now publicly available
import {
  parseCsv,
  formatCsv,
  parseCsvRows,
  parseCsvAsync,
  CsvParserStream,
  CsvFormatterStream,
  createCsvParserStream,
  createCsvFormatterStream
} from "@cj-tech-master/excelts/csv";

// Stream classes also available from main entry
import {
  CsvParserStream,
  CsvFormatterStream,
  createCsvParserStream,
  createCsvFormatterStream
} from "@cj-tech-master/excelts";
```

New standalone capabilities:

```ts
// Async parsing (string, AsyncIterable, or ReadableStream)
const result = await parseCsvAsync(input, options);

// Streaming row-by-row async generator (replaces parseCsvStream)
for await (const row of parseCsvRows(input, options)) { ... }

// Progress-aware parsing
const result = await parseCsvWithProgress(input, options, onProgress);

// Delimiter auto-detection
const delimiter = detectDelimiter(csvText);
```

---

## CSV: parse options renamed/restructured

### What changed

Several `CsvParseOptions` fields have been renamed or restructured:

| Old                                               | New                              | Notes                                                 |
| ------------------------------------------------- | -------------------------------- | ----------------------------------------------------- |
| `transform`                                       | `rowTransform`                   | Renamed to avoid confusion with stream `.transform()` |
| `strictColumnHandling` + `discardUnmappedColumns` | `columnMismatch: { less, more }` | Two booleans → structured config with strategies      |
| `ignoreEmpty`                                     | `skipEmptyLines`                 | Also accepts `"greedy"` for whitespace-only rows      |

### How to migrate

```ts
// Before
parseCsv(input, {
  transform: row => row.map(v => v.trim()),
  strictColumnHandling: true,
  discardUnmappedColumns: true,
  ignoreEmpty: true
});

// After
parseCsv(input, {
  rowTransform: row => row.map(v => v.trim()),
  columnMismatch: { less: "error", more: "truncate" },
  skipEmptyLines: true
});
```

---

## CSV: format options renamed

### What changed

| Old                      | New                                      |
| ------------------------ | ---------------------------------------- |
| `rowDelimiter`           | `lineEnding`                             |
| `writeBOM`               | `bom`                                    |
| `includeEndRowDelimiter` | `trailingNewline`                        |
| `alwaysWriteHeaders`     | `writeHeaders: true`                     |
| `transform`              | `typeTransform` (per-type transform map) |

### How to migrate

```ts
// Before
formatCsv(data, {
  rowDelimiter: "\r\n",
  writeBOM: true,
  includeEndRowDelimiter: true,
  alwaysWriteHeaders: true
});

// After
formatCsv(data, {
  lineEnding: "\r\n",
  bom: true,
  trailingNewline: true,
  writeHeaders: true
});
```

---

## Stream: `BufferChunk` renamed to `ByteChunk`

### What changed

The `BufferChunk` class has been renamed to `ByteChunk`.

### How to migrate

```ts
// Before
import { BufferChunk } from "@cj-tech-master/excelts";

// After
import { ByteChunk } from "@cj-tech-master/excelts";
```

---

## Stream: `normalizeWritable` / `Writeable` replaced by `toWritable`

### What changed

The `normalizeWritable` function (also exported as `Writeable`) has been replaced by `toWritable`.

### How to migrate

```ts
// Before
import { Writeable } from "@cj-tech-master/excelts";
const writable = Writeable(target);

// After
import { toWritable } from "@cj-tech-master/excelts";
const writable = toWritable(target);
```

---

## Stream: `EventEmitter` moved out of stream module

### What changed

`EventEmitter` is no longer exported from the stream module. It has been moved to `@utils/event-emitter`.

### How to migrate

If you were importing `EventEmitter` from the stream module internals, update the import path. Note: `EventEmitter` is not re-exported from the main entry point.

---

## Stream: binary utilities moved to `@utils/binary`

### What changed

The following utilities are no longer exported from the stream module:

- `textEncoder`, `textDecoder`
- `stringToUint8Array`, `uint8ArrayToString`
- `uint8ArrayEquals`, `uint8ArrayIndexOf`, `uint8ArraySlice`
- `toUint8Array`, `bufferToString`, `concatUint8Arrays`

### How to migrate

Some of these are now available from the main entry point:

```ts
// Available from main entry
import {
  concatUint8Arrays,
  toUint8Array,
  stringToUint8Array,
  uint8ArrayToString
} from "@cj-tech-master/excelts";
```

For others (`textEncoder`, `textDecoder`, `uint8ArrayEquals`, `uint8ArrayIndexOf`, `uint8ArraySlice`, `bufferToString`), import from the internal `@utils/binary` module if needed.

---

## Stream: `once` removed from stream exports

### What changed

The `once` function is no longer exported from the stream module.

### How to migrate

Use `onceEvent` instead (new export):

```ts
// Before
import { once } from "...stream module...";

// After
import { onceEvent } from "@cj-tech-master/excelts/stream";
```

---

## Stream: `BufferedStream.toUint8Array()` now consumes the buffer

### What changed

`BufferedStream.toUint8Array()` now **consumes** internal buffers (resets state to empty after the call). Previously it was non-consuming.

### How to migrate

If you called `toUint8Array()` multiple times on the same `BufferedStream`, you'll need to store the result from the first call. Subsequent calls will return an empty array.

```ts
// Before (could call multiple times)
const a = stream.toUint8Array();
const b = stream.toUint8Array(); // same result

// After (save the result)
const bytes = stream.toUint8Array();
// stream is now empty
```

---

## Stream: `PullStreamOptions` simplified

### What changed

`PullStreamOptions` was an interface with `{ objectMode?, highWaterMark? }`. It is now `type PullStreamOptions = object` (empty, reserved for future use).

### How to migrate

If you were passing `objectMode` or `highWaterMark` to `PullStream`, these options are no longer accepted. Remove them from your constructor calls.

---

## Stream: `ReadWriteBufferOptions` type removed

### What changed

The `ReadWriteBufferOptions` type has been deleted entirely.

### How to migrate

If you referenced this type, replace it with `StreamOptions` or inline the fields you need (`{ highWaterMark?, objectMode? }`).

---

## Stream: `ITransform._flush` is now optional

### What changed

`ITransform._flush` was a required method. It is now optional (`_flush?(callback)`).

### Impact

This is non-breaking for implementations that already define `_flush`. If you had type checks requiring `_flush` to exist, update them to handle `undefined`.

---

## Archive: `UnzipEntry.isDirectory` removed

### What changed

`UnzipEntry.isDirectory: boolean` has been removed. It is replaced by `UnzipEntry.type: ZipEntryType` where `ZipEntryType = "file" | "directory" | "symlink"`.

### How to migrate

```ts
// Before
if (entry.isDirectory) { ... }

// After
if (entry.type === "directory") { ... }

// Also available now:
if (entry.type === "symlink") { ... }
```

New properties on `UnzipEntry`:

- `mode: number` — Unix file permissions
- `linkTarget?: string` — symlink target (populated after `bytes()`)
- `readableStream()` — returns a WHATWG `ReadableStream<Uint8Array>`

---

## Archive: `ZipReader.entries()` signature changed

### What changed

`ZipReader.entries()` now accepts an optional `UnzipStreamOptions` parameter. It is no longer an `async *` generator directly, but returns an operation-backed `AsyncIterable`.

### How to migrate

```ts
// Before
for await (const entry of reader.entries()) { ... }

// After (still works, backward-compatible with no args)
for await (const entry of reader.entries()) { ... }

// New: with options
for await (const entry of reader.entries({ signal, onProgress })) { ... }

// New: WHATWG ReadableStream adapter
const stream = reader.entriesStream();
```

Similarly, `ZipArchive.stream()`, `.bytes()`, `.pipeTo()` now accept optional `ZipStreamOptions` (signal, onProgress). Existing no-arg calls are unaffected.

---

## Archive: browser entry no longer re-exports archive APIs

### What changed

The browser entry point (`index.browser.ts`) no longer re-exports archive APIs (`zip`, `unzip`, `ZipArchive`, `ZipReader`, `crc32`, `compress`, `decompress`, etc.).

### How to migrate

Use the new `@cj-tech-master/excelts/zip` subpath export:

```ts
// Before
import { zip, unzip, ZipArchive, ZipReader, crc32 } from "@cj-tech-master/excelts";

// After (browser)
import { zip, unzip, ZipArchive, ZipReader, crc32 } from "@cj-tech-master/excelts/zip";
```

The Node.js entry point is unaffected — archive APIs are still re-exported there.

---

## Excel: `Image` type renamed to `ImageData`

### What changed

The `Image` interface has been renamed to `ImageData`. A deprecated type alias `Image = ImageData` is provided for backward compatibility.

### How to migrate

```ts
// Before
import type { Image } from "@cj-tech-master/excelts";

// After
import type { ImageData } from "@cj-tech-master/excelts";
```

The `Workbook.addImage()` parameter type is now `ImageData` instead of `Image`.

---

## Excel: `ZipOptions` renamed to `WorkbookZipOptions`

### What changed

The `ZipOptions` interface (used in `WorkbookWriterOptions`) has been renamed to `WorkbookZipOptions`. A deprecated type alias `ZipOptions = WorkbookZipOptions` is provided for backward compatibility.

### How to migrate

```ts
// Before
import type { ZipOptions } from "@cj-tech-master/excelts";

// After
import type { WorkbookZipOptions } from "@cj-tech-master/excelts";
```

---

## Excel: error types changed to structured error classes

### What changed

All generic `throw new Error(...)` in the Excel module have been replaced with typed error subclasses extending `ExcelError` → `BaseError` → `Error`:

| Error Class              | Used For                     |
| ------------------------ | ---------------------------- |
| `ExcelError`             | General Excel errors         |
| `WorksheetNameError`     | Worksheet name validation    |
| `InvalidAddressError`    | Invalid cell address/range   |
| `ColumnOutOfBoundsError` | Column out of bounds         |
| `RowOutOfBoundsError`    | Row out of bounds            |
| `MergeConflictError`     | Merging already-merged cells |
| `InvalidValueTypeError`  | Invalid cell value type      |
| `XmlParseError`          | XML parsing errors           |
| `ExcelNotSupportedError` | Unsupported operation        |
| `ExcelFileError`         | File I/O failures            |
| `ExcelStreamStateError`  | Stream invalid state         |
| `ExcelDownloadError`     | HTTP download failures       |
| `PivotTableError`        | Pivot table errors           |
| `TableError`             | Table errors                 |
| `ImageError`             | Image processing errors      |
| `MaxItemsExceededError`  | Item limit exceeded          |

### Impact

- All new error classes extend `Error`, so existing `catch (e)` blocks will still work.
- Code that checks `error.message` text may need updating.
- You can now use `instanceof` or the `isExcelError()` type guard for precise error handling.

### How to migrate

```ts
// Before
try {
  worksheet.mergeCells("A1:B2");
} catch (e) {
  if (e.message.includes("Cannot merge")) { ... }
}

// After
import { MergeConflictError, isExcelError } from "@cj-tech-master/excelts";

try {
  worksheet.mergeCells("A1:B2");
} catch (e) {
  if (e instanceof MergeConflictError) { ... }
  // Or broadly:
  if (isExcelError(e)) { ... }
}
```

---

## Excel: sheet-utils removed

### What changed

The `sheet-utils` module (standalone utility functions ported from SheetJS) has been removed. All functionality is now available as native methods on `Worksheet` and `Workbook`, providing a more idiomatic API.

The following exports have been removed from the main entry point:

- `jsonToSheet()`, `sheetAddJson()`, `sheetToJson()`
- `aoaToSheet()`, `sheetAddAoa()`, `sheetToAoa()`
- `bookNew()`, `bookAppendSheet()`
- `JSONRow`, `JSON2SheetOpts`, `Sheet2JSONOpts`, `SheetAddJSONOpts`, `AOA2SheetOpts`

The following exports are **unchanged** (moved to a dedicated `address` module internally, but same signatures):

- `decodeCol`, `encodeCol`, `decodeRow`, `encodeRow`
- `decodeCell`, `encodeCell`, `decodeRange`, `encodeRange`
- `CellAddress`, `SheetRange`, `Origin`

### How to migrate

#### Worksheet data conversion

| Before                                     | After                                |
| ------------------------------------------ | ------------------------------------ |
| `sheetToJson(ws)`                          | `ws.toJSON()`                        |
| `sheetToJson(ws, { header: 1 })`           | `ws.toJSON({ header: 1 })`           |
| `sheetToJson(ws, { header: "A" })`         | `ws.toJSON({ header: "A" })`         |
| `sheetToJson(ws, { raw: false })`          | `ws.toJSON({ raw: false })`          |
| `sheetToJson(ws, { defval: "" })`          | `ws.toJSON({ defaultValue: "" })`    |
| `sheetToJson(ws, { blankrows: true })`     | `ws.toJSON({ blankRows: true })`     |
| `sheetAddJson(ws, data)`                   | `ws.addJSON(data)`                   |
| `sheetAddJson(ws, data, { origin: "C3" })` | `ws.addJSON(data, { origin: "C3" })` |
| `sheetToAoa(ws)`                           | `ws.toAOA()`                         |
| `sheetAddAoa(ws, data)`                    | `ws.addAOA(data)`                    |
| `sheetAddAoa(ws, data, { origin: -1 })`    | `ws.addAOA(data, { origin: -1 })`    |

#### Creating worksheets from data

| Before                                    | After                                                           |
| ----------------------------------------- | --------------------------------------------------------------- |
| `jsonToSheet(data)`                       | `wb.addWorksheet("Sheet1").addJSON(data)`                       |
| `jsonToSheet(data, { skipHeader: true })` | `wb.addWorksheet("Sheet1").addJSON(data, { skipHeader: true })` |
| `aoaToSheet(data)`                        | `wb.addWorksheet("Sheet1").addAOA(data)`                        |
| `aoaToSheet(data, { origin: "C3" })`      | `wb.addWorksheet("Sheet1").addAOA(data, { origin: "C3" })`      |

#### Workbook functions

| Before                            | After                        |
| --------------------------------- | ---------------------------- |
| `bookNew()`                       | `new Workbook()`             |
| `bookAppendSheet(wb, ws, "name")` | `wb.importSheet(ws, "name")` |

#### Option type renames

| Before             | After                       |
| ------------------ | --------------------------- |
| `JSON2SheetOpts`   | `AddJSONOptions`            |
| `SheetAddJSONOpts` | `AddJSONOptions` (merged)   |
| `Sheet2JSONOpts`   | `SheetToJSONOptions`        |
| `AOA2SheetOpts`    | `AddAOAOptions`             |
| `JSONRow`          | `Record<string, CellValue>` |
| `defval`           | `defaultValue`              |
| `blankrows`        | `blankRows`                 |

#### New: chaining support

`addJSON()` and `addAOA()` return `this`, enabling chaining:

```ts
const wb = new Workbook();
wb.addWorksheet("Data")
  .addJSON([{ name: "Alice", age: 30 }])
  .addAOA([["extra row"]], { origin: -1 });
```

### Notes

- Address functions (`decodeCol`, `encodeCol`, etc.) are unchanged — same signatures, same imports.
- `cellDates` option has been removed (it was declared but never implemented).
- The deprecated `Range` type alias (for `SheetRange`) has been removed. Use `SheetRange` directly.
- **`AutoFilter` object form**: `{ row, column }` has been renamed to `{ row, col }` to be consistent with all other address types (`Address`, `ImageAnchor`, `ImagePosition`, etc.). Update any code that uses the object form:
  ```ts
  // Before
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 3 } };
  // After
  ws.autoFilter = { from: { row: 1, col: 1 }, to: { row: 1, col: 3 } };
  ```
  String-form auto filters (`ws.autoFilter = "A1:C1"`) are unaffected.
- **`WorksheetModel.merges`** has been removed. The model now uses `mergeCells` exclusively (matching the OOXML `<mergeCells>` element name). If you access the model directly, replace `model.merges` with `model.mergeCells`.

---

## EventEmitter: behavioral changes for Node.js parity

### What changed

The `EventEmitter` class (moved to `@utils/event-emitter`) now has closer Node.js parity with several behavioral changes:

| Behavior                         | Before                                 | After                                                      |
| -------------------------------- | -------------------------------------- | ---------------------------------------------------------- |
| `emit("error")` with no listener | Returned `false` silently              | **Throws** the error (matches Node.js)                     |
| Listener errors during `emit()`  | Caught and re-emitted as `"error"`     | **Propagate** to caller (matches Node.js)                  |
| `removeAllListeners()`           | Silently deleted                       | Emits `removeListener` for each (matches Node.js)          |
| `off()` removal order            | Removed first/oldest match (`indexOf`) | Removes most-recent match (`lastIndexOf`, matches Node.js) |
| `on()` / `prependListener()`     | Emitted `newListener` after adding     | Emits `newListener` **before** adding (matches Node.js)    |

### Impact

- If you relied on `emit("error")` silently returning `false`, you must now add an error listener.
- If you relied on listener errors being caught and re-emitted, add your own try/catch in listeners.
- The `EventListener` type is no longer exported from `@utils/event-emitter`, but is still available from the stream module (`@cj-tech-master/excelts/stream`).

### How to migrate

```ts
// Ensure you have an error listener
emitter.on("error", err => {
  /* handle */
});
```

---

## Package: new subpath exports

### What changed

Three new subpath exports have been added to `package.json`:

| Subpath                          | Description                                         |
| -------------------------------- | --------------------------------------------------- |
| `@cj-tech-master/excelts/zip`    | Archive module (ZIP, TAR, compression, encryption)  |
| `@cj-tech-master/excelts/csv`    | CSV module (parse, format, stream, utilities)       |
| `@cj-tech-master/excelts/stream` | Stream module (cross-platform streaming primitives) |

### How to use

```ts
// Import only the archive module
import { ZipArchive, ZipReader, TarArchive } from "@cj-tech-master/excelts/zip";

// Import only the CSV module
import { parseCsv, formatCsv, CsvParserStream } from "@cj-tech-master/excelts/csv";

// Import only the stream module
import { Readable, Writable, Transform, pipeline } from "@cj-tech-master/excelts/stream";
```

Each subpath supports `import` (ESM), `require` (CJS), and browser-specific conditions.

---

## New capabilities (non-breaking)

These are additive features — no migration needed, but worth knowing about:

### Archive module

- **TAR support**: `TarArchive`, `TarReader`, `tar()`, `untar()`, `parseTar()`
- **TAR+Gzip** (Node-only): `TarGzArchive`, `targz()`, `parseTarGz()`
- **ZIP editing**: `ZipEditor`, `editZip()`, `ZipEditPlan`
- **HTTP Range reading**: `RemoteZipReader`, `HttpRangeReader`
- **Encryption**: ZipCrypto and AES-256 support (read + write)
- **Gzip/Zlib**: `gzip()`, `gunzip()`, `zlib()`, `unzlib()`, `decompressAuto()`
- **ZIP64**: Large file support via `zip64` option
- **Progress/abort**: All operations support `onProgress` and `signal: AbortSignal`
- **Structured errors**: `ArchiveError`, `ZipParseError`, `DecryptionError`, `PasswordRequiredError`, etc.
- **Worker pool** (browser): Parallel compression via Web Workers

### Excel module

- **Structured errors**: 15+ typed error classes (see above)
- **New exports**: `DefinedNames`, `DateParser`, `DateFormatter`, `BaseError`, binary utilities

### CSV module

- **Async parsing**: `parseCsvAsync()`, `parseCsvRows()`, `parseCsvWithProgress()`
- **Dynamic typing**: Auto-convert strings to `number`, `boolean`, `null`, `Date`
- **Delimiter detection**: `detectDelimiter()`, `detectLinebreak()`
- **CSV generation**: `csvGenerate()` for test data
- **Formula escaping**: `escapeFormulae: true` for CSV injection protection
- **Structured errors**: `CsvError`, `CsvWorkerError`

### Stream module

- **Type guards**: `isReadableStream()`, `isWritableStream()`, `isAsyncIterable()`, `isTransformStream()`
- **New errors**: `StreamError`, `StreamStateError`, `StreamTypeError`
- **ReadableStream interop**: `isReadableStreamLike()`, `readableStreamToAsyncIterable()`
- **Event utilities**: `onceEvent()`

---

## PDF: `pdf()`, `readPdf()`, `excelToPdf()` are now async

### What changed

All three PDF public APIs now return `Promise` instead of a synchronous result. They yield to the event loop between each output page during layout, rendering, and reading, preventing large documents from blocking the main thread.

| Function                         | Before          | After                    |
| -------------------------------- | --------------- | ------------------------ |
| `pdf(input, options?)`           | `Uint8Array`    | `Promise<Uint8Array>`    |
| `readPdf(data, options?)`        | `ReadPdfResult` | `Promise<ReadPdfResult>` |
| `excelToPdf(workbook, options?)` | `Uint8Array`    | `Promise<Uint8Array>`    |

The previous async variants (`pdfAsync`, `readPdfAsync`, `excelToPdfAsync`) have been removed — the base names are now async.

### How to migrate

Add `await` to every call site:

```ts
// Before
import { pdf, readPdf, excelToPdf } from "@cj-tech-master/excelts/pdf";

const bytes = pdf([
  ["Name", "Age"],
  ["Alice", 30]
]);
const result = readPdf(pdfBytes);
const pdfOut = excelToPdf(workbook);

// After
const bytes = await pdf([
  ["Name", "Age"],
  ["Alice", 30]
]);
const result = await readPdf(pdfBytes);
const pdfOut = await excelToPdf(workbook);
```

If you were using the short-lived `*Async` variants, drop the suffix:

```ts
// Before
import { pdfAsync, readPdfAsync, excelToPdfAsync } from "@cj-tech-master/excelts/pdf";

const bytes = await pdfAsync(data);
const result = await readPdfAsync(pdfBytes);

// After
import { pdf, readPdf, excelToPdf } from "@cj-tech-master/excelts/pdf";

const bytes = await pdf(data);
const result = await readPdf(pdfBytes);
```

### Notes

- The function signatures (parameters and return shape) are unchanged — only the return type is wrapped in `Promise`.
- The containing function must be `async`, or you must use `.then()`.
- Top-level `await` works in ESM modules (Node.js 14.8+, all modern browsers).

---

# Since v9.2.1

The changes below were introduced after the v9.2.1 release. Consumers upgrading from v9.2.1 or later should review these entries.

Also introduced in this window (**purely additive — no migration needed**, included here for awareness):

- **Formula calculation engine** at the `@cj-tech-master/excelts/formula` subpath — 433 Excel functions, tokenizer/parser, dependency graph, dynamic-array spill. Opt-in via `installFormulaEngine()` or used functionally via `calculateFormulas(workbookLike)`. The engine has zero excel runtime dependencies and can drive any `WorkbookLike` host.
- **`Workbook.calculateFormulas()`** — requires `installFormulaEngine()` from the formula subpath first; throws with a pointer to the installer otherwise.
- **Dynamic-array formulas** — `FILTER`, `SORT`, `UNIQUE`, `XLOOKUP`, `SEQUENCE`, spill-error detection, ghost cells.
- **External workbook links** — `[Book.xlsx]Sheet!A1` now round-trips through load/save.
- **Workbook structure protection**, public `Workbook.defaultFont`, `absoluteAnchor` image positioning, `ignoredErrors` support.

The breaking changes follow.

---

## <a id="excel-cellmodel-richtext-union-widened"></a>Excel: `CellModel.richText` union widened — branch on cell `type` before accessing

### What changed

`CellModel.richText` used to be `CellRichTextValue | undefined` (always the `{ richText: RichText[] }` wrapper shape). It is now `CellRichTextValue | RichText[] | undefined`:

- When `model.type === Cell.Types.RichText`, the wrapper shape is used (unchanged).
- When `model.type === Cell.Types.Hyperlink`, a bare `RichText[]` array is used to carry the hyperlink's formatted display runs.

This lets formula + hyperlink and rich-text hyperlink cells round-trip without data loss.

### How to migrate

If you read `cell.model.richText` directly, branch on the shape first.

```ts
// Before
const runs = cell.model.richText?.richText;

// After
const rt = cell.model.richText;
const runs = Array.isArray(rt) ? rt : rt?.richText;

// or branch on cell type
const runs =
  cell.model.type === Cell.Types.Hyperlink
    ? (cell.model.richText as RichText[] | undefined)
    : (cell.model.richText as CellRichTextValue | undefined)?.richText;
```

Reading `cell.value` (the public getter) is unaffected — it still returns the wrapper shape for `RichText` cells and the `CellHyperlinkValue` shape for `Hyperlink` cells.

---

## <a id="excel-empty-richtext-no-longer-classifies-as-richtext"></a>Excel: `{ richText: [] }` no longer classifies as `RichText`

### What changed

`Cell.Value.getType` previously returned `Cell.Types.RichText` for `{ richText: [] }` (empty array is truthy in JavaScript). It now requires `richText.length > 0`. An empty-runs object now falls through to `Cell.Types.JSON`.

### How to migrate

If you previously used `{ richText: [] }` to represent an empty rich-text cell, switch to one of:

```ts
// Before
cell.value = { richText: [] }; // classified as RichText

// After — options:
cell.value = null; // empty cell
cell.value = ""; // empty string
cell.value = { richText: [{ text: "" }] }; // RichText with a single empty run
```

---

## <a id="excel-hyperlinkvalue-text-setter-drops-richtext"></a>Excel: `HyperlinkValue.text` setter now drops `richText` to preserve invariants

### What changed

A hyperlink cell now enforces the invariant `text === flatten(richText)`. Assigning to `cell.value.text` on a hyperlink cell clears any previously-set `richText` runs. Previously, `text` and `richText` could drift out of sync silently.

### How to migrate

Set the whole `value` object in one go, or set `richText` (which auto-derives `text`):

```ts
// Before — could produce inconsistent text/richText state
cell.value = { text: "old", hyperlink: "https://example.com" };
(cell.value as CellHyperlinkValue).text = "new";
// richText (if any) was preserved even though text diverged

// After — option 1: set the whole value at once
cell.value = {
  text: "new",
  hyperlink: "https://example.com",
  richText: [{ text: "new" }]
};

// After — option 2: drive via richText (text is derived)
cell.value = {
  richText: [{ text: "new", font: { bold: true } }],
  hyperlink: "https://example.com"
};
```

---

## <a id="excel-tableproperties-rows-tightened"></a>Excel: `TableProperties.rows` tightened from `any[][]` to typed cell values

### What changed

`TableProperties.rows` was typed as `any[][]`. It is now `Array<Array<CellValue | CellFormulaValue>>`. This is a TypeScript-level tightening only — runtime behaviour is unchanged — but code passing non-cell-value shapes (plain objects, custom class instances) will now fail to type-check.

### How to migrate

Ensure table row values conform to `CellValue` (scalars, dates, rich text, hyperlinks, errors) or `CellFormulaValue`:

```ts
// Before — any[][] accepted anything
worksheet.addTable({
  name: "T1",
  ref: "A1",
  columns: [{ name: "A" }, { name: "B" }],
  rows: [[customObject, "text"]] // compiled, even though customObject was invalid
});

// After — narrow the value shape, or cast if you know what you're doing
worksheet.addTable({
  name: "T1",
  ref: "A1",
  columns: [{ name: "A" }, { name: "B" }],
  rows: [[42, "text"]] // all cells are CellValue
});

// Escape hatch if custom payloads are intentional:
worksheet.addTable({
  name: "T1",
  ref: "A1",
  columns: [{ name: "A" }, { name: "B" }],
  rows: [[customObject, "text"]] as unknown as Array<Array<CellValue | CellFormulaValue>>
});
```

---

## <a id="excel-empty-border-now-parses-to-undefined"></a>Excel: empty `<border>` now parses to `undefined` instead of `{}`

### What changed

Reading an XLSX that referenced `borderId=0` (the default empty border `<border><left/><right/><top/><bottom/><diagonal/></border>`) used to populate `cell.border` with an empty object `{}`. It now parses to `undefined`, matching the declared type and the semantic "no border".

### Impact

Code that used `cell.border` as a truthiness check will now see `undefined` (falsy) for no-border cells where previously it saw `{}` (truthy).

### How to migrate

```ts
// Before — worked accidentally because {} is truthy
if (cell.border) {
  // always ran for every parsed cell, even ones without any border
}

// After — use optional chaining or explicit shape check
if (cell.border?.top || cell.border?.right || cell.border?.bottom || cell.border?.left) {
  // runs only for cells that actually have a border
}
```

If you iterate border edges, the safer pattern is:

```ts
const { top, right, bottom, left } = cell.border ?? {};
```

---

## <a id="excel-table-names-must-be-unique-across-workbook"></a>Excel: table names must be unique across the workbook (case-insensitive)

### What changed

`Worksheet.addTable({ name, ... })` now throws when another table in the same workbook (any worksheet) already uses that name, compared case-insensitively. Excel itself requires workbook-wide table name uniqueness; previously the library silently allowed duplicates, which produced files that Excel refused to open.

The check also catches:

- Sanitize collisions (`"My Table"` and `"My_Table"` both normalise to `"My_Table"`).
- Same-worksheet duplicates.

### How to migrate

Ensure table names are unique per workbook. If your code was relying on duplicate names being silently accepted, rename so every `addTable` call uses a distinct identifier:

```ts
// Before — silently accepted, produced a broken XLSX
const ws1 = wb.addWorksheet("Sheet1");
const ws2 = wb.addWorksheet("Sheet2");
ws1.addTable({ name: "Data", ref: "A1", columns: [...], rows: [...] });
ws2.addTable({ name: "Data", ref: "A1", columns: [...], rows: [...] }); // now throws

// After — use distinct names
ws1.addTable({ name: "Sheet1Data", ref: "A1", columns: [...], rows: [...] });
ws2.addTable({ name: "Sheet2Data", ref: "A1", columns: [...], rows: [...] });
```

If you do not control the name (e.g. you are importing user data), check uniqueness before calling `addTable`:

```ts
const existing = wb.worksheets.flatMap(ws => ws.tables.map(t => t.name.toLowerCase()));
const unique = existing.includes(name.toLowerCase()) ? `${name}_${Date.now()}` : name;
ws.addTable({ name: unique, ... });
```

- Error handling is the same — errors are thrown (rejected) with the same `PdfError` / `PdfStructureError` types. Use `try/catch` or `.catch()`.
