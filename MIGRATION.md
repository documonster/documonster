# Migration Guide

Breaking changes for `documonster` will be documented here as they occur, release by release. Moving to `documonster`
from another library has its own guide: [from `exceljs`](MIGRATION_EXCELJS_TO_DOCUMONSTER.md) and
[from `@cj-tech-master/excelts`](MIGRATION_EXCELTS_TO_DOCUMONSTER.md).

## Unreleased

This section covers everything that changed since **0.21.1**, in `documonster` and in `@documonster/mcp`, and is
released as **0.22.0** (the changes are breaking, and before 1.0 a breaking release bumps the minor version). Until
0.22.0 is on npm — `npm view documonster version` — `npm install documonster` still installs 0.21.1; then upgrade with
`npm install documonster@^0.22.0` (`@documonster/mcp` moves with it: the two always share one version). It is long
because the public API was brought to one contract before 1.0; most of it is deprecations that keep working. Three
parts work together:

- **This section** says what to do, grouped by theme, with the reason for each change.
- **[docs/api-changes-since-v0.21.1.md](docs/api-changes-since-v0.21.1.md)** is the exhaustive per-name diff of the
  published declarations: every export of every subpath (Node and browser) and of `@documonster/mcp`, with each
  namespace member, class member and option field that was removed, changed signature, became deprecated or was
  added, old and new signatures side by side (a very long signature is shown around what changed), plus the MCP
  tools' input schemas.
- **[Every deprecation in this release](#every-deprecation-in-this-release)**, at the end of this section, lists each
  newly deprecated name and call form with its replacement.

### How to upgrade from 0.21.1

Work through these in order; each step names the part of this section that explains it.

1. **Upgrade and compile.** Very little was removed outright (see
   [Removed without a forward](#removed-without-a-forward)), so most code compiles unchanged. A compile error is
   one of: a name in that list, a field that is now optional (`PdfMetadata`, `PdfExtractedImage.alphaMask`,
   `PdfBookmark.pageIndex`), `new Pdf.EditorPage()` (pages come from an editor), a tar option that never applied
   (`password` / `encryptionMethod` on `ZipOptionsTar`), or an argument the stricter types now refuse (a value that
   was never valid).
2. **Find `null` checks that stopped matching.** TypeScript does not flag these. Search for `=== null` / `!== null`
   against the results listed under [Names and returns made uniform](#names-and-returns-made-uniform-before-10-breaking)
   and the PDF fields under [Read limits and PDF fields](#read-limits-pdf-fields-that-are-now-absent-and-other-late-fixes-breaking), and use
   `=== undefined` or `== null`.
3. **Check `catch` blocks.** Errors are now the module's own classes everywhere: argument mistakes are
   `<Module>ArgumentError` (with `field`), a refused file operation is `<Module>FileError`, a configured read limit is
   `<Module>LimitExceededError`. An Excel read past a container limit is `ExcelLimitExceededError` (the archive error is
   its `cause`). See [Every public function checks its arguments](#every-public-function-checks-its-arguments-behaviour-change)
   and [Contracts tightened across every module](#contracts-tightened-across-every-module-behaviour-change).
4. **Check `Buffer` use on Node.** `Workbook.toBuffer` resolves to a plain `Uint8Array`; `bytes.slice()` now copies.
5. **Review the behaviour changes for the modules you use** — Excel reads return a new workbook, references follow
   inserted/deleted rows and renamed sheets, facet getters return copies, values and notes are copied in, Word
   `Document.*` no longer reaches documents that share nodes, date handling, and the PDF export defaults. Each is a
   `(behaviour change)` heading below. One that tests trip on first: a six-digit `argb` (`"f56600"`) is now written
   as `FF` + rgb, as the file format requires, so it reads back as `"FFf56600"` — compare against the eight-digit form.
6. **Replace deprecated names.** Everything deprecated still works until 1.0, when all of it is removed together.
   Turn on a deprecation lint (`@typescript-eslint/no-deprecated`, or your editor's strike-through) and replace each hit
   with the name in [Every deprecation in this release](#every-deprecation-in-this-release) — that list is the
   complete one; the tables in the themed sections below explain the larger moves.
7. **If you run `@documonster/mcp`,** read [its section](#documonstermcp): no tool changed name or lost an argument,
   but `pdf_edit` now refuses to rebuild a signed PDF instead of invalidating its signature.

### The most common changes, before and after

Most code needs only these. Every "before" form below still compiles and runs until 1.0 (deprecated), so you can
migrate one call at a time.

**Excel: a read returns a new workbook.** The workbook-first form is deprecated.

```ts
import { Workbook } from "documonster/excel";

// Before
const old = Workbook.create();
await Workbook.read(old, bytes);

// After
const wb = await Workbook.read(bytes);
```

The same holds for `Workbook.readFile`, `readStream` and `readWithDiagnostics`, and for CSV and Markdown: `ExcelCsv.read(input)` / `ExcelMarkdown.read(input)` return a new workbook, and adding a sheet to an existing one is `ExcelCsv.addWorksheet(wb, input, options)` / `ExcelMarkdown.addWorksheet(wb, input, options)`. Both are async.

**Excel: a cell is one argument; settings go in an options object.**

```ts
import { Cell, Worksheet } from "documonster/excel";

// Before
Cell.setValue(ws, 2, 3, "x");
Worksheet.freeze(ws, 1, 1);
await Worksheet.protect(ws, "secret", { formatCells: true });

// After
Cell.setValue(ws, { row: 2, col: 3 }, "x"); // or "C2"
Worksheet.freeze(ws, { columns: 1, rows: 1 });
await Worksheet.protect(ws, { password: "secret", formatCells: true });
```

**Excel: one way to add a chart.** The 31 `Chart.add*` functions become `Chart.place`, with the placement before the
options and the chart family in `type`. It returns the chart's handle rather than a part number.

```ts
import { Chart } from "documonster/excel";

const series = [{ name: "Sales", categories: "Data!$A$2:$A$3", values: "Data!$B$2:$B$3" }];

// Before
Chart.addBar(ws, { series }, "E2:L16");

// After
const chart = Chart.place(ws, "E2:L16", { type: "bar", series });
```

**Errors carry their module's prefix; prefer the guard.** `TableError` is `ExcelTableError`, `WorksheetNameError` is
`ExcelWorksheetNameError`, and so on (the old names are the same classes, deprecated). Argument mistakes are
`<Module>ArgumentError` with a `field`.

```ts
import { Table, isExcelError, ExcelTableError } from "documonster/excel";

try {
  Table.add(ws, {
    name: "Sales",
    ref: "A1",
    headerRow: true,
    columns: [{ name: "a" }],
    rows: [[1]]
  });
} catch (error) {
  if (error instanceof ExcelTableError) {
    // the table was refused; nothing was changed
  } else if (isExcelError(error)) {
    // any other failure of this module
  }
}
```

**Word: reading and writing live on `Document`, and so do in-place edits.** `Io.*` and the editing members of `Query`
are deprecated forwards.

```ts
import { Io, Query } from "documonster/word";

// Before (deprecated, still works until 1.0)
const before = await Io.read(bytes);
Query.replaceText(before, "NAME", "World");
const outBefore = await Io.toBuffer(before);
```

```ts
import { Document } from "documonster/word";

// After
const doc = await Document.read(bytes);
Document.replaceText(doc, "NAME", "World");
const out = await Document.toBuffer(doc);
```

**PDF: pulling content out of a file is `Pdf.extract`** (`Pdf.read` is its deprecated name).

```ts
import { Pdf } from "documonster/pdf";

const { pages } = await Pdf.extract(pdfBytes);
```

**Absence is `undefined`.** Readers that returned `null` for "not there" now return `undefined`; the list is under
[Names and returns made uniform](#names-and-returns-made-uniform-before-10-breaking). `== null` matches both, so it is
the safest rewrite of a `=== null` test.

### Done? Check before you ship

- `tsc` is clean, and a deprecation lint (`@typescript-eslint/no-deprecated`) reports nothing — or only names you
  have chosen to migrate later, before 1.0.
- No `=== null` / `!== null` test remains against a result listed in
  [Names and returns made uniform](#names-and-returns-made-uniform-before-10-breaking).
- No `catch` tests for a class the error no longer is (an archive limit class around an Excel read, a
  `PdfStructureError` for a bad page index); prefer the module's guard (`isExcelError`, `isDocxError`, …).
- No `Buffer`-only method is called on bytes this library returns; `slice()` copies, `subarray()` is the view.
- No test or comparison expects a six-digit `argb` back from a written file; it reads back as `FF` + rgb.
- A real file of yours reads, edits and writes back, and opens in Excel or Word.

### Find what applies to you

| You use                                   | Read                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| ----------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Every module                              | [Names and returns made uniform](#names-and-returns-made-uniform-before-10-breaking), [Every public function checks its arguments](#every-public-function-checks-its-arguments-behaviour-change), [Removed without a forward](#removed-without-a-forward)                                                                                                                                                                                                                                                                                                                                                  |
| Excel                                     | [API shape](#api-shape-one-place-to-read-and-write-one-vocabulary-for-streaming-one-address-argument), [Excel: values, notes and models](#excel-values-notes-and-models-are-the-callers-own-behaviour-change-one-note-api), [references follow edits](#excel-inserting-deleting-and-renaming-keep-references-consistent-behaviour-change), [Excel — document model](#excel--document-model), [Excel — tables, charts, …](#excel--tables-charts-images-pivots-forms-streaming-csvmarkdown-bridges), [One vocabulary across modules](#public-api-one-vocabulary-across-modules-deprecations-nothing-removed) |
| Excel streaming (`Stream.WorkbookWriter`) | [API shape](#api-shape-one-place-to-read-and-write-one-vocabulary-for-streaming-one-address-argument)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| Word                                      | [API shape](#api-shape-one-place-to-read-and-write-one-vocabulary-for-streaming-one-address-argument), [Word](#word)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| PDF, Draw, Mermaid                        | [Read limits and PDF fields](#read-limits-pdf-fields-that-are-now-absent-and-other-late-fixes-breaking), [PDF, drawing and Mermaid](#pdf-drawing-and-mermaid)                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Archive, streams                          | [Archive and streams](#archive-and-streams)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| CSV, XML, Markdown, Formula               | [Option checks](#option-checks-are-declared-once-with-one-wording-behaviour-change), [CSV, XML, Markdown and Formula](#csv-xml-markdown-and-formula)                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| Encrypted Office files                    | [Excel and Word: encrypted files](#excel-and-word-encrypted-files-are-read-and-written-new-one-implementation)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| `@documonster/mcp`                        | [`@documonster/mcp`](#documonstermcp)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| A specific old name                       | Search [Every deprecation in this release](#every-deprecation-in-this-release), or the per-name [API diff](docs/api-changes-since-v0.21.1.md)                                                                                                                                                                                                                                                                                                                                                                                                                                                              |

At a glance, per entry — the counts of the rows in [docs/api-changes-since-v0.21.1.md](docs/api-changes-since-v0.21.1.md), where every one of them is listed with its signature:

<!-- api-changes:summary:start -->

| Entry                                                                                              | Removed | Became optional | Changed | Deprecated | Deprecated call forms | Added to existing | Added (new exports and their members) |
| -------------------------------------------------------------------------------------------------- | ------: | --------------: | ------: | ---------: | --------------------: | ----------------: | ------------------------------------: |
| [`documonster/archive`](docs/api-changes-since-v0.21.1.md#documonsterarchive)                      |       0 |               0 |     110 |        170 |                    12 |               165 |                                   327 |
| [`documonster/archive [browser]`](docs/api-changes-since-v0.21.1.md#documonsterarchive-browser)    |       0 |               0 |      12 |          6 |                     0 |                 8 |                                    13 |
| [`documonster/chart`](docs/api-changes-since-v0.21.1.md#documonsterchart)                          |       0 |               0 |       8 |         52 |                     0 |                12 |                                   104 |
| [`documonster/csv`](docs/api-changes-since-v0.21.1.md#documonstercsv)                              |       6 |               0 |      23 |         47 |                     5 |                10 |                                   256 |
| [`documonster/draw`](docs/api-changes-since-v0.21.1.md#documonsterdraw)                            |       0 |               0 |      13 |         45 |                     6 |                 3 |                                   181 |
| [`documonster/excel`](docs/api-changes-since-v0.21.1.md#documonsterexcel)                          |       1 |               0 |     185 |        281 |                    58 |               322 |                                   269 |
| [`documonster/excel/csv`](docs/api-changes-since-v0.21.1.md#documonsterexcelcsv)                   |       0 |               0 |       9 |         12 |                     0 |                 1 |                                   193 |
| [`documonster/excel/csv [browser]`](docs/api-changes-since-v0.21.1.md#documonsterexcelcsv-browser) |       0 |               0 |       1 |          1 |                     0 |                 0 |                                     2 |
| [`documonster/excel/formula`](docs/api-changes-since-v0.21.1.md#documonsterexcelformula)           |       0 |               0 |       0 |          1 |                     0 |                 1 |                                    18 |
| [`documonster/excel/markdown`](docs/api-changes-since-v0.21.1.md#documonsterexcelmarkdown)         |       0 |               0 |       3 |          7 |                     0 |                 0 |                                   108 |
| [`documonster/excel [browser]`](docs/api-changes-since-v0.21.1.md#documonsterexcel-browser)        |       0 |               0 |       3 |          0 |                     1 |                 1 |                                     0 |
| [`documonster/formula`](docs/api-changes-since-v0.21.1.md#documonsterformula)                      |       0 |               0 |       1 |          3 |                     0 |                 1 |                                   201 |
| [`documonster/markdown`](docs/api-changes-since-v0.21.1.md#documonstermarkdown)                    |       0 |               0 |       3 |          1 |                     0 |                 0 |                                     9 |
| [`documonster/mermaid`](docs/api-changes-since-v0.21.1.md#documonstermermaid)                      |       0 |               0 |       1 |         87 |                     0 |                24 |                                   487 |
| [`documonster/pdf`](docs/api-changes-since-v0.21.1.md#documonsterpdf)                              |       0 |               6 |      51 |         64 |                     6 |                47 |                                   509 |
| [`documonster/stream`](docs/api-changes-since-v0.21.1.md#documonsterstream)                        |       0 |               0 |      33 |         61 |                     2 |                 3 |                                   259 |
| [`documonster/stream [browser]`](docs/api-changes-since-v0.21.1.md#documonsterstream-browser)      |       0 |               0 |       0 |          0 |                     0 |                 0 |                                    44 |
| [`documonster/word`](docs/api-changes-since-v0.21.1.md#documonsterword)                            |       2 |               0 |     142 |        248 |                    31 |               347 |                                  1115 |
| [`documonster/word/crypto`](docs/api-changes-since-v0.21.1.md#documonsterwordcrypto)               |       0 |               0 |       9 |         22 |                     0 |                 2 |                                    87 |
| [`documonster/word/excel`](docs/api-changes-since-v0.21.1.md#documonsterwordexcel)                 |       0 |               0 |       3 |          7 |                     1 |                 0 |                                    66 |
| [`documonster/word/html`](docs/api-changes-since-v0.21.1.md#documonsterwordhtml)                   |       0 |               0 |       0 |          7 |                     0 |                 0 |                                    56 |
| [`documonster/word/markdown`](docs/api-changes-since-v0.21.1.md#documonsterwordmarkdown)           |       0 |               0 |       0 |          7 |                     0 |                 0 |                                    53 |
| [`documonster/xml`](docs/api-changes-since-v0.21.1.md#documonsterxml)                              |       0 |               0 |       1 |         19 |                     0 |                10 |                                    71 |
| [`@documonster/mcp`](docs/api-changes-since-v0.21.1.md#documonstermcp)                             |       0 |               0 |       0 |          0 |                     0 |                 0 |                                     3 |

<!-- api-changes:summary:end -->

### Names and returns made uniform before 1.0 (breaking)

**`Workbook.toBuffer` resolves to a `Uint8Array` on Node, as everywhere else.** It used to resolve to a `Buffer`
there. Node's own sinks (`fs.writeFile`, `res.end`, `Readable.from`) take the bytes unchanged, and the value is a
view of the same memory, so nothing is copied. A call to a `Buffer`-only method is now a type error; wrap the bytes
without copying where you need one: `Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)`. Note that
`bytes.slice()` now copies, as `Uint8Array#slice` does; `subarray` is the view.

**Absence is `undefined`, not `null`.** These now return `undefined` where they returned `null`:
`ZipReader#get` / `bytes`, `TarReader#get` / `bytes`, `ArchiveFile#getEntry` / `readEntry` / `readEntrySync` /
`readAsText` / `readAsTextSync` / `getZipEntryComment` / `sourcePath`, `RemoteZipReader#extract` /
`getRawCompressedData` / `getRawCompressedStream` / `getRawEntry` / `getRawEntryStream` / `verifyCrc32`,
`Fs.safeStats` / `safeStatsSync`, and Excel's `Watermark.get`. **TypeScript will not flag a strict `=== null` check
against these results**, and it now never matches — search your code for `=== null` and `!== null` on them and use
`=== undefined`, `== null` or a truthiness test. `RemoteZipReader#checkPassword` keeps its three-way `null`, and the
Node-compatible stream API keeps Node's.

**`Chart.copyTo` returns the copy's handle.** It returned the copy's part number; that number is now
`handle.chartNumber` (`chartExNumber` for an Office 2016 chart). The deprecated `Chart.clone` still returns the
number.

**Incremental readers take the `Stream` suffix.** `Csv.parseRows` → `Csv.parseStream`, `Xml.parseSax` →
`Xml.parseSaxStream`. The old names are deprecated forwards until 1.0; argument errors now name the new call.

**Smaller startup bundles.** With a code-splitting bundler (rolldown, rspack, Vite, webpack), a write-only Excel
export loads about 451 KB minified up front, down from 484 KB, and a Word one 319 KB, down from 372 KB. The ZIP AES
cipher, the hash and block-cipher library, the compression worker pool, TAR, Word's chart serialiser and its font
metrics now load only when a package needs them.

**`Archive.randomBytes()` without a length throws `ArchiveArgumentError`.** It used to return an empty array in a
browser and throw Node's `TypeError` on Node. In a browser it also accepts lengths above 65,536 now; it threw Web
Crypto's `QuotaExceededError` for them.

### Read limits, PDF fields that are now absent, and other late fixes (breaking)

**Every configured read limit is a `…LimitExceededError` with `limit` and `allowed`.**

- `ArchiveFileTooLargeError` is now a subclass of `ArchiveLimitExceededError`, and `ArchiveLimit` gains
  `"maxEntrySize"`. Catching `ArchiveLimitExceededError` sees every archive limit; a `catch` of
  `ArchiveFileTooLargeError` is unchanged.
- An Excel read past a container limit (`zip.maxEntries`, `zip.maxTotalUncompressedSize`, `maxEntrySize`) throws
  `ExcelLimitExceededError` — so `isExcelError` sees it — with the archive's error as `cause`. It threw the
  `ArchiveLimitExceededError` / `ArchiveFileTooLargeError` itself. **A `catch` testing for the archive class no longer
  matches; test for `ExcelLimitExceededError`, or read `error.cause`.**
- `ExcelLimitExceededError` gains `limit` (`"maxRows"`, `"maxCols"` or the `zip` option) and `allowed`; `maxItems` is a
  deprecated alias of `allowed`. New types `ExcelLimit`, `ExcelLimitExceededErrorOptions`.
- `DocxLimitExceededError` gains `allowed`; `maximum` is a deprecated alias of it.
- New `CsvLimitExceededError` (`limit: "maxRowBytes"`) for an oversized row, and `XmlLimitExceededError` for `maxDepth`
  / `maxEntityExpansions` (`limit: "depth" | "entityExpansions"`, plus `allowed`). Each subclasses the class thrown
  before (`CsvError`, `XmlParseError`), so existing `catch` blocks still match.

**Other errors.**

- New `PdfUnsupportedFeatureError` (a `PdfStructureError`) for an encryption filter, revision or crypt-filter method
  the reader has no handler for.
- `Streams.LowLevel.SinkAdapter#write` after `end()` rejects with `StreamStateError`, not a bare `Error`.
- `DefinedNames.addFormula` with an `expression` that is not a string throws `ExcelArgumentError` (`field:
"expression"`); an empty one is still `ExcelDefinedNameError`.
- `Archive.ZipEditPlan.deserialize` with a `version` other than 1 throws `ArchiveArgumentError` (`field:
"data.version"`), not a bare `ArchiveError`.
- `DocxEncryptedError` takes `{ cause }`.

**PDF: absence is `undefined` (breaking).** `PdfMetadata.creationDate`, `modDate`, `pageSize` and `xmpXml`,
`PdfExtractedImage.alphaMask`, and `PdfBookmark.pageIndex` are optional fields: absent where they held `null` (or, for
`pageIndex`, `-1`). **TypeScript will not flag `=== null` or `=== -1` against them**; test with `=== undefined`. The
text fields of `PdfMetadata` and `PdfAnnotation` keep `""` for a missing entry, now documented: a PDF can carry an entry
that is present and empty, and the reader cannot always tell the two apart.

**Behaviour fixes.**

- **A workbook read's options now hold on every path.** `zip.maxEntries`, `zip.maxTotalUncompressedSize`,
  `zip.maxEntrySize`, `zip.checkCrc32`, `maxRows` and `maxCols` were applied by the buffered XLSX reader only. An XLSB
  `read` / `readWithDiagnostics` / `readFile` accepted them and read the package anyway (`checkCrc32: false` was
  ignored too), and a streamed XLSX `readStream` on Node applied no container limit at all; each now refuses with
  `ExcelLimitExceededError` as the buffered XLSX reader does. The XLSB reader counts `maxRows` / `maxCols` as the XLSX
  one does: row records per sheet, cell records per row. Format detection applies `zip.maxEntries` too, before it walks
  the directory. **A file these used to accept past a limit is now refused.**
- **`Workbook.readStream` in a browser takes every input it declares** — a Web `ReadableStream` and an (async) iterable
  failed with a raw `TypeError` — and decrypts with `{ password }`, which it ignored.
- **An XLSB write applies `zip`** (`level`, `modTime`, `timestamps`), for `toBuffer`, `toStream` and `writeStream`; it
  ignored them. An unchanged package read from XLSB still comes back byte for byte, unless the write gives `zip`
  settings, which the original bytes cannot honour.
- `PdfMetadata.creationDate` / `modDate` are absent for a date the lenient XMP scan cannot read; that path handed out an
  `Invalid Date`.
- **`Column.setKey` on a streaming writer stays open until the sheet is committed.** A key is not written into the file
  — it only says which property of a row object `addRow` reads — so refusing it once the first row was committed, as
  the written column settings are, stopped later rows from being read by a new key.
- **Word: the copy a shared document makes before a tree edit is structural.** After a derivation (`Security.protect`,
  `Transform.*`, …) the first tree edit of either document (`replaceText`, accepting revisions, …) copies its nodes;
  that copy used `structuredClone`, so it duplicated every image and font and failed with the platform's
  `DataCloneError` on a value it cannot clone. It now shares bytes and anything that is not plain data. As before, a
  node taken from the document through `Query.*` before that edit belongs to the copy's predecessor afterwards.
- **Argument messages read alike across modules.** A boolean is `must be a boolean` everywhere (CSV and some Excel
  options said `true or false`), and a whole-number range reads `an integer from 1 to 5` (Excel said
  `a whole number from 1 to 5.`). A CSV `columns` entry that is an object names the field at fault
  (`options.columns[0].key`) rather than the whole entry. Match on `field`, not on the message.

- `RemoteZipReader#findEntries` matches by the shared glob rules `Fs.matchGlob` uses. `**/x` now includes a root-level
  `x`, `?` no longer matches `/`, `[abc]` and `{a,b}` work, and letters match without regard to case.
- `Compression.Workers.getDefaultPool(options)` refuses `options` with `ArchiveArgumentError` when the default pool
  already exists; it used to ignore them. The deprecated flat `getDefaultWorkerPool` keeps the old behaviour.
- `BasicRasterCanvas#getUncoveredCodePoints()` returns a sorted copy. The `uncoveredCodePoints` field is deprecated and
  now returns a copy too: clearing it no longer changes what the render reports.

**Renames (old names are deprecated forwards until 1.0).**

| Old                                                     | New                                                               |
| ------------------------------------------------------- | ----------------------------------------------------------------- |
| `PdfRect` (`documonster/pdf`)                           | `PdfAnnotationRect`                                               |
| `Archive.ZipEditor#getEntries`                          | `getSourceEntries` (the source's entries, ignoring pending edits) |
| `PositionErrorOptions`, `InvalidValueTypeErrorOptions`  | `ExcelPositionErrorOptions`, `ExcelInvalidValueTypeErrorOptions`  |
| `Compression.Workers.WorkerPool` (unreleased)           | `Compression.Workers.Pool`                                        |
| `UnrecognizedValue`, `DiscardUnrecognizedValuesOptions` | `DocxUnrecognizedValue`, `DocxDiscardUnrecognizedValuesOptions`   |

The last two rows were never released, so they have no forwards. `DocxLegacyTemplateOptions` is marked deprecated
with the template functions it serves. `DocxCompiledTemplate._doc` and `PdfContentStream#deferred` / `#preflight` are
deprecated: they are implementation details and leave the public types in 1.0.

**Limits and isolation, tightened at the edges.**

- `Workbook.readStream` charges every byte of a part to the `zip` limits, including parts it skips without parsing,
  as `Workbook.read` does; a limit crossed while draining such a part is now reported instead of ignored.
- A ZIP `maxTotalUncompressedSize` never widens the per-entry bound: an entry is still held to `maxEntrySize`, or its
  512 MiB default when unset.
- A Word document's nodes built in another realm (an iframe, `node:vm`) are copied like any other when a shared
  document is edited, so the edit no longer reaches a document derived from it.
- `Xml.parseSaxStream` (and its deprecated `Xml.parseSax`) checks its arguments when called, not on the first
  `next()`: a wrong argument now throws synchronously, so a `try` around only the `for await` no longer catches it.

### API shape: one place to read and write, one vocabulary for streaming, one address argument

**Word: `Document.*` no longer reaches a document that shares nodes with the one it edits (bug fix).** A function
that returns a new document (`Template.fill`, `Transform.*`, `Security.protect` / `unprotect`, `Vba.add` / `remove`,
`Font.addEmbedded`, `Ole.add`) may share unchanged arrays and nodes with its input, and `Document.*` used to edit
those in place — `Document.addParagraph(await Security.protect(a, …), "x")` appended to `a.body` too. `Document.*`
now copies whatever two such documents share before changing it (once; later edits are in place again), in both
directions. `Document.clone` is no longer needed to keep them apart; keep it for a snapshot or before editing nodes
by hand.

**Word: reading and writing moved to `Document`.** `Io.read`, `Io.toBuffer`, `Io.toBase64`, `Io.toFlatOpc` and, on
Node, `Io.readFile` / `Io.writeFile` are now `Document.read`, `Document.toBuffer`, `Document.toBase64`,
`Document.toFlatOpc`, `Document.readFile`, `Document.writeFile`, beside `Document.create` — as Excel has
`Workbook.create` / `read` / `toBuffer`. The `Io` names are deprecated forwards to the same functions; their argument
errors now start with `Document.…:`.

**Excel: a cell is one argument.** `Cell.*` takes `"A1"` or a `{ row, col }` position (`CellPosition`, now exported).
The `(ws, row, col, …)` form is deprecated — `Cell.setValue(ws, 2, 3, v)` → `Cell.setValue(ws, { row: 2, col: 3 }, v)`
— and keeps its behaviour until 1.0. The forms are told apart by the type of the second argument, not by argument
count. A position outside the grid is `ExcelInvalidAddressError` with `field` `addr.row` / `addr.col`.

**Excel: writing a range.** New `Range.setValues(ws, range, values)` writes a row-major matrix from the range's
top-left cell (rows may be shorter; larger is refused; checked in full before anything is written), and
`Range.setStyle`, `setFont`, `setFill`, `setBorder`, `setAlignment`, `setNumFmt`, `setProtection` apply a style to every
cell of a range with the matching `Cell` setter's semantics.

**Excel: a streaming writer takes the document model's functions.**

- `Workbook.addWorksheet` / `getWorksheet` / `listWorksheets` / `eachWorksheet` / `protect` / `unprotect` /
  `verifyPassword` / `getDefinedNames` / `addCellStyle` / `getCellStyle` / `listCellStyles` / `useBuiltInCellStyle`
  accept a `Stream.WorkbookWriter`; `Worksheet.merge` / `mergeWithoutStyle` / `protect` / `unprotect`, `Column`
  settings, `Cell.*`, `Row.*`, `Range.*`, `Image.add` / `place` / `list` / `remove` / `setBackground` /
  `getBackground` / `removeBackground`, `Watermark.*`, `ConditionalFormatting.*` and `DataValidation.*` accept a
  `Stream.WorksheetWriter`, addressing what the writer still holds and applying its own checks.
- A committed workbook or sheet refuses an edit with `ExcelStreamStateError`; so does a row the writer has written,
  for an edit **and** a read (`Row.setHeight` and the other `Row` setters threw `ExcelRowOutOfBoundsError` before), and
  a column setting once the columns are written with the first committed row. A range write that reaches a written
  row is refused whole.
- What needs rows already written or the whole document refuses a writer with `ExcelUnsupportedFeatureError`:
  `Worksheet.unmerge`, the `Image` shape functions, `Row.insert` / `insertMany` / `duplicate` / `splice` /
  `getCount` / `getActualCount` / `getLast` / `each` / `autoFit`, `Column.getValues` / `setValues` / `splice` /
  `each` / `eachAddress` / `has` / `getCount` / `getLast` / `getActualCount` / `autoFit`, `Workbook.removeWorksheet`,
  chartsheets, `importWorksheet`, external links, persons, registered functions, `getModel` / `setModel`,
  `removeCellStyle`, and `Workbook.toBuffer` / `toStream` / `writeFile` / `writeStream` (a writer finishes with its own
  `commit()` / `toBuffer()`).
- `Workbook.getWorksheet(writer)` with no id returns the first sheet; the deprecated `WorkbookWriter#getWorksheet()`
  still returns `undefined`.
- Deprecated, removed in 1.0: `WorksheetWriter#mergeCells` → `Worksheet.merge`, `#addImage` → `Image.place`,
  `#getImages` → `Image.list`, `#addBackgroundImage` → `Image.setBackground`, `#getBackgroundImageId` →
  `Image.getBackground` (the id as a string), `#addWatermark` / `#getWatermark` / `#removeWatermark` →
  `Watermark.set` / `get` / `remove`, `#addConditionalFormatting` / `#removeConditionalFormatting` →
  `ConditionalFormatting.add` / `remove` / `clear`, `#protect` / `#unprotect` → `Worksheet.protect` / `unprotect`;
  `WorkbookWriter#addImage` → `Image.add`, `#addWorksheet` / `#getWorksheet` → `Workbook.addWorksheet` /
  `getWorksheet`, `#protect` / `#unprotect` → `Workbook.protect` / `unprotect`, `#addCellStyle` /
  `#defineCellStyle` → `Workbook.addCellStyle`, `#useBuiltInCellStyle` → `Workbook.useBuiltInCellStyle`,
  `#definedNames` → `Workbook.getDefinedNames`.
- `Stream.setCell*` / `Stream.setRow*` / `Stream.commitRow` stay: they take the handle `addRow` returns, where
  `Cell` / `Row` / `Range` take a position.

### Option checks are declared once, with one wording (behaviour change)

Every module's options are now checked by one declared schema per options interface, so a wrong option reads the same
everywhere: `<call>: <field> must be <expected>, got <value>`, where the value is described (`got 5`, `got an object`,
`got NaN`) rather than given as its `typeof`, and there is no trailing period. Code that matched the old texts must be
updated (`true or false` is now `a boolean`, `a whole number` is `an integer`, `of at least N` is `>= N`; Mermaid
messages no longer end in `; got X.`) — better, match on the error's class and `field` rather than its message. `field` paths are unchanged
except where listed. Options that were accepted unchecked are now refused:

- **CSV.** `columnMismatch` must be an object (an array was accepted); `Csv.format` / `FormatterStream` check
  `firstRowIsHeader`; `Csv.generate*` check `headers`, `bom` and `objectMode`; `Csv.detectLinebreak` checks `quote`;
  `Csv.applyDynamicTypingToRow` refuses an array as `row`. A line break in `delimitersToGuess[i]` is reported at
  `options.delimitersToGuess[i]` (was `options.delimiter`); a bad `typeTransform.row` result at
  `options.typeTransform.row` (was `… result for data[i]`).
- **Markdown.** A `columns[i]` that is an array is refused at `options.columns[i]` (was `.header`).
- **Archive, Stream.** An array is no longer accepted as an options object, nor — by `Archive.createZip` and
  `Archive.packZip`, which let them through — `null` or a primitive. Archive checks `signal`, `onProgress`
  and `progressIntervalMs` where given; `RemoteZipReader` checks `password`, `decodeStrings`, `checkCrc32` and
  `signal`; `ZipArchive#add` refuses a `name` that is not a non-empty string. `StreamArgumentError.actualType` carries
  the described value.
  `Archive.createZip`, `Archive.packZip` and `Archive.packZipSync` now take `encryptionMethod` and `password`: they
  had neither, and an unknown option is not an error, so an archive asked to be encrypted was written in the clear.
  A `password` without an `encryptionMethod` — which would encrypt nothing — and a method without a password are
  refused (`field: "options.password"`); AES is written by the asynchronous builders only, as it runs on Web Crypto.
- **PDF, Mermaid.** A non-string entry of `lines` or `palette` is reported at its index (`lines[1]`,
  `options.palette[1]`); an array given as a Mermaid `box` / `style` or a PDF watermark `position` is reported at
  that field. `new Pdf.Builder` no longer checks `options.onWarning`, which it never had — use
  `Pdf.Builder#onWarning(handler)`.
- **Word.** The `include*` flags of `Transform.map`, `Query.each` and `Query.listTables` (`topLevelOnly`),
  `Convert.toSemantic`'s options, every option of `WordHtml.fromDocx` and `WordMarkdown.fromDocx`,
  `Validation.validate`'s `maxSeverity` / `maxErrors` / `compatibilityMode` / `strict`, and `Font.createHyphenator`'s
  `hyphenChar` are checked; a non-string `WordHtml.fromDocx` `title` or `classPrefix` is a `DocxArgumentError`, not a
  `TypeError`. A `WordHtml.toDocx` `options.classStyles` entry is reported as `options.classStyles["name"]`.
  `Document.toFlatOpc` no longer checks `modTime`, which it never used. `Build.mathFraction`, `mathLimit` and
  `mathAccent` take their setting in options (`{ fractionType }`, `{ limitType }`, `{ char }`), as `mathBar` does;
  for all four, a value that is not an options object is the deprecated positional setting, so `mathBar(c, 5)` is
  refused at `position` (was `options`).
- **Excel.** `Worksheet.protect` / `Workbook.protect` report a flag at `options.<flag>` and the password at
  `options.password` (was the bare key). `Stream.WorkbookWriter` checks every option — dates, booleans, `format`,
  `stream`, `filename`, and `zip.zlib.level` / `zip.compressionOptions.level` (0–9; it clamped silently).
  `WorkbookReader` reports `maxEntrySize` as `ExcelArgumentError` (was a bare `RangeError`) and checks
  `maxBufferedWorksheetBytes` and `checkCrc32`. `Workbook.addChartsheet` / `addPivotChartsheet` check their options as
  the `Chartsheet` setters do. `Column.autoFit({ start, end })` refuses a column outside A–XFD with
  `ExcelColumnOutOfBoundsError` (a huge `end` looped until it got there). `Workbook.toStream` checks
  `highWaterMark`; `Worksheet.toJson` checks `raw`, `blankRows` and `dateFormat`; `ExcelCsv` reads check
  `decimalSeparator`, `requestHeaders` and `signal`; `Sparkline.add` refuses an entry with no `cellRef`.
  `Workbook.toBuffer` / `writeStream` / `writeFile` accept `null` options as none (they threw a `TypeError`).
  A data-validation rule of `type: "any"` no longer checks `operator` / `formulae`, which neither writer writes
  for it; `DataValidationBase.showDropDown` is now in the published type (it was checked and written, but
  untyped — note that `true` hides the in-cell list).

When one call has several wrong options, which is reported first may differ from before.

### Every public function checks its arguments (behaviour change)

A wrong argument to a public function is now the module's own `…ArgumentError`, naming the parameter as the
signature does and starting its message with the call. A sweep over every namespace member of every entry found
about seventy that let a raw `TypeError` (or Node's `ERR_INVALID_ARG_TYPE`) escape instead, and four that accepted
`{}` as a worksheet and wrote onto it. All of them now report the module's own argument error.
Valid arguments behave exactly as before. A deprecated name keeps the behaviour it was published with.

- **Excel.** `Workbook.toBuffer`, `toStream`, `writeStream` and `writeFile` check the workbook first
  (`field: "wb"`). `Worksheet.protect`, `Worksheet.freeze` and `Worksheet.split` refuse something that is not a
  worksheet instead of writing onto it. `new Stream.WorkbookWriter(options)` and
  `new Stream.WorkbookReader(input, options)` refuse an `options` that is not an object (`WorkbookWriter(42)` was
  accepted), and the reader names a wrong option by its path (`options.password`, `options.worksheets`; it said
  `password`). `ChartClassic.fillNumRef` / `fillStrRef` / `fillCaches` / `toColorsXml` / `toStyleXml` /
  `fromColorsXml` / `fromStyleXml` / `createOptionsFrom*`, `ChartEx.toXml` / `fillCaches` / `createOptionsFrom*`
  and the `DrawingMl` readers and `create*` builders check their arguments.
- **Word.** `Font.createHyphenator` checks the pattern set, not only its type: `patterns.patterns` and
  `patterns.exceptions` must be arrays of strings (`field: "patterns.patterns[3]"`), `options.minLeft` /
  `options.minRight` integers ≥ 0 (a `NaN` quietly disabled hyphenation). `Font.shapeText`, `Font.detectScript`,
  `Font.detectDirection`, `Template.fill`, `Package.patchTemplate` and `WordCrypto.Agile.parseEncryptionInfo`
  throw `DocxArgumentError`. `WordExcel.toDocx` refuses a value that is not a workbook with `DocxArgumentError`
  (`field: "workbook"`) rather than Excel's `ExcelArgumentError`.
- **Draw and PDF.** `Draw.Geometry.roundedRectToPath` names its own parameters (`x`, `y`, `width`, `height`; it
  reported `box.x`, a path no caller wrote). `Draw.Backend.SvgSurface`, `Draw.Backend.BasicRasterCanvas` and the
  surfaces `Draw.Backend.createRasterSurface` and `Pdf.createDrawSurface` return throw `DrawArgumentError` /
  `PdfArgumentError` naming the call (`Draw.Backend.BasicRasterCanvas#fillRect: x must be a number, got
Symbol(x)`); a `NaN` coordinate still draws nothing on the SVG and raster surfaces, and the PDF surface refuses it
  under its own name rather than `Pdf.PageBuilder#drawRect`'s. `Draw.Backend.normalizeSamples` names itself.
  `Pdf.Editor#splitPages` refuses a value that is neither options nor a list of indices (`field: "options"`).
- **`BasicRasterCanvas` no longer hangs** on a far or `NaN` line end, or a large circle or ellipse: `drawLine`
  clips to the canvas and `fillCircle` / `fillEllipse` stop at its edge. `Draw.toRgba` of a line reaching `1e9`
  took 18 s, and of an ellipse of radius `1e5` did not finish.
- **Archive.** `Archive.toAsyncIterable` checks its arguments when it is called, not on the first `next()`
  (`field: "source"` / `"options"` / `"options.signal"` / `"options.onChunk"`); `toReadableStream` and
  `toNodeReadable` check `options` the same way, and a source chunk that is not bytes or a string is an
  `ArchiveArgumentError` (`field: "source"`). `Archive.Crypto.encryptZipCrypto`, `decryptZipCrypto`,
  `createZipCryptoKeys`, `encryptAes` and `decryptAes` check their data, password, `crc32`, `lastModTime` and
  `getRandomBytes` (which must return at least 11 bytes — a short one produced wrong ciphertext);
  `Archive.Crypto.getEncryptionMethodName` throws for a method it has no name for (it returned `undefined`, or a
  function for `"toString"`; the deprecated flat `Archive.getEncryptionMethodName` still answers `undefined`), and
  `getEncryptionMethodForAesKeyStrength` for a strength other than 128, 192 or 256. `new Archive.ZipEditPlan(ops)`
  checks every operation (`field: "ops[0].name"`) and needs an array; a plan's `delete`, `deleteDirectory` and
  `rename` check their names, and `ZipEditPlan.deserialize` reports `data.ops[i]`.
  `new Archive.ArchiveFile(null)` throws `ArchiveArgumentError` (`field: "options"`).
- **`Fs.*` check their arguments**: a wrong path, option, `data`, `content`, `encoding`, `mtime` or `prefix` is an
  `ArchiveArgumentError` (`Fs.readFileBytes: filePath must be …`); what the file system refuses is still passed on
  as it is. `Fs` does not publish `createReadStream` / `createWriteStream` (they only forwarded to `node:fs`; use
  those).
- **CSV.** `Csv.applyDynamicTyping` requires a string `value` and a boolean or function `columnConfig`;
  `Csv.applyDynamicTypingToRow` requires `dynamicTyping` and string values (`field: 'row["a"]'`), checked before the
  row is changed; `new Csv.FormatterStream(null)` throws `CsvArgumentError`.
- **Streams.** `Streams.fromWebReadable` / `fromWebWritable` / `toWebReadable` throw `StreamArgumentError` for
  options that are not an object or carry a wrong `objectMode`, `highWaterMark`, `encoding`, `decodeStrings`,
  `signal` or `strategy`, and leave the Web stream unlocked; `new Streams.PullStream`, `BufferedStream`,
  `ChunkedBuilder` and `TransactionalChunkedBuilder` check their `options` (and `chunkSize` / `initialCapacity`).

**Settings move into options objects** (the positional forms of published names are deprecated overloads with the
same results):

| Deprecated                                                           | Use                                                                                  |
| -------------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| `Font.hyphenateWord(word, hyphenator, "-")`, `hyphenateText(…, "-")` | `Font.hyphenateWord(word, hyphenator, { hyphenChar: "-" })` (`DocxHyphenateOptions`) |
| `HyphenationOptions.hyphenChar`                                      | none — `Font.createHyphenator` never read it                                         |
| `flattenPath(commands, scale)` (flat)                                | `Draw.Geometry.flattenPath(commands, { scale })` (`DrawFlattenPathOptions`)          |
| `parseRasterFont(data, index)` (flat)                                | `Draw.parseRasterFont(data)` or `Draw.parseRasterFont({ data, collectionIndex })`    |

**Excel's streaming writers stop advertising their internals.** These fields were only ever used by the library
itself and are deprecated (internal; removed in 1.0): `WorkbookWriter.sharedStrings` / `styles` / `zipOptions` /
`xlsbTables` / `compressionLevel` / `media` / `commentRefs` / `dynamicArrayCount` / `zip` / `stream` / `promise`, and
`WorksheetWriter._retainedComments` / `tables` / `pivotTables` / `threadedComments` / `formControls` / `_charts` /
`_shapes` / `_sparklineGroups` / `_worksheetMcIgnorable` / `_sortStateAutoFilterRef` / `orderNo` / `sheetNo` /
`conditionalFormattings` / `hasComments` / `startedData` / `rId`. `WorksheetWriter.xlsbUnsupported` is deprecated in
favour of `unsupported`, as on `WorkbookWriter`.

- **`WorksheetWriter#workbook` is typed as the `Stream.WorkbookWriter` that made the sheet.** It was typed as the
  internal protocol the sheet uses — shared strings, styles, `_openStream`, `packageSink` — so `ws.workbook` offered
  those and not `addWorksheet` or `commit`. The object is the same; only its declared type changed.
- **`new Stream.WorkbookReader(path)` reports a file the system will not open as `ExcelFileError`** (`path`,
  `operation: "read"`, the system error as `cause`), as `Workbook.readFile` does. It was the platform's own `ENOENT`
  error, which `isExcelError` did not recognise.
- **`new Archive.UnzipEntry(…)` is deprecated.** Entries come from a reader (`ZipReader.entries()` / `get()`, or
  iterating `Archive.openZip(…)`); the shapes the constructor takes are the reader's own. The class stays, for
  `instanceof`.

### Argument errors and names aligned across modules (deprecations)

Every entry point was checked against one set of naming and return conventions. As elsewhere, **nothing is removed**:
every renamed name stays as a `@deprecated` forward with its old behaviour until 1.0. The behaviour changes come
first.

**Behaviour changes**

- **A wrong page index is an argument error in `Pdf.Editor`.** `getPage`, `removePage`, `isPageRemoved` and
  `rotatePage` throw `PdfArgumentError` for an index that names no page — not a number, fractional, `NaN` or out of
  range — with `field: "index"`, and `rotatePage` for a rotation other than 0, 90, 180 or 270 with
  `field: "degrees"`. They threw `PdfStructureError`, which describes a malformed file, while an index naming a
  removed page was already a `PdfArgumentError`. Messages start with the call
  (`Pdf.Editor#getPage: index 3 names no page (0-2)`). The deprecated positional `splitPages(indices)` /
  `copyPagesFrom(pdf, indices)` keep `PdfStructureError`.
- **`Pdf.Builder#addBookmark`** throws `PdfArgumentError` (`field: "parent"`) for a `parent` that names no top-level
  bookmark. It threw `RangeError`.
- **`Layout.pageToSvg` / `Layout.computedPageToSvg`** throw `DocxArgumentError` (`field: "pageNumber"`) for a page
  number out of range, instead of `RangeError`; `computedPageToSvg` also checks its `layout` (`field: "layout"`,
  `"layout.pages"`, `"layout.totalPages"`).
- **`Transform.map`** throws `DocxArgumentError` (`field: "transformer.transformBodyContent"`) when a transform puts a
  non-paragraph into a text box, a shape's text or a TOC's cached paragraphs; **`Diff.compare`** throws `DocxError`
  (`Diff.compare: …`) when the documents are too large to compare, and checks `newDoc` as it checks `oldDoc`. Both
  threw a plain `Error`, which `isDocxError` did not recognise. `WordCrypto.parseSignature` and
  `WordCrypto.isWellFormedSignature` check their arguments.
- **`ArchiveFile` wraps every file-system refusal in the new `ArchiveFileError`** (`path`, `operation: "read" |
"write"`, the system error as `cause` — read `err.cause.code` where you read `err.code`): opening
  (`fromFile` / `fromFileSync`), writing (`writeFile`, `writeFileSync`, `streamToFile`), a local source read while
  building (`toBuffer`, `toBufferSync`, `stream`, `pipeTo`; the message says "source file") and extraction
  (`extractTo`, `extractToSync`, `extractEntryTo`, `extractEntryToSync`; "extracted file"). An `overwrite` refusal is
  still a plain `ArchiveError`, a missing file or permission error that extraction reports through `onWarning` is
  still a warning, and `Fs.*` still pass the platform's error on.
- **`Compression.Workers.*` on Node** throw (or reject with) `ArchiveWorkerUnsupportedError` instead of a plain
  `ArchiveError`; guard isomorphic code with `Compression.hasWorkerSupport()`, which is `false` there.
- **Error names aligned with the other modules:** `ExcelNotSupportedError` → `ExcelUnsupportedFeatureError` (as
  `DocxUnsupportedFeatureError`) and `MaxItemsExceededError` → `ExcelLimitExceededError` (as `DocxLimitExceededError`
  / `PdfLimitExceededError`). The old names are deprecated aliases of the same class; the error's `name` is the new
  one.
- `WorksheetWriter.removeWatermark()` returns `true` if the sheet had a watermark, as `Watermark.remove` does.
- Argument errors name the public parameter: `ExcelFormula.calculate` reports `field: "wb"` (was `"workbook"`);
  `Csv.applyDynamicTypingToRow` reports its second argument as `dynamicTyping` (was `options.dynamicTyping`);
  `Pdf.verifySignatures` names its first parameter `data`, the field it already reported.

**Deprecated, with the replacement**

| Deprecated                                                                                  | Use                                                                                      |
| ------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| `CellValueType`, `CellValueInputType`, `Cell.Value`                                         | `CellValue`, `CellValueInput`                                                            |
| `HeaderFooterImage.Entry` / `Options` / `Position`                                          | the flat `HeaderFooterImageEntry` / `…Options` / `…Position`                             |
| `Sparkline.AddSparklineGroupOptions`, `Sparkline.SparklineGroup`                            | the flat `AddSparklineGroupOptions`, `SparklineGroup`                                    |
| `Workbook.XlsxReadable` / `XlsxWritable` / `XlsxStreamOptions`                              | the flat types of the same names                                                         |
| `Row.commit(ws, n)`                                                                         | `Stream.commitRow(ws.getRow(n))` — a document-model sheet has nothing to commit          |
| `Cell.setName(ws, addr, n)`, `Cell.removeAllNames(ws, addr)`                                | `Cell.setNames(ws, addr, [n])`, `Cell.clearNames(ws, addr)`                              |
| `EXCEL_CHART_PRESETS`, `EXCEL_CHART_EX_PRESETS`                                             | `Object.keys(CHART_PRESETS)`, `Object.keys(CHART_EX_PRESETS)`                            |
| `VECTOR_PDF_CHART_EX_LAYOUT_IDS`                                                            | deprecated with `canRenderChartExAsVectorPdf`                                            |
| `WorkbookWriter.nextId`, `WorksheetWriter.stream` / `drawing` / `_name`                     | internal; removed in 1.0                                                                 |
| `CsvInput`, `CsvOptions` (`documonster/excel/csv`)                                          | `ExcelCsvInput`, `ExcelCsvOptions` (with `ExcelCsvReadOptions` / `ExcelCsvWriteOptions`) |
| `MarkdownOptions` (`documonster/markdown`)                                                  | `ExcelMarkdownReadOptions` / `ExcelMarkdownWriteOptions` (`documonster/excel/markdown`)  |
| `WordSecurityPolicy`, `WordRenderContext`, `WordChartExOptions`                             | `DocxSecurityPolicy`, `DocxRenderContext`, `DocxChartExOptions`                          |
| `Query.StyleIndex`                                                                          | `DocxStyleIndex`                                                                         |
| `HtmlToDocxResult` (`documonster/word/html`)                                                | `WordHtml.toDocxBody`, whose result is `DocxHtmlBodyResult`                              |
| `PdfExportOptions.font`                                                                     | `fonts: { default: { regular: bytes } }`                                                 |
| `Pdf.Builder#embedFont(bytes, index)`, `Pdf.Editor#embedFont(bytes, index)`                 | `embedFont(bytes)` or `embedFont({ data, collectionIndex })`                             |
| `BasicRasterCanvas` positional `drawText` / `drawPolyline` / `strokeEllipse` / `strokeRect` | the options forms: `drawText(x, y, text, { fontSize, color, anchor, … })`, …             |
| `BasicRasterCanvas#textWarnings()`, `SvgSurface#markup()`                                   | `getTextWarnings()`, `format()`                                                          |
| `Archive.editZip`, `Archive.editZipUrl`                                                     | `Archive.ZipEditor.open`, `Archive.ZipEditor.openUrl`                                    |
| `Archive.randomBytes`                                                                       | `crypto.getRandomValues(new Uint8Array(n))`                                              |
| `EOCDInfo`, `ZIP64EOCDInfo`                                                                 | none — no public signature uses them                                                     |
| a TAR entry's `data()` / `skip()`                                                           | `bytes()` / `discard()`, as `TarReaderEntry` and `UnzipEntry` name them                  |
| `Csv.formatNumber(v, sep)`, `Csv.parseNumber(v, sep)`                                       | `Csv.formatNumber(v, { decimalSeparator })`, `Csv.parseNumber(v, { decimalSeparator })`  |
| `Csv.applyDynamicTypingToRow(row, typing, castDate)`                                        | `Csv.applyDynamicTypingToRow(row, typing, { castDate })` (also takes `decimalSeparator`) |
| `Csv.createGenerator(…).asyncRows()`                                                        | `rowsAsync()`                                                                            |
| `Build.mathFraction(n, d, type)`, `mathLimit(b, l, type)`, `mathAccent(c, char)`            | `{ fractionType }`, `{ limitType }`, `{ char }` as the last argument                     |
| `Anchor.asInstance(input)` (answered `null` for no input)                                   | `Anchor.isAnchor(x) ? x : Anchor.create(undefined, x)`                                   |
| `Xml.Writer#save()`                                                                         | `Xml.Writer#snapshot()` (`commit` / `rollback` are unchanged)                            |
| `SaxTag`, `SaxEvent`, `SaxEventAny`, `SaxHandlers`, `SaxOptions`                            | `XmlSaxTag`, `XmlSaxEvent`, `XmlSaxEventAny`, `XmlSaxHandlers`, `XmlSaxOptions`          |
| `Token`, `AstNode`, `TokenType`, `NodeType` (`documonster/formula`)                         | `FormulaToken`, `FormulaAstNode`, `FormulaTokenType`, `FormulaNodeType`                  |

**Order and naming fixes in names that had not been released:** `WordCrypto.Agile.deriveKey` takes
`(info, password, blockKey)`, matching `Agile.verifyPassword(info, password)` (the deprecated flat
`deriveEncryptionKey(password, info, blockKey)` keeps its order); `Chart.getTypes`, `Form.Vml.getAnchor` /
`getStyle` / `getCheckedValue`, `DrawingMl.withSpPrFill` / `withSpPrLine` (they return a copy, which `set` did not
say), `WorkbookEncryptOptions`, `XlsxParseStream`, `DocxChartSvgOptions`; `Table.getColumn` returns `undefined` for
an index the table has no column at.

**New:** `Stream.ColumnHandle`; `ExcelArgumentError` from `documonster/chart`; `DocxError` / `DocxParseError` from
`documonster/word/crypto`; `ArchiveFileError`, `ArchiveWorkerUnsupportedError`, `UnpackedTarEntry`;
`CsvGenerator`, `CsvNumberOptions`, `CsvDynamicTypingOptions`; `PdfEncryptionOptions`, `PdfPermissions`,
`PdfCellCommentMode`; `DrawRasterTextOptions`, `DrawRasterTextRotation`, `DrawRasterStrokeOptions`,
`DrawRasterLineOptions`, `DrawGlyphPoint`.

### File-system errors are typed; Word's `Document` has no file-path members in the browser

- **`Document.readFile` / `Document.writeFile` wrap a file-system failure in `DocxFileError`** (new, a `DocxError`, with `path`
  and `operation`). It was the platform's own error, which `isDocxError` did not recognise; its code is now
  `error.cause.code`, as for `ExcelFileError`. A file that was read but is not a document is still a `DocxParseError`.
- **`Workbook.writeFile`, `ExcelCsv.writeFile` and `ExcelMarkdown.writeFile` wrap a file-system failure in
  `ExcelFileError`** (`operation: "write"`, system error as `cause`), as their `readFile` counterparts already did.
- **The browser build of `documonster/word` publishes no `Document.readFile` / `Document.writeFile`.** They threw at runtime
  there; read the bytes yourself and use `Document.read`, and persist the bytes of `Document.toBuffer`.
- Renamed: `InvalidZipSignatureError` → `ZipInvalidSignatureError` (`documonster/archive`) and
  `UnsupportedStreamTypeError` → `StreamUnsupportedTypeError` (`documonster/stream`). The old names are deprecated
  aliases of the same class, removed in 1.0; the error's `name` is now the new one.

### Excel: facet readers return copies; `Row.commit` creates nothing; `Chart.clone` deprecated

- **`Cell.getFont`, `getAlignment`, `getBorder`, `getFill`, `getProtection` and `getNumFmt` always return a copy
  (behaviour change).** Editing the object returned for an existing cell used to change the cell, while the same
  edit on an absent cell changed nothing, and `Row`/`Column` readers always returned copies. Change a facet with its
  setter (`Cell.setFont(ws, "A1", { ...Cell.getFont(ws, "A1"), bold: true })`) or edit the live record
  `Cell.getStyle(ws, "A1")` returns.
- **`Row.commit` no longer creates the row on a document-model sheet.** It is still a no-op there, and still writes
  the row out on a `Stream.WorksheetWriter`.
- Deprecated: `Chart.clone(c, range)` — use `Chart.copyTo(c, c.worksheet, range)`, which places the same copy at the
  same default position and returns its handle (the part number the deprecated `clone` returns is the handle's
  `chartNumber`, `chartExNumber` for an Office 2016 chart). `clone` suggested an independent copy, as
  `Document.clone` returns.
- `Anchor.setCol` / `Anchor.setRow` name their second parameter `value`, matching the `field` their argument
  errors report.

### Excel: values, notes and models are the caller's own (behaviour change); one note API

Completes the facet-reader change above for everything else a reader hands out.

- **What a reader returns is a copy (behaviour change).** `Cell.getValue`, `Cell.getResult`, `Row.getValues`,
  `Column.getValues`, `Range.getValues`, `Worksheet.getValues`, `Worksheet.toAoa`, `Worksheet.toJson` and
  `Stream.getCellValue` / `getRowValues` used to return a rich-text value, a `Date`, an error value or a
  formula's cached `Date` result as the object the cell stores — editing it changed the cell with none of
  `Cell.setValue`'s checks — while a hyperlink, formula or checkbox value was already a fresh object. Every kind is
  a copy now. Change a value with `Cell.setValue` / `Range.setValues`. A test that compared a returned `Date` with
  `toBe` compares with `toEqual`.
- **`Cell.getModel` and `Worksheet.getModel` return a copy (behaviour change).** `Cell.getModel` returned the
  cell's live model and `Worksheet.getModel` shared the cells' value records, views and page setup with the sheet,
  so an edit to either changed the workbook unchecked; `Row`, `Table` and `Image` already returned copies. Apply a
  changed model with `Cell.setModel` / `Worksheet.setModel`.
- **`Cell.getNote` returns a copy (behaviour change).** A note configuration it returned was the cell's own.
- **A value given to a setter is copied (behaviour change).** A rich-text value, a `Date`, an error value, a formula's
  cached result, a JSON value, or a note configuration passed to `Cell.setValue` / `Cell.setNote` used to become the
  cell's, so editing your object afterwards — or giving one object to two cells — changed what the cells wrote.
- **Fixed: a removed note came back.** Reading a cell's model — which every save does — stored the note into the
  cell's value model, and a note removed afterwards was written into the next file from there.
- **One note API.** A cell's note is written with `Cell.setNote(ws, addr, note, { author })`, read with
  `Cell.getNote` and the new `Cell.getNoteAuthor`, and removed with the new `Cell.removeNote`, which answers
  whether there was one. On a stream, `Stream.setCellNote(cell, note, { author })` and `Stream.removeCellNote`.
  Deprecated, removed in 1.0:

  | Deprecated                                           | Use                                                                                                                 |
  | ---------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------- |
  | `Cell.getComment(ws, a)?.note` / `.author`           | `Cell.getNote(ws, a)` / `Cell.getNoteAuthor(ws, a)`                                                                 |
  | `Cell.setComment(ws, a, Note.create(n, { author }))` | `Cell.setNote(ws, a, n, { author })`                                                                                |
  | `Cell.setComment(ws, a, undefined)`                  | `Cell.removeNote(ws, a)`                                                                                            |
  | `Stream.setCellComment(c, …)`                        | `Stream.setCellNote(c, n, { author })` / `Stream.removeCellNote(c)`                                                 |
  | the `Note` namespace                                 | the `Cell` members above; a note's round-trip model is `Cell.getModel(ws, a).comment`, applied with `Cell.setModel` |

  The deprecated members behave as before, including `Cell.getComment` returning the cell's own record.

- Deprecated: calling `Worksheet.getValues(ws)` without `options`. It answers 1-based, where `Row.getValues` and
  `Column.getValues` answer 0-based; state the base — `{ base: 1 }` for the same result.

### Word: image size in the options; paper sizes; header-row formatting

- Deprecated: `Document.addImage(doc, data, mediaType, width, height, options?)` and the same form of
  `Document.addFloatingImage` — pass `Document.addImage(doc, data, { type, width, height, … })`. The width and height
  are EMU, as before; convert with `Units.cmToEmu` / `inchesToEmu` / `pxToEmu`. Behaviour and return values are
  unchanged.
- New: `PageSizes` (`A3`, `A4`, `A5`, `LETTER`, `LEGAL`, `TABLOID`, in twips) for `SectionProperties.pageSize`.
- New: `SimpleTableOptions.headerRunProperties` formats the header row's text (`{ bold: true }`). `headerRow` still
  only marks the row as the table's header — repeated on every page — and does not change how it looks.

### Excel: inserting, deleting and renaming keep references consistent (behaviour change)

`Row.insert`, `Row.insertMany`, `Row.splice`, `Row.duplicate`, `Column.splice` (and the deprecated
`Worksheet.insertRow`, `insertRows`, `spliceRows`, `spliceColumns`, `duplicateRow`), `Worksheet.setName` and
`Workbook.removeWorksheet` now update every reference to the cells they move, on every sheet, as Excel does.
Before, only the cells, merges, images, chart anchors and (partly) defined names moved; everything else kept its
old text and silently pointed at the wrong cells — or, after a rename, at a sheet that no longer existed.

What follows the edit now: cell formulas on all sheets (plain, shared — master text, range and member pointers —
and array / dynamic-array formulas with their ranges), defined names, tables (their range; rows or columns inserted
strictly inside a table become part of it, a new column is named `Column1`, `Column2`, …; deleted ones shrink it;
a table whose cells are all deleted is removed), conditional-format ranges and formulas, data-validation ranges
and formulas, chart series, categories and titles (classic and ChartEx), sparklines, auto filters, print areas and
titles, threaded comments, ignored-error ranges, pivot table sources, the cells form controls are linked to, and
formulas in hyperlink cells. Pictures, charts, shapes and preserved form controls move by their own placement
(`twoCell` moves and resizes, `oneCell` moves and keeps its size, `absolute` stays), and manual page breaks move with
the line after them, going when it is deleted. A structured reference to a table the edit deletes becomes `#REF!`.
Removing a sheet (or a chartsheet) removes the names scoped to it and renumbers every later `localSheetId`.

The rules are Excel's and apply to absolute and relative references alike:

- inserting at a range's first line moves it; inserting strictly inside grows it; inserting just past its last line
  leaves it alone;
- deleting shrinks a range by the lines it loses; a reference whose cells are all deleted becomes `#REF!`;
- a whole column (`A:A`) is unaffected by row edits, a whole row (`3:3`) by column edits;
- a 3-D reference (`Jan:Mar!A1`) is not changed by a row or column edit on one of its sheets;
- a rename requotes the name where needed (`Data!A1` → `'Q1 Data'!A1`); removing a sheet turns references to it
  into `#REF!`, and a 3-D span ending on it shrinks to the next sheet inward;
- a splice that both removes and inserts overwrites the lines it replaces, so references to them stay put
  (`Row.splice(ws, 3, 1, values)` changes no reference to row 3).

Two results differ from before for defined names, which were the one thing partly updated: an insertion inside a
name's area now grows it (`$A$1:$A$4` → `$A$1:$A$5`) where it used to be cut in two around the new line, and a whole
column stays `$A:$A` after a row deletion instead of becoming `$A$1:$A$1048575`. A name whose every cell is deleted
now refers to `#REF!` instead of disappearing (which turned the formulas using it into `#NAME?`).

Internal hyperlink locations (`#Sheet1!A1`) are text, and — as in Excel — are not rewritten.

`Row.spliceCells(ws, row, start, count, ...values)` is new: it shifts the cells of one row left or right, each keeping
its value, style and note, as exceljs's `row.splice` did. It is deliberately not one of the edits above — no reference
follows those cells — so code ported from exceljs keeps exceljs's result. Replacing the row with `Row.setValues` is not
the same: it carries values only, and drops the moved cells' styles and notes.

Pictures move by their placement where they used to move by a fixed rule: a picture placed by a range (`"B2:D5"`,
written `editAs="oneCell"`) no longer grows when rows are inserted inside it, and one whose top rows are deleted now
starts at the deleted band instead of row 1.

### Excel: a checkbox moves with its cells, as Excel's own do (behaviour change)

`Form.addCheckbox` used to write every checkbox as "don't move or size with cells" — in all three places a file says
it: `x:MoveWithCells` and `x:SizeWithCells` in the VML, no `moveWithCells` on the `<control>` anchor, and an
`absolute` bridge anchor. Excel gives a checkbox it inserts "move but don't size with cells", so inserting rows
above one of ours left it behind while the cell it is linked to moved away. A read had the same effect on a
checkbox Excel wrote: its positioning was not read, so a round trip turned "move" into "stay".

A checkbox is now written "move but don't size" by default, and `editAs: "oneCell" | "twoCell" | "absolute"` in
the options chooses another; a read keeps the file's choice in `model.editAs`, and inserting or deleting rows and
columns moves the checkbox by it. Pass `editAs: "absolute"` for the old behaviour.

### Excel and Word: a file's value outside its field's vocabulary is kept as read (behaviour change)

An enumerated attribute a file states with a value the field does not take — `<w:jc w:val="bogus"/>`,
`<pageSetup orientation="sideways"/>`, a `zoomScale` that is not a number — used to be put in the typed field as
it was, so `alignment` could hold a string outside `Alignment`. It is now read as an absent field and the text is
kept beside it: writing the document puts it back unchanged while the field is absent, and a value set then is
checked like any other — one outside the vocabulary is refused by the write with `DocxArgumentError` or
`ExcelArgumentError`. Every vocabulary is the schema's (Transitional and Strict), held to the ECMA-376 schemas by a
test.

What was kept is reported: in Word by `Validation.validate`, as warnings with rule `"unrecognized-value"` and the
path to the field; in Excel by `Workbook.readWithDiagnostics`, in the new `unrecognizedValues` (a list of
`{ sheet, path, value }`). An XLSB write cannot
store such a text, so it names the setting in `unsupported` (refused under the default `unsupported: "error"`).
Clearing a field again brings the kept text back, since absent is what the reader left; `Document.discardUnrecognizedValues`
and `Workbook.discardUnrecognizedValues` remove what was kept — all of it, or with `{ filter }` the values you pick by
path and text (and, in Excel, by sheet).

Two number settings are read more strictly as part of this: a split pane's `xSplit` / `ySplit` keep their fraction
(they were truncated to whole numbers), and an Excel page or view number with trailing text (`scale="100%"`) is kept
as read rather than read as its leading digits.

### Excel and Word: encrypted files are read and written (new, one implementation)

Password-protected Office packages — ECMA-376 Agile Encryption in an OLE compound file, what Excel, Word and
PowerPoint write for "Encrypt with Password" — are read and written by both modules, from one implementation (an
internal module of `archive`, beside the ZIP code; it has no public namespace).

- **Excel reads them**: `Workbook.read`, `readFile`, `readStream`, `readWithDiagnostics` and `Stream.WorkbookReader`
  take `password` — the option name Word's `Document.read` uses — for XLSX and XLSB, the decrypted package's format detected
  as usual. Without it: `ExcelEncryptedError` (now pointing at the option). With a wrong one: the new
  `ExcelDecryptionError` (an `ExcelFileError`) with `reason: "password"`; other reasons are `"unsupported"`,
  `"malformed"`, `"integrity"`, `"limit"`. The decrypted package is bounded by `zip.maxTotalUncompressedSize`
  (default 2 GiB). An encrypted workbook cannot be streamed — it is one encrypted stream — so the streamed readers
  decrypt it whole first.
- **Both write them**: Excel `Workbook.toBuffer`, `writeFile`, `writeStream`, `toStream` (XLSX and XLSB) and Word
  `Document.toBuffer`, `toBase64`, `writeFile` take `password` (1–255 characters) and `encryption` (`EncryptOptions`:
  `keyBits`, `hashAlgorithm`, `spinCount` — the type `WordCrypto.encrypt` takes, now exported by
  `documonster/excel` too). Defaults are Office's: AES-256, SHA-512, 100,000 iterations. Excel 16 and Word 16 open the
  output with the password (AES-128/192/256, SHA-256/512 checked). **Word D7 is reversed: `Document.toBuffer` encrypts when
  given a password**; it wrote an unencrypted package and the TSDoc sent callers to `WordCrypto.encrypt`, which keeps
  working unchanged. `Package.*` keep `password` as a read option: their result is not encrypted. The encryption code
  loads only when a password is given (`import()`), so a consumer that never encrypts does not bundle it; there is no
  synchronous form.
- **Only Agile Encryption** (version 4.4) is implemented, in both modules. Standard Encryption (Office 2007), RC4
  CryptoAPI and Extensible Encryption are refused with `reason: "unsupported"`, the message naming the scheme.

Bug fixes in what Word already had (`WordCrypto.*`, `Document.read` with `password`):

- **Files Office encrypted decrypt.** `WordCrypto.decrypt`, `Document.read`, `Streaming.createReader` and
  `WordCrypto.Agile.decryptPackage` verified the password of a Word-, Excel- or third-party-encrypted file and then
  returned garbage (reported as a corrupt package): the package segments were keyed with the password key encryptor's
  salt instead of the key data's. Only files this library wrote, which used one salt for both, decrypted. It now writes
  two salts, as Office does. `AgileEncryptionInfo` gains `keyData` (the key data's salt, hash, sizes), and
  `dataIntegritySalt` / `dataIntegrityHmac` — declared but never filled — hold the file's `encryptedHmacKey` /
  `encryptedHmacValue`.
- **The data-integrity HMAC is checked** on decryption, as Office does: an altered package is a
  `DocxDecryptionError` (`reason: "integrity"`), not a silently damaged document.
- `DocxDecryptionError.reason` (new, optional): the same values as `ExcelDecryptionError.reason`.
- **A compound file that is not an encrypted package** (a legacy `.doc`) read by `Document.read` / `Streaming.createReader` is a
  `DocxParseError` saying so; it was a `DocxEncryptedError` asking for a password that could not help.
- **`keyBits: 192`** encrypts and decrypts; it failed with `RangeError: Invalid key length` on Node (the AES layer named
  every non-128-bit key AES-256).
- **Packages over about 7 MB** encrypt: `WordCrypto.Cfb.write` writes the DIFAT sector chain MS-CFB requires once the
  FAT outgrows the header; it refused with `DocxParseError`. The 256 MiB stream cap no longer applies while decrypting
  (the read's own package-size limit does).
- A malformed `EncryptionInfo` message names the scheme it found (`4.2 (Standard Encryption)`), and an unencrypted
  input to `WordCrypto.decrypt` says "Not an encrypted Office package" (it said "DOCX").

### Contracts tightened across every module (behaviour change)

Every change below fixes a bug: a call that was accepted and then wrote a file Office or a viewer refuses, failed
later with an untyped `TypeError`, or quietly did something other than what it says. Changes the maintainers chose
deliberately are marked _(owner decision)_. A call that is now refused changes nothing before it throws. Leniency that
v0.21.1 documented or tested is kept, and each case is named below as unchanged.

Each module lists **output changes** first — the same call now gives different bytes or values, so read these even
if your code never passes a bad argument — then **now refused**, then **new names and deprecations**.

Three rules hold in every module:

- **A wrong argument is the module's argument error, with a readonly `field`**: the path from the parameter name in
  the public signature (`options.delimiter`, `data[3]`, `source`, `options.theme.nodeFill`), or, for a document model
  checked while it is written, the path from the document root (`body[3].properties.alignment`). The message starts
  with the public call as written (`Csv.parse: …`, `Pdf.Builder#addPage: …`) wherever the check knows it, and never
  names an internal function; match on `field`, not on message text. The value it was given is described the same
  way everywhere: a primitive as its value, a built-in object by its kind (`got a Map`, `got a Uint8Array`, `got an
invalid Date`), a plain object as `an object` — not `got object`, nor "must be a valid Date, got a Date". Each extends its module's base
  error, so existing `catch` blocks and `is*Error` guards still match: `ExcelArgumentError` (Excel and its CSV,
  Markdown and formula bridges), `DocxArgumentError` (Word; `DocxInvalidDrawingError` and `DocxInvalidImageError`
  are its subclasses), `PdfArgumentError`, `DrawArgumentError`, `MermaidArgumentError`, `CsvArgumentError`,
  `MarkdownArgumentError`, `XmlArgumentError`, `ArchiveArgumentError` and `FormulaArgumentError`. Streams keep
  `StreamTypeError` / `UnsupportedStreamTypeError`, which now carry an optional `field`. Errors about the data
  itself — a corrupt file, a CSV record error, malformed XML, a formula syntax error — keep their own classes.
- **Handles are checked at the boundary.** A function that takes a workbook, worksheet, table, chart, document,
  glossary, archive source, sink or plan refuses `42`, `null`, `{}` or a name with that argument error, naming the
  parameter (`ws`, `wb`, `doc`, `workbook`, `chart`, …). It used to fail with a `TypeError` from inside, or return a
  result for `{}`. The check is one property read on the hot paths.
- **`error.toJSON()` carries the error's own data** (`field`, `line`, `position`, `path`, …), and `BaseErrorOptions`
  is exported from every entry that exports error classes.

#### Excel — document model

**Output changes**

- `Row.splice` / `Worksheet.spliceRows` deleting through the last row now delete it. A block reaching past the last
  row (`splice(4, 2)` on five rows) left the rows in place while every reference to them became `#REF!`. `Row.getLast`
  (and the deprecated `Worksheet.lastRow`) then answers the last row that remains; it was `undefined`.
- `Address.quoteSheetName` and `Formula.format` quote sheet names by one rule: unquoted only for ASCII word
  characters with no leading digit and nothing that reads as a cell, an R1C1 reference or a boolean (`A1`, `XFD1`,
  `R1C1`, `R`, `C12`, `TRUE`). Range and defined-name text follow it too: a sheet named `2023` is written
  `'2023'!$A$1` (it was `2023!$A$1`), `My_Sheet` is no longer quoted, and an XLSB formula with a cell-like sheet name
  reads back as `'S1'!A1`.
- Lowercase references are accepted everywhere and read as uppercase, as Excel's UI reads them (`b2` is `B2`):
  `Cell.*` refused `a1`, while `Chart.place` / `Image.place` accepted `d5:h10` and wrote `NaN`. Stored references are
  canonical (`Worksheet.setAutoFilter(ws, "a1:c4")` stores `A1:C4`).
- A six-digit `argb` is written as `FF` + rgb in XLSX, as XLSB always did (`rgb="FF0000"` was invalid); either
  format reads back `"FFFF0000"`.
- `Workbook.importWorksheet` gives a copied table whose name the target already uses the next free name (`Sales` →
  `Sales2`), as Excel's "Move or Copy" does, and restates the copied sheet's formulas, table-column formulas,
  conditional formats and validations under it. The copy kept the name and the workbook could not be written.
- `Column.splice` drops the cells it empties at the end of each row, so `Column.getCount` / `getLast`,
  `Row.getCellCount` and `Cell.has` agree with the written file. `Column.splice` (and the deprecated
  `Worksheet.spliceColumns`) writes every inserted value: values below the sheet's last row were dropped — on an empty
  sheet, all of them. A value no cell can hold is refused before anything moves; it was refused part-way, leaving some
  rows shifted and the references already rewritten.
- `Workbook.readWithDiagnostics(...).unrecognizedValues` is a list of `{ sheet, path, value }` (type
  `UnrecognizedValue`, also what the `discardUnrecognizedValues` filter receives), not sentences.
- `Workbook.readFile` keeps the file system's error as `cause` (`error.cause.code === "ENOENT"`), for XLSX and XLSB.
- `Worksheet.setModel` with a new `name` renames as `Worksheet.setName` does: formulas and defined names follow. The
  model replaces the sheet's merges (applying a sheet's own model twice threw `ExcelMergeConflictError`) and frees the
  names of the tables it drops.
- After `Workbook.setModel`, the handles of the previous worksheets are retired: an edit through one is refused. It
  used to vanish.
- Naming a cell with a formula-defined name (`Cell.addName` / `setName` / `setNames`) turns the name into a reference,
  as `DefinedNames.add` does. It used to shadow the formula, and removing the cell then deleted the name from the file.
- `Worksheet.merge` / `mergeWithoutStyle` on a single cell does nothing. The merge used to be recorded, reported by
  `getMergedRegions`, then dropped from the file.
- These now return copies, not live objects: `Row.getModel`, `Column.getDefinition` / `getDefinitions` (style
  included), `Worksheet.getAutoFilter`, `Workbook.addCellStyle` / `getCellStyle` / `listCellStyles`. The deprecated
  `Worksheet.autoFilter` still returns the live object. External links stay live (now documented).
- `Worksheet.getPanes` returns `{ state: "frozen", columns, rows }` or `{ state: "split", x, y }` (types
  `WorksheetPanes`, `FrozenPanes`, `SplitPanes`) — what `Worksheet.freeze` / `split` take, which now accept an
  optional matching `state` — instead of the raw view (`xSplit`, `topLeftCell`, …, still in `ws.views`). The
  deprecated `Worksheet.panes` still returns the live view.
- `Worksheet.getDimensions` returns `undefined` for a sheet with no cells. The deprecated `Worksheet.dimensions` keeps
  its zero range, which `Range.*` read as A1.
- A formula set through the API that produces an array is a dynamic array. That covers a function that returns one
  (`=SEQUENCE(3)`, `=FILTER(…)`, `=SORT(…)`, `=UNIQUE(…)`, `=TRANSPOSE(…)`, `=MMULT(…)`, `=MAP(…)`, `=TEXTSPLIT(…)`,
  `=XLOOKUP(…)`), an array constant (`={1,2,3}`), arithmetic on either (`=SEQUENCE(3)*2`), `=LET(x,SEQUENCE(3),x*2)`
  or an invoked `LAMBDA` returning an array, and a formula with a spill reference on the path from its root (`=A1#`,
  `=A1#*2`, `=ABS(A1#)`, and the same with `ANCHORARRAY(A1)`). The XLSX, XLSB and both streaming writers write it with
  dynamic-array metadata (`t="array"` over its own cell, `cm` → `XLDAPR`; XLSB `BrtArrFmla` + `BrtCellMeta`), and
  `ExcelFormula.calculate` spills it (`{2,4,6}` for `A1#*2` over `SEQUENCE(3)`). It was written as an ordinary
  formula, which Excel computes as its top-left value (and `A1#` on it is `#REF!`), while `calculate` spilled some of
  them and computed the rest as one intersected value. Such a formula can no longer be a shared-formula master: like
  other dynamic arrays it is written cell by cell. Under a parameter that takes a range whole a spill reference is a
  range like any other: `SUM(A1#*2)` is now 12 (it was 2) and is a dynamic array; `SUM(A1#)` is unchanged. References
  at the root (`=A1:A3`, `=INDEX(A1:C3,0,1)`, `=OFFSET(…)`) are not dynamic arrays: ordinary, intersected (for a
  multi-cell `OFFSET` this is new; see Formula below).
- A formula read from a file as an ordinary one is no longer spilled: `<f>_xlfn.SEQUENCE(3)</f>` without `cm`
  computes its top-left value, and `A1#*2` stored that way is evaluated with Excel's intersection (`@A1#*2`), as
  Excel does.
- The implicit-intersection operator `@` is written as `_xlfn.SINGLE(…)` in XLSX — cells, shared and array formulas,
  defined names, conditional formats, data validations, table column formulas — where it was written verbatim
  (`<f>@A1:A3</f>`, outside MS-XLSX's formula grammar, whose only prefix operators are `+` and `-`). `@x` wraps exactly
  the operand the parser gives `@` (`@A1:A3*2` is `_xlfn.SINGLE(A1:A3)*2`); `@` inside a structured reference
  (`Tb[@x]`) is untouched. On read, `_xlfn.SINGLE(x)` reads back as `@x` where that is the same formula, and as the call
  otherwise (`_xlfn.SINGLE(A1)^2`). The formula engine computes `_xlfn.SINGLE(x)` as `@x` (it was `#NAME?`). In XLSB,
  `@x` is written as Excel 16.114 writes it: a call to the `_xlfn.SINGLE` stub (`PtgName`, the operand, `PtgFuncVar`
  0x00FF) — in array and dynamic-array formulas, conditional formats, data validations, defined names, and cell
  formulas where `x` alone would not be intersected (`SUM(@A1:A3)`, `INDEX(@A1:A3,1)`, `@A1`) — and as `x` alone in the
  value class in a cell formula where `x` is intersected in its place anyway (`@A1:A3`, `@A1:A3+1`, `@Rng`), which
  reads back bare, as Excel's own XLSB does. It used to be written as the value-class operand everywhere, which outside
  a cell formula, and inside `SUM(…)`, is a different formula (the whole range).
- A reference to a sheet whose name holds an apostrophe (`'it''s'!$B$1`) names that sheet. The quoted name was taken
  with its doubled apostrophe (`it''s`), so `DefinedNames.add` covered cells of a sheet that does not exist and listed
  `'it''''s'!$B$1`, reading a file doubled the quotes again, and every other address taking a sheet
  (`DefinedNames.getNames`, data validations, chart caches, ranges) missed the sheet. Fixed in `colCache.decodeEx`.
- A defined name whose text has an `!` after something that is not a sheet name — `@S!$A$1:$A$3`, `-S!$A$1` — reads
  back as the formula it is; it was taken for a region on a sheet named `@S`.
- `A1#` (also `Sheet2!A1#`, `'My Sheet'!$A$1#`, inside any expression) is written as `_xlfn.ANCHORARRAY(A1)` in XLSX
  and XLSB — cells, shared and array formulas, defined names, conditional formats, data validations, table column
  formulas. It was written verbatim (in XLSX, a broken formula) or refused (XLSB). On read,
  `_xlfn.ANCHORARRAY(<cell ref>)` reads back as `<cell ref>#` (`SUM(_xlfn.ANCHORARRAY(D!H1))` reads as `SUM(D!H1#)`);
  this is the one exception to reads keeping the file's spelling — `_xlfn.XLOOKUP(…)` still reads as stored — and is
  lossless: the file written back is identical. A call on anything else (a name, a range, a call) keeps the function
  form. Deleting a spill anchor rewrites `A1#` to `#REF!` (it gave `#REF!#`).
- An XLSB array or dynamic-array formula whose cached result is an error reads back as that kind of formula; it read
  as an ordinary formula, losing its `shareType`, `ref` and `isDynamicArray`.
- An XLSX formula cell whose cached result is the empty string (`<c t="str"><f>…</f><v/></c>`, as Excel writes it)
  reads back with `result: ""`; it read as no result. A plain `t="str"` cell with `<v/>` (an array formula's cell
  holding `""`) reads as `""`. Both are written back as `<v></v>`; XLSB already read them.
- `ExcelFormula.calculate` writes every cell of a legacy array formula's range (`shareType: "array"`, `ref`): the
  master takes the result and every other cell becomes a plain cached value, as Excel stores it. Only cells already
  holding the formula were written, so `{=A1:A3}` over B1:B3 left B2:B3 empty or stale. A cell of the range that
  holds a plain value is overwritten (it is the array's); a merged cell is skipped. Empty elements are 0, as the
  engine now computes them.
- A dynamic array read from a file whose `ref` covers its spill (what Excel saves: the spill's cells are plain cached
  values) no longer calculates to `#SPILL!`. On load those cells become the master's spill; the engine takes the
  stored `ref` as the previous spill, reusing it and clearing what a smaller spill no longer reaches, and a value set
  over one of its cells blocks the spill as before. A spill cell inside the stored `ref` is written back as its cached
  value, so reading and writing an Excel file keeps the spill's values; one outside it (a spill calculated past the
  stored `ref`, or with no stored `ref`) is still not written.
- `Workbook.readStream` reads XLSB with `{ format: "xlsb" }`, buffering the stream; it accepted and ignored the
  option.
- `DefinedNames.list` reports `hidden: true` for a hidden name.
- A column key named like an `Object.prototype` member (`"toString"`) no longer resolves through the prototype:
  `Column.setWidth(ws, "toString", 5)` set a property on `Object.prototype.toString`. A key still wins over a column
  letter of the same text (now documented).
- `Workbook.read` accepts an `ArrayBuffer` from another realm (iframe, worker, `node:vm`).
- `Workbook.readStream` whose source fails rejects with the source's error. It used to wait forever.
- Named cell style names are compared without regard to case, as Excel compares them (Excel 16 merges `Brand Title`
  and `brand title` into one style on save). `Workbook.addCellStyle(wb, "brand title", …)` when `Brand Title` exists
  replaces that style's definition and keeps the name `Brand Title`; `getCellStyle`, `removeCellStyle` and
  `Cell.applyStyle` find a style in any case, and `applyStyle` records the style's own spelling. `"normal"` in any
  case is the built-in and is refused. A file carrying two styles that differ only in case is read with the first,
  and cells of the second point at it. The XLSX and XLSB writers and `WorkbookWriter.addCellStyle` follow the same
  rules.

**Now refused**

- At compile time: `PageSetup.orientation`, `pageOrder`, `cellComments` and `errors` are typed by their vocabularies
  (`PageSetupOrientation`, `PageSetupOrder`, `PageSetupCellComments`, `PageSetupErrors`) instead of `string`. A value
  outside them was already refused when the workbook is written and set aside when one is read, so only the type
  changed: assigning an unchecked `string` now needs a check (or the cast you would have needed for the runtime to
  accept it). This restores the types excelts and exceljs had.
- Splice and duplicate counts: `Row.splice` / `Column.splice` (and `Row.insert*`, `Row.duplicate`, the deprecated
  `Worksheet.splice*`) refuse a negative, fractional or `NaN` count with `ExcelArgumentError`, and a start outside the
  grid with `ExcelRowOutOfBoundsError` / `ExcelColumnOutOfBoundsError`. A negative count used to duplicate a row and shift every
  reference. `Row.splice` refuses a bare number as a row to insert, as `Row.add` does; `Column.splice` refuses a column
  that is not an array.
- An edit through a removed sheet's handle: `Row.splice`, `Column.splice`, `Column.setDefinitions`,
  `Worksheet.unmerge`, `Cell.removeName`, `Cell.removeAllNames`. The splices rewrote the references of the sheet that
  had since taken the removed sheet's name.
- A malformed model: `Worksheet.setModel` and `Workbook.setModel` throw `ExcelArgumentError` and change nothing.
  `Worksheet.setModel(ws, {})` wiped the cells and renamed the sheet, and `Workbook.setModel(wb, {})` removed every
  sheet, before a raw `TypeError`. The workbook model is applied to a scratch workbook first.
- `Cell.setModel` with a model for another address (a `Z99` model applied to `D4` wrote a file the library could not
  read back), an unknown `type`, or a value of the wrong kind.
- `Cell.addName` / `setName` / `setNames`: `ExcelDefinedNameError` for a name Excel would repair (`A1`, `1abc`,
  `has space`) or one a table uses. `setNames` checks every name before removing any.
- `Address.*` with coordinates outside the grid: `encodeCol(1.5)` returned `undefined`, `encodeRow(-5)` `"-4"`,
  `decodeCell("zz")` `{ c: 701, r: null }`. Any string that is not a reference is an `ExcelInvalidAddressError` (`"garbage"`
  was a `ExcelColumnOutOfBoundsError`); a well-formed reference past XFD stays `ExcelColumnOutOfBoundsError`.
- `Range.create` / `expand` / `expandToAddress` with text that names no cells (`"garbage"` gave `{}`, `"A0"` row 0).
  `Range.create()` with no arguments is still the empty accumulator.
- `Workbook.read` / `readWithDiagnostics` with something that is not bytes (`42`, `null`, `{}`, text without
  `{ base64: true }`) reject with `ExcelArgumentError` (field `data`), as `Workbook.readStream` does with something that
  is not a stream (its promise never settled). The read form used to be an `ExcelFileError`: test `ExcelArgumentError`
  or `isExcelError` instead. The message describes the argument, for example
  `Workbook.read: data must be bytes (Uint8Array, ArrayBuffer, a typed-array view) or a base64 string with { base64: true }, got number`;
  the old "Can't read the
  data of 'the loaded zip file'" sentence is gone. `Workbook.readFile` / `writeFile` refuse a path that is not a
  non-empty string (field `filename`); `writeFile` raised a Node `TypeError`.
- Read and write options: not an object, a `format` other than `"xlsx"` / `"xlsb"` (`toBuffer(wb, { format: "csv" })`
  wrote XLSX), a non-boolean `base64`, an unknown `unsupported`; `maxRows` / `maxCols` that are not a whole number
  ≥ 1 or `Infinity` (`NaN` and `0` disabled the guard); `ignoreNodes` that is not an array of names (a string matched
  by substring and dropped every cell); the `zip` limits, `checkCrc32`, `blankCells`, `formulas`, `zip.level` /
  `modTime` / `timestamps`, `useSharedStrings`, `useStyles`, `templateMode`, `strictTemplateMode`, `validate` and
  `formulaSyntaxProbe` (also in `Workbook.create`).
- `Workbook.toBuffer` / `writeFile` / `writeStream` / `toStream` with `zip.level` outside 0–9 reject; the promise never
  settled. A part that fails to compress now rejects the write instead of hanging it.
- A damaged package — bytes that are not a ZIP package, or a part that fails its CRC-32 — is an `ExcelFileError`
  (with the archive's error, such as `ZipParseError`, as `cause`) from every reader: `Workbook.read`, `readFile`,
  `readStream`, XLSB reads and `Stream.WorkbookReader`. `isExcelError` recognizes it; a CRC mismatch escaped as the
  archive's `Crc32MismatchError`. Archive limit and password errors keep their own classes.
- A malformed value of a known shape, in `Cell.setValue`, the `Row.*` / `Column.*` value setters, `Worksheet.addAoa`
  and the stream setters, is an `ExcelInvalidValueTypeError`: rich text that is not a list of `{ text, font? }`
  (`{ richText: 42 }`), an empty or non-text `formula` /
  `sharedFormula`, a `result` that is not a scalar, `Date` or `{ error }`, a non-text `hyperlink` / `text` /
  `tooltip`, an `error` that is not an Excel error literal (`{ error: "#BOGUS" }`), a non-boolean `checkbox`, and a
  shared formula whose master is not a formula cell above or to the left. An object of none of those shapes is still
  stored as its JSON text. `ExcelInvalidValueTypeError` now extends `ExcelArgumentError` and has a `field`;
  `instanceof ExcelError` and `instanceof InvalidValueTypeError` still match. Messages name the call and the argument
  (`Row.add: values[1]: …`).
- `Column.setValues` checks every value before writing any; a refused value left the earlier ones written.
- `field` is the path from the public parameter name: `Cell.setFont(ws, a, value)` names `value.size` (was
  `font.size`), `Table.add(ws, model)` `model.columns[0].name`, `Workbook.addWorksheet` `options.state`,
  `Sparkline.add` `options.type`, `Pivot.add` `model.rows`, `Chartsheet.setPageSetup` `value.orientation`,
  `Worksheet.freeze` `options.rows`.
- `Range.containsCell` with an address that is not `{ row, col }` (a string always answered `false`), and
  `Workbook.removeChartsheet` with a target that names nothing — a string, a fractional number, `null` — which
  answered `false`. An out-of-range position still answers `false`.
- `Workbook.importWorksheet` checks `source`, `name` and the sheet name before adding the sheet; a refused import left
  a sheet behind.
- `DefinedNames.setModel` with two entries that are one name in one scope (`N` and `n`): `ExcelDefinedNameError`. One of
  them was dropped. Reading a file is unchanged.
- A sheet name over 31 characters is refused, not shortened _(owner decision)_: `Workbook.addWorksheet`,
  `Worksheet.setName`, `Worksheet.setModel`, `Workbook.importWorksheet`, `Workbook.addChartsheet`,
  `Workbook.copyChartsheet` (a given name) and the deprecated `Workbook.renameChartsheet`. Reading a file still
  shortens a long name.
- Row and column setters: a non-boolean to `Row.setHidden` / `Column.setHidden`; `Row.setHeight` over 409 points and
  `Column.setWidth` over 255 characters (Excel's limits). `Column.setDefinitions` checks every definition (width,
  outline level, hidden, key, header, style) before changing anything; it wrote `width="-5"` and failed at write time
  on `NaN`. `Column.setHeader` checks its values first.
- `Column.setKey` with a key that is not a non-empty string, or that another column has.
- `{ base }` other than `0` or `1` on `Row.getValues` / `setValues` and `Column.getValues` / `setValues`;
  `{ base: 7 }` cleared the row.
- A `(ws, row, col)` form without a row number: `Cell.setValue(ws, "A1", 5, "x")` wrote to `EA1`. Now
  `ExcelInvalidAddressError`, for every `Cell` reader and writer, `setDateParts`, `getTemporal`, `applyStyle` and
  `removeName`.
- `Cell.setDateParts` checks `parts` and `kind` before creating the cell (a refused value left an empty cell); an
  unknown `kind` is refused instead of ignored.
- `Cell.setResult` on a cell that holds no formula: `ExcelArgumentError`, where it was dropped. `undefined` is still
  accepted.
- `Cell.applyStyle` with an unknown style: `ExcelArgumentError`, a subclass of the `ExcelError` it threw.
- `Worksheet.merge` / `mergeWithoutStyle` / `unmerge` with a range qualified by another sheet; it was applied to this
  sheet.
- An external link whose `target` an edit emptied: refused when the workbook is written.
- `Workbook.registerFunction` checks its name, function and arities, `unregisterFunction` its name.
- `Workbook.registerFunction(wb, name, fn, { returnsArray: true })` declares that the function returns an array: a
  formula calling it (`=F()`, `=F()*2`) is a dynamic array — spilled by `ExcelFormula.calculate` and stored with
  dynamic-array metadata by the XLSX and XLSB writers — as `SEQUENCE` is. A registered function's declaration decides
  over a built-in of the same name, and the default is `false`: **a function registered as `SEQUENCE` (or any other
  array-returning built-in) without `returnsArray: true` is now treated as returning one value** — its formula is
  ordinary and an array it returns is intersected to the top-left value, as Excel computes an ordinary formula (it
  used to spill, by the built-in's name). `calculateFormulas`' `FormulaFunction` takes the same `returnsArray`.
- A formula set with `{ formula, isDynamicArray: true }` and no `shareType` / `ref` is written to XLSX as a dynamic
  array over its own cell (`t="array"`, `ref` = its cell); it was written as an ordinary `<f>` carrying `cm`, which
  names an array formula that was not one. XLSB already wrote it that way.
  `Workbook.useBuiltInCellStyle` refuses an unknown id, `Workbook.addCellStyle` a style that is not an object.
- `Workbook.getExternalLink`, `Note.fromModel`, `Format.parseValue`, `Worksheet.toJson`, `DefinedNames.setModel`,
  `Workbook.copyChartsheet` / `getChartsheet` / `removeChartsheet` with a non-name, `Workbook.eachWorksheet` /
  `Row.each` / `Row.eachAddress` / `Column.each` / `Column.eachAddress` with a non-function, and
  `Address.quoteSheetName(42)`: an `ExcelArgumentError`, not a raw `TypeError`. `Worksheet.toJson` refuses an
  unbounded or unreadable `range` (`ExcelInvalidAddressError`) and an unknown `header`.
- A wrong handle to any function of `Workbook`, `Worksheet`, `Cell`, `Row`, `Column`, `DefinedNames`, `Note` and
  `Range` (field `wb`, `ws`, `dn`, `n`, `r`); `Worksheet.getName(42)` returned `undefined`.

**New names and deprecations**

- `DefinedNames.remove(dn, name, { localSheetId })` deletes a whole name (reference, formula or opaque); the
  `(dn, location, name)` form is unchanged. `DefinedNames.get(dn, name, { localSheetId })`, and the type
  `DefinedNames.NameScopeOptions`.
- `Workbook.listPersons(wb)`, `Workbook.listExternalLinks(wb)`.
- `Workbook.getDefaultFont(wb)` / `Workbook.setDefaultFont(wb, font)` — the font of the "Normal" style. It was settable
  only through `Workbook.setModel` since the internal `_defaultFont` field it lives in was hidden; excelts had it as
  `wb.defaultFont`. They take a streaming writer too (with `useStyles: true`).
- `Row.spliceCells(ws, row, start, count, ...values)` — exceljs's `row.splice` (see "inserting, deleting and renaming"
  above).
- `Range.toString(r, { absolute, compact })` — `$B$2:$D$5` and a one-cell range as `B2`: exceljs's and excelts's
  `$range`, `shortRange` and `$shortRange`, which nothing public wrote.
- `Column.isCollapsed(ws, col)` — excelts's `col.collapsed`, as `Row.isCollapsed` answers for a row.
- `ExcelCsvWriteOptions.mapRow(values, rowNumber)` — turn each row written into another, as exceljs's
  `formatterOptions.transform` did; `map` turns one value.
- `ExcelArgumentError.field` names the path from the offending parameter. Types `AddExternalLinkOptions`,
  `RegisterFunctionOptions`, `UnrecognizedValue`, `BaseErrorOptions`, `FrozenPanes`, `SplitPanes`, `WorksheetPanes`,
  `WorkbookStreamInput`, `ExcelFileErrorOptions`, `ExcelInvalidValueTypeErrorOptions`; error class `XlsxParseError`.
- `Column.has(ws, col)`.
- Deprecated: `DefinedNames.create()`. A registry it makes cannot be attached to a workbook; use
  `Workbook.getDefinedNames(wb)`.
- Widened: `Cell.setNumFmt` / `Row.setNumFmt` / `Column.setNumFmt` take `string | NumFmt` (what the getters return);
  `Cell.setNames`, `Column.setDefinitions`, `Column.splice`, `DefinedNames.setModel` and the row values of `Row.add` /
  `addMany` / `insert` / `insertMany` / `setValues` / `splice` take readonly arrays. `Row.setHeight(ws, row, undefined)`
  and `Column.setWidth(ws, col, undefined)` clear the height / width, and `Worksheet.setAutoFilter(ws, undefined)`
  clears it, as `null` does. Value parameters take `CellValueInput` (civil `Temporal` values): `Row.add` / `addMany` /
  `insert` / `insertMany` / `splice` / `setValues`, `Column.splice`, `Worksheet.addAoa` / `addJson`, `Table.addRow` /
  `addColumn`. `Workbook.readStream(stream: WorkbookStreamInput, options?: WorkbookReadOptions &
WorkbookCreateOptions)`. `Workbook.removeChartsheet` / `copyChartsheet` / `replaceChartsheetChart` accept the
  chartsheet's handle.
- **Positions are argument errors.** `ExcelRowOutOfBoundsError`, `ExcelColumnOutOfBoundsError` and `ExcelInvalidAddressError` extend
  `ExcelArgumentError` (still `ExcelError`s; `instanceof` either class and `isExcelError` hold as before) and carry
  `field` — the parameter (`row`, `col`, `addr`; `cstr`, `cell.r`, `range`, `cells`, `data`, `args`, … where the call
  says) — and the `Address.*`, `Range.*`, `Worksheet.merge` / `addAoa` / `addJson` and argument-check messages now start
  with the call (`Address.encodeCol: Column -1 is out of bounds: …`, where the call used to sit in the details). The
  constructors keep their positional forms and take `{ field, where }` in their options (`ExcelPositionErrorOptions`). A
  `catch (e instanceof ExcelArgumentError)` written for bad input now catches a bad position too.
- **A bad position in a file is the reader's parse error, not an argument error** (class change on read paths). A
  position the content puts outside the grid or not where it is written — `<c r="B7">` inside row 5, column `XFE`,
  `<row r="0">`, a merge from `C0` — is an `XlsxParseError` naming the part (`xl/worksheets/sheet1.xml: …`) from
  `Workbook.read` / `readFile` / `readStream` / `readWithDiagnostics` and `Stream.WorkbookReader`, with the position
  error as `cause`; it was the bare `ExcelInvalidAddressError` / `ExcelColumnOutOfBoundsError` / `ExcelRowOutOfBoundsError`, and an
  out-of-range `<row r>` a plain `ExcelError` ("Invalid row number"). A CSV or Markdown row a worksheet cannot hold
  (more than 16 384 fields) is an `ExcelFileError` naming the row (`row 2 does not fit a worksheet: …`, `cause` the
  `ExcelColumnOutOfBoundsError`) from `ExcelCsv.read` / `addWorksheet` / `readFile` / `createImportStream` and
  `ExcelMarkdown.read` / `addWorksheet(s)` / `readFile`. XLSB already reported these as `XlsbParseError`; a binary
  sheet read by `Stream.WorkbookReader` does too. The same positions given at an API call stay argument errors.
- `Stream.WorkbookReader`'s `maxEntrySize` refusal (`FileTooLargeError`) carries `allowed` — the configured limit — as
  the archive's own does; it was `NaN`.
- **Column setters check the column first.** `Column.setWidth` / `setHidden` / `setOutlineLevel` / `setKey` /
  `setHeader` / `setStyle` / `setNumFmt` / the facet setters / `addPageBreak` / `getStyle` and
  `WorksheetWriter#getColumn` refuse a number that is not a whole column from 1 to 16 384 with
  `ExcelColumnOutOfBoundsError` before anything is created, as the getters do. `0` or `1.5` was a raw `TypeError`; `16385`
  was accepted (and the getter then refused it), materialising every column up to it — `1e7` stalled the process —
  and `setStyle` applied the style before failing. `WorksheetWriter#getColumn("toString")` no longer finds
  `Object.prototype.toString`.
- **Sheet ids are never reused in a session**, as Excel does it (Excel 16.114: sheets 1–3, delete 3 and add → 4,
  delete 1 and add → 5). A new worksheet or chartsheet gets one past the highest id the workbook holds or has removed;
  the lowest free id used to be handed out, so a numeric id held across a remove named another sheet, and a worksheet
  could inherit a removed chartsheet's id. A read continues past the file's highest id.
- `Workbook.read` and `Workbook.readWithDiagnostics` accept a `Blob` (a `File` from `<input>`), XLSX and XLSB.
- **An encrypted workbook read without its password is an `ExcelEncryptedError`** (new, extends `ExcelFileError`) from
  `Workbook.read` / `readFile` / `readStream` / `readWithDiagnostics` and `Stream.WorkbookReader`, saying to pass the
  `password` read option; it was an `ExcelFileError` about a ZIP without a central directory. Another OLE compound file
  (a legacy `.xls`) is an `ExcelFileError` that says so. With its password it is read — see
  [Excel and Word: encrypted files are read and written](#excel-and-word-encrypted-files-are-read-and-written-new-one-implementation).
- `Workbook.writeStream` of an unmodified XLSB to a Web `WritableStream` resolves; it wrote the package and then waited
  forever, because the stream reported `finish` before the completion check was listening.
- `Workbook.writeStream` accepts a Web `WritableStream` (via the stream module's `SinkAdapter`) in both containers,
  and resolves once the stream has closed; it crashed with `this.pipedStream.write is not a function`. A `stream` that
  is neither a Node-style writable nor a Web stream is an `ExcelArgumentError` (`field: "stream"`).
- `Worksheet.verifyPassword` refuses a `password` that is not a string (`ExcelArgumentError`); a number reached the
  hash and failed there.
- An argument error from a chartsheet's chart names the caller's path: `options.chart.series[0].values` from
  `Workbook.addChartsheet` / `addPivotChartsheet`, `chart.series[0].values` from `replaceChartsheetChart` — it was
  the validator's `options.series[0].values` — and the message starts with the call.

#### Excel — tables, charts, images, pivots, forms, streaming, CSV/Markdown bridges

**Output changes**

- `ExcelMarkdown.read`, `addWorksheet`, `addWorksheets` and `readFile` type the data cells
  as `ExcelCsv` does _(owner decision)_: numbers, ISO dates, `true` / `false` and error literals become values, and
  `<br>` becomes a line break (`convertBr` defaults to `true` here). Header cells stay text. Pass `map` or
  `convertBr: false` for your own reading. The deprecated flat `readMarkdown*` functions and the deprecated
  workbook-first overloads keep reading text.
- `ExcelCsv.read*` no longer reads a whitespace-only cell as `0`, through the `Csv.parseNumber` fix below. It parses
  in the plain form, so a malformed CSV (text after a closing quote, an unterminated quote) now throws unless `onSkip`
  is passed — see CSV below.
- `ExcelCsv` and `ExcelMarkdown` reads take a date with no time zone as that calendar value in UTC: `2024-03-10` is
  the cell `2024-03-10 00:00`, serial 45361, in every time zone (it was local midnight), as are offset-less date-times
  such as `2024-03-10T14:30:00`. A text with `Z` or an offset is the instant it names, as before.
- `ExcelCsv` and `ExcelMarkdown` writes (`format`, `toBuffer`, `toBufferSync`, `toStream`, `writeStream`,
  `writeFile`, and the deprecated `writeCsv*` / `writeMarkdown*`) take a date cell's fields from UTC: `dateUTC` now
  defaults to `true`. For the cell `2024-03-10 00:00`, no options gave `2024-03-09T19:00:00.000-05:00` in New York
  and now give `2024-03-10T00:00:00.000Z` everywhere; `{ dateFormat: "YYYY-MM-DD" }` gave `2024-03-09` and now gives
  `2024-03-10`. In a UTC host only the no-`dateFormat` form changes (`+00:00` becomes `Z`). Old output still reads
  back exactly; pass `dateUTC: false` for the previous bytes.
- Pivot tables follow their source. A `sourceSheet` pivot covers the range of the sheet's populated cells when added;
  that range moves and grows with row and column edits, and the cache (items and records) is recomputed from the
  source cells when the workbook is written. A `sourceTable` pivot reads the table at write time. A value added to
  the source made `toBuffer` throw, and a row inserted above the source became its header.
- `DataValidation.add`, `find`, `Cell.setValidation` and `Cell.getValidation` store and return copies: editing the
  object you passed or got back changes nothing — re-add the rule to change it.
- `ConditionalFormatting.add` (and the deprecated `Worksheet.addConditionalFormatting`) stores a deep copy of the
  options and returns the stored block. It stored the caller's object, so one options object added to two sheets was
  one shared entry, and a row insert on one sheet rewrote the other's rule. `ConditionalFormatting.remove(ws, block)`
  takes the returned block; it no longer matches the caller's original object. It also takes the block's `ref`, and a
  predicate removes the first matching block (it removed all of them, against its TSDoc). `ConditionalFormatting.add`
  stores `ref` upper-cased.
- `Watermark.get` returns a copy; editing the returned object changed what `get` reported but never the file.
- `ExcelCsv.writeFile({ append: true })` appends every row, the sheet's first included, as it always did; nothing is
  guessed from the file's content. Its documentation claimed "the header row is skipped", which never happened — that
  was the bug, and it is fixed in the documentation. To continue a file that already has the header, pass
  `writeHeaders: false`, which now leaves out the sheet's first row in every writer (`format`, `toBuffer`,
  `toBufferSync`, `toStream`, `writeStream`, `writeFile`); it had no effect on a sheet before. A line break is written
  before the appended rows only when the file does not already end with one, read in the write `encoding` (a UTF-16LE
  file's `0a 00`, `\r\n`, a byte-order mark alone); it used to test the last byte only. `writeHeaders` must be a
  boolean, and `writeFile` refuses the byte encodings `hex`, `base64` and `base64url`.
- `ExcelCsv` `quoteHeaders` quotes the sheet's header row (its first row) as `Csv.format` quotes a header row — `true`,
  an array by column index, an object by header name; it had no effect. The header row is now quoted by
  `quoteHeaders`, not `quoteColumns`: `quoteColumns` as an array (`[true]`) or object no longer quotes row 1 (as
  `true` it still quotes every field, the header's included, as in `Csv.format`), and an object `quoteColumns` matches
  the header row's texts (`{ qty: true }`), which it never matched before.
- `ExcelCsv`'s `encoding` is the encoding of the input bytes for every byte input — `Uint8Array`, `ArrayBuffer`,
  `Blob`, `File`, byte streams, URLs, `readFile` — as `Csv.parse` takes it. It reached only the `File` path (a `Blob`
  was read as UTF-8 regardless). `ExcelCsv.read(file)` works where there is no `FileReader` (Node, Bun); it threw
  `ReferenceError`.
- A CSV or Markdown file the system will not read (`ExcelCsv.readFile`, `ExcelMarkdown.readFile`) is
  an `ExcelFileError` that says "CSV file" / "Markdown file" and keeps the system error as `cause`.
- `ExcelMarkdown.read({ tables: "all" })` / `addWorksheets` shorten the base name so `name_2`, `name_3`, … fit 31 characters, and
  check every name, a clash with an existing sheet included, before adding any sheet; a clash left the sheets before
  it behind.
- `Table.add` keeps a copy of its model; editing the object afterwards no longer changes the table. `Sparkline.add`
  copies the `sparklines` array.
- `Image.addShape` returns the `ShapeModel` and `Image.place` the placed image's handle (the one `Image.list`
  returns); both returned `void`.
- The deprecated `Table.columnSetName` forwards to `Table.setColumnName`, so a colliding name is refused rather than
  renumbered. The deprecated `column*` / `show*` / `headerRow` getters and setters forward to the checked functions,
  and the deprecated `Form.vmlAnchor` / `vmlStyle` / `vmlCheckedValue` to `Form.Vml.getAnchor` / `getStyle` / `getCheckedValue`, which also take the handle.
- `Stream.WorkbookReader`'s `parse(input, options)` / `read(input, options)` apply the defaults to what the call
  leaves out. `{ styles: "cache" }` dropped `worksheets: "emit"`, so no sheet was read.
- A committed `WorksheetWriter`'s `findRow` / `eachRow` return nothing instead of throwing; `mergeCells("A1")` does
  nothing.
- An `ExcelCsv` read parses its input before adding the sheet, so input the parser refuses leaves the workbook as it
  was.

**Now refused**

- A CSV record the parser leaves out — a row with the wrong number of fields under a header row, text after a closing
  quote — fails an `ExcelCsv` read (`read`, `readFile`, `addWorksheet`, `createImportStream`) with a `CsvError` naming
  its line, unless `onSkip` receives it. With a header row such records were dropped from the sheet without a word
  (without one, the read already failed). Pass `columnMismatch: { less: "pad", more: "truncate" }` to import them, or
  `onSkip` to leave them out knowingly.
- A wrong handle to any function of `Table`, `Chart`, `Chartsheet`, `Image`, `Anchor`, `Form`, `Pivot`, `Sparkline`,
  `DataValidation`, `ConditionalFormatting`, `HeaderFooterImage` and the `Stream.*` handle functions:
  `ExcelArgumentError`, `field` naming it (`ws`, `wb`, `table`, `column`, `chart`, `chartsheet`, `anchor`,
  `checkbox`, `image`, `row`, `cell`). `Image.add(42)` was an `ExcelImageError` and `Table.setName(42, …)` a
  `ExcelTableError`; others failed inside or returned `false` / `undefined`. `Watermark.get` checks its handle too.
- A removal target that names nothing — a string, a fractional number, `null` — in `Chart.remove`, `Sparkline.remove`,
  `HeaderFooterImage.remove` (also a section outside `LH` … `RF`), `Image.remove`, `Image.removeShape`,
  `Pivot.remove`, `Form.remove` and `ConditionalFormatting.remove` (`ExcelArgumentError`); they answered
  `false`. An out-of-range position or another sheet's item still answers `false`.
- A sheet-qualified reference where the target is the sheet's own cell or range (`ExcelInvalidAddressError`):
  `Table.add` / `setModel` / `setRef` `ref`, `DataValidation.add` / `find` / `remove`, `Pivot.add` `ref`,
  `Anchor.create`, `WorksheetWriter#mergeCells`. The qualifier was dropped or misread. `Anchor.setCol` / `setRow` /
  `setModel` / `create` refuse values off the grid.
- Tables (`ExcelTableError` unless noted):
  - `Table.addRow` / `addColumn` / `removeRows` / `removeColumns` with a position that is not a whole number in range.
    `-1`, `1.5`, `NaN` or `99` lost values, duplicated a row or copied the header into the body. `addRow` refuses more
    values than the table has columns, `addColumn` more than it has data rows; a negative or fractional `count` is
    refused (a `count` past the end still removes to the end); a table keeps at least one column.
  - An area another table on the sheet covers, in `Table.add`, `addRow`, `setRef` and `setModel`. The second table
    overwrote the first one's cells and Excel repaired the file.
  - A name or display name another table uses (as either) or a defined name holds, in `Table.add`, `setName`,
    `setDisplayName` and `setModel`. Two tables could share a `displayName`, which made structured references
    ambiguous.
  - Colliding column names in `Table.add`, `addColumn`, `setColumnName` and `setModel` _(owner decision)_; the message
    gives the name Excel would choose. Reading a file still renumbers repeated headers, as does the deprecated
    `Table.create`.
  - `Table.add` checks everything before it writes a cell, and `Table.setModel` checks the new rows before it clears
    the old cells; a refused call used to leave cells changed.
  - Values of the wrong kind to `setHeaderRow`, `setTotalsRow`, `setTheme`, the `setShow*` and `setColumn*` setters
    and the same fields of a model (`ExcelArgumentError`). `setColumnTotalsRowLabel(42)` failed at write time;
    `setTheme(42)` wrote `name="42"`.
  - An edit through the handle of a removed table — by `Table.remove` or by a row or column edit — or of a removed
    column (`ExcelArgumentError`). `setRef` on a removed table moved cells it no longer owned.
  - `Table.get` / `Table.remove` with a name that is not a string: `ExcelArgumentError`, not a raw `TypeError`.
- Conditional formats and validations: a `cellIs` rule without an `operator` (Excel 16 for macOS cannot open the
  file); data-bar and color-scale colors that are not colors (written as `auto`);
  `DataValidation.add(target, ref, undefined)`, which removed the rule and returned `undefined` — use
  `DataValidation.remove`.
- Images and shapes:
  - `Image.add` and `Stream.WorkbookWriter#addImage` refuse an image with no extension, with neither `buffer`,
    `base64`, `filename` nor `link`, or whose `buffer` is not bytes (`ExcelImageError`). These failed at write time, or
    wrote a string as a PNG.
  - `Image.place` checks the range: corners in the grid, non-negative extents, `br` not above or left of `tl`. They
    were written as `NaN` or negative EMU.
  - `Image.setModel` checks the type, the image id and the range, and refuses a removed image. An unknown id silently
    dropped the picture.
  - `Image.addShape` checks the geometry (ECMA-376 `ST_ShapeType`), colors, line width (0–1584 pt) and range.
- Form controls: a checkbox `link` (`Form.addCheckbox`, `Form.setLink`) that is not a cell, a range or a defined name,
  optionally after a sheet name (`"not a ref !!"` was written as `'not a ref !'!`); a non-boolean to
  `Form.setChecked`; an edit through a removed checkbox's handle (`ExcelArgumentError`).
- Charts and chartsheets: an edit through a removed chartsheet's handle; `Chartsheet.setPageMargins` with keys that
  are not margins; `Chart.copyTo` / `Chart.clone` of a removed chart now throw `ExcelArgumentError`, like its other
  edits (it was a `ExcelChartOptionsError`). The table and rows option builders and `createSeriesFromColumns` refuse a
  missing or non-object `options`, or rows that are not an array, with `ExcelArgumentError` from every entry
  (`Chart.*`, `ChartClassic.*`, `ChartEx.*`, the deprecated flat names), as do `Chart.addSeries(c, null)` and
  `Chart.setUserShapesXml(c, 42)`; `ChartClassic.toSvg({})` and the ChartEx renderers throw `ExcelChartOptionsError`.
  `Chart.toPng` (classic and Office 2016) checks `dpi` and `scale` itself and throws `ExcelChartOptionsError` for an
  invalid value, in Node and in browsers alike; Node surfaced an `ArchiveArgumentError` from the PNG encoder, browsers
  ignored the value, and the Office 2016 path did not check `scale`.
  `Chart.applyPreset` with an unknown preset names both preset families. `Chartsheet.setName` applies the
  31-character rule (`ExcelWorksheetNameError`).
- `Chart.*` argument errors name the call the caller wrote — `Chart.createOptionsFromRows: …`, `Chart.place: …` —
  rather than the shared builder that found the mistake; class and `field` are unchanged. `Chart.mutate` /
  `mutateEx` refuse a `mutator` that is not a function (field `mutator`; a `TypeError` from inside) and non-boolean
  options; `Chart.copyTo` / `clone` check `range` up front (a malformed one was a `TypeError`). `Chart.setTitle`
  refuses with field `value` (was `title`) and says `Chart.setTitle: value must be …` (it named the deprecated
  `Chart.title`); `Chart.setLegend`'s field is `value` (was `legend`). `ExcelChartOptionsError` is an
  `ExcelArgumentError` with a `field`, and the chart renderers check their options; see PDF, drawing and Mermaid
  below.
- Pivot tables:
  - Removing a pivot table's source _(owner decision)_ is a `ExcelPivotTableError`: its worksheet
    (`Workbook.removeWorksheet`, `Worksheet.destroy`), its table (`Table.remove`), or a sheet holding that table. This
    holds for pivots read from files too, by the cache's `worksheetSource`. A pivot table on the sheet being removed
    goes with it. A row or column edit that deletes a source table entirely is not covered.
  - `Pivot.add` refuses a placeholder block overlapping another new pivot on the sheet — typically two left at the
    default `ref` `A3`, which Excel cannot refresh — and a model of the wrong shape (`ExcelArgumentError`).
- Streaming (`Stream.*`):
  - `WorkbookWriter#addWorksheet` applies the document model's sheet-name rules: a duplicate (case-insensitive),
    illegal characters, a leading or trailing quote, `History`, and over 31 characters _(owner decision)_ are a
    `ExcelWorksheetNameError`. It wrote packages `Workbook.read` refused.
  - `Stream.setCellNote` / `setCellComment` check the note as `Cell.setNote` / `setComment` do; a bad note crashed
    `WorksheetWriter#commit`. The browser entry's style setters, which checked nothing, now share the Node
    implementation.
  - After `commit`, `ExcelStreamStateError` everywhere: a committed `WorksheetWriter` refuses `addRow`, `getRow` /
    `getCell`, `mergeCells`, `addConditionalFormatting`, `addImage`, `addBackgroundImage` and `addWatermark` (rows threw
    a raw `TypeError`; the rest were silently lost), the `Stream.*` setters refuse a committed row, and a committed
    `WorkbookWriter` refuses a second `commit`, `addWorksheet`, `addCellStyle`, `useBuiltInCellStyle`, `addImage`,
    `protect`, `unprotect` and setting `defaultFont`.
  - `WorksheetWriter#getRow(0)` is a `ExcelRowOutOfBoundsError`, `addRow(42)` an `ExcelArgumentError`, `getCell("zz")` and
    `mergeCells("Other!A1:B2")` an `ExcelInvalidAddressError`. `protect` checks its options, `addCellStyle` checks the
    style (and stores a copy), `useBuiltInCellStyle` the id.
  - `WorksheetWriter#addImage` / `addBackgroundImage` / `addWatermark` refuse an image id the workbook does not have,
    which was silently dropped. `addConditionalFormatting` checks the rule as `ConditionalFormatting.add` does and
    keeps a copy; an invalid rule vanished from the file.
  - `WorkbookReader` checks its options (`ExcelArgumentError`); `worksheets: "emitt"` read no sheets.
  - `Stream.setColumnWidth` once the sheet's columns are written (with its first committed row):
    `ExcelStreamStateError`. It accepted the width and lost it.
- `ExcelCsv` and `ExcelMarkdown` raise `ExcelArgumentError` for every argument error of their own; the message starts
  with the function called (`ExcelCsv.read:`) and `field` is `options.<name>` (`options.delimiter`,
  `options.sheetName`, …). A bad parser or formatter option was a `CsvError` / `MarkdownError` and is now this class
  with the original as `cause`; a non-string `sheetName` on a read was a `ExcelWorksheetNameError`. `null` data, a
  non-object `options`, `ExcelMarkdown.read(42)` (was `MarkdownError`) and `ExcelFormula.calculate(null | {})` are
  refused the same way. Errors about unreadable content are unchanged.
- `ExcelCsv.format` / `toBuffer` / `toBufferSync` / `toStream` / `writeStream` / `writeFile` (and the deprecated
  `writeCsv*`) refuse a `sheetName` / `sheetId` the workbook does not have; they wrote an empty CSV. A workbook with
  no sheets still gives `""`. `dateFormat` must be a string, `dateFormats` known `DateFormat`s, `map` a function, and
  `writeFile` checks its options before touching the file. `append`, `includeEmptyRows` and `dateUTC` must be booleans
  (`"false"` was read as true, so `append: "false"` appended), `highWaterMark` a whole number ≥ 1, and `encoding` a
  Node `Buffer` encoding for `writeFile`; `format` / `toBuffer` / `toBufferSync` / `toStream` / `writeStream` produce
  UTF-8 and refuse another `encoding` instead of ignoring it. A read refuses an unknown `encoding` label (field
  `options.encoding`).
- `ExcelCsv.readFile` / `writeFile` and `ExcelMarkdown.readFile` / `writeFile` refuse a path that is
  not a non-empty string (field `filename`).
- `ExcelMarkdown.format` / `toBuffer` / `toBufferSync` / `writeFile` refuse a `sheetName` / `sheetId` the workbook
  does not have, and a `dateFormat` that is not a string.
- The 31-character sheet-name rule (`ExcelWorksheetNameError`) applies to `ExcelCsv.read` / `addWorksheet` and
  `ExcelMarkdown.read` / `addWorksheet` / `addWorksheets`.

**New names and deprecations**

- `Table.getRef` returns the table's whole range (`"B2:C4"`) however it was made; the deprecated `Table.ref` still
  returns the reference as given. `Table.remove` also accepts the table's handle.
- `Image.removeShape(ws, shape | index)`; `DataValidation.list(target)` (copies of `{ ref, rule }`) and
  `DataValidation.clear(target)`; `Stream.setRowHeight`, `Stream.setRowHidden`, `Stream.setColumnWidth` (`undefined`
  clears), `Stream.setCellProtection`, `Stream.setRowNumFmt`, `Stream.setRowProtection`, `Stream.setRowOutlineLevel`,
  `Stream.setColumnHidden`, `Stream.setColumnOutlineLevel`, and `Stream.setCellNumFmt` takes a `NumFmt`.
- `Stream.getRowValues(r, { base })` (`Row.ValuesOptions`, the one values-options type) is 0-based by default, as `Row.getValues` is; the
  deprecated `Stream.rowValues` stays 1-based.
- `CsvInput` (`ExcelCsv.read`, `addWorksheet`) names what the readers take: a Web `ReadableStream<Uint8Array>` and an
  `AsyncIterable<Uint8Array | string>` besides a Node readable stream. A Web stream is read through its reader, so
  one that is not async-iterable (as in browsers without `ReadableStream` iteration) is read too; it was refused.
- `documonster/excel/csv` and `/excel/markdown` export the errors they throw — `ExcelError`, `ExcelArgumentError`,
  `ExcelFileError`, `ExcelWorksheetNameError`, `ExcelInvalidValueTypeError`, `isExcelError`, `BaseErrorOptions`, plus
  `ExcelDownloadError` and `CsvError` / `isCsvError` for CSV and `MarkdownError` / `MarkdownParseError` /
  `isMarkdownError` for Markdown — and `/excel/formula` exports `ExcelError`, `ExcelArgumentError`, `isExcelError` and
  `BaseErrorOptions`. They are the same class objects as `documonster/excel`, `/csv` and `/markdown`. On Node,
  `ExcelCsv.createImportStream` is typed `IWritable<string | Uint8Array> & stream.Writable`.
- Deprecated: `DataValidation.create` — the collection it makes belongs to no sheet and nothing can attach it; pass
  the worksheet (or `ws.dataValidations`) to `DataValidation.add`.
- `ChartClassic.Backend` holds `toScene`, `toSvgFilter`, `applyAxisTransform` and `resolveTopologyObject`; the
  deprecated flat `buildChartScene`, `buildEffectFilter`, `applyAxisTransform` and `resolveTopologyObject` point
  there. `ChartClassic.fillCaches` / `ChartEx.fillCaches` take `{ worksheet }` (the deprecated `fillChartCaches` /
  `fillChartExCaches` keep the positional form). A chart goes onto a PDF page as its display list
  (`toDrawList` + `Pdf.createDrawSurface` + `Draw.render`); see "PDF, drawing and Mermaid".
- Types used in published signatures: `Chart.PresetOptions`, `Chart.TableExOptions`, `Chart.RowsExOptions`,
  `Stream.WriterImageModel`, `Stream.PackageSink`, `Stream.EventListener`; `ExcelMarkdownReadOptions` /
  `ExcelMarkdownWriteOptions` from `documonster/excel/markdown` (`MarkdownOptions` is still accepted everywhere, deprecated).
- `Chartsheet.getModel`, `getPageMargins`, `getPageSetup` and `getHeaderFooter` return copies; the deprecated
  `Chartsheet.model`, `pageMargins`, … still return the live objects.
- Deprecated with no replacement, because each builds a handle that is attached to nothing and never written:
  `Image.create`, `Image.clone` (use `Image.place`), `Chartsheet.create` (use `Workbook.addChartsheet`) and
  `Form.checkboxFromModel` — a checkbox's model is `handle.model`, and the VML getters take the model directly.
- Deprecated: the `Stream.WorkbookWriter` pipeline steps `addStyles`, `addThemes`, `addOfficeRels`,
  `addContentTypes`, `addMedia`, `addApp`, `addCore`, `addSharedStrings`, `addFeaturePropertyBag`, `addMetadata`,
  `addWorkbookRels` and `addWorkbook` — `commit` runs them, and calling one again writes its part twice — and the
  sheet-writer hooks `xlsbFormulaContext`, `date1904Flag` and `packageSink`. Sheets come from `addWorksheet`.
- **A chart colour keeps every DrawingML transform, in order.** `ChartColor`, `ChartColorsEntry` and
  `ChartColorVariation` gain an optional `transforms` — the colour's full transform list (`ChartColorTransform`, new
  type: `{ kind, val? }`, raw OOXML values) in document order. The six named fields (`tint`, `shade`, `satMod`,
  `lumMod`, `lumOff`, `alpha`) stay and are views of it; a named field that is present wins (replaces the first entry
  of its kind, or is appended), an absent one leaves the list as it is. Before, only those six survived parsing:
  `DrawingMl.parseSpPr`, every structured edit of a read `spPr` (`withSpPrLine`, `withSpPrFill`, `createSpPr`) and
  `ChartClassic.fromColorsXml` → `toColorsXml` dropped `satOff`, `hue` / `hueMod` / `hueOff`, `sat`, `lum`,
  `alphaMod`, `alphaOff`, `comp`, `inv`, `gray`, `gamma`, `invGamma` and `red` / `green` / `blue` with their `Mod` /
  `Off`, and the chart renderers (SVG, PNG, PDF) ignored them. They are now read in order, written back in order by
  the classic and ChartEx writers and the colours sidecar, and drawn through `applyColorTransforms`. A chart part
  written back unchanged was not affected (its raw XML is kept). A malformed list entry is a `ExcelChartOptionsError` on
  `transforms[i]` at write time.
- `ChartClassic.fromColorsXml` reads a self-closing `<cs:variation/>` as one variation. It ran on to the next
  variation's close tag, so Excel's default `colors1.xml` (an empty variation, then eight) read as eight, each one
  index early, and a write from the model shifted every gradient colour.
- `Form.removeCheckbox` (unpublished) is `Form.remove(ws, control)`, the `remove` paired with `Form.list`.
- `ExcelMarkdown.read`, `readFile`, `addWorksheet` and `addWorksheets` take the text or its UTF-8 bytes
  (`ExcelMarkdownInput`) — what `toBuffer` / `toBufferSync` produce read back; bytes were refused. The unpublished
  `ExcelMarkdown.readAll` / `readAllFile` are `read` / `readFile` with `{ tables: "all" }`; `addWorksheet` refuses
  `tables: "all"` and `addWorksheets` `tables: "first"`.
- The read-side `map` of `ExcelCsvReadOptions` and `ExcelMarkdownReadOptions` is declared as it is called:
  `(text: string, column: number) => CellValueInput` — the field's text and its 0-based column; what it returns is
  checked as `Cell.setValue` checks a value. The write-side `ExcelCsvWriteOptions.map` and the combined `ExcelCsvOptions.map`
  keep `(value: CellValue, index: number) => CellValue`.
- `Stream.applyCellStyle(cell, name)` applies a named style registered with `WorkbookWriter#addCellStyle` /
  `useBuiltInCellStyle`, checked as `Cell.applyStyle` checks one; setting `cell.style.styleName` by hand was the only
  way. One `ValuesOptions` type serves `Row`, `Column`, `Worksheet` and `Stream` (`Stream.RowValuesOptions`,
  unpublished, is gone).

#### Word

**Output changes**

- `Transform.fillFormFields` (and the deprecated `Query.fillFormFields`) writes a text field's value into the field
  result, which is what Word shows and `Query.getText` reads; the default text is left alone. Only `<w:default>`
  changed, so a filled field looked unchanged in Word. A drop-down takes an entry's index or its text and sets the
  current selection (`w:result`); it set the default selection.
- `Query.listFormFields` (and the deprecated `extractFormFields`) reports a text field's result and a drop-down's
  current selection, each falling back to the default, and lists fields in footnotes, endnotes and comments too.
  `DropDownFormField.result` is read and written, so a selection survives a round trip.
- `Build.formTextField({ default })` without `cachedValue` shows the default as its result, as Word does.
- Text watermarks (`TextWatermark`): `fontSize` takes effect — it is written as a shape one line of that size tall
  and as wide as the text; every watermark was 468pt × 234pt. A watermark Word wrote (`font-size:1pt`) reads as
  `fontSize: undefined` (it read `2`), and its shape box round-trips in the new `widthPt` / `heightPt`. One with no
  rotation reads as `rotation: 0` (it read `-45`, so a horizontal watermark turned diagonal). VML color names read
  as hex (`silver` → `"C0C0C0"`; it was written back as the invalid `#silver`).
- `Document.setSettings` keeps, unless the new object has the key (even as `undefined`), what a settings object
  cannot express or other functions manage: `preservedXml`, `mailMergeRawXml`, `compatibilityMode`,
  `compatSettings`, `compatFlags`, `documentProtection`, `writeProtection`, `rsids`, the note separators and values
  kept as read, and `evenAndOddHeaders`. `setSettings(readDoc, { zoom })` dropped them, removing
  `Security.protect`'s protection.
- `Document.setSectionProperties` keeps the section's header and footer references when `props` has no `headers` /
  `footers` (pass `[]` to drop them), and keeps `titlePage` while a kept reference is a first-page one. A read
  document's headers and footers were written but shown on no page.
- `Document.setHeader` / `setFooter` with a type (`"default"`, `"first"`, `"even"`) replace the header the last
  section references for that type, or add and reference one. On a read document the new content was written as an
  unreferenced part. `"first"` turns on the section's `titlePage`, and `"even"` turns on
  `settings.evenAndOddHeaders`; the streaming writer does the same unless the caller set them. On a built document
  they always write an explicit section reference, so `Layout.*` and `Pdf.fromDocx` show what the file shows (a
  header set on a built document was missing from every laid-out page). `"odd"` is an ordinary key, not a type; it
  was written as the invalid `w:type="odd"`.
- Charts: a chart without `c:legend` reads as `legend: "none"` (it was written back with a right-hand legend), and
  data-label positions read as the model names them (`outEnd` → `"outsideEnd"`, `ctr` → `"center"`, `inEnd` →
  `"insideEnd"`); other `ST_DLblPos` values are kept as read. The anchor enumerations (`ST_RelFromH/V`,
  `ST_AlignH/V`, `ST_WrapText`, `ST_LegendPos`) are vocabularies too: a value outside one is kept and reported as
  `unrecognized-value`.
- `WordExcel.toDocx` writes a cell as its number format applied to its value — the text `Cell.getDisplayText` gives
  and `Pdf.fromExcel` renders. Every cell was its raw value's `toString()`: a date came out as `Date#toString()` in
  the machine's time zone (so it could land on the wrong day), and `0.25` formatted `0%` as `0.25`. Its `sheets`
  names are compared ignoring case, a very hidden sheet named explicitly is included, and numbers keep their
  published meaning, deprecated (see below): a 0-based position among the worksheets, very hidden ones not counted.
- `Theme.resolveColor`: `"auto"` (and `{ val: "auto" }` with no resolvable theme color) gives `undefined`, as its
  return type promises; it returned `"auto"`. A `themeColor` the theme defines wins over `val`, as in Word
  (ECMA-376 §17.3.2.6), with `val` the fallback; any `val` was returned and the theme never consulted. Word's
  `text1` / `background1` / `text2` / `background2` resolve through the default mapping (dk1 / lt1 / dk2 / lt2);
  they were unresolvable.
- Bytes may be any `ArrayBufferView` everywhere Word takes bytes: a `DataView` or `Uint16Array` is read as the bytes it
  views, as `Workbook.read` does (they failed inside or were read element by element). `Document.read`'s input type is
  widened to match. Every function that takes a package or file also takes an `ArrayBuffer`, as `Document.read` did:
  `Package.*`, `Streaming.createReader`, `WordCrypto.isEncrypted` / `decrypt` / `verifyPassword` / `encrypt` /
  `Cfb.read` / `Agile.decryptPackage`, `Convert.fromOdt` / `isFlatOpc` / `flatOpcToParts`.
- `Streaming.createReader` (`StreamingDocxReader`) reads Flat OPC, as `Document.read` does; it failed as a corrupt ZIP. It
  decrypts an encrypted file with `{ password }`, which `StreamingReadOptions` always carried, and refuses one without
  a password with `DocxEncryptedError` (a `DocxError`); every encrypted file was a plain `DocxError`. A package over
  `maxPackageSize` is a `DocxLimitExceededError` (a `DocxParseError`, as before).
- `Package.listParts` / `readPart` / `edit` (and the deprecated `Io.listDocxParts` / `readDocxPart` /
  `editDocxIncremental`) apply the security policy's `maxPartCount` / `maxPartSize` / `maxPackageSize` (they applied
  none), read Flat OPC as `Document.read` does, and refuse an encrypted file with `DocxEncryptedError`; it was a
  `DocxParseError` saying "End of Central Directory not found".
- `Package.fillTemplate` / `patch` / `compileTemplate` (and the deprecated `Io.patchDocument` / `compileTemplate`)
  honor `password`, `securityPolicy` and `checkCrc32` on the read, and `modTime` / `securityPolicy` on the write.
  `{ securityPolicy }` was accepted and ignored.
- `Document.setWatermark(doc, undefined)` / `setBackground(doc, undefined)` remove the watermark / background. They
  refused `undefined`, so there was no in-place way to drop a watermark read from a file.
- The writer names the node path: a hand-built node it cannot write (an unknown `type`, a list that is not an array, a
  number with no XML form) is refused with `DocxArgumentError` whose `field` is the exact property path from the
  document (`body[3].children[0].content[1].type`); it was dropped, or failed with `TypeError` or `XmlError` (now the
  `cause`). The same holds for `Document.toBuffer`, `Document.writeFile`, `Package.*` and the streaming writer, whose `add` /
  `addAsync` name the path from `element` (`addMany` / `addManyAsync`: `elements[i]`). Paths name the model's
  properties, not element names: `body[0].properties.alignment` (was `body[0].paragraph.alignment`),
  `body[0].children[0].properties.underline`, `body[0].rows[0].properties.height.rule`,
  `body[0].properties.width.type` (was `body[0].table width type`), `sectionProperties.breakType`,
  `styles[3].paragraphProperties.alignment` (was `styles.styles[3]…`), `body[0].chart.legend` (was `chart.legend`),
  `headers["rId7"].content.children[0].properties.alignment` for a header, `glossary.blocks[1].content[0]…` for a
  glossary block (was relative to the block), `element.properties.alignment` for a streamed element. A run after a
  hyperlink counts from the paragraph's children (`children[1]`), and one inside it is `…nestedChildren[0].child…`.
  Other parts start from their own name: `settings`, `styles`, `coreProperties`, `appProperties`,
  `customProperties`, `sectionProperties`.
  For a number with no XML form in an attribute written directly, `field` may name the nearest object holding it
  (`body[0].properties`) and the message names the attribute.
- Those write-time messages state the problem and name no caller (`type must be one of …, got "bogus" (at
body[0].type)`, `thumbnail.contentType must be one of …`); they began `write:`, `Writing DOCX:` or
  `Streaming writer:`. The streaming writer's `options.images` problems name `Streaming.createWriter` and
  `options.images[i]`.
- The deprecated `Query.walkDocument` / `walkBlocks` forward to `Query.each` / `eachBlock`, so they check their
  arguments as those do. The deprecated flat `decryptPackage` checks its password and cap as
  `WordCrypto.Agile.decryptPackage` does and keeps its positional `maxDecryptedSize`.
- `Document.addStyle` with a `styleId` the document has replaces that style in place — the way to customize a style
  `useDefaultStyles` added. `Document.addCustomProperty` with an existing name replaces it. Both wrote a duplicate.
- `Ole.add` gives an embedding whose relationship id or part path is taken unused ones (read the result's
  `oleObjects`). `Document.toBuffer` failed with `Relationship ID "rIdOle1" already exists`.
- `Document.read` reads a Flat OPC document (what `Document.toFlatOpc` writes) passed as bytes or as the string itself, under the
  same security limits; it failed with "End of Central Directory not found". With `{ base64: true }` it reads base64
  text, as `Workbook.read` does.
- Glossaries read from a file list their entries in `glossary.blocks`, so `Glossary.findBlock` / `listBlocks` /
  `listAutoTextEntries` / `listQuickParts` work on them; `blocks` was always `[]`. Entries in galleries
  `BuildingBlockGallery` does not name are kept but not listed, and Word's built-in/custom pairs read as one gallery
  (`autoTxt` + `custAutoTxt` → `autoText`, `docParts` + `custQuickParts` → `quickParts`). An unedited glossary is
  written back byte for byte; edited `blocks` are now written, where they were dropped.
- `Build.checkbox` blocks read back as a `CheckBox` (`{ type: "checkBox", checked, … }`), not a paragraph of opaque
  markup. A checkbox in a paragraph with other content is still kept as it was.
- `Package.readPart` finds a part with or without the leading `/`, and `Package.listParts` lists names without it.
- Style chains (`Query.getTableStyle`, `getParagraphStyle`, `getRunStyle`, the deprecated `resolve*Style`) list only
  styles the document defines; `getTableStyle(doc, "Nope").chain` was `["Nope"]`.
- `Template.createXmlDataSource`'s `rootTag` names the element whose content is the data, as its TSDoc always said.
- `Security.DEFAULT_SECURITY_POLICY` / `STRICT_SECURITY_POLICY` are frozen; assigning to them changed the defaults
  of every later `Document.read`.
- Error messages name the live API: `DocxEncryptedError` names every reader that takes `{ password }`, and
  `WordCrypto.encrypt`'s errors say `WordCrypto.encrypt`.
- A chart read from a file and written back loses nothing. Every element of the chart part comes back — what `Chart`
  types through its fields, everything else as it was read: scheme colours and every DrawingML colour transform
  (`lumMod`, `satOff`, `hueMod`, `comp`, `redMod`, …), gradients, line styles and effects in any `c:spPr`, text
  formatting (`c:txPr`), layouts, titles (including one with no text of its own), gridlines, axis scaling, number
  formats, tick marks, crossings and ids, data labels down to each `c:dLbl`, markers, `c:invertIfNegative`, trendline
  and error-bar detail, the series' data references and caches (`Tabelle1!$D$2:$D$5` stays; it became
  `Sheet1!$B$2:$B$5`), `c:view3D` / floor / walls, `c:dTable`, `c:pivotSource` / `c:pivotFmts`, `c:lang`,
  `c:dispBlanksAs`, `c:showDLblsOverMax`, `c:printSettings`, `c:extLst`, `mc:AlternateContent` and the root's
  namespaces and `mc:Ignorable`, every chart group (a combo chart's second group and its series were dropped; they
  now read as series with `chartType` / `plotOnSecondaryAxis`), every axis (a series axis too), and the part's
  relationships (its style and colour parts, embedded workbook and user shapes, which were left unreferenced).
  A field that is set rewrites only the element it governs — `values` the series' `c:val`, `valueAxis.min` the
  axis' `c:min` inside its kept `c:scaling`, `title` the title's `c:tx`, `legend` the `c:legendPos`, `color`,
  `pointColors[i]`, `plotAreaColor`, `chartAreaColor` the fill they name — and the rest stays as read.
  `rawXmlPolicy: "strip"` writes the fields only; `"reject"` refuses a read chart with `DocxRawXmlPolicyError`. A
  relationship whose target is not written (an embedded workbook under `preserveOleObjects: false`) is dropped
  together with the element that referred to it. A chart part the model cannot read (an `c:ofPieChart` alone) is
  kept as it was, with its relationships; it was dropped, leaving the drawing pointing at nothing.
- `ChartSeries.pointColors` is by point index, as the writer always wrote it (entry `i` → `c:dPt` `c:idx` `i`): a
  read keeps each colour at its point's index, leaving holes for points without one. It read the colours compactly,
  so a file colouring only point 3 read as `["FF8800"]` and was written back colouring point 0. The writer writes a
  `c:dPt` for each entry that is set and skips holes and `undefined` entries, which `Build.chart` and the write-time
  checks now accept.
- `WordExcel.toChartSvg` / `toChartModel` and `Pdf.fromDocx` draw a chart's colours: the kept formatting above with
  its transforms (scheme colours from `themeColors` / the document theme), else `color`, `pointColors`,
  `plotAreaColor` and `chartAreaColor`. Every series was drawn in the automatic palette whatever it said.
- A series with `plotOnSecondaryAxis: true` is written on secondary axes even when no series sets `chartType`; it was
  written on the primary ones. A Word chart's style number is read from Word's `mc:AlternateContent`-wrapped
  `c14:style` too.
- Chart parts are written in schema order: a line series' `c:marker` precedes its `c:dPt`, `c:trendline` /
  `c:errBars` precede the series data with `c:smooth` after it, and `c:scaling` has `c:max` before `c:min`.

**Now refused** — `DocxArgumentError` naming the field, unless noted:

- Dates must be an `xsd:dateTime` (`new Date().toISOString()` gives one): `Document.addComment`,
  `Build.insertedRun` / `deletedRun` / `movedFromRun` / `movedToRun`, `Build.moveFromRangeStart` /
  `moveToRangeStart`. `"not a date"` was written as `w:date="not a date"`. `Document.setCoreProperties` needs string
  text fields and valid `created` / `modified` dates; an invalid date failed inside `Document.toBuffer` with `ArchiveError`.
- `Build.chart` checks `type`, each series' `name`, `categories`, `values` (finite numbers), colors and `chartType`.
- `Template.bindChartData` checks every binding before applying any: finite values, and a `chartRef` of the right
  shape (`-1`, `0.5`, `NaN`, `""` are refused). A series/categories length mismatch is now this class, still a
  `DocxError`.
- `Document.setCompatibilityMode` outside 11, 12, 14, 15 (it silently wrote 15); `Document.addContent` /
  `insertContent` with an unknown `type` (dropped by the writer); `Document.replaceText` with a non-string
  `replacement` (`null` became `"null"`) or a `search` that is not a string or `RegExp`; `Document.mailMerge` with
  data that is not an object or `Map` of strings; `Document.acceptRevision` / `rejectRevision` with a non-integer id.
- `Transform.fillFormFields`, for a field the document has: a checkbox value that is not a boolean (`"false"` checked
  it), a drop-down value that names no entry (it became index 0, and an index past the entries was written), a text
  value longer than `maxLength`. Every value is checked before anything is filled. A text field whose result the
  reader keeps as field pieces throws `DocxUnsupportedFeatureError` instead of being skipped.
- `Document.setHeader` / `setFooter` with a key that is not a type, a stored header's key or a referenced key
  _(owner decision)_; it was written as a header no page showed.
- `Package.edit` with an unknown edit type or missing fields (ignored or a raw `TypeError`); `Package.patch` /
  `patchTemplate` with malformed patches.
- `Security.resolvePolicy` and every `securityPolicy` option: a limit must be a number ≥ 0 (`Infinity` lifts it;
  `NaN` switched it off), flags booleans, `rawXmlPolicy` one of its values. A bad `securityPolicy` given to `Document.read`,
  `Document.toBuffer` and the rest is reported by that call (`Document.read: options.securityPolicy.maxEntries …`, field
  `options.securityPolicy.maxEntries`), not as `Security.resolvePolicy: policy.maxEntries`.
- `compressionLevel` outside 0–9 or an invalid `modTime` in `Document.toBuffer` / `toBase64` / `writeFile` / `toFlatOpc`,
  `Package.edit` / `patch` / `patchTemplate` / `fillTemplate` and `Streaming.createWriter`. These failed later with
  `ArchiveError`, the streaming writer only at `finalize()`. `StreamingDocxWriter#add` / `addText` refuse a non-block
  and a non-string; an unknown block was dropped. `addMany` / `addManyAsync` check the whole batch before writing any
  element.
- `Build.arrow` / `flowchartShape` / `callout` / `star` with an unknown direction, kind, style or point count; they
  wrote a shape with no geometry.
- Bytes that are not what they claim: `Font.subset` / `embed` / `embedFamily` need an `sfnt` font
  (`Font.subset(null)` returned `null`; the documented fallback for a real font that cannot be subset is kept);
  `WordCrypto.encrypt` needs a ZIP package; `Glossary.createBlock` needs a non-empty name, a known gallery and an
  array of blocks, `Glossary.createDocument` an array.
- Data sources (`Template.createJsonDataSource`, `createXmlDataSource`, `createCsvDataSource`,
  `createCompositeDataSource`): non-object JSON, malformed XML (it read as no data, leaving every placeholder
  unresolved), a CSV delimiter that is not one character, composite members that are not data sources. The errors
  are `DocxArgumentError` named `Template.create…DataSource`; they were `DocxError`.
  `createXmlDataSource("")` still gives no data.
- `Query.*`, `Convert.*`, `Font.hyphenate*` and `Diff.compare` with a non-document argument (a raw `TypeError`).
  Every one of the 121 functions whose first parameter is a document, glossary or style map refuses another value
  (field `doc`, `glossary`, `map`); `WordCrypto.hasSignatures({})` returned `false` and now refuses (field
  `source`). A `null` `Template.fillFromSource` source, `Styles.match` map, `Glossary.findBlock` glossary and
  `WordExcel.toChartSvg` / `toChartModel` chart are refused too.
- `WordExcel.toDocx` with a `sheets` name or position that matches no worksheet (field `options.sheets[i]`); it was
  skipped, and a list of unknown names gave an empty document.
- `Document.setSettings` values outside each field's form (a `zoom` that is not an integer ≥ 1, an unknown
  `zoomType`, a negative or fractional `defaultTabStop`, a `compatibilityMode` other than 11, 12, 14, 15, an unknown
  protection `edit` or hash, a non-boolean flag); `Document.setAppProperties` with non-string text fields or
  statistics that are not integers ≥ 0.
- Drawings (`Document.addFloatingImage`, `Build.floatingImage`, `Build.shape`, hand-built models): an unknown
  `relativeTo`, `align` or `wrap.style` / `side`, and a `wrap` that is not an object. `Build.chart` with an unknown
  `legend`.
- Header and footer children that are not paragraphs or tables; content that is not paragraphs in `addFootnote`,
  `addEndnote`, `addComment` and `addTextBox`.
- Watermarks with `fontSize`, `widthPt` or `heightPt` of zero or less, or empty `text` (they drew nothing visible).
- Decryption (`WordCrypto.decrypt`, `verifyPassword`, `Agile.decryptPackage`, `Agile.verifyPassword`,
  `Document.read({ password })`): a password that is not a string, and a `maxDecryptedSize` that is `NaN`, negative or not a
  number (`NaN` switched the cap off; `Infinity` is still "no limit"). `Font.obfuscate` / `deobfuscate` refuse a key
  that is not 32 hex digits (fields `data`, `fontKey`); non-hex digits were XOR'd as `0`, and the error was a plain
  `DocxError`. `WordCrypto.Agile.deriveKey` / `verifyPassword` / `decryptPackage` and `WordCrypto.Cfb.read` /
  `write` check their arguments; the deprecated flat `deriveEncryptionKey` and `readCfb` stay unchecked.
- `Streaming.createWriter` checks every option as the matching `Document.set*` does, when the writer is created
  rather than at `finalize()`, plus `onProgress` and `addText(…, properties)`; `Streaming.createReader` checks its
  source and options. A non-object `options` is refused by `Document.read`, `Validation.validate`, `WordHtml.fromDocx`,
  `WordMarkdown.fromDocx`, `Convert.toSemantic` and `Query.listTables`.
- Unknown keys in a security policy (`{ maxEntries: 1 }`), where a typo left the real limit at its default. Unknown
  keys elsewhere are still accepted, because callers spread objects read from files.
- `Theme.resolveColor`, before resolving anything: a string that is neither `RRGGBB` nor `"auto"` (`"bogus"`,
  `"#FF0000"`, `"F00"`) or a non-object (field `color`), a `ColorSpec` whose `val` or `themeColor` is not a string or
  whose `themeTint` / `themeShade` is not two hex digits (they became `NaN` and were applied), and a `theme` that is
  not an object. Other strings used to pass through unchanged. The internal resolver behind the HTML renderer gives
  the same answers but treats a malformed value read from a file as unresolvable.
- Error classes that changed: a duplicate `numId` in `Document.addListDefinition` (was `RangeError`); an unknown key
  in `Document.setDocDefaults` (was a plain `DocxError`); non-OLE2 bytes to `Vba.add` (was a plain `Error`).
  `Package.listParts` / `readPart` / `edit` and `Convert.fromOdt` report non-bytes as `DocxArgumentError` and
  non-ZIP bytes as `DocxParseError` (cause: the archive error), as `Document.read` does; they threw `ArchiveError`.
  `Convert.flatOpcToParts` reports non-XML input as `DocxParseError` (cause: the `XmlParseError`).
  `DocxInvalidImageError` extends `DocxArgumentError` and has a `field`.
- `field` is the path from the parameter name throughout Word: `Build.chart`, `Build.floatingImage` and `Build.shape`
  name `options.type`, `options.series[0]…`, `options.rId`, `options.fill.color` (were `type`, `rId`, `fill.color`);
  `Build.rect` and its siblings say `Build.rect: …` and `options.outline.width` (were `Build.shape: …` and
  `outline.width`); `Build.table(rows, options)` names `options.layout`, `options.columnWidths[0]` (were
  `properties.layout`, `columnWidths[0]`; the deprecated three-argument form is unchanged);
  `Document.addFloatingImage` and `addTextBox` name `options.wrap`, `options.width`; `Document.read` names `input` (was
  `buffer`); `Transform.fillFormFields` names `values["x"]` (was `values.get("x")`); the streaming writer's
  `addText` says `Streaming.StreamingDocxWriter#addText: …` and names `content`; an image media type from
  `WordMarkdown.toDocx`'s `resolveImage` says `WordMarkdown.toDocx: …` (field `options.resolveImage`), and one from
  `Template.fill` says `Template.fill: …` (field `data`). Messages name a built-in value (`got a Map`, `got a Date`,
  `got a Uint8Array`, not `got an object`), print booleans as values, and truncate a long rejected string.

Unchanged, as in v0.21.1: `Template.bindChartData` skips a well-formed `chartRef` that names no chart;
`Transform.fillFormFields` skips a name no field has, so one set of values fits many forms; a `rootTag` naming no
element is ignored by `createXmlDataSource`; fields missing from `mailMerge` data are left as they are.

**New names and deprecations**

- `WordHtml.toDocx(html)` returns a complete `DocxDocument` with its styles and list numbering, and
  `WordHtml.toDocxBody(html)` returns a `DocxHtmlBodyResult`
  (`{ body, images, abstractNumberings, numberingInstances, styles, docDefaults }`). There is no
  `WordHtml.toDocument`. The deprecated flat `htmlToDocx` (`{ body, images }`) and `htmlToDocxBody`
  (`BodyContent[]`) keep their old results. A body's ids are fixed; to splice HTML into an existing document, use
  `Transform.merge([host, WordHtml.toDocx(html)])`. A picture they cannot embed is the italic grey `[Image: alt]` run
  `WordMarkdown.toDocx` writes; it was an image node with an empty `rId`, written as its alt text. The deprecated flat
  forms keep the old placeholder.
- `WordExcel.toChartSvg` and `WordExcel.toChartXlsx` (it returns `.xlsx` bytes) are the names of the Word-chart
  renderers; the deprecated `renderWordChartSvg` and `generateChartEmbeddedXlsx` forward to them.
- `WordCrypto.Agile.decryptPackage(pkg, info, password, { maxDecryptedSize })` takes the cap in `DecryptOptions`, as
  `WordCrypto.decrypt` does. Both default to `DEFAULT_SECURITY_POLICY.maxPackageSize` (500 MiB), the cap `Document.read`
  applies; the deprecated `decryptDocx` keeps 512 MiB.
- `Convert.flatOpcToParts`, the inverse of `Convert.partsToFlatOpc`; the deprecated `Convert.parseFlatOpc` forwards
  to it.
- `WordCrypto.hasSignatures(partNames)`: check a file from its part names without reading the document,
  `WordCrypto.hasSignatures(await Package.listParts(bytes))` — what the deprecated `hasDigitalSignatures` did.
- `DocxArgumentError` (`documonster/word`, and also `documonster/word/crypto` with `DocxDecryptionError` and
  `isDocxError`). `documonster/word/html`, `/word/markdown` and `/word/excel` export `DocxError`,
  `DocxArgumentError`, `isDocxError` and `BaseErrorOptions` (markdown also `DocxInvalidImageError`).
  `DiscardUnrecognizedValuesOptions` is a flat type export, as in Excel.
- `DocxWriteOptions` (the options of `Document.toBuffer` / `toBase64` / `writeFile`), `Build.SimpleTableOptions`,
  `PreservedLink`, `PreservedLinkRId`, `HtmlBodyResult`, and `DocxRawXmlPolicyError`, which was thrown but could not
  be imported. `FillFromSourceOptions` and `LegacyTemplateOptions` are exported, deprecated, for the deprecated
  functions that take them; `mergeArrays` never had an effect — pass it to `createCompositeDataSource`.
- Fields: `CreateShapeOptions.flipHorizontal` / `flipVertical`, `TextWatermark.widthPt` / `heightPt`,
  `DropDownFormField.result`.
- Units, now stated in TSDoc: `TextWatermark.fontSize` is in half-points (written as `font-size:<n/2>pt`; since 1pt
  is Word's auto-size marker, `fontSize: 2` reads back as absent); `defaultFontSize` / `codeFontSize` of
  `WordHtml.toDocx*`, `WordMarkdown.toDocx*` and the layout are half-points, while `Pdf.fromDocx`'s
  `defaultFontSize` is points.
- Deprecated: `Build.drawingShape` → `Build.shape`, which writes the same shape (`fillColor` →
  `fill: { type: "solid", color }`, `noFill` → `fill: { type: "none" }`, `outlineColor` / `outlineWidth` /
  `noOutline` → `outline: { color, width, noLine }`, `textContent` → `textBody: { paragraphs }`);
  `CreateShapeOptions.flipH` / `flipV` → `flipHorizontal` / `flipVertical`; `DocxOptions` → `DocxWriteOptions`;
  `AssetRef` and `MediaAsset`, which nothing takes or returns.
- Deprecated: numeric positions in `WordExcel.toDocx`'s `sheets` — pass names. Numbers are 0-based here, 1-based in
  `Pdf.fromExcel` and sheet ids in `Workbook.getWorksheet`; they work as before until 1.0. The deprecation lives in
  three overloads: `sheets` as names; a `@deprecated` one reached by a list that certainly carries a number (a
  literal with a number, a variable inferred as `number[]`), which editors strike through; and the published
  `ExcelToDocxOptions`, unchanged and not deprecated, so a variable annotated with it is not flagged.
- Deprecated: the positional forms of `Build.mathSum` / `mathIntegral` / `mathProduct` / `mathNary` (`sub`, `sup`) and
  `Build.mathBar` (`position`) → `mathSum(c, { sub, sup })` (`Build.MathNaryOptions`) and
  `mathBar(c, { position: "bottom" })` (`Build.MathBarOptions`).
- Widened: `ChartBinding.series` / `categories`, `ChartSeriesData.values`, `ChartTemplateData.charts`, the
  `columnWidths` of `Build.table` / `Build.simpleTable` / `Document.addTable` take readonly arrays;
  `Convert.partsToFlatOpc` takes `ReadonlyMap`s. Name → value data is a plain object or a `Map` in
  `Document.mailMerge`, `Template.fill`, `Package.fillTemplate`, `Transform.fillFormFields` and
  `Transform.resolveDataBindings` (each refused one of the two); `resolveDataBindings` now checks that its values are
  strings. `Query.getTableStyle` returns the named `ResolvedTableStyle`, and `ResolvedNumberingLevel.format` /
  `justification` are typed `NumberFormat` / `LevelJustification`.
- Hash algorithm names: `Security.protect`, `Document.setSettings({ documentProtection })`, `WordCrypto.encrypt` and
  `WordCrypto.Agile.deriveKey` accept both `"SHA-256"` (canonical — WebCrypto's spelling, what `Security.getState`
  reports) and the encryption format's `"SHA256"`; each refused the other spelling. `EncryptOptions.hashAlgorithm`
  and `DocumentProtectionOptions.hashAlgorithm` are widened to both; prefer `"SHA-256"`. `Security.protect` stores,
  and `Security.getState` reports, the canonical spelling; an `EncryptionInfo` part holds the format's spelling
  either way. Bug fixes on the way: a protection holding `"SHA256"` was written with `cryptAlgorithmSid` 4 (SHA-1),
  so Word could not check its password, and a protection read as MD2 / MD4 / MD5 or an unknown id was written back
  as SHA-1 — each is now written with the id it was read with.
- `Security.verifyPassword(doc, password)` is the protection-password check, no longer deprecated; the interim
  `Security.verifyProtectionPassword` is removed (it was never released). It now refuses a non-string `password`
  with `DocxArgumentError`. The file-encryption password is checked by `WordCrypto.verifyPassword`.
- `DigitalSignatureInfo.claimedSigner` holds the signature's display name (`SignatureText`), the name PDF's report
  uses for an unverified identity; `signer` is deprecated and holds the same value until 1.0.
- New: `Document.listUnrecognizedValues(doc)` lists the values a read kept outside their vocabulary, as they are kept
  now, each a `DocxUnrecognizedValue` (`{ path, value }`, exported); `Document.discardUnrecognizedValues`'s `filter`
  takes the same type.
- Argument errors describe a `Date` that holds no time as `an invalid Date` (they read "must be a valid Date, got a
  Date").
- `RenderOptions.background` (`Layout.toSvg`, `Layout.pageToSvg`, `Layout.computedPageToSvg`) takes what `Draw.toSvg`,
  `Mermaid.toSvg` and the chart renderers take — a CSS color or `{ r, g, b, a }`, checked the same way
  (`DocxArgumentError`, `field: "options.background"`); a fully transparent color, `transparent` or `none` draws no
  background. Deprecated: `backgroundColor`, which still works as before (`background` wins when both are given).

#### PDF, drawing and Mermaid

**Output changes**

- `Draw.parseColor` (and the deprecated `cssColour`) reads the 148 CSS named colors (case-insensitive) and `hsl()` /
  `hsla()`, which read as black. `none` and `transparent` are fully transparent, not opaque black. The same parser
  serves the PDF `drawSvg` importer and the raster canvas, and Mermaid themes and `classDef … fill:orange` now draw
  in the named color.
- `Mermaid.toSvg(src, { background })` (and `theme.background`) draws the backdrop once, as the display list's first
  node, in the requested color. It was drawn twice, and the list's own copy read `"white"` as black — the copy raster
  and PDF output use. `background` remains and still wins over `theme.background`, `"transparent"` included.
- A Gantt date that does not exist (`2024-13-45`, `2023-02-29`) is no longer rolled into another date: the task is
  left out and reported in the diagram's new `skipped`, as is a dependency on a missing task. A year below 100 reads
  as that year, not 19xx. A state diagram's bare `Name` / `state Name` declares a state (it was dropped), and
  `create participant X` declares the participant.
- `Pdf.Editor`'s full rewrite (`toBuffer()`, `splitPages`, `sign`, the deprecated `save`) keeps every `/Info` entry.
  It rebuilt `/Info` from five fields, losing `/Keywords`, `/ModDate`, `/Trapped` and custom entries, and overwrote
  `/Producer`. The editor invents nothing: neither a full rewrite nor an incremental update adds a `/Creator` or
  `/Producer` the document did not have (they defaulted to `documonster`, which the XMP packet did not say); they
  change only when `setMetadata` gives them. The builder still writes `documonster` as `/Producer` when none is given.
- A full rewrite and `copyPagesFrom` write each page's `/MediaBox` as the document has it, origin included, and carry an
  inherited `/Rotate` and `/CropBox` onto the page. `[0 0 w h]` moved every page whose box does not start at the
  origin, and a document landscape by inheritance came back portrait. Streams keep their filters (every JPEG lost
  `/DCTDecode`); those with a `/Crypt` filter are still decoded. XMP metadata streams are never compressed.
- `copyPagesFrom` keeps annotations and form fields; they were dropped. Each copy gets its own annotations, `/P`
  names the copy, references to other source pages become `null`, and the page's fields join the target's
  `/AcroForm /Fields`. `formFields: "drop"` copies the pages without their form fields (links and comments are still
  copied) — the way to merge the same form twice.
- Fields an editor adds are real fields: `EditorPage#addFormField` on an existing page wrote a widget no form listed,
  and annotations and fields on an `addPage()` page were dropped. An incremental update, which did not write added
  annotations or fields, falls back to a full rewrite for them (`incremental: "required"` throws).
- `Pdf.extract` reads XMP as XML: character references are decoded, attribute-form properties and a language
  alternative's `x-default` are read, only top-level `rdf:Description`s count, and `rdf:Seq` / `rdf:Bag` entries (such
  as several `dc:creator`s) are joined with `"; "`. UTF-16 and UTF-8-BOM packets are decoded; a packet that is not
  namespaced XML falls back to the previous lenient scan.
- `Pdf.fromExcel` shows a cell without a number format in "General", as `Cell.getDisplayText` and
  `WordExcel.toDocx` do, instead of `toLocaleDateString()` / `String()`: dates no longer depend on the machine's time
  zone or locale.
- `Pdf.fromTable` (and the deprecated `Pdf.create`) renders a `Date` cell by the same rule: UTC calendar fields,
  `yyyy-mm-dd` at midnight UTC, otherwise `yyyy-mm-dd hh:mm:ss`. It used `toLocaleDateString()`, which showed no
  time and gave a different day west of UTC.
- Filling a checkbox or radio group sets the widget's `/AS` and writes `/V` as a name (on the parent field for a
  radio group), in `toBuffer`, incremental updates, `sign` and `save`. The box was drawn unchecked and a radio
  group's value never reached the group.
- A JPEG is written in the color space its frame header declares (`/DeviceGray`, `/DeviceRGB`, `/DeviceCMYK`, with
  the Adobe-inverted `/Decode` for Adobe CMYK); it was always `/DeviceRGB`. A PNG drawn on an existing page in an
  incremental update is embedded as decoded pixels with its soft mask; the file bytes were written as pixels.
- The builder copies what it checked: changing an object after passing it to `setMetadata`, `setEncryption`,
  `sign`, `drawImage`, `addAnnotation` or `addFormField` no longer reaches the file.
- `Pdf.extract` with `extractMetadata: false` reports `metadata.pageCount` and `metadata.encrypted`; it said
  `pageCount: 0`.
- A missing or wrong password is a `PdfPasswordError` (a `PdfStructureError`, same messages) from `Pdf.extract`,
  `Pdf.Editor.read`, `copyPagesFrom` and `Pdf.verifySignatures`. The raster errors (`scale`, `samples`, output
  size) are `DrawArgumentError` / `DrawError`, not plain `Error`. Errors that were plain `Error`s now have a class:
  the `Pdf.fromDocx` "unhandled content" guard is a `PdfRenderError`, the `Pdf.fromChart` "no model" guard a
  `PdfArgumentError`, the chart PNG browser-path failures and the chart cache-fill invariant `ExcelError`s, and the
  PDF font plan's invariants `PdfFontError`s. Messages that began with an internal name (`signPdf:`,
  `buildIncremental:`, `renderLayoutDocumentToPdf:`, `chartToPdf:`, `resolveTopologyObject:`) describe the problem
  or start with the public call.
- Text that the standard faces cannot encode (CJK, most non-Latin scripts), drawn on an editor page, is embedded
  with a system face found on the host, as `Pdf.Builder` does; it was drawn with the Type3 fallback as `.notdef`
  boxes, while `Pdf.extract` still read it back. Pass `disableFontAutoDiscovery: true` to `Pdf.Editor.read` for
  host-independent output (see the new font options below).
- `Pdf.fromExcel` writes the workbook's own keywords when the `keywords` option is not given; it wrote none.
- `Pdf.fromDocx` writes the document's own keywords (`coreProperties.keywords`) when the `keywords` option is not
  given; it wrote none. `Pdf.fromDocx` and `Pdf.fromChart` write `/Creator` `documonster` unless `creator` says
  otherwise, as `Pdf.fromExcel` always has; they wrote none.
- A chart draws its chart area — the chart's own fill and border — from the model: `c:chartSpace/c:spPr` for a
  classic chart, `cx:chartSpace/cx:spPr` for an Office 2016 one, `noFill` and `ln/noFill` respected. Where the model
  says nothing it is what Office draws for an automatic chart area: a `bg1` fill (`dk1` for chart styles 41–48) and a
  0.75 pt `#D9D9D9` border (`tx1` at 15 % / 85 % luminance, which is also what Excel writes out), drawn inside the
  chart's box; a classic chart whose `c:roundedCorners` is absent (or true) has rounded corners, as Excel shows it
  — the library writes `val="0"`, so a chart it builds is square. It is the first node of `ChartClassic.toDrawList` /
  `ChartEx.toDrawList`, so `toSvg`, `toPng`, `Pdf.fromChart`, `Pdf.fromExcel` and `Pdf.fromDocx` all show the same one
  _(output change: SVG and PNG previews gain the border; they drew a white fill only)_. The PDF paths and the deprecated
  `drawChartPdf` / `drawChartExPdf` no longer draw their own frame — a white panel with a fixed grey border
  (`#CCCCCC` for a classic chart, `#D9D9D9` for an Office 2016 one) whatever the chart said. `ChartRenderOptions.background` replaces the chart area's fill;
  its border stays the model's. `ChartScene` gains `chartArea`.
- Chart previews apply DrawingML color transforms to every element — series fills and lines, data points, markers,
  gridlines, axes, axis and data labels, titles, legends, data tables, trendlines, error bars, effects, the chart area
  and Office 2016 charts — through one resolver (`resolveChartColor`) and one transform engine shared with Excel cells
  and Word (Office's semantics: `lum*` / `sat*` / `hue*` in HSL,
  `tint` / `shade` / `inv` / `gamma` in linear RGB). They were ignored: an Excel-saved gridline or axis
  (`tx1` lumMod 15 % / lumOff 85 %) was drawn black instead of `#D9D9D9`, axis labels (`tx1` 65 % / 35 %) black instead
  of `#595959`. A gridline or axis line with `<a:ln><a:noFill/>` is no longer drawn in a default grey. Scheme colors
  resolve against the workbook's theme in `Chart.toSvg` / `Chart.toPng`, `Pdf.fromChart` and `Pdf.fromExcel` — a
  workbook saved by a recent Excel has the 2023 Office theme (`accent1` = `#156082`), not the 2013 one the preview
  assumed (`#4472C4`) — and against the new `ChartRenderOptions.themeColors` for a model rendered on its own. A
  chart's automatic series colours are its theme's accents too (they were the default theme's whatever the workbook
  said). Word charts follow the document's theme (`doc.theme.colorScheme`) in `Pdf.fromDocx`, in
  `WordExcel.toChartSvg(chart, { themeColors })` (new `DocxChartSvgOptions`, checked as `DocxArgumentError`) and, through
  the new second argument of `RenderOptions.renderChart` (`RenderChartContext.themeColors`), in `Layout.toSvg` /
  `Layout.pageToSvg`.
- Excel cell tints (`Pdf.fromExcel` / `Pdf.fromTable` cell colors with `tint`, ECMA-376 §18.8.19) and Word's
  `themeTint` / `themeShade` (`Theme.resolveColor`, Word layout, HTML and PDF output) act on the HSL luminance, as Office
  does; they were an RGB mix toward white / black, and Word's `themeTint` the wrong way round — `themeTint="99"` (Word's
  "Lighter 40 %") came out as "Lighter 60 %". Neutral colors are unchanged; chromatic ones change (accent1 `4472C4`
  "Darker 25 %" is now `2F5496`, as in Word, not `335693`).
- `Pdf.fromChart({ title })` is the PDF's title only, on every path: an Office 2016 chart on the vector path drew it
  as the chart's visible title instead of its own, while a classic chart and the raster path kept theirs _(bug
  fix — a ChartEx chart's visible title changes)_.
- `Pdf.Editor` page indices are one model: the pages added with `addPage()` / `copyPagesFrom()` are numbered after the
  pages the document was read with (as `splitPages` already numbered them), and `getPage`, `removePage`,
  `rotatePage` and `isPageRemoved` take them as well — each of these threw `PdfStructureError` for such an index.
  `getPage` returns a handle that draws onto an added page (over a copied page's content); `removePage` leaves it out
  of the output and frees its form field names; `rotatePage` sets its `/Rotate`. The builder `addPage()` returns now
  refuses drawing once its page is removed (`PdfArgumentError`), as a held page handle does. `pageCount` keeps its
  meaning (the pages read); the new `outputPageCount` is the number of pages a write produces.
- A font supplied to `Draw.toRgba` that cannot be read is still skipped, and is now reported in
  `RasterizedImage.textWarnings` (`fonts[0] could not be read (…) and was skipped.`); so is a `collectionIndex` past
  the end of its collection, which still draws with face 0.
- `Pdf.fromDocx` draws every chart, classic and ChartEx, with the Excel chart engine without a `chartRenderer`, so
  `chartRenderer: await Pdf.wordChartRenderer()` changes nothing (the output is byte-identical). The Word README and
  the `chartRenderer` TSDoc no longer say charts degrade to a placeholder without one.
- `Pdf.PageSizes` (now typed `Readonly<…>`) and `Matrix.IDENTITY` are frozen.

**Now refused**

- One rule for argument errors in `PdfArgumentError`, `DrawArgumentError`, `MermaidArgumentError` and
  `ExcelChartOptionsError`: `field` is the path from the parameter name (`options.fontSize`,
  `options.watermark.pages[0]`, `options.sheets[1]`, `metadata.title`, `encryption.ownerPassword`, `rect.width`,
  `source`, `data`). PDF and Mermaid used paths relative to the options object (`fontSize`, `pages[0]`,
  `theme.nodeFill`), and Mermaid named `diagram` for the parameter named `source`. Messages start with the public call:
  `Pdf.fromExcel: …`, `Pdf.Builder#addPage: …`, `Pdf.PageBuilder#drawText: …` (a page from `Pdf.Builder#addPage` or
  `Pdf.Editor#addPage`), `Pdf.EditorPage#drawText: …`, `Pdf.Editor#splitPages: …`, `Pdf.Editor.read: …`,
  `Pdf.Signature.verify: …`, `PdfContentStream#setFont: …`, `Draw.toRgba: …`, `Draw.Geometry.flattenPath: …`,
  `new Draw.Backend.BasicRasterCanvas: …`, `Mermaid.toSvg: …` / `Mermaid.toDrawList: …` (they shared one prefix),
  `ChartClassic.toSvg: …`, `ChartEx.createModel: …`, `Chart.toPng: …`.
- `ExcelChartOptionsError` extends `ExcelArgumentError` (it extended `ExcelError`), so `instanceof ExcelArgumentError` is
  now true as well and `field` names the option (`options.series[0].values`, `preset`, `table`, `rows[2].price`,
  `model`). Validator messages name the path from `options`, not from `chart`, `chartEx` or `combo chart`:
  `options.series[0].trendline is not valid for pie charts` (was `chart.series[0].…`).
- Chart render options (`ChartClassic.toSvg` / `toPng` / `toDrawList` / `Backend.toScene`, `ChartEx.toSvg` /
  `toPng` / `toDrawList`, `Chart.toSvg` / `Chart.toPng`, and the deprecated `drawChartPdf` / `drawChartExPdf`), with
  `ExcelChartOptionsError` before drawing: `width`,
  `height`, `scale` and `dpi` must be finite numbers > 0 (`toSvg` wrote `width="NaN"`, `width="0"` or a negative
  `viewBox`; `toPng` drew a 1-pixel-wide image, or for `NaN` threw a `DrawArgumentError` naming the internal
  `BasicRasterCanvas`); `title` must be a string, `background` a CSS color or `{ r, g, b, a }` with channels in
  0..1 and the deprecated `backgroundColor` a CSS color (a number failed with a `TypeError`, an unreadable string was
  drawn black); `fonts` an array of fonts and `useSystemFonts` a boolean; `regionMap` is `{ topology, objectName,
match?, projection?, strokeColor? }`, each checked. `drawChartExPdf` needs a surface with `drawRect` / `drawLine` /
  `drawText` and a `rect` of finite numbers with an area, and checks the model (messages start `drawChartExPdf:`).
- A `regionMap.topology` that cannot be read — not an object, no `objects`, an `objectName` it does not have, `arcs`
  not a list of point lists, a `transform` that is not two finite pairs, a geometry whose `arcs` or `coordinates` do
  not have its type's shape — is a `ExcelChartOptionsError` (from `Pdf.fromChart`, a `PdfArgumentError`) naming the exact
  path (`options.regionMap.topology.arcs[3]`, `options.regionMap.objectName`). It was drawn as the built-in centroid
  preview without a word, failed with a `TypeError`, or, for a bad `transform`, read the coordinates as absolute. A
  valid topology that matches no category still draws the preview, and a geometry of an unknown type is still
  skipped.
- `documonster/draw` (`DrawArgumentError`, with `.field`): every `Draw.*`, `Draw.Backend.*`, `Draw.Geometry.*` and
  `Matrix.*` function checks its arguments — a `NaN` coordinate, a color that is not an `Rgba01`, an unknown node
  `kind`, a bad text `anchor`, a `null` child were written as `0`, `fill="#NaNNaNNaN"` or a `TypeError`.
  `Draw.withAlpha` refuses an alpha outside 0..1; `Draw.wrapText` a `NaN` width (0 or less still means "do not
  wrap"); `Draw.Backend.BasicRasterCanvas` a size that is not a positive integer. `Draw.parseRasterFont` refuses
  bytes that are not a TrueType font and a `collectionIndex` past the end of a `.ttc` — it returned an empty face or
  face 0 — while the deprecated flat `parseRasterFont` keeps the forgiving behavior. `Draw.toSvg` / `Draw.toRgba`
  refuse a `background` string no CSS reading accepts (`Draw.toSvg` wrote it into `fill` as given); the deprecated
  flat `toSvg` and `rasterizeToRgba` keep their old, unchecked reading (a 0×0 list rasterizes to a 1×1 transparent
  image, a background string is written as given).
- `documonster/mermaid` (`MermaidArgumentError`, a `MermaidError`, with `field` such as `source`,
  `options.theme.nodeFill`, `options.palette[1]`, and `diagram` for functions that take a parsed diagram): `fontSize` and `maxLabelWidth` must be > 0, the gaps and `padding` ≥ 0;
  `theme` must be an object of CSS strings, an unknown key (Mermaid's `primaryColor`) is refused, and a color that
  cannot be read is refused as documented (`"zz"` was drawn black); `palette` must not be empty; `background` is
  checked like a theme color. `Mermaid.parse` refuses a non-string; `Mermaid.resolveTheme`,
  `Mermaid.createCenteredText` and the deprecated `layoutFlowchart` / `centredText` check their arguments. Syntax
  errors are a `MermaidParseError` (see below), and the parser itself stays lenient.
- `documonster/draw` handles: `Draw.Backend.fontHasGlyph` (and its flat forward) refuses a font without `getOutline`
  (field `font`), and `Draw.Backend.createRasterSurface` (and the flat `createRasterSurface`) a canvas without
  `fillRect` / `pushClip` (field `canvas`).
- `documonster/pdf` (`PdfArgumentError` unless noted):
  - An edit to a removed page: `Editor#getPage(i)` / `rotatePage(i)` on a removed page (field `index`), and every
    `draw*`, `addAnnotation`, `addFormField` and `getContentStream` call through an `EditorPage` held from before
    `removePage` (field `page`; `measureText` still works). The edit was accepted and written nowhere. An index out of
    range stays a `PdfStructureError`.
  - `drawImage` (builder and editor page), `fromTable` sheet images and image watermarks check the bytes against
    `format`: PNG bytes passed as `"jpeg"` were written as an image no viewer decodes. A bad `format` is now this
    class too.
  - A form-field name already in use, at the call: the builder's `addFormField` refuses a name used on any page (a
    clash across pages surfaced at `toBuffer`); an editor page's `addFormField` (also on an `addPage()` page) refuses
    a name the document was read with, one already added to a page still in the document, or one a copied page
    brought. A field on a removed page no longer holds its name. `copyPagesFrom` refuses a copied field whose fully
    qualified name is taken (field `sourcePdf`), before any page is queued, unless `formFields: "drop"` is passed.
  - `Pdf.Editor#setFormField` / `setFormFields` with a field the document does not have (it was ignored) or a
    checkbox or radio state its widgets cannot show; `setFormFields` checks every entry first. Dynamic XFA forms
    still throw `PdfXfaFormError` at write time.
  - `Pdf.fromTable` (and the deprecated `Pdf.create`): a cell that is not a string, finite number, boolean, valid
    `Date`, `null` / `undefined` or `{ value, … }` (objects vanished, `NaN` printed "NaN"); a `fontColor` /
    `fillColor` that is not hex (`"red"` printed black); a column width ≤ 0 (a negative one erased the text).
  - `Pdf.fromExcel` / `Pdf.fromTable` `sheets`: `0` is refused (field `options.sheets[i]`; positions are 1-based,
    unlike `WordExcel.toDocx`'s 0-based ones; numbers are deprecated, see below), and a name or position that matches
    no sheet is refused instead of skipped (it dropped the sheet, or reported "No sheets to export"). `sheets: []`
    still means every visible sheet.
  - `watermark.sheets` with a name that matches no sheet (field `options.watermark.sheets[i]`), as an unknown name in
    `sheets` is; it left every page without the watermark _(owner decision)_. A page number past the last page in
    `watermark.pages` or `Pdf.extract({ pages })` is still skipped.
  - `Pdf.Editor#splitPages({ pageIndices })` and `copyPagesFrom(source, { pageIndices })` with an index that is out of
    range, names a removed page or is not an integer (field `options.pageIndices[i]`), before anything is written or
    copied _(owner decision)_. `splitPages` returns one part per index, in order. The deprecated positional forms
    still skip an index out of range.
  - The `fonts` option has one check at every entry point — `Pdf.fromTable`, `Pdf.fromExcel`, `Pdf.fromDocx`,
    `Pdf.fromChart`, `new Pdf.Builder`, `Pdf.Builder#embedFonts` (field `config…`) and `Pdf.Editor.read` — which
    names the path: `options.fonts.default.regular`, `options.fonts.families[0].name`,
    `options.fonts.fallbackFamilies[0]` (a family that is not configured), and a name or alias used twice.
    `Pdf.fromTable`, `Pdf.fromExcel` and `new Pdf.Builder` threw `PdfFontError` with no `field`; bytes that are not a
    font still do. `font` and `fonts` together is refused (field `options.fonts`), before `Pdf.fromExcel` runs
    `recalculate`.
  - `Pdf.fromDocx` needs a `body`, and checks `margins`, `defaultFontFamily`, `defaultFont`, `chartRenderer` and the
    document options. `Pdf.fromExcel` refuses anything that is not a workbook (field `workbook`; `{}` failed with a
    `TypeError`) and a `recalculate` that is not a function. `Pdf.fromChart` refuses a non-chart (field `chart`) and
    checks `forceRaster`, `regionMap` and the document options. The document options (`PdfDocumentOptions`, below) have
    one check at all four converters — metadata strings and valid dates, `encryption.ownerPassword`, the font options;
    `Pdf.fromDocx` and `Pdf.fromChart` also refuse metadata text XML cannot carry, as `setMetadata` does.
    `Pdf.parseSvgPath` refuses a non-string (field `d`), and the deprecated `createPdfDrawSurface` checks its page as
    `Pdf.createDrawSurface` does.
  - `Pdf.Builder`: `onWarning` must be a function, `embedFont` (builder and editor) bytes, and `setPdfACompliance`
    refuses anything but `"1b"`. `addBookmark` with no bookmark yet says so; it is still a `RangeError`. `embedFonts`
    with a non-object is this class (field `config`; it was `PdfFontError`), and `drawText` checks `bold` / `italic`.
  - `setMetadata` (builder and editor) refuses text XML 1.0 cannot carry (C0 controls other than tab, LF and CR,
    U+FFFE / U+FFFF, lone surrogates), naming the field, and a non-object (field `metadata`).
  - `Pdf.extract` checks `data` (an `ArrayBuffer` is reported as such, not as a corrupt file), `options`,
    `password`, `pages` (integers ≥ 1; `[0]` extracted nothing) and the `extract*` booleans. A `maxDecodedBytes` that is
    not a number `>= 0` (or `Infinity`) is a `PdfArgumentError` (field `options.maxDecodedBytes`, message
    `Pdf.extract: …`); it was a `RangeError` with no field _(the class changes)_. `Pdf.Editor.read`, `copyPagesFrom` and
    `Pdf.verifySignatures` check their bytes and options.
  - `copyPagesFrom` / `splitPages` given `pages` — the option is `pageIndices`, 0-based — so the old spelling cannot
    silently copy every page.
  - `Pdf.Builder#sign` / `Pdf.Editor#sign` check the certificate (DER X.509) and key (DER PKCS#8, RSA) at the call;
    they failed inside `toBuffer()`. `Pdf.Signature.fillPlaceholder`, `buildDictPlaceholder`, `parseAsn1` and
    `verify` check their arguments, and `verify` refuses a `byteRange` that is not four integers ≥ 0 (a `NaN` was
    reported as "content was modified"). Truncated ASN.1 and indefinite lengths are a `PdfStructureError`.
  - `PdfContentStream`: `setFont`, `setGraphicsState` and `doXObject` refuse a name that is not a plain PDF name
    (`setFont("F1) Tj (", 12)` injected operators); `setLineCap`, `setLineJoin` and `setTextRenderingMode` refuse
    values outside 0–2.
  - `Pdf.createDrawSurface` refuses a rect with no area and a `page` that is not a page builder. The list is placed
    at the rect's top-left and scaled by `scale`; nothing is clipped, and `rect.width` was never used (now
    documented).

**New names and deprecations**

- `Pdf.Editor#setMetadata(update)` takes the new type `PdfMetadataUpdate` — every `DocumentMetadata` field, each also
  `| null` — as a JSON merge patch (RFC 7396): an absent or `undefined` field keeps the document's value, a value sets
  it, and `null` removes it, deleting the `/Info` entry and the XMP property from every top-level description in
  every form (element, with its `rdf:Alt` / `rdf:Seq`, or attribute). Successive calls merge, a later value or `null`
  winning. The values are checked as `Pdf.Builder#setMetadata` checks them and copied at the call; it works on full
  rewrites and incremental updates (an update with only a metadata change appends just the Info and XMP objects).
  The XMP packet is parsed as XML and edited in place — only top-level descriptions, element or attribute form kept,
  absent properties added with their namespace bound, only the `x-default` entry of an `rdf:Alt` replaced,
  `dc:creator` written from `author` split on `;`, `xmp:MetadataDate` updated for any change — and written
  uncompressed in its own encoding; one that cannot be decoded or is not XMP is kept byte for byte and only `/Info`
  changes. `Pdf.Builder#setMetadata` and `DocumentMetadata` are unchanged.
- `DocumentMetadata.producer` (builder and editor) sets `/Producer` and `pdf:Producer`;
  `EditorCopyPagesOptions.formFields` (`"copy"`, the default, or `"drop"`).
- `PdfMetadataUpdate`; `PdfPasswordError`; `PdfSignaturePlaceholderMissingError`, thrown by `Pdf.Signature.fillPlaceholder` (and the
  deprecated `Pdf.sign`) when there is no placeholder; `EditorSplitOptions`; `ReadPdfPage.rotation`;
  `Pdf.EditorPage#drawSvg`; types `PdfCellValue`, `PdfHeaderFooterOptions`, `TextFragment`.
- `DrawError`, `DrawArgumentError`, `isDrawError` from `documonster/draw`; `MermaidArgumentError`,
  `MermaidSkippedStatement`, `MermaidParseReport` and `skipped?` on every diagram from `documonster/mermaid` —
  `Mermaid.parse` reports what it skipped (`{ text, line, column }`) and still does not throw.
- `Pdf.Editor#splitPages({ pageIndices, encryption, invalidateSignatures })`; the positional `splitPages(pageIndices)`
  is deprecated. `EditorCopyPagesOptions.pageIndices` and `PdfSignatureReport.reason` (matching
  `PdfSignatureOptions.reason`) are the names; the unreleased `pages` and `signingReason` are gone.
- Deprecated: numeric positions in `Pdf.fromExcel`'s `sheets` — pass names. Numbers are 1-based here, 0-based in
  `WordExcel.toDocx`, and sheet ids in `Workbook.getWorksheet`; they work as before until 1.0. As in
  `WordExcel.toDocx`, three overloads carry it: names; a `@deprecated` one reached by a list that certainly carries
  a number; and the published `ExcelToPdfOptions`, unchanged and not deprecated.
- Deprecated: `SignatureVerificationResult` (use `PdfSignatureCheck`), `CmsSignedData` and `SignOptions`, which no
  public function uses.
- Deprecated: `Pdf.wordChartRenderer` — omit `chartRenderer`: the default draws every chart with the Excel engine. It
  still returns the same callback. There is no `Pdf.createWordChartRenderer`.
- `new ChartOptionsError(field, message, options?)`; the `(message, options?)` form still works, deprecated, and
  leaves `field` as `""`.
- `MermaidParseError` is the name of Mermaid's syntax error — the `*ParseError` every module's parser throws — with
  `line` and `column` (1-based) locating the statement that names the diagram (`undefined` for an empty source). It
  takes `(message, { line?, column?, cause? })` (`MermaidParseErrorOptions`). Deprecated: `MermaidSyntaxError`, which
  is the same class object, so `instanceof` either answers alike; the error's `name` is now `"MermaidParseError"`.
- `PdfDocumentOptions`: the document-level options `Pdf.fromTable`, `Pdf.fromExcel`, `Pdf.fromDocx` and
  `Pdf.fromChart` all take — `title`, `author`, `subject`, `keywords`, `creator`, `creationDate`, `modDate`,
  `encryption`, `fonts`, `preferSystemFonts`, `disableFontAutoDiscovery`, `textLanguage`, `onWarning`.
  `PdfExportOptions`, `DocxToPdfOptions` and `ChartToPdfOptions` extend it, so `Pdf.fromDocx` and `Pdf.fromChart` gain
  metadata, fixed dates (byte-reproducible output) and encryption. `Pdf.fromDocx` takes each metadata field it is not
  given from the document's core properties (`title`, `creator`, `subject`, `keywords`), as `Pdf.fromExcel` takes the
  workbook's.
- `DocxToPdfOptions.margins` (`{ top, right, bottom, left }` in points, as `Pdf.fromExcel` takes it) and
  `DocxToPdfOptions.defaultFontFamily`. Deprecated: `marginTop`, `marginRight`, `marginBottom`, `marginLeft` and
  `defaultFont`, which keep working; the new name wins when both are given.
- `ChartRenderOptions.themeColors` (`ChartThemeColors`: six hex digits per theme slot, `dk1` … `folHlink`), the
  theme the chart's scheme colors resolve against; checked (field `options.themeColors.<slot>`).
  `ChartClassic.Backend.toSvgFilter(id, effects, { themeColors })` takes it too.
- One background check: `Draw.toSvg` / `Draw.toRgba`, the chart renderers and Word's `Layout.toSvg` /
  `Layout.pageToSvg` refuse the same `background` values with the same field and words, each in its own error class.
  `Draw`'s message for a background that is neither a string nor an object now reads like the others
  (`must be a CSS colour (…) or { r, g, b, a }`).
- `ChartRenderOptions.background` takes what `Draw.toSvg` / `Draw.toRgba` take — a CSS color or `{ r, g, b, a }`;
  deprecated: `backgroundColor`, which still works (`background` wins when both are given).
  `ChartRenderOptions.fonts` / `useSystemFonts` reach the rasteriser `ChartClassic.toPng`, `ChartEx.toPng` and
  `Chart.toPng` draw with, with `Draw.toRgba`'s meaning (`{ fonts, useSystemFonts: false }` gives the same PNG
  on every machine). In a browser `toPng` lets the platform draw the SVG with the page's fonts unless the caller
  chose the fonts — `fonts`, or `useSystemFonts: false` — and then uses the built-in rasteriser there too. For the
  characters no font covered, use `Draw.toRgba(ChartClassic.toDrawList(model), { fonts }).uncoveredCodePoints`.
- One route puts a chart on a PDF page, as for every other producer:
  `Draw.render(ChartClassic.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))` (likewise
  `ChartEx.toDrawList`, which draws every layout). The unreleased `ChartClassic.drawPdf`, `ChartEx.drawPdf`,
  `ChartEx.canDrawVectorPdf` and the types `ChartPdfRect`, `ChartDrawPdfOptions` and `ChartExDrawPdfOptions` are
  gone. The published `drawChartPdf`, `drawChartExPdf` and `canRenderChartExAsVectorPdf` stay, deprecated, and so
  do the types they use (`ChartPdfDrawingSurface`, `ChartPdfPathOp`, `PdfChartRenderOptions`, `ChartPdfColor`,
  `PdfColor`); the two drawing forwards draw the same display list (see the chart area, under "Output changes").
- `PdfLimitExceededError` gains `limit` (`PdfLimit`: `"maxDecodedBytes"`, `"depth"` or `"filterChain"`) and `allowed`
  (that limit's value), as `ArchiveLimitError` has, so which limit fired no longer has to be read from the message.
  It is built with `(limit, allowed, message, options?)`; the `(message, options?)` form still works, deprecated, and
  leaves both `undefined`.
- `ChartClassic.toDrawList(model, options?)` and `ChartEx.toDrawList(model, options?)` return the `DrawList` a chart
  draws as, like `Mermaid.toDrawList`. With `Draw.toRgba(list, { fonts, useSystemFonts: false })` it gives chart
  pixels that do not depend on the host's fonts.
- `EditorReadOptions` gains `fonts`, `preferSystemFonts`, `textLanguage`, `disableFontAutoDiscovery` and `onWarning`,
  with the builder's names, types, checks and behavior; `onWarning` reports the auto-embedded face and any character
  no face could draw. With `fonts` set, `editor.embedFont(…)` is refused, since the configuration already names the
  faces. `EditorCopyPagesOptions` declares its own `password` instead of extending `EditorReadOptions`.
- `DocumentMetadata.keywords` (so `Pdf.Builder#setMetadata` and `Pdf.Editor#setMetadata`, where `null` removes it)
  and `PdfExportOptions.keywords` write `/Keywords` and, in an XMP packet, `pdf:Keywords`.
- Widened: `Pdf.extract` and `Pdf.verifySignatures` also take a `Blob` (a browser `File`), whose bytes they read;
  the synchronous readers (`Pdf.Editor.read`, `copyPagesFrom`) take bytes only.
- Widened: `Pdf.extract`, `Pdf.Editor.read`, `editor.copyPagesFrom`, `Pdf.verifySignatures` and
  `Pdf.Signature.parseAsn1` / `verify` / `fillPlaceholder` take a `Uint8Array`, any other `ArrayBufferView` (a Node
  `Buffer` at an offset, a `DataView`, an `Int8Array`) or an `ArrayBuffer`, viewed rather than copied;
  `Pdf.Signature.*` took only a `Uint8Array`, and a typed array of wider elements such as `Uint16Array` is read as its
  bytes where it was refused. `ToSvgOptions.background` and `RasterizeOptions.background` take a CSS color string or
  an IR color `{ r, g, b, a }`. `PdfExportOptions.sheets`, `PdfWatermarkFilter.pages` and `.sheets` are readonly
  arrays.
- Widened: readonly arrays for annotation and form-field `rect`, `quadPoints`, dropdown `options`, radio `buttons`,
  `dashPattern`, `ReadPdfOptions.pages`, `PdfDrawPage.drawPath`, `BasicRasterCanvas#drawPolyline` / `fillPolygon` /
  `fillRings` and the `Pdf.Signature.verify` `byteRange`. Units are stated in TSDoc: `DocxToPdfOptions` page size,
  margins and `defaultFontSize` are points; `PdfExportOptions.margins` are points, and a side left out takes the
  sheet's own margin or 72 pt.

#### Archive and streams

**Output changes**

- ZIP DOS time is read as local time, as it was always written. Reading it as UTC shifted every `lastModified` by
  the reader's offset whenever the archive had no UT field — the default `"dos"` mode.
- A ZIP `modTime` the DOS fields cannot hold (1980–2107) but the UT field can (1970-01-01 to 2106-02-07) gets a UT
  field even in `"dos"` mode, so it round-trips; 1979 read back as 2107. More generally, `"dos"` mode adds a UT field
  whenever the DOS fields alone would not read back as `modTime` — the repeated hour of a DST change too — so every
  `modTime` from 1970 to 2106 reads back as written, to two seconds (such entries are 9 bytes larger). In UT modes a
  date only the DOS fields hold gets no UT field. A time taken from a file on disk is clamped to the DOS range, and
  Word's `Document.toBuffer` and the XLSB writer move the document's date to the nearest one a ZIP can store.
- `ZipEditor` keeps untouched entries readable by other tools: a copied ZipCrypto entry keeps its data-descriptor
  flag exactly as it had it (Info-ZIP, 7-Zip and Windows refused the password), and untouched entries keep their raw
  DOS date and time, which were recomputed in the editing machine's time zone.
- TAR headers: an mtime after 2242 or before 1970, and a uid/gid above 2097151, are written in GNU base-256 (2300 read
  back as 2011; the others were refused), and the reader decodes negative base-256 values; a link target over 99
  bytes goes in a GNU long-link record (it was cut); the ustar name/prefix split measures UTF-8 bytes (a 60-character
  CJK name was truncated).
- `ArchiveFile`:
  - Opened in TAR read mode, it keeps its entries on write (mode, uid/gid, names, mtime, type, link target), followed
    by the edits; an unchanged archive came out empty. A `.tgz` opened by path is written back gzip-compressed. PAX
    attributes beyond those fields are not carried over.
  - Every write builds the same archive and commits it: adding to a ZIP after `toBuffer()` dropped everything built
    before, and a second TAR `toBuffer()` threw.
  - ZIP `rename()` reaches the output. Streamed writes (`stream()`, `pipeTo()`, `streamToFile()`) write exactly what
    `toBuffer` writes: they wrote only the pending entries, and ignored `transform`, file stats and stamps.
  - Reads answer for the archive a write would produce now: the opened entries with pending adds, replacements,
    deletes, renames and directory/glob expansions applied, in `getEntryNames`, `entryCount`, `getEntries`,
    `getEntriesSync`, `getEntry`, `readEntry*`, `readAsText*`, `extractTo*`, `extractEntryTo*` and
    `hasEncryptedEntries`. They showed the archive as last opened or written; only `has()` saw pending edits. With
    pending edits, ZIP names are listed in the order the write stores them (by name); a new archive is readable before
    it is written (it threw "archive not loaded"); `entryCount` counts entries, not pending operations; a pending
    entry's size, CRC, dates and mode are what the write records (a pending ZIP entry is compressed once to describe
    it, cached until the next edit). A pending stream is buffered in memory the first time a read or write needs it
    and reused after; a synchronous read of a stream not yet read throws `ArchiveError`, and `toBufferSync()` works
    once an async read or write has buffered it.
  - `entryCount`, `getEntryNames()`, `getEntriesSync()`, and TAR `readEntrySync`, `readAsTextSync` and
    `extractToSync`, work for a TAR (0, "not supported for TAR", or a throw).
  - An entry with no date of its own is stamped when it is added, not when it is written: `modTime`, or 1980-01-01
    with `reproducible`. `reproducible` and `modTime` were ignored for buffers, text and symlinks; TAR directory
    entries now carry their file-system mtime.
  - The last write of a name wins, in reads and in the output; writing a name twice produced two entries. A TAR
    entry that replaces an opened entry keeps its position. `delete()` and `rename()` reach names a pending
    `addDirectory` / `addGlob` would add.
  - An opened ZIP's comment is kept on write (it was dropped), and `setZipComment` on an opened archive counts as a
    pending change.
  - `writeToFile` / `writeFile` honor `overwrite` for TAR too. TAR silently overwrote; pass
    `{ overwrite: "overwrite" }` for that.
- `compress` / `compressSync` (and `Compression.deflate*`) at `level: 0` return stored DEFLATE blocks; they returned
  the input unchanged, which no inflater can read.
- `ZipReader.get()` / `bytes()` buffer a streamed source, as documented, instead of throwing. `unpackTar` and
  `openTar` accept any `ArchiveSource`, and `unpackTarSync` / `unpackTarGzSync` an `ArrayBuffer` (`unpackTarSync`
  returned an empty map).
- A one-shot source — a `ReadableStream`, Node stream or async iterable given to `Archive.openZip` / `openTar` — is
  read once: by one `entries()` pass, or by buffering it the first time `list`, `get` or `bytes` needs it (later calls,
  `entries()` included, use that copy). A second pass is an `ArchiveError` saying the stream was consumed. It answered
  `[]` for TAR (an archive with entries reported empty), rejected with an `AbortError` for TAR over a Node stream, and
  with `EocdNotFoundError` ("Invalid ZIP file") for ZIP. `openTarGz` is unchanged.
- Duplicate names: `TarReader#get` / `bytes` (and `openTarGz`'s) return the later of two members with one name, as
  `unpackTar` and `tar -x` do; they returned the first. `RemoteZipReader#listFiles` lists one path per file entry,
  directories excluded, so its length is `getFileCount()`; it listed directories and folded duplicates. An unedited
  `editZip` writes every entry of a duplicated name back (it dropped all but one); editing the name leaves one.
  `ZipEditor#getSourceEntries` lists every entry.
- `ZipEditor.apply(plan)` changes nothing when it throws; it applied every operation before the refused one.
- Readers hand out frozen data: `RemoteZipReader.getEntries()` / `getEntry()` / `filterEntries()` / `findEntries()`
  and `ZipEditor.getSourceEntries()` (mutating them corrupted the reader); `ZipEditPlan.getOperations()` returns a frozen
  copy, and `new ZipEditPlan(ops)` copies `ops`.
- `pipeline(…, callback)` returns the last stage and `finished(stream, [options], callback)` a cleanup function, as
  `node:stream`'s do _(owner decision)_, on both entries; they returned a promise. Without a callback both still
  return a promise.
- `isReadable` / `isWritable` on Node answer as `node:stream`'s do, Web streams included (a readable
  `ReadableStream` was `false`); a non-stream is still `false`.
- `consumers.buffer()` resolves to a `Buffer` on Node. `consumers.text`, `Streams.text` and `Streams.json` decode as
  `node:stream/consumers.text` does: a leading BOM is stripped, string chunks are taken as they are, and a string or
  array is accepted as the source (the deprecated flat `streamToString` and `text` keep the BOM). `consumers.bytes` /
  `consumers.json` are the same functions as `Streams.bytes` / `Streams.json`. `addAbortSignal` cancels a Web
  `ReadableStream` and aborts a `WritableStream`. `ChunkedBuilder.toString()` no longer moves `cursor`.

**Now refused** — in `documonster/archive`, `ArchiveArgumentError` (an `ArchiveError`, with `field` such as
`options.level`, `options.modTime`, `options.encoding`, `source`, `entries[0]`, `sink`; options were named without
`options.`) for a wrong argument or option, while damaged data and state errors stay `ArchiveError`; in
`documonster/stream`, `StreamTypeError` / `UnsupportedStreamTypeError` with a `field`. Messages start with the public
call: `Archive.packZip: …`, `Archive.ZipArchive#add: …`, `Archive.ArchiveFile#addText: …`,
`Streams.bytes: …`, `Streams.createTransform: expected a function, got null` (they
began `createZip: …`, `ZipArchive.add: …`, `toBufferSync() only supports …`). `UnsupportedStreamTypeError.operation`
is the public name (`Streams.discard`, was `drainStream`). These were raw `TypeError`s, wrong results or nothing:

- `ZipArchive#add` / `addDirectory` / `addSymlink`, `packZip` and `packZipSync` with a name already added (field
  `name`); they wrote two entries of one name.
- Every archive reader checks its limit options when called, as `openZip` / `openTar` / `openTarGz` do:
  `unpackZip`, `unpackZipSync`, `unpackTar`, `unpackTarSync`, `parseTar`, `parseTarStream`, `unpackTarGz`,
  `unpackTarGzSync`, `parseTarGz`, `parseTarGzStream` and `ArchiveFile.fromBuffer` / `fromFile` / `fromFileSync`. A
  `maxEntrySize`, `maxEntries`, `maxTotalUncompressedSize` or (TAR) `maxFileSize` that is not a whole number from 0
  or `Infinity` is refused (field `options.maxEntries`, …), and so are options that are not an object. A mistyped
  limit (`"x"`, `null`) read the archive with no limit at all; `-1` was reported as the limit being hit
  (`ArchiveLimitError`), and in `ArchiveFile` as a `RangeError`. `unpackZip`, `unpackTarGz`, `parseTarGz` and
  `parseTarGzStream` also check the source at the call. `parseTarStream` and `parseTarGzStream` check their source
  and options when called rather than on the first `next()`.
- `openZip`, `openTar` and `openTarGz` check the source and options when called; they failed on first read. A lookup
  name (`get`, `bytes`, `has`, `delete`, `getEntry`, `readEntry`, `extract`, …) must be a non-empty string on every
  reader; `ArchiveFile` threw a raw `TypeError`, and `ZipReader`, `TarReader` and `RemoteZipReader` answered "not
  found". The builders refuse option values that are not objects. Under a `path` mode, `add` refuses a name the mode
  refuses, not `toBuffer`.
- `ArchiveFile`: an unknown `format` (it wrote a TAR); `addBuffer`, `addText`, `set`, `addZipComment` and
  `appendStream` given the wrong kind of value (`addText(42, …)` wrote `"42"`); ZIP options checked like
  `ZipArchive.add`'s and `createZip`'s; an empty name to `delete()` / `rename()` (`ArchiveArgumentError`, field
  `name`); and an entry that would be encrypted in a streamed write (`ArchiveError`) — it was silently written
  unencrypted.
- A ZIP date neither field can hold, in `createZip`, `add`, `packZip*`, `editZip` / `ZipEditor.set` and
  `ArchiveFile`; it wrapped (2108 read back as 1980).
- TAR `uname` / `gname` over 31 bytes and a `type` outside the entry types; they were truncated (`"bogus"` → `"b"`).
- A truncated TAR, in every reader (`unpackTar`, `unpackTarSync`, `parseTar`, `parseTarGz`, the flat
  `parseTarGzStream`, `openTar(bytes)`), including a cut header; a cut entry was returned short, and an input under
  one block read as empty. Length 0 is still an empty archive, and one ending after a whole entry without end blocks
  still reads.
- A gzip-compressed TAR given to a plain TAR reader, with a message pointing to `unpackTarGz` / `parseTarGz` /
  `openTarGz`. `parseTar` refuses anything but bytes, `parseTarStream` anything that is not an archive source,
  `unpackZipSync` a
  `Blob` or stream (naming `unpackZip`), and `ZipReader` a source that is not an `ArchiveSource`.
- Entry sources are checked at the call, not at `toBuffer()`: `ZipArchive.add`, `TarArchive.add`,
  `TarGzArchive.add`, `ZipEditor.set`, `ZipEditPlan.set`, `packZip*` / `packTar*`. `TarArchive.add(name, undefined)`
  wrote an empty file; `null` is still a directory. `addSymlink` targets must be strings, and
  `ZipEditor.setComment` / `ZipEditPlan.setComment` take a string or `undefined`.
- `Compression.*`: a level outside 0–9 or `NaN`, non-bytes, and damaged input, reported as
  `Invalid <codec> data: …` with the codec's error as `cause`. The deprecated flat forwards (`compress`, `gzip`, …)
  stay unchecked.
- `encodePng` with bad dimensions, a wrong buffer length or `rgba` that is not 8-bit (it accepts a canvas
  `Uint8ClampedArray`), a `dpi` that is `NaN`, 0, negative, `Infinity` or not a number (field `options.dpi`; it
  silently wrote no `pHYs` chunk — `dpi: undefined` still means none), and non-object `options`;
  `createBufferReader` with non-bytes. `createBufferReader().read(start, end)` and `HttpRangeReader.read(start, end)`
  follow one rule: an offset outside the source is a `RangeError` naming the parameter and the bounds
  (`end 11 is out of range: it must be > start (0) and <= 10`), so `instanceof RangeError` still catches every
  out-of-range case,
  while an offset that is not an integer (`NaN`, a fraction, `Infinity`, a non-number) is an `ArchiveArgumentError`
  with `field` `start` or `end` — it was a `RangeError`. Also a URL that is not a string or `URL` to
  `HttpRangeReader.open`, `RemoteZipReader.open` and
  `editZipUrl`; an AES key strength other than 128/192/256 in `Archive.Crypto`, a version other than 1/2 in
  `createAesExtraField`, and a negative or fractional length to the deprecated `Archive.randomBytes`.
- The byte utilities `crc32` / `crc32Update` / `crc32Finalize`, `base64ToUint8Array` (which now refuses anything
  outside the standard alphabet; Node decoded garbage to partial bytes and accepted base64url), `uint8ArrayToBase64`,
  `concatUint8Arrays` (a `totalLength` that is not the combined length was zero-padded) and the string conversions;
  the deprecated flat forms are unchanged. `Compression.create*Stream` with options that are not an object.
- Handles, where they are given: `packZip` / `packTar` / `packTarGz` (and `Sync`) with entries that are not a `Map`
  or an iterable of pairs or `{ name, source }`; `toReadableStream` / `toNodeReadable` with a non-source; `pipeTo`
  with a non-sink, before the archive starts streaming, so the builder is not sealed; `ZipEditor.apply(nonPlan)`;
  `ZipEditPlan.concat` / `applyTo` / `deserialize` / `fromJson` (invalid JSON has the `SyntaxError` as `cause`);
  `RemoteZipReader.fromReader(nonReader)`, `createTarReaderEntry` and the `isTar*` predicates.
- `addAbortSignal` with a non-stream or non-signal; a `null` entry in `Streams.createReadableFromArray` (it ended the
  stream there; `undefined` is still skipped); the wrong kind of argument to `Streams.collect`, `waitForFinished`,
  `toWritable` (it returned `null`), `LowLevel.SinkAdapter`, `fromString` (it stringified a number), `fromBytes` (it
  took a string) and `ChunkedBuilder.push` / `pushAll`; `BufferedStream.write` with a chunk that is neither bytes nor
  a string (field `chunk`; a bad one lost later data),
  `createStringChunk`, `createByteChunk` and `batchSize` (field `options.batchSize`); `PullStream.write`, its pattern
  and `highWaterMark` (field `options.highWaterMark`);
  `createTransform` (field `transformFn`, was `transform`), `filter`, `createReadableFromGenerator`,
  `createReadableFromAsyncIterable`,
  `createReadableFromPromise`; `fromWebReadable`, `toWebReadable`, `fromWebWritable`, `toWebWritable`,
  `readableStreamToAsyncIterable`; `Streams.waitForFinished` entries, `LowLevel.onceEvent`,
  `LowLevel.eventsToAsyncIterableNoDestroy` and `LowLevel.fromCallback`. `fromJson` names what it could not serialize
  (a function, a symbol, `undefined`, a BigInt; the flat `fromJSON(undefined)` gave an empty stream). The browser
  `finished(nonStream)` is an `UnsupportedStreamTypeError`. The `node:stream`-named functions (`pipeline`,
  `finished`, `compose`, `Readable.from`) keep Node's own `ERR_INVALID_ARG_TYPE` on Node.

**New names and deprecations**

- `Archive.parseTarGzStream` now parses incrementally, matching `parseTarStream`; the deprecated flat
  `parseTarGzStream` still decompresses everything first. The unreleased `parseTarGzIncremental` and
  `Streams.transform` (a duplicate of `Streams.createTransform`) are gone. `Archive.packZip` / `packZipSync` options
  are `Omit<ZipOptions, "format">`: they always write ZIP, and `format: "tar"` was accepted and ignored.
- `ArchiveBytes` (`ArrayBufferView | ArrayBuffer`): every archive reader takes any `ArrayBufferView` or `ArrayBuffer`,
  another realm's too — `openZip`, `unpackZip(Sync)`, `openTar`, `openTarGz`, `unpackTar(Sync)`, `unpackTarGz(Sync)`,
  `parseTar`, `parseTarStream`, `parseTarGz*`, `editZip`, `createBufferReader`, `ArchiveFile.fromBuffer`,
  `toReadableStream`, `toAsyncIterable`. `parseTarStream` takes any archive source, as `parseTarGzStream` does.
- Limits: `maxEntrySize` and `maxEntries` on `openZip` (`UnzipOptions`), `RemoteZipReaderOptions` and
  `ArchiveFile`'s ZIP open options; `maxEntrySize` on the TAR options (`TarParseOptions`, `TarReaderOptions`).
- `FileTooLargeError` (the `maxEntrySize` limit) carries the same `limit` / `allowed` fields as `ArchiveLimitError`,
  read the same way: `limit: "maxEntrySize"` and `allowed` (the configured `maxEntrySize`; `Number.MAX_SAFE_INTEGER`
  for a ZIP64 size or offset JavaScript cannot hold; `NaN` when constructed without it), plus `path`, the entry. It
  said both only in its message. It is still not an `ArchiveLimitError`, and `ArchiveLimit` is unchanged. The
  constructor is unchanged, with `allowed` in its options (`FileTooLargeErrorOptions`); messages are unchanged except
  that TAR's names the sizes.
- `packZip`, `packZipSync`, `packTar`, `packTarSync`, `packTarGz` and `packTarGzSync` declare their entries as a
  `ReadonlyMap` (any read-only map, which they read by iteration) or an iterable of pairs or `{ name, source }`
  objects. A `ReadonlyMap` already type-checked through the iterable arm; the declaration now says so.
- A string `source` given to an archive reader (`openZip`, `unpackZip`, `openTar`, `unpackTar`, `parseTarStream`,
  `openTarGz`, `unpackTarGz`, `parseTarGz`, `parseTarGzStream`, `editZip`, `toAsyncIterable`, `toReadableStream`,
  `toNodeReadable`) is the archive's content, encoded as UTF-8, never a file path or URL; this is unchanged and
  is now stated beside each parameter.
- `RemoteZipExtractOptions.onProgress`; `UnzipEntry.isDirectory`; `TarReaderEntry.readableStream()` and
  `pipeTo(WritableStream, options)`; `TarReader#entriesStream()`. `ArchiveFile` in TAR mode answers `has` and
  `getEntry`. `documonster/stream` exports `AbortError`, the same class as `documonster/archive`'s.
- Types: `ArchiveFile#setZipComment` returns `F extends "zip" ? this : never`, like `setPassword`; `ArchiveFile#has`
  returns `boolean` for both formats, and `getEntry` returns
  `(F extends "tar" ? ArchiveEntryInfo : ZipEntryInfo) | null`.
- `ArchiveArgumentError`. `Archive.ZipReader#list()`, `Archive.UnzipEntry#info`, `ArchiveFile#addSymlink(name, target, { mode })`,
  `#setZipComment`, `#writeFile`, `#writeFileSync`; `WorkerStreamStartError`, which `WorkerPool#openStream` rejects
  with and which could not be imported; types `RawEntryReadOptions`, `RawEntryStreamOptions`,
  `ArchiveOperationBase`, `UnzipOptionsZip`, `ZipOptionsZip`, `ArchivePipeToOptions`, `WorkerPoolStream`.
- `consumers.bytes`, `Streams.BufferedStream#readAll()`, `PullStream#pullUntil(pattern, { includePattern })`; types
  `StreamErrorOptions`, `PullUntilOptions`, `StreamByteChunk`, `StreamObjectModeOptions`, `StreamByteModeOptions`,
  `StreamDestroyHook`, `StreamConsumers`. `documonster/stream` does not export a `Stream` class; use
  `Streams.isStream` to ask whether a value is a stream.
- Deprecated, each keeping its behavior: `ArchiveFile#symlink` → `addSymlink`; `#addZipComment` → `setZipComment`
  (it sets, it does not add); `#writeToFile` / `writeToFileSync` → `writeFile` / `writeFileSync`;
  `#getEntryCountAsync()` → `entryCount`; `BufferedStream#toUint8Array()` → `readAll()` (it consumes the buffer
  despite its name); `PullStream#pull` and `pullUntil(pattern, boolean)` → `pullUntil(pattern, { includePattern })`;
  `RemoteZipReader#entriesGenerator(password)` → `entriesGenerator({ password })`; the TAR options' `maxFileSize` →
  `maxEntrySize` (still honored); `RemoteZipExtractOptions.onprogress` → `onProgress` (still called).

#### CSV, XML, Markdown and Formula

**Output changes**

- `Csv.format` / `Csv.createFormatterStream` _(owner decision)_: object rows and `[key, value]` rows get the first
  row's keys as a header row when `headers` is omitted, as the README and TSDoc always said, and every row is read by
  those keys. No header was written and each row kept its own key order, so `[{a:1,b:2},{b:3,a:4}]` gave `1,2\n3,4`.
  `headers: false` keeps the header-less output but still aligns columns. Array rows are unchanged.
- `castDate` reads a date-only value (`2020-01-02`) as UTC midnight _(owner decision)_, as ECMAScript does. It was
  local midnight, so the result depended on the machine and formatting it back moved the day east of UTC. Date-times
  without an offset are still local.
- Text after a closing quote is a record error _(owner decision)_: `"a"b,c` read as `a,c`, dropping `b`.
  `relaxQuotes` still reads the quote as a literal. A record error is never dropped silently (lead decision): the
  result-object forms of `Csv.parse` / `parseAsync` (with `headers`, `info` or `validate`) leave the record out of
  `rows`, report it in `errors` (`code: "ParseError"`, with `raw`), list it in `invalidRows` and pass it to `onSkip`
  — an unterminated-quote record is still kept and listed in `errors` there. The plain-array forms (`string[][]`, no
  `errors`), `Csv.parseRows` and so `ExcelCsv.read*` throw `CsvError` for a record error — text after a closing quote
  or an unterminated quote — unless `onSkip` is given, which then receives the record. In v0.21.1 an
  unterminated-quote record was returned with its partial text, and `parseRows` dropped bad records silently.
- Every header keeps its column: `constructor`, `prototype`, `__proto__`, a header holding a line break or tab, or
  one over 512 characters was dropped from every record. The same applies to `objname` keys and to
  `Csv.getRowValuesByHeaders` with more than ten headers.
- `onSkip` is called for every record left out because of an error, with or without `skipRecordsWithError`;
  `Csv.parseRows` dropped a column-mismatch record without a report.
- `dynamicTyping` keeps as strings integers no double holds exactly (`"9007199254740993"`) and literals outside the
  double range (`"1e400"`); they became a neighboring number. `Csv.parseNumber("")` and a blank string are `NaN`,
  not `0`.
- A field holding a character of a custom `lineEnding` is quoted. `Csv.forceQuote(v)` / `Csv.preventQuote(v)` work as
  cell values, not only as `typeTransform` results; they were written as JSON. A missing `constructor` / `toString` /
  `__proto__` key is written empty, not as `function Object() { [native code] }`, and `dynamicTyping: { … }` and
  `quoteColumns` maps are read by their own entries only (a column named `toString` was converted by
  `Object.prototype.toString`). `meta.renamedHeaders` has no prototype.
- `Xml.toPlainObject` returns a null-prototype object, like its nested objects and `Xml.parseToObject`.
- `Xml.encode`, and every writer's text, writes a carriage return as `&#xD;`: a literal CR is read back as LF by
  every XML parser. The change is in the shared encoder, so XLSX and DOCX text holding `\r` keeps it.
- `Xml.parse` keeps every attribute, including `constructor`, `__proto__` and `prototype` (the map has no
  prototype). `XmlDocument.prologue` holds only what precedes the root; what follows is in the new `epilogue`.
- `Xml.query` / `queryAll`: `//b//c` returns each node once, in document order, and `[@a="v"]` is accepted. Indexes
  stay 0-based. `Xml.Writer#rollback` after a `closeNode()` restores the closed elements (the next `closeNode()`
  threw), and `Xml.Writer#cursor` no longer decreases when the buffer is consolidated.
- `Markdown.format` makes as many columns as the longest of `headers`, `columns` and the rows, with an empty header
  cell for each extra one; extra columns and their values were dropped, and a table with rows but no headers was not
  written at all (`""`). A literal `<br>` in a cell is written `\<br>`, and `Markdown.parse` reads `\<` as `<` and never takes
  `\<br>` for a line break, so a format → parse round trip with `convertBr: true` returns the cells as written. Tables
  inside fenced or indented code blocks are not parsed.
- `Formula.format` writes external-workbook references back: `[1]Sheet1!A1` printed as `#REF!`. The node is still a
  `#REF!` error, now carrying `externalReference`. It escapes `[`, `]`, `#` and `'` in structured-reference column
  names with `'`, as Excel does, and the tokenizer reads those escapes in every bracket form (`T[a']b]` failed).
  Sheet names are quoted by the rule `Address.quoteSheetName` uses (`A1!B2`, `R1C1!A1`, `TRUE!A1` were written
  unquoted, and `.` and `\` let through); `_xlfn.`, `_xlws.` and `_xludf.` are written in lower case, as Excel writes
  them (`_xlfn.XLOOKUP`); a name may start with a backslash (`=\Name`), which lost the `\`.
- `FormulaParseError.position` is always a character offset into the text, as documented; parser errors reported a
  token index. Given tokens, there is no text, so `position` is absent and the message says `at token N`. Messages
  name token kinds (`Expected CloseParen`), not enum numbers.
- `A1#` (Excel's spill-range operator) parses as `UnaryOp { op: "#" }` and prints back as `A1#`; it was a `#NAME?`
  token and a parse error. The calculation engine evaluates `A1#` and `_xlfn.ANCHORARRAY(A1)` as the range the anchor
  spills into (`#REF!` when it does not spill); `ANCHORARRAY` was `#NAME?`. How a workbook stores `A1#` (as
  `_xlfn.ANCHORARRAY(A1)`, read back as `A1#`) is under Excel above.
- The engine decides what spills from the function table's return classes, replacing a hand-kept list of names: a
  formula that produces an array spills — newly `={1,2,3}`, `=SEQUENCE(3)*2`, `=TRANSPOSE(…)`, `=MMULT(…)`,
  `=MAP(…)`, `=TEXTSPLIT(…)`, `=LET(…)` or an invoked `LAMBDA` returning an array, and a formula with a spill
  reference on the path from its root (`=A1#*2`). Under a parameter that takes a range whole, a spill reference is a
  range: `SUM(A1#*2)` is 12 over `SEQUENCE(3)` (it was 2). A cell snapshot's `isDynamicArray: false` now means "an
  ordinary formula, as the host decided": its result is intersected with its cell even where the formula produces an
  array, which is how a formula read from a file as ordinary keeps Excel's result. Omit it to have the engine infer
  from the formula.
- A formula whose result is an empty cell holds 0, as Excel stores and shows it: `=A9` over an empty A9,
  `=IF(TRUE,A9)`, `=INDEX(A1:A3,2)`, a `VLOOKUP` or `LOOKUP` landing on an empty cell. Each empty element of a spill is
  0 too (`=A1:A3` with A2 empty spills `1, 0, 3`). It is the cell's value, not only its display: `ISBLANK(C1)` and
  `C1=""` are FALSE when C1 holds `=A2`, and over a spill `COUNTBLANK` is 0. The result was left empty
  (`Cell.getResult` was `undefined`). Inside an expression nothing changes (`A9=""`, `A9=0` and `ISBLANK(A9)` are
  TRUE, `A9&"x"` is `"x"`), and a formula returning `""` keeps `""`.
- A function argument Excel reads as an array — a range, an array constant, an array result — takes a single value
  (a one-cell reference or a literal) as a 1×1 array; it was `#VALUE!` or `#N/A`. A single error value is that error
  (`SORT(NA())` is `#N/A`), and a `LAMBDA` is `#VALUE!`. This covers `SORT`, `SORTBY`, `UNIQUE`, `FILTER`
  (`FILTER(A1,TRUE)` is A1, `FILTER(A1,FALSE)` is `#CALC!`), `TAKE`, `DROP`, `CHOOSEROWS`, `CHOOSECOLS`, `WRAPROWS`,
  `WRAPCOLS`, `EXPAND`; `INDEX` (`INDEX(A1,2)` is `#REF!`; it returned A1 whatever the index), `MATCH`, `XMATCH`,
  `XLOOKUP`, `VLOOKUP`, `HLOOKUP`, `LOOKUP`; every range of `COUNTIF`, `SUMIF`, `AVERAGEIF` and the `-IFS` functions;
  `MMULT`, `MINVERSE`, `MDETERM`, `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `FREQUENCY`, `CHISQ.TEST` / `CHITEST`, `IRR`,
  `MIRR`, `XNPV` and `XIRR`, where one cash flow reaches each function's own rule (`IRR` / `XIRR` `#NUM!`, `MIRR`
  `#DIV/0!`, `XNPV` the value). The database functions (`DSUM`, …) still require an array.
- Corrections to Microsoft's documented behavior: `INDEX` of a one-row array with one index counts along the row
  (`INDEX({1,2},2)`, `INDEX(A1:C1,2)` is B1; they were `#REF!`). `WRAPROWS` / `WRAPCOLS` give `#NUM!` for a
  `wrap_count` below 1 (was `#VALUE!`) and `#VALUE!` for a two-dimensional `vector` (it was flattened). Each `SORTBY`
  `by_array` must be one column as tall as `array` (sorting rows) or one row as wide as it (sorting columns, which is
  new), all keys the same way, else `#VALUE!`; a key of the wrong length read missing cells as blank, so
  `SORTBY(A1:A3, A1)` is now `#VALUE!`. `MIRR` without both a positive and a negative flow is `#DIV/0!` (a single
  value was `#NUM!`, an all-negative series -1).
- `OFFSET` and `INDEX` of a reference return a reference, as Microsoft documents; they returned the cells' values.
  `ROW(OFFSET(A1,2,0))`, `ROW(INDEX(A1:C3,2,2))` (2), `CELL("address",INDEX(H1:J3,2,2))` (`$I$2`), `AREAS(…)` and
  chained `OFFSET(OFFSET(A1,1,0),1,0)` / `OFFSET(INDEX(…),1,0)` work where they were `#VALUE!`, `ISREF(…)` is TRUE,
  and `A1:INDEX(A:A,5)` is a range (`SUM` over it is 15). `INDEX(ref,0,0)` returns the reference itself. With an
  array as its first argument `INDEX` still returns values (`ISREF(INDEX({1,2},1))` is FALSE). Every other use reads
  the same values as before, except that a multi-cell `OFFSET` as the whole formula of a cell that is not a dynamic
  array is intersected with the formula's row or column, like `=A1:A3`; it showed the block's top-left value.
- `SUMIF` / `AVERAGEIF` stretch the sum (average) range as Excel does: the cells used start at its top-left cell and
  take the criteria range's size and shape, so `SUMIF(A1:A3,">0",B1)`, `…,B1:B2)`, `…,B1:D9)` and
  `…,INDEX(B:B,1))` all sum B1:B3, and `SUMIF(A:A,">0",B1)` is `SUMIF(A:A,">0",B:B)`. A smaller sum range read its
  missing cells as blank, and a larger one was clipped — a one-cell sum range summed the criteria range. A stretch
  past the sheet's last row or column is cut there, as in Excel (`SUMIF(W!A:A,"<>",W!B2)` pairs A1 with B2). A union
  in either range is `#VALUE!`, and a sum range given as values (an array constant or result) must already be the
  criteria range's size, otherwise `#VALUE!` (it was read with blanks filled in). The `-IFS` functions do not
  stretch: every range must have the first one's size and shape. Dependency analysis includes the stretched cells.
- `LOOKUP` stretches `result_vector` from its first cell for as many cells as `lookup_vector` has: down a column when
  it is one column of two or more cells, otherwise across a row — a single cell runs across even for a vertical
  lookup (`LOOKUP(3,A1:A5,B1)` is D1). It is cut at the sheet's edge, and a two-dimensional `result_vector` or an
  array constant too short for the position is `#N/A`. These gave blank.
- `XLOOKUP` checks the return array's size: a one-column `lookup_array` is a vertical lookup and `return_array` needs
  as many rows (any number of columns, returned as a row); a one-row one is horizontal and needs as many columns; a
  one-cell one is vertical when `return_array` is one row tall. Anything else, and a two-dimensional `lookup_array`,
  is `#VALUE!`; a mismatch returned blank or a value from the wrong row. `MATCH` with a two-dimensional
  `lookup_array` is `#N/A` (it was read row by row) and `XMATCH` `#VALUE!` (it searched the first column).
- A whole column or row (`A:A`, `1:1`) is a reference to all of it, so positions in it match Excel. It was the
  sheet's used rows only, which shifted every position on a sheet whose data starts below row 1: with A3 = "a" and
  A4 = "b", `MATCH("b",A:A,0)` and `XMATCH` are 4 (were 2), `INDEX(A:A,4)` is "b" (was `#REF!`), `ROWS(A:A)` is
  1048576 and `COLUMNS(1:1)` 16384, and `CELL("address",A:A)` is `$A$1`. Whole columns on different sheets now have
  one length, so `XLOOKUP("b",S!A:A,T!A:A)`, `SUMIFS(T!A:A,S!A:A,"a")` and `COUNTIFS` across sheets work (they were
  `#VALUE!`) and `SUMIF(S!A:A,"",T!A:A)` reaches T's rows below S's data. Reading stops at the last used row of the
  sheets involved, so cells still pair up row by row. Counts include the empty cells of the full column or row, as
  in Excel: `COUNTIF(A:A,"")`, `COUNTIF(A:A,"<>x")`, `COUNTBLANK(A:A)` (1048573 with 3 values), `COUNTIFS(A:A,"",B:B,"")`.
  Sums, averages, `MINIFS` and `MAXIFS` are unchanged. A whole column beside a shorter range in an `-IFS` function is
  `#VALUE!`.

**Now refused** — argument and option checks throw `CsvArgumentError`, `XmlArgumentError`, `MarkdownArgumentError`
and `FormulaArgumentError` (each a subclass of the module's base error, with `field` such as `options.delimiter`,
`data[3]`, `rows[0][1]`, `input`, `element`, `node`). CSV messages start with the function called (`Csv.parse: …`,
`Csv.parseAsync: …`, `Csv.parseRows: …`, `Csv.createParserStream: …`, `Csv.format: …`,
`Csv.createFormatterStream: …`, `Csv.generate: …`, `Csv.generateRows: …`, …); they began `CSV parse:`,
`CSV format:` or `CSV generate:`. Problems in the input data keep their classes: CSV record
errors and a row over `maxRowBytes` (`CsvError`), malformed XML (`XmlParseError`), no Markdown table
(`MarkdownParseError`), formula syntax (`FormulaParseError`), and writer state (`XmlWriteError`).

- CSV: `chunk` / `chunkSize` in `Csv.parse` / `parseAsync`, which never called them (`Csv.parseRows` and the parser
  stream support them); `objname` without `headers` or naming no header — without headers it was ignored, and a
  missing column filed every record under `""` — and `objname` in `Csv.parseRows` / `createParserStream` (the stream
  ignored it, `parseRows` lost replaced records; use `Csv.parse` / `parseAsync` for the keyed map). Options of the
  wrong shape: a `delimiter` holding a line break, a `lineEnding` overlapping the delimiter or holding the quote,
  malformed `headers` / `columns` / `quoteColumns` / `quoteHeaders` / `typeTransform`, non-boolean flags, a
  `decimalSeparator` inside a multi-character delimiter. Data that is not an array, a row that is neither an array
  nor an object (a string row was split into one field per character), an `Invalid Date`, a `typeTransform` result
  that is not text, and a cell with no CSV text — a function (written as its source), a symbol, an object
  `JSON.stringify` cannot write (a circular one was `[object Object]`); plain objects and arrays are still JSON, and
  a `bigint` its digits. The inputs of `Csv.parseAsync`, `parseRows`, `detectDelimiter`, `detectLinebreak`,
  `stripBom`, `formatNumber`, `parseNumber`, `forceQuote`, `preventQuote`, `applyDynamicTypingToRow`, and a row,
  headers or columns that are not an array (or object) for `getRowValues` / `getRowHeaders` /
  `getRowValuesByHeaders` / `deduplicateHeaders` / `deduplicateHeadersWithRenames` / `processColumns` (the deprecated
  `rowHashArrayTo*` forwards stay unchecked). `Csv.generate*` and `createGenerator` options that are not an object,
  checked before they are read (`null` was a raw `TypeError`), a column type that is
  not built in (it generated strings), and `Csv.generateData({ rows: -1 })` (a plain `Error`).
- XML: names that are not QNames, in `Xml.validateName` and every writer method that takes one (`!--x`, `?pi`,
  `![CDATA[`, control characters, `:a`, `a:b:c` produced documents no parser accepts). `Xml.query` / `queryAll`
  refuse a leading single `/` (it searched every depth) and a path they do not support (empty step, `.`, `..`,
  `[@id]`, `[last()]`, other predicates, an empty path) instead of matching nothing. `Xml.Writer` /
  `Xml.StreamWriter` refuse `openXml()` once anything has been written (`XmlWriteError`); it wrote a declaration after
  content, such as `<a<?xml …?>`, which no parser accepts. `Xml.parse` stays lenient about a misplaced declaration.
  `writeRaw` refuses a non-string (field `xml`); `undefined` was written as nothing and `42` as `42`. Attribute errors
  name `attributes.<key>` (a map) or `value` (`addAttribute`). `maxDepth` / `maxEntityExpansions` that are not an integer ≥ 0
  (`NaN` and `-1` switched the limit off), an unknown `invalidCharHandling`, non-boolean flags, non-string input to
  `Xml.parse` / `parseToObject` / `encode` / `decode` / `encodeAttr` / `encodeCData` / `validateName` /
  `validateCommentText`, a non-element to the DOM helpers and `query*`, a `StreamWriter` target without `write`, an
  unknown event or non-function handler for `SaxParser#on`, a non-string chunk to `SaxParser#write`, a non-iterable
  source to `parseSax` / `feedSax`. A refused comment says which rule it broke. Invalid UTF-8 is an
  `XmlParseError`, where it was a `TypeError`.
- Markdown: non-boolean `padding`, `trailingNewline`, `escapeContent`, `trim`, `unescape`, `skipEmptyRows`,
  `convertBr`; a cell with no text (a function, a symbol, a circular object — plain objects are still JSON).
- Formula: input to `Formula.tokenize` / `Formula.parse` that is not text or tokens, and trees `Formula.format`
  cannot print (an unknown node type or operator, a non-finite number, a function name with a space).
  `Formula.tokenize` refuses a character it does not know with a `FormulaParseError` at its offset; it skipped it, so
  `=1~2` read as two numbers. A defined name's leading `!` before a reference (`!$A$1`) still reads as before. An
  error value Excel does not have (`#FOOBAR!`) is a `FormulaParseError` with its position; it was read as `#NAME?`,
  so the formula printed back was not the one written.

**New names and deprecations**

- `CsvArgumentError`, `MarkdownArgumentError`, `XmlArgumentError` and `FormulaArgumentError`, exported flat from
  their entries.
- CSV: `Csv.parseAsync` and `Csv.parseRows` take the whole input in bytes — any `ArrayBufferView`, an `ArrayBuffer`
  or a `Blob` / `File` — decoded like the same bytes streamed (`encoding`, BOM, `onProgress`), as `ExcelCsv.read`
  does; they took a byte stream but refused the same bytes whole. `CsvAsyncInput` includes them.
- CSV: `CsvParseOptions.decimalSeparator` (`"." | ","`), the parse counterpart used by `dynamicTyping`;
  `CsvFormatData`, the type `Csv.format` takes; `Csv.parse` / `parseAsync` overloads typing the `objname` result
  (`CsvParseResultWithObjname`) and the array-mode result with `validate`. Deprecated: `skipRecordsWithError`, which
  has no effect, and `CsvRecordInfo.invalid_field_length`, which is never set. Readonly arrays are accepted for
  `delimitersToGuess`, `headers`, `columns`, `Csv.getRowValuesByHeaders` and `Csv.processColumns`.
  `Csv.applyDynamicTyping(value)` types one value, as `dynamicTyping` does (`columnConfig` defaults to `true`).
- XML: `Xml.encodeCData` stays the name (there is no `Xml.encodeCdata`). `XmlDocument.epilogue`; `Xml.parseSax` is
  typed `AsyncGenerator<SaxEvent[], void, undefined>`. Deprecated: the `Xml.Writer#xml` getter — use `toString()`.
- Markdown: `Markdown.format` takes readonly headers, rows and `columns`.
- Formula: every member of `Token` and `AstNode` is exported under a `Formula` prefix (`FormulaNumberNode`,
  `FormulaCellRefNode`, `FormulaFunctionCallNode`, `FormulaErrorNode`, `FormulaCellRefToken`, …);
  `FormulaErrorNode.externalReference`, `FormulaErrorToken.externalReference`; `Formula.parse` takes `readonly Token[]`.

### Public API: one vocabulary across modules (deprecations, nothing removed)

Every public name now follows one set of conventions: namespaces of plain functions, `get` / `set` / `add` / `list` / `remove` verbs, `undefined` for absence, and module-prefixed error classes. **Nothing is
removed.** Every old name still works as a `@deprecated` forward with its old behaviour, so an
editor strikes it through and points at its replacement. All deprecated names are removed together in
**1.0**, not one minor at a time, so you migrate once.

Changes of meaning rather than of name — read these even if you migrate nothing else:

- **Excel reads return a new workbook.** `Workbook.read(bytes)`, `Workbook.readFile(path)`,
  `Workbook.readStream(stream)` and `Workbook.readWithDiagnostics(bytes)` take the input first and return
  the workbook (`report.workbook` for the diagnostics form). The forms that take a workbook first still
  work, deprecated; they replace everything that workbook held. `formulaSyntaxProbe`, which `Workbook.create`
  takes, is accepted in the read options. The same goes for CSV and Markdown: `ExcelCsv.read` / `readFile`
  and `ExcelMarkdown.read` / `readFile` (every table with `{ tables: "all" }`) take the input and return a new
  workbook. Their workbook-first forms still work, deprecated, and still _add_ worksheets to that workbook —
  which is now `ExcelCsv.addWorksheet` / `ExcelMarkdown.addWorksheet` / `ExcelMarkdown.addWorksheets`.
- **Word: `Document.*` edits in place, everything else does not.** The in-place edits that sat in `Query`
  moved to `Document` (`Document.replaceText`, `Document.mailMerge`, `Document.acceptAllRevisions`,
  `Document.rejectAllRevisions`, `Document.acceptRevision`, `Document.rejectRevision`,
  `Document.setCompatibilityMode`); `Query.fillFormFields` and `Query.resolveDataBindings`, which return a
  new document, moved to `Transform`. A function outside `Document` that returns a document leaves its input
  unchanged. The result may share unchanged nodes with it, but `Document.*` copies whatever the two share before
  editing it, so neither edit reaches the other and no `Document.clone` is needed (see
  [API shape](#api-shape-one-place-to-read-and-write-one-vocabulary-for-streaming-one-address-argument)). Only a node
  you change by hand — one taken through `Query.*` and mutated directly — is outside that guarantee.
- **Excel tables: the sheet's cells are what a table holds, and each edit applies as it is made.**
  - `Table.setRef`, `Table.addColumn`, `Table.setColumnName` and the other setters used to wait for
    `Table.commit`, so forgetting it wrote a table over cells that did not match it. They now apply
    immediately. `{ commit: false }` on `addRow` / `removeRows` no longer defers — a deferred row used to be
    left out of a file written before `commit` — and an appended row now costs only its own cells, so there
    is nothing to batch. `Table.commit` is deprecated and does nothing.
  - An edit validates before it changes anything: one that throws (`setRef` to an invalid address, an empty
    column name, a name the workbook already uses) leaves the table as it was. It used to leave the invalid
    value in the table.
  - **Bug fixes:** a table cell edited through `Cell.setValue` was reverted by the table's next edit, since
    the table wrote its own stale copy of the rows back over it; a pivot table or chart built from a table read
    that stale copy; inserting rows or columns, or moving the table, left each cell's own formatting at its old
    position; an inserted column took the formatting of the column it displaced.
  - `Table.getModel` returns a copy, its `rows` read from the sheet. `Table.setModel` now applies the model as
    `Table.add` would — the cells, the position and the name, checked against the workbook. It used to assign
    the model and nothing else, so the cells kept their old values and a second table could take the new name.
    The deprecated `Table.model` still returns the model the table uses.
  - `Table.getColumnStyle` returns a copy; `Table.setColumnStyle` applies a style to every cell of the column
    and takes back what the previous style gave them, keeping a cell's own formatting. The deprecated
    `Table.columnStyle` still returns the column's own object, applied on the table's next edit.
    `Table.getColumn` returns `undefined` for an index the table has no column at (and refuses one that is not an integer).
- **Excel handles' internal fields are deprecated.** The `_`-prefixed fields of `Workbook.Handle`,
  `Worksheet.Handle` and the streaming cell handle, plus `sourceFilePath`, `workbookContentType`,
  `userFunctions` and `pivotTables`, and a worksheet's `tables`, `pivotTables`, `conditionalFormattings`,
  `formControls` and `autoFilter`, are marked `@deprecated` and leave the published types in 1.0. They are
  bookkeeping or collections a namespace manages — `Table.list`, `Pivot.list`, `ConditionalFormatting.list`,
  `Form.list`, `Worksheet.getAutoFilter` read them. Document properties (`wb.title`, `wb.creator`, …),
  `properties`, `views`, `calcProperties`, and a worksheet's `pageSetup`, `headerFooter`, `properties` and
  `views` stay plain fields.
- **One way to add a chart.** `Chart.place(ws, range, options)` adds a chart of any family — the `type` chooses it,
  `groups` makes a combo, `pivotTable` binds it to a pivot table — and returns its handle. Its placement now comes
  before the options, as every other placement in the API does. The 31 `Chart.add*` functions are deprecated (they
  still return the part number); presets, tables and rows have option builders: `Chart.applyPreset`,
  `Chart.createOptionsFromTable`, `Chart.createOptionsFromRows`.
- New Excel names that pair with existing ones: `Workbook.listUnrecognizedValues(wb)` (beside
  `discardUnrecognizedValues`; no options — the list is filtered as any array is), `Workbook.verifyPassword(wb,
password)` (as `Worksheet.verifyPassword`), `Stream.applyCellStyle(cell, name)` (as `Cell.applyStyle`),
  `ExcelEncryptedError`, and the widenings `DefinedNames.add` / `addFormula` with
  `{ localSheetId }` returning the name as `get` reports it (they returned `void`) and `Worksheet.getValues(ws,
{ base })` — `base: 0` indexes as `Row.getValues` does; without options it stays 1-based.
- Duplicates deprecated in favour of the name they forward to: `Chart.setBuiltInStyle` (`setStyle`),
  `Chart.setTitleRichText` (`setTitle`), `Worksheet.destroy` (`Workbook.removeWorksheet`, which now also takes the
  sheet's handle) and `Workbook.renameChartsheet` (`Chartsheet.setName`).
- `DataValidation.add`, `find` and `remove` take the worksheet, as every other sheet-level namespace does; the rule
  collection (`ws.dataValidations`) is still accepted.
- Arguments refused at the call that used to pass, or fail later with an untyped `TypeError`: an invalid `Date`
  as a cell value or formula result (`ExcelInvalidValueTypeError`; the write failed on it), `Workbook.removeWorksheet`
  with no sheet (it removed the first one), `ExcelCsv.read` / `addWorksheet` with input that is not text, bytes, a URL,
  a Blob or a stream (it read `String(input)` — `42` became a one-cell sheet), `Csv.parse` with input that is not a
  string (`CsvError`), `Document.read` with input that is not bytes (`DocxArgumentError`; an `ArrayBuffer` is still
  read), and `DataValidation.*` with a target that is neither a worksheet nor a rule collection.
- `Chart.createOptionsFromRows` checks its start cell, that the rows fit on the sheet, and every value before it
  writes any cell, so a call that throws leaves the sheet as it was; it used to leave the header and the rows before
  the bad value behind.
- An edit to a worksheet that was removed from its workbook throws `ExcelArgumentError`. Only cell and row edits
  did; the rest — renaming the sheet, panes, filters, protection, column settings, charts, tables, pictures,
  checkboxes, validations, conditional formats, watermarks — were accepted and vanished, since a removed sheet is no
  longer written. The same holds for an edit through a handle on that sheet (a table, chart or checkbox), and for a
  chart handle after `Chart.remove`, whose edits were ignored or reported as the wrong kind of chart. Reads still
  answer. It also holds for every other sub-object a sheet hands out, which used to bypass the rule: the sheet's own
  data-validation collection (`DataValidation.add(ws.dataValidations, …)`), a placed image (`Image.setModel`,
  `Image.applyModel`), an image's anchor (`Anchor.setCol` / `setRow` / `setModel`), and a document-model row or cell
  handle given to the `Stream.*` setters — and to a sheet replaced by `Workbook.setModel`. A collection made on its
  own (`DataValidation.create`) belongs to no sheet and is edited freely.
- `Chart.createOptionsFromTable` refuses a table that is not on the sheet it is given. Its structured references
  named the table, but its plain ranges were built on the sheet passed, so the two forms charted different cells.
- `Chart.copyTo` names a missing target sheet instead of failing with a `TypeError`; its documentation said it
  defaulted to the chart's own sheet, which is what `Chart.clone` does.
- **A table name is never changed (behaviour change).** `Table.add` stored a name Excel would not accept under a
  corrected one — `"T1"` as `"_T1"`, `"My Table"` as `"My_Table"` — so a formula written with the name the caller chose
  named no table. It now throws a `ExcelTableError` naming a form Excel accepts; so do `Table.setName`,
  `Table.setDisplayName` and `Table.setModel`. A file whose table carries such a name is still read with it corrected.
  To migrate, pass a name Excel accepts (starts with a letter, `_` or `\`; letters, digits, `_` and `.` only; not a
  cell reference such as `T1`) — the error message suggests one — and use that same name in your formulas.
- `ExcelCsv` takes `ExcelCsvReadOptions` where it reads (`read`, `addWorksheet`, `readFile`, `createImportStream`) and
  `ExcelCsvWriteOptions` where it writes (`format`, `toBuffer`, `toBufferSync`, `toStream`, `writeStream`, `writeFile`), so an
  option the call ignores — a write's `dateFormat` on a read, a parser's `delimitersToGuess` on a write — is a type
  error in an object literal. `CsvOptions` is both halves and is accepted by every one of them. Its `info` and `raw`
  are deprecated: a sheet has nowhere to put them, and `info: true` made a read fail with a `TypeError`; a read now
  ignores both.
- `Pdf.Builder#toBuffer` refuses a document with both `setPdfACompliance` and `setEncryption`: PDF/A forbids
  encryption, and the file claimed a conformance it broke.
- Inputs widened, nothing narrowed: `Chart.createOptionsFromRows` takes rows of any object type (an `interface`
  needed an index signature) in a readonly array; the `Build.*` content builders, `Document.addHeading` /
  `addTable` / `addFootnote` / `addEndnote` / `addComment` / `addMath` / `addTextBox` / `addBulletList` /
  `addNumberedList`, `Template.bindChartData`, `Worksheet.addJson` / `addAoa`, `Table.addRow` / `addColumn`,
  `Csv.format` and the PDF `drawPath` / `splitPages` take readonly arrays, which they only read; `Document.read` declares
  the `ArrayBuffer` it has always read.
- `Table.create` and `Form.createCheckbox` are deprecated: each built an object for a sheet without adding it
  (a table that was never written; a checkbox that took a shape id another one could also take). Use
  `Table.add` and `Form.addCheckbox`.
- New error guards: `isArchiveError` (`documonster/archive`), `MermaidError` / `isMermaidError`
  (`documonster/mermaid`, with its syntax error — now `MermaidParseError` — extending `MermaidError`), and `documonster/chart`
  exports the `ExcelChartOptionsError` it throws.
- **Arguments are checked, and a few outputs change, in every module.** Every module has an `*ArgumentError` with a
  `field` naming the bad argument, and handles are checked at the boundary. Most of it refuses input that used to be
  accepted and then written as a broken file, but some calls now give different results for the same arguments —
  `Csv.format` writes a header row for object rows, `castDate` reads a date-only value as UTC, ZIP DOS times are read
  as local time, `pipeline` / `finished` with a callback return a stream and a cleanup function, a carriage return in
  XML text is written as `&#xD;`, `Pdf.Editor` keeps the file's `/Producer` and `/Creator`. See
  [Contracts tightened across every module](#contracts-tightened-across-every-module-behaviour-change).

**Positional forms deprecated in favour of an options object** (same name, the old overload is struck
through):

| Member                                                                                                                                      | Use                                                                                                          |
| ------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------ |
| `Workbook.read(wb, data, opts)`, `readFile(wb, path, opts)`, `readStream(wb, stream, opts)`, `readWithDiagnostics(wb, data, opts)`          | `read(data, opts)`, … — return a new workbook                                                                |
| `ExcelCsv.read(wb, input, opts)`, `ExcelCsv.readFile(wb, path, opts)`                                                                       | `ExcelCsv.addWorksheet(wb, input, opts)`; `read(input, opts)` / `readFile(path, opts)` return a new workbook |
| `ExcelMarkdown.read(wb, …)`, `readFile(wb, …)`                                                                                              | `ExcelMarkdown.addWorksheet(wb, …)` / `addWorksheets(wb, …)`; `read(input)`, … return a new workbook         |
| `Worksheet.fillFormula(ws, range, formula, results, shareType)`                                                                             | `fillFormula(ws, range, formula, { results, shareType })`                                                    |
| `Worksheet.freeze(ws, columns, rows)` / `Worksheet.split(ws, x, y)`                                                                         | `freeze(ws, { columns, rows })` / `split(ws, { x, y })`                                                      |
| `Worksheet.protect(ws, password, opts)` / `Workbook.protect(wb, password, opts)`                                                            | `protect(ws, { password, ...opts })`                                                                         |
| `Workbook.registerPerson(wb, name, userId, providerId)`                                                                                     | `registerPerson(wb, name, { userId, providerId })`                                                           |
| `Note.create(note, author)` / `Anchor.create(ws, address, offset)`                                                                          | `create(note, { author })` / `create(ws, address, { offset })`                                               |
| `Build.border(style, size, color)` / `Build.gridBorders(size, color)`                                                                       | `border({ style, size, color })` / `gridBorders({ size, color })`                                            |
| `Build.dateField(format, cachedValue)` / `Build.timeField(format, cachedValue)`                                                             | `dateField({ format, cachedValue })` / `timeField(...)`                                                      |
| `Build.field(instruction, cachedValue)`, `Build.pageNumberField(cachedValue)` and the ten other fields that took a positional `cachedValue` | `field(instruction, { cachedValue })`, `pageNumberField({ cachedValue })`, …                                 |
| `Build.table(rows, properties, columnWidths)` / `Build.underline(content, style, properties)`                                               | `table(rows, { ...properties, columnWidths })` / `underline(content, { ...properties, underline: style })`   |
| `Table.addRow(t, values, row, { commit })` / `Table.removeRows(t, index, count, { commit })`                                                | `addRow(t, values, row)` / `removeRows(t, index, count)` — `commit` never did anything                       |
| `Csv.detectDelimiter(input, quote, …)` / `Csv.detectLinebreak(input, quote)`                                                                | `detectDelimiter(input, { quote, delimitersToGuess, comment })` / `detectLinebreak(input, { quote })`        |
| `editor.copyPagesFrom(bytes, pages, opts)` (`Pdf.Editor`)                                                                                   | `copyPagesFrom(bytes, { pageIndices, password })`                                                            |
| `editor.splitPages(pageIndices)` (`Pdf.Editor`)                                                                                             | `splitPages({ pageIndices, encryption, invalidateSignatures })`                                              |
| `archiveFile.readEntrySync(path, password)` / `readAsTextSync(path, encoding)`                                                              | `readEntrySync(path, { password })` / `readAsTextSync(path, { encoding })`                                   |
| `archiveFile.symlink(name, target, mode)`                                                                                                   | `addSymlink(name, target, { mode })`                                                                         |
| `remoteZipReader.entriesGenerator(password)`                                                                                                | `entriesGenerator({ password })`                                                                             |
| `pullStream.pull(pattern, includePattern)` / `pullUntil(pattern, includePattern)`                                                           | `pullUntil(pattern, { includePattern })`                                                                     |

**Renamed or moved names**, by entry point:

<details>
<summary><code>documonster/excel</code> — 184 names</summary>

| Deprecated                              | Use instead                                                                                                                                                                                                                                                                                    |
| --------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `AddAOAOptions`                         | Use `AddAoaOptions`.                                                                                                                                                                                                                                                                           |
| `AddJSONOptions`                        | Use `AddJsonOptions`.                                                                                                                                                                                                                                                                          |
| `Anchor.col`                            | Use `Anchor.getCol`.                                                                                                                                                                                                                                                                           |
| `Anchor.colWidth`                       | Use `Anchor.getColWidth`.                                                                                                                                                                                                                                                                      |
| `Anchor.model`                          | Use `Anchor.getModel`.                                                                                                                                                                                                                                                                         |
| `Anchor.row`                            | Use `Anchor.getRow`.                                                                                                                                                                                                                                                                           |
| `Anchor.rowHeight`                      | Use `Anchor.getRowHeight`.                                                                                                                                                                                                                                                                     |
| `Cell.applyCellStyle`                   | Use `Cell.applyStyle`.                                                                                                                                                                                                                                                                         |
| `Cell.find`                             | Use `Cell.has` to test for a cell, `Cell.getView` to read one without creating it, and address it by position.                                                                                                                                                                                 |
| `Cell.getMergeMaster`                   | Use `Cell.getMergeMasterAddress`.                                                                                                                                                                                                                                                              |
| `Cell.Sheet`                            | Use `Worksheet.Handle`.                                                                                                                                                                                                                                                                        |
| `Cell.view`                             | The document model is addressed by position; use `Cell.getView` (`Cell.getView(ws, addr)`), or `Cell.getValue` and friends. For a streaming handle use `Stream.getCellView`.                                                                                                                   |
| `Chart.add`                             | Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                                           |
| `Chart.addArea`                         | Use `Chart.place(ws, range, { ...options, type: "area" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                      |
| `Chart.addBar`                          | Use `Chart.place(ws, range, { ...options, type: "bar", barDir: "bar" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                        |
| `Chart.addBoxWhisker`                   | Use `Chart.place(ws, range, { ...options, type: "boxWhisker" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                |
| `Chart.addBubble`                       | Use `Chart.place(ws, range, { ...options, type: "bubble" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                    |
| `Chart.addColumn`                       | Use `Chart.place(ws, range, { ...options, type: "bar", barDir: "col" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                        |
| `Chart.addColumnFromRows`               | Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, { ...options, type: "bar", barDir: "col" }))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). |
| `Chart.addCombo`                        | Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                                           |
| `Chart.addDoughnut`                     | Use `Chart.place(ws, range, { ...options, type: "doughnut" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                  |
| `Chart.addEx`                           | Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                                           |
| `Chart.addExFromRows`                   | Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                    |
| `Chart.addExFromTable`                  | Use `Chart.place(ws, range, Chart.createOptionsFromTable(ws, table, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                  |
| `Chart.addFromRows`                     | Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                    |
| `Chart.addFromTable`                    | Use `Chart.place(ws, range, Chart.createOptionsFromTable(ws, table, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                  |
| `Chart.addFunnel`                       | Use `Chart.place(ws, range, { ...options, type: "funnel" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                    |
| `Chart.addHistogram`                    | Use `Chart.place(ws, range, { ...options, type: "histogram" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                 |
| `Chart.addLine`                         | Use `Chart.place(ws, range, { ...options, type: "line" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                      |
| `Chart.addPareto`                       | Use `Chart.place(ws, range, { ...options, type: "pareto" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                    |
| `Chart.addPie`                          | Use `Chart.place(ws, range, { ...options, type: "pie" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                       |
| `Chart.addPivot`                        | Use `Chart.place(ws, range, { ...options, pivotTable })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                        |
| `Chart.addPivotCombo`                   | Use `Chart.place(ws, range, { ...options, pivotTable })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                        |
| `Chart.addPreset`                       | Use `Chart.place(ws, range, Chart.applyPreset(preset, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                |
| `Chart.addPresetEx`                     | Use `Chart.place(ws, range, Chart.applyPreset(preset, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                |
| `Chart.addRadar`                        | Use `Chart.place(ws, range, { ...options, type: "radar" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                     |
| `Chart.addRegionMap`                    | Use `Chart.place(ws, range, { ...options, type: "regionMap" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                 |
| `Chart.addScatter`                      | Use `Chart.place(ws, range, { ...options, type: "scatter" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                   |
| `Chart.addStock`                        | Use `Chart.place(ws, range, { ...options, type: "stock" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                     |
| `Chart.addSunburst`                     | Use `Chart.place(ws, range, { ...options, type: "sunburst" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                  |
| `Chart.addSurface`                      | Use `Chart.place(ws, range, { ...options, type: "surface" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                   |
| `Chart.addTreemap`                      | Use `Chart.place(ws, range, { ...options, type: "treemap" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                   |
| `Chart.addWaterfall`                    | Use `Chart.place(ws, range, { ...options, type: "waterfall" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types).                                                 |
| `Chart.anchorModel`                     | Use `Chart.getAnchorModel`.                                                                                                                                                                                                                                                                    |
| `Chart.axes`                            | Use `Chart.getAxes`.                                                                                                                                                                                                                                                                           |
| `Chart.categoryAxis`                    | Use `Chart.getCategoryAxis`.                                                                                                                                                                                                                                                                   |
| `Chart.chartExModel`                    | Use `Chart.getExModel`.                                                                                                                                                                                                                                                                        |
| `Chart.chartModel`                      | Use `Chart.getModel`.                                                                                                                                                                                                                                                                          |
| `Chart.chartTypes`                      | Use `Chart.getTypes`.                                                                                                                                                                                                                                                                          |
| `Chart.get`                             | Use `Chart.list`.                                                                                                                                                                                                                                                                              |
| `Chart.isChartEx`                       | Use `Chart.isEx`.                                                                                                                                                                                                                                                                              |
| `Chart.legend`                          | Use `Chart.getLegend`.                                                                                                                                                                                                                                                                         |
| `Chart.mutateChartEx`                   | Use `Chart.mutateEx`.                                                                                                                                                                                                                                                                          |
| `Chart.plotArea`                        | Use `Chart.getPlotArea`.                                                                                                                                                                                                                                                                       |
| `Chart.seriesFromColumns`               | Use `Chart.createSeriesFromColumns`.                                                                                                                                                                                                                                                           |
| `Chart.setBuiltInStyle`                 | Use `Chart.setStyle`, which this forwards to.                                                                                                                                                                                                                                                  |
| `Chart.setTitleRichText`                | Use `Chart.setTitle`, which takes rich text too.                                                                                                                                                                                                                                               |
| `Chart.spPr`                            | Use `Chart.getSpPr`.                                                                                                                                                                                                                                                                           |
| `Chart.title`                           | Use `Chart.getTitle`.                                                                                                                                                                                                                                                                          |
| `Chart.titleRichText`                   | Use `Chart.getTitleRichText`.                                                                                                                                                                                                                                                                  |
| `Chart.toPNG`                           | Use `Chart.toPng`.                                                                                                                                                                                                                                                                             |
| `Chart.toSVG`                           | Use `Chart.toSvg`.                                                                                                                                                                                                                                                                             |
| `Chart.totalSeriesCount`                | Use `Chart.getTotalSeriesCount`.                                                                                                                                                                                                                                                               |
| `Chart.unknownElements`                 | Use `Chart.getUnknownElements`.                                                                                                                                                                                                                                                                |
| `Chart.userShapesXml`                   | Use `Chart.getUserShapesXml`.                                                                                                                                                                                                                                                                  |
| `Chart.valueAxis`                       | Use `Chart.getValueAxis`.                                                                                                                                                                                                                                                                      |
| `Chartsheet.chart`                      | Use `Chartsheet.getChart`.                                                                                                                                                                                                                                                                     |
| `Chartsheet.chartExModel`               | Use `Chartsheet.getChartExModel`.                                                                                                                                                                                                                                                              |
| `Chartsheet.chartExNumber`              | Use `Chartsheet.getChartExNumber`.                                                                                                                                                                                                                                                             |
| `Chartsheet.chartModel`                 | Use `Chartsheet.getChartModel`.                                                                                                                                                                                                                                                                |
| `Chartsheet.chartNumber`                | Use `Chartsheet.getChartNumber`.                                                                                                                                                                                                                                                               |
| `Chartsheet.create`                     | Use `Workbook.addChartsheet`. This builds a chartsheet handle that is attached to nothing and never written.                                                                                                                                                                                   |
| `Chartsheet.headerFooter`               | Use `Chartsheet.getHeaderFooter`.                                                                                                                                                                                                                                                              |
| `Chartsheet.id`                         | Use `Chartsheet.getId`.                                                                                                                                                                                                                                                                        |
| `Chartsheet.model`                      | Use `Chartsheet.getModel`.                                                                                                                                                                                                                                                                     |
| `Chartsheet.name`                       | Use `Chartsheet.getName`.                                                                                                                                                                                                                                                                      |
| `Chartsheet.pageMargins`                | Use `Chartsheet.getPageMargins`.                                                                                                                                                                                                                                                               |
| `Chartsheet.pageSetup`                  | Use `Chartsheet.getPageSetup`.                                                                                                                                                                                                                                                                 |
| `Chartsheet.sheetNo`                    | Use `Chartsheet.getSheetNo`.                                                                                                                                                                                                                                                                   |
| `Chartsheet.state`                      | Use `Chartsheet.getState`.                                                                                                                                                                                                                                                                     |
| `Chartsheet.tabSelected`                | Use `Chartsheet.isTabSelected`.                                                                                                                                                                                                                                                                |
| `Chartsheet.workbookViewId`             | Use `Chartsheet.getWorkbookViewId`.                                                                                                                                                                                                                                                            |
| `Chartsheet.zoomScale`                  | Use `Chartsheet.getZoomScale`.                                                                                                                                                                                                                                                                 |
| `Chartsheet.zoomToFit`                  | Use `Chartsheet.isZoomToFit`.                                                                                                                                                                                                                                                                  |
| `Column.getHidden`                      | Use `Column.isHidden`.                                                                                                                                                                                                                                                                         |
| `Column.Sheet`                          | Use `Worksheet.Handle`.                                                                                                                                                                                                                                                                        |
| `Column.values`                         | Use `Column.getValues` with `{ base: 1 }`.                                                                                                                                                                                                                                                     |
| `DataValidation.create`                 | Pass the worksheet (or `ws.dataValidations`) to `DataValidation.add`. A collection made here belongs to no sheet, so its rules are never written.                                                                                                                                              |
| `DefinedNames.create`                   | Use `Workbook.getDefinedNames(wb)`. A registry this makes cannot be attached to a workbook.                                                                                                                                                                                                    |
| `DefinedNames.getAllEntries`            | Use `DefinedNames.list`.                                                                                                                                                                                                                                                                       |
| `DefinedNames.model`                    | Use `DefinedNames.getModel`.                                                                                                                                                                                                                                                                   |
| `Form.checkboxFromModel`                | No replacement: the handle this builds is never attached to a sheet or written. A checkbox's model is `handle.model`, and the VML getters take the model directly.                                                                                                                             |
| `Form.checked`                          | Use `Form.isChecked`.                                                                                                                                                                                                                                                                          |
| `Form.createCheckbox`                   | Use `Form.addCheckbox`. This builds a checkbox for `worksheet` without adding it, so two calls take the same shape id and neither is written.                                                                                                                                                  |
| `Form.link`                             | Use `Form.getLink`.                                                                                                                                                                                                                                                                            |
| `Form.listCheckboxes`                   | Use `Form.list`, which returns a copy. This form returns the sheet's live array, as it always did: before `removeCheckbox` existed, splicing it was how a checkbox was removed.                                                                                                                |
| `Form.text`                             | Use `Form.getText`.                                                                                                                                                                                                                                                                            |
| `Form.vmlAnchor`                        | Use `Form.Vml.getAnchor`.                                                                                                                                                                                                                                                                      |
| `Form.vmlCheckedValue`                  | Use `Form.Vml.getCheckedValue`.                                                                                                                                                                                                                                                                |
| `Form.vmlStyle`                         | Use `Form.Vml.getStyle`.                                                                                                                                                                                                                                                                       |
| `Format.parseValueByFormat`             | Use `Format.parseValue`.                                                                                                                                                                                                                                                                       |
| `Image.applyModel`                      | Use `Image.setModel`. Applies the model only, as it always did; `setModel` also records that a picture the image stops showing is no longer used.                                                                                                                                              |
| `Image.clone`                           | Use `Image.place`. This builds an image handle that is attached to nothing and never written.                                                                                                                                                                                                  |
| `Image.create`                          | Use `Image.place`. This builds an image handle that is attached to nothing and never written.                                                                                                                                                                                                  |
| `Image.getShapes`                       | Use `Image.listShapes`.                                                                                                                                                                                                                                                                        |
| `Image.model`                           | Use `Image.getModel`.                                                                                                                                                                                                                                                                          |
| `Note.model`                            | Use `Note.getModel`.                                                                                                                                                                                                                                                                           |
| `Range.count`                           | Use `Range.getCount`.                                                                                                                                                                                                                                                                          |
| `Range.forEachAddress`                  | Use `Range.eachAddress`.                                                                                                                                                                                                                                                                       |
| `Row.eachCell`                          | Use `Row.eachAddress`, which visits by position and takes its options last.                                                                                                                                                                                                                    |
| `Row.getCell`                           | Address the cell by position through the `Cell` namespace (`Cell.getValue(ws, row, col)`).                                                                                                                                                                                                     |
| `Row.getHidden`                         | Use `Row.isHidden`.                                                                                                                                                                                                                                                                            |
| `Row.Sheet`                             | Use `Worksheet.Handle`.                                                                                                                                                                                                                                                                        |
| `Row.values`                            | Use `Row.getValues` with `{ base: 1 }`.                                                                                                                                                                                                                                                        |
| `SheetToJSONOptions`                    | Use `SheetToJsonOptions`.                                                                                                                                                                                                                                                                      |
| `Stream.rowCell`                        | Use `Stream.getRowCell`.                                                                                                                                                                                                                                                                       |
| `Stream.rowValues`                      | Use `Stream.getRowValues(r, { base: 1 })`; `getRowValues` is 0-based by default, as `Row.getValues` is.                                                                                                                                                                                        |
| `Table.column`                          | Use `Table.getColumn`, which returns `undefined` for an index the table has no column at.                                                                                                                                                                                                      |
| `Table.columnFilterButton`              | Use `Table.hasColumnFilterButton`.                                                                                                                                                                                                                                                             |
| `Table.columnName`                      | Use `Table.getColumnName`.                                                                                                                                                                                                                                                                     |
| `Table.columnSetFilterButton`           | Use `Table.setColumnFilterButton`.                                                                                                                                                                                                                                                             |
| `Table.columnSetName`                   | Use `Table.setColumnName`, which this now forwards to: a name another column has is refused rather than renumbered.                                                                                                                                                                            |
| `Table.columnSetStyle`                  | Use `Table.setColumnStyle`.                                                                                                                                                                                                                                                                    |
| `Table.columnSetTotalsRowFormula`       | Use `Table.setColumnTotalsRowFormula`.                                                                                                                                                                                                                                                         |
| `Table.columnSetTotalsRowFunction`      | Use `Table.setColumnTotalsRowFunction`.                                                                                                                                                                                                                                                        |
| `Table.columnSetTotalsRowLabel`         | Use `Table.setColumnTotalsRowLabel`.                                                                                                                                                                                                                                                           |
| `Table.columnSetTotalsRowResult`        | Use `Table.setColumnTotalsRowResult`.                                                                                                                                                                                                                                                          |
| `Table.columnStyle`                     | Use `Table.getColumnStyle`, which returns a copy, and `Table.setColumnStyle` to change it. This returns the column's own style object.                                                                                                                                                         |
| `Table.columnTotalsRowFormula`          | Use `Table.getColumnTotalsRowFormula`.                                                                                                                                                                                                                                                         |
| `Table.columnTotalsRowFunction`         | Use `Table.getColumnTotalsRowFunction`.                                                                                                                                                                                                                                                        |
| `Table.columnTotalsRowLabel`            | Use `Table.getColumnTotalsRowLabel`.                                                                                                                                                                                                                                                           |
| `Table.columnTotalsRowResult`           | Use `Table.getColumnTotalsRowResult`.                                                                                                                                                                                                                                                          |
| `Table.commit`                          | Every edit applies to the sheet as it is made; there is never anything to commit.                                                                                                                                                                                                              |
| `Table.create`                          | Use `Table.add`. This builds a table bound to `worksheet` and writes its cells, but never registers it on the sheet, so it is not written, not found by `get` or `list`, and its name is not reserved.                                                                                         |
| `Table.displayName`                     | Use `Table.getDisplayName`.                                                                                                                                                                                                                                                                    |
| `Table.headerRow`                       | Use `Table.hasHeaderRow`.                                                                                                                                                                                                                                                                      |
| `Table.model`                           | Use `Table.getModel`, which returns a copy; change the table with `Table.setModel` or the setters. This returns the model the table uses.                                                                                                                                                      |
| `Table.name`                            | Use `Table.getName`.                                                                                                                                                                                                                                                                           |
| `Table.ref`                             | Use `Table.getRef`, which returns the table's whole range (`"B2:C4"`); this returns the reference as given (the top-left cell for a table built in memory).                                                                                                                                    |
| `Table.showColumnStripes`               | Use `Table.hasColumnStripes`.                                                                                                                                                                                                                                                                  |
| `Table.showFirstColumn`                 | Use `Table.hasFirstColumnStyle`.                                                                                                                                                                                                                                                               |
| `Table.showLastColumn`                  | Use `Table.hasLastColumnStyle`.                                                                                                                                                                                                                                                                |
| `Table.showRowStripes`                  | Use `Table.hasRowStripes`.                                                                                                                                                                                                                                                                     |
| `Table.theme`                           | Use `Table.getTheme`.                                                                                                                                                                                                                                                                          |
| `Table.totalsRow`                       | Use `Table.hasTotalsRow`.                                                                                                                                                                                                                                                                      |
| `Workbook.defineCellStyle`              | Use `Workbook.addCellStyle`, which returns the stored entry.                                                                                                                                                                                                                                   |
| `Workbook.eachSheet`                    | Use `Workbook.eachWorksheet`.                                                                                                                                                                                                                                                                  |
| `Workbook.getChartsheets`               | Use `Workbook.listChartsheets`.                                                                                                                                                                                                                                                                |
| `Workbook.getWorksheets`                | Use `Workbook.listWorksheets`.                                                                                                                                                                                                                                                                 |
| `Workbook.importSheet`                  | Use `Workbook.importWorksheet`.                                                                                                                                                                                                                                                                |
| `Workbook.renameChartsheet`             | Look the chartsheet up, then rename it: `const cs = Workbook.getChartsheet(wb, nameOrIndex); if (cs) Chartsheet.setName(cs, name);` — this answered `false` when there is no such chartsheet.                                                                                                  |
| `Workbook.useBuiltinCellStyle`          | Use `Workbook.useBuiltInCellStyle`.                                                                                                                                                                                                                                                            |
| `Worksheet.actualColumnCount`           | Use `Column.getActualCount`.                                                                                                                                                                                                                                                                   |
| `Worksheet.actualRowCount`              | Use `Row.getActualCount`.                                                                                                                                                                                                                                                                      |
| `Worksheet.addConditionalFormatting`    | Use `ConditionalFormatting.add`.                                                                                                                                                                                                                                                               |
| `Worksheet.addRow`                      | Use `Row.add`.                                                                                                                                                                                                                                                                                 |
| `Worksheet.addRows`                     | Use `Row.addMany`.                                                                                                                                                                                                                                                                             |
| `Worksheet.autoFilter`                  | Use `Worksheet.getAutoFilter`, which answers `undefined` rather than `null`.                                                                                                                                                                                                                   |
| `Worksheet.autoFitColumn`               | Use `Column.autoFit`.                                                                                                                                                                                                                                                                          |
| `Worksheet.autoFitColumns`              | Use `Column.autoFit`.                                                                                                                                                                                                                                                                          |
| `Worksheet.autoFitRow`                  | Use `Row.autoFit`.                                                                                                                                                                                                                                                                             |
| `Worksheet.autoFitRows`                 | Use `Row.autoFit`.                                                                                                                                                                                                                                                                             |
| `Worksheet.columnCount`                 | Use `Column.getCount`.                                                                                                                                                                                                                                                                         |
| `Worksheet.columnDefinitions`           | Use `Column.getDefinitions`.                                                                                                                                                                                                                                                                   |
| `Worksheet.columns`                     | Use `Column.getDefinitions`.                                                                                                                                                                                                                                                                   |
| `Worksheet.destroy`                     | Use `Workbook.removeWorksheet(wb, ws)`, which this does.                                                                                                                                                                                                                                       |
| `Worksheet.dimensions`                  | Use `Worksheet.getDimensions`, which returns `undefined` for a sheet with no cells; this returns a zero range that `Range.*` read as A1.                                                                                                                                                       |
| `Worksheet.duplicateRow`                | Use `Row.duplicate`.                                                                                                                                                                                                                                                                           |
| `Worksheet.eachRow`                     | Use `Row.each`.                                                                                                                                                                                                                                                                                |
| `Worksheet.findRow`                     | Use `Row.has`.                                                                                                                                                                                                                                                                                 |
| `Worksheet.findRows`                    | Use `Row.has`.                                                                                                                                                                                                                                                                                 |
| `Worksheet.getRow`                      | Use the `Row` and `Cell` namespaces.                                                                                                                                                                                                                                                           |
| `Worksheet.getRows`                     | Use the `Row` and `Cell` namespaces.                                                                                                                                                                                                                                                           |
| `Worksheet.insertRow`                   | Use `Row.insert`.                                                                                                                                                                                                                                                                              |
| `Worksheet.insertRows`                  | Use `Row.insertMany`.                                                                                                                                                                                                                                                                          |
| `Worksheet.lastColumn`                  | Use `Column.getLast`, which returns the column **number** (or `undefined` for an empty sheet) rather than a column view, and creates no column record.                                                                                                                                         |
| `Worksheet.lastRow`                     | Use `Row.getLast`, which returns the row **number** rather than the row record.                                                                                                                                                                                                                |
| `Worksheet.mergedRegions`               | Use `Worksheet.getMergedRegions`.                                                                                                                                                                                                                                                              |
| `Worksheet.panes`                       | Use `Worksheet.getPanes`, which returns `{ state: "frozen", columns, rows }` or `{ state: "split", x, y }` rather than the raw view (still in `ws.views`).                                                                                                                                     |
| `Worksheet.removeConditionalFormatting` | Use `ConditionalFormatting.remove` for an entry, a position or a predicate, and `ConditionalFormatting.clear` for every entry (what calling this without a filter does).                                                                                                                       |
| `Worksheet.rowCount`                    | Use `Row.getCount`.                                                                                                                                                                                                                                                                            |
| `Worksheet.setColumns`                  | Use `Column.setDefinitions`.                                                                                                                                                                                                                                                                   |
| `Worksheet.spliceColumns`               | Use `Column.splice`.                                                                                                                                                                                                                                                                           |
| `Worksheet.spliceRows`                  | Use `Row.splice`.                                                                                                                                                                                                                                                                              |

</details>

<details>
<summary><code>documonster/archive</code> — 113 names</summary>

| Deprecated                                   | Use instead                                                                                                                                                                     |
| -------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Archive.AES_AUTH_CODE_LENGTH`               | Use `Crypto.AES_AUTH_CODE_LENGTH`.                                                                                                                                              |
| `Archive.AES_EXTRA_FIELD_ID`                 | Use `Crypto.AES_EXTRA_FIELD_ID`.                                                                                                                                                |
| `Archive.AES_KEY_LENGTH`                     | Use `Crypto.AES_KEY_LENGTH`.                                                                                                                                                    |
| `Archive.AES_PASSWORD_VERIFY_LENGTH`         | Use `Crypto.AES_PASSWORD_VERIFY_LENGTH`.                                                                                                                                        |
| `Archive.AES_SALT_LENGTH`                    | Use `Crypto.AES_SALT_LENGTH`.                                                                                                                                                   |
| `Archive.AES_VENDOR_ID`                      | Use `Crypto.AES_VENDOR_ID`.                                                                                                                                                     |
| `Archive.AES_VERSION_AE1`                    | Use `Crypto.AES_VERSION_AE1`.                                                                                                                                                   |
| `Archive.AES_VERSION_AE2`                    | Use `Crypto.AES_VERSION_AE2`.                                                                                                                                                   |
| `Archive.aesDecrypt`                         | Use `Crypto.decryptAes`.                                                                                                                                                        |
| `Archive.aesEncrypt`                         | Use `Crypto.encryptAes`.                                                                                                                                                        |
| `Archive.aesEncryptedSize`                   | Use `Crypto.getAesEncryptedSize`.                                                                                                                                               |
| `Archive.buildAesExtraField`                 | Use `Crypto.createAesExtraField`.                                                                                                                                               |
| `Archive.COMPRESSION_METHOD_AES`             | Use `Crypto.COMPRESSION_METHOD_AES`.                                                                                                                                            |
| `Archive.createTarArchive`                   | Use `Archive.createTar`.                                                                                                                                                        |
| `Archive.createTarReader`                    | Use `Archive.openTar`.                                                                                                                                                          |
| `Archive.encryptionMethodFromAesKeyStrength` | Use `Crypto.getEncryptionMethodForAesKeyStrength`.                                                                                                                              |
| `Archive.getAesKeyStrength`                  | Use `Crypto.getAesKeyStrength`.                                                                                                                                                 |
| `Archive.getEncryptionMethodName`            | Use `Crypto.getEncryptionMethodName`.                                                                                                                                           |
| `Archive.isAesEncryption`                    | Use `Crypto.isAesEncryption`.                                                                                                                                                   |
| `Archive.randomBytes`                        | Use `crypto.getRandomValues(new Uint8Array(n))`.                                                                                                                                |
| `Archive.tar`                                | Use `Archive.packTar`.                                                                                                                                                          |
| `Archive.tarSync`                            | Use `Archive.packTarSync`.                                                                                                                                                      |
| `Archive.untar`                              | Use `Archive.unpackTar`.                                                                                                                                                        |
| `Archive.unzip`                              | Use `Archive.openZip`, or `Archive.openTar` for `{ format: "tar" }`. `unzip` named a reader while `untar` named a map.                                                          |
| `Archive.zip`                                | Use `Archive.createZip`, or `Archive.createTar` for `{ format: "tar" }`. `zip` named a builder while `tar` named bytes.                                                         |
| `Archive.ZIP_CRYPTO_HEADER_SIZE`             | Use `Crypto.ZIP_CRYPTO_HEADER_SIZE`.                                                                                                                                            |
| `Archive.zipCryptoDecrypt`                   | Use `Crypto.decryptZipCrypto`, which throws an `ArchiveError` where this returns `null`.                                                                                        |
| `Archive.zipCryptoEncrypt`                   | Use `Crypto.encryptZipCrypto`.                                                                                                                                                  |
| `Archive.zipCryptoInitKeys`                  | Use `Crypto.createZipCryptoKeys`.                                                                                                                                               |
| `ArchiveFile`                                | Use `Archive.ArchiveFile`.                                                                                                                                                      |
| `compress`                                   | Use `Compression.deflate`.                                                                                                                                                      |
| `compressSync`                               | Use `Compression.deflateSync`.                                                                                                                                                  |
| `copyFile`                                   | Use `Fs.copyFile`.                                                                                                                                                              |
| `copyFileSync`                               | Use `Fs.copyFileSync`.                                                                                                                                                          |
| `crc32`                                      | Use `Archive.crc32`.                                                                                                                                                            |
| `crc32Finalize`                              | Use `Archive.crc32Finalize`.                                                                                                                                                    |
| `crc32Update`                                | Use `Archive.crc32Update`.                                                                                                                                                      |
| `createDeflateStream`                        | Use `Compression.createDeflateStream`.                                                                                                                                          |
| `createGunzipStream`                         | Use `Compression.createGunzipStream`.                                                                                                                                           |
| `createGzipStream`                           | Use `Compression.createGzipStream`.                                                                                                                                             |
| `createInflateStream`                        | Use `Compression.createInflateStream`.                                                                                                                                          |
| `createReadStream`                           | Use `createReadStream` from `node:fs`.                                                                                                                                          |
| `createTempDir`                              | Use `Fs.createTempDir`.                                                                                                                                                         |
| `createTempDirSync`                          | Use `Fs.createTempDirSync`.                                                                                                                                                     |
| `createUnzlibStream`                         | Use `Compression.createUnzlibStream`.                                                                                                                                           |
| `createWriteStream`                          | Use `createWriteStream` from `node:fs`.                                                                                                                                         |
| `createZlibStream`                           | Use `Compression.createZlibStream`.                                                                                                                                             |
| `decompress`                                 | Use `Compression.inflate`.                                                                                                                                                      |
| `decompressAuto`                             | Use `Compression.decompressAuto`.                                                                                                                                               |
| `decompressAutoSync`                         | Use `Compression.decompressAutoSync`.                                                                                                                                           |
| `decompressSync`                             | Use `Compression.inflateSync`.                                                                                                                                                  |
| `deflateBatchWithPool`                       | Use `Compression.Workers.deflateBatch`.                                                                                                                                         |
| `deflateWithPool`                            | Use `Compression.Workers.deflate`.                                                                                                                                              |
| `detectCompressionFormat`                    | Use `Compression.detectFormat`.                                                                                                                                                 |
| `encodePng`                                  | Use `Archive.encodePng`.                                                                                                                                                        |
| `ensureDir`                                  | Use `Fs.ensureDir`.                                                                                                                                                             |
| `ensureDirSync`                              | Use `Fs.ensureDirSync`.                                                                                                                                                         |
| `FileEntry`                                  | Use `FsFileEntry`.                                                                                                                                                              |
| `fileExists`                                 | Use `Fs.fileExists`.                                                                                                                                                            |
| `fileExistsSync`                             | Use `Fs.fileExistsSync`.                                                                                                                                                        |
| `getDefaultWorkerPool`                       | Use `Compression.Workers.getDefaultPool`.                                                                                                                                       |
| `glob`                                       | Use `Fs.glob`.                                                                                                                                                                  |
| `GlobOptions`                                | Use `FsGlobOptions`.                                                                                                                                                            |
| `globSync`                                   | Use `Fs.globSync`.                                                                                                                                                              |
| `globToRegex`                                | Use `Fs.globToRegex`.                                                                                                                                                           |
| `gunzip`                                     | Use `Compression.gunzip`.                                                                                                                                                       |
| `gunzipSync`                                 | Use `Compression.gunzipSync`.                                                                                                                                                   |
| `gzip`                                       | Use `Compression.gzip`.                                                                                                                                                         |
| `gzipSync`                                   | Use `Compression.gzipSync`.                                                                                                                                                     |
| `hasCompressionStream`                       | Use `Compression.hasCompressionStream`.                                                                                                                                         |
| `hasDeflateRaw`                              | Use `Compression.hasDeflateRaw`.                                                                                                                                                |
| `hasGzipCompressionStream`                   | Use `Compression.hasGzipCompressionStream`.                                                                                                                                     |
| `hasGzipDecompressionStream`                 | Use `Compression.hasGzipDecompressionStream`.                                                                                                                                   |
| `hasWorkerSupport`                           | Use `Compression.hasWorkerSupport`.                                                                                                                                             |
| `inflateBatchWithPool`                       | Use `Compression.Workers.inflateBatch`.                                                                                                                                         |
| `inflateWithPool`                            | Use `Compression.Workers.inflate`.                                                                                                                                              |
| `isGzipData`                                 | Use `Compression.isGzipData`.                                                                                                                                                   |
| `isZlibData`                                 | Use `Compression.isZlibData`.                                                                                                                                                   |
| `matchGlob`                                  | Use `Fs.matchGlob`.                                                                                                                                                             |
| `matchGlobAny`                               | Use `Fs.matchGlobAny`.                                                                                                                                                          |
| `parseTarGz`                                 | Use `Archive.parseTarGz`.                                                                                                                                                       |
| `parseTarGzStream`                           | Use `Archive.parseTarGzStream`, which now parses incrementally. For the old behavior (decompress the whole archive, fail before the first entry), iterate `Archive.parseTarGz`. |
| `readFileBytes`                              | Use `Fs.readFileBytes`.                                                                                                                                                         |
| `readFileBytesSync`                          | Use `Fs.readFileBytesSync`.                                                                                                                                                     |
| `readFileText`                               | Use `Fs.readFileText`.                                                                                                                                                          |
| `readFileTextSync`                           | Use `Fs.readFileTextSync`.                                                                                                                                                      |
| `ReadStreamOptions`                          | Use `ReadStreamOptions` from `node:fs`.                                                                                                                                         |
| `remove`                                     | Use `Fs.remove`.                                                                                                                                                                |
| `removeSync`                                 | Use `Fs.removeSync`.                                                                                                                                                            |
| `safeStats`                                  | Use `Fs.safeStats`.                                                                                                                                                             |
| `safeStatsSync`                              | Use `Fs.safeStatsSync`.                                                                                                                                                         |
| `setFileTime`                                | Use `Fs.setFileTime`.                                                                                                                                                           |
| `setFileTimeSync`                            | Use `Fs.setFileTimeSync`.                                                                                                                                                       |
| `targz`                                      | Use `Archive.packTarGz`.                                                                                                                                                        |
| `TarGzArchive`                               | Use `Archive.TarGzArchive` (or build one with `Archive.createTarGz`).                                                                                                           |
| `TaskOptions`                                | Use `WorkerTaskOptions`.                                                                                                                                                        |
| `TaskResult`                                 | Use `WorkerTaskResult`.                                                                                                                                                         |
| `terminateDefaultWorkerPool`                 | Use `Compression.Workers.terminateDefaultPool`.                                                                                                                                 |
| `toNodeReadable`                             | Use `Archive.toNodeReadable`.                                                                                                                                                   |
| `traverseDirectory`                          | Use `Fs.traverseDirectory`.                                                                                                                                                     |
| `traverseDirectorySync`                      | Use `Fs.traverseDirectorySync`.                                                                                                                                                 |
| `TraverseOptions`                            | Use `FsTraverseOptions`.                                                                                                                                                        |
| `untargz`                                    | Use `Archive.unpackTarGz`.                                                                                                                                                      |
| `unzlib`                                     | Use `Compression.unzlib`.                                                                                                                                                       |
| `unzlibSync`                                 | Use `Compression.unzlibSync`.                                                                                                                                                   |
| `WorkerPool`                                 | Use `Compression.Workers.Pool`.                                                                                                                                                 |
| `writeFileBytes`                             | Use `Fs.writeFileBytes`.                                                                                                                                                        |
| `writeFileBytesSync`                         | Use `Fs.writeFileBytesSync`.                                                                                                                                                    |
| `writeFileText`                              | Use `Fs.writeFileText`.                                                                                                                                                         |
| `writeFileTextSync`                          | Use `Fs.writeFileTextSync`.                                                                                                                                                     |
| `WriteStreamOptions`                         | Use `WriteStreamOptions` from `node:fs`.                                                                                                                                        |
| `zlib`                                       | Use `Compression.zlib`.                                                                                                                                                         |
| `zlibSync`                                   | Use `Compression.zlibSync`.                                                                                                                                                     |

</details>

<details>
<summary><code>documonster/csv</code> — 40 names</summary>

| Deprecated                     | Use instead                                                                                                                                         |
| ------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| `BuiltinColumnType`            | Use `CsvBuiltinColumnType`.                                                                                                                         |
| `CastDateConfig`               | Use `CsvCastDateConfig`.                                                                                                                            |
| `ChunkMeta`                    | Use `CsvChunkMeta`.                                                                                                                                 |
| `ColumnConfig`                 | Use `CsvColumnConfig`.                                                                                                                              |
| `ColumnDef`                    | Use `CsvColumnDef`.                                                                                                                                 |
| `ColumnMismatchConfig`         | Use `CsvColumnMismatchConfig`.                                                                                                                      |
| `ColumnMismatchLess`           | Use `CsvColumnMismatchLess`.                                                                                                                        |
| `ColumnMismatchMore`           | Use `CsvColumnMismatchMore`.                                                                                                                        |
| `Csv.generateAsync`            | Use `Csv.generateRowsAsync`: it yields rows like `Csv.generateRows`, and `*Async` alone names a Promise elsewhere in this namespace (`parseAsync`). |
| `Csv.parseWithProgress`        | Use `Csv.parseAsync` with `{ onProgress }`.                                                                                                         |
| `Csv.quoted`                   | Use `Csv.forceQuote`.                                                                                                                               |
| `Csv.rowHashArrayMapByHeaders` | Use `Csv.getRowValuesByHeaders`.                                                                                                                    |
| `Csv.rowHashArrayToHeaders`    | Use `Csv.getRowHeaders`.                                                                                                                            |
| `Csv.rowHashArrayToValues`     | Use `Csv.getRowValues`.                                                                                                                             |
| `Csv.unquoted`                 | Use `Csv.preventQuote`.                                                                                                                             |
| `DecimalSeparator`             | Use `CsvDecimalSeparator`.                                                                                                                          |
| `DynamicTypingConfig`          | Use `CsvDynamicTypingConfig`.                                                                                                                       |
| `FormattedValue`               | Use `CsvFormattedValue`.                                                                                                                            |
| `GeneratorColumnConfig`        | Use `CsvGeneratorColumnConfig`.                                                                                                                     |
| `GeneratorContext`             | Use `CsvGeneratorContext`.                                                                                                                          |
| `GeneratorFn`                  | Use `CsvGeneratorFn`.                                                                                                                               |
| `HeaderArray`                  | Use `CsvHeaderArray`.                                                                                                                               |
| `HeaderTransformFunction`      | Use `CsvHeaderTransformFunction`.                                                                                                                   |
| `OnSkipCallback`               | Use `CsvOnSkipCallback`.                                                                                                                            |
| `ParsedRow`                    | Use `CsvParsedRow`.                                                                                                                                 |
| `RecordInfo`                   | Use `CsvRecordInfo`.                                                                                                                                |
| `RecordWithInfo`               | Use `CsvRecordWithInfo`.                                                                                                                            |
| `Row`                          | Use `CsvRow`.                                                                                                                                       |
| `RowArray`                     | Use `CsvRowArray`.                                                                                                                                  |
| `RowHashArray`                 | Use `CsvRowHashArray`.                                                                                                                              |
| `RowMap`                       | Use `CsvRowMap`.                                                                                                                                    |
| `RowTransformCallback`         | Use `CsvRowTransformCallback`.                                                                                                                      |
| `RowTransformFunction`         | Use `CsvRowTransformFunction`.                                                                                                                      |
| `RowValidateCallback`          | Use `CsvRowValidateCallback`.                                                                                                                       |
| `RowValidateFunction`          | Use `CsvRowValidateFunction`.                                                                                                                       |
| `StopCondition`                | Use `CsvStopCondition`.                                                                                                                             |
| `StopContext`                  | Use `CsvStopContext`.                                                                                                                               |
| `TransformContext`             | Use `CsvTransformContext`.                                                                                                                          |
| `TransformResult`              | Use `CsvTransformResult`.                                                                                                                           |
| `TypeTransformMap`             | Use `CsvTypeTransformMap`.                                                                                                                          |

</details>

<details>
<summary><code>documonster/mermaid</code> — 7 names</summary>

| Deprecated           | Use instead                                                                                                 |
| -------------------- | ----------------------------------------------------------------------------------------------------------- |
| `centredText`        | Use `Mermaid.createCenteredText`, which takes the font as one style object.                                 |
| `layoutFlowchart`    | Use `Mermaid.layoutGraph`, which also lays out state, class, ER, requirement, C4 and architecture diagrams. |
| `MermaidSyntaxError` | Use `MermaidParseError` — the same class, which also carries `line` and `column`.                           |
| `mermaidToDrawList`  | Use `Mermaid.toDrawList`.                                                                                   |
| `mermaidToSvg`       | Use `Mermaid.toSvg`.                                                                                        |
| `parseMermaid`       | Use `Mermaid.parse`.                                                                                        |
| `resolveTheme`       | Use `Mermaid.resolveTheme`.                                                                                 |

</details>

<details>
<summary><code>documonster/stream</code> — 55 names</summary>

| Deprecated                                | Use instead                                                                                                             |
| ----------------------------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| `BufferedStream`                          | Use `Streams.BufferedStream` (or build one with `Streams.createBufferedStream`).                                        |
| `bytes`                                   | Use `Streams.bytes`.                                                                                                    |
| `ChunkedBuilder`                          | Use `Streams.ChunkedBuilder`.                                                                                           |
| `collect`                                 | Use `Streams.collect`.                                                                                                  |
| `Collector`                               | Use `Streams.Collector` (or build one with `Streams.createCollector`).                                                  |
| `copyStream`                              | Use `pipeline`, which this duplicated.                                                                                  |
| `createBufferedStream`                    | Use `Streams.createBufferedStream`.                                                                                     |
| `createByteChunk`                         | Use `Streams.createByteChunk`.                                                                                          |
| `createCollector`                         | Use `Streams.createCollector`.                                                                                          |
| `createDuplex`                            | Use `Streams.createDuplex`.                                                                                             |
| `createEmptyReadable`                     | Use `Streams.createEmptyReadable`.                                                                                      |
| `createNullWritable`                      | Use `Streams.createNullWritable`.                                                                                       |
| `createPassThrough`                       | Use `Streams.createPassThrough`.                                                                                        |
| `createPullStream`                        | Use `Streams.createPullStream`.                                                                                         |
| `createReadable`                          | Use `Streams.createReadable`.                                                                                           |
| `createReadableFromArray`                 | Use `Streams.createReadableFromArray`.                                                                                  |
| `createReadableFromAsyncIterable`         | Use `Streams.createReadableFromAsyncIterable`.                                                                          |
| `createReadableFromGenerator`             | Use `Streams.createReadableFromGenerator`.                                                                              |
| `createReadableFromPromise`               | Use `Streams.createReadableFromPromise`.                                                                                |
| `createStringChunk`                       | Use `Streams.createStringChunk`.                                                                                        |
| `createTransform`                         | Use `Streams.createTransform`.                                                                                          |
| `createWritable`                          | Use `Streams.createWritable`.                                                                                           |
| `drainStream`                             | Use `Streams.discard`.                                                                                                  |
| `eventedReadableToAsyncIterableNoDestroy` | Use `Streams.LowLevel.eventsToAsyncIterableNoDestroy`.                                                                  |
| `filter`                                  | Use `Streams.filter`.                                                                                                   |
| `finishedAll`                             | Use `Streams.waitForFinished`.                                                                                          |
| `fromBytes`                               | Use `Streams.fromBytes`.                                                                                                |
| `fromJSON`                                | Use `Streams.fromJson`.                                                                                                 |
| `fromString`                              | Use `Streams.fromString`.                                                                                               |
| `isAsyncIterable`                         | Use `Streams.isAsyncIterable`.                                                                                          |
| `isDuplex`                                | Use `Streams.isDuplex`.                                                                                                 |
| `isReadableStream`                        | Use `Streams.isReadableStream`.                                                                                         |
| `isReadableStreamLike`                    | Use `Streams.isReadableStream`, which this duplicated.                                                                  |
| `isStream`                                | Use `Streams.isStream`.                                                                                                 |
| `isTransform`                             | Use `Streams.isTransform`.                                                                                              |
| `isTransformStream`                       | Use `Streams.isTransformStream`.                                                                                        |
| `isWritableStream`                        | Use `Streams.isWritableStream`.                                                                                         |
| `json`                                    | Use `Streams.json`.                                                                                                     |
| `onceEvent`                               | Use `Streams.LowLevel.onceEvent`.                                                                                       |
| `promisify`                               | Use `Streams.LowLevel.fromCallback`: `promisify` read as `node:util`'s, which wraps a function rather than running one. |
| `PullStream`                              | Use `Streams.PullStream` (or build one with `Streams.createPullStream`).                                                |
| `readableFromWeb`                         | Use `Streams.fromWebReadable`.                                                                                          |
| `readableStreamToAsyncIterable`           | Use `Streams.readableStreamToAsyncIterable`.                                                                            |
| `readableToWeb`                           | Use `Streams.toWebReadable`.                                                                                            |
| `SinkAdapter`                             | Use `Streams.LowLevel.SinkAdapter`.                                                                                     |
| `streamToBuffer`                          | Use `Streams.bytes`, which this duplicated: both resolve to a `Uint8Array`, not a `Buffer`.                             |
| `streamToPromise`                         | Use `finished`, which this duplicated.                                                                                  |
| `streamToString`                          | Use `Streams.text`; pass the encoding as `{ encoding }`.                                                                |
| `streamToUint8Array`                      | Use `Streams.bytes`.                                                                                                    |
| `text`                                    | Use `Streams.text`.                                                                                                     |
| `toWritable`                              | Use `Streams.toWritable`.                                                                                               |
| `TransactionalChunkedBuilder`             | Use `Streams.TransactionalChunkedBuilder`.                                                                              |
| `transform`                               | Use `Streams.createTransform`, which this duplicated.                                                                   |
| `writableFromWeb`                         | Use `Streams.fromWebWritable`.                                                                                          |
| `writableToWeb`                           | Use `Streams.toWebWritable`.                                                                                            |

</details>

<details>
<summary><code>documonster/pdf</code> — 8 names</summary>

| Deprecated                          | Use instead                                                                                                                      |
| ----------------------------------- | -------------------------------------------------------------------------------------------------------------------------------- |
| `createPdfDrawSurface`              | Use `Pdf.createDrawSurface(page, rect, { scale })`.                                                                              |
| `Pdf.asn1Parse`                     | Use `Pdf.Signature.parseAsn1`.                                                                                                   |
| `Pdf.buildSignatureDictPlaceholder` | Use `Pdf.Signature.buildDictPlaceholder`.                                                                                        |
| `Pdf.create`                        | Use `Pdf.fromTable`. `create` is the name every other namespace uses for a handle constructor, and this returns bytes.           |
| `Pdf.read`                          | Use `Pdf.extract`. `read` is the verb that opens a document as its editable object (`Pdf.Editor.read`); this extracts content.   |
| `Pdf.sign`                          | Use `Pdf.Signature.fillPlaceholder(bytes, { certificate, privateKey })` — this only completes a placeholder already in the file. |
| `Pdf.verifySignature`               | Use `Pdf.Signature.verify`, or `Pdf.verifySignatures` to find and check every signature in a file.                               |
| `Pdf.wordChartRenderer`             | Omit `chartRenderer`: the default draws every chart with the Excel engine.                                                       |

</details>

<details>
<summary><code>documonster/formula</code> — 1 name</summary>

| Deprecated      | Use instead           |
| --------------- | --------------------- |
| `Formula.print` | Use `Formula.format`. |

</details>

<details>
<summary><code>documonster/draw</code> — 32 names</summary>

| Deprecated            | Use instead                                                                       |
| --------------------- | --------------------------------------------------------------------------------- |
| `apply`               | Use `MatrixNs.applyMatrix.apply`.                                                 |
| `arcToCubics`         | Use `Draw.Geometry.arcToCubics`.                                                  |
| `BasicRasterCanvas`   | Use `Draw.Backend.BasicRasterCanvas`.                                             |
| `createRasterSurface` | Use `Draw.Backend.createRasterSurface`, which takes `{ scale }`.                  |
| `cssColour`           | Use `Draw.parseColor`.                                                            |
| `downsample`          | Use `Draw.Backend.downsample`.                                                    |
| `flattenPath`         | Use `Draw.Geometry.flattenPath`, which takes `{ scale }`.                         |
| `fontHasGlyph`        | Use `Draw.Backend.fontHasGlyph`.                                                  |
| `IDENTITY`            | Use `MatrixNs.IDENTITYMatrix.IDENTITY`.                                           |
| `isSimpleText`        | Use `Draw.isSimpleText`.                                                          |
| `measureText`         | Use `Draw.measureText`.                                                           |
| `multiply`            | Use `MatrixNs.multiplyMatrix.multiply`.                                           |
| `normalizeSamples`    | Use `Draw.Backend.normalizeSamples`.                                              |
| `parseRasterFont`     | Use `Draw.parseRasterFont`, which takes the bytes or `{ data, collectionIndex }`. |
| `rasterizeToRgba`     | Use `Draw.toRgba`.                                                                |
| `rectNode`            | Use `Draw.Geometry.createRectNode`.                                               |
| `relativeLuminance`   | Use `Draw.getRelativeLuminance`.                                                  |
| `renderDrawList`      | Use `Draw.render`.                                                                |
| `renderNode`          | Use `Draw.Backend.renderNode`.                                                    |
| `rotate`              | Use `MatrixNs.rotateMatrix.rotate`.                                               |
| `rotationOf`          | Use `MatrixNs.getRotationMatrix.getRotation`.                                     |
| `roundedRectToPath`   | Use `Draw.Geometry.roundedRectToPath`.                                            |
| `scale`               | Use `MatrixNs.scaleMatrix.scale`.                                                 |
| `sectorToPath`        | Use `Draw.Geometry.sectorToPath`.                                                 |
| `SvgSurface`          | Use `Draw.Backend.SvgSurface`.                                                    |
| `textFeaturesOf`      | Use `Draw.getTextFeatures`.                                                       |
| `toSvg`               | Use `Draw.toSvg`.                                                                 |
| `translate`           | Use `MatrixNs.translateMatrix.translate`.                                         |
| `translucent`         | Use `Draw.withAlpha`.                                                             |
| `uniformScale`        | Use `MatrixNs.getUniformScaleMatrix.getUniformScale`.                             |
| `widestText`          | Use `Draw.measureWidest`.                                                         |
| `wrapText`            | Use `Draw.wrapText`.                                                              |

</details>

<details>
<summary><code>documonster/chart</code> — 48 names</summary>

| Deprecated                    | Use instead                                                                                            |
| ----------------------------- | ------------------------------------------------------------------------------------------------------ |
| `applyAxisTransform`          | Use `ChartClassic.Backend.applyAxisTransform`.                                                         |
| `applyChartExPreset`          | Use `ChartEx.applyPreset`.                                                                             |
| `applyChartPreset`            | Use `ChartClassic.applyPreset`.                                                                        |
| `buildChartColors`            | Use `ChartClassic.toColorsXml`.                                                                        |
| `buildChartExModel`           | Use `ChartEx.createModel`.                                                                             |
| `buildChartModel`             | Use `ChartClassic.createModel`.                                                                        |
| `buildChartScene`             | Use `ChartClassic.Backend.toScene`.                                                                    |
| `buildChartStyle`             | Use `ChartClassic.toStyleXml`.                                                                         |
| `buildComboChartModel`        | Use `ChartClassic.createComboModel`.                                                                   |
| `buildEffectFilter`           | Use `ChartClassic.Backend.toSvgFilter`.                                                                |
| `buildSpPr`                   | Use `DrawingMl.createSpPr`.                                                                            |
| `buildTxPr`                   | Use `DrawingMl.createTxPr`.                                                                            |
| `canRenderChartExAsVectorPdf` | Nothing to ask: `ChartEx.toDrawList` draws every layout, and `Pdf.fromChart` picks its path itself.    |
| `ChartPdfColor`               | A type of the deprecated `drawChartPdf` / `drawChartExPdf`; draw a chart's display list instead.       |
| `ChartPdfDrawingSurface`      | A chart goes onto a PDF page through `Pdf.createDrawSurface` (a `PdfDrawPage`).                        |
| `ChartPdfPathOp`              | As `ChartPdfDrawingSurface`.                                                                           |
| `chartExOptionsFromRows`      | Use `ChartEx.createOptionsFromRows`.                                                                   |
| `chartExOptionsFromTable`     | Use `ChartEx.createOptionsFromTable`.                                                                  |
| `chartOptionsFromRows`        | Use `ChartClassic.createOptionsFromRows`.                                                              |
| `chartOptionsFromTable`       | Use `ChartClassic.createOptionsFromTable`.                                                             |
| `drawChartExPdf`              | `Draw.render(ChartEx.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`.        |
| `drawChartPdf`                | `Draw.render(ChartClassic.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`.   |
| `fillChartCaches`             | Use `ChartClassic.fillCaches`, which takes `{ worksheet }` instead of a positional `contextWorksheet`. |
| `fillChartExCaches`           | Use `ChartEx.fillCaches`, which takes `{ worksheet }` instead of a positional `contextWorksheet`.      |
| `fillNumRef`                  | Use `ChartClassic.fillNumRef(ref, workbook, { date1904, worksheet })`.                                 |
| `fillStrRef`                  | Use `ChartClassic.fillStrRef(ref, workbook, { worksheet })`.                                           |
| `getSpPrFillColor`            | Use `DrawingMl.getSpPrFillColor`.                                                                      |
| `getSpPrGradient`             | Use `DrawingMl.getSpPrGradient`.                                                                       |
| `getSpPrLine`                 | Use `DrawingMl.getSpPrLine`.                                                                           |
| `getSpPrPattern`              | Use `DrawingMl.getSpPrPattern`.                                                                        |
| `getTxPrColor`                | Use `DrawingMl.getTxPrColor`.                                                                          |
| `getTxPrFontSize`             | Use `DrawingMl.getTxPrFontSize`.                                                                       |
| `parseChartColors`            | Use `ChartClassic.fromColorsXml`.                                                                      |
| `parseChartEx`                | Use `ChartEx.fromXml`.                                                                                 |
| `parseChartStyle`             | Use `ChartClassic.fromStyleXml`.                                                                       |
| `parseSpPr`                   | Use `DrawingMl.parseSpPr`.                                                                             |
| `parseTxPr`                   | Use `DrawingMl.parseTxPr`.                                                                             |
| `PdfChartRenderOptions`       | Pass `ChartRenderOptions` to `ChartClassic.toDrawList` and the box to `Pdf.createDrawSurface`.         |
| `PdfColor`                    | As `ChartPdfColor`.                                                                                    |
| `renderChartEx`               | Use `ChartEx.toXml` — this returns the part's XML, not a picture.                                      |
| `renderChartExPng`            | Use `ChartEx.toPng`.                                                                                   |
| `renderChartExSvg`            | Use `ChartEx.toSvg`.                                                                                   |
| `renderChartPng`              | Use `ChartClassic.toPng`.                                                                              |
| `renderChartSvg`              | Use `ChartClassic.toSvg`.                                                                              |
| `resolveTopologyObject`       | Use `ChartClassic.Backend.resolveTopologyObject`.                                                      |
| `seriesFromColumns`           | Use `ChartClassic.createSeriesFromColumns`.                                                            |
| `setSpPrFill`                 | Use `DrawingMl.withSpPrFill` (returns a copy).                                                         |
| `setSpPrLine`                 | Use `DrawingMl.withSpPrLine` (returns a copy).                                                         |

</details>

<details>
<summary><code>documonster/excel/csv</code> — 7 names</summary>

| Deprecated             | Use instead                                                                                                                 |
| ---------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `createCsvReadStream`  | Use `ExcelCsv.toStream`.                                                                                                    |
| `createCsvWriteStream` | Use `ExcelCsv.createImportStream`.                                                                                          |
| `readCsv`              | Use `ExcelCsv.addWorksheet`.                                                                                                |
| `readCsvFile`          | Use `ExcelCsv.readFile` for a new workbook, or `ExcelCsv.addWorksheet` with a file stream to add the sheet to one you have. |
| `writeCsv`             | Use `ExcelCsv.format` for a string or `ExcelCsv.writeStream` for a stream.                                                  |
| `writeCsvBuffer`       | Use `ExcelCsv.toBuffer`.                                                                                                    |
| `writeCsvFile`         | Use `ExcelCsv.writeFile`.                                                                                                   |

</details>

<details>
<summary><code>documonster/excel/markdown</code> — 7 names</summary>

| Deprecated            | Use instead                                                                                                                                                          |
| --------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `readMarkdown`        | Use `ExcelMarkdown.addWorksheet`, which types the data cells as `ExcelCsv` does; this keeps them as text.                                                            |
| `readMarkdownAll`     | Use `ExcelMarkdown.addWorksheets`, which types the data cells as `ExcelCsv` does; this keeps them as text.                                                           |
| `readMarkdownAllFile` | Use `ExcelMarkdown.readFile(path, { tables: "all" })` for a new workbook, or `ExcelMarkdown.addWorksheets`; both type the data cells, where this keeps them as text. |
| `readMarkdownFile`    | Use `ExcelMarkdown.readFile` for a new workbook, or `ExcelMarkdown.addWorksheet`; both type the data cells, where this keeps them as text.                           |
| `writeMarkdown`       | Use `ExcelMarkdown.format`.                                                                                                                                          |
| `writeMarkdownBuffer` | Use `ExcelMarkdown.toBufferSync` (or the asynchronous `ExcelMarkdown.toBuffer`).                                                                                     |
| `writeMarkdownFile`   | Use `ExcelMarkdown.writeFile`.                                                                                                                                       |

</details>

<details>
<summary><code>documonster/excel/formula</code> — 1 name</summary>

| Deprecated          | Use instead                   |
| ------------------- | ----------------------------- |
| `calculateFormulas` | Use `ExcelFormula.calculate`. |

</details>

<details>
<summary><code>documonster/xml</code> — 6 names</summary>

| Deprecated            | Use instead             |
| --------------------- | ----------------------- |
| `Xml.attr`            | Use `Xml.getAttr`.      |
| `Xml.findChildren`    | Use `Xml.listChildren`. |
| `Xml.saxStream`       | Use `Xml.feedSax`.      |
| `Xml.textContent`     | Use `Xml.getText`.      |
| `Xml.validateXmlName` | Use `Xml.validateName`. |
| `Xml.walk`            | Use `Xml.eachElement`.  |

</details>

<details>
<summary><code>documonster/word</code> — 90 names</summary>

| Deprecated                         | Use instead                                                                                                                                                                   |
| ---------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Build.checkBox`                   | Use `Build.checkbox`.                                                                                                                                                         |
| `Build.createArrow`                | Use `Build.arrow`.                                                                                                                                                            |
| `Build.createCallout`              | Use `Build.callout`.                                                                                                                                                          |
| `Build.createEllipse`              | Use `Build.ellipse`.                                                                                                                                                          |
| `Build.createFlowchartShape`       | Use `Build.flowchartShape`.                                                                                                                                                   |
| `Build.createLine`                 | Use `Build.line`.                                                                                                                                                             |
| `Build.createRect`                 | Use `Build.rect`.                                                                                                                                                             |
| `Build.createRoundRect`            | Use `Build.roundRect`.                                                                                                                                                        |
| `Build.createShape`                | Use `Build.shape`.                                                                                                                                                            |
| `Build.createStar`                 | Use `Build.star`.                                                                                                                                                             |
| `Build.mathPreSubSuperScript`      | Use `Build.mathPreSubSuperscript`.                                                                                                                                            |
| `Build.mathSubScript`              | Use `Build.mathSubscript`.                                                                                                                                                    |
| `Build.mathSubSuperScript`         | Use `Build.mathSubSuperscript`.                                                                                                                                               |
| `Build.mathSuperScript`            | Use `Build.mathSuperscript`.                                                                                                                                                  |
| `Convert.docxToSemantic`           | Use `Convert.toSemantic`.                                                                                                                                                     |
| `Convert.mapDocument`              | Use `Transform.map`.                                                                                                                                                          |
| `Convert.mathMLToOmml`             | Use `Convert.mathMlToOmml`.                                                                                                                                                   |
| `Convert.ommlToMathML`             | Use `Convert.ommlToMathMl`.                                                                                                                                                   |
| `Convert.parseFlatOpc`             | Use `Convert.flatOpcToParts`, the inverse of `Convert.partsToFlatOpc`, or `Document.read`, which now reads Flat OPC bytes as a document.                                      |
| `Convert.readOdt`                  | Use `Convert.fromOdt`.                                                                                                                                                        |
| `Convert.toFlatOpc`                | Use `Convert.partsToFlatOpc` — this wraps a part map, not a document; `Document.toFlatOpc(doc)` is the document form.                                                         |
| `Convert.writeOdt`                 | Use `Convert.toOdt`.                                                                                                                                                          |
| `Diff.documents`                   | Use `Diff.compare`.                                                                                                                                                           |
| `Document.addParagraphElement`     | Use `Document.addParagraph`.                                                                                                                                                  |
| `Document.addTableElement`         | Use `Document.addTable`.                                                                                                                                                      |
| `Document.build`                   | A document needs no build step: pass the document itself to `Document.toBuffer`, `Layout.*`, `WordHtml.fromDocx` or `Pdf.fromDocx`, which apply the same write-time defaults. |
| `Document.DocumentHandle`          | Use `DocxDocument`.                                                                                                                                                           |
| `Document.insertContentAt`         | Use `Document.insertContent`.                                                                                                                                                 |
| `Document.nextBookmarkId`          | Use `Document.allocateBookmarkId`.                                                                                                                                            |
| `DocumentHandle`                   | Use `DocxDocument`.                                                                                                                                                           |
| `Glossary.autoTextEntries`         | Use `Glossary.listAutoTextEntries`.                                                                                                                                           |
| `Glossary.quickParts`              | Use `Glossary.listQuickParts`.                                                                                                                                                |
| `Io.compileTemplate`               | Use `Package.compileTemplate`.                                                                                                                                                |
| `Io.editDocxIncremental`           | Use `Package.edit`.                                                                                                                                                           |
| `Io.fillTemplateFromBuffer`        | Use `Package.fillTemplate`.                                                                                                                                                   |
| `Io.listDocxParts`                 | Use `Package.listParts`.                                                                                                                                                      |
| `Io.merge`                         | Use `Transform.merge`.                                                                                                                                                        |
| `Io.patchDocument`                 | Use `Package.patch`.                                                                                                                                                          |
| `Io.patchTemplate`                 | Use `Package.patchTemplate`.                                                                                                                                                  |
| `Io.readDocxPart`                  | Use `Package.readPart`.                                                                                                                                                       |
| `Io.split`                         | Use `Transform.split`.                                                                                                                                                        |
| `Io.toFlatOpcFromDoc`              | Use `Document.toFlatOpc`, which takes an options object (`{ compressionLevel }`).                                                                                             |
| `Io.updateFields`                  | Use `Transform.updateFields`.                                                                                                                                                 |
| `Io.updateTableOfContents`         | Use `Transform.updateTableOfContents`.                                                                                                                                        |
| `Layout.document`                  | Use `Layout.getPageMap`.                                                                                                                                                      |
| `Layout.documentFull`              | Use `Layout.compute`.                                                                                                                                                         |
| `Layout.renderDocumentToSvg`       | Use `Layout.toSvg`.                                                                                                                                                           |
| `Layout.renderPageFromLayout`      | Use `Layout.computedPageToSvg`.                                                                                                                                               |
| `Layout.renderPageToSvg`           | Use `Layout.pageToSvg`.                                                                                                                                                       |
| `Query.acceptAllRevisions`         | Use `Document.acceptAllRevisions`.                                                                                                                                            |
| `Query.acceptRevision`             | Use `Document.acceptRevision`.                                                                                                                                                |
| `Query.collectParagraphs`          | Use `Query.listParagraphs`.                                                                                                                                                   |
| `Query.collectRuns`                | Use `Query.listRuns`.                                                                                                                                                         |
| `Query.collectTables`              | Use `Query.listTables`, passing the parts to include — `listTables` searches only the body unless asked.                                                                      |
| `Query.countByFormat`              | Use `Query.getCountByFormat`.                                                                                                                                                 |
| `Query.countWords`                 | Use `Query.getWordCount`.                                                                                                                                                     |
| `Query.extractFormFields`          | Use `Query.listFormFields`.                                                                                                                                                   |
| `Query.extractText`                | Use `Query.getText`.                                                                                                                                                          |
| `Query.fillFormFields`             | Use `Transform.fillFormFields`.                                                                                                                                               |
| `Query.findBookmark`               | Use `Query.getBookmark`.                                                                                                                                                      |
| `Query.findComment`                | Use `Query.getComment`.                                                                                                                                                       |
| `Query.getHeadings`                | Use `Query.listHeadings`.                                                                                                                                                     |
| `Query.getUsedFormats`             | Use `Query.listUsedFormats`.                                                                                                                                                  |
| `Query.indexStyles`                | Use `Query.createStyleIndex`.                                                                                                                                                 |
| `Query.mailMerge`                  | Use `Document.mailMerge`.                                                                                                                                                     |
| `Query.paragraphCount`             | Use `Query.getParagraphCount`.                                                                                                                                                |
| `Query.rejectAllRevisions`         | Use `Document.rejectAllRevisions`.                                                                                                                                            |
| `Query.rejectRevision`             | Use `Document.rejectRevision`.                                                                                                                                                |
| `Query.replaceText`                | Use `Document.replaceText`.                                                                                                                                                   |
| `Query.resolveDataBindings`        | Use `Transform.resolveDataBindings`.                                                                                                                                          |
| `Query.resolveNumberingLevel`      | Use `Query.getNumberingLevel`.                                                                                                                                                |
| `Query.resolveRunStyle`            | Use `Query.getRunStyle`.                                                                                                                                                      |
| `Query.resolveStyle`               | Use `Query.getParagraphStyle`.                                                                                                                                                |
| `Query.resolveTableStyle`          | Use `Query.getTableStyle`.                                                                                                                                                    |
| `Query.setCompatibilityMode`       | Use `Document.setCompatibilityMode`.                                                                                                                                          |
| `Query.tableCount`                 | Use `Query.getTableCount`.                                                                                                                                                    |
| `Query.walkBlocks`                 | Use `Query.eachBlock`.                                                                                                                                                        |
| `Query.walkDocument`               | Use `Query.each`.                                                                                                                                                             |
| `Security.encrypt`                 | Use `WordCrypto.encrypt` from `documonster/word/crypto`, beside `decrypt` and `verifyPassword`.                                                                               |
| `Security.resolveSecurityPolicy`   | Use `Security.resolvePolicy`.                                                                                                                                                 |
| `Streaming.createDocxStream`       | Use `Streaming.createWriter`.                                                                                                                                                 |
| `Streaming.createDocxStreamReader` | Use `Streaming.createReader`.                                                                                                                                                 |
| `Template.fillTemplate`            | Use `Template.fill`, which returns a new document.                                                                                                                            |
| `Template.fillTemplateEnhanced`    | Use `Template.fill` with `{ richPlaceholders: true }`.                                                                                                                        |
| `Template.fillTemplateFromSource`  | Use `Template.fillFromSource`.                                                                                                                                                |
| `Template.isTemplateChart`         | Use `Template.isChart`.                                                                                                                                                       |
| `Template.listTemplateTags`        | Use `Template.listTags`.                                                                                                                                                      |
| `Template.TemplateError`           | Import `TemplateError` from `documonster/word` directly.                                                                                                                      |
| `Validation.document`              | Use `Validation.validate`.                                                                                                                                                    |

</details>

<details>
<summary><code>documonster/word/html</code> — 3 names</summary>

| Deprecated       | Use instead                                                                                                                                   |
| ---------------- | --------------------------------------------------------------------------------------------------------------------------------------------- |
| `htmlToDocx`     | Use `WordHtml.toDocxBody`, whose result also carries the styles and numbering the body refers to; `WordHtml.toDocx` returns a whole document. |
| `htmlToDocxBody` | Use `WordHtml.toDocxBody(html).body`.                                                                                                         |
| `renderToHtml`   | Use `WordHtml.fromDocx`.                                                                                                                      |

</details>

<details>
<summary><code>documonster/word/crypto</code> — 17 names</summary>

| Deprecated               | Use instead                                                                                                                                            |
| ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `AGILE_BLOCK_KEYS`       | Use `WordCrypto.Agile.BLOCK_KEYS`.                                                                                                                     |
| `decryptDocx`            | Use `WordCrypto.decrypt`.                                                                                                                              |
| `decryptPackage`         | Use `WordCrypto.Agile.decryptPackage`.                                                                                                                 |
| `deobfuscateFont`        | Use `Font.deobfuscate` from `documonster/word`.                                                                                                        |
| `deriveEncryptionKey`    | Use `WordCrypto.Agile.deriveKey`.                                                                                                                      |
| `encryptDocx`            | Use `WordCrypto.encrypt`.                                                                                                                              |
| `extractSignatures`      | Use `WordCrypto.listSignatures`.                                                                                                                       |
| `generateFontKey`        | Use `Font.generateObfuscationKey` from `documonster/word`.                                                                                             |
| `hasDigitalSignatures`   | Use `WordCrypto.hasSignatures`, which takes the document or, as this did, its part names (`WordCrypto.hasSignatures(await Package.listParts(bytes))`). |
| `isEncryptedDocx`        | Use `WordCrypto.isEncrypted`.                                                                                                                          |
| `isWellFormedSignature`  | Use `WordCrypto.isWellFormedSignature`.                                                                                                                |
| `obfuscateFont`          | Use `Font.obfuscate` from `documonster/word`.                                                                                                          |
| `parseEncryptionInfoXml` | Use `WordCrypto.Agile.parseEncryptionInfo`.                                                                                                            |
| `parseSignatureXml`      | Use `WordCrypto.parseSignature`.                                                                                                                       |
| `readCfb`                | Use `WordCrypto.Cfb.read`.                                                                                                                             |
| `verifyPassword`         | Use `WordCrypto.Agile.verifyPassword`, which takes `(info, password)` instead of `(password, info)`.                                                   |
| `writeCfb`               | Use `WordCrypto.Cfb.write`.                                                                                                                            |

</details>

<details>
<summary><code>documonster/word/markdown</code> — 3 names</summary>

| Deprecated           | Use instead                    |
| -------------------- | ------------------------------ |
| `markdownToDocx`     | Use `WordMarkdown.toDocx`.     |
| `markdownToDocxBody` | Use `WordMarkdown.toDocxBody`. |
| `renderToMarkdown`   | Use `WordMarkdown.fromDocx`.   |

</details>

<details>
<summary><code>documonster/word/excel</code> — 6 names</summary>

| Deprecated                  | Use instead                    |
| --------------------------- | ------------------------------ |
| `buildWordChartExXml`       | Use `WordExcel.toChartExXml`.  |
| `excelToDocx`               | Use `WordExcel.toDocx`.        |
| `extractTablesToExcel`      | Use `WordExcel.extractTables`. |
| `generateChartEmbeddedXlsx` | Use `WordExcel.toChartXlsx`.   |
| `renderWordChartSvg`        | Use `WordExcel.toChartSvg`.    |
| `wordChartToChartModel`     | Use `WordExcel.toChartModel`.  |

</details>

### Removed without a forward

These are the only things published in 0.21.1 that are gone, or narrowed, with no deprecated form left behind; they
are what a compile error after upgrading is most likely to point at. Everything else that was renamed or replaced
still works until 1.0. (The "Removed" counts in the summary table above count rows of the per-name diff — a class
and each of its members, an internal field — so they do not match the number of rows here.)

| Was                                                                                                         | Now                                                                                                                                                                                                                                                                                                             |
| ----------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `CsvWorkerError` (`documonster/csv`)                                                                        | No longer exported. The worker pool that throws it is not public, so nothing outside the module could produce it. It still subclasses `CsvError`: catch `CsvError` / test `isCsvError(error)`, and read `error.name === "CsvWorkerError"` if you need to tell it apart.                                         |
| `PdfMetadata.creationDate` / `modDate` / `pageSize` / `xmpXml`, `PdfExtractedImage.alphaMask` (`… \| null`) | The same fields, now optional and absent instead of `null`. Code that assigned the result to a `{ … : X \| null }` type fails to compile; code that tested `=== null` compiles and stops matching. See [Read limits and PDF fields](#read-limits-pdf-fields-that-are-now-absent-and-other-late-fixes-breaking). |
| `PdfBookmark.pageIndex: number` (`-1` when unresolved)                                                      | Optional; absent when the bookmark does not resolve to a page. Test `=== undefined`, not `=== -1`.                                                                                                                                                                                                              |
| `ZipOptionsTar.password` / `encryptionMethod` (`documonster/archive`)                                       | Removed from the type. TAR has no encryption, and neither option was ever applied: drop them from TAR options.                                                                                                                                                                                                  |
| `new Pdf.EditorPage()` (an implicit public constructor in the 0.21.1 declarations)                          | The constructor is private. A page was never usable unless an editor made it: get one from `Pdf.Editor.read(bytes).getPage(index)`.                                                                                                                                                                             |

Two internal fields also left the published types: `Table.Handle._cache` (Excel) and the `[_documentBrand]` marker on
`DocumentHandle` (Word). Neither was meant to be read or written.

### Also changed: signatures and return values

The themed sections above describe each change with its reason. This is the same ground cut the other way — by what
the compiler sees — for the changes that are easy to miss there. The full old/new signatures are in
[docs/api-changes-since-v0.21.1.md](docs/api-changes-since-v0.21.1.md).

**Byte results are `Uint8Array<ArrayBuffer>`.** Every byte result the library allocates — `toBuffer`, `Compression.*`,
`Archive.packZip` / `unpackZip` data, `stream()` / `operation().iterable` chunks, `Draw.toPng`, `Pdf` extraction,
`splitPages`, … — was typed `Uint8Array` (that is, `Uint8Array<ArrayBufferLike>`), which TypeScript 5.7+ with the DOM
library refuses for `new Blob([bytes])`, `crypto.subtle` and other `BufferSource` parameters. The narrower type removes
those casts. A result is assignable everywhere the old type was, so code that only receives one is unaffected. One case
can need a change: an object you write that implements one of these result interfaces (a custom `ZipOperation`, a
`WorkerTaskResult` mock) must now hold `Uint8Array<ArrayBuffer>` too — `new Uint8Array(bytes)` makes one from any
`Uint8Array`. Types a caller builds and hands in (`DrawRgbaImage`, the document models, `RandomAccessReader`) keep
accepting any `Uint8Array`. At runtime a result is always a plain `Uint8Array` (never a
`Buffer`); the one behavioural change is that bytes from a `SharedArrayBuffer` input — e.g. `Archive.unpackTarSync`
over shared memory — are now copied rather than viewed.

**A return value where there was none, or a more precise one.** Code that ignored the result is unaffected; code that
annotated it as `void` (a callback type, an interface it implements) needs the new type.

| Call                                                               | Returned              | Returns now                                                                                                                                                                              |
| ------------------------------------------------------------------ | --------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Workbook.removeWorksheet(wb, sheet)`                              | `void`                | `boolean` — and `sheet` may be the handle as well as a name or id                                                                                                                        |
| `Workbook.defineCellStyle` (deprecated) / `Workbook.addCellStyle`  | `void`                | the stored `NamedStyleEntry`                                                                                                                                                             |
| `Worksheet.addConditionalFormatting` / `ConditionalFormatting.add` | `void`                | the stored rule                                                                                                                                                                          |
| `Cell.removeName`                                                  | `void`                | `boolean`                                                                                                                                                                                |
| `Table.remove(ws, table)`                                          | `void`                | `boolean` — and `table` may be the handle as well as its name                                                                                                                            |
| `Image.place`                                                      | `void`                | the placed image's handle (what `Image.list` returns, and `Image.remove` takes)                                                                                                          |
| `Image.addShape`                                                   | `void`                | the stored `ShapeModel`                                                                                                                                                                  |
| `DataValidation.remove`                                            | `void`                | `boolean`                                                                                                                                                                                |
| `DefinedNames.add` / `DefinedNames.addFormula`                     | `void`                | the stored `DefinedNameModel`                                                                                                                                                            |
| `DefinedNames.remove`                                              | `void`                | `boolean`; `remove(dn, name, { localSheetId })` deletes the whole name                                                                                                                   |
| `Watermark.remove`, `Stream.WorksheetWriter#removeWatermark`       | `void`                | `boolean`                                                                                                                                                                                |
| `Chart.copyTo`                                                     | part number           | the copy's handle (`handle.chartNumber` is the number)                                                                                                                                   |
| `Csv.createGenerator`                                              | an inline object type | `CsvGenerator`, a named interface (`asyncRows()` is deprecated for `rowsAsync()`)                                                                                                        |
| `Table.get`                                                        | `TableData`           | `TableData \| undefined` — the runtime already answered `undefined`; the type now says so, and the name is matched as Excel matches names (case-insensitively, stored or sanitised form) |

**Handles accepted where only names were.** `Workbook.removeChartsheet`, `copyChartsheet` and
`replaceChartsheetChart` take the chartsheet handle as well as its name or index. Many Excel functions also accept a
streaming writer (`Stream.WorkbookWriter` / `Stream.WorksheetWriter`) beside the document handle — see
[API shape](#api-shape-one-place-to-read-and-write-one-vocabulary-for-streaming-one-address-argument).

**Options objects replace positional settings.** In every case the positional form is still accepted as a deprecated
overload, listed under [Deprecated call forms](#deprecated-call-forms). New option types that appear in signatures:
`WorkbookCreateOptions` (`Workbook.create`), `AddExternalLinkOptions`, `RegisterFunctionOptions`,
`RegisterPersonOptions`, `ProtectWorkbookOptions` / `ProtectSheetOptions` (`{ password, … }`), `FillFormulaOptions`,
`FreezePanesOptions` / `SplitPanesOptions`, `TableRowEditOptions`, `Anchor.CreateAnchorOptions`,
`Note.CreateNoteOptions`, `Cell.SetNoteOptions`, `DefinedNames.NameScopeOptions`, `Row.ValuesOptions` /
`Column.ValuesOptions` (`{ base }`) in Excel; `Build.FieldOptions`, `Build.MathNaryOptions`,
`Build.MathFractionOptions`, `Build.TableOptions`, `Build.SimpleTableOptions`, `DocxAddImageOptions` /
`DocxAddFloatingImageOptions`,
`DocxListTablesOptions` in Word; `RemoteZipExtractOptions`, `ArchiveFileReadOptions` / `ArchiveFileReadTextOptions` in
Archive; `PullUntilOptions` in Stream; `CsvNumberOptions`, `CsvDetectDelimiterOptions`, `CsvDetectLinebreakOptions`,
`CsvDynamicTypingOptions`, `CsvParseAsyncOptions` in CSV; `PdfEncryptionOptions`, `PdfEditorCopyPagesOptions` in PDF;
`DrawRasterLineOptions`, `DrawRasterStrokeOptions`, `DrawRasterTextOptions` for `BasicRasterCanvas`.

**Inputs widened.** Readers take more kinds of bytes — `ArrayBuffer`, any `ArrayBufferView` and (where the platform has
it) `Blob`, beside `Uint8Array`: `Workbook.read`, `Pdf.extract`, `Pdf.Editor.read`, `Document.read`,
`Package.patch`, the archive readers, `Archive.createBufferReader` (`ArchiveBytes`) and `ExcelCsv.read` /
`addWorksheet` (`ExcelCsvInput`, whose type now says what the reader always accepted). Array parameters are
`readonly` throughout, so a `readonly` array no longer needs a cast. `Csv.parseAsync` and `ExcelCsv` take an
`ArrayBuffer`, a view, a `Blob`, a Web `ReadableStream` or an async iterable.

**Types that narrowed.** Each accepted a value that never worked:

- `ZipOptionsTar` no longer has `password` or `encryptionMethod`. TAR has no encryption; they were ignored.
- `DocxResolvedNumberingLevel.format` is `NumberFormat` and `.justification` is `LevelJustification`, not `string`.
- `Stream.WorksheetWriter#workbook` is the `Stream.WorkbookWriter` it belongs to, not a structural stand-in.

**Types that widened** (no action): `DocxEncryptOptions.hashAlgorithm` and `DocxProtectionOptions.hashAlgorithm` also
take Web Crypto's spelling (`"SHA-256"`); `DrawToSvgOptions.background` and `DrawRasterizeOptions.background` take a
CSS color string or a `DrawRgba01` object (`{ r, g, b, a }`, each 0–1); `DocxRenderOptions.renderChart` receives a second `context` argument;
`ArchiveLimit` gained `"maxEntrySize"`; the Word content-tree nodes gained the optional range-marker fields of
`ParagraphChildRangeMarkers` / `BlockRangeMarkers` (bookmarks, comments, permissions and moves that open or close at
the node); `Archive.TarType` is now a value (the constant table) as well as a type.

**Callback forms of the `node:stream` helpers.** `pipeline(…, callback)` returns the last stage and
`finished(stream, [options], callback)` / `streamToPromise(…, callback)` return a cleanup function, as Node's do. Only
the callback-less calls return a promise. Code that passed a callback **and** awaited the result now awaits a
non-promise; drop one of the two.

### `@documonster/mcp`

The MCP server publishes the same 22 tools in the same groups, under the same names, titles and annotations; the
diff of their input schemas (as JSON Schema) and of the TypeScript entry is in
[docs/api-changes-since-v0.21.1.md](docs/api-changes-since-v0.21.1.md#documonstermcp). A client or a prompt written
for 0.21.1 keeps working. What changed:

- **TypeScript entry.** `PdfFontRef` (`{ path, collectionIndex? }`, the shape of a configured PDF font) is exported
  as a type. Nothing was removed.
- **`pdf_edit` — `append` takes `formFields: "copy" | "drop"`.** The default, `"copy"`, keeps the appended pages'
  form fields and refuses a field name the PDF already has; `"drop"` appends the pages without their fields (other
  annotations are kept), which is how the same form is appended twice.
- **`pdf_edit` — page numbers follow the document as the earlier operations in the call left it.** Pages appended
  earlier in a call can now be deleted, rotated or drawn on by their new numbers; that used to be refused ("an
  appended page cannot be removed in the same call").
- **`pdf_edit` — a signed PDF is never silently re-signed away.** Overlay edits (watermark, page numbers, stamp,
  diagram) are appended as an incremental update, classic and cross-reference-stream files alike, so bookmarks, form
  fields and signatures survive. A structural edit (delete, keep, rotate, append) rebuilds the file; on a signed file
  the tool now **refuses and writes nothing**, where it used to rebuild and report the signature as invalid. The
  report of how the file was saved is read off the output ("appended", "unchanged" or "rebuilt"), and a backup is taken
  only once the save has succeeded.
- **`form_fill` fills PDFs with a cross-reference stream.** They were refused, because the rewrite path dropped the
  filled values; the values are now appended as an incremental update (or survive a rewrite), and the output is still
  re-read and checked before anything is written. A signed PDF's earlier revision is left as it was.

### Every deprecation in this release

Every name and call form that became deprecated after 0.21.1, with the replacement its documentation gives. All of
them keep working, with their old behaviour, until 1.0, where they are removed together — so you can migrate in one
pass, at any point before then. Names already deprecated in 0.21.1 are not repeated. A row's replacement is relative
to its own namespace (`Use place` under `Image.*` means `Image.place`).

<!-- api-changes:deprecations:start -->

#### Deprecated names

**`documonster/archive`**

| Deprecated                                   | Replacement                                                                                                                                                                                                                                             |
| -------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `crc32`                                      | Use `Archive.crc32`. Removed in 1.0.                                                                                                                                                                                                                    |
| `crc32Update`                                | Use `Archive.crc32Update`. Removed in 1.0.                                                                                                                                                                                                              |
| `crc32Finalize`                              | Use `Archive.crc32Finalize`. Removed in 1.0.                                                                                                                                                                                                            |
| `compress`                                   | Use `Compression.deflate`. Removed in 1.0.                                                                                                                                                                                                              |
| `compressSync`                               | Use `Compression.deflateSync`. Removed in 1.0.                                                                                                                                                                                                          |
| `decompress`                                 | Use `Compression.inflate`. Removed in 1.0.                                                                                                                                                                                                              |
| `decompressSync`                             | Use `Compression.inflateSync`. Removed in 1.0.                                                                                                                                                                                                          |
| `gzip`                                       | Use `Compression.gzip`. Removed in 1.0.                                                                                                                                                                                                                 |
| `gunzip`                                     | Use `Compression.gunzip`. Removed in 1.0.                                                                                                                                                                                                               |
| `gzipSync`                                   | Use `Compression.gzipSync`. Removed in 1.0.                                                                                                                                                                                                             |
| `gunzipSync`                                 | Use `Compression.gunzipSync`. Removed in 1.0.                                                                                                                                                                                                           |
| `isGzipData`                                 | Use `Compression.isGzipData`. Removed in 1.0.                                                                                                                                                                                                           |
| `hasCompressionStream`                       | Use `Compression.hasCompressionStream`. Removed in 1.0.                                                                                                                                                                                                 |
| `hasWorkerSupport`                           | Use `Compression.hasWorkerSupport`. Removed in 1.0.                                                                                                                                                                                                     |
| `hasGzipCompressionStream`                   | Use `Compression.hasGzipCompressionStream`. Removed in 1.0.                                                                                                                                                                                             |
| `hasGzipDecompressionStream`                 | Use `Compression.hasGzipDecompressionStream`. Removed in 1.0.                                                                                                                                                                                           |
| `zlib`                                       | Use `Compression.zlib`. Removed in 1.0.                                                                                                                                                                                                                 |
| `unzlib`                                     | Use `Compression.unzlib`. Removed in 1.0.                                                                                                                                                                                                               |
| `zlibSync`                                   | Use `Compression.zlibSync`. Removed in 1.0.                                                                                                                                                                                                             |
| `unzlibSync`                                 | Use `Compression.unzlibSync`. Removed in 1.0.                                                                                                                                                                                                           |
| `isZlibData`                                 | Use `Compression.isZlibData`. Removed in 1.0.                                                                                                                                                                                                           |
| `decompressAuto`                             | Use `Compression.decompressAuto`. Removed in 1.0.                                                                                                                                                                                                       |
| `decompressAutoSync`                         | Use `Compression.decompressAutoSync`. Removed in 1.0.                                                                                                                                                                                                   |
| `detectCompressionFormat`                    | Use `Compression.detectFormat`. Removed in 1.0.                                                                                                                                                                                                         |
| `createDeflateStream`                        | Use `Compression.createDeflateStream`. Removed in 1.0.                                                                                                                                                                                                  |
| `createInflateStream`                        | Use `Compression.createInflateStream`. Removed in 1.0.                                                                                                                                                                                                  |
| `createGzipStream`                           | Use `Compression.createGzipStream`. Removed in 1.0.                                                                                                                                                                                                     |
| `createGunzipStream`                         | Use `Compression.createGunzipStream`. Removed in 1.0.                                                                                                                                                                                                   |
| `createZlibStream`                           | Use `Compression.createZlibStream`. Removed in 1.0.                                                                                                                                                                                                     |
| `createUnzlibStream`                         | Use `Compression.createUnzlibStream`. Removed in 1.0.                                                                                                                                                                                                   |
| `hasDeflateRaw`                              | Use `Compression.hasDeflateRaw`. Removed in 1.0.                                                                                                                                                                                                        |
| `WorkerPool`                                 | Use `Compression.Workers.Pool`. Removed in 1.0.                                                                                                                                                                                                         |
| `getDefaultWorkerPool`                       | Use `Compression.Workers.getDefaultPool`. Removed in 1.0.                                                                                                                                                                                               |
| `terminateDefaultWorkerPool`                 | Use `Compression.Workers.terminateDefaultPool`. Removed in 1.0.                                                                                                                                                                                         |
| `deflateWithPool`                            | Use `Compression.Workers.deflate`. Removed in 1.0.                                                                                                                                                                                                      |
| `inflateWithPool`                            | Use `Compression.Workers.inflate`. Removed in 1.0.                                                                                                                                                                                                      |
| `deflateBatchWithPool`                       | Use `Compression.Workers.deflateBatch`. Removed in 1.0.                                                                                                                                                                                                 |
| `inflateBatchWithPool`                       | Use `Compression.Workers.inflateBatch`. Removed in 1.0.                                                                                                                                                                                                 |
| `TaskOptions`                                | Use `WorkerTaskOptions`. Removed in 1.0.                                                                                                                                                                                                                |
| `TaskResult`                                 | Use `WorkerTaskResult`. Removed in 1.0.                                                                                                                                                                                                                 |
| `toNodeReadable`                             | Use `Archive.toNodeReadable`. Removed in 1.0.                                                                                                                                                                                                           |
| `TarGzArchive`                               | Use `Archive.TarGzArchive` (or build one with `Archive.createTarGz`). Removed in 1.0.                                                                                                                                                                   |
| `TarGzArchive#bytes`                         | Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                                                         |
| `TarGzArchive#bytesSync`                     | Use `toBufferSync`. Removed in 1.0.                                                                                                                                                                                                                     |
| `targz`                                      | Use `Archive.packTarGz`. Removed in 1.0.                                                                                                                                                                                                                |
| `parseTarGz`                                 | Use `Archive.parseTarGz`. Removed in 1.0.                                                                                                                                                                                                               |
| `parseTarGzStream`                           | Use `Archive.parseTarGzStream`, which decompresses and parses together in bounded memory; this one decompresses the whole archive before the first entry. For that all-or-nothing behaviour iterate `Archive.parseTarGz`. Removed in 1.0.               |
| `untargz`                                    | Use `Archive.unpackTarGz`. Removed in 1.0.                                                                                                                                                                                                              |
| `ParseTarGzOptions.maxFileSize?`             | Use `maxEntrySize`. Removed in 1.0.                                                                                                                                                                                                                     |
| `Archive.createAbortError`                   | Use `new AbortError(reason)` (a flat export of `documonster/archive` and `documonster/stream`); this one returns `reason` itself when it already is an `AbortError`. Removed in 1.0.                                                                    |
| `Archive.isAbortError`                       | Use the flat `isAbortError` export. Removed in 1.0.                                                                                                                                                                                                     |
| `Archive.zip`                                | Use `createZip`, or `createTar` for `{ format: "tar" }`. `zip` named a builder while `tar` named bytes. Removed in 1.0.                                                                                                                                 |
| `Archive.unzip`                              | Use `openZip`, or `openTar` for `{ format: "tar" }`. `unzip` named a reader while `untar` named a map. Removed in 1.0.                                                                                                                                  |
| `Archive.ZipArchive#bytes`                   | Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                                                         |
| `Archive.ZipArchive#bytesSync`               | Use `toBufferSync`. Removed in 1.0.                                                                                                                                                                                                                     |
| `Archive.ZipEditor#getEntries`               | Use `getSourceEntries`. The name read as the archive the edits produce — which `ArchiveFile#getEntries` returns — while this has always returned the source's entries. Removed in 1.0.                                                                  |
| `Archive.ZipEditor#bytes`                    | Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                                                         |
| `Archive.editZip`                            | Use `ZipEditor.open`; this only forwards to it. Removed in 1.0.                                                                                                                                                                                         |
| `Archive.editZipUrl`                         | Use `ZipEditor.openUrl`; this only forwards to it. Removed in 1.0.                                                                                                                                                                                      |
| `Archive.ZipEditPlan#toJSON`                 | Use `toJson`. Being named `toJSON`, this is also what `JSON.stringify(plan)` calls — and since it returns a Promise, that call serialises `{}`. Removed in 1.0.                                                                                         |
| `Archive.ZipEditPlan.fromJSON`               | Use `ZipEditPlan.fromJson`. Removed in 1.0.                                                                                                                                                                                                             |
| `Archive.UnzipEntry`                         | Entries come from a reader — `ZipReader.entries()` / `get()`, or iterating `Archive.openZip(…)` — and the shapes this takes are the reader's own. Removed in 1.0.                                                                                       |
| `Archive.TAR_TYPE`                           | Use `TarType`. Removed in 1.0.                                                                                                                                                                                                                          |
| `Archive.TarArchive#bytes`                   | Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                                                         |
| `Archive.TarArchive#bytesSync`               | Use `toBufferSync`. Removed in 1.0.                                                                                                                                                                                                                     |
| `Archive.createTarArchive`                   | Use `createTar`. Removed in 1.0.                                                                                                                                                                                                                        |
| `Archive.createTarReader`                    | Use `openTar`. Removed in 1.0.                                                                                                                                                                                                                          |
| `Archive.tar`                                | Use `packTar`. Removed in 1.0.                                                                                                                                                                                                                          |
| `Archive.tarSync`                            | Use `packTarSync`. Removed in 1.0.                                                                                                                                                                                                                      |
| `Archive.untar`                              | Use `unpackTar`. Removed in 1.0.                                                                                                                                                                                                                        |
| `Archive.base64ToUint8Array`                 | Not an archive operation. Decode natively: `Uint8Array.fromBase64(s)` where the runtime has it (not Node 22), `Buffer.from(s, "base64")` on Node, or `Uint8Array.from(atob(s), c => c.charCodeAt(0))` in a browser. Removed in 1.0.                     |
| `Archive.uint8ArrayToBase64`                 | Not an archive operation. Encode natively: `bytes.toBase64()` where the runtime has it (not Node 22), `Buffer.from(bytes).toString("base64")` on Node, or `btoa(Array.from(bytes, b => String.fromCharCode(b)).join(""))` in a browser. Removed in 1.0. |
| `Archive.concatUint8Arrays`                  | Not an archive operation. Natively: `const out = new Uint8Array(total)`, then `out.set(part, offset)` for each part (or `Buffer.concat(parts)` on Node). Removed in 1.0.                                                                                |
| `Archive.stringToUint8Array`                 | Not an archive operation. Use `new TextEncoder().encode(str)`. Removed in 1.0.                                                                                                                                                                          |
| `Archive.uint8ArrayToString`                 | Not an archive operation. Use `new TextDecoder(encoding).decode(bytes)`. Removed in 1.0.                                                                                                                                                                |
| `Archive.ZIP_CRYPTO_HEADER_SIZE`             | Use `Crypto.ZIP_CRYPTO_HEADER_SIZE`. Removed in 1.0.                                                                                                                                                                                                    |
| `Archive.zipCryptoInitKeys`                  | Use `Crypto.createZipCryptoKeys`. Removed in 1.0.                                                                                                                                                                                                       |
| `Archive.zipCryptoDecrypt`                   | Use `Crypto.decryptZipCrypto`, which throws an `ArchiveError` where this returns `null`. Removed in 1.0.                                                                                                                                                |
| `Archive.zipCryptoEncrypt`                   | Use `Crypto.encryptZipCrypto`. Removed in 1.0.                                                                                                                                                                                                          |
| `Archive.AES_VENDOR_ID`                      | Use `Crypto.AES_VENDOR_ID`. Removed in 1.0.                                                                                                                                                                                                             |
| `Archive.AES_VERSION_AE1`                    | Use `Crypto.AES_VERSION_AE1`. Removed in 1.0.                                                                                                                                                                                                           |
| `Archive.AES_VERSION_AE2`                    | Use `Crypto.AES_VERSION_AE2`. Removed in 1.0.                                                                                                                                                                                                           |
| `Archive.AES_EXTRA_FIELD_ID`                 | Use `Crypto.AES_EXTRA_FIELD_ID`. Removed in 1.0.                                                                                                                                                                                                        |
| `Archive.AES_SALT_LENGTH`                    | Use `Crypto.AES_SALT_LENGTH`. Removed in 1.0.                                                                                                                                                                                                           |
| `Archive.AES_KEY_LENGTH`                     | Use `Crypto.AES_KEY_LENGTH`. Removed in 1.0.                                                                                                                                                                                                            |
| `Archive.AES_AUTH_CODE_LENGTH`               | Use `Crypto.AES_AUTH_CODE_LENGTH`. Removed in 1.0.                                                                                                                                                                                                      |
| `Archive.AES_PASSWORD_VERIFY_LENGTH`         | Use `Crypto.AES_PASSWORD_VERIFY_LENGTH`. Removed in 1.0.                                                                                                                                                                                                |
| `Archive.COMPRESSION_METHOD_AES`             | Use `Crypto.COMPRESSION_METHOD_AES`. Removed in 1.0.                                                                                                                                                                                                    |
| `Archive.aesDecrypt`                         | Use `Crypto.decryptAes`. Removed in 1.0.                                                                                                                                                                                                                |
| `Archive.aesEncrypt`                         | Use `Crypto.encryptAes`. Removed in 1.0.                                                                                                                                                                                                                |
| `Archive.aesEncryptedSize`                   | Use `Crypto.getAesEncryptedSize`. Removed in 1.0.                                                                                                                                                                                                       |
| `Archive.buildAesExtraField`                 | Use `Crypto.createAesExtraField`. Removed in 1.0.                                                                                                                                                                                                       |
| `Archive.randomBytes`                        | Not an archive operation. Use the platform's generator: `crypto.getRandomValues(new Uint8Array(length))` (at most 65,536 bytes per call). Removed in 1.0.                                                                                               |
| `Archive.getEncryptionMethodName`            | Use `Crypto.getEncryptionMethodName`, which refuses a method it has no name for. This name keeps answering such a method with `undefined`, as it always did. Removed in 1.0.                                                                            |
| `Archive.isAesEncryption`                    | Use `Crypto.isAesEncryption`. Removed in 1.0.                                                                                                                                                                                                           |
| `Archive.getAesKeyStrength`                  | Use `Crypto.getAesKeyStrength`. Removed in 1.0.                                                                                                                                                                                                         |
| `Archive.encryptionMethodFromAesKeyStrength` | Use `Crypto.getEncryptionMethodForAesKeyStrength`. Removed in 1.0.                                                                                                                                                                                      |
| `ExtractOptions`                             | Use `RemoteZipExtractOptions`. Removed in 1.0.                                                                                                                                                                                                          |
| `ExtractOptions.onprogress?`                 | Use `onProgress`. Removed in 1.0.                                                                                                                                                                                                                       |
| `EOCDInfo`                                   | No public signature takes or returns it: the end-of-central-directory record is read inside `RemoteZipReader` and never handed out. Removed in 1.0, with no replacement.                                                                                |
| `ZIP64EOCDInfo`                              | No public signature takes or returns it: the ZIP64 end-of-central-directory record is read inside `RemoteZipReader` and never handed out. Removed in 1.0, with no replacement.                                                                          |
| `TarEntry.data`                              | Use `TarEntry.bytes`. Removed in 1.0.                                                                                                                                                                                                                   |
| `TarEntry.skip`                              | Use `TarEntry.discard`. Removed in 1.0.                                                                                                                                                                                                                 |
| `TarParseOptions.maxFileSize?`               | Use `maxEntrySize`. Removed in 1.0.                                                                                                                                                                                                                     |
| `TarReaderOptions.maxFileSize?`              | Use `maxEntrySize`. Removed in 1.0.                                                                                                                                                                                                                     |
| `encodePng`                                  | Use `Archive.encodePng`. Removed in 1.0.                                                                                                                                                                                                                |
| `InvalidZipSignatureError`                   | Use `ZipInvalidSignatureError`. Removed in 1.0.                                                                                                                                                                                                         |
| `EocdNotFoundError`                          | Use `ZipEocdNotFoundError`. Removed in 1.0.                                                                                                                                                                                                             |
| `Crc32MismatchError`                         | Use `ZipCrc32MismatchError`. Removed in 1.0.                                                                                                                                                                                                            |
| `DecryptionError`                            | Use `ZipDecryptionError`. Removed in 1.0.                                                                                                                                                                                                               |
| `PasswordRequiredError`                      | Use `ZipPasswordRequiredError`. Removed in 1.0.                                                                                                                                                                                                         |
| `RangeNotSupportedError`                     | Use `ArchiveRangeNotSupportedError`. Removed in 1.0.                                                                                                                                                                                                    |
| `HttpRangeError`                             | Use `ArchiveHttpRangeError`. Removed in 1.0.                                                                                                                                                                                                            |
| `FileTooLargeError`                          | Use `ArchiveFileTooLargeError`. Removed in 1.0.                                                                                                                                                                                                         |
| `UnsupportedCompressionError`                | Use `ZipUnsupportedCompressionError`. Removed in 1.0.                                                                                                                                                                                                   |
| `EntrySizeMismatchError`                     | Use `ZipEntrySizeMismatchError`. Removed in 1.0.                                                                                                                                                                                                        |
| `ArchiveLimitError`                          | Use `ArchiveLimitExceededError` — the same class, under the name every module's limit error has (`ExcelLimitExceededError`, `DocxLimitExceededError`, `PdfLimitExceededError`). Removed in 1.0.                                                         |
| `EntrySizeMismatchReason`                    | Use `ZipEntrySizeMismatchReason`. Removed in 1.0.                                                                                                                                                                                                       |
| `ArchiveFile`                                | Use `Archive.ArchiveFile`. Removed in 1.0.                                                                                                                                                                                                              |
| `ArchiveFile#getEntryCountAsync`             | Use `entryCount`, which now counts the entries of an opened TAR too. Removed in 1.0.                                                                                                                                                                    |
| `ArchiveFile#symlink`                        | Use `addSymlink`, which takes the mode as `{ mode }`. Removed in 1.0.                                                                                                                                                                                   |
| `ArchiveFile#bytes`                          | Use `toBuffer`. Removed in 1.0.                                                                                                                                                                                                                         |
| `ArchiveFile#bytesSync`                      | Use `toBufferSync`. Removed in 1.0.                                                                                                                                                                                                                     |
| `ArchiveFile#writeToFile`                    | Use `writeFile`. Removed in 1.0.                                                                                                                                                                                                                        |
| `ArchiveFile#writeToFileSync`                | Use `writeFileSync`. Removed in 1.0.                                                                                                                                                                                                                    |
| `ArchiveFile#addZipComment`                  | Use `setZipComment`: this sets the comment, it does not add one. Removed in 1.0.                                                                                                                                                                        |
| `OpenZipOptions`                             | Use `ArchiveFileOpenOptionsZip`. Removed in 1.0.                                                                                                                                                                                                        |
| `OpenTarOptions`                             | Use `ArchiveFileOpenOptionsTar`. Removed in 1.0.                                                                                                                                                                                                        |
| `OpenArchiveOptions`                         | Use `ArchiveFileOpenOptions`. Removed in 1.0.                                                                                                                                                                                                           |
| `OpenArchiveOptionsZip`                      | Use `ArchiveFileOpenOptionsZip`. Removed in 1.0.                                                                                                                                                                                                        |
| `OpenArchiveOptionsTar`                      | Use `ArchiveFileOpenOptionsTar`. Removed in 1.0.                                                                                                                                                                                                        |
| `TransformFunction`                          | Use `ArchiveTransformFunction`. Removed in 1.0.                                                                                                                                                                                                         |
| `TransformEntryData`                         | Use `ArchiveTransformEntryData`. Removed in 1.0.                                                                                                                                                                                                        |
| `TransformResult`                            | Use `ArchiveTransformResult`. Removed in 1.0.                                                                                                                                                                                                           |
| `traverseDirectory`                          | Use `Fs.traverseDirectory`. Removed in 1.0.                                                                                                                                                                                                             |
| `traverseDirectorySync`                      | Use `Fs.traverseDirectorySync`. Removed in 1.0.                                                                                                                                                                                                         |
| `glob`                                       | Use `Fs.glob`. Removed in 1.0.                                                                                                                                                                                                                          |
| `globSync`                                   | Use `Fs.globSync`. Removed in 1.0.                                                                                                                                                                                                                      |
| `globToRegex`                                | Use `Fs.globToRegex`. Removed in 1.0.                                                                                                                                                                                                                   |
| `matchGlob`                                  | Use `Fs.matchGlob`. Removed in 1.0.                                                                                                                                                                                                                     |
| `matchGlobAny`                               | Use `Fs.matchGlobAny`. Removed in 1.0.                                                                                                                                                                                                                  |
| `ensureDir`                                  | Use `Fs.ensureDir`. Removed in 1.0.                                                                                                                                                                                                                     |
| `ensureDirSync`                              | Use `Fs.ensureDirSync`. Removed in 1.0.                                                                                                                                                                                                                 |
| `fileExists`                                 | Use `Fs.fileExists`. Removed in 1.0.                                                                                                                                                                                                                    |
| `fileExistsSync`                             | Use `Fs.fileExistsSync`. Removed in 1.0.                                                                                                                                                                                                                |
| `readFileBytes`                              | Use `Fs.readFileBytes`. Removed in 1.0.                                                                                                                                                                                                                 |
| `readFileBytesSync`                          | Use `Fs.readFileBytesSync`. Removed in 1.0.                                                                                                                                                                                                             |
| `writeFileBytes`                             | Use `Fs.writeFileBytes`. Removed in 1.0.                                                                                                                                                                                                                |
| `writeFileBytesSync`                         | Use `Fs.writeFileBytesSync`. Removed in 1.0.                                                                                                                                                                                                            |
| `setFileTime`                                | Use `Fs.setFileTime`. Removed in 1.0.                                                                                                                                                                                                                   |
| `setFileTimeSync`                            | Use `Fs.setFileTimeSync`. Removed in 1.0.                                                                                                                                                                                                               |
| `safeStats`                                  | Use `Fs.safeStats`, which answers `undefined` for a missing file. This keeps the `null` it always answered. Removed in 1.0.                                                                                                                             |
| `safeStatsSync`                              | Use `Fs.safeStatsSync`, which answers `undefined` for a missing file. This keeps the `null` it always answered. Removed in 1.0.                                                                                                                         |
| `readFileText`                               | Use `Fs.readFileText`. Removed in 1.0.                                                                                                                                                                                                                  |
| `readFileTextSync`                           | Use `Fs.readFileTextSync`. Removed in 1.0.                                                                                                                                                                                                              |
| `writeFileText`                              | Use `Fs.writeFileText`. Removed in 1.0.                                                                                                                                                                                                                 |
| `writeFileTextSync`                          | Use `Fs.writeFileTextSync`. Removed in 1.0.                                                                                                                                                                                                             |
| `remove`                                     | Use `Fs.remove`. Removed in 1.0.                                                                                                                                                                                                                        |
| `removeSync`                                 | Use `Fs.removeSync`. Removed in 1.0.                                                                                                                                                                                                                    |
| `copyFile`                                   | Use `Fs.copyFile`. Removed in 1.0.                                                                                                                                                                                                                      |
| `copyFileSync`                               | Use `Fs.copyFileSync`. Removed in 1.0.                                                                                                                                                                                                                  |
| `createReadStream`                           | Use createReadStream from node:fs; this only forwards to it. Removed in 1.0.                                                                                                                                                                            |
| `createWriteStream`                          | Use createWriteStream from node:fs; this only forwards to it. Removed in 1.0.                                                                                                                                                                           |
| `createTempDir`                              | Use `Fs.createTempDir`. Removed in 1.0.                                                                                                                                                                                                                 |
| `createTempDirSync`                          | Use `Fs.createTempDirSync`. Removed in 1.0.                                                                                                                                                                                                             |
| `FileEntry`                                  | Use `FsFileEntry`. Removed in 1.0.                                                                                                                                                                                                                      |
| `TraverseOptions`                            | Use `FsTraverseOptions`. Removed in 1.0.                                                                                                                                                                                                                |
| `GlobOptions`                                | Use `FsGlobOptions`. Removed in 1.0.                                                                                                                                                                                                                    |
| `ReadStreamOptions`                          | Use `ReadStreamOptions` from `node:fs`, as `createReadStream` from `node:fs` replaces this module's. Removed in 1.0.                                                                                                                                    |
| `WriteStreamOptions`                         | Use `WriteStreamOptions` from `node:fs`, as `createWriteStream` from `node:fs` replaces this module's. Removed in 1.0.                                                                                                                                  |

**`documonster/archive [browser]`**

| Deprecated           | Replacement                                                     |
| -------------------- | --------------------------------------------------------------- |
| `createGzipStream`   | Use `Compression.createGzipStream`. Removed in 1.0.             |
| `createGunzipStream` | Use `Compression.createGunzipStream`. Removed in 1.0.           |
| `createZlibStream`   | Use `Compression.createZlibStream`. Removed in 1.0.             |
| `createUnzlibStream` | Use `Compression.createUnzlibStream`. Removed in 1.0.           |
| `Archive.editZip`    | Use `ZipEditor.open`; this only forwards to it. Removed in 1.0. |
| `Archive.untar`      | Use `unpackTar`. Removed in 1.0.                                |

**`documonster/chart`**

| Deprecated                               | Replacement                                                                                                                                                                                                                                                                                                            |
| ---------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `buildChartModel`                        | Use `ChartClassic.createModel`. Removed in 1.0.                                                                                                                                                                                                                                                                        |
| `buildComboChartModel`                   | Use `ChartClassic.createComboModel`. Removed in 1.0.                                                                                                                                                                                                                                                                   |
| `fillChartCaches`                        | Use `ChartClassic.fillCaches(model, workbook, { worksheet })`. Removed in 1.0.                                                                                                                                                                                                                                         |
| `fillChartExCaches`                      | Use `ChartEx.fillCaches(model, workbook, { worksheet })`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `fillNumRef`                             | Use `ChartClassic.fillNumRef(ref, workbook, { date1904, worksheet })`. Removed in 1.0.                                                                                                                                                                                                                                 |
| `fillStrRef`                             | Use `ChartClassic.fillStrRef(ref, workbook, { worksheet })`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `parseChartColors`                       | Use `ChartClassic.fromColorsXml`. Removed in 1.0.                                                                                                                                                                                                                                                                      |
| `buildChartColors`                       | Use `ChartClassic.toColorsXml`. Removed in 1.0.                                                                                                                                                                                                                                                                        |
| `parseChartStyle`                        | Use `ChartClassic.fromStyleXml`. Removed in 1.0.                                                                                                                                                                                                                                                                       |
| `buildChartStyle`                        | Use `ChartClassic.toStyleXml`. Removed in 1.0.                                                                                                                                                                                                                                                                         |
| `buildChartExModel`                      | Use `ChartEx.createModel`. Removed in 1.0.                                                                                                                                                                                                                                                                             |
| `parseChartEx`                           | Use `ChartEx.fromXml`. Removed in 1.0.                                                                                                                                                                                                                                                                                 |
| `canRenderChartExAsVectorPdf`            | Every Office 2016 chart has a vector picture: `ChartEx.toDrawList(model)` draws any model, and `Pdf.fromChart` picks its path itself. This still says whether each series' layout is one of `VECTOR_PDF_CHART_EX_LAYOUT_IDS`. Removed in 1.0.                                                                          |
| `drawChartExPdf`                         | Put the chart on the page as a display list, as every producer does: `Draw.render(ChartEx.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`. This draws that same display list — the chart and its chart area — onto a surface with `drawRect` / `drawLine` / `drawText`. Removed in 1.0.      |
| `renderChartExPng`                       | Use `ChartEx.toPng`. Removed in 1.0.                                                                                                                                                                                                                                                                                   |
| `renderChartExSvg`                       | Use `ChartEx.toSvg`. Removed in 1.0.                                                                                                                                                                                                                                                                                   |
| `VECTOR_PDF_CHART_EX_LAYOUT_IDS`         | It only answers the deprecated `canRenderChartExAsVectorPdf`: every Office 2016 chart now has a vector picture through `ChartEx.toDrawList`. Removed in 1.0.                                                                                                                                                           |
| `renderChartEx`                          | Use `ChartEx.toXml` — this returns the part's XML, not a picture. Removed in 1.0.                                                                                                                                                                                                                                      |
| `applyChartExPreset`                     | Use `ChartEx.applyPreset`. Removed in 1.0.                                                                                                                                                                                                                                                                             |
| `applyChartPreset`                       | Use `ChartClassic.applyPreset`. Removed in 1.0.                                                                                                                                                                                                                                                                        |
| `EXCEL_CHART_EX_PRESETS`                 | The preset names are the keys of the table: `Object.keys(CHART_EX_PRESETS)`. Removed in 1.0.                                                                                                                                                                                                                           |
| `EXCEL_CHART_PRESETS`                    | The preset names are the keys of the table: `Object.keys(CHART_PRESETS)`. Removed in 1.0.                                                                                                                                                                                                                              |
| `chartExOptionsFromRows`                 | Use `ChartEx.createOptionsFromRows`. Removed in 1.0.                                                                                                                                                                                                                                                                   |
| `chartExOptionsFromTable`                | Use `ChartEx.createOptionsFromTable`. Removed in 1.0.                                                                                                                                                                                                                                                                  |
| `chartOptionsFromRows`                   | Use `ChartClassic.createOptionsFromRows`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `chartOptionsFromTable`                  | Use `ChartClassic.createOptionsFromTable`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `seriesFromColumns`                      | Use `ChartClassic.createSeriesFromColumns`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `applyAxisTransform`                     | Use `ChartClassic.Backend.applyAxisTransform`. Removed in 1.0.                                                                                                                                                                                                                                                         |
| `buildChartScene`                        | Use `ChartClassic.Backend.toScene`. Removed in 1.0.                                                                                                                                                                                                                                                                    |
| `buildEffectFilter`                      | Use `ChartClassic.Backend.toSvgFilter`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `drawChartPdf`                           | Put the chart on the page as a display list, as every producer does: `Draw.render(ChartClassic.toDrawList(model, { width, height }), Pdf.createDrawSurface(page, rect))`. This draws that same display list — the chart and its chart area — onto a surface with `drawRect` / `drawLine` / `drawText`. Removed in 1.0. |
| `renderChartPng`                         | Use `ChartClassic.toPng`. Removed in 1.0.                                                                                                                                                                                                                                                                              |
| `renderChartSvg`                         | Use `ChartClassic.toSvg`. Removed in 1.0.                                                                                                                                                                                                                                                                              |
| `ChartPdfDrawingSurface`                 | A chart goes onto a PDF page as a display list, through `Pdf.createDrawSurface` (a `PdfDrawPage`). Removed in 1.0.                                                                                                                                                                                                     |
| `ChartPdfPathOp`                         | A path operation of `ChartPdfDrawingSurface`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `ChartRenderOptions.backgroundColor?`    | Use `background`, which also takes `{ r, g, b, a }`. When both are given, `background` wins. Removed in 1.0.                                                                                                                                                                                                           |
| `PdfChartRenderOptions`                  | Pass `ChartRenderOptions` to `ChartClassic.toDrawList` and the box to `Pdf.createDrawSurface`. Removed in 1.0.                                                                                                                                                                                                         |
| `PdfChartRenderOptions.backgroundColor?` | Use `background`, which also takes `{ r, g, b, a }`. When both are given, `background` wins. Removed in 1.0.                                                                                                                                                                                                           |
| `PdfColor`                               | See `ChartPdfColor`. Removed in 1.0.                                                                                                                                                                                                                                                                                   |
| `resolveTopologyObject`                  | Use `ChartClassic.Backend.resolveTopologyObject`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `parseSpPr`                              | Use `DrawingMl.parseSpPr`. Removed in 1.0.                                                                                                                                                                                                                                                                             |
| `parseTxPr`                              | Use `DrawingMl.parseTxPr`. Removed in 1.0.                                                                                                                                                                                                                                                                             |
| `getSpPrFillColor`                       | Use `DrawingMl.getSpPrFillColor`. Removed in 1.0.                                                                                                                                                                                                                                                                      |
| `getSpPrLine`                            | Use `DrawingMl.getSpPrLine`. Removed in 1.0.                                                                                                                                                                                                                                                                           |
| `getSpPrGradient`                        | Use `DrawingMl.getSpPrGradient`. Removed in 1.0.                                                                                                                                                                                                                                                                       |
| `getSpPrPattern`                         | Use `DrawingMl.getSpPrPattern`. Removed in 1.0.                                                                                                                                                                                                                                                                        |
| `getTxPrFontSize`                        | Use `DrawingMl.getTxPrFontSize`. Removed in 1.0.                                                                                                                                                                                                                                                                       |
| `getTxPrColor`                           | Use `DrawingMl.getTxPrColor`. Removed in 1.0.                                                                                                                                                                                                                                                                          |
| `buildSpPr`                              | Use `DrawingMl.createSpPr`. Removed in 1.0.                                                                                                                                                                                                                                                                            |
| `buildTxPr`                              | Use `DrawingMl.createTxPr`. Removed in 1.0.                                                                                                                                                                                                                                                                            |
| `setSpPrFill`                            | Use `DrawingMl.withSpPrFill`, which — as this always did — returns a copy rather than editing `spPr`. Removed in 1.0.                                                                                                                                                                                                  |
| `setSpPrLine`                            | Use `DrawingMl.withSpPrLine`, which — as this always did — returns a copy rather than editing `spPr`. Removed in 1.0.                                                                                                                                                                                                  |

**`documonster/csv`**

| Deprecated                                    | Replacement                                                                                                                                                                                                                                                                                     |
| --------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `HeaderArray`                                 | Use `CsvHeaderArray`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `RowHashArray`                                | Use `CsvRowHashArray`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `RowArray`                                    | Use `CsvRowArray`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `RowMap`                                      | Use `CsvRowMap`. Removed in 1.0.                                                                                                                                                                                                                                                                |
| `Row`                                         | Use `CsvRow`. Removed in 1.0.                                                                                                                                                                                                                                                                   |
| `HeaderTransformFunction`                     | Use `CsvHeaderTransformFunction`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `RowTransformCallback`                        | Use `CsvRowTransformCallback`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `RowTransformFunction`                        | Use `CsvRowTransformFunction`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `RowValidateCallback`                         | Use `CsvRowValidateCallback`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `RowValidateFunction`                         | Use `CsvRowValidateFunction`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `TransformContext`                            | Use `CsvTransformContext`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `FormattedValue`                              | Use `CsvFormattedValue`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `TransformResult`                             | Use `CsvTransformResult`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `TypeTransformMap`                            | Use `CsvTypeTransformMap`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `DynamicTypingConfig`                         | Use `CsvDynamicTypingConfig`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `CastDateConfig`                              | Use `CsvCastDateConfig`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `ColumnConfig`                                | Use `CsvColumnConfig`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `OnSkipCallback`                              | Use `CsvOnSkipCallback`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `CsvParseOptions.skipRecordsWithError?`       | Has no effect: a malformed record (a column-count mismatch the `columnMismatch` strategy rejects, or text after a closing quote) is always left out of `rows`, reported in `errors` / `invalidRows` and handed to `onSkip`. It only used to decide whether `onSkip` was called. Removed in 1.0. |
| `CsvParseArrayOptions.skipRecordsWithError?`  | Has no effect: a malformed record (a column-count mismatch the `columnMismatch` strategy rejects, or text after a closing quote) is always left out of `rows`, reported in `errors` / `invalidRows` and handed to `onSkip`. It only used to decide whether `onSkip` was called. Removed in 1.0. |
| `CsvParseObjectOptions.skipRecordsWithError?` | Has no effect: a malformed record (a column-count mismatch the `columnMismatch` strategy rejects, or text after a closing quote) is always left out of `rows`, reported in `errors` / `invalidRows` and handed to `onSkip`. It only used to decide whether `onSkip` was called. Removed in 1.0. |
| `ChunkMeta`                                   | Use `CsvChunkMeta`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `RecordInfo`                                  | Use `CsvRecordInfo`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `RecordInfo.invalid_field_length?`            | Never set by any parser. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `RecordWithInfo`                              | Use `CsvRecordWithInfo`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `ColumnMismatchConfig`                        | Use `CsvColumnMismatchConfig`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `ColumnMismatchLess`                          | Use `CsvColumnMismatchLess`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `ColumnMismatchMore`                          | Use `CsvColumnMismatchMore`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Csv.quoted`                                  | Use `forceQuote`. Removed in 1.0.                                                                                                                                                                                                                                                               |
| `Csv.unquoted`                                | Use `preventQuote`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `Csv.parseRows`                               | Use `parseStream`. An incremental reader takes the `Stream` suffix, as `Xml.parseSaxStream` and `Workbook.readStream` do. Removed in 1.0.                                                                                                                                                       |
| `Csv.parseWithProgress`                       | Use `parseAsync` with `{ onProgress }`. Removed in 1.0.                                                                                                                                                                                                                                         |
| `Csv.createParserStream`                      | Use `new Csv.ParserStream(options)`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `Csv.createFormatterStream`                   | Use `new Csv.FormatterStream(options)`. Removed in 1.0.                                                                                                                                                                                                                                         |
| `Csv.rowHashArrayToValues`                    | Use `getRowValues`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `Csv.rowHashArrayToHeaders`                   | Use `getRowHeaders`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `Csv.rowHashArrayMapByHeaders`                | Use `getRowValuesByHeaders`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Csv.generateAsync`                           | Use `generateRowsAsync`. `Async` marks the asynchronous sibling of a synchronous form, so this read as the sibling of `generate`, which returns the whole CSV; what it does is yield rows, as `generateRows` does. Removed in 1.0.                                                              |
| `ColumnDef`                                   | Use `CsvColumnDef`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `GeneratorColumnConfig`                       | Use `CsvGeneratorColumnConfig`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `BuiltinColumnType`                           | Use `CsvBuiltinColumnType`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `GeneratorFn`                                 | Use `CsvGeneratorFn`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `GeneratorContext`                            | Use `CsvGeneratorContext`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `StopCondition`                               | Use `CsvStopCondition`. Removed in 1.0.                                                                                                                                                                                                                                                         |
| `StopContext`                                 | Use `CsvStopContext`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `DecimalSeparator`                            | Use `CsvDecimalSeparator`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `ParsedRow`                                   | Use `CsvParsedRow`. Removed in 1.0.                                                                                                                                                                                                                                                             |

**`documonster/draw`**

| Deprecated                              | Replacement                                                                                                                                                                                                                   |
| --------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cssColour`                             | Use `Draw.parseColor`. Removed in 1.0.                                                                                                                                                                                        |
| `relativeLuminance`                     | Use `Draw.getRelativeLuminance`. Removed in 1.0.                                                                                                                                                                              |
| `translucent`                           | Use `Draw.withAlpha`. Removed in 1.0.                                                                                                                                                                                         |
| `measureText`                           | Use `Draw.measureText`. Removed in 1.0.                                                                                                                                                                                       |
| `widestText`                            | Use `Draw.measureWidest`. Removed in 1.0.                                                                                                                                                                                     |
| `wrapText`                              | Use `Draw.wrapText`. Removed in 1.0.                                                                                                                                                                                          |
| `renderDrawList`                        | Use `Draw.render`. Removed in 1.0.                                                                                                                                                                                            |
| `renderNode`                            | Use `Draw.Backend.renderNode`. Removed in 1.0.                                                                                                                                                                                |
| `SvgSurface`                            | Use `Draw.Backend.SvgSurface`. Removed in 1.0.                                                                                                                                                                                |
| `SvgSurface#markup`                     | Use `format`. Removed in 1.0.                                                                                                                                                                                                 |
| `toSvg`                                 | Use `Draw.toSvg`, which checks the list and its options. This form keeps writing what it is given, as it always did. Removed in 1.0.                                                                                          |
| `ToSvgOptions`                          | Use `DrawToSvgOptions`. Removed in 1.0.                                                                                                                                                                                       |
| `createRasterSurface`                   | Use `Draw.Backend.createRasterSurface`, which takes `{ scale }`. Removed in 1.0.                                                                                                                                              |
| `downsample`                            | Use `Draw.Backend.downsample`. Removed in 1.0.                                                                                                                                                                                |
| `normalizeSamples`                      | Use `Draw.Backend.normalizeSamples`. Removed in 1.0.                                                                                                                                                                          |
| `rasterizeToRgba`                       | Use `Draw.toRgba`, which checks the list and its options and refuses a list with no area. This form keeps its old reading — a 0×0 list rasterises to a 1×1 transparent image. Removed in 1.0.                                 |
| `RasterizeOptions`                      | Use `DrawRasterizeOptions`. Removed in 1.0.                                                                                                                                                                                   |
| `RasterizedImage`                       | Use `DrawRasterizedImage`. Removed in 1.0.                                                                                                                                                                                    |
| `RgbaImage`                             | Use `DrawRgbaImage`. Removed in 1.0.                                                                                                                                                                                          |
| `isSimpleText`                          | Use `Draw.isSimpleText`. Removed in 1.0.                                                                                                                                                                                      |
| `textFeaturesOf`                        | Use `Draw.getTextFeatures`. Removed in 1.0.                                                                                                                                                                                   |
| `BasicRasterCanvas`                     | Use `Draw.Backend.BasicRasterCanvas`. Removed in 1.0.                                                                                                                                                                         |
| `BasicRasterCanvas#uncoveredCodePoints` | Use `getUncoveredCodePoints`. This handed out the canvas's own set, so a caller could clear it and change what the render reports; it now returns a copy. Removed in 1.0.                                                     |
| `BasicRasterCanvas#textWarnings`        | Use `getTextWarnings`. Removed in 1.0.                                                                                                                                                                                        |
| `RasterPoint`                           | Use `DrawRasterPoint`. Removed in 1.0.                                                                                                                                                                                        |
| `GlyphOutline`                          | Use `DrawGlyphOutline`. Removed in 1.0.                                                                                                                                                                                       |
| `RasterFont`                            | Use `DrawRasterFont`. Removed in 1.0.                                                                                                                                                                                         |
| `fontHasGlyph`                          | Use `Draw.Backend.fontHasGlyph`. Removed in 1.0.                                                                                                                                                                              |
| `parseRasterFont`                       | Use `Draw.parseRasterFont`, which takes the bytes or `{ data, collectionIndex }` and throws for bytes it cannot read (this form returns an empty face, and face 0 for an index past the end of a collection). Removed in 1.0. |
| `RasterFontSource`                      | Use `DrawRasterFontSource`. Removed in 1.0.                                                                                                                                                                                   |
| `IDENTITY`                              | Use `Matrix.IDENTITY`. Removed in 1.0.                                                                                                                                                                                        |
| `apply`                                 | Use `Matrix.apply`. Removed in 1.0.                                                                                                                                                                                           |
| `arcToCubics`                           | Use `Draw.Geometry.arcToCubics`. Removed in 1.0.                                                                                                                                                                              |
| `multiply`                              | Use `Matrix.multiply`. Removed in 1.0.                                                                                                                                                                                        |
| `rotate`                                | Use `Matrix.rotate`. Removed in 1.0.                                                                                                                                                                                          |
| `flattenPath`                           | Use `Draw.Geometry.flattenPath`, which takes `{ scale }` and checks it. This form multiplies by whatever `scale` it is given, as it always did. Removed in 1.0.                                                               |
| `rectNode`                              | Use `Draw.Geometry.createRectNode`. Removed in 1.0.                                                                                                                                                                           |
| `roundedRectToPath`                     | Use `Draw.Geometry.roundedRectToPath`. Removed in 1.0.                                                                                                                                                                        |
| `rotationOf`                            | Use `Matrix.getRotation`. Removed in 1.0.                                                                                                                                                                                     |
| `scale`                                 | Use `Matrix.scale`. Removed in 1.0.                                                                                                                                                                                           |
| `sectorToPath`                          | Use `Draw.Geometry.sectorToPath`. Removed in 1.0.                                                                                                                                                                             |
| `translate`                             | Use `Matrix.translate`. Removed in 1.0.                                                                                                                                                                                       |
| `uniformScale`                          | Use `Matrix.getUniformScale`. Removed in 1.0.                                                                                                                                                                                 |
| `Rgba01`                                | Use `DrawRgba01`. Removed in 1.0.                                                                                                                                                                                             |
| `RasterTextStyle`                       | Use `DrawRasterTextStyle`. Removed in 1.0.                                                                                                                                                                                    |

**`documonster/excel`**

| Deprecated                                           | Replacement                                                                                                                                                                                                                                                                                                                                                                                       |
| ---------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Workbook.getWorksheets`                             | Use `listWorksheets`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Workbook.getChartsheets`                            | Use `listChartsheets`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                            |
| `Workbook.renameChartsheet`                          | Look the chartsheet up, then rename it: `const cs = Workbook.getChartsheet(wb, nameOrIndex); if (cs) Chartsheet.setName(cs, name);` — this answered `false` when there is no such chartsheet. Removed in 1.0.                                                                                                                                                                                     |
| `Workbook.eachSheet`                                 | Use `eachWorksheet`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Workbook.importSheet`                               | Use `importWorksheet`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                            |
| `Workbook.defineCellStyle`                           | Use `addCellStyle`, which returns (a copy of) the stored entry. Removed in 1.0.                                                                                                                                                                                                                                                                                                                   |
| `Workbook.useBuiltinCellStyle`                       | Use `useBuiltInCellStyle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                        |
| `Workbook.createStreamWriter`                        | Use `new Stream.WorkbookWriter(options)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                         |
| `Workbook.createStreamReader`                        | Use `new Stream.WorkbookReader(input, options)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                  |
| `Workbook.XlsxReadable`                              | Use the flat `XlsxReadable` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                 |
| `Workbook.XlsxWritable`                              | Use the flat `XlsxWritable` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                 |
| `Workbook.XlsxStreamOptions`                         | Use the flat `XlsxStreamOptions` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                            |
| `Workbook.Handle.pivotTables`                        | Use `Pivot.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                         |
| `Stream.WorkbookWriter#addMedia`                     | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes the media a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                   |
| `Stream.WorkbookWriter#xlsbUnsupported`              | Use `unsupported`. Removed in 1.0. What it always returned: the XLSB writer's losses only — empty for an XLSX stream — with a row's loss as `A1: …`.                                                                                                                                                                                                                                              |
| `Stream.WorkbookWriter#xlsbFormulaContext`           | The sheet writer's view of the workbook while it encodes rows; nothing outside the writer needs it. Removed in 1.0.                                                                                                                                                                                                                                                                               |
| `Stream.WorkbookWriter#date1904Flag`                 | The sheet writer's view of the workbook while it encodes rows; nothing outside the writer needs it. Removed in 1.0.                                                                                                                                                                                                                                                                               |
| `Stream.WorkbookWriter#definedNames`                 | Use `Workbook.getDefinedNames`: `Workbook.getDefinedNames(wb)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                   |
| `Stream.WorkbookWriter#defineCellStyle`              | Use `Workbook.addCellStyle`, which returns (a copy of) the stored entry. Removed in 1.0.                                                                                                                                                                                                                                                                                                          |
| `Stream.WorkbookWriter#packageSink`                  | The sheet writer's view of the workbook while it encodes rows; nothing outside the writer needs it. Removed in 1.0.                                                                                                                                                                                                                                                                               |
| `Stream.WorkbookWriter#nextId`                       | The id `addWorksheet` gives the next sheet; nothing outside the writer needs it. Removed in 1.0.                                                                                                                                                                                                                                                                                                  |
| `Stream.WorkbookWriter#addImage`                     | Use `Image.add`: `Image.add(wb, image)`, then `Image.place(ws, id, range)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                       |
| `Stream.WorkbookWriter#protect`                      | Use `Workbook.protect`: `Workbook.protect(wb, { password, ...flags })`. Removed in 1.0.                                                                                                                                                                                                                                                                                                           |
| `Stream.WorkbookWriter#unprotect`                    | Use `Workbook.unprotect`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                         |
| `Stream.WorkbookWriter#addWorksheet`                 | Use `Workbook.addWorksheet`: `Workbook.addWorksheet(wb, name, options)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                          |
| `Stream.WorkbookWriter#getWorksheet`                 | Use `Workbook.getWorksheet`: `Workbook.getWorksheet(wb, id)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                     |
| `Stream.WorkbookWriter#addStyles`                    | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addThemes`                    | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addOfficeRels`                | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addContentTypes`              | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addApp`                       | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addCore`                      | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addSharedStrings`             | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addFeaturePropertyBag`        | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addMetadata`                  | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addWorkbookRels`              | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.WorkbookWriter#addWorkbook`                  | One of the steps `commit` runs to write the package, and only correct there: called on its own it writes its part a second time. Use `commit`. Removed in 1.0.                                                                                                                                                                                                                                    |
| `Stream.setCellComment`                              | A comment is the cell's note with its author: `Stream.setCellNote(c, note, { author })`, and `removeCellNote` for `undefined`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Stream.rowValues`                                   | Use `getRowValues` with `{ base: 1 }` — this reads the row 1-based, by column number. Removed in 1.0.                                                                                                                                                                                                                                                                                             |
| `Stream.rowCell`                                     | Use `getRowCell`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                 |
| `Stream.WorksheetWriter#xlsbUnsupported`             | Use `unsupported`, which reports both formats and names the sheet. Removed in 1.0.                                                                                                                                                                                                                                                                                                                |
| `Stream.WorksheetWriter#stream`                      | The sheet part's output while the writer encodes it; writing to it directly corrupts the part, and nothing outside the writer needs it. Removed in 1.0.                                                                                                                                                                                                                                           |
| `Stream.WorksheetWriter#mergeCells`                  | Use `Worksheet.merge`: `Worksheet.merge(ws, "A1:B2")`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                            |
| `Stream.WorksheetWriter#addConditionalFormatting`    | Use `ConditionalFormatting.add`: `ConditionalFormatting.add(ws, cf)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                             |
| `Stream.WorksheetWriter#removeConditionalFormatting` | Use `ConditionalFormatting.remove` (one block) or `ConditionalFormatting.clear` (all). Removed in 1.0.                                                                                                                                                                                                                                                                                            |
| `Stream.WorksheetWriter#addBackgroundImage`          | Use `Image.setBackground`: `Image.setBackground(ws, imageId)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                    |
| `Stream.WorksheetWriter#getBackgroundImageId`        | Use `Image.getBackground`, which answers with the id as a string (`"0"`), as it does for a document-model sheet. Removed in 1.0.                                                                                                                                                                                                                                                                  |
| `Stream.WorksheetWriter#addImage`                    | Use `Image.place`: `Image.place(ws, imageId, range)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                             |
| `Stream.WorksheetWriter#getImages`                   | Use `Image.list`, which lists the placed pictures without the entry an overlay watermark adds. Removed in 1.0.                                                                                                                                                                                                                                                                                    |
| `Stream.WorksheetWriter#addWatermark`                | Use `Watermark.set`: `Watermark.set(ws, options)`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                |
| `Stream.WorksheetWriter#getWatermark`                | Use `Watermark.get`, which answers with a copy. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                   |
| `Stream.WorksheetWriter#removeWatermark`             | Use `Watermark.remove`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Stream.WorksheetWriter#protect`                     | Use `Worksheet.protect`: `Worksheet.protect(ws, { password, ...flags })`. Removed in 1.0.                                                                                                                                                                                                                                                                                                         |
| `Stream.WorksheetWriter#unprotect`                   | Use `Worksheet.unprotect`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                        |
| `Stream.WorksheetWriter#drawing`                     | The workbook writer's view of the sheet while it writes the package; nothing outside the writer needs it. `getImages()` lists what was added. Removed in 1.0.                                                                                                                                                                                                                                     |
| `Worksheet.spliceRows`                               | Use `Row.splice`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                 |
| `Worksheet.spliceColumns`                            | Use `Column.splice`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Worksheet.insertRow`                                | Use `Row.insert`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                 |
| `Worksheet.insertRows`                               | Use `Row.insertMany`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Worksheet.duplicateRow`                             | Use `Row.duplicate`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Worksheet.destroy`                                  | Use `Workbook.removeWorksheet(wb, ws)`, which this does. Removed in 1.0.                                                                                                                                                                                                                                                                                                                          |
| `Worksheet.autoFitColumn`                            | Use `Column.autoFit`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Worksheet.autoFitColumns`                           | Use `Column.autoFit`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Worksheet.autoFitRow`                               | Use `Row.autoFit`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                |
| `Worksheet.autoFitRows`                              | Use `Row.autoFit`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                |
| `Worksheet.addConditionalFormatting`                 | Use `ConditionalFormatting.add`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                  |
| `Worksheet.removeConditionalFormatting`              | Use `ConditionalFormatting.remove` for an entry, a position or a predicate, and `ConditionalFormatting.clear` for every entry (what calling this without a filter does). Removed in 1.0.                                                                                                                                                                                                          |
| `Worksheet.dimensions`                               | Use `getDimensions`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Worksheet.columnCount`                              | Use `Column.getCount`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                            |
| `Worksheet.actualColumnCount`                        | Use `Column.getActualCount`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                      |
| `Worksheet.rowCount`                                 | Use `Row.getCount`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                               |
| `Worksheet.actualRowCount`                           | Use `Row.getActualCount`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                         |
| `Worksheet.mergedRegions`                            | Use `getMergedRegions`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Worksheet.setColumns`                               | Use `Column.setDefinitions`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                      |
| `Worksheet.columns`                                  | Use `Column.getDefinitions`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                      |
| `Worksheet.columnDefinitions`                        | Use `Column.getDefinitions`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                      |
| `Worksheet.lastColumn`                               | Use `Column.getLast`, which returns the column **number** (or `undefined` for an empty sheet) rather than a column view, and creates no column record. Removed in 1.0.                                                                                                                                                                                                                            |
| `Worksheet.lastRow`                                  | Use `Row.getLast`, which returns the row **number** rather than the row record. Removed in 1.0.                                                                                                                                                                                                                                                                                                   |
| `Worksheet.panes`                                    | Use `getPanes`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Worksheet.autoFilter`                               | Use `getAutoFilter`, which answers `undefined` rather than `null`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                |
| `Worksheet.addRow`                                   | Use `Row.add`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Worksheet.addRows`                                  | Use `Row.addMany`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                |
| `Worksheet.getRow`                                   | Use the `Row` and `Cell` namespaces. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                              |
| `Worksheet.getRows`                                  | Use the `Row` and `Cell` namespaces. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                              |
| `Worksheet.findRow`                                  | Use `Row.has`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Worksheet.findRows`                                 | Use `Row.has`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Worksheet.eachRow`                                  | Use `Row.each`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Worksheet.Handle.autoFilter`                        | Use `Worksheet.getAutoFilter` / `Worksheet.setAutoFilter`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                |
| `Worksheet.Handle.tables`                            | Use `Table.list` or `Table.get`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                          |
| `Worksheet.Handle.pivotTables`                       | Use `Pivot.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                         |
| `Worksheet.Handle.conditionalFormattings`            | Use `ConditionalFormatting.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                         |
| `Worksheet.Handle.formControls`                      | Use `Form.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                          |
| `Cell.getMergeMaster`                                | Use `getMergeMasterAddress`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                      |
| `Cell.setName`                                       | Use `setNames` with a one-name array: `setNames(ws, addr, [name])`. Removed in 1.0.                                                                                                                                                                                                                                                                                                               |
| `Cell.getComment`                                    | A comment is the cell's note with its author: read them with `getNote` and `getNoteAuthor`. Removed in 1.0.                                                                                                                                                                                                                                                                                       |
| `Cell.setComment`                                    | A comment is the cell's note with its author: `Cell.setNote(ws, addr, note, { author })`, and `removeNote` for `undefined`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `Cell.find`                                          | Use `has` to test for a cell, `getView` to read one without creating it, and address it by position. Removed in 1.0.                                                                                                                                                                                                                                                                              |
| `Cell.view`                                          | The document model is addressed by position; use `getView` (`Cell.getView(ws, addr)`), or `getValue` and friends. For a streaming handle use `Stream.getCellView`. Removed in 1.0.                                                                                                                                                                                                                |
| `Cell.Sheet`                                         | Use `Worksheet.Handle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Cell.Sheet.autoFilter`                              | Use `Worksheet.getAutoFilter` / `Worksheet.setAutoFilter`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                |
| `Cell.Sheet.tables`                                  | Use `Table.list` or `Table.get`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                          |
| `Cell.Sheet.pivotTables`                             | Use `Pivot.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                         |
| `Cell.Sheet.conditionalFormattings`                  | Use `ConditionalFormatting.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                         |
| `Cell.Sheet.formControls`                            | Use `Form.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                          |
| `Cell.Value`                                         | Use the flat `CellValue` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                    |
| `Row.getHidden`                                      | Use `isHidden`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Row.values`                                         | Use `getValues` with `{ base: 1 }`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                               |
| `Row.getCell`                                        | Address the cell by position through the `Cell` namespace (`Cell.getValue(ws, { row, col })`). Removed in 1.0.                                                                                                                                                                                                                                                                                    |
| `Row.eachCell`                                       | Use `eachAddress`, which visits by position and takes its options last. Removed in 1.0.                                                                                                                                                                                                                                                                                                           |
| `Row.commit`                                         | A document-model sheet has nothing to commit — every edit applies as it is made — and a streamed row is committed through its handle: `Stream.commitRow(ws.getRow(row))`. Removed in 1.0.                                                                                                                                                                                                         |
| `Row.Sheet`                                          | Use `Worksheet.Handle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Row.Sheet.autoFilter`                               | Use `Worksheet.getAutoFilter` / `Worksheet.setAutoFilter`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                |
| `Row.Sheet.tables`                                   | Use `Table.list` or `Table.get`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                          |
| `Row.Sheet.pivotTables`                              | Use `Pivot.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                         |
| `Row.Sheet.conditionalFormattings`                   | Use `ConditionalFormatting.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                         |
| `Row.Sheet.formControls`                             | Use `Form.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                          |
| `Column.getHidden`                                   | Use `isHidden`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Column.values`                                      | Use `getValues` with `{ base: 1 }`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                               |
| `Column.Sheet`                                       | Use `Worksheet.Handle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Column.Sheet.autoFilter`                            | Use `Worksheet.getAutoFilter` / `Worksheet.setAutoFilter`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                |
| `Column.Sheet.tables`                                | Use `Table.list` or `Table.get`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                          |
| `Column.Sheet.pivotTables`                           | Use `Pivot.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                         |
| `Column.Sheet.conditionalFormattings`                | Use `ConditionalFormatting.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                         |
| `Column.Sheet.formControls`                          | Use `Form.list`. Removed from the published type in 1.0.                                                                                                                                                                                                                                                                                                                                          |
| `Range.forEachAddress`                               | Use `eachAddress`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                |
| `Range.count`                                        | Use `getCount`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Chart.add`                                          | Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                                              |
| `Chart.addColumn`                                    | Use `Chart.place(ws, range, { ...options, type: "bar", barDir: "col" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                           |
| `Chart.addBar`                                       | Use `Chart.place(ws, range, { ...options, type: "bar", barDir: "bar" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                           |
| `Chart.addLine`                                      | Use `Chart.place(ws, range, { ...options, type: "line" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                         |
| `Chart.addArea`                                      | Use `Chart.place(ws, range, { ...options, type: "area" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                         |
| `Chart.addPie`                                       | Use `Chart.place(ws, range, { ...options, type: "pie" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                          |
| `Chart.addDoughnut`                                  | Use `Chart.place(ws, range, { ...options, type: "doughnut" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                     |
| `Chart.addScatter`                                   | Use `Chart.place(ws, range, { ...options, type: "scatter" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                      |
| `Chart.addBubble`                                    | Use `Chart.place(ws, range, { ...options, type: "bubble" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                       |
| `Chart.addRadar`                                     | Use `Chart.place(ws, range, { ...options, type: "radar" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                        |
| `Chart.addStock`                                     | Use `Chart.place(ws, range, { ...options, type: "stock" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                        |
| `Chart.addSurface`                                   | Use `Chart.place(ws, range, { ...options, type: "surface" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                      |
| `Chart.addHistogram`                                 | Use `Chart.place(ws, range, { ...options, type: "histogram" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                    |
| `Chart.addPareto`                                    | Use `Chart.place(ws, range, { ...options, type: "pareto" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                       |
| `Chart.addWaterfall`                                 | Use `Chart.place(ws, range, { ...options, type: "waterfall" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                    |
| `Chart.addFunnel`                                    | Use `Chart.place(ws, range, { ...options, type: "funnel" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                       |
| `Chart.addTreemap`                                   | Use `Chart.place(ws, range, { ...options, type: "treemap" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                      |
| `Chart.addSunburst`                                  | Use `Chart.place(ws, range, { ...options, type: "sunburst" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                     |
| `Chart.addBoxWhisker`                                | Use `Chart.place(ws, range, { ...options, type: "boxWhisker" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                   |
| `Chart.addRegionMap`                                 | Use `Chart.place(ws, range, { ...options, type: "regionMap" })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                    |
| `Chart.addCombo`                                     | Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                                              |
| `Chart.addEx`                                        | Use `Chart.place(ws, range, options)`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                                              |
| `Chart.addPreset`                                    | Use `Chart.place(ws, range, Chart.applyPreset(preset, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                   |
| `Chart.addPresetEx`                                  | Use `Chart.place(ws, range, Chart.applyPreset(preset, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                   |
| `Chart.addFromTable`                                 | Use `Chart.place(ws, range, Chart.createOptionsFromTable(ws, table, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                     |
| `Chart.addFromRows`                                  | Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                       |
| `Chart.addColumnFromRows`                            | Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, { ...options, type: "bar", barDir: "col" }))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                    |
| `Chart.addExFromTable`                               | Use `Chart.place(ws, range, Chart.createOptionsFromTable(ws, table, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                     |
| `Chart.addExFromRows`                                | Use `Chart.place(ws, range, Chart.createOptionsFromRows(ws, rows, options))`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                       |
| `Chart.addPivot`                                     | Use `Chart.place(ws, range, { ...options, pivotTable })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                           |
| `Chart.addPivotCombo`                                | Use `Chart.place(ws, range, { ...options, pivotTable })`, which takes the placement before the chart and returns the chart's handle; this returns its part number (`handle.chartNumber`, or `chartExNumber` for the Office 2016 types). Removed in 1.0.                                                                                                                                           |
| `Chart.seriesFromColumns`                            | Use `createSeriesFromColumns`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                    |
| `Chart.get`                                          | Use `list`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                       |
| `Chart.chartModel`                                   | Use `getModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Chart.chartExModel`                                 | Use `getExModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                 |
| `Chart.isChartEx`                                    | Use `isEx`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                       |
| `Chart.anchorModel`                                  | Use `getAnchorModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Chart.title`                                        | Use `getTitle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Chart.mutateChartEx`                                | Use `mutateEx`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Chart.setBuiltInStyle`                              | Use `setStyle`, which this forwards to. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                           |
| `Chart.totalSeriesCount`                             | Use `getTotalSeriesCount`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                        |
| `Chart.chartTypes`                                   | Use `getTypes`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Chart.axes`                                         | Use `getAxes`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Chart.categoryAxis`                                 | Use `getCategoryAxis`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                            |
| `Chart.valueAxis`                                    | Use `getValueAxis`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                               |
| `Chart.plotArea`                                     | Use `getPlotArea`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                |
| `Chart.legend`                                       | Use `getLegend`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                  |
| `Chart.spPr`                                         | Use `getSpPr`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Chart.titleRichText`                                | Use `getTitleRichText`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Chart.setTitleRichText`                             | Use `setTitle`, which takes rich text too. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                        |
| `Chart.unknownElements`                              | Use `getUnknownElements`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                         |
| `Chart.userShapesXml`                                | Use `getUserShapesXml`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Chart.clone`                                        | Use `copyTo` with the chart's own worksheet — `Chart.copyTo(c, c.worksheet, range)`. It places the same copy at the same default position and returns its handle; the part number this returned is the handle's `chartNumber` (`chartExNumber` for an Office 2016 chart). This name suggested an independent copy, as `Document.clone` returns, but it adds a chart to the sheet. Removed in 1.0. |
| `Chart.toSVG`                                        | Use `toSvg`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                      |
| `Chart.toPNG`                                        | Use `toPng`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                      |
| `Table.columnStyle`                                  | Use `getColumnStyle`, which returns a copy, and `setColumnStyle` to change it. This returns the column's own style object. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `Table.create`                                       | Use `add`. This builds a table bound to `worksheet` and writes its cells, but never registers it on the sheet, so it is not written, not found by `get` or `list`, and its name is not reserved. Removed in 1.0.                                                                                                                                                                                  |
| `Table.model`                                        | Use `getModel`, which returns a copy; change the table with `setModel` or the setters. This returns the model the table uses. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Table.commit`                                       | Every edit applies as it is made; there is never anything to commit. Removed in 1.0.                                                                                                                                                                                                                                                                                                              |
| `Table.column`                                       | Use `getColumn`, which refuses an index the table has no column at. Removed in 1.0.                                                                                                                                                                                                                                                                                                               |
| `Table.ref`                                          | Use `getRef`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                     |
| `Table.name`                                         | Use `getName`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Table.displayName`                                  | Use `getDisplayName`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Table.headerRow`                                    | Use `hasHeaderRow`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                               |
| `Table.totalsRow`                                    | Use `hasTotalsRow`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                               |
| `Table.theme`                                        | Use `getTheme`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Table.showFirstColumn`                              | Use `hasFirstColumnStyle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                        |
| `Table.setShowFirstColumn`                           | Use `setFirstColumnStyle`, the name `hasFirstColumnStyle` reads it back with. Removed in 1.0.                                                                                                                                                                                                                                                                                                     |
| `Table.showLastColumn`                               | Use `hasLastColumnStyle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                         |
| `Table.setShowLastColumn`                            | Use `setLastColumnStyle`, the name `hasLastColumnStyle` reads it back with. Removed in 1.0.                                                                                                                                                                                                                                                                                                       |
| `Table.showRowStripes`                               | Use `hasRowStripes`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Table.setShowRowStripes`                            | Use `setRowStripes`, the name `hasRowStripes` reads it back with. Removed in 1.0.                                                                                                                                                                                                                                                                                                                 |
| `Table.showColumnStripes`                            | Use `hasColumnStripes`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Table.setShowColumnStripes`                         | Use `setColumnStripes`, the name `hasColumnStripes` reads it back with. Removed in 1.0.                                                                                                                                                                                                                                                                                                           |
| `Table.columnName`                                   | Use `getColumnName`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Table.columnSetName`                                | Use `setColumnName`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Table.columnFilterButton`                           | Use `hasColumnFilterButton`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                      |
| `Table.columnSetFilterButton`                        | Use `setColumnFilterButton`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                      |
| `Table.columnSetStyle`                               | Use `setColumnStyle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Table.columnTotalsRowLabel`                         | Use `getColumnTotalsRowLabel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                    |
| `Table.columnSetTotalsRowLabel`                      | Use `setColumnTotalsRowLabel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                    |
| `Table.columnTotalsRowFunction`                      | Use `getColumnTotalsRowFunction`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                 |
| `Table.columnSetTotalsRowFunction`                   | Use `setColumnTotalsRowFunction`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                 |
| `Table.columnTotalsRowResult`                        | Use `getColumnTotalsRowResult`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                   |
| `Table.columnSetTotalsRowResult`                     | Use `setColumnTotalsRowResult`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                   |
| `Table.columnTotalsRowFormula`                       | Use `getColumnTotalsRowFormula`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                  |
| `Table.columnSetTotalsRowFormula`                    | Use `setColumnTotalsRowFormula`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                  |
| `Image.getShapes`                                    | Use `listShapes`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                 |
| `Image.create`                                       | Use `place`. This builds an image record bound to `worksheet` without placing it, so it is not listed, not written, and nothing public attaches it. Removed in 1.0.                                                                                                                                                                                                                               |
| `Image.model`                                        | Use `getModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Image.clone`                                        | Use `place` with the image's id and range (`getModel` gives both). This copies the record without placing the copy, so it is not listed and not written. Removed in 1.0.                                                                                                                                                                                                                          |
| `Image.applyModel`                                   | Use `setModel`. Removed in 1.0. Applies the model only, as it always did; `setModel` also records that a picture the image stops showing is no longer used.                                                                                                                                                                                                                                       |
| `Sparkline.AddSparklineGroupOptions`                 | Use the flat `AddSparklineGroupOptions` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                     |
| `Sparkline.SparklineGroup`                           | Use the flat `SparklineGroup` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                               |
| `Form.listCheckboxes`                                | Use `list`, which returns a copy. Removed in 1.0. This form returns the sheet's live array, as it always did: before `remove` existed, splicing it was how a checkbox was removed.                                                                                                                                                                                                                |
| `Form.createCheckbox`                                | Use `addCheckbox`. This builds a checkbox for `worksheet` without adding it, so two calls take the same shape id and neither is written. Removed in 1.0.                                                                                                                                                                                                                                          |
| `Form.checkboxFromModel`                             | Use `addCheckbox`. This builds a handle from a model without adding it to the sheet, so it is not listed and not written, and nothing public attaches it; a checkbox's model is `handle.model`, which the VML getters take as it is. Removed in 1.0.                                                                                                                                              |
| `Form.checked`                                       | Use `isChecked`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                  |
| `Form.link`                                          | Use `getLink`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Form.text`                                          | Use `getText`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Form.vmlAnchor`                                     | Use `Form.Vml.getAnchor`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                         |
| `Form.vmlStyle`                                      | Use `Form.Vml.getStyle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                          |
| `Form.vmlCheckedValue`                               | Use `Form.Vml.getCheckedValue`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                   |
| `Chartsheet.create`                                  | Use `Workbook.addChartsheet`. This builds a handle over `model` without adding it to a workbook, so it is not listed and not written. Removed in 1.0.                                                                                                                                                                                                                                             |
| `Chartsheet.id`                                      | Use `getId`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                      |
| `Chartsheet.sheetNo`                                 | Use `getSheetNo`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                 |
| `Chartsheet.name`                                    | Use `getName`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |
| `Chartsheet.state`                                   | Use `getState`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Chartsheet.chartNumber`                             | Use `getChartNumber`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Chartsheet.chartExNumber`                           | Use `getChartExNumber`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `Chartsheet.chartModel`                              | Use `getChartModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Chartsheet.chartExModel`                            | Use `getChartExModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                            |
| `Chartsheet.chart`                                   | Use `getChart`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Chartsheet.model`                                   | Use `getModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Chartsheet.pageMargins`                             | Use `getPageMargins`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `Chartsheet.pageSetup`                               | Use `getPageSetup`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                               |
| `Chartsheet.headerFooter`                            | Use `getHeaderFooter`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                            |
| `Chartsheet.tabSelected`                             | Use `isTabSelected`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `Chartsheet.zoomScale`                               | Use `getZoomScale`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                               |
| `Chartsheet.workbookViewId`                          | Use `getWorkbookViewId`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                          |
| `Chartsheet.zoomToFit`                               | Use `isZoomToFit`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                |
| `DataValidation.create`                              | A collection made here belongs to no sheet and nothing can attach it to one, so its rules are never written; pass the worksheet (or its own `ws.dataValidations`) to `add` instead. Removed in 1.0.                                                                                                                                                                                               |
| `DefinedNames.create`                                | A registry made here belongs to no workbook and nothing can attach it to one; use `Workbook.getDefinedNames(wb)`, which is the workbook's own. Removed in 1.0.                                                                                                                                                                                                                                    |
| `DefinedNames.getAllEntries`                         | Use `list`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                       |
| `DefinedNames.model`                                 | Use `getModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Note.model`                                         | Use `getModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Anchor.asInstance`                                  | Test with `isAnchor` and build one with `create``(undefined, input)`. This answered `null` for no input, where every other reader answers `undefined`, and turned anything else into an anchor unchecked. Removed in 1.0.                                                                                                                                                                         |
| `Anchor.col`                                         | Use `getCol`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                     |
| `Anchor.row`                                         | Use `getRow`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                     |
| `Anchor.colWidth`                                    | Use `getColWidth`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                |
| `Anchor.rowHeight`                                   | Use `getRowHeight`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                               |
| `Anchor.model`                                       | Use `getModel`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                   |
| `Watermark.add`                                      | Use `set`: a sheet has one watermark, and this replaces it rather than adding a second. Removed in 1.0.                                                                                                                                                                                                                                                                                           |
| `HeaderFooterImage.Entry`                            | Use the flat `HeaderFooterImageEntry` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                       |
| `HeaderFooterImage.Options`                          | Use the flat `HeaderFooterImageOptions` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                     |
| `HeaderFooterImage.Position`                         | Use the flat `HeaderFooterImagePosition` type. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                    |
| `Format.parseValueByFormat`                          | Use `parseValue`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                 |
| `BuiltinCellStyle`                                   | Use `BuiltInCellStyle`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                           |
| `CellValueType`                                      | Use `CellValue`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                  |
| `CellValueInputType`                                 | Use `CellValueInput`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `AddAOAOptions`                                      | Use `AddAoaOptions`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                              |
| `AddJSONOptions`                                     | Use `AddJsonOptions`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                             |
| `SheetToJSONOptions`                                 | Use `SheetToJsonOptions`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                         |
| `WorksheetNameError`                                 | Use `ExcelWorksheetNameError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                    |
| `InvalidAddressError`                                | Use `ExcelInvalidAddressError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                   |
| `ColumnOutOfBoundsError`                             | Use `ExcelColumnOutOfBoundsError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                |
| `RowOutOfBoundsError`                                | Use `ExcelRowOutOfBoundsError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                   |
| `MergeConflictError`                                 | Use `ExcelMergeConflictError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                    |
| `InvalidValueTypeError`                              | Use `ExcelInvalidValueTypeError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                 |
| `ExcelNotSupportedError`                             | Use `ExcelUnsupportedFeatureError` — the same class, named as Word's `DocxUnsupportedFeatureError` is. Removed in 1.0.                                                                                                                                                                                                                                                                            |
| `PivotTableError`                                    | Use `ExcelPivotTableError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                       |
| `TableError`                                         | Use `ExcelTableError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                            |
| `ImageError`                                         | Use `ExcelImageError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                            |
| `MaxItemsExceededError`                              | Use `ExcelLimitExceededError`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                    |
| `MaxItemsExceededError#maxItems`                     | Use `allowed`. Removed in 1.0.                                                                                                                                                                                                                                                                                                                                                                    |

**`documonster/excel/csv`**

| Deprecated                         | Replacement                                                                                                                                                                                                                                                                                     |
| ---------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `createCsvReadStream`              | Use `ExcelCsv.toStream`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `readCsv`                          | Use `ExcelCsv.addWorksheet`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `writeCsv`                         | Use `ExcelCsv.format` for a string or `ExcelCsv.writeStream` for a stream. Removed in 1.0.                                                                                                                                                                                                      |
| `writeCsvBuffer`                   | Use `ExcelCsv.toBuffer`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `createCsvWriteStream`             | Use `ExcelCsv.createImportStream`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `CsvInput`                         | Use `ExcelCsvInput`. Removed in 1.0.                                                                                                                                                                                                                                                            |
| `CsvOptions`                       | Use `ExcelCsvOptions`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `CsvOptions.info?`                 | Has no effect on a workbook: a record's position has nowhere to go in a sheet. Removed in 1.0.                                                                                                                                                                                                  |
| `CsvOptions.raw?`                  | Has no effect on a workbook: a record's raw text has nowhere to go in a sheet. Removed in 1.0.                                                                                                                                                                                                  |
| `CsvOptions.skipRecordsWithError?` | Has no effect: a malformed record (a column-count mismatch the `columnMismatch` strategy rejects, or text after a closing quote) is always left out of `rows`, reported in `errors` / `invalidRows` and handed to `onSkip`. It only used to decide whether `onSkip` was called. Removed in 1.0. |
| `readCsvFile`                      | Use `ExcelCsv.readFile` for a new workbook, or `ExcelCsv.addWorksheet` with a file stream to add the sheet to one you have. Removed in 1.0.                                                                                                                                                     |
| `writeCsvFile`                     | Use `ExcelCsv.writeFile`. Removed in 1.0.                                                                                                                                                                                                                                                       |

**`documonster/excel/csv [browser]`**

| Deprecated            | Replacement                              |
| --------------------- | ---------------------------------------- |
| `createCsvReadStream` | Use `ExcelCsv.toStream`. Removed in 1.0. |

**`documonster/excel/formula`**

| Deprecated          | Replacement                                   |
| ------------------- | --------------------------------------------- |
| `calculateFormulas` | Use `ExcelFormula.calculate`. Removed in 1.0. |

**`documonster/excel/markdown`**

| Deprecated            | Replacement                                                                                                                 |
| --------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `readMarkdown`        | Use `ExcelMarkdown.addWorksheet`. Removed in 1.0.                                                                           |
| `readMarkdownAll`     | Use `ExcelMarkdown.addWorksheets`. Removed in 1.0.                                                                          |
| `writeMarkdown`       | Use `ExcelMarkdown.format`. Removed in 1.0.                                                                                 |
| `writeMarkdownBuffer` | Use `ExcelMarkdown.toBufferSync` (or the asynchronous `ExcelMarkdown.toBuffer`). Removed in 1.0.                            |
| `readMarkdownFile`    | Use `ExcelMarkdown.readFile` for a new workbook, or `ExcelMarkdown.addWorksheet`. Removed in 1.0.                           |
| `readMarkdownAllFile` | Use `ExcelMarkdown.readFile(path, { tables: "all" })` for a new workbook, or `ExcelMarkdown.addWorksheets`. Removed in 1.0. |
| `writeMarkdownFile`   | Use `ExcelMarkdown.writeFile`. Removed in 1.0.                                                                              |

**`documonster/formula`**

| Deprecated      | Replacement                           |
| --------------- | ------------------------------------- |
| `Formula.print` | Use `format`. Removed in 1.0.         |
| `Token`         | Use `FormulaToken`. Removed in 1.0.   |
| `AstNode`       | Use `FormulaAstNode`. Removed in 1.0. |

**`documonster/markdown`**

| Deprecated        | Replacement                                                                                                                                                                                                                  |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `MarkdownOptions` | Worksheet options do not belong to the standalone Markdown module. Use `ExcelMarkdownReadOptions` / `ExcelMarkdownWriteOptions` from `documonster/excel/markdown`, which the `ExcelMarkdown` functions take. Removed in 1.0. |

**`documonster/mermaid`**

| Deprecated             | Replacement                                                                                                                                                                               |
| ---------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `mermaidToDrawList`    | Use `Mermaid.toDrawList`. Removed in 1.0.                                                                                                                                                 |
| `mermaidToSvg`         | Use `Mermaid.toSvg`. Removed in 1.0.                                                                                                                                                      |
| `MermaidSyntaxError`   | Use `MermaidParseError` — the same class, under the name every module's parse error has (`XmlParseError`, `MarkdownParseError`, …), so `instanceof` either answers alike. Removed in 1.0. |
| `parseMermaid`         | Use `Mermaid.parse`. Removed in 1.0.                                                                                                                                                      |
| `layoutFlowchart`      | Use `Mermaid.layoutGraph`, which also lays out state, class, ER, requirement, C4 and architecture diagrams. Removed in 1.0.                                                               |
| `LayoutOptions`        | Use `MermaidLayoutOptions`. Removed in 1.0.                                                                                                                                               |
| `ThemeOptions`         | Use `MermaidThemeOptions`. Removed in 1.0.                                                                                                                                                |
| `EdgeRoute`            | Use `MermaidEdgeRoute`. Removed in 1.0.                                                                                                                                                   |
| `FlowLayout`           | Use `MermaidFlowLayout`. Removed in 1.0.                                                                                                                                                  |
| `GroupBox`             | Use `MermaidGroupBox`. Removed in 1.0.                                                                                                                                                    |
| `NodeBox`              | Use `MermaidNodeBox`. Removed in 1.0.                                                                                                                                                     |
| `centredText`          | Use `Mermaid.createCenteredText`, which takes the font as one style object. Removed in 1.0.                                                                                               |
| `resolveTheme`         | Use `Mermaid.resolveTheme`. Removed in 1.0.                                                                                                                                               |
| `Theme`                | Use `MermaidTheme`. Removed in 1.0.                                                                                                                                                       |
| `ClassBox`             | Use `MermaidClassBox`. Removed in 1.0.                                                                                                                                                    |
| `ClassDef`             | Use `MermaidClassDef`. Removed in 1.0.                                                                                                                                                    |
| `ClassDiagram`         | Use `MermaidClassDiagram`. Removed in 1.0.                                                                                                                                                |
| `ClassLink`            | Use `MermaidClassLink`. Removed in 1.0.                                                                                                                                                   |
| `ClassMember`          | Use `MermaidClassMember`. Removed in 1.0.                                                                                                                                                 |
| `ClassRelation`        | Use `MermaidClassRelation`. Removed in 1.0.                                                                                                                                               |
| `Cardinality`          | Use `MermaidCardinality`. Removed in 1.0.                                                                                                                                                 |
| `EdgeEnd`              | Use `MermaidEdgeEnd`. Removed in 1.0.                                                                                                                                                     |
| `Entity`               | Use `MermaidEntity`. Removed in 1.0.                                                                                                                                                      |
| `EntityAttribute`      | Use `MermaidEntityAttribute`. Removed in 1.0.                                                                                                                                             |
| `EntityRelation`       | Use `MermaidEntityRelation`. Removed in 1.0.                                                                                                                                              |
| `ErDiagram`            | Use `MermaidErDiagram`. Removed in 1.0.                                                                                                                                                   |
| `GanttDiagram`         | Use `MermaidGanttDiagram`. Removed in 1.0.                                                                                                                                                |
| `GanttTask`            | Use `MermaidGanttTask`. Removed in 1.0.                                                                                                                                                   |
| `JourneyDiagram`       | Use `MermaidJourneyDiagram`. Removed in 1.0.                                                                                                                                              |
| `JourneyTask`          | Use `MermaidJourneyTask`. Removed in 1.0.                                                                                                                                                 |
| `GitCommit`            | Use `MermaidGitCommit`. Removed in 1.0.                                                                                                                                                   |
| `GitGraphDiagram`      | Use `MermaidGitGraphDiagram`. Removed in 1.0.                                                                                                                                             |
| `MindNode`             | Use `MermaidMindNode`. Removed in 1.0.                                                                                                                                                    |
| `MindShape`            | Use `MermaidMindShape`. Removed in 1.0.                                                                                                                                                   |
| `MindmapDiagram`       | Use `MermaidMindmapDiagram`. Removed in 1.0.                                                                                                                                              |
| `ArchitectureDiagram`  | Use `MermaidArchitectureDiagram`. Removed in 1.0.                                                                                                                                         |
| `ArchitectureEdge`     | Use `MermaidArchitectureEdge`. Removed in 1.0.                                                                                                                                            |
| `ArchitectureNode`     | Use `MermaidArchitectureNode`. Removed in 1.0.                                                                                                                                            |
| `BlockCell`            | Use `MermaidBlockCell`. Removed in 1.0.                                                                                                                                                   |
| `BlockDiagram`         | Use `MermaidBlockDiagram`. Removed in 1.0.                                                                                                                                                |
| `C4Boundary`           | Use `MermaidC4Boundary`. Removed in 1.0.                                                                                                                                                  |
| `C4Diagram`            | Use `MermaidC4Diagram`. Removed in 1.0.                                                                                                                                                   |
| `C4Element`            | Use `MermaidC4Element`. Removed in 1.0.                                                                                                                                                   |
| `C4Relation`           | Use `MermaidC4Relation`. Removed in 1.0.                                                                                                                                                  |
| `KanbanCard`           | Use `MermaidKanbanCard`. Removed in 1.0.                                                                                                                                                  |
| `KanbanColumn`         | Use `MermaidKanbanColumn`. Removed in 1.0.                                                                                                                                                |
| `KanbanDiagram`        | Use `MermaidKanbanDiagram`. Removed in 1.0.                                                                                                                                               |
| `PacketDiagram`        | Use `MermaidPacketDiagram`. Removed in 1.0.                                                                                                                                               |
| `PacketField`          | Use `MermaidPacketField`. Removed in 1.0.                                                                                                                                                 |
| `QuadrantDiagram`      | Use `MermaidQuadrantDiagram`. Removed in 1.0.                                                                                                                                             |
| `QuadrantPoint`        | Use `MermaidQuadrantPoint`. Removed in 1.0.                                                                                                                                               |
| `SankeyDiagram`        | Use `MermaidSankeyDiagram`. Removed in 1.0.                                                                                                                                               |
| `RadarDiagram`         | Use `MermaidRadarDiagram`. Removed in 1.0.                                                                                                                                                |
| `RadarSeries`          | Use `MermaidRadarSeries`. Removed in 1.0.                                                                                                                                                 |
| `Requirement`          | Use `MermaidRequirement`. Removed in 1.0.                                                                                                                                                 |
| `RequirementDiagram`   | Use `MermaidRequirementDiagram`. Removed in 1.0.                                                                                                                                          |
| `RequirementElement`   | Use `MermaidRequirementElement`. Removed in 1.0.                                                                                                                                          |
| `RequirementLink`      | Use `MermaidRequirementLink`. Removed in 1.0.                                                                                                                                             |
| `SankeyLink`           | Use `MermaidSankeyLink`. Removed in 1.0.                                                                                                                                                  |
| `XyDiagram`            | Use `MermaidXyDiagram`. Removed in 1.0.                                                                                                                                                   |
| `XySeries`             | Use `MermaidXySeries`. Removed in 1.0.                                                                                                                                                    |
| `TimelineDiagram`      | Use `MermaidTimelineDiagram`. Removed in 1.0.                                                                                                                                             |
| `TimelinePeriod`       | Use `MermaidTimelinePeriod`. Removed in 1.0.                                                                                                                                              |
| `EdgeStroke`           | Use `MermaidEdgeStroke`. Removed in 1.0.                                                                                                                                                  |
| `FlowDirection`        | Use `MermaidFlowDirection`. Removed in 1.0.                                                                                                                                               |
| `FlowEdge`             | Use `MermaidFlowEdge`. Removed in 1.0.                                                                                                                                                    |
| `FlowNode`             | Use `MermaidFlowNode`. Removed in 1.0.                                                                                                                                                    |
| `FlowShape`            | Use `MermaidFlowShape`. Removed in 1.0.                                                                                                                                                   |
| `FlowSubgraph`         | Use `MermaidFlowSubgraph`. Removed in 1.0.                                                                                                                                                |
| `FlowchartDiagram`     | Use `MermaidFlowchartDiagram`. Removed in 1.0.                                                                                                                                            |
| `PieDiagram`           | Use `MermaidPieDiagram`. Removed in 1.0.                                                                                                                                                  |
| `StateDiagram`         | Use `MermaidStateDiagram`. Removed in 1.0.                                                                                                                                                |
| `StateNode`            | Use `MermaidStateNode`. Removed in 1.0.                                                                                                                                                   |
| `StateTransition`      | Use `MermaidStateTransition`. Removed in 1.0.                                                                                                                                             |
| `TaskState`            | Use `MermaidTaskState`. Removed in 1.0.                                                                                                                                                   |
| `Visibility`           | Use `MermaidVisibility`. Removed in 1.0.                                                                                                                                                  |
| `PieSlice`             | Use `MermaidPieSlice`. Removed in 1.0.                                                                                                                                                    |
| `SequenceDiagram`      | Use `MermaidSequenceDiagram`. Removed in 1.0.                                                                                                                                             |
| `SequenceLine`         | Use `MermaidSequenceLine`. Removed in 1.0.                                                                                                                                                |
| `SequenceMessage`      | Use `MermaidSequenceMessage`. Removed in 1.0.                                                                                                                                             |
| `SequenceParticipant`  | Use `MermaidSequenceParticipant`. Removed in 1.0.                                                                                                                                         |
| `SequenceBlock`        | Use `MermaidSequenceBlock`. Removed in 1.0.                                                                                                                                               |
| `SequenceEntry`        | Use `MermaidSequenceEntry`. Removed in 1.0.                                                                                                                                               |
| `SequenceMessageEntry` | Use `MermaidSequenceMessageEntry`. Removed in 1.0.                                                                                                                                        |
| `SequenceNote`         | Use `MermaidSequenceNote`. Removed in 1.0.                                                                                                                                                |
| `SequenceSection`      | Use `MermaidSequenceSection`. Removed in 1.0.                                                                                                                                             |
| `NotePlacement`        | Use `MermaidNotePlacement`. Removed in 1.0.                                                                                                                                               |

**`documonster/pdf`**

| Deprecated                          | Replacement                                                                                                                                                                             |
| ----------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Pdf.wordChartRenderer`             | Omit `chartRenderer`: the default draws every chart — classic and ChartEx — with the Excel engine, and this callback draws a classic chart exactly as the default does. Removed in 1.0. |
| `Pdf.create`                        | Use `fromTable`. `create` is the name every other namespace uses for a handle constructor, and this returns bytes. Removed in 1.0.                                                      |
| `Pdf.read`                          | Use `extract`. `read` is the verb that opens a document as its editable object (`Pdf.Editor.read`); this extracts content. Removed in 1.0.                                              |
| `Pdf.Builder#textLanguage`          | Use `PdfDocumentBuilder.setTextLanguage`. Removed in 1.0.                                                                                                                               |
| `Pdf.Builder#build`                 | Use `PdfDocumentBuilder.toBuffer`. Removed in 1.0.                                                                                                                                      |
| `Pdf.Editor#save`                   | Use `PdfEditor.toBuffer` (which keeps the encryption by default; pass `{ encryption: "remove" }` for this method's decrypted copy). Removed in 1.0.                                     |
| `Pdf.Editor#saveIncremental`        | Use `toBuffer({ incremental: true })` (whose fallback keeps the encryption by default). Removed in 1.0.                                                                                 |
| `Pdf.Editor.load`                   | Use `PdfEditor.read`. Removed in 1.0.                                                                                                                                                   |
| `Pdf.verifySignature`               | Use `Pdf.Signature.verify`, or `Pdf.verifySignatures` to find and check every signature in a file. Removed in 1.0.                                                                      |
| `Pdf.sign`                          | Use `Pdf.Signature.fillPlaceholder(bytes, { certificate, privateKey })` — this only completes a placeholder already in the file. Removed in 1.0.                                        |
| `Pdf.buildSignatureDictPlaceholder` | Use `Pdf.Signature.buildDictPlaceholder`. Removed in 1.0.                                                                                                                               |
| `Pdf.asn1Parse`                     | Use `Pdf.Signature.parseAsn1`. Removed in 1.0.                                                                                                                                          |
| `ChartToPdfOptions`                 | Use `PdfFromChartOptions`. Removed in 1.0.                                                                                                                                              |
| `ExcelToPdfOptions`                 | Use `PdfFromExcelOptions`. Removed in 1.0.                                                                                                                                              |
| `ExcelToPdfOptions.font?`           | Use `PdfDocumentOptions.fonts` (`{ fonts: { default: { regular: bytes } } }`). Removed in 1.0.                                                                                          |
| `DocxToPdfOptions`                  | Use `PdfFromDocxOptions`. Removed in 1.0.                                                                                                                                               |
| `DocxToPdfOptions.marginTop?`       | Use `margins: { top }`. Removed in 1.0. `margins.top` wins when both are given.                                                                                                         |
| `DocxToPdfOptions.marginBottom?`    | Use `margins: { bottom }`. Removed in 1.0.                                                                                                                                              |
| `DocxToPdfOptions.marginLeft?`      | Use `margins: { left }`. Removed in 1.0.                                                                                                                                                |
| `DocxToPdfOptions.marginRight?`     | Use `margins: { right }`. Removed in 1.0.                                                                                                                                               |
| `DocxToPdfOptions.defaultFont?`     | Use `defaultFontFamily`, which wins when both are given. Removed in 1.0.                                                                                                                |
| `CjkLanguage`                       | Use `PdfCjkLanguage`. Removed in 1.0.                                                                                                                                                   |
| `PdfExportOptions.font?`            | Use `PdfDocumentOptions.fonts` (`{ fonts: { default: { regular: bytes } } }`). Removed in 1.0.                                                                                          |
| `PageSizeName`                      | Use `PdfPageSizeName`. Removed in 1.0.                                                                                                                                                  |
| `ReadPdfOptions`                    | Use `PdfExtractOptions`. Removed in 1.0.                                                                                                                                                |
| `ReadPdfResult`                     | Use `PdfExtractResult`. Removed in 1.0.                                                                                                                                                 |
| `ReadPdfPage`                       | Use `PdfExtractedPage`. Removed in 1.0.                                                                                                                                                 |
| `ExtractedImage`                    | Use `PdfExtractedImage`. Removed in 1.0.                                                                                                                                                |
| `TextLine`                          | Use `PdfTextLine`. Removed in 1.0.                                                                                                                                                      |
| `PdfRect`                           | Use `PdfAnnotationRect`. A rectangle by its corners, as an annotation's `/Rect` gives it; the general name read as the `{ x, y, width, height }` shape. Removed in 1.0.                 |
| `PageOptions`                       | Use `PdfPageOptions`. Removed in 1.0.                                                                                                                                                   |
| `DrawSvgOptions`                    | Use `PdfDrawSvgOptions`. Removed in 1.0.                                                                                                                                                |
| `DrawTextOptions`                   | Use `PdfDrawTextOptions`. Removed in 1.0.                                                                                                                                               |
| `DrawRectOptions`                   | Use `PdfDrawRectOptions`. Removed in 1.0.                                                                                                                                               |
| `DrawCircleOptions`                 | Use `PdfDrawCircleOptions`. Removed in 1.0.                                                                                                                                             |
| `DrawEllipseOptions`                | Use `PdfDrawEllipseOptions`. Removed in 1.0.                                                                                                                                            |
| `DrawLineOptions`                   | Use `PdfDrawLineOptions`. Removed in 1.0.                                                                                                                                               |
| `DrawPathOptions`                   | Use `PdfDrawPathOptions`. Removed in 1.0.                                                                                                                                               |
| `DrawImageOptions`                  | Use `PdfDrawImageOptions`. Removed in 1.0.                                                                                                                                              |
| `DocumentMetadata`                  | Use `PdfDocumentMetadata`. Removed in 1.0.                                                                                                                                              |
| `PathOp`                            | Use `PdfPathOp`. Removed in 1.0.                                                                                                                                                        |
| `TocOptions`                        | Use `PdfTocOptions`. Removed in 1.0.                                                                                                                                                    |
| `AnnotationType`                    | Use `PdfAnnotationType`. Removed in 1.0.                                                                                                                                                |
| `AnnotationOptions`                 | Use `PdfAnnotationOptions`. Removed in 1.0.                                                                                                                                             |
| `TextMarkupAnnotationOptions`       | Use `PdfTextMarkupAnnotationOptions`. Removed in 1.0.                                                                                                                                   |
| `TextAnnotationOptions`             | Use `PdfTextAnnotationOptions`. Removed in 1.0.                                                                                                                                         |
| `FreeTextAnnotationOptions`         | Use `PdfFreeTextAnnotationOptions`. Removed in 1.0.                                                                                                                                     |
| `StampAnnotationOptions`            | Use `PdfStampAnnotationOptions`. Removed in 1.0.                                                                                                                                        |
| `FormFieldOptions`                  | Use `PdfFormFieldOptions`. Removed in 1.0.                                                                                                                                              |
| `TextFieldOptions`                  | Use `PdfTextFieldOptions`. Removed in 1.0.                                                                                                                                              |
| `CheckboxOptions`                   | Use `PdfCheckboxOptions`. Removed in 1.0.                                                                                                                                               |
| `DropdownOptions`                   | Use `PdfDropdownOptions`. Removed in 1.0.                                                                                                                                               |
| `RadioGroupOptions`                 | Use `PdfRadioGroupOptions`. Removed in 1.0.                                                                                                                                             |
| `LoadOptions`                       | Use `PdfEditorReadOptions`. Removed in 1.0.                                                                                                                                             |
| `EditorSaveOptions`                 | Use `PdfEditorToBufferOptions` with `toBuffer()`, or `PdfEditorWriteOptions` for what every write shares. Removed in 1.0.                                                               |
| `EditorSignOptions`                 | Use `PdfEditorSignOptions`. Removed in 1.0.                                                                                                                                             |
| `PdfContentStream#deferred`         | The builder's own machinery for text whose encoding is chosen at build time, not part of the drawing API; it leaves the public type in 1.0. Draw with the typed operators or `raw`.     |
| `PdfContentStream#preflight`        | The builder's own preparation step for deferred fragments, not part of the drawing API; it leaves the public type in 1.0.                                                               |
| `SignatureVerificationResult`       | What the deprecated `Pdf.verifySignature` returns; `Pdf.Signature.verify` returns `PdfSignatureCheck`. Removed in 1.0.                                                                  |
| `CmsSignedData`                     | No public function takes or returns this. Removed in 1.0.                                                                                                                               |
| `SignOptions`                       | No public function takes this; signing options are `PdfSignOptions`(`Pdf.Signature.fillPlaceholder`) and `PdfSignatureOptions` (`sign`). Removed in 1.0.                                |
| `Asn1Node`                          | Use `PdfAsn1Node`. Removed in 1.0.                                                                                                                                                      |
| `createPdfDrawSurface`              | Use `Pdf.createDrawSurface(page, rect, { scale })`. Removed in 1.0.                                                                                                                     |
| `DrawSurfaceRect`                   | Use `PdfDrawSurfaceRect`. Removed in 1.0.                                                                                                                                               |

**`documonster/stream`**

| Deprecated                                | Replacement                                                                                                                             |
| ----------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| `Collector`                               | Use `Streams.Collector` (`new Streams.Collector(options)`). Removed in 1.0.                                                             |
| `createCollector`                         | Use `new Streams.Collector(options)`. Removed in 1.0.                                                                                   |
| `PullStream`                              | Use `Streams.PullStream` (`new Streams.PullStream(options)`). Removed in 1.0.                                                           |
| `BufferedStream`                          | Use `Streams.BufferedStream` (`new Streams.BufferedStream(options)`). Removed in 1.0.                                                   |
| `BufferedStream#toUint8Array`             | Use `readAll`, which does the same: despite the `to` name this consumes the buffered data. Removed in 1.0.                              |
| `createStringChunk`                       | Use `Streams.createStringChunk`. Removed in 1.0.                                                                                        |
| `createByteChunk`                         | Use `Streams.createByteChunk`. Removed in 1.0.                                                                                          |
| `createReadable`                          | Use `Streams.createReadable`. Removed in 1.0.                                                                                           |
| `createReadableFromAsyncIterable`         | Use `Streams.createReadableFromAsyncIterable`. Removed in 1.0.                                                                          |
| `createReadableFromArray`                 | Use `Streams.createReadableFromArray`. Removed in 1.0.                                                                                  |
| `createWritable`                          | Use `Streams.createWritable`. Removed in 1.0.                                                                                           |
| `createTransform`                         | Use `Streams.createTransform`. Removed in 1.0.                                                                                          |
| `createPassThrough`                       | Use `Streams.createPassThrough`. Removed in 1.0.                                                                                        |
| `createPullStream`                        | Use `new Streams.PullStream(options)`. Removed in 1.0.                                                                                  |
| `createBufferedStream`                    | Use `new Streams.BufferedStream(options)`. Removed in 1.0.                                                                              |
| `createDuplex`                            | Use `Streams.createDuplex`. Removed in 1.0.                                                                                             |
| `createReadableFromGenerator`             | Use `Streams.createReadableFromGenerator`. Removed in 1.0.                                                                              |
| `createReadableFromPromise`               | Use `Streams.createReadableFromPromise`. Removed in 1.0.                                                                                |
| `createEmptyReadable`                     | Use `Streams.createEmptyReadable`. Removed in 1.0.                                                                                      |
| `createNullWritable`                      | Use `Streams.createNullWritable`. Removed in 1.0.                                                                                       |
| `filter`                                  | Use `Streams.filter`. Removed in 1.0.                                                                                                   |
| `fromBytes`                               | Use `Streams.fromBytes`. Removed in 1.0.                                                                                                |
| `fromJSON`                                | Use `Streams.fromJson`. Removed in 1.0.                                                                                                 |
| `fromString`                              | Use `Streams.fromString`. Removed in 1.0.                                                                                               |
| `transform`                               | Use `Streams.createTransform`, which this duplicated. Removed in 1.0.                                                                   |
| `finishedAll`                             | Use `Streams.waitForFinished`. Removed in 1.0.                                                                                          |
| `streamToPromise`                         | Use `finished`, which this duplicated. Removed in 1.0.                                                                                  |
| `streamToUint8Array`                      | Use `Streams.bytes`. Removed in 1.0.                                                                                                    |
| `streamToBuffer`                          | Use `Streams.bytes`, which this duplicated: both resolve to a `Uint8Array`, not a `Buffer`. Removed in 1.0.                             |
| `streamToString`                          | Use `Streams.text`; pass the encoding as `{ encoding }`. Removed in 1.0.                                                                |
| `drainStream`                             | Use `Streams.discard`. Removed in 1.0.                                                                                                  |
| `copyStream`                              | Use `pipeline`, which this duplicated. Removed in 1.0.                                                                                  |
| `isTransform`                             | Use `Streams.isTransform`. Removed in 1.0.                                                                                              |
| `isDuplex`                                | Use `Streams.isDuplex`. Removed in 1.0.                                                                                                 |
| `isStream`                                | Use `Streams.isStream`. Removed in 1.0.                                                                                                 |
| `readableFromWeb`                         | Use `Streams.fromWebReadable`. Removed in 1.0.                                                                                          |
| `readableToWeb`                           | Use `Streams.toWebReadable`. Removed in 1.0.                                                                                            |
| `toWritable`                              | Use `Streams.toWritable`. Removed in 1.0.                                                                                               |
| `writableFromWeb`                         | Use `Streams.fromWebWritable`. Removed in 1.0.                                                                                          |
| `writableToWeb`                           | Use `Streams.toWebWritable`. Removed in 1.0.                                                                                            |
| `IBufferedStream.toUint8Array`            | The class's `readAll()` is the same read under a name that says it consumes. Removed in 1.0.                                            |
| `DataChunk`                               | Use `StreamDataChunk`. Removed in 1.0.                                                                                                  |
| `EventListener`                           | Use `StreamEventListener`; the bare name shadows the DOM's global `EventListener`. Removed in 1.0.                                      |
| `promisify`                               | Use `Streams.LowLevel.fromCallback`: `promisify` read as `node:util`'s, which wraps a function rather than running one. Removed in 1.0. |
| `ChunkedBuilder`                          | Use `Streams.ChunkedBuilder`. Removed in 1.0.                                                                                           |
| `TransactionalChunkedBuilder`             | Use `Streams.TransactionalChunkedBuilder`. Removed in 1.0.                                                                              |
| `bytes`                                   | Use `Streams.bytes`. Removed in 1.0.                                                                                                    |
| `collect`                                 | Use `Streams.collect`. Removed in 1.0.                                                                                                  |
| `isReadableStreamLike`                    | Use `Streams.isReadableStream`, which this duplicated. Removed in 1.0.                                                                  |
| `json`                                    | Use `Streams.json`. Removed in 1.0.                                                                                                     |
| `readableStreamToAsyncIterable`           | Use `Streams.readableStreamToAsyncIterable`. Removed in 1.0.                                                                            |
| `text`                                    | Use `Streams.text`. Removed in 1.0.                                                                                                     |
| `StreamTypeError`                         | Use `StreamArgumentError`. Removed in 1.0.                                                                                              |
| `UnsupportedStreamTypeError`              | Use `StreamUnsupportedTypeError`. Removed in 1.0.                                                                                       |
| `isReadableStream`                        | Use `Streams.isReadableStream`. Removed in 1.0.                                                                                         |
| `isWritableStream`                        | Use `Streams.isWritableStream`. Removed in 1.0.                                                                                         |
| `isAsyncIterable`                         | Use `Streams.isAsyncIterable`. Removed in 1.0.                                                                                          |
| `isTransformStream`                       | Use `Streams.isTransformStream`. Removed in 1.0.                                                                                        |
| `onceEvent`                               | Use `Streams.LowLevel.onceEvent`. Removed in 1.0.                                                                                       |
| `eventedReadableToAsyncIterableNoDestroy` | Use `Streams.LowLevel.eventsToAsyncIterableNoDestroy`. Removed in 1.0.                                                                  |
| `SinkAdapter`                             | Use `Streams.LowLevel.SinkAdapter`. Removed in 1.0.                                                                                     |

**`documonster/word`**

| Deprecated                               | Replacement                                                                                                                                                                                                                                                                              |
| ---------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `DocxOptions`                            | No function takes this type; use `DocxWriteOptions` (the options of `Document.toBuffer`), which has `compressionLevel` too. Removed in 1.0.                                                                                                                                              |
| `DocxLimitExceededError#maximum`         | Use `allowed`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `Document.insertContentAt`               | Use `insertContent`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Document.addParagraphElement`           | Use `addParagraph`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `Document.addTableElement`               | Use `addTable`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `Document.nextBookmarkId`                | Use `allocateBookmarkId`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `Document.build`                         | A document needs no build step: pass the document itself to `Document.toBuffer`, `Layout.*`, `WordHtml.fromDocx` or `Pdf.fromDocx`, which apply the same write-time defaults. Removed in 1.0.                                                                                            |
| `Document.DocumentHandle`                | Use `DocxDocument`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `DocumentHandle`                         | Use `DocxDocument`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `Build.checkBox`                         | Use `checkbox`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `Build.mathSuperScript`                  | Use `mathSuperscript`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `Build.mathSubScript`                    | Use `mathSubscript`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Build.mathSubSuperScript`               | Use `mathSubSuperscript`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `Build.mathPreSubSuperScript`            | Use `mathPreSubSuperscript`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `Build.drawingShape`                     | Use `shape`, which builds the same shape and more: `fillColor` → `fill: { type: "solid", color }`, `noFill` → `fill: { type: "none" }`, `outlineColor` / `outlineWidth` / `noOutline` → `outline: { color, width, noLine }`, `textContent` → `textBody: { paragraphs }`. Removed in 1.0. |
| `Build.createShape`                      | Use `shape`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `Build.createRect`                       | Use `rect`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `Build.createRoundRect`                  | Use `roundRect`. Removed in 1.0.                                                                                                                                                                                                                                                         |
| `Build.createEllipse`                    | Use `ellipse`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `Build.createLine`                       | Use `line`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `Build.createArrow`                      | Use `arrow`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `Build.createFlowchartShape`             | Use `flowchartShape`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Build.createCallout`                    | Use `callout`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `Build.createStar`                       | Use `star`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `Query.paragraphCount`                   | Use `getParagraphCount`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Query.countWords`                       | Use `getWordCount`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `Query.getHeadings`                      | Use `listHeadings`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `Query.findBookmark`                     | Use `getBookmark`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `Query.findComment`                      | Use `getComment`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `Query.tableCount`                       | Use `getTableCount`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Query.extractText`                      | Use `getText`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `Query.replaceText`                      | Use `Document.replaceText`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `Query.mailMerge`                        | Use `Document.mailMerge`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `Query.acceptAllRevisions`               | Use `Document.acceptAllRevisions`. Removed in 1.0.                                                                                                                                                                                                                                       |
| `Query.rejectAllRevisions`               | Use `Document.rejectAllRevisions`. Removed in 1.0.                                                                                                                                                                                                                                       |
| `Query.acceptRevision`                   | Use `Document.acceptRevision`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `Query.rejectRevision`                   | Use `Document.rejectRevision`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `Query.indexStyles`                      | Use `createStyleIndex`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `Query.resolveStyle`                     | Use `getParagraphStyle`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Query.resolveRunStyle`                  | Use `getRunStyle`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `Query.resolveNumberingLevel`            | Use `getNumberingLevel`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Query.resolveTableStyle`                | Use `getTableStyle`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Query.StyleIndex`                       | Use `DocxStyleIndex` from `documonster/word`. Removed in 1.0.                                                                                                                                                                                                                            |
| `Query.setCompatibilityMode`             | Use `Document.setCompatibilityMode`. Removed in 1.0.                                                                                                                                                                                                                                     |
| `Query.resolveDataBindings`              | Use `Transform.resolveDataBindings`. Removed in 1.0.                                                                                                                                                                                                                                     |
| `Query.extractFormFields`                | Use `listFormFields`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Query.fillFormFields`                   | Use `Transform.fillFormFields`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `Query.countByFormat`                    | Use `getCountByFormat`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `Query.getUsedFormats`                   | Use `listUsedFormats`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `Query.walkDocument`                     | Use `each`. Removed in 1.0.                                                                                                                                                                                                                                                              |
| `Query.walkBlocks`                       | Use `eachBlock`. Removed in 1.0.                                                                                                                                                                                                                                                         |
| `Query.collectParagraphs`                | Use `listParagraphs`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Query.collectRuns`                      | Use `listRuns`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `Query.collectTables`                    | Use `listTables`, passing the parts to include — `listTables` searches only the body unless asked. Removed in 1.0.                                                                                                                                                                       |
| `Io.package`                             | Use `Document.toBuffer`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Io.read`                                | Use `Document.read`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Io.toBuffer`                            | Use `Document.toBuffer`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Io.toBase64`                            | Use `Document.toBase64`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Io.readFile`                            | Use `Document.readFile`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Io.writeFile`                           | Use `Document.writeFile`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `Io.patchDocument`                       | Use `Package.patch`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Io.compileTemplate`                     | Use `Package.compileTemplate`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `Io.patchTemplate`                       | Use `Package.patchTemplate`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `Io.fillTemplateFromBuffer`              | Use `Package.fillTemplate`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `Io.toFlatOpcFromDoc`                    | Use `toFlatOpc`, which takes an options object. Removed in 1.0.                                                                                                                                                                                                                          |
| `Io.editDocxIncremental`                 | Use `Package.edit`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `Io.listDocxParts`                       | Use `Package.listParts`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Io.readDocxPart`                        | Use `Package.readPart`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `Io.updateFields`                        | Use `Transform.updateFields`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `Io.updateTableOfContents`               | Use `Transform.updateTableOfContents`. Removed in 1.0.                                                                                                                                                                                                                                   |
| `Io.merge`                               | Use `Transform.merge`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `Io.split`                               | Use `Transform.split`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `Template.fillTemplate`                  | Use `fill`, which returns a new document. Removed in 1.0.                                                                                                                                                                                                                                |
| `Template.fillTemplateEnhanced`          | Use `fill` with `{ richPlaceholders: true }`. Removed in 1.0.                                                                                                                                                                                                                            |
| `Template.listTemplateTags`              | Use `listTags`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `Template.isTemplateChart`               | Use `isChart`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `Template.TemplateError`                 | Import `DocxTemplateError` from `documonster/word` directly. Removed in 1.0.                                                                                                                                                                                                             |
| `Template.fillTemplateFromSource`        | Use `fillFromSource`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Convert.parseFlatOpc`                   | Use `flatOpcToParts` for the part map, or `Document.read` for the document a Flat OPC file holds. Removed in 1.0.                                                                                                                                                                        |
| `Convert.toFlatOpc`                      | Use `partsToFlatOpc` — this wraps a part map, not a document; `Document.toFlatOpc(doc)` is the document form. Removed in 1.0.                                                                                                                                                            |
| `Convert.readOdt`                        | Use `fromOdt`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `Convert.writeOdt`                       | Use `toOdt`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `Convert.docxToSemantic`                 | Use `toSemantic`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `Convert.ommlToMathML`                   | Use `ommlToMathMl`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `Convert.mathMLToOmml`                   | Use `mathMlToOmml`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `Convert.mapDocument`                    | Use `Transform.map`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Layout.document`                        | Use `getPageMap`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `Layout.documentFull`                    | Use `compute`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `Layout.renderPageToSvg`                 | Use `pageToSvg`. Removed in 1.0.                                                                                                                                                                                                                                                         |
| `Layout.renderDocumentToSvg`             | Use `toSvg`. Removed in 1.0.                                                                                                                                                                                                                                                             |
| `Layout.renderPageFromLayout`            | Use `computedPageToSvg`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `Security.encrypt`                       | Use `WordCrypto.encrypt` from `documonster/word/crypto`, beside `decrypt` and `verifyPassword`. Removed in 1.0.                                                                                                                                                                          |
| `Security.DEFAULT_SECURITY_POLICY`       | Use `DEFAULT_POLICY`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Security.STRICT_SECURITY_POLICY`        | Use `STRICT_POLICY`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Security.resolveSecurityPolicy`         | Use `resolvePolicy`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `Glossary.autoTextEntries`               | Use `listAutoTextEntries`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `Glossary.quickParts`                    | Use `listQuickParts`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Diff.documents`                         | Use `compare`. Removed in 1.0.                                                                                                                                                                                                                                                           |
| `Validation.document`                    | Use `validate`. Removed in 1.0.                                                                                                                                                                                                                                                          |
| `Streaming.StreamingDocxWriter`          | Use `DocumentWriter`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Streaming.StreamingDocxWriter#finalize` | Use `commit` to finish the document and `toBuffer` to obtain its bytes. Removed in 1.0.                                                                                                                                                                                                  |
| `Streaming.createDocxStream`             | Use `new Stream.DocumentWriter(options)`. Removed in 1.0.                                                                                                                                                                                                                                |
| `Streaming.StreamingDocxReader`          | Use `DocumentReader`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `Streaming.createDocxStreamReader`       | Use `new Stream.DocumentReader(source, options)`. Removed in 1.0.                                                                                                                                                                                                                        |
| `SearchResult`                           | Use `DocxSearchResult`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `DocumentHeading`                        | Use `DocxHeading`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `DocumentSection`                        | Use `DocxSectionInfo`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `MergeOptions`                           | Use `DocxMergeOptions`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `SplitOptions`                           | Use `DocxSplitOptions`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `RevisionEntry`                          | Use `DocxRevisionEntry`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `StyleResolveContext`                    | Use `DocxStyleResolveContext`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `ResolvedParagraphStyle`                 | Use `DocxResolvedParagraphStyle`. Removed in 1.0.                                                                                                                                                                                                                                        |
| `ResolvedRunStyle`                       | Use `DocxResolvedRunStyle`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `ResolvedNumberingLevel`                 | Use `DocxResolvedNumberingLevel`. Removed in 1.0.                                                                                                                                                                                                                                        |
| `StyleIndex`                             | Use `DocxStyleIndex`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `CompatibilityMode`                      | Use `DocxCompatibilityMode`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `FormFieldEntry`                         | Use `DocxFormFieldEntry`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `PatchContent`                           | Use `DocxPatchContent`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `PatchOperation`                         | Use `DocxPatchOperation`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `PatchOptions`                           | Use `DocxPatchOptions`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `CompiledTemplate`                       | Use `DocxCompiledTemplate`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `IncrementalEdit`                        | Use `DocxIncrementalEdit`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `IncrementalEditOptions`                 | Use `DocxIncrementalEditOptions`. Removed in 1.0.                                                                                                                                                                                                                                        |
| `ReadDocxOptions`                        | Use `DocxReadOptions`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `DiffChangeType`                         | Use `DocxDiffChangeType`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `DiffEntry`                              | Use `DocxDiffEntry`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `DiffSummary`                            | Use `DocxDiffSummary`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `DiffResult`                             | Use `DocxDiffResult`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `TemplateOptions`                        | Use `DocxTemplateOptions`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `TemplateImage`                          | Use `DocxTemplateImage`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `TemplateRichText`                       | Use `DocxTemplateRichText`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `TemplateSubDocument`                    | Use `DocxTemplateSubDocument`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `TemplateChart`                          | Use `DocxTemplateChart`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `TemplateHtmlChunk`                      | Use `DocxTemplateHtmlChunk`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `TemplateTag`                            | Use `DocxTemplateTag`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `FlatOpcPart`                            | Use `DocxFlatOpcPart`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `StreamingDocxOptions`                   | Use `DocxStreamingWriteOptions`. Removed in 1.0.                                                                                                                                                                                                                                         |
| `StreamingProgressCallback`              | Use `DocxStreamingProgressCallback`. Removed in 1.0.                                                                                                                                                                                                                                     |
| `StreamingReadOptions`                   | Use `DocxStreamingReadOptions`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `StreamingDocxMetadata`                  | Use `DocxStreamingMetadata`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `ValidationSeverity`                     | Use `DocxValidationSeverity`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `ValidationIssue`                        | Use `DocxValidationIssue`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `ValidationResult`                       | Use `DocxValidationResult`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `ValidationOptions`                      | Use `DocxValidationOptions`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `FontEmbedStyle`                         | Use `DocxFontEmbedStyle`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `EmbedFontOptions`                       | Use `DocxEmbedFontOptions`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `EmbedFontResult`                        | Use `DocxEmbedFontResult`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `OleObject`                              | Use `DocxOleObject`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `OleObjectType`                          | Use `DocxOleObjectType`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `OleDisplayAs`                           | Use `DocxOleDisplayAs`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `OleExtractionResult`                    | Use `DocxOleExtractionResult`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `OleEmbeddingResult`                     | Use `DocxOleEmbeddingResult`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `BuildingBlockGallery`                   | Use `DocxBuildingBlockGallery`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `BuildingBlock`                          | Use `DocxBuildingBlock`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `GlossaryDocument`                       | Use `DocxGlossaryDocument`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `ProtectionEditType`                     | Use `DocxProtectionEditType`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `ProtectionHashAlgorithm`                | Use `DocxProtectionHashAlgorithm`. Removed in 1.0.                                                                                                                                                                                                                                       |
| `DocumentProtectionOptions`              | Use `DocxProtectionOptions`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `ProtectionState`                        | Use `DocxProtectionState`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `MappingSourceType`                      | Use `DocxMappingSourceType`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `MappingCondition`                       | Use `DocxMappingCondition`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `MappingTarget`                          | Use `DocxMappingTarget`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `StyleMappingRule`                       | Use `DocxStyleMappingRule`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `StyleMap`                               | Use `DocxStyleMap`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `StyleMapOptions`                        | Use `DocxStyleMapOptions`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `VbaProjectInfo`                         | Use `DocxVbaProjectInfo`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `LayoutResult`                           | Use `DocxLayoutResult`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `LayoutOptions`                          | Use `DocxLayoutOptions`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `FullLayoutOptions`                      | Use `DocxFullLayoutOptions`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `PageGeometryOverride`                   | Use `DocxPageGeometryOverride`. Removed in 1.0.                                                                                                                                                                                                                                          |
| `FieldUpdateOptions`                     | Use `DocxFieldUpdateOptions`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `StandardShapeType`                      | Use `DocxStandardShapeType`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `SolidFill`                              | Use `DocxSolidFill`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `GradientStop`                           | Use `DocxGradientStop`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `GradientFill`                           | Use `DocxGradientFill`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `PatternFill`                            | Use `DocxPatternFill`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `NoFill`                                 | Use `DocxNoFill`. Removed in 1.0.                                                                                                                                                                                                                                                        |
| `ShapeFill`                              | Use `DocxShapeFill`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `LineDash`                               | Use `DocxLineDash`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `LineEndType`                            | Use `DocxLineEndType`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `LineEndSize`                            | Use `DocxLineEndSize`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `ShapeOutline`                           | Use `DocxShapeOutline`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `ShadowEffect`                           | Use `DocxShadowEffect`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `GlowEffect`                             | Use `DocxGlowEffect`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `ReflectionEffect`                       | Use `DocxReflectionEffect`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `Effect3D`                               | Use `DocxEffect3D`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `ShapeEffects`                           | Use `DocxShapeEffects`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `TextVerticalAnchor`                     | Use `DocxTextVerticalAnchor`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `TextWrap`                               | Use `DocxTextWrap`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `ShapeTextBody`                          | Use `DocxShapeTextBody`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `CreateShapeOptions`                     | Use `DocxCreateShapeOptions`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `CreateShapeOptions.flipH?`              | Use `flipHorizontal`, as the shape node names it. Removed in 1.0.                                                                                                                                                                                                                        |
| `CreateShapeOptions.flipV?`              | Use `flipVertical`, as the shape node names it. Removed in 1.0.                                                                                                                                                                                                                          |
| `ScriptType`                             | Use `DocxScriptType`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `BiDiDirection`                          | Use `DocxBiDiDirection`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `ShapedCluster`                          | Use `DocxShapedCluster`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `ShapingOptions`                         | Use `DocxShapingOptions`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `HyphenationOptions`                     | Use `DocxHyphenationOptions`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `HyphenationOptions.hyphenChar?`         | Pass `{ hyphenChar }` to `Font.hyphenateWord` or `Font.hyphenateText`. Removed in 1.0.                                                                                                                                                                                                   |
| `HyphenationPatterns`                    | Use `DocxHyphenationPatterns`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `FormatCriteria`                         | Use `DocxFormatCriteria`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `FormatSearchResult`                     | Use `DocxFormatSearchResult`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `RenderOptions`                          | Use `DocxRenderOptions`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `RenderOptions.backgroundColor?`         | Use `background`, which also takes an `{ r, g, b, a }` colour and wins over this one. Removed in 1.0.                                                                                                                                                                                    |
| `DataSource`                             | Use `DocxDataSource`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `ChartBinding`                           | Use `DocxChartBinding`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `ChartSeriesData`                        | Use `DocxChartSeriesData`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `ChartTemplateData`                      | Use `DocxChartTemplateData`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `EncryptOptions`                         | Use `DocxEncryptOptions`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `WordSecurityPolicy`                     | Use `DocxSecurityPolicy`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `WordRenderContext`                      | Use `DocxRenderContext`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `IdGenerators`                           | Use `DocxIdGenerators`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `MapOptions`                             | Use `DocxMapOptions`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `WalkPath`                               | Use `DocxWalkPath`. Removed in 1.0.                                                                                                                                                                                                                                                      |
| `WalkOptions`                            | Use `DocxWalkOptions`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `VisitAction`                            | Use `DocxVisitAction`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `LayoutDocument`                         | Use `DocxLayoutDocument`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `LayoutPage`                             | Use `DocxLayoutPage`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `PageGeometry`                           | Use `DocxPageGeometry`. Removed in 1.0.                                                                                                                                                                                                                                                  |
| `PageContent`                            | Use `DocxPageContent`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `LayoutParagraph`                        | Use `DocxLayoutParagraph`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `LayoutTable`                            | Use `DocxLayoutTable`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `LayoutTableCell`                        | Use `DocxLayoutTableCell`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `LayoutImage`                            | Use `DocxLayoutImage`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `LayoutFloat`                            | Use `DocxLayoutFloat`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `LayoutFloatWrap`                        | Use `DocxLayoutFloatWrap`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `LayoutTextBox`                          | Use `DocxLayoutTextBox`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `LayoutShape`                            | Use `DocxLayoutShape`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `LayoutChart`                            | Use `DocxLayoutChart`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `LayoutSdt`                              | Use `DocxLayoutSdt`. Removed in 1.0.                                                                                                                                                                                                                                                     |
| `LayoutMath`                             | Use `DocxLayoutMath`. Removed in 1.0.                                                                                                                                                                                                                                                    |
| `LayoutCheckBox`                         | Use `DocxLayoutCheckBox`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `LayoutTableOfContents`                  | Use `DocxLayoutTableOfContents`. Removed in 1.0.                                                                                                                                                                                                                                         |
| `LayoutAltChunk`                         | Use `DocxLayoutAltChunk`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `LayoutOpaqueDrawing`                    | Use `DocxLayoutOpaqueDrawing`. Removed in 1.0.                                                                                                                                                                                                                                           |
| `LineBox`                                | Use `DocxLineBox`. Removed in 1.0.                                                                                                                                                                                                                                                       |
| `LineBoxItem`                            | Use `DocxLineBoxItem`. Removed in 1.0.                                                                                                                                                                                                                                                   |
| `PositionedInlineImage`                  | Use `DocxPositionedInlineImage`. Removed in 1.0.                                                                                                                                                                                                                                         |
| `PositionedRun`                          | Use `DocxPositionedRun`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `SemanticDocument`                       | Use `DocxSemanticDocument`. Removed in 1.0.                                                                                                                                                                                                                                              |
| `SemanticBlock`                          | Use `DocxSemanticBlock`. Removed in 1.0.                                                                                                                                                                                                                                                 |
| `SemanticInline`                         | Use `DocxSemanticInline`. Removed in 1.0.                                                                                                                                                                                                                                                |
| `ConversionContext`                      | Use `DocxConversionContext`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `ConversionWarning`                      | Use `DocxConversionWarning`. Removed in 1.0.                                                                                                                                                                                                                                             |
| `ConversionAsset`                        | Use `DocxConversionAsset`. Removed in 1.0.                                                                                                                                                                                                                                               |
| `ResolvedFormatting`                     | Use `DocxResolvedFormatting`. Removed in 1.0.                                                                                                                                                                                                                                            |
| `AssetRef`                               | Reserved for a package-model API that was never added: no function takes or returns it. Removed in 1.0.                                                                                                                                                                                  |
| `MediaAsset`                             | Reserved for a package-model API that was never added: no function takes or returns it; a document's images are `ImageDef`s. Removed in 1.0.                                                                                                                                             |

**`documonster/word/crypto`**

| Deprecated                     | Replacement                                                                                                                         |
| ------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------- |
| `isEncryptedDocx`              | Use `WordCrypto.isEncrypted`. Removed in 1.0.                                                                                       |
| `verifyPassword`               | Use `WordCrypto.Agile.verifyPassword`, which takes `(info, password)` instead of `(password, info)`. Removed in 1.0.                |
| `decryptPackage`               | Use `WordCrypto.Agile.decryptPackage`, which takes the cap in its options (`{ maxDecryptedSize }`). Removed in 1.0.                 |
| `decryptDocx`                  | Use `WordCrypto.decrypt`. Removed in 1.0.                                                                                           |
| `encryptDocx`                  | Use `WordCrypto.encrypt`. Removed in 1.0.                                                                                           |
| `parseEncryptionInfoXml`       | Use `WordCrypto.Agile.parseEncryptionInfo`. Removed in 1.0.                                                                         |
| `deriveEncryptionKey`          | Use `WordCrypto.Agile.deriveKey`, which takes `(info, password, blockKey)` instead of `(password, info, blockKey)`. Removed in 1.0. |
| `AGILE_BLOCK_KEYS`             | Use `WordCrypto.Agile.BLOCK_KEYS`. Removed in 1.0.                                                                                  |
| `AgileEncryptionInfo`          | Use `DocxAgileEncryptionInfo`. Removed in 1.0.                                                                                      |
| `EncryptOptions`               | Use `DocxEncryptOptions`. Removed in 1.0.                                                                                           |
| `readCfb`                      | Use `WordCrypto.Cfb.read`. Removed in 1.0.                                                                                          |
| `writeCfb`                     | Use `WordCrypto.Cfb.write`. Removed in 1.0.                                                                                         |
| `CfbEntry`                     | Use `DocxCfbEntry`. Removed in 1.0.                                                                                                 |
| `hasDigitalSignatures`         | Use `WordCrypto.hasSignatures`, which takes the document or the package's part names (`Package.listParts(bytes)`). Removed in 1.0.  |
| `parseSignatureXml`            | Use `WordCrypto.parseSignature`. Removed in 1.0.                                                                                    |
| `extractSignatures`            | Use `WordCrypto.listSignatures`. Removed in 1.0.                                                                                    |
| `isWellFormedSignature`        | Use `WordCrypto.isWellFormedSignature`. Removed in 1.0.                                                                             |
| `DigitalSignatureInfo`         | Use `DocxDigitalSignatureInfo`. Removed in 1.0.                                                                                     |
| `DigitalSignatureInfo.signer?` | Use `claimedSigner`, which holds the same value: the name is what the signature claims, not a verified signer. Removed in 1.0.      |
| `deobfuscateFont`              | Use `Font.deobfuscate` from `documonster/word`. Removed in 1.0.                                                                     |
| `obfuscateFont`                | Use `Font.obfuscate` from `documonster/word`. Removed in 1.0.                                                                       |
| `generateFontKey`              | Use `Font.generateObfuscationKey` from `documonster/word`. Removed in 1.0.                                                          |

**`documonster/word/excel`**

| Deprecated                  | Replacement                                    |
| --------------------------- | ---------------------------------------------- |
| `extractTablesToExcel`      | Use `WordExcel.extractTables`. Removed in 1.0. |
| `renderWordChartSvg`        | Use `WordExcel.toChartSvg`. Removed in 1.0.    |
| `buildWordChartExXml`       | Use `WordExcel.toChartExXml`. Removed in 1.0.  |
| `wordChartToChartModel`     | Use `WordExcel.toChartModel`. Removed in 1.0.  |
| `generateChartEmbeddedXlsx` | Use `WordExcel.toChartXlsx`. Removed in 1.0.   |
| `ExcelToDocxOptions`        | Use `DocxFromExcelOptions`. Removed in 1.0.    |
| `WordChartExOptions`        | Use `DocxChartExOptions`. Removed in 1.0.      |

**`documonster/word/html`**

| Deprecated          | Replacement                                                                                                                                                   |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `renderToHtml`      | Use `WordHtml.fromDocx`. Removed in 1.0.                                                                                                                      |
| `htmlToDocxBody`    | Use `WordHtml.toDocxBody``(html).body`. Removed in 1.0.                                                                                                       |
| `htmlToDocx`        | Use `WordHtml.toDocxBody`, which also returns the styles and list numbering the body refers to, or `WordHtml.toDocx` for a complete document. Removed in 1.0. |
| `HtmlRenderOptions` | Use `DocxHtmlRenderOptions`. Removed in 1.0.                                                                                                                  |
| `HtmlRenderResult`  | Use `DocxHtmlRenderResult`. Removed in 1.0.                                                                                                                   |
| `HtmlImportOptions` | Use `DocxHtmlImportOptions`. Removed in 1.0.                                                                                                                  |
| `HtmlToDocxResult`  | Use `WordHtml.toDocxBody`, whose result is `DocxHtmlBodyResult`. Removed in 1.0.                                                                              |

**`documonster/word/markdown`**

| Deprecated              | Replacement                                      |
| ----------------------- | ------------------------------------------------ |
| `renderToMarkdown`      | Use `WordMarkdown.fromDocx`. Removed in 1.0.     |
| `markdownToDocx`        | Use `WordMarkdown.toDocx`. Removed in 1.0.       |
| `markdownToDocxBody`    | Use `WordMarkdown.toDocxBody`. Removed in 1.0.   |
| `MarkdownRenderOptions` | Use `DocxMarkdownRenderOptions`. Removed in 1.0. |
| `MarkdownImportOptions` | Use `DocxMarkdownImportOptions`. Removed in 1.0. |
| `MarkdownImageData`     | Use `DocxMarkdownImageData`. Removed in 1.0.     |
| `MarkdownBodyResult`    | Use `DocxMarkdownBodyResult`. Removed in 1.0.    |

**`documonster/xml`**

| Deprecated                | Replacement                                                                                                                               |
| ------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| `SaxTag`                  | Use `XmlSaxTag`. Removed in 1.0.                                                                                                          |
| `SaxEvent`                | Use `XmlSaxEvent`. Removed in 1.0.                                                                                                        |
| `SaxEventAny`             | Use `XmlSaxEventAny`. Removed in 1.0.                                                                                                     |
| `SaxHandlers`             | Use `XmlSaxHandlers`. Removed in 1.0.                                                                                                     |
| `SaxOptions`              | Use `XmlSaxOptions`. Removed in 1.0.                                                                                                      |
| `InvalidCharHandling`     | Use `XmlInvalidCharHandling`. Removed in 1.0.                                                                                             |
| `WritableTarget`          | Use `XmlWritableTarget`. Removed in 1.0.                                                                                                  |
| `ToPlainObjectOptions`    | Use `XmlToPlainObjectOptions`. Removed in 1.0.                                                                                            |
| `ParseXmlToObjectOptions` | Use `XmlParseToObjectOptions`. Removed in 1.0.                                                                                            |
| `Xml.validateXmlName`     | Use `validateName`. Removed in 1.0.                                                                                                       |
| `Xml.Writer#xml`          | Use `toString`. Removed in 1.0.                                                                                                           |
| `Xml.Writer#save`         | Use `snapshot`. Removed in 1.0.                                                                                                           |
| `Xml.StdDocAttributes`    | Use `DEFAULT_DECLARATION`. Removed in 1.0.                                                                                                |
| `Xml.parseSax`            | Use `parseSaxStream`. An incremental reader takes the `Stream` suffix, as `Csv.parseStream` and `Workbook.readStream` do. Removed in 1.0. |
| `Xml.saxStream`           | Use `feedSax`. Removed in 1.0.                                                                                                            |
| `Xml.findChildren`        | Use `listChildren`. Removed in 1.0.                                                                                                       |
| `Xml.textContent`         | Use `getText`. Removed in 1.0.                                                                                                            |
| `Xml.attr`                | Use `getAttr`. Removed in 1.0.                                                                                                            |
| `Xml.walk`                | Use `eachElement`. Removed in 1.0.                                                                                                        |

#### Deprecated call forms

The name stays; the call form described is removed in 1.0.

**`documonster/archive`**

| Name                                       | Use instead                                          |
| ------------------------------------------ | ---------------------------------------------------- |
| `Archive.RemoteZipReader#extract`          | Pass the password as `{ password }`. Removed in 1.0. |
| `Archive.RemoteZipReader#extractEntry`     | Pass the password as `{ password }`. Removed in 1.0. |
| `Archive.RemoteZipReader#extractAll`       | Pass the password as `{ password }`. Removed in 1.0. |
| `Archive.RemoteZipReader#extractMultiple`  | Pass the password as `{ password }`. Removed in 1.0. |
| `Archive.RemoteZipReader#forEach`          | Pass the password as `{ password }`. Removed in 1.0. |
| `Archive.RemoteZipReader#entriesGenerator` | Pass the password as `{ password }`. Removed in 1.0. |
| `Archive.RemoteZipReader#extractToStream`  | Pass the password as `{ password }`. Removed in 1.0. |
| `Archive.RemoteZipReader#verifyCrc32`      | Pass the password as `{ password }`. Removed in 1.0. |
| `ArchiveFile#readEntry`                    | Pass the password as `{ password }`. Removed in 1.0. |
| `ArchiveFile#readEntrySync`                | Pass the password as `{ password }`. Removed in 1.0. |
| `ArchiveFile#readAsText`                   | Pass the encoding as `{ encoding }`. Removed in 1.0. |
| `ArchiveFile#readAsTextSync`               | Pass the encoding as `{ encoding }`. Removed in 1.0. |

**`documonster/csv`**

| Name                          | Use instead                                                                          |
| ----------------------------- | ------------------------------------------------------------------------------------ |
| `Csv.detectDelimiter`         | Pass `{ quote, delimitersToGuess, comment }`. The positional form is removed in 1.0. |
| `Csv.detectLinebreak`         | Pass `{ quote }`. The positional form is removed in 1.0.                             |
| `Csv.applyDynamicTypingToRow` | Pass `castDate` as `{ castDate }`. Removed in 1.0.                                   |
| `Csv.formatNumber`            | Pass the separator as `{ decimalSeparator }`. Removed in 1.0.                        |
| `Csv.parseNumber`             | Pass the separator as `{ decimalSeparator }`. Removed in 1.0.                        |

**`documonster/draw`**

| Name                              | Use instead                                                                                                                      |
| --------------------------------- | -------------------------------------------------------------------------------------------------------------------------------- |
| `BasicRasterCanvas#strokeRect`    | Pass `{ color, width }` as one options object. The positional form is removed in 1.0.                                            |
| `BasicRasterCanvas#drawPolyline`  | Pass `{ color, width, dash, join, cap }` as one options object. The positional form is removed in 1.0.                           |
| `BasicRasterCanvas#drawLine`      | Pass `{ color, width }` as one options object, as `strokeRect` takes it. The positional form is removed in 1.0.                  |
| `BasicRasterCanvas#strokeEllipse` | Pass `{ color, width, dash, join, cap }` as one options object. The positional form is removed in 1.0.                           |
| `BasicRasterCanvas#strokeCircle`  | Pass `{ color, width }` as one options object, as `strokeEllipse` takes it. The positional form is removed in 1.0.               |
| `BasicRasterCanvas#drawText`      | Pass `{ fontSize, color, anchor, rotation, family, bold, italic }` as one options object. The positional form is removed in 1.0. |

**`documonster/excel`**

| Name                           | Use instead                                                                                                                                                                                                                                  |
| ------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Workbook.protect`             | Pass the password in the options object: `protect(wb, { password, ...options })`. Removed in 1.0.                                                                                                                                            |
| `Workbook.registerPerson`      | Pass the identity in the options object: `registerPerson(wb, name, { userId, providerId })`. Removed in 1.0.                                                                                                                                 |
| `Workbook.read`                | Use `read(data, options)`, which returns a new workbook. This form replaces everything `wb` already holds — its sheets, names and properties — with what it reads. Removed in 1.0.                                                           |
| `Workbook.readWithDiagnostics` | Use `readWithDiagnostics(data, options)`, which reads into a new workbook. This form replaces everything `wb` already holds. Removed in 1.0.                                                                                                 |
| `Workbook.readFile`            | Use `readFile(filename, options)`, which returns a new workbook. This form replaces everything `wb` already holds. Removed in 1.0.                                                                                                           |
| `Workbook.readStream`          | Use `readStream(stream, options)`, which returns a new workbook. This form replaces everything `wb` already holds. Removed in 1.0.                                                                                                           |
| `Worksheet.fillFormula`        | Pass the results and the share type in the options object: `fillFormula(ws, range, formula, { results, shareType })`. Removed in 1.0.                                                                                                        |
| `Worksheet.protect`            | Pass the password in the options object: `protect(ws, { password, ...options })`. Removed in 1.0.                                                                                                                                            |
| `Worksheet.freeze`             | Pass the counts in the options object: `freeze(ws, { columns, rows })`. Removed in 1.0.                                                                                                                                                      |
| `Worksheet.split`              | Pass the offsets in the options object: `split(ws, { x, y })`. Removed in 1.0.                                                                                                                                                               |
| `Worksheet.getValues`          | Without `base` the result is 1-based, while `Row.getValues` and `Column.getValues` default to 0 — so state it: `{ base: 1 }` keeps this result, `{ base: 0 }` indexes as `Row` and `Column` do. Calling without `options` is removed in 1.0. |
| `Cell.getValue`                | Pass the cell as `{ row, col }`: `Cell.getValue(ws, { row, col })`. Removed in 1.0.                                                                                                                                                          |
| `Cell.setValue`                | Pass the cell as `{ row, col }`: `Cell.setValue(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                       |
| `Cell.getDateParts`            | Pass the cell as `{ row, col }`: `Cell.getDateParts(ws, { row, col })`. Removed in 1.0.                                                                                                                                                      |
| `Cell.setDateParts`            | Pass the cell as `{ row, col }`: `Cell.setDateParts(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                   |
| `Cell.getDateKind`             | Pass the cell as `{ row, col }`: `Cell.getDateKind(ws, { row, col })`. Removed in 1.0.                                                                                                                                                       |
| `Cell.getTemporal`             | Pass the cell as `{ row, col }`: `Cell.getTemporal(ws, { row, col })`. Removed in 1.0.                                                                                                                                                       |
| `Cell.getText`                 | Pass the cell as `{ row, col }`: `Cell.getText(ws, { row, col })`. Removed in 1.0.                                                                                                                                                           |
| `Cell.getDisplayText`          | Pass the cell as `{ row, col }`: `Cell.getDisplayText(ws, { row, col })`. Removed in 1.0.                                                                                                                                                    |
| `Cell.getType`                 | Pass the cell as `{ row, col }`: `Cell.getType(ws, { row, col })`. Removed in 1.0.                                                                                                                                                           |
| `Cell.getEffectiveType`        | Pass the cell as `{ row, col }`: `Cell.getEffectiveType(ws, { row, col })`. Removed in 1.0.                                                                                                                                                  |
| `Cell.getFormula`              | Pass the cell as `{ row, col }`: `Cell.getFormula(ws, { row, col })`. Removed in 1.0.                                                                                                                                                        |
| `Cell.getResult`               | Pass the cell as `{ row, col }`: `Cell.getResult(ws, { row, col })`. Removed in 1.0.                                                                                                                                                         |
| `Cell.getStyle`                | Pass the cell as `{ row, col }`: `Cell.getStyle(ws, { row, col })`. Removed in 1.0.                                                                                                                                                          |
| `Cell.setStyle`                | Pass the cell as `{ row, col }`: `Cell.setStyle(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                       |
| `Cell.applyCellStyle`          | Pass the cell as `{ row, col }`: `Cell.applyStyle(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                     |
| `Cell.isMerged`                | Pass the cell as `{ row, col }`: `Cell.isMerged(ws, { row, col })`. Removed in 1.0.                                                                                                                                                          |
| `Cell.getHyperlink`            | Pass the cell as `{ row, col }`: `Cell.getHyperlink(ws, { row, col })`. Removed in 1.0.                                                                                                                                                      |
| `Cell.getNote`                 | Pass the cell as `{ row, col }`: `Cell.getNote(ws, { row, col })`. Removed in 1.0.                                                                                                                                                           |
| `Cell.setNote`                 | Pass the cell as `{ row, col }`: `Cell.setNote(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                        |
| `Cell.getNames`                | Pass the cell as `{ row, col }`: `Cell.getNames(ws, { row, col })`. Removed in 1.0.                                                                                                                                                          |
| `Cell.addName`                 | Pass the cell as `{ row, col }`: `Cell.addName(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                        |
| `Cell.removeName`              | Pass the cell as `{ row, col }`: `Cell.removeName(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                     |
| `Cell.setNames`                | Pass the cell as `{ row, col }`: `Cell.setNames(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                       |
| `Cell.removeAllNames`          | Pass the cell as `{ row, col }`: `Cell.clearNames(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                     |
| `Cell.getValidation`           | Pass the cell as `{ row, col }`: `Cell.getValidation(ws, { row, col })`. Removed in 1.0.                                                                                                                                                     |
| `Cell.setValidation`           | Pass the cell as `{ row, col }`: `Cell.setValidation(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                  |
| `Cell.getModel`                | Pass the cell as `{ row, col }`: `Cell.getModel(ws, { row, col })`. Removed in 1.0.                                                                                                                                                          |
| `Cell.setModel`                | Pass the cell as `{ row, col }`: `Cell.setModel(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                       |
| `Cell.getFont`                 | Pass the cell as `{ row, col }`: `Cell.getFont(ws, { row, col })`. Removed in 1.0.                                                                                                                                                           |
| `Cell.setFont`                 | Pass the cell as `{ row, col }`: `Cell.setFont(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                        |
| `Cell.getNumFmt`               | Pass the cell as `{ row, col }`: `Cell.getNumFmt(ws, { row, col })`. Removed in 1.0.                                                                                                                                                         |
| `Cell.setNumFmt`               | Pass the cell as `{ row, col }`: `Cell.setNumFmt(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                      |
| `Cell.getAlignment`            | Pass the cell as `{ row, col }`: `Cell.getAlignment(ws, { row, col })`. Removed in 1.0.                                                                                                                                                      |
| `Cell.setAlignment`            | Pass the cell as `{ row, col }`: `Cell.setAlignment(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                   |
| `Cell.getBorder`               | Pass the cell as `{ row, col }`: `Cell.getBorder(ws, { row, col })`. Removed in 1.0.                                                                                                                                                         |
| `Cell.setBorder`               | Pass the cell as `{ row, col }`: `Cell.setBorder(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                      |
| `Cell.getFill`                 | Pass the cell as `{ row, col }`: `Cell.getFill(ws, { row, col })`. Removed in 1.0.                                                                                                                                                           |
| `Cell.setFill`                 | Pass the cell as `{ row, col }`: `Cell.setFill(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                        |
| `Cell.getProtection`           | Pass the cell as `{ row, col }`: `Cell.getProtection(ws, { row, col })`. Removed in 1.0.                                                                                                                                                     |
| `Cell.setProtection`           | Pass the cell as `{ row, col }`: `Cell.setProtection(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                  |
| `Cell.setResult`               | Pass the cell as `{ row, col }`: `Cell.setResult(ws, { row, col }, …)`. Removed in 1.0.                                                                                                                                                      |
| `Cell.getFullAddress`          | Pass the cell as `{ row, col }`: `Cell.getFullAddress(ws, { row, col })`. Removed in 1.0.                                                                                                                                                    |
| `Table.addRow`                 | `options.commit` does nothing — every edit applies as it is made. Omit it. Removed in 1.0.                                                                                                                                                   |
| `Table.removeRows`             | `options.commit` does nothing — every edit applies as it is made. Omit it. Removed in 1.0.                                                                                                                                                   |
| `Note.create`                  | Pass the author in the options object: `create(note, { author })`. Removed in 1.0.                                                                                                                                                           |
| `Anchor.create`                | Pass the offset in the options object: `create(ws, address, { offset })`. Removed in 1.0.                                                                                                                                                    |
| `ChartOptionsError`            | Pass the field first: `new ExcelChartOptionsError(field, message, options)`. This form leaves `field` empty. Removed in 1.0.                                                                                                                 |

**`documonster/excel [browser]`**

| Name                  | Use instead                                                                                                                        |
| --------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `Workbook.readStream` | Use `readStream(stream, options)`, which returns a new workbook. This form replaces everything `wb` already holds. Removed in 1.0. |

**`documonster/pdf`**

| Name                       | Use instead                                                                                                                                                                                                                                                               |
| -------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Pdf.fromExcel`            | Pass sheet names. Numeric positions are deprecated (removed in 1.0): they are 1-based here but 0-based in `WordExcel.toDocx` and sheet ids in `Workbook.getWorksheet`. The form that takes numeric positions in `sheets`; it works exactly as before until it is removed. |
| `Pdf.Builder#embedFont`    | Pass `{ data: fontBytes, collectionIndex }` as one `PdfFontSource`. The positional `collectionIndex` is removed in 1.0.                                                                                                                                                   |
| `Pdf.Editor#splitPages`    | Pass `{ pageIndices }` as one options object. The positional form is removed in 1.0. It skips an index out of range or naming a removed page, as it always has, and throws `PdfStructureError` for one in range that is not an integer.                                   |
| `Pdf.Editor#embedFont`     | Pass `{ data: fontBytes, collectionIndex }` as one `PdfFontSource`. The positional `collectionIndex` is removed in 1.0.                                                                                                                                                   |
| `Pdf.Editor#copyPagesFrom` | Pass `{ pageIndices, password }` as one options object. The positional form is removed in 1.0.                                                                                                                                                                            |
| `PdfLimitExceededError`    | Pass the limit and its value first: `new PdfLimitExceededError(limit, allowed, message, options)`. This form leaves `limit` and `allowed` undefined. Removed in 1.0.                                                                                                      |

**`documonster/stream`**

| Name                   | Use instead                                                        |
| ---------------------- | ------------------------------------------------------------------ |
| `PullStream#pull`      | Use `pullUntil`, which takes `{ includePattern }`. Removed in 1.0. |
| `PullStream#pullUntil` | Pass `{ includePattern }`. Removed in 1.0.                         |

**`documonster/word`**

| Name                        | Use instead                                                                                                                  |
| --------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `Document.addImage`         | Pass the format and size in the options: `Document.addImage(doc, data, { type, width, height })`. Removed in 1.0.            |
| `Document.addFloatingImage` | Pass the format and size in the options: `Document.addFloatingImage(doc, data, { type, width, height, … })`. Removed in 1.0. |
| `Build.underline`           | Pass the style in the properties: `underline(content, { ...properties, underline: style })`. Removed in 1.0.                 |
| `Build.field`               | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.pageNumberField`     | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.totalPagesField`     | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.sectionPagesField`   | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.sectionField`        | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.dateField`           | Pass an options object: `dateField({ format, cachedValue })`. Removed in 1.0.                                                |
| `Build.timeField`           | Pass an options object: `timeField({ format, cachedValue })`. Removed in 1.0.                                                |
| `Build.authorField`         | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.titleField`          | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.subjectField`        | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.keywordsField`       | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.fileSizeField`       | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.quoteField`          | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.ifField`             | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.includePictureField` | Pass `{ cachedValue }`. Removed in 1.0.                                                                                      |
| `Build.mathFraction`        | Pass the type in options: `mathFraction(numerator, denominator, { fractionType })`. Removed in 1.0.                          |
| `Build.mathSum`             | Pass the limits in options: `mathSum(content, { sub, sup })`. Removed in 1.0.                                                |
| `Build.mathIntegral`        | Pass the limits in options: `mathIntegral(content, { sub, sup })`. Removed in 1.0.                                           |
| `Build.mathProduct`         | Pass the limits in options: `mathProduct(content, { sub, sup })`. Removed in 1.0.                                            |
| `Build.mathNary`            | Pass the limits in options: `mathNary(char, content, { sub, sup })`. Removed in 1.0.                                         |
| `Build.mathLimit`           | Pass the type in options: `mathLimit(base, limit, { limitType })`. Removed in 1.0.                                           |
| `Build.mathAccent`          | Pass the character in options: `mathAccent(content, { char })`. Removed in 1.0.                                              |
| `Build.mathBar`             | Pass the position in options: `mathBar(content, { position })`. Removed in 1.0.                                              |
| `Build.border`              | Pass an options object: `border({ style, size, color })`. Removed in 1.0.                                                    |
| `Build.gridBorders`         | Pass an options object: `gridBorders({ size, color })`. Removed in 1.0.                                                      |
| `Build.table`               | Pass the widths in the options: `table(rows, { ...properties, columnWidths })`. Removed in 1.0.                              |
| `Font.hyphenateWord`        | Pass the hyphen character as `{ hyphenChar }`. Removed in 1.0.                                                               |
| `Font.hyphenateText`        | Pass the hyphen character as `{ hyphenChar }`. Removed in 1.0.                                                               |

**`documonster/word/excel`**

| Name          | Use instead                                                                                                                                                                                                                                                            |
| ------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `excelToDocx` | Pass sheet names. Numeric positions are deprecated (removed in 1.0): they are 0-based here but 1-based in `Pdf.fromExcel` and sheet ids in `Workbook.getWorksheet`. The form that takes numeric positions in `sheets`; it works exactly as before until it is removed. |

#### Internal fields marked deprecated

These were published by accident. They are stripped from the types in 1.0 and have no replacement — read and
change the object through its namespace instead.

- `Workbook.Handle` (`documonster/excel`): `sourceFilePath?`, `workbookContentType?`, `_worksheets`, `_definedNames`, `_themes?`, `_defaultFont?`, `_cellStyles?`, `_dxfs?`, `_writerExternalLinkCache`, `_tableNames`, `_chartEntries`, `_chartRels`, `_chartStyles`, `_chartColors`, `_chartExStyles`, `_chartExColors`, `_chartExEntries`, `_chartExRels`, `_chartExStructuredEntries`, `_chartsheets`, `_persons`, `_slicerParts`, `_slicerCacheParts`, `_timelineParts`, `_timelineCacheParts`, `_opaqueParts`, `_xlsbPivotCaches?`, `_opaqueContentTypeDefaults`, `_opaqueDrops`, `userFunctions?`
- `Stream.WorkbookWriter` (`documonster/excel`): `sharedStrings`, `styles`, `zipOptions?`, `xlsbTables?`, `compressionLevel`, `media`, `commentRefs`, `dynamicArrayCount`, `zip`, `stream`, `promise`
- `Stream.CellHandle` (`documonster/excel`): `_value`, `_mergeCount`, `_comment?`, `_formulaGhostOwner?`
- `Stream.WorksheetWriter` (`documonster/excel`): `_name`, `_retainedComments`, `tables`, `pivotTables`, `threadedComments`, `formControls`, `_charts`, `_shapes`, `_sparklineGroups`, `_worksheetMcIgnorable`, `_sortStateAutoFilterRef`, `orderNo?`, `sheetNo`, `conditionalFormattings`, `hasComments`, `startedData`, `rId?`
- `Worksheet.Handle` (`documonster/excel`): `_workbook`, `_name`, `_rows`, `_columns`, `_keys`, `_merges`, `_autoFilterCriteria?`, `_sortStateXml?`, `_opaqueRels?`, `_chartsheetPlaceholder?`, `_retainedComments?`, `_xlsbDrawingRelationshipId?`, `_worksheetNamespaceAttributes?`, `_worksheetMcIgnorable?`, `_sortStateAutoFilterRef?`, `_media`, `_shapes`, `_charts`, `_sparklineGroups`, `_headerRowCount?`, `_drawing`, `_watermark`, `_watermarkMedia`, `_commitRow?`
- `Cell.Sheet` (`documonster/excel`): `_workbook`, `_name`, `_rows`, `_columns`, `_keys`, `_merges`, `_autoFilterCriteria?`, `_sortStateXml?`, `_opaqueRels?`, `_chartsheetPlaceholder?`, `_retainedComments?`, `_xlsbDrawingRelationshipId?`, `_worksheetNamespaceAttributes?`, `_worksheetMcIgnorable?`, `_sortStateAutoFilterRef?`, `_media`, `_shapes`, `_charts`, `_sparklineGroups`, `_headerRowCount?`, `_drawing`, `_watermark`, `_watermarkMedia`, `_commitRow?`
- `Row.Sheet` (`documonster/excel`): `_workbook`, `_name`, `_rows`, `_columns`, `_keys`, `_merges`, `_autoFilterCriteria?`, `_sortStateXml?`, `_opaqueRels?`, `_chartsheetPlaceholder?`, `_retainedComments?`, `_xlsbDrawingRelationshipId?`, `_worksheetNamespaceAttributes?`, `_worksheetMcIgnorable?`, `_sortStateAutoFilterRef?`, `_media`, `_shapes`, `_charts`, `_sparklineGroups`, `_headerRowCount?`, `_drawing`, `_watermark`, `_watermarkMedia`, `_commitRow?`
- `Column.Sheet` (`documonster/excel`): `_workbook`, `_name`, `_rows`, `_columns`, `_keys`, `_merges`, `_autoFilterCriteria?`, `_sortStateXml?`, `_opaqueRels?`, `_chartsheetPlaceholder?`, `_retainedComments?`, `_xlsbDrawingRelationshipId?`, `_worksheetNamespaceAttributes?`, `_worksheetMcIgnorable?`, `_sortStateAutoFilterRef?`, `_media`, `_shapes`, `_charts`, `_sparklineGroups`, `_headerRowCount?`, `_drawing`, `_watermark`, `_watermarkMedia`, `_commitRow?`
- `CellData` (`documonster/excel`): `_value`, `_mergeCount`, `_comment?`, `_formulaGhostOwner?`
- `CompiledTemplate` (`documonster/word`): `_doc`

<!-- api-changes:deprecations:end -->

## 0.21.1 and earlier

These changes shipped in releases up to and including 0.21.1. They are kept for anyone upgrading from an older
version. **Read [Unreleased](#unreleased) first**: where it changed something described here, it wins. The code below
is written with today's names — `createCsvReadStream` / `writeCsv` (`documonster/excel/csv`) are now
`ExcelCsv.toStream` / `ExcelCsv.format`, `CsvFormatterStream` is `Csv.FormatterStream`, and `Workbook.toBuffer`
resolves to a plain `Uint8Array` on every platform.

### Dates: `Temporal` accepted, the `date1904` epoch honored where it was dropped

Most of this is additive. Three items change existing behavior, and all three
are cases where the previous behavior produced a wrong file or a wrong read.

**New, non-breaking.** `Cell.setValue` now accepts `Temporal.PlainDate`,
`Temporal.PlainTime` and `Temporal.PlainDateTime`, and there are four new
readers/writers: `Cell.getDateParts`, `Cell.setDateParts`, `Cell.getDateKind`
and `Cell.getTemporal`. Only `CellValueInput` grew — `CellValue`, the type a
read returns, is unchanged, so no exhaustive `switch` over a cell value breaks
and `Cell.getValue` still returns a `Date`. The exported `PlainDate` /
`PlainTime` / `PlainDateTime` types are derived structurally from `globalThis`,
so a consumer whose `lib` predates `esnext.temporal` gets `never` rather than an
error in a declaration file they did not write. No polyfill is added; `Temporal`
needs Node 26+, Chrome 144+, Firefox 139+, Bun 1.4+ or Deno 2.7+, and
`Cell.getDateParts` / `Cell.setDateParts` are the runtime-independent
equivalent.

#### 1. The streaming XLSX writer now honors `date1904` (bug fix, output changes)

`Stream.WorkbookWriter`'s `date1904` option reached the BIFF12 row encoder but
neither the XLSX row encoder nor `xl/workbook.xml` — which was built from a
literal `properties: {}`. Both halves were missing, so the package was a
self-consistent **1900** workbook that round-tripped perfectly and was simply
not the workbook that had been asked for. `date1904: true` now writes 1904
serials and `<workbookPr date1904="1">`, so a cell's serial changes by 1,462.

Only affects code that passed `date1904: true` (or assigned
`writer.properties.date1904`) to a **streaming XLSX** writer. If you were
compensating for this by shifting your dates, remove the compensation.

#### 2. Date bounds in data validations now use the workbook's epoch (bug fix)

`DataValidation` rules of `type: "date"` had their bounds converted against a
hard-coded 1900 epoch on both read and write. In a 1904 workbook the bounds
therefore sat 1,462 days from the cells they constrained — a rule reading "on or
after 2020-01-15" rejected that very date. The read side had the mirror-image
omission, so a round trip through this library agreed with itself and hid it.

#### 3. Strict-format `t="d"` date cells are no longer parsed as local time (bug fix)

A strict-OpenXML date cell stores an ISO string with no timezone
(`2020-01-15T00:00:00`), and this was read with `new Date(...)`, which every JS
engine interprets as _local_. The cell therefore read a day early anywhere east
of UTC, and disagreed with every other read path in the library. It now goes
through `parseOoxmlDate`, like the rest.

#### 4. `Stream.WorkbookWriter.xlsbDate1904()` is renamed to `date1904Flag()`

The accessor is consulted by the XLSX writer as well now, and the old name said
otherwise — which is part of why the XLSX branch went unfixed. `date1904` is not
an XLSB concept. Rename the call if you were using it; the workbook-level
`date1904` option and `properties.date1904` are unchanged and are the supported
way to set it. `date1904Flag()` is itself deprecated, with the other sheet-writer hooks (see
[Contracts tightened across every module](#contracts-tightened-across-every-module-behaviour-change)).

#### 5. Date-validation bounds in XLSB, and an ISO `Z` in CSV (bug fixes)

Two more of the same shape, both found while checking the above:

- **XLSB** wrote `type: "date"` validation bounds against the 1900 epoch too, so a
  1904 workbook's bounds sat 1,462 days from its cells. Now threaded, like XLSX.
- **`2020-01-15T00:00:00.000Z` was parsed a day early** anywhere east of UTC. The
  library's fixed-width ISO parsers were written to be selected by
  length and so checked only their separators — `parseISO` looked for dashes at
  positions 4 and 7 and nothing else. Any caller reaching them another way got a
  silent prefix match, and both `createIsoDateParser` and `createDateParser` do,
  so the 24-character string was consumed by the 10- and 19-character parsers and
  built with `new Date(y, m - 1, d, …)` — local time. Each parser now checks its
  own width. Consequently a malformed or unexpected-width value that previously
  parsed by prefix now returns `null` and stays a string, which is the intended
  behaviour of `castDate`.
- The Excel CSV bridge's default read formats also did not include
  `YYYY-MM-DD[T]HH:mm:ss.SSSZ` — the form its own writer emits. Added.

**Excel CSV and Markdown dates are the cell's calendar value, on both sides
(behaviour change).** A date cell is wall-clock time — its serial is computed from
the `Date`'s UTC fields, and those are what Excel shows — so:

- `ExcelCsv` / `ExcelMarkdown` **writes** now default to `dateUTC: true`. With no
  `dateFormat`, a date is written as UTC ISO (`2024-03-10T00:00:00.000Z`) instead of
  ISO with the host's offset (`2024-03-09T19:00:00.000-05:00` in New York); with a
  `dateFormat`, its fields are the UTC ones (`2024-03-10`, where New York used to get
  `2024-03-09`). Output is now the same on every machine. Pass `dateUTC: false` for
  the old local output.
- **Reads** take a date with no timezone (`2024-03-10`, `2024-03-10T14:30:00`,
  `10/03/2024 14:30:00`) as that calendar value in UTC, instead of local time, so
  the cell — and the serial in the file — no longer depends on the machine. A text
  with `Z` or an offset is the instant it names, as before.

A file written by the previous default with no `dateFormat` still reads back
exactly (it carries its offset). One written with a `dateFormat` on a non-UTC
machine holds local fields, and now reads as those fields in UTC.

#### 6. Dates before 1900-03-01 now match Excel (behaviour change)

The serial converter placed serial 0 at 1899-12-30 and ran linearly, which is
correct from 1900-03-01 onward and one day out below it — because Excel spends
serial 60 on the fictitious 1900-02-29 and the linear model had no room for it.

| Expression        | Before | Now  | Excel |
| ----------------- | ------ | ---- | ----- |
| `DATE(1900,1,1)`  | 2      | 1    | 1     |
| `DATE(1900,2,28)` | 60     | 59   | 59    |
| `DATE(1900,2,29)` | 60     | 60   | 60    |
| `MONTH(60)`       | 3      | 2    | 2     |
| `DAY(60)`         | 28     | 29   | 29    |
| `YEAR(TRUE)`      | 1899   | 1900 | 1900  |

Serials from 61 onward — every date a real workbook holds — are bit-for-bit
unchanged, as is the whole 1904 system, which has no phantom day. Only dates in
January and February 1900 move.

`Cell.setDateParts` **refuses** 1900-02-29: a cell's value is stored as a `Date`
and no `Date` names that day, so accepting it would silently store March 1.
`Cell.getDateParts` on a cell read from a file whose serial is 60 likewise
reports 1900-03-01. The formula engine, which works on serials and never routes
a date through a `Date`, is exact.

#### 7. `Cell.getDateKind` gained `"duration"` and `"unknown"`

It previously answered `"dateTime"` for an unformatted cell, a `General` cell, a
plain number format _and_ an elapsed-time format such as `[h]:mm:ss` — dressing
"I cannot tell" and "this is a length of time" up as a calendar kind.
`Cell.getTemporal` now refuses a duration cell instead of manufacturing a
`PlainDateTime` in 1899 for it; pass an explicit kind to override.

Format detection was fixed alongside: `[hh]:mm` (and `[mm]`, `[ss]`) are
recognised as elapsed time, only the first `;` section is consulted, and `\d0`
is read as an escaped literal rather than a day code.

#### 8. `Cell.setDateParts` validates its input

`{}`, `{ month: 13 }`, `{ day: 0 }`, `{ year: 2024, month: 2, day: 31 }`,
`{ hour: 99 }` and `NaN` now throw. They previously carried the way `Date.UTC`
does — 2024-02-31 became March 2 in silence, and `NaN` wrote a workbook Excel
cannot open from a call that reported success. Internal month arithmetic
(`EDATE`, coupon stepping) still relies on carrying and is unaffected.

#### A note on `new Date(y, m, d)`

Not a change, but the reason for the above. A spreadsheet date is a _calendar_
value and a `Date` is an _instant_, so this library reads a `Date`'s **UTC**
fields as the cell's value. `new Date(2024, 0, 15)` therefore stores a different
serial in every timezone — in UTC+8 it becomes 2024-01-14 16:00. Pass
`new Date(Date.UTC(2024, 0, 15))`, or better, a value that says what it is:

```ts
import { Cell } from "documonster/excel";

Cell.setDateParts(ws, "A1", { year: 2024, month: 1, day: 15 });
```

Where the runtime has `Temporal` (Node 26, Chrome 144), `Cell.setValue(ws, "A1", Temporal.PlainDate.from("2024-01-15"))`
works too. TypeScript's own `lib` does not declare `Temporal` yet, so in TypeScript that line needs the types of a
Temporal polyfill; `Cell.setDateParts` needs nothing.

### `Pdf.fromExcel` now honors the worksheet's print settings (behavior change)

The PDF bridge forwarded only a handful of `pageSetup` fields, so most of Excel's
Page Setup dialog had no effect on the output — and in three places the exporter
did something Excel does not. Everything below is now wired end to end, with an
explicit export option that overrides the sheet, resolved **per sheet** so sheets
in one workbook can differ.

Four changes alter the default output of existing code:

| Setting             | Before                                                 | Now                                                                  |
| ------------------- | ------------------------------------------------------ | -------------------------------------------------------------------- |
| Center on page      | Narrow sheets were **always** centered horizontally    | Starts at the left margin; opt in with `horizontalCentered`          |
| Page order          | Effectively `overThenDown`                             | `downThenOver`, Excel's default; override with `pageOrder`           |
| `pageSetup.scale`   | Silently dropped whenever `fitToPage` was on (default) | Honored when the sheet is in "Adjust to N %" mode                    |
| `repeatRows: false` | Could not suppress `printTitlesRow`                    | An explicit `false` now suppresses it (`undefined` still falls back) |

To keep the previous appearance:

```ts
import { Pdf } from "documonster/pdf";

await Pdf.fromExcel(wb, {
  horizontalCentered: true,
  pageOrder: "overThenDown",
  scale: 1 // ignore the sheet's own percentage
});
```

Newly honored settings (previously ignored entirely, so nothing regresses — the
output only changes for sheets that had already set them):

- `verticalCentered` — center on page, vertical axis
- `fitToPage` + `fitToWidth` / `fitToHeight` — "Fit to N pages wide by M tall".
  Like Excel, these only ever shrink; a grid that already fits is not enlarged.
  Both are solved against the real paginator, and exact for wrapped cells too
  (the wrap calculation scales width, padding and font size together).
- `printTitlesColumn` — "Columns to repeat at left", the counterpart of
  `printTitlesRow`. Both are now treated as **absolute** bands independent of the
  print area, so a sheet printing `E1:T3` still repeats `A:B` at the left of every
  page. Previously they were read as "the first N rows/columns of the printed
  range", which repeated the wrong tracks whenever a print area was set, and
  `printTitlesRow: "3:5"` was misread as "repeat the first five rows". A band
  inside the printed range still repeats, but is not hoisted to the leading edge,
  since that would reorder the first page's grid.
- `showRowColHeaders` — print row numbers and column letters
- `blackAndWhite` — grayscale both vector and raster content. Vector colors
  (cell text, fills, borders, gridlines, heading bands, chart vectors,
  `&K`-colored header/footer runs, text watermarks) are converted up-front, with
  opacity preserved. Raster payloads are converted per pixel when the image
  XObject is written: PNG samples collapse to one `/DeviceGray` component, JPEG
  keeps its `DCTDecode` data and gains a `/DeviceN` luma color space. Transparent
  PNG regions stay transparent — no blend overlay is painted.
- `draft` — omit images and charts. A chartsheet still emits its page, blank,
  so page numbering is unaffected.
- `errors` — print error cells as `"displayed"` / `"blank"` / `"dash"` / `"NA"`

**Type change (internal surface).** `ResolvedPdfOptions.repeatRows` /
`repeatCols` are now `PdfRepeatBand | false` (`{ first, last }`, absolute
1-based) instead of `number | false`. `ResolvedPdfOptions` is _not_ exported from
`documonster/pdf`, so this only affects code that reaches into
the package's internal layout engine to call `layoutSheet` with a hand-built options
object. The public `PdfExportOptions` inputs are unchanged and still take a count.

`PdfPageOrder` and `PdfCellErrorMode` are now exported from `documonster/pdf`, so
the new `pageOrder` / `errors` options can be named in user code.

**On `fitToPage`.** `PdfExportOptions.fitToPage` (default `true`) collided by
name with `pageSetup.fitToPage` but means something different: it is
documonster's own "shrink to one page wide" fallback. It now applies only when
neither the caller nor the sheet expresses a scaling intent, so a sheet asking
for 80% is no longer shrunk twice. Explicitly passing `fitToPage` still wins.

- `cellComments` — print cell comments and notes, in either of Excel's modes.
  `"atEnd"` appends a list of every comment (classic notes and threaded comments
  alike) after the sheet's pages; `"asDisplayed"` draws each one as a box where it
  sits on the sheet, decoding the note's VML anchor, plus the red corner marker on
  the commented cell.

**Not implemented.** `horizontalDpi` / `verticalDpi` are meaningless for vector
output.

### Charts with an EMU extent are no longer 4/3 too large (behavior change)

A chart anchored with an explicit `ext` (`cx`/`cy`, which Excel stores in EMU)
was converted EMU→**pixels** (÷9525) and the result used directly as PDF points.
Every such chart rendered 33% oversized — a 4in × 2in chart came out
384pt × 192pt instead of 288pt × 144pt — which also overflowed the content area
and distorted the aspect ratio against surrounding cells. The conversion is now
EMU→points (÷12700).

Charts sized from a `br` anchor (the common case when Excel writes a
two-cell anchor) were always correct and are unaffected.

**Also fixed:** `fitToPage` (default `true`) promises "shrink to one page wide",
but the ratio ignored `scale`, so `{ scale: 2 }` overflowed to two pages while
`{ scale: 0.8 }` compounded the two factors and under-filled the page. It now
behaves as `min(scale, contentWidth / tableWidth)`, matching `fitToWidth: 1`.

**Also fixed:** a chart fill and stroke with _different_ opacities collapsed to a
single value, because the alpha `ExtGState` sets `/ca` and `/CA` together and
`fillAndStroke()` paints both under one graphics state. Such paints are now split
into two passes, each with its own state.

### `createCsvReadStream` is now a byte stream (breaking)

`createCsvReadStream(workbook, options?)` (today `ExcelCsv.toStream`) from `documonster/excel/csv` used to
return an **object-mode** stream that emitted `string` chunks. It now returns a
normal **byte** stream: `objectMode` is off on the readable side and every chunk
is a `Uint8Array` (a `Buffer` on Node).

The declared return type changed from `IReadable<Uint8Array | string>` to
`IReadable<Uint8Array>`. On the Node entry it is narrowed further, to
`IReadable<Uint8Array> & Readable`, so the value is a nominal `stream.Readable`:

```ts
import { createWriteStream } from "node:fs";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { ExcelCsv } from "documonster/excel/csv";

// Node — no adapter, no cast
await pipeline(ExcelCsv.toStream(wb), createWriteStream("out.csv"));
const web = Readable.toWeb(ExcelCsv.toStream(wb));
```

The writable half of the underlying `Transform` is deliberately **not** exposed.
The stream is already attached to a producer, so `write()` would inject rows into
the CSV; `write()` / `end()` therefore do not typecheck on the returned value.

**Why.** The old shape leaked the CSV _formatter's_ internals into a public API.
It also made the stream indistinguishable by type from `Workbook.toStream()`
while behaving completely differently, so anything that consumes bytes silently
misbehaved instead of failing to compile:

```ts
import { Readable } from "node:stream";
import { ExcelCsv } from "documonster/excel/csv";

const chunks: Uint8Array[] = [];
for await (const chunk of ExcelCsv.toStream(wb)) {
  chunks.push(chunk);
}
// Before: threw `list[0] argument must be an instance of Buffer or Uint8Array`
Buffer.concat(chunks);

// Before: produced a ReadableStream of *strings*, which an SDK then rejected
// or uploaded as "[object Object]"-shaped garbage.
Readable.toWeb(ExcelCsv.toStream(wb));
```

Both now work. The CSV content is byte-for-byte identical to before — only the
chunk type changed.

**If you relied on string chunks**, decode explicitly, or use the string API:

```ts
import { ExcelCsv } from "documonster/excel/csv";

// A chunk boundary can fall inside a multi-byte character, so decode with a
// streaming decoder rather than per chunk.
const decoder = new TextDecoder();
let csv = "";
for await (const chunk of ExcelCsv.toStream(wb)) {
  csv += decoder.decode(chunk, { stream: true });
}
csv += decoder.decode();
```

Or skip streaming entirely when the whole document fits in memory:

```ts
import { ExcelCsv } from "documonster/excel/csv";

const csv = ExcelCsv.format(wb);
```

Building a string by `+=`-ing chunks is the one pattern that breaks **silently**
— `"" + new Uint8Array([97])` yields `"97"`, not `"a"`. Search for that shape.

`Csv.FormatterStream` from `documonster/csv` (then the flat `CsvFormatterStream`) is **unchanged** and still emits
strings by default. It gained a `readableObjectMode` option to opt into bytes:

```ts
import { Csv } from "documonster/csv";

new Csv.FormatterStream({ readableObjectMode: false }); // rows in, bytes out
```

### Node IO declarations now match the runtime (non-breaking)

Two Node-entry signatures were narrowed to the types Node already returned.
Both are strict narrowings, so existing code keeps compiling — but the wrappers
and casts it needed can be deleted.

**`Workbook.toBuffer` returned `Promise<Buffer>`** on `documonster/excel` in 0.21.1. **Superseded:** it now resolves
to a plain `Uint8Array` on every platform (see [How to upgrade from 0.21.1](#how-to-upgrade-from-0211), step 4). Pass
the bytes to anything that takes a `Uint8Array`; where an API insists on a `Buffer`, wrap without copying:

```ts
import { Workbook } from "documonster/excel";

const bytes = await Workbook.toBuffer(wb);
const buf = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength); // a view, not a copy
```

**`Workbook.toStream` returns a nominal `stream.Readable` on Node.** The shared
`XlsxReadable` type is unchanged and **identical in both builds** — it is still
`IReadable<Uint8Array>`, so cross-platform code written against that name keeps
compiling as-is. Node's `toStream` refines only its _return type_ to
`XlsxReadable & Readable`, which is assignable to `XlsxReadable`:

Before, callers needed an adapter (`Readable.from(Workbook.toStream(wb))`) or a cast
(`Workbook.toStream(wb) as unknown as Readable`). Now:

```ts
import { Readable } from "node:stream";
import { Workbook } from "documonster/excel";
import type { XlsxReadable } from "documonster/excel";

const body: Readable = Workbook.toStream(wb);
const web = Readable.toWeb(Workbook.toStream(wb));

// Still valid, and still means the same type on both platforms
const portable: XlsxReadable = Workbook.toStream(wb);
```

`for await` continues to yield `Uint8Array`.

### Note: leaving `for await` early (no change, previously undocumented)

Not a change in `documonster`, but it was never written down and it bites the
common "read the first N chunks" shape.

On **Node**, breaking out of a `for await` loop lets Node's own async iterator
destroy the stream with an `AbortError`, so `stream.errored` is set and any
attached `'error'` listener receives `The operation was aborted`. Calling
`destroy()` yourself destroys it with no error. Nothing crashes either way — the
iterator handles the event — but code that treats an `'error'` or a non-null
`errored` as a serialization failure will report a false one.

```ts
import { Workbook } from "documonster/excel";

// Sets errored to AbortError
for await (const chunk of Workbook.toStream(wb)) break;
```

```ts
import { Workbook } from "documonster/excel";

// Silent: destroy first, then break
const stream = Workbook.toStream(wb);
for await (const chunk of stream) {
  stream.destroy();
  break;
}
```

```ts
import { Workbook } from "documonster/excel";

// Silent: opt out of the iterator's teardown — but then you must destroy it,
// or the serializer stays parked
const stream = Workbook.toStream(wb);
try {
  for await (const chunk of stream.iterator({ destroyOnReturn: false })) break;
} finally {
  stream.destroy();
}
```

The **browser** build destroys silently on early exit, so `errored` is not a
portable way to distinguish "the consumer stopped" from "serialization failed".
Track that yourself if the same code runs on both platforms.

The same applies to `ExcelCsv.toStream` on Node, which is also a
`stream.Readable`.
