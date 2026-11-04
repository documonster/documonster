# Migration Guide

Breaking changes for `documonster` will be documented here as they occur.

## Unreleased

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
  take `password` — the option name Word's `Io.read` uses — for XLSX and XLSB, the decrypted package's format detected
  as usual. Without it: `ExcelEncryptedError` (now pointing at the option). With a wrong one: the new
  `ExcelDecryptionError` (an `ExcelFileError`) with `reason: "password"`; other reasons are `"unsupported"`,
  `"malformed"`, `"integrity"`, `"limit"`. The decrypted package is bounded by `zip.maxTotalUncompressedSize`
  (default 2 GiB). An encrypted workbook cannot be streamed — it is one encrypted stream — so the streamed readers
  decrypt it whole first.
- **Both write them**: Excel `Workbook.toBuffer`, `writeFile`, `writeStream`, `toStream` (XLSX and XLSB) and Word
  `Io.toBuffer`, `toBase64`, `writeFile` take `password` (1–255 characters) and `encryption` (`EncryptOptions`:
  `keyBits`, `hashAlgorithm`, `spinCount` — the type `WordCrypto.encrypt` takes, now exported by
  `documonster/excel` too). Defaults are Office's: AES-256, SHA-512, 100,000 iterations. Excel 16 and Word 16 open the
  output with the password (AES-128/192/256, SHA-256/512 checked). **Word D7 is reversed: `Io.toBuffer` encrypts when
  given a password**; it wrote an unencrypted package and the TSDoc sent callers to `WordCrypto.encrypt`, which keeps
  working unchanged. `Package.*` keep `password` as a read option: their result is not encrypted. The encryption code
  loads only when a password is given (`import()`), so a consumer that never encrypts does not bundle it; there is no
  synchronous form.
- **Only Agile Encryption** (version 4.4) is implemented, in both modules. Standard Encryption (Office 2007), RC4
  CryptoAPI and Extensible Encryption are refused with `reason: "unsupported"`, the message naming the scheme.

Bug fixes in what Word already had (`WordCrypto.*`, `Io.read` with `password`):

- **Files Office encrypted decrypt.** `WordCrypto.decrypt`, `Io.read`, `Streaming.createReader` and
  `WordCrypto.Agile.decryptPackage` verified the password of a Word-, Excel- or third-party-encrypted file and then
  returned garbage (reported as a corrupt package): the package segments were keyed with the password key encryptor's
  salt instead of the key data's. Only files this library wrote, which used one salt for both, decrypted. It now writes
  two salts, as Office does. `AgileEncryptionInfo` gains `keyData` (the key data's salt, hash, sizes), and
  `dataIntegritySalt` / `dataIntegrityHmac` — declared but never filled — hold the file's `encryptedHmacKey` /
  `encryptedHmacValue`.
- **The data-integrity HMAC is checked** on decryption, as Office does: an altered package is a
  `DocxDecryptionError` (`reason: "integrity"`), not a silently damaged document.
- `DocxDecryptionError.reason` (new, optional): the same values as `ExcelDecryptionError.reason`.
- **A compound file that is not an encrypted package** (a legacy `.doc`) read by `Io.read` / `Streaming.createReader` is a
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
  model replaces the sheet's merges (applying a sheet's own model twice threw `MergeConflictError`) and frees the
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

- Splice and duplicate counts: `Row.splice` / `Column.splice` (and `Row.insert*`, `Row.duplicate`, the deprecated
  `Worksheet.splice*`) refuse a negative, fractional or `NaN` count with `ExcelArgumentError`, and a start outside the
  grid with `RowOutOfBoundsError` / `ColumnOutOfBoundsError`. A negative count used to duplicate a row and shift every
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
- `Cell.addName` / `setName` / `setNames`: `DefinedNameError` for a name Excel would repair (`A1`, `1abc`,
  `has space`) or one a table uses. `setNames` checks every name before removing any.
- `Address.*` with coordinates outside the grid: `encodeCol(1.5)` returned `undefined`, `encodeRow(-5)` `"-4"`,
  `decodeCell("zz")` `{ c: 701, r: null }`. Any string that is not a reference is an `InvalidAddressError` (`"garbage"`
  was a `ColumnOutOfBoundsError`); a well-formed reference past XFD stays `ColumnOutOfBoundsError`.
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
  and the stream setters, is an `InvalidValueTypeError`: rich text that is not a list of `{ text, font? }`
  (`{ richText: 42 }`), an empty or non-text `formula` /
  `sharedFormula`, a `result` that is not a scalar, `Date` or `{ error }`, a non-text `hyperlink` / `text` /
  `tooltip`, an `error` that is not an Excel error literal (`{ error: "#BOGUS" }`), a non-boolean `checkbox`, and a
  shared formula whose master is not a formula cell above or to the left. An object of none of those shapes is still
  stored as its JSON text. `InvalidValueTypeError` now extends `ExcelArgumentError` and has a `field`;
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
- `DefinedNames.setModel` with two entries that are one name in one scope (`N` and `n`): `DefinedNameError`. One of
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
  `InvalidAddressError`, for every `Cell` reader and writer, `setDateParts`, `getTemporal`, `applyStyle` and
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
  unbounded or unreadable `range` (`InvalidAddressError`) and an unknown `header`.
- A wrong handle to any function of `Workbook`, `Worksheet`, `Cell`, `Row`, `Column`, `DefinedNames`, `Note` and
  `Range` (field `wb`, `ws`, `dn`, `n`, `r`); `Worksheet.getName(42)` returned `undefined`.

**New names and deprecations**

- `DefinedNames.remove(dn, name, { localSheetId })` deletes a whole name (reference, formula or opaque); the
  `(dn, location, name)` form is unchanged. `DefinedNames.get(dn, name, { localSheetId })`, and the type
  `DefinedNames.NameScopeOptions`.
- `Workbook.listPersons(wb)`, `Workbook.listExternalLinks(wb)`.
- `ExcelArgumentError.field` names the path from the offending parameter. Types `AddExternalLinkOptions`,
  `RegisterFunctionOptions`, `UnrecognizedValue`, `BaseErrorOptions`, `FrozenPanes`, `SplitPanes`, `WorksheetPanes`,
  `WorkbookStreamInput`, `ExcelFileErrorOptions`, `InvalidValueTypeErrorOptions`; error class `XlsxParseError`.
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
- **Positions are argument errors.** `RowOutOfBoundsError`, `ColumnOutOfBoundsError` and `InvalidAddressError` extend
  `ExcelArgumentError` (still `ExcelError`s; `instanceof` either class and `isExcelError` hold as before) and carry
  `field` — the parameter (`row`, `col`, `addr`; `cstr`, `cell.r`, `range`, `cells`, `data`, `args`, … where the call
  says) — and the `Address.*`, `Range.*`, `Worksheet.merge` / `addAoa` / `addJson` and argument-check messages now start
  with the call (`Address.encodeCol: Column -1 is out of bounds: …`, where the call used to sit in the details). The
  constructors keep their positional forms and take `{ field, where }` in their options (`PositionErrorOptions`). A
  `catch (e instanceof ExcelArgumentError)` written for bad input now catches a bad position too.
- **A bad position in a file is the reader's parse error, not an argument error** (class change on read paths). A
  position the content puts outside the grid or not where it is written — `<c r="B7">` inside row 5, column `XFE`,
  `<row r="0">`, a merge from `C0` — is an `XlsxParseError` naming the part (`xl/worksheets/sheet1.xml: …`) from
  `Workbook.read` / `readFile` / `readStream` / `readWithDiagnostics` and `Stream.WorkbookReader`, with the position
  error as `cause`; it was the bare `InvalidAddressError` / `ColumnOutOfBoundsError` / `RowOutOfBoundsError`, and an
  out-of-range `<row r>` a plain `ExcelError` ("Invalid row number"). A CSV or Markdown row a worksheet cannot hold
  (more than 16 384 fields) is an `ExcelFileError` naming the row (`row 2 does not fit a worksheet: …`, `cause` the
  `ColumnOutOfBoundsError`) from `ExcelCsv.read` / `addWorksheet` / `readFile` / `createImportStream` and
  `ExcelMarkdown.read` / `addWorksheet(s)` / `readFile`. XLSB already reported these as `XlsbParseError`; a binary
  sheet read by `Stream.WorkbookReader` does too. The same positions given at an API call stay argument errors.
- `Stream.WorkbookReader`'s `maxEntrySize` refusal (`FileTooLargeError`) carries `allowed` — the configured limit — as
  the archive's own does; it was `NaN`.
- **Column setters check the column first.** `Column.setWidth` / `setHidden` / `setOutlineLevel` / `setKey` /
  `setHeader` / `setStyle` / `setNumFmt` / the facet setters / `addPageBreak` / `getStyle` and
  `WorksheetWriter#getColumn` refuse a number that is not a whole column from 1 to 16 384 with
  `ColumnOutOfBoundsError` before anything is created, as the getters do. `0` or `1.5` was a raw `TypeError`; `16385`
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
  and the deprecated `Form.vmlAnchor` / `vmlStyle` / `vmlCheckedValue` to `Form.getVml*`, which also take the handle.
- `Stream.WorkbookReader`'s `parse(input, options)` / `read(input, options)` apply the defaults to what the call
  leaves out. `{ styles: "cache" }` dropped `worksheets: "emit"`, so no sheet was read.
- A committed `WorksheetWriter`'s `findRow` / `eachRow` return nothing instead of throwing; `mergeCells("A1")` does
  nothing.
- An `ExcelCsv` read parses its input before adding the sheet, so input the parser refuses leaves the workbook as it
  was.

**Now refused**

- A wrong handle to any function of `Table`, `Chart`, `Chartsheet`, `Image`, `Anchor`, `Form`, `Pivot`, `Sparkline`,
  `DataValidation`, `ConditionalFormatting`, `HeaderFooterImage` and the `Stream.*` handle functions:
  `ExcelArgumentError`, `field` naming it (`ws`, `wb`, `table`, `column`, `chart`, `chartsheet`, `anchor`,
  `checkbox`, `image`, `row`, `cell`). `Image.add(42)` was an `ImageError` and `Table.setName(42, …)` a
  `TableError`; others failed inside or returned `false` / `undefined`. `Watermark.get` checks its handle too.
- A removal target that names nothing — a string, a fractional number, `null` — in `Chart.remove`, `Sparkline.remove`,
  `HeaderFooterImage.remove` (also a section outside `LH` … `RF`), `Image.remove`, `Image.removeShape`,
  `Pivot.remove`, `Form.remove` and `ConditionalFormatting.remove` (`ExcelArgumentError`); they answered
  `false`. An out-of-range position or another sheet's item still answers `false`.
- A sheet-qualified reference where the target is the sheet's own cell or range (`InvalidAddressError`):
  `Table.add` / `setModel` / `setRef` `ref`, `DataValidation.add` / `find` / `remove`, `Pivot.add` `ref`,
  `Anchor.create`, `WorksheetWriter#mergeCells`. The qualifier was dropped or misread. `Anchor.setCol` / `setRow` /
  `setModel` / `create` refuse values off the grid.
- Tables (`TableError` unless noted):
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
    `base64`, `filename` nor `link`, or whose `buffer` is not bytes (`ImageError`). These failed at write time, or
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
  edits (it was a `ChartOptionsError`). The table and rows option builders and `createSeriesFromColumns` refuse a
  missing or non-object `options`, or rows that are not an array, with `ExcelArgumentError` from every entry
  (`Chart.*`, `ChartClassic.*`, `ChartEx.*`, the deprecated flat names), as do `Chart.addSeries(c, null)` and
  `Chart.setUserShapesXml(c, 42)`; `ChartClassic.toSvg({})` and the ChartEx renderers throw `ChartOptionsError`.
  `Chart.toPng` (classic and Office 2016) checks `dpi` and `scale` itself and throws `ChartOptionsError` for an
  invalid value, in Node and in browsers alike; Node surfaced an `ArchiveArgumentError` from the PNG encoder, browsers
  ignored the value, and the Office 2016 path did not check `scale`.
  `Chart.applyPreset` with an unknown preset names both preset families. `Chartsheet.setName` applies the
  31-character rule (`WorksheetNameError`).
- `Chart.*` argument errors name the call the caller wrote — `Chart.createOptionsFromRows: …`, `Chart.place: …` —
  rather than the shared builder that found the mistake; class and `field` are unchanged. `Chart.mutate` /
  `mutateEx` refuse a `mutator` that is not a function (field `mutator`; a `TypeError` from inside) and non-boolean
  options; `Chart.copyTo` / `clone` check `range` up front (a malformed one was a `TypeError`). `Chart.setTitle`
  refuses with field `value` (was `title`) and says `Chart.setTitle: value must be …` (it named the deprecated
  `Chart.title`); `Chart.setLegend`'s field is `value` (was `legend`). `ChartOptionsError` is an
  `ExcelArgumentError` with a `field`, and the chart renderers check their options; see PDF, drawing and Mermaid
  below.
- Pivot tables:
  - Removing a pivot table's source _(owner decision)_ is a `PivotTableError`: its worksheet
    (`Workbook.removeWorksheet`, `Worksheet.destroy`), its table (`Table.remove`), or a sheet holding that table. This
    holds for pivots read from files too, by the cache's `worksheetSource`. A pivot table on the sheet being removed
    goes with it. A row or column edit that deletes a source table entirely is not covered.
  - `Pivot.add` refuses a placeholder block overlapping another new pivot on the sheet — typically two left at the
    default `ref` `A3`, which Excel cannot refresh — and a model of the wrong shape (`ExcelArgumentError`).
- Streaming (`Stream.*`):
  - `WorkbookWriter#addWorksheet` applies the document model's sheet-name rules: a duplicate (case-insensitive),
    illegal characters, a leading or trailing quote, `History`, and over 31 characters _(owner decision)_ are a
    `WorksheetNameError`. It wrote packages `Workbook.read` refused.
  - `Stream.setCellNote` / `setCellComment` check the note as `Cell.setNote` / `setComment` do; a bad note crashed
    `WorksheetWriter#commit`. The browser entry's style setters, which checked nothing, now share the Node
    implementation.
  - After `commit`, `ExcelStreamStateError` everywhere: a committed `WorksheetWriter` refuses `addRow`, `getRow` /
    `getCell`, `mergeCells`, `addConditionalFormatting`, `addImage`, `addBackgroundImage` and `addWatermark` (rows threw
    a raw `TypeError`; the rest were silently lost), the `Stream.*` setters refuse a committed row, and a committed
    `WorkbookWriter` refuses a second `commit`, `addWorksheet`, `addCellStyle`, `useBuiltInCellStyle`, `addImage`,
    `protect`, `unprotect` and setting `defaultFont`.
  - `WorksheetWriter#getRow(0)` is a `RowOutOfBoundsError`, `addRow(42)` an `ExcelArgumentError`, `getCell("zz")` and
    `mergeCells("Other!A1:B2")` an `InvalidAddressError`. `protect` checks its options, `addCellStyle` checks the
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
  with the original as `cause`; a non-string `sheetName` on a read was a `WorksheetNameError`. `null` data, a
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
- The 31-character sheet-name rule (`WorksheetNameError`) applies to `ExcelCsv.read` / `addWorksheet` and
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
  `ExcelFileError`, `WorksheetNameError`, `InvalidValueTypeError`, `isExcelError`, `BaseErrorOptions`, plus
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
  `Stream.WriterImageModel`, `Stream.PackageSink`, `Stream.EventListener`; `MarkdownReadOptions` /
  `MarkdownWriteOptions` from `documonster/excel/markdown` (`MarkdownOptions` is still accepted everywhere).
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
  `DrawingMl.parseSpPr`, every structured edit of a read `spPr` (`setSpPrLine`, `setSpPrFill`, `createSpPr`) and
  `ChartClassic.fromColorsXml` → `toColorsXml` dropped `satOff`, `hue` / `hueMod` / `hueOff`, `sat`, `lum`,
  `alphaMod`, `alphaOff`, `comp`, `inv`, `gray`, `gamma`, `invGamma` and `red` / `green` / `blue` with their `Mod` /
  `Off`, and the chart renderers (SVG, PNG, PDF) ignored them. They are now read in order, written back in order by
  the classic and ChartEx writers and the colours sidecar, and drawn through `applyColorTransforms`. A chart part
  written back unchanged was not affected (its raw XML is kept). A malformed list entry is a `ChartOptionsError` on
  `transforms[i]` at write time.
- `ChartClassic.fromColorsXml` reads a self-closing `<cs:variation/>` as one variation. It ran on to the next
  variation's close tag, so Excel's default `colors1.xml` (an empty variation, then eight) read as eight, each one
  index early, and a write from the model shifted every gradient colour.
- `Form.removeCheckbox` (unpublished) is `Form.remove(ws, control)`, the `remove` paired with `Form.list`.
- `ExcelMarkdown.read`, `readFile`, `addWorksheet` and `addWorksheets` take the text or its UTF-8 bytes
  (`MarkdownInput`) — what `toBuffer` / `toBufferSync` produce read back; bytes were refused. The unpublished
  `ExcelMarkdown.readAll` / `readAllFile` are `read` / `readFile` with `{ tables: "all" }`; `addWorksheet` refuses
  `tables: "all"` and `addWorksheets` `tables: "first"`.
- The read-side `map` of `CsvReadOptions` and `MarkdownReadOptions` is declared as it is called:
  `(text: string, column: number) => CellValueInput` — the field's text and its 0-based column; what it returns is
  checked as `Cell.setValue` checks a value. The write-side `CsvWriteOptions.map` and the combined `CsvOptions.map`
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
  views, as `Workbook.read` does (they failed inside or were read element by element). `Io.read`'s input type is
  widened to match. Every function that takes a package or file also takes an `ArrayBuffer`, as `Io.read` did:
  `Package.*`, `Streaming.createReader`, `WordCrypto.isEncrypted` / `decrypt` / `verifyPassword` / `encrypt` /
  `Cfb.read` / `Agile.decryptPackage`, `Convert.fromOdt` / `isFlatOpc` / `flatOpcToParts`.
- `Streaming.createReader` (`StreamingDocxReader`) reads Flat OPC, as `Io.read` does; it failed as a corrupt ZIP. It
  decrypts an encrypted file with `{ password }`, which `StreamingReadOptions` always carried, and refuses one without
  a password with `DocxEncryptedError` (a `DocxError`); every encrypted file was a plain `DocxError`. A package over
  `maxPackageSize` is a `DocxLimitExceededError` (a `DocxParseError`, as before).
- `Package.listParts` / `readPart` / `edit` (and the deprecated `Io.listDocxParts` / `readDocxPart` /
  `editDocxIncremental`) apply the security policy's `maxPartCount` / `maxPartSize` / `maxPackageSize` (they applied
  none), read Flat OPC as `Io.read` does, and refuse an encrypted file with `DocxEncryptedError`; it was a
  `DocxParseError` saying "End of Central Directory not found".
- `Package.fillTemplate` / `patch` / `compileTemplate` (and the deprecated `Io.patchDocument` / `compileTemplate`)
  honor `password`, `securityPolicy` and `checkCrc32` on the read, and `modTime` / `securityPolicy` on the write.
  `{ securityPolicy }` was accepted and ignored.
- `Document.setWatermark(doc, undefined)` / `setBackground(doc, undefined)` remove the watermark / background. They
  refused `undefined`, so there was no in-place way to drop a watermark read from a file.
- The writer names the node path: a hand-built node it cannot write (an unknown `type`, a list that is not an array, a
  number with no XML form) is refused with `DocxArgumentError` whose `field` is the exact property path from the
  document (`body[3].children[0].content[1].type`); it was dropped, or failed with `TypeError` or `XmlError` (now the
  `cause`). The same holds for `Io.toBuffer`, `Io.writeFile`, `Package.*` and the streaming writer, whose `add` /
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
  `oleObjects`). `Io.toBuffer` failed with `Relationship ID "rIdOle1" already exists`.
- `Io.read` reads a Flat OPC document (what `Io.toFlatOpc` writes) passed as bytes or as the string itself, under the
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
  of every later `Io.read`.
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
  written on the primary ones. `Chart.style` reads Word's `mc:AlternateContent`-wrapped `c14:style` too.
- Chart parts are written in schema order: a line series' `c:marker` precedes its `c:dPt`, `c:trendline` /
  `c:errBars` precede the series data with `c:smooth` after it, and `c:scaling` has `c:max` before `c:min`.

**Now refused** — `DocxArgumentError` naming the field, unless noted:

- Dates must be an `xsd:dateTime` (`new Date().toISOString()` gives one): `Document.addComment`,
  `Build.insertedRun` / `deletedRun` / `movedFromRun` / `movedToRun`, `Build.moveFromRangeStart` /
  `moveToRangeStart`. `"not a date"` was written as `w:date="not a date"`. `Document.setCoreProperties` needs string
  text fields and valid `created` / `modified` dates; an invalid date failed inside `Io.toBuffer` with `ArchiveError`.
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
  `NaN` switched it off), flags booleans, `rawXmlPolicy` one of its values. A bad `securityPolicy` given to `Io.read`,
  `Io.toBuffer` and the rest is reported by that call (`Io.read: options.securityPolicy.maxEntries …`, field
  `options.securityPolicy.maxEntries`), not as `Security.resolvePolicy: policy.maxEntries`.
- `compressionLevel` outside 0–9 or an invalid `modTime` in `Io.toBuffer` / `toBase64` / `writeFile` / `toFlatOpc`,
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
  `Io.read({ password })`): a password that is not a string, and a `maxDecryptedSize` that is `NaN`, negative or not a
  number (`NaN` switched the cap off; `Infinity` is still "no limit"). `Font.obfuscate` / `deobfuscate` refuse a key
  that is not 32 hex digits (fields `data`, `fontKey`); non-hex digits were XOR'd as `0`, and the error was a plain
  `DocxError`. `WordCrypto.Agile.deriveKey` / `verifyPassword` / `decryptPackage` and `WordCrypto.Cfb.read` /
  `write` check their arguments; the deprecated flat `deriveEncryptionKey` and `readCfb` stay unchecked.
- `Streaming.createWriter` checks every option as the matching `Document.set*` does, when the writer is created
  rather than at `finalize()`, plus `onProgress` and `addText(…, properties)`; `Streaming.createReader` checks its
  source and options. A non-object `options` is refused by `Io.read`, `Validation.validate`, `WordHtml.fromDocx`,
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
  non-ZIP bytes as `DocxParseError` (cause: the archive error), as `Io.read` does; they threw `ArchiveError`.
  `Convert.flatOpcToParts` reports non-XML input as `DocxParseError` (cause: the `XmlParseError`).
  `DocxInvalidImageError` extends `DocxArgumentError` and has a `field`.
- `field` is the path from the parameter name throughout Word: `Build.chart`, `Build.floatingImage` and `Build.shape`
  name `options.type`, `options.series[0]…`, `options.rId`, `options.fill.color` (were `type`, `rId`, `fill.color`);
  `Build.rect` and its siblings say `Build.rect: …` and `options.outline.width` (were `Build.shape: …` and
  `outline.width`); `Build.table(rows, options)` names `options.layout`, `options.columnWidths[0]` (were
  `properties.layout`, `columnWidths[0]`; the deprecated three-argument form is unchanged);
  `Document.addFloatingImage` and `addTextBox` name `options.wrap`, `options.width`; `Io.read` names `input` (was
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
  `WordHtml.toDocxBody(html)` returns an `HtmlBodyResult`
  (`{ body, images, abstractNumberings, numberingInstances, styles, docDefaults }`). There is no
  `WordHtml.toDocument`. The deprecated flat `htmlToDocx` (`{ body, images }`) and `htmlToDocxBody`
  (`BodyContent[]`) keep their old results. A body's ids are fixed; to splice HTML into an existing document, use
  `Transform.merge([host, WordHtml.toDocx(html)])`. A picture they cannot embed is the italic grey `[Image: alt]` run
  `WordMarkdown.toDocx` writes; it was an image node with an empty `rId`, written as its alt text. The deprecated flat
  forms keep the old placeholder.
- `WordExcel.toChartSvg` and `WordExcel.toChartXlsx` (it returns `.xlsx` bytes) are the names of the Word-chart
  renderers; the deprecated `renderWordChartSvg` and `generateChartEmbeddedXlsx` forward to them.
- `WordCrypto.Agile.decryptPackage(pkg, info, password, { maxDecryptedSize })` takes the cap in `DecryptOptions`, as
  `WordCrypto.decrypt` does. Both default to `DEFAULT_SECURITY_POLICY.maxPackageSize` (500 MiB), the cap `Io.read`
  applies; the deprecated `decryptDocx` keeps 512 MiB.
- `Convert.flatOpcToParts`, the inverse of `Convert.partsToFlatOpc`; the deprecated `Convert.parseFlatOpc` forwards
  to it.
- `WordCrypto.hasSignatures(partNames)`: check a file from its part names without reading the document,
  `WordCrypto.hasSignatures(await Package.listParts(bytes))` — what the deprecated `hasDigitalSignatures` did.
- `DocxArgumentError` (`documonster/word`, and also `documonster/word/crypto` with `DocxDecryptionError` and
  `isDocxError`). `documonster/word/html`, `/word/markdown` and `/word/excel` export `DocxError`,
  `DocxArgumentError`, `isDocxError` and `BaseErrorOptions` (markdown also `DocxInvalidImageError`).
  `DiscardUnrecognizedValuesOptions` is a flat type export, as in Excel.
- `DocxWriteOptions` (the options of `Io.toBuffer` / `toBase64` / `writeFile`), `Build.SimpleTableOptions`,
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
  now, each an `UnrecognizedValue` (`{ path, value }`, exported); `Document.discardUnrecognizedValues`'s `filter`
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
  and Word (`applyColorTransforms` in `@utils/theme-colors`, Office's semantics: `lum*` / `sat*` / `hue*` in HSL,
  `tint` / `shade` / `inv` / `gamma` in linear RGB). They were ignored: an Excel-saved gridline or axis
  (`tx1` lumMod 15 % / lumOff 85 %) was drawn black instead of `#D9D9D9`, axis labels (`tx1` 65 % / 35 %) black instead
  of `#595959`. A gridline or axis line with `<a:ln><a:noFill/>` is no longer drawn in a default grey. Scheme colors
  resolve against the workbook's theme in `Chart.toSvg` / `Chart.toPng`, `Pdf.fromChart` and `Pdf.fromExcel` — a
  workbook saved by a recent Excel has the 2023 Office theme (`accent1` = `#156082`), not the 2013 one the preview
  assumed (`#4472C4`) — and against the new `ChartRenderOptions.themeColors` for a model rendered on its own. A
  chart's automatic series colours are its theme's accents too (they were the default theme's whatever the workbook
  said). Word charts follow the document's theme (`doc.theme.colorScheme`) in `Pdf.fromDocx`, in
  `WordExcel.toChartSvg(chart, { themeColors })` (new `WordChartSvgOptions`, checked as `DocxArgumentError`) and, through
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
  `ChartOptionsError`: `field` is the path from the parameter name (`options.fontSize`,
  `options.watermark.pages[0]`, `options.sheets[1]`, `metadata.title`, `encryption.ownerPassword`, `rect.width`,
  `source`, `data`). PDF and Mermaid used paths relative to the options object (`fontSize`, `pages[0]`,
  `theme.nodeFill`), and Mermaid named `diagram` for the parameter named `source`. Messages start with the public call:
  `Pdf.fromExcel: …`, `Pdf.Builder#addPage: …`, `Pdf.PageBuilder#drawText: …` (a page from `Pdf.Builder#addPage` or
  `Pdf.Editor#addPage`), `Pdf.EditorPage#drawText: …`, `Pdf.Editor#splitPages: …`, `Pdf.Editor.read: …`,
  `Pdf.Signature.verify: …`, `PdfContentStream#setFont: …`, `Draw.toRgba: …`, `Draw.Geometry.flattenPath: …`,
  `new Draw.Backend.BasicRasterCanvas: …`, `Mermaid.toSvg: …` / `Mermaid.toDrawList: …` (they shared one prefix),
  `ChartClassic.toSvg: …`, `ChartEx.createModel: …`, `Chart.toPng: …`.
- `ChartOptionsError` extends `ExcelArgumentError` (it extended `ExcelError`), so `instanceof ExcelArgumentError` is
  now true as well and `field` names the option (`options.series[0].values`, `preset`, `table`, `rows[2].price`,
  `model`). Validator messages name the path from `options`, not from `chart`, `chartEx` or `combo chart`:
  `options.series[0].trendline is not valid for pie charts` (was `chart.series[0].…`).
- Chart render options (`ChartClassic.toSvg` / `toPng` / `toDrawList` / `Backend.toScene`, `ChartEx.toSvg` /
  `toPng` / `toDrawList`, `Chart.toSvg` / `Chart.toPng`, and the deprecated `drawChartPdf` / `drawChartExPdf`), with
  `ChartOptionsError` before drawing: `width`,
  `height`, `scale` and `dpi` must be finite numbers > 0 (`toSvg` wrote `width="NaN"`, `width="0"` or a negative
  `viewBox`; `toPng` drew a 1-pixel-wide image, or for `NaN` threw a `DrawArgumentError` naming the internal
  `BasicRasterCanvas`); `title` must be a string, `background` a CSS color or `{ r, g, b, a }` with channels in
  0..1 and the deprecated `backgroundColor` a CSS color (a number failed with a `TypeError`, an unreadable string was
  drawn black); `fonts` an array of fonts and `useSystemFonts` a boolean; `regionMap` is `{ topology, objectName,
match?, projection?, strokeColor? }`, each checked. `drawChartExPdf` needs a surface with `drawRect` / `drawLine` /
  `drawText` and a `rect` of finite numbers with an area, and checks the model (messages start `drawChartExPdf:`).
- A `regionMap.topology` that cannot be read — not an object, no `objects`, an `objectName` it does not have, `arcs`
  not a list of point lists, a `transform` that is not two finite pairs, a geometry whose `arcs` or `coordinates` do
  not have its type's shape — is a `ChartOptionsError` (from `Pdf.fromChart`, a `PdfArgumentError`) naming the exact
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
  Word's `Io.toBuffer` and the XLSB writer move the document's date to the nearest one a ZIP can store.
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
  `ZipEditor#getEntries` lists every entry.
- `ZipEditor.apply(plan)` changes nothing when it throws; it applied every operation before the refused one.
- Readers hand out frozen data: `RemoteZipReader.getEntries()` / `getEntry()` / `filterEntries()` / `findEntries()`
  and `ZipEditor.getEntries()` (mutating them corrupted the reader); `ZipEditPlan.getOperations()` returns a frozen
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
`Archive.Crypto.randomBytes: …`, `Streams.bytes: …`, `Streams.createTransform: expected a function, got null` (they
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
  `createAesExtraField`, and a negative or fractional length to `Crypto.randomBytes`.
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

Every public name now follows the conventions in `AGENTS.md` ("Public API Conventions"). **Nothing is
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
  unchanged, but the result may share unchanged nodes with it — take `Document.clone(doc)` before editing a
  result in place while you still use the input.
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
    `Table.getColumn` throws for an index the table has no column at; the deprecated `Table.column` does not.
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
  `ExcelEncryptedError`, `PositionErrorOptions`, and the widenings `DefinedNames.add` / `addFormula` with
  `{ localSheetId }` returning the name as `get` reports it (they returned `void`) and `Worksheet.getValues(ws,
{ base })` — `base: 0` indexes as `Row.getValues` does; without options it stays 1-based.
- Duplicates deprecated in favour of the name they forward to: `Chart.setBuiltInStyle` (`setStyle`),
  `Chart.setTitleRichText` (`setTitle`), `Worksheet.destroy` (`Workbook.removeWorksheet`, which now also takes the
  sheet's handle) and `Workbook.renameChartsheet` (`Chartsheet.setName`).
- `DataValidation.add`, `find` and `remove` take the worksheet, as every other sheet-level namespace does; the rule
  collection (`ws.dataValidations`) is still accepted.
- Arguments refused at the call that used to pass, or fail later with an untyped `TypeError`: an invalid `Date`
  as a cell value or formula result (`InvalidValueTypeError`; the write failed on it), `Workbook.removeWorksheet`
  with no sheet (it removed the first one), `ExcelCsv.read` / `addWorksheet` with input that is not text, bytes, a URL,
  a Blob or a stream (it read `String(input)` — `42` became a one-cell sheet), `Csv.parse` with input that is not a
  string (`CsvError`), `Io.read` with input that is not bytes (`DocxArgumentError`; an `ArrayBuffer` is still
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
  named no table. It now throws a `TableError` naming a form Excel accepts; so do `Table.setName`,
  `Table.setDisplayName` and `Table.setModel`. A file whose table carries such a name is still read with it corrected.
- `ExcelCsv` takes `CsvReadOptions` where it reads (`read`, `addWorksheet`, `readFile`, `createImportStream`) and
  `CsvWriteOptions` where it writes (`format`, `toBuffer`, `toBufferSync`, `toStream`, `writeStream`, `writeFile`), so an
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
  `Csv.format` and the PDF `drawPath` / `splitPages` take readonly arrays, which they only read; `Io.read` declares
  the `ArrayBuffer` it has always read.
- `Table.create` and `Form.createCheckbox` are deprecated: each built an object for a sheet without adding it
  (a table that was never written; a checkbox that took a shape id another one could also take). Use
  `Table.add` and `Form.addCheckbox`.
- New error guards: `isArchiveError` (`documonster/archive`), `MermaidError` / `isMermaidError`
  (`documonster/mermaid`, with its syntax error — now `MermaidParseError` — extending `MermaidError`), and `documonster/chart`
  exports the `ChartOptionsError` it throws.
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
| `Chart.chartTypes`                      | Use `Chart.getChartTypes`.                                                                                                                                                                                                                                                                     |
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
| `Form.vmlAnchor`                        | Use `Form.getVmlAnchor`.                                                                                                                                                                                                                                                                       |
| `Form.vmlCheckedValue`                  | Use `Form.getVmlCheckedValue`.                                                                                                                                                                                                                                                                 |
| `Form.vmlStyle`                         | Use `Form.getVmlStyle`.                                                                                                                                                                                                                                                                        |
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
| `Table.column`                          | Use `Table.getColumn`, which refuses an index the table has no column at.                                                                                                                                                                                                                      |
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
| `Workbook.renameChartsheet`             | Use `Chartsheet.setName(Workbook.getChartsheet(wb, nameOrIndex), name)`.                                                                                                                                                                                                                       |
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
| `Archive.randomBytes`                        | Use `Crypto.randomBytes`.                                                                                                                                                       |
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
| `createReadStream`                           | Use `Fs.createReadStream`.                                                                                                                                                      |
| `createTempDir`                              | Use `Fs.createTempDir`.                                                                                                                                                         |
| `createTempDirSync`                          | Use `Fs.createTempDirSync`.                                                                                                                                                     |
| `createUnzlibStream`                         | Use `Compression.createUnzlibStream`.                                                                                                                                           |
| `createWriteStream`                          | Use `Fs.createWriteStream`.                                                                                                                                                     |
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
| `ReadStreamOptions`                          | Use `FsReadStreamOptions`.                                                                                                                                                      |
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
| `WorkerPool`                                 | Use `Compression.Workers.WorkerPool`.                                                                                                                                           |
| `writeFileBytes`                             | Use `Fs.writeFileBytes`.                                                                                                                                                        |
| `writeFileBytesSync`                         | Use `Fs.writeFileBytesSync`.                                                                                                                                                    |
| `writeFileText`                              | Use `Fs.writeFileText`.                                                                                                                                                         |
| `writeFileTextSync`                          | Use `Fs.writeFileTextSync`.                                                                                                                                                     |
| `WriteStreamOptions`                         | Use `FsWriteStreamOptions`.                                                                                                                                                     |
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

| Deprecated            | Use instead                                                      |
| --------------------- | ---------------------------------------------------------------- |
| `apply`               | Use `MatrixNs.applyMatrix.apply`.                                |
| `arcToCubics`         | Use `Draw.Geometry.arcToCubics`.                                 |
| `BasicRasterCanvas`   | Use `Draw.Backend.BasicRasterCanvas`.                            |
| `createRasterSurface` | Use `Draw.Backend.createRasterSurface`, which takes `{ scale }`. |
| `cssColour`           | Use `Draw.parseColor`.                                           |
| `downsample`          | Use `Draw.Backend.downsample`.                                   |
| `flattenPath`         | Use `Draw.Geometry.flattenPath`.                                 |
| `fontHasGlyph`        | Use `Draw.Backend.fontHasGlyph`.                                 |
| `IDENTITY`            | Use `MatrixNs.IDENTITYMatrix.IDENTITY`.                          |
| `isSimpleText`        | Use `Draw.isSimpleText`.                                         |
| `measureText`         | Use `Draw.measureText`.                                          |
| `multiply`            | Use `MatrixNs.multiplyMatrix.multiply`.                          |
| `normalizeSamples`    | Use `Draw.Backend.normalizeSamples`.                             |
| `parseRasterFont`     | Use `Draw.parseRasterFont`.                                      |
| `rasterizeToRgba`     | Use `Draw.toRgba`.                                               |
| `rectNode`            | Use `Draw.Geometry.createRectNode`.                              |
| `relativeLuminance`   | Use `Draw.getRelativeLuminance`.                                 |
| `renderDrawList`      | Use `Draw.render`.                                               |
| `renderNode`          | Use `Draw.Backend.renderNode`.                                   |
| `rotate`              | Use `MatrixNs.rotateMatrix.rotate`.                              |
| `rotationOf`          | Use `MatrixNs.getRotationMatrix.getRotation`.                    |
| `roundedRectToPath`   | Use `Draw.Geometry.roundedRectToPath`.                           |
| `scale`               | Use `MatrixNs.scaleMatrix.scale`.                                |
| `sectorToPath`        | Use `Draw.Geometry.sectorToPath`.                                |
| `SvgSurface`          | Use `Draw.Backend.SvgSurface`.                                   |
| `textFeaturesOf`      | Use `Draw.getTextFeatures`.                                      |
| `toSvg`               | Use `Draw.toSvg`.                                                |
| `translate`           | Use `MatrixNs.translateMatrix.translate`.                        |
| `translucent`         | Use `Draw.withAlpha`.                                            |
| `uniformScale`        | Use `MatrixNs.getUniformScaleMatrix.getUniformScale`.            |
| `widestText`          | Use `Draw.measureWidest`.                                        |
| `wrapText`            | Use `Draw.wrapText`.                                             |

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
| `setSpPrFill`                 | Use `DrawingMl.setSpPrFill`.                                                                           |
| `setSpPrLine`                 | Use `DrawingMl.setSpPrLine`.                                                                           |

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

| Deprecated                         | Use instead                                                                                                                                                             |
| ---------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Build.checkBox`                   | Use `Build.checkbox`.                                                                                                                                                   |
| `Build.createArrow`                | Use `Build.arrow`.                                                                                                                                                      |
| `Build.createCallout`              | Use `Build.callout`.                                                                                                                                                    |
| `Build.createEllipse`              | Use `Build.ellipse`.                                                                                                                                                    |
| `Build.createFlowchartShape`       | Use `Build.flowchartShape`.                                                                                                                                             |
| `Build.createLine`                 | Use `Build.line`.                                                                                                                                                       |
| `Build.createRect`                 | Use `Build.rect`.                                                                                                                                                       |
| `Build.createRoundRect`            | Use `Build.roundRect`.                                                                                                                                                  |
| `Build.createShape`                | Use `Build.shape`.                                                                                                                                                      |
| `Build.createStar`                 | Use `Build.star`.                                                                                                                                                       |
| `Build.mathPreSubSuperScript`      | Use `Build.mathPreSubSuperscript`.                                                                                                                                      |
| `Build.mathSubScript`              | Use `Build.mathSubscript`.                                                                                                                                              |
| `Build.mathSubSuperScript`         | Use `Build.mathSubSuperscript`.                                                                                                                                         |
| `Build.mathSuperScript`            | Use `Build.mathSuperscript`.                                                                                                                                            |
| `Convert.docxToSemantic`           | Use `Convert.toSemantic`.                                                                                                                                               |
| `Convert.mapDocument`              | Use `Transform.map`.                                                                                                                                                    |
| `Convert.mathMLToOmml`             | Use `Convert.mathMlToOmml`.                                                                                                                                             |
| `Convert.ommlToMathML`             | Use `Convert.ommlToMathMl`.                                                                                                                                             |
| `Convert.parseFlatOpc`             | Use `Convert.flatOpcToParts`, the inverse of `Convert.partsToFlatOpc`, or `Io.read`, which now reads Flat OPC bytes as a document.                                      |
| `Convert.readOdt`                  | Use `Convert.fromOdt`.                                                                                                                                                  |
| `Convert.toFlatOpc`                | Use `Convert.partsToFlatOpc` — this wraps a part map, not a document; `Io.toFlatOpc(doc)` is the document form.                                                         |
| `Convert.writeOdt`                 | Use `Convert.toOdt`.                                                                                                                                                    |
| `Diff.documents`                   | Use `Diff.compare`.                                                                                                                                                     |
| `Document.addParagraphElement`     | Use `Document.addParagraph`.                                                                                                                                            |
| `Document.addTableElement`         | Use `Document.addTable`.                                                                                                                                                |
| `Document.build`                   | A document needs no build step: pass the document itself to `Io.toBuffer`, `Layout.*`, `WordHtml.fromDocx` or `Pdf.fromDocx`, which apply the same write-time defaults. |
| `Document.DocumentHandle`          | Use `DocxDocument`.                                                                                                                                                     |
| `Document.insertContentAt`         | Use `Document.insertContent`.                                                                                                                                           |
| `Document.nextBookmarkId`          | Use `Document.allocateBookmarkId`.                                                                                                                                      |
| `DocumentHandle`                   | Use `DocxDocument`.                                                                                                                                                     |
| `Glossary.autoTextEntries`         | Use `Glossary.listAutoTextEntries`.                                                                                                                                     |
| `Glossary.quickParts`              | Use `Glossary.listQuickParts`.                                                                                                                                          |
| `Io.compileTemplate`               | Use `Package.compileTemplate`.                                                                                                                                          |
| `Io.editDocxIncremental`           | Use `Package.edit`.                                                                                                                                                     |
| `Io.fillTemplateFromBuffer`        | Use `Package.fillTemplate`.                                                                                                                                             |
| `Io.listDocxParts`                 | Use `Package.listParts`.                                                                                                                                                |
| `Io.merge`                         | Use `Transform.merge`.                                                                                                                                                  |
| `Io.patchDocument`                 | Use `Package.patch`.                                                                                                                                                    |
| `Io.patchTemplate`                 | Use `Package.patchTemplate`.                                                                                                                                            |
| `Io.readDocxPart`                  | Use `Package.readPart`.                                                                                                                                                 |
| `Io.split`                         | Use `Transform.split`.                                                                                                                                                  |
| `Io.toFlatOpcFromDoc`              | Use `Io.toFlatOpc`, which takes an options object (`{ compressionLevel }`).                                                                                             |
| `Io.updateFields`                  | Use `Transform.updateFields`.                                                                                                                                           |
| `Io.updateTableOfContents`         | Use `Transform.updateTableOfContents`.                                                                                                                                  |
| `Layout.document`                  | Use `Layout.getPageMap`.                                                                                                                                                |
| `Layout.documentFull`              | Use `Layout.compute`.                                                                                                                                                   |
| `Layout.renderDocumentToSvg`       | Use `Layout.toSvg`.                                                                                                                                                     |
| `Layout.renderPageFromLayout`      | Use `Layout.computedPageToSvg`.                                                                                                                                         |
| `Layout.renderPageToSvg`           | Use `Layout.pageToSvg`.                                                                                                                                                 |
| `Query.acceptAllRevisions`         | Use `Document.acceptAllRevisions`.                                                                                                                                      |
| `Query.acceptRevision`             | Use `Document.acceptRevision`.                                                                                                                                          |
| `Query.collectParagraphs`          | Use `Query.listParagraphs`.                                                                                                                                             |
| `Query.collectRuns`                | Use `Query.listRuns`.                                                                                                                                                   |
| `Query.collectTables`              | Use `Query.listTables`, passing the parts to include — `listTables` searches only the body unless asked.                                                                |
| `Query.countByFormat`              | Use `Query.getCountByFormat`.                                                                                                                                           |
| `Query.countWords`                 | Use `Query.getWordCount`.                                                                                                                                               |
| `Query.extractFormFields`          | Use `Query.listFormFields`.                                                                                                                                             |
| `Query.extractText`                | Use `Query.getText`.                                                                                                                                                    |
| `Query.fillFormFields`             | Use `Transform.fillFormFields`.                                                                                                                                         |
| `Query.findBookmark`               | Use `Query.getBookmark`.                                                                                                                                                |
| `Query.findComment`                | Use `Query.getComment`.                                                                                                                                                 |
| `Query.getHeadings`                | Use `Query.listHeadings`.                                                                                                                                               |
| `Query.getUsedFormats`             | Use `Query.listUsedFormats`.                                                                                                                                            |
| `Query.indexStyles`                | Use `Query.createStyleIndex`.                                                                                                                                           |
| `Query.mailMerge`                  | Use `Document.mailMerge`.                                                                                                                                               |
| `Query.paragraphCount`             | Use `Query.getParagraphCount`.                                                                                                                                          |
| `Query.rejectAllRevisions`         | Use `Document.rejectAllRevisions`.                                                                                                                                      |
| `Query.rejectRevision`             | Use `Document.rejectRevision`.                                                                                                                                          |
| `Query.replaceText`                | Use `Document.replaceText`.                                                                                                                                             |
| `Query.resolveDataBindings`        | Use `Transform.resolveDataBindings`.                                                                                                                                    |
| `Query.resolveNumberingLevel`      | Use `Query.getNumberingLevel`.                                                                                                                                          |
| `Query.resolveRunStyle`            | Use `Query.getRunStyle`.                                                                                                                                                |
| `Query.resolveStyle`               | Use `Query.getParagraphStyle`.                                                                                                                                          |
| `Query.resolveTableStyle`          | Use `Query.getTableStyle`.                                                                                                                                              |
| `Query.setCompatibilityMode`       | Use `Document.setCompatibilityMode`.                                                                                                                                    |
| `Query.tableCount`                 | Use `Query.getTableCount`.                                                                                                                                              |
| `Query.walkBlocks`                 | Use `Query.eachBlock`.                                                                                                                                                  |
| `Query.walkDocument`               | Use `Query.each`.                                                                                                                                                       |
| `Security.encrypt`                 | Use `WordCrypto.encrypt` from `documonster/word/crypto`, beside `decrypt` and `verifyPassword`.                                                                         |
| `Security.resolveSecurityPolicy`   | Use `Security.resolvePolicy`.                                                                                                                                           |
| `Streaming.createDocxStream`       | Use `Streaming.createWriter`.                                                                                                                                           |
| `Streaming.createDocxStreamReader` | Use `Streaming.createReader`.                                                                                                                                           |
| `Template.fillTemplate`            | Use `Template.fill`, which returns a new document.                                                                                                                      |
| `Template.fillTemplateEnhanced`    | Use `Template.fill` with `{ richPlaceholders: true }`.                                                                                                                  |
| `Template.fillTemplateFromSource`  | Use `Template.fillFromSource`.                                                                                                                                          |
| `Template.isTemplateChart`         | Use `Template.isChart`.                                                                                                                                                 |
| `Template.listTemplateTags`        | Use `Template.listTags`.                                                                                                                                                |
| `Template.TemplateError`           | Import `TemplateError` from `documonster/word` directly.                                                                                                                |
| `Validation.document`              | Use `Validation.validate`.                                                                                                                                              |

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
  fixed-width ISO parsers in `@utils/datetime` were written to be selected by
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
Cell.setDateParts(ws, "A1", { year: 2024, month: 1, day: 15 });
Cell.setValue(ws, "A1", Temporal.PlainDate.from("2024-01-15"));
```

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
await Pdf.fromExcel(workbook, {
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
`@pdf/render/layout-engine` to call `layoutSheet` with a hand-built options
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

`createCsvReadStream(workbook, options?)` from `documonster/excel/csv` used to
return an **object-mode** stream that emitted `string` chunks. It now returns a
normal **byte** stream: `objectMode` is off on the readable side and every chunk
is a `Uint8Array` (a `Buffer` on Node).

The declared return type changed from `IReadable<Uint8Array | string>` to
`IReadable<Uint8Array>`. On the Node entry it is narrowed further, to
`IReadable<Uint8Array> & Readable`, so the value is a nominal `stream.Readable`:

```ts
// Node — no adapter, no cast
await pipeline(createCsvReadStream(wb), createWriteStream("out.csv"));
const web = Readable.toWeb(createCsvReadStream(wb));
```

The writable half of the underlying `Transform` is deliberately **not** exposed.
The stream is already attached to a producer, so `write()` would inject rows into
the CSV; `write()` / `end()` therefore do not typecheck on the returned value.

**Why.** The old shape leaked the CSV _formatter's_ internals into a public API.
It also made the stream indistinguishable by type from `Workbook.toStream()`
while behaving completely differently, so anything that consumes bytes silently
misbehaved instead of failing to compile:

```ts
// Before: threw `list[0] argument must be an instance of Buffer or Uint8Array`
Buffer.concat(chunks);

// Before: produced a ReadableStream of *strings*, which an SDK then rejected
// or uploaded as "[object Object]"-shaped garbage.
Readable.toWeb(createCsvReadStream(wb));
```

Both now work. The CSV content is byte-for-byte identical to before — only the
chunk type changed.

**If you relied on string chunks**, decode explicitly, or use the string API:

```ts
// A chunk boundary can fall inside a multi-byte character, so decode with a
// streaming decoder rather than per chunk.
const decoder = new TextDecoder();
let csv = "";
for await (const chunk of createCsvReadStream(wb)) {
  csv += decoder.decode(chunk, { stream: true });
}
csv += decoder.decode();

// …or skip streaming entirely when the whole document fits in memory:
const csv = writeCsv(wb);
```

Building a string by `+=`-ing chunks is the one pattern that breaks **silently**
— `"" + new Uint8Array([97])` yields `"97"`, not `"a"`. Search for that shape.

`CsvFormatterStream` from `documonster/csv` is **unchanged** and still emits
strings by default. It gained a `readableObjectMode` option to opt into bytes:

```ts
new CsvFormatterStream({ readableObjectMode: false }); // rows in, bytes out
```

### Node IO declarations now match the runtime (non-breaking)

Two Node-entry signatures were narrowed to the types Node already returned.
Both are strict narrowings, so existing code keeps compiling — but the wrappers
and casts it needed can be deleted.

**`Workbook.toBuffer` returns `Promise<Buffer>`** on `documonster/excel` (the
browser entry still returns `Promise<Uint8Array>`). The runtime already produced
a `Buffer`; the `Uint8Array` declaration is what pushed callers into a defensive
`Buffer.from(bytes)`, which copies the entire package.

```ts
// Before
const bytes = await Workbook.toBuffer(wb);
await upload(Buffer.from(bytes)); // full copy, purely to satisfy the type

// Now
await upload(await Workbook.toBuffer(wb));
```

**`Workbook.toStream` returns a nominal `stream.Readable` on Node.** The shared
`XlsxReadable` type is unchanged and **identical in both builds** — it is still
`IReadable<Uint8Array>`, so cross-platform code written against that name keeps
compiling as-is. Node's `toStream` refines only its _return type_ to
`XlsxReadable & Readable`, which is assignable to `XlsxReadable`:

```ts
// Before: an adapter, or a bet on an implementation detail
const body = Readable.from(Workbook.toStream(wb));
const body = Workbook.toStream(wb) as unknown as Readable;

// Now
const body = Workbook.toStream(wb);
const web = Readable.toWeb(Workbook.toStream(wb));

// Still valid, and still means the same type on both platforms
const portable: Workbook.XlsxReadable = Workbook.toStream(wb);
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
// Sets errored to AbortError
for await (const chunk of Workbook.toStream(wb)) break;

// Silent: destroy first, then break
const stream = Workbook.toStream(wb);
for await (const chunk of stream) {
  stream.destroy();
  break;
}

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

The same applies to `createCsvReadStream` on Node, which is also a
`stream.Readable`.
