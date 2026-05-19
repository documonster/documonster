# Formula Module

[中文](README_zh.md)

Standalone Excel-compatible formula engine — tokenizer, parser, compiler, evaluator, dependency graph, dynamic-array spill materialiser, and 433 built-in functions. Zero runtime dependencies.

## Two usage modes

This module ships **two complementary entry points**. Pick whichever
matches your integration style — you never pay for the one you don't
import.

| Mode                        | Entry point              | Use when                                                                                            |
| --------------------------- | ------------------------ | --------------------------------------------------------------------------------------------------- |
| **Paired with `Workbook`**  | `installFormulaEngine()` | You use the excel module's `Workbook` class and want `wb.calculateFormulas()` + PDF auto-recalc.    |
| **Standalone / functional** | `calculateFormulas(wb)`  | You operate on a `WorkbookLike` object (custom host, server-side recalc, tests) — no excel runtime. |

The engine code itself is identical; only the host glue differs. See
[Why a separate subpath?](#why-a-separate-subpath) for the tree-shake
numbers.

## Features

- **Full expression pipeline** — tokenizer → AST → binder → compiled form → evaluator → writeback plan
- **Dependency graph** — topological evaluation, circular-ref detection, iterative calculation
- **Dynamic arrays with spill** — `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, spill-error detection, ghost-cell cleanup
- **Higher-order functions** — `LAMBDA`, `LET`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`
- **Array semantics** — implicit intersection, broadcasting, CSE array formulas
- **Structured references** — Tables, `[#This Row]`, `[@Column]`, `[#Totals]`
- **Shared formulas** — read, translate, recalculate
- **Defined names** — scoped, formula-based, range unions
- **Cross-sheet references** — `Sheet2!A1`, 3D ranges `Sheet1:Sheet3!A1`
- **R1C1 addressing** — both A1 and R1C1 modes via `INDIRECT`
- **433 built-in Excel functions** across 11 categories
- **Zero runtime dependencies** — no npm deps, no polyfills

## Function Coverage

| Category      | Count | Highlights                                                                              |
| ------------- | ----- | --------------------------------------------------------------------------------------- |
| Math & Trig   | ~70   | `SUM`, `PRODUCT`, `ROUND`, `CEILING`, `POWER`, `MMULT`, `MDETERM`, `SIN`, `ATAN2`       |
| Text          | ~55   | `CONCAT`, `TEXTJOIN`, `TEXT`, `LEFT`, `MID`, `SUBSTITUTE`, `REGEXTEST`, `REGEXEXTRACT`  |
| Logical       | ~15   | `IF`, `IFS`, `AND`, `OR`, `SWITCH`, `IFERROR`, `XOR`                                    |
| Date & Time   | ~30   | `TODAY`, `DATEDIF`, `EDATE`, `NETWORKDAYS`, `WEEKNUM`, `ISOWEEKNUM`, `WORKDAY.INTL`     |
| Lookup & Ref  | ~25   | `VLOOKUP`, `XLOOKUP`, `XMATCH`, `INDEX`, `OFFSET`, `INDIRECT`, `ADDRESS`, `CHOOSE`      |
| Statistical   | ~90   | `AVERAGE`, `STDEV`, `NORM.DIST`, `T.TEST`, `PERCENTILE`, `QUARTILE`, `CORREL`, `LINEST` |
| Financial     | ~50   | `PMT`, `IRR`, `NPV`, `XIRR`, `RATE`, `PRICE`, `DURATION`, `COUPNUM`, `ACCRINT`          |
| Dynamic Array | ~25   | `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `TAKE`, `DROP`, `VSTACK`, `TEXTSPLIT`, `LAMBDA` |
| Database      | ~12   | `DSUM`, `DCOUNT`, `DAVERAGE`, `DMAX`, `DMIN`, `DSTDEV`, `DPRODUCT`                      |
| Engineering   | ~45   | `DEC2BIN`, `BITAND`, `COMPLEX`, `IMSUM`, `ERF`, `BESSELJ`                               |
| Information   | ~20   | `ISNUMBER`, `ISBLANK`, `ISREF`, `N`, `TYPE`, `CELL`, `FORMULA`                          |

See `functions/` for the full list; `runtime/function-registry.ts` is the registration site.

## Quick Start

### Paired with `Workbook` (most common)

```typescript
import { Workbook } from "@cjnoname/excelts";
import { installFormulaEngine } from "@cjnoname/excelts/formula";

// Call once at startup — wires the engine into Workbook.calculateFormulas(),
// the PDF bridge's automatic recalc, and strict defined-name classification
// during XLSX load.
installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");
ws.getCell("A1").value = 10;
ws.getCell("A2").value = 20;
ws.getCell("A3").value = 30;
ws.getCell("A4").value = { formula: "SUM(A1:A3)" };

wb.calculateFormulas();
console.log(ws.getCell("A4").result); // 60
```

### Standalone / functional

The engine runs on any object shaped like `WorkbookLike` — you do **not**
have to use the excel module at all. A bundle that imports only
`calculateFormulas` pulls zero excel runtime code.

```typescript
import { calculateFormulas, type WorkbookLike } from "@cjnoname/excelts/formula";

// Your own data — any object implementing WorkbookLike works.
// No Workbook class, no installFormulaEngine() required.
const wb: WorkbookLike = buildMyWorkbookLike();

calculateFormulas(wb); // pure function, zero global side effects
```

This mode is ideal for:

- Server-side recalculation of cached XLSX files
- Custom spreadsheet hosts that already have their own data model
- Tests and benchmarks that want deterministic, per-instance behaviour
- Concurrent evaluation of multiple workbooks without touching process globals

### Per-Workbook syntax probe

If you use `Workbook` but want strict defined-name classification for a
specific instance without touching process-global state, inject a
probe explicitly:

```typescript
import { Workbook } from "@cjnoname/excelts";
import { createFormulaSyntaxProbe } from "@cjnoname/excelts/formula";

const wb = new Workbook({ formulaSyntaxProbe: createFormulaSyntaxProbe() });
await wb.xlsx.load(buffer);
// defined names classified strictly using the injected probe;
// installFormulaEngine() is not required for classification here.
```

### Tokenise / parse without evaluating

```typescript
import { tokenize, parse } from "@cjnoname/excelts/formula";

const tokens = tokenize("SUM(A1:B10) + VLOOKUP(key, table, 2, FALSE)");
const ast = parse(tokens); // throws on syntax errors
```

## Why a separate subpath?

The formula engine is ~200 KB minified. Most callers of `@cjnoname/excelts`
only read and write XLSX files and let Excel recalculate on open — pulling
the engine into those bundles unconditionally would be a large, invisible
cost.

The subpath gives you three tree-shaking outcomes:

| Imports                                             | Excel module | Formula engine |
| --------------------------------------------------- | ------------ | -------------- |
| `Workbook` from root only                           | ✓            | ✗              |
| `calculateFormulas` from `/formula`                 | ✗            | ✓              |
| `Workbook` + `installFormulaEngine` from `/formula` | ✓            | ✓              |

The functional `calculateFormulas` API operates on the structural
`WorkbookLike` interface and pulls **no** excel runtime code — you can
hand it any object shaped like a workbook. Server-side recalculation of
a cached XLSX loaded by the excel module works too; the excel import
stays in the excel bundle.

> **IIFE note:** The script-tag IIFE bundle
> (`dist/iife/excelts.iife.min.js`) intentionally excludes the formula
> engine so it stays lean. Script-tag users who need formula
> calculation should switch to ESM and call `installFormulaEngine()`
> from this subpath.

## Examples

Runnable examples live in `src/modules/formula/examples/`:

| File                         | What it demonstrates                                           |
| ---------------------------- | -------------------------------------------------------------- |
| `formula-math.ts`            | Arithmetic, rounding, trig, matrix, power & log                |
| `formula-text.ts`            | Slicing, search/replace, concat, formatting, regex             |
| `formula-logical.ts`         | `IF`/`IFS`, boolean ops, `IFERROR`, `SWITCH`, `CHOOSE`         |
| `formula-date.ts`            | Date construction, extract, duration, business days, format    |
| `formula-lookup.ts`          | `VLOOKUP`, `XLOOKUP`, `INDEX/MATCH`, `OFFSET`, `INDIRECT`      |
| `formula-statistical.ts`     | Descriptive stats, conditional aggregates, regression, dists   |
| `formula-financial.ts`       | Loans, TVM, NPV/IRR, depreciation                              |
| `formula-dynamic-array.ts`   | `FILTER`/`SORT`/`UNIQUE`, spill, `SEQUENCE`, `LAMBDA`/`REDUCE` |
| `formula-database.ts`        | `DSUM`/`DCOUNT`/`DAVERAGE` with criteria ranges                |
| `formula-engineering.ts`     | Base conversions, bitwise, complex numbers, ERF/BESSELJ        |
| `formula-standalone.ts`      | Functional API + `tokenize`/`parse` without evaluation         |
| `formula-pdf-integration.ts` | Automatic recalc during `excelToPdf()`                         |

Run any example:

```bash
npx tsx src/modules/formula/examples/formula-math.ts
npx tsx src/modules/formula/examples/formula-dynamic-array.ts
npx tsx src/modules/formula/examples/formula-pdf-integration.ts
# Output: tmp/formula-examples/formula-pdf-integration.pdf
```

## Architecture

The engine is a six-layer pipeline (see `AGENTS.md` Layer 3 for where
it sits in the overall module graph):

```
┌─ syntax/        tokenizer → parser → AST
├─ compile/       binder, dependency analysis, compiled form
├─ runtime/       evaluator, function registry, RuntimeValue
├─ functions/     433 function implementations (11 category files)
├─ materialize/   spill engine, ghost-cell tracking, writeback plan
└─ integration/   workbook adapter, snapshot, calculate-formulas entry
```

All layers depend on structural interfaces from `materialize/types.ts`
(`WorkbookLike`, `WorksheetLike`, `CellLike`, …) — not on the concrete
`Workbook`/`Worksheet`/`Cell` classes in the excel module. Any host
that implements those interfaces can drive the engine.

## API Surface

### `installFormulaEngine(): void`

Wires `calculateFormulas` into `Workbook.calculateFormulas()` and into
`excelToPdf()`'s automatic pre-render recalculation, and installs a
default syntax probe so XLSX defined names classify strictly. Safe to
call multiple times (last registration wins). Required only if you use
the excel module's `Workbook` class.

### `uninstallFormulaEngine(): void`

Symmetric reset of both slots populated by `installFormulaEngine()` —
engine and default probe. After calling this, `Workbook.calculateFormulas()`
throws and defined-name classification falls back to the conservative
"opaque" path. Primarily useful for tests that exercise the cold-start
classification path.

### `calculateFormulas(workbook: WorkbookLike): void`

Functional entry — the core of the **standalone** mode. Walks every
formula cell in `workbook`, evaluates it with full dependency
resolution, writes results back onto each cell's `result` property, and
materialises dynamic-array spills onto ghost cells. Mutates the
workbook in place. Zero global side effects; safe for concurrent calls
on different workbooks. Works on any `WorkbookLike` — the excel module
is not required.

### `createFormulaSyntaxProbe(): SyntaxProbe`

Build a standalone tokenizer+parser probe — a function that returns
`true` when its string argument parses as a formula expression.
Injecting this via `new Workbook({ formulaSyntaxProbe })` or
`new DefinedNames(probe)` makes defined-name classification strict for
that instance **without** touching process-global state. Useful for
tests and multi-host scenarios.

### `tokenize(source: string): Token[]`

Pure lexer — accepts a formula string (with or without leading `=`) and
returns a flat token stream. Throws on invalid characters.

### `parse(tokens: Token[]): AstNode`

Pratt parser — builds a typed AST from a token stream. Throws on
structural errors.

### Structural types

`WorkbookLike`, `WorksheetLike`, `CellLike`, `RowLike`, `CellErrorValueLike`,
`FormulaResultLike`, `DefinedNameEntry`, `DefinedNamesLike`,
`DimensionsLike`, `SpillRegion`, `SyntaxProbe`.

## Compatibility Notes

- **Date system** — honours the workbook's 1900 / 1904 setting, including
  the 1900 leap-year bug.
- **Error propagation** — matches Excel's precedence (`#N/A > #VALUE! > ...`).
- **Implicit intersection** — applied at non-dynamic-array sites, exactly
  where Excel 365 applies it.
- **Iterative calc** — disabled by default; enable via
  `workbook.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 }`.
- **External references** — `[book.xlsx]Sheet!A1` is parsed as `#REF!`;
  cached values are not followed.

## License

MIT — same as the parent package.
