# documonster 1.0 — Architecture & Refactor Design

Status: design document, the blueprint for the documonster → documonster rewrite.
This is the **single authoritative blueprint**. It absorbs the empirical
measurements, per-module audit, and design-pattern examples from
`FUTURE_ARCHITECTURE.md`, and supersedes its staged v10→v13 + codemod plan with
a clean-break 1.0. `FUTURE_ARCHITECTURE.md` is retained as a historical record
of the gradual-migration design; this document governs the actual work.

Last updated: 2026-06-12.

---

## 0. Goals

1. **Eliminate the tree-shaking disease.** Importing the core today drags
   ~2.1 MB into a consumer bundle (chart alone = 1.2 MB) even when the consumer
   only reads/writes cells. Target: core entry under ~200 KB; chart/pdf/formula
   only present when the consumer imports their subpath.
2. **One consistent usage style across all modules.** No more "Excel uses
   `new Workbook()`, Word uses `Document.create()`, PDF mixes
   `new PdfDocumentBuilder()` and `PdfEditor.load()`, Archive uses
   `zip().add()`." Everything follows the same shape.
3. **Zero `host-registry` / zero `install*()`.** No global mutable registration
   slots. Dependencies travel with data, not with module-level side effects.
4. **Clean break, no compatibility layer.** documonster 1.0 is a brand-new
   package; it does not preserve the documonster class API. No codemod, no
   deprecated re-exports. The old `documonster` package is frozen
   and deprecated, pointing users to documonster.

The hard rules from `AGENTS.md` still hold: zero runtime dependencies, prefer
native APIs, no circular imports, named exports only, respect layer direction.

---

## 0a. Baseline measurement & per-module audit

Empirical baseline. The original `FUTURE_ARCHITECTURE.md` measured **2146 KB**
for `new Workbook()` — that was _before_ the chart host-registry landed. A fresh
phase-0 measurement on the current tree (esbuild 0.28.0, `--bundle --minify
--tree-shaking --platform=node`, built `dist/esm`) shows the registry already
moved most chart code out of the core path:

| Import path                            | Consumer bundle (measured now) |
| -------------------------------------- | -----------------------------: |
| `new Workbook()` + write a cell (core) |                     **765 KB** |
| `parseCsv` (csv subpath)               |                        22.8 KB |
| `tokenize, parse` (formula subpath)    |                        15.6 KB |
| `import *` everything (worst case)     |                        1153 KB |

Key correction to earlier assumptions:

- The chart **host-registry already achieved most of the bundle isolation**:
  core dropped from 2146 KB → 765 KB. So the zero-register rewrite's main payoff
  is **architectural cleanliness** (no install slot, pure-data handle), **not**
  a large additional size cut. This is an honest re-scoping of the chart work's
  value.
- The remaining 765 KB is dominated by the **xlsx read/write engine + worksheet/
  workbook**, not chart. Therefore the original "< 200 KB core" target is **not
  realistic** — that mass is core spreadsheet functionality. A realistic target
  after detaching CSV/Markdown/pivot is **~400–550 KB** (confirmed/adjusted in
  §6 against phase-0 numbers).
- `formula` at 15.6 KB is already near-perfect — the reference.

Per-module audit (✅ already good / ⚠️ minor work / ❌ major rewrite):

| Module   |     LoC | Tree-shake | Verdict                                                              | Action in this plan                            |
| -------- | ------: | ---------: | -------------------------------------------------------------------- | ---------------------------------------------- |
| formula  |       — |       3.9% | ✅ reference                                                         | delete its registry (§2.1); flat named exports |
| excel    | 109,618 |      40.6% | ❌ full rewrite                                                      | phases 4–5                                     |
| chart    |  30,278 |        n/a | ❌ tree-shakeable excel feature (flat exports, bytes-on-handle §2.3) | phases 4–5                                     |
| word     |    ~20k |      28.4% | ⚠️ partial                                                           | phase 6 flat-export align                      |
| pdf      |       — |          — | ⚠️ public API → functions                                            | phase 6                                        |
| archive  |       — |          — | ✅ stream/error classes legit                                        | phase 6 flat-export wrap only                  |
| csv      |       — |      15.2% | ✅ mostly right                                                      | phase 3 flat exports                           |
| stream   |       — |      13.0% | ✅ native primitives                                                 | phase 6 flat-export wrap only                  |
| markdown |   1,695 |          — | ✅ clean                                                             | phase 3 flat exports                           |
| xml      |   5,063 |      37.4% | ⚠️ audit barrel                                                      | phase 6                                        |
| utils    |       — |          — | ✅ free functions                                                    | no change                                      |

Modules marked ✅ need only flat-named-export wrapping for API consistency
(§1.2), not structural rewrites. The heavy lifting is excel + chart.

---

## 1. The unified public API shape

### 1.1 Opaque handles + flat named exports (empirically chosen)

Every domain object is an **opaque handle** (a branded type). Consumers never
`new` anything. They call a `create()` factory and pass the handle to free
functions.

```ts
import { createWorkbook, addSheet, setCellValue, setCellStyle, writeXlsx } from "documonster/excel";

const wb = createWorkbook();
const ws = addSheet(wb, "Sheet1");
setCellValue(ws, "A1", "Hello");
setCellStyle(ws, "A1", { font: { bold: true }, fill: { type: "solid", color: "FFFF00" } });
await writeXlsx(wb, "out.xlsx");
```

The handle is opaque to the consumer. Internally it is a plain mutable record
(see §4). The public contract is: "pass the handle to a free function."

### 1.1a Why flat named exports — and the namespace caveat (empirical)

The earlier draft assumed **object-literal namespaces** (`export const Cell =
{ setValue, setStyle, … }`) would tree-shake by property. **A bundler spike
disproved this.** Measured with esbuild 0.28.0 and rolldown 1.1.0 (the project's
actual browser bundler), `bundle + minify + tree-shaking=true`, one used method

- two unused:

| Export style                            | esbuild: unused dropped?            | rolldown: unused dropped?       |
| --------------------------------------- | ----------------------------------- | ------------------------------- |
| **Object literal** `const Cell = {…}`   | ❌ no (whole object kept)           | ❌ no (whole object kept)       |
| **ESM ns re-export** `export * as Cell` | ❌ no (esbuild materializes the ns) | ✅ yes (byte-identical to flat) |
| **Flat named exports**                  | ✅ yes                              | ✅ yes                          |

Conclusions, now binding:

1. **Object-literal namespaces are forbidden** — an object literal is one
   indivisible runtime value; no bundler can prove `.setStyle` is unreachable.
   This is dead-code by construction. (This also means the root-cause analysis
   stands: classes _and_ object-literal namespaces both defeat tree-shaking.)
2. **Flat named exports are the only style that tree-shakes on every bundler**,
   and they match the `AGENTS.md` "Named exports only" rule. This is the
   baseline for all public API.
3. The `Foo.bar()` dot-ergonomics is **optional sugar** via
   `export * as Excel from "./excel"` (ESM namespace re-export, _not_ an object
   literal). It tree-shakes perfectly on rolldown; on esbuild it degrades to
   keeping the whole namespace. Therefore: ship **flat named exports as the
   canonical surface**; a namespace alias may be offered as a convenience but
   must never be the only way to reach a function, or esbuild consumers lose
   tree-shaking.

**Implementation pattern (branded handle + flat named exports):**

```ts
// brand.ts — opaque type, consumers cannot construct it directly.
declare const WorkbookBrand: unique symbol;
export type Workbook = { readonly [WorkbookBrand]: true };

// excel/workbook.ts — flat named exports
export function createWorkbook(): Workbook {
  return _internalCreateWorkbook() as unknown as Workbook;   // mutable record inside
}
export function addSheet(wb: Workbook, name: string): Sheet { ... }

// excel/cell.ts — flat named exports
export function setCellValue(ws: Sheet, addr: string, value: CellValue): void { ... }
export function getCellValue(ws: Sheet, addr: string): CellValue { ... }
export function setCellStyle(ws: Sheet, addr: string, style: Partial<Style>): void { ... }
export function cellInfo(ws: Sheet, addr: string): CellInfo { ... }   // lazy-getter object, §1.4

// excel/index.ts — re-export flat; optional dot-sugar alias for rolldown users
export * from "./workbook";
export * from "./cell";
// optional: export * as Excel from "./surface";  // sugar only, never the sole path
```

Benefit: bundler proves `setCellStyle` unused and drops it on **every** bundler;
matches `AGENTS.md`; no object-literal dead-code trap.

### 1.2 Naming conventions (applies to every module)

Flat named exports. The verb leads, the noun follows (so related functions are
discoverable by prefix in autocomplete: `setCell…`, `getCell…`, `addSheet…`).

| Concern                             | Convention                                    | Example                                               |
| ----------------------------------- | --------------------------------------------- | ----------------------------------------------------- |
| Create a handle                     | `create<Thing>()`                             | `createWorkbook()`, `createDocument()`, `createPdf()` |
| Mutate / add                        | `<verb><Thing>(handle, ...)`                  | `addSheet(wb, name)`, `setCellValue(ws, addr, v)`     |
| Read one scalar                     | `get<Thing><Field>(handle, ...)`              | `getCellValue(ws, "A1")`, `getCellText(ws, "A1")`     |
| Read aggregated info                | `<thing>Info(handle, ...)` lazy-getter object | `cellInfo(ws, "A1")` (see §1.4)                       |
| Serialize to bytes (cross-platform) | `to<Format>Buffer(handle)`                    | `toXlsxBuffer(wb)`                                    |
| Write to file (Node-only)           | `write<Format>(handle, path)`                 | `writeXlsx(wb, path)`                                 |
| Read from bytes (cross-platform)    | `load<Format>(bytes)`                         | `loadXlsx(bytes)`                                     |
| Read from file (Node-only)          | `read<Format>File(path)`                      | `readXlsxFile(path)`                                  |
| Parse text → handle                 | `read<Format>(text)`                          | `readCsv(text)`, `readMarkdown(text)`                 |

(Optional dot-sugar via `export * as Excel` is allowed per §1.1a but is never the
sole path to a function.)

### 1.3 Style API: collapse to `setCellStyle`

Today Cell/Row/Column expose ~18 getter/setters each (`numFmt`, `font`,
`alignment`, `border`, `fill`, `protection`). That is ~54 API points just for
style. documonster collapses these to one function per target:

```ts
setCellStyle(ws, "A1", { font, fill, border, alignment, numFmt, protection });
setRowStyle(ws, 3, { ... });
setColumnStyle(ws, "B", { ... });
```

Partial styles merge with existing. Reading style: `cellInfo(ws, "A1").style`.

### 1.4 Read info: lazy-getter info object (NOT an eager aggregate)

An earlier draft proposed an eager `cellInfo` returning all ~25 fields eagerly.
That is **wrong for tree-shaking and for runtime cost**: computing `displayText`
/ `html` / `formula` for every read pulls in formatting and HTML logic even when
the caller only wanted the value. Two correct options, used together:

1. **Scalar getters for the hot path** — `getCellValue`, `getCellText`,
   `getCellType`, `getCellFormula`. Each is independently tree-shakeable; the
   bundler drops `getCellHtml` if unused.
2. **`cellInfo(ws, addr)` returns an object with lazy getters** for the
   convenience case. Fields like `displayText` / `html` are defined as
   `get displayText()` that compute on first access, so reading `.value` never
   triggers HTML rendering. For maximum tree-shaking, prefer the scalar getters;
   `cellInfo()` is the ergonomic escape hatch.

This corrects the naive "1 fat function is smaller" assumption: a fat aggregate
defeats tree-shaking inside the function body.

### 1.5 Subpath = module, not feature (corrected)

**Top-level subpaths map to modules, never to a feature inside a module.**
`table` / `image` / `validation` / `comment` / `form` / `pivot` / `chart` are
**Excel features**, so they live inside the excel module's surface — not as
`documonster/table` siblings of `documonster/word`. The subpath granularity is
the nine modules:

```
documonster            main entry (re-exports the excel surface for convenience)
documonster/excel      the full Excel surface (workbook, sheet, cell, table,
                       image, validation, comment, form, pivot, chart, ...)
documonster/word
documonster/pdf
documonster/csv
documonster/markdown
documonster/xml
documonster/archive
documonster/stream
documonster/formula
```

Tree-shaking _within_ the excel module is achieved by **flat named exports**
(§1.1a), not by splitting each Excel feature into its own npm subpath. A
consumer that imports `Workbook` + `Cell` and never references `addChart` /
`addTable` gets those dropped by the bundler — that is the whole point of flat
named exports, and it is what the empirical spike (§1.1a) confirms.

The excel module's public surface — all **flat named exports** (grouped below by
topic only for readability; each row is a set of top-level functions, not a
namespace object):

```
workbook:  createWorkbook, addSheet, removeSheet, getSheet, listSheets, importSheet, protectWorkbook, ...
sheet:     addRow, addRows, insertRow, getRow, eachRow, getColumn, mergeCells,
           addTable, addImage, addNote, addValidation, addChart, addPivotTable, sheetToJson, addJson, ...
row:       setRowStyle, getRowValues, eachCell, ...
column:    setColumnStyle, setColumnWidth, ...
cell:      setCellValue, getCellValue, getCellText, getCellType, getCellFormula, setCellStyle, cellInfo, ...
io:        loadXlsx, toXlsxBuffer,                  // cross-platform
           readXlsxFile, writeXlsx,                 // Node-only conditional export
           createStreamWriter, createStreamReader   // streaming, kept
plus enums, errors, address helpers, types
```

#### 1.5a Public API: dot-namespaces over physical surface modules (binding)

The public surface is organised as **domain dot-namespaces** — `Workbook.`,
`Worksheet.`, `Cell.`, `Row.`, `Column.`, plus one namespace per Excel feature
(`Chart.`, `Table.`, `Image.`, `Pivot.`, `Sparkline.`, `Form.`, `Validation.`,
`Note.`, `Range.`). This mirrors Word's `Document.` surface.

```ts
import { Workbook, Cell, Chart } from "documonster/excel";

const wb = Workbook.create();
const ws = Workbook.addSheet(wb, "Sheet1");
Cell.setValue(ws, "A1", "Hello");
Cell.setStyle(ws, "A1", { font: { bold: true } });
Chart.addColumn(ws, { ... }, "D1:K20");
const buf = await Workbook.toXlsxBuffer(wb);
```

**Mechanism — empirically chosen (esbuild 0.28 + rolldown 1.1 + rspack 2.0,
re-verified):** each namespace is an **ESM namespace re-export**
(`export * as Cell from "./surface/cell"`) over a **physical surface module**
of flat functions. Measured tree-shaking (one used member, two unused, through
the `import { Cell } from "documonster/excel"` subpath then `Cell.member()`):

| Export style                              | esbuild | rolldown | rspack |
| ----------------------------------------- | ------- | -------- | ------ |
| flat named exports (internal)             | ✅      | ✅       | ✅     |
| **`export * as Ns` (the chosen surface)** | ❌      | ✅       | ✅     |
| object literal `const Ns = {…}`           | ❌      | ❌       | ❌     |
| class with methods                        | ❌      | ❌       | ❌     |
| TS `function f(){} namespace f{}` merge   | ❌      | ❌       | ❌     |

documonster targets **rolldown (browser) + rspack**; esbuild degradation is
accepted (documented). `export * as` is byte-identical to flat on rolldown/rspack
and **drops unused members**. Object-literal / class / namespace-merge are
forbidden — they are indivisible runtime values no bundler can tree-shake. This
matches every modern size-sensitive library surveyed (es-toolkit, valibot,
date-fns, zod: flat + `export *`; only size-indifferent backend ORMs like
drizzle use namespace-merge).

**Single style, no split.** Because the public surface is _only_ the
dot-namespaces, there is exactly one way to write each call (`Cell.setValue(…)`).
The flat functions are an internal implementation detail and are NOT re-exported
bare from the subpath, so consumers cannot mix `setCellValue` and
`Cell.setValue`. (Cost: no bare-function import; that is the price of a single,
uniform style — accepted.)

**Physical layering (binding).** Classify by meaning first, then modularise:
every namespace maps to one physical `surface/*.ts` file; the subpath
`documonster/excel` re-exports each via `export * as`. Each surface file
re-exports (and de-prefixes) the underlying flat implementation functions. No
namespace object is ever materialised in source.

```
src/modules/excel/
  surface/
    workbook.ts    export { create, addSheet, getSheet, removeSheet, eachSheet,
                            toXlsxBuffer, loadXlsx, readXlsxFile, writeXlsx,
                            protect, unprotect, addExternalLink, … }
    worksheet.ts   export { merge, unmerge, splice, insertRow, autoFit,
                            dimensions, rowCount, columnCount, toJson, addJson, … }
    cell.ts        export { setValue, getValue, setStyle, getStyle, getText,
                            getType, getFormula, getResult, setNote, getNote,
                            setHyperlink, addName, setValidation, … }
    row.ts         export { setHeight, getHeight, setHidden, setStyle, getValues }
    column.ts      export { setWidth, getWidth, setHeader, setKey, setHidden, setStyle }
    chart.ts       export { add, addColumn, addBar, addLine, …, get, remove }
    table.ts       export { add, get, list, remove }
    image.ts       export { add, list, setBackground }
    pivot.ts       export { add }
    sparkline.ts   export { add, list, remove }
    form.ts        export { addCheckbox, listCheckboxes }
    range.ts       export { create, contains, intersects, forEachAddress }
  index.ts         export * as Workbook  from "./surface/workbook";
                   export * as Worksheet from "./surface/worksheet";
                   export * as Cell      from "./surface/cell";
                   export * as Row       from "./surface/row";
                   export * as Column    from "./surface/column";
                   export * as Chart     from "./surface/chart";
                   export * as Table     from "./surface/table";
                   export * as Image     from "./surface/image";
                   export * as Pivot     from "./surface/pivot";
                   export * as Sparkline from "./surface/sparkline";
                   export * as Form      from "./surface/form";
                   export * as Range     from "./surface/range";
                   export type { Workbook as WorkbookHandle, ... };  // handle types
```

**Namespace member naming.** Inside a namespace, members drop the redundant
domain prefix (the namespace already supplies it): `cellSetValue` → `Cell.setValue`,
`addWorksheet` → `Workbook.addSheet`, `mergeCells` → `Worksheet.merge`,
`addColumnChart` → `Chart.addColumn`, `setRowHeight` → `Row.setHeight`.

**Two implementation layers (binding):**

1. **Public `(ws, addr, …)` flat functions** — `setCellValue(ws, addr, v)` =
   `cellSetValue(getCell(ws, addr), v)`. These compose the handle-level helpers
   and are what `surface/cell.ts` re-exports as `setValue`. This layer is what
   the test codemod migrates tests onto.
2. **Internal handle helpers** (`cellSetValue(cellData, v)`, `getCell(ws, addr)`,
   the value-boxing classes, the `*-core` container modules) — never exposed
   through a namespace.

**Authoritative member map (codemod contract).** `ws` = sheet handle, `wb` =
workbook handle, `a` = address.

| Today (handle access)             | documonster namespaced call                                  |
| --------------------------------- | ------------------------------------------------------------ |
| `new Workbook()`                  | `Workbook.create()`                                          |
| `wb.addWorksheet(name)`           | `Workbook.addWorksheet(wb, name)`                            |
| `wb.getWorksheet(id)`             | `Workbook.getWorksheet(wb, id)`                              |
| `wb.addChartsheet(...)`           | `Workbook.addChartsheet(wb, ...)`                            |
| `wb.eachSheet(cb)`                | `Workbook.eachSheet(wb, cb)`                                 |
| `await wb.protect(pw)`            | `Workbook.protect(wb, pw)`                                   |
| `wb.xlsx.writeBuffer()`           | `Workbook.toXlsxBuffer(wb)`                                  |
| `wb.xlsx.load(bytes)`             | `Workbook.loadXlsx(wb, bytes)`                               |
| `wb.xlsx.readFile(path)`          | `Workbook.readXlsxFile(wb, path)`                            |
| `wb.xlsx.writeFile(path)`         | `Workbook.writeXlsx(wb, path)`                               |
| `ws.getCell(a).value` (read)      | `Cell.getValue(ws, a)`                                       |
| `ws.getCell(a).value = v` (write) | `Cell.setValue(ws, a, v)`                                    |
| `ws.getCell(a).text`              | `Cell.getText(ws, a)`                                        |
| `ws.getCell(a).type`              | `Cell.getType(ws, a)`                                        |
| `ws.getCell(a).formula`           | `Cell.getFormula(ws, a)`                                     |
| `ws.getCell(a).result`            | `Cell.getResult(ws, a)`                                      |
| `ws.getCell(a).{font,fill,…}`     | `Cell.getStyle(ws, a).{…}` / `Cell.setStyle(ws, a, {…})`     |
| `ws.getCell(a).note`              | `Cell.getNote(ws, a)` / `Cell.setNote(ws, a, …)`             |
| `ws.getCell(a).dataValidation`    | `Cell.getValidation(ws, a)` / `Cell.setValidation(ws, a, …)` |
| `ws.getCell(a).master`            | `Cell.getMergeMaster(ws, a)`                                 |
| `ws.getRow(n).height`             | `Row.getHeight(ws, n)` / `Row.setHeight(ws, n, h)`           |
| `ws.getColumn(k).width`           | `Column.getWidth(ws, k)` / `Column.setWidth(ws, k, w)`       |
| `ws.mergeCells(range)`            | `Worksheet.merge(ws, range)`                                 |
| `ws.addRow(vals)`                 | `Worksheet.addRow(ws, vals)`                                 |
| `ws.eachRow(cb)`                  | `Worksheet.eachRow(ws, cb)`                                  |
| `await ws.protect(pw)`            | `Worksheet.protect(ws, pw)`                                  |
| `wb.addImage(payload)` (register) | `Image.add(wb, payload)` → imageId                           |
| `ws.addImage(id, range)` (place)  | `Image.place(ws, id, range)`                                 |
| `ws.addBackgroundImage(id)`       | `Image.setBackground(ws, id)`                                |
| `ws.addTable(model)`              | `Table.add(ws, model)`                                       |
| `ws.addChart(opts, range)`        | `Chart.add(ws, opts, range)`                                 |
| `ws.addColumnChart(opts, range)`  | `Chart.addColumn(ws, opts, range)`                           |
| `ws.addPivotTable(model)`         | `Pivot.add(ws, model)`                                       |

**Handle types** — exposed alongside the namespaces; the exact form (`Workbook.Handle`
inside the namespace vs a top-level `WorkbookHandle` type) is **TBD, decided last**
(does not block surface implementation). The surface functions are typed against
the existing `WorkbookData` / `WorksheetData` records internally regardless.

(Note: `protect` is `protect` inside both `Workbook.` and `Worksheet.` — the
namespace supplies the disambiguating context, so no `protectWorkbook` /
`protectSheet` suffix is needed. The internal flat implementation may still use
suffixed names; the surface module re-exports them de-suffixed.)

---

## 2. Zero host-registry — how each coupling is removed

The single hardest design question: today the core write engine reaches **back**
into optional modules (chart, formula) through a `host-registry` slot. The
"completely consistent + perfectly tree-shaken" goal requires removing all of
these. Below is the per-coupling resolution, grounded in the actual code.

### 2.1 Formula — delete the registry outright (zero cost)

Findings from code audit:

- formula never imports excel; there is no reverse dependency.
- `calculateFormulas(wb: WorkbookLike)` **already exists** as a self-contained
  free function (`formula/integration/calculate-formulas`) that does **not** go
  through the registry.
- No serialization/deserialization path ever auto-invokes the formula engine.
  Today's `Workbook.calculateFormulas()` method is a pure user-triggered thin
  wrapper around `invokeFormulaEngine(this)`.

Resolution:

- Delete `formula/host-registry.ts`, `formula/install.ts`,
  `installFormulaEngine()`, `invokeFormulaEngine()`. There is **no**
  `calculateFormulas` method on the workbook handle in documonster.
- Public API is the flat named function from the subpath:
  `import { calculateFormulas } from "documonster/formula"; calculateFormulas(wb);`
- Consumers who never import `documonster/formula` cannot reach the ~200 KB
  engine. Perfect tree-shaking, no slot.

One residual: `default-syntax-probe` is used **inside the XLSX read path** by
the defined-name parser to classify a defined-name as formula vs opaque.
Resolution: keep it as a **factory/option injection** (`createWorkbook({
formulaSyntaxProbe })`), drop the process-global default slot. Fallback when no
probe is supplied: treat all defined-names as opaque (byte round-trip stable;
just not evaluable). This is safe and already the behavior.

### 2.2 PDF formula recalc — explicit option, no registry

Today `excelToPdf()` calls `tryInvokeFormulaEngine(wb)` to recalc before export.
Resolution: `xlsxToPdf` (the documonster name) takes an explicit injected
recalculator so the engine is only pulled in when the caller asks:

```ts
import { calculateFormulas } from "documonster/formula";
await xlsxToPdf(wb, { recalculate: calculateFormulas }); // engine in bundle only here
await xlsxToPdf(wb); // uses cached XLSX values, no engine
```

### 2.3 Chart — rendered bytes travel with the handle (pure data, zero register)

**Decision rationale (with honest measurement).** Phase-0 measurement proved the
existing host-registry _already_ tree-shakes chart perfectly: a `new Workbook()`
that never imports chart is 765 KB with **0 KB of chart code**; importing chart
adds +367 KB. So the zero-register rewrite below yields **0 KB of additional
size benefit** over keeping the registry. Three options were on the table:

- **A — keep registry + manual `installChartSupport()`** (status quo): 0 work,
  but users can forget to install and hit a runtime error.
- **B — keep registry, auto-register on `import "documonster/excel/chart"`**:
  tiny change, same bytes, removes the footgun. Lowest risk.
- **C — delete the registry; chart renders bytes stored on the handle** (this
  section): highest work and risk, 0 size benefit, but the cleanest end state —
  no registry mechanism at all, and the workbook handle becomes pure
  serializable data.

**C is chosen deliberately** because documonster 1.0 is the **one-time window**
where breaking, deep changes to the excel internals are acceptable. A registry
left in place becomes a permanent fixture that can never be removed without
another breaking release; doing the de-register now, while the whole excel
module is already being rewritten, is the only chance. The phase-1 spike (§5)
de-risks the single hard assumption (byte-identical output); if the spike fails,
fall back to **B** — same bytes, no risk.

Chart is the hard one. The write engine (`xlsx.browser.ts`) reaches into chart
in ~12 places. Audit classified them:

- **Pure render (easy):** `renderChartEx(model, ctx)`, `buildChartStyle`,
  `buildChartColors`, `themeIndexToName`, `renderChartExLegendXml`. These are
  (or are nearly) pure `model → bytes`. `renderChartEx` does not touch the
  workbook today.
- **Defined-name rewrite (medium):** `rewriteChartExDataRefsToDefinedNames`
  registers hidden `_xlchart.vN.M` names. The audit confirmed **all its inputs
  are available at `addChart` time** (the chartEx index is assigned by
  `nextChartExNumber()`, formulas are already sheet-qualified, the
  `DefinedNames` instance exists from workbook construction). It does **not**
  depend on final sheet order or the final global defined-name table.
- **Raw-patch / passthrough decision tree (hard):** ~600 lines of
  `buildRawChartEx*` + `tryPatchChartExRawXml` + `spliceChartExLeadingComments`
  encode chartEx XML knowledge.
- **Read path (hard):** `parseChartEx`, `createChartSpaceXform`.

#### The chosen resolution: store rendered bytes, not closures

Three storage options were weighed (this is the design decision the user asked
to get right, not cut corners on):

| Option                                                           | Zero register? | Handle stays pure data?                                               | Verdict    |
| ---------------------------------------------------------------- | -------------- | --------------------------------------------------------------------- | ---------- |
| A. Store a `render(ctx)` **closure** on the handle               | ✅             | ❌ closure breaks `structuredClone` / worker transfer / serialization | rejected   |
| B. Chart keeps a `WeakMap`, core reads a tiny export point       | ✅ handle pure | ❌ that export point **is** a register by another name                | rejected   |
| C. Store **already-rendered bytes** (`Uint8Array`) on the handle | ✅             | ✅ bytes are pure serializable data                                   | **chosen** |

**Option C — rendered bytes travel with the handle:**

1. `addChart(ws, ...)` (a flat named export of the excel module; chart is an
   Excel feature per §1.5) renders the chart to bytes **immediately, in place**,
   and stores pure data on the handle:
   `ws.chartParts: Array<{ path: string; bytes: Uint8Array; sidecars: {...} }>`.
   The defined-name rewrite also runs now (its inputs are all available per the
   audit), registering `_xlchart.vN.M` into the workbook's defined names — also
   pure data on the handle.
2. The core write engine, serializing, just **copies `ws.chartParts` bytes into
   the zip** at the recorded paths. It references **no** chart symbol, holds
   **no** function. The handle is 100% serializable data.

   Key point: the chart _implementation_ lives in the excel module's chart
   files; `addChart` and its renderers are flat named exports. A consumer who
   never imports `addChart` (or any chart export) gets all ~1.2 MB of chart code
   tree-shaken — the core write loop only ever touches `ws.chartParts` as bytes.

3. **Mutation correctness** (the only real subtlety): if the user edits a chart
   after creating it, they go through the chart functions, and each edit
   re-renders the affected `chartPart` bytes. The core always sees current
   bytes. This trades the old "render at write time (last-moment snapshot)"
   semantics for "render at mutation time" — which the audit noted is actually
   _more_ idempotent (rewrite runs once per edit, not once per write).
4. **Read path** is symmetric: the core engine reads chart parts as **opaque
   bytes** into `ws.chartParts` (raw, unparsed). Only `getCharts(ws)` parses
   those bytes on demand. `parseChartEx` / `createChartSpaceXform` are reachable
   only through the chart exports and never run during core load.

Result: a consumer who never imports any chart export produces a workbook whose
`chartParts` is empty on write and carries raw bytes on read; the core never
references any chart symbol; the handle is pure serializable data (works with
`structuredClone`, worker transfer, streaming); the bundler tree-shakes all
~1.2 MB of chart code. No registry, no slot, no `install()`, no
closure-on-handle hazard.

Honest note: moving the ~600-line raw-patch/passthrough decision tree out of the
core write engine and behind the chart exports (so rendering — including patch
decisions — happens at chart-function call time and produces final bytes) plus
the lazy read path are the largest, highest-regression part of the project. They
are sequenced last (phase 5) so the rest of the wins land first. The chart
round-trip test corpus is the safety net: rendered output must be
byte-identical, only _when_ it is produced changes.

### 2.4 Pivot, CSV, Markdown, JSON/AOA — pure extraction (no reverse call)

These have **no** reverse call from the core write engine; they are just user
entry points that happen to live as `Workbook`/`Worksheet` methods today.
Resolution: pivot/JSON-AOA become flat excel-module exports (Excel features);
CSV/Markdown become their own module subpaths.

```ts
import { readCsv } from "documonster/csv"; // readCsv(text) -> wb
import { readMarkdown } from "documonster/markdown";
import { addPivotTable, sheetToJson } from "documonster/excel";

const wb = readCsv(text);
addPivotTable(ws, model);
```

Because the core namespaces never reference these functions, importing the
core does not pull them in. Zero registry, lowest risk, biggest immediate win.

---

## 3. Node vs Browser

The only core type with a `.browser` split is the workbook (file-path APIs).
documonster keeps a two-layer flat-function design:

- Cross-platform: `loadXlsx(bytes): Workbook`, `toXlsxBuffer(wb): Promise<Uint8Array>`.
- Node-only (conditional export from `documonster/excel`): `readXlsxFile(path)`,
  `writeXlsx(wb, path)`, streaming with file paths.

IO signatures use Web Platform primitives (`Uint8Array`, `ReadableStream`)
per `FUTURE_ARCHITECTURE.md` §6.4. Node users bridge with `Readable.toWeb()`.

---

## 4. Internal implementation: also de-class

**Decision (confirmed):** the internal implementation is also rewritten away
from classes to plain data records + module-internal functions — not merely
wrapped. This is the most expensive choice (the xlsx engine is deeply coupled to
class instances) but yields the cleanest final form and removes any risk of a
class instance leaking into the tree-shaking graph.

The trade-off was explicitly weighed: a cheaper alternative is "flat functions
on the outside, keep classes inside," which captures ~95% of the tree-shaking
win at much lower risk. It was **rejected** in favor of the fully de-classed
internals because the goal is the cleanest 1.0, not the smallest diff.

- The opaque public `Workbook` type (§1.1) is a brand over an internal mutable
  record, e.g. `interface WorkbookData { sheets: ...; definedNames: ...;
chartParts: ...; ... }`. `createWorkbook()` returns a `WorkbookData` typed as
  the branded `Workbook`. There is one runtime object; the brand only hides its
  shape from consumers.
- Operations are module-internal free functions over these records.
- The XLSX read/write engine (`xlsx/`) is rewritten to consume the record
  shapes instead of class instances.

Legitimate class uses are exempt (per `AGENTS.md` "prefer native APIs" pragmatism):
stream subclasses, parser state machines, and `BaseError` subclasses stay
classes — de-classing applies to the **Excel domain model** (Workbook/Sheet/
Cell/Row/Column/Table/…), not to every class in the codebase.

Scope reality: this touches the entire excel module (~110k LoC, of which the
xlsx engine is the hardest part). It is staged and incremental (§5) so each step
is independently verifiable, with the full test suite and benchmark as the
safety net.

---

## 4a. Error strategy (decide once, at 1.0)

1.0 sets the error contract once, and it is **uniform across the whole project**:
**throw**, do not return `Result`. An earlier draft proposed a `Result`
discriminated union for parser/validation failures; that was rejected because it
would create a two-style API alongside the existing `BaseError` hierarchy and
force a churn-heavy rewrite of every current parser boundary for marginal
benefit.

The single rule:

- **All failures throw a `BaseError` subclass** (`@utils/errors`) with `{ cause }`
  chaining, exactly as the codebase does today. Programmer errors (type
  mismatch, out-of-bounds address, null deref) and expected failures (malformed
  CSV/XML, invalid defined name, missing cell) both throw — they differ only in
  the error subclass, not in the mechanism.
- Callers use `try/catch` and may discriminate on `instanceof` (e.g.
  `CsvParseError`, `XmlParseError`, `InvalidAddressError`). The existing
  branded error classes (`isExcelError`, `getRootCause`, `errorToJSON`) stay.

Benefit: zero migration cost from today's behavior, one consistent style, no
mix of throw-and-return across namespaces. This is the conservative,
implementation-proven choice over the trendier `Result` type.

---

## 4b. Migration guide (documonster → documonster)

Clean break does **not** mean "leave users stranded." Even without a
compatibility layer, ship a migration aid so adoption is tractable:

1. **`docs/MIGRATION_FROM_EXCELTS.md`** — a complete old→new mapping table:
   `new Workbook()` → `createWorkbook()`,
   `wb.addWorksheet(n)` → `addSheet(wb, n)`,
   `cell.value = v` → `setCellValue(ws, addr, v)`,
   `cell.font = {...}` → `setCellStyle(ws, addr, { font })`,
   `workbook.calculateFormulas()` → `calculateFormulas(wb)` (from `documonster/formula`), etc.
   The Excel API inventory already gathered during design is the source for this
   table — it must be exhaustive.
2. **Optional codemod** (`ts-morph`) for the mechanical 80%: `new X()` →
   `X.create()` and method→namespace-function rewrites. Not required for 1.0,
   but the mapping table makes it straightforward to add later.
3. The existing `docs/FROM_*.md` (FROM_EXCELJS, FROM_OPENPYXL, …) competitor
   migration guides stay and get updated to the namespace API.

Without this, a renamed package with a 100%-incompatible API gets no adoption.
The migration guide is part of the 1.0 deliverable, not an afterthought.

---

## 5. Execution roadmap (de-risk the foundation first)

Reordered from a naive "easy → hard" sequence. The riskiest assumption in the
whole design is "the ~600-line chartEx raw-patch decision tree can move into the
chart module and still produce **byte-identical** output." If that is false, the
zero-register premise for chart collapses. So we **prove the chart foundation
first** with a thin spike, before investing in the rest. Everything else is
genuinely low-risk and can follow.

| Phase     | Work                                                                                                                                                                                                                                                                                                                                                                                                                         | Risk                      | Tree-shake win        |
| --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------- | --------------------- |
| 0         | Build baseline: run analyze, record bundle size for "read-only XLSX" and each scenario. Capture golden numbers. Capture write-1M-cells benchmark baseline.                                                                                                                                                                                                                                                                   | none                      | measure               |
| 1 (SPIKE) | **Chart foundation spike**: take ONE chartEx round-trip fixture; move its raw-patch/passthrough + render path so bytes are produced at chart-function call time and stored as pure data; prove output is byte-identical to current. Throwaway-quality is fine — the goal is to validate the premise, not ship. **If this fails, revisit the chart design (maybe a minimal slot after all) before committing to phases 4–5.** | highest (foundation)      | proves chart approach |
| 2         | Delete formula registry; `documonster/formula` = flat named exports (`calculateFormulas`, …); probe via factory injection; PDF recalc via explicit option.                                                                                                                                                                                                                                                                   | low                       | formula off core path |
| 3         | Move CSV / Markdown to their module subpaths; pivot / JSON-AOA become flat excel-module exports — all as flat named exports.                                                                                                                                                                                                                                                                                                 | low                       | big — CSV/MD off core |
| 4         | Excel core class→flat-named-exports (public AND internal de-class): `createWorkbook`/`addSheet`/`setCellValue`/`setCellStyle`/`get*`/IO; collapse style to `setCellStyle`; reads to scalar getters + lazy `cellInfo`; handle is a plain record carrying pure-data `chartParts`; rewrite tests.                                                                                                                               | high                      | core API consistent   |
| 5         | Apply the proven spike (phase 1) at full scale: chart exports render bytes in place, defined-name rewrite at call time; move full raw-patch tree + read path (`parseChartEx`) behind chart exports; core handles chart parts as opaque bytes. Whole-library zero-register achieved.                                                                                                                                          | high (de-risked by spike) | chart fully off core  |
| 6         | Align word/pdf/xml/archive/stream to the unified flat-named-export shape (§1.2).                                                                                                                                                                                                                                                                                                                                             | medium                    | consistency           |
| 7         | Update `treeshake-verify.ts`: assert read-only XLSX excludes csv/markdown/chart/formula/pdf subtrees. `pnpm check && build && test` all green; benchmark vs phase-0 baseline (no perf regression); produce before/after size table. Write `docs/MIGRATION_FROM_EXCELTS.md`.                                                                                                                                                  | —                         | verify                |
| 8         | Brand rename documonster → documonster (package.json name=`documonster`, homepage/bugs/repository=`documonster/documonster`, README/README_zh/AGENTS, rolldown iife filenames + global name, CI/release.yml package refs); version=`1.0.0`; reset CHANGELOG.                                                                                                                                                                 | low                       | —                     |
| 9         | GitHub: migrate repo to `documonster` org; reconfigure `NPM_TOKEN`, release-please component name `documonster`→`documonster`; publish `documonster@1.0.0`; `npm deprecate documonster "Renamed to documonster. Install: npm i documonster"`.                                                                                                                                                                                | medium                    | —                     |

Verification gate after every phase: `pnpm check` then `pnpm format`, plus the
relevant test subset, plus `treeshake-verify` once phase 3+ lands. **After
phase 4, also run the `benchmark/` suite and compare against the phase-0
baseline — flat-function dispatch must not regress the hot write path.**

---

## 6. Success metrics

| Metric                                    | Current (documonster) | documonster 1.0 target                 |
| ----------------------------------------- | --------------------- | -------------------------------------- |
| Core entry bundle (read/write cells only) | 2146 KB               | < 200 KB                               |
| chart code when not imported              | always in bundle      | 0 (tree-shaken)                        |
| Tree-shake ratio (core path)              | 40.6%                 | < 8%                                   |
| host-registry / install slots             | 2 (chart, formula)    | 0                                      |
| Public API object-creation paradigms      | 4+                    | 1 (`create*()` + flat named functions) |

---

## 7. Risks

1. **Chart phase 4–5 regressions.** The raw-patch decision tree is intricate
   (byte passthrough, strict template mode, leading-comment splicing). Mitigation:
   keep the existing chart round-trip test corpus; do not change rendering
   output, only relocate where it is invoked from.
2. **Internal de-class is large.** The xlsx engine is coupled to class
   instances. Mitigation: incremental per-phase, full test suite as the safety
   net; benchmark to catch perf regressions (record→function dispatch).
3. **Loss of IDE dot-discovery.** Flat named exports do not autocomplete off a
   dot. Mitigation: verb-noun naming groups by prefix (`setCell…`, `getCell…`,
   `addSheet…`) so prefix-typing discovers them; optional `export * as Excel`
   dot-sugar for rolldown users (§1.1a), never the sole path.
4. **Clean break breaks all existing documonster users.** Mitigation: ship
   `docs/MIGRATION_FROM_EXCELTS.md` (§4b) with an exhaustive old→new mapping;
   documonster stays frozen + deprecated pointing to documonster.
5. **`chartParts`-on-handle render timing.** Rendering at `Chart.*` mutation
   time instead of write time changes _when_ bytes are produced. Mitigation:
   the chart round-trip corpus asserts byte-identical output; every `Chart.*`
   mutation re-renders the affected part so the handle always holds current
   bytes. Verify multi-write (writing the same workbook twice) stays stable.
6. **Performance of namespace-function dispatch vs class methods.** Mitigation:
   phase-0 captures a write-1M-cells baseline; phase-3 gate re-runs the
   benchmark and must show no regression. V8 caches namespace property access;
   prior analysis (`FUTURE_ARCHITECTURE.md` §7.3) expects no measurable hot-path
   difference, but it is now an enforced gate, not an assumption.
7. **Error-style consistency.** Resolved by §4a: the whole project throws
   `BaseError` subclasses uniformly — no `Result` type, no mixed style. Zero
   migration from current behavior.
