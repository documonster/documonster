# Future Architecture — Namespace-Based Redesign

Status: design document, not yet implemented.
Owner: core team.
Last updated: 2026-05-05.

## 1. Motivation

Consumers importing the top-level `Workbook` from `excelts` get **2.1 MB** of
JavaScript in their final bundle even when they never draw a chart, never touch
PDF, and never evaluate a formula. Empirical measurement with `esbuild
--tree-shaking=true`:

| Import path                                      | Consumer bundle | Tree-shake ratio |
| ------------------------------------------------ | --------------: | ---------------: |
| `import { Workbook } from "@pkg"`                |         2146 KB |            40.6% |
| `import { parseCsv } from "@pkg/csv"`            |           54 KB |            15.2% |
| `import { tokenize, parse } from "@pkg/formula"` |           35 KB |             3.9% |
| `import { DocumentBuilder } from "@pkg/word"`    |          325 KB |            28.4% |
| `import { BufferedStream } from "@pkg/stream"`   |           15 KB |            13.0% |

The `formula` module retains 3.9% of its source in the final bundle, which is
effectively perfect tree-shaking. The `excel` module retains 40.6%, which is
dominated by unreachable code that the bundler cannot prove is dead.

Breakdown of what excel drags in (measured via `--metafile`):

```
Consumer bundle 2146 KB, of which:
  chart/              1203 KB  (56% of bundle — chart 14 files)
  worksheet+workbook   361 KB
  other xlsx/xform    ~580 KB
```

Consumers that only read/write cells pay for the whole chart renderer (SVG,
PNG, PDF), chart ex parser, chart presets, cache populator, and so on.

## 2. Root cause

`Worksheet` is a God class. 20 methods (`addChart`, `addLineChart`, `addBarChart`,
`addPieChart`, `addHistogramChart`, …) hang off `Worksheet.prototype`. Each
method body contains static ESM imports that name concrete chart runtime
symbols (`buildChartModel`, `fillChartCaches`, `new Chart(...)`). Because the
class is instantiated by the consumer, the bundler cannot statically prove that
any of these methods are unreachable. All prototype methods are retained,
therefore all of their static dependencies are retained.

This is not a limitation of `sideEffects: false` or Rolldown/esbuild
tree-shaking. It is a deliberate conservativism of static analysis: method
dispatch on an instance (`ws.addChart(...)`) is not knowable at compile time.

`sideEffects: false` is necessary but not sufficient. Free functions and
namespace objects of free functions preserve the property that the bundler can
prove a symbol is unreferenced. Class methods cannot be proven unreferenced
once the class is instantiated.

## 3. Design principles

1. **Plain data shapes, not classes, for domain types.** `Workbook`,
   `Worksheet`, `Cell`, `Row`, `Column` are opaque types (branded handles).
   Internal state may be mutable for performance; the public contract is a
   handle passed to free functions.

2. **Namespace objects of free functions** for ergonomics and logical
   grouping. `Workbook.addSheet(wb, name)` instead of `new Workbook().addSheet(name)`.

3. **Every optional capability is a subpath export.** Chart, formula, PDF
   bridge, pivot tables, slicers, timelines, form controls, data validations,
   images — each is its own subpath. The main entry stays minimal.

4. **No method-internal runtime imports from optional modules.** If a core
   type needs to invoke optional code, it goes through a host-registry
   pattern: small type-only interface in core, fat implementation registered
   by the optional module's `install()` function.

5. **Native Web Platform primitives for IO.** `ReadableStream`,
   `WritableStream`, `Uint8Array`, `AsyncIterable`. No hand-rolled
   stream types in the public API.

6. **No top-level side effects.** No module registers globals at import time.
   No `sideEffects`-breaking imports. Users call `install*()` explicitly if
   they want registration.

7. **Errors as branded discriminated unions**, not thrown exceptions, for
   expected failures (parse errors, validation errors, missing cells).
   Exceptions reserved for programmer errors (type mismatch, null deref).

8. **Small core, rich periphery.** The `@pkg` main entry should be under
   200 KB minified. Everything bigger ships from a subpath.

## 4. Current state per module

### 4.1 `formula` — reference implementation ✅

The closest to the target design. Subpath export `./formula`. Public API is
entirely free functions:

```ts
import { installFormulaEngine } from "@pkg/formula";
import { tokenize, parse, calculateFormulas } from "@pkg/formula";
```

One class exists (`EvalSession`) but it is internal state for a single
evaluation pass and is not exposed.

`installFormulaEngine()` registers implementation into a host-registry so
`Workbook.calculateFormulas()` can invoke it without the excel module
statically importing the 15,000-line function registry.

Tree-shake ratio: **3.9%**.

**Action**: none. Use as template.

### 4.2 `excel` — full rewrite ❌

| Measurement                 | Value                      |
| --------------------------- | -------------------------- |
| Source LoC                  | 109,618                    |
| Classes exported            | 22                         |
| Tree-shake ratio            | 40.6%                      |
| Bundle for `new Workbook()` | 2146 KB                    |
| chart code in that bundle   | 1203 KB (tree-shake fails) |

Problems:

- `Workbook`, `Worksheet`, `Cell`, `Row`, `Column`, `Range`, `Table`, `Image`,
  `Note`, `DataValidations`, `FormCheckbox`, `Chartsheet`, `PivotTable`,
  `DefinedNames` all exported as classes.
- `Worksheet` has 30+ methods spanning cells, rows, columns, charts, images,
  tables, pivot tables, slicers, timelines, form controls, page setup,
  hyperlinks, comments. Classic God object.
- `worksheet.ts` statically imports 9 runtime symbols from `@excel/chart/*`,
  which is why chart is unshakable.
- `workbook.browser.ts` statically imports chart similarly.
- `chartsheet.ts` statically imports `Chart` from chart module.
- Top-level barrel `src/index.ts` re-exports the entire chart public surface
  from `@excel/chart`, so even consumers who try to bypass the class have no
  dedicated `@pkg/chart` subpath.

Target design:

```ts
// @pkg  (main, small)
import { Workbook, Sheet, Cell, IO } from "@pkg";
const wb = Workbook.create();
Sheet.add(wb, "Sheet1");
Cell.set(wb, "Sheet1", "A1", "hello");
const buf = await IO.toBuffer(wb);

// @pkg/chart  (opt-in)
import { Chart } from "@pkg/chart";
Chart.add(wb, "Sheet1", { type: "line", range: "A1:B10" });

// @pkg/pivot  (opt-in)
import { Pivot } from "@pkg/pivot";
Pivot.create(wb, ...);

// @pkg/validation  (opt-in)
import { Validation } from "@pkg/validation";
Validation.addList(wb, "Sheet1", "A:A", ["x", "y", "z"]);

// @pkg/streaming  (existing subpath, kept)
import { Stream } from "@pkg/stream";
for await (const row of Stream.readRows(path)) { ... }
```

`Workbook`, `Sheet`, `Cell` are namespace objects of free functions. They
operate on opaque `Workbook` / `Sheet` / `Cell` handle types. The handle types
are plain records internally — mutation is allowed for performance, but the
public contract is "pass the handle to a function".

### 4.3 `chart` — extract to subpath ❌ → ✅

Currently 30,278 LoC. Exposed only via `excel` barrel and `excel/worksheet`
method calls.

Target:

- `@pkg/chart` subpath.
- `Chart` namespace with free functions: `Chart.add`, `Chart.remove`,
  `Chart.renderSvg`, `Chart.renderPng`, `Chart.applyPreset`, etc.
- Core `excel` module only declares `type ChartHandle = { /* opaque */ }` and
  routes through a host-registry when chart operations are initiated from
  Workbook write path.
- Chart ex, chart presets, chart renderers, chart sidecar all internal to the
  subpath.

Expected saving: consumers not importing `@pkg/chart` save 1.2 MB.

### 4.4 `word` — partial refactor ⚠

Tree-shake ratio 28.4%. Two God classes: `DocxDocument` and `DocumentBuilder`.
Similar issue to Workbook but smaller scale (20k LoC vs 109k).

Target: same pattern as excel. `Document`, `Paragraph`, `Run`, `Table`
namespaces. `IO.toBuffer`, `IO.fromBuffer`. Subpath for Word-specific advanced
features (styles, sections, comments).

### 4.5 `pdf` — internal class, public API should be functions ⚠

15 classes exported. `PdfDocument`, `PdfTokenizer`, `CMap`,
`PdfContentStream`, `PdfDict`, `PdfWriter` are internal state machines. Fine
to keep as classes internally, but public API should be functions:

```ts
// instead of
const doc = new PdfDocument();
doc.addPage(...);
const buf = doc.toBuffer();

// prefer
const doc = Pdf.create();
Pdf.addPage(doc, ...);
const buf = await Pdf.toBuffer(doc);
```

Existing `excelToPdf()` free function is a good example and should become
the canonical shape for the whole PDF API.

### 4.6 `archive` — stream compat classes are OK ✅

41 exported classes. Almost all are:

- Node/Web stream compatibility shims (`PullStream`, `ZipReader`, `ZipParser`,
  `Parse`).
- Error classes.
- Reader adapters (`RandomAccessReader`, `HttpRangeReader`, `BufferReader`).

These are legitimate class uses — state machines and stream subclasses that
Node/Web stream contracts require. The public API for archive operations is
already free functions: `extractAll`, `zip`, `unzip`, `tar`, `untar`. No
change needed.

### 4.7 `csv` — mostly right ✅

Tree-shake ratio 15.2%. Public API is free functions (`parseCsv`,
`formatCsv`, `parseCsvRows`). The few exported classes are stream
subclasses and errors. No change needed.

### 4.8 `stream` — mostly right ✅

Native Node/Web stream types and helpers. Tree-shake ratio 13%. No change
needed.

### 4.9 `markdown` — small and clean ✅

1695 LoC, 2 Error classes, everything else free functions. No change needed.

### 4.10 `xml` — mostly right ⚠

5063 LoC, 3 Error classes, 12 free functions. Tree-shake ratio 37.4%,
slightly worse than expected. Audit the public barrel; likely one or two
internal symbols are reachable from multiple exports and could be split.

### 4.11 `utils` — fine ✅

Free functions throughout.

## 5. Implementation roadmap

### Phase 0: preparation (1 week)

- Pin the public API shape with a written contract (this document).
- Add bundle-size regression tests: golden numbers per subpath, fail CI if
  they regress.
- Add `--metafile` analysis as a CI artifact.
- Document the `host-registry` pattern as the canonical decoupling mechanism.

### Phase 1: extract chart to subpath (breaking change, v10.0)

Dependencies: none.
Impact: consumers that use charts must add `installChartSupport()` or import
from `@pkg/chart`.

Steps:

1. New file `src/modules/excel/chart-host-registry.ts` — types + slot.
2. New file `src/modules/excel/chart/install.ts` — populates slot from
   concrete implementations.
3. Rewrite `worksheet.ts`, `workbook.browser.ts`, `chartsheet.ts` to depend
   only on `chart-host-registry` and call through `getChartSupport()`.
4. Remove `@excel/chart` re-exports from `src/index.ts`. Move them to
   `src/chart.ts` (the `./chart` subpath entry).
5. Add `./chart` to `package.json` exports map.
6. Update `src/test/setup-formula.ts` to also install chart support, or
   split into `setup.ts` and let individual chart tests install on demand.
7. Update MIGRATION.md with a v9 → v10 section.

Expected result: consumers not importing `@pkg/chart` save ~1.2 MB.

### Phase 2: Worksheet / Workbook to namespace objects (breaking change, v11.0)

Dependencies: Phase 1 complete.
Impact: method calls become function calls (`wb.addSheet()` →
`Workbook.addSheet(wb)`).

Steps:

1. Introduce `Workbook`, `Sheet`, `Cell`, `Row`, `Column`, `Range`, `IO` as
   exported namespace objects. Retain the class forms as deprecated re-exports
   during v11 so codemod can run incrementally.
2. Provide a codemod (`jscodeshift` or `ts-morph`) that rewrites:
   - `new Workbook()` → `Workbook.create()`
   - `wb.addWorksheet(name)` → `Workbook.addSheet(wb, name)`
   - `ws.getCell(addr).value = v` → `Cell.set(ws, addr, v)`
   - …
3. Keep internal implementation as classes if perf benchmarks show them
   faster. Public surface is free functions; private storage is flexible.
4. Delete deprecated class exports in v12.

Expected result: tree-shake ratio drops from 40% to <10% for consumers who
use only core operations. Bundle from 2.1 MB to <500 KB.

### Phase 3: extract remaining subpaths (v12.0)

- `@pkg/pivot` — PivotTable, CacheDefinition, cacheRecords
- `@pkg/validation` — DataValidations, ignored errors
- `@pkg/forms` — FormCheckbox
- `@pkg/threads` — Threaded comments
- `@pkg/slicer` — Slicers and timelines

Each extraction follows the same pattern: host-registry + subpath + free
functions.

### Phase 4: word module parity (v12.1)

Apply the same treatment to the word module. Same pattern.

### Phase 5: eliminate barrel re-exports from root (v13.0)

Main entry `@pkg` exports only core types and the IO namespace. Everything
else is a subpath. `src/index.ts` becomes a ~50-line file.

## 6. Design patterns used

### 6.1 Opaque handles

```ts
// core
declare const WorkbookBrand: unique symbol;
export type Workbook = { readonly [WorkbookBrand]: true };

// factory returns the branded type but internally is a record
export const Workbook = {
  create(): Workbook {
    return _internalCreateWorkbook() as unknown as Workbook;
  }
};
```

Consumers cannot construct `Workbook` directly, but `Workbook.create()`
returns one. Internally the object is mutable for performance.

### 6.2 Namespace objects of free functions

```ts
export const Sheet = {
  set(sheet: Sheet, addr: string, value: CellValue): void { ... },
  get(sheet: Sheet, addr: string): Cell | undefined { ... },
  rangeSet(sheet: Sheet, range: string, grid: CellValue[][]): void { ... }
};
```

Benefits:

- Tree-shakes (bundler can prove `Sheet.rangeSet` unused if consumer only
  calls `Sheet.set`).
- IDE autocomplete on `Sheet.` discovers methods.
- Mock/stub in tests by replacing `Sheet.set` on the namespace.

### 6.3 Host-registry for capability decoupling

Core declares an interface slot, defaulting empty:

```ts
// @pkg/chart-host-registry (tiny)
export interface ChartSupport {
  add: (wb: Workbook, sheet: string, opts: unknown) => number;
  renderSvg: (wb: Workbook, id: number) => string;
  // ...
}
let slot: ChartSupport | null = null;
export function registerChartSupport(impl: ChartSupport | null): void {
  slot = impl;
}
export function getChartSupport(): ChartSupport {
  if (!slot) throw new Error("Chart support not installed");
  return slot;
}
```

Optional module provides `install()`:

```ts
// @pkg/chart
import { registerChartSupport } from "@pkg/chart-host-registry";
import { addChart, renderSvg /* ... */ } from "./internal";
export function installChartSupport(): void {
  registerChartSupport({ add: addChart, renderSvg /* ... */ });
}
export const Chart = {
  /* public surface */
};
```

Consumer that uses charts:

```ts
import { installChartSupport, Chart } from "@pkg/chart";
installChartSupport();
// now Chart functions work
Chart.add(wb, "Sheet1", ...);
```

Consumer that does not use charts: no import from `@pkg/chart`. Bundler cannot
reach any chart runtime from the consumer's entry. 1.2 MB saved.

### 6.4 Web Platform primitives

```ts
export const IO = {
  toBuffer(wb: Workbook): Promise<Uint8Array> { ... },
  fromBuffer(buf: Uint8Array): Promise<Workbook> { ... },
  toStream(wb: Workbook): ReadableStream<Uint8Array> { ... },
  fromStream(s: ReadableStream<Uint8Array>): Promise<Workbook> { ... }
};

export const Stream = {
  readRows(src: ReadableStream<Uint8Array> | string): AsyncIterable<Row> { ... },
  writer(): WritableStream<Row> { ... }
};
```

No Node-specific types in the signature. Node users pass `fs.createReadStream(...)`
through a `Readable.toWeb()` bridge if needed.

## 7. Risks and mitigations

1. **User codebase impact.** Switching from methods to free functions is
   mechanical but large. Mitigation: ship a codemod with v11. Support both
   forms during v11.x.

2. **Loss of IDE "dot-discovery".** `wb.` today shows every method. With
   namespace objects, user types `Workbook.` — different trigger. Mitigation:
   namespace naming is predictable (`Workbook`, `Sheet`, `Cell`, `IO`, `Chart`,
   `Pivot`, …) and matches well-known patterns (Effect-TS, Array, Object,
   JSON). Documentation emphasizes these namespaces.

3. **Performance regression.** Calling `Sheet.set(s, addr, v)` vs
   `s.set(addr, v)` adds a property access on the namespace object. Mitigation:
   V8 caches this; no measurable difference in benchmarks for this kind of
   hot path. Verify with the existing `benchmark/` suite before each phase.

4. **npm dual-package hazard.** CJS consumers loading both the core and a
   subpath may get duplicate module state. Mitigation: the registry slots are
   the only mutable module-level state; they must be in a single `host-registry`
   module that both core and the optional module import from.

5. **Subpath resolution in older bundlers.** Webpack 4, parcel 1, older
   tsc resolutions do not fully support `exports` field. Mitigation: document
   minimum tooling versions (Node 18+, webpack 5+, Vite 2+, Rolldown, esbuild,
   TypeScript 4.7+).

6. **Breaking change fatigue.** 3 major versions over 6-12 months. Mitigation:
   each major has a clear win and a codemod. Changelog is the contract.

## 8. Success metrics

| Metric                                     | Current | After Phase 2 | After Phase 5 |
| ------------------------------------------ | ------: | ------------: | ------------: |
| `@pkg` main entry bundle (simple consumer) | 2146 KB |        300 KB |        150 KB |
| `@pkg/chart` only when needed              |     n/a |       1200 KB |       1200 KB |
| Tree-shake ratio (excel path)              |   40.6% |          <10% |           <8% |
| `pnpm test` wall time (isolate:true)       |    148s |          <30s |          <20s |
| `pnpm run generate:csv-worker`             |    slow |          fast |          fast |

## 9. Open questions

1. Do we support a compatibility layer (deprecated class methods that delegate
   to namespace functions) during v11, or hard-break at v11.0? Codemod quality
   determines this.

2. How do we expose chart types without making consumers install the chart
   subpath? Current plan: `type ChartHandle` and friends live in the core
   module as pure types; the runtime implementation lives in `@pkg/chart`.
   TypeScript-only references do not cost bundle bytes.

3. Do we keep `Workbook` as both a namespace object and a type? Yes, following
   the `Array`, `Object`, `Date` pattern from the standard library. The type
   is a branded record; the value is a namespace of free functions.

4. Should internal classes be renamed with an `_` prefix to discourage direct
   use? Or moved behind a `/internal` export pattern? Decision: use `internal`
   subpath not exported via `package.json` but importable from inside the
   monorepo.

---

**Summary**: `formula` is already at the target. `excel` needs the most work
because its Workbook/Worksheet god classes force tree-shaking to fail. The
refactor is staged across three major versions with codemods to keep the
migration tractable. Expected end state: main bundle reduced from 2.1 MB to
~150 KB, with optional features on independent subpaths.
