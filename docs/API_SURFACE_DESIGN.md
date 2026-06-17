# API Surface Design — Subpath Namespace Public Entry

Status: design locked, implementation in progress.
Owner: core team.
Last updated: 2026-06-17.
Supersedes the export-surface portion of `FUTURE_ARCHITECTURE.md`.

---

## 0. TL;DR

> The public API is **domain namespaces, accessed only through module
> subpaths** — there is **no root aggregate entry**. The `documonster` root
> is removed from `exports`; consumers import from `documonster/excel`,
> `documonster/word`, etc. Namespace **depth follows module complexity**:
> large multi-domain modules expose several namespaces (`Workbook`, `Cell`,
> `Chart` from `documonster/excel`); single-purpose modules expose **one
> flat namespace** (`Csv.parse`, `Csv.format` from `documonster/csv`).
> **Maximum depth is two** (`Workbook.create()` or `Csv.parse()`) — never
> three. Low-level handle functions (`cellSetValue`, `addChart`) are
> **internal** — used by tests and by the namespace facades, never in
> `package.json` `exports`. Each module ships its **own IIFE**
> (`Documonster.Excel`, `Documonster.Csv`, …). Everything is wired by static
> imports — no `register`/`install` for plumbing — and tree-shaken by the
> bundler, **verified by bundle-size tests**.

Satisfies every standing requirement:

1. Unified public surface (namespace everywhere), max two levels deep.
2. No three-level `Excel.Workbook.create()` ergonomics tax; no fake
   "everything" root entry.
3. Low-level functions still usable internally (tests).
4. No runtime registry for plumbing; legitimate extension points kept.
5. Tree-shakeable, and **measured**, not asserted.

### Why no root aggregate entry

A root `import { Excel } from "documonster"` forces `Excel.Workbook.create()`
— **three levels** (package → `Excel` → `Workbook` → method), strictly worse
than `Workbook.create()`. And a "whole-family" root is a **phantom need**:
nobody uses Excel + Word + Pdf + Formula in one file; an excel consumer
imports from `documonster/excel`. So the root entry earns its keep nowhere —
it is removed, not aggregated.

---

## 1. Why this document exists

After the class → free-function refactor the public surface drifted into
three inconsistent styles across the module entries. Measured state:

| Entry (`exports` subpath) | Entry file             | Style                         |
| ------------------------- | ---------------------- | ----------------------------- |
| `documonster` (`.`)       | `src/index.ts`         | **flat only** (9× `export *`) |
| `documonster/excel`       | `excel/index.ts`       | **namespace only** (12×)      |
| `documonster/word`        | `word/index.base.ts`   | **mixed** (1 ns + 50 flat)    |
| `documonster/formula`     | `formula/index.ts`     | flat only                     |
| `documonster/pdf`         | `pdf/index.ts`         | flat only                     |
| `documonster/csv`         | `csv/index.ts`         | flat only                     |
| `documonster/markdown`    | `markdown/index.ts`    | flat only                     |
| `documonster/xml`         | `xml/index.ts`         | flat only                     |
| `documonster/archive`     | `archive/index.ts`     | flat only                     |
| `documonster/chart`       | `excel/chart/index.ts` | flat (functional engine)      |

The original convention — **"public = namespace, internal/tests = flat"** —
was never enforced: the `documonster` main entry `export *`s hundreds of
low-level functions, and it is in practice an "excel + a few oddments" dump
(measured: it imports overwhelmingly `@excel/*`, not a true 9-module
aggregate).

---

## 2. Three orthogonal dimensions (the source of the confusion)

"namespace vs flat" conflates three independent things. Separating them is
the whole design:

| Dimension           | Question                                         | Decision                                            |
| ------------------- | ------------------------------------------------ | --------------------------------------------------- |
| **A. API layering** | high-level (sheet+address) vs low-level (handle) | Keep both. They are not redundant.                  |
| **B. Export shape** | `export * as Ns` vs flat named exports           | Public = namespace. Internal = flat (not exported). |
| **C. Visibility**   | public API vs internal                           | Public = fixed namespaces per module.               |

### A is real, not cosmetic

The namespaces are a **re-designed high-level layer**, not renamed flat
functions:

```ts
cellSetValue(cell, 42); // low-level (handle-centric) — internal
Excel.Cell.setValue(ws, "A1", 42); // high-level (sheet-centric) — public
```

`surface/cell.ts` defines 21 of its own `(ws, addr, …)` functions — it is not
a forwarding shim. Hiding the flat layer removes nothing the namespace lacks.

---

## 3. Target architecture — subpath namespaces, depth by complexity

```
ROOT  documonster (".")  — REMOVED from package.json "exports"
─────────────────────────────────────────────────────────────────
  (no root aggregate; see §0 "Why no root aggregate entry")

SUBPATHS  documonster/<module>  — the only public entries
─────────────────────────────────────────────────────────────────
  Large, multi-domain modules → several namespaces (one level of ns):
    documonster/excel    →  Workbook Worksheet Cell Row Column Range
                            Chart Table Pivot Sparkline Form Image      (already ✅)
    documonster/word     →  Document Query Convert Font Layout Docx …   (to build)
    documonster/pdf      →  Pdf  (+ sub-namespaces if it has distinct domains)
    documonster/archive  →  Zip Tar …
    documonster/stream   →  Stream …

  Single-purpose modules → ONE flat namespace (one level, no forced nesting):
    documonster/csv      →  Csv     ( Csv.parse, Csv.format, Csv.detectDelimiter )
    documonster/formula  →  Formula ( Formula.tokenize, Formula.parse, Formula.calculate )
    documonster/markdown →  Markdown
    documonster/xml      →  Xml

  documonster/chart      →  low-level functional render engine (stays — own product)

  →  import { Workbook } from "documonster/excel";  Workbook.create()    // 2 levels max
  →  import { Csv }      from "documonster/csv";     Csv.parse(text)      // 2 levels max

INTERNAL  (not in "exports"; reachable only via package-internal alias)
─────────────────────────────────────────────────────────────────
  @excel/cell      cellSetValue, cellGetValue, …    (tests + Cell.* facade)
  @excel/worksheet addChart, getCharts, addTable …  (tests + namespace facades)
  @excel/row …     rowSetValues, …
```

### Depth rule: every module gets a namespace; sub-namespaces by complexity

**Every public _domain_ module is exposed through a namespace — no bare flat
exports.** This is a deliberate consistency decision: a uniform `Ns.method`
mental model across modules, and clean IIFE globals (`Documonster.Csv.parse`).

**Exception — infrastructure/primitive modules stay flat.** `stream` (Node
stream primitives: `Transform`, `PassThrough`, `Readable`, `pipeline`, … — a
`node:stream` polyfill) and `archive` (`crc32`, `Zip`, `Tar`, `ZipDeflate`
primitives) are NOT domain APIs. They are foundational primitives imported by
20+ internal cross-module files as the equivalent of `node:stream` /
`node:zlib`. Forcing `Stream.Transform` / `Archive.crc32` would be high-churn
and semantically wrong (a primitive class is not a domain operation). These
modules keep flat exports.

Domain (namespaced): excel, word, csv, markdown, xml, pdf, formula, chart.
Infrastructure (flat): stream, archive.

Within a domain module, depth follows complexity (never exceed two levels):

- **Multi-domain module** (Excel: cells, charts, pivots, sparklines are
  genuinely different concerns) → several namespaces, each one level:
  `Workbook.create()`, `Cell.setValue()`, `Chart.add()`.
- **Single-purpose module** (csv = parse/format/stream facets of one job) →
  **one flat namespace**: `Csv.parse()`, `Csv.format()`,
  `Csv.detectDelimiter()`. It still gets a namespace (`Csv`), just no nested
  sub-namespaces. Forcing `Csv.Parse.sync()` (three effective levels) is
  rejected — the flat `Csv.*` is the right grain.

Heuristic: every module → exactly one namespace minimum. Introduce a
sub-namespace inside it only when two groups have genuinely distinct
domains/handles. Otherwise the module namespace stays flat.

### Orphan / cross-cutting / hard-to-classify functions

There is **no bare flat export** in the public surface — that is exactly what
this design eliminates. Every public function lands on _some_ namespace:

1. **Top-level whole-module operations** (operate on the entire document /
   workbook, belong to no sub-domain) → hang directly on the module's root
   namespace. e.g. `mergeDocuments`/`splitDocument`/`readDocx`/`packageDocx`
   → `Word.merge()` / `Word.split()` / `Word.read()` / `Word.package()`. The
   module namespace is both a container for sub-namespaces **and** a home for
   these top-level verbs.
2. **A related cluster that fits no existing sub-namespace** → open a **new**
   sub-namespace (threshold: ≥2 related functions with a clear domain name).
   e.g. `parseFlatOpc`/`toFlatOpc`/`isFlatOpc` → `Word.FlatOpc.*`;
   `embedFont`/`subsetFont` → `Word.Font.*`.
3. **A lone orphan utility** (single function, no neighbours) → also hangs on
   the module root namespace (`Word.foo`). Do not open a namespace for one
   function.

Rule: bare un-namespaced exports = **zero tolerance**. Worst case a function
lives on the module root namespace; it never escapes naked.

### Bridges (cross-module conversion)

Bridges convert A → B. Two sub-cases, classified by **what they are
centred on**, not where the file lives:

1. **True cross-module converters** (produce a different document type) →
   namespace of the **target/producing** module, named as a `from<Source>`
   verb. The bridge lives in the upper layer that already depends on the
   source (e.g. `pdf` depends on `excel`):
   - `excelToPdf` → `Pdf.fromExcel(wb, opts)`
   - `docxToPdf` → `Pdf.fromDocx(doc, opts)`
   - `excelToDocx` → `Word.fromExcel(wb, opts)`
   - `extractTablesToExcel` → `Word.tablesToExcel(...)` (word holds the bridge)
2. **Workbook IO into/out of a text format** (first arg is the workbook;
   the workbook _is_ the subject) → `Workbook` namespace:
   - `readCsv(wb,…)` / `writeCsv(wb,…)` → `Workbook.fromCsv` / `Workbook.toCsv`
   - `readMarkdown(wb,…)` / `writeMarkdown(wb,…)` → `Workbook.fromMarkdown` / `Workbook.toMarkdown`

So no bridge stays flat either: converters → target namespace, workbook IO →
`Workbook`.

### Why no root entry (recap)

`Excel.Workbook.create()` is three levels and strictly worse than
`Workbook.create()`. A whole-family root import is a phantom need. The root
`"."` is dropped from `exports`; subpaths are the public contract.

### Why low-level functions go internal (not deleted)

- They are the implementation substrate of the namespaces (`Cell.setValue`
  calls `cellSetValue`).
- Tests legitimately need handle-level access. Tests import the concrete
  module (`@excel/cell`) directly — aliases reach any package-internal file and
  are **not** constrained by `package.json` `exports`.

---

## 4. `documonster/chart` stays a functional engine (on purpose)

`documonster/chart` is **not** the excel high-level surface. It is the
standalone render engine (`buildChartModel`, `renderChartSvg`, `parseChartEx`)
for consumers who want charts **without** the excel object model. Its flat
exports are its product, not leakage.

| Layer                                           | Action             |
| ----------------------------------------------- | ------------------ |
| excel handle-level (`cellSetValue`, `addChart`) | → internal         |
| chart functional engine (`buildChartModel`)     | → **stays public** |

---

## 5. IIFE / browser globals — one IIFE per module

Current state (wrong): a single `ExcelTS` global built from `index.browser`
that drags in **all** modules (excel+word+pdf+formula+…), no per-module
opt-in.

Target: **one IIFE bundle per module**, each exposing that module's
namespace, under a shared `Documonster` global namespace:

| IIFE artifact                 | Global                        |
| ----------------------------- | ----------------------------- |
| `documonster.excel.iife.js`   | `Documonster.Excel.Workbook…` |
| `documonster.word.iife.js`    | `Documonster.Word.Document…`  |
| `documonster.formula.iife.js` | `Documonster.Formula…`        |
| `documonster.pdf.iife.js`     | `Documonster.Pdf…`            |
| `documonster.csv.iife.js`     | `Documonster.Csv…`            |
| … (one per public module)     | `Documonster.<Module>`        |

There is **no** whole-family IIFE (mirrors "no root aggregate entry" — a
single global carrying every module is the same phantom need, and it is what
the current single `ExcelTS` IIFE wrongly does today).

CDN users load only the script(s) they need. Global shape == ESM shape.
This also removes the class of browser-smoke breakage caused by an
undefined global API shape.

---

## 6. No plumbing registry; keep real extension points

Distinguish two kinds of "register":

| Kind                                                                                                          | Verdict                                                                                               |
| ------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| **Plumbing install** (`chart-host-registry`, `installFormulaEngine`) — manual wiring to break an import cycle | **Remove.** Replace with static imports.                                                              |
| **User extension** (`registerFunction(desc)` — register a custom formula function)                            | **Keep.** It is a legitimate API; expose as `Formula.registerFunction` (or a per-evaluator registry). |

- `chart`: registry **removed**, fully static, 16 260 tests green (done).
- `formula`: drop `installFormulaEngine` / `setDefaultSyntaxProbe` global slot
  (the probe is built from pure `tokenize`+`parse`; resolve statically or
  per-`Workbook`). **Preserve `registerFunction`** as the user extension point.

Never delete a useful extension registry just because the word "register"
appears.

---

## 7. Tree-shaking — must be measured

`export * as Ns` + `Ns.foo()` is member-level tree-shakeable under
**rolldown, rspack, webpack** and recent **esbuild**, given `sideEffects:false`
(set) and side-effect-free surface files (true). The real lever is the
**dependency graph** (why the chart registry was removed), not the export
shape.

There are currently **zero** bundle-size tests (measured). Add a
`bundle-size` fixture suite that asserts what each entry pulls in:

| Scenario                                                   | Must NOT include                        |
| ---------------------------------------------------------- | --------------------------------------- |
| `import { Cell } from "documonster/excel"; Cell.setValue`  | chart renderer, pdf, formula evaluator  |
| `import { Chart } from "documonster/excel"; Chart.add`     | SVG/PNG renderer (unless `toSVG/toPNG`) |
| `import { Formula } from "documonster/formula"; .tokenize` | the 433-function evaluator              |
| `import { Csv } from "documonster/csv"`                    | excel object model                      |
| `import { Excel } from "documonster"; Excel.Cell`          | word, pdf, formula                      |

Implementation: bundle each scenario with rolldown `--metafile` (or rspack),
assert byte budgets + absence of forbidden modules. Run in CI.

**esbuild caveat**: older esbuild shakes namespace members poorly. README must
state production bundling uses rolldown/rspack/webpack.

---

## 8. Execution plan (phased)

| #   | Phase                                                                                                                                            | Size               |
| --- | ------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------ |
| 1   | Remove root `"."` from `package.json` `exports`; delete `src/index.ts`/`index.browser.ts` (or keep only as internal dev aggregate, not exported) | small              |
| 2   | **word**: refactor to excel-isomorphic surface namespaces (Document/Query/Convert/Font/Layout/Docx)                                              | large (design)     |
| 3   | wrap formula/pdf/csv/markdown/xml/archive each in a namespace (flat unless genuinely multi-domain)                                               | medium             |
| 4   | **formula**: drop `installFormulaEngine` plumbing; keep `registerFunction` extension                                                             | small              |
| 5   | low-level handle fns → internal (out of `exports`); migrate 48 tests + 91 examples                                                               | large (mechanical) |
| 6   | `rolldown.config` → per-module IIFE (`Documonster.<Module>`); sync `package.json exports`                                                        | medium             |
| 7   | bundle-size / tree-shaking measurement test suite                                                                                                | medium             |
| 8   | full verify: `pnpm check` + node suite + browser suite green                                                                                     | —                  |

Phase 2 (word grouping) is the only genuine **design** work; the rest is
structural/mechanical.

---

## 9. Risks & trade-offs (honest)

1. **Breaking change.** Every external `import { … } from "documonster"` and
   the flat subpath imports break. Accepted (explicit decision).
2. **One extra `.` vs plain functions.** `Workbook.create()` /
   `Csv.parse()` add one namespace hop over a bare `createWorkbook()`. Worth
   it for discoverability, zero name collisions, and a crisp contract. Depth
   is capped at two — no three-level `Excel.Workbook.create()` (root removed).
3. **Namespace-only public API is a minority style** (date-fns/lodash/zod lean
   flat). For a **large multi-module** library the benefit — discoverability,
   no collisions, crisp contract, IDE completion — outweighs the cost.
4. **word grouping is a judgement call** — designed, not auto-generated.

---

## 10. Decisions locked

- Root `documonster` entry: **removed** from `exports`. No aggregate, no
  whole-family bundle (phantom need; three-level `Excel.Workbook.create()`).
- Public entries are **module subpaths only** (`documonster/excel`, …).
- Namespace **depth follows complexity**, max two levels: multi-domain
  modules expose several namespaces (`Workbook`, `Cell`, `Chart`);
  single-purpose modules expose one flat namespace (`Csv.parse`).
- Low-level handle functions: **internal** (tests + facades), not in `exports`.
- `documonster/chart` functional engine: **stays public** (separate product).
- IIFE: **one per module**, namespace globals under `Documonster`; no
  whole-family IIFE.
- Registry: remove plumbing installs; **keep `registerFunction`** extension.
- Tree-shaking: enforced by a **bundle-size test suite** (currently none).
