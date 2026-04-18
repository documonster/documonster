/**
 * Public entry for the excelts formula engine.
 *
 * Two complementary usage styles are supported:
 *
 * 1. **Functional, zero-side-effect**:
 *    ```ts
 *    import { calculateFormulas } from "@cj-tech-master/excelts/formula";
 *    calculateFormulas(workbook);
 *    ```
 *    Perfect tree-shaking: unused exports are dropped, no module
 *    initialisation runs. Bundles that import only `tokenize` or
 *    `parse` never pull the evaluator or function registry in.
 *
 * 2. **Engine installation**, to enable `Workbook.calculateFormulas()`
 *    and automatic recalculation in `excelToPdf()`:
 *    ```ts
 *    import { installFormulaEngine } from "@cj-tech-master/excelts/formula";
 *    installFormulaEngine();                // once, at startup
 *    workbook.calculateFormulas();          // now works
 *    ```
 *
 * The engine is **never installed implicitly** — consumers pay for what
 * they ask for. This keeps the root `package.json` `sideEffects: false`
 * contract intact, so bundles that only use the functional API include
 * exactly the code paths reachable from the exports they reference.
 *
 * Note: `installFormulaEngine` lives in a separate module
 * (`./install.ts`) so that its host-registry import and the
 * evaluator pipeline it carries don't get pulled in by consumers who
 * only need the tokenizer, parser, or the functional `calculateFormulas`
 * API. See `scripts/treeshake-verify.ts` for the bundler contracts.
 */

// Functional API — same capability exposed as a callable rather than
// going through `Workbook.calculateFormulas()`. Useful for server-side
// recalculation of cached XLSX files loaded via the excel module.
export { calculateFormulas } from "./integration/calculate-formulas";

// Low-level syntax surface — for tooling, static analysers and callers
// that want to pre-validate formulas without evaluating them.
export { tokenize } from "./syntax/tokenizer";
export { parse } from "./syntax/parser";

// Engine installer — isolated module so its imports don't leak into
// functional-only consumers. `createFormulaSyntaxProbe` is exported for
// callers that want a standalone probe (e.g. for per-Workbook injection
// via `new Workbook({ formulaSyntaxProbe })`) without touching
// process-global state. `uninstallFormulaEngine` is exported for
// symmetry and for test suites that exercise the cold-start
// classification path.
export { createFormulaSyntaxProbe, installFormulaEngine, uninstallFormulaEngine } from "./install";

// Re-export the probe type so consumers can type variables holding
// probes or constructing options objects that accept one.
export type { SyntaxProbe } from "./default-syntax-probe";

// Structural types callers may need to describe their host workbook.
export type {
  CellErrorValueLike,
  CellLike,
  DefinedNameEntry,
  DefinedNamesLike,
  DimensionsLike,
  FormulaResultLike,
  RowLike,
  SpillRegion,
  WorkbookLike,
  WorksheetLike
} from "./materialize/types";
