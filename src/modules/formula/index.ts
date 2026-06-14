/**
 * Public entry for the excelts formula engine.
 *
 * **Functional, zero-side-effect** is the only evaluation style:
 * ```ts
 * import { calculateFormulas } from "@cj-tech-master/excelts/formula";
 * calculateFormulas(workbook);
 * ```
 * Perfect tree-shaking: unused exports are dropped, no module
 * initialisation runs. Bundles that import only `tokenize` or `parse`
 * never pull the evaluator or function registry in. There is no
 * `Workbook.calculateFormulas()` method and no host-registry — the engine
 * is reached solely through this free function, so consumers who never
 * import it never pay for it.
 *
 * For PDF export recalculation, pass `calculateFormulas` to
 * `excelToPdf(wb, { recalculate: calculateFormulas })`.
 *
 * Separately, `installFormulaEngine()` installs the default **syntax
 * probe** (tokenizer+parser) used during XLSX load to classify
 * defined-name text as formula vs. opaque. It lives in a separate module
 * (`./install.ts`) so its parser import doesn't leak into functional-only
 * consumers. See `scripts/treeshake-verify.ts` for the bundler contracts.
 */

// Functional API — the sole formula-evaluation entry point. Useful for
// server-side recalculation of cached XLSX files loaded via the excel module.
export { calculateFormulas } from "./integration/calculate-formulas";

// Low-level syntax surface — for tooling, static analysers and callers
// that want to pre-validate formulas without evaluating them.
export { tokenize } from "./syntax/tokenizer";
export { parse } from "./syntax/parser";

// Syntax-probe installer — isolated module so its parser imports don't
// leak into functional-only consumers. `createFormulaSyntaxProbe` is
// exported for callers that want a standalone probe (e.g. per-Workbook
// injection via `new Workbook({ formulaSyntaxProbe })`) without touching
// process-global state. `uninstallFormulaEngine` is exported for symmetry
// and for test suites that exercise the cold-start classification path.
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
