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

// Public value API — the `Formula` domain namespace. Tree-shaken per-member
// on rolldown / rspack; a consumer that references only `Formula.tokenize`
// never pulls in the evaluator.
export * as Formula from "./surface/formula";

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
