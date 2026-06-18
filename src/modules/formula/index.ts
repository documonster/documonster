/**
 * Public entry for the excelts formula engine.
 *
 * **Functional, zero-side-effect** is the only evaluation style:
 * ```ts
 * import { Formula } from "@cj-tech-master/excelts/formula";
 * Formula.calculate(workbook);
 * ```
 * Perfect tree-shaking: unused exports are dropped, no module initialisation
 * runs, and there is **no install / registration step** — every export is used
 * directly. Bundles that import only `Formula.tokenize` / `Formula.parse` never
 * pull the evaluator or function registry in; the evaluator is reached solely
 * through `Formula.calculate`, so consumers who never call it never pay for it.
 *
 * For PDF export recalculation, pass `Formula.calculate` to
 * `excelToPdf(wb, { recalculate: Formula.calculate })`.
 */

// Public value API — the `Formula` domain namespace. Tree-shaken per-member
// on rolldown / rspack; a consumer that references only `Formula.tokenize`
// never pulls in the evaluator.
export * as Formula from "@formula/surface/formula";

// Errors — extend BaseError, consistent with every other module's errors.ts.
export { FormulaError, FormulaParseError } from "@formula/errors";

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
} from "@formula/materialize/types";
