/**
 * Formula Calculation Engine
 *
 * Provides `calculateFormulas()` to recalculate all formula cells in a workbook.
 * This is the sole public entry point for formula calculation.
 *
 * ## Architecture
 *
 * 1. **Snapshot** — immutable capture of all workbook state
 * 2. **Normalize** — uniform FormulaInstance objects
 * 3. **Parse** — tokenize → AST
 * 4. **Compile** — bind AST → BoundExpr (resolve names, structured refs, sheets)
 * 5. **Dependency Analysis** — topological sort
 * 6. **Evaluate** — execute BoundExpr with RuntimeValue system
 * 7. **Materialize** — build declarative WritebackPlan
 * 8. **Apply** — write plan to live workbook
 */

import type { WorkbookLike } from "../materialize/types";
import { calculateFormulasImpl } from "./calculate-formulas-impl";

// Re-export shared types for external consumers
export type { DefinedNamesLike, WorkbookLike } from "../materialize/types";

/**
 * Recalculate all formula cells in a workbook.
 *
 * Evaluates every formula cell using the built-in calculation engine
 * and updates each cell's `result` value. Formulas are evaluated lazily
 * with recursive dependency resolution, memoization, and circular
 * reference detection.
 *
 * All evaluation state is scoped to this invocation — concurrent calls
 * for different workbooks are safe.
 *
 * **Supported formula features:**
 * - Cell references: `A1`, `$B$2`, `Sheet1!A1`, `'Sheet Name'!A1:B10`
 * - Operators: `+ - * / ^`, `& (concat)`, `= <> < > <= >=`, `%`
 * - 120+ built-in functions (SUM, IF, VLOOKUP, SUMIF, FILTER, etc.)
 * - Shared formulas, array constants, nested expressions
 * - Dynamic array spill: FILTER, SORT, UNIQUE, SORTBY results are
 *   written to adjacent cells. #SPILL! error if target cells are occupied.
 * - CSE array formulas: `{=formula}` with a ref range distribute results
 *   across the designated range.
 * - Array arithmetic broadcasting: `{1,2,3} + {4;5;6}` produces a 3x3 matrix.
 * - Implicit intersection: range references in scalar context pick the
 *   value from the formula cell's row or column.
 *
 * **Unsupported formula behavior:**
 * - If a formula uses a function the engine does not implement, the engine
 *   returns `#NAME?`. However, if the cell already has a cached result
 *   (e.g., pre-computed by Excel when the XLSX was saved), that cached
 *   result is **preserved** — the engine will not overwrite usable data.
 * - If no cached result exists, the cell's result becomes `#NAME?`.
 *
 * **Volatile functions:**
 * - `RAND`, `RANDBETWEEN`, `NOW`, `TODAY` are re-evaluated on every call.
 *   This is intentional — these functions are expected to produce fresh values.
 *
 * **Side effects:**
 * - This function **mutates** the workbook by updating formula cells' `result`
 *   property in-place. For dynamic array formulas, adjacent cells are also
 *   written with spill results. If you need the original cached results
 *   preserved, clone the workbook before calling this function.
 *
 * @param workbook - The workbook whose formulas should be recalculated
 */
export function calculateFormulas(workbook: WorkbookLike): void {
  calculateFormulasImpl(workbook);
}
