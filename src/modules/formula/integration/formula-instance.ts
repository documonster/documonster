/**
 * Formula Instance — Unified model for all formula cell variants.
 *
 * The normalizer converts the heterogeneous formula representations found
 * in the snapshot (normal, shared-master, shared-slave, CSE, dynamic-array)
 * into a uniform `FormulaInstance` structure. After normalization, the
 * compile and runtime layers do not need to know how a formula originated.
 *
 * ## Invariants
 *
 * - Every `FormulaInstance` has a non-empty `sourceText` (the formula to parse).
 * - Shared-slave formulas have already been translated (offset applied) by
 *   the snapshot builder, so `sourceText` is the final formula text.
 * - `kind` is preserved so the materialize layer can determine the correct
 *   writeback strategy (scalar vs. CSE vs. spill).
 */

import type {
  CellSnapshot,
  FormulaCellKind,
  WorkbookSnapshot,
  WorksheetSnapshot
} from "./workbook-snapshot";

// ============================================================================
// FormulaInstance
// ============================================================================

/**
 * A normalized, engine-internal representation of a formula cell.
 * This is the unit of work that enters the compile → evaluate pipeline.
 */
export interface FormulaInstance {
  /** The worksheet name this formula belongs to. */
  readonly sheetName: string;
  /** The worksheet id (stable across renames). */
  readonly sheetId: number;
  /** 1-based row number. */
  readonly row: number;
  /** 1-based column number. */
  readonly col: number;
  /** The final formula text to parse and evaluate. */
  readonly sourceText: string;
  /**
   * The formula classification.
   * - `"normal"` — standard scalar formula
   * - `"shared-master"` — master of a shared formula group (treated as normal for eval)
   * - `"shared-slave"` — slave with translated formula (treated as normal for eval)
   * - `"cse"` — legacy CSE array formula; result distributed across `targetRef`
   * - `"dynamic-array"` — dynamic array formula; result spills to adjacent cells
   */
  readonly kind: FormulaCellKind;
  /**
   * For CSE formulas: the target range in "A1:B2" format.
   * For dynamic-array formulas: `undefined` (spill range is computed at runtime).
   * For all other kinds: `undefined`.
   */
  readonly targetRef?: string;
  /**
   * Whether this formula was explicitly flagged as a dynamic array formula
   * (from the XLSX `isDynamicArray` attribute).
   */
  readonly isDynamicArray: boolean;
}

// ============================================================================
// Normalizer
// ============================================================================

/**
 * Collect all formula instances from a workbook snapshot.
 *
 * Iterates every cell in every worksheet, selects cells with formulas,
 * and normalizes them into `FormulaInstance` objects. The result is a
 * flat list ordered by worksheet order then cell position.
 *
 * This replaces the old `collectFormulaCells()` function but operates
 * entirely on snapshot data — no live workbook objects are touched.
 */
export function collectFormulaInstances(snapshot: WorkbookSnapshot): FormulaInstance[] {
  const instances: FormulaInstance[] = [];

  for (const ws of snapshot.worksheets) {
    collectFromWorksheet(ws, instances);
  }

  return instances;
}

/**
 * Collect formula instances from a single worksheet snapshot.
 */
function collectFromWorksheet(ws: WorksheetSnapshot, out: FormulaInstance[]): void {
  for (const cell of ws.cells.values()) {
    if (cell.formulaKind === "none") {
      continue;
    }
    if (!cell.formula) {
      continue;
    }

    const instance = normalizeCell(ws, cell);
    if (instance) {
      out.push(instance);
    }
  }
}

/**
 * Normalize a single cell snapshot into a FormulaInstance.
 *
 * This is where the different formula types are unified:
 * - Normal formulas pass through directly.
 * - Shared-master and shared-slave formulas are both treated as normal
 *   for evaluation purposes (the slave's formula text is already translated).
 * - CSE formulas carry their target ref for the materialize layer.
 * - Dynamic-array formulas are flagged for spill handling.
 */
function normalizeCell(ws: WorksheetSnapshot, cell: CellSnapshot): FormulaInstance | null {
  const formula = cell.formula;
  if (!formula) {
    return null;
  }

  return {
    sheetName: ws.name,
    sheetId: ws.id,
    row: cell.row,
    col: cell.col,
    sourceText: formula,
    kind: cell.formulaKind,
    targetRef: cell.ref,
    isDynamicArray: cell.isDynamicArray ?? false
  };
}
