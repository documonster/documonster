/**
 * Structural interfaces the formula engine consumes.
 *
 * The engine talks to these interfaces only — never to the concrete
 * `Workbook` / `Worksheet` / `Cell` classes from the `excel` module. Any
 * host that wants to drive the engine (including the Node/browser
 * Workbook in `@excel/*`) must implement these shapes.
 *
 * This is what keeps the Layer 3 ↔ Layer 4 dependency one-way: the
 * `excel` module imports from `formula`, never the reverse.
 */

// ============================================================================
// Cell value / error / result shapes
// ============================================================================

/**
 * Excel-style error value carried inside a cell.
 *
 * Mirrors `CellErrorValue` in `@excel/types` so hosts can pass their
 * concrete error objects straight through.
 */
export interface CellErrorValueLike {
  error: string;
}

/**
 * The value the engine writes back as `cell.result` for formula cells.
 */
export type FormulaResultLike = number | string | boolean | Date | CellErrorValueLike | undefined;

// ============================================================================
// ValueType — numeric mirror of `@excel/enums` ValueType
// ============================================================================

/**
 * Numeric cell-type tag exposed by host cells. The engine only compares
 * against `Null` and `Formula`; any other value is treated as a scalar
 * literal.
 *
 * Kept as inline numeric literals (not an enum) so this file stays free
 * of runtime dependencies. The `const` object and `type` alias share a
 * name via TypeScript's declaration merging — the value form
 * (`CellValueTypeLike.Null`, `CellValueTypeLike.Formula`) is used at
 * comparison sites, the type form annotates `CellLike.type`.
 */
export const CellValueTypeLike = {
  Null: 0,
  Formula: 6
} as const;

export type CellValueTypeLike = number;

// ============================================================================
// Cell / Row / Worksheet / Workbook interfaces
// ============================================================================

/**
 * The portion of a cell's persisted `model` the engine inspects for
 * classifying a formula (shared / CSE / dynamic array, etc.).
 */
export interface CellModelLike {
  readonly shareType?: string;
  readonly ref?: string;
  readonly formula?: string;
  readonly sharedFormula?: string;
  readonly isDynamicArray?: boolean;
}

/**
 * Minimal cell surface consumed by the engine. A host cell must allow
 * both reading (for snapshot) and mutation (for writeback).
 */
export interface CellLike {
  readonly row: number;
  readonly col: number;
  readonly type: CellValueTypeLike;
  readonly formula?: string;
  readonly model: CellModelLike;
  /**
   * Host cells can carry any representation (hyperlinks, rich text,
   * checkboxes, dates, scalars, errors, …). The engine narrows to plain
   * scalars internally during snapshot construction, so this is typed
   * as `unknown` to stay structurally compatible with any host shape.
   */
  value: unknown;
  result: FormulaResultLike;
}

export interface RowLike {
  readonly hidden?: boolean;
  /**
   * Excel's `eachCell` accepts either a bare callback or options + callback.
   */
  eachCell(callback: (cell: CellLike, colNumber: number) => void): void;
  eachCell(
    opts: { includeEmpty?: boolean },
    callback: (cell: CellLike, colNumber: number) => void
  ): void;
}

export interface DimensionsLike {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export interface TableColumnLike {
  readonly name: string;
}

export interface TableDefinitionLike {
  readonly name?: string;
  readonly displayName?: string;
  readonly tl?: { readonly row: number; readonly col: number };
  readonly columns?: readonly TableColumnLike[];
  readonly rows?: readonly unknown[];
  readonly headerRow?: boolean;
  readonly totalsRow?: boolean;
}

export interface TableRefLike {
  readonly table?: TableDefinitionLike;
}

export interface WorksheetLike {
  readonly id: number;
  readonly name: string;
  readonly dimensions: DimensionsLike | null;
  /**
   * Excel's `eachRow` accepts either a bare callback or an options bag
   * followed by a callback. Mirror both signatures so concrete
   * `Worksheet` classes remain structurally assignable.
   */
  eachRow(callback: (row: RowLike, rowNumber: number) => void): void;
  eachRow(
    opts: { includeEmpty?: boolean },
    callback: (row: RowLike, rowNumber: number) => void
  ): void;
  findCell(row: number, col: number): CellLike | undefined;
  getCell(row: number, col: number): CellLike;
  getTables?(): TableRefLike[];
}

/**
 * A complete defined name entry with all details.
 */
export interface DefinedNameEntry {
  name: string;
  ranges: string[];
  localSheetId?: number;
  formulaExpression?: string;
}

export interface DefinedNamesLike {
  /**
   * Return all defined name entries with full details.
   * Each entry is self-contained — no second lookup needed.
   * Same bare name may appear multiple times with different scopes.
   */
  getAllEntries(): DefinedNameEntry[];
  /** Enumerate name+scope pairs (lightweight). */
  getAllNames(): { name: string; localSheetId?: number }[];
}

export interface WorkbookLike {
  worksheets: WorksheetLike[];
  getWorksheet(id?: number | string): WorksheetLike | undefined;
  definedNames?: DefinedNamesLike;
  /** Calculation properties — used for iterative calculation settings. */
  calcProperties?: {
    fullCalcOnLoad?: boolean;
    /** Enable iterative calculation for circular references. */
    iterate?: boolean;
    /** Maximum number of iterations (default 100). */
    iterateCount?: number;
    /** Maximum change threshold for convergence (default 0.001). */
    iterateDelta?: number;
  };
  /** Workbook properties including date system. */
  properties?: {
    date1904?: boolean;
  };
}

// ============================================================================
// Spill Region
// ============================================================================

/**
 * Tracks a spill region: the source formula cell and the range of cells it
 * has spilled into. Used for cleanup when a formula is re-evaluated with
 * different-sized results.
 */
export interface SpillRegion {
  /** The worksheet id — stable across renames */
  readonly worksheetId: number;
  /** The source formula cell's row */
  readonly sourceRow: number;
  /** The source formula cell's col */
  readonly sourceCol: number;
  /** Number of rows in the spill (including source) */
  readonly rows: number;
  /** Number of cols in the spill (including source) */
  readonly cols: number;
}
