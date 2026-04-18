/**
 * Shared types for the formula engine's workbook interface.
 *
 * These interfaces decouple the formula engine from the concrete
 * Workbook/Worksheet/Cell classes, preventing circular imports.
 */

import type { Worksheet } from "@excel/worksheet";

// ============================================================================
// Workbook Interface
// ============================================================================

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
  worksheets: Worksheet[];
  getWorksheet(id?: number | string): Worksheet | undefined;
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
