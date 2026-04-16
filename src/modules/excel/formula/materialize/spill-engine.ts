/**
 * Spill Engine — Persistent state management and dynamic array detection.
 *
 * This module manages persistent spill metadata that survives across
 * `calculateFormulas()` invocations. It also provides constants and
 * helpers for dynamic array formula detection.
 *
 * ## Persistent State
 *
 * Spill regions and ghost cell snapshots are stored in WeakMaps keyed by
 * the workbook object. This allows:
 * - State to persist across multiple `calculateFormulas()` calls
 * - Automatic garbage collection when the workbook is no longer referenced
 *
 * ## Usage
 *
 * - `getPersistentSpillMap()` / `getGhostSnapshots()` — read/write by
 *   `calculate-formulas-impl.ts` and `apply-writeback-plan.ts`
 * - `DYNAMIC_ARRAY_FUNCTIONS` / `isDynamicArrayFormula()` — used by
 *   the evaluation and materialization layers to detect formulas that
 *   should produce spill results
 */

import type { AstNode } from "../syntax/ast";
import { NodeType } from "../syntax/ast";
import type { SpillRegion, WorkbookLike } from "./types";

// ============================================================================
// Persistent State (WeakMap keyed by workbook — survives across invocations)
// ============================================================================

/**
 * Persistent spill metadata: survives across calculateFormulas invocations.
 * Key: "ws:<id>!row:col" of the source cell → SpillRegion.
 * Stored in a WeakMap keyed by the workbook object to allow GC.
 */
const persistentSpillRegions = new WeakMap<WeakKey, Map<string, SpillRegion>>();

export function getPersistentSpillMap(workbook: WorkbookLike): Map<string, SpillRegion> {
  let map = persistentSpillRegions.get(workbook);
  if (!map) {
    map = new Map();
    persistentSpillRegions.set(workbook, map);
  }
  return map;
}

/**
 * Persistent snapshot of values written to ghost (spill target) cells.
 * Key: ghost cell key "ws:<id>!row:col" → raw value written by spill.
 * Used to detect whether a ghost cell has been modified by the user since
 * the last calculation.
 */
const persistentGhostSnapshots = new WeakMap<WeakKey, Map<string, unknown>>();

export function getGhostSnapshots(workbook: WorkbookLike): Map<string, unknown> {
  let map = persistentGhostSnapshots.get(workbook);
  if (!map) {
    map = new Map();
    persistentGhostSnapshots.set(workbook, map);
  }
  return map;
}

// ============================================================================
// Dynamic Array Detection
// ============================================================================

/**
 * Set of function names that naturally produce arrays and should trigger
 * spill behavior even when isDynamicArray is not explicitly set.
 */
export const DYNAMIC_ARRAY_FUNCTIONS = new Set([
  "FILTER",
  "_XLFN._XLWS.FILTER",
  "SORT",
  "_XLFN._XLWS.SORT",
  "UNIQUE",
  "_XLFN._XLWS.UNIQUE",
  "SORTBY",
  "_XLFN._XLWS.SORTBY",
  "SEQUENCE",
  "_XLFN.SEQUENCE",
  "RANDARRAY",
  "_XLFN.RANDARRAY",
  "TOCOL",
  "_XLFN.TOCOL",
  "TOROW",
  "_XLFN.TOROW",
  "CHOOSEROWS",
  "_XLFN.CHOOSEROWS",
  "CHOOSECOLS",
  "_XLFN.CHOOSECOLS",
  "VSTACK",
  "_XLFN.VSTACK",
  "HSTACK",
  "_XLFN.HSTACK",
  "WRAPROWS",
  "_XLFN.WRAPROWS",
  "WRAPCOLS",
  "_XLFN.WRAPCOLS",
  "EXPAND",
  "_XLFN.EXPAND",
  "TAKE",
  "_XLFN.TAKE",
  "DROP",
  "_XLFN.DROP"
]);

/**
 * Check if an AST's top-level function is a known dynamic array function.
 */
export function isDynamicArrayFormula(ast: AstNode): boolean {
  if (ast.type === NodeType.FunctionCall) {
    return DYNAMIC_ARRAY_FUNCTIONS.has(ast.name.toUpperCase());
  }
  return false;
}
