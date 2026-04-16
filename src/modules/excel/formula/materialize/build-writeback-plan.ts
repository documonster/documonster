/**
 * Build Writeback Plan — Convert evaluation results into a WritebackPlan.
 *
 * This module takes the evaluation results from the evaluator and
 * produces a declarative `WritebackPlan` that describes all cell mutations.
 * The plan is then applied by `apply-writeback-plan.ts`.
 *
 * ## Responsibilities
 *
 * 1. Classify each formula's result as scalar, CSE, or dynamic-array.
 * 2. Check spill availability and detect #SPILL! conflicts.
 * 3. Generate cleanup operations for stale ghost cells.
 * 4. Track spill regions and ghost snapshots for persistence.
 *
 * ## Key Principle
 *
 * This module does NOT touch any live workbook objects. It reads only
 * from the `WorkbookSnapshot` and the evaluation results.
 */

import { parseRefRange } from "../compile/address-utils";
import type { CompiledFormula } from "../compile/compiled-formula";
import type { FormulaInstance } from "../integration/formula-instance";
import type { WorkbookSnapshot, SnapshotCellValue } from "../integration/workbook-snapshot";
import {
  snapshotCellKey,
  spillCellKeyFromId,
  formulaCellKey
} from "../integration/workbook-snapshot";
import type { EvalSession } from "../runtime/evaluator";
import type { RuntimeValue, ScalarValue, ArrayValue } from "../runtime/values";
import { RVKind } from "../runtime/values";
import type {
  WritebackPlan,
  WriteOperation,
  SpillWrite,
  SpillErrorWrite,
  CleanupWrite,
  CSEWrite,
  ScalarWrite,
  PreserveWrite
} from "./writeback-plan";

// ============================================================================
// Spill Tracking State
// ============================================================================

/**
 * Persistent spill metadata from a previous calculation.
 * Keyed by `"ws:<id>!row:col"` of the source formula cell.
 */
interface SpillRegionInfo {
  readonly worksheetId: number;
  readonly sourceRow: number;
  readonly sourceCol: number;
  readonly rows: number;
  readonly cols: number;
}

/**
 * Tracks which cells are ghost cells (spill targets).
 * Key: ghost cell key → source cell key.
 */
type GhostMap = Map<string, string>;

// ============================================================================
// Build Writeback Plan
// ============================================================================

/**
 * Build a complete `WritebackPlan` from evaluation results.
 *
 * @param snapshot - The workbook snapshot
 * @param compiled - All compiled formulas in evaluation order
 * @param results - Raw evaluation results, keyed by formula cell key
 * @param session - The eval session (for cache access)
 * @param previousSpills - Persistent spill regions from previous calculation
 * @param previousGhosts - Persistent ghost snapshots from previous calculation
 */
export function buildWritebackPlan(
  snapshot: WorkbookSnapshot,
  compiled: readonly CompiledFormula[],
  results: ReadonlyMap<string, RuntimeValue>,
  session: EvalSession,
  previousSpills: ReadonlyMap<string, SpillRegionInfo>,
  previousGhosts: ReadonlyMap<string, unknown>
): WritebackPlan {
  const operations: WriteOperation[] = [];
  const spillRegions = new Map<string, SpillRegionInfo>();
  const ghostSnapshots = new Map<string, SnapshotCellValue>();
  const removedSpillKeys: string[] = [];
  const cseSessionUpdates = new Map<string, SnapshotCellValue>();

  // Build a set of current formula cell keys (using worksheet-id-based keys)
  const formulaKeys = new Set<string>();
  for (const cf of compiled) {
    formulaKeys.add(spillCellKeyFromId(cf.instance.sheetId, cf.instance.row, cf.instance.col));
  }

  // Build ghost map from previous spills (validated against snapshot)
  const ghostMap: GhostMap = new Map();
  for (const [srcKey, region] of previousSpills) {
    for (let r = 0; r < region.rows; r++) {
      for (let c = 0; c < region.cols; c++) {
        if (r === 0 && c === 0) {
          continue;
        }
        const targetRow = region.sourceRow + r;
        const targetCol = region.sourceCol + c;
        const targetKey = spillCellKeyFromId(region.worksheetId, targetRow, targetCol);
        // Validate ghost cell is still unmodified
        const ws = snapshot.worksheetsById.get(region.worksheetId);
        if (ws) {
          const cell = ws.cells.get(snapshotCellKey(targetRow, targetCol));
          if (isGhostUnmodified(cell, targetKey, previousGhosts)) {
            ghostMap.set(targetKey, srcKey);
          }
        }
      }
    }
  }

  // Clean up stale spill regions (source formula no longer exists)
  for (const [srcKey, region] of previousSpills) {
    if (!formulaKeys.has(srcKey)) {
      const cleanupCells = collectStaleGhosts(region, previousGhosts, snapshot);
      if (cleanupCells.length > 0) {
        const ws = snapshot.worksheetsById.get(region.worksheetId);
        if (ws) {
          operations.push({
            type: "cleanup",
            sheetName: ws.name,
            sheetId: region.worksheetId,
            cells: cleanupCells
          } satisfies CleanupWrite);
        }
      }
      removedSpillKeys.push(srcKey);
    }
  }

  // Process each compiled formula's result
  for (const cf of compiled) {
    const inst = cf.instance;
    const fKey = formulaCellKey(inst.sheetName, inst.row, inst.col);
    const result = results.get(fKey);
    if (result === undefined) {
      continue;
    }

    const isCSE = inst.kind === "cse" && inst.targetRef;
    const isDynamic = inst.isDynamicArray || cf.isDynamicArrayFunction;

    if (isCSE) {
      // CSE array formula
      const op = buildCSEWrite(inst, result, session, cseSessionUpdates);
      if (op) {
        operations.push(op);
      }
    } else if (
      isDynamic &&
      result.kind === RVKind.Array &&
      (result.height > 1 || result.width > 1)
    ) {
      // Dynamic array formula with array result
      const srcKey = spillCellKeyFromId(inst.sheetId, inst.row, inst.col);

      // Check spill availability
      const spillResult = buildSpillWrite(
        inst,
        result,
        srcKey,
        ghostMap,
        previousSpills.get(srcKey),
        previousGhosts,
        snapshot,
        spillRegions,
        ghostSnapshots,
        operations
      );

      if (spillResult === "error") {
        operations.push({
          type: "spill-error",
          sheetName: inst.sheetName,
          sheetId: inst.sheetId,
          row: inst.row,
          col: inst.col
        } satisfies SpillErrorWrite);
      }
    } else {
      // Scalar result (or 1x1 array)
      const scalar = scalarFromResult(result);

      // If this formula previously had a spill region, clean it up
      const srcKey = spillCellKeyFromId(inst.sheetId, inst.row, inst.col);
      const prevRegion = previousSpills.get(srcKey);
      if (prevRegion) {
        const cleanupCells = collectStaleGhosts(prevRegion, previousGhosts, snapshot);
        if (cleanupCells.length > 0) {
          const ws = snapshot.worksheetsById.get(inst.sheetId);
          if (ws) {
            operations.push({
              type: "cleanup",
              sheetName: ws.name,
              sheetId: inst.sheetId,
              cells: cleanupCells
            } satisfies CleanupWrite);
          }
        }
        removedSpillKeys.push(srcKey);
      }

      if (shouldPreserve(scalar, inst, snapshot)) {
        operations.push({
          type: "preserve",
          sheetName: inst.sheetName,
          row: inst.row,
          col: inst.col
        } satisfies PreserveWrite);
      } else {
        operations.push({
          type: "scalar",
          sheetName: inst.sheetName,
          row: inst.row,
          col: inst.col,
          value: runtimeToSnapshotValue(scalar)
        } satisfies ScalarWrite);
      }
    }
  }

  return {
    operations,
    spillState: {
      spillRegions,
      ghostSnapshots,
      removedSpillKeys
    },
    sessionDelta: {
      cseUpdates: cseSessionUpdates
    }
  };
}

// ============================================================================
// CSE Write
// ============================================================================

function buildCSEWrite(
  inst: FormulaInstance,
  result: RuntimeValue,
  session: EvalSession,
  cseUpdates: Map<string, SnapshotCellValue>
): CSEWrite | null {
  const ref = inst.targetRef;
  if (!ref) {
    return null;
  }

  const range = parseRefRange(ref);
  if (!range) {
    return null;
  }

  if (result.kind === RVKind.Array) {
    const results: SnapshotCellValue[][] = [];
    const numRows = range.bottom - range.top + 1;
    const numCols = range.right - range.left + 1;
    for (let r = 0; r < numRows; r++) {
      const row: SnapshotCellValue[] = [];
      for (let c = 0; c < numCols; c++) {
        const val = result.rows[r]?.[c] ?? { kind: RVKind.Blank };
        const sv = runtimeToSnapshotValue(val);
        row.push(sv);
        // Update session cache for CSE cells
        const cellKey = formulaCellKey(inst.sheetName, range.top + r, range.left + c);
        cseUpdates.set(cellKey, sv);
      }
      results.push(row);
    }
    return {
      type: "cse",
      sheetName: inst.sheetName,
      top: range.top,
      left: range.left,
      bottom: range.bottom,
      right: range.right,
      results
    };
  }

  // Scalar fill
  const scalarVal = runtimeToSnapshotValue(scalarFromResult(result));
  const numRows = range.bottom - range.top + 1;
  const numCols = range.right - range.left + 1;
  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const cellKey = formulaCellKey(inst.sheetName, range.top + r, range.left + c);
      cseUpdates.set(cellKey, scalarVal);
    }
  }
  return {
    type: "cse",
    sheetName: inst.sheetName,
    top: range.top,
    left: range.left,
    bottom: range.bottom,
    right: range.right,
    results: [],
    scalarFill: scalarVal
  };
}

// ============================================================================
// Spill Write
// ============================================================================

function buildSpillWrite(
  inst: FormulaInstance,
  arr: ArrayValue,
  srcKey: string,
  ghostMap: GhostMap,
  previousRegion: SpillRegionInfo | undefined,
  previousGhosts: ReadonlyMap<string, unknown>,
  snapshot: WorkbookSnapshot,
  spillRegions: Map<string, SpillRegionInfo>,
  ghostSnapshotsOut: Map<string, SnapshotCellValue>,
  operations: WriteOperation[]
): "ok" | "error" {
  const ws = snapshot.worksheetsByName.get(inst.sheetName.toLowerCase());
  if (!ws) {
    return "error";
  }

  // 1x1 result: just write scalar, no spilling
  if (arr.height <= 1 && arr.width <= 1) {
    const val =
      arr.height > 0 && arr.width > 0 ? arr.rows[0][0] : ({ kind: RVKind.Blank } as ScalarValue);
    operations.push({
      type: "scalar",
      sheetName: inst.sheetName,
      row: inst.row,
      col: inst.col,
      value: runtimeToSnapshotValue(val)
    } satisfies ScalarWrite);
    // Clean up previous spill if any
    if (previousRegion) {
      const cleanups = collectStaleGhosts(previousRegion, previousGhosts, snapshot);
      if (cleanups.length > 0) {
        operations.push({
          type: "cleanup",
          sheetName: inst.sheetName,
          sheetId: inst.sheetId,
          cells: cleanups
        } satisfies CleanupWrite);
      }
    }
    return "ok";
  }

  // Check spill availability: verify all target ghost cells are unoccupied
  for (let r = 0; r < arr.height; r++) {
    for (let c = 0; c < arr.width; c++) {
      if (r === 0 && c === 0) {
        continue; // Source cell is always available
      }
      const targetRow = inst.row + r;
      const targetCol = inst.col + c;
      const targetKey = spillCellKeyFromId(inst.sheetId, targetRow, targetCol);

      // Check if the cell is a ghost from a previous spill of this same formula
      if (ghostMap.has(targetKey) && ghostMap.get(targetKey) === srcKey) {
        // It's our own ghost — check if user has modified it
        const cell = ws.cells.get(snapshotCellKey(targetRow, targetCol));
        if (!isGhostUnmodified(cell, targetKey, previousGhosts)) {
          return "error"; // User wrote into our ghost → #SPILL!
        }
        continue;
      }

      // Check if the cell is occupied by another formula or has a value
      const cell = ws.cells.get(snapshotCellKey(targetRow, targetCol));
      if (cell && cell.value !== null && cell.formulaKind === "none") {
        // Non-empty, non-formula cell → blocked
        return "error";
      }
      if (cell && cell.formulaKind !== "none") {
        // Formula cell → blocked (not our ghost)
        return "error";
      }
    }
  }

  // Clean up previous spill region if size changed
  if (previousRegion) {
    const cleanups = collectStaleGhosts(previousRegion, previousGhosts, snapshot);
    if (cleanups.length > 0) {
      operations.push({
        type: "cleanup",
        sheetName: inst.sheetName,
        sheetId: inst.sheetId,
        cells: cleanups
      } satisfies CleanupWrite);
    }
  }

  // Build spill write operation
  const results: SnapshotCellValue[][] = [];
  for (let r = 0; r < arr.height; r++) {
    const row: SnapshotCellValue[] = [];
    for (let c = 0; c < arr.width; c++) {
      const val = arr.rows[r]?.[c] ?? ({ kind: RVKind.Blank } as ScalarValue);
      const sv = runtimeToSnapshotValue(val);
      row.push(sv);

      // Record ghost snapshot
      if (r !== 0 || c !== 0) {
        const targetKey = spillCellKeyFromId(inst.sheetId, inst.row + r, inst.col + c);
        ghostSnapshotsOut.set(targetKey, sv);
      }
    }
    results.push(row);
  }

  operations.push({
    type: "spill",
    sheetName: inst.sheetName,
    sheetId: inst.sheetId,
    row: inst.row,
    col: inst.col,
    results
  } satisfies SpillWrite);

  // Record spill region
  spillRegions.set(srcKey, {
    worksheetId: inst.sheetId,
    sourceRow: inst.row,
    sourceCol: inst.col,
    rows: arr.height,
    cols: arr.width
  });

  return "ok";
}

// ============================================================================
// Helpers
// ============================================================================

function collectStaleGhosts(
  region: SpillRegionInfo,
  previousGhosts: ReadonlyMap<string, unknown>,
  snapshot: WorkbookSnapshot
): { row: number; col: number }[] {
  const cells: { row: number; col: number }[] = [];
  const ws = snapshot.worksheetsById.get(region.worksheetId);
  if (!ws) {
    return cells;
  }
  for (let r = 0; r < region.rows; r++) {
    for (let c = 0; c < region.cols; c++) {
      if (r === 0 && c === 0) {
        continue;
      }
      const targetRow = region.sourceRow + r;
      const targetCol = region.sourceCol + c;
      const targetKey = spillCellKeyFromId(region.worksheetId, targetRow, targetCol);
      const cell = ws.cells.get(snapshotCellKey(targetRow, targetCol));
      if (isGhostUnmodified(cell, targetKey, previousGhosts)) {
        cells.push({ row: targetRow, col: targetCol });
      }
    }
  }
  return cells;
}

function isGhostUnmodified(
  cell: { value: SnapshotCellValue; formulaKind: string } | undefined,
  ghostKey: string,
  previousGhosts: ReadonlyMap<string, unknown>
): boolean {
  if (!cell) {
    return true;
  }
  if (cell.value === null) {
    return true;
  }
  if (cell.formulaKind !== "none") {
    return false;
  }
  const snapshot = previousGhosts.get(ghostKey);
  if (snapshot === undefined) {
    return true; // No snapshot — assume unmodified (conservative)
  }
  return snapshotValuesEqual(cell.value, snapshot);
}

/**
 * Structural comparison for snapshot values.
 * Handles error objects (which are reference types and fail `===` comparison)
 * as well as primitives.
 */
function snapshotValuesEqual(a: SnapshotCellValue, b: unknown): boolean {
  if (a === b) {
    return true;
  }
  // Both are error objects — compare by error code
  if (
    a !== null &&
    typeof a === "object" &&
    "error" in a &&
    b !== null &&
    typeof b === "object" &&
    "error" in (b as Record<string, unknown>)
  ) {
    return a.error === (b as { error: string }).error;
  }
  return false;
}

function scalarFromResult(v: RuntimeValue): ScalarValue {
  if (v.kind === RVKind.Array) {
    if (v.height > 0 && v.width > 0) {
      return v.rows[0][0];
    }
    return { kind: RVKind.Blank };
  }
  if (
    v.kind === RVKind.Blank ||
    v.kind === RVKind.Number ||
    v.kind === RVKind.String ||
    v.kind === RVKind.Boolean ||
    v.kind === RVKind.Error
  ) {
    return v;
  }
  return { kind: RVKind.Blank };
}

function runtimeToSnapshotValue(v: ScalarValue | RuntimeValue): SnapshotCellValue {
  switch (v.kind) {
    case RVKind.Blank:
    case RVKind.MissingArg:
      return null;
    case RVKind.Number:
      return v.value;
    case RVKind.String:
      return v.value;
    case RVKind.Boolean:
      return v.value;
    case RVKind.Error:
      return { error: v.code };
    case RVKind.Array:
      // Take top-left
      if (v.height > 0 && v.width > 0) {
        return runtimeToSnapshotValue(v.rows[0][0]);
      }
      return null;
    default:
      return null;
  }
}

function shouldPreserve(
  computed: ScalarValue,
  inst: FormulaInstance,
  snapshot: WorkbookSnapshot
): boolean {
  if (computed.kind !== RVKind.Error || computed.code !== "#NAME?") {
    return false;
  }
  // Check if cell has a cached result in the snapshot
  const ws = snapshot.worksheetsByName.get(inst.sheetName.toLowerCase());
  if (!ws) {
    return false;
  }
  const cell = ws.cells.get(snapshotCellKey(inst.row, inst.col));
  return cell?.cachedResult !== undefined && cell.cachedResult !== null;
}
