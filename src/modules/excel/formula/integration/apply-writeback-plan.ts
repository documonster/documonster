/**
 * Apply Writeback Plan — Write formula results back to the live workbook.
 *
 * This is the **only** file that mutates live workbook/cell objects in the
 * new engine pipeline. It reads a `WritebackPlan` and applies each operation
 * to the workbook.
 *
 * ## Responsibilities
 *
 * 1. Apply `ScalarWrite` → set `cell.result`.
 * 2. Apply `CSEWrite` → distribute array across CSE range, update session cache.
 * 3. Apply `SpillWrite` → write source + ghost cells, update persistent maps.
 * 4. Apply `SpillErrorWrite` → set #SPILL! on source cell.
 * 5. Apply `CleanupWrite` → clear stale ghost cells.
 * 6. Apply `PreserveWrite` → no-op (cell keeps existing result).
 * 7. Update persistent spill maps and ghost snapshot maps.
 */

import type { FormulaResult } from "@excel/cell";
import { Enums } from "@excel/enums";
import type { CellErrorValue } from "@excel/types";

import { getGhostSnapshots, getPersistentSpillMap } from "../materialize/spill-engine";
import type { WorkbookLike } from "../materialize/types";
import type {
  WritebackPlan,
  WriteOperation,
  ScalarWrite,
  CSEWrite,
  SpillWrite,
  SpillErrorWrite,
  CleanupWrite
} from "../materialize/writeback-plan";
import type { SnapshotCellValue, SnapshotErrorValue } from "./workbook-snapshot";

// ============================================================================
// Apply Writeback Plan
// ============================================================================

/**
 * Apply a `WritebackPlan` to a live workbook.
 *
 * This mutates the workbook in-place. Operations are applied in order.
 * After all operations, persistent spill metadata is updated.
 */
export function applyWritebackPlan(workbook: WorkbookLike, plan: WritebackPlan): void {
  // Apply each operation
  for (const op of plan.operations) {
    applyOperation(workbook, op);
  }

  // Update persistent spill metadata
  const persistentSpills = getPersistentSpillMap(workbook);
  const ghostSnapshots = getGhostSnapshots(workbook);

  // Remove stale spill entries
  for (const key of plan.spillState.removedSpillKeys) {
    persistentSpills.delete(key);
  }

  // Update spill regions
  for (const [key, region] of plan.spillState.spillRegions) {
    persistentSpills.set(key, region);
  }

  // Update ghost value snapshots
  for (const [key, value] of plan.spillState.ghostSnapshots) {
    ghostSnapshots.set(key, snapshotValueToRaw(value));
  }
}

// ============================================================================
// Apply Individual Operations
// ============================================================================

function applyOperation(workbook: WorkbookLike, op: WriteOperation): void {
  switch (op.type) {
    case "scalar":
      applyScalarWrite(workbook, op);
      break;
    case "cse":
      applyCSEWrite(workbook, op);
      break;
    case "spill":
      applySpillWrite(workbook, op);
      break;
    case "spill-error":
      applySpillErrorWrite(workbook, op);
      break;
    case "cleanup":
      applyCleanupWrite(workbook, op);
      break;
    case "preserve":
      // No-op: keep existing cached result
      break;
  }
}

function applyScalarWrite(workbook: WorkbookLike, op: ScalarWrite): void {
  const ws = workbook.getWorksheet(op.sheetName);
  if (!ws) {
    return;
  }
  const cell = ws.findCell(op.row, op.col);
  if (!cell) {
    return;
  }
  cell.result = snapshotValueToResult(op.value);
}

function applyCSEWrite(workbook: WorkbookLike, op: CSEWrite): void {
  const ws = workbook.getWorksheet(op.sheetName);
  if (!ws) {
    return;
  }

  const numRows = op.bottom - op.top + 1;
  const numCols = op.right - op.left + 1;

  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const targetRow = op.top + r;
      const targetCol = op.left + c;
      const targetCell = ws.getCell(targetRow, targetCol);

      if (targetCell.type === Enums.ValueType.Formula) {
        if (op.scalarFill !== undefined) {
          targetCell.result = snapshotValueToResult(op.scalarFill);
        } else {
          const val = op.results[r]?.[c] ?? null;
          targetCell.result = snapshotValueToResult(val);
        }
      }
    }
  }
}

function applySpillWrite(workbook: WorkbookLike, op: SpillWrite): void {
  const ws = workbook.getWorksheet(op.sheetName);
  if (!ws) {
    return;
  }

  const numRows = op.results.length;
  const numCols = op.results[0]?.length ?? 0;

  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const targetRow = op.row + r;
      const targetCol = op.col + c;
      const val = op.results[r]?.[c] ?? null;

      if (r === 0 && c === 0) {
        // Source cell: set result
        const sourceCell = ws.findCell(targetRow, targetCol);
        if (sourceCell) {
          sourceCell.result = snapshotValueToResult(val);
        }
      } else {
        // Ghost cell: set value (not result)
        const targetCell = ws.getCell(targetRow, targetCol);
        targetCell.value = snapshotValueToRaw(val);
      }
    }
  }
}

function applySpillErrorWrite(workbook: WorkbookLike, op: SpillErrorWrite): void {
  const ws = workbook.getWorksheet(op.sheetName);
  if (!ws) {
    return;
  }
  const cell = ws.findCell(op.row, op.col);
  if (cell) {
    cell.result = { error: "#SPILL!" } as FormulaResult;
  }
}

function applyCleanupWrite(workbook: WorkbookLike, op: CleanupWrite): void {
  const ws = workbook.getWorksheet(op.sheetName);
  if (!ws) {
    return;
  }
  for (const { row, col } of op.cells) {
    const cell = ws.findCell(row, col);
    if (cell) {
      cell.value = null;
    }
  }
}

// ============================================================================
// Value Conversion (Snapshot → Live)
// ============================================================================

/**
 * Convert a snapshot cell value to a `FormulaResult` suitable for `cell.result`.
 */
function snapshotValueToResult(val: SnapshotCellValue): FormulaResult {
  if (val === null) {
    return 0;
  }
  if (typeof val === "number" || typeof val === "string" || typeof val === "boolean") {
    return val;
  }
  if (isSnapshotError(val)) {
    return val as unknown as FormulaResult;
  }
  return 0;
}

/**
 * Convert a snapshot cell value to a raw value suitable for `cell.value`.
 */
function snapshotValueToRaw(
  val: SnapshotCellValue
): number | string | boolean | CellErrorValue | null {
  if (val === null) {
    return null;
  }
  if (typeof val === "number" || typeof val === "string" || typeof val === "boolean") {
    return val;
  }
  if (isSnapshotError(val)) {
    return val as unknown as CellErrorValue;
  }
  return null;
}

function isSnapshotError(val: SnapshotCellValue): val is SnapshotErrorValue {
  return val !== null && typeof val === "object" && "error" in val;
}
