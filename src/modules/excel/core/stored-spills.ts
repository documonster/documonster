/**
 * The spill a file stores for a dynamic-array formula.
 *
 * Excel stores a spilled dynamic array as its master — `ref` covering the spill — and every other cell of the spill as
 * a plain cached value (`<c r="B2"><v>0</v></c>`; Excel 16's save of `tmp/office-tests/r4-formula/r4-blank.xlsx`).
 * Those values are the formula's own spill, not data in its way. Read as ordinary values they blocked it, and the
 * master Excel had spilled computed `#SPILL!`.
 *
 * So a loaded sheet marks them as the master's spill ghosts ({@link markStoredSpills}), the way a calculation marks
 * the cells it spills into: the engine reuses them, a calculation that spills less clears the rest, and a value set
 * over one makes it the caller's again. Unlike a ghost a calculation made, one inside its master's stored `ref` is
 * written back as the cached value the file carried ({@link isStoredSpillCell}) — the file Excel wrote, unchanged by a
 * read and a write.
 */
import type { CellData } from "@excel/core/cell";
import { cellCol, cellType } from "@excel/core/cell";
import { Enums } from "@excel/core/enums";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { decodeRange } from "@excel/utils/address";
import { parseSpillCellKey, spillCellKeyFromId } from "@utils/spill-key";

interface StoredMaster {
  readonly isDynamicArray?: boolean;
  readonly shareType?: string;
  readonly ref?: string;
}

/** The 1-based bounds of a dynamic-array master's stored spill, or `undefined` when it covers only its own cell. */
function storedSpill(
  model: StoredMaster | undefined
): { top: number; left: number; bottom: number; right: number } | undefined {
  if (
    model?.isDynamicArray !== true ||
    model.shareType !== "array" ||
    typeof model.ref !== "string"
  ) {
    return undefined;
  }
  const { s, e } = decodeRange(model.ref);
  if (s.r === e.r && s.c === e.c) {
    return undefined;
  }
  return { top: s.r + 1, left: s.c + 1, bottom: e.r + 1, right: e.c + 1 };
}

function masterModel(cell: CellData | undefined): StoredMaster | undefined {
  return cell === undefined || cellType(cell) !== Enums.ValueType.Formula
    ? undefined
    : (cell._value.model as unknown as StoredMaster);
}

/** Mark the cached values each dynamic-array master's stored `ref` covers as that master's spill. */
export function markStoredSpills(ws: WorksheetData): void {
  for (const row of ws._rows) {
    for (const cell of row?.cells ?? []) {
      const spill = storedSpill(masterModel(cell));
      if (spill === undefined) {
        continue;
      }
      const owner = spillCellKeyFromId(ws.id, spill.top, spill.left);
      for (let r = spill.top; r <= spill.bottom; r++) {
        for (let c = spill.left; c <= spill.right; c++) {
          const target = ws._rows[r - 1]?.cells[c - 1];
          if ((r === spill.top && c === spill.left) || target === undefined) {
            continue;
          }
          const type = cellType(target);
          if (type !== Enums.ValueType.Formula && type !== Enums.ValueType.Merge) {
            target._formulaGhostOwner = owner;
          }
        }
      }
    }
  }
}

/** Whether the spill ghost `cell` lies inside its master's stored `ref`, so its value is written as Excel stores it. */
export function isStoredSpillCell(cell: CellData): boolean {
  const owner =
    cell._formulaGhostOwner === undefined ? undefined : parseSpillCellKey(cell._formulaGhostOwner);
  const ws = cell.row?.worksheet as unknown as WorksheetData | undefined;
  if (owner === undefined || ws === undefined || owner.worksheetId !== ws.id) {
    return false;
  }
  const masterRow = owner.row;
  const masterCol = owner.col;
  const spill = storedSpill(masterModel(ws._rows[masterRow - 1]?.cells[masterCol - 1]));
  const row = cell.row.number;
  const col = cellCol(cell);
  return (
    spill !== undefined &&
    spill.top === masterRow &&
    spill.left === masterCol &&
    row >= spill.top &&
    row <= spill.bottom &&
    col >= spill.left &&
    col <= spill.right
  );
}
