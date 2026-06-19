import type { CellData } from "@excel/core/cell";
/**
 * `Row` namespace surface — row-level operations addressed by row number.
 *
 * `import { Row } from "documonster/excel"` → `Row.setHeight(ws, 1, 24)`,
 * `Row.setStyle(ws, 3, { font })`, `Row.getValues(ws, 2)`.
 */
import {
  rowGetStyle,
  rowGetValues,
  rowHidden,
  rowOutlineLevel,
  rowSetAlignment,
  rowSetBorder,
  rowSetFill,
  rowSetFont,
  rowSetHidden,
  rowSetOutlineLevel,
  rowSetStyle,
  rowValues
} from "@excel/core/row";
import {
  getRow,
  rowCommit,
  rowEachCell,
  rowGetCell,
  rowSetValues
} from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import type { Alignment, Borders, Fill, Font, RowValues, Style } from "@excel/types";

export type Sheet = WorksheetData;

// --- height ---

export function getHeight(ws: Sheet, row: number): number | undefined {
  return getRow(ws, row).height;
}
export function setHeight(ws: Sheet, row: number, height: number): void {
  const r = getRow(ws, row);
  r.height = height;
  r.customHeight = true;
}

// --- visibility / outline ---

export function getHidden(ws: Sheet, row: number): boolean {
  return rowHidden(getRow(ws, row));
}
export function setHidden(ws: Sheet, row: number, hidden: boolean): void {
  rowSetHidden(getRow(ws, row), hidden);
}
export function getOutlineLevel(ws: Sheet, row: number): number {
  return rowOutlineLevel(getRow(ws, row));
}
export function setOutlineLevel(ws: Sheet, row: number, level: number): void {
  rowSetOutlineLevel(getRow(ws, row), level);
}

// --- style ---

export function getStyle(ws: Sheet, row: number): Partial<Style> {
  return rowGetStyle(getRow(ws, row));
}
export function setStyle(ws: Sheet, row: number, style: Partial<Style>): void {
  rowSetStyle(getRow(ws, row), style);
}

// --- values ---

export function getValues(ws: Sheet, row: number): unknown[] {
  return rowGetValues(getRow(ws, row));
}
export function setValues(ws: Sheet, row: number, values: RowValues): void {
  rowSetValues(getRow(ws, row), values);
}
export function values(ws: Sheet, row: number): unknown[] {
  return rowValues(getRow(ws, row));
}

// --- individual style facets ---

export function setFont(ws: Sheet, row: number, value: Partial<Font> | undefined): void {
  rowSetFont(getRow(ws, row), value);
}
export function setAlignment(ws: Sheet, row: number, value: Partial<Alignment> | undefined): void {
  rowSetAlignment(getRow(ws, row), value);
}
export function setBorder(ws: Sheet, row: number, value: Partial<Borders> | undefined): void {
  rowSetBorder(getRow(ws, row), value);
}
export function setFill(ws: Sheet, row: number, value: Fill | undefined): void {
  rowSetFill(getRow(ws, row), value);
}

// --- cell access / iteration / commit ---

export function getCell(ws: Sheet, row: number, col: string | number): CellData {
  return rowGetCell(getRow(ws, row), col);
}
export function eachCell(
  ws: Sheet,
  row: number,
  optOrCallback: { includeEmpty?: boolean } | ((cell: CellData, colNumber: number) => void),
  maybeCallback?: (cell: CellData, colNumber: number) => void
): void {
  rowEachCell(getRow(ws, row), optOrCallback as never, maybeCallback);
}
export function commit(ws: Sheet, row: number): void {
  rowCommit(getRow(ws, row));
}
