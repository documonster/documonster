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
  rowSetHidden,
  rowSetOutlineLevel,
  rowSetStyle,
  rowValues
} from "@excel/row";
import type { RowValues, Style } from "@excel/types";
import { getRow, rowSetValues } from "@excel/worksheet-core";
import type { WorksheetData } from "@excel/worksheet-core";

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
