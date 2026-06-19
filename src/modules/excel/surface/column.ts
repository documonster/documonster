/**
 * `Column` namespace surface — column-level operations addressed by key, letter,
 * or 1-based number.
 *
 * `import { Column } from "documonster/excel"` → `Column.setWidth(ws, "A", 20)`,
 * `Column.setHeader(ws, 1, "Name")`, `Column.setStyle(ws, "B", { numFmt })`.
 */
import type { ColumnHeaderValue } from "@excel/core/column";
import {
  columnHidden,
  columnOutlineLevel,
  columnSetHidden,
  columnSetOutlineLevel
} from "@excel/core/column";
import {
  getColumn,
  columnSetAlignment,
  columnSetBorder,
  columnSetFill,
  columnSetFont,
  columnSetHeader,
  columnSetKey,
  columnSetNumFmt,
  columnSetProtection
} from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import type { Alignment, Borders, Fill, Font, Protection, Style } from "@excel/types";

export type Sheet = WorksheetData;
type ColRef = string | number;

// --- width ---

export function getWidth(ws: Sheet, col: ColRef): number | undefined {
  return getColumn(ws, col).width;
}
export function setWidth(ws: Sheet, col: ColRef, width: number): void {
  getColumn(ws, col).width = width;
}

// --- header / key ---

export function getHeader(ws: Sheet, col: ColRef): ColumnHeaderValue | undefined {
  return getColumn(ws, col).header;
}
export function setHeader(ws: Sheet, col: ColRef, header: ColumnHeaderValue): void {
  columnSetHeader(getColumn(ws, col), header);
}
export function getKey(ws: Sheet, col: ColRef): string | undefined {
  return getColumn(ws, col).key;
}
export function setKey(ws: Sheet, col: ColRef, key: string): void {
  columnSetKey(getColumn(ws, col), key);
}

// --- visibility / outline ---

export function getHidden(ws: Sheet, col: ColRef): boolean {
  return columnHidden(getColumn(ws, col));
}
export function setHidden(ws: Sheet, col: ColRef, hidden: boolean): void {
  columnSetHidden(getColumn(ws, col), hidden);
}
export function getOutlineLevel(ws: Sheet, col: ColRef): number {
  return columnOutlineLevel(getColumn(ws, col));
}
export function setOutlineLevel(ws: Sheet, col: ColRef, level: number): void {
  columnSetOutlineLevel(getColumn(ws, col), level);
}

// --- style ---

export function getStyle(ws: Sheet, col: ColRef): Partial<Style> {
  return getColumn(ws, col).style;
}
export function setStyle(ws: Sheet, col: ColRef, style: Partial<Style>): void {
  const c = getColumn(ws, col);
  if (style.numFmt !== undefined) {
    columnSetNumFmt(c, style.numFmt as string);
  }
  if (style.font !== undefined) {
    columnSetFont(c, style.font as Partial<Font>);
  }
  if (style.alignment !== undefined) {
    columnSetAlignment(c, style.alignment as Partial<Alignment>);
  }
  if (style.border !== undefined) {
    columnSetBorder(c, style.border as Partial<Borders>);
  }
  if (style.fill !== undefined) {
    columnSetFill(c, style.fill as Fill);
  }
  if (style.protection !== undefined) {
    columnSetProtection(c, style.protection as Partial<Protection>);
  }
}
