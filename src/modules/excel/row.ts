import {
  type CellData,
  type CellModel,
  cellCol,
  cellGetModel,
  cellGetValue,
  cellType
} from "@excel/cell";
import { Enums } from "@excel/enums";
import type {
  Style,
  NumFmt,
  Font,
  Alignment,
  Protection,
  Borders,
  Fill,
  CellValue,
  RowBreak
} from "@excel/types";
import type { Worksheet } from "@excel/worksheet";

interface RowDimensions {
  min: number;
  max: number;
}

export interface RowModel {
  cells: CellModel[];
  number: number;
  min: number;
  max: number;
  height?: number;
  customHeight?: boolean;
  style: Partial<Style>;
  hidden: boolean;
  outlineLevel: number;
  collapsed: boolean;
  dyDescent?: number;
}

/**
 * Plain-data row record. The entire state of a row — no class.
 * All operations are free functions in the {@link Row} namespace.
 */
export interface RowData {
  worksheet: Worksheet;
  number: number;
  cells: CellData[];
  style: Partial<Style>;
  hidden?: boolean;
  outlineLevel?: number;
  height?: number;
  customHeight?: boolean;
  dyDescent?: number;
}

function applyStyle<K extends keyof Style>(r: RowData, name: K, value: Style[K]): void {
  r.style[name] = value;
  r.cells.forEach(cell => {
    if (cell) {
      cell.style[name] =
        typeof value === "object" && value !== null ? structuredClone(value) : value;
    }
  });
}

/**
 * Row namespace — free functions over the plain-data {@link RowData}.
 * Replaces the former `Row` class.
 */
export function rowCreate(worksheet: Worksheet, number: number): RowData {
  return { worksheet, number, cells: [], style: {}, outlineLevel: 0 };
}

export function rowDestroy(r: RowData): void {
  r.worksheet = undefined!;
  r.cells = undefined!;
  r.style = undefined!;
}

export function rowFindCell(r: RowData, colNumber: number): CellData | undefined {
  return r.cells[colNumber - 1];
}

export function rowAddPageBreak(r: RowData, lft?: number, rght?: number): void {
  const ws = r.worksheet;
  const left = Math.max(0, (lft ?? 0) - 1) || 0;
  const right = Math.max(0, (rght ?? 0) - 1) || 16838;
  const pb: RowBreak = { id: r.number, max: right, man: 1 };
  if (left) {
    pb.min = left;
  }
  ws.rowBreaks.push(pb);
}

export function rowValues(r: RowData): CellValue[] {
  const values: CellValue[] = [];
  r.cells.forEach(cell => {
    if (cell && cellType(cell) !== Enums.ValueType.Null) {
      values[cellCol(cell)] = cellGetValue(cell);
    }
  });
  return values;
}

export function rowGetValues(r: RowData): CellValue[] {
  const values: CellValue[] = [];
  r.cells.forEach(cell => {
    if (cell && cellType(cell) !== Enums.ValueType.Null) {
      values[cellCol(cell) - 1] = cellGetValue(cell);
    }
  });
  return values;
}

export function rowValuesToString(r: RowData, separator = ","): string {
  return rowGetValues(r).join(separator);
}

export function rowHasValues(r: RowData): boolean {
  return r.cells.some(cell => cell && cellType(cell) !== Enums.ValueType.Null);
}

export function rowCellCount(r: RowData): number {
  return r.cells.length;
}

export function rowActualCellCount(r: RowData): number {
  let count = 0;
  r.cells.forEach(cell => {
    if (cell && cellType(cell) !== Enums.ValueType.Null) {
      count++;
    }
  });
  return count;
}

export function rowDimensions(r: RowData): RowDimensions | null {
  let min = 0;
  let max = 0;
  r.cells.forEach(cell => {
    if (cell && cellType(cell) !== Enums.ValueType.Null) {
      if (!min || min > cellCol(cell)) {
        min = cellCol(cell);
      }
      if (max < cellCol(cell)) {
        max = cellCol(cell);
      }
    }
  });
  return min > 0 ? { min, max } : null;
}

export const rowNumFmt = (r: RowData): string | NumFmt | undefined => r.style.numFmt;

export function rowSetNumFmt(r: RowData, value: string | undefined): void {
  if (value !== undefined) {
    applyStyle(r, "numFmt", value);
  }
}

export const rowFont = (r: RowData): Partial<Font> | undefined => r.style.font;

export function rowSetFont(r: RowData, value: Partial<Font> | undefined): void {
  if (value !== undefined) {
    applyStyle(r, "font", value);
  }
}

export const rowAlignment = (r: RowData): Partial<Alignment> | undefined => r.style.alignment;

export function rowSetAlignment(r: RowData, value: Partial<Alignment> | undefined): void {
  if (value !== undefined) {
    applyStyle(r, "alignment", value);
  }
}

export const rowProtection = (r: RowData): Partial<Protection> | undefined => r.style.protection;

export function rowSetProtection(r: RowData, value: Partial<Protection> | undefined): void {
  if (value !== undefined) {
    applyStyle(r, "protection", value);
  }
}

export const rowBorder = (r: RowData): Partial<Borders> | undefined => r.style.border;

export function rowSetBorder(r: RowData, value: Partial<Borders> | undefined): void {
  if (value !== undefined) {
    applyStyle(r, "border", value);
  }
}

export const rowFill = (r: RowData): Fill | undefined => r.style.fill;

export function rowSetFill(r: RowData, value: Fill | undefined): void {
  if (value !== undefined) {
    applyStyle(r, "fill", value);
  }
}

/** Read the row's style record. */
export const rowGetStyle = (r: RowData): Partial<Style> => r.style;

/** Merge a partial style into the row (propagates to existing cells). */
export function rowSetStyle(r: RowData, style: Partial<Style>): void {
  (Object.keys(style) as (keyof Style)[]).forEach(k => {
    const v = style[k];
    if (v !== undefined) {
      applyStyle(r, k, v as Style[keyof Style]);
    }
  });
}

export const rowHidden = (r: RowData): boolean => !!r.hidden;

export function rowSetHidden(r: RowData, value: boolean): void {
  r.hidden = value;
}

export const rowOutlineLevel = (r: RowData): number => r.outlineLevel ?? 0;

export function rowSetOutlineLevel(r: RowData, value: number): void {
  r.outlineLevel = value;
}

export const rowCollapsed = (r: RowData): boolean =>
  !!(r.outlineLevel && r.outlineLevel >= (r.worksheet.properties.outlineLevelRow ?? 0));

export function rowGetModel(r: RowData): RowModel | null {
  const cells: CellModel[] = [];
  let min = 0;
  let max = 0;
  r.cells.forEach(cell => {
    if (cell) {
      const cellModel = cellGetModel(cell);
      if (cellModel) {
        if (!min || min > cellCol(cell)) {
          min = cellCol(cell);
        }
        if (max < cellCol(cell)) {
          max = cellCol(cell);
        }
        cells.push(cellModel);
      }
    }
  });

  return r.height != null || cells.length
    ? {
        cells,
        number: r.number,
        min,
        max,
        height: r.height,
        customHeight: r.customHeight,
        style: r.style,
        hidden: rowHidden(r),
        outlineLevel: rowOutlineLevel(r),
        collapsed: rowCollapsed(r),
        dyDescent: r.dyDescent
      }
    : null;
}

export function resolveColumnKeyValue(obj: Record<string, unknown>, key: string): unknown {
  const direct = obj[key];
  if (direct !== undefined || !key.includes(".")) {
    return direct;
  }
  let current: unknown = obj;
  for (const segment of key.split(".")) {
    if (current === null || typeof current !== "object") {
      return undefined;
    }
    current = (current as Record<string, unknown>)[segment];
  }
  return current;
}
