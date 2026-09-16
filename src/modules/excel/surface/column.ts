/**
 * `Column` namespace surface — column-level operations addressed by key, letter,
 * or 1-based number.
 *
 * `import { Column } from "documonster/excel"` → `Column.setWidth(ws, "A", 20)`,
 * `Column.setHeader(ws, 1, "Name")`, `Column.setStyle(ws, "B", { numFmt })`,
 * `Column.values(ws, "total")`.
 *
 * The `col` parameter is a column key, an `"A"` letter, or a 1-based column
 * number — hence the `string | number` unions in the signatures below.
 */
import type { ColumnDefn, ColumnHeaderValue } from "@excel/core/column";
import {
  columnAddPageBreak,
  columnDefn,
  columnHidden,
  columnLetter,
  columnOutlineLevel,
  columnSetHidden,
  columnSetOutlineLevel
} from "@excel/core/column";
import { invalidateSharedCellStyle } from "@excel/core/style-sharing";
import {
  columnSetHeader,
  columnSetKey,
  columnSetNumFmt,
  columnSetStyle,
  columnSetValuesAt,
  columnValuesAt,
  findColumnNumber,
  getColumn
} from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import type { CellValue, CellValueInput, Style } from "@excel/types";

export type Sheet = WorksheetData;

// --- identity ---

/**
 * The 1-based column number for a key, letter or number.
 *
 * This is the bridge from a column *key* to the `(row, col)` form the `Cell`
 * namespace takes: `Cell.setValue(ws, row, Column.getNumber(ws, "total"), v)`.
 */
export function getNumber(ws: Sheet, col: string | number): number {
  return getColumn(ws, col).number;
}

/** The column letter (`"A"`, `"AB"`) for a key, letter or number. */
export function getLetter(ws: Sheet, col: string | number): string {
  return columnLetter(getColumn(ws, col));
}

// --- width ---

export function getWidth(ws: Sheet, col: string | number): number | undefined {
  return getColumn(ws, col).width;
}
export function setWidth(ws: Sheet, col: string | number, width: number): void {
  getColumn(ws, col).width = width;
}

// --- header / key ---

export function getHeader(ws: Sheet, col: string | number): ColumnHeaderValue | undefined {
  return getColumn(ws, col).header;
}
export function setHeader(ws: Sheet, col: string | number, header: ColumnHeaderValue): void {
  columnSetHeader(getColumn(ws, col), header);
}
export function getKey(ws: Sheet, col: string | number): string | undefined {
  return getColumn(ws, col).key;
}
export function setKey(ws: Sheet, col: string | number, key: string): void {
  columnSetKey(getColumn(ws, col), key);
}

// --- visibility / outline ---

export function getHidden(ws: Sheet, col: string | number): boolean {
  return columnHidden(getColumn(ws, col));
}
export function setHidden(ws: Sheet, col: string | number, hidden: boolean): void {
  columnSetHidden(getColumn(ws, col), hidden);
}
export function getOutlineLevel(ws: Sheet, col: string | number): number {
  return columnOutlineLevel(getColumn(ws, col));
}
export function setOutlineLevel(ws: Sheet, col: string | number, level: number): void {
  columnSetOutlineLevel(getColumn(ws, col), level);
}

// --- style ---

/**
 * The column's own style object, live — mutating it is how a caller changes the column.
 *
 * Cells the column has already styled point at a *snapshot* of these facets, so that
 * mutation does not reach them (as it does not today). The snapshot is dropped here
 * because the caller may mutate in place, and a cell created afterwards must copy
 * what the column says then — see `core/style-sharing.ts`.
 */
export function getStyle(ws: Sheet, col: string | number): Partial<Style> {
  const { style } = getColumn(ws, col);
  invalidateSharedCellStyle(style);
  return style;
}
export function setStyle(ws: Sheet, col: string | number, style: Partial<Style>): void {
  columnSetStyle(getColumn(ws, col), style);
}

// --- number format ---

/**
 * Set the column's number format.
 *
 * The mirror of `Cell.setNumFmt` / `Row.setNumFmt`, so a format can be set at
 * whichever level owns it. It is the one facet with its own setter here because
 * it is the one consumers reach for by column; every other facet goes through
 * {@link setStyle} (`setStyle(ws, col, { font })`), which also walks the column a
 * single time when several facets are set together.
 *
 * The format lands on the column *and* on the cells of every row that holds a
 * value, which is not quite "every existing cell": a row with values elsewhere
 * gains a cell in this column if it had none, and a materialised-but-empty cell in
 * a row with no values at all is skipped. That is the behaviour of every
 * `columnSet*`, not a property of this one.
 *
 * Passing `undefined` clears the format.
 *
 * There is no matching *getter*: a column's style is one record and
 * {@link getStyle} hands it over whole, so `getStyle(ws, col).numFmt` is the read.
 *
 * @example
 * ```typescript
 * import { Column, Workbook } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Sales");
 * Column.setNumFmt(ws, "C", "#,##0.00");
 * ```
 */
export function setNumFmt(ws: Sheet, col: string | number, value: string | undefined): void {
  columnSetNumFmt(getColumn(ws, col), value);
}

// --- definition ---

/**
 * The column's definition as plain data — the shape `Worksheet.setColumns`
 * takes, with `hidden` / `outlineLevel` normalised the way the library
 * normalises them internally.
 *
 * Use this (or `Worksheet.columnDefinitions` for every column) instead of copying
 * fields off a column handle by hand: hand-copying re-implements the
 * normalisation and silently drifts the moment a field or a fallback changes.
 */
export function getDefinition(ws: Sheet, col: string | number): ColumnDefn {
  return columnDefn(getColumn(ws, col));
}

// --- values ---

/**
 * Read the values down a column as a **0-based sparse** array: the value in row
 * 1 is at index `0`, row 2 at index `1`, and so on, and an empty cell is a hole
 * rather than a `null`.
 *
 * Contrast with {@link values}, which is indexed by row *number* and therefore
 * composes directly with the `Cell` namespace. That is also the one to pass back
 * to {@link setValues} — see the ambiguity noted there.
 *
 * The read is non-destructive: it walks the rows that exist rather than asking
 * each one for a cell, so it neither materialises cells down the column nor
 * declares the column itself. The metadata and style accessors in this namespace
 * resolve their reference through a column *record*, which pads
 * `Worksheet.columns` up to the column asked for.
 *
 * @example
 * ```typescript
 * import { Column, Workbook, Worksheet } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Sales");
 * Worksheet.setColumns(ws, [{ key: "total", header: "Total" }]);
 * Column.setValues(ws, "total", [10, 20, 30]);
 * const [first] = Column.getValues(ws, "total"); // 10 — row 1
 * ```
 */
export function getValues(ws: Sheet, col: string | number): CellValue[] {
  return columnValuesAt(ws, findColumnNumber(ws, col), 0);
}

/**
 * Read the values down a column as a **sparse** array indexed by **row number**:
 * index `0` is always an empty leading slot, the value in row 1 is at index `1`,
 * and so on. This mirrors Excel's 1-based row numbering, so the index can be
 * handed straight to the `Cell` namespace, and it is the form {@link setValues}
 * reads back unambiguously.
 *
 * Use {@link getValues} when the values are wanted as a plain 0-based array. Like
 * it, this read materialises nothing.
 *
 * @example
 * ```typescript
 * import { Cell, Column, Workbook } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Sales");
 * Column.setValues(ws, "C", [120, -40, 75]);
 *
 * // Colour the losses red. The index *is* the row number, and `Column.getNumber`
 * // turns the reference into the column number `Cell` takes.
 * const colNumber = Column.getNumber(ws, "C");
 * Column.values(ws, "C").forEach((value, rowNumber) => {
 *   if (typeof value === "number" && value < 0) {
 *     Cell.setFont(ws, rowNumber, colNumber, { color: { argb: "FFC00000" } });
 *   }
 * });
 * ```
 */
export function values(ws: Sheet, col: string | number): CellValue[] {
  return columnValuesAt(ws, findColumnNumber(ws, col));
}

/**
 * Write values down a column. Takes anything `Cell.setValue` takes — a number, a
 * string, a formula, rich text, a hyperlink, a `Date` or a civil `Temporal`.
 *
 * The array is read as **1-based by row number** when index 0 is a hole — the
 * shape {@link values} hands back, so that read round-trips — and as **0-based**
 * when index 0 is present, so a plain array starts at row 1:
 *
 * ```typescript
 * Column.setValues(ws, "A", [1, 2, 3]); // rows 1, 2, 3
 * Column.setValues(ws, "A", [, 1, 2, 3]); // rows 1, 2, 3 — the same thing
 * ```
 *
 * The two are indistinguishable when a 0-based array's own first slot is a hole,
 * so pass {@link values} rather than {@link getValues} when writing a read back —
 * a `getValues` array whose row 1 was empty lands a row high. `Row.setValues` has
 * the same ambiguity.
 *
 * **This writes, it does not replace.** Only the indices the array carries are
 * touched — unlike `Row.setValues`, which resets the row — so a shorter array
 * does not truncate the column and a value already in a row the array skips
 * survives. Clear a cell by passing `null` at its index.
 */
export function setValues(ws: Sheet, col: string | number, v: readonly CellValueInput[]): void {
  columnSetValuesAt(ws, findColumnNumber(ws, col), v);
}

// --- printing ---

/**
 * Add a manual vertical page break to the **right** of `col`, the equivalent of
 * Excel's *Page Layout → Breaks → Insert Page Break*. Applies to printing and
 * to `Pdf.fromExcel`; it does not affect the on-screen grid.
 *
 * The break spans the full height of the sheet, which is the only kind Excel can
 * author or render — see `columnAddPageBreak` for why no row band is offered.
 *
 * @example
 * ```typescript
 * import { Column, Workbook } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Report");
 * Column.addPageBreak(ws, "F"); // the next page starts at column G
 * ```
 */
export function addPageBreak(ws: Sheet, col: string | number): void {
  columnAddPageBreak(getColumn(ws, col));
}
