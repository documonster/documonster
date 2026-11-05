import type { CellData, CellView, NoteConfig } from "@excel/core/cell";
import {
  cellGetValue,
  cellSetAlignment,
  cellSetBorder,
  cellSetComment,
  cellSetFill,
  cellSetFont,
  cellSetNote,
  cellSetNumFmt,
  cellSetProtection,
  cellSetStyle,
  cellSetValue,
  cellView
} from "@excel/core/cell";
import type { ColumnData } from "@excel/core/column";
import { columnSetHidden, columnSetOutlineLevel } from "@excel/core/column";
/**
 * The streaming handle operations both `Stream` entries publish — Node and browser — so the two cannot drift apart.
 * Each setter checks what it is given, as the document model's do; the core setters they wrap are also what a reader
 * applies, and a file's own values are not the caller's to refuse.
 */
import { checkCellHandle, checkColumnHandle, checkRowHandle } from "@excel/core/feature-handles";
import type { NoteData } from "@excel/core/note";
import type { RowData } from "@excel/core/row";
import {
  rowAddPageBreak,
  rowGetValues as coreRowGetValues,
  rowValues as coreRowValues,
  rowSetAlignment,
  rowSetBorder,
  rowSetFill,
  rowSetFont,
  rowSetNumFmt,
  rowSetOutlineLevel,
  rowSetProtection
} from "@excel/core/row";
import { findCellStyleEntry } from "@excel/core/workbook-core";
import type { WorkbookData } from "@excel/core/workbook-core";
import {
  checkWorksheetAttached,
  rowGetCell as coreRowGetCell,
  rowCommit
} from "@excel/core/worksheet-core";
import { ExcelArgumentError, ExcelStreamStateError } from "@excel/errors";
import { checkComment, checkNote } from "@excel/surface/argument-checks";
import type {
  Alignment,
  Borders,
  CellValue,
  CellValueInput,
  Fill,
  Font,
  NumFmt,
  Protection
} from "@excel/types";
import { argumentChecks, checkNumber } from "@excel/utils/arguments";
import type { ValuesOptions } from "@excel/utils/numbered-values";
import { checkStyleFacet } from "@excel/utils/style-checks";

/**
 * Refuse an edit through the handle of a row the streaming writer has already written — by `commitRow`, or with its
 * sheet. The edit used to be accepted and to change nothing in the file: a note set after `commitRow` was lost.
 */
function checkRowOpen(operation: string, row: RowData): void {
  // A document-model row (the deprecated `Worksheet.getRow`) of a removed sheet: refused as the sheet refuses an edit.
  checkWorksheetAttached(row.worksheet as never);
  const writer = row.worksheet as unknown as { isRowCommitted?(rowNumber: number): boolean };
  if (writer.isRowCommitted?.(row.number)) {
    throw new ExcelStreamStateError(operation, `row ${row.number} has been committed`);
  }
}

// --- streaming cell handle operations (operate on a `CellData`) ---
/** A streamed cell's value. */
export function getCellValue(c: CellData): CellValue {
  checkCellHandle(c);
  return cellGetValue(c);
}

/**
 * Set a streamed cell's value — checked as `Cell.setValue` checks one: no `NaN`, `Infinity`, functions or invalid
 * dates.
 *
 * @throws {ExcelStreamStateError} when the cell's row has been committed.
 */
export function setCellValue(c: CellData, value: CellValueInput): void {
  checkCellHandle(c);
  checkRowOpen("set a cell value", c.row);
  cellSetValue(c, value);
}
/** A read-only view of a streamed cell: its address, value, type and formatting. */
export function getCellView(c: CellData): CellView {
  checkCellHandle(c);
  return cellView(c);
}

/**
 * Give a streamed cell the named cell style `name` — one `WorkbookWriter#addCellStyle` or `#useBuiltInCellStyle`
 * registered (on a document-model cell, one `Workbook.addCellStyle` defined), found in any case — as
 * `Cell.applyStyle` does. The cell records the style's own spelling.
 *
 * @throws {ExcelArgumentError} when no style of that name is registered — define it first.
 * @throws {ExcelStreamStateError} when the cell's row has been committed.
 */
export function applyCellStyle(c: CellData, name: string): void {
  checkCellHandle(c);
  checkRowOpen("apply a cell style", c.row);
  const check = argumentChecks("Stream.applyCellStyle");
  check.required("name", name);
  check.string("name", name, { nonEmpty: true });
  const workbook = (c.row.worksheet as { _workbook?: unknown })._workbook as
    | { _cellStyleName?(name: string): string | undefined }
    | WorkbookData;
  const stored =
    typeof (workbook as { _cellStyleName?: unknown })._cellStyleName === "function"
      ? (workbook as { _cellStyleName(name: string): string | undefined })._cellStyleName(name)
      : findCellStyleEntry(workbook as WorkbookData, name)?.name;
  if (stored === undefined) {
    throw new ExcelArgumentError(
      "name",
      `Stream.applyCellStyle: the named cell style "${name}" is not defined. Define it first with WorkbookWriter#addCellStyle.`
    );
  }
  cellSetStyle(c, { styleName: stored });
}

/** Set a streamed cell's font; `undefined` clears it. The font is checked as a cell style's is. */
export function setCellFont(c: CellData, value: Partial<Font> | undefined): void {
  checkCellHandle(c);
  checkRowOpen("set a cell font", c.row);
  checkStyleFacet("Stream.setCellFont", "font", value, "value");
  cellSetFont(c, value);
}
/** Set a streamed cell's fill; `undefined` clears it. The fill is checked as a cell style's is. */
export function setCellFill(c: CellData, value: Fill | undefined): void {
  checkCellHandle(c);
  checkRowOpen("set a cell fill", c.row);
  checkStyleFacet("Stream.setCellFill", "fill", value, "value");
  cellSetFill(c, value);
}
/** Set a streamed cell's borders; `undefined` clears them. They are checked as a cell style's are. */
export function setCellBorder(c: CellData, value: Partial<Borders> | undefined): void {
  checkCellHandle(c);
  checkRowOpen("set a cell border", c.row);
  checkStyleFacet("Stream.setCellBorder", "border", value, "value");
  cellSetBorder(c, value);
}
/** Set a streamed cell's alignment; `undefined` clears it. It is checked as a cell style's is. */
export function setCellAlignment(c: CellData, value: Partial<Alignment> | undefined): void {
  checkCellHandle(c);
  checkRowOpen("set a cell alignment", c.row);
  checkStyleFacet("Stream.setCellAlignment", "alignment", value, "value");
  cellSetAlignment(c, value);
}
/** Set a streamed cell's number format code; `undefined` clears it. */
export function setCellNumFmt(c: CellData, value: string | NumFmt | undefined): void {
  checkCellHandle(c);
  checkRowOpen("set a cell number format", c.row);
  checkStyleFacet("Stream.setCellNumFmt", "numFmt", value, "value");
  cellSetNumFmt(c, value);
}
/** Set a streamed cell's protection; `undefined` clears it. It is checked as a cell style's is. */
export function setCellProtection(c: CellData, value: Partial<Protection> | undefined): void {
  checkCellHandle(c);
  checkRowOpen("set a cell protection", c.row);
  checkStyleFacet("Stream.setCellProtection", "protection", value, "value");
  cellSetProtection(c, value);
}

/**
 * Set a streamed cell's note: text, or a note configuration — checked as `Cell.setNote` checks it. A note of the wrong
 * shape used to be accepted and failed when the row was committed, with a `TypeError` from inside the comment writer.
 */
export function setCellNote(c: CellData, note: string | NoteConfig): void {
  checkCellHandle(c);
  checkRowOpen("set a cell note", c.row);
  checkNote("Stream.setCellNote", "note", note);
  cellSetNote(c, note);
}

/** Set or clear (`undefined`) a streamed cell's comment, checked as `Cell.setComment` checks it. */
export function setCellComment(c: CellData, comment: NoteConfig | NoteData | undefined): void {
  checkCellHandle(c);
  checkRowOpen("set a cell comment", c.row);
  checkComment("Stream.setCellComment", "comment", comment);
  cellSetComment(c, comment);
}

// --- streaming row handle operations (operate on a `RowData`) ---

/**
 * A streamed row's values as a dense array with holes for empty cells — **0-based** (column A at index `0`) unless
 * `{ base: 1 }` asks for the index to be the column number, exactly as `Row.getValues` reads a document-model row.
 */
export function getRowValues(r: RowData, options?: ValuesOptions): CellValue[] {
  checkRowHandle(r);
  const check = argumentChecks("Stream.getRowValues");
  if (check.record("options", options) && options.base !== 0 && options.base !== 1) {
    check.fail("options.base", "0 or 1", options.base);
  }
  return options?.base === 1 ? coreRowValues(r) : coreRowGetValues(r);
}
/** Put a manual page break after the streamed row. @throws {ExcelStreamStateError} when the row has been committed. */
export function addRowPageBreak(r: RowData): void {
  checkRowHandle(r);
  checkRowOpen("add a page break", r);
  rowAddPageBreak(r);
}

/**
 * Set a streamed row's height in points — from 0 to 409, as `Row.setHeight` takes it; `undefined` clears it. Setting
 * `row.height` directly is unchecked.
 *
 * @throws {ExcelStreamStateError} when the row has been committed.
 */
export function setRowHeight(r: RowData, height: number | undefined): void {
  checkRowHandle(r);
  checkRowOpen("set a row height", r);
  if (height !== undefined) {
    checkNumber("Stream.setRowHeight", "height", height, { min: 0, max: 409 });
  }
  r.height = height;
  r.customHeight = height === undefined ? undefined : true;
}

/** Hide (`true`) or show (`false`) a streamed row. @throws {ExcelStreamStateError} when the row has been committed. */
export function setRowHidden(r: RowData, hidden: boolean): void {
  checkRowHandle(r);
  checkRowOpen("hide a row", r);
  const check = argumentChecks("Stream.setRowHidden");
  check.required("hidden", hidden);
  check.boolean("hidden", hidden);
  r.hidden = hidden;
}

/**
 * Refuse an edit to a streamed sheet's columns once they have been written — with the first committed row. The edit
 * used to be accepted and to change nothing in the file.
 */
function checkColumnsOpen(operation: string, column: ColumnData): void {
  checkWorksheetAttached(column.worksheet as never);
  const writer = column.worksheet as unknown as { areColumnsWritten?(): boolean };
  if (writer.areColumnsWritten?.()) {
    throw new ExcelStreamStateError(
      operation,
      "the sheet's columns were written with its first committed row"
    );
  }
}

/**
 * Set the width of a streamed sheet's column (from `WorksheetWriter#getColumn`), in characters — from 0 to 255, as
 * `Column.setWidth` takes it; `undefined` clears it. The columns are written with the first committed row.
 *
 * @throws {ExcelStreamStateError} when a row of the sheet has been committed.
 */
export function setColumnWidth(column: ColumnData, width: number | undefined): void {
  checkColumnHandle(column);
  if (width !== undefined) {
    checkNumber("Stream.setColumnWidth", "width", width, { min: 0, max: 255 });
  }
  checkColumnsOpen("set a column width", column);
  column.width = width;
  column.customWidth = undefined;
}
/**
 * Hide (`true`) or show (`false`) a streamed sheet's column.
 *
 * @throws {ExcelStreamStateError} when a row of the sheet has been committed.
 */
export function setColumnHidden(column: ColumnData, hidden: boolean): void {
  checkColumnHandle(column);
  const check = argumentChecks("Stream.setColumnHidden");
  check.required("hidden", hidden);
  check.boolean("hidden", hidden);
  checkColumnsOpen("hide a column", column);
  columnSetHidden(column, hidden);
}
/**
 * Set a streamed sheet's column outline (grouping) level, a whole number from 0 to 7.
 *
 * @throws {ExcelStreamStateError} when a row of the sheet has been committed.
 */
export function setColumnOutlineLevel(column: ColumnData, level: number): void {
  checkColumnHandle(column);
  checkNumber("Stream.setColumnOutlineLevel", "level", level, { min: 0, max: 7, integer: true });
  checkColumnsOpen("set a column outline level", column);
  columnSetOutlineLevel(column, level);
}
/**
 * Set a streamed row's outline (grouping) level, a whole number from 0 to 7, as `Row.setOutlineLevel` takes it.
 *
 * @throws {ExcelStreamStateError} when the row has been committed.
 */
export function setRowOutlineLevel(r: RowData, level: number): void {
  checkRowHandle(r);
  checkRowOpen("set a row outline level", r);
  checkNumber("Stream.setRowOutlineLevel", "level", level, { min: 0, max: 7, integer: true });
  rowSetOutlineLevel(r, level);
}
/** Set the number format of a streamed row; `undefined` clears it. It is checked as a cell style's is. */
export function setRowNumFmt(r: RowData, value: string | NumFmt | undefined): void {
  checkRowHandle(r);
  checkRowOpen("set a row number format", r);
  checkStyleFacet("Stream.setRowNumFmt", "numFmt", value, "value");
  rowSetNumFmt(r, value);
}
/** Set the protection of a streamed row; `undefined` clears it. It is checked as a cell style's is. */
export function setRowProtection(r: RowData, value: Partial<Protection> | undefined): void {
  checkRowHandle(r);
  checkRowOpen("set a row protection", r);
  checkStyleFacet("Stream.setRowProtection", "protection", value, "value");
  rowSetProtection(r, value);
}
/** Set the font of a streamed row; `undefined` clears it. The font is checked as a cell style's is. */
export function setRowFont(r: RowData, value: Partial<Font> | undefined): void {
  checkRowHandle(r);
  checkRowOpen("set a row font", r);
  checkStyleFacet("Stream.setRowFont", "font", value, "value");
  rowSetFont(r, value);
}
/** Set the fill of a streamed row; `undefined` clears it. The fill is checked as a cell style's is. */
export function setRowFill(r: RowData, value: Fill | undefined): void {
  checkRowHandle(r);
  checkRowOpen("set a row fill", r);
  checkStyleFacet("Stream.setRowFill", "fill", value, "value");
  rowSetFill(r, value);
}
/** Set the borders of a streamed row; `undefined` clears them. They are checked as a cell style's are. */
export function setRowBorder(r: RowData, value: Partial<Borders> | undefined): void {
  checkRowHandle(r);
  checkRowOpen("set a row border", r);
  checkStyleFacet("Stream.setRowBorder", "border", value, "value");
  rowSetBorder(r, value);
}
/** Set the alignment of a streamed row; `undefined` clears it. It is checked as a cell style's is. */
export function setRowAlignment(r: RowData, value: Partial<Alignment> | undefined): void {
  checkRowHandle(r);
  checkRowOpen("set a row alignment", r);
  checkStyleFacet("Stream.setRowAlignment", "alignment", value, "value");
  rowSetAlignment(r, value);
}
/**
 * The cell of a streamed row at a column (`3` or `"C"`), **created when the row has none there**.
 *
 * This hands out the handle every `setCell*` writes through, so — like `getStyle` on the document model — creating
 * what it names is its purpose: a streamed row is output being composed, not a record being read. It is the same
 * get-or-create the writer's own `getCell` does; `WorksheetWriter#findCell` looks a cell up without creating it.
 */
export function getRowCell(r: RowData, col: string | number): CellData {
  checkRowHandle(r);
  return coreRowGetCell(r, col);
}
/** Commit a streamed row: it and every row before it are written, and take no more edits. */
export function commitRow(r: RowData): void {
  checkRowHandle(r);
  rowCommit(r);
}

// --- streaming types ---
/** A streaming cell handle. */
export type { CellData as CellHandle } from "@excel/core/cell";
/** A streaming row handle. */
export type { RowData as RowHandle } from "@excel/core/row";
/** A streaming column handle — what `WorksheetWriter#getColumn` returns and the `setColumn*` operations take. */
export type { ColumnData as ColumnHandle } from "@excel/core/column";

/**
 * @deprecated Use {@link getRowValues} with `{ base: 1 }` — this reads the row 1-based, by column number. Removed in
 * 1.0.
 */
export const rowValues = (r: RowData): CellValue[] => getRowValues(r, { base: 1 });
/** @deprecated Use {@link getRowCell}. Removed in 1.0. */
export const rowCell = getRowCell;
