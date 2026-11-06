/**
 * `Row` namespace surface — row-level operations addressed by row number.
 *
 * `import { Row } from "documonster/excel"` → `Row.add(ws, [1, 2, 3])`,
 * `Row.setHeight(ws, 1, 24)`, `Row.setStyle(ws, 3, { font })`, `Row.getValues(ws, 2)`,
 * `Row.each(ws, rowNumber => …)`.
 *
 * Rows are addressed by their 1-based number and nothing here hands out a row
 * handle: a function that creates rows answers with their numbers. Handles belong
 * to the streaming API (`Stream.RowHandle`).
 */
import type { CellData } from "@excel/core/cell";
import { cellType, checkCellValue, checkRowList, checkRowValues } from "@excel/core/cell";
import { ValueType } from "@excel/core/enums";
import type { RowData, RowDimensions, RowModel } from "@excel/core/row";
import {
  rowActualCellCount,
  rowAddPageBreak,
  rowCellCount,
  rowCollapsed,
  rowDimensions,
  rowGetModel,
  rowHasValues,
  rowGetStyle,
  rowGetValues,
  rowHidden,
  rowOutlineLevel,
  rowSetAlignment,
  rowSetBorder,
  rowSetFill,
  rowSetFont,
  rowSetHidden,
  rowSetNumFmt,
  rowSetOutlineLevel,
  rowSetProtection,
  rowSetStyle,
  rowValues
} from "@excel/core/row";
import { invalidateSharedCellStyle } from "@excel/core/style-sharing";
import {
  autoFitRow,
  autoFitRows,
  duplicateRow,
  getActualRowCount,
  getRowCount,
  insertRow,
  insertRows,
  spliceRows
} from "@excel/core/worksheet";
import {
  addRow,
  checkWorksheetAttached,
  eachRow,
  findRow,
  getRow,
  rowCommit,
  rowEachCell,
  rowGetCell,
  rowSetValues,
  rowSplice
} from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { ExcelUnsupportedFeatureError } from "@excel/errors";
import type { StreamWorksheet } from "@excel/surface/stream-targets";
import {
  asStreamWorksheet,
  checkStreamRowsHeld,
  checkStreamRowsOpen,
  refuseWrittenRows
} from "@excel/surface/stream-targets";
import type {
  Alignment,
  Borders,
  CellValue,
  CellValueInput,
  Fill,
  Font,
  NumFmt,
  Protection,
  RowValues,
  Style
} from "@excel/types";
import { argumentChecks, checkNumber } from "@excel/utils/arguments";
import { colCache } from "@excel/utils/col-cache";
import { copyPlain } from "@excel/utils/copy-plain";
import { copyStyle, copyStyleFacet } from "@excel/utils/copy-style";
import { numberedValues } from "@excel/utils/numbered-values";
import type { ValuesBase, ValuesOptions } from "@excel/utils/numbered-values";
import { checkStyle, checkStyleFacet } from "@excel/utils/style-checks";

/** @deprecated Use `Worksheet.Handle`. Removed in 1.0. */
export type Sheet = WorksheetData;

const MAX_ROW = 1048576;
const MAX_COL = 16384;

/**
 * Where a new row takes its style from.
 *
 * `"above"` copies the row above it, `"below"` the row below it (which only
 * exists for an insert). Without it the new row is unstyled.
 */
export type RowStyleSource = "above" | "below";

/** Options for {@link add} and {@link addMany}. */
export interface AddOptions {
  /** Copy the style of the row above (the previous last row). */
  copyStyle?: "above";
  /** With `copyStyle`, also copy the style of the source row's empty cells. */
  includeEmptyCells?: boolean;
}

/** Options for {@link insert} and {@link insertMany}. */
export interface InsertOptions {
  /** Copy the style of the row above or below the inserted block. */
  copyStyle?: RowStyleSource;
  /** With `copyStyle`, also copy the style of the source row's empty cells. */
  includeEmptyCells?: boolean;
}

/** Options for {@link duplicate}. */
export interface DuplicateOptions {
  /**
   * Insert the copies below the row, pushing what follows down. Without it the
   * copies overwrite the rows below.
   */
  insert?: boolean;
}

/** Options for {@link each} and {@link eachAddress}. */
export interface EachOptions {
  /** Also visit rows (or cells) that hold no value. */
  includeEmpty?: boolean;
}

/** Options for the range form of {@link autoFit}. */
export interface AutoFitOptions {
  /** First row to fit. Defaults to the top of the sheet's used range. */
  start?: number;
  /** Last row to fit. Defaults to the bottom of the sheet's used range. */
  end?: number;
}

/** The legacy one-letter style code `addRow` / `insertRow` take. */
function styleCode(
  where: string,
  options: AddOptions | InsertOptions | undefined,
  allowed: readonly RowStyleSource[]
): string {
  if (options !== undefined && options !== null) {
    argumentChecks(where).oneOf("options.copyStyle", options.copyStyle, allowed);
    argumentChecks(where).boolean("options.includeEmptyCells", options.includeEmptyCells);
  }
  if (!options?.copyStyle) {
    return "n";
  }
  return (options.copyStyle === "above" ? "i" : "o") + (options.includeEmptyCells ? "+" : "");
}

// --- adding / inserting / removing rows ---

/**
 * Row `row` for an edit, created when absent (`create`) — refused once a streaming writer has written it, as
 * `Stream.setRow*` refuses an edit through a committed row's handle. The row number is checked first.
 */
function editRow(ws: WorksheetData | StreamWorksheet, where: string, row: number): RowData;
function editRow(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  row: number,
  create: false
): RowData | undefined;
function editRow(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  row: number,
  create = true
): RowData | undefined {
  checkStreamRowsOpen(ws, where, row);
  const sheet = ws as WorksheetData;
  return create ? getRow(sheet, row) : findRow(sheet, row);
}

/** Row `row` for a read, never created — refused once a streaming writer has written it, as it no longer holds it. */
function heldRow(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  row: number
): RowData | undefined {
  const r = findRow(ws as WorksheetData, row);
  checkStreamRowsHeld(ws, where, row);
  return r;
}

/**
 * A streaming writer adds a row as given: copying the style of the row above would need that row, which it may already
 * have written.
 */
function refuseStyleCopy(code: string, where: string): void {
  if (code !== "n") {
    throw new ExcelUnsupportedFeatureError(
      where,
      "a streaming worksheet writer cannot copy the style of the row above: it may already have written it"
    );
  }
}

/**
 * Append a row after the last one and return its number.
 *
 * @example
 * ```typescript
 * import { Row, Workbook } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Data");
 * const header = Row.add(ws, ["Name", "Total"]); // 1
 * Row.add(ws, ["Ada", 42], { copyStyle: "above" });
 * ```
 */
export function add(
  ws: WorksheetData | StreamWorksheet,
  values: RowValues | readonly CellValueInput[],
  options?: AddOptions
): number {
  checkRowValues("Row.add", values);
  const code = styleCode("Row.add", options, ["above"]);
  const writer = asStreamWorksheet(ws);
  if (writer) {
    refuseStyleCopy(code, "Row.add");
    return writer._appendCheckedRows([values as RowValues])[0].number;
  }
  return addRow(ws as WorksheetData, values as RowValues, code).number;
}

/** Append several rows and return their numbers. */
export function addMany(
  ws: WorksheetData | StreamWorksheet,
  values: readonly (RowValues | readonly CellValueInput[])[],
  options?: AddOptions
): number[] {
  const code = styleCode("Row.addMany", options, ["above"]);
  checkRowList("Row.addMany", values);
  const writer = asStreamWorksheet(ws);
  if (writer) {
    refuseStyleCopy(code, "Row.addMany");
    return writer._appendCheckedRows(values as RowValues[]).map(r => r.number);
  }
  return values.map(v => addRow(ws as WorksheetData, v as RowValues, code).number);
}

/**
 * Insert a row at `row`, pushing that row and everything below it down. Returns `row`.
 *
 * Everything that refers to cells follows them, on every sheet, as it does in Excel: formulas, defined names,
 * tables, conditional formats, data validations, charts, sparklines, pivot sources, the auto filter, print areas
 * and drawings. A reference to cells the edit deletes becomes `#REF!`.
 */
export function insert(
  ws: WorksheetData,
  row: number,
  values: RowValues | readonly CellValueInput[],
  options?: InsertOptions
): number {
  refuseWrittenRows(ws, "Row.insert", "insert a row");
  checkRowValues("Row.insert", values);
  insertRow(ws, row, values as RowValues, styleCode("Row.insert", options, ["above", "below"]));
  return row;
}

/** Insert several rows starting at `row`. Returns the numbers of the inserted rows. */
export function insertMany(
  ws: WorksheetData,
  row: number,
  values: readonly (RowValues | readonly CellValueInput[])[],
  options?: InsertOptions
): number[] {
  refuseWrittenRows(ws, "Row.insertMany", "insert rows");
  checkRowList("Row.insertMany", values);
  insertRows(
    ws,
    row,
    [...values] as RowValues[],
    styleCode("Row.insertMany", options, ["above", "below"])
  );
  return values.map((_, i) => row + i);
}

/**
 * Copy row `row` into the `count` rows below it — overwriting them, or with
 * `{ insert: true }` inserting the copies and pushing the rest down.
 */
export function duplicate(
  ws: WorksheetData,
  row: number,
  count: number,
  options?: DuplicateOptions
): void {
  refuseWrittenRows(ws, "Row.duplicate", "duplicate a row");
  duplicateRow(ws, row, count, options?.insert ?? false);
}

/**
 * Remove `count` rows starting at `start` and insert `inserts` in their place,
 * as `Array.prototype.splice` does. *
 * Everything that refers to cells follows them, on every sheet, as it does in Excel: formulas, defined names,
 * tables, conditional formats, data validations, charts, sparklines, pivot sources, the auto filter, print areas
 * and drawings. A reference to cells the edit deletes becomes `#REF!`.
 */
export function splice(
  ws: WorksheetData,
  start: number,
  count: number,
  ...inserts: (RowValues | readonly CellValueInput[])[]
): void {
  refuseWrittenRows(ws, "Row.splice", "splice rows");
  inserts.forEach((values, i) => checkRowValues("Row.splice", values, `inserts[${i}]`));
  spliceRows(ws, start, count, ...(inserts as RowValues[]));
}

/**
 * Within row `row`, remove `count` cells starting at column `start` (1-based) and insert `inserts` in their place,
 * as `Array.prototype.splice` does: the cells to the right move left or right, each taking its value, style and
 * note with it, and an inserted cell starts unstyled. Other rows are not touched, and nothing that refers to the
 * moved cells (a formula, a defined name, a merge) is rewritten — for an edit references follow, use
 * `Column.splice`.
 *
 * This is exceljs's `row.splice(start, count, ...values)`.
 */
export function spliceCells(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  start: number,
  count: number,
  ...inserts: CellValueInput[]
): void {
  const where = "Row.spliceCells";
  checkNumber(where, "start", start, { min: 1, max: MAX_COL, integer: true });
  checkNumber(where, "count", count, { min: 0, integer: true });
  inserts.forEach((value, i) => checkCellValue(value, { where, field: `inserts[${i}]` }));
  rowSplice(editRow(ws, where, row), start, count, ...(inserts as CellValue[]));
}

// --- counting / iteration ---

/** The number of the last row that exists — the sheet's row count. */
export function getCount(ws: WorksheetData): number {
  refuseWrittenRows(ws, "Row.getCount", "count its rows");
  return getRowCount(ws);
}

/** How many rows actually hold a value. */
export function getActualCount(ws: WorksheetData): number {
  refuseWrittenRows(ws, "Row.getActualCount", "count its rows");
  return getActualRowCount(ws);
}

/** The number of the last row that exists — {@link getCount} — or `undefined` for an empty sheet. */
export function getLast(ws: WorksheetData): number | undefined {
  refuseWrittenRows(ws, "Row.getLast", "find its last row");
  const n = getRowCount(ws);
  return n > 0 ? n : undefined;
}

/** Whether row `row` exists. Never creates it. */
export function has(ws: WorksheetData | StreamWorksheet, row: number): boolean {
  return heldRow(ws, "Row.has", row) !== undefined;
}

/**
 * Visit each row that holds a value, in order, with its number. With
 * `{ includeEmpty: true }` every row from 1 to the last is visited — without
 * creating the ones that do not exist.
 */
export function each(
  ws: WorksheetData,
  callback: (rowNumber: number) => void,
  options?: EachOptions
): void {
  refuseWrittenRows(ws, "Row.each", "walk its rows");
  if (typeof callback !== "function") {
    argumentChecks("Row.each").fail("callback", "a function", callback);
  }
  if (options?.includeEmpty) {
    const n = getRowCount(ws);
    for (let r = 1; r <= n; r++) {
      callback(r);
    }
    return;
  }
  eachRow(ws, (_row, rowNumber) => callback(rowNumber));
}

/**
 * Visit each cell of row `row` that holds a value, in column order, as
 * `(address, row, col)` — the shape `Range.eachAddress` hands over, so the
 * arguments go straight to the `Cell` namespace. With `{ includeEmpty: true }`
 * every column up to the row's last cell is visited, without creating cells.
 */
export function eachAddress(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  callback: (address: string, row: number, col: number) => void,
  options?: EachOptions
): void {
  if (typeof callback !== "function") {
    argumentChecks("Row.eachAddress").fail("callback", "a function", callback);
  }
  const r = heldRow(ws, "Row.eachAddress", row);
  if (!r) {
    return;
  }
  const n = r.cells.length;
  for (let col = 1; col <= n; col++) {
    const cell = r.cells[col - 1];
    if (options?.includeEmpty || (cell && cellType(cell) !== ValueType.Null)) {
      callback(colCache.n2l(col) + row, row, col);
    }
  }
}

// --- per-row facts (read without creating the row) ---

/**
 * Whether row `row` holds at least one value. A row that does not exist holds
 * none; asking never creates it.
 */
export function hasValues(ws: WorksheetData | StreamWorksheet, row: number): boolean {
  const r = heldRow(ws, "Row.hasValues", row);
  return r !== undefined && rowHasValues(r);
}

/**
 * How many cell slots row `row` has — the column number of its last cell,
 * empty or not. `0` for a row that does not exist.
 */
export function getCellCount(ws: WorksheetData | StreamWorksheet, row: number): number {
  const r = heldRow(ws, "Row.getCellCount", row);
  return r ? rowCellCount(r) : 0;
}

/** How many cells of row `row` actually hold a value. `0` for a row that does not exist. */
export function getActualCellCount(ws: WorksheetData | StreamWorksheet, row: number): number {
  const r = heldRow(ws, "Row.getActualCellCount", row);
  return r ? rowActualCellCount(r) : 0;
}

/**
 * The first and last column of row `row` that hold a value, or `undefined` when
 * it holds none (or does not exist).
 */
export function getDimensions(
  ws: WorksheetData | StreamWorksheet,
  row: number
): RowDimensions | undefined {
  const r = heldRow(ws, "Row.getDimensions", row);
  return (r && rowDimensions(r)) ?? undefined;
}

/**
 * Whether row `row` is collapsed — its outline level is at or above the sheet's
 * row outline level. `false` for a row that does not exist.
 */
export function isCollapsed(ws: WorksheetData | StreamWorksheet, row: number): boolean {
  const r = heldRow(ws, "Row.isCollapsed", row);
  return r !== undefined && rowCollapsed(r);
}

/**
 * Row `row` as a plain serialisable snapshot — cells, span, height, style,
 * visibility — or `undefined` when the row does not exist or has nothing to
 * write. It is a copy: editing it changes nothing.
 */
export function getModel(ws: WorksheetData | StreamWorksheet, row: number): RowModel | undefined {
  const r = heldRow(ws, "Row.getModel", row);
  const model = r && rowGetModel(r);
  return model ? copyPlain(model) : undefined;
}

// --- height ---

/** The row's height in points, or `undefined` when it has none of its own. Never creates the row. */
export function getHeight(ws: WorksheetData | StreamWorksheet, row: number): number | undefined {
  return heldRow(ws, "Row.getHeight", row)?.height;
}
/**
 * Set the row's height in points — from 0 to 409, Excel's tallest row ("Excel specifications and limits"). `undefined`
 * clears it, so the row takes the sheet's default height again, as {@link getHeight} reports a row with none.
 *
 * @throws {ExcelArgumentError} when `height` is not a number in that range.
 */
export function setHeight(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  height: number | undefined
): void {
  if (height === undefined) {
    // Clearing a row that does not exist has nothing to clear; do not create it to find that out.
    const r = editRow(ws, "Row.setHeight", row, false);
    checkWorksheetAttached(ws as WorksheetData);
    if (r) {
      r.height = undefined;
      r.customHeight = undefined;
    }
    return;
  }
  checkNumber("Row.setHeight", "height", height, { min: 0, max: 409 });
  const r = editRow(ws, "Row.setHeight", row);
  r.height = height;
  r.customHeight = true;
}

/**
 * Fit row heights to their content, as Excel's AutoFit does: one row by number,
 * or a span (`{ start, end }`, defaulting to the used range) when given options.
 */
export function autoFit(ws: WorksheetData, row: number): void;
export function autoFit(ws: WorksheetData, options?: AutoFitOptions): void;
export function autoFit(ws: WorksheetData, rowOrOptions?: number | AutoFitOptions): void {
  refuseWrittenRows(ws, "Row.autoFit", "fit rows to their cells");
  const check = argumentChecks("Row.autoFit");
  if (typeof rowOrOptions === "number") {
    check.number("row", rowOrOptions, { min: 1, max: MAX_ROW, integer: true });
    autoFitRow(ws, rowOrOptions);
  } else {
    const options: AutoFitOptions | undefined = rowOrOptions;
    if (check.record("options", rowOrOptions)) {
      check.number("options.start", rowOrOptions.start, { min: 1, max: MAX_ROW, integer: true });
      check.number("options.end", rowOrOptions.end, { min: 1, max: MAX_ROW, integer: true });
    }
    autoFitRows(ws, options?.start, options?.end);
  }
}

// --- visibility / outline ---

/** Whether row `row` is hidden. Never creates the row. */
export function isHidden(ws: WorksheetData | StreamWorksheet, row: number): boolean {
  const r = heldRow(ws, "Row.isHidden", row);
  return r !== undefined && rowHidden(r);
}
/** Hide or show row `row`. */
export function setHidden(ws: WorksheetData | StreamWorksheet, row: number, hidden: boolean): void {
  argumentChecks("Row.setHidden").required("hidden", hidden);
  argumentChecks("Row.setHidden").boolean("hidden", hidden);
  rowSetHidden(editRow(ws, "Row.setHidden", row), hidden);
}
/** Row `row`'s outline level, `0` when it has none. Never creates the row. */
export function getOutlineLevel(ws: WorksheetData | StreamWorksheet, row: number): number {
  const r = heldRow(ws, "Row.getOutlineLevel", row);
  return r ? rowOutlineLevel(r) : 0;
}
/** Set row `row`'s outline (grouping) level, a whole number from 0 to 7. */
export function setOutlineLevel(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  level: number
): void {
  checkNumber("Row.setOutlineLevel", "level", level, { min: 0, max: 7, integer: true });
  rowSetOutlineLevel(editRow(ws, "Row.setOutlineLevel", row), level);
}

// --- style ---

/**
 * The row's own style object, live — mutating it is how a caller changes the row.
 * Because it is handed out for editing, this is the one reader here that creates
 * the row when it does not exist yet. To *read* the style, use {@link findStyle};
 * to read one facet, {@link getFont}, {@link getFill}, … — they create nothing.
 *
 * Cells the row has already styled point at a *snapshot* of these facets, so that
 * mutation does not reach them (as it does not today). The snapshot is dropped here
 * because the caller may mutate in place, and a cell created afterwards must copy
 * what the row says then — see `core/style-sharing.ts`.
 */
export function getStyle(ws: WorksheetData | StreamWorksheet, row: number): Partial<Style> {
  const style = rowGetStyle(editRow(ws, "Row.getStyle", row));
  invalidateSharedCellStyle(style);
  return style;
}
/** Merge `style` into row `row`'s style and the style of every cell it already has. */
export function setStyle(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  style: Partial<Style>
): void {
  checkStyle("Row.setStyle", style);
  rowSetStyle(editRow(ws, "Row.setStyle", row), style);
}

/**
 * A copy of the row's own style, or `undefined` when the sheet has no such row — the
 * read-only counterpart of {@link getStyle}: it creates nothing, and editing what it
 * returns changes nothing.
 */
export function findStyle(
  ws: WorksheetData | StreamWorksheet,
  row: number
): Partial<Style> | undefined {
  const r = heldRow(ws, "Row.findStyle", row);
  return r ? (copyStyle(r.style) as Partial<Style>) : undefined;
}

// --- values ---

/** How a values array is indexed — see {@link ValuesOptions}. */
export type { ValuesBase, ValuesOptions };

/**
 * Read the row's cell values as a dense array with holes for empty cells — **0-based**
 * (column A at index `0`) unless `{ base: 1 }` asks for the index to be the column
 * number. Never creates the row.
 *
 * @example
 * ```typescript
 * Row.getValues(ws, 2); // ["Ada", 42]
 * Row.getValues(ws, 2, { base: 1 }); // [, "Ada", 42]
 * ```
 */
export function getValues(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  options?: ValuesOptions
): CellValue[] {
  checkBase("Row.getValues", options);
  const r = heldRow(ws, "Row.getValues", row);
  if (!r) {
    return [];
  }
  return options?.base === 1 ? rowValues(r) : rowGetValues(r);
}

/**
 * Replace the row's values. **This replaces the row** — cells the values do not carry
 * are cleared — unlike `Column.setValues`, which only writes the indices it is given.
 *
 * With `{ base }` an array is read exactly as {@link getValues} with the same `base`
 * returns it. Without it the array is read as column-numbered when index 0 is a hole
 * and as 0-based otherwise — which cannot tell a column-numbered array from a 0-based
 * one whose column A is empty, so pass `base` when the shape is known. An object is
 * read by column key.
 */
export function setValues(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  values: RowValues | readonly CellValueInput[],
  options?: ValuesOptions
): void {
  checkRowValues("Row.setValues", values);
  checkBase("Row.setValues", options);
  rowSetValues(
    editRow(ws, "Row.setValues", row),
    (options && Array.isArray(values) ? numberedValues(values, options.base) : values) as RowValues
  );
}

/** A values `base` is `0` or `1`; anything else re-indexed the array to nowhere — a write cleared the row. */
function checkBase(where: string, options: ValuesOptions | undefined): void {
  const check = argumentChecks(where);
  if (check.record("options", options)) {
    if (options.base !== 0 && options.base !== 1) {
      check.fail("options.base", "0 or 1", options.base);
    }
  }
}

// --- individual style facets ---

/**
 * Setting a facet on a row applies it to the row *and* to every cell that
 * already exists in it. Each call walks the row once, so {@link setStyle} is
 * the better call when setting several facets together — it walks the row a
 * single time for all of them.
 *
 * Passing `undefined` clears the facet, on the row and on its cells.
 */
export function setNumFmt(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  value: string | NumFmt | undefined
): void {
  checkStyleFacet("Row.setNumFmt", "numFmt", value, "value");
  rowSetNumFmt(editRow(ws, "Row.setNumFmt", row), value);
}
/** Set the font of the row and of every cell it already has; `undefined` clears it. See {@link setNumFmt}. */
export function setFont(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  value: Partial<Font> | undefined
): void {
  checkStyleFacet("Row.setFont", "font", value, "value");
  rowSetFont(editRow(ws, "Row.setFont", row), value);
}
/** Set the alignment of the row and of every cell it already has; `undefined` clears it. See {@link setNumFmt}. */
export function setAlignment(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  value: Partial<Alignment> | undefined
): void {
  checkStyleFacet("Row.setAlignment", "alignment", value, "value");
  rowSetAlignment(editRow(ws, "Row.setAlignment", row), value);
}
/** Set the borders of the row and of every cell it already has; `undefined` clears it. See {@link setNumFmt}. */
export function setBorder(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  value: Partial<Borders> | undefined
): void {
  checkStyleFacet("Row.setBorder", "border", value, "value");
  rowSetBorder(editRow(ws, "Row.setBorder", row), value);
}
/** Set the fill of the row and of every cell it already has; `undefined` clears it. See {@link setNumFmt}. */
export function setFill(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  value: Fill | undefined
): void {
  checkStyleFacet("Row.setFill", "fill", value, "value");
  rowSetFill(editRow(ws, "Row.setFill", row), value);
}
/** Set the protection of the row and of every cell it already has; `undefined` clears it. See {@link setNumFmt}. */
export function setProtection(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  value: Partial<Protection> | undefined
): void {
  checkStyleFacet("Row.setProtection", "protection", value, "value");
  rowSetProtection(editRow(ws, "Row.setProtection", row), value);
}

/** A copy of one facet of row `row`'s own style. Creates nothing and shares nothing. */
function facet<K extends keyof Style>(
  ws: WorksheetData | StreamWorksheet,
  where: string,
  row: number,
  key: K
): Style[K] | undefined {
  const value = heldRow(ws, where, row)?.style[key];
  return value === undefined ? undefined : copyStyleFacet(key, value);
}

/**
 * The per-facet readers answer what the row's own style says — `getFont(ws, row)`
 * equals `getStyle(ws, row).font` — but as a **read**: they never create the row,
 * and what they return is a copy, so editing it changes nothing. Change a facet with
 * its setter ({@link setFont}, …).
 */
export function getNumFmt(
  ws: WorksheetData | StreamWorksheet,
  row: number
): Style["numFmt"] | undefined {
  return facet(ws, "Row.getNumFmt", row, "numFmt");
}
/** A copy of row `row`'s own font, or `undefined`. Creates nothing; see {@link getNumFmt}. */
export function getFont(
  ws: WorksheetData | StreamWorksheet,
  row: number
): Partial<Font> | undefined {
  return facet(ws, "Row.getFont", row, "font");
}
/** A copy of row `row`'s own alignment, or `undefined`. Creates nothing. */
export function getAlignment(
  ws: WorksheetData | StreamWorksheet,
  row: number
): Partial<Alignment> | undefined {
  return facet(ws, "Row.getAlignment", row, "alignment");
}
/** A copy of row `row`'s own borders, or `undefined`. Creates nothing. */
export function getBorder(
  ws: WorksheetData | StreamWorksheet,
  row: number
): Partial<Borders> | undefined {
  return facet(ws, "Row.getBorder", row, "border");
}
/** A copy of row `row`'s own fill, or `undefined`. Creates nothing. */
export function getFill(ws: WorksheetData | StreamWorksheet, row: number): Fill | undefined {
  return facet(ws, "Row.getFill", row, "fill");
}
/** A copy of row `row`'s own protection, or `undefined`. Creates nothing. */
export function getProtection(
  ws: WorksheetData | StreamWorksheet,
  row: number
): Partial<Protection> | undefined {
  return facet(ws, "Row.getProtection", row, "protection");
}

// --- commit ---

/**
 * Commit row `row`. Meaningful to a streaming writer (`Stream.WorksheetWriter`), which writes the row out and releases
 * it; on a document-model sheet it does nothing — it does not create the row.
 *
 * @deprecated A document-model sheet has nothing to commit — every edit applies as it is made — and a streamed row is
 * committed through its handle: `Stream.commitRow(ws.getRow(row))`. Removed in 1.0.
 */
export function commit(ws: WorksheetData | StreamWorksheet, row: number): void {
  // Only a streaming writer has a commit step, and only there is the row materialised to be written out. A
  // document-model sheet is checked, as every reader checks it, and left as it was.
  if (asStreamWorksheet(ws)) {
    rowCommit(getRow(ws as WorksheetData, row));
  } else {
    findRow(ws as WorksheetData, row);
  }
}

// --- printing ---

/**
 * Add a manual horizontal page break **below** `row`, the equivalent of
 * Excel's *Page Layout → Breaks → Insert Page Break*. Applies to printing and
 * to `Pdf.fromExcel`; it does not affect the on-screen grid.
 *
 * The break spans the full width of the sheet, which is the only kind Excel can
 * author or render — see `rowAddPageBreak` for why no column band is offered.
 *
 * @example
 * ```typescript
 * import { Row, Workbook } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Report");
 * Row.addPageBreak(ws, 20); // page 2 starts at row 21
 * ```
 */
export function addPageBreak(ws: WorksheetData | StreamWorksheet, row: number): void {
  rowAddPageBreak(editRow(ws, "Row.addPageBreak", row));
}

// --- deprecated names ---

/** @deprecated Use {@link isHidden}. Removed in 1.0. */
export const getHidden = isHidden;
/**
 * @deprecated Use {@link getValues} with `{ base: 1 }`. Removed in 1.0.
 */
export function values(ws: WorksheetData | StreamWorksheet, row: number): CellValue[] {
  return getValues(ws, row, { base: 1 });
}

/**
 * @deprecated Address the cell by position through the `Cell` namespace
 * (`Cell.getValue(ws, { row, col })`). Removed in 1.0.
 */
export function getCell(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  col: string | number
): CellData {
  return rowGetCell(editRow(ws, "Row.getCell", row), col);
}
/**
 * @deprecated Use {@link eachAddress}, which visits by position and takes its
 * options last. Removed in 1.0.
 */
export function eachCell(
  ws: WorksheetData | StreamWorksheet,
  row: number,
  optOrCallback: { includeEmpty?: boolean } | ((cell: CellData, colNumber: number) => void),
  maybeCallback?: (cell: CellData, colNumber: number) => void
): void {
  // An absent row has no cells to visit, with or without `includeEmpty`; do not create it to find that out.
  const r = heldRow(ws, "Row.eachCell", row);
  if (r) {
    rowEachCell(r, optOrCallback as never, maybeCallback);
  }
}
