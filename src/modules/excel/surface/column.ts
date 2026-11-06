/**
 * `Column` namespace surface — column-level operations addressed by key, letter,
 * or 1-based number.
 *
 * `import { Column } from "documonster/excel"` → `Column.setWidth(ws, "A", 20)`,
 * `Column.setHeader(ws, 1, "Name")`, `Column.setStyle(ws, "B", { numFmt })`,
 * `Column.getValues(ws, "total")`, `Column.autoFit(ws, "B")`.
 *
 * The `col` parameter is a column key, an `"A"` letter, or a 1-based column
 * number — hence the `string | number` unions in the signatures below.
 *
 * Every reader creates nothing — no column record, no row, no cell — except
 * {@link getStyle}, which hands out the column's live style object for editing.
 */
import { cellType, checkCellValue } from "@excel/core/cell";
import type { ColumnDefn, ColumnHeaderValue } from "@excel/core/column";
import {
  columnAddPageBreak,
  columnDefn,
  columnCollapsed,
  columnHidden,
  columnIsDefault,
  columnOutlineLevel,
  columnSetHidden,
  columnSetOutlineLevel
} from "@excel/core/column";
import { ValueType } from "@excel/core/enums";
import { invalidateSharedCellStyle } from "@excel/core/style-sharing";
import {
  autoFitColumn,
  autoFitColumns,
  getActualColumnCount,
  getColumnCount,
  getColumnDefinitions,
  setColumns,
  spliceColumns
} from "@excel/core/worksheet";
import type { AutoFitColumnOptions } from "@excel/core/worksheet";
import {
  columnSetAlignment,
  columnSetBorder,
  columnSetFill,
  columnSetFont,
  columnSetHeader,
  columnSetKey,
  columnSetNumFmt,
  columnSetProtection,
  columnSetStyle,
  columnCreate,
  columnSetValuesAt,
  columnValuesAt,
  checkWorksheet,
  checkWorksheetAttached,
  findColumn,
  findColumnNumber,
  getColumn,
  getColumnKey
} from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { ExcelColumnOutOfBoundsError, ExcelStreamStateError } from "@excel/errors";
import type { StreamWorksheet } from "@excel/surface/stream-targets";
import { asStreamWorksheet, refuseWrittenRows } from "@excel/surface/stream-targets";
import type {
  Alignment,
  Borders,
  CellValue,
  CellValueInput,
  Fill,
  Font,
  NumFmt,
  Protection,
  Style
} from "@excel/types";
import { argumentChecks, checkNumber, checkSchema } from "@excel/utils/arguments";
import { colCache } from "@excel/utils/col-cache";
import { copyPlain } from "@excel/utils/copy-plain";
import { copyStyle, copyStyleFacet } from "@excel/utils/copy-style";
import { numberedValues } from "@excel/utils/numbered-values";
import type { ValuesBase, ValuesOptions } from "@excel/utils/numbered-values";
import { checkStyle, checkStyleFacet } from "@excel/utils/style-checks";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

/** @deprecated Use `Worksheet.Handle`. Removed in 1.0. */
export type Sheet = WorksheetData;

/** Options for {@link each} and {@link eachAddress}. */
export interface EachOptions {
  /** Also visit columns (or cells) that hold no value. */
  includeEmpty?: boolean;
}

/** How a values array is indexed — the one type `Row`, `Column`, `Worksheet` and `Stream` share. */
export type { ValuesBase, ValuesOptions };

/** Options for {@link autoFit}. */
export interface AutoFitOptions extends AutoFitColumnOptions {
  /** First column of the range form. Defaults to the left of the used range. */
  start?: number | string;
  /** Last column of the range form. Defaults to the right of the used range. */
  end?: number | string;
}

// --- definitions / counting ---

/**
 * Every declared column's definition as plain data — the shape
 * {@link setDefinitions} takes. It is a copy, style included: editing it changes nothing.
 */
export function getDefinitions(ws: WorksheetData | StreamWorksheet): ColumnDefn[] {
  readColumns(ws);
  return copyPlain(getColumnDefinitions(ws));
}

/**
 * Declare the sheet's columns — header, key, width, style — from column 1
 * onward. Headers are written into the first rows.
 *
 * Every definition is checked before any column changes, as the single-field setters check the same fields: a
 * width from 0 to 255, an outline level from 0 to 7, a boolean `hidden`, keys that are distinct non-empty strings,
 * headers a cell can hold and a valid style.
 */
export function setDefinitions(
  ws: WorksheetData | StreamWorksheet,
  definitions: readonly ColumnDefn[]
): void {
  editColumns(ws, "set the column definitions");
  const where = "Column.setDefinitions";
  const check = argumentChecks(where);
  check.required("definitions", definitions);
  check.array("definitions", definitions);
  const keys = new Set<string>();
  definitions.forEach((definition, i) => {
    const at = `definitions[${i}]`;
    check.required(at, definition);
    check.record(at, definition);
    check.number(`${at}.width`, definition.width, { min: 0, max: MAX_WIDTH });
    check.number(`${at}.outlineLevel`, definition.outlineLevel, { min: 0, max: 7, integer: true });
    check.boolean(`${at}.hidden`, definition.hidden);
    if (definition.key !== undefined) {
      check.string(`${at}.key`, definition.key, { nonEmpty: true });
      if (keys.has(definition.key)) {
        check.fail(`${at}.key`, "a key no other column has", definition.key);
      }
      keys.add(definition.key);
    }
    checkHeader(definition.header, where, `${at}.header`);
    if (definition.style !== undefined) {
      checkStyle(where, definition.style, `${at}.style`);
    }
  });
  setColumns(ws, definitions);
}

/**
 * `ws` as the column helpers take it. A streaming worksheet writer stands in for a worksheet there — it carries the same
 * column fields — so the readers of a column's own settings serve it as they are.
 */
function readColumns(ws: WorksheetData | StreamWorksheet): asserts ws is WorksheetData {
  // Nothing to check: the assertion is the point.
  void ws;
}

/**
 * {@link readColumns}, for an edit: a streaming writer writes its columns with its first committed row, so a setting
 * written into the file (width, style, header…) is refused after that rather than accepted and lost. A key is not
 * written — it only says which property of a row object `addRow` reads into the column — so a `keyOnly` edit stays
 * open until the sheet is committed.
 */
function editColumns(
  ws: WorksheetData | StreamWorksheet,
  operation: string,
  { keyOnly = false }: { keyOnly?: boolean } = {}
): asserts ws is WorksheetData {
  const writer = asStreamWorksheet(ws);
  if (writer) {
    writer._checkOpen(operation);
    if (!keyOnly && writer.areColumnsWritten()) {
      throw new ExcelStreamStateError(
        operation,
        "the sheet's columns were written with its first committed row"
      );
    }
  }
}

/** Excel's widest column, in characters ("Excel specifications and limits"). */
const MAX_WIDTH = 255;

/** A header is written into the column's first cells, so each of its values is checked as `Cell.setValue` checks one. */
function checkHeader(header: ColumnHeaderValue | undefined, where: string, argument: string): void {
  if (Array.isArray(header)) {
    header.forEach((value, i) => checkCellValue(value, { where, field: `${argument}[${i}]` }));
  } else if (header !== undefined) {
    checkCellValue(header, { where, field: argument });
  }
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

/**
 * The number of the last column that exists — some row has a cell in it, or it has a setting of its own (see
 * {@link has}) — the sheet's column count, and the mirror of `Row.getCount`.
 */
export function getCount(ws: WorksheetData): number {
  refuseWrittenRows(ws, "Column.getCount", "count columns by their cells");
  return lastColumn(ws);
}

/** The last column {@link has} answers `true` for, or `0`. */
function lastColumn(ws: WorksheetData): number {
  let last = getColumnCount(ws);
  for (let n = ws._columns.length; n > last; n--) {
    const column = ws._columns[n - 1];
    if (column !== undefined && !columnIsDefault(column)) {
      last = n;
      break;
    }
  }
  return last;
}

/** How many columns hold a value in at least one row. */
export function getActualCount(ws: WorksheetData): number {
  refuseWrittenRows(ws, "Column.getActualCount", "count columns by their cells");
  return getActualColumnCount(ws);
}

/**
 * Whether column `col` exists — some row has a cell in it, or it has a setting of its own (a width, a style, hidden, an
 * outline level) — the mirror of `Row.has`. Creates nothing.
 */
export function has(ws: WorksheetData, col: string | number): boolean {
  refuseWrittenRows(ws, "Column.has", "read the cells of a column");
  const n = findColumnNumber(ws, col);
  const own = ws._columns[n - 1];
  if (own !== undefined && !columnIsDefault(own)) {
    return true;
  }
  return ws._rows.some(row => row !== undefined && row.cells[n - 1] !== undefined);
}

/**
 * The number of the last column that exists — {@link getCount}, the mirror of `Row.getLast` — or `undefined` when
 * none does. A column exists when some row has a cell in it, even one holding no value (a styled empty cell, say), or
 * when it has a setting of its own (a width, a style, hidden, an outline level); use {@link getActualCount} to count
 * only columns holding a value.
 */
export function getLast(ws: WorksheetData): number | undefined {
  refuseWrittenRows(ws, "Column.getLast", "find the last column by its cells");
  const n = lastColumn(ws);
  return n > 0 ? n : undefined;
}

/**
 * Remove `count` columns starting at `start` and insert the given columns of
 * values in their place, as `Array.prototype.splice` does. *
 * Everything that refers to cells follows them, on every sheet, as it does in Excel: formulas, defined names,
 * tables, conditional formats, data validations, charts, sparklines, pivot sources, the auto filter, print areas
 * and drawings. A reference to cells the edit deletes becomes `#REF!`.
 */
export function splice(
  ws: WorksheetData,
  start: number,
  count: number,
  ...inserts: (readonly CellValueInput[])[]
): void {
  refuseWrittenRows(ws, "Column.splice", "splice columns");
  spliceColumns(ws, start, count, ...(inserts as CellValue[][]));
}

/**
 * Visit each column that holds a value in at least one row, in order, with its
 * number — the mirror of `Row.each`. With `{ includeEmpty: true }` every column
 * from 1 to {@link getLast} is visited, a column with only a width included. Creates nothing.
 */
export function each(
  ws: WorksheetData,
  callback: (colNumber: number) => void,
  options?: EachOptions
): void {
  refuseWrittenRows(ws, "Column.each", "walk the cells of its columns");
  if (typeof callback !== "function") {
    argumentChecks("Column.each").fail("callback", "a function", callback);
  }
  const n = lastColumn(ws);
  const used: boolean[] = [];
  if (!options?.includeEmpty) {
    for (const row of ws._rows) {
      row?.cells.forEach((cell, index) => {
        if (cell && cellType(cell) !== ValueType.Null) {
          used[index + 1] = true;
        }
      });
    }
  }
  for (let c = 1; c <= n; c++) {
    if (options?.includeEmpty || used[c]) {
      callback(c);
    }
  }
}

/**
 * Visit each cell of column `col` that holds a value, in row order, as
 * `(address, row, col)` — the shape `Range.eachAddress` and `Row.eachAddress`
 * hand over. With `{ includeEmpty: true }` every row up to the column's last
 * cell is visited — the mirror of `Row.eachAddress`, which stops at the row's
 * last cell. Creates nothing.
 */
export function eachAddress(
  ws: WorksheetData,
  col: string | number,
  callback: (address: string, row: number, col: number) => void,
  options?: EachOptions
): void {
  refuseWrittenRows(ws, "Column.eachAddress", "walk the cells of a column");
  if (typeof callback !== "function") {
    argumentChecks("Column.eachAddress").fail("callback", "a function", callback);
  }
  const colNumber = findColumnNumber(ws, col);
  const letter = colCache.n2l(colNumber);
  const rows = ws._rows;
  let last = rows.length;
  while (last > 0 && !rows[last - 1]?.cells[colNumber - 1]) {
    last--;
  }
  for (let r = 1; r <= last; r++) {
    const cell = rows[r - 1]?.cells[colNumber - 1];
    if (options?.includeEmpty || (cell && cellType(cell) !== ValueType.Null)) {
      callback(letter + r, r, colNumber);
    }
  }
}

const autoFitColumnShape = { includeHiddenRows: s.boolean().optional() };
const autoFitColumnOptionsSchema = s.object(autoFitColumnShape);
const _autoFitColumnOptionsMatch: SchemaMatches<
  typeof autoFitColumnOptionsSchema,
  AutoFitColumnOptions
> = true;
void _autoFitColumnOptionsMatch;

const column = () =>
  s.union([s.string(), s.number()], { expected: "a column letter or number" }).optional();
const autoFitOptionsSchema = s.object({ ...autoFitColumnShape, start: column(), end: column() });
const _autoFitOptionsMatch: SchemaMatches<typeof autoFitOptionsSchema, AutoFitOptions> = true;
void _autoFitOptionsMatch;

/**
 * Fit column widths to their content, as Excel's AutoFit does: one column, or a
 * span (`{ start, end }`, defaulting to the used range) when given only options.
 */
export function autoFit(
  ws: WorksheetData,
  col: string | number,
  options?: AutoFitColumnOptions
): void;
export function autoFit(ws: WorksheetData, options?: AutoFitOptions): void;
export function autoFit(
  ws: WorksheetData,
  colOrOptions?: string | number | AutoFitOptions,
  options?: AutoFitColumnOptions
): void {
  refuseWrittenRows(ws, "Column.autoFit", "fit a column to its cells");
  if (typeof colOrOptions === "string" || typeof colOrOptions === "number") {
    const col = autoFitColumnNumber(ws, colOrOptions, "col");
    checkSchema("Column.autoFit", "options", autoFitColumnOptionsSchema.optional(), options);
    autoFitColumn(ws, col, options);
    return;
  }
  checkWorksheet(ws);
  checkSchema("Column.autoFit", "options", autoFitOptionsSchema.optional(), colOrOptions);
  const { start, end, ...rest } = colOrOptions ?? {};
  // Resolved here, so a column past XFD is refused rather than counted up to.
  autoFitColumns(
    ws,
    start === undefined ? undefined : autoFitColumnNumber(ws, start, "options.start"),
    end === undefined ? undefined : autoFitColumnNumber(ws, end, "options.end"),
    rest
  );
}

/** `findColumnNumber`, with a column off the sheet reported at `field` of `Column.autoFit`. */
function autoFitColumnNumber(ws: WorksheetData, col: string | number, field: string): number {
  try {
    return findColumnNumber(ws, col);
  } catch (cause) {
    if (cause instanceof ExcelColumnOutOfBoundsError) {
      throw new ExcelColumnOutOfBoundsError(col, undefined, {
        field,
        where: "Column.autoFit",
        cause
      });
    }
    throw cause;
  }
}

// --- identity ---

/**
 * The 1-based column number for a key, letter or number.
 *
 * This is the bridge from a column *key* to the `{ row, col }` position the `Cell`
 * namespace takes: `Cell.setValue(ws, { row, col: Column.getNumber(ws, "total") }, v)`.
 */
export function getNumber(ws: WorksheetData | StreamWorksheet, col: string | number): number {
  readColumns(ws);
  return findColumnNumber(ws, col);
}

/** The column letter (`"A"`, `"AB"`) for a key, letter or number. */
export function getLetter(ws: WorksheetData | StreamWorksheet, col: string | number): string {
  readColumns(ws);
  return colCache.n2l(findColumnNumber(ws, col));
}

// --- width ---

/** The column's width in characters, or `undefined` when it has none of its own. Creates nothing. */
export function getWidth(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): number | undefined {
  readColumns(ws);
  return findColumn(ws, col)?.width;
}
/**
 * Set the column's width in characters — from 0 to 255, Excel's widest column. `undefined` clears it, so the column
 * takes the sheet's default width again, as {@link getWidth} reports a column with none.
 *
 * @throws {ExcelArgumentError} when `width` is not a number in that range.
 */
export function setWidth(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  width: number | undefined
): void {
  editColumns(ws, "set a column width");
  if (width === undefined) {
    const existing = findColumn(ws, col);
    checkWorksheetAttached(ws);
    if (existing) {
      existing.width = undefined;
      existing.customWidth = undefined;
    }
    return;
  }
  checkNumber("Column.setWidth", "width", width, { min: 0, max: MAX_WIDTH });
  const column = getColumn(ws, col);
  column.width = width;
  // The caller's width now, whatever the file said about the one before.
  column.customWidth = undefined;
}

// --- header / key ---

/** The column's header, as {@link setHeader} or a definition set it, or `undefined`. Creates nothing. */
export function getHeader(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): ColumnHeaderValue | undefined {
  readColumns(ws);
  return findColumn(ws, col)?.header;
}
/** Set the column's header and write it into the column's first cell(s) — one per value of an array header. */
export function setHeader(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  header: ColumnHeaderValue
): void {
  editColumns(ws, "set a column header");
  checkHeader(header, "Column.setHeader", "header");
  columnSetHeader(getColumn(ws, col), header);
}
/** The column's key, or `undefined`. Creates nothing. */
export function getKey(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): string | undefined {
  readColumns(ws);
  return findColumn(ws, col)?.key;
}
/**
 * Give the column a key, which then names it wherever a column reference is taken — `Column.setWidth(ws, "total", 12)`.
 *
 * **A key wins over a column letter of the same text**: after `setKey(ws, 1, "F")`, `"F"` names column A for every
 * `Column`, `Cell`-by-key and `Row`-by-key call. Choose keys that do not read as column letters to keep both usable.
 *
 * @throws {ExcelArgumentError} when `key` is not a non-empty string, or another column already has it.
 */
export function setKey(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  key: string
): void {
  editColumns(ws, "set a column key", { keyOnly: true });
  const check = argumentChecks("Column.setKey");
  check.required("key", key);
  check.string("key", key, { nonEmpty: true });
  const holder = getColumnKey(ws, key);
  if (holder !== undefined && holder.number !== findColumnNumber(ws, col)) {
    check.fail(
      "key",
      `a key no other column has (column ${colCache.n2l(holder.number)} has it)`,
      key
    );
  }
  columnSetKey(getColumn(ws, col), key);
}

// --- visibility / outline ---

/** Whether the column is hidden. Creates nothing. */
export function isHidden(ws: WorksheetData | StreamWorksheet, col: string | number): boolean {
  readColumns(ws);
  const c = findColumn(ws, col);
  return c !== undefined && columnHidden(c);
}
/**
 * Whether the column is collapsed — its outline level is at or above the sheet's column outline level, as
 * `Row.isCollapsed` answers for a row. `false` for a column that does not exist. Creates nothing.
 */
export function isCollapsed(ws: WorksheetData | StreamWorksheet, col: string | number): boolean {
  readColumns(ws);
  const c = findColumn(ws, col);
  return c !== undefined && columnCollapsed(c);
}
/** Hide or show the column. */
export function setHidden(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  hidden: boolean
): void {
  editColumns(ws, "hide a column");
  argumentChecks("Column.setHidden").required("hidden", hidden);
  argumentChecks("Column.setHidden").boolean("hidden", hidden);
  columnSetHidden(getColumn(ws, col), hidden);
}
/** The column's outline (grouping) level, `0` when it has none. Creates nothing. */
export function getOutlineLevel(ws: WorksheetData | StreamWorksheet, col: string | number): number {
  readColumns(ws);
  const c = findColumn(ws, col);
  return c ? columnOutlineLevel(c) : 0;
}
/** Set the column's outline (grouping) level, a whole number from 0 to 7. */
export function setOutlineLevel(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  level: number
): void {
  editColumns(ws, "set a column outline level");
  checkNumber("Column.setOutlineLevel", "level", level, { min: 0, max: 7, integer: true });
  columnSetOutlineLevel(getColumn(ws, col), level);
}

// --- style ---

/**
 * The column's own style object, live — mutating it is how a caller changes the column.
 * Because it is handed out for editing, this is the one reader here that creates the
 * column record when it does not exist yet. To *read* the style, use {@link findStyle};
 * to read one facet, {@link getFont}, {@link getFill}, … — they create nothing.
 *
 * Cells the column has already styled point at a *snapshot* of these facets, so that
 * mutation does not reach them (as it does not today). The snapshot is dropped here
 * because the caller may mutate in place, and a cell created afterwards must copy
 * what the column says then — see `core/style-sharing.ts`.
 */
export function getStyle(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): Partial<Style> {
  editColumns(ws, "edit a column style");
  const { style } = getColumn(ws, col);
  invalidateSharedCellStyle(style);
  return style;
}
/** Merge `style` into the column's style and the style of the cells it applies to; see {@link setNumFmt}. */
export function setStyle(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  style: Partial<Style>
): void {
  editColumns(ws, "set a column style");
  checkStyle("Column.setStyle", style);
  columnSetStyle(getColumn(ws, col), style);
}

/**
 * A copy of the column's own style, or `undefined` when the sheet has no record for the
 * column — the read-only counterpart of {@link getStyle}: it creates nothing, and editing
 * what it returns changes nothing.
 */
export function findStyle(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): Partial<Style> | undefined {
  readColumns(ws);
  const column = findColumn(ws, col);
  return column ? (copyStyle(column.style) as Partial<Style>) : undefined;
}

// --- number format ---

/**
 * Set the column's number format.
 *
 * The mirror of `Cell.setNumFmt` / `Row.setNumFmt`, so a format can be set at
 * whichever level owns it. Each facet setter walks the column once, so
 * {@link setStyle} is the better call when setting several facets together.
 *
 * The format lands on the column *and* on the cells of every row that holds a
 * value, which is not quite "every existing cell": a row with values elsewhere
 * gains a cell in this column if it had none, and a materialised-but-empty cell in
 * a row with no values at all is skipped. That is the behaviour of every
 * `columnSet*`, not a property of this one.
 *
 * Passing `undefined` clears the format.
 *
 * {@link getNumFmt} reads it back.
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
export function setNumFmt(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  value: string | NumFmt | undefined
): void {
  editColumns(ws, "set a column number format");
  checkStyleFacet("Column.setNumFmt", "numFmt", value, "value");
  columnSetNumFmt(getColumn(ws, col), value);
}

/** Set the font of the column and of the cells it applies to; `undefined` clears it. See {@link setNumFmt}. */
export function setFont(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  value: Partial<Font> | undefined
): void {
  editColumns(ws, "set a column font");
  checkStyleFacet("Column.setFont", "font", value, "value");
  columnSetFont(getColumn(ws, col), value);
}
/** Set the alignment of the column and of the cells it applies to; `undefined` clears it. See {@link setNumFmt}. */
export function setAlignment(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  value: Partial<Alignment> | undefined
): void {
  editColumns(ws, "set a column alignment");
  checkStyleFacet("Column.setAlignment", "alignment", value, "value");
  columnSetAlignment(getColumn(ws, col), value);
}
/** Set the borders of the column and of the cells it applies to; `undefined` clears it. See {@link setNumFmt}. */
export function setBorder(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  value: Partial<Borders> | undefined
): void {
  editColumns(ws, "set a column border");
  checkStyleFacet("Column.setBorder", "border", value, "value");
  columnSetBorder(getColumn(ws, col), value);
}
/** Set the fill of the column and of the cells it applies to; `undefined` clears it. See {@link setNumFmt}. */
export function setFill(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  value: Fill | undefined
): void {
  editColumns(ws, "set a column fill");
  checkStyleFacet("Column.setFill", "fill", value, "value");
  columnSetFill(getColumn(ws, col), value);
}
/** Set the protection of the column and of the cells it applies to; `undefined` clears it. See {@link setNumFmt}. */
export function setProtection(
  ws: WorksheetData | StreamWorksheet,
  col: string | number,
  value: Partial<Protection> | undefined
): void {
  editColumns(ws, "set a column protection");
  checkStyleFacet("Column.setProtection", "protection", value, "value");
  columnSetProtection(getColumn(ws, col), value);
}

/** A copy of one facet of column `col`'s own style. Creates nothing and shares nothing. */
function facet<K extends keyof Style>(
  ws: WorksheetData,
  col: string | number,
  key: K
): Style[K] | undefined {
  const value = findColumn(ws, col)?.style[key];
  return value === undefined ? undefined : copyStyleFacet(key, value);
}

/**
 * The per-facet readers answer what the column's own style says — `getFont(ws, col)`
 * equals `getStyle(ws, col).font` — but as a **read**: they never create the column
 * record, and what they return is a copy, so editing it changes nothing. Change a
 * facet with its setter ({@link setFont}, …).
 */
export function getNumFmt(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): Style["numFmt"] | undefined {
  readColumns(ws);
  return facet(ws, col, "numFmt");
}
/** A copy of the column's own font, or `undefined`. Creates nothing; see {@link getNumFmt}. */
export function getFont(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): Partial<Font> | undefined {
  readColumns(ws);
  return facet(ws, col, "font");
}
/** A copy of the column's own alignment, or `undefined`. Creates nothing. */
export function getAlignment(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): Partial<Alignment> | undefined {
  readColumns(ws);
  return facet(ws, col, "alignment");
}
/** A copy of the column's own borders, or `undefined`. Creates nothing. */
export function getBorder(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): Partial<Borders> | undefined {
  readColumns(ws);
  return facet(ws, col, "border");
}
/** A copy of the column's own fill, or `undefined`. Creates nothing. */
export function getFill(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): Fill | undefined {
  readColumns(ws);
  return facet(ws, col, "fill");
}
/** A copy of the column's own protection, or `undefined`. Creates nothing. */
export function getProtection(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): Partial<Protection> | undefined {
  readColumns(ws);
  return facet(ws, col, "protection");
}

// --- definition ---

/**
 * The column's definition as plain data — an element of what {@link setDefinitions}
 * takes, with `hidden` / `outlineLevel` normalised the way the library
 * normalises them internally.
 *
 * It is a copy, style included: editing it changes nothing.
 *
 * Use this (or {@link getDefinitions} for every column) instead of copying
 * fields off a column handle by hand: hand-copying re-implements the
 * normalisation and silently drifts the moment a field or a fallback changes.
 */
export function getDefinition(
  ws: WorksheetData | StreamWorksheet,
  col: string | number
): ColumnDefn {
  readColumns(ws);
  return copyPlain(columnDefn(findColumn(ws, col) ?? columnCreate(ws, findColumnNumber(ws, col))));
}

// --- values ---

/**
 * Read the values down a column as a **sparse** array — an empty cell is a hole
 * rather than a `null` — indexed **0-based** (row 1 at index `0`) unless
 * `{ base: 1 }` asks for the index to be the row number.
 *
 * The read is non-destructive: it walks the rows that exist rather than asking
 * each one for a cell, so it neither materialises cells down the column nor
 * declares the column itself.
 *
 * @example
 * ```typescript
 * import { Cell, Column, Workbook } from "documonster/excel";
 *
 * const wb = Workbook.create();
 * const ws = Workbook.addWorksheet(wb, "Sales");
 * Column.setValues(ws, "C", [120, -40, 75]);
 * const [first] = Column.getValues(ws, "C"); // 120 — row 1
 *
 * // Colour the losses red. With `base: 1` the index *is* the row number, and
 * // `Column.getNumber` turns the reference into the column number `Cell` takes.
 * const colNumber = Column.getNumber(ws, "C");
 * Column.getValues(ws, "C", { base: 1 }).forEach((value, rowNumber) => {
 *   if (typeof value === "number" && value < 0) {
 *     Cell.setFont(ws, { row: rowNumber, col: colNumber }, { color: { argb: "FFC00000" } });
 *   }
 * });
 * ```
 */
export function getValues(
  ws: WorksheetData,
  col: string | number,
  options?: ValuesOptions
): CellValue[] {
  refuseWrittenRows(ws, "Column.getValues", "read the values of a column");
  checkBase("Column.getValues", options);
  return columnValuesAt(ws, findColumnNumber(ws, col), options?.base === 1 ? 1 : 0);
}

/**
 * Write values down a column. Takes anything `Cell.setValue` takes — a number, a
 * string, a formula, rich text, a hyperlink, a `Date` or a civil `Temporal`.
 *
 * With `{ base }` the array is read exactly as {@link getValues} with the same
 * `base` returns it. Without it the array is read as **1-based by row number** when
 * index 0 is a hole and as **0-based** when index 0 is present, so a plain array
 * starts at row 1:
 *
 * ```typescript
 * Column.setValues(ws, "A", [1, 2, 3]); // rows 1, 2, 3
 * Column.setValues(ws, "A", [, 1, 2, 3]); // rows 1, 2, 3 — the same thing
 * Column.setValues(ws, "A", [, 1, 2], { base: 0 }); // rows 2, 3
 * ```
 *
 * The two unbased readings are indistinguishable when a 0-based array's own first
 * slot is a hole — pass `base` whenever the shape is known.
 *
 * **This writes, it does not replace.** Only the indices the array carries are
 * touched — unlike `Row.setValues`, which resets the row — so a shorter array
 * does not truncate the column and a value already in a row the array skips
 * survives. Clear a cell by passing `null` at its index.
 */
export function setValues(
  ws: WorksheetData,
  col: string | number,
  values: readonly CellValueInput[],
  options?: ValuesOptions
): void {
  refuseWrittenRows(ws, "Column.setValues", "write values down a column");
  const check = argumentChecks("Column.setValues");
  check.required("values", values);
  check.array("values", values);
  checkBase("Column.setValues", options);
  // Every value before any cell is written, as `Row.setValues` does: a refusal used to leave the earlier ones written.
  values.forEach((value, i) =>
    checkCellValue(value, { where: "Column.setValues", field: `values[${i}]` })
  );
  columnSetValuesAt(
    ws,
    findColumnNumber(ws, col),
    options ? numberedValues(values, options.base) : values
  );
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
export function addPageBreak(ws: WorksheetData | StreamWorksheet, col: string | number): void {
  readColumns(ws);
  asStreamWorksheet(ws)?._checkOpen("add a page break");
  columnAddPageBreak(getColumn(ws, col));
}

// --- deprecated names ---

/** @deprecated Use {@link isHidden}. Removed in 1.0. */
export const getHidden = isHidden;
/**
 * @deprecated Use {@link getValues} with `{ base: 1 }`. Removed in 1.0.
 */
export function values(ws: WorksheetData, col: string | number): CellValue[] {
  return getValues(ws, col, { base: 1 });
}
