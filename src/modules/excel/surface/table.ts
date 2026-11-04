import { checkTable, checkTableColumn } from "@excel/core/feature-handles";
import { invalidateSharedCellStyle } from "@excel/core/style-sharing";
import type {
  TableColumnView,
  TableData,
  TableModel,
  TableRowEditOptions
} from "@excel/core/table";
import {
  checkTableName,
  createTable as createTableImpl,
  tableAddColumn,
  tableAddRow,
  tableApplyModel,
  tableColumnFilterButton,
  tableColumnName,
  tableColumnSetFilterButton,
  tableColumnSetName,
  tableColumnSetStyle,
  tableColumnSetTotalsRowFormula,
  tableColumnSetTotalsRowFunction,
  tableColumnSetTotalsRowLabel,
  tableColumnSetTotalsRowResult,
  tableColumnStyle,
  tableColumnTotalsRowFormula,
  tableColumnTotalsRowFunction,
  tableColumnTotalsRowLabel,
  tableColumnTotalsRowResult,
  tableDisplayName,
  tableCommit,
  tableGetColumn,
  tableGetColumnChecked,
  tableGetModel,
  tableHeaderRow,
  tableMarkLiveStyle,
  tableModel,
  tableName,
  tableRef,
  tableRemoveColumns,
  tableRemoveRows,
  tableSetDisplayName,
  tableSetHeaderRow,
  tableSetName,
  tableSetRef,
  tableSetShowColumnStripes,
  tableSetShowFirstColumn,
  tableSetShowLastColumn,
  tableSetShowRowStripes,
  tableSetTheme,
  tableSetTotalsRow,
  tableShowColumnStripes,
  tableShowFirstColumn,
  tableShowLastColumn,
  tableShowRowStripes,
  tableTheme,
  tableTotalsRow
} from "@excel/core/table";
import { addTable, getTable, removeTable } from "@excel/core/worksheet";
/**
 * `Table` namespace surface.
 *
 * Worksheet-level management:
 *   `Table.add(ws, model)`, `Table.get(ws, name)`, `Table.list(ws)`,
 *   `Table.remove(ws, name)`.
 * Table-handle operations:
 *   `Table.addRow(t, row)`, `Table.setName(t, name)`,
 *   `Table.getModel(t)`, `Table.getColumn(t, idx)`, …
 *
 * Every edit applies to the sheet as it is made, and is validated first: an edit
 * that throws leaves the table as it was. The sheet's cells are what the table
 * holds — `getModel` reads its rows from them, and an edit moves each cell's value
 * and formatting with it.
 *
 * The boolean readers (`hasHeaderRow`, `hasTotalsRow`, `hasFirstColumnStyle`,
 * `hasLastColumnStyle`, `hasRowStripes`, `hasColumnStripes`, `hasColumnFilterButton`)
 * answer `boolean | undefined`. `undefined` means the table carries no setting, so
 * Excel's default applies — and the defaults differ (a header row and a filter button
 * are on unless switched off; a totals row is off) — and it is what the matching
 * setter accepts to clear the setting. Collapsing it to the default would make a
 * read-then-write round trip turn an absent attribute into an explicit one, so the
 * third state is kept; test with `=== false` / `=== true` when the default matters.
 */
import { checkWorksheet, getTables } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { TableError } from "@excel/errors";
import type {
  CellValue,
  CellValueInput,
  Style,
  TableColumnProperties,
  TableProperties
} from "@excel/types";
import type { ArgumentChecks } from "@excel/utils/arguments";
import { argumentChecks, normalizeReference } from "@excel/utils/arguments";
import { copyStyle } from "@excel/utils/copy-style";
import { checkStyle } from "@excel/utils/style-checks";
import { excelNameKey } from "@utils/name-key";

const TOTALS_FUNCTIONS = [
  "none",
  "average",
  "countNums",
  "count",
  "max",
  "min",
  "stdDev",
  "var",
  "sum",
  "custom"
];

/** A totals row's cached result: what a formula's cached result may be. */
function checkTotalsResult(check: ArgumentChecks, field: string, value: unknown): void {
  const ok =
    value === undefined ||
    typeof value === "string" ||
    typeof value === "boolean" ||
    (typeof value === "number" && Number.isFinite(value)) ||
    (value instanceof Date && !Number.isNaN(value.getTime())) ||
    (value !== null &&
      typeof value === "object" &&
      typeof (value as { error?: unknown }).error === "string");
  if (!ok) {
    check.fail(field, "a finite number, a string, a boolean, a valid Date or { error }", value);
  }
}

/**
 * A column's properties, each of the type the table part writes. A label of the wrong type failed only when the
 * workbook was written, with a `TypeError` from inside the XML writer; a theme or flag of the wrong type was written as
 * given.
 */
function checkColumn(where: string, field: string, column: unknown): void {
  const check = argumentChecks(where);
  check.required(field, column);
  if (!check.record(field, column)) {
    return;
  }
  check.required(`${field}.name`, column.name);
  check.string(`${field}.name`, column.name);
  check.boolean(`${field}.filterButton`, column.filterButton);
  check.string(`${field}.totalsRowLabel`, column.totalsRowLabel);
  check.oneOf(`${field}.totalsRowFunction`, column.totalsRowFunction, TOTALS_FUNCTIONS);
  check.string(`${field}.totalsRowFormula`, column.totalsRowFormula);
  check.string(`${field}.calculatedColumnFormula`, column.calculatedColumnFormula);
  checkTotalsResult(check, `${field}.totalsRowResult`, column.totalsRowResult);
  checkStyle(where, column.style, `${field}.style`);
}

/** The table-level settings of a model given to `add` or `setModel`. */
function checkTableShape(where: string, model: unknown, argument: string): void {
  const check = argumentChecks(where);
  check.required(argument, model);
  if (!check.record(argument, model)) {
    return;
  }
  check.required(`${argument}.ref`, model.ref);
  check.string(`${argument}.ref`, model.ref);
  // A cell or range of the table's own sheet. The lenient decoder read `Other!A1` as column `OA`, and the table was
  // written there over whatever the sheet held.
  check.reference(`${argument}.ref`, model.ref, { range: true });
  check.boolean(`${argument}.headerRow`, model.headerRow);
  check.boolean(`${argument}.totalsRow`, model.totalsRow);
  check.boolean(
    `${argument}.qualifyImplicitStructuredReferences`,
    model.qualifyImplicitStructuredReferences
  );
  if (check.record(`${argument}.style`, model.style)) {
    check.string(`${argument}.style.theme`, model.style.theme);
    for (const flag of [
      "showFirstColumn",
      "showLastColumn",
      "showRowStripes",
      "showColumnStripes"
    ]) {
      check.boolean(`${argument}.style.${flag}`, model.style[flag]);
    }
  }
  check.required(`${argument}.columns`, model.columns);
  if (check.array(`${argument}.columns`, model.columns, { min: 1 })) {
    model.columns.forEach((column, i) => checkColumn(where, `${argument}.columns[${i}]`, column));
    checkColumnNamesUnique(
      where,
      model.columns.map(column => (column as { name: string }).name)
    );
  }
  check.required(`${argument}.rows`, model.rows);
  if (check.array(`${argument}.rows`, model.rows)) {
    model.rows.forEach((row, i) => {
      check.required(`${argument}.rows[${i}]`, row);
      check.array(`${argument}.rows[${i}]`, row);
    });
  }
}

/**
 * Refuse column names that repeat, compared as Excel compares them. Excel itself numbers a repeated header when one
 * is typed (`a`, `a` → `a`, `a2`), and a table read from a file is still named that way; but a caller who chose the
 * names writes formulas with them, and `T[a2]` names a column they never asked for — so the API refuses, naming the
 * name Excel would give.
 */
function checkColumnNamesUnique(where: string, names: readonly string[]): void {
  const taken = new Set(names.map(name => excelNameKey(name)));
  const seen = new Set<string>();
  for (const name of names) {
    const key = excelNameKey(name);
    if (!seen.has(key)) {
      seen.add(key);
      continue;
    }
    let n = 2;
    while (taken.has(excelNameKey(`${name}${n}`))) {
      n++;
    }
    throw new TableError(
      `${where}: column name ${JSON.stringify(name)} is already used by another column of the table ` +
        `(Excel compares column names without case); Excel would name it ${JSON.stringify(`${name}${n}`)}.`
    );
  }
}

/**
 * Add a table to the sheet, and return its handle.
 *
 * Everything is checked before a cell is written: the name, every column's properties (a style as a cell style is),
 * every value as a cell's is, that the column names are distinct, that the area overlaps no other table, and that the
 * name and display name are free in the workbook. A call that throws leaves the sheet as it was. The model is copied,
 * so changing the object afterwards does not change the table.
 *
 * @throws {TableError} when the name is not one Excel accepts — a space, a character other than a letter, digit,
 * underscore or period, or a name that reads as a cell reference (`T1`, `R1C1`); the message gives the nearest name
 * Excel does accept, and the table is not added under a name the caller did not choose. Also when two columns share a
 * name, the table would overlap another one, or the name is already taken.
 * @throws {ExcelArgumentError} when a property is of the wrong type.
 */
export function add(ws: WorksheetData, model: TableProperties): TableData {
  checkWorksheet(ws);
  checkTableShape("Table.add", model, "model");
  checkTableName("Table.add", "model.name", model.name);
  if (model.displayName !== undefined) {
    checkTableName("Table.add", "model.displayName", model.displayName);
  }
  return addTable(ws, { ...model, ref: normalizeReference(model.ref) });
}

/**
 * Remove a table from the sheet — by its handle, or by its name (compared as Excel compares names). Returns `true` if
 * the sheet had it. The cells keep their values, as when Excel converts a table to a range; the handle refuses edits
 * afterwards.
 *
 * @throws {PivotTableError} when a pivot table reads its data from the table.
 */
export function remove(ws: WorksheetData, table: string | TableData): boolean {
  checkWorksheet(ws);
  if (typeof table !== "string") {
    const check = argumentChecks("Table.remove");
    check.required("table", table);
    if (!check.record("table", table) || !("table" in table)) {
      check.fail("table", "a table name or a table handle", table);
    }
  }
  return removeTable(ws, table);
}

/** The table called `name` (compared as Excel compares names; the stored name or the one it was added under), or `undefined`. */
export function get(ws: WorksheetData, name: string): TableData | undefined {
  checkWorksheet(ws);
  argumentChecks("Table.get").string("name", name);
  return getTable(ws, name);
}

/** Set a column's style; see {@link getColumnStyle}. The style is checked as a cell style is. */
export function setColumnStyle(view: TableColumnView, value: Partial<Style> | undefined): void {
  checkTableColumn(view);
  checkStyle("Table.setColumnStyle", value, "value");
  tableColumnSetStyle(view, value);
}

/** The sheet's tables, in the order they were added. */
export function list(ws: WorksheetData): TableData[] {
  checkWorksheet(ws);
  return getTables(ws);
}

/**
 * @deprecated Use {@link add}. This builds a table bound to `worksheet` and writes its cells, but never registers it on
 * the sheet, so it is not written, not found by `get` or `list`, and its name is not reserved. Removed in 1.0.
 */
export function create(worksheet: WorksheetData, table?: TableModel): TableData {
  checkWorksheet(worksheet);
  return createTableImpl(worksheet, table);
}

/** @deprecated Every edit applies to the sheet as it is made; there is never anything to commit. Removed in 1.0. */
export const commit = tableCommit;

/**
 * Replace the table with `value`, as {@link add} would build it — cells, position and name. Everything is checked
 * first, as `add` checks it, so a model that is refused leaves the table and the sheet as they were.
 */
export function setModel(t: TableData, value: TableModel): void {
  checkTable(t);
  checkTableShape("Table.setModel", value, "value");
  tableApplyModel(t, { ...value, ref: normalizeReference(value.ref) });
}

/**
 * Insert a data row at the 0-based `rowNumber` — from `0` to the number of data rows; the end when absent. A row
 * shorter than the table is padded with empty cells.
 *
 * @throws {TableError} when `rowNumber` is not such a position, or the row has more values than the table has columns.
 */
export function addRow(t: TableData, values: readonly CellValueInput[], rowNumber?: number): void;
/** @deprecated `options.commit` does nothing — every edit applies as it is made. Omit it. Removed in 1.0. */
export function addRow(
  t: TableData,
  values: readonly CellValueInput[],
  rowNumber: number | undefined,
  options: TableRowEditOptions
): void;
export function addRow(t: TableData, values: readonly CellValueInput[], rowNumber?: number): void {
  checkTable(t);
  tableAddRow(t, values as readonly CellValue[], rowNumber);
}

/**
 * Remove `count` (default 1) data rows from the 0-based `rowIndex`, or as many as there are from it to the end.
 *
 * @throws {TableError} when `rowIndex` is not a data row of the table or `count` is not a whole number of at least 0.
 */
export function removeRows(t: TableData, rowIndex: number, count?: number): void;
/** @deprecated `options.commit` does nothing — every edit applies as it is made. Omit it. Removed in 1.0. */
export function removeRows(
  t: TableData,
  rowIndex: number,
  count: number | undefined,
  options: TableRowEditOptions
): void;
export function removeRows(t: TableData, rowIndex: number, count?: number): void {
  checkTable(t);
  tableRemoveRows(t, rowIndex, count);
}

/**
 * Insert a column at the 0-based `colIndex` — from `0` to the number of columns; the end when absent — with its
 * data-row `values` (padded with empty cells when shorter).
 *
 * @throws {TableError} when `colIndex` is not such a position, there are more values than data rows, or the name is
 * already a column's.
 */
export function addColumn(
  t: TableData,
  column: TableColumnProperties,
  values: readonly CellValueInput[],
  colIndex?: number
): void {
  checkTable(t);
  checkColumn("Table.addColumn", "column", column);
  checkColumnNamesUnique("Table.addColumn", [
    ...t.table.columns.map(existing => existing.name),
    column.name
  ]);
  tableAddColumn(t, column, values as readonly CellValue[], colIndex);
}

/**
 * Remove `count` (default 1) columns from the 0-based `colIndex`, or as many as there are from it to the end. A table
 * keeps at least one column.
 *
 * @throws {TableError} when `colIndex` is not a column of the table, `count` is not a whole number of at least 0, or
 * no column would be left.
 */
export function removeColumns(t: TableData, colIndex: number, count?: number): void {
  checkTable(t);
  tableRemoveColumns(t, colIndex, count);
}

/**
 * Move the table so its top-left cell is `value` (`"B2"`, or `"b2"`; a range's top-left cell is taken). The cells move
 * with it.
 *
 * @throws {InvalidAddressError} when `value` is not a cell or range of the table's sheet (a sheet name is refused).
 * @throws {TableError} when the table would overlap another one.
 */
export function setRef(t: TableData, value: string): void {
  checkTable(t);
  const check = argumentChecks("Table.setRef");
  check.required("value", value);
  check.string("value", value);
  check.reference("value", value, { range: true });
  tableSetRef(t, normalizeReference(value));
}

/**
 * Rename the table. The totals row's formulas follow.
 *
 * @throws {TableError} when the name is not one Excel accepts, or another table or a defined name already has it.
 */
export function setName(t: TableData, value: string): void {
  checkTable(t);
  tableSetName(t, value);
}

/**
 * Set the name Excel shows and structured references use.
 *
 * @throws {TableError} when the name is not one Excel accepts, or another table or a defined name already has it.
 */
export function setDisplayName(t: TableData, value: string): void {
  checkTable(t);
  tableSetDisplayName(t, value);
}

function flag(where: string, value: unknown): void {
  argumentChecks(where).boolean("value", value);
}

/** Show (`true`) or hide (`false`) the header row; `undefined` clears the setting (Excel's default: shown). */
export function setHeaderRow(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  flag("Table.setHeaderRow", value);
  tableSetHeaderRow(t, value);
}

/** Show (`true`) or hide (`false`) the totals row; `undefined` clears the setting (Excel's default: hidden). */
export function setTotalsRow(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  flag("Table.setTotalsRow", value);
  tableSetTotalsRow(t, value);
}

/** Set the table style by name (`"TableStyleMedium2"`, or a custom style's name); `undefined` clears it. */
export function setTheme(t: TableData, value: string | undefined): void {
  checkTable(t);
  argumentChecks("Table.setTheme").string("value", value);
  tableSetTheme(t, value);
}

/** Set whether the first column is styled apart; `undefined` clears the setting. */
export function setShowFirstColumn(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  flag("Table.setShowFirstColumn", value);
  tableSetShowFirstColumn(t, value);
}

/** Set whether the last column is styled apart; `undefined` clears the setting. */
export function setShowLastColumn(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  flag("Table.setShowLastColumn", value);
  tableSetShowLastColumn(t, value);
}

/** Set whether rows are banded; `undefined` clears the setting. */
export function setShowRowStripes(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  flag("Table.setShowRowStripes", value);
  tableSetShowRowStripes(t, value);
}

/** Set whether columns are banded; `undefined` clears the setting. */
export function setShowColumnStripes(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  flag("Table.setShowColumnStripes", value);
  tableSetShowColumnStripes(t, value);
}

/**
 * The table's range — header, data rows and totals row (`"B2:D9"`) — however the table was made. The deprecated
 * `Table.ref` returns the reference as it was given, which for a table built in memory is its top-left cell.
 */
export function getRef(t: TableData): string {
  checkTable(t);
  return t.table.tableRef ?? t.table.ref;
}

/** The table's name. */
export function getName(t: TableData): string {
  checkTable(t);
  return tableName(t);
}

/** The name Excel shows and structured references use: the display name, or the name when there is none. */
export function getDisplayName(t: TableData): string {
  checkTable(t);
  return tableDisplayName(t);
}

/** Whether the header row is shown; `undefined` when the table carries no setting (Excel's default: shown). */
export function hasHeaderRow(t: TableData): boolean | undefined {
  checkTable(t);
  return tableHeaderRow(t);
}

/** Whether the totals row is shown; `undefined` when the table carries no setting (Excel's default: hidden). */
export function hasTotalsRow(t: TableData): boolean | undefined {
  checkTable(t);
  return tableTotalsRow(t);
}

/** The table style's name, or `undefined`. */
export function getTheme(t: TableData): string | undefined {
  checkTable(t);
  return tableTheme(t);
}

/** Whether the first column is styled apart; `undefined` when the table carries no setting. */
export function hasFirstColumnStyle(t: TableData): boolean | undefined {
  checkTable(t);
  return tableShowFirstColumn(t);
}

/** Whether the last column is styled apart; `undefined` when the table carries no setting. */
export function hasLastColumnStyle(t: TableData): boolean | undefined {
  checkTable(t);
  return tableShowLastColumn(t);
}

/** Whether rows are banded; `undefined` when the table carries no setting. */
export function hasRowStripes(t: TableData): boolean | undefined {
  checkTable(t);
  return tableShowRowStripes(t);
}

/** Whether columns are banded; `undefined` when the table carries no setting. */
export function hasColumnStripes(t: TableData): boolean | undefined {
  checkTable(t);
  return tableShowColumnStripes(t);
}

/** The column's name. */
export function getColumnName(view: TableColumnView): string {
  checkTableColumn(view);
  return tableColumnName(view);
}

/**
 * Rename the column; its header cell follows.
 *
 * @throws {TableError} when another column of the table already has the name (compared as Excel compares names).
 */
export function setColumnName(view: TableColumnView, value: string): void {
  checkTableColumn(view);
  const check = argumentChecks("Table.setColumnName");
  check.required("value", value);
  check.string("value", value);
  checkColumnNamesUnique("Table.setColumnName", [
    ...view.table.table.columns.filter(column => column !== view.column).map(column => column.name),
    value
  ]);
  tableColumnSetName(view, value);
}

/** Whether the column shows a filter button; `undefined` when the column carries no setting (Excel's default: shown). */
export function hasColumnFilterButton(view: TableColumnView): boolean | undefined {
  checkTableColumn(view);
  return tableColumnFilterButton(view);
}

/** Show (`true`) or hide (`false`) the column's filter button; `undefined` clears the setting. */
export function setColumnFilterButton(view: TableColumnView, value: boolean | undefined): void {
  checkTableColumn(view);
  flag("Table.setColumnFilterButton", value);
  tableColumnSetFilterButton(view, value);
}

/** The text the totals row shows under the column, or `undefined`. */
export function getColumnTotalsRowLabel(view: TableColumnView): string | undefined {
  checkTableColumn(view);
  return tableColumnTotalsRowLabel(view);
}

/** Set the text the totals row shows under the column; `undefined` clears it. */
export function setColumnTotalsRowLabel(view: TableColumnView, value: string | undefined): void {
  checkTableColumn(view);
  argumentChecks("Table.setColumnTotalsRowLabel").string("value", value);
  tableColumnSetTotalsRowLabel(view, value);
}

/** The function the totals row computes for the column, or `undefined`. */
export function getColumnTotalsRowFunction(
  view: TableColumnView
): TableColumnProperties["totalsRowFunction"] {
  checkTableColumn(view);
  return tableColumnTotalsRowFunction(view);
}

/** Set the function the totals row computes for the column (`"custom"` uses the totals-row formula). */
export function setColumnTotalsRowFunction(
  view: TableColumnView,
  value: TableColumnProperties["totalsRowFunction"]
): void {
  checkTableColumn(view);
  argumentChecks("Table.setColumnTotalsRowFunction").oneOf("value", value, TOTALS_FUNCTIONS);
  tableColumnSetTotalsRowFunction(view, value);
}

/** The totals cell's cached result, or `undefined`. */
export function getColumnTotalsRowResult(view: TableColumnView): CellValue {
  checkTableColumn(view);
  return tableColumnTotalsRowResult(view);
}

/** Set the totals cell's cached result — what Excel shows until it recalculates; `undefined` clears it. */
export function setColumnTotalsRowResult(
  view: TableColumnView,
  value: TableColumnProperties["totalsRowResult"]
): void {
  checkTableColumn(view);
  checkTotalsResult(argumentChecks("Table.setColumnTotalsRowResult"), "value", value);
  tableColumnSetTotalsRowResult(view, value);
}

/** The totals cell's formula, or `undefined`. */
export function getColumnTotalsRowFormula(view: TableColumnView): string | undefined {
  checkTableColumn(view);
  return tableColumnTotalsRowFormula(view);
}

/** Set the formula a `"custom"` totals cell computes; `undefined` clears it. */
export function setColumnTotalsRowFormula(view: TableColumnView, value: string | undefined): void {
  checkTableColumn(view);
  argumentChecks("Table.setColumnTotalsRowFormula").string("value", value);
  tableColumnSetTotalsRowFormula(view, value);
}

export { tableGetModel as getModel, tableGetColumnChecked as getColumn };

/** A table handle. */
export type { TableData as Handle } from "@excel/core/table";

/**
 * A copy of a table column's style. Change the column with {@link setColumnStyle}, which applies the new style to
 * every cell of the column and takes back what the previous one gave them.
 */
export function getColumnStyle(view: TableColumnView): Partial<Style> | undefined {
  checkTableColumn(view);
  return copyStyle(tableColumnStyle(view)) as Partial<Style> | undefined;
}

/**
 * The column's own style object: mutating it changes the column, and the table's next edit applies it to every cell
 * of the column.
 */
function liveColumnStyle(view: TableColumnView): Partial<Style> | undefined {
  const style = tableColumnStyle(view);
  if (style) {
    // Cells the column has already written point at a snapshot of these facets, so a mutation does not reach them;
    // the snapshot is dropped so the next edit applies what the column says then — see `core/style-sharing.ts`.
    invalidateSharedCellStyle(style);
    tableMarkLiveStyle(view);
  }
  return style;
}

// --- deprecated names ---

/**
 * @deprecated Use {@link getModel}, which returns a copy; change the table with {@link setModel} or the setters. This
 * returns the model the table uses. Removed in 1.0.
 */
export const model = tableModel;
/**
 * @deprecated Use {@link getColumn}, which refuses an index the table has no column at. Removed in 1.0.
 */
export const column = tableGetColumn;
/** @deprecated Use {@link getRef}. Removed in 1.0. */
export const ref = tableRef;
/** @deprecated Use {@link getName}. Removed in 1.0. */
export const name = getName;
/** @deprecated Use {@link getDisplayName}. Removed in 1.0. */
export const displayName = getDisplayName;
/** @deprecated Use {@link hasHeaderRow}. Removed in 1.0. */
export const headerRow = hasHeaderRow;
/** @deprecated Use {@link hasTotalsRow}. Removed in 1.0. */
export const totalsRow = hasTotalsRow;
/** @deprecated Use {@link getTheme}. Removed in 1.0. */
export const theme = getTheme;
/** @deprecated Use {@link hasFirstColumnStyle}. Removed in 1.0. */
export const showFirstColumn = hasFirstColumnStyle;
/** @deprecated Use {@link hasLastColumnStyle}. Removed in 1.0. */
export const showLastColumn = hasLastColumnStyle;
/** @deprecated Use {@link hasRowStripes}. Removed in 1.0. */
export const showRowStripes = hasRowStripes;
/** @deprecated Use {@link hasColumnStripes}. Removed in 1.0. */
export const showColumnStripes = hasColumnStripes;
/** @deprecated Use {@link getColumnName}. Removed in 1.0. */
export const columnName = getColumnName;
/** @deprecated Use {@link setColumnName}. Removed in 1.0. */
export const columnSetName = setColumnName;
/** @deprecated Use {@link hasColumnFilterButton}. Removed in 1.0. */
export const columnFilterButton = hasColumnFilterButton;
/** @deprecated Use {@link setColumnFilterButton}. Removed in 1.0. */
export const columnSetFilterButton = setColumnFilterButton;
/** @deprecated Use {@link setColumnStyle}. Removed in 1.0. */
export const columnSetStyle = setColumnStyle;
/** @deprecated Use {@link getColumnTotalsRowLabel}. Removed in 1.0. */
export const columnTotalsRowLabel = getColumnTotalsRowLabel;
/** @deprecated Use {@link setColumnTotalsRowLabel}. Removed in 1.0. */
export const columnSetTotalsRowLabel = setColumnTotalsRowLabel;
/** @deprecated Use {@link getColumnTotalsRowFunction}. Removed in 1.0. */
export const columnTotalsRowFunction = getColumnTotalsRowFunction;
/** @deprecated Use {@link setColumnTotalsRowFunction}. Removed in 1.0. */
export const columnSetTotalsRowFunction = setColumnTotalsRowFunction;
/** @deprecated Use {@link getColumnTotalsRowResult}. Removed in 1.0. */
export const columnTotalsRowResult = getColumnTotalsRowResult;
/** @deprecated Use {@link setColumnTotalsRowResult}. Removed in 1.0. */
export const columnSetTotalsRowResult = setColumnTotalsRowResult;
/** @deprecated Use {@link getColumnTotalsRowFormula}. Removed in 1.0. */
export const columnTotalsRowFormula = getColumnTotalsRowFormula;
/** @deprecated Use {@link setColumnTotalsRowFormula}. Removed in 1.0. */
export const columnSetTotalsRowFormula = setColumnTotalsRowFormula;
/**
 * @deprecated Use {@link getColumnStyle}, which returns a copy, and {@link setColumnStyle} to change it. This returns
 * the column's own style object. Removed in 1.0.
 */
export const columnStyle = liveColumnStyle;
