import type { CellData } from "@excel/core/cell";
import {
  cellFormula,
  cellGetValue,
  cellOwnStyle,
  cellResult,
  cellSetValue,
  checkCellValue,
  setFacetShared
} from "@excel/core/cell";
import { definedNamesHasName } from "@excel/core/defined-names";
import { checkTable, checkTableColumn } from "@excel/core/feature-handles";
import { sharedCellFacet } from "@excel/core/style-sharing";
import { sanitizeTableName, tableNameProblem } from "@excel/core/table-name";
import {
  checkWorksheetAttached,
  findCell,
  getRow,
  getSheetWorkbook,
  rowGetCell
} from "@excel/core/worksheet-core";
import type { WorksheetData as Worksheet } from "@excel/core/worksheet-core";
import { ExcelArgumentError, TableError } from "@excel/errors";
import type {
  DecodedAddress,
  CellFormulaValue,
  CellValue,
  Style,
  TableColumnProperties,
  TableStyleProperties
} from "@excel/types";
import { colCache } from "@excel/utils/col-cache";
import { copyStyle } from "@excel/utils/copy-style";
import type { LineEdit, ReferenceEdit } from "@excel/utils/reference-rewrite";
import { editLine, editSpan, rewriteFormula } from "@excel/utils/reference-rewrite";
import { excelNameKey } from "@utils/name-key";

interface LoadedTableColumnProperties extends TableColumnProperties {
  rawFilterXml?: string[];
  namespaceAttributes?: Record<string, string>;
}

/** The round-trip table model — `Table.getModel` / `Table.setModel`. */
export interface TableModel {
  ref: string;
  name: string;
  displayName?: string;
  columns: LoadedTableColumnProperties[];
  rows: CellValue[][];
  headerRow?: boolean;
  totalsRow?: boolean;
  qualifyImplicitStructuredReferences?: boolean;
  style?: TableStyleProperties;
  tl?: DecodedAddress;
  autoFilterRef?: string;
  tableRef?: string;
  sortStateXml?: string;
  sortStateAutoFilterRef?: string;
  autoFilterSortStateXml?: string;
  autoFilterSortStateRef?: string;
  autoFilterExtLstXml?: string;
  rawAttributes?: Record<string, string>;
  extLstXml?: string;
  autoFilterNamespaceAttributes?: Record<string, string>;
}

function clearStaleSortState(t: TableData): void {
  t.table.sortStateXml = undefined;
  t.table.sortStateAutoFilterRef = undefined;
  t.table.autoFilterSortStateXml = undefined;
  t.table.autoFilterSortStateRef = undefined;
}

// ============================================================================
// Table — de-classed domain model (data record + flat helpers)
// ============================================================================

/**
 * Plain-data worksheet table (de-classed domain model). Holds the owning
 * worksheet and the {@link TableModel}. All former methods/getters/setters are
 * flat `table*` helpers.
 */
export interface TableData {
  worksheet: Worksheet;
  table: TableModel;
}

/**
 * View over a single table column (de-classed). Carries the owning table, the
 * column properties record, and the column index; its setters apply to the sheet
 * as they are made.
 */
export interface TableColumnView {
  table: TableData;
  column: TableColumnProperties;
  index: number;
}

// SUBTOTAL function codes per OOXML/Excel (see original notes):
//   average → 101, countNums → 102, count → 103, max → 104, min → 105,
//   stdDev → 107, var → 110, sum → 109.
const SUBTOTAL_FUNCTIONS: Record<string, number> = {
  average: 101,
  countNums: 102,
  count: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  var: 110,
  sum: 109
};

/**
 * Carry the totals row's cached results from the loaded sheet into a table model — **for a reader only**.
 *
 * The results are `<v>` values in the sheet's cells; the table part has nowhere to keep them, and
 * {@link createTable} rewrites the totals cells from the model. A reader that does not call this loses
 * every cached total on load.
 *
 * It belongs to the load path and nowhere else. A cached value is evidence about the data the *file*
 * held; a table a caller builds with `Table.add` brings its own rows, and an old result sitting in the
 * same cell — even under an identical formula — says nothing about them.
 *
 * `range` is the table's full extent (header and totals included). Only a result cached under the exact
 * formula this table writes is taken, and a result the model already carries wins.
 */
export function adoptLoadedTotals(worksheet: Worksheet, table: TableModel, range: string): void {
  if (table.totalsRow !== true) {
    return;
  }
  const decoded = colCache.decode(range);
  if (!("dimensions" in decoded)) {
    return;
  }
  const t: TableData = { worksheet, table };
  table.columns.forEach((column, j) => {
    if (j === 0 || column.totalsRowResult !== undefined) {
      return;
    }
    const cell = findCell(worksheet, decoded.bottom, decoded.left + j);
    if (!cell) {
      return;
    }
    const result = cellResult(cell);
    if (result !== undefined && cellFormula(cell) === tableGetFormula(t, column)) {
      column.totalsRowResult = result;
    }
  });
}

/** Create a table bound to a worksheet, validating + storing on-sheet if a model is given. */
export function createTable(worksheet: Worksheet, table?: TableModel): TableData {
  const t: TableData = { worksheet, table: table as TableModel };
  if (table) {
    // When loading tables from xlsx, Excel stores table ranges and cell values in the worksheet,
    // but may not embed row data into the table definition. Hydrate rows from the worksheet so
    // table mutations (e.g. addRow) can correctly expand table ranges and serialize.
    if (Array.isArray(table.rows) && table.rows.length === 0 && table.tableRef) {
      const decoded = colCache.decode(table.tableRef);
      if ("dimensions" in decoded) {
        const startRow = decoded.top + (table.headerRow === false ? 0 : 1);
        const endRow = decoded.bottom - (table.totalsRow === true ? 1 : 0);

        if (endRow >= startRow) {
          for (let r = startRow; r <= endRow; r++) {
            const row = getRow(worksheet, r);
            const values: CellValue[] = [];
            for (let c = decoded.left; c <= decoded.right; c++) {
              values.push(cellGetValue(rowGetCell(row, c)));
            }
            table.rows.push(values);
          }
        }
      }
    }
    // check things are ok first
    tableValidate(t);

    tableStore(t);
  }
  return t;
}

/**
 * Build a table a caller asked for, without writing anything: the model is validated, its values are checked as a
 * cell's are, and its area is checked against the sheet's other tables. {@link tableStore} then writes it. Splitting
 * the two is what lets `Table.add` refuse a duplicate name or an overlap before a single cell is overwritten.
 */
export function tablePrepare(worksheet: Worksheet, table: TableModel): TableData {
  const t: TableData = { worksheet, table };
  tableValidate(t);
  checkTableValues(table, "Table.add", "model");
  checkNoOverlap(t, layoutOf(table));
  return t;
}

/** Tables taken off their sheet — by `Table.remove`, or by a sheet edit that deleted all of them. */
const removedTables = new WeakSet<TableData>();

/** Record that `t` is no longer on its sheet, so its handle refuses edits from now on. */
export function markTableRemoved(t: TableData): void {
  checkTable(t);
  removedTables.add(t);
}

/**
 * Refuse an edit through a handle whose table was removed, or whose sheet was. A removed table's handle used to keep
 * editing the sheet: `setRef` moved cells the table no longer owned.
 */
export function checkTableEditable(t: TableData): void {
  checkTable(t);
  checkWorksheetAttached(t.worksheet);
  if (removedTables.has(t)) {
    throw new ExcelArgumentError(
      "table",
      `Table "${t.table.name}" was removed from its sheet; it can no longer be edited.`
    );
  }
}

/** Refuse a value a cell cannot hold, in any data row or cached total of `table`, before any cell is written. */
function checkTableValues(table: TableModel, where: string, argument: string): void {
  table.rows.forEach((row, r) =>
    row.forEach((value, c) =>
      checkCellValue(value, { where, field: `${argument}.rows[${r}][${c}]` })
    )
  );
  table.columns.forEach((column, c) =>
    checkCellValue(column.totalsRowResult, {
      where,
      field: `${argument}.columns[${c}].totalsRowResult`
    })
  );
}

/**
 * Refuse a table area that intersects another table of the sheet. Excel cannot hold two tables over one cell — the
 * file opens behind a repair prompt — and writing the second table's cells overwrote the first one's data.
 */
function checkNoOverlap(t: TableData, layout: Layout, self: TableData = t): void {
  const tables = t.worksheet?.tables;
  if (!tables) {
    return;
  }
  const last = lastRowOf(layout);
  for (const other of Object.values(tables)) {
    if (other === self || other.table === t.table) {
      continue;
    }
    const o = layoutOf(other.table);
    if (
      layout.col < o.col + o.width &&
      o.col < layout.col + layout.width &&
      layout.row <= lastRowOf(o) &&
      o.row <= last
    ) {
      throw new TableError(
        `Table "${t.table.name}" would cover ${colCache.encode(layout.row, layout.col, last, layout.col + layout.width - 1)}, ` +
          `which overlaps table "${other.table.name}" (${other.table.tableRef ?? other.table.ref}); tables cannot share cells.`
      );
    }
  }
}

export function tableGetFormula(t: TableData, column: TableColumnProperties): string | null {
  checkTable(t);
  if (column.totalsRowFunction === "none") {
    return null;
  }
  if (column.totalsRowFunction === "custom") {
    return column.totalsRowFormula ?? null;
  }
  const fnNum = column.totalsRowFunction ? SUBTOTAL_FUNCTIONS[column.totalsRowFunction] : undefined;
  if (fnNum !== undefined) {
    return `SUBTOTAL(${fnNum},${t.table.name}[${column.name}])`;
  }
  throw new TableError(`Invalid Totals Row Function: ${column.totalsRowFunction}`);
}

function hasTotalsFunction(column: TableColumnProperties): boolean {
  return column.totalsRowFunction !== undefined && column.totalsRowFunction !== "none";
}

/** Number of columns. */
export function tableWidth(t: TableData): number {
  checkTable(t);
  return t.table.columns.length;
}

/** Number of data rows. */
export function tableHeight(t: TableData): number {
  checkTable(t);
  return t.table.rows.length;
}

/** Data rows + optional header row. */
export function tableFilterHeight(t: TableData): number {
  checkTable(t);
  return tableHeight(t) + (t.table.headerRow ? 1 : 0);
}

/** Full on-sheet height (data + header + optional totals row). */
export function tableTableHeight(t: TableData): number {
  checkTable(t);
  return tableFilterHeight(t) + (t.table.totalsRow ? 1 : 0);
}

export function tableValidate(t: TableData): void {
  checkTable(t);
  const { table } = t;
  // set defaults and check is valid
  const assign = <T extends object, K extends keyof T>(o: T, name: K, dflt: T[K]): void => {
    if (o[name] === undefined) {
      o[name] = dflt;
    }
  };
  assign(table, "headerRow", true);
  assign(table, "totalsRow", false);

  assign(table, "style", {});
  const style = table.style!;
  assign(style, "theme", "TableStyleMedium2");
  assign(style, "showFirstColumn", false);
  assign(style, "showLastColumn", false);
  assign(style, "showRowStripes", false);
  assign(style, "showColumnStripes", false);

  // A name Excel would not accept is corrected here, for a table read from a file; a caller's name reaches this
  // only after `checkTableName` has refused one, so it is never changed in silence.
  if (table.name) {
    table.name = sanitizeTableName(table.name);
  }
  if (table.displayName) {
    table.displayName = sanitizeTableName(table.displayName);
  }

  const assert = (test: boolean, message: string) => {
    if (!test) {
      throw new TableError(message);
    }
  };
  assert(!!table.name, "Table must have a name");
  assert(!!table.ref, "Table must have ref");
  assert(!!table.columns, "Table must have column definitions");
  assert(table.columns.length > 0, "Table must have at least one column");
  assert(!!table.rows, "Table must have row definitions");

  table.tl = colCache.decodeAddress(table.ref);
  const { row, col } = table.tl;
  assert(row > 0, "Table must be on valid row");
  assert(col > 0, "Table must be on valid col");

  const width = tableWidth(t);
  const filterHeight = tableFilterHeight(t);
  const tableHt = tableTableHeight(t);

  // autoFilterRef spans the header + all data rows (excludes the optional
  // totals row). See original notes: matches what Excel emits and avoids
  // "Removed Records: Table" on open. A table without a header row has no
  // filter — its buttons would have nowhere to sit, and Excel refuses a
  // header-less table that carries an <autoFilter>.
  table.autoFilterRef = table.headerRow
    ? colCache.encode(row, col, row + filterHeight - 1, col + width - 1)
    : undefined;

  // tableRef is a range that includes optional headers and totals
  table.tableRef = colCache.encode(row, col, row + tableHt - 1, col + width - 1);

  // Excel refuses a table whose column names repeat, compared as `excelNameKey` does, and numbers a repeated header itself when
  // one is typed: `a`, `a` becomes `a`, `a2`. Named the same way here, so the header cells say the names the part does.
  // The first column with a name keeps it; a later one takes the first number no column uses.
  table.columns.forEach((column, i) => {
    assert(!!column.name, `Column ${i} must have a name`);
  });
  const taken = new Set(table.columns.map(column => excelNameKey(column.name)));
  const kept = new Set<string>();
  for (const column of table.columns) {
    const key = excelNameKey(column.name);
    if (!kept.has(key)) {
      kept.add(key);
      continue;
    }
    let n = 2;
    while (taken.has(excelNameKey(`${column.name}${n}`))) {
      n++;
    }
    column.name = `${column.name}${n}`;
    taken.add(excelNameKey(column.name));
  }

  // A totals cell shows a function or a label, never both. Excel labels the first column "Total" when the row is switched
  // on, and drops that label when a function is chosen for the column — so the default applies only where none is.
  table.columns.forEach((column, i) => {
    if (i === 0 && !hasTotalsFunction(column)) {
      assign(column, "totalsRowLabel", "Total");
    } else {
      assign(column, "totalsRowFunction", "none");
      column.totalsRowFormula = tableGetFormula(t, column) ?? undefined;
    }
  });
}

export function tableStore(t: TableData): void {
  checkTable(t);
  // where the table needs to store table data, headers, footers in the sheet...
  const layout = layoutOf(t.table);
  for (let j = 0; j < layout.width; j++) {
    const column = t.table.columns[j];
    if (layout.headerRow) {
      const cell = cellAt(t, headerRowOf(layout), layout.col + j);
      cellSetValue(cell, column.name);
      assignColumnStyle(cell, column.style);
    }
    t.table.rows.forEach((data, i) => {
      if (j < data.length) {
        const cell = cellAt(t, bodyRowOf(layout, i), layout.col + j);
        cellSetValue(cell, storedValue(t, data[j]));
        assignColumnStyle(cell, column.style);
      }
    });
    if (layout.totalsRow) {
      const cell = cellAt(t, totalsRowOf(layout), layout.col + j);
      cellSetValue(cell, totalsValue(t, column));
      assignColumnStyle(cell, column.style);
    }
  }
}

/**
 * A table column's style is one object applied to its header, every body row and
 * its totals cell, so it is a style owner in the sense of `style-sharing.ts`: the
 * cells share one frozen snapshot of it rather than each holding a copy, and they
 * are marked so that `Cell.get*` separates them again.
 *
 * This used to be `Object.assign(cell.style, style)`, which shared the facets *and*
 * said nothing about it, so every cell in a table column was permanently aliased —
 * `Cell.getStyle(a).font.bold = false` rewrote the whole column.
 */
function assignColumnStyle(cell: CellData, style: Partial<Style> | undefined): void {
  if (style) {
    for (const key of Object.keys(style) as (keyof Style)[]) {
      setFacetShared(cell, key, sharedCellFacet(style, key));
    }
  }
}

/** A body value as the sheet stores it: an implicit `[@Col]` qualified with the table's name when the model asks. */
function storedValue(t: TableData, value: CellValue): CellValue {
  const isFormulaValue = typeof value === "object" && value !== null && "formula" in value;
  if (!isFormulaValue || typeof (value as CellFormulaValue).formula !== "string") {
    return value;
  }
  const formulaValue = value as CellFormulaValue;
  if (t.table.qualifyImplicitStructuredReferences !== true) {
    return formulaValue;
  }
  return {
    ...formulaValue,
    formula: formulaValue.formula.replace(
      /(^|[^A-Za-z0-9_])\[@\[?([^[\]]+?)\]?\]/g,
      `$1${t.table.name}[[#This Row],[$2]]`
    )
  } as CellFormulaValue;
}

/** What a totals cell shows: the column's function as a formula, or its label. */
function totalsValue(t: TableData, column: TableColumnProperties): CellValue {
  const formula = hasTotalsFunction(column) ? tableGetFormula(t, column) : null;
  if (formula) {
    return { formula, result: column.totalsRowResult } as CellValue;
  }
  return column.totalsRowLabel ?? null;
}

// ============================================================================
// Layout — where a table's cells are, before and after an edit
// ============================================================================

/**
 * Where a table's cells sit on the sheet.
 *
 * **The sheet's cells are what a table holds.** Its model keeps the shape — position, columns, how many data rows —
 * and an edit moves cells from where one layout puts them to where the next one does, carrying each cell's value and
 * its own formatting. The model's `rows` are input when a table is built and output when it is read; nothing is ever
 * stored back from them, so there is no second copy of the data to fall out of step with the cells.
 */
interface Layout {
  row: number;
  col: number;
  headerRow: boolean;
  totalsRow: boolean;
  width: number;
  /** Data rows. */
  height: number;
}

function layoutOf(table: TableModel): Layout {
  const { row, col } = colCache.decodeAddress(table.ref);
  return {
    row,
    col,
    headerRow: table.headerRow !== false,
    totalsRow: table.totalsRow === true,
    width: table.columns.length,
    height: table.rows.length
  };
}

function headerRowOf(layout: Layout): number {
  return layout.row;
}

function bodyRowOf(layout: Layout, index: number): number {
  return layout.row + (layout.headerRow ? 1 : 0) + index;
}

function totalsRowOf(layout: Layout): number {
  return bodyRowOf(layout, layout.height);
}

function lastRowOf(layout: Layout): number {
  return bodyRowOf(layout, layout.height) - (layout.totalsRow ? 0 : 1);
}

function contains(layout: Layout, row: number, col: number): boolean {
  return (
    col >= layout.col &&
    col < layout.col + layout.width &&
    row >= layout.row &&
    row <= lastRowOf(layout)
  );
}

function cellAt(t: TableData, row: number, col: number): CellData {
  return rowGetCell(getRow(t.worksheet, row), col);
}

/**
 * What an edit does to the table's rows and columns: for each data row and column of the result, the one it came from
 * before the edit — or, for one the edit adds, its values. Rows and columns default to staying where they were.
 */
interface EditPlan {
  /** The earlier data row a data row of the result came from, or the values of a row the edit adds. */
  rowSource?: (index: number) => number | readonly CellValue[];
  /** The earlier column a column of the result came from, or `undefined` for a column the edit adds. */
  columnSource?: (index: number) => number | undefined;
  /** The data-row values of an added column. */
  addedColumnValues?: readonly CellValue[];
  /** Columns whose style the edit replaced, with the style they had. */
  restyled?: ReadonlyMap<TableColumnProperties, Partial<Style> | undefined>;
}

interface Moved {
  row: number;
  col: number;
  value: CellValue;
  style: Partial<Style> | undefined;
}

/**
 * Change the table's model and move its cells to match.
 *
 * `edit` changes the model; the result is validated **before anything on the sheet changes**, and a failed validation
 * restores the model as it was — so an edit that throws leaves the table exactly as usable as before the call.
 */
function editTable(t: TableData, edit: () => void, plan: EditPlan = {}): void {
  checkTableEditable(t);
  const before = layoutOf(t.table);
  const saved = saveModel(t.table);
  try {
    edit();
    tableValidate(t);
    checkNoOverlap(t, layoutOf(t.table));
  } catch (error) {
    restoreModel(t.table, saved);
    throw error;
  }
  relayout(t, before, layoutOf(t.table), plan);
  applyLiveStyles(t);
}

function relayout(t: TableData, before: Layout, after: Layout, plan: EditPlan): void {
  const { worksheet } = t;
  const columns = t.table.columns;
  const rowSource = plan.rowSource ?? ((i: number) => i);
  const columnSource = plan.columnSource ?? ((j: number) => j);
  const identityColumns =
    before.col === after.col &&
    before.width === after.width &&
    Array.from({ length: after.width }, (_, j) => columnSource(j) === j).every(Boolean);

  // Read every cell that moves before writing any: a destination is often another move's source.
  const moved: Moved[] = [];
  const written: { row: number; col: number; value: CellValue; column: TableColumnProperties }[] =
    [];
  const take = (row: number, col: number, fromRow: number, fromCol: number) => {
    if (row === fromRow && col === fromCol) {
      return;
    }
    const source = findCell(worksheet, fromRow, fromCol);
    moved.push({
      row,
      col,
      value: source ? cellGetValue(source) : null,
      style: source?.style
    });
  };

  for (let i = 0; i < after.height; i++) {
    const from = rowSource(i);
    const row = bodyRowOf(after, i);
    if (typeof from === "number" && identityColumns && bodyRowOf(before, from) === row) {
      continue;
    }
    for (let j = 0; j < after.width; j++) {
      const fromColumn = columnSource(j);
      if (typeof from !== "number") {
        written.push({ row, col: after.col + j, value: from[j] ?? null, column: columns[j] });
      } else if (fromColumn === undefined) {
        written.push({
          row,
          col: after.col + j,
          value: plan.addedColumnValues?.[i] ?? null,
          column: columns[j]
        });
      } else {
        take(row, after.col + j, bodyRowOf(before, from), before.col + fromColumn);
      }
    }
  }
  // The header and totals rows are written from the model; only their formatting moves.
  const edgeStyles: {
    row: number;
    col: number;
    from?: [number, number];
    column: TableColumnProperties;
  }[] = [];
  for (let j = 0; j < after.width; j++) {
    const fromColumn = columnSource(j);
    const col = after.col + j;
    if (after.headerRow) {
      edgeStyles.push({
        row: headerRowOf(after),
        col,
        from:
          before.headerRow && fromColumn !== undefined
            ? [headerRowOf(before), before.col + fromColumn]
            : undefined,
        column: columns[j]
      });
    }
    if (after.totalsRow) {
      edgeStyles.push({
        row: totalsRowOf(after),
        col,
        from:
          before.totalsRow && fromColumn !== undefined
            ? [totalsRowOf(before), before.col + fromColumn]
            : undefined,
        column: columns[j]
      });
    }
  }
  const edgeSources = edgeStyles.map(({ row, col, from }) =>
    from && (from[0] !== row || from[1] !== col)
      ? findCell(worksheet, from[0], from[1])?.style
      : undefined
  );

  // What the table no longer covers keeps its formatting and loses its value, as a sheet does when a table shrinks.
  const covers =
    after.col <= before.col &&
    after.col + after.width >= before.col + before.width &&
    after.row <= before.row &&
    lastRowOf(after) >= lastRowOf(before);
  for (let row = covers ? Infinity : before.row; row <= lastRowOf(before); row++) {
    for (let col = before.col; col < before.col + before.width; col++) {
      if (!contains(after, row, col) && findCell(worksheet, row, col)) {
        cellSetValue(cellAt(t, row, col), null);
      }
    }
  }

  for (const { row, col, value, style } of moved) {
    const cell = cellAt(t, row, col);
    cellSetValue(cell, value);
    cell.style = (copyStyle(style) as Partial<Style> | undefined) ?? {};
    cell._sharedStyle = false;
  }
  for (const { row, col, value, column } of written) {
    const cell = cellAt(t, row, col);
    cellSetValue(cell, storedValue(t, value));
    styleNewCell(cell, contains(before, row, col), column);
  }
  edgeStyles.forEach(({ row, col, from, column }, k) => {
    const cell = cellAt(t, row, col);
    if (from === undefined) {
      styleNewCell(cell, contains(before, row, col), column);
    } else if (from[0] !== row || from[1] !== col) {
      cell.style = (copyStyle(edgeSources[k]) as Partial<Style> | undefined) ?? {};
      cell._sharedStyle = false;
    }
  });
  for (let j = 0; j < after.width; j++) {
    if (after.headerRow) {
      cellSetValue(cellAt(t, headerRowOf(after), after.col + j), columns[j].name);
    }
    if (after.totalsRow) {
      cellSetValue(cellAt(t, totalsRowOf(after), after.col + j), totalsValue(t, columns[j]));
    }
  }

  plan.restyled?.forEach((previous, column) => restyleColumn(t, after, column, previous));
}

/**
 * Give a cell the table just started covering its column's style.
 *
 * A position that was already inside the table belonged to a cell the edit moved away, so that cell's formatting is not
 * this one's and it starts clean. A position the table grew onto keeps whatever the sheet gave it, with the column's
 * style over it.
 */
function styleNewCell(cell: CellData, wasInTable: boolean, column: TableColumnProperties): void {
  if (wasInTable) {
    cell.style = {};
    cell._sharedStyle = false;
  }
  assignColumnStyle(cell, column.style);
}

/**
 * Apply a column's new style to every cell of the column, taking back what its previous style gave them.
 *
 * A facet a cell holds that is equal to the previous column style's is the column's, and is replaced or removed; one
 * that differs is the cell's own formatting and stays — which is what lets clearing a column's style clear it without
 * clearing what a caller set on a cell.
 */
function restyleColumn(
  t: TableData,
  layout: Layout,
  column: TableColumnProperties,
  previous: Partial<Style> | undefined
): void {
  const j = t.table.columns.indexOf(column);
  if (j < 0) {
    return;
  }
  const col = layout.col + j;
  for (let row = layout.row; row <= lastRowOf(layout); row++) {
    const cell = cellAt(t, row, col);
    if (previous) {
      for (const key of Object.keys(previous) as (keyof Style)[]) {
        if (column.style?.[key] === undefined && sameFacet(cell.style[key], previous[key])) {
          cellOwnStyle(cell);
          delete cell.style[key];
        }
      }
    }
    assignColumnStyle(cell, column.style);
  }
}

function sameFacet(a: unknown, b: unknown): boolean {
  if (a === b) {
    return true;
  }
  if (typeof a !== "object" || typeof b !== "object" || a === null || b === null) {
    return false;
  }
  return JSON.stringify(a) === JSON.stringify(b);
}

// ============================================================================
// Model snapshots
// ============================================================================

interface SavedModel {
  fields: Record<string, unknown>;
  style?: TableStyleProperties;
  columns: LoadedTableColumnProperties[];
  columnFields: Record<string, unknown>[];
  rows: CellValue[][];
  rowValues: CellValue[][];
}

function saveModel(model: TableModel): SavedModel {
  return {
    fields: { ...model },
    style: model.style ? { ...model.style } : undefined,
    columns: model.columns.slice(),
    columnFields: model.columns.map(column => ({ ...column })),
    rows: model.rows.slice(),
    rowValues: model.rows.map(row => row.slice())
  };
}

/** Put the model back as {@link saveModel} found it — the same objects, so column views stay attached. */
function restoreModel(model: TableModel, saved: SavedModel): void {
  const record = model as unknown as Record<string, unknown>;
  for (const key of Object.keys(record)) {
    if (!(key in saved.fields)) {
      delete record[key];
    }
  }
  Object.assign(record, saved.fields);
  if (model.style && saved.style) {
    for (const key of Object.keys(model.style)) {
      delete (model.style as Record<string, unknown>)[key];
    }
    Object.assign(model.style, saved.style);
  }
  model.columns.splice(0, model.columns.length, ...saved.columns);
  saved.columns.forEach((column, i) => {
    const fields = column as unknown as Record<string, unknown>;
    for (const key of Object.keys(fields)) {
      if (!(key in saved.columnFields[i])) {
        delete fields[key];
      }
    }
    Object.assign(fields, saved.columnFields[i]);
  });
  model.rows.splice(0, model.rows.length, ...saved.rows);
  saved.rows.forEach((row, i) => row.splice(0, row.length, ...saved.rowValues[i]));
}

/** The data rows as the sheet holds them. */
function readRows(t: TableData): CellValue[][] {
  if (!t.worksheet || !t.table.ref || !Array.isArray(t.table.rows)) {
    return t.table.rows;
  }
  const layout = layoutOf(t.table);
  return t.table.rows.map((_, i) => {
    const values: CellValue[] = [];
    for (let j = 0; j < layout.width; j++) {
      const cell = findCell(t.worksheet, bodyRowOf(layout, i), layout.col + j);
      values.push(cell ? cellGetValue(cell) : null);
    }
    return values;
  });
}

/**
 * The table's model, its `rows` read from the sheet — the live object, which the table keeps using.
 *
 * Internal readers and the deprecated `Table.model` use this. A value edited through `Cell.setValue` is what the rows
 * report, and what a pivot table or a chart built from this table reads. The rows are fresh arrays each time, so the
 * ones a caller passed to `Table.add` are never rewritten.
 */
export function tableModel(t: TableData): TableModel {
  checkTable(t);
  t.table.rows = readRows(t);
  return t.table;
}

/**
 * A copy of the table's model, its `rows` read from the sheet.
 *
 * A copy, because a reader must not hand out what the table uses: editing the copy changes nothing until it is given
 * to {@link tableApplyModel}.
 */
export function tableGetModel(t: TableData): TableModel {
  checkTable(t);
  const rows = readRows(t);
  const copy = structuredClone({ ...t.table, rows: [] as CellValue[][] });
  copy.rows = structuredClone(rows);
  return copy;
}

/**
 * Replace the table with `value`, as `Table.add` would build it: the model is validated and its name checked against
 * the workbook before anything changes, then the old cells are cleared and the model's header, rows and totals are
 * written.
 *
 * It used to assign the model and touch nothing else, so the sheet kept the old cells, the workbook kept the old name —
 * a second table could then take the new one — and reading the table back reported the cells, not the model.
 */
export function tableApplyModel(t: TableData, value: TableModel): void {
  checkTable(t);
  checkTableEditable(t);
  checkTableName("Table.setModel", "value.name", value?.name);
  if (value?.displayName !== undefined) {
    checkTableName("Table.setModel", "value.displayName", value.displayName);
  }
  const candidate = structuredClone(value);
  const next: TableData = { worksheet: t.worksheet, table: candidate };
  tableValidate(next);
  // Everything that can refuse the model does so before the old cells are cleared: a value no cell can hold used to
  // be found while the new rows were being written, after the old ones were gone.
  checkTableValues(candidate, "Table.setModel", "value");
  checkNoOverlap(next, layoutOf(candidate), t);
  if (candidate.displayName !== undefined) {
    checkTableNameFree(t, candidate.displayName, "displayName");
  }
  const ws = t.worksheet;
  const workbook = ws ? getSheetWorkbook(ws) : undefined;
  const tables = ws?.tables;
  const registeredAs = tables ? Object.keys(tables).find(name => tables[name] === t) : undefined;
  const newKey = excelNameKey(candidate.name);
  if (registeredAs !== undefined && newKey !== excelNameKey(registeredAs)) {
    if (workbook?._tableNames.has(newKey)) {
      throw new TableError(
        `Table name "${candidate.name}" already exists in the workbook (case-insensitive).`
      );
    }
    if (workbook && definedNamesHasName(workbook._definedNames, candidate.name)) {
      throw new TableError(
        `Table name "${candidate.name}" is a defined name in the workbook (case-insensitive).`
      );
    }
    checkTableNameFree(t, candidate.name, "name");
  }
  const before = layoutOf(t.table);
  for (let row = before.row; row <= lastRowOf(before); row++) {
    for (let col = before.col; col < before.col + before.width; col++) {
      if (findCell(ws, row, col)) {
        cellSetValue(cellAt(t, row, col), null);
      }
    }
  }
  t.table = candidate;
  tableStore(t);
  if (tables && registeredAs !== undefined && registeredAs !== candidate.name) {
    delete tables[registeredAs];
    tables[candidate.name] = t;
    workbook?._tableNames.delete(excelNameKey(registeredAs));
    workbook?._tableNames.add(newKey);
  }
}

/**
 * Replace the underlying {@link TableModel} without touching the sheet or the workbook — what a reader does once the
 * cells it loaded are in place.
 */
export function tableSetModel(t: TableData, value: TableModel): void {
  checkTable(t);
  t.table = value;
}

// ============================================================================
// Structural edits of the sheet
// ============================================================================

/**
 * Follow a row or column edit of the table's sheet, made after the sheet's cells have already moved.
 *
 * Only the model changes: the cells are where the sheet edit put them. The table moves, grows and shrinks by the same
 * rules as every other reference (`reference-rewrite`) — so lines inserted strictly inside it become its own rows or
 * columns, including rows inserted between the last data row and the totals row, while lines inserted at its first
 * line or just past its last one leave it the same size. A column the edit adds takes the first name `ColumnN` no
 * column of the table uses, as Excel names one, and its header cell is written.
 *
 * A deleted header row turns the header row off, and a deleted totals row the totals row; the rows in between stay a
 * table. Returns `false` when the edit deletes the whole table, or every data row of one with no header row left —
 * the caller removes it, as Excel does.
 */
export function tableApplyLineEdit(t: TableData, edit: LineEdit): boolean {
  checkTable(t);
  if (!tableSurvivesLineEdit(t, edit)) {
    return false;
  }
  const before = layoutOf(t.table);
  const rangeForm = t.table.ref.includes(":");
  if (edit.kind === "rows") {
    const top = before.row;
    const bottom = lastRowOf(before);
    const span = editSpan(top, bottom, edit)!;
    const headerRow = before.headerRow && editLine(top, edit) !== undefined;
    const totalsRow = before.totalsRow && editLine(bottom, edit) !== undefined;
    const height = span[1] - span[0] + 1 - (headerRow ? 1 : 0) - (totalsRow ? 1 : 0);
    if (span[0] === top && span[1] === bottom && headerRow === before.headerRow) {
      return true;
    }
    clearStaleSortState(t);
    t.table.headerRow = headerRow;
    t.table.totalsRow = totalsRow;
    t.table.rows = Array.from({ length: height }, (_, i) => t.table.rows[i] ?? []);
    t.table.ref = colCache.encodeAddress(span[0], before.col);
  } else {
    const left = before.col;
    const right = left + before.width - 1;
    const span = editSpan(left, right, edit)!;
    if (span[0] === left && span[1] === right) {
      return true;
    }
    const kept = new Map<number, LoadedTableColumnProperties>();
    t.table.columns.forEach((column, j) => {
      const col = editLine(left + j, edit);
      if (col !== undefined) {
        kept.set(col, column);
      }
    });
    const taken = new Set(t.table.columns.map(column => excelNameKey(column.name)));
    const added: number[] = [];
    const columns: LoadedTableColumnProperties[] = [];
    for (let col = span[0]; col <= span[1]; col++) {
      let column = kept.get(col);
      if (column === undefined) {
        let n = 1;
        while (taken.has(excelNameKey(`Column${n}`))) {
          n++;
        }
        column = { name: `Column${n}` };
        taken.add(excelNameKey(column.name));
        added.push(col - span[0]);
      }
      columns.push(column);
    }
    clearStaleSortState(t);
    t.table.columns = columns;
    t.table.rows = t.table.rows.map(() => []);
    t.table.ref = colCache.encodeAddress(before.row, span[0]);
    tableValidate(t);
    const after = layoutOf(t.table);
    for (const j of added) {
      if (after.headerRow) {
        cellSetValue(cellAt(t, headerRowOf(after), after.col + j), columns[j].name);
      }
      if (after.totalsRow) {
        cellSetValue(cellAt(t, totalsRowOf(after), after.col + j), totalsValue(t, columns[j]));
      }
    }
  }
  tableValidate(t);
  if (rangeForm) {
    t.table.ref = t.table.tableRef!;
  }
  return true;
}

/**
 * Whether the table is still there after a row or column edit of its sheet (see {@link tableApplyLineEdit}). Decided
 * before anything moves, so references to a table the edit removes can be turned into `#REF!` with the rest.
 */
export function tableSurvivesLineEdit(t: TableData, edit: LineEdit): boolean {
  checkTable(t);
  const layout = layoutOf(t.table);
  if (edit.kind === "columns") {
    return editSpan(layout.col, layout.col + layout.width - 1, edit) !== undefined;
  }
  const top = layout.row;
  const bottom = lastRowOf(layout);
  const span = editSpan(top, bottom, edit);
  if (span === undefined) {
    return false;
  }
  const headerRow = layout.headerRow && editLine(top, edit) !== undefined;
  const totalsRow = layout.totalsRow && editLine(bottom, edit) !== undefined;
  const height = span[1] - span[0] + 1 - (headerRow ? 1 : 0) - (totalsRow ? 1 : 0);
  return height > 0 || (height === 0 && headerRow);
}

/** Rewrite the formulas the table's model keeps — calculated columns and custom totals — through `edit`. */
export function tableRewriteFormulas(t: TableData, edit: ReferenceEdit, host: string): void {
  checkTable(t);
  for (const column of t.table.columns) {
    if (column.calculatedColumnFormula !== undefined) {
      column.calculatedColumnFormula = rewriteFormula(column.calculatedColumnFormula, edit, host);
    }
    if (column.totalsRowFormula !== undefined && column.totalsRowFunction === "custom") {
      column.totalsRowFormula = rewriteFormula(column.totalsRowFormula, edit, host);
    }
  }
}

// ============================================================================
// Live column styles (deprecated `Table.columnStyle`)
// ============================================================================

const liveStyles = new WeakMap<TableData, Set<TableColumnProperties>>();

/**
 * The deprecated `Table.columnStyle` handed out the column's own style object, and the table's next edit applied
 * whatever the caller had changed in it to every cell of the column. That is kept for the deprecated name only.
 */
export function tableMarkLiveStyle(view: TableColumnView): void {
  checkTableColumn(view);
  let set = liveStyles.get(view.table);
  if (!set) {
    set = new Set();
    liveStyles.set(view.table, set);
  }
  set.add(view.column);
}

function applyLiveStyles(t: TableData): void {
  const set = liveStyles.get(t);
  if (!set) {
    return;
  }
  liveStyles.delete(t);
  const layout = layoutOf(t.table);
  for (const column of set) {
    const j = t.table.columns.indexOf(column);
    if (j < 0) {
      continue;
    }
    for (let row = layout.row; row <= lastRowOf(layout); row++) {
      assignColumnStyle(cellAt(t, row, layout.col + j), column.style);
    }
  }
}

// ============================================================================
// Edits
// ============================================================================

/** What `Table.addRow` and `Table.removeRows` take. */
export interface TableRowEditOptions {
  /**
   * @deprecated Edits apply as they are made, and an appended row costs only its own cells, so there is nothing to
   * defer; this is ignored. Removed in 1.0.
   */
  readonly commit?: boolean;
}

/** @deprecated Every edit applies as it is made; there is never anything to commit. Removed in 1.0. */
export function tableCommit(t: TableData): void {
  checkTable(t);
  applyLiveStyles(t);
}

/**
 * Refuse a 0-based position that is not a whole number from 0 to `max`. A position outside the table used to reach
 * `splice`, which reads `-1` as "the last", truncates `1.5` and clamps `99` — so a row was inserted empty and its
 * values dropped, or the header row was copied into the body.
 */
function checkIndex(
  t: TableData,
  argument: string,
  value: unknown,
  max: number,
  what: string
): void {
  if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > max) {
    throw new TableError(
      max < 0
        ? `Table "${t.table.name}" has no ${what.replace(/^an? /, "")} at ${String(value)}: it has none.`
        : `Table "${t.table.name}": ${argument} must be ${what} from 0 to ${max}, got ${String(value)}.`
    );
  }
}

/** Refuse a count of rows or columns that is not a whole number of at least 0. */
function checkCount(t: TableData, count: unknown): void {
  if (typeof count !== "number" || !Number.isInteger(count) || count < 0) {
    throw new TableError(
      `Table "${t.table.name}": count must be a whole number of at least 0, got ${String(count)}.`
    );
  }
}

/** Refuse a data row that is not an array, has more values than the table has columns, or holds a value no cell can. */
function checkRowValues(t: TableData, where: string, argument: string, values: unknown): void {
  if (!Array.isArray(values)) {
    throw new TableError(
      `Table "${t.table.name}": ${argument} must be an array of cell values, got ${String(values)}.`
    );
  }
  if (values.length > t.table.columns.length) {
    throw new TableError(
      `Table "${t.table.name}" has ${t.table.columns.length} column(s); a row of ${values.length} value(s) was given.`
    );
  }
  values.forEach((value, i) => checkCellValue(value, { where, field: `${argument}[${i}]` }));
}

/** Insert a data row at 0-based `rowNumber` (the end when absent). */
export function tableAddRow(t: TableData, values: readonly CellValue[], rowNumber?: number): void;
/** @deprecated `options.commit` does nothing — every edit applies as it is made. Omit it. Removed in 1.0. */
export function tableAddRow(
  t: TableData,
  values: readonly CellValue[],
  rowNumber: number | undefined,
  options: TableRowEditOptions
): void;
export function tableAddRow(t: TableData, values: readonly CellValue[], rowNumber?: number): void {
  checkTable(t);
  const height = t.table.rows.length;
  checkRowValues(t, "Table.addRow", "values", values);
  const at = rowNumber ?? height;
  checkIndex(t, "rowNumber", at, height, "a data row position");
  editTable(
    t,
    () => {
      clearStaleSortState(t);
      t.table.rows.splice(at, 0, values.slice());
    },
    { rowSource: i => (i < at ? i : i === at ? values : i - 1) }
  );
}

/** Remove `count` (default 1) data rows from 0-based `rowIndex`. */
export function tableRemoveRows(t: TableData, rowIndex: number, count?: number): void;
/** @deprecated `options.commit` does nothing — every edit applies as it is made. Omit it. Removed in 1.0. */
export function tableRemoveRows(
  t: TableData,
  rowIndex: number,
  count: number | undefined,
  options: TableRowEditOptions
): void;
export function tableRemoveRows(t: TableData, rowIndex: number, count: number = 1): void {
  checkTable(t);
  checkCount(t, count);
  const height = t.table.rows.length;
  checkIndex(t, "rowIndex", rowIndex, count === 0 ? height : height - 1, "a data row");
  const removed = Math.max(0, Math.min(count, t.table.rows.length - rowIndex));
  editTable(
    t,
    () => {
      if (removed > 0) {
        clearStaleSortState(t);
      }
      t.table.rows.splice(rowIndex, removed);
    },
    { rowSource: i => (i < rowIndex ? i : i + removed) }
  );
}

/**
 * A view of the column at `colIndex`.
 *
 * @throws {TableError} when the table has no such column. The deprecated `Table.column` returns a view with no column
 * instead, which every setter then fails on.
 */
export function tableGetColumnChecked(t: TableData, colIndex: number): TableColumnView {
  checkTable(t);
  if (!Number.isInteger(colIndex) || colIndex < 0 || colIndex >= t.table.columns.length) {
    throw new TableError(
      `Table "${t.table.name}" has no column ${colIndex}; it has ${t.table.columns.length}.`
    );
  }
  return tableGetColumn(t, colIndex);
}

export function tableGetColumn(t: TableData, colIndex: number): TableColumnView {
  checkTable(t);
  const column = t.table.columns[colIndex];
  return { table: t, column, index: colIndex };
}

export function tableAddColumn(
  t: TableData,
  column: TableColumnProperties,
  values: readonly CellValue[],
  colIndex?: number
): void {
  checkTable(t);
  const width = t.table.columns.length;
  if (column === null || typeof column !== "object") {
    throw new TableError(
      `Table "${t.table.name}": the column must be an object with a name, got ${String(column)}.`
    );
  }
  if (!Array.isArray(values)) {
    throw new TableError(
      `Table "${t.table.name}": the column's values must be an array, got ${String(values)}.`
    );
  }
  if (values.length > t.table.rows.length) {
    throw new TableError(
      `Table "${t.table.name}" has ${t.table.rows.length} data row(s); ${values.length} value(s) were given for the new column.`
    );
  }
  values.forEach((value, i) =>
    checkCellValue(value, { where: "Table.addColumn", field: `values[${i}]` })
  );
  const at = colIndex ?? width;
  checkIndex(t, "colIndex", at, width, "a column position");
  editTable(
    t,
    () => {
      clearStaleSortState(t);
      t.table.columns.splice(at, 0, column);
      t.table.rows.forEach((row, i) => {
        row.splice(at, 0, values[i]);
      });
    },
    {
      columnSource: j => (j < at ? j : j === at ? undefined : j - 1),
      addedColumnValues: values
    }
  );
}

export function tableRemoveColumns(t: TableData, colIndex: number, count: number = 1): void {
  checkTable(t);
  checkCount(t, count);
  const width = t.table.columns.length;
  checkIndex(t, "colIndex", colIndex, count === 0 ? width : width - 1, "a column");
  const removed = Math.max(0, Math.min(count, t.table.columns.length - colIndex));
  editTable(
    t,
    () => {
      if (removed > 0) {
        clearStaleSortState(t);
      }
      t.table.columns.splice(colIndex, removed);
      t.table.rows.forEach(row => {
        row.splice(colIndex, removed);
      });
    },
    { columnSource: j => (j < colIndex ? j : j + removed) }
  );
}

export function tableRef(t: TableData): string {
  checkTable(t);
  return t.table.ref;
}
export function tableSetRef(t: TableData, value: string): void {
  checkTable(t);
  editTable(t, () => {
    if (t.table.ref !== value) {
      clearStaleSortState(t);
    }
    t.table.ref = value;
  });
}

/**
 * Refuse a table name a caller chose that Excel would not accept, naming the one it would. Such a name used to be
 * changed in silence (`T1` stored as `_T1`), so a formula the caller wrote with it named no table.
 */
export function checkTableName(where: string, field: string, name: unknown): void {
  const problem = tableNameProblem(name);
  if (problem !== undefined) {
    throw new TableError(
      `${where}: ${field} ${JSON.stringify(name)} is not a valid table name — ${problem}.`
    );
  }
}

export function tableName(t: TableData): string {
  checkTable(t);
  return t.table.name;
}
export function tableSetName(t: TableData, value: string): void {
  checkTable(t);
  checkTableName("Table.setName", "value", value);
  const newName = value;
  const oldName = t.table.name;
  if (newName === oldName) {
    return;
  }
  // Synchronise the worksheet's table map and the workbook-wide name set
  // so subsequent getTable(newName)/duplicate-name checks remain correct.
  // Falls back to a bare assignment if the worksheet hasn't registered
  // this table (e.g. transient instances built by Worksheet.set model).
  const ws = t.worksheet;
  const tables = ws?.tables;
  const tableNames = getSheetWorkbook(ws)?._tableNames;
  const registered = tables !== undefined && tables[oldName] === t;
  const newKey = excelNameKey(newName);
  const oldKey = excelNameKey(oldName);
  if (registered) {
    if (newKey !== oldKey && tableNames?.has(newKey)) {
      throw new TableError(
        `Table name "${newName}" already exists in the workbook (case-insensitive).`
      );
    }
    const definedNames = getSheetWorkbook(ws)?._definedNames;
    if (
      newKey !== oldKey &&
      definedNames !== undefined &&
      definedNamesHasName(definedNames, newName)
    ) {
      throw new TableError(
        `Table name "${newName}" is a defined name in the workbook (case-insensitive).`
      );
    }
    if (newKey !== oldKey) {
      checkTableNameFree(t, newName, "name");
    }
  }
  // The totals row's `SUBTOTAL` formulas name the table; nothing else on the sheet does.
  editTable(t, () => {
    t.table.name = newName;
  });
  if (registered) {
    delete tables[oldName];
    tables[newName] = t;
    if (tableNames) {
      tableNames.delete(oldKey);
      tableNames.add(newKey);
    }
  }
}

export function tableDisplayName(t: TableData): string {
  checkTable(t);
  return t.table.displayName || t.table.name;
}
export function tableSetDisplayName(t: TableData, value: string): void {
  checkTable(t);
  checkTableName("Table.setDisplayName", "value", value);
  checkTableEditable(t);
  checkTableNameFree(t, value, "displayName");
  // The display name lives in the table part only; no cell shows it.
  t.table.displayName = value;
}

/**
 * Refuse a name or display name another table of the workbook already answers to, or a defined name holds. Excel
 * resolves a structured reference by display name, so two tables showing the same one make every `T[col]` ambiguous
 * and the workbook opens behind a repair prompt.
 */
export function checkTableNameFree(
  t: TableData,
  value: string,
  field: "name" | "displayName"
): void {
  checkTable(t);
  const ws = t.worksheet;
  const workbook = ws ? getSheetWorkbook(ws) : undefined;
  if (!workbook) {
    return;
  }
  const key = excelNameKey(value);
  for (const sheet of workbook._worksheets) {
    for (const other of Object.values(sheet?.tables ?? {})) {
      if (other === t) {
        continue;
      }
      const names = [other.table.name, other.table.displayName].filter(
        (name): name is string => typeof name === "string"
      );
      if (names.some(name => excelNameKey(name) === key)) {
        throw new TableError(
          `Table ${field} "${value}" already exists in the workbook (case-insensitive): table "${other.table.name}" uses it.`
        );
      }
    }
  }
  if (definedNamesHasName(workbook._definedNames, value)) {
    throw new TableError(
      `Table ${field} "${value}" is a defined name in the workbook (case-insensitive).`
    );
  }
}

export function tableHeaderRow(t: TableData): boolean | undefined {
  checkTable(t);
  return t.table.headerRow;
}
export function tableSetHeaderRow(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  editTable(t, () => {
    if (t.table.headerRow !== value) {
      clearStaleSortState(t);
    }
    t.table.headerRow = value;
  });
}

export function tableTotalsRow(t: TableData): boolean | undefined {
  checkTable(t);
  return t.table.totalsRow;
}
export function tableSetTotalsRow(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  editTable(t, () => {
    if (t.table.totalsRow !== value && t.table.rawAttributes) {
      // `totalsRowShown` is preserved from the loaded file; once the public
      // totals-row state is changed the stale attribute must not survive.
      delete t.table.rawAttributes.totalsRowShown;
    }
    t.table.totalsRow = value;
  });
}

/** The table's style settings, for a setter to change: created when absent, refused on a removed sheet. */
function tableEnsureStyle(t: TableData): TableStyleProperties {
  checkTableEditable(t);
  if (!t.table.style) {
    t.table.style = {};
  }
  return t.table.style;
}

export function tableTheme(t: TableData): TableStyleProperties["theme"] {
  checkTable(t);
  return t.table.style?.theme;
}
export function tableSetTheme(t: TableData, value: TableStyleProperties["theme"]): void {
  checkTable(t);
  tableEnsureStyle(t).theme = value;
}

export function tableShowFirstColumn(t: TableData): boolean | undefined {
  checkTable(t);
  return t.table.style?.showFirstColumn;
}
export function tableSetShowFirstColumn(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  tableEnsureStyle(t).showFirstColumn = value;
}

export function tableShowLastColumn(t: TableData): boolean | undefined {
  checkTable(t);
  return t.table.style?.showLastColumn;
}
export function tableSetShowLastColumn(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  tableEnsureStyle(t).showLastColumn = value;
}

export function tableShowRowStripes(t: TableData): boolean | undefined {
  checkTable(t);
  return t.table.style?.showRowStripes;
}
export function tableSetShowRowStripes(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  tableEnsureStyle(t).showRowStripes = value;
}

export function tableShowColumnStripes(t: TableData): boolean | undefined {
  checkTable(t);
  return t.table.style?.showColumnStripes;
}
export function tableSetShowColumnStripes(t: TableData, value: boolean | undefined): void {
  checkTable(t);
  tableEnsureStyle(t).showColumnStripes = value;
}

// --- TableColumnView accessors (former Column getters/setters) ---------------

/** Refuse an edit through the view of a column the table no longer has: it changed an object nothing wrote. */
function checkColumnPresent(view: TableColumnView): void {
  if (!view.table.table.columns.includes(view.column)) {
    throw new ExcelArgumentError(
      "column",
      `Table "${view.table.table.name}" no longer has the column this view was taken for; take a new one with Table.getColumn.`
    );
  }
}

/** Set one column property and apply it to the sheet. Only a style reaches the column's data cells. */
function tableColumnSet<K extends keyof TableColumnProperties>(
  view: TableColumnView,
  name: K,
  value: TableColumnProperties[K]
): void {
  checkColumnPresent(view);
  const previous = view.column.style;
  editTable(
    view.table,
    () => {
      view.column[name] = value;
    },
    name === "style" ? { restyled: new Map([[view.column, previous]]) } : {}
  );
}

export function tableColumnName(view: TableColumnView): string {
  checkTableColumn(view);
  return view.column.name;
}
export function tableColumnSetName(view: TableColumnView, value: string): void {
  checkTableColumn(view);
  tableColumnSet(view, "name", value);
}

export function tableColumnFilterButton(view: TableColumnView): boolean | undefined {
  checkTableColumn(view);
  return view.column.filterButton;
}
export function tableColumnSetFilterButton(
  view: TableColumnView,
  value: boolean | undefined
): void {
  checkTableColumn(view);
  checkTableEditable(view.table);
  checkColumnPresent(view);
  view.column.filterButton = value;
}

export function tableColumnStyle(view: TableColumnView): Partial<Style> | undefined {
  checkTableColumn(view);
  return view.column.style;
}
export function tableColumnSetStyle(
  view: TableColumnView,
  value: Partial<Style> | undefined
): void {
  checkTableColumn(view);
  tableColumnSet(view, "style", value);
}

export function tableColumnTotalsRowLabel(view: TableColumnView): string | undefined {
  checkTableColumn(view);
  return view.column.totalsRowLabel;
}
export function tableColumnSetTotalsRowLabel(
  view: TableColumnView,
  value: string | undefined
): void {
  checkTableColumn(view);
  tableColumnSet(view, "totalsRowLabel", value);
}

export function tableColumnTotalsRowFunction(
  view: TableColumnView
): TableColumnProperties["totalsRowFunction"] {
  checkTableColumn(view);
  return view.column.totalsRowFunction;
}
export function tableColumnSetTotalsRowFunction(
  view: TableColumnView,
  value: TableColumnProperties["totalsRowFunction"]
): void {
  checkTableColumn(view);
  tableColumnSet(view, "totalsRowFunction", value);
}

export function tableColumnTotalsRowResult(view: TableColumnView): CellValue {
  checkTableColumn(view);
  return view.column.totalsRowResult;
}
export function tableColumnSetTotalsRowResult(
  view: TableColumnView,
  value: CellFormulaValue["result"]
): void {
  checkTableColumn(view);
  tableColumnSet(view, "totalsRowResult", value);
}

export function tableColumnTotalsRowFormula(view: TableColumnView): string | undefined {
  checkTableColumn(view);
  return view.column.totalsRowFormula;
}
export function tableColumnSetTotalsRowFormula(
  view: TableColumnView,
  value: string | undefined
): void {
  checkTableColumn(view);
  tableColumnSet(view, "totalsRowFormula", value);
}

export { sanitizeTableName };
