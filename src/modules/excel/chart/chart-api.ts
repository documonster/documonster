import { ChartOptionsError } from "@excel/errors";
import { quoteSheetName } from "@excel/utils/address";
import { colCache } from "@excel/utils/col-cache";
import { getCell, getSheetName, getTable } from "@excel/worksheet-core";

import { type CellValueInputType, cellSetValue } from "../cell";
import { tableModel, type TableData } from "../table";
import type { Worksheet } from "../worksheet";
import type { AddChartExOptions, AddChartExSeriesOptions } from "./chart-ex-types";
import type { AddChartOptions, AddChartSeriesOptions } from "./types";

export interface SeriesFromColumnsOptions {
  categories?: string;
  values: string;
  name?: AddChartSeriesOptions["name"];
}

export interface AddChartFromTableOptions extends Omit<AddChartOptions, "series"> {
  categoryColumn?: string | number;
  valueColumns?: Array<string | number>;
  /** Use Excel structured references so charts expand with the table. Default: true. */
  structuredReferences?: boolean;
}

export interface AddChartFromRowsOptions<T extends Record<string, unknown>> extends Omit<
  AddChartOptions,
  "series"
> {
  x: keyof T & string;
  y: (keyof T & string) | Array<keyof T & string>;
  sheetName?: string;
  startCell?: string;
  includeHeaders?: boolean;
}

export function seriesFromColumns(
  sheetName: string,
  options: SeriesFromColumnsOptions
): AddChartSeriesOptions {
  return {
    name: options.name,
    categories: options.categories ? qualifyRange(sheetName, options.categories) : undefined,
    values: qualifyRange(sheetName, options.values)
  };
}

/**
 * Common layout resolved from a worksheet Table — shared by the classic and
 * ChartEx helpers so the two paths cannot drift. Previously the two helpers
 * re-implemented ~50 lines of near-identical code (column resolution, range
 * derivation, structured-vs-absolute reference emission); now both consume
 * this structure and only differ in how they assemble the final series.
 */
interface ResolvedTableLayout {
  dataStartRow: number;
  dataEndRow: number;
  tableLeft: number;
  tableName: string;
  columnNames: Array<string | undefined>;
  categoryIndex: number;
  valueColumns: number[];
  sheetName: string;
  structured: boolean;
}

function resolveTableLayout(
  worksheet: Worksheet,
  table: TableData | string,
  options: Pick<
    AddChartFromTableOptions,
    "categoryColumn" | "valueColumns" | "structuredReferences"
  >
): ResolvedTableLayout {
  const resolved = typeof table === "string" ? getTable(worksheet, table) : table;
  if (!resolved) {
    throw new ChartOptionsError(`Table not found: ${String(table)}.`);
  }
  const model = tableModel(resolved);
  const tableRef = colCache.decode(model.tableRef ?? model.ref);
  if (!("top" in tableRef)) {
    throw new ChartOptionsError(`Invalid table range: ${model.ref}.`);
  }
  if (tableRef.bottom < tableRef.top) {
    throw new ChartOptionsError(`Invalid table range: ${model.ref}.`);
  }
  const headerRow = model.headerRow !== false;
  const dataStartRow = tableRef.top + (headerRow ? 1 : 0);
  const dataEndRow = tableRef.bottom - (model.totalsRow ? 1 : 0);
  if (dataEndRow < dataStartRow) {
    throw new ChartOptionsError(`Table has no data rows: ${model.name}.`);
  }
  const columnNames = model.columns.map(c => c.name);
  const categoryIndex = resolveTableColumn(columnNames, options.categoryColumn ?? 0);
  const explicitValueCols = options.valueColumns;
  const valueColumns =
    explicitValueCols && explicitValueCols.length > 0
      ? explicitValueCols.map(col => resolveTableColumn(columnNames, col))
      : model.columns.map((_, i) => i).filter(i => i !== categoryIndex);
  if (valueColumns.length === 0) {
    throw new ChartOptionsError(`Table has no value columns: ${model.name}.`);
  }
  return {
    dataStartRow,
    dataEndRow,
    tableLeft: tableRef.left,
    tableName: model.name,
    columnNames,
    categoryIndex,
    valueColumns,
    sheetName: getSheetName(worksheet),
    structured: options.structuredReferences !== false
  };
}

function buildColumnReference(layout: ResolvedTableLayout, columnIndex: number): string {
  if (layout.structured) {
    const name = layout.columnNames[columnIndex];
    if (name === undefined) {
      throw new ChartOptionsError(
        `Table column at index ${columnIndex} has no name; cannot emit a structured reference.`
      );
    }
    return tableColumnReference(layout.tableName, name);
  }
  return absoluteRange(
    layout.sheetName,
    layout.dataStartRow,
    layout.tableLeft + columnIndex,
    layout.dataEndRow,
    layout.tableLeft + columnIndex
  );
}

export function chartOptionsFromTable(
  worksheet: Worksheet,
  table: TableData | string,
  options: AddChartFromTableOptions
): AddChartOptions {
  const layout = resolveTableLayout(worksheet, table, options);
  const categoryReference = buildColumnReference(layout, layout.categoryIndex);
  const series = layout.valueColumns.map(index => ({
    name: layout.columnNames[index],
    categories: categoryReference,
    values: buildColumnReference(layout, index)
  }));
  return { ...stripTableHelperOptions(options), series };
}

export function chartOptionsFromRows<T extends Record<string, unknown>>(
  worksheet: Worksheet,
  rows: T[],
  options: AddChartFromRowsOptions<T>
): AddChartOptions {
  if (rows.length === 0) {
    throw new ChartOptionsError("chartOptionsFromRows requires at least one row.");
  }
  const staged = stageRowsIntoWorksheet(worksheet, rows, options);
  const sheetName = getSheetName(worksheet);
  const series = staged.yKeys.map((key, i) => ({
    name: key,
    categories: absoluteRange(
      sheetName,
      staged.dataStartRow,
      staged.startCol,
      staged.dataEndRow,
      staged.startCol
    ),
    values: absoluteRange(
      sheetName,
      staged.dataStartRow,
      staged.startCol + i + 1,
      staged.dataEndRow,
      staged.startCol + i + 1
    )
  }));
  return { ...stripRowsHelperOptions(options), series };
}

/**
 * Table column helper — validates a column descriptor against a table's
 * header list. Strings match case-insensitively; numbers are zero-based
 * indexes. Missing column names are treated as unmatchable (not a
 * TypeError) so a table with a blank header column is diagnosable.
 */
function tableColumnReference(tableName: string, columnName: string): string {
  return `${escapeStructuredReferenceName(tableName)}[${escapeStructuredReferenceColumn(columnName)}]`;
}

/**
 * Quote a table name for use in a structured reference. Excel table names
 * are validated when they are created — they cannot contain whitespace,
 * `[`, `]`, `#`, `'`, `@`, or start with a digit — so the legal-name set
 * is very restricted. Any name that fails validation throws instead of
 * silently round-tripping invalid XML.
 *
 * Allows Unicode characters in the Basic Multilingual Plane (BMP)
 * above `\u00A0` — matches `cache-populator.ts` which accepts CJK
 * table names from Chinese / Japanese / Korean workbooks. The prior
 * ASCII-only regex threw on `"销售表"` or `"商品マスター"` even though
 * Excel itself permits them.
 */
function escapeStructuredReferenceName(name: string): string {
  // Canonical structure check — identifier must start with a letter /
  // underscore, then letters / digits / underscore / dot / non-ASCII.
  if (!/^[A-Za-z_\u00A0-\uFFFF][A-Za-z0-9_.\u00A0-\uFFFF]*$/.test(name)) {
    throw new ChartOptionsError(
      `Invalid Excel table name for structured reference: ${JSON.stringify(name)}.`
    );
  }
  // Reject Unicode whitespace, line / paragraph separators, zero-width
  // characters and the BOM. The structural regex above admits them via
  // the `\u00A0-\uFFFF` range, but Excel rejects them in real table
  // names — without this guard a name like `"Tab\u2028le"` would pass
  // our validation but break the formula parser downstream.
  if (/[\u00A0\u1680\u2000-\u200F\u2028\u2029\u202F\u205F\u3000\uFEFF]/.test(name)) {
    throw new ChartOptionsError(
      `Excel table name contains whitespace / separator / zero-width character: ${JSON.stringify(name)}.`
    );
  }
  return name;
}

/**
 * Escape the special characters (`[ ] # ' @`) in a structured-reference
 * column specifier by prefixing each with a single apostrophe, per
 * §18.17.2.4 of the Excel formula grammar.
 */
function escapeStructuredReferenceColumn(name: string): string {
  return name.replace(/([[\]#'@])/g, "'$1");
}

function resolveTableColumn(names: Array<string | undefined>, column: string | number): number {
  if (typeof column === "number") {
    if (column < 0 || column >= names.length) {
      throw new ChartOptionsError(`Table column index out of range: ${column}.`);
    }
    return column;
  }
  const target = column.toLowerCase();
  const idx = names.findIndex(name => name !== undefined && name.toLowerCase() === target);
  if (idx < 0) {
    throw new ChartOptionsError(`Table column not found: ${column}.`);
  }
  return idx;
}

function qualifyRange(sheetName: string, range: string): string {
  // Always produce an absolute A1 range. Chart formulas are expected
  // to remain stable across insertions/deletions, so relative refs
  // would silently shift rows/columns on the first edit.
  //
  // Split on the first `!` so pre-qualified refs (`"Sheet1!A1:B2"`)
  // still get their range part normalised to `$A$1:$B$2`. The old
  // early-return kept the caller's relative form verbatim, producing
  // divergent output depending on whether the input was sheet-
  // qualified or not.
  //
  // Normalise the caller-supplied sheet prefix through `quoteSheetName`
  // so `"My Sheet!A1"` and `"'My Sheet'!A1"` both emit the same
  // single-quoted form, matching what the unqualified branch produces
  // via `quoteSheetName(sheetName)`. `quoteSheetName` is idempotent
  // for safely-named sheets, so pre-quoted input stays pre-quoted.
  const bang = range.indexOf("!");
  if (bang >= 0) {
    const rawSheet = range.slice(0, bang);
    const ref = range.slice(bang + 1);
    // Strip one optional layer of single quotes + collapse escaped `''`
    // back to a literal single quote so we can re-quote through the
    // canonical helper. `'` is the only character OOXML escapes in
    // sheet names.
    const unquoted =
      rawSheet.length >= 2 && rawSheet.startsWith("'") && rawSheet.endsWith("'")
        ? rawSheet.slice(1, -1).replace(/''/g, "'")
        : rawSheet;
    // A reference that isn't a plain A1 cell/range — a structured
    // reference (`Table1[Sales]`) or a defined-name reference
    // (`MyData`) — must pass through unchanged. `colCache.decode`
    // doesn't throw on these inputs but silently returns garbage
    // (e.g. `Table1[Sales]` decodes as if it were a cell address "T1"),
    // which `absoluteA1Range` then emits as `$T$1`, silently corrupting
    // the formula. Detect non-A1 shapes and keep them verbatim.
    if (isA1RangeOrCell(ref)) {
      return `${quoteSheetName(unquoted)}!${absoluteA1Range(ref)}`;
    }
    // Pass through structured/named refs verbatim — the sheet prefix is
    // still legal (e.g. `Sheet1!Table1[Sales]` is accepted by Excel).
    return `${quoteSheetName(unquoted)}!${ref}`;
  }
  if (!isA1RangeOrCell(range)) {
    // Bare structured / named refs resolve against the workbook as a
    // whole; they don't take a sheet prefix.
    return range;
  }
  return `${quoteSheetName(sheetName)}!${absoluteA1Range(range)}`;
}

/**
 * Whether a reference string is a classic A1 cell or range (optionally
 * with `$` absolute markers). Returns false for structured references
 * like `Table1[Sales]`, defined-name references like `MyRange`, and
 * other non-A1 shapes so callers can route them away from
 * `colCache.decode` (which silently produces garbage on bad input).
 *
 * Also rejects column letters beyond Excel's maximum column XFD (16384)
 * so downstream `colCache.decode` / `n2l` calls don't throw cryptic
 * out-of-bounds errors.
 */
function isA1RangeOrCell(ref: string): boolean {
  // Cell: optional $, letters, optional $, digits. Range: two of those
  // joined by `:`. Full-column (`A:A`) and full-row (`1:1`) references
  // are also A1-shaped and accepted.
  if (
    !/^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$|^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d+:\$?\d+$/.test(
      ref
    )
  ) {
    return false;
  }
  // Validate column letters are within Excel's XFD (16384) limit.
  const colParts = ref.replace(/\$/g, "").match(/[A-Za-z]+/g);
  if (colParts) {
    for (const col of colParts) {
      const n = colCache.l2n(col.toUpperCase());
      if (!n || n > 16384) {
        return false;
      }
    }
  }
  return true;
}

function absoluteRange(
  sheetName: string,
  top: number,
  left: number,
  bottom: number,
  right: number
): string {
  return `${quoteSheetName(sheetName)}!$${colCache.n2l(left)}$${top}:$${colCache.n2l(right)}$${bottom}`;
}

function absoluteA1Range(range: string): string {
  // Normalize to uppercase before decoding — `colCache.decodeAddress`
  // only parses A-Z column letters; lowercase input (e.g. "b1:b2")
  // would produce garbage or throw without this.
  const decoded = colCache.decode(range.toUpperCase());
  if ("top" in decoded) {
    return `$${colCache.n2l(decoded.left)}$${decoded.top}:$${colCache.n2l(decoded.right)}$${decoded.bottom}`;
  }
  return `$${colCache.n2l(decoded.col)}$${decoded.row}`;
}

// ---------------------------------------------------------------------------
// Shared row-staging helpers
// ---------------------------------------------------------------------------

interface StagedRows<T extends Record<string, unknown>> {
  yKeys: Array<keyof T & string>;
  startCol: number;
  dataStartRow: number;
  dataEndRow: number;
}

/**
 * Write a rows-style dataset into the target worksheet and return the
 * resolved data extents. Shared by `chartOptionsFromRows` /
 * `chartExOptionsFromRows` so the two paths agree on cell coercion,
 * sheet-name validation, and header handling.
 */
function stageRowsIntoWorksheet<T extends Record<string, unknown>>(
  worksheet: Worksheet,
  rows: T[],
  options: {
    x: keyof T & string;
    y: (keyof T & string) | Array<keyof T & string>;
    sheetName?: string;
    startCell?: string;
    includeHeaders?: boolean;
  }
): StagedRows<T> {
  if (
    options.sheetName !== undefined &&
    !sheetNamesEqual(options.sheetName, getSheetName(worksheet))
  ) {
    throw new ChartOptionsError(
      `sheetName must match the target worksheet: got ${JSON.stringify(options.sheetName)}, expected ${JSON.stringify(getSheetName(worksheet))}.`
    );
  }
  const start = colCache.decodeAddress(options.startCell ?? "A1");
  const includeHeaders = options.includeHeaders !== false;
  const yKeys: Array<keyof T & string> = Array.isArray(options.y) ? options.y : [options.y];
  // Reject `y` lists that include the `x` key — without this check the
  // header row gets two copies of the same column name and the series
  // references `x` twice, producing a chart where the author sees the
  // same column as both the category and value axis without any
  // diagnostic.
  if (yKeys.some(k => k === (options.x as unknown))) {
    throw new ChartOptionsError(
      `chart y columns must not include the x key: x=${JSON.stringify(options.x)} appears in y=${JSON.stringify(yKeys)}.`
    );
  }
  const keys: Array<keyof T & string> = [options.x, ...yKeys];
  const startRow = start.row;
  const startCol = start.col;
  if (includeHeaders) {
    keys.forEach((key, i) => {
      cellSetValue(
        getCell(worksheet, startRow, startCol + i),
        key as unknown as CellValueInputType
      );
    });
  }
  const dataRowOffset = includeHeaders ? 1 : 0;
  rows.forEach((row, r) => {
    keys.forEach((key, c) => {
      cellSetValue(
        getCell(worksheet, startRow + dataRowOffset + r, startCol + c),
        coerceCellValue(row[key], key)
      );
    });
  });
  const dataStartRow = startRow + dataRowOffset;
  return {
    yKeys,
    startCol,
    dataStartRow,
    dataEndRow: dataStartRow + rows.length - 1
  };
}

/**
 * Narrow a value produced by indexing into a user-supplied row object to
 * one of the types `Worksheet.getCell().value` accepts. Unsupported shapes
 * (plain objects that aren't Date, Symbols, functions) throw with a
 * clear message rather than silently flowing into the worksheet as an
 * `as never` cast and surfacing elsewhere as a confusing writer error.
 */
function coerceCellValue(value: unknown, key: string): CellValueInputType {
  if (value === undefined || value === null) {
    return null;
  }
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    value instanceof Date
  ) {
    return value as CellValueInputType;
  }
  throw new ChartOptionsError(
    `Unsupported row value for column ${JSON.stringify(key)}: ${typeof value}. ` +
      `Expected string | number | boolean | Date | null | undefined.`
  );
}

function sheetNamesEqual(a: string, b: string): boolean {
  // Excel sheet names are case-insensitive.
  return a.toLowerCase() === b.toLowerCase();
}

/**
 * Strip the helper-only table options before spreading the remainder onto
 * an `AddChart(Ex)Options`. Previously the classic and ChartEx helpers
 * each had their own byte-for-byte duplicate of this function; the
 * structural constraint is all that matters at runtime, so we unify them
 * on a single generic.
 */
function stripTableHelperOptions<
  O extends {
    categoryColumn?: string | number;
    valueColumns?: Array<string | number>;
    structuredReferences?: boolean;
  }
>(options: O): Omit<O, "categoryColumn" | "valueColumns" | "structuredReferences"> {
  const { categoryColumn, valueColumns, structuredReferences, ...rest } = options;
  void categoryColumn;
  void valueColumns;
  void structuredReferences;
  return rest;
}

/**
 * Strip the helper-only rows options. Shared between the classic and
 * ChartEx rows helpers (see {@link stripTableHelperOptions}).
 */
function stripRowsHelperOptions<
  O extends {
    x: unknown;
    y: unknown;
    sheetName?: string;
    startCell?: string;
    includeHeaders?: boolean;
  }
>(options: O): Omit<O, "x" | "y" | "sheetName" | "startCell" | "includeHeaders"> {
  const { x, y, sheetName, startCell, includeHeaders, ...rest } = options;
  void x;
  void y;
  void sheetName;
  void startCell;
  void includeHeaders;
  return rest;
}

// ============================================================================
// ChartEx helpers
// ============================================================================

/**
 * Shared options for ChartEx helpers that derive series references from an
 * existing worksheet Table. The `type` is narrowed per-helper (histogram,
 * sunburst, treemap, waterfall, funnel, pareto, boxWhisker) because
 * `regionMap` requires geographic labels that don't map cleanly from a
 * generic category column. Column resolution mirrors
 * {@link AddChartFromTableOptions}: string matches the header
 * case-insensitively, number is a 0-based index.
 */
export interface AddChartExFromTableOptions extends Omit<
  AddChartExOptions,
  "series" | "categories"
> {
  /**
   * The category column. String is matched case-insensitively against the
   * table header; number is a 0-based index. Default: 0.
   */
  categoryColumn?: string | number;
  /**
   * The value columns. Each produces one `cx:series`. Defaults to every
   * non-category column.
   */
  valueColumns?: Array<string | number>;
  /**
   * Use Excel structured references (`Table1[Col]`) so the chart tracks
   * table expansion. Default: true. When `false`, classic absolute
   * ranges (`Sheet1!$B$2:$B$10`) are emitted instead.
   */
  structuredReferences?: boolean;
}

/**
 * Shared options for ChartEx helpers that derive series from a plain
 * JavaScript object array, mirroring {@link AddChartFromRowsOptions}. The
 * helper writes the rows into the worksheet first (headers optional) and
 * then emits the corresponding absolute ranges.
 */
export interface AddChartExFromRowsOptions<T extends Record<string, unknown>> extends Omit<
  AddChartExOptions,
  "series" | "categories"
> {
  /** Key used as the category column. */
  x: keyof T & string;
  /**
   * Key(s) used as value column(s). Each produces one `cx:series`.
   */
  y: (keyof T & string) | Array<keyof T & string>;
  /** Target worksheet name. Must match the worksheet this helper is invoked on. */
  sheetName?: string;
  /** Top-left cell for the staged data. Default: `A1`. */
  startCell?: string;
  /** Emit a header row before the data. Default: `true`. */
  includeHeaders?: boolean;
}

/**
 * Build an {@link AddChartExOptions} bundle by pointing at a worksheet Table.
 *
 * Mirrors {@link chartOptionsFromTable} for classic charts. The builder
 * resolves the table's data range and populates `categories` + one
 * series per value column. `regionMap` is intentionally unsupported
 * because its data model expects geographic labels that don't fit a
 * flat-column helper — use {@link buildChartExModel} directly.
 */
export function chartExOptionsFromTable(
  worksheet: Worksheet,
  table: TableData | string,
  options: AddChartExFromTableOptions & { type: Exclude<AddChartExOptions["type"], "regionMap"> }
): AddChartExOptions {
  const layout = resolveTableLayout(worksheet, table, options);
  const categories = buildColumnReference(layout, layout.categoryIndex);
  const series: AddChartExSeriesOptions[] = layout.valueColumns.map(index => ({
    name: layout.columnNames[index],
    values: buildColumnReference(layout, index)
  }));
  return { ...stripTableHelperOptions(options), categories, series };
}

/**
 * Build an {@link AddChartExOptions} bundle from a plain object-array
 * dataset, staging the rows into the worksheet before emitting ranges.
 * Mirrors {@link chartOptionsFromRows} for classic charts.
 */
export function chartExOptionsFromRows<T extends Record<string, unknown>>(
  worksheet: Worksheet,
  rows: T[],
  options: AddChartExFromRowsOptions<T> & { type: Exclude<AddChartExOptions["type"], "regionMap"> }
): AddChartExOptions {
  if (rows.length === 0) {
    throw new ChartOptionsError("chartExOptionsFromRows requires at least one row.");
  }
  const staged = stageRowsIntoWorksheet(worksheet, rows, options);
  const sheetName = getSheetName(worksheet);
  const categories = absoluteRange(
    sheetName,
    staged.dataStartRow,
    staged.startCol,
    staged.dataEndRow,
    staged.startCol
  );
  const series: AddChartExSeriesOptions[] = staged.yKeys.map((key, i) => ({
    name: key,
    values: absoluteRange(
      sheetName,
      staged.dataStartRow,
      staged.startCol + i + 1,
      staged.dataEndRow,
      staged.startCol + i + 1
    )
  }));
  return { ...stripRowsHelperOptions(options), categories, series };
}
