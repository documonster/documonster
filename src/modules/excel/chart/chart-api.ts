import { colCache } from "@excel/utils/col-cache";

import type { Table } from "../table";
import type { Worksheet } from "../worksheet";
import type { AddChartExOptions, AddChartExSeriesOptions } from "./chart-ex-types";
import type { AddChartOptions, AddChartSeriesOptions } from "./types";

export interface SeriesFromColumnsOptions {
  categories?: string;
  values: string;
  name?: string | { formula: string };
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

export function chartOptionsFromTable(
  worksheet: Worksheet,
  table: Table | string,
  options: AddChartFromTableOptions
): AddChartOptions {
  const resolved = typeof table === "string" ? worksheet.getTable(table) : table;
  if (!resolved) {
    throw new Error(`Table not found: ${String(table)}`);
  }
  const model = resolved.model;
  const tableRef = colCache.decode(model.tableRef ?? model.ref);
  if (!("top" in tableRef)) {
    throw new Error(`Invalid table range: ${model.ref}`);
  }
  if (tableRef.bottom < tableRef.top) {
    throw new Error(`Invalid table range: ${model.ref}`);
  }
  const headerRow = model.headerRow !== false;
  const dataStartRow = tableRef.top + (headerRow ? 1 : 0);
  const dataEndRow = tableRef.bottom - (model.totalsRow ? 1 : 0);
  if (dataEndRow < dataStartRow) {
    throw new Error(`Table has no data rows: ${model.name}`);
  }
  const categoryIndex = resolveTableColumn(
    model.columns.map(c => c.name),
    options.categoryColumn ?? 0
  );
  const valueColumns = options.valueColumns?.length
    ? options.valueColumns.map(col =>
        resolveTableColumn(
          model.columns.map(c => c.name),
          col
        )
      )
    : model.columns.map((_, i) => i).filter(i => i !== categoryIndex);
  const sheetName = worksheet.name;
  const series = valueColumns.map(index => ({
    name: model.columns[index]?.name,
    categories:
      options.structuredReferences === false
        ? absoluteRange(
            sheetName,
            dataStartRow,
            tableRef.left + categoryIndex,
            dataEndRow,
            tableRef.left + categoryIndex
          )
        : tableColumnReference(model.name, model.columns[categoryIndex]?.name),
    values:
      options.structuredReferences === false
        ? absoluteRange(
            sheetName,
            dataStartRow,
            tableRef.left + index,
            dataEndRow,
            tableRef.left + index
          )
        : tableColumnReference(model.name, model.columns[index]?.name)
  }));
  return { ...options, series };
}

export function chartOptionsFromRows<T extends Record<string, unknown>>(
  worksheet: Worksheet,
  rows: T[],
  options: AddChartFromRowsOptions<T>
): AddChartOptions {
  if (rows.length === 0) {
    throw new Error("chartOptionsFromRows requires at least one row");
  }
  if (options.sheetName !== undefined && options.sheetName !== worksheet.name) {
    const target = worksheet.workbook.getWorksheet(options.sheetName);
    if (!target || target !== worksheet) {
      throw new Error(
        `chartOptionsFromRows sheetName must match the target worksheet (${worksheet.name})`
      );
    }
  }
  const start = colCache.decodeAddress(options.startCell ?? "A1");
  const includeHeaders = options.includeHeaders !== false;
  const keys = [options.x, ...(Array.isArray(options.y) ? options.y : [options.y])];
  const startRow = start.row;
  const startCol = start.col;
  if (includeHeaders) {
    keys.forEach((key, i) => {
      worksheet.getCell(startRow, startCol + i).value = key;
    });
  }
  rows.forEach((row, r) => {
    keys.forEach((key, c) => {
      worksheet.getCell(startRow + (includeHeaders ? 1 : 0) + r, startCol + c).value = row[
        key
      ] as never;
    });
  });
  const dataStartRow = startRow + (includeHeaders ? 1 : 0);
  const dataEndRow = dataStartRow + rows.length - 1;
  const sheetName = options.sheetName ?? worksheet.name;
  const yKeys = Array.isArray(options.y) ? options.y : [options.y];
  const series = yKeys.map((key, i) => ({
    name: key,
    categories: absoluteRange(sheetName, dataStartRow, startCol, dataEndRow, startCol),
    values: absoluteRange(sheetName, dataStartRow, startCol + i + 1, dataEndRow, startCol + i + 1)
  }));
  return { ...options, series };
}

function tableColumnReference(tableName: string, columnName: string | undefined): string {
  if (!columnName) {
    throw new Error("Table column has no name");
  }
  return `${escapeStructuredReferenceName(tableName)}[${escapeStructuredReferenceColumn(columnName)}]`;
}

function escapeStructuredReferenceName(name: string): string {
  return /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

function escapeStructuredReferenceColumn(name: string): string {
  return name.replace(/([[\]#'@])/g, "'$1");
}

function resolveTableColumn(names: string[], column: string | number): number {
  if (typeof column === "number") {
    if (column < 0 || column >= names.length) {
      throw new Error(`Table column index out of range: ${column}`);
    }
    return column;
  }
  const idx = names.findIndex(name => name.toLowerCase() === column.toLowerCase());
  if (idx < 0) {
    throw new Error(`Table column not found: ${column}`);
  }
  return idx;
}

function qualifyRange(sheetName: string, range: string): string {
  return range.includes("!") ? range : `${quoteSheetName(sheetName)}!${absoluteA1Range(range)}`;
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
  const decoded = colCache.decode(range);
  if ("top" in decoded) {
    return `$${colCache.n2l(decoded.left)}$${decoded.top}:$${colCache.n2l(decoded.right)}$${decoded.bottom}`;
  }
  return `$${colCache.n2l(decoded.col)}$${decoded.row}`;
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z0-9_]+$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
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
  table: Table | string,
  options: AddChartExFromTableOptions & { type: Exclude<AddChartExOptions["type"], "regionMap"> }
): AddChartExOptions {
  const resolved = typeof table === "string" ? worksheet.getTable(table) : table;
  if (!resolved) {
    throw new Error(`Table not found: ${String(table)}`);
  }
  const model = resolved.model;
  const tableRef = colCache.decode(model.tableRef ?? model.ref);
  if (!("top" in tableRef)) {
    throw new Error(`Invalid table range: ${model.ref}`);
  }
  if (tableRef.bottom < tableRef.top) {
    throw new Error(`Invalid table range: ${model.ref}`);
  }
  const headerRow = model.headerRow !== false;
  const dataStartRow = tableRef.top + (headerRow ? 1 : 0);
  const dataEndRow = tableRef.bottom - (model.totalsRow ? 1 : 0);
  if (dataEndRow < dataStartRow) {
    throw new Error(`Table has no data rows: ${model.name}`);
  }
  const columnNames = model.columns.map(c => c.name);
  const categoryIndex = resolveTableColumn(columnNames, options.categoryColumn ?? 0);
  const valueColumns = options.valueColumns?.length
    ? options.valueColumns.map(col => resolveTableColumn(columnNames, col))
    : model.columns.map((_, i) => i).filter(i => i !== categoryIndex);
  if (valueColumns.length === 0) {
    throw new Error(`Table has no value columns: ${model.name}`);
  }
  const sheetName = worksheet.name;
  const structured = options.structuredReferences !== false;
  const categories = structured
    ? tableColumnReference(model.name, columnNames[categoryIndex])
    : absoluteRange(
        sheetName,
        dataStartRow,
        tableRef.left + categoryIndex,
        dataEndRow,
        tableRef.left + categoryIndex
      );
  const series: AddChartExSeriesOptions[] = valueColumns.map(index => ({
    name: columnNames[index],
    values: structured
      ? tableColumnReference(model.name, columnNames[index])
      : absoluteRange(
          sheetName,
          dataStartRow,
          tableRef.left + index,
          dataEndRow,
          tableRef.left + index
        )
  }));
  // Strip helper-only keys before spreading.
  const {
    categoryColumn: _catCol,
    valueColumns: _valCols,
    structuredReferences: _structured,
    ...rest
  } = options;
  void _catCol;
  void _valCols;
  void _structured;
  return { ...rest, categories, series };
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
    throw new Error("chartExOptionsFromRows requires at least one row");
  }
  if (options.sheetName !== undefined && options.sheetName !== worksheet.name) {
    const target = worksheet.workbook.getWorksheet(options.sheetName);
    if (!target || target !== worksheet) {
      throw new Error(
        `chartExOptionsFromRows sheetName must match the target worksheet (${worksheet.name})`
      );
    }
  }
  const start = colCache.decodeAddress(options.startCell ?? "A1");
  const includeHeaders = options.includeHeaders !== false;
  const yKeys = Array.isArray(options.y) ? options.y : [options.y];
  const keys = [options.x, ...yKeys];
  const startRow = start.row;
  const startCol = start.col;
  if (includeHeaders) {
    keys.forEach((key, i) => {
      worksheet.getCell(startRow, startCol + i).value = key;
    });
  }
  rows.forEach((row, r) => {
    keys.forEach((key, c) => {
      worksheet.getCell(startRow + (includeHeaders ? 1 : 0) + r, startCol + c).value = row[
        key
      ] as never;
    });
  });
  const dataStartRow = startRow + (includeHeaders ? 1 : 0);
  const dataEndRow = dataStartRow + rows.length - 1;
  const sheetName = options.sheetName ?? worksheet.name;
  const categories = absoluteRange(sheetName, dataStartRow, startCol, dataEndRow, startCol);
  const series: AddChartExSeriesOptions[] = yKeys.map((key, i) => ({
    name: key,
    values: absoluteRange(sheetName, dataStartRow, startCol + i + 1, dataEndRow, startCol + i + 1)
  }));
  const {
    x: _x,
    y: _y,
    sheetName: _sheet,
    startCell: _startCell,
    includeHeaders: _headers,
    ...rest
  } = options;
  void _x;
  void _y;
  void _sheet;
  void _startCell;
  void _headers;
  return { ...rest, categories, series };
}
