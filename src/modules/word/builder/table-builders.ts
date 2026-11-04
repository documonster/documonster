/**
 * Table builder functions for DOCX documents.
 *
 * Includes border, gridBorders, cell, row, table, simpleTable.
 */

import { isRecordArgument } from "@utils/object";
import { checkedTextParagraph } from "@word/builder/paragraph-builders";
import {
  checkArray,
  checkBorderStyle,
  checkHexColor,
  checkInteger,
  checkNodeArray,
  checkRecord,
  checkString,
  checkTableCellProperties,
  checkTableProperties,
  checkTableRowProperties,
  checkTableWidth
} from "@word/core/argument-checks";
import type {
  Border,
  Table,
  TableRow,
  TableCell,
  TableProperties,
  TableRowProperties,
  TableCellProperties,
  TableBorders,
  TableWidth,
  Paragraph,
  Twips
} from "@word/types";

// =============================================================================
// Table Builders
// =============================================================================

/** Options for {@link border}. */
export interface BorderOptions {
  /** Line style. Default: `"single"`. */
  readonly style?: Border["style"];
  /** Width in eighths of a point. Default: 4. */
  readonly size?: number;
  /** Hex color, or `"auto"`. Default: `"auto"`. */
  readonly color?: string;
}

/** Options for {@link gridBorders}. */
export interface GridBordersOptions {
  /** Width in eighths of a point. Default: 4. */
  readonly size?: number;
  /** Hex color, or `"auto"`. Default: `"auto"`. */
  readonly color?: string;
}

/** Shorthand border: `border({ style: "dashed", size: 8, color: "0070C0" })`. */
export function border(options?: BorderOptions): Border;
/**
 * @deprecated Pass an options object: `border({ style, size, color })`. Removed in 1.0.
 */
export function border(style: Border["style"] | undefined, size?: number, color?: string): Border;
export function border(
  styleOrOptions?: Border["style"] | BorderOptions,
  size?: number,
  color?: string
): Border {
  if (isRecordArgument(styleOrOptions)) {
    return borderOf(
      "Build.border",
      "options.",
      styleOrOptions.style,
      styleOrOptions.size,
      styleOrOptions.color
    );
  }
  return borderOf("Build.border", "", styleOrOptions, size, color);
}

/** `prefix` is how the caller named the values: `options.` for an options object. */
function borderOf(
  where: string,
  prefix: string,
  style: Border["style"] = "single",
  size = 4,
  color = "auto"
): Border {
  checkBorderStyle(style, `${prefix}style`, where);
  checkInteger(size, `${prefix}size`, where, { min: 0 });
  checkHexColor(color, `${prefix}color`, where);
  return { style, size, space: 0, color };
}

/** Create standard grid borders for a table: single lines, `{ size: 4, color: "auto" }` by default. */
export function gridBorders(options?: GridBordersOptions): TableBorders;
/**
 * @deprecated Pass an options object: `gridBorders({ size, color })`. Removed in 1.0.
 */
export function gridBorders(size: number | undefined, color?: string): TableBorders;
export function gridBorders(
  sizeOrOptions?: number | GridBordersOptions,
  color?: string
): TableBorders {
  const b = isRecordArgument(sizeOrOptions)
    ? borderOf("Build.gridBorders", "options.", "single", sizeOrOptions.size, sizeOrOptions.color)
    : borderOf("Build.gridBorders", "", "single", sizeOrOptions, color);
  return { top: b, left: b, bottom: b, right: b, insideH: b, insideV: b };
}

/** Create a table cell. */
export function cell(
  content: string | readonly (Paragraph | Table)[],
  properties?: TableCellProperties
): TableCell {
  return checkedCell("Build.cell", "content", content, properties);
}

function checkedCell(
  where: string,
  field: string,
  content: string | readonly (Paragraph | Table)[],
  properties: TableCellProperties | undefined
): TableCell {
  checkTableCellProperties(properties, "properties", where);
  if (typeof content === "string") {
    return { properties, content: [checkedTextParagraph(where, content, undefined)] };
  }
  if (!Array.isArray(content)) {
    // Text or blocks; anything else would be written as a cell with no content.
    checkString(content, field, where);
  }
  checkNodeArray(content, field, where);
  return { properties, content };
}

/** Create a table row. */
export function row(cells: readonly TableCell[], properties?: TableRowProperties): TableRow {
  checkNodeArray(cells, "cells", "Build.row");
  checkTableRowProperties(properties, "properties", "Build.row");
  return { properties, cells };
}

/** What {@link table} takes: the table's properties, and the width of each grid column. */
export interface TableOptions extends TableProperties {
  /** The width of each grid column (`w:tblGrid`), in twips. */
  readonly columnWidths?: readonly Twips[];
}

/** Create a table. */
export function table(rows: readonly TableRow[], options?: TableOptions): Table;
/** @deprecated Pass the widths in the options: `table(rows, { ...properties, columnWidths })`. Removed in 1.0. */
export function table(
  rows: readonly TableRow[],
  properties: TableProperties | undefined,
  columnWidths: readonly Twips[] | undefined
): Table;
export function table(
  rows: readonly TableRow[],
  options?: TableOptions,
  columnWidths?: readonly Twips[]
): Table {
  if (arguments.length > 2) {
    // The deprecated form names its own parameters.
    return checkedTable("Build.table", rows, options, columnWidths, "properties", "columnWidths");
  }
  if (options === undefined || !isRecordArgument(options) || !("columnWidths" in options)) {
    return checkedTable("Build.table", rows, options, undefined, "options", "options.columnWidths");
  }
  const { columnWidths: widths, ...properties } = options;
  return checkedTable("Build.table", rows, properties, widths, "options", "options.columnWidths");
}

function checkedTable(
  where: string,
  rows: readonly TableRow[],
  properties: TableProperties | undefined,
  columnWidths: readonly Twips[] | undefined,
  propertiesField: string,
  columnWidthsField: string
): Table {
  checkNodeArray(rows, "rows", where);
  checkTableProperties(properties, propertiesField, where);
  checkColumnWidths(columnWidths, where, columnWidthsField);
  return { type: "table", properties, columnWidths, rows };
}

function checkColumnWidths(columnWidths: unknown, where: string, field: string): void {
  if (columnWidths === undefined) {
    return;
  }
  checkArray(columnWidths, field, where);
  (columnWidths as unknown[]).forEach((w, i) =>
    checkInteger(w, `${field}[${i}]`, where, { min: 0 })
  );
}

/** Options for {@link simpleTable}. */
export interface SimpleTableOptions {
  /** Mark the first row as a header row, repeated at the top of every page (default: true). */
  readonly headerRow?: boolean;
  /** Draw single grid borders (default: true). */
  readonly borders?: boolean;
  /** The table's width (default: the full text width, `{ value: 5000, type: "pct" }`). */
  readonly width?: TableWidth;
  /** Column widths in twips, one per column. */
  readonly columnWidths?: readonly Twips[];
}

/** Create a simple table from a 2D string array. */
export function simpleTable(
  data: readonly (readonly string[])[],
  options?: SimpleTableOptions
): Table {
  return checkedSimpleTable("Build.simpleTable", data, options);
}

/**
 * A table of text cells, its arguments checked in the name of `where` — the
 * public function that was called.
 */
export function checkedSimpleTable(
  where: string,
  data: readonly (readonly string[])[],
  options: SimpleTableOptions | undefined
): Table {
  checkArray(data, "data", where);
  data.forEach((rowData, r) => {
    checkArray(rowData, `data[${r}]`, where);
    rowData.forEach((cellText, c) => checkString(cellText, `data[${r}][${c}]`, where));
  });
  checkRecord(options, "options", where, { optional: true });
  checkTableWidth(options?.width, "options.width", where, { optional: true });
  checkColumnWidths(options?.columnWidths, where, "options.columnWidths");
  const opts = { headerRow: true, borders: true, ...options };
  const tableRows: TableRow[] = data.map((rowData, rowIndex) => ({
    properties: rowIndex === 0 && opts.headerRow ? { tableHeader: true } : undefined,
    cells: rowData.map(cellText => checkedCell(where, "data", cellText, undefined))
  }));

  return {
    type: "table",
    properties: {
      width: opts.width ?? { value: 5000, type: "pct" },
      borders: opts.borders ? gridBorders() : undefined
    },
    columnWidths: opts.columnWidths,
    rows: tableRows
  };
}
