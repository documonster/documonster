/**
 * Table builder functions for DOCX documents.
 *
 * Includes border, gridBorders, cell, row, table, simpleTable.
 */

import { textParagraph } from "@word/builder/paragraph-builders";
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

/** Shorthand border. */
export function border(style: Border["style"] = "single", size = 4, color = "auto"): Border {
  return { style, size, space: 0, color };
}

/** Create standard grid borders for a table. */
export function gridBorders(size = 4, color = "auto"): TableBorders {
  const b = border("single", size, color);
  return { top: b, left: b, bottom: b, right: b, insideH: b, insideV: b };
}

/** Create a table cell. */
export function cell(
  content: string | (Paragraph | Table)[],
  properties?: TableCellProperties
): TableCell {
  if (typeof content === "string") {
    return { properties, content: [textParagraph(content)] };
  }
  return { properties, content };
}

/** Create a table row. */
export function row(cells: TableCell[], properties?: TableRowProperties): TableRow {
  return { properties, cells };
}

/** Create a table. */
export function table(
  rows: TableRow[],
  properties?: TableProperties,
  columnWidths?: Twips[]
): Table {
  return { type: "table", properties, columnWidths, rows };
}

/** Create a simple table from a 2D string array. */
export function simpleTable(
  data: string[][],
  options?: {
    headerRow?: boolean;
    borders?: boolean;
    width?: TableWidth;
    columnWidths?: Twips[];
  }
): Table {
  const opts = { headerRow: true, borders: true, ...options };
  const tableRows: TableRow[] = data.map((rowData, rowIndex) => {
    const cells = rowData.map(cellText => cell(cellText));
    return row(cells, rowIndex === 0 && opts.headerRow ? { tableHeader: true } : undefined);
  });

  return table(
    tableRows,
    {
      width: opts.width ?? { value: 5000, type: "pct" },
      borders: opts.borders ? gridBorders() : undefined
    },
    opts.columnWidths
  );
}
