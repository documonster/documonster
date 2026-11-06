/**
 * Boundary checks for the handles of the feature namespaces — `Table`, `Chart`, `Chartsheet`, `Anchor`, `Form`,
 * `Image` and the streaming row and cell handles — in the form `checkWorksheet` and `checkWorkbook` give the document
 * model's: something that is not one (`42`, `null`, `{}`, a name) is an `ExcelArgumentError` naming `field`, rather
 * than a `TypeError` from inside or a silent `undefined`.
 *
 * Each is one or two property reads, on what every handle of the kind carries and nothing else that reaches it does.
 */
import type { AnchorData } from "@excel/core/anchor";
import type { CellData } from "@excel/core/cell";
import type { ChartsheetData } from "@excel/core/chartsheet";
import type { ColumnData } from "@excel/core/column";
import type { FormCheckboxData } from "@excel/core/form-control";
import type { WorksheetImage } from "@excel/core/image";
import type { RowData } from "@excel/core/row";
import type { TableColumnView, TableData } from "@excel/core/table";
import type { ChartHandle } from "@excel/core/worksheet-core";
import { describeHandle } from "@excel/core/worksheet-core";
import { ExcelArgumentError } from "@excel/errors";

type Fields = Record<string, unknown> | null | undefined;

function refuse(field: string, kind: string, value: unknown): never {
  throw new ExcelArgumentError(field, `${field} must be ${kind}, got ${describeHandle(value)}.`);
}

function isObject(value: unknown): boolean {
  return typeof value === "object" && value !== null;
}

/** A table, as `Table.add` / `Table.get` return it. */
export function checkTable(t: unknown, field = "table"): asserts t is TableData {
  if (!Array.isArray(((t as Fields)?.table as Fields)?.columns)) {
    refuse(field, "a table", t);
  }
}

/** A table column, as `Table.getColumn` returns it. */
export function checkTableColumn(view: unknown, field = "column"): asserts view is TableColumnView {
  if (
    !isObject((view as Fields)?.column) ||
    !isObject(((view as Fields)?.table as Fields)?.table)
  ) {
    refuse(field, "a table column", view);
  }
}

/** A chart, as `Chart.place` / `Chart.list` return it. */
export function checkChart(c: unknown, field = "chart"): asserts c is ChartHandle {
  if (typeof (c as Fields)?.chartNumber !== "number" || !isObject((c as Fields)?.worksheet)) {
    refuse(field, "a chart", c);
  }
}

/** A chartsheet, as `Workbook.addChartsheet` returns it. */
export function checkChartsheet(cs: unknown, field = "chartsheet"): asserts cs is ChartsheetData {
  if (!isObject((cs as Fields)?._model)) {
    refuse(field, "a chartsheet", cs);
  }
}

/** A drawing anchor, as `Anchor.create` returns it. */
export function checkAnchor(a: unknown, field = "anchor"): asserts a is AnchorData {
  if (typeof (a as Fields)?.nativeCol !== "number") {
    refuse(field, "an anchor", a);
  }
}

/** A form checkbox, as `Form.addCheckbox` returns it. */
export function checkFormCheckbox(c: unknown, field = "checkbox"): asserts c is FormCheckboxData {
  if (!isObject((c as Fields)?.model) || !isObject((c as Fields)?.worksheet)) {
    refuse(field, "a checkbox", c);
  }
}

/** A placed image, as `Image.place` / `Image.list` return it. */
export function checkWorksheetImage(img: unknown, field = "image"): asserts img is WorksheetImage {
  if (!isObject((img as Fields)?.worksheet)) {
    refuse(field, "an image", img);
  }
}

/** A streamed row, as `WorksheetWriter#addRow` returns it. */
export function checkRowHandle(r: unknown, field = "row"): asserts r is RowData {
  if (typeof (r as Fields)?.number !== "number" || !Array.isArray((r as Fields)?.cells)) {
    refuse(field, "a row", r);
  }
}

/** A streamed cell, as `Stream.getRowCell` returns it. */
export function checkCellHandle(c: unknown, field = "cell"): asserts c is CellData {
  if (typeof (c as Fields)?.address !== "string" || !isObject((c as Fields)?.row)) {
    refuse(field, "a cell", c);
  }
}

/** A streamed sheet's column, as `WorksheetWriter#getColumn` returns it. */
export function checkColumnHandle(column: unknown, field = "column"): asserts column is ColumnData {
  if (typeof (column as Fields)?.number !== "number" || !isObject((column as Fields)?.worksheet)) {
    refuse(field, "a column", column);
  }
}
