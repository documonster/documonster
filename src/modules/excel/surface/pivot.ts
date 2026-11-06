import type { PivotTable, PivotTableModel } from "@excel/core/pivot-table";
import { addPivotTable, getPivotTables, removePivotTable } from "@excel/core/worksheet";
/**
 * `Pivot` namespace surface.
 *
 * `import { Pivot } from "documonster/excel"` → `Pivot.add(ws, model)`,
 * `Pivot.list(ws)`, `Pivot.remove(ws, pivot)`.
 *
 * Also re-exports the public pivot-table model types so consumers reference
 * them through the namespace (`Pivot.Value`, `Pivot.Model`, …) instead of an
 * internal module path.
 */
import { checkWorksheet } from "@excel/core/worksheet-core";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { checkListTarget } from "@excel/surface/argument-checks";
import { argumentChecks, normalizeReference } from "@excel/utils/arguments";

function checkNames(check: ReturnType<typeof argumentChecks>, field: string, names: unknown): void {
  if (check.array(field, names)) {
    names.forEach((name, i) => {
      check.required(`${field}[${i}]`, name);
      check.string(`${field}[${i}]`, name);
    });
  }
}

/**
 * Add a pivot table to the sheet, built from `model.sourceSheet` (its first row names the fields) or
 * `model.sourceTable`, and return it.
 *
 * The source is followed, not copied: a sheet source is the range of the sheet's populated cells when the pivot table
 * is added, and it moves and grows with row and column edits as any other reference does; a table source is the
 * table. The cache — the fields' items and the records — is computed from the source when the workbook is written,
 * so values changed in it or rows it gained since are there; Excel refreshes it again on open.
 *
 * @throws {ExcelArgumentError} when the model is not of the documented shape.
 * @throws {ExcelInvalidAddressError} when `ref` is not a cell or range of this sheet (a sheet name is refused).
 * @throws {ExcelPivotTableError} when a field is not found in the source, the metric is not one Excel has, or the pivot
 * table would overlap another one on the sheet — give each its own `ref`, with room to grow.
 */
export function add(ws: WorksheetData, model: PivotTableModel): PivotTable {
  checkWorksheet(ws);
  const check = argumentChecks("Pivot.add");
  check.required("model", model);
  if (check.record("model", model)) {
    check.required("model.rows", model.rows);
    checkNames(check, "model.rows", model.rows);
    checkNames(check, "model.columns", model.columns);
    checkNames(check, "model.pages", model.pages);
    check.required("model.values", model.values);
    if (check.array("model.values", model.values, { min: 1 })) {
      model.values.forEach((value, i) => {
        if (typeof value !== "string") {
          check.required(`model.values[${i}]`, value);
          if (check.record(`model.values[${i}]`, value)) {
            check.required(`model.values[${i}].name`, value.name);
            check.string(`model.values[${i}].name`, value.name);
          }
        }
      });
    }
    check.string("model.ref", model.ref);
    // A cell (or a range, whose top-left is taken) of this sheet. A sheet name used to be dropped in silence.
    check.reference("model.ref", model.ref, { range: true });
    check.oneOf("model.applyWidthHeightFormats", model.applyWidthHeightFormats, ["0", "1"]);
  }
  return addPivotTable(
    ws,
    typeof model.ref === "string" ? { ...model, ref: normalizeReference(model.ref) } : model
  );
}

/** The sheet's pivot tables, in the order they were added: the handles themselves, which {@link remove} takes. */
export function list(ws: WorksheetData): PivotTable[] {
  checkWorksheet(ws);
  return getPivotTables(ws);
}

/**
 * Remove a pivot table — the handle {@link add} or {@link list} returned, or its index in that list. Returns `true` if
 * one was removed. A pivot chart or a slicer bound to it keeps it, and this throws a `ExcelPivotTableError`.
 */
export function remove(ws: WorksheetData, pivot: PivotTable | number): boolean {
  checkWorksheet(ws);
  checkListTarget("Pivot.remove", "pivot", pivot, "a pivot table");
  return removePivotTable(ws, pivot);
}

export type {
  PivotTable as Handle,
  PivotTableSource as Source,
  PivotTableValue as Value,
  PivotTableModel as Model,
  PivotTableSubtotal as Subtotal,
  PivotTableChartFormat as ChartFormat,
  DataField,
  CacheField,
  SharedItemValue,
  ParsedCacheDefinition,
  RecordValue
} from "@excel/core/pivot-table";
