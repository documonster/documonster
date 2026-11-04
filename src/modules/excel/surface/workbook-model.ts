/**
 * `Workbook.setModel`, shared by the Node and browser `Workbook` surfaces. A module of its own because it reaches the
 * model-load path — the formula tokenizer and parser, the chart handle — which a write-only bundle must not carry;
 * see `core/model-load.ts`.
 */
import { loadWorkbookModel } from "@excel/core/model-load";
import type { WorkbookData } from "@excel/core/workbook-core";
import { getWorksheets } from "@excel/core/workbook-core";
import type { WorkbookModel } from "@excel/core/workbook-model";
import { createWorkbook } from "@excel/core/workbook.browser";
import { checkWorkbookModel } from "@excel/surface/model-checks";

/**
 * Replace the workbook's contents with `value` — the inverse of `getModel`.
 *
 * The model is checked, and then applied to a scratch workbook first: a model that fails part-way — a sheet name
 * Excel refuses, merges that conflict, a table the workbook cannot hold — leaves this workbook as it was, every sheet
 * included. Applying it replaces the worksheets, so the handles of the previous ones are retired: an edit through one
 * is refused, as it is for a removed sheet, instead of vanishing.
 *
 * @throws {ExcelArgumentError} when the model is malformed.
 */
export function setModel(wb: WorkbookData, value: WorkbookModel): void {
  checkWorkbookModel("Workbook.setModel", value);
  const probe = wb._definedNames._explicitProbe;
  loadWorkbookModel(createWorkbook(probe ? { formulaSyntaxProbe: probe } : undefined), value);
  const previous = getWorksheets(wb);
  loadWorkbookModel(wb, value);
  for (const worksheet of previous) {
    worksheet._removed = true;
  }
}
