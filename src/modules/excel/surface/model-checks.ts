/**
 * Checks for the models the `setModel` members apply. A model is a round-trip shape, so these check what the
 * model-load path relies on and what no workbook can hold — the same refusals `Cell.setValue` makes — and not
 * every field a file may carry.
 */
import { checkCellValue, checkFormulaResult } from "@excel/core/cell";
import type { CellModel } from "@excel/core/cell";
import { ValueType } from "@excel/core/enums";
import { argumentChecks } from "@excel/utils/arguments";
import { colCache } from "@excel/utils/col-cache";

const KNOWN_TYPES = new Set<number>(Object.values(ValueType));

function fail(where: string, argument: string, expected: string, value: unknown): never {
  return argumentChecks(where).fail(argument, expected, value);
}

/** The stored value of a cell model must be of the kind its `type` says, and a value a workbook can hold. */
function checkModelValue(where: string, argument: string, model: Record<string, unknown>): void {
  const { type, value } = model;
  const check = argumentChecks(where);
  switch (type) {
    case ValueType.Number:
    case ValueType.SharedString:
      if (typeof value !== "number" || !Number.isFinite(value)) {
        fail(where, `${argument}.value`, "a finite number", value);
      }
      break;
    case ValueType.String:
      // A rich-text cell and a JSON cell are stored as strings too; the former holds its runs.
      if (
        typeof value !== "string" &&
        !(
          typeof value === "object" &&
          value !== null &&
          Array.isArray((value as { richText?: unknown }).richText)
        )
      ) {
        fail(where, `${argument}.value`, "a string or rich text", value);
      }
      break;
    case ValueType.RichText:
      if (
        typeof value !== "object" ||
        value === null ||
        !Array.isArray((value as { richText?: unknown }).richText)
      ) {
        fail(where, `${argument}.value`, "rich text ({ richText: [...] })", value);
      }
      break;
    case ValueType.Date:
      if (!(value instanceof Date) || Number.isNaN(value.getTime())) {
        fail(where, `${argument}.value`, "a valid Date", value);
      }
      break;
    case ValueType.Boolean:
    case ValueType.Checkbox:
      if (typeof value !== "boolean") {
        fail(where, `${argument}.value`, "a boolean", value);
      }
      break;
    case ValueType.Error:
      if (
        typeof value !== "object" ||
        value === null ||
        typeof (value as { error?: unknown }).error !== "string"
      ) {
        fail(where, `${argument}.value`, 'an error value ({ error: "#N/A" })', value);
      }
      break;
    case ValueType.Formula:
      check.string(`${argument}.formula`, model.formula);
      check.string(`${argument}.sharedFormula`, model.sharedFormula);
      if (model.formula === undefined && model.sharedFormula === undefined) {
        fail(where, `${argument}.formula`, "a formula or a shared formula", undefined);
      }
      checkFormulaResult(model.result, { where, field: `${argument}.result` });
      break;
    case ValueType.Hyperlink:
      check.string(`${argument}.text`, model.text);
      check.string(`${argument}.hyperlink`, model.hyperlink);
      break;
    default:
      // Null, Merge and JSON carry no value of their own to check.
      if (type === ValueType.JSON) {
        checkCellValue(value, { where, field: `${argument}.value` });
      }
  }
}

/**
 * A cell model: an object with a known `type`, a value of that kind, and — when `address` is given — that address.
 * A model filed under another address was written into the wrong `<row>`, and the file could not be read back.
 */
export function checkCellModel(
  where: string,
  argument: string,
  model: unknown,
  address?: string
): asserts model is CellModel {
  const check = argumentChecks(where);
  check.required(argument, model);
  check.record(argument, model);
  const record = model as Record<string, unknown>;
  if (typeof record.type !== "number" || !KNOWN_TYPES.has(record.type)) {
    fail(where, `${argument}.type`, "a ValueType", record.type);
  }
  if (address !== undefined && record.address !== address) {
    fail(where, `${argument}.address`, JSON.stringify(address), record.address);
  }
  checkModelValue(where, argument, record);
}

/**
 * A worksheet model: the collections the load path walks unconditionally are arrays, and every cell sits in the row
 * its address names.
 */
export function checkWorksheetModel(where: string, argument: string, model: unknown): void {
  const check = argumentChecks(where);
  check.required(argument, model);
  check.record(argument, model);
  const record = model as Record<string, unknown>;
  check.string(`${argument}.name`, record.name);
  check.required(`${argument}.media`, record.media);
  check.array(`${argument}.media`, record.media);
  check.required(`${argument}.tables`, record.tables);
  check.array(`${argument}.tables`, record.tables);
  for (const key of [
    "rows",
    "cols",
    "mergeCells",
    "pivotTables",
    "formControls",
    "rowBreaks",
    "colBreaks",
    "views",
    "conditionalFormattings",
    "ignoredErrors"
  ]) {
    check.array(`${argument}.${key}`, record[key] ?? undefined);
  }
  (record.rows as unknown[] | undefined)?.forEach((row, i) => {
    const at = `${argument}.rows[${i}]`;
    if (row === undefined || row === null) {
      return;
    }
    check.record(at, row);
    const { number, cells } = row as { number?: unknown; cells?: unknown };
    check.required(`${at}.number`, number);
    check.number(`${at}.number`, number, { min: 1, max: 1048576, integer: true });
    check.array(`${at}.cells`, cells ?? undefined);
    (cells as unknown[] | undefined)?.forEach((cell, j) => {
      if (cell === undefined || cell === null) {
        return;
      }
      checkCellModel(where, `${at}.cells[${j}]`, cell);
      const decoded = decodeAddress((cell as { address?: unknown }).address);
      if (decoded?.row !== number) {
        fail(
          where,
          `${at}.cells[${j}].address`,
          `a cell of row ${number}`,
          (cell as CellModel).address
        );
      }
    });
  });
  (record.mergeCells as unknown[] | undefined)?.forEach((merge, i) =>
    check.reference(`${argument}.mergeCells[${i}]`, merge, { range: true })
  );
}

function decodeAddress(address: unknown): { row: number } | undefined {
  if (typeof address !== "string") {
    return undefined;
  }
  try {
    return colCache.decodeAddress(address);
  } catch {
    return undefined;
  }
}

/** A workbook model: what the load path walks unconditionally is there, and each worksheet model is one. */
export function checkWorkbookModel(where: string, model: unknown): void {
  const check = argumentChecks(where);
  check.required("model", model);
  check.record("model", model);
  const record = model as Record<string, unknown>;
  check.required("model.worksheets", record.worksheets);
  check.array("model.worksheets", record.worksheets);
  check.required("model.definedNames", record.definedNames);
  check.array("model.definedNames", record.definedNames);
  (record.worksheets as unknown[]).forEach((sheet, i) =>
    checkWorksheetModel(where, `model.worksheets[${i}]`, sheet)
  );
}
