/**
 * `DataValidation` namespace surface — a sheet's data-validation rules, by cell or range.
 *
 * `import { DataValidation } from "documonster/excel"` →
 *   `DataValidation.add(ws, "A1:A10", rule)`, `DataValidation.find(ws, "A3")`,
 *   `DataValidation.remove(ws, "A1:A10")`, `DataValidation.list(ws)`, `DataValidation.clear(ws)`.
 *
 * Rules go in and come out as copies: the collection holds its own, so one rule object can be given to several
 * sheets, and a rule `find` returned can be edited and added back without affecting anything until it is.
 *
 * Each also takes a sheet's rule collection by itself — `ws.dataValidations` — in place of the sheet. Cell-level validation (`Cell.getValidation` / `Cell.setValidation`) lives on the `Cell`
 * namespace.
 */
import type { DataValidationModel, DataValidationsData } from "@excel/core/data-validations";
import {
  createDataValidations,
  dataValidationAdd,
  dataValidationFind,
  dataValidationRemove,
  dataValidationsOwner
} from "@excel/core/data-validations";
import type { WorksheetData } from "@excel/core/worksheet-core";
import { checkWorksheetAttached } from "@excel/core/worksheet-core";
import { checkDataValidationRule } from "@excel/surface/argument-checks";
import type { DataValidationRule } from "@excel/types";
import { argumentChecks, normalizeReference } from "@excel/utils/arguments";
import { isRecordArgument } from "@utils/object";

/**
 * @deprecated A collection made here belongs to no sheet and nothing can attach it to one, so its rules are never
 * written; pass the worksheet (or its own `ws.dataValidations`) to {@link add} instead. Removed in 1.0.
 */
export const create = (model?: DataValidationModel): DataValidationsData =>
  createDataValidations(model);

/** The rule collection a call names: a sheet's, or the collection itself. An edit refuses a removed sheet. */
function rules(
  target: WorksheetData | DataValidationsData,
  where: string,
  edit: boolean
): DataValidationsData {
  const check = argumentChecks(where);
  check.required("target", target);
  check.record("target", target);
  if ("dataValidations" in target) {
    if (edit) {
      checkWorksheetAttached(target);
    }
    return target.dataValidations;
  }
  if (!isRecordArgument((target as { model?: unknown }).model)) {
    check.fail("target", "a worksheet or a data-validation collection", target);
  }
  // A collection a worksheet holds is that sheet's: an edit through it is refused as one through the sheet is.
  const owner = dataValidationsOwner(target);
  if (edit && owner !== undefined) {
    checkWorksheetAttached(owner as WorksheetData);
  }
  return target;
}

/**
 * A reference of the collection's own sheet, upper-cased as Excel reads it. A sheet qualifier used to be dropped in
 * silence, so `"Other!E1"` put the rule on this sheet's `E1`.
 */
function localReference(where: string, argument: string, ref: unknown, range: boolean): string {
  const check = argumentChecks(where);
  check.required(argument, ref);
  check.reference(argument, ref, { range, lines: range });
  return normalizeReference(ref as string);
}

/**
 * Register a validation for a cell or A1 range of `target` — a worksheet, or a rule collection — and return a copy of
 * it.
 *
 * Whole-column (`"A:A"`) and whole-row (`"1:3"`) references are expanded to Excel's sheet limits. The collection keeps
 * its own copy of `validation`: the same rule can be added to several sheets, and editing it afterwards — or the copy
 * returned — changes nothing that was added. Replace a rule by adding another at the same `ref`.
 *
 * @throws {ExcelArgumentError} when the rule is not one a workbook can hold — an unknown type, operator or error
 * style, or bounds that are not formulas, numbers or dates.
 * @throws {ExcelInvalidAddressError} when `ref` is not a cell or range of the sheet (a sheet name is refused).
 */
export function add(
  target: WorksheetData | DataValidationsData,
  ref: string,
  validation: DataValidationRule
): DataValidationRule {
  // `undefined` used to remove the rule and return `undefined`, which the return type does not allow; `remove` does that.
  argumentChecks("DataValidation.add").required("validation", validation);
  checkDataValidationRule("DataValidation.add", validation);
  const local = localReference("DataValidation.add", "ref", ref, true);
  const collection = rules(target, "DataValidation.add", true);
  return structuredClone(dataValidationAdd(collection, local, structuredClone(validation)));
}

/** A copy of the validation covering the cell `address`, or `undefined`. */
export function find(
  target: WorksheetData | DataValidationsData,
  address: string
): DataValidationRule | undefined {
  const local = localReference("DataValidation.find", "address", address, false);
  const rule = dataValidationFind(rules(target, "DataValidation.find", false), local);
  return rule === undefined ? undefined : structuredClone(rule);
}

/** Remove the validation registered for `ref`. Returns `true` if there was one. */
export function remove(target: WorksheetData | DataValidationsData, ref: string): boolean {
  const local = localReference("DataValidation.remove", "ref", ref, true);
  return dataValidationRemove(rules(target, "DataValidation.remove", true), local);
}

/** One registered validation, as {@link list} reports it. */
export interface Entry {
  /** The cell or range the rule covers (`"B2:B9"`). */
  ref: string;
  /** A copy of the rule. */
  rule: DataValidationRule;
}

/** Every validation of `target`, in the order they were registered — copies, so editing one changes nothing. */
export function list(target: WorksheetData | DataValidationsData): Entry[] {
  const { model } = rules(target, "DataValidation.list", false);
  const entries: Entry[] = [];
  for (const [key, rule] of Object.entries(model)) {
    if (rule !== undefined && rule !== null) {
      entries.push({
        ref: key.startsWith("range:") ? key.slice("range:".length) : key,
        rule: structuredClone(rule)
      });
    }
  }
  return entries;
}

/** Remove every validation of `target`. */
export function clear(target: WorksheetData | DataValidationsData): void {
  const { model } = rules(target, "DataValidation.clear", true);
  for (const key of Object.keys(model)) {
    delete model[key];
  }
}

/** A data-validations handle. */
export type { DataValidationsData as Handle } from "@excel/core/data-validations";
