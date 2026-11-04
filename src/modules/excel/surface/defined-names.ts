/**
 * `DefinedNames` namespace surface — workbook/worksheet defined names.
 *
 * `import { DefinedNames } from "documonster/excel"` →
 *   `const dn = Workbook.getDefinedNames(wb)`, `DefinedNames.add(dn, "Sheet1!A1", "MyName")`,
 *   `DefinedNames.addFormula(dn, "Tax", "0.2")`, `DefinedNames.list(dn)`,
 *   `DefinedNames.getNames(dn, "Sheet1!A1")`, `DefinedNames.remove(...)`.
 *
 * Exposes the high-level operations; lower-level `*Ex` / matrix helpers stay
 * internal (import from `@excel/defined-names` directly if needed in tests).
 */
import type { DefinedNameModel, DefinedNamesData } from "@excel/core/defined-names";
import {
  checkDefinedNames,
  createDefinedNames,
  definedNamesGetNames,
  definedNamesAdd,
  definedNamesDelete,
  definedNamesRemove,
  definedNamesAddFormula,
  definedNamesCheckName,
  definedNamesGetAllEntries,
  definedNamesIsPlaceable,
  definedNamesModel
} from "@excel/core/defined-names";
import { loadDefinedNamesModel } from "@excel/core/model-load";
import { DefinedNameError } from "@excel/errors";
import type { ArgumentChecks } from "@excel/utils/arguments";
import { argumentChecks } from "@excel/utils/arguments";
import { excelNameKey } from "@utils/name-key";

/** The names that cover the cell at `addressStr` (`"Sheet1!A1"`); `[]` for an invalid reference or a range. */
export function getNames(dn: DefinedNamesData, addressStr: string): string[] {
  checkDefinedNames(dn);
  return definedNamesGetNames(dn, addressStr);
}

/**
 * @deprecated A registry made here belongs to no workbook and nothing can attach it to one; use
 * `Workbook.getDefinedNames(wb)`, which is the workbook's own. Removed in 1.0.
 */
export const create = createDefinedNames;

/** The scope of a name: for {@link add}, {@link addFormula}, {@link get} and the whole-name form of {@link remove}. */
export interface NameScopeOptions {
  /** The 0-based tab position of the sheet the name is scoped to; omit for a workbook-scoped name. */
  localSheetId?: number;
}

/**
 * Delete the name `name` — all of it, whether it refers to cells, is a formula, or is text this library could not
 * read — in the workbook scope, or the sheet scope `options.localSheetId`. Returns `true` if the name existed.
 * Formulas that use it read `#NAME?` afterwards, as they do when Excel's Name Manager deletes one.
 */
export function remove(dn: DefinedNamesData, name: string, options?: NameScopeOptions): boolean;
/**
 * Remove `name` from the cells at `location` (`"Sheet1!$A$1"`), keeping it on any other cells it covers. Returns
 * `true` if any cell at `location` carried it.
 */
export function remove(dn: DefinedNamesData, location: string, name: string): boolean;
export function remove(
  dn: DefinedNamesData,
  nameOrLocation: string,
  nameOrOptions?: string | NameScopeOptions
): boolean {
  checkDefinedNames(dn);
  const check = argumentChecks("DefinedNames.remove");
  if (typeof nameOrOptions === "string") {
    check.string("location", nameOrLocation);
    return definedNamesRemove(dn, nameOrLocation, nameOrOptions);
  }
  check.required("name", nameOrLocation);
  check.string("name", nameOrLocation);
  const localSheetId = checkScope(check, nameOrOptions);
  return definedNamesDelete(dn, nameOrLocation, localSheetId);
}

/**
 * The name `name` in the workbook scope, or the sheet scope `options.localSheetId`, as {@link list} reports it — a
 * reference name whose cells have all been removed included, with `ranges: []` — or `undefined` when there is none.
 * Compared as Excel compares names (`Rate` and `RATE` are one). A copy.
 */
export function get(
  dn: DefinedNamesData,
  name: string,
  options?: NameScopeOptions
): DefinedNameModel | undefined {
  checkDefinedNames(dn);
  const check = argumentChecks("DefinedNames.get");
  check.required("name", name);
  check.string("name", name);
  const localSheetId = checkScope(check, options);
  const key = excelNameKey(name);
  return list(dn).find(
    entry => excelNameKey(entry.name) === key && entry.localSheetId === localSheetId
  );
}

function checkScope(check: ArgumentChecks, options: unknown): number | undefined {
  if (!check.record("options", options)) {
    return undefined;
  }
  check.number("options.localSheetId", options.localSheetId, { min: 0, integer: true });
  return options.localSheetId as number | undefined;
}

/**
 * Give the cell or range at `location` (`"Sheet1!$A$1"`, `"Sheet1!$A$1:$B$3"`) the name `name` — in the workbook
 * scope, or the sheet scope `options.localSheetId` (the sheet's 0-based tab position), where a formula on that sheet
 * finds it before a workbook name of the same spelling. Returns the name as {@link get} reports it then.
 *
 * @throws {DefinedNameError} when `name` is not a name Excel accepts — one that reads as a cell reference (`A1`), or
 * starts with a digit — or `location` is not a cell or a range of the grid. Both used to be accepted: the first
 * produced a file Excel repairs, the second was dropped when the workbook was written.
 * @throws {ExcelArgumentError} when `options.localSheetId` is not a whole number from 0.
 */
export function add(
  dn: DefinedNamesData,
  location: string,
  name: string,
  options?: NameScopeOptions
): DefinedNameModel {
  checkDefinedNames(dn);
  definedNamesCheckName(dn, "DefinedNames.add", name);
  if (!definedNamesIsPlaceable(location)) {
    throw new DefinedNameError(
      `DefinedNames.add: ${JSON.stringify(location)} is not a cell or range reference such as "Sheet1!$A$1".`
    );
  }
  const localSheetId = checkScope(argumentChecks("DefinedNames.add"), options);
  definedNamesAdd(dn, location, name, localSheetId);
  return attached(dn, name, localSheetId);
}

/** The name just attached, as `get` reports it. */
function attached(
  dn: DefinedNamesData,
  name: string,
  localSheetId: number | undefined
): DefinedNameModel {
  return get(dn, name, localSheetId === undefined ? undefined : { localSheetId })!;
}

/**
 * Define `name` as a formula (`"0.2"`, `"Sheet1!$A$1*2"`) — in the workbook scope, or the sheet scope
 * `options.localSheetId`, as {@link add} does. Returns the name as {@link get} reports it then.
 *
 * @throws {DefinedNameError} when `name` is not a name Excel accepts, or `expression` is not a string.
 * @throws {ExcelArgumentError} when `options.localSheetId` is not a whole number from 0.
 */
export function addFormula(
  dn: DefinedNamesData,
  name: string,
  expression: string,
  options?: NameScopeOptions
): DefinedNameModel {
  checkDefinedNames(dn);
  definedNamesCheckName(dn, "DefinedNames.addFormula", name);
  if (typeof expression !== "string" || expression.length === 0) {
    throw new DefinedNameError(
      `DefinedNames.addFormula: the expression for "${name}" must be a non-empty string.`
    );
  }
  const localSheetId = checkScope(argumentChecks("DefinedNames.addFormula"), options);
  definedNamesAddFormula(dn, name, expression, localSheetId);
  return attached(dn, name, localSheetId);
}

/**
 * Replace every name with `value` — the inverse of {@link getModel}. Like a read, it keeps what the model holds, so a
 * model taken from another workbook applies as it is.
 *
 * @throws {ExcelArgumentError} when `value` is not an array of entries, each with a string `name`.
 * @throws {DefinedNameError} when two entries are one name in one scope (`N` and `n`).
 */
export function setModel(dn: DefinedNamesData, value: readonly DefinedNameModel[]): void {
  checkDefinedNames(dn);
  const check = argumentChecks("DefinedNames.setModel");
  check.required("value", value);
  check.array("value", value);
  value.forEach((entry, i) => {
    check.required(`value[${i}]`, entry);
    check.record(`value[${i}]`, entry);
    check.required(`value[${i}].name`, entry.name);
    check.string(`value[${i}].name`, entry.name, { nonEmpty: true });
    check.array(`value[${i}].ranges`, entry.ranges);
    check.number(`value[${i}].localSheetId`, entry.localSheetId, { min: 0, integer: true });
  });
  // Two entries for one name in one scope — `N` and `n` are one name to Excel — would be one entry in the registry,
  // and the first was dropped without a word.
  const seen = new Map<string, string>();
  for (const entry of value) {
    const key = `${excelNameKey(entry.name)}\0${entry.localSheetId ?? ""}`;
    const first = seen.get(key);
    if (first !== undefined) {
      throw new DefinedNameError(
        `DefinedNames.setModel: ${JSON.stringify(first)} and ${JSON.stringify(entry.name)} are one name ` +
          `(names compare without regard to case) in the same scope; give each name one entry.`
      );
    }
    seen.set(key, entry.name);
  }
  loadDefinedNamesModel(dn, value);
}

/**
 * Every name the registry holds, one entry per name and scope — including a reference name whose cells have all been
 * removed (`ranges: []`), which is no longer written. A copy: editing it changes nothing.
 */
export function list(dn: DefinedNamesData): DefinedNameModel[] {
  checkDefinedNames(dn);
  return definedNamesGetAllEntries(dn);
}

/**
 * The names as they are written — the round-trip model {@link setModel} takes, without the names that no longer refer
 * to anything. A copy: editing it changes nothing.
 */
export function getModel(dn: DefinedNamesData): DefinedNameModel[] {
  checkDefinedNames(dn);
  return definedNamesModel(dn);
}

/** A defined-names handle. */
export type { DefinedNamesData as Handle } from "@excel/core/defined-names";

// --- deprecated names ---

/** @deprecated Use {@link list}. Removed in 1.0. */
export const getAllEntries = list;
/** @deprecated Use {@link getModel}. Removed in 1.0. */
export const model = getModel;
