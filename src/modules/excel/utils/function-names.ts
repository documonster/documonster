/**
 * The prefixes a file spells function and parameter names with, and the two predicates on them.
 *
 * Apart from `future-functions.ts` because the XLSB codec needs them too, and
 * `future-functions.ts` needs the `Ftab` to tell a worksheet function's name from a user's.
 */

import { excelNameKey, sheetNameKey } from "@utils/name-key";

export const XLFN = "_xlfn.";
export const XLWS = "_xlws.";
export const XLPM = "_xlpm.";

/** Whether `name` is a `LET` / `LAMBDA` parameter as a file spells it, `_xlpm.x`. */
export function isLambdaParameter(name: string): boolean {
  return name.toLowerCase().startsWith(XLPM);
}

/**
 * The key a name in a workbook's defined-name list is compared by: a `LET` / `LAMBDA` parameter by the sheet-name
 * rule, every other name by the defined-name rule (`@utils/name-key`).
 */
export function definedNameListKey(name: string): string {
  return isLambdaParameter(name) ? sheetNameKey(name) : excelNameKey(name);
}

/** `name` without any `_xlfn.` / `_xlws.` prefix, upper-cased. */
export function bareFunctionName(name: string): string {
  let rest = name;
  if (rest.toLowerCase().startsWith(XLFN)) {
    rest = rest.slice(XLFN.length);
  }
  if (rest.toLowerCase().startsWith(XLWS)) {
    rest = rest.slice(XLWS.length);
  }
  return rest.toUpperCase();
}
