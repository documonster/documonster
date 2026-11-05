/**
 * Argument errors named for the public call.
 *
 * The chart builders and validators are shared by several public names — `ChartClassic.*`,
 * `ChartEx.*` here, `Chart.*` in `documonster/excel`, the deprecated flat names — so the
 * code that finds a mistake cannot know which of them the caller wrote. The public entry
 * does: it runs the work through {@link namedCall}, and an argument error raised inside starts
 * its message with that call (`ChartEx.createModel: options.series must …`). The error object,
 * its class and its `field` are the ones that were thrown.
 */

import { checkWorkbook } from "@excel/core/workbook-core";
import { checkWorksheet } from "@excel/core/worksheet-core";
import { ExcelArgumentError } from "@excel/errors";
import { argumentChecks } from "@excel/utils/arguments";

/** A message that already starts with a public call — a dotted or `#` name, `Chart.copyTo: …`. */
const PUBLIC_PREFIX = /^[A-Za-z][\w]*[.#][\w.#]*: /;

/** A message that starts with an internal name — `createOptionsFromRows: …`. */
const INTERNAL_PREFIX = /^[A-Za-z]\w*: /;

/** Start the message of an argument error with `where`, unless it names a public call already. */
export function nameCall(err: unknown, where: string): unknown {
  if (!(err instanceof ExcelArgumentError) || PUBLIC_PREFIX.test(err.message)) {
    return err;
  }
  const before = err.message;
  const internal = INTERNAL_PREFIX.exec(before);
  err.message = `${where}: ${internal ? before.slice(internal[0].length) : before}`;
  if (typeof err.stack === "string" && err.stack.includes(before)) {
    err.stack = err.stack.replace(before, err.message);
  }
  return err;
}

/** Run `run`, naming any argument error it raises for the public call `where`. */
export function namedCall<T>(where: string, run: () => T): T {
  try {
    return run();
  } catch (err) {
    throw nameCall(err, where);
  }
}

/**
 * Refuse a missing argument, or one that is not an object, with an `ExcelArgumentError` naming `argument` — for the
 * entries whose argument goes straight into the work, where a wrong one used to surface as a `TypeError` from inside.
 */
export function checkObjectArgument(
  where: string,
  argument: string,
  value: unknown
): asserts value is Record<string, unknown> {
  const check = argumentChecks(where);
  check.required(argument, value);
  check.record(argument, value);
}

/** Refuse a missing argument, or one that is not a string, as {@link checkObjectArgument} refuses an object. */
export function checkStringArgument(
  where: string,
  argument: string,
  value: unknown
): asserts value is string {
  const check = argumentChecks(where);
  check.required(argument, value);
  check.string(argument, value);
}

/**
 * Check the workbook and `options.worksheet` a cache filler reads from, before anything is filled. Messages start with
 * `where`, and `field` is the public parameter's name (`workbook`, not the `wb` the shared handle check defaults to).
 */
export function checkFillArguments(where: string, workbook: unknown, options: unknown): void {
  namedCall(where, () => {
    checkWorkbook(workbook, "workbook");
    if (argumentChecks(where).record("options", options) && options.worksheet !== undefined) {
      checkWorksheet(options.worksheet, "options.worksheet");
    }
  });
}
