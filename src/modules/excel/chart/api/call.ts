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

import { ExcelArgumentError } from "@excel/errors";

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
