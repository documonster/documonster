/**
 * Date context shared between formula functions and the evaluator.
 *
 * Native function signatures (`NativeFn = (args: RuntimeValue[]) => RuntimeValue`)
 * have no access to the evaluation context, so date/time functions cannot
 * receive the workbook's `date1904` flag through a parameter. Rather than
 * plumb a new argument through 200+ function signatures, we store the flag
 * in a module-local variable. The evaluator sets it before evaluating any
 * formula in a calculation session.
 *
 * ## Thread safety
 *
 * JavaScript runtimes are single-threaded and a single calculation session
 * operates on exactly one workbook at a time, so a module-level variable is
 * safe in practice. If the engine is ever extended to support concurrent
 * calculations across different workbooks (e.g. via worker threads), each
 * worker will have its own module instance and remain correctly isolated.
 * Nested synchronous calls within the same session all share the same
 * workbook, so the flag remains correct throughout.
 */
let currentDate1904 = false;

/** Set the workbook-wide `date1904` mode for the current calculation session. */
export function setDate1904(v: boolean): void {
  currentDate1904 = v;
}

/** Read the workbook-wide `date1904` mode. */
export function isDate1904(): boolean {
  return currentDate1904;
}
