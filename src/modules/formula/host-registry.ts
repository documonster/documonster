/**
 * Formula Engine Registry
 *
 * A tiny indirection layer that lets `Workbook.calculateFormulas()` and
 * the PDF bridge invoke a full formula engine **only if the host
 * application has opted in** by calling `installFormulaEngine()` from
 * `@cj-tech-master/excelts/formula`.
 *
 * ## Why a registry
 *
 * The formula engine ships 433 Excel functions, a tokenizer, a parser,
 * a compiler, an evaluator, a dependency graph and a spill materialiser.
 * Minified this is ~200 KB. Most Workbook users only need to read /
 * write XLSX files and let Excel recalculate on open; pulling the
 * engine into their bundle unconditionally would be a large, invisible
 * cost.
 *
 * This file is the single point of indirection. Importing `Workbook`
 * pulls it in (~2.5 KB) but does NOT pull in the engine; the engine's
 * implementation only arrives once someone calls `installFormulaEngine()`,
 * which bundlers can then resolve to the subpath import graph.
 *
 * If `calculateFormulas()` is called without `installFormulaEngine()`
 * having run, a clear error is thrown telling the developer how to fix it.
 */

import type { WorkbookLike } from "./materialize/types";

/** The function shape {@link installFormulaEngine} registers here. */
export type FormulaEngine = (workbook: WorkbookLike) => void;

let installed: FormulaEngine | null = null;

/**
 * Install a formula engine implementation. Called from
 * `installFormulaEngine()` in the `@cj-tech-master/excelts/formula`
 * subpath.
 *
 * Re-installing is allowed — the last registration wins. This keeps
 * hot-reload tooling and tests simple. Passing `null` uninstalls the
 * engine, which is useful for tests that exercise the
 * "no engine" error path.
 */
export function registerFormulaEngine(engine: FormulaEngine | null): void {
  installed = engine;
}

/** Returns true when a formula engine has been installed. */
export function hasFormulaEngine(): boolean {
  return installed !== null;
}

/**
 * Invoke the registered engine on `workbook`. Throws a descriptive error
 * if no engine has been installed.
 */
export function invokeFormulaEngine(workbook: WorkbookLike): void {
  if (!installed) {
    throw new Error(
      "No formula engine is installed. " +
        "Call `installFormulaEngine()` from `@cj-tech-master/excelts/formula` " +
        "once at startup to enable `Workbook.calculateFormulas()` and " +
        "automatic recalculation during `excelToPdf()`."
    );
  }
  installed(workbook);
}

/**
 * Like {@link invokeFormulaEngine} but returns silently when no engine is
 * installed. Used by `excelToPdf()` so that PDF export still works for
 * workbooks whose cached formula results are already up to date (the
 * common case when the XLSX was saved by Excel itself).
 */
export function tryInvokeFormulaEngine(workbook: WorkbookLike): boolean {
  if (!installed) {
    return false;
  }
  installed(workbook);
  return true;
}
