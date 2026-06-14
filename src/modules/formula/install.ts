/**
 * Opt-in engine installer.
 *
 * This file carries the imports that wire a concrete formula engine and
 * a tokenizer+parser probe into the excel host. Keep these imports
 * isolated in a separate module so callers who import only the
 * tokenizer, parser, or the functional `calculateFormulas` API from
 * `./index` never trigger this module — and its evaluator + function
 * registry — at bundle time.
 *
 * Two process-wide slots are populated:
 *
 * 1. The formula-engine slot (`@formula/host-registry`) — enables
 *    `Workbook.calculateFormulas()` and the PDF bridge's automatic
 *    recalculation.
 *
 * 2. The default syntax-probe slot (`@excel/default-syntax-probe`) —
 *    enables strict classification of defined-name text (formula vs.
 *    opaque) during XLSX load. Without this, `DefinedNames` preserves
 *    unrecognised text verbatim as opaque, which is correct for
 *    round-trip but cannot evaluate.
 *
 * Both slots accept `null` to uninstall. See `createFormulaSyntaxProbe`
 * for constructing a probe without touching the process-global slot
 * (useful for tests and for per-Workbook injection via
 * `new Workbook({ formulaSyntaxProbe })`).
 */

import { setDefaultSyntaxProbe, type SyntaxProbe } from "./default-syntax-probe";
import { parse } from "./syntax/parser";
import { tokenize } from "./syntax/tokenizer";

/**
 * Build a standalone formula-syntax probe backed by the real
 * tokenizer+parser. The returned function reports whether its argument
 * parses as a formula expression.
 *
 * Unlike {@link installFormulaEngine}, calling this does **not** touch
 * any process-global state. Inject the returned probe explicitly via
 * `new Workbook({ formulaSyntaxProbe })` or `new DefinedNames(probe)`
 * for deterministic, instance-scoped behaviour.
 */
export function createFormulaSyntaxProbe(): SyntaxProbe {
  return (text: string): boolean => {
    try {
      const tokens = tokenize(text);
      if (tokens.length === 0) {
        return false;
      }
      parse(tokens);
      return true;
    } catch {
      return false;
    }
  };
}

/**
 * Install the formula engine and the default syntax probe into the
 * excel host.
 *
 * After calling this once, `Workbook.calculateFormulas()` and the PDF
 * bridge's automatic recalculation run the full 433-function engine
 * instead of throwing / using stale cached values, and every
 * subsequently-loaded workbook's defined-name classification uses the
 * real tokenizer+parser instead of the conservative "opaque" fallback.
 *
 * Safe to call more than once — the registry accepts the last
 * registration.
 */
export function installFormulaEngine(): void {
  setDefaultSyntaxProbe(createFormulaSyntaxProbe());
}

/**
 * Uninstall both slots, restoring the cold-start state.
 *
 * Mainly useful for tests that need to exercise the "no engine" /
 * "no probe" classification path. In production, calling this is
 * rarely necessary — subsequent `installFormulaEngine()` calls simply
 * overwrite the previous registration.
 */
export function uninstallFormulaEngine(): void {
  setDefaultSyntaxProbe(null);
}
