/**
 * Opt-in syntax-probe installer.
 *
 * This file carries the imports that wire a concrete tokenizer+parser probe
 * into the process-global default-syntax-probe slot. Keep these imports
 * isolated in a separate module so callers who import only the tokenizer,
 * parser, or the functional `calculateFormulas` API from `./index` never
 * trigger this module at bundle time.
 *
 * One process-wide slot is populated:
 *
 *   The default syntax-probe slot (`./default-syntax-probe`) — enables strict
 *   classification of defined-name text (formula vs. opaque) during XLSX load.
 *   Without this, `DefinedNames` preserves unrecognised text verbatim as
 *   opaque, which is correct for round-trip but cannot evaluate.
 *
 * The slot accepts `null` to uninstall. See `createFormulaSyntaxProbe` for
 * constructing a probe without touching the process-global slot (useful for
 * tests and for per-Workbook injection via
 * `new Workbook({ formulaSyntaxProbe })`).
 *
 * Note: there is NO formula-engine "host registry". `Workbook.calculateFormulas`
 * and the PDF bridge call the functional `calculateFormulas` API directly —
 * no install step or global engine slot is involved.
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
 * Install the default syntax probe (backed by the real tokenizer+parser)
 * into the process-global slot.
 *
 * After calling this once, every subsequently-loaded workbook's defined-name
 * classification uses the real tokenizer+parser instead of the conservative
 * "opaque" fallback. This does NOT wire any formula engine: `calculateFormulas`
 * already works standalone via the functional API.
 *
 * Safe to call more than once — the slot accepts the last registration.
 */
export function installFormulaEngine(): void {
  setDefaultSyntaxProbe(createFormulaSyntaxProbe());
}

/**
 * Uninstall the syntax-probe slot, restoring the cold-start state.
 *
 * Mainly useful for tests that need to exercise the "no probe" classification
 * path. In production, calling this is rarely necessary — subsequent
 * `installFormulaEngine()` calls simply overwrite the previous registration.
 */
export function uninstallFormulaEngine(): void {
  setDefaultSyntaxProbe(null);
}
