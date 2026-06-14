/**
 * Formula-syntax probe slot.
 *
 * ## What this is
 *
 * `DefinedNames` (in the excel module) needs a way to tell whether a raw
 * defined-name text like `OFFSET(Sheet1!$A$1,0,0,3,1)` is a parseable
 * formula or opaque content to preserve verbatim. The only authoritative
 * answer comes from the formula tokenizer + parser, which ship in the
 * `@cj-tech-master/excelts/formula` subpath and are an opt-in
 * dependency.
 *
 * This file is a tiny passive registration slot — a single `let` with
 * a getter/setter pair — carrying **no** formula-engine code. Importing
 * `DefinedNames` pulls in this slot (~0.5 KB minified) but never drags
 * in the tokenizer, parser, or evaluator.
 *
 * ## Classification semantics
 *
 * - **No probe installed** (default): `DefinedNames` classifies any
 *   non-range, non-wrapper text as **opaque**. The `rawText` is
 *   preserved so XLSX round-trip bytes are stable; the entry simply
 *   cannot be evaluated — which is correct, because no formula engine
 *   is available anyway.
 *
 * - **Probe installed** (after `installFormulaEngine()`): `DefinedNames`
 *   classifies strictly. Parseable text becomes `formula`, unparseable
 *   text becomes `opaque`.
 *
 * ## Why the slot lives in the formula module
 *
 * Module dependency direction: `excel` may import from `formula`
 * (layer 4 → layer 3), but `formula` must not import from `excel`. The
 * probe slot is a *formula-module concept* (it wraps a formula
 * tokenizer+parser) that excel consults; keeping it here preserves the
 * one-way dependency.
 *
 * ## Construction-time injection is also supported
 *
 * `Workbook` and `DefinedNames` each accept an explicit
 * `formulaSyntaxProbe` option. That is the preferred API for callers
 * who want deterministic, per-instance behaviour (e.g. a test that
 * needs a specific probe without touching process-global state). The
 * default-probe slot here is the convenience layer for
 * `installFormulaEngine()` — call it once at startup and every
 * subsequent `new Workbook()` picks up strict classification
 * automatically.
 *
 * ## Lookup timing
 *
 * `DefinedNames` reads the default probe lazily, at the moment `model`
 * is assigned (i.e. during XLSX parsing). This means the common
 * sequence
 *
 * ```ts
 * const wb = new Workbook();
 * installFormulaEngine();            // later
 * await Workbook.loadXlsx(wb, buffer);        // sees the installed probe
 * ```
 *
 * works correctly — the probe installed before `load()` is the one
 * used, regardless of when `Workbook` itself was constructed.
 */

export type SyntaxProbe = (text: string) => boolean;

let defaultProbe: SyntaxProbe | null = null;

/**
 * Install (or clear) the process-wide default syntax probe.
 *
 * Called from `installFormulaEngine()` in `./install`. Passing `null`
 * uninstalls the probe — symmetric with the formula-engine registry and
 * useful for tests that exercise the "no probe" classification path.
 */
export function setDefaultSyntaxProbe(probe: SyntaxProbe | null): void {
  defaultProbe = probe;
}

/**
 * Retrieve the currently-installed default probe, or `null` if none is
 * installed. Consumers (chiefly `DefinedNames`) should treat `null` as
 * a signal to use conservative (opaque) classification rather than
 * guessing.
 */
export function getDefaultSyntaxProbe(): SyntaxProbe | null {
  return defaultProbe;
}
