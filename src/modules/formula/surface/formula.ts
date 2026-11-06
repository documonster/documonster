/**
 * `Formula` namespace surface — the formula engine's value API.
 *
 * `import { Formula } from "documonster/formula"` →
 *   `Formula.tokenize(src)`, `Formula.parse(src)`, `Formula.format(ast)`.
 *
 * Everything is used directly — there is no install / registration step.
 * Single flat namespace (formula is a single-purpose module). Re-exported
 * via `export * as Formula`, tree-shaken per-member on rolldown / rspack.
 */
import { printAst } from "@formula/syntax/print";

export { tokenize } from "@formula/syntax/tokenizer";
// Accepts source text or the tokens `tokenize` returned.
export { parseFormula as parse } from "@formula/syntax/parser";
// The inverse of `parse`, so the syntax layer is symmetric: anything that transforms,
// normalises or reconstructs a formula can hand back text instead of assembling it.
// Parentheses are minimal, so formatting normalises redundant ones away.
export { printAst as format } from "@formula/syntax/print";

/** @deprecated Use {@link format}. Removed in 1.0. */
export const print = printAst;
