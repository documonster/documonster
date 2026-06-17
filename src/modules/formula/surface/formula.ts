/**
 * `Formula` namespace surface — the formula engine's value API.
 *
 * `import { Formula } from "documonster/formula"` →
 *   `Formula.calculate(workbook)`, `Formula.tokenize(src)`,
 *   `Formula.parse(tokens)`, `Formula.createSyntaxProbe()`.
 *
 * Single flat namespace (formula is a single-purpose module). Re-exported
 * via `export * as Formula`, tree-shaken per-member on rolldown / rspack.
 */
export { calculateFormulas as calculate } from "../integration/calculate-formulas";
export { tokenize } from "../syntax/tokenizer";
export { parse } from "../syntax/parser";
export {
  createFormulaSyntaxProbe as createSyntaxProbe,
  installFormulaEngine as install,
  uninstallFormulaEngine as uninstall
} from "../install";
