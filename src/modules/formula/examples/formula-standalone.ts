/**
 * Example: Functional / Standalone API
 *
 * Covers:
 * - `calculateFormulas(workbook)` — functional equivalent of
 *   `Workbook.calculateFormulas()`, works without calling
 *   `installFormulaEngine()` first. This path is fully tree-shakeable:
 *   bundlers ship only the code paths reachable from the exports you
 *   reference.
 *
 * - `tokenize` / `parse` — pure syntax inspection. Use these for
 *   linters, formula migration tools, or static analysis that does not
 *   need to evaluate anything.
 */
import { Workbook } from "../../../index";
import { calculateFormulas, parse, tokenize } from "../index";

// =============================================================================
// 1. Functional calculation — zero side effects
// =============================================================================
//
// Note we DO NOT call `installFormulaEngine()`. The functional
// `calculateFormulas(workbook)` is self-contained and has no effect on
// `Workbook.calculateFormulas()` (which would still throw without install).

const wb = new Workbook();
const ws = wb.addWorksheet("Func");
ws.getCell("A1").value = 10;
ws.getCell("A2").value = 20;
ws.getCell("A3").value = 30;
ws.getCell("B1").value = { formula: "SUM(A1:A3)" };
ws.getCell("B2").value = { formula: "AVERAGE(A1:A3)" };
ws.getCell("B3").value = { formula: "MAX(A1:A3)" };

calculateFormulas(wb);

console.log("Functional calculateFormulas():");
console.log("  B1 =", ws.getCell("B1").result); // 60
console.log("  B2 =", ws.getCell("B2").result); // 20
console.log("  B3 =", ws.getCell("B3").result); // 30

// =============================================================================
// 2. Syntax inspection — tokenize + parse
// =============================================================================

const tokens = tokenize("SUM(A1:B10) + VLOOKUP(key, table, 2, FALSE)");
console.log("\ntokenize():");
for (const tok of tokens) {
  console.log(`  ${tok.type} ${"value" in tok ? JSON.stringify(tok.value) : ""}`);
}

// Structural check — can the parser handle it?
try {
  const ast = parse(tokens);
  console.log("\nparse():", ast.type, "root node — syntactically valid");
} catch (err) {
  console.log("\nparse() failed:", (err as Error).message);
}

// A deliberately malformed formula
try {
  parse(tokenize("SUM(A1:"));
  console.log("unexpected success");
} catch (err) {
  console.log("Bad formula rejected as expected:", (err as Error).message);
}
