import { calculateFormulas } from "@excel/formula-adapter";
import { Cell, Workbook } from "@excel/index";

/**
 * Example: Functional / Standalone API
 *
 * Covers:
 * - `calculateFormulas(workbook)` — functional equivalent of
 *   `Workbook.calculateFormulas()`, works without calling
 *   `Formula.install()` first. This path is fully tree-shakeable:
 *   bundlers ship only the code paths reachable from the exports you
 *   reference.
 *
 * - `tokenize` / `parse` — pure syntax inspection. Use these for
 *   linters, formula migration tools, or static analysis that does not
 *   need to evaluate anything.
 */
import { Formula } from "../index";

// =============================================================================
// 1. Functional calculation — zero side effects
// =============================================================================
//
// Note we DO NOT call `Formula.install()`. The functional
// `calculateFormulas(workbook)` is self-contained and has no effect on
// `Workbook.calculateFormulas()` (which would still throw without install).

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Func");
Cell.setValue(ws, "A1", 10);
Cell.setValue(ws, "A2", 20);
Cell.setValue(ws, "A3", 30);
Cell.setValue(ws, "B1", { formula: "SUM(A1:A3)" });
Cell.setValue(ws, "B2", { formula: "AVERAGE(A1:A3)" });
Cell.setValue(ws, "B3", { formula: "MAX(A1:A3)" });

calculateFormulas(wb);

console.log("Functional calculateFormulas():");
console.log("  B1 =", Cell.getResult(ws, "B1")); // 60
console.log("  B2 =", Cell.getResult(ws, "B2")); // 20
console.log("  B3 =", Cell.getResult(ws, "B3")); // 30

// =============================================================================
// 2. Syntax inspection — tokenize + parse
// =============================================================================

const tokens = Formula.tokenize("SUM(A1:B10) + VLOOKUP(key, table, 2, FALSE)");
console.log("\ntokenize():");
for (const tok of tokens) {
  console.log(`  ${tok.type} ${"value" in tok ? JSON.stringify(tok.value) : ""}`);
}

// Structural check — can the parser handle it?
try {
  const ast = Formula.parse(tokens);
  console.log("\nparse():", ast.type, "root node — syntactically valid");
} catch (err) {
  console.log("\nparse() failed:", (err as Error).message);
}

// A deliberately malformed formula
try {
  Formula.parse(Formula.tokenize("SUM(A1:"));
  console.log("unexpected success");
} catch (err) {
  console.log("Bad formula rejected as expected:", (err as Error).message);
}
