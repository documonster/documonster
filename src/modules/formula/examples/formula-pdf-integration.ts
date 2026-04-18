/**
 * Example: Formula + PDF Integration
 *
 * Covers:
 * - `installFormulaEngine()` makes `excelToPdf()` automatically
 *   recalculate stale formula results before rendering.
 * - Without install, `excelToPdf()` silently falls back to the cached
 *   results saved in the XLSX (safe default for files last opened in
 *   Excel itself).
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook, excelToPdf } from "../../../index";
import { installFormulaEngine } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/formula-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// Enable automatic recalculation inside excelToPdf().
installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Invoice");

ws.columns = [
  { header: "Item", key: "item", width: 24 },
  { header: "Qty", key: "qty", width: 8 },
  { header: "Price", key: "price", width: 12 },
  { header: "Subtotal", key: "subtotal", width: 14 }
];
ws.addRows([
  { item: "Widget A", qty: 3, price: 9.99 },
  { item: "Gadget B", qty: 2, price: 14.5 },
  { item: "Gizmo C", qty: 5, price: 3.25 }
]);

// Live subtotal formula per row — results are stale (zero) until calc runs
for (let r = 2; r <= 4; r++) {
  ws.getCell(`D${r}`).value = { formula: `B${r}*C${r}` };
}
// Grand total
ws.getCell("D5").value = { formula: "SUM(D2:D4)" };
ws.getCell("C5").value = "Total";

// `excelToPdf` calls `tryInvokeFormulaEngine(workbook)` internally;
// because `installFormulaEngine` was called above, subtotals and the
// grand total are computed fresh right before rendering.
const pdf = await excelToPdf(wb, {
  title: "Invoice (live formula results)",
  showGridLines: true
});
fs.writeFileSync(path.join(outDir, "formula-pdf-integration.pdf"), pdf);
console.log("Wrote tmp/formula-examples/formula-pdf-integration.pdf");
console.log(
  "  Subtotals:",
  ws.getCell("D2").result,
  ws.getCell("D3").result,
  ws.getCell("D4").result
);
console.log("  Grand total:", ws.getCell("D5").result);
