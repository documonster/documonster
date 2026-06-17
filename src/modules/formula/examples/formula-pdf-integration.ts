/**
 * Example: Formula + PDF Integration
 *
 * Covers:
 * - `Formula.install()` makes `Pdf.fromExcel()` automatically
 *   recalculate stale formula results before rendering.
 * - Without install, `Pdf.fromExcel()` silently falls back to the cached
 *   results saved in the XLSX (safe default for files last opened in
 *   Excel itself).
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Cell, Workbook, Worksheet } from "@excel/index";

import { Pdf } from "../../../index";
import { Formula } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/formula-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// Enable automatic recalculation inside Pdf.fromExcel().
Formula.install();

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Invoice");

Worksheet.setColumns(ws, [
  { header: "Item", key: "item", width: 24 },
  { header: "Qty", key: "qty", width: 8 },
  { header: "Price", key: "price", width: 12 },
  { header: "Subtotal", key: "subtotal", width: 14 }
]);
Worksheet.addRows(ws, [
  { item: "Widget A", qty: 3, price: 9.99 },
  { item: "Gadget B", qty: 2, price: 14.5 },
  { item: "Gizmo C", qty: 5, price: 3.25 }
]);

// Live subtotal formula per row — results are stale (zero) until calc runs
for (let r = 2; r <= 4; r++) {
  Cell.setValue(ws, `D${r}`, { formula: `B${r}*C${r}` });
}
// Grand total
Cell.setValue(ws, "D5", { formula: "SUM(D2:D4)" });
Cell.setValue(ws, "C5", "Total");

// `excelToPdf` calls `tryInvokeFormulaEngine(workbook)` internally;
// because `installFormulaEngine` was called above, subtotals and the
// grand total are computed fresh right before rendering.
const pdf = await Pdf.fromExcel(wb, {
  title: "Invoice (live formula results)",
  showGridLines: true
});
fs.writeFileSync(path.join(outDir, "formula-pdf-integration.pdf"), pdf);
console.log("Wrote tmp/formula-examples/formula-pdf-integration.pdf");
console.log(
  "  Subtotals:",
  Cell.getResult(ws, "D2"),
  Cell.getResult(ws, "D3"),
  Cell.getResult(ws, "D4")
);
console.log("  Grand total:", Cell.getResult(ws, "D5"));
