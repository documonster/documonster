/**
 * Example: Database Formulas
 *
 * Covers:
 * - DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN
 * - DPRODUCT, DSTDEV, DVAR
 * - Criteria ranges (header + condition rows)
 * - Multi-column criteria (AND across columns, OR across rows)
 *
 * All D-functions share the same shape:
 *     DFUNC(database, field, criteria)
 *
 * - database: range including headers
 * - field: column header (string) or 1-based column index
 * - criteria: range including headers and one or more condition rows
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("DB");

// Database: A1:D7  (header row + 6 data rows)
// Region  Product  Sales  Units
ws.getCell("A1").value = "Region";
ws.getCell("B1").value = "Product";
ws.getCell("C1").value = "Sales";
ws.getCell("D1").value = "Units";
const data: [string, string, number, number][] = [
  ["East", "Widget", 120, 10],
  ["West", "Gadget", 95, 8],
  ["East", "Gadget", 140, 12],
  ["South", "Widget", 60, 5],
  ["West", "Widget", 110, 9],
  ["East", "Gadget", 75, 6]
];
data.forEach((row, i) => {
  ws.getCell(i + 2, 1).value = row[0];
  ws.getCell(i + 2, 2).value = row[1];
  ws.getCell(i + 2, 3).value = row[2];
  ws.getCell(i + 2, 4).value = row[3];
});

// Criteria — East + Gadget
ws.getCell("F1").value = "Region";
ws.getCell("G1").value = "Product";
ws.getCell("F2").value = "East";
ws.getCell("G2").value = "Gadget";

// DSUM — total sales of East-region Gadgets
ws.getCell("I1").value = { formula: 'DSUM(A1:D7, "Sales", F1:G2)' }; // 215
// DCOUNT — matching rows
ws.getCell("I2").value = { formula: 'DCOUNT(A1:D7, "Sales", F1:G2)' }; // 2
// DAVERAGE
ws.getCell("I3").value = { formula: 'DAVERAGE(A1:D7, "Sales", F1:G2)' }; // 107.5
// DMAX / DMIN
ws.getCell("I4").value = { formula: 'DMAX(A1:D7, "Sales", F1:G2)' }; // 140
ws.getCell("I5").value = { formula: 'DMIN(A1:D7, "Sales", F1:G2)' }; // 75

// OR-style criteria — East OR West
ws.getCell("F5").value = "Region";
ws.getCell("F6").value = "East";
ws.getCell("F7").value = "West";
ws.getCell("I7").value = { formula: 'DSUM(A1:D7, "Sales", F5:F7)' };
ws.getCell("I8").value = { formula: 'DCOUNTA(A1:D7, "Product", F5:F7)' };

wb.calculateFormulas();

for (const addr of ["I1", "I2", "I3", "I4", "I5", "I7", "I8"]) {
  const c = ws.getCell(addr);
  console.log(`${addr}  ${String(c.formula).padEnd(40)}  = ${JSON.stringify(c.result)}`);
}
