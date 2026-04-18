/**
 * Example: Lookup & Reference Formulas
 *
 * Covers:
 * - VLOOKUP / HLOOKUP (legacy)
 * - INDEX / MATCH (flexible classic)
 * - XLOOKUP / XMATCH (modern)
 * - LOOKUP (vector form)
 * - OFFSET (dynamic range anchoring)
 * - INDIRECT (string-built references)
 * - ROW / COLUMN / ADDRESS (positional metadata)
 * - ROWS / COLUMNS (dimension of a range)
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Lookup");

// Price table
// A      B       C
// SKU    Name    Price
// A001   Widget  9.99
// A002   Gadget  14.50
// A003   Gizmo   3.25
ws.getCell("A1").value = "SKU";
ws.getCell("B1").value = "Name";
ws.getCell("C1").value = "Price";
ws.getCell("A2").value = "A001";
ws.getCell("B2").value = "Widget";
ws.getCell("C2").value = 9.99;
ws.getCell("A3").value = "A002";
ws.getCell("B3").value = "Gadget";
ws.getCell("C3").value = 14.5;
ws.getCell("A4").value = "A003";
ws.getCell("B4").value = "Gizmo";
ws.getCell("C4").value = 3.25;

// Query
ws.getCell("E1").value = "A002";

// Classic VLOOKUP — get price by SKU
ws.getCell("F1").value = { formula: "VLOOKUP(E1, A2:C4, 3, FALSE)" }; // 14.5

// INDEX + MATCH — the same, but right-to-left safe
ws.getCell("F2").value = { formula: "INDEX(B2:B4, MATCH(E1, A2:A4, 0))" }; // "Gadget"

// XLOOKUP — modern, supports not-found default
ws.getCell("F3").value = { formula: 'XLOOKUP(E1, A2:A4, C2:C4, "n/a")' }; // 14.5
ws.getCell("F4").value = { formula: 'XLOOKUP("ZZZ", A2:A4, C2:C4, "n/a")' }; // "n/a"

// XMATCH
ws.getCell("F5").value = { formula: "XMATCH(E1, A2:A4)" }; // 2

// OFFSET — dynamic anchor (2nd row, 2nd col, 1x1)
ws.getCell("F6").value = { formula: "OFFSET(A1, 2, 1)" }; // "Gadget"

// INDIRECT — build a reference from a string (issue #140 style)
ws.getCell("F7").value = { formula: 'SUM(C2:INDIRECT("C" & ROWS(A1:A4)))' }; // 27.74

// Metadata
ws.getCell("G1").value = { formula: "ROW(E1)" }; // 1
ws.getCell("G2").value = { formula: "COLUMN(E1)" }; // 5
ws.getCell("G3").value = { formula: "ADDRESS(G1, G2)" }; // "$E$1"
ws.getCell("G4").value = { formula: "ROWS(A2:C4)" }; // 3
ws.getCell("G5").value = { formula: "COLUMNS(A2:C4)" }; // 3

wb.calculateFormulas();

for (const addr of ["F1", "F2", "F3", "F4", "F5", "F6", "F7", "G1", "G2", "G3", "G4", "G5"]) {
  const c = ws.getCell(addr);
  console.log(`${addr}  ${String(c.formula).padEnd(48)}  = ${JSON.stringify(c.result)}`);
}
