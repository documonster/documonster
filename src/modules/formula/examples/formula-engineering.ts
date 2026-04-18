/**
 * Example: Engineering Formulas
 *
 * Covers:
 * - Number base conversions (DEC2BIN, DEC2HEX, DEC2OCT, BIN2DEC, HEX2DEC, OCT2DEC)
 * - Bitwise operations (BITAND, BITOR, BITXOR, BITLSHIFT, BITRSHIFT)
 * - Complex numbers (COMPLEX, IMSUM, IMABS, IMARGUMENT)
 * - Special functions (ERF, BESSELJ)
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Eng");

// Base conversions
ws.getCell("A1").value = { formula: "DEC2BIN(42)" }; // "101010"
ws.getCell("A2").value = { formula: "DEC2HEX(255)" }; // "FF"
ws.getCell("A3").value = { formula: "DEC2OCT(64)" }; // "100"
ws.getCell("A4").value = { formula: 'BIN2DEC("101010")' }; // 42
ws.getCell("A5").value = { formula: 'HEX2DEC("FF")' }; // 255
ws.getCell("A6").value = { formula: 'OCT2DEC("777")' }; // 511

// Bitwise
ws.getCell("B1").value = { formula: "BITAND(12, 10)" }; // 8
ws.getCell("B2").value = { formula: "BITOR(12, 10)" }; // 14
ws.getCell("B3").value = { formula: "BITXOR(12, 10)" }; // 6
ws.getCell("B4").value = { formula: "BITLSHIFT(1, 4)" }; // 16
ws.getCell("B5").value = { formula: "BITRSHIFT(64, 2)" }; // 16

// Complex numbers (strings with "i" suffix)
ws.getCell("C1").value = { formula: "COMPLEX(3, 4)" }; // "3+4i"
ws.getCell("C2").value = { formula: 'IMSUM("3+4i", "1+2i")' }; // "4+6i"
ws.getCell("C3").value = { formula: 'IMABS("3+4i")' }; // 5
ws.getCell("C4").value = { formula: 'IMARGUMENT("1+1i")' }; // π/4 ≈ 0.7854

// Special functions
ws.getCell("D1").value = { formula: "ERF(1)" }; // 0.8427
ws.getCell("D2").value = { formula: "BESSELJ(0.5, 0)" };

wb.calculateFormulas();

for (const addr of [
  "A1",
  "A2",
  "A3",
  "A4",
  "A5",
  "A6",
  "B1",
  "B2",
  "B3",
  "B4",
  "B5",
  "C1",
  "C2",
  "C3",
  "C4",
  "D1",
  "D2"
]) {
  const c = ws.getCell(addr);
  console.log(`${addr}  ${String(c.formula).padEnd(32)}  = ${JSON.stringify(c.result)}`);
}
