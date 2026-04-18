/**
 * Example: Math & Trigonometry Formulas
 *
 * Covers:
 * - Basic arithmetic formulas (SUM, PRODUCT, ABS)
 * - Rounding (ROUND, ROUNDUP, ROUNDDOWN, INT, CEILING, FLOOR)
 * - Trigonometry (SIN, COS, TAN, ATAN2, RADIANS, DEGREES, PI)
 * - Power & logarithms (POWER, SQRT, EXP, LN, LOG, LOG10)
 * - Matrix math (MMULT, MDETERM, MINVERSE, TRANSPOSE)
 * - Modular arithmetic (MOD, QUOTIENT)
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

// One-time install so `workbook.calculateFormulas()` wires to the engine.
installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Math");

// Inputs
ws.getCell("A1").value = 10;
ws.getCell("A2").value = 20;
ws.getCell("A3").value = 30;
ws.getCell("B1").value = -5.678;

// Aggregates and basic arithmetic
ws.getCell("C1").value = { formula: "SUM(A1:A3)" }; // 60
ws.getCell("C2").value = { formula: "PRODUCT(A1:A3)" }; // 6000
ws.getCell("C3").value = { formula: "ABS(B1)" }; // 5.678

// Rounding family
ws.getCell("D1").value = { formula: "ROUND(B1, 2)" }; // -5.68
ws.getCell("D2").value = { formula: "ROUNDUP(B1, 0)" }; // -6
ws.getCell("D3").value = { formula: "INT(B1)" }; // -6
ws.getCell("D4").value = { formula: "CEILING(4.3, 1)" }; // 5

// Trigonometry — 30° in radians, then back
ws.getCell("E1").value = { formula: "RADIANS(30)" };
ws.getCell("E2").value = { formula: "SIN(E1)" }; // 0.5
ws.getCell("E3").value = { formula: "DEGREES(ATAN2(1, 1))" }; // 45

// Power & log
ws.getCell("F1").value = { formula: "POWER(2, 10)" }; // 1024
ws.getCell("F2").value = { formula: "SQRT(144)" }; // 12
ws.getCell("F3").value = { formula: "LOG(1000, 10)" }; // 3

// Modular
ws.getCell("G1").value = { formula: "MOD(17, 5)" }; // 2
ws.getCell("G2").value = { formula: "QUOTIENT(17, 5)" }; // 3

// Matrix — populate a 2x2 and compute determinant / inverse via MMULT
ws.getCell("H1").value = 1;
ws.getCell("I1").value = 2;
ws.getCell("H2").value = 3;
ws.getCell("I2").value = 4;
ws.getCell("J1").value = { formula: "MDETERM(H1:I2)" }; // -2

wb.calculateFormulas();

// Report — single pass, tabulated so you can eyeball the results.
const cells = [
  "C1",
  "C2",
  "C3",
  "D1",
  "D2",
  "D3",
  "D4",
  "E1",
  "E2",
  "E3",
  "F1",
  "F2",
  "F3",
  "G1",
  "G2",
  "J1"
];
for (const addr of cells) {
  const cell = ws.getCell(addr);
  console.log(`${addr}  ${String(cell.formula).padEnd(28)}  = ${cell.result}`);
}
