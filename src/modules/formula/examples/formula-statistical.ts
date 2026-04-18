/**
 * Example: Statistical Formulas
 *
 * Covers:
 * - Basic descriptive (AVERAGE, MEDIAN, MODE, STDEV, VAR)
 * - Extremes (MIN, MAX, LARGE, SMALL)
 * - Conditional aggregates (COUNTIF, COUNTIFS, SUMIF, SUMIFS, AVERAGEIF)
 * - Rank & percentile (RANK, PERCENTILE, QUARTILE, PERCENTRANK)
 * - Correlation & regression (CORREL, COVAR, SLOPE, INTERCEPT, LINEST)
 * - Distributions (NORM.DIST, NORM.INV, T.TEST, CHISQ.TEST)
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Stats");

// Sample dataset
const nums = [12, 15, 14, 10, 18, 22, 19, 24, 13, 21];
for (let i = 0; i < nums.length; i++) {
  ws.getCell(i + 1, 1).value = nums[i]; // A1:A10
}
// Second dataset for correlation
const ys = [30, 33, 31, 29, 40, 42, 39, 45, 32, 44];
for (let i = 0; i < ys.length; i++) {
  ws.getCell(i + 1, 2).value = ys[i]; // B1:B10
}

// Descriptive
ws.getCell("D1").value = { formula: "AVERAGE(A1:A10)" };
ws.getCell("D2").value = { formula: "MEDIAN(A1:A10)" };
ws.getCell("D3").value = { formula: "MODE(A1:A10)" };
ws.getCell("D4").value = { formula: "STDEV.S(A1:A10)" };
ws.getCell("D5").value = { formula: "VAR.S(A1:A10)" };

// Extremes
ws.getCell("E1").value = { formula: "MIN(A1:A10)" };
ws.getCell("E2").value = { formula: "MAX(A1:A10)" };
ws.getCell("E3").value = { formula: "LARGE(A1:A10, 2)" }; // 2nd largest
ws.getCell("E4").value = { formula: "SMALL(A1:A10, 3)" }; // 3rd smallest

// Conditional — count / sum values ≥ 18
ws.getCell("F1").value = { formula: 'COUNTIF(A1:A10, ">=18")' };
ws.getCell("F2").value = { formula: 'SUMIF(A1:A10, ">=18")' };
ws.getCell("F3").value = { formula: 'AVERAGEIF(A1:A10, ">=18")' };
// COUNTIFS with two criteria
ws.getCell("F4").value = { formula: 'COUNTIFS(A1:A10, ">=14", B1:B10, ">=35")' };

// Rank & percentile
ws.getCell("G1").value = { formula: "RANK(21, A1:A10)" }; // descending rank
ws.getCell("G2").value = { formula: "PERCENTILE.INC(A1:A10, 0.9)" };
ws.getCell("G3").value = { formula: "QUARTILE.INC(A1:A10, 3)" }; // Q3
ws.getCell("G4").value = { formula: "PERCENTRANK.INC(A1:A10, 18)" };

// Correlation & simple regression
ws.getCell("H1").value = { formula: "CORREL(A1:A10, B1:B10)" };
ws.getCell("H2").value = { formula: "SLOPE(B1:B10, A1:A10)" };
ws.getCell("H3").value = { formula: "INTERCEPT(B1:B10, A1:A10)" };
ws.getCell("H4").value = { formula: "FORECAST(16, B1:B10, A1:A10)" };

// Distributions
ws.getCell("I1").value = { formula: "NORM.DIST(1, 0, 1, TRUE)" }; // ≈ 0.8413
ws.getCell("I2").value = { formula: "NORM.INV(0.975, 0, 1)" }; // ≈ 1.96

wb.calculateFormulas();

for (const row of [1, 2, 3, 4, 5]) {
  for (const col of ["D", "E", "F", "G", "H", "I"]) {
    const addr = `${col}${row}`;
    const c = ws.getCell(addr);
    if (c.formula) {
      console.log(`${addr}  ${String(c.formula).padEnd(40)}  = ${JSON.stringify(c.result)}`);
    }
  }
}
