/**
 * Example: Dynamic Array Formulas (Excel 365)
 *
 * Covers:
 * - Spill semantics — one formula, many output cells
 * - FILTER / SORT / SORTBY / UNIQUE
 * - SEQUENCE / RANDARRAY
 * - TAKE / DROP / CHOOSECOLS / CHOOSEROWS
 * - VSTACK / HSTACK
 * - TEXTSPLIT / TEXTJOIN (paired)
 * - Higher-order (MAP, REDUCE, SCAN, LAMBDA)
 *
 * Dynamic arrays write their full rectangle into the sheet. Inspect the
 * `ghost` cells around the source formula to see the spilled values.
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Dynamic");

// Source data (A1:C6)
// Region   Product  Sales
ws.getCell("A1").value = "Region";
ws.getCell("B1").value = "Product";
ws.getCell("C1").value = "Sales";
const rows: [string, string, number][] = [
  ["East", "Widget", 120],
  ["West", "Gadget", 95],
  ["East", "Gadget", 140],
  ["South", "Widget", 60],
  ["West", "Widget", 110]
];
rows.forEach((row, i) => {
  ws.getCell(i + 2, 1).value = row[0];
  ws.getCell(i + 2, 2).value = row[1];
  ws.getCell(i + 2, 3).value = row[2];
});

// FILTER — rows where Sales ≥ 100
ws.getCell("E1").value = { formula: "FILTER(A2:C6, C2:C6>=100)" };

// SORT — by Sales descending
ws.getCell("E10").value = { formula: "SORT(A2:C6, 3, -1)" };

// UNIQUE — distinct regions
ws.getCell("J1").value = { formula: "UNIQUE(A2:A6)" };

// SEQUENCE — 1..10
ws.getCell("L1").value = { formula: "SEQUENCE(10)" };

// TAKE — first 2 rows, last 2 cols
ws.getCell("N1").value = { formula: "TAKE(A2:C6, 2, -2)" };

// VSTACK — combine uniques from two sources
ws.getCell("P1").value = { formula: "VSTACK(UNIQUE(A2:A6), UNIQUE(B2:B6))" };

// TEXTSPLIT — break a CSV row into cells (row-wise)
ws.getCell("R1").value = "Alice,Bob,Carol,Dave";
ws.getCell("R2").value = { formula: 'TEXTSPLIT(R1, ",")' };

// LAMBDA / REDUCE — sum of squares of 1..10
ws.getCell("T1").value = {
  formula: "REDUCE(0, SEQUENCE(10), LAMBDA(acc,x, acc + x*x))"
}; // 385

// MAP — double each value in a range
ws.getCell("T3").value = {
  formula: "MAP(SEQUENCE(5), LAMBDA(x, x*2))"
};

// SCAN — running total 1..5
ws.getCell("T5").value = {
  formula: "SCAN(0, SEQUENCE(5), LAMBDA(acc,x, acc + x))"
};

wb.calculateFormulas();

// Dynamic array sources print their scalar top-left; the full spill
// lives in surrounding cells.
for (const addr of ["E1", "E10", "J1", "L1", "T1"]) {
  const c = ws.getCell(addr);
  console.log(`${addr}  ${String(c.formula).padEnd(52)}  top-left = ${JSON.stringify(c.result)}`);
}

// Demonstrate the spill — FILTER writes its full rectangle into E1:G{n}.
// Source cell's `result` holds the top-left scalar; ghost cells hold the
// spilled values in their `value` (not `result`).
console.log("\nFILTER spill E1:G?");
for (let rn = 1; rn <= 6; rn++) {
  const row: unknown[] = [];
  for (let cn = 5; cn <= 7; cn++) {
    const cell = ws.findCell(rn, cn);
    if (!cell) {
      row.push(null);
      continue;
    }
    // For the source cell, the spilled value is carried in `result`;
    // ghost cells carry it in `value`.
    row.push(rn === 1 && cn === 5 ? cell.result : cell.value);
  }
  if (row.some(v => v !== null)) {
    console.log(`  row ${rn}: ${JSON.stringify(row)}`);
  }
}
