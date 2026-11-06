import { Cell, Workbook } from "@excel/index";

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "A1 Notation");

// A1 addressing
Cell.setValue(ws, "A1", "A1");
Cell.setValue(ws, "B2", "B2");

// Row/column addressing (1-based)
Cell.setValue(ws, { row: 1, col: 3 }, "C1");
Cell.setValue(ws, { row: 3, col: 1 }, "A3");

// Verify we can read values back with either style
console.log("A1 =", Cell.getValue(ws, "A1"));
console.log("B2 =", Cell.getValue(ws, { row: 2, col: 2 }));
console.log("C1 =", Cell.getValue(ws, "C1"));
console.log("A3 =", Cell.getValue(ws, { row: 3, col: 1 }));

console.log("Done.");
