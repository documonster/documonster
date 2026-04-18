/**
 * Example: Logical & Conditional Formulas
 *
 * Covers:
 * - IF / nested IF / IFS
 * - AND / OR / NOT / XOR
 * - IFERROR / IFNA / ISERROR
 * - SWITCH
 * - CHOOSE
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Logic");

ws.getCell("A1").value = 85;
ws.getCell("A2").value = "OK";
ws.getCell("A3").value = 0;
ws.getCell("A4").value = 42;

// Basic IF
ws.getCell("B1").value = { formula: 'IF(A1>=60, "pass", "fail")' };

// Nested IF → grade letter
ws.getCell("B2").value = {
  formula: 'IF(A1>=90,"A",IF(A1>=80,"B",IF(A1>=70,"C",IF(A1>=60,"D","F"))))'
};

// Same idea with IFS — cleaner
ws.getCell("B3").value = {
  formula: 'IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C",A1>=60,"D",TRUE,"F")'
};

// AND / OR / NOT
ws.getCell("C1").value = { formula: "AND(A1>=60, A1<=100)" }; // TRUE
ws.getCell("C2").value = { formula: 'OR(A2="OK", A2="YES")' }; // TRUE
ws.getCell("C3").value = { formula: "NOT(A3)" }; // TRUE  (A3 is 0 → FALSE)

// IFERROR / IFNA guard against #DIV/0!, #N/A, etc.
ws.getCell("D1").value = { formula: 'IFERROR(A4/A3, "n/a")' }; // "n/a"
ws.getCell("D2").value = { formula: 'IFNA(VLOOKUP("x", A1:A4, 1, FALSE), "not found")' };

// SWITCH — dispatch on equality
ws.getCell("E1").value = {
  formula: 'SWITCH(A2, "OK", 1, "WARN", 2, "ERR", 3, 0)'
}; // 1

// CHOOSE — index-based dispatch (1-based)
ws.getCell("E2").value = { formula: 'CHOOSE(2, "low", "mid", "high")' }; // "mid"

wb.calculateFormulas();

for (const addr of ["B1", "B2", "B3", "C1", "C2", "C3", "D1", "D2", "E1", "E2"]) {
  const c = ws.getCell(addr);
  console.log(`${addr}  ${String(c.formula).padEnd(60)}  = ${JSON.stringify(c.result)}`);
}
