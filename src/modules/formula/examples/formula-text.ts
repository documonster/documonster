/**
 * Example: Text Formulas
 *
 * Covers:
 * - Case & trim (UPPER, LOWER, PROPER, TRIM, CLEAN)
 * - Slicing (LEFT, RIGHT, MID, LEN)
 * - Search / replace (FIND, SEARCH, SUBSTITUTE, REPLACE)
 * - Concatenation (CONCAT, TEXTJOIN, `&` operator)
 * - Formatting (TEXT, VALUE, FIXED, DOLLAR)
 * - Character codes (CHAR, UNICHAR, CODE, UNICODE)
 * - Regex (REGEXTEST, REGEXEXTRACT, REGEXREPLACE)
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Text");

ws.getCell("A1").value = "  hello WORLD  ";
ws.getCell("A2").value = "name@example.com";
ws.getCell("A3").value = "Phone: +1-415-555-0123";

// Case & trim
ws.getCell("B1").value = { formula: "TRIM(A1)" }; // "hello WORLD"
ws.getCell("B2").value = { formula: "UPPER(TRIM(A1))" }; // "HELLO WORLD"
ws.getCell("B3").value = { formula: "PROPER(TRIM(A1))" }; // "Hello World"

// Slicing
ws.getCell("C1").value = { formula: "LEFT(A2, 4)" }; // "name"
ws.getCell("C2").value = { formula: 'MID(A2, FIND("@",A2)+1, 100)' }; // "example.com"
ws.getCell("C3").value = { formula: "LEN(A2)" }; // 16

// Search / substitute
ws.getCell("D1").value = { formula: 'SUBSTITUTE(A2,"example","acme")' };
ws.getCell("D2").value = { formula: 'FIND("@", A2)' }; // 5

// Concatenation
ws.getCell("E1").value = { formula: 'CONCAT("Hi, ", PROPER(TRIM(A1)), "!")' };
ws.getCell("E2").value = { formula: 'TEXTJOIN(", ", TRUE, B1, B2, B3)' };

// Formatting
ws.getCell("F1").value = { formula: 'TEXT(1234567.89, "#,##0.00")' }; // "1,234,567.89"
ws.getCell("F2").value = { formula: 'VALUE("42.5")' }; // 42.5

// Regex (Excel 365)
ws.getCell("G1").value = { formula: 'REGEXTEST(A2, "^[^@]+@[^@]+$")' }; // TRUE
ws.getCell("G2").value = { formula: 'REGEXEXTRACT(A3, "[0-9]{3}-[0-9]{3}-[0-9]{4}")' }; // "415-555-0123"

wb.calculateFormulas();

for (const addr of [
  "B1",
  "B2",
  "B3",
  "C1",
  "C2",
  "C3",
  "D1",
  "D2",
  "E1",
  "E2",
  "F1",
  "F2",
  "G1",
  "G2"
]) {
  const c = ws.getCell(addr);
  console.log(`${addr}  ${String(c.formula).padEnd(48)}  = ${JSON.stringify(c.result)}`);
}
