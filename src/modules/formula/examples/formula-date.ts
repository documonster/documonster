/**
 * Example: Date & Time Formulas
 *
 * Covers:
 * - Current date/time (TODAY, NOW)
 * - Construct / extract (DATE, TIME, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND)
 * - Weekday / week number (WEEKDAY, WEEKNUM, ISOWEEKNUM)
 * - Duration (DATEDIF, DAYS, DAYS360, EDATE, EOMONTH)
 * - Business days (NETWORKDAYS, WORKDAY, NETWORKDAYS.INTL)
 * - Formatting via TEXT
 *
 * Note: Excel stores dates as serial numbers (days since 1900-01-01). The
 * engine works with those serials; set `cell.numFmt` when reading back if
 * you need a formatted display.
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Dates");

// Fixed reference points so the example is deterministic
ws.getCell("A1").value = { formula: "DATE(2026, 4, 18)" };
ws.getCell("A2").value = { formula: "DATE(2026, 12, 25)" };
ws.getCell("A3").value = { formula: "TIME(14, 30, 0)" };

// Extract components
ws.getCell("B1").value = { formula: "YEAR(A1)" }; // 2026
ws.getCell("B2").value = { formula: "MONTH(A1)" }; // 4
ws.getCell("B3").value = { formula: "DAY(A1)" }; // 18
ws.getCell("B4").value = { formula: "HOUR(A3)" }; // 14

// Weekday / week-of-year
ws.getCell("C1").value = { formula: "WEEKDAY(A1, 2)" }; // 1..7, Mon=1
ws.getCell("C2").value = { formula: "WEEKNUM(A1)" };
ws.getCell("C3").value = { formula: "ISOWEEKNUM(A1)" };

// Duration
ws.getCell("D1").value = { formula: 'DATEDIF(A1, A2, "D")' }; // 251 days
ws.getCell("D2").value = { formula: "DAYS(A2, A1)" }; // 251
ws.getCell("D3").value = { formula: "EOMONTH(A1, 0)" }; // last day of April 2026
ws.getCell("D4").value = { formula: "EDATE(A1, 6)" }; // 6 months later

// Business days
ws.getCell("E1").value = { formula: "NETWORKDAYS(A1, A2)" };
ws.getCell("E2").value = { formula: "WORKDAY(A1, 10)" }; // 10 business days out

// Formatting (YYYY-MM-DD)
ws.getCell("F1").value = { formula: 'TEXT(A1, "yyyy-mm-dd")' };

wb.calculateFormulas();

for (const addr of [
  "B1",
  "B2",
  "B3",
  "B4",
  "C1",
  "C2",
  "C3",
  "D1",
  "D2",
  "D3",
  "D4",
  "E1",
  "E2",
  "F1"
]) {
  const c = ws.getCell(addr);
  console.log(`${addr}  ${String(c.formula).padEnd(32)}  = ${JSON.stringify(c.result)}`);
}
