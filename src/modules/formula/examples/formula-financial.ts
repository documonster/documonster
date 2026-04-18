/**
 * Example: Financial Formulas
 *
 * Covers:
 * - Loan / mortgage (PMT, IPMT, PPMT, NPER, RATE)
 * - Time value of money (PV, FV)
 * - Investment evaluation (NPV, IRR, XNPV, XIRR, MIRR)
 * - Depreciation (SLN, DB, DDB, SYD)
 * - Bonds (PRICE, YIELD, DURATION, ACCRINT)
 */
import { Workbook } from "../../../index";
import { installFormulaEngine } from "../index";

installFormulaEngine();

const wb = new Workbook();
const ws = wb.addWorksheet("Finance");

// Loan assumptions
ws.getCell("A1").value = 300000; // principal
ws.getCell("A2").value = 0.065 / 12; // monthly rate (6.5% annual)
ws.getCell("A3").value = 30 * 12; // periods (30 years)

// Monthly payment — negative of PMT because of cashflow convention
ws.getCell("B1").value = { formula: "-PMT(A2, A3, A1)" };
ws.getCell("B2").value = { formula: "IPMT(A2, 1, A3, A1)" }; // interest, period 1
ws.getCell("B3").value = { formula: "PPMT(A2, 1, A3, A1)" }; // principal, period 1

// How long until a $500k goal at 8% earning $2k/month?
ws.getCell("B4").value = { formula: "NPER(0.08/12, -2000, 0, 500000)" };
// What rate makes a $10k loan turn into $15k over 5 yr?
ws.getCell("B5").value = { formula: "RATE(5, 0, -10000, 15000)" };

// Investment — cashflows in C1:C5
ws.getCell("C1").value = -1000;
ws.getCell("C2").value = 300;
ws.getCell("C3").value = 400;
ws.getCell("C4").value = 500;
ws.getCell("C5").value = 200;

// NPV applies to future cashflows; add the initial outlay separately
ws.getCell("D1").value = { formula: "NPV(0.1, C2:C5) + C1" };
ws.getCell("D2").value = { formula: "IRR(C1:C5)" };

// Time value — PV of $1000 received in 10 years at 5%
ws.getCell("D3").value = { formula: "PV(0.05, 10, 0, -1000)" };
ws.getCell("D4").value = { formula: "FV(0.05, 10, -100)" }; // $100/yr deposits

// Depreciation — $10k asset over 5 years, $1k salvage
ws.getCell("E1").value = { formula: "SLN(10000, 1000, 5)" }; // straight line
ws.getCell("E2").value = { formula: "DDB(10000, 1000, 5, 1)" }; // double declining
ws.getCell("E3").value = { formula: "SYD(10000, 1000, 5, 1)" }; // sum-of-years

wb.calculateFormulas();

for (const addr of ["B1", "B2", "B3", "B4", "B5", "D1", "D2", "D3", "D4", "E1", "E2", "E3"]) {
  const c = ws.getCell(addr);
  console.log(`${addr}  ${String(c.formula).padEnd(36)}  = ${JSON.stringify(c.result)}`);
}
