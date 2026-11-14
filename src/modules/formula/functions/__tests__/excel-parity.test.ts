/**
 * Formulas whose results differed from Excel's, checked end to end through the workbook
 * calculator so argument passing (omitted arguments, string masks) is covered too.
 *
 * Expected values are Excel's, read back from Microsoft Excel for Mac (16.x).
 */

import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { fnWEEKDAY } from "@formula/functions/date";
import { DEFAULT_FUNCTION_CONTEXT } from "@formula/runtime/function-context";
import { rvNumber } from "@formula/runtime/values";
import { describe, it, expect } from "vitest";

function runFormula(formula: string): unknown {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Sheet1");
  Cell.setValue(ws, "A1", { formula, result: 0 });
  calculateFormulas(wb);
  return Cell.getResult(ws, "A1");
}

describe("rounding reads the value at Excel's 15-digit precision", () => {
  it.each([
    ["ROUND(1.005,2)", 1.01],
    ["ROUND(0.285,2)", 0.29],
    ["ROUND(-1.005,2)", -1.01],
    ["ROUND(2.675,2)", 2.68],
    ["ROUND(1234.5,-2)", 1200],
    ["ROUND(-2.5,0)", -3],
    ["ROUNDUP(0.1+0.2,1)", 0.3],
    ["ROUNDDOWN(1.005*1000,0)", 1005],
    ["TRUNC(4.35*100)", 435],
    ["MROUND(1.3,0.2)", 1.4],
    ["MROUND(10,3)", 9],
    ["ROUND(1E+300,-299)", 1e300],
    ["ROUND(1.5,400)", 1.5]
  ])("%s = %s", (formula, expected) => {
    expect(runFormula(formula)).toBe(expected);
  });

  it("FIXED and DOLLAR round the same way", () => {
    expect(runFormula("FIXED(1.005,2)")).toBe("1.01");
    expect(runFormula("DOLLAR(-1.005)")).toBe("($1.01)");
    expect(runFormula("FIXED(1234.5,-2)")).toBe("1,200");
  });
});

describe("DAYS360 (US method) treats the last day of February as the 30th", () => {
  it.each([
    ["DAYS360(DATE(2024,2,29),DATE(2024,3,31))", 30],
    ["DAYS360(DATE(2023,2,28),DATE(2023,3,31))", 30],
    ["DAYS360(DATE(2024,2,28),DATE(2024,3,31))", 33],
    ["DAYS360(DATE(2023,2,28),DATE(2023,3,15))", 15],
    // Excel's 1900-02-29 (serial 60) is February's last day; 1900-02-28 is not.
    ["DAYS360(60,DATE(1900,3,31))", 30],
    ["DAYS360(59,DATE(1900,3,31))", 33],
    // The European method leaves February alone.
    ["DAYS360(DATE(2023,2,28),DATE(2023,3,31),TRUE)", 32]
  ])("%s = %s", (formula, expected) => {
    expect(runFormula(formula)).toBe(expected);
  });
});

describe("NETWORKDAYS.INTL / WORKDAY.INTL weekend argument", () => {
  const week = "DATE(2024,1,1),DATE(2024,1,7)"; // Monday through Sunday
  it.each([
    [`NETWORKDAYS.INTL(${week},"0000011")`, 5],
    [`NETWORKDAYS.INTL(${week},"1000000")`, 6],
    [`NETWORKDAYS.INTL(${week},"0000000")`, 7],
    [`NETWORKDAYS.INTL(${week},"1111111")`, 0],
    [`NETWORKDAYS.INTL(${week},11)`, 6],
    [`NETWORKDAYS.INTL(${week},)`, 5],
    [`NETWORKDAYS.INTL(${week},"111")`, { error: "#VALUE!" }],
    [`NETWORKDAYS.INTL(${week},"0000021")`, { error: "#VALUE!" }],
    [`NETWORKDAYS.INTL(${week},"")`, { error: "#VALUE!" }],
    [`NETWORKDAYS.INTL(${week},8)`, { error: "#NUM!" }],
    [`NETWORKDAYS.INTL(${week},99)`, { error: "#NUM!" }],
    ['WORKDAY.INTL(DATE(2024,1,5),1,"0000011")', 45299],
    ['WORKDAY.INTL(DATE(2024,1,5),1,"0000001")', 45297],
    ['WORKDAY.INTL(DATE(2024,1,5),1,"1111111")', { error: "#VALUE!" }],
    ["WORKDAY.INTL(DATE(2024,1,5),1,0)", { error: "#NUM!" }]
  ])("%s = %j", (formula, expected) => {
    expect(runFormula(formula)).toEqual(expected);
  });
});

describe("an omitted rows / columns argument keeps the whole dimension", () => {
  it.each([
    ['TEXTJOIN(",",,TAKE({1,2,3},,-2))', "2,3"],
    ['TEXTJOIN(",",,TAKE({1,2;3,4},-1,))', "3,4"],
    ['TEXTJOIN(",",,EXPAND({1,2},,3,0))', "1,2,0"],
    ['TEXTJOIN(",",,EXPAND({1;2},3,,0))', "1,2,0"]
  ])("%s = %s", (formula, expected) => {
    expect(runFormula(formula)).toBe(expected);
  });
});

describe("weekdays follow Excel's 1900 calendar, which counts 1900-02-29", () => {
  it.each([
    ["WEEKDAY(0)", 7],
    ["WEEKDAY(1)", 1],
    ["WEEKDAY(59)", 3],
    ["WEEKDAY(60)", 4],
    ["WEEKDAY(61)", 5],
    ["WEEKDAY(1,2)", 7],
    ["WEEKDAY(-1)", { error: "#NUM!" }],
    ["WEEKNUM(0)", 0],
    ["WEEKNUM(1)", 1],
    ["WEEKNUM(7)", 1],
    ["WEEKNUM(8)", 2],
    ["WEEKNUM(DATE(2024,12,31))", 53],
    ["ISOWEEKNUM(1)", 52],
    ["ISOWEEKNUM(2)", 1],
    ["ISOWEEKNUM(DATE(2024,12,30))", 1],
    ["ISOWEEKNUM(DATE(2021,1,3))", 53],
    // Working days in the first week of 1900: Monday 1900-01-02 through Saturday the 6th.
    ["NETWORKDAYS(1,7)", 5],
    ["WORKDAY(1,1)", 2]
  ])("%s = %j", (formula, expected) => {
    expect(runFormula(formula)).toEqual(expected);
  });

  it("agrees with TEXT for every day around the phantom leap day", () => {
    const names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    for (let serial = 1; serial <= 70; serial++) {
      const weekday = runFormula(`WEEKDAY(${serial})`) as number;
      expect(runFormula(`TEXT(${serial},"dddd")`)).toBe(names[weekday - 1]);
    }
  });

  it("uses the 1904 epoch when the workbook does", () => {
    // Serial 0 is 1904-01-01, a Friday.
    const context = { ...DEFAULT_FUNCTION_CONTEXT, date1904: true };
    expect(fnWEEKDAY([rvNumber(0)], context)).toEqual(rvNumber(6));
  });
});

describe("FIXED and DOLLAR write every requested place exactly", () => {
  it.each([
    ["FIXED(1.01,20)", "1.01000000000000000000"],
    ["FIXED(0.1,3)", "0.100"],
    ["FIXED(1E+21,0)", "1,000,000,000,000,000,000,000"],
    ["FIXED(1E+21,0,TRUE)", "1000000000000000000000"],
    ["FIXED(1.5E-7,9)", "0.000000150"],
    ["FIXED(-1234.567,1)", "-1,234.6"],
    ["FIXED(-0.001,2)", "0.00"],
    ["FIXED(5,1.9)", "5.0"],
    ["DOLLAR(-0.001)", "($0.00)"],
    ["DOLLAR(1234.567,-2)", "$1,200"]
  ])("%s = %s", (formula, expected) => {
    expect(runFormula(formula)).toBe(expected);
  });

  it("accepts up to 127 places and refuses more", () => {
    expect(runFormula("LEN(FIXED(1,127))")).toBe(129);
    expect(runFormula("FIXED(1,128)")).toEqual({ error: "#VALUE!" });
    expect(runFormula("DOLLAR(1,128)")).toEqual({ error: "#VALUE!" });
  });
});

// Excel counts its 1900-02-29 in DATE, DATEVALUE, DAYS360, DATEDIF and the weekday functions, but
// not in EOMONTH, EDATE or YEARFRAC. Every expectation here was read back from Excel.
describe("date arithmetic follows Excel's calendar, 1900-02-29 and 1900-01-00 included", () => {
  it.each([
    ["DATE(1900,2,29)", 60],
    ["DATE(1900,3,0)", 60],
    ["DATE(1900,1,60)", 60],
    ["DATE(1900,1,0)", 0],
    ["DATE(1900,1,-1)", { error: "#NUM!" }],
    ["DATE(9999,12,31)", 2958465],
    ["DATE(9999,12,32)", { error: "#NUM!" }],
    ["DATE(2024,1.9,1.9)", 45292],
    ["YEAR(0)", 1900],
    ["MONTH(0)", 1],
    ["DAY(0)", 0],
    ["EOMONTH(1,1)", 59],
    ["EOMONTH(0,0)", 31],
    ["EOMONTH(DATE(1900,1,15),0)", 31],
    ["EOMONTH(-1,0)", { error: "#NUM!" }],
    ["EDATE(DATE(1900,1,31),1)", 59],
    ["EDATE(60,1)", 89],
    ["EDATE(DATE(2024,1,31),1)", 45351],
    ['DATEDIF(DATE(1900,2,15),DATE(1900,3,1),"MD")', 15],
    ['DATEDIF(1,61,"D")', 60],
    ['DATEDIF(1,61,"YD")', 60],
    ['DATEVALUE("2/29/1900")', 60],
    ['DATEVALUE("1900-02-29")', 60],
    ['DATEVALUE("February 29, 1900")', 60],
    ['DATEVALUE("2/30/1900")', { error: "#VALUE!" }],
    ['DATEVALUE("2/29/1901")', { error: "#VALUE!" }],
    ["YEARFRAC(DATE(1900,1,1),DATE(1901,1,1),1)", 366 / 365],
    ["YEARFRAC(DATE(1900,1,1),DATE(1900,12,31),1)", 1],
    ["YEARFRAC(DATE(1900,1,15),DATE(1900,3,31),1)", 76 / 365],
    ["YEARFRAC(60,DATE(1900,12,31),0)", 302 / 360],
    ["YEARFRAC(60,DATE(1900,12,31),1)", 306 / 365]
  ])("%s = %j", (formula, expected) => {
    const result = runFormula(formula);
    if (typeof expected === "number") {
      expect(result as number).toBeCloseTo(expected, 12);
    } else {
      expect(result).toEqual(expected);
    }
  });
});

describe("YEARFRAC basis 1 uses Excel's actual/actual rule", () => {
  it.each([
    ["YEARFRAC(DATE(2020,1,1),DATE(2021,1,1),1)", 1],
    ["YEARFRAC(DATE(2023,1,1),DATE(2024,1,1),1)", 1],
    ["YEARFRAC(DATE(2023,6,1),DATE(2024,6,1),1)", 366 / 366],
    ["YEARFRAC(DATE(2024,3,1),DATE(2025,3,1),1)", 365 / 365],
    ["YEARFRAC(DATE(2023,3,1),DATE(2024,2,29),1)", 365 / 366],
    // More than a year apart: the average length of 2020…2023.
    ["YEARFRAC(DATE(2020,6,1),DATE(2023,6,1),1)", 1095 / ((366 + 365 + 365 + 365) / 4)]
  ])("%s = %s", (formula, expected) => {
    expect(runFormula(formula) as number).toBeCloseTo(expected as number, 12);
  });
});
