/**
 * Every distribution function and its compatibility name, against values real Excel computed.
 *
 * `data/statistical-distributions.excel.json` is `[formula, value]` for a matrix of inputs chosen to find where Excel's
 * rules differ from the textbook: fractional degrees of freedom and counts (Excel truncates them — `T.DIST(0,1.5,FALSE)`
 * is 1/π), the legacy names with their own argument orders and tails (`TDIST`, `TINV`, `CHIINV`, `FINV`, `BETADIST`,
 * `CRITBINOM`, …), probabilities of 0, 1 and beyond, densities at the ends of their support, and tails down to 1e-300.
 * The workbook was written by this library with the formulas alone, opened in Excel 16 (macOS), which computed them, and
 * the saved values were read back. A value is a number or the error Excel showed.
 *
 * RANK, RANK.EQ and RANK.AVG are given references (C1:C3, D1:D6), not array constants: their `ref` argument must be a
 * reference, and Excel declines to open a workbook containing `RANK(2,{2,2,1},0)` at all — no workbook appears.
 *
 * `data/statistical-distributions.exact.json` lists the cases where Excel itself is wrong: its inverses are iterative
 * and, deep in a tail, stop short (`CHISQ.INV.RT(1E-20,1)` is `#NUM!`, though χ² with 1 degree of freedom exceeds
 * 87.16 with probability 1e-20), wander (`T.INV(1E-300,1000000)` is −77.97 where the answer is −37.06), or return a
 * value for an answer that is past the largest double (`F.INV.RT(1E-300,1,1)` is about 4e599). Those cases are held to
 * the exact answer instead, computed to 40+ digits with mpmath (bisection over `betainc` / `gammainc`, regularized), and
 * Excel's value is kept beside it. An exact answer below the smallest normal double is 0, as Excel stores it.
 */
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

import exactCases from "./data/statistical-distributions.exact.json";
import excel from "./data/statistical-distributions.excel.json";

const cases = excel as [string, number | string][];
const EXCEL_INACCURATE = exactCases as Record<
  string,
  { readonly exact: number | string; readonly excel: number | string }
>;

function computed(): (number | string | undefined)[] {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "M");
  // The data the RANK cases refer to, where the Excel workbook held it.
  [2, 2, 1].forEach((value, index) => Cell.setValue(sheet, `C${index + 1}`, value));
  [1, 2, 3, 4, 9, 2.5].forEach((value, index) => Cell.setValue(sheet, `D${index + 1}`, value));
  cases.forEach(([formula], index) => Cell.setValue(sheet, `A${index + 1}`, { formula }));
  calculateFormulas(workbook);
  return cases.map((_, index) => {
    const result = Cell.getResult(sheet, `A${index + 1}`) as number | { error: string } | undefined;
    return typeof result === "object" ? result.error : result;
  });
}

describe("statistical distributions against Excel", () => {
  const results = computed();

  it("lists only cases in the matrix as Excel-inaccurate", () => {
    const formulas = new Set(cases.map(([formula]) => formula));
    expect(Object.keys(EXCEL_INACCURATE).filter(formula => !formulas.has(formula))).toEqual([]);
  });

  it.each(cases.map((entry, index) => [...entry, index] as const))(
    "%s = %s",
    (formula, excelValue, index) => {
      const actual = results[index];
      const known = EXCEL_INACCURATE[formula];
      if (known !== undefined) {
        expect(excelValue).toBe(known.excel);
      }
      const expected = known?.exact ?? excelValue;
      if (typeof expected === "string") {
        expect(actual).toBe(expected);
        return;
      }
      expect(typeof actual).toBe("number");
      const tolerance = expected === 0 ? 0 : 1e-9 * Math.abs(expected);
      expect(Math.abs((actual as number) - expected)).toBeLessThanOrEqual(tolerance);
    }
  );
});
