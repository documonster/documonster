/**
 * The statistical inverses `CONFIDENCE.T` builds on, against values real Excel computed.
 *
 * `data/statistical-inverses.excel.json` is `[formula, value]` for every formula of a matrix of inputs: the workbook was
 * written by this library with the formulas alone, opened in Excel 16.114 (macOS), which computed them, and the saved
 * values were read back. A value is a number or the error Excel showed.
 *
 * Reported from `scientific-analysis`: `CONFIDENCE.T(0.05,8.27,350)` was `#NUM!`, because `lnGamma` overflowed past
 * 171 and every distribution with large degrees of freedom returned NaN.
 */
import { calculateFormulas } from "@excel/core/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

import excel from "./data/statistical-inverses.excel.json";

const cases = excel as [string, number | string][];

/**
 * Where Excel itself is further than 1e-9 from the exact answer, the case is held to the exact answer instead, computed
 * to 40 digits with mpmath (`findroot` over `gammainc(…, regularized=True)`). Excel's own value is kept beside it.
 */
const EXCEL_INACCURATE: Record<string, { readonly exact: number; readonly excel: number }> = {
  // Excel is 3.9e-9 relative low here: Q(1/2, x/2) = 1e-10 at 41.821456364761294…
  "CHISQ.INV.RT(1e-10,1)": { exact: 41.82145636476129, excel: 41.82145620298276 }
};

function computed(): (number | string | undefined)[] {
  const workbook = Workbook.create();
  const sheet = Workbook.addWorksheet(workbook, "M");
  cases.forEach(([formula], index) => Cell.setValue(sheet, `A${index + 1}`, { formula }));
  calculateFormulas(workbook);
  return cases.map((_, index) => {
    const result = Cell.getResult(sheet, `A${index + 1}`) as number | { error: string } | undefined;
    return typeof result === "object" ? result.error : result;
  });
}

describe("statistical inverses against Excel", () => {
  const results = computed();

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
