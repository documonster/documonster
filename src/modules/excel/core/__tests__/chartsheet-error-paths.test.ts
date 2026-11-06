/**
 * An argument error from a chartsheet's chart names the path the caller wrote: `options.chart.series[0].values` for
 * `Workbook.addChartsheet(wb, name, options)` and `addPivotChartsheet`, `chart.series[0].values` for
 * `replaceChartsheetChart(wb, sheet, chart)` — not the validator's own `options.series[0].values` — and the message
 * starts with the call. Nothing is added or replaced.
 */
import { ExcelChartOptionsError, Chartsheet, ExcelArgumentError, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

const GOOD = { type: "bar", series: [{ values: "A!$A$1:$A$2" }] } as never;
const BAD = { type: "bar", series: [{ values: 42 }] } as never;

function book() {
  const wb = Workbook.create();
  Workbook.addWorksheet(wb, "A");
  return wb;
}

function caught(run: () => unknown): ExcelArgumentError {
  try {
    run();
  } catch (error) {
    return error as ExcelArgumentError;
  }
  throw new Error("did not throw");
}

describe("chartsheet chart errors are rooted at the caller's parameter", () => {
  it("Workbook.addChartsheet", () => {
    const wb = book();
    const error = caught(() => Workbook.addChartsheet(wb, "CS", { chart: BAD }));
    expect(error).toBeInstanceOf(ExcelChartOptionsError);
    expect(error.field).toBe("options.chart.series[0].values");
    expect(error.message).toMatch(/^Workbook\.addChartsheet: /);
    expect(error.message).toContain("options.chart.series[0].values");
    expect(Workbook.listChartsheets(wb)).toEqual([]);
  });

  it("Workbook.replaceChartsheetChart", () => {
    const wb = book();
    const cs = Workbook.addChartsheet(wb, "CS", { chart: GOOD });
    const before = Chartsheet.getChartModel(cs);
    const error = caught(() => Workbook.replaceChartsheetChart(wb, cs, BAD));
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect(error.field).toBe("chart.series[0].values");
    expect(error.message).toMatch(/^Workbook\.replaceChartsheetChart: /);
    expect(Chartsheet.getChartModel(Workbook.getChartsheet(wb, "CS")!)).toEqual(before);
  });
});
