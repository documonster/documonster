/**
 * The chart PNG renderers check `dpi` and `scale` themselves, before choosing a backend, so a bad value is the same
 * `ExcelChartOptionsError` in Node and in a browser — not an archive error from the PNG encoder on one side and a silently
 * dropped resolution on the other.
 */
import { ExcelChartOptionsError } from "@excel/errors";
import { Cell, Chart, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function chartHandle(type: "bar" | "waterfall") {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  for (const [i, [label, value]] of [
    ["a", 1],
    ["b", 2],
    ["c", 3]
  ].entries()) {
    Cell.setValue(ws, `A${i + 1}`, label);
    Cell.setValue(ws, `B${i + 1}`, value);
  }
  // "waterfall" is an Office 2016 (ChartEx) chart, rendered by the other PNG path.
  return Chart.place(ws, "D2:J14", {
    type,
    series: [{ categories: "S!$A$1:$A$3", values: "S!$B$1:$B$3" }]
  } as never);
}

describe.each(["bar", "waterfall"] as const)("Chart.toPng (%s) raster options", type => {
  it.each([NaN, 0, -72, Infinity, "300"])(
    "refuses dpi %s with ExcelChartOptionsError",
    async dpi => {
      await expect(Chart.toPng(chartHandle(type), { dpi: dpi as number })).rejects.toBeInstanceOf(
        ExcelChartOptionsError
      );
    }
  );

  it.each([NaN, 0, -1])("refuses scale %s with ExcelChartOptionsError", async scale => {
    await expect(Chart.toPng(chartHandle(type), { scale })).rejects.toBeInstanceOf(
      ExcelChartOptionsError
    );
  });

  it("accepts a positive dpi and no dpi", async () => {
    expect((await Chart.toPng(chartHandle(type), { dpi: 144 })).byteLength).toBeGreaterThan(0);
    expect((await Chart.toPng(chartHandle(type))).byteLength).toBeGreaterThan(0);
  });
});
