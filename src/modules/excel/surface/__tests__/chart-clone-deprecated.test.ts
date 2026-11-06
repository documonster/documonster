import { Cell, Chart, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function sheetWithChart() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", "x");
  Cell.setValue(ws, "B1", 1);
  Chart.add(ws, { type: "bar", series: [{ name: "S", values: "S!$B$1:$B$1" }] }, "C1:J10");
  return ws;
}

describe("Chart.clone (deprecated forward)", () => {
  it("places the copy beside the chart, as Chart.copyTo onto its own sheet does", () => {
    const viaClone = sheetWithChart();
    const viaCopy = sheetWithChart();
    const a = Chart.clone(Chart.list(viaClone)[0]);
    const b = Chart.copyTo(Chart.list(viaCopy)[0], viaCopy);
    expect(a).toBe(b.chartNumber);
    expect(b).toBe(Chart.list(viaCopy)[1]);
    expect(Chart.getAnchorModel(Chart.list(viaClone)[1])).toEqual(
      Chart.getAnchorModel(Chart.list(viaCopy)[1])
    );
  });

  it("keeps its call name in argument errors", () => {
    const ws = sheetWithChart();
    expect(() => Chart.clone(Chart.list(ws)[0], 42 as never)).toThrow(/^Chart\.clone:/);
  });
});
