/**
 * Every remove that takes a position refuses a fractional one — it names no item — and answers `false` for an
 * out-of-range one. `splice` used to truncate `1.5` to `1` and remove an item nobody asked for; answering `false`
 * hid the mistake as an absence.
 */
import { Cell, Chart, ExcelArgumentError, Form, Image, Sparkline, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

const PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function sheet() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Cell.setValue(ws, "A1", 1);
  Cell.setValue(ws, "A2", 2);
  return { wb, ws };
}

describe("position-based removes reject a non-integer position", () => {
  it("Chart.remove", () => {
    const { ws } = sheet();
    Chart.place(ws, "C1:J10", { type: "bar", series: [{ values: "S!$A$1:$A$2" }] });
    Chart.place(ws, "C12:J20", { type: "line", series: [{ values: "S!$A$1:$A$2" }] });
    expect(() => Chart.remove(ws, 1.5)).toThrow(ExcelArgumentError);
    expect(() => Chart.remove(ws, Number.NaN)).toThrow(ExcelArgumentError);
    expect(Chart.remove(ws, 5)).toBe(false);
    expect(Chart.list(ws)).toHaveLength(2);
    expect(Chart.remove(ws, 1)).toBe(true);
  });

  it("Sparkline.remove", () => {
    const { ws } = sheet();
    Sparkline.add(ws, { type: "line", sparklines: [{ cellRef: "B1", dataRef: "S!A1:A2" }] });
    Sparkline.add(ws, { type: "line", sparklines: [{ cellRef: "B2", dataRef: "S!A1:A2" }] });
    expect(() => Sparkline.remove(ws, 0.5)).toThrow(ExcelArgumentError);
    expect(Sparkline.list(ws)).toHaveLength(2);
  });

  it("Image.remove and Form.remove", () => {
    const { wb, ws } = sheet();
    const id = Image.add(wb, { base64: PNG, extension: "png" });
    Image.place(ws, id, "A1:B2");
    Image.place(ws, id, "D1:E2");
    Form.addCheckbox(ws, "C2:D3");
    Form.addCheckbox(ws, "C5:D6");
    expect(() => Image.remove(ws, 1.5)).toThrow(ExcelArgumentError);
    expect(() => Form.remove(ws, 1.5)).toThrow(ExcelArgumentError);
    expect(Image.list(ws)).toHaveLength(2);
    expect(Form.list(ws)).toHaveLength(2);
  });

  it("Workbook.removeChartsheet", () => {
    const { wb } = sheet();
    for (const name of ["C1", "C2"]) {
      Workbook.addChartsheet(wb, name, {
        chart: { type: "bar", series: [{ values: "S!$A$1:$A$2" }] }
      });
    }
    expect(() => Workbook.removeChartsheet(wb, 1.5)).toThrow(ExcelArgumentError);
    expect(Workbook.removeChartsheet(wb, 5)).toBe(false);
    expect(Workbook.listChartsheets(wb)).toHaveLength(2);
  });
});

describe("Chart.place", () => {
  it("does not treat `pivotTable: undefined` as a pivot chart", () => {
    const { ws } = sheet();
    const chart = Chart.place(ws, "C1:J10", {
      type: "bar",
      series: [{ values: "S!$A$1:$A$2" }],
      pivotTable: undefined
    } as never);
    expect(Chart.list(ws)).toEqual([chart]);
  });
});
