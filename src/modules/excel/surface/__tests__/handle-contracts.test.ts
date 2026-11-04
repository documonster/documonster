/**
 * Handles of removed objects refuse edits, the checkbox link is a reference, and chartsheet readers hand out copies.
 */
import {
  Chart,
  Chartsheet,
  ExcelArgumentError,
  Form,
  InvalidAddressError,
  Workbook,
  Worksheet
} from "@excel";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "Data");
  Worksheet.addAoa(ws, [
    ["k", "v"],
    ["a", 1],
    ["b", 2]
  ]);
  return { wb, ws };
}

const CHART = {
  type: "bar" as const,
  barDir: "col" as const,
  series: [{ values: "Data!$B$2:$B$3", categories: "Data!$A$2:$A$3" }]
};

describe("Form checkboxes", () => {
  it("a removed checkbox's handle refuses edits", () => {
    const { ws } = book();
    const cb = Form.addCheckbox(ws, "D2", { text: "One", link: "A1" });
    expect(Form.remove(ws, cb)).toBe(true);
    expect(() => Form.setChecked(cb, true)).toThrow(ExcelArgumentError);
    expect(() => Form.setText(cb, "x")).toThrow(ExcelArgumentError);
    expect(() => Form.setLink(cb, "C1")).toThrow(ExcelArgumentError);
    expect(Form.getText(cb)).toBe("One");
  });

  it("the link is a cell, a range or a defined name", () => {
    const { ws } = book();
    const cb = Form.addCheckbox(ws, "D2");
    expect(() => Form.setLink(cb, "not a ref !!")).toThrow(InvalidAddressError);
    expect(() => Form.addCheckbox(ws, "D5", { link: "x y" })).toThrow(InvalidAddressError);
    Form.setLink(cb, "=$C$3");
    expect(Form.getLink(cb)).toBe("$C$3");
    Form.setLink(cb, "'My Sheet'!B2");
    expect(Form.getLink(cb)).toBe("'My Sheet'!$B$2");
    Form.setLink(cb, "Choice");
    expect(Form.getLink(cb)).toBe("Choice");
    Form.setLink(cb, "Other!Choice");
    expect(Form.getLink(cb)).toBe("Other!Choice");
    expect(() => Form.setLink(cb, "Other!not a name")).toThrow(InvalidAddressError);
  });

  it("setChecked takes a boolean", () => {
    const { ws } = book();
    const cb = Form.addCheckbox(ws, "D2", { checked: true });
    expect(() => Form.setChecked(cb, "yes" as never)).toThrow(ExcelArgumentError);
    expect(() => Form.setChecked(cb, undefined as never)).toThrow(ExcelArgumentError);
    expect(Form.isChecked(cb)).toBe(true);
  });
});

describe("Chartsheets", () => {
  it("a removed chartsheet's handle refuses edits", () => {
    const { wb } = book();
    const cs = Workbook.addChartsheet(wb, "CS", { chart: CHART });
    Workbook.removeChartsheet(wb, "CS");
    expect(() => Chartsheet.setName(cs, "Again")).toThrow(ExcelArgumentError);
    expect(() => Chartsheet.setState(cs, "hidden")).toThrow(ExcelArgumentError);
    expect(() => Chartsheet.setZoomScale(cs, 50)).toThrow(ExcelArgumentError);
    expect(Chartsheet.getName(cs)).toBe("CS");
  });

  it("getModel and the setting readers return copies", () => {
    const { wb } = book();
    const cs = Workbook.addChartsheet(wb, "CS", { chart: CHART });
    Chartsheet.setPageMargins(cs, { left: 1, right: 1 });
    const model = Chartsheet.getModel(cs);
    model.name = "Hacked/Name:*";
    expect(Chartsheet.getName(cs)).toBe("CS");
    Chartsheet.getPageMargins(cs)!.left = 9;
    expect(Chartsheet.getPageMargins(cs)).toEqual({ left: 1, right: 1 });
    // The deprecated name keeps returning the live model.
    expect(Chartsheet.model(cs)).toBe(Chartsheet.model(cs));
  });

  it("setPageMargins refuses keys the margins do not have", () => {
    const { wb } = book();
    const cs = Workbook.addChartsheet(wb, "CS", { chart: CHART });
    expect(() => Chartsheet.setPageMargins(cs, { bogus: 1 } as never)).toThrow(ExcelArgumentError);
    expect(Chartsheet.getPageMargins(cs)).toBeUndefined();
  });
});

describe("Chart.copyTo and Chart.clone of a removed chart", () => {
  it("refuse it as every other edit does", () => {
    const { ws } = book();
    const chart = Chart.place(ws, "E1:L15", CHART);
    Chart.remove(ws, chart);
    expect(() => Chart.copyTo(chart, ws)).toThrow(ExcelArgumentError);
    expect(() => Chart.clone(chart)).toThrow(ExcelArgumentError);
  });
});

describe("Chart.applyPreset", () => {
  it("names both families when the preset is neither", () => {
    expect(() => Chart.applyPreset("nope" as never, { series: [] } as never)).toThrow(
      /classic preset .*Office 2016/
    );
  });
});

describe("the chart option builders check their arguments, whichever entry calls them", () => {
  it("createOptionsFromRows and createSeriesFromColumns refuse a missing options object", async () => {
    const { ChartClassic } = await import("@excel/chart");
    const { ws } = book();
    expect(() => ChartClassic.createOptionsFromRows(ws, [{ a: 1 }], null as never)).toThrow(
      ExcelArgumentError
    );
    expect(() => ChartClassic.createSeriesFromColumns("S", {} as never)).toThrow(
      ExcelArgumentError
    );
    expect(() => Chart.createSeriesFromColumns(ws, {} as never)).toThrow(ExcelArgumentError);
  });
});
