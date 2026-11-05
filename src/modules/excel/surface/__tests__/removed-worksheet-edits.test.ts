/**
 * A worksheet removed from its workbook is no longer written, so an edit to it would vanish. Every edit refuses it —
 * sheet-level ones and those made through a handle the sheet holds — while reads still answer.
 */
import {
  Cell,
  Chart,
  Column,
  ConditionalFormatting,
  DataValidation,
  ExcelArgumentError,
  Form,
  Image,
  Row,
  Table,
  Watermark,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
  ),
  c => c.charCodeAt(0)
);

function removedSheet() {
  const wb = Workbook.create();
  Workbook.addWorksheet(wb, "A");
  const ws = Workbook.addWorksheet(wb, "B");
  Row.addMany(ws, [
    [1, 2],
    [3, 4]
  ]);
  const table = Table.add(ws, {
    name: "Kpi",
    ref: "H1",
    headerRow: true,
    columns: [{ name: "a" }],
    rows: [[1]]
  });
  const chart = Chart.place(ws, "D1:F8", { type: "line", series: [{ values: "B!$A$1:$A$2" }] });
  const checkbox = Form.addCheckbox(ws, "J1", {});
  const imageId = Image.add(wb, { buffer: PNG, extension: "png" });
  Workbook.removeWorksheet(wb, ws);
  return { wb, ws, table, chart, checkbox, imageId };
}

const REMOVED = /was removed from its workbook|was removed from its sheet/;

describe("a removed worksheet cannot be edited", () => {
  it.each<[string, (s: ReturnType<typeof removedSheet>) => unknown]>([
    ["Worksheet.setName", ({ ws }) => Worksheet.setName(ws, "Z")],
    ["Worksheet.freeze", ({ ws }) => Worksheet.freeze(ws, { rows: 1 })],
    ["Worksheet.unfreeze", ({ ws }) => Worksheet.unfreeze(ws)],
    ["Worksheet.setAutoFilter", ({ ws }) => Worksheet.setAutoFilter(ws, "A1:B2")],
    ["Worksheet.unprotect", ({ ws }) => Worksheet.unprotect(ws)],
    ["Column.setWidth", ({ ws }) => Column.setWidth(ws, 1, 10)],
    ["Column.autoFit", ({ ws }) => Column.autoFit(ws, 1)],
    ["Row.autoFit", ({ ws }) => Row.autoFit(ws, 1)],
    [
      "Chart.place",
      ({ ws }) => Chart.place(ws, "D1:K10", { type: "line", series: [{ values: "B!$A$1:$A$2" }] })
    ],
    ["Chart.remove", ({ ws, chart }) => Chart.remove(ws, chart)],
    ["Table.remove", ({ ws }) => Table.remove(ws, "Kpi")],
    ["Image.place", ({ ws, imageId }) => Image.place(ws, imageId, "A1:B2")],
    ["Image.setBackground", ({ ws, imageId }) => Image.setBackground(ws, imageId)],
    ["Watermark.set", ({ ws, imageId }) => Watermark.set(ws, { imageId })],
    ["Form.addCheckbox", ({ ws }) => Form.addCheckbox(ws, "A1", {})],
    [
      "DataValidation.add",
      ({ ws }) =>
        DataValidation.add(ws, "A1", { type: "whole", operator: "between", formulae: [1, 2] })
    ],
    ["DataValidation.remove", ({ ws }) => DataValidation.remove(ws, "A1")],
    [
      "ConditionalFormatting.add",
      ({ ws }) =>
        ConditionalFormatting.add(ws, {
          ref: "A1",
          rules: [{ type: "expression", formulae: ["TRUE"], priority: 1 }]
        })
    ],
    ["ConditionalFormatting.clear", ({ ws }) => ConditionalFormatting.clear(ws)],
    // Through a handle the sheet holds.
    ["Table.setTheme", ({ table }) => Table.setTheme(table, "TableStyleLight1")],
    ["Chart.setTitle", ({ chart }) => Chart.setTitle(chart, "x")],
    ["Chart.setStyle", ({ chart }) => Chart.setStyle(chart, 3)],
    [
      "Chart.addSeriesFromOptions",
      ({ chart }) => Chart.addSeriesFromOptions(chart, { values: "B!$A$1:$A$2" })
    ],
    ["Form.setChecked", ({ checkbox }) => Form.setChecked(checkbox, true)]
  ])("%s", (_name, edit) => {
    const sheet = removedSheet();
    expect(() => edit(sheet)).toThrow(ExcelArgumentError);
    expect(() => edit(removedSheet())).toThrow(REMOVED);
  });

  it("while reads still answer", () => {
    const { ws, table, checkbox } = removedSheet();
    expect(Cell.getValue(ws, "B2")).toBe(4);
    expect(Worksheet.getName(ws)).toBe("B");
    expect(Table.getName(table)).toBe("Kpi");
    expect(Form.isChecked(checkbox)).toBe(false);
    expect(DataValidation.find(ws, "A1")).toBeUndefined();
  });
});

describe("a chart handle whose chart is gone", () => {
  it("refuses an edit by name instead of ignoring it", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "B");
    Row.addMany(ws, [[1], [2]]);
    const chart = Chart.place(ws, "D1:F8", { type: "line", series: [{ values: "B!$A$1:$A$2" }] });
    Chart.remove(ws, chart);
    for (const edit of [
      () => Chart.setTitle(chart, "x"),
      () => Chart.setStyle(chart, 3),
      () => Chart.setSeriesName(chart, 0, "n"),
      () => Chart.mutate(chart, () => {})
    ]) {
      expect(edit).toThrow(/Chart\.\w+: the chart was removed from its sheet/);
    }
  });

  it("Chart.copyTo names its missing target", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "B");
    const chart = Chart.place(ws, "D1:F8", { type: "line", series: [{ values: "B!$A$1:$A$1" }] });
    expect(() => Chart.copyTo(chart, undefined as never)).toThrow(/targetWs/);
  });
});
