/**
 * An edit through a sub-object a sheet hands out — its data-validation collection, a placed image, an image's
 * anchor, a table column — is refused once the sheet is removed, exactly as the same edit through the sheet is: the
 * edit would change something no file holds. A chartsheet's handle is refused once the chartsheet is removed. Reads
 * still answer.
 */
import {
  Anchor,
  Stream,
  Worksheet,
  Chartsheet,
  DataValidation,
  ExcelArgumentError,
  Image,
  Row,
  Table,
  Workbook
} from "@excel/index";
import type { DataValidationRule } from "@excel/types";
import { describe, expect, it } from "vitest";

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
  ),
  c => c.charCodeAt(0)
);

const RULE: DataValidationRule = { type: "whole", operator: "between", formulae: [1, 2] };

function sheet() {
  const wb = Workbook.create();
  Workbook.addWorksheet(wb, "A");
  const ws = Workbook.addWorksheet(wb, "B");
  Row.addMany(ws, [
    [1, 2],
    [3, 4]
  ]);
  DataValidation.add(ws, "C1", RULE);
  const imageId = Image.add(wb, { buffer: PNG, extension: "png" });
  Image.place(ws, imageId, "D1:E3");
  const table = Table.add(ws, {
    name: "Kpi",
    ref: "H1",
    headerRow: true,
    columns: [{ name: "a" }],
    rows: [[1]]
  });
  const collection = ws.dataValidations;
  const image = Image.list(ws)[0]!;
  const anchor = image.range!.tl;
  const column = Table.getColumn(table, 0)!;
  // The deprecated get-or-create handles: plain records the `Stream` setters also take.
  const row = Worksheet.getRow(ws, 1);
  const cell = Stream.getRowCell(row, 1);
  return { wb, ws, imageId, collection, image, anchor, column, row, cell };
}

type Sheet = ReturnType<typeof sheet>;

function removed(): Sheet {
  const s = sheet();
  Workbook.removeWorksheet(s.wb, s.ws);
  return s;
}

/** The sheet's place taken by a model set on its workbook — its handles are detached the same way. */
function replaced(): Sheet {
  const s = sheet();
  Workbook.setModel(s.wb, Workbook.getModel(Workbook.create()));
  return s;
}

const EDITS: [string, (s: Sheet) => unknown][] = [
  ["DataValidation.add on the collection", s => DataValidation.add(s.collection, "A1", RULE)],
  ["DataValidation.remove on the collection", s => DataValidation.remove(s.collection, "C1")],
  [
    "Image.setModel",
    s => Image.setModel(s.image, { type: "image", imageId: s.imageId, range: "A1:B2" } as never)
  ],
  [
    "Image.applyModel",
    s => Image.applyModel(s.image, { type: "image", imageId: s.imageId, range: "A1:B2" } as never)
  ],
  ["Anchor.setCol", s => Anchor.setCol(s.anchor, 2)],
  ["Anchor.setRow", s => Anchor.setRow(s.anchor, 2)],
  [
    "Anchor.setModel",
    s => Anchor.setModel(s.anchor, { nativeCol: 1, nativeColOff: 0, nativeRow: 1, nativeRowOff: 0 })
  ],
  ["Table.setColumnName", s => Table.setColumnName(s.column, "b")],
  ["Table.setColumnStyle", s => Table.setColumnStyle(s.column, { font: { bold: true } })],
  ["Table.setColumnTotalsRowLabel", s => Table.setColumnTotalsRowLabel(s.column, "x")],
  ["Stream.setCellValue on a sheet's cell", s => Stream.setCellValue(s.cell, 9)],
  ["Stream.setRowHeight on a sheet's row", s => Stream.setRowHeight(s.row, 20)]
];

describe.each([
  ["removed", removed],
  ["replaced by Workbook.setModel", replaced]
])("a sub-object of a sheet %s", (_how, make) => {
  it.each(EDITS)("%s is refused", (_name, edit) => {
    const s = make();
    expect(() => edit(s)).toThrow(ExcelArgumentError);
    expect(() => edit(make())).toThrow(/was removed from its workbook/);
  });

  it("refuses before changing anything, and reads still answer", () => {
    const s = make();
    expect(() => DataValidation.add(s.collection, "A1", RULE)).toThrow(ExcelArgumentError);
    expect(DataValidation.find(s.collection, "A1")).toBeUndefined();
    expect(DataValidation.find(s.collection, "C1")).toMatchObject({ type: "whole" });
    expect(() => Anchor.setCol(s.anchor, 5)).toThrow(ExcelArgumentError);
    expect(Anchor.getModel(s.anchor).nativeCol).toBe(3);
    expect(Image.getModel(s.image)).toMatchObject({ type: "image" });
    expect(Table.getColumnName(s.column)).toBe("a");
  });
});

describe("edits through the sub-objects of an attached sheet", () => {
  it.each(EDITS)("%s still works", (_name, edit) => {
    expect(() => edit(sheet())).not.toThrow();
  });

  it("a collection that belongs to no sheet is edited freely", () => {
    const collection = DataValidation.create();
    DataValidation.add(collection, "A1", RULE);
    expect(DataValidation.find(collection, "A1")).toMatchObject({ type: "whole" });
  });
});

describe("a removed chartsheet", () => {
  function chartsheet() {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "D");
    Row.addMany(ws, [[1], [2]]);
    const cs = Workbook.addChartsheet(wb, "C", {
      chart: { type: "line", series: [{ values: "D!$A$1:$A$2" }] }
    });
    return { wb, cs };
  }

  it.each<[string, (cs: ReturnType<typeof chartsheet>["cs"]) => unknown]>([
    ["Chartsheet.setName", cs => Chartsheet.setName(cs, "Z")],
    ["Chartsheet.setState", cs => Chartsheet.setState(cs, "hidden")],
    ["Chartsheet.setZoomScale", cs => Chartsheet.setZoomScale(cs, 120)],
    ["Chartsheet.setTabSelected", cs => Chartsheet.setTabSelected(cs, true)],
    [
      "Chartsheet.setPageMargins",
      cs =>
        Chartsheet.setPageMargins(cs, {
          left: 1,
          right: 1,
          top: 1,
          bottom: 1,
          header: 0.5,
          footer: 0.5
        })
    ]
  ])("%s is refused; reads still answer", (_name, edit) => {
    const { wb, cs } = chartsheet();
    expect(() => edit(cs)).not.toThrow();
    const again = chartsheet();
    Workbook.removeChartsheet(again.wb, again.cs);
    expect(() => edit(again.cs)).toThrow(ExcelArgumentError);
    expect(() => edit(again.cs)).toThrow(/chartsheet "C" was removed from its workbook/);
    expect(Chartsheet.getModel(again.cs)).toMatchObject({ name: "C" });
    void wb;
  });
});
