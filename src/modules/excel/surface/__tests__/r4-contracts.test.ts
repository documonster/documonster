/**
 * Round-4 review fixes in the document model: edits that refuse leave the sheet as it was, readers answer in the words
 * their setters take, collection `remove`s refuse what names nothing, and a refused cell value is one error class.
 */
import {
  Address,
  Anchor,
  Cell,
  Chart,
  Column,
  ConditionalFormatting,
  DataValidation,
  ExcelArgumentError,
  Form,
  HeaderFooterImage,
  Image,
  ExcelInvalidAddressError,
  ExcelInvalidValueTypeError,
  Note,
  Pivot,
  Range,
  Row,
  Sparkline,
  Table,
  Watermark,
  Workbook,
  Worksheet
} from "@excel";
import { describe, expect, it } from "vitest";

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
  ),
  c => c.charCodeAt(0)
);

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  return { wb, ws };
}

const hasTemporal = typeof (globalThis as { Temporal?: unknown }).Temporal !== "undefined";
const temporal = (globalThis as { Temporal?: any }).Temporal;

describe("R4-E1 a refused Column.splice changes nothing", () => {
  it("checks every inserted value before a cell moves or a reference is rewritten", () => {
    const { ws } = book();
    Row.addMany(ws, [
      ["a", "b"],
      ["c", "d"],
      ["e", "f"]
    ]);
    Cell.setValue(ws, "D1", { formula: "B1" });
    const before = Worksheet.toAoa(ws);
    const error = (() => {
      try {
        Column.splice(ws, 1, 0, [1, Number.NaN]);
      } catch (e) {
        return e;
      }
    })();
    expect(error).toBeInstanceOf(ExcelInvalidValueTypeError);
    expect((error as ExcelInvalidValueTypeError).field).toBe("inserts[0][1]");
    expect((error as Error).message).toMatch(/^Column\.splice: inserts\[0\]\[1\]: /);
    expect(Worksheet.toAoa(ws)).toEqual(before);
    expect(Cell.getFormula(ws, "D1")).toBe("B1");
  });
});

describe("R4-E2 Column.splice writes every inserted value", () => {
  it("below the sheet's last row, and on an empty sheet", () => {
    const { wb, ws } = book();
    Row.add(ws, ["x"]);
    Column.splice(ws, 2, 0, [10, 20, 30]);
    expect([1, 2, 3].map(r => Cell.getValue(ws, { row: r, col: 2 }))).toEqual([10, 20, 30]);

    const empty = Workbook.addWorksheet(wb, "Empty");
    Column.splice(empty, 1, 0, [1, 2]);
    expect([Cell.getValue(empty, "A1"), Cell.getValue(empty, "A2")]).toEqual([1, 2]);
  });
});

describe.runIf(hasTemporal)("R4-E3 Column.splice keeps the format a Temporal value sets", () => {
  it("as Cell.setValue gives it", () => {
    const { ws } = book();
    Row.addMany(ws, [["x"], ["y"]]);
    const time = temporal.PlainTime.from("14:30");
    const dateTime = temporal.PlainDateTime.from("2024-03-10T14:30");
    Column.splice(ws, 2, 0, [time, dateTime]);
    Cell.setValue(ws, "C1", time);
    Cell.setValue(ws, "C2", dateTime);
    expect(Cell.getNumFmt(ws, "B1")).toBe(Cell.getNumFmt(ws, "C1"));
    expect(Cell.getDateKind(ws, "B1")).toBe("time");
    expect(Cell.getNumFmt(ws, "B2")).toBe(Cell.getNumFmt(ws, "C2"));
    expect(Cell.getDateKind(ws, "B2")).toBe("dateTime");
  });
});

describe("R4-E4 a refused Column.setValues writes nothing", () => {
  it("checks every value first, as Row.setValues does", () => {
    const { ws } = book();
    expect(() => Column.setValues(ws, "D", [1, Number.NaN], { base: 0 })).toThrow(
      ExcelInvalidValueTypeError
    );
    expect(Cell.has(ws, "D1")).toBe(false);
    expect(Column.getValues(ws, "D")).toEqual([]);
  });
});

describe("R4-S4 a refused cell value is one error class, naming the call and the argument", () => {
  it("is an ExcelInvalidValueTypeError, which is an ExcelArgumentError, for scalars and shapes alike", () => {
    const { ws } = book();
    const caught = (run: () => void) => {
      try {
        run();
      } catch (e) {
        return e as ExcelInvalidValueTypeError;
      }
      throw new Error("did not throw");
    };
    const nan = caught(() => Cell.setValue(ws, "A1", Number.NaN));
    const shape = caught(() => Cell.setValue(ws, "A1", { richText: [{ text: 42 }] } as never));
    for (const error of [nan, shape]) {
      expect(error).toBeInstanceOf(ExcelInvalidValueTypeError);
      expect(error).toBeInstanceOf(ExcelArgumentError);
      expect(error.message).toMatch(/^Cell\.setValue: /);
    }
    expect(nan.field).toBe("value");
    expect(shape.field).toBe("value.richText[0].text");

    const row = caught(() => Row.add(ws, [1, Number.POSITIVE_INFINITY]));
    expect(row.field).toBe("values[1]");
    expect(row.message).toMatch(/^Row\.add: values\[1\]: /);
    const aoa = caught(() => Worksheet.addAoa(ws, [[1], [() => 1]] as never));
    expect(aoa.field).toBe("data[1][0]");
    expect(aoa.message).toMatch(/^Worksheet\.addAoa: /);
  });

  it("keeps a bare value check's message when no call is named", () => {
    expect(new ExcelInvalidValueTypeError("number").field).toBe("value");
  });
});

describe("R4-C1 remove(…) refuses a target that names nothing", () => {
  function populated() {
    const { wb, ws } = book();
    Row.addMany(ws, [
      ["Region", "Amt"],
      ["N", 1],
      ["S", 2]
    ]);
    const pivots = Workbook.addWorksheet(wb, "P");
    const imageId = Image.add(wb, { buffer: PNG, extension: "png" });
    Image.place(ws, imageId, "D1:E3");
    Image.addShape(ws, { type: "rect", range: "D5:E6" });
    Pivot.add(pivots, { sourceSheet: ws, rows: ["Region"], values: ["Amt"], metric: "sum" });
    Sparkline.add(ws, { type: "line", sparklines: [{ dataRef: "S!B2:B3", cellRef: "F2" }] });
    Form.addCheckbox(ws, "H1:I2", { text: "x" });
    ConditionalFormatting.add(ws, {
      ref: "A1:B3",
      rules: [{ type: "expression", formulae: ["TRUE"], style: {} }]
    });
    Chart.place(ws, "K1:P10", { type: "bar", series: [{ values: "S!$B$2:$B$3" }] });
    return { wb, ws, pivots };
  }

  it.each(["x", 1.5, null, true])(
    "Image, shapes, Pivot, Sparkline, Form, Chart, CF refuse %j",
    target => {
      const { ws, pivots } = populated();
      const removes: [string, () => boolean][] = [
        ["image", () => Image.remove(ws, target as never)],
        ["shape", () => Image.removeShape(ws, target as never)],
        ["pivot", () => Pivot.remove(pivots, target as never)],
        ["groupOrIndex", () => Sparkline.remove(ws, target as never)],
        ["control", () => Form.remove(ws, target as never)],
        ["chart", () => Chart.remove(ws, target as never)]
      ];
      for (const [field, remove] of removes) {
        expect(remove).toThrow(expect.objectContaining({ field }));
        expect(remove).toThrow(ExcelArgumentError);
      }
      if (typeof target !== "string") {
        expect(() => ConditionalFormatting.remove(ws, target as never)).toThrow(ExcelArgumentError);
      }
      // Nothing was removed.
      expect(Image.list(ws)).toHaveLength(1);
      expect(Chart.list(ws)).toHaveLength(1);
    }
  );

  it("still answers false for a position or a handle that is not in the list", () => {
    const { wb, ws } = populated();
    const other = Workbook.addWorksheet(wb, "Other");
    expect(Image.remove(ws, 9)).toBe(false);
    expect(Image.remove(other, Image.list(ws)[0]!)).toBe(false);
    expect(Chart.remove(ws, -1)).toBe(false);
  });

  it("HeaderFooterImage.remove refuses a section that is not one", () => {
    const { ws } = book();
    expect(() => HeaderFooterImage.remove(ws, "bogus" as never)).toThrow(
      expect.objectContaining({ field: "position" })
    );
    expect(HeaderFooterImage.remove(ws, "CH")).toBe(false);
  });

  it("Range.containsCell refuses a string, which it compared as a decoded address", () => {
    const region = Range.create("A3:B4");
    expect(() => Range.containsCell(region, "A3" as never)).toThrow(ExcelArgumentError);
    expect(Range.containsCell(region, { address: "A3", row: 3, col: 1 })).toBe(true);
  });
});

describe("R4-C1 ConditionalFormatting.remove by ref, and a predicate removes one", () => {
  it("takes the block's ref, compared as Excel reads a reference", () => {
    const { ws } = book();
    const rules = [{ type: "expression" as const, formulae: ["TRUE"], style: {} }];
    ConditionalFormatting.add(ws, { ref: "A1:B3", rules });
    ConditionalFormatting.add(ws, { ref: "C1:C3", rules });
    expect(ConditionalFormatting.remove(ws, "a1:b3")).toBe(true);
    expect(ConditionalFormatting.list(ws).map(cf => cf.ref)).toEqual(["C1:C3"]);
    expect(ConditionalFormatting.remove(ws, "A1:B3")).toBe(false);
    expect(() => ConditionalFormatting.remove(ws, "not a ref")).toThrow(ExcelInvalidAddressError);
  });

  it("removes the first entry a predicate accepts, as the TSDoc says", () => {
    const { ws } = book();
    const rules = [{ type: "expression" as const, formulae: ["TRUE"], style: {} }];
    ConditionalFormatting.add(ws, { ref: "A1", rules });
    ConditionalFormatting.add(ws, { ref: "A2", rules });
    expect(ConditionalFormatting.remove(ws, () => true)).toBe(true);
    expect(ConditionalFormatting.list(ws).map(cf => cf.ref)).toEqual(["A2"]);
  });
});

describe("R4-C2 ConditionalFormatting.add stores its ref canonical", () => {
  it("upper-cases a lowercase ref, as every other reference is", () => {
    const { ws } = book();
    const added = ConditionalFormatting.add(ws, {
      ref: "a1:b2 d4",
      rules: [{ type: "expression", formulae: ["TRUE"], style: {} }]
    });
    expect(added.ref).toBe("A1:B2 D4");
    expect(ConditionalFormatting.list(ws)[0]!.ref).toBe("A1:B2 D4");
  });

  it("names its fields from the parameter (`cf.ref`)", () => {
    const { ws } = book();
    expect(() => ConditionalFormatting.add(ws, { ref: "zz", rules: [] } as never)).toThrow(
      ExcelInvalidAddressError
    );
    expect(() => ConditionalFormatting.add(ws, { ref: "A1", rules: [] })).toThrow(
      expect.objectContaining({ field: "cf.rules" })
    );
  });
});

describe("R4-C5 Watermark.get checks its handle and returns a copy", () => {
  it("refuses what is not a worksheet", () => {
    expect(() => Watermark.get(42 as never)).toThrow(ExcelArgumentError);
    expect(() => Watermark.get(null as never)).toThrow(ExcelArgumentError);
  });

  it("answers null for none, and a copy whose edits change nothing", () => {
    const { wb, ws } = book();
    expect(Watermark.get(ws)).toBeUndefined();
    const imageId = Image.add(wb, { buffer: PNG, extension: "png" });
    Watermark.set(ws, { imageId, mode: "overlay", opacity: 0.3 });
    const got = Watermark.get(ws)!;
    got.opacity = 9;
    expect(Watermark.get(ws)!.opacity).toBe(0.3);
  });

  it("keeps the deprecated Watermark.add as a forward that replaces, as set does", () => {
    const { wb, ws } = book();
    const imageId = Image.add(wb, { buffer: PNG, extension: "png" });
    Watermark.add(ws, { imageId, mode: "overlay", opacity: 0.3 });
    Watermark.add(ws, { imageId, mode: "overlay", opacity: 0.5 });
    expect(Watermark.get(ws)!.opacity).toBe(0.5);
    expect(() => Watermark.add(ws, { imageId, opacity: 5 })).toThrow(/^Watermark\.add: /);
  });
});

describe("R4-C6 Worksheet.getDimensions of an empty sheet", () => {
  it("is undefined, as Row.getLast and Column.getLast answer", () => {
    const { ws } = book();
    expect(Worksheet.getDimensions(ws)).toBeUndefined();
    Cell.setValue(ws, "B2", 1);
    expect(Range.toString(Worksheet.getDimensions(ws)!)).toBe("B2:B2");
  });
});

describe("R4-S2 a row height and a column width can be cleared", () => {
  it("with undefined, so the row or column takes the sheet's default again", () => {
    const { ws } = book();
    Row.setHeight(ws, 1, 30);
    Column.setWidth(ws, "B", 20);
    Row.setHeight(ws, 1, undefined);
    Column.setWidth(ws, "B", undefined);
    expect(Row.getHeight(ws, 1)).toBeUndefined();
    expect(Column.getWidth(ws, "B")).toBeUndefined();
    // Clearing what does not exist creates nothing.
    Row.setHeight(ws, 5, undefined);
    expect(Row.has(ws, 5)).toBe(false);
    expect(() => Row.setHeight(ws, 1, Number.NaN)).toThrow(ExcelArgumentError);
  });
});

describe("R4-S3 getPanes / getAutoFilter answer in their setters' words", () => {
  it("frozen panes round-trip through freeze", () => {
    const { wb, ws } = book();
    Worksheet.freeze(ws, { columns: 2, rows: 1 });
    const panes = Worksheet.getPanes(ws);
    expect(panes).toEqual({ state: "frozen", columns: 2, rows: 1 });
    const target = Workbook.addWorksheet(wb, "T");
    if (panes?.state === "frozen") {
      Worksheet.freeze(target, panes);
    }
    expect(Worksheet.getPanes(target)).toEqual(panes);
  });

  it("split panes round-trip through split", () => {
    const { wb, ws } = book();
    Worksheet.split(ws, { x: 2000, y: 0 });
    const panes = Worksheet.getPanes(ws);
    expect(panes).toEqual({ state: "split", x: 2000, y: 0 });
    const target = Workbook.addWorksheet(wb, "T");
    if (panes?.state === "split") {
      Worksheet.split(target, panes);
    }
    expect(Worksheet.getPanes(target)).toEqual(panes);
  });

  it("refuses panes of the other kind", () => {
    const { ws } = book();
    expect(() => Worksheet.freeze(ws, { state: "split" } as never)).toThrow(
      expect.objectContaining({ field: "options.state" })
    );
    expect(() => Worksheet.split(ws, { state: "frozen" } as never)).toThrow(ExcelArgumentError);
    expect(() => Worksheet.freeze(ws, { rows: -1 })).toThrow(
      expect.objectContaining({ field: "options.rows" })
    );
  });

  it("setAutoFilter takes back what getAutoFilter answers for none", () => {
    const { wb, ws } = book();
    Worksheet.setAutoFilter(ws, "A1:B1");
    const target = Workbook.addWorksheet(wb, "T");
    Worksheet.setAutoFilter(target, "A1:C1");
    Worksheet.setAutoFilter(target, Worksheet.getAutoFilter(Workbook.addWorksheet(wb, "None")));
    expect(Worksheet.getAutoFilter(target)).toBeUndefined();
  });
});

describe("R4-G2 the chartsheet functions take a handle", () => {
  it("remove, copy and replace take what addChartsheet returned", () => {
    const wb = Workbook.create();
    Workbook.addWorksheet(wb, "S");
    const chart = { type: "bar" as const, series: [{ values: "S!$A$1:$A$2" }] };
    const cs = Workbook.addChartsheet(wb, "C", { chart });
    const copy = Workbook.copyChartsheet(wb, cs, "C2")!;
    expect(Workbook.listChartsheets(wb)).toHaveLength(2);
    expect(Workbook.replaceChartsheetChart(wb, copy, chart)).toBe(true);
    expect(Workbook.removeChartsheet(wb, cs)).toBe(true);
    expect(Workbook.removeChartsheet(wb, cs)).toBe(false);
    expect(Workbook.listChartsheets(wb)).toHaveLength(1);
    expect(() => Workbook.removeChartsheet(wb, {} as never)).toThrow(
      expect.objectContaining({ field: "chartsheet" })
    );
  });
});

describe("R4-G3 Column.has", () => {
  it("mirrors Row.has and Cell.has, creating nothing", () => {
    const { ws } = book();
    Cell.setValue(ws, "C2", 1);
    expect(Column.has(ws, "C")).toBe(true);
    expect(Column.has(ws, "D")).toBe(false);
    Column.setWidth(ws, "E", 12);
    expect(Column.has(ws, "E")).toBe(true);
    expect(Column.has(ws, "F")).toBe(false);
  });
});

describe("R4-D2 Note.getModel and Anchor.getModel are copies", () => {
  it("editing what they return changes nothing", () => {
    const { ws } = book();
    const anchor = Anchor.create(ws, "B2");
    const col = Anchor.getCol(anchor);
    Anchor.getModel(anchor).nativeCol = 99;
    expect(Anchor.getCol(anchor)).toBe(col);
    const note = Note.create("hello");
    expect(Note.getModel(note)).not.toBe(Note.getModel(note));
  });
});

describe("R4-S5 argument errors name the parameter path and the public call", () => {
  it.each<[string, () => unknown, string, RegExp]>([
    [
      "Sparkline.add",
      () => Sparkline.add(book().ws, { type: "bogus" } as never),
      "options.type",
      /^Sparkline\.add: /
    ],
    ["Pivot.add", () => Pivot.add(book().ws, { rows: 42 } as never), "model.rows", /^Pivot\.add: /],
    [
      "Image.addShape",
      () => Image.addShape(book().ws, { type: "nope", range: "A1:B2" } as never),
      "options.type",
      /^Image\.addShape: /
    ],
    [
      "Form.setText",
      () => Form.setText(Form.addCheckbox(book().ws, "A1:B2"), 42 as never),
      "value",
      /^Form\.setText: /
    ],
    [
      "Workbook.addExternalLink",
      () => Workbook.addExternalLink(Workbook.create(), { target: "" } as never),
      "input.target",
      /^Workbook\.addExternalLink: /
    ],
    ["Range.getCount", () => Range.getCount(42 as never), "r", /^Range\.getCount: /],
    [
      "Address.encodeCell",
      () => Address.encodeCell({ c: "x", r: 0 } as never),
      "cell.c",
      /^Address\.encodeCell: /
    ],
    [
      "Workbook.copyChartsheet",
      () => Workbook.copyChartsheet(Workbook.create(), null as never),
      "chartsheet",
      /^Workbook\.copyChartsheet: /
    ],
    [
      "DataValidation.add",
      () => DataValidation.add(book().ws, "A1", null as never),
      "validation",
      /^DataValidation\.add: /
    ]
  ])("%s", (_name, run, field, message) => {
    let error: unknown;
    try {
      run();
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect((error as ExcelArgumentError).field).toBe(field);
    expect((error as Error).message).toMatch(message);
  });

  it("Table.addRow names the call and the value's position", () => {
    const { ws } = book();
    const table = Table.add(ws, { name: "T_x", ref: "A1", columns: [{ name: "a" }], rows: [] });
    expect(() => Table.addRow(table, [Number.NaN])).toThrow(
      expect.objectContaining({ field: "values[0]" })
    );
  });
});
