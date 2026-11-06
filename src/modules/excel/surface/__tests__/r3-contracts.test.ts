/**
 * Round-3 review fixes: references are checked and normalised at the boundary, rule collections and pivot caches
 * stay consistent with the sheet, and boundaries raise the module's own errors.
 */
import { Archive } from "@archive";
import {
  Anchor,
  Cell,
  Chart,
  Chartsheet,
  Column,
  Form,
  DataValidation,
  ExcelArgumentError,
  Image,
  ExcelInvalidAddressError,
  Pivot,
  Row,
  Table,
  Workbook,
  Worksheet,
  ExcelWorksheetNameError
} from "@excel";
import { ExcelFormula } from "@excel/bridge/formula";
import { ChartClassic, ExcelChartOptionsError } from "@excel/chart";
import { describe, expect, it } from "vitest";

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
  ),
  c => c.charCodeAt(0)
);

async function part(bytes: Uint8Array, name: RegExp): Promise<string[]> {
  const files = await Archive.unpackZip(bytes);
  return [...files]
    .filter(([path]) => name.test(path))
    .map(([, f]) => new TextDecoder().decode(f.data));
}

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  const other = Workbook.addWorksheet(wb, "Other");
  return { wb, ws, other };
}

describe("R3-B4 a table's ref is a cell or range of its own sheet", () => {
  it.each(["Other!A1", "'My Sheet'!B2", "S!C3", "zz"])(
    "Table.add refuses %j and writes nothing",
    ref => {
      const { ws } = book();
      expect(() =>
        Table.add(ws, { name: "T", ref, columns: [{ name: "h" }], rows: [[1]] })
      ).toThrow(ExcelInvalidAddressError);
      expect(Table.list(ws)).toEqual([]);
    }
  );

  it("setRef refuses a sheet qualifier; lowercase is read as uppercase", () => {
    const { ws } = book();
    const t = Table.add(ws, { name: "T", ref: "b2", columns: [{ name: "h" }], rows: [[1]] });
    expect(Table.getRef(t)).toBe("B2:B3");
    expect(Cell.getValue(ws, "B2")).toBe("h");
    expect(() => Table.setRef(t, "Other!H1")).toThrow(ExcelInvalidAddressError);
    expect(Table.getRef(t)).toBe("B2:B3");
    Table.setRef(t, "d4");
    expect(Table.getRef(t)).toBe("D4:D5");
  });
});

describe("R3-B5 lowercase ranges are accepted and normalised", () => {
  it("Chart.place and Image.place", async () => {
    const { wb, ws } = book();
    Worksheet.addAoa(ws, [
      ["k", "v"],
      ["a", 1],
      ["b", 2]
    ]);
    const chart = Chart.place(ws, "d5:h10", {
      type: "bar",
      barDir: "col",
      series: [{ values: "S!$B$2:$B$3", categories: "S!$A$2:$A$3" }]
    });
    expect(Chart.getAnchorModel(chart).range).toMatchObject({
      tl: { nativeCol: 3, nativeRow: 4 },
      br: { nativeCol: 8, nativeRow: 10 }
    });
    const id = Image.add(wb, { buffer: PNG, extension: "png" });
    Image.place(ws, id, "b2:c3");
    Image.place(ws, id, { tl: "b2", br: "c3" });
    for (const image of Image.list(ws)) {
      expect((Image.getModel(image) as { range?: unknown }).range).toMatchObject({
        tl: { nativeCol: 1, nativeRow: 1 },
        br: { nativeCol: 3, nativeRow: 3 }
      });
    }
    await expect(Workbook.toBuffer(wb)).resolves.toBeInstanceOf(Uint8Array);
  });
});

describe("R3-G3 Image.place returns the placed image", () => {
  it("the handle Image.list lists", () => {
    const { wb, ws } = book();
    const id = Image.add(wb, { buffer: PNG, extension: "png" });
    const placed = Image.place(ws, id, "B2:C3");
    expect(Image.list(ws)).toEqual([placed]);
    expect(Image.remove(ws, placed!)).toBe(true);
  });
});

describe("R3-C4 Anchor setters and create check what they are given", () => {
  it.each([Number.NaN, -5, "x", 1.5e9])("setCol / setRow refuse %j", value => {
    const { ws } = book();
    const a = Anchor.create(ws, "B2");
    expect(() => Anchor.setCol(a, value as never)).toThrow(ExcelArgumentError);
    expect(() => Anchor.setRow(a, value as never)).toThrow(ExcelArgumentError);
    expect(Anchor.getModel(a)).toMatchObject({ nativeCol: 2, nativeRow: 2 });
  });

  it("create refuses an address that is not a cell, a sheet qualifier and a bad offset", () => {
    const { ws } = book();
    expect(() => Anchor.create(ws, "zz")).toThrow(ExcelInvalidAddressError);
    expect(() => Anchor.create(ws, "Other!D5")).toThrow(ExcelInvalidAddressError);
    expect(() => Anchor.create(ws, "B2", { offset: Number.NaN })).toThrow(ExcelArgumentError);
    expect(() => Anchor.create(ws, { col: Number.NaN, row: 1 })).toThrow(ExcelArgumentError);
    expect(Anchor.getModel(Anchor.create(ws, "b2"))).toMatchObject({ nativeCol: 2, nativeRow: 2 });
  });

  it("setModel refuses a model that is not one", () => {
    const { ws } = book();
    const a = Anchor.create(ws, "B2");
    expect(() => Anchor.setModel(a, null as never)).toThrow(ExcelArgumentError);
    expect(() =>
      Anchor.setModel(a, { nativeCol: -1, nativeColOff: 0, nativeRow: 0, nativeRowOff: 0 })
    ).toThrow(ExcelArgumentError);
    Anchor.setModel(a, { nativeCol: 4, nativeColOff: 10, nativeRow: 3, nativeRowOff: 0 });
    expect(Anchor.getCol(a)).toBeGreaterThan(4);
  });
});

describe("R3-B8 DataValidation stores and returns copies", () => {
  it("one options object on two sheets makes two rules; readers return copies", async () => {
    const { wb, ws, other } = book();
    const rule = { type: "whole" as const, operator: "between" as const, formulae: [1, 10] };
    const added = DataValidation.add(ws, "B2:B5", rule);
    DataValidation.add(other, "B2:B5", rule);
    expect(added).not.toBe(rule);
    rule.formulae[1] = 99;
    added.type = "decimal" as never;
    (DataValidation.find(ws, "B3") as { operator?: string }).operator = "notBetween";
    expect(DataValidation.find(ws, "B3")).toMatchObject({
      type: "whole",
      operator: "between",
      formulae: [1, 10]
    });
    expect(DataValidation.find(other, "B3")).toMatchObject({ formulae: [1, 10] });
    const sheets = await part(await Workbook.toBuffer(wb), /worksheets\/sheet\d\.xml$/);
    expect(sheets.join("")).not.toMatch(/decimal|99|notBetween/);
  });

  it("Cell.setValidation and Cell.getValidation copy too; adjacent cells still write one rule", async () => {
    const { wb, ws } = book();
    const rule = { type: "list" as const, formulae: ['"a,b"'] };
    Cell.setValidation(ws, "A1", rule);
    Cell.setValidation(ws, "A2", rule);
    rule.formulae[0] = '"x"';
    Cell.getValidation(ws, "A1")!.type = "bogus" as never;
    expect(Cell.getValidation(ws, "A1")).toMatchObject({ type: "list", formulae: ['"a,b"'] });
    const [sheet] = await part(await Workbook.toBuffer(wb), /worksheets\/sheet1\.xml$/);
    expect(sheet).toMatch(/<dataValidation type="list"[^>]*sqref="A1:A2"/);
    expect(sheet).not.toMatch(/bogus|"x"/);
  });

  it("list and clear", () => {
    const { ws } = book();
    DataValidation.add(ws, "A1", { type: "whole", operator: "greaterThan", formulae: [1] });
    DataValidation.add(ws, "C1:C3", { type: "list", formulae: ['"a"'] });
    expect(DataValidation.list(ws)).toEqual([
      { ref: "A1", rule: { type: "whole", operator: "greaterThan", formulae: [1] } },
      { ref: "C1:C3", rule: { type: "list", formulae: ['"a"'] } }
    ]);
    DataValidation.clear(ws);
    expect(DataValidation.list(ws)).toEqual([]);
    expect(DataValidation.find(ws, "A1")).toBeUndefined();
  });
});

describe("R3-C1 sheet-qualified targets are refused", () => {
  it("DataValidation.add / remove / find", () => {
    const { ws } = book();
    const rule = { type: "whole" as const, operator: "greaterThan" as const, formulae: [1] };
    expect(() => DataValidation.add(ws, "Other!E1", rule)).toThrow(ExcelInvalidAddressError);
    expect(() => DataValidation.remove(ws, "Other!E1")).toThrow(ExcelInvalidAddressError);
    expect(() => DataValidation.find(ws, "Other!E1")).toThrow(ExcelInvalidAddressError);
    expect(DataValidation.find(ws, "E1")).toBeUndefined();
    DataValidation.add(ws, "e1", rule);
    expect(DataValidation.find(ws, "E1")).toMatchObject({ type: "whole" });
  });

  it("Pivot.add ref", () => {
    const { ws, other } = book();
    Worksheet.addAoa(ws, [
      ["Region", "Amt"],
      ["N", 1]
    ]);
    expect(() =>
      Pivot.add(other, { sourceSheet: ws, rows: ["Region"], values: ["Amt"], ref: "Other!D5" })
    ).toThrow(ExcelInvalidAddressError);
    expect(Pivot.list(other)).toEqual([]);
    expect(
      Pivot.add(other, { sourceSheet: ws, rows: ["Region"], values: ["Amt"], ref: "d5" }).ref
    ).toBe("D5");
  });
});

describe("R3-B2 a sheet-sourced pivot follows edits to its source", () => {
  async function written(edit: (data: Worksheet.Handle) => void) {
    const wb = Workbook.create();
    const data = Workbook.addWorksheet(wb, "Data");
    Worksheet.addAoa(data, [
      ["Region", "Amt"],
      ["N", 1],
      ["S", 2]
    ]);
    const out = Workbook.addWorksheet(wb, "Out");
    Pivot.add(out, { sourceSheet: data, rows: ["Region"], values: ["Amt"], ref: "A3" });
    edit(data);
    const bytes = await Workbook.toBuffer(wb);
    const [definition] = await part(bytes, /pivotCacheDefinition1\.xml$/);
    const [records] = await part(bytes, /pivotCacheRecords1\.xml$/);
    return { definition, records, back: await Workbook.read(bytes) };
  }

  it("a new item in a source cell", async () => {
    const { definition } = await written(data => Cell.setValue(data, "A3", "W"));
    expect(definition).toMatch(/<s v="W" ?\/>/);
    expect(definition).toMatch(/<worksheetSource ref="A1:B3" sheet="Data"\/>/);
  });

  it("a row inserted above the source moves it", async () => {
    const { definition } = await written(data => Row.insert(data, 1, ["new"]));
    expect(definition).toMatch(/<worksheetSource ref="A2:B4" sheet="Data"\/>/);
    expect(definition).toMatch(/<cacheField name="Region"/);
  });

  it("a row added below the source is not part of it", async () => {
    const { definition } = await written(data => Row.insert(data, 5, ["below"]));
    expect(definition).toMatch(/<worksheetSource ref="A1:B3" sheet="Data"\/>/);
    expect(definition).not.toMatch(/below/);
  });

  it("a column inserted before the source moves it without widening the cache", async () => {
    const { definition } = await written(data => Column.splice(data, 1, 0, ["x"]));
    expect(definition).toMatch(/<worksheetSource ref="B1:C3" sheet="Data"\/>/);
    expect([...definition.matchAll(/<cacheField name="([^"]*)"/g)].map(m => m[1])).toEqual([
      "Region",
      "Amt"
    ]);
  });

  it("a renamed header renames the field", async () => {
    const { definition } = await written(data => Cell.setValue(data, "A1", "Area"));
    expect(definition).toMatch(/<cacheField name="Area"/);
  });
});

describe("R3-C2 the 31-character rule on every entry that names a sheet", () => {
  it("Chartsheet.setName", () => {
    const { wb, ws } = book();
    Worksheet.addAoa(ws, [
      ["k", "v"],
      ["a", 1]
    ]);
    const cs = Workbook.addChartsheet(wb, "CS", {
      chart: { type: "bar", barDir: "col", series: [{ values: "S!$B$2:$B$2" }] }
    });
    expect(() => Chartsheet.setName(cs, "c".repeat(40))).toThrow(ExcelWorksheetNameError);
    expect(Chartsheet.getName(cs)).toBe("CS");
  });
});

describe("R3-C5 typed errors at the boundary", () => {
  it("ExcelFormula.calculate", () => {
    expect(() => ExcelFormula.calculate(null as never)).toThrow(ExcelArgumentError);
    expect(() => ExcelFormula.calculate(42 as never)).toThrow(ExcelArgumentError);
    expect(() => ExcelFormula.calculate({} as never)).toThrow(ExcelArgumentError);
  });

  it("Chart.addSeries and Chart.setUserShapesXml", () => {
    const { ws } = book();
    Worksheet.addAoa(ws, [
      ["k", "v"],
      ["a", 1]
    ]);
    const chart = Chart.place(ws, "D1:H10", {
      type: "bar",
      barDir: "col",
      series: [{ values: "S!$B$2:$B$2" }]
    });
    expect(() => Chart.addSeries(chart, null as never)).toThrow(ExcelArgumentError);
    expect(() => Chart.setUserShapesXml(chart, 42 as never)).toThrow(ExcelArgumentError);
    expect(Chart.getSeriesCount(chart)).toBe(1);
  });

  it("ChartClassic.toSvg", () => {
    expect(() => ChartClassic.toSvg({} as never)).toThrow(ExcelChartOptionsError);
    expect(() => ChartClassic.toSvg(null as never)).toThrow(ExcelChartOptionsError);
  });
});

describe("R3-D2 error text", () => {
  it("Table.removeRows on a table with no data row, and a form link that is not a reference", () => {
    const { ws } = book();
    const t = Table.add(ws, { name: "T", ref: "A1", columns: [{ name: "h" }], rows: [] });
    expect(() => Table.removeRows(t, 50)).toThrow(/has no data row at 50/);
    expect(() => Form.addCheckbox(ws, "B2", { link: "not a ref!!" })).toThrow(
      /Invalid address "not a ref!!"/
    );
  });
});
