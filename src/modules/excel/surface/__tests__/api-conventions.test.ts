import {
  Anchor,
  Cell,
  Chart,
  Chartsheet,
  Column,
  ConditionalFormatting,
  DataValidation,
  DefinedNames,
  Form,
  Image,
  Note,
  Pivot,
  Range,
  Row,
  Sparkline,
  Stream,
  Table,
  Watermark,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  return { wb, ws: Workbook.addWorksheet(wb, "S") };
}

describe("Row — position-based row work", () => {
  it("add / addMany return row numbers and copy style from above on request", () => {
    const { ws } = sheet();
    expect(Row.add(ws, ["a"])).toBe(1);
    Row.setFont(ws, 1, { bold: true });
    expect(Row.addMany(ws, [["b"], ["c"]], { copyStyle: "above" })).toEqual([2, 3]);
    expect(Row.getFont(ws, 3)?.bold).toBe(true);
    expect(Row.getCount(ws)).toBe(3);
    expect(Row.getLast(ws)).toBe(3);
  });

  it("insert / insertMany push rows down and return their numbers", () => {
    const { ws } = sheet();
    Row.addMany(ws, [["a"], ["d"]]);
    expect(Row.insert(ws, 2, ["b"])).toBe(2);
    expect(Row.insertMany(ws, 3, [["c"]])).toEqual([3]);
    expect(
      Worksheet.getValues(ws)
        .slice(1)
        .map(r => r[1])
    ).toEqual(["a", "b", "c", "d"]);
  });

  it("duplicate takes { insert } and splice behaves like Array#splice", () => {
    const { ws } = sheet();
    Row.addMany(ws, [["a"], ["z"]]);
    Row.duplicate(ws, 1, 2, { insert: true });
    expect(Cell.getValue(ws, "A3")).toBe("a");
    expect(Cell.getValue(ws, "A4")).toBe("z");
    Row.splice(ws, 2, 2);
    expect(Cell.getValue(ws, "A2")).toBe("z");
  });

  it("each / eachAddress visit positions without creating anything", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "B2", 1);
    Cell.setValue(ws, "D2", 2);
    const rows: number[] = [];
    Row.each(ws, r => rows.push(r));
    expect(rows).toEqual([2]);
    const all: number[] = [];
    Row.each(ws, r => all.push(r), { includeEmpty: true });
    expect(all).toEqual([1, 2]);
    expect(Row.has(ws, 1)).toBe(false);
    const seen: [string, number, number][] = [];
    Row.eachAddress(ws, 2, (a, r, c) => seen.push([a, r, c]));
    expect(seen).toEqual([
      ["B2", 2, 2],
      ["D2", 2, 4]
    ]);
    expect(Cell.has(ws, "C2")).toBe(false);
  });

  it("autoFit takes a row or a span", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "A1", "x");
    expect(() => Row.autoFit(ws, 1)).not.toThrow();
    expect(() => Row.autoFit(ws, { start: 1, end: 1 })).not.toThrow();
    expect(() => Row.autoFit(ws)).not.toThrow();
  });

  it("per-facet getters read the row style", () => {
    const { ws } = sheet();
    Row.setFill(ws, 1, { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } });
    Row.setProtection(ws, 1, { locked: false });
    Row.setNumFmt(ws, 1, "0.00");
    expect(Row.getFill(ws, 1)).toMatchObject({ pattern: "solid" });
    expect(Row.getProtection(ws, 1)).toEqual({ locked: false });
    expect(Row.getNumFmt(ws, 1)).toBe("0.00");
    expect(Row.isHidden(ws, 1)).toBe(false);
  });

  it("getValues with base 1 is indexed by column number", () => {
    const { ws } = sheet();
    Row.add(ws, [1, 2]);
    expect(Row.getValues(ws, 1, { base: 1 })).toEqual([undefined, 1, 2]);
    expect(Row.getValues(ws, 1)).toEqual([1, 2]);
  });
});

describe("Column — position-based column work", () => {
  it("definitions, counts and last", () => {
    const { ws } = sheet();
    expect(Column.getLast(ws)).toBeUndefined();
    Column.setDefinitions(ws, [{ header: "A", key: "a" }, { header: "B" }]);
    expect(Column.getDefinitions(ws).map(d => d.header)).toEqual(["A", "B"]);
    expect(Column.getCount(ws)).toBe(2);
    expect(Column.getActualCount(ws)).toBe(2);
    expect(Column.getLast(ws)).toBe(2);
  });

  it("splice, eachAddress and autoFit", () => {
    const { ws } = sheet();
    Row.add(ws, [1, 2, 3]);
    Column.splice(ws, 2, 1);
    expect(Cell.getValue(ws, "B1")).toBe(3);
    Cell.setValue(ws, "A3", 9);
    const seen: string[] = [];
    Column.eachAddress(ws, "A", a => seen.push(a));
    expect(seen).toEqual(["A1", "A3"]);
    const all: string[] = [];
    Column.eachAddress(ws, 1, a => all.push(a), { includeEmpty: true });
    expect(all).toEqual(["A1", "A2", "A3"]);
    Column.autoFit(ws, "A");
    Column.autoFit(ws, { start: 1, end: 2, includeHiddenRows: true });
    expect(Column.getWidth(ws, "A")).toBeGreaterThan(0);
  });

  it("facet setters and getters mirror Cell and Row", () => {
    const { ws } = sheet();
    Column.setFont(ws, "B", { italic: true });
    Column.setAlignment(ws, "B", { horizontal: "center" });
    Column.setBorder(ws, "B", { top: { style: "thin" } });
    Column.setFill(ws, "B", { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF0000" } });
    Column.setProtection(ws, "B", { hidden: true });
    expect(Column.getFont(ws, "B")?.italic).toBe(true);
    expect(Column.getAlignment(ws, "B")?.horizontal).toBe("center");
    expect(Column.getBorder(ws, "B")?.top?.style).toBe("thin");
    expect(Column.getFill(ws, "B")).toMatchObject({ pattern: "solid" });
    expect(Column.getProtection(ws, "B")).toEqual({ hidden: true });
    expect(Column.getNumFmt(ws, "B")).toBeUndefined();
    expect(Column.isHidden(ws, "B")).toBe(false);
  });
});

describe("Cell — existence and merges by position", () => {
  it("has creates nothing; getMergeMasterAddress answers an address", () => {
    const { ws } = sheet();
    expect(Cell.has(ws, "A1")).toBe(false);
    expect(Cell.has(ws, 1, 1)).toBe(false);
    Worksheet.merge(ws, "B2:C3");
    expect(Cell.getMergeMasterAddress(ws, "C3")).toBe("B2");
    expect(Cell.getMergeMasterAddress(ws, 3, 3)).toBe("B2");
    expect(Cell.getMergeMasterAddress(ws, "E5")).toBe("E5");
    expect(Cell.has(ws, "E5")).toBe(false);
  });
});

describe("collections", () => {
  it("removers answer whether anything was removed", () => {
    const { wb, ws } = sheet();
    expect(Table.remove(ws, "Nope")).toBe(false);
    expect(Table.get(ws, "Nope")).toBeUndefined();
    expect(Watermark.remove(ws)).toBe(false);
    expect(Watermark.get(ws)).toBeNull();
    expect(Workbook.removeWorksheet(wb, "Nope")).toBe(false);
    expect(Workbook.removeWorksheet(wb, "S")).toBe(true);
    const dv = DataValidation.create();
    expect(DataValidation.remove(dv, "A1")).toBe(false);
    DataValidation.add(dv, "A1", { type: "whole", operator: "equal", formulae: [1] });
    expect(DataValidation.remove(dv, "A1")).toBe(true);
    const dn = DefinedNames.create();
    DefinedNames.add(dn, "S!A1", "n");
    expect(DefinedNames.remove(dn, "S!A1", "n")).toBe(true);
    expect(DefinedNames.remove(dn, "S!A1", "n")).toBe(false);
    expect(DefinedNames.list(dn)).toEqual([{ name: "n", ranges: [] }]);
  });

  it("ConditionalFormatting add / list / remove", () => {
    const { ws } = sheet();
    const cf = ConditionalFormatting.add(ws, {
      ref: "A1:A3",
      rules: [{ type: "expression", formulae: ["A1>1"], priority: 1 }]
    });
    expect(ConditionalFormatting.list(ws)).toEqual([cf]);
    expect(ConditionalFormatting.remove(ws, 0)).toBe(true);
    expect(ConditionalFormatting.remove(ws, 0)).toBe(false);
  });

  it("Pivot.list, Chart.list, Workbook.listWorksheets", () => {
    const { wb, ws } = sheet();
    expect(Pivot.list(ws)).toEqual([]);
    expect(Chart.list(ws)).toEqual([]);
    expect(Workbook.listWorksheets(wb)).toEqual([ws]);
    expect(Workbook.listChartsheets(wb)).toEqual([]);
    expect(Worksheet.getAutoFilter(ws)).toBeUndefined();
  });
});

describe("Form — the VML getters take the handle like the other getters", () => {
  it("getVmlAnchor / getVmlStyle / getVmlCheckedValue accept a handle or its model", () => {
    const { ws } = sheet();
    const cb = Form.addCheckbox(ws, "B2:C3", { text: "Ok", checked: true });
    expect(Form.getVmlAnchor(cb)).toBe(Form.getVmlAnchor(cb.model));
    expect(Form.getVmlAnchor(cb)).toMatch(/^\d+, \d+, \d+, \d+, \d+, \d+, \d+, \d+$/);
    expect(Form.getVmlStyle(cb)).toBe(Form.getVmlStyle(cb.model));
    expect(Form.getVmlStyle(cb)).toContain("position:absolute");
    expect(Form.getVmlCheckedValue(cb)).toBe(1);
    expect(Form.getVmlCheckedValue(cb.model)).toBe(1);
    // The deprecated names keep taking the model, as they always did.
    expect(Form.vmlAnchor(cb.model)).toBe(Form.getVmlAnchor(cb));
    expect(Form.vmlStyle(cb.model)).toBe(Form.getVmlStyle(cb));
    expect(Form.vmlCheckedValue(cb.model)).toBe(1);
  });
});

describe("deprecated forwards", () => {
  it("pure aliases point at the new names", () => {
    expect(Table.name).toBe(Table.getName);
    expect(Table.columnSetName).toBe(Table.setColumnName);
    expect(Chart.get).toBe(Chart.list);
    expect(Chart.toSVG).toBe(Chart.toSvg);
    expect(Chart.toPNG).toBe(Chart.toPng);
    expect(Chart.title).toBe(Chart.getTitle);
    expect(Chart.totalSeriesCount).toBe(Chart.getTotalSeriesCount);
    expect(Chartsheet.name).toBe(Chartsheet.getName);
    // Not an alias: `getModel` returns a copy; the deprecated name keeps returning the live model.
    expect(Chartsheet.model).not.toBe(Chartsheet.getModel);
    expect(Anchor.col).toBe(Anchor.getCol);
    expect(Anchor.model).toBe(Anchor.getModel);
    expect(Form.checked).toBe(Form.isChecked);
    // `fromCheckboxModel` was never shipped: the handle it built was not on the sheet.
    expect("fromCheckboxModel" in Form).toBe(false);
    expect(Image.model).toBe(Image.getModel);
    // Not an alias: `setModel` also records that a picture the image stops showing is no longer
    // used here; the deprecated name keeps the old behaviour of only applying the model.
    expect(Image.applyModel).not.toBe(Image.setModel);
    expect(Image.getShapes).toBe(Image.listShapes);
    expect(Note.model).toBe(Note.getModel);
    expect(DefinedNames.getAllEntries).toBe(DefinedNames.list);
    expect(DefinedNames.model).toBe(DefinedNames.getModel);
    expect(Range.forEachAddress).toBe(Range.eachAddress);
    expect(Range.count).toBe(Range.getCount);
    expect(Workbook.getWorksheets).toBe(Workbook.listWorksheets);
    expect(Workbook.getChartsheets).toBe(Workbook.listChartsheets);
    // Not an alias: `getDimensions` answers `undefined` for an empty sheet; the deprecated name keeps its range of zeros.
    expect(Worksheet.dimensions).not.toBe(Worksheet.getDimensions);
    expect(Worksheet.mergedRegions).toBe(Worksheet.getMergedRegions);
    // Not an alias: `getPanes` hands out a copy; the deprecated name keeps handing out the live view.
    expect(Worksheet.panes).not.toBe(Worksheet.getPanes);
    expect(Worksheet.addConditionalFormatting).toBe(ConditionalFormatting.add);
    expect(Row.getHidden).toBe(Row.isHidden);
    expect(Column.getHidden).toBe(Column.isHidden);
    // Not an alias: `getRowValues` reads 0-based by default, as `Row.getValues` does; the deprecated name stays 1-based.
    expect(Stream.rowValues).not.toBe(Stream.getRowValues);
    expect(Stream.rowCell).toBe(Stream.getRowCell);
    expect(Sparkline.list).toBeTypeOf("function");
  });

  it("moved Worksheet members keep their old behaviour", () => {
    const { ws } = sheet();
    const row = Worksheet.addRow(ws, ["a", "b"]);
    expect(row.number).toBe(1);
    expect(Worksheet.addRows(ws, [["c"]])[0].number).toBe(2);
    expect(Worksheet.getRow(ws, 1)).toBe(row);
    expect(Worksheet.findRow(ws, 9)).toBeUndefined();
    expect(Worksheet.lastRow(ws)).toBe(Worksheet.getRow(ws, 2));
    expect(Worksheet.rowCount(ws)).toBe(Row.getCount(ws));
    expect(Worksheet.columnCount(ws)).toBe(Column.getCount(ws));
    expect(Worksheet.autoFitColumns(ws)).toBe(ws);
    expect(Worksheet.autoFilter(ws)).toBeNull();
    Worksheet.setColumns(ws, [{ header: "H" }]);
    expect(Worksheet.columnDefinitions(ws)).toEqual(Column.getDefinitions(ws));
    const seen: number[] = [];
    Worksheet.eachRow(ws, (_r, n) => seen.push(n));
    expect(seen).toEqual([1, 2]);
  });

  it("handle-returning Row / Cell members keep their old behaviour", () => {
    const { ws } = sheet();
    Cell.setValue(ws, "A1", 5);
    Worksheet.merge(ws, "B1:C1");
    const cell = Row.getCell(ws, 1, 1);
    expect(Cell.view(cell).value).toBe(5);
    expect(Stream.getCellView(cell).value).toBe(5);
    expect(Cell.find(ws, "A1")).toBe(cell);
    expect(Cell.getMergeMaster(ws, "C1").address).toBe("B1");
    const cols: number[] = [];
    Row.eachCell(ws, 1, (_c, n) => cols.push(n));
    expect(cols).toEqual([1, 3]);
  });
});

describe("deprecated value forwards keep their reading", () => {
  it("Row.values / Column.values are the base-1 reads", () => {
    const { ws } = sheet();
    Row.add(ws, [1, 2]);
    expect(Row.values(ws, 1)).toEqual(Row.getValues(ws, 1, { base: 1 }));
    expect(Column.values(ws, "B")).toEqual(Column.getValues(ws, "B", { base: 1 }));
    expect(Column.values(ws, "B")).toEqual([, 2]);
  });
});

describe("ConditionalFormatting.remove / clear", () => {
  const rule = { type: "expression" as const, formulae: ["TRUE"], priority: 1 };
  it("remove takes an entry, a position or a predicate; clear takes everything", () => {
    const { ws } = sheet();
    const a = ConditionalFormatting.add(ws, { ref: "A1", rules: [rule] });
    const b = ConditionalFormatting.add(ws, { ref: "B1", rules: [rule] });
    const c = ConditionalFormatting.add(ws, { ref: "C1", rules: [rule] });
    expect(ConditionalFormatting.remove(ws, b)).toBe(true);
    expect(ConditionalFormatting.remove(ws, b)).toBe(false);
    expect(ConditionalFormatting.remove(ws, 5)).toBe(false);
    expect(ConditionalFormatting.remove(ws, -1)).toBe(false);
    expect(ConditionalFormatting.list(ws)).toEqual([a, c]);
    expect(ConditionalFormatting.remove(ws, e => e.ref === "C1")).toBe(true);
    expect(ConditionalFormatting.remove(ws, 0)).toBe(true);
    ConditionalFormatting.add(ws, { ref: "D1", rules: [rule] });
    ConditionalFormatting.clear(ws);
    expect(ConditionalFormatting.list(ws)).toEqual([]);
  });

  it("the deprecated Worksheet form keeps clearing without a filter", () => {
    const { ws } = sheet();
    ConditionalFormatting.add(ws, { ref: "A1", rules: [rule] });
    expect(Worksheet.removeConditionalFormatting(ws)).toBeUndefined();
    expect(ConditionalFormatting.list(ws)).toEqual([]);
  });

  it("the deprecated Worksheet form keeps its baseline results", () => {
    const { ws } = sheet();
    ConditionalFormatting.add(ws, { ref: "A1", rules: [rule] });
    ConditionalFormatting.add(ws, { ref: "B1", rules: [rule] });
    expect(Worksheet.removeConditionalFormatting(ws, 0)).toBeUndefined();
    expect(Worksheet.removeConditionalFormatting(ws, () => true)).toBeUndefined();
    expect(ConditionalFormatting.list(ws)).toEqual([]);
    // Anything but an index or a predicate clears, as before — whatever `ws` holds.
    const bare = {} as Parameters<typeof Worksheet.removeConditionalFormatting>[0];
    expect(Worksheet.removeConditionalFormatting(bare)).toBeUndefined();
  });
});
