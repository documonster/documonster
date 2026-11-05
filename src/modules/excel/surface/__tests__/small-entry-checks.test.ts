/**
 * The small stateless entry points — `Address`, `Range`, `Note`, `Format` — and a few single setters refuse an input
 * they cannot use with the module's typed error, rather than answering with `undefined`, `NaN` or a raw `TypeError`.
 */
import {
  Address,
  Cell,
  Column,
  ExcelColumnOutOfBoundsError,
  ExcelArgumentError,
  ExcelError,
  Format,
  ExcelInvalidAddressError,
  Note,
  Range,
  Row,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

describe("Address", () => {
  it.each<[string, () => unknown]>([
    ["encodeCol(1.5)", () => Address.encodeCol(1.5)],
    ["encodeCol(-1)", () => Address.encodeCol(-1)],
    ["encodeRow(-5)", () => Address.encodeRow(-5)],
    ["encodeRow(1.5)", () => Address.encodeRow(1.5)],
    ["encodeRow(1048576)", () => Address.encodeRow(1048576)],
    ["decodeRow('abc')", () => Address.decodeRow("abc")],
    ["decodeRow('0')", () => Address.decodeRow("0")],
    ["decodeCell('zz')", () => Address.decodeCell("zz")],
    ["decodeCell('A0')", () => Address.decodeCell("A0")],
    ["decodeRange('A1:zz')", () => Address.decodeRange("A1:zz")],
    ["encodeCell({ r: -1, c: 0 })", () => Address.encodeCell({ r: -1, c: 0 })]
  ])("%s is refused", (_name, call) => {
    expect(call).toThrow(ExcelError);
  });

  it("keeps answering for the grid", () => {
    expect(Address.encodeCol(16383)).toBe("XFD");
    expect(Address.encodeRow(1048575)).toBe("1048576");
    expect(Address.decodeCell("$B$3")).toEqual({ c: 1, r: 2 });
    expect(Address.decodeCol("a")).toBe(0);
    expect(Address.decodeRange("A1")).toEqual({ s: { c: 0, r: 0 }, e: { c: 0, r: 0 } });
  });
});

describe("Range.create", () => {
  it.each<[string, () => unknown]>([
    ["'garbage'", () => Range.create("garbage")],
    ["'A0:B2'", () => Range.create("A0:B2")],
    ["NaN", () => Range.create(NaN, 1, 2, 2)],
    ["negatives", () => Range.create(-1, -1, -5, -5)],
    ["past the last row", () => Range.create(1, 1, 1048577, 1)]
  ])("refuses %s", (_name, call) => {
    expect(call).toThrow(ExcelError);
  });

  it("still makes the empty range an accumulator starts from", () => {
    const r = Range.create();
    Range.expandToAddress(r, "B2");
    expect(Range.toString(r)).toBe("B2:B2");
  });

  it("refuses an address it cannot expand to", () => {
    const r = Range.create("A1");
    expect(() => Range.expandToAddress(r, "zz")).toThrow(ExcelError);
    expect(() => Range.expand(r, NaN, 1, 1, 1)).toThrow(ExcelError);
    expect(Range.toString(r)).toBe("A1:A1");
  });
});

describe("Note.fromModel and Format.parseValue", () => {
  it("refuse what is not a model or not text", () => {
    expect(() => Note.fromModel(null as never)).toThrow(ExcelArgumentError);
    expect(() => Note.fromModel({} as never)).toThrow(ExcelArgumentError);
    expect(() => Format.parseValue(42 as never, "1")).toThrow(ExcelArgumentError);
    expect(() => Format.parseValue("0", null as never)).toThrow(ExcelArgumentError);
  });
});

describe("Cell.setResult", () => {
  it("refuses a result for a cell that holds no formula, instead of dropping it", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 5);
    expect(() => Cell.setResult(ws, "A1", 7)).toThrow(ExcelArgumentError);
    expect(Cell.has(ws, "B1")).toBe(false);
    expect(() => Cell.setResult(ws, "B1", 7)).toThrow(ExcelArgumentError);
    expect(Cell.has(ws, "B1")).toBe(false);
    Cell.setValue(ws, "C1", { formula: "1+1" });
    Cell.setResult(ws, "C1", 2);
    expect(Cell.getResult(ws, "C1")).toBe(2);
  });
});

describe("Worksheet.toJson", () => {
  it("refuses a range it cannot read and a header form it does not know", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    expect(() => Worksheet.toJson(ws, { range: "zz" })).toThrow(ExcelInvalidAddressError);
    expect(() => Worksheet.toJson(ws, { header: "bogus" as never })).toThrow(ExcelArgumentError);
    expect(() => Worksheet.toJson(ws, { range: -1 })).toThrow(ExcelArgumentError);
  });
});

describe("Address error classes", () => {
  it("a string that is not a reference is an ExcelInvalidAddressError; a column past XFD is out of bounds", () => {
    for (const bad of ["garbage", "zz", "", "1A", "A1:", "GARBAGE"]) {
      expect(() => Address.decodeCell(bad), bad).toThrow(ExcelInvalidAddressError);
    }
    expect(() => Address.decodeRange("garbage")).toThrow(ExcelInvalidAddressError);
    expect(() => Address.decodeRange("A1:zz")).toThrow(ExcelInvalidAddressError);
    expect(() => Address.decodeCol("garbage")).toThrow(ExcelInvalidAddressError);
    expect(() => Address.decodeCell("XFE1")).toThrow(ExcelColumnOutOfBoundsError);
    expect(() => Address.decodeCol("XFE")).toThrow(ExcelColumnOutOfBoundsError);
  });

  it("names what it was given when it is not a cell object, and refuses a sheet name that is not text", () => {
    expect(() => Address.encodeRange(0 as never, 0 as never)).toThrow(ExcelArgumentError);
    expect(() => Address.encodeCell(null as never)).toThrow(ExcelArgumentError);
    expect(() => Address.quoteSheetName(42 as never)).toThrow(ExcelArgumentError);
  });

  it("reads lowercase as uppercase", () => {
    expect(Address.decodeCell("b3")).toEqual({ c: 1, r: 2 });
    expect(Address.decodeRange("a1:b2")).toEqual({ s: { c: 0, r: 0 }, e: { c: 1, r: 1 } });
  });
});

describe("callbacks and handles that are not what they must be", () => {
  it("are argument errors, not raw TypeErrors", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Cell.setValue(ws, "A1", 1);
    expect(() => Workbook.eachWorksheet(wb, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Row.each(ws, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Row.eachAddress(ws, 1, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Column.each(ws, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Column.eachAddress(ws, 1, 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Workbook.copyChartsheet(wb, null as never)).toThrow(ExcelArgumentError);
    expect(() => Workbook.getChartsheet(wb, null as never)).toThrow(ExcelArgumentError);
    expect(() => Workbook.removeChartsheet(wb, {} as never)).toThrow(ExcelArgumentError);
    expect(() => Workbook.importWorksheet(wb, null as never)).toThrow(ExcelArgumentError);
  });
});
