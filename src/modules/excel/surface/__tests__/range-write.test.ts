/**
 * `Range.setValues` and the range style setters write every cell of a range as the matching `Cell.*` call writes one,
 * and check everything before writing anything.
 */
import {
  Cell,
  ExcelArgumentError,
  ExcelInvalidAddressError,
  Range,
  Row,
  Workbook
} from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  return Workbook.addWorksheet(wb, "S");
}

describe("Range.setValues", () => {
  it("writes from the top-left cell and reads back with Range.getValues", () => {
    const ws = sheet();
    Range.setValues(ws, "B2:C3", [
      ["a", 1],
      [true, null]
    ]);
    expect(Range.getValues(ws, "B2:C3")).toEqual([
      ["a", 1],
      [true, null]
    ]);
    expect(Cell.has(ws, "A1")).toBe(false);
  });

  it("takes a range handle", () => {
    const ws = sheet();
    Range.setValues(ws, Range.create("A1:B1"), [[1, 2]]);
    expect(Range.getValues(ws, "A1:B1")).toEqual([[1, 2]]);
  });

  it("leaves the cells shorter rows do not reach as they were", () => {
    const ws = sheet();
    Cell.setValue(ws, "B1", "keep");
    Cell.setValue(ws, "B2", "keep too");
    Range.setValues(ws, "A1:B3", [["x"], ["y"]]);
    expect(Range.getValues(ws, "A1:B3")).toEqual([
      ["x", "keep"],
      ["y", "keep too"],
      [null, null]
    ]);
    expect(Cell.has(ws, "A3")).toBe(false);
  });

  it("writes formulas, dates and rich text as Cell.setValue does", () => {
    const ws = sheet();
    const date = new Date(Date.UTC(2024, 0, 15));
    Range.setValues(ws, "A1:C1", [[{ formula: "1+1" }, date, { richText: [{ text: "r" }] }]]);
    expect(Cell.getFormula(ws, "A1")).toBe("1+1");
    expect(Cell.getValue(ws, "B1")).toEqual(date);
    expect(Cell.getText(ws, "C1")).toBe("r");
  });

  it.each<[string, unknown, string]>([
    ["too many rows", [[1], [2], [3]], "values"],
    ["a row too long", [[1, 2, 3]], "values[0]"],
    ["a row that is not an array", [[1], 2], "values[1]"],
    ["not an array", "nope", "values"],
    [
      "a value Cell.setValue refuses",
      [
        [1, 2],
        [3, Number.NaN]
      ],
      "values[1][1]"
    ]
  ])("refuses %s and writes nothing", (_name, values, field) => {
    const ws = sheet();
    let error: unknown;
    try {
      Range.setValues(ws, "A1:B2", values as never);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect((error as ExcelArgumentError).field).toBe(field);
    expect((error as Error).message).toMatch(/^Range\.setValues: /);
    expect(Row.getCount(ws)).toBe(0);
  });

  it.each(["A:A", "garbage", 42, null])("refuses the range %s", range => {
    const ws = sheet();
    expect(() => Range.setValues(ws, range as never, [[1]])).toThrow(ExcelArgumentError);
    expect(Row.getCount(ws)).toBe(0);
  });

  it("names the call and the parameter for a malformed range", () => {
    const ws = sheet();
    expect(() => Range.setValues(ws, "A:A", [[1]])).toThrow(
      expect.objectContaining({
        name: "ExcelInvalidAddressError",
        field: "range",
        message: expect.stringMatching(/^Range\.setValues: /)
      })
    );
    expect(() => Range.setValues(ws, "A:A", [[1]])).toThrow(ExcelInvalidAddressError);
  });

  it("writes a large range in linear time", () => {
    const ws = sheet();
    const values = Array.from({ length: 500 }, (_, r) =>
      Array.from({ length: 40 }, (_, c) => r * 40 + c)
    );
    const start = performance.now();
    Range.setValues(ws, "A1:AN500", values);
    expect(performance.now() - start).toBeLessThan(2000);
    expect(Cell.getValue(ws, "AN500")).toBe(19999);
  });
});

describe("Range style setters", () => {
  it("Range.setStyle merges into every cell as Cell.setStyle does", () => {
    const ws = sheet();
    Cell.setNumFmt(ws, "A1", "0.0");
    Range.setStyle(ws, "A1:B2", { font: { bold: true } });
    for (const addr of ["A1", "B1", "A2", "B2"]) {
      expect(Cell.getFont(ws, addr)).toEqual({ bold: true });
    }
    expect(Cell.getNumFmt(ws, "A1")).toBe("0.0");
  });

  it("the facet setters replace the facet on every cell, and undefined clears it", () => {
    const ws = sheet();
    Range.setFont(ws, "A1:B1", { italic: true });
    Range.setFill(ws, "A1:B1", {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF0000" }
    });
    Range.setBorder(ws, "A1:B1", { top: { style: "thin" } });
    Range.setAlignment(ws, "A1:B1", { wrapText: true });
    Range.setNumFmt(ws, "A1:B1", "0.00");
    Range.setProtection(ws, "A1:B1", { locked: false });
    for (const addr of ["A1", "B1"]) {
      expect(Cell.getFont(ws, addr)).toEqual({ italic: true });
      expect(Cell.getFill(ws, addr)).toMatchObject({ pattern: "solid" });
      expect(Cell.getBorder(ws, addr)).toEqual({ top: { style: "thin" } });
      expect(Cell.getAlignment(ws, addr)).toEqual({ wrapText: true });
      expect(Cell.getNumFmt(ws, addr)).toBe("0.00");
      expect(Cell.getProtection(ws, addr)).toEqual({ locked: false });
    }
    Range.setFont(ws, "A1:B1", undefined);
    expect(Cell.getFont(ws, "B1")).toBeUndefined();
  });

  it("editing one cell's live style afterwards leaves the others alone", () => {
    const ws = sheet();
    Range.setFont(ws, "A1:B1", { bold: true });
    Cell.getStyle(ws, "A1").font!.italic = true;
    expect(Cell.getFont(ws, "B1")).toEqual({ bold: true });
  });

  it("refuses a malformed facet before creating any cell", () => {
    const ws = sheet();
    expect(() => Range.setFont(ws, "A1:C3", { size: "big" } as never)).toThrow(
      expect.objectContaining({ field: "value.size" })
    );
    expect(() => Range.setStyle(ws, "A1:C3", 5 as never)).toThrow(ExcelArgumentError);
    expect(Row.getCount(ws)).toBe(0);
  });
});
