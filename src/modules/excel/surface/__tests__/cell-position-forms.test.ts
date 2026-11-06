/**
 * The `(ws, row, col)` form takes a row number, and the date writers check what they are given before the cell is
 * created.
 */
import {
  Cell,
  ExcelArgumentError,
  ExcelInvalidAddressError,
  Row,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  return Workbook.addWorksheet(wb, "S");
}

describe('an "A1" address with a column after it', () => {
  it.each<[string, (ws: ReturnType<typeof sheet>) => unknown]>([
    ["Cell.setValue", ws => (Cell.setValue as (...a: unknown[]) => void)(ws, "A1", 5, "x")],
    ["Cell.getValue", ws => (Cell.getValue as (...a: unknown[]) => unknown)(ws, "A1", 5)],
    [
      "Cell.setFont",
      ws => (Cell.setFont as (...a: unknown[]) => void)(ws, "A1", 5, { bold: true })
    ],
    ["Cell.getFont", ws => (Cell.getFont as (...a: unknown[]) => unknown)(ws, "A1", 5)],
    ["Cell.has", ws => (Cell.has as (...a: unknown[]) => unknown)(ws, "A1", 5)],
    ["Cell.getTemporal", ws => (Cell.getTemporal as (...a: unknown[]) => unknown)(ws, "A1", 5)],
    [
      "Cell.setDateParts",
      ws => (Cell.setDateParts as (...a: unknown[]) => void)(ws, "A1", 5, { year: 2024 })
    ],
    ["Cell.removeName", ws => (Cell.removeName as (...a: unknown[]) => unknown)(ws, "A1", 5, "n")]
  ])("is refused by %s rather than addressing another cell", (_name, call) => {
    const ws = sheet();
    expect(() => call(ws)).toThrow(ExcelInvalidAddressError);
    expect(Worksheet.getDimensions(ws)).toBeUndefined();
    expect(Row.getCount(ws)).toBe(0);
  });

  it("the (row, col) form still works", () => {
    const ws = sheet();
    Cell.setValue(ws, 1, 5, "x");
    expect(Cell.getValue(ws, "E1")).toBe("x");
    Cell.setDateParts(ws, 2, 5, { year: 2024, month: 1, day: 15 });
    expect(Cell.getDateParts(ws, "E2")).toMatchObject({ year: 2024, month: 1, day: 15 });
  });
});

describe("a { row, col } position", () => {
  it("addresses the same cell as its A1 text, in readers and writers", () => {
    const ws = sheet();
    Cell.setValue(ws, { row: 2, col: 3 }, "x");
    Cell.setFont(ws, { row: 2, col: 3 }, { bold: true });
    Cell.setDateParts(ws, { row: 3, col: 3 }, { year: 2024, month: 1, day: 15 });
    expect(Cell.getValue(ws, "C2")).toBe("x");
    expect(Cell.getFont(ws, { row: 2, col: 3 })).toEqual({ bold: true });
    expect(Cell.has(ws, { row: 2, col: 3 })).toBe(true);
    expect(Cell.getDateParts(ws, { row: 3, col: 3 })).toMatchObject({ year: 2024 });
    expect(Cell.getFullAddress(ws, { row: 2, col: 3 }).address).toBe("C2");
  });

  it("accepts what Cell.getFullAddress returns", () => {
    const ws = sheet();
    Cell.setValue(ws, Cell.getFullAddress(ws, "D4"), 1);
    expect(Cell.getValue(ws, "D4")).toBe(1);
  });

  it.each<[string, unknown, string]>([
    ["row 0", { row: 0, col: 1 }, "addr.row"],
    ["a fractional row", { row: 1.5, col: 1 }, "addr.row"],
    ["a row past the grid", { row: 1048577, col: 1 }, "addr.row"],
    ["a missing col", { row: 1 }, "addr.col"],
    ["a col past XFD", { row: 1, col: 16385 }, "addr.col"],
    ["a string col", { row: 1, col: "A" }, "addr.col"]
  ])("refuses %s, naming the field and the call", (_name, addr, field) => {
    const ws = sheet();
    for (const call of [
      () => Cell.setValue(ws, addr as never, 1),
      () => Cell.getValue(ws, addr as never)
    ]) {
      expect(call).toThrow(ExcelInvalidAddressError);
      expect(call).toThrow(expect.objectContaining({ field }));
      expect(call).toThrow(/^Cell\.(setValue|getValue): /);
    }
    expect(Row.getCount(ws)).toBe(0);
  });
});

describe("the deprecated (row, col) form", () => {
  // Expected values are what v0.21.1 answered for the same calls.
  it("keeps addressing (row, col) in readers and writers", () => {
    const ws = sheet();
    Cell.setValue(ws, 2, 3, "x");
    Cell.setFont(ws, 2, 3, { bold: true });
    Cell.setStyle(ws, 2, 3, { numFmt: "0.0" });
    Cell.setNote(ws, 2, 3, "n");
    expect(Cell.getValue(ws, 2, 3)).toBe("x");
    expect(Cell.getFont(ws, 2, 3)).toEqual({ bold: true });
    expect(Cell.getNumFmt(ws, 2, 3)).toBe("0.0");
    expect(Cell.getNote(ws, 2, 3)).toBe("n");
    expect(Cell.has(ws, 2, 3)).toBe(true);
    expect(Cell.getFullAddress(ws, 2, 3).address).toBe("C2");
    Cell.setFont(ws, 2, 3, undefined);
    expect(Cell.getFont(ws, "C2")).toBeUndefined();
  });
});

describe("Cell.setDateParts", () => {
  it.each([
    ["NaN year", { year: NaN }, undefined],
    ["null parts", null, undefined],
    ["month 13", { year: 2024, month: 13 }, undefined],
    ["an unknown kind", { year: 2024, month: 1, day: 15 }, "bogus"]
  ])("refuses %s and creates no cell", (_name, parts, kind) => {
    const ws = sheet();
    expect(() => Cell.setDateParts(ws, "J10", parts as never, kind as never)).toThrow();
    expect(Cell.has(ws, "J10")).toBe(false);
    expect(Row.getCount(ws)).toBe(0);
  });

  it("refuses an unknown kind as an argument error", () => {
    const ws = sheet();
    expect(() => Cell.setDateParts(ws, "A1", { year: 2024 }, "bogus" as never)).toThrow(
      ExcelArgumentError
    );
  });
});

describe("Cell.applyStyle", () => {
  it("refuses a name that is not a defined style as an argument error", () => {
    const ws = sheet();
    expect(() => Cell.applyStyle(ws, "A1", "Nope")).toThrow(ExcelArgumentError);
    expect(() => Cell.applyStyle(ws, "A1", 42 as never)).toThrow(ExcelArgumentError);
    expect(Cell.has(ws, "A1")).toBe(false);
  });
});
