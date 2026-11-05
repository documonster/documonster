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
