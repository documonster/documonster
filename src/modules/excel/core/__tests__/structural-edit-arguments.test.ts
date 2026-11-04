/**
 * Row and column splices: a deletion that reaches past the last row deletes what it names, the counts are checked
 * before anything moves, and a splice through the handle of a removed sheet is refused rather than rewriting the
 * references of whichever sheet now carries its name.
 */
import {
  Cell,
  Column,
  DefinedNames,
  ExcelArgumentError,
  ExcelError,
  Row,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

function sheetWithRows(n: number) {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  Row.addMany(
    ws,
    Array.from({ length: n }, (_, i) => [String.fromCharCode(97 + i)])
  );
  return { wb, ws };
}

function firstColumn(ws: ReturnType<typeof sheetWithRows>["ws"], n: number) {
  return Array.from({ length: n }, (_, i) => Row.getValues(ws, i + 1)[0] ?? null);
}

describe("Row.splice deleting through the last row", () => {
  it.each([
    [4, 2, ["a", "b", "c", null, null]],
    [5, 1, ["a", "b", "c", "d", null]],
    [3, 3, ["a", "b", null, null, null]],
    [1, 5, [null, null, null, null, null]],
    [4, 9, ["a", "b", "c", null, null]]
  ])("splice(%i, %i) on five rows", (start, count, expected) => {
    const { ws } = sheetWithRows(5);
    Row.splice(ws, start, count);
    expect(firstColumn(ws, 5)).toEqual(expected);
    expect(Row.getCount(ws)).toBe(expected.filter(v => v !== null).length);
  });

  it("writes the inserted rows and clears the rest of the deleted block", () => {
    const { ws } = sheetWithRows(5);
    Row.splice(ws, 4, 3, ["x"]);
    expect(firstColumn(ws, 5)).toEqual(["a", "b", "c", "x", null]);
  });

  it("moves the cells the same way the references move", async () => {
    const { wb, ws } = sheetWithRows(5);
    Cell.setValue(ws, "C1", { formula: "SUM(A4:A5)" });
    Row.splice(ws, 4, 2);
    expect(Cell.getFormula(ws, "C1")).toBe("SUM(#REF!)");
    const back = Workbook.getWorksheet(await Workbook.read(await Workbook.toBuffer(wb)), "S")!;
    expect([Cell.getValue(back, "A4"), Cell.getValue(back, "A5")]).toEqual([null, null]);
  });
});

describe("Row.getLast", () => {
  it("is the last row that exists after the last row is deleted", () => {
    const { ws } = sheetWithRows(5);
    Row.splice(ws, 5, 1);
    expect(Row.getCount(ws)).toBe(4);
    expect(Row.getLast(ws)).toBe(4);
  });
});

describe("splice and duplicate counts are checked before anything moves", () => {
  it.each<[string, (ws: ReturnType<typeof sheetWithRows>["ws"]) => unknown]>([
    ["Row.splice negative count", ws => Row.splice(ws, 3, -1)],
    ["Row.splice fractional count", ws => Row.splice(ws, 3, 1.5)],
    ["Row.splice NaN start", ws => Row.splice(ws, NaN, 1)],
    ["Row.splice start 0", ws => Row.splice(ws, 0, 1)],
    ["Row.splice a bare number to insert", ws => Row.splice(ws, 2, 0, 42 as never)],
    ["Row.duplicate NaN count", ws => Row.duplicate(ws, 1, NaN)],
    ["Row.duplicate negative count", ws => Row.duplicate(ws, 1, -2)],
    ["Row.duplicate fractional count", ws => Row.duplicate(ws, 1, 1.5)],
    ["Column.splice negative count", ws => Column.splice(ws, 1, -1)],
    ["Column.splice NaN start", ws => Column.splice(ws, NaN, 1)],
    ["Column.splice a non-array column", ws => Column.splice(ws, 1, 0, 5 as never)],
    ["Row.autoFit row 0", ws => Row.autoFit(ws, 0)],
    ["Row.autoFit NaN start", ws => Row.autoFit(ws, { start: NaN })],
    ["Row.add copyStyle 'below'", ws => Row.add(ws, [1], { copyStyle: "below" as never })],
    ["Row.insert copyStyle 'sideways'", ws => Row.insert(ws, 1, [1], { copyStyle: "x" as never })]
  ])("%s", (_name, edit) => {
    const { ws } = sheetWithRows(4);
    Cell.setValue(ws, "C1", { formula: "A4" });
    // An argument error, or the position error a row or column outside the grid gives.
    expect(() => edit(ws)).toThrow(ExcelError);
    expect(firstColumn(ws, 5)).toEqual(["a", "b", "c", "d", null]);
    expect(Cell.getFormula(ws, "C1")).toBe("A4");
  });
});

describe("a splice through a removed sheet's handle", () => {
  it("is refused, and leaves the references of the sheet that now has its name alone", () => {
    const wb = Workbook.create();
    const old = Workbook.addWorksheet(wb, "Data");
    const calc = Workbook.addWorksheet(wb, "Calc");
    Workbook.removeWorksheet(wb, old);
    const fresh = Workbook.addWorksheet(wb, "Data");
    Row.addMany(fresh, [["a"], ["b"], ["c"]]);
    Cell.setValue(calc, "A1", { formula: "Data!A3" });
    DefinedNames.add(Workbook.getDefinedNames(wb), "Data!$A$1", "First");

    expect(() => Row.splice(old, 1, 2)).toThrow(ExcelArgumentError);
    expect(() => Column.splice(old, 1, 1)).toThrow(ExcelArgumentError);
    expect(() => Cell.removeAllNames(old, "A1")).toThrow(ExcelArgumentError);
    expect(() => Cell.removeName(old, "A1", "First")).toThrow(ExcelArgumentError);
    expect(() => Column.setDefinitions(old, [{ width: 3 }])).toThrow(ExcelArgumentError);
    expect(() => Worksheet.unmerge(old, "C1:D1")).toThrow(ExcelArgumentError);

    expect(Cell.getFormula(calc, "A1")).toBe("Data!A3");
    expect(Cell.getNames(fresh, "A1")).toEqual(["First"]);
  });
});

describe("Column.splice deleting through the last column", () => {
  it("drops the emptied cells, so the counts agree with the written file", async () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.add(ws, ["a", "b", "c"]);
    Column.splice(ws, 2, 2);
    expect(Row.getValues(ws, 1)).toEqual(["a"]);
    expect([Column.getCount(ws), Column.getLast(ws), Row.getCellCount(ws, 1)]).toEqual([1, 1, 1]);
    expect([Cell.has(ws, "B1"), Cell.has(ws, "C1")]).toEqual([false, false]);
    const back = Workbook.getWorksheet(await Workbook.read(await Workbook.toBuffer(wb)), "S")!;
    expect(Column.getCount(back)).toBe(1);
  });

  it("keeps a row's cells when the deletion starts past its end", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.add(ws, ["a", "b", "c"]);
    Row.add(ws, ["x", "y", "z", "w", "v"]);
    Column.splice(ws, 5, 1);
    expect(Row.getValues(ws, 1)).toEqual(["a", "b", "c"]);
    expect(Row.getValues(ws, 2)).toEqual(["x", "y", "z", "w"]);
    expect(Column.getCount(ws)).toBe(4);
  });
});

describe("Column.splice inserts", () => {
  it("accepts read-only columns of values, as Row.splice and Column.setValues do", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Row.add(ws, ["a", "b"]);
    const columns: readonly (readonly string[])[] = Object.freeze([Object.freeze(["x"])]);
    Column.splice(ws, 2, 0, ...columns);
    expect(Row.getValues(ws, 1)).toEqual(["a", "x", "b"]);
  });
});
