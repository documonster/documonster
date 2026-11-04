/**
 * Row and column setters refuse what a workbook cannot hold before anything changes, and a column key resolves only
 * through the sheet's own keys.
 */
import { Cell, Column, ExcelArgumentError, ExcelError, Row, Workbook } from "@excel/index";
import { describe, expect, it } from "vitest";

function sheet() {
  const wb = Workbook.create();
  return { wb, ws: Workbook.addWorksheet(wb, "S") };
}

describe("column keys", () => {
  it("resolve only through the sheet's own keys, never through Object.prototype", () => {
    const { ws } = sheet();
    expect(() => Column.setWidth(ws, "toString", 5)).toThrow();
    expect(() => Column.setHidden(ws, "constructor", true)).toThrow();
    expect((Object.prototype.toString as { width?: number }).width).toBeUndefined();
    expect(Column.getDefinitions(ws)).toEqual([]);
  });

  it("may be any string, including one that names a member of Object.prototype", () => {
    const { ws } = sheet();
    Column.setKey(ws, 2, "toString");
    Column.setKey(ws, 3, "__proto__");
    Column.setWidth(ws, "toString", 12);
    Column.setWidth(ws, "__proto__", 14);
    expect([Column.getWidth(ws, "B"), Column.getWidth(ws, "C")]).toEqual([12, 14]);
  });

  it("win over a column letter of the same text", () => {
    const { ws } = sheet();
    Column.setKey(ws, 1, "F");
    Column.setWidth(ws, "F", 30);
    expect(Column.getWidth(ws, 1)).toBe(30);
    expect(Column.getNumber(ws, "F")).toBe(1);
  });

  it("refuses a key that is not a non-empty string", () => {
    const { ws } = sheet();
    expect(() => Column.setKey(ws, "D", 42 as never)).toThrow(ExcelArgumentError);
    expect(() => Column.setKey(ws, "D", "")).toThrow(ExcelArgumentError);
    expect(Column.getKey(ws, "D")).toBeUndefined();
  });

  it("refuses a key another column already has", () => {
    const { ws } = sheet();
    Column.setKey(ws, 2, "k");
    expect(() => Column.setKey(ws, 3, "k")).toThrow(ExcelArgumentError);
    expect(Column.getNumber(ws, "k")).toBe(2);
    Column.setKey(ws, 2, "k");
    expect(Column.getNumber(ws, "k")).toBe(2);
  });

  it("refuses duplicate keys in a set of definitions, before changing any column", () => {
    const { ws } = sheet();
    Column.setWidth(ws, "A", 10);
    expect(() => Column.setDefinitions(ws, [{ key: "a" }, { key: "a" }])).toThrow(
      ExcelArgumentError
    );
    expect(Column.getWidth(ws, "A")).toBe(10);
  });
});

describe("Row and Column setters check what they are given", () => {
  it.each<[string, (ws: ReturnType<typeof sheet>["ws"]) => unknown]>([
    ["Row.setHidden with a string", ws => Row.setHidden(ws, 2, "yes" as never)],
    ["Row.setHidden with undefined", ws => Row.setHidden(ws, 2, undefined as never)],
    ["Column.setHidden with a string", ws => Column.setHidden(ws, "B", "no" as never)],
    ["Row.setHeight past Excel's 409 points", ws => Row.setHeight(ws, 2, 410)],
    ["Column.setWidth past Excel's 255 characters", ws => Column.setWidth(ws, "B", 256)],
    ["Column.setDefinitions null", ws => Column.setDefinitions(ws, null as never)],
    ["Column.setDefinitions a negative width", ws => Column.setDefinitions(ws, [{ width: -5 }])],
    ["Column.setDefinitions NaN width", ws => Column.setDefinitions(ws, [{ width: NaN }])],
    [
      "Column.setDefinitions an outline level of 9",
      ws => Column.setDefinitions(ws, [{ outlineLevel: 9 }])
    ],
    [
      "Column.setDefinitions hidden 'yes'",
      ws => Column.setDefinitions(ws, [{ hidden: "yes" as never }])
    ],
    [
      "Column.setDefinitions a header of NaN",
      ws => Column.setDefinitions(ws, [{ header: ["ok", NaN] }])
    ],
    [
      "Column.setDefinitions a style that is not one",
      ws => Column.setDefinitions(ws, [{ style: { font: { size: -1 } } }])
    ],
    ["Column.setHeader NaN in a list", ws => Column.setHeader(ws, "B", ["ok", NaN])],
    ["Column.setValues a non-array", ws => Column.setValues(ws, "H", 5 as never)],
    ["Column.setValues base 3", ws => Column.setValues(ws, "H", [9], { base: 3 as never })],
    ["Row.setValues base 7", ws => Row.setValues(ws, 1, ["x"], { base: 7 as never })],
    ["Row.getValues base 2", ws => Row.getValues(ws, 1, { base: 2 as never })],
    ["Column.getValues base 2", ws => Column.getValues(ws, "A", { base: 2 as never })]
  ])("%s is refused and changes nothing", (_name, edit) => {
    const { ws } = sheet();
    Row.setValues(ws, 1, ["a", "b", "c"]);
    Column.setWidth(ws, "A", 10);
    // An argument error, or the value error `Cell.setValue` gives a value no cell can hold.
    expect(() => edit(ws)).toThrow(ExcelError);
    expect(Row.getValues(ws, 1)).toEqual(["a", "b", "c"]);
    expect(Column.getWidth(ws, "A")).toBe(10);
    expect([Row.isHidden(ws, 2), Column.isHidden(ws, "B"), Row.getHeight(ws, 2)]).toEqual([
      false,
      false,
      undefined
    ]);
    expect(Cell.has(ws, "B2")).toBe(false);
  });

  it("accepts Excel's own limits", () => {
    const { ws } = sheet();
    Row.setHeight(ws, 2, 409);
    Column.setWidth(ws, "B", 255);
    expect([Row.getHeight(ws, 2), Column.getWidth(ws, "B")]).toEqual([409, 255]);
  });
});
