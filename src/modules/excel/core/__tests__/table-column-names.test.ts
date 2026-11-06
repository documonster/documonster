/**
 * A table's column names are unique, compared without case, and the header cells say them. Excel refuses a table part
 * whose `tableColumn` names repeat — the workbook opens behind a repair prompt — and gives a repeated header the next
 * free number when one is typed (`a`, `a` → `a`, `a2`). A table read from a file is renamed the same way; a caller's
 * repeated names are refused, since the formulas they write name the columns they chose.
 */
import { Cell, Table, ExcelTableError, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

describe("table column names", () => {
  it("are refused by Table.add when they repeat, naming the name Excel would give", () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    expect(() =>
      Table.add(sheet, {
        name: "D",
        ref: "A1",
        columns: [{ name: "a" }, { name: "A" }, { name: "a2" }, { name: "b" }],
        rows: [[1, 2, 3, 4]]
      })
    ).toThrow(ExcelTableError);
    expect(() =>
      Table.add(sheet, {
        name: "D",
        ref: "A1",
        columns: [{ name: "a" }, { name: "A" }, { name: "a2" }, { name: "b" }],
        rows: [[1, 2, 3, 4]]
      })
    ).toThrow(/"A3"/);
    // Nothing was written, and the name is still free.
    expect(Cell.getValue(sheet, "A1")).toBeNull();
    expect(Table.list(sheet)).toEqual([]);
  });

  it("are refused by addColumn and setColumnName, leaving the table as it was", () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    const t = Table.add(sheet, {
      name: "D",
      ref: "A1",
      columns: [{ name: "a" }, { name: "b" }],
      rows: [[1, 2]]
    });
    expect(() => Table.addColumn(t, { name: "A" }, [3])).toThrow(/"A2"/);
    expect(() => Table.setColumnName(Table.getColumn(t, 1)!, "A")).toThrow(ExcelTableError);
    // Renaming a column to its own name in another case is not a collision.
    Table.setColumnName(Table.getColumn(t, 0)!, "A");
    expect(Table.getModel(t).columns.map(column => column.name)).toEqual(["A", "b"]);
    expect(["A1", "B1", "C1"].map(address => Cell.getValue(sheet, address))).toEqual([
      "A",
      "b",
      null
    ]);
  });

  it("are made unique the way Excel makes them when the model is not the caller's", async () => {
    // A model read from a file reaches the same validation as the deprecated `Table.create`, which keeps renaming.
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Table.create(sheet, {
      name: "D",
      ref: "A1",
      columns: [{ name: "a" }, { name: "A" }, { name: "a2" }, { name: "b" }],
      rows: [[1, 2, 3, 4]]
    });
    expect(["A1", "B1", "C1", "D1"].map(address => Cell.getValue(sheet, address))).toEqual([
      "a",
      "A3",
      "a2",
      "b"
    ]);
  });

  // Each pair below is what Excel itself did when the second header was typed under the first (Excel 16, macOS):
  // renumbered means the two names are the same name to Excel, kept means they are not.
  const SAME: [string, string][] = [
    ["ß", "SS"],
    ["ẞ", "ß"],
    ["ß", "sS"],
    ["ﬁ", "FI"],
    ["ς", "Σ"],
    ["ς", "σ"],
    ["ａ", "A"],
    ["ｶﾞ", "ガ"],
    ["ア", "あ"],
    ["ǅ", "ǆ"],
    ["ſ", "S"],
    ["\u212A", "k"],
    ["µ", "Μ"],
    ["ŉ", "ʼN"],
    ["ǰ", "J\u030C"],
    ["e\u0301", "é"],
    ["x²", "x2"],
    ["①", "1"],
    ["Ⅻ", "XII"],
    ["İ", "i\u0307"],
    ["ᾳ", "ᾼ"],
    ["𐐀", "𐐨"],
    ["ꭰ", "Ꭰ"]
  ];
  const DIFFERENT: [string, string][] = [
    ["ı", "I"],
    ["ı", "i"],
    ["i", "İ"],
    ["é", "e"],
    ["Ä", "A"],
    ["æ", "AE"],
    ["ゝ", "ヽ"],
    ["ﾟ", "゜"],
    ["ᾳ", "ΑΙ"]
  ];

  function headers(names: string[]): unknown[] {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Table.create(sheet, {
      name: "T",
      ref: "A1",
      columns: names.map(name => ({ name })),
      rows: [names.map(() => 1)]
    });
    return names.map((_, i) => Cell.getValue(sheet, `${String.fromCharCode(65 + i)}1`));
  }

  it.each(SAME)("compare %s and %s as Excel does: the same name", (a, b) => {
    expect(headers([a, b])).toEqual([a, `${b}2`]);
  });

  it.each(DIFFERENT)("compare %s and %s as Excel does: different names", (a, b) => {
    expect(headers([a, b])).toEqual([a, b]);
  });

  function add(names: string[]): void {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Table.add(sheet, {
      name: "T",
      ref: "A1",
      columns: names.map(name => ({ name })),
      rows: [names.map(() => 1)]
    });
  }

  it.each(SAME)("Table.add refuses %s beside %s: the same name to Excel", (a, b) => {
    expect(() => add([a, b])).toThrow(ExcelTableError);
  });

  it.each(DIFFERENT)("Table.add accepts %s beside %s: different names to Excel", (a, b) => {
    expect(() => add([a, b])).not.toThrow();
  });

  it("number a third collision past the names already taken", () => {
    expect(headers(["ß", "ẞ", "SS"])).toEqual(["ß", "ẞ2", "SS3"]);
  });

  it("make table names unique by the same comparison", () => {
    const workbook = Workbook.create();
    const sheet = Workbook.addWorksheet(workbook, "S");
    Table.add(sheet, { name: "ß", ref: "A1", columns: [{ name: "a" }], rows: [[1]] });
    expect(() =>
      Table.add(sheet, { name: "SS", ref: "C1", columns: [{ name: "a" }], rows: [[1]] })
    ).toThrow(/already exists/);
    Table.add(sheet, { name: "ı", ref: "E1", columns: [{ name: "a" }], rows: [[1]] });
    expect(() =>
      Table.add(sheet, { name: "I", ref: "G1", columns: [{ name: "a" }], rows: [[1]] })
    ).not.toThrow();
  });
});
