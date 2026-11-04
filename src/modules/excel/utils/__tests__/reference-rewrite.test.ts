import type { LineEdit, ReferenceEdit } from "@excel/utils/reference-rewrite";
import {
  editAnchorCorner,
  editSpan,
  moveAnchorCorners,
  rewriteArea,
  rewriteFormula,
  rewriteSqref,
  sheetPrefix
} from "@excel/utils/reference-rewrite";
import { describe, expect, it } from "vitest";

const insertRows = (at: number, n: number, sheet = "S"): LineEdit => ({
  kind: "rows",
  sheet,
  at,
  deleted: 0,
  inserted: n
});
const deleteRows = (at: number, n: number, sheet = "S"): LineEdit => ({
  kind: "rows",
  sheet,
  at,
  deleted: n,
  inserted: 0
});
const insertCols = (at: number, n: number, sheet = "S"): LineEdit => ({
  kind: "columns",
  sheet,
  at,
  deleted: 0,
  inserted: n
});
const deleteCols = (at: number, n: number, sheet = "S"): LineEdit => ({
  kind: "columns",
  sheet,
  at,
  deleted: n,
  inserted: 0
});

describe("editSpan", () => {
  it.each([
    // [lo, hi, edit, expected]
    [2, 5, insertRows(2, 1), [3, 6]], // at the first line: moves
    [2, 5, insertRows(3, 2), [2, 7]], // strictly inside: grows
    [2, 5, insertRows(5, 1), [2, 6]], // at the last line: grows
    [2, 5, insertRows(6, 1), [2, 5]], // just past: unchanged
    [2, 5, deleteRows(1, 1), [1, 4]],
    [2, 5, deleteRows(3, 2), [2, 3]],
    [2, 5, deleteRows(1, 3), [1, 2]], // top deleted: shrinks from the top
    [2, 5, deleteRows(4, 5), [2, 3]], // bottom deleted
    [2, 5, deleteRows(2, 4), undefined], // exactly deleted
    [2, 5, deleteRows(1, 10), undefined],
    [2, 5, deleteRows(6, 3), [2, 5]],
    [1, 1_048_576, insertRows(1, 3), [1, 1_048_576]], // whole axis
    [5, 1_048_576, insertRows(1, 1), [6, 1_048_576]], // pushed past the end: stops there
    [1_048_576, 1_048_576, insertRows(2, 1), undefined]
  ] as const)("%i..%i", (lo, hi, edit, expected) => {
    expect(editSpan(lo, hi, edit)).toEqual(expected);
  });

  it("splices as a deletion followed by an insertion", () => {
    expect(editSpan(1, 4, { kind: "rows", sheet: "S", at: 2, deleted: 2, inserted: 1 })).toEqual([
      1, 3
    ]);
    expect(
      editSpan(2, 3, { kind: "rows", sheet: "S", at: 2, deleted: 2, inserted: 1 })
    ).toBeUndefined();
  });
});

describe("rewriteArea / rewriteSqref", () => {
  it("moves, grows and shrinks areas on the edited axis only", () => {
    expect(rewriteArea("A3", insertRows(1, 2))).toBe("A5");
    expect(rewriteArea("$B$2:$C$4", insertRows(3, 1))).toBe("$B$2:$C$5");
    expect(rewriteArea("A:A", insertRows(1, 1))).toBe("A:A");
    expect(rewriteArea("3:5", insertRows(1, 1))).toBe("4:6");
    expect(rewriteArea("B:D", insertCols(1, 1))).toBe("C:E");
    expect(rewriteArea("3:5", insertCols(1, 1))).toBe("3:5");
    expect(rewriteArea("A3", deleteRows(3, 1))).toBeUndefined();
  });

  it("drops the areas of a list the edit deletes", () => {
    expect(rewriteSqref("A1:A2 C5 D7:D9", deleteRows(5, 3))).toBe("A1:A2 D5:D6");
    expect(rewriteSqref("C5", deleteRows(5, 1))).toBeUndefined();
  });
});

describe("rewriteFormula — row and column edits", () => {
  it.each([
    // same sheet, unqualified, relative and absolute alike
    ["SUM(A4:A5)", insertRows(1, 1), "SUM(A5:A6)"],
    ["=$A$4*A4", insertRows(4, 2), "=$A$6*A6"],
    ["SUM(A2:A5)", insertRows(3, 1), "SUM(A2:A6)"],
    ["SUM(A2:A5)", insertRows(6, 1), "SUM(A2:A5)"],
    ["A3+1", deleteRows(3, 1), "#REF!+1"],
    ["SUM(A2:A5)", deleteRows(2, 2), "SUM(A2:A3)"],
    ["SUM(A:A)+SUM(3:4)", insertRows(1, 1), "SUM(A:A)+SUM(4:5)"],
    ["SUM(B:C)", insertCols(2, 1), "SUM(C:D)"],
    ["SUM(B:C)", insertCols(3, 1), "SUM(B:D)"],
    ["B1*C1", deleteCols(2, 1), "#REF!*B1"],
    ["$A1+XFD1", insertCols(1, 1), "$B1+#REF!"],
    // lower case is read, and written back in Excel's case only where it changed
    ["a1+a9", insertRows(5, 1), "a1+A10"]
  ] as const)("%s through %o", (formula, edit, expected) => {
    expect(rewriteFormula(formula, edit, "S")).toBe(expected);
  });

  it("follows qualified references to the edited sheet from any sheet", () => {
    expect(rewriteFormula("S!A4*2+A4", insertRows(1, 1), "Other")).toBe("S!A5*2+A4");
    expect(rewriteFormula("'My S'!$A$4", insertRows(1, 1, "My S"), "Other")).toBe("'My S'!$A$5");
    expect(rewriteFormula("s!A4", insertRows(1, 1), "Other")).toBe("s!A5");
    expect(rewriteFormula("S!A4+1", deleteRows(4, 1), "Other")).toBe("#REF!+1");
    // a qualified reference to another sheet, written on the edited one, does not move
    expect(rewriteFormula("Other!A4+A4", insertRows(1, 1), "S")).toBe("Other!A4+A5");
  });

  it("leaves what only looks like a reference alone", () => {
    const edit = insertRows(1, 1);
    expect(rewriteFormula('LOG10(A1)&"A1"&Table1[A1]&Name1', edit, "S")).toBe(
      'LOG10(A2)&"A1"&Table1[A1]&Name1'
    );
    expect(rewriteFormula("[1]S!A1+Table1[[#This Row],[Col]]", edit, "S")).toBe(
      "[1]S!A1+Table1[[#This Row],[Col]]"
    );
    expect(rewriteFormula("SUM(S1:S3!A1)", insertRows(1, 1, "S1"), "S2")).toBe("SUM(S1:S3!A1)");
    expect(rewriteFormula('"unterminated', edit, "S")).toBe('"unterminated');
  });

  it("leaves unqualified references alone where they belong to no sheet", () => {
    expect(rewriteFormula("A4+S!A4", insertRows(1, 1), undefined)).toBe("A4+S!A5");
  });
});

describe("rewriteFormula — sheet rename and removal", () => {
  const rename = (from: string, to: string): ReferenceEdit => ({ kind: "rename", from, to });

  it("renames every prefix naming the sheet, quoting as the new name needs", () => {
    expect(rewriteFormula("A!A1*2+B!A1", rename("A", "Data"), "C")).toBe("Data!A1*2+B!A1");
    expect(rewriteFormula("a!A1", rename("A", "My Data"), "C")).toBe("'My Data'!A1");
    expect(rewriteFormula("'My Data'!A1", rename("My Data", "Plain"), "C")).toBe("Plain!A1");
    expect(rewriteFormula("A!A1", rename("A", "It's"), "C")).toBe("'It''s'!A1");
    expect(rewriteFormula("'It''s'!A1", rename("It's", "B"), "C")).toBe("B!A1");
    expect(rewriteFormula("A!A1", rename("A", "B2"))).toBe("'B2'!A1");
    expect(rewriteFormula("SUM(A:C!B1)", rename("C", "Z Z"))).toBe("SUM('A:Z Z'!B1)");
    expect(rewriteFormula("SUM(Jan:Mar!B1)", rename("Mar", "D"))).toBe("SUM('Jan:D'!B1)");
    expect(rewriteFormula("A!LocalName", rename("A", "B"))).toBe("B!LocalName");
    expect(rewriteFormula('"A!A1"&[1]A!A1', rename("A", "B"))).toBe('"A!A1"&[1]A!A1');
  });

  it("turns references to a removed sheet into #REF!", () => {
    const sheets = ["North", "South", "East", "West"];
    const remove = (sheet: string): ReferenceEdit => ({ kind: "removeSheet", sheet, sheets });
    expect(rewriteFormula("South!A1*2+East!A1", remove("South"), "North")).toBe("#REF!*2+East!A1");
    expect(rewriteFormula("SUM(South!A1:A3)", remove("South"), "North")).toBe("SUM(#REF!)");
    expect(rewriteFormula("South!Local+1", remove("South"), "North")).toBe("#REF!+1");
    expect(rewriteFormula("SUM(South:West!A1)", remove("South"), "North")).toBe(
      "SUM(East:West!A1)"
    );
    expect(rewriteFormula("SUM(South:West!A1)", remove("West"), "North")).toBe(
      "SUM(South:East!A1)"
    );
    expect(rewriteFormula("SUM(South:East!A1)", remove("East"), "North")).toBe("SUM(South!A1)");
    expect(rewriteFormula("SUM(South:West!A1)", remove("East"), "North")).toBe(
      "SUM(South:West!A1)"
    );
    expect(rewriteFormula("SUM('South:South'!A1)", remove("South"), "North")).toBe("SUM(#REF!)");
  });
});

describe("sheetPrefix", () => {
  it.each([
    ["Sheet1", "Sheet1!"],
    ["My Sheet", "'My Sheet'!"],
    ["A1", "'A1'!"],
    ["R1C1", "'R1C1'!"],
    ["TRUE", "'TRUE'!"],
    ["2024", "'2024'!"],
    ["销售", "'销售'!"]
  ])("%s", (name, expected) => {
    expect(sheetPrefix(name)).toBe(expected);
  });
});

describe("drawing corners", () => {
  it("moves a corner with the line it lies on", () => {
    // index 2 = row 3
    expect(editAnchorCorner({ index: 2, offset: 50 }, insertRows(3, 2), false)).toEqual({
      index: 4,
      offset: 50
    });
    expect(editAnchorCorner({ index: 2, offset: 50 }, insertRows(4, 2), false)).toEqual({
      index: 2,
      offset: 50
    });
    expect(editAnchorCorner({ index: 2, offset: 50 }, deleteRows(3, 1), false)).toEqual({
      index: 2,
      offset: 0
    });
    expect(editAnchorCorner({ index: 6, offset: 50 }, deleteRows(3, 2), false)).toEqual({
      index: 4,
      offset: 50
    });
  });

  it("reads a bottom-right corner at offset 0 as the edge of the line before it", () => {
    // `A1:C3` ends at the top of row 4: inserting at row 4 leaves it, inserting at row 3 moves it
    expect(editAnchorCorner({ index: 3, offset: 0 }, insertRows(4, 1), true)).toEqual({
      index: 3,
      offset: 0
    });
    expect(editAnchorCorner({ index: 3, offset: 0 }, insertRows(3, 1), true)).toEqual({
      index: 4,
      offset: 0
    });
  });

  it("follows the placement", () => {
    const tl = { index: 2, offset: 0 };
    const br = { index: 6, offset: 0 };
    const edit = insertRows(4, 1);
    expect(moveAnchorCorners(tl, br, "twoCell", edit)).toEqual({ tl, br: { index: 7, offset: 0 } });
    expect(moveAnchorCorners(tl, br, "oneCell", edit)).toEqual({ tl, br });
    expect(moveAnchorCorners(tl, br, "oneCell", insertRows(1, 1))).toEqual({
      tl: { index: 3, offset: 0 },
      br: { index: 7, offset: 0 }
    });
    expect(moveAnchorCorners(tl, br, "absolute", insertRows(1, 1))).toEqual({ tl, br });
  });
});

describe("structured references to removed tables", () => {
  it("become #REF!, whatever the case, and nothing else does", () => {
    const edit: LineEdit = { ...deleteRows(1, 5), removedTables: new Set(["sales"]) };
    expect(rewriteFormula('SUM(Sales[Q1])+SUM(Kept[x])&"Sales[Q1]"', edit, "Other")).toBe(
      'SUM(#REF!)+SUM(Kept[x])&"Sales[Q1]"'
    );
    expect(rewriteFormula("SALES[[#Totals],[Q1]]", edit)).toBe("#REF!");
  });
});
