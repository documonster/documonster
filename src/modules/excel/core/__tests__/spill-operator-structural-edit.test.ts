/**
 * A spill-range reference `A1#` follows its anchor through a structural edit, as Excel's does: inserting a row above
 * it makes it `A2#`, and deleting the anchor makes the reference `#REF!` — not `#REF!#`, which is no formula.
 */
import { Cell, Column, Row, Workbook, Worksheet } from "@excel";
import { describe, expect, it } from "vitest";

function book() {
  const wb = Workbook.create();
  const ws = Workbook.addWorksheet(wb, "S");
  const other = Workbook.addWorksheet(wb, "My Sheet");
  Row.addMany(ws, [[1], [2], [3]]);
  return { wb, ws, other };
}

describe("structural edits move `A1#`", () => {
  it("a row inserted above the anchor", () => {
    const { ws } = book();
    Cell.setValue(ws, "C1", { formula: "SUM(A1#)+ROWS($A$1#)" });
    Row.insert(ws, 1, []);
    expect(Cell.getFormula(ws, "C2")).toBe("SUM(A2#)+ROWS($A$2#)");
  });

  it("a column inserted before the anchor, on another sheet", () => {
    const { ws, other } = book();
    Cell.setValue(other, "B1", { formula: "SUM(S!A1#)" });
    Cell.setValue(other, "B2", { formula: "SUM('My Sheet'!A5#)" });
    Column.splice(ws, 1, 0, []);
    Column.splice(other, 1, 0, []);
    expect(Cell.getFormula(other, "C1")).toBe("SUM(S!B1#)");
    expect(Cell.getFormula(other, "C2")).toBe("SUM('My Sheet'!B5#)");
  });

  it("deleting the anchor makes the reference #REF!", () => {
    const { ws, other } = book();
    Cell.setValue(ws, "C5", { formula: "SUM(A1#)" });
    Cell.setValue(other, "A1", { formula: "IFERROR(S!$A$1#,0)" });
    Row.splice(ws, 1, 1);
    expect(Cell.getFormula(ws, "C4")).toBe("SUM(#REF!)");
    // Qualified, the error is written alone, as `reference-rewrite.ts` writes every deleted qualified reference.
    expect(Cell.getFormula(other, "A1")).toBe("IFERROR(#REF!,0)");
  });

  it("renaming the sheet renames the qualifier", () => {
    const { ws, other } = book();
    Cell.setValue(other, "A1", { formula: "S!A1#" });
    Worksheet.setName(ws, "Data Sheet");
    expect(Cell.getFormula(other, "A1")).toBe("'Data Sheet'!A1#");
  });
});
