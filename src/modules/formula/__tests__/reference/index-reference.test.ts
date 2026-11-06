import { Cell, Workbook } from "@excel/index";
/**
 * INDEX of a reference is a reference (Microsoft: "Reference form … Returns the reference of the
 * cell at the intersection of a particular row and column"); of an array it is values. Whole
 * columns keep their position. Values Excel gave for the same cases are pinned in
 * `excel-ground-truth.test.ts`; these cover the edges around them.
 */
import { check } from "@formula/__tests__/reference/sheet-check.helpers";
import { describe, it } from "vitest";

// S!A1:A3 = 1, 2, =A2*10 (20); S!B1:B3 = 10, 20, 30
function seed(wb: ReturnType<typeof Workbook.create>): void {
  const s = Workbook.addWorksheet(wb, "S");
  Cell.setValue(s, "A1", 1);
  Cell.setValue(s, "A2", 2);
  Cell.setValue(s, "A3", { formula: "A2*10", result: 0 });
  [10, 20, 30].forEach((v, i) => Cell.setValue(s, `B${i + 1}`, v));
}

describe("INDEX reference form", () => {
  it("selects a cell, a row, a column or the whole reference", () => {
    check(
      [
        ["INDEX(S!A1:B3,3,1)", 20],
        ["SUM(S!A1:INDEX(S!A1:A3,3))", 23],
        ["ROWS(INDEX(S!A1:B3,0,0))", 3],
        ["COLUMNS(INDEX(S!A1:B3,2,0))", 2],
        ["ROWS(INDEX(S!A:A,0,1))", 1048576],
        ["ROW(INDEX(S!A:B,3,2))", 3],
        ["INDEX(S!A1:B3,4,1)", { error: "#REF!" }],
        ["INDEX(S!A1:B3,1,3)", { error: "#REF!" }],
        ["INDEX(S!A1:B3,-1,1)", { error: "#VALUE!" }],
        ["INDEX((S!A1:A2,S!B1:B3),3,1,2)", 30],
        ["INDEX(S!A1:B3,1,1,2)", { error: "#REF!" }],
        ["ISREF(INDEX({1,2;3,4},2,2))", false],
        ["INDEX({1,2;3,4},2,2)", 4]
      ],
      seed
    );
  });
});

describe("criteria ranges must still be the same size", () => {
  it("a whole column beside a shorter range is #VALUE!", () => {
    check(
      [
        ['COUNTIFS(S!A:A,"",S!B1:B5,"")', { error: "#VALUE!" }],
        ['SUMIFS(S!B:B,S!A1:A3,">0")', { error: "#VALUE!" }],
        ['COUNTIFS(S!A:A,">0",S!B:B,">0")', 3]
      ],
      seed
    );
  });
});
