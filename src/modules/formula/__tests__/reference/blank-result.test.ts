import { Cell, Workbook } from "@excel/index";
/**
 * A formula whose result is an empty cell holds 0, as Excel stores it — alone, or as an element of
 * a spilled array. That is the cell's value, so another formula reading the cell sees 0 too
 * (`ISBLANK` is FALSE, `=""` is FALSE, `COUNTBLANK` skips it). Inside an expression an empty cell
 * stays empty: `A2=""`, `A2=0` and `ISBLANK(A2)` are TRUE, `LEN(A2)` is 0, `A2&"x"` is "x". An empty
 * string stays an empty string.
 *
 * Every expected value is the one Excel for Mac 16.114 saved for the same formula over the same
 * data (`tmp/office-tests/r4-formula/README.md`, second run).
 */
import { check, type Case } from "@formula/__tests__/reference/sheet-check.helpers";
import { describe, it } from "vitest";

// E!A1 = 1, E!A2 empty, E!A3 = 3; E!B1 = "p", E!B2 empty, E!B3 = "q"; E!C1 = =E!A2; E!C2 = ="";
// E!C3 = =F!A2 (a cell of the first case's spill); E!D1:D3 = 1, 2, 3; E!E3 = =E!A2.
function seed(wb: ReturnType<typeof Workbook.create>): void {
  const e = Workbook.addWorksheet(wb, "E");
  Cell.setValue(e, "A1", 1);
  Cell.setValue(e, "A3", 3);
  Cell.setValue(e, "B1", "p");
  Cell.setValue(e, "B3", "q");
  Cell.setValue(e, "C1", { formula: "E!A2", result: 0 });
  Cell.setValue(e, "C2", { formula: '""', result: 0 });
  Cell.setValue(e, "C3", { formula: "F!A2", result: 0 });
  [1, 2, 3].forEach((v, i) => Cell.setValue(e, `D${i + 1}`, v));
  Cell.setValue(e, "E3", { formula: "E!A2", result: 0 });
}

const run = (cases: ReadonlyArray<Case>) => check(cases, seed);

describe("a formula's empty result is 0", () => {
  it("in a spilled array, each empty element", () => {
    // Case 0 spills into F!A1:A3 and case 2 into F!A21:A23; later cases read them.
    run([
      ["E!A1:A3", [[1], [0], [3]], "dynamic"],
      ["FILTER(E!A1:A3,{TRUE;TRUE;TRUE})", [[1], [0], [3]], "dynamic"],
      ["E!B1:B3", [["p"], [0], ["q"]], "dynamic"],
      ["INDEX(E!A1:A3,0,1)", [[1], [0], [3]], "dynamic"],
      ["IF({1;1;1},E!A1:A3)", [[1], [0], [3]], "dynamic"],
      ["ISBLANK(F!A2)", false],
      ['F!A2=""', false],
      ["F!A2", 0],
      ["COUNTBLANK(F!A1:A3)", 0],
      ["COUNTIF(F!A1:A3,0)", 1],
      ["ISTEXT(F!A22)", false],
      ["F!A22", 0],
      ["SUM(F!A1#)", 4],
      ["COUNTBLANK(F!A1#)", 0],
      ["ISBLANK(E!C3)", false],
      ["E!C3", 0]
    ]);
  });

  it("as a single result, and as the value other formulas read from that cell", () => {
    run([
      ["E!A2", 0],
      ["E!C1", 0],
      ['E!C1=""', false],
      ["ISBLANK(E!C1)", false],
      ["E!C1=0", true],
      ["IF(TRUE,E!A2)", 0],
      ["INDEX(E!A1:A3,2)", 0],
      ["VLOOKUP(2,E!D1:E3,2,FALSE)", 0],
      ["VLOOKUP(3,E!D1:E3,2,FALSE)", 0]
    ]);
  });
});

describe("inside an expression an empty cell stays empty", () => {
  it("and an empty string stays an empty string", () => {
    run([
      ['E!A2=""', true],
      ["E!A2=0", true],
      ["ISBLANK(E!A2)", true],
      ["LEN(E!A2)", 0],
      ['E!A2&"x"', "x"],
      // Excel saves `t="str"` with an empty `<v/>`: the empty string, not 0.
      ["E!C2", ""],
      ['E!C2=""', true]
    ]);
  });
});
