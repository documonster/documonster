/**
 * OFFSET returns a reference (Microsoft: "Returns a reference to a range that is a specified
 * number of rows and columns from a cell or range of cells"), so a function that reads its
 * argument as a reference sees where it points. It used to return the cells' values.
 */
import { check } from "@formula/__tests__/reference/sheet-check.helpers";
import { describe, it } from "vitest";

describe("OFFSET is a reference", () => {
  it("to reference-aware functions", () => {
    // S!A1:A3 = 5, 3, 9; S!B1 = "x"
    check([
      ["ROW(OFFSET(S!A1,2,0))", 3],
      ["COLUMN(OFFSET(S!A1,0,2))", 3],
      ["ISREF(OFFSET(S!A1,0,0))", true],
      ['CELL("address",OFFSET(A1,1,1))', "$B$2"],
      ["AREAS(OFFSET(S!A1,0,0))", 1],
      ["OFFSET(OFFSET(S!A1,1,0),1,0)", 9],
      ["ROWS(OFFSET(S!A1,0,0,3))", 3]
    ]);
  });

  it("and its cells' values to everything else", () => {
    check([
      ["OFFSET(S!A1,1,0)", 3],
      ["SUM(OFFSET(S!A1,0,0,3))", 17],
      ["INDEX(OFFSET(S!A1,0,0,3),2)", 3],
      ["SUM(OFFSET(S!A1:A2,1,0))", 12],
      ["OFFSET(S!A1,0,0,1,1)+1", 6],
      ["OFFSET(S!A1,-1,0)", { error: "#REF!" }]
    ]);
  });
});
