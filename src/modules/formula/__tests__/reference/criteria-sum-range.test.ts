import { Cell, Workbook } from "@excel/index";
/**
 * `SUMIF` / `AVERAGEIF` read their sum (average) range as Excel does: from its top-left cell,
 * in the shape of the criteria range — `SUMIF(A1:A3,">0",B1)` sums B1:B3 (Microsoft:
 * "The actual cells that are added are determined by using the upper leftmost cell in the
 * sum_range argument as the beginning cell, and then including cells that correspond in size
 * and shape to the range argument"). The -IFS functions do not stretch: every range must be the
 * size and shape of the first, otherwise #VALUE!.
 */
import { check, type Expected } from "@formula/__tests__/reference/sheet-check.helpers";
import { describe, it } from "vitest";

// S!A1:A3 = 1, -2, 3; S!B1:B3 = 10, 20, =C1+1 (101); S!C1 = 100; T!B1:B3 = 7, 8, 9
function seed(wb: ReturnType<typeof Workbook.create>): void {
  const s = Workbook.addWorksheet(wb, "S");
  [1, -2, 3].forEach((v, i) => Cell.setValue(s, `A${i + 1}`, v));
  Cell.setValue(s, "B1", 10);
  Cell.setValue(s, "B2", 20);
  Cell.setValue(s, "B3", { formula: "C1+1", result: 0 });
  Cell.setValue(s, "C1", 100);
  const t = Workbook.addWorksheet(wb, "T");
  [7, 8, 9].forEach((v, i) => Cell.setValue(t, `B${i + 1}`, v));
}

const run = (cases: ReadonlyArray<readonly [string, Expected]>) => check(cases, seed);
const VALUE = { error: "#VALUE!" };

describe("SUMIF / AVERAGEIF stretch the sum range from its top-left cell", () => {
  it("a one-cell, a shorter and a same-size sum range give the same sum", () => {
    run([
      ['SUMIF(S!A1:A3,">0",S!B1:B3)', 111],
      ['SUMIF(S!A1:A3,">0",S!B1)', 111],
      ['SUMIF(S!A1:A3,">0",S!B1:B2)', 111],
      ['SUMIF(S!A1:A3,">0",S!B1:D9)', 111],
      ['AVERAGEIF(S!A1:A3,">0",S!B1)', 55.5],
      ['AVERAGEIF(S!A1:A3,">0",S!B1:B2)', 55.5]
    ]);
  });

  it("from wherever the sum range starts, on any sheet", () => {
    run([
      ['SUMIF(S!A1:A3,">0",S!B2)', 20],
      ['SUMIF(S!A1:A3,">0",T!B1)', 16],
      ['SUMIF(S!A1:A2,">0",S!A2:B2)', -2]
    ]);
  });

  it("an array sum range must be the criteria range's size", () => {
    run([
      ['SUMIF(S!A1:A3,">0",{10;20;30})', 40],
      ['SUMIF(S!A1:A3,">0",{10;20})', VALUE],
      ['AVERAGEIF(S!A1:A3,">0",{10,20,30})', VALUE]
    ]);
  });

  it("a range stretched past the sheet's last row is cut there (Excel: excel-ground-truth)", () => {
    run([
      ['SUMIF(S!A1:A3,">0",S!B1048576)', 0],
      ['SUMIF(S!A:A,">0",S!B2)', 20]
    ]);
  });

  it("whole columns and rows pair each cell with its partner", () => {
    run([
      ['SUMIF(S!A:A,">0",S!B1)', 111],
      ['SUMIF(S!A:A,">0",S!B:B)', 111],
      ['AVERAGEIF(S!A:A,">0",T!B1)', 8],
      ['SUMIF(S!1:1,">5",S!2:2)', 20],
      ['SUMIF(S!1:1,">5",S!A2)', 20]
    ]);
  });

  it("a sum range produced at run time is stretched too; a union is #VALUE!", () => {
    run([
      ['SUMIF(S!A1:A3,">0",OFFSET(S!B1,0,0))', 111],
      ['SUMIF(S!A1:A3,">0",INDIRECT("S!B1"))', 111],
      ['SUMIF(S!A1:A3,">0",(S!B1,S!B2))', VALUE],
      ['SUMIF((S!A1,S!A2),">0",S!B1)', VALUE]
    ]);
  });

  it("the criteria range is not stretched: a missing sum range sums it", () => {
    run([
      ['SUMIF(S!A1:A3,">0")', 4],
      ['AVERAGEIF(S!A1:A3,"<0")', -2]
    ]);
  });
});

describe("the -IFS functions require same-size ranges", () => {
  it("#VALUE! for a range of another size", () => {
    run([
      ['SUMIFS(S!B1,S!A1:A3,">0")', VALUE],
      ['AVERAGEIFS(S!B1:B2,S!A1:A3,">0")', VALUE],
      ['MAXIFS(S!B1,S!A1:A3,">0")', VALUE],
      ['MINIFS(S!B1,S!A1:A3,">0")', VALUE],
      ['COUNTIFS(S!A1:A3,">0",S!B1:B2,">0")', VALUE],
      ['SUMIFS(S!B1:B3,S!A1:A3,">0")', 111]
    ]);
  });
});
