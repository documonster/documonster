import { Cell, Workbook } from "@excel/index";
/**
 * Ground truth from Excel (Microsoft 365 for Mac, 16.114) for behaviour Microsoft does not
 * document. The library wrote a workbook holding these formulas over this data, Excel
 * recalculated and saved it, and the values below are the ones Excel saved
 * (`tmp/office-tests/r4-formula/README.md` lists them beside the engine's).
 */
import { check, type Expected } from "@formula/__tests__/reference/sheet-check.helpers";
import { describe, it } from "vitest";

function seed(wb: ReturnType<typeof Workbook.create>): void {
  const f = Workbook.addWorksheet(wb, "F");
  const d = Workbook.addWorksheet(wb, "D");
  const w = Workbook.addWorksheet(wb, "W");
  // D!A1:A5 = 1..5, D!B1:B9 = 10..90, D!C1:C5 = 100..500, D!D1 = 1000
  for (let i = 1; i <= 9; i++) {
    if (i <= 5) {
      Cell.setValue(d, `A${i}`, i);
      Cell.setValue(d, `C${i}`, i * 100);
    }
    Cell.setValue(d, `B${i}`, i * 10);
  }
  Cell.setValue(d, "D1", 1000);
  // D!E1:I1 = 1..5, D!E2:I2 = 10..50, D!E3:E6 = 11..14
  ["E", "F", "G", "H", "I"].forEach((col, i) => {
    Cell.setValue(d, `${col}1`, i + 1);
    Cell.setValue(d, `${col}2`, (i + 1) * 10);
  });
  [11, 12, 13, 14].forEach((v, i) => Cell.setValue(d, `E${i + 3}`, v));
  // D!A1048570:A1048576 = 1; D!B1048570 = 3, D!B1048575 = 5, D!B1048576 = 7
  for (let r = 1048570; r <= 1048576; r++) {
    Cell.setValue(d, `A${r}`, 1);
  }
  Cell.setValue(d, "B1048570", 3);
  Cell.setValue(d, "B1048575", 5);
  Cell.setValue(d, "B1048576", 7);
  // D!XFA1:XFD1 = 1; D!XFC2 = 4, D!XFD2 = 9; D!XFC3 = 66, D!XFD3 = 77
  ["XFA", "XFB", "XFC", "XFD"].forEach(col => Cell.setValue(d, `${col}1`, 1));
  Cell.setValue(d, "XFC2", 4);
  Cell.setValue(d, "XFD2", 9);
  Cell.setValue(d, "XFC3", 66);
  Cell.setValue(d, "XFD3", 77);
  // W!A1 = "x", W!A3 = 1, W!A5 = "y"; W!B1:B4 = 10, 6, 20, 5; W!C2 = -1, W!C9 = -7
  Cell.setValue(w, "A1", "x");
  Cell.setValue(w, "A3", 1);
  Cell.setValue(w, "A5", "y");
  [10, 6, 20, 5].forEach((v, i) => Cell.setValue(w, `B${i + 1}`, v));
  Cell.setValue(w, "C2", -1);
  Cell.setValue(w, "C9", -7);
  // S!A3 = "a", S!A4 = "b" (data not starting in row 1); T!A1:A8 = 1..8
  const s = Workbook.addWorksheet(wb, "S");
  Cell.setValue(s, "A3", "a");
  Cell.setValue(s, "A4", "b");
  const t = Workbook.addWorksheet(wb, "T");
  for (let i = 1; i <= 8; i++) {
    Cell.setValue(t, `A${i}`, i);
  }
  // F!H1:J3 = 1..9, so CELL's address carries no sheet name
  for (let i = 0; i < 9; i++) {
    Cell.setValue(f, `${"HIJ"[i % 3]}${Math.floor(i / 3) + 1}`, i + 1);
  }
}

const run = (cases: ReadonlyArray<readonly [string, Expected]>) => check(cases, seed);

describe("LOOKUP stretches its result vector from its first cell", () => {
  it("down a column, or across a row (a single cell goes across); 2-D or a short array is #N/A", () => {
    run([
      ["LOOKUP(3,D!A1:A5,D!B1)", 1000],
      ["LOOKUP(3,D!A1:A5,D!B1:B2)", 30],
      ["LOOKUP(5,D!A1:A5,D!B1:B2)", 50],
      ["LOOKUP(3,D!A1:A5,D!B1:B9)", 30],
      ["LOOKUP(3,D!A1:A5,D!B1:D1)", 1000],
      ["LOOKUP(3,D!A1:A5,D!B1:C2)", { error: "#N/A" }],
      ["LOOKUP(3,D!E1:I1,D!E2)", 30],
      ["LOOKUP(3,D!E1:I1,D!E2:F2)", 30],
      ["LOOKUP(3,D!E1:I1,D!E2:E6)", 12],
      ["LOOKUP(3,D!A1:A5,{10;20})", { error: "#N/A" }],
      ["LOOKUP(3,D!A1:A5,{10,20,30,40,50})", 30],
      ["LOOKUP(2,D!A1:A5,D!XFC3)", 77]
    ]);
  });

  it("past the sheet's last column or row it reads an empty cell, which a result shows as 0", () => {
    run([
      ["LOOKUP(3,D!A1:A5,D!XFC3)", 0],
      ["LOOKUP(3,D!A1:A5,D!B1048575:B1048576)", 0]
    ]);
  });
});

describe("SUMIF / AVERAGEIF stretch the sum range", () => {
  it("and clip it at the sheet's last row or column", () => {
    run([
      ['SUMIF(D!A1048570:A1048576,">0",D!B1048576)', 7],
      ['SUMIF(D!A1048570:A1048576,">0",D!B1048575)', 12],
      ['SUMIF(D!A1048570:A1048576,">0",D!B1048570)', 15],
      ['AVERAGEIF(D!A1048570:A1048576,">0",D!B1048576)', 7],
      ['SUMIF(D!XFA1:XFD1,">0",D!XFD2)', 9],
      ['SUMIF(D!XFA1:XFD1,">0",D!XFC2)', 13],
      ['SUMIF(W!A:A,"<>",W!B2)', 11],
      ['SUMIF(W!1:1,"<>",W!B2)', 5]
    ]);
  });

  it("from its first cell whatever its shape", () => {
    run([
      ['SUMIF(D!A1:A3,">0",D!B1:C1)', 60],
      ['SUMIF(D!A1:A3,">0",D!C1:C2)', 600],
      ['SUMIF(D!E1:G1,">0",D!E2:E4)', 60],
      ['AVERAGEIF(D!A1:A3,">1",D!B1:D1)', 25]
    ]);
  });
});

describe("whole columns and rows count the empty cells below and beside the data", () => {
  it("for blank-matching criteria in COUNTIF, COUNTIFS and COUNTBLANK", () => {
    run([
      ['COUNTIF(W!A:A,"")', 1048573],
      ['COUNTIF(W!A:A,"=")', 1048573],
      ['COUNTIF(W!A:A,"<>")', 3],
      ['COUNTIF(W!A:A,"<>x")', 1048575],
      ['COUNTIF(W!A:A,"<>1")', 1048575],
      ["COUNTBLANK(W!A:A)", 1048573],
      ['COUNTIFS(W!A:A,"",W!B:B,"")', 1048571],
      ['COUNTIFS(W!A:A,"<>x",W!B:B,"")', 1048572],
      ['SUMIF(W!A:A,"",W!B:B)', 11],
      ['AVERAGEIF(W!A:A,"",W!B:B)', 5.5],
      ['MINIFS(W!B:B,W!A:A,"")', 5],
      ['MAXIFS(W!B:B,W!A:A,"<>x")', 20],
      ['MINIFS(W!C:C,W!A:A,"")', -7],
      ['COUNTIF(W!1:1,"")', 16382],
      ["COUNTBLANK(W!1:1)", 16382],
      ['COUNTIF(W!A:B,"")', 2097145],
      ['SUMIFS(W!B:B,W!A:A,"")', 11],
      ['AVERAGEIFS(W!B:B,W!A:A,"<>x")', 10.333333333333334]
    ]);
  });
});

describe("whole columns and rows keep their position", () => {
  it("when the data does not start in row 1, and across sheets with different used ranges", () => {
    run([
      ['MATCH("b",S!A:A,0)', 4],
      ["INDEX(S!A:A,4)", "b"],
      ["ROWS(S!A:A)", 1048576],
      ["COLUMNS(S!1:1)", 16384],
      ['XMATCH("b",S!A:A)', 4],
      ['XLOOKUP("b",S!A:A,T!A:A)', 4],
      ['VLOOKUP("b",S!A:B,1,FALSE)', "b"],
      ["OFFSET(S!A:A,3,0,1,1)", "b"],
      ['COUNTIFS(S!A:A,"",T!A:A,"")', 1048568],
      ['SUMIFS(T!A:A,S!A:A,"a")', 3],
      ['SUMIF(S!A:A,"",T!A:A)', 29],
      ['SUMIF(S!A:A,"",T!A1)', 29],
      ['COUNTIF(S!A:A,"<>a")', 1048575],
      ["COUNTBLANK(S!A1:A10)", 8],
      ['COUNTIF(W!A1:A1048576,"")', 1048573],
      ["COUNTBLANK(W!A:B)", 2097145]
    ]);
  });
});

describe("INDEX of a reference is a reference", () => {
  it("to ROW, COLUMN, ISREF, CELL, AREAS, OFFSET, SUMIF and the range operator", () => {
    run([
      ["ROW(INDEX(D!A1:C3,2,2))", 2],
      ["COLUMN(INDEX(D!A1:C3,2,2))", 2],
      ["SUM(INDEX(D!A1:C3,0,1))", 6],
      ["SUM(INDEX(D!A1:C3,2,0))", 222],
      ["ROWS(INDEX(D!A1:C3,0,2))", 3],
      ["SUM(D!A1:INDEX(D!A:A,5))", 15],
      ["ISREF(INDEX(D!A1:C3,1,1))", true],
      ["ISREF(INDEX({1,2},1))", false],
      ['CELL("address",INDEX(H1:J3,2,2))', "$I$2"],
      ['SUMIF(D!A1:A3,">0",INDEX(D!B:B,1))', 60],
      ["AREAS(INDEX((D!A1:B2,D!C1:C3),1,1,2))", 1],
      ["SUM(INDEX((D!A1:B2,D!C1:C3),0,0,2))", 600],
      ["OFFSET(INDEX(D!A1:C3,2,2),1,0)", 30],
      ["INDEX(D!A1:C3,2,2)", 20],
      ["INDEX(D!A1:A5,4)", 4],
      ["INDEX(D!A1:E1,3)", 100],
      ["ROW(INDEX(S!A:A,4))", 4],
      ["SUM(INDEX(T!A:A,0,1))", 36]
    ]);
  });
});
