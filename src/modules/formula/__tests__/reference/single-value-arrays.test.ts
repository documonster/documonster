/**
 * A single value where a function wants an array is a 1×1 array, as in Excel: a one-cell
 * reference (which reaches a function as the cell's value) or a literal. These functions
 * answered `#VALUE!` (or `#N/A`) for one. A single error value is propagated as the error.
 * Expected values are Excel's, per Microsoft's documentation of each function.
 */
import { check } from "@formula/__tests__/reference/sheet-check.helpers";
import { describe, it } from "vitest";

const NA = { error: "#N/A" };

describe("dynamic-array functions take a single value as a 1×1 array", () => {
  it("SORT, UNIQUE, FILTER, TAKE, DROP, CHOOSEROWS, CHOOSECOLS", () => {
    check([
      ["SORT(S!A1)", 5],
      ["SORT(5)", 5],
      ["UNIQUE(S!A1)", 5],
      ["FILTER(S!A1,TRUE)", 5],
      ["FILTER(S!A1,FALSE)", { error: "#CALC!" }],
      ['FILTER(S!A1,FALSE,"none")', "none"],
      ["TAKE(S!A1,1)", 5],
      ["DROP(S!A1,0)", 5],
      ["DROP(S!A1,1)", { error: "#CALC!" }],
      ["CHOOSEROWS(S!A1,1)", 5],
      ["CHOOSECOLS(S!A1,-1)", 5],
      ["CHOOSEROWS(S!A1,2)", { error: "#VALUE!" }]
    ]);
  });

  it("a single error value is that error", () => {
    check([
      ["SORT(NA())", NA],
      ["FILTER(S!A1,NA())", NA],
      ["EXPAND(1/0,2,2,0)", { error: "#DIV/0!" }],
      ["MATCH(5,NA(),0)", NA]
    ]);
  });

  it("WRAPROWS / WRAPCOLS: one value is a vector; a 2-D array is #VALUE!, wrap_count < 1 #NUM!", () => {
    check([
      ["WRAPROWS(S!A1,1)", 5],
      ["WRAPROWS(S!A1,2)", [[5, NA]]],
      ["WRAPCOLS(S!A1,1)", 5],
      ["WRAPCOLS(S!A1,2)", [[5], [NA]]],
      ["WRAPROWS(S!A1,0)", { error: "#NUM!" }],
      ["WRAPCOLS(S!A1:A3,0)", { error: "#NUM!" }],
      ["WRAPROWS({1,2;3,4},2)", { error: "#VALUE!" }],
      ["WRAPCOLS({1,2;3,4},2)", { error: "#VALUE!" }],
      [
        "WRAPROWS(S!A1:A3,2)",
        [
          [5, 3],
          [9, NA]
        ]
      ]
    ]);
  });

  it("EXPAND", () => {
    check([
      [
        "EXPAND(S!A1,2,2,0)",
        [
          [5, 0],
          [0, 0]
        ]
      ],
      ["EXPAND(S!A1,1,2)", [[5, NA]]]
    ]);
  });

  it("SORTBY: a 1×1 key sorts a 1×1 array; a key of the wrong shape is #VALUE!; a row key sorts columns", () => {
    check([
      ["SORTBY(S!A1,S!A1)", 5],
      ["SORTBY(S!A1:A3,S!A1:A3)", [[3], [5], [9]]],
      ["SORTBY(S!A1:A3,S!A1:A3,-1)", [[9], [5], [3]]],
      ["SORTBY(S!A1:A3,S!A1)", { error: "#VALUE!" }],
      ["SORTBY(S!A1:A3,{1,2,3})", { error: "#VALUE!" }],
      ["SORTBY({3,1,2},{3,1,2})", [[1, 2, 3]]],
      [
        "SORTBY({3,1,2;30,10,20},{3,1,2})",
        [
          [1, 2, 3],
          [10, 20, 30]
        ]
      ],
      ["SORTBY(S!A1:A3,S!A1:A3,1,{1,2,3})", { error: "#VALUE!" }]
    ]);
  });
});

describe("INDEX", () => {
  it("indexes a one-row array along its only dimension", () => {
    check([
      ["INDEX({1,2},2)", 2],
      ["INDEX({1;2},2)", 2],
      ["INDEX({1,2},1,2)", 2],
      ["INDEX(S!A1:C1,2)", "x"],
      ["INDEX(S!A1:A3,2)", 3],
      ["INDEX({1,2},3)", { error: "#REF!" }],
      ["INDEX({1,2;3,4},2,1)", 3]
    ]);
  });

  it("a single value is a 1×1 array", () => {
    check([
      ["INDEX(S!A1,1)", 5],
      ["INDEX(S!A1,1,1)", 5],
      ["INDEX(S!A1,2)", { error: "#REF!" }]
    ]);
  });
});

describe("lookups take a single-cell lookup or table", () => {
  it("MATCH, XMATCH, XLOOKUP, VLOOKUP, HLOOKUP, LOOKUP", () => {
    check([
      ["MATCH(5,S!A1,0)", 1],
      ["MATCH(4,S!A1,0)", NA],
      ["XMATCH(5,S!A1)", 1],
      ["XLOOKUP(5,S!A1,S!C1)", 7],
      ["XLOOKUP(4,S!A1,S!C1)", NA],
      ["VLOOKUP(5,S!A1,1,FALSE)", 5],
      ["HLOOKUP(5,S!A1,1,FALSE)", 5],
      ["LOOKUP(5,S!A1)", 5],
      ["LOOKUP(5,S!A1,S!C1)", 7]
    ]);
  });
});

describe("criteria functions take a single-cell range", () => {
  it("COUNTIF, SUMIF, AVERAGEIF and the -IFS forms", () => {
    check([
      ["COUNTIF(S!A1,5)", 1],
      ["COUNTIF(S!A1,4)", 0],
      ["SUMIF(S!A1,5)", 5],
      ["SUMIF(S!A1,5,S!C1)", 7],
      ["AVERAGEIF(S!A1,5,S!C1)", 7],
      ["COUNTIFS(S!A1,5)", 1],
      ["COUNTIFS(S!A1,5,S!C1,7)", 1],
      ["SUMIFS(S!C1,S!A1,5)", 7],
      ["AVERAGEIFS(S!C1,S!A1,5)", 7],
      ["MAXIFS(S!C1,S!A1,5)", 7],
      ["MINIFS(S!C1,S!A1,5)", 7]
    ]);
  });
});

describe("matrix, paired-array and cash-flow functions take a single value", () => {
  it("MMULT, MINVERSE, MDETERM, SUMX2MY2 / SUMX2PY2 / SUMXMY2, FREQUENCY", () => {
    check([
      ["MMULT(S!A1,S!C1)", 35],
      ["MINVERSE(S!A1)", 0.2],
      ["MDETERM(S!A1)", 5],
      ["SUMX2MY2(S!A1,S!C1)", -24],
      ["SUMX2PY2(S!A1,S!C1)", 74],
      ["SUMXMY2(S!A1,S!C1)", 4],
      ["FREQUENCY(S!A1,S!C1)", [[1], [0]]]
    ]);
  });

  it("IRR / MIRR / XIRR need a positive and a negative flow: a single value is that error, not #VALUE!", () => {
    check([
      ["IRR(S!A1)", { error: "#NUM!" }],
      ["MIRR(S!A1,0.1,0.1)", { error: "#DIV/0!" }],
      ["MIRR({-1,-2},0.1,0.1)", { error: "#DIV/0!" }],
      ["XIRR(S!A1,S!C1)", { error: "#NUM!" }]
    ]);
  });
});
