/**
 * The spill-range operator `A1#` as a file stores it: `_xlfn.ANCHORARRAY(A1)` (the file grammar has no `#` operator;
 * Excel's own XLSB save of `future-functions.xlsx` calls `_xlfn.ANCHORARRAY` through its hidden stub name).
 */
import { fromFileFormula, toFileFormula } from "@excel/utils/future-functions";
import { describe, expect, it } from "vitest";

const PAIRS: [string, string][] = [
  ["A1#", "_xlfn.ANCHORARRAY(A1)"],
  ["$B$2#", "_xlfn.ANCHORARRAY($B$2)"],
  ["Sheet2!A1#", "_xlfn.ANCHORARRAY(Sheet2!A1)"],
  ["'My Sheet'!$A$1#", "_xlfn.ANCHORARRAY('My Sheet'!$A$1)"],
  ["SUM(A1#)*2", "SUM(_xlfn.ANCHORARRAY(A1))*2"],
  ["-A1#+B1#", "-_xlfn.ANCHORARRAY(A1)+_xlfn.ANCHORARRAY(B1)"],
  ["ROWS(Data!C3#)", "ROWS(_xlfn.ANCHORARRAY(Data!C3))"],
  ["_xlfn._xlws.SORT(A1#)", "_xlfn._xlws.SORT(_xlfn.ANCHORARRAY(A1))"],
  ["IFERROR(A1#,#N/A)", "IFERROR(_xlfn.ANCHORARRAY(A1),#N/A)"]
];

describe("toFileFormula writes `A1#` as `_xlfn.ANCHORARRAY(A1)`", () => {
  it.each(PAIRS)("%s", (text, file) => {
    expect(toFileFormula(text)).toBe(file);
    expect(toFileFormula(`=${text}`)).toBe(file);
    // Already in the file's form: unchanged.
    expect(toFileFormula(file)).toBe(file);
  });

  it("leaves a `#` that is not the operator alone", () => {
    for (const text of [
      'A1&"#"',
      "#REF!+1",
      "SUM(#REF!)",
      "Tb[[#This Row],[x]]",
      "IF(A1=1,#N/A,2)",
      "ISERROR(A1)"
    ]) {
      expect(toFileFormula(text)).toBe(text);
    }
  });
});

describe("fromFileFormula reads `_xlfn.ANCHORARRAY(ref)` back as `ref#`", () => {
  it.each(PAIRS)("%s", (text, file) => {
    expect(fromFileFormula(file)).toBe(text);
  });

  it("keeps the function form for an argument the operator cannot take", () => {
    for (const file of [
      "_xlfn.ANCHORARRAY(Name)",
      "_xlfn.ANCHORARRAY(INDEX(A1:B2,1,1))",
      "_xlfn.ANCHORARRAY(A1:B2)",
      'LEN("_xlfn.ANCHORARRAY(A1)")'
    ]) {
      expect(fromFileFormula(file)).toBe(file);
    }
  });

  it("is the inverse of toFileFormula on every other formula", () => {
    for (const file of ["_xlfn.XLOOKUP(1,A1:A3,B1:B3)", "SUM(A1:A3)"]) {
      expect(fromFileFormula(file)).toBe(file);
    }
    // `_xlfn.SINGLE` is the file's `@` (see implicit-intersection-operator.test.ts).
    expect(fromFileFormula("_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))")).toBe("@A1#");
  });
});
