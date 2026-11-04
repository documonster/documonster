/**
 * The implicit-intersection operator `@` as a file stores it: `_xlfn.SINGLE(x)`. MS-XLSX's formula grammar has `+` and
 * `-` as its only prefix operators, and Excel writes `@` as that call (its XLSB save of `future-functions.xlsx` holds
 * `_xlfn.SINGLE` as a future-function stub). Read back as `@x` wherever `@x` reads as the same formula — the operand
 * is what `@` takes in the parser, nothing more, nothing less — and as the call otherwise.
 */
import { fromFileFormula, toFileFormula } from "@excel/utils/future-functions";
import { describe, expect, it } from "vitest";

const PAIRS: [string, string][] = [
  ["@A1:A3", "_xlfn.SINGLE(A1:A3)"],
  ["@$B$2:$B$9", "_xlfn.SINGLE($B$2:$B$9)"],
  ["@Sheet2!A1:A3", "_xlfn.SINGLE(Sheet2!A1:A3)"],
  ["@'My Sheet'!A:A", "_xlfn.SINGLE('My Sheet'!A:A)"],
  ["@Data", "_xlfn.SINGLE(Data)"],
  ["@Tb[x]", "_xlfn.SINGLE(Tb[x])"],
  ["@SEQUENCE(3)", "_xlfn.SINGLE(_xlfn.SEQUENCE(3))"],
  ["@INDEX(A1:C3,0,1)", "_xlfn.SINGLE(INDEX(A1:C3,0,1))"],
  ["@(A1:A3)", "_xlfn.SINGLE((A1:A3))"],
  ["SUM(@A1:A3)", "SUM(_xlfn.SINGLE(A1:A3))"],
  ["@A1:A3*2", "_xlfn.SINGLE(A1:A3)*2"],
  ["@A1:A3+@B1:B3", "_xlfn.SINGLE(A1:A3)+_xlfn.SINGLE(B1:B3)"],
  // `@` takes what binds at its power: `^` and the reference operators bind tighter.
  ["@A1^2", "_xlfn.SINGLE(A1^2)"],
  ["@A1:INDEX(B1:B3,2)", "_xlfn.SINGLE(A1:INDEX(B1:B3,2))"],
  ["@A1:A3 B1:B3", "_xlfn.SINGLE(A1:A3 B1:B3)"],
  ["@A1%", "_xlfn.SINGLE(A1%)"],
  ["@-A1:A3", "_xlfn.SINGLE(-A1:A3)"],
  ["-@A1:A3", "-_xlfn.SINGLE(A1:A3)"],
  ["@@A1:A3", "_xlfn.SINGLE(_xlfn.SINGLE(A1:A3))"],
  ["@A1#", "_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))"],
  ["IF(@A1:A3>1,@B1:B3,0)", "IF(_xlfn.SINGLE(A1:A3)>1,_xlfn.SINGLE(B1:B3),0)"],
  ["@LAMBDA(x,x)(A1:A3)", "_xlfn.SINGLE(_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(A1:A3))"]
];

describe("toFileFormula writes `@x` as `_xlfn.SINGLE(x)`", () => {
  it.each(PAIRS)("%s", (text, file) => {
    expect(toFileFormula(text)).toBe(file);
    expect(toFileFormula(`=${text}`)).toBe(file);
    expect(toFileFormula(file)).toBe(file);
  });

  it("leaves an `@` that is not the operator alone", () => {
    expect(toFileFormula('A1&"@"')).toBe('A1&"@"');
    expect(toFileFormula("Tb[@x]*2")).toBe("Tb[[#This Row],[x]]*2");
    expect(toFileFormula("SUM(Tb[@[a b]:[c]])")).toBe("SUM(Tb[[#This Row],[a b]:[c]])");
    expect(toFileFormula("[@x]", { table: "Tb" })).toBe("Tb[[#This Row],[x]]");
    expect(toFileFormula("'a@b'!A1")).toBe("'a@b'!A1");
  });
});

describe("fromFileFormula reads `_xlfn.SINGLE(x)` back as `@x`", () => {
  it.each(PAIRS)("%s", (text, file) => {
    // A future function keeps the prefix the file gives it, as every read does: only `@` changes spelling.
    const read = fromFileFormula(file);
    expect(read).toBe(
      /SEQUENCE|LAMBDA/.test(text) ? file.replace(/_xlfn\.SINGLE\(([^]*)\)$/, "@$1") : text
    );
    expect(toFileFormula(read)).toBe(file);
  });

  it("keeps the call where `@x` would read as a different formula", () => {
    for (const file of [
      // What follows would join the operand: `@A1^2` is `@(A1^2)`.
      "_xlfn.SINGLE(A1)^2",
      "_xlfn.SINGLE(A1)%",
      "_xlfn.SINGLE(A1:A3) B1:B3",
      "_xlfn.SINGLE(A1):B2",
      // More than the operand `@` takes.
      "_xlfn.SINGLE(A1+B1)",
      "_xlfn.SINGLE(A1:A3*2)",
      // Not one argument, or not written as a file writes it.
      "_xlfn.SINGLE(A1,B1)",
      "_xlfn.SINGLE()",
      "_xlfn.SINGLE( A1)",
      "_xlfn.SINGLE(A1 )",
      'LEN("_xlfn.SINGLE(A1)")'
    ]) {
      expect(fromFileFormula(file)).toBe(file);
    }
    expect(fromFileFormula("_xlfn.SINGLE(A1+B1)*_xlfn.SINGLE(C1:C3)")).toBe(
      "_xlfn.SINGLE(A1+B1)*@C1:C3"
    );
  });
});
