import { FUTURE_FUNCTION_PREFIXES, toFileFormula } from "@excel/utils/future-functions";
import { describe, expect, it } from "vitest";

describe("toFileFormula", () => {
  it.each([
    ["=1+1", "1+1"],
    ["=A1", "A1"],
    ["==1", "=1"],
    ["=SUM(A1:A3)", "SUM(A1:A3)"]
  ])("drops one leading = from %s, which a file never stores", (text, stored) => {
    expect(toFileFormula(text)).toBe(stored);
  });

  it.each([
    ["FILTER(B2:B5,B2:B5>1)", "_xlfn._xlws.FILTER(B2:B5,B2:B5>1)"],
    ["SORT(A1:A3)", "_xlfn._xlws.SORT(A1:A3)"],
    ["XLOOKUP(2,A1:A3,B1:B3)", "_xlfn.XLOOKUP(2,A1:A3,B1:B3)"],
    ["SUM(SEQUENCE(3))", "SUM(_xlfn.SEQUENCE(3))"],
    ["=vstack(1, 2)", "_xlfn.VSTACK(1, 2)"],
    ["CONCAT (A1,B1)", "_xlfn.CONCAT (A1,B1)"],
    ["STDEV.S(A1:A3)+STDEV(A1:A3)", "_xlfn.STDEV.S(A1:A3)+STDEV(A1:A3)"]
  ])("prefixes %s", (text, stored) => {
    expect(toFileFormula(text)).toBe(stored);
  });

  it.each([
    "_xlfn._xlws.FILTER(B2:B5,B2:B5>1)",
    "_xlfn.XLOOKUP(2,A1:A3,B1:B3)",
    "SUM(A1:A3)",
    'IF(A1="FILTER(x)",1,2)',
    "'SEQUENCE(1)'!A1+Sheet1!A1",
    "Tbl[FILTER]",
    "SEQUENCE",
    "MyUdf(1)",
    "_xlfn.LET(_xlpm.x,2,_xlpm.x*3)"
  ])("leaves %s alone", text => {
    expect(toFileFormula(text)).toBe(text);
  });

  it("adds the worksheet prefix a stored _xlfn.SORT is missing", () => {
    expect(toFileFormula("_xlfn.SORT(A1:A3)")).toBe("_xlfn._xlws.SORT(A1:A3)");
  });

  it.each([
    ["LET(x,2,x*3)", "_xlfn.LET(_xlpm.x,2,_xlpm.x*3)"],
    ["LET(x, 2, y, x+1, x*y)", "_xlfn.LET(_xlpm.x, 2, _xlpm.y, _xlpm.x+1, _xlpm.x*_xlpm.y)"],
    ["LET(x,B2:B5,SUM(x))", "_xlfn.LET(_xlpm.x,B2:B5,SUM(_xlpm.x))"],
    ["MAP(B2:B5,LAMBDA(v,v*2))", "_xlfn.MAP(B2:B5,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))"],
    [
      "LET(f,LAMBDA(a,b,a+b),f(1,2))",
      "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b),_xlpm.f(1,2))"
    ],
    // A name is in scope only after its definition, and only inside the call.
    ["LET(x,x,x)+x", "_xlfn.LET(_xlpm.x,x,_xlpm.x)+x"],
    // A sheet-qualified name or a string is not a parameter.
    ['LET(x,1,Sheet1!x+"x")', '_xlfn.LET(_xlpm.x,1,Sheet1!x+"x")']
  ])("names LET and LAMBDA parameters in %s", (text, stored) => {
    expect(toFileFormula(text)).toBe(stored);
  });

  // MS-XLSX lists `_xlfn._xlws.PY` and `_xlfn.PQSOURCE`, but Excel refuses the whole workbook when a cell spells
  // either with its prefix, and opens it (showing `#NAME?`) when the call is bare — so they are written bare.
  it.each(["PY(0,0)", "PQSOURCE(1)", "SUM(PY(0,0),1)"])("leaves %s unprefixed", text => {
    expect(toFileFormula(text)).toBe(text);
  });

  // `@` is how Excel's formula bar shows a this-row reference; a file spells it `[#This Row]`, and Excel refuses a
  // workbook whose cell holds `Tb[@[x]:[y]]`.
  it.each([
    ["Tb[@x]*2", "Tb[[#This Row],[x]]*2"],
    ["Tb[@[x y]]", "Tb[[#This Row],[x y]]"],
    ["SUM(Tb[@x]:Tb[@y])", "SUM(Tb[[#This Row],[x]]:Tb[[#This Row],[y]])"],
    ["SUM(Tb[@[x]:[y]])", "SUM(Tb[[#This Row],[x]:[y]])"],
    ["Tb[ @x ]", "Tb[[#This Row],[x]]"],
    ["Tb[@]", "Tb[#This Row]"],
    ["Tb[@[a']b]]", "Tb[[#This Row],[a']b]]"],
    ["Tb[[#This Row],[x]]", "Tb[[#This Row],[x]]"],
    ['IF(A1="Tb[@x]",1,2)', 'IF(A1="Tb[@x]",1,2)']
  ])("writes the this-row reference in %s as a file spells it", (text, stored) => {
    expect(toFileFormula(text)).toBe(stored);
  });

  // Without a table name the reference is left as written: no Excel-written file shows the unnamed file form.
  it.each(["[@x]+1", "CONCAT([@A])"])("leaves the unnamed this-row reference in %s alone", text => {
    expect(toFileFormula(text)).toBe(text.startsWith("CONCAT") ? `_xlfn.${text}` : text);
  });

  it("spells every prefix the way the specification lists it", () => {
    expect(FUTURE_FUNCTION_PREFIXES.get("FILTER")).toBe("_xlfn._xlws.");
    expect(FUTURE_FUNCTION_PREFIXES.get("XLOOKUP")).toBe("_xlfn.");
    expect(FUTURE_FUNCTION_PREFIXES.has("SUM")).toBe(false);
  });
});
