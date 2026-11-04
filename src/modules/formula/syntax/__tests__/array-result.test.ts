/**
 * Whether a formula produces an array that dynamic-array Excel spills: the static decision both the engine and the
 * Excel writers make, so that a formula the engine spills is exactly a formula a file stores as a dynamic array.
 */
import { producesArray } from "@formula/syntax/array-result";
import { parseFormula } from "@formula/syntax/parser";
import { describe, expect, it } from "vitest";

const produces = (formula: string): boolean => producesArray(parseFormula(formula));

describe("producesArray", () => {
  it.each([
    // Functions whose result is an array, from the function table's return classes.
    "SEQUENCE(3)",
    "_xlfn.SEQUENCE(3)",
    "FILTER(A1:A5,B1:B5>1)",
    "_xlfn._xlws.SORT(A1:A5)",
    "SORTBY(A1:A5,B1:B5)",
    "UNIQUE(A1:A5)",
    "RANDARRAY(2,2)",
    "TOCOL(A1:B2)",
    "TOROW(A1:B2)",
    "VSTACK(A1:A2,B1:B2)",
    "HSTACK(A1:A2,B1:B2)",
    "TAKE(A1:A5,2)",
    "DROP(A1:A5,2)",
    "CHOOSECOLS(A1:C3,1)",
    "CHOOSEROWS(A1:C3,1)",
    "EXPAND(A1,2,2)",
    "WRAPROWS(A1:A6,2)",
    "WRAPCOLS(A1:A6,2)",
    'TEXTSPLIT("a,b",",")',
    "MAKEARRAY(2,2,LAMBDA(r,c,r*c))",
    "MAP(A1:A3,LAMBDA(x,x*2))",
    "SCAN(0,A1:A3,LAMBDA(a,x,a+x))",
    "BYROW(A1:B3,LAMBDA(r,SUM(r)))",
    "BYCOL(A1:B3,LAMBDA(c,SUM(c)))",
    "GROUPBY(A1:A5,B1:B5,SUM)",
    "MUNIT(3)",
    "MODE.MULT(A1:A9)",
    "TRANSPOSE(A1:C1)",
    "MMULT(A1:B2,A1:B2)",
    "MINVERSE(A1:B2)",
    "FREQUENCY(A1:A9,B1:B3)",
    "LINEST(A1:A9)",
    "TREND(A1:A9)",
    'STOCKHISTORY("MSFT",TODAY())',
    // Decided at evaluation time, so decided for the array: a 1×1 dynamic array is the value itself.
    "XLOOKUP(1,A1:A3,B1:C3)",
    "REDUCE(0,A1:A3,LAMBDA(a,x,a+x))",
    "TRIMRANGE(A1:A9)",
    'REGEXEXTRACT(A1,"\\d+",1)',
    'REGEXEXTRACT(A1,"(\\d)(\\d)",2)',
    'REGEXEXTRACT(A1,"\\d+",B1)',
    // An array constant of more than one element.
    "{1,2,3}",
    "{1;2}",
    // Lifted: an operator or a value parameter applied to an array.
    "SEQUENCE(3)*2",
    "-SEQUENCE(3)",
    "SEQUENCE(3)%",
    "ABS(SEQUENCE(3))",
    "{1,2}+A1:A2",
    "IF(SEQUENCE(3)>1,1,0)",
    "IF(A1,SEQUENCE(3),0)",
    "CHOOSE(1,SEQUENCE(3),0)",
    "INDEX(SEQUENCE(3,3),0,1)",
    "INDEX(SEQUENCE(3),2)",
    // The spill-range operator.
    "A1#",
    "A1#*2",
    "ANCHORARRAY(A1)",
    // LET and LAMBDA follow what they return.
    "LET(x,SEQUENCE(3),x*2)",
    "LET(x,SEQUENCE(3),y,x+1,y)",
    "LAMBDA(x,x*2)(SEQUENCE(3))",
    "LAMBDA(x,SEQUENCE(x))(3)",
    "LAMBDA(x,LAMBDA(y,SEQUENCE(x+y)))(1)(2)",
    "LET(f,LAMBDA(x,SEQUENCE(x)),f(3))",
    "LET(f,LAMBDA(x,x*2),f(SEQUENCE(3)))"
  ])("%s does", formula => {
    expect(produces(formula)).toBe(true);
    expect(produces(`=${formula}`)).toBe(true);
  });

  it.each([
    "1+2",
    "SUM(A1:A5)",
    "SUM(SEQUENCE(3))",
    "ROWS(SEQUENCE(3))",
    "SUMPRODUCT(SEQUENCE(3)*2)",
    "COUNTA({1,2,3})",
    "SUM(A1#)",
    "{1}",
    "@SEQUENCE(3)",
    "_xlfn.SINGLE(SEQUENCE(3))",
    // A reference: what it covers is known only when it is evaluated, and an ordinary formula intersects it — the rule
    // for every reference on the path from the root, a range or a function returning one.
    "A1:A3",
    "A1:A3*2",
    "ABS(A1:A3)",
    "Name",
    "INDEX(A1:C3,0,1)",
    "OFFSET(A1,0,0,3)",
    'INDIRECT("A1:A3")',
    "ROW(A1:A3)",
    // A value whatever its arguments, as Microsoft documents them.
    'RTD("x",,"y")',
    'TEXTAFTER("a-b","-")',
    'TEXTBEFORE("a-b","-")',
    'REGEXEXTRACT(A1,"\\d+")',
    'REGEXEXTRACT(A1,"\\d+",0)',
    'TEXTJOIN(",",TRUE,SEQUENCE(3))',
    "LET(x,SEQUENCE(3),SUM(x))",
    "LET(x,1,x+1)",
    "LAMBDA(x,x*2)(3)",
    "LAMBDA(x,LAMBDA(y,x+y))(1)(2)",
    "LET(f,LAMBDA(x,x*2),f(3))",
    "XMATCH(1,A1:A3)"
  ])("%s does not", formula => {
    expect(produces(formula)).toBe(false);
  });
});

describe("producesArray with registered functions", () => {
  const registered = new Map([
    ["CUSTOM", { returnsArray: true }],
    ["PLAIN", {}],
    ["SEQUENCE", { returnsArray: false }]
  ]);
  const with_ = (formula: string): boolean => producesArray(parseFormula(formula), registered);

  it("takes a registered function's own declaration, over a built-in of the same name", () => {
    expect(with_("CUSTOM()")).toBe(true);
    expect(with_("_xlfn.CUSTOM()")).toBe(true);
    expect(with_("CUSTOM()*2")).toBe(true);
    expect(with_("SUM(CUSTOM())")).toBe(false);
    expect(with_("PLAIN()")).toBe(false);
    expect(with_("PLAIN(SEQUENCE(3))")).toBe(false);
    expect(with_("SEQUENCE(3)")).toBe(false);
    expect(with_("FILTER(A1:A3,B1:B3)")).toBe(true);
    // A `LET` name shadows a registered function, as it shadows a built-in.
    expect(with_("LET(custom,LAMBDA(x,x),custom(1))")).toBe(false);
  });

  it("without a registry, a name is the built-in's", () => {
    expect(producesArray(parseFormula("SEQUENCE(3)"))).toBe(true);
    expect(producesArray(parseFormula("CUSTOM()"))).toBe(false);
  });
});
