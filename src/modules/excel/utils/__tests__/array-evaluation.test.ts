import { needsArrayEvaluation, ordinaryEvaluationFormula } from "@excel/utils/array-evaluation";
import { describe, expect, it } from "vitest";

describe("needsArrayEvaluation", () => {
  it.each([
    // A range compared inside a function that consumes the result whole: legacy Excel compares one cell.
    'MEDIAN(IF(Samples[Group]="A",Samples[Response]))',
    'CORREL(IF(T[g]="x",T[a]),IF(T[g]="x",T[b]))',
    'MAX(IF(T[g]="x",T[r]))-MIN(IF(T[g]="x",T[r]))',
    "SUM(A1:A3*2)",
    "SUM(IFERROR(A1:A3,0))",
    // An array parameter evaluates over arrays in legacy Excel too — except through IF's condition.
    "SUMPRODUCT(IF(A1:A3>1,1,0))",
    "INDEX(A:A,MATCH(1,(B:B=1)*(C:C=2),0))",
    'TEXTJOIN(",",TRUE,IF(A1:A3>1,A1:A3,""))'
  ])("%s", formula => {
    expect(needsArrayEvaluation(formula)).toBe(true);
    expect(needsArrayEvaluation(`=${formula}`)).toBe(true);
  });

  it.each([
    "SUM(A1:A3)",
    "SUMPRODUCT(A1:A3*B1:B3)",
    'COUNTIF(T[g],"x")',
    'AVERAGEIF(T[g],"x",T[r])',
    "SLOPE(T[r],T[d])",
    "VLOOKUP(A1,B:C,2,0)",
    "T[@a]*2",
    "A1+B1",
    // A range on the path from the root: as a dynamic array it would spill, so it stays ordinary and is intersected.
    "A1:A3",
    "ABS(A1:A3)",
    // An array parameter takes an expression over a range whole, as Excel's own tokens show (`65`).
    "ROWS(A1:A3+0)",
    "INDEX(A1:A3*3,2)",
    // Explicit intersection means what it says in both.
    "SUM(@A1:A3*2)",
    // Would need an array inside and spill at the root: left alone rather than trade one difference for another.
    "IF(SUM(A1:A3*2)>0,B1:B3)",
    "not a (formula"
  ])("%s", formula => {
    expect(needsArrayEvaluation(formula)).toBe(false);
  });
});

describe("ordinaryEvaluationFormula", () => {
  it("marks each operand Excel intersects in a formula stored as an ordinary one", () => {
    expect(ordinaryEvaluationFormula('MEDIAN(IF(T[g]="A",T[r]))', true)).toBe(
      'MEDIAN(IF(@T[g]="A",T[r]))'
    );
    expect(ordinaryEvaluationFormula("IF(SUM(A1:A3*2)>0,B1:B3)", true)).toBe(
      "IF(SUM(@A1:A3*2)>0,B1:B3)"
    );
    expect(ordinaryEvaluationFormula("ABS(A1:A3)", true)).toBe("ABS(@A1:A3)");
    expect(ordinaryEvaluationFormula("SUM(Rng*2)", true)).toBe("SUM(@Rng*2)");
    expect(ordinaryEvaluationFormula("SUM(LEN(B1:B5))", true)).toBe("SUM(LEN(@B1:B5))");
  });

  it("leaves what Excel takes whole", () => {
    expect(ordinaryEvaluationFormula("SUM(A1:A3)", true)).toBeUndefined();
    expect(ordinaryEvaluationFormula("A1:A3", true)).toBeUndefined();
    expect(ordinaryEvaluationFormula("ROWS(A1:A5+0)", true)).toBeUndefined();
    expect(ordinaryEvaluationFormula("SUMPRODUCT(LEN(B1:B5))", true)).toBeUndefined();
    expect(ordinaryEvaluationFormula("SUM(@A1:A3*2)", true)).toBeUndefined();
  });

  it("leaves a formula set through the API that is written as a dynamic array", () => {
    expect(ordinaryEvaluationFormula("SUM(LEN(B1:B5))", false)).toBeUndefined();
    expect(ordinaryEvaluationFormula("ABS(A1:A3)", false)).toBe("ABS(@A1:A3)");
  });
});
