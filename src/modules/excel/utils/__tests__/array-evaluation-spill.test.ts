/**
 * A spill reference — `A1#`, or `ANCHORARRAY(A1)` as a file stores it — is a range: the anchor's spill range. It exists
 * only in dynamic-array Excel, which spills any formula that reaches it with an array, so a formula set with one on the
 * path from its root is written as a dynamic array and spills here too. Under a parameter that takes it whole it is a
 * range like any other.
 */
import { needsArrayEvaluation, ordinaryEvaluationFormula } from "@excel/utils/array-evaluation";
import { describe, expect, it } from "vitest";

describe("needsArrayEvaluation with a spill reference", () => {
  it.each([
    "A1#",
    "A1#*2",
    "ANCHORARRAY(A1)*2",
    "_xlfn.ANCHORARRAY(A1)*2",
    "S!$A$1#+B1",
    "ABS(A1#)",
    "IF(A1#>1,1,0)",
    "A1#*B1:B3",
    // Below a parameter that takes it whole it is intersected by legacy Excel, as `SUM(A1:A3*2)` is.
    "SUM(A1#*2)",
    "SUM(ANCHORARRAY(A1)*2)",
    // One element or a whole row of it, depending on the spill's shape: decided for the array.
    "INDEX(A1#,2)"
  ])("%s is a dynamic array", formula => {
    expect(needsArrayEvaluation(formula)).toBe(true);
  });

  it.each(["SUM(A1#)", "ROWS(A1#)", "@A1#*2", "SUM(@A1#*2)", "SUM(ANCHORARRAY(A1))"])(
    "%s is not",
    formula => {
      expect(needsArrayEvaluation(formula)).toBe(false);
    }
  );
});

describe("ordinaryEvaluationFormula with a spill reference", () => {
  it("evaluates a formula set through the API as written", () => {
    for (const formula of ["A1#*2", "ANCHORARRAY(A1)*2", "SUM(A1#*2)", "ABS(A1#)"]) {
      expect(ordinaryEvaluationFormula(formula, false)).toBeUndefined();
    }
  });

  it("intersects it where Excel intersects a range, in a formula a file stored as an ordinary one", () => {
    expect(ordinaryEvaluationFormula("A1#*2", true)).toBe("@A1#*2");
    expect(ordinaryEvaluationFormula("SUM(A1#*2)", true)).toBe("SUM(@A1#*2)");
    expect(ordinaryEvaluationFormula("_xlfn.ANCHORARRAY(A1)*2", true)).toBe(
      "@_xlfn.ANCHORARRAY(A1)*2"
    );
    expect(ordinaryEvaluationFormula("SUM(A1#)", true)).toBeUndefined();
    expect(ordinaryEvaluationFormula("A1#", true)).toBeUndefined();
  });
});
