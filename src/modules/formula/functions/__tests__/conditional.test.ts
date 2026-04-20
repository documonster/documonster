/**
 * Unit tests for conditional aggregates (SUMIF / COUNTIF / AVERAGEIF /
 * MAXIFS / MINIFS and their multi-range counterparts) in `../conditional.ts`.
 *
 * The task highlights several criteria-predicate regressions that must be
 * verified:
 *   - `<>abc` must route through the `<>` operator branch, not through `<`
 *     with `>abc` as the value — a greedy regex prefix was silently
 *     mishandling every not-equal criterion.
 *   - `~*`, `~?`, `~~` must be honoured as literal `*`, `?`, `~` (wildcard
 *     escape). A plain `*` or `?` triggers the wildcard path.
 *   - A blank criteria value must match blank cells.
 */

import { describe, it, expect } from "vitest";

import {
  BLANK,
  ERRORS,
  RVKind,
  rvArray,
  rvBoolean,
  rvNumber,
  rvString,
  type NumberValue,
  type RuntimeValue,
  type ScalarValue
} from "../../runtime/values";
import {
  fnSUMIF,
  fnSUMIFS,
  fnCOUNTIF,
  fnCOUNTIFS,
  fnAVERAGEIF,
  fnAVERAGEIFS,
  fnMAXIFS,
  fnMINIFS,
  buildCriteriaPredicateRV
} from "../conditional";

function asNumber(v: RuntimeValue): number {
  expect(v.kind).toBe(RVKind.Number);
  return (v as NumberValue).value;
}

// Representative ranges reused across tests.
const fruits = rvArray([
  [rvString("apple"), rvString("banana"), rvString("cherry"), rvString("apple"), rvString("banana")]
]);
const quantities = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4), rvNumber(5)]]);

describe("buildCriteriaPredicateRV — operator parsing", () => {
  it("`<>abc` routes to the not-equal branch (regression)", () => {
    // Previously `/^[<>]=?/` would match just `<` and silently treat
    // `<>abc` as the less-than branch. With the ordered alternation, the
    // <> prefix must bind first.
    const pred = buildCriteriaPredicateRV(rvString("<>abc"));
    expect(pred(rvString("abc"))).toBe(false);
    expect(pred(rvString("xyz"))).toBe(true);
    // Number cell against string criterion compares via toStringRV.
    expect(pred(rvNumber(123))).toBe(true);
  });

  it("`>=5` uses the numeric branch", () => {
    const pred = buildCriteriaPredicateRV(rvString(">=5"));
    expect(pred(rvNumber(5))).toBe(true);
    expect(pred(rvNumber(10))).toBe(true);
    expect(pred(rvNumber(4))).toBe(false);
  });

  it("`<5` matches numeric cells only", () => {
    const pred = buildCriteriaPredicateRV(rvString("<5"));
    expect(pred(rvNumber(3))).toBe(true);
    expect(pred(rvNumber(7))).toBe(false);
  });

  it("`<>5` matches non-numeric cells (Excel: NaN !== numVal is true)", () => {
    // Regression: a specialised numeric fast-path must NOT early-return
    // `false` for NaN-coerced cells. Excel counts text cells against a
    // numeric `<>` criterion.
    const pred = buildCriteriaPredicateRV(rvString("<>5"));
    expect(pred(rvNumber(5))).toBe(false);
    expect(pred(rvNumber(7))).toBe(true);
    expect(pred(rvString("abc"))).toBe(true);
    expect(pred(BLANK)).toBe(true);
  });

  it("`=5` with text cells returns false (NaN === 5 is false)", () => {
    const pred = buildCriteriaPredicateRV(rvString("=5"));
    expect(pred(rvNumber(5))).toBe(true);
    expect(pred(rvString("abc"))).toBe(false);
    expect(pred(rvString("5"))).toBe(false); // string "5" stays textual
  });

  it("`>5` against text returns false (NaN > 5 is false)", () => {
    const pred = buildCriteriaPredicateRV(rvString(">5"));
    expect(pred(rvNumber(10))).toBe(true);
    expect(pred(rvString("zzz"))).toBe(false);
    expect(pred(BLANK)).toBe(false); // 0 > 5 is false
  });

  it("`>=0` matches blanks (blank coerces to 0)", () => {
    const pred = buildCriteriaPredicateRV(rvString(">=0"));
    expect(pred(BLANK)).toBe(true);
    expect(pred(rvNumber(-1))).toBe(false);
    expect(pred(rvBoolean(true))).toBe(true); // TRUE → 1
    expect(pred(rvBoolean(false))).toBe(true); // FALSE → 0
  });

  it("`>abc` with string criterion does a string comparison", () => {
    const pred = buildCriteriaPredicateRV(rvString(">abc"));
    expect(pred(rvString("abd"))).toBe(true);
    expect(pred(rvString("abc"))).toBe(false);
    expect(pred(rvString("ABD"))).toBe(true); // case-insensitive
  });

  it("=TEXT matches case-insensitively", () => {
    const pred = buildCriteriaPredicateRV(rvString("=abc"));
    expect(pred(rvString("ABC"))).toBe(true);
    expect(pred(rvString("abcd"))).toBe(false);
  });

  it("blank criteria matches blank or empty-string cells", () => {
    const pred = buildCriteriaPredicateRV(BLANK);
    expect(pred(BLANK)).toBe(true);
    expect(pred(rvString(""))).toBe(true);
    expect(pred(rvString("x"))).toBe(false);
  });

  it("boolean criteria requires exact boolean", () => {
    const pred = buildCriteriaPredicateRV(rvBoolean(true));
    expect(pred(rvBoolean(true))).toBe(true);
    expect(pred(rvBoolean(false))).toBe(false);
    expect(pred(rvNumber(1))).toBe(false);
  });
});

describe("buildCriteriaPredicateRV — wildcards", () => {
  it("`*` matches any sequence", () => {
    const pred = buildCriteriaPredicateRV(rvString("a*"));
    expect(pred(rvString("apple"))).toBe(true);
    expect(pred(rvString("banana"))).toBe(false);
  });

  it("`?` matches any single character", () => {
    const pred = buildCriteriaPredicateRV(rvString("a??le"));
    expect(pred(rvString("apple"))).toBe(true);
    expect(pred(rvString("angle"))).toBe(true);
    expect(pred(rvString("ale"))).toBe(false);
  });

  it("`~*` is a literal asterisk", () => {
    const pred = buildCriteriaPredicateRV(rvString("a~*b"));
    expect(pred(rvString("a*b"))).toBe(true);
    expect(pred(rvString("aXb"))).toBe(false);
  });

  it("`~?` is a literal question mark", () => {
    const pred = buildCriteriaPredicateRV(rvString("a~?b"));
    expect(pred(rvString("a?b"))).toBe(true);
    expect(pred(rvString("aXb"))).toBe(false);
  });

  it("`~~` is a literal tilde", () => {
    const pred = buildCriteriaPredicateRV(rvString("a~~b"));
    expect(pred(rvString("a~b"))).toBe(true);
  });

  it("unescaped wildcards are case-insensitive", () => {
    const pred = buildCriteriaPredicateRV(rvString("A*"));
    expect(pred(rvString("apple"))).toBe(true);
  });

  it("wildcard criteria only match text cells, not numbers (Excel behaviour)", () => {
    // Regression: Excel's COUNTIF / SUMIF wildcards operate on text cells
    // only. A criterion like `"1*"` must NOT match the number 1 (even
    // though its string form is `"1"`); Excel would simply return 0.
    const pred = buildCriteriaPredicateRV(rvString("1*"));
    expect(pred(rvNumber(1))).toBe(false);
    expect(pred(rvNumber(15))).toBe(false);
    expect(pred(rvString("15"))).toBe(true);
    expect(pred(rvString("1abc"))).toBe(true);
  });

  it("`?` wildcard also excludes numeric cells", () => {
    const pred = buildCriteriaPredicateRV(rvString("?"));
    expect(pred(rvNumber(5))).toBe(false);
    expect(pred(rvString("x"))).toBe(true);
    expect(pred(rvBoolean(true))).toBe(false);
  });
});

describe("SUMIF", () => {
  it("sums values in the same range", () => {
    // values > 2 → 3 + 4 + 5 = 12
    expect(asNumber(fnSUMIF([quantities, rvString(">2")]))).toBe(12);
  });

  it("sums a separate sum_range", () => {
    // where fruit = "apple" → qtys[0], qtys[3] = 1 + 4 = 5
    expect(asNumber(fnSUMIF([fruits, rvString("apple"), quantities]))).toBe(5);
  });

  it("supports `<>` criterion (regression)", () => {
    // not "apple" → qty[1]+qty[2]+qty[4] = 2+3+5 = 10
    expect(asNumber(fnSUMIF([fruits, rvString("<>apple"), quantities]))).toBe(10);
  });

  it("propagates errors from the sum_range (Excel behaviour)", () => {
    // Regression: previously silently skipped error cells in the
    // sum-range, hiding `#DIV/0!` / `#VALUE!` under the aggregation.
    // Excel propagates errors.
    const range = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const sumRange = rvArray([[rvNumber(10), ERRORS.DIV0, rvNumber(30)]]);
    // Criterion `>0` matches all three cells; the middle one is #DIV/0!.
    expect(fnSUMIF([range, rvString(">0"), sumRange])).toEqual(ERRORS.DIV0);
  });

  it("returns 0 when no cells match", () => {
    expect(asNumber(fnSUMIF([quantities, rvString(">100")]))).toBe(0);
  });

  it("returns #VALUE! when range is not an array", () => {
    expect(fnSUMIF([rvNumber(5), rvString(">0")])).toEqual(ERRORS.VALUE);
  });
});

describe("COUNTIF", () => {
  it("counts values matching a criterion", () => {
    expect(asNumber(fnCOUNTIF([quantities, rvString(">2")]))).toBe(3);
  });

  it("counts with `<>abc` criterion (regression)", () => {
    expect(asNumber(fnCOUNTIF([fruits, rvString("<>apple")]))).toBe(3);
  });

  it("counts with wildcards", () => {
    // banana, cherry → both contain "na" or "ry"; count apples with a*e
    expect(asNumber(fnCOUNTIF([fruits, rvString("a*e")]))).toBe(2);
  });

  it("counts with literal asterisk via escape", () => {
    const arr = rvArray([[rvString("a*b"), rvString("axb"), rvString("a*b")]]);
    expect(asNumber(fnCOUNTIF([arr, rvString("a~*b")]))).toBe(2);
  });

  it("counts blank cells", () => {
    const arr = rvArray([[rvString("x"), BLANK, rvString(""), rvString("y")]]);
    expect(asNumber(fnCOUNTIF([arr, BLANK]))).toBe(2);
  });
});

describe("SUMIFS / COUNTIFS / AVERAGEIFS", () => {
  const col1 = rvArray([
    [rvString("A"), rvString("B"), rvString("A"), rvString("A"), rvString("B")]
  ]);
  const col2 = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(1), rvNumber(3)]]);
  const sum = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30), rvNumber(40), rvNumber(50)]]);

  it("SUMIFS with two criteria", () => {
    // col1="A" AND col2=1 → rows 0, 3 → 10 + 40 = 50
    expect(asNumber(fnSUMIFS([sum, col1, rvString("A"), col2, rvNumber(1)]))).toBe(50);
  });

  it("SUMIFS with `<>` criterion", () => {
    // col1 <> "A" → rows 1, 4 → 20 + 50 = 70
    expect(asNumber(fnSUMIFS([sum, col1, rvString("<>A")]))).toBe(70);
  });

  it("COUNTIFS with two criteria", () => {
    expect(asNumber(fnCOUNTIFS([col1, rvString("A"), col2, rvNumber(1)]))).toBe(2);
  });

  it("AVERAGEIFS with criteria", () => {
    // col1="A" → rows 0, 2, 3 → sum 10+30+40 = 80, count 3 → 80/3
    expect(asNumber(fnAVERAGEIFS([sum, col1, rvString("A")]))).toBeCloseTo(80 / 3, 10);
  });

  it("AVERAGEIFS returns #DIV/0! when no rows match", () => {
    expect(fnAVERAGEIFS([sum, col1, rvString("Z")])).toEqual(ERRORS.DIV0);
  });
});

describe("MAXIFS / MINIFS", () => {
  const col = rvArray([
    [rvString("A"), rvString("B"), rvString("A"), rvString("A"), rvString("B")]
  ]);
  const vals = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30), rvNumber(40), rvNumber(50)]]);

  it("MAXIFS returns the largest matching value", () => {
    // A → 10, 30, 40 → 40
    expect(asNumber(fnMAXIFS([vals, col, rvString("A")]))).toBe(40);
  });

  it("MINIFS returns the smallest matching value", () => {
    expect(asNumber(fnMINIFS([vals, col, rvString("A")]))).toBe(10);
  });

  it("MAXIFS returns 0 when nothing matches", () => {
    expect(asNumber(fnMAXIFS([vals, col, rvString("Z")]))).toBe(0);
  });

  it("MINIFS returns 0 when nothing matches", () => {
    expect(asNumber(fnMINIFS([vals, col, rvString("Z")]))).toBe(0);
  });
});

describe("AVERAGEIF delegation", () => {
  // AVERAGEIF lives in conditional.ts; the statistical.test.ts file checks it
  // through a different lens. Here we verify the same regression path
  // (`<>`) goes through.
  it("AVERAGEIF with `<>` criterion", () => {
    const range = rvArray([[rvString("x"), rvString("y"), rvString("x")]]);
    const vals = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
    // where cell <> "x" → row 1 → value 20
    expect(asNumber(fnAVERAGEIF([range, rvString("<>x"), vals]))).toBe(20);
  });
});

// ============================================================================
// Comprehensive coverage — each function gets >= 5 dedicated cases.
// ============================================================================

describe("SUMIF comprehensive", () => {
  it("numeric criteria exact match", () => {
    const arr = rvArray([[rvNumber(1), rvNumber(2), rvNumber(2), rvNumber(3)]]);
    expect(asNumber(fnSUMIF([arr, rvNumber(2)]))).toBe(4);
  });

  it("`<=` numeric", () => {
    const arr = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4)]]);
    expect(asNumber(fnSUMIF([arr, rvString("<=2")]))).toBe(3);
  });

  it("wildcard '*abc*'", () => {
    const names = rvArray([[rvString("xabcy"), rvString("bcd"), rvString("abcd")]]);
    const qty = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
    expect(asNumber(fnSUMIF([names, rvString("*abc*"), qty]))).toBe(40);
  });

  it("wildcard '?bc'", () => {
    const names = rvArray([[rvString("abc"), rvString("xbc"), rvString("bcd")]]);
    const qty = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
    expect(asNumber(fnSUMIF([names, rvString("?bc"), qty]))).toBe(30);
  });

  it("literal '~*'", () => {
    const names = rvArray([[rvString("*"), rvString("a"), rvString("*")]]);
    const qty = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    expect(asNumber(fnSUMIF([names, rvString("~*"), qty]))).toBe(4);
  });

  it("non-numeric sum cells are ignored", () => {
    const range = rvArray([[rvString("a"), rvString("a")]]);
    const sumR = rvArray([[rvString("not num"), rvNumber(5)]]);
    expect(asNumber(fnSUMIF([range, rvString("a"), sumR]))).toBe(5);
  });

  it("single-cell range", () => {
    const r = rvArray([[rvNumber(5)]]);
    expect(asNumber(fnSUMIF([r, rvString(">0")]))).toBe(5);
  });

  it("M×N array", () => {
    const data = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)]
    ]);
    expect(asNumber(fnSUMIF([data, rvString(">1")]))).toBe(9);
  });

  it("criteria is error -> error", () => {
    const arr = rvArray([[rvNumber(1)]]);
    expect(fnSUMIF([arr, ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("range is scalar -> #VALUE!", () => {
    expect(fnSUMIF([rvNumber(5), rvNumber(5)])).toEqual(ERRORS.VALUE);
  });
});

describe("SUMIFS comprehensive", () => {
  const c1 = rvArray([[rvString("A"), rvString("A"), rvString("B")]]);
  const c2 = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
  const sums = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);

  it("two criteria AND", () => {
    expect(asNumber(fnSUMIFS([sums, c1, rvString("A"), c2, rvString(">=2")]))).toBe(20);
  });

  it("less than 3 args -> #VALUE!", () => {
    expect(fnSUMIFS([sums])).toEqual(ERRORS.VALUE);
    expect(fnSUMIFS([sums, c1])).toEqual(ERRORS.VALUE);
  });

  it("criteria range shape mismatch -> #VALUE!", () => {
    const short = rvArray([[rvString("A"), rvString("A")]]);
    expect(fnSUMIFS([sums, short, rvString("A")])).toEqual(ERRORS.VALUE);
  });

  it("empty result -> 0", () => {
    expect(asNumber(fnSUMIFS([sums, c1, rvString("Z")]))).toBe(0);
  });

  it("criteria error propagates", () => {
    expect(fnSUMIFS([sums, c1, ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("sum range scalar -> #VALUE!", () => {
    expect(fnSUMIFS([rvNumber(5), c1, rvString("A")])).toEqual(ERRORS.VALUE);
  });

  it("wildcards in criteria", () => {
    const names = rvArray([[rvString("apple"), rvString("apricot"), rvString("banana")]]);
    const q = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    expect(asNumber(fnSUMIFS([q, names, rvString("ap*")]))).toBe(3);
  });
});

describe("COUNTIF comprehensive", () => {
  it("counts exact number", () => {
    const arr = rvArray([[rvNumber(1), rvNumber(2), rvNumber(1)]]);
    expect(asNumber(fnCOUNTIF([arr, rvNumber(1)]))).toBe(2);
  });

  it("counts with `>`", () => {
    const arr = rvArray([[rvNumber(1), rvNumber(5), rvNumber(10)]]);
    expect(asNumber(fnCOUNTIF([arr, rvString(">3")]))).toBe(2);
  });

  it("counts boolean cells with numeric compare", () => {
    const arr = rvArray([[rvBoolean(true), rvBoolean(false), rvBoolean(true)]]);
    expect(asNumber(fnCOUNTIF([arr, rvString(">0")]))).toBe(2);
  });

  it("COUNTIF with '*' wildcard matches only text cells (Excel behaviour)", () => {
    // `*` matches 0+ of any character — but Excel restricts wildcard
    // matching to TEXT cells. Numbers, blanks, booleans are not matched
    // by `*` even though their stringified forms would be non-empty.
    // Previously we stringified every cell through `toStringRV` which
    // made `*` match numbers and emptied cells — a soft deviation from
    // Excel. The tightened semantics match Excel's documented rule.
    const arr = rvArray([[rvString("a"), rvNumber(1), rvString(""), BLANK]]);
    expect(asNumber(fnCOUNTIF([arr, rvString("*")]))).toBe(2); // "a" + ""
  });

  it("empty result -> 0", () => {
    const arr = rvArray([[rvNumber(1)]]);
    expect(asNumber(fnCOUNTIF([arr, rvString(">100")]))).toBe(0);
  });

  it("range scalar -> #VALUE!", () => {
    expect(fnCOUNTIF([rvNumber(1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("criteria error -> error", () => {
    expect(fnCOUNTIF([rvArray([[rvNumber(1)]]), ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("N×1 array", () => {
    const col = rvArray([[rvString("a")], [rvString("b")], [rvString("a")]]);
    expect(asNumber(fnCOUNTIF([col, rvString("a")]))).toBe(2);
  });

  it("M×N array", () => {
    const m = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(1)]
    ]);
    expect(asNumber(fnCOUNTIF([m, rvNumber(1)]))).toBe(2);
  });
});

describe("COUNTIFS comprehensive", () => {
  const c1 = rvArray([[rvString("A"), rvString("B"), rvString("A")]]);
  const c2 = rvArray([[rvNumber(1), rvNumber(1), rvNumber(2)]]);

  it("two-criteria AND", () => {
    expect(asNumber(fnCOUNTIFS([c1, rvString("A"), c2, rvNumber(1)]))).toBe(1);
  });

  it("single-criterion (same as COUNTIF)", () => {
    expect(asNumber(fnCOUNTIFS([c1, rvString("A")]))).toBe(2);
  });

  it("< 2 args -> #VALUE!", () => {
    expect(fnCOUNTIFS([c1])).toEqual(ERRORS.VALUE);
  });

  it("first arg scalar -> #VALUE!", () => {
    expect(fnCOUNTIFS([rvNumber(1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("shape mismatch -> #VALUE!", () => {
    const short = rvArray([[rvNumber(1), rvNumber(1)]]);
    expect(fnCOUNTIFS([c1, rvString("A"), short, rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("empty result -> 0", () => {
    expect(asNumber(fnCOUNTIFS([c1, rvString("Z")]))).toBe(0);
  });

  it("criteria-error propagates", () => {
    expect(fnCOUNTIFS([c1, ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("AVERAGEIF comprehensive", () => {
  it("basic", () => {
    const arr = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
    expect(asNumber(fnAVERAGEIF([arr, rvString(">15")]))).toBe(25);
  });

  it("with separate avg_range", () => {
    const crit = rvArray([[rvString("a"), rvString("b"), rvString("a")]]);
    const vals = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
    expect(asNumber(fnAVERAGEIF([crit, rvString("a"), vals]))).toBe(20);
  });

  it("no match -> #DIV/0!", () => {
    const arr = rvArray([[rvNumber(1), rvNumber(2)]]);
    expect(fnAVERAGEIF([arr, rvString(">100")])).toEqual(ERRORS.DIV0);
  });

  it("range scalar -> #VALUE!", () => {
    expect(fnAVERAGEIF([rvNumber(1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("criteria error -> error", () => {
    expect(fnAVERAGEIF([rvArray([[rvNumber(1)]]), ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("non-numeric matched cells excluded from average", () => {
    const range = rvArray([[rvString("x"), rvString("x"), rvString("x")]]);
    const vals = rvArray([[rvNumber(10), rvString("not"), rvNumber(30)]]);
    expect(asNumber(fnAVERAGEIF([range, rvString("x"), vals]))).toBe(20);
  });

  it("single-cell range match", () => {
    const arr = rvArray([[rvNumber(5)]]);
    expect(asNumber(fnAVERAGEIF([arr, rvString(">0")]))).toBe(5);
  });
});

describe("AVERAGEIFS comprehensive", () => {
  const avg = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30), rvNumber(40)]]);
  const c1 = rvArray([[rvString("A"), rvString("A"), rvString("B"), rvString("B")]]);

  it("basic", () => {
    expect(asNumber(fnAVERAGEIFS([avg, c1, rvString("A")]))).toBe(15);
  });

  it("no match -> #DIV/0!", () => {
    expect(fnAVERAGEIFS([avg, c1, rvString("Z")])).toEqual(ERRORS.DIV0);
  });

  it("< 3 args -> #VALUE!", () => {
    expect(fnAVERAGEIFS([avg])).toEqual(ERRORS.VALUE);
  });

  it("avg scalar -> #VALUE!", () => {
    expect(fnAVERAGEIFS([rvNumber(1), c1, rvString("A")])).toEqual(ERRORS.VALUE);
  });

  it("shape mismatch -> #VALUE!", () => {
    const short = rvArray([[rvString("A"), rvString("A")]]);
    expect(fnAVERAGEIFS([avg, short, rvString("A")])).toEqual(ERRORS.VALUE);
  });

  it("criteria error -> error", () => {
    expect(fnAVERAGEIFS([avg, c1, ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("MAXIFS comprehensive", () => {
  const vals = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
  const c = rvArray([[rvString("A"), rvString("A"), rvString("B")]]);

  it("basic", () => {
    expect(asNumber(fnMAXIFS([vals, c, rvString("A")]))).toBe(20);
  });

  it("no match returns 0", () => {
    expect(asNumber(fnMAXIFS([vals, c, rvString("Z")]))).toBe(0);
  });

  it("negative values", () => {
    const v = rvArray([[rvNumber(-5), rvNumber(-10), rvNumber(-1)]]);
    expect(asNumber(fnMAXIFS([v, c, rvString("A")]))).toBe(-5);
  });

  it("< 3 args -> #VALUE!", () => {
    expect(fnMAXIFS([vals])).toEqual(ERRORS.VALUE);
  });

  it("max-range scalar -> #VALUE!", () => {
    expect(fnMAXIFS([rvNumber(1), c, rvString("A")])).toEqual(ERRORS.VALUE);
  });

  it("criteria error propagates", () => {
    expect(fnMAXIFS([vals, c, ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("with multi-criteria", () => {
    const c2 = rvArray([[rvNumber(1), rvNumber(2), rvNumber(1)]]);
    expect(asNumber(fnMAXIFS([vals, c, rvString("A"), c2, rvNumber(2)]))).toBe(20);
  });
});

describe("MINIFS comprehensive", () => {
  const vals = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
  const c = rvArray([[rvString("A"), rvString("A"), rvString("B")]]);

  it("basic", () => {
    expect(asNumber(fnMINIFS([vals, c, rvString("A")]))).toBe(10);
  });

  it("no match returns 0", () => {
    expect(asNumber(fnMINIFS([vals, c, rvString("Z")]))).toBe(0);
  });

  it("only non-numeric matches -> 0", () => {
    const v = rvArray([[rvString("x"), rvString("y"), rvString("z")]]);
    expect(asNumber(fnMINIFS([v, c, rvString("A")]))).toBe(0);
  });

  it("< 3 args -> #VALUE!", () => {
    expect(fnMINIFS([vals])).toEqual(ERRORS.VALUE);
  });

  it("criteria error propagates", () => {
    expect(fnMINIFS([vals, c, ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });

  it("negative values", () => {
    const v = rvArray([[rvNumber(-5), rvNumber(-1), rvNumber(100)]]);
    expect(asNumber(fnMINIFS([v, c, rvString("A")]))).toBe(-5);
  });
});

describe("buildCriteriaPredicateRV — extended", () => {
  it("numeric criteria matches only numeric cells (not strings)", () => {
    const pred = buildCriteriaPredicateRV(rvNumber(5));
    expect(pred(rvNumber(5))).toBe(true);
    expect(pred(rvString("5"))).toBe(false);
  });

  it("`>0` accepts booleans (TRUE=1 > 0)", () => {
    const pred = buildCriteriaPredicateRV(rvString(">0"));
    expect(pred(rvBoolean(true))).toBe(true);
    expect(pred(rvBoolean(false))).toBe(false);
  });

  it("`>=10` numeric-only — falls back to string compare for non-num cells", () => {
    const pred = buildCriteriaPredicateRV(rvString(">=10"));
    // Number cell with value 10 matches
    expect(pred(rvNumber(10))).toBe(true);
    expect(pred(rvNumber(9))).toBe(false);
  });

  it("error criteria always false", () => {
    const pred = buildCriteriaPredicateRV(ERRORS.NA);
    expect(pred(rvNumber(1))).toBe(false);
    expect(pred(rvString("a"))).toBe(false);
  });

  it("numeric literal criteria string '42' matches number 42", () => {
    const pred = buildCriteriaPredicateRV(rvString("42"));
    expect(pred(rvNumber(42))).toBe(true);
    expect(pred(rvString("42"))).toBe(false);
  });
});

// ============================================================================
// R8 deep coverage: operator parsing and type coercion matrix
// ============================================================================

describe("criteria operator matrix", () => {
  const range = rvArray([[rvNumber(1)], [rvNumber(5)], [rvNumber(10)], [rvNumber(15)]]);

  it("'=5' matches exact", () => {
    expect(asNumber(fnCOUNTIF([range, rvString("=5")]))).toBe(1);
  });

  it("'>5' matches strictly greater", () => {
    expect(asNumber(fnCOUNTIF([range, rvString(">5")]))).toBe(2);
  });

  it("'<5' matches strictly less", () => {
    expect(asNumber(fnCOUNTIF([range, rvString("<5")]))).toBe(1);
  });

  it("'>=5' matches greater-or-equal", () => {
    expect(asNumber(fnCOUNTIF([range, rvString(">=5")]))).toBe(3);
  });

  it("'<=5' matches less-or-equal", () => {
    expect(asNumber(fnCOUNTIF([range, rvString("<=5")]))).toBe(2);
  });

  it("'<>5' matches not-equal (R2 regression)", () => {
    expect(asNumber(fnCOUNTIF([range, rvString("<>5")]))).toBe(3);
  });

  it("boolean coercion: '>0' with TRUE=1, FALSE=0", () => {
    const bools = rvArray([[rvBoolean(true)], [rvBoolean(false)], [rvBoolean(true)]]);
    expect(asNumber(fnCOUNTIF([bools, rvString(">0")]))).toBe(2);
  });

  it("blank coercion: blank = 0", () => {
    const blanks = rvArray([[{ kind: RVKind.Blank } as ScalarValue], [rvNumber(1)]]);
    expect(asNumber(fnCOUNTIF([blanks, rvString("=0")]))).toBe(1);
  });

  it("direct number criterion (not string)", () => {
    expect(asNumber(fnCOUNTIF([range, rvNumber(5)]))).toBe(1);
  });

  it("string criterion without operator = exact equality", () => {
    const arr = rvArray([[rvString("abc")], [rvString("def")], [rvString("abc")]]);
    expect(asNumber(fnCOUNTIF([arr, rvString("abc")]))).toBe(2);
  });
});

describe("wildcard matching", () => {
  const names = rvArray([
    [rvString("apple")],
    [rvString("banana")],
    [rvString("apricot")],
    [rvString("cherry")]
  ]);

  it("'*' matches all strings", () => {
    expect(asNumber(fnCOUNTIF([names, rvString("*")]))).toBe(4);
  });

  it("'a*' matches prefix", () => {
    expect(asNumber(fnCOUNTIF([names, rvString("a*")]))).toBe(2);
  });

  it("'*e*' matches containing", () => {
    expect(asNumber(fnCOUNTIF([names, rvString("*e*")]))).toBe(2); // apple, cherry
  });

  it("'?p*' matches single char", () => {
    expect(asNumber(fnCOUNTIF([names, rvString("?p*")]))).toBe(2); // apple, apricot
  });

  it("'~*' matches literal asterisk", () => {
    const arr = rvArray([[rvString("a*b")], [rvString("ab")], [rvString("a*")]]);
    expect(asNumber(fnCOUNTIF([arr, rvString("a~*")]))).toBe(1); // only "a*"
  });

  it("'~?' matches literal question mark", () => {
    const arr = rvArray([[rvString("a?b")], [rvString("axb")]]);
    expect(asNumber(fnCOUNTIF([arr, rvString("a~?b")]))).toBe(1);
  });

  it("case-insensitive matching", () => {
    expect(asNumber(fnCOUNTIF([names, rvString("APPLE")]))).toBe(1);
  });
});

describe("SUMIFS criteria range shape validation (R4 regression)", () => {
  const sumRange = rvArray([[rvNumber(10)], [rvNumber(20)], [rvNumber(30)]]);

  it("requires matching criteria range shape", () => {
    const wrongShape = rvArray([[rvString("A"), rvString("B")]]); // 1×2 vs 3×1
    expect(fnSUMIFS([sumRange, wrongShape, rvString("A")])).toEqual(ERRORS.VALUE);
  });

  it("accepts matching shape", () => {
    const crit = rvArray([[rvString("A")], [rvString("A")], [rvString("B")]]);
    expect(asNumber(fnSUMIFS([sumRange, crit, rvString("A")]))).toBe(30); // rows 1+2
  });

  it("multiple criteria all must match", () => {
    const c1 = rvArray([[rvString("x")], [rvString("y")], [rvString("x")]]);
    const c2 = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(1)]]);
    expect(asNumber(fnSUMIFS([sumRange, c1, rvString("x"), c2, rvNumber(1)]))).toBe(40); // rows 1+3
  });
});

describe("COUNTIFS with multiple conditions", () => {
  it("two conditions AND", () => {
    const a = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    const b = rvArray([[rvString("x")], [rvString("x")], [rvString("y")]]);
    expect(asNumber(fnCOUNTIFS([a, rvString(">=2"), b, rvString("x")]))).toBe(1);
  });

  it("zero matches returns 0", () => {
    const a = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    expect(asNumber(fnCOUNTIFS([a, rvString(">99")]))).toBe(0);
  });
});

describe("AVERAGEIFS edge cases", () => {
  it("no matches → #DIV/0!", () => {
    const avg = rvArray([[rvNumber(10)], [rvNumber(20)]]);
    const crit = rvArray([[rvString("x")], [rvString("y")]]);
    expect(fnAVERAGEIFS([avg, crit, rvString("z")])).toEqual(ERRORS.DIV0);
  });

  it("averages only numeric cells matching criteria", () => {
    const avg = rvArray([[rvNumber(10)], [rvString("text")], [rvNumber(30)]]);
    const crit = rvArray([[rvString("a")], [rvString("a")], [rvString("a")]]);
    expect(asNumber(fnAVERAGEIFS([avg, crit, rvString("a")]))).toBe(20); // (10+30)/2
  });
});

describe("MAXIFS / MINIFS edge cases", () => {
  it("MAXIFS no matches returns 0", () => {
    const m = rvArray([[rvNumber(10)]]);
    const crit = rvArray([[rvString("a")]]);
    expect(asNumber(fnMAXIFS([m, crit, rvString("b")]))).toBe(0);
  });

  it("MINIFS no matches returns 0", () => {
    const m = rvArray([[rvNumber(10)]]);
    const crit = rvArray([[rvString("a")]]);
    expect(asNumber(fnMINIFS([m, crit, rvString("b")]))).toBe(0);
  });

  it("MAXIFS with multiple conditions", () => {
    const m = rvArray([[rvNumber(10)], [rvNumber(20)], [rvNumber(30)]]);
    const c1 = rvArray([[rvString("a")], [rvString("a")], [rvString("b")]]);
    const c2 = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(1)]]);
    expect(asNumber(fnMAXIFS([m, c1, rvString("a"), c2, rvNumber(2)]))).toBe(20);
  });
});

describe("SUMIF avg_range longer than check range", () => {
  it("uses corresponding cells", () => {
    const check = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    const avg = rvArray([[rvNumber(10)], [rvNumber(20)], [rvNumber(30)]]);
    expect(asNumber(fnSUMIF([check, rvString(">=2"), avg]))).toBe(50);
  });
});

describe("MINIFS extras (R9 saturation)", () => {
  const min = rvArray([[rvNumber(5)], [rvNumber(10)], [rvNumber(3)], [rvNumber(8)]]);
  const c1 = rvArray([[rvString("a")], [rvString("b")], [rvString("a")], [rvString("b")]]);

  it("single condition filters", () => {
    expect(asNumber(fnMINIFS([min, c1, rvString("a")]))).toBe(3);
  });
  it("no match returns 0", () => {
    expect(asNumber(fnMINIFS([min, c1, rvString("z")]))).toBe(0);
  });
  it("numeric criterion < threshold", () => {
    expect(asNumber(fnMINIFS([min, min, rvString(">=5")]))).toBe(5);
  });
  it("multiple conditions narrow further", () => {
    const c2 = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(1)], [rvNumber(2)]]);
    expect(asNumber(fnMINIFS([min, c1, rvString("a"), c2, rvNumber(1)]))).toBe(3);
  });
  it("skips non-numeric in target range", () => {
    const mixed = rvArray([[rvString("x")], [rvNumber(10)], [rvNumber(3)]]);
    const crit = rvArray([[rvString("a")], [rvString("a")], [rvString("a")]]);
    expect(asNumber(fnMINIFS([mixed, crit, rvString("a")]))).toBe(3);
  });
  it("shape mismatch → #VALUE!", () => {
    const bad = rvArray([[rvString("a"), rvString("b")]]);
    expect(fnMINIFS([min, bad, rvString("a")])).toEqual(ERRORS.VALUE);
  });
});
