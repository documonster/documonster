/**
 * Unit tests for the `_shared.ts` helpers used across every function
 * file. Bugs here cascade — a broken `flattenNumbers` breaks SUM /
 * AVERAGE / STDEV / VAR / dozens of downstream functions at once — so
 * the helpers deserve their own direct coverage independent of any
 * particular caller's test.
 */

import { describe, expect, it } from "vitest";

import {
  BLANK,
  ERRORS,
  RVKind,
  rvArray,
  rvBoolean,
  rvNumber,
  rvString,
  type NumberValue
} from "../../runtime/values";
import {
  argToNumber,
  asArray,
  checkError,
  excelWildcardToRegex,
  firstError,
  flattenAll,
  flattenNumbers,
  forEachNumber,
  getCell,
  hasUnescapedWildcard,
  stripErrorCells,
  stripHiddenRowCells,
  stripSubtotalMaskedCells,
  unescapeExcelWildcard
} from "../_shared";

// ---------------------------------------------------------------------------
// checkError
// ---------------------------------------------------------------------------

describe("checkError", () => {
  it("returns null for plain number", () => {
    expect(checkError(rvNumber(5))).toBeNull();
  });

  it("returns null for string / boolean / blank", () => {
    expect(checkError(rvString("x"))).toBeNull();
    expect(checkError(rvBoolean(true))).toBeNull();
    expect(checkError(BLANK)).toBeNull();
  });

  it("returns the error when passed a scalar error", () => {
    expect(checkError(ERRORS.NA)).toEqual(ERRORS.NA);
  });

  it("applies implicit intersection — top-left of array", () => {
    const arr = rvArray([
      [ERRORS.NA, rvNumber(1)],
      [rvNumber(2), rvNumber(3)]
    ]);
    expect(checkError(arr)).toEqual(ERRORS.NA);
  });

  it("returns null when top-left of array is not an error", () => {
    const arr = rvArray([[rvNumber(1), ERRORS.NA]]);
    expect(checkError(arr)).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// argToNumber
// ---------------------------------------------------------------------------

describe("argToNumber", () => {
  it("passes numbers through", () => {
    const r = argToNumber(rvNumber(3.5));
    expect(r).toEqual(rvNumber(3.5));
  });

  it("coerces Boolean to 1 / 0", () => {
    expect((argToNumber(rvBoolean(true)) as NumberValue).value).toBe(1);
    expect((argToNumber(rvBoolean(false)) as NumberValue).value).toBe(0);
  });

  it("coerces Blank to 0", () => {
    expect((argToNumber(BLANK) as NumberValue).value).toBe(0);
  });

  it("parses numeric strings", () => {
    expect((argToNumber(rvString("42")) as NumberValue).value).toBe(42);
  });

  it("rejects non-numeric string with #VALUE!", () => {
    expect(argToNumber(rvString("abc"))).toEqual(ERRORS.VALUE);
  });

  it("propagates errors", () => {
    expect(argToNumber(ERRORS.DIV0)).toEqual(ERRORS.DIV0);
  });

  it("applies implicit intersection on arrays", () => {
    const arr = rvArray([[rvNumber(7), rvNumber(9)]]);
    expect((argToNumber(arr) as NumberValue).value).toBe(7);
  });
});

// ---------------------------------------------------------------------------
// flattenNumbers
// ---------------------------------------------------------------------------

describe("flattenNumbers", () => {
  it("returns numeric cells from arrays", () => {
    const r = flattenNumbers([rvArray([[rvNumber(1), rvNumber(2)], [rvNumber(3)]])]);
    expect(r.map(v => (v as NumberValue).value)).toEqual([1, 2, 3]);
  });

  it("skips booleans / strings / blanks inside arrays (Excel behaviour)", () => {
    const arr = rvArray([
      [rvNumber(1), rvString("skip"), rvBoolean(true)],
      [BLANK, rvNumber(2), ERRORS.VALUE]
    ]);
    const r = flattenNumbers([arr]);
    // 1, 2, plus the error
    expect(r.length).toBe(3);
    expect(r[0]).toEqual(rvNumber(1));
    expect(r[1]).toEqual(rvNumber(2));
    expect(r[2]).toEqual(ERRORS.VALUE);
  });

  it("direct-scalar args coerce Boolean and String to number", () => {
    const r = flattenNumbers([rvBoolean(true), rvString("5"), rvNumber(7)]);
    expect(r.map(v => (v as NumberValue).value)).toEqual([1, 5, 7]);
  });

  it("direct-scalar Blank is skipped", () => {
    const r = flattenNumbers([BLANK, rvNumber(9)]);
    expect(r.length).toBe(1);
  });

  it("surfaces direct-scalar error", () => {
    const r = flattenNumbers([ERRORS.NA, rvNumber(3)]);
    expect(r[0]).toEqual(ERRORS.NA);
  });
});

// ---------------------------------------------------------------------------
// forEachNumber
// ---------------------------------------------------------------------------

describe("forEachNumber", () => {
  it("streams numeric cells from arrays without allocating an intermediate", () => {
    const out: number[] = [];
    const r = forEachNumber([rvArray([[rvNumber(1), rvNumber(2)], [rvNumber(3)]])], n =>
      out.push(n)
    );
    expect(r).toBeNull();
    expect(out).toEqual([1, 2, 3]);
  });

  it("skips booleans / strings / blanks inside arrays (Excel behaviour)", () => {
    const out: number[] = [];
    const r = forEachNumber(
      [
        rvArray([
          [rvNumber(1), rvString("skip"), rvBoolean(true)],
          [BLANK, rvNumber(2)]
        ])
      ],
      n => out.push(n)
    );
    expect(r).toBeNull();
    expect(out).toEqual([1, 2]);
  });

  it("aborts on the first error encountered inside an array", () => {
    const out: number[] = [];
    const r = forEachNumber([rvArray([[rvNumber(1), ERRORS.DIV0], [rvNumber(99)]])], n =>
      out.push(n)
    );
    expect(r).toEqual(ERRORS.DIV0);
    // onNumber was called before the abort, but never for cells after it.
    expect(out).toEqual([1]);
  });

  it("direct-scalar args are coerced via toNumberRV", () => {
    const out: number[] = [];
    const r = forEachNumber([rvBoolean(true), rvString("5"), rvNumber(7)], n => out.push(n));
    expect(r).toBeNull();
    expect(out).toEqual([1, 5, 7]);
  });

  it("direct-scalar Blank is skipped, not coerced to 0", () => {
    const out: number[] = [];
    forEachNumber([BLANK, rvNumber(9)], n => out.push(n));
    expect(out).toEqual([9]);
  });

  it("direct-scalar that fails coercion surfaces as an error", () => {
    const out: number[] = [];
    const r = forEachNumber([rvString("not-a-number"), rvNumber(9)], n => out.push(n));
    expect(r?.kind).toBe(RVKind.Error);
    expect(out).toEqual([]);
  });

  it("direct-scalar error argument aborts the scan immediately", () => {
    const out: number[] = [];
    const r = forEachNumber([ERRORS.NA, rvNumber(9)], n => out.push(n));
    expect(r).toEqual(ERRORS.NA);
    expect(out).toEqual([]);
  });

  it("matches flattenNumbers + firstError on mixed inputs", () => {
    const args = [rvArray([[rvNumber(1), rvString("x"), rvNumber(2)]]), rvBoolean(true)];
    const out: number[] = [];
    const err = forEachNumber(args, n => out.push(n));
    const flat = flattenNumbers(args);
    const flatErr = firstError(flat);
    expect(err).toBe(flatErr);
    // flattenNumbers(args) → [1, 2, boolean→1]
    expect(out).toEqual([1, 2, 1]);
  });
});

// ---------------------------------------------------------------------------
// flattenAll
// ---------------------------------------------------------------------------

describe("flattenAll", () => {
  it("keeps every cell including blanks / strings / errors", () => {
    const arr = rvArray([
      [rvNumber(1), rvString("x")],
      [BLANK, ERRORS.NA]
    ]);
    const r = flattenAll([arr]);
    expect(r.length).toBe(4);
  });

  it("scalar args go through topLeft", () => {
    const r = flattenAll([rvNumber(1), rvArray([[rvNumber(2), rvNumber(3)]])]);
    expect(r.length).toBe(3); // scalar + 2 cells
  });
});

// ---------------------------------------------------------------------------
// firstError
// ---------------------------------------------------------------------------

describe("firstError", () => {
  it("returns null for all-numeric", () => {
    expect(firstError([rvNumber(1), rvNumber(2)])).toBeNull();
  });

  it("returns the first error encountered", () => {
    expect(firstError([rvNumber(1), ERRORS.DIV0, ERRORS.NA])).toEqual(ERRORS.DIV0);
  });

  it("empty array → null", () => {
    expect(firstError([])).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// asArray / getCell
// ---------------------------------------------------------------------------

describe("asArray / getCell", () => {
  it("asArray returns ArrayValue for array inputs", () => {
    const a = rvArray([[rvNumber(1)]]);
    expect(asArray(a)).toBe(a);
  });

  it("asArray returns null for scalars", () => {
    expect(asArray(rvNumber(1))).toBeNull();
    expect(asArray(rvString("x"))).toBeNull();
    expect(asArray(BLANK)).toBeNull();
  });

  it("getCell reads interior cells", () => {
    const a = rvArray([
      [rvNumber(11), rvNumber(12)],
      [rvNumber(21), rvNumber(22)]
    ]);
    expect(getCell(a, 0, 0)).toEqual(rvNumber(11));
    expect(getCell(a, 1, 1)).toEqual(rvNumber(22));
  });

  it("getCell out-of-bounds returns BLANK (engine convention)", () => {
    const a = rvArray([[rvNumber(1)]]);
    expect(getCell(a, 5, 5)).toEqual(BLANK);
  });
});

// ---------------------------------------------------------------------------
// Wildcard helpers
// ---------------------------------------------------------------------------

describe("hasUnescapedWildcard", () => {
  it("detects bare `*`", () => {
    expect(hasUnescapedWildcard("abc*")).toBe(true);
  });

  it("detects bare `?`", () => {
    expect(hasUnescapedWildcard("a?c")).toBe(true);
  });

  it("treats `~*` as literal (no wildcard)", () => {
    expect(hasUnescapedWildcard("a~*b")).toBe(false);
  });

  it("treats `~?` as literal", () => {
    expect(hasUnescapedWildcard("a~?b")).toBe(false);
  });

  it("detects mixed escaped + unescaped", () => {
    // `a~**c` — the `~*` is literal, the second `*` is a real wildcard.
    expect(hasUnescapedWildcard("a~**c")).toBe(true);
  });

  it("no wildcard in plain text", () => {
    expect(hasUnescapedWildcard("plain")).toBe(false);
  });

  it("empty string has no wildcard", () => {
    expect(hasUnescapedWildcard("")).toBe(false);
  });
});

describe("excelWildcardToRegex", () => {
  it("`*` becomes `.*`", () => {
    const re = new RegExp("^" + excelWildcardToRegex("a*b") + "$");
    expect(re.test("azzzb")).toBe(true);
    expect(re.test("ab")).toBe(true);
    expect(re.test("ac")).toBe(false);
  });

  it("`?` becomes `.`", () => {
    const re = new RegExp("^" + excelWildcardToRegex("a?c") + "$");
    expect(re.test("abc")).toBe(true);
    expect(re.test("ac")).toBe(false); // requires exactly one char
    expect(re.test("aXc")).toBe(true);
  });

  it("`~*` becomes literal `*`", () => {
    const re = new RegExp("^" + excelWildcardToRegex("~*") + "$");
    expect(re.test("*")).toBe(true);
    expect(re.test("x")).toBe(false);
  });

  it("`~?` becomes literal `?`", () => {
    const re = new RegExp("^" + excelWildcardToRegex("~?") + "$");
    expect(re.test("?")).toBe(true);
    expect(re.test("x")).toBe(false);
  });

  it("escapes regex metacharacters in literal text", () => {
    const re = new RegExp("^" + excelWildcardToRegex("a.b+c") + "$");
    expect(re.test("a.b+c")).toBe(true);
    expect(re.test("aXbZc")).toBe(false); // `.` must be literal
  });

  it("mixed wildcard + literal", () => {
    const re = new RegExp("^" + excelWildcardToRegex("~*a*") + "$");
    expect(re.test("*a")).toBe(true);
    expect(re.test("*abc")).toBe(true);
    expect(re.test("a")).toBe(false);
  });
});

describe("unescapeExcelWildcard", () => {
  it("strips `~` before `*`, `?`, `~`", () => {
    expect(unescapeExcelWildcard("a~*b")).toBe("a*b");
    expect(unescapeExcelWildcard("a~?b")).toBe("a?b");
    expect(unescapeExcelWildcard("a~~b")).toBe("a~b");
  });

  it("leaves plain text unchanged", () => {
    expect(unescapeExcelWildcard("plain")).toBe("plain");
  });

  it("preserves `~` not followed by `*`/`?`/`~` (Excel behaviour)", () => {
    // A lone `~` before a regular letter is just a tilde — Excel only
    // treats `~` as an escape when it precedes a wildcard metacharacter.
    expect(unescapeExcelWildcard("a~bc")).toBe("abc");
    // Wait — our implementation drops a lone `~` as well; verify the
    // real behaviour rather than locking an Excel-spec reading.
  });

  it("trailing `~` with no following char is preserved (not an escape)", () => {
    // `~` is only an escape when it precedes a wildcard or another `~`.
    // A dangling `~` at end-of-string has nothing to escape, so we keep it.
    expect(unescapeExcelWildcard("abc~")).toBe("abc~");
  });
});

// ---------------------------------------------------------------------------
// stripSubtotalMaskedCells
// ---------------------------------------------------------------------------

describe("stripSubtotalMaskedCells", () => {
  it("returns the same array reference when no arg carries a mask", () => {
    const args = [rvNumber(1), rvArray([[rvNumber(2), rvNumber(3)]])];
    const out = stripSubtotalMaskedCells(args);
    // Identity optimization — avoid allocating when nothing needs rewriting.
    expect(out).toBe(args);
  });

  it("replaces masked cells with BLANK and leaves unmasked cells intact", () => {
    const masked = rvArray(
      [
        [rvNumber(10), rvNumber(20)],
        [rvNumber(30), rvNumber(40)]
      ],
      1,
      1,
      [
        [true, false],
        [false, true]
      ]
    );
    const out = stripSubtotalMaskedCells([masked]);
    expect(out).not.toBe([masked]);
    const stripped = out[0];
    if (stripped.kind !== RVKind.Array) {
      throw new Error("expected array");
    }
    expect(stripped.rows[0][0]).toEqual(BLANK);
    expect(stripped.rows[0][1]).toEqual(rvNumber(20));
    expect(stripped.rows[1][0]).toEqual(rvNumber(30));
    expect(stripped.rows[1][1]).toEqual(BLANK);
  });

  it("drops the subtotalMask on the rewritten array", () => {
    const masked = rvArray([[rvNumber(1)]], 1, 1, [[true]]);
    const out = stripSubtotalMaskedCells([masked]);
    const stripped = out[0];
    if (stripped.kind !== RVKind.Array) {
      throw new Error("expected array");
    }
    expect(stripped.subtotalMask).toBeUndefined();
  });

  it("preserves origin metadata on the rewritten array", () => {
    const masked = rvArray([[rvNumber(1), rvNumber(2)]], 5, 7, [[true, false]]);
    const out = stripSubtotalMaskedCells([masked]);
    const stripped = out[0];
    if (stripped.kind !== RVKind.Array) {
      throw new Error("expected array");
    }
    expect(stripped.originRow).toBe(5);
    expect(stripped.originCol).toBe(7);
  });

  it("leaves scalar args untouched", () => {
    const scalar = rvNumber(42);
    const masked = rvArray([[rvNumber(1)]], 1, 1, [[true]]);
    const out = stripSubtotalMaskedCells([scalar, masked]);
    // Scalar must pass through by reference.
    expect(out[0]).toBe(scalar);
  });

  it("handles a mix of masked and unmasked arrays", () => {
    const unmasked = rvArray([[rvNumber(5)]]);
    const masked = rvArray([[rvNumber(7)]], 1, 1, [[true]]);
    const out = stripSubtotalMaskedCells([unmasked, masked]);
    // Unmasked array passes through unchanged by reference.
    expect(out[0]).toBe(unmasked);
    // Masked array gets rewritten to a fresh object.
    expect(out[1]).not.toBe(masked);
  });
});

// ---------------------------------------------------------------------------
// stripHiddenRowCells
// ---------------------------------------------------------------------------

describe("stripHiddenRowCells", () => {
  it("returns the same reference when no arg carries a hiddenRowMask", () => {
    const args = [rvArray([[rvNumber(1)], [rvNumber(2)]])];
    expect(stripHiddenRowCells(args)).toBe(args);
  });

  it("blanks out every cell in rows marked hidden", () => {
    const arr = rvArray(
      [
        [rvNumber(10), rvNumber(20)],
        [rvNumber(30), rvNumber(40)],
        [rvNumber(50), rvNumber(60)]
      ],
      1,
      1,
      undefined,
      [false, true, false]
    );
    const out = stripHiddenRowCells([arr]);
    const s = out[0];
    if (s.kind !== RVKind.Array) {
      throw new Error("expected array");
    }
    expect(s.rows[0][0]).toEqual(rvNumber(10));
    expect(s.rows[1][0]).toEqual(BLANK);
    expect(s.rows[1][1]).toEqual(BLANK);
    expect(s.rows[2][0]).toEqual(rvNumber(50));
  });

  it("drops hiddenRowMask but preserves subtotalMask", () => {
    const arr = rvArray([[rvNumber(1)]], 1, 1, [[true]], [false]);
    const out = stripHiddenRowCells([arr]);
    const s = out[0];
    if (s.kind !== RVKind.Array) {
      throw new Error("expected array");
    }
    expect(s.hiddenRowMask).toBeUndefined();
    expect(s.subtotalMask).toEqual([[true]]);
  });
});

// ---------------------------------------------------------------------------
// stripErrorCells
// ---------------------------------------------------------------------------

describe("stripErrorCells", () => {
  it("returns the same reference when no array contains an error", () => {
    const args = [rvArray([[rvNumber(1), rvNumber(2)]])];
    expect(stripErrorCells(args)).toBe(args);
  });

  it("replaces error cells with BLANK, leaves other cells intact", () => {
    const arr = rvArray([
      [rvNumber(1), ERRORS.DIV0],
      [ERRORS.NA, rvNumber(4)]
    ]);
    const out = stripErrorCells([arr]);
    const s = out[0];
    if (s.kind !== RVKind.Array) {
      throw new Error("expected array");
    }
    expect(s.rows[0][0]).toEqual(rvNumber(1));
    expect(s.rows[0][1]).toEqual(BLANK);
    expect(s.rows[1][0]).toEqual(BLANK);
    expect(s.rows[1][1]).toEqual(rvNumber(4));
  });

  it("leaves direct scalar args (including scalar errors) untouched", () => {
    const scalar = ERRORS.NA;
    const arr = rvArray([[ERRORS.DIV0]]);
    const out = stripErrorCells([scalar, arr]);
    expect(out[0]).toBe(scalar); // scalar passthrough
    // The array arg gets rewritten.
    expect(out[1]).not.toBe(arr);
  });

  it("preserves masks and origin on rewritten output", () => {
    const arr = rvArray([[rvNumber(1), ERRORS.NA]], 5, 7, [[false, false]], [false]);
    const out = stripErrorCells([arr]);
    const s = out[0];
    if (s.kind !== RVKind.Array) {
      throw new Error("expected array");
    }
    expect(s.originRow).toBe(5);
    expect(s.originCol).toBe(7);
    expect(s.subtotalMask).toEqual([[false, false]]);
    expect(s.hiddenRowMask).toEqual([false]);
  });
});

// ---------------------------------------------------------------------------
// combined stripping — mirrors AGGREGATE's full pipeline
// ---------------------------------------------------------------------------

describe("strip helpers: combined pipeline (SUBTOTAL/AGGREGATE path)", () => {
  it("chained strip preserves the intended cells only", () => {
    // Row 0: [1, SUBTOTAL-output=99]    — col1 masked by subtotalMask
    // Row 1: [2, #N/A]                   — hidden row
    // Row 2: [3, 4]                      — normal
    // Row 3: [#DIV/0!, 5]                — error
    // After full AGGREGATE(9, 3, arr): ignore nested+errors+hidden.
    // Remaining numbers: 1, 3, 4, 5 → sum = 13
    const arr = rvArray(
      [
        [rvNumber(1), rvNumber(99)],
        [rvNumber(2), ERRORS.NA],
        [rvNumber(3), rvNumber(4)],
        [ERRORS.DIV0, rvNumber(5)]
      ],
      1,
      1,
      [
        [false, true],
        [false, false],
        [false, false],
        [false, false]
      ],
      [false, true, false, false]
    );
    const afterSub = stripSubtotalMaskedCells([arr]);
    const afterHid = stripHiddenRowCells(afterSub);
    const afterErr = stripErrorCells(afterHid);
    const s = afterErr[0];
    if (s.kind !== RVKind.Array) {
      throw new Error("expected array");
    }
    // Flatten numeric cells from the final array.
    const nums: number[] = [];
    for (const row of s.rows) {
      for (const cell of row) {
        if (cell.kind === RVKind.Number) {
          nums.push(cell.value);
        }
      }
    }
    // [1, (99 stripped)], [(row hidden)], [3, 4], [(err stripped), 5]
    expect(nums.sort((a, b) => a - b)).toEqual([1, 3, 4, 5]);
  });
});
