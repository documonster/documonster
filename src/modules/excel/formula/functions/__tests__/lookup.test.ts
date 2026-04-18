/**
 * Unit tests for lookup / reference functions in `../lookup.ts`.
 *
 * The reference-aware branch of ROW/COLUMN lives in the evaluator
 * (`tryEvaluateRefFunction`) — direct calls to `fnROW` / `fnCOLUMN` with a
 * non-reference argument always return `#VALUE!`, and we test that baseline
 * behaviour here. End-to-end tests through a Workbook (covering
 * `ROW(INDIRECT("A5"))` and `COLUMN(OFFSET(A1, 0, 3))`) are split into
 * `evaluator-integration.test.ts` so this file remains purely mechanical.
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
  type ArrayValue,
  type NumberValue,
  type RuntimeValue,
  type StringValue
} from "../../runtime/values";
import {
  fnROW,
  fnCOLUMN,
  fnROWS,
  fnCOLUMNS,
  fnINDEX,
  fnMATCH,
  fnVLOOKUP,
  fnHLOOKUP,
  fnXLOOKUP,
  fnXMATCH,
  fnADDRESS,
  fnLOOKUP,
  fnTRANSPOSE,
  fnAREAS
} from "../lookup";

function asNumber(v: RuntimeValue): number {
  expect(v.kind).toBe(RVKind.Number);
  return (v as NumberValue).value;
}

function asString(v: RuntimeValue): string {
  expect(v.kind).toBe(RVKind.String);
  return (v as StringValue).value;
}

describe("ROW / COLUMN (non-reference fallback)", () => {
  it("fnROW with a non-reference argument returns #VALUE!", () => {
    expect(fnROW([rvArray([[rvNumber(1)]])])).toEqual(ERRORS.VALUE);
    expect(fnROW([rvNumber(5)])).toEqual(ERRORS.VALUE);
  });

  it("fnCOLUMN with a non-reference argument returns #VALUE!", () => {
    expect(fnCOLUMN([rvArray([[rvNumber(1)]])])).toEqual(ERRORS.VALUE);
  });
});

describe("ROWS / COLUMNS", () => {
  it("ROWS counts the height of an array", () => {
    const a = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)],
      [rvNumber(5), rvNumber(6)]
    ]);
    expect(asNumber(fnROWS([a]))).toBe(3);
  });

  it("COLUMNS counts the width of an array", () => {
    const a = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    expect(asNumber(fnCOLUMNS([a]))).toBe(3);
  });

  it("ROWS / COLUMNS return 1 for a scalar", () => {
    expect(asNumber(fnROWS([rvNumber(1)]))).toBe(1);
    expect(asNumber(fnCOLUMNS([rvString("x")]))).toBe(1);
  });
});

describe("INDEX", () => {
  const a = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvNumber(4), rvNumber(5), rvNumber(6)],
    [rvNumber(7), rvNumber(8), rvNumber(9)]
  ]);

  it("INDEX(a, 2, 3) returns the cell at row 2, column 3", () => {
    expect(asNumber(fnINDEX([a, rvNumber(2), rvNumber(3)]))).toBe(6);
  });

  it("INDEX truncates fractional indices toward zero (regression)", () => {
    // INDEX(a, 1.9, 2.5) should match INDEX(a, 1, 2) → value 2.
    expect(asNumber(fnINDEX([a, rvNumber(1.9), rvNumber(2.5)]))).toBe(2);
  });

  it("INDEX(a, 0, col) returns the whole column as an array", () => {
    const r = fnINDEX([a, rvNumber(0), rvNumber(1)]);
    expect(r.kind).toBe(RVKind.Array);
    const arr = r as ArrayValue;
    expect(arr.height).toBe(3);
    expect(arr.width).toBe(1);
    expect((arr.rows[0][0] as NumberValue).value).toBe(1);
    expect((arr.rows[2][0] as NumberValue).value).toBe(7);
  });

  it("INDEX(a, row, 0) returns the whole row as an array", () => {
    const r = fnINDEX([a, rvNumber(2), rvNumber(0)]);
    expect(r.kind).toBe(RVKind.Array);
    const arr = r as ArrayValue;
    expect(arr.height).toBe(1);
    expect(arr.width).toBe(3);
    expect((arr.rows[0][0] as NumberValue).value).toBe(4);
  });

  it("INDEX rejects out-of-bounds indices with #REF!", () => {
    expect(fnINDEX([a, rvNumber(5), rvNumber(1)])).toEqual(ERRORS.REF);
    expect(fnINDEX([a, rvNumber(1), rvNumber(5)])).toEqual(ERRORS.REF);
  });

  it("INDEX rejects negative indices with #VALUE!", () => {
    expect(fnINDEX([a, rvNumber(-1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });
});

describe("MATCH", () => {
  it("matchType 0 — exact match", () => {
    const arr = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
    expect(asNumber(fnMATCH([rvNumber(20), arr, rvNumber(0)]))).toBe(2);
  });

  it("matchType 0 — string match with wildcard", () => {
    const arr = rvArray([[rvString("apple"), rvString("banana"), rvString("cherry")]]);
    expect(asNumber(fnMATCH([rvString("ban*"), arr, rvNumber(0)]))).toBe(2);
  });

  it("matchType 1 — largest ≤ lookup (sorted ascending)", () => {
    const arr = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30), rvNumber(40)]]);
    expect(asNumber(fnMATCH([rvNumber(25), arr, rvNumber(1)]))).toBe(2);
  });

  it("matchType -1 — smallest ≥ lookup (sorted descending)", () => {
    const arr = rvArray([[rvNumber(40), rvNumber(30), rvNumber(20), rvNumber(10)]]);
    expect(asNumber(fnMATCH([rvNumber(25), arr, rvNumber(-1)]))).toBe(2);
  });

  it("returns #N/A when not found", () => {
    const arr = rvArray([[rvNumber(10), rvNumber(20)]]);
    expect(fnMATCH([rvNumber(99), arr, rvNumber(0)])).toEqual(ERRORS.NA);
  });
});

describe("VLOOKUP", () => {
  const table = rvArray([
    [rvNumber(1), rvString("a"), rvString("x")],
    [rvNumber(2), rvString("b"), rvString("y")],
    [rvNumber(3), rvString("c"), rvString("z")]
  ]);

  it("exact match returns the matching row's column", () => {
    expect(asString(fnVLOOKUP([rvNumber(2), table, rvNumber(2), rvBoolean(false)]))).toBe("b");
    expect(asString(fnVLOOKUP([rvNumber(3), table, rvNumber(3), rvBoolean(false)]))).toBe("z");
  });

  it("truncates fractional column index (regression)", () => {
    // VLOOKUP(x, t, 1.5, FALSE) should behave like col_index = 1
    expect(asNumber(fnVLOOKUP([rvNumber(2), table, rvNumber(1.5), rvBoolean(false)]))).toBe(2);
    expect(asNumber(fnVLOOKUP([rvNumber(2), table, rvNumber(1), rvBoolean(false)]))).toBe(2);
  });

  it("returns #N/A for exact match miss", () => {
    expect(fnVLOOKUP([rvNumber(99), table, rvNumber(2), rvBoolean(false)])).toEqual(ERRORS.NA);
  });

  it("approximate match finds largest ≤ lookup", () => {
    expect(asString(fnVLOOKUP([rvNumber(2.5), table, rvNumber(2), rvBoolean(true)]))).toBe("b");
  });

  it("returns #REF! on out-of-range column index", () => {
    expect(fnVLOOKUP([rvNumber(1), table, rvNumber(99), rvBoolean(false)])).toEqual(ERRORS.REF);
    expect(fnVLOOKUP([rvNumber(1), table, rvNumber(0), rvBoolean(false)])).toEqual(ERRORS.REF);
  });
});

describe("HLOOKUP", () => {
  const table = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvString("a"), rvString("b"), rvString("c")],
    [rvString("x"), rvString("y"), rvString("z")]
  ]);

  it("exact match", () => {
    expect(asString(fnHLOOKUP([rvNumber(2), table, rvNumber(2), rvBoolean(false)]))).toBe("b");
  });

  it("#N/A on miss (exact)", () => {
    expect(fnHLOOKUP([rvNumber(99), table, rvNumber(2), rvBoolean(false)])).toEqual(ERRORS.NA);
  });

  it("#REF! on out-of-range row", () => {
    expect(fnHLOOKUP([rvNumber(1), table, rvNumber(99), rvBoolean(false)])).toEqual(ERRORS.REF);
  });
});

describe("XLOOKUP", () => {
  it("vertical lookup — single return column", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    expect(asString(fnXLOOKUP([rvNumber(2), lookup, ret]))).toBe("b");
  });

  it("vertical lookup — multi-column return yields a row vector", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    const ret = rvArray([
      [rvString("a"), rvString("x")],
      [rvString("b"), rvString("y")]
    ]);
    const r = fnXLOOKUP([rvNumber(2), lookup, ret]);
    expect(r.kind).toBe(RVKind.Array);
    const arr = r as ArrayValue;
    expect(arr.height).toBe(1);
    expect(arr.width).toBe(2);
    expect((arr.rows[0][0] as StringValue).value).toBe("b");
    expect((arr.rows[0][1] as StringValue).value).toBe("y");
  });

  it("horizontal lookup with multi-row return yields a column vector (regression)", () => {
    // lookup_array is a 1xN row; return_array is MxN — each column is a
    // column in the return.
    const lookup = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const ret = rvArray([
      [rvString("a"), rvString("b"), rvString("c")],
      [rvString("x"), rvString("y"), rvString("z")]
    ]);
    const r = fnXLOOKUP([rvNumber(2), lookup, ret]);
    expect(r.kind).toBe(RVKind.Array);
    const arr = r as ArrayValue;
    // Found index = 1 → column 1 of return. The output must be a COLUMN
    // (2 rows × 1 col), NOT a row — this was the regression.
    expect(arr.height).toBe(2);
    expect(arr.width).toBe(1);
    expect((arr.rows[0][0] as StringValue).value).toBe("b");
    expect((arr.rows[1][0] as StringValue).value).toBe("y");
  });

  it("returns if_not_found when provided", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")]]);
    expect(asString(fnXLOOKUP([rvNumber(99), lookup, ret, rvString("missing")]))).toBe("missing");
  });

  it("returns #N/A on miss when no if_not_found", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")]]);
    expect(fnXLOOKUP([rvNumber(99), lookup, ret])).toEqual(ERRORS.NA);
  });

  it("matchMode -1 — next smaller", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(5)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    // Lookup 4, next smaller = 3 → "b"
    expect(asString(fnXLOOKUP([rvNumber(4), lookup, ret, BLANK, rvNumber(-1)]))).toBe("b");
  });

  it("matchMode 2 — wildcard", () => {
    const lookup = rvArray([[rvString("apple")], [rvString("banana")], [rvString("cherry")]]);
    const ret = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    expect(asNumber(fnXLOOKUP([rvString("ban*"), lookup, ret, BLANK, rvNumber(2)]))).toBe(2);
  });
});

describe("XMATCH", () => {
  it("exact match default", () => {
    const arr = rvArray([[rvNumber(10)], [rvNumber(20)], [rvNumber(30)]]);
    expect(asNumber(fnXMATCH([rvNumber(20), arr]))).toBe(2);
  });

  it("matchMode -1 — next smaller", () => {
    const arr = rvArray([[rvNumber(10)], [rvNumber(20)], [rvNumber(30)]]);
    expect(asNumber(fnXMATCH([rvNumber(25), arr, rvNumber(-1)]))).toBe(2);
  });

  it("matchMode 1 — next larger", () => {
    const arr = rvArray([[rvNumber(10)], [rvNumber(20)], [rvNumber(30)]]);
    expect(asNumber(fnXMATCH([rvNumber(25), arr, rvNumber(1)]))).toBe(3);
  });

  it("not-found returns #N/A", () => {
    const arr = rvArray([[rvNumber(10)]]);
    expect(fnXMATCH([rvNumber(99), arr])).toEqual(ERRORS.NA);
  });
});

describe("ADDRESS", () => {
  it("default produces $A$1-style absolute reference", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(1)]))).toBe("$A$1");
    expect(asString(fnADDRESS([rvNumber(10), rvNumber(3)]))).toBe("$C$10");
  });

  it("abs_num=4 — relative", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(4)]))).toBe("A1");
  });

  it("abs_num=2 — col relative, row absolute", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(2)]))).toBe("A$1");
  });

  it("handles columns beyond Z", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(27)]))).toBe("$AA$1");
  });

  it("R1C1 style when a1=false", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(1), rvBoolean(false)]))).toBe(
      "R1C1"
    );
  });
});

describe("LOOKUP", () => {
  it("vector form", () => {
    const lookup = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const ret = rvArray([[rvString("a"), rvString("b"), rvString("c")]]);
    expect(asString(fnLOOKUP([rvNumber(2), lookup, ret]))).toBe("b");
  });

  it("approximate match", () => {
    const lookup = rvArray([[rvNumber(1), rvNumber(3), rvNumber(5)]]);
    const ret = rvArray([[rvString("a"), rvString("b"), rvString("c")]]);
    // Looking up 4 — largest ≤ 4 is 3 → "b"
    expect(asString(fnLOOKUP([rvNumber(4), lookup, ret]))).toBe("b");
  });
});

describe("TRANSPOSE", () => {
  it("swaps rows and columns", () => {
    const a = rvArray([
      [rvNumber(1), rvNumber(2), rvNumber(3)],
      [rvNumber(4), rvNumber(5), rvNumber(6)]
    ]);
    const r = fnTRANSPOSE([a]);
    expect(r.kind).toBe(RVKind.Array);
    const arr = r as ArrayValue;
    expect(arr.height).toBe(3);
    expect(arr.width).toBe(2);
    expect((arr.rows[0][0] as NumberValue).value).toBe(1);
    expect((arr.rows[0][1] as NumberValue).value).toBe(4);
    expect((arr.rows[2][1] as NumberValue).value).toBe(6);
  });

  it("wraps a scalar in a 1x1 array", () => {
    const r = fnTRANSPOSE([rvNumber(42)]);
    expect(r.kind).toBe(RVKind.Array);
    const arr = r as ArrayValue;
    expect(arr.height).toBe(1);
    expect(arr.width).toBe(1);
  });
});

describe("AREAS", () => {
  it("returns 1 for any single argument", () => {
    expect(asNumber(fnAREAS([rvNumber(1)]))).toBe(1);
  });

  it("returns #VALUE! for no arguments", () => {
    expect(fnAREAS([])).toEqual(ERRORS.VALUE);
  });
});

// ============================================================================
// Comprehensive function-level tests — append-only.
//
// These blocks fill the task checklist: normal inputs, shape edges (scalar
// vs 1×N vs M×1 vs M×N), error propagation, out-of-range indices, and the
// specific Excel behaviours the checklist flags (ADDRESS abs_num > 4,
// VLOOKUP fractional col_index_num truncation, XLOOKUP binary search etc.).
//
// Note: ROW / COLUMN / INDIRECT / OFFSET need an EvalContext to do any
// useful work and are not directly exported as `fnROW`/`fnCOLUMN` in a
// reference-aware form. We only exercise the direct-call fallback (→ #VALUE!)
// and leave end-to-end tests to `evaluator-integration.test.ts`.
// ============================================================================

describe("ROW / COLUMN comprehensive", () => {
  it("ROW with no args falls through to #VALUE! (needs ctx)", () => {
    expect(fnROW([])).toEqual(ERRORS.VALUE);
  });

  it("ROW on a scalar is #VALUE!", () => {
    expect(fnROW([rvNumber(5)])).toEqual(ERRORS.VALUE);
    expect(fnROW([rvString("x")])).toEqual(ERRORS.VALUE);
  });

  it("COLUMN with no args / scalar / string / array all return #VALUE!", () => {
    expect(fnCOLUMN([])).toEqual(ERRORS.VALUE);
    expect(fnCOLUMN([rvNumber(5)])).toEqual(ERRORS.VALUE);
    expect(fnCOLUMN([rvString("hi")])).toEqual(ERRORS.VALUE);
    expect(fnCOLUMN([rvArray([[rvNumber(1)]])])).toEqual(ERRORS.VALUE);
  });

  it("ROW/COLUMN treat errors the same — still #VALUE! (not propagated)", () => {
    expect(fnROW([ERRORS.NA])).toEqual(ERRORS.VALUE);
    expect(fnCOLUMN([ERRORS.NA])).toEqual(ERRORS.VALUE);
  });
});

describe("ROWS / COLUMNS comprehensive", () => {
  it("ROWS on a 1×N row is 1", () => {
    expect(asNumber(fnROWS([rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]])]))).toBe(1);
  });

  it("ROWS on an M×1 column counts the height", () => {
    expect(asNumber(fnROWS([rvArray([[rvNumber(1)], [rvNumber(2)]])]))).toBe(2);
  });

  it("COLUMNS on M×1 is 1", () => {
    expect(asNumber(fnCOLUMNS([rvArray([[rvNumber(1)], [rvNumber(2)]])]))).toBe(1);
  });

  it("ROWS / COLUMNS on a blank still return 1 (scalar shape)", () => {
    expect(asNumber(fnROWS([BLANK]))).toBe(1);
    expect(asNumber(fnCOLUMNS([BLANK]))).toBe(1);
  });

  it("ROWS / COLUMNS on an error treat it as a scalar (return 1 — NOT error-propagating)", () => {
    expect(asNumber(fnROWS([ERRORS.NA]))).toBe(1);
    expect(asNumber(fnCOLUMNS([ERRORS.NA]))).toBe(1);
  });
});

describe("INDEX comprehensive", () => {
  const a = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvNumber(4), rvNumber(5), rvNumber(6)],
    [rvNumber(7), rvNumber(8), rvNumber(9)]
  ]);

  it("INDEX with 1 arg returns the full array", () => {
    const r = fnINDEX([a]);
    expect(r.kind).toBe(RVKind.Array);
    const arr = r as ArrayValue;
    expect(arr.height).toBe(3);
    expect(arr.width).toBe(3);
  });

  it("INDEX with row only (2 args) returns the whole row", () => {
    const r = fnINDEX([a, rvNumber(2)]) as ArrayValue;
    expect(r.kind).toBe(RVKind.Array);
    expect(r.height).toBe(1);
    expect(r.width).toBe(3);
    expect((r.rows[0][0] as NumberValue).value).toBe(4);
  });

  it("INDEX on a scalar returns the scalar", () => {
    expect(asNumber(fnINDEX([rvNumber(42), rvNumber(1), rvNumber(1)]))).toBe(42);
  });

  it("INDEX row=0 col=0 returns the full array", () => {
    const r = fnINDEX([a, rvNumber(0), rvNumber(0)]);
    expect(r.kind).toBe(RVKind.Array);
  });

  it("INDEX row=0 with out-of-range col → #REF!", () => {
    expect(fnINDEX([a, rvNumber(0), rvNumber(99)])).toEqual(ERRORS.REF);
  });

  it("INDEX col=0 with out-of-range row → #REF!", () => {
    expect(fnINDEX([a, rvNumber(99), rvNumber(0)])).toEqual(ERRORS.REF);
  });

  it("INDEX propagates error from row_num / col_num argument", () => {
    expect(fnINDEX([a, ERRORS.NA, rvNumber(1)])).toEqual(ERRORS.NA);
    expect(fnINDEX([a, rvNumber(1), ERRORS.REF])).toEqual(ERRORS.REF);
  });

  it("INDEX truncates fractional indices toward zero (0.5 → 0 → full column)", () => {
    const r = fnINDEX([a, rvNumber(0.5), rvNumber(1)]);
    // 0.5 truncates to 0 → whole column
    expect(r.kind).toBe(RVKind.Array);
    const arr = r as ArrayValue;
    expect(arr.height).toBe(3);
    expect(arr.width).toBe(1);
  });
});

describe("MATCH comprehensive", () => {
  it("propagates error from lookup value", () => {
    const arr = rvArray([[rvNumber(1), rvNumber(2)]]);
    expect(fnMATCH([ERRORS.NA, arr, rvNumber(0)])).toEqual(ERRORS.NA);
  });

  it("non-array lookup range → #N/A", () => {
    expect(fnMATCH([rvNumber(1), rvNumber(1), rvNumber(0)])).toEqual(ERRORS.NA);
  });

  it("matchType 0 on a column (M×1) works", () => {
    const arr = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    expect(asNumber(fnMATCH([rvString("b"), arr, rvNumber(0)]))).toBe(2);
  });

  it("matchType 0 is case-insensitive for strings", () => {
    const arr = rvArray([[rvString("Banana")]]);
    expect(asNumber(fnMATCH([rvString("banana"), arr, rvNumber(0)]))).toBe(1);
  });

  it("matchType 1 on sorted strings finds largest ≤ lookup (case-insensitive)", () => {
    const arr = rvArray([[rvString("apple"), rvString("mango"), rvString("zebra")]]);
    expect(asNumber(fnMATCH([rvString("orange"), arr, rvNumber(1)]))).toBe(2);
  });

  it("matchType default (no arg) = 1", () => {
    const arr = rvArray([[rvNumber(10), rvNumber(20), rvNumber(30)]]);
    expect(asNumber(fnMATCH([rvNumber(25), arr]))).toBe(2);
  });

  it("matchType 0 wildcard * matches anything", () => {
    const arr = rvArray([[rvString("apple"), rvString("banana"), rvString("cherry")]]);
    expect(asNumber(fnMATCH([rvString("c*"), arr, rvNumber(0)]))).toBe(3);
  });

  it("matchType 0 returns #N/A if no match", () => {
    const arr = rvArray([[rvNumber(10), rvNumber(20)]]);
    expect(fnMATCH([rvNumber(99), arr, rvNumber(0)])).toEqual(ERRORS.NA);
  });
});

describe("VLOOKUP comprehensive", () => {
  const table = rvArray([
    [rvNumber(1), rvString("a"), rvString("x")],
    [rvNumber(2), rvString("b"), rvString("y")],
    [rvNumber(3), rvString("c"), rvString("z")]
  ]);

  it("propagates error from lookup value", () => {
    expect(fnVLOOKUP([ERRORS.NA, table, rvNumber(2), rvBoolean(false)])).toEqual(ERRORS.NA);
  });

  it("non-array table_array → #N/A", () => {
    expect(fnVLOOKUP([rvNumber(1), rvNumber(1), rvNumber(1), rvBoolean(false)])).toEqual(ERRORS.NA);
  });

  it("col_index_num = 0 → #REF!", () => {
    expect(fnVLOOKUP([rvNumber(1), table, rvNumber(0), rvBoolean(false)])).toEqual(ERRORS.REF);
  });

  it("col_index_num fractional truncation (2.9 → 2)", () => {
    expect(asString(fnVLOOKUP([rvNumber(2), table, rvNumber(2.9), rvBoolean(false)]))).toBe("b");
  });

  it("rangeLookup defaults to TRUE (approximate)", () => {
    // no 4th arg → approx search for 2.5 → finds 2 → "b"
    expect(asString(fnVLOOKUP([rvNumber(2.5), table, rvNumber(2)]))).toBe("b");
  });

  it("approximate with value below all rows → #N/A", () => {
    expect(fnVLOOKUP([rvNumber(-5), table, rvNumber(2), rvBoolean(true)])).toEqual(ERRORS.NA);
  });

  it("exact string match case-insensitive", () => {
    const t = rvArray([
      [rvString("Banana"), rvNumber(100)],
      [rvString("Apple"), rvNumber(200)]
    ]);
    expect(asNumber(fnVLOOKUP([rvString("banana"), t, rvNumber(2), rvBoolean(false)]))).toBe(100);
  });
});

describe("HLOOKUP comprehensive", () => {
  const table = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvString("a"), rvString("b"), rvString("c")]
  ]);

  it("approximate match on 1st row", () => {
    expect(asString(fnHLOOKUP([rvNumber(2.5), table, rvNumber(2), rvBoolean(true)]))).toBe("b");
  });

  it("row_index_num truncated (2.9 → 2)", () => {
    expect(asString(fnHLOOKUP([rvNumber(1), table, rvNumber(2.9), rvBoolean(false)]))).toBe("a");
  });

  it("error from lookup value propagates", () => {
    expect(fnHLOOKUP([ERRORS.NA, table, rvNumber(2), rvBoolean(false)])).toEqual(ERRORS.NA);
  });

  it("non-array table → #N/A", () => {
    expect(fnHLOOKUP([rvNumber(1), rvNumber(1), rvNumber(1), rvBoolean(false)])).toEqual(ERRORS.NA);
  });

  it("row_index_num = 0 → #REF!", () => {
    expect(fnHLOOKUP([rvNumber(1), table, rvNumber(0), rvBoolean(false)])).toEqual(ERRORS.REF);
  });

  it("rangeLookup defaults to TRUE", () => {
    expect(asString(fnHLOOKUP([rvNumber(2.5), table, rvNumber(2)]))).toBe("b");
  });
});

describe("XLOOKUP comprehensive", () => {
  it("non-array lookup_array → #VALUE!", () => {
    expect(fnXLOOKUP([rvNumber(1), rvNumber(1), rvArray([[rvString("x")]])])).toEqual(ERRORS.VALUE);
  });

  it("non-array return_array → #VALUE!", () => {
    expect(fnXLOOKUP([rvNumber(1), rvArray([[rvNumber(1)]]), rvNumber(5)])).toEqual(ERRORS.VALUE);
  });

  it("searchMode -1 reverse scan finds last match", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(1)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    expect(asString(fnXLOOKUP([rvNumber(1), lookup, ret, BLANK, rvNumber(0), rvNumber(-1)]))).toBe(
      "c"
    );
  });

  it("searchMode 2 binary (ascending sorted)", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)], [rvNumber(4)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")], [rvString("c")], [rvString("d")]]);
    expect(asString(fnXLOOKUP([rvNumber(3), lookup, ret, BLANK, rvNumber(0), rvNumber(2)]))).toBe(
      "c"
    );
  });

  it("searchMode -2 binary descending", () => {
    const lookup = rvArray([[rvNumber(4)], [rvNumber(3)], [rvNumber(2)], [rvNumber(1)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")], [rvString("c")], [rvString("d")]]);
    expect(asString(fnXLOOKUP([rvNumber(3), lookup, ret, BLANK, rvNumber(0), rvNumber(-2)]))).toBe(
      "b"
    );
  });

  it("matchMode 1 next-larger (linear)", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(5)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    expect(asString(fnXLOOKUP([rvNumber(4), lookup, ret, BLANK, rvNumber(1)]))).toBe("c");
  });

  it("horizontal lookup with narrow return_array yields BLANK for out-of-range index", () => {
    const lookup = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    // return has only 1 column — found index 1 (for value 2) is out of width
    const ret = rvArray([[rvString("x")]]);
    expect(fnXLOOKUP([rvNumber(2), lookup, ret])).toEqual(BLANK);
  });

  it("propagates error from lookup value", () => {
    const lookup = rvArray([[rvNumber(1)]]);
    const ret = rvArray([[rvString("a")]]);
    expect(fnXLOOKUP([ERRORS.NA, lookup, ret])).toEqual(ERRORS.NA);
  });
});

describe("XMATCH comprehensive", () => {
  it("non-array lookup_array → #VALUE!", () => {
    expect(fnXMATCH([rvNumber(1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("propagates error from lookup value", () => {
    expect(fnXMATCH([ERRORS.NA, rvArray([[rvNumber(1)]])])).toEqual(ERRORS.NA);
  });

  it("searchMode -1 reverse scan finds last occurrence", () => {
    const arr = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(1)]]);
    expect(asNumber(fnXMATCH([rvNumber(1), arr, rvNumber(0), rvNumber(-1)]))).toBe(3);
  });

  it("string match_mode=0 case-insensitive", () => {
    const arr = rvArray([[rvString("Banana")]]);
    expect(asNumber(fnXMATCH([rvString("banana"), arr]))).toBe(1);
  });

  it("unsupported matchMode returns #N/A (value outside {-1,0,1})", () => {
    const arr = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    expect(fnXMATCH([rvNumber(1), arr, rvNumber(99)])).toEqual(ERRORS.NA);
  });

  it("matchMode 1 returns smallest ≥ lookup", () => {
    const arr = rvArray([[rvNumber(10)], [rvNumber(20)], [rvNumber(30)]]);
    expect(asNumber(fnXMATCH([rvNumber(15), arr, rvNumber(1)]))).toBe(2);
  });
});

describe("ADDRESS comprehensive", () => {
  it("row < 1 → #VALUE! (regression)", () => {
    expect(fnADDRESS([rvNumber(0), rvNumber(1)])).toEqual(ERRORS.VALUE);
    expect(fnADDRESS([rvNumber(-1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("col < 1 → #VALUE!", () => {
    expect(fnADDRESS([rvNumber(1), rvNumber(0)])).toEqual(ERRORS.VALUE);
  });

  it("abs_num outside {1..4} → #VALUE!", () => {
    expect(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(5)])).toEqual(ERRORS.VALUE);
    expect(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(0)])).toEqual(ERRORS.VALUE);
  });

  it("abs_num=3 — col absolute, row relative → $C5", () => {
    expect(asString(fnADDRESS([rvNumber(5), rvNumber(3), rvNumber(3)]))).toBe("$C5");
  });

  it("R1C1 absolute (a1=false, abs_num=1)", () => {
    expect(asString(fnADDRESS([rvNumber(5), rvNumber(3), rvNumber(1), rvBoolean(false)]))).toBe(
      "R5C3"
    );
  });

  it("R1C1 relative (abs_num=4)", () => {
    expect(asString(fnADDRESS([rvNumber(5), rvNumber(3), rvNumber(4), rvBoolean(false)]))).toBe(
      "R[5]C[3]"
    );
  });

  it("plain sheet name prefix (no quotes)", () => {
    expect(
      asString(
        fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(1), rvBoolean(true), rvString("Sheet1")])
      )
    ).toBe("Sheet1!$A$1");
  });

  it("sheet name with space is quoted", () => {
    expect(
      asString(
        fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(1), rvBoolean(true), rvString("My Sheet")])
      )
    ).toBe("'My Sheet'!$A$1");
  });

  it("column letters beyond ZZ (col 703 → AAA)", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(703)]))).toBe("$AAA$1");
  });

  it("fractional row/col truncated", () => {
    expect(asString(fnADDRESS([rvNumber(5.9), rvNumber(1.2)]))).toBe("$A$5");
  });

  it("propagates errors from row/col args", () => {
    expect(fnADDRESS([ERRORS.NA, rvNumber(1)])).toEqual(ERRORS.NA);
    expect(fnADDRESS([rvNumber(1), ERRORS.REF])).toEqual(ERRORS.REF);
  });
});

describe("LOOKUP comprehensive", () => {
  it("propagates error from lookup value", () => {
    const la = rvArray([[rvNumber(1)]]);
    expect(fnLOOKUP([ERRORS.NA, la])).toEqual(ERRORS.NA);
  });

  it("non-array lookup_vector → #N/A", () => {
    expect(fnLOOKUP([rvNumber(1), rvNumber(1)])).toEqual(ERRORS.NA);
  });

  it("vector form — vertical M×1 lookup vector", () => {
    const lookup = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    const ret = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    expect(asString(fnLOOKUP([rvNumber(2), lookup, ret]))).toBe("b");
  });

  it("below-smallest lookup returns #N/A", () => {
    const lookup = rvArray([[rvNumber(10), rvNumber(20)]]);
    const ret = rvArray([[rvString("a"), rvString("b")]]);
    expect(fnLOOKUP([rvNumber(5), lookup, ret])).toEqual(ERRORS.NA);
  });

  it("array form with 3×2 (rows > cols) scans first column, returns last column", () => {
    const arr = rvArray([
      [rvNumber(1), rvString("a")],
      [rvNumber(2), rvString("b")],
      [rvNumber(3), rvString("c")]
    ]);
    expect(asString(fnLOOKUP([rvNumber(2), arr]))).toBe("b");
  });

  it("array form with 2×3 (cols >= rows) scans first row, returns last row", () => {
    const arr = rvArray([
      [rvNumber(1), rvNumber(2), rvNumber(3)],
      [rvString("a"), rvString("b"), rvString("c")]
    ]);
    expect(asString(fnLOOKUP([rvNumber(2), arr]))).toBe("b");
  });

  it("vector form — case-insensitive string approximate match", () => {
    const lookup = rvArray([[rvString("apple"), rvString("mango"), rvString("zebra")]]);
    const ret = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    expect(asNumber(fnLOOKUP([rvString("orange"), lookup, ret]))).toBe(2);
  });
});

describe("TRANSPOSE comprehensive", () => {
  it("transpose of 1×3 → 3×1", () => {
    const r = fnTRANSPOSE([rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]])]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[2][0] as NumberValue).value).toBe(3);
  });

  it("transpose of M×1 → 1×M", () => {
    const r = fnTRANSPOSE([
      rvArray([[rvString("a")], [rvString("b")], [rvString("c")]])
    ]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(3);
  });

  it("transpose string scalar → 1×1 array containing the string", () => {
    const r = fnTRANSPOSE([rvString("x")]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as StringValue).value).toBe("x");
  });

  it("transpose of error propagates the error (R8-P1)", () => {
    // Excel's TRANSPOSE propagates scalar errors so callers that
    // aggregate the result (e.g. `SUM(TRANSPOSE(#N/A))`) see the error
    // surface naturally rather than operating on a 1×1 array.
    expect(fnTRANSPOSE([ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("transpose preserves cell kinds (string, number, blank)", () => {
    const r = fnTRANSPOSE([rvArray([[rvString("a"), rvNumber(1), BLANK]])]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as StringValue).value).toBe("a");
    expect((r.rows[1][0] as NumberValue).value).toBe(1);
    expect(r.rows[2][0].kind).toBe(RVKind.Blank);
  });
});

describe("AREAS comprehensive", () => {
  it("always returns 1 regardless of argument kind", () => {
    expect(asNumber(fnAREAS([rvString("x")]))).toBe(1);
    expect(asNumber(fnAREAS([rvBoolean(true)]))).toBe(1);
    expect(asNumber(fnAREAS([rvArray([[rvNumber(1)]])]))).toBe(1);
    expect(asNumber(fnAREAS([BLANK]))).toBe(1);
  });

  it("even an error argument returns 1 (not propagated — current limitation)", () => {
    expect(asNumber(fnAREAS([ERRORS.NA]))).toBe(1);
  });
});

// ============================================================================
// R8 deep coverage: VLOOKUP — approximate vs exact match, corners
// ============================================================================

describe("VLOOKUP deep coverage", () => {
  const table = rvArray([
    [rvNumber(1), rvString("one")],
    [rvNumber(3), rvString("three")],
    [rvNumber(5), rvString("five")],
    [rvNumber(7), rvString("seven")]
  ]);

  it("exact match hit", () => {
    expect(asString(fnVLOOKUP([rvNumber(3), table, rvNumber(2), rvBoolean(false)]))).toBe("three");
  });

  it("exact match miss → #N/A", () => {
    expect(fnVLOOKUP([rvNumber(4), table, rvNumber(2), rvBoolean(false)])).toEqual(ERRORS.NA);
  });

  it("approximate (default) finds largest ≤ lookup_value", () => {
    // value 4 → between 3 and 5 → should match row for 3
    expect(asString(fnVLOOKUP([rvNumber(4), table, rvNumber(2), rvBoolean(true)]))).toBe("three");
  });

  it("approximate lookup value smaller than first → #N/A", () => {
    expect(fnVLOOKUP([rvNumber(0), table, rvNumber(2), rvBoolean(true)])).toEqual(ERRORS.NA);
  });

  it("approximate lookup value larger than last → last row", () => {
    expect(asString(fnVLOOKUP([rvNumber(999), table, rvNumber(2), rvBoolean(true)]))).toBe("seven");
  });

  it("col_index fractional → truncated toward zero", () => {
    expect(asString(fnVLOOKUP([rvNumber(3), table, rvNumber(2.9), rvBoolean(false)]))).toBe(
      "three"
    );
  });

  it("col_index < 1 → #REF! (engine uses #REF! for out-of-range index)", () => {
    expect(fnVLOOKUP([rvNumber(3), table, rvNumber(0), rvBoolean(false)])).toEqual(ERRORS.REF);
  });

  it("col_index > width → #REF!", () => {
    expect(fnVLOOKUP([rvNumber(3), table, rvNumber(10), rvBoolean(false)])).toEqual(ERRORS.REF);
  });

  it("string lookup exact returns found row", () => {
    const strTable = rvArray([
      [rvString("apple"), rvNumber(1)],
      [rvString("banana"), rvNumber(2)],
      [rvString("cherry"), rvNumber(3)]
    ]);
    expect(asNumber(fnVLOOKUP([rvString("banana"), strTable, rvNumber(2), rvBoolean(false)]))).toBe(
      2
    );
  });

  it("error in lookup_value propagates", () => {
    expect(fnVLOOKUP([ERRORS.NA, table, rvNumber(2), rvBoolean(false)])).toEqual(ERRORS.NA);
  });
});

// ============================================================================
// HLOOKUP deep coverage
// ============================================================================

describe("HLOOKUP deep coverage", () => {
  const h = rvArray([
    [rvNumber(1), rvNumber(3), rvNumber(5), rvNumber(7)],
    [rvString("a"), rvString("b"), rvString("c"), rvString("d")]
  ]);

  it("exact hit", () => {
    expect(asString(fnHLOOKUP([rvNumber(5), h, rvNumber(2), rvBoolean(false)]))).toBe("c");
  });

  it("approximate 6 → b (matches 3 in sorted row)", () => {
    expect(asString(fnHLOOKUP([rvNumber(4), h, rvNumber(2), rvBoolean(true)]))).toBe("b");
  });

  it("row_index fractional truncated", () => {
    expect(asString(fnHLOOKUP([rvNumber(1), h, rvNumber(2.5), rvBoolean(false)]))).toBe("a");
  });

  it("row_index < 1 → #REF!", () => {
    expect(fnHLOOKUP([rvNumber(1), h, rvNumber(0), rvBoolean(false)])).toEqual(ERRORS.REF);
  });

  it("row_index too big → #REF!", () => {
    expect(fnHLOOKUP([rvNumber(1), h, rvNumber(10), rvBoolean(false)])).toEqual(ERRORS.REF);
  });

  it("not found exact → #N/A", () => {
    expect(fnHLOOKUP([rvNumber(2), h, rvNumber(2), rvBoolean(false)])).toEqual(ERRORS.NA);
  });
});

// ============================================================================
// XLOOKUP deep coverage
// ============================================================================

describe("XLOOKUP deep coverage", () => {
  const keys = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(5)], [rvNumber(7)]]);
  const vals = rvArray([[rvString("a")], [rvString("b")], [rvString("c")], [rvString("d")]]);

  it("exact match found (default match_mode=0)", () => {
    expect(asString(fnXLOOKUP([rvNumber(3), keys, vals]))).toBe("b");
  });

  it("exact match not found → #N/A", () => {
    expect(fnXLOOKUP([rvNumber(4), keys, vals])).toEqual(ERRORS.NA);
  });

  it("if_not_found takes precedence", () => {
    expect(asString(fnXLOOKUP([rvNumber(4), keys, vals, rvString("MISSING")]))).toBe("MISSING");
  });

  it("match_mode=-1 returns next smaller", () => {
    expect(asString(fnXLOOKUP([rvNumber(4), keys, vals, rvString("NF"), rvNumber(-1)]))).toBe("b");
  });

  it("match_mode=1 returns next larger", () => {
    expect(asString(fnXLOOKUP([rvNumber(4), keys, vals, rvString("NF"), rvNumber(1)]))).toBe("c");
  });

  it("match_mode=2 wildcard", () => {
    const strKeys = rvArray([[rvString("apple")], [rvString("banana")], [rvString("cherry")]]);
    const strVals = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    expect(
      asNumber(fnXLOOKUP([rvString("ban*"), strKeys, strVals, rvString("NF"), rvNumber(2)]))
    ).toBe(2);
  });

  it("search_mode=-1 reverses search", () => {
    const dupKeys = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(3)], [rvNumber(5)]]);
    const dupVals = rvArray([[rvString("a")], [rvString("b1")], [rvString("b2")], [rvString("d")]]);
    // search_mode=1 (default) finds first = "b1"
    expect(asString(fnXLOOKUP([rvNumber(3), dupKeys, dupVals]))).toBe("b1");
    // search_mode=-1 finds last = "b2"
    expect(
      asString(
        fnXLOOKUP([rvNumber(3), dupKeys, dupVals, rvString("NF"), rvNumber(0), rvNumber(-1)])
      )
    ).toBe("b2");
  });

  it("search_mode=2 binary search (ascending assumed)", () => {
    expect(
      asString(fnXLOOKUP([rvNumber(5), keys, vals, rvString("NF"), rvNumber(0), rvNumber(2)]))
    ).toBe("c");
  });

  it("horizontal lookup: 1-row lookup, multi-row return", () => {
    const hKeys = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const hReturn = rvArray([
      [rvString("a1"), rvString("a2"), rvString("a3")],
      [rvString("b1"), rvString("b2"), rvString("b3")]
    ]);
    const r = fnXLOOKUP([rvNumber(2), hKeys, hReturn]);
    expect(r.kind).toBe(RVKind.Array);
    const a = r as ArrayValue;
    // 水平查找 + 多行 return: 应返回对应列，2 行 1 列
    expect(a.height).toBe(2);
    expect(a.width).toBe(1);
    expect((a.rows[0][0] as StringValue).value).toBe("a2");
    expect((a.rows[1][0] as StringValue).value).toBe("b2");
  });

  it("error in lookup_value propagates", () => {
    expect(fnXLOOKUP([ERRORS.NA, keys, vals])).toEqual(ERRORS.NA);
  });
});

// ============================================================================
// XMATCH deep coverage
// ============================================================================

describe("XMATCH deep coverage", () => {
  const arr = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(5)], [rvNumber(7)]]);

  it("exact match returns 1-based index", () => {
    expect(asNumber(fnXMATCH([rvNumber(3), arr]))).toBe(2);
  });

  it("not found exact → #N/A", () => {
    expect(fnXMATCH([rvNumber(4), arr])).toEqual(ERRORS.NA);
  });

  it("match_mode=-1 next smaller", () => {
    expect(asNumber(fnXMATCH([rvNumber(4), arr, rvNumber(-1)]))).toBe(2);
  });

  it("match_mode=1 next larger", () => {
    expect(asNumber(fnXMATCH([rvNumber(4), arr, rvNumber(1)]))).toBe(3);
  });

  it("search_mode=-1 last match", () => {
    const dup = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(3)], [rvNumber(5)]]);
    expect(asNumber(fnXMATCH([rvNumber(3), dup, rvNumber(0), rvNumber(-1)]))).toBe(3);
  });

  it("lookup_value error propagates", () => {
    const arr = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    expect(fnXMATCH([ERRORS.NA, arr])).toEqual(ERRORS.NA);
  });
});

// ============================================================================
// MATCH deep coverage
// ============================================================================

describe("MATCH deep coverage", () => {
  it("match_type=0 exact hit", () => {
    const arr = rvArray([[rvNumber(1), rvNumber(5), rvNumber(3)]]);
    expect(asNumber(fnMATCH([rvNumber(5), arr, rvNumber(0)]))).toBe(2);
  });

  it("match_type=0 no hit → #N/A", () => {
    const arr = rvArray([[rvNumber(1), rvNumber(5), rvNumber(3)]]);
    expect(fnMATCH([rvNumber(99), arr, rvNumber(0)])).toEqual(ERRORS.NA);
  });

  it("match_type=1 (default) finds largest ≤ lookup in ascending", () => {
    const asc = rvArray([[rvNumber(1), rvNumber(3), rvNumber(5), rvNumber(7)]]);
    // lookup 4 → largest ≤ 4 is 3, index 2
    expect(asNumber(fnMATCH([rvNumber(4), asc, rvNumber(1)]))).toBe(2);
  });

  it("match_type=-1 finds smallest ≥ lookup in descending", () => {
    const desc = rvArray([[rvNumber(7), rvNumber(5), rvNumber(3), rvNumber(1)]]);
    // lookup 4 → smallest ≥ 4 is 5, index 2
    expect(asNumber(fnMATCH([rvNumber(4), desc, rvNumber(-1)]))).toBe(2);
  });

  it("match_type=0 with wildcards", () => {
    const arr = rvArray([[rvString("apple"), rvString("banana")]]);
    expect(asNumber(fnMATCH([rvString("ban*"), arr, rvNumber(0)]))).toBe(2);
  });

  it("MATCH with escaped wildcard ~* matches literal asterisk (R8 fix)", () => {
    const arr = rvArray([[rvString("a*b"), rvString("banana")]]);
    expect(asNumber(fnMATCH([rvString("a~*b"), arr, rvNumber(0)]))).toBe(1);
  });

  it("lookup value error propagates", () => {
    const arr = rvArray([[rvNumber(1)]]);
    expect(fnMATCH([ERRORS.NA, arr, rvNumber(0)])).toEqual(ERRORS.NA);
  });
});

// ============================================================================
// INDEX deep coverage
// ============================================================================

describe("INDEX deep coverage", () => {
  const arr = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvNumber(4), rvNumber(5), rvNumber(6)],
    [rvNumber(7), rvNumber(8), rvNumber(9)]
  ]);

  it("cell index", () => {
    expect(asNumber(fnINDEX([arr, rvNumber(2), rvNumber(3)]))).toBe(6);
  });

  it("row=0 returns entire row", () => {
    const r = fnINDEX([arr, rvNumber(0), rvNumber(2)]);
    expect(r.kind).toBe(RVKind.Array);
    const a = r as ArrayValue;
    expect(a.height).toBe(3);
    expect(a.width).toBe(1);
  });

  it("col=0 returns entire row", () => {
    const r = fnINDEX([arr, rvNumber(2), rvNumber(0)]);
    expect(r.kind).toBe(RVKind.Array);
    const a = r as ArrayValue;
    expect(a.height).toBe(1);
    expect(a.width).toBe(3);
  });

  it("fractional indices truncated", () => {
    expect(asNumber(fnINDEX([arr, rvNumber(2.9), rvNumber(3.2)]))).toBe(6);
  });

  it("row > height → #REF!", () => {
    expect(fnINDEX([arr, rvNumber(10), rvNumber(1)])).toEqual(ERRORS.REF);
  });

  it("col > width → #REF!", () => {
    expect(fnINDEX([arr, rvNumber(1), rvNumber(10)])).toEqual(ERRORS.REF);
  });

  it("negative row → #VALUE!", () => {
    expect(fnINDEX([arr, rvNumber(-1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("scalar array passes through topLeft", () => {
    expect(asNumber(fnINDEX([rvNumber(42), rvNumber(1), rvNumber(1)]))).toBe(42);
  });
});

// ============================================================================
// ADDRESS deep coverage
// ============================================================================

describe("ADDRESS deep coverage", () => {
  it("absolute A1 style (default)", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(1)]))).toBe("$A$1");
  });

  it("abs_num=2: absolute row, relative col", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(2)]))).toBe("A$1");
  });

  it("abs_num=3: relative row, absolute col", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(3)]))).toBe("$A1");
  });

  it("abs_num=4: relative", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(4)]))).toBe("A1");
  });

  it("R1C1 style", () => {
    expect(asString(fnADDRESS([rvNumber(5), rvNumber(3), rvNumber(1), rvBoolean(false)]))).toBe(
      "R5C3"
    );
  });

  it("R1C1 relative", () => {
    expect(asString(fnADDRESS([rvNumber(5), rvNumber(3), rvNumber(4), rvBoolean(false)]))).toBe(
      "R[5]C[3]"
    );
  });

  it("with plain sheet name", () => {
    expect(
      asString(
        fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(1), rvBoolean(true), rvString("Sheet1")])
      )
    ).toBe("Sheet1!$A$1");
  });

  it("with sheet name requiring quotes", () => {
    expect(
      asString(
        fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(1), rvBoolean(true), rvString("My Sheet")])
      )
    ).toBe("'My Sheet'!$A$1");
  });

  it("large column → multi-letter", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(27)]))).toBe("$AA$1");
  });

  it("col 16384 → XFD", () => {
    expect(asString(fnADDRESS([rvNumber(1), rvNumber(16384)]))).toBe("$XFD$1");
  });

  it("row=0 → #VALUE!", () => {
    expect(fnADDRESS([rvNumber(0), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("abs_num outside [1,4] → #VALUE!", () => {
    expect(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(5)])).toEqual(ERRORS.VALUE);
    expect(fnADDRESS([rvNumber(1), rvNumber(1), rvNumber(0)])).toEqual(ERRORS.VALUE);
  });
});

// ============================================================================
// LOOKUP deep coverage
// ============================================================================

describe("LOOKUP deep coverage", () => {
  it("vector form: numeric search in ascending column", () => {
    const vec = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(5)]]);
    const res = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    expect(asString(fnLOOKUP([rvNumber(3), vec, res]))).toBe("b");
  });

  it("vector form: string search case-insensitive", () => {
    const vec = rvArray([[rvString("Apple")], [rvString("Banana")], [rvString("Cherry")]]);
    const res = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    expect(asNumber(fnLOOKUP([rvString("banana"), vec, res]))).toBe(2);
  });

  it("vector form: horizontal search (1-row vector)", () => {
    const vec = rvArray([[rvNumber(1), rvNumber(3), rvNumber(5)]]);
    const res = rvArray([[rvString("a"), rvString("b"), rvString("c")]]);
    expect(asString(fnLOOKUP([rvNumber(3), vec, res]))).toBe("b");
  });

  it("array form: 2D array, more columns than rows → horizontal", () => {
    const arr = rvArray([
      [rvNumber(1), rvNumber(3), rvNumber(5)],
      [rvString("a"), rvString("b"), rvString("c")]
    ]);
    expect(asString(fnLOOKUP([rvNumber(3), arr]))).toBe("b");
  });

  it("array form: more rows than cols → vertical", () => {
    const arr = rvArray([
      [rvNumber(1), rvString("a")],
      [rvNumber(3), rvString("b")],
      [rvNumber(5), rvString("c")]
    ]);
    expect(asString(fnLOOKUP([rvNumber(3), arr]))).toBe("b");
  });

  it("lookup value smaller than first → #N/A", () => {
    const vec = rvArray([[rvNumber(1)], [rvNumber(3)]]);
    const res = rvArray([[rvString("a")], [rvString("b")]]);
    expect(fnLOOKUP([rvNumber(0), vec, res])).toEqual(ERRORS.NA);
  });

  it("returns largest ≤ lookup when exact miss", () => {
    const vec = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(5)]]);
    const res = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    expect(asString(fnLOOKUP([rvNumber(4), vec, res]))).toBe("b");
  });
});

// ============================================================================
// TRANSPOSE deep coverage
// ============================================================================

describe("TRANSPOSE deep coverage", () => {
  it("2×3 → 3×2", () => {
    const r = fnTRANSPOSE([
      rvArray([
        [rvNumber(1), rvNumber(2), rvNumber(3)],
        [rvNumber(4), rvNumber(5), rvNumber(6)]
      ])
    ]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[0][1] as NumberValue).value).toBe(4);
    expect((r.rows[2][1] as NumberValue).value).toBe(6);
  });

  it("1×N → N×1", () => {
    const r = fnTRANSPOSE([rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]])]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(1);
  });

  it("N×1 → 1×N", () => {
    const r = fnTRANSPOSE([rvArray([[rvNumber(1)], [rvNumber(2)]])]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(2);
  });

  it("scalar → 1×1 array", () => {
    const r = fnTRANSPOSE([rvNumber(42)]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(1);
  });

  it("preserves mixed cell kinds", () => {
    const r = fnTRANSPOSE([
      rvArray([
        [rvNumber(1), rvString("a")],
        [BLANK, rvBoolean(true)]
      ])
    ]) as ArrayValue;
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[1][1] as RuntimeValue).kind).toBe(RVKind.Boolean);
  });
});

// ============================================================================
// ROWS / COLUMNS deep coverage
// ============================================================================

describe("ROWS / COLUMNS deep coverage", () => {
  it("ROWS of 5×3 array = 5", () => {
    const a = rvArray(Array.from({ length: 5 }, () => [rvNumber(1), rvNumber(2), rvNumber(3)]));
    expect(asNumber(fnROWS([a]))).toBe(5);
  });

  it("COLUMNS of 5×3 array = 3", () => {
    const a = rvArray(Array.from({ length: 5 }, () => [rvNumber(1), rvNumber(2), rvNumber(3)]));
    expect(asNumber(fnCOLUMNS([a]))).toBe(3);
  });

  it("ROWS of 1×N = 1", () => {
    expect(asNumber(fnROWS([rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]])]))).toBe(1);
  });

  it("COLUMNS of N×1 = 1", () => {
    expect(asNumber(fnCOLUMNS([rvArray([[rvNumber(1)], [rvNumber(2)]])]))).toBe(1);
  });

  it("ROWS / COLUMNS of scalar = 1", () => {
    expect(asNumber(fnROWS([rvNumber(42)]))).toBe(1);
    expect(asNumber(fnCOLUMNS([rvString("x")]))).toBe(1);
  });
});

// ============================================================================
// AREAS comprehensive
// ============================================================================

describe("AREAS comprehensive", () => {
  it("returns 1 for single reference (engine limitation)", () => {
    expect(asNumber(fnAREAS([rvArray([[rvNumber(1)]])]))).toBe(1);
  });

  it("returns 1 for scalar", () => {
    expect(asNumber(fnAREAS([rvNumber(42)]))).toBe(1);
  });

  it("requires at least one argument", () => {
    expect(fnAREAS([])).toEqual(ERRORS.VALUE);
  });
});

// ============================================================================
// COLUMNS saturation — bring per-function reference count to 10+.
// The existing tests exercise the common 3x3 / 1xN / Nx1 cases; these
// focus on boundary shapes and argument coercion rules.
// ============================================================================

describe("COLUMNS saturation", () => {
  it("COLUMNS of a 1×1 array = 1", () => {
    expect(asNumber(fnCOLUMNS([rvArray([[rvNumber(1)]])]))).toBe(1);
  });

  it("COLUMNS of a 10-column single-row array", () => {
    const cells = Array.from({ length: 10 }, (_, i) => rvNumber(i));
    expect(asNumber(fnCOLUMNS([rvArray([cells])]))).toBe(10);
  });

  it("COLUMNS of a 100-column array (stress)", () => {
    const cells = Array.from({ length: 100 }, (_, i) => rvNumber(i));
    expect(asNumber(fnCOLUMNS([rvArray([cells])]))).toBe(100);
  });

  it("COLUMNS of a 5×5 square array = 5", () => {
    const rows = Array.from({ length: 5 }, (_, r) =>
      Array.from({ length: 5 }, (_, c) => rvNumber(r * 5 + c))
    );
    expect(asNumber(fnCOLUMNS([rvArray(rows)]))).toBe(5);
  });

  it("COLUMNS of a BLANK scalar = 1", () => {
    expect(asNumber(fnCOLUMNS([BLANK]))).toBe(1);
  });

  it("COLUMNS of a numeric scalar = 1", () => {
    expect(asNumber(fnCOLUMNS([rvNumber(42)]))).toBe(1);
  });

  it("COLUMNS of a boolean scalar = 1", () => {
    expect(asNumber(fnCOLUMNS([rvBoolean(true)]))).toBe(1);
  });

  it("COLUMNS of a mixed-type array still reports shape", () => {
    const r = rvArray([[rvNumber(1), rvString("a"), rvBoolean(true), BLANK]]);
    expect(asNumber(fnCOLUMNS([r]))).toBe(4);
  });
});

describe("ROW / COLUMN / ROWS extras (R9 saturation)", () => {
  // fnROW / fnCOLUMN return #VALUE! when called without an evaluator-
  // supplied BoundCellRef context; the reference-aware path lives in
  // tryEvaluateRefFunction in the evaluator. These tests pin the
  // function-level behaviour — end-to-end coverage via Workbook is in
  // the evaluator integration tests.
  it("ROW on non-reference returns #VALUE!", () => {
    expect(fnROW([rvNumber(42)])).toEqual(ERRORS.VALUE);
  });
  it("ROW on array returns #VALUE!", () => {
    expect(fnROW([rvArray([[rvNumber(1)]])])).toEqual(ERRORS.VALUE);
  });
  it("ROW on error still #VALUE!", () => {
    expect(fnROW([ERRORS.NA])).toEqual(ERRORS.VALUE);
  });
  it("ROW on boolean returns #VALUE!", () => {
    expect(fnROW([rvBoolean(true)])).toEqual(ERRORS.VALUE);
  });
  it("COLUMN on non-reference returns #VALUE!", () => {
    expect(fnCOLUMN([rvNumber(42)])).toEqual(ERRORS.VALUE);
  });
  it("COLUMN on array returns #VALUE!", () => {
    expect(fnCOLUMN([rvArray([[rvNumber(1)]])])).toEqual(ERRORS.VALUE);
  });
  it("ROWS of 3×2 = 3", () => {
    const arr = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)],
      [rvNumber(5), rvNumber(6)]
    ]);
    expect(asNumber(fnROWS([arr]))).toBe(3);
  });
  it("ROWS of string scalar = 1", () => {
    expect(asNumber(fnROWS([rvString("x")]))).toBe(1);
  });
  it("ROWS of boolean scalar = 1", () => {
    expect(asNumber(fnROWS([rvBoolean(true)]))).toBe(1);
  });
  it("ROWS of 1×N row vector = 1", () => {
    expect(asNumber(fnROWS([rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]])]))).toBe(1);
  });
  it("COLUMNS of 3×2 = 2", () => {
    const arr = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)]
    ]);
    expect(asNumber(fnCOLUMNS([arr]))).toBe(2);
  });
  it("COLUMNS of string scalar = 1", () => {
    expect(asNumber(fnCOLUMNS([rvString("x")]))).toBe(1);
  });
});
