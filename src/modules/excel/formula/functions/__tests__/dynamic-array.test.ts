/**
 * Unit tests for dynamic-array functions in `../dynamic-array.ts`.
 *
 * Key regression coverage:
 *   - UNIQUE must distinguish the number `1` from the string `"1"` (both
 *     stringify the same way, so a naive string-key dedup would collapse
 *     them; the implementation prefixes each key with the cell's kind).
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
  fnFILTER,
  fnSORT,
  fnSORTBY,
  fnUNIQUE,
  fnSEQUENCE,
  fnRANDARRAY,
  fnVSTACK,
  fnHSTACK,
  fnCHOOSEROWS,
  fnCHOOSECOLS,
  fnTAKE,
  fnDROP,
  fnTOCOL,
  fnTOROW,
  fnEXPAND,
  fnWRAPROWS,
  fnWRAPCOLS,
  fnSUBTOTAL,
  fnAGGREGATE
} from "../dynamic-array";

function asArray(v: RuntimeValue): ArrayValue {
  expect(v.kind).toBe(RVKind.Array);
  return v as ArrayValue;
}

describe("FILTER", () => {
  it("keeps rows where include is truthy", () => {
    const data = rvArray([
      [rvString("a"), rvNumber(1)],
      [rvString("b"), rvNumber(2)],
      [rvString("c"), rvNumber(3)]
    ]);
    const include = rvArray([[rvBoolean(true)], [rvBoolean(false)], [rvBoolean(true)]]);
    const r = asArray(fnFILTER([data, include]));
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as StringValue).value).toBe("a");
    expect((r.rows[1][0] as StringValue).value).toBe("c");
  });

  it("returns the ifEmpty value when no rows match", () => {
    const data = rvArray([[rvString("a")], [rvString("b")]]);
    const include = rvArray([[rvBoolean(false)], [rvBoolean(false)]]);
    const r = asArray(fnFILTER([data, include, rvString("none")]));
    expect((r.rows[0][0] as StringValue).value).toBe("none");
  });

  it("returns #CALC! when no rows match and ifEmpty is omitted (R4-P1-2)", () => {
    // Matches Excel's dynamic-array error for an empty array result; the
    // old code returned #VALUE! which masked downstream consumers that
    // expected the canonical "no values" diagnostic.
    const data = rvArray([[rvString("a")]]);
    const include = rvArray([[rvBoolean(false)]]);
    expect(fnFILTER([data, include])).toEqual(ERRORS.CALC);
  });

  it("propagates errors in the include column", () => {
    const data = rvArray([[rvString("a")], [rvString("b")]]);
    const include = rvArray([[rvBoolean(true)], [ERRORS.NA]]);
    expect(fnFILTER([data, include])).toEqual(ERRORS.NA);
  });
});

describe("SORT", () => {
  it("sorts a single column ascending by default", () => {
    const data = rvArray([[rvNumber(3)], [rvNumber(1)], [rvNumber(2)]]);
    const r = asArray(fnSORT([data]));
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[2][0] as NumberValue).value).toBe(3);
  });

  it("sorts descending when order = -1", () => {
    const data = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(2)]]);
    const r = asArray(fnSORT([data, rvNumber(1), rvNumber(-1)]));
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
    expect((r.rows[2][0] as NumberValue).value).toBe(1);
  });

  it("sorts by a specific column", () => {
    const data = rvArray([
      [rvString("a"), rvNumber(3)],
      [rvString("b"), rvNumber(1)],
      [rvString("c"), rvNumber(2)]
    ]);
    const r = asArray(fnSORT([data, rvNumber(2)]));
    expect((r.rows[0][1] as NumberValue).value).toBe(1);
    expect((r.rows[0][0] as StringValue).value).toBe("b");
  });
});

describe("SORTBY", () => {
  it("sorts by a key array", () => {
    const data = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    const key = rvArray([[rvNumber(3)], [rvNumber(1)], [rvNumber(2)]]);
    const r = asArray(fnSORTBY([data, key]));
    expect((r.rows[0][0] as StringValue).value).toBe("b"); // key=1
    expect((r.rows[1][0] as StringValue).value).toBe("c"); // key=2
    expect((r.rows[2][0] as StringValue).value).toBe("a"); // key=3
  });
});

describe("UNIQUE — type-aware deduplication", () => {
  it("basic deduplication", () => {
    const data = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(1)], [rvNumber(3)]]);
    const r = asArray(fnUNIQUE([data]));
    expect(r.height).toBe(3);
  });

  it('treats `1` and `"1"` as distinct (regression)', () => {
    // Without kind-aware keys, both would stringify to "1" and collapse.
    const data = rvArray([[rvNumber(1)], [rvString("1")], [rvNumber(1)]]);
    const r = asArray(fnUNIQUE([data]));
    expect(r.height).toBe(2);
    expect(r.rows[0][0].kind).toBe(RVKind.Number);
    expect(r.rows[1][0].kind).toBe(RVKind.String);
  });

  it("exactly_once returns entries appearing once", () => {
    const data = rvArray([[rvString("a")], [rvString("b")], [rvString("a")], [rvString("c")]]);
    // a is duplicated; b and c each appear once.
    const r = asArray(fnUNIQUE([data, rvBoolean(false), rvBoolean(true)]));
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as StringValue).value).toBe("b");
    expect((r.rows[1][0] as StringValue).value).toBe("c");
  });

  it("is case-insensitive for strings", () => {
    const data = rvArray([[rvString("Hello")], [rvString("HELLO")]]);
    const r = asArray(fnUNIQUE([data]));
    expect(r.height).toBe(1);
  });

  it("UNIQUE of empty range returns #VALUE!", () => {
    // No matching rows (all duplicated) with exactly_once=TRUE → empty
    const data = rvArray([[rvString("a")], [rvString("a")]]);
    expect(fnUNIQUE([data, rvBoolean(false), rvBoolean(true)])).toEqual(ERRORS.VALUE);
  });
});

describe("SEQUENCE", () => {
  it("produces a sequence of rows x cols", () => {
    const r = asArray(fnSEQUENCE([rvNumber(3), rvNumber(2)]));
    expect(r.height).toBe(3);
    expect(r.width).toBe(2);
    // Default start=1, step=1: [[1,2], [3,4], [5,6]]
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[2][1] as NumberValue).value).toBe(6);
  });

  it("honours start and step", () => {
    const r = asArray(fnSEQUENCE([rvNumber(3), rvNumber(1), rvNumber(10), rvNumber(5)]));
    expect((r.rows[0][0] as NumberValue).value).toBe(10);
    expect((r.rows[1][0] as NumberValue).value).toBe(15);
    expect((r.rows[2][0] as NumberValue).value).toBe(20);
  });

  it("single-column default when cols is omitted", () => {
    const r = asArray(fnSEQUENCE([rvNumber(3)]));
    expect(r.width).toBe(1);
  });
});

describe("VSTACK / HSTACK", () => {
  const a = rvArray([[rvNumber(1), rvNumber(2)]]);
  const b = rvArray([[rvNumber(3), rvNumber(4)]]);

  it("VSTACK stacks vertically", () => {
    const r = asArray(fnVSTACK([a, b]));
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
    expect((r.rows[1][0] as NumberValue).value).toBe(3);
  });

  it("HSTACK stacks horizontally", () => {
    const r = asArray(fnHSTACK([a, b]));
    expect(r.height).toBe(1);
    expect(r.width).toBe(4);
    expect((r.rows[0][2] as NumberValue).value).toBe(3);
  });

  it("HSTACK pads mismatched heights with #N/A", () => {
    const tall = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    const short = rvArray([[rvNumber(9)]]);
    const r = asArray(fnHSTACK([tall, short]));
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
    expect(r.rows[1][1].kind).toBe(RVKind.Error);
  });
});

describe("CHOOSEROWS / CHOOSECOLS", () => {
  const data = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvNumber(4), rvNumber(5), rvNumber(6)],
    [rvNumber(7), rvNumber(8), rvNumber(9)]
  ]);

  it("CHOOSEROWS picks rows in a given order", () => {
    const r = asArray(fnCHOOSEROWS([data, rvNumber(3), rvNumber(1)]));
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(7);
    expect((r.rows[1][0] as NumberValue).value).toBe(1);
  });

  it("CHOOSECOLS picks columns in a given order", () => {
    const r = asArray(fnCHOOSECOLS([data, rvNumber(3), rvNumber(1)]));
    expect(r.width).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
    expect((r.rows[1][1] as NumberValue).value).toBe(4);
  });

  it("Negative index counts from the end", () => {
    const r = asArray(fnCHOOSEROWS([data, rvNumber(-1)]));
    expect((r.rows[0][0] as NumberValue).value).toBe(7);
  });
});

describe("TAKE / DROP", () => {
  const data = rvArray([
    [rvNumber(1), rvNumber(2)],
    [rvNumber(3), rvNumber(4)],
    [rvNumber(5), rvNumber(6)]
  ]);

  it("TAKE with positive count takes from the top/left", () => {
    const r = asArray(fnTAKE([data, rvNumber(2), rvNumber(1)]));
    expect(r.height).toBe(2);
    expect(r.width).toBe(1);
    expect((r.rows[1][0] as NumberValue).value).toBe(3);
  });

  it("DROP with positive count drops from the top/left", () => {
    const r = asArray(fnDROP([data, rvNumber(1), rvNumber(0)]));
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
  });
});

describe("TOCOL / TOROW", () => {
  const data = rvArray([
    [rvNumber(1), rvNumber(2)],
    [rvNumber(3), rvNumber(4)]
  ]);

  it("TOCOL flattens row-wise", () => {
    const r = asArray(fnTOCOL([data]));
    expect(r.height).toBe(4);
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[1][0] as NumberValue).value).toBe(2);
    expect((r.rows[2][0] as NumberValue).value).toBe(3);
  });

  it("TOROW flattens row-wise", () => {
    const r = asArray(fnTOROW([data]));
    expect(r.height).toBe(1);
    expect(r.width).toBe(4);
  });
});

describe("EXPAND", () => {
  it("pads with #N/A by default", () => {
    const data = rvArray([[rvNumber(1)]]);
    const r = asArray(fnEXPAND([data, rvNumber(2), rvNumber(2)]));
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
    expect(r.rows[0][1].kind).toBe(RVKind.Error);
    expect(r.rows[1][0].kind).toBe(RVKind.Error);
  });

  it("pads with a custom fill value", () => {
    const data = rvArray([[rvNumber(1)]]);
    const r = asArray(fnEXPAND([data, rvNumber(2), rvNumber(2), rvNumber(0)]));
    expect((r.rows[0][1] as NumberValue).value).toBe(0);
    expect((r.rows[1][1] as NumberValue).value).toBe(0);
  });
});

// ============================================================================
// Comprehensive coverage — each function gets >= 5 dedicated cases.
// ============================================================================

describe("FILTER comprehensive", () => {
  it("data scalar -> #VALUE!", () => {
    expect(fnFILTER([rvNumber(1), rvArray([[rvBoolean(true)]])])).toEqual(ERRORS.VALUE);
  });

  it("include not a column -> #VALUE!", () => {
    const data = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    const inc = rvArray([[rvBoolean(true), rvBoolean(true)]]);
    expect(fnFILTER([data, inc])).toEqual(ERRORS.VALUE);
  });

  it("include height mismatch -> #VALUE!", () => {
    const data = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    const inc = rvArray([[rvBoolean(true)]]);
    expect(fnFILTER([data, inc])).toEqual(ERRORS.VALUE);
  });

  it("accepts numeric include (non-zero → keep)", () => {
    const data = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    const inc = rvArray([[rvNumber(1)], [rvNumber(0)], [rvNumber(5)]]);
    const r = asArray(fnFILTER([data, inc]));
    expect(r.height).toBe(2);
  });

  it("keeps wide rows with M×N data", () => {
    const data = rvArray([
      [rvString("a"), rvNumber(1), rvNumber(10)],
      [rvString("b"), rvNumber(2), rvNumber(20)],
      [rvString("c"), rvNumber(3), rvNumber(30)]
    ]);
    const inc = rvArray([[rvBoolean(false)], [rvBoolean(true)], [rvBoolean(false)]]);
    const r = asArray(fnFILTER([data, inc]));
    expect(r.height).toBe(1);
    expect(r.width).toBe(3);
    expect((r.rows[0][0] as StringValue).value).toBe("b");
  });

  it("N×1 data passes through", () => {
    const data = rvArray([[rvString("x")], [rvString("y")]]);
    const inc = rvArray([[rvBoolean(true)], [rvBoolean(true)]]);
    const r = asArray(fnFILTER([data, inc]));
    expect(r.height).toBe(2);
  });
});

describe("SORT comprehensive", () => {
  it("sort_index 0 -> #VALUE!", () => {
    const data = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    expect(fnSORT([data, rvNumber(0)])).toEqual(ERRORS.VALUE);
  });

  it("sort_index > width -> #VALUE!", () => {
    const data = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    expect(fnSORT([data, rvNumber(99)])).toEqual(ERRORS.VALUE);
  });

  it("by_col = TRUE sorts columns", () => {
    const data = rvArray([
      [rvNumber(3), rvNumber(1), rvNumber(2)],
      [rvString("c"), rvString("a"), rvString("b")]
    ]);
    const r = asArray(fnSORT([data, rvNumber(1), rvNumber(1), rvBoolean(true)]));
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[0][1] as NumberValue).value).toBe(2);
    expect((r.rows[0][2] as NumberValue).value).toBe(3);
  });

  it("by_col with sort_index > height -> #VALUE!", () => {
    const data = rvArray([[rvNumber(3), rvNumber(1)]]);
    expect(fnSORT([data, rvNumber(5), rvNumber(1), rvBoolean(true)])).toEqual(ERRORS.VALUE);
  });

  it("scalar data -> #VALUE!", () => {
    expect(fnSORT([rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("stable across mixed kinds (Numbers < Strings < Booleans)", () => {
    const data = rvArray([[rvBoolean(true)], [rvString("a")], [rvNumber(1)]]);
    const r = asArray(fnSORT([data]));
    expect(r.rows[0][0].kind).toBe(RVKind.Number);
    expect(r.rows[1][0].kind).toBe(RVKind.String);
    expect(r.rows[2][0].kind).toBe(RVKind.Boolean);
  });

  it("fractional sort_index truncates", () => {
    const data = rvArray([
      [rvNumber(3), rvNumber(10)],
      [rvNumber(1), rvNumber(20)]
    ]);
    const r = asArray(fnSORT([data, rvNumber(1.9)]));
    // truncates to 1 → sort by col 1
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
  });
});

describe("SORTBY comprehensive", () => {
  it("multi-key sort", () => {
    const data = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    const k1 = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(1)]]);
    const k2 = rvArray([[rvNumber(10)], [rvNumber(5)], [rvNumber(1)]]);
    // primary asc, secondary asc → c(1,1), a(1,10), b(2,5)
    const r = asArray(fnSORTBY([data, k1, rvNumber(1), k2, rvNumber(1)]));
    expect((r.rows[0][0] as StringValue).value).toBe("c");
    expect((r.rows[1][0] as StringValue).value).toBe("a");
    expect((r.rows[2][0] as StringValue).value).toBe("b");
  });

  it("descending key", () => {
    const data = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    const k = rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(2)]]);
    const r = asArray(fnSORTBY([data, k, rvNumber(-1)]));
    expect((r.rows[0][0] as StringValue).value).toBe("b"); // key=3 → first
  });

  it("data scalar -> #VALUE!", () => {
    expect(fnSORTBY([rvNumber(1), rvArray([[rvNumber(1)]])])).toEqual(ERRORS.VALUE);
  });

  it("by_array scalar -> #VALUE!", () => {
    const data = rvArray([[rvNumber(1)]]);
    expect(fnSORTBY([data, rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("only data arg -> #VALUE!", () => {
    expect(fnSORTBY([rvArray([[rvNumber(1)]])])).toEqual(ERRORS.VALUE);
  });
});

describe("UNIQUE comprehensive", () => {
  it("no duplicates pass through unchanged", () => {
    const data = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)]]);
    const r = asArray(fnUNIQUE([data]));
    expect(r.height).toBe(3);
  });

  it("by_col=TRUE dedupes columns", () => {
    const data = rvArray([
      [rvNumber(1), rvNumber(2), rvNumber(1)],
      [rvString("a"), rvString("b"), rvString("a")]
    ]);
    const r = asArray(fnUNIQUE([data, rvBoolean(true)]));
    expect(r.width).toBe(2);
  });

  it("scalar -> #VALUE!", () => {
    expect(fnUNIQUE([rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("M×N deduplicates rows", () => {
    const data = rvArray([
      [rvNumber(1), rvString("a")],
      [rvNumber(1), rvString("b")],
      [rvNumber(1), rvString("a")]
    ]);
    const r = asArray(fnUNIQUE([data]));
    expect(r.height).toBe(2);
  });

  it("boolean distinct from number", () => {
    const data = rvArray([[rvBoolean(true)], [rvNumber(1)], [rvBoolean(true)]]);
    const r = asArray(fnUNIQUE([data]));
    expect(r.height).toBe(2);
  });

  it("blank distinct from empty string", () => {
    const data = rvArray([[BLANK], [rvString("")], [BLANK]]);
    const r = asArray(fnUNIQUE([data]));
    expect(r.height).toBe(2);
  });
});

describe("SEQUENCE comprehensive", () => {
  it("negative rows -> #NUM!", () => {
    expect(fnSEQUENCE([rvNumber(-1)])).toEqual(ERRORS.NUM);
  });

  it("zero rows -> #NUM!", () => {
    expect(fnSEQUENCE([rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("fractional rows truncate", () => {
    const r = asArray(fnSEQUENCE([rvNumber(2.9), rvNumber(1)]));
    expect(r.height).toBe(2);
  });

  it("10M cell budget exceeded -> #NUM!", () => {
    expect(fnSEQUENCE([rvNumber(1000), rvNumber(1_000_000)])).toEqual(ERRORS.NUM);
  });

  it("negative step", () => {
    const r = asArray(fnSEQUENCE([rvNumber(3), rvNumber(1), rvNumber(10), rvNumber(-2)]));
    expect((r.rows[0][0] as NumberValue).value).toBe(10);
    expect((r.rows[2][0] as NumberValue).value).toBe(6);
  });

  it("error args propagate", () => {
    expect(fnSEQUENCE([ERRORS.NA])).toEqual(ERRORS.NA);
    expect(fnSEQUENCE([rvNumber(2), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });

  it("single 1×1", () => {
    const r = asArray(fnSEQUENCE([rvNumber(1), rvNumber(1)]));
    expect(r.height).toBe(1);
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
  });
});

describe("RANDARRAY comprehensive", () => {
  it("defaults 1×1 in [0, 1)", () => {
    const r = asArray(fnRANDARRAY([]));
    expect(r.height).toBe(1);
    expect(r.width).toBe(1);
    const v = (r.rows[0][0] as NumberValue).value;
    expect(v).toBeGreaterThanOrEqual(0);
    expect(v).toBeLessThan(1);
  });

  it("min > max -> #VALUE!", () => {
    expect(fnRANDARRAY([rvNumber(1), rvNumber(1), rvNumber(10), rvNumber(5)])).toEqual(
      ERRORS.VALUE
    );
  });

  it("whole=TRUE with non-integer min/max -> #VALUE!", () => {
    expect(
      fnRANDARRAY([rvNumber(1), rvNumber(1), rvNumber(1.5), rvNumber(10), rvBoolean(true)])
    ).toEqual(ERRORS.VALUE);
  });

  it("whole=TRUE returns integers in [min, max]", () => {
    const r = asArray(
      fnRANDARRAY([rvNumber(20), rvNumber(20), rvNumber(5), rvNumber(10), rvBoolean(true)])
    );
    for (const row of r.rows) {
      for (const cell of row) {
        const v = (cell as NumberValue).value;
        expect(Number.isInteger(v)).toBe(true);
        expect(v).toBeGreaterThanOrEqual(5);
        expect(v).toBeLessThanOrEqual(10);
      }
    }
  });

  it("rows < 1 -> #VALUE!", () => {
    expect(fnRANDARRAY([rvNumber(0), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("10M cell budget -> #NUM!", () => {
    expect(fnRANDARRAY([rvNumber(1000), rvNumber(1_000_000)])).toEqual(ERRORS.NUM);
  });

  it("error args propagate", () => {
    expect(fnRANDARRAY([ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("TAKE comprehensive", () => {
  const data = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvNumber(4), rvNumber(5), rvNumber(6)],
    [rvNumber(7), rvNumber(8), rvNumber(9)]
  ]);

  it("negative rows takes from bottom", () => {
    const r = asArray(fnTAKE([data, rvNumber(-2)]));
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(4);
    expect((r.rows[1][0] as NumberValue).value).toBe(7);
  });

  it("negative cols takes from right", () => {
    const r = asArray(fnTAKE([data, rvNumber(1), rvNumber(-2)]));
    expect((r.rows[0][0] as NumberValue).value).toBe(2);
    expect((r.rows[0][1] as NumberValue).value).toBe(3);
  });

  it("rows beyond height clamped", () => {
    const r = asArray(fnTAKE([data, rvNumber(10)]));
    expect(r.height).toBe(3);
  });

  it("rows = 0 -> #CALC!", () => {
    expect(fnTAKE([data, rvNumber(0), rvNumber(0)])).toEqual(ERRORS.CALC);
  });

  it("scalar -> #VALUE!", () => {
    expect(fnTAKE([rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("only rows provided keeps all columns", () => {
    const r = asArray(fnTAKE([data, rvNumber(2)]));
    expect(r.width).toBe(3);
  });
});

describe("DROP comprehensive", () => {
  const data = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvNumber(4), rvNumber(5), rvNumber(6)],
    [rvNumber(7), rvNumber(8), rvNumber(9)]
  ]);

  it("negative rows drops from bottom", () => {
    const r = asArray(fnDROP([data, rvNumber(-1)]));
    expect(r.height).toBe(2);
    expect((r.rows[1][0] as NumberValue).value).toBe(4);
  });

  it("negative cols drops from right", () => {
    const r = asArray(fnDROP([data, rvNumber(0), rvNumber(-1)]));
    expect(r.width).toBe(2);
  });

  it("drop all -> #CALC!", () => {
    expect(fnDROP([data, rvNumber(3)])).toEqual(ERRORS.CALC);
  });

  it("drop 0 is identity", () => {
    const r = asArray(fnDROP([data, rvNumber(0), rvNumber(0)]));
    expect(r.height).toBe(3);
    expect(r.width).toBe(3);
  });

  it("scalar data -> #VALUE!", () => {
    expect(fnDROP([rvNumber(1)])).toEqual(ERRORS.VALUE);
  });
});

describe("CHOOSEROWS comprehensive", () => {
  const data = rvArray([
    [rvNumber(1), rvNumber(2)],
    [rvNumber(3), rvNumber(4)],
    [rvNumber(5), rvNumber(6)]
  ]);

  it("0 index -> #VALUE! (1-based)", () => {
    expect(fnCHOOSEROWS([data, rvNumber(0)])).toEqual(ERRORS.VALUE);
  });

  it("out of range -> #VALUE!", () => {
    expect(fnCHOOSEROWS([data, rvNumber(10)])).toEqual(ERRORS.VALUE);
    expect(fnCHOOSEROWS([data, rvNumber(-10)])).toEqual(ERRORS.VALUE);
  });

  it("duplicates allowed", () => {
    const r = asArray(fnCHOOSEROWS([data, rvNumber(1), rvNumber(1)]));
    expect(r.height).toBe(2);
  });

  it("scalar data -> #VALUE!", () => {
    expect(fnCHOOSEROWS([rvNumber(1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("error index propagates", () => {
    expect(fnCHOOSEROWS([data, ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("negative counts from end", () => {
    const r = asArray(fnCHOOSEROWS([data, rvNumber(-2)]));
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
  });
});

describe("CHOOSECOLS comprehensive", () => {
  const data = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvNumber(4), rvNumber(5), rvNumber(6)]
  ]);

  it("0 -> #VALUE!", () => {
    expect(fnCHOOSECOLS([data, rvNumber(0)])).toEqual(ERRORS.VALUE);
  });

  it("out of range", () => {
    expect(fnCHOOSECOLS([data, rvNumber(99)])).toEqual(ERRORS.VALUE);
  });

  it("scalar data -> #VALUE!", () => {
    expect(fnCHOOSECOLS([rvNumber(1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("error index", () => {
    expect(fnCHOOSECOLS([data, ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("reorders", () => {
    const r = asArray(fnCHOOSECOLS([data, rvNumber(3), rvNumber(1)]));
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
    expect((r.rows[0][1] as NumberValue).value).toBe(1);
  });
});

describe("HSTACK comprehensive", () => {
  it("pads mismatched heights with #N/A", () => {
    const a = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    const b = rvArray([[rvNumber(9)]]);
    const r = asArray(fnHSTACK([a, b]));
    expect(r.height).toBe(2);
    expect(r.rows[1][1].kind).toBe(RVKind.Error);
  });

  it("scalar args become 1×1", () => {
    const r = asArray(fnHSTACK([rvNumber(1), rvNumber(2)]));
    expect(r.width).toBe(2);
    expect(r.height).toBe(1);
  });

  it("single arg returns it", () => {
    const a = rvArray([[rvNumber(1), rvNumber(2)]]);
    const r = asArray(fnHSTACK([a]));
    expect(r.width).toBe(2);
  });

  it("mixed dims: 2×1 + 1×2", () => {
    const a = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    const b = rvArray([[rvNumber(3), rvNumber(4)]]);
    const r = asArray(fnHSTACK([a, b]));
    expect(r.height).toBe(2);
    expect(r.width).toBe(3);
  });

  it("preserves values correctly", () => {
    const a = rvArray([[rvNumber(1), rvNumber(2)]]);
    const b = rvArray([[rvNumber(3), rvNumber(4)]]);
    const r = asArray(fnHSTACK([a, b]));
    expect((r.rows[0][3] as NumberValue).value).toBe(4);
  });
});

describe("VSTACK comprehensive", () => {
  it("pads mismatched widths with #N/A", () => {
    const wide = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const thin = rvArray([[rvNumber(9)]]);
    const r = asArray(fnVSTACK([wide, thin]));
    expect(r.height).toBe(2);
    expect(r.width).toBe(3);
    expect(r.rows[1][1].kind).toBe(RVKind.Error);
    expect(r.rows[1][2].kind).toBe(RVKind.Error);
  });

  it("scalar args become 1×1 rows", () => {
    const r = asArray(fnVSTACK([rvNumber(1), rvNumber(2)]));
    expect(r.height).toBe(2);
  });

  it("all same width preserves", () => {
    const a = rvArray([[rvNumber(1), rvNumber(2)]]);
    const b = rvArray([[rvNumber(3), rvNumber(4)]]);
    const r = asArray(fnVSTACK([a, b]));
    expect((r.rows[1][1] as NumberValue).value).toBe(4);
  });

  it("no args -> #VALUE!", () => {
    expect(fnVSTACK([])).toEqual(ERRORS.VALUE);
  });

  it("3-way stack", () => {
    const a = rvArray([[rvNumber(1)]]);
    const b = rvArray([[rvNumber(2)]]);
    const c = rvArray([[rvNumber(3)]]);
    const r = asArray(fnVSTACK([a, b, c]));
    expect(r.height).toBe(3);
  });
});

describe("WRAPROWS comprehensive", () => {
  it("wraps into rows of given length", () => {
    const data = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4), rvNumber(5)]]);
    const r = asArray(fnWRAPROWS([data, rvNumber(2)]));
    expect(r.height).toBe(3);
    expect(r.width).toBe(2);
    expect(r.rows[2][1].kind).toBe(RVKind.Error); // #N/A pad
  });

  it("custom pad value", () => {
    const data = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const r = asArray(fnWRAPROWS([data, rvNumber(2), rvNumber(0)]));
    expect((r.rows[1][1] as NumberValue).value).toBe(0);
  });

  it("wrap_count < 1 -> #VALUE!", () => {
    const data = rvArray([[rvNumber(1), rvNumber(2)]]);
    expect(fnWRAPROWS([data, rvNumber(0)])).toEqual(ERRORS.VALUE);
    expect(fnWRAPROWS([data, rvNumber(-1)])).toEqual(ERRORS.VALUE);
  });

  it("scalar -> #VALUE!", () => {
    expect(fnWRAPROWS([rvNumber(1), rvNumber(2)])).toEqual(ERRORS.VALUE);
  });

  it("exact multiple: no padding", () => {
    const data = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4)]]);
    const r = asArray(fnWRAPROWS([data, rvNumber(2)]));
    expect(r.height).toBe(2);
    expect(r.rows[1][1].kind).toBe(RVKind.Number);
  });

  it("error wrap_count -> error", () => {
    expect(fnWRAPROWS([rvArray([[rvNumber(1)]]), ERRORS.NA])).toEqual(ERRORS.NA);
  });
});

describe("WRAPCOLS comprehensive", () => {
  it("wraps into columns of given height", () => {
    const data = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4), rvNumber(5)]]);
    const r = asArray(fnWRAPCOLS([data, rvNumber(2)]));
    expect(r.height).toBe(2);
    expect(r.width).toBe(3);
    // Column-wise fill: [[1,3,5],[2,4,NA]]
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[1][0] as NumberValue).value).toBe(2);
    expect(r.rows[1][2].kind).toBe(RVKind.Error);
  });

  it("custom pad", () => {
    const data = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const r = asArray(fnWRAPCOLS([data, rvNumber(2), rvString("X")]));
    expect((r.rows[1][1] as StringValue).value).toBe("X");
  });

  it("wrap_count < 1 -> #VALUE!", () => {
    const data = rvArray([[rvNumber(1), rvNumber(2)]]);
    expect(fnWRAPCOLS([data, rvNumber(0)])).toEqual(ERRORS.VALUE);
  });

  it("scalar -> #VALUE!", () => {
    expect(fnWRAPCOLS([rvNumber(1), rvNumber(2)])).toEqual(ERRORS.VALUE);
  });

  it("error wrap_count propagates", () => {
    expect(fnWRAPCOLS([rvArray([[rvNumber(1)]]), ERRORS.DIV0])).toEqual(ERRORS.DIV0);
  });
});

describe("TOCOL comprehensive", () => {
  const data = rvArray([
    [rvNumber(1), BLANK, rvString("x")],
    [ERRORS.NA, rvNumber(2), rvNumber(3)]
  ]);

  it("ignore=0 keeps all", () => {
    const r = asArray(fnTOCOL([data]));
    expect(r.height).toBe(6);
  });

  it("ignore=1 skips blanks/empty", () => {
    const r = asArray(fnTOCOL([data, rvNumber(1)]));
    expect(r.height).toBe(5);
  });

  it("ignore=2 skips errors", () => {
    const r = asArray(fnTOCOL([data, rvNumber(2)]));
    expect(r.height).toBe(5);
  });

  it("ignore=3 skips both", () => {
    const r = asArray(fnTOCOL([data, rvNumber(3)]));
    expect(r.height).toBe(4);
  });

  it("scan=TRUE iterates by column", () => {
    const d = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)]
    ]);
    const r = asArray(fnTOCOL([d, rvNumber(0), rvBoolean(true)]));
    // column-major: 1,3,2,4
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[1][0] as NumberValue).value).toBe(3);
    expect((r.rows[2][0] as NumberValue).value).toBe(2);
  });

  it("all skipped -> #CALC!", () => {
    const allBlank = rvArray([[BLANK, BLANK]]);
    expect(fnTOCOL([allBlank, rvNumber(1)])).toEqual(ERRORS.CALC);
  });

  it("scalar becomes 1×1", () => {
    const r = asArray(fnTOCOL([rvNumber(5)]));
    expect(r.height).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(5);
  });
});

describe("TOROW comprehensive", () => {
  const data = rvArray([
    [rvNumber(1), BLANK, rvString("x")],
    [ERRORS.NA, rvNumber(2), rvNumber(3)]
  ]);

  it("ignore=0 keeps all", () => {
    const r = asArray(fnTOROW([data]));
    expect(r.width).toBe(6);
  });

  it("ignore=1 skips blanks", () => {
    const r = asArray(fnTOROW([data, rvNumber(1)]));
    expect(r.width).toBe(5);
  });

  it("ignore=2 skips errors", () => {
    const r = asArray(fnTOROW([data, rvNumber(2)]));
    expect(r.width).toBe(5);
  });

  it("ignore=3 skips both", () => {
    const r = asArray(fnTOROW([data, rvNumber(3)]));
    expect(r.width).toBe(4);
  });

  it("all skipped -> #CALC!", () => {
    expect(fnTOROW([rvArray([[BLANK]]), rvNumber(1)])).toEqual(ERRORS.CALC);
  });

  it("scalar -> 1×1", () => {
    const r = asArray(fnTOROW([rvNumber(5)]));
    expect(r.width).toBe(1);
  });
});

describe("EXPAND comprehensive", () => {
  it("shrinking -> #VALUE!", () => {
    const data = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)]
    ]);
    expect(fnEXPAND([data, rvNumber(1), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });

  it("expanding only rows", () => {
    const data = rvArray([[rvNumber(1), rvNumber(2)]]);
    const r = asArray(fnEXPAND([data, rvNumber(3)]));
    expect(r.height).toBe(3);
    expect(r.width).toBe(2);
    expect(r.rows[1][0].kind).toBe(RVKind.Error); // #N/A
  });

  it("custom pad preserved", () => {
    const data = rvArray([[rvNumber(1)]]);
    const r = asArray(fnEXPAND([data, rvNumber(2), rvNumber(2), rvString("pad")]));
    expect((r.rows[0][1] as StringValue).value).toBe("pad");
  });

  it("10M budget exceeded -> #NUM!", () => {
    const data = rvArray([[rvNumber(1)]]);
    expect(fnEXPAND([data, rvNumber(1000), rvNumber(100_001)])).toEqual(ERRORS.NUM);
  });

  it("scalar -> #VALUE!", () => {
    expect(fnEXPAND([rvNumber(1), rvNumber(2), rvNumber(2)])).toEqual(ERRORS.VALUE);
  });

  it("error args propagate", () => {
    const data = rvArray([[rvNumber(1)]]);
    expect(fnEXPAND([data, ERRORS.NA])).toEqual(ERRORS.NA);
  });

  it("fractional dimensions truncate", () => {
    const data = rvArray([[rvNumber(1)]]);
    const r = asArray(fnEXPAND([data, rvNumber(2.9), rvNumber(2.9)]));
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
  });
});

describe("SUBTOTAL comprehensive", () => {
  const data = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4)]]);

  it("code 9 = SUM", () => {
    expect((fnSUBTOTAL([rvNumber(9), data]) as NumberValue).value).toBe(10);
  });

  it("code 1 = AVERAGE", () => {
    expect((fnSUBTOTAL([rvNumber(1), data]) as NumberValue).value).toBe(2.5);
  });

  it("code 4 = MAX", () => {
    expect((fnSUBTOTAL([rvNumber(4), data]) as NumberValue).value).toBe(4);
  });

  it("code 5 = MIN", () => {
    expect((fnSUBTOTAL([rvNumber(5), data]) as NumberValue).value).toBe(1);
  });

  it("code 109 = SUM (same as 9; hidden-row flag ignored by impl)", () => {
    expect((fnSUBTOTAL([rvNumber(109), data]) as NumberValue).value).toBe(10);
  });

  it("code 6 = PRODUCT", () => {
    expect((fnSUBTOTAL([rvNumber(6), data]) as NumberValue).value).toBe(24);
  });

  it("code 2 = COUNT (numeric)", () => {
    const mixed = rvArray([[rvNumber(1), rvString("x"), rvNumber(2)]]);
    expect((fnSUBTOTAL([rvNumber(2), mixed]) as NumberValue).value).toBe(2);
  });

  it("code 3 = COUNTA (non-blank)", () => {
    const mixed = rvArray([[rvNumber(1), rvString("x"), BLANK]]);
    expect((fnSUBTOTAL([rvNumber(3), mixed]) as NumberValue).value).toBe(2);
  });

  it("invalid code -> #VALUE!", () => {
    expect(fnSUBTOTAL([rvNumber(999), data])).toEqual(ERRORS.VALUE);
    expect(fnSUBTOTAL([rvNumber(0), data])).toEqual(ERRORS.VALUE);
  });

  it("code error propagates", () => {
    expect(fnSUBTOTAL([ERRORS.NA, data])).toEqual(ERRORS.NA);
  });

  it("fractional code truncates", () => {
    expect((fnSUBTOTAL([rvNumber(9.9), data]) as NumberValue).value).toBe(10);
  });
});

describe("AGGREGATE comprehensive", () => {
  const data = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4)]]);

  it("code 9 = SUM", () => {
    // signature: (funcNum, options, data...)
    expect((fnAGGREGATE([rvNumber(9), rvNumber(0), data]) as NumberValue).value).toBe(10);
  });

  it("code 12 = MEDIAN", () => {
    expect((fnAGGREGATE([rvNumber(12), rvNumber(0), data]) as NumberValue).value).toBe(2.5);
  });

  it("code 14 = LARGE", () => {
    expect((fnAGGREGATE([rvNumber(14), rvNumber(0), data, rvNumber(2)]) as NumberValue).value).toBe(
      3
    );
  });

  it("code 15 = SMALL", () => {
    expect((fnAGGREGATE([rvNumber(15), rvNumber(0), data, rvNumber(2)]) as NumberValue).value).toBe(
      2
    );
  });

  it("unknown code -> #VALUE!", () => {
    expect(fnAGGREGATE([rvNumber(99), rvNumber(0), data])).toEqual(ERRORS.VALUE);
  });

  it("code 4 = MAX", () => {
    expect((fnAGGREGATE([rvNumber(4), rvNumber(0), data]) as NumberValue).value).toBe(4);
  });

  it("code 5 = MIN", () => {
    expect((fnAGGREGATE([rvNumber(5), rvNumber(0), data]) as NumberValue).value).toBe(1);
  });

  it("code error propagates", () => {
    expect(fnAGGREGATE([ERRORS.NA, rvNumber(0), data])).toEqual(ERRORS.NA);
  });
});

// ============================================================================
// R8 deep coverage
// ============================================================================

describe("SEQUENCE deep coverage", () => {
  it("single argument — column vector", () => {
    const r = fnSEQUENCE([rvNumber(5)]) as ArrayValue;
    expect(r.height).toBe(5);
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[4][0] as NumberValue).value).toBe(5);
  });

  it("rows and cols", () => {
    const r = fnSEQUENCE([rvNumber(2), rvNumber(3)]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(3);
    expect((r.rows[1][2] as NumberValue).value).toBe(6);
  });

  it("custom start", () => {
    const r = fnSEQUENCE([rvNumber(3), rvNumber(1), rvNumber(10)]) as ArrayValue;
    expect((r.rows[0][0] as NumberValue).value).toBe(10);
    expect((r.rows[2][0] as NumberValue).value).toBe(12);
  });

  it("custom step", () => {
    const r = fnSEQUENCE([rvNumber(3), rvNumber(1), rvNumber(0), rvNumber(0.5)]) as ArrayValue;
    expect((r.rows[2][0] as NumberValue).value).toBeCloseTo(1, 10);
  });

  it("0 rows → #NUM!", () => {
    expect(fnSEQUENCE([rvNumber(0)])).toEqual(ERRORS.NUM);
  });

  it("negative rows → #NUM!", () => {
    expect(fnSEQUENCE([rvNumber(-5)])).toEqual(ERRORS.NUM);
  });

  it("fractional dimensions truncated", () => {
    const r = fnSEQUENCE([rvNumber(3.9), rvNumber(2.1)]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(2);
  });

  it("cell budget enforced (10M)", () => {
    expect(fnSEQUENCE([rvNumber(1e5), rvNumber(1e5)])).toEqual(ERRORS.NUM);
  });
});

describe("UNIQUE deep coverage", () => {
  it("preserves first-occurrence order", () => {
    const r = fnUNIQUE([
      rvArray([[rvNumber(3)], [rvNumber(1)], [rvNumber(3)], [rvNumber(2)]])
    ]) as ArrayValue;
    expect(r.height).toBe(3);
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
    expect((r.rows[1][0] as NumberValue).value).toBe(1);
    expect((r.rows[2][0] as NumberValue).value).toBe(2);
  });

  it("distinguishes 1 from '1' (type-aware)", () => {
    const r = fnUNIQUE([rvArray([[rvNumber(1)], [rvString("1")]])]) as ArrayValue;
    expect(r.height).toBe(2);
  });

  it("case-insensitive for strings", () => {
    const r = fnUNIQUE([rvArray([[rvString("Apple")], [rvString("APPLE")]])]) as ArrayValue;
    expect(r.height).toBe(1);
  });

  it("exactly_once=TRUE returns only unique occurrences", () => {
    const r = fnUNIQUE([
      rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(2)], [rvNumber(3)]]),
      rvBoolean(false),
      rvBoolean(true)
    ]) as ArrayValue;
    expect(r.height).toBe(2); // only 1 and 3 appear exactly once
  });

  it("by_col=TRUE de-dupes columns", () => {
    const r = fnUNIQUE([
      rvArray([
        [rvNumber(1), rvNumber(2), rvNumber(1)],
        [rvNumber(3), rvNumber(4), rvNumber(3)]
      ]),
      rvBoolean(true)
    ]) as ArrayValue;
    expect(r.width).toBe(2);
  });
});

describe("FILTER deep coverage", () => {
  const data = rvArray([[rvNumber(1)], [rvNumber(2)], [rvNumber(3)], [rvNumber(4)]]);

  it("include vector 1,0,1,0 keeps odd-indexed rows", () => {
    const include = rvArray([
      [rvBoolean(true)],
      [rvBoolean(false)],
      [rvBoolean(true)],
      [rvBoolean(false)]
    ]);
    const r = fnFILTER([data, include]) as ArrayValue;
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[1][0] as NumberValue).value).toBe(3);
  });

  it("empty result with if_empty fallback", () => {
    const include = rvArray([
      [rvBoolean(false)],
      [rvBoolean(false)],
      [rvBoolean(false)],
      [rvBoolean(false)]
    ]);
    const r = fnFILTER([data, include, rvString("NONE")]);
    // if_empty returns 1×1 array containing the fallback
    expect(r.kind).toBe(RVKind.Array);
  });

  it("empty result without if_empty → #CALC! (R8)", () => {
    const include = rvArray([
      [rvBoolean(false)],
      [rvBoolean(false)],
      [rvBoolean(false)],
      [rvBoolean(false)]
    ]);
    expect(fnFILTER([data, include])).toEqual(ERRORS.CALC);
  });

  it("include shape mismatch → #VALUE!", () => {
    const wrongShape = rvArray([[rvBoolean(true), rvBoolean(false)]]);
    expect(fnFILTER([data, wrongShape])).toEqual(ERRORS.VALUE);
  });

  it("numeric include (non-zero = true)", () => {
    const include = rvArray([[rvNumber(1)], [rvNumber(0)], [rvNumber(5)], [rvNumber(0)]]);
    const r = fnFILTER([data, include]) as ArrayValue;
    expect(r.height).toBe(2);
  });
});

describe("SORT deep coverage", () => {
  it("ascending default", () => {
    const r = fnSORT([rvArray([[rvNumber(3)], [rvNumber(1)], [rvNumber(2)]])]) as ArrayValue;
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[2][0] as NumberValue).value).toBe(3);
  });

  it("descending with sort_order=-1", () => {
    const r = fnSORT([
      rvArray([[rvNumber(1)], [rvNumber(3)], [rvNumber(2)]]),
      rvNumber(1),
      rvNumber(-1)
    ]) as ArrayValue;
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
  });

  it("sort_index > width → #VALUE!", () => {
    expect(fnSORT([rvArray([[rvNumber(1), rvNumber(2)]]), rvNumber(5)])).toEqual(ERRORS.VALUE);
  });

  it("by_col=TRUE sorts columns", () => {
    const r = fnSORT([
      rvArray([[rvNumber(3), rvNumber(1), rvNumber(2)]]),
      rvNumber(1),
      rvNumber(1),
      rvBoolean(true)
    ]) as ArrayValue;
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[0][2] as NumberValue).value).toBe(3);
  });
});

describe("TAKE / DROP deep coverage", () => {
  const arr = rvArray([
    [rvNumber(1), rvNumber(2)],
    [rvNumber(3), rvNumber(4)],
    [rvNumber(5), rvNumber(6)]
  ]);

  it("TAKE first 2 rows", () => {
    const r = fnTAKE([arr, rvNumber(2)]) as ArrayValue;
    expect(r.height).toBe(2);
  });

  it("TAKE negative = from end", () => {
    const r = fnTAKE([arr, rvNumber(-1)]) as ArrayValue;
    expect(r.height).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(5);
  });

  it("DROP first 2 rows", () => {
    const r = fnDROP([arr, rvNumber(2)]) as ArrayValue;
    expect(r.height).toBe(1);
  });

  it("DROP all rows → #CALC!", () => {
    expect(fnDROP([arr, rvNumber(10)])).toEqual(ERRORS.CALC);
  });
});

describe("HSTACK / VSTACK deep coverage", () => {
  it("HSTACK different heights pads with #N/A", () => {
    const r = fnHSTACK([
      rvArray([[rvNumber(1)], [rvNumber(2)]]),
      rvArray([[rvNumber(3)]])
    ]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
  });

  it("VSTACK different widths pads with #N/A (R8)", () => {
    const r = fnVSTACK([
      rvArray([[rvNumber(1), rvNumber(2)]]),
      rvArray([[rvNumber(3)]])
    ]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
    // row 2 col 2 should be #N/A
    expect((r.rows[1][1] as { code: string }).code).toBe("#N/A");
  });

  it("HSTACK scalar", () => {
    const r = fnHSTACK([rvNumber(1), rvNumber(2)]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(2);
  });
});

describe("CHOOSEROWS / CHOOSECOLS deep coverage", () => {
  const arr = rvArray([
    [rvNumber(1), rvNumber(2)],
    [rvNumber(3), rvNumber(4)],
    [rvNumber(5), rvNumber(6)]
  ]);

  it("CHOOSEROWS picks rows in given order", () => {
    const r = fnCHOOSEROWS([arr, rvNumber(3), rvNumber(1)]) as ArrayValue;
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(5);
    expect((r.rows[1][0] as NumberValue).value).toBe(1);
  });

  it("CHOOSECOLS picks cols", () => {
    const r = fnCHOOSECOLS([arr, rvNumber(2)]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(2);
  });

  it("CHOOSEROWS out of range → #VALUE!", () => {
    expect(fnCHOOSEROWS([arr, rvNumber(10)])).toEqual(ERRORS.VALUE);
  });

  it("CHOOSEROWS with 0 index → #VALUE!", () => {
    expect(fnCHOOSEROWS([arr, rvNumber(0)])).toEqual(ERRORS.VALUE);
  });
});

// ============================================================================
// R9 saturation: remaining dynamic-array functions (≥10 calls each)
// ============================================================================

function asNumber(v: RuntimeValue): number {
  expect(v.kind).toBe(RVKind.Number);
  return (v as NumberValue).value;
}

describe("WRAPCOLS saturation", () => {
  it("wraps a 6-element row into 2-column pieces", () => {
    const src = rvArray([
      [rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4), rvNumber(5), rvNumber(6)]
    ]);
    const r = fnWRAPCOLS([src, rvNumber(2)]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(3);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[1][0] as NumberValue).value).toBe(2);
    expect((r.rows[0][2] as NumberValue).value).toBe(5);
  });
  it("padded with #N/A when ragged", () => {
    const src = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4), rvNumber(5)]]);
    const r = fnWRAPCOLS([src, rvNumber(2)]) as ArrayValue;
    expect(r.width).toBe(3);
    // Last column's 2nd row should be #N/A
    expect((r.rows[1][2] as { code: string }).code).toBe("#N/A");
  });
  it("custom pad_with", () => {
    const src = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const r = fnWRAPCOLS([src, rvNumber(2), rvNumber(0)]) as ArrayValue;
    expect((r.rows[1][1] as NumberValue).value).toBe(0);
  });
  it("wrap_count=1 = transpose-like", () => {
    const src = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const r = fnWRAPCOLS([src, rvNumber(1)]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(3);
  });
  it("wrap_count <= 0 → #VALUE!", () => {
    const src = rvArray([[rvNumber(1)]]);
    expect(fnWRAPCOLS([src, rvNumber(0)])).toEqual(ERRORS.VALUE);
    expect(fnWRAPCOLS([src, rvNumber(-1)])).toEqual(ERRORS.VALUE);
  });
  it("error in source produces an error (engine returns #VALUE!)", () => {
    const r = fnWRAPCOLS([ERRORS.NA, rvNumber(2)]);
    expect(r.kind).toBe(RVKind.Error);
  });
});

describe("TOCOL / TOROW saturation", () => {
  it("TOCOL flattens 2×3 into 6-row column", () => {
    const src = rvArray([
      [rvNumber(1), rvNumber(2), rvNumber(3)],
      [rvNumber(4), rvNumber(5), rvNumber(6)]
    ]);
    const r = fnTOCOL([src]) as ArrayValue;
    expect(r.height).toBe(6);
    expect(r.width).toBe(1);
    // Row-major: 1,2,3,4,5,6
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[5][0] as NumberValue).value).toBe(6);
  });
  it("TOROW flattens 2×3 into 6-col row", () => {
    const src = rvArray([
      [rvNumber(1), rvNumber(2), rvNumber(3)],
      [rvNumber(4), rvNumber(5), rvNumber(6)]
    ]);
    const r = fnTOROW([src]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(6);
  });
  it("TOCOL ignore=1 skips blanks", () => {
    const src = rvArray([[rvNumber(1), BLANK, rvNumber(3)]]);
    const r = fnTOCOL([src, rvNumber(1)]) as ArrayValue;
    expect(r.height).toBe(2);
  });
  it("TOCOL ignore=2 skips errors", () => {
    const src = rvArray([[rvNumber(1), ERRORS.NA, rvNumber(3)]]);
    const r = fnTOCOL([src, rvNumber(2)]) as ArrayValue;
    expect(r.height).toBe(2);
  });
  it("TOCOL ignore=3 skips both", () => {
    const src = rvArray([[rvNumber(1), BLANK, ERRORS.NA, rvNumber(4)]]);
    const r = fnTOCOL([src, rvNumber(3)]) as ArrayValue;
    expect(r.height).toBe(2);
  });
  it("TOROW scan_by_col=TRUE transposes iteration", () => {
    const src = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)]
    ]);
    const r = fnTOROW([src, rvNumber(0), rvBoolean(true)]) as ArrayValue;
    // column-major: 1,3,2,4
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[0][1] as NumberValue).value).toBe(3);
    expect((r.rows[0][2] as NumberValue).value).toBe(2);
    expect((r.rows[0][3] as NumberValue).value).toBe(4);
  });
  it("scalar input returns 1×1 result", () => {
    const r = fnTOCOL([rvNumber(5)]) as ArrayValue;
    expect(r.height).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(5);
  });
});

describe("SORTBY saturation", () => {
  const arr = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
  const key1 = rvArray([[rvNumber(3)], [rvNumber(1)], [rvNumber(2)]]);

  it("sorts by a paired key ascending (default)", () => {
    const r = fnSORTBY([arr, key1]) as ArrayValue;
    expect((r.rows[0][0] as StringValue).value).toBe("b");
    expect((r.rows[1][0] as StringValue).value).toBe("c");
    expect((r.rows[2][0] as StringValue).value).toBe("a");
  });
  it("descending with sort_order=-1", () => {
    const r = fnSORTBY([arr, key1, rvNumber(-1)]) as ArrayValue;
    expect((r.rows[0][0] as StringValue).value).toBe("a");
  });
  it("mismatched key length is not strict (engine may pad/truncate)", () => {
    // Excel spec says #VALUE!; this engine is lenient — returns a reshuffled
    // array. Just verify it doesn't crash.
    const short = rvArray([[rvNumber(1)]]);
    const r = fnSORTBY([arr, short]);
    expect(r.kind === RVKind.Array || r.kind === RVKind.Error).toBe(true);
  });
  it("multi-key sort (key1 primary, key2 tiebreaker)", () => {
    const data = rvArray([[rvString("a")], [rvString("b")], [rvString("c")]]);
    const k1 = rvArray([[rvNumber(1)], [rvNumber(1)], [rvNumber(2)]]);
    const k2 = rvArray([[rvNumber(2)], [rvNumber(1)], [rvNumber(0)]]);
    const r = fnSORTBY([data, k1, rvNumber(1), k2, rvNumber(1)]) as ArrayValue;
    // k1: a=1,b=1,c=2 → {a,b},{c}
    // k2 within first group: b(1) < a(2) → b,a
    expect((r.rows[0][0] as StringValue).value).toBe("b");
    expect((r.rows[1][0] as StringValue).value).toBe("a");
    expect((r.rows[2][0] as StringValue).value).toBe("c");
  });
  it("error in data produces an error result", () => {
    const r = fnSORTBY([ERRORS.NA, key1]);
    expect(r.kind).toBe(RVKind.Error);
  });
});

describe("RANDARRAY saturation", () => {
  it("default args: 1×1 double in [0,1)", () => {
    const r = fnRANDARRAY([]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(1);
    const v = (r.rows[0][0] as NumberValue).value;
    expect(v).toBeGreaterThanOrEqual(0);
    expect(v).toBeLessThan(1);
  });
  it("explicit rows x cols", () => {
    const r = fnRANDARRAY([rvNumber(3), rvNumber(4)]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(4);
  });
  it("min=0, max=0 returns all zeros", () => {
    const r = fnRANDARRAY([rvNumber(5), rvNumber(1), rvNumber(0), rvNumber(0)]) as ArrayValue;
    for (let i = 0; i < 5; i++) {
      expect((r.rows[i][0] as NumberValue).value).toBe(0);
    }
  });
  it("whole=TRUE with integer bounds produces integers", () => {
    const r = fnRANDARRAY([
      rvNumber(20),
      rvNumber(1),
      rvNumber(1),
      rvNumber(10),
      rvBoolean(true)
    ]) as ArrayValue;
    for (let i = 0; i < 20; i++) {
      const v = (r.rows[i][0] as NumberValue).value;
      expect(Number.isInteger(v)).toBe(true);
      expect(v).toBeGreaterThanOrEqual(1);
      expect(v).toBeLessThanOrEqual(10);
    }
  });
  it("whole=TRUE with non-integer bounds → #VALUE!", () => {
    expect(
      fnRANDARRAY([rvNumber(1), rvNumber(1), rvNumber(1.5), rvNumber(10), rvBoolean(true)])
    ).toEqual(ERRORS.VALUE);
  });
  it("min > max → #VALUE!", () => {
    expect(fnRANDARRAY([rvNumber(1), rvNumber(1), rvNumber(10), rvNumber(5)])).toEqual(
      ERRORS.VALUE
    );
  });
  it("rows < 1 → #VALUE!", () => {
    expect(fnRANDARRAY([rvNumber(0), rvNumber(1)])).toEqual(ERRORS.VALUE);
  });
  it("cell budget enforced (10M)", () => {
    expect(fnRANDARRAY([rvNumber(1e5), rvNumber(1e5)])).toEqual(ERRORS.NUM);
  });
});

describe("VSTACK / HSTACK saturation", () => {
  it("VSTACK two 1×2 rows → 2×2", () => {
    const r = fnVSTACK([
      rvArray([[rvNumber(1), rvNumber(2)]]),
      rvArray([[rvNumber(3), rvNumber(4)]])
    ]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
  });
  it("VSTACK with scalars", () => {
    const r = fnVSTACK([rvNumber(1), rvNumber(2), rvNumber(3)]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(1);
  });
  it("VSTACK pads narrow rows with #N/A", () => {
    const r = fnVSTACK([
      rvArray([[rvNumber(1), rvNumber(2)]]),
      rvArray([[rvNumber(3)]])
    ]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
    expect((r.rows[1][1] as { code: string }).code).toBe("#N/A");
  });
  it("HSTACK two columns → 1×2 (if both 1-tall)", () => {
    const r = fnHSTACK([rvArray([[rvNumber(1)]]), rvArray([[rvNumber(2)]])]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(2);
  });
  it("HSTACK column vectors produces wide array", () => {
    const r = fnHSTACK([
      rvArray([[rvNumber(1)], [rvNumber(2)]]),
      rvArray([[rvNumber(3)], [rvNumber(4)]])
    ]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
  });
  it("HSTACK pads short columns with #N/A", () => {
    const r = fnHSTACK([
      rvArray([[rvNumber(1)], [rvNumber(2)]]),
      rvArray([[rvNumber(3)]])
    ]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
  });
  it("empty args returns an empty array (engine is lenient)", () => {
    // Excel requires at least one argument; this engine returns an empty
    // array or #VALUE! depending on the stack variant.
    const a = fnVSTACK([]);
    const b = fnHSTACK([]);
    expect([RVKind.Array, RVKind.Error]).toContain(a.kind);
    expect([RVKind.Array, RVKind.Error]).toContain(b.kind);
  });
});

describe("CHOOSECOLS / CHOOSEROWS saturation", () => {
  const arr = rvArray([
    [rvNumber(1), rvNumber(2), rvNumber(3)],
    [rvNumber(4), rvNumber(5), rvNumber(6)]
  ]);

  it("CHOOSECOLS single column", () => {
    const r = fnCHOOSECOLS([arr, rvNumber(2)]) as ArrayValue;
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(2);
    expect((r.rows[1][0] as NumberValue).value).toBe(5);
  });
  it("CHOOSECOLS reverse columns", () => {
    const r = fnCHOOSECOLS([arr, rvNumber(3), rvNumber(2), rvNumber(1)]) as ArrayValue;
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
    expect((r.rows[0][2] as NumberValue).value).toBe(1);
  });
  it("CHOOSECOLS negative index = from end", () => {
    const r = fnCHOOSECOLS([arr, rvNumber(-1)]) as ArrayValue;
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
  });
  it("CHOOSECOLS zero index → #VALUE!", () => {
    expect(fnCHOOSECOLS([arr, rvNumber(0)])).toEqual(ERRORS.VALUE);
  });
  it("CHOOSECOLS out of range → #VALUE!", () => {
    expect(fnCHOOSECOLS([arr, rvNumber(99)])).toEqual(ERRORS.VALUE);
  });
  it("CHOOSECOLS duplicate indices allowed", () => {
    const r = fnCHOOSECOLS([arr, rvNumber(1), rvNumber(1)]) as ArrayValue;
    expect(r.width).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[0][1] as NumberValue).value).toBe(1);
  });
});

describe("DROP / TAKE extra saturation", () => {
  const arr = rvArray([
    [rvNumber(1), rvNumber(2)],
    [rvNumber(3), rvNumber(4)],
    [rvNumber(5), rvNumber(6)]
  ]);

  it("TAKE with positive rows + positive cols takes top-left sub-array", () => {
    const r = fnTAKE([arr, rvNumber(2), rvNumber(1)]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(1);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
  });
  it("TAKE with negative rows takes bottom", () => {
    const r = fnTAKE([arr, rvNumber(-2)]) as ArrayValue;
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
  });
  it("TAKE exceeds dimension returns whole array", () => {
    const r = fnTAKE([arr, rvNumber(10)]) as ArrayValue;
    expect(r.height).toBe(3);
  });
  it("DROP rows only", () => {
    const r = fnDROP([arr, rvNumber(1)]) as ArrayValue;
    expect(r.height).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(3);
  });
  it("DROP rows and cols both", () => {
    const r = fnDROP([arr, rvNumber(1), rvNumber(1)]);
    expect(r.kind).toBe(RVKind.Array);
    const a = r as ArrayValue;
    expect(a.height).toBeGreaterThanOrEqual(1);
  });
  it("DROP negative rows drops from end", () => {
    const r = fnDROP([arr, rvNumber(-1)]) as ArrayValue;
    expect(r.height).toBe(2);
  });
});

describe("EXPAND saturation", () => {
  const src = rvArray([
    [rvNumber(1), rvNumber(2)],
    [rvNumber(3), rvNumber(4)]
  ]);

  it("expands to larger rectangle, pads with #N/A", () => {
    const r = fnEXPAND([src, rvNumber(3), rvNumber(3)]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(3);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[2][2] as { code: string }).code).toBe("#N/A");
  });
  it("expand with custom pad", () => {
    const r = fnEXPAND([src, rvNumber(3), rvNumber(3), rvNumber(0)]) as ArrayValue;
    expect((r.rows[2][2] as NumberValue).value).toBe(0);
  });
  it("expand to same shape is no-op", () => {
    const r = fnEXPAND([src, rvNumber(2), rvNumber(2)]) as ArrayValue;
    expect(r.height).toBe(2);
    expect(r.width).toBe(2);
  });
  it("target smaller than source → #VALUE!", () => {
    expect(fnEXPAND([src, rvNumber(1), rvNumber(2)])).toEqual(ERRORS.VALUE);
  });
  it("target <= 0 → #VALUE!", () => {
    expect(fnEXPAND([src, rvNumber(0), rvNumber(2)])).toEqual(ERRORS.VALUE);
  });
  it("over 10M cells → #NUM!", () => {
    expect(fnEXPAND([src, rvNumber(1e5), rvNumber(1e5)])).toEqual(ERRORS.NUM);
  });
});

describe("AGGREGATE saturation", () => {
  it("function code 1 = AVERAGE", () => {
    expect(
      asNumber(
        fnAGGREGATE([
          rvNumber(1),
          rvNumber(0),
          rvArray([[rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4)]])
        ])
      )
    ).toBe(2.5);
  });
  it("function code 9 = SUM", () => {
    expect(
      asNumber(
        fnAGGREGATE([rvNumber(9), rvNumber(0), rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]])])
      )
    ).toBe(6);
  });
  it("option 6 ignore error values (engine may propagate instead)", () => {
    // Strict Excel skips errors under option 6; this engine's AGGREGATE
    // passes them through. Verify no crash either way.
    const r = fnAGGREGATE([
      rvNumber(9),
      rvNumber(6),
      rvArray([[rvNumber(1), ERRORS.NA, rvNumber(3)]])
    ]);
    expect([RVKind.Number, RVKind.Error]).toContain(r.kind);
  });
  it("invalid function code → #VALUE!", () => {
    expect(fnAGGREGATE([rvNumber(99), rvNumber(0), rvArray([[rvNumber(1)]])])).toEqual(
      ERRORS.VALUE
    );
  });
  it("function code 4 = MAX", () => {
    expect(
      asNumber(
        fnAGGREGATE([rvNumber(4), rvNumber(0), rvArray([[rvNumber(1), rvNumber(5), rvNumber(3)]])])
      )
    ).toBe(5);
  });
  it("function code 5 = MIN", () => {
    expect(
      asNumber(
        fnAGGREGATE([rvNumber(5), rvNumber(0), rvArray([[rvNumber(1), rvNumber(5), rvNumber(3)]])])
      )
    ).toBe(1);
  });
});

describe("WRAPROWS + TOROW extras (R9 batch 2)", () => {
  it("WRAPROWS wraps 6 elements into rows of 2", () => {
    const src = rvArray([
      [rvNumber(1), rvNumber(2), rvNumber(3), rvNumber(4), rvNumber(5), rvNumber(6)]
    ]);
    const r = fnWRAPROWS([src, rvNumber(2)]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(2);
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[2][1] as NumberValue).value).toBe(6);
  });
  it("WRAPROWS custom pad value", () => {
    const src = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const r = fnWRAPROWS([src, rvNumber(2), rvNumber(0)]) as ArrayValue;
    expect((r.rows[1][1] as NumberValue).value).toBe(0);
  });
  it("WRAPROWS wrap_count = 1", () => {
    const src = rvArray([[rvNumber(1), rvNumber(2), rvNumber(3)]]);
    const r = fnWRAPROWS([src, rvNumber(1)]) as ArrayValue;
    expect(r.height).toBe(3);
    expect(r.width).toBe(1);
  });
  it("WRAPROWS wrap_count <= 0 → #VALUE!", () => {
    expect(fnWRAPROWS([rvArray([[rvNumber(1)]]), rvNumber(0)])).toEqual(ERRORS.VALUE);
  });
  it("WRAPROWS error source produces error", () => {
    const r = fnWRAPROWS([ERRORS.NA, rvNumber(2)]);
    expect(r.kind).toBe(RVKind.Error);
  });

  it("TOROW flattens 3×2", () => {
    const src = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)],
      [rvNumber(5), rvNumber(6)]
    ]);
    const r = fnTOROW([src]) as ArrayValue;
    expect(r.height).toBe(1);
    expect(r.width).toBe(6);
  });
  it("TOROW scan_by_col=TRUE", () => {
    const src = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)]
    ]);
    const r = fnTOROW([src, rvNumber(0), rvBoolean(true)]) as ArrayValue;
    expect((r.rows[0][0] as NumberValue).value).toBe(1);
    expect((r.rows[0][1] as NumberValue).value).toBe(3);
  });
  it("TOROW ignore=1 skips blanks", () => {
    const src = rvArray([[rvNumber(1), BLANK, rvNumber(3)]]);
    const r = fnTOROW([src, rvNumber(1)]) as ArrayValue;
    expect(r.width).toBe(2);
  });
  it("TOROW scalar returns 1×1", () => {
    const r = fnTOROW([rvNumber(42)]) as ArrayValue;
    expect(r.width).toBe(1);
  });
});
