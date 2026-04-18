/**
 * Unit tests for the RuntimeValue system.
 *
 * Covers constructor helpers, singletons, the rectangular-padding behaviour
 * of `rvArray`, coercion (`toNumberRV`, `toStringRV`, `toBooleanRV`),
 * `scalarEquals`, `topLeft`, the error table, type guards, and snapshot
 * conversion (`fromSnapshotValue`).
 */

import { describe, it, expect } from "vitest";

import {
  BLANK,
  ERRORS,
  RVKind,
  fromSnapshotValue,
  isArray,
  isError,
  isLambda,
  isScalar,
  rvArray,
  rvBoolean,
  rvCellRef,
  rvError,
  rvNumber,
  rvRef,
  rvString,
  scalarEquals,
  toBooleanRV,
  toNumberRV,
  toStringRV,
  topLeft,
  type ArrayValue,
  type BooleanValue,
  type ErrorValue,
  type NumberValue,
  type ReferenceValue,
  type RuntimeValue,
  type ScalarValue,
  type StringValue
} from "../values";

describe("values — constructor helpers", () => {
  it("rvNumber produces a NumberValue", () => {
    const v = rvNumber(3.14);
    expect(v.kind).toBe(RVKind.Number);
    expect(v.value).toBe(3.14);
  });

  it("rvString produces a StringValue", () => {
    const v = rvString("hi");
    expect(v.kind).toBe(RVKind.String);
    expect(v.value).toBe("hi");
  });

  it("rvBoolean produces a BooleanValue", () => {
    expect(rvBoolean(true).value).toBe(true);
    expect(rvBoolean(false).value).toBe(false);
  });

  it("rvError produces an ErrorValue", () => {
    const v = rvError("#DIV/0!");
    expect(v.kind).toBe(RVKind.Error);
    expect(v.code).toBe("#DIV/0!");
  });

  it("rvRef wraps a rectangular area", () => {
    const r: ReferenceValue = rvRef("Sheet1", 2, 3, 5, 7);
    expect(r.kind).toBe(RVKind.Reference);
    expect(r.areas).toHaveLength(1);
    expect(r.areas[0]).toEqual({ sheet: "Sheet1", top: 2, left: 3, bottom: 5, right: 7 });
    expect(r.singleCell).toBeUndefined();
  });

  it("rvCellRef marks single-cell refs with singleCell=true", () => {
    const r = rvCellRef("Sheet1", 4, 8);
    expect(r.singleCell).toBe(true);
    expect(r.areas[0]).toEqual({ sheet: "Sheet1", top: 4, left: 8, bottom: 4, right: 8 });
  });
});

describe("values — singleton constants", () => {
  it("BLANK is the single canonical Blank value", () => {
    expect(BLANK.kind).toBe(RVKind.Blank);
  });

  it("ERRORS table contains the expected codes", () => {
    expect(ERRORS.VALUE.code).toBe("#VALUE!");
    expect(ERRORS.REF.code).toBe("#REF!");
    expect(ERRORS.NAME.code).toBe("#NAME?");
    expect(ERRORS.DIV0.code).toBe("#DIV/0!");
    expect(ERRORS.NA.code).toBe("#N/A");
    expect(ERRORS.NUM.code).toBe("#NUM!");
    expect(ERRORS.NULL.code).toBe("#NULL!");
    expect(ERRORS.SPILL.code).toBe("#SPILL!");
    expect(ERRORS.CALC.code).toBe("#CALC!");
  });
});

describe("values — rvArray rectangularisation", () => {
  it("records height and width from a uniformly-shaped input", () => {
    const arr = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)]
    ]);
    expect(arr.height).toBe(2);
    expect(arr.width).toBe(2);
    expect(arr.rows[0][0]).toEqual(rvNumber(1));
    expect(arr.rows[1][1]).toEqual(rvNumber(4));
  });

  it("pads short rows with BLANK to reach the maximum width", () => {
    const arr: ArrayValue = rvArray([[rvNumber(1)], [rvNumber(2), rvNumber(3), rvNumber(4)]]);
    expect(arr.width).toBe(3);
    expect(arr.height).toBe(2);
    expect(arr.rows[0]).toHaveLength(3);
    // Padding is the canonical BLANK singleton.
    expect(arr.rows[0][1]).toEqual(BLANK);
    expect(arr.rows[0][2]).toEqual(BLANK);
  });

  it("attaches origin metadata when both origins are supplied", () => {
    const arr = rvArray([[rvNumber(1)]], 5, 7);
    expect(arr.originRow).toBe(5);
    expect(arr.originCol).toBe(7);
  });
});

describe("values — toNumberRV coercion", () => {
  it("returns the same NumberValue when input is a number", () => {
    const n = rvNumber(5);
    const out = toNumberRV(n);
    expect(out.kind).toBe(RVKind.Number);
    expect((out as NumberValue).value).toBe(5);
  });

  it("coerces Blank to 0", () => {
    const out = toNumberRV(BLANK);
    expect(out.kind).toBe(RVKind.Number);
    expect((out as NumberValue).value).toBe(0);
  });

  it("coerces Boolean to 1 / 0", () => {
    expect((toNumberRV(rvBoolean(true)) as NumberValue).value).toBe(1);
    expect((toNumberRV(rvBoolean(false)) as NumberValue).value).toBe(0);
  });

  it("coerces a numeric string", () => {
    expect((toNumberRV(rvString("3.5")) as NumberValue).value).toBe(3.5);
  });

  it("returns #VALUE! for an empty string", () => {
    expect(toNumberRV(rvString(""))).toBe(ERRORS.VALUE);
  });

  it("returns #VALUE! for a non-numeric string", () => {
    expect(toNumberRV(rvString("hello"))).toBe(ERRORS.VALUE);
  });

  it("propagates errors unchanged", () => {
    const err = rvError("#N/A");
    expect(toNumberRV(err)).toEqual(err);
  });
});

describe("values — toStringRV coercion", () => {
  it("leaves strings unchanged", () => {
    expect(toStringRV(rvString("hi"))).toBe("hi");
  });

  it("formats a number via String(n)", () => {
    expect(toStringRV(rvNumber(42))).toBe("42");
    expect(toStringRV(rvNumber(1.5))).toBe("1.5");
  });

  it("formats booleans as uppercase TRUE/FALSE", () => {
    expect(toStringRV(rvBoolean(true))).toBe("TRUE");
    expect(toStringRV(rvBoolean(false))).toBe("FALSE");
  });

  it("returns empty string for Blank", () => {
    expect(toStringRV(BLANK)).toBe("");
  });

  it("returns the error code for an error value", () => {
    expect(toStringRV(rvError("#REF!"))).toBe("#REF!");
  });
});

describe("values — toBooleanRV coercion", () => {
  it("leaves booleans unchanged", () => {
    expect((toBooleanRV(rvBoolean(true)) as BooleanValue).value).toBe(true);
  });

  it("converts 0 to false and non-zero to true", () => {
    expect((toBooleanRV(rvNumber(0)) as BooleanValue).value).toBe(false);
    expect((toBooleanRV(rvNumber(-1)) as BooleanValue).value).toBe(true);
  });

  it("converts Blank to false", () => {
    expect((toBooleanRV(BLANK) as BooleanValue).value).toBe(false);
  });

  it("parses the strings 'TRUE' and 'FALSE' case-insensitively", () => {
    expect((toBooleanRV(rvString("true")) as BooleanValue).value).toBe(true);
    expect((toBooleanRV(rvString("False")) as BooleanValue).value).toBe(false);
  });

  it("returns #VALUE! for an unparseable string", () => {
    expect(toBooleanRV(rvString("maybe"))).toBe(ERRORS.VALUE);
  });

  it("propagates errors", () => {
    const e = rvError("#NUM!");
    expect(toBooleanRV(e)).toEqual(e);
  });
});

describe("values — scalarEquals", () => {
  it("returns true for two equal numbers", () => {
    expect(scalarEquals(rvNumber(3), rvNumber(3))).toBe(true);
  });

  it("returns false for different kinds", () => {
    expect(scalarEquals(rvNumber(1), rvString("1"))).toBe(false);
  });

  it("is case-insensitive for strings", () => {
    expect(scalarEquals(rvString("Foo"), rvString("FOO"))).toBe(true);
  });

  it("treats Blank as equal to Blank", () => {
    expect(scalarEquals(BLANK, BLANK)).toBe(true);
  });

  it("returns false when comparing errors (errors are not equal to each other)", () => {
    expect(scalarEquals(rvError("#N/A"), rvError("#N/A"))).toBe(false);
  });
});

describe("values — topLeft", () => {
  it("returns a scalar unchanged", () => {
    const v = rvNumber(7);
    expect(topLeft(v)).toBe(v);
  });

  it("returns the [0][0] element of an array", () => {
    const arr = rvArray([
      [rvNumber(1), rvNumber(2)],
      [rvNumber(3), rvNumber(4)]
    ]);
    const tl = topLeft(arr);
    expect(tl).toEqual(rvNumber(1));
  });

  it("returns BLANK for an empty array", () => {
    // rvArray of shape 0x0
    const arr = rvArray([]);
    expect(topLeft(arr)).toEqual(BLANK);
  });

  it("returns #VALUE! for a Reference (cannot resolve without context)", () => {
    const ref = rvCellRef("Sheet1", 1, 1);
    expect(topLeft(ref)).toEqual(ERRORS.VALUE);
  });
});

describe("values — type guards", () => {
  it("isError identifies errors only", () => {
    expect(isError(rvError("#N/A"))).toBe(true);
    expect(isError(rvNumber(1))).toBe(false);
    expect(isError(BLANK)).toBe(false);
  });

  it("isArray identifies arrays only", () => {
    expect(isArray(rvArray([[rvNumber(1)]]))).toBe(true);
    expect(isArray(rvNumber(1))).toBe(false);
  });

  it("isLambda identifies lambdas only", () => {
    // We don't need a real BoundExpr body for the type-guard check to run.
    const fake: RuntimeValue = {
      kind: RVKind.Lambda,
      params: ["X"],
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      body: { kind: 1 as any, value: 0 },
      closureBindings: undefined
    } as RuntimeValue;
    expect(isLambda(fake)).toBe(true);
    expect(isLambda(rvNumber(1))).toBe(false);
  });

  it("isScalar accepts Blank/Number/String/Boolean/Error but rejects Array/Reference", () => {
    const scalars: ScalarValue[] = [
      BLANK,
      rvNumber(1),
      rvString("x"),
      rvBoolean(true),
      rvError("#N/A")
    ];
    for (const s of scalars) {
      expect(isScalar(s)).toBe(true);
    }
    expect(isScalar(rvArray([[rvNumber(1)]]))).toBe(false);
    expect(isScalar(rvCellRef("S", 1, 1))).toBe(false);
  });
});

describe("values — fromSnapshotValue", () => {
  it("maps null to BLANK", () => {
    expect(fromSnapshotValue(null)).toEqual(BLANK);
  });

  it("maps a number", () => {
    expect(fromSnapshotValue(5)).toEqual(rvNumber(5));
  });

  it("maps a string", () => {
    expect(fromSnapshotValue("hi")).toEqual(rvString("hi"));
  });

  it("maps a boolean", () => {
    expect(fromSnapshotValue(true)).toEqual(rvBoolean(true));
  });

  it("maps an { error } object to an ErrorValue", () => {
    const out = fromSnapshotValue({ error: "#REF!" }) as ErrorValue;
    expect(out.kind).toBe(RVKind.Error);
    expect(out.code).toBe("#REF!");
  });
});

describe("values — preserved shapes", () => {
  it("NumberValue objects have exactly kind and value properties (no monkey-patching)", () => {
    const n: NumberValue = rvNumber(1);
    expect(Object.keys(n).sort()).toEqual(["kind", "value"]);
  });

  it("StringValue objects have exactly kind and value properties", () => {
    const s: StringValue = rvString("x");
    expect(Object.keys(s).sort()).toEqual(["kind", "value"]);
  });
});

describe("rvArray — subtotalMask", () => {
  it("omits the subtotalMask field entirely when no mask is provided", () => {
    const a = rvArray([[rvNumber(1), rvNumber(2)]]);
    expect(a.subtotalMask).toBeUndefined();
    expect("subtotalMask" in a).toBe(false);
  });

  it("attaches a subtotalMask when one is supplied", () => {
    const mask = [[true, false]];
    const a = rvArray([[rvNumber(1), rvNumber(2)]], 1, 1, mask);
    expect(a.subtotalMask).toBe(mask);
    expect(a.subtotalMask?.[0][0]).toBe(true);
    expect(a.subtotalMask?.[0][1]).toBe(false);
  });

  it("array origin + mask coexist", () => {
    const mask = [[false], [true]];
    const a = rvArray([[rvNumber(10)], [rvNumber(20)]], 5, 3, mask);
    expect(a.originRow).toBe(5);
    expect(a.originCol).toBe(3);
    expect(a.subtotalMask).toBe(mask);
    expect(a.height).toBe(2);
    expect(a.width).toBe(1);
  });
});

describe("rvArray — hiddenRowMask", () => {
  it("omits hiddenRowMask when none is provided", () => {
    const a = rvArray([[rvNumber(1)], [rvNumber(2)]]);
    expect(a.hiddenRowMask).toBeUndefined();
    expect("hiddenRowMask" in a).toBe(false);
  });

  it("attaches a hiddenRowMask when supplied (5th arg)", () => {
    const mask = [true, false];
    const a = rvArray([[rvNumber(1)], [rvNumber(2)]], 1, 1, undefined, mask);
    expect(a.hiddenRowMask).toBe(mask);
    expect(a.hiddenRowMask?.[0]).toBe(true);
    expect(a.hiddenRowMask?.[1]).toBe(false);
  });

  it("both masks coexist with origin", () => {
    const subMask = [[true], [false]];
    const hiddenMask = [false, true];
    const a = rvArray([[rvNumber(10)], [rvNumber(20)]], 3, 4, subMask, hiddenMask);
    expect(a.originRow).toBe(3);
    expect(a.originCol).toBe(4);
    expect(a.subtotalMask).toBe(subMask);
    expect(a.hiddenRowMask).toBe(hiddenMask);
  });
});
