import { describe, expect, it } from "vitest";

import { describeValue } from "../describe-value";

describe("describeValue", () => {
  it("names what it was given in words that tell it from what was wanted", () => {
    expect(describeValue(new Date(Number.NaN))).toBe("an invalid Date");
    expect(describeValue(new Date(0))).toBe("a Date");
    expect(describeValue(new Map())).toBe("a Map");
    expect(describeValue(new Uint8Array(1))).toBe("a Uint8Array");
    expect(describeValue(new ArrayBuffer(1))).toBe("an ArrayBuffer");
    expect(describeValue(new Int8Array(1))).toBe("an Int8Array");
    expect(describeValue({})).toBe("an object");
    expect(describeValue([])).toBe("an array");
    expect(describeValue(() => 1)).toBe("a function");
  });

  it("shows a primitive as its value, a long string cut short", () => {
    expect(describeValue(42)).toBe("42");
    expect(describeValue(Number.NaN)).toBe("NaN");
    expect(describeValue(true)).toBe("true");
    expect(describeValue(null)).toBe("null");
    expect(describeValue(undefined)).toBe("undefined");
    expect(describeValue(10n)).toBe("10n");
    expect(describeValue(Symbol("s"))).toBe("Symbol(s)");
    expect(describeValue("x")).toBe('"x"');
    expect(describeValue("a".repeat(100))).toBe(`"${"a".repeat(60)}…"`);
  });
});
