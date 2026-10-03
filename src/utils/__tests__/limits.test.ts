import { describe, expect, it } from "vitest";

import { assertLimitOption } from "../limits";

describe("assertLimitOption", () => {
  it("accepts undefined, zero, positive numbers and Infinity", () => {
    for (const value of [undefined, 0, 1, 1.5, Infinity]) {
      expect(() => assertLimitOption("maxX", value)).not.toThrow();
    }
  });

  it("rejects NaN, negatives and non-numbers with a RangeError naming the option", () => {
    for (const value of [Number.NaN, -1, -Infinity, "10", null]) {
      expect(() => assertLimitOption("maxX", value)).toThrow(/maxX must be a non-negative/);
      expect(() => assertLimitOption("maxX", value)).toThrow(RangeError);
    }
  });
});
