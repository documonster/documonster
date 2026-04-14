import util from "util";

import { isEqual } from "@excel/utils/under-dash";
import { describe, it, expect } from "vitest";

describe("under-dash", () => {
  describe("isEqual", () => {
    function showVal(o: any) {
      return util.inspect(o, { compact: true });
    }

    it("compares primitives and references", () => {
      const fn1 = () => {};
      const fn2 = () => {};
      const sym1 = Symbol("foo");
      const sym2 = Symbol("foo");

      const cases: Array<{ a: any; b: any; expected: boolean; label: string }> = [
        { a: 0, b: 0, expected: true, label: "0 === 0" },
        { a: 0, b: 1, expected: false, label: "0 !== 1" },
        { a: true, b: true, expected: true, label: "true === true" },
        { a: true, b: false, expected: false, label: "true !== false" },
        { a: "string", b: "string", expected: true, label: "same string" },
        { a: "string", b: "other string", expected: false, label: "different strings" },
        { a: null, b: null, expected: true, label: "null === null" },
        { a: undefined, b: undefined, expected: true, label: "undefined === undefined" },
        { a: null, b: undefined, expected: false, label: "null !== undefined" },
        { a: fn1, b: fn1, expected: true, label: "same function reference" },
        { a: fn1, b: fn2, expected: false, label: "different functions" },
        { a: sym1, b: sym1, expected: true, label: "same symbol reference" },
        { a: sym1, b: sym2, expected: false, label: "different symbols" },
        { a: 1, b: "1", expected: false, label: "different types" }
      ];

      for (const c of cases) {
        const assertion = `${c.label}: ${showVal(c.a)} ${c.expected ? "==" : "!="} ${showVal(c.b)}`;
        expect(isEqual(c.a, c.b), assertion).toBe(c.expected);
        expect(isEqual(c.b, c.a), assertion + " (symmetric)").toBe(c.expected);
      }
    });

    it("compares arrays deeply", () => {
      const cases: Array<{ a: any; b: any; expected: boolean; label: string }> = [
        { a: [], b: [], expected: true, label: "empty arrays" },
        { a: ["array"], b: ["array"], expected: true, label: "same single element" },
        { a: ["array"], b: ["array2"], expected: false, label: "different element" },
        {
          a: ["array", "foobar"],
          b: ["array", "foobar"],
          expected: true,
          label: "same two elements"
        },
        {
          a: ["array", "foobar"],
          b: ["array", "quux"],
          expected: false,
          label: "different second element"
        },
        { a: [{ object: 1 }], b: [{ object: 1 }], expected: true, label: "nested object" },
        { a: [{ object: 1 }], b: [{ object: 2 }], expected: false, label: "nested object differs" },
        { a: [null, undefined], b: [null, undefined], expected: true, label: "null + undefined" },
        { a: [null], b: [undefined], expected: false, label: "null vs undefined" }
      ];

      for (const c of cases) {
        const assertion = `${c.label}: ${showVal(c.a)} ${c.expected ? "==" : "!="} ${showVal(c.b)}`;
        expect(isEqual(c.a, c.b), assertion).toBe(c.expected);
        expect(isEqual(c.b, c.a), assertion + " (symmetric)").toBe(c.expected);
      }
    });

    it("compares objects deeply", () => {
      const cases: Array<{ a: any; b: any; expected: boolean; label: string }> = [
        { a: {}, b: {}, expected: true, label: "empty objects" },
        { a: { object: 1 }, b: { object: 1 }, expected: true, label: "same key/value" },
        { a: { object: 1 }, b: { object: 2 }, expected: false, label: "different values" },
        {
          a: { object: 1, foobar: "quux" },
          b: { foobar: "quux", object: 1 },
          expected: true,
          label: "same keys different order"
        },
        {
          a: { nested: { object: 1, foobar: "quux" } },
          b: { nested: { object: 1, foobar: "quux" } },
          expected: true,
          label: "nested object"
        },
        {
          a: { nested: { object: 1, foobar: "quux" } },
          b: { nested: { object: 2, foobar: "quux" } },
          expected: false,
          label: "nested object differs"
        },
        {
          a: { key: undefined },
          b: {},
          expected: false,
          label: "missing key vs explicit undefined"
        }
      ];

      for (const c of cases) {
        const assertion = `${c.label}: ${showVal(c.a)} ${c.expected ? "==" : "!="} ${showVal(c.b)}`;
        expect(isEqual(c.a, c.b), assertion).toBe(c.expected);
        expect(isEqual(c.b, c.a), assertion + " (symmetric)").toBe(c.expected);
      }
    });
  });
});
