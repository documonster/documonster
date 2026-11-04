import util from "util";

import {
  deepEqual,
  isForbiddenKey,
  isRecordArgument,
  isSafeDynamicKey,
  MAX_DYNAMIC_KEY_LENGTH
} from "@utils/object";
import { describe, it, expect } from "vitest";

describe("@utils/object", () => {
  describe("deepEqual", () => {
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
        expect(deepEqual(c.a, c.b), assertion).toBe(c.expected);
        expect(deepEqual(c.b, c.a), assertion + " (symmetric)").toBe(c.expected);
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
        expect(deepEqual(c.a, c.b), assertion).toBe(c.expected);
        expect(deepEqual(c.b, c.a), assertion + " (symmetric)").toBe(c.expected);
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
        expect(deepEqual(c.a, c.b), assertion).toBe(c.expected);
        expect(deepEqual(c.b, c.a), assertion + " (symmetric)").toBe(c.expected);
      }
    });
  });

  describe("isForbiddenKey", () => {
    it("flags prototype-pollution keys", () => {
      expect(isForbiddenKey("__proto__")).toBe(true);
      expect(isForbiddenKey("constructor")).toBe(true);
      expect(isForbiddenKey("prototype")).toBe(true);
    });

    it("allows ordinary keys", () => {
      expect(isForbiddenKey("total")).toBe(false);
      expect(isForbiddenKey("")).toBe(false);
    });
  });

  describe("isSafeDynamicKey", () => {
    it("accepts legitimate column names / aliases", () => {
      for (const key of ["count", "total", "First Name", "价格", "sum_x", "a.b-c", ""]) {
        expect(isSafeDynamicKey(key), key).toBe(true);
      }
    });

    it("rejects prototype-pollution keys", () => {
      expect(isSafeDynamicKey("__proto__")).toBe(false);
      expect(isSafeDynamicKey("constructor")).toBe(false);
      expect(isSafeDynamicKey("prototype")).toBe(false);
    });

    it("rejects control characters", () => {
      expect(isSafeDynamicKey("a\u0000b")).toBe(false);
      expect(isSafeDynamicKey("a\tb")).toBe(false);
      expect(isSafeDynamicKey("a\nb")).toBe(false);
      expect(isSafeDynamicKey("a\u007fb")).toBe(false);
    });

    it("rejects over-length keys but accepts at-limit keys", () => {
      expect(isSafeDynamicKey("x".repeat(MAX_DYNAMIC_KEY_LENGTH))).toBe(true);
      expect(isSafeDynamicKey("x".repeat(MAX_DYNAMIC_KEY_LENGTH + 1))).toBe(false);
    });
  });

  describe("isRecordArgument", () => {
    it("accepts option objects, with or without a prototype", () => {
      expect(isRecordArgument({})).toBe(true);
      expect(isRecordArgument({ columns: 1 })).toBe(true);
      expect(isRecordArgument(Object.create(null))).toBe(true);
      class Options {
        columns = 1;
      }
      expect(isRecordArgument(new Options())).toBe(true);
    });

    it("rejects every positional placeholder", () => {
      for (const value of [null, undefined, 0, 1, "", "x", true, false, 1n, Symbol("s")]) {
        expect(isRecordArgument(value), String(value)).toBe(false);
      }
    });

    it("rejects sequences, binary data and dates", () => {
      for (const value of [
        [],
        [0],
        new Set([0]),
        new Map(),
        new Int32Array(1),
        new Uint8Array(1),
        new DataView(new ArrayBuffer(1)),
        new ArrayBuffer(1),
        new Date(0),
        // A boxed string is iterable, so it is a sequence too.
        Object("x")
      ]) {
        expect(isRecordArgument(value), Object.prototype.toString.call(value)).toBe(false);
      }
    });
  });
});
