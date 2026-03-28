/**
 * Focused tests for page renderer helpers.
 */
import { describe, expect, it } from "vitest";
import {
  computeTextStartY,
  computeTextX,
  wrapTextLines,
  alphaGsName
} from "@pdf/render/page-renderer";

describe("page-renderer helpers", () => {
  describe("computeTextX", () => {
    it("should apply indent for left-aligned text", () => {
      expect(computeTextX("left", { x: 10, width: 100 }, 20, 8)).toBe(21);
    });

    it("should ignore indent for centered text", () => {
      expect(computeTextX("center", { x: 10, width: 100 }, 20, 8)).toBe(50);
    });
  });

  describe("computeTextStartY", () => {
    const rect = { x: 0, y: 10, width: 100, height: 40 };

    it("should order top above middle above bottom", () => {
      const top = computeTextStartY("top", rect, 12, 8);
      const middle = computeTextStartY("middle", rect, 12, 8);
      const bottom = computeTextStartY("bottom", rect, 12, 8);

      expect(top).toBeGreaterThan(middle);
      expect(middle).toBeGreaterThan(bottom);
    });
  });

  describe("wrapTextLines", () => {
    const measure = (s: string) => s.length;

    it("should wrap greedily by words", () => {
      expect(wrapTextLines("aa bb ccc", measure, 5)).toEqual(["aa bb", "ccc"]);
    });

    it("should preserve explicit blank lines", () => {
      expect(wrapTextLines("a\n\nb", measure, 10)).toEqual(["a", "", "b"]);
    });

    it("should keep a single oversized word on one line", () => {
      expect(wrapTextLines("superlongword", measure, 3)).toEqual(["superlongword"]);
    });
  });

  describe("alphaGsName", () => {
    it("should avoid collisions for close alpha values", () => {
      expect(alphaGsName(0.5001)).not.toBe(alphaGsName(0.5002));
    });
  });
});
