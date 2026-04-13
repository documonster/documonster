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

    it("should use asymmetric padding when provided", () => {
      const rect = { x: 10, width: 100 };
      const padLeft = 5;
      const padRight = 10;

      // left-aligned: x + padLeft + indent
      expect(computeTextX("left", rect, 20, 0, padLeft, padRight)).toBe(15);

      // right-aligned: x + width - padRight - textWidth
      expect(computeTextX("right", rect, 20, 0, padLeft, padRight)).toBe(80);

      // center: unchanged by padding
      expect(computeTextX("center", rect, 20, 0, padLeft, padRight)).toBe(50);
    });

    it("should clamp right-aligned text to padLeft boundary", () => {
      // Right-aligned with very wide text that would start before cell left
      const rect = { x: 10, width: 50 };
      // textWidth=100 > rect.width, so right-align would put x at 10+50-5-100 = -45
      // Clamp to minX = x + padLeft = 15
      expect(computeTextX("right", rect, 100, 0, 5, 5)).toBe(15);
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

    it("should use asymmetric vertical padding", () => {
      const padTop = 5;
      const padBottom = 10;

      // top-aligned: y + height - padTop - ascent
      const topY = computeTextStartY("top", rect, 12, 8, padTop, padBottom);
      expect(topY).toBe(10 + 40 - 5 - 8); // 37

      // bottom-aligned: y + padBottom + (totalTextHeight - ascent)
      const bottomY = computeTextStartY("bottom", rect, 12, 8, padTop, padBottom);
      expect(bottomY).toBe(10 + 10 + (12 - 8)); // 24
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
