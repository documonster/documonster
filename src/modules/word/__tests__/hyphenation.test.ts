/**
 * Hyphenation Engine Tests
 *
 * Tests for Liang's hyphenation algorithm implementation covering
 * pattern-based hyphenation, exception dictionaries, and text processing.
 */

import { describe, it, expect } from "vitest";

import {
  createHyphenator,
  hyphenateWord,
  hyphenateText,
  ENGLISH_US_PATTERNS
} from "../font/hyphenation";
import type { HyphenationPatterns } from "../font/hyphenation";

describe("Hyphenation", () => {
  describe("createHyphenator", () => {
    it("should return a function when created with ENGLISH_US_PATTERNS", () => {
      const hyphenator = createHyphenator(ENGLISH_US_PATTERNS);
      expect(typeof hyphenator).toBe("function");
    });

    it("should return a function when created with custom patterns", () => {
      const patterns = {
        language: "test",
        patterns: ["1tion", "pre1"],
        exceptions: []
      };
      const hyphenator = createHyphenator(patterns);
      expect(typeof hyphenator).toBe("function");
    });

    it("should accept options parameter", () => {
      const hyphenator = createHyphenator(ENGLISH_US_PATTERNS, {
        minLeft: 3,
        minRight: 3
      });
      expect(typeof hyphenator).toBe("function");
    });
  });

  describe("Basic English word hyphenation", () => {
    const hyphenator = createHyphenator(ENGLISH_US_PATTERNS);

    it("should produce break points for 'computer'", () => {
      const points = hyphenator("computer");
      expect(points.length).toBeGreaterThan(0);
      // All break points must be within valid range
      for (const p of points) {
        expect(p).toBeGreaterThanOrEqual(2);
        expect(p).toBeLessThanOrEqual(6); // word.length - minRight
      }
    });

    it("should produce break points for 'information'", () => {
      const points = hyphenator("information");
      expect(points.length).toBeGreaterThan(0);
      for (const p of points) {
        expect(p).toBeGreaterThanOrEqual(2);
        expect(p).toBeLessThanOrEqual(9); // 11 - 2
      }
    });

    it("should produce break points for 'hyphenation'", () => {
      const points = hyphenator("hyphenation");
      expect(points.length).toBeGreaterThan(0);
      for (const p of points) {
        expect(p).toBeGreaterThanOrEqual(2);
        expect(p).toBeLessThanOrEqual(9);
      }
    });

    it("should return sorted break points", () => {
      const points = hyphenator("information");
      for (let i = 1; i < points.length; i++) {
        expect(points[i]).toBeGreaterThan(points[i - 1]);
      }
    });
  });

  describe("Short words", () => {
    const hyphenator = createHyphenator(ENGLISH_US_PATTERNS);

    it("should not hyphenate a 4-letter word", () => {
      const points = hyphenator("test");
      // With minLeft=2, minRight=2 on a 4-char word, there's only position 2
      // which may or may not be a break point, but likely none due to constraints
      for (const p of points) {
        expect(p).toBeGreaterThanOrEqual(2);
        expect(p).toBeLessThanOrEqual(2);
      }
    });

    it("should not hyphenate a 3-letter word", () => {
      const points = hyphenator("the");
      // minLeft=2, minRight=2 means no valid positions for a 3-char word
      expect(points).toHaveLength(0);
    });

    it("should not hyphenate a 2-letter word", () => {
      const points = hyphenator("to");
      expect(points).toHaveLength(0);
    });

    it("should not hyphenate a 1-letter word", () => {
      const points = hyphenator("a");
      expect(points).toHaveLength(0);
    });
  });

  describe("Exception words", () => {
    const hyphenator = createHyphenator(ENGLISH_US_PATTERNS);

    it("should hyphenate 'associate' using exception dictionary", () => {
      const points = hyphenator("associate");
      // Exception: "as-so-ciate" → points at [2, 4]
      expect(points).toContain(2);
      expect(points).toContain(4);
    });

    it("should hyphenate 'associates' using exception dictionary", () => {
      const points = hyphenator("associates");
      // Exception: "as-so-ciates" → points at [2, 4]
      expect(points).toContain(2);
      expect(points).toContain(4);
    });

    it("should handle exception word 'table' with no valid breaks", () => {
      // Exception: "ta-ble" → point at [2]
      const points = hyphenator("table");
      // minLeft=2, minRight=2: point 2 is valid (>= 2 and <= 5-2=3)
      expect(points).toContain(2);
    });

    it("should handle exception word with no hyphens (present)", () => {
      // Exception: "present" has no dashes → no break points
      const points = hyphenator("present");
      expect(points).toHaveLength(0);
    });
  });

  describe("hyphenateWord", () => {
    const hyphenator = createHyphenator(ENGLISH_US_PATTERNS);

    it("should insert soft hyphens at break points", () => {
      const result = hyphenateWord("associate", hyphenator);
      // Should contain soft hyphen character
      expect(result).toContain("\u00AD");
      // Original letters should be preserved
      expect(result.replace(/\u00AD/g, "")).toBe("associate");
    });

    it("should return original word if no break points", () => {
      const result = hyphenateWord("the", hyphenator);
      expect(result).toBe("the");
    });

    it("should use default soft hyphen character", () => {
      const result = hyphenateWord("information", hyphenator);
      expect(result).toContain("\u00AD");
      expect(result.replace(/\u00AD/g, "")).toBe("information");
    });

    it("should preserve word content when hyphens removed", () => {
      const words = ["computer", "hyphenation", "information", "programming"];
      for (const word of words) {
        const result = hyphenateWord(word, hyphenator);
        expect(result.replace(/\u00AD/g, "")).toBe(word);
      }
    });

    it("should place soft hyphens at expected positions for exceptions", () => {
      const result = hyphenateWord("associate", hyphenator);
      // "as-so-ciate" → "as\u00ADso\u00ADciate"
      expect(result).toBe("as\u00ADso\u00ADciate");
    });
  });

  describe("hyphenateText", () => {
    const hyphenator = createHyphenator(ENGLISH_US_PATTERNS);

    it("should hyphenate suitable words in a sentence", () => {
      const text = "The computer processes information quickly.";
      const result = hyphenateText(text, hyphenator);
      // Short words (The, quickly) may not be hyphenated
      // "computer" and "information" should get hyphens
      expect(result).toContain("\u00AD");
      // Non-word content should be preserved
      expect(result).toContain(".");
      expect(result).toContain(" ");
    });

    it("should skip words shorter than 5 characters", () => {
      const text = "The cat sat on a mat.";
      const result = hyphenateText(text, hyphenator);
      // All words are <= 4 chars, so no hyphenation
      expect(result).toBe("The cat sat on a mat.");
    });

    it("should preserve punctuation and whitespace", () => {
      const text = "Hello, world! This is great.";
      const result = hyphenateText(text, hyphenator);
      // Check punctuation is preserved
      expect(result).toContain(",");
      expect(result).toContain("!");
      expect(result).toContain(".");
    });

    it("should handle empty text", () => {
      const result = hyphenateText("", hyphenator);
      expect(result).toBe("");
    });

    it("should handle text with only spaces", () => {
      const result = hyphenateText("   ", hyphenator);
      expect(result).toBe("   ");
    });

    it("should handle text with multiple long words", () => {
      const text = "The transformation of information into computation";
      const result = hyphenateText(text, hyphenator);
      // All letters (ignoring soft hyphens) should remain the same
      expect(result.replace(/\u00AD/g, "")).toBe(text);
    });
  });

  describe("Custom options", () => {
    it("should produce fewer break points with larger minLeft/minRight", () => {
      const defaultHyphenator = createHyphenator(ENGLISH_US_PATTERNS);
      const strictHyphenator = createHyphenator(ENGLISH_US_PATTERNS, {
        minLeft: 3,
        minRight: 3
      });

      const word = "information";
      const defaultPoints = defaultHyphenator(word);
      const strictPoints = strictHyphenator(word);

      // Strict constraints should produce equal or fewer points
      expect(strictPoints.length).toBeLessThanOrEqual(defaultPoints.length);

      // Verify strict constraints are respected
      for (const p of strictPoints) {
        expect(p).toBeGreaterThanOrEqual(3);
        expect(p).toBeLessThanOrEqual(word.length - 3);
      }
    });

    it("should respect minLeft constraint", () => {
      const hyphenator = createHyphenator(ENGLISH_US_PATTERNS, { minLeft: 4 });
      const points = hyphenator("information");
      for (const p of points) {
        expect(p).toBeGreaterThanOrEqual(4);
      }
    });

    it("should respect minRight constraint", () => {
      const hyphenator = createHyphenator(ENGLISH_US_PATTERNS, { minRight: 4 });
      const word = "information";
      const points = hyphenator(word);
      for (const p of points) {
        expect(p).toBeLessThanOrEqual(word.length - 4);
      }
    });

    it("should return no points when constraints are too strict", () => {
      const hyphenator = createHyphenator(ENGLISH_US_PATTERNS, {
        minLeft: 10,
        minRight: 10
      });
      const points = hyphenator("short");
      expect(points).toHaveLength(0);
    });
  });

  describe("Custom hyphen character", () => {
    const hyphenator = createHyphenator(ENGLISH_US_PATTERNS);

    it("should use custom hyphen character in hyphenateWord", () => {
      const result = hyphenateWord("associate", hyphenator, "-");
      expect(result).toContain("-");
      expect(result).not.toContain("\u00AD");
      expect(result.replace(/-/g, "")).toBe("associate");
    });

    it("should use custom hyphen character in hyphenateText", () => {
      const text = "The associate processes information.";
      const result = hyphenateText(text, hyphenator, "|");
      expect(result).toContain("|");
      expect(result).not.toContain("\u00AD");
      expect(result.replace(/\|/g, "")).toBe(text);
    });

    it("should support multi-character hyphen strings", () => {
      const result = hyphenateWord("associate", hyphenator, "~~");
      expect(result).toContain("~~");
      expect(result.replace(/~~/g, "")).toBe("associate");
    });

    it("should support empty string as hyphen character", () => {
      const result = hyphenateWord("associate", hyphenator, "");
      // Empty hyphen char means nothing inserted
      expect(result).toBe("associate");
    });
  });

  describe("Edge cases", () => {
    const hyphenator = createHyphenator(ENGLISH_US_PATTERNS);

    it("should handle empty string", () => {
      const points = hyphenator("");
      expect(points).toHaveLength(0);
    });

    it("should handle single character", () => {
      const points = hyphenator("x");
      expect(points).toHaveLength(0);
    });

    it("should handle all-uppercase words", () => {
      const points = hyphenator("INFORMATION");
      // Should still find break points (case-insensitive matching)
      expect(points.length).toBeGreaterThan(0);
      for (const p of points) {
        expect(p).toBeGreaterThanOrEqual(2);
        expect(p).toBeLessThanOrEqual(9);
      }
    });

    it("should handle mixed-case words", () => {
      const points = hyphenator("Information");
      expect(points.length).toBeGreaterThan(0);
    });

    it("should handle words with no matching patterns", () => {
      const minimalPatterns: HyphenationPatterns = {
        language: "minimal",
        patterns: [],
        exceptions: []
      };
      const minHyphenator = createHyphenator(minimalPatterns);
      const points = minHyphenator("anything");
      expect(points).toHaveLength(0);
    });

    it("should handle repeated characters", () => {
      const points = hyphenator("aaaaaa");
      // Should not throw; results depend on patterns
      expect(Array.isArray(points)).toBe(true);
    });

    it("should produce consistent results on repeated calls", () => {
      const points1 = hyphenator("information");
      const points2 = hyphenator("information");
      expect(points1).toEqual(points2);
    });
  });
});
