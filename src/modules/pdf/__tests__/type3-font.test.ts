import { describe, it, expect } from "vitest";

import { lookupGlyph, hasGlyph, glyphCount, NOTDEF_GLYPH } from "../font/type3-glyphs";

describe("Type3 fallback glyphs", () => {
  describe("lookupGlyph", () => {
    it("should return non-null for WHITE CIRCLE (U+25CB)", () => {
      const glyph = lookupGlyph(0x25cb);
      expect(glyph).toBeDefined();
      expect(glyph!.width).toBeGreaterThan(0);
    });

    it("should return non-null for BALLOT BOX (U+2610)", () => {
      const glyph = lookupGlyph(0x2610);
      expect(glyph).toBeDefined();
      expect(glyph!.width).toBeGreaterThan(0);
    });

    it("should return non-null for SQUARED SMALL CIRCLE (U+29C7)", () => {
      const glyph = lookupGlyph(0x29c7);
      expect(glyph).toBeDefined();
      expect(glyph!.width).toBeGreaterThan(0);
    });

    it("should return undefined for ASCII 'A' (U+0041)", () => {
      const glyph = lookupGlyph(0x0041);
      expect(glyph).toBeUndefined();
    });
  });

  describe("glyphCount", () => {
    it("should have at least 2800 glyphs", () => {
      expect(glyphCount()).toBeGreaterThanOrEqual(2800);
    });
  });

  describe("hasGlyph", () => {
    it("should return true for CHECK MARK (U+2713)", () => {
      expect(hasGlyph(0x2713)).toBe(true);
    });

    it("should return false for SPACE (U+0020)", () => {
      expect(hasGlyph(0x0020)).toBe(false);
    });
  });

  describe("NOTDEF_GLYPH", () => {
    it("should have width > 0", () => {
      expect(NOTDEF_GLYPH.width).toBeGreaterThan(0);
    });
  });
});
