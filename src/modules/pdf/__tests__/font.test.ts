/**
 * Tests for PDF font metrics and font manager.
 */
import { describe, it, expect } from "vitest";
import {
  measureText,
  getFontAscent,
  getFontDescent,
  getLineHeight,
  getCharWidth,
  isStandardFont,
  getStandardFontNames
} from "@pdf/font/metrics";
import { FontManager, resolvePdfFontName } from "@pdf/font/font-manager";

describe("Font Metrics", () => {
  describe("getCharWidth", () => {
    it("should return width for ASCII space in Helvetica", () => {
      const width = getCharWidth(32, "Helvetica"); // space
      expect(width).toBe(278);
    });

    it("should return width for letter A in Helvetica", () => {
      const width = getCharWidth(65, "Helvetica"); // 'A'
      expect(width).toBe(667);
    });

    it("should return monospace width for Courier", () => {
      const widthA = getCharWidth(65, "Courier");
      const widthZ = getCharWidth(90, "Courier");
      expect(widthA).toBe(600);
      expect(widthZ).toBe(600);
    });

    it("should fallback to Helvetica for unknown fonts", () => {
      const width = getCharWidth(65, "UnknownFont");
      expect(width).toBe(getCharWidth(65, "Helvetica"));
    });

    it("should return average width for non-Latin characters", () => {
      const width = getCharWidth(0x4e2d, "Helvetica"); // Chinese character
      expect(width).toBe(513); // average width
    });
  });

  describe("measureText", () => {
    it("should return 0 for empty string", () => {
      expect(measureText("", "Helvetica", 12)).toBe(0);
    });

    it("should measure a simple word", () => {
      const width = measureText("Hello", "Helvetica", 12);
      expect(width).toBeGreaterThan(0);
      expect(width).toBeLessThan(100);
    });

    it("should scale with font size", () => {
      const width12 = measureText("Test", "Helvetica", 12);
      const width24 = measureText("Test", "Helvetica", 24);
      expect(width24).toBeCloseTo(width12 * 2, 1);
    });

    it("should vary by font family", () => {
      const helvetica = measureText("Test", "Helvetica", 12);
      const times = measureText("Test", "Times-Roman", 12);
      expect(helvetica).not.toBe(times);
    });

    it("should give same width for all Courier chars", () => {
      const widthA = measureText("A", "Courier", 12);
      const widthW = measureText("W", "Courier", 12);
      expect(widthA).toBe(widthW);
    });
  });

  describe("getFontAscent", () => {
    it("should return positive ascent", () => {
      const ascent = getFontAscent("Helvetica", 12);
      expect(ascent).toBeGreaterThan(0);
    });

    it("should scale with font size", () => {
      const a12 = getFontAscent("Helvetica", 12);
      const a24 = getFontAscent("Helvetica", 24);
      expect(a24).toBeCloseTo(a12 * 2, 1);
    });
  });

  describe("getFontDescent", () => {
    it("should return negative descent", () => {
      const descent = getFontDescent("Helvetica", 12);
      expect(descent).toBeLessThan(0);
    });
  });

  describe("getLineHeight", () => {
    it("should return ascent minus descent", () => {
      const lineHeight = getLineHeight("Helvetica", 12);
      const ascent = getFontAscent("Helvetica", 12);
      const descent = getFontDescent("Helvetica", 12);
      expect(lineHeight).toBeCloseTo(ascent - descent, 4);
    });
  });

  describe("isStandardFont", () => {
    it("should recognize standard fonts", () => {
      expect(isStandardFont("Helvetica")).toBe(true);
      expect(isStandardFont("Helvetica-Bold")).toBe(true);
      expect(isStandardFont("Times-Roman")).toBe(true);
      expect(isStandardFont("Courier")).toBe(true);
    });

    it("should reject non-standard fonts", () => {
      expect(isStandardFont("Arial")).toBe(false);
      expect(isStandardFont("Calibri")).toBe(false);
    });
  });

  describe("getStandardFontNames", () => {
    it("should return all 12 standard fonts", () => {
      const names = getStandardFontNames();
      expect(names.length).toBe(12);
      expect(names).toContain("Helvetica");
      expect(names).toContain("Helvetica-Bold");
      expect(names).toContain("Helvetica-Oblique");
      expect(names).toContain("Helvetica-BoldOblique");
      expect(names).toContain("Times-Roman");
      expect(names).toContain("Times-Bold");
      expect(names).toContain("Times-Italic");
      expect(names).toContain("Times-BoldItalic");
      expect(names).toContain("Courier");
      expect(names).toContain("Courier-Bold");
      expect(names).toContain("Courier-Oblique");
      expect(names).toContain("Courier-BoldOblique");
    });
  });
});

describe("Font Name Resolution", () => {
  describe("resolvePdfFontName", () => {
    it("should map Arial to Helvetica family", () => {
      expect(resolvePdfFontName("Arial", false, false)).toBe("Helvetica");
      expect(resolvePdfFontName("Arial", true, false)).toBe("Helvetica-Bold");
      expect(resolvePdfFontName("Arial", false, true)).toBe("Helvetica-Oblique");
      expect(resolvePdfFontName("Arial", true, true)).toBe("Helvetica-BoldOblique");
    });

    it("should map Calibri to Helvetica family", () => {
      expect(resolvePdfFontName("Calibri", false, false)).toBe("Helvetica");
      expect(resolvePdfFontName("Calibri", true, false)).toBe("Helvetica-Bold");
    });

    it("should map Times New Roman to Times family", () => {
      expect(resolvePdfFontName("Times New Roman", false, false)).toBe("Times-Roman");
      expect(resolvePdfFontName("Times New Roman", true, false)).toBe("Times-Bold");
      expect(resolvePdfFontName("Times New Roman", false, true)).toBe("Times-Italic");
      expect(resolvePdfFontName("Times New Roman", true, true)).toBe("Times-BoldItalic");
    });

    it("should map Courier New to Courier family", () => {
      expect(resolvePdfFontName("Courier New", false, false)).toBe("Courier");
      expect(resolvePdfFontName("Courier New", true, false)).toBe("Courier-Bold");
      expect(resolvePdfFontName("Courier New", false, true)).toBe("Courier-Oblique");
    });

    it("should map Consolas to Courier family", () => {
      expect(resolvePdfFontName("Consolas", false, false)).toBe("Courier");
    });

    it("should fall back to Helvetica for unknown fonts", () => {
      expect(resolvePdfFontName("FancyFont", false, false)).toBe("Helvetica");
      expect(resolvePdfFontName("FancyFont", true, false)).toBe("Helvetica-Bold");
    });

    it("should be case-insensitive", () => {
      expect(resolvePdfFontName("ARIAL", false, false)).toBe("Helvetica");
      expect(resolvePdfFontName("times new roman", true, true)).toBe("Times-BoldItalic");
    });
  });
});

describe("FontManager", () => {
  it("should register and track fonts", () => {
    const fm = new FontManager();
    const r1 = fm.ensureFont("Helvetica");
    const r2 = fm.ensureFont("Helvetica-Bold");
    const r3 = fm.ensureFont("Helvetica"); // should return same as r1

    expect(r1).toBe("F1");
    expect(r2).toBe("F2");
    expect(r3).toBe("F1"); // same font, same resource name
  });

  it("should resolve Excel font names", () => {
    const fm = new FontManager();
    const r1 = fm.resolveFont("Arial", false, false);
    const r2 = fm.resolveFont("Arial", true, false);

    expect(r1).not.toBe(r2);
    expect(fm.getPdfFontName(r1)).toBe("Helvetica");
    expect(fm.getPdfFontName(r2)).toBe("Helvetica-Bold");
  });

  it("should measure text through font manager", () => {
    const fm = new FontManager();
    const resource = fm.ensureFont("Helvetica");
    const width = fm.measureText("Hello", resource, 12);
    expect(width).toBeGreaterThan(0);
  });

  it("should return font ascent/descent", () => {
    const fm = new FontManager();
    const resource = fm.ensureFont("Helvetica");
    expect(fm.getFontAscent(resource, 12)).toBeGreaterThan(0);
    expect(fm.getFontDescent(resource, 12)).toBeLessThan(0);
    expect(fm.getLineHeight(resource, 12)).toBeGreaterThan(0);
  });

  it("should list registered fonts", () => {
    const fm = new FontManager();
    fm.ensureFont("Helvetica");
    fm.ensureFont("Courier");

    const fonts = fm.getRegisteredFonts();
    expect(fonts).toHaveLength(2);
    expect(fonts.find(f => f.pdfFontName === "Helvetica")).toBeDefined();
    expect(fonts.find(f => f.pdfFontName === "Courier")).toBeDefined();
  });
});
