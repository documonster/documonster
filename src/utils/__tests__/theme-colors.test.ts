import {
  hexToRgb01,
  applyTint,
  applyShade,
  applyTintRgb01,
  resolveOoxmlThemeColor,
  DEFAULT_OFFICE_THEME,
  CHART_THEME_PALETTE,
  CELL_THEME_PALETTE
} from "@utils/theme-colors";
import { describe, it, expect } from "vitest";

describe("@utils/theme-colors", () => {
  describe("hexToRgb01", () => {
    it("parses 6-digit RRGGBB", () => {
      expect(hexToRgb01("FF0000")).toEqual({ r: 1, g: 0, b: 0 });
      expect(hexToRgb01("#00FF00")).toEqual({ r: 0, g: 1, b: 0 });
    });

    it("parses 8-digit AARRGGBB, surfacing alpha only when < 1", () => {
      // Fully opaque alpha (FF) is omitted.
      expect(hexToRgb01("FF0000FF")).toEqual({ r: 0, g: 0, b: 1 });
      // Half alpha (80 ≈ 0.502) is surfaced.
      const half = hexToRgb01("800000FF");
      expect(half?.r).toBe(0);
      expect(half?.b).toBe(1);
      expect(half?.a).toBeCloseTo(128 / 255, 5);
    });

    it("expands 3-digit CSS shorthand", () => {
      expect(hexToRgb01("abc")).toEqual(hexToRgb01("aabbcc"));
    });

    it("returns null for malformed input", () => {
      expect(hexToRgb01(undefined)).toBeNull();
      expect(hexToRgb01("")).toBeNull();
      expect(hexToRgb01("12")).toBeNull();
      expect(hexToRgb01("ZZZZZZ")).toBeNull();
    });
  });

  describe("applyTint (hex)", () => {
    it("lightens toward white for positive tint", () => {
      expect(applyTint("000000", 1)).toBe("FFFFFF");
      expect(applyTint("808080", 0)).toBe("808080");
    });
  });

  describe("applyShade (hex)", () => {
    it("darkens toward black", () => {
      expect(applyShade("FFFFFF", 0)).toBe("000000");
      expect(applyShade("FFFFFF", 1)).toBe("FFFFFF");
    });
  });

  describe("applyTintRgb01", () => {
    it("lightens toward white for positive tint", () => {
      expect(applyTintRgb01({ r: 0, g: 0, b: 0 }, 1)).toEqual({ r: 1, g: 1, b: 1 });
    });

    it("darkens toward black for negative tint", () => {
      expect(applyTintRgb01({ r: 1, g: 1, b: 1 }, -1)).toEqual({ r: 0, g: 0, b: 0 });
      expect(applyTintRgb01({ r: 1, g: 1, b: 1 }, -0.5)).toEqual({ r: 0.5, g: 0.5, b: 0.5 });
    });

    it("returns colour unchanged for tint 0 and preserves alpha", () => {
      expect(applyTintRgb01({ r: 0.4, g: 0.5, b: 0.6, a: 0.5 }, 0)).toEqual({
        r: 0.4,
        g: 0.5,
        b: 0.6,
        a: 0.5
      });
    });
  });

  describe("default Office theme (2013) palettes", () => {
    it("uses the modern 2013 Office defaults", () => {
      expect(DEFAULT_OFFICE_THEME.dk2).toBe("44546A");
      expect(DEFAULT_OFFICE_THEME.accent1).toBe("4472C4");
      expect(DEFAULT_OFFICE_THEME.accent6).toBe("70AD47");
    });

    it("orders the chart palette as dk1, lt1, dk2, lt2, accent1.. (CT_ColorMapping)", () => {
      expect(CHART_THEME_PALETTE[0]).toBe(DEFAULT_OFFICE_THEME.dk1); // black
      expect(CHART_THEME_PALETTE[1]).toBe(DEFAULT_OFFICE_THEME.lt1); // white
      expect(CHART_THEME_PALETTE[2]).toBe(DEFAULT_OFFICE_THEME.dk2);
      expect(CHART_THEME_PALETTE[3]).toBe(DEFAULT_OFFICE_THEME.lt2);
      expect(CHART_THEME_PALETTE[4]).toBe(DEFAULT_OFFICE_THEME.accent1);
    });

    it("orders the cell palette as lt1, dk1, lt2, dk2, accent1.. (SpreadsheetML)", () => {
      expect(CELL_THEME_PALETTE[0]).toBe(DEFAULT_OFFICE_THEME.lt1); // white
      expect(CELL_THEME_PALETTE[1]).toBe(DEFAULT_OFFICE_THEME.dk1); // black
      expect(CELL_THEME_PALETTE[2]).toBe(DEFAULT_OFFICE_THEME.lt2);
      expect(CELL_THEME_PALETTE[3]).toBe(DEFAULT_OFFICE_THEME.dk2);
      expect(CELL_THEME_PALETTE[4]).toBe(DEFAULT_OFFICE_THEME.accent1);
    });

    it("swaps only the first two slot-pairs between chart and cell order", () => {
      // Chart 0/1 = dk1/lt1; cell 0/1 = lt1/dk1 (background/text vs dark/light).
      expect(CELL_THEME_PALETTE[0]).toBe(CHART_THEME_PALETTE[1]);
      expect(CELL_THEME_PALETTE[1]).toBe(CHART_THEME_PALETTE[0]);
      expect(CELL_THEME_PALETTE[2]).toBe(CHART_THEME_PALETTE[3]);
      expect(CELL_THEME_PALETTE[3]).toBe(CHART_THEME_PALETTE[2]);
      // Accents 4..9 share the same order.
      for (let i = 4; i < CHART_THEME_PALETTE.length; i++) {
        expect(CELL_THEME_PALETTE[i]).toBe(CHART_THEME_PALETTE[i]);
      }
    });
  });

  describe("resolveOoxmlThemeColor", () => {
    it("resolves a scheme key with tint", () => {
      const scheme = { accent1: "4472C4" };
      expect(resolveOoxmlThemeColor("accent1", scheme)).toBe("4472C4");
      expect(resolveOoxmlThemeColor("accent1", scheme, 0)).toBe("4472C4");
    });
  });
});
