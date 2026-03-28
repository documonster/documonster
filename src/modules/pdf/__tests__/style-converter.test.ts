/**
 * Tests for style conversion (Excel -> PDF).
 */
import { describe, it, expect } from "vitest";
import {
  argbToPdfColor,
  excelColorToPdf,
  extractFontProperties,
  applyTint,
  excelFillToPdfColor,
  excelBordersToPdf,
  excelHAlignToPdf,
  excelVAlignToPdf,
  DEFAULT_COLORS
} from "@pdf/render/style-converter";

describe("Style Converter", () => {
  describe("argbToPdfColor", () => {
    it("should convert ARGB to PDF color", () => {
      const color = argbToPdfColor("FF000000"); // black
      expect(color).toEqual({ r: 0, g: 0, b: 0 });
    });

    it("should convert red ARGB", () => {
      const color = argbToPdfColor("FFFF0000"); // red
      expect(color).toEqual({ r: 1, g: 0, b: 0 });
    });

    it("should convert 6-char hex", () => {
      const color = argbToPdfColor("00FF00"); // green
      expect(color).toEqual({ r: 0, g: 1, b: 0 });
    });

    it("should handle white", () => {
      const color = argbToPdfColor("FFFFFFFF");
      expect(color).toEqual({ r: 1, g: 1, b: 1 });
    });

    it("should return null for undefined", () => {
      expect(argbToPdfColor(undefined)).toBeNull();
    });

    it("should return null for short strings", () => {
      expect(argbToPdfColor("FF")).toBeNull();
    });

    it("should handle mid-range values", () => {
      const color = argbToPdfColor("FF808080"); // gray
      expect(color!.r).toBeCloseTo(0.502, 2);
      expect(color!.g).toBeCloseTo(0.502, 2);
      expect(color!.b).toBeCloseTo(0.502, 2);
    });
  });

  describe("excelColorToPdf", () => {
    it("should convert ARGB color", () => {
      const color = excelColorToPdf({ argb: "FFFF0000" });
      expect(color).toEqual({ r: 1, g: 0, b: 0 });
    });

    it("should convert theme color", () => {
      const color = excelColorToPdf({ theme: 1 }); // dk1 -> black
      expect(color).toEqual({ r: 0, g: 0, b: 0 });
    });

    it("should prefer ARGB over theme", () => {
      const color = excelColorToPdf({ argb: "FFFF0000", theme: 1 });
      expect(color).toEqual({ r: 1, g: 0, b: 0 });
    });

    it("should return null for undefined", () => {
      expect(excelColorToPdf(undefined)).toBeNull();
    });

    it("should return null for empty object", () => {
      expect(excelColorToPdf({})).toBeNull();
    });

    it("should apply positive theme tint", () => {
      const color = excelColorToPdf({ theme: 1, tint: 0.5 } as never);
      expect(color).toEqual({ r: 0.5, g: 0.5, b: 0.5 });
    });

    it("should apply negative theme tint", () => {
      const color = excelColorToPdf({ theme: 0, tint: -0.25 } as never);
      expect(color).toEqual({ r: 0.75, g: 0.75, b: 0.75 });
    });
  });

  describe("applyTint", () => {
    it("should lighten colors", () => {
      const color = applyTint({ r: 0.2, g: 0.4, b: 0.6 }, 0.5);
      expect(color.r).toBeCloseTo(0.6, 10);
      expect(color.g).toBeCloseTo(0.7, 10);
      expect(color.b).toBeCloseTo(0.8, 10);
    });

    it("should darken colors", () => {
      const color = applyTint({ r: 0.2, g: 0.4, b: 0.6 }, -0.5);
      expect(color.r).toBeCloseTo(0.1, 10);
      expect(color.g).toBeCloseTo(0.2, 10);
      expect(color.b).toBeCloseTo(0.3, 10);
    });
  });

  describe("extractFontProperties", () => {
    it("should use defaults when font is undefined", () => {
      const props = extractFontProperties(undefined, "Helvetica", 11);
      expect(props.fontFamily).toBe("Helvetica");
      expect(props.fontSize).toBe(11);
      expect(props.bold).toBe(false);
      expect(props.italic).toBe(false);
      expect(props.strike).toBe(false);
      expect(props.underline).toBe(false);
    });

    it("should extract font properties", () => {
      const props = extractFontProperties(
        {
          name: "Arial",
          size: 14,
          bold: true,
          italic: true,
          strike: true,
          underline: "single",
          color: { argb: "FFFF0000" }
        },
        "Helvetica",
        11
      );
      expect(props.fontFamily).toBe("Arial");
      expect(props.fontSize).toBe(14);
      expect(props.bold).toBe(true);
      expect(props.italic).toBe(true);
      expect(props.strike).toBe(true);
      expect(props.underline).toBe(true);
      expect(props.textColor).toEqual({ r: 1, g: 0, b: 0 });
    });

    it("should treat 'none' underline as false", () => {
      const props = extractFontProperties({ underline: "none" }, "Helvetica", 11);
      expect(props.underline).toBe(false);
    });
  });

  describe("excelFillToPdfColor", () => {
    it("should convert solid fill", () => {
      const color = excelFillToPdfColor({
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0000FF" }
      });
      expect(color).toEqual({ r: 0, g: 0, b: 1 });
    });

    it("should return null for 'none' pattern", () => {
      const color = excelFillToPdfColor({ type: "pattern", pattern: "none" });
      expect(color).toBeNull();
    });

    it("should return null for undefined", () => {
      expect(excelFillToPdfColor(undefined)).toBeNull();
    });

    it("should handle gradient fills (use first stop)", () => {
      const color = excelFillToPdfColor({
        type: "gradient",
        stops: [
          { position: 0, color: { argb: "FFFF0000" } },
          { position: 1, color: { argb: "FF0000FF" } }
        ]
      });
      expect(color).toEqual({ r: 1, g: 0, b: 0 });
    });
  });

  describe("excelBordersToPdf", () => {
    it("should return null borders for undefined", () => {
      const borders = excelBordersToPdf(undefined);
      expect(borders.top).toBeNull();
      expect(borders.right).toBeNull();
      expect(borders.bottom).toBeNull();
      expect(borders.left).toBeNull();
    });

    it("should convert thin borders", () => {
      const borders = excelBordersToPdf({
        top: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "medium", color: { argb: "FF000000" } }
      });
      expect(borders.top).not.toBeNull();
      expect(borders.top!.width).toBe(0.5);
      expect(borders.bottom!.width).toBe(1);
      expect(borders.right).toBeNull();
    });

    it("should convert dashed border styles", () => {
      const borders = excelBordersToPdf({
        top: { style: "dashed", color: { argb: "FF000000" } }
      });
      expect(borders.top!.dashPattern.length).toBeGreaterThan(0);
    });
  });

  describe("excelHAlignToPdf", () => {
    it("should map left alignment", () => {
      expect(excelHAlignToPdf({ horizontal: "left" })).toBe("left");
    });

    it("should map center alignment", () => {
      expect(excelHAlignToPdf({ horizontal: "center" })).toBe("center");
    });

    it("should map right alignment", () => {
      expect(excelHAlignToPdf({ horizontal: "right" })).toBe("right");
    });

    it("should map centerContinuous to center", () => {
      expect(excelHAlignToPdf({ horizontal: "centerContinuous" })).toBe("center");
    });

    it("should default to left", () => {
      expect(excelHAlignToPdf(undefined)).toBe("left");
      expect(excelHAlignToPdf({})).toBe("left");
    });
  });

  describe("excelVAlignToPdf", () => {
    it("should map top alignment", () => {
      expect(excelVAlignToPdf({ vertical: "top" })).toBe("top");
    });

    it("should map middle alignment", () => {
      expect(excelVAlignToPdf({ vertical: "middle" })).toBe("middle");
    });

    it("should map bottom alignment", () => {
      expect(excelVAlignToPdf({ vertical: "bottom" })).toBe("bottom");
    });

    it("should default to bottom (Excel default)", () => {
      expect(excelVAlignToPdf(undefined)).toBe("bottom");
      expect(excelVAlignToPdf({})).toBe("bottom");
    });
  });

  describe("DEFAULT_COLORS", () => {
    it("should have black", () => {
      expect(DEFAULT_COLORS.black).toEqual({ r: 0, g: 0, b: 0 });
    });

    it("should have white", () => {
      expect(DEFAULT_COLORS.white).toEqual({ r: 1, g: 1, b: 1 });
    });
  });
});
