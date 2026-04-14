import {
  pdfString,
  pdfHexString,
  pdfName,
  pdfNumber,
  pdfBoolean,
  pdfArray,
  pdfRef,
  pdfDate,
  PdfDict
} from "@pdf/core/pdf-object";
/**
 * Tests for PDF low-level object serialization.
 */
import { describe, it, expect } from "vitest";

describe("PDF Object Primitives", () => {
  describe("pdfString", () => {
    it("should wrap in parentheses", () => {
      expect(pdfString("Hello")).toBe("(Hello)");
    });

    it("should escape backslashes", () => {
      expect(pdfString("a\\b")).toBe("(a\\\\b)");
    });

    it("should escape parentheses", () => {
      expect(pdfString("a(b)c")).toBe("(a\\(b\\)c)");
    });

    it("should escape newlines and carriage returns", () => {
      expect(pdfString("line1\nline2")).toBe("(line1\\nline2)");
      expect(pdfString("line1\rline2")).toBe("(line1\\nline2)");
    });

    it("should handle empty string", () => {
      expect(pdfString("")).toBe("()");
    });

    it("should encode non-ASCII strings as UTF-16BE hex", () => {
      expect(pdfString("报告")).toBe("<feff62a5544a>");
    });
  });

  describe("pdfHexString", () => {
    it("should format as hex with angle brackets", () => {
      const bytes = new Uint8Array([0x48, 0x65, 0x6c, 0x6c, 0x6f]);
      expect(pdfHexString(bytes)).toBe("<48656c6c6f>");
    });

    it("should pad single-digit hex values", () => {
      const bytes = new Uint8Array([0x0a, 0x0b]);
      expect(pdfHexString(bytes)).toBe("<0a0b>");
    });

    it("should handle empty bytes", () => {
      expect(pdfHexString(new Uint8Array([]))).toBe("<>");
    });
  });

  describe("pdfName", () => {
    it("should prefix with /", () => {
      expect(pdfName("Type")).toBe("/Type");
    });

    it("should encode special characters", () => {
      const result = pdfName("Name With Space");
      expect(result).toContain("#20");
    });

    it("should encode # character", () => {
      const result = pdfName("A#B");
      expect(result).toBe("/A#23B");
    });

    it("should handle simple ASCII names", () => {
      expect(pdfName("Font")).toBe("/Font");
      expect(pdfName("BaseFont")).toBe("/BaseFont");
    });
  });

  describe("pdfNumber", () => {
    it("should format integers without decimal", () => {
      expect(pdfNumber(42)).toBe("42");
      expect(pdfNumber(0)).toBe("0");
      expect(pdfNumber(-5)).toBe("-5");
    });

    it("should format floats with precision", () => {
      const result = pdfNumber(3.14159);
      expect(result).toBe("3.1416");
    });

    it("should handle very small precision", () => {
      // Should not produce floating point artifacts
      const result = pdfNumber(0.1 + 0.2);
      expect(parseFloat(result)).toBeCloseTo(0.3, 4);
    });

    it("should return 0 for NaN", () => {
      expect(pdfNumber(NaN)).toBe("0");
    });

    it("should return 0 for Infinity", () => {
      expect(pdfNumber(Infinity)).toBe("0");
      expect(pdfNumber(-Infinity)).toBe("0");
    });
  });

  describe("pdfBoolean", () => {
    it("should return true/false strings", () => {
      expect(pdfBoolean(true)).toBe("true");
      expect(pdfBoolean(false)).toBe("false");
    });
  });

  describe("pdfArray", () => {
    it("should format as space-separated in brackets", () => {
      expect(pdfArray(["1", "2", "3"])).toBe("[1 2 3]");
    });

    it("should handle empty array", () => {
      expect(pdfArray([])).toBe("[]");
    });

    it("should handle single element", () => {
      expect(pdfArray(["42"])).toBe("[42]");
    });
  });

  describe("pdfRef", () => {
    it("should format indirect reference", () => {
      expect(pdfRef(5)).toBe("5 0 R");
      expect(pdfRef(12, 0)).toBe("12 0 R");
    });
  });

  describe("pdfDate", () => {
    it("should format as PDF date string", () => {
      const date = new Date(Date.UTC(2024, 0, 15, 10, 30, 0)); // Jan 15, 2024 10:30 UTC
      const result = pdfDate(date);
      expect(result).toBe("(D:20240115103000Z)");
    });
  });

  describe("PdfDict", () => {
    it("should build a dictionary", () => {
      const dict = new PdfDict().set("Type", "/Catalog").set("Pages", "5 0 R");

      const result = dict.toString();
      expect(result).toContain("<<");
      expect(result).toContain(">>");
      expect(result).toContain("/Type /Catalog");
      expect(result).toContain("/Pages 5 0 R");
    });

    it("should handle conditional entries", () => {
      const dict = new PdfDict()
        .set("Type", "/Page")
        .setIf(true, "Width", "100")
        .setIf(false, "Height", "200");

      const result = dict.toString();
      expect(result).toContain("/Width 100");
      expect(result).not.toContain("/Height");
    });

    it("should handle empty dictionary", () => {
      const dict = new PdfDict();
      const result = dict.toString();
      expect(result).toContain("<<");
      expect(result).toContain(">>");
    });

    it("should overwrite duplicate keys", () => {
      const dict = new PdfDict().set("Type", "/Page").set("Type", "/Catalog");

      const result = dict.toString();
      expect(result).toContain("/Type /Catalog");
      expect(result).not.toContain("/Type /Page");
      // Should only have one /Type entry
      const typeCount = (result.match(/\/Type/g) || []).length;
      expect(typeCount).toBe(1);
    });
  });
});
