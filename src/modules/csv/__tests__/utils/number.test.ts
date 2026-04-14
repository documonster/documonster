/**
 * CSV Number Utilities Unit Tests
 *
 * Tests for number formatting and parsing with different decimal separators:
 * - formatNumberForCsv: Format numbers for CSV output
 * - parseNumberFromCsv: Parse numbers from CSV input
 */

import { formatNumberForCsv, parseNumberFromCsv } from "@csv/utils/number";
import { describe, it, expect } from "vitest";

// =============================================================================
// formatNumberForCsv Tests
// =============================================================================

describe("formatNumberForCsv", () => {
  describe("standard decimal point (.)", () => {
    it("formats integers", () => {
      expect(formatNumberForCsv(0, ".")).toBe("0");
      expect(formatNumberForCsv(42, ".")).toBe("42");
      expect(formatNumberForCsv(-100, ".")).toBe("-100");
    });

    it("formats floating point numbers", () => {
      expect(formatNumberForCsv(3.14, ".")).toBe("3.14");
      expect(formatNumberForCsv(0.5, ".")).toBe("0.5");
      expect(formatNumberForCsv(-2.718, ".")).toBe("-2.718");
    });

    it("formats scientific notation", () => {
      expect(formatNumberForCsv(1e10, ".")).toBe("10000000000");
      expect(formatNumberForCsv(1.5e-5, ".")).toBe("0.000015");
    });

    it("formats special values", () => {
      expect(formatNumberForCsv(Infinity, ".")).toBe("Infinity");
      expect(formatNumberForCsv(-Infinity, ".")).toBe("-Infinity");
      expect(formatNumberForCsv(NaN, ".")).toBe("NaN");
    });
  });

  describe("European decimal comma (,)", () => {
    it("formats integers unchanged", () => {
      expect(formatNumberForCsv(0, ",")).toBe("0");
      expect(formatNumberForCsv(42, ",")).toBe("42");
      expect(formatNumberForCsv(-100, ",")).toBe("-100");
    });

    it("replaces decimal point with comma", () => {
      expect(formatNumberForCsv(3.14, ",")).toBe("3,14");
      expect(formatNumberForCsv(0.5, ",")).toBe("0,5");
      expect(formatNumberForCsv(-2.718, ",")).toBe("-2,718");
    });

    it("formats high precision numbers", () => {
      expect(formatNumberForCsv(1.23456789, ",")).toBe("1,23456789");
    });

    it("formats special values", () => {
      expect(formatNumberForCsv(Infinity, ",")).toBe("Infinity");
      expect(formatNumberForCsv(-Infinity, ",")).toBe("-Infinity");
      expect(formatNumberForCsv(NaN, ",")).toBe("NaN");
    });
  });
});

// =============================================================================
// parseNumberFromCsv Tests
// =============================================================================

describe("parseNumberFromCsv", () => {
  describe("standard decimal point (.)", () => {
    it("parses integers", () => {
      expect(parseNumberFromCsv("0", ".")).toBe(0);
      expect(parseNumberFromCsv("42", ".")).toBe(42);
      expect(parseNumberFromCsv("-100", ".")).toBe(-100);
    });

    it("parses floating point numbers", () => {
      expect(parseNumberFromCsv("3.14", ".")).toBe(3.14);
      expect(parseNumberFromCsv("0.5", ".")).toBe(0.5);
      expect(parseNumberFromCsv("-2.718", ".")).toBe(-2.718);
    });

    it("parses scientific notation", () => {
      expect(parseNumberFromCsv("1e10", ".")).toBe(1e10);
      expect(parseNumberFromCsv("1.5e-5", ".")).toBe(1.5e-5);
    });

    it("returns NaN for invalid strings", () => {
      expect(parseNumberFromCsv("abc", ".")).toBeNaN();
    });

    it("returns 0 for empty string (Number behavior)", () => {
      expect(parseNumberFromCsv("", ".")).toBe(0);
    });
  });

  describe("European decimal comma (,)", () => {
    it("parses integers", () => {
      expect(parseNumberFromCsv("0", ",")).toBe(0);
      expect(parseNumberFromCsv("42", ",")).toBe(42);
      expect(parseNumberFromCsv("-100", ",")).toBe(-100);
    });

    it("parses numbers with comma as decimal separator", () => {
      expect(parseNumberFromCsv("3,14", ",")).toBe(3.14);
      expect(parseNumberFromCsv("0,5", ",")).toBe(0.5);
      expect(parseNumberFromCsv("-2,718", ",")).toBe(-2.718);
    });

    it("parses with scientific notation", () => {
      expect(parseNumberFromCsv("1,5e-5", ",")).toBe(1.5e-5);
      expect(parseNumberFromCsv("2,5E+3", ",")).toBe(2500);
    });

    it("handles whitespace", () => {
      expect(parseNumberFromCsv("  3,14  ", ",")).toBe(3.14);
      expect(parseNumberFromCsv("\t42\t", ",")).toBe(42);
    });

    it("returns NaN for invalid strings", () => {
      expect(parseNumberFromCsv("abc", ",")).toBeNaN();
    });

    it("returns 0 for empty string (Number behavior)", () => {
      expect(parseNumberFromCsv("", ",")).toBe(0);
    });
  });

  describe("round-trip consistency", () => {
    it("formats and parses back to same value (standard)", () => {
      const values = [0, 42, -100, 3.14, 0.001, -99.99];
      for (const v of values) {
        const formatted = formatNumberForCsv(v, ".");
        const parsed = parseNumberFromCsv(formatted, ".");
        expect(parsed).toBe(v);
      }
    });

    it("formats and parses back to same value (European)", () => {
      const values = [0, 42, -100, 3.14, 0.001, -99.99];
      for (const v of values) {
        const formatted = formatNumberForCsv(v, ",");
        const parsed = parseNumberFromCsv(formatted, ",");
        expect(parsed).toBe(v);
      }
    });
  });
});
