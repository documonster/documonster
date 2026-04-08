import { describe, it, expect } from "vitest";
import { format, cellFormat } from "@excel/utils/cell-format";

describe("cell-format", () => {
  describe("format", () => {
    describe("General format", () => {
      it("should format integers", () => {
        expect(format("General", 123)).toBe("123");
        expect(format("General", -456)).toBe("-456");
        expect(format("General", 0)).toBe("0");
      });

      it("should format decimals", () => {
        expect(format("General", 123.456)).toBe("123.456");
        expect(format("General", 0.1)).toBe("0.1");
      });

      it("should format strings", () => {
        expect(format("General", "hello")).toBe("hello");
        expect(format("General", "")).toBe("");
      });

      it("should format booleans", () => {
        expect(format("General", true)).toBe("TRUE");
        expect(format("General", false)).toBe("FALSE");
      });
    });

    describe("Percentage format", () => {
      it("should format basic percentages", () => {
        expect(format("0%", 0.25)).toBe("25%");
        expect(format("0%", 1)).toBe("100%");
        expect(format("0%", 0)).toBe("0%");
      });

      it("should format percentages with decimals", () => {
        expect(format("0.00%", 0.25)).toBe("25.00%");
        expect(format("0.00%", 0.1234)).toBe("12.34%");
        expect(format("0.0%", 0.256)).toBe("25.6%");
      });

      it("should format negative percentages", () => {
        expect(format("0%", -0.25)).toBe("-25%");
        expect(format("0.00%", -0.1234)).toBe("-12.34%");
      });
    });

    describe("Number format with decimals", () => {
      it("should format with fixed decimal places", () => {
        expect(format("0.00", 123.456)).toBe("123.46");
        expect(format("0.00", 123)).toBe("123.00");
        expect(format("0.0", 1.234)).toBe("1.2");
      });

      it("should format integers", () => {
        expect(format("0", 123.456)).toBe("123");
        expect(format("0", 123)).toBe("123");
      });
    });

    describe("Number format with thousand separators", () => {
      it("should add thousand separators", () => {
        expect(format("#,##0", 1234)).toBe("1,234");
        expect(format("#,##0", 1234567)).toBe("1,234,567");
        expect(format("#,##0", 123)).toBe("123");
      });

      it("should add thousand separators with decimals", () => {
        expect(format("#,##0.00", 1234.56)).toBe("1,234.56");
        expect(format("#,##0.00", 1234567.89)).toBe("1,234,567.89");
      });
    });

    describe("Date format", () => {
      // Excel serial number for 2025-10-22 is approximately 45952
      const dateSerial = 45952; // 2025-10-22

      it("should format yyyy-mm-dd", () => {
        expect(format("yyyy-mm-dd", dateSerial)).toBe("2025-10-22");
      });

      it("should format dd/mm/yyyy", () => {
        expect(format("dd/mm/yyyy", dateSerial)).toBe("22/10/2025");
      });

      it("should format m/d/yy", () => {
        expect(format("m/d/yy", dateSerial)).toBe("10/22/25");
      });

      it("should format with month names", () => {
        expect(format("d-mmm-yy", dateSerial)).toBe("22-Oct-25");
        expect(format("mmmm d, yyyy", dateSerial)).toBe("October 22, 2025");
      });

      it("should format with day names", () => {
        expect(format("ddd", dateSerial)).toBe("Wed");
        expect(format("dddd", dateSerial)).toBe("Wednesday");
      });
    });

    describe("Currency format", () => {
      it("should format with dollar sign", () => {
        expect(format("$#,##0.00", 1234.56)).toBe("$1,234.56");
        expect(format("$#,##0", 1234)).toBe("$1,234");
      });
    });

    describe("Negative number handling", () => {
      it("should format negative numbers with minus sign", () => {
        expect(format("#,##0", -1234)).toBe("-1,234");
        expect(format("0.00", -123.45)).toBe("-123.45");
      });

      it("should handle multi-section formats", () => {
        expect(format("#,##0;(#,##0)", 1234)).toBe("1,234");
        expect(format("#,##0;(#,##0)", -1234)).toBe("(1,234)");
      });
    });

    describe("Trailing commas (scale by 1000)", () => {
      it("should scale numbers by 1000 for each trailing comma", () => {
        expect(format("#,##0,", 1234000)).toBe("1,234");
        expect(format("#,##0,,", 1234000000)).toBe("1,234");
      });
    });

    describe("Leading zeros", () => {
      it("should pad with leading zeros", () => {
        expect(format("00000", 123)).toBe("00123");
        expect(format("000", 7)).toBe("007");
      });
    });

    describe("Color codes", () => {
      it("should ignore color codes", () => {
        expect(format("[Red]0.00", 123.45)).toBe("123.45");
        expect(format("[Green]#,##0", 1234)).toBe("1,234");
      });
    });

    describe("Scientific notation", () => {
      it("should format basic scientific notation", () => {
        expect(format("0.00E+00", 1234)).toBe("1.23E+03");
        expect(format("0.00E+00", 0.00123)).toBe("1.23E-03");
      });

      it("should handle zero", () => {
        expect(format("0.00E+00", 0)).toBe("0.00E+00");
      });

      it("should handle negative numbers", () => {
        expect(format("0.00E+00", -1234)).toBe("-1.23E+03");
      });
    });

    describe("Fraction format", () => {
      it("should format as fraction with fixed denominator", () => {
        expect(format("# ?/8", 1.5)).toBe("1 4/8");
        expect(format("# ?/4", 0.25)).toBe("1/4");
      });

      it("should format as fraction with variable denominator", () => {
        expect(format("# ?/?", 1.5)).toBe("1 1/2");
        expect(format("# ??/??", 0.333)).toBe("1/3");
      });

      it("should handle whole numbers", () => {
        expect(format("# ?/?", 5)).toBe("5");
      });
    });

    describe("Elapsed time format", () => {
      it("should format elapsed hours", () => {
        // 1.5 days = 36 hours
        expect(format("[h]:mm:ss", 1.5)).toBe("36:00:00");
      });

      it("should format elapsed minutes", () => {
        // 0.5 days = 720 minutes
        expect(format("[m]:ss", 0.5)).toBe("720:00");
      });
    });

    describe("Text placeholder", () => {
      it("should handle @ placeholder for numbers", () => {
        expect(format("@", 123)).toBe("123");
      });

      it("should handle @ placeholder in text format section", () => {
        expect(format('0;0;0;"Text: "@', "hello")).toBe("Text: hello");
      });
    });

    describe("Edge cases", () => {
      it("should handle zero", () => {
        expect(format("0.00", 0)).toBe("0.00");
        expect(format("#,##0", 0)).toBe("0");
      });

      it("should handle very small numbers", () => {
        expect(format("0.00", 0.001)).toBe("0.00");
        expect(format("0.000", 0.001)).toBe("0.001");
      });

      it("should handle very large numbers", () => {
        expect(format("#,##0", 1234567890)).toBe("1,234,567,890");
      });
    });

    describe("? placeholder (space padding)", () => {
      it("should pad integer part with leading spaces", () => {
        expect(format("??0.00", 0)).toBe("  0.00");
        expect(format("??0.00", 5)).toBe("  5.00");
        expect(format("??0.00", 42)).toBe(" 42.00");
        expect(format("??0.00", 123)).toBe("123.00");
      });

      it("should pad decimal part with trailing spaces for zero", () => {
        expect(format("0.??", 1)).toBe("1.  ");
        expect(format("0.??", 1.5)).toBe("1.5 ");
        expect(format("0.??", 1.23)).toBe("1.23");
      });

      it("should handle mixed ? and 0 in decimal part", () => {
        expect(format("0.0?", 1.5)).toBe("1.5 ");
        expect(format("0.0?", 1.53)).toBe("1.53");
        expect(format("0.?0", 1.5)).toBe("1.50");
      });

      it("should produce space-dot-space for all-? format with zero", () => {
        expect(format("???.???", 0)).toBe("   .   ");
      });
    });

    describe("# placeholder (suppress zeros)", () => {
      it("should suppress leading zero in integer part", () => {
        expect(format("#.00", 0)).toBe(".00");
        expect(format("#.00", 1.5)).toBe("1.50");
      });

      it("should strip trailing zeros in decimal part", () => {
        expect(format("#.##", 1.5)).toBe("1.5");
        expect(format("#.##", 1.23)).toBe("1.23");
        expect(format("#.##", 1)).toBe("1");
      });

      it("should produce empty string for #.## with zero", () => {
        expect(format("#.##", 0)).toBe("");
      });

      it("should handle mixed # and 0 in decimal part", () => {
        expect(format("#.0#", 0)).toBe(".0");
        expect(format("#.0#", 1.5)).toBe("1.5");
        expect(format("#.0#", 1.56)).toBe("1.56");
      });

      it("should handle 0.## (required integer, optional decimals)", () => {
        expect(format("0.##", 0)).toBe("0");
        expect(format("0.##", 1)).toBe("1");
        expect(format("0.##", 1.5)).toBe("1.5");
        expect(format("0.##", 1.23)).toBe("1.23");
      });

      it("should suppress leading zero with decimal value", () => {
        expect(format("#.##", 0.25)).toBe(".25");
        expect(format("#.##", 0.1)).toBe(".1");
      });
    });

    describe("Zero-value section in multi-section formats", () => {
      it("should use third section for zero with literal dash", () => {
        expect(format('#,##0.00;-#,##0.00;"-"??', 0)).toBe("-  ");
      });

      it("should use third section for zero with literal dash (no spaces)", () => {
        expect(format('0.00;-0.00;"-"', 0)).toBe("-");
      });

      it("should still format positive and negative correctly", () => {
        expect(format('#,##0.00;-#,##0.00;"-"??', 1234.5)).toBe("1,234.50");
        expect(format('#,##0.00;-#,##0.00;"-"??', -1234.5)).toBe("-1,234.50");
      });
    });
  });

  describe("isDateFormat", () => {
    it("should detect date formats", () => {
      expect(cellFormat.isDateFormat("yyyy-mm-dd")).toBe(true);
      expect(cellFormat.isDateFormat("m/d/yy")).toBe(true);
      expect(cellFormat.isDateFormat("dd/mm/yyyy")).toBe(true);
      expect(cellFormat.isDateFormat("h:mm:ss")).toBe(true);
    });

    it("should not detect number formats as date", () => {
      expect(cellFormat.isDateFormat("0.00")).toBe(false);
      expect(cellFormat.isDateFormat("#,##0")).toBe(false);
      expect(cellFormat.isDateFormat("0%")).toBe(false);
    });
  });

  describe("isGeneral", () => {
    it("should detect General format", () => {
      expect(cellFormat.isGeneral("General")).toBe(true);
      expect(cellFormat.isGeneral("GENERAL")).toBe(true);
      expect(cellFormat.isGeneral("general")).toBe(true);
    });

    it("should not detect other formats as General", () => {
      expect(cellFormat.isGeneral("0.00")).toBe(false);
      expect(cellFormat.isGeneral("General Text")).toBe(false);
    });
  });

  describe("getFormat", () => {
    it("should return format string for known numFmtId", () => {
      expect(cellFormat.getFormat(0)).toBe("General");
      expect(cellFormat.getFormat(1)).toBe("0");
      expect(cellFormat.getFormat(2)).toBe("0.00");
      expect(cellFormat.getFormat(9)).toBe("0%");
      expect(cellFormat.getFormat(14)).toBe("m/d/yy");
    });

    it("should return General for unknown numFmtId", () => {
      expect(cellFormat.getFormat(999)).toBe("General");
    });

    it("should handle default mapping for certain numFmtIds", () => {
      // 5-8 map to 37-40
      expect(cellFormat.getFormat(5)).toBe("#,##0 ;(#,##0)");
      // 27-31 map to 14
      expect(cellFormat.getFormat(27)).toBe("m/d/yy");
      // 59-62 map to 1-4
      expect(cellFormat.getFormat(59)).toBe("0");
    });
  });

  describe("Conditional formats", () => {
    it("should handle conditional format [>100]", () => {
      expect(format("[>100]0.00;0", 150)).toBe("150.00");
      expect(format("[>100]0.00;0", 50)).toBe("50");
    });

    it("should handle conditional format [<=50]", () => {
      expect(format("[<=50]0.00;0", 30)).toBe("30.00");
      expect(format("[<=50]0.00;0", 80)).toBe("80");
    });
  });

  describe("Placeholder characters", () => {
    it("should handle underscore _ placeholder for spacing", () => {
      expect(format("0_)", 123)).toBe("123 ");
    });

    it("should handle asterisk * placeholder", () => {
      // Asterisk fill is simplified to empty in our implementation
      expect(format("0*-", 123)).toBe("123");
    });
  });

  describe("Accounting formats", () => {
    it("should format accounting format 41", () => {
      const fmt = cellFormat.getFormat(41);
      expect(fmt).toContain("#,##0");
    });

    it("should format accounting format 44", () => {
      const fmt = cellFormat.getFormat(44);
      expect(fmt).toContain("$");
    });
  });

  describe("Fractional seconds", () => {
    it("should format ss.0", () => {
      // Test with a time that has fractional seconds
      // 0.50001157407 = 12:00:01.0 (approximately)
      const result = format("h:mm:ss.0", 0.50001157407);
      expect(result).toMatch(/\d+:\d+:\d+\.\d/);
    });

    it("should format ss.00", () => {
      const result = format("h:mm:ss.00", 0.50001157407);
      expect(result).toMatch(/\d+:\d+:\d+\.\d{2}/);
    });
  });

  describe("Locale codes", () => {
    it("should strip locale codes like [$-804]", () => {
      expect(format("[$-804]#,##0", 1234)).toBe("1,234");
    });

    it("should strip currency locale codes", () => {
      // Currency symbol with locale is stripped, only the format remains
      expect(format("[$€-407]#,##0.00", 1234.56)).toBe("1,234.56");
    });
  });

  describe("Single letter month (mmmmm)", () => {
    it("should format mmmmm as single letter", () => {
      // January 15, 2024 (Excel serial: 45306)
      const result = format("mmmmm", 45306);
      expect(result).toBe("J");
    });
  });

  describe("Negative number handling", () => {
    it("should not double negative sign with multi-section format", () => {
      // With two sections, negative should use second section without adding another minus
      expect(format("#,##0;(#,##0)", -1234)).toBe("(1,234)");
    });

    it("should show negative sign with single section format", () => {
      expect(format("#,##0", -1234)).toBe("-1,234");
    });
  });

  describe("Backslash escape", () => {
    it("should handle backslash escaped characters", () => {
      // 0\-0 means: digit + literal "-" + digit
      expect(format("0\\-0", 12)).toBe("1-2");
    });

    it("should handle phone number format", () => {
      expect(format("000-0000", 1234567)).toBe("123-4567");
    });
  });

  describe("AM/PM time format", () => {
    it("should format midnight (12 AM) as 12:00:32 AM", () => {
      // Excel serial for midnight: 0 or any integer (time portion is 0)
      // 0.000370370... = 32 seconds after midnight
      const midnightSerial = 32 / 86400; // 00:00:32
      expect(format("h:mm:ss AM/PM", midnightSerial)).toBe("12:00:32 AM");
    });

    it("should format midnight with hh as 12:00:32 AM", () => {
      const midnightSerial = 32 / 86400; // 00:00:32
      expect(format("hh:mm:ss AM/PM", midnightSerial)).toBe("12:00:32 AM");
    });

    it("should format noon (12 PM) as 12:00:00 PM", () => {
      // Excel serial for noon: 0.5 (half a day)
      // 0.5 + 32/86400 = 12:00:32 PM
      const noonSerial = 0.5 + 32 / 86400; // 12:00:32
      expect(format("h:mm:ss AM/PM", noonSerial)).toBe("12:00:32 PM");
    });

    it("should format noon with hh as 12:00:00 PM", () => {
      const noonSerial = 0.5 + 32 / 86400; // 12:00:32
      expect(format("hh:mm:ss AM/PM", noonSerial)).toBe("12:00:32 PM");
    });

    it("should format 1 AM correctly", () => {
      // 1:00:32 AM = 1 hour + 32 seconds = (3600 + 32) / 86400
      const serial = (3600 + 32) / 86400;
      expect(format("h:mm:ss AM/PM", serial)).toBe("1:00:32 AM");
    });

    it("should format 1 PM correctly", () => {
      // 1:00:32 PM = 13 hours + 32 seconds = (13 * 3600 + 32) / 86400
      const serial = (13 * 3600 + 32) / 86400;
      expect(format("h:mm:ss AM/PM", serial)).toBe("1:00:32 PM");
    });

    it("should format 11 AM correctly", () => {
      // 11:00:32 AM = 11 hours + 32 seconds
      const serial = (11 * 3600 + 32) / 86400;
      expect(format("h:mm:ss AM/PM", serial)).toBe("11:00:32 AM");
    });

    it("should format 11 PM correctly", () => {
      // 11:00:32 PM = 23 hours + 32 seconds
      const serial = (23 * 3600 + 32) / 86400;
      expect(format("h:mm:ss AM/PM", serial)).toBe("11:00:32 PM");
    });
  });
});
