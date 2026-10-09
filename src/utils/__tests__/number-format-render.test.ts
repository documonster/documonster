import { renderNumberFormat as render } from "@utils/number-format-render";
import { describe, expect, it } from "vitest";

/** 2024-01-15, a Monday, under the 1900 epoch. */
const JAN_15_2024 = 45306;

describe("renderNumberFormat", () => {
  describe("literal units", () => {
    it.each([
      ["\\$0.0,,\\M", 5e9, "$5000.0M"],
      ['$0.0,,"M"', 5e9, "$5000.0M"],
      ["0.0,,\\M", 999999999, "1000.0M"],
      ['[>=1000]0.0,"K";0', 12345, "12.3K"],
      ["0.0\\s", 12, "12.0s"],
      ["0.0\\h", 12, "12.0h"],
      ['0" days"', 12, "12 days"],
      ["0 \\d", 12, "12 d"],
      ["0_m", 12, "12 "],
      ["*-0", 5, "5"]
    ])("%s renders %s as %s", (fmt, value, expected) => {
      expect(render(fmt, value)).toBe(expected);
    });

    it("puts the minus sign in front of a currency symbol", () => {
      expect(render("\\$0.0,,\\M", -5e9)).toBe("-$5000.0M");
      expect(render("$#,##0", -1234)).toBe("-$1,234");
      expect(render("[$$-409]#,##0.00", -1234.5)).toBe("-$1,234.50");
    });

    it("shows a currency locale's symbol and drops the locale id", () => {
      expect(render("[$€-407]#,##0.00", 1234.56)).toBe("€1,234.56");
      expect(render("[$-804]#,##0", 1234)).toBe("1,234");
    });
  });

  describe("rounding", () => {
    it("rounds half away from zero at fifteen significant digits, as Excel does", () => {
      expect(render("0.00", 1.005)).toBe("1.01");
      expect(render("0.00", 2.675)).toBe("2.68");
      expect(render("0", 0.5)).toBe("1");
      expect(render("0", -0.5)).toBe("-1");
    });

    it("drops the sign of a value that rounds to zero", () => {
      expect(render("0.00", -0.001)).toBe("0.00");
      expect(render("#,###", 0.4)).toBe("");
    });

    it("keeps the negative section's own sign even when the value rounds to zero", () => {
      // Excel's own `-0`: the section is chosen by the value, before rounding.
      expect(render("#,##0;-#,##0", -0.4)).toBe("-0");
      expect(render("0.00;(0.00)", -0.001)).toBe("(0.00)");
    });

    it("writes large integers out in full", () => {
      expect(render("0", 1e21)).toBe("1000000000000000000000");
    });
  });

  describe("digit placeholders", () => {
    it("keeps the decimal point when every digit after it is suppressed", () => {
      expect(render("#.##", 12)).toBe("12.");
      expect(render("#.##", 0)).toBe(".");
    });

    it("pads with spaces under ?", () => {
      expect(render("??0.0?", 1.5)).toBe("  1.5 ");
    });

    it("groups across overflow digits", () => {
      expect(render("#,##0", 1234567890)).toBe("1,234,567,890");
    });

    it("spreads digits across literals", () => {
      expect(render("0000-000", 1234567)).toBe("1234-567");
      expect(render("(###) ###-####", 5551234)).toBe("() 555-1234");
    });

    it("scales by a thousand per trailing comma, after decimals too", () => {
      expect(render("#,##0.00,,", 1234567890)).toBe("1,234.57");
      // A comma ending the integer placeholders scales too, even with decimals after it.
      expect(render("0,.0", 1234)).toBe("1.2");
      expect(render("0,,.00", 1234567)).toBe("1.23");
      expect(render("#,##0,.0", 1234567)).toBe("1,234.6");
    });

    it("multiplies by a hundred per percent sign and keeps surrounding literals", () => {
      expect(render("0.0%", 0.1234)).toBe("12.3%");
      expect(render('0%" off"', 0.25)).toBe("25% off");
      expect(render("0%%", 0.01)).toBe("100%%");
    });
  });

  describe("scientific", () => {
    it.each([
      ["0.00E+00", 0, "0.00E+00"],
      ["0.00E+00", 9.999, "1.00E+01"],
      ["0.00E-00", 0.001234, "1.23E-03"],
      ["0.00E-00", 1234, "1.23E03"],
      ["0.0e+0", 1234, "1.2e+3"],
      ["##0.0E+0", 12345, "12.3E+3"],
      ["0.000E+000", 5e7, "5.000E+007"],
      // The mantissa is found by moving a decimal point, not by dividing a double that underflows.
      ["0.00E+00", Number.MIN_VALUE, "4.94E-324"],
      ["##0.0E+0", 1e-320, "10.0E-321"],
      ["0.00E+00", Number.MAX_VALUE, "1.80E+308"]
    ])("%s renders %s as %s", (fmt, value, expected) => {
      expect(render(fmt, value)).toBe(expected);
    });
  });

  describe("fractions", () => {
    it.each([
      ["# ?/?", 3.25, "3 1/4"],
      // Excel shows the last continued-fraction convergent, not the closest fraction: 1/7, not 14/99.
      ["# ??/??", 3.14159, "3  1/7 "],
      ["# ???/???", 3.14159, "3  16/113"],
      ["# ?/?", 0.3, " 2/7"],
      // A blank whole number keeps its separator, so `?` still lines fractions up.
      ["# ?/?", 0.5, " 1/2"],
      ["# ??/??", 0.333, "  1/3 "],
      ["# ?/?", 5, "5    "],
      ["# ?/?", 0.97, "1    "],
      ["# ?/?", 0, "0    "],
      ["# ?/?", -3.25, "-3 1/4"],
      ["?/?", 1.5, "3/2"],
      ["# ?/8", 1.5, "1 4/8"],
      ["# ?/16", 0.3, " 5/16"],
      // The first denominator placeholder decides the padding, and padding never changes the value:
      // zeros go in front (1/02), spaces behind (1/2 ), and `#` adds nothing.
      ["?/00", 0.5, "1/02"],
      ["?/000", 0.5, "1/002"],
      ["?/0#", 0.5, "1/02"],
      ["?/?0", 0.5, "1/2 "],
      ["?/#0", 0.5, "1/2"],
      ["00/??", 0.5, "01/2 "],
      ["# ??/##", 0.5, "  1/2"],
      // A quoted or escaped digit is text, not a fixed denominator.
      ['0/"8"', 12, "12/8"],
      ["0/\\8", 12, "12/8"],
      ["# ?/10", 1.25, "1 3/10"],
      // The whole-number part of a fraction still groups thousands.
      ["#,##0 ?/?", 1234.5, "1,234 1/2"],
      ["#,##0 ?/?", 1234, "1,234    "]
    ])("%s renders %s as %j", (fmt, value, expected) => {
      expect(render(fmt, value)).toBe(expected);
    });
  });

  describe("dates and times", () => {
    it.each([
      ["yyyy-mm-dd hh:mm:ss", JAN_15_2024 + 0.5208333, "2024-01-15 12:30:00"],
      ['h "at" mm', 0.75, "18 at 00"],
      // `mm` after an hour is a minute even with a day between them, as Excel reads it.
      ["hh dd mm", JAN_15_2024, "00 15 00"],
      ["dd mm", JAN_15_2024, "15 01"],
      ["h:mm:ss mm", JAN_15_2024 + 0.5, "12:00:00 01"],
      ["yyyy-mm-dd", 0, "1900-01-00"],
      ["dddd", 0, "Saturday"],
      ["mmmmm yyy", JAN_15_2024, "J 2024"],
      ["ddd, mmm d", JAN_15_2024, "Mon, Jan 15"],
      ["dddd", JAN_15_2024, "Monday"],
      ["bbbb", JAN_15_2024, "2567"],
      ["yyyy\\-mm", JAN_15_2024, "2024-01"],
      ['"Day "d', JAN_15_2024, "Day 15"],
      ["dd/mm/yyyy", 60, "29/02/1900"],
      // Weekdays follow the serial, so Excel's phantom 1900-02-29 sits between Tuesday and Thursday.
      ["dddd", 1, "Sunday"],
      ["dddd", 59, "Tuesday"],
      ["dddd", 60, "Wednesday"],
      ["dddd", 61, "Thursday"],
      ["h:mm am/pm", 0.6, "2:24 pm"],
      ["h A/P", 0.2, "4 A"],
      ["mm:ss.00", 0.0001157, "00:10.00"]
    ])("%s renders %s as %s", (fmt, value, expected) => {
      expect(render(fmt, value)).toBe(expected);
    });

    it("rounds to the second even when no time is shown, as Excel does", () => {
      expect(render("yyyy-mm-dd", JAN_15_2024 + 0.999999)).toBe("2024-01-16");
      expect(render("yyyy-mm-dd", JAN_15_2024 + 0.9999)).toBe("2024-01-15");
    });

    it("keeps a far date displayable at three second decimals", () => {
      expect(render("yyyy-mm-dd hh:mm:ss.000", 200000)).toBe("2447-07-30 00:00:00.000");
      expect(render("yyyy-mm-dd hh:mm:ss.000", 2958465.5)).toBe("9999-12-31 12:00:00.000");
    });

    it("rejects more than three second decimals, as Excel does", () => {
      expect(render("ss.0000", 0.0001)).toBe("########");
      expect(render("ss.000", 0.0001)).toBe("08.640");
    });

    it("computes the second decimals a format asks for", () => {
      expect(render("ss.000", 1.234 / 86400)).toBe("01.234");
      expect(render("h:mm:ss.00", (16 * 3600 + 36 * 60 + 3.75) / 86400)).toBe("16:36:03.75");
      expect(render("[ss].00", 3735.8 / 86400)).toBe("3735.80");
    });

    it("shows hashes for a negative date or time, but a sign for a lone elapsed count", () => {
      expect(render("yyyy-mm-dd", -1)).toBe("########");
      expect(render("[h]:mm", -1.5)).toBe("########");
      expect(render("0;[h]:mm", -1)).toBe("########");
      expect(render("[h]", -1)).toBe("-24");
      expect(render("[ss]", -1.5)).toBe("-129600");
    });

    it("bounds dates at 9999-12-31 in the workbook's own epoch, after rounding", () => {
      expect(render("yyyy-mm-dd", 2958465)).toBe("9999-12-31");
      expect(render("yyyy-mm-dd", 2958465.99999)).toBe("9999-12-31");
      expect(render("yyyy-mm-dd hh:mm:ss", 2958465.99999999)).toBe("########");
      expect(render("yyyy-mm-dd", 2957003, { date1904: true })).toBe("9999-12-31");
      expect(render("yyyy-mm-dd", 2957004, { date1904: true })).toBe("########");
      expect(render("[h]:mm", 3e6)).toBe("########");
      // A lone elapsed count is a plain number, without the date bound.
      expect(render("[ss]", 3e6)).toBe("259200000000");
    });

    it("rounds a fraction of a second into the seconds, then drops seconds the format hides", () => {
      // As SheetJS (calibrated against Excel) shows: 10:59:59 is 10:59, but 10:59:59.6 is 11:00.
      expect(render("h:mm", (10 * 3600 + 59 * 60 + 59) / 86400)).toBe("10:59");
      expect(render("h:mm", (10 * 3600 + 59 * 60 + 59.6) / 86400)).toBe("11:00");
    });

    it("rounds to the shown precision before splitting fields", () => {
      // 23:59:59.99999 shows as the next midnight, with the date carried.
      expect(render("yyyy-mm-dd hh:mm:ss", JAN_15_2024 + 0.9999999)).toBe("2024-01-16 00:00:00");
      expect(render("ss.0", 0.9999999)).toBe("00.0");
    });

    it("counts elapsed time in total units", () => {
      expect(render("[h]:mm:ss", 1.5)).toBe("36:00:00");
      expect(render("[mm]:ss", 0.05)).toBe("72:00");
      // After a calendar field, an elapsed count is the whole hours of that day.
      expect(render("dd [h]", 1.5)).toBe("01 12");
      expect(render("[h] dd", 1.5)).toBe("36 01");
    });

    it("honours the 1904 epoch", () => {
      expect(render("yyyy-mm-dd", 1, { date1904: true })).toBe("1904-01-02");
      expect(render("yyyy-mm-dd", 1)).toBe("1900-01-01");
      expect(render("dddd", 0, { date1904: true })).toBe("Friday");
    });
  });

  describe("sections", () => {
    it("shows an empty section as nothing", () => {
      expect(render("0;-0;;@", 0)).toBe("");
      expect(render(";;;", 5)).toBe("");
    });

    it("chooses by condition", () => {
      expect(render("[>100]0.00;0", 150)).toBe("150.00");
      expect(render("[>100]0.00;0", 50)).toBe("50");
    });

    it("hides text under an empty text section", () => {
      expect(render("0;0;0;", "hello")).toBe("");
      expect(render(";;;", "hello")).toBe("");
    });

    it("formats text through the text section, or shows it as is", () => {
      expect(render("0;@", "hi")).toBe("hi");
      expect(render('@" units"', "5")).toBe("5 units");
      expect(render('0;0;0;"t: "@', "x")).toBe("t: x");
      expect(render("0", "hi")).toBe("hi");
    });

    it("shows the sign of a section without digits, as Excel does", () => {
      expect(render('"x"', -5)).toBe("-x");
      expect(render('"x";"y"', -5)).toBe("y");
      expect(render('"neg "0', -5)).toBe("-neg 5");
    });

    it("pads with the width of the character after _, even a separator", () => {
      expect(render("0_;0", 12)).toBe("1 2");
      expect(render('0_";0', 12)).toBe("12 ");
    });

    it("shows a number under a text-only format as General", () => {
      expect(render('"n="@', 5)).toBe("5");
      expect(render("@", 1.5)).toBe("1.5");
    });

    it("drops the sign where Excel does in conditional formats", () => {
      // A condition admitting only negatives drops it; one admitting zero or positives keeps it.
      expect(render("[<0]0;0", -5)).toBe("5");
      expect(render('[<0]"neg";0', -5)).toBe("neg");
      expect(render("[<-5]0;0", -7)).toBe("7");
      expect(render("[<=0]0;0", -3)).toBe("-3");
      expect(render("[<100]0;0", -5)).toBe("-5");
      // A two-section "else" keeps it unless the first condition covered every positive or only negatives.
      expect(render("[>100]0;0", -5)).toBe("-5");
      expect(render("[>0]0;0", -5)).toBe("5");
      expect(render("[<-10]0;0", -5)).toBe("5");
      // With three sections, a bare second section means [<0].
      expect(render("[>100]0;0;0", -5)).toBe("5");
      expect(render('[>100]0;[<0]"m"0;0', -5)).toBe("m5");
    });

    it("shows hashes when no conditional section matches", () => {
      expect(render('[>100]"A"0;[<-10]"B"0', -5)).toBe("########");
    });
  });

  describe("values a format cannot show", () => {
    it("shows hashes for a date past 9999 or a number past double range", () => {
      expect(render("yyyy-mm-dd", 3e6)).toBe("########");
      expect(render("[h]:mm", 1e300)).toBe("########");
      expect(render("0" + "%".repeat(200), 1)).toBe("########");
      // Zero stays zero under any scaling.
      expect(render("0" + "%".repeat(200), 0)).toBe("0" + "%".repeat(200));
    });

    it("finds the closest fraction under the denominator bound", () => {
      expect(render("?/???", Math.PI)).toBe("355/113");
      // A small value is exact when the bound allows its denominator.
      expect(render("?/??????????????", 1e-13)).toBe("1/10000000000000");
    });
  });

  describe("General", () => {
    it.each([
      ["General", 1.5, "1.5"],
      // At most eleven characters, as Excel shows it.
      ["General", 1 / 3, "0.333333333"],
      ["General", 123456789012, "1.23457E+11"],
      ["General", 99999999999, "99999999999"],
      ["General", 1234567890.5, "1234567891"],
      ["General", 1e21, "1E+21"],
      ["General", 1e-12, "1E-12"],
      ["[Red]General", -1.5, "-1.5"],
      ['General" kg"', 3, "3 kg"],
      ["", 42, "42"]
    ])("%j renders %s as %s", (fmt, value, expected) => {
      expect(render(fmt, value)).toBe(expected);
    });

    it("shows booleans as TRUE and FALSE", () => {
      expect(render("0.00", true)).toBe("TRUE");
      expect(render("General", false)).toBe("FALSE");
    });
  });
});
