import { parseValueByFormat } from "@excel/utils/cell-format-parse";
import { describe, it, expect } from "vitest";

function utcDate(y: number, m: number, d: number): Date {
  return new Date(Date.UTC(y, m - 1, d));
}

describe("cell-format-parse", () => {
  describe("parseValueByFormat — date formats", () => {
    it("parses day-first dd/mm/yyyy", () => {
      expect(parseValueByFormat("dd/mm/yyyy", "09/07/2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("parses month-first mm/dd/yyyy", () => {
      expect(parseValueByFormat("mm/dd/yyyy", "07/09/2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("parses ISO yyyy-mm-dd", () => {
      expect(parseValueByFormat("yyyy-mm-dd", "2026-07-09")).toEqual(utcDate(2026, 7, 9));
    });

    it("parses dotted day-first dd.mm.yyyy regardless of the input's own separator", () => {
      // Format uses dots, input uses dashes — token order still lines up.
      expect(parseValueByFormat("dd.mm.yyyy", "09-07-2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("parses a named-month format d-mmm-yyyy", () => {
      expect(parseValueByFormat("d-mmm-yyyy", "9-Jul-2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("parses a full month name mmmm d, yyyy", () => {
      expect(parseValueByFormat("mmmm d, yyyy", "July 9, 2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("accepts a numeric month for a named-month format (token gates display, not input)", () => {
      // d-mmm-yy DISPLAYS as 2-Jul-26, but a numeric month is still a valid
      // date on input — mirroring Excel's own manual-entry recognizer.
      expect(parseValueByFormat("d-mmm-yy", "02.07.2026")).toEqual(utcDate(2026, 7, 2));
      expect(parseValueByFormat("d-mmm-yy", "2-7-2026")).toEqual(utcDate(2026, 7, 2));
    });

    it("accepts a month name for a numeric-month format", () => {
      expect(parseValueByFormat("dd/mm/yyyy", "09/Jul/2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("does not run a full 4-digit year through the 2-digit pivot", () => {
      expect(parseValueByFormat("dd/mm/yy", "09/07/2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("parses a 2-digit year with 1900/2000 pivot", () => {
      expect(parseValueByFormat("dd/mm/yy", "09/07/26")).toEqual(utcDate(2026, 7, 9));
      expect(parseValueByFormat("dd/mm/yy", "09/07/85")).toEqual(utcDate(1985, 7, 9));
    });

    it("rejects an overflowed day for the given month", () => {
      expect(parseValueByFormat("dd/mm/yyyy", "31/02/2026")).toBeUndefined();
    });

    it("rejects input that doesn't have a part for every token", () => {
      expect(parseValueByFormat("dd/mm/yyyy", "09/07")).toBeUndefined();
      expect(parseValueByFormat("dd/mm/yyyy", "not a date")).toBeUndefined();
    });

    it("rejects an unrecognized month name", () => {
      expect(parseValueByFormat("d-mmm-yyyy", "9-Xyz-2026")).toBeUndefined();
    });
  });

  describe("parseValueByFormat — time formats", () => {
    it("parses 24h h:mm as a fraction of a day", () => {
      expect(parseValueByFormat("h:mm", "09:00")).toBeCloseTo(9 / 24);
      expect(parseValueByFormat("h:mm", "23:30")).toBeCloseTo(23.5 / 24);
    });

    it("parses h:mm:ss", () => {
      expect(parseValueByFormat("h:mm:ss", "01:02:03")).toBeCloseTo(
        (1 * 3600 + 2 * 60 + 3) / 86400
      );
    });

    it("parses 12h AM/PM format", () => {
      expect(parseValueByFormat("h:mm AM/PM", "9:00 AM")).toBeCloseTo(9 / 24);
      expect(parseValueByFormat("h:mm AM/PM", "9:00 PM")).toBeCloseTo(21 / 24);
      expect(parseValueByFormat("h:mm AM/PM", "12:00 AM")).toBeCloseTo(0);
      expect(parseValueByFormat("h:mm AM/PM", "12:00 PM")).toBeCloseTo(12 / 24);
    });

    it("rejects out-of-range minutes", () => {
      expect(parseValueByFormat("h:mm", "09:75")).toBeUndefined();
    });

    it("defaults omitted trailing seconds to 0 for an h:mm:ss cell, matching Excel's own manual-entry behavior", () => {
      expect(parseValueByFormat("h:mm:ss", "09:00")).toBeCloseTo(9 / 24);
      expect(parseValueByFormat("hh:mm:ss", "23:30")).toBeCloseTo(23.5 / 24);
    });

    it("defaults an omitted trailing AM/PM to 24h interpretation", () => {
      expect(parseValueByFormat("h:mm AM/PM", "14:00")).toBeCloseTo(14 / 24);
    });

    it("still rejects input with more parts than the format has tokens", () => {
      expect(parseValueByFormat("h:mm", "09:00:00")).toBeUndefined();
    });

    it("rejects 24-hour values when AM/PM requires a 1-12 hour", () => {
      expect(parseValueByFormat("h:mm AM/PM", "13:00 PM")).toBeUndefined();
      expect(parseValueByFormat("h:mm AM/PM", "0:30 AM")).toBeUndefined();
    });

    it("supports the short A/P marker", () => {
      expect(parseValueByFormat("h:mm A/P", "9:00 P")).toBeCloseTo(21 / 24);
    });

    it("does not turn empty input into midnight", () => {
      expect(parseValueByFormat("h:mm", "")).toBeUndefined();
      expect(parseValueByFormat("h:mm", "   ")).toBeUndefined();
    });

    it("parses elapsed-time formats beyond 24 hours", () => {
      expect(parseValueByFormat("[h]:mm:ss", "25:30:00")).toBeCloseTo(25.5 / 24);
      expect(parseValueByFormat("[m]:ss", "90:30")).toBeCloseTo(90.5 / 1440);
    });
  });

  describe("parseValueByFormat — non-date/time formats", () => {
    it("returns undefined for General/plain-number formats", () => {
      expect(parseValueByFormat("General", "09.07.2026")).toBeUndefined();
      expect(parseValueByFormat("0.00", "09.07.2026")).toBeUndefined();
      expect(parseValueByFormat("@", "09.07.2026")).toBeUndefined();
    });
  });

  describe("parseValueByFormat — combined datetime", () => {
    it("parses a full datetime format when the input carries both parts", () => {
      expect(parseValueByFormat("yyyy-mm-dd hh:mm:ss", "2026-07-09 09:30:00")).toEqual(
        new Date(Date.UTC(2026, 6, 9, 9, 30, 0))
      );
    });

    it("ignores quoted and escaped word literals in display text", () => {
      expect(parseValueByFormat('yyyy-mm-dd "at" hh:mm', "2026-07-09 at 09:30")).toEqual(
        new Date(Date.UTC(2026, 6, 9, 9, 30))
      );
      expect(parseValueByFormat("yyyy-mm-dd\\T hh:mm", "2026-07-09T 09:30")).toEqual(
        new Date(Date.UTC(2026, 6, 9, 9, 30))
      );
    });

    it("does not split a format section at a quoted semicolon", () => {
      expect(parseValueByFormat('yyyy";"mm";"dd', "2026;07;09")).toEqual(
        new Date(Date.UTC(2026, 6, 9))
      );
    });
  });

  // These exercise the anchored-literal matcher rather than any per-format
  // special case: the format's separators/literals are match anchors, so new
  // literal shapes need no new code path.
  describe("parseValueByFormat — anchored-literal matching", () => {
    it("tolerates a different separator in the input than the format", () => {
      expect(parseValueByFormat("dd-mm-yyyy", "09.07.2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("consumes an embedded word literal instead of miscounting it as a value", () => {
      expect(parseValueByFormat('dd/mm/yyyy" at "hh:mm', "09/07/2026 at 14:30")).toEqual(
        new Date(Date.UTC(2026, 6, 9, 14, 30))
      );
    });

    it("skips a display-only locale/currency bracket", () => {
      expect(parseValueByFormat("[$-409]d/m/yyyy", "9/7/2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("parses an elapsed [mm]:ss duration", () => {
      expect(parseValueByFormat("[mm]:ss", "90:30")).toBeCloseTo(90.5 / 1440);
    });

    it("parses a full date + 12h time with AM/PM and rejects an out-of-range hour", () => {
      expect(parseValueByFormat("m/d/yyyy h:mm AM/PM", "7/9/2026 2:05 PM")).toEqual(
        new Date(Date.UTC(2026, 6, 9, 14, 5))
      );
      expect(parseValueByFormat("m/d/yyyy h:mm AM/PM", "7/9/2026 13:05 PM")).toBeUndefined();
    });

    it("rejects input with unconsumed trailing content", () => {
      expect(parseValueByFormat("dd/mm/yyyy", "09/07/2026 garbage")).toBeUndefined();
    });

    it("parses a full month name", () => {
      expect(parseValueByFormat("d mmmm yyyy", "9 July 2026")).toEqual(utcDate(2026, 7, 9));
    });

    it("treats format tokens case-insensitively (YYYY, am/pm, A/P)", () => {
      expect(parseValueByFormat("YYYY-MM-DD", "2026-07-09")).toEqual(utcDate(2026, 7, 9));
      expect(parseValueByFormat("h:mm am/pm", "9:00 pm")).toBeCloseTo(21 / 24);
      expect(parseValueByFormat("h:mm A/P", "9:00 p")).toBeCloseTo(21 / 24);
    });
  });
});
