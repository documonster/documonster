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

    it("accepts a numeric month for a named-month format (d-mmm-yy displays as 2-Jul-26, but a numeric month is still a valid date)", () => {
      // The format's month TOKEN (mmm vs mm) only controls how the value
      // later displays - it must not gate whether a numeric month is
      // accepted, mirroring what Excel itself does on manual entry.
      expect(parseValueByFormat("d-mmm-yy", "02.07.2026")).toEqual(utcDate(2026, 7, 2));
      expect(parseValueByFormat("d-mmm-yy", "2-7-2026")).toEqual(utcDate(2026, 7, 2));
    });

    it("accepts a month name for a numeric-month format", () => {
      expect(parseValueByFormat("dd/mm/yyyy", "09/Jul/2026")).toEqual(utcDate(2026, 7, 9));
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
  });
});
