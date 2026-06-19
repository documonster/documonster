import {
  getSupportedFormats,
  createDateParser,
  createIsoDateParser,
  createDateFormatter,
  createIsoDateFormatter
} from "@utils/datetime";
import { describe, it, expect } from "vitest";

describe("datetime", () => {
  describe("DateParser", () => {
    it("creates a reusable parser with specified formats", () => {
      const parser = createDateParser(["YYYY-MM-DD", "MM-DD-YYYY"]);
      expect(parser.parse("2024-12-26")).toBeInstanceOf(Date);
      expect(parser.parse("12-26-2024")).toBeInstanceOf(Date);
      expect(parser.parse("26-12-2024")).toBeNull();
    });

    it("creates ISO auto-detect parser", () => {
      const parser = createIsoDateParser();
      expect(parser.parse("2024-12-26")).toBeInstanceOf(Date);
      expect(parser.parse("2024-12-26T10:30:00Z")).toBeInstanceOf(Date);
      expect(parser.parse("2024-12-26 10:30:00")).toBeInstanceOf(Date);
    });

    it("parses all values in batch", () => {
      const parser = createDateParser(["YYYY-MM-DD"]);
      const results = parser.parseAll(["2024-01-01", "invalid", "2024-12-31"]);
      expect(results[0]).toBeInstanceOf(Date);
      expect(results[1]).toBeNull();
      expect(results[2]).toBeInstanceOf(Date);
    });

    it("filters valid dates", () => {
      const parser = createDateParser(["YYYY-MM-DD"]);
      const results = parser.parseValid(["2024-01-01", "invalid", "2024-12-31"]);
      expect(results.length).toBe(2);
      expect(results.every(d => d instanceof Date)).toBe(true);
    });
  });

  describe("DateFormatter", () => {
    it("creates ISO formatter", () => {
      const formatter = createIsoDateFormatter(true);
      const date = new Date(Date.UTC(2024, 11, 26, 10, 30, 45));
      expect(formatter.format(date)).toBe("2024-12-26T10:30:45.000Z");
    });

    it("creates custom format formatter", () => {
      const formatter = createDateFormatter("YYYY-MM-DD", { utc: true });
      const date = new Date(Date.UTC(2024, 11, 26));
      expect(formatter.format(date)).toBe("2024-12-26");
    });

    it("formats all dates in batch", () => {
      const formatter = createDateFormatter("MM/DD/YYYY", { utc: true });
      const dates = [
        new Date(Date.UTC(2024, 0, 1)),
        new Date(Date.UTC(2024, 5, 15)),
        new Date(Date.UTC(2024, 11, 31))
      ];
      const results = formatter.formatAll(dates);
      expect(results).toEqual(["01/01/2024", "06/15/2024", "12/31/2024"]);
    });

    it("handles invalid date in formatter", () => {
      const formatter = createDateFormatter("YYYY-MM-DD", { utc: true });
      expect(formatter.format(new Date("invalid"))).toBe("");
    });
  });

  describe("getSupportedFormats", () => {
    it("returns array of supported format strings", () => {
      const formats = getSupportedFormats();
      expect(Array.isArray(formats)).toBe(true);
      expect(formats.length).toBeGreaterThan(0);
      expect(formats).toContain("YYYY-MM-DD");
      expect(formats).toContain("YYYY-MM-DD[T]HH:mm:ssZ");
    });
  });

  describe("performance", () => {
    it("handles large batches efficiently", () => {
      const parser = createDateParser(["YYYY-MM-DD"]);
      const formatter = createDateFormatter("YYYY-MM-DD", { utc: true });

      const dateStrings: string[] = [];
      for (let i = 0; i < 10000; i++) {
        const month = String((i % 12) + 1).padStart(2, "0");
        const day = String((i % 28) + 1).padStart(2, "0");
        dateStrings.push(`2024-${month}-${day}`);
      }

      const start = performance.now();
      const dates = parser.parseAll(dateStrings);
      const formatted = formatter.formatAll(dates.filter((d): d is Date => d !== null));
      const elapsed = performance.now() - start;

      expect(dates.length).toBe(10000);
      expect(formatted.length).toBe(10000);
      // Avoid hard micro-bench thresholds in unit tests (CI and developer machines vary).
      // Still guard against pathological regressions.
      expect(elapsed).toBeGreaterThanOrEqual(0);
      expect(elapsed).toBeLessThan(5000);
    });
  });
});
