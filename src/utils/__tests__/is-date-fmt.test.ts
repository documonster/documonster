import { isDateFmt, splitFormatSections } from "@utils/utils.base";
import { describe, it, expect } from "vitest";

// =============================================================================
// splitFormatSections
// =============================================================================
describe("splitFormatSections", () => {
  it("returns a single-element array for a simple format", () => {
    expect(splitFormatSections("General")).toEqual(["General"]);
  });

  it("splits plain sections by semicolons", () => {
    expect(splitFormatSections("#,##0;-#,##0;0;@")).toEqual(["#,##0", "-#,##0", "0", "@"]);
  });

  it("does not split on semicolons inside double quotes", () => {
    expect(splitFormatSections('"a;b"yyyy')).toEqual(['"a;b"yyyy']);
  });

  it("does not split on semicolons inside brackets", () => {
    expect(splitFormatSections("[$;-409]yyyy")).toEqual(["[$;-409]yyyy"]);
  });

  it("handles mixed quoted, bracketed and bare semicolons", () => {
    // first section has a quoted semicolon, second section is plain
    expect(splitFormatSections('"x;y"mm;@')).toEqual(['"x;y"mm', "@"]);
  });

  it("returns [''] for an empty string", () => {
    expect(splitFormatSections("")).toEqual([""]);
  });

  it("handles consecutive semicolons", () => {
    expect(splitFormatSections(";;")).toEqual(["", "", ""]);
  });
});

// =============================================================================
// isDateFmt
// =============================================================================
describe("isDateFmt", () => {
  describe("falsy / empty inputs", () => {
    it("returns false for null", () => {
      expect(isDateFmt(null)).toBe(false);
    });

    it("returns false for undefined", () => {
      expect(isDateFmt(undefined)).toBe(false);
    });

    it("returns false for empty string", () => {
      expect(isDateFmt("")).toBe(false);
    });
  });

  describe("standard date formats", () => {
    it.each([
      "yyyy-mm-dd",
      "mm/dd/yyyy",
      "dd/mm/yyyy",
      "dd-mmm-yy",
      "dd-mmm-yyyy",
      "mmm-yy",
      "yyyy",
      "mm",
      "dd",
      "d-mmm",
      "m/d/yy"
    ])("returns true for %s", fmt => {
      expect(isDateFmt(fmt)).toBe(true);
    });
  });

  describe("standard time formats", () => {
    it.each(["h:mm:ss", "hh:mm", "h:mm AM/PM", "hh:mm:ss.00", "mm:ss", "mm:ss.0"])(
      "returns true for %s",
      fmt => {
        expect(isDateFmt(fmt)).toBe(true);
      }
    );
  });

  describe("date-time combination formats", () => {
    it.each(["yyyy-mm-dd hh:mm:ss", "m/d/yy h:mm", "dd/mm/yyyy hh:mm:ss"])(
      "returns true for %s",
      fmt => {
        expect(isDateFmt(fmt)).toBe(true);
      }
    );
  });

  describe("formats with brackets (locale / color tags)", () => {
    it.each(["[Red]yyyy-mm-dd", "[DBNum1]yyyy-mm-dd", "[$-409]mm/dd/yyyy", "[h]:mm:ss", "[mm]:ss"])(
      "returns true for %s",
      fmt => {
        expect(isDateFmt(fmt)).toBe(true);
      }
    );
  });

  describe("formats with quoted literal text", () => {
    it('returns true for yyyy"年"mm"月"dd"日"', () => {
      expect(isDateFmt('yyyy"年"mm"月"dd"日"')).toBe(true);
    });

    it('returns false for "Date:" # (date chars only in quotes)', () => {
      expect(isDateFmt('"Date:" #')).toBe(false);
    });

    it('returns false for "yyyy" (all date chars inside quotes)', () => {
      expect(isDateFmt('"yyyy"')).toBe(false);
    });
  });

  describe("non-date number formats", () => {
    it.each([
      "#,##0",
      "#,##0.00",
      "0%",
      "0.00%",
      "0.00E+00",
      "$#,##0",
      "0.0",
      "#,##0;-#,##0",
      "#,##0.00;[Red]-#,##0.00"
    ])("returns false for %s", fmt => {
      expect(isDateFmt(fmt)).toBe(false);
    });
  });

  describe("pure text format", () => {
    it("returns false for @", () => {
      expect(isDateFmt("@")).toBe(false);
    });
  });

  // =========================================================================
  // The core bug fix: multi-section formats with text fallback
  // =========================================================================
  describe("multi-section formats with text fallback", () => {
    it('returns true for "mm/dd/yyyy;@" (date + text fallback)', () => {
      expect(isDateFmt("mm/dd/yyyy;@")).toBe(true);
    });

    it('returns true for "yyyy-mm-dd;@" (date + text fallback)', () => {
      expect(isDateFmt("yyyy-mm-dd;@")).toBe(true);
    });

    it('returns true for "dd/mm/yyyy;@;@;@" (date + 3 text fallback sections)', () => {
      expect(isDateFmt("dd/mm/yyyy;@;@;@")).toBe(true);
    });

    it('returns true for "h:mm:ss;@" (time + text fallback)', () => {
      expect(isDateFmt("h:mm:ss;@")).toBe(true);
    });

    it('returns true for "yyyy-mm-dd hh:mm;@" (datetime + text fallback)', () => {
      expect(isDateFmt("yyyy-mm-dd hh:mm;@")).toBe(true);
    });

    it('returns true for "[$-409]mm/dd/yyyy;@" (locale + date + text fallback)', () => {
      expect(isDateFmt("[$-409]mm/dd/yyyy;@")).toBe(true);
    });
  });

  describe("multi-section: first section is text, later sections have date chars", () => {
    it('returns false for "@;yyyy-mm-dd" (text first section)', () => {
      expect(isDateFmt("@;yyyy-mm-dd")).toBe(false);
    });

    it('returns false for "@;@;yyyy" (text first section)', () => {
      expect(isDateFmt("@;@;yyyy")).toBe(false);
    });
  });

  describe("multi-section number formats without date chars", () => {
    it('returns false for "#,##0;-#,##0;0;@"', () => {
      expect(isDateFmt("#,##0;-#,##0;0;@")).toBe(false);
    });

    it('returns false for "0.00;-0.00;0;@"', () => {
      expect(isDateFmt("0.00;-0.00;0;@")).toBe(false);
    });
  });

  // =========================================================================
  // Edge cases: semicolons inside quotes / brackets (safe splitting)
  // =========================================================================
  describe("semicolons inside quotes or brackets", () => {
    it("handles semicolons inside quoted text in the first section", () => {
      // "a;b"yyyy  — the semicolon is quoted, so the whole string is one section
      expect(isDateFmt('"a;b"yyyy')).toBe(true);
    });

    it("handles semicolons inside bracket expressions", () => {
      // [$;-409]yyyy — the semicolon is inside brackets (exotic currency symbol)
      expect(isDateFmt("[$;-409]yyyy")).toBe(true);
    });

    it("does not misinterpret a quoted semicolon as a section separator", () => {
      // ";" is quoted literal text, followed by number format — not a date
      expect(isDateFmt('";" #,##0')).toBe(false);
    });
  });

  describe("@ inside quotes or brackets (literal, not text placeholder)", () => {
    it('returns true for "@"yyyy (@ is a quoted literal, not a text placeholder)', () => {
      expect(isDateFmt('"@"yyyy')).toBe(true);
    });

    it("returns true for [@]yyyy (@ inside brackets)", () => {
      expect(isDateFmt("[@]yyyy")).toBe(true);
    });
  });

  describe("Excel built-in date format IDs (as format strings)", () => {
    // These correspond to built-in numFmtId values in Excel
    it.each([
      "m/d/yy", // 14
      "d-mmm-yy", // 15
      "d-mmm", // 16
      "mmm-yy", // 17
      "h:mm AM/PM", // 18
      "h:mm:ss AM/PM", // 19
      "h:mm", // 20
      "h:mm:ss", // 21
      "m/d/yy h:mm" // 22
    ])("returns true for built-in format %s", fmt => {
      expect(isDateFmt(fmt)).toBe(true);
    });
  });
});
