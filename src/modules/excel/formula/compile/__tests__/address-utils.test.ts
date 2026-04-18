/**
 * Unit tests for the address-parsing helpers. These pure functions
 * underpin defined-name resolution, CSE target-range decoding, and
 * spill-region book-keeping — incorrect parsing here cascades into
 * silent wrong-answer bugs far from the source.
 */

import { describe, expect, it } from "vitest";

import {
  colLetterToNumber,
  parseDefinedNameRange,
  parseRefRange,
  parseSimpleAddress
} from "../address-utils";

// ---------------------------------------------------------------------------
// colLetterToNumber
// ---------------------------------------------------------------------------

describe("colLetterToNumber", () => {
  it("single letters A..Z → 1..26", () => {
    expect(colLetterToNumber("A")).toBe(1);
    expect(colLetterToNumber("B")).toBe(2);
    expect(colLetterToNumber("Z")).toBe(26);
  });

  it("double letters AA..ZZ", () => {
    expect(colLetterToNumber("AA")).toBe(27);
    expect(colLetterToNumber("AZ")).toBe(52);
    expect(colLetterToNumber("BA")).toBe(53);
    expect(colLetterToNumber("ZZ")).toBe(702);
  });

  it("triple letters up to XFD", () => {
    expect(colLetterToNumber("AAA")).toBe(703);
    // Excel max column
    expect(colLetterToNumber("XFD")).toBe(16384);
  });

  it("round-trips via letter grammar (sanity)", () => {
    // Arbitrary column letters must be monotonic.
    expect(colLetterToNumber("AB")).toBeGreaterThan(colLetterToNumber("AA"));
    expect(colLetterToNumber("BA")).toBeGreaterThan(colLetterToNumber("AZ"));
  });
});

// ---------------------------------------------------------------------------
// parseSimpleAddress
// ---------------------------------------------------------------------------

describe("parseSimpleAddress", () => {
  it("parses A1", () => {
    expect(parseSimpleAddress("A1")).toEqual({ row: 1, col: 1 });
  });

  it("parses $A$1 (absolute)", () => {
    expect(parseSimpleAddress("$A$1")).toEqual({ row: 1, col: 1 });
  });

  it("parses mixed absolute", () => {
    expect(parseSimpleAddress("$A1")).toEqual({ row: 1, col: 1 });
    expect(parseSimpleAddress("A$1")).toEqual({ row: 1, col: 1 });
  });

  it("parses multi-letter columns", () => {
    expect(parseSimpleAddress("AA99")).toEqual({ row: 99, col: 27 });
    expect(parseSimpleAddress("XFD1048576")).toEqual({ row: 1048576, col: 16384 });
  });

  it("rejects malformed addresses", () => {
    expect(parseSimpleAddress("")).toBeNull();
    expect(parseSimpleAddress("A")).toBeNull();
    expect(parseSimpleAddress("1A")).toBeNull();
    expect(parseSimpleAddress("A1B")).toBeNull();
    expect(parseSimpleAddress("A 1")).toBeNull();
    expect(parseSimpleAddress("a1")).toBeNull(); // lowercase not accepted
  });
});

// ---------------------------------------------------------------------------
// parseRefRange
// ---------------------------------------------------------------------------

describe("parseRefRange", () => {
  it("parses A1:B2", () => {
    expect(parseRefRange("A1:B2")).toEqual({ top: 1, left: 1, bottom: 2, right: 2 });
  });

  it("normalises reversed bounds (B2:A1 == A1:B2)", () => {
    expect(parseRefRange("B2:A1")).toEqual({ top: 1, left: 1, bottom: 2, right: 2 });
  });

  it("handles absolute references", () => {
    expect(parseRefRange("$A$1:$B$2")).toEqual({ top: 1, left: 1, bottom: 2, right: 2 });
  });

  it("single-cell range A1:A1 valid", () => {
    expect(parseRefRange("A1:A1")).toEqual({ top: 1, left: 1, bottom: 1, right: 1 });
  });

  it("rejects malformed ranges", () => {
    expect(parseRefRange("A1")).toBeNull(); // no colon
    expect(parseRefRange("A1:")).toBeNull();
    expect(parseRefRange(":B2")).toBeNull();
    expect(parseRefRange("A1:B2:C3")).toBeNull();
    expect(parseRefRange("bad:ref")).toBeNull();
  });
});

// ---------------------------------------------------------------------------
// parseDefinedNameRange
// ---------------------------------------------------------------------------

describe("parseDefinedNameRange", () => {
  it("parses sheet + absolute cell", () => {
    expect(parseDefinedNameRange("Sheet1!$A$1")).toEqual({
      sheet: "Sheet1",
      startRow: 1,
      startCol: 1,
      endRow: 1,
      endCol: 1
    });
  });

  it("parses sheet + absolute range", () => {
    expect(parseDefinedNameRange("Sheet1!$A$1:$C$10")).toEqual({
      sheet: "Sheet1",
      startRow: 1,
      startCol: 1,
      endRow: 10,
      endCol: 3
    });
  });

  it("parses quoted sheet with spaces", () => {
    expect(parseDefinedNameRange("'My Data'!$B$5")).toEqual({
      sheet: "My Data",
      startRow: 5,
      startCol: 2,
      endRow: 5,
      endCol: 2
    });
  });

  it("unescapes doubled single-quotes in sheet name", () => {
    expect(parseDefinedNameRange("'Mc''Donald'!$A$1")).toEqual({
      sheet: "Mc'Donald",
      startRow: 1,
      startCol: 1,
      endRow: 1,
      endCol: 1
    });
  });

  it("rejects relative references (only $row$col form is a defined-name range)", () => {
    expect(parseDefinedNameRange("Sheet1!A1")).toBeNull();
  });

  it("rejects whole-column and whole-row refs", () => {
    expect(parseDefinedNameRange("Sheet1!$A:$A")).toBeNull();
    expect(parseDefinedNameRange("Sheet1!$1:$5")).toBeNull();
  });

  it("rejects names without sheet prefix", () => {
    expect(parseDefinedNameRange("$A$1")).toBeNull();
  });

  it("rejects malformed input", () => {
    expect(parseDefinedNameRange("")).toBeNull();
    expect(parseDefinedNameRange("Sheet1")).toBeNull();
    expect(parseDefinedNameRange("Sheet1!garbage")).toBeNull();
  });
});
