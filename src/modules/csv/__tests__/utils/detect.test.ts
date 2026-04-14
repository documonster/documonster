/**
 * CSV Detection Utilities Unit Tests
 *
 * Tests for detection and normalization functions:
 * - escapeRegex: Escape special regex characters
 * - normalizeQuoteOption: Normalize quote config
 * - normalizeEscapeOption: Normalize escape config
 * - startsWithFormulaChar: CSV injection prevention
 * - stripBom: UTF-8 BOM removal (tested in encoding.test.ts but expanded here)
 * - detectLinebreak: Line ending detection (tested in encoding.test.ts but expanded here)
 * - detectDelimiter: Delimiter auto-detection (tested in parse-options.test.ts but expanded here)
 */

import {
  escapeRegex,
  normalizeQuoteOption,
  normalizeEscapeOption,
  startsWithFormulaChar,
  stripBom,
  detectLinebreak,
  detectDelimiter
} from "@csv/utils/detect";
import { describe, it, expect } from "vitest";

// =============================================================================
// escapeRegex Tests
// =============================================================================

describe("escapeRegex", () => {
  it("escapes special regex characters", () => {
    expect(escapeRegex(".")).toBe("\\.");
    expect(escapeRegex("*")).toBe("\\*");
    expect(escapeRegex("+")).toBe("\\+");
    expect(escapeRegex("?")).toBe("\\?");
    expect(escapeRegex("^")).toBe("\\^");
    expect(escapeRegex("$")).toBe("\\$");
    expect(escapeRegex("{")).toBe("\\{");
    expect(escapeRegex("}")).toBe("\\}");
    expect(escapeRegex("(")).toBe("\\(");
    expect(escapeRegex(")")).toBe("\\)");
    expect(escapeRegex("|")).toBe("\\|");
    expect(escapeRegex("[")).toBe("\\[");
    expect(escapeRegex("]")).toBe("\\]");
    expect(escapeRegex("\\")).toBe("\\\\");
  });

  it("escapes multiple special characters", () => {
    expect(escapeRegex("a.b*c")).toBe("a\\.b\\*c");
    expect(escapeRegex("[a-z]+")).toBe("\\[a-z\\]\\+");
    expect(escapeRegex("(foo|bar)")).toBe("\\(foo\\|bar\\)");
  });

  it("leaves normal characters unchanged", () => {
    expect(escapeRegex("hello")).toBe("hello");
    expect(escapeRegex("abc123")).toBe("abc123");
    expect(escapeRegex("")).toBe("");
  });

  it("produces valid regex pattern", () => {
    const input = "test.*pattern+";
    const escaped = escapeRegex(input);
    const regex = new RegExp(escaped);
    expect(regex.test(input)).toBe(true);
    expect(regex.test("test_any_pattern")).toBe(false);
  });
});

// =============================================================================
// normalizeQuoteOption Tests
// =============================================================================

describe("normalizeQuoteOption", () => {
  it("returns disabled for false", () => {
    expect(normalizeQuoteOption(false)).toEqual({ enabled: false, char: "" });
  });

  it("returns disabled for null", () => {
    expect(normalizeQuoteOption(null)).toEqual({ enabled: false, char: "" });
  });

  it("returns default double-quote for undefined", () => {
    expect(normalizeQuoteOption(undefined)).toEqual({ enabled: true, char: '"' });
  });

  it("returns custom quote character when provided", () => {
    expect(normalizeQuoteOption("'")).toEqual({ enabled: true, char: "'" });
    expect(normalizeQuoteOption("`")).toEqual({ enabled: true, char: "`" });
  });

  it("handles empty string as custom quote", () => {
    expect(normalizeQuoteOption("")).toEqual({ enabled: true, char: "" });
  });
});

// =============================================================================
// normalizeEscapeOption Tests
// =============================================================================

describe("normalizeEscapeOption", () => {
  it("returns disabled for false", () => {
    expect(normalizeEscapeOption(false, '"')).toEqual({ enabled: false, char: "" });
  });

  it("returns disabled for null", () => {
    expect(normalizeEscapeOption(null, '"')).toEqual({ enabled: false, char: "" });
  });

  it("uses quote char as default when undefined", () => {
    expect(normalizeEscapeOption(undefined, '"')).toEqual({ enabled: true, char: '"' });
    expect(normalizeEscapeOption(undefined, "'")).toEqual({ enabled: true, char: "'" });
  });

  it("returns custom escape character when provided", () => {
    expect(normalizeEscapeOption("\\", '"')).toEqual({ enabled: true, char: "\\" });
    expect(normalizeEscapeOption("~", '"')).toEqual({ enabled: true, char: "~" });
  });
});

// =============================================================================
// startsWithFormulaChar Tests
// =============================================================================

describe("startsWithFormulaChar", () => {
  describe("formula prefix characters", () => {
    it("detects equals sign", () => {
      expect(startsWithFormulaChar("=SUM(A1:A10)")).toBe(true);
      expect(startsWithFormulaChar("=1+1")).toBe(true);
    });

    it("detects plus sign", () => {
      expect(startsWithFormulaChar("+100")).toBe(true);
      expect(startsWithFormulaChar("+cmd|' /C calc'!A0")).toBe(true);
    });

    it("detects minus sign", () => {
      expect(startsWithFormulaChar("-100")).toBe(true);
      expect(startsWithFormulaChar("-@SUM(A1)")).toBe(true);
    });

    it("detects at sign", () => {
      expect(startsWithFormulaChar("@SUM(A1:A10)")).toBe(true);
    });
  });

  describe("whitespace characters", () => {
    it("detects tab", () => {
      expect(startsWithFormulaChar("\tdata")).toBe(true);
    });

    it("detects carriage return", () => {
      expect(startsWithFormulaChar("\rdata")).toBe(true);
    });

    it("detects line feed", () => {
      expect(startsWithFormulaChar("\ndata")).toBe(true);
    });
  });

  describe("full-width characters (Unicode)", () => {
    it("detects full-width equals (＝)", () => {
      expect(startsWithFormulaChar("\uFF1DSUM")).toBe(true);
    });

    it("detects full-width plus (＋)", () => {
      expect(startsWithFormulaChar("\uFF0B100")).toBe(true);
    });

    it("detects full-width minus (－)", () => {
      expect(startsWithFormulaChar("\uFF0D100")).toBe(true);
    });

    it("detects full-width at (＠)", () => {
      expect(startsWithFormulaChar("\uFF20SUM")).toBe(true);
    });
  });

  describe("safe strings", () => {
    it("returns false for empty string", () => {
      expect(startsWithFormulaChar("")).toBe(false);
    });

    it("returns false for normal text", () => {
      expect(startsWithFormulaChar("Hello")).toBe(false);
      expect(startsWithFormulaChar("123")).toBe(false);
      expect(startsWithFormulaChar("test@example.com")).toBe(false); // @ not at start
    });

    it("returns false for formula chars not at start", () => {
      expect(startsWithFormulaChar("a=b")).toBe(false);
      expect(startsWithFormulaChar("x+y")).toBe(false);
      expect(startsWithFormulaChar("1-2")).toBe(false);
    });
  });
});

// =============================================================================
// stripBom Extended Tests
// =============================================================================

describe("stripBom", () => {
  const UTF8_BOM = "\ufeff";

  it("removes BOM from start", () => {
    expect(stripBom(UTF8_BOM + "hello")).toBe("hello");
  });

  it("returns unchanged if no BOM", () => {
    expect(stripBom("hello")).toBe("hello");
  });

  it("handles empty string", () => {
    expect(stripBom("")).toBe("");
  });

  it("handles BOM-only string", () => {
    expect(stripBom(UTF8_BOM)).toBe("");
  });

  it("does not remove BOM from middle of string", () => {
    expect(stripBom("hello" + UTF8_BOM + "world")).toBe("hello" + UTF8_BOM + "world");
  });

  it("works with CSV content", () => {
    const csvWithBom = UTF8_BOM + "name,age\nAlice,30";
    expect(stripBom(csvWithBom)).toBe("name,age\nAlice,30");
  });
});

// =============================================================================
// detectLinebreak Extended Tests
// =============================================================================

describe("detectLinebreak", () => {
  it("detects LF (Unix)", () => {
    expect(detectLinebreak("a,b\nc,d")).toBe("\n");
  });

  it("detects CRLF (Windows)", () => {
    expect(detectLinebreak("a,b\r\nc,d")).toBe("\r\n");
  });

  it("detects CR (old Mac)", () => {
    expect(detectLinebreak("a,b\rc,d")).toBe("\r");
  });

  it("returns LF as default when no newlines", () => {
    expect(detectLinebreak("a,b,c")).toBe("\n");
    expect(detectLinebreak("")).toBe("\n");
  });

  it("uses first newline found for mixed line endings", () => {
    expect(detectLinebreak("a\nb\r\nc")).toBe("\n"); // LF comes first
    expect(detectLinebreak("a\r\nb\nc")).toBe("\r\n"); // CRLF comes first
  });

  it("handles quoted fields with newlines", () => {
    // Note: detectLinebreak uses simple detection (doesn't parse quotes)
    // It finds the first newline character regardless of context
    expect(detectLinebreak('"field\nwith\nnewlines"')).toBe("\n");
  });
});

// =============================================================================
// detectDelimiter Extended Tests
// =============================================================================

describe("detectDelimiter", () => {
  describe("common delimiters", () => {
    it("detects comma", () => {
      expect(detectDelimiter("a,b,c\n1,2,3")).toBe(",");
    });

    it("detects semicolon", () => {
      expect(detectDelimiter("a;b;c\n1;2;3")).toBe(";");
    });

    it("detects tab", () => {
      expect(detectDelimiter("a\tb\tc\n1\t2\t3")).toBe("\t");
    });

    it("detects pipe", () => {
      expect(detectDelimiter("a|b|c\n1|2|3")).toBe("|");
    });
  });

  describe("custom delimiters", () => {
    it("detects custom delimiter from provided list", () => {
      expect(detectDelimiter("a:b:c\n1:2:3", '"', [":"])).toBe(":");
      expect(detectDelimiter("a~b~c\n1~2~3", '"', ["~"])).toBe("~");
    });
  });

  describe("edge cases", () => {
    it("returns default for empty input", () => {
      expect(detectDelimiter("")).toBe(",");
    });

    it("returns default for single column", () => {
      expect(detectDelimiter("a\nb\nc")).toBe(",");
    });

    it("handles quoted fields with delimiters inside", () => {
      // The comma inside quotes should not count
      expect(detectDelimiter('"a,b";c\n"1,2";3')).toBe(";");
    });

    it("handles escaped quotes", () => {
      expect(detectDelimiter('"""a""",b,c\n1,2,3')).toBe(",");
    });
  });

  describe("with options", () => {
    it("skips comment lines", () => {
      const csv = "#header\na,b,c\n1,2,3";
      expect(detectDelimiter(csv, '"', undefined, "#")).toBe(",");
    });

    it("handles skipEmptyLines", () => {
      const csv = "a,b,c\n\n\n1,2,3";
      expect(detectDelimiter(csv, '"', undefined, undefined, true)).toBe(",");
    });
  });

  describe("consistency scoring", () => {
    it("prefers delimiter with consistent field counts", () => {
      // Semicolon has consistent 3 fields per row
      // Comma appears inconsistently
      const csv = "a;b;c\n1;2;3\n4;5;6";
      expect(detectDelimiter(csv)).toBe(";");
    });

    it("prefers delimiter with more fields when consistent", () => {
      // Both comma and semicolon are present, but comma gives more fields
      const csv = "a,b,c,d\n1,2,3,4";
      expect(detectDelimiter(csv)).toBe(",");
    });
  });
});
