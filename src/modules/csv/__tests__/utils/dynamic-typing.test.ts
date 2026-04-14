/**
 * CSV Dynamic Typing Unit Tests
 *
 * Tests for automatic type conversion functions:
 * - convertValue: Core string-to-type conversion
 * - tryParseDate: ISO date parsing
 * - shouldCastDate: Date casting config check
 * - applyDynamicTyping: Single value with config
 * - applyDynamicTypingToRow: Object row conversion
 * - applyDynamicTypingToArrayRow: Array row conversion
 */

import {
  convertValue,
  tryParseDate,
  shouldCastDate,
  applyDynamicTyping,
  applyDynamicTypingToRow,
  applyDynamicTypingToArrayRow
} from "@csv/utils/dynamic-typing";
import { describe, it, expect } from "vitest";

// =============================================================================
// convertValue Tests
// =============================================================================

describe("convertValue", () => {
  describe("empty string", () => {
    it("returns empty string unchanged", () => {
      expect(convertValue("")).toBe("");
    });
  });

  describe("boolean conversion", () => {
    it("converts 'true' to true (case-insensitive)", () => {
      expect(convertValue("true")).toBe(true);
      expect(convertValue("TRUE")).toBe(true);
      expect(convertValue("True")).toBe(true);
      expect(convertValue("TrUe")).toBe(true);
    });

    it("converts 'false' to false (case-insensitive)", () => {
      expect(convertValue("false")).toBe(false);
      expect(convertValue("FALSE")).toBe(false);
      expect(convertValue("False")).toBe(false);
      expect(convertValue("FaLsE")).toBe(false);
    });

    it("does not convert partial matches", () => {
      expect(convertValue("trueish")).toBe("trueish");
      expect(convertValue("falsey")).toBe("falsey");
      expect(convertValue("tru")).toBe("tru");
      expect(convertValue("fals")).toBe("fals");
    });
  });

  describe("null conversion", () => {
    it("converts 'null' to null (case-insensitive)", () => {
      expect(convertValue("null")).toBe(null);
      expect(convertValue("NULL")).toBe(null);
      expect(convertValue("Null")).toBe(null);
      expect(convertValue("NuLl")).toBe(null);
    });

    it("does not convert partial matches", () => {
      expect(convertValue("nullable")).toBe("nullable");
      expect(convertValue("nul")).toBe("nul");
    });
  });

  describe("number conversion", () => {
    it("converts integers", () => {
      expect(convertValue("0")).toBe(0);
      expect(convertValue("1")).toBe(1);
      expect(convertValue("42")).toBe(42);
      expect(convertValue("-1")).toBe(-1);
      expect(convertValue("-999")).toBe(-999);
    });

    it("converts floating point numbers", () => {
      expect(convertValue("3.14")).toBe(3.14);
      expect(convertValue("0.5")).toBe(0.5);
      expect(convertValue("-2.718")).toBe(-2.718);
      expect(convertValue(".5")).toBe(0.5);
      expect(convertValue("-.25")).toBe(-0.25);
    });

    it("converts scientific notation", () => {
      expect(convertValue("1e10")).toBe(1e10);
      expect(convertValue("1E10")).toBe(1e10);
      expect(convertValue("2.5e-3")).toBe(0.0025);
      expect(convertValue("-1.5E+2")).toBe(-150);
    });

    it("converts special numeric values", () => {
      expect(convertValue("Infinity")).toBe(Infinity);
      expect(convertValue("-Infinity")).toBe(-Infinity);
      expect(convertValue("NaN")).toBeNaN();
    });
  });

  describe("leading zeros preservation", () => {
    it("preserves leading zeros (for zip codes, IDs)", () => {
      expect(convertValue("007")).toBe("007");
      expect(convertValue("00123")).toBe("00123");
      expect(convertValue("0123456789")).toBe("0123456789");
    });

    it("preserves negative leading zeros", () => {
      expect(convertValue("-007")).toBe("-007");
      expect(convertValue("-00123")).toBe("-00123");
    });

    it("allows single zero", () => {
      expect(convertValue("0")).toBe(0);
    });

    it("allows zero with decimal", () => {
      expect(convertValue("0.5")).toBe(0.5);
      expect(convertValue("-0.5")).toBe(-0.5);
    });
  });

  describe("whitespace handling", () => {
    it("preserves strings with leading whitespace", () => {
      expect(convertValue(" 123")).toBe(" 123");
      expect(convertValue("\t456")).toBe("\t456");
    });

    it("preserves strings with trailing whitespace", () => {
      expect(convertValue("123 ")).toBe("123 ");
      expect(convertValue("456\n")).toBe("456\n");
    });
  });

  describe("invalid number formats", () => {
    it("preserves strings with letters mixed in", () => {
      expect(convertValue("123abc")).toBe("123abc");
      expect(convertValue("abc123")).toBe("abc123");
      expect(convertValue("12.34.56")).toBe("12.34.56");
    });

    it("preserves special characters", () => {
      expect(convertValue("$100")).toBe("$100");
      expect(convertValue("100%")).toBe("100%");
      expect(convertValue("1,000")).toBe("1,000");
    });
  });

  describe("regular strings", () => {
    it("preserves regular text", () => {
      expect(convertValue("hello")).toBe("hello");
      expect(convertValue("Hello World")).toBe("Hello World");
      expect(convertValue("test@example.com")).toBe("test@example.com");
    });
  });
});

// =============================================================================
// tryParseDate Tests
// =============================================================================

describe("tryParseDate", () => {
  describe("valid ISO dates", () => {
    it("parses YYYY-MM-DD format", () => {
      const date = tryParseDate("2024-01-15");
      expect(date).toBeInstanceOf(Date);
      expect(date!.getFullYear()).toBe(2024);
      expect(date!.getMonth()).toBe(0); // January
      expect(date!.getDate()).toBe(15);
    });

    it("parses YYYY-MM-DDTHH:mm:ss format", () => {
      const date = tryParseDate("2024-06-20T14:30:00");
      expect(date).toBeInstanceOf(Date);
      expect(date!.getFullYear()).toBe(2024);
      expect(date!.getMonth()).toBe(5); // June
      // Note: hours depend on timezone, just verify it's a valid date
      expect(date!.getTime()).not.toBeNaN();
    });

    it("parses YYYY-MM-DDTHH:mm:ssZ format", () => {
      const date = tryParseDate("2024-12-31T23:59:59Z");
      expect(date).toBeInstanceOf(Date);
    });

    it("parses YYYY-MM-DDTHH:mm:ss.SSSZ format", () => {
      const date = tryParseDate("2024-03-15T10:20:30.456Z");
      expect(date).toBeInstanceOf(Date);
    });
  });

  describe("invalid inputs", () => {
    it("returns null for empty string", () => {
      expect(tryParseDate("")).toBe(null);
    });

    it("returns null for short strings", () => {
      expect(tryParseDate("2024")).toBe(null);
      expect(tryParseDate("2024-01")).toBe(null);
    });

    it("returns null for non-date strings", () => {
      expect(tryParseDate("hello world")).toBe(null);
      expect(tryParseDate("not a date")).toBe(null);
    });

    it("returns null for invalid date formats", () => {
      expect(tryParseDate("01/15/2024")).toBe(null); // US format
      expect(tryParseDate("15-01-2024")).toBe(null); // European format
    });
  });
});

// =============================================================================
// shouldCastDate Tests
// =============================================================================

describe("shouldCastDate", () => {
  describe("undefined/false config", () => {
    it("returns false when castDate is undefined", () => {
      expect(shouldCastDate(undefined, "date")).toBe(false);
    });

    it("returns false when castDate is false", () => {
      expect(shouldCastDate(false, "date")).toBe(false);
    });
  });

  describe("true config", () => {
    it("returns true for any column when castDate is true", () => {
      expect(shouldCastDate(true, "date")).toBe(true);
      expect(shouldCastDate(true, "created_at")).toBe(true);
      expect(shouldCastDate(true, "anything")).toBe(true);
      expect(shouldCastDate(true, undefined)).toBe(true);
    });
  });

  describe("array config", () => {
    it("returns true for columns in the array", () => {
      const config = ["date", "created_at", "updated_at"];
      expect(shouldCastDate(config, "date")).toBe(true);
      expect(shouldCastDate(config, "created_at")).toBe(true);
      expect(shouldCastDate(config, "updated_at")).toBe(true);
    });

    it("returns false for columns not in the array", () => {
      const config = ["date", "created_at"];
      expect(shouldCastDate(config, "name")).toBe(false);
      expect(shouldCastDate(config, "email")).toBe(false);
    });

    it("returns false for numeric column index", () => {
      const config = ["date", "created_at"];
      expect(shouldCastDate(config, 0)).toBe(false);
      expect(shouldCastDate(config, 1)).toBe(false);
    });

    it("returns false for undefined column name", () => {
      const config = ["date"];
      expect(shouldCastDate(config, undefined)).toBe(false);
    });
  });
});

// =============================================================================
// applyDynamicTyping Tests
// =============================================================================

describe("applyDynamicTyping", () => {
  describe("config = false", () => {
    it("returns value unchanged", () => {
      expect(applyDynamicTyping("123", false)).toBe("123");
      expect(applyDynamicTyping("true", false)).toBe("true");
      expect(applyDynamicTyping("null", false)).toBe("null");
    });
  });

  describe("config = true", () => {
    it("applies default conversion", () => {
      expect(applyDynamicTyping("123", true)).toBe(123);
      expect(applyDynamicTyping("true", true)).toBe(true);
      expect(applyDynamicTyping("null", true)).toBe(null);
    });
  });

  describe("custom converter function", () => {
    it("uses custom function for conversion", () => {
      const toUpperCase = (v: string) => v.toUpperCase();
      expect(applyDynamicTyping("hello", toUpperCase)).toBe("HELLO");
    });

    it("can return any type", () => {
      const toArray = (v: string) => v.split(",");
      expect(applyDynamicTyping("a,b,c", toArray)).toEqual(["a", "b", "c"]);
    });

    it("receives original string value", () => {
      const spy = (v: string) => {
        expect(v).toBe("test");
        return v;
      };
      applyDynamicTyping("test", spy);
    });
  });
});

// =============================================================================
// applyDynamicTypingToRow Tests
// =============================================================================

describe("applyDynamicTypingToRow", () => {
  describe("fast path", () => {
    it("returns same object when dynamicTyping=false and no castDate", () => {
      const row = { a: "1", b: "2" };
      const result = applyDynamicTypingToRow(row, false);
      expect(result).toBe(row); // Same reference
    });
  });

  describe("global dynamicTyping = true", () => {
    it("converts all fields", () => {
      const row = { name: "Alice", age: "30", active: "true" };
      const result = applyDynamicTypingToRow(row, true);
      expect(result).toEqual({ name: "Alice", age: 30, active: true });
    });
  });

  describe("per-column config", () => {
    it("converts only specified columns", () => {
      const row = { name: "Bob", age: "25", zip: "00123" };
      const result = applyDynamicTypingToRow(row, { age: true, zip: false });
      expect(result).toEqual({ name: "Bob", age: 25, zip: "00123" });
    });

    it("uses custom converter for specific column", () => {
      const row = { name: "alice", score: "100" };
      const result = applyDynamicTypingToRow(row, {
        name: (v: string) => v.toUpperCase(),
        score: true
      });
      expect(result).toEqual({ name: "ALICE", score: 100 });
    });

    it("leaves unconfigured columns as strings", () => {
      const row = { a: "1", b: "2", c: "3" };
      const result = applyDynamicTypingToRow(row, { a: true });
      expect(result).toEqual({ a: 1, b: "2", c: "3" });
    });
  });

  describe("with castDate", () => {
    it("parses dates when castDate=true", () => {
      const row = { date: "2024-01-15", name: "Event" };
      const result = applyDynamicTypingToRow(row, false, true);
      expect(result.date).toBeInstanceOf(Date);
      expect(result.name).toBe("Event");
    });

    it("parses dates for specified columns only", () => {
      const row = { created: "2024-01-15", note: "2024-01-15" };
      const result = applyDynamicTypingToRow(row, false, ["created"]);
      expect(result.created).toBeInstanceOf(Date);
      expect(result.note).toBe("2024-01-15");
    });

    it("combines with dynamicTyping", () => {
      const row = { date: "2024-01-15", count: "42", active: "true" };
      const result = applyDynamicTypingToRow(row, true, true);
      expect(result.date).toBeInstanceOf(Date);
      expect(result.count).toBe(42);
      expect(result.active).toBe(true);
    });
  });
});

// =============================================================================
// applyDynamicTypingToArrayRow Tests
// =============================================================================

describe("applyDynamicTypingToArrayRow", () => {
  describe("fast path", () => {
    it("returns same array when dynamicTyping=false and no castDate", () => {
      const row = ["1", "2", "3"];
      const result = applyDynamicTypingToArrayRow(row, null, false);
      expect(result).toBe(row); // Same reference
    });
  });

  describe("global dynamicTyping = true", () => {
    it("converts all values", () => {
      const row = ["Alice", "30", "true", "null"];
      const result = applyDynamicTypingToArrayRow(row, null, true);
      expect(result).toEqual(["Alice", 30, true, null]);
    });

    it("works without headers", () => {
      const row = ["123", "456"];
      const result = applyDynamicTypingToArrayRow(row, null, true);
      expect(result).toEqual([123, 456]);
    });
  });

  describe("per-column config with headers", () => {
    it("converts columns by header name", () => {
      const row = ["Bob", "25", "00123"];
      const headers = ["name", "age", "zip"];
      const result = applyDynamicTypingToArrayRow(row, headers, { age: true, zip: false });
      expect(result).toEqual(["Bob", 25, "00123"]);
    });

    it("returns unchanged when per-column config but no headers", () => {
      const row = ["1", "2", "3"];
      const result = applyDynamicTypingToArrayRow(row, null, { col0: true });
      expect(result).toBe(row); // Same reference
    });
  });

  describe("with castDate", () => {
    it("parses dates when castDate=true", () => {
      const row = ["2024-01-15", "Event"];
      const result = applyDynamicTypingToArrayRow(row, null, false, true);
      expect(result[0]).toBeInstanceOf(Date);
      expect(result[1]).toBe("Event");
    });

    it("parses dates for specified columns with headers", () => {
      const row = ["2024-01-15", "2024-01-15"];
      const headers = ["created", "note"];
      const result = applyDynamicTypingToArrayRow(row, headers, false, ["created"]);
      expect(result[0]).toBeInstanceOf(Date);
      expect(result[1]).toBe("2024-01-15");
    });
  });
});
