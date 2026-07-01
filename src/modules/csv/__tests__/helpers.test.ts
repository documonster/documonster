/**
 * CSV Helper Functions Tests
 *
 * Tests for utility functions:
 * - RowHashArray helpers (isRowHashArray, rowHashArrayToMap, etc.)
 * - Header deduplication
 * - Csv.quoted/Csv.unquoted helpers
 * - isFormattedValue
 */

import { isFormattedValue } from "@csv/format/formatted-value";
import { Csv } from "@csv/index";
// Internal helpers - not part of public API
import { rowHashArrayToMap, rowHashArrayGet } from "@csv/utils/row";
import { describe, it, expect } from "vitest";

// =============================================================================
// RowHashArray Helper Functions
// =============================================================================
describe("RowHashArray Helper Functions", () => {
  describe("isRowHashArray", () => {
    it("should return true for valid RowHashArray", () => {
      expect(
        Csv.isRowHashArray([
          ["name", "Alice"],
          ["age", 30]
        ])
      ).toBe(true);
    });

    it("should return false for plain array", () => {
      expect(Csv.isRowHashArray(["Alice", "30"])).toBe(false);
    });

    it("should return false for empty array", () => {
      expect(Csv.isRowHashArray([])).toBe(false);
    });

    it("should return false for object", () => {
      expect(Csv.isRowHashArray({ name: "Alice" })).toBe(false);
    });

    it("should return false for array with non-string keys", () => {
      expect(
        Csv.isRowHashArray([
          [123, "value"],
          ["key", "value"]
        ])
      ).toBe(false);
    });

    it("should return false for array with wrong tuple length", () => {
      expect(Csv.isRowHashArray([["name", "Alice", "extra"]])).toBe(false);
    });
  });

  describe("rowHashArrayToMap", () => {
    it("should convert RowHashArray to object", () => {
      const result = rowHashArrayToMap<string | number>([
        ["name", "Alice"],
        ["age", 30]
      ]);
      expect(result).toEqual({ name: "Alice", age: 30 });
    });

    it("should handle empty RowHashArray", () => {
      expect(rowHashArrayToMap([])).toEqual({});
    });
  });

  describe("rowHashArrayToValues", () => {
    it("should extract values from RowHashArray", () => {
      const result = Csv.rowHashArrayToValues<string | number>([
        ["name", "Alice"],
        ["age", 30]
      ]);
      expect(result).toEqual(["Alice", 30]);
    });

    it("should handle empty RowHashArray", () => {
      expect(Csv.rowHashArrayToValues([])).toEqual([]);
    });
  });

  describe("rowHashArrayToHeaders", () => {
    it("should extract headers from RowHashArray", () => {
      const data: [string, string | number][] = [
        ["name", "Alice"],
        ["age", 30]
      ];
      const result = Csv.rowHashArrayToHeaders(data);
      expect(result).toEqual(["name", "age"]);
    });

    it("should handle empty RowHashArray", () => {
      expect(Csv.rowHashArrayToHeaders([])).toEqual([]);
    });
  });

  describe("rowHashArrayGet", () => {
    it("should get value by key", () => {
      const row: [string, any][] = [
        ["name", "Alice"],
        ["age", 30]
      ];
      expect(rowHashArrayGet(row, "name")).toBe("Alice");
      expect(rowHashArrayGet(row, "age")).toBe(30);
    });

    it("should return undefined for non-existent key", () => {
      const row: [string, any][] = [["name", "Alice"]];
      expect(rowHashArrayGet(row, "city")).toBeUndefined();
    });
  });

  describe("rowHashArrayMapByHeaders", () => {
    it("should map values according to header order", () => {
      const row: [string, any][] = [
        ["name", "Alice"],
        ["age", 30],
        ["city", "NYC"]
      ];
      const result = Csv.rowHashArrayMapByHeaders(row, ["city", "name", "age"]);
      expect(result).toEqual(["NYC", "Alice", 30]);
    });

    it("should return undefined for missing keys", () => {
      const row: [string, any][] = [["name", "Alice"]];
      const result = Csv.rowHashArrayMapByHeaders(row, ["name", "age", "city"]);
      expect(result).toEqual(["Alice", undefined, undefined]);
    });

    it("should handle empty RowHashArray", () => {
      const result = Csv.rowHashArrayMapByHeaders([], ["name", "age"]);
      expect(result).toEqual([undefined, undefined]);
    });

    it("should handle empty headers", () => {
      const row: [string, any][] = [["name", "Alice"]];
      const result = Csv.rowHashArrayMapByHeaders(row, []);
      expect(result).toEqual([]);
    });

    it("should use optimized linear search for small headers", () => {
      // With <= 10 headers, uses linear search per header
      const row: [string, any][] = [
        ["a", 1],
        ["b", 2],
        ["c", 3]
      ];
      const result = Csv.rowHashArrayMapByHeaders(row, ["c", "a", "b"]);
      expect(result).toEqual([3, 1, 2]);
    });

    it("should handle large headers efficiently", () => {
      // With > 10 headers, builds a map once
      const row: [string, any][] = [];
      const headers: string[] = [];
      for (let i = 0; i < 15; i++) {
        row.push([`key${i}`, `value${i}`]);
        headers.push(`key${14 - i}`); // Reverse order
      }
      const result = Csv.rowHashArrayMapByHeaders(row, headers);
      // Should be reversed values
      expect(result[0]).toBe("value14");
      expect(result[14]).toBe("value0");
    });
  });
});

// =============================================================================
// Header Deduplication
// =============================================================================
describe("Header Deduplication", () => {
  describe("deduplicateHeaders", () => {
    it("should rename duplicate headers with suffix", () => {
      expect(Csv.deduplicateHeaders(["A", "B", "A", "A"])).toEqual(["A", "B", "A_1", "A_2"]);
    });

    it("should handle no duplicates", () => {
      expect(Csv.deduplicateHeaders(["A", "B", "C"])).toEqual(["A", "B", "C"]);
    });

    it("should handle multiple different duplicates", () => {
      expect(Csv.deduplicateHeaders(["A", "B", "A", "B", "C"])).toEqual([
        "A",
        "B",
        "A_1",
        "B_1",
        "C"
      ]);
    });

    it("should preserve null/undefined", () => {
      expect(Csv.deduplicateHeaders(["A", null, "A", undefined])).toEqual([
        "A",
        null,
        "A_1",
        undefined
      ]);
    });

    it("should handle empty array", () => {
      expect(Csv.deduplicateHeaders([])).toEqual([]);
    });

    it("should handle all same headers", () => {
      expect(Csv.deduplicateHeaders(["X", "X", "X"])).toEqual(["X", "X_1", "X_2"]);
    });
  });

  describe("deduplicateHeadersWithRenames", () => {
    it("should return null renamedHeaders when no renames occur", () => {
      const { headers, renamedHeaders } = Csv.deduplicateHeadersWithRenames(["A", "B", "C"]);
      expect(headers).toEqual(["A", "B", "C"]);
      expect(renamedHeaders).toBeNull();
    });

    it("should return renamedHeaders mapping candidate -> original", () => {
      const { headers, renamedHeaders } = Csv.deduplicateHeadersWithRenames(["A", "A"]);
      expect(headers).toEqual(["A", "A_1"]);
      expect(renamedHeaders).toEqual({ A_1: "A" });
    });

    it("should avoid collisions with headers that already exist in the input", () => {
      const { headers, renamedHeaders } = Csv.deduplicateHeadersWithRenames(["A", "A", "A_1"]);
      expect(headers).toEqual(["A", "A_2", "A_1"]);
      expect(renamedHeaders).toEqual({ A_2: "A" });
    });

    it("should preserve null/undefined and only rename duplicates", () => {
      const { headers, renamedHeaders } = Csv.deduplicateHeadersWithRenames([
        "A",
        null,
        "A",
        undefined
      ]);
      expect(headers).toEqual(["A", null, "A_1", undefined]);
      expect(renamedHeaders).toEqual({ A_1: "A" });
    });
  });
});

// =============================================================================
// Quoting Helpers
// =============================================================================
describe("Quoting Helpers", () => {
  describe("quoted", () => {
    it("should mark value as pre-quoted", () => {
      const result = Csv.quoted("test");
      expect(isFormattedValue(result)).toBe(true);
      expect(result.value).toBe("test");
      expect(result.quote).toBe(true);
    });

    it("should handle empty string", () => {
      const result = Csv.quoted("");
      expect(result.value).toBe("");
      expect(result.quote).toBe(true);
    });
  });

  describe("unquoted", () => {
    it("should mark value as never quoted", () => {
      const result = Csv.unquoted("test,with,commas");
      expect(isFormattedValue(result)).toBe(true);
      expect(result.value).toBe("test,with,commas");
      expect(result.quote).toBe(false);
    });

    it("should handle empty string", () => {
      const result = Csv.unquoted("");
      expect(result.value).toBe("");
      expect(result.quote).toBe(false);
    });
  });

  describe("isFormattedValue", () => {
    it("should return true for FormattedValue objects", () => {
      expect(isFormattedValue(Csv.quoted("test"))).toBe(true);
      expect(isFormattedValue(Csv.unquoted("test"))).toBe(true);
    });

    it("should return false for plain values", () => {
      expect(isFormattedValue("test")).toBe(false);
      expect(isFormattedValue(123)).toBe(false);
      expect(isFormattedValue(null)).toBe(false);
      expect(isFormattedValue(undefined)).toBe(false);
      expect(isFormattedValue({})).toBe(false);
      expect(isFormattedValue([])).toBe(false);
    });
  });
});
