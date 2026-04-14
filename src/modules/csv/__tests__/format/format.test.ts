/**
 * CSV Format Tests
 *
 * Tests for formatCsv function - basic formatting, quoting, headers, and round-trip.
 *
 * Coverage:
 * - Basic formatting (2D arrays, empty data, single rows)
 * - Quoting (RFC 4180 compliance, special characters)
 * - Custom options (delimiter, quote, rowDelimiter, BOM)
 * - Headers (custom headers, objects, header ordering)
 * - Unicode support
 * - RowHashArray support
 * - Round-trip parsing/formatting
 */

import { formatCsv, parseCsv } from "@csv/index";
import { describe, it, expect } from "vitest";

// =============================================================================
// Basic Formatting
// =============================================================================

describe("formatCsv - Basic Formatting", () => {
  it("should format simple 2D array", () => {
    const data = [
      ["a", "b", "c"],
      ["1", "2", "3"]
    ];
    const result = formatCsv(data);
    expect(result).toBe("a,b,c\n1,2,3");
  });

  it("should use LF as row delimiter by default", () => {
    const data = [["a"], ["b"]];
    const result = formatCsv(data);
    expect(result).toBe("a\nb");
  });

  it("should handle empty data", () => {
    const result = formatCsv([]);
    expect(result).toBe("");
  });

  it("should handle single row", () => {
    const result = formatCsv([["a", "b", "c"]]);
    expect(result).toBe("a,b,c");
  });

  it("should handle null and undefined values", () => {
    const data = [[null, undefined, "value"]];
    const result = formatCsv(data as any);
    expect(result).toBe(",,value");
  });

  it("should convert numbers and booleans to strings", () => {
    const data = [[1, 2.5, true, false]];
    const result = formatCsv(data as any);
    expect(result).toBe("1,2.5,true,false");
  });

  it("should format numbers with comma decimalSeparator", () => {
    const data = [[1, 2.5]];
    const result = formatCsv(data as any, { delimiter: ";", decimalSeparator: "," });
    expect(result).toBe("1;2,5");
  });
});

// =============================================================================
// Quoting (RFC 4180 Section 2.5, 2.6, 2.7)
// =============================================================================

describe("formatCsv - Quoting", () => {
  it("should quote fields containing commas", () => {
    const data = [["hello, world", "test"]];
    const result = formatCsv(data);
    expect(result).toBe('"hello, world",test');
  });

  it("should quote fields containing double-quotes and escape them", () => {
    const data = [['He said "Hello"', "test"]];
    const result = formatCsv(data);
    expect(result).toBe('"He said ""Hello""",test');
  });

  it("should quote fields containing newlines", () => {
    const data = [["line1\nline2", "test"]];
    const result = formatCsv(data);
    expect(result).toBe('"line1\nline2",test');
  });

  it("should quote fields containing CRLF", () => {
    const data = [["line1\r\nline2", "test"]];
    const result = formatCsv(data);
    expect(result).toBe('"line1\r\nline2",test');
  });

  it("should quote fields containing only quotes", () => {
    const data = [['"']];
    const result = formatCsv(data);
    expect(result).toBe('""""');
  });

  it("should handle quoteColumns: true option", () => {
    const data = [["a", "b", "c"]];
    const result = formatCsv(data, { quoteColumns: true });
    expect(result).toBe('"a","b","c"');
  });

  it("should not quote when quote is disabled (false)", () => {
    const data = [["hello, world", "test"]];
    const result = formatCsv(data, { quote: false });
    // When quoting is disabled, commas are literal (may break parsing)
    expect(result).toBe("hello, world,test");
  });

  it("should not quote when quote is disabled (null)", () => {
    const data = [["hello, world", "test"]];
    const result = formatCsv(data, { quote: null });
    expect(result).toBe("hello, world,test");
  });
});

// =============================================================================
// Custom Options
// =============================================================================

describe("formatCsv - Custom Options", () => {
  it("should support custom delimiter", () => {
    const data = [["a", "b", "c"]];
    const result = formatCsv(data, { delimiter: "\t" });
    expect(result).toBe("a\tb\tc");
  });

  it("should support custom quote character", () => {
    const data = [["hello, world", "test"]];
    const result = formatCsv(data, { quote: "'" });
    expect(result).toBe("'hello, world',test");
  });

  it("should support custom row delimiter", () => {
    const data = [
      ["a", "b"],
      ["1", "2"]
    ];
    const result = formatCsv(data, { lineEnding: "\r\n" });
    expect(result).toBe("a,b\r\n1,2");
  });

  it("should add BOM when bom is true", () => {
    const data = [["a", "b"]];
    const result = formatCsv(data, { bom: true });
    expect(result.charCodeAt(0)).toBe(0xfeff);
    expect(result).toBe("\uFEFFa,b");
  });
});

// =============================================================================
// Headers
// =============================================================================

describe("formatCsv - Headers", () => {
  it("should add custom headers to 2D array", () => {
    const data = [
      ["1", "2", "3"],
      ["4", "5", "6"]
    ];
    const result = formatCsv(data, { headers: ["a", "b", "c"] });
    expect(result).toBe("a,b,c\n1,2,3\n4,5,6");
  });

  it("should format array of objects with headers: true", () => {
    const data = [
      { name: "Alice", age: "30" },
      { name: "Bob", age: "25" }
    ];
    const result = formatCsv(data, { headers: true });
    expect(result).toBe("name,age\nAlice,30\nBob,25");
  });

  it("should format objects with custom header order", () => {
    const data = [
      { name: "Alice", age: "30", city: "NYC" },
      { name: "Bob", age: "25", city: "LA" }
    ];
    const result = formatCsv(data, { headers: ["city", "name", "age"] });
    expect(result).toBe("city,name,age\nNYC,Alice,30\nLA,Bob,25");
  });

  it("should format empty data with writeHeaders: true", () => {
    const data: string[][] = [];
    const result = formatCsv(data, {
      headers: ["name", "age"],
      writeHeaders: true
    });
    expect(result).toBe("name,age");
  });

  it("should not write headers when writeHeaders: false", () => {
    const data = [
      { name: "Alice", age: "30" },
      { name: "Bob", age: "25" }
    ];
    const result = formatCsv(data, { headers: true, writeHeaders: false });
    expect(result).toBe("Alice,30\nBob,25");
  });
});

// =============================================================================
// Unicode
// =============================================================================

describe("formatCsv - Unicode", () => {
  it("should format UTF-8 characters", () => {
    const data = [
      ["名前", "年齢"],
      ["田中", "30"]
    ];
    const result = formatCsv(data);
    expect(result).toBe("名前,年齢\n田中,30");
  });

  it("should quote Unicode fields containing delimiters", () => {
    const data = [["你好,世界", "测试"]];
    const result = formatCsv(data);
    expect(result).toBe('"你好,世界",测试');
  });
});

// =============================================================================
// RowHashArray Support
// =============================================================================

describe("formatCsv - RowHashArray Support", () => {
  it("should format RowHashArray (array of [key, value] tuples)", () => {
    const data: [string, any][][] = [
      [
        ["name", "Alice"],
        ["age", 30]
      ],
      [
        ["name", "Bob"],
        ["age", 25]
      ]
    ];
    const result = formatCsv(data);
    expect(result).toBe("Alice,30\nBob,25");
  });

  it("should format RowHashArray with headers: true", () => {
    const data: [string, any][][] = [
      [
        ["name", "Alice"],
        ["age", 30]
      ],
      [
        ["name", "Bob"],
        ["age", 25]
      ]
    ];
    const result = formatCsv(data, { headers: true });
    expect(result).toBe("name,age\nAlice,30\nBob,25");
  });

  it("should format RowHashArray with custom headers (reorder columns)", () => {
    const data: [string, any][][] = [
      [
        ["name", "Alice"],
        ["age", 30],
        ["city", "NYC"]
      ],
      [
        ["name", "Bob"],
        ["age", 25],
        ["city", "LA"]
      ]
    ];
    const result = formatCsv(data, { headers: ["city", "age", "name"] });
    expect(result).toBe("city,age,name\nNYC,30,Alice\nLA,25,Bob");
  });

  it("should handle RowHashArray with missing keys when using custom headers", () => {
    const data: [string, any][][] = [
      [
        ["name", "Alice"],
        ["age", 30]
      ],
      [
        ["name", "Bob"],
        ["city", "LA"]
      ]
    ];
    const result = formatCsv(data, { headers: ["name", "age", "city"] });
    expect(result).toBe("name,age,city\nAlice,30,\nBob,,LA");
  });

  it("should format single RowHashArray row", () => {
    const data: [string, any][][] = [
      [
        ["firstName", "John"],
        ["lastName", "Doe"]
      ]
    ];
    const result = formatCsv(data, { headers: true });
    expect(result).toBe("firstName,lastName\nJohn,Doe");
  });

  it("should format RowHashArray with special characters", () => {
    const data: [string, any][][] = [
      [
        ["message", "Hello, World"],
        ["note", 'He said "hi"']
      ]
    ];
    const result = formatCsv(data);
    expect(result).toBe('"Hello, World","He said ""hi"""');
  });

  it("should format empty RowHashArray with writeHeaders: true", () => {
    const data: [string, any][][] = [];
    const result = formatCsv(data, {
      headers: ["name", "age"],
      writeHeaders: true
    });
    expect(result).toBe("name,age");
  });

  it("should format RowHashArray without writing headers when writeHeaders: false", () => {
    const data: [string, any][][] = [
      [
        ["name", "Alice"],
        ["age", 30]
      ]
    ];
    const result = formatCsv(data, { headers: true, writeHeaders: false });
    expect(result).toBe("Alice,30");
  });
});

// =============================================================================
// Round-trip Tests (Parse + Format)
// =============================================================================

describe("formatCsv - Round-trip Tests", () => {
  it("should round-trip simple data", () => {
    const original = [
      ["a", "b", "c"],
      ["1", "2", "3"]
    ];
    const csv = formatCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });

  it("should round-trip data with commas", () => {
    const original = [["hello, world", "test"]];
    const csv = formatCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });

  it("should round-trip data with quotes", () => {
    const original = [['He said "Hello"', "test"]];
    const csv = formatCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });

  it("should round-trip data with newlines", () => {
    const original = [["line1\nline2", "test"]];
    const csv = formatCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });

  it("should round-trip Unicode data", () => {
    const original = [
      ["名前", "年齢"],
      ["田中,太郎", "30"]
    ];
    const csv = formatCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });

  it("should round-trip complex data", () => {
    const original = [
      ["Name", "Description", "Price"],
      ["Widget", 'A "great" product, really!', "19.99"],
      ["Gadget", "Multi-line\ndescription", "29.99"],
      ["Thing", "", "9.99"]
    ];
    const csv = formatCsv(original);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(original);
  });

  it("should round-trip very long fields", () => {
    const longString = "a".repeat(10000);
    const data = [[longString]];
    const csv = formatCsv(data);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(data);
  });

  it("should round-trip many columns", () => {
    const cols = Array.from({ length: 100 }, (_, i) => `col${i}`);
    const data = [cols];
    const csv = formatCsv(data);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(data);
  });

  it("should round-trip many rows", () => {
    const data = Array.from({ length: 1000 }, (_, i) => [`row${i}`, `value${i}`]);
    const csv = formatCsv(data);
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(data);
  });
});

// =============================================================================
// RFC 4180 Specific Compliance
// =============================================================================

describe("formatCsv - RFC 4180 Compliance", () => {
  it("Rule 1: Each record on separate line with LF", () => {
    const data = [
      ["a", "b"],
      ["c", "d"]
    ];
    const csv = formatCsv(data);
    expect(csv).toBe("a,b\nc,d");
  });

  it("Rule 2: Optional header line", () => {
    const data = [{ col1: "val1", col2: "val2" }];
    const withHeader = formatCsv(data, { headers: true });
    expect(withHeader).toContain("col1,col2");
  });

  it("Rule 5: Fields with CRLF, comma, or quotes MUST be quoted", () => {
    const data = [["hello,world", 'say "hi"', "line1\r\nline2"]];
    const csv = formatCsv(data);
    expect(csv).toContain('"hello,world"');
    expect(csv).toContain('"say ""hi"""');
    expect(csv).toContain('"line1\r\nline2"');
  });

  it("Rule 6: Quote character is double-quote", () => {
    const data = [["test"]];
    const csv = formatCsv(data, { quoteColumns: true });
    expect(csv).toBe('"test"');
  });

  it("Rule 7: Quotes escaped by doubling", () => {
    const data = [['"quoted"']];
    const csv = formatCsv(data);
    expect(csv).toBe('"""quoted"""');

    // Verify round-trip
    const parsed = parseCsv(csv);
    expect(parsed).toEqual(data);
  });
});
