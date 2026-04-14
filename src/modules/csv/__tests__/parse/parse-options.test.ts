/**
 * CSV Parse Options Tests
 *
 * Consolidated tests for all parse options:
 * - dynamicTyping: Automatic type conversion
 * - info/raw: Record metadata
 * - toLine: Stop at specific line
 * - castDate: Automatic date parsing
 * - skipRecordsWithError/onSkip: Error recovery
 * - skipRecordsWithEmptyValues: Skip empty records
 * - relaxQuotes: Tolerant quote handling
 * - Delimiter auto-detection
 * - Column mismatch handling
 * - transform/validate callbacks
 */

import { parseCsv, detectDelimiter, type CsvParseResult, type RecordWithInfo } from "@csv/index";
import { CsvParserStream } from "@csv/stream";
import { describe, it, expect } from "vitest";

import { parseStreamCsv } from "../csv-test-utils";

// =============================================================================
// dynamicTyping Tests
// =============================================================================
describe("dynamicTyping Option", () => {
  it("should convert numbers when dynamicTyping is true", () => {
    const csv = "name,age,score\nAlice,25,98.5\nBob,30,87.2";
    const result = parseCsv(csv, {
      headers: true,
      dynamicTyping: true
    }) as CsvParseResult<Record<string, unknown>>;

    expect(result.rows[0]).toEqual({ name: "Alice", age: 25, score: 98.5 });
    expect(typeof result.rows[0].age).toBe("number");
    expect(typeof result.rows[0].score).toBe("number");
    expect(typeof result.rows[0].name).toBe("string");
  });

  it("should convert booleans when dynamicTyping is true", () => {
    const csv = "name,active,verified\nAlice,true,false\nBob,TRUE,FALSE";
    const result = parseCsv(csv, {
      headers: true,
      dynamicTyping: true
    }) as CsvParseResult<Record<string, unknown>>;

    expect(result.rows[0]).toEqual({ name: "Alice", active: true, verified: false });
    expect(typeof result.rows[0].active).toBe("boolean");
  });

  it("should convert null when dynamicTyping is true", () => {
    const csv = "name,value\nAlice,null\nBob,NULL";
    const result = parseCsv(csv, {
      headers: true,
      dynamicTyping: true
    }) as CsvParseResult<Record<string, unknown>>;

    expect(result.rows[0].value).toBeNull();
    expect(result.rows[1].value).toBeNull();
  });

  it("should apply dynamicTyping per column by name", () => {
    const csv = "name,age,zip\nAlice,25,02134\nBob,30,10001";
    const result = parseCsv(csv, {
      headers: true,
      dynamicTyping: { age: true, zip: false }
    }) as CsvParseResult<Record<string, unknown>>;

    expect(result.rows[0]).toEqual({ name: "Alice", age: 25, zip: "02134" });
    expect(typeof result.rows[0].age).toBe("number");
    expect(typeof result.rows[0].zip).toBe("string");
  });

  it("should handle negative numbers", () => {
    const csv = "name,balance\nAlice,-100.50\nBob,200";
    const result = parseCsv(csv, {
      headers: true,
      dynamicTyping: true
    }) as CsvParseResult<Record<string, unknown>>;

    expect(result.rows[0].balance).toBe(-100.5);
    expect(result.rows[1].balance).toBe(200);
  });

  it("should handle scientific notation", () => {
    const csv = "name,value\nAlice,1.5e10\nBob,2E-5";
    const result = parseCsv(csv, {
      headers: true,
      dynamicTyping: true
    }) as CsvParseResult<Record<string, unknown>>;

    expect(result.rows[0].value).toBe(1.5e10);
    expect(result.rows[1].value).toBe(2e-5);
  });

  it("should preserve strings with leading zeros", () => {
    const csv = "name,code\nAlice,00123\nBob,0045";
    const result = parseCsv(csv, {
      headers: true,
      dynamicTyping: true
    }) as CsvParseResult<Record<string, unknown>>;

    // Leading zeros are preserved as strings (important for zip codes, IDs etc.)
    expect(result.rows[0].code).toBe("00123");
    expect(result.rows[1].code).toBe("0045");
  });

  it("should work with array mode (no headers)", () => {
    const csv = "Alice,25,true\nBob,30,false";
    const result = parseCsv(csv, { headers: false, dynamicTyping: true });

    expect(result).toEqual([
      ["Alice", 25, true],
      ["Bob", 30, false]
    ]);
  });

  it("should apply custom converter function", () => {
    const csv = "name,date\nAlice,2024-01-15\nBob,2024-12-25";
    const result = parseCsv(csv, {
      headers: true,
      dynamicTyping: {
        date: (value: string) => new Date(value)
      }
    }) as CsvParseResult<Record<string, unknown>>;

    expect(result.rows[0].date).toBeInstanceOf(Date);
    expect((result.rows[0].date as Date).getFullYear()).toBe(2024);
  });

  it("should work in streaming mode", async () => {
    const csv = "name,age,active\nAlice,25,true\nBob,30,false";
    const parser = new CsvParserStream({ headers: true, dynamicTyping: true });
    const rows = await parseStreamCsv<Record<string, unknown>>(csv, parser);

    expect(rows[0]).toEqual({ name: "Alice", age: 25, active: true });
    expect(typeof rows[0].age).toBe("number");
    expect(typeof rows[0].active).toBe("boolean");
  });
});

// =============================================================================
// info/raw Options Tests
// =============================================================================
describe("info/raw Options", () => {
  it("should return record with info when info: true (headers mode)", () => {
    const csv = "name,age\nAlice,30\nBob,25";
    const result = parseCsv(csv, { headers: true, info: true }) as CsvParseResult<
      RecordWithInfo<Record<string, unknown>>
    >;

    expect(result.rows[0]).toHaveProperty("record");
    expect(result.rows[0]).toHaveProperty("info");
    expect(result.rows[0].record).toEqual({ name: "Alice", age: "30" });
    expect(result.rows[0].info.index).toBe(0);
    expect(result.rows[0].info.line).toBe(2); // 1-based, header is line 1
    expect(result.rows[0].info.quoted).toEqual([false, false]);
  });

  it("should return record with info when info: true (array mode)", () => {
    const csv = "Alice,30\nBob,25";
    const result = parseCsv(csv, { info: true }) as CsvParseResult<RecordWithInfo<string[]>>;

    expect(result.rows[0].record).toEqual(["Alice", "30"]);
    expect(result.rows[0].info.index).toBe(0);
    expect(result.rows[0].info.line).toBe(1);
  });

  it("should track quoted fields correctly", () => {
    const csv = '"Alice",30\nBob,"25"';
    const result = parseCsv(csv, { info: true }) as CsvParseResult<RecordWithInfo<string[]>>;

    expect(result.rows[0].info.quoted).toEqual([true, false]);
    expect(result.rows[1].info.quoted).toEqual([false, true]);
  });

  it("should track character offset correctly", () => {
    const csv = "a,b\n1,2\n3,4";
    const result = parseCsv(csv, { headers: true, info: true }) as CsvParseResult<
      RecordWithInfo<Record<string, unknown>>
    >;

    expect(result.rows[0].info.offset).toBe(4); // "a,b\n" = 4 characters
    expect(result.rows[1].info.offset).toBe(8); // "a,b\n1,2\n" = 8 characters
  });

  it("should include raw string when raw: true", () => {
    const csv = '"Alice",30\nBob,"25"';
    const result = parseCsv(csv, { info: true, raw: true }) as CsvParseResult<
      RecordWithInfo<string[]>
    >;

    expect(result.rows[0].info.raw).toBe('"Alice",30');
    expect(result.rows[1].info.raw).toBe('Bob,"25"');
  });

  it("should not include raw string when raw: false", () => {
    const csv = "Alice,30";
    const result = parseCsv(csv, { info: true, raw: false }) as CsvParseResult<
      RecordWithInfo<string[]>
    >;

    expect(result.rows[0].info.raw).toBeUndefined();
  });
});

// =============================================================================
// toLine Option Tests
// =============================================================================
describe("toLine Option", () => {
  it("should stop parsing at specified line number (no headers)", () => {
    const csv = "a,b\n1,2\n3,4\n5,6\n7,8";
    const result = parseCsv(csv, { toLine: 3 }) as string[][];
    expect(result).toEqual([
      ["a", "b"],
      ["1", "2"],
      ["3", "4"]
    ]);
  });

  it("should stop parsing at specified line number (with headers)", () => {
    const csv = "name,age\nAlice,30\nBob,25\nCharlie,35\nDave,40";
    const result = parseCsv(csv, { headers: true, toLine: 3 }) as any;
    expect(result.rows).toEqual([
      { name: "Alice", age: "30" },
      { name: "Bob", age: "25" }
    ]);
    expect(result.meta.truncated).toBe(true);
  });

  it("should work with skipLines + toLine", () => {
    const csv = "meta\na,b\n1,2\n3,4\n5,6";
    const result = parseCsv(csv, { skipLines: 1, toLine: 4 }) as string[][];
    expect(result).toEqual([
      ["a", "b"],
      ["1", "2"],
      ["3", "4"]
    ]);
  });

  it("should work in fastMode", () => {
    const csv = "a,b\n1,2\n3,4\n5,6";
    const result = parseCsv(csv, { fastMode: true, toLine: 2 }) as string[][];
    expect(result).toEqual([
      ["a", "b"],
      ["1", "2"]
    ]);
  });

  it("should handle toLine: 1 (only first line)", () => {
    const csv = "a,b\n1,2\n3,4";
    const result = parseCsv(csv, { toLine: 1 }) as string[][];
    expect(result).toEqual([["a", "b"]]);
  });
});

// =============================================================================
// castDate Option Tests
// =============================================================================
describe("castDate Option", () => {
  it("should parse ISO dates when castDate: true", () => {
    const csv = "date,value\n2024-01-15,100\n2024-06-30,200";
    const result = parseCsv(csv, { headers: true, castDate: true }) as any;

    expect(result.rows[0].date).toBeInstanceOf(Date);
    const date = result.rows[0].date as Date;
    expect(date.getFullYear()).toBe(2024);
    expect(date.getMonth()).toBe(0); // January is 0
    expect(date.getDate()).toBe(15);
    expect(result.rows[0].value).toBe("100"); // Not a date, stays string
  });

  it("should parse ISO datetime with T separator", () => {
    const csv = "timestamp\n2024-01-15T10:30:00";
    const result = parseCsv(csv, { headers: true, castDate: true }) as any;

    expect(result.rows[0].timestamp).toBeInstanceOf(Date);
  });

  it("should parse ISO datetime with Z suffix (UTC)", () => {
    const csv = "timestamp\n2024-01-15T10:30:00Z";
    const result = parseCsv(csv, { headers: true, castDate: true }) as any;

    expect(result.rows[0].timestamp).toBeInstanceOf(Date);
  });

  it("should not convert non-date strings", () => {
    const csv = "value\nhello\n12345\ntrue";
    const result = parseCsv(csv, { headers: true, castDate: true }) as any;

    expect(result.rows[0].value).toBe("hello");
    expect(result.rows[1].value).toBe("12345");
    expect(result.rows[2].value).toBe("true");
  });
});

// =============================================================================
// Delimiter Auto-Detection Tests
// =============================================================================
describe("Delimiter Auto-Detection", () => {
  describe("detectDelimiter", () => {
    it("should detect comma delimiter", () => {
      expect(detectDelimiter("a,b,c\n1,2,3\n4,5,6")).toBe(",");
    });

    it("should detect semicolon delimiter", () => {
      expect(detectDelimiter("a;b;c\n1;2;3\n4;5;6")).toBe(";");
    });

    it("should detect tab delimiter", () => {
      expect(detectDelimiter("a\tb\tc\n1\t2\t3\n4\t5\t6")).toBe("\t");
    });

    it("should detect pipe delimiter", () => {
      expect(detectDelimiter("a|b|c\n1|2|3\n4|5|6")).toBe("|");
    });

    it("should prefer comma when counts are equal", () => {
      expect(detectDelimiter("a\nb\nc")).toBe(",");
    });

    it("should handle empty input", () => {
      expect(detectDelimiter("")).toBe(",");
    });
  });

  describe("parseCsv with auto-detect delimiter", () => {
    it("should auto-detect comma delimiter when delimiter is empty string", () => {
      const input = "a,b,c\n1,2,3";
      const result = parseCsv(input, { delimiter: "" });
      expect(result).toEqual([
        ["a", "b", "c"],
        ["1", "2", "3"]
      ]);
    });

    it("should auto-detect semicolon delimiter when delimiter is empty string", () => {
      const input = "a;b;c\n1;2;3";
      const result = parseCsv(input, { delimiter: "" });
      expect(result).toEqual([
        ["a", "b", "c"],
        ["1", "2", "3"]
      ]);
    });

    it("should work with headers and auto-detect", () => {
      const input = "name;age;city\nAlice;30;NYC\nBob;25;LA";
      const result = parseCsv(input, { delimiter: "", headers: true });
      expect(result).toMatchObject({
        headers: ["name", "age", "city"],
        rows: [
          { name: "Alice", age: "30", city: "NYC" },
          { name: "Bob", age: "25", city: "LA" }
        ]
      });
      expect((result as any).meta?.delimiter).toBe(";");
    });
  });
});

// =============================================================================
// Header Rename Meta Tests
// =============================================================================
describe("Header Rename Meta (renamedHeaders)", () => {
  it("should expose renamedHeaders for header row duplicates", () => {
    const input = "A,A,A_1\n1,2,3";
    const result = parseCsv(input, { headers: true }) as any;

    expect(result.headers).toEqual(["A", "A_2", "A_1"]);
    expect(result.rows).toEqual([{ A: "1", A_2: "2", A_1: "3" }]);
    expect(result.meta?.renamedHeaders).toEqual({ A_2: "A" });
  });

  it("should expose renamedHeaders for explicit headers array", () => {
    const input = "1,2\n3,4";
    const result = parseCsv(input, { headers: ["A", "A"], delimiter: "," }) as any;

    expect(result.headers).toEqual(["A", "A_1"]);
    expect(result.rows).toEqual([
      { A: "1", A_1: "2" },
      { A: "3", A_1: "4" }
    ]);
    expect(result.meta?.renamedHeaders).toEqual({ A_1: "A" });
  });

  it("should expose renamedHeaders for header transform function", () => {
    const input = "a,b\n1,2";
    const result = parseCsv(input, {
      delimiter: ",",
      headers: () => ["X", "X"]
    }) as any;

    expect(result.headers).toEqual(["X", "X_1"]);
    expect(result.rows).toEqual([{ X: "1", X_1: "2" }]);
    expect(result.meta?.renamedHeaders).toEqual({ X_1: "X" });
  });
});

// =============================================================================
// relaxQuotes Option Tests
// =============================================================================
describe("relaxQuotes Option", () => {
  it("should allow unescaped quotes in unquoted field when relaxQuotes is true", () => {
    const csv = 'name,description\nTest,He said "hello" to them';
    const result = parseCsv(csv, { headers: true, relaxQuotes: true }) as any;

    expect(result.rows[0].description).toBe('He said "hello" to them');
  });

  it("should handle multiple quotes in field when relaxQuotes is true", () => {
    // When field starts with quote, relaxQuotes allows unescaped quotes inside
    // The outer quotes are consumed as field delimiters
    const csv = 'text\nShe said "A" and "B"';
    const result = parseCsv(csv, { headers: true, relaxQuotes: true }) as any;

    expect(result.rows[0].text).toBe('She said "A" and "B"');
  });

  it("should still parse properly quoted fields correctly", () => {
    const csv = 'name,value\n"hello, world",test';
    const result = parseCsv(csv, { headers: true, relaxQuotes: true }) as any;

    expect(result.rows[0].name).toBe("hello, world");
    expect(result.rows[0].value).toBe("test");
  });
});

// =============================================================================
// transform/validate Callback Tests
// =============================================================================
describe("rowTransform/validate Callbacks", () => {
  describe("rowTransform", () => {
    it("should transform rows", () => {
      const csv = "name,age\nAlice,30\nBob,25";
      const result = parseCsv(csv, {
        headers: true,
        rowTransform: row => {
          const r = row as Record<string, string>;
          return { ...r, age: String(parseInt(r.age) * 2) };
        }
      }) as any;

      expect(result.rows[0]).toEqual({ name: "Alice", age: "60" });
      expect(result.rows[1]).toEqual({ name: "Bob", age: "50" });
    });

    it("should filter rows by returning null", () => {
      const csv = "name,age\nAlice,30\nBob,25\nCharlie,35";
      const result = parseCsv(csv, {
        headers: true,
        rowTransform: row => {
          const r = row as Record<string, string>;
          return parseInt(r.age) > 25 ? row : null;
        }
      }) as any;

      expect(result.rows).toHaveLength(2);
      expect(result.rows[0].name).toBe("Alice");
      expect(result.rows[1].name).toBe("Charlie");
    });
  });

  describe("validate", () => {
    it("should filter rows by returning false", () => {
      const csv = "name,age\nAlice,30\nBob,invalid\nCharlie,35";
      const result = parseCsv(csv, {
        headers: true,
        validate: row => {
          const r = row as Record<string, string>;
          return !isNaN(parseInt(r.age));
        }
      }) as any;

      expect(result.rows).toHaveLength(2);
      expect(result.rows[0].name).toBe("Alice");
      expect(result.rows[1].name).toBe("Charlie");
    });
  });
});

// =============================================================================
// skipRecordsWithError/onSkip Tests
// =============================================================================
describe("skipRecordsWithError/onSkip", () => {
  it("should skip invalid rows and invoke onSkip", () => {
    const csv = "a,b,c\n1,2\n3,4,5\n6,7\n8,9,10";
    const skipped: { line: number; code: string }[] = [];

    const result = parseCsv(csv, {
      headers: true,
      skipRecordsWithError: true,
      onSkip: (error, _record) => {
        skipped.push({ line: error.line, code: error.code });
      }
    }) as any;

    expect(result.rows).toHaveLength(2);
    expect(skipped).toEqual([
      { line: 2, code: "TooFewFields" },
      { line: 4, code: "TooFewFields" }
    ]);
  });
});

// =============================================================================
// skipRecordsWithEmptyValues Tests
// =============================================================================
describe("skipRecordsWithEmptyValues Option", () => {
  it("should skip records where all values are empty strings", () => {
    const csv = "a,b,c\n1,2,3\n,,\n4,5,6";
    const result = parseCsv(csv, {
      headers: true,
      skipRecordsWithEmptyValues: true
    }) as any;

    expect(result.rows).toHaveLength(2);
    expect(result.rows[0]).toEqual({ a: "1", b: "2", c: "3" });
    expect(result.rows[1]).toEqual({ a: "4", b: "5", c: "6" });
  });

  it("should not skip records with at least one non-empty value", () => {
    const csv = "a,b,c\n,value,\n,,";
    const result = parseCsv(csv, {
      headers: true,
      skipRecordsWithEmptyValues: true
    }) as any;

    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]).toEqual({ a: "", b: "value", c: "" });
  });
});

// =============================================================================
// columnMismatch Option Tests
// =============================================================================
describe("columnMismatch Option", () => {
  it("should pad rows with too few columns", () => {
    const csv = "a,b,c\n1,2\n3,4,5";
    const result = parseCsv(csv, {
      headers: true,
      columnMismatch: { less: "pad", more: "error" }
    }) as any;

    expect(result.rows).toHaveLength(2);
    expect(result.rows[0]).toEqual({ a: "1", b: "2", c: "" });
    expect(result.rows[1]).toEqual({ a: "3", b: "4", c: "5" });
  });

  it("should truncate rows with too many columns", () => {
    const csv = "a,b\n1,2,3,4\n5,6";
    const result = parseCsv(csv, {
      headers: true,
      columnMismatch: { less: "error", more: "truncate" }
    }) as any;

    expect(result.rows).toHaveLength(2);
    expect(result.rows[0]).toEqual({ a: "1", b: "2" });
    expect(result.rows[1]).toEqual({ a: "5", b: "6" });
  });
});
