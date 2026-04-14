/**
 * CSV Test Utilities
 *
 * Shared helpers and utilities for CSV module tests.
 * Provides consistent test data generation, assertions, and helpers.
 */

import { parseCsv, formatCsv } from "@csv/index";
import type { CsvParseResult } from "@csv/index";

// =============================================================================
// Test Data Generation
// =============================================================================

/**
 * Generate a large CSV string for performance testing
 */
export function generateLargeCsv(rows: number, cols: number, options?: { seed?: number }): string {
  const lines: string[] = [];

  // Header row
  const headers = Array.from({ length: cols }, (_, i) => `col${i + 1}`);
  lines.push(headers.join(","));

  // Data rows
  const seed = options?.seed ?? 42;
  let random = seed;
  const nextRandom = () => {
    random = (random * 1103515245 + 12345) & 0x7fffffff;
    return random;
  };

  for (let r = 0; r < rows; r++) {
    const row = Array.from({ length: cols }, () => {
      const type = nextRandom() % 4;
      switch (type) {
        case 0:
          return String(nextRandom() % 10000);
        case 1:
          return `text_${nextRandom() % 1000}`;
        case 2:
          return (nextRandom() / 100).toFixed(2);
        default:
          return "";
      }
    });
    lines.push(row.join(","));
  }

  return lines.join("\n");
}

/**
 * Generate a CSV with specific characteristics for edge case testing
 */
export function generateEdgeCaseCsv(
  type:
    | "wide" // Many columns
    | "deep" // Many rows
    | "large-field" // Single field with lots of data
    | "unicode-heavy" // Full of unicode characters
    | "quoted-heavy" // Many quoted fields with special chars
): string {
  switch (type) {
    case "wide": {
      // 1000 columns
      const cols = 1000;
      const headers = Array.from({ length: cols }, (_, i) => `h${i}`);
      const row = Array.from({ length: cols }, (_, i) => `v${i}`);
      return headers.join(",") + "\n" + row.join(",");
    }
    case "deep": {
      // 10000 rows
      const rows = 10000;
      const lines = ["a,b,c"];
      for (let i = 0; i < rows; i++) {
        lines.push(`${i},${i * 2},${i * 3}`);
      }
      return lines.join("\n");
    }
    case "large-field": {
      // Single field with 100KB of data
      const largeContent = "x".repeat(100 * 1024);
      return `name,content\ntest,"${largeContent}"`;
    }
    case "unicode-heavy": {
      const unicodeChars = [
        "你好",
        "世界",
        "🎉",
        "🌍",
        "こんにちは",
        "مرحبا",
        "שלום",
        "Привет",
        "🚀",
        "💻"
      ];
      const rows = unicodeChars.map((char, i) => `${i},${char},test${i}`);
      return "id,unicode,name\n" + rows.join("\n");
    }
    case "quoted-heavy": {
      return [
        "a,b,c",
        '"hello, world","line1\nline2","test"',
        '"say ""hi""","a\r\nb\r\nc","quoted"',
        '"comma,here","newline\nhere","quote""here"'
      ].join("\n");
    }
    default:
      return "a,b,c\n1,2,3";
  }
}

/**
 * Generate malformed CSV data for error handling tests
 */
export function generateMalformedCsv(
  type:
    | "unclosed-quote" // Quote never closed
    | "mismatched-columns" // Inconsistent column counts
    | "binary-data" // Contains binary characters
    | "truncated" // Ends mid-field
): string {
  switch (type) {
    case "unclosed-quote":
      return 'a,b,c\n"unclosed,value\n1,2,3';
    case "mismatched-columns":
      return "a,b,c\n1,2\n3,4,5,6\n7,8,9";
    case "binary-data":
      // Contains null bytes and other binary chars
      return `a,b\ntest\x00data,value\n1,2`;
    case "truncated":
      return 'a,b,c\n1,"incomplete';
    default:
      return "a,b,c\n1,2,3";
  }
}

// =============================================================================
// Test Assertions
// =============================================================================

/**
 * Assert that two CSV strings are equivalent (ignoring trailing newlines)
 */
export function assertCsvEqual(actual: string, expected: string): void {
  const normalizedActual = actual.replace(/\r\n/g, "\n").replace(/\n+$/, "");
  const normalizedExpected = expected.replace(/\r\n/g, "\n").replace(/\n+$/, "");

  if (normalizedActual !== normalizedExpected) {
    throw new Error(
      `CSV mismatch:\nActual:\n${JSON.stringify(normalizedActual)}\n\nExpected:\n${JSON.stringify(normalizedExpected)}`
    );
  }
}

/**
 * Assert that parsing and formatting a CSV produces the same result (roundtrip)
 */
export function assertCsvRoundtrip(csv: string, options?: { delimiter?: string }): void {
  const parsed = parseCsv(csv, { delimiter: options?.delimiter }) as string[][];
  const formatted = formatCsv(parsed, {
    delimiter: options?.delimiter,
    trailingNewline: false
  });

  assertCsvEqual(formatted, csv);
}

/**
 * Assert that a parse result has expected structure
 */
export function assertParseResult<T>(
  result: CsvParseResult<T>,
  expectations: {
    rowCount?: number;
    headers?: string[];
    truncated?: boolean;
    delimiter?: string;
  }
): void {
  if (expectations.rowCount !== undefined) {
    if (result.rows.length !== expectations.rowCount) {
      throw new Error(`Expected ${expectations.rowCount} rows, got ${result.rows.length}`);
    }
  }
  if (expectations.headers !== undefined) {
    const actualHeaders = JSON.stringify(result.headers);
    const expectedHeaders = JSON.stringify(expectations.headers);
    if (actualHeaders !== expectedHeaders) {
      throw new Error(`Headers mismatch:\nActual: ${actualHeaders}\nExpected: ${expectedHeaders}`);
    }
  }
  if (expectations.truncated !== undefined && result.meta.truncated !== expectations.truncated) {
    throw new Error(`Expected truncated=${expectations.truncated}, got ${result.meta.truncated}`);
  }
  if (expectations.delimiter !== undefined && result.meta.delimiter !== expectations.delimiter) {
    throw new Error(`Expected delimiter=${expectations.delimiter}, got ${result.meta.delimiter}`);
  }
}

// =============================================================================
// Stream Test Helpers
// =============================================================================

import { Readable } from "node:stream";

import type { CsvParserStream } from "@csv/stream";

/**
 * Collect all rows from a CsvParserStream
 */
export async function collectStreamRows<T = unknown>(parser: CsvParserStream): Promise<T[]> {
  const rows: T[] = [];
  return new Promise((resolve, reject) => {
    parser.on("data", row => rows.push(row as T));
    parser.on("end", () => resolve(rows));
    parser.on("error", reject);
  });
}

/**
 * Pipe CSV string to parser and collect results
 */
export async function parseStreamCsv<T = unknown>(
  csv: string,
  parser: CsvParserStream
): Promise<T[]> {
  const readable = Readable.from(csv);
  readable.pipe(parser);
  return collectStreamRows<T>(parser);
}

/**
 * Create a slow readable stream for testing backpressure
 */
export function createSlowReadable(chunks: string[], delayMs: number): Readable {
  let index = 0;
  return new Readable({
    read() {
      if (index >= chunks.length) {
        this.push(null);
        return;
      }
      setTimeout(() => {
        this.push(chunks[index++]);
      }, delayMs);
    }
  });
}

// =============================================================================
// Performance Helpers
// =============================================================================

/**
 * Measure execution time of a function
 */
export async function measureTime<T>(fn: () => T | Promise<T>): Promise<{ result: T; ms: number }> {
  const start = performance.now();
  const result = await fn();
  const ms = performance.now() - start;
  return { result, ms };
}

/**
 * Run a function multiple times and return average time
 */
export async function benchmarkAverage<T>(
  fn: () => T | Promise<T>,
  iterations: number
): Promise<{ avgMs: number; minMs: number; maxMs: number }> {
  const times: number[] = [];

  for (let i = 0; i < iterations; i++) {
    const { ms } = await measureTime(fn);
    times.push(ms);
  }

  return {
    avgMs: times.reduce((a, b) => a + b, 0) / times.length,
    minMs: Math.min(...times),
    maxMs: Math.max(...times)
  };
}

// =============================================================================
// Common Test Data
// =============================================================================

/**
 * Standard test CSVs for consistent testing across files
 */
export const TEST_CSV = {
  simple: "a,b,c\n1,2,3\n4,5,6",
  withHeaders: "name,age,city\nAlice,30,NYC\nBob,25,LA",
  quoted: '"hello, world",test\n"line1\nline2",value',
  escaped: '"say ""hello""",normal\n"quote""here",test',
  unicode: "姓名,年龄\n张三,30\n李四,25",
  emoji: "emoji,text\n😀,happy\n😢,sad",
  empty: "",
  singleField: "value",
  singleRow: "a,b,c",
  withBom: "\ufeffa,b,c\n1,2,3",
  crlf: "a,b\r\n1,2\r\n3,4",
  mixedLineEndings: "a,b\n1,2\r\n3,4\r5,6"
} as const;

/**
 * Standard malformed test CSVs
 */
export const MALFORMED_CSV = {
  unclosedQuote: 'a,b\n"unclosed,value',
  mismatchedColumns: "a,b,c\n1,2\n3,4,5,6",
  emptyQuotes: '"",,""',
  onlyNewlines: "\n\n\n"
} as const;
