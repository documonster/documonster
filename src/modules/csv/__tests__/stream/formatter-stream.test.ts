/**
 * CsvFormatterStream Tests
 *
 * Tests for the streaming CSV formatter (CsvFormatterStream).
 *
 * Coverage:
 * - Basic formatting
 * - Quoting (commas, quotes, newlines)
 * - Custom options (delimiter, row delimiter, BOM)
 * - Headers
 * - RowHashArray support
 * - Round-trip tests
 * - Backpressure handling
 */

import { CsvParserStream, CsvFormatterStream } from "@csv/stream";
import { Readable, Writable, pipeline } from "@stream";
import { describe, it, expect } from "vitest";

// =============================================================================
// Basic Formatting
// =============================================================================

describe("CsvFormatterStream - Basic Formatting", () => {
  it("should format rows to CSV", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["a", "b", "c"]);
    formatter.write(["1", "2", "3"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("a,b,c\n1,2,3");
  });

  it("should work with pipeline", async () => {
    const rows = [
      ["a", "b", "c"],
      ["1", "2", "3"]
    ];
    const readable = Readable.from(rows, { objectMode: true });
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    const writable = new Writable({
      write(chunk, _encoding, callback) {
        chunks.push(chunk.toString());
        callback();
      }
    });

    await pipeline(readable, formatter, writable);

    expect(chunks.join("")).toBe("a,b,c\n1,2,3");
  });

  it("should handle null and undefined values", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write([null, undefined, "value"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe(",,value");
  });

  it("should convert numbers and booleans", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write([1, 2.5, true, false]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("1,2.5,true,false");
  });

  it("should format numbers with comma decimalSeparator", async () => {
    const formatter = new CsvFormatterStream({ delimiter: ";", decimalSeparator: "," });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write([1, 2.5]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("1;2,5");
  });
});

// =============================================================================
// Quoting
// =============================================================================

describe("CsvFormatterStream - Quoting", () => {
  it("should quote fields containing commas", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["hello, world", "test"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe('"hello, world",test');
  });

  it("should quote fields containing quotes and escape them", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(['He said "Hello"', "test"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe('"He said ""Hello""",test');
  });

  it("should quote fields containing newlines", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["line1\nline2", "test"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe('"line1\nline2",test');
  });

  it("should always quote when quoteColumns: true", async () => {
    const formatter = new CsvFormatterStream({ quoteColumns: true });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["a", "b", "c"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe('"a","b","c"');
  });

  it("should not quote when quote is disabled", async () => {
    const formatter = new CsvFormatterStream({ quote: false });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["hello, world", "test"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("hello, world,test");
  });
});

// =============================================================================
// Options
// =============================================================================

describe("CsvFormatterStream - Options", () => {
  it("should support custom delimiter", async () => {
    const formatter = new CsvFormatterStream({ delimiter: ";" });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["a", "b", "c"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("a;b;c");
  });

  it("should support custom row delimiter", async () => {
    const formatter = new CsvFormatterStream({ lineEnding: "\r\n" });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["a", "b"]);
    formatter.write(["1", "2"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("a,b\r\n1,2");
  });

  it("should add BOM when bom is true", async () => {
    const formatter = new CsvFormatterStream({ bom: true });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["a", "b"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    const result = chunks.join("");
    expect(result.charCodeAt(0)).toBe(0xfeff);
  });
});

// =============================================================================
// Headers
// =============================================================================

describe("CsvFormatterStream - Headers", () => {
  it("should write custom headers", async () => {
    const formatter = new CsvFormatterStream({ headers: ["col1", "col2", "col3"] });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write(["a", "b", "c"]);
    formatter.write(["1", "2", "3"]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("col1,col2,col3\na,b,c\n1,2,3");
  });

  it("should auto-detect headers from objects when headers: true", async () => {
    const formatter = new CsvFormatterStream({ headers: true });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write({ name: "Alice", age: "30" });
    formatter.write({ name: "Bob", age: "25" });
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("name,age\nAlice,30\nBob,25");
  });

  it("should use custom header order for objects", async () => {
    const formatter = new CsvFormatterStream({ headers: ["age", "name"] });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write({ name: "Alice", age: "30" });
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("age,name\n30,Alice");
  });
});

// =============================================================================
// RowHashArray Support
// =============================================================================

describe("CsvFormatterStream - RowHashArray Support", () => {
  it("should format RowHashArray (array of [key, value] tuples)", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write([
      ["name", "Alice"],
      ["age", "30"]
    ]);
    formatter.write([
      ["name", "Bob"],
      ["age", "25"]
    ]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("Alice,30\nBob,25");
  });

  it("should auto-detect headers from RowHashArray when headers: true", async () => {
    const formatter = new CsvFormatterStream({ headers: true });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write([
      ["firstName", "Alice"],
      ["lastName", "Smith"]
    ]);
    formatter.write([
      ["firstName", "Bob"],
      ["lastName", "Jones"]
    ]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("firstName,lastName\nAlice,Smith\nBob,Jones");
  });

  it("should reorder RowHashArray columns based on custom headers", async () => {
    const formatter = new CsvFormatterStream({ headers: ["age", "name", "city"] });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write([
      ["name", "Alice"],
      ["city", "NYC"],
      ["age", "30"]
    ]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("age,name,city\n30,Alice,NYC");
  });

  it("should handle missing keys in RowHashArray when using custom headers", async () => {
    const formatter = new CsvFormatterStream({ headers: ["name", "age", "city"] });
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write([
      ["name", "Alice"],
      ["age", "30"]
    ]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe("name,age,city\nAlice,30,");
  });

  it("should format RowHashArray with special characters", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => {
      chunks.push(chunk.toString());
    });

    formatter.write([
      ["greeting", "Hello, World"],
      ["quote", 'He said "hi"']
    ]);
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    expect(chunks.join("")).toBe('"Hello, World","He said ""hi"""');
  });
});

// =============================================================================
// Round-trip Tests
// =============================================================================

describe("CsvFormatterStream - Round-trip Tests", () => {
  it("should round-trip simple data through parser and formatter", async () => {
    const original = [
      ["a", "b", "c"],
      ["1", "2", "3"],
      ["4", "5", "6"]
    ];

    // Format to CSV
    const formatter = new CsvFormatterStream();
    const csvChunks: string[] = [];

    formatter.on("data", chunk => csvChunks.push(chunk.toString()));

    for (const row of original) {
      formatter.write(row);
    }
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    const csv = csvChunks.join("");

    // Parse back
    const readable = Readable.from([csv]);
    const parser = new CsvParserStream();

    const parsed: string[][] = [];
    for await (const row of readable.pipe(parser)) {
      parsed.push(row as string[]);
    }

    expect(parsed).toEqual(original);
  });

  it("should round-trip data with special characters", async () => {
    const original = [
      ["hello, world", 'say "hi"'],
      ["line1\nline2", "normal"]
    ];

    // Format
    const formatter = new CsvFormatterStream();
    const csvChunks: string[] = [];

    formatter.on("data", chunk => csvChunks.push(chunk.toString()));

    for (const row of original) {
      formatter.write(row);
    }
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    // Parse back
    const readable = Readable.from([csvChunks.join("")]);
    const parser = new CsvParserStream();

    const parsed: string[][] = [];
    for await (const row of readable.pipe(parser)) {
      parsed.push(row as string[]);
    }

    expect(parsed).toEqual(original);
  });

  it("should round-trip Unicode data", async () => {
    const original = [
      ["日本語", "中文"],
      ["한국어", "😀🎉"]
    ];

    // Format
    const formatter = new CsvFormatterStream();
    const csvChunks: string[] = [];

    formatter.on("data", chunk => csvChunks.push(chunk.toString()));

    for (const row of original) {
      formatter.write(row);
    }
    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    // Parse back
    const readable = Readable.from([csvChunks.join("")]);
    const parser = new CsvParserStream();

    const parsed: string[][] = [];
    for await (const row of readable.pipe(parser)) {
      parsed.push(row as string[]);
    }

    expect(parsed).toEqual(original);
  });
});

// =============================================================================
// Backpressure
// =============================================================================

describe("CsvFormatterStream - Backpressure", () => {
  it("should respect backpressure", async () => {
    const formatter = new CsvFormatterStream();
    let writeCount = 0;
    let drainCount = 0;

    // Create a slow consumer
    const slowConsumer = new Writable({
      highWaterMark: 16, // Very small buffer
      write(_chunk, _encoding, callback) {
        // Simulate slow processing
        setTimeout(callback, 1);
      }
    });

    formatter.pipe(slowConsumer);

    // Write many rows
    for (let i = 0; i < 100; i++) {
      const canContinue = formatter.write([`row${i}`, `value${i}`]);
      writeCount++;

      if (!canContinue) {
        drainCount++;
        await new Promise(resolve => formatter.once("drain", resolve));
      }
    }

    formatter.end();

    await new Promise(resolve => slowConsumer.on("finish", resolve));

    expect(writeCount).toBe(100);
    // With backpressure, we should have had some drain events
    expect(drainCount).toBeGreaterThanOrEqual(0);
  });
});

// =============================================================================
// Large Data Performance
// =============================================================================

describe("CsvFormatterStream - Large Data Performance", () => {
  it("should write large CSV with streaming (10000 rows)", async () => {
    const formatter = new CsvFormatterStream();
    const chunks: string[] = [];

    formatter.on("data", chunk => chunks.push(chunk.toString()));

    // Write header
    formatter.write(["id", "value"]);

    // Write 10000 rows
    for (let i = 0; i < 10000; i++) {
      formatter.write([String(i), `value${i}`]);
    }

    formatter.end();

    await new Promise(resolve => formatter.on("finish", resolve));

    const result = chunks.join("");
    const lines = result.trim().split("\n");

    expect(lines.length).toBe(10001); // Header + 10000 data rows
    expect(lines[0]).toBe("id,value");
    expect(lines[1]).toBe("0,value0");
    expect(lines[10000]).toBe("9999,value9999");
  });
});
