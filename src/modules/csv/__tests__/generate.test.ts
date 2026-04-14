/**
 * CSV Generate Unit Tests
 */

import { describe, it, expect } from "vitest";

import {
  csvGenerate,
  csvGenerateRows,
  csvGenerateData,
  createCsvGenerator,
  csvGenerateAsync
} from "../utils/generate";

describe("csvGenerate", () => {
  describe("basic generation", () => {
    it("should generate default CSV with 5 columns and 10 rows", () => {
      const { csv, headers, data } = csvGenerate();

      expect(headers).toHaveLength(5);
      expect(data).toHaveLength(10);
      expect(csv).toBeTruthy();
      expect(csv.split("\n")).toHaveLength(11); // 1 header + 10 data rows
    });

    it("should generate specified number of columns and rows", () => {
      const { headers, data } = csvGenerate({ columns: 3, rows: 5 });

      expect(headers).toHaveLength(3);
      expect(data).toHaveLength(5);
      data.forEach(row => expect(row).toHaveLength(3));
    });

    it("should generate reproducible data with seed", () => {
      const result1 = csvGenerate({ columns: 3, rows: 5, seed: 12345 });
      const result2 = csvGenerate({ columns: 3, rows: 5, seed: 12345 });

      expect(result1.csv).toBe(result2.csv);
      expect(result1.data).toEqual(result2.data);
    });

    it("should generate different data with different seeds", () => {
      const result1 = csvGenerate({ columns: 3, rows: 5, seed: 12345 });
      const result2 = csvGenerate({ columns: 3, rows: 5, seed: 54321 });

      expect(result1.csv).not.toBe(result2.csv);
    });
  });

  describe("column types", () => {
    it("should generate string columns", () => {
      const { data } = csvGenerate({ columns: ["string"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).length).toBeGreaterThan(0);
      });
    });

    it("should generate int columns", () => {
      const { data } = csvGenerate({ columns: ["int"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("number");
        expect(Number.isInteger(row[0])).toBe(true);
      });
    });

    it("should generate float columns", () => {
      const { data } = csvGenerate({ columns: ["float"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("number");
      });
    });

    it("should generate bool columns", () => {
      const { data } = csvGenerate({ columns: ["bool"], rows: 100, seed: 42 });
      const values = data.map(row => row[0]);

      expect(values.some(v => v === true)).toBe(true);
      expect(values.some(v => v === false)).toBe(true);
    });

    it("should generate date columns", () => {
      const { data } = csvGenerate({ columns: ["date"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect(row[0]).toMatch(/^\d{4}-\d{2}-\d{2}$/);
      });
    });

    it("should generate datetime columns", () => {
      const { data } = csvGenerate({ columns: ["datetime"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        // ISO format
        expect(new Date(row[0] as string).toISOString()).toBe(row[0]);
      });
    });

    it("should generate uuid columns", () => {
      const { data } = csvGenerate({ columns: ["uuid"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect(row[0]).toMatch(
          /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i
        );
      });
    });

    it("should generate email columns", () => {
      const { data } = csvGenerate({ columns: ["email"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect(row[0]).toMatch(/^[a-z]+@[a-z.]+$/);
      });
    });

    it("should generate name columns", () => {
      const { data } = csvGenerate({ columns: ["name"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).split(" ")).toHaveLength(2);
      });
    });

    it("should generate word columns", () => {
      const { data } = csvGenerate({ columns: ["word"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).length).toBeGreaterThan(0);
      });
    });

    it("should generate sentence columns", () => {
      const { data } = csvGenerate({ columns: ["sentence"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect(row[0]).toMatch(/^[A-Z].+\.$/);
      });
    });

    it("should generate index columns", () => {
      const { data } = csvGenerate({ columns: ["index"], rows: 10 });
      data.forEach((row, i) => {
        expect(row[0]).toBe(i);
      });
    });

    it("should generate firstName columns", () => {
      const { data } = csvGenerate({ columns: ["firstName"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).length).toBeGreaterThan(0);
      });
    });

    it("should generate lastName columns", () => {
      const { data } = csvGenerate({ columns: ["lastName"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).length).toBeGreaterThan(0);
      });
    });

    it("should generate paragraph columns", () => {
      const { data } = csvGenerate({ columns: ["paragraph"], rows: 5 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        // Multiple sentences
        expect((row[0] as string).split(". ").length).toBeGreaterThanOrEqual(2);
      });
    });

    it("should generate phone columns", () => {
      const { data } = csvGenerate({ columns: ["phone"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect(row[0]).toMatch(/^\+1-\d{3}-\d{3}-\d{4}$/);
      });
    });

    it("should generate url columns", () => {
      const { data } = csvGenerate({ columns: ["url"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect(row[0]).toMatch(/^https?:\/\//);
      });
    });

    it("should generate ip columns", () => {
      const { data } = csvGenerate({ columns: ["ip"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect(row[0]).toMatch(/^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$/);
      });
    });

    it("should generate ipv6 columns", () => {
      const { data } = csvGenerate({ columns: ["ipv6"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).split(":")).toHaveLength(8);
      });
    });

    it("should generate hex columns", () => {
      const { data } = csvGenerate({ columns: ["hex"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect(row[0]).toMatch(/^[0-9a-f]+$/);
      });
    });

    it("should generate company columns", () => {
      const { data } = csvGenerate({ columns: ["company"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).length).toBeGreaterThan(0);
      });
    });

    it("should generate country columns", () => {
      const { data } = csvGenerate({ columns: ["country"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).length).toBeGreaterThan(0);
      });
    });

    it("should generate currency columns", () => {
      const { data } = csvGenerate({ columns: ["currency"], rows: 10 });
      data.forEach(row => {
        expect(row[0]).toMatch(/^[A-Z]{3} \d+\.\d{2}$/);
      });
    });

    it("should generate percent columns", () => {
      const { data } = csvGenerate({ columns: ["percent"], rows: 10 });
      data.forEach(row => {
        expect(row[0]).toMatch(/^\d+(\.\d)?%$/);
      });
    });

    it("should generate timestamp columns", () => {
      const now = Date.now();
      const { data } = csvGenerate({ columns: ["timestamp"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("number");
        expect(row[0] as number).toBeGreaterThanOrEqual(now);
      });
    });

    it("should generate city columns", () => {
      const { data } = csvGenerate({ columns: ["city"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).length).toBeGreaterThan(0);
      });
    });

    it("should generate zipCode columns", () => {
      const { data } = csvGenerate({ columns: ["zipCode"], rows: 10 });
      data.forEach(row => {
        expect(row[0]).toMatch(/^\d{5}$/);
      });
    });

    it("should generate color columns", () => {
      const { data } = csvGenerate({ columns: ["color"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).length).toBeGreaterThan(0);
      });
    });

    it("should generate username columns", () => {
      const { data } = csvGenerate({ columns: ["username"], rows: 10 });
      data.forEach(row => {
        expect(row[0]).toMatch(/^[a-z]+_[a-z]+\d+$/);
      });
    });

    it("should generate slug columns", () => {
      const { data } = csvGenerate({ columns: ["slug"], rows: 10 });
      data.forEach(row => {
        expect(typeof row[0]).toBe("string");
        expect((row[0] as string).includes("-")).toBe(true);
      });
    });
  });

  describe("column config", () => {
    it("should respect min/max for int", () => {
      const { data } = csvGenerate({
        columns: [{ type: "int", min: 100, max: 200 }],
        rows: 100
      });
      data.forEach(row => {
        expect(row[0]).toBeGreaterThanOrEqual(100);
        expect(row[0]).toBeLessThanOrEqual(200);
      });
    });

    it("should respect min/max for float", () => {
      const { data } = csvGenerate({
        columns: [{ type: "float", min: 10, max: 20 }],
        rows: 100
      });
      data.forEach(row => {
        expect(row[0]).toBeGreaterThanOrEqual(10);
        expect(row[0]).toBeLessThanOrEqual(20);
      });
    });

    it("should respect length for string", () => {
      const { data } = csvGenerate({
        columns: [{ type: "string", length: 5 }],
        rows: 10
      });
      data.forEach(row => {
        expect((row[0] as string).length).toBe(5);
      });
    });

    it("should pick from values array", () => {
      const { data } = csvGenerate({
        columns: [{ type: "string", values: ["A", "B", "C"] }],
        rows: 100
      });
      data.forEach(row => {
        expect(["A", "B", "C"]).toContain(row[0]);
      });
    });

    it("should generate nullable values", () => {
      const { data } = csvGenerate({
        columns: [{ type: "string", nullable: 0.5 }],
        rows: 100,
        seed: 42
      });
      const nullCount = data.filter(row => row[0] === null).length;
      expect(nullCount).toBeGreaterThan(0);
      expect(nullCount).toBeLessThan(100);
    });

    it("should use column name for headers", () => {
      const { headers } = csvGenerate({
        columns: [
          { type: "string", name: "FirstName" },
          { type: "int", name: "Age" }
        ],
        rows: 1
      });
      expect(headers).toEqual(["FirstName", "Age"]);
    });
  });

  describe("custom generator function", () => {
    it("should support custom generator function", () => {
      const { data } = csvGenerate({
        columns: [ctx => `row-${ctx.rowIndex}-col-${ctx.colIndex}`],
        rows: 5
      });
      expect(data[0][0]).toBe("row-0-col-0");
      expect(data[4][0]).toBe("row-4-col-0");
    });

    it("should provide random utilities in context", () => {
      const { data } = csvGenerate({
        columns: [ctx => ctx.randomInt(1, 10)],
        rows: 100
      });
      data.forEach(row => {
        expect(row[0]).toBeGreaterThanOrEqual(1);
        expect(row[0]).toBeLessThanOrEqual(10);
      });
    });

    it("should provide randomPick in context", () => {
      const options = ["X", "Y", "Z"];
      const { data } = csvGenerate({
        columns: [ctx => ctx.randomPick(options)],
        rows: 100
      });
      data.forEach(row => {
        expect(options).toContain(row[0]);
      });
    });
  });

  describe("headers", () => {
    it("should auto-generate headers by default", () => {
      const { headers, csv } = csvGenerate({
        columns: ["string", "int", "bool"],
        rows: 1
      });
      expect(headers).toEqual(["string_1", "int_2", "bool_3"]);
      expect(csv.split("\n")[0]).toBe("string_1,int_2,bool_3");
    });

    it("should use custom headers", () => {
      const { headers, csv } = csvGenerate({
        columns: ["string", "int"],
        rows: 1,
        headers: ["Name", "Age"]
      });
      expect(headers).toEqual(["Name", "Age"]);
      expect(csv.split("\n")[0]).toBe("Name,Age");
    });

    it("should skip headers when headers: false", () => {
      const { csv } = csvGenerate({
        columns: 3,
        rows: 2,
        headers: false
      });
      expect(csv.split("\n")).toHaveLength(2);
    });
  });

  describe("delimiters", () => {
    it("should use custom field delimiter", () => {
      const { csv } = csvGenerate({
        columns: 3,
        rows: 1,
        delimiter: ";"
      });
      expect(csv.split("\n")[0]).toMatch(/;.*;/);
    });

    it("should use custom row delimiter", () => {
      const { csv } = csvGenerate({
        columns: 2,
        rows: 2,
        lineEnding: "\r\n"
      });
      expect(csv.split("\r\n")).toHaveLength(3);
    });
  });

  describe("CSV escaping", () => {
    it("should quote fields containing delimiter", () => {
      const { csv } = csvGenerate({
        columns: [ctx => "hello,world"],
        rows: 1,
        headers: false
      });
      expect(csv).toBe('"hello,world"');
    });

    it("should escape double quotes", () => {
      const { csv } = csvGenerate({
        columns: [ctx => 'say "hello"'],
        rows: 1,
        headers: false
      });
      expect(csv).toBe('"say ""hello"""');
    });

    it("should quote fields containing newlines", () => {
      const { csv } = csvGenerate({
        columns: [ctx => "line1\nline2"],
        rows: 1,
        headers: false
      });
      expect(csv).toBe('"line1\nline2"');
    });
  });
});

describe("csvGenerateRows", () => {
  it("should generate rows as iterator", () => {
    const rows = [...csvGenerateRows({ columns: 3, rows: 5 })];
    expect(rows).toHaveLength(6); // 1 header + 5 data rows
  });

  it("should be memory efficient for large datasets", () => {
    let count = 0;
    for (const _ of csvGenerateRows({ columns: 3, rows: 10000 })) {
      count++;
      if (count > 100) {
        break;
      } // Only consume first 100
    }
    expect(count).toBe(101);
  });
});

describe("csvGenerateData", () => {
  it("should return raw data without CSV formatting", () => {
    const data = csvGenerateData({
      columns: ["name", "int", "bool"],
      rows: 5
    });
    expect(data).toHaveLength(5);
    data.forEach(row => {
      expect(row).toHaveLength(3);
      expect(typeof row[0]).toBe("string");
      expect(typeof row[1]).toBe("number");
      expect(typeof row[2]).toBe("boolean");
    });
  });
});

describe("createCsvGenerator", () => {
  it("should create reusable generator", () => {
    const gen = createCsvGenerator({
      columns: ["name", "int"],
      seed: 42
    });

    const batch1 = gen.generate(5);
    const batch2 = gen.data(5);

    expect(batch1.data).toHaveLength(5);
    expect(batch2).toHaveLength(5);
  });

  it("should support rows generator", () => {
    const gen = createCsvGenerator({ columns: 3 });
    const rows = [...gen.rows(5)];
    expect(rows).toHaveLength(6); // header + 5 rows
  });
});

describe("csvGenerateAsync", () => {
  it("should generate rows asynchronously", async () => {
    const rows: string[] = [];
    for await (const row of csvGenerateAsync({ columns: 3, rows: 5 })) {
      rows.push(row);
    }
    expect(rows).toHaveLength(6);
  });

  it("should support delay between rows", async () => {
    const start = Date.now();
    const rows: string[] = [];
    for await (const row of csvGenerateAsync({ columns: 2, rows: 3, delay: 10 })) {
      rows.push(row);
    }
    const elapsed = Date.now() - start;
    expect(elapsed).toBeGreaterThanOrEqual(30); // At least 3 delays of 10ms
  });

  it("should support duration-based stopping", async () => {
    const rows: string[] = [];
    const start = Date.now();
    for await (const row of csvGenerateAsync({ columns: 2, duration: 50, delay: 5 })) {
      rows.push(row);
    }
    const elapsed = Date.now() - start;
    expect(elapsed).toBeGreaterThanOrEqual(50);
    expect(elapsed).toBeLessThan(200); // Should stop reasonably soon after duration
  });
});

describe("stop conditions", () => {
  it("should support unlimited rows with manual break", () => {
    let count = 0;
    for (const _ of csvGenerateRows({ columns: 2, rows: -1 })) {
      count++;
      if (count >= 100) {
        break;
      }
    }
    expect(count).toBe(100);
  });

  it("should support Infinity for unlimited", () => {
    let count = 0;
    for (const _ of csvGenerateRows({ columns: 2, rows: Infinity })) {
      count++;
      if (count >= 50) {
        break;
      }
    }
    expect(count).toBe(50);
  });

  it("should throw error for unlimited in sync csvGenerate", () => {
    expect(() => csvGenerate({ columns: 2, rows: -1 })).toThrow(/Unlimited generation/);
    expect(() => csvGenerate({ columns: 2, rows: Infinity })).toThrow(/Unlimited generation/);
  });

  it("should support duration-based stopping", () => {
    const start = Date.now();
    let count = 0;
    for (const _ of csvGenerateRows({ columns: 2, duration: 50 })) {
      count++;
    }
    const elapsed = Date.now() - start;
    expect(elapsed).toBeGreaterThanOrEqual(50);
    expect(count).toBeGreaterThan(0);
  });

  it("should support until-based stopping with Date", () => {
    const until = new Date(Date.now() + 30);
    let count = 0;
    for (const _ of csvGenerateRows({ columns: 2, until })) {
      count++;
    }
    expect(count).toBeGreaterThan(0);
  });

  it("should support until-based stopping with timestamp", () => {
    const until = Date.now() + 30;
    let count = 0;
    for (const _ of csvGenerateRows({ columns: 2, until })) {
      count++;
    }
    expect(count).toBeGreaterThan(0);
  });
});

describe("eof option", () => {
  it("should append eof string to output", () => {
    const { csv } = csvGenerate({ columns: 2, rows: 2, eof: "\n" });
    expect(csv.endsWith("\n")).toBe(true);
  });

  it("should append custom eof string", () => {
    const { csv } = csvGenerate({ columns: 2, rows: 2, eof: "\r\n---END---\r\n" });
    expect(csv.endsWith("\r\n---END---\r\n")).toBe(true);
  });

  it("should not append eof if not specified", () => {
    const { csv } = csvGenerate({ columns: 2, rows: 2 });
    expect(csv.endsWith("\n")).toBe(false);
  });
});

describe("bom option", () => {
  it("should prepend UTF-8 BOM when bom: true", () => {
    const { csv } = csvGenerate({ columns: 2, rows: 2, bom: true });
    expect(csv.charCodeAt(0)).toBe(0xfeff);
  });

  it("should not prepend BOM when bom is not specified", () => {
    const { csv } = csvGenerate({ columns: 2, rows: 2 });
    expect(csv.charCodeAt(0)).not.toBe(0xfeff);
  });

  it("should work with eof together", () => {
    const { csv } = csvGenerate({ columns: 2, rows: 2, bom: true, eof: "\n" });
    expect(csv.charCodeAt(0)).toBe(0xfeff);
    expect(csv.endsWith("\n")).toBe(true);
  });
});

describe("quote option", () => {
  it("should always quote with quote: 'always'", () => {
    const { csv } = csvGenerate({ columns: ["string"], rows: 2, quote: "always", seed: 42 });
    const lines = csv.split("\n");
    lines.forEach(line => {
      expect(line.startsWith('"')).toBe(true);
      expect(line.endsWith('"')).toBe(true);
    });
  });

  it("should never quote with quote: 'never'", () => {
    const { csv } = csvGenerate({
      columns: [() => "hello,world"],
      rows: 2,
      quote: "never"
    });
    // Contains unquoted comma
    expect(csv.includes("hello,world")).toBe(true);
  });

  it("should auto quote by default (quote: 'auto')", () => {
    const { csv } = csvGenerate({
      columns: [() => "hello,world"],
      rows: 1,
      headers: false
    });
    // Should be quoted because it contains comma
    expect(csv).toBe('"hello,world"');
  });
});

describe("skipRows option", () => {
  it("should skip first N rows", () => {
    const { data } = csvGenerate({
      columns: ["index"],
      rows: 5,
      skipRows: 3,
      seed: 42
    });
    // Should have 5 rows but starting from index 3
    expect(data).toHaveLength(5);
    expect(data[0][0]).toBe(3);
    expect(data[4][0]).toBe(7);
  });

  it("should work with iterator API", () => {
    const rows: string[] = [];
    for (const row of csvGenerateRows({
      columns: ["index"],
      rows: 3,
      skipRows: 2,
      headers: false
    })) {
      rows.push(row);
    }
    expect(rows).toHaveLength(3);
    expect(rows[0]).toBe("2");
    expect(rows[2]).toBe("4");
  });
});

describe("enhanced context", () => {
  it("should provide randomFloat in context", () => {
    const { data } = csvGenerate({
      columns: [ctx => ctx.randomFloat(10, 20, 2)],
      rows: 100
    });
    data.forEach(row => {
      expect(row[0]).toBeGreaterThanOrEqual(10);
      expect(row[0]).toBeLessThanOrEqual(20);
    });
  });

  it("should provide randomDate in context", () => {
    const { data } = csvGenerate({
      columns: [ctx => ctx.randomDate().toISOString()],
      rows: 10
    });
    data.forEach(row => {
      expect(new Date(row[0] as string)).toBeInstanceOf(Date);
    });
  });

  it("should provide randomBool in context", () => {
    const { data } = csvGenerate({
      columns: [ctx => ctx.randomBool(0.8)],
      rows: 100,
      seed: 42
    });
    const trueCount = data.filter(row => row[0] === true).length;
    // With 80% probability, expect roughly 80 trues (with some variance)
    expect(trueCount).toBeGreaterThan(60);
    expect(trueCount).toBeLessThan(95);
  });
});

describe("objectMode option", () => {
  it("should return objects when objectMode: true", () => {
    const data = csvGenerateData({
      columns: [
        { type: "name", name: "fullName" },
        { type: "int", name: "age" },
        { type: "email", name: "email" }
      ],
      rows: 5,
      objectMode: true
    }) as Record<string, unknown>[];

    expect(data).toHaveLength(5);
    data.forEach(row => {
      expect(row).toHaveProperty("fullName");
      expect(row).toHaveProperty("age");
      expect(row).toHaveProperty("email");
      expect(typeof row.fullName).toBe("string");
      expect(typeof row.age).toBe("number");
      expect(typeof row.email).toBe("string");
    });
  });

  it("should return arrays when objectMode is not specified", () => {
    const data = csvGenerateData({ columns: 3, rows: 5 });
    expect(Array.isArray(data[0])).toBe(true);
  });
});

describe("transform option", () => {
  it("should transform each row", () => {
    const { data } = csvGenerate({
      columns: [{ type: "int", min: 1, max: 10 }],
      rows: 5,
      transform: row => [(row[0] as number) * 10]
    });

    data.forEach(row => {
      expect(row[0] as number).toBeGreaterThanOrEqual(10);
      expect(row[0] as number).toBeLessThanOrEqual(100);
    });
  });

  it("should receive row context in transform", () => {
    const rowIndices: number[] = [];
    csvGenerate({
      columns: 2,
      rows: 5,
      transform: (row, ctx) => {
        rowIndices.push(ctx.rowIndex);
        return row;
      }
    });

    expect(rowIndices).toEqual([0, 1, 2, 3, 4]);
  });

  it("should work with iterator APIs", () => {
    let count = 0;
    for (const row of csvGenerateRows({
      columns: [{ type: "int", min: 1, max: 10 }],
      rows: 3,
      transform: row => [(row[0] as number) * 100]
    })) {
      if (count > 0) {
        // Skip header
        const val = parseInt(row);
        expect(val).toBeGreaterThanOrEqual(100);
      }
      count++;
    }
  });
});

describe("mixed column types", () => {
  it("should handle mixed column definitions", () => {
    const { data, headers } = csvGenerate({
      columns: [
        "name",
        { type: "int", min: 18, max: 65, name: "age" },
        "email",
        ctx => `ID-${ctx.rowIndex.toString().padStart(4, "0")}`,
        { type: "bool", name: "active" }
      ],
      rows: 10,
      seed: 42
    });

    expect(headers).toEqual(["name_1", "age", "email_3", "col_4", "active"]);
    expect(data).toHaveLength(10);

    data.forEach((row, i) => {
      expect(typeof row[0]).toBe("string"); // name
      expect(row[1]).toBeGreaterThanOrEqual(18); // age
      expect(row[1]).toBeLessThanOrEqual(65);
      expect(row[2] as string).toMatch(/@/); // email
      expect(row[3]).toBe(`ID-${i.toString().padStart(4, "0")}`); // custom
      expect(typeof row[4]).toBe("boolean"); // active
    });
  });
});

describe("createCsvGenerator enhanced API", () => {
  it("should accept options object for rows()", () => {
    const gen = createCsvGenerator({ columns: 2 });
    let count = 0;
    for (const _ of gen.rows({ rows: 5 })) {
      count++;
    }
    expect(count).toBe(6); // 1 header + 5 data rows
  });

  it("should accept options object for generate()", () => {
    const gen = createCsvGenerator({ columns: 2 });
    const result = gen.generate({ rows: 3, eof: "\n" });
    expect(result.data).toHaveLength(3);
    expect(result.csv.endsWith("\n")).toBe(true);
  });

  it("should support asyncRows()", async () => {
    const gen = createCsvGenerator({ columns: 2 });
    const rows: string[] = [];
    for await (const row of gen.asyncRows({ rows: 3, delay: 5 })) {
      rows.push(row);
    }
    expect(rows).toHaveLength(4); // 1 header + 3 data rows
  });
});
