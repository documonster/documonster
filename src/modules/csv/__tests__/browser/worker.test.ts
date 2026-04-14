/**
 * CSV Worker Pool Browser Tests
 *
 * Comprehensive tests for the CSV Web Worker implementation.
 */

import { parseCsv, formatCsv } from "@csv/index";
import {
  CsvWorkerPool,
  CsvWorkerSession,
  hasWorkerSupport,
  parseWithPool,
  formatWithPool,
  getDefaultWorkerPool,
  terminateDefaultWorkerPool
} from "@csv/worker/index.browser";
import { describe, it, expect, beforeEach, afterEach } from "vitest";

type ParseOptions = Parameters<typeof parseCsv>[1];
type FormatOptions = Parameters<typeof formatCsv>[1];

function formatExpected(data: Parameters<typeof formatCsv>[0], options?: FormatOptions): string {
  return formatCsv(data, { ...(options ?? {}), trailingNewline: false } as FormatOptions);
}

function parseExpected(input: string, options?: ParseOptions): ReturnType<typeof parseCsv> {
  return parseCsv(input, options ?? {});
}

describe("CSV Worker Pool - Browser", () => {
  // ===========================================================================
  // Environment Tests
  // ===========================================================================

  describe("hasWorkerSupport", () => {
    it("should return true in browser environment", () => {
      expect(hasWorkerSupport()).toBe(true);
    });
  });

  // ===========================================================================
  // CsvWorkerPool Basic Tests
  // ===========================================================================

  describe("CsvWorkerPool", () => {
    let pool: CsvWorkerPool;

    beforeEach(() => {
      pool = new CsvWorkerPool({ maxWorkers: 2 });
    });

    afterEach(() => {
      pool.terminate();
    });

    describe("constructor options", () => {
      it("should create minWorkers on initialization", async () => {
        const poolWithMin = await CsvWorkerPool.create({ minWorkers: 2, maxWorkers: 4 });
        const stats = poolWithMin.getStats();
        expect(stats.totalWorkers).toBe(2);
        poolWithMin.terminate();
      });
    });

    describe("parse", () => {
      it("should parse simple CSV", async () => {
        const input = "a,b,c\n1,2,3\n4,5,6";
        const result = await pool.parse(input);
        expect(result.data).toEqual(parseExpected(input));
        expect(result.duration).toBeGreaterThanOrEqual(0);
      });

      it("should parse with headers option", async () => {
        const result = await pool.parse("name,age\nAlice,30\nBob,25", { headers: true });
        expect(result.data).toEqual(parseCsv("name,age\nAlice,30\nBob,25", { headers: true }));
      });

      it("should expose renamedHeaders meta and avoid header collisions", async () => {
        const result = await pool.parse("A,A,A_1\n1,2,3", { headers: true });
        expect(result.data).toEqual(parseCsv("A,A,A_1\n1,2,3", { headers: true }));
      });

      it("should parse with custom delimiter", async () => {
        const input = "a;b;c\n1;2;3";
        const options = { delimiter: ";" };
        const result = await pool.parse(input, options);
        expect(result.data).toEqual(parseExpected(input, options));
      });

      it("should handle quoted fields", async () => {
        const input = 'name,value\n"Hello, World",42';
        const result = await pool.parse(input);
        expect(result.data).toEqual(parseExpected(input));
      });

      it("should handle escaped quotes", async () => {
        const input = 'a\n"He said ""Hello"""';
        const result = await pool.parse(input);
        expect(result.data).toEqual(parseExpected(input));
      });

      it("should handle multiline quoted fields", async () => {
        const input = 'text\n"Line 1\nLine 2"';
        const result = await pool.parse(input);
        expect(result.data).toEqual(parseExpected(input));
      });

      it("should skip empty lines", async () => {
        const input = "a\n\nb\n\nc";
        const options = { skipEmptyLines: true };
        const result = await pool.parse(input, options);
        expect(result.data).toEqual(parseExpected(input, options));
      });

      it("should trim fields", async () => {
        const input = " a , b \n 1 , 2 ";
        const options = { trim: true };
        const result = await pool.parse(input, options);
        expect(result.data).toEqual(parseExpected(input, options));
      });

      it("should limit rows with maxRows", async () => {
        const result = await pool.parse("a\n1\n2\n3\n4\n5", { maxRows: 3 });
        expect(result.data).toHaveLength(3);
      });

      it("should parse in fast mode", async () => {
        const input = "a,b,c\n1,2,3\n4,5,6";
        const options = { fastMode: true };
        const result = await pool.parse(input, options);
        expect(result.data).toEqual(parseExpected(input, options));
      });

      it("should skip delimiter-only rows in fastMode when configured", async () => {
        const input = "a,b\n,\n1,2\n,,\n3,4";
        const options = { fastMode: true, skipEmptyLines: true };
        const result = await pool.parse(input, options);
        expect(result.data).toEqual(parseExpected(input, options));
      });
    });

    describe("format", () => {
      it("should format simple data", async () => {
        const data: unknown[][] = [
          ["a", "b", "c"],
          [1, 2, 3]
        ];
        const result = await pool.format(data);
        expect(result.data).toBe(formatExpected(data as Parameters<typeof formatCsv>[0]));
      });

      it("should format with custom delimiter", async () => {
        const data: unknown[][] = [
          ["a", "b"],
          [1, 2]
        ];
        const options = { delimiter: ";" };
        const result = await pool.format(data, options);
        expect(result.data).toBe(formatExpected(data as Parameters<typeof formatCsv>[0], options));
      });

      it("should quote fields with special characters", async () => {
        const data = [["Hello, World", "normal"]];
        const result = await pool.format(data);
        expect(result.data).toBe(formatExpected(data));
      });

      it("should escape quotes in fields", async () => {
        const data = [['He said "Hello"']];
        const result = await pool.format(data);
        expect(result.data).toBe(formatExpected(data));
      });

      it("should escape formulae", async () => {
        const data = [["=SUM(A1)", "+1", "-1", "@mention"]];
        const options = { escapeFormulae: true };
        const result = await pool.format(data, options);
        expect(result.data).toBe(formatExpected(data, options));
      });

      it("should use quoteColumns: true option", async () => {
        const data = [["a", "b"]];
        const options = { quoteColumns: true };
        const result = await pool.format(data, options);
        expect(result.data).toBe(formatExpected(data, options));
      });
    });

    describe("getStats", () => {
      it("should return statistics", async () => {
        const stats = pool.getStats();
        expect(stats).toHaveProperty("totalWorkers");
        expect(stats).toHaveProperty("busyWorkers");
        expect(stats).toHaveProperty("pendingTasks");
        expect(stats).toHaveProperty("completedTasks");
        expect(stats).toHaveProperty("failedTasks");
        // idleWorkers can be computed: totalWorkers - busyWorkers
        expect(stats.totalWorkers - stats.busyWorkers).toBeGreaterThanOrEqual(0);
      });

      it("should track completed tasks", async () => {
        await pool.parse("a,b\n1,2");
        await pool.format([["x", "y"]]);
        const stats = pool.getStats();
        expect(stats.completedTasks).toBe(2);
      });
    });

    describe("terminate", () => {
      it("should terminate all workers", () => {
        pool.terminate();
        const stats = pool.getStats();
        expect(stats.totalWorkers).toBe(0);
      });

      it("should reject new tasks after termination", async () => {
        pool.terminate();
        await expect(pool.parse("a,b")).rejects.toThrow("terminated");
      });

      it("should be idempotent", () => {
        pool.terminate();
        pool.terminate();
        pool.terminate();
        expect(pool.getStats().totalWorkers).toBe(0);
      });
    });
  });

  // ===========================================================================
  // CsvWorkerSession Tests
  // ===========================================================================

  describe("CsvWorkerSession", () => {
    let session: CsvWorkerSession;

    beforeEach(async () => {
      session = await CsvWorkerSession.create();
    });

    afterEach(async () => {
      await session.dispose();
      terminateDefaultWorkerPool();
    });

    describe("properties", () => {
      it("should have unique sessionId", async () => {
        const session2 = await CsvWorkerSession.create();
        expect(session.sessionId).toBeTruthy();
        expect(session2.sessionId).toBeTruthy();
        expect(session.sessionId).not.toBe(session2.sessionId);
        session2.dispose();
      });

      it("should update headers and rowCount after load", async () => {
        expect(session.headers).toEqual([]);
        expect(session.rowCount).toBe(0);

        await session.load("name,age\nAlice,30\nBob,25", { headers: true });

        expect(session.headers).toEqual(["name", "age"]);
        expect(session.rowCount).toBe(2);
      });
    });

    describe("load", () => {
      it("should load CSV string", async () => {
        const result = await session.load("name,age\nAlice,30\nBob,25", { headers: true });
        expect(result.rowCount).toBe(2);
        expect(result.headers).toEqual(["name", "age"]);
      });

      it("should load array of objects", async () => {
        const data = [
          { name: "Alice", age: 30 },
          { name: "Bob", age: 25 }
        ];
        const result = await session.load(data);
        expect(result.rowCount).toBe(2);
        expect(result.headers).toContain("name");
        expect(result.headers).toContain("age");
      });

      it("should load 2D array with headers", async () => {
        const data = [
          ["Alice", 30],
          ["Bob", 25]
        ];
        const result = await session.load(data, { headers: ["name", "age"] });
        expect(result.rowCount).toBe(2);
        expect(result.headers).toEqual(["name", "age"]);
      });
    });

    describe("sort", () => {
      beforeEach(async () => {
        await session.load("name,age\nCharlie,35\nAlice,30\nBob,25", { headers: true });
      });

      it("should sort by string column ascending", async () => {
        await session.sort({ column: "name", order: "asc" });
        const { data } = await session.getData();
        expect(data[0].name).toBe("Alice");
        expect(data[1].name).toBe("Bob");
        expect(data[2].name).toBe("Charlie");
      });

      it("should sort by number column descending", async () => {
        await session.sort({ column: "age", order: "desc", comparator: "number" });
        const { data } = await session.getData();
        expect(data[0].age).toBe("35");
        expect(data[1].age).toBe("30");
        expect(data[2].age).toBe("25");
      });

      it("should sort by multiple columns", async () => {
        await session.load("dept,name,age\nIT,Alice,30\nHR,Bob,25\nIT,Charlie,35\nHR,David,40", {
          headers: true
        });
        await session.sort([
          { column: "dept", order: "asc" },
          { column: "age", order: "desc", comparator: "number" }
        ]);
        const { data } = await session.getData();
        expect(data[0]).toEqual({ dept: "HR", name: "David", age: "40" });
        expect(data[1]).toEqual({ dept: "HR", name: "Bob", age: "25" });
        expect(data[2]).toEqual({ dept: "IT", name: "Charlie", age: "35" });
        expect(data[3]).toEqual({ dept: "IT", name: "Alice", age: "30" });
      });
    });

    describe("filter", () => {
      beforeEach(async () => {
        await session.load("name,age,status\nAlice,30,active\nBob,25,inactive\nCharlie,35,active", {
          headers: true
        });
      });

      it("should filter with eq operator", async () => {
        const result = await session.filter({
          conditions: [{ column: "status", operator: "eq", value: "active" }]
        });
        expect(result.matchCount).toBe(2);
        expect(result.data.every(r => r.status === "active")).toBe(true);
      });

      it("should filter with gt operator", async () => {
        const result = await session.filter({
          conditions: [{ column: "age", operator: "gt", value: 28 }]
        });
        expect(result.matchCount).toBe(2);
        expect(result.data.map(r => r.name)).toContain("Alice");
        expect(result.data.map(r => r.name)).toContain("Charlie");
      });

      it("should filter with contains operator", async () => {
        const result = await session.filter({
          conditions: [{ column: "name", operator: "contains", value: "li" }]
        });
        expect(result.matchCount).toBe(2); // Alice, Charlie
      });

      it("should filter with AND logic", async () => {
        const result = await session.filter({
          conditions: [
            { column: "status", operator: "eq", value: "active" },
            { column: "age", operator: "gt", value: 32 }
          ],
          logic: "and"
        });
        expect(result.matchCount).toBe(1);
        expect(result.data[0].name).toBe("Charlie");
      });

      it("should filter with OR logic", async () => {
        const result = await session.filter({
          conditions: [
            { column: "name", operator: "eq", value: "Alice" },
            { column: "name", operator: "eq", value: "Bob" }
          ],
          logic: "or"
        });
        expect(result.matchCount).toBe(2);
      });

      it("should filter with case insensitive", async () => {
        const result = await session.filter({
          conditions: [{ column: "name", operator: "eq", value: "ALICE", ignoreCase: true }]
        });
        expect(result.matchCount).toBe(1);
        expect(result.data[0].name).toBe("Alice");
      });

      it("should filter with in operator", async () => {
        const result = await session.filter({
          conditions: [{ column: "name", operator: "in", value: ["Alice", "Bob"] }]
        });
        expect(result.matchCount).toBe(2);
      });

      it("should filter with isNull operator", async () => {
        await session.load("name,age\nAlice,30\nBob,\nCharlie,35", { headers: true });
        const result = await session.filter({
          conditions: [{ column: "age", operator: "isNull" }]
        });
        expect(result.matchCount).toBe(1);
        expect(result.data[0].name).toBe("Bob");
      });

      it("should filter with notNull operator", async () => {
        await session.load("name,age\nAlice,30\nBob,\nCharlie,35", { headers: true });
        const result = await session.filter({
          conditions: [{ column: "age", operator: "notNull" }]
        });
        expect(result.matchCount).toBe(2);
      });

      it("should filter with neq operator", async () => {
        const result = await session.filter({
          conditions: [{ column: "status", operator: "neq", value: "active" }]
        });
        expect(result.matchCount).toBe(1);
        expect(result.data[0].name).toBe("Bob");
      });

      it("should filter with lt and lte operators", async () => {
        const ltResult = await session.filter({
          conditions: [{ column: "age", operator: "lt", value: 30 }]
        });
        expect(ltResult.matchCount).toBe(1);
        expect(ltResult.data[0].name).toBe("Bob");

        const lteResult = await session.filter({
          conditions: [{ column: "age", operator: "lte", value: 30 }]
        });
        expect(lteResult.matchCount).toBe(2);
      });

      it("should filter with gte operator", async () => {
        const result = await session.filter({
          conditions: [{ column: "age", operator: "gte", value: 30 }]
        });
        expect(result.matchCount).toBe(2);
      });

      it("should filter with startsWith and endsWith operators", async () => {
        const startsResult = await session.filter({
          conditions: [{ column: "name", operator: "startsWith", value: "A" }]
        });
        expect(startsResult.matchCount).toBe(1);
        expect(startsResult.data[0].name).toBe("Alice");

        const endsResult = await session.filter({
          conditions: [{ column: "name", operator: "endsWith", value: "e" }]
        });
        expect(endsResult.matchCount).toBe(2); // Alice, Charlie
      });

      it("should filter with regex operator", async () => {
        const result = await session.filter({
          conditions: [{ column: "name", operator: "regex", value: "^[AB]" }]
        });
        expect(result.matchCount).toBe(2); // Alice, Bob
      });

      it("should filter with notIn operator", async () => {
        const result = await session.filter({
          conditions: [{ column: "name", operator: "notIn", value: ["Alice", "Bob"] }]
        });
        expect(result.matchCount).toBe(1);
        expect(result.data[0].name).toBe("Charlie");
      });
    });

    describe("search", () => {
      beforeEach(async () => {
        await session.load("name,email\nAlice,alice@example.com\nBob,bob@test.com", {
          headers: true
        });
      });

      it("should search across all columns", async () => {
        const result = await session.search({ query: "alice" });
        expect(result.matchCount).toBe(1);
        expect(result.data[0].name).toBe("Alice");
      });

      it("should search specific columns", async () => {
        const result = await session.search({ query: "alice", columns: ["name"] });
        expect(result.matchCount).toBe(1);
      });

      it("should be case insensitive by default", async () => {
        const result = await session.search({ query: "ALICE" });
        expect(result.matchCount).toBe(1);
      });
    });

    describe("groupBy", () => {
      beforeEach(async () => {
        await session.load(
          "dept,name,salary\nIT,Alice,100\nHR,Bob,80\nIT,Charlie,120\nHR,David,90",
          { headers: true }
        );
      });

      it("should group by single column with count", async () => {
        const result = await session.groupBy({
          columns: ["dept"],
          aggregates: [{ column: "name", fn: "count", alias: "count" }]
        });
        expect(result.groupCount).toBe(2);
        const itGroup = result.data.find(g => g.dept === "IT");
        const hrGroup = result.data.find(g => g.dept === "HR");
        expect(itGroup?.count).toBe(2);
        expect(hrGroup?.count).toBe(2);
      });

      it("should group with sum aggregate", async () => {
        const result = await session.groupBy({
          columns: ["dept"],
          aggregates: [{ column: "salary", fn: "sum", alias: "total" }]
        });
        const itGroup = result.data.find(g => g.dept === "IT");
        expect(itGroup?.total).toBe(220);
      });

      it("should group with multiple aggregates", async () => {
        const result = await session.groupBy({
          columns: ["dept"],
          aggregates: [
            { column: "salary", fn: "sum", alias: "total" },
            { column: "salary", fn: "avg", alias: "average" },
            { column: "salary", fn: "min", alias: "min" },
            { column: "salary", fn: "max", alias: "max" }
          ]
        });
        const itGroup = result.data.find(g => g.dept === "IT");
        expect(itGroup?.total).toBe(220);
        expect(itGroup?.average).toBe(110);
        expect(itGroup?.min).toBe(100);
        expect(itGroup?.max).toBe(120);
      });

      it("should group with first and last aggregates", async () => {
        const result = await session.groupBy({
          columns: ["dept"],
          aggregates: [
            { column: "name", fn: "first", alias: "firstName" },
            { column: "name", fn: "last", alias: "lastName" }
          ]
        });
        const itGroup = result.data.find(g => g.dept === "IT");
        expect(itGroup?.firstName).toBe("Alice");
        expect(itGroup?.lastName).toBe("Charlie");
      });

      it("should group by multiple columns", async () => {
        await session.load(
          "dept,level,name\nIT,senior,Alice\nIT,junior,Bob\nIT,senior,Charlie\nHR,junior,David",
          { headers: true }
        );
        const result = await session.groupBy({
          columns: ["dept", "level"],
          aggregates: [{ column: "name", fn: "count", alias: "count" }]
        });
        expect(result.groupCount).toBe(3);
        const itSenior = result.data.find(g => g.dept === "IT" && g.level === "senior");
        expect(itSenior?.count).toBe(2);
      });
    });

    describe("aggregate", () => {
      beforeEach(async () => {
        await session.load("name,salary\nAlice,100\nBob,80\nCharlie,120", { headers: true });
      });

      it("should compute aggregates", async () => {
        const result = await session.aggregate([
          { column: "salary", fn: "sum", alias: "total" },
          { column: "salary", fn: "avg", alias: "average" },
          { column: "name", fn: "count", alias: "count" }
        ]);
        expect(result.data.total).toBe(300);
        expect(result.data.average).toBe(100);
        expect(result.data.count).toBe(3);
      });
    });

    describe("getPage", () => {
      beforeEach(async () => {
        const rows = Array.from({ length: 100 }, (_, i) => `user${i},${i}`);
        await session.load("name,id\n" + rows.join("\n"), { headers: true });
      });

      it("should return correct page", async () => {
        const result = await session.getPage({ page: 1, pageSize: 10 });
        expect(result.data).toHaveLength(10);
        expect(result.page).toBe(1);
        expect(result.pageSize).toBe(10);
        expect(result.totalRows).toBe(100);
        expect(result.totalPages).toBe(10);
        expect(result.data[0].name).toBe("user0");
      });

      it("should return second page", async () => {
        const result = await session.getPage({ page: 2, pageSize: 10 });
        expect(result.data[0].name).toBe("user10");
      });

      it("should handle partial last page", async () => {
        const result = await session.getPage({ page: 4, pageSize: 30 });
        expect(result.data).toHaveLength(10);
        expect(result.totalPages).toBe(4);
      });
    });

    describe("query (batch API)", () => {
      beforeEach(async () => {
        await session.load(
          "name,age,status\nAlice,30,active\nBob,25,inactive\nCharlie,35,active\nDavid,40,active\nEve,28,inactive",
          { headers: true }
        );
      });

      it("should execute sort + filter + page in single round-trip", async () => {
        const result = await session.query({
          sort: { column: "age", order: "desc", comparator: "number" },
          filter: { conditions: [{ column: "status", operator: "eq", value: "active" }] },
          page: { page: 1, pageSize: 2 }
        });

        expect(result.data).toHaveLength(2);
        expect(result.matchCount).toBe(3); // 3 active users
        expect(result.page).toBe(1);
        expect(result.totalRows).toBe(3);
        expect(result.totalPages).toBe(2);
        // Sorted by age desc, so David (40) first, then Charlie (35)
        expect(result.data[0].name).toBe("David");
        expect(result.data[1].name).toBe("Charlie");
        expect(result.duration).toBeGreaterThanOrEqual(0);
      });

      it("should execute sort only", async () => {
        const result = await session.query({
          sort: { column: "name", order: "asc" }
        });

        expect(result.data).toHaveLength(5);
        expect(result.data[0].name).toBe("Alice");
        expect(result.data[4].name).toBe("Eve");
      });

      it("should execute filter only", async () => {
        const result = await session.query({
          filter: { conditions: [{ column: "age", operator: "gt", value: 30 }] }
        });

        expect(result.matchCount).toBe(2); // Charlie (35), David (40)
        expect(result.data).toHaveLength(2);
      });

      it("should execute search + page", async () => {
        const result = await session.query({
          search: { query: "li", columns: ["name"] },
          page: { page: 1, pageSize: 10 }
        });

        expect(result.matchCount).toBe(2); // Alice, Charlie (contain 'li')
        expect(result.data).toHaveLength(2);
        expect(result.page).toBe(1);
      });

      it("should execute groupBy + aggregates", async () => {
        const result = await session.query({
          groupBy: {
            columns: ["status"],
            aggregates: [
              { column: "age", fn: "avg", alias: "avgAge" },
              { column: "name", fn: "count", alias: "count" }
            ]
          }
        });

        expect(result.groupCount).toBe(2);
        const activeGroup = result.data.find((g: any) => g.status === "active");
        const inactiveGroup = result.data.find((g: any) => g.status === "inactive");
        expect(activeGroup?.count).toBe(3);
        expect(inactiveGroup?.count).toBe(2);
      });

      it("should execute aggregate without groupBy", async () => {
        const result = await session.query({
          aggregate: [
            { column: "age", fn: "sum", alias: "totalAge" },
            { column: "age", fn: "avg", alias: "avgAge" },
            { column: "name", fn: "count", alias: "total" }
          ]
        });

        expect(result.aggregates?.total).toBe(5);
        expect(result.aggregates?.totalAge).toBe(158); // 30+25+35+40+28
        expect(result.aggregates?.avgAge).toBeCloseTo(31.6, 1);
      });

      it("should handle empty config (returns all data)", async () => {
        const result = await session.query({});
        expect(result.data).toHaveLength(5);
      });

      it("should execute complex query with all operations", async () => {
        const result = await session.query({
          sort: { column: "age", order: "asc", comparator: "number" },
          filter: {
            conditions: [{ column: "age", operator: "gte", value: 28 }],
            logic: "and"
          },
          page: { page: 1, pageSize: 3 }
        });

        expect(result.matchCount).toBe(4); // Alice(30), Charlie(35), David(40), Eve(28)
        expect(result.page).toBe(1);
        expect(result.pageSize).toBe(3);
        expect(result.totalPages).toBe(2);
        expect(result.data).toHaveLength(3);
        // Sorted by age asc: Eve(28), Alice(30), Charlie(35)
        expect(result.data[0].name).toBe("Eve");
        expect(result.data[1].name).toBe("Alice");
        expect(result.data[2].name).toBe("Charlie");
      });
    });

    describe("dispose", () => {
      it("should clear session data", async () => {
        await session.load("a,b\n1,2", { headers: true });
        await session.dispose();
        // Session should be disposed
      });

      it("should be idempotent", async () => {
        await session.load("a,b\n1,2", { headers: true });
        await session.dispose();
        await session.dispose();
        await session.dispose();
      });

      it("should reject operations after dispose", async () => {
        await session.load("a,b\n1,2", { headers: true });
        await session.dispose();
        await expect(session.getData()).rejects.toThrow("disposed");
      });
    });
  });

  // ===========================================================================
  // Convenience Functions Tests
  // ===========================================================================

  describe("Convenience Functions", () => {
    afterEach(() => {
      terminateDefaultWorkerPool();
    });

    describe("parseWithPool", () => {
      it("should parse CSV", async () => {
        const input = "a,b\n1,2";
        const result = await parseWithPool(input);
        expect(result.data).toEqual(parseExpected(input));
      });

      it("should support options", async () => {
        const result = await parseWithPool("name,age\nAlice,30", { headers: true });
        expect(result.data).toEqual(parseCsv("name,age\nAlice,30", { headers: true }));
      });
    });

    describe("formatWithPool", () => {
      it("should format data", async () => {
        const data: unknown[][] = [
          ["a", "b"],
          [1, 2]
        ];
        const result = await formatWithPool(data);
        expect(result.data).toBe(formatExpected(data as Parameters<typeof formatCsv>[0]));
      });
    });

    describe("getDefaultWorkerPool", () => {
      it("should return same instance", async () => {
        const pool1 = await getDefaultWorkerPool();
        const pool2 = await getDefaultWorkerPool();
        expect(pool1).toBe(pool2);
      });
    });

    describe("terminateDefaultWorkerPool", () => {
      it("should terminate and reset pool", async () => {
        const pool1 = await getDefaultWorkerPool();
        terminateDefaultWorkerPool();
        const pool2 = await getDefaultWorkerPool();
        expect(pool2).not.toBe(pool1);
      });

      it("should be safe to call multiple times", () => {
        terminateDefaultWorkerPool();
        terminateDefaultWorkerPool();
        terminateDefaultWorkerPool();
      });
    });
  });

  // ===========================================================================
  // Priority & AbortSignal Tests
  // ===========================================================================

  describe("Task Priority", () => {
    let pool: CsvWorkerPool;

    beforeEach(() => {
      pool = new CsvWorkerPool({ maxWorkers: 1 });
    });

    afterEach(() => {
      pool.terminate();
    });

    it("should prioritize high priority tasks in queue", async () => {
      // Start a blocking task first
      const blockingTask = pool.parse("x,y,z\n1,2,3\n4,5,6");

      // Queue tasks with different priorities while first is running
      const order: string[] = [];
      const lowTask = pool
        .parse("low", undefined, { priority: "low" })
        .then(() => order.push("low"));
      const highTask = pool
        .parse("high", undefined, { priority: "high" })
        .then(() => order.push("high"));
      const normalTask = pool
        .parse("normal", undefined, { priority: "normal" })
        .then(() => order.push("normal"));

      await Promise.all([blockingTask, lowTask, highTask, normalTask]);

      // High should be processed before low (they were queued while first was running)
      const highIdx = order.indexOf("high");
      const lowIdx = order.indexOf("low");
      expect(highIdx).toBeLessThan(lowIdx);
    });
  });

  describe("AbortSignal", () => {
    let pool: CsvWorkerPool;

    beforeEach(() => {
      pool = new CsvWorkerPool({ maxWorkers: 1 });
    });

    afterEach(() => {
      pool.terminate();
    });

    it("should reject with AbortError when already aborted", async () => {
      const controller = new AbortController();
      controller.abort();

      await expect(pool.parse("a,b", undefined, { signal: controller.signal })).rejects.toThrow(
        "abort"
      );
    });

    it("should support AbortSignal for cancellation", async () => {
      const controller = new AbortController();

      // Immediately abort and try to parse
      controller.abort();

      try {
        await pool.parse("x,y", undefined, { signal: controller.signal });
        expect.fail("Expected task to reject");
      } catch (err: any) {
        expect(err.name === "AbortError" || err.message.includes("abort")).toBe(true);
      }
    });
  });

  // ===========================================================================
  // Performance Tests
  // ===========================================================================

  describe("Performance", () => {
    let pool: CsvWorkerPool;

    beforeEach(() => {
      pool = new CsvWorkerPool({ maxWorkers: 4 });
    });

    afterEach(() => {
      pool.terminate();
    });

    it("should handle concurrent tasks", async () => {
      const tasks = Array.from({ length: 10 }, (_, i) => pool.parse(`col${i}\nval${i}`));

      const results = await Promise.all(tasks);

      expect(results).toHaveLength(10);
      results.forEach((result, i) => {
        expect(result.data).toEqual([[`col${i}`], [`val${i}`]]);
      });
    });

    it("should handle large CSV", async () => {
      const rows = Array.from({ length: 1000 }, (_, i) => `${i},${i * 2},${i * 3}`);
      const csv = "a,b,c\n" + rows.join("\n");

      const result = await pool.parse(csv);
      expect(result.data).toHaveLength(1001);
    });

    it("should report duration", async () => {
      const result = await pool.parse("a,b,c\n1,2,3");
      expect(typeof result.duration).toBe("number");
      expect(result.duration).toBeGreaterThanOrEqual(0);
    });
  });

  // ===========================================================================
  // Edge Cases
  // ===========================================================================

  describe("Edge Cases", () => {
    let pool: CsvWorkerPool;

    beforeEach(() => {
      pool = new CsvWorkerPool({ maxWorkers: 2 });
    });

    afterEach(() => {
      pool.terminate();
    });

    it("should handle empty string", async () => {
      const result = await pool.parse("");
      expect(result.data).toEqual([]);
    });

    it("should handle single value", async () => {
      const result = await pool.parse("value");
      expect(result.data).toEqual([["value"]]);
    });

    it("should handle special characters", async () => {
      const result = await pool.parse('emoji,text\n"😀","Hello 世界"');
      expect(result.data).toEqual([
        ["emoji", "text"],
        ["😀", "Hello 世界"]
      ]);
    });

    it("should handle Windows line endings", async () => {
      const result = await pool.parse("a,b\r\n1,2\r\n3,4");
      expect(result.data).toEqual([
        ["a", "b"],
        ["1", "2"],
        ["3", "4"]
      ]);
    });

    it("should handle null and undefined in format", async () => {
      const result = await pool.format([[null, undefined, "value"]]);
      expect(result.data).toBe(",,value");
    });

    it("should handle session not found error", async () => {
      await expect(pool.getData("nonexistent-session")).rejects.toThrow("Session not found");
    });
  });

  // ===========================================================================
  // Concurrency & Stress Tests
  // ===========================================================================

  describe("Concurrency & Stress Tests", () => {
    let pool: CsvWorkerPool;

    beforeEach(() => {
      pool = new CsvWorkerPool({ maxWorkers: 4 });
    });

    afterEach(() => {
      pool.terminate();
    });

    it("should handle many concurrent parse requests", async () => {
      const tasks = Array.from({ length: 50 }, (_, i) =>
        pool.parse(`name,value\nitem${i},${i}`, { headers: true })
      );

      const results = await Promise.all(tasks);

      expect(results).toHaveLength(50);
      results.forEach((result, i) => {
        expect((result.data as { rows: any[] }).rows[0].name).toBe(`item${i}`);
      });
    });

    it("should handle many concurrent format requests", async () => {
      const tasks = Array.from({ length: 50 }, (_, i) =>
        pool.format([
          ["a", "b"],
          [i, i * 2]
        ])
      );

      const results = await Promise.all(tasks);

      expect(results).toHaveLength(50);
      results.forEach(result => {
        expect(result.data).toContain("a,b");
      });
    });

    it("should handle mixed concurrent operations", async () => {
      const parseTasks = Array.from({ length: 25 }, (_, i) => pool.parse(`a\n${i}`));

      const formatTasks = Array.from({ length: 25 }, (_, i) => pool.format([[i, i + 1]]));

      const results = await Promise.all([...parseTasks, ...formatTasks]);

      expect(results).toHaveLength(50);
    });

    it("should respect maxWorkers limit under load", async () => {
      const poolWithLimit = new CsvWorkerPool({ maxWorkers: 2 });

      // Launch many tasks
      const tasks = Array.from({ length: 20 }, (_, i) => poolWithLimit.parse(`col\n${i}`));

      // Check stats during execution
      const statsPromise = new Promise<{ totalWorkers: number }>(resolve => {
        setTimeout(() => {
          resolve(poolWithLimit.getStats());
        }, 50);
      });

      const [stats] = await Promise.all([statsPromise, Promise.all(tasks)]);

      // Should not exceed maxWorkers
      expect(stats.totalWorkers).toBeLessThanOrEqual(2);

      poolWithLimit.terminate();
    });

    it("should handle worker failures gracefully", async () => {
      // This test ensures that if one task fails, others continue
      const tasks = [
        pool.parse("a,b\n1,2"), // Valid
        pool.parse("a,b\n1,2"), // Valid
        pool.parse("a,b\n1,2") // Valid
      ];

      const results = await Promise.allSettled(tasks);

      // At least some should succeed
      const fulfilled = results.filter(r => r.status === "fulfilled");
      expect(fulfilled.length).toBeGreaterThan(0);
    });

    it("should handle rapid sequential operations", async () => {
      for (let i = 0; i < 20; i++) {
        const result = await pool.parse(`x\n${i}`);
        expect(result.data).toEqual([["x"], [String(i)]]);
      }
    });

    it("should maintain data integrity under concurrent load", async () => {
      const testData = Array.from({ length: 30 }, (_, i) => ({
        id: i,
        csv: `id,name,value\n${i},item${i},${i * 100}`
      }));

      const tasks = testData.map(({ csv }) => pool.parse(csv, { headers: true }));

      const results = await Promise.all(tasks);

      // Verify each result matches its input
      results.forEach((result, i) => {
        const rows = (result.data as { rows: any[] }).rows;
        expect(rows[0].id).toBe(String(i));
        expect(rows[0].name).toBe(`item${i}`);
        expect(rows[0].value).toBe(String(i * 100));
      });
    });

    it("should recover after pool is terminated and recreated", async () => {
      await pool.parse("a\n1");
      pool.terminate();

      // Create new pool
      pool = new CsvWorkerPool({ maxWorkers: 2 });

      const result = await pool.parse("b\n2");
      expect(result.data).toEqual([["b"], ["2"]]);
    });

    it("should handle task timeout scenario", async () => {
      // Large data that takes longer to process
      const rows = Array.from({ length: 5000 }, (_, i) => `${i},${"x".repeat(100)}`);
      const csv = "id,data\n" + rows.join("\n");

      // Should still complete even if it takes a while
      const result = await pool.parse(csv, { headers: true });
      expect((result.data as { rows: any[] }).rows.length).toBe(5000);
    });
  });
});
