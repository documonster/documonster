import { Workbook } from "@excel/index";
import { addPivotTable, addTable } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

describe("makeCacheFields", () => {
  describe("field categorization", () => {
    it("should create sharedItems for row fields", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Category" }, { name: "Value" }],
        rows: [
          ["A", 10],
          ["B", 20],
          ["A", 30]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // Category (index 0) should have sharedItems
      expect(pivotTable.cacheFields[0].name).toBe("Category");
      expect(pivotTable.cacheFields[0].sharedItems).toEqual(["A", "B"]);

      // Value (index 1) should be a value-only field with minMax
      expect(pivotTable.cacheFields[1].name).toBe("Value");
      expect(pivotTable.cacheFields[1].sharedItems).toBeNull();
      expect(pivotTable.cacheFields[1].minValue).toBe(10);
      expect(pivotTable.cacheFields[1].maxValue).toBe(30);
    });

    it("should create sharedItems for column fields", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Row" }, { name: "Col" }, { name: "Val" }],
        rows: [
          ["R1", "C1", 10],
          ["R1", "C2", 20]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Row"],
        columns: ["Col"],
        values: ["Val"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // Row (index 0) should have sharedItems
      expect(pivotTable.cacheFields[0].sharedItems).toEqual(["R1"]);

      // Col (index 1) should have sharedItems
      expect(pivotTable.cacheFields[1].sharedItems).toEqual(["C1", "C2"]);

      // Val (index 2) should be value-only
      expect(pivotTable.cacheFields[2].sharedItems).toBeNull();
      expect(pivotTable.cacheFields[2].minValue).toBe(10);
      expect(pivotTable.cacheFields[2].maxValue).toBe(20);
    });

    it("should create empty sharedItems for unused fields", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Used" }, { name: "Unused" }, { name: "Value" }],
        rows: [
          ["A", "X", 10],
          ["B", "Y", 20]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Used"],
        columns: [],
        values: ["Value"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // Used (index 0) should have sharedItems
      expect(pivotTable.cacheFields[0].name).toBe("Used");
      expect(pivotTable.cacheFields[0].sharedItems).toEqual(["A", "B"]);

      // Unused (index 1) should have null sharedItems and no minMax
      expect(pivotTable.cacheFields[1].name).toBe("Unused");
      expect(pivotTable.cacheFields[1].sharedItems).toBeNull();
      expect(pivotTable.cacheFields[1].minValue).toBeUndefined();
      expect(pivotTable.cacheFields[1].maxValue).toBeUndefined();

      // Value (index 2) should be value-only with minMax
      expect(pivotTable.cacheFields[2].name).toBe("Value");
      expect(pivotTable.cacheFields[2].sharedItems).toBeNull();
      expect(pivotTable.cacheFields[2].minValue).toBe(10);
      expect(pivotTable.cacheFields[2].maxValue).toBe(20);
    });

    it("should create numeric sharedItems when field is both row and value", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
        rows: [
          ["a1", "b1", 5],
          ["a1", "b2", 5],
          ["a2", "b1", 24],
          ["a2", "b2", 35],
          ["a3", "b1", 45],
          ["a3", "b2", 45]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      // Same field "C" for both rows and values
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["C"],
        columns: ["B"],
        values: ["C"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // A (index 0) should be unused (null, no minMax)
      expect(pivotTable.cacheFields[0].name).toBe("A");
      expect(pivotTable.cacheFields[0].sharedItems).toBeNull();
      expect(pivotTable.cacheFields[0].minValue).toBeUndefined();

      // B (index 1) should have string sharedItems (column field)
      expect(pivotTable.cacheFields[1].name).toBe("B");
      expect(pivotTable.cacheFields[1].sharedItems).toEqual(["b1", "b2"]);

      // C (index 2) should have numeric sharedItems (both row and value)
      expect(pivotTable.cacheFields[2].name).toBe("C");
      expect(pivotTable.cacheFields[2].sharedItems).toEqual([5, 24, 35, 45]);
    });
  });

  describe("pivotField attributes", () => {
    it("should set dataField=1 when field is both row and value", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
        rows: [
          ["a1", "b1", 5],
          ["a2", "b2", 10]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["C"],
        columns: ["B"],
        values: ["C"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // Field C (index 2) should be in both rows and values
      expect(pivotTable.rows).toContain(2);
      expect(pivotTable.values).toContain(2);
    });

    it("should set dataField=1 when field is both column and value", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "A" }, { name: "B" }, { name: "C" }],
        rows: [
          ["a1", "b1", 5],
          ["a2", "b2", 10]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["A"],
        columns: ["C"],
        values: ["C"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // Field C (index 2) should be in both columns and values
      expect(pivotTable.columns).toContain(2);
      expect(pivotTable.values).toContain(2);
    });
  });

  describe("edge cases", () => {
    it("should handle all values being the same (single unique value)", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Category" }, { name: "Value" }],
        rows: [
          ["Same", 100],
          ["Same", 100],
          ["Same", 100]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // Category should have only one unique value
      expect(pivotTable.cacheFields[0].sharedItems).toEqual(["Same"]);
      // Value should have minValue = maxValue
      expect(pivotTable.cacheFields[1].minValue).toBe(100);
      expect(pivotTable.cacheFields[1].maxValue).toBe(100);
    });

    it("should handle negative numbers in value fields", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Category" }, { name: "Value" }],
        rows: [
          ["A", -50],
          ["B", 0],
          ["C", 50]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      expect(pivotTable.cacheFields[1].minValue).toBe(-50);
      expect(pivotTable.cacheFields[1].maxValue).toBe(50);
    });

    it("should handle decimal values", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Category" }, { name: "Price" }],
        rows: [
          ["A", 10.5],
          ["B", 20.75],
          ["C", 0.01]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Category"],
        columns: [],
        values: ["Price"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      expect(pivotTable.cacheFields[1].minValue).toBe(0.01);
      expect(pivotTable.cacheFields[1].maxValue).toBe(20.75);
    });

    it("should handle many columns (wide data)", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      // Create 10 columns
      const columns: { name: string }[] = [];
      for (let i = 1; i <= 10; i++) {
        columns.push({ name: `Col${i}` });
      }

      const rows = [
        ["A", 1, 2, 3, 4, 5, 6, 7, 8, 9],
        ["B", 10, 20, 30, 40, 50, 60, 70, 80, 90]
      ];

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns,
        rows
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Col1"],
        columns: [],
        values: ["Col10"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // Should have 10 cache fields
      expect(pivotTable.cacheFields.length).toBe(10);

      // Col1 should have sharedItems, rest should be unused or value
      expect(pivotTable.cacheFields[0].sharedItems).toEqual(["A", "B"]);
      expect(pivotTable.cacheFields[9].sharedItems).toBeNull();
      expect(pivotTable.cacheFields[9].minValue).toBe(9);
      expect(pivotTable.cacheFields[9].maxValue).toBe(90);
    });

    it("should handle mixed null/undefined values in data", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Category" }, { name: "Value" }],
        rows: [
          ["A", 10],
          ["B", null as any],
          ["C", undefined as any],
          ["A", 20]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      // Category should have unique values (null/undefined filtered)
      expect(pivotTable.cacheFields[0].sharedItems).toEqual(["A", "B", "C"]);
      // Value minMax should only consider numeric values
      expect(pivotTable.cacheFields[1].minValue).toBe(10);
      expect(pivotTable.cacheFields[1].maxValue).toBe(20);
    });

    it("should handle zero values correctly", async () => {
      const workbook = Workbook.create();
      const worksheet = Workbook.addWorksheet(workbook);

      const table = addTable(worksheet, {
        name: "TestTable",
        ref: "A1",
        headerRow: true,
        columns: [{ name: "Category" }, { name: "Value" }],
        rows: [
          ["A", 0],
          ["B", 0],
          ["C", 0]
        ]
      });

      const worksheet2 = Workbook.addWorksheet(workbook, "Pivot");
      addPivotTable(worksheet2, {
        sourceTable: table,
        rows: ["Category"],
        columns: [],
        values: ["Value"],
        metric: "sum"
      });

      const pivotTable = workbook.pivotTables[0];

      expect(pivotTable.cacheFields[1].minValue).toBe(0);
      expect(pivotTable.cacheFields[1].maxValue).toBe(0);
    });
  });
});
