import { definedNamesAdd } from "@excel/defined-names";
import { calculateFormulas } from "@excel/formula-adapter";
import { Cell, Workbook } from "@excel/index";
import { getDefinedNames } from "@excel/workbook";
import { addTable } from "@excel/worksheet";
import { describe, it, expect } from "vitest";

/**
 * End-to-end verification of the Excel formula engine.
 *
 * Each test creates a workbook from scratch, sets up formulas,
 * calls calculateFormulas(), and verifies that the computed results
 * match expected Excel behavior.
 */
describe("Formula Engine E2E Verification", () => {
  // ==========================================================================
  // 1. Named Ranges
  // ==========================================================================
  describe("named ranges", () => {
    it("should resolve a named range pointing to a single cell", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Set up: TaxRate = 0.08 in B1
      Cell.setValue(ws, "B1", 0.08);
      // Register the defined name: "TaxRate" → Sheet1!$B$1
      definedNamesAdd(getDefinedNames(wb), "Sheet1!$B$1", "TaxRate");

      // A1 = 100 (the base amount)
      Cell.setValue(ws, "A1", 100);

      // Formula uses the named range: =TaxRate * A1
      Cell.setValue(ws, "C1", { formula: "TaxRate*A1", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "C1")).toBe(8);
    });

    it("should resolve a named range pointing to a multi-cell range", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", 10);
      Cell.setValue(ws, "A2", 20);
      Cell.setValue(ws, "A3", 30);
      // Define "MyData" as Sheet1!$A$1:$A$3
      definedNamesAdd(getDefinedNames(wb), "Sheet1!$A$1", "MyData");
      definedNamesAdd(getDefinedNames(wb), "Sheet1!$A$2", "MyData");
      definedNamesAdd(getDefinedNames(wb), "Sheet1!$A$3", "MyData");

      Cell.setValue(ws, "B1", { formula: "SUM(MyData)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B1")).toBe(60);
    });
  });

  // ==========================================================================
  // 2. Whole Column/Row References
  // ==========================================================================
  describe("whole column/row references", () => {
    it("SUM(A:A) should sum all data in column A", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", 10);
      Cell.setValue(ws, "A2", 20);
      Cell.setValue(ws, "A3", 30);
      Cell.setValue(ws, "A4", 40);
      Cell.setValue(ws, "A5", 50);

      // Place formula in a different column to avoid self-reference
      Cell.setValue(ws, "B1", { formula: "SUM(A:A)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B1")).toBe(150);
    });

    it("SUM(1:1) should sum all data in row 1", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "B1", 2);
      Cell.setValue(ws, "C1", 3);

      // Place formula in row 2 to avoid self-reference
      Cell.setValue(ws, "A2", { formula: "SUM(1:1)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A2")).toBe(6);
    });
  });

  // ==========================================================================
  // 3. XLOOKUP
  // ==========================================================================
  describe("XLOOKUP", () => {
    it("should find a value by exact match lookup", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Lookup table: fruits and prices
      Cell.setValue(ws, "A1", "Apple");
      Cell.setValue(ws, "A2", "Banana");
      Cell.setValue(ws, "A3", "Cherry");
      Cell.setValue(ws, "B1", 1.5);
      Cell.setValue(ws, "B2", 0.75);
      Cell.setValue(ws, "B3", 3.0);

      // XLOOKUP("Banana", A1:A3, B1:B3) should return 0.75
      Cell.setValue(ws, "D1", {
        formula: 'XLOOKUP("Banana",A1:A3,B1:B3)',
        result: 0
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "D1")).toBe(0.75);
    });

    it("should return if_not_found value when no match", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", "Apple");
      Cell.setValue(ws, "A2", "Banana");
      Cell.setValue(ws, "B1", 1.5);
      Cell.setValue(ws, "B2", 0.75);

      // XLOOKUP("Mango", A1:A2, B1:B2, "Not found")
      Cell.setValue(ws, "D1", {
        formula: 'XLOOKUP("Mango",A1:A2,B1:B2,"Not found")',
        result: ""
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "D1")).toBe("Not found");
    });
  });

  // ==========================================================================
  // 4. Financial Functions
  // ==========================================================================
  describe("financial functions", () => {
    it("PMT should compute monthly mortgage payment", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // PMT(rate, nper, pv) = PMT(0.08/12, 360, 200000)
      // Expected: approximately -1467.53
      Cell.setValue(ws, "A1", { formula: "PMT(0.08/12,360,200000)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBeCloseTo(-1467.53, 1);
    });

    it("FV should compute future value of annuity", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // FV(rate, nper, pmt) = FV(0.06/12, 120, -200)
      // Expected: approximately 32775.87
      Cell.setValue(ws, "A1", { formula: "FV(0.06/12,120,-200)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBeCloseTo(32775.87, 0);
    });
  });

  // ==========================================================================
  // 5. Structured References (Table[Column])
  // ==========================================================================
  describe("structured references", () => {
    it("should resolve Table[Column] references in formulas", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Create a Table named "Products" with columns Name and Price
      addTable(ws, {
        name: "Products",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "Name" }, { name: "Price" }],
        rows: [
          ["Apple", 1.5],
          ["Banana", 0.75],
          ["Cherry", 3.0]
        ]
      });

      // Formula referencing the table column: SUM of Products[Price]
      Cell.setValue(ws, "D1", { formula: "SUM(Products[Price])", result: 0 });

      calculateFormulas(wb);

      // 1.5 + 0.75 + 3.0 = 5.25
      expect(Cell.getResult(ws, "D1")).toBeCloseTo(5.25);
    });
  });

  // ==========================================================================
  // 6. Dynamic Array Spill
  // ==========================================================================
  describe("dynamic array spill", () => {
    it("FILTER should spill filtered values to adjacent cells", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Data: names in A1:A5, include flags in B1:B5
      Cell.setValue(ws, "A1", "Apple");
      Cell.setValue(ws, "A2", "Banana");
      Cell.setValue(ws, "A3", "Cherry");
      Cell.setValue(ws, "A4", "Date");
      Cell.setValue(ws, "A5", "Elderberry");
      Cell.setValue(ws, "B1", 1); // TRUE
      Cell.setValue(ws, "B2", 0); // FALSE
      Cell.setValue(ws, "B3", 1); // TRUE
      Cell.setValue(ws, "B4", 0); // FALSE
      Cell.setValue(ws, "B5", 1); // TRUE

      // Dynamic array formula with isDynamicArray flag
      Cell.setValue(ws, "D1", {
        formula: "_xlfn._xlws.FILTER(A1:A5,B1:B5)",
        result: "",
        shareType: "array",
        ref: "D1",
        isDynamicArray: true
      });

      calculateFormulas(wb);

      // Should spill: D1=Apple, D2=Cherry, D3=Elderberry
      expect(Cell.getResult(ws, "D1")).toBe("Apple");
      expect(Cell.getValue(ws, "D2")).toBe("Cherry");
      expect(Cell.getValue(ws, "D3")).toBe("Elderberry");
    });
  });

  // ==========================================================================
  // 7. Implicit Intersection
  // ==========================================================================
  describe("implicit intersection", () => {
    it("should pick value from same row when range used in scalar context", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Column data A1:A3
      Cell.setValue(ws, "A1", 10);
      Cell.setValue(ws, "A2", 20);
      Cell.setValue(ws, "A3", 30);

      // Formulas in B1:B3 all reference A1:A3 but in scalar context
      // Each should pick the value from its own row via implicit intersection
      Cell.setValue(ws, "B1", { formula: "A1:A3*2", result: 0 });
      Cell.setValue(ws, "B2", { formula: "A1:A3*2", result: 0 });
      Cell.setValue(ws, "B3", { formula: "A1:A3*2", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B1")).toBe(20); // 10 * 2
      expect(Cell.getResult(ws, "B2")).toBe(40); // 20 * 2
      expect(Cell.getResult(ws, "B3")).toBe(60); // 30 * 2
    });
  });

  // ==========================================================================
  // 8. CSE Array Formula
  // ==========================================================================
  describe("CSE array formula", () => {
    it("{=A1:A3*10} should distribute across B1:B3", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", 1);
      Cell.setValue(ws, "A2", 2);
      Cell.setValue(ws, "A3", 3);

      // CSE array formula: {=A1:A3*10} with ref="B1:B3"
      Cell.setValue(ws, "B1", {
        formula: "A1:A3*10",
        result: 0,
        shareType: "array",
        ref: "B1:B3"
      });
      Cell.setValue(ws, "B2", {
        formula: "A1:A3*10",
        result: 0,
        shareType: "array"
      });
      Cell.setValue(ws, "B3", {
        formula: "A1:A3*10",
        result: 0,
        shareType: "array"
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B1")).toBe(10); // 1*10
      expect(Cell.getResult(ws, "B2")).toBe(20); // 2*10
      expect(Cell.getResult(ws, "B3")).toBe(30); // 3*10
    });
  });

  // ==========================================================================
  // 9. 3D References
  // ==========================================================================
  describe("3D references", () => {
    it("SUM(Sheet1:Sheet3!A1) should sum A1 across three sheets", () => {
      const wb = Workbook.create();
      const ws1 = Workbook.addWorksheet(wb, "Sheet1");
      const ws2 = Workbook.addWorksheet(wb, "Sheet2");
      const ws3 = Workbook.addWorksheet(wb, "Sheet3");
      const summary = Workbook.addWorksheet(wb, "Summary");

      Cell.setValue(ws1, "A1", 10);
      Cell.setValue(ws2, "A1", 20);
      Cell.setValue(ws3, "A1", 30);

      // 3D reference formula on the Summary sheet
      Cell.setValue(summary, "A1", {
        formula: "SUM(Sheet1:Sheet3!A1)",
        result: 0
      });

      calculateFormulas(wb);

      expect(Cell.getResult(summary, "A1")).toBe(60);
    });
  });

  // ==========================================================================
  // 10. LET Function
  // ==========================================================================
  describe("LET function", () => {
    it("LET(x, 10, y, 20, x+y) should return 30", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "LET(x,10,y,20,x+y)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(30);
    });

    it("LET should allow later bindings to reference earlier ones", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // LET(a, 5, b, a*2, a+b) → 5 + 10 = 15
      Cell.setValue(ws, "A1", { formula: "LET(a,5,b,a*2,a+b)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(15);
    });
  });

  // ==========================================================================
  // 11. Engineering Functions
  // ==========================================================================
  describe("engineering functions", () => {
    it('BIN2DEC("1010") should return 10', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: 'BIN2DEC("1010")', result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(10);
    });

    it('HEX2DEC("FF") should return 255', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: 'HEX2DEC("FF")', result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(255);
    });
  });

  // ==========================================================================
  // 12. Statistical Functions: NORM.S.INV
  // ==========================================================================
  describe("statistical functions", () => {
    it("NORM.S.INV(0.975) should be approximately 1.96", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "NORM.S.INV(0.975)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBeCloseTo(1.96, 1);
    });
  });

  // ==========================================================================
  // 13. Database Functions: DSUM
  // ==========================================================================
  describe("database functions", () => {
    it("DSUM should sum matching records based on criteria", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Database table: A1:C5 (header + 4 rows)
      //   Region | Product | Sales
      //   East   | Widget  | 100
      //   West   | Widget  | 200
      //   East   | Gadget  | 150
      //   West   | Gadget  | 300
      Cell.setValue(ws, "A1", "Region");
      Cell.setValue(ws, "B1", "Product");
      Cell.setValue(ws, "C1", "Sales");
      Cell.setValue(ws, "A2", "East");
      Cell.setValue(ws, "B2", "Widget");
      Cell.setValue(ws, "C2", 100);
      Cell.setValue(ws, "A3", "West");
      Cell.setValue(ws, "B3", "Widget");
      Cell.setValue(ws, "C3", 200);
      Cell.setValue(ws, "A4", "East");
      Cell.setValue(ws, "B4", "Gadget");
      Cell.setValue(ws, "C4", 150);
      Cell.setValue(ws, "A5", "West");
      Cell.setValue(ws, "B5", "Gadget");
      Cell.setValue(ws, "C5", 300);

      // Criteria table: E1:E2 (header + 1 criterion)
      // Region = "East"
      Cell.setValue(ws, "E1", "Region");
      Cell.setValue(ws, "E2", "East");

      // DSUM(database, "Sales", criteria)
      // Should sum Sales where Region="East" → 100 + 150 = 250
      Cell.setValue(ws, "G1", {
        formula: 'DSUM(A1:C5,"Sales",E1:E2)',
        result: 0
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "G1")).toBe(250);
    });
  });

  // ==========================================================================
  // 14. YEARFRAC
  // ==========================================================================
  describe("YEARFRAC", () => {
    it("YEARFRAC(DATE(2020,1,1), DATE(2020,7,1), 0) should be approximately 0.5", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // US 30/360 basis (basis=0)
      Cell.setValue(ws, "A1", {
        formula: "YEARFRAC(DATE(2020,1,1),DATE(2020,7,1),0)",
        result: 0
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBeCloseTo(0.5, 2);
    });
  });

  // ==========================================================================
  // 15. IRR
  // ==========================================================================
  describe("IRR", () => {
    it("IRR({-1000, 300, 420, 680}) should be approximately 0.1665", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Put cash flows in cells and use a range reference
      Cell.setValue(ws, "A1", -1000);
      Cell.setValue(ws, "A2", 300);
      Cell.setValue(ws, "A3", 420);
      Cell.setValue(ws, "A4", 680);

      Cell.setValue(ws, "B1", { formula: "IRR(A1:A4)", result: 0 });

      calculateFormulas(wb);

      // IRR should be approximately 0.1665 (~16.65%)
      expect(Cell.getResult(ws, "B1")).toBeCloseTo(0.1665, 2);
    });
  });

  // ==========================================================================
  // 16. INDEX with row=0 (return entire column)
  // ==========================================================================
  describe("INDEX with row=0", () => {
    it("INDEX(A1:B3, 0, 1) should return entire first column as array", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // 3x2 data block
      Cell.setValue(ws, "A1", 10);
      Cell.setValue(ws, "B1", 100);
      Cell.setValue(ws, "A2", 20);
      Cell.setValue(ws, "B2", 200);
      Cell.setValue(ws, "A3", 30);
      Cell.setValue(ws, "B3", 300);

      // INDEX(A1:B3, 0, 1) returns the first column as an array
      // Wrap in SUM to verify: SUM should be 10+20+30=60
      Cell.setValue(ws, "D1", { formula: "SUM(INDEX(A1:B3,0,1))", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "D1")).toBe(60);
    });

    it("INDEX with specific row and col returns a scalar", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", 10);
      Cell.setValue(ws, "B1", 100);
      Cell.setValue(ws, "A2", 20);
      Cell.setValue(ws, "B2", 200);

      // INDEX(A1:B2, 2, 2) should return 200
      Cell.setValue(ws, "D1", { formula: "INDEX(A1:B2,2,2)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "D1")).toBe(200);
    });
  });

  // ==========================================================================
  // 17. MATCH with Wildcards
  // ==========================================================================
  describe("MATCH with wildcards", () => {
    it('MATCH("app*", {"apple","banana","apricot"}, 0) should return 1', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", "apple");
      Cell.setValue(ws, "A2", "banana");
      Cell.setValue(ws, "A3", "apricot");

      // MATCH with wildcard — "app*" matches "apple" at position 1
      Cell.setValue(ws, "B1", { formula: 'MATCH("app*",A1:A3,0)', result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B1")).toBe(1);
    });

    it('MATCH("?????", {"hi","hello","hey"}, 0) should return 2', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", "hi");
      Cell.setValue(ws, "A2", "hello");
      Cell.setValue(ws, "A3", "hey");

      // MATCH with ? wildcard — "?????" matches "hello" (5 chars) at position 2
      Cell.setValue(ws, "B1", { formula: 'MATCH("?????",A1:A3,0)', result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B1")).toBe(2);
    });
  });

  // ==========================================================================
  // 18. INDIRECT Function
  // ==========================================================================
  describe("INDIRECT function", () => {
    it('INDIRECT("A1") should return the value of cell A1', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", 42);
      Cell.setValue(ws, "B1", { formula: 'INDIRECT("A1")', result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B1")).toBe(42);
    });

    it('INDIRECT("R1C1", FALSE) should return the value using R1C1 style', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Row 1, Col 1 = A1
      Cell.setValue(ws, "A1", 42);
      Cell.setValue(ws, "B1", { formula: 'INDIRECT("R1C1",FALSE)', result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B1")).toBe(42);
    });
  });

  // ==========================================================================
  // 19. OFFSET Function
  // ==========================================================================
  describe("OFFSET function", () => {
    it("SUM(OFFSET(A1, 1, 0, 3, 1)) should sum A2:A4", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", 0);
      Cell.setValue(ws, "A2", 10);
      Cell.setValue(ws, "A3", 20);
      Cell.setValue(ws, "A4", 30);

      Cell.setValue(ws, "C1", { formula: "SUM(OFFSET(A1,1,0,3,1))", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "C1")).toBe(60);
    });
  });

  // ==========================================================================
  // 20. @ Implicit Intersection Operator
  // ==========================================================================
  describe("@ implicit intersection operator", () => {
    it("=@A1:A3 in row 2 should pick A2's value", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", 100);
      Cell.setValue(ws, "A2", 200);
      Cell.setValue(ws, "A3", 300);

      // The @ operator forces implicit intersection — picks value from same row
      Cell.setValue(ws, "B2", { formula: "@A1:A3", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "B2")).toBe(200);
    });
  });

  // ==========================================================================
  // 21. Trigonometric Functions
  // ==========================================================================
  describe("trigonometric functions", () => {
    it("SIN(0) should be 0", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "SIN(0)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(0);
    });

    it("COS(0) should be 1", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "COS(0)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(1);
    });

    it("TAN(PI()/4) should be approximately 1", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "TAN(PI()/4)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBeCloseTo(1, 10);
    });

    it("ATAN2(1,1) should be approximately PI/4", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "ATAN2(1,1)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBeCloseTo(Math.PI / 4, 10);
    });
  });

  // ==========================================================================
  // 22. MAP / REDUCE / SCAN
  // ==========================================================================
  describe("MAP / REDUCE / SCAN", () => {
    it("MAP({1,2,3}, LAMBDA(x, x*2)) should return {2,4,6}", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: "MAP({1,2,3},LAMBDA(x,x*2))",
        result: 0,
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      });

      calculateFormulas(wb);

      // Dynamic array spills horizontally: A1=2, B1=4, C1=6
      expect(Cell.getResult(ws, "A1")).toBe(2);
      expect(Cell.getValue(ws, "B1")).toBe(4);
      expect(Cell.getValue(ws, "C1")).toBe(6);
    });

    it("REDUCE(0, {1,2,3,4}, LAMBDA(a,b, a+b)) should return 10", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: "REDUCE(0,{1,2,3,4},LAMBDA(a,b,a+b))",
        result: 0
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(10);
    });
  });

  // ==========================================================================
  // 23. Lotus 1-2-3 Bug (DATE(1900,2,29))
  // ==========================================================================
  describe("Lotus 1-2-3 bug", () => {
    it("DATE(1900,2,29) should return serial number 60", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "DATE(1900,2,29)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(60);
    });
  });

  // ==========================================================================
  // 24. TEXT Format Codes
  // ==========================================================================
  describe("TEXT format codes", () => {
    it('TEXT(1234.5, "#,##0.00") should return "1,234.50"', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: 'TEXT(1234.5,"#,##0.00")',
        result: ""
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe("1,234.50");
    });

    it('TEXT(0.75, "0.00%") should return "75.00%"', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: 'TEXT(0.75,"0.00%")',
        result: ""
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe("75.00%");
    });

    it('TEXT(44927, "YYYY-MM-DD") should return "2023-01-01"', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Serial 44927 = 2023-01-01
      // Use uppercase format tokens — the formatter matches MM case-sensitively
      Cell.setValue(ws, "A1", {
        formula: 'TEXT(44927,"YYYY-MM-DD")',
        result: ""
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe("2023-01-01");
    });
  });

  // ==========================================================================
  // 25. Complex Engineering Functions
  // ==========================================================================
  describe("complex engineering functions", () => {
    it('COMPLEX(3, 4) should return "3+4i"', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "COMPLEX(3,4)", result: "" });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe("3+4i");
    });

    it('IMABS("3+4i") should return 5', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: 'IMABS("3+4i")', result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(5);
    });
  });

  // ==========================================================================
  // 26. Statistical Distributions
  // ==========================================================================
  describe("statistical distributions", () => {
    it("POISSON.DIST(3, 5, TRUE) should be approximately 0.2650", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: "POISSON.DIST(3,5,TRUE)",
        result: 0
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBeCloseTo(0.265, 3);
    });

    it("BINOM.DIST(3, 10, 0.5, FALSE) should be approximately 0.1172", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: "BINOM.DIST(3,10,0.5,FALSE)",
        result: 0
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBeCloseTo(0.1172, 3);
    });
  });

  // ==========================================================================
  // 27. TRANSPOSE
  // ==========================================================================
  describe("TRANSPOSE", () => {
    it("TRANSPOSE({1,2,3;4,5,6}) should swap rows and columns", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // {1,2,3;4,5,6} is a 2×3 matrix. Transposed → 3×2 matrix.
      Cell.setValue(ws, "A1", {
        formula: "TRANSPOSE({1,2,3;4,5,6})",
        result: 0,
        shareType: "array",
        ref: "A1:B3",
        isDynamicArray: true
      });

      calculateFormulas(wb);

      // Transposed:
      // A1=1, B1=4
      // A2=2, B2=5
      // A3=3, B3=6
      expect(Cell.getResult(ws, "A1")).toBe(1);
      expect(Cell.getValue(ws, "B1")).toBe(4);
      expect(Cell.getValue(ws, "A2")).toBe(2);
      expect(Cell.getValue(ws, "B2")).toBe(5);
      expect(Cell.getValue(ws, "A3")).toBe(3);
      expect(Cell.getValue(ws, "B3")).toBe(6);
    });
  });

  // ==========================================================================
  // 28. Named LAMBDA
  // ==========================================================================
  describe("named LAMBDA", () => {
    it("a defined name DOUBLE pointing to a LAMBDA cell, =DOUBLE(5) should return 10", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Put the LAMBDA formula in a helper cell (Z1) to hold the lambda value
      Cell.setValue(ws, "Z1", { formula: "LAMBDA(x,x*2)", result: 0 });

      // Register the defined name "DOUBLE" → Sheet1!$Z$1
      definedNamesAdd(getDefinedNames(wb), "Sheet1!$Z$1", "DOUBLE");

      // Use the named lambda: =DOUBLE(5)
      Cell.setValue(ws, "A1", { formula: "DOUBLE(5)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(10);
    });
  });

  // ==========================================================================
  // 29. MAKEARRAY
  // ==========================================================================
  describe("MAKEARRAY", () => {
    it("MAKEARRAY(2, 3, LAMBDA(r,c, r*10+c)) should return [[11,12,13],[21,22,23]]", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: "MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))",
        result: 0,
        shareType: "array",
        ref: "A1:C2",
        isDynamicArray: true
      });

      calculateFormulas(wb);

      // Row 1: A1=11, B1=12, C1=13
      // Row 2: A2=21, B2=22, C2=23
      expect(Cell.getResult(ws, "A1")).toBe(11);
      expect(Cell.getValue(ws, "B1")).toBe(12);
      expect(Cell.getValue(ws, "C1")).toBe(13);
      expect(Cell.getValue(ws, "A2")).toBe(21);
      expect(Cell.getValue(ws, "B2")).toBe(22);
      expect(Cell.getValue(ws, "C2")).toBe(23);
    });
  });

  // ==========================================================================
  // 30. Empty/Omitted Arguments (Round 3 — CRITICAL fix)
  // ==========================================================================
  describe("empty/omitted arguments", () => {
    it("IF(TRUE,,0) should return blank (omitted second arg)", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "IF(TRUE,,0)", result: 0 });

      calculateFormulas(wb);

      // Omitted argument → BLANK. In R5 we changed the writeback path to
      // surface BLANK as `undefined` (previously it was coerced to 0,
      // which collapsed "blank" and "literal zero"). Either representation
      // is acceptable — Excel itself displays it as an empty cell when
      // formatted as text and as 0 in numeric contexts.
      const result = Cell.getResult(ws, "A1");
      expect(result === null || result === undefined || result === 0).toBe(true);
    });

    it("IF(FALSE,,5) should return 5", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "IF(FALSE,,5)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe(5);
    });

    it("VLOOKUP with trailing comma (omitted 4th arg) should use default exact match", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Set up lookup table: A1:B3
      Cell.setValue(ws, "A1", "x");
      Cell.setValue(ws, "B1", 100);
      Cell.setValue(ws, "A2", "y");
      Cell.setValue(ws, "B2", 200);
      Cell.setValue(ws, "A3", "z");
      Cell.setValue(ws, "B3", 300);

      // VLOOKUP("x", A1:B3, 2,) — trailing comma = omitted 4th arg
      Cell.setValue(ws, "D1", { formula: 'VLOOKUP("x",A1:B3,2,)', result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "D1")).toBe(100);
    });
  });

  // ==========================================================================
  // 31. Infinity → #NUM! (Round 3 — HIGH fix)
  // ==========================================================================
  describe("infinity to #NUM!", () => {
    it("9.99E+307*10 should return #NUM!", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "9.99E+307*10", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toEqual({ error: "#NUM!" });
    });

    it("POWER(10, 309) should return #NUM!", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: "POWER(10,309)", result: 0 });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toEqual({ error: "#NUM!" });
    });
  });

  // ==========================================================================
  // 32. Boolean Concatenation (Round 3 — MEDIUM fix)
  // ==========================================================================
  describe("boolean concatenation", () => {
    it('TRUE&" value" should return "TRUE value"', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: 'TRUE&" value"', result: "" });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe("TRUE value");
    });

    it('FALSE&"" should return "FALSE"', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", { formula: 'FALSE&""', result: "" });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe("FALSE");
    });
  });

  // ==========================================================================
  // 33. UNIQUE with exactly_once (Round 3)
  // ==========================================================================
  describe("UNIQUE with exactly_once", () => {
    it("UNIQUE({1;2;1;3;2}, FALSE, TRUE) should return only {3}", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: "UNIQUE({1;2;1;3;2},FALSE,TRUE)",
        result: 0,
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      });

      calculateFormulas(wb);

      // Only 3 appears exactly once
      expect(Cell.getResult(ws, "A1")).toBe(3);
    });
  });

  // ==========================================================================
  // 34. SORT by_col (Round 3)
  // ==========================================================================
  describe("SORT by_col", () => {
    it("SORT({3,1,2}, 1, 1, TRUE) should sort columns: {1,2,3}", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: "SORT({3,1,2},1,1,TRUE)",
        result: 0,
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      });

      calculateFormulas(wb);

      // Columns sorted ascending: 1, 2, 3
      expect(Cell.getResult(ws, "A1")).toBe(1);
      expect(Cell.getValue(ws, "B1")).toBe(2);
      expect(Cell.getValue(ws, "C1")).toBe(3);
    });
  });

  // ==========================================================================
  // 35. TEXTBEFORE / TEXTAFTER / TEXTSPLIT (Round 3)
  // ==========================================================================
  describe("TEXTBEFORE / TEXTAFTER / TEXTSPLIT", () => {
    it('TEXTBEFORE("hello-world", "-") should return "hello"', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: 'TEXTBEFORE("hello-world","-")',
        result: ""
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe("hello");
    });

    it('TEXTAFTER("hello-world", "-") should return "world"', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: 'TEXTAFTER("hello-world","-")',
        result: ""
      });

      calculateFormulas(wb);

      expect(Cell.getResult(ws, "A1")).toBe("world");
    });

    it('TEXTSPLIT("a,b,c", ",") should return array ["a","b","c"]', () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      Cell.setValue(ws, "A1", {
        formula: 'TEXTSPLIT("a,b,c",",")',
        result: "",
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      });

      calculateFormulas(wb);

      // TEXTSPLIT spills horizontally: A1="a", B1="b", C1="c"
      expect(Cell.getResult(ws, "A1")).toBe("a");
      expect(Cell.getValue(ws, "B1")).toBe("b");
      expect(Cell.getValue(ws, "C1")).toBe("c");
    });
  });
});
