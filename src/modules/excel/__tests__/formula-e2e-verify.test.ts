import { Workbook } from "@excel/workbook";
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
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Set up: TaxRate = 0.08 in B1
      ws.getCell("B1").value = 0.08;
      // Register the defined name: "TaxRate" → Sheet1!$B$1
      wb.definedNames.add("Sheet1!$B$1", "TaxRate");

      // A1 = 100 (the base amount)
      ws.getCell("A1").value = 100;

      // Formula uses the named range: =TaxRate * A1
      ws.getCell("C1").value = { formula: "TaxRate*A1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe(8);
    });

    it("should resolve a named range pointing to a multi-cell range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      // Define "MyData" as Sheet1!$A$1:$A$3
      wb.definedNames.add("Sheet1!$A$1", "MyData");
      wb.definedNames.add("Sheet1!$A$2", "MyData");
      wb.definedNames.add("Sheet1!$A$3", "MyData");

      ws.getCell("B1").value = { formula: "SUM(MyData)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(60);
    });
  });

  // ==========================================================================
  // 2. Whole Column/Row References
  // ==========================================================================
  describe("whole column/row references", () => {
    it("SUM(A:A) should sum all data in column A", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("A4").value = 40;
      ws.getCell("A5").value = 50;

      // Place formula in a different column to avoid self-reference
      ws.getCell("B1").value = { formula: "SUM(A:A)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(150);
    });

    it("SUM(1:1) should sum all data in row 1", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = 2;
      ws.getCell("C1").value = 3;

      // Place formula in row 2 to avoid self-reference
      ws.getCell("A2").value = { formula: "SUM(1:1)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(6);
    });
  });

  // ==========================================================================
  // 3. XLOOKUP
  // ==========================================================================
  describe("XLOOKUP", () => {
    it("should find a value by exact match lookup", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Lookup table: fruits and prices
      ws.getCell("A1").value = "Apple";
      ws.getCell("A2").value = "Banana";
      ws.getCell("A3").value = "Cherry";
      ws.getCell("B1").value = 1.5;
      ws.getCell("B2").value = 0.75;
      ws.getCell("B3").value = 3.0;

      // XLOOKUP("Banana", A1:A3, B1:B3) should return 0.75
      ws.getCell("D1").value = {
        formula: 'XLOOKUP("Banana",A1:A3,B1:B3)',
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toBe(0.75);
    });

    it("should return if_not_found value when no match", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "Apple";
      ws.getCell("A2").value = "Banana";
      ws.getCell("B1").value = 1.5;
      ws.getCell("B2").value = 0.75;

      // XLOOKUP("Mango", A1:A2, B1:B2, "Not found")
      ws.getCell("D1").value = {
        formula: 'XLOOKUP("Mango",A1:A2,B1:B2,"Not found")',
        result: ""
      };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toBe("Not found");
    });
  });

  // ==========================================================================
  // 4. Financial Functions
  // ==========================================================================
  describe("financial functions", () => {
    it("PMT should compute monthly mortgage payment", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // PMT(rate, nper, pv) = PMT(0.08/12, 360, 200000)
      // Expected: approximately -1467.53
      ws.getCell("A1").value = { formula: "PMT(0.08/12,360,200000)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(-1467.53, 1);
    });

    it("FV should compute future value of annuity", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // FV(rate, nper, pmt) = FV(0.06/12, 120, -200)
      // Expected: approximately 32775.87
      ws.getCell("A1").value = { formula: "FV(0.06/12,120,-200)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(32775.87, 0);
    });
  });

  // ==========================================================================
  // 5. Structured References (Table[Column])
  // ==========================================================================
  describe("structured references", () => {
    it("should resolve Table[Column] references in formulas", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Create a Table named "Products" with columns Name and Price
      ws.addTable({
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
      ws.getCell("D1").value = { formula: "SUM(Products[Price])", result: 0 };

      wb.calculateFormulas();

      // 1.5 + 0.75 + 3.0 = 5.25
      expect(ws.getCell("D1").result).toBeCloseTo(5.25);
    });
  });

  // ==========================================================================
  // 6. Dynamic Array Spill
  // ==========================================================================
  describe("dynamic array spill", () => {
    it("FILTER should spill filtered values to adjacent cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Data: names in A1:A5, include flags in B1:B5
      ws.getCell("A1").value = "Apple";
      ws.getCell("A2").value = "Banana";
      ws.getCell("A3").value = "Cherry";
      ws.getCell("A4").value = "Date";
      ws.getCell("A5").value = "Elderberry";
      ws.getCell("B1").value = 1; // TRUE
      ws.getCell("B2").value = 0; // FALSE
      ws.getCell("B3").value = 1; // TRUE
      ws.getCell("B4").value = 0; // FALSE
      ws.getCell("B5").value = 1; // TRUE

      // Dynamic array formula with isDynamicArray flag
      ws.getCell("D1").value = {
        formula: "_xlfn._xlws.FILTER(A1:A5,B1:B5)",
        result: "",
        shareType: "array",
        ref: "D1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Should spill: D1=Apple, D2=Cherry, D3=Elderberry
      expect(ws.getCell("D1").result).toBe("Apple");
      expect(ws.getCell("D2").value).toBe("Cherry");
      expect(ws.getCell("D3").value).toBe("Elderberry");
    });
  });

  // ==========================================================================
  // 7. Implicit Intersection
  // ==========================================================================
  describe("implicit intersection", () => {
    it("should pick value from same row when range used in scalar context", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Column data A1:A3
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;

      // Formulas in B1:B3 all reference A1:A3 but in scalar context
      // Each should pick the value from its own row via implicit intersection
      ws.getCell("B1").value = { formula: "A1:A3*2", result: 0 };
      ws.getCell("B2").value = { formula: "A1:A3*2", result: 0 };
      ws.getCell("B3").value = { formula: "A1:A3*2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(20); // 10 * 2
      expect(ws.getCell("B2").result).toBe(40); // 20 * 2
      expect(ws.getCell("B3").result).toBe(60); // 30 * 2
    });
  });

  // ==========================================================================
  // 8. CSE Array Formula
  // ==========================================================================
  describe("CSE array formula", () => {
    it("{=A1:A3*10} should distribute across B1:B3", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("A3").value = 3;

      // CSE array formula: {=A1:A3*10} with ref="B1:B3"
      ws.getCell("B1").value = {
        formula: "A1:A3*10",
        result: 0,
        shareType: "array",
        ref: "B1:B3"
      };
      ws.getCell("B2").value = {
        formula: "A1:A3*10",
        result: 0,
        shareType: "array"
      };
      ws.getCell("B3").value = {
        formula: "A1:A3*10",
        result: 0,
        shareType: "array"
      };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(10); // 1*10
      expect(ws.getCell("B2").result).toBe(20); // 2*10
      expect(ws.getCell("B3").result).toBe(30); // 3*10
    });
  });

  // ==========================================================================
  // 9. 3D References
  // ==========================================================================
  describe("3D references", () => {
    it("SUM(Sheet1:Sheet3!A1) should sum A1 across three sheets", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");
      const ws3 = wb.addWorksheet("Sheet3");
      const summary = wb.addWorksheet("Summary");

      ws1.getCell("A1").value = 10;
      ws2.getCell("A1").value = 20;
      ws3.getCell("A1").value = 30;

      // 3D reference formula on the Summary sheet
      summary.getCell("A1").value = {
        formula: "SUM(Sheet1:Sheet3!A1)",
        result: 0
      };

      wb.calculateFormulas();

      expect(summary.getCell("A1").result).toBe(60);
    });
  });

  // ==========================================================================
  // 10. LET Function
  // ==========================================================================
  describe("LET function", () => {
    it("LET(x, 10, y, 20, x+y) should return 30", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "LET(x,10,y,20,x+y)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(30);
    });

    it("LET should allow later bindings to reference earlier ones", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // LET(a, 5, b, a*2, a+b) → 5 + 10 = 15
      ws.getCell("A1").value = { formula: "LET(a,5,b,a*2,a+b)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(15);
    });
  });

  // ==========================================================================
  // 11. Engineering Functions
  // ==========================================================================
  describe("engineering functions", () => {
    it('BIN2DEC("1010") should return 10', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: 'BIN2DEC("1010")', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(10);
    });

    it('HEX2DEC("FF") should return 255', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: 'HEX2DEC("FF")', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(255);
    });
  });

  // ==========================================================================
  // 12. Statistical Functions: NORM.S.INV
  // ==========================================================================
  describe("statistical functions", () => {
    it("NORM.S.INV(0.975) should be approximately 1.96", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "NORM.S.INV(0.975)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(1.96, 1);
    });
  });

  // ==========================================================================
  // 13. Database Functions: DSUM
  // ==========================================================================
  describe("database functions", () => {
    it("DSUM should sum matching records based on criteria", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Database table: A1:C5 (header + 4 rows)
      //   Region | Product | Sales
      //   East   | Widget  | 100
      //   West   | Widget  | 200
      //   East   | Gadget  | 150
      //   West   | Gadget  | 300
      ws.getCell("A1").value = "Region";
      ws.getCell("B1").value = "Product";
      ws.getCell("C1").value = "Sales";
      ws.getCell("A2").value = "East";
      ws.getCell("B2").value = "Widget";
      ws.getCell("C2").value = 100;
      ws.getCell("A3").value = "West";
      ws.getCell("B3").value = "Widget";
      ws.getCell("C3").value = 200;
      ws.getCell("A4").value = "East";
      ws.getCell("B4").value = "Gadget";
      ws.getCell("C4").value = 150;
      ws.getCell("A5").value = "West";
      ws.getCell("B5").value = "Gadget";
      ws.getCell("C5").value = 300;

      // Criteria table: E1:E2 (header + 1 criterion)
      // Region = "East"
      ws.getCell("E1").value = "Region";
      ws.getCell("E2").value = "East";

      // DSUM(database, "Sales", criteria)
      // Should sum Sales where Region="East" → 100 + 150 = 250
      ws.getCell("G1").value = {
        formula: 'DSUM(A1:C5,"Sales",E1:E2)',
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("G1").result).toBe(250);
    });
  });

  // ==========================================================================
  // 14. YEARFRAC
  // ==========================================================================
  describe("YEARFRAC", () => {
    it("YEARFRAC(DATE(2020,1,1), DATE(2020,7,1), 0) should be approximately 0.5", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // US 30/360 basis (basis=0)
      ws.getCell("A1").value = {
        formula: "YEARFRAC(DATE(2020,1,1),DATE(2020,7,1),0)",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(0.5, 2);
    });
  });

  // ==========================================================================
  // 15. IRR
  // ==========================================================================
  describe("IRR", () => {
    it("IRR({-1000, 300, 420, 680}) should be approximately 0.1665", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Put cash flows in cells and use a range reference
      ws.getCell("A1").value = -1000;
      ws.getCell("A2").value = 300;
      ws.getCell("A3").value = 420;
      ws.getCell("A4").value = 680;

      ws.getCell("B1").value = { formula: "IRR(A1:A4)", result: 0 };

      wb.calculateFormulas();

      // IRR should be approximately 0.1665 (~16.65%)
      expect(ws.getCell("B1").result).toBeCloseTo(0.1665, 2);
    });
  });

  // ==========================================================================
  // 16. INDEX with row=0 (return entire column)
  // ==========================================================================
  describe("INDEX with row=0", () => {
    it("INDEX(A1:B3, 0, 1) should return entire first column as array", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // 3x2 data block
      ws.getCell("A1").value = 10;
      ws.getCell("B1").value = 100;
      ws.getCell("A2").value = 20;
      ws.getCell("B2").value = 200;
      ws.getCell("A3").value = 30;
      ws.getCell("B3").value = 300;

      // INDEX(A1:B3, 0, 1) returns the first column as an array
      // Wrap in SUM to verify: SUM should be 10+20+30=60
      ws.getCell("D1").value = { formula: "SUM(INDEX(A1:B3,0,1))", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toBe(60);
    });

    it("INDEX with specific row and col returns a scalar", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("B1").value = 100;
      ws.getCell("A2").value = 20;
      ws.getCell("B2").value = 200;

      // INDEX(A1:B2, 2, 2) should return 200
      ws.getCell("D1").value = { formula: "INDEX(A1:B2,2,2)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toBe(200);
    });
  });

  // ==========================================================================
  // 17. MATCH with Wildcards
  // ==========================================================================
  describe("MATCH with wildcards", () => {
    it('MATCH("app*", {"apple","banana","apricot"}, 0) should return 1', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "apple";
      ws.getCell("A2").value = "banana";
      ws.getCell("A3").value = "apricot";

      // MATCH with wildcard — "app*" matches "apple" at position 1
      ws.getCell("B1").value = { formula: 'MATCH("app*",A1:A3,0)', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(1);
    });

    it('MATCH("?????", {"hi","hello","hey"}, 0) should return 2', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "hi";
      ws.getCell("A2").value = "hello";
      ws.getCell("A3").value = "hey";

      // MATCH with ? wildcard — "?????" matches "hello" (5 chars) at position 2
      ws.getCell("B1").value = { formula: 'MATCH("?????",A1:A3,0)', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(2);
    });
  });

  // ==========================================================================
  // 18. INDIRECT Function
  // ==========================================================================
  describe("INDIRECT function", () => {
    it('INDIRECT("A1") should return the value of cell A1', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 42;
      ws.getCell("B1").value = { formula: 'INDIRECT("A1")', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(42);
    });

    it('INDIRECT("R1C1", FALSE) should return the value using R1C1 style', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Row 1, Col 1 = A1
      ws.getCell("A1").value = 42;
      ws.getCell("B1").value = { formula: 'INDIRECT("R1C1",FALSE)', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(42);
    });
  });

  // ==========================================================================
  // 19. OFFSET Function
  // ==========================================================================
  describe("OFFSET function", () => {
    it("SUM(OFFSET(A1, 1, 0, 3, 1)) should sum A2:A4", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 0;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("A4").value = 30;

      ws.getCell("C1").value = { formula: "SUM(OFFSET(A1,1,0,3,1))", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe(60);
    });
  });

  // ==========================================================================
  // 20. @ Implicit Intersection Operator
  // ==========================================================================
  describe("@ implicit intersection operator", () => {
    it("=@A1:A3 in row 2 should pick A2's value", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 100;
      ws.getCell("A2").value = 200;
      ws.getCell("A3").value = 300;

      // The @ operator forces implicit intersection — picks value from same row
      ws.getCell("B2").value = { formula: "@A1:A3", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B2").result).toBe(200);
    });
  });

  // ==========================================================================
  // 21. Trigonometric Functions
  // ==========================================================================
  describe("trigonometric functions", () => {
    it("SIN(0) should be 0", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "SIN(0)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(0);
    });

    it("COS(0) should be 1", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "COS(0)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(1);
    });

    it("TAN(PI()/4) should be approximately 1", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "TAN(PI()/4)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(1, 10);
    });

    it("ATAN2(1,1) should be approximately PI/4", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "ATAN2(1,1)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(Math.PI / 4, 10);
    });
  });

  // ==========================================================================
  // 22. MAP / REDUCE / SCAN
  // ==========================================================================
  describe("MAP / REDUCE / SCAN", () => {
    it("MAP({1,2,3}, LAMBDA(x, x*2)) should return {2,4,6}", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: "MAP({1,2,3},LAMBDA(x,x*2))",
        result: 0,
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Dynamic array spills horizontally: A1=2, B1=4, C1=6
      expect(ws.getCell("A1").result).toBe(2);
      expect(ws.getCell("B1").value).toBe(4);
      expect(ws.getCell("C1").value).toBe(6);
    });

    it("REDUCE(0, {1,2,3,4}, LAMBDA(a,b, a+b)) should return 10", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: "REDUCE(0,{1,2,3,4},LAMBDA(a,b,a+b))",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(10);
    });
  });

  // ==========================================================================
  // 23. Lotus 1-2-3 Bug (DATE(1900,2,29))
  // ==========================================================================
  describe("Lotus 1-2-3 bug", () => {
    it("DATE(1900,2,29) should return serial number 60", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "DATE(1900,2,29)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(60);
    });
  });

  // ==========================================================================
  // 24. TEXT Format Codes
  // ==========================================================================
  describe("TEXT format codes", () => {
    it('TEXT(1234.5, "#,##0.00") should return "1,234.50"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: 'TEXT(1234.5,"#,##0.00")',
        result: ""
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe("1,234.50");
    });

    it('TEXT(0.75, "0.00%") should return "75.00%"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: 'TEXT(0.75,"0.00%")',
        result: ""
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe("75.00%");
    });

    it('TEXT(44927, "YYYY-MM-DD") should return "2023-01-01"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Serial 44927 = 2023-01-01
      // Use uppercase format tokens — the formatter matches MM case-sensitively
      ws.getCell("A1").value = {
        formula: 'TEXT(44927,"YYYY-MM-DD")',
        result: ""
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe("2023-01-01");
    });
  });

  // ==========================================================================
  // 25. Complex Engineering Functions
  // ==========================================================================
  describe("complex engineering functions", () => {
    it('COMPLEX(3, 4) should return "3+4i"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "COMPLEX(3,4)", result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe("3+4i");
    });

    it('IMABS("3+4i") should return 5', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: 'IMABS("3+4i")', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(5);
    });
  });

  // ==========================================================================
  // 26. Statistical Distributions
  // ==========================================================================
  describe("statistical distributions", () => {
    it("POISSON.DIST(3, 5, TRUE) should be approximately 0.2650", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: "POISSON.DIST(3,5,TRUE)",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(0.265, 3);
    });

    it("BINOM.DIST(3, 10, 0.5, FALSE) should be approximately 0.1172", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: "BINOM.DIST(3,10,0.5,FALSE)",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(0.1172, 3);
    });
  });

  // ==========================================================================
  // 27. TRANSPOSE
  // ==========================================================================
  describe("TRANSPOSE", () => {
    it("TRANSPOSE({1,2,3;4,5,6}) should swap rows and columns", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // {1,2,3;4,5,6} is a 2×3 matrix. Transposed → 3×2 matrix.
      ws.getCell("A1").value = {
        formula: "TRANSPOSE({1,2,3;4,5,6})",
        result: 0,
        shareType: "array",
        ref: "A1:B3",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Transposed:
      // A1=1, B1=4
      // A2=2, B2=5
      // A3=3, B3=6
      expect(ws.getCell("A1").result).toBe(1);
      expect(ws.getCell("B1").value).toBe(4);
      expect(ws.getCell("A2").value).toBe(2);
      expect(ws.getCell("B2").value).toBe(5);
      expect(ws.getCell("A3").value).toBe(3);
      expect(ws.getCell("B3").value).toBe(6);
    });
  });

  // ==========================================================================
  // 28. Named LAMBDA
  // ==========================================================================
  describe("named LAMBDA", () => {
    it("a defined name DOUBLE pointing to a LAMBDA cell, =DOUBLE(5) should return 10", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Put the LAMBDA formula in a helper cell (Z1) to hold the lambda value
      ws.getCell("Z1").value = { formula: "LAMBDA(x,x*2)", result: 0 };

      // Register the defined name "DOUBLE" → Sheet1!$Z$1
      wb.definedNames.add("Sheet1!$Z$1", "DOUBLE");

      // Use the named lambda: =DOUBLE(5)
      ws.getCell("A1").value = { formula: "DOUBLE(5)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(10);
    });
  });

  // ==========================================================================
  // 29. MAKEARRAY
  // ==========================================================================
  describe("MAKEARRAY", () => {
    it("MAKEARRAY(2, 3, LAMBDA(r,c, r*10+c)) should return [[11,12,13],[21,22,23]]", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: "MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))",
        result: 0,
        shareType: "array",
        ref: "A1:C2",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Row 1: A1=11, B1=12, C1=13
      // Row 2: A2=21, B2=22, C2=23
      expect(ws.getCell("A1").result).toBe(11);
      expect(ws.getCell("B1").value).toBe(12);
      expect(ws.getCell("C1").value).toBe(13);
      expect(ws.getCell("A2").value).toBe(21);
      expect(ws.getCell("B2").value).toBe(22);
      expect(ws.getCell("C2").value).toBe(23);
    });
  });

  // ==========================================================================
  // 30. Empty/Omitted Arguments (Round 3 — CRITICAL fix)
  // ==========================================================================
  describe("empty/omitted arguments", () => {
    it("IF(TRUE,,0) should return null (omitted second arg)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "IF(TRUE,,0)", result: 0 };

      wb.calculateFormulas();

      // Omitted argument → null; in Excel this displays as 0 in numeric context
      const result = ws.getCell("A1").result;
      expect(result === null || result === 0).toBe(true);
    });

    it("IF(FALSE,,5) should return 5", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "IF(FALSE,,5)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(5);
    });

    it("VLOOKUP with trailing comma (omitted 4th arg) should use default exact match", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Set up lookup table: A1:B3
      ws.getCell("A1").value = "x";
      ws.getCell("B1").value = 100;
      ws.getCell("A2").value = "y";
      ws.getCell("B2").value = 200;
      ws.getCell("A3").value = "z";
      ws.getCell("B3").value = 300;

      // VLOOKUP("x", A1:B3, 2,) — trailing comma = omitted 4th arg
      ws.getCell("D1").value = { formula: 'VLOOKUP("x",A1:B3,2,)', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toBe(100);
    });
  });

  // ==========================================================================
  // 31. Infinity → #NUM! (Round 3 — HIGH fix)
  // ==========================================================================
  describe("infinity to #NUM!", () => {
    it("9.99E+307*10 should return #NUM!", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "9.99E+307*10", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toEqual({ error: "#NUM!" });
    });

    it("POWER(10, 309) should return #NUM!", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "POWER(10,309)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toEqual({ error: "#NUM!" });
    });
  });

  // ==========================================================================
  // 32. Boolean Concatenation (Round 3 — MEDIUM fix)
  // ==========================================================================
  describe("boolean concatenation", () => {
    it('TRUE&" value" should return "TRUE value"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: 'TRUE&" value"', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe("TRUE value");
    });

    it('FALSE&"" should return "FALSE"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: 'FALSE&""', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe("FALSE");
    });
  });

  // ==========================================================================
  // 33. UNIQUE with exactly_once (Round 3)
  // ==========================================================================
  describe("UNIQUE with exactly_once", () => {
    it("UNIQUE({1;2;1;3;2}, FALSE, TRUE) should return only {3}", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: "UNIQUE({1;2;1;3;2},FALSE,TRUE)",
        result: 0,
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Only 3 appears exactly once
      expect(ws.getCell("A1").result).toBe(3);
    });
  });

  // ==========================================================================
  // 34. SORT by_col (Round 3)
  // ==========================================================================
  describe("SORT by_col", () => {
    it("SORT({3,1,2}, 1, 1, TRUE) should sort columns: {1,2,3}", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: "SORT({3,1,2},1,1,TRUE)",
        result: 0,
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Columns sorted ascending: 1, 2, 3
      expect(ws.getCell("A1").result).toBe(1);
      expect(ws.getCell("B1").value).toBe(2);
      expect(ws.getCell("C1").value).toBe(3);
    });
  });

  // ==========================================================================
  // 35. TEXTBEFORE / TEXTAFTER / TEXTSPLIT (Round 3)
  // ==========================================================================
  describe("TEXTBEFORE / TEXTAFTER / TEXTSPLIT", () => {
    it('TEXTBEFORE("hello-world", "-") should return "hello"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: 'TEXTBEFORE("hello-world","-")',
        result: ""
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe("hello");
    });

    it('TEXTAFTER("hello-world", "-") should return "world"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: 'TEXTAFTER("hello-world","-")',
        result: ""
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe("world");
    });

    it('TEXTSPLIT("a,b,c", ",") should return array ["a","b","c"]', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: 'TEXTSPLIT("a,b,c",",")',
        result: "",
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // TEXTSPLIT spills horizontally: A1="a", B1="b", C1="c"
      expect(ws.getCell("A1").result).toBe("a");
      expect(ws.getCell("B1").value).toBe("b");
      expect(ws.getCell("C1").value).toBe("c");
    });
  });
});
