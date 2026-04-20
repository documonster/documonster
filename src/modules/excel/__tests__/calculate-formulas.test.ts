import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

describe("calculateFormulas", () => {
  // ==========================================================================
  // Basic Arithmetic
  // ==========================================================================
  describe("basic arithmetic", () => {
    it("should calculate simple addition", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = { formula: "A1+A2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A3").result).toBe(30);
    });

    it("should calculate subtraction, multiplication, division", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 100;
      ws.getCell("A2").value = 25;
      ws.getCell("B1").value = { formula: "A1-A2", result: 0 };
      ws.getCell("B2").value = { formula: "A1*A2", result: 0 };
      ws.getCell("B3").value = { formula: "A1/A2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(75);
      expect(ws.getCell("B2").result).toBe(2500);
      expect(ws.getCell("B3").result).toBe(4);
    });

    it("should handle exponentiation", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 2;
      ws.getCell("A2").value = { formula: "A1^10", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(1024);
    });

    it("should respect operator precedence", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 2;
      ws.getCell("A2").value = 3;
      ws.getCell("A3").value = 4;
      ws.getCell("B1").value = { formula: "A1+A2*A3", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(14); // 2 + (3*4) = 14
    });

    it("should handle parenthesized expressions", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 2;
      ws.getCell("A2").value = 3;
      ws.getCell("A3").value = 4;
      ws.getCell("B1").value = { formula: "(A1+A2)*A3", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(20); // (2+3)*4 = 20
    });

    it("should handle division by zero", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 0;
      ws.getCell("A3").value = { formula: "A1/A2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A3").result).toEqual({ error: "#DIV/0!" });
    });
  });

  // ==========================================================================
  // SUM and Aggregate Functions
  // ==========================================================================
  describe("SUM and aggregate functions", () => {
    it("should calculate SUM over a range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("A4").value = { formula: "SUM(A1:A3)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A4").result).toBe(60);
    });

    it("should calculate SUM with multiple arguments", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("B1").value = 20;
      ws.getCell("C1").value = { formula: "SUM(A1,B1,5)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe(35);
    });

    it("should calculate AVERAGE", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("A4").value = { formula: "AVERAGE(A1:A3)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A4").result).toBe(20);
    });

    it("should calculate MIN and MAX", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 5;
      ws.getCell("A2").value = 15;
      ws.getCell("A3").value = 10;
      ws.getCell("B1").value = { formula: "MIN(A1:A3)", result: 0 };
      ws.getCell("B2").value = { formula: "MAX(A1:A3)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(5);
      expect(ws.getCell("B2").result).toBe(15);
    });

    it("should calculate COUNT and COUNTA", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = "hello";
      ws.getCell("A3").value = 30;
      // A4 left empty
      ws.getCell("B1").value = { formula: "COUNT(A1:A4)", result: 0 };
      ws.getCell("B2").value = { formula: "COUNTA(A1:A4)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(2); // Only numbers
      expect(ws.getCell("B2").result).toBe(3); // Non-empty cells
    });

    it("should calculate PRODUCT", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 2;
      ws.getCell("A2").value = 3;
      ws.getCell("A3").value = 5;
      ws.getCell("A4").value = { formula: "PRODUCT(A1:A3)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A4").result).toBe(30);
    });
  });

  // ==========================================================================
  // The Issue #140 Scenario
  // ==========================================================================
  describe("issue #140: formula values after cell modification", () => {
    it("should recalculate SUM after modifying referenced cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Setup: create a SUM formula with initial values
      ws.getCell("B1").value = 0;
      ws.getCell("B2").value = 0;
      ws.getCell("B3").value = 0;
      ws.getCell("B4").value = { formula: "SUM(B1:B3)", result: 0 };

      // Modify values (simulating the user's code in the issue)
      ws.getCell("B1").value = 1;
      ws.getCell("B2").value = 2;
      ws.getCell("B3").value = 3;

      wb.calculateFormulas();

      expect(ws.getCell("B4").result).toBe(6);
    });

    it("should handle multiple columns with formulas", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Setup column D with SUM formula
      ws.getCell("D1").value = 0;
      ws.getCell("D2").value = 0;
      ws.getCell("D3").value = 0;
      ws.getCell("D4").value = { formula: "SUM(D1:D3)", result: 0 };

      // Setup column E with SUM formula
      ws.getCell("E1").value = 0;
      ws.getCell("E2").value = 0;
      ws.getCell("E3").value = 0;
      ws.getCell("E4").value = { formula: "SUM(E1:E3)", result: 0 };

      // Modify values (as in the issue)
      ws.getCell("D1").value = 10.5;
      ws.getCell("D2").value = 23.75;
      ws.getCell("D3").value = 7.001;
      ws.getCell("E1").value = 100;
      ws.getCell("E2").value = 3.14;
      ws.getCell("E3").value = 99.99;

      wb.calculateFormulas();

      expect(ws.getCell("D4").result).toBeCloseTo(41.251);
      expect(ws.getCell("E4").result).toBeCloseTo(203.13);
    });
  });

  // ==========================================================================
  // Logical Functions
  // ==========================================================================
  describe("logical functions", () => {
    it("should evaluate IF", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = { formula: 'IF(A1>5,"big","small")', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe("big");
    });

    it("should evaluate nested IF", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 75;
      ws.getCell("A2").value = {
        formula: 'IF(A1>=90,"A",IF(A1>=80,"B",IF(A1>=70,"C","D")))',
        result: ""
      };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe("C");
    });

    it("should evaluate AND and OR", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = true;
      ws.getCell("A2").value = false;
      ws.getCell("B1").value = { formula: "AND(A1,A2)", result: false };
      ws.getCell("B2").value = { formula: "OR(A1,A2)", result: false };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(false);
      expect(ws.getCell("B2").result).toBe(true);
    });

    it("should evaluate IFERROR", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 0;
      ws.getCell("B1").value = { formula: 'IFERROR(A1/A2,"Error")', result: "" };
      ws.getCell("B2").value = { formula: "IFERROR(A1*A1,0)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe("Error");
      expect(ws.getCell("B2").result).toBe(100);
    });
  });

  // ==========================================================================
  // Text Functions
  // ==========================================================================
  describe("text functions", () => {
    it("should evaluate CONCATENATE", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "Hello";
      ws.getCell("A2").value = "World";
      ws.getCell("A3").value = { formula: 'CONCATENATE(A1,", ",A2,"!")', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("A3").result).toBe("Hello, World!");
    });

    it("should evaluate string concatenation operator &", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "foo";
      ws.getCell("A2").value = "bar";
      ws.getCell("A3").value = { formula: 'A1&" "&A2', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("A3").result).toBe("foo bar");
    });

    it("should evaluate LEN, LEFT, RIGHT, MID", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "Hello World";
      ws.getCell("B1").value = { formula: "LEN(A1)", result: 0 };
      ws.getCell("B2").value = { formula: "LEFT(A1,5)", result: "" };
      ws.getCell("B3").value = { formula: "RIGHT(A1,5)", result: "" };
      ws.getCell("B4").value = { formula: "MID(A1,7,5)", result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(11);
      expect(ws.getCell("B2").result).toBe("Hello");
      expect(ws.getCell("B3").result).toBe("World");
      expect(ws.getCell("B4").result).toBe("World");
    });

    it("should evaluate UPPER, LOWER, TRIM", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "  hello  WORLD  ";
      ws.getCell("B1").value = { formula: "UPPER(A1)", result: "" };
      ws.getCell("B2").value = { formula: "LOWER(A1)", result: "" };
      ws.getCell("B3").value = { formula: "TRIM(A1)", result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe("  HELLO  WORLD  ");
      expect(ws.getCell("B2").result).toBe("  hello  world  ");
      expect(ws.getCell("B3").result).toBe("hello WORLD");
    });
  });

  // ==========================================================================
  // Math Functions
  // ==========================================================================
  describe("math functions", () => {
    it("should evaluate ROUND, ROUNDUP, ROUNDDOWN", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 3.456;
      ws.getCell("B1").value = { formula: "ROUND(A1,2)", result: 0 };
      ws.getCell("B2").value = { formula: "ROUNDUP(A1,2)", result: 0 };
      ws.getCell("B3").value = { formula: "ROUNDDOWN(A1,2)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(3.46);
      expect(ws.getCell("B2").result).toBe(3.46);
      expect(ws.getCell("B3").result).toBe(3.45);
    });

    it("should evaluate ABS, SQRT, POWER, MOD", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = -16;
      ws.getCell("B1").value = { formula: "ABS(A1)", result: 0 };
      ws.getCell("B2").value = { formula: "SQRT(ABS(A1))", result: 0 };
      ws.getCell("B3").value = { formula: "POWER(2,8)", result: 0 };
      ws.getCell("B4").value = { formula: "MOD(17,5)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(16);
      expect(ws.getCell("B2").result).toBe(4);
      expect(ws.getCell("B3").result).toBe(256);
      expect(ws.getCell("B4").result).toBe(2);
    });

    it("should evaluate INT, CEILING, FLOOR", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 7.8;
      ws.getCell("B1").value = { formula: "INT(A1)", result: 0 };
      ws.getCell("B2").value = { formula: "CEILING(A1,1)", result: 0 };
      ws.getCell("B3").value = { formula: "FLOOR(A1,1)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(7);
      expect(ws.getCell("B2").result).toBe(8);
      expect(ws.getCell("B3").result).toBe(7);
    });
  });

  // ==========================================================================
  // Cross-sheet References
  // ==========================================================================
  describe("cross-sheet references", () => {
    it("should resolve references to another sheet", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Data");
      const ws2 = wb.addWorksheet("Summary");

      ws1.getCell("A1").value = 100;
      ws1.getCell("A2").value = 200;
      ws1.getCell("A3").value = 300;

      ws2.getCell("A1").value = { formula: "SUM(Data!A1:A3)", result: 0 };

      wb.calculateFormulas();

      expect(ws2.getCell("A1").result).toBe(600);
    });

    it("should resolve single cell cross-sheet reference", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");

      ws1.getCell("A1").value = 42;
      ws2.getCell("A1").value = { formula: "Sheet1!A1*2", result: 0 };

      wb.calculateFormulas();

      expect(ws2.getCell("A1").result).toBe(84);
    });
  });

  // ==========================================================================
  // Chained Dependencies
  // ==========================================================================
  describe("chained formula dependencies", () => {
    it("should handle a chain of dependent formulas", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 5;
      ws.getCell("A2").value = { formula: "A1*2", result: 0 }; // 10
      ws.getCell("A3").value = { formula: "A2+A1", result: 0 }; // 15
      ws.getCell("A4").value = { formula: "SUM(A1:A3)", result: 0 }; // 30

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(10);
      expect(ws.getCell("A3").result).toBe(15);
      expect(ws.getCell("A4").result).toBe(30);
    });

    it("should detect and handle circular references", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "A2+1", result: 0 };
      ws.getCell("A2").value = { formula: "A1+1", result: 0 };

      wb.calculateFormulas();

      // Should not hang — circular refs return 0 (Excel behavior)
      const r1 = ws.getCell("A1").result;
      const r2 = ws.getCell("A2").result;
      // Both should be numbers (circular refs get fallback 0, then evaluate)
      expect(typeof r1).toBe("number");
      expect(typeof r2).toBe("number");
    });
  });

  // ==========================================================================
  // Comparison Operators
  // ==========================================================================
  describe("comparison operators", () => {
    it("should evaluate = <> < > <= >=", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("B1").value = { formula: "A1=A2", result: false };
      ws.getCell("B2").value = { formula: "A1<>A2", result: false };
      ws.getCell("B3").value = { formula: "A1<A2", result: false };
      ws.getCell("B4").value = { formula: "A1>A2", result: false };
      ws.getCell("B5").value = { formula: "A1<=A2", result: false };
      ws.getCell("B6").value = { formula: "A1>=A2", result: false };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(false);
      expect(ws.getCell("B2").result).toBe(true);
      expect(ws.getCell("B3").result).toBe(true);
      expect(ws.getCell("B4").result).toBe(false);
      expect(ws.getCell("B5").result).toBe(true);
      expect(ws.getCell("B6").result).toBe(false);
    });
  });

  // ==========================================================================
  // Absolute References
  // ==========================================================================
  describe("absolute references", () => {
    it("should handle $A$1 style references", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 100;
      ws.getCell("B1").value = { formula: "$A$1+50", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(150);
    });
  });

  // ==========================================================================
  // Unary Operators
  // ==========================================================================
  describe("unary operators", () => {
    it("should handle unary minus", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 42;
      ws.getCell("A2").value = { formula: "-A1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(-42);
    });

    it("should handle percentage", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "50%", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(0.5);
    });
  });

  // ==========================================================================
  // VLOOKUP
  // ==========================================================================
  describe("VLOOKUP", () => {
    it("should perform exact match lookup", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "Apple";
      ws.getCell("B1").value = 1.5;
      ws.getCell("A2").value = "Banana";
      ws.getCell("B2").value = 0.75;
      ws.getCell("A3").value = "Cherry";
      ws.getCell("B3").value = 3.0;

      ws.getCell("D1").value = { formula: 'VLOOKUP("Banana",A1:B3,2,FALSE)', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toBe(0.75);
    });
  });

  // ==========================================================================
  // Information Functions
  // ==========================================================================
  describe("information functions", () => {
    it("should evaluate ISNUMBER, ISTEXT, ISBLANK", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 42;
      ws.getCell("A2").value = "text";
      // A3 left blank

      ws.getCell("B1").value = { formula: "ISNUMBER(A1)", result: false };
      ws.getCell("B2").value = { formula: "ISTEXT(A2)", result: false };
      ws.getCell("B3").value = { formula: "ISBLANK(A3)", result: false };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(true);
      expect(ws.getCell("B2").result).toBe(true);
      expect(ws.getCell("B3").result).toBe(true);
    });
  });

  // ==========================================================================
  // Constant Formulas (no cell references)
  // ==========================================================================
  describe("constant formulas", () => {
    it("should evaluate formulas with only constants", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "1+1", result: 0 };
      ws.getCell("A2").value = { formula: "PI()", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(2);
      expect(ws.getCell("A2").result).toBeCloseTo(Math.PI);
    });
  });

  // ==========================================================================
  // Unsupported Formulas (graceful fallback)
  // ==========================================================================
  describe("unsupported formulas", () => {
    it("should preserve cached result when function is not implemented", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Use a completely made-up function name that will never be implemented
      ws.getCell("A1").value = {
        formula: "XYZNONEXISTENT(1,2,3)",
        result: 42
      };

      wb.calculateFormulas();

      // The engine cannot evaluate this formula, so it must preserve
      // the original cached result (42) rather than overwriting with #NAME?
      expect(ws.getCell("A1").result).toBe(42);
    });

    it("should return #NAME? when no cached result exists", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = {
        formula: "XYZNONEXISTENT(1,2,3)"
      };

      wb.calculateFormulas();

      // No cached result to preserve — #NAME? is the only option
      expect(ws.getCell("A1").result).toEqual({ error: "#NAME?" });
    });
  });

  // ==========================================================================
  // ROW() and COLUMN()
  // ==========================================================================
  describe("ROW and COLUMN functions", () => {
    it("should evaluate ROW with cell reference", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "ROW(B5)", result: 0 };
      ws.getCell("A2").value = { formula: "COLUMN(D1)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(5);
      expect(ws.getCell("A2").result).toBe(4);
    });
  });

  // ==========================================================================
  // Empty cells in ranges
  // ==========================================================================
  describe("empty cells in ranges", () => {
    it("should treat empty cells as 0 in SUM", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      // A2 intentionally left empty
      ws.getCell("A3").value = 30;
      ws.getCell("A4").value = { formula: "SUM(A1:A3)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A4").result).toBe(40);
    });
  });

  // ==========================================================================
  // Boolean values in formulas
  // ==========================================================================
  describe("boolean values", () => {
    it("should handle TRUE and FALSE in arithmetic", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "TRUE+TRUE", result: 0 };
      ws.getCell("A2").value = { formula: "FALSE+1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(2);
      expect(ws.getCell("A2").result).toBe(1);
    });
  });

  // ==========================================================================
  // C1 fix: formula with no initial cached result
  // ==========================================================================
  describe("formula with undefined initial result", () => {
    it("should write result even when formula had no cached result", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      // Formula with NO initial result (result is undefined)
      ws.getCell("A3").value = { formula: "A1+A2" };

      expect(ws.getCell("A3").result).toBeUndefined();

      wb.calculateFormulas();

      expect(ws.getCell("A3").result).toBe(30);
    });
  });

  // ==========================================================================
  // C2 fix: recursive dependency persists results
  // ==========================================================================
  describe("recursive dependency result persistence", () => {
    it("should persist results for cells evaluated as dependencies", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 5;
      // A2 depends on A1, A3 depends on A2, A4 depends on A3
      // If A4 is evaluated first, it should trigger A3 → A2 recursively
      // and all intermediate results should be persisted
      ws.getCell("A2").value = { formula: "A1*2", result: 0 }; // 10
      ws.getCell("A3").value = { formula: "A2+3", result: 0 }; // 13
      ws.getCell("A4").value = { formula: "A3+7", result: 0 }; // 20

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(10);
      expect(ws.getCell("A3").result).toBe(13);
      expect(ws.getCell("A4").result).toBe(20);
    });
  });

  // ==========================================================================
  // I2 fix: circular reference returns 0 (not #REF!)
  // ==========================================================================
  describe("circular reference returns 0", () => {
    it("should return 0 for circular references like Excel", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "A2+1", result: 0 };
      ws.getCell("A2").value = { formula: "A1+1", result: 0 };

      wb.calculateFormulas();

      // Excel returns 0 for circular references, not #REF!
      // At least one cell in the cycle should resolve to a number
      const r1 = ws.getCell("A1").result;
      const r2 = ws.getCell("A2").result;
      // Both cells in the cycle should resolve to numbers
      expect(typeof r1).toBe("number");
      expect(typeof r2).toBe("number");
    });
  });

  // ==========================================================================
  // I6 fix: RIGHT(text, 0) returns ""
  // ==========================================================================
  describe("RIGHT edge cases", () => {
    it("should return empty string for RIGHT(text, 0)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "Hello";
      ws.getCell("B1").value = { formula: "RIGHT(A1,0)", result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe("");
    });
  });

  // ==========================================================================
  // I1 fix: session isolation (no global state pollution)
  // ==========================================================================
  describe("session isolation", () => {
    it("should correctly calculate two workbooks independently", () => {
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Sheet1");
      ws1.getCell("A1").value = 100;
      ws1.getCell("A2").value = { formula: "A1*2", result: 0 };

      const wb2 = new Workbook();
      const ws2 = wb2.addWorksheet("Sheet1");
      ws2.getCell("A1").value = 999;
      ws2.getCell("A2").value = { formula: "A1*3", result: 0 };

      wb1.calculateFormulas();
      wb2.calculateFormulas();

      expect(ws1.getCell("A2").result).toBe(200);
      expect(ws2.getCell("A2").result).toBe(2997);
    });
  });

  // ==========================================================================
  // Memoization: same formula cell is not evaluated twice
  // ==========================================================================
  describe("memoization", () => {
    it("should not produce different results when a cell is referenced multiple times", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 7;
      ws.getCell("A2").value = { formula: "A1*A1", result: 0 }; // 49
      ws.getCell("B1").value = { formula: "A2+A2", result: 0 }; // 98 (references A2 twice)

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(49);
      expect(ws.getCell("B1").result).toBe(98);
    });
  });

  // ==========================================================================
  // Conditional Functions: SUMIF, COUNTIF, AVERAGEIF, SUMIFS, COUNTIFS
  // ==========================================================================
  describe("conditional aggregate functions", () => {
    it("should evaluate SUMIF", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "Apple";
      ws.getCell("A2").value = "Banana";
      ws.getCell("A3").value = "Apple";
      ws.getCell("B1").value = 10;
      ws.getCell("B2").value = 20;
      ws.getCell("B3").value = 30;
      ws.getCell("C1").value = { formula: 'SUMIF(A1:A3,"Apple",B1:B3)', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe(40);
    });

    it("should evaluate SUMIF with operator criteria", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 5;
      ws.getCell("A2").value = 15;
      ws.getCell("A3").value = 25;
      ws.getCell("B1").value = { formula: 'SUMIF(A1:A3,">10")', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(40);
    });

    it("should evaluate COUNTIF", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "Yes";
      ws.getCell("A2").value = "No";
      ws.getCell("A3").value = "Yes";
      ws.getCell("A4").value = "Yes";
      ws.getCell("B1").value = { formula: 'COUNTIF(A1:A4,"Yes")', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(3);
    });

    it("should evaluate COUNTIFS", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "A";
      ws.getCell("A2").value = "B";
      ws.getCell("A3").value = "A";
      ws.getCell("B1").value = 10;
      ws.getCell("B2").value = 20;
      ws.getCell("B3").value = 30;
      ws.getCell("C1").value = { formula: 'COUNTIFS(A1:A3,"A",B1:B3,">5")', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe(2);
    });

    it("should evaluate AVERAGEIF", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("B1").value = { formula: 'AVERAGEIF(A1:A3,">15")', result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(25);
    });
  });

  // ==========================================================================
  // Statistical Functions: MEDIAN, LARGE, SMALL, STDEV
  // ==========================================================================
  describe("statistical functions", () => {
    it("should evaluate MEDIAN", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 5;
      ws.getCell("A3").value = 3;
      ws.getCell("A4").value = 7;
      ws.getCell("A5").value = 9;
      ws.getCell("B1").value = { formula: "MEDIAN(A1:A5)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(5);
    });

    it("should evaluate LARGE and SMALL", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 30;
      ws.getCell("A3").value = 20;
      ws.getCell("B1").value = { formula: "LARGE(A1:A3,1)", result: 0 };
      ws.getCell("B2").value = { formula: "SMALL(A1:A3,1)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(30);
      expect(ws.getCell("B2").result).toBe(10);
    });
  });

  // ==========================================================================
  // Dynamic Array Functions: FILTER, SORT, UNIQUE
  // ==========================================================================
  describe("dynamic array functions", () => {
    it("should evaluate FILTER", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "Apple";
      ws.getCell("A2").value = "Banana";
      ws.getCell("A3").value = "Cherry";
      ws.getCell("B1").value = 1;
      ws.getCell("B2").value = 0;
      ws.getCell("B3").value = 1;
      // _xlfn._xlws.FILTER — the dynamic array function
      ws.getCell("D1").value = { formula: "_xlfn._xlws.FILTER(A1:A3,B1:B3)", result: "" };

      wb.calculateFormulas();

      // FILTER returns an array; result should reflect the first element
      // since evaluateFormula unwraps to scalar
      expect(ws.getCell("D1").result).toBe("Apple");
    });

    it("should evaluate SORT", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("B1").value = { formula: "_xlfn._xlws.SORT(A1:A3)", result: 0 };

      wb.calculateFormulas();

      // Result is first element of sorted array
      expect(ws.getCell("B1").result).toBe(10);
    });

    it("should evaluate UNIQUE", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "A";
      ws.getCell("A2").value = "B";
      ws.getCell("A3").value = "A";
      ws.getCell("A4").value = "C";
      ws.getCell("B1").value = { formula: "_xlfn._xlws.UNIQUE(A1:A4)", result: "" };

      wb.calculateFormulas();

      // First unique element
      expect(ws.getCell("B1").result).toBe("A");
    });
  });

  // ==========================================================================
  // Date Functions: DATE, TIME, HOUR, MINUTE, SECOND, WEEKDAY, EOMONTH
  // ==========================================================================
  describe("date/time functions", () => {
    it("should evaluate DATE", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "YEAR(DATE(2024,3,15))", result: 0 };
      ws.getCell("A2").value = { formula: "MONTH(DATE(2024,3,15))", result: 0 };
      ws.getCell("A3").value = { formula: "DAY(DATE(2024,3,15))", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(2024);
      expect(ws.getCell("A2").result).toBe(3);
      expect(ws.getCell("A3").result).toBe(15);
    });

    it("should evaluate TIME, HOUR, MINUTE, SECOND", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "TIME(14,30,45)", result: 0 };
      ws.getCell("A2").value = { formula: "HOUR(TIME(14,30,45))", result: 0 };
      ws.getCell("A3").value = { formula: "MINUTE(TIME(14,30,45))", result: 0 };
      ws.getCell("A4").value = { formula: "SECOND(TIME(14,30,45))", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(14);
      expect(ws.getCell("A3").result).toBe(30);
      expect(ws.getCell("A4").result).toBe(45);
    });
  });

  // ==========================================================================
  // Additional Math: TRUNC, GCD, LCM, SUMSQ
  // ==========================================================================
  describe("additional math functions", () => {
    it("should evaluate TRUNC", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "TRUNC(3.789,2)", result: 0 };
      ws.getCell("A2").value = { formula: "TRUNC(-3.789)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(3.78);
      expect(ws.getCell("A2").result).toBe(-3);
    });

    it("should evaluate GCD and LCM", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "GCD(12,18)", result: 0 };
      ws.getCell("A2").value = { formula: "LCM(4,6)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(6);
      expect(ws.getCell("A2").result).toBe(12);
    });
  });

  // ==========================================================================
  // TEXT function enhancements
  // ==========================================================================
  describe("TEXT function", () => {
    it("should format percentages", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 0.75;
      ws.getCell("B1").value = { formula: 'TEXT(A1,"0%")', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe("75%");
    });

    it("should format with thousands separator", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1234567.89;
      ws.getCell("B1").value = { formula: 'TEXT(A1,"#,##0.00")', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe("1,234,567.89");
    });
  });

  // ==========================================================================
  // MATCH with approximate match
  // ==========================================================================
  describe("MATCH approximate match", () => {
    it("should find approximate match in sorted ascending data", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("A4").value = 40;
      ws.getCell("B1").value = { formula: "MATCH(25,A1:A4,1)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(2); // 20 is the largest <= 25
    });
  });

  // ==========================================================================
  // T2: IF short-circuit — should not evaluate unused branch
  // ==========================================================================
  describe("IF short-circuit", () => {
    it("should not return error from un-taken IF branch", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 0;
      // IF(TRUE, 1, 1/A1) — false branch would be #DIV/0! but should not be evaluated
      ws.getCell("B1").value = { formula: "IF(TRUE,1,1/A1)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(1);
    });

    it("should evaluate IFERROR correctly with error in first arg", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 0;
      ws.getCell("B1").value = { formula: 'IFERROR(1/A1,"safe")', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe("safe");
    });

    it("should evaluate IFERROR correctly with non-error in first arg", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 2;
      ws.getCell("B1").value = { formula: 'IFERROR(1/A1,"safe")', result: "" };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(0.5);
    });
  });

  // ==========================================================================
  // T4: SUMPRODUCT test
  // ==========================================================================
  describe("SUMPRODUCT", () => {
    it("should calculate SUMPRODUCT of two ranges", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("A3").value = 3;
      ws.getCell("B1").value = 4;
      ws.getCell("B2").value = 5;
      ws.getCell("B3").value = 6;
      // SUMPRODUCT = 1*4 + 2*5 + 3*6 = 4+10+18 = 32
      ws.getCell("C1").value = { formula: "SUMPRODUCT(A1:A3,B1:B3)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe(32);
    });
  });

  // ==========================================================================
  // T5: Shared formulas
  // ==========================================================================
  describe("shared formulas", () => {
    it("should calculate shared formulas via fillFormula", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      // B1:B3 as shared formula B1=A1*2, B2=A2*2, B3=A3*2
      ws.fillFormula("B1:B3", "A1*2", [20, 40, 60]);

      // Override A values — cached results should be stale
      ws.getCell("A1").value = 100;
      ws.getCell("A2").value = 200;
      ws.getCell("A3").value = 300;

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(200);
      expect(ws.getCell("B2").result).toBe(400);
      expect(ws.getCell("B3").result).toBe(600);
    });
  });

  // ==========================================================================
  // B3 regression: SUM should ignore booleans in ranges
  // ==========================================================================
  describe("SUM ignores booleans in ranges", () => {
    it("should not count TRUE as 1 in a range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = true;
      ws.getCell("A3").value = 20;
      ws.getCell("B1").value = { formula: "SUM(A1:A3)", result: 0 };

      wb.calculateFormulas();

      // Excel: SUM over a range ignores TRUE — result should be 30, not 31
      expect(ws.getCell("B1").result).toBe(30);
    });

    it("should count TRUE as 1 when passed as direct argument", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "SUM(10,TRUE,20)", result: 0 };

      wb.calculateFormulas();

      // Direct argument: TRUE → 1, so 10+1+20 = 31
      expect(ws.getCell("A1").result).toBe(31);
    });
  });

  // ==========================================================================
  // Unary minus vs exponentiation precedence
  // ==========================================================================
  describe("unary minus vs exponentiation", () => {
    it("should parse -1^2 as (-1)^2 = 1 (Excel's unique precedence)", () => {
      // Regression: Excel's precedence table ranks "Negation (as in –1)"
      // at rank 1 — tighter than exponentiation (rank 4). So `=-1^2` is
      // `(-1)^2 = 1`, not `-(1^2) = -1`. The engine used to parse
      // unary-minus at precedence 55 (below `^`'s 60/61), producing
      // `-(1^2)` — matching Python/C but not Excel.
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = { formula: "-1^2", result: 0 };
      ws.getCell("A2").value = { formula: "-2^3", result: 0 };
      ws.getCell("A3").value = { formula: "-2^2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(1); // (-1)^2 = 1
      expect(ws.getCell("A2").result).toBe(-8); // (-2)^3 = -8 (happens to match either way)
      expect(ws.getCell("A3").result).toBe(4); // (-2)^2 = 4 — key case
    });

    it("should parse -A1^2 as (-A1)^2 (Excel)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 3;
      ws.getCell("B1").value = { formula: "-A1^2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(9); // (-3)^2 = 9
    });
  });

  // ==========================================================================
  // Quoted sheet names with spaces
  // ==========================================================================
  describe("quoted sheet names", () => {
    it("should resolve 'Sheet Name'!A1 with spaces", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("My Data");
      const ws2 = wb.addWorksheet("Results");

      ws1.getCell("A1").value = 100;
      ws1.getCell("A2").value = 200;
      ws2.getCell("A1").value = { formula: "SUM('My Data'!A1:A2)", result: 0 };

      wb.calculateFormulas();

      expect(ws2.getCell("A1").result).toBe(300);
    });
  });

  // ==========================================================================
  // XLSX Round-trip Integration: write → read → recalculate
  // ==========================================================================
  describe("xlsx round-trip", () => {
    it("should recalculate formulas after write+read round-trip", async () => {
      // Build a workbook with formulas
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Sheet1");
      ws1.getCell("A1").value = 10;
      ws1.getCell("A2").value = 20;
      ws1.getCell("A3").value = { formula: "SUM(A1:A2)", result: 30 };
      ws1.getCell("B1").value = { formula: "A3*2", result: 60 };

      // Write to buffer
      const buffer = await wb1.xlsx.writeBuffer();

      // Read back
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);
      const ws2 = wb2.getWorksheet("Sheet1")!;

      // Modify data cells — cached results are now stale
      ws2.getCell("A1").value = 100;
      ws2.getCell("A2").value = 200;

      // Recalculate
      wb2.calculateFormulas();

      expect(ws2.getCell("A3").result).toBe(300);
      expect(ws2.getCell("B1").result).toBe(600);
    });
  });

  // ==========================================================================
  // Edge cases: empty workbook, no formulas, formula-only
  // ==========================================================================
  describe("edge cases", () => {
    it("should handle workbook with no worksheets", () => {
      const wb = new Workbook();
      expect(() => wb.calculateFormulas()).not.toThrow();
    });

    it("should handle worksheet with no formulas", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 42;
      ws.getCell("A2").value = "hello";

      expect(() => wb.calculateFormulas()).not.toThrow();
    });

    it("should handle formula referencing empty cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      // Formula references cells that don't exist at all
      ws.getCell("Z1").value = { formula: "SUM(A1:A100)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("Z1").result).toBe(0);
    });
  });

  // ==========================================================================
  // Idempotency: calling calculateFormulas multiple times
  // ==========================================================================
  describe("idempotency", () => {
    it("should produce same results when called multiple times", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 5;
      ws.getCell("A2").value = { formula: "A1*3", result: 0 };
      ws.getCell("A3").value = { formula: "A2+A1", result: 0 };

      wb.calculateFormulas();
      const r1a = ws.getCell("A2").result;
      const r1b = ws.getCell("A3").result;

      wb.calculateFormulas();
      expect(ws.getCell("A2").result).toBe(r1a);
      expect(ws.getCell("A3").result).toBe(r1b);

      // Verify actual values
      expect(r1a).toBe(15);
      expect(r1b).toBe(20);
    });
  });

  // ==========================================================================
  // Deep dependency chain (>10 levels)
  // ==========================================================================
  describe("deep dependency chain", () => {
    it("should handle 50-level deep formula chain without stack overflow", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      for (let i = 2; i <= 50; i++) {
        ws.getCell(`A${i}`).value = { formula: `A${i - 1}+1`, result: 0 };
      }

      wb.calculateFormulas();

      expect(ws.getCell("A50").result).toBe(50);
    });
  });

  // ==========================================================================
  // INDEX + MATCH combination
  // ==========================================================================
  describe("INDEX+MATCH", () => {
    it("should evaluate INDEX(MATCH()) combination", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Lookup table
      ws.getCell("A1").value = "Alpha";
      ws.getCell("A2").value = "Beta";
      ws.getCell("A3").value = "Gamma";
      ws.getCell("B1").value = 100;
      ws.getCell("B2").value = 200;
      ws.getCell("B3").value = 300;

      // INDEX(B1:B3, MATCH("Beta", A1:A3, 0))
      ws.getCell("D1").value = {
        formula: 'INDEX(B1:B3,MATCH("Beta",A1:A3,0))',
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toBe(200);
    });
  });

  // ==========================================================================
  // Complex nested: IF + AND + VLOOKUP
  // ==========================================================================
  describe("complex nested formulas", () => {
    it("should evaluate IF(AND(...), VLOOKUP(...), ...)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Lookup table
      ws.getCell("A1").value = "X";
      ws.getCell("B1").value = 10;
      ws.getCell("A2").value = "Y";
      ws.getCell("B2").value = 20;

      // Conditions
      ws.getCell("D1").value = true;
      ws.getCell("D2").value = true;

      // IF(AND(D1,D2), VLOOKUP("Y",A1:B2,2,FALSE), -1)
      ws.getCell("E1").value = {
        formula: 'IF(AND(D1,D2),VLOOKUP("Y",A1:B2,2,FALSE),-1)',
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("E1").result).toBe(20);
    });

    it("should take else branch when condition is false", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "X";
      ws.getCell("B1").value = 10;
      ws.getCell("D1").value = true;
      ws.getCell("D2").value = false; // AND will be false

      ws.getCell("E1").value = {
        formula: 'IF(AND(D1,D2),VLOOKUP("X",A1:B1,2,FALSE),-1)',
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("E1").result).toBe(-1);
    });
  });

  // ==========================================================================
  // Statistical functions: STDEV, VAR
  // ==========================================================================
  describe("STDEV and VAR", () => {
    it("should calculate STDEV.S (sample standard deviation)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 2;
      ws.getCell("A2").value = 4;
      ws.getCell("A3").value = 4;
      ws.getCell("A4").value = 4;
      ws.getCell("A5").value = 5;
      ws.getCell("A6").value = 5;
      ws.getCell("A7").value = 7;
      ws.getCell("A8").value = 9;
      ws.getCell("B1").value = { formula: "STDEV(A1:A8)", result: 0 };

      wb.calculateFormulas();

      // stdev.s of [2,4,4,4,5,5,7,9]: mean=5, Σ(xi-mean)²=32, s²=32/7≈4.571, s≈2.138
      expect(ws.getCell("B1").result).toBeCloseTo(2.138, 2);
    });

    it("should calculate VAR.S (sample variance)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 2;
      ws.getCell("A2").value = 4;
      ws.getCell("A3").value = 4;
      ws.getCell("A4").value = 4;
      ws.getCell("A5").value = 5;
      ws.getCell("A6").value = 5;
      ws.getCell("A7").value = 7;
      ws.getCell("A8").value = 9;
      ws.getCell("B1").value = { formula: "VAR(A1:A8)", result: 0 };

      wb.calculateFormulas();

      // var.s = 32/7 ≈ 4.571
      expect(ws.getCell("B1").result).toBeCloseTo(4.571, 2);
    });
  });

  // ==========================================================================
  // Date functions: DATEDIF, EOMONTH, EDATE
  // ==========================================================================
  describe("DATEDIF, EOMONTH, EDATE", () => {
    it("should calculate DATEDIF in years", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // DATEDIF(DATE(2020,1,1), DATE(2024,6,15), "Y") = 4 years
      ws.getCell("A1").value = {
        formula: 'DATEDIF(DATE(2020,1,1),DATE(2024,6,15),"Y")',
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(4);
    });

    it("should calculate EOMONTH", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // EOMONTH(DATE(2024,1,15), 2) → end of March 2024 = 2024-03-31
      // Then extract DAY to verify
      ws.getCell("A1").value = {
        formula: "DAY(EOMONTH(DATE(2024,1,15),2))",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(31); // March has 31 days
    });

    it("should calculate EDATE", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // EDATE(DATE(2024,1,15), 3) → 2024-04-15
      ws.getCell("A1").value = {
        formula: "MONTH(EDATE(DATE(2024,1,15),3))",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(4); // April
    });
  });

  // ==========================================================================
  // Dynamic Array Spill Engine
  // ==========================================================================
  describe("dynamic array spill engine", () => {
    it("should spill FILTER results to adjacent cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Data in A1:B3
      ws.getCell("A1").value = "Apple";
      ws.getCell("A2").value = "Banana";
      ws.getCell("A3").value = "Cherry";
      ws.getCell("B1").value = 1;
      ws.getCell("B2").value = 0;
      ws.getCell("B3").value = 1;

      // Dynamic array formula in D1
      ws.getCell("D1").value = {
        formula: "_xlfn._xlws.FILTER(A1:A3,B1:B3)",
        result: "",
        shareType: "array",
        ref: "D1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Should spill: D1=Apple, D2=Cherry
      expect(ws.getCell("D1").result).toBe("Apple");
      expect(ws.getCell("D2").value).toBe("Cherry");
    });

    it("should spill SORT results to adjacent cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;

      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);
    });

    it("should spill UNIQUE results to adjacent cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "A";
      ws.getCell("A2").value = "B";
      ws.getCell("A3").value = "A";
      ws.getCell("A4").value = "C";

      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.UNIQUE(A1:A4)",
        result: "",
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe("A");
      expect(ws.getCell("C2").value).toBe("B");
      expect(ws.getCell("C3").value).toBe("C");
    });

    it("should produce #SPILL! error when target cells are occupied", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;

      // Put data in C2 — this will conflict with spill from C1
      ws.getCell("C2").value = "blocked";

      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Should get #SPILL! because C2 is occupied
      expect(ws.getCell("C1").result).toEqual({ error: "#SPILL!" });
    });

    it("should spill multi-column results", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // 2-column data
      ws.getCell("A1").value = "Apple";
      ws.getCell("B1").value = 1.5;
      ws.getCell("A2").value = "Banana";
      ws.getCell("B2").value = 0.75;
      ws.getCell("A3").value = "Cherry";
      ws.getCell("B3").value = 3.0;

      // Include mask
      ws.getCell("C1").value = 1;
      ws.getCell("C2").value = 0;
      ws.getCell("C3").value = 1;

      ws.getCell("E1").value = {
        formula: "_xlfn._xlws.FILTER(A1:B3,C1:C3)",
        result: "",
        shareType: "array",
        ref: "E1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Should spill 2x2: E1:F2
      expect(ws.getCell("E1").result).toBe("Apple");
      expect(ws.getCell("F1").value).toBe(1.5);
      expect(ws.getCell("E2").value).toBe("Cherry");
      expect(ws.getCell("F2").value).toBe(3.0);
    });

    it("should handle scalar result from dynamic array formula", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "OnlyOne";
      ws.getCell("B1").value = 1;

      ws.getCell("D1").value = {
        formula: "_xlfn._xlws.FILTER(A1:A1,B1:B1)",
        result: "",
        shareType: "array",
        ref: "D1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Single result — no spill needed
      expect(ws.getCell("D1").result).toBe("OnlyOne");
    });

    it("should drop stale spill ghosts when source changes to non-dynamic formula", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("X1").value = 10;
      ws.getCell("X2").value = 20;
      ws.getCell("X3").value = 30;
      // First: A1 is a dynamic array that spills to A1:A3
      ws.getCell("A1").value = {
        formula: "_xlfn._xlws.SORT(X1:X3)",
        result: 0,
        shareType: "array",
        ref: "A1",
        isDynamicArray: true
      };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toBe(10);
      // Ghost cells populated
      expect(ws.getCell("A2").value).toBe(20);
      expect(ws.getCell("A3").value).toBe(30);

      // Replace A1 with a non-dynamic formula
      ws.getCell("A1").value = { formula: "SUM(X1:X3)", result: 0 };
      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(60);
      // Ghost cells should be cleared
      expect(ws.findCell(2, 1)?.value ?? null).toBeNull();
      expect(ws.findCell(3, 1)?.value ?? null).toBeNull();
    });
  });

  // ==========================================================================
  // CSE Array Formulas
  // ==========================================================================
  describe("CSE array formulas", () => {
    it("should distribute array formula results across ref range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Data
      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("A3").value = 3;
      ws.getCell("B1").value = 10;
      ws.getCell("B2").value = 20;
      ws.getCell("B3").value = 30;

      // CSE array formula: {=A1:A3*B1:B3} in C1:C3
      // Master cell at C1 with ref="C1:C3"
      ws.getCell("C1").value = {
        formula: "A1:A3*B1:B3",
        result: 0,
        shareType: "array",
        ref: "C1:C3"
      };
      // Slave cells — shared formula referencing master
      ws.getCell("C2").value = {
        formula: "A1:A3*B1:B3",
        result: 0,
        shareType: "array"
      };
      ws.getCell("C3").value = {
        formula: "A1:A3*B1:B3",
        result: 0,
        shareType: "array"
      };

      wb.calculateFormulas();

      // Array multiplication: [1*10, 2*20, 3*30] = [10, 40, 90]
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").result).toBe(40);
      expect(ws.getCell("C3").result).toBe(90);
    });

    it("should fill CSE range with scalar when formula returns scalar", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // CSE formula that returns a scalar: {=SUM(A1:A3)} in B1:B3
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;

      ws.getCell("B1").value = {
        formula: "SUM(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "B1:B3"
      };
      ws.getCell("B2").value = {
        formula: "SUM(A1:A3)",
        result: 0,
        shareType: "array"
      };
      ws.getCell("B3").value = {
        formula: "SUM(A1:A3)",
        result: 0,
        shareType: "array"
      };

      wb.calculateFormulas();

      // All cells in ref range get the same scalar value
      expect(ws.getCell("B1").result).toBe(60);
      expect(ws.getCell("B2").result).toBe(60);
      expect(ws.getCell("B3").result).toBe(60);
    });
  });

  // ==========================================================================
  // Array Arithmetic Broadcasting
  // ==========================================================================
  describe("array arithmetic broadcasting", () => {
    it("should broadcast scalar * array", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("A3").value = 3;

      // CSE: {=A1:A3*10} — scalar broadcast
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

      expect(ws.getCell("B1").result).toBe(10);
      expect(ws.getCell("B2").result).toBe(20);
      expect(ws.getCell("B3").result).toBe(30);
    });

    it("should broadcast row + column to produce matrix", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Row: A1:C1 = {1, 2, 3}
      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = 2;
      ws.getCell("C1").value = 3;

      // Column: A3:A5 = {10; 20; 30}
      ws.getCell("A3").value = 10;
      ws.getCell("A4").value = 20;
      ws.getCell("A5").value = 30;

      // Dynamic array: A1:C1 + A3:A5 should produce 3x3 matrix
      ws.getCell("E1").value = {
        formula: "A1:C1+A3:A5",
        result: 0,
        shareType: "array",
        ref: "E1",
        isDynamicArray: true
      };

      wb.calculateFormulas();

      // Broadcasting: each row of column + each col of row
      // Row 0: 10+1=11, 10+2=12, 10+3=13
      expect(ws.getCell("E1").result).toBe(11);
      expect(ws.getCell("F1").value).toBe(12);
      expect(ws.getCell("G1").value).toBe(13);
      // Row 1: 20+1=21, 20+2=22, 20+3=23
      expect(ws.getCell("E2").value).toBe(21);
      expect(ws.getCell("F2").value).toBe(22);
      expect(ws.getCell("G2").value).toBe(23);
      // Row 2: 30+1=31, 30+2=32, 30+3=33
      expect(ws.getCell("E3").value).toBe(31);
      expect(ws.getCell("F3").value).toBe(32);
      expect(ws.getCell("G3").value).toBe(33);
    });

    it("should handle element-wise operations on same-sized arrays", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("A3").value = 3;
      ws.getCell("B1").value = 10;
      ws.getCell("B2").value = 20;
      ws.getCell("B3").value = 30;

      // CSE: {=A1:A3+B1:B3}
      ws.getCell("C1").value = {
        formula: "A1:A3+B1:B3",
        result: 0,
        shareType: "array",
        ref: "C1:C3"
      };
      ws.getCell("C2").value = {
        formula: "A1:A3+B1:B3",
        result: 0,
        shareType: "array"
      };
      ws.getCell("C3").value = {
        formula: "A1:A3+B1:B3",
        result: 0,
        shareType: "array"
      };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe(11);
      expect(ws.getCell("C2").result).toBe(22);
      expect(ws.getCell("C3").result).toBe(33);
    });

    it("should apply unary minus element-wise on arrays", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("A3").value = 3;

      // CSE: {=-A1:A3}
      ws.getCell("B1").value = {
        formula: "-A1:A3",
        result: 0,
        shareType: "array",
        ref: "B1:B3"
      };
      ws.getCell("B2").value = {
        formula: "-A1:A3",
        result: 0,
        shareType: "array"
      };
      ws.getCell("B3").value = {
        formula: "-A1:A3",
        result: 0,
        shareType: "array"
      };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(-1);
      expect(ws.getCell("B2").result).toBe(-2);
      expect(ws.getCell("B3").result).toBe(-3);
    });
  });

  // ==========================================================================
  // Implicit Intersection
  // ==========================================================================
  describe("implicit intersection", () => {
    it("should use implicit intersection for range in scalar context", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Column data A1:A3
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;

      // Formulas in B1:B3 that reference A1:A3 in a scalar context
      // B1 should pick A1 (same row), B2→A2, B3→A3
      ws.getCell("B1").value = { formula: "A1:A3+100", result: 0 };
      ws.getCell("B2").value = { formula: "A1:A3+100", result: 0 };
      ws.getCell("B3").value = { formula: "A1:A3+100", result: 0 };

      wb.calculateFormulas();

      // Implicit intersection: each formula picks the element from its own row
      expect(ws.getCell("B1").result).toBe(110);
      expect(ws.getCell("B2").result).toBe(120);
      expect(ws.getCell("B3").result).toBe(130);
    });
  });

  // ==========================================================================
  // Spill cleanup on re-evaluation
  // ==========================================================================
  describe("spill idempotency", () => {
    it("should produce correct results on repeated calculateFormulas calls", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;

      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);

      // Recalculate — should produce identical results
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);
    });
  });

  // ==========================================================================
  // Spill data safety: user-written data must not be silently destroyed
  // ==========================================================================
  describe("spill data safety", () => {
    it("should return #SPILL! when user writes into a former ghost cell", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Set up source data and a SORT formula that spills 3 rows
      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      // First calculation — spills to C1:C3
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);

      // User writes a formula into C2 (a former ghost cell)
      ws.getCell("C2").value = { formula: "42+1", result: 0 };

      // Recalculate — should detect conflict and return #SPILL!, not overwrite
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toEqual({ error: "#SPILL!" });
      // User's formula in C2 must be preserved
      expect(ws.getCell("C2").result).toBe(43);
    });

    it("should clean up old ghost cells when spill region shrinks", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // FILTER that initially returns 3 results
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("B1").value = 1;
      ws.getCell("B2").value = 1;
      ws.getCell("B3").value = 1;
      ws.getCell("D1").value = {
        formula: "_xlfn._xlws.FILTER(A1:A3,B1:B3)",
        result: 0,
        shareType: "array",
        ref: "D1",
        isDynamicArray: true
      };

      wb.calculateFormulas();
      expect(ws.getCell("D1").result).toBe(10);
      expect(ws.getCell("D2").value).toBe(20);
      expect(ws.getCell("D3").value).toBe(30);

      // Now change filter so only 1 result matches
      ws.getCell("B2").value = 0;
      ws.getCell("B3").value = 0;

      wb.calculateFormulas();
      expect(ws.getCell("D1").result).toBe(10);
      // Old ghost cells must be cleared
      const d2 = ws.findCell(2, 4); // D2
      const d3 = ws.findCell(3, 4); // D3
      const d2Val = d2 ? d2.value : null;
      const d3Val = d3 ? d3.value : null;
      expect(d2Val).toBeNull();
      expect(d3Val).toBeNull();
    });

    it("should return #SPILL! when user writes a plain value into a former ghost cell", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Set up source data and a SORT formula that spills 3 rows
      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      // First calculation — spills to C1:C3
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);

      // User writes a plain number into C2 (a former ghost cell)
      ws.getCell("C2").value = 42;

      // Recalculate — should detect conflict and return #SPILL!, not overwrite
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toEqual({ error: "#SPILL!" });
      // User's value in C2 must be preserved
      expect(ws.getCell("C2").value).toBe(42);
    });

    it("should return #SPILL! when user writes a string into a former ghost cell", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);

      // User writes a string into C3
      ws.getCell("C3").value = "user data";

      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toEqual({ error: "#SPILL!" });
      expect(ws.getCell("C3").value).toBe("user data");
    });

    it("should not clear user-modified ghost cells when source formula is deleted", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      // First calculation — spills to C1:C3
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);

      // User writes a plain value into C2
      ws.getCell("C2").value = 99;

      // Delete the source formula (replace with a plain value)
      ws.getCell("C1").value = "no formula";

      // Recalculate — stale cleanup should NOT clear user's C2 value
      wb.calculateFormulas();
      expect(ws.getCell("C2").value).toBe(99);
      // C3 was not modified by the user, so it should be cleaned up
      const c3 = ws.findCell(3, 3); // C3
      const c3Val = c3 ? c3.value : null;
      expect(c3Val).toBeNull();
    });
  });

  // ==========================================================================
  // Spill stability across sheet rename / delete
  // ==========================================================================
  describe("spill stability across sheet rename and delete", () => {
    it("should preserve spill ghosts after renaming the sheet", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);

      // Rename the sheet
      ws.name = "Renamed";

      // Recalculate — spill should still work correctly
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);
    });

    it("should detect #SPILL! after rename when user modifies a ghost cell", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);

      // Rename the sheet, then user writes into a ghost cell
      ws.name = "Renamed";
      ws.getCell("C2").value = "user data";

      // Recalculate — should detect conflict and return #SPILL!
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toEqual({ error: "#SPILL!" });
      expect(ws.getCell("C2").value).toBe("user data");
    });

    it("should clean up stale ghosts after rename when source formula is deleted", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 30;
      ws.getCell("A2").value = 10;
      ws.getCell("A3").value = 20;
      ws.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
      expect(ws.getCell("C2").value).toBe(20);
      expect(ws.getCell("C3").value).toBe(30);

      // User modifies C2 before rename
      ws.getCell("C2").value = 99;

      // Rename the sheet, then delete the source formula
      ws.name = "Renamed";
      ws.getCell("C1").value = "no formula";

      // Recalculate — stale cleanup should clear unmodified C3, preserve user C2
      wb.calculateFormulas();
      expect(ws.getCell("C2").value).toBe(99);
      const c3 = ws.findCell(3, 3);
      const c3Val = c3 ? c3.value : null;
      expect(c3Val).toBeNull();
    });

    it("should not error when a sheet with spill data is deleted", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");

      ws1.getCell("A1").value = 30;
      ws1.getCell("A2").value = 10;
      ws1.getCell("A3").value = 20;
      ws1.getCell("C1").value = {
        formula: "_xlfn._xlws.SORT(A1:A3)",
        result: 0,
        shareType: "array",
        ref: "C1",
        isDynamicArray: true
      };

      // Put a formula on Sheet2 so calculateFormulas has something to process
      ws2.getCell("A1").value = { formula: "1+1", result: 0 };

      wb.calculateFormulas();
      expect(ws1.getCell("C1").result).toBe(10);
      expect(ws1.getCell("C2").value).toBe(20);

      // Delete Sheet1 — spill metadata should be silently discarded
      wb.removeWorksheet(ws1.id);

      // Recalculate — should not throw
      expect(() => wb.calculateFormulas()).not.toThrow();
      expect(ws2.getCell("A1").result).toBe(2);
    });
  });
  describe("whole-column and whole-row implicit intersection", () => {
    it("should use implicit intersection for whole-column reference (A:A*2)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;

      // B2 = A:A * 2 — implicit intersection should pick A2 (row 2)
      ws.getCell("B1").value = { formula: "A:A*2", result: 0 };
      ws.getCell("B2").value = { formula: "A:A*2", result: 0 };
      ws.getCell("B3").value = { formula: "A:A*2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(20);
      expect(ws.getCell("B2").result).toBe(40);
      expect(ws.getCell("B3").result).toBe(60);
    });

    it("should use implicit intersection for whole-row reference (1:1+1)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 100;
      ws.getCell("B1").value = 200;
      ws.getCell("C1").value = 300;

      // A2, B2, C2 = 1:1 + 1 — implicit intersection picks column
      ws.getCell("A2").value = { formula: "1:1+1", result: 0 };
      ws.getCell("B2").value = { formula: "1:1+1", result: 0 };
      ws.getCell("C2").value = { formula: "1:1+1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(101);
      expect(ws.getCell("B2").result).toBe(201);
      expect(ws.getCell("C2").result).toBe(301);
    });
  });

  // ==========================================================================
  // Self-referencing + iterative calculation
  // ==========================================================================
  describe("self-reference with iterative calculation", () => {
    it("should handle A1=A1+1 with iterate enabled", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Enable iterative calculation
      wb.calcProperties = {
        fullCalcOnLoad: false,
        iterate: true,
        iterateCount: 10,
        iterateDelta: 0
      };

      ws.getCell("A1").value = { formula: "A1+1", result: 0 };

      wb.calculateFormulas();

      // Initial pass: circular ref returns 0 → 0+1=1
      // Then 10 iterations: 1→2→3→...→11
      expect(ws.getCell("A1").result).toBe(11);
    });

    it("should converge with iterateDelta", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      wb.calcProperties = {
        fullCalcOnLoad: false,
        iterate: true,
        iterateCount: 100,
        iterateDelta: 0.001
      };

      // A1 = A1 / 2 + 1 converges to 2
      ws.getCell("A1").value = { formula: "A1/2+1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(2, 2);
    });
  });

  // ==========================================================================
  // Multi-area defined names
  // ==========================================================================
  describe("multi-area defined names", () => {
    it("should not silently truncate multi-area defined names", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("C1").value = 30;
      ws.getCell("C2").value = 40;

      // Add a defined name with two non-adjacent ranges (separated by column B)
      wb.definedNames.add("Sheet1!$A$1:$A$2", "MyRange");
      wb.definedNames.add("Sheet1!$C$1:$C$2", "MyRange");

      // Formula using the multi-area name
      ws.getCell("D1").value = { formula: "SUM(MyRange)", result: 0 };

      wb.calculateFormulas();

      // Multi-area names are not supported — should NOT silently return
      // SUM of just the first range (which would be 30).
      // Our implementation returns #VALUE! for multi-area names.
      const result = ws.getCell("D1").result;
      expect(result).toEqual({ error: "#VALUE!" });
    });

    it("should work correctly with single-area defined names", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;

      wb.definedNames.add("Sheet1!$A$1:$A$3", "MyRange");

      ws.getCell("B1").value = { formula: "SUM(MyRange)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(60);
    });
  });

  // ==========================================================================
  // XLOOKUP reverse search and approximate match
  // ==========================================================================
  describe("XLOOKUP advanced modes", () => {
    it("should support reverse search (searchMode = -1)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Duplicate values — reverse search should find the last occurrence
      ws.getCell("A1").value = "Apple";
      ws.getCell("A2").value = "Banana";
      ws.getCell("A3").value = "Apple";
      ws.getCell("B1").value = 1;
      ws.getCell("B2").value = 2;
      ws.getCell("B3").value = 3;

      ws.getCell("C1").value = {
        formula: 'XLOOKUP("Apple",A1:A3,B1:B3,,0,-1)',
        result: 0
      };

      wb.calculateFormulas();

      // Reverse search: should find the last "Apple" at row 3, return 3
      expect(ws.getCell("C1").result).toBe(3);
    });

    it("should support approximate match - next smaller (matchMode = -1)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("B1").value = "low";
      ws.getCell("B2").value = "mid";
      ws.getCell("B3").value = "high";

      // Look for 25 — next smaller is 20 → "mid"
      ws.getCell("C1").value = {
        formula: "XLOOKUP(25,A1:A3,B1:B3,,-1)",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe("mid");
    });

    it("should support approximate match - next larger (matchMode = 1)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("B1").value = "low";
      ws.getCell("B2").value = "mid";
      ws.getCell("B3").value = "high";

      // Look for 25 — next larger is 30 → "high"
      ws.getCell("C1").value = {
        formula: "XLOOKUP(25,A1:A3,B1:B3,,1)",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toBe("high");
    });
  });

  // ==========================================================================
  // Named range implicit intersection
  // ==========================================================================
  describe("named range implicit intersection", () => {
    it("should use implicit intersection when named range resolves to a column range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;

      // Define MyCol as A1:A3
      wb.definedNames.add("Sheet1!$A$1:$A$3", "MyCol");

      // B1:B3 = MyCol * 2 — should pick the value from the formula's own row
      ws.getCell("B1").value = { formula: "MyCol*2", result: 0 };
      ws.getCell("B2").value = { formula: "MyCol*2", result: 0 };
      ws.getCell("B3").value = { formula: "MyCol*2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(20); // A1*2
      expect(ws.getCell("B2").result).toBe(40); // A2*2
      expect(ws.getCell("B3").result).toBe(60); // A3*2
    });

    it("should use implicit intersection when named range resolves to a row range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("B1").value = 20;
      ws.getCell("C1").value = 30;

      // Define MyRow as A1:C1
      wb.definedNames.add("Sheet1!$A$1:$C$1", "MyRow");

      // A2:C2 = MyRow + 1 — should pick the value from the formula's own column
      ws.getCell("A2").value = { formula: "MyRow+1", result: 0 };
      ws.getCell("B2").value = { formula: "MyRow+1", result: 0 };
      ws.getCell("C2").value = { formula: "MyRow+1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(11); // A1+1
      expect(ws.getCell("B2").result).toBe(21); // B1+1
      expect(ws.getCell("C2").result).toBe(31); // C1+1
    });
  });

  // ==========================================================================
  // Named range alias with iterative calculation
  // ==========================================================================
  describe("named range alias with iterative calculation", () => {
    it("should use circularFallback for self-reference through named range", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      wb.calcProperties = {
        fullCalcOnLoad: false,
        iterate: true,
        iterateCount: 10,
        iterateDelta: 0
      };

      // Foo -> A1, A1 = Foo + 1 (self-reference through named range)
      wb.definedNames.add("Sheet1!$A$1", "Foo");
      ws.getCell("A1").value = { formula: "Foo+1", result: 0 };

      wb.calculateFormulas();

      // Same behavior as A1=A1+1: initial pass 0→1, then 10 iterations → 11
      expect(ws.getCell("A1").result).toBe(11);
    });

    it("should converge through named range alias", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      wb.calcProperties = {
        fullCalcOnLoad: false,
        iterate: true,
        iterateCount: 100,
        iterateDelta: 0.001
      };

      // Bar -> A1, A1 = Bar/2 + 1 converges to 2
      wb.definedNames.add("Sheet1!$A$1", "Bar");
      ws.getCell("A1").value = { formula: "Bar/2+1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBeCloseTo(2, 2);
    });
  });

  // ==========================================================================
  // Formula-based defined names
  // ==========================================================================
  describe("formula-based defined names", () => {
    it("should evaluate a formula-based defined name", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Register a formula-based name that evaluates to 42
      wb.definedNames.addFormula("TheAnswer", "40+2");

      ws.getCell("A1").value = { formula: "TheAnswer", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(42);
    });

    it("should not degrade array result to scalar on second use in same formula round", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Register a formula-based name that returns an array (SUM over range)
      // We use a cell-reference name + a formula that uses it twice to test caching
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;

      wb.definedNames.add("Sheet1!$A$1:$A$3", "MyRange");

      // Both B1 and B2 use MyRange — the second evaluation must not get a
      // degraded scalar from the cache
      ws.getCell("B1").value = { formula: "SUM(MyRange)", result: 0 };
      ws.getCell("B2").value = { formula: "SUM(MyRange)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(60);
      expect(ws.getCell("B2").result).toBe(60);
    });

    it("addFormula overrides a previous add, and vice versa", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;

      // First bind as cell reference
      wb.definedNames.add("Sheet1!$A$1", "MyName");
      // Override with formula
      wb.definedNames.addFormula("MyName", "99");

      ws.getCell("B1").value = { formula: "MyName", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(99);

      // Now override back to cell reference
      wb.definedNames.add("Sheet1!$A$1", "MyName");

      ws.getCell("C1").value = { formula: "MyName", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("C1").result).toBe(10);
    });

    it("should handle Unicode characters in defined name identifiers", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 100;
      wb.definedNames.add("Sheet1!A1", "销售额");
      ws.getCell("B1").value = { formula: "销售额*2", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(200);
    });
  });

  // ==========================================================================
  // Scoped Defined Names
  // ==========================================================================
  describe("scoped defined names", () => {
    it("should resolve sheet-scoped name over workbook-scoped name", async () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");

      ws1.getCell("A1").value = 100;
      ws2.getCell("A1").value = 200;

      // Global "MyVal" → Sheet1!A1, Sheet1-local "MyVal" → Sheet2!A1
      wb.definedNames.model = [
        { name: "MyVal", ranges: ["Sheet1!$A$1"], rawText: "Sheet1!$A$1" },
        {
          name: "MyVal",
          ranges: ["Sheet2!$A$1"],
          rawText: "Sheet2!$A$1",
          localSheetId: 0
        }
      ];

      // Verify the DefinedNames layer works correctly
      const allNames = wb.definedNames.getAllNames();
      expect(allNames.length).toBe(2);

      // Check snapshot construction
      const { buildWorkbookSnapshot } = await import("@formula/integration/workbook-adapter");
      const snapshot = buildWorkbookSnapshot(wb as any);

      // Verify we have both entries in the snapshot
      expect(snapshot.definedNames.size).toBe(2);

      // Verify scope resolution
      const { resolveDefinedName } = await import("@formula/integration/workbook-snapshot");
      const fromSheet1 = resolveDefinedName(snapshot.definedNames, "MyVal", "Sheet1");
      const fromSheet2 = resolveDefinedName(snapshot.definedNames, "MyVal", "Sheet2");
      expect(fromSheet1).toBeDefined();
      expect(fromSheet1!.ranges[0]).toContain("Sheet2"); // sheet-local on Sheet1 → Sheet2
      expect(fromSheet2).toBeDefined();
      expect(fromSheet2!.ranges[0]).toContain("Sheet1"); // global fallback → Sheet1

      // Formula on Sheet1 should see the sheet-local "MyVal" → 200
      ws1.getCell("B1").value = { formula: "MyVal", result: 0 };
      // Formula on Sheet2 should see the global "MyVal" → 100
      ws2.getCell("B1").value = { formula: "MyVal", result: 0 };

      wb.calculateFormulas();

      expect(ws1.getCell("B1").result).toBe(200);
      expect(ws2.getCell("B1").result).toBe(100);
    });

    it("should fall back to global name when no sheet-local exists", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");

      ws1.getCell("A1").value = 42;

      // Only a global "GlobalOnly" name, no sheet-local override
      wb.definedNames.model = [
        { name: "GlobalOnly", ranges: ["Sheet1!$A$1"], rawText: "Sheet1!$A$1" }
      ];

      ws1.getCell("B1").value = { formula: "GlobalOnly", result: 0 };
      ws2.getCell("B1").value = { formula: "GlobalOnly", result: 0 };

      wb.calculateFormulas();

      // Both sheets should resolve to the global name
      expect(ws1.getCell("B1").result).toBe(42);
      expect(ws2.getCell("B1").result).toBe(42);
    });

    it("should not cross-contaminate scoped and global name content in snapshot", async () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");

      ws1.getCell("A1").value = 10;
      ws2.getCell("A1").value = 20;

      // Global "X" = Sheet1!A1, Sheet2-local "X" = Sheet2!A1
      wb.definedNames.model = [
        { name: "X", ranges: ["Sheet1!$A$1"], rawText: "Sheet1!$A$1" },
        {
          name: "X",
          ranges: ["Sheet2!$A$1"],
          rawText: "Sheet2!$A$1",
          localSheetId: 1
        }
      ];

      const { buildWorkbookSnapshot } = await import("@formula/integration/workbook-adapter");
      const snapshot = buildWorkbookSnapshot(wb as any);

      // Verify distinct entries exist and don't cross-contaminate
      const globalEntry = snapshot.definedNames.get("X");
      expect(globalEntry).toBeDefined();
      expect(globalEntry!.ranges[0]).toContain("Sheet1");
      expect(globalEntry!.scope).toBeUndefined();

      const { scopedNameKey } = await import("@formula/integration/workbook-snapshot");
      const scopedEntry = snapshot.definedNames.get(scopedNameKey("Sheet2", "X"));
      expect(scopedEntry).toBeDefined();
      expect(scopedEntry!.ranges[0]).toContain("Sheet2");
      expect(scopedEntry!.scope).toBe("Sheet2");
    });

    it("should propagate deps for formula-based scoped defined name", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");

      // A1=5, A2=10, formula name "MySum" = Sheet1!A1+Sheet1!A2
      ws1.getCell("A1").value = 5;
      ws1.getCell("A2").value = 10;
      wb.definedNames.addFormula("MySum", "Sheet1!$A$1+Sheet1!$A$2");

      ws1.getCell("B1").value = { formula: "MySum", result: 0 };

      wb.calculateFormulas();

      expect(ws1.getCell("B1").result).toBe(15);
    });
  });

  // ==========================================================================
  // Critical Edge Cases
  // ==========================================================================
  describe("critical edge cases", () => {
    it("should resolve INDIRECT to correct sheet", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");
      ws1.getCell("A1").value = 10;
      ws2.getCell("A1").value = 20;
      // INDIRECT("A1") on Sheet1 should get Sheet1!A1
      ws1.getCell("B1").value = { formula: 'INDIRECT("A1")', result: 0 };
      // INDIRECT("A1") on Sheet2 should get Sheet2!A1
      ws2.getCell("B1").value = { formula: 'INDIRECT("A1")', result: 0 };
      wb.calculateFormulas();
      expect(ws1.getCell("B1").result).toBe(10);
      expect(ws2.getCell("B1").result).toBe(20);
    });

    it("should handle OFFSET with negative height", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("A3").value = 3;
      // OFFSET(A3,0,0,-3,1) should reference A1:A3
      ws.getCell("B1").value = { formula: "SUM(OFFSET(A3,0,0,-3,1))", result: 0 };
      // OFFSET(A1,0,0,-1,1) → just A1
      ws.getCell("B2").value = { formula: "SUM(OFFSET(A1,0,0,-1,1))", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(6);
      expect(ws.getCell("B2").result).toBe(1);
    });

    it("should return #REF! for OFFSET with zero height", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = { formula: "OFFSET(A1,0,0,0,1)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toEqual({ error: "#REF!" });
    });

    it("should re-evaluate non-circular dependents after iterative convergence", () => {
      const wb = new Workbook();
      wb.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 };
      const ws = wb.addWorksheet("Sheet1");
      // Circular: A1 = IF(A1>0, A1, 1) — self-ref that seeds to 1
      ws.getCell("A1").value = { formula: "IF(A1>0,A1,1)", result: 0 };
      // Non-circular dependent
      ws.getCell("B1").value = { formula: "A1*10", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toBe(1);
      expect(ws.getCell("B1").result).toBe(10);
    });

    it("should propagate errors through text functions", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = { formula: "1/0", result: 0 };
      ws.getCell("B1").value = { formula: "LEFT(A1, 2)", result: 0 };
      ws.getCell("C1").value = { formula: 'CONCATENATE("x", A1)', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toEqual({ error: "#DIV/0!" });
      expect(ws.getCell("C1").result).toEqual({ error: "#DIV/0!" });
    });

    it("should merge INDIRECT dynamic deps and re-evaluate dependents", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // A1 holds a cell address, B1 holds the value at that address
      ws.getCell("A1").value = "C1";
      ws.getCell("C1").value = 42;
      // INDIRECT(A1) resolves "C1" → C1 at runtime (dynamic dep)
      ws.getCell("B1").value = { formula: "INDIRECT(A1)", result: 0 };
      // D1 depends on B1 (downstream of dynamic dep)
      ws.getCell("D1").value = { formula: "B1*2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(42);
      expect(ws.getCell("D1").result).toBe(84);
    });

    it("should evaluate position-dependent formula names per cell", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Formula name "R" = ROW() — depends on calling cell's position
      wb.definedNames.addFormula("R", "ROW()");

      ws.getCell("A1").value = { formula: "R", result: 0 };
      ws.getCell("A2").value = { formula: "R", result: 0 };
      ws.getCell("A3").value = { formula: "R", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(1);
      expect(ws.getCell("A2").result).toBe(2);
      expect(ws.getCell("A3").result).toBe(3);
    });

    it("should re-evaluate defined name dependents after iterative convergence", () => {
      const wb = new Workbook();
      wb.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 };
      const ws = wb.addWorksheet("Sheet1");

      // Circular: A1 = IF(A1>0,A1,1)
      ws.getCell("A1").value = { formula: "IF(A1>0,A1,1)", result: 0 };
      // Name "N" references A1
      wb.definedNames.addFormula("N", "Sheet1!$A$1");
      // B1 depends on N (indirect dependency on circular A1)
      ws.getCell("B1").value = { formula: "N*10", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(1);
      expect(ws.getCell("B1").result).toBe(10);
    });

    it("should handle OFFSET with negative width", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = 2;
      ws.getCell("C1").value = 3;

      // OFFSET(C1,0,0,1,-3) → A1:C1
      ws.getCell("A2").value = { formula: "SUM(OFFSET(C1,0,0,1,-3))", result: 0 };
      // OFFSET(A1,0,0,1,0) → #REF!
      ws.getCell("B2").value = { formula: "OFFSET(A1,0,0,1,0)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A2").result).toBe(6);
      expect(ws.getCell("B2").result).toEqual({ error: "#REF!" });
    });

    it("should handle 3D cell reference across sheets", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");
      const ws3 = wb.addWorksheet("Sheet3");
      const wsResult = wb.addWorksheet("Result");

      ws1.getCell("A1").value = 1;
      ws2.getCell("A1").value = 2;
      ws3.getCell("A1").value = 3;

      // 3D cell reference: SUM across Sheet1:Sheet3!A1
      wsResult.getCell("A1").value = { formula: "SUM(Sheet1:Sheet3!A1)", result: 0 };

      wb.calculateFormulas();

      expect(wsResult.getCell("A1").result).toBe(6);
    });

    it("should handle 3D whole-row reference across sheets", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");
      const wsResult = wb.addWorksheet("Result");

      ws1.getCell("A1").value = 10;
      ws1.getCell("B1").value = 20;
      ws2.getCell("A1").value = 30;
      ws2.getCell("B1").value = 40;

      // 3D row range: SUM(Sheet1:Sheet2!1:1)
      wsResult.getCell("A1").value = { formula: "SUM(Sheet1:Sheet2!1:1)", result: 0 };

      wb.calculateFormulas();

      // 10 + 20 + 30 + 40 = 100
      expect(wsResult.getCell("A1").result).toBe(100);
    });

    it("should propagate errors through FILTER include array", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("B1").value = true;
      ws.getCell("B2").value = { formula: "1/0", result: 0 };

      ws.getCell("C1").value = {
        formula: "FILTER(A1:A2,B1:B2)",
        result: 0,
        isDynamicArray: true
      };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toEqual({ error: "#DIV/0!" });
    });

    it("should return correct weekday for return types 11-17", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // 2024-01-01 is a Monday. Excel serial = 45292
      ws.getCell("A1").value = 45292;
      // Type 11: Mon=1
      ws.getCell("B1").value = { formula: "WEEKDAY(A1,11)", result: 0 };
      // Type 17: Sun=1, Mon=2
      ws.getCell("C1").value = { formula: "WEEKDAY(A1,17)", result: 0 };
      // Invalid type → #NUM!
      ws.getCell("D1").value = { formula: "WEEKDAY(A1,4)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(1); // Monday → 1 for type 11
      expect(ws.getCell("C1").result).toBe(2); // Monday → 2 for type 17
      expect(ws.getCell("D1").result).toEqual({ error: "#NUM!" });
    });

    it("should handle structured ref #Headers+#Data combination", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.addTable({
        name: "Sales",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "Item" }, { name: "Price" }],
        rows: [
          ["Apple", 1],
          ["Banana", 2],
          ["Cherry", 3]
        ]
      });

      // COUNTA of #Headers + #Data for the Price column = header + 3 data rows = 4
      ws.getCell("D1").value = {
        formula: "COUNTA(Sales[[#Headers],[#Data],[Price]])",
        result: 0
      };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toBe(4);
    });

    it("should handle structured ref #Data+#Totals combination", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.addTable({
        name: "Inventory",
        ref: "A1",
        headerRow: true,
        totalsRow: true,
        columns: [{ name: "Item" }, { name: "Qty", totalsRowFunction: "sum" }],
        rows: [
          ["Apple", 10],
          ["Banana", 20]
        ]
      });

      // #Data + #Totals for Qty = data rows + totals row = 3 cells
      ws.getCell("D1").value = {
        formula: "COUNTA(Inventory[[#Data],[#Totals],[Qty]])",
        result: 0
      };

      wb.calculateFormulas();

      // 2 data cells + 1 totals cell = 3
      expect(ws.getCell("D1").result).toBe(3);
    });

    it("should re-evaluate INDIRECT when target address changes", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = "C1";
      ws.getCell("C1").value = 100;
      ws.getCell("E1").value = 200;
      ws.getCell("B1").value = { formula: "INDIRECT(A1)", result: 0 };

      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(100);

      // Change the target address
      ws.getCell("A1").value = "E1";
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(200);
    });

    it("should re-evaluate multi-hop downstream after iterative convergence", () => {
      const wb = new Workbook();
      wb.calcProperties = { iterate: true, iterateCount: 100, iterateDelta: 0.001 };
      const ws = wb.addWorksheet("Sheet1");

      // A1 = circular self-ref that converges to 1
      ws.getCell("A1").value = { formula: "IF(A1>0,A1,1)", result: 0 };
      // B1 = first hop
      ws.getCell("B1").value = { formula: "A1*10", result: 0 };
      // C1 = second hop
      ws.getCell("C1").value = { formula: "B1+5", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toBe(1);
      expect(ws.getCell("B1").result).toBe(10);
      expect(ws.getCell("C1").result).toBe(15);
    });

    it("should handle structured ref [#This Row]", () => {
      // Regression: `#This Row` is supported end-to-end. The original
      // variant of this test used `"T1"` as the table name, which the
      // library's Table constructor sanitises to `"_T1"` because
      // bare-letter+digit names collide with cell refs — the formula
      // then referenced a non-existent table. Using an unambiguous
      // table name (`MyTable`) exercises the intended path.
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.addTable({
        name: "MyTable",
        ref: "A1",
        headerRow: true,
        totalsRow: false,
        columns: [{ name: "Val" }, { name: "Double" }],
        rows: [
          [10, null],
          [20, null],
          [30, null]
        ]
      });

      // Formulas in the Double column reference [#This Row]
      ws.getCell("B2").value = { formula: "MyTable[[#This Row],[Val]]*2", result: 0 };
      ws.getCell("B3").value = { formula: "MyTable[[#This Row],[Val]]*2", result: 0 };
      ws.getCell("B4").value = { formula: "MyTable[[#This Row],[Val]]*2", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B2").result).toBe(20);
      expect(ws.getCell("B3").result).toBe(40);
      expect(ws.getCell("B4").result).toBe(60);
    });

    it("should propagate error from SUMIF criteria", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("B1").value = 10;
      ws.getCell("B2").value = 20;

      // Criteria is an error
      ws.getCell("C1").value = { formula: "1/0", result: 0 };
      ws.getCell("D1").value = { formula: "SUMIF(A1:A2,C1,B1:B2)", result: 0 };
      ws.getCell("D2").value = { formula: "COUNTIF(A1:A2,C1)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("D1").result).toEqual({ error: "#DIV/0!" });
      expect(ws.getCell("D2").result).toEqual({ error: "#DIV/0!" });
    });
  });

  // ==========================================================================
  // Newly added functions — ISFORMULA, FORMULATEXT, HYPERLINK,
  // F.DIST.RT, F.INV.RT, SKEW, SKEW.P, KURT, PRICE, YIELD, DURATION,
  // MDURATION, ACCRINT.
  // ==========================================================================
  describe("newly added functions dispatch", () => {
    it("should dispatch ISFORMULA, FORMULATEXT, HYPERLINK, stats, and bond functions", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // A1 contains a value; A2 contains a formula referring to A1.
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = { formula: "A1*2", result: 0 };

      // ISFORMULA / FORMULATEXT.
      ws.getCell("B1").value = { formula: "ISFORMULA(A1)", result: 0 };
      ws.getCell("B2").value = { formula: "ISFORMULA(A2)", result: 0 };
      ws.getCell("B3").value = { formula: "FORMULATEXT(A2)", result: 0 };
      ws.getCell("B4").value = { formula: "FORMULATEXT(A1)", result: 0 };

      // HYPERLINK(url, friendly) and HYPERLINK(url) alone.
      ws.getCell("C1").value = {
        formula: 'HYPERLINK("http://x.com","click me")',
        result: 0
      };
      ws.getCell("C2").value = { formula: 'HYPERLINK("http://x.com")', result: 0 };

      // Statistical data for SKEW / SKEW.P / KURT — known values:
      // sample 1..10: SKEW ≈ 0, SKEW.P = 0, KURT ≈ -1.2
      ws.getCell("D1").value = 1;
      ws.getCell("D2").value = 2;
      ws.getCell("D3").value = 3;
      ws.getCell("D4").value = 4;
      ws.getCell("D5").value = 5;
      ws.getCell("D6").value = 6;
      ws.getCell("D7").value = 7;
      ws.getCell("D8").value = 8;
      ws.getCell("D9").value = 9;
      ws.getCell("D10").value = 10;
      ws.getCell("E1").value = { formula: "SKEW(D1:D10)", result: 0 };
      ws.getCell("E2").value = { formula: "SKEW.P(D1:D10)", result: 0 };
      ws.getCell("E3").value = { formula: "KURT(D1:D10)", result: 0 };

      // F.DIST.RT and F.INV.RT — self-consistency check:
      // F.INV.RT(F.DIST.RT(2, 5, 10), 5, 10) ≈ 2.
      ws.getCell("F1").value = { formula: "F.DIST.RT(2,5,10)", result: 0 };
      ws.getCell("F2").value = { formula: "F.INV.RT(F.DIST.RT(2,5,10),5,10)", result: 0 };

      // Bond functions using simple parameters (basis 0, semi-annual).
      // Settlement 2020-01-01 = 43831, maturity 2025-01-01 = 45658.
      ws.getCell("G1").value = 43831; // settlement
      ws.getCell("G2").value = 45658; // maturity
      ws.getCell("G3").value = {
        formula: "PRICE(G1,G2,0.05,0.05,100,2,0)",
        result: 0
      };
      // For equal rate and yield, price should be ~100.
      ws.getCell("G4").value = {
        formula: "YIELD(G1,G2,0.05,100,100,2,0)",
        result: 0
      };
      ws.getCell("G5").value = {
        formula: "DURATION(G1,G2,0.05,0.05,2,0)",
        result: 0
      };
      ws.getCell("G6").value = {
        formula: "MDURATION(G1,G2,0.05,0.05,2,0)",
        result: 0
      };
      // ACCRINT: issue 2020-01-01, settlement 2020-07-01 (= 44013),
      // rate 0.05, par 1000, basis 0 → 1000 * 0.05 * (180/360) = 25.
      ws.getCell("H1").value = 44013;
      ws.getCell("H2").value = {
        formula: "ACCRINT(G1,G1,H1,0.05,1000,2,0)",
        result: 0
      };

      wb.calculateFormulas();

      // ISFORMULA / FORMULATEXT.
      expect(ws.getCell("B1").result).toBe(false);
      expect(ws.getCell("B2").result).toBe(true);
      expect(ws.getCell("B3").result).toBe("=A1*2");
      expect(ws.getCell("B4").result).toEqual({ error: "#N/A" });

      // HYPERLINK — friendly name wins; otherwise URL.
      expect(ws.getCell("C1").result).toBe("click me");
      expect(ws.getCell("C2").result).toBe("http://x.com");

      // SKEW/SKEW.P should be ~0 for the symmetric sample 1..10.
      expect(Math.abs(ws.getCell("E1").result as number)).toBeLessThan(1e-9);
      expect(Math.abs(ws.getCell("E2").result as number)).toBeLessThan(1e-9);
      // KURT for 1..10 is approximately -1.2.
      expect(ws.getCell("E3").result as number).toBeCloseTo(-1.2, 5);

      // F.DIST.RT in (0,1); round-trip through F.INV.RT returns ~2.
      const frt = ws.getCell("F1").result as number;
      expect(frt).toBeGreaterThan(0);
      expect(frt).toBeLessThan(1);
      expect(ws.getCell("F2").result as number).toBeCloseTo(2, 4);

      // PRICE at par when yield == coupon.
      expect(ws.getCell("G3").result as number).toBeCloseTo(100, 4);
      // YIELD given price 100 should return 0.05.
      expect(ws.getCell("G4").result as number).toBeCloseTo(0.05, 4);
      // 5-year par bond, 5% coupon semi-annual → duration ≈ 4.49 years.
      expect(ws.getCell("G5").result as number).toBeGreaterThan(4);
      expect(ws.getCell("G5").result as number).toBeLessThan(5);
      // Modified duration slightly less than Macaulay duration.
      expect(ws.getCell("G6").result as number).toBeLessThan(ws.getCell("G5").result as number);
      // ACCRINT on a 30/360 half-year → $25.
      expect(ws.getCell("H2").result as number).toBeCloseTo(25, 6);
    });

    it("should handle F.DIST.RT at x=0 boundary", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = { formula: "F.DIST.RT(0,5,10)", result: 0 };
      ws.getCell("A2").value = { formula: "F.DIST.RT(-1,5,10)", result: 0 };
      wb.calculateFormulas();
      // x=0 → entire distribution above 0 → probability = 1
      expect(ws.getCell("A1").result).toBe(1);
      // x<0 → #NUM! (F-distribution only defined for x ≥ 0)
      expect(ws.getCell("A2").result).toEqual({ error: "#NUM!" });
    });
  });

  // ==========================================================================
  // CELL — information about a referenced cell.
  // Covers the implemented info-type subset: address/row/col/contents/type/
  // width/filename, plus the #N/A fallback for unsupported info types.
  // ==========================================================================
  describe("CELL function", () => {
    it("should return $A$1-style absolute address strings", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("B1").value = { formula: 'CELL("address",A1)', result: 0 };
      ws.getCell("B2").value = { formula: 'CELL("address",Z99)', result: 0 };
      ws.getCell("B3").value = { formula: 'CELL("address",AA1)', result: 0 };
      ws.getCell("B4").value = { formula: 'CELL("address",XFD1048576)', result: 0 };
      // Area ref → top-left.
      ws.getCell("B5").value = { formula: 'CELL("address",C3:D4)', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe("$A$1");
      expect(ws.getCell("B2").result).toBe("$Z$99");
      expect(ws.getCell("B3").result).toBe("$AA$1");
      expect(ws.getCell("B4").result).toBe("$XFD$1048576");
      expect(ws.getCell("B5").result).toBe("$C$3");
    });

    it("should return row and column numbers", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("B1").value = { formula: 'CELL("row",A1)', result: 0 };
      ws.getCell("B2").value = { formula: 'CELL("row",B5)', result: 0 };
      ws.getCell("B3").value = { formula: 'CELL("col",A1)', result: 0 };
      ws.getCell("B4").value = { formula: 'CELL("col",Z7)', result: 0 };
      // Area ref top-left wins.
      ws.getCell("B5").value = { formula: 'CELL("row",D3:E4)', result: 0 };
      ws.getCell("B6").value = { formula: 'CELL("col",D3:E4)', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(1);
      expect(ws.getCell("B2").result).toBe(5);
      expect(ws.getCell("B3").result).toBe(1);
      expect(ws.getCell("B4").result).toBe(26);
      expect(ws.getCell("B5").result).toBe(3);
      expect(ws.getCell("B6").result).toBe(4);
    });

    it("should return cell contents for various value kinds", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 42;
      ws.getCell("A2").value = "hello";
      ws.getCell("A3").value = true;
      // A4 is intentionally empty.
      ws.getCell("B1").value = { formula: 'CELL("contents",A1)', result: 0 };
      ws.getCell("B2").value = { formula: 'CELL("contents",A2)', result: 0 };
      ws.getCell("B3").value = { formula: 'CELL("contents",A3)', result: 0 };
      ws.getCell("B4").value = { formula: 'CELL("contents",A4)', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(42);
      expect(ws.getCell("B2").result).toBe("hello");
      expect(ws.getCell("B3").result).toBe(true);
      // Excel represents empty contents as 0 in numeric contexts, but
      // our writeback path preserves BLANK as `undefined` so that
      // downstream consumers can distinguish "empty" from "literal
      // zero" (see R5-P1-2).
      const b4 = ws.getCell("B4").result;
      expect(b4 === undefined || b4 === 0).toBe(true);
    });

    it('should classify cell type as "b"/"l"/"v"', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10; // number → "v"
      ws.getCell("A2").value = "text"; // string → "l"
      // A3 blank → "b"
      ws.getCell("A4").value = true; // boolean → "v"
      ws.getCell("B1").value = { formula: 'CELL("type",A1)', result: 0 };
      ws.getCell("B2").value = { formula: 'CELL("type",A2)', result: 0 };
      ws.getCell("B3").value = { formula: 'CELL("type",A3)', result: 0 };
      ws.getCell("B4").value = { formula: 'CELL("type",A4)', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe("v");
      expect(ws.getCell("B2").result).toBe("l");
      expect(ws.getCell("B3").result).toBe("b");
      expect(ws.getCell("B4").result).toBe("v");
    });

    it("should return default width=8 and empty filename", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("B1").value = { formula: 'CELL("width",A1)', result: 0 };
      ws.getCell("B2").value = { formula: 'CELL("filename",A1)', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(8);
      expect(ws.getCell("B2").result).toBe("");
    });

    it("should return #N/A for unsupported info types", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("B1").value = { formula: 'CELL("format",A1)', result: 0 };
      ws.getCell("B2").value = { formula: 'CELL("color",A1)', result: 0 };
      ws.getCell("B3").value = { formula: 'CELL("parentheses",A1)', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toEqual({ error: "#N/A" });
      expect(ws.getCell("B2").result).toEqual({ error: "#N/A" });
      expect(ws.getCell("B3").result).toEqual({ error: "#N/A" });
    });

    it("should resolve the reference produced by INDIRECT", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 99;
      ws.getCell("B1").value = { formula: 'CELL("address",INDIRECT("A1"))', result: 0 };
      ws.getCell("B2").value = { formula: 'CELL("row",INDIRECT("B5"))', result: 0 };
      ws.getCell("B3").value = { formula: 'CELL("contents",INDIRECT("A1"))', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe("$A$1");
      expect(ws.getCell("B2").result).toBe(5);
      expect(ws.getCell("B3").result).toBe(99);
    });

    it("should default to the current cell when reference is omitted", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("C5").value = { formula: 'CELL("address")', result: 0 };
      ws.getCell("C6").value = { formula: 'CELL("row")', result: 0 };
      ws.getCell("C7").value = { formula: 'CELL("col")', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("C5").result).toBe("$C$5");
      expect(ws.getCell("C6").result).toBe(6);
      expect(ws.getCell("C7").result).toBe(3);
    });
  });

  // ==========================================================================
  // ISREF — reference detection at the evaluator level.
  // ==========================================================================
  describe("ISREF function", () => {
    it("should return TRUE for direct cell and range references", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("B1").value = { formula: "ISREF(A1)", result: 0 };
      ws.getCell("B2").value = { formula: "ISREF(A1:B3)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(true);
      expect(ws.getCell("B2").result).toBe(true);
    });

    it("should return FALSE for literals and computed scalars", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("B1").value = { formula: "ISREF(1)", result: 0 };
      ws.getCell("B2").value = { formula: 'ISREF("text")', result: 0 };
      ws.getCell("B3").value = { formula: "ISREF(1+2)", result: 0 };
      ws.getCell("B4").value = { formula: "ISREF(TRUE)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(false);
      expect(ws.getCell("B2").result).toBe(false);
      expect(ws.getCell("B3").result).toBe(false);
      expect(ws.getCell("B4").result).toBe(false);
    });

    it("should return TRUE when INDIRECT produces a valid reference", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 1;
      // INDIRECT yields a ReferenceValue at runtime.
      ws.getCell("B1").value = { formula: 'ISREF(INDIRECT("A1"))', result: 0 };
      ws.getCell("B2").value = { formula: 'ISREF(INDIRECT("A1:B2"))', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(true);
      expect(ws.getCell("B2").result).toBe(true);
    });

    it("should return FALSE when INDIRECT cannot resolve a reference", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      // INDIRECT on an invalid address yields #REF! — ISREF suppresses to FALSE.
      ws.getCell("B1").value = { formula: 'ISREF(INDIRECT("not a ref"))', result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(false);
    });

    it("should return #VALUE! for wrong arity", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("B1").value = { formula: "ISREF(A1,A2)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toEqual({ error: "#VALUE!" });
    });
  });

  // ==========================================================================
  // Language features: intersection operator and external references
  // ==========================================================================
  describe("intersection operator and external references", () => {
    it("should compute intersection of two overlapping ranges (space operator)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // A1:A4 contain 1..4. A1:A3 ∩ A2:A4 = A2:A3 → sum = 2+3 = 5.
      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("A3").value = 3;
      ws.getCell("A4").value = 4;

      ws.getCell("B1").value = { formula: "SUM(A1:A3 A2:A4)", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("B1").result).toBe(5);
    });

    it("should return #NULL! when intersected ranges do not overlap", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("B1").value = 20;

      // A1 and B1 are distinct single-cell refs with no overlap.
      ws.getCell("C1").value = { formula: "A1 B1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("C1").result).toEqual({ error: "#NULL!" });
    });

    it("should return #REF! for an external workbook reference", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");

      // Cross-workbook references like [Book1]Sheet1!A1 are recognised but
      // unsupported — the engine reports #REF! instead of silently
      // swallowing the text.
      ws.getCell("A1").value = { formula: "[Book1]Sheet1!A1", result: 0 };

      wb.calculateFormulas();

      expect(ws.getCell("A1").result).toEqual({ error: "#REF!" });
    });

    it("should return #REF! for OFFSET beyond sheet upper bound", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 1;
      // Row overflow
      ws.getCell("B1").value = { formula: "OFFSET(A1,1048577,0)", result: 0 };
      // Column overflow
      ws.getCell("B2").value = { formula: "OFFSET(A1,0,16384)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toEqual({ error: "#REF!" });
      expect(ws.getCell("B2").result).toEqual({ error: "#REF!" });
    });

    it("should return #NULL! for cross-sheet intersection", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("Sheet1");
      const ws2 = wb.addWorksheet("Sheet2");
      ws1.getCell("A1").value = 1;
      ws1.getCell("A2").value = 2;
      ws1.getCell("A3").value = 3;
      ws2.getCell("A1").value = 10;
      ws2.getCell("A2").value = 20;
      ws2.getCell("A3").value = 30;
      // Intersection across different sheets should be #NULL!
      ws1.getCell("B1").value = {
        formula: "SUM(Sheet1!A1:A2 Sheet2!A1:A3)",
        result: 0
      };
      wb.calculateFormulas();
      expect(ws1.getCell("B1").result).toEqual({ error: "#NULL!" });
    });

    it("should detect all nodes in a diamond dependency cycle", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      // Diamond: A→B→C→A, plus A→D→C
      // All four cells should be in the cycle (not just A, B, C)
      ws.getCell("A1").value = { formula: "B1+D1", result: 0 };
      ws.getCell("B1").value = { formula: "C1+1", result: 0 };
      ws.getCell("C1").value = { formula: "A1*0.5", result: 0 };
      ws.getCell("D1").value = { formula: "C1+2", result: 0 };
      wb.calculateFormulas();
      // All four should resolve to numbers (circular fallback = 0, then one pass)
      expect(typeof ws.getCell("A1").result).toBe("number");
      expect(typeof ws.getCell("B1").result).toBe("number");
      expect(typeof ws.getCell("C1").result).toBe("number");
      expect(typeof ws.getCell("D1").result).toBe("number");
    });

    it("should return #VALUE! when coercing empty string to number", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = { formula: '1+""', result: 0 };
      // Empty cell should coerce to 0
      ws.getCell("A2").value = { formula: "1+B1", result: 0 };
      // Explicit empty string should error
      ws.getCell("B1").value = "";
      ws.getCell("A3").value = { formula: "1+B1", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toEqual({ error: "#VALUE!" });
      // A2 references B1 which is empty string; A3 references B1 too.
      // Depending on whether empty cell and empty string are distinguished,
      // the behavior differs. If B1 reads as empty string, expect #VALUE!
      expect(ws.getCell("A3").result).toEqual({ error: "#VALUE!" });
    });

    it("should return #REF! for external reference nested in a function", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = { formula: "SUM([Book1]Sheet1!A1:A5)", result: 0 };
      ws.getCell("A2").value = { formula: "[Book1]Sheet1!A1 + 1", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toEqual({ error: "#REF!" });
      expect(ws.getCell("A2").result).toEqual({ error: "#REF!" });
    });

    it("should return #NULL! for rectangles disjoint in column axis", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = 2;
      ws.getCell("C1").value = 3;
      ws.getCell("D1").value = 4;
      // A1:B2 and C1:D2 share no cells
      ws.getCell("E1").value = { formula: "SUM(A1:B2 C1:D2)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("E1").result).toEqual({ error: "#NULL!" });
    });
  });

  // ==========================================================================
  // User-registered custom functions
  // ==========================================================================
  describe("registerFunction", () => {
    it("invokes a simple user function inside a formula", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      wb.registerFunction(
        "DOUBLE",
        args => {
          const v = args[0] as { kind: number; value: number };
          return { kind: 1 /* Number */, value: v.value * 2 };
        },
        { minArity: 1, maxArity: 1 }
      );
      ws.getCell("A1").value = 7;
      ws.getCell("B1").value = { formula: "DOUBLE(A1)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(14);
    });

    it("is case-insensitive and accepts _XLFN. prefixed names", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      wb.registerFunction("answer", () => ({ kind: 1, value: 42 }));
      ws.getCell("A1").value = { formula: "ANSWER()", result: 0 };
      ws.getCell("A2").value = { formula: "_xlfn.ANSWER()", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toBe(42);
      expect(ws.getCell("A2").result).toBe(42);
    });

    it("shadows built-in functions of the same name", () => {
      // Registering `SUM` as a custom always-returns-99 function should
      // override the built-in aggregator for this workbook only.
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      wb.registerFunction("SUM", () => ({ kind: 1, value: 99 }));
      ws.getCell("A1").value = 1;
      ws.getCell("A2").value = 2;
      ws.getCell("B1").value = { formula: "SUM(A1:A2)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(99);
    });

    it("arity validation rejects wrong arg count with #VALUE!", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      wb.registerFunction("TWO", () => ({ kind: 1, value: 0 }), {
        minArity: 2,
        maxArity: 2
      });
      ws.getCell("A1").value = { formula: "TWO(1)", result: 0 };
      ws.getCell("A2").value = { formula: "TWO(1,2)", result: 0 };
      ws.getCell("A3").value = { formula: "TWO(1,2,3)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toEqual({ error: "#VALUE!" });
      expect(ws.getCell("A2").result).toBe(0);
      expect(ws.getCell("A3").result).toEqual({ error: "#VALUE!" });
    });

    it("unregisterFunction removes the entry", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      wb.registerFunction("TMP", () => ({ kind: 1, value: 42 }));
      // With TMP registered, the formula evaluates to 42.
      ws.getCell("A1").value = { formula: "TMP()", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toBe(42);

      // Remove the registration; subsequent calculate should surface
      // #NAME? via the engine's unknown-function path.
      expect(wb.unregisterFunction("TMP")).toBe(true);
      expect(wb.unregisterFunction("TMP")).toBe(false);
      // Reset the cached result so the engine has no prior value to
      // preserve (the "preserve cached result on #NAME?" path would
      // otherwise hide the behaviour we're testing).
      ws.getCell("A1").value = { formula: "TMP()" };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toEqual({ error: "#NAME?" });
    });

    it("throwing user function surfaces as #VALUE! (doesn't tear down calc)", () => {
      // Regression: without a try/catch boundary a buggy custom
      // function would throw through the evaluator and leave the
      // whole calculation pass half-done.
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      wb.registerFunction("BOOM", () => {
        throw new Error("boom");
      });
      ws.getCell("A1").value = { formula: "BOOM()", result: 0 };
      // Other formulas should still complete successfully.
      ws.getCell("A2").value = 5;
      ws.getCell("A3").value = { formula: "A2 * 2", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("A1").result).toEqual({ error: "#VALUE!" });
      expect(ws.getCell("A3").result).toBe(10);
    });

    it("user function composes with built-ins (IF, SUM, etc.)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      wb.registerFunction("NEGATE", args => {
        const v = args[0] as { kind: number; value: number };
        return { kind: 1, value: -v.value };
      });
      ws.getCell("A1").value = 5;
      ws.getCell("A2").value = 3;
      ws.getCell("B1").value = { formula: "IF(A1>0, NEGATE(A1), A2)", result: 0 };
      ws.getCell("B2").value = { formula: "SUM(NEGATE(A1), A2)", result: 0 };
      wb.calculateFormulas();
      expect(ws.getCell("B1").result).toBe(-5); // IF true branch runs NEGATE
      expect(ws.getCell("B2").result).toBe(-2); // -5 + 3 = -2
    });
  });
});
