import { Workbook } from "@excel/workbook";
import { describe, it, expect } from "vitest";

/**
 * XLSX Round-Trip Formula Tests
 *
 * These tests exercise the FULL pipeline:
 *   1. Create a workbook with data + formulas
 *   2. Write to XLSX buffer via wb.xlsx.writeBuffer()
 *   3. Read back via new Workbook() + wb.xlsx.load(buffer)
 *   4. Call wb.calculateFormulas()
 *   5. Verify ALL formula results match expected values
 *
 * This is not a unit test with mock data — it tests real XLSX
 * serialization → deserialization → formula evaluation.
 */
describe("Formula XLSX Round-Trip", () => {
  // ==========================================================================
  // Test 1: Complex workbook with mixed data types and diverse formulas
  // ==========================================================================
  describe("complex workbook round-trip", () => {
    it("should create, write, read, and calculate a workbook with diverse formulas", async () => {
      // ── Step 1: Create the workbook ──────────────────────────────────────
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Sheet1");

      // Row 1: Headers
      ws1.getCell("A1").value = "Name";
      ws1.getCell("B1").value = "Score";
      ws1.getCell("C1").value = "Grade";
      ws1.getCell("D1").value = "Active";

      // Row 2–6: Data (numbers, strings, booleans)
      ws1.getCell("A2").value = "Alice";
      ws1.getCell("B2").value = 92;
      ws1.getCell("C2").value = "A";
      ws1.getCell("D2").value = true;

      ws1.getCell("A3").value = "Bob";
      ws1.getCell("B3").value = 78;
      ws1.getCell("C3").value = "C";
      ws1.getCell("D3").value = true;

      ws1.getCell("A4").value = "Charlie";
      ws1.getCell("B4").value = 85;
      ws1.getCell("C4").value = "B";
      ws1.getCell("D4").value = false;

      ws1.getCell("A5").value = "Diana";
      ws1.getCell("B5").value = 95;
      ws1.getCell("C5").value = "A";
      ws1.getCell("D5").value = true;

      ws1.getCell("A6").value = "Eve";
      ws1.getCell("B6").value = 62;
      ws1.getCell("C6").value = "D";
      ws1.getCell("D6").value = false;

      // Row 7–10: More numeric data
      ws1.getCell("A7").value = "Frank";
      ws1.getCell("B7").value = 88;
      ws1.getCell("C7").value = "B";
      ws1.getCell("D7").value = true;

      ws1.getCell("A8").value = "Grace";
      ws1.getCell("B8").value = 71;
      ws1.getCell("C8").value = "C";
      ws1.getCell("D8").value = true;

      ws1.getCell("A9").value = "Hank";
      ws1.getCell("B9").value = 43;
      ws1.getCell("C9").value = "F";
      ws1.getCell("D9").value = false;

      ws1.getCell("A10").value = "Ivy";
      ws1.getCell("B10").value = 99;
      ws1.getCell("C10").value = "A";
      ws1.getCell("D10").value = true;

      // ── Formula cells ────────────────────────────────────────────────────

      // F1: SUM of all scores
      ws1.getCell("F1").value = { formula: "SUM(B2:B10)", result: 0 };

      // F2: AVERAGE of scores
      ws1.getCell("F2").value = { formula: "AVERAGE(B2:B10)", result: 0 };

      // F3: MAX score
      ws1.getCell("F3").value = { formula: "MAX(B2:B10)", result: 0 };

      // F4: MIN score
      ws1.getCell("F4").value = { formula: "MIN(B2:B10)", result: 0 };

      // F5: COUNT of numeric scores
      ws1.getCell("F5").value = { formula: "COUNT(B2:B10)", result: 0 };

      // F6: COUNTA (all non-empty in A column)
      ws1.getCell("F6").value = { formula: "COUNTA(A2:A10)", result: 0 };

      // F7: IF — check if average > 80
      ws1.getCell("F7").value = {
        formula: 'IF(AVERAGE(B2:B10)>80,"Pass","Fail")',
        result: ""
      };

      // F8: CONCATENATE
      ws1.getCell("F8").value = {
        formula: 'CONCATENATE(A2," scored ",B2)',
        result: ""
      };

      // F9: VLOOKUP — look up Bob's score
      ws1.getCell("F9").value = {
        formula: 'VLOOKUP("Bob",A2:B10,2,FALSE)',
        result: 0
      };

      // F10: Nested IF (grade based on average)
      ws1.getCell("F10").value = {
        formula:
          'IF(AVERAGE(B2:B10)>=90,"A",IF(AVERAGE(B2:B10)>=80,"B",IF(AVERAGE(B2:B10)>=70,"C","D")))',
        result: ""
      };

      // F11: COUNTIF — count "A" grades
      ws1.getCell("F11").value = {
        formula: 'COUNTIF(C2:C10,"A")',
        result: 0
      };

      // F12: SUMIF — sum scores of active students
      ws1.getCell("F12").value = {
        formula: "SUMIF(D2:D10,TRUE,B2:B10)",
        result: 0
      };

      // F13: ROUND of average
      ws1.getCell("F13").value = {
        formula: "ROUND(AVERAGE(B2:B10),1)",
        result: 0
      };

      // F14: Chained dependency — F1 * 2
      ws1.getCell("F14").value = { formula: "F1*2", result: 0 };

      // F15: DATE function
      ws1.getCell("F15").value = {
        formula: "YEAR(DATE(2024,6,15))",
        result: 0
      };

      // F16: LEN of a name
      ws1.getCell("F16").value = { formula: "LEN(A4)", result: 0 };

      // F17: LEFT
      ws1.getCell("F17").value = { formula: "LEFT(A4,3)", result: "" };

      // F18: IFERROR with safe expression
      ws1.getCell("F18").value = {
        formula: 'IFERROR(B2/B3,"Error")',
        result: 0
      };

      // F19: PRODUCT
      ws1.getCell("F19").value = { formula: "PRODUCT(B2,B3)", result: 0 };

      // ── Named range ──────────────────────────────────────────────────────
      // Define "Scores" as Sheet1!$B$2:$B$10
      wb1.definedNames.add("Sheet1!$B$2", "Scores");
      wb1.definedNames.add("Sheet1!$B$3", "Scores");
      wb1.definedNames.add("Sheet1!$B$4", "Scores");
      wb1.definedNames.add("Sheet1!$B$5", "Scores");
      wb1.definedNames.add("Sheet1!$B$6", "Scores");
      wb1.definedNames.add("Sheet1!$B$7", "Scores");
      wb1.definedNames.add("Sheet1!$B$8", "Scores");
      wb1.definedNames.add("Sheet1!$B$9", "Scores");
      wb1.definedNames.add("Sheet1!$B$10", "Scores");

      // F20: Formula using named range
      ws1.getCell("F20").value = { formula: "SUM(Scores)", result: 0 };

      // ── Step 2: Write to buffer ──────────────────────────────────────────
      const buffer = await wb1.xlsx.writeBuffer();
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(0);

      // ── Step 3: Read back ────────────────────────────────────────────────
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);

      const ws2 = wb2.getWorksheet("Sheet1")!;
      expect(ws2).toBeDefined();

      // Verify raw data survived the round-trip
      expect(ws2.getCell("A2").value).toBe("Alice");
      expect(ws2.getCell("B2").value).toBe(92);
      expect(ws2.getCell("D2").value).toBe(true);

      // ── Step 4: Calculate formulas ───────────────────────────────────────
      wb2.calculateFormulas();

      // ── Step 5: Verify ALL formula results ───────────────────────────────
      // Expected scores: 92, 78, 85, 95, 62, 88, 71, 43, 99
      const expectedSum = 92 + 78 + 85 + 95 + 62 + 88 + 71 + 43 + 99; // 713
      const expectedAvg = expectedSum / 9; // ~79.222...

      // F1: SUM
      expect(ws2.getCell("F1").result).toBe(expectedSum);

      // F2: AVERAGE
      expect(ws2.getCell("F2").result).toBeCloseTo(expectedAvg, 5);

      // F3: MAX
      expect(ws2.getCell("F3").result).toBe(99);

      // F4: MIN
      expect(ws2.getCell("F4").result).toBe(43);

      // F5: COUNT
      expect(ws2.getCell("F5").result).toBe(9);

      // F6: COUNTA
      expect(ws2.getCell("F6").result).toBe(9);

      // F7: IF — average is ~79.2, which is < 80
      expect(ws2.getCell("F7").result).toBe("Fail");

      // F8: CONCATENATE
      expect(ws2.getCell("F8").result).toBe("Alice scored 92");

      // F9: VLOOKUP
      expect(ws2.getCell("F9").result).toBe(78);

      // F10: Nested IF — avg ~79.2, so falls into >=70 branch → "C"
      expect(ws2.getCell("F10").result).toBe("C");

      // F11: COUNTIF "A" grades — Alice, Diana, Ivy → 3
      expect(ws2.getCell("F11").result).toBe(3);

      // F12: SUMIF active=TRUE scores: 92+78+95+88+71+99 = 523
      expect(ws2.getCell("F12").result).toBe(523);

      // F13: ROUND(avg, 1) = ROUND(79.222..., 1) = 79.2
      expect(ws2.getCell("F13").result).toBeCloseTo(79.2, 1);

      // F14: chained — SUM * 2
      expect(ws2.getCell("F14").result).toBe(expectedSum * 2);

      // F15: YEAR(DATE(2024,6,15)) = 2024
      expect(ws2.getCell("F15").result).toBe(2024);

      // F16: LEN("Charlie") = 7
      expect(ws2.getCell("F16").result).toBe(7);

      // F17: LEFT("Charlie", 3) = "Cha"
      expect(ws2.getCell("F17").result).toBe("Cha");

      // F18: IFERROR(92/78, "Error") ≈ 1.179...
      expect(ws2.getCell("F18").result).toBeCloseTo(92 / 78, 5);

      // F19: PRODUCT(92, 78) = 7176
      expect(ws2.getCell("F19").result).toBe(92 * 78);

      // F20: SUM(Scores) via named range = same as SUM(B2:B10)
      expect(ws2.getCell("F20").result).toBe(expectedSum);
    });
  });

  // ==========================================================================
  // Test 2: Every reference type
  // ==========================================================================
  describe("all reference types round-trip", () => {
    it("should handle simple, absolute, cross-sheet, range, and named range references", async () => {
      // ── Step 1: Create workbook with two sheets ──────────────────────────
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Sheet1");
      const ws2sheet = wb1.addWorksheet("Sheet2");

      // Sheet1 data
      ws1.getCell("A1").value = 10;
      ws1.getCell("A2").value = 20;
      ws1.getCell("A3").value = 30;
      ws1.getCell("B1").value = 100;
      ws1.getCell("B2").value = 200;
      ws1.getCell("B3").value = 300;

      // Sheet2 data
      ws2sheet.getCell("A1").value = 1000;
      ws2sheet.getCell("A2").value = 2000;
      ws2sheet.getCell("A3").value = 3000;

      // Named range: "MyValues" → Sheet1!$A$1:$A$3
      wb1.definedNames.add("Sheet1!$A$1", "MyValues");
      wb1.definedNames.add("Sheet1!$A$2", "MyValues");
      wb1.definedNames.add("Sheet1!$A$3", "MyValues");

      // Named range: "Sheet2Total" → single cell Sheet2!$A$1
      wb1.definedNames.add("Sheet2!$A$1", "Sheet2Val");

      // ── Formulas exercising every reference type ─────────────────────────

      // D1: Simple reference — A1
      ws1.getCell("D1").value = { formula: "A1+5", result: 0 };

      // D2: Absolute reference — $A$1
      ws1.getCell("D2").value = { formula: "$A$1+$B$1", result: 0 };

      // D3: Mixed reference — $A2 (absolute col, relative row)
      ws1.getCell("D3").value = { formula: "$A2*3", result: 0 };

      // D4: Range reference — SUM(A1:A3)
      ws1.getCell("D4").value = { formula: "SUM(A1:A3)", result: 0 };

      // D5: Multi-column range — SUM(A1:B3)
      ws1.getCell("D5").value = { formula: "SUM(A1:B3)", result: 0 };

      // D6: Cross-sheet reference — Sheet2!A1
      ws1.getCell("D6").value = { formula: "Sheet2!A1", result: 0 };

      // D7: Cross-sheet range — SUM(Sheet2!A1:A3)
      ws1.getCell("D7").value = { formula: "SUM(Sheet2!A1:A3)", result: 0 };

      // D8: Named range — SUM(MyValues)
      ws1.getCell("D8").value = { formula: "SUM(MyValues)", result: 0 };

      // D9: Named single-cell range
      ws1.getCell("D9").value = { formula: "Sheet2Val*2", result: 0 };

      // D10: Complex — combines cross-sheet + range + arithmetic
      ws1.getCell("D10").value = {
        formula: "SUM(A1:A3)+Sheet2!A1",
        result: 0
      };

      // D11: Absolute in function — AVERAGE($B$1:$B$3)
      ws1.getCell("D11").value = {
        formula: "AVERAGE($B$1:$B$3)",
        result: 0
      };

      // D12: Cross-sheet with absolute — Sheet2!$A$2
      ws1.getCell("D12").value = { formula: "Sheet2!$A$2+1", result: 0 };

      // D13: Named range in expression — SUM(MyValues) * 10
      ws1.getCell("D13").value = {
        formula: "SUM(MyValues)*10",
        result: 0
      };

      // ── Also add a formula on Sheet2 referencing Sheet1 ──────────────────
      ws2sheet.getCell("B1").value = {
        formula: "SUM(Sheet1!A1:A3)",
        result: 0
      };

      ws2sheet.getCell("B2").value = {
        formula: "Sheet1!B2+A1",
        result: 0
      };

      // ── Step 2: Write to buffer ──────────────────────────────────────────
      const buffer = await wb1.xlsx.writeBuffer();

      // ── Step 3: Read back ────────────────────────────────────────────────
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);

      const s1 = wb2.getWorksheet("Sheet1")!;
      const s2 = wb2.getWorksheet("Sheet2")!;
      expect(s1).toBeDefined();
      expect(s2).toBeDefined();

      // Verify data survived
      expect(s1.getCell("A1").value).toBe(10);
      expect(s1.getCell("B3").value).toBe(300);
      expect(s2.getCell("A1").value).toBe(1000);

      // ── Step 4: Calculate ────────────────────────────────────────────────
      wb2.calculateFormulas();

      // ── Step 5: Verify results ───────────────────────────────────────────

      // D1: A1 + 5 = 10 + 5 = 15
      expect(s1.getCell("D1").result).toBe(15);

      // D2: $A$1 + $B$1 = 10 + 100 = 110
      expect(s1.getCell("D2").result).toBe(110);

      // D3: $A2 * 3 = 20 * 3 = 60
      expect(s1.getCell("D3").result).toBe(60);

      // D4: SUM(A1:A3) = 10 + 20 + 30 = 60
      expect(s1.getCell("D4").result).toBe(60);

      // D5: SUM(A1:B3) = (10+20+30) + (100+200+300) = 660
      expect(s1.getCell("D5").result).toBe(660);

      // D6: Sheet2!A1 = 1000
      expect(s1.getCell("D6").result).toBe(1000);

      // D7: SUM(Sheet2!A1:A3) = 1000 + 2000 + 3000 = 6000
      expect(s1.getCell("D7").result).toBe(6000);

      // D8: SUM(MyValues) = SUM(A1:A3) = 60
      expect(s1.getCell("D8").result).toBe(60);

      // D9: Sheet2Val * 2 = 1000 * 2 = 2000
      expect(s1.getCell("D9").result).toBe(2000);

      // D10: SUM(A1:A3) + Sheet2!A1 = 60 + 1000 = 1060
      expect(s1.getCell("D10").result).toBe(1060);

      // D11: AVERAGE($B$1:$B$3) = (100+200+300)/3 = 200
      expect(s1.getCell("D11").result).toBe(200);

      // D12: Sheet2!$A$2 + 1 = 2000 + 1 = 2001
      expect(s1.getCell("D12").result).toBe(2001);

      // D13: SUM(MyValues) * 10 = 60 * 10 = 600
      expect(s1.getCell("D13").result).toBe(600);

      // Sheet2 B1: SUM(Sheet1!A1:A3) = 60
      expect(s2.getCell("B1").result).toBe(60);

      // Sheet2 B2: Sheet1!B2 + A1 = 200 + 1000 = 1200
      expect(s2.getCell("B2").result).toBe(1200);
    });
  });

  // ==========================================================================
  // Test 3: Modify data after load, then recalculate
  // ==========================================================================
  describe("modify after load and recalculate", () => {
    it("should produce correct results after modifying data post-load", async () => {
      // Create original workbook with known values
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Sheet1");
      ws1.getCell("A1").value = 10;
      ws1.getCell("A2").value = 20;
      ws1.getCell("A3").value = 30;
      ws1.getCell("B1").value = { formula: "SUM(A1:A3)", result: 60 };
      ws1.getCell("B2").value = { formula: "AVERAGE(A1:A3)", result: 20 };
      ws1.getCell("B3").value = { formula: "A1*A2+A3", result: 230 };
      ws1.getCell("B4").value = {
        formula: 'IF(B1>100,"big","small")',
        result: "small"
      };
      ws1.getCell("B5").value = { formula: "B1+B3", result: 290 };

      // Write
      const buffer = await wb1.xlsx.writeBuffer();

      // Read back
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);
      const ws2 = wb2.getWorksheet("Sheet1")!;

      // Modify the data cells — cached formula results are now stale
      ws2.getCell("A1").value = 100;
      ws2.getCell("A2").value = 200;
      ws2.getCell("A3").value = 300;

      // Recalculate
      wb2.calculateFormulas();

      // Verify with new data
      expect(ws2.getCell("B1").result).toBe(600); // SUM(100,200,300)
      expect(ws2.getCell("B2").result).toBe(200); // AVERAGE(100,200,300)
      expect(ws2.getCell("B3").result).toBe(100 * 200 + 300); // 20300
      expect(ws2.getCell("B4").result).toBe("big"); // 600 > 100
      expect(ws2.getCell("B5").result).toBe(600 + 20300); // B1 + B3 = 20900
    });
  });

  // ==========================================================================
  // Test 4: Multi-sheet cross-references through round-trip
  // ==========================================================================
  describe("multi-sheet cross-reference round-trip", () => {
    it("should maintain cross-sheet formula relationships after round-trip", async () => {
      const wb1 = new Workbook();

      // Sheet: Revenue
      const revenue = wb1.addWorksheet("Revenue");
      revenue.getCell("A1").value = "Q1";
      revenue.getCell("B1").value = 10000;
      revenue.getCell("A2").value = "Q2";
      revenue.getCell("B2").value = 15000;
      revenue.getCell("A3").value = "Q3";
      revenue.getCell("B3").value = 12000;
      revenue.getCell("A4").value = "Q4";
      revenue.getCell("B4").value = 20000;
      revenue.getCell("A5").value = "Total";
      revenue.getCell("B5").value = { formula: "SUM(B1:B4)", result: 0 };

      // Sheet: Expenses
      const expenses = wb1.addWorksheet("Expenses");
      expenses.getCell("A1").value = "Q1";
      expenses.getCell("B1").value = 8000;
      expenses.getCell("A2").value = "Q2";
      expenses.getCell("B2").value = 9000;
      expenses.getCell("A3").value = "Q3";
      expenses.getCell("B3").value = 7500;
      expenses.getCell("A4").value = "Q4";
      expenses.getCell("B4").value = 11000;
      expenses.getCell("B5").value = { formula: "SUM(B1:B4)", result: 0 };

      // Sheet: Summary — references both other sheets
      const summary = wb1.addWorksheet("Summary");
      summary.getCell("A1").value = "Total Revenue";
      summary.getCell("B1").value = {
        formula: "Revenue!B5",
        result: 0
      };
      summary.getCell("A2").value = "Total Expenses";
      summary.getCell("B2").value = {
        formula: "Expenses!B5",
        result: 0
      };
      summary.getCell("A3").value = "Net Profit";
      summary.getCell("B3").value = { formula: "B1-B2", result: 0 };
      summary.getCell("A4").value = "Profit Margin";
      summary.getCell("B4").value = {
        formula: "ROUND(B3/B1*100,1)",
        result: 0
      };

      // Write → Read
      const buffer = await wb1.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);

      // Calculate
      wb2.calculateFormulas();

      const s = wb2.getWorksheet("Summary")!;
      const r = wb2.getWorksheet("Revenue")!;
      const e = wb2.getWorksheet("Expenses")!;

      // Revenue total: 10000+15000+12000+20000 = 57000
      expect(r.getCell("B5").result).toBe(57000);

      // Expenses total: 8000+9000+7500+11000 = 35500
      expect(e.getCell("B5").result).toBe(35500);

      // Summary references
      expect(s.getCell("B1").result).toBe(57000);
      expect(s.getCell("B2").result).toBe(35500);
      expect(s.getCell("B3").result).toBe(21500); // Net profit
      // Profit margin: ROUND(21500/57000*100, 1) = ROUND(37.719..., 1) = 37.7
      expect(s.getCell("B4").result).toBeCloseTo(37.7, 1);
    });
  });

  // ==========================================================================
  // Test 5: Formula types variety — INDEX/MATCH, logical, text, math
  // ==========================================================================
  describe("formula variety round-trip", () => {
    it("should correctly evaluate INDEX/MATCH, SUMPRODUCT, and text formulas after round-trip", async () => {
      const wb1 = new Workbook();
      const ws = wb1.addWorksheet("Sheet1");

      // Lookup table
      ws.getCell("A1").value = "Product";
      ws.getCell("B1").value = "Price";
      ws.getCell("C1").value = "Qty";
      ws.getCell("A2").value = "Widget";
      ws.getCell("B2").value = 9.99;
      ws.getCell("C2").value = 5;
      ws.getCell("A3").value = "Gadget";
      ws.getCell("B3").value = 24.5;
      ws.getCell("C3").value = 3;
      ws.getCell("A4").value = "Doohickey";
      ws.getCell("B4").value = 14.75;
      ws.getCell("C4").value = 8;

      // E1: INDEX+MATCH — find Gadget's price
      ws.getCell("E1").value = {
        formula: 'INDEX(B2:B4,MATCH("Gadget",A2:A4,0))',
        result: 0
      };

      // E2: SUMPRODUCT — total revenue
      ws.getCell("E2").value = {
        formula: "SUMPRODUCT(B2:B4,C2:C4)",
        result: 0
      };

      // E3: UPPER
      ws.getCell("E3").value = { formula: "UPPER(A2)", result: "" };

      // E4: MID
      ws.getCell("E4").value = { formula: "MID(A4,4,3)", result: "" };

      // E5: ABS of negative
      ws.getCell("E5").value = { formula: "ABS(-42.5)", result: 0 };

      // E6: MOD
      ws.getCell("E6").value = { formula: "MOD(17,5)", result: 0 };

      // E7: POWER
      ws.getCell("E7").value = { formula: "POWER(2,10)", result: 0 };

      // Write → Read → Calculate
      const buffer = await wb1.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);
      wb2.calculateFormulas();

      const ws2 = wb2.getWorksheet("Sheet1")!;

      // E1: Gadget price = 24.5
      expect(ws2.getCell("E1").result).toBe(24.5);

      // E2: SUMPRODUCT = 9.99*5 + 24.5*3 + 14.75*8 = 49.95+73.5+118 = 241.45
      expect(ws2.getCell("E2").result).toBeCloseTo(241.45, 2);

      // E3: UPPER("Widget") = "WIDGET"
      expect(ws2.getCell("E3").result).toBe("WIDGET");

      // E4: MID("Doohickey", 4, 3) = "hic"
      expect(ws2.getCell("E4").result).toBe("hic");

      // E5: ABS(-42.5) = 42.5
      expect(ws2.getCell("E5").result).toBe(42.5);

      // E6: MOD(17, 5) = 2
      expect(ws2.getCell("E6").result).toBe(2);

      // E7: POWER(2, 10) = 1024
      expect(ws2.getCell("E7").result).toBe(1024);
    });
  });

  // ==========================================================================
  // Test 6: Shared formulas through round-trip
  // ==========================================================================
  describe("shared formulas round-trip", () => {
    it("should handle fillFormula through round-trip correctly", async () => {
      const wb1 = new Workbook();
      const ws = wb1.addWorksheet("Sheet1");

      // Data
      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 20;
      ws.getCell("A3").value = 30;
      ws.getCell("A4").value = 40;
      ws.getCell("A5").value = 50;

      // Shared formula: B1=A1*2, B2=A2*2, ... B5=A5*2
      ws.fillFormula("B1:B5", "A1*2", [20, 40, 60, 80, 100]);

      // Write → Read → Calculate
      const buffer = await wb1.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);
      wb2.calculateFormulas();

      const ws2 = wb2.getWorksheet("Sheet1")!;

      expect(ws2.getCell("B1").result).toBe(20);
      expect(ws2.getCell("B2").result).toBe(40);
      expect(ws2.getCell("B3").result).toBe(60);
      expect(ws2.getCell("B4").result).toBe(80);
      expect(ws2.getCell("B5").result).toBe(100);
    });
  });

  // ==========================================================================
  // Test 7: Error formulas survive round-trip
  // ==========================================================================
  describe("error handling round-trip", () => {
    it("should correctly produce errors after round-trip", async () => {
      const wb1 = new Workbook();
      const ws = wb1.addWorksheet("Sheet1");

      ws.getCell("A1").value = 10;
      ws.getCell("A2").value = 0;

      // Division by zero
      ws.getCell("B1").value = { formula: "A1/A2", result: 0 };

      // IFERROR wrapping division by zero
      ws.getCell("B2").value = {
        formula: 'IFERROR(A1/A2,"Div/0!")',
        result: ""
      };

      // VLOOKUP not found
      ws.getCell("C1").value = "Apple";
      ws.getCell("D1").value = 1;
      ws.getCell("B3").value = {
        formula: 'VLOOKUP("Mango",C1:D1,2,FALSE)',
        result: 0
      };

      // Write → Read → Calculate
      const buffer = await wb1.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer as Buffer);
      wb2.calculateFormulas();

      const ws2 = wb2.getWorksheet("Sheet1")!;

      // B1: #DIV/0!
      expect(ws2.getCell("B1").result).toEqual({ error: "#DIV/0!" });

      // B2: IFERROR catches it
      expect(ws2.getCell("B2").result).toBe("Div/0!");

      // B3: #N/A (not found)
      expect(ws2.getCell("B3").result).toEqual({ error: "#N/A" });
    });
  });

  // ==========================================================================
  // Test 8: Double round-trip (write → read → write → read → calculate)
  // ==========================================================================
  describe("double round-trip", () => {
    it("should survive write→read→write→read→calculate", async () => {
      // First workbook
      const wb1 = new Workbook();
      const ws1 = wb1.addWorksheet("Sheet1");
      ws1.getCell("A1").value = 7;
      ws1.getCell("A2").value = 8;
      ws1.getCell("A3").value = { formula: "A1+A2", result: 15 };
      ws1.getCell("A4").value = { formula: "A3*A1", result: 105 };

      // First round-trip
      const buf1 = await wb1.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buf1 as Buffer);

      // Second round-trip
      const buf2 = await wb2.xlsx.writeBuffer();
      const wb3 = new Workbook();
      await wb3.xlsx.load(buf2 as Buffer);

      // Calculate on the doubly-round-tripped workbook
      wb3.calculateFormulas();

      const ws3 = wb3.getWorksheet("Sheet1")!;
      expect(ws3.getCell("A3").result).toBe(15);
      expect(ws3.getCell("A4").result).toBe(105);
    });
  });
});
