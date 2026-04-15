import { testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

const TEST_FILE = testFilePath("ignored-errors.test");

describe("ignoredErrors", () => {
  it("round-trips ignoredErrors through write and read", async () => {
    // Create workbook with ignoredErrors
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");

    ws.getCell("A1").value = "123";
    ws.getCell("A2").value = "456";
    ws.getCell("B1").value = "=SUM(A1:A2)";

    ws.ignoredErrors = [
      { ref: "A1:A2", numberStoredAsText: true },
      { ref: "B1:B10", formula: true, evalError: true }
    ];

    // Write
    await wb.xlsx.writeFile(TEST_FILE);

    // Read back
    const wb2 = new Workbook();
    await wb2.xlsx.readFile(TEST_FILE);
    const ws2 = wb2.getWorksheet("Sheet1")!;

    // Verify
    expect(ws2.ignoredErrors).toHaveLength(2);
    expect(ws2.ignoredErrors[0]).toEqual({
      ref: "A1:A2",
      numberStoredAsText: true
    });
    expect(ws2.ignoredErrors[1]).toEqual({
      ref: "B1:B10",
      formula: true,
      evalError: true
    });
  });

  it("preserves empty ignoredErrors (no extra XML)", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "test";
    // ignoredErrors defaults to [] — should not write <ignoredErrors> tag

    const buffer = await wb.xlsx.writeBuffer();

    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet("Sheet1")!;
    expect(ws2.ignoredErrors).toEqual([]);
  });

  it("round-trips all boolean attributes", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "data";

    ws.ignoredErrors = [
      {
        ref: "A1:Z100",
        numberStoredAsText: true,
        formula: true,
        formulaRange: true,
        unlockedFormula: true,
        emptyCellReference: true,
        listDataValidation: true,
        calculatedColumn: true,
        evalError: true,
        twoDigitTextYear: true
      }
    ];

    const buffer = await wb.xlsx.writeBuffer();

    const wb2 = new Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet("Sheet1")!;

    expect(ws2.ignoredErrors).toHaveLength(1);
    const entry = ws2.ignoredErrors[0];
    expect(entry.ref).toBe("A1:Z100");
    expect(entry.numberStoredAsText).toBe(true);
    expect(entry.formula).toBe(true);
    expect(entry.formulaRange).toBe(true);
    expect(entry.unlockedFormula).toBe(true);
    expect(entry.emptyCellReference).toBe(true);
    expect(entry.listDataValidation).toBe(true);
    expect(entry.calculatedColumn).toBe(true);
    expect(entry.evalError).toBe(true);
    expect(entry.twoDigitTextYear).toBe(true);
  });

  it("works with streaming writer", async () => {
    const { WorkbookWriter } = await import("@excel/stream/workbook-writer");

    const wb = new WorkbookWriter({ filename: TEST_FILE });
    const ws = wb.addWorksheet("Sheet1");

    ws.ignoredErrors = [{ ref: "A1:A100", numberStoredAsText: true }];

    ws.getCell("A1").value = "123";
    ws.commit();
    await wb.commit();

    // Read back with standard reader
    const wb2 = new Workbook();
    await wb2.xlsx.readFile(TEST_FILE);
    const ws2 = wb2.getWorksheet("Sheet1")!;

    expect(ws2.ignoredErrors).toHaveLength(1);
    expect(ws2.ignoredErrors[0]).toEqual({
      ref: "A1:A100",
      numberStoredAsText: true
    });
  });
});
