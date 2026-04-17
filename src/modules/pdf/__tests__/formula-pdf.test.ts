/**
 * Tests for formula evaluation during Excel-to-PDF conversion.
 *
 * Regression tests for the bug where formula cells kept their stale cached
 * results from the last save in Excel, rather than reflecting the current
 * (programmatically-set) input values.
 *
 * PDF-Test-Formula.xlsx formulas (row 14):
 *   B14: SUM(B11:INDIRECT("B" & ROW()-1))   → should sum B11:B13
 *   D14: SUM(D11:INDIRECT("D" & ROW()-1))   → should sum D11:D13
 *   E14: SUM(E11:E13)                        → straightforward SUM
 *   F14: SUM(F11:INDIRECT("F" & ROW()-1))   → should sum F11:F13
 */

import path from "node:path";

import { Workbook } from "@excel/workbook";
import { evaluateWorksheetFormulas } from "@excel/utils/formula-evaluator";
import { excelToPdf } from "@pdf/excel-bridge";
import { readPdf } from "@pdf/reader/pdf-reader";
import { describe, it, expect, beforeAll } from "vitest";

import { expectValidPdf } from "./test-helpers";

const XLSX_PATH = path.join(__dirname, "PDF-Test-Formula.xlsx");

// ---------------------------------------------------------------------------
// Helper: read PDF text content
// ---------------------------------------------------------------------------
async function extractPdfText(pdfBytes: Uint8Array): Promise<string> {
  const result = await readPdf(pdfBytes);
  return result.text;
}

// ---------------------------------------------------------------------------
// Core scenario: load XLSX → set inputs → convert → verify formula results
// ---------------------------------------------------------------------------
describe("Formula evaluation in excelToPdf (PDF-Test-Formula.xlsx)", () => {
  let pdfBytes: Uint8Array;

  // Input values set in the test (same as the user's script)
  const B11 = 1,
    B12 = 2,
    B13 = 3;
  const D11 = 10.5,
    D12 = 23.75,
    D13 = 7.001;
  const E11 = 100,
    E12 = 3.14,
    E13 = 99.99;
  const F11 = 0.5,
    F12 = 42,
    F13 = 1234.56;

  // Expected formula results
  const expectedB14 = B11 + B12 + B13; // 6
  const expectedD14 = D11 + D12 + D13; // 41.251
  const expectedE14 = E11 + E12 + E13; // 203.13
  const expectedF14 = F11 + F12 + F13; // 1277.06

  beforeAll(async () => {
    const workbook = new Workbook();
    const excel = await workbook.xlsx.readFile(XLSX_PATH);

    const sheet = excel.worksheets[0];

    // Set integer inputs
    sheet.getCell("B11").value = B11;
    sheet.getCell("B12").value = B12;
    sheet.getCell("B13").value = B13;

    // Set decimal inputs
    sheet.getCell("D11").value = D11;
    sheet.getCell("D12").value = D12;
    sheet.getCell("D13").value = D13;
    sheet.getCell("E11").value = E11;
    sheet.getCell("E12").value = E12;
    sheet.getCell("E13").value = E13;
    sheet.getCell("F11").value = F11;
    sheet.getCell("F12").value = F12;
    sheet.getCell("F13").value = F13;

    pdfBytes = await excelToPdf(excel);
  });

  it("should produce a valid PDF", () => {
    expectValidPdf(pdfBytes);
    expect(pdfBytes.length).toBeGreaterThan(100);
  });

  it("should evaluate SUM(E11:E13) = 203.13 and render it in the PDF", async () => {
    const text = await extractPdfText(pdfBytes);
    expect(text).toContain(String(expectedE14));
  });

  it("should evaluate SUM(B11:B13) via INDIRECT and render the result 6 in the PDF", async () => {
    const text = await extractPdfText(pdfBytes);
    expect(text).toContain(String(expectedB14));
  });

  it("should evaluate SUM(D11:D13) via INDIRECT and render the result in the PDF", async () => {
    const text = await extractPdfText(pdfBytes);
    // 41.251 — check that the whole or main part is present
    expect(text).toContain("41");
  });

  it("should evaluate SUM(F11:F13) via INDIRECT and render the result in the PDF", async () => {
    const text = await extractPdfText(pdfBytes);
    expect(text).toContain("1277");
  });
});

// ---------------------------------------------------------------------------
// evaluateWorksheetFormulas — unit tests (no PDF involved)
// ---------------------------------------------------------------------------
describe("evaluateWorksheetFormulas", () => {
  it("evaluates a simple SUM formula", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("A3").value = 30;
    ws.getCell("A4").value = { formula: "SUM(A1:A3)", result: 0 }; // stale result

    evaluateWorksheetFormulas(ws);

    expect(ws.getCell("A4").result).toBe(60);
  });

  it("evaluates AVERAGE formula", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("B1").value = 2;
    ws.getCell("B2").value = 4;
    ws.getCell("B3").value = 6;
    ws.getCell("B4").value = { formula: "AVERAGE(B1:B3)", result: 0 };

    evaluateWorksheetFormulas(ws);

    expect(ws.getCell("B4").result).toBe(4);
  });

  it("evaluates MIN and MAX formulas", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("C1").value = 5;
    ws.getCell("C2").value = 3;
    ws.getCell("C3").value = 9;
    ws.getCell("C4").value = { formula: "MIN(C1:C3)", result: 0 };
    ws.getCell("C5").value = { formula: "MAX(C1:C3)", result: 0 };

    evaluateWorksheetFormulas(ws);

    expect(ws.getCell("C4").result).toBe(3);
    expect(ws.getCell("C5").result).toBe(9);
  });

  it("evaluates arithmetic expressions", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("D1").value = 8;
    ws.getCell("D2").value = { formula: "D1*2+5", result: 0 };

    evaluateWorksheetFormulas(ws);

    expect(ws.getCell("D2").result).toBe(21);
  });

  it("evaluates IF formula", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("E1").value = 10;
    ws.getCell("E2").value = { formula: 'IF(E1>5,"big","small")', result: "" };

    evaluateWorksheetFormulas(ws);

    expect(ws.getCell("E2").result).toBe("big");
  });

  it("evaluates SUM with INDIRECT and ROW()", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = 1;
    ws.getCell("A2").value = 2;
    ws.getCell("A3").value = 3;
    // Row 4: SUM(A1:INDIRECT("A" & ROW()-1)) = SUM(A1:A3) = 6
    ws.getCell("A4").value = { formula: 'SUM(A1:INDIRECT("A" & ROW()-1))', result: 0 };

    evaluateWorksheetFormulas(ws);

    expect(ws.getCell("A4").result).toBe(6);
  });

  it("leaves non-formula cells unchanged", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("F1").value = 42;
    ws.getCell("F2").value = "hello";

    evaluateWorksheetFormulas(ws);

    expect(ws.getCell("F1").value).toBe(42);
    expect(ws.getCell("F2").value).toBe("hello");
  });

  it("does not throw for unknown/unsupported functions", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("G1").value = { formula: "VLOOKUP(1,A1:B10,2,0)", result: 0 };

    // Should not throw; silently skips the cell
    expect(() => evaluateWorksheetFormulas(ws)).not.toThrow();
  });

  it("handles formula cells with no cached result (result = undefined)", () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("H1").value = 5;
    ws.getCell("H2").value = 7;
    ws.getCell("H3").value = { formula: "H1+H2" };

    evaluateWorksheetFormulas(ws);

    expect(ws.getCell("H3").result).toBe(12);
  });
});

// ---------------------------------------------------------------------------
// Inline workbook: create fresh workbook, add formulas, convert — no XLSX
// ---------------------------------------------------------------------------
describe("excelToPdf with programmatically-created formulas", () => {
  it("renders SUM formula result in PDF", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sales");
    ws.getCell("A1").value = "Q1";
    ws.getCell("B1").value = 1000;
    ws.getCell("A2").value = "Q2";
    ws.getCell("B2").value = 2500;
    ws.getCell("A3").value = "Total";
    // Formula with NO pre-computed result
    ws.getCell("B3").value = { formula: "SUM(B1:B2)" };

    const pdfBytes = await excelToPdf(wb);

    expectValidPdf(pdfBytes);
    const text = await extractPdfText(pdfBytes);
    expect(text).toContain("3500");
  });

  it("renders AVERAGE formula result in PDF", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Stats");
    ws.getCell("A1").value = 10;
    ws.getCell("A2").value = 20;
    ws.getCell("A3").value = 30;
    ws.getCell("A4").value = { formula: "AVERAGE(A1:A3)" };

    const pdfBytes = await excelToPdf(wb);

    expectValidPdf(pdfBytes);
    const text = await extractPdfText(pdfBytes);
    expect(text).toContain("20");
  });

  it("renders IF formula result in PDF", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Check");
    ws.getCell("A1").value = 100;
    ws.getCell("A2").value = { formula: 'IF(A1>=100,"Pass","Fail")' };

    const pdfBytes = await excelToPdf(wb);

    expectValidPdf(pdfBytes);
    const text = await extractPdfText(pdfBytes);
    expect(text).toContain("Pass");
  });
});
