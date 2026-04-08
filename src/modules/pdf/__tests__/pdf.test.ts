/**
 * Tests for PDF APIs: pdf, excelToPdf, readPdf.
 *
 * Verifies correct results, event loop yielding between pages,
 * and roundtrip (write → read) integrity.
 */
import { describe, it, expect, beforeAll } from "vitest";
import { Workbook } from "@excel/workbook";
import { pdf } from "@pdf/pdf";
import { excelToPdf } from "@pdf/excel-bridge";
import { readPdf } from "@pdf/reader/pdf-reader";
import { PdfError, PdfStructureError } from "@pdf/errors";
import { pdfToString, expectValidPdf } from "./test-helpers";

// =============================================================================
// readPdf
// =============================================================================

describe("readPdf", () => {
  let pdfBytes: Uint8Array;

  beforeAll(async () => {
    pdfBytes = await pdf(
      [
        ["Name", "Score"],
        ["Alice", 95],
        ["Bob", 87]
      ],
      { title: "Async Test", author: "Test Author", showGridLines: true }
    );
  });

  it("should support options (extractImages: false)", async () => {
    const result = await readPdf(pdfBytes, { extractImages: false });
    expect(result.pages.length).toBeGreaterThan(0);
    for (const page of result.pages) {
      expect(page.images).toEqual([]);
    }
  });

  it("should support page selection", async () => {
    const result = await readPdf(pdfBytes, { pages: [1] });
    expect(result.pages.length).toBe(1);
    expect(result.pages[0].pageNumber).toBe(1);
  });

  it("should throw on invalid PDF data", async () => {
    await expect(readPdf(new Uint8Array([1, 2, 3]))).rejects.toThrow();
  });

  it("should yield to event loop between pages", async () => {
    // Generate a multi-page PDF
    const rows: [string, number][] = [];
    for (let i = 0; i < 200; i++) {
      rows.push([`Row ${i}`, i]);
    }
    const bigPdf = await pdf([["Name", "Value"], ...rows]);

    let yielded = false;
    const timer = setTimeout(() => {
      yielded = true;
    }, 0);

    await readPdf(bigPdf);
    clearTimeout(timer);

    // readPdf should have allowed setTimeout(0) callbacks to run
    expect(yielded).toBe(true);
  });
});

// =============================================================================
// pdf
// =============================================================================

describe("pdf", () => {
  it("should produce a valid PDF", async () => {
    const input: [string, number][] = [
      ["Product", 100],
      ["Widget", 200]
    ];
    const asyncBytes = await pdf([["Name", "Value"], ...input]);

    expectValidPdf(asyncBytes);
    expect(asyncBytes.length).toBeGreaterThan(100);

    // Verify text when read back
    const asyncRead = await readPdf(asyncBytes);
    expect(asyncRead.text).toContain("Name");
    expect(asyncRead.text).toContain("Value");
    expect(asyncRead.pages.length).toBeGreaterThan(0);
  });

  it("should accept sheet object input", async () => {
    const result = await pdf({
      name: "Test",
      data: [
        ["A", "B"],
        [1, 2]
      ]
    });
    expectValidPdf(result);
  });

  it("should accept workbook object input", async () => {
    const result = await pdf({
      sheets: [
        { name: "Sheet1", data: [["A"], [1]] },
        { name: "Sheet2", data: [["B"], [2]] }
      ]
    });
    expectValidPdf(result);
  });

  it("should respect options", async () => {
    const result = await pdf(
      [
        ["Name", "Score"],
        ["Alice", 95]
      ],
      { title: "My Title", showGridLines: true }
    );
    expectValidPdf(result);
    const text = pdfToString(result);
    expect(text).toContain("My Title");
  });

  it("should throw for empty workbook", async () => {
    await expect(pdf({ sheets: [] })).rejects.toThrow(PdfError);
  });
});

// =============================================================================
// excelToPdf
// =============================================================================

describe("excelToPdf", () => {
  it("should produce a valid PDF", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sales");
    ws.getCell("A1").value = "Product";
    ws.getCell("B1").value = "Revenue";
    ws.getCell("A2").value = "Widget";
    ws.getCell("B2").value = 1000;

    const asyncBytes = await excelToPdf(wb);

    expectValidPdf(asyncBytes);

    // Verify readable text
    const asyncRead = await readPdf(asyncBytes);
    expect(asyncRead.text).toContain("Product");
    expect(asyncRead.pages.length).toBeGreaterThan(0);
  });

  it("should export multiple worksheets", async () => {
    const wb = new Workbook();
    const ws1 = wb.addWorksheet("Sheet1");
    ws1.getCell("A1").value = "Hello";
    const ws2 = wb.addWorksheet("Sheet2");
    ws2.getCell("A1").value = "World";

    const result = await excelToPdf(wb);
    expectValidPdf(result);

    const read = await readPdf(result);
    expect(read.pages.length).toBeGreaterThanOrEqual(2);
  });

  it("should throw for workbook with no sheets", async () => {
    const wb = new Workbook();
    await expect(excelToPdf(wb)).rejects.toThrow(PdfError);
  });

  it("should support PDF export options", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Sheet1");
    ws.getCell("A1").value = "Test";

    const result = await excelToPdf(wb, {
      title: "Async Export",
      showPageNumbers: true
    });
    expectValidPdf(result);
    const text = pdfToString(result);
    expect(text).toContain("Async Export");
  });
});

// =============================================================================
// Roundtrip: pdf → readPdf
// =============================================================================

describe("Roundtrip (pdf → readPdf)", () => {
  it("should write and read back correctly", async () => {
    const bytes = await pdf([
      ["Name", "Score"],
      ["Alice", 95],
      ["Bob", 87]
    ]);

    const result = await readPdf(bytes);
    expect(result.pages.length).toBeGreaterThan(0);
    expect(result.text).toContain("Alice");
    expect(result.text).toContain("95");
    expect(result.text).toContain("Bob");
    expect(result.text).toContain("87");
  });

  it("should handle encrypted roundtrip", async () => {
    const bytes = await pdf(
      [
        ["Secret", "Data"],
        ["A", 1]
      ],
      { encryption: { ownerPassword: "owner123", userPassword: "user456" } }
    );

    // Without password should throw
    await expect(readPdf(bytes)).rejects.toThrow(PdfStructureError);

    // With wrong password should throw
    await expect(readPdf(bytes, { password: "wrong" })).rejects.toThrow(PdfStructureError);

    // With correct user password should work
    const result = await readPdf(bytes, { password: "user456" });
    expect(result.text).toContain("Secret");

    // With correct owner password should also work
    const result2 = await readPdf(bytes, { password: "owner123" });
    expect(result2.text).toContain("Secret");
  });
});
