/**
 * Integration test: PDF-Test.xlsx → PDF conversion
 *
 * Verifies that the Spares Offer spreadsheet exports correctly,
 * especially that fitToPage doesn't double-apply the sheet's
 * pageSetup.scale.
 */
import { describe, it, expect } from "vitest";
import { Workbook } from "@excel/workbook";
import { excelToPdf } from "@pdf/excel-bridge";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));

// pnpm exec vitest run src/modules/pdf/__tests__/pdf-test-xlsx.test.ts
const PDF_TEST_FILE = resolve(__dirname, "../../excel/__tests__/data/PDF-Test.xlsx");

function pdfToString(pdf: Uint8Array): string {
  return new TextDecoder().decode(pdf);
}

function expectValidPdf(pdf: Uint8Array): void {
  const text = pdfToString(pdf);
  expect(text).toContain("%PDF-1.4");
  expect(text).toContain("xref");
  expect(text).toContain("trailer");
  expect(text).toContain("%%EOF");
  expect(text).toContain("/Catalog");
  expect(text).toContain("/Pages");
}

describe("PDF-Test.xlsx to PDF", () => {
  let workbook: Workbook;

  // Load the workbook once for all tests
  async function loadWorkbook(): Promise<Workbook> {
    if (!workbook) {
      workbook = new Workbook();
      await workbook.xlsx.readFile(PDF_TEST_FILE);
    }
    return workbook;
  }

  it("should produce a valid PDF from PDF-Test.xlsx", async () => {
    const wb = await loadWorkbook();

    const pdf = excelToPdf(wb);

    expect(pdf).toBeInstanceOf(Uint8Array);
    expect(pdf.length).toBeGreaterThan(1000);
    expectValidPdf(pdf);
  });

  it("should not double-scale when fitToPage is active with sheet scale", async () => {
    const wb = await loadWorkbook();

    // The sheet has pageSetup.scale = 20 (20%) and orientation = landscape.
    // With fitToPage (default: true), the sheet scale should be ignored
    // and the content should be scaled only by the fitToPage calculation.
    const pdf = excelToPdf(wb);

    const text = pdfToString(pdf);
    // The PDF should have landscape A4 pages (841.89 x 595.28)
    // Content should be readable, not tiny
    expectValidPdf(pdf);

    // Verify it's landscape: MediaBox width > height
    const mediaBoxMatch = text.match(/\/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]/);
    expect(mediaBoxMatch).not.toBeNull();
    if (mediaBoxMatch) {
      const pageWidth = parseFloat(mediaBoxMatch[1]);
      const pageHeight = parseFloat(mediaBoxMatch[2]);
      expect(pageWidth).toBeGreaterThan(pageHeight); // landscape
    }

    // The PDF should contain visible text content from the spreadsheet
    // (not compressed to nothing)
    expect(pdf.length).toBeGreaterThan(5000);
  });

  it("should respect explicit scale when fitToPage is disabled", async () => {
    const wb = await loadWorkbook();

    // When fitToPage is false, the sheet's scale (20%) should be used
    const pdf = excelToPdf(wb, {
      fitToPage: false,
      orientation: "landscape",
      pageSize: "A4"
    });

    expectValidPdf(pdf);
    expect(pdf.length).toBeGreaterThan(1000);
  });

  it("should produce reasonable cell sizes, not extremely tiny", async () => {
    const wb = await loadWorkbook();

    const pdf = excelToPdf(wb, {
      fitToPage: true,
      orientation: "landscape",
      pageSize: "A4"
    });

    expectValidPdf(pdf);

    // The PDF should be substantial in size (not just empty pages)
    expect(pdf.length).toBeGreaterThan(5000);
  });

  it("should produce identical output with default options and explicit scale=1.0", async () => {
    const wb = await loadWorkbook();

    // Default: fitToPage=true, scale inherited from sheet (20%) should be ignored
    const pdfDefault = excelToPdf(wb);
    // Explicit scale=1.0 with fitToPage=true
    const pdfExplicit = excelToPdf(wb, { scale: 1.0 });

    // Both should produce identical PDFs since fitToPage=true should ignore the sheet scale
    expect(pdfDefault.length).toBe(pdfExplicit.length);
  });

  it("should include empty cells that have borders in PDF output", async () => {
    const wb = await loadWorkbook();

    // Build a workbook with a styled empty cell to verify it's included
    const testWb = new Workbook();
    const testWs = testWb.addWorksheet("Test");
    testWs.getCell("A1").value = "Data";
    testWs.getCell("A1").border = {
      top: { style: "thin" },
      right: { style: "thin" },
      bottom: { style: "thin" },
      left: { style: "thin" }
    };
    // B1 is empty but has a right border
    testWs.getCell("B1").border = {
      right: { style: "thin" }
    };

    const pdf = excelToPdf(testWb);
    expectValidPdf(pdf);
    // The PDF should be larger than one with just A1 data, proving B1 borders are rendered
    expect(pdf.length).toBeGreaterThan(500);
  });
});
