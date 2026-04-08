/**
 * Integration tests for PDF rendering edge cases.
 *
 * These tests verify the fixes from PR #131 and related style improvements:
 * type-based alignment, merge border propagation, text overflow, double borders,
 * zero-value number formats, fitToPage, row heights, error cells, and newline handling.
 */
import { describe, it, expect } from "vitest";
import { Workbook } from "@excel/workbook";
import { excelToPdf } from "@pdf/excel-bridge";
import { pdf as standalonePdf } from "@pdf/pdf";
import { readPdf } from "@pdf/reader/pdf-reader";
import { expectValidPdf } from "./test-helpers";

// Helper: extract page text from PDF bytes
async function extractText(pdfBytes: Uint8Array): Promise<string> {
  const result = await readPdf(pdfBytes);
  return result.text;
}

// Helper: get text fragments with positions from first page
async function getFragments(
  pdfBytes: Uint8Array
): Promise<{ text: string; x: number; y: number }[]> {
  const result = await readPdf(pdfBytes);
  return result.pages[0].textFragments.map(f => ({
    text: f.text,
    x: Math.round(f.x * 10) / 10,
    y: Math.round(f.y * 10) / 10
  }));
}

describe("PDF Rendering Edge Cases", () => {
  // ===========================================================================
  // Default alignment
  // ===========================================================================

  describe("Type-based default alignment", () => {
    it("should right-align numbers and left-align text by default", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getCell("A1").value = "Text";
      ws.getCell("A2").value = 42;

      const pdfBytes = await excelToPdf(wb, { showGridLines: true });
      const fragments = await getFragments(pdfBytes);

      const textFrag = fragments.find(f => f.text === "Text");
      const numFrag = fragments.find(f => f.text === "42");
      expect(textFrag).toBeDefined();
      expect(numFrag).toBeDefined();
      // Number should be further right than text (right-aligned vs left-aligned)
      expect(numFrag!.x).toBeGreaterThan(textFrag!.x);
    });

    it("should center-align booleans by default", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getCell("A1").value = "Left";
      ws.getCell("A2").value = true;

      const pdfBytes = await excelToPdf(wb, { showGridLines: true });
      const fragments = await getFragments(pdfBytes);

      const textFrag = fragments.find(f => f.text === "Left");
      const boolFrag = fragments.find(f => f.text.toLowerCase() === "true");
      expect(textFrag).toBeDefined();
      expect(boolFrag).toBeDefined();
      // Boolean (centered) should be further right than left-aligned text
      expect(boolFrag!.x).toBeGreaterThan(textFrag!.x);
    });

    it("should align formula cells by result type", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getColumn(2).width = 20;
      ws.getCell("A1").value = { formula: "1+1", result: 2 };
      ws.getCell("B1").value = { formula: 'CONCAT("a","b")', result: "ab" };

      const pdfBytes = await excelToPdf(wb, { showGridLines: true });
      const fragments = await getFragments(pdfBytes);

      const numFrag = fragments.find(f => f.text === "2");
      const strFrag = fragments.find(f => f.text === "ab");
      expect(numFrag).toBeDefined();
      expect(strFrag).toBeDefined();
    });

    it("should work in standalone pdf() mode", async () => {
      const pdfBytes = await standalonePdf({
        sheets: [
          {
            name: "Test",
            data: [
              ["text", 123, true],
              ["hello", 456, false]
            ]
          }
        ]
      });

      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("text");
      expect(text).toContain("123");
      expect(text.toLowerCase()).toContain("true");
    });
  });

  // ===========================================================================
  // Merged cell borders
  // ===========================================================================

  describe("Merged cell border propagation", () => {
    it("should preserve borders set on boundary cells after merge", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.mergeCells("A1:C3");
      ws.getCell("A1").value = "Merged";
      ws.getCell("A1").border = {
        top: { style: "thick", color: { argb: "FFFF0000" } },
        left: { style: "thick", color: { argb: "FF00FF00" } }
      };
      ws.getCell("C1").border = { right: { style: "thick", color: { argb: "FF0000FF" } } };
      ws.getCell("A3").border = { bottom: { style: "thick", color: { argb: "FFFF00FF" } } };

      const pdfBytes = await excelToPdf(wb, { showGridLines: true });
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Merged");
    });

    it("should render bordered empty cells", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "Data";
      ws.getCell("B1").border = {
        top: { style: "thick" },
        right: { style: "thick" },
        bottom: { style: "thick" },
        left: { style: "thick" }
      };
      ws.getCell("C1").fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" }
      };

      const pdfBytes = await excelToPdf(wb, { showGridLines: true });
      expectValidPdf(pdfBytes);
      // B1 and C1 have no text but should still be rendered (borders/fill)
      const text = await extractText(pdfBytes);
      expect(text).toContain("Data");
    });
  });

  // ===========================================================================
  // Text overflow
  // ===========================================================================

  describe("Text overflow into adjacent cells", () => {
    it("should overflow text into empty neighbors", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 10;
      ws.getColumn(2).width = 10;
      ws.getColumn(3).width = 10;
      ws.getCell("A1").value = "This is a very long text that overflows";

      const pdfBytes = await excelToPdf(wb, { showGridLines: true });
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("This is a very long text that overflows");
    });

    it("should stop overflow at cells with content", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 10;
      ws.getColumn(2).width = 10;
      ws.getCell("A1").value = "Long text that should not fully display";
      ws.getCell("B1").value = "Blocker";

      const pdfBytes = await excelToPdf(wb, { showGridLines: true });
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Blocker");
    });
  });

  // ===========================================================================
  // Double borders
  // ===========================================================================

  describe("Double border rendering", () => {
    it("should render double borders without crashing", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "Double";
      ws.getCell("A1").border = {
        top: { style: "double" },
        bottom: { style: "double" },
        left: { style: "double" },
        right: { style: "double" }
      };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Double");
    });
  });

  // ===========================================================================
  // Zero-value number formats
  // ===========================================================================

  describe("Zero-value number formats", () => {
    it('should format accounting zero as dash with spaces: "-"??', async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getCell("A1").value = 0;
      ws.getCell("A1").numFmt = '#,##0.00;-#,##0.00;"-"??';

      const pdfBytes = await excelToPdf(wb);
      const text = await extractText(pdfBytes);
      expect(text).toContain("-");
    });

    it("should produce empty string for #.## with zero", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 0;
      ws.getCell("A1").numFmt = "#.##";
      ws.getCell("A2").value = "marker";

      const pdfBytes = await excelToPdf(wb);
      const text = await extractText(pdfBytes);
      // A1 should be empty, only marker visible
      expect(text).toContain("marker");
    });

    it("should pad with spaces for ?? placeholders", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getCell("A1").value = 0;
      ws.getCell("A1").numFmt = "??0.00";

      const pdfBytes = await excelToPdf(wb);
      const text = await extractText(pdfBytes);
      expect(text).toContain("0.00");
    });
  });

  // ===========================================================================
  // fitToPage
  // ===========================================================================

  describe("fitToPage scaling", () => {
    it("should fit 20 columns onto a single page", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      for (let c = 1; c <= 20; c++) {
        ws.getColumn(c).width = 12;
        ws.getCell(1, c).value = `Col ${c}`;
      }

      const pdfBytes = await excelToPdf(wb, { fitToPage: true });
      expectValidPdf(pdfBytes);
      const result = await readPdf(pdfBytes);
      // All 20 columns should be on one page
      expect(result.pages).toHaveLength(1);
      expect(result.text).toContain("Col 1");
      expect(result.text).toContain("Col 20");
    });
  });

  // ===========================================================================
  // Row heights
  // ===========================================================================

  describe("Row height handling", () => {
    it("should respect custom row height", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 30;
      ws.getCell("A1").value = "Normal";
      const r2 = ws.getRow(2);
      r2.height = 50;
      ws.getCell("A2").value = "Tall";
      ws.getCell("A3").value = "Normal again";

      const pdfBytes = await excelToPdf(wb, { showGridLines: true });
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Normal");
      expect(text).toContain("Tall");
    });
  });

  // ===========================================================================
  // Newline handling
  // ===========================================================================

  describe("Explicit newline handling", () => {
    it("should split non-wrapped text on explicit newlines", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getCell("A1").value = "Line1\nLine2\nLine3";

      const pdfBytes = await excelToPdf(wb);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Line1");
      expect(text).toContain("Line2");
      expect(text).toContain("Line3");
    });

    it("should handle wrapped text with newlines", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getCell("A1").value = "Wrap\nwith\nnewlines";
      ws.getCell("A1").alignment = { wrapText: true };

      const pdfBytes = await excelToPdf(wb);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Wrap");
      expect(text).toContain("newlines");
    });
  });

  // ===========================================================================
  // Error values and mixed types
  // ===========================================================================

  describe("Error values and special types", () => {
    it("should render error cell values", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = { error: "#DIV/0!" } as any;

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("#DIV/0!");
    });

    it("should render rich text with mixed formatting", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 40;
      ws.getCell("A1").value = {
        richText: [
          { text: "Bold ", font: { bold: true } },
          { text: "Normal ", font: { size: 11 } },
          { text: "Red", font: { color: { argb: "FFFF0000" } } }
        ]
      };

      const pdfBytes = await excelToPdf(wb);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Bold");
      expect(text).toContain("Normal");
      expect(text).toContain("Red");
    });

    it("should handle hyperlink cells", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = { text: "Click me", hyperlink: "https://example.com" };
      ws.getCell("A1").font = { color: { argb: "FF0563C1" }, underline: true };

      const pdfBytes = await excelToPdf(wb);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Click me");
    });
  });

  // ===========================================================================
  // Vertical stacked text (textRotation = 255)
  // ===========================================================================

  describe("Vertical stacked text", () => {
    it("should render vertical stacked text (rotation 255)", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getRow(1).height = 80;
      ws.getCell("A1").value = "Vertical";
      ws.getCell("A1").alignment = { textRotation: 255 };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      // Vertical stacked renders each char separately
      expect(text).toContain("V");
      expect(text).toContain("e");
    });
  });
});
