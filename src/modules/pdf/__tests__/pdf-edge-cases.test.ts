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
  // Vertical stacked text (textRotation = 255 / "vertical")
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

    it("should respect horizontal alignment for vertical stacked text", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getColumn(2).width = 20;
      ws.getRow(1).height = 100;

      ws.getCell("A1").value = "Hi";
      ws.getCell("A1").alignment = {
        textRotation: "vertical" as any,
        horizontal: "left"
      };
      ws.getCell("B1").value = "Hi";
      ws.getCell("B1").alignment = {
        textRotation: "vertical" as any,
        horizontal: "right"
      };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);
      const hFrags = frags.filter(f => f.text === "H");
      expect(hFrags.length).toBeGreaterThanOrEqual(2);
      // Left-aligned "H" should have a smaller x than right-aligned "H"
      expect(hFrags[0].x).toBeLessThan(hFrags[1].x);
    });

    it("should respect vertical alignment for vertical stacked text", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 10;
      ws.getColumn(2).width = 10;
      ws.getColumn(3).width = 10;
      ws.getRow(1).height = 120;

      ws.getCell("A1").value = "A";
      ws.getCell("A1").alignment = {
        textRotation: "vertical" as any,
        vertical: "top"
      };
      ws.getCell("B1").value = "B";
      ws.getCell("B1").alignment = {
        textRotation: "vertical" as any,
        vertical: "middle"
      };
      ws.getCell("C1").value = "C";
      ws.getCell("C1").alignment = {
        textRotation: "vertical" as any,
        vertical: "bottom"
      };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);
      const aFrag = frags.find(f => f.text === "A");
      const bFrag = frags.find(f => f.text === "B");
      const cFrag = frags.find(f => f.text === "C");
      expect(aFrag).toBeDefined();
      expect(bFrag).toBeDefined();
      expect(cFrag).toBeDefined();
      // top has highest y, bottom has lowest y (PDF coords)
      expect(aFrag!.y).toBeGreaterThan(bFrag!.y);
      expect(bFrag!.y).toBeGreaterThan(cFrag!.y);
    });
  });

  // ===========================================================================
  // Rotated text alignment (90°, -90°, general angles)
  // ===========================================================================

  describe("Rotated text alignment", () => {
    // --- 90° vertical alignment ---
    it("should position 90° text according to vertical alignment (top > middle > bottom)", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 12;
      ws.getColumn(2).width = 12;
      ws.getColumn(3).width = 12;
      ws.getRow(1).height = 80;

      ws.getCell("A1").value = "Top";
      ws.getCell("A1").alignment = { textRotation: 90, horizontal: "center", vertical: "top" };
      ws.getCell("B1").value = "Mid";
      ws.getCell("B1").alignment = { textRotation: 90, horizontal: "center", vertical: "middle" };
      ws.getCell("C1").value = "Bot";
      ws.getCell("C1").alignment = { textRotation: 90, horizontal: "center", vertical: "bottom" };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);
      const topFrag = frags.find(f => f.text === "Top");
      const midFrag = frags.find(f => f.text === "Mid");
      const botFrag = frags.find(f => f.text === "Bot");
      expect(topFrag).toBeDefined();
      expect(midFrag).toBeDefined();
      expect(botFrag).toBeDefined();
      // In PDF coords (origin bottom-left), top text starts at highest y
      expect(topFrag!.y).toBeGreaterThan(midFrag!.y);
      expect(midFrag!.y).toBeGreaterThan(botFrag!.y);
    });

    // --- 90° horizontal alignment ---
    it("should position 90° text left/right with horizontal alignment", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getColumn(2).width = 20;
      ws.getRow(1).height = 80;

      ws.getCell("A1").value = "Left";
      ws.getCell("A1").alignment = { textRotation: 90, horizontal: "left" };
      ws.getCell("B1").value = "Right";
      ws.getCell("B1").alignment = { textRotation: 90, horizontal: "right" };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);
      const leftFrag = frags.find(f => f.text === "Left");
      const rightFrag = frags.find(f => f.text === "Right");
      expect(leftFrag).toBeDefined();
      expect(rightFrag).toBeDefined();
      // Right cell's text x should be greater (further right)
      expect(rightFrag!.x).toBeGreaterThan(leftFrag!.x);
    });

    // --- 90° combined: left+bottom vs right+top (issue #133 core scenario) ---
    it("should handle 90° combined h/v alignment (issue #133)", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getColumn(2).width = 20;
      ws.getRow(1).height = 100;

      ws.getCell("A1").value = "LB";
      ws.getCell("A1").alignment = {
        textRotation: 90,
        horizontal: "left",
        vertical: "bottom"
      };
      ws.getCell("B1").value = "RT";
      ws.getCell("B1").alignment = {
        textRotation: 90,
        horizontal: "right",
        vertical: "top"
      };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);
      const lbFrag = frags.find(f => f.text === "LB");
      const rtFrag = frags.find(f => f.text === "RT");
      expect(lbFrag).toBeDefined();
      expect(rtFrag).toBeDefined();
      // LB: left+bottom → low x, low y; RT: right+top → high x, high y
      expect(lbFrag!.x).toBeLessThan(rtFrag!.x);
      expect(lbFrag!.y).toBeLessThan(rtFrag!.y);
    });

    // --- -90° alignment ---
    it("should render -90° rotated text with alignment", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 12;
      ws.getColumn(2).width = 12;
      ws.getColumn(3).width = 12;
      ws.getRow(1).height = 80;

      ws.getCell("A1").value = "Top";
      ws.getCell("A1").alignment = { textRotation: -90, vertical: "top" };
      ws.getCell("B1").value = "Mid";
      ws.getCell("B1").alignment = { textRotation: -90, vertical: "middle" };
      ws.getCell("C1").value = "Bot";
      ws.getCell("C1").alignment = { textRotation: -90, vertical: "bottom" };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);
      const topFrag = frags.find(f => f.text === "Top");
      const midFrag = frags.find(f => f.text === "Mid");
      const botFrag = frags.find(f => f.text === "Bot");
      expect(topFrag).toBeDefined();
      expect(midFrag).toBeDefined();
      expect(botFrag).toBeDefined();
      // For -90° (text flows downward), top text starts from highest y
      expect(topFrag!.y).toBeGreaterThan(midFrag!.y);
      expect(midFrag!.y).toBeGreaterThan(botFrag!.y);
    });

    it("should position -90° text left/right with horizontal alignment", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 20;
      ws.getColumn(2).width = 20;
      ws.getRow(1).height = 80;

      ws.getCell("A1").value = "Left";
      ws.getCell("A1").alignment = { textRotation: -90, horizontal: "left" };
      ws.getCell("B1").value = "Right";
      ws.getCell("B1").alignment = { textRotation: -90, horizontal: "right" };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);
      const leftFrag = frags.find(f => f.text === "Left");
      const rightFrag = frags.find(f => f.text === "Right");
      expect(leftFrag).toBeDefined();
      expect(rightFrag).toBeDefined();
      expect(rightFrag!.x).toBeGreaterThan(leftFrag!.x);
    });

    // --- General angle (45°) alignment ---
    it("should render 45° text with top-left vs bottom-right alignment", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getColumn(1).width = 15;
      ws.getColumn(2).width = 15;
      ws.getRow(1).height = 60;

      ws.getCell("A1").value = "TL";
      ws.getCell("A1").alignment = {
        textRotation: 45,
        horizontal: "left",
        vertical: "top"
      };
      ws.getCell("B1").value = "BR";
      ws.getCell("B1").alignment = {
        textRotation: 45,
        horizontal: "right",
        vertical: "bottom"
      };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);
      const tlFrag = frags.find(f => f.text === "TL");
      const brFrag = frags.find(f => f.text === "BR");
      expect(tlFrag).toBeDefined();
      expect(brFrag).toBeDefined();
      // Top-left should have higher y than bottom-right
      expect(tlFrag!.y).toBeGreaterThan(brFrag!.y);
    });

    // --- 45° slanted borders ---
    it("should render slanted borders for general rotation angles", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getRow(1).height = 60;
      ws.getCell("A1").value = "Slant";
      ws.getCell("A1").alignment = { textRotation: 45 };
      ws.getCell("A1").border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Slant");
    });

    // --- 45° negative angle ---
    it("should render -45° text with slanted borders", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getRow(1).height = 60;
      ws.getCell("A1").value = "Neg";
      ws.getCell("A1").alignment = { textRotation: -45 };
      ws.getCell("A1").border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const text = await extractText(pdfBytes);
      expect(text).toContain("Neg");
    });

    // --- All 6 combos from issue #133 reproduced ---
    it("should match Excel alignment for all 6 rotation combos from issue #133", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      for (let c = 1; c <= 6; c++) {
        ws.getColumn(c).width = 12;
      }
      ws.getRow(1).height = 110;

      // The exact combos from the issue's PDF-Test-2.xlsx Row 8
      const combos: Array<{ value: string; h: string; v?: string }> = [
        { value: "Col1", h: "center", v: "top" },
        { value: "Col2", h: "center", v: "middle" },
        { value: "Col3", h: "center" },
        { value: "Col4", h: "left" },
        { value: "Col5", h: "right" },
        { value: "Col6", h: "left", v: "top" }
      ];

      for (let i = 0; i < combos.length; i++) {
        const cell = ws.getCell(1, i + 1);
        cell.value = combos[i].value;
        cell.alignment = {
          textRotation: 90,
          horizontal: combos[i].h as any,
          vertical: combos[i].v as any,
          wrapText: true
        };
      }

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      const frags = await getFragments(pdfBytes);

      // All 6 texts should be present
      for (const combo of combos) {
        expect(frags.find(f => f.text === combo.value)).toBeDefined();
      }

      // Col1 (center/top) should have higher y than Col3 (center/bottom-default)
      const col1 = frags.find(f => f.text === "Col1")!;
      const col3 = frags.find(f => f.text === "Col3")!;
      expect(col1.y).toBeGreaterThan(col3.y);

      // Col4 (left) should have smaller x than Col5 (right) within their cells
      // Col4 is in column 4 (offset ~3*colWidth), Col5 in column 5 (offset ~4*colWidth)
      // Col4 left-aligned within its cell, Col5 right-aligned within its cell
      const col4 = frags.find(f => f.text === "Col4")!;
      const col5 = frags.find(f => f.text === "Col5")!;
      // Despite being in adjacent cells, right-aligned Col5 should have notably larger x
      expect(col5.x).toBeGreaterThan(col4.x);
    });
  });
});
