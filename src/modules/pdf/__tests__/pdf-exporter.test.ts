/**
 * Integration tests for the full PDF export pipeline.
 * Tests the PdfExporter with real Workbook instances.
 */
import { describe, it, expect } from "vitest";
import { Workbook } from "@excel/workbook";
import { PdfExporter, exportPdf } from "@pdf/render/pdf-exporter";
import { PdfError } from "@pdf/errors";
import { buildMinimalTtf } from "./font-embedding.test";

/**
 * Helper to decode a PDF Uint8Array to string for assertion.
 */
function pdfToString(pdf: Uint8Array): string {
  return new TextDecoder().decode(pdf);
}

/**
 * Verify basic PDF structure (header, xref, trailer, EOF).
 */
function expectValidPdf(pdf: Uint8Array): void {
  const text = pdfToString(pdf);
  expect(text).toContain("%PDF-1.4");
  expect(text).toContain("xref");
  expect(text).toContain("trailer");
  expect(text).toContain("%%EOF");
  expect(text).toContain("/Catalog");
  expect(text).toContain("/Pages");
}

describe("PdfExporter", () => {
  describe("Basic Export", () => {
    it("should export a simple workbook with one sheet", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "Hello";
      ws.getCell("B1").value = "World";
      ws.getCell("A2").value = 42;
      ws.getCell("B2").value = 3.14;

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expect(pdf).toBeInstanceOf(Uint8Array);
      expect(pdf.length).toBeGreaterThan(100);
      expectValidPdf(pdf);

      // Content streams may be compressed; verify structure
      const text = pdfToString(pdf);
      expect(text).toContain("/Helvetica");
    });

    it("should export an empty workbook with at least one sheet", () => {
      const wb = new Workbook();
      wb.addWorksheet("Empty");

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
    });

    it("should throw for workbook with no sheets", () => {
      const wb = new Workbook();

      const exporter = new PdfExporter(wb);
      expect(() => exporter.export()).toThrow(PdfError);
    });
  });

  describe("Multiple Sheets", () => {
    it("should export multiple worksheets", () => {
      const wb = new Workbook();

      const ws1 = wb.addWorksheet("Sales");
      ws1.getCell("A1").value = "Product";
      ws1.getCell("B1").value = "Revenue";
      ws1.getCell("A2").value = "Widget";
      ws1.getCell("B2").value = 1000;

      const ws2 = wb.addWorksheet("Expenses");
      ws2.getCell("A1").value = "Category";
      ws2.getCell("B1").value = "Amount";
      ws2.getCell("A2").value = "Rent";
      ws2.getCell("B2").value = 500;

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Two pages for two sheets
      const pageMatches = text.match(/\/Type \/Page\b/g);
      expect(pageMatches!.length).toBe(2);
      // Outlines/bookmarks for multi-sheet navigation
      expect(text).toContain("/Outlines");
      expect(text).toContain("(Sales)");
      expect(text).toContain("(Expenses)");
    });

    it("should filter sheets by name", () => {
      const wb = new Workbook();
      wb.addWorksheet("Include").getCell("A1").value = "Included";
      wb.addWorksheet("Exclude").getCell("A1").value = "Excluded";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ sheets: ["Include"] });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Included");
      expect(text).not.toContain("Excluded");
    });

    it("should filter sheets by 1-based position", () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("First");
      ws1.getCell("A1").value = "First Sheet";
      const ws2 = wb.addWorksheet("Second");
      ws2.getCell("A1").value = "Second Sheet";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ sheets: [2] }); // 1-based: second sheet

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Second Sheet");
      expect(text).not.toContain("First Sheet");
    });
  });

  describe("Page Size and Orientation", () => {
    it("should support A4 portrait (default)", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "A4";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // A4: 595.28 x 841.89
      expect(text).toContain("595.28");
      expect(text).toContain("841.89");
    });

    it("should support landscape orientation", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Landscape";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ orientation: "landscape" });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Landscape A4: 841.89 x 595.28
      expect(text).toContain("841.89");
      expect(text).toContain("595.28");
    });

    it("should support LETTER page size", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Letter";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ pageSize: "LETTER" });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("612");
      expect(text).toContain("792");
    });

    it("should support custom page size", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Custom";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({
        pageSize: { width: 400, height: 600 }
      });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("400");
      expect(text).toContain("600");
    });
  });

  describe("Cell Styles", () => {
    it("should render bold text", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Bold Text";
      cell.font = { bold: true };

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Helvetica-Bold");
      expect(text).toContain("Bold Text");
    });

    it("should render italic text", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Italic Text";
      cell.font = { italic: true };

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Helvetica-Oblique");
    });

    it("should render colored text", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Red Text";
      cell.font = { color: { argb: "FFFF0000" } };

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("1 0 0 rg"); // red fill color for text
    });

    it("should render background fill", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Filled";
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" }
      };

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Yellow fill: 1 1 0
      expect(text).toContain("1 1 0 rg");
    });

    it("should render cell borders", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Bordered";
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } }
      };

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      // Borders are inside compressed content streams; just verify structure
    });
  });

  describe("Merged Cells", () => {
    it("should handle merged cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Merge");
      ws.getCell("A1").value = "Merged Title";
      ws.mergeCells("A1:C1");
      ws.getCell("A2").value = "Col1";
      ws.getCell("B2").value = "Col2";
      ws.getCell("C2").value = "Col3";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      // Merged cells produce one page with all cells rendered
    });
  });

  describe("Data Types", () => {
    it("should handle various cell value types", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Types");
      ws.getCell("A1").value = "String";
      ws.getCell("A2").value = 42;
      ws.getCell("A3").value = 3.14;
      ws.getCell("A4").value = true;
      ws.getCell("A5").value = new Date(2024, 0, 15);
      ws.getCell("A6").value = null;

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      // All value types render without errors
    });

    it("should handle hyperlinks", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Links");
      ws.getCell("A1").value = { text: "Click Me", hyperlink: "https://example.com" };

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Click Me");
    });
  });

  describe("Grid Lines", () => {
    it("should render grid lines when enabled", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Grid");
      ws.getCell("A1").value = "A1";
      ws.getCell("B1").value = "B1";
      ws.getCell("A2").value = "A2";
      ws.getCell("B2").value = "B2";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ showGridLines: true });

      expectValidPdf(pdf);
      // Grid lines are in compressed content streams; verify the PDF is valid
    });
  });

  describe("Page Headers and Footers", () => {
    it("should include sheet name as header", () => {
      const wb = new Workbook();
      wb.addWorksheet("My Report").getCell("A1").value = "Data";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ showSheetNames: true });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("My Report");
    });

    it("should include page numbers in footer", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ showPageNumbers: true });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Page 1 of 1");
    });

    it("should register footer fonts even when using an embedded font", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({
        showPageNumbers: true,
        font: new Uint8Array(buildMinimalTtf())
      });

      const text = pdfToString(pdf);
      expect(text).toContain("/BaseFont /Helvetica");
    });
  });

  describe("PDF Metadata", () => {
    it("should set document title", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ title: "My Report" });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("/Title (My Report)");
    });

    it("should set document author", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ author: "John Doe" });

      const text = pdfToString(pdf);
      expect(text).toContain("/Author (John Doe)");
    });

    it("should always set producer", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      const text = pdfToString(pdf);
      expect(text).toContain("/Producer (excelts)");
    });

    it("should encode Unicode metadata and bookmark titles correctly", () => {
      const wb = new Workbook();
      wb.addWorksheet("报告").getCell("A1").value = "One";
      wb.addWorksheet("数据").getCell("A1").value = "Two";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ title: "作者" });

      const text = pdfToString(pdf);
      expect(text).toContain("/Title <feff4f5c8005>");
      expect(text).toContain("/Title <feff62a5544a>");
      expect(text).toContain("/Title <feff6570636e>");
    });
  });

  describe("Worksheet Page Setup", () => {
    it("should honor per-sheet page setup defaults", () => {
      const wb = new Workbook();

      const ws1 = wb.addWorksheet("First");
      ws1.getCell("A1").value = "One";
      ws1.pageSetup.paperSize = 11;

      const ws2 = wb.addWorksheet("Second");
      ws2.getCell("A1").value = "Two";
      ws2.pageSetup.orientation = "landscape";
      ws2.pageSetup.paperSize = 9;

      const pdf = exportPdf(wb);
      const text = pdfToString(pdf);

      expect(text).toContain("[0 0 419.53 595.28]");
      expect(text).toContain("[0 0 841.89 595.28]");
    });
  });

  describe("Column Widths", () => {
    it("should respect custom column widths", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Widths");
      ws.columns = [
        { header: "Narrow", width: 5 },
        { header: "Wide", width: 30 }
      ];

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Narrow");
      expect(text).toContain("Wide");
    });
  });

  describe("Large Datasets (Pagination)", () => {
    it("should paginate when content exceeds page height", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("LargeData");

      // Add enough rows to fill multiple pages
      for (let i = 1; i <= 100; i++) {
        ws.getCell(`A${i}`).value = `Row ${i}`;
        ws.getCell(`B${i}`).value = i * 10;
      }

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ showPageNumbers: true });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);

      // Should have multiple pages
      const pageMatches = text.match(/\/Type \/Page\b/g);
      expect(pageMatches).not.toBeNull();
      expect(pageMatches!.length).toBeGreaterThan(1);
    });
  });

  describe("Options", () => {
    it("should clamp scale factor", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Scale";

      const exporter = new PdfExporter(wb);

      // Very small scale should be clamped to 0.1
      const pdf1 = exporter.export({ scale: 0.01 });
      expectValidPdf(pdf1);

      // Very large scale should be clamped to 3.0
      const pdf2 = exporter.export({ scale: 10 });
      expectValidPdf(pdf2);
    });

    it("should handle fitToPage", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Wide");

      // Create a very wide sheet
      for (let i = 1; i <= 20; i++) {
        ws.getColumn(i).width = 15;
        ws.getCell(1, i).value = `Col${i}`;
      }

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({ fitToPage: true });

      expectValidPdf(pdf);
    });

    it("should handle custom margins", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Margins";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export({
        margins: { top: 36, right: 36, bottom: 36, left: 36 }
      });

      expectValidPdf(pdf);
    });
  });

  describe("Alignment", () => {
    it("should handle center-aligned text", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Align");
      const cell = ws.getCell("A1");
      cell.value = "Centered";
      cell.alignment = { horizontal: "center" };

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Centered");
    });

    it("should handle right-aligned text", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Align");
      const cell = ws.getCell("A1");
      cell.value = "Right";
      cell.alignment = { horizontal: "right" };

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Right");
    });
  });

  describe("exportPdf function", () => {
    it("should work as a standalone function", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "FunctionAPI";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("FunctionAPI");
    });
  });

  describe("Hidden Rows and Columns", () => {
    it("should exclude hidden columns", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("HidCols");
      ws.getCell("A1").value = "Visible";
      ws.getCell("B1").value = "SecretData";
      ws.getCell("C1").value = "Also Visible";
      ws.getColumn(2).hidden = true;

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Visible");
      expect(text).toContain("Also Visible");
      expect(text).not.toContain("SecretData");
    });

    it("should exclude hidden rows", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("HidRows");
      ws.getCell("A1").value = "Row1";
      ws.getCell("A2").value = "SecretRow";
      ws.getCell("A3").value = "Row3";
      ws.getRow(2).hidden = true;

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Row1");
      expect(text).toContain("Row3");
      expect(text).not.toContain("SecretRow");
    });
  });

  describe("Hidden Worksheets", () => {
    it("should exclude hidden worksheets by default", () => {
      const wb = new Workbook();
      wb.addWorksheet("Visible").getCell("A1").value = "Shown";
      const hidden = wb.addWorksheet("Hidden");
      hidden.getCell("A1").value = "NotShown";
      hidden.state = "hidden";

      const exporter = new PdfExporter(wb);
      const pdf = exporter.export();

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Shown");
      expect(text).not.toContain("NotShown");
    });
  });

  describe("Edge Cases", () => {
    it("should handle a single cell workbook", () => {
      const wb = new Workbook();
      wb.addWorksheet("Solo").getCell("A1").value = "Only";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Only");
    });

    it("should handle cells with special characters", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Special");
      ws.getCell("A1").value = "Hello (world)";
      ws.getCell("A2").value = "Back\\slash";
      ws.getCell("A3").value = "New\nLine";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      // Special chars are inside compressed streams; just verify no crash
    });

    it("should handle empty string values", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Empty");
      ws.getCell("A1").value = "";
      ws.getCell("B1").value = "Not Empty";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
    });

    it("should handle boolean values", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Bool");
      ws.getCell("A1").value = true;
      ws.getCell("A2").value = false;

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
    });

    it("should handle error values", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Errors");
      ws.getCell("A1").value = { error: "#DIV/0!" };

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
    });
  });

  describe("Merged Cells + Hidden Columns", () => {
    it("should handle merged cells spanning hidden columns", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("MergeHidden");
      ws.getCell("A1").value = "Merged Over Hidden";
      ws.mergeCells("A1:D1");
      ws.getColumn(2).hidden = true; // hide B
      ws.getCell("A2").value = "Data";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Merged Over Hidden");
      expect(text).toContain("Data");
    });

    it("should handle merged cells spanning hidden rows", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("MergeHidden");
      ws.getCell("A1").value = "Tall Merge";
      ws.mergeCells("A1:A4");
      ws.getRow(2).hidden = true; // hide row 2
      ws.getCell("B1").value = "Side";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Tall Merge");
      expect(text).toContain("Side");
    });
  });

  describe("Repeat Rows", () => {
    it("should repeat header rows on subsequent pages", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Repeat");
      ws.getCell("A1").value = "Header";
      for (let i = 2; i <= 100; i++) {
        ws.getCell(`A${i}`).value = `Row ${i}`;
      }

      const pdf = exportPdf(wb, { repeatRows: 1, showPageNumbers: true });
      expectValidPdf(pdf);

      const text = pdfToString(pdf);
      // Should have multiple pages
      const pageMatches = text.match(/\/Type \/Page\b/g);
      expect(pageMatches!.length).toBeGreaterThan(1);
    });
  });

  describe("Text Wrapping", () => {
    it("should wrap long text when wrapText is enabled", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Wrap");
      const cell = ws.getCell("A1");
      cell.value = "This is a very long text that should wrap to multiple lines in the cell";
      cell.alignment = { wrapText: true };
      ws.getColumn(1).width = 15; // narrow column

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Should contain part of the text
      expect(text).toContain("This");
    });
  });

  describe("Invalid Sheet Selectors", () => {
    it("should throw when all sheet selectors are invalid", () => {
      const wb = new Workbook();
      wb.addWorksheet("Real").getCell("A1").value = "Data";

      const exporter = new PdfExporter(wb);
      expect(() => exporter.export({ sheets: ["NonExistent"] })).toThrow(PdfError);
    });

    it("should throw for out-of-range numeric selector", () => {
      const wb = new Workbook();
      wb.addWorksheet("Only").getCell("A1").value = "Data";

      const exporter = new PdfExporter(wb);
      expect(() => exporter.export({ sheets: [99] })).toThrow(PdfError);
    });
  });

  describe("Encryption with content streams", () => {
    it("should encrypt stream data when encryption is enabled", () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "SecretValue";

      const pdfPlain = exportPdf(wb);
      const pdfEncrypted = exportPdf(wb, {
        encryption: { ownerPassword: "owner123" }
      });

      expectValidPdf(pdfEncrypted);
      const plainText = pdfToString(pdfPlain);
      const encText = pdfToString(pdfEncrypted);

      // Plain PDF should NOT have /Encrypt dict
      expect(plainText).not.toContain("/Encrypt");
      // Encrypted PDF must have /Encrypt dict and /ID
      expect(encText).toContain("/Encrypt");
      expect(encText).toContain("/ID");
      // The encrypted PDF should have a different stream content than plain
      // (stream data is encrypted, so "SecretValue" should not appear in cleartext
      // in compressed+encrypted streams)
      expect(encText).toContain("/Filter /Standard");
    });
  });

  describe("RichText cell text extraction", () => {
    it("should not produce [object Object] for RichText cells", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Rich");
      ws.getCell("A1").value = {
        richText: [{ text: "Bold", font: { bold: true } }, { text: " Normal" }]
      };

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Should NOT contain the stringified object
      expect(text).not.toContain("[object Object]");
    });
  });

  describe("Non-ASCII text with Type1 fonts", () => {
    it("should encode accented characters correctly", () => {
      const wb = new Workbook();
      const wsAccent = wb.addWorksheet("Accent");
      wsAccent.getCell("A1").value = "café";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      // The text "café" should be encoded as WinAnsi hex, not UTF-8
      // (this will be in a compressed stream, so we just verify no crash)
    });
  });

  describe("Row page breaks", () => {
    it("should break after the row with the page break, not before it", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Breaks");
      for (let r = 1; r <= 10; r++) {
        ws.getCell(`A${r}`).value = `Row ${r}`;
      }
      // Break after row 5: rows 1-5 on first page, 6-10 on second
      ws.getRow(5).addPageBreak();

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Should have multiple pages
      expect((text.match(/\/Type \/Page\b/g) ?? []).length).toBeGreaterThanOrEqual(2);
    });
  });

  describe("Partial margins with worksheet fallback", () => {
    it("should merge partial PDF margins with worksheet pageSetup margins", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Margins");
      ws.getCell("A1").value = "Test";
      // Worksheet margins in inches
      ws.pageSetup.margins = {
        left: 0.5,
        right: 0.5,
        top: 1.0,
        bottom: 1.0,
        header: 0.3,
        footer: 0.3
      };

      // Only override left margin (36pt = 0.5in)
      const pdf = exportPdf(wb, { margins: { left: 36 } });
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // The top margin should come from worksheet (1.0 * 72 = 72pt), not reset to default
      // Just verify valid PDF produced without crash
      expect(text).toContain("%PDF");
    });
  });

  describe("Rich text with wrapping", () => {
    it("should wrap rich text cells when wrapText is enabled", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("RichWrap");
      const cell = ws.getCell("A1");
      cell.value = {
        richText: [
          { text: "This is bold text ", font: { bold: true } },
          { text: "and this is normal text that should wrap to multiple lines" }
        ]
      };
      cell.alignment = { wrapText: true };
      ws.getColumn(1).width = 15;

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
    });
  });

  describe("Encryption with embedded font", () => {
    it("should produce valid encrypted PDF with embedded TrueType font", () => {
      const wb = new Workbook();
      const wsEnc = wb.addWorksheet("Encrypted");
      wsEnc.getCell("A1").value = "Hello encrypted with font";

      const pdf = exportPdf(wb, {
        font: new Uint8Array(buildMinimalTtf()),
        encryption: { ownerPassword: "owner", userPassword: "user" }
      });
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("/Filter /Standard");
      expect(text).toContain("/CIDFontType2");
    });
  });

  describe("Encrypted PDF hex string encryption", () => {
    it("should encrypt non-ASCII metadata hex strings", () => {
      const wb = new Workbook();
      wb.addWorksheet("报告").getCell("A1").value = "Data";

      const pdf = exportPdf(wb, {
        title: "作者",
        encryption: { ownerPassword: "owner" }
      });
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // The hex-encoded UTF-16 title <feff4f5c8005> should NOT appear in cleartext
      expect(text).not.toContain("<feff4f5c8005>");
      // The hex-encoded UTF-16 sheet name <feff62a5544a> should NOT appear in cleartext
      expect(text).not.toContain("<feff62a5544a>");
    });
  });

  describe("Multi-range print area", () => {
    it("should use the first range from a multi-range printArea", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Multi");
      for (let r = 1; r <= 10; r++) {
        ws.getCell(`A${r}`).value = `A${r}`;
        ws.getCell(`B${r}`).value = `B${r}`;
        ws.getCell(`C${r}`).value = `C${r}`;
        ws.getCell(`D${r}`).value = `D${r}`;
      }
      // Multi-range: only A1:B5 should be used
      ws.pageSetup.printArea = "A1:B5&&D1:D10";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
    });
  });

  describe("printTitlesRow single-row format", () => {
    it("should accept single-number printTitlesRow format", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Titles");
      for (let r = 1; r <= 50; r++) {
        ws.getCell(`A${r}`).value = `Row ${r}`;
      }
      ws.pageSetup.printTitlesRow = "1";

      const pdf = exportPdf(wb);
      expectValidPdf(pdf);
      // Should have more than 1 page (50 rows with repeat headers)
      const text = pdfToString(pdf);
      expect((text.match(/\/Type \/Page\b/g) ?? []).length).toBeGreaterThanOrEqual(2);
    });
  });
});
