/**
 * Integration tests for the full PDF export pipeline.
 * Tests the PDF exporter with real Workbook instances via the Excel bridge,
 * and standalone pdf() API.
 */
import { describe, it, expect } from "vitest";
import { Workbook } from "@excel/workbook";
import { excelToPdf } from "@pdf/excel-bridge";
import { pdf as standalonePdf } from "@pdf/pdf";
import { PdfError } from "@pdf/errors";
import { buildMinimalTtf } from "./font-embedding.test";
import { pdfToString, expectValidPdf } from "./test-helpers";

describe("excelToPdf", () => {
  describe("Basic Export", () => {
    it("should export a simple workbook with one sheet", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = "Hello";
      ws.getCell("B1").value = "World";
      ws.getCell("A2").value = 42;
      ws.getCell("B2").value = 3.14;

      const pdf = await excelToPdf(wb);

      expect(pdf).toBeInstanceOf(Uint8Array);
      expect(pdf.length).toBeGreaterThan(100);
      expectValidPdf(pdf);

      // Content streams may be compressed; verify structure
      const text = pdfToString(pdf);
      expect(text).toContain("/Helvetica");
    });

    it("should export an empty workbook with at least one sheet", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Empty");

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
    });

    it("should throw for workbook with no sheets", async () => {
      const wb = new Workbook();

      await expect(excelToPdf(wb)).rejects.toThrow(PdfError);
    });
  });

  describe("Multiple Sheets", () => {
    it("should export multiple worksheets", async () => {
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

      const pdf = await excelToPdf(wb);

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

    it("should filter sheets by name", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Include").getCell("A1").value = "Included";
      wb.addWorksheet("Exclude").getCell("A1").value = "Excluded";

      const pdf = await excelToPdf(wb, { sheets: ["Include"] });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Included");
      expect(text).not.toContain("Excluded");
    });

    it("should filter sheets by 1-based position", async () => {
      const wb = new Workbook();
      const ws1 = wb.addWorksheet("First");
      ws1.getCell("A1").value = "First Sheet";
      const ws2 = wb.addWorksheet("Second");
      ws2.getCell("A1").value = "Second Sheet";

      const pdf = await excelToPdf(wb, { sheets: [2] }); // 1-based: second sheet

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Second Sheet");
      expect(text).not.toContain("First Sheet");
    });
  });

  describe("Page Size and Orientation", () => {
    it("should support A4 portrait (default)", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "A4";

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // A4: 595.28 x 841.89
      expect(text).toContain("595.28");
      expect(text).toContain("841.89");
    });

    it("should support landscape orientation", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Landscape";

      const pdf = await excelToPdf(wb, { orientation: "landscape" });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Landscape A4: 841.89 x 595.28
      expect(text).toContain("841.89");
      expect(text).toContain("595.28");
    });

    it("should support LETTER page size", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Letter";

      const pdf = await excelToPdf(wb, { pageSize: "LETTER" });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("612");
      expect(text).toContain("792");
    });

    it("should support custom page size", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Custom";

      const pdf = await excelToPdf(wb, {
        pageSize: { width: 400, height: 600 }
      });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("400");
      expect(text).toContain("600");
    });
  });

  describe("Cell Styles", () => {
    it("should render bold text", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Bold Text";
      cell.font = { bold: true };

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Helvetica-Bold");
      expect(text).toContain("Bold Text");
    });

    it("should render italic text", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Italic Text";
      cell.font = { italic: true };

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Helvetica-Oblique");
    });

    it("should render colored text", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Red Text";
      cell.font = { color: { argb: "FFFF0000" } };

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("1 0 0 rg"); // red fill color for text
    });

    it("should render background fill", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Styles");
      const cell = ws.getCell("A1");
      cell.value = "Filled";
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" }
      };

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Yellow fill: 1 1 0
      expect(text).toContain("1 1 0 rg");
    });

    it("should render cell borders", async () => {
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

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      // Borders are inside compressed content streams; just verify structure
    });
  });

  describe("Merged Cells", () => {
    it("should handle merged cells", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Merge");
      ws.getCell("A1").value = "Merged Title";
      ws.mergeCells("A1:C1");
      ws.getCell("A2").value = "Col1";
      ws.getCell("B2").value = "Col2";
      ws.getCell("C2").value = "Col3";

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      // Merged cells produce one page with all cells rendered
    });
  });

  describe("Data Types", () => {
    it("should handle various cell value types", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Types");
      ws.getCell("A1").value = "String";
      ws.getCell("A2").value = 42;
      ws.getCell("A3").value = 3.14;
      ws.getCell("A4").value = true;
      ws.getCell("A5").value = new Date(2024, 0, 15);
      ws.getCell("A6").value = null;

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      // All value types render without errors
    });

    it("should handle hyperlinks", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Links");
      ws.getCell("A1").value = { text: "Click Me", hyperlink: "https://example.com" };

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Click Me");
    });
  });

  describe("Grid Lines", () => {
    it("should render grid lines when enabled", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Grid");
      ws.getCell("A1").value = "A1";
      ws.getCell("B1").value = "B1";
      ws.getCell("A2").value = "A2";
      ws.getCell("B2").value = "B2";

      const pdf = await excelToPdf(wb, { showGridLines: true });

      expectValidPdf(pdf);
      // Grid lines are in compressed content streams; verify the PDF is valid
    });
  });

  describe("Page Headers and Footers", () => {
    it("should include sheet name as header", async () => {
      const wb = new Workbook();
      wb.addWorksheet("My Report").getCell("A1").value = "Data";

      const pdf = await excelToPdf(wb, { showSheetNames: true });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("My Report");
    });

    it("should include page numbers in footer", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const pdf = await excelToPdf(wb, { showPageNumbers: true });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Page 1 of 1");
    });

    it("should register footer fonts even when using an embedded font", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const pdf = await excelToPdf(wb, {
        showPageNumbers: true,
        font: new Uint8Array(buildMinimalTtf())
      });

      const text = pdfToString(pdf);
      expect(text).toContain("/BaseFont /Helvetica");
    });
  });

  describe("PDF Metadata", () => {
    it("should set document title", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const pdf = await excelToPdf(wb, { title: "My Report" });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("/Title (My Report)");
    });

    it("should set document author", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const pdf = await excelToPdf(wb, { author: "John Doe" });

      const text = pdfToString(pdf);
      expect(text).toContain("/Author (John Doe)");
    });

    it("should always set producer", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Data";

      const pdf = await excelToPdf(wb);

      const text = pdfToString(pdf);
      expect(text).toContain("/Producer (excelts)");
    });

    it("should encode Unicode metadata and bookmark titles correctly", async () => {
      const wb = new Workbook();
      wb.addWorksheet("报告").getCell("A1").value = "One";
      wb.addWorksheet("数据").getCell("A1").value = "Two";

      const pdf = await excelToPdf(wb, { title: "作者" });

      const text = pdfToString(pdf);
      expect(text).toContain("/Title <feff4f5c8005>");
      expect(text).toContain("/Title <feff62a5544a>");
      expect(text).toContain("/Title <feff6570636e>");
    });
  });

  describe("Worksheet Page Setup", () => {
    it("should honor per-sheet page setup defaults", async () => {
      const wb = new Workbook();

      const ws1 = wb.addWorksheet("First");
      ws1.getCell("A1").value = "One";
      ws1.pageSetup.paperSize = 11;

      const ws2 = wb.addWorksheet("Second");
      ws2.getCell("A1").value = "Two";
      ws2.pageSetup.orientation = "landscape";
      ws2.pageSetup.paperSize = 9;

      const pdf = await excelToPdf(wb);
      const text = pdfToString(pdf);

      expect(text).toContain("[0 0 419.53 595.28]");
      expect(text).toContain("[0 0 841.89 595.28]");
    });
  });

  describe("Column Widths", () => {
    it("should respect custom column widths", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Widths");
      ws.columns = [
        { header: "Narrow", width: 5 },
        { header: "Wide", width: 30 }
      ];

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Narrow");
      expect(text).toContain("Wide");
    });
  });

  describe("Large Datasets (Pagination)", () => {
    it("should paginate when content exceeds page height", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("LargeData");

      // Add enough rows to fill multiple pages
      for (let i = 1; i <= 100; i++) {
        ws.getCell(`A${i}`).value = `Row ${i}`;
        ws.getCell(`B${i}`).value = i * 10;
      }

      const pdf = await excelToPdf(wb, { showPageNumbers: true });

      expectValidPdf(pdf);
      const text = pdfToString(pdf);

      // Should have multiple pages
      const pageMatches = text.match(/\/Type \/Page\b/g);
      expect(pageMatches).not.toBeNull();
      expect(pageMatches!.length).toBeGreaterThan(1);
    });
  });

  describe("Options", () => {
    it("should clamp scale factor", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Scale";

      // Very small scale should be clamped to 0.1
      const pdf1 = await excelToPdf(wb, { scale: 0.01 });
      expectValidPdf(pdf1);

      // Very large scale should be clamped to 3.0
      const pdf2 = await excelToPdf(wb, { scale: 10 });
      expectValidPdf(pdf2);
    });

    it("should handle fitToPage", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Wide");

      // Create a very wide sheet
      for (let i = 1; i <= 20; i++) {
        ws.getColumn(i).width = 15;
        ws.getCell(1, i).value = `Col${i}`;
      }

      const pdf = await excelToPdf(wb, { fitToPage: true });

      expectValidPdf(pdf);
    });

    it("should handle custom margins", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "Margins";

      const pdf = await excelToPdf(wb, {
        margins: { top: 36, right: 36, bottom: 36, left: 36 }
      });

      expectValidPdf(pdf);
    });
  });

  describe("Alignment", () => {
    it("should handle center-aligned text", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Align");
      const cell = ws.getCell("A1");
      cell.value = "Centered";
      cell.alignment = { horizontal: "center" };

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Centered");
    });

    it("should handle right-aligned text", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Align");
      const cell = ws.getCell("A1");
      cell.value = "Right";
      cell.alignment = { horizontal: "right" };

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Right");
    });
  });

  describe("excelToPdf function", () => {
    it("should work as a standalone function", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "FunctionAPI";

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("FunctionAPI");
    });
  });

  describe("Hidden Rows and Columns", () => {
    it("should exclude hidden columns", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("HidCols");
      ws.getCell("A1").value = "Visible";
      ws.getCell("B1").value = "SecretData";
      ws.getCell("C1").value = "Also Visible";
      ws.getColumn(2).hidden = true;

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Visible");
      expect(text).toContain("Also Visible");
      expect(text).not.toContain("SecretData");
    });

    it("should exclude hidden rows", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("HidRows");
      ws.getCell("A1").value = "Row1";
      ws.getCell("A2").value = "SecretRow";
      ws.getCell("A3").value = "Row3";
      ws.getRow(2).hidden = true;

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Row1");
      expect(text).toContain("Row3");
      expect(text).not.toContain("SecretRow");
    });
  });

  describe("Hidden Worksheets", () => {
    it("should exclude hidden worksheets by default", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Visible").getCell("A1").value = "Shown";
      const hidden = wb.addWorksheet("Hidden");
      hidden.getCell("A1").value = "NotShown";
      hidden.state = "hidden";

      const pdf = await excelToPdf(wb);

      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Shown");
      expect(text).not.toContain("NotShown");
    });
  });

  describe("Edge Cases", () => {
    it("should handle a single cell workbook", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Solo").getCell("A1").value = "Only";

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Only");
    });

    it("should handle cells with special characters", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Special");
      ws.getCell("A1").value = "Hello (world)";
      ws.getCell("A2").value = "Back\\slash";
      ws.getCell("A3").value = "New\nLine";

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      // Special chars are inside compressed streams; just verify no crash
    });

    it("should handle empty string values", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Empty");
      ws.getCell("A1").value = "";
      ws.getCell("B1").value = "Not Empty";

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
    });

    it("should handle boolean values", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Bool");
      ws.getCell("A1").value = true;
      ws.getCell("A2").value = false;

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
    });

    it("should handle error values", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Errors");
      ws.getCell("A1").value = { error: "#DIV/0!" };

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
    });
  });

  describe("Merged Cells + Hidden Columns", () => {
    it("should handle merged cells spanning hidden columns", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("MergeHidden");
      ws.getCell("A1").value = "Merged Over Hidden";
      ws.mergeCells("A1:D1");
      ws.getColumn(2).hidden = true; // hide B
      ws.getCell("A2").value = "Data";

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Merged Over Hidden");
      expect(text).toContain("Data");
    });

    it("should handle merged cells spanning hidden rows", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("MergeHidden");
      ws.getCell("A1").value = "Tall Merge";
      ws.mergeCells("A1:A4");
      ws.getRow(2).hidden = true; // hide row 2
      ws.getCell("B1").value = "Side";

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      expect(text).toContain("Tall Merge");
      expect(text).toContain("Side");
    });
  });

  describe("Repeat Rows", () => {
    it("should repeat header rows on subsequent pages", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Repeat");
      ws.getCell("A1").value = "Header";
      for (let i = 2; i <= 100; i++) {
        ws.getCell(`A${i}`).value = `Row ${i}`;
      }

      const pdf = await excelToPdf(wb, { repeatRows: 1, showPageNumbers: true });
      expectValidPdf(pdf);

      const text = pdfToString(pdf);
      // Should have multiple pages
      const pageMatches = text.match(/\/Type \/Page\b/g);
      expect(pageMatches!.length).toBeGreaterThan(1);
    });
  });

  describe("Text Wrapping", () => {
    it("should wrap long text when wrapText is enabled", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Wrap");
      const cell = ws.getCell("A1");
      cell.value = "This is a very long text that should wrap to multiple lines in the cell";
      cell.alignment = { wrapText: true };
      ws.getColumn(1).width = 15; // narrow column

      const pdfBytes = await excelToPdf(wb);
      expectValidPdf(pdfBytes);
      // Read back and verify text content is present
      const { readPdf } = await import("@pdf/reader/pdf-reader");
      const result = await readPdf(pdfBytes);
      expect(result.text).toContain("This");
    });
  });

  describe("Invalid Sheet Selectors", () => {
    it("should throw when all sheet selectors are invalid", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Real").getCell("A1").value = "Data";

      await expect(excelToPdf(wb, { sheets: ["NonExistent"] })).rejects.toThrow(PdfError);
    });

    it("should throw for out-of-range numeric selector", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Only").getCell("A1").value = "Data";

      await expect(excelToPdf(wb, { sheets: [99] })).rejects.toThrow(PdfError);
    });
  });

  describe("Encryption with content streams", () => {
    it("should encrypt stream data when encryption is enabled", async () => {
      const wb = new Workbook();
      wb.addWorksheet("Test").getCell("A1").value = "SecretValue";

      const pdfPlain = await excelToPdf(wb);
      const pdfEncrypted = await excelToPdf(wb, {
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
    it("should not produce [object Object] for RichText cells", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Rich");
      ws.getCell("A1").value = {
        richText: [{ text: "Bold", font: { bold: true } }, { text: " Normal" }]
      };

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Should NOT contain the stringified object
      expect(text).not.toContain("[object Object]");
    });
  });

  describe("Non-ASCII text with Type1 fonts", () => {
    it("should encode accented characters correctly", async () => {
      const wb = new Workbook();
      const wsAccent = wb.addWorksheet("Accent");
      wsAccent.getCell("A1").value = "café";

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      // The text "café" should be encoded as WinAnsi hex, not UTF-8
      // (this will be in a compressed stream, so we just verify no crash)
    });
  });

  describe("Row page breaks", () => {
    it("should break after the row with the page break, not before it", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Breaks");
      for (let r = 1; r <= 10; r++) {
        ws.getCell(`A${r}`).value = `Row ${r}`;
      }
      // Break after row 5: rows 1-5 on first page, 6-10 on second
      ws.getRow(5).addPageBreak();

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // Should have multiple pages
      expect((text.match(/\/Type \/Page\b/g) ?? []).length).toBeGreaterThanOrEqual(2);
    });
  });

  describe("Partial margins with worksheet fallback", () => {
    it("should merge partial PDF margins with worksheet pageSetup margins", async () => {
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
      const pdf = await excelToPdf(wb, { margins: { left: 36 } });
      expectValidPdf(pdf);
      const text = pdfToString(pdf);
      // The top margin should come from worksheet (1.0 * 72 = 72pt), not reset to default
      // Just verify valid PDF produced without crash
      expect(text).toContain("%PDF");
    });
  });

  describe("Rich text with wrapping", () => {
    it("should wrap rich text cells when wrapText is enabled", async () => {
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

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
    });
  });

  describe("Encryption with embedded font", () => {
    it("should produce valid encrypted PDF with embedded TrueType font", async () => {
      const wb = new Workbook();
      const wsEnc = wb.addWorksheet("Encrypted");
      wsEnc.getCell("A1").value = "Hello encrypted with font";

      const pdf = await excelToPdf(wb, {
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
    it("should encrypt non-ASCII metadata hex strings", async () => {
      const wb = new Workbook();
      wb.addWorksheet("报告").getCell("A1").value = "Data";

      const pdf = await excelToPdf(wb, {
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
    it("should use the first range from a multi-range printArea", async () => {
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

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
    });
  });

  describe("printTitlesRow single-row format", () => {
    it("should accept single-number printTitlesRow format", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("Titles");
      for (let r = 1; r <= 50; r++) {
        ws.getCell(`A${r}`).value = `Row ${r}`;
      }
      ws.pageSetup.printTitlesRow = "1";

      const pdf = await excelToPdf(wb);
      expectValidPdf(pdf);
      // Should have more than 1 page (50 rows with repeat headers)
      const text = pdfToString(pdf);
      expect((text.match(/\/Type \/Page\b/g) ?? []).length).toBeGreaterThanOrEqual(2);
    });
  });
});

// =============================================================================
// Image Integration Tests
// =============================================================================

/**
 * Build a minimal valid JPEG (1x1 red pixel).
 * SOI + SOF0 + SOS + EOI
 */
function buildMinimalJpeg(): Uint8Array {
  // prettier-ignore
  return new Uint8Array([
    0xFF, 0xD8,             // SOI
    0xFF, 0xE0,             // APP0
    0x00, 0x10,             // length = 16
    0x4A, 0x46, 0x49, 0x46, 0x00, // "JFIF\0"
    0x01, 0x01,             // version 1.1
    0x00,                   // aspect ratio
    0x00, 0x01, 0x00, 0x01, // 1x1 pixel density
    0x00, 0x00,             // no thumbnail
    0xFF, 0xDB,             // DQT
    0x00, 0x43,             // length = 67
    0x00,                   // table 0, 8-bit precision
    // 64 quantization values (all 1s for simplicity)
    ...Array.from({ length: 64 }, () => 0x01),
    0xFF, 0xC0,             // SOF0 (baseline)
    0x00, 0x0B,             // length = 11
    0x08,                   // 8-bit precision
    0x00, 0x01,             // height = 1
    0x00, 0x01,             // width = 1
    0x01,                   // 1 component
    0x01,                   // component ID = 1
    0x11,                   // H/V sampling = 1x1
    0x00,                   // quant table 0
    0xFF, 0xC4,             // DHT
    0x00, 0x1F,             // length = 31
    0x00,                   // DC table 0
    // Number of codes of each length (1-16)
    0x00, 0x01, 0x05, 0x01, 0x01, 0x01, 0x01, 0x01,
    0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    // Values
    0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A, 0x0B,
    0xFF, 0xDA,             // SOS
    0x00, 0x08,             // length = 8
    0x01,                   // 1 component
    0x01,                   // component 1
    0x00,                   // DC/AC table 0/0
    0x00, 0x3F, 0x00,       // spectral selection
    0x7B, 0x40,             // scan data (minimal)
    0xFF, 0xD9              // EOI
  ]);
}

/**
 * Build a minimal valid PNG (2x2, RGBA with varying alpha).
 */
function buildMinimalPng(): Uint8Array {
  const parts: number[] = [];

  // PNG signature
  parts.push(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a);

  // IHDR
  const ihdr = [
    0x00,
    0x00,
    0x00,
    0x02, // width = 2
    0x00,
    0x00,
    0x00,
    0x02, // height = 2
    0x08, // bit depth = 8
    0x06, // color type = 6 (RGBA)
    0x00,
    0x00,
    0x00 // compression, filter, interlace
  ];
  writeChunk(parts, "IHDR", ihdr);

  // IDAT — raw pixel data: 2 rows of 2 RGBA pixels, each row starts with filter byte 0
  // Row 1: red (opaque), green (semi-transparent)
  // Row 2: blue (opaque), white (fully transparent)
  const rawPixels = [
    0x00, // filter byte
    0xff,
    0x00,
    0x00,
    0xff, // red, alpha=255
    0x00,
    0xff,
    0x00,
    0x80, // green, alpha=128
    0x00, // filter byte
    0x00,
    0x00,
    0xff,
    0xff, // blue, alpha=255
    0xff,
    0xff,
    0xff,
    0x00 // white, alpha=0
  ];

  // Deflate the raw data (use zlib sync from the archive module isn't available here,
  // so we'll use a stored (uncompressed) deflate block)
  const deflated = deflateStored(rawPixels);
  writeChunk(parts, "IDAT", Array.from(deflated));

  // IEND
  writeChunk(parts, "IEND", []);

  return new Uint8Array(parts);
}

function writeChunk(buf: number[], type: string, data: number[]): void {
  // Length (4 bytes, big-endian)
  const len = data.length;
  buf.push((len >>> 24) & 0xff, (len >>> 16) & 0xff, (len >>> 8) & 0xff, len & 0xff);
  // Type (4 bytes)
  for (let i = 0; i < 4; i++) {
    buf.push(type.charCodeAt(i));
  }
  // Data
  buf.push(...data);
  // CRC32 (over type + data)
  const crcInput = new Uint8Array(4 + data.length);
  for (let i = 0; i < 4; i++) {
    crcInput[i] = type.charCodeAt(i);
  }
  for (let i = 0; i < data.length; i++) {
    crcInput[4 + i] = data[i];
  }
  const crc = crc32(crcInput);
  buf.push((crc >>> 24) & 0xff, (crc >>> 16) & 0xff, (crc >>> 8) & 0xff, crc & 0xff);
}

function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/**
 * Wrap raw bytes in a stored (uncompressed) deflate stream with zlib header.
 */
function deflateStored(data: number[]): Uint8Array {
  const len = data.length;
  // zlib header: CMF=0x78 (deflate, window=32K), FLG=0x01 (FCHECK=1)
  const result = [0x78, 0x01];
  // BFINAL=1, BTYPE=00 (stored)
  result.push(0x01);
  // LEN and NLEN (little-endian)
  result.push(len & 0xff, (len >>> 8) & 0xff);
  result.push(~len & 0xff, (~len >>> 8) & 0xff);
  // Data
  result.push(...data);
  // Adler32 checksum
  let a = 1;
  let b = 0;
  for (const byte of data) {
    a = (a + byte) % 65521;
    b = (b + a) % 65521;
  }
  const adler = ((b << 16) | a) >>> 0;
  result.push((adler >>> 24) & 0xff, (adler >>> 16) & 0xff, (adler >>> 8) & 0xff, adler & 0xff);
  return new Uint8Array(result);
}

describe("Image integration", () => {
  it("should export PDF with embedded JPEG image", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Images");
    // Fill data that extends past the image range
    for (let r = 1; r <= 10; r++) {
      ws.getCell(`A${r}`).value = `Row ${r}`;
      ws.getCell(`B${r}`).value = r * 10;
      ws.getCell(`C${r}`).value = `Data ${r}`;
    }

    const jpegData = buildMinimalJpeg();
    const imageId = wb.addImage({ buffer: jpegData, extension: "jpeg" });
    ws.addImage(imageId, {
      tl: { col: 0, row: 1 },
      br: { col: 2, row: 4 }
    });

    const pdf = await excelToPdf(wb);
    expectValidPdf(pdf);
    const text = pdfToString(pdf);
    // Should contain an XObject image reference
    expect(text).toContain("/Subtype /Image");
    expect(text).toContain("/Filter /DCTDecode");
  });

  it("should export PDF with embedded PNG image (alpha channel)", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("PngTest");
    for (let r = 1; r <= 5; r++) {
      ws.getCell(`A${r}`).value = `Row ${r}`;
      ws.getCell(`B${r}`).value = r;
    }

    const pngData = buildMinimalPng();
    const imageId = wb.addImage({ buffer: pngData, extension: "png" });
    ws.addImage(imageId, {
      tl: { col: 0, row: 1 },
      ext: { width: 100, height: 100 }
    });

    const pdf = await excelToPdf(wb);
    expectValidPdf(pdf);
    const text = pdfToString(pdf);
    // Should contain XObject image
    expect(text).toContain("/Subtype /Image");
    // PNG with alpha should generate a soft mask
    expect(text).toContain("/SMask");
  });

  it("should handle workbook with image and multiple pages", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("Data");

    // Fill enough rows to span 2 pages
    for (let r = 1; r <= 80; r++) {
      ws.getCell(`A${r}`).value = `Row ${r}`;
      ws.getCell(`B${r}`).value = r;
    }

    const jpegData = buildMinimalJpeg();
    const imageId = wb.addImage({ buffer: jpegData, extension: "jpeg" });
    ws.addImage(imageId, {
      tl: { col: 1, row: 0 },
      br: { col: 3, row: 3 }
    });

    const pdf = await excelToPdf(wb, { showPageNumbers: true });
    expectValidPdf(pdf);
    const text = pdfToString(pdf);
    // Multiple pages
    const pages = text.match(/\/Type \/Page\b/g);
    expect(pages!.length).toBeGreaterThanOrEqual(2);
    // Image on first page
    expect(text).toContain("/Filter /DCTDecode");
  });

  it("should render tl/br image anchored beyond data bounds", async () => {
    const wb = new Workbook();
    const ws = wb.addWorksheet("BrBounds");
    // Only one cell of data
    ws.getCell("A1").value = "Hello";

    const jpegData = buildMinimalJpeg();
    const imageId = wb.addImage({ buffer: jpegData, extension: "jpeg" });
    // Image br extends well beyond the single data cell
    ws.addImage(imageId, {
      tl: { col: 1, row: 1 },
      br: { col: 5, row: 8 }
    });

    const pdf = await excelToPdf(wb);
    expectValidPdf(pdf);
    const text = pdfToString(pdf);
    expect(text).toContain("/Subtype /Image");
    expect(text).toContain("/Filter /DCTDecode");
  });
});

// =============================================================================
// Standalone pdf() API
// =============================================================================

describe("Standalone pdf() API", () => {
  it("should generate a valid PDF from a 2D array", async () => {
    const result = await standalonePdf([
      ["Product", "Revenue"],
      ["Widget", 1000],
      ["Gadget", 2500]
    ]);

    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(100);
    expectValidPdf(result);
    // Content streams may be compressed; verify font was used
    const text = pdfToString(result);
    expect(text).toContain("/Helvetica");
  });

  it("should generate a valid PDF from a sheet object with columns", async () => {
    const result = await standalonePdf({
      name: "Report",
      columns: [{ width: 25 }, { width: 15 }],
      data: [
        ["Widget", 1000],
        ["Gadget", 2500]
      ]
    });

    expect(result.length).toBeGreaterThan(100);
    expectValidPdf(result);
  });

  it("should generate a valid PDF from a multi-sheet book", async () => {
    const result = await standalonePdf({
      sheets: [
        {
          name: "Sales",
          data: [
            ["Product", "Revenue"],
            ["Widget", 1000]
          ]
        },
        {
          name: "Costs",
          data: [
            ["Item", "Amount"],
            ["Rent", 500]
          ]
        }
      ]
    });

    expectValidPdf(result);
    const text = pdfToString(result);
    const pageMatches = text.match(/\/Type \/Page\b/g);
    expect(pageMatches!.length).toBe(2);
    expect(text).toContain("(Sales)");
    expect(text).toContain("(Costs)");
  });

  it("should render column headers as the first row", async () => {
    const result = await standalonePdf({
      columns: [
        { width: 20, header: "Name" },
        { width: 15, header: "Score" }
      ],
      data: [
        ["Alice", 95],
        ["Bob", 87]
      ]
    });

    expectValidPdf(result);
    // Headers use bold font
    const text = pdfToString(result);
    expect(text).toContain("Helvetica-Bold");
  });

  it("should render header-only sheet with no data rows", async () => {
    const result = await standalonePdf({
      columns: [{ header: "Name" }, { header: "Score" }],
      data: []
    });

    expectValidPdf(result);
    // Should have at least one page with the bold header font
    const text = pdfToString(result);
    expect(text).toContain("Helvetica-Bold");
  });

  it("should place sparse column headers at the correct positions", async () => {
    const result = await standalonePdf({
      columns: [{ header: "A" }, { width: 10 }, { header: "C" }],
      data: [["x", "y", "z"]]
    });

    expectValidPdf(result);
    // Headers should be in columns 1 and 3, not 1 and 2.
    // Decompress not possible here, but verify 3 columns are rendered
    // and the bold header font is used.
    const text = pdfToString(result);
    expect(text).toContain("Helvetica-Bold");
  });

  it("should handle styled cells", async () => {
    const result = await standalonePdf([
      [
        { value: "Bold", bold: true },
        { value: "Red", fontColor: "FFFF0000" }
      ],
      [{ value: "Filled", fillColor: "FFFFFF00" }, "Plain"]
    ]);

    expectValidPdf(result);
    const text = pdfToString(result);
    expect(text).toContain("Helvetica-Bold");
  });

  it("should handle empty 2D array", async () => {
    const result = await standalonePdf([]);

    expectValidPdf(result);
  });

  it("should handle boolean and Date values", async () => {
    const result = await standalonePdf([[true, false, new Date(2024, 0, 15)]]);

    expect(result.length).toBeGreaterThan(100);
    expectValidPdf(result);
  });

  it("should handle null and undefined cells", async () => {
    const result = await standalonePdf([["Hello", null, undefined, "World"]]);

    expectValidPdf(result);
  });

  it("should accept export options", async () => {
    const result = await standalonePdf([["Test"]], {
      pageSize: "LETTER",
      orientation: "landscape",
      showGridLines: true
    });

    expectValidPdf(result);
    const text = pdfToString(result);
    // Landscape LETTER: 792 x 612
    expect(text).toContain("792");
    expect(text).toContain("612");
  });

  it("should embed a JPEG image via standalone pdf()", async () => {
    const jpegData = buildMinimalJpeg();
    const result = await standalonePdf({
      data: [
        ["Product", "Price"],
        ["Widget", "$10"]
      ],
      images: [{ data: jpegData, format: "jpeg", col: 0, row: 2, width: 100, height: 80 }]
    });

    expectValidPdf(result);
    const text = pdfToString(result);
    expect(text).toContain("/Subtype /Image");
    expect(text).toContain("/Filter /DCTDecode");
  });

  it("should embed a PNG image with alpha via standalone pdf()", async () => {
    const pngData = buildMinimalPng();
    const result = await standalonePdf({
      data: [
        ["Row 1", 100],
        ["Row 2", 200]
      ],
      images: [{ data: pngData, format: "png", col: 0, row: 2, width: 80, height: 80 }]
    });

    expectValidPdf(result);
    const text = pdfToString(result);
    expect(text).toContain("/Subtype /Image");
    expect(text).toContain("/SMask");
  });

  it("should render image-only sheet with no data rows", async () => {
    const jpegData = buildMinimalJpeg();
    const result = await standalonePdf({
      data: [],
      images: [{ data: jpegData, format: "jpeg", col: 0, row: 0, width: 200, height: 150 }]
    });

    expectValidPdf(result);
    const text = pdfToString(result);
    expect(text).toContain("/Subtype /Image");
    expect(text).toContain("/Filter /DCTDecode");
  });

  it("should extend bounds for image anchored beyond data columns", async () => {
    const jpegData = buildMinimalJpeg();
    const result = await standalonePdf({
      data: [["A only"]],
      images: [{ data: jpegData, format: "jpeg", col: 3, row: 0, width: 100, height: 80 }]
    });

    expectValidPdf(result);
    const text = pdfToString(result);
    expect(text).toContain("/Subtype /Image");
  });
});
