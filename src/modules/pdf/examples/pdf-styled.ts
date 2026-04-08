/**
 * Example: Styled PDF Export
 *
 * Covers:
 * - Font styles: bold, italic, underline, strikethrough, color, size
 * - Cell fill / background colors (solid pattern)
 * - Cell borders (thin, medium, thick, dashed, dotted)
 * - Horizontal alignment: left, center, right
 * - Vertical alignment: top, middle, bottom
 * - Text wrapping
 * - Text indentation
 * - Text rotation (angled and vertical stacked)
 * - Merged cells
 * - Number formatting (currency, percentage, dates)
 * - Hyperlinks
 * - Rich text (mixed formatting in one cell)
 * - Theme-like color usage
 * - Column widths and row heights
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Workbook, excelToPdf } from "../../../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/pdf-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const wb = new Workbook();

// =============================================================================
// Sheet 1: Font Styles
// =============================================================================

const wsFonts = wb.addWorksheet("Fonts");
wsFonts.columns = [
  { header: "Style", width: 20 },
  { header: "Sample", width: 30 }
];

wsFonts.getCell("A2").value = "Bold";
wsFonts.getCell("B2").value = "Bold text";
wsFonts.getCell("B2").font = { bold: true };

wsFonts.getCell("A3").value = "Italic";
wsFonts.getCell("B3").value = "Italic text";
wsFonts.getCell("B3").font = { italic: true };

wsFonts.getCell("A4").value = "Underline";
wsFonts.getCell("B4").value = "Underlined text";
wsFonts.getCell("B4").font = { underline: true };

wsFonts.getCell("A5").value = "Strikethrough";
wsFonts.getCell("B5").value = "Struck through";
wsFonts.getCell("B5").font = { strike: true };

wsFonts.getCell("A6").value = "Bold + Italic";
wsFonts.getCell("B6").value = "Bold italic text";
wsFonts.getCell("B6").font = { bold: true, italic: true };

wsFonts.getCell("A7").value = "Large Red";
wsFonts.getCell("B7").value = "Size 18 red";
wsFonts.getCell("B7").font = { size: 18, color: { argb: "FFFF0000" } };

wsFonts.getCell("A8").value = "Small Blue";
wsFonts.getCell("B8").value = "Size 8 blue";
wsFonts.getCell("B8").font = { size: 8, color: { argb: "FF0000FF" } };

wsFonts.getCell("A9").value = "Font Family";
wsFonts.getCell("B9").value = "Courier New";
wsFonts.getCell("B9").font = { name: "Courier New" };

// =============================================================================
// Sheet 2: Fills and Borders
// =============================================================================

const wsFills = wb.addWorksheet("Fills & Borders");
wsFills.columns = [
  { header: "Description", width: 25 },
  { header: "Cell", width: 25 }
];

// Solid fills
wsFills.getCell("A2").value = "Yellow fill";
wsFills.getCell("B2").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFFF00" }
};
wsFills.getCell("B2").value = "Yellow";

wsFills.getCell("A3").value = "Green fill";
wsFills.getCell("B3").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF00CC00" }
};
wsFills.getCell("B3").value = "Green";
wsFills.getCell("B3").font = { color: { argb: "FFFFFFFF" } };

wsFills.getCell("A4").value = "Dark blue fill";
wsFills.getCell("B4").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF003366" }
};
wsFills.getCell("B4").value = "Dark Blue";
wsFills.getCell("B4").font = { color: { argb: "FFFFFFFF" }, bold: true };

// Borders
wsFills.getCell("A6").value = "Thin border";
wsFills.getCell("B6").value = "All sides thin";
wsFills.getCell("B6").border = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" }
};

wsFills.getCell("A7").value = "Medium border";
wsFills.getCell("B7").value = "All sides medium";
wsFills.getCell("B7").border = {
  top: { style: "medium" },
  left: { style: "medium" },
  bottom: { style: "medium" },
  right: { style: "medium" }
};

wsFills.getCell("A8").value = "Thick border";
wsFills.getCell("B8").value = "All sides thick";
wsFills.getCell("B8").border = {
  top: { style: "thick" },
  left: { style: "thick" },
  bottom: { style: "thick" },
  right: { style: "thick" }
};

wsFills.getCell("A9").value = "Colored border";
wsFills.getCell("B9").value = "Red dashed border";
wsFills.getCell("B9").border = {
  top: { style: "dashed", color: { argb: "FFFF0000" } },
  left: { style: "dashed", color: { argb: "FFFF0000" } },
  bottom: { style: "dashed", color: { argb: "FFFF0000" } },
  right: { style: "dashed", color: { argb: "FFFF0000" } }
};

wsFills.getCell("A10").value = "Dotted border";
wsFills.getCell("B10").value = "Blue dotted";
wsFills.getCell("B10").border = {
  top: { style: "dotted", color: { argb: "FF0000FF" } },
  left: { style: "dotted", color: { argb: "FF0000FF" } },
  bottom: { style: "dotted", color: { argb: "FF0000FF" } },
  right: { style: "dotted", color: { argb: "FF0000FF" } }
};

// =============================================================================
// Sheet 3: Alignment, Wrapping, Indent, Rotation, Merged Cells
// =============================================================================

const wsAlign = wb.addWorksheet("Alignment");
wsAlign.columns = [{ width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }];

// Horizontal alignment
wsAlign.getCell("A1").value = "Left";
wsAlign.getCell("A1").alignment = { horizontal: "left" };
wsAlign.getCell("B1").value = "Center";
wsAlign.getCell("B1").alignment = { horizontal: "center" };
wsAlign.getCell("C1").value = "Right";
wsAlign.getCell("C1").alignment = { horizontal: "right" };

// Vertical alignment
wsAlign.getRow(3).height = 40;
wsAlign.getCell("A3").value = "Top";
wsAlign.getCell("A3").alignment = { vertical: "top" };
wsAlign.getCell("B3").value = "Middle";
wsAlign.getCell("B3").alignment = { vertical: "middle" };
wsAlign.getCell("C3").value = "Bottom";
wsAlign.getCell("C3").alignment = { vertical: "bottom" };

// Text wrapping
wsAlign.getRow(5).height = 50;
wsAlign.getCell("A5").value =
  "This is a long piece of text that should wrap within the cell boundaries";
wsAlign.getCell("A5").alignment = { wrapText: true };

// Text indentation
wsAlign.getCell("A7").value = "No indent";
wsAlign.getCell("A8").value = "Indent 1";
wsAlign.getCell("A8").alignment = { indent: 1 };
wsAlign.getCell("A9").value = "Indent 3";
wsAlign.getCell("A9").alignment = { indent: 3 };

// Text rotation
wsAlign.getRow(11).height = 60;
wsAlign.getCell("A11").value = "45° rotation";
wsAlign.getCell("A11").alignment = { textRotation: 45 };
wsAlign.getCell("B11").value = "90° rotation";
wsAlign.getCell("B11").alignment = { textRotation: 90 };
wsAlign.getCell("C11").value = "Vertical";
wsAlign.getCell("C11").alignment = { textRotation: 255 }; // 255 = vertical stacked

// Merged cells
wsAlign.mergeCells("A13:D13");
wsAlign.getCell("A13").value = "Merged across 4 columns";
wsAlign.getCell("A13").alignment = { horizontal: "center" };
wsAlign.getCell("A13").font = { bold: true, size: 14 };
wsAlign.getCell("A13").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFE0E0E0" }
};

wsAlign.mergeCells("A15:A18");
wsAlign.getCell("A15").value = "Merged down 4 rows";
wsAlign.getCell("A15").alignment = { vertical: "middle", horizontal: "center" };

// =============================================================================
// Sheet 4: Number Formatting, Hyperlinks, Rich Text
// =============================================================================

const wsData = wb.addWorksheet("Data Types");
wsData.columns = [
  { header: "Type", width: 20 },
  { header: "Value", width: 30 }
];

// Number formats
wsData.getCell("A2").value = "Currency";
wsData.getCell("B2").value = 1234.56;
wsData.getCell("B2").numFmt = "$#,##0.00";

wsData.getCell("A3").value = "Percentage";
wsData.getCell("B3").value = 0.8523;
wsData.getCell("B3").numFmt = "0.0%";

wsData.getCell("A4").value = "Date";
wsData.getCell("B4").value = new Date(2025, 2, 28);
wsData.getCell("B4").numFmt = "yyyy-mm-dd";

wsData.getCell("A5").value = "Accounting";
wsData.getCell("B5").value = -4500;
wsData.getCell("B5").numFmt = "_($#,##0.00_);[Red]_($#,##0.00)";

wsData.getCell("A6").value = "Number";
wsData.getCell("B6").value = 3141592.65358;
wsData.getCell("B6").numFmt = "#,##0.00";

// Hyperlink
wsData.getCell("A8").value = "Hyperlink";
wsData.getCell("B8").value = { text: "Visit GitHub", hyperlink: "https://github.com" };
wsData.getCell("B8").font = { color: { argb: "FF0563C1" }, underline: true };

// Rich text
wsData.getCell("A10").value = "Rich Text";
wsData.getCell("B10").value = {
  richText: [
    { text: "Bold", font: { bold: true } },
    { text: " and " },
    { text: "italic", font: { italic: true } },
    { text: " and " },
    { text: "red", font: { color: { argb: "FFFF0000" } } },
    { text: " and " },
    { text: "large", font: { size: 16 } },
    { text: " mixed." }
  ]
};

// Rich text with wrapping
wsData.getRow(12).height = 60;
wsData.getCell("A12").value = "Rich Wrapped";
wsData.getCell("B12").value = {
  richText: [
    { text: "This is bold text ", font: { bold: true } },
    { text: "followed by normal text that should wrap to multiple lines in the cell" }
  ]
};
wsData.getCell("B12").alignment = { wrapText: true };

// =============================================================================
// Export
// =============================================================================

const pdf = await excelToPdf(wb, {
  showGridLines: true,
  showSheetNames: true,
  showPageNumbers: true,
  title: "Styled PDF Export Demo",
  author: "excelts"
});

fs.writeFileSync(path.join(outDir, "pdf-styled.pdf"), pdf);
console.log("pdf-styled.pdf generated — fonts, fills, borders, alignment, merged, rich text");
