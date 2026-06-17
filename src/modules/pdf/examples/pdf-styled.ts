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

import { Cell, Row, Workbook, Worksheet } from "@excel/index";

import { Pdf } from "../../../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/pdf-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const wb = Workbook.create();

// =============================================================================
// Sheet 1: Font Styles
// =============================================================================

const wsFonts = Workbook.addWorksheet(wb, "Fonts");
Worksheet.setColumns(wsFonts, [
  { header: "Style", width: 20 },
  { header: "Sample", width: 30 }
]);

Cell.setValue(wsFonts, "A2", "Bold");
Cell.setValue(wsFonts, "B2", "Bold text");
Cell.setStyle(wsFonts, "B2", { font: { bold: true } });

Cell.setValue(wsFonts, "A3", "Italic");
Cell.setValue(wsFonts, "B3", "Italic text");
Cell.setStyle(wsFonts, "B3", { font: { italic: true } });

Cell.setValue(wsFonts, "A4", "Underline");
Cell.setValue(wsFonts, "B4", "Underlined text");
Cell.setStyle(wsFonts, "B4", { font: { underline: true } });

Cell.setValue(wsFonts, "A5", "Strikethrough");
Cell.setValue(wsFonts, "B5", "Struck through");
Cell.setStyle(wsFonts, "B5", { font: { strike: true } });

Cell.setValue(wsFonts, "A6", "Bold + Italic");
Cell.setValue(wsFonts, "B6", "Bold italic text");
Cell.setStyle(wsFonts, "B6", { font: { bold: true, italic: true } });

Cell.setValue(wsFonts, "A7", "Large Red");
Cell.setValue(wsFonts, "B7", "Size 18 red");
Cell.setStyle(wsFonts, "B7", { font: { size: 18, color: { argb: "FFFF0000" } } });

Cell.setValue(wsFonts, "A8", "Small Blue");
Cell.setValue(wsFonts, "B8", "Size 8 blue");
Cell.setStyle(wsFonts, "B8", { font: { size: 8, color: { argb: "FF0000FF" } } });

Cell.setValue(wsFonts, "A9", "Font Family");
Cell.setValue(wsFonts, "B9", "Courier New");
Cell.setStyle(wsFonts, "B9", { font: { name: "Courier New" } });

// =============================================================================
// Sheet 2: Fills and Borders
// =============================================================================

const wsFills = Workbook.addWorksheet(wb, "Fills & Borders");
Worksheet.setColumns(wsFills, [
  { header: "Description", width: 25 },
  { header: "Cell", width: 25 }
]);

// Solid fills
Cell.setValue(wsFills, "A2", "Yellow fill");
Cell.setStyle(wsFills, "B2", {
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" }
  }
});
Cell.setValue(wsFills, "B2", "Yellow");

Cell.setValue(wsFills, "A3", "Green fill");
Cell.setStyle(wsFills, "B3", {
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF00CC00" }
  }
});
Cell.setValue(wsFills, "B3", "Green");
Cell.setStyle(wsFills, "B3", { font: { color: { argb: "FFFFFFFF" } } });

Cell.setValue(wsFills, "A4", "Dark blue fill");
Cell.setStyle(wsFills, "B4", {
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF003366" }
  }
});
Cell.setValue(wsFills, "B4", "Dark Blue");
Cell.setStyle(wsFills, "B4", { font: { color: { argb: "FFFFFFFF" }, bold: true } });

// Borders
Cell.setValue(wsFills, "A6", "Thin border");
Cell.setValue(wsFills, "B6", "All sides thin");
Cell.setStyle(wsFills, "B6", {
  border: {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" }
  }
});

Cell.setValue(wsFills, "A7", "Medium border");
Cell.setValue(wsFills, "B7", "All sides medium");
Cell.setStyle(wsFills, "B7", {
  border: {
    top: { style: "medium" },
    left: { style: "medium" },
    bottom: { style: "medium" },
    right: { style: "medium" }
  }
});

Cell.setValue(wsFills, "A8", "Thick border");
Cell.setValue(wsFills, "B8", "All sides thick");
Cell.setStyle(wsFills, "B8", {
  border: {
    top: { style: "thick" },
    left: { style: "thick" },
    bottom: { style: "thick" },
    right: { style: "thick" }
  }
});

Cell.setValue(wsFills, "A9", "Colored border");
Cell.setValue(wsFills, "B9", "Red dashed border");
Cell.setStyle(wsFills, "B9", {
  border: {
    top: { style: "dashed", color: { argb: "FFFF0000" } },
    left: { style: "dashed", color: { argb: "FFFF0000" } },
    bottom: { style: "dashed", color: { argb: "FFFF0000" } },
    right: { style: "dashed", color: { argb: "FFFF0000" } }
  }
});

Cell.setValue(wsFills, "A10", "Dotted border");
Cell.setValue(wsFills, "B10", "Blue dotted");
Cell.setStyle(wsFills, "B10", {
  border: {
    top: { style: "dotted", color: { argb: "FF0000FF" } },
    left: { style: "dotted", color: { argb: "FF0000FF" } },
    bottom: { style: "dotted", color: { argb: "FF0000FF" } },
    right: { style: "dotted", color: { argb: "FF0000FF" } }
  }
});

// =============================================================================
// Sheet 3: Alignment, Wrapping, Indent, Rotation, Merged Cells
// =============================================================================

const wsAlign = Workbook.addWorksheet(wb, "Alignment");
Worksheet.setColumns(wsAlign, [{ width: 20 }, { width: 20 }, { width: 20 }, { width: 20 }]);

// Horizontal alignment
Cell.setValue(wsAlign, "A1", "Left");
Cell.setStyle(wsAlign, "A1", { alignment: { horizontal: "left" } });
Cell.setValue(wsAlign, "B1", "Center");
Cell.setStyle(wsAlign, "B1", { alignment: { horizontal: "center" } });
Cell.setValue(wsAlign, "C1", "Right");
Cell.setStyle(wsAlign, "C1", { alignment: { horizontal: "right" } });

// Vertical alignment
Row.setHeight(wsAlign, 3, 40);
Cell.setValue(wsAlign, "A3", "Top");
Cell.setStyle(wsAlign, "A3", { alignment: { vertical: "top" } });
Cell.setValue(wsAlign, "B3", "Middle");
Cell.setStyle(wsAlign, "B3", { alignment: { vertical: "middle" } });
Cell.setValue(wsAlign, "C3", "Bottom");
Cell.setStyle(wsAlign, "C3", { alignment: { vertical: "bottom" } });

// Text wrapping
Row.setHeight(wsAlign, 5, 50);
Cell.setValue(
  wsAlign,
  "A5",
  "This is a long piece of text that should wrap within the cell boundaries"
);
Cell.setStyle(wsAlign, "A5", { alignment: { wrapText: true } });

// Text indentation
Cell.setValue(wsAlign, "A7", "No indent");
Cell.setValue(wsAlign, "A8", "Indent 1");
Cell.setStyle(wsAlign, "A8", { alignment: { indent: 1 } });
Cell.setValue(wsAlign, "A9", "Indent 3");
Cell.setStyle(wsAlign, "A9", { alignment: { indent: 3 } });

// Text rotation — basic
Row.setHeight(wsAlign, 11, 60);
Cell.setValue(wsAlign, "A11", "45° rotation");
Cell.setStyle(wsAlign, "A11", { alignment: { textRotation: 45 } });
Cell.setValue(wsAlign, "B11", "90° rotation");
Cell.setStyle(wsAlign, "B11", { alignment: { textRotation: 90 } });
Cell.setValue(wsAlign, "C11", "Vertical");
Cell.setStyle(wsAlign, "C11", { alignment: { textRotation: "vertical" } });
Cell.setValue(wsAlign, "D11", "-90° rotation");
Cell.setStyle(wsAlign, "D11", { alignment: { textRotation: -90 } });

const thinBorder = {
  top: { style: "thin" as const },
  bottom: { style: "thin" as const },
  left: { style: "thin" as const },
  right: { style: "thin" as const }
};

// Text rotation — 90° with alignment combinations
Row.setHeight(wsAlign, 13, 80);
Cell.setValue(wsAlign, "A13", "center/top");
Cell.setStyle(wsAlign, "A13", {
  alignment: { textRotation: 90, horizontal: "center", vertical: "top" }
});
Cell.setStyle(wsAlign, "A13", { border: thinBorder });
Cell.setValue(wsAlign, "B13", "center/mid");
Cell.setStyle(wsAlign, "B13", {
  alignment: {
    textRotation: 90,
    horizontal: "center",
    vertical: "middle"
  }
});
Cell.setStyle(wsAlign, "B13", { border: thinBorder });
Cell.setValue(wsAlign, "C13", "left/bot");
Cell.setStyle(wsAlign, "C13", {
  alignment: { textRotation: 90, horizontal: "left", vertical: "bottom" }
});
Cell.setStyle(wsAlign, "C13", { border: thinBorder });
Cell.setValue(wsAlign, "D13", "right/bot");
Cell.setStyle(wsAlign, "D13", {
  alignment: { textRotation: 90, horizontal: "right", vertical: "bottom" }
});
Cell.setStyle(wsAlign, "D13", { border: thinBorder });

// Text rotation — 45° with slanted borders and alignment
Row.setHeight(wsAlign, 15, 60);
Cell.setValue(wsAlign, "A15", "45° top");
Cell.setStyle(wsAlign, "A15", {
  alignment: { textRotation: 45, horizontal: "center", vertical: "top" }
});
Cell.setStyle(wsAlign, "A15", { border: thinBorder });
Cell.setValue(wsAlign, "B15", "45° mid");
Cell.setStyle(wsAlign, "B15", {
  alignment: {
    textRotation: 45,
    horizontal: "center",
    vertical: "middle"
  }
});
Cell.setStyle(wsAlign, "B15", { border: thinBorder });
Cell.setValue(wsAlign, "C15", "45° left");
Cell.setStyle(wsAlign, "C15", {
  alignment: { textRotation: 45, horizontal: "left", vertical: "bottom" }
});
Cell.setStyle(wsAlign, "C15", { border: thinBorder });
Cell.setValue(wsAlign, "D15", "45° right");
Cell.setStyle(wsAlign, "D15", {
  alignment: { textRotation: 45, horizontal: "right", vertical: "bottom" }
});
Cell.setStyle(wsAlign, "D15", { border: thinBorder });

// Merged cells
Worksheet.merge(wsAlign, "A17:D17");
Cell.setValue(wsAlign, "A17", "Merged across 4 columns");
Cell.setStyle(wsAlign, "A17", { alignment: { horizontal: "center" } });
Cell.setStyle(wsAlign, "A17", { font: { bold: true, size: 14 } });
Cell.setStyle(wsAlign, "A17", {
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0E0E0" }
  }
});

Worksheet.merge(wsAlign, "A19:A22");
Cell.setValue(wsAlign, "A19", "Merged down 4 rows");
Cell.setStyle(wsAlign, "A19", { alignment: { vertical: "middle", horizontal: "center" } });

// =============================================================================
// Sheet 4: Number Formatting, Hyperlinks, Rich Text
// =============================================================================

const wsData = Workbook.addWorksheet(wb, "Data Types");
Worksheet.setColumns(wsData, [
  { header: "Type", width: 20 },
  { header: "Value", width: 30 }
]);

// Number formats
Cell.setValue(wsData, "A2", "Currency");
Cell.setValue(wsData, "B2", 1234.56);
Cell.setStyle(wsData, "B2", { numFmt: "$#,##0.00" });

Cell.setValue(wsData, "A3", "Percentage");
Cell.setValue(wsData, "B3", 0.8523);
Cell.setStyle(wsData, "B3", { numFmt: "0.0%" });

Cell.setValue(wsData, "A4", "Date");
Cell.setValue(wsData, "B4", new Date(2025, 2, 28));
Cell.setStyle(wsData, "B4", { numFmt: "yyyy-mm-dd" });

Cell.setValue(wsData, "A5", "Accounting");
Cell.setValue(wsData, "B5", -4500);
Cell.setStyle(wsData, "B5", { numFmt: "_($#,##0.00_);[Red]_($#,##0.00)" });

Cell.setValue(wsData, "A6", "Number");
Cell.setValue(wsData, "B6", 3141592.65358);
Cell.setStyle(wsData, "B6", { numFmt: "#,##0.00" });

// Hyperlink
Cell.setValue(wsData, "A8", "Hyperlink");
Cell.setValue(wsData, "B8", { text: "Visit GitHub", hyperlink: "https://github.com" });
Cell.setStyle(wsData, "B8", { font: { color: { argb: "FF0563C1" }, underline: true } });

// Rich text
Cell.setValue(wsData, "A10", "Rich Text");
Cell.setValue(wsData, "B10", {
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
});

// Rich text with wrapping
Row.setHeight(wsData, 12, 60);
Cell.setValue(wsData, "A12", "Rich Wrapped");
Cell.setValue(wsData, "B12", {
  richText: [
    { text: "This is bold text ", font: { bold: true } },
    { text: "followed by normal text that should wrap to multiple lines in the cell" }
  ]
});
Cell.setStyle(wsData, "B12", { alignment: { wrapText: true } });

// =============================================================================
// Export
// =============================================================================

const pdf = await Pdf.fromExcel(wb, {
  showGridLines: true,
  showSheetNames: true,
  showPageNumbers: true,
  title: "Styled PDF Export Demo",
  author: "excelts"
});

fs.writeFileSync(path.join(outDir, "pdf-styled.pdf"), pdf);
console.log("pdf-styled.pdf generated — fonts, fills, borders, alignment, merged, rich text");
