/**
 * Example: Rich text overflow and wrapping in PDF rendering
 *
 * Demonstrates correct handling of:
 *
 * 1. Rich text with 2-3 different sizes in a single (non-merged) cell
 *    overflowing into adjacent empty cells (same as plain text).
 *
 * 2. Gridlines/borders hidden under text overflow regions
 *    (white fill drawn after borders, before text).
 *
 * 3. Per-run font size measurement for word-wrap decisions
 *    (small-font runs wrap at their actual width, not max font width).
 *
 * 4. Layout row height consistent with render line count for rich text.
 *
 * Run: npx tsx src/modules/pdf/examples/pdf-rich-text-overflow.ts
 * Output: tmp/pdf-examples/rich-text-overflow.pdf
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
// Sheet 1: Rich text overflow in single (non-merged) cells
//
// Rich text wider than the cell overflows into adjacent empty cells,
// just like plain text does.
// =============================================================================

const ws1 = wb.addWorksheet("RichText-Overflow");
ws1.columns = [
  { header: "Description", width: 18 },
  { header: "Narrow Single Cell", width: 12 },
  { header: "C (empty)", width: 12 },
  { header: "D (empty)", width: 12 },
  { header: "Wide Merged Cell", width: 12 }
];

// Row 2: Rich text exceeds narrow cell width → overflows into C2, D2
ws1.getCell("A2").value = "8pt + 16pt (no wrap)";
ws1.getCell("B2").value = {
  richText: [
    { text: "Small(8)", font: { size: 8 } },
    { text: " BIG(16)", font: { size: 16 } }
  ]
};
// C2, D2 empty — rich text overflows into them
// Merged cell for comparison (wider, so text fits without overflow)
ws1.mergeCells("E2:G2");
ws1.getCell("E2").value = {
  richText: [
    { text: "Small(8)", font: { size: 8 } },
    { text: " BIG(16)", font: { size: 16 } }
  ]
};

// Row 3: Three sizes, narrow cell
ws1.getCell("A3").value = "7 + 8 + 14 (no wrap)";
ws1.getCell("B3").value = {
  richText: [
    { text: "Tiny", font: { size: 7 } },
    { text: " Mid", font: { size: 8 } },
    { text: " BIG", font: { size: 14 } }
  ]
};
ws1.mergeCells("E3:G3");
ws1.getCell("E3").value = {
  richText: [
    { text: "Tiny", font: { size: 7 } },
    { text: " Mid", font: { size: 8 } },
    { text: " BIG", font: { size: 14 } }
  ]
};

// Row 4: Plain text comparison — overflow works correctly
ws1.getCell("A4").value = "Plain text overflow";
ws1.getCell("B4").value = "This plain text overflows into C4 and D4 correctly";

// =============================================================================
// Sheet 2: Overflow region hides gridlines and borders
//
// When cell A has text that overflows into cells B, C, D (which are empty),
// gridlines and borders in the overflow area are hidden — matching Excel.
// =============================================================================

const ws2 = wb.addWorksheet("Overflow-Erase");
ws2.columns = [
  { header: "A", width: 12 },
  { header: "B", width: 8 },
  { header: "C", width: 8 },
  { header: "D", width: 8 },
  { header: "E", width: 8 },
  { header: "F", width: 8 }
];

// Long text overflows → gridlines visible in overflow region
ws2.getCell("A2").value =
  "This very long text overflows across several columns. In Excel the gridlines disappear under the text. In PDF they remain visible.";

// Same but with explicit borders on some cells
ws2.getCell("A4").value = "Bordered cell with overflow:";
ws2.getCell("A4").border = {
  top: { style: "thin" },
  bottom: { style: "thin" },
  left: { style: "thin" },
  right: { style: "thin" }
};
// B4 has border but is empty → border drawn even though text from A4 covers it visually in Excel
ws2.getCell("B4").border = {
  top: { style: "thin" },
  bottom: { style: "thin" },
  left: { style: "thin" },
  right: { style: "thin" }
};

// =============================================================================
// Sheet 3: Per-run font size measurement for wrapped rich text
//
// When wrapping rich text with different font sizes, each run's text is
// measured at its own font size for line-break decisions. This produces
// correct character density per line regardless of size differences.
// =============================================================================

const ws3 = wb.addWorksheet("PerRun-WrapSize");
ws3.columns = [
  { header: "Case", width: 25 },
  { header: "Wrapped Rich Text", width: 25 },
  { header: "Expected Layout", width: 40 }
];

// Case from issue: 8pt vs 7pt — only 1pt difference but looks huge
ws3.getRow(2).height = 30;
ws3.getCell("A2").value = "8pt vs 7pt (wrap)";
ws3.getCell("B2").value = {
  richText: [
    { text: "1TEXT-XD", font: { size: 8 } },
    { text: "(ex.2)(ex=1)", font: { size: 7 } }
  ]
};
ws3.getCell("B2").alignment = { wrapText: true };
ws3.getCell("C2").value = "Both runs should be nearly same visual size (8:7 ratio)";
ws3.getCell("C2").alignment = { wrapText: true };

// Extreme case: 16pt header + 7pt body
ws3.getRow(3).height = 60;
ws3.getCell("A3").value = "16pt + 7pt body (wrap)";
ws3.getCell("B3").value = {
  richText: [
    { text: "TITLE ", font: { size: 16 } },
    {
      text: "The body text is 7pt and should wrap normally at its own size, fitting many more characters per line than it currently does.",
      font: { size: 7 }
    }
  ]
};
ws3.getCell("B3").alignment = { wrapText: true };
ws3.getCell("C3").value =
  "Body text wraps as if it were 16pt wide → only ~4 chars/line instead of ~8. Actual render at 7pt leaves huge gaps.";
ws3.getCell("C3").alignment = { wrapText: true };

// Reference: same text at uniform 7pt (correct wrap behavior)
ws3.getRow(4).height = 60;
ws3.getCell("A4").value = "All 7pt (reference)";
ws3.getCell("B4").value =
  "The body text is 7pt and should wrap normally at its own size, fitting many more characters per line than it currently does.";
ws3.getCell("B4").font = { size: 7 };
ws3.getCell("B4").alignment = { wrapText: true };

// =============================================================================
// Sheet 4: Alignment with mixed-size rich text
//
// All runs share the same Y baseline (computed from the largest font's ascent).
// Vertical alignment (top/middle/bottom) positions the text block correctly
// within the cell, with line height based on the largest run's font size.
// =============================================================================

const ws4 = wb.addWorksheet("Alignment");
ws4.columns = [
  { header: "Vertical Align", width: 15 },
  { header: "Mixed Rich Text", width: 40 },
  { header: "Plain (reference)", width: 40 }
];

// Middle alignment — Y position depends on totalTextHeight which uses maxFontSize lineHeight
ws4.getRow(2).height = 40;
ws4.getCell("A2").value = "Middle";
ws4.getCell("B2").value = {
  richText: [
    { text: "BIG(16)", font: { size: 16 } },
    { text: " tiny(7)", font: { size: 7 } }
  ]
};
ws4.getCell("B2").alignment = { horizontal: "left", vertical: "middle" };
ws4.getCell("C2").value = "Reference: 11pt middle";
ws4.getCell("C2").alignment = { horizontal: "left", vertical: "middle" };

// Bottom alignment
ws4.getRow(3).height = 40;
ws4.getCell("A3").value = "Bottom";
ws4.getCell("B3").value = {
  richText: [
    { text: "BIG(16)", font: { size: 16 } },
    { text: " tiny(7)", font: { size: 7 } }
  ]
};
ws4.getCell("B3").alignment = { horizontal: "left", vertical: "bottom" };
ws4.getCell("C3").value = "Reference: 11pt bottom";
ws4.getCell("C3").alignment = { horizontal: "left", vertical: "bottom" };

// Center horizontal with mixed sizes
ws4.getRow(4).height = 30;
ws4.getCell("A4").value = "H-Center";
ws4.getCell("B4").value = {
  richText: [
    { text: "Big(14)", font: { size: 14 } },
    { text: " Small(8)", font: { size: 8 } }
  ]
};
ws4.getCell("B4").alignment = { horizontal: "center", vertical: "middle" };
ws4.getCell("C4").value = "Reference: center";
ws4.getCell("C4").alignment = { horizontal: "center", vertical: "middle" };

// =============================================================================
// Export
// =============================================================================

const pdf = await excelToPdf(wb, {
  showGridLines: true,
  showSheetNames: true,
  showPageNumbers: true,
  title: "Rich Text Overflow Demo"
});

const filename = "rich-text-overflow.pdf";
fs.writeFileSync(path.join(outDir, filename), pdf);
console.log(`${filename} generated — ${pdf.length} bytes`);
console.log(`Output: ${path.join(outDir, filename)}`);

// Also export as xlsx for comparison in Excel
await wb.xlsx.writeFile(path.join(outDir, "rich-text-overflow.xlsx"));
console.log("rich-text-overflow.xlsx generated for comparison in Excel");
