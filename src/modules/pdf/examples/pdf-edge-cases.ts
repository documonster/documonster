/**
 * Example: PDF Edge Cases
 *
 * Demonstrates advanced rendering scenarios not covered by the basic styled example.
 *
 * Covers:
 * - Type-based default alignment (numbers right, booleans center, formula result inference)
 * - Merged cell boundary borders (different colors per edge, double border)
 * - Text overflow into adjacent empty cells
 * - Zero-value number formats (#.##, ??0.00, accounting "-"??)
 * - fitToPage scaling with many columns
 * - Long word overflow vs wrap
 * - Custom vs auto row heights
 * - Error values and empty styled rows
 * - Explicit newlines in non-wrapped cells
 * - Rich text with mixed styles
 *
 * Run: npx tsx src/modules/pdf/examples/pdf-edge-cases.ts
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook, excelToPdf } from "../../../index";
import { pdf } from "../pdf";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/pdf-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const wb = new Workbook();

// =============================================================================
// Sheet 1: Default Alignment by Type
// =============================================================================

const wsAlign = wb.addWorksheet("Default Alignment");
wsAlign.columns = [
  { header: "Type", width: 15 },
  { header: "Value", width: 20 },
  { header: "Expected", width: 15 }
];
wsAlign.getCell("A2").value = "Text";
wsAlign.getCell("B2").value = "Hello World";
wsAlign.getCell("C2").value = "LEFT";

wsAlign.getCell("A3").value = "Number";
wsAlign.getCell("B3").value = 1234.56;
wsAlign.getCell("C3").value = "RIGHT";

wsAlign.getCell("A4").value = "Date";
wsAlign.getCell("B4").value = new Date(2025, 0, 15);
wsAlign.getCell("B4").numFmt = "m/d/yyyy";
wsAlign.getCell("C4").value = "RIGHT";

wsAlign.getCell("A5").value = "Boolean";
wsAlign.getCell("B5").value = true;
wsAlign.getCell("C5").value = "CENTER";

wsAlign.getCell("A6").value = "Num formula";
wsAlign.getCell("B6").value = { formula: "1+1", result: 2 };
wsAlign.getCell("C6").value = "RIGHT";

wsAlign.getCell("A7").value = "Str formula";
wsAlign.getCell("B7").value = { formula: 'CONCAT("a","b")', result: "ab" };
wsAlign.getCell("C7").value = "LEFT";

wsAlign.getCell("A8").value = "Bool formula";
wsAlign.getCell("B8").value = { formula: "TRUE()", result: true };
wsAlign.getCell("C8").value = "CENTER";

// =============================================================================
// Sheet 2: Merge Borders + Double Border
// =============================================================================

const wsMerge = wb.addWorksheet("Merge Borders");
wsMerge.columns = [{ width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }];

// 4-column merge with different colored border per edge
wsMerge.mergeCells("A1:D3");
wsMerge.getCell("A1").value = "Merged A1:D3";
wsMerge.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
wsMerge.getCell("A1").border = {
  top: { style: "thick", color: { argb: "FFFF0000" } },
  left: { style: "thick", color: { argb: "FF00CC00" } }
};
wsMerge.getCell("D1").border = { right: { style: "thick", color: { argb: "FF0000FF" } } };
wsMerge.getCell("D2").border = { right: { style: "thick", color: { argb: "FF0000FF" } } };
wsMerge.getCell("D3").border = {
  right: { style: "thick", color: { argb: "FF0000FF" } },
  bottom: { style: "double", color: { argb: "FFFF00FF" } }
};
wsMerge.getCell("A3").border = { bottom: { style: "double", color: { argb: "FFFF00FF" } } };
wsMerge.getCell("B3").border = { bottom: { style: "double", color: { argb: "FFFF00FF" } } };
wsMerge.getCell("C3").border = { bottom: { style: "double", color: { argb: "FFFF00FF" } } };

// Bordered empty cells
wsMerge.getCell("A5").value = "Data";
wsMerge.getCell("B5").border = {
  top: { style: "thick", color: { argb: "FFFF0000" } },
  right: { style: "thick", color: { argb: "FFFF0000" } },
  bottom: { style: "thick", color: { argb: "FFFF0000" } },
  left: { style: "thick", color: { argb: "FFFF0000" } }
};
wsMerge.getCell("C5").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFFFF00" }
};
wsMerge.getCell("C5").border = {
  top: { style: "thin" },
  right: { style: "thin" },
  bottom: { style: "thin" },
  left: { style: "thin" }
};

// =============================================================================
// Sheet 3: Text Overflow + Long Word + Newlines
// =============================================================================

const wsOverflow = wb.addWorksheet("Overflow & Wrap");
wsOverflow.columns = [{ width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }];

// Overflow into empty neighbors
wsOverflow.getCell("A1").value = "This text overflows into B1 and C1";
wsOverflow.getCell("D1").value = "Blocker";

// Long single word
wsOverflow.getCell("A3").value = "Supercalifragilisticexpialidocious";
wsOverflow.getCell("D3").value = "Stop";

// Same word but wrapped
wsOverflow.getCell("A5").value = "Supercalifragilisticexpialidocious";
wsOverflow.getCell("A5").alignment = { wrapText: true };

// Explicit newlines without wrapText
wsOverflow.getCell("C5").value = "Line1\nLine2\nLine3";

// Explicit newlines with wrapText
wsOverflow.getCell("A7").value = "Wrap\nwith\nnewlines";
wsOverflow.getCell("A7").alignment = { wrapText: true };

// =============================================================================
// Sheet 4: Zero-Value Number Formats
// =============================================================================

const wsZero = wb.addWorksheet("Zero Formats");
wsZero.columns = [
  { header: "Format", width: 30 },
  { header: "Value=0", width: 20 }
];

const zeroFormats: [string, string][] = [
  ['#,##0.00;-#,##0.00;"-"??', 'Dash + 2 spaces: "-  "'],
  ["??0.00", 'Space-padded: "  0.00"'],
  ["#.##", "Empty string"],
  ['0.00;-0.00;"-"', 'Just a dash: "-"'],
  ["#,##0", '"0"'],
  ["???.???", '"   .   " (spaces + dot + spaces)'],
  ["0.##", '"0" (no trailing decimals)'],
  ["#.0#", '".0"'],
  ["0.0?", '"0.0 " (trailing space)']
];
for (const [fmt, desc] of zeroFormats) {
  const row = wsZero.addRow([`${fmt}  →  ${desc}`, 0]);
  row.getCell(2).numFmt = fmt;
}

// =============================================================================
// Sheet 5: Row Heights + fitToPage
// =============================================================================

const wsHeight = wb.addWorksheet("Row Heights");
wsHeight.getColumn(1).width = 30;

wsHeight.getCell("A1").value = "Auto height (default)";
const r2 = wsHeight.getRow(2);
r2.height = 40;
wsHeight.getCell("A2").value = "Custom height = 40pt";
wsHeight.getCell("A3").value = "Tall font (auto expand)";
wsHeight.getCell("A3").font = { size: 24 };
const r4 = wsHeight.getRow(4);
r4.height = 10;
wsHeight.getCell("A4").value = "Cramped: height=10pt";

// =============================================================================
// Sheet 6: Error Value + Empty Styled Row + Hyperlink
// =============================================================================

const wsMixed = wb.addWorksheet("Mixed Types");
wsMixed.getColumn(1).width = 30;

wsMixed.getCell("A1").value = { text: "Click me (hyperlink)", hyperlink: "https://example.com" };
wsMixed.getCell("A1").font = { color: { argb: "FF0563C1" }, underline: true };

wsMixed.getCell("A2").value = { error: "#DIV/0!" } as any;

// Empty row with styling
const r3 = wsMixed.getRow(3);
r3.font = { bold: true, size: 14 };
r3.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFCCCC" } };

wsMixed.getCell("A4").value = "Normal cell after styled empty row";

// =============================================================================
// Export: multi-sheet workbook
// =============================================================================

const mainPdf = await excelToPdf(wb, {
  showGridLines: true,
  showSheetNames: true,
  title: "PDF Edge Cases",
  author: "excelts"
});
fs.writeFileSync(path.join(outDir, "pdf-edge-cases.pdf"), mainPdf);
console.log("pdf-edge-cases.pdf — 6 sheets covering advanced rendering edge cases");

// =============================================================================
// Separate file: fitToPage with 20 columns
// =============================================================================

const wbFit = new Workbook();
const wsFit = wbFit.addWorksheet("20 Columns");
for (let c = 1; c <= 20; c++) {
  wsFit.getColumn(c).width = 12;
  wsFit.getCell(1, c).value = `Col ${c}`;
  wsFit.getCell(1, c).font = { bold: true };
  wsFit.getCell(2, c).value = c * 100;
}
const fitPdf = await excelToPdf(wbFit, {
  showGridLines: true,
  fitToPage: true,
  title: "Fit To Page: 20 Columns"
});
fs.writeFileSync(path.join(outDir, "pdf-fit-to-page.pdf"), fitPdf);
console.log("pdf-fit-to-page.pdf — 20 columns shrunk to fit one page");

// =============================================================================
// Separate file: standalone pdf() with type-based alignment
// =============================================================================

const standalonePdf = await pdf(
  {
    sheets: [
      {
        name: "Standalone Alignment",
        data: [
          ["Text", "Number", "Boolean"],
          ["hello", 123.45, true],
          ["world", -67.89, false],
          ["test", 0, true]
        ]
      }
    ]
  },
  { showGridLines: true, title: "Standalone PDF: Type Alignment" }
);
fs.writeFileSync(path.join(outDir, "pdf-standalone-align.pdf"), standalonePdf);
console.log("pdf-standalone-align.pdf — standalone pdf() with type-based alignment");

console.log(`\nAll files written to: ${outDir}`);
