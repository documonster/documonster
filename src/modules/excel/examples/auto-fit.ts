/**
 * Auto-Fit Example: Demonstrates autoFitColumns() and autoFitRows()
 *
 * Generates 3 xlsx files showcasing different auto-fit scenarios:
 * 1. auto-fit-basic.xlsx      — Mixed content types with autoFitColumns
 * 2. auto-fit-fonts.xlsx      — Different fonts and sizes
 * 3. auto-fit-advanced.xlsx   — Wrap text, multi-line, rich text, merged cells, indent
 */
import { Workbook } from "../../../index";

const outDir = "out";

// =============================================================================
// 1. Basic auto-fit: mixed content types
// =============================================================================
async function generateBasic() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Auto-Fit Basic");

  // Header row (bold)
  const headers = ["ID", "Name", "Description", "Amount", "Date", "Active"];
  const headerRow = ws.addRow(headers);
  headerRow.font = { bold: true };

  // Data rows
  ws.addRow([1, "Alice", "Software Engineer", 95000.5, new Date(2024, 0, 15), true]);
  ws.addRow([
    2,
    "Bob",
    "Senior Product Manager with a very long title",
    120000,
    new Date(2024, 5, 1),
    false
  ]);
  ws.addRow([3, "Charlie", "QA", 78000.99, new Date(2023, 11, 20), true]);
  ws.addRow([4, "Diana", "VP of Engineering and Operations", 185000, new Date(2024, 2, 10), true]);
  ws.addRow([5, "Eve", "Intern", 45000, new Date(2024, 8, 1), false]);

  // Apply number format to Amount column
  ws.getColumn("D").numFmt = "#,##0.00";
  // Apply date format
  ws.getColumn("E").numFmt = "yyyy-mm-dd";

  // Auto-fit all columns then all rows
  ws.autoFitColumns().autoFitRows();

  const path = `${outDir}/auto-fit-basic.xlsx`;
  await wb.xlsx.writeFile(path);
  console.log(`Written: ${path}`);
}

// =============================================================================
// 2. Different fonts and sizes
// =============================================================================
async function generateFonts() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Auto-Fit Fonts");

  ws.getCell("A1").value = "Calibri 11 (default)";

  ws.getCell("A2").value = "Calibri 16 Bold";
  ws.getCell("A2").font = { name: "Calibri", size: 16, bold: true };

  ws.getCell("A3").value = "Arial 14";
  ws.getCell("A3").font = { name: "Arial", size: 14 };

  ws.getCell("A4").value = "Arial 11 Bold";
  ws.getCell("A4").font = { name: "Arial", size: 11, bold: true };

  ws.getCell("A5").value = "Times New Roman 12";
  ws.getCell("A5").font = { name: "Times New Roman", size: 12 };

  ws.getCell("A6").value = "Courier New 11 (monospace)";
  ws.getCell("A6").font = { name: "Courier New", size: 11 };

  ws.getCell("A7").value = "Verdana 10";
  ws.getCell("A7").font = { name: "Verdana", size: 10 };

  ws.getCell("A8").value = "Tahoma 13 Italic";
  ws.getCell("A8").font = { name: "Tahoma", size: 13, italic: true };

  ws.getCell("A9").value = "Georgia 14";
  ws.getCell("A9").font = { name: "Georgia", size: 14 };

  ws.getCell("A10").value = "Trebuchet MS 11";
  ws.getCell("A10").font = { name: "Trebuchet MS", size: 11 };

  // Second column: numbers
  ws.getCell("B1").value = "Numbers";
  ws.getCell("B1").font = { bold: true };
  for (let i = 2; i <= 10; i++) {
    ws.getCell(`B${i}`).value = 1234567.89;
    ws.getCell(`B${i}`).numFmt = "#,##0.00";
    ws.getCell(`B${i}`).font = ws.getCell(`A${i}`).font;
  }

  ws.autoFitColumns().autoFitRows();

  const path = `${outDir}/auto-fit-fonts.xlsx`;
  await wb.xlsx.writeFile(path);
  console.log(`Written: ${path}`);
}

// =============================================================================
// 3. Advanced: wrap text, multi-line, rich text, merged, indent, hidden
// =============================================================================
async function generateAdvanced() {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Auto-Fit Advanced");

  // Column A: Labels
  ws.getCell("A1").value = "Feature";
  ws.getCell("A1").font = { bold: true };
  ws.getCell("B1").value = "Content";
  ws.getCell("B1").font = { bold: true };
  ws.getCell("C1").value = "Notes";
  ws.getCell("C1").font = { bold: true };

  // Row 2: Multi-line explicit newlines
  ws.getCell("A2").value = "Multi-line (\\n)";
  ws.getCell("B2").value = "Line 1\nLine 2\nLine 3";
  ws.getCell("C2").value = "Height should accommodate 3 lines";

  // Row 3: Wrap text
  ws.getCell("A3").value = "Wrap Text";
  ws.getCell("B3").value =
    "This is a long sentence that should wrap within the column width when wrapText is enabled.";
  ws.getCell("B3").alignment = { wrapText: true };
  ws.getCell("C3").value = "Column B is 30 chars wide";

  // Row 4: Rich text
  ws.getCell("A4").value = "Rich Text";
  ws.getCell("B4").value = {
    richText: [
      { text: "Bold ", font: { bold: true, size: 14 } },
      { text: "Normal ", font: { size: 11 } },
      { text: "Red Italic", font: { italic: true, color: { argb: "FFFF0000" } } }
    ]
  };
  ws.getCell("C4").value = "Mixed fonts in one cell";

  // Row 5: Indent
  ws.getCell("A5").value = "Indent";
  ws.getCell("B5").value = "Indented text (level 3)";
  ws.getCell("B5").alignment = { indent: 3 };
  ws.getCell("C5").value = "Extra width for indent";

  // Row 6: Merged cells (should be skipped for width)
  ws.getCell("A6").value = "Merged";
  ws.getCell("B6").value = "This cell is merged across B6:C6";
  ws.mergeCells("B6:C6");

  // Row 7: ShrinkToFit (should be skipped for width)
  ws.getCell("A7").value = "ShrinkToFit";
  ws.getCell("B7").value = "This long text has shrinkToFit — should NOT affect column width";
  ws.getCell("B7").alignment = { shrinkToFit: true };
  ws.getCell("C7").value = "Skipped by auto-fit";

  // Row 8: Hidden row (should be skipped for width)
  ws.getCell("A8").value = "Hidden Row";
  ws.getCell("B8").value =
    "VERY LONG TEXT THAT IS HIDDEN AND SHOULD NOT AFFECT COLUMN WIDTH AT ALL";
  ws.getCell("C8").value = "This row is hidden";
  ws.getRow(8).hidden = true;

  // Row 9: Large font
  ws.getCell("A9").value = "Large Font";
  ws.getCell("B9").value = "Big Text";
  ws.getCell("B9").font = { size: 24, bold: true };
  ws.getCell("C9").value = "24pt bold";

  // Row 10: Numbers with format
  ws.getCell("A10").value = "Formatted Number";
  ws.getCell("B10").value = 9876543.21;
  ws.getCell("B10").numFmt = "$#,##0.00";
  ws.getCell("C10").value = "Currency format";

  // Set column B to a reasonable width for wrap text demo
  ws.getColumn("B").width = 30;

  // Auto-fit columns A and C (skip B since we set it manually for wrap demo)
  ws.autoFitColumn("A");
  ws.autoFitColumn("C");

  // Auto-fit all rows (this will use column B's width for wrap calculation)
  ws.autoFitRows();

  const path = `${outDir}/auto-fit-advanced.xlsx`;
  await wb.xlsx.writeFile(path);
  console.log(`Written: ${path}`);
}

// =============================================================================
// Run all
// =============================================================================
await generateBasic();
await generateFonts();
await generateAdvanced();
console.log("Done!");
