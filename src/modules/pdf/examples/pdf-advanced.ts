/**
 * Example: Advanced PDF Export
 *
 * Covers:
 * - Multi-page pagination (many rows)
 * - Repeat header rows (printTitlesRow)
 * - Manual row page breaks
 * - Print area
 * - Page setup from worksheet (paperSize, orientation, margins)
 * - Column pagination (wide sheets with horizontal splitting)
 * - Hidden rows and columns (skipped in PDF)
 * - Password-protected / encrypted PDF
 * - Custom TrueType font embedding (Unicode/CJK support)
 * - Transparency / alpha colors
 * - Multiple worksheets with different page setups
 * - Bookmarks / outlines (auto-generated per sheet)
 * - Custom grid line color
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

// =============================================================================
// 1. Pagination with repeat header rows
// =============================================================================

const wb1 = new Workbook();
const ws1 = wb1.addWorksheet("Inventory");
ws1.columns = [
  { header: "ID", key: "id", width: 8 },
  { header: "Product Name", key: "name", width: 25 },
  { header: "Category", key: "category", width: 15 },
  { header: "Stock", key: "stock", width: 10 },
  { header: "Price", key: "price", width: 12 }
];

// Style the header row
ws1.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
ws1.getRow(1).fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF2F5496" }
};
ws1.getRow(1).alignment = { horizontal: "center" };

// Add 100 rows to force pagination
const categories = ["Electronics", "Clothing", "Food", "Tools", "Books"];
for (let i = 1; i <= 100; i++) {
  ws1.addRow({
    id: i,
    name: `Product ${i}`,
    category: categories[i % categories.length],
    stock: Math.floor(Math.random() * 500),
    price: Math.round(Math.random() * 10000) / 100
  });
}

// Format the price column
ws1.getColumn("price").numFmt = "$#,##0.00";

// Repeat row 1 on every page
ws1.pageSetup.printTitlesRow = "1:1";

const pdf1 = await excelToPdf(wb1, {
  showGridLines: true,
  showPageNumbers: true,
  showSheetNames: true,
  repeatRows: 1,
  title: "Inventory Report"
});
fs.writeFileSync(path.join(outDir, "advanced-pagination.pdf"), pdf1);
console.log("1. advanced-pagination.pdf — 100 rows, repeat headers, page numbers");

// =============================================================================
// 2. Manual row page breaks + print area
// =============================================================================

const wb2 = new Workbook();
const ws2 = wb2.addWorksheet("Sections");
for (let r = 1; r <= 30; r++) {
  ws2.getCell(`A${r}`).value = `Section ${Math.ceil(r / 10)}`;
  ws2.getCell(`B${r}`).value = `Row ${r} data`;
  ws2.getCell(`C${r}`).value = r * 100;
}
ws2.getColumn(1).width = 15;
ws2.getColumn(2).width = 20;
ws2.getColumn(3).width = 10;

// Break after row 10 and row 20
ws2.getRow(10).addPageBreak();
ws2.getRow(20).addPageBreak();

// Print area: only columns A-C, rows 1-30
ws2.pageSetup.printArea = "A1:C30";

const pdf2 = await excelToPdf(wb2, {
  showGridLines: true,
  showPageNumbers: true
});
fs.writeFileSync(path.join(outDir, "advanced-pagebreaks.pdf"), pdf2);
console.log("2. advanced-pagebreaks.pdf — manual page breaks + print area");

// =============================================================================
// 3. Wide sheet with column pagination
// =============================================================================

const wb3 = new Workbook();
const ws3 = wb3.addWorksheet("Wide Data");
// 20 columns × 30 rows — will need horizontal page splits
for (let c = 1; c <= 20; c++) {
  ws3.getColumn(c).width = 12;
  ws3.getCell(1, c).value = `Col ${c}`;
  ws3.getCell(1, c).font = { bold: true };
}
for (let r = 2; r <= 30; r++) {
  for (let c = 1; c <= 20; c++) {
    ws3.getCell(r, c).value = (r - 1) * c;
  }
}

const pdf3 = await excelToPdf(wb3, {
  fitToPage: false, // Don't shrink — let it paginate horizontally
  showGridLines: true,
  showPageNumbers: true,
  showSheetNames: true
});
fs.writeFileSync(path.join(outDir, "advanced-wide.pdf"), pdf3);
console.log("3. advanced-wide.pdf — 20 columns, horizontal pagination");

// =============================================================================
// 4. Hidden rows and columns
// =============================================================================

const wb4 = new Workbook();
const ws4 = wb4.addWorksheet("Hidden");
for (let r = 1; r <= 10; r++) {
  ws4.getCell(`A${r}`).value = `Visible A${r}`;
  ws4.getCell(`B${r}`).value = `Hidden B${r}`;
  ws4.getCell(`C${r}`).value = `Visible C${r}`;
}
ws4.getColumn(2).hidden = true;
ws4.getRow(3).hidden = true;
ws4.getRow(7).hidden = true;

const pdf4 = await excelToPdf(wb4, { showGridLines: true });
fs.writeFileSync(path.join(outDir, "advanced-hidden.pdf"), pdf4);
console.log("4. advanced-hidden.pdf — hidden column B, hidden rows 3 and 7");

// =============================================================================
// 5. Encrypted PDF with permissions
// =============================================================================

const wb5 = new Workbook();
const ws5 = wb5.addWorksheet("Confidential");
ws5.getCell("A1").value = "This document is password-protected.";
ws5.getCell("A1").font = { bold: true, size: 14, color: { argb: "FFCC0000" } };
ws5.getCell("A3").value = "Owner password: owner123";
ws5.getCell("A4").value = "User password: (none — opens without password)";
ws5.getCell("A6").value = "Permissions: print=yes, copy=no, modify=no";
ws5.getColumn(1).width = 50;

const pdf5 = await excelToPdf(wb5, {
  encryption: {
    ownerPassword: "owner123",
    // No userPassword — document opens without a password
    permissions: {
      print: true,
      copy: false,
      modify: false,
      annotate: false
    }
  },
  title: "Confidential Document",
  author: "Security Team"
});
fs.writeFileSync(path.join(outDir, "advanced-encrypted.pdf"), pdf5);
console.log("5. advanced-encrypted.pdf — encrypted, print-only, no copy");

// =============================================================================
// 6. Encrypted PDF with open password
// =============================================================================

const wb6 = new Workbook();
const ws6 = wb6.addWorksheet("Locked");
ws6.getCell("A1").value = "You need a password to open this PDF.";
ws6.getCell("A2").value = 'User password: "hello"';
ws6.getColumn(1).width = 40;

const pdf6 = await excelToPdf(wb6, {
  encryption: {
    ownerPassword: "admin",
    userPassword: "hello"
  }
});
fs.writeFileSync(path.join(outDir, "advanced-password.pdf"), pdf6);
console.log('6. advanced-password.pdf — requires password "hello" to open');

// =============================================================================
// 7. Multiple sheets with different page setups + bookmarks
// =============================================================================

const wb7 = new Workbook();

const wsPortrait = wb7.addWorksheet("Portrait A4");
wsPortrait.getCell("A1").value = "Portrait A4 sheet";
wsPortrait.getCell("A1").font = { bold: true, size: 16 };
wsPortrait.pageSetup.paperSize = 9; // A4
wsPortrait.pageSetup.orientation = "portrait";

const wsLandscape = wb7.addWorksheet("Landscape A4");
wsLandscape.getCell("A1").value = "Landscape A4 sheet";
wsLandscape.getCell("A1").font = { bold: true, size: 16 };
wsLandscape.pageSetup.paperSize = 9;
wsLandscape.pageSetup.orientation = "landscape";

const wsLetter = wb7.addWorksheet("Letter");
wsLetter.getCell("A1").value = "US Letter sheet";
wsLetter.getCell("A1").font = { bold: true, size: 16 };
wsLetter.pageSetup.paperSize = 1; // Letter

const wsSmall = wb7.addWorksheet("A5");
wsSmall.getCell("A1").value = "A5 sheet";
wsSmall.getCell("A1").font = { bold: true, size: 16 };
wsSmall.pageSetup.paperSize = 11; // A5

// Export — each sheet gets its own page setup; bookmarks are auto-generated
const pdf7 = await excelToPdf(wb7, {
  showSheetNames: true,
  showPageNumbers: true
});
fs.writeFileSync(path.join(outDir, "advanced-multi-setup.pdf"), pdf7);
console.log("7. advanced-multi-setup.pdf — per-sheet page sizes + bookmarks");

// =============================================================================
// 8. Transparency / alpha fills
// =============================================================================

const wb8 = new Workbook();
const ws8 = wb8.addWorksheet("Transparency");
ws8.getColumn(1).width = 25;
ws8.getColumn(2).width = 25;

ws8.getCell("A1").value = "Opaque red fill";
ws8.getCell("A1").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFFF0000" }
};

ws8.getCell("A2").value = "Semi-transparent blue";
ws8.getCell("A2").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "800000FF" } // alpha=0x80 (~50%)
};

ws8.getCell("A3").value = "Light transparent green";
ws8.getCell("A3").fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "4000CC00" } // alpha=0x40 (~25%)
};

ws8.getCell("A4").value = "Semi-transparent text";
ws8.getCell("A4").font = { color: { argb: "80FF0000" }, size: 14, bold: true };

const pdf8 = await excelToPdf(wb8, { showGridLines: true });
fs.writeFileSync(path.join(outDir, "advanced-transparency.pdf"), pdf8);
console.log("8. advanced-transparency.pdf — alpha fills and text");

// =============================================================================
// 9. Custom grid line color
// =============================================================================

const wb9 = new Workbook();
const ws9 = wb9.addWorksheet("GridColors");
for (let r = 1; r <= 5; r++) {
  for (let c = 1; c <= 5; c++) {
    ws9.getCell(r, c).value = `${r},${c}`;
  }
}

const pdf9 = await excelToPdf(wb9, {
  showGridLines: true,
  gridLineColor: "FF3366CC" // blue grid lines
});
fs.writeFileSync(path.join(outDir, "advanced-grid-color.pdf"), pdf9);
console.log("9. advanced-grid-color.pdf — blue grid lines");

// =============================================================================
// 10. Worksheet margins from pageSetup
// =============================================================================

const wb10 = new Workbook();
const ws10 = wb10.addWorksheet("WS Margins");
ws10.getCell("A1").value = "This sheet uses worksheet-level margins (0.5in all sides)";
ws10.getColumn(1).width = 50;
ws10.pageSetup.margins = {
  left: 0.5,
  right: 0.5,
  top: 0.5,
  bottom: 0.5,
  header: 0.3,
  footer: 0.3
};

const pdf10 = await excelToPdf(wb10, { showPageNumbers: true });
fs.writeFileSync(path.join(outDir, "advanced-ws-margins.pdf"), pdf10);
console.log("10. advanced-ws-margins.pdf — worksheet pageSetup margins");

console.log("\nAll advanced examples generated.");
