/**
 * Example: Basic PDF Export
 *
 * Covers:
 * - Simple workbook creation with data
 * - excelToPdf() function
 * - Page size, orientation, margins
 * - Grid lines
 * - Sheet name headers and page number footers
 * - PDF metadata (title, author, subject)
 * - Selecting specific sheets
 * - fitToPage and scale
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
// 1. Minimal export — one line
// =============================================================================

const wb1 = new Workbook();
const ws1 = wb1.addWorksheet("Sales");
ws1.columns = [
  { header: "Product", key: "product", width: 20 },
  { header: "Q1", key: "q1", width: 12 },
  { header: "Q2", key: "q2", width: 12 },
  { header: "Q3", key: "q3", width: 12 },
  { header: "Q4", key: "q4", width: 12 }
];
ws1.addRows([
  { product: "Widget A", q1: 1200, q2: 1350, q3: 1100, q4: 1500 },
  { product: "Widget B", q1: 800, q2: 950, q3: 1020, q4: 870 },
  { product: "Gadget C", q1: 3200, q2: 2900, q3: 3100, q4: 3400 },
  { product: "Gadget D", q1: 450, q2: 520, q3: 480, q4: 610 }
]);

// Simplest possible export
const pdf1 = await excelToPdf(wb1);
fs.writeFileSync(path.join(outDir, "basic-minimal.pdf"), pdf1);
console.log("1. basic-minimal.pdf — default settings");

// =============================================================================
// 2. With grid lines, headers, footers, and metadata
// =============================================================================

const pdf2 = await excelToPdf(wb1, {
  showGridLines: true,
  showSheetNames: true,
  showPageNumbers: true,
  title: "Quarterly Sales Report",
  author: "Finance Team",
  subject: "Q1-Q4 Revenue"
});
fs.writeFileSync(path.join(outDir, "basic-gridlines.pdf"), pdf2);
console.log("2. basic-gridlines.pdf — grid lines + headers + footers + metadata");

// =============================================================================
// 3. Landscape A3 with custom margins
// =============================================================================

const pdf3 = await excelToPdf(wb1, {
  pageSize: "A3",
  orientation: "landscape",
  margins: { top: 36, right: 36, bottom: 36, left: 36 },
  showGridLines: true
});
fs.writeFileSync(path.join(outDir, "basic-landscape-a3.pdf"), pdf3);
console.log("3. basic-landscape-a3.pdf — landscape A3, tight margins");

// =============================================================================
// 4. Custom page size (US Half Letter)
// =============================================================================

const pdf4 = await excelToPdf(wb1, {
  pageSize: { width: 396, height: 612 }, // 5.5" × 8.5" in points
  fitToPage: true,
  scale: 0.9
});
fs.writeFileSync(path.join(outDir, "basic-custom-size.pdf"), pdf4);
console.log("4. basic-custom-size.pdf — custom page size, scale 0.9");

// =============================================================================
// 5. Multi-sheet workbook, export specific sheets
// =============================================================================

const wb5 = new Workbook();
const wsJan = wb5.addWorksheet("January");
wsJan.getColumn(1).width = 20;
wsJan.getCell("A1").value = "January Data";
const wsFeb = wb5.addWorksheet("February");
wsFeb.getColumn(1).width = 20;
wsFeb.getCell("A1").value = "February Data";
const wsMar = wb5.addWorksheet("March");
wsMar.getColumn(1).width = 20;
wsMar.getCell("A1").value = "March Data";
const wsSum = wb5.addWorksheet("Summary");
wsSum.getColumn(1).width = 20;
wsSum.getCell("A1").value = "Q1 Summary";

// Export only January and Summary by name
const pdf5a = await excelToPdf(wb5, { sheets: ["January", "Summary"] });
fs.writeFileSync(path.join(outDir, "basic-select-by-name.pdf"), pdf5a);

// Export by 1-based position (sheets 2 and 4)
const pdf5b = await excelToPdf(wb5, { sheets: [2, 4] });
fs.writeFileSync(path.join(outDir, "basic-select-by-index.pdf"), pdf5b);
console.log("5. basic-select-by-name.pdf + basic-select-by-index.pdf — sheet selection");

// =============================================================================
// 6. excelToPdf with options
// =============================================================================

const pdf6 = await excelToPdf(wb1, {
  showGridLines: true,
  showPageNumbers: true,
  title: "Via excelToPdf"
});
fs.writeFileSync(path.join(outDir, "basic-with-options.pdf"), pdf6);
console.log("6. basic-with-options.pdf — excelToPdf with options");

console.log("\nAll basic examples generated.");
