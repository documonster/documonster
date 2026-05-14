/**
 * Word Example 27 — Excel ↔ DOCX
 *
 * Covers:
 *   - excelToDocx: render every visible sheet as a Word table with the
 *     original cell formatting preserved (bold, italic, fonts, colours,
 *     alignment).
 *   - Selecting specific sheets, capping rows/columns, including the
 *     workbook title page, suppressing borders.
 *   - extractTablesToExcel: pull tabular data out of a DOCX and feed it
 *     back into a Workbook (round-trip).
 *
 * Output: tmp/word-examples/27-excel/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Workbook } from "../../../index";
import { excelToDocx, extractTablesToExcel } from "../excel";
import { Document, toBuffer, readDocx } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/27-excel"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Build a small workbook with formatting
// ---------------------------------------------------------------------------
const wb = new Workbook();
wb.creator = "OpenCode";
wb.created = new Date("2026-05-01");

const ws1 = wb.addWorksheet("Sales");
ws1.columns = [
  { header: "Region", key: "region", width: 12 },
  { header: "Q1", key: "q1", width: 10 },
  { header: "Q2", key: "q2", width: 10 },
  { header: "Q3", key: "q3", width: 10 },
  { header: "Q4", key: "q4", width: 10 }
];
// Style header row
ws1.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
ws1.getRow(1).fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF1F4E79" }
};
const data = [
  { region: "North", q1: 1200, q2: 1400, q3: 1600, q4: 1900 },
  { region: "South", q1: 900, q2: 1100, q3: 1300, q4: 1500 },
  { region: "East", q1: 2000, q2: 2100, q3: 2200, q4: 2400 },
  { region: "West", q1: 750, q2: 820, q3: 950, q4: 1100 }
];
for (const r of data) {
  ws1.addRow(r);
}
ws1.getColumn("q1").numFmt = "$#,##0";
ws1.getColumn("q2").numFmt = "$#,##0";
ws1.getColumn("q3").numFmt = "$#,##0";
ws1.getColumn("q4").numFmt = "$#,##0";

const ws2 = wb.addWorksheet("Inventory");
ws2.addRow(["Item", "Qty", "Note"]);
ws2.addRow(["Widget", 100, "in stock"]);
ws2.addRow(["Gadget", 0, "OUT OF STOCK"]);
ws2.getCell("C3").font = { color: { argb: "FFC00000" }, bold: true };

// Add a hidden sheet — by default excelToDocx skips it
const ws3 = wb.addWorksheet("Internal");
ws3.state = "hidden";
ws3.addRow(["secret", 42]);

// ---------------------------------------------------------------------------
// 2. Convert — default: every visible sheet, headings included
// ---------------------------------------------------------------------------
const docModel = excelToDocx(wb);
fs.writeFileSync(path.join(outDir, "01-default.docx"), await toBuffer(docModel));
console.log(`  → 01-default.docx (${docModel.body.length} body items)`);

// ---------------------------------------------------------------------------
// 3. Custom options — only the Sales sheet, with title page, no borders
// ---------------------------------------------------------------------------
const docModel2 = excelToDocx(wb, {
  sheets: ["Sales"],
  includeSheetHeadings: false,
  includeTitlePage: true,
  includeBorders: false
});
fs.writeFileSync(path.join(outDir, "02-sales-only.docx"), await toBuffer(docModel2));
console.log(`  → 02-sales-only.docx`);

// ---------------------------------------------------------------------------
// 4. Cap rows and columns — useful for previewing huge workbooks
// ---------------------------------------------------------------------------
const docModel3 = excelToDocx(wb, { maxRows: 3, maxColumns: 3 });
fs.writeFileSync(path.join(outDir, "03-capped.docx"), await toBuffer(docModel3));
console.log(`  → 03-capped.docx`);

// ---------------------------------------------------------------------------
// 5. Round-trip: take a hand-built Word document with tables and pull the
//    table data back into Excel via extractTablesToExcel.
// ---------------------------------------------------------------------------
const wordDoc = Document.create();
Document.useDefaultStyles(wordDoc);
Document.addHeading(wordDoc, "Quarterly results", 1);
Document.addTable(
  wordDoc,
  [
    ["Quarter", "Revenue", "Profit"],
    ["Q1", "1.2", "0.2"],
    ["Q2", "1.5", "0.3"],
    ["Q3", "1.8", "0.4"]
  ],
  { headerRow: true, borders: true }
);
Document.addParagraph(wordDoc, "");
Document.addTable(
  wordDoc,
  [
    ["Region", "Sales"],
    ["North", "120"],
    ["South", "90"]
  ],
  { headerRow: true, borders: true }
);

const wordBuf = await toBuffer(Document.build(wordDoc));
fs.writeFileSync(path.join(outDir, "04-source.docx"), wordBuf);
const reread = await readDocx(wordBuf);

const tables = extractTablesToExcel(reread);
console.log(`  extracted ${tables.length} tables from DOCX:`);
for (const t of tables) {
  console.log(`    - ${t.name}: ${t.data.length} rows`);
}

// Fold each extracted table into a sheet of a new workbook
const wb2 = new Workbook();
for (const tbl of tables) {
  const sheet = wb2.addWorksheet(tbl.name);
  for (const row of tbl.data) {
    sheet.addRow(row);
  }
}
await wb2.xlsx.writeFile(path.join(outDir, "05-extracted.xlsx"));
console.log(`  → 05-extracted.xlsx`);
