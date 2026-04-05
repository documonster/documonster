/**
 * Example: Round-Trip & Workbook Integration
 *
 * Covers:
 * - format → parse → format round-trip fidelity
 * - Workbook.readMarkdown / writeMarkdown
 * - Workbook.readMarkdownAll (multi-table → worksheets)
 * - Workbook.writeMarkdownBuffer
 * - Value mapper on readMarkdown
 * - Multiline round-trip through Workbook
 * - Error handling with MarkdownParseError
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseMarkdown, formatMarkdown, MarkdownParseError } from "../index";
import { Workbook } from "../../excel/workbook";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/md-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// =============================================================================
// 1. Round-trip fidelity
// =============================================================================

console.log("=== 1. Round-Trip ===\n");

const headers = ["Name", "Formula", "Path"];
const rows: string[][] = [
  ["Alice", "a | b", "C:\\Users"],
  ["Bob", "x > y", "D:\\Data"]
];

const md1 = formatMarkdown(headers, rows, {
  columns: [
    { header: "Name", alignment: "left" },
    { header: "Formula", alignment: "center" },
    { header: "Path", alignment: "right" }
  ]
});
console.log("Formatted:");
console.log(md1);

const parsed = parseMarkdown(md1);
const md2 = formatMarkdown(parsed.headers, parsed.rows, {
  columns: parsed.headers.map((h, i) => ({
    header: h,
    alignment: parsed.alignments[i]
  }))
});

console.log("Round-trip match:", md1 === md2 ? "PASS" : "FAIL");
fs.writeFileSync(path.join(outDir, "round-trip.md"), md1, "utf8");
console.log();

// =============================================================================
// 2. Workbook.readMarkdown / writeMarkdown
// =============================================================================

console.log("=== 2. Workbook readMarkdown / writeMarkdown ===\n");

const mdInput = `
| Name  | Age | Department  |
| :---- | --: | :---------: |
| Alice | 30  | Engineering |
| Bob   | 25  | Marketing   |
| Carol | 35  | Engineering |
`.trim();

const wb1 = new Workbook();
const ws = wb1.readMarkdown(mdInput, { sheetName: "Employees" });

console.log("Sheet:", ws.name, "| Rows:", ws.rowCount);
for (let r = 1; r <= ws.rowCount; r++) {
  const row = ws.getRow(r);
  console.log(
    `  Row ${r}: ${row.getCell(1).value} | ${row.getCell(2).value} | ${row.getCell(3).value}`
  );
}

const mdOut = wb1.writeMarkdown({ sheetName: "Employees" });
console.log("\nWritten back:");
console.log(mdOut);
fs.writeFileSync(path.join(outDir, "workbook-single.md"), mdOut, "utf8");

// =============================================================================
// 3. Workbook.readMarkdownAll
// =============================================================================

console.log("=== 3. readMarkdownAll ===\n");

const multiDoc = `
# Q1 Report

| Product | Revenue |
| :------ | ------: |
| Widget  | $10,000 |
| Gadget  | $25,000 |

| Category  | Amount  |
| :-------- | ------: |
| Salaries  | $50,000 |
| Rent      | $5,000  |

| Metric  | Value    |
| :------ | -------: |
| Revenue | $35,000  |
| Net     | -$28,000 |
`.trim();

const wb2 = new Workbook();
const sheets = wb2.readMarkdownAll(multiDoc, { sheetName: "Q1" });

console.log(`Created ${sheets.length} worksheets:`);
const allMd: string[] = [];
for (const s of sheets) {
  const md = wb2.writeMarkdown({ sheetName: s.name });
  console.log(`\n--- ${s.name} ---`);
  console.log(md);
  allMd.push(`## ${s.name}\n\n${md}`);
}
fs.writeFileSync(
  path.join(outDir, "workbook-multi.md"),
  "# Multi-Table Workbook\n\n" + allMd.join("\n"),
  "utf8"
);

// =============================================================================
// 4. writeMarkdownBuffer
// =============================================================================

console.log("=== 4. writeMarkdownBuffer ===\n");
const buffer = wb2.writeMarkdownBuffer({ sheetName: "Q1" });
console.log(`Buffer: ${buffer.length} bytes`);
fs.writeFileSync(path.join(outDir, "workbook-buffer.md"), buffer);
console.log();

// =============================================================================
// 5. Value mapper
// =============================================================================

console.log("=== 5. Value Mapper ===\n");

const wb3 = new Workbook();
const ws3 = wb3.readMarkdown(mdInput, {
  sheetName: "Typed",
  map: (value: string, _col: number) => {
    const n = Number(value);
    return Number.isNaN(n) ? value : n;
  }
});
console.log("Age type:", typeof ws3.getRow(2).getCell(2).value);
console.log("Age value:", ws3.getRow(2).getCell(2).value);
console.log();

// =============================================================================
// 6. Multiline round-trip through Workbook
// =============================================================================

console.log("=== 6. Multiline Workbook Round-Trip ===\n");

const wb4 = new Workbook();
const ws4 = wb4.addWorksheet("Notes");
ws4.addRow(["Name", "Address"]);
ws4.addRow(["Alice", "123 Main St\nApt 4\nNew York"]);
ws4.addRow(["Bob", "456 Oak Ave\nLondon"]);

const mdMultiline = wb4.writeMarkdown({ sheetName: "Notes" });
console.log("Written:");
console.log(mdMultiline);

const wb5 = new Workbook();
const ws5 = wb5.readMarkdown(mdMultiline, { sheetName: "Notes", convertBr: true });
console.log("Parsed back (address):", JSON.stringify(ws5.getRow(2).getCell(2).value));
fs.writeFileSync(path.join(outDir, "multiline-workbook.md"), mdMultiline, "utf8");
console.log();

// =============================================================================
// 7. Error handling
// =============================================================================

console.log("=== 7. Error Handling ===\n");

try {
  parseMarkdown("no table here");
} catch (e) {
  if (e instanceof MarkdownParseError) {
    console.log("MarkdownParseError:", e.message);
    console.log("  line:", e.line);
  }
}

console.log();

// =============================================================================
// 8. File I/O (Node.js)
// =============================================================================

console.log("=== 8. File I/O ===\n");

const wb6 = new Workbook();
wb6.readMarkdown("| X | Y |\n| - | - |\n| 1 | 2 |", { sheetName: "Data" });
const outFile = path.join(outDir, "file-io-output.md");
await wb6.writeMarkdownFile(outFile, { sheetName: "Data" });
console.log("Wrote:", outFile);

const wb7 = new Workbook();
await wb7.readMarkdownFile(outFile, { sheetName: "FromFile" });
console.log("Read back:", wb7.getWorksheet("FromFile")?.getRow(2).getCell(1).value);

// readMarkdownAllFile
const multiFile = path.join(outDir, "multi-input.md");
fs.writeFileSync(multiFile, multiDoc, "utf8");
const wb8 = new Workbook();
const allSheets = await wb8.readMarkdownAllFile(multiFile, { sheetName: "Sheet" });
console.log(
  `readMarkdownAllFile: ${allSheets.length} worksheets → ${allSheets.map(s => s.name).join(", ")}`
);

console.log("\nAll output files written to:", outDir);
