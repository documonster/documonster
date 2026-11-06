/**
 * Example: Round-Trip & Workbook Integration
 *
 * Covers:
 * - format → parse → format round-trip fidelity
 * - ExcelMarkdown.read / ExcelMarkdown.format (bridge functions — `read` returns a new
 *   workbook, `format` takes one; not members of the `Workbook` namespace — it has none)
 * - ExcelMarkdown.read({ tables: "all" }) (multi-table → a workbook with one worksheet per table)
 * - ExcelMarkdown.addWorksheet (a table added to an existing workbook)
 * - ExcelMarkdown.toBuffer / ExcelMarkdown.toBufferSync
 * - Value mapper on ExcelMarkdown.read
 * - Multiline round-trip through Workbook
 * - Error handling with MarkdownParseError
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { ExcelMarkdown } from "@excel/bridge/markdown.node";
import { Cell, Row, Workbook, Worksheet } from "@excel/index";

import { Markdown, MarkdownParseError } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/markdown-examples"
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

const markdown1 = Markdown.format(headers, rows, {
  columns: [
    { header: "Name", alignment: "left" },
    { header: "Formula", alignment: "center" },
    { header: "Path", alignment: "right" }
  ]
});
console.log("Formatted:");
console.log(markdown1);

const parsed = Markdown.parse(markdown1);
const markdown2 = Markdown.format(parsed.headers, parsed.rows, {
  columns: parsed.headers.map((h, i) => ({
    header: h,
    alignment: parsed.alignments[i]
  }))
});

console.log("Round-trip match:", markdown1 === markdown2 ? "PASS" : "FAIL");
fs.writeFileSync(path.join(outDir, "round-trip.md"), markdown1, "utf8");
console.log();

// =============================================================================
// 2. ExcelMarkdown.read / ExcelMarkdown.format
// =============================================================================

console.log("=== 2. ExcelMarkdown.read / ExcelMarkdown.format ===\n");

const markdownInput = `
| Name  | Age | Department  |
| :---- | --: | :---------: |
| Alice | 30  | Engineering |
| Bob   | 25  | Marketing   |
| Carol | 35  | Engineering |
`.trim();

const wb1 = ExcelMarkdown.read(markdownInput, { sheetName: "Employees" });
const ws = Workbook.getWorksheet(wb1, "Employees")!;

console.log("Sheet:", Worksheet.getName(ws), "| Rows:", Row.getCount(ws));
for (let r = 1; r <= Row.getCount(ws); r++) {
  console.log(
    `  Row ${r}: ${Cell.getValue(ws, { row: r, col: 1 })} | ${Cell.getValue(ws, { row: r, col: 2 })} | ${Cell.getValue(ws, { row: r, col: 3 })}`
  );
}

const mdOut = ExcelMarkdown.format(wb1, { sheetName: "Employees" });
console.log("\nWritten back:");
console.log(mdOut);
fs.writeFileSync(path.join(outDir, "workbook-single.md"), mdOut, "utf8");

// =============================================================================
// 3. ExcelMarkdown.read({ tables: "all" })
// =============================================================================

console.log('=== 3. ExcelMarkdown.read({ tables: "all" }) ===\n');

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

const wb2 = ExcelMarkdown.read(multiDoc, { sheetName: "Q1", tables: "all" });
const sheets = Workbook.listWorksheets(wb2);

console.log(`Created ${sheets.length} worksheets:`);
const allMarkdown: string[] = [];
for (const s of sheets) {
  const markdown = ExcelMarkdown.format(wb2, { sheetName: Worksheet.getName(s) });
  console.log(`\n--- ${Worksheet.getName(s)} ---`);
  console.log(markdown);
  allMarkdown.push(`## ${Worksheet.getName(s)}\n\n${markdown}`);
}
fs.writeFileSync(
  path.join(outDir, "workbook-multi.md"),
  "# Multi-Table Workbook\n\n" + allMarkdown.join("\n"),
  "utf8"
);

// =============================================================================
// 4. ExcelMarkdown.toBuffer
// =============================================================================

console.log("=== 4. ExcelMarkdown.toBuffer ===\n");
const buffer = await ExcelMarkdown.toBuffer(wb2, { sheetName: "Q1" });
console.log(`Buffer: ${buffer.length} bytes`);
const bufferSync = ExcelMarkdown.toBufferSync(wb2, { sheetName: "Q1" });
console.log(`toBufferSync: ${bufferSync.length} bytes (same bytes, no await)`);
fs.writeFileSync(path.join(outDir, "workbook-buffer.md"), buffer);
console.log();

// =============================================================================
// 5. Value mapper
// =============================================================================

console.log("=== 5. Value Mapper ===\n");

const wb3 = ExcelMarkdown.read(markdownInput, {
  sheetName: "Typed",
  map: (value: string, _col: number) => {
    const n = Number(value);
    return Number.isNaN(n) ? value : n;
  }
});
const ws3 = Workbook.getWorksheet(wb3, "Typed")!;
console.log("Age type:", typeof Cell.getValue(ws3, { row: 2, col: 2 }));
console.log("Age value:", Cell.getValue(ws3, { row: 2, col: 2 }));
console.log();

// =============================================================================
// 6. Multiline round-trip through Workbook
// =============================================================================

console.log("=== 6. Multiline Workbook Round-Trip ===\n");

const wb4 = Workbook.create();
const ws4 = Workbook.addWorksheet(wb4, "Notes");
Row.add(ws4, ["Name", "Address"]);
Row.add(ws4, ["Alice", "123 Main St\nApt 4\nNew York"]);
Row.add(ws4, ["Bob", "456 Oak Ave\nLondon"]);

const markdownMultiline = ExcelMarkdown.format(wb4, { sheetName: "Notes" });
console.log("Written:");
console.log(markdownMultiline);

// The table goes onto the workbook it came from, beside the original sheet
const ws5 = ExcelMarkdown.addWorksheet(wb4, markdownMultiline, {
  sheetName: "NotesParsed",
  convertBr: true
});
console.log("Parsed back (address):", JSON.stringify(Cell.getValue(ws5, { row: 2, col: 2 })));
fs.writeFileSync(path.join(outDir, "multiline-workbook.md"), markdownMultiline, "utf8");
console.log();

// =============================================================================
// 7. Error handling
// =============================================================================

console.log("=== 7. Error Handling ===\n");

try {
  Markdown.parse("no table here");
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

const wb6 = ExcelMarkdown.read("| X | Y |\n| - | - |\n| 1 | 2 |", { sheetName: "Data" });
const outFile = path.join(outDir, "file-io-output.md");
await ExcelMarkdown.writeFile(wb6, outFile, { sheetName: "Data" });
console.log("Wrote:", outFile);

const wb7 = await ExcelMarkdown.readFile(outFile, { sheetName: "FromFile" });
const fromFileSheet = Workbook.getWorksheet(wb7, "FromFile")!;
if (fromFileSheet) {
  console.log("Read back:", Cell.getValue(fromFileSheet, { row: 2, col: 1 }));
}

// ExcelMarkdown.readFile({ tables: "all" })
const multiFile = path.join(outDir, "multi-input.md");
fs.writeFileSync(multiFile, multiDoc, "utf8");
const wb8 = await ExcelMarkdown.readFile(multiFile, { sheetName: "Sheet", tables: "all" });
const allSheets = Workbook.listWorksheets(wb8);
console.log(
  `ExcelMarkdown.readFile (all tables): ${allSheets.length} worksheets → ${allSheets.map(s => Worksheet.getName(s)).join(", ")}`
);

console.log("\nAll output files written to:", outDir);
