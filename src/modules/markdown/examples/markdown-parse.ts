/**
 * Example: Markdown Parsing
 *
 * Covers:
 * - Basic table parsing
 * - Alignment detection
 * - Tables without leading/trailing pipes
 * - Extracting tables from Markdown documents
 * - Parse options: trim, unescape, skipEmptyRows, maxRows, convertBr
 * - Multi-table parsing with Markdown.parseAll
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Markdown } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/markdown-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// =============================================================================
// 1. Basic parsing
// =============================================================================

const markdown1 = `
| Name  | Age | City     |
| ----- | --- | -------- |
| Alice | 30  | New York |
| Bob   | 25  | London   |
| Carol | 35  | Tokyo    |
`.trim();

const result1 = Markdown.parse(markdown1);
console.log("=== 1. Basic Parsing ===");
console.log("Headers:", result1.headers);
console.log("Rows:", result1.rows);
console.log();

// =============================================================================
// 2. Alignment detection
// =============================================================================

const markdown2 = `
| Left   | Center | Right  | Default |
| :----- | :----: | -----: | ------- |
| Alice  | 30     | $1,000 | hello   |
`.trim();

const result2 = Markdown.parse(markdown2);
console.log("=== 2. Alignment Detection ===");
for (let i = 0; i < result2.headers.length; i++) {
  console.log(`  ${result2.headers[i]}: ${result2.alignments[i]}`);
}
console.log();

// =============================================================================
// 3. Without pipes + from document
// =============================================================================

const markdown3 = `Name | Age\n--- | ---\nAlice | 30\nBob | 25`;
const result3 = Markdown.parse(markdown3);
console.log("=== 3. Without Pipes ===");
console.log("Headers:", result3.headers);
console.log("Rows:", result3.rows);
console.log();

const doc = `# Title\n\nSome text.\n\n| ID | Name |\n| -- | ---- |\n| 1  | Alice |\n\nEnd.`;
const result4 = Markdown.parse(doc);
console.log("From document:", result4.headers, result4.rows);
console.log();

// =============================================================================
// 4. Parse options
// =============================================================================

console.log("=== 4. Parse Options ===");

// trim
const markdownTrim = "| A |\n| --- |\n|  hello  |";
console.log("trim:true →", JSON.stringify(Markdown.parse(markdownTrim, { trim: true }).rows[0]));
console.log("trim:false →", JSON.stringify(Markdown.parse(markdownTrim, { trim: false }).rows[0]));

// unescape
const markdownEsc = "| A |\n| --- |\n| a \\| b |";
console.log("unescape:true →", Markdown.parse(markdownEsc, { unescape: true }).rows[0][0]);
console.log("unescape:false →", Markdown.parse(markdownEsc, { unescape: false }).rows[0][0]);

// maxRows
const bigRows = Array.from({ length: 100 }, (_, i) => `| row${i} |`).join("\n");
const markdownBig = `| Data |\n| --- |\n${bigRows}`;
console.log("maxRows:3 →", Markdown.parse(markdownBig, { maxRows: 3 }).rows.length, "rows");

// convertBr
const markdownBr = "| Note |\n| --- |\n| Line1<br>Line2 |";
console.log(
  "convertBr:true →",
  JSON.stringify(Markdown.parse(markdownBr, { convertBr: true }).rows[0])
);
console.log(
  "convertBr:false →",
  JSON.stringify(Markdown.parse(markdownBr, { convertBr: false }).rows[0])
);
console.log();

// =============================================================================
// 5. Multi-table parsing
// =============================================================================

const multiDoc = `
# Report

| Product | Revenue |
| ------- | ------: |
| Widget  | $10,000 |
| Gadget  | $25,000 |

Some text.

| Category | Amount  |
| -------- | ------: |
| Salaries | $50,000 |
| Rent     | $5,000  |
`.trim();

const tables = Markdown.parseAll(multiDoc);
console.log("=== 5. Multi-Table ===");
console.log(`Found ${tables.length} tables`);
for (let i = 0; i < tables.length; i++) {
  console.log(`  Table ${i + 1}: ${tables[i].headers.join(", ")} (${tables[i].rows.length} rows)`);
}

// Write parsed results to file
const output = tables
  .map(
    (t, i) =>
      `## Table ${i + 1}\n\nHeaders: ${t.headers.join(", ")}\nAlignments: ${t.alignments.join(", ")}\nRows:\n${t.rows.map(r => "  " + r.join(" | ")).join("\n")}`
  )
  .join("\n\n");
fs.writeFileSync(path.join(outDir, "parsed-tables.md"), output, "utf8");
console.log(`\nWrote: ${path.join(outDir, "parsed-tables.md")}`);
