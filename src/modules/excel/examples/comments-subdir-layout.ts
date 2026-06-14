/**
 * Example: Comments with subdirectory layout (Issue #148)
 *
 * Demonstrates that excelts can read comments from xlsx files that use
 * the subdirectory layout (xl/comments/comment1.xml) produced by tools
 * like openpyxl, in addition to the standard flat layout (xl/comments1.xml).
 *
 * This example:
 * 1. Creates a workbook with comments (with authors) and writes it
 * 2. Reads it back and verifies comments + authors survive the round-trip
 * 3. Re-packs the xlsx into subdirectory layout with absolute rel targets
 * 4. Reads the subdirectory-layout file and verifies comments are parsed
 */
import fs from "node:fs";
import path from "node:path";

import { cellNote, cellComment, cellSetComment } from "@excel/cell";
import { Cell, Workbook } from "@excel/index";
import { getCell } from "@excel/worksheet";

import { noteCreate } from "../../../index";
import { ArchiveFile } from "../../archive/fs/archive-file";

const tmpDir = path.resolve(import.meta.dirname, "../../../../tmp");
fs.mkdirSync(tmpDir, { recursive: true });

const flatFile = path.join(tmpDir, "comments-flat.xlsx");
const subdirFile = path.join(tmpDir, "comments-subdir.xlsx");

// ---------------------------------------------------------------------------
// Step 1: Create a workbook with comments (with different authors)
// ---------------------------------------------------------------------------
console.log("Step 1: Creating workbook with comments...");
const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Sheet1");

Cell.setValue(ws, "A1", "Hello");
cellSetComment(getCell(ws, "A1"), noteCreate({ texts: [{ text: "Comment by Alice" }] }, "Alice"));

Cell.setValue(ws, "B2", "World");
cellSetComment(getCell(ws, "B2"), noteCreate({ texts: [{ text: "Comment by Bob" }] }, "Bob"));

await Workbook.writeXlsx(wb, flatFile);
console.log(`  Written flat-layout file: ${flatFile}`);

// ---------------------------------------------------------------------------
// Step 2: Read back the flat-layout file and verify
// ---------------------------------------------------------------------------
console.log("\nStep 2: Reading flat-layout file...");
const wb2 = Workbook.create();
await Workbook.readXlsxFile(wb2, flatFile);
const ws2 = Workbook.getWorksheet(wb2, "Sheet1")!;

const a1 = getCell(ws2, "A1");
const b2 = getCell(ws2, "B2");
console.log(`  A1 note: ${JSON.stringify(cellNote(a1))}, author: ${cellComment(a1)?.author}`);
console.log(`  B2 note: ${JSON.stringify(cellNote(b2))}, author: ${cellComment(b2)?.author}`);
console.log(`  A1 has comment: ${cellNote(a1) != null ? "YES" : "MISSING"}`);
console.log(`  B2 has comment: ${cellNote(b2) != null ? "YES" : "MISSING"}`);
console.log(
  `  A1 author correct: ${cellComment(a1)?.author === "Alice" ? "YES" : `NO (got "${cellComment(a1)?.author}")`}`
);
console.log(
  `  B2 author correct: ${cellComment(b2)?.author === "Bob" ? "YES" : `NO (got "${cellComment(b2)?.author}")`}`
);

// ---------------------------------------------------------------------------
// Step 3: Re-pack into subdirectory layout (simulating openpyxl output)
// ---------------------------------------------------------------------------
console.log("\nStep 3: Re-packing into subdirectory layout...");

const flatBytes = fs.readFileSync(flatFile);
const flatZip = ArchiveFile.fromBuffer(flatBytes);
const subdirZip = new ArchiveFile();

for (const entry of flatZip.getEntriesSync()) {
  const data = flatZip.readEntrySync(entry.path);
  if (!data) {
    continue;
  }

  // Rename xl/comments1.xml → xl/comments/comment1.xml
  if (/^xl\/comments(\d+)\.xml$/.test(entry.path)) {
    const idx = entry.path.match(/comments(\d+)/)?.[1];
    const newPath = `xl/comments/comment${idx}.xml`;
    console.log(`  Renaming: ${entry.path} → ${newPath}`);
    subdirZip.addBuffer(data, newPath);
    continue;
  }

  // Rewrite [Content_Types].xml to reference the new path
  if (entry.path === "[Content_Types].xml") {
    let xml = new TextDecoder().decode(data);
    xml = xml.replace(
      /PartName="\/xl\/comments(\d+)\.xml"/g,
      'PartName="/xl/comments/comment$1.xml"'
    );
    subdirZip.addBuffer(new TextEncoder().encode(xml), entry.path);
    continue;
  }

  // Rewrite worksheet rels to use absolute target for comments
  if (/^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(entry.path)) {
    let xml = new TextDecoder().decode(data);
    xml = xml.replace(/Target="\.\.\/comments(\d+)\.xml"/g, 'Target="/xl/comments/comment$1.xml"');
    console.log(`  Rewrote rels: ${entry.path}`);
    subdirZip.addBuffer(new TextEncoder().encode(xml), entry.path);
    continue;
  }

  subdirZip.addBuffer(data, entry.path);
}

const subdirBytes = subdirZip.toBufferSync();
fs.writeFileSync(subdirFile, subdirBytes);
console.log(`  Written subdirectory-layout file: ${subdirFile}`);

// ---------------------------------------------------------------------------
// Step 4: Read the subdirectory-layout file and verify comments are parsed
// ---------------------------------------------------------------------------
console.log("\nStep 4: Reading subdirectory-layout file...");
const wb3 = Workbook.create();
await Workbook.readXlsxFile(wb3, subdirFile);
const ws3 = Workbook.getWorksheet(wb3, "Sheet1")!;

const a1s = getCell(ws3, "A1");
const b2s = getCell(ws3, "B2");
console.log(`  A1 note: ${JSON.stringify(cellNote(a1s))}, author: ${cellComment(a1s)?.author}`);
console.log(`  B2 note: ${JSON.stringify(cellNote(b2s))}, author: ${cellComment(b2s)?.author}`);
console.log(`  A1 has comment: ${cellNote(a1s) != null ? "YES" : "MISSING"}`);
console.log(`  B2 has comment: ${cellNote(b2s) != null ? "YES" : "MISSING"}`);
console.log(
  `  A1 author correct: ${cellComment(a1s)?.author === "Alice" ? "YES" : `NO (got "${cellComment(a1s)?.author}")`}`
);
console.log(
  `  B2 author correct: ${cellComment(b2s)?.author === "Bob" ? "YES" : `NO (got "${cellComment(b2s)?.author}")`}`
);

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------
const allPassed =
  cellNote(a1) != null &&
  cellNote(b2) != null &&
  cellComment(a1)?.author === "Alice" &&
  cellComment(b2)?.author === "Bob" &&
  cellNote(a1s) != null &&
  cellNote(b2s) != null &&
  cellComment(a1s)?.author === "Alice" &&
  cellComment(b2s)?.author === "Bob";

console.log(
  `\n${allPassed ? "SUCCESS" : "FAILURE"}: Comments and authors ${allPassed ? "correctly round-tripped" : "NOT fully preserved"} across both layouts.`
);

if (!allPassed) {
  process.exit(1);
}
