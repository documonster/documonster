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

import { Workbook, Note } from "../../../index";
import { ArchiveFile } from "../../archive/fs/archive-file";

const tmpDir = path.resolve(import.meta.dirname, "../../../../tmp");
fs.mkdirSync(tmpDir, { recursive: true });

const flatFile = path.join(tmpDir, "comments-flat.xlsx");
const subdirFile = path.join(tmpDir, "comments-subdir.xlsx");

// ---------------------------------------------------------------------------
// Step 1: Create a workbook with comments (with different authors)
// ---------------------------------------------------------------------------
console.log("Step 1: Creating workbook with comments...");
const wb = new Workbook();
const ws = wb.addWorksheet("Sheet1");

ws.getCell("A1").value = "Hello";
ws.getCell("A1").comment = new Note({ texts: [{ text: "Comment by Alice" }] }, "Alice");

ws.getCell("B2").value = "World";
ws.getCell("B2").comment = new Note({ texts: [{ text: "Comment by Bob" }] }, "Bob");

await wb.xlsx.writeFile(flatFile);
console.log(`  Written flat-layout file: ${flatFile}`);

// ---------------------------------------------------------------------------
// Step 2: Read back the flat-layout file and verify
// ---------------------------------------------------------------------------
console.log("\nStep 2: Reading flat-layout file...");
const wb2 = new Workbook();
await wb2.xlsx.readFile(flatFile);
const ws2 = wb2.getWorksheet("Sheet1")!;

const a1 = ws2.getCell("A1");
const b2 = ws2.getCell("B2");
console.log(`  A1 note: ${JSON.stringify(a1.note)}, author: ${a1.comment?.author}`);
console.log(`  B2 note: ${JSON.stringify(b2.note)}, author: ${b2.comment?.author}`);
console.log(`  A1 has comment: ${a1.note != null ? "YES" : "MISSING"}`);
console.log(`  B2 has comment: ${b2.note != null ? "YES" : "MISSING"}`);
console.log(
  `  A1 author correct: ${a1.comment?.author === "Alice" ? "YES" : `NO (got "${a1.comment?.author}")`}`
);
console.log(
  `  B2 author correct: ${b2.comment?.author === "Bob" ? "YES" : `NO (got "${b2.comment?.author}")`}`
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
const wb3 = new Workbook();
await wb3.xlsx.readFile(subdirFile);
const ws3 = wb3.getWorksheet("Sheet1")!;

const a1s = ws3.getCell("A1");
const b2s = ws3.getCell("B2");
console.log(`  A1 note: ${JSON.stringify(a1s.note)}, author: ${a1s.comment?.author}`);
console.log(`  B2 note: ${JSON.stringify(b2s.note)}, author: ${b2s.comment?.author}`);
console.log(`  A1 has comment: ${a1s.note != null ? "YES" : "MISSING"}`);
console.log(`  B2 has comment: ${b2s.note != null ? "YES" : "MISSING"}`);
console.log(
  `  A1 author correct: ${a1s.comment?.author === "Alice" ? "YES" : `NO (got "${a1s.comment?.author}")`}`
);
console.log(
  `  B2 author correct: ${b2s.comment?.author === "Bob" ? "YES" : `NO (got "${b2s.comment?.author}")`}`
);

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------
const allPassed =
  a1.note != null &&
  b2.note != null &&
  a1.comment?.author === "Alice" &&
  b2.comment?.author === "Bob" &&
  a1s.note != null &&
  b2s.note != null &&
  a1s.comment?.author === "Alice" &&
  b2s.comment?.author === "Bob";

console.log(
  `\n${allPassed ? "SUCCESS" : "FAILURE"}: Comments and authors ${allPassed ? "correctly round-tripped" : "NOT fully preserved"} across both layouts.`
);

if (!allPassed) {
  process.exit(1);
}
