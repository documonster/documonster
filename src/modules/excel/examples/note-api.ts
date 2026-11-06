/**
 * Cell notes — write, read and remove a note, its author included, through `Cell`.
 *
 * - Cell.setNote       — text or a note configuration, the author in the options
 * - Cell.getNote       — the text or configuration back, creating nothing
 * - Cell.getNoteAuthor — who wrote it
 * - Cell.removeNote    — take it off; answers whether there was one
 *
 * The workbook is written and read back so the round trip is observable.
 *
 * Usage:
 *   pnpm example --filter note-api
 *
 * Output:
 *   tmp/excel-examples/note-api.xlsx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { writeBothFormats } from "@excel/examples/utils/write-both";
import { Cell, Workbook } from "@excel/index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/excel-examples"
);
fs.mkdirSync(outDir, { recursive: true });
const filename = process.argv[2] ?? path.join(outDir, "note-api.xlsx");

const wb = Workbook.create();
const ws = Workbook.addWorksheet(wb, "Notes");
Cell.setValue(ws, "B2", 42);
Cell.setValue(ws, "B3", 43);

// 1. Plain text, with an author
Cell.setNote(ws, "B2", "Reviewed and approved.", { author: "Alice" });

// 2. Rich text and a box size, no author
Cell.setNote(ws, "B3", {
  texts: [{ text: "Check: ", font: { bold: true } }, { text: "source is the Q3 ledger" }],
  width: 180,
  height: 60
});

console.log("B2 note:", Cell.getNote(ws, "B2"), "by", Cell.getNoteAuthor(ws, "B2"));
console.log("B3 note:", JSON.stringify(Cell.getNote(ws, "B3")));
console.log("C9 note (no cell there):", Cell.getNote(ws, "C9")); // undefined, and C9 is not created

// 3. Remove one — answers whether there was a note
Cell.setNote(ws, "B4", "temporary");
console.log("removed B4:", Cell.removeNote(ws, "B4")); // true
console.log("removed B4 again:", Cell.removeNote(ws, "B4")); // false

await writeBothFormats(wb, filename);

const reread = await Workbook.readFile(filename);
const sheet = Workbook.getWorksheet(reread, "Notes")!;
console.log("read back:", Cell.getNote(sheet, "B2"), "by", Cell.getNoteAuthor(sheet, "B2"));
console.log(`Done. Wrote ${filename}`);
