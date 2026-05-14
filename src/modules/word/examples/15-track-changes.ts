/**
 * Word Example 15 — Track Changes (revisions)
 *
 * Covers:
 *   - Inserted run (insertion mark)
 *   - Deleted run (deletion mark)
 *   - Moved-from / moved-to (relocation)
 *   - Move range start / end markers
 *   - listRevisions, acceptRevision, rejectRevision (selectively)
 *   - acceptAllRevisions, rejectAllRevisions
 *   - Edge case: nested inserted+deleted (a deletion within an inserted block)
 *
 * Output:
 *   - 15-revisions-pending.docx — all revisions still present
 *   - 15-revisions-accepted.docx — every revision accepted
 *   - 15-revisions-rejected.docx — every revision rejected
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  text,
  bold,
  insertedRun,
  deletedRun,
  movedFromRun,
  movedToRun,
  moveFromRangeStart,
  moveFromRangeEnd,
  moveToRangeStart,
  moveToRangeEnd,
  acceptAllRevisions,
  rejectAllRevisions,
  listRevisions,
  acceptRevision,
  rejectRevision,
  toBuffer
} from "../index";
import type { RevisionInfo } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const baseDate = "2026-05-01T09:00:00Z";
const rev = (id: number, author = "Alice"): RevisionInfo => ({ id, author, date: baseDate });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Track Changes Demo", 1);

// ---------------------------------------------------------------------------
// 1. Plain insertion + deletion within a paragraph
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  paragraph([
    text("The quick "),
    deletedRun(text("brown "), rev(1, "Alice")),
    insertedRun(text("red "), rev(2, "Bob")),
    text("fox jumps over the lazy dog.")
  ])
);

// ---------------------------------------------------------------------------
// 2. Insert a whole new sentence
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  paragraph([
    text("Original sentence kept. "),
    insertedRun(text("Brand-new sentence inserted by Bob. "), rev(3, "Bob")),
    text("Second original sentence kept.")
  ])
);

// ---------------------------------------------------------------------------
// 3. Move (paragraph relocation)
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  paragraph([
    moveFromRangeStart(10, "Carol", { date: baseDate, name: "swap-paragraphs" }),
    movedFromRun(text("This sentence has been moved away from here. "), rev(10, "Carol")),
    moveFromRangeEnd(10),
    text("(remaining text stays put.)")
  ])
);
Document.addParagraphElement(
  doc,
  paragraph([
    moveToRangeStart(10, "Carol", { date: baseDate, name: "swap-paragraphs" }),
    movedToRun(text("This sentence has been moved away from here. "), rev(10, "Carol")),
    moveToRangeEnd(10),
    text("It now belongs with the second paragraph.")
  ])
);

// ---------------------------------------------------------------------------
// 4. Edge case: deletion of a previously-inserted run.  In Word's data
// model an inserted run that was later deleted shows as both marks. We
// approximate by using a deletedRun that wraps an inserted run sub-tree
// in two separate paragraphs.
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  paragraph([
    text("Edge case: "),
    insertedRun(text("inserted then "), rev(20, "Dave")),
    deletedRun(text("deleted "), rev(21, "Dave")),
    bold("(both marks visible in pending file)")
  ])
);

const built = Document.build(doc);

// Pending — write as-is
const pending = await toBuffer(built);
fs.writeFileSync(path.join(outDir, "15-revisions-pending.docx"), pending);
console.log(`  → 15-revisions-pending.docx (${pending.length} bytes)`);

// List the revisions (proves the API can introspect them)
const revisions = listRevisions(built);
console.log(`  detected ${revisions.length} revisions`);

// Selectively accept & reject specific IDs
{
  const cloned = JSON.parse(JSON.stringify(built));
  acceptRevision(cloned, 2); // accept the "red" insertion
  rejectRevision(cloned, 1); // keep the original "brown"
  const buf = await toBuffer(cloned);
  fs.writeFileSync(path.join(outDir, "15-revisions-partial.docx"), buf);
  console.log(`  → 15-revisions-partial.docx (${buf.length} bytes)`);
}

// Accept all
{
  const cloned = JSON.parse(JSON.stringify(built));
  acceptAllRevisions(cloned);
  const buf = await toBuffer(cloned);
  fs.writeFileSync(path.join(outDir, "15-revisions-accepted.docx"), buf);
  console.log(`  → 15-revisions-accepted.docx (${buf.length} bytes)`);
}

// Reject all
{
  const cloned = JSON.parse(JSON.stringify(built));
  rejectAllRevisions(cloned);
  const buf = await toBuffer(cloned);
  fs.writeFileSync(path.join(outDir, "15-revisions-rejected.docx"), buf);
  console.log(`  → 15-revisions-rejected.docx (${buf.length} bytes)`);
}
