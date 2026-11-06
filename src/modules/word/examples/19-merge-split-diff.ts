/**
 * Word Example 19 — Merge / Split / Diff
 *
 * Covers:
 *   - Transform.merge — combine N documents with section breaks between them
 *   - Transform.split — by section / by pageBreak / by heading
 *   - Diff.compare — structural diff between two related documents
 *   - Edge cases: merging a single document, splitting a document with no
 *     splittable boundary (returns the original), diffing identical docs.
 *
 * Output: tmp/word-examples/19-merge-split-diff/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Build, Diff, Document, Transform } from "../index";
import type { DocxDocument } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/19-merge-split-diff"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Helpers — produce three small chapters
// ---------------------------------------------------------------------------
function makeChapter(title: string, body: string[]): DocxDocument {
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, title, 1);
  for (const p of body) {
    Document.addParagraph(d, p);
  }
  return d;
}

const ch1 = makeChapter("Chapter 1 — Origins", [
  "It begins with a single line.",
  "Then another, and another."
]);
const ch2 = makeChapter("Chapter 2 — Conflict", [
  "Things grow more complicated.",
  "There is tension."
]);
const ch3 = makeChapter("Chapter 3 — Resolution", ["The hero prevails.", "Curtain falls."]);

// ---------------------------------------------------------------------------
// 1. Merge
// ---------------------------------------------------------------------------
const merged = Transform.merge([ch1, ch2, ch3], { sectionBreak: "nextPage" });
const mergedBuf = await Document.toBuffer(merged);
fs.writeFileSync(path.join(outDir, "01-merged.docx"), mergedBuf);
console.log(`  → 01-merged.docx (${mergedBuf.length} bytes)`);

// Edge: merge a single document
const oneMerged = Transform.merge([ch1]);
fs.writeFileSync(path.join(outDir, "01-single-merged.docx"), await Document.toBuffer(oneMerged));

// Edge: empty merge — invalid, requires at least one doc; we guard:
try {
  Transform.merge([]);
} catch (err) {
  console.log(`  empty merge correctly rejected: ${(err as Error).message}`);
}

// ---------------------------------------------------------------------------
// 2. Split — by section
// ---------------------------------------------------------------------------
const splits = Transform.split(merged, { by: "section" });
console.log(`  Transform.split(by:"section") produced ${splits.length} parts`);
for (const [i, part] of splits.entries()) {
  const buf = await Document.toBuffer(part);
  fs.writeFileSync(path.join(outDir, `02-split-${i + 1}.docx`), buf);
}

// Build a doc with explicit page breaks → split by page break
const pageDoc = Document.create();
Document.useDefaultStyles(pageDoc);
Document.addParagraph(pageDoc, "Page A content.");
Document.addParagraph(pageDoc, Build.paragraph([Build.pageBreak()]));
Document.addParagraph(pageDoc, "Page B content.");
Document.addParagraph(pageDoc, Build.paragraph([Build.pageBreak()]));
Document.addParagraph(pageDoc, "Page C content.");
const pageSplits = Transform.split(pageDoc, { by: "pageBreak" });
console.log(`  Transform.split(by:"pageBreak") produced ${pageSplits.length} parts`);
for (const [i, part] of pageSplits.entries()) {
  const buf = await Document.toBuffer(part);
  fs.writeFileSync(path.join(outDir, `03-pages-${i + 1}.docx`), buf);
}

// Build a doc with multiple Heading 1 → split by heading
const headingDoc = Document.create();
Document.useDefaultStyles(headingDoc);
Document.addHeading(headingDoc, "First topic", 1);
Document.addParagraph(headingDoc, "First topic body.");
Document.addHeading(headingDoc, "Second topic", 1);
Document.addParagraph(headingDoc, "Second topic body.");
Document.addHeading(headingDoc, "Third topic", 1);
Document.addParagraph(headingDoc, "Third topic body.");
const headingSplits = Transform.split(headingDoc, {
  by: "heading",
  headingLevel: 1
});
console.log(`  Transform.split(by:"heading", level:1) produced ${headingSplits.length} parts`);
for (const [i, part] of headingSplits.entries()) {
  const buf = await Document.toBuffer(part);
  fs.writeFileSync(path.join(outDir, `04-heading-${i + 1}.docx`), buf);
}

// Edge: split a doc with no splittable boundary → returns original
const tinyDoc = Document.create();
Document.useDefaultStyles(tinyDoc);
Document.addParagraph(tinyDoc, "single paragraph");
const tinySplits = Transform.split(tinyDoc, { by: "heading" });
console.log(`  Transform.split(no boundary) → ${tinySplits.length} part(s)`);

// ---------------------------------------------------------------------------
// 3. Diff
// ---------------------------------------------------------------------------
const oldDoc = (() => {
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Recipe", 1);
  Document.addParagraph(d, "Step 1: Mix dry ingredients.");
  Document.addParagraph(d, "Step 2: Add eggs.");
  Document.addParagraph(d, "Step 3: Bake at 180°C for 25 minutes.");
  return d;
})();
const newDoc = (() => {
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Recipe (revised)", 1);
  Document.addParagraph(d, "Step 1: Mix dry ingredients in a large bowl.");
  // Step 2 deleted
  Document.addParagraph(d, "Step 3: Bake at 200°C for 30 minutes.");
  Document.addParagraph(d, "Step 4: Cool before slicing.");
  return d;
})();

const diff = Diff.compare(oldDoc, newDoc);
console.log(
  `  Diff.compare summary: +${diff.summary.added}  -${diff.summary.deleted}  ~${diff.summary.modified}`
);
for (const change of diff.entries.slice(0, 10)) {
  console.log(`    [${change.type}] ${JSON.stringify(change).slice(0, 120)}`);
}

// Edge: identical docs
const identical = Diff.compare(oldDoc, oldDoc);
console.log(
  `  Diff.compare identical: +${identical.summary.added}  -${identical.summary.deleted}  ~${identical.summary.modified}`
);

// Render a "redline" demo: take the new doc and write a header summary
const redline = Document.create();
Document.useDefaultStyles(redline);
Document.addHeading(redline, "Diff summary", 1);
Document.addParagraph(
  redline,
  Build.paragraph([
    Build.text(`Added: ${diff.summary.added}, `),
    Build.text(`Deleted: ${diff.summary.deleted}, `),
    Build.text(`Modified: ${diff.summary.modified}`)
  ])
);
Document.addHeading(redline, "Old → New paragraphs", 2);
for (const change of diff.entries) {
  Document.addParagraph(
    redline,
    Build.paragraph([Build.bold(`[${change.type}] `), Build.text(JSON.stringify(change))])
  );
}
fs.writeFileSync(path.join(outDir, "05-diff-summary.docx"), await Document.toBuffer(redline));
console.log(`  → 05-diff-summary.docx`);

// Quiet linter
void Build.textParagraph;
