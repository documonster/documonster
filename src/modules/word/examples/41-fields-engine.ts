/**
 * Word Example 41 — Field engine + index/TOC fields
 *
 * Covers:
 *   - tcField — table of contents entry (alternative to deriving from headings)
 *   - indexEntryField (XE) — mark a word for the index
 *   - indexField (INDEX) — generate the index from XE entries
 *   - noteRefField — cross-reference to a footnote/endnote
 *   - updateFields — compute every field's cachedValue (PAGE, NUMPAGES, …)
 *     after layout
 *   - updateTableOfContents — refresh just the TOC entries
 *
 * Output: tmp/word-examples/41-fields/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/41-fields"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Document with TOC + XE entries + INDEX + NOTEREF
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);

// Cover/TOC
Document.addHeading(d, "Contents", 1);
Document.addTableOfContents(d, {
  headingStyleRange: "1-3",
  hyperlink: true,
  cachedParagraphs: [
    Build.paragraph([Build.text("Chapter 1\t1")]),
    Build.paragraph([Build.text("Chapter 2\t2")]),
    Build.paragraph([Build.text("Index\t3")])
  ]
});
Document.addParagraphElement(d, Build.paragraph([Build.pageBreak()]));

// Chapter 1 with TC marker for fine-grained TOC entry control
Document.addHeading(d, "Chapter 1 — Introduction", 1);
Document.addParagraphElement(
  d,
  Build.paragraph([
    Build.tcField("Special: An overlooked first principle", { level: 2 }),
    Build.text("This chapter introduces a "),
    Build.text("widget", { color: "0070C0" }),
    Build.indexEntryField("widget"),
    Build.text(" — a unit of inventory.")
  ])
);
const fnId = Document.addFootnote(d, "An older spelling: 'widgit'.");
Document.addParagraphElement(
  d,
  Build.paragraph([
    Build.text("See also "),
    {
      properties: { vertAlign: "superscript" },
      content: [{ type: "footnoteRef", id: fnId }]
    },
    Build.text(" for historical context.")
  ])
);
Document.addParagraphElement(d, Build.paragraph([Build.pageBreak()]));

// Chapter 2 with another XE
Document.addHeading(d, "Chapter 2 — Components", 1);
Document.addParagraphElement(
  d,
  Build.paragraph([
    Build.text("Each "),
    Build.text("gadget", { color: "0070C0" }),
    Build.indexEntryField("gadget"),
    Build.text(" is composed of multiple "),
    Build.text("widgets", { color: "0070C0" }),
    Build.indexEntryField("widget"),
    Build.text(".")
  ])
);

// NOTEREF cross-reference back to the footnote
Document.addParagraphElement(
  d,
  Build.paragraph([
    Build.text("Cross-reference back to the earlier footnote: "),
    Build.noteRefField("widgit-note")
  ])
);
Document.addParagraphElement(d, Build.paragraph([Build.pageBreak()]));

// Index page
Document.addHeading(d, "Index", 1);
Document.addParagraphElement(d, Build.paragraph([Build.indexField({ columns: 2 })]));

// Page X of Y in body so we can see updateFields populate the cachedValue
Document.addParagraphElement(
  d,
  Build.paragraph([
    Build.bold("Page "),
    Build.pageNumberField(),
    Build.text(" of "),
    Build.totalPagesField()
  ])
);

const built = Document.build(d);

// ---------------------------------------------------------------------------
// 2. Save raw (cachedValues are placeholders)
// ---------------------------------------------------------------------------
fs.writeFileSync(path.join(outDir, "01-raw.docx"), await Io.toBuffer(built));
console.log(`  → 01-raw.docx`);

// ---------------------------------------------------------------------------
// 3. updateFields — compute all field cached values via a layout pass
// ---------------------------------------------------------------------------
const updated = Io.updateFields(built);
fs.writeFileSync(path.join(outDir, "02-updated.docx"), await Io.toBuffer(updated));
console.log(`  → 02-updated.docx (every field cached value computed)`);

// ---------------------------------------------------------------------------
// 4. updateTableOfContents — only refresh the TOC entries
// ---------------------------------------------------------------------------
const tocOnly = Io.updateTableOfContents(built);
fs.writeFileSync(path.join(outDir, "03-toc-refreshed.docx"), await Io.toBuffer(tocOnly));
console.log(`  → 03-toc-refreshed.docx`);

// ---------------------------------------------------------------------------
// 5. Verify the model — count cached fields after update
// ---------------------------------------------------------------------------
let fieldsWithCache = 0;
for (const block of updated.body) {
  if (block.type === "paragraph") {
    for (const run of block.children) {
      if ("content" in run) {
        for (const c of run.content) {
          if (c.type === "field" && c.cachedValue) {
            fieldsWithCache++;
          }
        }
      }
    }
  }
}
console.log(`  fields with cached values after updateFields: ${fieldsWithCache}`);
