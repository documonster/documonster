/**
 * Word Example 11 — Footnotes, Endnotes & Comments
 *
 * Covers:
 *   - Footnote with simple text content
 *   - Footnote with rich content (multiple paragraphs, formatting)
 *   - Endnote
 *   - Comment range (start + reference + end)
 *   - Reply-style threaded comments via parentId
 *   - Multiple comments overlapping the same range
 *   - Edge cases: comment with empty body, comment without range
 *     (just a reference at a point)
 *
 * Output: tmp/word-examples/11-footnotes-endnotes-comments.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io } from "../index";
import type { Run } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Footnotes, Endnotes & Comments", 1);

// ---------------------------------------------------------------------------
// 1. Footnotes
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Footnotes", 2);

// Simple footnote
const fn1 = Document.addFootnote(doc, "First footnote — citation goes here.");

// Rich footnote (multiple paragraphs)
const fn2 = Document.addFootnote(doc, [
  Build.paragraph([
    Build.italic("First paragraph "),
    Build.text("of a richly formatted footnote.")
  ]),
  Build.paragraph([
    Build.text("Second paragraph (with "),
    Build.bold("bold"),
    Build.text(" and an URL: "),
    Build.text("https://example.com", { color: "0563C1", underline: "single" }),
    Build.text(").")
  ])
]);

// Reference the footnotes inline
function footnoteRef(id: number): Run {
  return {
    properties: { vertAlign: "superscript" },
    content: [{ type: "footnoteRef", id }]
  };
}

Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("This sentence has a footnote"),
    footnoteRef(fn1),
    Build.text(" attached, and another"),
    footnoteRef(fn2),
    Build.text(" with rich content.")
  ])
);

// ---------------------------------------------------------------------------
// 2. Endnotes
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Endnotes", 2);

const en1 = Document.addEndnote(doc, "First endnote — appears at the end of the document.");
const en2 = Document.addEndnote(doc, [
  Build.paragraph([
    Build.text("A multi-paragraph endnote. "),
    Build.text("Endnotes typically collect citations.")
  ])
]);

function endnoteRef(id: number): Run {
  return {
    properties: { vertAlign: "superscript" },
    content: [{ type: "endnoteRef", id }]
  };
}

Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("This sentence references endnote one"),
    endnoteRef(en1),
    Build.text(", and another endnote"),
    endnoteRef(en2),
    Build.text(" follows here.")
  ])
);

// ---------------------------------------------------------------------------
// 3. Comments
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Comments (review)", 2);

// Three independent comments
const c1 = Document.addComment(doc, "Alice", "Should we capitalise this term?", {
  initials: "AB",
  date: "2026-05-01T09:00:00Z"
});
const c2 = Document.addComment(doc, "Bob", "Agreed — also rephrase the second sentence.", {
  initials: "BC",
  date: "2026-05-01T10:30:00Z"
});
const c3 = Document.addComment(doc, "Alice", "Done.", {
  initials: "AB",
  date: "2026-05-01T11:00:00Z"
});

// Body content with overlapping comment ranges
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.commentRangeStart(c1),
    Build.text("This entire sentence is the target of comment 1."),
    Build.commentRangeEnd(c1),
    Build.commentReference(c1),
    Build.text(" Following normal text. "),
    Build.commentRangeStart(c2),
    Build.commentRangeStart(c3),
    Build.text("Both c2 and c3 cover this fragment."),
    Build.commentRangeEnd(c2),
    Build.commentRangeEnd(c3),
    Build.commentReference(c2),
    Build.commentReference(c3),
    Build.text(" End of paragraph.")
  ])
);

// Comment without range — point reference
const pointComment = Document.addComment(
  doc,
  "Reviewer",
  "Inserted-at-point comment (no range — pinned to this caret position)."
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Point-style comment placed mid-sentence "),
    Build.commentReference(pointComment),
    Build.text(" — the sticky note has no underlined range.")
  ])
);

// Edge case: comment with empty body (Word still renders the marker)
const emptyComment = Document.addComment(doc, "Carol", "");
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.commentRangeStart(emptyComment),
    Build.text("Even empty-bodied comments produce a valid review marker."),
    Build.commentRangeEnd(emptyComment),
    Build.commentReference(emptyComment)
  ])
);

// ---------------------------------------------------------------------------
// 4. Edge case: many footnote references inside a single table cell
//    (the writer must serialise each ref into the same cell paragraph).
// ---------------------------------------------------------------------------
Document.addHeading(doc, "4. Multiple footnotes inside one table cell", 2);
const fnA = Document.addFootnote(doc, "Footnote A — first citation in the cell.");
const fnB = Document.addFootnote(doc, "Footnote B — second citation in the cell.");
const fnC = Document.addFootnote(doc, "Footnote C — third citation in the cell.");
Document.addTableElement(doc, {
  type: "table",
  properties: { width: { value: 5000, type: "pct" } },
  rows: [
    {
      cells: [
        { content: [Build.textParagraph("Source")] },
        {
          content: [
            Build.paragraph([
              Build.text("Cited multiple times: "),
              footnoteRef(fnA),
              Build.text(", "),
              footnoteRef(fnB),
              Build.text(", "),
              footnoteRef(fnC)
            ])
          ]
        }
      ]
    }
  ]
});

// ---------------------------------------------------------------------------
// 5. Edge case: 100 footnotes in one document — stress the numbering pipeline
// ---------------------------------------------------------------------------
Document.addHeading(doc, "5. 100 footnotes (stress)", 2);
const stressIds: number[] = [];
for (let i = 1; i <= 100; i++) {
  stressIds.push(Document.addFootnote(doc, `Stress footnote #${i}`));
}
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text(`There are ${stressIds.length} footnotes attached to this paragraph.`)
  ])
);
// Reference every 10th one inline so the file actually links them
const sampledRefs = stressIds.filter((_, idx) => idx % 10 === 0).map(footnoteRef);
Document.addParagraphElement(
  doc,
  Build.paragraph([Build.text("Sampled refs:"), ...sampledRefs.flatMap(r => [Build.text(" "), r])])
);

const buf = await Io.toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "11-footnotes-endnotes-comments.docx"), buf);
console.log(`  → 11-footnotes-endnotes-comments.docx (${buf.length} bytes)`);
