/**
 * Word Example 10 — Table of Contents & document fields
 *
 * Covers:
 *   - addTableOfContents() — auto-update on open
 *   - TOC with custom heading range, custom leader, custom styles
 *   - Cached fallback content (so TOC shows entries even before Word's
 *     "Update Field" runs)
 *   - PAGE / NUMPAGES / SECTIONPAGES fields
 *   - DATE / TIME field with format string
 *   - SEQ field (auto-numbering captions, e.g. Figure 1, Figure 2)
 *   - Document property fields (AUTHOR / TITLE / FILENAME / FILESIZE / SUBJECT / KEYWORDS)
 *   - STYLEREF (running header / chapter title)
 *   - QUOTE / IF (computed fields)
 *   - INCLUDETEXT / INCLUDEPICTURE — placeholders only
 *   - REF / PAGEREF (already covered in 08, brief reuse here)
 *
 * Output: tmp/word-examples/10-toc-fields.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

// Set core properties so the property fields render meaningful values
Document.setCoreProperties(doc, {
  title: "Word Module Field Demo",
  creator: "OpenCode",
  subject: "Demonstrating fields and TOC",
  keywords: "docx, fields, TOC",
  description: "Sample DOCX exercising every field helper."
});

// ---------------------------------------------------------------------------
// 1. Table of Contents at the top
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Contents", 1);
Document.addTableOfContents(doc, {
  headingStyleRange: "1-3",
  hyperlink: true,
  leader: "dot",
  // Cached fallback so the TOC isn't blank on first open. Word will replace
  // these entries when the user runs "Update Field". The exact text doesn't
  // need to match the heading text — it's just a placeholder.
  cachedParagraphs: [
    Build.textParagraph("1. Introduction\t1"),
    Build.textParagraph("2. Property fields\t2"),
    Build.textParagraph("3. Sequence captions\t3"),
    Build.textParagraph("4. Computed fields\t4")
  ]
});
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));

// ---------------------------------------------------------------------------
// 1b. tocField() run-helper — same TOC field, but as a Run inside a paragraph
//     instead of a top-level helper. Useful when embedding a TOC field
//     mid-paragraph (e.g. inside a SDT) or composing custom TOC layouts.
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.bold("Inline TOC field: "),
    Build.tocField({
      headingLevels: "1-2",
      hyperlink: true,
      tabLeader: ".",
      cachedValue: "(updates on field refresh)"
    })
  ])
);
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));

// ---------------------------------------------------------------------------
// 2. Heading content so the TOC has something to point at
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Introduction", 1);
Document.addParagraph(doc, "This document exercises every field helper.");
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.bold("Page count: "),
    Build.pageNumberField(),
    Build.text(" of "),
    Build.totalPagesField(),
    Build.text("  (section "),
    Build.sectionField(),
    Build.text("/page within section: "),
    Build.sectionPagesField(),
    Build.text(")")
  ])
);
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));

// ---------------------------------------------------------------------------
// 3. Property fields
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Property fields", 1);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Author: "),
    Build.authorField(),
    Build.text(" | Title: "),
    Build.titleField(),
    Build.text(" | Subject: "),
    Build.subjectField()
  ])
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Keywords: "),
    Build.keywordsField(),
    Build.text(" | Filename: "),
    Build.fileNameField(),
    Build.text(" | Size: "),
    Build.fileSizeField()
  ])
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Generated on: "),
    Build.dateField("yyyy-MM-dd"),
    Build.text("  at "),
    Build.timeField("HH:mm")
  ])
);
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));

// ---------------------------------------------------------------------------
// 4. Sequence captions (Figure 1, Figure 2, …)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Sequence captions", 1);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.bold("Figure "),
    Build.sequenceField("Figure"),
    Build.text(" — first figure")
  ])
);
Document.addParagraph(doc, "[figure 1 placeholder]");
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.bold("Figure "),
    Build.sequenceField("Figure"),
    Build.text(" — second figure")
  ])
);
Document.addParagraph(doc, "[figure 2 placeholder]");

Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.bold("Table "),
    Build.sequenceField("Table"),
    Build.text(" — first table")
  ])
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.bold("Table "),
    Build.sequenceField("Table"),
    Build.text(" — second table")
  ])
);
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));

// ---------------------------------------------------------------------------
// 5. Computed fields — IF, QUOTE, STYLEREF
// ---------------------------------------------------------------------------
Document.addHeading(doc, "4. Computed fields", 1);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("IF field — branches on a condition: "),
    // Note: nested fields inside IF (e.g. NUMPAGES) require a nested-field
    // syntax that's only recognised by Word during interactive editing. We
    // demonstrate a simple literal comparison here so the cached value the
    // writer emits is always meaningful when opened in Word for the first
    // time.
    Build.ifField('"a" = "a"', "(true branch shown)", "(false branch shown)", "(true branch shown)")
  ])
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("QUOTE field — literal text: "),
    Build.quoteField("hello world", "hello world")
  ])
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("STYLEREF — most recent Heading 1 above this run: "),
    Build.styleRefField("Heading1", { cachedValue: "(updates when fields refresh)" })
  ])
);

// ---------------------------------------------------------------------------
// 6. INCLUDE fields — placeholders only (Word fetches the contents on
//    update; we ship a sensible cached value so the document opens cleanly).
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Included text from another file: "),
    Build.includeTextField("./fragments/intro.docx", { cachedValue: "(cached: included intro)" })
  ])
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Linked picture: "),
    Build.includePictureField("./images/logo.png", "(cached: external logo)")
  ])
);

const buf = await Io.toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "10-toc-fields.docx"), buf);
console.log(`  → 10-toc-fields.docx (${buf.length} bytes)`);
