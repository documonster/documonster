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

import {
  Document,
  paragraph,
  textParagraph,
  text,
  bold,
  pageBreak,
  pageNumberField,
  totalPagesField,
  sectionPagesField,
  sectionField,
  dateField,
  timeField,
  authorField,
  titleField,
  subjectField,
  keywordsField,
  fileNameField,
  fileSizeField,
  styleRefField,
  sequenceField,
  ifField,
  quoteField,
  includeTextField,
  includePictureField,
  toBuffer
} from "../index";

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
    textParagraph("1. Introduction\t1"),
    textParagraph("2. Property fields\t2"),
    textParagraph("3. Sequence captions\t3"),
    textParagraph("4. Computed fields\t4")
  ]
});
Document.addParagraphElement(doc, paragraph([pageBreak()]));

// ---------------------------------------------------------------------------
// 2. Heading content so the TOC has something to point at
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Introduction", 1);
Document.addParagraph(doc, "This document exercises every field helper.");
Document.addParagraphElement(
  doc,
  paragraph([
    bold("Page count: "),
    pageNumberField(),
    text(" of "),
    totalPagesField(),
    text("  (section "),
    sectionField(),
    text("/page within section: "),
    sectionPagesField(),
    text(")")
  ])
);
Document.addParagraphElement(doc, paragraph([pageBreak()]));

// ---------------------------------------------------------------------------
// 3. Property fields
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Property fields", 1);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Author: "),
    authorField(),
    text(" | Title: "),
    titleField(),
    text(" | Subject: "),
    subjectField()
  ])
);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Keywords: "),
    keywordsField(),
    text(" | Filename: "),
    fileNameField(),
    text(" | Size: "),
    fileSizeField()
  ])
);
Document.addParagraphElement(
  doc,
  paragraph([text("Generated on: "), dateField("yyyy-MM-dd"), text("  at "), timeField("HH:mm")])
);
Document.addParagraphElement(doc, paragraph([pageBreak()]));

// ---------------------------------------------------------------------------
// 4. Sequence captions (Figure 1, Figure 2, …)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Sequence captions", 1);
Document.addParagraphElement(
  doc,
  paragraph([bold("Figure "), sequenceField("Figure"), text(" — first figure")])
);
Document.addParagraph(doc, "[figure 1 placeholder]");
Document.addParagraphElement(
  doc,
  paragraph([bold("Figure "), sequenceField("Figure"), text(" — second figure")])
);
Document.addParagraph(doc, "[figure 2 placeholder]");

Document.addParagraphElement(
  doc,
  paragraph([bold("Table "), sequenceField("Table"), text(" — first table")])
);
Document.addParagraphElement(
  doc,
  paragraph([bold("Table "), sequenceField("Table"), text(" — second table")])
);
Document.addParagraphElement(doc, paragraph([pageBreak()]));

// ---------------------------------------------------------------------------
// 5. Computed fields — IF, QUOTE, STYLEREF
// ---------------------------------------------------------------------------
Document.addHeading(doc, "4. Computed fields", 1);
Document.addParagraphElement(
  doc,
  paragraph([
    text("IF field — branches on a condition: "),
    // Note: nested fields inside IF (e.g. NUMPAGES) require a nested-field
    // syntax that's only recognised by Word during interactive editing. We
    // demonstrate a simple literal comparison here so the cached value the
    // writer emits is always meaningful when opened in Word for the first
    // time.
    ifField('"a" = "a"', "(true branch shown)", "(false branch shown)", "(true branch shown)")
  ])
);
Document.addParagraphElement(
  doc,
  paragraph([text("QUOTE field — literal text: "), quoteField("hello world", "hello world")])
);
Document.addParagraphElement(
  doc,
  paragraph([
    text("STYLEREF — most recent Heading 1 above this run: "),
    styleRefField("Heading1", { cachedValue: "(updates when fields refresh)" })
  ])
);

// ---------------------------------------------------------------------------
// 6. INCLUDE fields — placeholders only (Word fetches the contents on
//    update; we ship a sensible cached value so the document opens cleanly).
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  paragraph([
    text("Included text from another file: "),
    includeTextField("./fragments/intro.docx", { cachedValue: "(cached: included intro)" })
  ])
);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Linked picture: "),
    includePictureField("./images/logo.png", "(cached: external logo)")
  ])
);

const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "10-toc-fields.docx"), buf);
console.log(`  → 10-toc-fields.docx (${buf.length} bytes)`);
