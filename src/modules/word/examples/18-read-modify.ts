/**
 * Word Example 18 — Read, query, modify
 *
 * Covers:
 *   - readDocx → introspect / mutate / write back
 *   - extractText (whole document) and paragraph-by-paragraph access
 *   - searchText with string and regex queries
 *   - replaceText (string / regex with $1 backrefs)
 *   - mailMerge (MERGEFIELD substitution)
 *   - paragraphCount, countWords, getHeadings, listImages, listHyperlinks,
 *     listTables, listSections, tableCount
 *   - searchByFormat / countByFormat / getUsedFormats
 *   - Edge case: replace across runs, replace inside a hyperlink, replace
 *     inside a footnote.
 *
 * Output: tmp/word-examples/18-read-modify/...
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
  italic,
  hyperlink,
  field,
  readDocx,
  toBuffer,
  extractText,
  searchText,
  replaceText,
  mailMerge,
  paragraphCount,
  countWords,
  getHeadings,
  listImages,
  listHyperlinks,
  listTables,
  listSections,
  tableCount,
  searchByFormat,
  countByFormat,
  getUsedFormats,
  cmToEmu
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/18-read-modify"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Step 1: produce a moderately complex DOCX to read back
// ---------------------------------------------------------------------------
const src = Document.create();
Document.useDefaultStyles(src);
Document.addHeading(src, "Source document", 1);
Document.addHeading(src, "Section A", 2);
Document.addParagraph(src, "Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
Document.addParagraphElement(
  src,
  paragraph([
    text("Mixed run: "),
    bold("the brown fox "),
    italic("jumps over"),
    text(" the lazy dog."),
    hyperlink(" example.com", { url: "https://example.com" })
  ])
);
Document.addHeading(src, "Section B", 2);
Document.addTable(
  src,
  [
    ["Column 1", "Column 2"],
    ["a", "b"],
    ["c", "d"]
  ],
  { headerRow: true, borders: true }
);

// MERGEFIELD instruction — two flavours: quoted name and bare name
Document.addParagraphElement(
  src,
  paragraph([
    text("Hello "),
    field(' MERGEFIELD "FullName" ', "Name"),
    text(", balance: "),
    field(" MERGEFIELD AccountBalance ", "$0.00"),
    text(".")
  ])
);

// Add an inline image so listImages has something to find
const tinyPng = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82
]);
Document.addImage(src, tinyPng, "png", cmToEmu(1), cmToEmu(1), { altText: "test pixel" });

// Footnote so we can demonstrate replaceText reaching into footnotes
const fnId = Document.addFootnote(src, "Footnote with PLACEHOLDER token to replace.");
Document.addParagraphElement(
  src,
  paragraph([
    text("This sentence has a footnote ref."),
    {
      properties: { vertAlign: "superscript" },
      content: [{ type: "footnoteRef", id: fnId }]
    }
  ])
);

const sourceBytes = await toBuffer(Document.build(src));
fs.writeFileSync(path.join(outDir, "00-source.docx"), sourceBytes);
console.log(`  → 00-source.docx (${sourceBytes.length} bytes)`);

// ---------------------------------------------------------------------------
// Step 2: read it back
// ---------------------------------------------------------------------------
const doc = await readDocx(sourceBytes);

// Whole-document plain text
const fullText = extractText(doc);
console.log("  Full text (first 120 chars):", JSON.stringify(fullText.slice(0, 120)));

// Stats
console.log(`  paragraphCount: ${paragraphCount(doc)}`);
console.log(`  wordCount:      ${countWords(doc)}`);
console.log(`  tableCount:     ${tableCount(doc)}`);
console.log(`  imageCount:     ${listImages(doc).length}`);
console.log(`  hyperlinkCount: ${listHyperlinks(doc).length}`);
console.log(`  sectionCount:   ${listSections(doc).length}`);

// Headings outline
const headings = getHeadings(doc);
console.log(`  ${headings.length} headings:`);
for (const h of headings) {
  console.log(`    - L${h.level}: "${h.text}"`);
}

// Tables list with shape
const tables = listTables(doc);
console.log(`  ${tables.length} tables:`);
for (const t of tables) {
  console.log(`    - ${t.rows.length} rows × ${t.rows[0]?.cells.length ?? 0} cols`);
}

// ---------------------------------------------------------------------------
// Step 3: searchText (string and regex)
// ---------------------------------------------------------------------------
const stringHits = searchText(doc, "fox");
console.log(`  searchText("fox"): ${stringHits.length} hit(s)`);
const regexHits = searchText(doc, /\b[A-Z][a-z]+\b/g);
console.log(`  searchText(/\\b[A-Z]\\w+\\b/g): ${regexHits.length} hit(s)`);

// ---------------------------------------------------------------------------
// Step 4: replaceText — both across runs and inside hyperlinks/footnotes
// ---------------------------------------------------------------------------
const replaced1 = await readDocx(sourceBytes);
const c1 = replaceText(replaced1, "brown", "RED"); // simple
const c2 = replaceText(replaced1, /lazy (dog)/, "vigilant $1"); // regex with backref
const c3 = replaceText(replaced1, "PLACEHOLDER", "REPLACED"); // inside footnote
console.log(`  replaceText counts: brown→RED:${c1}, regex:${c2}, footnote:${c3}`);
const buf1 = await toBuffer(replaced1);
fs.writeFileSync(path.join(outDir, "01-replaced.docx"), buf1);
console.log(`  → 01-replaced.docx (${buf1.length} bytes)`);

// ---------------------------------------------------------------------------
// Step 5: mailMerge — MERGEFIELDs become real values
// ---------------------------------------------------------------------------
const merged = await readDocx(sourceBytes);
const mergedCount = mailMerge(merged, {
  FullName: "Jane Doe",
  AccountBalance: "$2,345.67"
});
console.log(`  mailMerge replaced ${mergedCount} fields`);
const buf2 = await toBuffer(merged);
fs.writeFileSync(path.join(outDir, "02-merged.docx"), buf2);
console.log(`  → 02-merged.docx (${buf2.length} bytes)`);

// mailMerge with removeUnmatched: missing keys → field is cleared instead of
// keeping its placeholder text.
const partial = await readDocx(sourceBytes);
const partialCount = mailMerge(
  partial,
  { FullName: "Solo" /* no AccountBalance */ },
  { removeUnmatched: true }
);
console.log(`  mailMerge({removeUnmatched:true}) replaced ${partialCount}`);
fs.writeFileSync(path.join(outDir, "02b-merged-partial.docx"), await toBuffer(partial));

// ---------------------------------------------------------------------------
// Step 6: format-search — find every bold or italic run
// ---------------------------------------------------------------------------
const usedFormats = getUsedFormats(doc);
console.log(`  ${usedFormats.length} distinct formats in use (sample):`);
for (const f of usedFormats.slice(0, 4)) {
  console.log(`    · ${JSON.stringify(f)}`);
}
const boldHits = searchByFormat(doc, { bold: true });
const italicCount = countByFormat(doc, { italic: true });
console.log(`  searchByFormat({ bold:true }): ${boldHits.length} hits`);
console.log(`  countByFormat({ italic:true }): ${italicCount}`);

// ---------------------------------------------------------------------------
// Edge cases — replace across runs (search string straddles formatting
// boundaries) and replace nothing
// ---------------------------------------------------------------------------
const edge = await readDocx(sourceBytes);
// "the brown fox" sits across normal + bold runs — replaceText must stitch
const acrossCount = replaceText(edge, "the brown fox", "BIG_RED_FOX");
const noopCount = replaceText(edge, "this string is not in the doc", "x");
console.log(`  edge: across-run replace=${acrossCount}, noop=${noopCount}`);
const buf3 = await toBuffer(edge);
fs.writeFileSync(path.join(outDir, "03-edge.docx"), buf3);

// ---------------------------------------------------------------------------
// Edge cases — replaceText reaches every container
// ---------------------------------------------------------------------------

// Build a doc with the same placeholder PLACEHOLDER inside:
// 1) the body, 2) a header, 3) a footer, 4) a table cell, 5) a footnote,
// 6) an endnote, 7) a comment, 8) inside a hyperlink, 9) inside an SDT.
{
  const richDoc = Document.create();
  Document.useDefaultStyles(richDoc);
  Document.addParagraph(richDoc, "Body PLACEHOLDER once.");
  Document.addParagraphElement(
    richDoc,
    paragraph([
      text("Inside hyperlink: "),
      hyperlink("PLACEHOLDER", { url: "https://example.com" })
    ])
  );

  // Header / footer
  Document.setHeader(richDoc, "default", {
    children: [textParagraph("Header has PLACEHOLDER too.")]
  });
  Document.setFooter(richDoc, "default", {
    children: [textParagraph("Footer PLACEHOLDER")]
  });

  // Inline table with the placeholder
  Document.addTable(
    richDoc,
    [
      ["Header", "Value"],
      ["row 1", "PLACEHOLDER in cell"]
    ],
    { headerRow: true, borders: true }
  );

  // Footnote + endnote
  Document.addFootnote(richDoc, "Footnote with PLACEHOLDER inside.");
  Document.addEndnote(richDoc, "Endnote PLACEHOLDER.");

  // Comment
  const cId = Document.addComment(richDoc, "Reviewer", "Comment body has PLACEHOLDER somewhere.");
  Document.addParagraphElement(richDoc, paragraph([{ type: "commentReference", id: cId }]));

  const richBytes = await toBuffer(Document.build(richDoc));
  const richModel = await readDocx(richBytes);
  const total = replaceText(richModel, "PLACEHOLDER", "FILLED");
  console.log(`  multi-container replace replaced ${total} occurrence(s)`);
  const richOut = await toBuffer(richModel);
  fs.writeFileSync(path.join(outDir, "04-multi-container.docx"), richOut);
  console.log(`  → 04-multi-container.docx (${richOut.length} bytes)`);
}
console.log(`  → 03-edge.docx (${buf3.length} bytes)`);
