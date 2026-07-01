/**
 * Word Example 26 — Markdown ↔ DOCX
 *
 * Covers:
 *   - markdownToDocx — full GFM features (async):
 *     · ATX & Setext headings, paragraphs, hard line breaks
 *     · Inline emphasis, strong, strike-through, code, links, images
 *     · Bullet, ordered and task lists (`- [ ]`, `- [x]`)
 *     · Block quotes, fenced code blocks (with language), thematic breaks
 *     · Tables (with column alignment)
 *     · Footnotes (`[^id]` references with `[^id]: …` definitions)
 *     · Custom image resolver (sync or async) — supply image data for
 *       `![](url)` and the image is embedded into word/media/
 *   - markdownToDocxBody — same pipeline returning body + supporting
 *     numbering / footnote / image definitions for splicing
 *   - renderToMarkdown — DOCX → GFM Markdown, choice of heading style
 *
 * Output:
 *   - 26-md-imported.docx
 *   - 26-md-roundtrip.md
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Io, Units } from "../index";
import { markdownToDocx, markdownToDocxBody, renderToMarkdown } from "../markdown";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/26-markdown"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Markdown → DOCX
// ---------------------------------------------------------------------------

const tinyPng = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82
]);

// We build the markdown by concatenation rather than a template literal —
// dollar-signs, back-ticks and backslashes inside Markdown collide with TS
// template-literal escaping rules.
const lines: string[] = [
  "# Markdown demo",
  "",
  "A paragraph with **bold**, *italic*, ~~strike~~, `code`, and [a link](https://example.com).",
  "",
  "Setext heading",
  "==============",
  "",
  "Subheading",
  "----------",
  "",
  "> A nested",
  "> > deeply nested",
  "> blockquote.",
  "",
  "## Lists",
  "",
  "- First bullet",
  "- Second bullet",
  "  - Nested 2.1",
  "  - Nested 2.2",
  "- Third bullet",
  "",
  "1. Step one",
  "2. Step two",
  "3. Step three",
  "",
  "Task list:",
  "",
  "- [x] Done",
  "- [ ] Pending",
  "- [x] Also done",
  "",
  "## Tables",
  "",
  "| Region | Q1   | Q2   | Q3   |",
  "|:------ | ---: | :--: | ---- |",
  "| North  |  120 | 140  | 160  |",
  "| South  |   90 | 110  | 130  |",
  "",
  "## Code block",
  "",
  "```typescript",
  "function hello(name: string): string {",
  "  return `Hello, $" + "{name}!`;",
  "}",
  "```",
  "",
  "## Image",
  "",
  '![placeholder](logo.png "Tooltip text")',
  "",
  "## Footnote",
  "",
  "Here is a fact[^1].",
  "",
  "[^1]: Source: World Bank 2024.",
  "",
  "## Edge cases",
  "",
  "- Empty paragraph follows:",
  "",
  "- (paragraph above is empty)",
  "- Mixed scripts: 你好 / مرحبا / שלום / 🎉 / *italic & bold* combined: ***triple***",
  "- Backslash escapes: \\*not italic\\* and \\`not code\\`",
  "",
  "---",
  "",
  "End of markdown.",
  ""
];
const md = lines.join("\n");

const importedDoc = await markdownToDocx(md, {
  defaultFont: "Calibri",
  defaultFontSize: 22,
  resolveImage(url, alt) {
    void alt;
    if (url === "logo.png") {
      return {
        data: tinyPng,
        mediaType: "png",
        width: Units.cmToEmu(2),
        height: Units.cmToEmu(2)
      };
    }
    return undefined;
  }
});
const buf = await Io.toBuffer(importedDoc);
fs.writeFileSync(path.join(outDir, "26-md-imported.docx"), buf);
console.log(`  markdownToDocx → ${importedDoc.body.length} body items, ${buf.length} bytes`);
console.log(
  `  embedded images: ${importedDoc.images?.length ?? 0}, footnotes: ${importedDoc.footnotes?.length ?? 0}`
);
console.log(`  → 26-md-imported.docx`);

// ---------------------------------------------------------------------------
// 2. DOCX → Markdown round-trip (re-read the just-written DOCX so the
//    pipeline goes Markdown → DOCX → file → DOCX → Markdown)
// ---------------------------------------------------------------------------
const reread = await Io.read(buf);
const mdRoundtrip = renderToMarkdown(reread, {
  headingStyle: "atx",
  includeImages: true,
  includeNotes: true
});
fs.writeFileSync(path.join(outDir, "26-md-roundtrip.md"), mdRoundtrip);
console.log(`  → 26-md-roundtrip.md (${mdRoundtrip.length} chars)`);

// ---------------------------------------------------------------------------
// 3. Setext heading style
// ---------------------------------------------------------------------------
const mdSetext = renderToMarkdown(reread, { headingStyle: "setext" });
fs.writeFileSync(path.join(outDir, "26-md-setext.md"), mdSetext);
console.log(`  → 26-md-setext.md`);

// ---------------------------------------------------------------------------
// 4. Build a fresh DOCX with mixed-format content, then export to Markdown
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Hand-built document", 1);
  Document.addParagraph(d, "Plain paragraph.");
  Document.addBulletList(d, ["A", "B", "C"]);
  Document.addNumberedList(d, ["1st", "2nd", "3rd"]);
  Document.addTable(d, [
    ["Header 1", "Header 2"],
    ["1", "2"]
  ]);
  const md2 = renderToMarkdown(Document.build(d));
  fs.writeFileSync(path.join(outDir, "26-md-from-docx.md"), md2);
  console.log(`  → 26-md-from-docx.md`);
}

// ---------------------------------------------------------------------------
// 5. markdownToDocxBody — same parser, returns body content PLUS the
//    supporting numbering / footnote / image definitions it references.
//    Useful when you want to splice Markdown content into an existing
//    DOCX you've already built (templates, headers/footers, etc.) rather
//    than producing a stand-alone document.
//
//    The result's `body` references document-level definitions (numbering for
//    lists, FootnoteDef for footnotes, ImageDef for images) returned in the
//    same result object. When splicing into a host document you must merge
//    those definitions too; here the fragment is kept to paragraphs +
//    headings + inline formatting, so only the body needs splicing.
// ---------------------------------------------------------------------------
{
  const fragmentMd = [
    "## Imported fragment",
    "",
    "First paragraph with **bold** and *italic* text.",
    "",
    "Second paragraph with [a link](https://example.com) inline.",
    "",
    "### Sub-heading",
    "",
    "Final paragraph — end of fragment."
  ].join("\n");
  const fragment = await markdownToDocxBody(fragmentMd);

  const host = Document.create();
  Document.useDefaultStyles(host);
  Document.addHeading(host, "Host document", 1);
  Document.addParagraph(host, "Below is content spliced in from Markdown:");
  for (const item of fragment.body) {
    Document.addContent(host, item);
  }
  Document.addParagraph(host, "(End of imported fragment.)");
  const buf3 = await Io.toBuffer(Document.build(host));
  fs.writeFileSync(path.join(outDir, "26-md-fragment-spliced.docx"), buf3);
  console.log(`  markdownToDocxBody → ${fragment.body.length} body items spliced into host doc`);
  console.log(`  → 26-md-fragment-spliced.docx`);
}
