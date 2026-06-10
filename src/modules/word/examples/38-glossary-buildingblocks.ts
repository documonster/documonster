/**
 * Word Example 38 — Glossary / Building Blocks (AutoText, Quick Parts)
 *
 * Building blocks are reusable snippets that Word's UI exposes as
 * "AutoText", "Quick Parts", "Cover Pages", etc.  This example shows how
 * to construct, query and serialise them.
 *
 * Assigning the assembled `GlossaryDocument` to `doc.glossary` makes the
 * packager serialise it to `word/glossary/document.xml`, register the
 * `glossaryDocument` relationship and add the `[Content_Types].xml`
 * override — the canonical OOXML location Word reads Quick Parts / AutoText
 * from. The same blocks are also inlined into the body so the assembled
 * letter is visible directly.
 *
 * Output: tmp/word-examples/38-glossary/...
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
  toBuffer,
  createBuildingBlock,
  createGlossaryDocument,
  findBuildingBlock,
  listBuildingBlocks,
  getAutoTextEntries,
  getQuickParts
} from "../index";
import type { BuildingBlock, BodyContent } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/38-glossary"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Build several building blocks
// ---------------------------------------------------------------------------
const greeting = createBuildingBlock(
  "FormalGreeting",
  "autoText",
  [textParagraph("Dear Sir/Madam,")],
  { category: "Letters", description: "Formal opening for a letter" }
);
const signoff = createBuildingBlock(
  "SignOff",
  "autoText",
  [textParagraph("Yours faithfully,"), textParagraph("(Signature)")],
  { category: "Letters", description: "Closing for a formal letter" }
);
const disclaimer = createBuildingBlock(
  "Disclaimer",
  "quickParts",
  [paragraph([bold("Disclaimer: "), text("This document is provided as-is.")])],
  { category: "Legal" }
);
const cover = createBuildingBlock(
  "TitleCover",
  "coverPages",
  [paragraph([text("REPORT TITLE")], { style: "Heading1" })],
  { category: "Reports" }
);

// ---------------------------------------------------------------------------
// 2. Compose into a glossary document and query
// ---------------------------------------------------------------------------
const glossary = createGlossaryDocument([greeting, signoff, disclaimer, cover]);

console.log(
  `  AutoText:     ${getAutoTextEntries(glossary)
    .map(b => b.name)
    .join(", ")}`
);
console.log(
  `  Quick Parts:  ${getQuickParts(glossary)
    .map(b => b.name)
    .join(", ")}`
);
console.log(
  `  Cover pages:  ${listBuildingBlocks(glossary, "coverPages")
    .map(b => b.name)
    .join(", ")}`
);

const found = findBuildingBlock(glossary, "Disclaimer");
console.log(`  findBuildingBlock("Disclaimer"): ${found?.gallery} / ${found?.category}`);
const notFound = findBuildingBlock(glossary, "DoesNotExist");
console.log(`  findBuildingBlock("DoesNotExist"): ${notFound}`);

// ---------------------------------------------------------------------------
// 3. Use one of the blocks as document content
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);
Document.addHeading(d, "Letter assembled from building blocks", 1);

const insertBlock = (block: BuildingBlock): void => {
  for (const item of block.content) {
    Document.addContent(d, item as BodyContent);
  }
};
insertBlock(greeting);
Document.addParagraph(d, "Body of the letter goes here. " + "Lorem ipsum… ".repeat(8));
insertBlock(signoff);
insertBlock(disclaimer);

// Attach the glossary to the document so the building blocks are *also*
// embedded at the canonical OOXML location (word/glossary/document.xml) —
// Word then exposes them as Quick Parts / AutoText, not just as inlined body
// content. The packager registers the glossaryDocument relationship + content
// type automatically.
const builtDoc = { ...Document.build(d), glossary };
const buf = await toBuffer(builtDoc);
fs.writeFileSync(path.join(outDir, "01-letter-from-blocks.docx"), buf);
console.log(`  → 01-letter-from-blocks.docx (${buf.length} bytes)`);

// ---------------------------------------------------------------------------
// 4. Edge case: building block with rich nested content
// ---------------------------------------------------------------------------
const richBlock = createBuildingBlock("RichSnippet", "quickParts", [
  paragraph([bold("Rich snippet:")], { style: "Heading2" }),
  textParagraph("First line of the snippet."),
  textParagraph("Second line.")
]);
console.log(
  `  rich block has ${richBlock.content.length} content blocks, GUID=${(richBlock.guid ?? "(none)").slice(0, 8)}…`
);
