/**
 * Word Example 33 — Incremental editing
 *
 * Covers:
 *   - listDocxParts — enumerate every part in a .docx without parsing
 *   - readDocxPart — extract one part as raw bytes
 *   - editDocxIncremental with each edit type:
 *     · replaceBody (most common) — keeps section properties, swaps body
 *     · replacePart — replace one binary part (e.g. an embedded image)
 *     · replacePartText — replace one XML part (e.g. core props)
 *     · deletePart — remove a part (must be one with no rels)
 *     · replaceHeader / replaceFooter — rewrite an existing header/footer
 *
 * This is the fastest path for "open a docx, change a few things, save"
 * because unchanged parts are passed through without parsing.
 *
 * Output: tmp/word-examples/33-incremental/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  textParagraph,
  bold,
  toBuffer,
  listDocxParts,
  readDocxPart,
  editDocxIncremental
} from "../index";
import type { BodyContent } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/33-incremental"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Build a base document with header/footer & metadata
// ---------------------------------------------------------------------------
const seed = Document.create();
Document.useDefaultStyles(seed);
Document.addHeading(seed, "Original title", 1);
Document.addParagraph(seed, "First paragraph.");
Document.addParagraph(seed, "Second paragraph.");
Document.setHeader(seed, "default", { children: [textParagraph("Header v1")] });
Document.setFooter(seed, "default", { children: [textParagraph("Footer v1")] });
Document.setCoreProperties(seed, { title: "Original", creator: "OpenCode" });
const baseBytes = await toBuffer(Document.build(seed));
fs.writeFileSync(path.join(outDir, "00-base.docx"), baseBytes);
console.log(`  base.docx: ${baseBytes.length} bytes`);

// ---------------------------------------------------------------------------
// 2. listDocxParts — peek at the package without parsing
// ---------------------------------------------------------------------------
const parts = await listDocxParts(baseBytes);
console.log(`  ${parts.length} parts:`);
for (const p of parts) {
  console.log(`    · ${p}`);
}

// ---------------------------------------------------------------------------
// 3. readDocxPart — read one part's bytes
// ---------------------------------------------------------------------------
const docXmlBytes = await readDocxPart(baseBytes, "word/document.xml");
console.log(`  word/document.xml: ${docXmlBytes?.length ?? 0} bytes`);
const headerBytes = await readDocxPart(baseBytes, "word/header1.xml");
console.log(`  word/header1.xml:  ${headerBytes?.length ?? 0} bytes`);
const missing = await readDocxPart(baseBytes, "word/no-such-part.xml");
console.log(`  no-such-part:      ${missing === undefined ? "undefined ✓" : "unexpected"}`);

// ---------------------------------------------------------------------------
// 4. Edit body in place — fastest path
// ---------------------------------------------------------------------------
const newBody: BodyContent[] = [
  paragraph([bold("Edited title — body replaced")], { style: "Heading1" }),
  textParagraph("Brand-new first paragraph."),
  textParagraph("Brand-new second paragraph.")
];
const edited1 = await editDocxIncremental(baseBytes, [{ type: "replaceBody", body: newBody }]);
fs.writeFileSync(path.join(outDir, "01-replaced-body.docx"), edited1);
console.log(`  → 01-replaced-body.docx (${edited1.length} bytes)`);

// ---------------------------------------------------------------------------
// 5. Replace header & footer simultaneously
// ---------------------------------------------------------------------------
const edited2 = await editDocxIncremental(baseBytes, [
  {
    type: "replaceHeader",
    path: "word/header1.xml",
    children: [textParagraph("Header v2 — incrementally edited")]
  },
  {
    type: "replaceFooter",
    path: "word/footer1.xml",
    children: [textParagraph("Footer v2 — incrementally edited")]
  }
]);
fs.writeFileSync(path.join(outDir, "02-replaced-header-footer.docx"), edited2);
console.log(`  → 02-replaced-header-footer.docx (${edited2.length} bytes)`);

// ---------------------------------------------------------------------------
// 6. replacePartText — rewrite docProps/core.xml directly with new metadata
// ---------------------------------------------------------------------------
const newCoreXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Patched title</dc:title>
  <dc:creator>Patched creator</dc:creator>
  <cp:lastModifiedBy>Patched modifier</cp:lastModifiedBy>
</cp:coreProperties>`;
const edited3 = await editDocxIncremental(baseBytes, [
  { type: "replacePartText", path: "docProps/core.xml", text: newCoreXml }
]);
fs.writeFileSync(path.join(outDir, "03-replaced-core-props.docx"), edited3);
console.log(`  → 03-replaced-core-props.docx (${edited3.length} bytes)`);

// ---------------------------------------------------------------------------
// 7. Combine multiple edits in one round-trip
// ---------------------------------------------------------------------------
const edited4 = await editDocxIncremental(baseBytes, [
  { type: "replaceBody", body: newBody },
  {
    type: "replaceHeader",
    path: "word/header1.xml",
    children: [textParagraph("Combined edit header")]
  },
  { type: "replacePartText", path: "docProps/core.xml", text: newCoreXml }
]);
fs.writeFileSync(path.join(outDir, "04-combined.docx"), edited4);
console.log(`  → 04-combined.docx (${edited4.length} bytes)`);

// ---------------------------------------------------------------------------
// 8. Replace a binary/XML part — demonstrate `replacePart` on an existing
//    part. We rewrite `word/settings.xml` to enable an additional setting,
//    keeping the part's existing content type so [Content_Types].xml
//    stays consistent.
//
//    NOTE: editDocxIncremental does NOT update [Content_Types].xml for
//    you. Adding a brand-new part type (e.g. injecting docProps/custom.xml
//    when the base doc has none) requires the caller to also patch the
//    Content Types — typically easier with the bulk packager workflow.
// ---------------------------------------------------------------------------
const newSettingsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="125"/>
</w:settings>`;
const edited5 = await editDocxIncremental(baseBytes, [
  {
    type: "replacePartText",
    path: "word/settings.xml",
    text: newSettingsXml
  }
]);
fs.writeFileSync(path.join(outDir, "05-replaced-settings.docx"), edited5);
console.log(`  → 05-replaced-settings.docx (${edited5.length} bytes)`);

// Edge case: list every part to confirm no new file types were added
const after = await listDocxParts(edited5);
console.log(`  parts after edit: ${after.length}`);
