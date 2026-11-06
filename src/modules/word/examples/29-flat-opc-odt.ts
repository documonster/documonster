/**
 * Word Example 29 — Flat OPC & ODT
 *
 * Covers:
 *   - Document.toFlatOpc: serialize a DocxDocument as a single XML file.
 *   - Document.read reads it back as a document; Convert.flatOpcToParts / isFlatOpc
 *     give its parts.
 *   - Convert.fromOdt / writeOdt: round-trip an OpenDocument Text document.
 *   - Edge case: feeding a non-Flat-OPC string to isFlatOpc.
 *
 * Output: tmp/word-examples/29-formats/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Convert } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/29-formats"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Build a small document we can convert through every format
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);
Document.addHeading(d, "Format round-trip", 1);
Document.addParagraph(d, "This document will be serialized to multiple formats.");
Document.addParagraph(
  d,
  Build.paragraph([
    Build.text("Mixed run: "),
    Build.bold("bold"),
    Build.text(" + "),
    Build.text("color", { color: "C00000" })
  ])
);
Document.addBulletList(d, ["Foo", "Bar", "Baz"]);

// ---------------------------------------------------------------------------
// 1. Flat OPC
// ---------------------------------------------------------------------------
{
  const flatXml = await Document.toFlatOpc(d);
  console.log(`  Flat OPC output: ${flatXml.length} characters`);
  console.log(`  starts with: ${flatXml.slice(0, 80)}`);
  fs.writeFileSync(path.join(outDir, "01-document.xml"), flatXml);
  console.log(`  → 01-document.xml`);

  console.log(`  isFlatOpc(flatXml): ${Convert.isFlatOpc(flatXml)}`);
  console.log(`  isFlatOpc("hello"): ${Convert.isFlatOpc("hello")}`);

  // Document.read takes the Flat OPC bytes as it takes a .docx's.
  const readBack = await Document.read(new TextEncoder().encode(flatXml));
  console.log(`  Document.read(Flat OPC): ${readBack.body.length} body blocks`);

  // Re-parse — produces the same Map of part path → bytes that the writer
  // would emit into a ZIP.
  const partsBack = Convert.flatOpcToParts(flatXml);
  console.log(`  re-parsed Flat OPC: ${partsBack.size} parts`);
  for (const partName of [...partsBack.keys()].slice(0, 6)) {
    console.log(`    · ${partName}`);
  }

  // Round-trip Flat OPC → its own XML again (lossless serialization)
  const reSerialized = Convert.partsToFlatOpc(partsBack);
  console.log(`  re-serialized length: ${reSerialized.length}`);
}

// ---------------------------------------------------------------------------
// 2. ODT (OpenDocument Text)
// ---------------------------------------------------------------------------
{
  const odtBytes = await Convert.toOdt(d);
  fs.writeFileSync(path.join(outDir, "02-document.odt"), odtBytes);
  console.log(`  → 02-document.odt (${odtBytes.length} bytes)`);

  // Round-trip: read the ODT back
  const reread = await Convert.fromOdt(odtBytes);
  console.log(`  ODT re-read body length: ${reread.body.length}`);

  // Convert the round-tripped doc back to .docx
  const buf = await Document.toBuffer(reread);
  fs.writeFileSync(path.join(outDir, "03-from-odt.docx"), buf);
  console.log(`  → 03-from-odt.docx`);
}

// ---------------------------------------------------------------------------
// 3. Sanity: re-read the original .docx so the example covers the most
//    common path in the same file.
// ---------------------------------------------------------------------------
{
  const docxBytes = await Document.toBuffer(d);
  fs.writeFileSync(path.join(outDir, "04-original.docx"), docxBytes);
  const reread = await Document.read(docxBytes);
  console.log(`  re-read original .docx body length: ${reread.body.length}`);
}
