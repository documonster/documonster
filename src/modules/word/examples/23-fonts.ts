/**
 * Word Example 23 — Font embedding & subsetting
 *
 * Covers:
 *   - Font.embed with ODTTF obfuscation (default — required by Word for OEM fonts)
 *   - Font.embed without obfuscation
 *   - Font.embedFamily with regular + bold + italic variants in one call
 *   - Font.addEmbedded: merging into the document model
 *   - Font.subset: shrink the embedded glyph set to characters actually used
 *   - Document.addFont: register a font name without embedding the binary
 *   - Font.obfuscate / deobfuscate / generateObfuscationKey: the ODTTF
 *     obfuscation primitives behind Font.embed
 *
 * Note: Real-world usage requires a valid TTF/OTF binary. This example uses
 * a deterministic fake TTF stub (starts with the TTF magic bytes 00 01 00 00)
 * so the API surface can be exercised without shipping a font file. The
 * resulting docx contains the obfuscated bytes — Word will reject it as a
 * malformed font, but the round-trip through the writer/reader still works.
 *
 * Output:
 *   - 23-fonts-embedded.docx
 *   - 23-fonts-subset.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Font } from "../index";
import type { FontDef } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// Synthetic TTF stub. Real fonts work the same way — pass the file bytes.
function fakeTtf(size = 200): Uint8Array {
  const data = new Uint8Array(size);
  data[0] = 0x00;
  data[1] = 0x01;
  data[2] = 0x00;
  data[3] = 0x00;
  for (let i = 4; i < size; i++) {
    data[i] = (i * 17) & 0xff;
  }
  return data;
}

// ---------------------------------------------------------------------------
// 1. Embed a single font (obfuscated) and a family with multiple variants
// ---------------------------------------------------------------------------
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "Font embedding demo", 1);
  Document.addParagraph(doc, "The text below uses fonts embedded inside the .docx.");
  Document.addParagraph(
    doc,
    Build.paragraph([Build.text("Hello in Custom font.", { font: "MyCustom" })])
  );
  Document.addParagraph(
    doc,
    Build.paragraph([
      Build.text("Regular ", { font: "FamilyA" }),
      Build.text("Bold ", { font: "FamilyA", bold: true }),
      Build.text("Italic ", { font: "FamilyA", italic: true }),
      Build.text("BoldItalic", { font: "FamilyA", bold: true, italic: true })
    ])
  );

  // Single font
  const single = Font.embed({
    name: "MyCustom",
    data: fakeTtf(),
    style: "regular",
    obfuscate: true,
    family: "swiss",
    pitch: "variable"
  });

  // Whole family in one call
  const family = Font.embedFamily("FamilyA", {
    regular: fakeTtf(180),
    bold: fakeTtf(190),
    italic: fakeTtf(195),
    boldItalic: fakeTtf(200)
  });

  // Merge into the model
  const withFonts = Font.addEmbedded(doc, [single, ...family]);
  const buf = await Document.toBuffer(withFonts);
  fs.writeFileSync(path.join(outDir, "23-fonts-embedded.docx"), buf);
  console.log(`  → 23-fonts-embedded.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 2. Subsetted embed (only the characters that appear in the doc)
// ---------------------------------------------------------------------------
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "Subset demo", 1);
  const visibleText = "Hello, world!";
  Document.addParagraph(doc, Build.paragraph([Build.text(visibleText, { font: "Subsetted" })]));

  const fullData = fakeTtf(400);
  const trimmed = Font.subset(fullData, visibleText);
  console.log(`  full size: ${fullData.length} bytes, subsetted size: ${trimmed.length} bytes`);

  const single = Font.embed({
    name: "Subsetted",
    data: trimmed,
    style: "regular",
    usedCharacters: visibleText
  });

  const withFonts = Font.addEmbedded(doc, [single]);
  const buf = await Document.toBuffer(withFonts);
  fs.writeFileSync(path.join(outDir, "23-fonts-subset.docx"), buf);
  console.log(`  → 23-fonts-subset.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 3. Register a font *name* without embedding (font.xml only)
// ---------------------------------------------------------------------------
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "Reference-only font", 1);
  Document.addParagraph(
    doc,
    "If the user has 'CustomMono' installed locally, Word will use it. Otherwise it falls back."
  );
  const fontDef: FontDef = {
    name: "CustomMono",
    family: "modern",
    pitch: "fixed",
    panose1: "020B0609000000000000"
  };
  Document.addFont(doc, fontDef);
  Document.addParagraph(
    doc,
    Build.paragraph([Build.text("Sample monospace text.", { font: "CustomMono" })])
  );

  const buf = await Document.toBuffer(doc);
  fs.writeFileSync(path.join(outDir, "23-fonts-reference-only.docx"), buf);
  console.log(`  → 23-fonts-reference-only.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 4. Low-level ODTTF obfuscation primitives (applied by Font.embed when
//    `obfuscate: true`). Useful for round-tripping fonts that were written by
//    another tool, or for verifying our own output.
// ---------------------------------------------------------------------------
{
  const original = fakeTtf(64);
  const key = Font.generateObfuscationKey();
  console.log(`\n  Font.generateObfuscationKey() → ${key}`);

  const obf = Font.obfuscate(original, key);
  // The first 32 bytes must differ; the rest must be unchanged.
  const firstHalfChanged = obf.slice(0, 32).some((b, i) => b !== original[i]);
  const secondHalfSame = obf.slice(32).every((b, i) => b === original[i + 32]);
  console.log(
    `  obfuscateFont: first 32 bytes changed=${firstHalfChanged}, tail unchanged=${secondHalfSame}`
  );

  const round = Font.deobfuscate(obf, key);
  const isRoundTrip = round.every((b, i) => b === original[i]);
  console.log(`  deobfuscateFont round-trip identical: ${isRoundTrip}`);

  // Edge case: applying the wrong key produces garbled output (still 64 bytes
  // but with corrupt header).
  const wrongKey = Font.generateObfuscationKey();
  const wrong = Font.deobfuscate(obf, wrongKey);
  const headerMatches = [0x00, 0x01, 0x00, 0x00].every((b, i) => wrong[i] === b);
  console.log(`  deobfuscate with wrong key produces valid TTF magic: ${headerMatches}`);
}
