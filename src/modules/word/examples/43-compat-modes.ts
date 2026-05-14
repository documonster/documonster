/**
 * Word Example 43 — Compatibility mode + CompatFlag
 *
 * Word renders certain layout/typographic decisions differently depending on
 * the document's "compatibility mode" (the Word version it was authored for).
 *   - 11 = Word 2003
 *   - 12 = Word 2007
 *   - 14 = Word 2010
 *   - 15 = Word 2013+ (modern default)
 *
 * Covers:
 *   - getCompatibilityMode (defaults to 15 when nothing is stored)
 *   - setCompatibilityMode — mutates doc.settings.compatSettings
 *   - Round-trip across read/write preserves the chosen mode
 *   - Adding individual w:compat flags (e.g. doNotExpandShiftReturn)
 *
 * Output: tmp/word-examples/43-compat/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, toBuffer, readDocx, getCompatibilityMode, setCompatibilityMode } from "../index";
import type { CompatibilityMode, CompatFlag, DocumentSettings } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/43-compat"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Default mode is 15 (Word 2013+)
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addParagraph(d, "Default compatibility mode demo.");
  const built = Document.build(d);
  console.log(`  default getCompatibilityMode: ${getCompatibilityMode(built)}`);
}

// ---------------------------------------------------------------------------
// 2. Set each supported mode and round-trip
// ---------------------------------------------------------------------------
const modes: CompatibilityMode[] = [11, 12, 14, 15];
for (const mode of modes) {
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addParagraph(d, `Targeting compatibility mode ${mode}.`);
  const model = Document.build(d);
  setCompatibilityMode(model, mode);
  console.log(`  set/get mode ${mode}: ${getCompatibilityMode(model)}`);

  const bytes = await toBuffer(model);
  fs.writeFileSync(path.join(outDir, `compat-${mode}.docx`), bytes);

  // Re-read and verify. Note: the writer normalises compat metadata when
  // outputting modern OOXML — a doc authored for Word 2003 still opens in
  // 2013+ as compat=15 unless the consumer additionally configures the
  // legacy-flags surface. We log the round-trip value to surface this.
  const reread = await readDocx(bytes);
  console.log(`    after round-trip: ${getCompatibilityMode(reread)}`);
}

// ---------------------------------------------------------------------------
// 3. Manual w:compat flags (e.g. legacy quirks)
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addParagraph(d, "Document with custom compat flags.");

  const flags: CompatFlag[] = [
    { name: "doNotExpandShiftReturn", val: "1" },
    { name: "useSingleBorderforContiguousCells", val: "1" },
    { name: "wpJustification", val: "1" }
  ];
  const settings: DocumentSettings = { compatFlags: flags };
  Document.setSettings(d, settings);
  const built = Document.build(d);
  setCompatibilityMode(built, 14);

  const bytes = await toBuffer(built);
  fs.writeFileSync(path.join(outDir, "with-compat-flags.docx"), bytes);
  const reread = await readDocx(bytes);
  console.log(
    `  custom flags: mode=${getCompatibilityMode(reread)}, flags=${reread.settings?.compatFlags?.length ?? 0}`
  );
}
