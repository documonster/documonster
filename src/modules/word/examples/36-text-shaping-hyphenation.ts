/**
 * Word Example 36 — Text shaping & hyphenation
 *
 * Covers:
 *   - detectScript / detectDirection — quick classification of Unicode strings
 *   - shapeText — apply Arabic joining + BiDi reordering, returns visual-order
 *     glyph clusters
 *   - createHyphenator + hyphenateWord + hyphenateText with the bundled
 *     ENGLISH_US_PATTERNS (no external corpus required)
 *
 * These are CPU-bound utilities meant for layout pipelines; they don't
 * produce a .docx by themselves.  We embed the result of hyphenation into a
 * docx so the soft hyphens are visible if Word/LibreOffice need to break
 * lines.
 *
 * Output: tmp/word-examples/36-shaping.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  text,
  paragraph,
  toBuffer,
  shapeText,
  detectScript,
  detectDirection,
  createHyphenator,
  hyphenateWord,
  hyphenateText,
  ENGLISH_US_PATTERNS
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. detectScript / detectDirection
// ---------------------------------------------------------------------------
const samples: { name: string; value: string }[] = [
  { name: "Latin", value: "Hello, World" },
  { name: "Arabic", value: "مرحبا بالعالم" },
  { name: "Hebrew", value: "שלום עולם" },
  { name: "CJK", value: "你好,世界" },
  { name: "Hiragana", value: "ひらがな" },
  { name: "Mixed Latin+CJK", value: "Hello 世界!" },
  { name: "Mixed Latin+RTL", value: "Order #42 (طلب)" }
];
console.log("  Script / direction detection:");
for (const { name, value } of samples) {
  console.log(
    `    ${name.padEnd(18)} script=${detectScript(value).padEnd(8)} dir=${detectDirection(value)}`
  );
}

// ---------------------------------------------------------------------------
// 2. shapeText — visual-order clusters with Arabic joining
// ---------------------------------------------------------------------------
console.log("\n  Shape Arabic mixed with Latin:");
const arabicShaped = shapeText("Order طلب 42", { direction: "ltr" });
for (const c of arabicShaped) {
  console.log(
    `    chars=${JSON.stringify(c.chars)} visual=${JSON.stringify(c.visual)} script=${c.script} dir=${c.direction}`
  );
}

// Same with rtl base direction — the visual reordering changes
console.log("\n  Same string, rtl base direction:");
const arabicShapedRtl = shapeText("Order طلب 42", { direction: "rtl" });
for (const c of arabicShapedRtl) {
  console.log(`    chars=${JSON.stringify(c.chars)} visual=${JSON.stringify(c.visual)}`);
}

// ---------------------------------------------------------------------------
// 3. Hyphenation
// ---------------------------------------------------------------------------
const hyph = createHyphenator(ENGLISH_US_PATTERNS, { minLeft: 2, minRight: 3 });

console.log("\n  Hyphenation points:");
const words = [
  "association",
  "computer",
  "international",
  "extraordinary",
  "supercalifragilistic",
  "of",
  "abc"
];
for (const w of words) {
  console.log(`    ${w.padEnd(28)} points=${JSON.stringify(hyph(w))}`);
}

console.log("\n  hyphenateWord (default soft hyphen):");
for (const w of words) {
  // Print with visible "·" instead of soft hyphen so the result is readable
  const result = hyphenateWord(w, hyph).replace(/\u00AD/g, "·");
  console.log(`    ${w.padEnd(28)} → ${result}`);
}

console.log("\n  hyphenateText paragraph:");
const sample =
  "The international association of professional consultants demonstrated extraordinary intelligence.";
const hyphenated = hyphenateText(sample, hyph);
console.log(`    visible: ${hyphenated.replace(/\u00AD/g, "·")}`);

// ---------------------------------------------------------------------------
// 4. Embed the hyphenated paragraph in a .docx so Word can use the soft
//    hyphens for line-breaking
// ---------------------------------------------------------------------------
const doc = Document.create();
Document.useDefaultStyles(doc);
Document.addHeading(doc, "Hyphenated paragraph (with soft hyphens)", 1);
Document.addParagraphElement(doc, paragraph([text(hyphenated)]));
Document.addHeading(doc, "Same paragraph without hyphenation", 2);
Document.addParagraphElement(doc, paragraph([text(sample)]));

const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "36-shaping.docx"), buf);
console.log(`\n  → 36-shaping.docx (${buf.length} bytes)`);
