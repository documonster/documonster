/**
 * Word Example 01 — Basics
 *
 * Covers the fundamental document building blocks and their edge cases:
 *   - Empty / minimum-sized document
 *   - Plain paragraphs, headings (1-9)
 *   - Page breaks & line breaks
 *   - Tabs (regular and positional)
 *   - Unicode (CJK / emoji / RTL)
 *   - Very long paragraph
 *   - Special characters that must be XML-escaped (<, >, &, ", ', NBSP, control chars)
 *   - Soft hyphen / no-break hyphen / carriage return
 *   - Empty paragraph (used as a vertical spacer)
 *
 * Output: tmp/word-examples/01-basics.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  text,
  tab,
  pageBreak,
  lineBreak,
  positionalTab,
  noBreakHyphen,
  softHyphen,
  carriageReturn,
  heading,
  toBuffer,
  // Inverse unit conversions — useful when round-tripping values out of OOXML
  twipsToInches,
  twipsToCm,
  twipsToPt,
  emuToInches,
  emuToCm,
  emuToPx,
  halfPointToPt,
  eighthPointToPt,
  spacingToLineMultiplier,
  tablePctToPercent,
  inchesToTwips,
  cmToTwips,
  ptToTwips,
  inchesToEmu,
  cmToEmu,
  pxToEmu,
  ptToEmu,
  ptToHalfPoint,
  ptToEighthPoint,
  lineMultiplierToSpacing,
  percentToTablePct,
  mmToTwips
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Edge case 1: an empty document still produces a valid .docx (Word requires
// at least one paragraph in document.xml — Document.build() handles that for
// callers, so building an empty handle is a no-op and the writer adds the
// implicit empty paragraph).
// ---------------------------------------------------------------------------
const emptyDoc = Document.create();
const emptyBuf = await toBuffer(Document.build(emptyDoc));
fs.writeFileSync(path.join(outDir, "01-basics-empty.docx"), emptyBuf);
console.log(`  → 01-basics-empty.docx (${emptyBuf.length} bytes)`);

// ---------------------------------------------------------------------------
// Main basics document
// ---------------------------------------------------------------------------
const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word Module — Basics", 1);

// Plain paragraph
Document.addParagraph(doc, "A plain paragraph using Document.addParagraph().");

// Empty paragraph as vertical spacer
Document.addParagraph(doc, "");

// All heading levels
for (const lvl of [1, 2, 3, 4, 5, 6, 7, 8, 9] as const) {
  Document.addHeading(doc, `Heading level ${lvl}`, lvl);
}

// Same thing through the standalone heading() builder — which produces a
// fully-formed Paragraph object you can compose, store or transform before
// inserting. Useful when generating headings programmatically.
for (const lvl of [1, 2, 3] as const) {
  Document.addParagraphElement(doc, heading(`Built via heading() — level ${lvl}`, lvl));
}

// Tabs and positional tabs
Document.addHeading(doc, "Tabs & breaks", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Left"),
    tab(),
    text("Center"),
    tab(),
    text("Right (with regular tabs — relies on default tab stops).")
  ])
);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Positional tab → "),
    positionalTab({ alignment: "right", relativeTo: "margin", leader: "dot" }),
    text("end")
  ])
);

// Line break vs paragraph break
Document.addParagraphElement(
  doc,
  paragraph([
    text("Line one"),
    lineBreak(),
    text("Line two (same paragraph, soft return)"),
    lineBreak(),
    text("Line three")
  ])
);

// Carriage return (treated like a line break by Word)
Document.addParagraphElement(
  doc,
  paragraph([text("CR before:"), carriageReturn(), text("CR after.")])
);

// Page break inside a run
Document.addParagraph(doc, "Below this paragraph a page break is inserted.");
Document.addParagraphElement(doc, paragraph([pageBreak(), text("This is on a new page.")]));

// ---------------------------------------------------------------------------
// Edge cases — special chars, Unicode
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);

// XML-sensitive characters
Document.addParagraph(
  doc,
  `XML special chars must be escaped by the writer: < > & " '   <p>tag-like</p> &amp; "smart quotes" \u2018'\u2019`
);

// NBSP and other whitespace
Document.addParagraph(
  doc,
  "NBSP\u00A0between\u00A0words; thin\u2009space; em\u2003space; ZWJ\u200dhere; RLM\u200fhere."
);

// Mix of scripts: Latin / CJK / Cyrillic / Arabic / Hebrew
Document.addParagraph(
  doc,
  "Multi-script: Hello / 你好世界 / Привет / مرحبا / שלום / こんにちは / 안녕하세요."
);

// Emoji (the surrogate pairs must round-trip through the writer)
Document.addParagraph(doc, "Emoji: 😀 🚀 🇨🇳 🇺🇸 🧑‍💻 (ZWJ sequence) ⚙️");

// RTL paragraph (set bidi flag so layout engines treat it accordingly)
Document.addParagraphElement(
  doc,
  paragraph([text("هذا نص عربي RTL مع كلمة latin بداخلها.")], { bidi: true })
);

// Long single paragraph (5,000 chars) — exercises the writer's text node sizing
const long = "Lorem ipsum dolor sit amet. ".repeat(200);
Document.addParagraph(doc, `Long paragraph (${long.length} chars): ${long}`);

// noBreakHyphen / softHyphen
Document.addParagraphElement(
  doc,
  paragraph([
    text("part-one"),
    noBreakHyphen(),
    text("part-two; soft\u00ADhyphenated word built explicitly: super"),
    softHyphen(),
    text("califragilistic"),
    softHyphen(),
    text("expialidocious.")
  ])
);

// Whitespace-only run + control-character run.  Control chars (0x00-0x08,
// 0x0B, 0x0C, 0x0E-0x1F) are illegal in XML; the writer drops them, so the
// example above tests that we don't crash on accidental input.
Document.addParagraph(doc, "Trailing spaces follow →   ");
Document.addParagraph(doc, "Control char dropped: [\u0007] bell, [\u0001] SOH.");

// ---------------------------------------------------------------------------
// Unit conversion helpers — every forward / inverse pair must round-trip
// inside its declared precision so callers can freely move between OOXML's
// native units (twips, EMU, half-points, eighth-points, line-spacing 240ths,
// table-pct fiftieths) and human-friendly units (inches, cm, mm, pt, px,
// percent). The values below are the assertions that keep the helpers honest
// for downstream consumers.
// ---------------------------------------------------------------------------
{
  const checks: [string, number, number][] = [
    ["inchesToTwips(2) == 2880", inchesToTwips(2), 2880],
    ["twipsToInches(2880) == 2", twipsToInches(2880), 2],
    ["cmToTwips(2.54) ≈ inchesToTwips(1)", cmToTwips(2.54), inchesToTwips(1)],
    ["twipsToCm(1440) ≈ 2.54", Number(twipsToCm(1440).toFixed(2)), 2.54],
    ["ptToTwips(12) == 240", ptToTwips(12), 240],
    ["twipsToPt(240) == 12", twipsToPt(240), 12],
    ["mmToTwips(254) == cmToTwips(25.4)", mmToTwips(254), cmToTwips(25.4)],
    ["inchesToEmu(1) == 914400", inchesToEmu(1), 914_400],
    ["emuToInches(914400) == 1", emuToInches(914_400), 1],
    ["cmToEmu(2.54) == inchesToEmu(1)", cmToEmu(2.54), inchesToEmu(1)],
    ["emuToCm(914400) ≈ 2.54", Number(emuToCm(914_400).toFixed(2)), 2.54],
    ["ptToEmu(72) == inchesToEmu(1)", ptToEmu(72), inchesToEmu(1)],
    ["pxToEmu(96) == inchesToEmu(1)", pxToEmu(96), inchesToEmu(1)],
    ["emuToPx(914400) == 96", emuToPx(914_400), 96],
    ["ptToHalfPoint(12) == 24", ptToHalfPoint(12), 24],
    ["halfPointToPt(24) == 12", halfPointToPt(24), 12],
    ["ptToEighthPoint(1.5) == 12", ptToEighthPoint(1.5), 12],
    ["eighthPointToPt(12) == 1.5", eighthPointToPt(12), 1.5],
    ["lineMultiplierToSpacing(1.5) == 360", lineMultiplierToSpacing(1.5), 360],
    ["spacingToLineMultiplier(360) == 1.5", spacingToLineMultiplier(360), 1.5],
    ["percentToTablePct(75) == 3750", percentToTablePct(75), 3750],
    ["tablePctToPercent(3750) == 75", tablePctToPercent(3750), 75]
  ];
  for (const [name, got, want] of checks) {
    if (Math.abs(got - want) > 0.01) {
      throw new Error(`Unit-conversion regression: ${name} → got ${got}, want ${want}`);
    }
  }
  console.log(`  unit-conversion round-trips: ${checks.length} checks ✓`);
}

// ---------------------------------------------------------------------------
// Build & write
// ---------------------------------------------------------------------------
const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "01-basics.docx"), buf);
console.log(`  → 01-basics.docx (${buf.length} bytes)`);
