/**
 * Word Example 02 — Text Formatting
 *
 * Covers run-level formatting and its edge cases:
 *   - Bold / italic / underline (every underline style)
 *   - Strikethrough / double strikethrough
 *   - Caps / smallCaps / hidden (vanish)
 *   - Color (RGB hex, "auto", theme color)
 *   - Background highlight + custom shading
 *   - Font size, family (ASCII / EastAsia / CS), font hint
 *   - Superscript / subscript
 *   - Character spacing, kerning, scale, position
 *   - Border around a run
 *   - Outline / shadow / emboss / imprint
 *   - Emphasis mark (East Asian)
 *   - Symbol (Wingdings)
 *   - Mixed nested runs in one paragraph
 *   - Edge case: zero-length / repeated identical runs / empty RunProperties
 *
 * Output: tmp/word-examples/02-formatting.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  text,
  bold,
  italic,
  underline,
  strikethrough,
  symbol,
  ptToHalfPoint,
  ptToTwips,
  toBuffer
} from "../index";
import type { UnderlineStyle, HighlightColor, RunProperties, Run } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word Module — Text Formatting", 1);

// ---------------------------------------------------------------------------
// Basic toggles (helper builders)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Basic toggles", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Plain, "),
    bold("bold, "),
    italic("italic, "),
    underline("underlined, "),
    strikethrough("struck-through.")
  ])
);

// Combined: bold + italic + underline at once via text()
Document.addParagraphElement(
  doc,
  paragraph([
    text("Combined: "),
    text("bold + italic + underline", { bold: true, italic: true, underline: "single" })
  ])
);

// ---------------------------------------------------------------------------
// Every underline style
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Every underline style", 2);
const allUnderlineStyles: UnderlineStyle[] = [
  "single",
  "words",
  "double",
  "thick",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashedHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dashDotHeavy",
  "dotDotDash",
  "dashDotDotHeavy",
  "wave",
  "wavyHeavy",
  "wavyDouble",
  "none"
];
Document.addParagraphElement(
  doc,
  paragraph(
    allUnderlineStyles.flatMap(s => [text(s, { underline: s }), text("  ")] as const) as Run[]
  )
);

// Underline with explicit color (independent of text color)
Document.addParagraphElement(
  doc,
  paragraph([
    text("Red wavy underline on black text: "),
    text("attention", { underline: { style: "wave", color: "FF0000" } })
  ])
);

// ---------------------------------------------------------------------------
// Color
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Color", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("RGB ", { color: "1F4E79" }),
    text("named\u00A0", { color: { val: "1F4E79", themeColor: "accent1" } }),
    text("auto", { color: "auto" })
  ])
);

// ---------------------------------------------------------------------------
// Highlight + custom shading
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Highlight & shading", 2);
const highlights: HighlightColor[] = [
  "yellow",
  "green",
  "cyan",
  "magenta",
  "red",
  "blue",
  "darkRed",
  "darkGreen",
  "lightGray",
  "white",
  "none"
];
Document.addParagraphElement(
  doc,
  paragraph(
    highlights.flatMap(h => [
      text(`${h} `, { highlight: h, color: h === "white" ? "000000" : undefined })
    ]) as Run[]
  )
);
// Custom shading (background fill, more flexible than highlight)
Document.addParagraphElement(
  doc,
  paragraph([
    text("Custom shading via shading.fill: "),
    text(" salmon ", { shading: { fill: "FA8072", pattern: "clear" } }),
    text("  "),
    text(" diagonal ", { shading: { fill: "FFFFCC", pattern: "diagStripe", color: "808080" } })
  ])
);

// ---------------------------------------------------------------------------
// Size & font family
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Size & font family", 2);
for (const pt of [8, 10, 12, 14, 18, 24, 36]) {
  Document.addParagraphElement(
    doc,
    paragraph([text(`${pt}pt sample`, { size: ptToHalfPoint(pt) })])
  );
}
Document.addParagraphElement(
  doc,
  paragraph([
    text("Calibri ", { font: { ascii: "Calibri", hAnsi: "Calibri" } }),
    text("Times ", { font: { ascii: "Times New Roman", hAnsi: "Times New Roman" } }),
    text("Courier ", { font: "Courier New" }),
    text("CJK 中文(SimSun) ", { font: { eastAsia: "SimSun", hint: "eastAsia" } }),
    text("CS العربية", { font: { cs: "Traditional Arabic" }, complexScript: true })
  ])
);

// ---------------------------------------------------------------------------
// Sub / super script
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Super/Subscript", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("E = mc"),
    text("2", { vertAlign: "superscript" }),
    text("    H"),
    text("2", { vertAlign: "subscript" }),
    text("O")
  ])
);

// ---------------------------------------------------------------------------
// Caps / small caps / hidden
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Caps & visibility", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("All caps: ", { caps: true }),
    text("hello world", { caps: true }),
    text("  Small caps: ", { smallCaps: true }),
    text("hello world", { smallCaps: true })
  ])
);
Document.addParagraphElement(
  doc,
  paragraph([text("Visible "), text("[hidden run]", { vanish: true }), text(" — visible again.")])
);

// ---------------------------------------------------------------------------
// Spacing / kerning / scale / position / fitText
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Spacing & glyph metrics", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("S P A C E D ", { spacing: ptToTwips(2) }),
    text("kerned ", { kern: ptToHalfPoint(8) }),
    text("scaled200% ", { scale: 200 }),
    text("scaled50% ", { scale: 50 }),
    text("raised", { position: ptToHalfPoint(4) }),
    text(" lowered", { position: -ptToHalfPoint(4) })
  ])
);
Document.addParagraphElement(
  doc,
  paragraph([
    text("FitText: "),
    text("squeezed into 1.5\u201d", {
      fitText: { val: ptToTwips(108), id: 1 }
    })
  ])
);

// ---------------------------------------------------------------------------
// Decorative effects (outline, shadow, emboss, imprint)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Decorative effects", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Outline ", { outline: true, size: ptToHalfPoint(20), bold: true }),
    text("Shadow ", { shadow: true, size: ptToHalfPoint(20), bold: true }),
    text("Emboss ", { emboss: true, size: ptToHalfPoint(20), bold: true }),
    text("Imprint", { imprint: true, size: ptToHalfPoint(20), bold: true })
  ])
);

// ---------------------------------------------------------------------------
// Run-level border
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Character border", 2);
Document.addParagraphElement(
  doc,
  paragraph([
    text("Plain "),
    text("[bordered]", {
      border: { style: "single", size: 8, color: "C00000", space: 0 }
    }),
    text(" plain.")
  ])
);

// ---------------------------------------------------------------------------
// Emphasis mark (East Asian typographic dots)
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  paragraph([
    text("着重号:"),
    text("重要内容", {
      emphasisMark: "dot",
      font: { eastAsia: "SimSun", hint: "eastAsia" }
    })
  ])
);

// ---------------------------------------------------------------------------
// Symbol (Wingdings glyph by codepoint)
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  paragraph([
    text("Symbols (Wingdings): "),
    symbol("Wingdings", "F0FC"), // ✓
    text("  "),
    symbol("Wingdings", "F0FB"), // ✗
    text("  "),
    symbol("Wingdings", "F046") // ⌚
  ])
);

// ---------------------------------------------------------------------------
// Edge case 1: empty RunProperties object — should produce a clean run
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);
const emptyProps: RunProperties = {};
Document.addParagraphElement(doc, paragraph([text("Run with empty properties.", emptyProps)]));

// Edge case 2: zero-length text run (writer should still emit valid XML)
Document.addParagraphElement(
  doc,
  paragraph([text("Before["), text("", { bold: true }), text("]After (zero-length bold run).")])
);

// Edge case 3: many consecutive identical-formatted runs
Document.addParagraphElement(
  doc,
  paragraph(
    Array.from({ length: 20 }, (_, i) =>
      text(`r${i} `, { color: "0070C0", italic: i % 2 === 0 })
    ) as Run[]
  )
);

// Edge case 4: deeply mixed inline formatting
Document.addParagraphElement(
  doc,
  paragraph([
    text("Mixed: "),
    bold("B"),
    italic("I"),
    underline("U"),
    strikethrough("S"),
    text("…and "),
    text("RED-LARGE-BOLD-ITALIC-CAPS", {
      color: "C00000",
      size: ptToHalfPoint(16),
      bold: true,
      italic: true,
      caps: true
    })
  ])
);

const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "02-formatting.docx"), buf);
console.log(`  → 02-formatting.docx (${buf.length} bytes)`);
