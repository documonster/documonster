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

import { Document, Build, Io, Units } from "../index";
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
  Build.paragraph([
    Build.text("Plain, "),
    Build.bold("bold, "),
    Build.italic("italic, "),
    Build.underline("underlined, "),
    Build.strikethrough("struck-through.")
  ])
);

// Combined: bold + italic + underline at once via text()
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Combined: "),
    Build.text("bold + italic + underline", { bold: true, italic: true, underline: "single" })
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
  Build.paragraph(
    allUnderlineStyles.flatMap(
      s => [Build.text(s, { underline: s }), Build.text("  ")] as const
    ) as Run[]
  )
);

// Underline with explicit color (independent of text color)
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Red wavy underline on black text: "),
    Build.text("attention", { underline: { style: "wave", color: "FF0000" } })
  ])
);

// ---------------------------------------------------------------------------
// Color
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Color", 2);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("RGB ", { color: "1F4E79" }),
    Build.text("named\u00A0", { color: { val: "1F4E79", themeColor: "accent1" } }),
    Build.text("auto", { color: "auto" })
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
  Build.paragraph(
    highlights.flatMap(h => [
      Build.text(`${h} `, { highlight: h, color: h === "white" ? "000000" : undefined })
    ]) as Run[]
  )
);
// Custom shading (background fill, more flexible than highlight)
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Custom shading via shading.fill: "),
    Build.text(" salmon ", { shading: { fill: "FA8072", pattern: "clear" } }),
    Build.text("  "),
    Build.text(" diagonal ", {
      shading: { fill: "FFFFCC", pattern: "diagStripe", color: "808080" }
    })
  ])
);

// ---------------------------------------------------------------------------
// Size & font family
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Size & font family", 2);
for (const pt of [8, 10, 12, 14, 18, 24, 36]) {
  Document.addParagraphElement(
    doc,
    Build.paragraph([Build.text(`${pt}pt sample`, { size: Units.ptToHalfPoint(pt) })])
  );
}
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Calibri ", { font: { ascii: "Calibri", hAnsi: "Calibri" } }),
    Build.text("Times ", { font: { ascii: "Times New Roman", hAnsi: "Times New Roman" } }),
    Build.text("Courier ", { font: "Courier New" }),
    Build.text("CJK 中文(SimSun) ", { font: { eastAsia: "SimSun", hint: "eastAsia" } }),
    Build.text("CS العربية", { font: { cs: "Traditional Arabic" }, complexScript: true })
  ])
);

// ---------------------------------------------------------------------------
// Sub / super script
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Super/Subscript", 2);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("E = mc"),
    Build.text("2", { vertAlign: "superscript" }),
    Build.text("    H"),
    Build.text("2", { vertAlign: "subscript" }),
    Build.text("O")
  ])
);

// ---------------------------------------------------------------------------
// Caps / small caps / hidden
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Caps & visibility", 2);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("All caps: ", { caps: true }),
    Build.text("hello world", { caps: true }),
    Build.text("  Small caps: ", { smallCaps: true }),
    Build.text("hello world", { smallCaps: true })
  ])
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Visible "),
    Build.text("[hidden run]", { vanish: true }),
    Build.text(" — visible again.")
  ])
);

// ---------------------------------------------------------------------------
// Spacing / kerning / scale / position / fitText
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Spacing & glyph metrics", 2);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("S P A C E D ", { spacing: Units.ptToTwips(2) }),
    Build.text("kerned ", { kern: Units.ptToHalfPoint(8) }),
    Build.text("scaled200% ", { scale: 200 }),
    Build.text("scaled50% ", { scale: 50 }),
    Build.text("raised", { position: Units.ptToHalfPoint(4) }),
    Build.text(" lowered", { position: -Units.ptToHalfPoint(4) })
  ])
);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("FitText: "),
    Build.text("squeezed into 1.5\u201d", {
      fitText: { val: Units.ptToTwips(108), id: 1 }
    })
  ])
);

// ---------------------------------------------------------------------------
// Decorative effects (outline, shadow, emboss, imprint)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Decorative effects", 2);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Outline ", { outline: true, size: Units.ptToHalfPoint(20), bold: true }),
    Build.text("Shadow ", { shadow: true, size: Units.ptToHalfPoint(20), bold: true }),
    Build.text("Emboss ", { emboss: true, size: Units.ptToHalfPoint(20), bold: true }),
    Build.text("Imprint", { imprint: true, size: Units.ptToHalfPoint(20), bold: true })
  ])
);

// ---------------------------------------------------------------------------
// Run-level border
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Character border", 2);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Plain "),
    Build.text("[bordered]", {
      border: { style: "single", size: 8, color: "C00000", space: 0 }
    }),
    Build.text(" plain.")
  ])
);

// ---------------------------------------------------------------------------
// Emphasis mark (East Asian typographic dots)
// ---------------------------------------------------------------------------
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("着重号:"),
    Build.text("重要内容", {
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
  Build.paragraph([
    Build.text("Symbols (Wingdings): "),
    Build.symbol("Wingdings", "F0FC"), // ✓
    Build.text("  "),
    Build.symbol("Wingdings", "F0FB"), // ✗
    Build.text("  "),
    Build.symbol("Wingdings", "F046") // ⌚
  ])
);

// ---------------------------------------------------------------------------
// Edge case 1: empty RunProperties object — should produce a clean run
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);
const emptyProps: RunProperties = {};
Document.addParagraphElement(
  doc,
  Build.paragraph([Build.text("Run with empty properties.", emptyProps)])
);

// Edge case 2: zero-length text run (writer should still emit valid XML)
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Before["),
    Build.text("", { bold: true }),
    Build.text("]After (zero-length bold run).")
  ])
);

// Edge case 3: many consecutive identical-formatted runs
Document.addParagraphElement(
  doc,
  Build.paragraph(
    Array.from({ length: 20 }, (_, i) =>
      Build.text(`r${i} `, { color: "0070C0", italic: i % 2 === 0 })
    ) as Run[]
  )
);

// Edge case 4: deeply mixed inline formatting
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Mixed: "),
    Build.bold("B"),
    Build.italic("I"),
    Build.underline("U"),
    Build.strikethrough("S"),
    Build.text("…and "),
    Build.text("RED-LARGE-BOLD-ITALIC-CAPS", {
      color: "C00000",
      size: Units.ptToHalfPoint(16),
      bold: true,
      italic: true,
      caps: true
    })
  ])
);

const buf = await Io.toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "02-formatting.docx"), buf);
console.log(`  → 02-formatting.docx (${buf.length} bytes)`);
