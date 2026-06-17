/**
 * Word Example 05 — Lists & numbering
 *
 * Covers:
 *   - Bullet list (built-in convenience API)
 *   - Numbered list (built-in convenience API)
 *   - Multi-level nested mix (level 0 / 1 / 2)
 *   - Custom abstract numbering (e.g. legal "1.1.1", upperRoman, lettered)
 *   - Restart numbering after a heading
 *   - Continuous numbering across a heading (using the same numId twice)
 *   - Task list (rendered with check-box symbols since real "task lists" are
 *     a docx convention, not a Word primitive)
 *   - Rich-formatted list items (mixed runs in one list item)
 *   - Edge case: deeply nested levels (0..8), empty list, single-item list
 *
 * Output: tmp/word-examples/05-lists.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io } from "../index";
import type {
  AbstractNumbering,
  NumberingInstance,
  NumberingLevel,
  ParagraphChild
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word Module — Lists & Numbering", 1);

// ---------------------------------------------------------------------------
// 1. Built-in bullet & numbered lists (uses internal abstract num definition)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Bullet & numbered lists", 2);

Document.addBulletList(doc, ["First bullet", "Second bullet", "Third bullet"]);
Document.addNumberedList(doc, ["Step one", "Step two", "Step three"]);

// Rich-formatted items (mix of runs)
Document.addBulletList(doc, [
  [Build.text("Plain bullet item")],
  [Build.bold("Bold-prefix"), Build.text(" with normal continuation")],
  [Build.italic("Italic"), Build.text(" + "), Build.text("colored", { color: "C00000" })]
] as (string | ParagraphChild[])[]);

// ---------------------------------------------------------------------------
// 2. Mixed multi-level (built-in API supports level 0..2)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Multi-level mix", 2);
Document.addBulletList(doc, ["Top-level A"], 0);
Document.addBulletList(doc, ["Indented A.1", "Indented A.2"], 1);
Document.addBulletList(doc, ["Deepest A.1.a"], 2);
Document.addBulletList(doc, ["Top-level B"], 0);
Document.addNumberedList(doc, ["Step 1"], 0);
Document.addNumberedList(doc, ["sub a", "sub b"], 1);
Document.addNumberedList(doc, ["sub-sub i"], 2);

// ---------------------------------------------------------------------------
// 3. Custom abstract numbering — legal style (1, 1.1, 1.1.1)
//    The bullet/numbered helpers only seed level 0..2 templates, so for
//    deeper or custom formats we attach AbstractNumbering + NumberingInstance
//    onto the built model directly and reference them by numId from each
//    paragraph's numbering property.
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Legal numbering (1.1.1)", 2);

const legalLevels: NumberingLevel[] = [];
for (let lvl = 0; lvl < 4; lvl++) {
  let txt = "";
  for (let i = 1; i <= lvl + 1; i++) {
    txt += `%${i}.`;
  }
  legalLevels.push({
    level: lvl,
    start: 1,
    format: "decimal",
    text: txt,
    justification: "left",
    paragraphProperties: { indent: { left: 720 + lvl * 360, hanging: 360 } }
  });
}
const legalAbstract: AbstractNumbering = {
  abstractNumId: 100,
  multiLevelType: "multilevel",
  levels: legalLevels
};
const legalNumId = 100;
const legalInstance: NumberingInstance = { numId: legalNumId, abstractNumId: 100 };

const docModel = Document.build(doc);
const merged = {
  ...docModel,
  abstractNumberings: [...(docModel.abstractNumberings ?? []), legalAbstract],
  numberingInstances: [...(docModel.numberingInstances ?? []), legalInstance],
  body: [
    ...docModel.body,
    Build.textParagraph("Article one — overview.", {
      numbering: { numId: legalNumId, level: 0 }
    }),
    Build.textParagraph("First definition.", { numbering: { numId: legalNumId, level: 1 } }),
    Build.textParagraph("Second definition.", { numbering: { numId: legalNumId, level: 1 } }),
    Build.textParagraph("A footnote.", { numbering: { numId: legalNumId, level: 2 } }),
    Build.textParagraph("Article two — scope.", { numbering: { numId: legalNumId, level: 0 } }),
    Build.textParagraph("Sub-clause.", { numbering: { numId: legalNumId, level: 1 } })
  ]
};

// ---------------------------------------------------------------------------
// 4. Custom abstract numbering — upperRoman / upperLetter mix
// ---------------------------------------------------------------------------
const romanAbstract: AbstractNumbering = {
  abstractNumId: 101,
  multiLevelType: "multilevel",
  levels: [
    {
      level: 0,
      start: 1,
      format: "upperRoman",
      text: "%1.",
      justification: "right",
      paragraphProperties: { indent: { left: 720, hanging: 360 } }
    },
    {
      level: 1,
      start: 1,
      format: "upperLetter",
      text: "%2.",
      justification: "left",
      paragraphProperties: { indent: { left: 1440, hanging: 360 } }
    },
    {
      level: 2,
      start: 1,
      format: "decimal",
      text: "%3)",
      justification: "left",
      paragraphProperties: { indent: { left: 2160, hanging: 360 } }
    }
  ]
};
const romanNumId = 101;
const merged2 = {
  ...merged,
  abstractNumberings: [...(merged.abstractNumberings ?? []), romanAbstract],
  numberingInstances: [
    ...(merged.numberingInstances ?? []),
    { numId: romanNumId, abstractNumId: 101 }
  ],
  body: [
    ...merged.body,
    Build.textParagraph("Roman / Letter / Decimal:", { style: "Heading2" }),
    Build.textParagraph("Introduction", { numbering: { numId: romanNumId, level: 0 } }),
    Build.textParagraph("First sub-section", { numbering: { numId: romanNumId, level: 1 } }),
    Build.textParagraph("First note", { numbering: { numId: romanNumId, level: 2 } }),
    Build.textParagraph("Second note", { numbering: { numId: romanNumId, level: 2 } }),
    Build.textParagraph("Second sub-section", { numbering: { numId: romanNumId, level: 1 } }),
    Build.textParagraph("Methodology", { numbering: { numId: romanNumId, level: 0 } })
  ]
};

// ---------------------------------------------------------------------------
// 5. Restart vs continue across a heading
//    Two separate numbering instances → restart;
//    Same instance reused → continues.
// ---------------------------------------------------------------------------
const restartInstance = { numId: 200, abstractNumId: 0 }; // reuse default numbered abstract (0)
const merged3 = {
  ...merged2,
  // duplicate abstract not needed — reuse abstractNumId=0 created by the
  // built-in addNumberedList helper above
  numberingInstances: [...(merged2.numberingInstances ?? []), restartInstance],
  body: [
    ...merged2.body,
    Build.textParagraph("4. Restart vs continue", { style: "Heading2" }),
    Build.textParagraph("First (continuous): item 1", { numbering: { numId: 1, level: 0 } }),
    Build.textParagraph("First (continuous): item 2", { numbering: { numId: 1, level: 0 } }),
    Build.textParagraph("— heading interrupts —", { style: "Heading3" }),
    Build.textParagraph("First (continuous): item 3", { numbering: { numId: 1, level: 0 } }),
    Build.textParagraph("Second (restart): item 1", {
      numbering: { numId: restartInstance.numId, level: 0 }
    }),
    Build.textParagraph("Second (restart): item 2", {
      numbering: { numId: restartInstance.numId, level: 0 }
    })
  ]
};

// ---------------------------------------------------------------------------
// 6. Task list (visual: ☐ / ☑ Wingdings glyph + plain text)
// ---------------------------------------------------------------------------
const merged4 = {
  ...merged3,
  body: [
    ...merged3.body,
    Build.textParagraph("5. Task list (visual)", { style: "Heading2" }),
    Build.paragraph([Build.symbol("Wingdings", "F0FE"), Build.text("  Buy groceries")]),
    Build.paragraph([Build.symbol("Wingdings", "F0FE"), Build.text("  Reply to e-mails")]),
    Build.paragraph([Build.symbol("Wingdings", "F0FC"), Build.text("  Take out trash (done)")]),
    Build.paragraph([Build.symbol("Wingdings", "F0FE"), Build.text("  Walk the dog")])
  ]
};

// ---------------------------------------------------------------------------
// 7. Edge: very deep nesting using built-in API (levels 0-2 supported by
// helper; for deeper levels we build our own abstract num)
// ---------------------------------------------------------------------------
const deepLevels: NumberingLevel[] = [];
for (let lvl = 0; lvl < 9; lvl++) {
  deepLevels.push({
    level: lvl,
    start: 1,
    format: lvl % 2 === 0 ? "decimal" : "lowerLetter",
    text: lvl % 2 === 0 ? `%${lvl + 1}.` : `%${lvl + 1})`,
    justification: "left",
    paragraphProperties: { indent: { left: 360 * (lvl + 2), hanging: 360 } }
  });
}
const deepAbstract: AbstractNumbering = {
  abstractNumId: 300,
  multiLevelType: "multilevel",
  levels: deepLevels
};
const deepNumId = 300;
const final = {
  ...merged4,
  abstractNumberings: [...(merged4.abstractNumberings ?? []), deepAbstract],
  numberingInstances: [
    ...(merged4.numberingInstances ?? []),
    { numId: deepNumId, abstractNumId: 300 }
  ],
  body: [
    ...merged4.body,
    Build.textParagraph("6. Deep nesting (level 0..8)", { style: "Heading2" }),
    ...Array.from({ length: 9 }, (_, lvl) =>
      Build.textParagraph(`Depth ${lvl}`, { numbering: { numId: deepNumId, level: lvl } })
    )
  ]
};

const buf = await Io.toBuffer(final);
fs.writeFileSync(path.join(outDir, "05-lists.docx"), buf);
console.log(`  → 05-lists.docx (${buf.length} bytes)`);
