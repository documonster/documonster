/**
 * Word Example 35 — MathML conversion + ruby + mathPhantom + mathBlock builder
 *
 * Picks up math-related APIs not exercised by 12-math:
 *   - mathBlock builder (used directly to attach a math block)
 *   - ommlToMathML — turn an internal OMML model into a MathML string
 *   - mathMLToOmml — round-trip a MathML string back to OMML
 *   - mathPhantom — invisible expression that takes up space
 *   - ruby (Japanese furigana / Chinese pinyin)
 *
 * Output: tmp/word-examples/35-math-extras.docx + .mathml dumps
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Convert, Io } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/35-math-extras"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

// ---------------------------------------------------------------------------
// 1. Direct mathBlock builder (instead of Document.addMath helper)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Math extras", 1);
Document.addContent(
  doc,
  Build.mathBlock([
    Build.mathRun("a + b = "),
    Build.mathFraction([Build.mathRun("c")], [Build.mathRun("d")])
  ])
);

// ---------------------------------------------------------------------------
// 2. mathPhantom — reserves space without showing the expression. Useful for
// alignment in equation arrays. `show: false` (serialized as <m:show
// m:val="0"/>) is what actually hides the base — in OOXML `show` defaults to
// ON, so omitting it would leave the placeholder text visible. `transparent`
// additionally marks it transparent for spacing.
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Phantom (invisible spacing)", 2);
Document.addContent(
  doc,
  Build.mathBlock([
    Build.mathRun("|x| = "),
    Build.mathPhantom([Build.mathRun("placeholder same width")], {
      show: false,
      transparent: true
    }),
    Build.mathRun(" (visible result)")
  ])
);

// ---------------------------------------------------------------------------
// 3. Ruby — phonetic guide (Japanese furigana / Chinese pinyin)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Ruby (phonetic guide)", 2);
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Japanese: "),
    Build.ruby("漢字", "かんじ", { align: "center", language: "ja-JP" }),
    Build.text("、"),
    Build.ruby("日本語", "にほんご", { align: "center", language: "ja-JP" }),
    Build.text(" — Chinese pinyin: "),
    Build.ruby("中文", "zhōngwén", { align: "center", language: "zh-CN" })
  ])
);

// ---------------------------------------------------------------------------
// 4. Save the doc & extract its math blocks for MathML conversion
// ---------------------------------------------------------------------------
const built = Document.build(doc);
const buf = await Io.toBuffer(built);
fs.writeFileSync(path.join(outDir, "35-math-extras.docx"), buf);
console.log(`  → 35-math-extras.docx (${buf.length} bytes)`);

// ---------------------------------------------------------------------------
// 5. ommlToMathML — convert each math block to MathML
// ---------------------------------------------------------------------------
const mathBlocks = built.body.filter(b => "type" in b && b.type === "math");
console.log(`  ${mathBlocks.length} math blocks found`);
mathBlocks.forEach((block, i) => {
  if ("content" in block) {
    const mathml = Convert.ommlToMathML(block.content);
    fs.writeFileSync(path.join(outDir, `block-${i + 1}.mathml`), mathml);
    console.log(`    block ${i + 1}: ${mathml.length} chars (block-${i + 1}.mathml)`);
  }
});

// ---------------------------------------------------------------------------
// 6. mathMLToOmml — round-trip a hand-written MathML string into OMML model
// ---------------------------------------------------------------------------
const mathmlInput = `<math xmlns="http://www.w3.org/1998/Math/MathML">
  <mrow>
    <mfrac><mn>1</mn><mn>2</mn></mfrac>
    <mo>+</mo>
    <msqrt><mn>2</mn></msqrt>
    <mo>=</mo>
    <msup><mi>x</mi><mn>2</mn></msup>
  </mrow>
</math>`;
const omml = Convert.mathMLToOmml(mathmlInput);
console.log(`  mathMLToOmml: ${omml.length} OMML node(s)`);
const docFromMathml = Document.create();
Document.useDefaultStyles(docFromMathml);
Document.addHeading(docFromMathml, "Imported from MathML", 1);
Document.addMath(docFromMathml, omml);
const buf2 = await Io.toBuffer(Document.build(docFromMathml));
fs.writeFileSync(path.join(outDir, "imported-from-mathml.docx"), buf2);
console.log(`  → imported-from-mathml.docx (${buf2.length} bytes)`);

// ---------------------------------------------------------------------------
// 7. Build a more complex expression to see it survive both directions
// ---------------------------------------------------------------------------
const complex = [
  Build.mathSum(
    [
      Build.mathFraction(
        [Build.mathSuperScript([Build.mathRun("k")], [Build.mathRun("2")])],
        [Build.mathRun("k!")]
      )
    ],
    [Build.mathRun("k=0")],
    [Build.mathRun("∞")]
  ),
  Build.mathRun(" = "),
  Build.mathDelimiter(
    [
      [Build.mathSqrt([Build.mathRun("1 + π")])],
      [
        Build.mathMatrix([
          [[Build.mathRun("1")], [Build.mathRun("0")]],
          [[Build.mathRun("0")], [Build.mathRun("1")]]
        ])
      ]
    ],
    { beginChar: "(", endChar: ")", separatorChar: "|" }
  )
];
const complexMl = Convert.ommlToMathML(complex);
const roundTrip = Convert.mathMLToOmml(complexMl);
fs.writeFileSync(path.join(outDir, "complex.mathml"), complexMl);
console.log(
  `  complex round-trip: original=${complex.length} nodes → MathML=${complexMl.length} chars → OMML=${roundTrip.length} nodes`
);
