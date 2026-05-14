/**
 * Word Example 12 — Math equations (OOML)
 *
 * Covers every math primitive supplied by the builder helpers:
 *   - Inline math (a math run inside a paragraph)
 *   - Math block (display equation as its own paragraph)
 *   - Fraction (regular & "linear" %1/%2 / "no bar" stacking)
 *   - Square root and nth root
 *   - Subscript / superscript / sub-superscript / pre-sub-superscript
 *   - n-ary operators: sum, integral, product, custom (∮)
 *   - Functions (sin, log)
 *   - Limit
 *   - Delimiters (parens, square brackets, curly with separator)
 *   - Matrix (2x2, 3x3)
 *   - Accent / bar / group character (over-brace) / border box
 *   - Equation array (alignment of multiple equations)
 *   - Edge cases: empty math block, deeply nested
 *
 * Output: tmp/word-examples/12-math.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  mathRun,
  mathFraction,
  mathSqrt,
  mathRoot,
  mathSum,
  mathIntegral,
  mathProduct,
  mathSuperScript,
  mathSubScript,
  mathSubSuperScript,
  mathPreSubSuperScript,
  mathNary,
  mathFunction,
  mathLimit,
  mathDelimiter,
  mathMatrix,
  mathAccent,
  mathBar,
  mathBorderBox,
  mathBox,
  mathGroupChar,
  mathEquationArray,
  toBuffer
} from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

Document.addHeading(doc, "Word — Math equations (OMML)", 1);

// ---------------------------------------------------------------------------
// Inline / display math: paragraphs that mix prose and a math block
// ---------------------------------------------------------------------------
Document.addHeading(doc, "1. Inline & display equations", 2);
Document.addParagraph(doc, "Mass–energy equivalence:");
Document.addMath(doc, [mathSuperScript([mathRun("E = mc")], [mathRun("2")])]);

// ---------------------------------------------------------------------------
// 2. Fractions
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Fractions", 2);
Document.addMath(doc, [
  mathRun("y = "),
  mathFraction(
    [mathRun("a"), mathRun(" + "), mathSuperScript([mathRun("x")], [mathRun("2")])],
    [mathRun("b"), mathRun(" - "), mathRun("c")]
  )
]);
// Linear fraction
Document.addMath(doc, [mathRun("linear: "), mathFraction([mathRun("a")], [mathRun("b")], "lin")]);
// No-bar (stack)
Document.addMath(doc, [mathRun("stack: "), mathFraction([mathRun("a")], [mathRun("b")], "noBar")]);

// ---------------------------------------------------------------------------
// 3. Roots
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Roots", 2);
Document.addMath(doc, [
  mathRun("x = "),
  mathFraction(
    [
      mathRun("-b ± "),
      mathSqrt([mathSuperScript([mathRun("b")], [mathRun("2")]), mathRun(" - 4ac")])
    ],
    [mathRun("2a")]
  )
]);
Document.addMath(doc, [
  mathRun("∛8 = "),
  mathRoot([mathRun("3")], [mathRun("8")]),
  mathRun(" = 2")
]);

// ---------------------------------------------------------------------------
// 4. Sub/super scripts (and pre-sub-super)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "4. Subscript / Superscript", 2);
Document.addMath(doc, [
  mathSubScript([mathRun("a")], [mathRun("ij")]),
  mathRun(" · "),
  mathSuperScript([mathRun("b")], [mathRun("k")]),
  mathRun(" · "),
  mathSubSuperScript([mathRun("c")], [mathRun("p")], [mathRun("q")]),
  mathRun(" · "),
  mathPreSubSuperScript([mathRun("X")], [mathRun("a")], [mathRun("b")])
]);

// ---------------------------------------------------------------------------
// 5. N-ary operators
// ---------------------------------------------------------------------------
Document.addHeading(doc, "5. Sums, integrals, products", 2);
Document.addMath(doc, [
  mathSum([mathSuperScript([mathRun("k")], [mathRun("2")])], [mathRun("k=1")], [mathRun("n")]),
  mathRun(" = "),
  mathFraction([mathRun("n(n+1)(2n+1)")], [mathRun("6")])
]);
Document.addMath(doc, [
  mathIntegral([mathSuperScript([mathRun("e")], [mathRun("-x")])], [mathRun("0")], [mathRun("∞")]),
  mathRun(" dx = 1")
]);
Document.addMath(doc, [
  mathProduct([mathRun("k")], [mathRun("k=1")], [mathRun("n")]),
  mathRun(" = n!")
]);
// Custom nary — contour integral ∮
Document.addMath(doc, [mathNary("\u222E", [mathRun("F · dr")])]);

// ---------------------------------------------------------------------------
// 6. Functions, limits
// ---------------------------------------------------------------------------
Document.addHeading(doc, "6. Functions & limits", 2);
Document.addMath(doc, [mathFunction([mathRun("sin")], [mathRun("(2x + π)")])]);
Document.addMath(doc, [
  mathLimit([mathRun("lim")], [mathRun("x → 0")]),
  mathRun(" "),
  mathFraction([mathRun("sin x")], [mathRun("x")]),
  mathRun(" = 1")
]);

// ---------------------------------------------------------------------------
// 7. Delimiters
// ---------------------------------------------------------------------------
Document.addHeading(doc, "7. Delimiters", 2);
Document.addMath(doc, [
  mathDelimiter([[mathRun("a"), mathRun(", "), mathRun("b")]], { beginChar: "(", endChar: ")" })
]);
Document.addMath(doc, [
  mathDelimiter([[mathRun("a")], [mathRun("b")], [mathRun("c")]], {
    beginChar: "{",
    endChar: "}",
    separatorChar: "|"
  })
]);
Document.addMath(doc, [
  mathDelimiter([[mathRun("v"), mathSubScript([mathRun("")], [mathRun("max")])]], {
    beginChar: "[",
    endChar: "]"
  })
]);

// ---------------------------------------------------------------------------
// 8. Matrices
// ---------------------------------------------------------------------------
Document.addHeading(doc, "8. Matrices", 2);
Document.addMath(doc, [
  mathDelimiter(
    [
      [
        mathMatrix([
          [[mathRun("a")], [mathRun("b")]],
          [[mathRun("c")], [mathRun("d")]]
        ])
      ]
    ],
    { beginChar: "[", endChar: "]" }
  )
]);
Document.addMath(doc, [
  mathDelimiter(
    [
      [
        mathMatrix([
          [[mathRun("1")], [mathRun("2")], [mathRun("3")]],
          [[mathRun("4")], [mathRun("5")], [mathRun("6")]],
          [[mathRun("7")], [mathRun("8")], [mathRun("9")]]
        ])
      ]
    ],
    { beginChar: "(", endChar: ")" }
  )
]);

// ---------------------------------------------------------------------------
// 9. Accents / bars / group char
// ---------------------------------------------------------------------------
Document.addHeading(doc, "9. Accents & decorations", 2);
Document.addMath(doc, [
  mathAccent([mathRun("v")], "→"),
  mathRun(" + "),
  mathAccent([mathRun("u")], "^"),
  mathRun(" + "),
  mathBar([mathRun("z")], "top"),
  mathRun(" + "),
  mathBar([mathRun("w")], "bottom")
]);
Document.addMath(doc, [
  mathGroupChar([mathRun("x + y + z")], { char: "\u23DE", position: "top" }),
  mathRun(" — "),
  mathGroupChar([mathRun("a · b")], { char: "\u23DF", position: "bottom" })
]);
Document.addMath(doc, [
  mathBorderBox([mathRun("answer = 42")], { strikeTlBr: false }),
  mathRun(" "),
  mathBox([mathRun("just a box")])
]);

// ---------------------------------------------------------------------------
// 10. Equation array (aligned multi-line equations)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "10. Equation array", 2);
Document.addMath(doc, [
  mathEquationArray([
    [mathRun("x = a + b")],
    [mathRun("  = a + (c - d)")],
    [mathRun("  = (a + c) - d")]
  ])
]);

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);
Document.addMath(doc, []); // empty math block
Document.addMath(doc, [
  // very deep nesting: a fraction inside a sum inside an integral
  mathIntegral(
    [
      mathSum(
        [mathFraction([mathSuperScript([mathRun("k")], [mathRun("2")])], [mathRun("k!")])],
        [mathRun("k=0")],
        [mathRun("∞")]
      )
    ],
    [mathRun("0")],
    [mathRun("1")]
  )
]);

// Math mixed with normal prose in same paragraph using a math block placed
// adjacent to text paragraphs (Word treats mathBlock as its own paragraph).
Document.addParagraph(doc, "End of math examples.");

const buf = await toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "12-math.docx"), buf);
console.log(`  → 12-math.docx (${buf.length} bytes)`);
