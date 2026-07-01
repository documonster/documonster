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

import { Document, Build, Io } from "../index";

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
Document.addMath(doc, [Build.mathSuperScript([Build.mathRun("E = mc")], [Build.mathRun("2")])]);

// ---------------------------------------------------------------------------
// 2. Fractions
// ---------------------------------------------------------------------------
Document.addHeading(doc, "2. Fractions", 2);
Document.addMath(doc, [
  Build.mathRun("y = "),
  Build.mathFraction(
    [
      Build.mathRun("a"),
      Build.mathRun(" + "),
      Build.mathSuperScript([Build.mathRun("x")], [Build.mathRun("2")])
    ],
    [Build.mathRun("b"), Build.mathRun(" - "), Build.mathRun("c")]
  )
]);
// Linear fraction
Document.addMath(doc, [
  Build.mathRun("linear: "),
  Build.mathFraction([Build.mathRun("a")], [Build.mathRun("b")], "lin")
]);
// No-bar (stack)
Document.addMath(doc, [
  Build.mathRun("stack: "),
  Build.mathFraction([Build.mathRun("a")], [Build.mathRun("b")], "noBar")
]);

// ---------------------------------------------------------------------------
// 3. Roots
// ---------------------------------------------------------------------------
Document.addHeading(doc, "3. Roots", 2);
Document.addMath(doc, [
  Build.mathRun("x = "),
  Build.mathFraction(
    [
      Build.mathRun("-b ± "),
      Build.mathSqrt([
        Build.mathSuperScript([Build.mathRun("b")], [Build.mathRun("2")]),
        Build.mathRun(" - 4ac")
      ])
    ],
    [Build.mathRun("2a")]
  )
]);
Document.addMath(doc, [
  Build.mathRun("∛8 = "),
  Build.mathRoot([Build.mathRun("3")], [Build.mathRun("8")]),
  Build.mathRun(" = 2")
]);

// ---------------------------------------------------------------------------
// 4. Sub/super scripts (and pre-sub-super)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "4. Subscript / Superscript", 2);
Document.addMath(doc, [
  Build.mathSubScript([Build.mathRun("a")], [Build.mathRun("ij")]),
  Build.mathRun(" · "),
  Build.mathSuperScript([Build.mathRun("b")], [Build.mathRun("k")]),
  Build.mathRun(" · "),
  Build.mathSubSuperScript([Build.mathRun("c")], [Build.mathRun("p")], [Build.mathRun("q")]),
  Build.mathRun(" · "),
  Build.mathPreSubSuperScript([Build.mathRun("X")], [Build.mathRun("a")], [Build.mathRun("b")])
]);

// ---------------------------------------------------------------------------
// 5. N-ary operators
// ---------------------------------------------------------------------------
Document.addHeading(doc, "5. Sums, integrals, products", 2);
Document.addMath(doc, [
  Build.mathSum(
    [Build.mathSuperScript([Build.mathRun("k")], [Build.mathRun("2")])],
    [Build.mathRun("k=1")],
    [Build.mathRun("n")]
  ),
  Build.mathRun(" = "),
  Build.mathFraction([Build.mathRun("n(n+1)(2n+1)")], [Build.mathRun("6")])
]);
Document.addMath(doc, [
  Build.mathIntegral(
    [Build.mathSuperScript([Build.mathRun("e")], [Build.mathRun("-x")])],
    [Build.mathRun("0")],
    [Build.mathRun("∞")]
  ),
  Build.mathRun(" dx = 1")
]);
Document.addMath(doc, [
  Build.mathProduct([Build.mathRun("k")], [Build.mathRun("k=1")], [Build.mathRun("n")]),
  Build.mathRun(" = n!")
]);
// Custom nary — contour integral ∮
Document.addMath(doc, [Build.mathNary("\u222E", [Build.mathRun("F · dr")])]);

// ---------------------------------------------------------------------------
// 6. Functions, limits
// ---------------------------------------------------------------------------
Document.addHeading(doc, "6. Functions & limits", 2);
Document.addMath(doc, [Build.mathFunction([Build.mathRun("sin")], [Build.mathRun("(2x + π)")])]);
Document.addMath(doc, [
  Build.mathLimit([Build.mathRun("lim")], [Build.mathRun("x → 0")]),
  Build.mathRun(" "),
  Build.mathFraction([Build.mathRun("sin x")], [Build.mathRun("x")]),
  Build.mathRun(" = 1")
]);

// ---------------------------------------------------------------------------
// 7. Delimiters
// ---------------------------------------------------------------------------
Document.addHeading(doc, "7. Delimiters", 2);
Document.addMath(doc, [
  Build.mathDelimiter([[Build.mathRun("a"), Build.mathRun(", "), Build.mathRun("b")]], {
    beginChar: "(",
    endChar: ")"
  })
]);
Document.addMath(doc, [
  Build.mathDelimiter([[Build.mathRun("a")], [Build.mathRun("b")], [Build.mathRun("c")]], {
    beginChar: "{",
    endChar: "}",
    separatorChar: "|"
  })
]);
Document.addMath(doc, [
  Build.mathDelimiter(
    [[Build.mathRun("v"), Build.mathSubScript([Build.mathRun("")], [Build.mathRun("max")])]],
    {
      beginChar: "[",
      endChar: "]"
    }
  )
]);

// ---------------------------------------------------------------------------
// 8. Matrices
// ---------------------------------------------------------------------------
Document.addHeading(doc, "8. Matrices", 2);
Document.addMath(doc, [
  Build.mathDelimiter(
    [
      [
        Build.mathMatrix([
          [[Build.mathRun("a")], [Build.mathRun("b")]],
          [[Build.mathRun("c")], [Build.mathRun("d")]]
        ])
      ]
    ],
    { beginChar: "[", endChar: "]" }
  )
]);
Document.addMath(doc, [
  Build.mathDelimiter(
    [
      [
        Build.mathMatrix([
          [[Build.mathRun("1")], [Build.mathRun("2")], [Build.mathRun("3")]],
          [[Build.mathRun("4")], [Build.mathRun("5")], [Build.mathRun("6")]],
          [[Build.mathRun("7")], [Build.mathRun("8")], [Build.mathRun("9")]]
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
  Build.mathAccent([Build.mathRun("v")], "→"),
  Build.mathRun(" + "),
  Build.mathAccent([Build.mathRun("u")], "^"),
  Build.mathRun(" + "),
  Build.mathBar([Build.mathRun("z")], "top"),
  Build.mathRun(" + "),
  Build.mathBar([Build.mathRun("w")], "bottom")
]);
Document.addMath(doc, [
  Build.mathGroupChar([Build.mathRun("x + y + z")], { char: "\u23DE", position: "top" }),
  Build.mathRun(" — "),
  Build.mathGroupChar([Build.mathRun("a · b")], { char: "\u23DF", position: "bottom" })
]);
Document.addMath(doc, [
  Build.mathBorderBox([Build.mathRun("answer = 42")], { strikeTlBr: false }),
  Build.mathRun(" "),
  Build.mathBox([Build.mathRun("just a box")])
]);

// ---------------------------------------------------------------------------
// 10. Equation array (aligned multi-line equations)
// ---------------------------------------------------------------------------
Document.addHeading(doc, "10. Equation array", 2);
Document.addMath(doc, [
  Build.mathEquationArray([
    [Build.mathRun("x = a + b")],
    [Build.mathRun("  = a + (c - d)")],
    [Build.mathRun("  = (a + c) - d")]
  ])
]);

// ---------------------------------------------------------------------------
// Edge cases
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Edge cases", 2);
Document.addMath(doc, []); // empty math block
Document.addMath(doc, [
  // very deep nesting: a fraction inside a sum inside an integral
  Build.mathIntegral(
    [
      Build.mathSum(
        [
          Build.mathFraction(
            [Build.mathSuperScript([Build.mathRun("k")], [Build.mathRun("2")])],
            [Build.mathRun("k!")]
          )
        ],
        [Build.mathRun("k=0")],
        [Build.mathRun("∞")]
      )
    ],
    [Build.mathRun("0")],
    [Build.mathRun("1")]
  )
]);

// Math mixed with normal prose in same paragraph using a math block placed
// adjacent to text paragraphs (Word treats mathBlock as its own paragraph).
Document.addParagraph(doc, "End of math examples.");

const buf = await Io.toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "12-math.docx"), buf);
console.log(`  → 12-math.docx (${buf.length} bytes)`);
