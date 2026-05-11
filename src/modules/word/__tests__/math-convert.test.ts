/**
 * DOCX Module - MathML ↔ OMML Conversion Tests
 */

import { describe, it, expect } from "vitest";

import { ommlToMathML, mathMLToOmml } from "../advanced/math-convert";
import type { MathContent } from "../types";

// =============================================================================
// OMML → MathML
// =============================================================================

describe("ommlToMathML", () => {
  it("should convert a simple math run", () => {
    const content: MathContent[] = [{ type: "mathRun", text: "x" }];
    const result = ommlToMathML(content);
    expect(result).toContain("<mi>x</mi>");
    expect(result).toMatch(/^<math xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML">/);
    expect(result).toMatch(/<\/math>$/);
  });

  it("should convert a number", () => {
    const content: MathContent[] = [{ type: "mathRun", text: "42" }];
    const result = ommlToMathML(content);
    expect(result).toContain("<mn>42</mn>");
  });

  it("should convert an operator", () => {
    const content: MathContent[] = [{ type: "mathRun", text: "+" }];
    const result = ommlToMathML(content);
    expect(result).toContain("<mo>+</mo>");
  });

  it("should convert a fraction", () => {
    const content: MathContent[] = [
      {
        type: "mathFraction",
        numerator: [{ type: "mathRun", text: "a" }],
        denominator: [{ type: "mathRun", text: "b" }]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<mfrac>");
    expect(result).toContain("<mi>a</mi>");
    expect(result).toContain("<mi>b</mi>");
    expect(result).toContain("</mfrac>");
  });

  it("should convert superscript", () => {
    const content: MathContent[] = [
      {
        type: "mathSuperScript",
        base: [{ type: "mathRun", text: "x" }],
        superScript: [{ type: "mathRun", text: "2" }]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<msup>");
    expect(result).toContain("<mi>x</mi>");
    expect(result).toContain("<mn>2</mn>");
  });

  it("should convert subscript", () => {
    const content: MathContent[] = [
      {
        type: "mathSubScript",
        base: [{ type: "mathRun", text: "a" }],
        subScript: [{ type: "mathRun", text: "i" }]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<msub>");
    expect(result).toContain("<mi>a</mi>");
    expect(result).toContain("<mi>i</mi>");
  });

  it("should convert square root", () => {
    const content: MathContent[] = [
      {
        type: "mathRadical",
        content: [{ type: "mathRun", text: "x" }]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<msqrt>");
    expect(result).toContain("<mi>x</mi>");
  });

  it("should convert nth root", () => {
    const content: MathContent[] = [
      {
        type: "mathRadical",
        content: [{ type: "mathRun", text: "x" }],
        degree: [{ type: "mathRun", text: "3" }]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<mroot>");
    expect(result).toContain("<mi>x</mi>");
    expect(result).toContain("<mn>3</mn>");
  });

  it("should convert summation (nary)", () => {
    const content: MathContent[] = [
      {
        type: "mathNary",
        char: "\u2211",
        sub: [
          { type: "mathRun", text: "i" },
          { type: "mathRun", text: "=" },
          { type: "mathRun", text: "0" }
        ],
        sup: [{ type: "mathRun", text: "n" }],
        content: [{ type: "mathRun", text: "i" }]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<munderover>");
    expect(result).toContain("\u2211");
  });

  it("should convert integral", () => {
    const content: MathContent[] = [
      {
        type: "mathNary",
        char: "\u222B",
        sub: [{ type: "mathRun", text: "0" }],
        sup: [{ type: "mathRun", text: "1" }],
        content: [
          { type: "mathRun", text: "f" },
          { type: "mathRun", text: "(" },
          { type: "mathRun", text: "x" },
          { type: "mathRun", text: ")" }
        ]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("\u222B");
    expect(result).toContain("<munderover>");
  });

  it("should convert delimiter (parentheses)", () => {
    const content: MathContent[] = [
      {
        type: "mathDelimiter",
        beginChar: "(",
        endChar: ")",
        content: [
          [
            { type: "mathRun", text: "x" },
            { type: "mathRun", text: "+" },
            { type: "mathRun", text: "y" }
          ]
        ]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<mo>(</mo>");
    expect(result).toContain("<mo>)</mo>");
    expect(result).toContain("<mi>x</mi>");
  });

  it("should convert matrix", () => {
    const content: MathContent[] = [
      {
        type: "mathMatrix",
        rows: [
          [[{ type: "mathRun", text: "a" }], [{ type: "mathRun", text: "b" }]],
          [[{ type: "mathRun", text: "c" }], [{ type: "mathRun", text: "d" }]]
        ]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<mtable>");
    expect(result).toContain("<mtr>");
    expect(result).toContain("<mtd>");
    expect(result).toContain("<mi>a</mi>");
    expect(result).toContain("<mi>d</mi>");
  });

  it("should convert sub-superscript", () => {
    const content: MathContent[] = [
      {
        type: "mathSubSuperScript",
        base: [{ type: "mathRun", text: "x" }],
        subScript: [{ type: "mathRun", text: "i" }],
        superScript: [{ type: "mathRun", text: "2" }]
      }
    ];
    const result = ommlToMathML(content);
    expect(result).toContain("<msubsup>");
    expect(result).toContain("<mi>x</mi>");
    expect(result).toContain("<mi>i</mi>");
    expect(result).toContain("<mn>2</mn>");
  });
});

// =============================================================================
// MathML → OMML
// =============================================================================

describe("mathMLToOmml", () => {
  it("should parse a simple identifier", () => {
    const result = mathMLToOmml("<math><mi>x</mi></math>");
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathRun");
    expect((result[0] as any).text).toBe("x");
  });

  it("should parse a number", () => {
    const result = mathMLToOmml("<math><mn>42</mn></math>");
    expect(result.length).toBe(1);
    expect((result[0] as any).text).toBe("42");
  });

  it("should parse a fraction", () => {
    const result = mathMLToOmml(
      "<math><mfrac><mrow><mi>a</mi></mrow><mrow><mi>b</mi></mrow></mfrac></math>"
    );
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathFraction");
    const frac = result[0] as any;
    expect(frac.numerator.length).toBe(1);
    expect(frac.numerator[0].text).toBe("a");
    expect(frac.denominator.length).toBe(1);
    expect(frac.denominator[0].text).toBe("b");
  });

  it("should parse superscript", () => {
    const result = mathMLToOmml(
      "<math><msup><mrow><mi>x</mi></mrow><mrow><mn>2</mn></mrow></msup></math>"
    );
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathSuperScript");
    const sup = result[0] as any;
    expect(sup.base[0].text).toBe("x");
    expect(sup.superScript[0].text).toBe("2");
  });

  it("should parse subscript", () => {
    const result = mathMLToOmml(
      "<math><msub><mrow><mi>a</mi></mrow><mrow><mi>i</mi></mrow></msub></math>"
    );
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathSubScript");
    const sub = result[0] as any;
    expect(sub.base[0].text).toBe("a");
    expect(sub.subScript[0].text).toBe("i");
  });

  it("should parse square root", () => {
    const result = mathMLToOmml("<math><msqrt><mi>x</mi></msqrt></math>");
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathRadical");
    const rad = result[0] as any;
    expect(rad.content[0].text).toBe("x");
    expect(rad.degree).toBeUndefined();
  });

  it("should parse nth root", () => {
    const result = mathMLToOmml(
      "<math><mroot><mrow><mi>x</mi></mrow><mrow><mn>3</mn></mrow></mroot></math>"
    );
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathRadical");
    const rad = result[0] as any;
    expect(rad.content[0].text).toBe("x");
    expect(rad.degree[0].text).toBe("3");
  });

  it("should parse summation (munderover)", () => {
    const mathml = `<math><munderover><mo>\u2211</mo><mrow><mi>i</mi></mrow><mrow><mi>n</mi></mrow></munderover></math>`;
    const result = mathMLToOmml(mathml);
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathNary");
    const nary = result[0] as any;
    expect(nary.char).toBe("\u2211");
  });

  it("should parse matrix", () => {
    const mathml = `<math><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable></math>`;
    const result = mathMLToOmml(mathml);
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathMatrix");
    const matrix = result[0] as any;
    expect(matrix.rows.length).toBe(2);
    expect(matrix.rows[0].length).toBe(2);
  });

  it("should parse msubsup", () => {
    const mathml = `<math><msubsup><mrow><mi>x</mi></mrow><mrow><mi>i</mi></mrow><mrow><mn>2</mn></mrow></msubsup></math>`;
    const result = mathMLToOmml(mathml);
    expect(result.length).toBe(1);
    expect(result[0]!.type).toBe("mathSubSuperScript");
  });

  it("should handle complex expression round-trip", () => {
    // Create OMML: x^2 + y^2
    const omml: MathContent[] = [
      {
        type: "mathSuperScript",
        base: [{ type: "mathRun", text: "x" }],
        superScript: [{ type: "mathRun", text: "2" }]
      },
      { type: "mathRun", text: "+" },
      {
        type: "mathSuperScript",
        base: [{ type: "mathRun", text: "y" }],
        superScript: [{ type: "mathRun", text: "2" }]
      }
    ];

    // Convert to MathML and back
    const mathml = ommlToMathML(omml);
    expect(mathml).toContain("<msup>");
    expect(mathml).toContain("<mi>x</mi>");
    expect(mathml).toContain("<mi>y</mi>");

    // Parse back
    const parsed = mathMLToOmml(mathml);
    expect(parsed.length).toBe(3);
    expect(parsed[0]!.type).toBe("mathSuperScript");
    expect(parsed[1]!.type).toBe("mathRun");
    expect(parsed[2]!.type).toBe("mathSuperScript");
  });
});
