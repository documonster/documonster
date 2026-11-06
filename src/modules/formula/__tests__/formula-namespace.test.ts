import { Formula } from "@formula/index";
import { parse as parseTokens } from "@formula/syntax/parser";
import { printAst } from "@formula/syntax/print";
import { describe, expect, it } from "vitest";

describe("Formula namespace", () => {
  it("parse accepts source text and tokens alike", () => {
    const fromText = Formula.parse("=SUM(A1:B2)+1");
    expect(fromText).toEqual(Formula.parse(Formula.tokenize("=SUM(A1:B2)+1")));
  });

  it("format is the inverse of parse", () => {
    expect(Formula.format(Formula.parse("=(1+2)*A1"))).toBe("(1+2)*A1");
  });

  it("keeps print as a deprecated forward to format", () => {
    expect(Formula.format).toBe(printAst);
    expect(Formula.print).toBe(printAst);
  });

  it("parse(tokens) is exactly the old token parser", () => {
    for (const src of ["=SUM(A1:B2)+1", 'IF(A1>0,"y","n")', "=-A1^2%"]) {
      const tokens = Formula.tokenize(src);
      expect(Formula.parse(tokens)).toEqual(parseTokens(tokens));
    }
  });
});
