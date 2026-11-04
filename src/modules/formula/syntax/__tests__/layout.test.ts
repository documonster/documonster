/**
 * Formula layout — the spacing and display parentheses a formula's text carries beside its tree — and the parsing
 * details the XLSB codec needed fixed to reproduce Excel's text exactly.
 */
import { NodeType } from "@formula/syntax/ast";
import { parse, parseWithLayout } from "@formula/syntax/parser";
import { printAst, printAstWithLayout } from "@formula/syntax/print";
import { tokenize, tokenizeWithGaps } from "@formula/syntax/tokenizer";
import { describe, expect, it } from "vitest";

function reprint(text: string): string {
  const { tokens, gaps } = tokenizeWithGaps(text);
  const { ast, layout } = parseWithLayout(tokens, gaps);
  return printAstWithLayout(ast, layout);
}

describe("layout", () => {
  it.each([
    "A1 + B1",
    "( A1+B1 )*2",
    "(A1+B1) * 2",
    "SUM( A1, B1 )",
    'IFERROR((D5-C5)/ABS(C5), "")',
    "-  A1",
    "A1 %",
    "A1\n+B1",
    "IF(A1>0,\n  1,\n  2)",
    "A1:A3  B2:C2",
    "((A1))",
    "-(A1)",
    "1+ (A1)",
    "SUM( (A1,B1) )",
    "AND( A1>0, B1>0 )",
    "TODAY()-ROUNDDOWN(D1,0)>=(WEEKDAY(TODAY()))"
  ])("keeps the text of %j", text => {
    expect(reprint(text)).toBe(text);
  });

  it("drops only what a token stream cannot hold: space before a comma", () => {
    // Excel drops it too: its own XLSX of `SUM(A1 ,B1)` reads `SUM(A1,B1)`.
    expect(reprint("SUM(A1 ,B1)")).toBe("SUM(A1,B1)");
  });

  it("gives the same tree as parse", () => {
    const { tokens, gaps } = tokenizeWithGaps("( A1 + B1 ) * 2");
    expect(parseWithLayout(tokens, gaps).ast).toEqual(parse(tokenize("(A1+B1)*2")));
  });

  it("adds the parentheses precedence needs when a layout has none", () => {
    expect(printAstWithLayout(parse(tokenize("A1*(B1+C1)")), new Map())).toBe("A1*(B1+C1)");
  });
});

describe("Excel's precedence", () => {
  it("binds negation tighter than percent", () => {
    // Microsoft's table ranks negation first and percent second, and Excel's own tokens for `-A1%` are
    // `A1, PtgUminus, PtgPercent`. This parsed as `-(A1%)`, so the same tokens printed as `(-A1)%`.
    expect(parse(tokenize("-A1%"))).toEqual({
      type: NodeType.Percent,
      operand: { type: NodeType.UnaryOp, op: "-", operand: expect.objectContaining({ col: "A" }) }
    });
    expect(printAst(parse(tokenize("-A1%")))).toBe("-A1%");
    expect(printAst(parse(tokenize("-(A1%)")))).toBe("-(A1%)");
    // Still above `^`.
    expect(printAst(parse(tokenize("-2^2")))).toBe("-2^2");
    expect(parse(tokenize("-2^2"))).toMatchObject({ type: NodeType.BinaryOp, op: "^" });
  });
});

describe("sheet-qualified names and errors", () => {
  it("parses and prints Sheet2!Total", () => {
    const ast = parse(tokenize("Other!LocalO+1"));
    expect(ast).toMatchObject({ left: { type: NodeType.Name, name: "LocalO", sheet: "Other" } });
    expect(printAst(ast)).toBe("Other!LocalO+1");
    expect(printAst(parse(tokenize("'My Sheet'!Total")))).toBe("'My Sheet'!Total");
  });

  it("parses and prints Sheet2!#REF!, which a deleted reference becomes", () => {
    const ast = parse(tokenize("SUM(Other!#REF!)"));
    expect(ast).toMatchObject({ args: [{ type: NodeType.Error, value: "#REF!", sheet: "Other" }] });
    expect(printAst(ast)).toBe("SUM(Other!#REF!)");
  });
});

describe("structured references", () => {
  it("prints a lone item specifier without a second bracket, as Excel writes it", () => {
    expect(printAst(parse(tokenize("COUNTA(Tbl[#All])")))).toBe("COUNTA(Tbl[#All])");
    expect(printAst(parse(tokenize("Tbl[[#Totals],[Qty]]")))).toBe("Tbl[[#Totals],[Qty]]");
  });
});
