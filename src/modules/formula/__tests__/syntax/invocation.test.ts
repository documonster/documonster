/**
 * A call applied to the result of a call — `LAMBDA(n,n*2)(3)`.
 *
 * Excel accepts an argument list after any function call whose result is a lambda: an immediately invoked `LAMBDA`,
 * a curried one (`LAMBDA(x,LAMBDA(y,x+y))(1)(2)`), and a `LET`-bound lambda that returns a lambda (`f(1)(2)`). The
 * parser used to stop at the first `)` and report the second `(` as a trailing token.
 *
 * The public `AstNode` union is unchanged: the argument lists ride on the call they follow, as `invocations`.
 */
import { NodeType, type AstNode } from "@formula/syntax/ast";
import { parse, parseWithLayout } from "@formula/syntax/parser";
import { printAst, printAstWithLayout } from "@formula/syntax/print";
import { tokenize, tokenizeWithGaps } from "@formula/syntax/tokenizer";
import { describe, expect, it } from "vitest";

const ast = (source: string): AstNode => parse(tokenize(source));

describe("an argument list applied to a call's result", () => {
  it("parses an immediately invoked LAMBDA", () => {
    expect(ast("LAMBDA(n,n*2)(3)")).toEqual({
      type: NodeType.FunctionCall,
      name: "LAMBDA",
      args: [
        { type: NodeType.Name, name: "n" },
        {
          type: NodeType.BinaryOp,
          op: "*",
          left: { type: NodeType.Name, name: "n" },
          right: { type: NodeType.Number, value: 2 }
        }
      ],
      invocations: [[{ type: NodeType.Number, value: 3 }]]
    });
  });

  it("parses a chain of them, in order", () => {
    const node = ast("LAMBDA(x,LAMBDA(y,x+y))(1)(2)");
    expect(node).toMatchObject({
      type: NodeType.FunctionCall,
      name: "LAMBDA",
      invocations: [[{ value: 1 }], [{ value: 2 }]]
    });
  });

  it("parses one after a LET-bound name, and an empty argument list", () => {
    const node = ast("LET(F,LAMBDA(x,LAMBDA(y,x+y)),F(1)(2))") as { args: AstNode[] };
    expect(node.args[2]).toMatchObject({ name: "F", invocations: [[{ value: 2 }]] });
    expect(ast("LAMBDA(1)()")).toMatchObject({ name: "LAMBDA", invocations: [[]] });
  });

  it("binds tighter than every operator", () => {
    expect(ast("-LAMBDA(n,n)(3)^2")).toMatchObject({
      type: NodeType.BinaryOp,
      op: "^",
      left: { type: NodeType.UnaryOp, operand: { name: "LAMBDA", invocations: [[{ value: 3 }]] } }
    });
  });

  it("leaves a call with no invocation exactly as it was", () => {
    expect(Object.keys(ast("SUM(1,2)"))).toEqual(["type", "name", "args"]);
  });

  it("prints back to the same text and the same tree", () => {
    for (const source of [
      "LAMBDA(n,n*2)(3)",
      "LAMBDA(x,LAMBDA(y,x+y))(1)(2)",
      "LET(F,LAMBDA(x,LAMBDA(y,x+y)),F(1)(2))",
      "LAMBDA(1)()",
      "LAMBDA(a,b,a+b)(A1:A3,)"
    ]) {
      const printed = printAst(ast(source));
      expect(printed).toBe(source);
      expect(ast(printed)).toEqual(ast(source));
    }
  });

  it("keeps the author's whitespace inside the applied argument list", () => {
    const source = "LAMBDA(n, n*2)( 3 )";
    const { tokens, gaps } = tokenizeWithGaps(source);
    const { ast: tree, layout } = parseWithLayout(tokens, gaps);
    expect(printAstWithLayout(tree, layout)).toBe(source);
  });
});
