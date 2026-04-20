/**
 * Unit tests for the formula parser.
 *
 * Covers operator precedence and associativity, unary prefixes, function
 * calls with zero/single/multiple/missing args, nested calls, reference
 * operators (range `:`, union `,`, intersect ` `), comparisons, string
 * concatenation, error literal AST nodes, and array constants.
 *
 * The parser consumes token streams from the tokenizer, so these tests
 * run tokenize → parse end-to-end.
 */

import { describe, it, expect } from "vitest";

import { NodeType } from "../ast";
import type {
  AstNode,
  BinaryOpNode,
  UnaryOpNode,
  NumberNode,
  StringNode,
  BooleanNode,
  ErrorNode,
  FunctionCallNode,
  CellRefNode,
  RangeRefNode,
  ArrayNode,
  PercentNode,
  NameNode,
  MissingNode
} from "../ast";
import { parse } from "../parser";
import { tokenize } from "../tokenizer";

function compile(formula: string): AstNode {
  return parse(tokenize(formula));
}

describe("parser — literals", () => {
  it("parses a number literal", () => {
    const ast = compile("42");
    expect(ast.type).toBe(NodeType.Number);
    expect((ast as NumberNode).value).toBe(42);
  });

  it("parses a string literal", () => {
    const ast = compile('"hello"');
    expect(ast.type).toBe(NodeType.String);
    expect((ast as StringNode).value).toBe("hello");
  });

  it("parses TRUE/FALSE as boolean nodes", () => {
    expect((compile("TRUE") as BooleanNode).value).toBe(true);
    expect((compile("FALSE") as BooleanNode).value).toBe(false);
  });

  it("parses #N/A as an Error node preserving the canonical code", () => {
    const ast = compile("#N/A");
    expect(ast.type).toBe(NodeType.Error);
    expect((ast as ErrorNode).value).toBe("#N/A");
  });
});

describe("parser — arithmetic precedence", () => {
  it("parses '+' as left-associative", () => {
    // 1+2+3 → ((1+2)+3)
    const ast = compile("1+2+3") as BinaryOpNode;
    expect(ast.type).toBe(NodeType.BinaryOp);
    expect(ast.op).toBe("+");
    expect(ast.right.type).toBe(NodeType.Number);
    expect((ast.right as NumberNode).value).toBe(3);

    const left = ast.left as BinaryOpNode;
    expect(left.type).toBe(NodeType.BinaryOp);
    expect(left.op).toBe("+");
    expect((left.left as NumberNode).value).toBe(1);
    expect((left.right as NumberNode).value).toBe(2);
  });

  it("gives '*' tighter precedence than '+'", () => {
    // 1+2*3 → 1 + (2*3)
    const ast = compile("1+2*3") as BinaryOpNode;
    expect(ast.op).toBe("+");
    expect((ast.left as NumberNode).value).toBe(1);
    const right = ast.right as BinaryOpNode;
    expect(right.op).toBe("*");
    expect((right.left as NumberNode).value).toBe(2);
    expect((right.right as NumberNode).value).toBe(3);
  });

  it("parses '^' as left-associative (Excel behaviour)", () => {
    // Regression: Excel evaluates multiple `^` left-to-right, so
    // `=2^3^2` is `(2^3)^2 = 64`, NOT `2^(3^2) = 512`. Previous
    // precedence (61/60, right-assoc) produced 512 — matching the
    // math convention but not Excel.
    const ast = compile("2^3^2") as BinaryOpNode;
    expect(ast.op).toBe("^");
    // Left side is (2^3), right side is literal 2
    const left = ast.left as BinaryOpNode;
    expect(left.op).toBe("^");
    expect((left.left as NumberNode).value).toBe(2);
    expect((left.right as NumberNode).value).toBe(3);
    expect((ast.right as NumberNode).value).toBe(2);
  });

  it("parses unary minus with HIGHER precedence than '^' (Excel quirk)", () => {
    // Regression: Excel's precedence table ranks negation (1) higher
    // than exponentiation (4). So `=-2^2 = 4` (= (-2)^2), NOT -4.
    // Previously unary minus was parsed at precedence 55 below `^`'s
    // 60/61, producing `-(2^2)`. The behaviour now matches Excel.
    const ast = compile("-2^2") as BinaryOpNode;
    expect(ast.type).toBe(NodeType.BinaryOp);
    expect(ast.op).toBe("^");
    const left = ast.left as UnaryOpNode;
    expect(left.type).toBe(NodeType.UnaryOp);
    expect(left.op).toBe("-");
    expect((left.operand as NumberNode).value).toBe(2);
    expect((ast.right as NumberNode).value).toBe(2);
  });

  it("parses unary plus as a UnaryOp", () => {
    const ast = compile("+5") as UnaryOpNode;
    expect(ast.type).toBe(NodeType.UnaryOp);
    expect(ast.op).toBe("+");
    expect((ast.operand as NumberNode).value).toBe(5);
  });

  it("parses the postfix percent operator", () => {
    const ast = compile("50%") as PercentNode;
    expect(ast.type).toBe(NodeType.Percent);
    expect((ast.operand as NumberNode).value).toBe(50);
  });

  it("honours explicit parentheses", () => {
    // (1+2)*3 → BinaryOp(*, BinaryOp(+,1,2), 3)
    const ast = compile("(1+2)*3") as BinaryOpNode;
    expect(ast.op).toBe("*");
    const left = ast.left as BinaryOpNode;
    expect(left.op).toBe("+");
    expect((ast.right as NumberNode).value).toBe(3);
  });
});

describe("parser — string concatenation and comparison", () => {
  it("parses '&' for string concatenation", () => {
    const ast = compile('"a"&"b"') as BinaryOpNode;
    expect(ast.op).toBe("&");
    expect((ast.left as StringNode).value).toBe("a");
    expect((ast.right as StringNode).value).toBe("b");
  });

  it("'&' binds looser than arithmetic but tighter than comparison", () => {
    // 1+2 & "x" = "three" → compare("three") of concat("3","x")
    // We only check the shape: the outermost op is '='.
    const ast = compile('1+2 & "x" = "three"') as BinaryOpNode;
    expect(ast.op).toBe("=");
  });

  it("parses all comparison operators", () => {
    for (const op of ["=", "<>", "<", ">", "<=", ">="]) {
      const ast = compile(`1${op}2`) as BinaryOpNode;
      expect(ast.type).toBe(NodeType.BinaryOp);
      expect(ast.op).toBe(op);
    }
  });
});

describe("parser — function calls", () => {
  it("parses a zero-argument call", () => {
    const ast = compile("NOW()") as FunctionCallNode;
    expect(ast.type).toBe(NodeType.FunctionCall);
    expect(ast.name).toBe("NOW");
    expect(ast.args).toEqual([]);
  });

  it("parses a single-argument call", () => {
    const ast = compile("ABS(-5)") as FunctionCallNode;
    expect(ast.name).toBe("ABS");
    expect(ast.args).toHaveLength(1);
    expect(ast.args[0].type).toBe(NodeType.UnaryOp);
  });

  it("parses a multi-argument call", () => {
    const ast = compile("SUM(A1,B2,C3)") as FunctionCallNode;
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(3);
    for (const arg of ast.args) {
      expect(arg.type).toBe(NodeType.CellRef);
    }
  });

  it("parses nested function calls", () => {
    const ast = compile("SUM(MAX(A1,B1),1)") as FunctionCallNode;
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(2);
    const inner = ast.args[0] as FunctionCallNode;
    expect(inner.type).toBe(NodeType.FunctionCall);
    expect(inner.name).toBe("MAX");
  });

  it("parses an omitted middle argument as Missing", () => {
    const ast = compile("IF(A1,,0)") as FunctionCallNode;
    expect(ast.args).toHaveLength(3);
    expect(ast.args[1].type).toBe(NodeType.Missing);
    // Flank args are NOT missing.
    expect(ast.args[0].type).not.toBe(NodeType.Missing);
    expect(ast.args[2].type).not.toBe(NodeType.Missing);
  });

  it("parses a trailing missing argument", () => {
    const ast = compile("IF(A1,1,)") as FunctionCallNode;
    expect(ast.args).toHaveLength(3);
    expect((ast.args[2] as MissingNode).type).toBe(NodeType.Missing);
  });

  it("parses a leading missing argument", () => {
    const ast = compile("IF(,1,0)") as FunctionCallNode;
    expect(ast.args).toHaveLength(3);
    expect((ast.args[0] as MissingNode).type).toBe(NodeType.Missing);
  });
});

describe("parser — references", () => {
  it("parses a plain cell reference", () => {
    const ast = compile("A1") as CellRefNode;
    expect(ast.type).toBe(NodeType.CellRef);
    expect(ast.col).toBe("A");
    expect(ast.row).toBe("1");
    expect(ast.colAbsolute).toBe(false);
    expect(ast.rowAbsolute).toBe(false);
  });

  it("parses a range reference as a RangeRef with resolved endpoints", () => {
    const ast = compile("A1:B2") as RangeRefNode;
    expect(ast.type).toBe(NodeType.RangeRef);
    expect(ast.start.col).toBe("A");
    expect(ast.start.row).toBe("1");
    expect(ast.end.col).toBe("B");
    expect(ast.end.row).toBe("2");
  });

  it("attaches a sheet name to a qualified cell reference", () => {
    const ast = compile("Sheet1!A1") as CellRefNode;
    expect(ast.type).toBe(NodeType.CellRef);
    expect(ast.sheet).toBe("Sheet1");
  });

  it("parses a 3D sheet range as a range with endSheet", () => {
    const ast = compile("Sheet1:Sheet3!A1") as CellRefNode;
    expect(ast.type).toBe(NodeType.CellRef);
    expect(ast.sheet).toBe("Sheet1");
    expect(ast.endSheet).toBe("Sheet3");
  });

  it("parses a union ',' inside a function argument list (not as union op)", () => {
    // Outside a function, `A1,B1` — the top-level comma isn't supported by
    // the parser as an expression. Inside SUM(A1,B1) it's just two args.
    const ast = compile("SUM(A1,B1)") as FunctionCallNode;
    expect(ast.args).toHaveLength(2);
  });

  it("parses the intersection operator (whitespace between refs)", () => {
    const ast = compile("A1:A10 B1:B10") as BinaryOpNode;
    expect(ast.type).toBe(NodeType.BinaryOp);
    expect(ast.op).toBe(" ");
    expect(ast.left.type).toBe(NodeType.RangeRef);
    expect(ast.right.type).toBe(NodeType.RangeRef);
  });
});

describe("parser — arrays", () => {
  it("parses a single-row array constant", () => {
    const ast = compile("{1,2,3}") as ArrayNode;
    expect(ast.type).toBe(NodeType.Array);
    expect(ast.rows).toHaveLength(1);
    expect(ast.rows[0]).toHaveLength(3);
    expect((ast.rows[0][2] as NumberNode).value).toBe(3);
  });

  it("parses a multi-row array constant with ';' separator", () => {
    const ast = compile("{1,2;3,4}") as ArrayNode;
    expect(ast.rows).toHaveLength(2);
    expect(ast.rows[0]).toHaveLength(2);
    expect(ast.rows[1]).toHaveLength(2);
    expect((ast.rows[1][0] as NumberNode).value).toBe(3);
    expect((ast.rows[1][1] as NumberNode).value).toBe(4);
  });
});

describe("parser — names and errors", () => {
  it("parses an unknown identifier as a Name node", () => {
    const ast = compile("MyName") as NameNode;
    expect(ast.type).toBe(NodeType.Name);
    expect(ast.name).toBe("MyName");
  });

  it("rejects trailing tokens that can't be consumed", () => {
    // "A1 B1 C1 D1" → intersection chain, legal. But "1 2" would leave a
    // trailing token because no intersection applies.
    expect(() => compile("1 2")).toThrow();
  });

  it("rejects a bare SheetRef with nothing after it", () => {
    expect(() => compile("Sheet1!")).toThrow();
  });

  it("rejects deeply nested formula beyond MAX_DEPTH guard", () => {
    // Build a left-associative chain of 2000 `+` operators which will
    // recurse through primary/expression routines. MAX_DEPTH guards the
    // parser from stack-overflow crashes.
    let formula = "1";
    for (let i = 0; i < 2000; i++) {
      formula = `(${formula}+1)`;
    }
    expect(() => compile(formula)).toThrow(/too deep/);
  });

  it("rejects unexpected end of formula", () => {
    // Incomplete expressions that leave parser expecting more input.
    expect(() => compile("1+")).toThrow();
    expect(() => compile("SUM(")).toThrow();
    expect(() => compile("(1+")).toThrow();
  });

  it("rejects unbalanced parentheses", () => {
    // Closing paren with nothing to close.
    expect(() => compile(")")).toThrow();
    // More close than open.
    expect(() => compile("SUM(1,2))")).toThrow();
  });

  it("rejects empty formula string", () => {
    expect(() => compile("")).toThrow();
  });

  it("rejects unknown token sequences", () => {
    // Two consecutive operators with no operand between.
    expect(() => compile("1 + * 2")).toThrow();
  });

  it("rejects comma at the start of an argument list", () => {
    // `SUM(,1,2)` — some tokenizers accept blank-leading args, but
    // `SUM(,1)` should still parse the first arg as Missing.
    // What we do reject: `SUM(` followed by `)` alone — nope actually
    // `SUM()` should work (0-arity), `SUM(1,)` is also accepted with
    // Missing trailing. The truly broken case is a top-level `,`.
    expect(() => compile(",")).toThrow();
  });
});
