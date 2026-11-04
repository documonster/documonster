/**
 * Contracts of `Formula.tokenize` / `parse` / `format` that the API review found broken: an
 * external-workbook reference turned into `#REF!` by a parse → format round trip, inputs and
 * hand-built trees that surfaced as a `TypeError` or printed as text that does not parse
 * back, and `FormulaParseError.position` meaning a token index or a character offset
 * depending on where the error was found.
 */
import {
  Formula,
  FormulaArgumentError,
  FormulaError,
  FormulaParseError,
  NodeType
} from "@formula/index";
import type {
  AstNode,
  FormulaCellRefNode as CellRefNode,
  FormulaErrorNode as ErrorNode,
  FormulaStructuredRefNode as StructuredRefNode
} from "@formula/index";
import { describe, expect, it } from "vitest";

describe("external-workbook references survive a round trip", () => {
  it.each([
    ["=[1]Sheet1!A1", "[1]Sheet1!A1"],
    ["=[Book2.xlsx]Sheet1!A1", "[Book2.xlsx]Sheet1!A1"],
    ["=SUM([1]Sheet1!A1:B2)", "SUM([1]Sheet1!A1:B2)"],
    ["=[1]Data!$A$1+1", "[1]Data!$A$1+1"]
  ])("%s", (source, printed) => {
    const ast = Formula.parse(source);
    expect(Formula.format(ast)).toBe(printed);
    expect(Formula.parse(Formula.format(ast))).toEqual(ast);
  });

  it("still evaluates as #REF!, which the node says", () => {
    const node = Formula.parse("=[1]Sheet1!A1") as ErrorNode;
    expect(node).toEqual({
      type: NodeType.Error,
      value: "#REF!",
      externalReference: "[1]Sheet1!A1"
    });
  });
});

describe("structured-reference column names are escaped", () => {
  it.each(["a]b", "a[b", "Item #", "it's"])("%j", column => {
    const node: StructuredRefNode = {
      type: NodeType.StructuredRef,
      tableName: "T",
      columns: [column],
      specials: []
    };
    expect(Formula.parse(Formula.format(node))).toEqual(node);
  });

  it("reads Excel's escaped single-bracket form", () => {
    expect(Formula.parse("=T[a']b]")).toEqual({
      type: NodeType.StructuredRef,
      tableName: "T",
      columns: ["a]b"],
      specials: []
    });
    expect((Formula.parse("=T[@a'#b]") as StructuredRefNode).columns).toEqual(["a#b"]);
  });
});

describe("inputs are checked with FormulaError", () => {
  it.each([42, null, undefined])("Formula.tokenize(%s)", input => {
    expect(() => Formula.tokenize(input as never)).toThrow(FormulaError);
  });

  it.each([42, null, undefined, {}])("Formula.parse(%s)", input => {
    expect(() => Formula.parse(input as never)).toThrow(FormulaError);
  });

  it.each<[string, unknown]>([
    ["null", null],
    ["an empty object", {}],
    ["an unknown node type", { type: 999 }],
    ["a NaN number", { type: NodeType.Number, value: NaN }],
    ["Infinity", { type: NodeType.Number, value: Infinity }],
    [
      "an unknown operator",
      {
        type: NodeType.BinaryOp,
        op: "$$",
        left: { type: NodeType.Number, value: 1 },
        right: { type: NodeType.Number, value: 2 }
      }
    ],
    [
      "an unknown prefix",
      { type: NodeType.UnaryOp, op: "!", operand: { type: NodeType.Number, value: 1 } }
    ],
    ["a function name with a space", { type: NodeType.FunctionCall, name: "MY FN", args: [] }]
  ])("Formula.format refuses %s", (_label, node) => {
    expect(() => Formula.format(node as AstNode)).toThrow(FormulaError);
  });

  it("formats every node Formula.parse produces", () => {
    for (const source of [
      "=-1^2",
      "=SUM((A1,B1))",
      "=@A1:A10",
      "=A1:B2 B1:C3",
      "=_xlfn.XLOOKUP(1,A:A,B:B)"
    ]) {
      expect(() => Formula.format(Formula.parse(source))).not.toThrow();
    }
  });
});

describe("FormulaParseError.position", () => {
  it("is a character offset when parsing text, from the parser as from the tokenizer", () => {
    const parserError = catchParse("=LONGFUNCTIONNAME(1) 2");
    expect(parserError.position).toBe("=LONGFUNCTIONNAME(1) ".length);
    const tokenizerError = catchParse('="abc');
    expect(tokenizerError.position).toBe(5);
  });

  it("names the token, without a position, when parsing tokens", () => {
    const error = catchParse(Formula.tokenize("=1 2"));
    expect(error.position).toBeUndefined();
    expect(error.message).toMatch(/at token 1/);
  });

  it("names token kinds rather than numbers", () => {
    expect(catchParse("=SUM(1,2").message).toMatch(/Expected CloseParen/);
    expect(catchParse("=1)").message).toMatch(/Unexpected trailing CloseParen/);
  });
});

describe("the AST member types are published", () => {
  it("lets a visitor name the node it handles", () => {
    const cell = Formula.parse("=$A$1") as CellRefNode;
    expect(cell.colAbsolute && cell.rowAbsolute).toBe(true);
  });
});

function catchParse(source: Parameters<typeof Formula.parse>[0]): FormulaParseError {
  try {
    Formula.parse(source);
  } catch (error) {
    expect(error).toBeInstanceOf(FormulaParseError);
    return error as FormulaParseError;
  }
  throw new Error("expected a parse error");
}

describe("r3: sheet names are quoted by Excel's rule (TX-B3)", () => {
  it.each([
    ["='A1'!B2", "'A1'!B2"],
    ["='R1C1'!A1", "'R1C1'!A1"],
    ["='TRUE'!A1", "'TRUE'!A1"],
    ["='a.b'!A1", "'a.b'!A1"],
    ["=Sheet1!A1", "Sheet1!A1"]
  ])("%s", (source, printed) => {
    expect(Formula.format(Formula.parse(source))).toBe(printed);
  });
});

describe("r3: unknown characters are refused, and a name may start with a backslash (TX-B4)", () => {
  it("refuses a character no token starts with, at its offset", () => {
    const error = catchParse("=1~2");
    expect(error.position).toBe(2);
    expect(() => Formula.tokenize("=a?b")).toThrow(FormulaParseError);
  });

  it("reads \\Name as one name", () => {
    expect(Formula.parse("=\\Name")).toEqual({ type: NodeType.Name, name: "\\Name" });
    expect(Formula.format(Formula.parse("=\\Name"))).toBe("\\Name");
    expect(() => Formula.parse("=a\\b")).toThrow(FormulaParseError);
  });

  it("still reads a defined name's sheetless !$A$1", () => {
    expect(() => Formula.tokenize("!$A$1:$B$2")).not.toThrow();
  });
});

describe("r3: future-function prefixes are written in lower case (question 5)", () => {
  it.each([
    ["=_xlfn.XLOOKUP(1,A:A,B:B)", "_xlfn.XLOOKUP(1,A:A,B:B)"],
    ["=_xlfn._xlws.SORT(A1:A3)", "_xlfn._xlws.SORT(A1:A3)"],
    ["=_XLUDF.MYFN(1)", "_xludf.MYFN(1)"],
    ["=sum(1)", "SUM(1)"]
  ])("%s", (source, printed) => {
    expect(Formula.format(Formula.parse(source))).toBe(printed);
  });
});

describe("X12 / X10: argument errors are FormulaArgumentError with a field", () => {
  it.each<[string, () => unknown, string]>([
    ["tokenize", () => Formula.tokenize(42 as never), "formula"],
    ["parse", () => Formula.parse({} as never), "source"],
    ["format", () => Formula.format(null as never), "node"],
    ["format a bad tree", () => Formula.format({ type: NodeType.Number, value: NaN }), "node"]
  ])("%s", (_label, fn, field) => {
    let error: unknown;
    try {
      fn();
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(FormulaArgumentError);
    expect(error).toBeInstanceOf(FormulaError);
    expect((error as FormulaArgumentError).field).toBe(field);
    expect((error as FormulaArgumentError).toJSON()).toMatchObject({ field });
  });

  it("is not used for a syntax error, and the internal detail is not serialised", () => {
    const error = catchParse("=1+");
    expect(error).not.toBeInstanceOf(FormulaArgumentError);
    expect(Object.keys(error.toJSON())).not.toContain("detail");
    expect(error.toJSON()).toMatchObject({ position: 3 });
  });
});
