import { FormulaArgumentError, FormulaParseError } from "@formula/errors";
import { NodeType } from "@formula/syntax/ast";
import { parseFormula as parse } from "@formula/syntax/parser";
import { printAst } from "@formula/syntax/print";
import { TokenType } from "@formula/syntax/token-types";
import { tokenize, tokenizeWithSpans } from "@formula/syntax/tokenizer";
import { describe, expect, it } from "vitest";

const roundTrip = (formula: string): string => printAst(parse(formula));

describe("spill-range operator `A1#`", () => {
  it("tokenizes as an operator straight after the cell reference", () => {
    expect(tokenize("=A1#")).toEqual([
      { type: TokenType.CellRef, col: "A", row: "1", colAbsolute: false, rowAbsolute: false },
      { type: TokenType.Operator, value: "#" }
    ]);
    const { spans } = tokenizeWithSpans("SUM($B$2#)");
    expect(spans[3]).toEqual([8, 9]);
  });

  it("parses to a postfix unary operation on the reference", () => {
    expect(parse("=A1#")).toMatchObject({
      type: NodeType.UnaryOp,
      op: "#",
      operand: { type: NodeType.CellRef, colAbsolute: false, rowAbsolute: false }
    });
    const node = parse("=Sheet2!$C$3#") as { type: NodeType; operand: { sheet?: string } };
    expect(node.type).toBe(NodeType.UnaryOp);
    expect(node.operand.sheet).toBe("Sheet2");
  });

  it("round-trips exactly in both spellings", () => {
    for (const formula of [
      "A1#",
      "SUM(A1#)",
      "A1#*2",
      "-A1#",
      "Sheet1!A1#",
      "'My Sheet'!$B$2#+1",
      "ROWS(A1#)",
      "A1# B1",
      "_xlfn.ANCHORARRAY(A1)",
      "SUM(_xlfn.ANCHORARRAY(Sheet1!A1))"
    ]) {
      expect(roundTrip(formula)).toBe(formula);
    }
  });

  it("binds tighter than every other operator", () => {
    const node = parse("=-A1#") as { op: string; operand: { op: string } };
    expect(node.op).toBe("-");
    expect(node.operand.op).toBe("#");
  });

  it("refuses `#` anywhere else", () => {
    for (const formula of ["=A1 #", "=A1:B2#", "=(A1)#", "=Name#", "=#"]) {
      expect(() => parse(formula), formula).toThrow(FormulaParseError);
    }
  });

  it("refuses to print the operator on something that is not a cell reference", () => {
    expect(() =>
      printAst({
        type: NodeType.UnaryOp,
        op: "#",
        operand: { type: NodeType.Number, value: 1 }
      })
    ).toThrow(FormulaArgumentError);
  });
});

describe("error literals", () => {
  it("keeps every error value Excel has", () => {
    for (const literal of [
      "#NULL!",
      "#DIV/0!",
      "#VALUE!",
      "#REF!",
      "#NAME?",
      "#NUM!",
      "#N/A",
      "#GETTING_DATA",
      "#SPILL!",
      "#CALC!",
      "#FIELD!",
      "#BLOCKED!",
      "#CONNECT!",
      "#UNKNOWN!",
      "#BUSY!"
    ]) {
      expect(roundTrip(`${literal}+1`)).toBe(`${literal}+1`);
    }
  });

  it("refuses an unknown one instead of reading it as #NAME?", () => {
    expect(() => tokenize("=#FOOBAR!+1")).toThrow(FormulaParseError);
    try {
      tokenize("=1+#FOOBAR!");
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(FormulaParseError);
      expect((error as FormulaParseError).message).toContain('"#FOOBAR!"');
      expect((error as FormulaParseError).position).toBe(3);
    }
  });
});
