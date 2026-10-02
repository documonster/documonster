import { colCache } from "@excel/utils/col-cache";
import { slideFormula } from "@excel/utils/shared-formula";
import { TokenType } from "@formula/syntax/token-types";
import type { Token } from "@formula/syntax/token-types";
import { tokenize } from "@formula/syntax/tokenizer";
import { describe, it, expect } from "vitest";

describe("shared-formula", () => {
  describe("slideFormula", () => {
    const expectations = [
      { args: ["A1+1", "A2", "A3"], result: "A2+1" },
      { args: ["A1+1", "A2", "B2"], result: "B1+1" },
      { args: ["SUM(A1:A10)", "A11", "B11"], result: "SUM(B1:B10)" },
      { args: ["$A$1+A1", "A2", "A3"], result: "$A$1+A2" },
      { args: ["$A$1+A1", "A2", "B2"], result: "$A$1+B1" },
      { args: ["$A1+A1", "A2", "A3"], result: "$A2+A2" },
      { args: ["$A1+A1", "A2", "B2"], result: "$A1+B1" },
      { args: ["A$1+A1", "A2", "A3"], result: "A$1+A2" },
      { args: ["A$1+A1", "A2", "B2"], result: "B$1+B1" }
    ];
    expectations.forEach(({ args, result }) => {
      it(`${args[0]} from ${args[1]} to ${args[2]}`, () => {
        expect(slideFormula(args[0] as string, args[1] as string, args[2] as string)).toBe(result);
      });
    });
  });
});

describe("slideFormula leaves literals alone", () => {
  it.each([
    ['IF(A1="A1",1,0)', 'IF(A2="A1",1,0)'],
    ['"x""A1"&A1', '"x""A1"&A2'],
    ["'Q1 A1'!A1+Sheet2!B1", "'Q1 A1'!A2+Sheet2!B2"],
    ["'it''s A1'!A1", "'it''s A1'!A2"],
    ["Table1[Col A1]+A1", "Table1[Col A1]+A2"],
    ["SUM(Table1[[#This Row],[B2]])+B2", "SUM(Table1[[#This Row],[B2]])+B3"],
    ["Table1[Col']A1]+A1", "Table1[Col']A1]+A2"],
    ["[1]Sheet1!A1+A1", "[1]Sheet1!A2+A2"],
    ["A1:B2+$A$1+A$1", "A2:B3+$A$1+A$1"]
  ])("%s", (formula, expected) => {
    expect(slideFormula(formula, "C1", "C2")).toBe(expected);
  });
});

describe("slideFormula references", () => {
  it.each([
    ["SUM(A:A)", "C1", "D1", "SUM(B:B)"],
    ["SUM($A:B)", "C1", "D1", "SUM($A:C)"],
    ["SUM(1:1)", "C1", "C2", "SUM(2:2)"],
    ["SUM(1:$3)", "C1", "C2", "SUM(2:$3)"],
    ["Sheet2!A:A+'My Sheet'!1:2", "C1", "D2", "Sheet2!B:B+'My Sheet'!2:3"],
    ["_A1+A1+A1_B", "C1", "C2", "_A1+A2+A1_B"],
    ["LOG10(A1)+AB2!A1", "C1", "C2", "LOG10(A2)+AB2!A2"],
    ["1E5+1.5+A1", "C1", "C2", "1E5+1.5+A2"],
    ["ABCD1+A1", "C1", "C2", "ABCD1+A2"],
    ["a1+$b$2", "C1", "D2", "B2+$B$2"]
  ])("%s from %s to %s", (formula, from, to, expected) => {
    expect(slideFormula(formula, from, to)).toBe(expected);
  });

  it.each([
    ["A1+B1", "C2", "C1", "#REF!+#REF!"],
    ["SUM(A1:B2)", "C2", "C1", "SUM(#REF!)"],
    ["XFD1", "C1", "D1", "#REF!"],
    ["SUM(A:B)", "D1", "C1", "SUM(#REF!)"],
    ["Sheet2!A1", "C2", "C1", "Sheet2!#REF!"]
  ])("turns %s slid off the sheet into #REF!", (formula, from, to, expected) => {
    expect(slideFormula(formula, from, to)).toBe(expected);
  });
});

/**
 * The formula engine's tokenizer decides what a reference is; `slideFormula` must agree with it. Tokenizing a slid
 * formula has to give the original tokens with exactly the reference tokens moved — anything else is a reference
 * missed, or a name, string or number changed.
 *
 * Unquoted external references (`[1]Sheet1!A1`) are left out: the engine does not evaluate them, and its tokenizer
 * reads only part of one, so it cannot judge them.
 */
describe("slideFormula agrees with the formula tokenizer", () => {
  function shiftCell(ref: string, dc: number, dr: number): string {
    const m = /^(\$?)([A-Z]+)(\$?)(\d+)$/.exec(ref.toUpperCase())!;
    const col = m[1] ? colCache.l2n(m[2]) : colCache.l2n(m[2]) + dc;
    const row = m[3] ? Number(m[4]) : Number(m[4]) + dr;
    return m[1] + colCache.n2l(col) + m[3] + row;
  }

  function moved(token: Token, dc: number, dr: number): Token {
    switch (token.type) {
      case TokenType.CellRef: {
        const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(
          shiftCell(
            `${token.colAbsolute ? "$" : ""}${token.col}${token.rowAbsolute ? "$" : ""}${token.row}`,
            dc,
            dr
          )
        )!;
        return { ...token, col: m[1], row: m[2] };
      }
      case TokenType.Range:
        return {
          ...token,
          value: token.value
            .split(":")
            .map(p => shiftCell(p, dc, dr))
            .join(":")
        };
      case TokenType.ColRange:
        return {
          ...token,
          value: token.value
            .toUpperCase()
            .split(":")
            .map(p => (p.startsWith("$") ? p : colCache.n2l(colCache.l2n(p) + dc)))
            .join(":")
        };
      case TokenType.RowRange:
        return {
          ...token,
          value: token.value
            .split(":")
            .map(p => (p.startsWith("$") ? p : String(Number(p) + dr)))
            .join(":")
        };
      default:
        return token;
    }
  }

  it("on 5,000 generated formulas", () => {
    let seed = 12345;
    const random = () => (seed = (seed * 1103515245 + 12345) >>> 0) / 2 ** 32;
    const pick = <T>(items: readonly T[]): T => items[Math.floor(random() * items.length)];
    const atom = (): string =>
      pick([
        () =>
          pick(["", "$"]) +
          pick(["A", "B", "Z", "AA", "XF", "c"]) +
          pick(["", "$"]) +
          pick(["1", "9", "100"]),
        () => "A1:" + pick(["B2", "$C$3", "c4"]),
        () => pick(["A:A", "$B:C", "1:1", "$2:3"]),
        () =>
          pick(["Sheet1!", "'My Sheet'!", "'it''s'!", "Sheet1:Sheet3!", "'[a b.xlsx]S 1'!"]) +
          pick(["A1", "B2:C3", "D:D", "4:5"]),
        () => pick(['"A1"', '"x""B2"', '""']),
        () => pick(["Table1[Col]", "Table1[[#This Row],[A1]]", "Table1[@Q1]", "[@[B2]]"]),
        () =>
          pick(["1", "1.5", "1E5", "2.5E-3", ".5", "TRUE", "FALSE", "#REF!", "#N/A", "#DIV/0!"]),
        () => pick(["MyName", "_A1", "A1B", "Q1_2024", "ABCD1"]),
        () => "{1,2;3,4}"
      ])();
    const expression = (depth: number): string =>
      depth > 2 || random() < 0.4
        ? atom()
        : pick([
            () =>
              `${pick(["SUM", "IF", "LOG10", "_xlfn.XLOOKUP", "INDEX"])}(${expression(depth + 1)},${expression(depth + 1)})`,
            () =>
              `${expression(depth + 1)}${pick(["+", "-", "*", "&", "=", "<>", " "])}${expression(depth + 1)}`,
            () => `(${expression(depth + 1)})`,
            () => `-${expression(depth + 1)}%`
          ])();

    for (let i = 0; i < 5000; i++) {
      const formula = expression(0);
      const slid = slideFormula(formula, "A1", "C4");
      expect(tokenize(slid), `${formula} → ${slid}`).toEqual(
        tokenize(formula).map(t => moved(t, 2, 3))
      );
    }
  });
});
