/**
 * `spill-scan.ts` reads the tokens the spill-range operator is made of without the formula tokenizer, which a
 * write-only bundle must not carry. Held here to `tokenizeWithSpans`: for every formula of a corpus — hand-picked edge
 * cases and every pairing of a set of fragments — the scanner emits the same tokens at the same spans, classed alike,
 * and gives up exactly where the tokenizer throws.
 */
import { ScanKind, scanFormula, type ScanToken } from "@excel/utils/spill-scan";
import { isSpillOperator, type Token, TokenType } from "@formula/syntax/token-types";
import { tokenizeWithSpans } from "@formula/syntax/tokenizer";
import { describe, expect, it } from "vitest";

const VALUE_TYPES = new Set([
  TokenType.Number,
  TokenType.String,
  TokenType.Boolean,
  TokenType.CellRef,
  TokenType.Range,
  TokenType.CloseParen,
  TokenType.CloseBrace,
  TokenType.Percent,
  TokenType.Name,
  TokenType.Error,
  TokenType.ColRange,
  TokenType.RowRange,
  TokenType.StructuredRef
]);

const REF_TYPES = new Set([
  TokenType.CellRef,
  TokenType.Range,
  TokenType.ColRange,
  TokenType.RowRange,
  TokenType.StructuredRef,
  TokenType.CloseParen,
  TokenType.Name
]);

function kindOf(token: Token): ScanKind {
  if (isSpillOperator(token)) {
    return ScanKind.Spill;
  }
  switch (token.type) {
    case TokenType.SheetRef:
      return ScanKind.Sheet;
    case TokenType.CellRef:
      return ScanKind.Cell;
    case TokenType.Function:
      return ScanKind.Function;
    case TokenType.OpenParen:
      return ScanKind.Open;
    case TokenType.CloseParen:
      return ScanKind.Close;
    case TokenType.AtSign:
      return ScanKind.At;
    case TokenType.UnaryPrefix:
      return ScanKind.Prefix;
    case TokenType.Operator:
      return ScanKind.Infix;
    case TokenType.Percent:
      return ScanKind.Percent;
    case TokenType.Colon:
      return ScanKind.Colon;
    case TokenType.Intersect:
      return ScanKind.Intersect;
    case TokenType.Comma:
      return ScanKind.Comma;
    case TokenType.OpenBrace:
      return ScanKind.OpenBrace;
    case TokenType.CloseBrace:
      return ScanKind.CloseBrace;
    case TokenType.Semicolon:
      return ScanKind.Semicolon;
    default:
      return ScanKind.Other;
  }
}

/** The operator a prefix or infix token carries, as the scanner names it. */
function operatorOf(token: Token): string | undefined {
  return !isSpillOperator(token) &&
    (token.type === TokenType.UnaryPrefix || token.type === TokenType.Operator)
    ? (token as { value: string }).value
    : undefined;
}

/** What the tokenizer says, in the scanner's terms; `undefined` where it throws. */
function expected(text: string): Omit<ScanToken, never>[] | undefined {
  let result: ReturnType<typeof tokenizeWithSpans>;
  try {
    result = tokenizeWithSpans(text);
  } catch {
    return undefined;
  }
  return result.tokens.map((token, i) => {
    const kind = kindOf(token);
    return {
      kind,
      start: result.spans[i]![0],
      end: result.spans[i]![1],
      isValue: VALUE_TYPES.has(token.type) || isSpillOperator(token),
      isRef: REF_TYPES.has(token.type) || isSpillOperator(token),
      ...(kind === ScanKind.Function
        ? { name: (token as { name: string }).name }
        : operatorOf(token) !== undefined
          ? { name: operatorOf(token) }
          : {})
    };
  });
}

const EDGE_CASES = [
  "A1#",
  "=A1#",
  "  =A1#*2",
  "$A$1#",
  "A$1#",
  "$A1#",
  "a1#",
  "XFD1048576#",
  "XFE1#",
  "A0#",
  "A1048577#",
  "AAAA1#",
  "A1 #",
  "A1##",
  "A1#B1",
  "A1#:B1",
  "A1:B2#",
  "S!A1:B2#",
  "Sheet2!A1#",
  "'My Sheet'!$A$1#",
  "'x#y'!B2#",
  "'a''b'!A1#",
  "'it''s #1'!C3#+1",
  "Sheet1:Sheet3!A1#",
  "Q1:Q4!A1#",
  "A1!B2#",
  "!A1#",
  "!$A$1#",
  "Sheet 1!A1#",
  "A1:A3 B2#",
  "-A1#",
  "+A1#%",
  "A1%#",
  "LOG10#",
  "LOG10(A1#)",
  "R1C1#",
  "_A1#",
  "A1.B#",
  "A1_#",
  "A1e5#",
  "Ä1#",
  "Name1#",
  "TRUE#",
  "SUM(A1)#",
  "{1,2}#",
  "{1,2;A1#,3}",
  "A:A#",
  "1:1#",
  "$1:$3#",
  "A:B",
  "$A:$B+1",
  "1:5",
  "$1:$5",
  '"A1#"&A1#',
  '"un"terminated',
  '"unterminated',
  'A1&"#"',
  '""""&A1#',
  "#REF!",
  "#REF!#",
  "#ref!+A1#",
  "#N/A",
  "#n/a&A1#",
  "#DIV/0!",
  "#GETTING_DATA",
  "#SPILL!",
  "#FOO",
  "IFERROR(A1#,#N/A)",
  "SUM(#REF!)",
  "Tb[[#This Row],[x]]",
  "Tb[[#This Row],[x]]#",
  "Tb[@x]*A1#",
  "Tb[@[a b]]+A1#",
  "Tb[[a]:[b]]",
  "Tb['#x]+A1#",
  "Tb[a'[b]+A1#",
  "[@x]+A1#",
  "[#Headers]",
  "[[#Data],[x]]+B2#",
  "[1]Sheet1!A1+A1#",
  "[Book.xlsx]S!A1#",
  "[Book.xlsx]S!A1:B2+C1#",
  "['My Book.xlsx']'S x'!A1+B1#",
  "[x]",
  "A1[x]",
  "SUM(A1:A3)[1]",
  "_xlfn.ANCHORARRAY(A1)",
  "_xlfn.ANCHORARRAY('My Sheet'!$A$1)+1",
  "_XLFN.anchorarray(S!B2)",
  "_xlfn.ANCHORARRAY( A1)",
  "_xlfn.ANCHORARRAY(A1 )",
  "_xlfn.ANCHORARRAY (A1)",
  "_xlfn.ANCHORARRAY(A1:B2)",
  "_xlfn.ANCHORARRAY(Name)",
  "SUM(_xlfn.ANCHORARRAY(A1),_xlfn.ANCHORARRAY(B1))",
  'LEN("_xlfn.ANCHORARRAY(A1)")',
  "1.5E+3*A1#",
  ".5+A1#",
  "1.+A1#",
  "1E",
  "1e5",
  "A1<>B1#",
  "A1<=B1#",
  "A1>=B1#",
  "A1>B1#",
  "A1^2&B1#",
  "@A1:A3+A1#",
  "A1,B1#",
  "(A1,B1#)",
  "A1;B1",
  "1~2",
  "\\Name+A1#",
  "\\\\x",
  "Name?",
  "A1]",
  "A1.",
  "A1`",
  "销售额+A1#",
  "e\u0301x+A1#",
  "A1\t#",
  "A1\n+B1#",
  "LET(x,A1#,x*2)",
  "_xlfn.LET(_xlpm.x,A1#,_xlpm.x)",
  "@A1:A3",
  "@ A1",
  "-@A1:A3",
  "1-@A1",
  "@@A1",
  "@-A1",
  "SUM(@A1:A3)*2",
  "Tb[@x]+@Tb[x]",
  "[@[a b]]",
  "@(A1)",
  "@{1,2}",
  "A1@",
  "1+-2",
  "-+1",
  "(1)-2",
  "{1,2}-1",
  "A1%-1",
  "TRUE-1",
  '"a"-1',
  "#N/A-1",
  "A1 -1",
  "1 - -1",
  "A1:A3 B1:B3",
  "A1:INDEX(B1:B3,2)",
  "{1,2;3,4}"
];

const FRAGMENTS = [
  "@",
  "-",
  "^",
  "A1",
  "#",
  "$B$2",
  "S!",
  "'a b'!",
  "(",
  ")",
  ",",
  " ",
  "+",
  "*",
  ":",
  '"x#"',
  "#N/A",
  "#REF!",
  "Tb[x]",
  "[@y]",
  "{1,2}",
  "1",
  "2.5",
  "SUM",
  "_xlfn.ANCHORARRAY",
  "Name",
  "%",
  "!",
  "[1]S!A1",
  "XFD9"
];

function corpus(): string[] {
  const out = [...EDGE_CASES];
  for (const a of FRAGMENTS) {
    for (const b of FRAGMENTS) {
      out.push(a + b);
      for (const c of ["#", "A1#", ")", "(A1)", " B1#"]) {
        out.push(a + b + c);
      }
    }
  }
  return out;
}

describe("scanFormula reads what tokenizeWithSpans reads", () => {
  it.each(EDGE_CASES)("%j", text => {
    expect(scanFormula(text)).toEqual(expected(text));
  });

  it("over every pairing of the fragments", () => {
    const mismatches: string[] = [];
    let read = 0;
    let rejected = 0;
    for (const text of corpus()) {
      const want = expected(text);
      if (want === undefined) {
        rejected++;
      } else {
        read++;
      }
      if (JSON.stringify(scanFormula(text)) !== JSON.stringify(want)) {
        mismatches.push(text);
      }
    }
    expect(mismatches).toEqual([]);
    // Both outcomes are exercised.
    expect(read).toBeGreaterThan(1000);
    expect(rejected).toBeGreaterThan(1000);
  });

  it("over random text from formula characters", () => {
    // A fixed seed: the same strings on every run.
    let seed = 0x2545f491;
    const random = (n: number): number => {
      seed ^= seed << 13;
      seed ^= seed >>> 17;
      seed ^= seed << 5;
      return (seed >>> 0) % n;
    };
    const pieces = [
      ..."AB1$#!:'\" ()[],@%+-*/&<>={};.eE_x9 \t",
      "A1",
      "XFD1",
      "S!",
      "#N/A",
      "#REF!",
      "Tb",
      "''",
      '""',
      "SUM(",
      "[@",
      "[#Data]",
      "[1]",
      "_xlfn.ANCHORARRAY(",
      "1048577",
      "销",
      "TRUE"
    ];
    const mismatches: string[] = [];
    for (let n = 0; n < 100000; n++) {
      let text = "";
      for (let k = random(12) + 1; k > 0; k--) {
        text += pieces[random(pieces.length)];
      }
      if (JSON.stringify(scanFormula(text)) !== JSON.stringify(expected(text))) {
        mismatches.push(text);
      }
    }
    expect(mismatches.slice(0, 20)).toEqual([]);
  });
});
