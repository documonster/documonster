/**
 * Token Types for Formula Tokenizer
 *
 * Defines a proper discriminated union of all token types.
 * Each token variant carries only the data relevant to its type.
 */

// ============================================================================
// Token Type Enum
// ============================================================================

export const enum TokenType {
  Number = 1,
  String = 2,
  Boolean = 3,
  Error = 4,
  CellRef = 5,
  Range = 6, // e.g. A1:B10
  SheetRef = 7, // sheet prefix, e.g. Sheet1! or 'Sheet Name'!
  Function = 8, // function name (followed by '(')
  Operator = 9, // + - * / ^ & = < > <> <= >=
  OpenParen = 10,
  CloseParen = 11,
  Comma = 12,
  Colon = 13, // standalone : (for building ranges)
  Percent = 14,
  OpenBrace = 15, // { for array constants
  CloseBrace = 16, // } for array constants
  Semicolon = 17, // ; for array row separators
  UnaryPrefix = 18, // unary + or -
  Name = 19, // named range / defined name
  ColRange = 20, // whole-column range e.g. A:B, $C:$D
  RowRange = 21, // whole-row range e.g. 1:5, $3:$7
  StructuredRef = 22, // structured reference e.g. Table1[Column], [@Column]
  AtSign = 23 // @ implicit intersection prefix (Excel 365)
}

// ============================================================================
// Individual Token Interfaces
// ============================================================================

export interface NumberToken {
  type: TokenType.Number;
  value: string;
}

export interface StringToken {
  type: TokenType.String;
  value: string;
}

export interface BooleanToken {
  type: TokenType.Boolean;
  value: "TRUE" | "FALSE";
}

export interface ErrorToken {
  type: TokenType.Error;
  value: string;
}

export interface CellRefToken {
  type: TokenType.CellRef;
  /** Uppercase column letters (e.g. "A") */
  col: string;
  /** Row number as string (e.g. "1") */
  row: string;
  /** Whether column is absolute ($A) */
  colAbsolute: boolean;
  /** Whether row is absolute ($1) */
  rowAbsolute: boolean;
}

export interface RangeToken {
  type: TokenType.Range;
  /** Raw range string, e.g. "A1:B2" or "$A$1:$B$2" */
  value: string;
}

export interface SheetRefToken {
  type: TokenType.SheetRef;
  /** The sheet name (unquoted) */
  sheetName: string;
  /** The end sheet name for 3D references (e.g. Sheet1:Sheet3!) */
  endSheetName?: string;
}

export interface FunctionToken {
  type: TokenType.Function;
  /** Uppercase function name */
  name: string;
}

export interface OperatorToken {
  type: TokenType.Operator;
  value: string;
}

export interface OpenParenToken {
  type: TokenType.OpenParen;
}

export interface CloseParenToken {
  type: TokenType.CloseParen;
}

export interface CommaToken {
  type: TokenType.Comma;
}

export interface ColonToken {
  type: TokenType.Colon;
}

export interface PercentToken {
  type: TokenType.Percent;
}

export interface OpenBraceToken {
  type: TokenType.OpenBrace;
}

export interface CloseBraceToken {
  type: TokenType.CloseBrace;
}

export interface SemicolonToken {
  type: TokenType.Semicolon;
}

export interface UnaryPrefixToken {
  type: TokenType.UnaryPrefix;
  value: "+" | "-";
}

export interface NameToken {
  type: TokenType.Name;
  value: string;
}

export interface ColRangeToken {
  type: TokenType.ColRange;
  value: string;
}

export interface RowRangeToken {
  type: TokenType.RowRange;
  value: string;
}

export interface StructuredRefToken {
  type: TokenType.StructuredRef;
  /** Table name (empty string for implicit [@Col] syntax) */
  tableName: string;
  /** Column references (e.g. ["Column1", "Column2"]) */
  columns: string[];
  /** Special items (e.g. ["#Headers", "#Data"]) */
  specials: string[];
}

export interface AtSignToken {
  type: TokenType.AtSign;
}

// ============================================================================
// Discriminated Union
// ============================================================================

export type Token =
  | NumberToken
  | StringToken
  | BooleanToken
  | ErrorToken
  | CellRefToken
  | RangeToken
  | SheetRefToken
  | FunctionToken
  | OperatorToken
  | OpenParenToken
  | CloseParenToken
  | CommaToken
  | ColonToken
  | PercentToken
  | OpenBraceToken
  | CloseBraceToken
  | SemicolonToken
  | UnaryPrefixToken
  | NameToken
  | ColRangeToken
  | RowRangeToken
  | StructuredRefToken
  | AtSignToken;
