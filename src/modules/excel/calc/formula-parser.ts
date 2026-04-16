/**
 * Formula Parser
 *
 * Converts a token stream into an Abstract Syntax Tree (AST).
 * Implements a Pratt parser (top-down operator precedence) for correct
 * handling of operator precedence and associativity.
 *
 * Excel operator precedence (highest to lowest):
 *   1. Unary + / - / %
 *   2. ^  (exponentiation, right-assoc)
 *   3. * /
 *   4. + -
 *   5. &  (concatenation)
 *   6. = <> < > <= >=  (comparison)
 */

import { TokenType, type Token } from "./formula-tokenizer";

// ============================================================================
// AST Node Types
// ============================================================================

export const enum NodeType {
  Number = 1,
  String = 2,
  Boolean = 3,
  Error = 4,
  CellRef = 5,
  RangeRef = 6,
  BinaryOp = 7,
  UnaryOp = 8,
  FunctionCall = 9,
  Array = 10,
  Percent = 11,
  Name = 12,
  ColRangeRef = 13,
  RowRangeRef = 14,
  StructuredRef = 15,
  Missing = 16
}

export interface NumberNode {
  type: NodeType.Number;
  value: number;
}

export interface StringNode {
  type: NodeType.String;
  value: string;
}

export interface BooleanNode {
  type: NodeType.Boolean;
  value: boolean;
}

export interface ErrorNode {
  type: NodeType.Error;
  value: string;
}

export interface CellRefNode {
  type: NodeType.CellRef;
  /** Uppercase column letters, e.g. "A" */
  col: string;
  /** Row number as string, e.g. "1" */
  row: string;
  colAbsolute: boolean;
  rowAbsolute: boolean;
  /** Optional sheet name for cross-sheet references */
  sheet?: string;
  /** Optional end sheet name for 3D references (e.g. Sheet1:Sheet3!A1) */
  endSheet?: string;
}

export interface RangeRefNode {
  type: NodeType.RangeRef;
  /** Top-left cell */
  start: CellRefNode;
  /** Bottom-right cell */
  end: CellRefNode;
  /** Sheet name — both start and end share same sheet */
  sheet?: string;
  /** Optional end sheet name for 3D references (e.g. Sheet1:Sheet3!A1:B2) */
  endSheet?: string;
}

export interface BinaryOpNode {
  type: NodeType.BinaryOp;
  op: string;
  left: AstNode;
  right: AstNode;
}

export interface UnaryOpNode {
  type: NodeType.UnaryOp;
  op: string;
  operand: AstNode;
}

export interface FunctionCallNode {
  type: NodeType.FunctionCall;
  name: string;
  args: AstNode[];
}

export interface ArrayNode {
  type: NodeType.Array;
  /** rows[i][j] = element */
  rows: AstNode[][];
}

export interface PercentNode {
  type: NodeType.Percent;
  operand: AstNode;
}

export interface NameNode {
  type: NodeType.Name;
  name: string;
}

export interface ColRangeRefNode {
  type: NodeType.ColRangeRef;
  /** Start column (e.g. "A") */
  startCol: string;
  /** End column (e.g. "B") */
  endCol: string;
  /** Optional sheet name */
  sheet?: string;
  /** Optional end sheet name for 3D references */
  endSheet?: string;
}

export interface RowRangeRefNode {
  type: NodeType.RowRangeRef;
  /** Start row (1-based) */
  startRow: number;
  /** End row (1-based) */
  endRow: number;
  /** Optional sheet name */
  sheet?: string;
  /** Optional end sheet name for 3D references */
  endSheet?: string;
}

export interface StructuredRefNode {
  type: NodeType.StructuredRef;
  /** Table name (empty string for implicit [@Col] syntax) */
  tableName: string;
  /** Column references (e.g. ["Column1", "Column2"] for a column range) */
  columns: string[];
  /** Special items (e.g. ["#Headers", "#Data", "#Totals", "#All", "#This Row"]) */
  specials: string[];
}

/** Represents an omitted/missing argument in a function call (e.g., the middle arg in IF(A1,,0)). */
export interface MissingNode {
  type: NodeType.Missing;
}

export type AstNode =
  | NumberNode
  | StringNode
  | BooleanNode
  | ErrorNode
  | CellRefNode
  | RangeRefNode
  | BinaryOpNode
  | UnaryOpNode
  | FunctionCallNode
  | ArrayNode
  | PercentNode
  | NameNode
  | ColRangeRefNode
  | RowRangeRefNode
  | StructuredRefNode
  | MissingNode;

// ============================================================================
// Operator Precedence
// ============================================================================

function prefixBindingPower(op: string): number {
  switch (op) {
    case "+":
    case "-":
      // Must be lower than ^ (60/61) so that -2^3 parses as -(2^3), not (-2)^3.
      // Excel: -2^2 = -4, not 4.
      return 55;
    default:
      return 0;
  }
}

function infixBindingPower(op: string): [number, number] {
  // Returns [left bp, right bp]. Higher = tighter binding.
  // Left < Right for left-associative, Left > Right for right-associative.
  switch (op) {
    case "=":
    case "<>":
    case "<":
    case ">":
    case "<=":
    case ">=":
      return [10, 11];
    case "&":
      return [20, 21];
    case "+":
    case "-":
      return [30, 31];
    case "*":
    case "/":
      return [40, 41];
    case "^":
      return [61, 60]; // right-associative
    default:
      return [0, 0];
  }
}

// ============================================================================
// Parser
// ============================================================================

class Parser {
  private tokens: Token[];
  private pos: number;
  /** Pending sheet name from a SheetRef token */
  private pendingSheet: string | undefined;
  /** Pending end sheet name for 3D references */
  private pendingEndSheet: string | undefined;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
    this.pos = 0;
  }

  peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  next(): Token {
    return this.tokens[this.pos++];
  }

  expect(type: TokenType): Token {
    const t = this.next();
    if (!t || t.type !== type) {
      throw new Error(`Expected token type ${type}, got ${t?.type ?? "EOF"} (${t?.value})`);
    }
    return t;
  }

  // Entry point
  parse(): AstNode {
    const node = this.parseExpr(0);
    return node;
  }

  parseExpr(minBp: number): AstNode {
    let left = this.parsePrefix();

    while (true) {
      const t = this.peek();
      if (!t) {
        break;
      }

      // Postfix: %
      if (t.type === TokenType.Percent) {
        this.next();
        left = { type: NodeType.Percent, operand: left };
        continue;
      }

      // Infix operators
      if (t.type === TokenType.Operator) {
        const [lbp, rbp] = infixBindingPower(t.value);
        if (lbp === 0 || lbp < minBp) {
          break;
        }
        this.next();
        const right = this.parseExpr(rbp);
        left = { type: NodeType.BinaryOp, op: t.value, left, right };
        continue;
      }

      break;
    }

    return left;
  }

  parsePrefix(): AstNode {
    const t = this.peek();
    if (!t) {
      throw new Error("Unexpected end of formula");
    }

    // SheetRef: consume and set pending sheet for next cell/range
    if (t.type === TokenType.SheetRef) {
      this.next();
      this.pendingSheet = t.sheetName;
      this.pendingEndSheet = t.endSheetName;
      return this.parsePrefix();
    }

    // Unary prefix: + -
    if (t.type === TokenType.UnaryPrefix) {
      this.next();
      const bp = prefixBindingPower(t.value);
      const operand = this.parseExpr(bp);
      return { type: NodeType.UnaryOp, op: t.value, operand };
    }

    // @ implicit intersection prefix (Excel 365)
    if (t.type === TokenType.AtSign) {
      this.next();
      const operand = this.parseExpr(55); // same precedence as unary +/-
      return { type: NodeType.UnaryOp, op: "@", operand };
    }

    // Number literal
    if (t.type === TokenType.Number) {
      this.next();
      return { type: NodeType.Number, value: parseFloat(t.value) };
    }

    // String literal
    if (t.type === TokenType.String) {
      this.next();
      return { type: NodeType.String, value: t.value };
    }

    // Boolean literal
    if (t.type === TokenType.Boolean) {
      this.next();
      return { type: NodeType.Boolean, value: t.value === "TRUE" };
    }

    // Error literal
    if (t.type === TokenType.Error) {
      this.next();
      return { type: NodeType.Error, value: t.value };
    }

    // Parenthesized expression
    if (t.type === TokenType.OpenParen) {
      this.next();
      const expr = this.parseExpr(0);
      this.expect(TokenType.CloseParen);
      return expr;
    }

    // Array constant: {1,2;3,4}
    if (t.type === TokenType.OpenBrace) {
      return this.parseArrayConstant();
    }

    // Function call: NAME(args)
    if (t.type === TokenType.Function) {
      return this.parseFunctionCall();
    }

    // Cell reference (may become range via tokenizer)
    if (t.type === TokenType.CellRef) {
      this.next();
      const ref = this.parseCellRefFromToken(t);
      const sheet = this.pendingSheet;
      const endSheet = this.pendingEndSheet;
      this.pendingSheet = undefined;
      this.pendingEndSheet = undefined;
      ref.sheet = sheet;
      ref.endSheet = endSheet;
      return ref;
    }

    // Range reference (already parsed by tokenizer as A1:B2)
    if (t.type === TokenType.Range) {
      this.next();
      const sheet = this.pendingSheet;
      const endSheet = this.pendingEndSheet;
      this.pendingSheet = undefined;
      this.pendingEndSheet = undefined;
      return this.parseRangeFromToken(t, sheet, endSheet);
    }

    // Named range / defined name
    if (t.type === TokenType.Name) {
      this.next();
      return { type: NodeType.Name, name: t.value };
    }

    // Whole-column range (e.g. A:B, $C:$D)
    if (t.type === TokenType.ColRange) {
      this.next();
      const sheet = this.pendingSheet;
      const endSheet = this.pendingEndSheet;
      this.pendingSheet = undefined;
      this.pendingEndSheet = undefined;
      const parts = t.value.split(":");
      const startCol = parts[0].replace(/\$/g, "").toUpperCase();
      const endCol = parts[1].replace(/\$/g, "").toUpperCase();
      return { type: NodeType.ColRangeRef, startCol, endCol, sheet, endSheet };
    }

    // Whole-row range (e.g. 1:5, $3:$7)
    if (t.type === TokenType.RowRange) {
      this.next();
      const sheet = this.pendingSheet;
      const endSheet = this.pendingEndSheet;
      this.pendingSheet = undefined;
      this.pendingEndSheet = undefined;
      const parts = t.value.split(":");
      const startRow = parseInt(parts[0].replace(/\$/g, ""), 10);
      const endRow = parseInt(parts[1].replace(/\$/g, ""), 10);
      return { type: NodeType.RowRangeRef, startRow, endRow, sheet, endSheet };
    }

    // Structured reference (e.g. Table1[Column], [@Column])
    if (t.type === TokenType.StructuredRef) {
      this.next();
      return {
        type: NodeType.StructuredRef,
        tableName: t.tableName ?? "",
        columns: t.columns ?? [],
        specials: t.specials ?? []
      };
    }

    throw new Error(`Unexpected token: ${t.type} "${t.value}"`);
  }

  private parseCellRefFromToken(t: Token): CellRefNode {
    return {
      type: NodeType.CellRef,
      col: t.col!,
      row: t.row!,
      colAbsolute: t.colAbsolute ?? false,
      rowAbsolute: t.rowAbsolute ?? false
    };
  }

  private parseRangeFromToken(t: Token, sheet?: string, endSheet?: string): RangeRefNode {
    // Token value is like "A1:B2" or "$A$1:$B$2"
    const parts = t.value.split(":");
    const startRef = parseCellRefStr(parts[0]);
    const endRef = parseCellRefStr(parts[1]);

    return {
      type: NodeType.RangeRef,
      start: { ...startRef, sheet },
      end: { ...endRef, sheet },
      sheet,
      endSheet
    };
  }

  private parseFunctionCall(): FunctionCallNode {
    const nameToken = this.next(); // function name
    this.expect(TokenType.OpenParen);

    const args: AstNode[] = [];
    if (this.peek()?.type !== TokenType.CloseParen) {
      // First argument: might be missing if comma is next
      if (this.peek()?.type === TokenType.Comma) {
        args.push({ type: NodeType.Missing });
      } else {
        args.push(this.parseExpr(0));
      }
      while (this.peek()?.type === TokenType.Comma) {
        this.next(); // skip comma
        // Next argument: missing if followed by comma or close paren
        if (this.peek()?.type === TokenType.Comma || this.peek()?.type === TokenType.CloseParen) {
          args.push({ type: NodeType.Missing });
        } else {
          args.push(this.parseExpr(0));
        }
      }
    }
    this.expect(TokenType.CloseParen);

    return {
      type: NodeType.FunctionCall,
      name: nameToken.value,
      args
    };
  }

  private parseArrayConstant(): ArrayNode {
    this.next(); // skip {
    const rows: AstNode[][] = [];
    let currentRow: AstNode[] = [];

    if (this.peek()?.type !== TokenType.CloseBrace) {
      currentRow.push(this.parseExpr(0));

      while (this.peek()) {
        const t = this.peek()!;
        if (t.type === TokenType.Comma) {
          this.next();
          currentRow.push(this.parseExpr(0));
        } else if (t.type === TokenType.Semicolon) {
          this.next();
          rows.push(currentRow);
          currentRow = [];
          currentRow.push(this.parseExpr(0));
        } else {
          break;
        }
      }
    }
    rows.push(currentRow);
    this.expect(TokenType.CloseBrace);

    return { type: NodeType.Array, rows };
  }
}

// Helper to parse a cell reference string like "$A$1" or "B2"
function parseCellRefStr(s: string): CellRefNode {
  let i = 0;
  let colAbsolute = false;
  let rowAbsolute = false;

  if (s[i] === "$") {
    colAbsolute = true;
    i++;
  }
  const colStart = i;
  while (i < s.length && ((s[i] >= "A" && s[i] <= "Z") || (s[i] >= "a" && s[i] <= "z"))) {
    i++;
  }
  const col = s.slice(colStart, i).toUpperCase();

  if (i < s.length && s[i] === "$") {
    rowAbsolute = true;
    i++;
  }
  const row = s.slice(i);

  return {
    type: NodeType.CellRef,
    col,
    row,
    colAbsolute,
    rowAbsolute
  };
}

// ============================================================================
// Public API
// ============================================================================

export function parse(tokens: Token[]): AstNode {
  const parser = new Parser(tokens);
  return parser.parse();
}
