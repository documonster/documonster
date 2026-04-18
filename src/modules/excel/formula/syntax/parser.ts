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

import {
  NodeType,
  type AstNode,
  type CellRefNode,
  type RangeRefNode,
  type FunctionCallNode,
  type ArrayNode
} from "./ast";
import {
  TokenType,
  type Token,
  type CellRefToken,
  type RangeToken,
  type FunctionToken
} from "./token-types";

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
    // Intersection operator — whitespace between two refs. In Excel
    // precedence this sits between `:` (range, already handled at the
    // tokenizer level) and unary +/-. Left-associative, binds tighter
    // than every arithmetic operator.
    case " ":
      return [80, 81];
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
  /**
   * Current recursion depth for parseExpr / parsePrefix. Excel caps
   * formula nesting at 64; we allow a generous 256 to tolerate legitimate
   * LAMBDA bodies, deeply nested IF chains, etc., while still refusing
   * adversarial `((((..(A1)..))))` patterns that would otherwise blow
   * V8's JS call stack.
   */
  private depth: number;
  private readonly MAX_DEPTH = 256;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
    this.pos = 0;
    this.depth = 0;
  }

  private enter(): void {
    if (++this.depth > this.MAX_DEPTH) {
      throw new Error(`Formula nested too deep (> ${this.MAX_DEPTH} levels)`);
    }
  }

  private leave(): void {
    this.depth--;
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
      throw new Error(`Expected token type ${type}, got ${t?.type ?? "EOF"}`);
    }
    return t;
  }

  // Entry point
  parse(): AstNode {
    const node = this.parseExpr(0);
    // After parsing a complete expression, every token must have been
    // consumed. Otherwise inputs like "A1:B2:C3" or "1+2 3" would silently
    // produce an AST that drops the trailing tokens.
    if (this.pos !== this.tokens.length) {
      const t = this.tokens[this.pos];
      throw new Error(`Unexpected trailing token at position ${this.pos} (type ${t.type})`);
    }
    return node;
  }

  parseExpr(minBp: number): AstNode {
    this.enter();
    try {
      return this.parseExprInner(minBp);
    } finally {
      this.leave();
    }
  }

  private parseExprInner(minBp: number): AstNode {
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

      // Intersection: whitespace separating two refs. Modelled as a binary
      // operation with op=" ". The evaluator resolves both sides to refs
      // and returns their rectangle intersection (or `#NULL!`).
      if (t.type === TokenType.Intersect) {
        const [lbp, rbp] = infixBindingPower(" ");
        if (lbp < minBp) {
          break;
        }
        this.next();
        const right = this.parseExpr(rbp);
        left = { type: NodeType.BinaryOp, op: " ", left, right };
        continue;
      }

      // Range operator `:` — binds tighter than intersection. Normally
      // `A1:B2` is merged into a single Range token by the tokenizer,
      // but patterns like `B11:INDIRECT("B" & ROW()-1)` where one side
      // is a function call leave the colon as a standalone token. The
      // evaluator then coerces both sides to references and constructs
      // the union rectangle.
      if (t.type === TokenType.Colon) {
        // Right-biased binding power to guarantee higher precedence than
        // every other operator in the table above.
        const lbp = 100;
        const rbp = 101;
        if (lbp < minBp) {
          break;
        }
        this.next();
        const right = this.parseExpr(rbp);
        left = { type: NodeType.BinaryOp, op: ":", left, right };
        continue;
      }

      break;
    }

    return left;
  }

  parsePrefix(): AstNode {
    this.enter();
    try {
      const result = this.parsePrefixInner();
      // If the SheetRef consumer below set pendingSheet but the prefix that
      // followed wasn't a ref (CellRef / Range / ColRange / RowRange), the
      // sheet qualifier would leak into the next parse and silently attach
      // itself to an unrelated node. Detect it here.
      if (this.pendingSheet !== undefined) {
        const sheet = this.pendingSheet;
        this.pendingSheet = undefined;
        this.pendingEndSheet = undefined;
        throw new Error(`Sheet reference '${sheet}' not followed by a cell or range`);
      }
      return result;
    } finally {
      this.leave();
    }
  }

  private parsePrefixInner(): AstNode {
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
      const n = parseFloat(t.value);
      // Reject non-finite literals at parse time. `1e400` would otherwise
      // become `Infinity` and flow through arithmetic, requiring every
      // downstream consumer to guard. Surfacing the overflow as a #NUM!
      // error node keeps the engine's invariant that numeric values are
      // always finite.
      if (!Number.isFinite(n)) {
        return { type: NodeType.Error, value: "#NUM!" };
      }
      return { type: NodeType.Number, value: n };
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
        tableName: t.tableName,
        columns: t.columns,
        specials: t.specials
      };
    }

    throw new Error(`Unexpected token: ${t.type}`);
  }

  private parseCellRefFromToken(t: CellRefToken): CellRefNode {
    return {
      type: NodeType.CellRef,
      col: t.col,
      row: t.row,
      colAbsolute: t.colAbsolute,
      rowAbsolute: t.rowAbsolute
    };
  }

  private parseRangeFromToken(t: RangeToken, sheet?: string, endSheet?: string): RangeRefNode {
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
    const nameToken = this.next() as FunctionToken; // function name
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
      name: nameToken.name,
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
