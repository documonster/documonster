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

import { formulaArgument } from "@formula/arguments";
import { FormulaParseError } from "@formula/errors";
import type {
  AstNode,
  CellRefNode,
  RangeRefNode,
  FunctionCallNode,
  ArrayNode
} from "@formula/syntax/ast";
import { NodeType } from "@formula/syntax/ast";
import { layoutOf, type FormulaLayout } from "@formula/syntax/layout";
import type { Token, CellRefToken, RangeToken, FunctionToken } from "@formula/syntax/token-types";
import { isSpillOperator, SPILL_OPERATOR, TokenType } from "@formula/syntax/token-types";
import { tokenize, tokenizeWithSpans } from "@formula/syntax/tokenizer";
import { check, s } from "@utils/schema";

// ============================================================================
// Operator Precedence
// ============================================================================

function prefixBindingPower(op: string): number {
  switch (op) {
    case "+":
    case "-":
      // Excel's unary `-` binds TIGHTER than `^` — unique among
      // spreadsheets and most programming languages. Microsoft's
      // published precedence table places "Negation (as in –1)" at
      // rank 1 (highest) and "Exponentiation" at rank 4.
      //   =-2^2  →  (-2)^2  →  4    (NOT -(2^2) = -4)
      //   =-2^3  →  (-2)^3  →  -8
      // Previously we used 55 which routed via `^` (60/61) and produced
      // `-(2^3)` — matching most languages but not Excel.
      //
      // Above `%` (75) too, which the same table ranks third: `-A1%` is `(-A1)%`. Excel's own token stream for it
      // is `A1, PtgUminus, PtgPercent`, and at 70 this parsed it as `-(A1%)` — the same number, a different tree,
      // and one that printed back as `(-A1)%` from Excel's tokens.
      return PREFIX_BINDING_POWER;
    default:
      return 0;
  }
}

/** Unary `+`/`-`: Excel's precedence table ranks negation above `%` and `^`. */
const PREFIX_BINDING_POWER = 77;
/** Postfix `%`. */
const PERCENT_BINDING_POWER = 75;
/** Postfix `#` (spill range): above the range operator `:` (100), so `A1#` is one operand wherever it stands. */
const SPILL_BINDING_POWER = 110;

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
      // Excel is unusual: `^` is LEFT-associative (not right-associative
      // like the math convention). `=2^3^2` evaluates to `(2^3)^2 = 64`,
      // not `2^(3^2) = 512`. Using right-associative precedence silently
      // diverged from Excel for any stacked exponent.
      return [60, 61];
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
  /** Whitespace before each token, when a layout was asked for — see `parseWithLayout`. */
  private readonly gaps: readonly string[] | undefined;
  /** Where the layout is collected; present exactly when `gaps` is. */
  readonly layout: FormulaLayout | undefined;

  constructor(tokens: Token[], gaps?: readonly string[]) {
    this.tokens = tokens;
    this.pos = 0;
    this.depth = 0;
    this.gaps = gaps;
    this.layout = gaps === undefined ? undefined : new Map();
  }

  /** The whitespace before the token at `index`, or `""` when no layout is being kept. */
  private gapAt(index: number): string {
    return this.gaps?.[index] ?? "";
  }

  /** Record a layout field, skipping empty whitespace so a plain formula produces an empty layout. */
  private note(node: AstNode, field: "lead" | "op" | "close", text: string): void {
    if (this.layout !== undefined && text !== "") {
      const entry = layoutOf(this.layout, node);
      entry[field] = text + (field === "lead" ? (entry.lead ?? "") : "");
    }
  }

  private enter(): void {
    if (++this.depth > this.MAX_DEPTH) {
      throw new FormulaParseError(`Formula nested too deep (> ${this.MAX_DEPTH} levels)`, this.pos);
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
      throw new FormulaParseError(
        `Expected ${TokenType[type]}, got ${t ? TokenType[t.type] : "the end of the formula"}`,
        this.pos
      );
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
      throw new FormulaParseError(`Unexpected trailing ${TokenType[t.type]}`, this.pos);
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

      // Postfix `#`, Excel's spill-range operator: `A1#` is the range `A1`'s dynamic array spills into. The
      // tokenizer emits it only straight after a cell reference, and it binds tighter than everything else.
      if (isSpillOperator(t)) {
        if (SPILL_BINDING_POWER < minBp) {
          break;
        }
        this.next();
        left = { type: NodeType.UnaryOp, op: SPILL_OPERATOR, operand: left };
        continue;
      }

      // Postfix: %. Binds tighter than every arithmetic operator and looser than negation and the reference
      // operators, so it ends an operand being parsed at a higher power — `-A1%` is `(-A1)%`, `A1 B1%` is `(A1 B1)%`.
      if (t.type === TokenType.Percent) {
        if (PERCENT_BINDING_POWER < minBp) {
          break;
        }
        const gap = this.gapAt(this.pos);
        this.next();
        left = { type: NodeType.Percent, operand: left };
        this.note(left, "op", gap);
        continue;
      }

      // Infix operators
      if (t.type === TokenType.Operator) {
        const [lbp, rbp] = infixBindingPower(t.value);
        if (lbp === 0 || lbp < minBp) {
          break;
        }
        const gap = this.gapAt(this.pos);
        this.next();
        const right = this.parseExpr(rbp);
        left = { type: NodeType.BinaryOp, op: t.value, left, right };
        this.note(left, "op", gap);
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
        // The operator *is* the whitespace, so only a run longer than one character says anything.
        const gap = this.gapAt(this.pos);
        this.next();
        const right = this.parseExpr(rbp);
        left = { type: NodeType.BinaryOp, op: " ", left, right };
        if (gap !== " ") {
          this.note(left, "op", gap);
        }
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
        const gap = this.gapAt(this.pos);
        this.next();
        const right = this.parseExpr(rbp);
        left = { type: NodeType.BinaryOp, op: ":", left, right };
        this.note(left, "op", gap);
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
        throw new FormulaParseError(
          `Sheet reference '${sheet}' not followed by a cell or range`,
          this.pos
        );
      }
      return result;
    } finally {
      this.leave();
    }
  }

  private parsePrefixInner(): AstNode {
    // Whitespace before the prefix's first token belongs to the node it starts — except for a parenthesis, whose
    // whitespace is the parenthesis's own and is recorded with it.
    const gap = this.gapAt(this.pos);
    if (this.peek()?.type === TokenType.OpenParen) {
      return this.parseParenthesised(gap);
    }
    const node = this.parsePrefixCore();
    this.note(node, "lead", gap);
    return node;
  }

  /**
   * A parenthesised expression — also the syntactic entry point for reference unions: `(A1:B2, D4:E5)` is a
   * multi-area reference that `INDEX(..., area_num)` can index into. A single expression inside parens is just an
   * expression group (no UnionRef wrapper), and when a layout is being kept the group is recorded on the node it
   * encloses.
   */
  private parseParenthesised(openGap: string): AstNode {
    this.next();
    const first = this.parseExpr(0);
    if (this.peek()?.type === TokenType.Comma) {
      const areas: AstNode[] = [first];
      while (this.peek()?.type === TokenType.Comma) {
        this.next(); // consume ','
        areas.push(this.parseExpr(0));
      }
      const closeGap = this.gapAt(this.pos);
      this.expect(TokenType.CloseParen);
      const union: AstNode = { type: NodeType.UnionRef, areas };
      this.note(union, "lead", openGap);
      this.note(union, "close", closeGap);
      return union;
    }
    const closeGap = this.gapAt(this.pos);
    this.expect(TokenType.CloseParen);
    if (this.layout !== undefined) {
      const entry = layoutOf(this.layout, first);
      (entry.parens ??= []).push({
        ...(openGap === "" ? {} : { open: openGap }),
        ...(closeGap === "" ? {} : { close: closeGap })
      });
    }
    return first;
  }

  private parsePrefixCore(): AstNode {
    const t = this.peek();
    if (!t) {
      throw new FormulaParseError("Unexpected end of formula", this.pos);
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

    // Error literal. `Sheet2!#REF!` is what a reference becomes once the cells it named are deleted; the error is
    // the value, and the qualifying sheet is kept so the text survives.
    if (t.type === TokenType.Error) {
      this.next();
      if (t.externalReference !== undefined) {
        return { type: NodeType.Error, value: t.value, externalReference: t.externalReference };
      }
      const sheet = this.takeSheetFor(t.value === "#REF!");
      return { type: NodeType.Error, value: t.value, ...(sheet === undefined ? {} : { sheet }) };
    }

    // Parenthesised expressions are handled by `parseParenthesised`, before this is reached.

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

    // Named range / defined name. `Sheet2!Total` names the one scoped to that sheet.
    if (t.type === TokenType.Name) {
      this.next();
      const sheet = this.takeSheetFor(true);
      return { type: NodeType.Name, name: t.value, ...(sheet === undefined ? {} : { sheet }) };
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
      return {
        type: NodeType.ColRangeRef,
        startCol,
        endCol,
        ...absoluteMarkers(parts),
        sheet,
        endSheet
      };
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
      return {
        type: NodeType.RowRangeRef,
        startRow,
        endRow,
        ...absoluteMarkers(parts),
        sheet,
        endSheet
      };
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

    throw new FormulaParseError(
      `Unexpected ${TokenType[t.type] ?? `token type ${String(t.type)}`}`,
      this.pos
    );
  }

  /**
   * Consume a pending single-sheet qualifier for a node that can carry one. Left pending otherwise, so
   * `parsePrefix` reports the qualifier as dangling.
   */
  private takeSheetFor(allowed: boolean): string | undefined {
    if (!allowed || this.pendingSheet === undefined || this.pendingEndSheet !== undefined) {
      return undefined;
    }
    const sheet = this.pendingSheet;
    this.pendingSheet = undefined;
    return sheet;
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
    const first = this.parseArgumentList();
    const call: FunctionCallNode = {
      type: NodeType.FunctionCall,
      name: nameToken.name,
      args: first.args
    };
    this.note(call, "close", first.closeGap);
    // An argument list straight after the `)` applies the call's result — `LAMBDA(n,n*2)(3)`, `f(1)(2)`. Only an
    // adjacent `(`: with whitespace between, the tokenizer has already made it an intersection.
    while (this.peek()?.type === TokenType.OpenParen) {
      const applied = this.parseArgumentList();
      (call.invocations ??= []).push(applied.args);
      if (this.layout !== undefined && applied.closeGap !== "") {
        const entry = layoutOf(this.layout, call);
        (entry.invocationClose ??= [])[call.invocations.length - 1] = applied.closeGap;
      }
    }
    return call;
  }

  /** `(arg, arg, …)`, with a missing argument wherever a comma or the `)` leaves one empty. */
  private parseArgumentList(): { args: AstNode[]; closeGap: string } {
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
    const closeGap = this.gapAt(this.pos);
    this.expect(TokenType.CloseParen);
    return { args, closeGap };
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
/** The `$` markers on the two halves of a whole-row or whole-column range, only where present. */
function absoluteMarkers(parts: readonly string[]): {
  startAbsolute?: boolean;
  endAbsolute?: boolean;
} {
  return {
    ...(parts[0]?.startsWith("$") ? { startAbsolute: true } : {}),
    ...(parts[1]?.startsWith("$") ? { endAbsolute: true } : {})
  };
}

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

/**
 * Parse, and also collect the formula's layout — the whitespace and redundant parentheses its text carries.
 *
 * `gaps` is what `tokenizeWithGaps` returns beside the tokens. The tree is the one `parse` builds; the layout is
 * keyed by its nodes. See `syntax/layout.ts`.
 */
export function parseWithLayout(
  tokens: Token[],
  gaps: readonly string[]
): { ast: AstNode; layout: FormulaLayout } {
  const parser = new Parser(tokens, gaps);
  const ast = parser.parse();
  return { ast, layout: parser.layout! };
}

/**
 * Parse a formula into an AST, from its source text (`"=SUM(A1:B2)"`, the
 * leading `=` optional) or from tokens `Formula.tokenize` already produced.
 *
 * @example
 * ```ts
 * import { Formula } from "documonster/formula";
 * const ast = Formula.parse("=A1+1");
 *
 * // Tokens you already have (e.g. after inspecting or filtering them) skip re-tokenizing
 * const tokens = Formula.tokenize("SUM(A1:B2)");
 * const same = Formula.parse(tokens);
 * ```
 *
 * The result is the same either way; `Formula.format(ast)` turns it back into text.
 */
export function parseFormula(source: string | readonly Token[]): AstNode {
  if (typeof source === "string") {
    const tokens = tokenize(source);
    try {
      return parse(tokens);
    } catch (error) {
      // The parser counts tokens; a caller holding the text wants to know where in it.
      throw relocate(error, source);
    }
  }
  check(
    s.array(s.unknown(), { expected: "formula text or the tokens Formula.tokenize returned" }),
    source,
    formulaArgument("Formula.parse", "source")
  );
  try {
    return parse(source as Token[]);
  } catch (error) {
    if (error instanceof FormulaParseError && error.position !== undefined) {
      // No text to point into: name the token instead of passing its index off as an offset.
      throw new FormulaParseError(`${error.detail} (at token ${error.position})`, undefined, {
        cause: error
      });
    }
    throw error;
  }
}

/** A parser error re-expressed with the character offset of the token it names. */
function relocate(error: unknown, source: string): unknown {
  if (!(error instanceof FormulaParseError) || error.position === undefined) {
    return error;
  }
  const { spans } = tokenizeWithSpans(source);
  const index = error.position;
  const offset = index < spans.length ? spans[index][0] : source.length;
  return new FormulaParseError(error.detail, offset, { cause: error });
}
