/**
 * Basic Excel formula evaluator.
 *
 * Evaluates formula cells in a worksheet and writes computed results back to
 * each cell, so that PDF export (and any other consumer) sees up-to-date
 * values even when the workbook was never opened in Excel after the last
 * programmatic cell edit.
 *
 * Supported:
 *  - Arithmetic operators  + - * / ^ &
 *  - Comparison operators  = <> < > <= >=
 *  - Range references      A1:B3
 *  - INDIRECT / ROW / COLUMN
 *  - SUM AVERAGE COUNT COUNTA MIN MAX ABS ROUND ROUNDUP ROUNDDOWN
 *    INT CEILING FLOOR MOD POWER SQRT
 *  - IF IFERROR IFNA AND OR NOT
 *  - LEN LEFT RIGHT MID UPPER LOWER TRIM CONCATENATE CONCAT TEXT VALUE
 *  - ISBLANK ISNUMBER ISTEXT ISERROR
 *
 * Unknown functions are silently skipped (result stays null).
 */

import { ValueType } from "@excel/enums";
import type { Worksheet } from "@excel/worksheet";

import { colCache } from "./col-cache";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type FVal = number | string | boolean | null;

/** Result of evaluating one expression node. */
type EvalResult = { v: FVal } | { addr: string } | { err: string };

/** A resolved range argument passed to a function. */
type RangeArg = { from: string; to: string };

/** Anything a function argument can be. */
type FuncArg = EvalResult | RangeArg;

function isRef(r: FuncArg): r is { addr: string } {
  return "addr" in r && !("from" in r);
}
function isRange(r: FuncArg): r is RangeArg {
  return "from" in r && "to" in r;
}
function isErr(r: FuncArg): r is { err: string } {
  return "err" in r;
}

// ---------------------------------------------------------------------------
// Evaluation context
// ---------------------------------------------------------------------------

interface EvalCtx {
  getCellValue(addr: string): FVal;
  getRangeValues(from: string, to: string): FVal[];
  currentRow: number;
  currentCol: number;
}

function resultToValue(r: FuncArg, ctx: EvalCtx): FVal {
  if ("v" in r) return r.v;
  if ("addr" in r && !("from" in r)) return ctx.getCellValue(r.addr);
  if ("err" in r) return null;
  return null;
}

// ---------------------------------------------------------------------------
// Tokenizer
// ---------------------------------------------------------------------------

const enum TK {
  Num,
  Str,
  Name,
  LP,
  RP,
  Comma,
  Colon,
  Op,
  EOF
}

interface Token {
  type: TK;
  value?: string | number;
}

class Tokenizer {
  private i = 0;
  constructor(private readonly s: string) {}

  next(): Token {
    const s = this.s;
    // Skip spaces
    while (this.i < s.length && s[this.i] === " ") this.i++;
    if (this.i >= s.length) return { type: TK.EOF };

    const c = s[this.i];

    // String literal  "..."
    if (c === '"') {
      let str = "";
      this.i++;
      while (this.i < s.length && s[this.i] !== '"') str += s[this.i++];
      this.i++; // closing "
      return { type: TK.Str, value: str };
    }

    // Numeric literal
    if ((c >= "0" && c <= "9") || c === ".") {
      const start = this.i;
      while (this.i < s.length) {
        const ch = s[this.i];
        if (ch >= "0" && ch <= "9") {
          this.i++;
        } else if (ch === ".") {
          this.i++;
        } else if ((ch === "E" || ch === "e") && this.i > start) {
          this.i++;
          if (this.i < s.length && (s[this.i] === "+" || s[this.i] === "-")) this.i++;
        } else {
          break;
        }
      }
      return { type: TK.Num, value: parseFloat(s.slice(start, this.i)) };
    }

    // Punctuation
    if (c === "(") {
      this.i++;
      return { type: TK.LP };
    }
    if (c === ")") {
      this.i++;
      return { type: TK.RP };
    }
    if (c === ",") {
      this.i++;
      return { type: TK.Comma };
    }
    if (c === ":") {
      this.i++;
      return { type: TK.Colon };
    }

    // Multi-char operators
    if (c === "<") {
      if (s[this.i + 1] === ">") {
        this.i += 2;
        return { type: TK.Op, value: "<>" };
      }
      if (s[this.i + 1] === "=") {
        this.i += 2;
        return { type: TK.Op, value: "<=" };
      }
      this.i++;
      return { type: TK.Op, value: "<" };
    }
    if (c === ">") {
      if (s[this.i + 1] === "=") {
        this.i += 2;
        return { type: TK.Op, value: ">=" };
      }
      this.i++;
      return { type: TK.Op, value: ">" };
    }

    // Single-char operators
    if (c === "=" || c === "+" || c === "-" || c === "*" || c === "/" || c === "&" || c === "^") {
      this.i++;
      return { type: TK.Op, value: c };
    }

    // Name (function name or cell reference)
    if (
      (c >= "A" && c <= "Z") ||
      (c >= "a" && c <= "z") ||
      c === "_" ||
      c === "$"
    ) {
      let name = "";
      while (this.i < s.length) {
        const ch = s[this.i];
        if (
          (ch >= "A" && ch <= "Z") ||
          (ch >= "a" && ch <= "z") ||
          (ch >= "0" && ch <= "9") ||
          ch === "_" ||
          ch === "$" ||
          ch === "."
        ) {
          name += ch;
          this.i++;
        } else {
          break;
        }
      }
      return { type: TK.Name, value: name.toUpperCase() };
    }

    // Skip unknown character and try again
    this.i++;
    return this.next();
  }
}

// ---------------------------------------------------------------------------
// Recursive-descent evaluator
// ---------------------------------------------------------------------------

const OP_PREC: Record<string, number> = {
  "=": 1,
  "<>": 1,
  "<": 1,
  ">": 1,
  "<=": 1,
  ">=": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5
};

/** Strip $ signs and upper-case an address like $B$11 → B11 */
function normalizeAddr(addr: string): string {
  return addr.replace(/\$/g, "").toUpperCase();
}

const CELL_REF_RE = /^[A-Z]{1,3}[0-9]{1,7}$/;

class FormulaEvaluator {
  private cur: Token;
  constructor(
    private readonly tok: Tokenizer,
    private readonly ctx: EvalCtx
  ) {
    this.cur = this.tok.next();
  }

  private advance(): Token {
    const prev = this.cur;
    this.cur = this.tok.next();
    return prev;
  }

  /** Check current token type without TypeScript narrowing side-effects. */
  private peekIs(type: TK): boolean {
    return this.cur.type === type;
  }

  /** Evaluate the entire formula, returning a scalar. */
  evaluate(): FVal {
    const result = this.parseExpr();
    return resultToValue(result, this.ctx);
  }

  // -------------------------------------------------------------------------
  // parseExpr — full expression with binary operators
  // -------------------------------------------------------------------------

  private parseExpr(): EvalResult {
    return this.parseBinOp(this.parsePrimary(), 0);
  }

  private parseBinOp(left: EvalResult, minPrec: number): EvalResult {
    while (true) {
      if (this.cur.type !== TK.Op) break;
      const op = this.cur.value as string;
      const prec = OP_PREC[op] ?? -1;
      if (prec < minPrec) break;
      this.advance();

      let right = this.parsePrimary();
      // Right-associative look-ahead for higher-precedence operators
      while (this.cur.type === TK.Op && (OP_PREC[this.cur.value as string] ?? -1) > prec) {
        right = this.parseBinOp(right, OP_PREC[this.cur.value as string]);
      }
      left = this.applyBinOp(op, left, right);
    }
    return left;
  }

  private applyBinOp(op: string, l: EvalResult, r: EvalResult): EvalResult {
    const lv = resultToValue(l, this.ctx);
    const rv = resultToValue(r, this.ctx);
    switch (op) {
      case "+":
        return { v: (lv as number) + (rv as number) };
      case "-":
        return { v: (lv as number) - (rv as number) };
      case "*":
        return { v: (lv as number) * (rv as number) };
      case "/":
        return (rv as number) === 0
          ? { err: "#DIV/0!" }
          : { v: (lv as number) / (rv as number) };
      case "^":
        return { v: Math.pow(lv as number, rv as number) };
      case "&":
        return { v: String(lv ?? "") + String(rv ?? "") };
      case "=":
        return { v: lv === rv };
      case "<>":
        return { v: lv !== rv };
      case "<":
        return { v: (lv as number) < (rv as number) };
      case ">":
        return { v: (lv as number) > (rv as number) };
      case "<=":
        return { v: (lv as number) <= (rv as number) };
      case ">=":
        return { v: (lv as number) >= (rv as number) };
      default:
        return { v: null };
    }
  }

  // -------------------------------------------------------------------------
  // parsePrimary
  // -------------------------------------------------------------------------

  private parsePrimary(): EvalResult {
    const tok = this.cur;

    // Unary minus / plus
    if (tok.type === TK.Op && tok.value === "-") {
      this.advance();
      const operand = this.parsePrimary();
      return { v: -(resultToValue(operand, this.ctx) as number) };
    }
    if (tok.type === TK.Op && tok.value === "+") {
      this.advance();
      return this.parsePrimary();
    }

    // Number literal
    if (tok.type === TK.Num) {
      this.advance();
      return { v: tok.value as number };
    }

    // String literal
    if (tok.type === TK.Str) {
      this.advance();
      return { v: tok.value as string };
    }

    // Parenthesised expression
    if (tok.type === TK.LP) {
      this.advance();
      const expr = this.parseExpr();
      if (this.cur.type === TK.RP) this.advance();
      return expr;
    }

    // Name — function call OR cell reference OR boolean literal
    if (tok.type === TK.Name) {
      this.advance();
      const name = tok.value as string;

      // Function call
      if (this.cur.type === TK.LP) {
        this.advance(); // consume (
        const args: FuncArg[] = [];
        while (!this.peekIs(TK.RP) && !this.peekIs(TK.EOF)) {
          args.push(this.parseFuncArg());
          if (this.peekIs(TK.Comma)) this.advance();
        }
        if (this.peekIs(TK.RP)) this.advance();
        return this.callFn(name, args);
      }

      // Boolean literals
      if (name === "TRUE") return { v: true };
      if (name === "FALSE") return { v: false };

      // Cell reference (e.g. B11, $B$11)
      const clean = normalizeAddr(name);
      if (CELL_REF_RE.test(clean)) {
        return { addr: clean };
      }

      // Named range or unknown identifier — treat as null
      return { v: null };
    }

    return { v: null };
  }

  // -------------------------------------------------------------------------
  // parseFuncArg — handles range syntax "A1:B3" or "A1:INDIRECT(...)"
  // -------------------------------------------------------------------------

  private parseFuncArg(): FuncArg {
    const lhs = this.parsePrimary();

    // Range colon?
    if (this.cur.type === TK.Colon) {
      this.advance();
      const rhs = this.parsePrimary();

      const fromAddr = isRef(lhs) ? lhs.addr : null;
      // RHS can be a ref directly, or a string returned by INDIRECT
      const toAddr = isRef(rhs)
        ? rhs.addr
        : typeof (rhs as { v?: FVal }).v === "string"
          ? ((rhs as { v: string }).v as string).toUpperCase()
          : null;

      if (fromAddr && toAddr) {
        return { from: fromAddr, to: toAddr };
      }
      return { v: null };
    }

    // Not a range — parse the rest as a binary expression
    return this.parseBinOp(lhs, 0);
  }

  // -------------------------------------------------------------------------
  // Function implementations
  // -------------------------------------------------------------------------

  private callFn(name: string, args: FuncArg[]): EvalResult {
    /** Collect all numeric values from args (ranges expanded). */
    const nums = (): number[] => {
      const out: number[] = [];
      for (const a of args) {
        if (isRange(a)) {
          for (const v of this.ctx.getRangeValues(a.from, a.to)) {
            if (typeof v === "number") out.push(v);
          }
        } else {
          const v = resultToValue(a as EvalResult, this.ctx);
          if (typeof v === "number") out.push(v);
        }
      }
      return out;
    };

    /** Collect all values from args (ranges expanded). */
    const allVals = (): FVal[] => {
      const out: FVal[] = [];
      for (const a of args) {
        if (isRange(a)) {
          out.push(...this.ctx.getRangeValues(a.from, a.to));
        } else {
          out.push(resultToValue(a as EvalResult, this.ctx));
        }
      }
      return out;
    };

    const v0 = (): FVal => resultToValue(args[0] as EvalResult, this.ctx);
    const v1 = (): FVal => resultToValue(args[1] as EvalResult, this.ctx);
    const v2 = (): FVal => resultToValue(args[2] as EvalResult, this.ctx);

    switch (name) {
      // ---- Aggregates ----
      case "SUM": {
        const n = nums();
        return { v: n.reduce((a, b) => a + b, 0) };
      }
      case "AVERAGE": {
        const n = nums();
        return { v: n.length ? n.reduce((a, b) => a + b, 0) / n.length : 0 };
      }
      case "COUNT":
        return { v: nums().length };
      case "COUNTA":
        return { v: allVals().filter(v => v !== null && v !== undefined && v !== "").length };
      case "COUNTBLANK":
        return { v: allVals().filter(v => v === null || v === undefined || v === "").length };
      case "MIN": {
        const n = nums();
        return { v: n.length ? Math.min(...n) : 0 };
      }
      case "MAX": {
        const n = nums();
        return { v: n.length ? Math.max(...n) : 0 };
      }

      // ---- Math ----
      case "ABS":
        return { v: Math.abs(v0() as number) };
      case "ROUND": {
        const factor = Math.pow(10, v1() as number);
        return { v: Math.round((v0() as number) * factor) / factor };
      }
      case "ROUNDUP": {
        const factor = Math.pow(10, v1() as number);
        const val = v0() as number;
        return { v: (Math.ceil(Math.abs(val) * factor) / factor) * Math.sign(val) };
      }
      case "ROUNDDOWN": {
        const factor = Math.pow(10, v1() as number);
        const val = v0() as number;
        return { v: (Math.floor(Math.abs(val) * factor) / factor) * Math.sign(val) };
      }
      case "INT":
        return { v: Math.floor(v0() as number) };
      case "CEILING":
      case "CEILING.MATH": {
        const sig = args[1] ? (v1() as number) : 1;
        return { v: Math.ceil((v0() as number) / sig) * sig };
      }
      case "FLOOR":
      case "FLOOR.MATH": {
        const sig = args[1] ? (v1() as number) : 1;
        return { v: Math.floor((v0() as number) / sig) * sig };
      }
      case "MOD":
        return { v: (v0() as number) % (v1() as number) };
      case "POWER":
        return { v: Math.pow(v0() as number, v1() as number) };
      case "SQRT":
        return { v: Math.sqrt(v0() as number) };
      case "LOG": {
        const base = args[1] ? (v1() as number) : 10;
        return { v: Math.log(v0() as number) / Math.log(base) };
      }
      case "LN":
        return { v: Math.log(v0() as number) };
      case "EXP":
        return { v: Math.exp(v0() as number) };

      // ---- Logic ----
      case "IF": {
        const cond = v0();
        const yes = args[1] ? v1() : true;
        const no = args[2] ? v2() : false;
        return { v: cond ? yes : no };
      }
      case "IFERROR": {
        const a = args[0] as EvalResult;
        return isErr(a) ? { v: args[1] ? v1() : "" } : { v: resultToValue(a, this.ctx) };
      }
      case "IFNA": {
        const a = args[0] as EvalResult;
        return isErr(a) && (a as { err: string }).err === "#N/A"
          ? { v: args[1] ? v1() : "" }
          : { v: resultToValue(a as EvalResult, this.ctx) };
      }
      case "AND":
        return { v: allVals().every(v => !!v) };
      case "OR":
        return { v: allVals().some(v => !!v) };
      case "NOT":
        return { v: !v0() };

      // ---- Text ----
      case "LEN":
        return { v: String(v0() ?? "").length };
      case "LEFT": {
        const n = args[1] ? (v1() as number) : 1;
        return { v: String(v0() ?? "").slice(0, n) };
      }
      case "RIGHT": {
        const n = args[1] ? (v1() as number) : 1;
        const s = String(v0() ?? "");
        return { v: s.slice(Math.max(0, s.length - n)) };
      }
      case "MID": {
        const s = String(v0() ?? "");
        const start = (v1() as number) - 1;
        const count = v2() as number;
        return { v: s.slice(start, start + count) };
      }
      case "UPPER":
        return { v: String(v0() ?? "").toUpperCase() };
      case "LOWER":
        return { v: String(v0() ?? "").toLowerCase() };
      case "TRIM":
        return { v: String(v0() ?? "").trim() };
      case "SUBSTITUTE": {
        const src = String(v0() ?? "");
        const old = String(v1() ?? "");
        const rep = String(v2() ?? "");
        return { v: src.split(old).join(rep) };
      }
      case "REPT": {
        const s = String(v0() ?? "");
        const n = v1() as number;
        return { v: s.repeat(Math.max(0, Math.floor(n))) };
      }
      case "CONCATENATE":
      case "CONCAT":
        return { v: allVals().map(v => String(v ?? "")).join("") };
      case "TEXT":
        return { v: String(v0() ?? "") };
      case "VALUE":
        return { v: parseFloat(String(v0() ?? "0")) || 0 };
      case "EXACT":
        return { v: String(v0()) === String(v1()) };
      case "FIND":
      case "SEARCH": {
        const needle = String(v0() ?? "");
        const haystack = String(v1() ?? "");
        const idx =
          name === "SEARCH"
            ? haystack.toLowerCase().indexOf(needle.toLowerCase())
            : haystack.indexOf(needle);
        return { v: idx >= 0 ? idx + 1 : 0 };
      }

      // ---- Reference & info ----
      case "ROW": {
        if (args.length === 0) return { v: this.ctx.currentRow };
        const a = args[0];
        if (isRef(a)) {
          const dec = colCache.decode(a.addr) as any;
          return { v: dec?.row ?? this.ctx.currentRow };
        }
        return { v: this.ctx.currentRow };
      }
      case "COLUMN": {
        if (args.length === 0) return { v: this.ctx.currentCol };
        const a = args[0];
        if (isRef(a)) {
          const dec = colCache.decode(a.addr) as any;
          return { v: dec?.col ?? this.ctx.currentCol };
        }
        return { v: this.ctx.currentCol };
      }
      case "ROWS": {
        if (isRange(args[0])) {
          const from = colCache.decode(args[0].from) as any;
          const to = colCache.decode(args[0].to) as any;
          return { v: Math.abs(to.row - from.row) + 1 };
        }
        return { v: 1 };
      }
      case "COLUMNS": {
        if (isRange(args[0])) {
          const from = colCache.decode(args[0].from) as any;
          const to = colCache.decode(args[0].to) as any;
          return { v: Math.abs(to.col - from.col) + 1 };
        }
        return { v: 1 };
      }
      case "INDIRECT": {
        const ref = resultToValue(args[0] as EvalResult, this.ctx);
        if (typeof ref === "string") {
          const clean = normalizeAddr(ref);
          // Could be a cell address or a range address
          if (CELL_REF_RE.test(clean)) {
            return { addr: clean };
          }
        }
        return { err: "#REF!" };
      }
      case "ADDRESS": {
        const row = v0() as number;
        const col = v1() as number;
        return { v: colCache.encodeAddress(row, col) };
      }
      case "OFFSET": {
        // OFFSET(ref, rows, cols[, height, width])
        const baseRef = args[0];
        if (isRef(baseRef)) {
          const base = colCache.decode(baseRef.addr) as any;
          const rowOffset = (v1() as number) || 0;
          const colOffset = v2() as number;
          const newAddr = colCache.encodeAddress(
            base.row + rowOffset,
            base.col + (colOffset || 0)
          );
          return { addr: newAddr };
        }
        return { err: "#REF!" };
      }

      // ---- Type checks ----
      case "ISBLANK":
        return { v: v0() === null || v0() === undefined || v0() === "" };
      case "ISNUMBER":
        return { v: typeof v0() === "number" };
      case "ISTEXT":
        return { v: typeof v0() === "string" };
      case "ISERROR":
      case "ISERR":
        return { v: isErr(args[0] as EvalResult) };
      case "ISNA":
        return { v: isErr(args[0] as EvalResult) && (args[0] as { err: string }).err === "#N/A" };

      default:
        // Unknown function — leave result as null
        return { v: null };
    }
  }
}

// ---------------------------------------------------------------------------
// Worksheet helpers
// ---------------------------------------------------------------------------

/** Resolve the "effective" scalar value of a cell (follows formula results). */
function getCellValueAt(row: number, col: number, ws: Worksheet): FVal {
  const cell = ws.findCell(row, col) as any;
  if (!cell) return null;

  const type = cell.type as number;
  if (type === ValueType.Null || type === ValueType.Merge) return null;

  if (type === ValueType.Formula) {
    const result = cell.result;
    if (result === undefined || result === null) return null;
    if (result instanceof Date) return null; // not a number for arithmetic
    return result as FVal;
  }

  const v = cell.value;
  if (v === null || v === undefined) return null;
  if (typeof v === "number") return v;
  if (typeof v === "string") return v;
  if (typeof v === "boolean") return v;
  if (v instanceof Date) return null;
  // FormulaValueData object without result
  if (typeof v === "object" && "result" in v) return (v as any).result ?? null;
  return null;
}

function getCellValueByAddr(addr: string, ws: Worksheet): FVal {
  try {
    const decoded = colCache.decode(addr) as any;
    return getCellValueAt(decoded.row, decoded.col, ws);
  } catch {
    return null;
  }
}

function getRangeValues(fromAddr: string, toAddr: string, ws: Worksheet): FVal[] {
  try {
    const from = colCache.decode(fromAddr) as any;
    const to = colCache.decode(toAddr) as any;
    const minRow = Math.min(from.row, to.row);
    const maxRow = Math.max(from.row, to.row);
    const minCol = Math.min(from.col, to.col);
    const maxCol = Math.max(from.col, to.col);
    const vals: FVal[] = [];
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        vals.push(getCellValueAt(r, c, ws));
      }
    }
    return vals;
  } catch {
    return [];
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Evaluate all formula cells in `ws` and write their computed result back,
 * so that PDF export (and other consumers) see current values.
 *
 * Safe to call multiple times; silently skips cells whose formulas cannot
 * be evaluated (e.g. they use unsupported functions like VLOOKUP).
 */
export function evaluateWorksheetFormulas(ws: Worksheet): void {
  const dims = ws.dimensions as any;
  if (!dims || !dims.model || dims.model.top === 0) return;

  const { top, bottom } = dims.model;

  for (let r = top; r <= bottom; r++) {
    const row = ws.findRow(r);
    if (!row) continue;

    (row as any).eachCell({ includeEmpty: false }, (cell: any) => {
      if (cell.type !== ValueType.Formula) return;

      // Prefer cell.formula (handles shared formulas via translation)
      const formula: string | undefined = cell.formula || (cell.value as any)?.formula;
      if (!formula) return;

      const ctx: EvalCtx = {
        getCellValue: (addr: string) => getCellValueByAddr(addr, ws),
        getRangeValues: (from: string, to: string) => getRangeValues(from, to, ws),
        currentRow: cell.row as number,
        currentCol: cell.col as number
      };

      try {
        const evaluator = new FormulaEvaluator(new Tokenizer(formula), ctx);
        const result = evaluator.evaluate();

        // Write back: preserve formula & all metadata, only update result
        const currentVal = cell.value;
        const updatedVal =
          currentVal !== null &&
          typeof currentVal === "object" &&
          ("formula" in currentVal || "sharedFormula" in currentVal)
            ? { ...(currentVal as object), result }
            : { formula, result };

        cell.value = updatedVal;
      } catch {
        // Silently skip unevaluable formulas
      }
    });
  }
}
