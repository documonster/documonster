/**
 * BIFF12 `Ptg` token streams ⇄ formula AST.
 *
 * ## What this file does and does not own
 *
 * BIFF12 stores a formula as a reverse-polish stream of `Ptg` tokens, not as text. Turning that into something
 * usable needs two things: a mapping between tokens and expression structure, and a way to get from structure to
 * text. Only the first is specific to XLSB, and this file is only the first.
 *
 * The second is `@formula`: `tokenizeWithGaps` + `parseWithLayout` for text → AST, and `printAstWithLayout` for
 * AST → text. Sharing that is not opportunistic reuse — the alternative is assembling text while walking the token
 * stack, which means reimplementing operator precedence and parenthesisation. A mistake there does not produce a
 * syntax error; it produces a different formula that parses cleanly and computes something else.
 *
 * ```text
 * read:   Ptg bytes ──► AST + layout ──► text      (decodeFormula, then formulaText)
 * write:  text ──► AST + layout ──► Ptg bytes      (encodeFormula)
 * ```
 *
 * The *layout* is the part of a formula's text the AST does not hold — the author's spaces and line breaks and the
 * parentheses precedence does not need — and Excel stores exactly that as tokens of its own (`PtgAttrSpace`,
 * `PtgParen`). It travels beside the tree (`@formula/syntax/layout`), so `( A1 + B1 ) * 2` written by Excel reads
 * back as Excel shows it and is written back the same way.
 *
 * ## Three properties Excel checks and a round trip cannot
 *
 * Every defect this codec has had passed its own round trip, because the encoder and the decoder agreed with each
 * other. The tests that matter compare against bytes Excel wrote (`__tests__/data/*-excel.*`). The three things
 * most easily got wrong:
 *
 * - **Operand classes.** A reference token's class (`0x20` reference, `0x40` value, `0x60` array) states what its
 *   consumer does with it. `=A1+B1` with reference-class operands is `#VALUE!` in Excel. The class of every token is
 *   computed from the root down — see {@link assignClasses}.
 * - **Control tokens.** Excel writes `IF` and `CHOOSE` with jump tables (`PtgAttrIf`, `PtgAttrChoose`,
 *   `PtgAttrGoto`), a one-argument `SUM` as `PtgAttrSum`, and a volatile formula with a leading `PtgAttrSemi`. The
 *   reader must skip or interpret each, and the writer emits them as Excel does.
 * - **Display tokens.** Without `PtgParen` Excel *shows* `A1*(B1+C1)` as `A1*B1+C1` — it still computes the right
 *   thing, from the token order, and displays something else.
 *
 * ## Unsupported tokens fail loudly
 *
 * A token this codec does not model throws `ExcelNotSupportedError` naming the token. It never returns a partial
 * expression: half a formula is worse than no formula, because it looks like it worked.
 */

import { ExcelNotSupportedError, XlsbParseError } from "@excel/errors";
import { decodeCol, encodeCol } from "@excel/utils/address";
import { definedNameListKey } from "@excel/utils/function-names";
import {
  bareFunctionName,
  isLambdaParameter,
  isWorksheetFunctionName,
  storedFunctionName,
  fromFileFormula,
  toFileFormula
} from "@excel/utils/future-functions";
import { checkSheetIndex, XLSB_MAX_COLUMNS, XLSB_MAX_ROWS, XlsbReader } from "@excel/xlsb/binary";
import { errorCodeOf, errorTextOf } from "@excel/xlsb/error-values";
import {
  ARRAY_CLASS,
  assignClasses,
  bareIntersections,
  nonReferenceArgument,
  builtInId,
  INVOKE,
  invocationOperands,
  isAttrSum,
  isInvocation,
  needsMemToken,
  REFERENCE_CLASS,
  VALUE_CLASS,
  type FormulaRole,
  type TokenClasses
} from "@excel/xlsb/formula/operand-classes";
import {
  NodeType,
  type AstNode,
  type CellRefNode,
  type ColRangeRefNode,
  type FunctionCallNode,
  type RowRangeRefNode,
  type StructuredRefNode
} from "@formula/syntax/ast";
import {
  ARITY_BY_ID,
  FUNCTION_NAME_BY_ID,
  FUNCTION_TABLES,
  functionClasses,
  futureFunctionClasses
} from "@formula/syntax/function-table";
import { layoutOf, type FormulaLayout, type NodeLayout } from "@formula/syntax/layout";
import { parseWithLayout } from "@formula/syntax/parser";
import { printAstWithLayout } from "@formula/syntax/print";
import { tokenizeWithGaps } from "@formula/syntax/tokenizer";
import type { BinaryReader } from "@utils/binary";
import { BinaryWriter, concatUint8Arrays } from "@utils/binary";
import { excelNameKey, sheetNameKey } from "@utils/name-key";

export { FUNCTION_TABLES };

/** Ptg identifiers this codec understands. Base values, before the class bits. */
const Ptg = {
  Exp: 0x01,
  Tbl: 0x02,
  Add: 0x03,
  Sub: 0x04,
  Mul: 0x05,
  Div: 0x06,
  Power: 0x07,
  Concat: 0x08,
  Lt: 0x09,
  Le: 0x0a,
  Eq: 0x0b,
  Ge: 0x0c,
  Gt: 0x0d,
  Ne: 0x0e,
  Isect: 0x0f,
  Union: 0x10,
  Range: 0x11,
  Uplus: 0x12,
  Uminus: 0x13,
  Percent: 0x14,
  Paren: 0x15,
  MissArg: 0x16,
  Str: 0x17,
  Attr: 0x19,
  Err: 0x1c,
  Bool: 0x1d,
  Int: 0x1e,
  Num: 0x1f,
  // Class-tagged tokens: the low five bits identify the token, the upper bits its class
  // (reference / value / array), which does not change the payload.
  Array: 0x00,
  Func: 0x01,
  FuncVar: 0x02,
  Name: 0x03,
  Ref: 0x04,
  Area: 0x05,
  MemArea: 0x06,
  MemErr: 0x07,
  MemNoMem: 0x08,
  MemFunc: 0x09,
  RefErr: 0x0a,
  AreaErr: 0x0b,
  RefN: 0x0c,
  AreaN: 0x0d,
  NameX: 0x19,
  Ref3d: 0x1a,
  Area3d: 0x1b,
  // Tagged only: untagged, these two bytes are `PtgErr` and `PtgBool`, matched first by `decodePtg`.
  RefErr3d: 0x1c,
  AreaErr3d: 0x1d,
  /**
   * `PtgList`. Reserved as 0x18 by MS-XLSB 2.5.98.52 and always followed by an `eptg` of 0x19 — it is
   * one of the "extended" tokens, whose real identity is the second byte.
   */
  List: 0x18
} as const;

/** `PtgAttr` flag bits — MS-XLSB 2.5.98.24 to 2.5.98.33. Exactly one is set, except `Space | Semi`. */
const Attr = {
  Semi: 0x01,
  If: 0x02,
  Choose: 0x04,
  Goto: 0x08,
  Sum: 0x10,
  Baxcel: 0x20,
  Space: 0x40,
  IfError: 0x80
} as const;

/**
 * `PtgAttrSpaceType.type` — MS-XLSB 2.5.98.32. Even values are spaces, odd ones carriage returns, and the pair
 * says where: before an operand or operator, before an opening display parenthesis, before a closing one, or before
 * the whole formula.
 */
const SpaceAt = {
  Expression: 0x00,
  OpenParen: 0x02,
  CloseParen: 0x04,
  Formula: 0x06
} as const;

const BINARY_OP_BY_PTG: ReadonlyMap<number, string> = new Map([
  [Ptg.Add, "+"],
  [Ptg.Sub, "-"],
  [Ptg.Mul, "*"],
  [Ptg.Div, "/"],
  [Ptg.Power, "^"],
  [Ptg.Concat, "&"],
  [Ptg.Lt, "<"],
  [Ptg.Le, "<="],
  [Ptg.Eq, "="],
  [Ptg.Ge, ">="],
  [Ptg.Gt, ">"],
  [Ptg.Ne, "<>"],
  [Ptg.Isect, " "],
  [Ptg.Range, ":"]
]);

const PTG_BY_BINARY_OP: ReadonlyMap<string, number> = new Map(
  [...BINARY_OP_BY_PTG].map(([ptg, op]) => [op, ptg])
);

/** The `iftab`s with control-token encodings of their own. */
const IF_ID = 1;
const CHOOSE_ID = 100;
const IFERROR_ID = 480;

/** Context a token stream needs to be interpretable. */
export interface PtgContext {
  /** Sheet names in workbook declaration order — the order `BrtBundleSh` emits them. */
  readonly sheetNames?: readonly string[];
  /**
   * The `BrtExternSheet` table: what a 3D reference's `ixti` actually indexes.
   *
   * A `PtgRef3d` does **not** carry a sheet index. It carries an index into this table, whose
   * entries then name a range of sheets in a supporting book. Treating `ixti` as a sheet index
   * directly is self-consistent between a reader and a writer and wrong against Excel: in
   * `issues.xlsb` the table is `[{first:0,last:0}, {first:2,last:2}]`, so `ixti = 1` means the
   * *third* sheet, and reading it as the second silently retargets the reference.
   *
   * `first` is signed: `-2` is a workbook-level entry (what a `PtgNameX` to another sheet's local name points
   * through) and `-1` a sheet that no longer exists. `supportingLink` indexes {@link supportingLinks}; absent means
   * 0, this workbook.
   *
   * Absent means the identity mapping, which is what a file with no `BrtExternSheet` implies and
   * what this library's own writer emits a real table for.
   */
  readonly externSheets?: readonly {
    readonly first: number;
    readonly last: number;
    readonly supportingLink?: number;
  }[];
  /**
   * The supporting links `BrtExternSheet` entries point into, in file order. Absent means a single `BrtSupSelf`.
   *
   * Only this workbook (`self`) and add-ins (`addin`, whose function names a `PtgNameX` indexes) are resolvable; an
   * external workbook is recognised so a reference to one fails by name rather than resolving into this one.
   */
  readonly supportingLinks?: readonly SupportingLink[];
  /** Defined names in declaration order; `PtgName` is a one-based index into this. */
  readonly definedNames?: readonly string[];
  /**
   * Each defined name's scope, parallel to {@link definedNames}: the sheet position of a sheet-local name, or
   * `undefined` for a workbook-wide one. A formula naming `Total` on a sheet with its own `Total` means that one,
   * and `Sheet2!Total` means Sheet2's.
   */
  readonly definedNameScopes?: readonly (number | undefined)[];
  /**
   * Whether each defined name's definition is empty, parallel to {@link definedNames}. A call through such a name is a
   * user-defined function, which XLSX spells `_xludf.NAME(…)` — Excel writes a plain, empty `BrtName` for each.
   */
  readonly definedNameEmpty?: readonly boolean[];
  /** Position of the sheet the formula sits on, which decides which of two same-named names it means. */
  readonly sheetIndex?: number;
  /** What the formula is for. Defaults to `cell`. See {@link FormulaRole}. */
  readonly role?: FormulaRole;
  /**
   * Tables by name, for `PtgList`.
   *
   * A structured reference names a table and a column *by name*, and the token carries an `idList` and a
   * column *index* — so resolving one needs the table's id and its column order, which only the package
   * writer knows.
   *
   * `sheet` is here because `PtgList` carries an `ixti` — it is a *3D* reference, like `PtgRef3d`, and
   * naming the table is not enough to locate it.
   */
  readonly tables?: ReadonlyMap<string, TableEntry>;
  /** Cell the formula sits in, for the relative forms `PtgRefN` / `PtgAreaN`. */
  readonly origin?: { readonly row: number; readonly column: number };
  /**
   * Whether to *emit* the relative forms rather than only read them.
   *
   * A conditional-formatting rule's formula is stored once for a whole range and shifted per cell, so its
   * references are offsets from `origin`. A cell formula's are positions. The decoder has always needed
   * `origin` to read a `PtgRefN`; this is the encoder's half, and it is opt-in because turning every cell
   * formula into offsets would be catastrophic and silent.
   */
  readonly relativeToOrigin?: boolean;
}

/** One supporting link — `BrtSupSelf`, `BrtSupSame`, `BrtSupAddin` or `BrtSupBookSrc`. */
export type SupportingLink =
  | { readonly kind: "self" }
  | { readonly kind: "same" }
  | { readonly kind: "addin"; readonly names: readonly string[] }
  | { readonly kind: "external" };

/** A shared-formula reference: this cell defers to the formula at `row`/`column`. */
export interface SharedFormulaReference {
  readonly sharedRow: number;
  readonly sharedColumn: number;
}

/** A decoded formula: the tree, its layout, and whitespace written before the whole formula. */
export interface DecodedFormula {
  readonly ast: AstNode;
  readonly layout: FormulaLayout;
  readonly prefix: string;
}

/** An encoded formula: the `Rgce` and the `RgbExtra` its array constants need. */
export interface EncodedFormula {
  readonly rgce: Uint8Array;
  readonly extra: Uint8Array;
}

/** Bytes a `PtgExtraCol` occupies — a single `Col`, MS-XLSB 2.5.98.42. */
export const PTG_EXTRA_COL_SIZE = 4;

/** Bytes a `PtgExp` occupies in the `Rgce`: the token and a four-byte `Rw`, and nothing else. */
export const PTG_EXP_SIZE = 5;

/**
 * The `Rgce` and `RgbExtra` of a cell that defers to the shared or array formula at `row`/`column`.
 *
 * Both halves are returned together because neither is meaningful alone — the token says which row and the
 * extra says which column, and a writer that emitted one without the other would produce a formula pointing
 * at column 0 of the right row. Verified against Excel: `01 08 00 00 00` with an `RgbExtra` of
 * `02 00 00 00` is cell C9 deferring to the array formula whose master is C9.
 */
export function encodeSharedFormulaReference(
  row: number,
  column: number
): { readonly rgce: Uint8Array; readonly rgbExtra: Uint8Array } {
  return {
    rgce: new BinaryWriter().writeUint8(Ptg.Exp).writeUint32(row).toUint8Array(),
    rgbExtra: new BinaryWriter().writeUint32(column).toUint8Array()
  };
}

// =============================================================================
// Decode: Ptg → AST
// =============================================================================

/**
 * Decode a token stream into an AST, or report that it defers to a shared formula.
 *
 * Kept for callers that want the tree alone; {@link decodeFormula} also returns the layout, and
 * {@link formulaText} the text Excel would show.
 */
export function decodePtg(
  tokens: Uint8Array,
  context: PtgContext,
  where: string,
  extra?: Uint8Array
): AstNode | SharedFormulaReference {
  const decoded = decodeFormula(tokens, context, where, extra);
  return "sharedRow" in decoded ? decoded : decoded.ast;
}

/**
 * The text of a token stream — what Excel shows for it — or the shared formula it defers to.
 */
export function formulaText(
  tokens: Uint8Array,
  context: PtgContext,
  where: string,
  extra?: Uint8Array
): string | SharedFormulaReference {
  const decoded = decodeFormula(tokens, context, where, extra);
  return "sharedRow" in decoded ? decoded : printDecoded(decoded);
}

/** Print a decoded formula with its layout. */
export function printDecoded(decoded: DecodedFormula): string {
  // `_xlfn.ANCHORARRAY(A1)` is how the package stores the spill-range operator; it reads back as `A1#`, as XLSX does.
  return fromFileFormula(decoded.prefix + printAstWithLayout(decoded.ast, decoded.layout));
}

/**
 * Decode a token stream and its `RgbExtra`.
 *
 * `PtgExp` alone means the cell's formula lives elsewhere — Excel's way of storing one expression for a filled
 * range — so that is returned rather than treated as an expression. The token carries only the row; the column
 * lives in a `PtgExtraCol` in `extra` (MS-XLSB 2.5.98.40 and 2.5.98.42).
 *
 * `extra` is also where a `PtgArray`'s values and a `PtgMemArea`'s cached areas live, one structure per such token
 * in token order (MS-XLSB 2.5.98.87), so a caller that has a `RgbExtra` must pass it.
 */
export function decodeFormula(
  tokens: Uint8Array,
  context: PtgContext,
  where: string,
  extra?: Uint8Array
): DecodedFormula | SharedFormulaReference {
  if (tokens.length > 0 && tokens[0] === Ptg.Exp) {
    if (tokens.length !== PTG_EXP_SIZE) {
      throw new XlsbParseError(where, "PtgExp must be the only token in the stream");
    }
    if (extra === undefined || extra.length < PTG_EXTRA_COL_SIZE) {
      throw new XlsbParseError(
        where,
        "PtgExp needs a PtgExtraCol in RgbExtra to supply its column"
      );
    }
    return {
      sharedRow: new DataView(tokens.buffer, tokens.byteOffset, tokens.byteLength).getUint32(
        1,
        true
      ),
      sharedColumn: new DataView(extra.buffer, extra.byteOffset, extra.byteLength).getUint32(
        0,
        true
      )
    };
  }
  const state: DecodeState = {
    reader: new XlsbReader(tokens, 0, where),
    extra: new XlsbReader(extra ?? new Uint8Array(0), 0, `${where} RgbExtra`),
    stack: [],
    layout: new Map(),
    pending: [],
    closedUnions: new Set(),
    emptyNames: new Set(),
    tokenClasses: new Map(),
    prefix: "",
    context,
    where
  };
  while (state.reader.remaining > 0) {
    decodeToken(state);
  }
  if (state.stack.length !== 1) {
    throw new XlsbParseError(
      where,
      `token stream left ${state.stack.length} expressions on the stack instead of one`
    );
  }
  const ast =
    (context.role ?? "cell") === "cell"
      ? withIntersections(state.stack[0]!, state.tokenClasses)
      : state.stack[0]!;
  // A union with no parenthesis of its own is a defined name's list of areas, `S!$A$1,S!$B$2`. Anywhere else the
  // printer supplies the parentheses, because without them the commas would separate arguments.
  if (ast.type === NodeType.UnionRef && !state.closedUnions.has(ast)) {
    layoutOf(state.layout, ast).bare = true;
  }
  return { ast, layout: state.layout, prefix: state.prefix };
}

interface DecodeState {
  readonly reader: BinaryReader;
  readonly extra: BinaryReader;
  readonly stack: AstNode[];
  readonly layout: FormulaLayout;
  /** `PtgAttrSpace` tokens seen since the last token they could belong to. */
  pending: { readonly type: number; readonly count: number }[];
  /** Unions whose own parenthesis has been read — a further one is a display parenthesis. */
  readonly closedUnions: Set<AstNode>;
  /** Name operands naming an empty defined name — a user-defined function when called. */
  readonly emptyNames: Set<AstNode>;
  /** The class each class-tagged token was written in, by the node it decoded to — see {@link withIntersections}. */
  readonly tokenClasses: Map<AstNode, number>;
  prefix: string;
  readonly context: PtgContext;
  readonly where: string;
}

/**
 * The whitespace pending spaces of the given positions describe, removing them.
 *
 * `0x00`/`0x02`/`0x04`/`0x06` are spaces and the odd values carriage returns. A carriage return is a line break,
 * and the XLSX text of the same formula carries it as one — `\r\n` in Excel's XML, which XML parsing reads as `\n`.
 */
function takeSpaces(state: DecodeState, at: number): string {
  let text = "";
  const rest: { type: number; count: number }[] = [];
  for (const entry of state.pending) {
    if ((entry.type & ~1) === at) {
      text += (entry.type & 1 ? "\n" : " ").repeat(entry.count);
    } else {
      rest.push(entry);
    }
  }
  state.pending = rest;
  return text;
}

/**
 * Attach the pending whitespace an operand-like token owns, and push it.
 *
 * Only the kind it owns is taken. The rest stays pending for a later token: Excel writes the space before a
 * function's `)` in front of whatever token ends the last argument — `AND( A1>0, B1>0 )` has it before the `>` —
 * so an operator that cleared everything pending lost it.
 */
function pushOperand(state: DecodeState, node: AstNode): void {
  setLayout(state, node, "lead", takeSpaces(state, SpaceAt.Expression));
  state.stack.push(node);
}

function setLayout(
  state: DecodeState,
  node: AstNode,
  field: "lead" | "op" | "close",
  text: string
): void {
  if (text !== "") {
    const entry = layoutOf(state.layout, node);
    entry[field] = (entry[field] ?? "") + text;
  }
}

function decodeToken(state: DecodeState): void {
  const { reader, stack, where } = state;
  const raw = reader.readUint8();

  // Untagged tokens first: their identifier is the whole byte, so they must be matched
  // before the low-five-bit form or `PtgStr` (0x17) collides with a tagged token.
  if (raw === Ptg.Exp || raw === Ptg.Tbl) {
    throw new ExcelNotSupportedError(
      `Read XLSB formula at ${where}`,
      raw === Ptg.Exp
        ? "a PtgExp that is not the whole formula"
        : "a data table (PtgTbl) is not supported yet"
    );
  }
  const binaryOp = BINARY_OP_BY_PTG.get(raw);
  if (binaryOp !== undefined) {
    const right = pop(stack, where);
    const left = pop(stack, where);
    const node: AstNode = { type: NodeType.BinaryOp, op: binaryOp, left, right };
    const spaces = takeSpaces(state, SpaceAt.Expression);
    // The intersection operator *is* a space, so what Excel records is the run beyond the first.
    setLayout(state, node, "op", binaryOp === " " && spaces !== "" ? ` ${spaces}` : spaces);
    stack.push(node);
    return;
  }
  if (raw === Ptg.Union) {
    // The comma operator produces a union. Flattened so `(a,b,c)` is one node rather
    // than a nest, matching what the parser builds from the same text.
    const right = pop(stack, where);
    const left = pop(stack, where);
    const areas =
      left.type === NodeType.UnionRef && !state.closedUnions.has(left)
        ? [...left.areas, right]
        : [left, right];
    // Whitespace before a comma has no place in the text Excel would show; it is dropped, as Excel drops it.
    takeSpaces(state, SpaceAt.Expression);
    stack.push({ type: NodeType.UnionRef, areas });
    return;
  }

  switch (raw) {
    case Ptg.Uplus:
    case Ptg.Uminus: {
      const node: AstNode = {
        type: NodeType.UnaryOp,
        op: raw === Ptg.Uminus ? "-" : "+",
        operand: pop(stack, where)
      };
      setLayout(state, node, "lead", takeSpaces(state, SpaceAt.Expression));
      stack.push(node);
      return;
    }
    case Ptg.Percent: {
      const node: AstNode = { type: NodeType.Percent, operand: pop(stack, where) };
      setLayout(state, node, "op", takeSpaces(state, SpaceAt.Expression));
      stack.push(node);
      return;
    }
    case Ptg.Paren: {
      const inner = pop(stack, where);
      const open = takeSpaces(state, SpaceAt.OpenParen);
      const close = takeSpaces(state, SpaceAt.CloseParen);
      if (inner.type === NodeType.UnionRef && !state.closedUnions.has(inner)) {
        // A union's own parentheses — the ones that make its commas a union.
        state.closedUnions.add(inner);
        setLayout(state, inner, "lead", open);
        setLayout(state, inner, "close", close);
      } else {
        (layoutOf(state.layout, inner).parens ??= []).push({
          ...(open === "" ? {} : { open }),
          ...(close === "" ? {} : { close })
        });
      }
      stack.push(inner);
      return;
    }
    case Ptg.MissArg:
      pushOperand(state, { type: NodeType.Missing });
      return;
    case Ptg.Str:
      pushOperand(state, { type: NodeType.String, value: readPtgString(reader, where) });
      return;
    case Ptg.Err: {
      const code = reader.readUint8();
      const error = errorTextOf(code);
      if (error === undefined) {
        throw new XlsbParseError(where, `unknown error code 0x${code.toString(16)} in PtgErr`);
      }
      pushOperand(state, { type: NodeType.Error, value: error });
      return;
    }
    case Ptg.Bool:
      pushOperand(state, { type: NodeType.Boolean, value: reader.readUint8() !== 0 });
      return;
    case Ptg.Int:
      pushOperand(state, { type: NodeType.Number, value: reader.readUint16() });
      return;
    case Ptg.Num:
      pushOperand(state, { type: NodeType.Number, value: reader.readFloat64() });
      return;
    case Ptg.Attr:
      decodeAttribute(state);
      return;
    default:
      break;
  }

  const depth = stack.length;
  const before = stack[depth - 1];
  decodeClassTagged(state, raw, raw & 0x1f);
  const after = stack[stack.length - 1];
  if (after !== undefined && (stack.length !== depth || after !== before)) {
    state.tokenClasses.set(after, raw & 0x60);
  }
}

/** A reference, or a call that returns one — what `@` intersects; not a function passed by name. */
function isIntersectable(node: AstNode): boolean {
  switch (node.type) {
    case NodeType.CellRef:
    case NodeType.RangeRef:
    case NodeType.ColRangeRef:
    case NodeType.RowRangeRef:
    case NodeType.StructuredRef:
      return true;
    case NodeType.Name:
      return !node.name.toLowerCase().startsWith("_xleta.");
    case NodeType.FunctionCall:
      return true;
    default:
      return false;
  }
}

/**
 * A cell formula's `tree` with an `@` around each reference the file wrote in the value class where its text alone is
 * written in another. Excel writes `@x` (XLSX `_xlfn.SINGLE(x)`) as a call to the `_xlfn.SINGLE` stub, which reads
 * back as `@x`, except where `x` alone is intersected in its place anyway (`bareIntersections`) — and those bytes read
 * back bare, as the XLSX reader of `<f>A1:A3</f>` leaves it. A value-class reference where the text alone would be in
 * another class is an intersection written without the call, and reading it bare would sum a whole range where the
 * file sums one cell (`SUM` of a `45` area). A cell formula only: the encoder's classes are Excel's there
 * (`operand-classes.node.test.ts`), and not checked against Excel's in conditional formats, validations or names,
 * where Excel writes the call. Only references — an `@` changes nothing else — and not a function passed by name
 * (`_xleta.SUM`), which is always in the value class. Unchanged parts keep their identity, so their layout still
 * applies.
 */
function withIntersections(tree: AstNode, written: ReadonlyMap<AstNode, number>): AstNode {
  let assigned: ReadonlyMap<AstNode, number> | undefined;
  for (const tokenClass of written.values()) {
    if (tokenClass === VALUE_CLASS) {
      assigned = assignClasses(tree, "cell").own;
      break;
    }
  }
  if (assigned === undefined) {
    return tree;
  }
  const expected = assigned;
  const rewrite = (value: unknown): unknown => {
    if (Array.isArray(value)) {
      let changed = false;
      const mapped = value.map(item => {
        const next = rewrite(item);
        changed ||= next !== item;
        return next;
      });
      return changed ? mapped : value;
    }
    if (typeof value !== "object" || value === null || !("type" in value)) {
      return value;
    }
    const node = value as AstNode;
    let copy: Record<string, unknown> | undefined;
    for (const [key, item] of Object.entries(node)) {
      const next = rewrite(item);
      if (next !== item) {
        copy ??= { ...node };
        copy[key] = next;
      }
    }
    const inner = (copy ?? node) as AstNode;
    const wanted = expected.get(node);
    return written.get(node) === VALUE_CLASS &&
      wanted !== undefined &&
      wanted !== VALUE_CLASS &&
      isIntersectable(node)
      ? { type: NodeType.UnaryOp, op: "@", operand: inner }
      : inner;
  };
  return rewrite(tree) as AstNode;
}

/**
 * `PtgAttr` — MS-XLSB 2.5.98.24 to 2.5.98.33.
 *
 * Most are control tokens the reader steps over: `PtgAttrIf`, `PtgAttrIfError`, `PtgAttrChoose` and `PtgAttrGoto`
 * are jumps that let an evaluator skip a branch, and the function token after them already says what was called;
 * `PtgAttrSemi` marks a volatile formula. Two carry meaning. `PtgAttrSum` *is* a call — Excel writes `SUM` with one
 * argument as this four-byte token, and leaving the operand on the stack read `=SUM(F5-I5)` as `=F5-I5`, 93 times
 * across the Excel-converted corpus. `PtgAttrSpace` is the author's whitespace, kept until the token it precedes.
 */
function decodeAttribute(state: DecodeState): void {
  const { reader, where } = state;
  const flags = reader.readUint8();
  if ((flags & Attr.Space) !== 0) {
    // `PtgAttrSpace`, or `PtgAttrSpaceSemi` (0x41), which is the same plus "volatile".
    const type = reader.readUint8();
    const count = reader.readUint8();
    if (type === SpaceAt.Formula || type === SpaceAt.Formula + 1) {
      state.prefix += (type & 1 ? "\n" : " ").repeat(count);
    } else if (type <= 0x05) {
      state.pending.push({ type, count });
    }
    return;
  }
  if ((flags & Attr.Choose) !== 0) {
    // A jump table whose length is one more than `cOffset`.
    const count = reader.readUint16();
    reader.skip((count + 1) * 2);
    return;
  }
  if ((flags & Attr.Sum) !== 0) {
    reader.skip(2);
    const call: FunctionCallNode = {
      type: NodeType.FunctionCall,
      name: "SUM",
      args: [pop(state.stack, where)]
    };
    finishCall(state, call);
    return;
  }
  // `PtgAttrSemi`, `PtgAttrIf`, `PtgAttrIfError`, `PtgAttrGoto`, `PtgAttrBaxcel`: two bytes, nothing to build. Any
  // whitespace pending stays pending — Excel writes the space before a function's `)` *before* the jump that
  // precedes the function token.
  if ((flags & (Attr.Semi | Attr.If | Attr.IfError | Attr.Goto | Attr.Baxcel)) === 0) {
    throw new XlsbParseError(where, `PtgAttr with unknown flags 0x${flags.toString(16)}`);
  }
  reader.skip(2);
}

/** Give a decoded call the whitespace its function token owns, and push it. */
function finishCall(state: DecodeState, call: FunctionCallNode): void {
  setLayout(state, call, "lead", takeSpaces(state, SpaceAt.Expression));
  setLayout(state, call, "close", takeSpaces(state, SpaceAt.CloseParen));
  state.stack.push(call);
}

function decodeClassTagged(state: DecodeState, raw: number, base: number): void {
  const { reader, stack, context, where } = state;
  switch (base) {
    case Ptg.Array:
      pushOperand(state, readArrayConstant(state));
      return;
    case Ptg.Func: {
      // A fixed-arity call. The token carries no argument count — the count is implied by the function.
      const functionId = reader.readUint16() & 0x7fff;
      const name = FUNCTION_NAME_BY_ID.get(functionId);
      if (name === undefined) {
        throw new ExcelNotSupportedError(
          `Read XLSB formula at ${where}`,
          `built-in function id ${functionId} is not in this codec's table`
        );
      }
      const arity = ARITY_BY_ID.get(functionId);
      if (arity === undefined) {
        // A `PtgFunc` naming a function this codec knows to be variadic is a contradiction: the token says the
        // count is implied and the table says there is none to imply. Reported rather than guessed.
        throw new XlsbParseError(
          where,
          `${name} was called through PtgFunc, which implies a fixed argument count, but it is variadic`
        );
      }
      if (stack.length < arity) {
        throw new XlsbParseError(
          where,
          `${name} wants ${arity} argument(s) but the stack holds ${stack.length}`
        );
      }
      finishCall(state, {
        type: NodeType.FunctionCall,
        name,
        args: stack.splice(stack.length - arity, arity)
      });
      return;
    }
    case Ptg.FuncVar: {
      const argumentCount = reader.readUint8() & 0x7f;
      const functionId = reader.readUint16() & 0x7fff;
      if (stack.length < argumentCount) {
        throw new XlsbParseError(
          where,
          `a call wants ${argumentCount} operand(s) but the stack holds ${stack.length}`
        );
      }
      // `0x00FF` is not a function: it means "the callee is the first operand", which is how Excel calls
      // anything the `Ftab` has no id for. The first of the `cparams` is therefore the name of the function.
      if (functionId === IFTAB_CALL_BY_NAME) {
        if (argumentCount < 1) {
          throw new XlsbParseError(
            where,
            "a call through a name has no operand naming the function"
          );
        }
        const operands = stack.splice(stack.length - argumentCount, argumentCount);
        const callee = operands[0]!;
        // An argument list applied to the result of a call — `LAMBDA(x,x*2)(5)`; see `INVOCATION_CLASSES`.
        if (callee.type === NodeType.FunctionCall) {
          (callee.invocations ??= []).push(operands.slice(1));
          stack.push(callee);
          return;
        }
        if (callee.type !== NodeType.Name) {
          throw new XlsbParseError(
            where,
            `a call through iftab 0x00FF must name its function, and the first operand is not a name`
          );
        }
        // **The `_xlfn.` prefix stays.** It is how the XLSX container spells the same call — Excel's own XLSX of
        // this cell reads `_xlfn.CONCAT(…)` — and a reader that dropped it produced `CONCAT(…)`, which the XLSX
        // writer then wrote verbatim and Excel opened as `#NAME?`.
        //
        // A call through a plain *empty* name is a user-defined function, which XLSX spells `_xludf.NAME(…)`: Excel's
        // own XLSX of an unknown `CONCAT(A1,B1)` holds `_xludf.CONCAT(A1,B1)` where its XLSB has `PtgName` to an empty
        // `CONCAT`.
        const call: FunctionCallNode = {
          type: NodeType.FunctionCall,
          name: state.emptyNames.has(callee)
            ? `${USER_FUNCTION_PREFIX}${callee.name}`
            : callee.name,
          args: operands.slice(1)
        };
        // The whitespace before the name was written before the name's token.
        const calleeLead = state.layout.get(callee)?.lead;
        if (calleeLead !== undefined) {
          setLayout(state, call, "lead", calleeLead);
        }
        finishCall(state, call);
        return;
      }
      const name = FUNCTION_NAME_BY_ID.get(functionId);
      if (name === undefined) {
        throw new ExcelNotSupportedError(
          `Read XLSB formula at ${where}`,
          `built-in function id ${functionId} is not in this codec's table`
        );
      }
      finishCall(state, {
        type: NodeType.FunctionCall,
        name,
        args: stack.splice(stack.length - argumentCount, argumentCount)
      });
      return;
    }
    case Ptg.Name: {
      const index = reader.readUint32();
      const name = context.definedNames?.[index - 1];
      if (name === undefined) {
        throw new XlsbParseError(
          where,
          `PtgName references defined name ${index}, which is absent`
        );
      }
      const node: AstNode = { type: NodeType.Name, name };
      if (context.definedNameEmpty?.[index - 1] === true && !name.toLowerCase().startsWith("_xl")) {
        state.emptyNames.add(node);
      }
      pushOperand(state, node);
      return;
    }
    case Ptg.NameX: {
      pushOperand(state, readNameX(state));
      return;
    }
    // In a defined name every reference to a sheet is 3D, so a plain one is Excel's `!$A$1` — "the sheet the name is
    // used on". The empty sheet prints as that bare `!`; read without it, the definition became a different one.
    case Ptg.Ref:
      pushOperand(state, readRef(reader, currentSheetOf(context), undefined, where));
      return;
    case Ptg.Area:
      pushOperand(state, readArea(reader, currentSheetOf(context), undefined, where));
      return;
    case Ptg.RefN:
      pushOperand(state, readRelativeRef(reader, requireOrigin(context, where, "PtgRefN"), where));
      return;
    case Ptg.AreaN:
      pushOperand(
        state,
        readRelativeArea(reader, requireOrigin(context, where, "PtgAreaN"), where)
      );
      return;
    case Ptg.Ref3d: {
      const { sheet, endSheet } = readSheetPair(reader, context, where);
      pushOperand(state, readRef(reader, sheet, endSheet, where));
      return;
    }
    case Ptg.Area3d: {
      const { sheet, endSheet } = readSheetPair(reader, context, where);
      pushOperand(state, readArea(reader, sheet, endSheet, where));
      return;
    }
    case Ptg.List:
      pushOperand(state, readList(reader, context, where));
      return;
    case Ptg.RefErr:
      reader.skip(6);
      pushOperand(state, { type: NodeType.Error, value: "#REF!" });
      return;
    case Ptg.AreaErr:
      reader.skip(12);
      pushOperand(state, { type: NodeType.Error, value: "#REF!" });
      return;
    // The 3D forms of the two above: an `ixti`, then the same unused bytes (MS-XLSB 2.5.98.71 and 2.5.98.21). What
    // Excel writes once the cells a cross-sheet reference named have been deleted — and it shows them qualified,
    // `Other!#REF!`, so the sheet is kept.
    case Ptg.RefErr3d:
    case Ptg.AreaErr3d: {
      const ixti = reader.readUint16();
      reader.skip(base === Ptg.RefErr3d ? 6 : 12);
      // In a conditional format or a validation the formula's own sheet is not named: Excel writes the rule
      // `#REF!>1` as `5c 00 00 …`, through the entry for the sheet it sits on, and shows it unqualified. A defined
      // name keeps it — a local name's `Feststellanlagen!#REF!` reads back qualified in Excel's XLSX.
      const resolved = resolveSheet(context, ixti);
      const template =
        context.role === "conditional" || context.role === "validation" || context.role === "list";
      const sheet =
        template &&
        resolved !== undefined &&
        resolved === context.sheetNames?.[context.sheetIndex ?? -1]
          ? undefined
          : resolved;
      pushOperand(state, {
        type: NodeType.Error,
        value: "#REF!",
        ...(sheet === undefined ? {} : { sheet })
      });
      return;
    }
    // Mem tokens precede a reference subexpression and say whether its result was cached. The subexpression's
    // own tokens follow and are decoded like any others; the cache is not needed to say what the formula is.
    case Ptg.MemArea: {
      reader.skip(4 + 2);
      // Its cached areas are a `PtgExtraMem` in `RgbExtra`, which has to be stepped over to keep any later
      // `PtgArray` aligned with its own values.
      const count = state.extra.readUint32();
      state.extra.skip(count * 16);
      return;
    }
    case Ptg.MemErr:
    case Ptg.MemNoMem:
      reader.skip(4 + 2);
      return;
    case Ptg.MemFunc:
      reader.skip(2);
      return;
    default:
      throw new ExcelNotSupportedError(
        `Read XLSB formula at ${where}`,
        `Ptg token 0x${raw.toString(16).padStart(2, "0")} is not supported yet`
      );
  }
}

/**
 * `PtgArray` and its `PtgExtraArray` — MS-XLSB 2.5.98.23 and 2.5.98.41.
 *
 * The token's fourteen bytes are unused; the values are in `RgbExtra`: a row count, a column count, then one
 * `SerAr` per element in row-major order, each led by a byte naming its kind (2.5.98.93). Excel's own
 * `{1,"a";TRUE,#N/A}` is `02 00 00 00 02 00 00 00 | 00 <f64 1> | 01 01 00 61 00 | 02 01 | 04 2a 00 00 00`.
 */
function readArrayConstant(state: DecodeState): AstNode {
  const { reader, extra, where } = state;
  reader.skip(14);
  const rowCount = extra.readUint32();
  const columnCount = extra.readUint32();
  if (rowCount === 0 || columnCount === 0 || rowCount * columnCount > extra.remaining) {
    throw new XlsbParseError(
      where,
      `PtgExtraArray of ${rowCount}×${columnCount} does not fit its data`
    );
  }
  const rows: AstNode[][] = [];
  for (let row = 0; row < rowCount; row++) {
    const values: AstNode[] = [];
    for (let column = 0; column < columnCount; column++) {
      const kind = extra.readUint8();
      switch (kind) {
        case 0x00:
          values.push({ type: NodeType.Number, value: extra.readFloat64() });
          break;
        case 0x01:
          values.push({ type: NodeType.String, value: readPtgString(extra, where) });
          break;
        case 0x02:
          values.push({ type: NodeType.Boolean, value: extra.readUint8() !== 0 });
          break;
        case 0x04: {
          const code = extra.readUint8();
          extra.skip(3);
          const error = errorTextOf(code);
          if (error === undefined) {
            throw new XlsbParseError(where, `unknown error code 0x${code.toString(16)} in SerErr`);
          }
          values.push({ type: NodeType.Error, value: error });
          break;
        }
        default:
          throw new XlsbParseError(where, `SerAr of unknown kind 0x${kind.toString(16)}`);
      }
    }
    rows.push(values);
  }
  return { type: NodeType.Array, rows };
}

/**
 * `PtgNameX` — MS-XLSB 2.5.98.61: an `ixti` and a one-based name index.
 *
 * The specification describes it as a name in another workbook, and Excel also uses it for a name *scoped to
 * another sheet of this one*: `Other!LocalO` on sheet `S` is `59 02 00 0b 00 00 00`, where `ixti` 2 is a
 * workbook-level entry (`BrtSupSelf`, -2, -2) and 11 indexes this workbook's own `BrtName` list. The sheet in the
 * text is the name's own scope.
 */
function readNameX(state: DecodeState): AstNode {
  const { reader, context, where } = state;
  const ixti = reader.readUint16();
  const index = reader.readUint32();
  const entry = context.externSheets?.[ixti];
  const link = context.supportingLinks?.[entry?.supportingLink ?? 0] ?? { kind: "self" };
  if (link.kind === "addin") {
    const name = link.names[index - 1];
    if (name === undefined) {
      throw new XlsbParseError(where, `PtgNameX names add-in function ${index}, which is absent`);
    }
    return { type: NodeType.Name, name };
  }
  if (link.kind !== "self") {
    throw new ExcelNotSupportedError(
      `Read XLSB formula at ${where}`,
      "a defined name in another workbook is not supported yet"
    );
  }
  const name = context.definedNames?.[index - 1];
  if (name === undefined) {
    throw new XlsbParseError(where, `PtgNameX references defined name ${index}, which is absent`);
  }
  const scope = context.definedNameScopes?.[index - 1];
  const sheet = scope === undefined ? undefined : context.sheetNames?.[scope];
  return { type: NodeType.Name, name, ...(sheet === undefined ? {} : { sheet }) };
}

/**
 * `PtgList` — a structured reference; see {@link emitStructuredRef} for the layout.
 */
function readList(reader: BinaryReader, context: PtgContext, where: string): AstNode {
  // The `eptg` that follows, which the specification fixes at 0x19. Consumed even though its value is
  // known, because skipping it shifts every field after it.
  const eptg = reader.readUint8();
  if (eptg !== 0x19) {
    throw new ExcelNotSupportedError(
      `Read XLSB formula at ${where}`,
      `extended Ptg 0x18 0x${eptg.toString(16)} is not supported yet`
    );
  }
  reader.readUint16(); // ixti
  const flags = reader.readUint16();
  const listIndex = reader.readUint32();
  const first = reader.readUint16();
  const last = reader.readUint16();
  // `invalid` (bit 12): the reference no longer names a table, and every other field is unused. Excel writes it for
  // `Tb[@x]` in a cell outside the table — `18 19 ff ff 00 30 ff ff ff ff …` — and shows `#REF!`.
  if ((flags & 0x1000) !== 0) {
    return { type: NodeType.Error, value: "#REF!" };
  }
  // `nonresident` (bit 13): a table in another workbook, with its name in a `PtgExtraList` this codec does not read.
  if ((flags & 0x2000) !== 0) {
    throw new ExcelNotSupportedError(
      `Read XLSB formula at ${where}`,
      "a structured reference to a table in another workbook is not supported yet"
    );
  }
  const table = [...(context.tables ?? [])].find(([, entry]) => entry.id === listIndex);
  if (table === undefined) {
    throw new XlsbParseError(
      where,
      `PtgList names table ${listIndex}, which this file does not declare`
    );
  }
  const [tableName, entry] = table;
  const columnCount = flags & 0x03;
  const columns =
    columnCount === 0
      ? []
      : columnCount === 1
        ? [entry.columns[first] ?? ""]
        : [entry.columns[first] ?? "", entry.columns[last] ?? ""];
  return {
    type: NodeType.StructuredRef,
    tableName,
    columns: columns.filter(name => name !== ""),
    specials: specialsOf((flags >>> 2) & 0x1f)
  };
}

/**
 * `PtgStr`: a two-byte character count, then UTF-16LE. No flag byte.
 *
 * Established from Excel's own output. On a real `=("a"&"b")` the tokens are `17 01 00 61 00`, five bytes: the
 * ptg, a count of one, and one UTF-16 code unit. A BIFF8-style flag byte would have consumed the `61`.
 */
function readPtgString(reader: BinaryReader, where: string): string {
  const length = reader.readUint16();
  if (length > Math.floor(reader.remaining / 2)) {
    throw new XlsbParseError(where, `PtgStr declares ${length} character(s) beyond the stream`);
  }
  let text = "";
  for (let index = 0; index < length; index++) {
    text += String.fromCharCode(reader.readUint16());
  }
  return text;
}

/** The sheet a single-sheet `ixti` resolves to, or `undefined` when it names none. */
function resolveSheet(context: PtgContext, ixti: number): string | undefined {
  const entry = context.externSheets?.[ixti] ?? { first: ixti, last: ixti };
  if ((entry.supportingLink ?? 0) !== 0 || entry.first < 0) {
    return undefined;
  }
  return context.sheetNames?.[entry.first];
}

function readSheetPair(
  reader: BinaryReader,
  context: PtgContext,
  where: string
): { sheet: string; endSheet: string | undefined } {
  const ixti = reader.readUint16();
  // Resolved through `BrtExternSheet`, not used as a sheet index. See `PtgContext.externSheets`.
  const entry = context.externSheets?.[ixti] ?? { first: ixti, last: ixti };
  const link = context.supportingLinks?.[entry.supportingLink ?? 0] ?? { kind: "self" };
  if (link.kind !== "self") {
    throw new ExcelNotSupportedError(
      `Read XLSB formula at ${where}`,
      "a reference to another workbook is not supported yet"
    );
  }
  const sheet = entry.first < 0 ? undefined : context.sheetNames?.[entry.first];
  if (sheet === undefined) {
    throw new XlsbParseError(
      where,
      `3D reference ixti ${ixti} resolves to sheet ${entry.first}, which is not in the workbook`
    );
  }
  // A table entry spanning several sheets is a `Sheet1:Sheet3!A1` style reference.
  const endSheet = entry.last !== entry.first ? context.sheetNames?.[entry.last] : undefined;
  return { sheet, endSheet };
}

function requireOrigin(
  context: PtgContext,
  where: string,
  token: string
): { row: number; column: number } {
  if (!context.origin) {
    throw new XlsbParseError(where, `${token} is relative but no origin cell was supplied`);
  }
  return context.origin;
}

/** `RgceLoc`: row as `u32`, then column with its two absolute-reference flags. */
/** The sheet a plain `PtgRef` / `PtgArea` names: none in a formula, and the using sheet (`""`) in a defined name. */
function currentSheetOf(context: PtgContext): string | undefined {
  return context.role === "name" ? "" : undefined;
}

function readRef(
  reader: BinaryReader,
  sheet: string | undefined,
  endSheet: string | undefined,
  where: string
): CellRefNode {
  const row = checkSheetIndex(reader.readUint32(), "row", where, "PtgRef");
  const columnAndFlags = reader.readUint16();
  return cellRef(row, columnAndFlags, sheet, endSheet);
}

/**
 * `RgceArea` — MS-XLSB 2.5.98.89: `rowFirst`, `rowLast`, `columnFirst`, `columnLast`, all positions.
 *
 * The two rows come before the two columns. A cell pair (`row, col, row, col`) is the obvious misreading.
 */
function readArea(
  reader: BinaryReader,
  sheet: string | undefined,
  endSheet: string | undefined,
  where: string
): AstNode {
  // An `RwLongU` is at most 0xFFFFF. A larger one is a malformed token, not a reference to be modelled: a defined
  // name holding one reached a consumer that walks its area cell by cell.
  const firstRow = checkSheetIndex(reader.readUint32(), "row", where, "PtgArea");
  const lastRow = checkSheetIndex(reader.readUint32(), "row", where, "PtgArea");
  const firstColumn = reader.readUint16();
  const lastColumn = reader.readUint16();
  return areaNode(firstRow, lastRow, firstColumn, lastColumn, sheet, endSheet);
}

/**
 * An area from resolved positions, each column still carrying its two `ColRelShort` flags.
 *
 * An axis pinned to the sheet's full extent is a *whole-row* or *whole-column* reference, and reading it back as
 * an ordinary range loses the shape: `$1:$1` becomes `$A$1:$XFD$1`, which computes the same thing but is not what
 * the author wrote — and for `_xlnm.Print_Titles` it is not even valid. The axis that *is* named keeps its `$`
 * markers: `C:C` in a template moves with the cell and `$C:$C` does not.
 */
function areaNode(
  firstRow: number,
  lastRow: number,
  firstColumn: number,
  lastColumn: number,
  sheet: string | undefined,
  endSheet: string | undefined
): AstNode {
  const scope = {
    ...(sheet === undefined ? {} : { sheet }),
    ...(endSheet === undefined ? {} : { endSheet })
  };
  const spansAllColumns =
    (firstColumn & 0x3fff) === 0 && (lastColumn & 0x3fff) === XLSB_MAX_COLUMNS - 1;
  const spansAllRows = firstRow === 0 && lastRow === XLSB_MAX_ROWS - 1;
  // The whole sheet is shown as every *row* — Excel's `SUM(1:1048576)` is the same token as `A1:XFD1048576` and
  // reads back as rows.
  if (spansAllColumns) {
    return {
      type: NodeType.RowRangeRef,
      startRow: firstRow + 1,
      endRow: lastRow + 1,
      ...((firstColumn & 0x8000) === 0 ? { startAbsolute: true } : {}),
      ...((lastColumn & 0x8000) === 0 ? { endAbsolute: true } : {}),
      ...scope
    };
  }
  if (spansAllRows && !spansAllColumns) {
    return {
      type: NodeType.ColRangeRef,
      startCol: encodeCol(firstColumn & 0x3fff),
      endCol: encodeCol(lastColumn & 0x3fff),
      ...((firstColumn & 0x4000) === 0 ? { startAbsolute: true } : {}),
      ...((lastColumn & 0x4000) === 0 ? { endAbsolute: true } : {}),
      ...scope
    };
  }

  return {
    type: NodeType.RangeRef,
    start: cellRef(firstRow, firstColumn, undefined, undefined),
    end: cellRef(lastRow, lastColumn, undefined, undefined),
    ...scope
  };
}

/** `RgceLocRel` — MS-XLSB 2.5.98.92: a row, then a `ColRelShort`. */
function readRelativeRef(
  reader: BinaryReader,
  origin: { row: number; column: number },
  where: string
): CellRefNode {
  const row = reader.readUint32();
  const columnAndFlags = reader.readUint16();
  const resolved = resolveRelative(row, columnAndFlags, origin, where);
  return cellRef(resolved.row, resolved.columnAndFlags, undefined, undefined);
}

/**
 * `RgceAreaRel` — MS-XLSB 2.5.98.90: `rowFirst`, `rowLast`, `columnFirst`, `columnLast`.
 *
 * **Rows first, then columns — the same order as `RgceArea`, and not two `RgceLocRel`s back to back.**
 */
function readRelativeArea(
  reader: BinaryReader,
  origin: { row: number; column: number },
  where: string
): AstNode {
  const firstRow = reader.readUint32();
  const lastRow = reader.readUint32();
  const firstColumn = reader.readUint16();
  const lastColumn = reader.readUint16();
  const first = resolveRelative(firstRow, firstColumn, origin, where);
  const last = resolveRelative(lastRow, lastColumn, origin, where);
  return areaNode(
    first.row,
    last.row,
    first.columnAndFlags,
    last.columnAndFlags,
    undefined,
    undefined
  );
}

/**
 * One row/column pair of a relative structure, resolved against the cell the formula is evaluated in.
 *
 * Each axis is an offset only when its own flag says so (`fRwRel`, `fColRel`); otherwise it is an ordinary
 * position. **Offsets wrap through the sheet, as MS-XLSB 2.5.98.92 specifies** — a modulus, which is also what makes
 * both spellings of a negative row offset read the same: Excel writes -1 as `0xFFFFF`.
 */
function resolveRelative(
  row: number,
  columnAndFlags: number,
  origin: { row: number; column: number },
  where: string
): { row: number; columnAndFlags: number } {
  const flags = columnAndFlags & 0xc000;
  const column = columnAndFlags & 0x3fff;
  return {
    // An offset wraps; an absolute row is a position and must be on the sheet.
    row:
      (columnAndFlags & 0x8000) !== 0
        ? (origin.row + row) & ROW_MASK
        : checkSheetIndex(row, "row", where, "PtgRefN"),
    columnAndFlags:
      flags | ((columnAndFlags & 0x4000) !== 0 ? (origin.column + column) & 0x3fff : column)
  };
}

function cellRef(
  row: number,
  columnAndFlags: number,
  sheet: string | undefined,
  endSheet: string | undefined
): CellRefNode {
  // The two high bits are the absolute-reference flags — set means *relative*, which is
  // the inversion that makes a hand-written decoder produce `$A$1` for `A1`.
  const rowRelative = (columnAndFlags & 0x8000) !== 0;
  const columnRelative = (columnAndFlags & 0x4000) !== 0;
  return {
    type: NodeType.CellRef,
    col: encodeCol(columnAndFlags & 0x3fff),
    row: String(row + 1),
    colAbsolute: !columnRelative,
    rowAbsolute: !rowRelative,
    ...(sheet === undefined ? {} : { sheet }),
    ...(endSheet === undefined ? {} : { endSheet })
  };
}

function pop(stack: AstNode[], where: string): AstNode {
  const node = stack.pop();
  if (node === undefined) {
    throw new XlsbParseError(where, "operator with no operand on the stack");
  }
  return node;
}

// =============================================================================
// Encode: AST → Ptg
// =============================================================================

/**
 * Whether a formula names a cell on a sheet, explicitly — `Other!A1`, `SUM(S1:S3!B2)`.
 *
 * Such a formula cannot be a *template*: `SharedParsedFormula` and `CFParsedFormula` both forbid `PtgRef3d` and
 * `PtgArea3d` (MS-XLSB 2.5.98.98 and 2.5.98.6), and there is no relative 3D form for them to use instead.
 */
export function referencesAnotherSheet(node: AstNode): boolean {
  switch (node.type) {
    case NodeType.CellRef:
    case NodeType.RangeRef:
    case NodeType.ColRangeRef:
    case NodeType.RowRangeRef:
    case NodeType.Error:
    case NodeType.Name:
      return node.sheet !== undefined;
    case NodeType.BinaryOp:
      return referencesAnotherSheet(node.left) || referencesAnotherSheet(node.right);
    case NodeType.UnaryOp:
    case NodeType.Percent:
      return referencesAnotherSheet(node.operand);
    case NodeType.FunctionCall:
      return (
        node.args.some(referencesAnotherSheet) ||
        (node.invocations ?? []).some(args => args.some(referencesAnotherSheet))
      );
    case NodeType.UnionRef:
      return node.areas.some(referencesAnotherSheet);
    default:
      return false;
  }
}

/**
 * Encode formula text — the whole pipeline, keeping the author's whitespace and parentheses.
 *
 * Whitespace before the first token and after the last is not kept, which is what Excel does with the same text:
 * its XLSX reading of `" A1+B1"` shows `A1+B1`.
 */
export function encodeFormula(text: string, context: PtgContext, where: string): EncodedFormula {
  // The stored spelling: `FILTER` → `_xlfn._xlws.FILTER`, `LET(x,…)` → `_xlfn.LET(_xlpm.x,…)`. The stubs and the
  // `_xlpm.` names a formula calls are created from the same spelling (`write/package.ts`).
  const { tokens, gaps } = tokenizeWithGaps(toFileFormula(text));
  gaps[0] = "";
  const { ast, layout } = parseWithLayout(tokens, gaps);
  return encodeFormulaAst(ast, context, where, layout);
}

/**
 * Encode an AST as a `Rgce` token stream.
 *
 * Kept for callers holding a tree with no array constant in it; {@link encodeFormulaAst} also returns the
 * `RgbExtra` an array constant needs, and refusing one here is better than writing a `PtgArray` with no values.
 */
export function encodePtg(node: AstNode, context: PtgContext, where: string): Uint8Array {
  const encoded = encodeFormulaAst(node, context, where, undefined);
  if (encoded.extra.length > 0) {
    unsupported(where, "an array constant without its RgbExtra");
  }
  return encoded.rgce;
}

/** The call a file stores each operator as, by operator — see {@link operatorsAsCalls}. */
const OPERATOR_CALLS: ReadonlyMap<string, string> = new Map([
  ["#", "_xlfn.ANCHORARRAY"],
  ["@", "_xlfn.SINGLE"]
]);

/**
 * `value` with every spill-range operator `A1#` made the `_xlfn.ANCHORARRAY(A1)` call it is stored as, and every
 * implicit-intersection operator `@x` the `_xlfn.SINGLE(x)` call — the operand, then `PtgName` of the stub and
 * `PtgFuncVar(255)`, as Excel's own XLSB saves of `future-functions.xlsx` and `single.xlsx` show
 * (`future-function-tokens.node.test.ts`, `single-tokens.node.test.ts`); `bareIntersections` decides where `@` is
 * written as its operand alone. Text reaches the encoder with both already calls (`toFileFormula`); a tree may not.
 * Unchanged parts keep their identity, and a rebuilt node takes over its original's layout.
 */
function operatorsAsCalls(value: unknown, layout: FormulaLayout): unknown {
  if (Array.isArray(value)) {
    let changed = false;
    const mapped = value.map(item => {
      const next = operatorsAsCalls(item, layout);
      changed ||= next !== item;
      return next;
    });
    return changed ? mapped : value;
  }
  if (typeof value !== "object" || value === null || !("type" in value)) {
    return value;
  }
  const node = value as AstNode;
  const call = node.type === NodeType.UnaryOp ? OPERATOR_CALLS.get(node.op) : undefined;
  let rebuilt: AstNode | undefined;
  if (call !== undefined && node.type === NodeType.UnaryOp) {
    rebuilt = {
      type: NodeType.FunctionCall,
      name: call,
      args: [operatorsAsCalls(node.operand, layout) as AstNode]
    };
  } else {
    let copy: Record<string, unknown> | undefined;
    for (const [key, item] of Object.entries(node)) {
      const next = operatorsAsCalls(item, layout);
      if (next !== item) {
        copy ??= { ...node };
        copy[key] = next;
      }
    }
    rebuilt = copy as AstNode | undefined;
  }
  if (rebuilt === undefined) {
    return node;
  }
  const entry = layout.get(node);
  if (entry !== undefined) {
    layout.set(rebuilt, entry);
  }
  return rebuilt;
}

/** Encode an AST, with its layout when there is one. */
export function encodeFormulaAst(
  node: AstNode,
  context: PtgContext,
  where: string,
  layout: FormulaLayout | undefined
): EncodedFormula {
  // A file Excel will not open, however faithfully it is encoded — see `nonReferenceArgument`.
  const refused = nonReferenceArgument(node);
  if (refused !== undefined) {
    throw new ExcelNotSupportedError(`Write XLSB formula at ${where}`, refused);
  }
  // A copy: an invoked call hands its whitespace to the call it applies to, which is a node of the encoder's own.
  const ownLayout: FormulaLayout = new Map(layout);
  const tree = operatorsAsCalls(node, ownLayout) as AstNode;
  const role = context.role ?? "cell";
  const bare = bareIntersections(tree, role);
  const encoder: Encoder = {
    out: new TokenBuffer(),
    extra: new TokenBuffer(),
    classes: assignClasses(tree, role, bare),
    bare,
    layout: ownLayout,
    context,
    where,
    volatile: false
  };
  emit(encoder, tree, undefined);
  const body = encoder.out.toUint8Array();
  // **A volatile formula starts with `PtgAttrSemi`.** Excel writes `19 01 00 00` in front of every formula that
  // calls `NOW`, `RAND`, `OFFSET`, `INDIRECT`, `CELL` or another volatile function; it is the token that tells a
  // recalculation engine the cell must be recomputed even when nothing it reads has changed.
  const rgce = encoder.volatile
    ? concatUint8Arrays([Uint8Array.of(Ptg.Attr, Attr.Semi, 0, 0), body])
    : body;
  return { rgce, extra: encoder.extra.toUint8Array() };
}

/** A growable little-endian byte buffer whose earlier bytes can be patched — a jump offset is known only later. */
class TokenBuffer {
  private bytes = new Uint8Array(64);
  private size = 0;

  get length(): number {
    return this.size;
  }

  private reserve(count: number): void {
    if (this.size + count > this.bytes.length) {
      const grown = new Uint8Array(Math.max(this.bytes.length * 2, this.size + count));
      grown.set(this.bytes.subarray(0, this.size));
      this.bytes = grown;
    }
  }

  u8(value: number): this {
    this.reserve(1);
    this.bytes[this.size++] = value & 0xff;
    return this;
  }

  u16(value: number): this {
    return this.u8(value).u8(value >>> 8);
  }

  u32(value: number): this {
    return this.u16(value & 0xffff).u16(value >>> 16);
  }

  f64(value: number): this {
    const chunk = new Uint8Array(8);
    new DataView(chunk.buffer).setFloat64(0, value, true);
    return this.write(chunk);
  }

  write(chunk: Uint8Array): this {
    this.reserve(chunk.length);
    this.bytes.set(chunk, this.size);
    this.size += chunk.length;
    return this;
  }

  patchU16(at: number, value: number): void {
    this.bytes[at] = value & 0xff;
    this.bytes[at + 1] = (value >>> 8) & 0xff;
  }

  /** Remove and return everything from `at` on. */
  cut(at: number): Uint8Array {
    const tail = this.bytes.slice(at, this.size);
    this.size = at;
    return tail;
  }

  toUint8Array(): Uint8Array {
    return this.bytes.slice(0, this.size);
  }
}

interface Encoder {
  readonly out: TokenBuffer;
  readonly extra: TokenBuffer;
  readonly classes: TokenClasses;
  /** The `_xlfn.SINGLE` calls written as their operand alone — see `bareIntersections`. */
  readonly bare: ReadonlySet<AstNode>;
  readonly layout: FormulaLayout;
  readonly context: PtgContext;
  readonly where: string;
  volatile: boolean;
  /** Where the operand that ends a call's last argument finished — see {@link emitClosingArgument}. */
  mark?: { readonly node: AstNode; at: number };
}

/** Emit an invoked call — see {@link INVOCATION_CLASSES} for the shape. */
function emitInvocation(
  encoder: Encoder,
  node: FunctionCallNode,
  entry: NodeLayout | undefined
): void {
  const operands = invocationOperands(node);
  if (node.name !== INVOKE && entry !== undefined && !encoder.layout.has(operands[0]!)) {
    // The whitespace written around the call belongs to the call the lists are applied to.
    encoder.layout.set(operands[0]!, entry);
  }
  for (const operand of operands) {
    emit(encoder, operand, node);
  }
  encoder.out
    .u8(Ptg.FuncVar | classOf(encoder, node))
    .u8(operands.length)
    .u16(IFTAB_CALL_BY_NAME);
}

// -----------------------------------------------------------------------------
// Emission
// -----------------------------------------------------------------------------

/** Emit `node`, then its mem token's length, its own parenthesis if it is a union, then its display parentheses. */
function emit(encoder: Encoder, node: AstNode, parent: AstNode | undefined): void {
  const { out } = encoder;
  const entry = encoder.layout.get(node);
  const memClass = encoder.classes.mem.get(node);
  if (encoder.bare.has(node) && node.type === NodeType.FunctionCall) {
    // `@x` written as `x` in its place, which intersects it there.
    emitSpaces(out, SpaceAt.Expression, entry?.lead);
    emit(encoder, node.args[0]!, parent);
  } else if (memClass !== undefined && needsMemToken(node, parent)) {
    out.u8(Ptg.MemFunc | memClass);
    const at = out.length;
    out.u16(0);
    emitCore(encoder, node, entry);
    out.patchU16(at, out.length - at - 2);
  } else {
    emitCore(encoder, node, entry);
  }
  const bareRoot =
    parent === undefined &&
    (entry?.bare === true || encoder.context.role === "name" || encoder.context.role === "list");
  if (node.type === NodeType.UnionRef && !bareRoot) {
    // A union's own parentheses are a `PtgParen` after the mem expression, as Excel writes them — except as a
    // whole defined name, which Excel stores without them: `S!$A$1,S!$B$2` is `29 13 00 3a … 3a … 10`.
    emitSpaces(out, SpaceAt.OpenParen, entry?.lead);
    emitSpaces(out, SpaceAt.CloseParen, entry?.close);
    out.u8(Ptg.Paren);
  }
  for (const paren of entry?.parens ?? []) {
    emitSpaces(out, SpaceAt.OpenParen, paren.open);
    emitSpaces(out, SpaceAt.CloseParen, paren.close);
    out.u8(Ptg.Paren);
  }
  // A child that precedence needs parenthesised but whose layout has none — a tree built rather than parsed —
  // still gets its parenthesis, so Excel displays what it computes.
  if (
    parent !== undefined &&
    (entry?.parens?.length ?? 0) === 0 &&
    needsParentheses(node, parent)
  ) {
    out.u8(Ptg.Paren);
  }
  if (encoder.mark?.node === node) {
    encoder.mark.at = out.length;
  }
}

/**
 * `PtgAttrSpace` tokens for a run of whitespace at `at`: one per run of spaces or of line breaks, each at most 255.
 * A `\r\n` pair is one line break, and a tab is treated as a space.
 */
function emitSpaces(out: TokenBuffer, at: number, text: string | undefined): void {
  if (text === undefined || text === "") {
    return;
  }
  const normalised = text.replaceAll("\r\n", "\n").replaceAll("\r", "\n").replaceAll("\t", " ");
  for (const run of normalised.match(/ +|\n+/g) ?? []) {
    const type = at + (run[0] === "\n" ? 1 : 0);
    for (let left = run.length; left > 0; left -= 255) {
      out.u8(Ptg.Attr).u8(Attr.Space).u8(type).u8(Math.min(left, 255));
    }
  }
}

function emitCore(encoder: Encoder, node: AstNode, entry: NodeLayout | undefined): void {
  const { out, where } = encoder;
  const lead = entry?.lead;
  switch (node.type) {
    case NodeType.Number:
      emitSpaces(out, SpaceAt.Expression, lead);
      emitNumber(out, node.value);
      return;
    case NodeType.String: {
      if (node.value.length > 0xffff) {
        unsupported(where, "a string literal longer than 65535 characters");
      }
      emitSpaces(out, SpaceAt.Expression, lead);
      out.u8(Ptg.Str).u16(node.value.length);
      for (let index = 0; index < node.value.length; index++) {
        out.u16(node.value.charCodeAt(index));
      }
      return;
    }
    case NodeType.Boolean:
      emitSpaces(out, SpaceAt.Expression, lead);
      out.u8(Ptg.Bool).u8(node.value ? 1 : 0);
      return;
    case NodeType.Error: {
      emitSpaces(out, SpaceAt.Expression, lead);
      if (node.value === "#REF!") {
        // Excel parses a typed `#REF!` as a reference error rather than as an error constant: `=#REF!` is
        // `4a 00 00 00 00 00 00` and `=Other!#REF!` is `5c 03 00 …`, the 3D form naming the sheet.
        const tokenClass = classOf(encoder, node);
        if (node.sheet !== undefined) {
          out.u8(Ptg.RefErr3d | tokenClass).u16(sheetIndex(node.sheet, encoder.context, where));
        } else {
          out.u8(Ptg.RefErr | tokenClass);
        }
        out.u32(0).u16(0);
        return;
      }
      const code = errorCodeOf(node.value);
      if (code === undefined) {
        unsupported(where, `the error value ${node.value}`);
      }
      out.u8(Ptg.Err).u8(code!);
      return;
    }
    case NodeType.Missing:
      emitSpaces(out, SpaceAt.Expression, lead);
      out.u8(Ptg.MissArg);
      return;
    case NodeType.BinaryOp: {
      const ptg = PTG_BY_BINARY_OP.get(node.op);
      if (ptg === undefined) {
        unsupported(where, `the operator "${node.op}"`);
      }
      emit(encoder, node.left, node);
      emit(encoder, node.right, node);
      // The intersection operator is its first space; anything after it is recorded whitespace.
      emitSpaces(out, SpaceAt.Expression, node.op === " " ? entry?.op?.slice(1) : entry?.op);
      out.u8(ptg!);
      return;
    }
    case NodeType.UnaryOp:
      // `A1#` is not among them: it is stored as the `_xlfn.ANCHORARRAY` call, which `encodeFormula` and `encodePtg`
      // write it as (`operatorsAsCalls`).
      // Nor is `@`: it is the `_xlfn.SINGLE` call, which `encodeFormulaAst` makes it (`operatorsAsCalls`).
      if (node.op !== "-" && node.op !== "+") {
        unsupported(where, `the operator "${node.op}"`);
      }
      if (
        node.op === "-" &&
        node.operand.type === NodeType.Number &&
        encoder.layout.get(node.operand) === undefined
      ) {
        // **A negated literal is one constant**, as Excel stores it: `=-5` is `1f 00 00 00 00 00 00 14 c0`, not
        // `5` and `PtgUminus`. Both compute the same; this is the one Excel writes and reads back as `-5`.
        emitSpaces(out, SpaceAt.Expression, lead);
        out.u8(Ptg.Num).f64(-node.operand.value);
        return;
      }
      emit(encoder, node.operand, node);
      emitSpaces(out, SpaceAt.Expression, lead);
      out.u8(node.op === "-" ? Ptg.Uminus : Ptg.Uplus);
      return;
    case NodeType.Percent:
      emit(encoder, node.operand, node);
      emitSpaces(out, SpaceAt.Expression, entry?.op);
      out.u8(Ptg.Percent);
      return;
    case NodeType.CellRef:
      emitSpaces(out, SpaceAt.Expression, lead);
      emitCellRef(encoder, node);
      return;
    case NodeType.RangeRef: {
      emitSpaces(out, SpaceAt.Expression, lead);
      const start = decodeCellRef(node.start, where);
      const end = decodeCellRef(node.end, where);
      emitArea(encoder, node, start, end, node.sheet, node.endSheet);
      return;
    }
    // A whole-row or whole-column reference is an ordinary `PtgArea` with the *other* axis pinned to the
    // sheet's limit — that is what `A:A` and `$1:$1` are in a token stream, and Excel writes them that way.
    case NodeType.ColRangeRef:
    case NodeType.RowRangeRef: {
      emitSpaces(out, SpaceAt.Expression, lead);
      const whole = wholeAxisArea(node);
      emitArea(encoder, node, whole.start, whole.end, node.sheet, node.endSheet);
      return;
    }
    case NodeType.FunctionCall:
      if (isInvocation(node)) {
        emitInvocation(encoder, node, entry);
        return;
      }
      emitCall(encoder, node, entry);
      return;
    case NodeType.Name:
      emitSpaces(out, SpaceAt.Expression, lead);
      emitName(encoder, node);
      return;
    case NodeType.UnionRef:
      // The members, then one union operator per join — the shape the decoder flattens. The union's own
      // parenthesis is emitted by `emit`, outside the mem token's expression.
      node.areas.forEach((area, index) => {
        emit(encoder, area, node);
        if (index > 0) {
          out.u8(Ptg.Union);
        }
      });
      return;
    case NodeType.StructuredRef:
      emitSpaces(out, SpaceAt.Expression, lead);
      emitStructuredRef(node, out, encoder.context, where, classOf(encoder, node));
      return;
    case NodeType.Array:
      emitSpaces(out, SpaceAt.Expression, lead);
      emitArray(encoder, node);
      return;
  }
}

/** The compact integer form where it fits, exactly as Excel does. */
function emitNumber(out: TokenBuffer, value: number): void {
  if (Number.isInteger(value) && value >= 0 && value <= 0xffff && !Object.is(value, -0)) {
    out.u8(Ptg.Int).u16(value);
  } else {
    out.u8(Ptg.Num).f64(value);
  }
}

/**
 * Whether `child` needs parentheses to keep its meaning as an operand of `parent` — the same rule the printer
 * applies, so a tree built rather than parsed is displayed by Excel as it computes.
 */
function needsParentheses(child: AstNode, parent: AstNode): boolean {
  const power = bindingPower(child);
  if (power === undefined) {
    return false;
  }
  switch (parent.type) {
    case NodeType.BinaryOp: {
      const [left, right] = infixPower(parent.op);
      return parent.left === child ? power < left : power < right;
    }
    case NodeType.UnaryOp:
      return parent.op !== "@" && power < PREFIX_POWER;
    case NodeType.Percent:
      return power < PERCENT_POWER;
    default:
      return false;
  }
}

const PREFIX_POWER = 77;
const PERCENT_POWER = 75;

/** Mirrors `@formula/syntax/print`. */
function infixPower(op: string): readonly [number, number] {
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
      return [60, 61];
    case " ":
      return [80, 81];
    case ":":
      return [90, 91];
    default:
      return [0, 0];
  }
}

function bindingPower(node: AstNode): number | undefined {
  switch (node.type) {
    case NodeType.BinaryOp:
      return infixPower(node.op)[0];
    case NodeType.UnaryOp:
      return node.op === "@" ? undefined : PREFIX_POWER;
    case NodeType.Percent:
      return PERCENT_POWER;
    default:
      return undefined;
  }
}

function classOf(encoder: Encoder, node: AstNode): number {
  return encoder.classes.own.get(node) ?? REFERENCE_CLASS;
}

/**
 * A function call, in the form Excel writes it.
 *
 * `IF` and `CHOOSE` carry jump tables, `IFERROR` a jump, and a one-argument `SUM` is `PtgAttrSum`. The jumps are
 * byte offsets, measured on Excel's own output: `PtgAttrIf` points past the first `PtgAttrGoto` (the start of the
 * false branch), `PtgAttrIfError` at its `PtgAttrGoto`, a `PtgAttrChoose` entry at its branch counted from the start
 * of the offset table, and every `PtgAttrGoto` at one byte short of the end of the function token. The whitespace
 * before the function's `)` goes where Excel puts it — see {@link emitClosingArgument}.
 */
function emitCall(encoder: Encoder, node: FunctionCallNode, entry: NodeLayout | undefined): void {
  const { out, where } = encoder;
  const id = builtInId(node.name);
  const closing = (): void => {
    emitSpaces(out, SpaceAt.Expression, entry?.lead);
    emitSpaces(out, SpaceAt.CloseParen, entry?.close);
  };
  /** Emit argument `index`; the last one carries the whitespace before `)` — see {@link emitClosingArgument}. */
  const argument = (index: number): void => {
    if (index === node.args.length - 1) {
      emitClosingArgument(encoder, node.args[index]!, node, closing);
    } else {
      emit(encoder, node.args[index]!, node);
    }
  };
  // With no argument, the whitespace can only go in front of the function token.
  const bare = node.args.length === 0 ? closing : (): void => {};
  if (id === undefined) {
    emitFutureFunctionCall(encoder, node, argument, bare);
    return;
  }
  bare();
  if (functionClasses(id).volatile) {
    encoder.volatile = true;
  }
  const functionToken = (): void => {
    const arity = ARITY_BY_ID.get(id);
    const tokenClass = classOf(encoder, node);
    if (arity === undefined) {
      out
        .u8(Ptg.FuncVar | tokenClass)
        .u8(node.args.length)
        .u16(id);
      return;
    }
    if (node.args.length !== arity) {
      // The count is implied by the token, so a call with the wrong number of arguments cannot be *expressed*.
      unsupported(
        where,
        `a call to ${node.name} with ${node.args.length} argument(s), which takes exactly ${arity}`
      );
    }
    out.u8(Ptg.Func | tokenClass).u16(id);
  };
  const gotos: number[] = [];
  const jump = (): void => {
    out.u8(Ptg.Attr).u8(Attr.Goto);
    gotos.push(out.length);
    out.u16(0);
  };
  const finish = (): void => {
    functionToken();
    for (const at of gotos) {
      out.patchU16(at, out.length - (at + 2) - 1);
    }
  };

  if (id === IF_ID && (node.args.length === 2 || node.args.length === 3)) {
    argument(0);
    out.u8(Ptg.Attr).u8(Attr.If);
    const ifAt = out.length;
    out.u16(0);
    argument(1);
    jump();
    out.patchU16(ifAt, out.length - (ifAt + 2));
    if (node.args.length === 3) {
      argument(2);
      jump();
    }
    finish();
    return;
  }
  if (id === CHOOSE_ID && node.args.length >= 2) {
    argument(0);
    const choices = node.args.length - 1;
    out.u8(Ptg.Attr).u8(Attr.Choose).u16(choices);
    const table = out.length;
    for (let index = 0; index <= choices; index++) {
      out.u16(0);
    }
    for (let index = 1; index <= choices; index++) {
      out.patchU16(table + (index - 1) * 2, out.length - table);
      argument(index);
      jump();
    }
    out.patchU16(table + choices * 2, out.length - table);
    finish();
    return;
  }
  if (id === IFERROR_ID && node.args.length === 2) {
    argument(0);
    out.u8(Ptg.Attr).u8(Attr.IfError);
    const at = out.length;
    out.u16(0);
    argument(1);
    out.patchU16(at, out.length - (at + 2));
    jump();
    finish();
    return;
  }
  node.args.forEach((_argument, index) => argument(index));
  if (isAttrSum(node)) {
    out.u8(Ptg.Attr).u8(Attr.Sum).u16(0);
    return;
  }
  finish();
}

/**
 * Emit a call's last argument with the whitespace before the call's `)` where Excel puts it: **after the last
 * operand, before the operators that complete the argument.**
 *
 * Excel records that whitespace when it reaches the `)`, which is before it has emitted the operators still pending
 * on the right of the argument — so `AND( A1>0, B1>0 )` is `… 44 B1 1e 00 00 19 40 04 01 0d 42 02 24 00`, the
 * space ahead of the `>`, while `SUM( A1 )` has it after `A1` and `IF( A1>0,1,2 )` ahead of the final jump. Reading
 * either placement gives the same text; this one makes the bytes Excel's.
 */
function emitClosingArgument(
  encoder: Encoder,
  argument: AstNode,
  parent: AstNode,
  closing: () => void
): void {
  const previous = encoder.mark;
  const mark = { node: rightmostOperand(encoder, argument), at: -1 };
  encoder.mark = mark;
  emit(encoder, argument, parent);
  encoder.mark = previous;
  const tail = encoder.out.cut(mark.at < 0 ? encoder.out.length : mark.at);
  closing();
  encoder.out.write(tail);
}

/**
 * The node whose tokens end an expression before its trailing operators: down the right-hand side through every
 * operator that is emitted after its operands, stopping at a leaf, a call, a parenthesised group or a mem-wrapped
 * reference expression.
 */
function rightmostOperand(encoder: Encoder, node: AstNode): AstNode {
  let current = node;
  for (;;) {
    if (
      (encoder.layout.get(current)?.parens?.length ?? 0) > 0 ||
      encoder.classes.mem.has(current)
    ) {
      return current;
    }
    if (current.type === NodeType.BinaryOp) {
      current = current.right;
    } else if (current.type === NodeType.Percent) {
      current = current.operand;
    } else if (
      current.type === NodeType.UnaryOp &&
      !(
        current.op === "-" &&
        current.operand.type === NodeType.Number &&
        !encoder.layout.has(current.operand)
      )
    ) {
      current = current.operand;
    } else {
      return current;
    }
  }
}

/** `PtgArray` and its `PtgExtraArray` — the inverse of `readArrayConstant`. */
function emitArray(encoder: Encoder, node: Extract<AstNode, { type: NodeType.Array }>): void {
  const { out, extra, where } = encoder;
  const rows = node.rows;
  const columns = rows[0]?.length ?? 0;
  if (rows.length === 0 || columns === 0 || rows.some(row => row.length !== columns)) {
    unsupported(where, "an array constant whose rows differ in length");
  }
  // The fourteen bytes are unused. Excel's own carry the column and row counts less one in the first four, which
  // is written the same way so the token matches Excel's byte for byte where it matters to nobody but a diff.
  out
    .u8(Ptg.Array | classOf(encoder, node))
    .u16(columns - 1)
    .u16(rows.length - 1);
  for (let index = 0; index < 10; index++) {
    out.u8(0);
  }
  extra.u32(rows.length).u32(columns);
  for (const row of rows) {
    for (const value of row) {
      const literal =
        value.type === NodeType.UnaryOp &&
        (value.op === "-" || value.op === "+") &&
        value.operand.type === NodeType.Number
          ? {
              type: NodeType.Number as const,
              value: value.op === "-" ? -value.operand.value : value.operand.value
            }
          : value;
      switch (literal.type) {
        case NodeType.Number:
          extra.u8(0x00).f64(literal.value);
          break;
        case NodeType.String:
          if (literal.value.length > 255) {
            unsupported(where, "an array constant string longer than 255 characters");
          }
          extra.u8(0x01).u16(literal.value.length);
          for (let index = 0; index < literal.value.length; index++) {
            extra.u16(literal.value.charCodeAt(index));
          }
          break;
        case NodeType.Boolean:
          extra.u8(0x02).u8(literal.value ? 1 : 0);
          break;
        case NodeType.Error: {
          const code = errorCodeOf(literal.value);
          if (code === undefined) {
            unsupported(where, `the error value ${literal.value} in an array constant`);
          }
          extra.u8(0x04).u8(code!).u8(0).u16(0);
          break;
        }
        default:
          unsupported(where, "an array constant element that is not a literal");
      }
    }
  }
}

/**
 * A defined name: `PtgName` with its one-based index, or `PtgNameX` for a name scoped to another sheet.
 *
 * **Scope decides which name.** A sheet may define its own `Total` beside the workbook's, and an unqualified
 * `Total` on that sheet means the sheet's; elsewhere it means the workbook's. `Other!Total` means Other's, and
 * Excel writes it as `PtgNameX` through a workbook-level `BrtExternSheet` entry (`BrtSupSelf`, -2, -2) — the same
 * token a formula reaching another workbook's name uses.
 */
function emitName(encoder: Encoder, node: Extract<AstNode, { type: NodeType.Name }>): void {
  const { out, context, where } = encoder;
  const index = indexOf(context);
  const candidates = index.names.get(definedNameListKey(node.name)) ?? [];
  if (node.sheet !== undefined) {
    const sheet = index.sheets.get(sheetNameKey(node.sheet));
    if (sheet === undefined) {
      throw new XlsbParseError(where, `sheet ${node.sheet} is not in the workbook`);
    }
    const local =
      candidates.find(candidate => candidate.scope === sheet) ??
      candidates.find(candidate => candidate.scope === undefined);
    if (local === undefined) {
      unsupported(where, `the name ${node.sheet}!${node.name}, which the workbook does not define`);
    }
    out
      .u8(Ptg.NameX | classOf(encoder, node))
      .u16(internExternSheet(context, -2, -2, where))
      .u32(local!.index + 1);
    return;
  }
  const eta = etaStubName(node.name, candidates.length > 0);
  if (eta !== undefined) {
    // Always the value class, whatever the argument it fills: Excel's `GROUPBY(…,_xleta.SUM)`, `BYROW(…,_xleta.SUM)`
    // and `MAP(…,_xleta.ABS)` all carry `43`, where a reference argument of the same function is `25`.
    const stub = index.names.get(definedNameListKey(eta))?.[0];
    if (stub === undefined) {
      unsupported(
        where,
        `the function ${node.name} as an argument, which needs a ${eta} defined name`
      );
    }
    out.u8(Ptg.Name | VALUE_CLASS).u32(stub!.index + 1);
    return;
  }
  const chosen =
    candidates.find(
      candidate => candidate.scope !== undefined && candidate.scope === context.sheetIndex
    ) ??
    candidates.find(candidate => candidate.scope === undefined) ??
    candidates[0];
  if (chosen === undefined) {
    unsupported(where, `the name ${node.name}, which the workbook does not define`);
  }
  out.u8(Ptg.Name | classOf(encoder, node)).u32(chosen!.index + 1);
}

/** A row and a `ColRelShort` — one corner of an area, as the token stores it. */
interface AreaCorner {
  readonly row: number;
  readonly columnAndFlags: number;
}

/**
 * A whole-row or whole-column reference as `PtgArea` corners.
 *
 * **The pinned axis runs the sheet's full extent and is absolute**: `A:A` means *every* row, so its row bounds
 * must not shift when the formula is copied. **The named axis keeps the reference's own `$` markers.** Bits set
 * mean *relative*, the opposite of how `$` reads.
 */
function wholeAxisArea(node: ColRangeRefNode | RowRangeRefNode): {
  readonly start: AreaCorner;
  readonly end: AreaCorner;
} {
  if (node.type === NodeType.RowRangeRef) {
    const relativeRow = (absolute: boolean | undefined): number => (absolute ? 0 : 0x8000);
    return {
      start: { row: node.startRow - 1, columnAndFlags: relativeRow(node.startAbsolute) },
      end: {
        row: node.endRow - 1,
        columnAndFlags: relativeRow(node.endAbsolute) | (XLSB_MAX_COLUMNS - 1)
      }
    };
  }
  const relativeColumn = (absolute: boolean | undefined): number => (absolute ? 0 : 0x4000);
  return {
    start: {
      row: 0,
      columnAndFlags: relativeColumn(node.startAbsolute) | (decodeCol(node.startCol) & 0x3fff)
    },
    end: {
      row: XLSB_MAX_ROWS - 1,
      columnAndFlags: relativeColumn(node.endAbsolute) | (decodeCol(node.endCol) & 0x3fff)
    }
  };
}

/**
 * `PtgArea`, `PtgArea3d` or `PtgAreaN` for two corners whose flags say which axes are relative.
 *
 * **A range in a template is `PtgAreaN` only when something in it is relative.** A shared `=SUM(B2:D2)` filled
 * downwards has to name its own row from every row, so its relative axes are offsets from the origin. A fully
 * absolute one is an ordinary `PtgArea`, as Excel writes it: a conditional format's `MAX($A$1:$A$10)` is `25 …`
 * and a validation list `A:A` — relative columns — is `2d …`.
 */
function emitArea(
  encoder: Encoder,
  node: AstNode,
  start: AreaCorner,
  end: AreaCorner,
  sheet: string | undefined,
  endSheet: string | undefined
): void {
  const { out, context, where } = encoder;
  const tokenClass = classOf(encoder, node);
  if (sheet !== undefined) {
    // **No `PtgArea3d` in a template.** `SharedParsedFormula` and `CFParsedFormula` both forbid it (MS-XLSB
    // 2.5.98.98 and 2.5.98.6), and its `RgceArea` holds positions, so a relative corner could not shift anyway.
    if (context.relativeToOrigin === true) {
      unsupported(where, `the reference to another sheet ${sheet}!…, which a template cannot hold`);
    }
    out.u8(Ptg.Area3d | tokenClass).u16(sheetIndex(sheet, context, where, endSheet));
    out.u32(start.row).u32(end.row).u16(start.columnAndFlags).u16(end.columnAndFlags);
    return;
  }
  const relative = ((start.columnAndFlags | end.columnAndFlags) & 0xc000) !== 0;
  if (context.relativeToOrigin === true && context.origin !== undefined && relative) {
    out.u8(Ptg.AreaN | tokenClass);
    out
      .u32(relativeRowOffset(start, context.origin))
      .u32(relativeRowOffset(end, context.origin))
      .u16(relativeColumnOffset(start.columnAndFlags, context.origin))
      .u16(relativeColumnOffset(end.columnAndFlags, context.origin));
    return;
  }
  out.u8(Ptg.Area | tokenClass);
  out.u32(start.row).u32(end.row).u16(start.columnAndFlags).u16(end.columnAndFlags);
}

function emitCellRef(encoder: Encoder, node: CellRefNode): void {
  const { out, context, where } = encoder;
  const corner = decodeCellRef(node, where);
  const { row, columnAndFlags } = corner;
  const tokenClass = classOf(encoder, node);
  if (node.sheet !== undefined) {
    // Forbidden in a template for the same reason as `PtgArea3d` — see `emitArea`.
    if (context.relativeToOrigin === true) {
      unsupported(
        where,
        `the reference to another sheet ${node.sheet}!…, which a template cannot hold`
      );
    }
    out.u8(Ptg.Ref3d | tokenClass).u16(sheetIndex(node.sheet, context, where, node.endSheet));
    out.u32(row).u16(columnAndFlags);
    return;
  }
  // **`PtgRefN` for a relative reference in a template.** A conditional-formatting rule's formula is written once
  // and applied to every cell in its range, so its relative references are *offsets from the range's anchor*.
  // A fully absolute one stays `PtgRef`, as Excel writes it — a validation bound `$B$1+1` is `44 00 00 00 00 01 00`.
  if (
    context.relativeToOrigin === true &&
    context.origin !== undefined &&
    (columnAndFlags & 0xc000) !== 0
  ) {
    // **Offsets wrap at the width of the field they name.** A row offset of -1 is `0xFFFFF` and a column offset of
    // -1 is `0x3FFF`, because that is what Excel writes: `poi-62815.xlsb` stores `=<the cell above> + 1` as
    // `4c ff ff 0f 00 00 c0 …`.
    out
      .u8(Ptg.RefN | tokenClass)
      .u32(relativeRowOffset(corner, context.origin))
      .u16(relativeColumnOffset(columnAndFlags, context.origin));
    return;
  }
  out
    .u8(Ptg.Ref | tokenClass)
    .u32(row)
    .u16(columnAndFlags);
}

function decodeCellRef(node: CellRefNode, where: string): { row: number; columnAndFlags: number } {
  const row = Number(node.row) - 1;
  const column = decodeCol(node.col);
  if (!Number.isInteger(row) || row < 0 || column < 0) {
    throw new XlsbParseError(where, `cannot encode the reference ${node.col}${node.row}`);
  }
  // Set means relative, so an absolute reference clears the bit.
  const flags = (node.rowAbsolute ? 0 : 0x8000) | (node.colAbsolute ? 0 : 0x4000);
  return { row, columnAndFlags: flags | (column & 0x3fff) };
}

/**
 * `PtgList` — MS-XLSB 2.5.98.52. Twelve bytes.
 *
 * ```text
 * ptg          u8    0x18, then a mandatory 0x19 `eptg`
 * eptg         u8
 * ixti         u16
 * flags        u16   columns (2 bits), rowType (5), then seven more
 * listIndex    u32   the table's own `idList`
 * colFirst     u16   zero-based, *relative to the table*
 * colLast      u16
 * ```
 *
 * **The column indices are relative to the table, not to the sheet.** `rowType` is what `[#Data]`, `[#Headers]`,
 * `[#Totals]` and `[#All]` become; a bare `Table1[Amount]` means the data region, which is `DATA` — 0.
 *
 * Two columns in a `StructuredRefNode` are a span, `Table1[[Qty]:[Note]]`: `columns` 2, from `colFirst` to `colLast`.
 */
function emitStructuredRef(
  node: StructuredRefNode,
  out: TokenBuffer,
  context: PtgContext,
  where: string,
  tokenClass: number
): void {
  const table = indexOf(context).tables.get(excelNameKey(node.tableName));
  if (table === undefined) {
    // A reference to a table this workbook does not define. Refused by name rather than written with a
    // guessed id: an `idList` matching no `BrtBeginList` is a reference Excel reports as broken. One with no name at
    // all is an unnamed `[@x]` that reached here without the enclosing table being supplied.
    unsupported(
      where,
      node.tableName === ""
        ? "a structured reference with no table name"
        : `a structured reference to the unknown table ${node.tableName}`
    );
    return;
  }
  const rowType = rowTypeOf(node.specials);
  // `columns`: 0 for the whole table, 1 for one column, 2 for a span.
  // Columns, like tables, are compared by `excelNameKey`.
  const columnKeys = table.columns.map(excelNameKey);
  const first = node.columns.length === 0 ? -1 : columnKeys.indexOf(excelNameKey(node.columns[0]));
  const last =
    node.columns.length < 2
      ? first
      : columnKeys.indexOf(excelNameKey(node.columns[node.columns.length - 1]));
  if (node.columns.length > 0 && (first < 0 || last < 0)) {
    unsupported(where, `a structured reference to a column ${node.tableName} does not have`);
    return;
  }
  const columns = node.columns.length === 0 ? 0 : node.columns.length === 1 ? 1 : 2;
  let flags = columns & 0x03;
  flags |= (rowType & 0x1f) << 2;
  // `type` (bits 10–11) is the token's operand class — 0 reference, 1 value, 2 array — carried here because the
  // `ptg` byte has no room for it. Excel's `=Tbl[[#Totals],[Qty]]` is `21 04`: one column, `TOTALS`, value; the
  // reference type there made Excel answer `#VALUE!` for the same text.
  flags |= (tokenClass === VALUE_CLASS ? 1 : tokenClass === ARRAY_CLASS ? 2 : 0) << 10;
  // Bit 9 — `unused` in MS-XLSB 2.5.98.52 ("undefined, MUST be ignored") — is set by Excel on every this-row
  // reference it writes, named or typed as `[@x]`: `Tb[[#This Row],[x]]` is `41 06`, where the same reference with
  // the bit clear would be `41 04` (`data/table-this-row.excel.xlsb`, in the cells and in `BrtListCCFmla`). It sits
  // beside `rowType`'s `#This Row`, and the reader ignores it as the specification says.
  if (rowType === 0x10) {
    flags |= 0x0200;
  }
  // **The byte is exactly 0x18 — `PtgList` carries no operand class.** Its `ptg` field is *seven* bits wide, with
  // the eighth a reserved flag that MUST be 0. Its class is the `type` field inside `flags`, set above.
  out.u8(Ptg.List);
  // `eptg`, which the specification fixes at 0x19. It is not a length or a variant — omitting it shifts
  // every field after it.
  out.u8(0x19);
  // `ixti` — the sheet the *table* is on, resolved through `BrtExternSheet` exactly as `PtgRef3d` does.
  out.u16(sheetIndex(table.sheet, context, where));
  out.u16(flags);
  out.u32(table.id);
  out.u16(Math.max(0, first));
  out.u16(Math.max(0, last));
}

/** The inverse of {@link rowTypeOf}: the `[#…]` specifiers a `rowType` stands for. */
function specialsOf(rowType: number): string[] {
  switch (rowType) {
    case 0x01:
      return ["#All"];
    case 0x02:
      return ["#Headers"];
    case 0x04:
      return ["#Data"];
    case 0x06:
      return ["#Headers", "#Data"];
    case 0x08:
      return ["#Totals"];
    case 0x0c:
      return ["#Data", "#Totals"];
    case 0x10:
      return ["#This Row"];
    default:
      // `DATA`: the data region, which is what a bare reference means.
      return [];
  }
}

/** `PtgRowType`, MS-XLSB 2.5.98.73, from the `[#…]` specifiers a reference carries. */
function rowTypeOf(specials: readonly string[]): number {
  const has = (name: string): boolean =>
    specials.some(entry => entry.toLowerCase() === name.toLowerCase());
  if (has("#All")) {
    return 0x01;
  }
  if (has("#This Row") || has("@")) {
    return 0x10;
  }
  const headers = has("#Headers");
  const totals = has("#Totals");
  const data = has("#Data");
  if (headers && data) {
    return 0x06;
  }
  if (data && totals) {
    return 0x0c;
  }
  if (headers) {
    return 0x02;
  }
  if (totals) {
    return 0x08;
  }
  // An explicit `[#Data]` is `DATA2`: Excel writes `Tb[[#Data],[x]]` as `11 00` — one column, rowType 4 — and a bare
  // `Tb[x]` as `01 00`. Both mean the data region; the specifier is what Excel shows.
  if (data) {
    return 0x04;
  }
  // No specifier means the data region, which is `DATA` — and `DATA` is 0.
  return 0x00;
}

/** A table as a formula context knows it. */
interface TableEntry {
  readonly id: number;
  readonly columns: readonly string[];
  readonly sheet: string;
}

/**
 * Lookup tables for a context, built once and cached on it.
 *
 * Keyed by identity, so a caller that mutates a context between calls gets a fresh one by handing over a fresh
 * object — which is what `writeXlsbPackage` does.
 */
interface ContextIndex {
  /** `definedNameListKey` → every defined name of that key, with its position and scope. */
  readonly names: Map<string, { readonly index: number; readonly scope: number | undefined }[]>;
  /** `sheetNameKey` → position. */
  readonly sheets: Map<string, number>;
  /** `excelNameKey` → table. */
  readonly tables: Map<string, TableEntry>;
  readonly externBySheet: Map<number, number>;
  /**
   * Whether anything encoded through this context resolved an `ixti` against the identity mapping a context with
   * no table implies. With a table, the table itself is marked — see {@link markUsed}.
   */
  usedExternSheet: boolean;
}

/** The `Ftab` id reserved for a function called through a `PtgName` — see {@link emitFutureFunctionCall}. */
const IFTAB_CALL_BY_NAME = 0x00ff;

/**
 * `#NAME?` as an `Rgce` — the body Excel gives a future function's stub.
 *
 * `PtgErr` (`0x1c`) then the `BErr` byte `0x1D`. Verified against Excel's own output, whose `_xlfn.CONCAT` and
 * `_xlfn.SINGLE` both carry exactly `1c 1d` with a `cce` of 2.
 */
export const FUTURE_FUNCTION_STUB_RGCE = Uint8Array.of(0x1c, 0x1d);

/** The widest row an XLSB addresses, and therefore the modulus a negative row offset wraps through. */
const ROW_MASK = 0xfffff;

/**
 * A corner's row as an offset from the template's origin, unless the reference pinned it with `$`.
 */
function relativeRowOffset(
  corner: AreaCorner,
  origin: { readonly row: number; readonly column: number }
): number {
  return (corner.columnAndFlags & 0x8000) === 0 ? corner.row : (corner.row - origin.row) & ROW_MASK;
}

/** A column as an offset from the template's origin, preserving the two flags in the high bits. */
function relativeColumnOffset(
  columnAndFlags: number,
  origin: { readonly row: number; readonly column: number }
): number {
  const flags = columnAndFlags & 0xc000;
  const column = columnAndFlags & 0x3fff;
  const offset = (columnAndFlags & 0x4000) === 0 ? column : column - origin.column;
  return (flags | (offset & 0x3fff)) & 0xffff;
}

/**
 * `BrtName.flags` for a future function's stub.
 *
 * `fHidden | fFunc | fProc | fFutureFunction`, and MS-XLSB 2.4.674 makes that combination mandatory rather
 * than stylistic. Excel's own stubs are `0x0002000b`, which is this value.
 */
export const FUTURE_FUNCTION_FLAGS = 0x0002000b;

/**
 * `BrtName.flags` for a `LET`/`LAMBDA` parameter, `_xlpm.x`, without `fCalcExp`: `fHidden` (bit 0), `fProc` (bit 3) and
 * `fFutureFunction` (bit 17), with the same `#NAME?` body and trailing strings as a future function's stub.
 *
 * **Copied from Excel, not derived from the specification** — which this combination contradicts: MS-XLSB 2.4.718
 * says `fFutureFunction` MUST be 0 when `fFunc` is 0 (it is here) or `fCalcExp` is 1 (it often is, see
 * {@link NAME_FLAG_CALC_EXP}). Excel writes it anyway, and a parameter written with flags 0 and no body was dropped
 * when Excel re-saved the file, every `PtgName` to it becoming index 0.
 */
export const LAMBDA_PARAMETER_FLAGS = 0x00020009;

/**
 * `fCalcExp`, bit 4 — per MS-XLSB "whether formula.rgce contains a call to a function that can return an array".
 *
 * That meaning does not fit where Excel sets it: on stubs whose body is only `#NAME?`. Excel sets it on the stubs of
 * `ISOMITTED`, `LAMBDA`, `MODE.MULT`, `SINGLE` and `SORTBY` and no other future function in its files — not on
 * `SEQUENCE` or `FILTER`, which return arrays — and on a `_xlpm.` parameter by the rule in `parameterDeclarations`.
 * Both are reproduced as observed.
 */
export const NAME_FLAG_CALC_EXP = 0x10;

/** The future functions whose stub Excel writes with {@link NAME_FLAG_CALC_EXP} — see there. */
export const CALC_EXP_FUTURE_FUNCTIONS: ReadonlySet<string> = new Set([
  "_XLFN.ISOMITTED",
  "_XLFN.LAMBDA",
  "_XLFN.MODE.MULT",
  "_XLFN.SINGLE",
  "_XLFN.SORTBY"
]);

/** The prefix Excel gives a future function's stub name, and the second one a worksheet function carries. */
const FUTURE_FUNCTION_PREFIX = "_xlfn.";
const WORKSHEET_FUNCTION_PREFIX = "_xlws.";
/** How XLSX spells a call to a user-defined function — one the workbook names with an empty defined name. */
const USER_FUNCTION_PREFIX = "_xludf.";
/** How XLSX spells a function passed by name, an eta-reduced lambda: `GROUPBY(A2:A9,B2:B9,_xleta.SUM)`. */
const ETA_PREFIX = "_xleta.";

/**
 * The `_xleta.` name a function named as a value stands for — `SUM` in `GROUPBY(A2:A9,B2:B9,SUM)` — or `undefined`
 * when `name` is not one.
 *
 * A name spelled `_xleta.X` is one, as a file stores it. A bare name is one when it is a worksheet function's and the
 * workbook defines no name of that spelling (`defined`): such a formula used to be refused for naming an undefined
 * name, so reading it as the function changes nothing a workbook could already write. Excel's own XLSB calls it
 * through a hidden `_xleta.SUM` name with the bare function name — no `_xlfn.` even for `PERCENTOF` or `ARRAYTOTEXT`.
 */
export function etaStubName(name: string, defined: boolean): string | undefined {
  if (name.toLowerCase().startsWith(ETA_PREFIX)) {
    return ETA_PREFIX + name.slice(ETA_PREFIX.length).toUpperCase();
  }
  if (defined) {
    return undefined;
  }
  return isWorksheetFunctionName(name) ? ETA_PREFIX + bareFunctionName(name) : undefined;
}

/** `BrtName.flags` of an `_xleta.` name — Excel's, `fHidden | fProc | fCalcExp | fFutureFunction`. */
export const ETA_NAME_FLAGS = 0x00020019;

/**
 * Whether `name` is XLSX's spelling of a user-defined function call, `_xludf.NAME`. Such a call goes through a
 * plain defined name `NAME` with an empty definition, which is what Excel writes in XLSB.
 */
export function isUserFunction(name: string): boolean {
  return name.toLowerCase().startsWith(USER_FUNCTION_PREFIX);
}

/**
 * Whether `name` is a function this codec must call through a stub rather than by id.
 *
 * Exported so the workbook writer can create the stubs it will need *before* the formulas are encoded — a
 * `PtgName` is a one-based index into the names list, so a name discovered mid-encoding would have to be
 * appended, and the list has already been counted by then.
 */
export function isFutureFunction(name: string): boolean {
  return builtInId(name) === undefined;
}

/**
 * The stub name for a future function, spelled as Excel spells it: `_xlfn.` — and `_xlws.` for a worksheet
 * function — in lower case, then the name in upper case.
 *
 * The prefix comes from the one list of future functions (`@excel/utils/future-functions`) that the XLSX writer
 * uses too, so `FILTER(…)` gets `_xlfn._xlws.FILTER` whether the formula was written bare, half-prefixed
 * (`_xlfn.FILTER`) or as XLSX stores it. A name that list does not know — a function newer than it, or a
 * user-defined one — keeps whatever prefix it was written with and gets `_xlfn.` if it had none, which is the
 * common case.
 *
 * **The case is normalised deliberately.** A caller may type `=xlookup(…)` in lower case, and `_xlfn.xlookup` is a
 * spelling Excel has never been observed to produce.
 */
export function futureFunctionStubName(name: string): string {
  const known = storedFunctionName(name);
  if (known !== undefined) {
    return known;
  }
  let rest = name;
  let worksheet = false;
  if (rest.toLowerCase().startsWith(FUTURE_FUNCTION_PREFIX)) {
    rest = rest.slice(FUTURE_FUNCTION_PREFIX.length);
  }
  if (rest.toLowerCase().startsWith(WORKSHEET_FUNCTION_PREFIX)) {
    rest = rest.slice(WORKSHEET_FUNCTION_PREFIX.length);
    worksheet = true;
  }
  return `${FUTURE_FUNCTION_PREFIX}${worksheet ? WORKSHEET_FUNCTION_PREFIX : ""}${rest.toUpperCase()}`;
}

/**
 * Functions this writer cannot call in XLSB: `PY` (Python in Excel) and `PQSOURCE` (Power Query).
 *
 * A future function is called through a stub `BrtName`, and Excel refuses the whole package when that stub is
 * `_xlfn.PY` or `_xlfn.PQSOURCE` (`lib-py.xlsb` in the fx21 Excel checks) — the XLSX spelling with the prefix is
 * refused the same way, which is why `@excel/utils/future-functions` leaves both bare. No file Excel itself wrote shows
 * what a working Python or Power Query cell needs in either container, so a call to one is reported rather than
 * written into a file Excel will not open.
 */
const NO_XLSB_FORM: ReadonlySet<string> = new Set(["PY", "PQSOURCE"]);

/** Whether a call to `name` has no XLSB form this writer can produce — see {@link NO_XLSB_FORM}. */
export function hasNoXlsbForm(name: string): boolean {
  return NO_XLSB_FORM.has(bareFunctionName(name));
}

/**
 * Emit a call to a function the `Ftab` cannot name.
 *
 * The shape is Excel's: `23 02 00 00 00 … 42 03 ff 00` for `_xlfn.CONCAT(A1,B1)` — a `PtgName` for the stub, the
 * arguments, then `PtgFuncVar` whose `tab` is `0x00FF` and whose `cparams` **counts the name as one of them**.
 *
 * The stub must already exist. It is refused rather than invented here because the name index has to be stable
 * across the whole workbook, and {@link isFutureFunction} exists so the writer can collect them up front.
 */
function emitFutureFunctionCall(
  encoder: Encoder,
  node: FunctionCallNode,
  argument: (index: number) => void,
  closing: () => void
): void {
  const { out, context, where } = encoder;
  if (!isUserFunction(node.name) && !isLambdaParameter(node.name) && hasNoXlsbForm(node.name)) {
    unsupported(where, `a call to ${bareFunctionName(node.name)}`);
    return;
  }
  if (futureFunctionClasses(node.name).volatile) {
    encoder.volatile = true;
  }
  // A `LET`/`LAMBDA` parameter holding a function — `_xlpm.f(1,2)` — is called through its own name.
  const stub = isUserFunction(node.name)
    ? node.name.slice(USER_FUNCTION_PREFIX.length)
    : isLambdaParameter(node.name)
      ? node.name
      : futureFunctionStubName(node.name);
  const index = indexOf(context).names.get(definedNameListKey(stub))?.[0]?.index ?? -1;
  if (index === -1) {
    unsupported(
      where,
      `the function ${node.name}, which needs a ${stub} defined name this workbook does not carry`
    );
    return;
  }
  out.u8(Ptg.Name | REFERENCE_CLASS).u32(index + 1);
  node.args.forEach((_argument, position) => argument(position));
  closing();
  out
    .u8(Ptg.FuncVar | classOf(encoder, node))
    .u8(node.args.length + 1)
    .u16(IFTAB_CALL_BY_NAME);
}

const CONTEXT_INDEX = new WeakMap<object, ContextIndex>();

function indexOf(context: PtgContext): ContextIndex {
  const cached = CONTEXT_INDEX.get(context);
  if (cached !== undefined) {
    return cached;
  }
  // **Keyed by the rules Excel compares by** (`@utils/name-key`): `MyRange` and `myrange` are one defined name, and
  // `Straße` and `STRASSE` one, while `ı` and `I` are two — which `toUpperCase` made one, writing every `I` as `ı`.
  // A sheet, and a `LET` / `LAMBDA` parameter, is compared by its own, narrower rule.
  const names = new Map<string, { index: number; scope: number | undefined }[]>();
  (context.definedNames ?? []).forEach((name, index) => {
    const key = definedNameListKey(name);
    const entries = names.get(key) ?? [];
    entries.push({ index, scope: context.definedNameScopes?.[index] });
    names.set(key, entries);
  });
  const sheets = new Map<string, number>();
  (context.sheetNames ?? []).forEach((sheet, index) => {
    const key = sheetNameKey(sheet);
    if (!sheets.has(key)) {
      sheets.set(key, index);
    }
  });
  const tables = new Map<string, TableEntry>();
  for (const [name, table] of context.tables ?? []) {
    const key = excelNameKey(name);
    if (!tables.has(key)) {
      tables.set(key, table);
    }
  }
  const externBySheet = new Map<number, number>();
  (context.externSheets ?? []).forEach((entry, index) => {
    if (
      (entry.supportingLink ?? 0) === 0 &&
      entry.first === entry.last &&
      !externBySheet.has(entry.first)
    ) {
      externBySheet.set(entry.first, index);
    }
  });
  const built: ContextIndex = { names, sheets, tables, externBySheet, usedExternSheet: false };
  CONTEXT_INDEX.set(context, built);
  return built;
}

/**
 * The `ixti` for a sheet — an index into `BrtExternSheet`, not into the sheet list.
 *
 * With an explicit table, the entry naming exactly that one sheet is the answer. Without one the
 * identity mapping is assumed, which is only sound because this library's writer emits an identity
 * table alongside.
 */
function sheetIndex(sheet: string, context: PtgContext, where: string, endSheet?: string): number {
  const index = indexOf(context);
  const sheetPosition = index.sheets.get(sheetNameKey(sheet)) ?? -1;
  if (sheetPosition === -1) {
    throw new XlsbParseError(where, `sheet ${sheet} is not in the workbook`);
  }
  if (context.externSheets === undefined) {
    index.usedExternSheet = true;
    return sheetPosition;
  }
  markUsed(context);
  // A reference across a *span* of sheets — `Sheet1:Sheet3!A1` — is an entry whose `itabFirst` and `itabLast`
  // differ.
  if (endSheet !== undefined) {
    const endPosition = index.sheets.get(sheetNameKey(endSheet)) ?? -1;
    if (endPosition === -1) {
      throw new XlsbParseError(where, `sheet ${endSheet} is not in the workbook`);
    }
    return internExternSheet(context, sheetPosition, endPosition, where);
  }
  const ixti = index.externBySheet.get(sheetPosition) ?? -1;
  if (ixti === -1) {
    throw new XlsbParseError(
      where,
      `sheet ${sheet} has no single-sheet BrtExternSheet entry to reference it by`
    );
  }
  return ixti;
}

/**
 * `BrtExternSheet` tables some token resolved an `ixti` through.
 *
 * Marked on the *table*, not on the context: a writer derives one context per sheet and per template from the
 * workbook's (`{ ...context, sheetIndex }`), and every copy shares the one table the workbook part will carry. A
 * flag on the context object was invisible from the original as soon as a copy did the resolving.
 */
const USED_TABLES = new WeakSet<object>();

function markUsed(context: PtgContext): void {
  if (context.externSheets !== undefined) {
    USED_TABLES.add(context.externSheets);
  }
}

/**
 * Whether anything encoded through `context` resolved an `ixti`.
 *
 * The extern-sheet block — `BrtBeginExternals`, `BrtSupSelf`, `BrtExternSheet`, `BrtEndExternals` — is written only
 * when a token needs it, and only the encoder knows whether one did.
 */
export function externSheetWasUsed(context: PtgContext): boolean {
  return context.externSheets !== undefined
    ? USED_TABLES.has(context.externSheets)
    : (CONTEXT_INDEX.get(context)?.usedExternSheet ?? false);
}

/**
 * The `ixti` for an entry of this workbook's `BrtExternSheet`, appending one when the table has none.
 *
 * The table is built by the writer as an identity mapping — one single-sheet entry per sheet. A span
 * (`Sheet1:Sheet3`) and the workbook-level entry a `PtgNameX` goes through (-2, -2) have no such entry, so one is
 * added; the worksheets are serialised before the workbook part, so an entry appended while a formula is being
 * encoded still reaches the file. Excel's own table for a workbook with both carries `{0,1,1}`, `{0,-2,-2}` and
 * `{0,1,2}`.
 */
function internExternSheet(
  context: PtgContext,
  first: number,
  last: number,
  where: string
): number {
  const table = context.externSheets;
  if (!Array.isArray(table)) {
    // No table means the identity mapping, which has no room for anything but single sheets; and a caller that
    // handed a frozen table cannot be given an entry. Quietly narrowing the reference is the outcome this function
    // exists to remove.
    throw new XlsbParseError(
      where,
      `a reference through BrtExternSheet entry ${first}..${last} needs a table that can take one`
    );
  }
  // Searched in the live table rather than in a cached map: every context derived from the workbook's shares it.
  const entries = table as { first: number; last: number; supportingLink?: number }[];
  let ixti = entries.findIndex(
    entry => (entry.supportingLink ?? 0) === 0 && entry.first === first && entry.last === last
  );
  if (ixti === -1) {
    ixti = entries.length;
    entries.push({ first, last });
  }
  markUsed(context);
  return ixti;
}

function unsupported(where: string, what: string): never {
  throw new ExcelNotSupportedError(
    `Write XLSB formula at ${where}`,
    `${what} is not supported yet`
  );
}

/**
 * Wrap a token stream as a parsed formula: `cce`, the tokens, `cb`, then the `RgbExtra`.
 *
 * `CellParsedFormula` (MS-XLSB 2.5.98.3), `DVParsedFormula` (2.5.98.8), `NameParsedFormula` and their kind have
 * the same four-part layout, so one function serves them all. The `RgbExtra` carries a `PtgExp`'s column and an
 * array constant's values.
 */
export function encodeParsedFormula(tokens: Uint8Array, extra?: Uint8Array): Uint8Array {
  return concatUint8Arrays([
    new BinaryWriter().writeUint32(tokens.length).toUint8Array(),
    tokens,
    new BinaryWriter().writeUint32(extra?.length ?? 0).toUint8Array(),
    ...(extra === undefined ? [] : [extra])
  ]);
}
