/**
 * Print an AST back to formula text.
 *
 * The missing half of this module. `tokenize` and `parse` turn text into a tree and there
 * was no way back, which is an odd shape for a syntax layer to have: anything that
 * *transforms* a formula, normalises one, or reconstructs one from another representation
 * has to build strings by hand, and every such caller reimplements the one genuinely
 * difficult part — deciding where parentheses are required.
 *
 * The immediate consumer is the XLSB formula codec. BIFF12 stores a formula as a
 * reverse-polish token stream rather than as text, so reading one means `Ptg → AST → text`
 * and writing one means `text → AST → Ptg`. Without a printer the read direction has to
 * assemble text while walking the token stack, which means reimplementing precedence — and
 * a mistake there does not produce a syntax error, it produces a *different formula* that
 * parses fine and computes something else.
 *
 * ## Parenthesisation is minimal, and that is a normalisation
 *
 * Parentheses are emitted only where precedence or associativity requires them. `=(A1)`
 * prints as `=A1`; `=(A1+A2)*3` keeps its parentheses because dropping them would change
 * the tree. This is a deliberate choice and it is worth being precise about the
 * consequence: the AST does not record redundant parentheses, so a round trip through it
 * cannot preserve them. Excel does the same thing when a formula is edited.
 *
 * The invariant that matters is therefore structural, not textual, and it is what the
 * tests assert: `parse(tokenize(print(ast)))` is deep-equal to `ast`, and printing is
 * idempotent.
 *
 * ## Precedence lives in one place
 *
 * The binding powers below mirror `parser.ts`. They are not a second opinion about Excel's
 * precedence — if the two ever disagree, the round-trip test fails, which is the point of
 * expressing the invariant that way rather than by comparing strings.
 */

import { FormulaArgumentError } from "@formula/errors";
import { NodeType, type AstNode, type CellRefNode } from "@formula/syntax/ast";
import type { FormulaLayout, NodeLayout } from "@formula/syntax/layout";
import { SPILL_OPERATOR } from "@formula/syntax/token-types";
import { describeValue } from "@utils/describe-value";
import { sheetNameNeedsQuotes } from "@utils/sheet-name";

const BINARY_OPERATORS: ReadonlySet<string> = new Set([
  "=",
  "<>",
  "<",
  ">",
  "<=",
  ">=",
  "&",
  "+",
  "-",
  "*",
  "/",
  "^",
  ",",
  ":",
  " "
]);
const PREFIX_OPERATORS: ReadonlySet<string> = new Set(["+", "-", "@"]);
/** A function name as Excel writes one: `SUM`, `_xlfn.XLOOKUP`, a LAMBDA bound by `LET`. */
const FUNCTION_NAME = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;

function invalid(what: string): never {
  throw new FormulaArgumentError("node", `Formula.format: ${what}`);
}

/**
 * Left/right binding power of an infix operator, mirroring `parser.ts`.
 *
 * Higher binds tighter. `left < right` is left-associative.
 */
function infixBindingPower(op: string): readonly [number, number] {
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
      // Left-associative in Excel, unlike the mathematical convention.
      return [60, 61];
    case ",":
      // Reference union. Only legal inside parentheses, which `UnionRef` supplies.
      return [8, 9];
    case ":":
      return [90, 91];
    case " ":
      // Intersection. Binds tighter than arithmetic, looser than range.
      return [80, 81];
    default:
      return [0, 0];
  }
}

/** Unary `+`/`-`, which bind tighter than `%` and `^` in Excel — `-A1%` is `(-A1)%`. */
const PREFIX_BINDING_POWER = 77;
/** Postfix `%`, looser than negation and tighter than every binary arithmetic operator. */
const PERCENT_BINDING_POWER = 75;
/** A leaf never needs parenthesising. */
const ATOM_BINDING_POWER = 100;

/**
 * Print an AST as formula text, without a leading `=`. A tree no formula text can express —
 * an unknown node or operator, a non-finite number, a malformed function name — is refused
 * with a {@link FormulaError} rather than printed as text that does not parse back.
 */
export function printAst(node: AstNode): string {
  return print(node, undefined).text;
}

/**
 * Print an AST with the whitespace and display parentheses `layout` records — the inverse of `parseWithLayout`.
 *
 * Parentheses precedence requires are still inserted where the layout has none, so a layout can only add to what
 * `printAst` produces, never make it mean something else. A top-level union recorded as `bare` prints without its
 * parentheses, which is how a multi-area defined name is spelled.
 */
export function printAstWithLayout(node: AstNode, layout: FormulaLayout | undefined): string {
  if (layout !== undefined && node.type === NodeType.UnionRef && layout.get(node)?.bare) {
    const entry = layout.get(node)!;
    return `${entry.lead ?? ""}${node.areas.map(area => print(area, layout).text).join(",")}`;
  }
  return print(node, layout).text;
}

interface Printed {
  readonly text: string;
  /**
   * Binding power of the outermost operator, so a parent can decide whether this needs
   * wrapping. An atom reports the maximum, so it never does.
   */
  readonly power: number;
}

function print(node: AstNode, layout: FormulaLayout | undefined): Printed {
  const entry = layout?.get(node);
  const core = printCore(node, layout, entry);
  if (entry?.parens === undefined || entry.parens.length === 0) {
    return core;
  }
  let text = core.text;
  for (const paren of entry.parens) {
    text = `${paren.open ?? ""}(${text}${paren.close ?? ""})`;
  }
  return atom(text);
}

function printCore(
  node: AstNode,
  layout: FormulaLayout | undefined,
  entry: NodeLayout | undefined
): Printed {
  const lead = entry?.lead ?? "";
  if (typeof node !== "object" || node === null) {
    invalid(`expected an AST node, got ${describeValue(node)}`);
  }
  switch (node.type) {
    case NodeType.Number:
      if (typeof node.value !== "number" || !Number.isFinite(node.value)) {
        invalid(`a Number node must hold a finite number, got ${String(node.value)}`);
      }
      return atom(lead + printNumber(node.value));
    case NodeType.String:
      if (typeof node.value !== "string") {
        invalid("a String node must hold a string");
      }
      // Excel escapes a quote by doubling it.
      return atom(`${lead}"${node.value.replaceAll('"', '""')}"`);
    case NodeType.Boolean:
      return atom(lead + (node.value ? "TRUE" : "FALSE"));
    case NodeType.Error:
      if (node.externalReference !== undefined) {
        return atom(lead + node.externalReference);
      }
      return atom(lead + sheetPrefix(node.sheet, undefined) + node.value);
    case NodeType.Missing:
      return atom(lead);
    case NodeType.Name:
      return atom(lead + sheetPrefix(node.sheet, undefined) + node.name);
    case NodeType.CellRef:
      return atom(lead + printCellRef(node));
    case NodeType.RangeRef:
      return atom(
        `${lead}${sheetPrefix(node.sheet, node.endSheet)}${printLocalCellRef(node.start)}:` +
          `${printLocalCellRef(node.end)}`
      );
    case NodeType.ColRangeRef:
      return atom(
        `${lead}${sheetPrefix(node.sheet, node.endSheet)}${node.startAbsolute ? "$" : ""}${node.startCol}:` +
          `${node.endAbsolute ? "$" : ""}${node.endCol}`
      );
    case NodeType.RowRangeRef:
      return atom(
        `${lead}${sheetPrefix(node.sheet, node.endSheet)}${node.startAbsolute ? "$" : ""}${node.startRow}:` +
          `${node.endAbsolute ? "$" : ""}${node.endRow}`
      );
    case NodeType.StructuredRef:
      return atom(lead + printStructuredRef(node.tableName, node.columns, node.specials));
    case NodeType.Array:
      return atom(
        `${lead}{${node.rows.map(row => row.map(cell => print(cell, layout).text).join(",")).join(";")}}`
      );
    case NodeType.FunctionCall:
      if (typeof node.name !== "string" || !FUNCTION_NAME.test(node.name)) {
        invalid(`${JSON.stringify(node.name)} is not a function name`);
      }
      // The argument list supplies its own delimiters, so the call is an atom.
      return atom(
        `${lead}${printFunctionName(node.name)}(${node.args.map(arg => print(arg, layout).text).join(",")}${entry?.close ?? ""})` +
          (node.invocations ?? [])
            .map(
              (args, index) =>
                `(${args.map(arg => print(arg, layout).text).join(",")}${entry?.invocationClose?.[index] ?? ""})`
            )
            .join("")
      );
    case NodeType.UnionRef:
      // The parentheses are what make it a union; without them the commas would be
      // argument separators.
      return atom(
        `${lead}(${node.areas.map(area => print(area, layout).text).join(",")}${entry?.close ?? ""})`
      );
    case NodeType.Percent: {
      const operand = print(node.operand, layout);
      return {
        text: `${wrap(operand, PERCENT_BINDING_POWER)}${entry?.op ?? ""}%`,
        power: PERCENT_BINDING_POWER
      };
    }
    case NodeType.UnaryOp: {
      if (node.op === SPILL_OPERATOR) {
        // `A1#` — only a cell reference takes the spill operator, so anything else has no text that reads back.
        if (node.operand?.type !== NodeType.CellRef) {
          invalid("the spill operator `#` applies to a cell reference only");
        }
        return atom(`${print(node.operand, layout).text}${SPILL_OPERATOR}`);
      }
      if (!PREFIX_OPERATORS.has(node.op)) {
        invalid(`${JSON.stringify(node.op)} is not a prefix operator`);
      }
      const operand = print(node.operand, layout);
      return {
        text: `${lead}${node.op}${wrap(operand, PREFIX_BINDING_POWER)}`,
        power: PREFIX_BINDING_POWER
      };
    }
    case NodeType.BinaryOp: {
      if (!BINARY_OPERATORS.has(node.op)) {
        invalid(`${JSON.stringify(node.op)} is not an operator`);
      }
      const [leftPower, rightPower] = infixBindingPower(node.op);
      const left = print(node.left, layout);
      const right = print(node.right, layout);
      // A left operand may bind exactly as tightly as this operator and still not need
      // parentheses, because evaluation is left-to-right at equal precedence. A right
      // operand at equal precedence does need them, or the tree re-associates.
      //
      // The intersection operator is whitespace, so its recorded run *is* the operator.
      const separator = node.op === " " ? (entry?.op ?? " ") : (entry?.op ?? "") + node.op;
      return {
        text: `${wrap(left, leftPower)}${separator}${wrap(right, rightPower)}`,
        power: leftPower
      };
    }
    default:
      return invalid(`unknown node type ${String((node as { type?: unknown }).type)}`);
  }
}

function atom(text: string): Printed {
  return { text, power: ATOM_BINDING_POWER };
}

/**
 * Wrap when the child binds strictly looser than its position requires.
 *
 * One predicate serves both operands of an infix operator, because the asymmetry is
 * already in the binding powers: `left < right` by one, so a left operand at equal
 * precedence passes and a right operand at equal precedence does not. That is what makes
 * `1-2-3` print without parentheses while `1-(2-3)` keeps them.
 */
function wrap(child: Printed, required: number): string {
  return child.power < required ? `(${child.text})` : child.text;
}

/**
 * Print a number the way Excel writes it.
 *
 * `String(x)` is right for everything a formula can hold except the exponent form, where
 * JavaScript writes `1e+21` and Excel writes `1E+21`. Numbers are the one leaf where a
 * printer that "just works" silently produces text the tokenizer would read back
 * differently.
 */
function printNumber(value: number): string {
  const text = String(value);
  return text.includes("e") ? text.toUpperCase() : text;
}

function printCellRef(node: CellRefNode): string {
  return `${sheetPrefix(node.sheet, node.endSheet)}${printLocalCellRef(node)}`;
}

function printLocalCellRef(node: CellRefNode): string {
  return `${node.colAbsolute ? "$" : ""}${node.col}${node.rowAbsolute ? "$" : ""}${node.row}`;
}

/**
 * `Sheet1!`, `Sheet1:Sheet3!`, or `'My Sheet'!` when quoting is required.
 *
 * A sheet name needs quoting when it contains anything outside the unquoted set, and an
 * apostrophe inside it is doubled — the same rule as a string literal.
 */
function sheetPrefix(sheet: string | undefined, endSheet: string | undefined): string {
  if (sheet === undefined) {
    return "";
  }
  // A defined name's `!$A$1`: a reference to whichever sheet the name is used on.
  if (sheet === "" && endSheet === undefined) {
    return "!";
  }
  const first = quoteSheetName(sheet);
  const last = endSheet === undefined ? undefined : quoteSheetName(endSheet);
  // A quoted 3D reference quotes the pair, not each half: 'A B:C D'!A1.
  if (last !== undefined) {
    return needsQuoting(sheet) || needsQuoting(endSheet!)
      ? `'${escapeSheetName(sheet)}:${escapeSheetName(endSheet!)}'!`
      : `${sheet}:${endSheet}!`;
  }
  return `${first}!`;
}

function needsQuoting(name: string): boolean {
  // Excel's rule, shared with its reference writers: ASCII word characters only, no leading
  // digit, and nothing that reads as a cell, an R1C1 reference or a boolean (`'A1'!B2`).
  return sheetNameNeedsQuotes(name);
}

function escapeSheetName(name: string): string {
  return name.replaceAll("'", "''");
}

function quoteSheetName(name: string): string {
  return needsQuoting(name) ? `'${escapeSheetName(name)}'` : name;
}

/**
 * `Table1[Column]`, `Table1[[#Data],[Column]]`, `[@Column]`.
 *
 * The bracket form depends on how many parts there are: a single column with no special
 * item is written bare, and anything else needs each part in its own brackets.
 */
function printStructuredRef(
  tableName: string,
  columns: readonly string[],
  specials: readonly string[]
): string {
  // Two columns are a span from the first to the last, `[Qty]:[Note]`.
  const parts = [
    ...specials.map(special => `[${special}]`),
    ...(columns.length === 0
      ? []
      : [columns.map(column => `[${escapeStructuredItem(column)}]`).join(":")])
  ];
  if (parts.length === 0) {
    return `${tableName}[]`;
  }
  // A single part needs no second bracket, whether it is a column or an item specifier: Excel writes `Tbl[#All]`.
  // A span is one part but two brackets, so it keeps the outer pair: `Tbl[[Qty]:[Note]]`.
  if (parts.length === 1 && columns.length < 2) {
    return `${tableName}${parts[0]}`;
  }
  return `${tableName}[${parts.join(",")}]`;
}

/**
 * A column name inside a structured reference, with `[`, `]`, `#` and `'` escaped by a
 * preceding `'` as Excel writes them: `Tbl[a']b]` is the column `a]b`.
 */
function escapeStructuredItem(column: string): string {
  return column.replace(/[[\]#']/g, "'$&");
}

/** Excel's future-function prefixes, which it writes in lowercase: `_xlfn.XLOOKUP`. */
const FUTURE_PREFIXES = /^(?:_xlfn\.|_xlws\.|_xludf\.)+/i;

/**
 * A function name as Excel writes it. The tokenizer reads a name case-insensitively and
 * keeps it in upper case, prefix included; the `_xlfn.` / `_xlws.` / `_xludf.` prefixes go
 * back to lower case, as Excel writes them in a file.
 */
function printFunctionName(name: string): string {
  const prefix = FUTURE_PREFIXES.exec(name);
  return prefix === null ? name : prefix[0].toLowerCase() + name.slice(prefix[0].length);
}
