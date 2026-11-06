/**
 * Future functions: the ones newer than the file formats' own function lists, which a file names with a prefix.
 *
 * A user writes `FILTER(A1:A3,B1:B3>2)`; the XLSX file has to hold `_xlfn._xlws.FILTER(A1:A3,B1:B3>2)`. Without the
 * prefix Excel reads the call as a user-defined function — `XLOOKUP`, `SEQUENCE`, `UNIQUE`, `VSTACK` and `LET` show
 * `#NAME?` — and for `FILTER` and `SORT` it refuses the whole file. XLSB spells the same call through a hidden
 * defined name carrying the same prefixed text, so both writers take the prefix from here.
 *
 * `LET` and `LAMBDA` parameters are prefixed too: Excel stores `LET(x,2,x*3)` as `_xlfn.LET(_xlpm.x,2,_xlpm.x*3)`.
 *
 * Reading leaves the stored text alone — `_xlfn.XLOOKUP(…)` reads back as written in the file, which is what every
 * reader of an Excel-made workbook has always returned, and what the formula engine accepts. A formula *set* bare is
 * stored prefixed, so after a write and a read it comes back in the stored, prefixed form.
 *
 * **The one exception is the spill-range operator.** `A1#` is not a function a formula is written with: it is an
 * operator Excel's formula bar shows, and the file grammar has no such operator — MS-XLSX's `Formulas` ABNF knows `%`
 * as its only postfix operator — so a file stores it as the future function `_xlfn.ANCHORARRAY(A1)`. `ANCHORARRAY` is
 * newer than MS-XLSX's future-function list (see below); Excel treats it as one — its own XLSB save of
 * `xlsb/__tests__/data/future-functions.excel.xlsb` calls `_xlfn.ANCHORARRAY` through the hidden stub name, exactly
 * as it calls `_xlfn.XLOOKUP`, and `future-function-tokens.node.test.ts` pins those bytes. {@link toFileFormula}
 * writes the operator in that form and {@link fromFileFormula} reads it back as the operator, so `=A1#` round-trips
 * as `A1#`, and the file is the same either way.
 */
import { bareFunctionName, isLambdaParameter, XLFN, XLPM, XLWS } from "@excel/utils/function-names";
import {
  AT_POWER,
  operandEnd,
  ScanKind,
  type ScanToken,
  scanFormula
} from "@excel/utils/spill-scan";
import { FUNCTION_ID_BY_NAME } from "@formula/syntax/function-table";
import { excelNameKey, sheetNameKey } from "@utils/name-key";

/** Functions called through `_xlfn._xlws.`: worksheet-only functions in MS-XLSX's list. */
const WORKSHEET_FUNCTIONS = ["FILTER", "SORT"];

/**
 * Functions called through `_xlfn.`. The list is MS-XLSX 2.2.2 ("Functions", the future-function list), then the
 * functions Excel added after that list was last revised, each confirmed against Excel by opening a file that calls
 * it with the prefix and reading the computed value.
 *
 * Four names on MS-XLSX's list carry no prefix — `ECMA.CEILING`, `ISO.CEILING`, `NETWORKDAYS.INTL`,
 * `WORKDAY.INTL` — so XLSX text needs nothing for them, and they are not here.
 *
 * Two names on the list are deliberately left out: `_xlfn._xlws.PY` and `_xlfn.PQSOURCE`. Excel refuses the whole
 * workbook when a cell spells either with that prefix, and opens it — the cell showing `#NAME?` — when the call is
 * bare. No file Excel itself wrote shows what else a working Python or Power Query cell needs, so they are written
 * as given.
 */
const FUTURE_FUNCTIONS = [
  // MS-XLSX 2.2.2.
  "ACOT",
  "ACOTH",
  "AGGREGATE",
  "ARABIC",
  "BASE",
  "BETA.DIST",
  "BETA.INV",
  "BINOM.DIST",
  "BINOM.DIST.RANGE",
  "BINOM.INV",
  "BITAND",
  "BITLSHIFT",
  "BITOR",
  "BITRSHIFT",
  "BITXOR",
  "BYCOL",
  "BYROW",
  "CEILING.MATH",
  "CEILING.PRECISE",
  "CHISQ.DIST",
  "CHISQ.DIST.RT",
  "CHISQ.INV",
  "CHISQ.INV.RT",
  "CHISQ.TEST",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "COMBINA",
  "CONCAT",
  "CONFIDENCE.NORM",
  "CONFIDENCE.T",
  "COT",
  "COTH",
  "COVARIANCE.P",
  "COVARIANCE.S",
  "CSC",
  "CSCH",
  "DAYS",
  "DECIMAL",
  "DROP",
  "ERF.PRECISE",
  "ERFC.PRECISE",
  "EXPAND",
  "EXPON.DIST",
  "F.DIST",
  "F.DIST.RT",
  "F.INV",
  "F.INV.RT",
  "F.TEST",
  "FIELDVALUE",
  "FILTERXML",
  "FLOOR.MATH",
  "FLOOR.PRECISE",
  "FORECAST.ETS",
  "FORECAST.ETS.CONFINT",
  "FORECAST.ETS.SEASONALITY",
  "FORECAST.ETS.STAT",
  "FORECAST.LINEAR",
  "FORMULATEXT",
  "GAMMA",
  "GAMMA.DIST",
  "GAMMA.INV",
  "GAMMALN.PRECISE",
  "GAUSS",
  "HSTACK",
  "HYPGEOM.DIST",
  "IFNA",
  "IFS",
  "IMCOSH",
  "IMCOT",
  "IMCSC",
  "IMCSCH",
  "IMSEC",
  "IMSECH",
  "IMSINH",
  "IMTAN",
  "ISFORMULA",
  "ISOMITTED",
  "ISOWEEKNUM",
  "LAMBDA",
  "LET",
  "LOGNORM.DIST",
  "LOGNORM.INV",
  "MAKEARRAY",
  "MAP",
  "MAXIFS",
  "MINIFS",
  "MODE.MULT",
  "MODE.SNGL",
  "MUNIT",
  "NEGBINOM.DIST",
  "NORM.DIST",
  "NORM.INV",
  "NORM.S.DIST",
  "NORM.S.INV",
  "NUMBERVALUE",
  "PDURATION",
  "PERCENTILE.EXC",
  "PERCENTILE.INC",
  "PERCENTRANK.EXC",
  "PERCENTRANK.INC",
  "PERMUTATIONA",
  "PHI",
  "POISSON.DIST",
  "PYTHON_STR",
  "PYTHON_TYPE",
  "PYTHON_TYPENAME",
  "QUARTILE.EXC",
  "QUARTILE.INC",
  "QUERYSTRING",
  "RANDARRAY",
  "RANK.AVG",
  "RANK.EQ",
  "REDUCE",
  "RRI",
  "SCAN",
  "SEC",
  "SECH",
  "SEQUENCE",
  "SHEET",
  "SHEETS",
  "SKEW.P",
  "SORTBY",
  "STDEV.P",
  "STDEV.S",
  "SWITCH",
  "T.DIST",
  "T.DIST.2T",
  "T.DIST.RT",
  "T.INV",
  "T.INV.2T",
  "T.TEST",
  "TAKE",
  "TEXTAFTER",
  "TEXTBEFORE",
  "TEXTJOIN",
  "TEXTSPLIT",
  "TOCOL",
  "TOROW",
  "UNICHAR",
  "UNICODE",
  "UNIQUE",
  "VAR.P",
  "VAR.S",
  "VSTACK",
  "WEBSERVICE",
  "WEIBULL.DIST",
  "WRAPCOLS",
  "WRAPROWS",
  "XLOOKUP",
  "XOR",
  "Z.TEST",
  // Newer than MS-XLSX's list.
  "ANCHORARRAY",
  "ARRAYTOTEXT",
  "GROUPBY",
  "IMAGE",
  "PERCENTOF",
  "PIVOTBY",
  "REGEXEXTRACT",
  "REGEXREPLACE",
  "REGEXTEST",
  "SINGLE",
  "STOCKHISTORY",
  "TRIMRANGE",
  "VALUETOTEXT",
  "XMATCH"
];

const XLETA = "_xleta.";

export { bareFunctionName, isLambdaParameter };

/** Whether `name` (with or without `_xlfn.`) is a worksheet function's: one the `Ftab` lists, or a future function. */
export function isWorksheetFunctionName(name: string): boolean {
  const bare = bareFunctionName(name);
  return FUNCTION_ID_BY_NAME.has(bare) || FUTURE_FUNCTION_PREFIXES.has(bare);
}

/**
 * Where each built-in that takes a `LAMBDA` takes it, by zero-based argument index — `-1` for the last argument, which
 * is where `MAP` takes it after any number of arrays. Only there can a function be passed by name: `BYROW(SUM,…)` names
 * an array `SUM`, and `GROUPBY`'s fourth argument is a header setting, so neither is `_xleta.`.
 */
const LAMBDA_ARGUMENTS: ReadonlyMap<string, number> = new Map([
  ["BYCOL", 1],
  ["BYROW", 1],
  ["GROUPBY", 2],
  ["MAKEARRAY", 2],
  ["MAP", -1],
  ["PIVOTBY", 3],
  ["REDUCE", 2],
  ["SCAN", 2]
]);

/**
 * Whether argument `argument` (zero-based) of `count` in a call to `name` can be passed a function by name: the
 * argument a built-in takes its `LAMBDA` in, or any argument of a function that is not a worksheet function — a
 * `LAMBDA` behind a defined name. Anywhere else a bare function name is a name: `COS(NA)` refers to a name `NA`, and
 * Excel reports `#NAME?` for it rather than reading `_xleta.NA`.
 */
export function acceptsFunctionArgument(name: string, argument: number, count: number): boolean {
  const bare = bareFunctionName(name.replace(/^_xlpm\./i, ""));
  const at = LAMBDA_ARGUMENTS.get(bare);
  if (at !== undefined) {
    return argument === (at === -1 ? count - 1 : at);
  }
  return !isWorksheetFunctionName(bare);
}

/** The workbook's defined names, by `excelNameKey`, while one of its parts is rendered — see {@link withFormulaNames}. */
let renderingNames: ReadonlySet<string> | undefined;

/**
 * Run `render` — one part's synchronous serialisation — with the workbook's defined names in force for
 * {@link toFileFormula}.
 *
 * A function passed by name is written `_xleta.SUM`, but a bare `SUM` is a user's name if the workbook defines one,
 * and only the workbook knows. The writers carry the names in their write context and set them around each part they
 * render, so every formula in the part — cells, shared and array formulas, validations, conditional formats, defined
 * names, table columns — is converted by the same rule without each xform being handed them. `render` is synchronous,
 * so two writes in flight cannot see each other's names.
 */
export function withFormulaNames<T>(names: ReadonlySet<string> | undefined, render: () => T): T {
  const previous = renderingNames;
  renderingNames = names;
  try {
    return render();
  } finally {
    renderingNames = previous;
  }
}

/** The keys (`excelNameKey`) of a defined-name list, as {@link withFormulaNames} takes them. */
export function formulaNameSet(names: Iterable<{ readonly name: string }>): ReadonlySet<string> {
  return new Set([...names].map(entry => excelNameKey(entry.name)));
}

/** Every future function, by its bare upper-case name, to the prefix a file spells it with. */
export const FUTURE_FUNCTION_PREFIXES: ReadonlyMap<string, string> = new Map([
  ...FUTURE_FUNCTIONS.map(name => [name, XLFN] as const),
  ...WORKSHEET_FUNCTIONS.map(name => [name, XLFN + XLWS] as const)
]);

/** The `_xlpm.` parameter names stored formula text uses, as written. */
export function lambdaParameters(stored: string): string[] {
  return [...stored.matchAll(/(?<![\p{L}\p{N}_.\\$])_xlpm\.[\p{L}\p{N}_.\\?]+/giu)].map(
    match => match[0]
  );
}

/**
 * The name a file calls `name` by: the future-function prefix and the upper-case name, or `undefined` when `name`
 * is not a future function.
 */
export function storedFunctionName(name: string): string | undefined {
  const bare = bareFunctionName(name);
  const prefix = FUTURE_FUNCTION_PREFIXES.get(bare);
  return prefix === undefined ? undefined : prefix + bare;
}

/** One lexical piece of formula text: verbatim text, or a bare identifier. */
interface Piece {
  readonly text: string;
  readonly identifier: boolean;
}

const IDENTIFIER_START = /[\p{L}_\\$]/u;
const IDENTIFIER_PART = /[\p{L}\p{N}_.\\$?]/u;

/**
 * Split formula text into identifiers and everything else, keeping every character. Strings, quoted sheet names and
 * bracketed parts (structured and external references) are verbatim: nothing inside them is a function or a name.
 */
function split(text: string, thisRow?: ThisRowContext): Piece[] {
  const pieces: Piece[] = [];
  let verbatim = "";
  const flush = (): void => {
    if (verbatim !== "") {
      pieces.push({ text: verbatim, identifier: false });
      verbatim = "";
    }
  };
  let at = 0;
  while (at < text.length) {
    const ch = text[at]!;
    if (ch === '"' || ch === "'") {
      let end = at + 1;
      while (end < text.length) {
        if (text[end] === ch) {
          if (text[end + 1] === ch) {
            end += 2;
            continue;
          }
          break;
        }
        end++;
      }
      verbatim += text.slice(at, end + 1);
      at = end + 1;
      continue;
    }
    if (ch === "[") {
      let depth = 0;
      let end = at;
      for (; end < text.length; end++) {
        if (text[end] === "'") {
          end++;
        } else if (text[end] === "[") {
          depth++;
        } else if (text[end] === "]" && --depth === 0) {
          break;
        }
      }
      const bracket = text.slice(at, end + 1);
      if (IDENTIFIER_PART.test(text[at - 1] ?? "")) {
        verbatim += thisRowFileForm(bracket);
      } else if (thisRow !== undefined && isThisRowBracket(bracket)) {
        // An unnamed this-row reference. Excel stores it with the enclosing table's name and refuses — or reads as
        // `#REF!` — the unnamed form, so it is written with the name when there is one to give.
        thisRow.found = true;
        verbatim +=
          thisRow.table === undefined ? bracket : thisRow.table + thisRowFileForm(bracket);
      } else {
        verbatim += bracket;
      }
      at = end + 1;
      continue;
    }
    if (/[0-9]/.test(ch)) {
      // A number, or the row part of a reference — never the start of a name.
      let end = at + 1;
      while (end < text.length && /[0-9.]/.test(text[end]!)) {
        end++;
      }
      if (/[eE]/.test(text[end] ?? "") && /[-+0-9]/.test(text[end + 1] ?? "")) {
        end += 2;
        while (end < text.length && /[0-9]/.test(text[end]!)) {
          end++;
        }
      }
      verbatim += text.slice(at, end);
      at = end;
      continue;
    }
    if (IDENTIFIER_START.test(ch)) {
      let end = at + 1;
      while (end < text.length && IDENTIFIER_PART.test(text[end]!)) {
        end++;
      }
      flush();
      pieces.push({ text: text.slice(at, end), identifier: true });
      at = end;
      continue;
    }
    verbatim += ch;
    at++;
  }
  flush();
  return pieces;
}

/** What `toFileFormula` knows about the table enclosing a formula, and what it found. */
interface ThisRowContext {
  /** The enclosing table's name, or `undefined` when the formula is in none. */
  readonly table: string | undefined;
  /** Set when the text holds an unnamed this-row reference. */
  found: boolean;
}

/** Whether a bracketed part with no table name before it is a this-row reference: `[@…]` or `[#This Row]…`. */
function isThisRowBracket(bracket: string): boolean {
  return /^\[\s*(?:@|\[?\s*#This Row\s*\])/i.test(bracket);
}

/**
 * Whether formula text holds a this-row reference with no table name — `[@x]`, `[@[x]:[y]]`, `[@]` or
 * `[[#This Row],[x]]`. Such text is only meaningful inside a table, and a file can only store it with that table's
 * name: see `toFileFormula`'s `table` option.
 */
export function hasUnnamedThisRow(text: string): boolean {
  if (!text.includes("[")) {
    return false;
  }
  const context: ThisRowContext = { table: undefined, found: false };
  split(text, context);
  return context.found;
}

/**
 * A structured reference's bracketed part with its `@` written out: `[@x]` → `[[#This Row],[x]]`,
 * `[@[x]:[y]]` → `[[#This Row],[x]:[y]]`, `[@]` → `[#This Row]`.
 *
 * `@` is how Excel's formula bar shows a this-row reference, not how a file stores one: Excel refuses a workbook
 * whose cell holds `Tb[@[x]:[y]]`, and reads `Tb[@x]:Tb[@y]` as `#REF!:#REF!`. Anything else is returned as given.
 *
 * Inside a table `[@x]` may leave the table out, but Excel's own files always name it (`Tb[[#This Row],[x]]`, in
 * the cells and the `calculatedColumnFormula` of `sb-tbl.xlsx`), and Excel refuses a `calculatedColumnFormula` — and
 * reads a cell formula as `#REF!` — that does not. An unnamed one is therefore given its table's name by the caller
 * that knows it; see `toFileFormula`'s `table` option.
 */
function thisRowFileForm(bracket: string): string {
  const inner = /^\[\s*@([^]*)\]$/.exec(bracket)?.[1]?.trim();
  if (inner === undefined) {
    return bracket;
  }
  if (inner === "") {
    return "[#This Row]";
  }
  return `[[#This Row],${inner.startsWith("[") ? inner : `[${inner}]`}]`;
}

/** The index of the next piece after `index` that is not only whitespace, or `-1`. */
function nextSignificant(pieces: readonly Piece[], index: number): number {
  for (let at = index + 1; at < pieces.length; at++) {
    if (pieces[at]!.identifier || pieces[at]!.text.trim() !== "") {
      return at;
    }
  }
  return -1;
}

/** The first non-blank character of a verbatim piece. */
function leading(piece: Piece | undefined): string {
  return piece === undefined || piece.identifier ? "" : (piece.text.trimStart()[0] ?? "");
}

/** The last non-blank character of a verbatim piece. */
function trailing(piece: Piece | undefined): string {
  if (piece === undefined || piece.identifier) {
    return "";
  }
  const trimmed = piece.text.trimEnd();
  return trimmed[trimmed.length - 1] ?? "";
}

interface Frame {
  /** The bare upper-case name of the call, or `undefined` for a parenthesis or an array constant. */
  readonly call: string | undefined;
  /** The index of the argument being read, by commas at this level. */
  argument: number;
  /** The identifier pieces making up each argument when that argument is a lone identifier. */
  readonly lone: (number | undefined)[];
  /** Whether the current argument has seen anything yet. */
  started: boolean;
}

/**
 * For each `LET` / `LAMBDA` call, the piece indices that open it and the arguments it has, found before rewriting:
 * whether an argument declares a parameter depends on whether it is the last one.
 */
function callShapes(pieces: readonly Piece[]): Map<number, { lone: (number | undefined)[] }> {
  const shapes = new Map<number, { lone: (number | undefined)[] }>();
  const stack: { frame: Frame; opener: number }[] = [];
  let pendingCall: { name: string; at: number } | undefined;
  for (let index = 0; index < pieces.length; index++) {
    const piece = pieces[index]!;
    const top = stack[stack.length - 1]?.frame;
    if (piece.identifier) {
      const after = nextSignificant(pieces, index);
      if (after !== -1 && leading(pieces[after]) === "(") {
        pendingCall = { name: bareFunctionName(piece.text.replace(/^_xlpm\./i, "")), at: index };
      }
      if (top !== undefined) {
        top.lone[top.argument] = top.started ? undefined : index;
        top.started = true;
      }
      continue;
    }
    for (const ch of piece.text) {
      const frame = stack[stack.length - 1]?.frame;
      if (ch === "(" || ch === "{") {
        if (frame !== undefined) {
          frame.lone[frame.argument] = undefined;
          frame.started = true;
        }
        const call = ch === "(" ? pendingCall : undefined;
        stack.push({
          frame: { call: call?.name, argument: 0, lone: [], started: false },
          opener: call?.at ?? -1
        });
        pendingCall = undefined;
      } else if (ch === ")" || ch === "}") {
        const closed = stack.pop();
        if (closed !== undefined && closed.opener !== -1) {
          closed.frame.lone.length = closed.frame.argument + 1;
          shapes.set(closed.opener, { lone: closed.frame.lone });
        }
      } else if (ch === "," && frame !== undefined && frame.call !== undefined) {
        frame.argument++;
        frame.started = false;
      } else if (ch.trim() !== "" && frame !== undefined) {
        frame.lone[frame.argument] = undefined;
        frame.started = true;
      }
    }
  }
  return shapes;
}

/** The parameter names each `LET` / `LAMBDA` declares, by the piece index of the declaring identifier. */
function declarations(
  pieces: readonly Piece[],
  shapes: Map<number, { lone: (number | undefined)[] }>
): Map<number, { call: number; argument: number }> {
  const declared = new Map<number, { call: number; argument: number }>();
  for (const [opener, { lone }] of shapes) {
    const name = bareFunctionName(pieces[opener]!.text);
    const last = lone.length - 1;
    lone.forEach((piece, argument) => {
      const declares =
        name === "LET"
          ? argument % 2 === 0 && argument < last
          : name === "LAMBDA" && argument < last;
      if (declares && piece !== undefined) {
        declared.set(piece, { call: opener, argument });
      }
    });
  }
  return declared;
}

/**
 * Formula text as a file stores it: the spill-range operator `A1#` is `_xlfn.ANCHORARRAY(A1)`, future functions carry
 * their `_xlfn.` / `_xlfn._xlws.` prefix, `LET` /
 * `LAMBDA` parameters their `_xlpm.` one, and a this-row reference is `[#This Row]` rather than `@`. Text already in that form is returned unchanged, and anything that is not
 * a recognised call or parameter — a user-defined function, a defined name, a string — is left as written. One
 * leading `=` is dropped.
 */
/** `#` straight after a cell reference's last character — the only place the operator can be. */
const SPILL_OPERATOR_HINT = /[A-Za-z0-9$]#/;
const ANCHORARRAY_HINT = /ANCHORARRAY/i;
const ANCHORARRAY = "_XLFN.ANCHORARRAY";

/**
 * `A1#` → `_xlfn.ANCHORARRAY(A1)`, for every spill-range operator in `text` — on a cell reference, optionally
 * sheet-qualified (`'My Sheet'!$A$1#`), which is the only operand Excel and the formula parser accept. Read with
 * `spill-scan.ts` rather than the formula tokenizer, which a write-only bundle must not carry; text the tokenizer
 * cannot read is left as given.
 */
function spillOperatorsToFileForm(text: string): string {
  if (!SPILL_OPERATOR_HINT.test(text)) {
    return text;
  }
  const tokens = scanFormula(text);
  if (tokens === undefined) {
    return text;
  }
  let out = "";
  let at = 0;
  for (let i = 1; i < tokens.length; i++) {
    if (tokens[i]!.kind !== ScanKind.Spill || tokens[i - 1]!.kind !== ScanKind.Cell) {
      continue;
    }
    const qualified = i >= 2 && tokens[i - 2]!.kind === ScanKind.Sheet;
    const start = tokens[qualified ? i - 2 : i - 1]!.start;
    const end = tokens[i - 1]!.end;
    out += `${text.slice(at, start)}_xlfn.ANCHORARRAY(${text.slice(start, end)})`;
    at = tokens[i]!.end;
  }
  return at === 0 ? text : out + text.slice(at);
}

const SINGLE = "_XLFN.SINGLE";
const SINGLE_HINT = /SINGLE/i;

/**
 * `@x` → `_xlfn.SINGLE(x)`, for every implicit-intersection operator in `text`, `x` being exactly the operand the
 * parser gives `@` — `@A1:A3*2` is `_xlfn.SINGLE(A1:A3)*2`, `@A1^2` is `_xlfn.SINGLE(A1^2)`. MS-XLSX's formula grammar
 * has `+` and `-` as its only prefix operators, and the call is how Excel writes `@` in a file (its XLSB of this
 * library's `future-functions.xlsx` calls `_xlfn.SINGLE` through a stub, `ptg.ts` `CALC_EXP_FUTURE_FUNCTIONS`). An
 * `@` inside a structured reference — `Tb[@x]` — is part of that reference, not the operator. Text whose operand
 * cannot be found is left as given.
 */
function intersectionOperatorsToFileForm(text: string): string {
  if (!text.includes("@")) {
    return text;
  }
  const tokens = scanFormula(text);
  if (tokens === undefined || !tokens.some(token => token.kind === ScanKind.At)) {
    return text;
  }
  return rewriteOperators(text, tokens, 0, tokens.length) ?? text;
}

/** `text` from token `from` to token `to` with each `@` written as the call; `undefined` for an operand not found. */
function rewriteOperators(
  text: string,
  tokens: readonly ScanToken[],
  from: number,
  to: number
): string | undefined {
  const textEnd = to < tokens.length ? tokens[to]!.start : text.length;
  let out = "";
  let at = from < tokens.length ? tokens[from]!.start : text.length;
  for (let i = from; i < to; i++) {
    if (tokens[i]!.kind !== ScanKind.At) {
      continue;
    }
    const end = operandEnd(tokens, i + 1, AT_POWER);
    if (end === undefined || end > to) {
      return undefined;
    }
    const inner = rewriteOperators(text, tokens, i + 1, end);
    if (inner === undefined) {
      return undefined;
    }
    // From just after the `@`, so whitespace between it and its operand is kept, inside the call.
    out += `${text.slice(at, tokens[i]!.start)}_xlfn.SINGLE(${text.slice(tokens[i]!.end, tokens[i + 1]!.start)}${inner})`;
    at = tokens[end - 1]!.end;
    i = end - 1;
  }
  return out + text.slice(at, textEnd);
}

/**
 * `_xlfn.SINGLE(x)` → `@x`, where `@x` reads as the same formula: one argument, written as a file writes it (no
 * whitespace inside the call), exactly the operand `@` takes, and nothing after the call that would join it — `^`,
 * `%`, `:` and an intersection bind tighter than `@`, so `_xlfn.SINGLE(A1)^2` stays a call.
 */
function intersectionCallsFromFileForm(text: string): string {
  if (!SINGLE_HINT.test(text)) {
    return text;
  }
  const tokens = scanFormula(text);
  if (tokens === undefined) {
    return text;
  }
  return readCalls(text, tokens, 0, tokens.length);
}

function readCalls(text: string, tokens: readonly ScanToken[], from: number, to: number): string {
  const textEnd = to < tokens.length ? tokens[to]!.start : text.length;
  let out = "";
  let at = from < tokens.length ? tokens[from]!.start : text.length;
  for (let i = from; i < to; i++) {
    const token = tokens[i]!;
    if (
      token.kind !== ScanKind.Function ||
      token.name !== SINGLE ||
      tokens[i + 1]?.kind !== ScanKind.Open
    ) {
      continue;
    }
    const open = i + 1;
    const end = operandEnd(tokens, open + 1, AT_POWER);
    const close = tokens[end ?? -1];
    const after = tokens[(end ?? 0) + 1];
    if (
      end === undefined ||
      close?.kind !== ScanKind.Close ||
      tokens[open]!.end !== tokens[open + 1]!.start ||
      tokens[end - 1]!.end !== close.start ||
      (after !== undefined &&
        (after.kind === ScanKind.Percent ||
          after.kind === ScanKind.Colon ||
          after.kind === ScanKind.Intersect ||
          after.kind === ScanKind.Spill ||
          (after.kind === ScanKind.Infix && after.name === "^")))
    ) {
      continue;
    }
    out += `${text.slice(at, token.start)}@${readCalls(text, tokens, open + 1, end)}`;
    at = close.end;
    i = end;
  }
  return out + text.slice(at, textEnd);
}

/**
 * Formula text as this library reads it from a file: the file's text, with `_xlfn.ANCHORARRAY(ref)` — how a file
 * stores the spill-range operator — read back as `ref#`, the operator Excel shows. Only a call on one cell
 * reference, optionally sheet-qualified, is the operator; any other argument (a name, a range, a call) keeps the
 * function form, which the formula engine evaluates alike. The inverse of {@link toFileFormula} for it.
 */
export function fromFileFormula(text: string): string {
  return intersectionCallsFromFileForm(anchorCallsFromFileForm(text));
}

/** `_xlfn.ANCHORARRAY(ref)` → `ref#`; see {@link fromFileFormula}. */
function anchorCallsFromFileForm(text: string): string {
  if (!ANCHORARRAY_HINT.test(text)) {
    return text;
  }
  const tokens = scanFormula(text);
  if (tokens === undefined) {
    return text;
  }
  let out = "";
  let at = 0;
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i]!;
    if (token.kind !== ScanKind.Function || token.name !== ANCHORARRAY) {
      continue;
    }
    const qualified = tokens[i + 2]?.kind === ScanKind.Sheet;
    const cell = i + (qualified ? 3 : 2);
    const close = cell + 1;
    if (
      tokens[i + 1]?.kind !== ScanKind.Open ||
      tokens[cell]?.kind !== ScanKind.Cell ||
      tokens[close]?.kind !== ScanKind.Close
    ) {
      continue;
    }
    // Contiguous, as a file writes it: whitespace inside the call would be lost by the operator form.
    let contiguous = true;
    for (let k = i; k < close; k++) {
      if (tokens[k]!.end !== tokens[k + 1]!.start) {
        contiguous = false;
      }
    }
    if (!contiguous) {
      continue;
    }
    out += `${text.slice(at, token.start)}${text.slice(tokens[i + 2]!.start, tokens[cell]!.end)}#`;
    at = tokens[close]!.end;
    i = close;
  }
  return at === 0 ? text : out + text.slice(at);
}

export function toFileFormula(
  text: string,
  options?: {
    /**
     * The name of the table the formula sits in — a cell inside its range, a calculated column, a totals row. An
     * unnamed this-row reference (`[@x]`) is then written with it (`Tb[[#This Row],[x]]`), the only form Excel stores;
     * without it such a reference is written as given.
     */
    readonly table?: string;
    /**
     * The workbook's defined names, by `excelNameKey`. A function name passed as an argument — `GROUPBY(A2:A9,B2:B9,SUM)` —
     * is then written `_xleta.SUM`, as Excel writes it, unless the workbook defines that name. Defaults to the names
     * {@link withFormulaNames} put in force; with neither, such a name is written as given.
     */
    readonly definedNames?: ReadonlySet<string>;
  }
): string {
  // A formula is often written as typed into a cell, `=1+1`, but a file stores it without the `=`, which is not part
  // of the formula.
  if (typeof text === "string" && text.startsWith("=")) {
    text = text.slice(1);
  }
  // `A1#` is `_xlfn.ANCHORARRAY(A1)` in a file: Excel reads the operator written as it is as a broken formula.
  if (typeof text === "string") {
    text = intersectionOperatorsToFileForm(spillOperatorsToFileForm(text));
  }
  const definedNames = options?.definedNames ?? renderingNames;
  if (!/[[\p{L}_]/u.test(text)) {
    return text;
  }
  const pieces = split(
    text,
    options?.table === undefined ? undefined : { table: options.table, found: false }
  );
  const shapes = callShapes(pieces);
  const declared = declarations(pieces, shapes);
  const loneArguments = new Set<number>();
  for (const [call, { lone }] of shapes) {
    lone.forEach((piece, argument) => {
      if (
        piece !== undefined &&
        acceptsFunctionArgument(pieces[call]!.text, argument, lone.length)
      ) {
        loneArguments.add(piece);
      }
    });
  }
  /** Scopes open at the current point: the call they belong to, its argument index, and what they declare. */
  interface Scope {
    readonly call: number | undefined;
    argument: number;
    readonly names: { name: string; from: number }[];
  }
  const scopes: Scope[] = [];
  let pendingCall: number | undefined;
  const visible = (name: string): boolean =>
    scopes.some(scope =>
      scope.names.some(entry => entry.name === name && scope.argument >= entry.from)
    );
  let out = "";
  for (let index = 0; index < pieces.length; index++) {
    const piece = pieces[index]!;
    if (!piece.identifier) {
      for (const ch of piece.text) {
        const scope = scopes[scopes.length - 1];
        if (ch === "(" || ch === "{") {
          scopes.push({ call: ch === "(" ? pendingCall : undefined, argument: 0, names: [] });
          pendingCall = undefined;
        } else if (ch === ")" || ch === "}") {
          scopes.pop();
        } else if (ch === "," && scope !== undefined && scope.call !== undefined) {
          scope.argument++;
        }
      }
      out += piece.text;
      continue;
    }
    const after = nextSignificant(pieces, index);
    const isCall = after !== -1 && leading(pieces[after]) === "(";
    if (isCall) {
      pendingCall = index;
    }
    const qualified = trailing(pieces[index - 1]) === "!" || leading(pieces[after]) === "!";
    const bareParameter = piece.text.replace(/^_xlpm\./i, "");
    const declaration = declared.get(index);
    if (declaration !== undefined) {
      // A `LET` name is in scope from the argument after its value; a `LAMBDA` parameter in the body.
      const owner = scopes[scopes.length - 1]!;
      const isLet = bareFunctionName(pieces[declaration.call]!.text) === "LET";
      owner.names.push({
        name: sheetNameKey(bareParameter),
        from: declaration.argument + (isLet ? 2 : 1)
      });
      out += XLPM + bareParameter;
      continue;
    }
    if (!qualified && visible(sheetNameKey(bareParameter))) {
      out += XLPM + bareParameter;
      continue;
    }
    if (isCall && !qualified) {
      out += storedFunctionName(piece.text) ?? piece.text;
      continue;
    }
    // A function passed by name: Excel's XLSX of `GROUPBY(A2:A9,B2:B9,SUM)` holds `_xleta.SUM` — the bare name, no
    // `_xlfn.` even for `PERCENTOF`. Without the prefix Excel reads `SUM` as an undefined name.
    if (
      definedNames !== undefined &&
      !qualified &&
      loneArguments.has(index) &&
      !definedNames.has(excelNameKey(piece.text)) &&
      // `TRUE` and `FALSE` are functions too, and as an argument they are the constants.
      !/^(_xl|TRUE$|FALSE$)/i.test(piece.text) &&
      isWorksheetFunctionName(piece.text)
    ) {
      out += XLETA + piece.text.toUpperCase();
      continue;
    }
    out += piece.text;
  }
  return out;
}
