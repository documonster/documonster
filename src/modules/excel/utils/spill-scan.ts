/**
 * The tokens of a formula that the spill-range operator `A1#` and its file form `_xlfn.ANCHORARRAY(A1)` are made of —
 * cell references, sheet prefixes, the operator itself, function names and parentheses — with their spans, read the
 * way `@formula/syntax/tokenizer` reads them.
 *
 * **Why not the tokenizer.** `toFileFormula` runs on every formula every writer stores, including a write-only XLSX
 * export, and that bundle must not carry the formula engine's 30 KB tokenizer (`scripts/treeshake-verify.ts`, the
 * write-only and issue #241 scenarios). This scanner follows the tokenizer's branches in its order, emitting only what
 * the operator needs and a placeholder for everything else, and `spill-scan.test.ts` holds it to `tokenizeWithSpans`
 * over a corpus: the same cell references, sheet prefixes, operators and calls at the same spans, and `undefined`
 * exactly where the tokenizer throws.
 */

export const enum ScanKind {
  /** A sheet prefix — `Sheet1!`, `'My Sheet'!`, `S1:S3!`. */
  Sheet,
  Cell,
  /** The spill-range operator `#` after a cell reference. */
  Spill,
  /** A function name directly followed by `(`; `name` is upper-cased. */
  Function,
  Open,
  Close,
  /** The implicit-intersection operator `@`, a prefix. */
  At,
  /** A unary `+` or `-`; `name` is the operator. */
  Prefix,
  /** A binary operator — `+ - * / ^ & = <> < > <= >=`; `name` is the operator. */
  Infix,
  /** The postfix `%`. */
  Percent,
  /** A standalone range operator `:` (one the tokenizer does not fold into a range). */
  Colon,
  /** The intersection operator: whitespace between two references. */
  Intersect,
  Comma,
  OpenBrace,
  CloseBrace,
  Semicolon,
  /** An operand: a number, string, boolean, error, range, name, structured or external reference. */
  Other
}

export interface ScanToken {
  readonly kind: ScanKind;
  readonly start: number;
  readonly end: number;
  readonly name?: string;
  /** What the tokenizer's `lastTokenIsValue` answers for the token. */
  readonly isValue: boolean;
  /** What the tokenizer's `lastTokenIsRef` answers for the token. */
  readonly isRef: boolean;
}

const UNICODE_LETTER = /\p{L}/u;
const UNICODE_MARK = /\p{M}/u;

function isDigit(ch: string | undefined): boolean {
  return ch !== undefined && ch >= "0" && ch <= "9";
}

function isAsciiAlpha(ch: string | undefined): boolean {
  return ch !== undefined && ((ch >= "A" && ch <= "Z") || (ch >= "a" && ch <= "z"));
}

function isAlpha(ch: string | undefined): boolean {
  if (ch === undefined) {
    return false;
  }
  if (isAsciiAlpha(ch)) {
    return true;
  }
  return ch.charCodeAt(0) >= 128 && UNICODE_LETTER.test(ch);
}

function isWordPart(ch: string | undefined): boolean {
  return (
    ch !== undefined &&
    (isAlpha(ch) ||
      isDigit(ch) ||
      ch === "_" ||
      ch === "." ||
      ch === "$" ||
      (ch.charCodeAt(0) >= 0x300 && UNICODE_MARK.test(ch)))
  );
}

function isWhitespace(ch: string | undefined): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
}

/** The error literals the tokenizer knows, longest first. */
const ERROR_LITERALS = [
  "#GETTING_DATA",
  "#BLOCKED!",
  "#CONNECT!",
  "#UNKNOWN!",
  "#SPILL!",
  "#VALUE!",
  "#FIELD!",
  "#DIV/0!",
  "#NAME?",
  "#NULL!",
  "#CALC!",
  "#BUSY!",
  "#REF!",
  "#NUM!",
  "#N/A"
].sort((a, b) => b.length - a.length);

function errorLiteralEnd(text: string, at: number): number | undefined {
  for (const literal of ERROR_LITERALS) {
    if (text.slice(at, at + literal.length).toUpperCase() === literal) {
      return at + literal.length;
    }
  }
  return undefined;
}

function isCellRef(word: string): boolean {
  const match = /^\$?([A-Za-z]{1,3})\$?([0-9]+)$/.exec(word);
  if (match === null) {
    return false;
  }
  const column = match[1]!.toUpperCase();
  const row = Number.parseInt(match[2]!, 10);
  return (column.length < 3 || column <= "XFD") && row >= 1 && row <= 1048576;
}

function looksLikeCellRefStart(text: string, at: number): boolean {
  let j = at;
  if (text[j] === "$") {
    j++;
  }
  if (isAsciiAlpha(text[j])) {
    let letters = 0;
    while (isAsciiAlpha(text[j]) && letters < 4) {
      j++;
      letters++;
    }
    return letters <= 3;
  }
  return isDigit(text[j]);
}

function skipHorizontal(text: string, at: number): number {
  let i = at;
  while (text[i] === " " || text[i] === "\t") {
    i++;
  }
  return i;
}

/** The index of the next unescaped `]` from `at` (`'` escapes the next character), or the end. */
function escapedClose(text: string, at: number): number {
  let i = at;
  while (i < text.length && text[i] !== "]") {
    i += text[i] === "'" && i + 1 < text.length ? 2 : 1;
  }
  return i;
}

/** Past the `]` of a bracketed item whose content starts at `at`. */
function bracketedItemEnd(text: string, at: number): number {
  const close = escapedClose(text, at);
  return close < text.length ? close + 1 : close;
}

/**
 * The end of a structured reference's brackets starting at `at` (`[`), as the tokenizer's
 * `parseStructuredRefBrackets` ends them — malformed text included, since where it stops decides what comes next.
 */
function bracketEnd(text: string, at: number): number {
  const len = text.length;
  let i = skipHorizontal(text, at + 1);
  if (text[i] === "@") {
    i = skipHorizontal(text, i + 1);
    if (text[i] === "[") {
      i = bracketedItemEnd(text, i + 1);
      const afterFirst = skipHorizontal(text, i);
      if (text[afterFirst] === ":") {
        const second = skipHorizontal(text, afterFirst + 1);
        if (text[second] === "[") {
          i = bracketedItemEnd(text, second + 1);
        }
      }
    } else if (i < len && text[i] !== "]") {
      i = escapedClose(text, i);
    }
    i = skipHorizontal(text, i);
    return text[i] === "]" ? i + 1 : i;
  }
  if (text[i] === "[") {
    while (text[i] === "[") {
      i = skipHorizontal(text, bracketedItemEnd(text, i + 1));
      if (text[i] === ":") {
        i = skipHorizontal(text, i + 1);
        if (text[i] === "[") {
          i = bracketedItemEnd(text, i + 1);
        }
        i = skipHorizontal(text, i);
      } else if (text[i] === ",") {
        i = skipHorizontal(text, i + 1);
      }
    }
    i = skipHorizontal(text, i);
    return text[i] === "]" ? i + 1 : i;
  }
  if (text[i] === "#") {
    while (i < len && text[i] !== "]") {
      i++;
    }
    return text[i] === "]" ? i + 1 : i;
  }
  return bracketedItemEnd(text, i);
}

/** The end of a `'…'` run starting at `at` (the opening quote), `''` being an escaped quote. */
function quotedEnd(text: string, at: number): number {
  let i = at + 1;
  while (i < text.length) {
    if (text[i] === "'") {
      if (text[i + 1] === "'") {
        i += 2;
        continue;
      }
      return i + 1;
    }
    i++;
  }
  return i;
}

/** Characters that end an external reference's tail, as the tokenizer stops at them. */
const EXTERNAL_STOP = new Set([
  "+",
  "-",
  "*",
  "/",
  "^",
  "&",
  "=",
  "<",
  ">",
  ")",
  ",",
  " ",
  "}",
  ":",
  "%",
  ";",
  "{"
]);

/**
 * The tokens of `text`, or `undefined` where `tokenizeWithSpans` throws — an unterminated string, an unknown error
 * value, a character no token starts with.
 */
export function scanFormula(text: string): ScanToken[] | undefined {
  const tokens: ScanToken[] = [];
  const len = text.length;
  const lead = /^\s*=/.exec(text);
  let i = lead === null ? 0 : lead[0].length;
  const last = (): ScanToken | undefined => tokens[tokens.length - 1];
  const push = (
    kind: ScanKind,
    start: number,
    end: number,
    flags: { value?: boolean; ref?: boolean; name?: string } = {}
  ): void => {
    tokens.push({
      kind,
      start,
      end,
      isValue: flags.value === true,
      isRef: flags.ref === true,
      ...(flags.name === undefined ? {} : { name: flags.name })
    });
  };
  const value = { value: true };
  const reference = { value: true, ref: true };

  while (i < len) {
    const start = i;
    const ch = text[i]!;

    if (isWhitespace(ch)) {
      while (isWhitespace(text[i])) {
        i++;
      }
      const next = text[i];
      if (
        i < len &&
        last()?.isRef === true &&
        (isAlpha(next) ||
          next === "$" ||
          next === "_" ||
          next === "'" ||
          next === "[" ||
          next === "(")
      ) {
        push(ScanKind.Intersect, start, i);
      }
      continue;
    }

    if (ch === '"') {
      i++;
      let closed = false;
      while (i < len) {
        if (text[i] === '"') {
          if (text[i + 1] === '"') {
            i += 2;
            continue;
          }
          closed = true;
          i++;
          break;
        }
        i++;
      }
      if (!closed) {
        return undefined;
      }
      push(ScanKind.Other, start, i, value);
      continue;
    }

    if (ch === "#") {
      if (last()?.kind === ScanKind.Cell && !isWhitespace(text[i - 1])) {
        i++;
        push(ScanKind.Spill, start, i, reference);
        continue;
      }
      const end = errorLiteralEnd(text, i);
      if (end === undefined) {
        return undefined;
      }
      i = end;
      push(ScanKind.Other, start, i, value);
      continue;
    }

    if (ch === "[" && last()?.isValue !== true) {
      let j = start + 1;
      while (text[j] === " " || text[j] === "\t") {
        j++;
      }
      const first = text[j];
      if (first !== "@" && first !== "#" && first !== "[") {
        const close = text.indexOf("]", start + 1);
        let external = false;
        if (close !== -1 && close + 1 < len) {
          const p = close + 1;
          if (text[p] === "'") {
            external = text[quotedEnd(text, p)] === "!";
          } else {
            let q = p;
            while (isWordPart(text[q])) {
              q++;
            }
            external = q > p && text[q] === "!";
          }
        }
        if (external) {
          i = close + 1;
          while (i < len && !EXTERNAL_STOP.has(text[i]!)) {
            i++;
          }
          if (text[i] === ":") {
            let q = i + 1;
            if (text[q] === "$") {
              q++;
            }
            const columnStart = q;
            while (isAsciiAlpha(text[q])) {
              q++;
            }
            if (q > columnStart) {
              if (text[q] === "$") {
                q++;
              }
              const rowStart = q;
              while (isDigit(text[q])) {
                q++;
              }
              if (q > rowStart) {
                i = q;
              }
            }
          }
          push(ScanKind.Other, start, i, value);
          continue;
        }
      }
    }

    if (ch === "{" || ch === ";") {
      i++;
      push(ch === "{" ? ScanKind.OpenBrace : ScanKind.Semicolon, start, i);
      continue;
    }
    if (ch === "}") {
      i++;
      push(ScanKind.CloseBrace, start, i, value);
      continue;
    }

    if (isDigit(ch) || (ch === "." && isDigit(text[i + 1]))) {
      while (isDigit(text[i])) {
        i++;
      }
      if (i > start && ch !== "." && text[i] === ":") {
        const first = Number.parseInt(text.slice(start, i), 10);
        if (first >= 1 && first <= 1048576) {
          let j = i + 1;
          if (text[j] === "$") {
            j++;
          }
          const rowStart = j;
          while (isDigit(text[j])) {
            j++;
          }
          const second = Number.parseInt(text.slice(rowStart, j), 10);
          if (j > rowStart && second >= 1 && second <= 1048576) {
            i = j;
            push(ScanKind.Other, start, i, reference);
            continue;
          }
        }
      }
      if (text[i] === "." && isDigit(text[i + 1])) {
        i++;
        while (isDigit(text[i])) {
          i++;
        }
      }
      if (text[i] === "E" || text[i] === "e") {
        let j = i + 1;
        if (text[j] === "+" || text[j] === "-") {
          j++;
        }
        if (isDigit(text[j])) {
          i = j;
          while (isDigit(text[i])) {
            i++;
          }
        }
      }
      push(ScanKind.Other, start, i, value);
      continue;
    }

    if (ch === "'") {
      i = quotedEnd(text, i);
      if (text[i] === "!") {
        i++;
      }
      push(ScanKind.Sheet, start, i);
      continue;
    }

    if (ch === "[") {
      i = bracketEnd(text, i);
      push(ScanKind.Other, start, i, reference);
      continue;
    }

    if (isAlpha(ch) || ch === "$" || ch === "_" || ch === "\\") {
      if (ch === "\\") {
        i++;
      }
      while (isWordPart(text[i])) {
        i++;
      }
      const word = text.slice(start, i);
      if (text[i] === "!") {
        i++;
        push(ScanKind.Sheet, start, i);
        continue;
      }
      if (last()?.kind !== ScanKind.Sheet && text[i] === ":") {
        let j = i + 1;
        while (isWordPart(text[j])) {
          j++;
        }
        if (j > i + 1 && text[j] === "!" && j + 1 < len && looksLikeCellRefStart(text, j + 1)) {
          i = j + 1;
          push(ScanKind.Sheet, start, i);
          continue;
        }
      }
      if (text[i] === "[") {
        i = bracketEnd(text, i);
        push(ScanKind.Other, start, i, reference);
        continue;
      }
      if (text[i] === "(") {
        push(ScanKind.Function, start, i, { name: word.toUpperCase() });
        continue;
      }
      const upper = word.toUpperCase();
      if (upper === "TRUE" || upper === "FALSE") {
        push(ScanKind.Other, start, i, value);
        continue;
      }
      if (isCellRef(word)) {
        if (text[i] === ":") {
          let j = i + 1;
          while (isWordPart(text[j])) {
            j++;
          }
          if (isCellRef(text.slice(i + 1, j))) {
            i = j;
            push(ScanKind.Other, start, i, reference);
            continue;
          }
        }
        push(ScanKind.Cell, start, i, reference);
        continue;
      }
      if (text[i] === ":") {
        const columns = /^\$?[A-Za-z]{1,3}$/.test(word) ? columnRangeEnd(text, word, i) : undefined;
        if (columns !== undefined) {
          i = columns;
          push(ScanKind.Other, start, i, reference);
          continue;
        }
        const rows = /^\$\d+$/.test(word) ? rowRangeEnd(text, word, i) : undefined;
        if (rows !== undefined) {
          i = rows;
          push(ScanKind.Other, start, i, reference);
          continue;
        }
      }
      push(ScanKind.Other, start, i, reference);
      continue;
    }

    if (ch === "(") {
      i++;
      push(ScanKind.Open, start, i);
      continue;
    }
    if (ch === ")") {
      i++;
      push(ScanKind.Close, start, i, reference);
      continue;
    }
    if (ch === "%") {
      i++;
      push(ScanKind.Percent, start, i, value);
      continue;
    }
    if (ch === "<" || ch === ">") {
      i += text[i + 1] === "=" || (ch === "<" && text[i + 1] === ">") ? 2 : 1;
      push(ScanKind.Infix, start, i, { name: text.slice(start, i) });
      continue;
    }
    if (ch === "+" || ch === "-") {
      // Binary after a value, as the tokenizer decides it; unary otherwise.
      const binary = last()?.isValue === true;
      i++;
      push(binary ? ScanKind.Infix : ScanKind.Prefix, start, i, { name: ch });
      continue;
    }
    if ("=&^*/".includes(ch)) {
      i++;
      push(ScanKind.Infix, start, i, { name: ch });
      continue;
    }
    if (ch === "@" || ch === "," || ch === ":") {
      i++;
      push(ch === "@" ? ScanKind.At : ch === "," ? ScanKind.Comma : ScanKind.Colon, start, i);
      continue;
    }
    if (ch === "!" && i + 1 < len && looksLikeCellRefStart(text, i + 1)) {
      i++;
      continue;
    }
    return undefined;
  }
  return tokens;
}

/** The end of a whole-column range `A:B` whose first column is `word`, ending before `colon`; or `undefined`. */
function columnRangeEnd(text: string, word: string, colon: number): number | undefined {
  const first = word.replace("$", "").toUpperCase();
  if (first.length === 3 && first > "XFD") {
    return undefined;
  }
  let i = colon + 1;
  const start = i;
  if (text[i] === "$") {
    i++;
  }
  const letters = i;
  while (isAsciiAlpha(text[i])) {
    i++;
  }
  if (isDigit(text[i]) || i === letters) {
    return undefined;
  }
  const second = /^\$?([A-Za-z]{1,3})$/.exec(text.slice(start, i))?.[1]?.toUpperCase();
  return second === undefined || (second.length === 3 && second > "XFD") ? undefined : i;
}

/** The end of an absolute whole-row range `$1:$5` whose first row is `word`, ending before `colon`; or `undefined`. */
function rowRangeEnd(text: string, word: string, colon: number): number | undefined {
  const first = Number.parseInt(word.slice(1), 10);
  if (first < 1 || first > 1048576) {
    return undefined;
  }
  let i = colon + 1;
  if (text[i] === "$") {
    i++;
  }
  const digits = i;
  while (isDigit(text[i])) {
    i++;
  }
  const second = Number.parseInt(text.slice(digits, i), 10);
  return i > digits && second >= 1 && second <= 1048576 ? i : undefined;
}

// =============================================================================
// Operand extents — the parser's binding powers over scanned tokens
// =============================================================================

/** The power `@` parses its operand at (`@A1:A3*2` is `(@A1:A3)*2`, `@A1^2` is `@(A1^2)`). */
export const AT_POWER = 55;

const PREFIX_POWER = 77;
const PERCENT_POWER = 75;
const SPILL_POWER = 110;
const INTERSECT_POWERS: readonly [number, number] = [80, 81];
const COLON_POWERS: readonly [number, number] = [100, 101];

/** The `[left, right]` binding powers of a binary operator, as `@formula/syntax/parser` gives them. */
function infixPowers(op: string | undefined): readonly [number, number] {
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
    default:
      return [0, 0];
  }
}

/** The index past the bracket that closes the one at `at` (`(` or `{`), or `undefined` when none does. */
function closing(tokens: readonly ScanToken[], at: number): number | undefined {
  const open = tokens[at]!.kind;
  const close = open === ScanKind.Open ? ScanKind.Close : ScanKind.CloseBrace;
  let depth = 0;
  for (let i = at; i < tokens.length; i++) {
    const kind = tokens[i]!.kind;
    if (kind === open) {
      depth++;
    } else if (kind === close && --depth === 0) {
      return i + 1;
    }
  }
  return undefined;
}

function prefixEnd(tokens: readonly ScanToken[], at: number): number | undefined {
  const token = tokens[at];
  switch (token?.kind) {
    case ScanKind.Sheet:
      return prefixEnd(tokens, at + 1);
    case ScanKind.Prefix:
      return operandEnd(tokens, at + 1, PREFIX_POWER);
    case ScanKind.At:
      return operandEnd(tokens, at + 1, AT_POWER);
    case ScanKind.Cell:
    case ScanKind.Other:
      return at + 1;
    case ScanKind.Function: {
      if (tokens[at + 1]?.kind !== ScanKind.Open) {
        return undefined;
      }
      let end = closing(tokens, at + 1);
      // An argument list straight after the `)` applies the call's result: `LAMBDA(x,x)(1)`.
      while (end !== undefined && tokens[end]?.kind === ScanKind.Open) {
        end = closing(tokens, end);
      }
      return end;
    }
    case ScanKind.Open:
    case ScanKind.OpenBrace:
      return closing(tokens, at);
    default:
      return undefined;
  }
}

/**
 * The index past the operand the parser reads from `at` at binding power `power` — what `parseExpr(power)` consumes —
 * or `undefined` when the tokens there do not start one. Mirrors `@formula/syntax/parser`, so `@` is wrapped around
 * exactly what it applies to.
 */
export function operandEnd(
  tokens: readonly ScanToken[],
  at: number,
  power: number
): number | undefined {
  let end = prefixEnd(tokens, at);
  while (end !== undefined && end < tokens.length) {
    const token = tokens[end]!;
    let powers: readonly [number, number] | undefined;
    if (token.kind === ScanKind.Spill || token.kind === ScanKind.Percent) {
      if ((token.kind === ScanKind.Spill ? SPILL_POWER : PERCENT_POWER) < power) {
        break;
      }
      end++;
      continue;
    }
    if (token.kind === ScanKind.Infix) {
      powers = infixPowers(token.name);
    } else if (token.kind === ScanKind.Intersect) {
      powers = INTERSECT_POWERS;
    } else if (token.kind === ScanKind.Colon) {
      powers = COLON_POWERS;
    }
    if (powers === undefined || powers[0] === 0 || powers[0] < power) {
      break;
    }
    end = operandEnd(tokens, end + 1, powers[1]);
  }
  return end;
}
