/**
 * The `IF` field: `IF operand operator operand "true text" "false text"`,
 * evaluated as Word does. Shared by field update and mail merge, which
 * evaluates again an `IF` whose operands were merged.
 */

import { formatFieldResult } from "@word/core/field-format";

/**
 * The result of an `IF` field whose code — nested fields replaced by what they
 * show — is `instruction`, formatted by its general switches; `undefined` when
 * it is not an `IF` Word could evaluate.
 */
export function evaluateIfField(instruction: string): string | undefined {
  const m = /^\s*IF\b/i.exec(instruction);
  if (!m) {
    return undefined;
  }
  const parsed = parseIfField(instruction.slice(m[0].length));
  if (!parsed) {
    return undefined;
  }
  const holds = evaluateCondition(parsed.left, parsed.operator, parsed.right, parsed.rightQuoted);
  return formatFieldResult(holds ? parsed.trueText : parsed.falseText, instruction);
}

/** Parse IF field: IF expr1 op expr2 "trueText" "falseText" */
function parseIfField(args: string): {
  left: string;
  operator: string;
  right: string;
  /** Whether the right operand was quoted (`"…"`): `?` and `*` then match as wildcards. */
  rightQuoted: boolean;
  trueText: string;
  falseText: string;
} | null {
  // Hand-rolled parser used in place of the previous chained regex
  // (`/^"?([^"=<>!]*?)"?\s*(<=|>=|<>|=|<|>)\s*"?([^"]*?)"?\s+"([^"]*)"\s+"([^"]*)"/`).
  // CodeQL flagged that regex as polynomial-redos. The grammar is also
  // permissive in ways the regex captured implicitly: real Word IF
  // fields contain operands such as `MERGEFIELD foo` (with internal
  // whitespace) and operands wrapped in quotes. The scanner below
  // mirrors the regex's accepted shape:
  //
  //   args := SP* leftOperand SP* op SP* rightOperand SP+ "trueText" SP+ "falseText" …
  //
  // where `leftOperand` runs up to the first comparison operator that
  // is not inside a quoted span, and `rightOperand` runs up to the
  // first `"` that begins the trueText literal.

  // 1. Find the comparison operator outside any quoted span.
  const opPos = findIfOperator(args, 0);
  if (!opPos) {
    return null;
  }

  // 2. Left operand: everything before the operator, with surrounding
  //    whitespace and outer quotes stripped.
  const left = stripOuterQuotes(args.slice(0, opPos.start).trim());

  // 3. Right operand: text between the operator and the first quoted
  //    literal, with surrounding whitespace and outer quotes stripped.
  let cursor = opPos.next;
  // Scan to the next `"` that is not the immediate value-quoted operand.
  // We have to be careful: the right operand itself may be quoted, e.g.
  // `1 = "bar" "y" "n"`. To match the previous regex we adopt: skip
  // whitespace, optionally consume one quoted span as the right operand,
  // otherwise consume up to the next whitespace+`"` boundary.
  cursor = skipSpaces(args, cursor);
  let right: string;
  let rightQuoted = false;
  if (args.charCodeAt(cursor) === 0x22 /* '"' */) {
    const close = args.indexOf('"', cursor + 1);
    if (close < 0) {
      return null;
    }
    right = args.slice(cursor + 1, close);
    rightQuoted = true;
    cursor = close + 1;
  } else {
    // Read until the next `"` (which begins the trueText literal).
    const nextQuote = args.indexOf('"', cursor);
    if (nextQuote < 0) {
      return null;
    }
    right = args.slice(cursor, nextQuote).trim();
    cursor = nextQuote;
  }

  // 4. trueText (required quoted string).
  cursor = skipSpaces(args, cursor);
  const trueRead = readQuotedString(args, cursor);
  if (!trueRead) {
    return null;
  }
  cursor = skipSpaces(args, trueRead.next);

  // 5. falseText (required quoted string).
  const falseRead = readQuotedString(args, cursor);
  if (!falseRead) {
    return null;
  }

  return {
    left,
    operator: opPos.value,
    right,
    rightQuoted,
    trueText: trueRead.value,
    falseText: falseRead.value
  };
}

/**
 * Find the first IF-field comparison operator (`<=`, `>=`, `<>`, `=`,
 * `<`, `>`) starting at `from`, skipping over any quoted (`"…"`) spans
 * so that operators inside operand strings are not mistaken for the
 * top-level comparator.
 *
 * Bare `!` is reported as "no operator" rather than absorbed into the
 * preceding operand: the previous regex excluded `!` from the left
 * operand character class, so `1 != 1 …` was rejected outright. We
 * preserve that rejection here to avoid silently parsing `!=` (not a
 * Word IF-field operator) as `=` with `!` glued to the left operand.
 */
function findIfOperator(
  s: string,
  from: number
): { start: number; next: number; value: string } | null {
  const n = s.length;
  let i = from;
  while (i < n) {
    const c = s.charCodeAt(i);
    if (c === 0x22 /* '"' */) {
      // Skip quoted span.
      const close = s.indexOf('"', i + 1);
      if (close < 0) {
        return null;
      }
      i = close + 1;
      continue;
    }
    if (c === 0x21 /* '!' */) {
      // Reject `!` outside quotes — matches the previous regex behaviour.
      return null;
    }
    if (c === 0x3c /* '<' */) {
      const next = s.charCodeAt(i + 1);
      if (next === 0x3d) {
        return { start: i, next: i + 2, value: "<=" };
      }
      if (next === 0x3e) {
        return { start: i, next: i + 2, value: "<>" };
      }
      return { start: i, next: i + 1, value: "<" };
    }
    if (c === 0x3e /* '>' */) {
      if (s.charCodeAt(i + 1) === 0x3d) {
        return { start: i, next: i + 2, value: ">=" };
      }
      return { start: i, next: i + 1, value: ">" };
    }
    if (c === 0x3d /* '=' */) {
      return { start: i, next: i + 1, value: "=" };
    }
    i++;
  }
  return null;
}

/**
 * Strip a single matched pair of outer quotes (`"…"`) from a trimmed
 * operand. Mirrors the implicit `"?…"?` shape of the original regex.
 */
function stripOuterQuotes(s: string): string {
  if (s.length >= 2 && s.charCodeAt(0) === 0x22 && s.charCodeAt(s.length - 1) === 0x22) {
    return s.slice(1, -1);
  }
  return s;
}

function skipSpaces(s: string, from: number): number {
  const n = s.length;
  let i = from;
  while (i < n) {
    const c = s.charCodeAt(i);
    if (c !== 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) {
      break;
    }
    i++;
  }
  return i;
}

/** Read a quoted (`"…"`) string starting at `from`, or return null. */
function readQuotedString(s: string, from: number): { value: string; next: number } | null {
  if (s.charCodeAt(from) !== 0x22 /* '"' */) {
    return null;
  }
  const close = s.indexOf('"', from + 1);
  if (close < 0) {
    return null;
  }
  return { value: s.slice(from + 1, close), next: close + 1 };
}

/**
 * Compare two operands: as numbers when both are, otherwise as text, where
 * `=` and `<>` against a quoted operand match `?` (one character) and `*`
 * (any run) as wildcards.
 */
function evaluateCondition(
  left: string,
  operator: string,
  right: string,
  rightQuoted: boolean
): boolean {
  const a = asNumber(left);
  const b = asNumber(right);
  if (a !== undefined && b !== undefined) {
    return compareWith(operator, a, b);
  }
  if ((operator === "=" || operator === "<>") && rightQuoted && /[?*]/.test(right)) {
    const pattern = new RegExp(
      `^${right
        .replace(/[.+^${}()|[\]\\]/g, "\\$&")
        .replace(/\?/g, ".")
        .replace(/\*/g, ".*")}$`,
      "s"
    );
    return pattern.test(left) === (operator === "=");
  }
  return compareWith(operator, left, right);
}

function asNumber(s: string): number | undefined {
  const t = s.trim();
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(t)) {
    return undefined;
  }
  return Number(t);
}

function compareWith<T extends number | string>(operator: string, a: T, b: T): boolean {
  switch (operator) {
    case "=":
      return a === b;
    case "<>":
      return a !== b;
    case "<":
      return a < b;
    case ">":
      return a > b;
    case "<=":
      return a <= b;
    case ">=":
      return a >= b;
    default:
      return false;
  }
}
