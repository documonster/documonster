/**
 * Excel number formats, read once.
 *
 * **Why this exists.** A number format such as `\$0.0,,"M"` or `[h]:mm:ss` is a small language,
 * and this codebase used to read it with a separate scanner per consumer: the XLSX/XLSB readers
 * deciding whether a serial is a date, the display formatter, the date-kind classifier, the
 * pivot cache writer, the `TEXT` function and the input parser. Each one stripped a different
 * subset of the constructs that make a character *literal* — quotes, backslash escapes,
 * `_x` padding, `*x` fill, bracket tags — so each one mistook a different literal for a code.
 * `\$0.0,,\M` turned a five-billion-dollar figure into `Invalid Date` on load because one of them
 * read the escaped `M` as a month; `0_m` was a date to one reader and a number to another.
 *
 * Patching each scanner is how they drifted apart in the first place. So there is now one
 * tokenizer: every consumer asks it what a format says, and a character is a code or a literal
 * for all of them or for none.
 *
 * What it models:
 *
 * - **Sections** split on `;`, except inside quotes, brackets or after a backslash.
 * - **Literals**: `"quoted"`, `\x`, and any character that is not a code.
 * - **Layout**: `_x` (space the width of `x`) and `*x` (repeat `x` to fill).
 * - **Bracket tags**: colours, conditions (`[>=100]`), locale/currency (`[$€-407]`), elapsed time
 *   (`[h]`, `[mm]`, `[ss]`) and anything else (`[DBNum1]`), each kept apart so a consumer can act
 *   on the one it needs.
 * - **Date and time codes**, with each `m`/`mm` resolved to month or minute where it stands —
 *   Excel's rule is positional (minute directly after an hour or directly before a second), so
 *   `yyyy-mm-dd hh:mm` holds one of each and no format-wide flag can describe it.
 * - **Number codes**: `0 # ?`, `.`, `,`, `%`, `E+`/`E-`, `/` and `@`, and the word `General`.
 *
 * `e` and `g` (era years in East Asian locales) are deliberately literals: an unquoted word such
 * as `zero` in `0;-0;zero` must stay a word, and no locale this library renders uses them.
 * `b` (Buddhist year) is a code, as it was for the XLSX reader before this module existed.
 */

/** English month names, January first: what `mmm`/`mmmm`/`mmmmm` print and what input parsing accepts. */
export const MONTH_NAMES: readonly string[] = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December"
];

/** English weekday names, Sunday first: what `ddd`/`dddd` print and what input parsing accepts. */
export const DAY_NAMES: readonly string[] = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday"
];

/** A comparison in a conditional section such as `[>=100]`. */
export type FormatCompareOp = "=" | "<>" | "<" | "<=" | ">" | ">=";

/** What a date/time code stands for, after `m` has been resolved to month or minute. */
export type FormatDatePart =
  | "year"
  | "buddhistYear"
  | "month"
  | "day"
  | "hour"
  | "minute"
  | "second";

/** One lexical element of a number-format section. */
export type FormatToken =
  /** Text shown as written: a quoted run, an escaped character, or a character that is not a code. */
  | { readonly kind: "literal"; readonly text: string }
  /**
   * An unquoted digit 1–9. It prints as itself, and is kept apart from `literal` because it is also the one
   * place a digit is syntax: the fixed denominator of `# ?/16`. A quoted or escaped digit (`"8"`, `\8`) is a
   * `literal` and never a denominator.
   */
  | { readonly kind: "numeral"; readonly text: string }
  /** `_x`: a space as wide as `x`. */
  | { readonly kind: "pad"; readonly char: string }
  /** `*x`: `x` repeated to fill the cell. */
  | { readonly kind: "fill"; readonly char: string }
  | { readonly kind: "color"; readonly name: string }
  | { readonly kind: "condition"; readonly op: FormatCompareOp; readonly value: number }
  /** `[$…]`: a currency symbol and/or locale id. `symbol` is the text before any `-`. */
  | { readonly kind: "locale"; readonly symbol: string }
  /** A bracket tag with no effect on the rendered text, such as `[DBNum1]`. */
  | { readonly kind: "bracket"; readonly text: string }
  /**
   * `[h]`, `[mm]`, `[ss]`: a total count of units rather than a clock reading. `fraction` is the number of
   * decimal places after `[ss]` (`[ss].00`), and 0 otherwise.
   */
  | {
      readonly kind: "elapsed";
      readonly unit: "h" | "m" | "s";
      readonly width: number;
      readonly fraction: number;
    }
  /**
   * A date or time code. `fraction` is the number of decimal places of a seconds code (`ss.000`); the point
   * and its zeros belong to the code, not to the number layout, and 0 means there are none.
   */
  | {
      readonly kind: "date";
      readonly part: FormatDatePart;
      readonly width: number;
      readonly fraction: number;
    }
  /** `AM/PM` or `A/P`; `lower` when written in lower case, which is how Excel prints it back. */
  | { readonly kind: "ampm"; readonly short: boolean; readonly lower: boolean }
  | { readonly kind: "digit"; readonly char: "0" | "#" | "?" }
  | { readonly kind: "decimal" }
  | { readonly kind: "thousands" }
  | { readonly kind: "percent" }
  | { readonly kind: "exponent"; readonly sign: "+" | "-"; readonly upper: boolean }
  | { readonly kind: "slash" }
  /** `@`: the cell's text. */
  | { readonly kind: "text" }
  | { readonly kind: "general" };

/** One section of a format, with the facts every consumer asks about computed once. */
export interface FormatSection {
  /** The section as written. */
  readonly source: string;
  readonly tokens: readonly FormatToken[];
  /** A year, month or day code. */
  readonly hasDate: boolean;
  /** An hour, minute or second code, or an AM/PM marker. */
  readonly hasTime: boolean;
  /** An elapsed-time tag such as `[h]`. */
  readonly hasElapsed: boolean;
  /** The `@` text placeholder. */
  readonly hasText: boolean;
  /** The `General` keyword. */
  readonly hasGeneral: boolean;
  /** The section's condition, if it carries one. */
  readonly condition: { readonly op: FormatCompareOp; readonly value: number } | undefined;
}

/** A whole format: its sections, and which of them formats text. */
export interface ParsedNumberFormat {
  readonly source: string;
  readonly sections: readonly FormatSection[];
  /**
   * The sections that format numbers, in order. A section carrying `@` is the *text* section and
   * is excluded, which is what makes `0.00;@` a one-section number format rather than one whose
   * negative numbers are rendered as text.
   */
  readonly numberSections: readonly FormatSection[];
  /** The section that formats a text value, or `undefined` when text is shown as is. */
  readonly textSection: FormatSection | undefined;
}

// =============================================================================
// Sections
// =============================================================================

/**
 * Split a format into its `;`-separated sections (positive; negative; zero; text).
 *
 * A `;` inside `"…"`, inside `[…]`, or taken as the argument of `\`, `_` or `*` is part of a section,
 * not a separator — the same rules `tokenizeFormatSection` reads a character by.
 */
export function splitFormatSections(fmt: string): string[] {
  const sections: string[] = [];
  let start = 0;
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (inQuote) {
      if (ch === '"') {
        inQuote = false;
      }
    } else if (inBracket) {
      if (ch === "]") {
        inBracket = false;
      }
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      // Each takes the next character as its argument, whatever it is: `0_;0` pads with the width
      // of `;`, and `0_";0` with the width of `"` — neither ends a section or opens a literal.
      i++;
    } else if (ch === '"') {
      inQuote = true;
    } else if (ch === "[") {
      inBracket = true;
    } else if (ch === ";") {
      sections.push(fmt.slice(start, i));
      start = i + 1;
    }
  }
  sections.push(fmt.slice(start));
  return sections;
}

// =============================================================================
// Tokenizer
// =============================================================================

const COLOR_RE = /^(?:black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)$/i;
const CONDITION_RE = /^(<>|<=|>=|=|<|>)\s*(-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)$/i;
const ELAPSED_RE = /^(h+|m+|s+)$/i;

function bracketToken(body: string): FormatToken {
  if (ELAPSED_RE.test(body)) {
    return {
      kind: "elapsed",
      unit: body[0].toLowerCase() as "h" | "m" | "s",
      width: body.length,
      fraction: 0
    };
  }
  const condition = CONDITION_RE.exec(body);
  if (condition) {
    return { kind: "condition", op: condition[1] as FormatCompareOp, value: Number(condition[2]) };
  }
  if (body.startsWith("$")) {
    const dash = body.indexOf("-");
    return { kind: "locale", symbol: body.slice(1, dash === -1 ? body.length : dash) };
  }
  if (COLOR_RE.test(body)) {
    return { kind: "color", name: body };
  }
  return { kind: "bracket", text: body };
}

/** Length of the run of `letter` (case-insensitive) starting at `i`. */
function runLength(section: string, i: number, letter: string): number {
  let j = i;
  while (j < section.length && section[j].toLowerCase() === letter) {
    j++;
  }
  return j - i;
}

function startsWithIgnoreCase(section: string, i: number, word: string): boolean {
  return section.slice(i, i + word.length).toLowerCase() === word.toLowerCase();
}

/** A provisional token: `m` runs are resolved to month or minute once the whole section is read. */
type RawToken = FormatToken | { readonly kind: "m"; readonly width: number };

/**
 * Read one section into tokens.
 *
 * Every construct that makes a character literal is recognised here and only here, so no consumer
 * can mistake `\M`, `"M"`, `_M` or `*M` for a month.
 */
export function tokenizeFormatSection(section: string): FormatToken[] {
  const raw: RawToken[] = [];
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    const lower = ch.toLowerCase();

    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const stop = end === -1 ? section.length : end;
      raw.push({ kind: "literal", text: section.slice(i + 1, stop) });
      i = stop + 1;
      continue;
    }
    if (ch === "\\") {
      if (i + 1 < section.length) {
        raw.push({ kind: "literal", text: section[i + 1] });
      }
      i += 2;
      continue;
    }
    if (ch === "_" || ch === "*") {
      if (i + 1 < section.length) {
        raw.push(
          ch === "_"
            ? { kind: "pad", char: section[i + 1] }
            : { kind: "fill", char: section[i + 1] }
        );
      }
      i += 2;
      continue;
    }
    if (ch === "[") {
      const end = section.indexOf("]", i + 1);
      if (end === -1) {
        raw.push({ kind: "literal", text: section.slice(i) });
        break;
      }
      raw.push(bracketToken(section.slice(i + 1, end)));
      i = end + 1;
      continue;
    }
    if (startsWithIgnoreCase(section, i, "General")) {
      raw.push({ kind: "general" });
      i += 7;
      continue;
    }
    if (startsWithIgnoreCase(section, i, "AM/PM")) {
      raw.push({ kind: "ampm", short: false, lower: ch === "a" });
      i += 5;
      continue;
    }
    if (startsWithIgnoreCase(section, i, "A/P")) {
      raw.push({ kind: "ampm", short: true, lower: ch === "a" });
      i += 3;
      continue;
    }
    if (lower === "e" && (section[i + 1] === "+" || section[i + 1] === "-")) {
      raw.push({ kind: "exponent", sign: section[i + 1] as "+" | "-", upper: ch === "E" });
      i += 2;
      continue;
    }
    if (lower === "b" && (section[i + 1] === "1" || section[i + 1] === "2")) {
      // `B1`/`B2` switch calendars (Gregorian/Hijri); they print nothing.
      raw.push({ kind: "bracket", text: section.slice(i, i + 2) });
      i += 2;
      continue;
    }
    if (lower === "y" || lower === "d" || lower === "h" || lower === "s" || lower === "b") {
      const width = runLength(section, i, lower);
      const part: FormatDatePart =
        lower === "y"
          ? "year"
          : lower === "b"
            ? "buddhistYear"
            : lower === "d"
              ? "day"
              : lower === "h"
                ? "hour"
                : "second";
      raw.push({ kind: "date", part, width, fraction: 0 });
      i += width;
      continue;
    }
    if (lower === "m") {
      const width = runLength(section, i, "m");
      raw.push({ kind: "m", width });
      i += width;
      continue;
    }
    switch (ch) {
      case "0":
      case "#":
      case "?":
        raw.push({ kind: "digit", char: ch });
        break;
      case ".":
        raw.push({ kind: "decimal" });
        break;
      case ",":
        raw.push({ kind: "thousands" });
        break;
      case "%":
        raw.push({ kind: "percent" });
        break;
      case "/":
        raw.push({ kind: "slash" });
        break;
      case "@":
        raw.push({ kind: "text" });
        break;
      case "1":
      case "2":
      case "3":
      case "4":
      case "5":
      case "6":
      case "7":
      case "8":
      case "9":
        raw.push({ kind: "numeral", text: ch });
        break;
      default:
        raw.push({ kind: "literal", text: ch });
    }
    i++;
  }
  return foldSecondFractions(resolveMonthOrMinute(raw));
}

/**
 * Attach `.0`, `.00`, … after a seconds code to that code.
 *
 * In `ss.00` and `[ss].00` the point and zeros are fractions of a second, not a decimal number layout,
 * and every consumer — the renderer rounding to them, the input parser reading them back — needs to know
 * that. Deciding it here, once, is what keeps the two from reading the same format differently.
 */
function foldSecondFractions(tokens: FormatToken[]): FormatToken[] {
  const out: FormatToken[] = [];
  for (let k = 0; k < tokens.length; k++) {
    const token = tokens[k];
    const isSecond =
      (token.kind === "date" && token.part === "second") ||
      (token.kind === "elapsed" && token.unit === "s");
    if (!isSecond || tokens[k + 1]?.kind !== "decimal") {
      out.push(token);
      continue;
    }
    let zeros = 0;
    for (let j = k + 2; j < tokens.length; j++) {
      const next = tokens[j];
      if (next.kind !== "digit" || next.char !== "0") {
        break;
      }
      zeros++;
    }
    if (zeros === 0) {
      out.push(token);
      continue;
    }
    out.push({ ...token, fraction: zeros });
    k += 1 + zeros;
  }
  return out;
}

/**
 * Excel's rule for `m` and `mm`, measured against Excel itself (33 formats). `mmm` and longer are always
 * months. A shorter run is a minute when
 *
 * - the time code before it, skipping year, day and month-name codes, is an hour (`h dd mm`, `h mmm mm`,
 *   `[h] mm`);
 * - the code right after it is a seconds code (`mm:ss`), but not an elapsed `[ss]` and not across a day or year
 *   code (`mm dd ss` is a month);
 * - or the time code before it is a seconds code and no minute has appeared yet (`ss mm`, but the second
 *   `mm` of `h:mm:ss mm` is a month).
 *
 * Anything else is a month (`dd mm`, `mm h`).
 */
function resolveMonthOrMinute(raw: readonly RawToken[]): FormatToken[] {
  const isCode = (t: RawToken): boolean =>
    t.kind === "m" || t.kind === "date" || t.kind === "elapsed";
  const isCalendar = (t: RawToken): boolean =>
    (t.kind === "m" && t.width >= 3) ||
    (t.kind === "date" && t.part !== "hour" && t.part !== "minute" && t.part !== "second");
  let minuteSeen = false;
  const out: FormatToken[] = [];
  raw.forEach((token, index) => {
    if (token.kind !== "m") {
      out.push(token);
      return;
    }
    if (token.width >= 3) {
      out.push({ kind: "date", part: "month", width: token.width, fraction: 0 });
      return;
    }
    let before: RawToken | undefined;
    for (let k = index - 1; k >= 0; k--) {
      if (isCode(raw[k]) && !isCalendar(raw[k])) {
        before = raw[k];
        break;
      }
    }
    let after: RawToken | undefined;
    for (let k = index + 1; k < raw.length; k++) {
      if (isCode(raw[k])) {
        after = raw[k];
        break;
      }
    }
    const isPlain = (t: RawToken | undefined, part: FormatDatePart): boolean =>
      t?.kind === "date" && t.part === part;
    const minute =
      isPlain(before, "hour") ||
      (before?.kind === "elapsed" && before.unit === "h") ||
      isPlain(after, "second") ||
      (isPlain(before, "second") && !minuteSeen);
    minuteSeen ||= minute;
    out.push({ kind: "date", part: minute ? "minute" : "month", width: token.width, fraction: 0 });
  });
  return out;
}

// =============================================================================
// Parsed formats
// =============================================================================

function describeSection(source: string): FormatSection {
  const tokens = tokenizeFormatSection(source);
  let hasDate = false;
  let hasTime = false;
  let hasElapsed = false;
  let hasText = false;
  let hasGeneral = false;
  let condition: FormatSection["condition"];
  for (const token of tokens) {
    switch (token.kind) {
      case "date":
        if (token.part === "hour" || token.part === "minute" || token.part === "second") {
          hasTime = true;
        } else {
          hasDate = true;
        }
        break;
      case "ampm":
        hasTime = true;
        break;
      case "elapsed":
        hasElapsed = true;
        break;
      case "text":
        hasText = true;
        break;
      case "general":
        hasGeneral = true;
        break;
      case "condition":
        condition ??= { op: token.op, value: token.value };
        break;
    }
  }
  return { source, tokens, hasDate, hasTime, hasElapsed, hasText, hasGeneral, condition };
}

/** Formats seen in a workbook are few and asked about very often; past this the cache restarts. */
const PARSE_CACHE_LIMIT = 1024;
const parseCache = new Map<string, ParsedNumberFormat>();

/**
 * Parse a whole format. Results are cached: a reader asks about the same handful of formats once
 * per cell.
 */
export function parseNumberFormat(fmt: string): ParsedNumberFormat {
  const cached = parseCache.get(fmt);
  if (cached !== undefined) {
    return cached;
  }
  const sections = splitFormatSections(fmt).map(describeSection);
  let textSection: FormatSection | undefined;
  let numberSections: readonly FormatSection[] = sections;
  if (sections.length >= 4) {
    textSection = sections[3];
    numberSections = sections.slice(0, 3);
  } else if (sections[sections.length - 1].hasText) {
    textSection = sections[sections.length - 1];
    numberSections = sections.length > 1 ? sections.slice(0, -1) : [];
  }
  const parsed: ParsedNumberFormat = { source: fmt, sections, numberSections, textSection };
  if (parseCache.size >= PARSE_CACHE_LIMIT) {
    parseCache.clear();
  }
  parseCache.set(fmt, parsed);
  return parsed;
}

/** What a format's first section says about dates and times. */
export interface NumberFormatFacets {
  /** A year, month or day code. */
  readonly date: boolean;
  /** An hour, minute or second code, or AM/PM. */
  readonly time: boolean;
  /** An elapsed-time tag such as `[h]`. */
  readonly elapsed: boolean;
  /** The `@` text placeholder. */
  readonly text: boolean;
  /** `General`, or no format at all. */
  readonly general: boolean;
}

/**
 * The date and time facts of a format, read from its **first** section — the one that applies to
 * a positive serial, and therefore to every date. `";;;dd"` hides numbers; it is not a date format.
 */
export function numberFormatFacets(fmt: string | null | undefined): NumberFormatFacets {
  if (fmt === null || fmt === undefined || fmt.trim() === "") {
    return { date: false, time: false, elapsed: false, text: false, general: true };
  }
  const first = parseNumberFormat(fmt).sections[0];
  return {
    date: first.hasDate,
    time: first.hasTime,
    elapsed: first.hasElapsed,
    text: first.hasText,
    general: first.hasGeneral
  };
}

/**
 * Whether a cell carrying this format holds a date: its first section names a calendar or clock code
 * outside every literal, and is not a text format.
 *
 * This is the rule every reader applies when deciding whether a stored serial becomes a `Date`.
 *
 * **An elapsed tag on its own does not count.** `[h]` or `[ss]` is a quantity of time, and a `Date` would
 * misstate it — and could not hold it past a few thousand years. `[h]:mm:ss` still counts, through its
 * clock codes, which is the reading these readers have always given both forms.
 */
export function isDateFmt(fmt: string | null | undefined): boolean {
  if (!fmt) {
    return false;
  }
  // Read the cached section directly: the readers ask this once per numeric cell.
  const first = parseNumberFormat(fmt).sections[0];
  return !first.hasText && (first.hasDate || first.hasTime);
}

/** Whether a format is `General` — or empty, which Excel treats the same way. */
export function isGeneralFormat(fmt: string | null | undefined): boolean {
  if (fmt === null || fmt === undefined || fmt.trim() === "") {
    return true;
  }
  const { sections } = parseNumberFormat(fmt);
  return (
    sections.length === 1 &&
    sections[0].hasGeneral &&
    sections[0].tokens.every(
      t =>
        t.kind === "general" ||
        t.kind === "color" ||
        (t.kind === "locale" && t.symbol === "") ||
        (t.kind === "literal" && t.text.trim() === "")
    )
  );
}
