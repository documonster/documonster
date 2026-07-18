/**
 * Parse a raw input value (typically a string) into a real Date or a
 * fraction-of-day number, driven entirely by the *target cell's own*
 * `numFmt` — not by guessing the shape of the input string.
 *
 * This is the inverse of the render-direction logic in `cell-format.ts`
 * (`formatDate`, `isDateDisplayFormat`, `isTimeOnlyFormat`): instead of
 * turning a value into display text per a format, it turns display text
 * back into a value per that same format's token order.
 */

const MONTHS_LONG = [
  "january",
  "february",
  "march",
  "april",
  "may",
  "june",
  "july",
  "august",
  "september",
  "october",
  "november",
  "december"
];
const MONTHS_SHORT = MONTHS_LONG.map(m => m.slice(0, 3));

type TokenRole = "year2" | "year4" | "month" | "day" | "hour" | "minute" | "second" | "ampm";

interface Token {
  role: TokenRole;
}

/**
 * Strip quoted literal spans (`"..."`), bracketed sections (`[Red]`,
 * `[h]`, etc.), and backslash-escaped single characters — none of these
 * carry date/time token meaning and would otherwise confuse the scanner.
 */
function stripNonTokenSpans(fmt: string): string {
  return fmt
    .replace(/"[^"]*"/g, " ")
    .replace(/\[[^\]]*\]/g, " ")
    .replace(/\\./g, " ");
}

/**
 * Split an Excel format string into per-section format (Excel format codes
 * can have up to 4 `;`-separated sections: positive;negative;zero;text).
 * Only the first (positive) section matters for date/time values.
 */
function firstSection(fmt: string): string {
  return fmt.split(";")[0];
}

/**
 * Scan a format code left-to-right and extract the ordered sequence of
 * date/time tokens it contains. `mm`/`m` is disambiguated as month vs.
 * minute by adjacency to an hour or second token, mirroring the render-side
 * `resolveMonthOrMinute` heuristic.
 */
export function extractTokenOrder(fmt: string): Token[] {
  const cleaned = stripNonTokenSpans(firstSection(fmt));
  const tokens: Token[] = [];
  const pending: { index: number }[] = [];

  const re = /yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|AM\/PM|A\/P/gi;
  let match: RegExpExecArray | null;
  while ((match = re.exec(cleaned)) !== null) {
    const raw = match[0].toLowerCase();
    if (raw === "yyyy") {
      tokens.push({ role: "year4" });
    } else if (raw === "yy") {
      tokens.push({ role: "year2" });
    } else if (raw === "mmmm" || raw === "mmm") {
      tokens.push({ role: "month" });
    } else if (raw === "dd" || raw === "d") {
      tokens.push({ role: "day" });
    } else if (raw === "hh" || raw === "h") {
      tokens.push({ role: "hour" });
    } else if (raw === "ss" || raw === "s") {
      tokens.push({ role: "second" });
    } else if (raw === "am/pm" || raw === "a/p") {
      tokens.push({ role: "ampm" });
    } else if (raw === "mm" || raw === "m") {
      // Ambiguous — resolved to month vs minute in a second pass below,
      // once every token's rough position is known.
      pending.push({ index: tokens.length });
      tokens.push({ role: "month" }); // placeholder, may flip to "minute"
    }
  }

  // Resolve each ambiguous `m`/`mm` occurrence: minute if adjacent (ignoring
  // literal separators, which were already stripped from the scan) to an
  // hour token before it or a second token after it; month otherwise.
  for (const { index } of pending) {
    const prev = tokens[index - 1];
    const next = tokens[index + 1];
    if (prev?.role === "hour" || next?.role === "second" || next?.role === "hour") {
      tokens[index] = { role: "minute" };
    }
  }

  return tokens;
}

function monthIndexFromName(word: string): number | undefined {
  const lower = word.toLowerCase();
  const fullIdx = MONTHS_LONG.indexOf(lower);
  if (fullIdx !== -1) {
    return fullIdx;
  }
  const shortIdx = MONTHS_SHORT.indexOf(lower.slice(0, 3));
  return shortIdx !== -1 ? shortIdx : undefined;
}

/**
 * Split an input string into its constituent alphabetic/numeric parts, in
 * order, ignoring whatever separator characters appear between them. This
 * deliberately does not require the input's separators to match the
 * format's separators — only the *count and order* of value-bearing parts
 * must line up with the format's token order.
 */
function splitInputParts(input: string): string[] {
  return input.match(/[A-Za-z]+|\d+/g) ?? [];
}

export interface ParsedDateTime {
  year?: number;
  month?: number; // 1-12
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
}

/** Trailing time tokens may be omitted from the input (Excel itself treats
 * "09:00" as valid for an `h:mm:ss` cell, defaulting seconds to 0) — but a
 * missing date component (day/month/year) is never safe to guess, so only
 * these roles may be dropped off the end. */
const OMITTABLE_TRAILING_ROLES: ReadonlySet<TokenRole> = new Set([
  "hour",
  "minute",
  "second",
  "ampm"
]);

/**
 * Parse `input` against the token order derived from `fmt`. Returns
 * `undefined` if the input has more parts than the format has tokens, is
 * missing anything but trailing time tokens, or any part fails to
 * parse/validate — callers should leave the original value untouched in
 * that case rather than guess further.
 */
function parseByTokens(tokens: Token[], input: string): ParsedDateTime | undefined {
  if (tokens.length === 0) {
    return undefined;
  }
  const parts = splitInputParts(input);
  if (parts.length > tokens.length) {
    return undefined;
  }
  if (parts.length < tokens.length) {
    const missing = tokens.slice(parts.length);
    if (!missing.every(t => OMITTABLE_TRAILING_ROLES.has(t.role))) {
      return undefined;
    }
  }
  const usedTokens = tokens.slice(0, parts.length);

  const result: ParsedDateTime = {};
  let ampm: "am" | "pm" | undefined;

  for (let i = 0; i < usedTokens.length; i++) {
    const part = parts[i];
    const role = usedTokens[i].role;

    if (role === "ampm") {
      const lower = part.toLowerCase();
      if (lower !== "am" && lower !== "pm") {
        return undefined;
      }
      ampm = lower;
      continue;
    }

    if (role === "month") {
      // The format's own month token (mmm vs mm) only controls how the
      // value later *displays* - Excel's manual-entry recognizer accepts
      // either a name or a number regardless, and stores whichever was
      // typed as a plain numeric month. Match that: try numeric first,
      // fall back to a name lookup.
      if (/^\d+$/.test(part)) {
        const num = Number(part);
        if (num < 1 || num > 12) {
          return undefined;
        }
        result.month = num;
      } else {
        const idx = monthIndexFromName(part);
        if (idx === undefined) {
          return undefined;
        }
        result.month = idx + 1;
      }
      continue;
    }

    if (!/^\d+$/.test(part)) {
      return undefined;
    }
    const num = Number(part);

    switch (role) {
      case "year4":
        result.year = num;
        break;
      case "year2":
        // The format's yy token only controls display width - typing a
        // full 4-digit year into a 2-digit-year cell is still valid input
        // (Excel just displays it truncated to 2 digits). Only apply the
        // 1900/2000 pivot to an actual 1-2 digit year.
        result.year = part.length >= 4 ? num : num <= 49 ? 2000 + num : 1900 + num;
        break;
      case "day":
        if (num < 1 || num > 31) {
          return undefined;
        }
        result.day = num;
        break;
      case "hour":
        if (num < 0 || num > 23) {
          return undefined;
        }
        result.hour = num;
        break;
      case "minute":
        if (num < 0 || num > 59) {
          return undefined;
        }
        result.minute = num;
        break;
      case "second":
        if (num < 0 || num > 59) {
          return undefined;
        }
        result.second = num;
        break;
    }
  }

  if (ampm && result.hour !== undefined) {
    const h = result.hour % 12;
    result.hour = ampm === "pm" ? h + 12 : h;
  }

  return result;
}

/** True if `fmt` renders as a pure time-of-day (no date component). */
export function isPureTimeFormat(tokens: Token[]): boolean {
  const hasDate = tokens.some(
    t => t.role === "day" || t.role === "month" || t.role === "year2" || t.role === "year4"
  );
  const hasTime = tokens.some(t => t.role === "hour" || t.role === "minute" || t.role === "second");
  return hasTime && !hasDate;
}

/** True if `fmt` has at least a date component (may also carry time). */
export function isDateFormat(tokens: Token[]): boolean {
  return tokens.some(
    t => t.role === "day" || t.role === "month" || t.role === "year2" || t.role === "year4"
  );
}

/**
 * Parse `input` against the given cell number-format code. Returns a `Date`
 * for date (or date+time) formats, a fraction-of-day number (0..1) for
 * pure time-of-day formats, or `undefined` when the format carries no
 * date/time meaning or the input doesn't fit the format's token shape.
 */
export function parseValueByFormat(fmt: string, input: string): Date | number | undefined {
  const tokens = extractTokenOrder(fmt);
  if (tokens.length === 0) {
    return undefined;
  }

  const parsed = parseByTokens(tokens, input);
  if (!parsed) {
    return undefined;
  }

  if (isPureTimeFormat(tokens)) {
    const h = parsed.hour ?? 0;
    const m = parsed.minute ?? 0;
    const s = parsed.second ?? 0;
    return (h * 3600 + m * 60 + s) / 86400;
  }

  if (isDateFormat(tokens)) {
    if (parsed.year === undefined || parsed.month === undefined || parsed.day === undefined) {
      return undefined;
    }
    const date = new Date(
      Date.UTC(
        parsed.year,
        parsed.month - 1,
        parsed.day,
        parsed.hour ?? 0,
        parsed.minute ?? 0,
        parsed.second ?? 0
      )
    );
    // Reject overflowed components (e.g. day 31 in a 30-day month) rather
    // than silently rolling over into the next month.
    if (date.getUTCMonth() !== parsed.month - 1 || date.getUTCDate() !== parsed.day) {
      return undefined;
    }
    return date;
  }

  return undefined;
}
