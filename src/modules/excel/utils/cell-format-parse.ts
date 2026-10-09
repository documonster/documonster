/**
 * Parse a raw input value (typically a string) into a real Date or a
 * fraction-of-day number, driven entirely by the *target cell's own* `numFmt`
 * — not by guessing the shape of the input string.
 *
 * This is the inverse of `@utils/number-format-render`: instead of turning a
 * value into display text per a format, it turns display text back into a
 * value per that same format. Both read the format through the one tokenizer
 * in `@utils/number-format`, so a literal on one side is a literal on the other.
 *
 * ## How it works (and why it is not a heuristic)
 *
 * The format string is *compiled* into an ordered list of segments — either a
 * value **field** (year/month/day/hour/minute/second/AM-PM) or a fixed
 * **literal** (every separator, quoted span and escaped char between fields).
 * The input is then matched against that segment list with a single
 * left-to-right cursor: a field consumes the run of characters of its kind
 * (digits, or letters for names/AM-PM), a literal is matched leniently
 * (any non-field run is accepted as the separator, so `-` in the format still
 * matches `.` in the input — a deliberate, useful tolerance).
 *
 * Because the literals are *anchors* rather than something reverse-engineered
 * out of the input, adding a new literal/quoted/section construct to a format
 * needs no special case here — it just becomes another literal segment. That
 * is what makes this robust across arbitrary formats instead of accreting a
 * patch per newly-discovered format shape.
 */

import {
  DAY_NAMES,
  MONTH_NAMES,
  numberFormatFacets,
  splitFormatSections,
  tokenizeFormatSection
} from "@utils/number-format";

const MONTHS_LONG = MONTH_NAMES.map(m => m.toLowerCase());
const MONTHS_SHORT = MONTHS_LONG.map(m => m.slice(0, 3));
const WEEKDAYS = DAY_NAMES.map(d => d.toLowerCase());
const WEEKDAYS_SHORT = WEEKDAYS.map(d => d.slice(0, 3));

type FieldRole =
  | "year2"
  | "year4"
  | "month"
  | "day"
  | "weekday"
  | "hour"
  | "minute"
  | "second"
  | "secondFraction"
  | "ampm"
  | "ampmShort"
  | "elapsedHour"
  | "elapsedMinute"
  | "elapsedSecond";

interface FieldSegment {
  kind: "field";
  role: FieldRole;
}

interface LiteralSegment {
  kind: "literal";
  /** Whether this literal contains letters (must be consumed from the input's
   *  letter runs) versus pure punctuation/space (skippable). */
  hasLetters: boolean;
}

type Segment = FieldSegment | LiteralSegment;

/** A field consumes letters (names, AM/PM) rather than digits. */
function isAlphaField(role: FieldRole): boolean {
  return role === "month" || role === "weekday" || role === "ampm" || role === "ampmShort";
}

/**
 * Time-tail roles that may be omitted from the *end* of the input. Excel treats
 * `"09:00"` as valid for an `h:mm:ss` cell (seconds default to 0). A missing
 * date component (day/month/year) is never guessed.
 */
const OMITTABLE_TRAILING_ROLES: ReadonlySet<FieldRole> = new Set([
  "hour",
  "minute",
  "second",
  "secondFraction",
  "ampm",
  "ampmShort"
]);

/**
 * Compile a format's first section into ordered field/literal segments.
 *
 * The section is read by the shared number-format tokenizer, the same one that renders it and that
 * decides on load whether a serial is a date — so a character is a field here exactly when it is a
 * code there, and `m`/`mm` is a month or a minute by the same positional rule. Consecutive literal
 * characters coalesce into one literal segment.
 */
export function compileFormat(fmt: string): Segment[] {
  const tokens = tokenizeFormatSection(splitFormatSections(fmt)[0]);
  const segments: Segment[] = [];

  let literal = "";
  const flushLiteral = (): void => {
    if (literal.length > 0) {
      segments.push({ kind: "literal", hasLetters: /\p{L}/u.test(literal) });
      literal = "";
    }
  };
  const pushField = (role: FieldRole): void => {
    flushLiteral();
    segments.push({ kind: "field", role });
  };

  for (const token of tokens) {
    switch (token.kind) {
      case "date":
        switch (token.part) {
          case "year":
            pushField(token.width <= 2 ? "year2" : "year4");
            break;
          case "month":
            pushField("month");
            break;
          case "day":
            pushField(token.width <= 2 ? "day" : "weekday");
            break;
          case "hour":
            pushField("hour");
            break;
          case "minute":
            pushField("minute");
            break;
          case "second":
            pushField("second");
            break;
        }
        break;
      case "elapsed":
        pushField(
          token.unit === "h"
            ? "elapsedHour"
            : token.unit === "m"
              ? "elapsedMinute"
              : "elapsedSecond"
        );
        break;
      case "numeral":
        literal += token.text;
        break;
      case "ampm":
        pushField(token.short ? "ampmShort" : "ampm");
        break;
      case "literal":
        literal += token.text;
        break;
      case "locale":
        literal += token.symbol;
        break;
      case "pad":
        literal += " ";
        break;
      case "digit":
        literal += token.char;
        break;
      case "decimal":
        literal += ".";
        break;
      case "thousands":
        literal += ",";
        break;
      case "percent":
        literal += "%";
        break;
      case "slash":
        literal += "/";
        break;
      default:
        // Fill, colour, condition and other bracket tags are display-only.
        break;
    }
    // `ss.00` and `[ss].00`: the tokenizer has already attached the decimal places to the seconds code.
    if ((token.kind === "date" || token.kind === "elapsed") && token.fraction > 0) {
      pushField("secondFraction");
    }
  }
  flushLiteral();
  return segments;
}

/** A full name or its three-letter abbreviation — not any word that happens to start like one. */
function nameIndex(word: string, full: readonly string[], short: readonly string[]): number {
  const lower = word.toLowerCase();
  const fullIdx = full.indexOf(lower);
  return fullIdx !== -1 ? fullIdx : short.indexOf(lower);
}

function monthIndexFromName(word: string): number | undefined {
  const idx = nameIndex(word, MONTHS_LONG, MONTHS_SHORT);
  return idx !== -1 ? idx : undefined;
}

export interface ParsedDateTime {
  year?: number;
  month?: number; // 1-12
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
  /** The digits typed after the seconds' decimal point. */
  secondFraction?: string;
  elapsedSeconds?: number;
}

/** True if the compiled format carries at least one date field. */
function hasDateField(segments: Segment[]): boolean {
  return segments.some(
    s =>
      s.kind === "field" &&
      (s.role === "day" || s.role === "month" || s.role === "year2" || s.role === "year4")
  );
}

/** True if the compiled format carries at least one time field. */
function hasTimeField(segments: Segment[]): boolean {
  return segments.some(
    s =>
      s.kind === "field" &&
      (s.role === "hour" ||
        s.role === "minute" ||
        s.role === "second" ||
        s.role === "secondFraction" ||
        s.role === "elapsedHour" ||
        s.role === "elapsedMinute" ||
        s.role === "elapsedSecond")
  );
}

const LETTER_RUN = /^[A-Za-z]+/;
const DIGIT_RUN = /^\d+/;

/**
 * Match `input` against the compiled `segments` with a single cursor. Returns
 * the extracted components, or `undefined` if the input does not conform (extra
 * trailing content, a field with no matching run, or a value out of range).
 */
function matchSegments(segments: Segment[], input: string): ParsedDateTime | undefined {
  const result: ParsedDateTime = {};
  let ampm: "am" | "pm" | undefined;
  let pos = 0;

  const skipSeparators = () => {
    // A separator run is any leading non-alphanumeric characters.
    while (pos < input.length && !/[A-Za-z0-9]/.test(input[pos])) {
      pos++;
    }
  };

  for (let s = 0; s < segments.length; s++) {
    const seg = segments[s];

    if (seg.kind === "literal") {
      // Letters in a literal (e.g. the "at" in `"at"`) are consumed from a
      // matching letter run; punctuation/space literals just skip separators.
      if (seg.hasLetters) {
        skipSeparators();
        const lm = LETTER_RUN.exec(input.slice(pos));
        if (lm) {
          pos += lm[0].length;
        }
      } else {
        skipSeparators();
      }
      continue;
    }

    skipSeparators();
    const rest = input.slice(pos);

    if (isAlphaField(seg.role)) {
      // AM/PM accepts letters; month accepts a name OR a number.
      const lm = LETTER_RUN.exec(rest);
      if (seg.role === "month" && !lm) {
        const dm = DIGIT_RUN.exec(rest);
        if (!dm) {
          return remainingAreOptional(segments, s) ? finalize() : undefined;
        }
        const num = Number(dm[0]);
        if (num < 1 || num > 12) {
          return undefined;
        }
        result.month = num;
        pos += dm[0].length;
        continue;
      }
      if (!lm) {
        return remainingAreOptional(segments, s) ? finalize() : undefined;
      }
      const word = lm[0];
      if (seg.role === "weekday") {
        // A weekday name is redundant with the date it sits beside; it must only be a real one.
        if (nameIndex(word, WEEKDAYS, WEEKDAYS_SHORT) === -1) {
          return undefined;
        }
      } else if (seg.role === "month") {
        const idx = monthIndexFromName(word);
        if (idx === undefined) {
          return undefined;
        }
        result.month = idx + 1;
      } else {
        const lower = word.toLowerCase();
        const normalized =
          seg.role === "ampmShort" ? (lower === "a" ? "am" : lower === "p" ? "pm" : lower) : lower;
        if (normalized !== "am" && normalized !== "pm") {
          return undefined;
        }
        ampm = normalized;
      }
      pos += word.length;
      continue;
    }

    // Numeric field.
    const dm = DIGIT_RUN.exec(rest);
    if (!dm) {
      return remainingAreOptional(segments, s) ? finalize() : undefined;
    }
    const digits = dm[0];
    const num = Number(digits);
    switch (seg.role) {
      case "year4":
        result.year = num;
        break;
      case "secondFraction":
        result.secondFraction = digits;
        break;
      case "year2":
        // The `yy` token only controls display width; a typed 4-digit year is
        // taken verbatim, a 1–2(–3) digit year uses the 1900/2000 pivot.
        result.year = digits.length >= 4 ? num : num <= 49 ? 2000 + num : 1900 + num;
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
      case "elapsedHour":
        result.elapsedSeconds = (result.elapsedSeconds ?? 0) + num * 3600;
        break;
      case "elapsedMinute":
        result.elapsedSeconds = (result.elapsedSeconds ?? 0) + num * 60;
        break;
      case "elapsedSecond":
        result.elapsedSeconds = (result.elapsedSeconds ?? 0) + num;
        break;
    }
    pos += digits.length;
  }

  skipSeparators();
  if (pos !== input.length) {
    // Unconsumed trailing content — the input has more than the format allows.
    return undefined;
  }
  return finalize();

  function remainingAreOptional(segs: Segment[], fromIndex: number): boolean {
    for (let k = fromIndex; k < segs.length; k++) {
      const seg = segs[k];
      if (seg.kind === "field" && !OMITTABLE_TRAILING_ROLES.has(seg.role)) {
        return false;
      }
    }
    return true;
  }

  function finalize(): ParsedDateTime | undefined {
    if (ampm !== undefined) {
      if (result.hour === undefined || result.hour < 1 || result.hour > 12) {
        return undefined;
      }
      const h = result.hour % 12;
      result.hour = ampm === "pm" ? h + 12 : h;
    }
    return result;
  }
}

/**
 * Parse `input` against the given cell number-format code. Returns a `Date` for
 * date (or date+time) formats, a fraction-of-day number (0..1) for pure
 * time-of-day formats, or `undefined` when the format carries no date/time
 * meaning or the input does not conform to it.
 */
export function parseValueByFormat(fmt: string, input: string): Date | number | undefined {
  // A text format (`yyyy-mm-dd@`) keeps what is typed as text, as the readers keep its values numbers.
  if (input.trim() === "" || numberFormatFacets(fmt).text) {
    return undefined;
  }
  const segments = compileFormat(fmt);
  const hasDate = hasDateField(segments);
  const hasTime = hasTimeField(segments);
  if (!hasDate && !hasTime) {
    return undefined;
  }

  const parsed = matchSegments(segments, input);
  if (!parsed) {
    return undefined;
  }

  if (!hasDate) {
    // Pure time-of-day → fraction of a day.
    const h = parsed.hour ?? 0;
    const m = parsed.minute ?? 0;
    const s = parsed.second ?? 0;
    const fraction = Number("0." + (parsed.secondFraction ?? "0"));
    return ((parsed.elapsedSeconds ?? 0) + h * 3600 + m * 60 + s + fraction) / 86400;
  }

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
      parsed.second ?? 0,
      // A `Date` holds whole milliseconds: take the first three typed digits, so .9999… stays on its
      // second however many nines follow — converting through a double first rounded it up to 1.
      Number((parsed.secondFraction ?? "").slice(0, 3).padEnd(3, "0"))
    )
  );
  // Reject overflowed components (e.g. day 31 in a 30-day month) rather than
  // silently rolling over into the next month.
  if (date.getUTCMonth() !== parsed.month - 1 || date.getUTCDate() !== parsed.day) {
    return undefined;
  }
  return date;
}
