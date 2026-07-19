/**
 * Parse a raw input value (typically a string) into a real Date or a
 * fraction-of-day number, driven entirely by the *target cell's own* `numFmt`
 * — not by guessing the shape of the input string.
 *
 * This is the inverse of the render-direction logic in `cell-format.ts`
 * (`formatDate`, `isDateDisplayFormat`, `isTimeOnlyFormat`): instead of turning
 * a value into display text per a format, it turns display text back into a
 * value per that same format.
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

type FieldRole =
  | "year2"
  | "year4"
  | "month"
  | "day"
  | "hour"
  | "minute"
  | "second"
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
  return role === "month" || role === "ampm" || role === "ampmShort";
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
  "ampm",
  "ampmShort"
]);

/**
 * Take the first (positive) of a format's up-to-four `;`-separated sections,
 * honouring quoted spans and backslash escapes so a `;` inside a literal does
 * not split the section.
 */
function firstSection(fmt: string): string {
  let quoted = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === "\\") {
      i++;
    } else if (ch === '"') {
      quoted = !quoted;
    } else if (ch === ";" && !quoted) {
      return fmt.slice(0, i);
    }
  }
  return fmt;
}

// Sticky (`y`) so it only matches at the exact cursor, and case-insensitive
// (`i`) because Excel format tokens are case-insensitive (`YYYY-MM-DD`,
// `AM/PM`, `am/pm`, `A/P` are all valid and mean the same thing).
const TOKEN_RE = /yyyy|yy|mmmmm|mmmm|mmm|AM\/PM|A\/P|\[h+\]|\[m+\]|\[s+\]|mm|m|dd|d|hh|h|ss|s/iy;

/**
 * Compile a format section into ordered field/literal segments. Consecutive
 * literal characters coalesce into one literal segment; `mm`/`m` is resolved to
 * month vs. minute by adjacency to an hour (before) or second (after) field,
 * matching the render side's `resolveMonthOrMinute`.
 */
export function compileFormat(fmt: string): Segment[] {
  const section = firstSection(fmt);
  const segments: Segment[] = [];
  // Indices (into `segments`) of month/minute placeholders awaiting resolution.
  const ambiguous: number[] = [];

  let literal = "";
  let literalHasLetters = false;
  const flushLiteral = () => {
    if (literal.length > 0) {
      segments.push({ kind: "literal", hasLetters: literalHasLetters });
      literal = "";
      literalHasLetters = false;
    }
  };
  const pushLiteralChar = (ch: string) => {
    literal += ch;
    if (/[A-Za-z]/.test(ch)) {
      literalHasLetters = true;
    }
  };
  const pushField = (role: FieldRole) => {
    flushLiteral();
    segments.push({ kind: "field", role });
  };

  for (let i = 0; i < section.length; ) {
    const ch = section[i];

    // Quoted literal span.
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      const inner = section.slice(i + 1, end === -1 ? section.length : end);
      for (const c of inner) {
        pushLiteralChar(c);
      }
      i = end === -1 ? section.length : end + 1;
      continue;
    }
    // Backslash escape — next char is a literal.
    if (ch === "\\" && i + 1 < section.length) {
      pushLiteralChar(section[i + 1]);
      i += 2;
      continue;
    }
    // Bracketed span: elapsed-time ([h]/[mm]/[s]) is a field; anything else
    // ([Red], [$-409], locale/condition) is display-only — skip entirely.
    if (ch === "[") {
      const end = section.indexOf("]", i + 1);
      const body = section.slice(i + 1, end === -1 ? section.length : end);
      const head = body[0]?.toLowerCase();
      if (/^h+$/i.test(body)) {
        pushField("elapsedHour");
      } else if (/^m+$/i.test(body)) {
        pushField("elapsedMinute");
      } else if (/^s+$/i.test(body)) {
        pushField("elapsedSecond");
      }
      void head;
      i = end === -1 ? section.length : end + 1;
      continue;
    }

    // Date/time token?
    TOKEN_RE.lastIndex = i;
    const m = TOKEN_RE.exec(section);
    if (m && m.index === i) {
      const raw = m[0].toLowerCase();
      switch (raw) {
        case "yyyy":
          pushField("year4");
          break;
        case "yy":
          pushField("year2");
          break;
        case "mmmmm":
        case "mmmm":
        case "mmm":
          pushField("month");
          break;
        case "dd":
        case "d":
          pushField("day");
          break;
        case "hh":
        case "h":
          pushField("hour");
          break;
        case "ss":
        case "s":
          pushField("second");
          break;
        case "am/pm":
          pushField("ampm");
          break;
        case "a/p":
          pushField("ampmShort");
          break;
        case "mm":
        case "m":
          // Provisionally month; resolved after the whole scan.
          flushLiteral();
          ambiguous.push(segments.length);
          segments.push({ kind: "field", role: "month" });
          break;
      }
      i += m[0].length;
      continue;
    }

    // Anything else is a literal separator character.
    pushLiteralChar(ch);
    i++;
  }
  flushLiteral();

  // Resolve each ambiguous m/mm: it is a minute if the nearest field before it
  // is an hour, or the nearest field after it is a second (mirroring the
  // render-side adjacency rule).
  const fieldRoleAt = (segIndex: number, dir: -1 | 1): FieldRole | undefined => {
    for (let k = segIndex + dir; k >= 0 && k < segments.length; k += dir) {
      const seg = segments[k];
      if (seg.kind === "field") {
        return seg.role;
      }
    }
    return undefined;
  };
  for (const idx of ambiguous) {
    const before = fieldRoleAt(idx, -1);
    const after = fieldRoleAt(idx, 1);
    if (before === "hour" || before === "elapsedHour" || after === "second") {
      (segments[idx] as FieldSegment).role = "minute";
    }
  }

  return segments;
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

export interface ParsedDateTime {
  year?: number;
  month?: number; // 1-12
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
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
      if (seg.role === "month") {
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
  if (input.trim() === "") {
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
    return ((parsed.elapsedSeconds ?? 0) + h * 3600 + m * 60 + s) / 86400;
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
      parsed.second ?? 0
    )
  );
  // Reject overflowed components (e.g. day 31 in a 30-day month) rather than
  // silently rolling over into the next month.
  if (date.getUTCMonth() !== parsed.month - 1 || date.getUTCDate() !== parsed.day) {
    return undefined;
  }
  return date;
}
