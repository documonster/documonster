/**
 * A field's code read as Word reads it, and the general switches every field
 * result goes through: `\#` (numeric picture), `\@` (date-time picture) and
 * `\*` (case and number formats). Shared by mail merge and field update, so a
 * `MERGEFIELD`, `REF`, `SEQ`, `IF` or `=` shows its value the same way.
 *
 * Word applies a `\#` or `\@` picture first, then each `\*` in order.
 * `\* MERGEFORMAT` and `\* CHARFORMAT` are about formatting, not text, and
 * change nothing here. A switch that cannot apply — a numeric picture on a
 * value that is not a number, a date picture on one that is not a date, a
 * `\*` format not listed below — leaves the value as it is.
 *
 * `\*` formats: `Upper`, `Lower`, `Caps`, `FirstCap`, `Arabic`, `roman`,
 * `ROMAN`, `alphabetic`, `ALPHABETIC`, `Ordinal`, `Hex`, and in English words
 * `CardText`, `OrdText` and `DollarText` (0 to 999,999, as Word).
 */

import { formatListCounter } from "@word/query/page-field";

/** One token of a field code: a word, a quoted string (unquoted), or a switch. */
export interface FieldToken {
  readonly text: string;
  readonly quoted: boolean;
}

/**
 * The tokens of a field code, split on white space outside quotes. Inside
 * quotes, `\"` is a quote and `\\` a backslash, as in Word.
 */
export function tokenizeFieldCode(code: string): FieldToken[] {
  const out: FieldToken[] = [];
  let i = 0;
  const n = code.length;
  while (i < n) {
    while (i < n && /\s/.test(code[i]!)) {
      i++;
    }
    if (i >= n) {
      break;
    }
    if (code[i] === '"') {
      let text = "";
      i++;
      while (i < n && code[i] !== '"') {
        if (code[i] === "\\" && (code[i + 1] === '"' || code[i + 1] === "\\")) {
          i++;
        }
        text += code[i];
        i++;
      }
      i++;
      out.push({ text, quoted: true });
      continue;
    }
    let text = "";
    while (i < n && !/\s/.test(code[i]!) && code[i] !== '"') {
      text += code[i];
      i++;
    }
    out.push({ text, quoted: false });
  }
  return out;
}

/** The argument of switch `name` (e.g. `"\\b"`) in `tokens`, or `undefined`. */
export function fieldSwitchArgument(
  tokens: readonly FieldToken[],
  name: string
): string | undefined {
  for (let i = 0; i < tokens.length - 1; i++) {
    const t = tokens[i]!;
    if (!t.quoted && t.text.toLowerCase() === name.toLowerCase()) {
      const next = tokens[i + 1]!;
      return next.quoted || !next.text.startsWith("\\") ? next.text : undefined;
    }
  }
  return undefined;
}

/**
 * `value` — a field's result — formatted by the general switches of the field
 * whose code is `code`.
 */
export function formatFieldResult(value: string, code: string): string {
  const tokens = tokenizeFieldCode(code);
  let out = value;
  for (let i = 0; i < tokens.length - 1; i++) {
    const t = tokens[i]!;
    if (t.quoted) {
      continue;
    }
    const arg = tokens[i + 1]!.text;
    if (t.text === "\\#") {
      const n = parseFieldNumber(out);
      if (n !== undefined) {
        out = formatNumberPicture(n, arg);
      }
    } else if (t.text === "\\@") {
      const d = parseFieldDate(out);
      if (d) {
        out = formatDatePicture(d, arg);
      }
    }
  }
  for (let i = 0; i < tokens.length - 1; i++) {
    const t = tokens[i]!;
    if (!t.quoted && t.text === "\\*") {
      out = formatGeneral(out, tokens[i + 1]!.text);
    }
  }
  return out;
}

// =============================================================================
// \* — case and number formats
// =============================================================================

function formatGeneral(value: string, format: string): string {
  switch (format.toLowerCase()) {
    case "upper":
      return value.toUpperCase();
    case "lower":
      return value.toLowerCase();
    case "caps":
      // As Word: each word's first letter upper case and the rest lower case,
      // a word starting after a space or a hyphen (not an apostrophe).
      return value.replace(/(^|[\s-])([^\s-]+)/g, (_, lead: string, word: string) => {
        const first = String.fromCodePoint(word.codePointAt(0)!);
        return lead + first.toUpperCase() + word.slice(first.length).toLowerCase();
      });
    case "firstcap":
      return value.replace(/\S/, c => c.toUpperCase());
  }
  const n = parseFieldNumber(value);
  if (n === undefined) {
    return value;
  }
  const whole = Math.round(n);
  switch (format) {
    case "Arabic":
    case "arabic":
      return String(whole);
    case "roman":
      return formatListCounter(whole, "lowerRoman");
    case "ROMAN":
    case "Roman":
      return formatListCounter(whole, "upperRoman");
    case "alphabetic":
      return formatListCounter(whole, "lowerLetter");
    case "ALPHABETIC":
    case "Alphabetic":
      return formatListCounter(whole, "upperLetter");
    case "Ordinal":
    case "ordinal":
      return ordinal(whole);
    case "Hex":
    case "hex":
      return whole >= 0 ? whole.toString(16).toUpperCase() : value;
  }
  switch (format.toLowerCase()) {
    case "cardtext":
      return cardinalWords(whole) ?? value;
    case "ordtext": {
      const words = cardinalWords(whole);
      return words === undefined ? value : ordinalWords(words);
    }
    case "dollartext": {
      const cents = Math.round(Math.abs(n) * 100);
      const words = n < 0 ? undefined : cardinalWords(Math.floor(cents / 100));
      return words === undefined
        ? value
        : `${words} and ${String(cents % 100).padStart(2, "0")}/100`;
    }
    default:
      return value;
  }
}

const ONES = [
  "zero",
  "one",
  "two",
  "three",
  "four",
  "five",
  "six",
  "seven",
  "eight",
  "nine",
  "ten",
  "eleven",
  "twelve",
  "thirteen",
  "fourteen",
  "fifteen",
  "sixteen",
  "seventeen",
  "eighteen",
  "nineteen"
];
const TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];

/** `n` in English words, as Word's `CardText`; `undefined` outside 0–999,999. */
function cardinalWords(n: number): string | undefined {
  if (n < 0 || n > 999_999) {
    return undefined;
  }
  const below1000 = (k: number): string => {
    const parts: string[] = [];
    if (k >= 100) {
      parts.push(`${ONES[Math.floor(k / 100)]} hundred`);
      k %= 100;
    }
    if (k >= 20) {
      parts.push(TENS[Math.floor(k / 10)]! + (k % 10 ? `-${ONES[k % 10]}` : ""));
    } else if (k > 0) {
      parts.push(ONES[k]!);
    }
    return parts.join(" ");
  };
  if (n === 0) {
    return "zero";
  }
  const thousands = Math.floor(n / 1000);
  const rest = n % 1000;
  return [thousands ? `${below1000(thousands)} thousand` : "", below1000(rest)]
    .filter(Boolean)
    .join(" ");
}

const ORDINAL_WORDS: Record<string, string> = {
  one: "first",
  two: "second",
  three: "third",
  five: "fifth",
  eight: "eighth",
  nine: "ninth",
  twelve: "twelfth"
};

/** Cardinal words made ordinal by their last word, as Word's `OrdText`. */
function ordinalWords(words: string): string {
  return words.replace(/[a-z]+$/, last => {
    if (ORDINAL_WORDS[last]) {
      return ORDINAL_WORDS[last];
    }
    return last.endsWith("y") ? `${last.slice(0, -1)}ieth` : `${last}th`;
  });
}

function ordinal(n: number): string {
  const tens = Math.abs(n) % 100;
  const suffix =
    tens >= 11 && tens <= 13 ? "th" : (["th", "st", "nd", "rd"][Math.abs(n) % 10] ?? "th");
  return `${n}${suffix}`;
}

// =============================================================================
// \# — numeric pictures
// =============================================================================

/** A field value as a number: digits with an optional sign, grouping commas and decimals. */
function parseFieldNumber(value: string): number | undefined {
  const trimmed = value.trim();
  if (!/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(trimmed)) {
    return undefined;
  }
  if (!/\d/.test(trimmed)) {
    return undefined;
  }
  const n = Number(trimmed.replace(/,/g, ""));
  return Number.isFinite(n) ? n : undefined;
}

/** `n` rounded to `decimals` places, half away from zero, as digits. */
function roundedDigits(n: number, decimals: number): { int: string; frac: string } {
  const [int, frac = ""] = (Math.round(Math.abs(n) * 10 ** decimals + 1e-9) / 10 ** decimals)
    .toFixed(decimals)
    .split(".");
  return { int: int!, frac };
}

/** The sections of a picture (positive;negative;zero), split outside quotes. */
function pictureSections(picture: string): string[] {
  const out: string[] = [""];
  let quoted = false;
  for (const c of picture) {
    if (c === "'") {
      quoted = !quoted;
    }
    if (c === ";" && !quoted) {
      out.push("");
    } else {
      out[out.length - 1] += c;
    }
  }
  return out;
}

const DIGIT = /[0#x]/;

/**
 * `n` drawn in Word's numeric picture: `0` a digit, `#` a digit or a space,
 * `x` a digit after which (to the left) digits are dropped — or, after the
 * decimal point, the place to round at; `.` the decimal point, `,` digit
 * grouping, `-` / `+` the sign, `'text'` literal text, anything else itself.
 * Sections `positive;negative;zero` pick the picture by sign; with one, a
 * negative number gets a minus.
 */
export function formatNumberPicture(n: number, picture: string): string {
  const sections = pictureSections(picture);
  let section = sections[0]!;
  // With a section of its own for negatives, the picture states the sign.
  let signed = true;
  if (n < 0 && sections[1] !== undefined) {
    section = sections[1];
    signed = false;
  } else if (n === 0 && sections[2] !== undefined) {
    section = sections[2];
  }
  const chars: PictureChar[] = [];
  let quoted = false;
  for (const c of section) {
    if (c === "'") {
      quoted = !quoted;
    } else {
      chars.push({ c, literal: quoted });
    }
  }
  const isPlace = (x: PictureChar): boolean =>
    !x.literal && (DIGIT.test(x.c) || x.c === "." || x.c === ",");
  const first = chars.findIndex(isPlace);
  if (first < 0) {
    return chars.map(x => x.c).join("");
  }
  let last = chars.length - 1;
  while (!isPlace(chars[last]!)) {
    last--;
  }
  const core = chars.slice(first, last + 1).filter(x => !x.literal);
  const dot = core.findIndex(x => x.c === ".");
  const intPlaces = (dot < 0 ? core : core.slice(0, dot)).map(x => x.c);
  const fracPlaces = dot < 0 ? [] : core.slice(dot + 1).map(x => x.c);
  const rounded = roundedDigits(n, fracPlaces.filter(c => DIGIT.test(c)).length);
  const negative = n < 0 && /[1-9]/.test(rounded.int + rounded.frac);

  let slots = intPlaces.filter(c => DIGIT.test(c));
  let digits = rounded.int === "0" ? "" : rounded.int;
  const x = slots.indexOf("x");
  if (x >= 0) {
    slots = slots.slice(x);
    digits = digits.slice(-slots.length);
  }
  let padded = digits;
  for (let i = slots.length - digits.length - 1; i >= 0; i--) {
    padded = (slots[i] === "#" ? " " : "0") + padded;
  }
  if (intPlaces.includes(",")) {
    let grouped = "";
    for (let i = 0; i < padded.length; i++) {
      if (i > 0 && (padded.length - i) % 3 === 0) {
        // A separator with no digit to its left is dropped, not padded.
        grouped += padded[i - 1] === " " ? "" : ",";
      }
      grouped += padded[i];
    }
    padded = grouped;
  }
  const body = dot >= 0 ? `${padded}.${rounded.frac}` : padded;

  let stated = false;
  const literal = (list: readonly PictureChar[]): string =>
    list
      .map(({ c, literal: lit }) => {
        if (lit || !signed || (c !== "-" && c !== "+")) {
          return c;
        }
        stated = true;
        return negative ? "-" : c === "+" && n !== 0 ? "+" : " ";
      })
      .join("");
  const prefix = literal(chars.slice(0, first));
  const suffix = literal(chars.slice(last + 1));
  const minus = signed && negative && !stated ? "-" : "";
  return prefix + minus + body + suffix;
}

interface PictureChar {
  readonly c: string;
  /** Inside `'…'`. */
  readonly literal: boolean;
}

// =============================================================================
// \@ — date-time pictures
// =============================================================================

const MONTHS = [
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
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/** The parts of a date-time value, as written (no time zone applies). */
interface FieldDate {
  readonly year: number;
  readonly month: number;
  readonly day: number;
  readonly hour: number;
  readonly minute: number;
  readonly second: number;
}

/**
 * A field value as a date: ISO (`2024-03-05`, `2024-03-05T13:08[:09]`, a space
 * for the `T`) or US (`3/5/2024`, with an optional `13:08`). The time is taken
 * as written, in no time zone.
 */
function parseFieldDate(value: string): FieldDate | undefined {
  const v = value.trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(v);
  const us = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(v);
  const m = iso ?? us;
  if (!m) {
    return undefined;
  }
  const [year, month, day] = iso
    ? [Number(m[1]), Number(m[2]), Number(m[3])]
    : [Number(m[3]), Number(m[1]), Number(m[2])];
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    return undefined;
  }
  return {
    year,
    month,
    day,
    hour: Number(m[4] ?? 0),
    minute: Number(m[5] ?? 0),
    second: Number(m[6] ?? 0)
  };
}

const DATE_TOKEN = /'[^']*'|AM\/PM|am\/pm|A\/P|a\/p|y+|M+|d+|H+|h+|m+|s+/g;

/**
 * `date` drawn in Word's date-time picture: `M`–`MMMM` month, `d`–`dddd`
 * day, `yy` / `yyyy` year, `H`/`HH` 24-hour and `h`/`hh` 12-hour, `m`/`mm`
 * minute, `s`/`ss` second, `AM/PM` `am/pm` `A/P` `a/p`, `'text'` literal.
 * English month and day names.
 */
export function formatDatePicture(date: FieldDate, picture: string): string {
  const pad = (n: number): string => String(n).padStart(2, "0");
  const weekday = new Date(Date.UTC(date.year, date.month - 1, date.day)).getUTCDay();
  const h12 = date.hour % 12 === 0 ? 12 : date.hour % 12;
  const pm = date.hour >= 12;
  return picture.replace(DATE_TOKEN, token => {
    switch (token[0]) {
      case "'":
        return token.slice(1, -1);
      case "y":
        return token.length >= 3 ? String(date.year) : pad(date.year % 100);
      case "M":
        return token.length >= 4
          ? MONTHS[date.month - 1]!
          : token.length === 3
            ? MONTHS[date.month - 1]!.slice(0, 3)
            : token.length === 2
              ? pad(date.month)
              : String(date.month);
      case "d":
        return token.length >= 4
          ? DAYS[weekday]!
          : token.length === 3
            ? DAYS[weekday]!.slice(0, 3)
            : token.length === 2
              ? pad(date.day)
              : String(date.day);
      case "H":
        return token.length >= 2 ? pad(date.hour) : String(date.hour);
      case "h":
        return token.length >= 2 ? pad(h12) : String(h12);
      case "m":
        return token.length >= 2 ? pad(date.minute) : String(date.minute);
      case "s":
        return token.length >= 2 ? pad(date.second) : String(date.second);
      case "A":
        return token === "AM/PM" ? (pm ? "PM" : "AM") : pm ? "P" : "A";
      case "a":
        return token === "am/pm" ? (pm ? "pm" : "am") : pm ? "p" : "a";
      default:
        return token;
    }
  });
}
