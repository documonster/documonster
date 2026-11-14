/**
 * Base utility functions shared between Node.js and Browser
 * All functions use standard Web APIs that work in both environments
 * (Node.js 16+ supports atob/btoa/TextEncoder/TextDecoder globally)
 */

import { isNode } from "@utils/env";
import { partsToSerial, partsToUtcDate, serialToParts, utcDateToParts } from "@utils/excel-serial";

// =============================================================================
// Base64 utilities (with native Buffer optimization for Node.js)
// =============================================================================

/**
 * Convert base64 string to Uint8Array
 * Uses native Buffer in Node.js for better performance
 */
export function base64ToUint8Array(base64: string): Uint8Array {
  // Node.js: use native Buffer (fast, C++ implementation)
  if (isNode()) {
    return Buffer.from(base64, "base64");
  }
  // Browser: use atob
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

// =============================================================================
// Basic utilities
// =============================================================================

export function delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// =============================================================================
// Date utilities
// =============================================================================

/**
 * A `Date` as an Excel date serial.
 *
 * **The `Date` is read as a calendar value, not as an instant** — its *UTC* fields are taken to
 * be the spreadsheet's wall-clock date and time. That is the convention the whole library
 * follows (see `@utils/excel-serial`), and it is why `new Date(Date.UTC(2024, 0, 15))` is the
 * form to pass. A locally-constructed `new Date(2024, 0, 15)` names a different instant in
 * every timezone, so it converts to a different serial in every timezone; in UTC+8 it yields
 * 45305.667, which Excel displays as 2024-01-14 16:00.
 *
 * `Date` cannot express which of its two readings is meant, which is the whole problem. Prefer
 * {@link partsToSerial} internally, and `Cell.setValue` with a Temporal `Plain*` or an
 * `ExcelDateTimeParts` externally, where the value says what it is.
 */
export function dateToExcel(d: Date, date1904?: boolean): number {
  return partsToSerial(utcDateToParts(d), date1904);
}

/**
 * An Excel date serial as a `Date` carrying the calendar value in its **UTC** fields.
 *
 * The counterpart of {@link dateToExcel}, and the reason every consumer of a date cell in this
 * codebase reads it with `getUTC*`. A caller who wants the value without the ambiguity should
 * use `Cell.getDateParts` or `Cell.getTemporal`.
 */
export function excelToDate(v: number, date1904?: boolean): Date {
  return partsToUtcDate(serialToParts(v, date1904));
}

/**
 * Parse an OOXML date string into a `Date` carrying the calendar value in its UTC fields, as
 * {@link excelToDate} does — the same value whatever zone the reader runs in.
 *
 * An OOXML datetime (`2024-01-15T00:00:00`, or with a space for the `T`) carries no timezone, and
 * is read as UTC, as is an ISO date (`2024-01-15`, `2024-01`, `2024`). Text naming its zone
 * (`…Z`, `…+02:00`, `… GMT+0200`) names an instant, which is kept. ISO text is read here and its
 * fields checked — a day past the month's end or a 60th second is invalid — because engines
 * disagree on it: V8 accepted `2024-02-09Z` and rolled `2024-02-30` into March, Safari read the
 * one as invalid and a 60th second as the next minute.
 *
 * Any other text the engine reads (`1/15/2020`, `Jan 15 2020 10:30 PM`) is its calendar value:
 * the engine reads its date, and its time of day is read here. Read by the engine, a time is
 * local, and one inside a daylight-saving gap is moved past it — `00:30` on a day that skips its
 * midnight reads as `01:30` — which nothing can undo afterwards, as both name one instant. Text
 * holding a time this cannot read (`1:2`), or two, is invalid rather than read as midnight, and
 * so is a bare number, which engines read as different years.
 */
export function parseOoxmlDate(raw: string): Date {
  const text = raw.trim();
  if (ISO_LIKE.test(text)) {
    return isoDate(text);
  }
  if (ZONE.test(text)) {
    // Text naming its zone names an instant, which is what the engine reads.
    return new Date(text);
  }
  const time = TIME_OF_DAY.exec(text);
  // The text without its time — `1/15/2020, 10:30 AM` leaves `1/15/2020` — is what the engine reads the date from.
  const rest =
    time === null
      ? text
      : (
          text.slice(0, time.index + time[1]!.length) +
          " " +
          text.slice(time.index + time[0].length)
        )
          .replace(TRAILING_CONNECTIVE, "")
          .trim();
  // A time left over is one this cannot read, or a second one: the engine would read it, or drop it, its own way.
  if (/\d:\d/.test(rest) || /^\d*$/.test(rest)) {
    return new Date(NaN);
  }
  const date = calendarDate(rest);
  if (date === undefined) {
    return new Date(NaN);
  }
  if (time === null) {
    return utcDate(date.year, date.month, date.day, 0, 0, 0, 0);
  }
  let hours = Number(time[2]);
  const meridiem = time[6]?.toUpperCase();
  if (meridiem !== undefined) {
    if (hours < 1 || hours > 12) {
      return new Date(NaN);
    }
    hours = (hours % 12) + (meridiem === "PM" ? 12 : 0);
  }
  const minutes = Number(time[3]);
  const seconds = time[4] === undefined ? 0 : Number(time[4]);
  if (hours > 23 || minutes > 59 || seconds > 59) {
    return new Date(NaN);
  }
  return utcDate(date.year, date.month, date.day, hours, minutes, seconds, milliseconds(time[5]));
}

/**
 * ISO 8601 text, read strictly: a year (`2024`, or six digits with a sign), its month and day, then a time — after
 * `T`, `t` or a space — to the second with any fraction, and a zone, `Z` or an offset. A zone needs a time, but for
 * `Z` after a date, which is UTC as the date is.
 */
const ISO =
  /^([+-]\d{6}|\d{4})(?:-(\d{2})(?:-(\d{2})(?:(?:[Tt ](\d{2}):(\d{2})(?::(\d{2})(?:[.,](\d+))?)?)?([Zz]|[+-]\d{2}(?::?\d{2})?)?)?)?)?$/;

/** Text that is ISO 8601 or meant to be: a year, then a month. Read by {@link isoDate}, never by the engine. */
const ISO_LIKE = /^[+-]?\d{4,6}(?:-\d{2}|$)/;

/** {@link ISO} text as the instant it names, a time without a zone in UTC; invalid where a field is out of range. */
function isoDate(text: string): Date {
  const m = ISO.exec(text);
  if (
    m === null ||
    m[1] === "-000000" ||
    (m[8] !== undefined && m[4] === undefined && !/^z$/i.test(m[8]))
  ) {
    return new Date(NaN);
  }
  const year = Number(m[1]);
  const month = m[2] === undefined ? 1 : Number(m[2]);
  const day = m[3] === undefined ? 1 : Number(m[3]);
  const hours = m[4] === undefined ? 0 : Number(m[4]);
  const minutes = m[5] === undefined ? 0 : Number(m[5]);
  const seconds = m[6] === undefined ? 0 : Number(m[6]);
  const ms = milliseconds(m[7]);
  const endOfDay = hours === 24 && minutes === 0 && seconds === 0 && ms === 0;
  if (
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > daysInMonth(year, month - 1) ||
    (hours > 23 && !endOfDay) ||
    minutes > 59 ||
    seconds > 59
  ) {
    return new Date(NaN);
  }
  const utc = utcDate(year, month - 1, day, hours, minutes, seconds, ms);
  const zone =
    m[8] === undefined || /^z$/i.test(m[8]) ? undefined : /^([+-])(\d{2}):?(\d{2})?$/.exec(m[8]);
  if (zone == null) {
    return utc;
  }
  const offsetHours = Number(zone[2]);
  const offsetMinutes = zone[3] === undefined ? 0 : Number(zone[3]);
  if (offsetHours > 23 || offsetMinutes > 59) {
    return new Date(NaN);
  }
  const offset = (zone[1] === "-" ? -1 : 1) * (offsetHours * 60 + offsetMinutes) * 60_000;
  return new Date(utc.getTime() - offset);
}

/** The milliseconds of a fraction of a second's digits (`5` is 500), truncated as engines truncate them. */
function milliseconds(fraction: string | undefined): number {
  return fraction === undefined ? 0 : Number(fraction.slice(0, 3).padEnd(3, "0"));
}

function daysInMonth(year: number, month: number): number {
  return utcDate(year, month + 1, 0, 0, 0, 0, 0).getUTCDate();
}

/** `Date.UTC` for any year: it reads a year from 0 to 99 as 1900 to 1999. */
function utcDate(
  year: number,
  month: number,
  day: number,
  hours: number,
  minutes: number,
  seconds: number,
  ms: number
): Date {
  const date = new Date(Date.UTC(2000, month, day, hours, minutes, seconds, ms));
  date.setUTCFullYear(date.getUTCFullYear() - 2000 + year);
  return date;
}

/**
 * The calendar date the engine reads `text` as, which names no time of day: its local date — a date that skips its
 * midnight reads as 01:00, on the same date. `undefined` for text it cannot read.
 */
function calendarDate(text: string): { year: number; month: number; day: number } | undefined {
  const read = new Date(text);
  if (Number.isNaN(read.getTime())) {
    return undefined;
  }
  return { year: read.getFullYear(), month: read.getMonth(), day: read.getDate() };
}

/**
 * A time of day in date text: `10:30`, `10:30:15`, `10:30:15.250`, with or without `AM` / `PM` — not inside a longer
 * run of digits, colons and points. What precedes it is the first group, as Safari 15.4 has no lookbehind.
 */
const TIME_OF_DAY =
  /(^|[^\d:.])(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?(?![\d:])(?:\s*([AaPp][Mm])(?![A-Za-z]))?/;

/** What joined a time to its date — `, ` or ` at ` — left at the end once the time is taken out. */
const TRAILING_CONNECTIVE = /(?:[\s,]+|\s+at)+$/i;

/** A zone in date text: `Z`, `GMT`, `UTC`, an offset (`+0200`, `-05:00`), or a named zone (`EST`, `(CEST)`). */
const ZONE =
  /\b(?:Z|GMT|UTC|UT)\b|[+-]\d{2}:?\d{2}(?!\d)|\([^)]*\)\s*$|\d(?:\s*[AP]M)?\s+[A-Z]{3,4}\s*$/;

// =============================================================================
// OOXML escape utilities (ST_Xstring, ISO/IEC 29500 clause 22.4.2.4)
// =============================================================================

/**
 * Pattern matching OOXML `_xHHHH_` escape sequences (case-insensitive hex).
 *
 * Per the OOXML spec, `_xHHHH_` encodes a Unicode code point where HHHH is
 * a 4-digit hexadecimal number. The spec uses uppercase, but real-world files
 * from third-party tools (Google Sheets, LibreOffice, etc.) may use lowercase.
 */
const ooxmlEscapeRegex = /_x([0-9A-Fa-f]{4})_/g;

/**
 * Decode OOXML `_xHHHH_` escape sequences in a string.
 *
 * Used when reading text content from `<t>` elements in shared strings,
 * rich text runs, and inline strings. The replacement works left-to-right,
 * so `_x005F_x000D_` correctly decodes to the literal string `_x000D_`
 * (the `_x005F_` decodes to `_`, consuming the match).
 */
export function decodeOoxmlEscape(text: string): string {
  return text.replace(ooxmlEscapeRegex, (match, $1) => {
    const code = parseInt($1, 16);
    // Reject characters that are invalid in XML 1.0:
    // - NUL (0x00)
    // - ASCII control chars except TAB(0x09), LF(0x0A), CR(0x0D)
    // - DEL (0x7F)
    // - Lone surrogates (0xD800-0xDFFF)
    // - Non-characters (0xFFFE, 0xFFFF)
    if (
      code === 0 ||
      (code >= 0x01 && code <= 0x08) ||
      code === 0x0b ||
      code === 0x0c ||
      (code >= 0x0e && code <= 0x1f) ||
      code === 0x7f ||
      (code >= 0xd800 && code <= 0xdfff) ||
      code === 0xfffe ||
      code === 0xffff
    ) {
      // Invalid character — leave the escape sequence as-is
      return match;
    }
    return String.fromCharCode(code);
  });
}

/**
 * Encode literal `_xHHHH_` patterns in a string for OOXML output.
 *
 * If a string naturally contains the pattern `_xHHHH_` (e.g., the user typed
 * `_x000D_`), the leading underscore must be escaped as `_x005F_` to prevent
 * readers from misinterpreting it as an escape sequence.
 *
 * Roundtrip guarantee: `decodeOoxmlEscape(encodeOoxmlEscape(s)) === s`
 */
export function encodeOoxmlEscape(text: string): string {
  return text.replace(ooxmlEscapeRegex, "_x005F_x$1_");
}

/**
 * Characters that XML attribute-value normalisation replaces with spaces
 * (XML 1.0 §3.3.3). When writing OOXML attribute values we must encode
 * these as `_xHHHH_` so that the original characters survive a round-trip.
 */
const xmlAttrUnsafeRe = /[\t\n\r]/g;
const xmlAttrUnsafeMap: Record<string, string> = {
  "\t": "_x0009_",
  "\n": "_x000A_",
  "\r": "_x000D_"
};

/**
 * Encode a string for safe use in an OOXML **XML attribute** value.
 *
 * Two transformations are applied (order matters):
 * 1. Literal `_xHHHH_` patterns are escaped (`_x005F_xHHHH_`) so readers
 *    do not misinterpret them as escape sequences.
 * 2. Characters that XML attribute-value normalisation would mangle
 *    (`\t`, `\n`, `\r`) are encoded as `_x0009_`, `_x000A_`, `_x000D_`.
 *
 * This is the write-side counterpart of {@link decodeOoxmlEscape}.
 * Use `encodeOoxmlEscape` for element **text** content and this function
 * for **attribute** values.
 */
export function encodeOoxmlAttr(text: string): string {
  // Step 1 – protect literal _xHHHH_ patterns (must come first so that
  // the _xHHHH_ sequences produced in step 2 are not double-escaped).
  let result = text.replace(ooxmlEscapeRegex, "_x005F_x$1_");
  // Step 2 – encode characters unsafe in XML attributes.
  result = result.replace(xmlAttrUnsafeRe, ch => xmlAttrUnsafeMap[ch]);
  return result;
}

// =============================================================================
// Parsing utilities
// =============================================================================

export function validInt(value: string | number): number {
  const i = typeof value === "number" ? value : parseInt(value, 10);
  return Number.isNaN(i) ? 0 : i;
}

// Number-format reading lives in one place; these names are kept here because every reader imports
// them from `@utils/utils`.
export { isDateFmt, splitFormatSections } from "./number-format";

export function parseBoolean(value: unknown): boolean {
  return value === true || value === "true" || value === 1 || value === "1";
}

// =============================================================================
// Collection utilities
// =============================================================================

export function* range(start: number, stop: number, step: number = 1): Generator<number> {
  const compareOrder = step > 0 ? (a: number, b: number) => a < b : (a: number, b: number) => a > b;
  for (let value = start; compareOrder(value, stop); value += step) {
    yield value;
  }
}

export function toSortedArray<T>(values: Iterable<T>): T[] {
  const result = Array.from(values);
  if (result.length <= 1) {
    return result;
  }
  // All numbers → numeric sort
  if (result.every(item => Number.isFinite(item))) {
    return result.sort((a, b) => (a as number) - (b as number));
  }
  // All Dates → chronological sort
  if (result.every(item => item instanceof Date)) {
    return result.sort((a, b) => (a as Date).getTime() - (b as Date).getTime());
  }
  // Mixed types → type-aware sort: numbers first (numerically), then dates (chronologically), then strings (lexicographic)
  return result.sort((a, b) => {
    const ta = sortTypeRank(a);
    const tb = sortTypeRank(b);
    if (ta !== tb) {
      return ta - tb;
    }
    // Same type group
    if (ta === 0) {
      return (a as number) - (b as number);
    }
    if (ta === 1) {
      return (a as Date).getTime() - (b as Date).getTime();
    }
    return String(a).localeCompare(String(b));
  });
}

/** Rank for mixed-type sort: numbers=0, dates=1, everything else=2 */
function sortTypeRank(v: unknown): number {
  if (Number.isFinite(v)) {
    return 0;
  }
  if (v instanceof Date) {
    return 1;
  }
  return 2;
}

// =============================================================================
// Buffer utilities (cross-platform)
// =============================================================================

const textDecoder = new TextDecoder("utf-8");

let latin1Decoder: TextDecoder | undefined;
let _latin1DecoderResolved = false;

function getLatin1Decoder(): TextDecoder | undefined {
  if (!_latin1DecoderResolved) {
    _latin1DecoderResolved = true;
    try {
      latin1Decoder = new TextDecoder("latin1");
    } catch {
      latin1Decoder = undefined;
    }
  }
  return latin1Decoder;
}

/**
 * Convert a Buffer, ArrayBuffer, or Uint8Array to a UTF-8 string
 * Works in both Node.js and browser environments
 */
export function bufferToString(chunk: ArrayBuffer | Uint8Array | string): string {
  if (typeof chunk === "string") {
    return chunk;
  }
  return textDecoder.decode(chunk);
}

/**
 * Convert Uint8Array to base64 string
 * Uses native Buffer in Node.js, optimized chunked conversion in browser
 */
export function uint8ArrayToBase64(bytes: Uint8Array): string {
  if (isNode()) {
    return Buffer.from(bytes).toString("base64");
  }

  // Browser: fastest path when latin1 TextDecoder exists.
  // Some environments can still throw on `btoa(...)` (e.g. if decoding yields non-Latin1 chars),
  // so fall back to a guaranteed-binary string conversion.
  if (getLatin1Decoder()) {
    try {
      return btoa(latin1Decoder!.decode(bytes));
    } catch {
      // fall through
    }
  }

  // Browser: chunked String.fromCharCode to avoid stack overflow and reduce string concatenation
  const CHUNK_SIZE = 0x8000; // 32KB chunks
  const chunks: string[] = [];
  for (let i = 0; i < bytes.length; i += CHUNK_SIZE) {
    // Spread the chunk (a Uint8Array is an iterable of numbers) so the call is
    // type-checked against `fromCharCode(...codes: number[])` without an `any`
    // cast; chunk size stays within the engine's argument-count limit.
    chunks.push(String.fromCharCode(...bytes.subarray(i, i + CHUNK_SIZE)));
  }
  return btoa(chunks.join(""));
}

/**
 * Convert string to UTF-16LE Uint8Array (used for Excel password hashing)
 */
export function stringToUtf16Le(str: string): Uint8Array {
  const bytes = new Uint8Array(str.length * 2);
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    bytes[i * 2] = code & 0xff;
    bytes[i * 2 + 1] = (code >> 8) & 0xff;
  }
  return bytes;
}
