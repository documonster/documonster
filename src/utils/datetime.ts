/**
 * High-performance native date parsing and formatting utilities
 *
 * Zero external dependencies. Optimized for CSV batch processing.
 * Uses character code operations and lookup tables for maximum speed.
 */

// ============================================================================
// Types
// ============================================================================

/** Supported date format strings */
export type DateFormat =
  | "YYYY-MM-DD[T]HH:mm:ssZ"
  | "YYYY-MM-DD[T]HH:mm:ss"
  | "YYYY-MM-DD[T]HH:mm:ss.SSSZ"
  | "YYYY-MM-DD"
  | "YYYY-MM-DD HH:mm:ss"
  | "MM-DD-YYYY"
  | "MM-DD-YYYY HH:mm:ss"
  | "MM/DD/YYYY HH:mm:ss"
  | "DD-MM-YYYY"
  | "DD-MM-YYYY HH:mm:ss"
  | "DD/MM/YYYY HH:mm:ss";

// ============================================================================
// Constants - Pre-computed lookup tables for zero-allocation operations
// ============================================================================

// Padding lookup (0-59 covers hours, minutes, seconds, and months/days)
const PAD2 = Array.from({ length: 60 }, (_, i) => (i < 10 ? `0${i}` : `${i}`));

// Character codes for fast comparison
const C_0 = 48;
const C_DASH = 45;
const C_SLASH = 47;
const C_COLON = 58;
const C_T = 84;
const C_SPACE = 32;
const C_Z = 90;
const C_PLUS = 43;
const C_DOT = 46;

// ============================================================================
// Low-level utilities (inlined for JIT optimization)
// ============================================================================

// Inline digit extraction - avoid function call overhead
// Using bitwise OR for integer coercion (faster than Math.floor)
const digit2 = (s: string, i: number) =>
  ((s.charCodeAt(i) - C_0) * 10 + s.charCodeAt(i + 1) - C_0) | 0;
const digit4 = (s: string, i: number) =>
  ((s.charCodeAt(i) - C_0) * 1000 +
    (s.charCodeAt(i + 1) - C_0) * 100 +
    (s.charCodeAt(i + 2) - C_0) * 10 +
    s.charCodeAt(i + 3) -
    C_0) |
  0;

// Days in month lookup (index 0 unused, 1-12 for Jan-Dec)
// Using 31 for Feb; actual validation done by Date constructor
const DAYS_IN_MONTH = [0, 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

function validateDate(y: number, m: number, d: number): Date | null {
  // Fast bounds check using lookup table
  if (m < 1 || m > 12 || d < 1 || d > DAYS_IN_MONTH[m]) {
    return null;
  }
  const date = new Date(y, m - 1, d);
  // Check for overflow (e.g., Feb 30 -> Mar 2)
  return date.getMonth() === m - 1 ? date : null;
}

function validateDateTime(
  y: number,
  m: number,
  d: number,
  h: number,
  min: number,
  s: number
): Date | null {
  if (m < 1 || m > 12 || d < 1 || d > DAYS_IN_MONTH[m]) {
    return null;
  }
  if (h > 23 || min > 59 || s > 59) {
    return null;
  }
  return new Date(y, m - 1, d, h, min, s);
}

// ============================================================================
// Specialized parsers (length-based dispatch for speed)
// ============================================================================

// YYYY-MM-DD (10 chars)
function parseISO(s: string): Date | null {
  if (s.charCodeAt(4) !== C_DASH || s.charCodeAt(7) !== C_DASH) {
    return null;
  }
  return validateDate(digit4(s, 0), digit2(s, 5), digit2(s, 8));
}

// YYYY-MM-DDTHH:mm:ss (19 chars)
function parseISOT(s: string): Date | null {
  if (
    s.charCodeAt(4) !== C_DASH ||
    s.charCodeAt(7) !== C_DASH ||
    s.charCodeAt(10) !== C_T ||
    s.charCodeAt(13) !== C_COLON ||
    s.charCodeAt(16) !== C_COLON
  ) {
    return null;
  }
  return validateDateTime(
    digit4(s, 0),
    digit2(s, 5),
    digit2(s, 8),
    digit2(s, 11),
    digit2(s, 14),
    digit2(s, 17)
  );
}

// YYYY-MM-DD HH:mm:ss (19 chars, space separator)
function parseISOSpace(s: string): Date | null {
  if (
    s.charCodeAt(4) !== C_DASH ||
    s.charCodeAt(7) !== C_DASH ||
    s.charCodeAt(10) !== C_SPACE ||
    s.charCodeAt(13) !== C_COLON ||
    s.charCodeAt(16) !== C_COLON
  ) {
    return null;
  }
  return validateDateTime(
    digit4(s, 0),
    digit2(s, 5),
    digit2(s, 8),
    digit2(s, 11),
    digit2(s, 14),
    digit2(s, 17)
  );
}

// YYYY-MM-DDTHH:mm:ssZ (20 chars)
function parseISOZ(s: string): Date | null {
  if (s.charCodeAt(19) !== C_Z) {
    return null;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// YYYY-MM-DDTHH:mm:ss.SSSZ (24 chars)
function parseISOMsZ(s: string): Date | null {
  if (s.charCodeAt(19) !== C_DOT || s.charCodeAt(23) !== C_Z) {
    return null;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// YYYY-MM-DDTHH:mm:ss+HH:mm (25 chars)
function parseISOOffset(s: string): Date | null {
  const c = s.charCodeAt(19);
  if (c !== C_PLUS && c !== C_DASH) {
    return null;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// YYYY-MM-DDTHH:mm:ss.SSS+HH:mm (29 chars)
function parseISOMsOffset(s: string): Date | null {
  if (s.charCodeAt(19) !== C_DOT) {
    return null;
  }
  const c = s.charCodeAt(23);
  if (c !== C_PLUS && c !== C_DASH) {
    return null;
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// MM-DD-YYYY or MM/DD/YYYY (10 chars)
function parseUS(s: string): Date | null {
  const sep = s.charCodeAt(2);
  if ((sep !== C_DASH && sep !== C_SLASH) || s.charCodeAt(5) !== sep) {
    return null;
  }
  return validateDate(digit4(s, 6), digit2(s, 0), digit2(s, 3));
}

// DD-MM-YYYY or DD/MM/YYYY (10 chars)
function parseEU(s: string): Date | null {
  const sep = s.charCodeAt(2);
  if ((sep !== C_DASH && sep !== C_SLASH) || s.charCodeAt(5) !== sep) {
    return null;
  }
  return validateDate(digit4(s, 6), digit2(s, 3), digit2(s, 0));
}

// MM-DD-YYYY HH:mm:ss or MM/DD/YYYY HH:mm:ss (19 chars)
function parseUSTime(s: string): Date | null {
  const sep = s.charCodeAt(2);
  if ((sep !== C_DASH && sep !== C_SLASH) || s.charCodeAt(5) !== sep) {
    return null;
  }
  if (
    s.charCodeAt(10) !== C_SPACE ||
    s.charCodeAt(13) !== C_COLON ||
    s.charCodeAt(16) !== C_COLON
  ) {
    return null;
  }
  return validateDateTime(
    digit4(s, 6),
    digit2(s, 0),
    digit2(s, 3),
    digit2(s, 11),
    digit2(s, 14),
    digit2(s, 17)
  );
}

// DD-MM-YYYY HH:mm:ss or DD/MM/YYYY HH:mm:ss (19 chars)
function parseEUTime(s: string): Date | null {
  const sep = s.charCodeAt(2);
  if ((sep !== C_DASH && sep !== C_SLASH) || s.charCodeAt(5) !== sep) {
    return null;
  }
  if (
    s.charCodeAt(10) !== C_SPACE ||
    s.charCodeAt(13) !== C_COLON ||
    s.charCodeAt(16) !== C_COLON
  ) {
    return null;
  }
  return validateDateTime(
    digit4(s, 6),
    digit2(s, 3),
    digit2(s, 0),
    digit2(s, 11),
    digit2(s, 14),
    digit2(s, 17)
  );
}

// ============================================================================
// Format dispatch tables
// ============================================================================

type Parser = (s: string) => Date | null;

const PARSERS: Record<DateFormat, Parser> = {
  "YYYY-MM-DD": parseISO,
  "YYYY-MM-DD[T]HH:mm:ss": parseISOT,
  "YYYY-MM-DD HH:mm:ss": parseISOSpace,
  "YYYY-MM-DD[T]HH:mm:ssZ": s =>
    s.length === 20 ? parseISOZ(s) : s.length === 25 ? parseISOOffset(s) : null,
  "YYYY-MM-DD[T]HH:mm:ss.SSSZ": s =>
    s.length === 24 ? parseISOMsZ(s) : s.length === 29 ? parseISOMsOffset(s) : null,
  "MM-DD-YYYY": parseUS,
  "MM-DD-YYYY HH:mm:ss": parseUSTime,
  "MM/DD/YYYY HH:mm:ss": parseUSTime,
  "DD-MM-YYYY": parseEU,
  "DD-MM-YYYY HH:mm:ss": parseEUTime,
  "DD/MM/YYYY HH:mm:ss": parseEUTime
};

// Length-based auto-detection (ISO formats only, US/EU excluded due to ambiguity)
const AUTO_DETECT: Array<[number, Parser[]]> = [
  [10, [parseISO]],
  [19, [parseISOT, parseISOSpace]],
  [20, [parseISOZ]],
  [24, [parseISOMsZ]],
  [25, [parseISOOffset]],
  [29, [parseISOMsOffset]]
];

// ============================================================================
// High-performance batch processors (class-based for state encapsulation)
// ============================================================================

/**
 * Optimized date parser for batch processing
 *
 * @example
 * const parser = DateParser.create(["YYYY-MM-DD"]);
 * const dates = parser.parseAll(csvStrings);
 */
export class DateParser {
  private readonly fns: Parser[];
  private readonly single: boolean;
  private readonly fn0: Parser;

  private constructor(fns: Parser[]) {
    this.fns = fns;
    this.single = fns.length === 1;
    this.fn0 = fns[0];
  }

  /** Create parser for specific formats */
  static create(formats: readonly DateFormat[]): DateParser {
    return new DateParser(formats.map(f => PARSERS[f]).filter(Boolean));
  }

  /** Create parser for auto-detecting ISO formats */
  static iso(): DateParser {
    const fns: Parser[] = [];
    for (const [, parsers] of AUTO_DETECT) {
      fns.push(...parsers);
    }
    return new DateParser(fns);
  }

  /** Parse single value */
  parse = (value: string): Date | null => {
    if (!value) {
      return null;
    }
    const s = value.trim();
    if (!s) {
      return null;
    }
    // Fast path for single parser
    if (this.single) {
      return this.fn0(s);
    }
    // Multi-parser path
    for (let i = 0, len = this.fns.length; i < len; i++) {
      const r = this.fns[i](s);
      if (r) {
        return r;
      }
    }
    return null;
  };

  /** Parse array of values */
  parseAll(values: string[]): (Date | null)[] {
    const len = values.length;
    const out = new Array<Date | null>(len);
    const parse = this.parse;
    for (let i = 0; i < len; i++) {
      out[i] = parse(values[i]);
    }
    return out;
  }

  /** Parse and filter valid dates */
  parseValid(values: string[]): Date[] {
    const out: Date[] = [];
    const parse = this.parse;
    for (let i = 0, len = values.length; i < len; i++) {
      const d = parse(values[i]);
      if (d) {
        out.push(d);
      }
    }
    return out;
  }
}

function tzOffset(d: Date): string {
  const off = -d.getTimezoneOffset();
  const sign = off >= 0 ? "+" : "-";
  const h = (Math.abs(off) / 60) | 0; // Bitwise OR faster than Math.floor
  const m = Math.abs(off) % 60;
  return `${sign}${PAD2[h]}:${PAD2[m]}`;
}

/**
 * Optimized date formatter for batch processing
 *
 * @example
 * const formatter = DateFormatter.create("YYYY-MM-DD", { utc: true });
 * const strings = formatter.formatAll(dates);
 */
export class DateFormatter {
  private readonly fn: (d: Date) => string;

  private constructor(fn: (d: Date) => string) {
    this.fn = fn;
  }

  /** Create ISO formatter (fastest) */
  static iso(utc = false): DateFormatter {
    // Direct string building is faster than toISOString() + slice
    return utc
      ? new DateFormatter(d => {
          if (!(d instanceof Date)) {
            return "";
          }
          const t = d.getTime();
          if (Number.isNaN(t)) {
            return "";
          } // NaN check
          const y = d.getUTCFullYear();
          const M = d.getUTCMonth() + 1;
          const D = d.getUTCDate();
          const H = d.getUTCHours();
          const m = d.getUTCMinutes();
          const s = d.getUTCSeconds();
          const ms = d.getUTCMilliseconds();
          return `${y}-${PAD2[M]}-${PAD2[D]}T${PAD2[H]}:${PAD2[m]}:${PAD2[s]}.${ms < 10 ? "00" + ms : ms < 100 ? "0" + ms : ms}Z`;
        })
      : new DateFormatter(d => {
          if (!(d instanceof Date)) {
            return "";
          }
          const t = d.getTime();
          if (Number.isNaN(t)) {
            return "";
          } // NaN check
          const y = d.getFullYear();
          const M = d.getMonth() + 1;
          const D = d.getDate();
          const H = d.getHours();
          const m = d.getMinutes();
          const s = d.getSeconds();
          const ms = d.getMilliseconds();
          return `${y}-${PAD2[M]}-${PAD2[D]}T${PAD2[H]}:${PAD2[m]}:${PAD2[s]}.${ms < 10 ? "00" + ms : ms < 100 ? "0" + ms : ms}${tzOffset(d)}`;
        });
  }

  /** Create custom format formatter */
  static create(format: string, options?: { utc?: boolean }): DateFormatter {
    const utc = options?.utc ?? false;

    // Fast paths for common formats (no regex, direct string building)
    if (format === "YYYY-MM-DD") {
      return utc
        ? new DateFormatter(d => {
            if (!(d instanceof Date)) {
              return "";
            }
            const t = d.getTime();
            if (Number.isNaN(t)) {
              return "";
            }
            return `${d.getUTCFullYear()}-${PAD2[d.getUTCMonth() + 1]}-${PAD2[d.getUTCDate()]}`;
          })
        : new DateFormatter(d => {
            if (!(d instanceof Date)) {
              return "";
            }
            const t = d.getTime();
            if (Number.isNaN(t)) {
              return "";
            }
            return `${d.getFullYear()}-${PAD2[d.getMonth() + 1]}-${PAD2[d.getDate()]}`;
          });
    }

    if (format === "YYYY-MM-DD HH:mm:ss") {
      return utc
        ? new DateFormatter(d => {
            if (!(d instanceof Date)) {
              return "";
            }
            const t = d.getTime();
            if (Number.isNaN(t)) {
              return "";
            }
            return `${d.getUTCFullYear()}-${PAD2[d.getUTCMonth() + 1]}-${PAD2[d.getUTCDate()]} ${PAD2[d.getUTCHours()]}:${PAD2[d.getUTCMinutes()]}:${PAD2[d.getUTCSeconds()]}`;
          })
        : new DateFormatter(d => {
            if (!(d instanceof Date)) {
              return "";
            }
            const t = d.getTime();
            if (Number.isNaN(t)) {
              return "";
            }
            return `${d.getFullYear()}-${PAD2[d.getMonth() + 1]}-${PAD2[d.getDate()]} ${PAD2[d.getHours()]}:${PAD2[d.getMinutes()]}:${PAD2[d.getSeconds()]}`;
          });
    }

    if (format === "MM-DD-YYYY" || format === "MM/DD/YYYY") {
      const sep = format.charAt(2);
      return utc
        ? new DateFormatter(d => {
            if (!(d instanceof Date)) {
              return "";
            }
            const t = d.getTime();
            if (Number.isNaN(t)) {
              return "";
            }
            return `${PAD2[d.getUTCMonth() + 1]}${sep}${PAD2[d.getUTCDate()]}${sep}${d.getUTCFullYear()}`;
          })
        : new DateFormatter(d => {
            if (!(d instanceof Date)) {
              return "";
            }
            const t = d.getTime();
            if (Number.isNaN(t)) {
              return "";
            }
            return `${PAD2[d.getMonth() + 1]}${sep}${PAD2[d.getDate()]}${sep}${d.getFullYear()}`;
          });
    }

    if (format === "DD-MM-YYYY" || format === "DD/MM/YYYY") {
      const sep = format.charAt(2);
      return utc
        ? new DateFormatter(d => {
            if (!(d instanceof Date)) {
              return "";
            }
            const t = d.getTime();
            if (Number.isNaN(t)) {
              return "";
            }
            return `${PAD2[d.getUTCDate()]}${sep}${PAD2[d.getUTCMonth() + 1]}${sep}${d.getUTCFullYear()}`;
          })
        : new DateFormatter(d => {
            if (!(d instanceof Date)) {
              return "";
            }
            const t = d.getTime();
            if (Number.isNaN(t)) {
              return "";
            }
            return `${PAD2[d.getDate()]}${sep}${PAD2[d.getMonth() + 1]}${sep}${d.getFullYear()}`;
          });
    }

    // Generic path with pre-compiled template
    return DateFormatter.createGeneric(format, utc);
  }

  /** Generic formatter for arbitrary formats */
  private static createGeneric(format: string, utc: boolean): DateFormatter {
    // Pre-process escaped sections (e.g. [literal text] → placeholder)
    // Manual scan avoids regex backtracking on adversarial input
    const esc: string[] = [];
    const parts: string[] = [];
    let i = 0;
    while (i < format.length) {
      if (format[i] === "[") {
        const closeIdx = format.indexOf("]", i + 1);
        if (closeIdx !== -1) {
          esc.push(format.slice(i + 1, closeIdx));
          parts.push(`\x00${esc.length - 1}\x00`);
          i = closeIdx + 1;
        } else {
          // No closing bracket — treat '[' as literal
          parts.push(format[i]);
          i++;
        }
      } else {
        parts.push(format[i]);
        i++;
      }
    }
    const tpl = parts.join("");

    // Detect used tokens for conditional computation
    const hasY = tpl.includes("YYYY");
    const hasMs = tpl.includes("SSS");
    const hasM = tpl.includes("MM");
    const hasD = tpl.includes("DD");
    const hasH = tpl.includes("HH");
    const hasMin = tpl.includes("mm");
    const hasS = tpl.includes("ss");
    const hasZ = tpl.includes("Z");

    return new DateFormatter(d => {
      if (!(d instanceof Date)) {
        return "";
      }
      const t = d.getTime();
      if (Number.isNaN(t)) {
        return "";
      }

      let out = tpl;
      if (hasY) {
        out = out.replace(/YYYY/g, String(utc ? d.getUTCFullYear() : d.getFullYear()));
      }
      if (hasMs) {
        const ms = utc ? d.getUTCMilliseconds() : d.getMilliseconds();
        out = out.replace(/SSS/g, ms < 10 ? `00${ms}` : ms < 100 ? `0${ms}` : String(ms));
      }
      if (hasM) {
        out = out.replace(/MM/g, PAD2[utc ? d.getUTCMonth() + 1 : d.getMonth() + 1]);
      }
      if (hasD) {
        out = out.replace(/DD/g, PAD2[utc ? d.getUTCDate() : d.getDate()]);
      }
      if (hasH) {
        out = out.replace(/HH/g, PAD2[utc ? d.getUTCHours() : d.getHours()]);
      }
      if (hasMin) {
        out = out.replace(/mm/g, PAD2[utc ? d.getUTCMinutes() : d.getMinutes()]);
      }
      if (hasS) {
        out = out.replace(/ss/g, PAD2[utc ? d.getUTCSeconds() : d.getSeconds()]);
      }
      if (hasZ) {
        out = out.replace(/Z/g, utc ? "Z" : tzOffset(d));
      }

      if (esc.length) {
        // oxlint-disable-next-line no-control-regex
        out = out.replace(/\x00(\d+)\x00/g, (_, i) => esc[+i]);
      }
      return out;
    });
  }

  /** Format single date */
  format = (date: Date): string => this.fn(date);

  /** Format array of dates */
  formatAll(dates: Date[]): string[] {
    const len = dates.length;
    const out = new Array<string>(len);
    const fn = this.fn;
    for (let i = 0; i < len; i++) {
      out[i] = fn(dates[i]);
    }
    return out;
  }
}

/** Get supported format strings */
export function getSupportedFormats(): DateFormat[] {
  return Object.keys(PARSERS) as DateFormat[];
}
