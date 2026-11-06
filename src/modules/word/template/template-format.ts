/**
 * Format switches in template placeholders: `{{amount \# "#,##0.00"}}`,
 * `{{date \@ "d MMMM yyyy"}}`, `{{name \* Upper}}`, `{{phone \b "Tel: "}}`.
 * They are Word's own field switches, read and written as a `MERGEFIELD`
 * reads and writes them, so a template author who knows one knows the other.
 *
 * A locale (`TemplateOptions.locale`) does what Word's regional settings do: a
 * picture is written in that locale's symbols (`#.##0,00` in German), digits
 * are grouped as it groups them (`12,34,567` in Indian English), and months
 * and days are named in its language. The data is machine-written whatever
 * the locale — `1234.5` and `2026-11-07`, as XML and JSON write them.
 */

import type { DateNames, FieldDate, FieldFormat, NumberSymbols } from "@word/core/field-format";
import {
  fieldDate,
  fieldSwitchArgument,
  formatFieldResult,
  isGeneralFormat,
  parseFieldDate,
  parseFieldNumber,
  tokenizeFieldCode
} from "@word/core/field-format";

/** The switches a template placeholder accepts, as Word names them. */
const SWITCHES = new Set(["\\#", "\\@", "\\*", "\\b", "\\f"]);

/** `locale`'s symbols and names, or English for `undefined`. */
export function localeFieldFormat(locale: string | undefined): FieldFormat {
  if (locale === undefined) {
    return {};
  }
  let format = FORMAT_CACHE.get(locale);
  if (!format) {
    format = { symbols: localeNumberSymbols(locale), names: localeDateNames(locale) };
    FORMAT_CACHE.set(locale, format);
  }
  return format;
}

const FORMAT_CACHE = new Map<string, FieldFormat>();

/** Whether `locale` is a locale tag this platform can format numbers and dates in. */
export function isSupportedLocale(locale: unknown): boolean {
  if (typeof locale !== "string" || locale === "") {
    return false;
  }
  try {
    return (
      Intl.NumberFormat.supportedLocalesOf(locale).length > 0 &&
      Intl.DateTimeFormat.supportedLocalesOf(locale).length > 0
    );
  } catch {
    return false;
  }
}

function localeNumberSymbols(locale: string): NumberSymbols {
  // `useGrouping: "always"`: some locales leave four-digit numbers ungrouped.
  const parts = new Intl.NumberFormat(locale, {
    useGrouping: "always",
    numberingSystem: "latn"
  } as Intl.NumberFormatOptions).formatToParts(123456789.5);
  const decimal = parts.find(p => p.type === "decimal")?.value ?? ".";
  const group = parts.find(p => p.type === "group")?.value ?? ",";
  const integers = parts.filter(p => p.type === "integer").map(p => p.value.length);
  const primary = integers.at(-1) ?? 3;
  const secondary = integers.length > 2 ? integers.at(-2)! : primary;
  return { decimal, group, groupSizes: [primary, secondary] };
}

function localeDateNames(locale: string): DateNames {
  // The names as they stand in a date (`7 ноября`, not `ноябрь`), which is
  // where a picture writes them.
  const monthOf = (month: number, style: "long" | "short"): string => {
    const parts = new Intl.DateTimeFormat(locale, {
      day: "numeric",
      month: style,
      timeZone: "UTC"
    }).formatToParts(Date.UTC(2001, month, 15));
    return parts.find(p => p.type === "month")?.value ?? "";
  };
  // 2001-01-07 was a Sunday.
  const dayOf = (day: number, style: "long" | "short"): string =>
    new Intl.DateTimeFormat(locale, { weekday: style, timeZone: "UTC" }).format(
      Date.UTC(2001, 0, 7 + day)
    );
  const twelve = Array.from({ length: 12 }, (_, i) => i);
  const seven = Array.from({ length: 7 }, (_, i) => i);
  return {
    months: twelve.map(m => monthOf(m, "long")),
    shortMonths: twelve.map(m => monthOf(m, "short")),
    days: seven.map(d => dayOf(d, "long")),
    shortDays: seven.map(d => dayOf(d, "short"))
  };
}

/**
 * The switches of a placeholder expression — `amount` and `\# "0.00"` from
 * `amount \# "0.00"` — or `undefined` when it has none. A switch starts at a
 * backslash after white space, so a key that holds a space still reads whole.
 */
export function splitSwitches(expr: string): { path: string; switches: string } | undefined {
  const m = /^(.*?\S)\s+(\\[\s\S]*)$/.exec(expr);
  if (!m) {
    return undefined;
  }
  // Word turns the quotes an author types into typographic ones.
  const switches = m[2]!
    .replace(/[\u201C\u201D\u201E\u201F]/g, '"')
    .replace(/[\u2018\u2019\u201A\u201B]/g, "'");
  return { path: m[1]!, switches };
}

/** What is wrong with `switches` as a placeholder's switches, or `undefined`. */
export function checkSwitches(switches: string): string | undefined {
  const tokens = tokenizeFieldCode(switches);
  const open = tokens.find(t => t.unterminated);
  if (open !== undefined) {
    return `the quote before "${open.text}" is not closed`;
  }
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    if (t.quoted || !t.text.startsWith("\\")) {
      return `"${t.text}" is not a switch (one of \\# \\@ \\* \\b \\f, each with its argument)`;
    }
    if (!SWITCHES.has(t.text)) {
      return `${t.text} is not a switch a placeholder takes (one of \\# \\@ \\* \\b \\f)`;
    }
    const arg = tokens[i + 1];
    if (!arg || (!arg.quoted && arg.text.startsWith("\\"))) {
      return `${t.text} needs an argument`;
    }
    if (t.text === "\\*" && !isGeneralFormat(arg.text)) {
      return `\\* ${arg.text} is not a format Word has`;
    }
    i++;
  }
  return undefined;
}

/**
 * A value's text as a field reads it: a `Date` as its UTC date and time,
 * anything else as the template writes it.
 */
function switchInput(value: unknown, text: string): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return isoText(fieldDateOf(value));
  }
  return text;
}

function fieldDateOf(d: Date): FieldDate {
  return {
    year: d.getUTCFullYear(),
    month: d.getUTCMonth() + 1,
    day: d.getUTCDate(),
    hour: d.getUTCHours(),
    minute: d.getUTCMinutes(),
    second: d.getUTCSeconds()
  };
}

function isoText(d: FieldDate): string {
  const pad = (n: number): string => String(n).padStart(2, "0");
  return `${String(d.year).padStart(4, "0")}-${pad(d.month)}-${pad(d.day)}T${pad(d.hour)}:${pad(d.minute)}:${pad(d.second)}`;
}

/**
 * `text` as a date, read the ways data writes one: what Word reads (ISO, US),
 * an XML Schema `date` or `dateTime` — with fractional seconds, a zone or both,
 * the time taken as written (the zone does not move it), and `24:00:00` as the
 * start of the next day — and SAP's `YYYYMMDD`. A part out of range (a 30
 * February, a minute 60, a zone `+15:00`) is not a date.
 */
function readDate(text: string): FieldDate | undefined {
  const t = text.trim();
  const word = parseFieldDate(t);
  if (word) {
    return word;
  }
  // The zone is optional: .NET and Java write `2026-11-07T10:00:00.123` without one. Fractional seconds follow
  // seconds, as XML Schema writes them.
  const xs =
    /^(\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?)?(Z|[+-](\d{2}):(\d{2}))?$/.exec(
      t
    );
  if (xs) {
    const [, y, mo, d, h, mi, sec, fraction, zone, zh, zm] = xs;
    if (zone !== undefined && zone !== "Z") {
      const hours = Number(zh);
      const minutes = Number(zm);
      if (minutes > 59 || hours > 14 || (hours === 14 && minutes !== 0)) {
        return undefined;
      }
    }
    if (h === "24") {
      // XML Schema's end of day: only `24:00:00`, the next day's midnight.
      if (mi !== "00" || sec !== "00" || !/^0*$/.test(fraction ?? "")) {
        return undefined;
      }
      const day = fieldDate(Number(y), Number(mo), Number(d));
      return day === undefined ? undefined : nextDay(day);
    }
    return fieldDate(
      Number(y),
      Number(mo),
      Number(d),
      Number(h ?? 0),
      Number(mi ?? 0),
      Number(sec ?? 0)
    );
  }
  const compact = /^(\d{4})(\d{2})(\d{2})$/.exec(t);
  if (compact) {
    return fieldDate(Number(compact[1]), Number(compact[2]), Number(compact[3]));
  }
  return undefined;
}

/** Midnight of the day after `date`. */
function nextDay(date: FieldDate): FieldDate {
  const d = new Date(0);
  d.setUTCFullYear(date.year, date.month - 1, date.day + 1);
  return fieldDate(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate())!;
}

/**
 * `text` — what `value` reads as — formatted by `switches` (already checked),
 * or the problem when the value cannot take a picture: `\#` on a value that
 * is not a number, `\@` on one that is not a date. An empty value stays empty,
 * without the `\b` / `\f` text, as in a merge field.
 */
export function applySwitches(
  value: unknown,
  text: string,
  switches: string,
  format: FieldFormat
): { readonly text: string } | { readonly problem: string } {
  let input = switchInput(value, text);
  if (input === "") {
    return { text: "" };
  }
  const tokens = tokenizeFieldCode(switches);
  if (fieldSwitchArgument(tokens, "\\#") !== undefined && parseFieldNumber(input) === undefined) {
    return { problem: `${JSON.stringify(input)} is not a number` };
  }
  if (fieldSwitchArgument(tokens, "\\@") !== undefined) {
    const date = readDate(input);
    if (!date) {
      return { problem: `${JSON.stringify(input)} is not a date` };
    }
    input = isoText(date);
  }
  const before = fieldSwitchArgument(tokens, "\\b") ?? "";
  const after = fieldSwitchArgument(tokens, "\\f") ?? "";
  return { text: before + formatFieldResult(input, switches, format) + after };
}
