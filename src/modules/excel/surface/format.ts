/**
 * `Format.*` — parse a raw value against a specific cell's own `numFmt`,
 * the inverse of the library's internal display-formatting logic. Useful
 * for callers ingesting external data (CSV, API JSON, etc.) into a
 * pre-formatted template: convert a string to a real Date/number using
 * whatever date/time layout the target cell was already set up with,
 * rather than guessing the input string's shape.
 */
import { argumentChecks } from "@excel/utils/arguments";
import { parseValueByFormat as parseValueImpl } from "@excel/utils/cell-format-parse";

/**
 * Parse `input` against the number-format code `fmt`. Returns a `Date` for a date (or date-and-time) format, a
 * fraction-of-day number (0..1) for a time-of-day format, or `undefined` when the format carries no date or time
 * meaning or the input does not conform to it.
 *
 * @throws {ExcelArgumentError} when `fmt` or `input` is not a string.
 */
export function parseValue(fmt: string, input: string): number | Date | undefined {
  const check = argumentChecks("Format.parseValue");
  check.required("fmt", fmt);
  check.string("fmt", fmt);
  check.required("input", input);
  check.string("input", input);
  return parseValueImpl(fmt, input);
}

/** @deprecated Use {@link parseValue}. Removed in 1.0. */
export const parseValueByFormat = parseValue;
