/**
 * Display text for a cell value under its number format.
 *
 * Reading and rendering a number format is not done here: both live at Layer 0 in
 * `@utils/number-format` and `@utils/number-format-render`, shared with the readers' date
 * detection and the `TEXT` worksheet function, so that a character is a literal or a code for all
 * of them at once. This file adds only what is specific to a cell — a `Date` value, a default
 * date format for an unformatted date, and the date-kind classification `Cell.getDateKind` reports.
 */

import type { NumFmt } from "@excel/types";
import { isGeneralFormat, numberFormatFacets } from "@utils/number-format";
import { renderNumberFormat } from "@utils/number-format-render";
import { dateToExcel } from "@utils/utils";

/**
 * Format a value according to an Excel number format, as Excel displays it.
 * @param fmt The Excel number format string (e.g. `"0.00%"`, `"#,##0"`, `"yyyy-mm-dd"`)
 * @param val The value to format; a number is an Excel serial when the format names a date
 */
export function format(fmt: string, val: number | string | boolean): string {
  if (val == null) {
    return "";
  }
  return renderNumberFormat(fmt, val);
}

/**
 * Check if format is a pure time format: a clock reading with no date part.
 * Elapsed-time formats such as `[h]:mm:ss` are excluded — they need the whole serial.
 */
export function isTimeOnlyFormat(fmt: string): boolean {
  const { elapsed, date, time, text } = numberFormatFacets(fmt);
  return !text && !elapsed && !date && time;
}

/**
 * Check if format names a calendar date: a year, a day, or an `m` that reads as a month.
 * Elapsed-time formats are not dates.
 */
export function isDateDisplayFormat(fmt: string): boolean {
  const { elapsed, date, text } = numberFormatFacets(fmt);
  return !text && !elapsed && date;
}

/**
 * What a number format says a date-typed cell actually holds.
 *
 * **The format is the only record of the distinction.** A serial does not say whether its whole-day part or
 * its fraction is the point, and neither does the `Date` the model stores — so this is where
 * `Cell.getDateKind` and `Cell.getTemporal` get their answer, and why a time-of-day written without a format
 * renders as `12-30-1899`.
 *
 * **`"unknown"` and `"duration"` are real answers, not failures to produce one.** This returned `"dateTime"`
 * for an unformatted cell, a `General` cell, a plain number format and an elapsed-time format alike — which
 * dressed "I cannot tell" and "this is a length of time" up as a specific calendar kind, and left
 * `Cell.getTemporal` manufacturing a `PlainDateTime` for `[h]:mm:ss` that named a moment in 1899 for a value
 * meaning thirty-six hours. A caller can now see the difference and decide.
 */
export type DateFormatKind = "date" | "time" | "dateTime" | "duration" | "unknown";

export function dateFormatKind(fmt: string | undefined): DateFormatKind {
  if (fmt === undefined || isGeneralFormat(fmt)) {
    return "unknown";
  }
  const { elapsed, date, time, text } = numberFormatFacets(fmt);
  // A text format (`yyyy@`) shows a number as text, so a reader keeps it a number: no date kind either.
  if (text) {
    return "unknown";
  }
  if (elapsed) {
    return "duration";
  }
  if (date && time) {
    return "dateTime";
  }
  if (date) {
    return "date";
  }
  return time ? "time" : "unknown";
}

/**
 * Default format applied to Date values whose numFmt is `General` or empty.
 *
 * Excel itself substitutes a locale-dependent short date in this case (US:
 * `m/d/yyyy`). We pick an ISO-like `yyyy-mm-dd` so consumers who never set a
 * `numFmt` still get a sensible, unambiguous rendering instead of the raw
 * Excel serial number.
 */
const DEFAULT_DATE_FORMAT = "yyyy-mm-dd";
const DEFAULT_DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss";

/**
 * Format a value according to the given format string.
 * Handles Date objects with timezone-independent Excel serial conversion.
 *
 * `date1904` is the workbook's date system. A `Date` read from a 1904 workbook has to become the same serial
 * again: a calendar date survives either epoch, because its fields come back out unchanged, but an elapsed
 * format reads the whole serial — and `[h]:mm` showed 36 hours as 35124 when the epoch was dropped here.
 */
export function formatCellValue(
  value: Date | number | boolean | string,
  fmt: string,
  dateFormat?: string,
  date1904 = false
): string {
  if (value instanceof Date) {
    let serial = dateToExcel(value, date1904);
    if (isTimeOnlyFormat(fmt)) {
      serial = serial % 1;
      if (serial < 0) {
        serial += 1;
      }
      return renderNumberFormat(fmt, serial, { date1904 });
    }
    // For Date values whose numFmt is missing or General, Excel substitutes a
    // default short-date format. Without this, `format("General", serial)`
    // would emit the raw Excel serial (e.g. "43567") — almost never what the
    // caller wants. Pick a datetime-aware default based on whether the value
    // carries a non-midnight time component.
    let effectiveFmt: string;
    if (dateFormat && isDateDisplayFormat(fmt)) {
      effectiveFmt = dateFormat;
    } else if (isGeneralFormat(fmt)) {
      effectiveFmt = serial % 1 === 0 ? DEFAULT_DATE_FORMAT : DEFAULT_DATETIME_FORMAT;
    } else {
      effectiveFmt = fmt;
    }
    return renderNumberFormat(effectiveFmt, serial, { date1904 });
  }
  return renderNumberFormat(fmt, value, { date1904 });
}

// =============================================================================
// Cell Display Text
// =============================================================================

/** Minimal cell shape needed by {@link getCellDisplayText} — avoids importing Cell. */
interface CellLike {
  value: unknown;
  numFmt: string | NumFmt | undefined;
  text: string;
  /** The workbook's date system; the 1900 system when absent. */
  date1904?: boolean;
}

/**
 * Get the formatted display text for a cell value.
 *
 * Handles primitive values, Date objects, formula results, and falls back to
 * `cell.text` for complex types (rich text, hyperlinks, errors, etc.).
 *
 * @param cell       - A cell (or cell-like object) with `.value`, `.numFmt`, and `.text`
 * @param dateFormat - Optional custom date format override
 */
export function getCellDisplayText(cell: CellLike, dateFormat?: string): string {
  const value = cell.value;
  const numFmt = cell.numFmt;
  const fmt = typeof numFmt === "string" ? numFmt : (numFmt?.formatCode ?? "General");

  if (value == null) {
    return "";
  }

  if (
    value instanceof Date ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "string"
  ) {
    return formatCellValue(value, fmt, dateFormat, cell.date1904);
  }

  // Formula type — use the result value. A shared-formula clone carries `sharedFormula` and no `formula`,
  // and is formatted exactly like the cell that defines it.
  if (typeof value === "object" && ("formula" in value || "sharedFormula" in value)) {
    const result = (value as { result?: unknown }).result;
    if (result == null) {
      return "";
    }
    if (
      result instanceof Date ||
      typeof result === "number" ||
      typeof result === "boolean" ||
      typeof result === "string"
    ) {
      return formatCellValue(result, fmt, dateFormat, cell.date1904);
    }
  }

  // Fallback to cell.text for other types (rich text, hyperlink, error, etc.)
  return cell.text;
}
