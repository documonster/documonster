import { ExcelArgumentError } from "@excel/errors";
import { checkSchema } from "@excel/utils/arguments";
/**
 * `DateSerial.*` — Excel date serials that are not in a cell: a number from a CSV export, a formula
 * result, a streaming reader's raw value, or one bound for any of those.
 *
 * `import { DateSerial } from "documonster/excel"` →
 *   `DateSerial.toParts(45306)`, `DateSerial.fromParts({ year: 2024, month: 1, day: 15 })`.
 *
 * A serial is a calendar value — a date on a wall calendar, a time on a wall clock — so it converts
 * to and from calendar fields (`ExcelDateTimeParts`) without a `Date` anywhere, and no timezone can
 * reach the answer. For a value in a cell, `Cell.getDateParts` / `Cell.setDateParts` read and write
 * the same fields, and also the number format that says which kind of date the cell is.
 *
 * The 1900 epoch is the default; pass `{ date1904: true }` for a workbook that uses the 1904 one
 * (`Workbook.getModel(wb).properties.date1904`). Excel's fictitious 1900-02-29 is serial 60, as it
 * is to Excel: `toParts(60)` names it and `fromParts` accepts it. The serials before it are a day
 * apart from the calendar, again as in Excel.
 *
 * **The range is Excel's**: from serial 0 — 1900-01-00 under the 1900 epoch, 1904-01-01 under the
 * 1904 one — to the last millisecond of 9999-12-31. Outside it Excel shows `####`, so a number
 * there is not a date to Excel, and is refused rather than read as one: extrapolated, `1e9` is
 * the year 2,739,807 and `1e305` is no date at all.
 */
import type { ExcelDateTimeParts } from "@utils/excel-serial";
import { excelDaysInMonth, partsToSerial, serialToParts } from "@utils/excel-serial";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

/** Options of `DateSerial.toParts` and `DateSerial.fromParts`. */
export interface ExcelDateSerialOptions {
  /** The serial counts from the 1904 epoch (1904-01-01) rather than the 1900 one. Default `false`. */
  readonly date1904?: boolean;
}

const optionsShape = s.object({ date1904: s.boolean().optional() }, { readonly: true });
const _optionsMatch: SchemaMatches<typeof optionsShape, ExcelDateSerialOptions> = true;
void _optionsMatch;
const optionsSchema = optionsShape.optional();

/** The last year Excel has a serial for. */
const MAX_YEAR = 9999;

/** The first serial past Excel's range: 10000-01-01. */
const END_1900 = partsToSerial({ year: MAX_YEAR + 1, month: 1, day: 1 }, false);
const END_1904 = partsToSerial({ year: MAX_YEAR + 1, month: 1, day: 1 }, true);

const serialSchema = s.number({ finite: true });

const field = (min?: number, max?: number) => s.number({ integer: true, min, max }).optional();
const partsSchema = s.object(
  {
    year: field(1900, MAX_YEAR),
    month: field(1, 12),
    day: field(1, 31),
    hour: field(0, 23),
    minute: field(0, 59),
    second: field(0, 59),
    millisecond: field(0, 999)
  },
  { unknownKeys: "reject", expected: "calendar fields ({ year, month, day, hour, … })" }
);

/** The refusal of a serial outside Excel's range, from {@link toParts}. */
function outOfRange(serial: number, end: number, date1904: boolean): ExcelArgumentError {
  return new ExcelArgumentError(
    "serial",
    `DateSerial.toParts: serial must be at least 0 and below ${end} (10000-01-01 under the ` +
      `${date1904 ? "1904" : "1900"} epoch), got ${serial}`
  );
}

/**
 * The calendar fields serial `serial` names. The fraction is the time of day, to the millisecond.
 *
 * @throws {ExcelArgumentError} When `serial` is not a number in Excel's range (see the namespace).
 */
export function toParts(serial: number, options?: ExcelDateSerialOptions): ExcelDateTimeParts {
  checkSchema("DateSerial.toParts", "serial", serialSchema, serial);
  checkSchema("DateSerial.toParts", "options", optionsSchema, options);
  const date1904 = options?.date1904 === true;
  const end = date1904 ? END_1904 : END_1900;
  if (serial < 0 || serial >= end) {
    throw outOfRange(serial, end, date1904);
  }
  const parts = serialToParts(serial, date1904);
  // A fraction within a millisecond of the end rounds to 10000-01-01
  if (parts.year > MAX_YEAR) {
    throw outOfRange(serial, end, date1904);
  }
  return parts;
}

/**
 * The serial of calendar fields: a date (`year`, `month`, `day`), a time of day (`hour`, `minute`, …), or
 * both. A missing `month` or `day` is 1 and a missing time field 0, so `{ year: 2024 }` is
 * 2024-01-01 at midnight; a time alone is a fraction of a day.
 *
 * Fields are checked rather than carried: `{ month: 2, day: 31 }` is refused, not read as March 2.
 * (Excel's `DATE` carries; this converts.)
 *
 * @throws {ExcelArgumentError} When a field is not an integer in its range — `year` from 1900 (1904
 *   under the 1904 epoch) to 9999 — a day is past the end of its month, `month` or `day` is given
 *   without `year`, or no field is given at all.
 */
export function fromParts(
  parts: Partial<ExcelDateTimeParts>,
  options?: ExcelDateSerialOptions
): number {
  checkSchema("DateSerial.fromParts", "parts", partsSchema, parts);
  checkSchema("DateSerial.fromParts", "options", optionsSchema, options);
  const date1904 = options?.date1904 === true;
  const refuse = (message: string): never => {
    throw new ExcelArgumentError("parts", `DateSerial.fromParts: ${message}`);
  };
  const { year, month, day } = parts;
  if (Object.values(parts).every(value => value === undefined)) {
    refuse("parts must hold at least a date (year, month, day) or a time (hour, minute, …)");
  }
  if (year === undefined && (month !== undefined || day !== undefined)) {
    refuse("parts.year is required when parts.month or parts.day is given");
  }
  if (date1904 && year !== undefined && year < 1904) {
    refuse(`the 1904 epoch starts at 1904-01-01, got parts.year ${year}`);
  }
  if (year !== undefined && day !== undefined) {
    // Excel's calendar, in which February 1900 has the 29 days Excel gives it
    const limit = excelDaysInMonth(year, month ?? 1, date1904);
    if (day > limit) {
      refuse(
        `${year}-${String(month ?? 1).padStart(2, "0")} has ${limit} days, got parts.day ${day}`
      );
    }
  }
  // A time of day is the fraction of a day after serial 0, whose date depends on the epoch
  const zero = serialToParts(0, date1904);
  return partsToSerial(
    {
      year: year ?? zero.year,
      month: month ?? (year === undefined ? zero.month : 1),
      day: day ?? (year === undefined ? zero.day : 1),
      hour: parts.hour ?? 0,
      minute: parts.minute ?? 0,
      second: parts.second ?? 0,
      millisecond: parts.millisecond ?? 0
    },
    date1904
  );
}
