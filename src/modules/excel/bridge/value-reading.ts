/**
 * How a text read — CSV or a Markdown table — turns a cell's text into a cell value: a number, a date in one of the
 * named forms, `true` / `false`, an error literal, or the text itself; an empty cell is empty.
 */
import type { DecimalSeparator } from "@csv/utils/number";
import { parseNumberFromCsv } from "@csv/utils/number";
import type { Worksheet } from "@excel/core/worksheet";
import { addRow } from "@excel/core/worksheet";
import { get_nextRow } from "@excel/core/worksheet-core";
import type { CellErrorValue, CellValue, CellValueInput } from "@excel/types";
import { tableRowError } from "@excel/utils/file-content-errors";
import type { DateFormat } from "@utils/datetime";
import { createDateParser } from "@utils/datetime";

/**
 * The formats a CSV date column is tried against, when the caller names none.
 *
 * **The millisecond form is first because it is the one this bridge's own writer emits.** It was absent, and
 * the omission was invisible only because the parsers it fell through to did not check their input's width:
 * `2020-01-15T00:00:00.000Z` was matched by the *nineteen-character* `YYYY-MM-DD[T]HH:mm:ss` parser, which
 * consumed the prefix, ignored the `Z` and built the date in **local time**. A writer and a reader that cannot
 * agree on a format, held together by a parser being too permissive to notice.
 */
export const DEFAULT_DATE_FORMATS: readonly DateFormat[] = [
  "YYYY-MM-DD[T]HH:mm:ss.SSSZ",
  "YYYY-MM-DD[T]HH:mm:ssZ",
  "YYYY-MM-DD[T]HH:mm:ss",
  "MM-DD-YYYY",
  "YYYY-MM-DD"
];

const SpecialValues: Record<string, boolean | CellErrorValue> = {
  true: true,
  false: false,
  "#N/A": { error: "#N/A" },
  "#REF!": { error: "#REF!" },
  "#NAME?": { error: "#NAME?" },
  "#DIV/0!": { error: "#DIV/0!" },
  "#NULL!": { error: "#NULL!" },
  "#VALUE!": { error: "#VALUE!" },
  "#NUM!": { error: "#NUM!" }
};

// =============================================================================
// Value mappers
// =============================================================================

export function createDefaultValueMapper(
  dateFormats: readonly DateFormat[],
  options?: { decimalSeparator?: DecimalSeparator }
): (datum: CellValue) => CellValue {
  // A text with no timezone is the calendar value it names, read in UTC: a cell's `Date` is UTC wall-clock time
  // everywhere else in Excel, so a local reading put a different serial in the file on every machine.
  const dateParser = createDateParser(dateFormats, { utc: true });
  const decimalSeparator: DecimalSeparator = options?.decimalSeparator ?? ".";

  return function mapValue(datum: CellValue): CellValue {
    if (datum === "") {
      return null;
    }

    if (typeof datum === "string") {
      const datumNumber = parseNumberFromCsv(datum, decimalSeparator);
      if (!Number.isNaN(datumNumber) && datumNumber !== Infinity) {
        return datumNumber;
      }
    } else {
      const datumNumber = Number(datum);
      if (!Number.isNaN(datumNumber) && datumNumber !== Infinity) {
        return datumNumber;
      }
    }

    if (typeof datum === "string") {
      const date = dateParser.parse(datum);
      if (date) {
        return date;
      }

      const special = SpecialValues[datum];
      if (special !== undefined) {
        return special;
      }
    }

    return datum;
  };
}

/**
 * Add one row of a CSV or Markdown input to `worksheet` — the bridges' only way in, so a row the input holds that a
 * worksheet cannot (wider than 16 384 columns, past the last row) is reported as the input's `ExcelFileError` naming
 * the row, not as the argument error `addRow` raises for a caller's own values (`utils/file-content-errors`).
 */
export function addInputRow(
  worksheet: Worksheet,
  values: readonly (CellValue | CellValueInput)[],
  fileKind: "CSV" | "Markdown"
): void {
  const row = get_nextRow(worksheet);
  try {
    addRow(worksheet, values as CellValue[]);
  } catch (error) {
    throw tableRowError(error, row, fileKind);
  }
}
