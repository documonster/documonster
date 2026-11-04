/**
 * CSV Number Utilities
 *
 * Functions for parsing and formatting numbers in CSV files,
 * with support for different decimal separators (e.g., European format).
 */

import { CsvArgumentError } from "@csv/errors";
import { describeValue } from "@csv/utils/check-options";

/**
 * Decimal separator type.
 * - "." - Standard decimal point (default)
 * - "," - European decimal comma
 */
export type DecimalSeparator = "." | ",";

/** Pre-compiled regex for European decimal comma format */
const COMMA_DECIMAL_REGEX = /^-?\d+(,\d+)?([eE][+-]?\d+)?$/;

/**
 * Format a number for CSV output with the specified decimal separator.
 *
 * @param value - The number to format
 * @param decimalSeparator - The decimal separator to use
 * @returns Formatted string representation
 *
 * @example
 * formatNumberForCsv(3.14, ".") // "3.14"
 * formatNumberForCsv(3.14, ",") // "3,14"
 */
export function formatNumberForCsv(value: number, decimalSeparator: DecimalSeparator): string {
  if (decimalSeparator !== ",") {
    return String(value);
  }
  // Replace decimal point with comma - faster than split().join()
  return String(value).replace(".", ",");
}

/**
 * Parse a CSV string value as a number with the specified decimal separator.
 *
 * @param value - The string value to parse
 * @param decimalSeparator - The decimal separator used in the value
 * @returns Parsed number; NaN when the text is not a number, including an empty or blank one
 *
 * @example
 * parseNumberFromCsv("3.14", ".") // 3.14
 * parseNumberFromCsv("3,14", ",") // 3.14
 */
export function parseNumberFromCsv(value: string, decimalSeparator: DecimalSeparator): number {
  const trimmed = value.trim();
  // `Number("")` is 0: an empty or blank field holds no number.
  if (trimmed === "") {
    return NaN;
  }

  if (decimalSeparator !== ",") {
    return Number(trimmed);
  }

  // Minimal locale support: treat a single comma as the decimal separator.
  // Common EU CSV uses delimiter ';' and decimal ',' (e.g. 12,34).
  if (COMMA_DECIMAL_REGEX.test(trimmed)) {
    return Number(trimmed.replace(",", "."));
  }

  return Number(trimmed);
}

function checkDecimalSeparator(where: string, decimalSeparator: unknown): void {
  if (decimalSeparator !== "." && decimalSeparator !== ",") {
    throw new CsvArgumentError(
      "decimalSeparator",
      `${where}: decimalSeparator must be "." or ",", got ${describeValue(decimalSeparator)}`
    );
  }
}

/**
 * Format a number with the given decimal separator — {@link formatNumberForCsv}, checking its
 * arguments. Published as `Csv.formatNumber`.
 *
 * @example
 * ```ts
 * import { Csv } from "documonster/csv";
 * Csv.formatNumber(3.14, ","); // "3,14"
 * ```
 */
export function formatNumber(value: number, decimalSeparator: DecimalSeparator): string {
  if (typeof value !== "number") {
    throw new CsvArgumentError(
      "value",
      `Csv.formatNumber: value must be a number, got ${describeValue(value)}`
    );
  }
  checkDecimalSeparator("Csv.formatNumber", decimalSeparator);
  return formatNumberForCsv(value, decimalSeparator);
}

/**
 * Parse a number written with the given decimal separator — {@link parseNumberFromCsv},
 * checking its arguments. `NaN` when the text is not a number, including an empty one.
 * Published as `Csv.parseNumber`.
 *
 * @example
 * ```ts
 * import { Csv } from "documonster/csv";
 * Csv.parseNumber("3,14", ","); // 3.14
 * ```
 */
export function parseNumber(value: string, decimalSeparator: DecimalSeparator): number {
  if (typeof value !== "string") {
    throw new CsvArgumentError(
      "value",
      `Csv.parseNumber: value must be a string, got ${describeValue(value)}`
    );
  }
  checkDecimalSeparator("Csv.parseNumber", decimalSeparator);
  return parseNumberFromCsv(value, decimalSeparator);
}
