/**
 * `Format.*` — parse a raw value against a specific cell's own `numFmt`,
 * the inverse of the library's internal display-formatting logic. Useful
 * for callers ingesting external data (CSV, API JSON, etc.) into a
 * pre-formatted template: convert a string to a real Date/number using
 * whatever date/time layout the target cell was already set up with,
 * rather than guessing the input string's shape.
 */
import { parseValueByFormat } from "@excel/utils/cell-format-parse";

export { parseValueByFormat };
