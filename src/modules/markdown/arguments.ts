/**
 * Argument checks for `Markdown.parse`, `Markdown.parseAll` and `Markdown.format`: a value
 * with no reading is a {@link MarkdownArgumentError} naming the call and the field, thrown before
 * any output is built, rather than a `TypeError` from deep inside the formatter or an
 * alignment that silently became `"none"`.
 */

import { MarkdownArgumentError } from "@markdown/errors";
import type { MarkdownFormatOptions } from "@markdown/types";
import { describeValue } from "@utils/describe-value";
export { describeValue } from "@utils/describe-value";

const ALIGNMENTS: readonly string[] = ["left", "center", "right", "none"];

/** The argument path a name stands for: `alignment` → `options.alignment`. */
function fieldOf(name: string): string {
  return /^(?:headers|rows|options)\b/.test(name) ? name : `options.${name}`;
}

function fail(field: string, expected: string, value: unknown): never {
  throw new MarkdownArgumentError(
    fieldOf(field),
    `Markdown.format: ${field} must be ${expected}, got ${describeValue(value)}`
  );
}

function checkAlignment(field: string, value: unknown): void {
  if (value !== undefined && (typeof value !== "string" || !ALIGNMENTS.includes(value))) {
    fail(field, `one of ${ALIGNMENTS.map(a => JSON.stringify(a)).join(", ")}`, value);
  }
}

/** Check the arguments of `Markdown.format`. */
export function checkFormatArguments(
  headers: unknown,
  rows: unknown,
  options: MarkdownFormatOptions
): void {
  if (!Array.isArray(headers)) {
    fail("headers", "an array of strings", headers);
  }
  headers.forEach((header, i) => {
    if (typeof header !== "string") {
      fail(`headers[${i}]`, "a string", header);
    }
  });
  if (!Array.isArray(rows)) {
    fail("rows", "an array of rows", rows);
  }
  rows.forEach((row, i) => {
    if (!Array.isArray(row)) {
      fail(`rows[${i}]`, "an array of cell values", row);
    }
  });
  if (typeof options !== "object" || options === null || Array.isArray(options)) {
    fail("options", "an options object", options);
  }
  checkAlignment("alignment", options.alignment);
  for (const name of ["padding", "trailingNewline", "escapeContent"] as const) {
    const value = options[name];
    if (value !== undefined && typeof value !== "boolean") {
      fail(name, "true or false", value);
    }
  }
  if (options.stringify !== undefined && typeof options.stringify !== "function") {
    fail("stringify", "a function", options.stringify);
  }
  if (options.columns === undefined) {
    return;
  }
  if (!Array.isArray(options.columns)) {
    fail("columns", "an array of headers or column configs", options.columns);
  }
  options.columns.forEach((column, i) => {
    if (typeof column === "string") {
      return;
    }
    if (typeof column !== "object" || column === null) {
      fail(`columns[${i}]`, "a header string or { header, alignment?, minWidth? }", column);
    }
    if (typeof column.header !== "string") {
      fail(`columns[${i}].header`, "a string", column.header);
    }
    checkAlignment(`columns[${i}].alignment`, column.alignment);
    const { minWidth } = column;
    if (
      minWidth !== undefined &&
      (typeof minWidth !== "number" || !Number.isInteger(minWidth) || minWidth < 0)
    ) {
      fail(`columns[${i}].minWidth`, "an integer >= 0", minWidth);
    }
  });
}

/** The text `stringify` returned for a cell, which has to be a string. */
export function checkStringified(value: unknown, row: number, column: number): string {
  if (typeof value !== "string") {
    fail(`rows[${row}][${column}]`, "turned into a string by options.stringify", value);
  }
  return value;
}
