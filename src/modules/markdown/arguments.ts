/**
 * Argument checks for `Markdown.parse`, `Markdown.parseAll` and `Markdown.format`: a value
 * with no reading is a {@link MarkdownArgumentError} naming the call and the field, thrown before
 * any output is built, rather than a `TypeError` from deep inside the formatter or an
 * alignment that silently became `"none"`.
 */

import { MarkdownArgumentError } from "@markdown/errors";
import type {
  MarkdownColumnConfig,
  MarkdownFormatOptions,
  MarkdownParseOptions
} from "@markdown/types";
import { check, fail, s } from "@utils/schema";
import type { SchemaCheckOptions, SchemaMatches } from "@utils/schema";

function argumentError(field: string, message: string): MarkdownArgumentError {
  return new MarkdownArgumentError(field, message);
}

/** `field` is the parameter, named in the message by its path (`options.trim`). */
function at(where: string, field: string): SchemaCheckOptions {
  return { where, field, error: argumentError };
}

const flag = () => s.boolean().optional();

const parseOptionsSchema = s.object(
  {
    trim: flag(),
    unescape: flag(),
    skipEmptyRows: flag(),
    maxRows: s.number({ integer: true, min: 0 }).optional(),
    convertBr: flag()
  },
  { expected: "an options object" }
);
const _parseOptionsMatch: SchemaMatches<typeof parseOptionsSchema, MarkdownParseOptions> = true;
void _parseOptionsMatch;

const alignment = () => s.oneOf(["left", "center", "right", "none"]).optional();

const columnSchema = s.object({
  header: s.string(),
  alignment: alignment(),
  minWidth: s.number({ integer: true, min: 0 }).optional()
});
const _columnMatch: SchemaMatches<typeof columnSchema, MarkdownColumnConfig> = true;
void _columnMatch;

const formatOptionsSchema = s.object(
  {
    columns: s
      .readonlyArray(
        s.union([s.string(), columnSchema], {
          expected: "a header string or { header, alignment?, minWidth? }"
        }),
        { expected: "an array of headers or column configs" }
      )
      .optional(),
    alignment: alignment(),
    padding: flag(),
    trailingNewline: flag(),
    escapeContent: flag(),
    stringify: s.function<(value: unknown) => string>().optional()
  },
  { expected: "an options object" }
);
const _formatOptionsMatch: SchemaMatches<typeof formatOptionsSchema, MarkdownFormatOptions> = true;
void _formatOptionsMatch;

const inputSchema = s.string();

/** Check the arguments of `Markdown.parse` and `Markdown.parseAll`. */
export function checkParseArguments(
  where: string,
  input: unknown,
  options: MarkdownParseOptions
): void {
  check(inputSchema, input, at(where, "input"));
  check(parseOptionsSchema, options, at(where, "options"));
}

const headersSchema = s.array(s.string(), { expected: "an array of strings" });
const rowsSchema = s.array(s.array(s.unknown(), { expected: "an array of cell values" }), {
  expected: "an array of rows"
});

/** Check the arguments of `Markdown.format`. */
export function checkFormatArguments(
  headers: unknown,
  rows: unknown,
  options: MarkdownFormatOptions
): void {
  const where = "Markdown.format";
  check(headersSchema, headers, at(where, "headers"));
  check(rowsSchema, rows, at(where, "rows"));
  check(formatOptionsSchema, options, at(where, "options"));
}

/** The text `stringify` returned for a cell, which has to be a string. */
export function checkStringified(value: unknown, row: number, column: number): string {
  if (typeof value !== "string") {
    throw fail(
      at("Markdown.format", "rows"),
      `[${row}][${column}]`,
      "turned into a string by options.stringify",
      value
    );
  }
  return value;
}
