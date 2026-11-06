/**
 * Option checks shared by every CSV parser and formatter.
 *
 * Parsing and formatting each resolve their options in one place (`createParseConfig`,
 * `createFormatConfig`), so the checks live there and cover `Csv.parse`, `Csv.parseAsync`,
 * `Csv.parseStream`, the streams and the worker alike. Each refusal is a {@link CsvArgumentError}
 * naming the option, thrown before any input is read: a formatter given `delimiter: ""`
 * wrote rows with nothing between the fields, and `quote` equal to the delimiter wrote a
 * file no reader can split.
 */

import { CsvArgumentError } from "@csv/errors";
import type {
  CastDateConfig,
  CsvFormatOptions,
  CsvParseOptions,
  HeaderTransformFunction,
  OnSkipCallback,
  TypeTransformMap
} from "@csv/types";
import type { CsvDetectDelimiterOptions, CsvDetectLinebreakOptions } from "@csv/utils/detect";
import { describeValue } from "@utils/describe-value";
import { check, fail as schemaFail, s } from "@utils/schema";
import type { SchemaCheckOptions, SchemaMatches } from "@utils/schema";
export { describeValue } from "@utils/describe-value";

/** The {@link SchemaCheckOptions.error} of every CSV check. */
export function argumentError(field: string, message: string): CsvArgumentError {
  return new CsvArgumentError(field, message);
}

/** How an option of an options object is reported: `field` and message both `options.x`. */
function optionsAt(where: string): SchemaCheckOptions {
  return { where, field: "options", error: argumentError };
}

function fail(where: string, option: string, expected: string, value: unknown): never {
  throw schemaFail(optionsAt(where), `.${option}`, expected, value);
}

const hasNoLineBreak = (value: string): boolean => !value.includes("\n") && !value.includes("\r");

/** A quote or escape option: one character, or `false`/`null` to disable. */
function charOption(allowEmpty: boolean) {
  return s
    .custom<string | false | null>(
      v =>
        v === false ||
        v === null ||
        (typeof v === "string" && v.length <= 1 && (allowEmpty || v !== "")),
      "a single character, or false to disable"
    )
    .refine(v => v !== "\r" && v !== "\n", "a character other than a line break")
    .optional();
}

const count = (min: number) => s.number({ integer: true, min }).optional();
const flag = () => s.boolean().optional();

/**
 * A `dynamicTyping` setting: a boolean, or a map whose own entries are booleans or converter
 * functions. Only own entries are ever looked up, so a column named `toString` is not read
 * through `Object.prototype`.
 */
export const dynamicTypingSchema = s.union(
  [
    s.boolean(),
    s.record(
      s.union([s.boolean(), s.function<(value: string) => unknown>()], {
        expected: "true, false or a function"
      })
    )
  ],
  { expected: "true, false or a map of column name to boolean or function" }
);

export const castDateSchema = s.custom<CastDateConfig>(
  v => typeof v === "boolean" || (Array.isArray(v) && v.every(c => typeof c === "string")),
  "true, false or an array of column names"
);

const parseOptionsSchema = s.object(
  {
    delimiter: s.string().refine(hasNoLineBreak, "a string without line breaks").optional(),
    // An empty quote or escape has always meant that no character matches, so it is kept.
    quote: charOption(true),
    escape: charOption(true),
    objectMode: flag(),
    delimitersToGuess: s
      .readonlyArray(
        s.string({ nonEmpty: true }).refine(hasNoLineBreak, "a string without line breaks"),
        { expected: "an array of delimiters" }
      )
      .optional(),
    lineEnding: s.string().optional(),
    comment: s.string().optional(),
    skipEmptyLines: s.literal(true, false, "greedy").optional(),
    maxRows: count(0),
    toLine: count(0),
    skipLines: count(0),
    skipRows: count(0),
    maxRowBytes: count(1),
    chunkSize: count(1),
    columnMismatch: s
      .object(
        {
          less: s.literal("error", "pad").orUndefined(),
          more: s.literal("error", "truncate", "keep").orUndefined()
        },
        { expected: "{ less, more }" }
      )
      .optional(),
    encoding: s.string({ nonEmpty: true }).optional(),
    rowTransform: s.function<CsvParseOptions["rowTransform"]>().optional(),
    validate: s.function<CsvParseOptions["validate"]>().optional(),
    chunk: s.function<CsvParseOptions["chunk"]>().optional(),
    beforeFirstChunk: s.function<CsvParseOptions["beforeFirstChunk"]>().optional(),
    onSkip: s.function<OnSkipCallback>().optional(),
    trim: flag(),
    ltrim: flag(),
    rtrim: flag(),
    groupColumnsByName: flag(),
    fastMode: flag(),
    info: flag(),
    raw: flag(),
    relaxQuotes: flag(),
    skipRecordsWithError: flag(),
    skipRecordsWithEmptyValues: flag(),
    headers: s
      .union(
        [
          s.boolean(),
          s.function<HeaderTransformFunction>(),
          s.array(
            s.union([s.string(), s.literal(null, undefined)], {
              expected: "a string, null or undefined"
            })
          )
        ],
        { expected: "true, false, an array of header names or a function" }
      )
      .optional(),
    dynamicTyping: dynamicTypingSchema.optional(),
    castDate: castDateSchema.optional(),
    objname: s.string({ nonEmpty: true }).optional(),
    decimalSeparator: s.literal(".", ",").optional()
  },
  { expected: "an options object" }
);
const _parseOptionsMatch: SchemaMatches<typeof parseOptionsSchema, CsvParseOptions> = true;
void _parseOptionsMatch;

const quoteConfig = () =>
  s
    .union([s.boolean(), s.array(s.boolean().optional()), s.record(s.boolean().optional())], {
      expected: "true, false, an array of booleans or a map of header to boolean"
    })
    .optional();

const optionalFunction = <F>() => s.function<F>().optional();

const formatOptionsSchema = s.object(
  {
    delimiter: s
      .string({ nonEmpty: true })
      .refine(hasNoLineBreak, "a string without line breaks")
      .optional(),
    quote: charOption(false),
    // An empty escape has always fallen back to the quote character.
    escape: charOption(true),
    objectMode: flag(),
    readableObjectMode: flag(),
    lineEnding: s.string({ nonEmpty: true }).optional(),
    decimalSeparator: s.oneOf([".", ","], { expected: '"." or ","' }).optional(),
    quoteColumns: quoteConfig(),
    quoteHeaders: quoteConfig(),
    headers: s
      .union([s.boolean(), s.literal(null), s.readonlyArray(s.string())], {
        expected: "true, false or an array of header names"
      })
      .optional(),
    columns: s
      .readonlyArray(
        s.union([s.string(), s.object({ key: s.string(), header: s.string().optional() })], {
          expected: "a key or { key, header? }"
        }),
        { expected: "an array of keys or { key, header? } objects" }
      )
      .optional(),
    writeHeaders: flag(),
    firstRowIsHeader: flag(),
    bom: flag(),
    trailingNewline: flag(),
    escapeFormulae: flag(),
    typeTransform: s
      .object(
        {
          boolean: optionalFunction<TypeTransformMap["boolean"]>(),
          date: optionalFunction<TypeTransformMap["date"]>(),
          number: optionalFunction<TypeTransformMap["number"]>(),
          bigint: optionalFunction<TypeTransformMap["bigint"]>(),
          object: optionalFunction<TypeTransformMap["object"]>(),
          string: optionalFunction<TypeTransformMap["string"]>(),
          row: optionalFunction<TypeTransformMap["row"]>()
        },
        { expected: "an object of transform functions", unknownKeys: "reject" }
      )
      .optional()
  },
  { expected: "an options object" }
);
const _formatOptionsMatch: SchemaMatches<typeof formatOptionsSchema, CsvFormatOptions> = true;
void _formatOptionsMatch;

export function checkObject(where: string, value: unknown): void {
  check(s.object({}, { expected: "an options object" }), value, optionsAt(where));
}

/** The quote and escape characters must not be the delimiter, or a field cannot be read back. */
function checkDistinct(where: string, delimiter: string, quote: unknown, escape: unknown): void {
  if (delimiter === "") {
    return;
  }
  for (const [option, value] of [
    ["quote", quote],
    ["escape", escape]
  ] as const) {
    if (typeof value === "string" && value !== "" && delimiter.includes(value)) {
      fail(where, option, `a character not in the delimiter ${describeValue(delimiter)}`, value);
    }
  }
}

/**
 * A line ending that overlaps the delimiter splits fields instead of records, so it cannot be
 * read back. (A delimiter holding a line break is refused by the schema.)
 */
function checkSeparators(where: string, delimiter: string, lineEnding: string | undefined): void {
  if (
    delimiter !== "" &&
    lineEnding !== undefined &&
    lineEnding !== "" &&
    (delimiter.includes(lineEnding) || lineEnding.includes(delimiter))
  ) {
    fail(
      where,
      "lineEnding",
      `a string that does not overlap the delimiter ${describeValue(delimiter)}`,
      lineEnding
    );
  }
}

/** Check {@link CsvParseOptions}. `delimiter: ""` and `lineEnding: ""` mean auto-detect. */
export function checkParseOptions(options: CsvParseOptions, where = "Csv.parse"): void {
  check(parseOptionsSchema, options, optionsAt(where));
  const delimiter = options.delimiter ?? ",";
  checkDistinct(where, delimiter, options.quote ?? '"', options.escape);
  checkSeparators(where, delimiter, options.lineEnding);
  if (
    options.objname !== undefined &&
    (options.headers === undefined || options.headers === false)
  ) {
    fail(where, "objname", "used with headers (it names a header column)", options.objname);
  }
  if (options.decimalSeparator === "," && delimiter !== "" && delimiter.includes(",")) {
    fail(where, "decimalSeparator", "different from the delimiter", options.decimalSeparator);
  }
}

/**
 * Refuse the options only a streaming parser can honour, for the parsers that read the whole
 * input at once: `chunk` hands rows over while parsing continues, which a call that returns
 * every row at the end has no moment for. They were accepted and never called.
 */
export function checkNoStreamingOptions(where: string, options: CsvParseOptions): void {
  for (const option of ["chunk", "chunkSize"] as const) {
    if (options[option] !== undefined) {
      throw new CsvArgumentError(
        `options.${option}`,
        `${where}: ${option} is only supported by Csv.parseStream and Csv.ParserStream`
      );
    }
  }
}

/** Check {@link CsvFormatOptions}. */
export function checkFormatOptions(options: CsvFormatOptions, where = "Csv.format"): void {
  check(formatOptionsSchema, options, optionsAt(where));
  const delimiter = options.delimiter ?? ",";
  checkDistinct(where, delimiter, options.quote ?? '"', options.escape);
  checkSeparators(where, delimiter, options.lineEnding);
  const quote = options.quote === undefined ? '"' : options.quote;
  if (
    typeof quote === "string" &&
    quote !== "" &&
    options.lineEnding !== undefined &&
    options.lineEnding.includes(quote)
  ) {
    fail(
      where,
      "lineEnding",
      `a string without the quote character ${describeValue(quote)}`,
      options.lineEnding
    );
  }
}

const rowsSchema = s.custom<readonly unknown[]>(Array.isArray, "an array of rows");

/** The data handed to `Csv.format`: an array of rows. */
export function checkFormatRows(data: unknown, where = "Csv.format"): void {
  check(rowsSchema, data, {
    where,
    field: "data",
    error: argumentError
  });
}

/**
 * One row for a formatter: an array of values, an object, or a `[key, value]` list. A string
 * row was split into one field per character, and a primitive or `null` lost the row.
 */
export function checkFormatRow(
  row: unknown,
  index: number,
  from?: "typeTransform.row",
  where = "Csv.format"
): void {
  if (typeof row !== "object" || row === null) {
    // A row the caller passed is `data[i]`; one `typeTransform.row` returned is a wrong return
    // value of that callback, reported at the option as every other callback result is.
    throw from === undefined
      ? new CsvArgumentError(
          `data[${index}]`,
          `${where}: data[${index}] must be an array or an object, got ${describeValue(row)}`
        )
      : new CsvArgumentError(
          `options.${from}`,
          `${where}: ${from} must return an array or an object for data[${index}], got ${describeValue(row)}`
        );
  }
}

const textInputSchema = s.string();

/** The text a CSV function reads: a `string`, refused with a {@link CsvArgumentError} otherwise. */
export function checkTextInput(where: string, input: unknown): void {
  check(textInputSchema, input, { where, field: "input", error: argumentError });
}

const detectDelimiterOptionsSchema = s.object({
  quote: s.string().optional(),
  delimitersToGuess: s
    .readonlyArray(s.string({ nonEmpty: true }), { expected: "an array of delimiters" })
    .optional(),
  comment: s.string().optional()
});
const _detectDelimiterOptionsMatch: SchemaMatches<
  typeof detectDelimiterOptionsSchema,
  CsvDetectDelimiterOptions
> = true;
void _detectDelimiterOptionsMatch;

const detectLinebreakOptionsSchema = s.object({ quote: s.string().optional() });
const _detectLinebreakOptionsMatch: SchemaMatches<
  typeof detectLinebreakOptionsSchema,
  CsvDetectLinebreakOptions
> = true;
void _detectLinebreakOptionsMatch;

/** The options object of `Csv.detectLinebreak`, already known to be a record. */
export function checkDetectLinebreakOptions(options: CsvDetectLinebreakOptions): void {
  check(detectLinebreakOptionsSchema, options, optionsAt("Csv.detectLinebreak"));
}

/** The options object of `Csv.detectDelimiter`, already known to be a record. */
export function checkDetectDelimiterOptions(options: CsvDetectDelimiterOptions): void {
  check(detectDelimiterOptionsSchema, options, optionsAt("Csv.detectDelimiter"));
}

/**
 * The decoder for an `encoding` option, refusing a label `TextDecoder` does not know with a
 * {@link CsvArgumentError} rather than its `RangeError`.
 */
export function createCsvTextDecoder(
  encoding: string | undefined,
  where = "Csv.parse"
): TextDecoder {
  try {
    return new TextDecoder(encoding || "utf-8");
  } catch (cause) {
    throw new CsvArgumentError(
      "options.encoding",
      `${where}: options.encoding must be a label TextDecoder supports, got ${describeValue(encoding)}`,
      { cause }
    );
  }
}
