/**
 * Option checks shared by every CSV parser and formatter.
 *
 * Parsing and formatting each resolve their options in one place (`createParseConfig`,
 * `createFormatConfig`), so the checks live there and cover `Csv.parse`, `Csv.parseAsync`,
 * `Csv.parseRows`, the streams and the worker alike. Each refusal is a {@link CsvArgumentError}
 * naming the option, thrown before any input is read: a formatter given `delimiter: ""`
 * wrote rows with nothing between the fields, and `quote` equal to the delimiter wrote a
 * file no reader can split.
 */

import { CsvArgumentError } from "@csv/errors";
import type { CsvFormatOptions, CsvParseOptions } from "@csv/types";
import { describeValue } from "@utils/describe-value";
export { describeValue } from "@utils/describe-value";

/** The argument path an option or argument name stands for: `delimiter` → `options.delimiter`. */
function fieldOf(option: string): string {
  return option === "options" ||
    option === "data" ||
    option === "input" ||
    option.startsWith("data[") ||
    option.startsWith("options.")
    ? option
    : `options.${option}`;
}

function fail(where: string, option: string, expected: string, value: unknown): never {
  throw new CsvArgumentError(
    fieldOf(option),
    `${where}: ${option} must be ${expected}, got ${describeValue(value)}`
  );
}

function checkString(where: string, option: string, value: unknown, allowEmpty: boolean): void {
  if (value === undefined) {
    return;
  }
  if (typeof value !== "string" || (!allowEmpty && value === "")) {
    fail(where, option, allowEmpty ? "a string" : "a non-empty string", value);
  }
}

/** A quote or escape option: one character, or `false`/`null` to disable. */
function checkCharOption(where: string, option: string, value: unknown, allowEmpty: boolean): void {
  if (value === undefined || value === false || value === null) {
    return;
  }
  if (typeof value !== "string" || value.length > 1 || (!allowEmpty && value === "")) {
    fail(where, option, "a single character, or false to disable", value);
  }
  if (value === "\r" || value === "\n") {
    fail(where, option, "a character other than a line break", value);
  }
}

function checkCount(where: string, option: string, value: unknown, min: number): void {
  if (
    value !== undefined &&
    (typeof value !== "number" || !Number.isInteger(value) || value < min)
  ) {
    fail(where, option, `an integer >= ${min}`, value);
  }
}

function checkFunction(where: string, option: string, value: unknown): void {
  if (value !== undefined && typeof value !== "function") {
    fail(where, option, "a function", value);
  }
}

function checkEnum(
  where: string,
  option: string,
  value: unknown,
  allowed: readonly unknown[]
): void {
  if (value !== undefined && !allowed.includes(value)) {
    fail(where, option, `one of ${allowed.map(v => describeValue(v)).join(", ")}`, value);
  }
}

export function checkObject(where: string, value: unknown): void {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    fail(where, "options", "an options object", value);
  }
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
 * A delimiter holding a line break splits records instead of fields, and a line ending that
 * overlaps the delimiter splits fields instead of records: neither can be read back.
 */
function checkSeparators(where: string, delimiter: string, lineEnding: string | undefined): void {
  if (delimiter.includes("\n") || delimiter.includes("\r")) {
    fail(where, "delimiter", "a string without line breaks", delimiter);
  }
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

function checkBoolean(where: string, option: string, value: unknown): void {
  if (value !== undefined && typeof value !== "boolean") {
    fail(where, option, "true or false", value);
  }
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * A `dynamicTyping` setting: a boolean, or a map whose own entries are booleans or converter
 * functions. Only own entries are ever looked up, so a column named `toString` is not read
 * through `Object.prototype`.
 *
 * `field` is where the setting sits in the public signature: `options.dynamicTyping` for a
 * parse option, `dynamicTyping` where it is a parameter of its own
 * (`Csv.applyDynamicTypingToRow`).
 */
export function checkDynamicTyping(
  where: string,
  dynamicTyping: unknown,
  field = "options.dynamicTyping"
): void {
  if (dynamicTyping === undefined || typeof dynamicTyping === "boolean") {
    return;
  }
  if (!isPlainRecord(dynamicTyping)) {
    throw new CsvArgumentError(
      field,
      `${where}: dynamicTyping must be true, false or a map of column name to boolean or function, got ${describeValue(dynamicTyping)}`
    );
  }
  for (const key of Object.keys(dynamicTyping)) {
    const v = dynamicTyping[key];
    if (typeof v !== "boolean" && typeof v !== "function") {
      const entry = `[${JSON.stringify(key)}]`;
      throw new CsvArgumentError(
        `${field}${entry}`,
        `${where}: dynamicTyping${entry} must be true, false or a function, got ${describeValue(v)}`
      );
    }
  }
}

/** Check {@link CsvParseOptions}. `delimiter: ""` and `lineEnding: ""` mean auto-detect. */
export function checkParseOptions(options: CsvParseOptions, where = "Csv.parse"): void {
  checkObject(where, options);
  checkString(where, "delimiter", options.delimiter, true);
  // An empty quote or escape has always meant that no character matches, so it is kept.
  checkCharOption(where, "quote", options.quote, true);
  checkCharOption(where, "escape", options.escape, true);
  checkDistinct(where, options.delimiter ?? ",", options.quote ?? '"', options.escape);
  if (options.delimitersToGuess !== undefined) {
    if (!Array.isArray(options.delimitersToGuess)) {
      fail(where, "delimitersToGuess", "an array of delimiters", options.delimitersToGuess);
    }
    options.delimitersToGuess.forEach((d, i) => {
      checkString(where, `delimitersToGuess[${i}]`, d, false);
      checkSeparators(where, d, undefined);
    });
  }
  checkString(where, "lineEnding", options.lineEnding, true);
  checkSeparators(where, options.delimiter ?? ",", options.lineEnding);
  checkString(where, "comment", options.comment, true);
  checkEnum(where, "skipEmptyLines", options.skipEmptyLines, [true, false, "greedy"]);
  checkCount(where, "maxRows", options.maxRows, 0);
  checkCount(where, "toLine", options.toLine, 0);
  checkCount(where, "skipLines", options.skipLines, 0);
  checkCount(where, "skipRows", options.skipRows, 0);
  checkCount(where, "maxRowBytes", options.maxRowBytes, 1);
  checkCount(where, "chunkSize", options.chunkSize, 1);
  if (options.columnMismatch !== undefined) {
    if (typeof options.columnMismatch !== "object" || options.columnMismatch === null) {
      fail(where, "columnMismatch", "{ less, more }", options.columnMismatch);
    }
    checkEnum(where, "columnMismatch.less", options.columnMismatch.less, ["error", "pad"]);
    checkEnum(where, "columnMismatch.more", options.columnMismatch.more, [
      "error",
      "truncate",
      "keep"
    ]);
  }
  checkString(where, "encoding", options.encoding, false);
  for (const option of [
    "rowTransform",
    "validate",
    "chunk",
    "beforeFirstChunk",
    "onSkip"
  ] as const) {
    checkFunction(where, option, options[option]);
  }
  for (const option of [
    "trim",
    "ltrim",
    "rtrim",
    "groupColumnsByName",
    "fastMode",
    "info",
    "raw",
    "relaxQuotes",
    "skipRecordsWithError",
    "skipRecordsWithEmptyValues",
    "objectMode"
  ] as const) {
    checkBoolean(where, option, options[option]);
  }
  const { headers } = options;
  if (
    headers !== undefined &&
    typeof headers !== "boolean" &&
    typeof headers !== "function" &&
    !Array.isArray(headers)
  ) {
    fail(where, "headers", "true, false, an array of header names or a function", headers);
  }
  if (Array.isArray(headers)) {
    headers.forEach((h, i) => {
      if (h !== null && h !== undefined && typeof h !== "string") {
        fail(where, `headers[${i}]`, "a string, null or undefined", h);
      }
    });
  }
  checkDynamicTyping(where, options.dynamicTyping);
  const { castDate } = options;
  if (castDate !== undefined && typeof castDate !== "boolean") {
    if (!Array.isArray(castDate) || castDate.some(c => typeof c !== "string")) {
      fail(where, "castDate", "true, false or an array of column names", castDate);
    }
  }
  if (options.objname !== undefined) {
    checkString(where, "objname", options.objname, false);
    if (headers === undefined || headers === false) {
      fail(where, "objname", "used with headers (it names a header column)", options.objname);
    }
  }
  checkEnum(where, "decimalSeparator", options.decimalSeparator, [".", ","]);
  if (
    options.decimalSeparator === "," &&
    (options.delimiter ?? ",") !== "" &&
    (options.delimiter ?? ",").includes(",")
  ) {
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
        `${where}: ${option} is only supported by Csv.parseRows and Csv.ParserStream`
      );
    }
  }
}

const TYPE_TRANSFORM_KEYS: readonly string[] = [
  "boolean",
  "date",
  "number",
  "bigint",
  "object",
  "string",
  "row"
];

function checkQuoteConfig(where: string, option: string, value: unknown): void {
  if (value === undefined || typeof value === "boolean") {
    return;
  }
  if (Array.isArray(value)) {
    value.forEach((v, i) => checkBoolean(where, `${option}[${i}]`, v));
    return;
  }
  if (!isPlainRecord(value)) {
    fail(where, option, "true, false, an array of booleans or a map of header to boolean", value);
  }
  for (const key of Object.keys(value)) {
    checkBoolean(where, `${option}[${JSON.stringify(key)}]`, value[key]);
  }
}

/** Check {@link CsvFormatOptions}. */
export function checkFormatOptions(options: CsvFormatOptions, where = "Csv.format"): void {
  checkObject(where, options);
  checkString(where, "delimiter", options.delimiter, false);
  checkCharOption(where, "quote", options.quote, false);
  // An empty escape has always fallen back to the quote character.
  checkCharOption(where, "escape", options.escape, true);
  const delimiter = options.delimiter ?? ",";
  checkDistinct(where, delimiter, options.quote ?? '"', options.escape);
  checkString(where, "lineEnding", options.lineEnding, false);
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
  for (const option of [
    "objectMode",
    "readableObjectMode",
    "writeHeaders",
    "bom",
    "trailingNewline",
    "escapeFormulae"
  ] as const) {
    checkBoolean(where, option, options[option]);
  }
  const { headers, columns, typeTransform } = options;
  if (headers !== undefined && headers !== null && typeof headers !== "boolean") {
    if (!Array.isArray(headers)) {
      fail(where, "headers", "true, false or an array of header names", headers);
    }
    headers.forEach((h, i) => {
      if (typeof h !== "string") {
        fail(where, `headers[${i}]`, "a string", h);
      }
    });
  }
  if (columns !== undefined) {
    if (!Array.isArray(columns)) {
      fail(where, "columns", "an array of keys or { key, header? } objects", columns);
    }
    columns.forEach((c, i) => {
      if (typeof c === "string") {
        return;
      }
      if (!isPlainRecord(c) || typeof c.key !== "string") {
        fail(where, `columns[${i}]`, "a key or { key, header? }", c);
      }
      if (c.header !== undefined && typeof c.header !== "string") {
        fail(where, `columns[${i}].header`, "a string", c.header);
      }
    });
  }
  checkQuoteConfig(where, "quoteColumns", options.quoteColumns);
  checkQuoteConfig(where, "quoteHeaders", options.quoteHeaders);
  if (typeTransform !== undefined) {
    if (!isPlainRecord(typeTransform)) {
      fail(where, "typeTransform", "an object of transform functions", typeTransform);
    }
    for (const key of Object.keys(typeTransform)) {
      if (!TYPE_TRANSFORM_KEYS.includes(key)) {
        fail(
          where,
          `typeTransform[${JSON.stringify(key)}]`,
          `a key in ${TYPE_TRANSFORM_KEYS.join(", ")}`,
          key
        );
      }
      checkFunction(where, `typeTransform.${key}`, typeTransform[key]);
    }
  }
}

/** The data handed to `Csv.format`: an array of rows. */
export function checkFormatRows(data: unknown, where = "Csv.format"): void {
  if (!Array.isArray(data)) {
    fail(where, "data", "an array of rows", data);
  }
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
    fail(
      where,
      from === undefined ? `data[${index}]` : `${from} result for data[${index}]`,
      "an array or an object",
      row
    );
  }
}

/** The text a CSV function reads: a `string`, refused with a {@link CsvArgumentError} otherwise. */
export function checkTextInput(where: string, input: unknown): void {
  if (typeof input !== "string") {
    throw new CsvArgumentError(
      "input",
      `${where}: input must be a string, got ${describeValue(input)}`
    );
  }
}

/** The options object of `Csv.detectDelimiter`. */
export function checkDetectDelimiterOptions(options: Record<string, unknown>): void {
  const where = "Csv.detectDelimiter";
  checkString(where, "quote", options.quote, true);
  checkString(where, "comment", options.comment, true);
  const guesses = options.delimitersToGuess;
  if (guesses !== undefined) {
    if (!Array.isArray(guesses)) {
      fail(where, "delimitersToGuess", "an array of delimiters", guesses);
    }
    guesses.forEach((d, i) => checkString(where, `delimitersToGuess[${i}]`, d, false));
  }
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
      `${where}: encoding must be a label TextDecoder supports, got ${describeValue(encoding)}`,
      { cause }
    );
  }
}
