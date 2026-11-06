/**
 * CSV ↔ Workbook bridge — free functions.
 *
 * These functions implement the CSV import/export capability as
 * tree-shakeable free functions that take a `Workbook` handle, instead of
 * methods on the `Workbook` class. A consumer who never imports this module
 * pays nothing for the CSV parser/formatter and the `@stream` pipeline they
 * pull in — the core `Workbook` no longer references `@csv` at all.
 *
 * Layer note: this file lives inside the excel module (layer 4), so it may
 * import from `@csv` (layer 2). The csv module never imports excel, so the
 * `readCsv(workbook, ...)` direction (which constructs worksheets) must live
 * here, not in the csv module.
 *
 * Node-only file-path variants (`readCsvFile` / `writeCsvFile`) live in
 * `./csv-bridge.node.ts`.
 */

import { CsvArgumentError } from "@csv/errors";
import { formatCsv, headerNames } from "@csv/format";
import { createFormatConfig } from "@csv/format/config";
import { parseCsv } from "@csv/parse";
import { recordErrorToThrow } from "@csv/parse/sync";
import { CsvParserStream, CsvFormatterStream } from "@csv/stream";
import type { CsvParseOptions, CsvFormatOptions, CsvRecordError, Row as CsvRow } from "@csv/types";
import { checkFormatOptions, checkParseOptions } from "@csv/utils/check-options";
import {
  DEFAULT_DATE_FORMATS,
  addInputRow,
  createDefaultValueMapper
} from "@excel/bridge/value-reading";
import { rowHasValues, rowValues } from "@excel/core/row";
import type { RowData } from "@excel/core/row";
import { addWorksheet, createWorkbook, getWorksheet } from "@excel/core/workbook";
import { refuseLongSheetName } from "@excel/core/workbook-core";
import { isReadTarget } from "@excel/core/workbook-io-types";
import type { Workbook } from "@excel/core/workbook.browser";
import type { Worksheet } from "@excel/core/worksheet";
import { ExcelArgumentError, ExcelDownloadError } from "@excel/errors";
import type { CellValue, CellValueInput } from "@excel/types";
import { argumentFailure, checkSchema } from "@excel/utils/arguments";
import { finished, pipeline } from "@stream";
import { isReadableStreamLike, readableStreamToAsyncIterable } from "@stream/core/stream-helpers";
import type { IReadable, IWritable } from "@stream/types";
import type { DateFormat } from "@utils/datetime";
import { createDateFormatter, createIsoDateFormatter, isDateFormat } from "@utils/datetime";
import { describeValue as describeInput } from "@utils/describe-value";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

// =============================================================================
// Public CSV option / input types
// =============================================================================

export type ExcelCsvInput =
  | string // CSV string or URL (http:// or https://)
  | ArrayBuffer
  | ArrayBufferView // Bytes: a Uint8Array, or any other typed array or DataView, read as the bytes it views
  | File // Browser File object
  | Blob // Browser Blob object
  | IReadable<Uint8Array | string> // Node readable stream
  | ReadableStream<Uint8Array> // Web stream (a `fetch` body), with or without async iteration
  | AsyncIterable<Uint8Array | string>; // Chunks of bytes or text

type CsvReadParseFields = Pick<
  CsvParseOptions,
  | "delimiter"
  | "quote"
  | "escape"
  | "delimitersToGuess"
  | "lineEnding"
  | "headers"
  | "skipEmptyLines"
  | "trim"
  | "ltrim"
  | "rtrim"
  | "comment"
  | "maxRows"
  | "toLine"
  | "skipLines"
  | "skipRows"
  | "columnMismatch"
  | "groupColumnsByName"
  | "relaxQuotes"
  | "fastMode"
  | "skipRecordsWithError"
  | "skipRecordsWithEmptyValues"
  | "onSkip"
>;

type CsvWriteFormatFields = Pick<
  CsvFormatOptions,
  | "delimiter"
  | "quote"
  | "escape"
  | "lineEnding"
  | "decimalSeparator"
  | "quoteColumns"
  | "quoteHeaders"
  | "escapeFormulae"
>;

/**
 * Options for reading CSV into a workbook — `ExcelCsv.read`, `addWorksheet`, `readFile` and `createImportStream` —
 * the parser's options plus where the rows go and how text becomes cell values.
 */
export interface ExcelCsvReadOptions extends CsvReadParseFields {
  /** The name of the worksheet the rows are written to. */
  sheetName?: string;
  /**
   * The forms a text is tried against to read it as a date. Default: the ISO forms with and without a time and an
   * offset, and `MM-DD-YYYY`.
   *
   * A text that states no timezone — `2024-03-10`, `2024-03-10T14:30:00`, `10/03/2024 14:30:00` — is the calendar
   * value it names, in UTC, which is how a cell holds a `Date`: the cell is `2024-03-10 14:30` on every machine. A
   * text with `Z` or an offset is the instant it names. (A date-only text used to be read as the host's local
   * midnight, so the same CSV made a different cell in every timezone.)
   */
  dateFormats?: readonly DateFormat[];
  /** The decimal separator of the numbers in the text. */
  decimalSeparator?: CsvFormatOptions["decimalSeparator"];
  /**
   * Turn each parsed text into the cell value written, in place of the default number / date / boolean reading.
   * Called with the field's text and its 0-based column; what it returns is checked as `Cell.setValue` checks a value.
   */
  map?(text: string, column: number): CellValueInput;
  /** For a URL: the request's headers. */
  requestHeaders?: Record<string, string>;
  /** For a URL: the request's body. */
  requestBody?: NonNullable<RequestInit["body"]>;
  /** For a URL: send credentials with the request. */
  withCredentials?: boolean;
  /** For a URL: cancels the request. */
  signal?: AbortSignal;
  /**
   * The encoding of the input bytes — any label `TextDecoder` accepts (`"utf-8"`, `"windows-1252"`, `"shift_jis"`, …),
   * as `Csv.parse` takes it. Default `"utf-8"`. Applies to every byte input: an `ArrayBuffer` or any `ArrayBufferView`, a `File`
   * or `Blob`, a byte stream, a URL's response, a file path. Text input is already decoded.
   */
  encoding?: string;
  /** For a `File` or `Blob`: called as it is read, with the bytes read and the total. */
  onProgress?: (loaded: number, total: number) => void;
  /** For a URL or a `File`: parse as the bytes arrive rather than after all of them. */
  stream?: boolean;
  /** For a file path: the read chunk size in bytes, a whole number of at least 1. */
  highWaterMark?: number;
}

/**
 * Options for writing a worksheet as CSV — `ExcelCsv.format`, `toBuffer`, `toBufferSync`, `toStream`, `writeStream`
 * and `writeFile` — the formatter's options plus which sheet is written and how cell values become text.
 */
export interface ExcelCsvWriteOptions extends CsvWriteFormatFields {
  /**
   * Whether to write the sheet's first row — its header row, where `ExcelCsv.read` with `headers` puts a file's
   * header. `false` leaves it out, in every writer: the CSV holds the data rows only, and `ExcelCsv.writeFile` with
   * `append` continues a file that already has its header. Rows are not otherwise affected: the blank line of an empty
   * row 2 is still written (`includeEmptyRows`).
   *
   * The header row is quoted by `quoteHeaders` — `true`, an array by column index, an object by header name — as
   * `Csv.format` quotes one, and `quoteColumns` quotes the rows below it; the header row's texts are the names both
   * objects match, whether the row is written or not. `quoteColumns: true` quotes every field, the header's included.
   * @default true
   */
  writeHeaders?: boolean;
  /** The worksheet written, by name. Default: the first. */
  sheetName?: string;
  /** The worksheet written, by id, when no `sheetName` is given. */
  sheetId?: number;
  /**
   * For a file path (`ExcelCsv.writeFile`): append to the file when it exists and is not empty, rather than replace it.
   *
   * Every row is appended, the sheet's first row included; nothing is guessed from what the file holds. To continue a
   * file that already has the header row, leave it out with `writeHeaders: false`. A line break is written first
   * unless the file already ends with one — `\n`, `\r\n` or `\r`, read in `encoding`, so a UTF-16LE file is read as
   * UTF-16LE — or holds nothing but a byte-order mark. A file that does not exist, or is empty, is simply written.
   * @default false
   */
  append?: boolean;
  /**
   * The form dates are written in, as `YYYY-MM-DD HH:mm:ss` tokens. Default: ISO 8601 with milliseconds, in UTC
   * (`…Z`). A cell's number format is not applied — `Cell.getDisplayText` is the reading that applies it.
   *
   * See {@link ExcelCsvWriteOptions.dateUTC} for which clock the fields are taken from.
   */
  dateFormat?: string;
  /**
   * Take a date's fields from UTC — the calendar value the cell holds, and what Excel shows — rather than from the
   * host's timezone. Default `true`; `false` writes the host's local time.
   *
   * A date cell is wall-clock time with no timezone of its own. For the cell `2024-03-10 00:00`, written in New York
   * (UTC−5):
   *
   * | options                                        | written                         |
   * | ---------------------------------------------- | ------------------------------- |
   * | (none)                                         | `2024-03-10T00:00:00.000Z`      |
   * | `{ dateFormat: "YYYY-MM-DD" }`                 | `2024-03-10`                    |
   * | `{ dateUTC: false }`                           | `2024-03-09T19:00:00.000-05:00` |
   * | `{ dateFormat: "YYYY-MM-DD", dateUTC: false }` | `2024-03-09`                    |
   *
   * The default output is the same on every machine and reads back as the same cell. A read takes a text with no
   * timezone as the calendar value in UTC, so `dateFormat` with `dateUTC: false` reads back shifted by the host's
   * offset. (Before this release the default was `false`.)
   */
  dateUTC?: boolean;
  /** Turn each cell value into the value written, in place of the default formatting. */
  map?(value: CellValue, index: number): CellValue;
  /**
   * Turn each row written into the row to write instead — the values after `map`, and the row's number on the sheet —
   * so a row can gain, lose or reorder fields: exceljs's `formatterOptions.transform`. Every row written passes through
   * it, the header row and the blank line of an empty row (`[]`) included.
   */
  mapRow?(values: CellValue[], rowNumber: number): readonly CellValue[];
  /** Write a blank line for each empty row between rows that have values. Default `true`. */
  includeEmptyRows?: boolean;
  /**
   * For a file path (`ExcelCsv.writeFile`): the file's encoding — a Node `Buffer` encoding (`"utf8"`, `"utf16le"`,
   * `"latin1"`, …), where the WHATWG spellings `"utf-8"` and `"utf-16le"` are the same encodings. Default `"utf8"`.
   * The other writers produce UTF-8 (or text) and refuse any other encoding rather than ignore it.
   */
  encoding?: string;
  /** For a file path (`ExcelCsv.writeFile`): the write chunk size in bytes, a whole number of at least 1. */
  highWaterMark?: number;
}

/**
 * Both {@link ExcelCsvReadOptions} and {@link ExcelCsvWriteOptions}, for one bag shared by a read and a write. Each function
 * takes the half it uses.
 */
export interface ExcelCsvOptions extends ExcelCsvReadOptions, ExcelCsvWriteOptions {
  /**
   * Turn each value into another — a read's text into the cell value written, a write's cell value into the value
   * written. One function for both halves, so it takes a cell value; see {@link ExcelCsvReadOptions.map} and
   * {@link ExcelCsvWriteOptions.map} for what each calls it with.
   */
  map?(value: CellValue, index: number): CellValue;
  /** @deprecated Has no effect on a workbook: a record's position has nowhere to go in a sheet. Removed in 1.0. */
  info?: CsvParseOptions["info"];
  /** @deprecated Has no effect on a workbook: a record's raw text has nowhere to go in a sheet. Removed in 1.0. */
  raw?: CsvParseOptions["raw"];
}

// =============================================================================
// Value mappers
// =============================================================================

/**
 * Map a cell value to what the CSV should carry.
 *
 * **A date is written from its UTC fields unless `dateUTC: false`.** A date cell is wall-clock time: its serial is
 * computed from the `Date`'s UTC fields (`dateToExcel`), so those fields are what Excel shows, and what its own
 * "Save as CSV" writes. The default used to be the host's local time — paired with a reader that built local time,
 * so a file round-tripped, but in New York a cell Excel shows as `2024-03-10` was written `2024-03-09`. The reader
 * now reads a text with no timezone as UTC, so the writer does too. See {@link ExcelCsvWriteOptions.dateUTC}.
 *
 * `2020-01-15T00:00:00.000Z` was once read a day early, because the fixed-width parsers did not check their input's
 * width and the millisecond form was missing from {@link DEFAULT_DATE_FORMATS}. Both are fixed at the root, in
 * `@utils/datetime` and above.
 */
function createDefaultWriteMapper(
  dateFormat: string | undefined,
  dateUTC: boolean
): (value: CellValue) => CellValue {
  const formatter = dateFormat
    ? createDateFormatter(dateFormat, { utc: dateUTC })
    : createIsoDateFormatter(dateUTC);

  return function mapValue(value: CellValue): CellValue {
    if (value === null || value === undefined) {
      return value;
    }
    if (value instanceof Date) {
      return formatter.format(value);
    }
    if (typeof value !== "object") {
      return value;
    }

    const maybeLink = value as { hyperlink?: unknown; text?: unknown };
    if (typeof maybeLink.hyperlink === "string" || typeof maybeLink.text === "string") {
      const url = typeof maybeLink.hyperlink === "string" ? maybeLink.hyperlink : "";
      const text = typeof maybeLink.text === "string" ? maybeLink.text : "";
      return url || text || "";
    }
    if ("formula" in value || "sharedFormula" in value) {
      return (value as { result?: CellValue }).result ?? "";
    }
    const richTextValue = value as { richText?: { text: string }[] };
    if ("richText" in value && Array.isArray(richTextValue.richText)) {
      return richTextValue.richText.map(r => r.text).join("");
    }
    const checkboxValue = value as { checkbox?: unknown };
    if ("checkbox" in value && typeof checkboxValue.checkbox === "boolean") {
      return checkboxValue.checkbox;
    }
    const errorValue = value as { error?: unknown };
    if ("error" in value && typeof errorValue.error === "string") {
      return errorValue.error;
    }
    return JSON.stringify(value);
  };
}

// =============================================================================
// Input detection
// =============================================================================

function isArrayBuffer(input: unknown): input is ArrayBuffer {
  return Object.prototype.toString.call(input) === "[object ArrayBuffer]";
}

function isUrl(input: unknown): input is string {
  return typeof input === "string" && /^https?:\/\//i.test(input);
}

/** A `Blob`, or a `File` (which is one). */
function isBlob(input: unknown): input is Blob {
  return typeof Blob !== "undefined" && input instanceof Blob;
}

function isReadableStream(input: unknown): input is IReadable<Uint8Array | string> {
  if (!input || typeof input !== "object") {
    return false;
  }
  const obj = input as {
    [Symbol.asyncIterator]?: unknown;
    pipe?: unknown;
    on?: unknown;
  };
  return (
    typeof obj[Symbol.asyncIterator] === "function" ||
    (typeof obj.pipe === "function" && typeof obj.on === "function")
  );
}

// =============================================================================
// Stream helpers
// =============================================================================

function* iterateWorksheetRows(
  worksheet: Worksheet
): Generator<{ row: RowData; rowNumber: number }> {
  const rows = worksheet._rows;
  if (!rows || rows.length === 0) {
    return;
  }
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (row && rowHasValues(row)) {
      yield { row, rowNumber: i + 1 };
    }
  }
}

/**
 * The rows a worksheet is written as: each row's values, mapped, with a blank row for each empty row between two
 * written ones (`includeEmptyRows`), and without the sheet's first row — its header row — when `writeHeaders` is
 * `false`. One sequence for every writer, so `format`, the streams and the file agree on what a sheet's CSV is.
 */
function* csvRows(
  worksheet: Worksheet,
  options: ExcelCsvWriteOptions | undefined,
  map: (value: CellValue, index: number) => CellValue
): Generator<CellValue[]> {
  const includeEmptyRows = options?.includeEmptyRows !== false;
  const writeHeaders = options?.writeHeaders !== false;
  const mapRow = options?.mapRow;
  const out = (values: CellValue[], rowNumber: number): CellValue[] => {
    if (!mapRow) {
      return values;
    }
    const mapped: unknown = mapRow(values, rowNumber);
    if (!Array.isArray(mapped)) {
      throw new ExcelArgumentError(
        "options.mapRow",
        `options.mapRow must return an array of values for row ${rowNumber}, got ${describeInput(mapped)}.`
      );
    }
    return [...(mapped as CellValue[])];
  };
  let lastRow = 1;
  for (const { row, rowNumber } of iterateWorksheetRows(worksheet)) {
    if (!writeHeaders && rowNumber === 1) {
      continue;
    }
    if (includeEmptyRows) {
      while (lastRow++ < rowNumber - 1) {
        yield out([], lastRow);
      }
    }
    yield out(rowValues(row).slice(1).map(map), rowNumber);
    lastRow = rowNumber;
  }
}

// Structural shape of the CSV formatter's event API. The `any[]` is required
// for interop: `CsvFormatterStream extends Transform`, whose `once`/`off` carry
// Node's overloaded EventEmitter signatures; only `any[]` is assignable to that
// overload set here. `unknown[]`/`never[]` would reject the formatter.
function createDrainRacer(emitter: {
  once(event: string, listener: (...args: any[]) => void): void;
  off(event: string, listener: (...args: any[]) => void): void;
}): () => Promise<void> {
  return () =>
    new Promise<void>((resolve, reject) => {
      const cleanup = (): void => {
        emitter.off("drain", onDrain);
        emitter.off("error", onError);
        emitter.off("close", onClose);
      };
      const onDrain = (): void => {
        cleanup();
        resolve();
      };
      const onError = (err: Error): void => {
        cleanup();
        reject(err);
      };
      const onClose = (): void => {
        cleanup();
        reject(new Error("stream closed before drain"));
      };
      emitter.once("drain", onDrain);
      emitter.once("error", onError);
      emitter.once("close", onClose);
    });
}

/**
 * The parser options for a read, and the records the parser leaves out (a row with the wrong number of fields under a
 * header, text after a closing quote) when the caller gives no `onSkip` to receive them. A sheet has nowhere to put
 * such a record, so a read throws for it rather than import the rest without it — as `Csv.parseStream` does. With a
 * header row they used to be dropped without a word: `Csv.parse` reports them in its result there, and the result was
 * read for its rows only.
 */
function buildParserOptions(options?: ExcelCsvReadOptions): {
  parserOptions: Partial<CsvParseOptions>;
  skipped: CsvRecordError[];
} {
  const skipped: CsvRecordError[] = [];
  const onSkip = options?.onSkip ?? ((error: CsvRecordError) => void skipped.push(error));
  return { parserOptions: parserOptionsOf(options, onSkip), skipped };
}

function parserOptionsOf(
  options: ExcelCsvReadOptions | undefined,
  onSkip: CsvParseOptions["onSkip"]
): Partial<CsvParseOptions> {
  return {
    delimiter: options?.delimiter ?? ",",
    quote: options?.quote,
    escape: options?.escape,
    delimitersToGuess: options?.delimitersToGuess,
    lineEnding: options?.lineEnding,
    headers: options?.headers,
    skipEmptyLines: options?.skipEmptyLines,
    trim: options?.trim,
    ltrim: options?.ltrim,
    rtrim: options?.rtrim,
    comment: options?.comment,
    maxRows: options?.maxRows,
    toLine: options?.toLine,
    skipLines: options?.skipLines,
    skipRows: options?.skipRows,
    columnMismatch: options?.columnMismatch,
    groupColumnsByName: options?.groupColumnsByName,
    relaxQuotes: options?.relaxQuotes,
    fastMode: options?.fastMode,
    // `info` and `raw` change the shape of what the parser returns, which a sheet has no place for: `info` made the
    // import fail on the records it could not read as rows.
    skipRecordsWithError: options?.skipRecordsWithError,
    skipRecordsWithEmptyValues: options?.skipRecordsWithEmptyValues,
    onSkip,
    // Byte chunks are decoded by the parser stream; text arrives already decoded.
    encoding: options?.encoding
  };
}

/** Whether a quoting option names columns — its object form, which matches the header row's texts. */
function hasQuoteNames(options: ExcelCsvWriteOptions | undefined): boolean {
  const named = (value: unknown): boolean =>
    typeof value === "object" && value !== null && !Array.isArray(value);
  return named(options?.quoteColumns) || named(options?.quoteHeaders);
}

/** The header row's texts, as the formatter writes them unquoted — the names its columns go by. */
function headerTexts(values: CellValue[], options: ExcelCsvWriteOptions | undefined): string[] {
  return headerNames(
    values,
    createFormatConfig({ decimalSeparator: options?.decimalSeparator ?? "." }, "ExcelCsv.format")
  );
}

function buildFormatterOptions(
  options?: ExcelCsvWriteOptions,
  worksheet?: Worksheet,
  map?: (value: CellValue, index: number) => CellValue
) {
  const firstRow = worksheet?._rows[0];
  const sheetHasHeaderRow = firstRow !== undefined && rowHasValues(firstRow);
  const leftOut = options?.writeHeaders === false && sheetHasHeaderRow;
  return {
    delimiter: options?.delimiter ?? ",",
    quote: options?.quote,
    escape: options?.escape,
    lineEnding: options?.lineEnding,
    quoteColumns: options?.quoteColumns,
    quoteHeaders: options?.quoteHeaders,
    decimalSeparator: options?.decimalSeparator ?? ".",
    escapeFormulae: options?.escapeFormulae ?? true,
    // The sheet's first row is its header row: `quoteHeaders` quotes it, and its texts name the columns for both
    // quoting options' object forms, as `Csv.format`'s headers do. Without it the row is left out (`csvRows`).
    firstRowIsHeader: options?.writeHeaders !== false && sheetHasHeaderRow,
    // Left out (`writeHeaders: false`), the header row still names the columns: given as `headers`, which
    // `writeHeaders: false` keeps out of the text, as `Csv.format` does.
    ...(leftOut && map !== undefined && hasQuoteNames(options)
      ? {
          headers: headerTexts(rowValues(firstRow).slice(1).map(map), options),
          writeHeaders: false
        }
      : {})
  };
}

// =============================================================================
// Read (cross-platform)
// =============================================================================

/** The bridge's argument error: `where` is the function called (`"ExcelCsv.read"`), `field` the argument at fault. */
function fail(where: string, field: string, expected: string, value: unknown): never {
  throw argumentFailure(where, field, expected, value);
}

/**
 * The CSV module's own check of the parser or formatter options, raised as this bridge's error. Its
 * `CsvArgumentError` named `CSV parse`, which the caller never called, and was missed by `isExcelError`.
 */
function asBridgeError(where: string, check: () => void): void {
  try {
    check();
  } catch (error) {
    if (error instanceof CsvArgumentError) {
      const detail = error.message.slice(error.message.indexOf(": ") + 2);
      throw new ExcelArgumentError(error.field, `${where}: ${detail}`, {
        cause: error
      });
    }
    throw error;
  }
}

/**
 * The worksheet a write names: `sheetName`, else `sheetId`, else the first. `undefined` only for a workbook with no
 * sheets. A name or id the workbook does not have is refused — it used to write an empty CSV, so a misspelt name
 * produced an empty file without a word.
 */
function selectWorksheet(
  where: string,
  workbook: Workbook,
  options: ExcelCsvWriteOptions | undefined
): Worksheet | undefined {
  if (!isReadTarget(workbook)) {
    fail(where, "workbook", "a workbook", workbook);
  }
  checkSchema(where, "options", writeOptionsSchema.optional(), options);
  const sheetName = options?.sheetName;
  const sheetId = options?.sheetId;
  const key = sheetName || sheetId;
  const worksheet = getWorksheet(workbook, key);
  if (worksheet === undefined && key !== undefined) {
    throw new ExcelArgumentError(
      sheetName ? "options.sheetName" : "options.sheetId",
      `${where}: the workbook has no worksheet ${JSON.stringify(key)}.`
    );
  }
  return worksheet;
}

// The bridge's own options. Those it hands to the CSV parser or formatter are checked there (`asBridgeError`).
const flag = () => s.boolean().optional();
// The file and stream settings both directions share: a chunk size is a whole number of bytes, at least 1.
const highWaterMark = () => s.number({ integer: true, min: 1 }).optional();

const readOptionsSchema = s.object({
  sheetName: s.string().optional(),
  // An unknown format was dropped in silence, so a column of dates read as text; checked below with the default.
  dateFormats: s.unknown<readonly DateFormat[]>().optional(),
  decimalSeparator: s.literal(".", ",").optional(),
  map: s.function<(text: string, column: number) => CellValueInput>().optional(),
  requestHeaders: s.record(s.string()).optional(),
  requestBody: s.unknown<NonNullable<RequestInit["body"]>>().optional(),
  withCredentials: flag(),
  signal: s
    .custom<AbortSignal>(
      v => typeof v === "object" && v !== null && typeof (v as AbortSignal).aborted === "boolean",
      "an AbortSignal"
    )
    .optional(),
  encoding: s.string({ expected: 'a TextDecoder label such as "utf-8"' }).optional(),
  onProgress: s.function<(loaded: number, total: number) => void>().optional(),
  stream: flag(),
  highWaterMark: highWaterMark()
});
const _readOptionsMatch: SchemaMatches<
  typeof readOptionsSchema,
  Omit<ExcelCsvReadOptions, keyof CsvReadParseFields>
> = true;
void _readOptionsMatch;

const writeOptionsSchema = s.object({
  writeHeaders: flag(),
  sheetName: s.string().optional(),
  sheetId: s.number({ integer: true }).optional(),
  append: flag(),
  // A format that is not a string wrote every date cell empty.
  dateFormat: s.string({ expected: 'a format string such as "YYYY-MM-DD"' }).optional(),
  dateUTC: flag(),
  map: s.function<(value: CellValue, index: number) => CellValue>().optional(),
  mapRow: s.function<(values: CellValue[], rowNumber: number) => readonly CellValue[]>().optional(),
  includeEmptyRows: flag(),
  encoding: s.string().optional(),
  highWaterMark: highWaterMark()
});
const _writeOptionsMatch: SchemaMatches<
  typeof writeOptionsSchema,
  Omit<ExcelCsvWriteOptions, keyof CsvWriteFormatFields>
> = true;
void _writeOptionsMatch;

/**
 * Check a read's options, and say how it turns text into cell values: `map`, or the default reading with `dateFormats`
 * checked. Everything a read is given is checked here, before anything is read or added.
 */
function readValueMapper(
  where: string,
  options: ExcelCsvReadOptions | undefined
): (text: string, column: number) => CellValueInput {
  checkSchema(where, "options", readOptionsSchema.optional(), options);
  if (options?.encoding !== undefined) {
    decoderFor(where, options.encoding);
  }
  const dateFormats = options?.dateFormats ?? DEFAULT_DATE_FORMATS;
  // An unknown format was dropped in silence, so a column of dates read as text.
  if (!Array.isArray(dateFormats) || !dateFormats.every(isDateFormat)) {
    throw new ExcelArgumentError(
      "options.dateFormats",
      `${where}: options.dateFormats must be an array of the date forms a CSV read knows, such as "YYYY-MM-DD", got ${JSON.stringify(dateFormats)}.`
    );
  }
  asBridgeError(where, () =>
    checkParseOptions(parserOptionsOf(options, options?.onSkip) as CsvParseOptions)
  );
  return (
    options?.map ||
    (createDefaultValueMapper(dateFormats, {
      decimalSeparator: options?.decimalSeparator
    }) as (text: string) => CellValue)
  );
}

/**
 * The decoder for a read's `encoding` — a `TextDecoder` label, as `Csv.parse` takes it. An unknown label is refused at
 * the call; the option used to reach only the File path, and every other byte input was read as UTF-8 regardless.
 */
export function decoderFor(where: string, encoding: unknown): TextDecoder {
  if (encoding !== undefined && typeof encoding !== "string") {
    fail(where, "options.encoding", 'a TextDecoder label such as "utf-8"', encoding);
  }
  try {
    return new TextDecoder((encoding as string | undefined) || "utf-8");
  } catch (cause) {
    throw new ExcelArgumentError(
      "options.encoding",
      `${where}: options.encoding must be a label TextDecoder supports, such as "utf-8" or "windows-1252", got ${describeInput(encoding)}.`,
      { cause }
    );
  }
}

/** The UTF-8 spellings a writer that only produces UTF-8 accepts for `encoding`. */
const UTF8_LABELS = new Set(["utf8", "utf-8"]);

/**
 * Check a write's options, and say how it turns cell values into text: `map`, or the default writing with
 * `dateFormat` checked. `fileEncoding` is whether `encoding` names the file's encoding (`writeFile`, which checks it
 * against the platform's); every other writer produces UTF-8 and refuses another encoding rather than ignore it.
 */
function writeValueMapper(
  where: string,
  options: ExcelCsvWriteOptions | undefined,
  fileEncoding = false
): (value: CellValue, index: number) => CellValue {
  checkSchema(where, "options", writeOptionsSchema.optional(), options);
  const encoding = options?.encoding;
  if (encoding !== undefined && !fileEncoding && !UTF8_LABELS.has(encoding.toLowerCase())) {
    fail(
      where,
      "options.encoding",
      `"utf-8" — ${where} produces UTF-8; options.encoding chooses the encoding of ExcelCsv.writeFile only`,
      encoding
    );
  }
  asBridgeError(where, () =>
    checkFormatOptions(buildFormatterOptions(options) as CsvFormatOptions)
  );
  return options?.map || createDefaultWriteMapper(options?.dateFormat, options?.dateUTC ?? true);
}

/** @internal Check what a read is given before it touches a file — for the Node file variant. */
export function checkCsvReadArguments(
  where: string,
  options: ExcelCsvReadOptions | undefined
): void {
  readValueMapper(where, options);
}

/**
 * @internal Check what a write is given before it opens a file — for the Node file variant, which checks the file
 * `encoding` against the platform's own list.
 */
export function checkCsvWriteArguments(
  where: string,
  workbook: Workbook,
  options: ExcelCsvWriteOptions | undefined
): void {
  selectWorksheet(where, workbook, options);
  writeValueMapper(where, options, true);
}

/** @internal — shared by read entry points and the Node file variant. */
export function readCsvContent(
  workbook: Workbook,
  content: string | ArrayBuffer | ArrayBufferView,
  options?: ExcelCsvReadOptions,
  where = "ExcelCsv.addWorksheet"
): Worksheet {
  let str: string;
  if (typeof content === "string") {
    str = content;
  } else if (ArrayBuffer.isView(content) || isArrayBuffer(content)) {
    // Recognised structurally: bytes from another realm fail `instanceof` and were read as their `String()`. Decoded
    // as `options.encoding` says, as `Csv.parse` decodes bytes: this path used to read every byte input as UTF-8.
    str = decoderFor(where, options?.encoding).decode(content);
  } else {
    // `String(content)` turned `42`, `null` or a plain object into a one-cell sheet holding that text.
    fail(where, "input", "CSV text, bytes, a URL, a Blob or a readable stream", content);
  }

  const map = readValueMapper(where, options);
  // Parsed before the sheet is added, so text the parser refuses leaves the workbook as it was.
  const { parserOptions, skipped } = buildParserOptions(options);
  const result = parseCsv(str, parserOptions);
  if (skipped.length > 0) {
    throw recordErrorToThrow(where, skipped);
  }
  refuseLongSheetName(options?.sheetName);
  const worksheet = addWorksheet(workbook, options?.sheetName);

  if (Array.isArray(result)) {
    for (const row of result) {
      addInputRow(worksheet, row.map(map), "CSV");
    }
  } else {
    if (result.headers) {
      addInputRow(worksheet, result.headers, "CSV");
    }
    for (const rowObj of result.rows) {
      const rowArray = result.headers!.map(h => rowObj[h]);
      addInputRow(worksheet, rowArray.map(map), "CSV");
    }
  }

  return worksheet;
}

/** @internal — shared by stream read entry points and the Node file variant. */
export function readCsvStream(
  workbook: Workbook,
  stream: IReadable<Uint8Array | string> | AsyncIterable<Uint8Array | string>,
  options?: ExcelCsvReadOptions,
  where = "ExcelCsv.addWorksheet"
): Promise<Worksheet> {
  const map = readValueMapper(where, options);
  refuseLongSheetName(options?.sheetName);
  const worksheet = addWorksheet(workbook, options?.sheetName);
  const { parserOptions, skipped } = buildParserOptions(options);
  const parser = new CsvParserStream(parserOptions);
  const useHeaders = !!options?.headers;
  let headerRow: string[] | null = null;

  return new Promise((resolve, reject) => {
    if (useHeaders) {
      parser.on("headers", (headers: string[]) => {
        headerRow = headers;
        addInputRow(worksheet, headers, "CSV");
      });
    }

    parser.on("data", (row: unknown) => {
      if (useHeaders && headerRow && row && typeof row === "object" && !Array.isArray(row)) {
        const rowObj = row as Record<string, CellValue>;
        const rowArray: CellValue[] = headerRow.map(h => rowObj[h]);
        addInputRow(worksheet, (rowArray as string[]).map(map) as CellValue[], "CSV");
      } else if (Array.isArray(row)) {
        addInputRow(worksheet, (row as string[]).map(map) as CellValue[], "CSV");
      }
    });

    // `stream` may be a plain async-iterable (e.g. a Web ReadableStream
    // adapted via readableStreamToAsyncIterable); pipeline accepts it as a
    // source at runtime even though its type union lists concrete streams.
    pipeline(stream as IReadable<Uint8Array | string>, parser)
      .then(() =>
        skipped.length > 0 ? reject(recordErrorToThrow(where, skipped)) : resolve(worksheet)
      )
      .catch(reject);
  });
}

async function readCsvUrl(
  where: string,
  workbook: Workbook,
  url: string,
  options?: ExcelCsvReadOptions
): Promise<Worksheet> {
  // Checked before the download, so a wrong option does not cost a request.
  readValueMapper(where, options);
  const fetchOptions: RequestInit = {
    method: options?.requestBody ? "POST" : "GET",
    headers: options?.requestHeaders,
    body: options?.requestBody,
    credentials: options?.withCredentials ? "include" : "same-origin",
    signal: options?.signal
  };

  const response = await fetch(url, fetchOptions);
  if (!response.ok) {
    throw new ExcelDownloadError(url, response.status, response.statusText);
  }

  if (options?.stream && response.body) {
    const readable = readableStreamToAsyncIterable<Uint8Array>(response.body);
    return readCsvStream(workbook, readable, options, where);
  }

  // `response.text()` is always UTF-8; the bytes are decoded as `options.encoding` says.
  return readCsvContent(workbook, new Uint8Array(await response.arrayBuffer()), options, where);
}

/**
 * Read a `File` or `Blob`. Its bytes are decoded as `options.encoding` says; the `FileReader` this used does not exist
 * in Node or Bun, where `File` and `Blob` are globals, so a `File` crashed there with a `ReferenceError` — and the
 * `Blob` path, through `blob.text()`, read every encoding as UTF-8.
 */
async function readCsvBlob(
  where: string,
  workbook: Workbook,
  blob: Blob,
  options?: ExcelCsvReadOptions
): Promise<Worksheet> {
  readValueMapper(where, options);
  const LARGE_FILE_THRESHOLD = 10 * 1024 * 1024;
  if ((options?.stream || blob.size > LARGE_FILE_THRESHOLD) && typeof blob.stream === "function") {
    return readCsvStream(
      workbook,
      withProgress(blob.stream(), blob.size, options?.onProgress),
      options,
      where
    );
  }
  const bytes =
    options?.onProgress && typeof blob.stream === "function"
      ? await collectBytes(withProgress(blob.stream(), blob.size, options.onProgress))
      : new Uint8Array(await blob.arrayBuffer());
  return readCsvContent(workbook, bytes, options, where);
}

/** The chunks of a Blob's stream, reporting `(loaded, total)` to `onProgress` as each arrives. */
async function* withProgress(
  stream: ReadableStream<Uint8Array>,
  total: number,
  onProgress: ((loaded: number, total: number) => void) | undefined
): AsyncGenerator<Uint8Array> {
  let loaded = 0;
  for await (const chunk of readableStreamToAsyncIterable<Uint8Array>(stream)) {
    loaded += chunk.byteLength;
    onProgress?.(loaded, total);
    yield chunk;
  }
}

async function collectBytes(chunks: AsyncIterable<Uint8Array>): Promise<Uint8Array> {
  const parts: Uint8Array[] = [];
  let length = 0;
  for await (const chunk of chunks) {
    parts.push(chunk);
    length += chunk.byteLength;
  }
  const bytes = new Uint8Array(length);
  let offset = 0;
  for (const part of parts) {
    bytes.set(part, offset);
    offset += part.byteLength;
  }
  return bytes;
}

/** {@link readCsv}, with the public name the errors give. */
async function addCsvWorksheet(
  where: string,
  workbook: Workbook,
  input: ExcelCsvInput,
  options?: ExcelCsvReadOptions
): Promise<Worksheet> {
  if (isUrl(input)) {
    return readCsvUrl(where, workbook, input, options);
  }
  if (isBlob(input)) {
    return readCsvBlob(where, workbook, input, options);
  }
  // A Web stream is read through its reader: not every browser's `ReadableStream` is
  // async-iterable, and the reader cancels it if the read stops early.
  if (isReadableStreamLike(input)) {
    return readCsvStream(
      workbook,
      readableStreamToAsyncIterable<Uint8Array | string>(input as ReadableStream<Uint8Array>),
      options,
      where
    );
  }
  if (isReadableStream(input)) {
    return readCsvStream(workbook, input, options, where);
  }
  return readCsvContent(workbook, input as string | ArrayBuffer | ArrayBufferView, options, where);
}

/**
 * Read CSV into a new worksheet on `workbook`. Accepts a CSV string, URL,
 * an `ArrayBuffer` or any `ArrayBufferView`, browser `File`/`Blob`, a Node or Web readable stream, or an
 * async iterable of text or byte chunks ({@link ExcelCsvInput}).
 *
 * @example
 * ```ts
 * import { ExcelCsv } from "documonster/excel/csv";
 * await ExcelCsv.addWorksheet(workbook, "a,b,c\n1,2,3");
 * await ExcelCsv.addWorksheet(workbook, "https://example.com/data.csv");
 * await ExcelCsv.addWorksheet(workbook, input, { delimiter: ";", sheetName: "Data" });
 * ```
 */
export async function readCsv(
  workbook: Workbook,
  input: ExcelCsvInput,
  options?: ExcelCsvReadOptions
): Promise<Worksheet> {
  if (!isReadTarget(workbook)) {
    fail("ExcelCsv.addWorksheet", "workbook", "a workbook", workbook);
  }
  return addCsvWorksheet("ExcelCsv.addWorksheet", workbook, input, options);
}

/**
 * Read CSV and return it as a new workbook holding one worksheet. Accepts what
 * {@link readCsv} accepts.
 *
 * Text is read as numbers, dates (see {@link ExcelCsvReadOptions.dateFormats} — a date with no timezone is the calendar
 * value in UTC, as a cell holds it), `true` / `false` and error literals, unless `options.map` says otherwise.
 *
 * @throws {ExcelArgumentError} for an argument or option this bridge cannot use — the message names the function
 * called, and `field` the option (`options.delimiter`). Text the CSV parser cannot read keeps its `CsvError`.
 *
 * @example
 * ```ts
 * import { ExcelCsv } from "documonster/excel/csv";
 * const workbook = await ExcelCsv.read("a,b,c\n1,2,3", { sheetName: "Data" });
 * ```
 */
export async function readCsvWorkbook(
  input: ExcelCsvInput,
  options?: ExcelCsvReadOptions
): Promise<Workbook>;
/**
 * @deprecated Use `ExcelCsv.addWorksheet(workbook, input, options)`, which this form does — it adds a
 * worksheet to `workbook` and returns the worksheet. `ExcelCsv.read(input)` returns a new workbook.
 * Removed in 1.0.
 */
export async function readCsvWorkbook(
  workbook: Workbook,
  input: ExcelCsvInput,
  options?: ExcelCsvReadOptions
): Promise<Worksheet>;
export async function readCsvWorkbook(
  first: Workbook | ExcelCsvInput,
  second?: ExcelCsvInput | ExcelCsvReadOptions,
  third?: ExcelCsvReadOptions
): Promise<Workbook | Worksheet> {
  if (isReadTarget(first)) {
    return addCsvWorksheet("ExcelCsv.read", first, second as ExcelCsvInput, third);
  }
  const workbook = createWorkbook();
  await addCsvWorksheet(
    "ExcelCsv.read",
    workbook,
    first,
    second as ExcelCsvReadOptions | undefined
  );
  return workbook;
}

// =============================================================================
// Write (cross-platform)
// =============================================================================

/**
 * Format a worksheet as CSV text.
 *
 * A date cell is written as the calendar value it holds — ISO 8601 in UTC (`2024-03-10T00:00:00.000Z`), or
 * `dateFormat` filled from the UTC fields — unless `dateUTC: false`; see {@link ExcelCsvWriteOptions.dateUTC} for the four
 * renderings. A cell's number format is not applied.
 *
 * @throws {ExcelArgumentError} for an argument or option this bridge cannot use — the message names the function
 * called, and `field` the option (`options.delimiter`). Text the CSV parser cannot read keeps its `CsvError`.
 *
 * @example
 * ```ts
 * import { ExcelCsv } from "documonster/excel/csv";
 * const csv = ExcelCsv.format(workbook, { sheetName: "Data" });
 * ```
 */
export function writeCsvString(workbook: Workbook, options?: ExcelCsvWriteOptions): string {
  return formatWorksheet("ExcelCsv.format", workbook, options);
}

/** {@link writeCsvString}, with the public name the errors give. */
function formatWorksheet(
  where: string,
  workbook: Workbook,
  options?: ExcelCsvWriteOptions
): string {
  const worksheet = selectWorksheet(where, workbook, options);
  const map = writeValueMapper(where, options);
  if (!worksheet) {
    return "";
  }

  return formatCsv(
    [...csvRows(worksheet, options, map)] as CsvRow[],
    buildFormatterOptions(options, worksheet, map)
  );
}

/**
 * Write a worksheet as CSV into a writable stream. Resolves once the stream
 * has finished; the stream is ended either way.
 *
 * @example
 * ```ts
 * import { ExcelCsv } from "documonster/excel/csv";
 * await ExcelCsv.writeStream(workbook, outputStream, { sheetId: 1 });
 * ```
 */
export async function writeCsvStream(
  workbook: Workbook,
  stream: IWritable<Uint8Array | string>,
  options?: ExcelCsvWriteOptions
): Promise<void> {
  return streamWorksheet("ExcelCsv.writeStream", workbook, stream, options);
}

/**
 * @internal {@link writeCsvStream}, with the public name the errors give — for the Node file variant too
 * (`fileEncoding`), which has checked its own file `encoding`.
 */
export async function streamWorksheet(
  where: string,
  workbook: Workbook,
  stream: IWritable<Uint8Array | string>,
  options?: ExcelCsvWriteOptions,
  fileEncoding = false
): Promise<void> {
  const worksheet = selectWorksheet(where, workbook, options);
  const map = writeValueMapper(where, options, fileEncoding);
  if (!worksheet) {
    // Settled when the stream is, as the pipeline below is: a file that cannot be opened must not resolve as written.
    stream.end();
    await finished(stream as never);
    return;
  }

  const formatter = new CsvFormatterStream(buildFormatterOptions(options, worksheet, map));
  const pipelinePromise = pipeline(formatter, stream);

  const awaitFormatterDrain = createDrainRacer(formatter);

  const writeAndDrain = async (values: CellValue[]): Promise<void> => {
    if (!formatter.write(values as CsvRow)) {
      await awaitFormatterDrain();
    }
  };

  try {
    for (const values of csvRows(worksheet, options, map)) {
      await writeAndDrain(values);
    }

    formatter.end();
    await pipelinePromise;
  } catch (err) {
    formatter.destroy(err instanceof Error ? err : new Error(String(err)));
    await pipelinePromise.catch(() => {});
    throw err;
  }
}

/**
 * Write a worksheet to CSV. Returns a string, or writes to a provided
 * writable stream.
 *
 * Kept for the deprecated `writeCsv` export; one name with two return shapes is
 * what `ExcelCsv.format` and `ExcelCsv.writeStream` replace.
 */
export function writeCsv(workbook: Workbook, options?: ExcelCsvWriteOptions): string;
export function writeCsv(
  workbook: Workbook,
  stream: IWritable<Uint8Array | string>,
  options?: ExcelCsvWriteOptions
): Promise<void>;
export function writeCsv(
  workbook: Workbook,
  streamOrOptions?: IWritable<Uint8Array | string> | ExcelCsvWriteOptions,
  options?: ExcelCsvWriteOptions
): string | Promise<void> {
  if (streamOrOptions && typeof (streamOrOptions as { write?: unknown }).write === "function") {
    return writeCsvStream(workbook, streamOrOptions as IWritable<Uint8Array | string>, options);
  }
  return writeCsvString(workbook, streamOrOptions as ExcelCsvWriteOptions | undefined);
}

/**
 * Write a worksheet to a CSV buffer (`Uint8Array`) synchronously. Published as
 * `ExcelCsv.toBufferSync`; `ExcelCsv.toBuffer` is the asynchronous form.
 */
export function writeCsvBufferSync(
  workbook: Workbook,
  options?: ExcelCsvWriteOptions
): Uint8Array<ArrayBuffer> {
  return new TextEncoder().encode(formatWorksheet("ExcelCsv.toBufferSync", workbook, options));
}

/**
 * Write a worksheet to a CSV buffer (`Uint8Array`). Published as `ExcelCsv.toBuffer`;
 * asynchronous like every other `toBuffer` in the excel entries, with
 * `ExcelCsv.toBufferSync` as the synchronous form.
 */
export async function writeCsvBuffer(
  workbook: Workbook,
  options?: ExcelCsvWriteOptions
): Promise<Uint8Array<ArrayBuffer>> {
  return new TextEncoder().encode(formatWorksheet("ExcelCsv.toBuffer", workbook, options));
}

// =============================================================================
// Streaming surfaces
// =============================================================================

/**
 * Create a readable stream that outputs the worksheet as CSV.
 *
 * This is a **byte stream**: chunks are `Uint8Array` (a `Buffer` on Node) and
 * `objectMode` is off on the readable side, so it can be piped or handed
 * straight to anything that consumes bytes — `stream.pipeline()` into a file,
 * an HTTP response, an SDK upload body, `Readable.toWeb()`, `Buffer.concat`.
 *
 * It used to emit `string` chunks from an object-mode readable side, which is
 * what a CSV *formatter* produces internally. That leaked an implementation
 * detail into a public API and made this stream indistinguishable by type from
 * `Workbook.toStream()` while behaving completely differently: every byte
 * consumer silently received strings instead. Use `ExcelCsv.format` when a
 * `string` is what you want.
 *
 * The return type is deliberately read-only. The concrete object is a
 * `CsvFormatterStream`, i.e. a `Transform`, but exposing that would put `write()`
 * / `end()` on a stream this function has already wired to a producer: a caller
 * who wrote a row would inject it into the output (or trip `write after end`,
 * depending on where the producer had got to). The Node entry re-narrows the
 * concrete stream to a nominal `stream.Readable`, which adds only read-side
 * members — see {@link createCsvByteStream}.
 */
export function createCsvReadStream(
  workbook: Workbook,
  options?: ExcelCsvWriteOptions
): IReadable<Uint8Array> {
  return createCsvByteStream(workbook, options, "ExcelCsv.toStream");
}

/**
 * The concrete stream behind {@link createCsvReadStream}.
 *
 * Exists so the Node entry can type its return as the platform's nominal
 * `stream.Readable` without an `as` assertion: on the Node build
 * `CsvFormatterStream extends Transform extends Readable`, so returning the
 * concrete type is enough for the compiler to check the narrowing itself.
 * Consumers get the read-only {@link createCsvReadStream} signature instead.
 *
 * @internal
 */
export function createCsvByteStream(
  workbook: Workbook,
  options?: ExcelCsvWriteOptions,
  where = "ExcelCsv.toStream"
): CsvFormatterStream {
  const worksheet = selectWorksheet(where, workbook, options);
  const map = writeValueMapper(where, options);
  // Rows in, bytes out: the writable side stays object-mode so `write(row)`
  // keeps taking arrays, while the readable side encodes the CSV text it pushes.
  const formatter = new CsvFormatterStream({
    ...buildFormatterOptions(options, worksheet, map),
    readableObjectMode: false
  });

  if (!worksheet) {
    setTimeout(() => formatter.end(), 0);
    return formatter;
  }

  const awaitFormatterDrain = createDrainRacer(formatter);

  const writeAndDrain = (values: CellValue[]): Promise<void> | void => {
    if (formatter.write(values as CsvRow)) {
      return;
    }
    return awaitFormatterDrain();
  };

  (async () => {
    try {
      for (const values of csvRows(worksheet, options, map)) {
        if (formatter.destroyed) {
          return;
        }
        const p = writeAndDrain(values);
        if (p) {
          await p;
        }
      }
      formatter.end();
    } catch (err) {
      formatter.destroy(err instanceof Error ? err : new Error(String(err)));
    }
  })();

  return formatter;
}

/**
 * Create a writable stream that imports the CSV written into it into a new worksheet on
 * `workbook`. Data flows CSV → worksheet. Published as `ExcelCsv.createImportStream`; the
 * standalone parser that yields rows instead is `Csv.ParserStream`.
 *
 * @example
 * ```ts
 * import { ExcelCsv } from "documonster/excel/csv";
 * const sink = ExcelCsv.createImportStream(workbook, { sheetName: "Imported" });
 * sink.write("a,b\n1,2\n");
 * sink.end();
 * ```
 */
export function createCsvWriteStream(
  workbook: Workbook,
  options?: ExcelCsvReadOptions
): IWritable<Uint8Array | string> {
  return createCsvImportParser(workbook, options);
}

/**
 * The concrete stream behind {@link createCsvWriteStream} — for the Node entry, which types it as the
 * `stream.Writable` it is (as {@link createCsvByteStream} is for `toStream`), so `readable.pipe(sink)` type-checks.
 *
 * @internal
 */
export function createCsvImportParser(
  workbook: Workbook,
  options?: ExcelCsvReadOptions
): CsvParserStream {
  if (!isReadTarget(workbook)) {
    fail("ExcelCsv.createImportStream", "workbook", "a workbook", workbook);
  }
  const map = readValueMapper("ExcelCsv.createImportStream", options);
  refuseLongSheetName(options?.sheetName);
  const worksheet = addWorksheet(workbook, options?.sheetName);
  // A record the parser leaves out, with no `onSkip` to receive it, fails the import stream as it is met.
  const parser: CsvParserStream = new CsvParserStream(
    parserOptionsOf(
      options,
      options?.onSkip ??
        (error =>
          parser.destroy(recordErrorToThrow("ExcelCsv.createImportStream", [error]) as Error))
    )
  );
  const useHeaders = !!options?.headers;
  let headerRow: string[] | null = null;

  if (useHeaders) {
    parser.on("headers", (headers: string[]) => {
      headerRow = headers;
      addInputRow(worksheet, headers, "CSV");
    });
  }

  parser.on("data", (row: unknown) => {
    if (useHeaders && headerRow && row && typeof row === "object" && !Array.isArray(row)) {
      const rowObj = row as Record<string, CellValue>;
      const rowArray: CellValue[] = headerRow.map(h => rowObj[h]);
      addInputRow(worksheet, (rowArray as string[]).map(map) as CellValue[], "CSV");
    } else if (Array.isArray(row)) {
      addInputRow(worksheet, (row as string[]).map(map) as CellValue[], "CSV");
    }
  });

  return parser;
}
