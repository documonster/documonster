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

import { formatCsv } from "@csv/format";
import { parseCsv } from "@csv/parse";
import { CsvParserStream, CsvFormatterStream } from "@csv/stream";
import type { CsvParseOptions, CsvFormatOptions } from "@csv/types";
import type { DecimalSeparator } from "@csv/utils/number";
import { parseNumberFromCsv } from "@csv/utils/number";
import { rowHasValues, rowValues } from "@excel/core/row";
import { addWorksheet, getWorksheet } from "@excel/core/workbook";
import type { Workbook } from "@excel/core/workbook.browser";
import type { Worksheet } from "@excel/core/worksheet";
import { addRow, eachRow } from "@excel/core/worksheet";
import { ExcelDownloadError } from "@excel/errors";
import { pipeline } from "@stream";
import type { IReadable, IWritable } from "@stream/types";
import { readableStreamToAsyncIterable } from "@stream/utils.base";
import type { DateFormat } from "@utils/datetime";
import { DateParser, DateFormatter } from "@utils/datetime";

import type { CellValue, CellErrorValue } from "../types";

// =============================================================================
// Public CSV option / input types
// =============================================================================

export type CsvInput =
  | string // CSV string or URL (http:// or https://)
  | ArrayBuffer
  | Uint8Array
  | File // Browser File object
  | Blob // Browser Blob object
  | IReadable<any>; // Readable stream

type CsvOptionsParseFields = Pick<
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
  | "info"
  | "raw"
  | "skipRecordsWithError"
  | "skipRecordsWithEmptyValues"
  | "onSkip"
>;

type CsvOptionsFormatFields = Pick<
  CsvFormatOptions,
  | "lineEnding"
  | "decimalSeparator"
  | "quoteColumns"
  | "quoteHeaders"
  | "writeHeaders"
  | "escapeFormulae"
>;

interface CsvOptionsExtras {
  sheetName?: string;
  sheetId?: number;
  /**
   * Append mode - when true, data is appended to existing file.
   * Header row is automatically skipped in append mode.
   * @default false
   */
  append?: boolean;
  dateFormats?: readonly DateFormat[];
  dateFormat?: string;
  dateUTC?: boolean;
  map?(value: CellValue, index: number): CellValue;
  includeEmptyRows?: boolean;
  requestHeaders?: Record<string, string>;
  requestBody?: NonNullable<RequestInit["body"]>;
  withCredentials?: boolean;
  signal?: AbortSignal;
  encoding?: string;
  onProgress?: (loaded: number, total: number) => void;
  stream?: boolean;
  highWaterMark?: number;
}

/** Unified CSV options for both parsing and formatting. */
export interface CsvOptions
  extends CsvOptionsParseFields, CsvOptionsFormatFields, CsvOptionsExtras {}

// =============================================================================
// Constants
// =============================================================================

const DEFAULT_DATE_FORMATS: readonly DateFormat[] = [
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

function createDefaultValueMapper(
  dateFormats: readonly DateFormat[],
  options?: { decimalSeparator?: DecimalSeparator }
): (datum: CellValue) => CellValue {
  const dateParser = DateParser.create(dateFormats);
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

function createDefaultWriteMapper(
  dateFormat?: string,
  dateUTC?: boolean
): (value: CellValue) => CellValue {
  const formatter = dateFormat
    ? DateFormatter.create(dateFormat, { utc: dateUTC })
    : DateFormatter.iso(dateUTC);

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
    if ("richText" in value && Array.isArray((value as any).richText)) {
      return (value as any).richText.map((r: { text: string }) => r.text).join("");
    }
    if ("checkbox" in value && typeof (value as any).checkbox === "boolean") {
      return (value as any).checkbox;
    }
    if ("error" in value && typeof (value as any).error === "string") {
      return (value as any).error;
    }
    return JSON.stringify(value);
  };
}

// =============================================================================
// Input detection
// =============================================================================

function isUrl(input: unknown): input is string {
  return typeof input === "string" && /^https?:\/\//i.test(input);
}

function isFile(input: unknown): input is File {
  return typeof File !== "undefined" && input instanceof File;
}

function isBlob(input: unknown): input is Blob {
  return typeof Blob !== "undefined" && input instanceof Blob && !isFile(input);
}

function isReadableStream(input: unknown): input is IReadable<any> {
  if (!input || typeof input !== "object") {
    return false;
  }
  const obj = input as any;
  return (
    typeof obj[Symbol.asyncIterator] === "function" ||
    (typeof obj.pipe === "function" && typeof obj.on === "function")
  );
}

// =============================================================================
// Stream helpers
// =============================================================================

function* iterateWorksheetRows(worksheet: any): Generator<{ row: any; rowNumber: number }> {
  const rows = (worksheet as { _rows?: any[] })._rows;
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

function buildParserOptions(options?: CsvOptions): Partial<CsvParseOptions> {
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
    info: options?.info,
    raw: options?.raw,
    skipRecordsWithError: options?.skipRecordsWithError,
    skipRecordsWithEmptyValues: options?.skipRecordsWithEmptyValues,
    onSkip: options?.onSkip
  };
}

function buildFormatterOptions(options?: CsvOptions) {
  return {
    delimiter: options?.delimiter ?? ",",
    quote: options?.quote,
    escape: options?.escape,
    lineEnding: options?.lineEnding,
    quoteColumns: options?.quoteColumns,
    quoteHeaders: options?.quoteHeaders,
    decimalSeparator: options?.decimalSeparator ?? ".",
    escapeFormulae: options?.escapeFormulae ?? true,
    writeHeaders: options?.writeHeaders
  };
}

// =============================================================================
// Read (cross-platform)
// =============================================================================

/** @internal — shared by read entry points and the Node file variant. */
export function readCsvContent(
  workbook: Workbook,
  content: string | ArrayBuffer | Uint8Array,
  options?: CsvOptions
): Worksheet {
  let str: string;
  if (typeof content === "string") {
    str = content;
  } else if (content instanceof ArrayBuffer || content instanceof Uint8Array) {
    str = new TextDecoder().decode(content);
  } else {
    str = String(content);
  }

  const worksheet = addWorksheet(workbook, options?.sheetName);
  const dateFormats = options?.dateFormats ?? DEFAULT_DATE_FORMATS;
  const decimalSeparator = options?.decimalSeparator;
  const map = options?.map || createDefaultValueMapper(dateFormats, { decimalSeparator });
  const result = parseCsv(str, buildParserOptions(options));

  if (Array.isArray(result)) {
    for (const row of result) {
      addRow(worksheet, row.map(map));
    }
  } else {
    if (result.headers) {
      addRow(worksheet, result.headers);
    }
    for (const rowObj of result.rows) {
      const rowArray = result.headers!.map(h => rowObj[h]);
      addRow(worksheet, rowArray.map(map));
    }
  }

  return worksheet;
}

/** @internal — shared by stream read entry points and the Node file variant. */
export function readCsvStream(
  workbook: Workbook,
  stream: IReadable<any>,
  options?: CsvOptions
): Promise<Worksheet> {
  const worksheet = addWorksheet(workbook, options?.sheetName);
  const dateFormats = options?.dateFormats ?? DEFAULT_DATE_FORMATS;
  const decimalSeparator = options?.decimalSeparator;
  const map = options?.map || createDefaultValueMapper(dateFormats, { decimalSeparator });
  const parser = new CsvParserStream(buildParserOptions(options));
  const useHeaders = !!options?.headers;
  let headerRow: string[] | null = null;

  return new Promise((resolve, reject) => {
    if (useHeaders) {
      parser.on("headers", (headers: string[]) => {
        headerRow = headers;
        addRow(worksheet, headers);
      });
    }

    parser.on("data", (row: unknown) => {
      if (useHeaders && headerRow && row && typeof row === "object" && !Array.isArray(row)) {
        const rowObj = row as Record<string, CellValue>;
        const rowArray: CellValue[] = headerRow.map(h => rowObj[h]);
        addRow(worksheet, rowArray.map(map));
      } else if (Array.isArray(row)) {
        addRow(worksheet, (row as CellValue[]).map(map));
      }
    });

    pipeline(stream, parser)
      .then(() => resolve(worksheet))
      .catch(reject);
  });
}

async function readCsvUrl(
  workbook: Workbook,
  url: string,
  options?: CsvOptions
): Promise<Worksheet> {
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
    return readCsvStream(workbook, readable as any, options);
  }

  const text = await response.text();
  return readCsvContent(workbook, text, options);
}

async function readCsvFileObject(
  workbook: Workbook,
  file: File,
  options?: CsvOptions
): Promise<Worksheet> {
  const LARGE_FILE_THRESHOLD = 10 * 1024 * 1024;
  if ((options?.stream || file.size > LARGE_FILE_THRESHOLD) && typeof file.stream === "function") {
    const readable = readableStreamToAsyncIterable<Uint8Array>(file.stream());
    return readCsvStream(workbook, readable as any, options);
  }

  return new Promise<Worksheet>((resolve, reject) => {
    const reader = new FileReader();
    const encoding = options?.encoding ?? "UTF-8";

    if (options?.onProgress) {
      reader.onprogress = event => {
        options.onProgress!(event.loaded, event.total || file.size);
      };
    }

    reader.onload = event => {
      try {
        const content = event.target?.result as string;
        resolve(readCsvContent(workbook, content, options));
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(new Error(`Failed to read file: ${file.name}`));
    reader.readAsText(file, encoding);
  });
}

async function readCsvBlob(
  workbook: Workbook,
  blob: Blob,
  options?: CsvOptions
): Promise<Worksheet> {
  const text = await blob.text();
  return readCsvContent(workbook, text, options);
}

/**
 * Read CSV into a new worksheet on `workbook`. Accepts a CSV string, URL,
 * `ArrayBuffer`/`Uint8Array`, browser `File`/`Blob`, or a readable stream.
 *
 * @example
 * ```ts
 * import { readCsv } from "documonster/excel/csv";
 * await readCsv(workbook, "a,b,c\n1,2,3");
 * await readCsv(workbook, "https://example.com/data.csv");
 * await readCsv(workbook, input, { delimiter: ";", sheetName: "Data" });
 * ```
 */
export async function readCsv(
  workbook: Workbook,
  input: CsvInput,
  options?: CsvOptions
): Promise<Worksheet> {
  if (isUrl(input)) {
    return readCsvUrl(workbook, input, options);
  }
  if (isFile(input)) {
    return readCsvFileObject(workbook, input, options);
  }
  if (isBlob(input)) {
    return readCsvBlob(workbook, input, options);
  }
  if (isReadableStream(input)) {
    return readCsvStream(workbook, input, options);
  }
  return readCsvContent(workbook, input, options);
}

// =============================================================================
// Write (cross-platform)
// =============================================================================

/** @internal — shared by write entry points and the Node file variant. */
export function writeCsvString(workbook: Workbook, options?: CsvOptions): string {
  const worksheet = getWorksheet(workbook, options?.sheetName || options?.sheetId);
  if (!worksheet) {
    return "";
  }

  const map = options?.map || createDefaultWriteMapper(options?.dateFormat, options?.dateUTC);
  const includeEmptyRows = options?.includeEmptyRows !== false;
  const rows: any[][] = [];
  let lastRow = 1;

  eachRow(worksheet, (row: any, rowNumber: number) => {
    if (includeEmptyRows) {
      while (lastRow++ < rowNumber - 1) {
        rows.push([]);
      }
    }
    const values = rowValues(row);
    values.shift();
    rows.push(values.map(map));
    lastRow = rowNumber;
  });

  return formatCsv(rows, buildFormatterOptions(options));
}

/** @internal — shared by write entry points and the Node file variant. */
export async function writeCsvStream(
  workbook: Workbook,
  stream: IWritable<any>,
  options?: CsvOptions
): Promise<void> {
  const worksheet = getWorksheet(workbook, options?.sheetName || options?.sheetId);
  if (!worksheet) {
    stream.end();
    return;
  }

  const map = options?.map || createDefaultWriteMapper(options?.dateFormat, options?.dateUTC);
  const includeEmptyRows = options?.includeEmptyRows !== false;
  const formatter = new CsvFormatterStream(buildFormatterOptions(options));
  const pipelinePromise = pipeline(formatter, stream);

  const awaitFormatterDrain = createDrainRacer(formatter);

  const writeAndDrain = async (values: any[]): Promise<void> => {
    if (!formatter.write(values)) {
      await awaitFormatterDrain();
    }
  };

  try {
    let lastRow = 1;
    for (const { row, rowNumber } of iterateWorksheetRows(worksheet)) {
      const dataValues = rowValues(row).slice(1).map(map);

      if (includeEmptyRows) {
        while (lastRow++ < rowNumber - 1) {
          await writeAndDrain([]);
        }
      }
      await writeAndDrain(dataValues);
      lastRow = rowNumber;
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
 * @example
 * ```ts
 * const csv = writeCsv(workbook);
 * await writeCsv(workbook, outputStream, { sheetId: 1 });
 * ```
 */
export function writeCsv(workbook: Workbook, options?: CsvOptions): string;
export function writeCsv(
  workbook: Workbook,
  stream: IWritable<any>,
  options?: CsvOptions
): Promise<void>;
export function writeCsv(
  workbook: Workbook,
  streamOrOptions?: IWritable<any> | CsvOptions,
  options?: CsvOptions
): string | Promise<void> {
  if (streamOrOptions && typeof (streamOrOptions as any).write === "function") {
    return writeCsvStream(workbook, streamOrOptions as IWritable<any>, options);
  }
  return writeCsvString(workbook, streamOrOptions as CsvOptions | undefined);
}

/** Write a worksheet to a CSV buffer (`Uint8Array`). */
export async function writeCsvBuffer(
  workbook: Workbook,
  options?: CsvOptions
): Promise<Uint8Array> {
  return new TextEncoder().encode(writeCsvString(workbook, options));
}

// =============================================================================
// Streaming surfaces
// =============================================================================

/** Create a readable stream that outputs the worksheet as CSV. */
export function createCsvReadStream(workbook: Workbook, options?: CsvOptions): IReadable<any> {
  const worksheet = getWorksheet(workbook, options?.sheetName || options?.sheetId);
  const map = options?.map || createDefaultWriteMapper(options?.dateFormat, options?.dateUTC);
  const includeEmptyRows = options?.includeEmptyRows !== false;
  const formatter = new CsvFormatterStream(buildFormatterOptions(options));

  if (!worksheet) {
    setTimeout(() => formatter.end(), 0);
    return formatter;
  }

  const awaitFormatterDrain = createDrainRacer(formatter);

  const writeAndDrain = (values: any[]): Promise<void> | void => {
    if (formatter.write(values)) {
      return;
    }
    return awaitFormatterDrain();
  };

  (async () => {
    try {
      let lastRow = 1;
      for (const { row, rowNumber } of iterateWorksheetRows(worksheet)) {
        if (formatter.destroyed) {
          return;
        }
        const dataValues = rowValues(row).slice(1).map(map);

        if (includeEmptyRows) {
          while (lastRow++ < rowNumber - 1) {
            const p = writeAndDrain([]);
            if (p) {
              await p;
            }
            if (formatter.destroyed) {
              return;
            }
          }
        }
        const p = writeAndDrain(dataValues);
        if (p) {
          await p;
        }
        lastRow = rowNumber;
      }
      formatter.end();
    } catch (err) {
      formatter.destroy(err instanceof Error ? err : new Error(String(err)));
    }
  })();

  return formatter;
}

/** Create a writable stream that parses CSV into a new worksheet. */
export function createCsvWriteStream(workbook: Workbook, options?: CsvOptions): IWritable<any> {
  const worksheet = addWorksheet(workbook, options?.sheetName);
  const dateFormats = options?.dateFormats ?? DEFAULT_DATE_FORMATS;
  const decimalSeparator = options?.decimalSeparator;
  const map = options?.map || createDefaultValueMapper(dateFormats, { decimalSeparator });
  const parser = new CsvParserStream(buildParserOptions(options));
  const useHeaders = !!options?.headers;
  let headerRow: string[] | null = null;

  if (useHeaders) {
    parser.on("headers", (headers: string[]) => {
      headerRow = headers;
      addRow(worksheet, headers);
    });
  }

  parser.on("data", (row: unknown) => {
    if (useHeaders && headerRow && row && typeof row === "object" && !Array.isArray(row)) {
      const rowObj = row as Record<string, CellValue>;
      const rowArray: CellValue[] = headerRow.map(h => rowObj[h]);
      addRow(worksheet, rowArray.map(map));
    } else if (Array.isArray(row)) {
      addRow(worksheet, (row as CellValue[]).map(map));
    }
  });

  return parser;
}
