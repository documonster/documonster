/**
 * Workbook - Cross-platform Excel Workbook (Browser Version)
 *
 * Base implementation with all platform-agnostic functionality.
 * Node.js version (workbook.ts) extends this with file system support for CSV.
 *
 * Features:
 * - xlsx: File/stream/buffer support (file operations Node.js only)
 * - csv: CSV read/write support (file operations Node.js only)
 * - streaming: createStreamWriter/createStreamReader for large files
 */

import { Worksheet, type WorksheetModel } from "@excel/worksheet";
import { DefinedNames, type DefinedNameModel } from "@excel/defined-names";
import { XLSX } from "@excel/xlsx/xlsx";
import { WorkbookWriter, type WorkbookWriterOptions } from "@excel/stream/workbook-writer";
import { WorkbookReader, type WorkbookReaderOptions } from "@excel/stream/workbook-reader";
import { DateParser, DateFormatter, type DateFormat } from "@utils/datetime";
import { parseCsv } from "@csv/parse";
import { formatCsv } from "@csv/format";
import type { CsvParseOptions, CsvFormatOptions } from "@csv/types";
import { CsvParserStream, CsvFormatterStream } from "@csv/stream";
import { parseNumberFromCsv, type DecimalSeparator } from "@csv/utils/number";
import { ExcelDownloadError, ExcelNotSupportedError } from "@excel/errors";
import { worksheetsToPdf, type ToPdfOptions } from "@pdf/pdf-converter";
import { pipeline } from "@stream";
import { readableStreamToAsyncIterable } from "@stream/utils.base";
import type { Readable } from "@stream";
import type { IReadable, IWritable } from "@stream/types";
import type { PivotTable } from "@excel/pivot-table";
import type {
  AddWorksheetOptions,
  CalculationProperties,
  CellErrorValue,
  ImageData,
  WorkbookProperties,
  WorkbookView,
  Buffer as ExcelBuffer
} from "@excel/types";

// =============================================================================
// Internal Types
// =============================================================================

/** Internal media type - more flexible than public Media type */
export interface WorkbookMedia {
  type: string;
  extension: string;
  filename?: string;
  buffer?: ExcelBuffer | Uint8Array;
  base64?: string;
  name?: string;
}

/** Internal model type for serialization */
export interface WorkbookModel {
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  created: Date;
  modified: Date;
  properties: Partial<WorkbookProperties>;
  worksheets: WorksheetModel[];
  sheets?: WorksheetModel[];
  definedNames: DefinedNameModel[];
  views: WorkbookView[];
  company: string;
  manager: string;
  title: string;
  subject: string;
  keywords: string;
  category: string;
  description: string;
  language?: string;
  revision?: number;
  contentStatus?: string;
  themes?: unknown;
  media: WorkbookMedia[];
  pivotTables: PivotTable[];
  /** Loaded pivot tables from file - used during reconciliation */
  loadedPivotTables?: any[];
  calcProperties: Partial<CalculationProperties>;
  /** Passthrough files (charts, etc.) preserved for round-trip */
  passthrough?: Record<string, Uint8Array>;
  /** Raw drawing XML data for passthrough (when drawing contains chart references) */
  rawDrawings?: Record<string, Uint8Array>;
  /** Default font preserved from the original file for round-trip fidelity */
  defaultFont?: any;
}

// =============================================================================
// CSV Types
// =============================================================================

/**
 * Supported input types for CSV parsing
 */
export type CsvInput =
  | string // CSV string or URL (http:// or https://)
  | ArrayBuffer
  | Uint8Array
  | File // Browser File object
  | Blob // Browser Blob object
  | IReadable<any>; // Readable stream

/**
 * Parse options from CsvParseOptions that are exposed in CsvOptions.
 * Internal fields like objectMode, transform, validate, chunk, etc. are excluded.
 */
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

/**
 * Format options from CsvFormatOptions that are exposed in CsvOptions.
 */
type CsvOptionsFormatFields = Pick<
  CsvFormatOptions,
  | "lineEnding"
  | "decimalSeparator"
  | "quoteColumns"
  | "quoteHeaders"
  | "writeHeaders"
  | "escapeFormulae"
>;

/**
 * CsvOptions-specific fields not present in CsvParseOptions or CsvFormatOptions.
 */
interface CsvOptionsExtras {
  // === Worksheet ===
  sheetName?: string;
  sheetId?: number;

  // === File write options ===
  /**
   * Append mode - when true, data is appended to existing file.
   * Header row is automatically skipped in append mode.
   * If file doesn't exist, it will be created (with headers if configured).
   * @default false
   */
  append?: boolean;

  // === Value mapping ===
  dateFormats?: readonly DateFormat[];
  dateFormat?: string;
  dateUTC?: boolean;
  map?(value: any, index: number): any;
  includeEmptyRows?: boolean;

  // === Network options (for URL input) ===
  requestHeaders?: Record<string, string>;
  requestBody?: BodyInit;
  withCredentials?: boolean;
  signal?: AbortSignal;

  // === File options ===
  encoding?: string;
  onProgress?: (loaded: number, total: number) => void;

  // === Stream options ===
  stream?: boolean;
  highWaterMark?: number;
}

/**
 * Unified CSV options for both parsing and formatting
 */
export interface CsvOptions
  extends CsvOptionsParseFields, CsvOptionsFormatFields, CsvOptionsExtras {}

// =============================================================================
// CSV Constants
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
// CSV Value Mappers (Internal)
// =============================================================================

/**
 * Create a value mapper for parsing CSV values into Excel-compatible types.
 * Converts strings to numbers, dates, booleans, and Excel error values.
 */
function createDefaultValueMapper(
  dateFormats: readonly DateFormat[],
  options?: { decimalSeparator?: DecimalSeparator }
) {
  const dateParser = DateParser.create(dateFormats);
  const decimalSeparator: DecimalSeparator = options?.decimalSeparator ?? ".";

  return function mapValue(datum: any): any {
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

    const date = dateParser.parse(datum);
    if (date) {
      return date;
    }

    const special = SpecialValues[datum];
    if (special !== undefined) {
      return special;
    }

    return datum;
  };
}

/**
 * Create a value mapper for writing Excel values to CSV format.
 * Handles hyperlinks, formulas, rich text, dates, errors, and objects.
 */
function createDefaultWriteMapper(dateFormat?: string, dateUTC?: boolean) {
  const formatter = dateFormat
    ? DateFormatter.create(dateFormat, { utc: dateUTC })
    : DateFormatter.iso(dateUTC);

  return function mapValue(value: any): any {
    if (value) {
      if (value.text || value.hyperlink) {
        return value.hyperlink || value.text || "";
      }
      if (value.formula || value.result) {
        return value.result || "";
      }
      // Handle rich text - extract and concatenate all text fragments
      if (value.richText && Array.isArray(value.richText)) {
        return value.richText.map((r: { text: string }) => r.text).join("");
      }
      if (value instanceof Date) {
        return formatter.format(value);
      }
      if (value.error) {
        return value.error;
      }
      if (typeof value === "object") {
        return JSON.stringify(value);
      }
    }
    return value;
  };
}

// =============================================================================
// CSV Input Type Detection (Internal)
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
// CSV Helper Functions (Internal)
// =============================================================================

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
    escapeFormulae: options?.escapeFormulae,
    writeHeaders: options?.writeHeaders
  };
}

// =============================================================================
// Workbook Class
// =============================================================================

class Workbook {
  // ===========================================================================
  // Static Properties
  // ===========================================================================

  /**
   * Streaming workbook writer class for large files.
   * @example
   * // Node.js: new Workbook.Writer({ filename: "large.xlsx" })
   * // Browser: new Workbook.Writer({ stream: writableStream })
   */
  static Writer = WorkbookWriter;

  /**
   * Streaming workbook reader class for large files.
   * @example
   * // Node.js: new Workbook.Reader("large.xlsx")
   * // Browser: new Workbook.Reader(readableStream)
   */
  static Reader = WorkbookReader;

  // ===========================================================================
  // Instance Properties - Metadata
  // ===========================================================================

  declare public category: string;
  declare public company: string;
  declare public created: Date;
  declare public description: string;
  declare public keywords: string;
  declare public manager: string;
  declare public modified: Date;
  declare public subject: string;
  declare public title: string;
  declare public creator?: string;
  declare public lastModifiedBy?: string;
  declare public lastPrinted?: Date;
  declare public language?: string;
  declare public revision?: number;
  declare public contentStatus?: string;

  // ===========================================================================
  // Instance Properties - Data
  // ===========================================================================

  declare public properties: Partial<WorkbookProperties>;
  declare public calcProperties: Partial<CalculationProperties>;
  declare public views: WorkbookView[];
  declare public media: WorkbookMedia[];
  declare public pivotTables: PivotTable[];

  // ===========================================================================
  // Private Properties
  // ===========================================================================

  declare protected _worksheets: Worksheet[];
  declare protected _definedNames: DefinedNames;
  declare protected _themes?: unknown;
  /** Passthrough files (charts, etc.) preserved for round-trip */
  declare protected _passthrough: Record<string, Uint8Array>;
  /** Raw drawing XML data for passthrough (when drawing contains chart references) */
  declare protected _rawDrawings: Record<string, Uint8Array>;
  /** Default font preserved from original file for round-trip fidelity */
  declare protected _defaultFont?: any;
  private _xlsx?: XLSX;

  // ===========================================================================
  // Constructor
  // ===========================================================================

  constructor() {
    this.category = "";
    this.company = "";
    this.created = new Date();
    this.description = "";
    this.keywords = "";
    this.manager = "";
    this.modified = this.created;
    this.properties = {};
    this.calcProperties = {};
    this._worksheets = [];
    this.subject = "";
    this.title = "";
    this.views = [];
    this.media = [];
    this.pivotTables = [];
    this._passthrough = {};
    this._rawDrawings = {};
    this._definedNames = new DefinedNames();
  }

  // ===========================================================================
  // Sheet Import
  // ===========================================================================

  /**
   * Import a worksheet from another workbook (or a standalone worksheet).
   * Deep-copies all worksheet properties via the model getter/setter, including
   * cell values, styles, merges, row heights, column widths, data validations,
   * conditional formatting, images, views, page setup, auto filter, tables,
   * sheet protection, page breaks, and drawing data.
   *
   * @param source - The worksheet to import
   * @param name - Optional name for the new worksheet (defaults to source name)
   * @returns The newly created worksheet
   */
  importSheet(source: Worksheet, name?: string): Worksheet {
    const newWs = this.addWorksheet(name ?? source.name);

    // Deep copy via model: the getter serializes ALL worksheet properties and the
    // setter deserializes them, so future properties are automatically included.
    const sourceModel = source.model;
    newWs.model = {
      ...sourceModel,
      id: newWs.id,
      name: newWs.name
    };

    return newWs;
  }

  // ===========================================================================
  // Format Operations (xlsx)
  // ===========================================================================

  /**
   * xlsx file format operations
   * Node.js: readFile, writeFile, read (stream), write (stream), load (buffer), writeBuffer
   * Browser: load (buffer), writeBuffer
   */
  get xlsx(): XLSX {
    if (!this._xlsx) {
      this._xlsx = new XLSX(this);
    }
    return this._xlsx;
  }

  // ===========================================================================
  // CSV Operations
  // ===========================================================================

  /**
   * Read CSV from any supported input source and add as worksheet
   *
   * @example
   * ```ts
   * // String
   * await workbook.readCsv("a,b,c\n1,2,3");
   *
   * // URL
   * await workbook.readCsv("https://example.com/data.csv");
   *
   * // File (browser)
   * await workbook.readCsv(fileInput.files[0]);
   *
   * // Stream
   * await workbook.readCsv(readableStream);
   *
   * // With options
   * await workbook.readCsv(input, { delimiter: ";", sheetName: "Data" });
   * ```
   */
  async readCsv(input: CsvInput, options?: CsvOptions): Promise<Worksheet> {
    if (isUrl(input)) {
      return this._readCsvUrl(input, options);
    }
    if (isFile(input)) {
      return this._readCsvFile(input, options);
    }
    if (isBlob(input)) {
      return this._readCsvBlob(input, options);
    }
    if (isReadableStream(input)) {
      return this._readCsvStream(input, options);
    }
    return this._readCsvContent(input, options);
  }

  /**
   * Write worksheet to CSV
   *
   * @example
   * ```ts
   * // Return CSV string
   * const csvString = workbook.writeCsv();
   * const csvString = workbook.writeCsv({ delimiter: ";", sheetName: "Data" });
   *
   * // Write to stream
   * await workbook.writeCsv(outputStream);
   * await workbook.writeCsv(outputStream, { sheetId: 1 });
   * ```
   */
  writeCsv(options?: CsvOptions): string;
  writeCsv(stream: IWritable<any>, options?: CsvOptions): Promise<void>;
  writeCsv(
    streamOrOptions?: IWritable<any> | CsvOptions,
    options?: CsvOptions
  ): string | Promise<void> {
    // If first argument is a stream (has write/end methods)
    if (streamOrOptions && typeof (streamOrOptions as any).write === "function") {
      return this._writeCsvStream(streamOrOptions as IWritable<any>, options);
    }
    // Otherwise treat first argument as options
    return this._writeCsvString(streamOrOptions as CsvOptions | undefined);
  }

  /**
   * Write worksheet to CSV buffer (Uint8Array)
   *
   * @example
   * ```ts
   * const buffer = await workbook.writeCsvBuffer();
   * const buffer = await workbook.writeCsvBuffer({ delimiter: ";", sheetName: "Data" });
   * ```
   */
  async writeCsvBuffer(options?: CsvOptions): Promise<Uint8Array> {
    const csvString = this._writeCsvString(options);
    return new TextEncoder().encode(csvString);
  }

  /**
   * Read CSV from file (Node.js only - throws in browser)
   */
  async readCsvFile(_filename: string, _options?: CsvOptions): Promise<Worksheet> {
    throw new ExcelNotSupportedError(
      "readCsvFile()",
      "not available in browser. Use readCsv(url) or readCsv(file) instead."
    );
  }

  /**
   * Write CSV to file (Node.js only - throws in browser)
   */
  async writeCsvFile(_filename: string, _options?: CsvOptions): Promise<void> {
    throw new ExcelNotSupportedError(
      "writeCsvFile()",
      "not available in browser. Use writeCsv() and trigger a download instead."
    );
  }

  /**
   * Create a readable stream that outputs CSV data
   *
   * @example
   * ```ts
   * const csvStream = workbook.createCsvReadStream();
   * csvStream.pipe(response); // pipe to HTTP response
   * ```
   */
  createCsvReadStream(options?: CsvOptions): IReadable<any> {
    const worksheet = this.getWorksheet(options?.sheetName || options?.sheetId);
    const map = options?.map || createDefaultWriteMapper(options?.dateFormat, options?.dateUTC);
    const includeEmptyRows = options?.includeEmptyRows !== false;
    const formatter = new CsvFormatterStream(buildFormatterOptions(options));

    if (worksheet) {
      setTimeout(() => {
        let lastRow = 1;
        worksheet.eachRow((row: any, rowNumber: number) => {
          if (includeEmptyRows) {
            while (lastRow++ < rowNumber - 1) {
              formatter.write([]);
            }
          }
          const { values } = row;
          values.shift();
          formatter.write(values.map(map));
          lastRow = rowNumber;
        });
        formatter.end();
      }, 0);
    } else {
      setTimeout(() => formatter.end(), 0);
    }

    return formatter;
  }

  /**
   * Create a writable stream that accepts CSV data and adds to worksheet
   *
   * @example
   * ```ts
   * const csvStream = workbook.createCsvWriteStream({ sheetName: "Data" });
   * inputStream.pipe(csvStream);
   * ```
   */
  createCsvWriteStream(options?: CsvOptions): IWritable<any> {
    const worksheet = this.addWorksheet(options?.sheetName);
    const dateFormats = options?.dateFormats ?? DEFAULT_DATE_FORMATS;
    const decimalSeparator = options?.decimalSeparator;
    const map = options?.map || createDefaultValueMapper(dateFormats, { decimalSeparator });
    const parser = new CsvParserStream(buildParserOptions(options));
    const useHeaders = !!options?.headers;
    let headerRow: string[] | null = null;

    // When headers option is enabled, listen for headers event to write header row first
    if (useHeaders) {
      parser.on("headers", (headers: string[]) => {
        headerRow = headers;
        worksheet.addRow(headers);
      });
    }

    parser.on("data", (row: unknown) => {
      // When headers: true, CsvParserStream emits objects; otherwise arrays
      if (useHeaders && headerRow && row && typeof row === "object" && !Array.isArray(row)) {
        // Convert object row to array using header order
        const rowObj = row as Record<string, unknown>;
        const rowArray = headerRow.map(h => rowObj[h]);
        worksheet.addRow(rowArray.map(map));
      } else if (Array.isArray(row)) {
        worksheet.addRow((row as unknown[]).map(map));
      }
    });

    return parser;
  }

  // ===========================================================================
  // CSV Internal Methods
  // ===========================================================================

  /** @internal */
  protected _readCsvContent(
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

    const worksheet = this.addWorksheet(options?.sheetName);
    const dateFormats = options?.dateFormats ?? DEFAULT_DATE_FORMATS;
    const decimalSeparator = options?.decimalSeparator;
    const map = options?.map || createDefaultValueMapper(dateFormats, { decimalSeparator });
    const result = parseCsv(str, buildParserOptions(options));

    if (Array.isArray(result)) {
      for (const row of result) {
        worksheet.addRow(row.map(map));
      }
    } else {
      if (result.headers) {
        worksheet.addRow(result.headers);
      }
      for (const rowObj of result.rows) {
        const rowArray = result.headers!.map(h => rowObj[h]);
        worksheet.addRow(rowArray.map(map));
      }
    }

    return worksheet;
  }

  /** @internal */
  protected async _readCsvStream(stream: IReadable<any>, options?: CsvOptions): Promise<Worksheet> {
    const worksheet = this.addWorksheet(options?.sheetName);
    const dateFormats = options?.dateFormats ?? DEFAULT_DATE_FORMATS;
    const decimalSeparator = options?.decimalSeparator;
    const map = options?.map || createDefaultValueMapper(dateFormats, { decimalSeparator });
    const parser = new CsvParserStream(buildParserOptions(options));
    const useHeaders = !!options?.headers;
    let headerRow: string[] | null = null;

    return new Promise((resolve, reject) => {
      // When headers option is enabled, listen for headers event to write header row first
      if (useHeaders) {
        parser.on("headers", (headers: string[]) => {
          headerRow = headers;
          worksheet.addRow(headers);
        });
      }

      parser.on("data", (row: unknown) => {
        // When headers: true, CsvParserStream emits objects; otherwise arrays
        if (useHeaders && headerRow && row && typeof row === "object" && !Array.isArray(row)) {
          // Convert object row to array using header order
          const rowObj = row as Record<string, unknown>;
          const rowArray = headerRow.map(h => rowObj[h]);
          worksheet.addRow(rowArray.map(map));
        } else if (Array.isArray(row)) {
          worksheet.addRow(row.map(map));
        }
      });

      pipeline(stream, parser)
        .then(() => resolve(worksheet))
        .catch(reject);
    });
  }

  private async _readCsvUrl(url: string, options?: CsvOptions): Promise<Worksheet> {
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
      return this._readCsvStream(readable as any, options);
    }

    const text = await response.text();
    return this._readCsvContent(text, options);
  }

  private async _readCsvFile(file: File, options?: CsvOptions): Promise<Worksheet> {
    const LARGE_FILE_THRESHOLD = 10 * 1024 * 1024;
    if (
      (options?.stream || file.size > LARGE_FILE_THRESHOLD) &&
      typeof file.stream === "function"
    ) {
      const readable = readableStreamToAsyncIterable<Uint8Array>(file.stream());
      return this._readCsvStream(readable as any, options);
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
          resolve(this._readCsvContent(content, options));
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = () => reject(new Error(`Failed to read file: ${file.name}`));
      reader.readAsText(file, encoding);
    });
  }

  private async _readCsvBlob(blob: Blob, options?: CsvOptions): Promise<Worksheet> {
    const text = await blob.text();
    return this._readCsvContent(text, options);
  }

  private _writeCsvString(options?: CsvOptions): string {
    const worksheet = this.getWorksheet(options?.sheetName || options?.sheetId);
    if (!worksheet) {
      return "";
    }

    const map = options?.map || createDefaultWriteMapper(options?.dateFormat, options?.dateUTC);
    const includeEmptyRows = options?.includeEmptyRows !== false;
    const rows: any[][] = [];
    let lastRow = 1;

    worksheet.eachRow((row: any, rowNumber: number) => {
      if (includeEmptyRows) {
        while (lastRow++ < rowNumber - 1) {
          rows.push([]);
        }
      }
      const { values } = row;
      values.shift();
      rows.push(values.map(map));
      lastRow = rowNumber;
    });

    return formatCsv(rows, buildFormatterOptions(options));
  }

  /** @internal */
  protected async _writeCsvStream(stream: IWritable<any>, options?: CsvOptions): Promise<void> {
    const worksheet = this.getWorksheet(options?.sheetName || options?.sheetId);
    if (!worksheet) {
      stream.end();
      return;
    }

    const map = options?.map || createDefaultWriteMapper(options?.dateFormat, options?.dateUTC);
    const includeEmptyRows = options?.includeEmptyRows !== false;
    const formatter = new CsvFormatterStream(buildFormatterOptions(options));
    const pipelinePromise = pipeline(formatter, stream);

    let lastRow = 1;
    worksheet.eachRow((row: any, rowNumber: number) => {
      if (includeEmptyRows) {
        while (lastRow++ < rowNumber - 1) {
          formatter.write([]);
        }
      }
      const { values } = row;
      values.shift();
      formatter.write(values.map(map));
      lastRow = rowNumber;
    });

    formatter.end();
    await pipelinePromise;
  }

  // ===========================================================================
  // Static Factory Methods for Streaming
  // ===========================================================================

  /**
   * Create a streaming workbook writer for large files.
   * This is more memory-efficient than using Workbook for large datasets.
   *
   * @param options - Options for the workbook writer
   *   - Node.js: can use { filename } or { stream }
   *   - Browser: must use { stream }
   * @returns A new WorkbookWriter instance
   *
   * @example
   * ```ts
   * // Node.js with filename
   * const writer = Workbook.createStreamWriter({ filename: "large-file.xlsx" });
   *
   * // Browser or Node.js with stream
   * const writer = Workbook.createStreamWriter({ stream: writableStream });
   *
   * const sheet = writer.addWorksheet("Sheet1");
   * for (let i = 0; i < 1000000; i++) {
   *   sheet.addRow([i, `Row ${i}`]).commit();
   * }
   * await writer.commit();
   * ```
   */
  static createStreamWriter(options?: WorkbookWriterOptions): WorkbookWriter {
    return new WorkbookWriter(options);
  }

  /**
   * Create a streaming workbook reader for large files.
   * This is more memory-efficient than using Workbook.xlsx.readFile for large datasets.
   *
   * @param input - File path (Node.js only) or readable stream
   * @param options - Options for the workbook reader
   * @returns A new WorkbookReader instance
   *
   * @example
   * ```ts
   * // Node.js with file path
   * const reader = Workbook.createStreamReader("large-file.xlsx");
   *
   * // Browser or Node.js with stream
   * const reader = Workbook.createStreamReader(readableStream);
   *
   * for await (const event of reader) {
   *   if (event.eventType === "worksheet") {
   *     const worksheet = event.value;
   *     for await (const row of worksheet) {
   *       console.log(row.values);
   *     }
   *   }
   * }
   * ```
   */
  static createStreamReader(
    input: string | Readable,
    options?: WorkbookReaderOptions
  ): WorkbookReader {
    return new WorkbookReader(input, options);
  }

  // ===========================================================================
  // Worksheet Management
  // ===========================================================================

  get nextId(): number {
    // Find the next unique spot to add worksheet
    for (let i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  }

  /**
   * Add a new worksheet and return a reference to it
   */
  addWorksheet(name?: string, options?: AddWorksheetOptions): Worksheet {
    const id = this.nextId;

    const lastOrderNo = this._worksheets.reduce(
      (acc, ws) => ((ws && ws.orderNo) > acc ? ws.orderNo : acc),
      0
    );
    const worksheetOptions = {
      ...options,
      id,
      name,
      orderNo: lastOrderNo + 1,
      workbook: this as any
    };

    const worksheet = new Worksheet(worksheetOptions);

    this._worksheets[id] = worksheet;
    return worksheet;
  }

  removeWorksheetEx(worksheet: Worksheet): void {
    delete this._worksheets[worksheet.id];
  }

  removeWorksheet(id: number | string): void {
    const worksheet = this.getWorksheet(id);
    if (worksheet) {
      worksheet.destroy();
    }
  }

  /**
   * Fetch sheet by name or id
   */
  getWorksheet(id?: number | string): Worksheet | undefined {
    if (id === undefined) {
      return this._worksheets.find(Boolean);
    }
    if (typeof id === "number") {
      return this._worksheets[id];
    }
    if (typeof id === "string") {
      const idLower = id.toLowerCase();
      return this._worksheets.find(
        worksheet => worksheet && worksheet.name.toLowerCase() === idLower
      );
    }
    return undefined;
  }

  /**
   * Return a clone of worksheets in order
   */
  get worksheets(): Worksheet[] {
    return this._worksheets
      .slice(1)
      .sort((a, b) => a.orderNo - b.orderNo)
      .filter(Boolean);
  }

  /**
   * Iterate over all sheets.
   *
   * Note: `workbook.worksheets.forEach` will still work but this is better.
   */
  eachSheet(callback: (sheet: Worksheet, id: number) => void): void {
    this.worksheets.forEach(sheet => {
      callback(sheet, sheet.id);
    });
  }

  // ===========================================================================
  // Defined Names
  // ===========================================================================

  get definedNames(): DefinedNames {
    return this._definedNames;
  }

  // ===========================================================================
  // Themes
  // ===========================================================================

  clearThemes(): void {
    // Note: themes are not an exposed feature, meddle at your peril!
    this._themes = undefined;
  }

  // ===========================================================================
  // Images
  // ===========================================================================

  /**
   * Add Image to Workbook and return the id
   */
  addImage(image: ImageData): number {
    const id = this.media.length;
    this.media.push({ ...image, type: "image" });
    return id;
  }

  getImage(id: number | string): WorkbookMedia | undefined {
    return this.media[Number(id)];
  }

  // ===========================================================================
  // PDF Conversion
  // ===========================================================================

  /**
   * Convert this workbook to a PDF Uint8Array.
   *
   * Preserves styles (fonts, fills, borders, alignment), merged cells, and images.
   *
   * @example
   * ```ts
   * const pdfBuffer = workbook.toPdf();
   * // Node.js: write to file
   * fs.writeFileSync("output.pdf", pdfBuffer);
   * // Browser: create download link
   * const blob = new Blob([pdfBuffer], { type: "application/pdf" });
   * ```
   *
   * @param options - PDF generation options
   * @returns PDF file contents as Uint8Array
   */
  toPdf(options?: ToPdfOptions): Uint8Array {
    return worksheetsToPdf(this.worksheets, this.media, options);
  }

  // ===========================================================================
  // Model (Serialization)
  // ===========================================================================

  get model(): WorkbookModel {
    return {
      creator: this.creator || "Unknown",
      lastModifiedBy: this.lastModifiedBy || "Unknown",
      lastPrinted: this.lastPrinted,
      created: this.created,
      modified: this.modified,
      properties: this.properties,
      worksheets: this.worksheets.map(worksheet => worksheet.model),
      sheets: this.worksheets.map(ws => ws.model).filter(Boolean),
      definedNames: this._definedNames.model,
      views: this.views,
      company: this.company,
      manager: this.manager,
      title: this.title,
      subject: this.subject,
      keywords: this.keywords,
      category: this.category,
      description: this.description,
      language: this.language,
      revision: this.revision,
      contentStatus: this.contentStatus,
      themes: this._themes,
      media: this.media,
      pivotTables: this.pivotTables,
      calcProperties: this.calcProperties,
      passthrough: this._passthrough,
      rawDrawings: this._rawDrawings,
      defaultFont: this._defaultFont
    };
  }

  set model(value: WorkbookModel) {
    this.creator = value.creator;
    this.lastModifiedBy = value.lastModifiedBy;
    this.lastPrinted = value.lastPrinted;
    this.created = value.created;
    this.modified = value.modified;
    this.company = value.company;
    this.manager = value.manager;
    this.title = value.title;
    this.subject = value.subject;
    this.keywords = value.keywords;
    this.category = value.category;
    this.description = value.description;
    this.language = value.language;
    this.revision = value.revision;
    this.contentStatus = value.contentStatus;

    this.properties = value.properties;
    this.calcProperties = value.calcProperties;
    this._worksheets = [];
    value.worksheets.forEach(worksheetModel => {
      const { id, name, state } = worksheetModel;
      const orderNo = value.sheets && value.sheets.findIndex(ws => ws.id === id);
      const worksheet = (this._worksheets[id] = new Worksheet({
        id,
        name,
        orderNo: orderNo !== -1 ? orderNo : undefined,
        state,
        workbook: this as any
      }));
      worksheet.model = worksheetModel;
    });

    this._definedNames.model = value.definedNames;
    this.views = value.views;
    this._themes = value.themes;
    this.media = value.media || [];

    // Handle pivot tables - either newly created or loaded from file
    // Loaded pivot tables come from loadedPivotTables after reconciliation
    this.pivotTables = value.pivotTables || value.loadedPivotTables || [];

    // Preserve passthrough files (charts, etc.) for round-trip preservation
    this._passthrough = value.passthrough || {};
    // Preserve raw drawing data for drawings with chart references
    this._rawDrawings = value.rawDrawings || {};
    // Preserve default font for round-trip fidelity
    this._defaultFont = value.defaultFont;
  }
}

export { Workbook };
export type { ToPdfOptions } from "@pdf/pdf-converter";
