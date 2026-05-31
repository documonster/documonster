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

import { formatCsv } from "@csv/format";
import { parseCsv } from "@csv/parse";
import { CsvParserStream, CsvFormatterStream } from "@csv/stream";
import type { CsvParseOptions, CsvFormatOptions } from "@csv/types";
import { parseNumberFromCsv, type DecimalSeparator } from "@csv/utils/number";
// Chart runtime accessed through the host-registry slot; see
// `chart-host-registry.ts` and `chart/install.ts`. Keeps chart code
// out of consumer bundles when they don't import
// `@cj-tech-master/excelts/chart`.
import { getChartSupport } from "@excel/chart-host-registry";
import type { ChartEntry, ChartExEntry } from "@excel/chart/chart";
import type { AddChartExOptions, ChartExModel } from "@excel/chart/chart-ex-types";
import type { AddComboChartOptions, ChartModel } from "@excel/chart/types";
import {
  Chartsheet,
  type AddChartsheetOptions,
  type AddPivotChartsheetOptions
} from "@excel/chartsheet";
import { DefinedNames, type DefinedNameModel } from "@excel/defined-names";
import { ExcelDownloadError, ExcelNotSupportedError, WorksheetNameError } from "@excel/errors";
import { withPivotChartSource } from "@excel/pivot-chart";
import type { PivotTable } from "@excel/pivot-table";
import {
  WorkbookReader,
  type WorkbookReaderOptions,
  type CommonInput
} from "@excel/stream/workbook-reader.browser";
import { WorkbookWriter, type WorkbookWriterOptions } from "@excel/stream/workbook-writer.browser";
import type {
  AddWorksheetOptions,
  CalculationProperties,
  CellErrorValue,
  CellValue,
  Font,
  ImageData,
  ThreadedCommentPerson,
  WorkbookProperties,
  WorkbookProtection,
  WorkbookView,
  Buffer as ExcelBuffer
} from "@excel/types";
import { synthGuid } from "@excel/utils/guid";
import { buildWorkbookProtection } from "@excel/utils/workbook-protection";
import { Worksheet, type WorksheetModel } from "@excel/worksheet";
import { RelType } from "@excel/xlsx/rel-type";
import type { ChartsheetModel } from "@excel/xlsx/xform/sheet/chartsheet-xform";
import { XLSX } from "@excel/xlsx/xlsx.browser";
import type { SyntaxProbe } from "@formula/default-syntax-probe";
import { invokeFormulaEngine } from "@formula/host-registry";
import { formatMarkdown } from "@markdown/format/index";
import { parseMarkdown, parseMarkdownAll } from "@markdown/parse/index";
import type { MarkdownOptions, MarkdownAlignment, MarkdownParseResult } from "@markdown/types";
import { pipeline } from "@stream";
import type { IReadable, IWritable } from "@stream/types";
import { readableStreamToAsyncIterable } from "@stream/utils.base";
import { DateParser, DateFormatter, type DateFormat } from "@utils/datetime";

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
  /** External link target — when set, the image is referenced, not embedded. */
  link?: string;
}

/** Internal model type for serialization */
export interface WorkbookModel {
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  created: Date;
  modified: Date;
  properties: Partial<WorkbookProperties>;
  protection?: WorkbookProtectionModel;
  worksheets: WorksheetModel[];
  sheets?: WorksheetModel[];
  definedNames: DefinedNameModel[];
  /**
   * Live `DefinedNames` instance — used by the write-time chartEx
   * transform (`prepareChartExSidecars`) which registers hidden
   * `_xlchart.vN.M` defined names on the fly and needs an object
   * it can mutate in place. The serialised `definedNames` array
   * above is re-materialised from this instance after the
   * transform runs. Optional because the model is also used for
   * input paths that don't carry the live instance.
   */
  definedNamesInstance?: unknown;
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
  loadedPivotTables?: PivotTable[];
  calcProperties: Partial<CalculationProperties>;
  /** Default font preserved from the original file for round-trip fidelity */
  defaultFont?: Partial<Font>;
  /** Chart entries indexed by 1-based chart number */
  chartEntries?: Record<number, ChartEntry>;
  /** Chart rels indexed by chart number — preserved for round-trip */
  chartRels?: Record<number, any[]>;
  /** Chart style XML raw bytes indexed by style number — preserved for round-trip */
  chartStyles?: Record<number, Uint8Array>;
  /** Chart colors XML raw bytes indexed by colors number — preserved for round-trip */
  chartColors?: Record<number, Uint8Array>;
  chartExStyles?: Record<number, Uint8Array>;
  chartExColors?: Record<number, Uint8Array>;
  /** ChartEx raw bytes (Office 2016+ extended charts) indexed by chartEx number */
  chartExEntries?: Record<number, Uint8Array>;
  /** ChartEx rels indexed by chartEx number */
  chartExRels?: Record<number, any[]>;
  /** Structured chartEx entries (loaded or programmatically built) indexed by chartEx number */
  chartExStructuredEntries?: Record<number, ChartExEntry>;
  /** Chartsheets parsed from the XLSX file — preserved for round-trip */
  chartsheets?: any[];
  /**
   * Office 365 threaded-comment person directory, hydrated from
   * `xl/persons/person.xml` on load and serialised back on save when
   * non-empty. See {@link Workbook.persons}.
   */
  persons?: ThreadedCommentPerson[];
  /**
   * Raw-passthrough slicer parts keyed by zip-relative path. Excelts
   * does not structurally model slicers yet but preserves the bytes on
   * round-trip so dashboards continue to work.
   */
  slicerParts?: Record<string, Uint8Array>;
  slicerCacheParts?: Record<string, Uint8Array>;
  timelineParts?: Record<string, Uint8Array>;
  timelineCacheParts?: Record<string, Uint8Array>;
  /**
   * External workbook references in declaration order. Matches the on-disk
   * `[N]Sheet!Ref` indexing (1-based). Empty or undefined when the workbook
   * has no external references.
   */
  externalLinks?: ExternalLinkModel[];
}

/** Internal model for workbook-level protection (serialized to <workbookProtection>) */
export interface WorkbookProtectionModel {
  lockStructure?: boolean;
  lockWindows?: boolean;
  lockRevision?: boolean;
  workbookPassword?: string;
  revisionsPassword?: string;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

// =============================================================================
// External Workbook Link Types
// =============================================================================

/**
 * Cached values for a single sheet of an external workbook. Keys are the
 * A1-notation cell addresses *in uppercase* (e.g. `"A1"`, `"B12"`). Values
 * are the cached primitives Excel displays when the external file is not
 * currently available — must be JSON primitives: string, number, boolean, or
 * null for an explicitly blank cell.
 */
export type ExternalLinkCachedSheet = Record<string, string | number | boolean | null>;

/**
 * A single external workbook reference. Each entry corresponds to one
 * `xl/externalLinks/externalLink{N}.xml` part in the output file, and to
 * one `<externalReference r:id="...">` entry in `xl/workbook.xml`.
 *
 * The on-disk formula syntax for referring to this workbook is `[N]Sheet!A1`
 * where `N` is the 1-based `index` below.
 */
export interface ExternalLinkModel {
  /**
   * The 1-based index used in `[N]Sheet!A1` formulas. This is the position
   * in the workbook's `<externalReferences>` list (in declaration order).
   * Assigned automatically on read/write; treat as read-only when produced
   * by the library.
   */
  index: number;
  /**
   * The rel Target that will be written into
   * `xl/externalLinks/_rels/externalLink{N}.xml.rels`. For relative paths
   * (which is what users almost always want), pass the bare filename or a
   * path relative to the current workbook: `"测试.xlsx"`, `"data/ref.xlsx"`.
   * Office resolves bare relative paths from the current workbook's
   * directory — *that* is the fix for the "Office goes to the Documents
   * folder" problem with external links.
   *
   * Absolute `file:///` or `http(s)://` URIs are accepted and written
   * through unchanged.
   */
  target: string;
  /**
   * Almost always `"External"`. `"Internal"` is for embedded workbooks
   * (rare) and is preserved on round-trip when present in the source file.
   */
  targetMode: "External" | "Internal";
  /**
   * The relationship id inside `xl/_rels/workbook.xml.rels` pointing to this
   * external link's XML part. Populated automatically on read and
   * re-assigned on write. Callers should leave this undefined.
   */
  rId?: string;
  /**
   * The sheet names exposed by the external workbook, in declaration order.
   * Excel writes one `<sheetName val="..."/>` per entry under
   * `<sheetNames>` inside the externalLink part.
   *
   * At minimum you must declare every sheet that appears in a formula
   * targeting this external workbook, otherwise Excel will fail to link
   * the cached values and show `#REF!`.
   */
  sheetNames: string[];
  /**
   * Cached primitive values per sheet. Key is the *sheet name* (matching an
   * entry in `sheetNames`), value is a map from A1 address to primitive.
   *
   * Cached values are what Excel displays when the referenced external file
   * is not available (e.g. freshly-downloaded workbook on another machine).
   * Writing them turns your file from "opens with errors" into "opens,
   * shows values, offers to update links".
   */
  cachedValues?: Record<string, ExternalLinkCachedSheet>;
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
  /**
   * Transform each cell value as rows are parsed from CSV or formatted for
   * CSV output.
   *
   * - During parse: `value` is the raw JS value produced by the CSV parser
   *   (string, number, boolean, Date, ...) and already narrows to `CellValue`.
   * - During format: `value` is the worksheet cell's `CellValue`.
   *
   * The function should return a `CellValue`; returning non-`CellValue`
   * types (functions, symbols, ...) is unsupported and will break downstream
   * serialization.
   */
  map?(value: CellValue, index: number): CellValue;
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

/**
 * Create a value mapper for writing Excel values to CSV format.
 *
 * Branches below correspond to each member of the `CellValue` union, in an
 * order that prefers the most specific shape:
 *   1. Hyperlink   — emit the URL (falling back to display text)
 *   2. Formula / SharedFormula / ArrayFormula — emit the evaluated `result`
 *   3. RichText    — flatten runs to plain text
 *   4. Checkbox    — emit the boolean state
 *   5. Error       — emit the Excel error token (e.g. "#N/A")
 *   6. Date        — format via the configured DateFormatter
 *   7. Primitive   — pass through (`string | number | boolean | null | undefined`)
 *
 * An object that does not match any of the above is stringified as JSON so
 * the CSV remains round-trippable rather than producing `[object Object]`.
 */
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
      // string | number | boolean
      return value;
    }

    // --- Object variants of CellValue ---

    // Hyperlink: prefer URL, fall back to display text.
    // Accepts objects carrying either `hyperlink` or just `text` — the latter
    // is a historical input shape some callers use to denote a "link-like"
    // cell value without an actual URL; preserving it maintains backward
    // compatibility with pre-existing CSV output.
    const maybeLink = value as { hyperlink?: unknown; text?: unknown };
    if (typeof maybeLink.hyperlink === "string" || typeof maybeLink.text === "string") {
      const url = typeof maybeLink.hyperlink === "string" ? maybeLink.hyperlink : "";
      const text = typeof maybeLink.text === "string" ? maybeLink.text : "";
      return url || text || "";
    }
    // Formula / SharedFormula / ArrayFormula — all carry an evaluated `result`
    if ("formula" in value || "sharedFormula" in value) {
      return value.result ?? "";
    }
    // Rich text — flatten runs
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map(r => r.text).join("");
    }
    // Checkbox
    if ("checkbox" in value && typeof value.checkbox === "boolean") {
      return value.checkbox;
    }
    // Error
    if ("error" in value && typeof value.error === "string") {
      return value.error;
    }
    // Unknown object shape — round-trippable fallback
    return JSON.stringify(value);
  };
}

// =============================================================================
// Markdown Value Mapper (Internal)
// =============================================================================

/**
 * Create a stringify function for Markdown output.
 * Handles hyperlinks, formulas, rich text, dates, errors, and objects.
 */
function createMarkdownStringify(
  dateFormat?: string,
  dateUTC?: boolean
): (value: unknown) => string {
  const formatter = dateFormat
    ? DateFormatter.create(dateFormat, { utc: dateUTC })
    : DateFormatter.iso(dateUTC);

  return function stringify(value: unknown): string {
    if (value === null || value === undefined) {
      return "";
    }
    if (typeof value === "string") {
      return value;
    }
    if (typeof value === "number" || typeof value === "bigint") {
      return String(value);
    }
    if (typeof value === "boolean") {
      return value ? "true" : "false";
    }
    if (value instanceof Date) {
      return formatter.format(value);
    }
    if (typeof value === "object") {
      const v = value as any;
      if (v.text || v.hyperlink) {
        return v.hyperlink || v.text || "";
      }
      if (v.formula || v.result) {
        return v.result != null ? String(v.result) : "";
      }
      if (v.richText && Array.isArray(v.richText)) {
        return v.richText.map((r: { text: string }) => r.text).join("");
      }
      if (v.error) {
        return v.error;
      }
      try {
        return JSON.stringify(value);
      } catch {
        return "[object Object]";
      }
    }
    return String(value);
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

/**
 * Iterate worksheet rows lazily — yields `{row, rowNumber}` for every row
 * with values, in sheet order. Reaches into the worksheet's internal
 * `_rows` array directly so we don't have to materialise the whole row
 * set up front (matters for very large worksheets piped to slow sinks).
 *
 * Note: `_rows` is 0-based but XLSX row numbers are 1-based. Row N lives
 * at `_rows[N - 1]`, so the yielded `rowNumber` is `i + 1`.
 */
function* iterateWorksheetRows(worksheet: any): Generator<{ row: any; rowNumber: number }> {
  const rows = (worksheet as { _rows?: any[] })._rows;
  if (!rows || rows.length === 0) {
    return;
  }
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (row && row.hasValues) {
      yield { row, rowNumber: i + 1 };
    }
  }
}

/**
 * Build a `() => Promise<void>` that resolves on the next `'drain'` event
 * of `emitter`, but rejects promptly if `'error'` or `'close'` fires first.
 *
 * Without the error/close races, a producer parked on `once('drain')`
 * would hang forever after the downstream sink errors mid-write — the
 * Transform is destroyed and never emits drain again.
 */
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
    // Default to true for CSV injection protection when writing through Workbook.csv.
    // The low-level formatCsv() keeps its own default (false) for backward compatibility.
    escapeFormulae: options?.escapeFormulae ?? true,
    writeHeaders: options?.writeHeaders
  };
}

function isComboChartOptions(chart: AddChartsheetOptions["chart"]): chart is AddComboChartOptions {
  return !!chart && typeof chart === "object" && "groups" in chart;
}

function isChartExOptions(chart: AddChartsheetOptions["chart"]): chart is AddChartExOptions {
  return !!chart && typeof chart === "object" && "type" in chart && isChartExType(chart.type);
}

function isChartExType(type: string): boolean {
  return (
    type === "sunburst" ||
    type === "treemap" ||
    type === "waterfall" ||
    type === "funnel" ||
    type === "histogram" ||
    type === "pareto" ||
    type === "boxWhisker" ||
    type === "regionMap"
  );
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
  declare public protection?: WorkbookProtectionModel;
  /**
   * External workbook references, in declaration order. The 1-based index
   * of each entry matches the `[N]` prefix used inside formula strings
   * (e.g. the first entry is referenced as `[1]Sheet1!A1` on disk).
   *
   * Prefer {@link addExternalLink} for appending — it handles index
   * assignment and sheet-name deduplication. Direct mutation of this array
   * is supported but callers must keep indices contiguous starting at 1.
   */
  declare public externalLinks: ExternalLinkModel[];

  // ===========================================================================
  // Private Properties
  // ===========================================================================

  declare protected _worksheets: Worksheet[];
  declare protected _definedNames: DefinedNames;
  declare protected _themes?: unknown;
  /** Default font preserved from original file for round-trip fidelity */
  declare protected _defaultFont?: Partial<Font>;
  /**
   * Cache of external-workbook references auto-discovered from formula
   * strings during previous `writeBuffer()` calls. This is an internal
   * stash used to keep subsequent writes fixed-point stable: once a
   * formula has been normalised to `[N]Sheet!A1`, the writer needs the
   * corresponding link metadata on the next write too, but we don't want
   * those auto-discovered entries to appear on the user-facing
   * `externalLinks` list. Indexed by lower-cased target path.
   *
   * Entries explicitly added via `addExternalLink()` live on `externalLinks`
   * instead — the writer combines both at serialisation time.
   */
  declare protected _writerExternalLinkCache: Map<string, ExternalLinkModel>;
  /** Global registry of table names (lowercase) for cross-worksheet uniqueness checks. */
  readonly _tableNames = new Set<string>();
  /** Chart entries indexed by 1-based chart number */
  declare protected _chartEntries: Record<number, ChartEntry>;
  /** Chart rels indexed by chart number — preserved for round-trip */
  declare protected _chartRels: Record<number, any[]>;
  /** Chart style XML raw bytes indexed by style number — preserved for round-trip */
  declare protected _chartStyles: Record<number, Uint8Array>;
  /** Chart colors XML raw bytes indexed by colors number — preserved for round-trip */
  declare protected _chartColors: Record<number, Uint8Array>;
  declare protected _chartExStyles: Record<number, Uint8Array>;
  declare protected _chartExColors: Record<number, Uint8Array>;
  /** ChartEx raw bytes (Office 2016+ extended charts) indexed by chartEx number */
  declare protected _chartExEntries: Record<number, Uint8Array>;
  /** ChartEx rels indexed by chartEx number */
  declare protected _chartExRels: Record<number, any[]>;
  /** ChartEx structured entries (loaded or built programmatically via addChartEx). */
  declare protected _chartExStructuredEntries: Record<number, ChartExEntry>;
  /** Chartsheets parsed from the XLSX file — preserved for round-trip */
  declare protected _chartsheets: ChartsheetModel[];
  /**
   * Office 365 threaded-comment person directory (`xl/persons/person.xml`).
   * Referenced by per-sheet `threadedComment/@personId`. Hydrated on
   * load, preserved across save. Programmatic comments that don't supply
   * a personId get auto-registered against this list.
   */
  declare protected _persons: ThreadedCommentPerson[];
  /**
   * Raw XML passthrough for Office 2010+ slicers and timelines.
   *
   * Structured creation of these controls is out of scope for the
   * current release — the OOXML surface is large (four coordinated
   * part families: `xl/slicers`, `xl/slicerCaches`, `xl/timelines`,
   * `xl/timelineCaches`, plus sheet-level extensions and workbook-
   * level cache list entries). The passthrough here prevents silent
   * data loss when an Excel dashboard with slicers or timelines
   * travels through excelts: we capture every part verbatim on read
   * and re-emit them on write, along with their rels and Content
   * Types overrides.
   *
   * The map key is the full zip-relative path (e.g.
   * `"xl/slicers/slicer1.xml"`), the value is the exact bytes we
   * found in the input. Loaders that construct a workbook
   * programmatically never touch these maps.
   */
  declare protected _slicerParts: Record<string, Uint8Array>;
  declare protected _slicerCacheParts: Record<string, Uint8Array>;
  declare protected _timelineParts: Record<string, Uint8Array>;
  declare protected _timelineCacheParts: Record<string, Uint8Array>;
  private _xlsx?: XLSX;

  // ===========================================================================
  // Constructor
  // ===========================================================================

  /**
   * @param options Optional construction options.
   *   - `formulaSyntaxProbe`: An explicit tokenizer+parser probe used to
   *     classify defined-name text during XLSX load. Providing this makes
   *     classification deterministic for *this* workbook regardless of
   *     whether `installFormulaEngine()` has been called. Most callers
   *     don't need it — `installFormulaEngine()` registers a
   *     process-wide default probe that is picked up automatically.
   */
  constructor(options?: { formulaSyntaxProbe?: SyntaxProbe }) {
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
    this.externalLinks = [];
    this._chartEntries = {};
    this._chartRels = {};
    this._chartStyles = {};
    this._chartColors = {};
    this._chartExStyles = {};
    this._chartExColors = {};
    this._chartExEntries = {};
    this._chartExRels = {};
    this._chartExStructuredEntries = {};
    this._chartsheets = [];
    this._persons = [];
    this._slicerParts = {};
    this._slicerCacheParts = {};
    this._timelineParts = {};
    this._timelineCacheParts = {};
    this._writerExternalLinkCache = new Map();
    this._definedNames = new DefinedNames(options?.formulaSyntaxProbe);
  }

  // ===========================================================================
  // Default Font
  // ===========================================================================

  /**
   * The default font for the workbook (fontId=0 / "Normal" style).
   * Cells without explicit font styles will inherit this font in Excel.
   *
   * @example
   * ```ts
   * wb.defaultFont = { name: "Arial", size: 12 };
   * ```
   *
   * When reading an existing XLSX file, this preserves the original default font
   * for round-trip fidelity. Setting it on a new workbook changes the default
   * from Calibri 11 to your chosen font.
   */
  get defaultFont(): Partial<Font> | undefined {
    return this._defaultFont;
  }

  set defaultFont(font: Partial<Font> | undefined) {
    this._defaultFont = font;
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
    // Remap chart numbers so the source's `chartNumber` / `chartExNumber`
    // references point at entries we actually copy into the target
    // workbook. Build the map here so the rewritten `charts` array and
    // the copied entries use consistent ids.
    const chartMap = new Map<number, number>();
    const chartExMap = new Map<number, number>();
    const sourceWorkbook = source.workbook as unknown as Workbook;
    const differentWorkbook = sourceWorkbook !== (this as unknown as Workbook);
    const sourceCharts = sourceModel.charts ?? [];
    // `nextChartNumber()` / `nextChartExNumber()` compute `max(existing) + 1`
    // from the entry maps — they do NOT reserve a slot. Calling them in
    // a tight loop without an intervening `addChartEntry` therefore
    // returns the SAME number N times, and the second loop below then
    // overwrites `_chartEntries[dstNum]` repeatedly — only the last
    // cloned entry survives, the others are silently lost. Track the
    // allocator locally so each source chart gets a unique target slot.
    let nextChartAlloc = this.nextChartNumber();
    let nextChartExAlloc = this.nextChartExNumber();
    for (const anchor of sourceCharts) {
      if (anchor.chartNumber && anchor.chartNumber > 0 && !chartMap.has(anchor.chartNumber)) {
        chartMap.set(anchor.chartNumber, nextChartAlloc++);
      }
      if (
        anchor.chartExNumber &&
        anchor.chartExNumber > 0 &&
        !chartExMap.has(anchor.chartExNumber)
      ) {
        chartExMap.set(anchor.chartExNumber, nextChartExAlloc++);
      }
    }
    const remappedCharts = sourceCharts.map(anchor => ({
      ...anchor,
      chartNumber: anchor.chartNumber
        ? (chartMap.get(anchor.chartNumber) ?? anchor.chartNumber)
        : anchor.chartNumber,
      chartExNumber: anchor.chartExNumber
        ? (chartExMap.get(anchor.chartExNumber) ?? anchor.chartExNumber)
        : anchor.chartExNumber
    }));
    newWs.model = {
      ...sourceModel,
      id: newWs.id,
      name: newWs.name,
      charts: remappedCharts
    };

    // Copy the actual chart parts + sidecars into the target workbook
    // so the remapped `charts` array references live entries. Without
    // this, `importSheet` left the target with chart anchors but no
    // backing chart XML, producing a broken package on save. We copy
    // both the structured model (via `getChartEntry` / `addChartEntry`
    // — the public API) and all sidecars (`copyChartSidecars` /
    // `copyChartExSidecars`).
    if (chartMap.size > 0 || chartExMap.size > 0) {
      for (const [srcNum, dstNum] of chartMap) {
        const entry = sourceWorkbook.getChartEntry(srcNum);
        if (!entry) {
          continue;
        }
        // Deep-clone the entry with every metadata field preserved —
        // rawData / userShapesXml (byte slices), modelSnapshot and the
        // dirty / preferRawPatch / requireRawPatch writer hints, plus
        // per-entry `rels`. Previously only `model`, `rawData`, and
        // `userShapesXml` were copied, so the cross-workbook import
        // path produced charts where the raw-patch fast path couldn't
        // run and the change-detection snapshot didn't reflect the
        // source entry's load-time state.
        this.addChartEntry(cloneChartEntry(entry, dstNum));
        if (differentWorkbook) {
          sourceWorkbook.copyChartSidecars(srcNum, dstNum, this as unknown as Workbook);
        } else {
          this.copyChartSidecars(srcNum, dstNum);
        }
      }
      for (const [srcNum, dstNum] of chartExMap) {
        const exEntry = sourceWorkbook.getChartExStructuredEntry?.(srcNum);
        if (exEntry) {
          this.addChartExStructuredEntry(cloneChartExEntry(exEntry, dstNum));
        } else {
          const rawBytes = (
            sourceWorkbook as unknown as { _chartExEntries?: Record<number, Uint8Array> }
          )._chartExEntries?.[srcNum];
          if (rawBytes) {
            (this as unknown as { _chartExEntries: Record<number, Uint8Array> })._chartExEntries[
              dstNum
            ] = rawBytes.slice();
          }
        }
        if (differentWorkbook) {
          sourceWorkbook.copyChartExSidecars(srcNum, dstNum, this as unknown as Workbook);
        } else {
          this.copyChartExSidecars(srcNum, dstNum);
        }
      }
    }

    return newWs;
  }

  // ===========================================================================
  // Workbook Protection
  // ===========================================================================

  /**
   * Protect the workbook structure with an optional password.
   * Prevents users from adding, deleting, renaming, moving, or copying worksheets.
   *
   * @param password - Optional password to protect the structure
   * @param options  - Optional protection flags (lockStructure, lockWindows, lockRevision)
   */
  async protect(password?: string, options?: Partial<WorkbookProtection>): Promise<void> {
    this.protection = await buildWorkbookProtection(password, options);
  }

  /**
   * Remove workbook structure protection.
   */
  unprotect(): void {
    this.protection = undefined;
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

    if (!worksheet) {
      setTimeout(() => formatter.end(), 0);
      return formatter;
    }

    // Drive rows asynchronously so the formatter's backpressure signal can
    // throttle production. The drain wait races against `'error'` / `'close'`
    // so a downstream sink failure unwinds the producer instead of hanging.
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
        // Iterate worksheet rows lazily — no snapshot. Each row's `values`
        // are mapped at iteration time, so already-yielded data is GC-able
        // as the consumer drains the formatter.
        for (const { row, rowNumber } of iterateWorksheetRows(worksheet)) {
          if (formatter.destroyed) {
            return;
          }
          // First slot is the 1-based padding cell — skip without mutating
          // the row (mutating would corrupt subsequent reads / writes).
          const dataValues = row.values.slice(1).map(map);

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
      // When headers: true, CsvParserStream emits objects; otherwise arrays.
      // The CSV parser only ever emits primitive CellValue-compatible shapes
      // (string/number/boolean/Date/null). Narrow once here so the rest of
      // the pipeline — including the user-supplied `map` — sees CellValue.
      if (useHeaders && headerRow && row && typeof row === "object" && !Array.isArray(row)) {
        // Convert object row to array using header order
        const rowObj = row as Record<string, CellValue>;
        const rowArray: CellValue[] = headerRow.map(h => rowObj[h]);
        worksheet.addRow(rowArray.map(map));
      } else if (Array.isArray(row)) {
        worksheet.addRow((row as CellValue[]).map(map));
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
        // When headers: true, CsvParserStream emits objects; otherwise arrays.
        // See createCsvReadStream for the rationale on narrowing to CellValue.
        if (useHeaders && headerRow && row && typeof row === "object" && !Array.isArray(row)) {
          // Convert object row to array using header order
          const rowObj = row as Record<string, CellValue>;
          const rowArray: CellValue[] = headerRow.map(h => rowObj[h]);
          worksheet.addRow(rowArray.map(map));
        } else if (Array.isArray(row)) {
          worksheet.addRow((row as CellValue[]).map(map));
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

    // Race drain against error / close so a mid-stream sink failure makes
    // `writeAndDrain` reject (and the for-loop unwind) instead of hanging
    // on a 'drain' the destroyed formatter will never emit.
    const awaitFormatterDrain = createDrainRacer(formatter);

    const writeAndDrain = async (values: any[]): Promise<void> => {
      if (!formatter.write(values)) {
        await awaitFormatterDrain();
      }
    };

    try {
      let lastRow = 1;
      // Iterate worksheet rows directly without pre-collecting them. The
      // Workbook model is already in memory, so reaching into `_rows` here
      // adds no per-row allocation — just a Row reference per iteration,
      // immediately reassigned. The async loop honours formatter
      // backpressure between rows so the formatter's internal buffer can't
      // grow unbounded against a slow sink.
      for (const { row, rowNumber } of iterateWorksheetRows(worksheet)) {
        // First slot is the 1-based padding cell — skip without mutating
        // the row (mutating would corrupt subsequent reads / writes).
        const dataValues = row.values.slice(1).map(map);

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
      // Sink errored mid-write (or pipeline tore down for any reason).
      // Destroy the formatter so the pipeline unwinds, swallow the
      // pipeline rejection (the original error is what we want to surface),
      // and rethrow.
      formatter.destroy(err instanceof Error ? err : new Error(String(err)));
      await pipelinePromise.catch(() => {});
      throw err;
    }
  }

  /**
   * Populate a worksheet from a parsed Markdown table result.
   * Shared by readMarkdown and readMarkdownAll.
   */
  private _populateMarkdownWorksheet(
    worksheet: Worksheet,
    result: MarkdownParseResult,
    map?: (value: string, column: number) => unknown
  ): void {
    worksheet.addRow(result.headers);
    (worksheet as any)._markdownAlignments = result.alignments;
    for (const row of result.rows) {
      if (map) {
        worksheet.addRow(row.map((v, i) => map(v, i)));
      } else {
        worksheet.addRow(row);
      }
    }
  }

  /**
   * Read a Markdown table and add as worksheet.
   *
   * @example
   * ```ts
   * // From a Markdown string
   * workbook.readMarkdown("| Name | Age |\n| --- | --- |\n| Alice | 30 |");
   *
   * // With options
   * workbook.readMarkdown(markdownString, { sheetName: "Data", map: (v, col) => Number(v) || v });
   * ```
   */
  readMarkdown(input: string, options?: MarkdownOptions): Worksheet {
    const parseResult = parseMarkdown(input, {
      trim: options?.trim,
      unescape: options?.unescape,
      skipEmptyRows: options?.skipEmptyRows,
      maxRows: options?.maxRows,
      convertBr: options?.convertBr
    });

    const worksheet = this.addWorksheet(options?.sheetName);
    this._populateMarkdownWorksheet(worksheet, parseResult, options?.map);
    return worksheet;
  }

  /**
   * Read all Markdown tables from a document, each becoming a separate worksheet.
   *
   * @param input - Markdown string containing one or more tables
   * @param options - Parse options (sheetName is used as prefix: "sheetName", "sheetName_2", ...)
   * @returns Array of created worksheets (empty if no tables found)
   *
   * @example
   * ```ts
   * // Parse a document with multiple tables
   * const sheets = workbook.readMarkdownAll(markdownDoc);
   * console.log(`Created ${sheets.length} worksheets`);
   *
   * // With a naming prefix
   * const sheets = workbook.readMarkdownAll(markdownDoc, { sheetName: "Table" });
   * // Creates "Table", "Table_2", "Table_3", ...
   * ```
   */
  readMarkdownAll(input: string, options?: MarkdownOptions): Worksheet[] {
    const parseResults = parseMarkdownAll(input, {
      trim: options?.trim,
      unescape: options?.unescape,
      skipEmptyRows: options?.skipEmptyRows,
      maxRows: options?.maxRows,
      convertBr: options?.convertBr
    });

    const baseName = options?.sheetName;
    const map = options?.map;
    const worksheets: Worksheet[] = [];

    for (let t = 0; t < parseResults.length; t++) {
      const name = baseName ? (t === 0 ? baseName : `${baseName}_${t + 1}`) : undefined;
      const worksheet = this.addWorksheet(name);
      this._populateMarkdownWorksheet(worksheet, parseResults[t], map);
      worksheets.push(worksheet);
    }

    return worksheets;
  }

  /**
   * Write worksheet as a Markdown table string.
   *
   * @example
   * ```ts
   * // Write first worksheet
   * const markdownText = workbook.writeMarkdown();
   *
   * // Write specific worksheet with options
   * const markdownText = workbook.writeMarkdown({ sheetName: "Data", padding: true });
   * ```
   */
  writeMarkdown(options?: MarkdownOptions): string {
    const worksheet = this.getWorksheet(options?.sheetName || options?.sheetId);
    if (!worksheet) {
      return "";
    }

    const dateFormat = options?.dateFormat;
    const dateUTC = options?.dateUTC;
    const includeEmptyRows = options?.includeEmptyRows !== false;

    // Build stringify function
    const stringify = options?.stringify ?? createMarkdownStringify(dateFormat, dateUTC);

    // Collect all rows from worksheet
    const allRows: unknown[][] = [];
    let lastRow = 1;

    worksheet.eachRow((row: any, rowNumber: number) => {
      if (includeEmptyRows) {
        while (lastRow++ < rowNumber - 1) {
          allRows.push([]);
        }
      }
      // row.values is a 1-indexed sparse array — use Array.from to fill holes
      // with undefined, then slice(1) to remove the leading 1-indexed slot
      const values = Array.from(row.values as unknown[]).slice(1);
      allRows.push(values);
      lastRow = rowNumber;
    });

    if (allRows.length === 0) {
      return "";
    }

    // First row is the header
    const headerRow = allRows[0];
    const headers: string[] = headerRow.map(v => stringify(v));
    const dataRows = allRows.slice(1);

    // Check for stored alignments from a previous readMarkdown
    const storedAlignments: MarkdownAlignment[] | undefined = (worksheet as any)
      ._markdownAlignments;

    // Build column configs
    const columns = options?.columns;
    let resolvedColumns: { header: string; alignment?: MarkdownAlignment }[] | undefined;

    if (!columns && storedAlignments) {
      // Use stored alignments from parsed Markdown
      resolvedColumns = headers.map((h, i) => ({
        header: h,
        alignment: i < storedAlignments.length ? storedAlignments[i] : undefined
      }));
    }

    return formatMarkdown(headers, dataRows, {
      columns: resolvedColumns ?? columns,
      alignment: options?.alignment,
      padding: options?.padding,
      trailingNewline: options?.trailingNewline,
      escapeContent: options?.escapeContent,
      stringify
    });
  }

  /**
   * Write worksheet to Markdown buffer (Uint8Array).
   *
   * @example
   * ```ts
   * const buffer = workbook.writeMarkdownBuffer();
   * ```
   */
  writeMarkdownBuffer(options?: MarkdownOptions): Uint8Array {
    const markdownString = this.writeMarkdown(options);
    return new TextEncoder().encode(markdownString);
  }

  /**
   * Read Markdown from file (Node.js only - throws in browser)
   */
  async readMarkdownFile(_filename: string, _options?: MarkdownOptions): Promise<Worksheet> {
    throw new ExcelNotSupportedError(
      "readMarkdownFile()",
      "not available in browser. Use readMarkdown(string) instead."
    );
  }

  /**
   * Read all Markdown tables from file (Node.js only - throws in browser)
   */
  async readMarkdownAllFile(_filename: string, _options?: MarkdownOptions): Promise<Worksheet[]> {
    throw new ExcelNotSupportedError(
      "readMarkdownAllFile()",
      "not available in browser. Use readMarkdownAll(string) instead."
    );
  }

  /**
   * Write Markdown to file (Node.js only - throws in browser)
   */
  async writeMarkdownFile(_filename: string, _options?: MarkdownOptions): Promise<void> {
    throw new ExcelNotSupportedError(
      "writeMarkdownFile()",
      "not available in browser. Use writeMarkdown() and trigger a download instead."
    );
  }

  // ===========================================================================
  // Static Factory Methods for Streaming
  // ===========================================================================

  /**
   * Create a streaming workbook writer for large files.
   * This is more memory-efficient than using Workbook for large datasets.
   *
   * File-path output (`{ filename }`) is a Node.js-only feature and is
   * exposed by the Node `Workbook` subclass, which overrides this
   * factory to return the Node `WorkbookWriter`. This browser base
   * only accepts `{ stream }`.
   *
   * @example
   * ```ts
   * // Browser (or Node.js with an explicit stream)
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
   * File-path input (`string`) is a Node.js-only feature and is exposed
   * by the Node `Workbook` subclass, which overrides this factory to
   * return the Node `WorkbookReader`. This browser base accepts the
   * cross-platform `CommonInput` type
   * (`Uint8Array | ArrayBuffer | Readable | ReadableStream`).
   *
   * @example
   * ```ts
   * // Browser or Node.js with stream / buffer
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
  static createStreamReader(input: CommonInput, options?: WorkbookReaderOptions): WorkbookReader {
    return new WorkbookReader(input, options);
  }

  // ===========================================================================
  // Worksheet Management
  // ===========================================================================

  get nextId(): number {
    // Worksheets and chartsheets share a single `sheetId` namespace in
    // `workbook.xml`'s `<sheets>` element (OOXML requires each
    // `sheetId` to be globally unique across both families). Allocating
    // from `_worksheets` alone used to hand out an id already claimed
    // by a chartsheet whenever the author interleaved their calls —
    // e.g. `addWorksheet(×16)` → ids 1-16; `addChartsheet(×2)` → ids
    // 17-18 (via `_nextSheetId()`); then `addWorksheet("X")` walked
    // `_worksheets` slots 1..16, found them full, and returned
    // `_worksheets.length = 17`, colliding with the first chartsheet.
    // Excel rejects the resulting workbook as corrupt. Collect
    // chartsheet ids up front so the search honours the shared pool.
    const chartsheetIds = new Set<number>();
    for (const cs of this._chartsheets) {
      if (cs && typeof cs.id === "number" && Number.isFinite(cs.id)) {
        chartsheetIds.add(cs.id);
      }
    }
    // Prefer reusing vacated `_worksheets` slots (left as holes by
    // `removeWorksheetEx`) so ids stay stable across delete+add cycles.
    for (let i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i] && !chartsheetIds.has(i)) {
        return i;
      }
    }
    // No reusable hole — hand out the next id beyond the current
    // tail, skipping any slots already taken by chartsheets.
    let candidate = this._worksheets.length || 1;
    while (chartsheetIds.has(candidate)) {
      candidate++;
    }
    return candidate;
  }

  /**
   * Add a new worksheet and return a reference to it
   */
  addWorksheet(name?: string, options?: AddWorksheetOptions): Worksheet {
    const id = this.nextId;

    // Allocate `orderNo` from the unified worksheet+chartsheet counter.
    // Looking only at `_worksheets` here (the previous implementation)
    // silently collides when a chartsheet has been added in between:
    // e.g. `addWorksheet("A")` → orderNo 0; `addChartsheet(…)` → 1
    // (via `_nextSheetOrderNo()`); `addWorksheet("B")` → 1 again
    // (because `max(worksheets.orderNo) + 1 = 0 + 1 = 1`), so A and
    // B share an ordinal with the chartsheet. The writer's stable
    // sort then interleaves them non-deterministically, scrambling
    // the user's tab order (`[A, CS, B]` could come out as
    // `[A, B, CS]` or `[A, CS, B]` across runs).
    const orderNo = this._nextSheetOrderNo();
    const worksheetOptions = {
      ...options,
      id,
      name,
      orderNo,
      workbook: this as any
    };

    const worksheet = new Worksheet(worksheetOptions);

    this._worksheets[id] = worksheet;
    return worksheet;
  }

  removeWorksheetEx(worksheet: Worksheet): void {
    // Release any workbook-wide table names this sheet held so the names can
    // be reused by future tables on other sheets without spurious "name
    // already exists" errors.
    const tables = worksheet.tables;
    if (tables) {
      for (const tableName of Object.keys(tables)) {
        this._tableNames.delete(tableName.toLowerCase());
      }
    }
    this._worksheets[worksheet.id] = undefined!;
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
   * Add a chartsheet containing a single chart and return the created chartsheet.
   */
  addChartsheet(name: string | undefined, options: AddChartsheetOptions): Chartsheet {
    const sheetName = this._validateChartsheetName(name ?? `Chart${this._chartsheets.length + 1}`);
    const sheetNo = this._nextChartsheetNo();
    const id = this._nextSheetId();
    // Assign a unified `orderNo` across worksheets and chartsheets so
    // the writer can preserve the author's interleaved tab layout.
    // Without this, workbook-xform `prepare()` sorted by `sheetNo`
    // (file-path number, independent per family) and reordered
    // `[ws1, cs1, ws2]` into `[ws1, ws2, cs1]`.
    const orderNo = this._nextSheetOrderNo();
    const chartsheet: ChartsheetModel = {
      sheetNo,
      id,
      name: sheetName,
      orderNo,
      state: options.state ?? "visible",
      tabSelected: options.tabSelected,
      zoomScale: options.zoomScale,
      workbookViewId: options.workbookViewId,
      zoomToFit: options.zoomToFit,
      pageMargins: options.pageMargins,
      pageSetup: options.pageSetup,
      drawing: { rId: "rId1" }
    };

    if (isChartExOptions(options.chart)) {
      const chartSupport = getChartSupport();
      const chartExNumber = this.nextChartExNumber();
      const model = chartSupport.buildChartExModel(options.chart);
      try {
        chartSupport.fillChartExCaches(model, this as any);
      } catch {
        // Cache population is best-effort; never let it break chart creation.
      }
      this.addChartExStructuredEntry({ chartExNumber, model });
      chartsheet.chartExNumber = chartExNumber;
    } else {
      const chartSupport = getChartSupport();
      const chartNumber = this.nextChartNumber();
      const chartModel = isComboChartOptions(options.chart)
        ? chartSupport.buildComboChartModel(options.chart)
        : chartSupport.buildChartModel(options.chart);
      try {
        chartSupport.fillChartCaches(chartModel, this as any);
      } catch {
        // Cache population is best-effort; never let it break chart creation.
      }
      const entry: ChartEntry = { chartNumber, model: chartModel };
      // Resolve programmatic `series.spPr.fill.blip._pendingImage`
      // payloads into workbook media entries and chart rels. The
      // worksheet-embedded `addChart` path does this immediately
      // after `fillChartCaches`; chartsheets ran the same builder
      // output but skipped the image-resolution step entirely, so a
      // picture-fill series authored via `addChartsheet` was
      // registered with its `_pendingImage` stuck on the model and
      // never reached `media/imageN.{ext}` — Excel rendered the
      // series as a transparent fill. Safe to call before
      // `addChartEntry` so the stored entry carries its resolved
      // `entry.rels` from the start.
      try {
        chartSupport.resolvePendingChartImages(entry, this as any, chartNumber);
      } catch {
        // Image resolution is best-effort; a broken image payload
        // should never take down chart creation — the series keeps
        // its `pictureOptions`, just without the blipFill.
      }
      this.addChartEntry(entry);
      this._applyChartsheetSidecars(chartNumber, options.chart);
      chartsheet.chartNumber = chartNumber;
    }

    this._chartsheets.push(chartsheet);
    return new Chartsheet(chartsheet, this);
  }

  /**
   * Add a chartsheet containing a classic pivot chart linked to an existing pivot table.
   */
  addPivotChartsheet(
    name: string | undefined,
    pivotTable: PivotTable,
    options: AddPivotChartsheetOptions
  ): Chartsheet {
    return this.addChartsheet(name, {
      ...options,
      chart: withPivotChartSource(pivotTable, options.chart)
    });
  }

  /** Return chartsheets in workbook order. */
  get chartsheets(): Chartsheet[] {
    return this._chartsheets.map(model => new Chartsheet(model, this));
  }

  getChartsheet(nameOrIndex: string | number): Chartsheet | undefined {
    const model = this._getChartsheetModel(nameOrIndex);
    return model ? new Chartsheet(model, this) : undefined;
  }

  removeChartsheet(nameOrIndex: string | number): boolean {
    const index =
      typeof nameOrIndex === "number"
        ? nameOrIndex
        : this._chartsheets.findIndex(
            sheet => sheet.name.toLowerCase() === nameOrIndex.toLowerCase()
          );
    if (index < 0 || index >= this._chartsheets.length) {
      return false;
    }
    const [removed] = this._chartsheets.splice(index, 1);
    if (removed.chartNumber) {
      this.removeChartEntry?.(removed.chartNumber);
    }
    if (removed.chartExNumber) {
      this.removeChartExStructuredEntry?.(removed.chartExNumber);
    }
    return true;
  }

  private _getChartsheetModel(nameOrIndex: string | number): ChartsheetModel | undefined {
    return typeof nameOrIndex === "number"
      ? this._chartsheets[nameOrIndex]
      : this._chartsheets.find(sheet => sheet.name.toLowerCase() === nameOrIndex.toLowerCase());
  }

  renameChartsheet(nameOrIndex: string | number, name: string): boolean {
    const model = this._getChartsheetModel(nameOrIndex);
    if (!model) {
      return false;
    }
    const currentName = model.name;
    if (currentName === name) {
      return true;
    }
    model.name = "__excelts_pending_chartsheet_rename__";
    try {
      model.name = this._validateChartsheetName(name);
      return true;
    } catch (error) {
      model.name = currentName;
      throw error;
    }
  }

  copyChartsheet(nameOrIndex: string | number, name?: string): Chartsheet | undefined {
    const source = this._getChartsheetModel(nameOrIndex);
    if (!source) {
      return undefined;
    }
    const cloneName = this._validateChartsheetName(name ?? `${source.name} Copy`);
    const clone: ChartsheetModel = {
      ...deepClone(source),
      id: this._nextSheetId(),
      sheetNo: this._nextChartsheetNo(),
      // New tab position — the clone goes to the tail of the tab
      // bar, matching Excel's "Duplicate" behaviour. Drop the
      // deep-cloned `orderNo` from the source.
      orderNo: this._nextSheetOrderNo(),
      name: cloneName,
      drawingName: undefined,
      relationships: source.relationships ? deepClone(source.relationships) : undefined
    };
    if (source.chartNumber) {
      const entry = this.getChartEntry(source.chartNumber);
      if (entry) {
        const chartNumber = this.nextChartNumber();
        // Clone the entry with ALL metadata: rawData, modelSnapshot,
        // dirty, preferRawPatch, requireRawPatch, rels (per-entry),
        // userShapesXml. A freshly-created entry carrying only `model`
        // would lose Excel-authored user-shape overlays, the raw-patch
        // fast path, and any per-entry rels that aren't in
        // `_chartRels`. Keeping them in lockstep means a clone of a
        // just-loaded chart matches the source byte-for-byte.
        this.addChartEntry(cloneChartEntry(entry, chartNumber));
        this.copyChartSidecars(source.chartNumber, chartNumber);
        clone.chartNumber = chartNumber;
        clone.chartExNumber = undefined;
      }
    } else if (source.chartExNumber) {
      const entry = this.getChartExStructuredEntry(source.chartExNumber);
      const chartExNumber = this.nextChartExNumber();
      if (entry) {
        // Same rationale as the classic branch — carry dirty /
        // preferRawPatch / requireRawPatch / rawData / modelSnapshot
        // across the clone so the raw-patch path keeps working on
        // the duplicate.
        this.addChartExStructuredEntry(cloneChartExEntry(entry, chartExNumber));
      } else if (this._chartExEntries[source.chartExNumber]) {
        this._chartExEntries[chartExNumber] = this._chartExEntries[source.chartExNumber].slice();
      }
      // Copy the chartEx sidecars (authored rels) so the cloned
      // chartsheet's XML references stay valid. Previously a chartEx
      // with `cx14:` / media rels on the source lost every relationship
      // on the clone.
      this.copyChartExSidecars(source.chartExNumber, chartExNumber);
      clone.chartExNumber = chartExNumber;
      clone.chartNumber = undefined;
    }
    this._chartsheets.push(clone);
    return new Chartsheet(clone, this);
  }

  replaceChartsheetChart(
    nameOrIndex: string | number,
    chart: AddChartsheetOptions["chart"]
  ): boolean {
    const wrapper = this.getChartsheet(nameOrIndex);
    if (!wrapper) {
      return false;
    }
    const model = wrapper.model;
    const chartSupport = getChartSupport();
    // Build the replacement first so a malformed options object throws
    // *before* we remove the existing chart entry. Without this, a
    // failed `buildChartExModel` / `buildChartModel` would leave the
    // chartsheet empty (old chart nuked, new chart never registered).
    let newChartExModel: ChartExModel | undefined;
    let newChartModel: ChartModel | undefined;
    if (isChartExOptions(chart)) {
      newChartExModel = chartSupport.buildChartExModel(chart);
    } else if (isComboChartOptions(chart)) {
      newChartModel = chartSupport.buildComboChartModel(chart);
    } else {
      newChartModel = chartSupport.buildChartModel(chart);
    }
    // Remove existing entries only after the new model builds cleanly.
    if (model.chartNumber) {
      this.removeChartEntry?.(model.chartNumber);
      model.chartNumber = undefined;
    }
    if (model.chartExNumber) {
      this.removeChartExStructuredEntry?.(model.chartExNumber);
      model.chartExNumber = undefined;
    }
    if (newChartExModel) {
      const chartExNumber = this.nextChartExNumber();
      try {
        chartSupport.fillChartExCaches(newChartExModel, this as any);
      } catch {
        // Cache population is best-effort; never let it break chart replacement.
      }
      this.addChartExStructuredEntry({ chartExNumber, model: newChartExModel });
      model.chartExNumber = chartExNumber;
    } else if (newChartModel) {
      const chartNumber = this.nextChartNumber();
      try {
        chartSupport.fillChartCaches(newChartModel, this as any);
      } catch {
        // Cache population is best-effort; never let it break chart replacement.
      }
      const entry: ChartEntry = { chartNumber, model: newChartModel };
      // Resolve programmatic `series.spPr.fill.blip._pendingImage`
      // payloads — matches the classic `addChart` and `addChartsheet`
      // paths. Previously replacement via `replaceChartsheetChart`
      // silently dropped picture-fill payloads on the floor.
      try {
        chartSupport.resolvePendingChartImages(entry, this as any, chartNumber);
      } catch {
        // Image resolution is best-effort; a broken image payload
        // should never take down chart replacement.
      }
      this.addChartEntry(entry);
      this._applyChartsheetSidecars(chartNumber, chart);
      model.chartNumber = chartNumber;
    }
    return true;
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

  /**
   * Workbook-level directory of people referenced by threaded comments.
   * Mutating the returned array adds/removes entries in the persistent
   * state; writers emit `xl/persons/person.xml` only when this list is
   * non-empty.
   *
   * Most callers don't need to touch this directly — creating a
   * {@link ThreadedComment} through `cell.note` handles registration
   * automatically.
   */
  get persons(): ThreadedCommentPerson[] {
    return this._persons;
  }

  /**
   * Register a person in the workbook persons list and return its id.
   *
   * When an entry with the same {@link displayName} + {@link userId}
   * already exists, its existing id is returned so duplicate
   * commenters collapse onto a single entry. New entries receive a
   * synthesised `{GUID}` id.
   *
   * @param displayName — shown in the comment bubble author line
   * @param userId — optional identity-provider user id (email / SID)
   * @param providerId — optional provider identifier ("AD", …)
   */
  registerPerson(displayName: string, userId?: string, providerId?: string): string {
    const existing = this._persons.find(p => p.displayName === displayName && p.userId === userId);
    if (existing) {
      return existing.id;
    }
    const id = `{${synthGuid()}}`;
    const entry: ThreadedCommentPerson = { id, displayName };
    if (userId !== undefined) {
      entry.userId = userId;
    }
    if (providerId !== undefined) {
      entry.providerId = providerId;
    }
    this._persons.push(entry);
    return id;
  }

  // ===========================================================================
  // Formula Calculation
  // ===========================================================================

  /**
   * Recalculate all formula cells in this workbook.
   *
   * Evaluates every formula cell using the built-in calculation engine and updates
   * each cell's cached `result` value in-place. Formulas are evaluated with
   * recursive dependency resolution, memoization, and circular reference detection.
   *
   * Call this after programmatically modifying cell values that are referenced by
   * formulas, to ensure formula results reflect the latest data.
   *
   * Unsupported functions preserve their original cached result if one exists.
   *
   * ## Tree-shaking note
   *
   * The formula engine ships ~200KB of code (433 Excel functions, parser,
   * evaluator, dependency graph, spill materialiser). To keep it out of
   * bundles that don't need it, the engine is registered at runtime
   * rather than imported by the core `Workbook` module. Call
   * {@link installFormulaEngine} once at startup before the first call
   * to this method, or a clear error will be thrown explaining what to do.
   *
   * ```ts
   * import { installFormulaEngine } from "@cj-tech-master/excelts/formula";
   *
   * installFormulaEngine();                   // once, at startup
   *
   * sheet.getCell("A1").value = 100;
   * workbook.calculateFormulas();             // now works
   * ```
   *
   * Callers who prefer a zero-side-effect, tree-shakeable surface can
   * use the functional equivalent directly:
   *
   * ```ts
   * import { calculateFormulas } from "@cj-tech-master/excelts/formula";
   * calculateFormulas(workbook);
   * ```
   */
  calculateFormulas(): void {
    invokeFormulaEngine(this);
  }

  // ===========================================================================
  // User-registered formula functions
  // ===========================================================================

  /**
   * Per-workbook registry of user-defined functions. The formula engine
   * consults this map before the built-in 433-function registry, so a
   * registered name either adds a new function (`MYFN`) or shadows a
   * built-in (`IRR` → project-specific variant).
   *
   * Populated by {@link registerFunction}; read by the formula engine
   * when the host calls `calculateFormulas()` — see
   * `@formula/runtime/evaluator.ts::evaluateCall`.
   */
  userFunctions?: Map<
    string,
    {
      minArity: number;
      maxArity: number;
      invoke: (args: unknown[]) => unknown;
      volatile?: boolean;
    }
  >;

  /**
   * Register (or replace) a custom formula function on this workbook.
   *
   * The function becomes visible to `calculateFormulas()` on this
   * workbook only — the built-in registry stays untouched. Names are
   * case-insensitive (normalised to uppercase) and must not include
   * the `_XLFN.` prefix — the engine strips that automatically.
   *
   * @param name    Function name (case-insensitive).
   * @param fn      Implementation. Receives already-evaluated RuntimeValue
   *                arguments; return a RuntimeValue. Wrap failures with
   *                `rvError("#VALUE!")` rather than throwing — throws are
   *                caught at the evaluator boundary and surface as
   *                `#VALUE!` so a buggy custom function doesn't tear
   *                down the whole calculation pass.
   * @param options Optional arity bounds. Defaults to `minArity=0`,
   *                `maxArity=255` (Excel's universal argument cap), so
   *                simple variadic functions work without extra config.
   *                Set `volatile: true` when the function should be
   *                re-evaluated on every calc cycle (analogous to
   *                built-in `RAND`, `NOW`). Currently reserved for
   *                future use; the engine recomputes every formula on
   *                each `calculateFormulas()` call regardless.
   *
   * ```ts
   * import { rvNumber } from "@cj-tech-master/excelts/formula";
   * workbook.registerFunction("DOUBLE", ([x]) => {
   *   return rvNumber((x as any).value * 2);
   * }, { minArity: 1, maxArity: 1 });
   * ```
   */
  registerFunction(
    name: string,
    fn: (args: unknown[]) => unknown,
    options?: { minArity?: number; maxArity?: number; volatile?: boolean }
  ): void {
    if (!this.userFunctions) {
      this.userFunctions = new Map();
    }
    this.userFunctions.set(name.toUpperCase(), {
      minArity: options?.minArity ?? 0,
      maxArity: options?.maxArity ?? 255,
      invoke: fn,
      volatile: options?.volatile ?? false
    });
  }

  /**
   * Remove a user-registered function. No-op when the name isn't
   * registered; returns `true` when an entry was removed.
   */
  unregisterFunction(name: string): boolean {
    return this.userFunctions?.delete(name.toUpperCase()) ?? false;
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
   * Register an image with the workbook and return its numeric id. Pass the id
   * to {@link Worksheet.addImage}, {@link Worksheet.addBackgroundImage}, or
   * {@link Worksheet.addWatermark} to place it.
   *
   * The image is either **embedded** or **linked (external)**:
   *
   * - **Embedded** — supply `buffer`, `base64`, or `filename`. The bytes are
   *   written into the `.xlsx` package (`xl/media/imageN.ext`). Self-contained,
   *   but inflates file size.
   * - **Linked (external)** — supply only `link` (a URL or local file path).
   *   No bytes are stored; the package keeps a relationship with
   *   `TargetMode="External"` and the picture is rendered via `<a:blip r:link>`.
   *   Keeps the file small, but the image is resolved by Excel at open time.
   *
   * If both bytes and a `link` are provided, **embedding wins**.
   *
   * Linked images work with **cell pictures** ({@link Worksheet.addImage}) and
   * **overlay watermarks** ({@link Worksheet.addWatermark} with `mode:
   * "overlay"`). Worksheet background images and header/footer (VML) watermarks
   * cannot be linked — they require an embedded image.
   *
   * Note: Excel treats linked images as volatile — a moved/missing target
   * shows a broken-image placeholder, and modern Excel may not auto-load
   * remote URLs for security reasons. Prefer embedding for self-contained files.
   *
   * @example Embedded image
   * ```typescript
   * const id = workbook.addImage({ buffer: pngBytes, extension: "png" });
   * worksheet.addImage(id, "B2:D6");
   * ```
   *
   * @example Linked (external) image — no bytes stored
   * ```typescript
   * const id = workbook.addImage({ extension: "png", link: "https://example.com/logo.png" });
   * worksheet.addImage(id, "B2:D6");
   * ```
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
  // Charts
  // ===========================================================================

  /**
   * Return the next available 1-based chart number.
   */
  nextChartNumber(): number {
    const existing = Object.keys(this._chartEntries).map(Number);
    return existing.length > 0 ? Math.max(...existing) + 1 : 1;
  }

  /**
   * Store a chart entry in the workbook (keyed by chartNumber).
   */
  addChartEntry(entry: ChartEntry): void {
    this._chartEntries[entry.chartNumber] = entry;
  }

  setChartStyle(chartNumber: number, data: Uint8Array): void {
    this._chartStyles[chartNumber] = data;
  }

  setChartColors(chartNumber: number, data: Uint8Array): void {
    this._chartColors[chartNumber] = data;
  }

  copyChartSidecars(
    sourceChartNumber: number,
    targetChartNumber: number,
    targetWorkbook: Workbook = this
  ): void {
    const style = this._chartStyles[sourceChartNumber];
    if (style) {
      targetWorkbook.setChartStyle(targetChartNumber, style.slice());
    }
    const colors = this._chartColors[sourceChartNumber];
    if (colors) {
      targetWorkbook.setChartColors(targetChartNumber, colors.slice());
    }
    // Copy the full chart rels bag (`_chartRels`), not just the
    // style/colors pair. A classic chart can carry rels to embedded
    // images (pictureFill), external data links, and `<c:userShapes>`
    // drawing parts — without copying those the clone ends up with
    // dangling rIds. Deep-copy each rel so a later mutation on the
    // source doesn't leak into the clone.
    //
    // Rewrite style/colors Targets to the destination chart number —
    // verbatim copy would leave the rel pointing at the source's
    // `style{src}.xml`, while the writer emits `style{dst}.xml` and
    // produces a chart whose .rels references a non-existent file.
    //
    // For image rels on a cross-workbook copy (`targetWorkbook !==
    // this`), re-register each referenced image in the destination
    // workbook and rewrite the Target to point at the new media
    // file. Without this, a pictureFill that round-tripped through
    // `importSheet` pointed at the source workbook's media array —
    // which the destination package doesn't ship, so Excel shows a
    // broken image icon.
    const srcRels = this._chartRels[sourceChartNumber];
    if (Array.isArray(srcRels) && srcRels.length > 0) {
      const crossWorkbook = targetWorkbook !== this;
      targetWorkbook._chartRels[targetChartNumber] = srcRels.map(rel => {
        if (typeof rel !== "object" || rel === null) {
          return rel;
        }
        const cloned = { ...rel } as { Type?: string; Target?: string; [k: string]: unknown };
        const target = typeof cloned.Target === "string" ? cloned.Target : undefined;
        if (target) {
          if (/^style\d+\.xml$/.test(target)) {
            cloned.Target = `style${targetChartNumber}.xml`;
          } else if (/^colors\d+\.xml$/.test(target)) {
            cloned.Target = `colors${targetChartNumber}.xml`;
          } else if (crossWorkbook && cloned.Type === RelType.Image) {
            const rewritten = this._rewriteCrossWorkbookImageTarget(target, targetWorkbook);
            if (rewritten !== undefined) {
              cloned.Target = rewritten;
            }
          }
        }
        return cloned;
      });
    }
  }

  /**
   * Copy the media referenced by `target` (e.g. `../media/image3.png`)
   * from this workbook's media collection into `targetWorkbook`, then
   * return the rewritten Target pointing at the destination workbook's
   * copy. Returns `undefined` when the source media can't be
   * resolved — callers leave the original target in place and let
   * the writer emit a broken rel (same degradation as before the
   * cross-workbook rewrite landed).
   *
   * The on-disk naming convention is determined by the destination
   * workbook's `addImage`, so we take the `id` the new image gets and
   * compute `../media/image{id+1}.{ext}` (the same formula used in
   * `chart-images.ts:resolvePendingChartImages`). Centralising the
   * mapping here keeps the cross-workbook copy robust against future
   * media-naming changes on either workbook.
   */
  private _rewriteCrossWorkbookImageTarget(
    target: string,
    targetWorkbook: Workbook
  ): string | undefined {
    const match = /\/media\/image(\d+)\.([a-zA-Z0-9]+)$/.exec(target);
    if (!match) {
      return undefined;
    }
    const sourceMediaIndex = parseInt(match[1], 10) - 1;
    if (!Number.isFinite(sourceMediaIndex) || sourceMediaIndex < 0) {
      return undefined;
    }
    const medium = this.getImage?.(sourceMediaIndex) as
      | { extension?: string; buffer?: Uint8Array; base64?: string }
      | undefined;
    if (!medium) {
      return undefined;
    }
    const ext = medium.extension as "png" | "jpeg" | "gif" | undefined;
    if (ext !== "png" && ext !== "jpeg" && ext !== "gif") {
      return undefined;
    }
    const payload: { extension: "png" | "jpeg" | "gif"; buffer?: Uint8Array; base64?: string } = {
      extension: ext
    };
    // `instanceof Uint8Array` is realm-sensitive: buffers that crossed
    // a Worker / iframe / `structuredClone` boundary carry a different
    // `Uint8Array` prototype and fail the operator even though they
    // are byte-granular typed arrays. Duck-type via `ArrayBuffer.isView`
    // + `BYTES_PER_ELEMENT === 1` so cross-workbook copies from a
    // worker-loaded Workbook preserve the image bytes; otherwise the
    // copy path silently falls through to `return undefined`, dropping
    // every image from the chart. Matches `chart-images.ts`'s handling
    // of the same realm-crossing issue.
    const buf = medium.buffer as ArrayBufferView | undefined;
    if (
      buf &&
      ArrayBuffer.isView(buf) &&
      (buf as unknown as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1
    ) {
      payload.buffer =
        buf instanceof Uint8Array
          ? buf.slice()
          : new Uint8Array(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
    } else if (typeof medium.base64 === "string") {
      payload.base64 = medium.base64;
    } else {
      return undefined;
    }
    const newId = targetWorkbook.addImage(payload);
    return `../media/image${newId + 1}.${ext}`;
  }

  /**
   * Copy the ChartEx-specific sidecar state from one `chartExNumber` slot
   * to another. Classic charts have `copyChartSidecars` for chart-style /
   * chart-colors; ChartEx charts carry their own `_chartExRels`
   * (relationship entries for extension packages, embedded images, etc.)
   * that the classic helper does not touch. Without this, cloning a
   * ChartEx via {@link Chart.copyTo} / {@link Chartsheet.clone} silently
   * dropped every relationship the source had — authored `cx14:`
   * extensions, embedded custom geometry, or linked media ended up
   * pointing at an undefined rel id on the clone.
   */
  copyChartExSidecars(
    sourceChartExNumber: number,
    targetChartExNumber: number,
    targetWorkbook: Workbook = this
  ): void {
    const rels = this._chartExRels[sourceChartExNumber];
    if (rels && rels.length > 0) {
      // Rewrite `Target` for rels that point at numbered sidecars
      // (styleEx / colorsEx / userShapes). Those files get different
      // on-disk numbers on the clone — copying the rel verbatim
      // leaves it pointing at the source's sidecar, so saving the
      // package produces a chartEx whose .rels references
      // `styleEx{src}.xml` while the writer emits `styleEx{dst}.xml`.
      // Strip the number from the source Target and re-stamp it with
      // the target's number.
      //
      // For image rels on a cross-workbook copy (`targetWorkbook !==
      // this`), re-register each referenced image in the destination
      // workbook and rewrite the Target — same logic as classic chart
      // sidecars. Without this, a ChartEx with embedded images (e.g.
      // pictureFill or custom geometry) would reference media that
      // doesn't exist in the destination package.
      const crossWorkbook = targetWorkbook !== this;
      targetWorkbook._chartExRels[targetChartExNumber] = rels.map(r => {
        if (typeof r !== "object" || r === null) {
          return r;
        }
        const cloned = { ...r };
        const target: string | undefined =
          typeof cloned.Target === "string" ? cloned.Target : undefined;
        if (target) {
          const styleExMatch = /^styleEx\d+\.xml$/.exec(target);
          if (styleExMatch) {
            cloned.Target = `styleEx${targetChartExNumber}.xml`;
          } else if (/^colorsEx\d+\.xml$/.exec(target)) {
            cloned.Target = `colorsEx${targetChartExNumber}.xml`;
          } else if (crossWorkbook && cloned.Type === RelType.Image) {
            const rewritten = this._rewriteCrossWorkbookImageTarget(target, targetWorkbook);
            if (rewritten !== undefined) {
              cloned.Target = rewritten;
            }
          }
        }
        return cloned;
      });
    }
    // ChartEx style / colors sidecars (matching `_chartStyles` /
    // `_chartColors` for classic charts). Previously only `_chartExRels`
    // was copied — a cloned chartEx lost its authored chartExStyle and
    // chartExColors bytes, so the saved package re-derived them from
    // defaults and the clone looked different from the source.
    const exStyle = this._chartExStyles[sourceChartExNumber];
    if (exStyle) {
      targetWorkbook._chartExStyles[targetChartExNumber] = exStyle.slice();
    }
    const exColors = this._chartExColors[sourceChartExNumber];
    if (exColors) {
      targetWorkbook._chartExColors[targetChartExNumber] = exColors.slice();
    }
  }

  private _applyChartsheetSidecars(
    chartNumber: number,
    chartOptions: AddChartsheetOptions["chart"]
  ): void {
    if (isChartExOptions(chartOptions)) {
      return;
    }
    if (!chartOptions.chartStyle && !chartOptions.chartColors) {
      return;
    }
    const chartSupport = getChartSupport();
    if (chartOptions.chartStyle) {
      this.setChartStyle(
        chartNumber,
        new TextEncoder().encode(chartSupport.buildChartStyle(chartOptions.chartStyle))
      );
    }
    if (chartOptions.chartColors) {
      this.setChartColors(
        chartNumber,
        new TextEncoder().encode(chartSupport.buildChartColors(chartOptions.chartColors))
      );
    }
  }

  /**
   * Retrieve a chart entry by its 1-based chart number.
   */
  getChartEntry(chartNumber: number): ChartEntry | undefined {
    return this._chartEntries[chartNumber];
  }

  /**
   * Remove a chart entry from the workbook.
   * Safe to call even if the chart number doesn't exist.
   */
  removeChartEntry(chartNumber: number): void {
    delete this._chartEntries[chartNumber];
    delete this._chartRels[chartNumber];
    delete this._chartStyles[chartNumber];
    delete this._chartColors[chartNumber];
  }

  // ===========================================================================
  // ChartEx (Office 2016+) structured entries
  // ===========================================================================

  /** Return the next available 1-based chartEx number. */
  nextChartExNumber(): number {
    const rawKeys = Object.keys(this._chartExEntries ?? {}).map(Number);
    const structKeys = Object.keys(this._chartExStructuredEntries ?? {}).map(Number);
    const combined = [...rawKeys, ...structKeys];
    return combined.length > 0 ? Math.max(...combined) + 1 : 1;
  }

  /**
   * Store a structured chartEx entry.
   * Loaded entries may also keep raw bytes for clean passthrough.
   */
  addChartExStructuredEntry(entry: ChartExEntry): void {
    if (!this._chartExStructuredEntries) {
      this._chartExStructuredEntries = {};
    }
    this._chartExStructuredEntries[entry.chartExNumber] = entry;
  }

  /** Get a structured chartEx entry by number. */
  getChartExStructuredEntry(chartExNumber: number): ChartExEntry | undefined {
    return this._chartExStructuredEntries?.[chartExNumber];
  }

  private _nextChartsheetNo(): number {
    const existing = this._chartsheets.map(cs => cs.sheetNo).filter(Number.isFinite);
    return existing.length > 0 ? Math.max(...existing) + 1 : 1;
  }

  /**
   * Next value for the unified `orderNo` (tab-bar position) counter
   * shared between worksheets and chartsheets. Used by the writer's
   * `prepare()` to emit `<sheets>` in the author's insertion order,
   * preserving interleaved `[ws, cs, ws]` layouts that the old
   * sheetNo-based sort used to reshuffle.
   */
  private _nextSheetOrderNo(): number {
    let max = -1;
    for (const ws of this._worksheets) {
      if (ws && typeof ws.orderNo === "number" && ws.orderNo > max) {
        max = ws.orderNo;
      }
    }
    for (const cs of this._chartsheets) {
      if (typeof cs.orderNo === "number" && cs.orderNo > max) {
        max = cs.orderNo;
      }
    }
    return max + 1;
  }

  private _nextSheetId(): number {
    const worksheetIds = this.worksheets.map(ws => ws.id);
    const chartsheetIds = this._chartsheets.map(cs => cs.id).filter(Number.isFinite);
    const ids = [...worksheetIds, ...chartsheetIds];
    return ids.length > 0 ? Math.max(...ids) + 1 : 1;
  }

  /**
   * Validate a sheet name (worksheet OR chartsheet) against Excel's
   * single unified namespace. Returns the (possibly truncated) name on
   * success; throws {@link WorksheetNameError} on invalid input.
   *
   * Unifying the check at the workbook level fixes three related
   * regressions that used to exist in the per-family validators:
   *   1. `Worksheet.name` setter only cross-checked against other
   *      worksheets, so `addChartsheet("S")` followed by
   *      `addWorksheet("S")` silently produced a duplicate tab name.
   *   2. The chartsheet regex was missing the backslash, so
   *      `addChartsheet("A\\B")` sneaked through — Excel rejects it.
   *   3. `Chartsheet.name = …` bypassed validation entirely, letting
   *      users mutate the model into a corrupt state.
   *
   * @param name - Proposed sheet name. `undefined` / empty / over-31
   *   chars / containing any of `* ? : \\ / [ ]` / leading or trailing
   *   single-quote is rejected. Names ≤31 chars are passed through;
   *   longer ones are truncated (non-production builds emit a warning).
   * @param existing - The sheet being renamed (if any) — it is
   *   excluded from the duplicate check so `sheet.name = sheet.name`
   *   is a no-op rather than a self-collision.
   */
  validateSheetName(name: string, existing?: Worksheet | { name: string }): string {
    if (typeof name !== "string") {
      throw new WorksheetNameError("The name has to be a string.");
    }
    if (name === "") {
      throw new WorksheetNameError("The name can't be empty.");
    }
    if (name === "History") {
      throw new WorksheetNameError('The name "History" is protected. Please use a different name.');
    }
    // Illegal characters per Excel's own naming rules: asterisk (*),
    // question mark (?), colon (:), forward slash (/), backslash (\),
    // left bracket ([), right bracket (]). The chartsheet regex used
    // to omit `\\`; unified here so both families enforce the same
    // char set.
    if (/[*?:/\\[\]]/.test(name)) {
      throw new WorksheetNameError(
        `Sheet name ${name} cannot include any of the following characters: * ? : \\ / [ ]`
      );
    }
    if (/(^')|('$)/.test(name)) {
      throw new WorksheetNameError(
        `The first or last character of sheet name cannot be a single quotation mark: ${name}`
      );
    }
    if (name.length > 31) {
      if (process.env.NODE_ENV !== "production") {
        console.warn(`Sheet name ${name} exceeds 31 chars. This will be truncated`);
      }
      name = name.substring(0, 31);
    }
    const nameLower = name.toLowerCase();
    const duplicateWorksheet = this.worksheets.find(
      ws => ws && ws !== existing && ws.name.toLowerCase() === nameLower
    );
    const duplicateChartsheet = this._chartsheets.find(
      cs => cs && cs !== existing && cs.name.toLowerCase() === nameLower
    );
    if (duplicateWorksheet || duplicateChartsheet) {
      throw new WorksheetNameError(`Sheet name already exists: ${name}`);
    }
    return name;
  }

  private _validateChartsheetName(name: string): string {
    return this.validateSheetName(name);
  }
  /** Remove a structured chartEx entry. */
  removeChartExStructuredEntry(chartExNumber: number): void {
    if (this._chartExStructuredEntries) {
      delete this._chartExStructuredEntries[chartExNumber];
    }
    delete this._chartExEntries[chartExNumber];
    delete this._chartExRels[chartExNumber];
    delete this._chartExStyles[chartExNumber];
    delete this._chartExColors[chartExNumber];
  }

  // ===========================================================================
  // External Workbook Links
  // ===========================================================================

  /**
   * Declare that formulas in this workbook may reference an external
   * workbook. Registers the target so the output file contains the required
   * `xl/externalLinks/externalLink{N}.xml` part plus its `.rels` sibling
   * and Office/WPS can resolve the reference correctly.
   *
   * When Office opens the file, it resolves a relative `target` like
   * `"测试.xlsx"` **relative to the current workbook's directory** — which
   * is the exact behaviour the user expects when they write
   * `=[测试.xlsx]Sheet1!A1`. Absolute `file:///…` or `http(s)://…` URIs
   * are accepted and written through unchanged.
   *
   * @returns the registered {@link ExternalLinkModel}. Its `index` field is
   *   the 1-based number used inside the `[N]` prefix of on-disk formula
   *   strings (the library rewrites `[target]` forms to `[index]` at write
   *   time automatically).
   *
   * @example
   * ```ts
   * const wb = new Workbook();
   * const ws = wb.addWorksheet("Main");
   *
   * // Declare the link once — sheet names and cached values are optional
   * // but improve interoperability (Excel displays cached values when the
   * // external file is unavailable).
   * wb.addExternalLink({
   *   target: "测试.xlsx",
   *   sheetNames: ["Sheet1"],
   *   cachedValues: { Sheet1: { A1: 42 } }
   * });
   *
   * // Write the formula using either the target name OR the numeric index;
   * // the library normalises both to the on-disk `[N]` form.
   * ws.getCell("A1").value = { formula: "=[测试.xlsx]Sheet1!A1", result: 42 };
   * ```
   */
  addExternalLink(input: {
    target: string;
    sheetNames?: string[];
    cachedValues?: ExternalLinkModel["cachedValues"];
    targetMode?: ExternalLinkModel["targetMode"];
  }): ExternalLinkModel {
    const link: ExternalLinkModel = {
      index: this.externalLinks.length + 1,
      target: input.target,
      targetMode: input.targetMode ?? "External",
      sheetNames: input.sheetNames ? [...input.sheetNames] : [],
      cachedValues: input.cachedValues ? { ...input.cachedValues } : {}
    };
    this.externalLinks.push(link);
    return link;
  }

  /**
   * Retrieve an external link by its 1-based on-disk index (the number
   * inside the `[N]` formula prefix) or by matching target path.
   */
  getExternalLink(indexOrTarget: number | string): ExternalLinkModel | undefined {
    if (typeof indexOrTarget === "number") {
      return this.externalLinks[indexOrTarget - 1];
    }
    const lower = indexOrTarget.toLowerCase();
    return this.externalLinks.find(link => link.target.toLowerCase() === lower);
  }

  /**
   * @internal — used by the writer to obtain the full list of external
   * links to serialise, including entries auto-discovered from formula
   * strings during earlier writes. User-visible `externalLinks` always
   * comes first (in declaration order) so explicit `addExternalLink()`
   * indices are stable across writes.
   */
  _collectExternalLinksForWrite(): ExternalLinkModel[] {
    const userLower = new Set(this.externalLinks.map(l => l.target.toLowerCase()));
    const combined: ExternalLinkModel[] = this.externalLinks.map((link, i) => ({
      ...link,
      index: i + 1,
      sheetNames: [...(link.sheetNames ?? [])],
      cachedValues: { ...(link.cachedValues ?? {}) },
      targetMode: link.targetMode ?? "External"
    }));
    for (const cached of this._writerExternalLinkCache.values()) {
      if (userLower.has(cached.target.toLowerCase())) {
        // User explicitly added a link with the same target after an
        // auto-discovery pass — prefer the user's definition, drop the
        // cached one.
        continue;
      }
      combined.push({
        ...cached,
        index: combined.length + 1,
        sheetNames: [...cached.sheetNames],
        cachedValues: { ...cached.cachedValues }
      });
    }
    return combined;
  }

  /**
   * @internal — record an auto-discovered external link (seen in a
   * formula but not explicitly declared). Idempotent by target; the
   * sheet name is upserted onto the existing cached entry when present.
   * Returns the 1-based index the link will carry in the output file.
   */
  _recordAutoExternalLink(target: string, sheetName: string): number {
    const lower = target.toLowerCase();
    // If the user explicitly declared a link with this target, we respect
    // their definition verbatim: no sheetName upserts, no cache entry.
    // Excel needs the user-declared sheetNames to match the refs, and
    // augmenting them on the user's behalf could silently hide a typo.
    const existingUserIdx = this.externalLinks.findIndex(l => l.target.toLowerCase() === lower);
    if (existingUserIdx !== -1) {
      return existingUserIdx + 1;
    }
    let cached = this._writerExternalLinkCache.get(lower);
    if (!cached) {
      cached = {
        // Index is provisional — the real on-disk index is recomputed by
        // `_collectExternalLinksForWrite()` at serialisation time.
        index: 0,
        target,
        targetMode: "External",
        sheetNames: [],
        cachedValues: {}
      };
      this._writerExternalLinkCache.set(lower, cached);
    }
    if (sheetName && !cached.sheetNames.includes(sheetName)) {
      cached.sheetNames.push(sheetName);
    }
    // Recompute final index: user entries first, then cache entries in
    // insertion order. The caller needs the *on-disk* index so that the
    // formula it's rewriting matches the link that will be serialised.
    const userCount = this.externalLinks.length;
    let cacheIdx = 0;
    for (const key of this._writerExternalLinkCache.keys()) {
      cacheIdx++;
      if (key === lower) {
        return userCount + cacheIdx;
      }
    }
    // Unreachable — we just inserted.
    return userCount + this._writerExternalLinkCache.size;
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
      protection: this.protection,
      worksheets: this.worksheets.map(worksheet => worksheet.model),
      sheets: this.worksheets.map(ws => ws.model).filter(Boolean),
      definedNames: this._definedNames.model,
      // Live `DefinedNames` instance — required by the write-time
      // chartEx transform `prepareChartExSidecars`, which registers
      // hidden `_xlchart.vN.M` names on the fly and needs an object
      // that can mutate in place. The serialised `definedNames`
      // array above is re-materialised after the transform runs.
      definedNamesInstance: this._definedNames,
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
      defaultFont: this._defaultFont,
      externalLinks: this.externalLinks,
      chartEntries: this._chartEntries,
      chartRels: this._chartRels,
      chartStyles: this._chartStyles,
      chartColors: this._chartColors,
      chartExStyles: this._chartExStyles,
      chartExColors: this._chartExColors,
      chartExEntries: this._chartExEntries,
      chartExRels: this._chartExRels,
      chartExStructuredEntries: this._chartExStructuredEntries,
      chartsheets: this._chartsheets,
      persons: this._persons,
      slicerParts: this._slicerParts,
      slicerCacheParts: this._slicerCacheParts,
      timelineParts: this._timelineParts,
      timelineCacheParts: this._timelineCacheParts
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
    this.protection = value.protection;
    this.calcProperties = value.calcProperties;
    this._worksheets = [];
    this._tableNames.clear();
    value.worksheets.forEach(worksheetModel => {
      const { id, name, state } = worksheetModel;
      // API invariant: `_worksheets` is keyed by a positive integer
      // sheet id. A worksheet model with a missing or non-integer id
      // would be stored under a string pseudo key like `"undefined"`
      // or `"NaN"`, making it unreachable via `getWorksheet(name)`
      // (issue #166). The xlsx reconciler enforces the same invariant
      // before reaching this point; programmatic callers assigning
      // `model` directly with a malformed payload land here instead.
      if (!Number.isInteger(id) || (id as number) <= 0) {
        return;
      }
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

    // Preserve default font for round-trip fidelity
    this._defaultFont = value.defaultFont;
    // Restore chart entries
    this._chartEntries = value.chartEntries || {};
    this._chartRels = value.chartRels || {};
    this._chartStyles = value.chartStyles || {};
    this._chartColors = value.chartColors || {};
    this._chartExStyles = (value as any).chartExStyles || {};
    this._chartExColors = (value as any).chartExColors || {};
    this._chartExEntries = value.chartExEntries || {};
    this._chartExRels = value.chartExRels || {};
    this._chartExStructuredEntries = value.chartExStructuredEntries || {};
    // Restore chartsheets. Populate each chartsheet's `orderNo` from
    // the position in `value.sheets` (workbook.xml tab order) so the
    // writer's `prepare()` can sort interleaved worksheets +
    // chartsheets back into the author's layout. Matches the
    // equivalent loop above for worksheets.
    this._chartsheets = value.chartsheets || [];
    if (value.sheets) {
      for (const cs of this._chartsheets) {
        const idx = value.sheets.findIndex((s: { id?: number }) => s.id === cs.id);
        if (idx !== -1) {
          cs.orderNo = idx;
        }
      }
    }
    // Restore threaded-comment person directory. Always assign a new
    // list so callers editing the previous value don't mutate the
    // newly-loaded workbook by accident.
    this._persons = value.persons ? [...value.persons] : [];
    // Restore raw-passthrough slicer/timeline parts so dashboards
    // survive round-trip. The maps are stored by reference — loaders
    // and writers treat them as read-only; mutating them between
    // load and save is not supported.
    this._slicerParts = value.slicerParts ?? {};
    this._slicerCacheParts = value.slicerCacheParts ?? {};
    this._timelineParts = value.timelineParts ?? {};
    this._timelineCacheParts = value.timelineCacheParts ?? {};
    // Preserve external workbook references (empty array if none)
    this.externalLinks = value.externalLinks ? [...value.externalLinks] : [];
    // Reset the writer-scoped auto-discovery cache — loading a fresh
    // workbook replaces any accumulated state from previous writes.
    this._writerExternalLinkCache = new Map();
  }
}

function deepClone<T>(value: T): T {
  if (typeof structuredClone === "function") {
    return structuredClone(value);
  }
  return JSON.parse(JSON.stringify(value));
}

/**
 * Deep-copy a {@link ChartEntry}, preserving every field that affects
 * write-time behaviour — rawData (for the raw-patch fast path),
 * modelSnapshot (change detection), dirty / preferRawPatch /
 * requireRawPatch (writer hints), style / colors (ancillary parts),
 * rels (per-entry relationship bag), and userShapesXml (annotation
 * overlay). The caller supplies the new `chartNumber`; everything
 * else is a structural clone so later mutations on one entry don't
 * leak into the other.
 *
 * `rawData` and `userShapesXml` are `Uint8Array`s — `.slice()` is
 * used instead of `structuredClone` to keep the fast path cheap.
 */
function cloneChartEntry(entry: ChartEntry, chartNumber: number): ChartEntry {
  return {
    chartNumber,
    model: deepClone(entry.model),
    ...(entry.rawData ? { rawData: entry.rawData.slice() } : {}),
    ...(entry.modelSnapshot !== undefined ? { modelSnapshot: entry.modelSnapshot } : {}),
    ...(entry.dirty !== undefined ? { dirty: entry.dirty } : {}),
    ...(entry.preferRawPatch !== undefined ? { preferRawPatch: entry.preferRawPatch } : {}),
    ...(entry.requireRawPatch !== undefined ? { requireRawPatch: entry.requireRawPatch } : {}),
    ...(entry.style ? { style: deepClone(entry.style) } : {}),
    ...(entry.colors ? { colors: deepClone(entry.colors) } : {}),
    ...(entry.rels ? { rels: entry.rels.map(r => ({ ...r })) } : {}),
    ...(entry.userShapesXml ? { userShapesXml: entry.userShapesXml.slice() } : {})
  };
}

/**
 * Deep-copy a {@link ChartExEntry}, preserving the same write-time
 * fields as {@link cloneChartEntry} but for the ChartEx family
 * (structured model + rawData + dirty / preferRawPatch /
 * requireRawPatch + rels).
 */
function cloneChartExEntry(entry: ChartExEntry, chartExNumber: number): ChartExEntry {
  return {
    chartExNumber,
    model: deepClone(entry.model),
    ...(entry.rawData ? { rawData: entry.rawData.slice() } : {}),
    ...(entry.modelSnapshot !== undefined ? { modelSnapshot: entry.modelSnapshot } : {}),
    ...(entry.dirty !== undefined ? { dirty: entry.dirty } : {}),
    ...(entry.preferRawPatch !== undefined ? { preferRawPatch: entry.preferRawPatch } : {}),
    ...(entry.requireRawPatch !== undefined ? { requireRawPatch: entry.requireRawPatch } : {}),
    ...(entry.rels ? { rels: entry.rels.map(r => ({ ...r })) } : {})
  };
}

export { Workbook };
