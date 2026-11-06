/**
 * CSV Formatter Stream
 *
 * True streaming CSV formatter using cross-platform stream module.
 * Works identically in both Node.js and Browser environments.
 */

import type { FormatConfig, FormatRowOptions } from "@csv/format";
import { createFormatConfig, formatRowWithLookup, headerNames } from "@csv/format";
import type { CsvFormatOptions, Row } from "@csv/types";
import { checkFormatRow, checkObject } from "@csv/utils/check-options";
import {
  extractRowValues,
  detectRowKeys,
  processColumns,
  deduplicateHeaders
} from "@csv/utils/row";
import { Transform } from "@stream";

/** One row the formatter stream takes: the shapes `Csv.format` accepts for each row. */
type FormatterRow = Row | readonly unknown[] | Readonly<Record<string, unknown>>;

/**
 * Transform stream that formats rows to CSV
 *
 * @example
 * ```ts
 * const formatter = new CsvFormatterStream({ headers: ['name', 'age'] });
 * formatter.pipe(writable);
 * formatter.write(['Alice', 30]);
 * formatter.write(['Bob', 25]);
 * formatter.end();
 * ```
 */
export class CsvFormatterStream extends Transform {
  private options: CsvFormatOptions;
  /** Unified format configuration (shared with batch formatter) */
  private formatConfig: FormatConfig;
  private headerWritten: boolean = false;
  /** The first row was written as the header row (`firstRowIsHeader`). */
  private headerRowWritten = false;
  /** That row's texts, which name the columns. */
  private firstRowNames: string[] | undefined;
  /** Keys to access data from source objects */
  private keys: string[] | null = null;
  /** Headers to write to output (may differ from keys) */
  private displayHeaders: string[] | null = null;
  /** Index of source row (before filtering), passed to transform.row */
  private sourceRowIndex: number = 0;
  /** Index of output data row (after filtering, excludes header), used for ctx.index */
  private outputRowIndex: number = 0;
  /** Pre-allocated options object to avoid per-row allocation in streaming */
  declare private rowOptions: FormatRowOptions;

  // Improve public typing without relying on generic Transform types. The
  // declared chunk types WIDEN the base signature (they include the base's
  // byte chunk) rather than replace it: `Transform` resolves to Node's
  // non-generic class on the Node build and to the generic
  // `Transform<Uint8Array, Uint8Array>` on the browser build, and a member that
  // dropped the byte arm was not assignable to the latter.
  declare push: (chunk: string | Uint8Array | null, encoding?: string) => boolean;
  declare write: {
    (chunk: FormatterRow | Uint8Array, callback?: (error?: Error | null) => void): boolean;
    (
      chunk: FormatterRow | Uint8Array,
      encoding?: string,
      callback?: (error?: Error | null) => void
    ): boolean;
  };

  constructor(options: CsvFormatOptions = {}) {
    // `objectMode` is read before `createFormatConfig` checks the rest.
    checkObject("new Csv.FormatterStream", options);
    const rowsIn = options.objectMode !== false;
    super({
      // The two sides are set independently: rows arrive on the writable side,
      // while the readable side may be asked to produce bytes instead of
      // strings (see `CsvFormatOptions.readableObjectMode`). Defaults to the
      // writable setting, so existing callers keep getting string chunks.
      readableObjectMode: options.readableObjectMode ?? rowsIn,
      writableObjectMode: rowsIn
    });
    this.options = options;

    // Use shared config factory (same as batch formatter)
    this.formatConfig = createFormatConfig(options, "new Csv.FormatterStream");

    // Pre-allocate row options object (mutated per-row to avoid GC pressure)
    const cfg = this.formatConfig;
    this.rowOptions = {
      quoteLookup: cfg.shouldQuoteColumn,
      delimiter: cfg.delimiter,
      headers: undefined,
      isHeader: false,
      outputRowIndex: 0,
      quoteAll: cfg.quoteAll,
      escapeFormulae: cfg.escapeFormulae,
      decimalSeparator: cfg.decimalSeparator,
      transform: cfg.typeTransform,
      where: cfg.where
    };

    // Process columns config (takes precedence over headers)
    const columnsConfig = processColumns(options.columns);
    if (columnsConfig) {
      this.keys = columnsConfig.keys;
      this.displayHeaders = columnsConfig.headers;
    } else if (Array.isArray(options.headers)) {
      this.keys = options.headers;
      this.displayHeaders = options.headers;
    }

    // Deduplicate headers (consistent with batch formatCsv)
    if (this.displayHeaders) {
      this.displayHeaders = deduplicateHeaders(this.displayHeaders) as string[];
    }
  }

  /**
   * Auto-detect keys/headers from a row (object or RowHashArray)
   */
  private detectHeadersFromRow(chunk: Row): void {
    const detectedKeys = detectRowKeys(chunk);
    if (detectedKeys.length > 0) {
      this.keys = detectedKeys;
      this.displayHeaders = deduplicateHeaders(detectedKeys) as string[];
    }
  }

  _transform(
    chunkOrBytes: Row | Uint8Array,
    _encoding: string,
    // `data` is `never`: this stream reports errors through the callback and
    // emits rows through `this.push`, never through the callback's data channel.
    // Typing it so also keeps the override assignable to the byte-typed base
    // `_transform` on the browser build (see the note on `push` above).
    callback: (error?: Error | null, data?: never) => void
  ): void {
    // This stream is object-mode: only rows ever arrive. The byte arm exists so
    // the override stays assignable to the byte-typed base `_transform` (see the
    // note on `push` above).
    const chunk = chunkOrBytes as Row;
    try {
      checkFormatRow(chunk, this.sourceRowIndex, undefined, this.formatConfig.where);
      // Write BOM if first chunk
      if (!this.headerWritten && this.formatConfig.bom) {
        this.push("\uFEFF");
      }

      // Handle header writing on first row
      if (!this.headerWritten) {
        // Keyed rows (objects, RowHashArrays) are mapped by the first row's keys unless
        // `headers`/`columns` name them; those keys are the header row unless `headers: false`.
        if (!this.keys) {
          this.detectHeadersFromRow(chunk);
          if (this.options.headers === false) {
            this.displayHeaders = null;
          }
        }

        // Write headers if we should and have them
        if (this.formatConfig.writeHeaders && this.displayHeaders) {
          this.push(this.formatRow(this.displayHeaders, true));
        } else if (this.options.firstRowIsHeader === true && !this.displayHeaders) {
          // The rows carry their own header row: this one, written as a header row (`firstRowIsHeader`).
          const values = extractRowValues(chunk, this.keys);
          this.headerWritten = true;
          this.sourceRowIndex++;
          this.firstRowNames = headerNames(values, this.formatConfig);
          this.push(this.formatRow(values, true));
          this.headerRowWritten = true;
          callback();
          return;
        }
        this.headerWritten = true;
      }

      // Apply row-level transform if provided
      let processedChunk: Row | null = chunk;
      const sourceIndex = this.sourceRowIndex++;
      if (this.formatConfig.typeTransform?.row) {
        processedChunk = this.formatConfig.typeTransform.row(chunk, sourceIndex);
        if (processedChunk === null) {
          callback();
          return;
        }
        checkFormatRow(processedChunk, sourceIndex, "typeTransform.row", this.formatConfig.where);
      }

      this.formatAndPush(processedChunk);
      this.outputRowIndex++;
      callback();
    } catch (error) {
      callback(error as Error);
    }
  }

  _destroy(error: Error | null, callback: (error: Error | null) => void): void {
    this.keys = null;
    this.displayHeaders = null;
    callback(error);
  }

  _flush(callback: (error?: Error | null) => void): void {
    // Handle writeHeaders: true with no data - still write headers
    if (!this.headerWritten && this.displayHeaders && this.formatConfig.writeHeaders) {
      if (this.formatConfig.bom) {
        this.push("\uFEFF");
      }
      this.push(this.formatRow(this.displayHeaders, true));
      this.headerWritten = true;
    }

    // Add trailing newline if trailingNewline is true
    // hasOutput = wrote header OR wrote any data row
    const hasOutput =
      (this.formatConfig.writeHeaders && this.displayHeaders) ||
      this.headerRowWritten ||
      this.outputRowIndex > 0;
    if (this.formatConfig.trailingNewline && hasOutput) {
      this.push(this.formatConfig.lineEnding);
    }

    callback();
  }

  private formatAndPush(chunk: Row): void {
    const row = extractRowValues(chunk, this.keys);
    this.push(this.formatRow(row, false));
  }

  private formatRow(row: unknown[], isHeader: boolean = false): string {
    const cfg = this.formatConfig;
    // Mutate pre-allocated options to avoid per-row object allocation
    const opts = this.rowOptions;
    opts.quoteLookup = isHeader ? cfg.shouldQuoteHeader : cfg.shouldQuoteColumn;
    opts.headers = this.firstRowNames ?? this.displayHeaders ?? undefined;
    opts.isHeader = isHeader;
    opts.outputRowIndex = this.outputRowIndex;

    const formattedRow = formatRowWithLookup(row, cfg.regex, opts);

    // Use row delimiter as prefix (except for first output)
    // First output = header row OR (no header AND first data row)
    const isFirstLine =
      isHeader ||
      (!(cfg.writeHeaders && this.displayHeaders) &&
        !this.headerRowWritten &&
        this.outputRowIndex === 0);
    return isFirstLine ? formattedRow : cfg.lineEnding + formattedRow;
  }
}
