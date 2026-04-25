/**
 * XLSX - Abstract base class for XLSX operations
 *
 * Contains all platform-agnostic logic shared between Node.js and Browser versions:
 * - reconcile: Reconcile model after parsing
 * - _process*Entry: Process individual ZIP entries
 * - add*: Add content to ZIP during writing
 * - prepareModel: Prepare model for writing
 * - loadFromFiles: Load from pre-extracted ZIP data
 */

import { ZipParser } from "@archive/unzip/zip-parser";
import type { ZipTimestampMode } from "@archive/zip-spec/timestamps";
import { StreamingZip, ZipDeflateFile } from "@archive/zip/stream";
import type { ChartExEntry } from "@excel/chart/chart";
import { renderChartEx } from "@excel/chart/chart-ex-renderer";
import {
  ExcelStreamStateError,
  ExcelFileError,
  ImageError,
  ExcelNotSupportedError,
  XmlParseError,
  TableError
} from "@excel/errors";
import type { PivotTable, PivotTableSubtotal, ParsedCacheDefinition } from "@excel/pivot-table";
import { filterDrawingAnchors } from "@excel/utils/drawing-utils";
import { rewriteExternalRefs } from "@excel/utils/external-link-formula";
import {
  commentsPath,
  chartsheetPath,
  chartsheetRelsPath,
  getChartsheetNoFromPath,
  getChartsheetNoFromRelsPath,
  ctrlPropPath,
  drawingPath,
  drawingRelsPath,
  externalLinkPath,
  externalLinkRelsPath,
  externalLinkRelTargetFromWorkbook,
  OOXML_REL_TARGETS,
  pivotCacheDefinitionRelTargetFromWorkbook,
  pivotTablePathFromName,
  isCommentsPath,
  chartPath,
  chartRelsPath,
  chartStylePath,
  chartColorsPath,
  chartStyleRelTarget,
  chartExPath,
  chartExRelsPath,
  getChartExNumberFromPath,
  getChartExNumberFromRelsPath,
  chartColorsRelTarget,
  chartRelTargetFromDrawing,
  chartExRelTargetFromDrawing,
  getChartNumberFromPath,
  getChartNumberFromRelsPath,
  getChartStyleNumberFromPath,
  getChartColorsNumberFromPath,
  getDrawingNameFromPath,
  getDrawingNameFromRelsPath,
  getExternalLinkIndexFromPath,
  getExternalLinkIndexFromRelsPath,
  getMediaFilenameFromPath,
  mediaPath,
  getPivotCacheDefinitionNameFromPath,
  getPivotCacheDefinitionNameFromRelsPath,
  getPivotCacheRecordsNameFromPath,
  getPivotTableNameFromPath,
  getPivotTableNameFromRelsPath,
  pivotCacheDefinitionPath,
  pivotCacheDefinitionRelsPath,
  pivotCacheDefinitionRelTargetFromPivotTable,
  pivotCacheRecordsPath,
  pivotCacheRecordsRelTarget,
  pivotTablePath,
  pivotTableRelsPath,
  getTableNameFromPath,
  tablePath,
  themePath,
  getThemeNameFromPath,
  getVmlDrawingNameFromPath,
  getVmlDrawingHFNameFromPath,
  getWorksheetNoFromWorksheetPath,
  getWorksheetNoFromWorksheetRelsPath,
  isBinaryEntryPath,
  normalizeZipPath,
  OOXML_PATHS,
  vmlDrawingPath,
  vmlDrawingHFPath,
  vmlDrawingHFRelsPath,
  worksheetPath,
  worksheetRelsPath,
  worksheetRelTarget
} from "@excel/utils/ooxml-paths";
import { StreamBuf } from "@excel/utils/stream-buf";
import type { Workbook } from "@excel/workbook";
import type { ExternalLinkModel } from "@excel/workbook.browser";
import { RelType } from "@excel/xlsx/rel-type";
import {
  ExternalLinkXform,
  type ParsedExternalLink
} from "@excel/xlsx/xform/book/external-link-xform";
import { WorkbookXform } from "@excel/xlsx/xform/book/workbook-xform";
import { ChartSpaceXform } from "@excel/xlsx/xform/chart/chart-space-xform";
import { CommentsXform } from "@excel/xlsx/xform/comment/comments-xform";
import { AppXform } from "@excel/xlsx/xform/core/app-xform";
import { ContentTypesXform } from "@excel/xlsx/xform/core/content-types-xform";
import { CoreXform } from "@excel/xlsx/xform/core/core-xform";
import { FeaturePropertyBagXform } from "@excel/xlsx/xform/core/feature-property-bag-xform";
import { MetadataXform } from "@excel/xlsx/xform/core/metadata-xform";
import { RelationshipsXform } from "@excel/xlsx/xform/core/relationships-xform";
import { CtrlPropXform } from "@excel/xlsx/xform/drawing/ctrl-prop-xform";
import { DrawingXform } from "@excel/xlsx/xform/drawing/drawing-xform";
import { VmlDrawingXform } from "@excel/xlsx/xform/drawing/vml-drawing-xform";
import { PivotCacheDefinitionXform } from "@excel/xlsx/xform/pivot-table/pivot-cache-definition-xform";
import { PivotCacheRecordsXform } from "@excel/xlsx/xform/pivot-table/pivot-cache-records-xform";
import {
  PivotTableXform,
  type ParsedPivotTableModel
} from "@excel/xlsx/xform/pivot-table/pivot-table-xform";
import { ChartsheetXform } from "@excel/xlsx/xform/sheet/chartsheet-xform";
import { WorkSheetXform } from "@excel/xlsx/xform/sheet/worksheet-xform";
import { SharedStringsXform } from "@excel/xlsx/xform/strings/shared-strings-xform";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { TableXform } from "@excel/xlsx/xform/table/table-xform";
import { theme1Xml } from "@excel/xlsx/xml/theme1";
import { PassThrough, type IEventEmitter } from "@stream";
import { concatUint8Arrays } from "@utils/binary";
import { bufferToString, base64ToUint8Array } from "@utils/utils";
import { XmlStreamWriter } from "@xml/stream-writer";

type StreamListener = Parameters<IEventEmitter["on"]>[1];

interface EmitterLike {
  on(event: string, listener: StreamListener): this;
  once(event: string, listener: StreamListener): this;
  off(event: string, listener: StreamListener): this;
}

export interface IParseStream extends EmitterLike {
  pipe(dest: any): any;
  [Symbol.asyncIterator]?: () => AsyncIterator<Uint8Array | string>;
}

/**
 * Minimal write-side shape required to receive XLSX bytes. Anything that
 * behaves like a Node `WritableStream` (a `write()` method plus `end()` and
 * event emitter basics) satisfies this — including Node's `fs.WriteStream`,
 * `PassThrough`, and our `@stream` Writable class.
 */
export interface IWritableStream extends EmitterLike {
  write(data: string | Uint8Array): boolean | void | Promise<boolean>;
  end(): void;
  // `pipe` is the Node stream ecosystem's polymorphic dispatcher; its return
  // type depends entirely on the destination. Typed as `any` so callers can
  // freely chain `.pipe(next).pipe(another)` without forced type assertions.
  pipe?(dest: any): any;
}

/**
 * An in-memory buffered stream. Extends the write-side shape with `read()`
 * and optional `toBuffer()` for callers that use the stream as a sink and
 * then harvest the accumulated bytes (e.g. `writeBuffer()`'s internal buffer).
 */
export interface IStreamBuf extends IWritableStream {
  read(): Uint8Array | null;
  toBuffer?(): Uint8Array | null;
}

export interface IZipWriter extends EmitterLike {
  append(data: string | Uint8Array, options: { name: string; base64?: boolean }): void;
  /** Create a streaming entry: write chunks incrementally, then call end(). */
  createEntry(name: string): { write(chunk: string): void; end(): void };
  pipe(stream: IWritableStream): void;
  finalize(): void;
  /** Wait for downstream backpressure to clear. Resolves immediately if no backpressure. */
  waitForDrain(): Promise<void>;
}

class StreamingZipWriterAdapter implements IZipWriter {
  private static textEncoder = new TextEncoder();

  private readonly zip: StreamingZip;
  private readonly events: Map<string, Set<StreamListener>> = new Map();
  private pipedStream: Pick<IWritableStream, "write" | "end"> | null = null;
  private level: number;
  private modTime: Date | undefined;
  private timestamps: ZipTimestampMode | undefined;
  private finalized = false;

  // Backpressure tracking
  private _needsDrain = false;
  private _drainResolvers: Array<() => void> = [];
  // Count of in-flight async write() calls whose backpressure result is unknown.
  // waitForDrain() must wait for this to reach 0 before checking _needsDrain.
  private _pendingWrites = 0;
  private _pendingWriteResolvers: Array<() => void> = [];

  // Buffer errors that occur before _finalize registers its error listener,
  // so async compression errors during writeToZip() are never silently lost.
  private _earlyError: Error | null = null;

  constructor(options?: ZipWriterOptions) {
    this.level = options?.level ?? 6;
    this.modTime = options?.modTime;
    this.timestamps = options?.timestamps;
    this.zip = new StreamingZip((err: Error | null, data: Uint8Array, final: boolean) => {
      if (err) {
        this._emit("error", err);
        return;
      }

      if (data && data.length > 0) {
        this._emit("data", data);
        if (this.pipedStream) {
          this._checkBackpressure(this.pipedStream.write(data));
        }
      }

      if (final) {
        if (this.pipedStream) {
          this.pipedStream.end();
        }
        this._emit("finish");
      }
    });
  }

  private _emit(event: string, ...args: any[]): void {
    // Buffer error events that fire before any listener is registered,
    // so _finalize() can surface them even if it registers late.
    if (event === "error") {
      const callbacks = this.events.get(event);
      if (!callbacks || callbacks.size === 0) {
        this._earlyError = args[0] instanceof Error ? args[0] : new Error(String(args[0]));
        return;
      }
    }
    const callbacks = this.events.get(event);
    if (!callbacks) {
      return;
    }
    for (const cb of callbacks) {
      cb(...args);
    }
  }

  /**
   * Handle backpressure from pipedStream.write().
   * Accepts both sync (boolean) and async (Promise<boolean>) return values.
   */
  private _checkBackpressure(ok: boolean | void | Promise<boolean>): void {
    if (ok instanceof Promise) {
      this._pendingWrites++;
      ok.then(
        result => {
          if (!result) {
            this._needsDrain = true;
          }
        },
        () => {} // write errors surface via the stream's error event
      ).finally(() => {
        this._pendingWrites--;
        if (this._pendingWrites === 0) {
          const resolvers = this._pendingWriteResolvers.splice(0);
          for (const resolve of resolvers) {
            resolve();
          }
        }
      });
      return;
    }
    if (ok === false) {
      this._needsDrain = true;
    }
  }

  on(event: string, callback: StreamListener): this {
    const callbacks = this.events.get(event) || new Set<StreamListener>();
    callbacks.add(callback);
    this.events.set(event, callbacks);

    // If an error was buffered before any listener was registered, deliver it now.
    if (event === "error" && this._earlyError) {
      const err = this._earlyError;
      this._earlyError = null;
      callback(err);
    }

    return this;
  }

  once(event: string, callback: StreamListener): this {
    const wrapped: StreamListener = (...args: any[]) => {
      this.off(event, wrapped);
      callback(...args);
    };
    return this.on(event, wrapped);
  }

  off(event: string, callback: StreamListener): this {
    const callbacks = this.events.get(event);
    if (!callbacks) {
      return this;
    }
    callbacks.delete(callback);
    if (callbacks.size === 0) {
      this.events.delete(event);
    }
    return this;
  }

  pipe(stream: IWritableStream): void {
    this.pipedStream = stream;
    // Listen for drain events to resolve backpressure waiters
    if (stream && typeof stream.on === "function") {
      stream.on("drain", () => {
        this._needsDrain = false;
        const resolvers = this._drainResolvers.splice(0);
        for (const resolve of resolvers) {
          resolve();
        }
      });
    }
  }

  /**
   * Wait for the downstream writable to drain if it signaled backpressure.
   * If any write() calls are still in-flight (returned a Promise that hasn't
   * settled), waits for all of them first so the backpressure signal isn't missed.
   */
  async waitForDrain(): Promise<void> {
    // Wait for all in-flight async writes to settle so _needsDrain is up to date.
    if (this._pendingWrites > 0) {
      await new Promise<void>(resolve => {
        this._pendingWriteResolvers.push(resolve);
      });
    }
    if (!this._needsDrain || !this.pipedStream) {
      return;
    }
    return new Promise<void>(resolve => {
      this._drainResolvers.push(resolve);
    });
  }

  append(data: any, options: { name: string; base64?: boolean }): void {
    if (this.finalized) {
      throw new ExcelStreamStateError("append", "stream already finalized");
    }

    let buffer: Uint8Array;
    if (options.base64) {
      buffer = base64ToUint8Array(typeof data === "string" ? data : String(data));
    } else if (typeof data === "string") {
      buffer = StreamingZipWriterAdapter.textEncoder.encode(data);
    } else if (data instanceof Uint8Array) {
      buffer = data;
    } else if (ArrayBuffer.isView(data)) {
      buffer = new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    } else if (data instanceof ArrayBuffer) {
      buffer = new Uint8Array(data);
    } else {
      buffer = data;
    }

    const file = new ZipDeflateFile(options.name, {
      level: this.level,
      modTime: this.modTime,
      timestamps: this.timestamps
    });
    this.zip.add(file);

    file.push(buffer, true);
  }

  createEntry(name: string): { write(chunk: string): void; end(): void } {
    if (this.finalized) {
      throw new ExcelStreamStateError("createEntry", "stream already finalized");
    }
    const file = new ZipDeflateFile(name, {
      level: this.level,
      modTime: this.modTime,
      timestamps: this.timestamps
    });
    this.zip.add(file);
    const encoder = StreamingZipWriterAdapter.textEncoder;
    return {
      write(chunk: string): void {
        file.push(encoder.encode(chunk));
      },
      end(): void {
        file.push(new Uint8Array(0), true);
      }
    };
  }

  finalize(): void {
    if (this.finalized) {
      return;
    }
    this.finalized = true;
    this.zip.end();
  }
}

// =============================================================================
// Minimal shared types (keep internal model flexible)
// =============================================================================

/**
 * Options for reading (loading) an XLSX workbook.
 *
 * All officially supported options are declared below with proper types so
 * callers get IDE completion and type-checking. Additional fields are
 * permitted via the index signature for forward compatibility and for
 * callers who subclass `XLSX` to pass through private flags.
 */
export interface XlsxReadOptions {
  /**
   * When the input to `load()` is a string, interpret it as a base64-encoded
   * zip archive instead of a binary buffer. Defaults to `false`.
   */
  base64?: boolean;
  /**
   * Maximum number of rows to parse from each worksheet. Rows beyond this
   * limit are silently skipped. Useful for previewing very large sheets
   * without loading everything into memory.
   */
  maxRows?: number;
  /**
   * Maximum number of columns to parse from each worksheet. Columns beyond
   * this limit are silently skipped. Useful for previewing very wide sheets.
   */
  maxCols?: number;
  /**
   * List of worksheet XML node names to skip while parsing (e.g.
   * `"dataValidations"`, `"conditionalFormatting"`). Use for workbooks that
   * contain corrupted or unsupported elements you want to ignore.
   */
  ignoreNodes?: string[];
  /**
   * Forward-compatibility / subclass extension escape hatch. Unknown keys are
   * passed through to internal loaders; unrecognised keys are ignored.
   */
  [key: string]: unknown;
}

export interface ZipWriterOptions {
  level?: number;
  /** ZIP entry modification time (optional). If omitted, defaults to current time. */
  modTime?: Date;
  /** Timestamp writing strategy for ZIP entry metadata (optional). */
  timestamps?: ZipTimestampMode;
}

/**
 * Options for writing an XLSX workbook.
 *
 * All officially supported options are declared below with proper types so
 * callers get IDE completion and type-checking. Additional fields are
 * permitted via the index signature for forward compatibility and for
 * callers who subclass `XLSX` to pass through private flags.
 */
export interface XlsxWriteOptions {
  /** ZIP archive options (compression level, timestamps, ...). */
  zip?: ZipWriterOptions;
  /**
   * Use a shared-string table for cell text values. Defaults to `true`.
   * Set to `false` to write string values inline (larger file, but streams
   * better for very large sheets).
   */
  useSharedStrings?: boolean;
  /**
   * Emit style definitions (fonts, fills, borders, number formats, …).
   * Defaults to `true`. Set to `false` to skip style blocks for maximum
   * compatibility with minimal readers.
   */
  useStyles?: boolean;
  /**
   * Forward-compatibility / subclass extension escape hatch. Unknown keys are
   * passed through to internal writers; unrecognised keys are ignored.
   */
  [key: string]: unknown;
}

export type XlsxOptions = XlsxReadOptions & XlsxWriteOptions;

export interface WorkbookMediaLike {
  type: string;
  extension: string;
  name?: string;
  filename?: string;
  buffer?: Uint8Array;
  base64?: string;
}

export interface MediaModel {
  media: WorkbookMediaLike[];
}

interface ZipEntryLike {
  name: string;
  // Mirrors the `type` field on streaming ZipEntry. Symlink detection in
  // streaming parsers is best-effort; for XLSX (which contains no symlinks)
  // this is effectively "Directory" | "File" at runtime.
  type: "Directory" | "File" | "Symlink";
  stream: IParseStream;
  drain: () => Promise<void>;
}

/**
 * Extract the trailing integer from a workbook-rels Target like
 * `"externalLinks/externalLink12.xml"` (a path relative to `xl/`). Mirror
 * of {@link getExternalLinkIndexFromPath} which takes the full
 * `xl/externalLinks/…` form. Used during reconcile to bridge the
 * workbook.xml.rels entry to the parsed externalLinkN.xml part.
 */
function externalLinkIndexFromRelTarget(target: string): number | undefined {
  const match = /(?:^|\/)externalLink(\d+)[.]xml$/.exec(target);
  return match ? parseInt(match[1], 10) : undefined;
}

/**
 * Add `sheetName` to an ExternalLinkModel's `sheetNames` list if it isn't
 * already present. Ordering is preserved — the first-seen sheet wins
 * position 0, which matches what Excel does when writing externalLinks
 * itself.
 */
function upsertSheet(link: { sheetNames: string[] }, sheetName: string): void {
  if (!sheetName) {
    return;
  }
  if (!link.sheetNames.includes(sheetName)) {
    link.sheetNames.push(sheetName);
  }
}

/**
 * Scratch state used during `_normaliseExternalLinks`. `links` is the
 * write-scoped ExternalLinkModel array (user-declared + auto-discovered)
 * the writer will consume; `byTarget` indexes it by lower-cased target
 * for O(1) lookup during formula rewriting; `workbook` is used to
 * persist auto-discoveries to the Workbook's private cache so subsequent
 * writes stay consistent without mutating `wb.externalLinks`.
 */
interface NormaliseScratch {
  links: ExternalLinkModel[];
  byTarget: Map<string, ExternalLinkModel>;
  workbook: Workbook;
}

/**
 * Shape of a parsed externalLinkN.xml.rels entry. Kept broad so any
 * relationship type passes through verbatim — the writer only cares about
 * the ExternalLinkPath entry, but we preserve the rest for round-trip.
 */
type ExternalLinkRelsEntry = {
  Id: string;
  Type: string;
  Target: string;
  TargetMode?: string;
};

/**
 * XLSX class - handles Excel file operations
 * Works in both Node.js and Browser environments
 */
class XLSX {
  declare public workbook: Workbook;

  static RelType = RelType;

  constructor(workbook: Workbook) {
    this.workbook = workbook;
  }

  // ===========================================================================
  // Stream creation - cross-platform implementation using modules/stream
  // ===========================================================================

  /**
   * Create a stream from binary data (for media/themes)
   */
  protected createBinaryStream(data: Uint8Array): IParseStream {
    const stream = new PassThrough();
    stream.end(data);
    return stream;
  }

  /**
   * Create a stream from string content (for XML parsing)
   */
  protected createTextStream(content: string): IParseStream {
    const stream = new PassThrough();
    stream.end(content);
    return stream;
  }

  // ===========================================================================
  // Shared implementations - used by all platforms
  // ===========================================================================

  /**
   * Create a StreamBuf instance for buffering data
   */
  protected createStreamBuf(): IStreamBuf {
    return new StreamBuf();
  }

  /**
   * Convert buffer/Uint8Array to string
   */
  protected bufferToString(data: string | ArrayBuffer | Uint8Array): string {
    return bufferToString(data);
  }

  /**
   * Create a ZIP writer adapter.
   * Can be overridden by subclasses for platform-specific implementations.
   */
  protected createZipWriter(options?: XlsxWriteOptions["zip"]): IZipWriter {
    return new StreamingZipWriterAdapter(options);
  }

  /**
   * Write all workbook content to a ZIP writer
   * Shared by both Node.js write() and browser writeBuffer()
   */
  protected async writeToZip(zip: IZipWriter, options?: XlsxWriteOptions): Promise<void> {
    const { model } = this.workbook;
    this.prepareModel(model, options);

    await this.addContentTypes(zip, model);
    await this.addOfficeRels(zip, model);
    await this.addWorkbookRels(zip, model);
    // Write workbook.xml before worksheets so that streaming readers can
    // resolve worksheet names/ids/state from workbook metadata before
    // processing worksheet entries.
    await this.addWorkbook(zip, model);
    await this.addWorksheets(zip, model);
    await this.addChartsheets(zip, model);
    await this.addSharedStrings(zip, model);
    await this.addDrawings(zip, model);
    await this.addCharts(zip, model);
    await this.addChartExEntries(zip, model);
    await this.addTables(zip, model);
    await this.addPivotTables(zip, model);
    await this.addExternalLinks(zip, model);
    await this.addThemes(zip, model);
    await this.addStyles(zip, model);
    await this.addFeaturePropertyBag(zip, model);
    await this.addMetadata(zip, model);
    await this.addMedia(zip, model);
    await this.addApp(zip, model);
    await this.addCore(zip, model);
  }

  // ===========================================================================
  // Stream/Buffer operations - shared by all platforms
  // ===========================================================================

  /**
   * Read workbook from a stream
   */
  async read(stream: IParseStream, options?: XlsxReadOptions): Promise<Workbook> {
    // Collect all stream data into a single buffer
    const chunks: Uint8Array[] = [];

    await new Promise<void>((resolve, reject) => {
      const onData = (chunk: Uint8Array) => {
        chunks.push(chunk);
      };

      const onEnd = () => {
        stream.off("data", onData);
        stream.off("end", onEnd);
        stream.off("error", onError);
        resolve();
      };

      const onError = (err: Error) => {
        stream.off("data", onData);
        stream.off("end", onEnd);
        stream.off("error", onError);
        reject(err);
      };

      stream.on("data", onData);
      stream.on("end", onEnd);
      stream.on("error", onError);
    });

    return this.loadBuffer(concatUint8Arrays(chunks), options);
  }

  /**
   * Write workbook to a stream
   */
  async write(stream: IWritableStream, options?: XlsxWriteOptions): Promise<XLSX> {
    options = options || {};

    options.zip = options.zip || {};
    options.zip.modTime ??= this.workbook.modified ?? this.workbook.created;

    const zip = this.createZipWriter(options.zip);
    zip.pipe(stream);
    await this.writeToZip(zip, options);
    return this._finalize(zip) as Promise<XLSX>;
  }

  /**
   * Load a workbook from binary data.
   *
   * Accepted inputs:
   *  - `Uint8Array` (and `Buffer`, which is a Uint8Array at runtime)
   *  - `ArrayBuffer` / `SharedArrayBuffer`
   *  - Any `ArrayBufferView` (DataView, Int8Array, Float32Array, …) — the
   *    underlying bytes are reinterpreted as a zip archive
   *  - `string` — treated as base64-encoded data when `options.base64 === true`;
   *    raw binary cannot be round-tripped through a JS string and is rejected
   *    to prevent silent corruption.
   */
  async load(
    data: Uint8Array | ArrayBuffer | ArrayBufferView | string,
    options?: XlsxReadOptions
  ): Promise<Workbook> {
    if (data === null || data === undefined) {
      throw new ExcelFileError(
        "<input>",
        "read",
        "Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?"
      );
    }

    let buffer: Uint8Array;

    if (typeof data === "string") {
      // Strings must be base64-encoded — binary zip bytes cannot be round-tripped
      // through a JS string without corruption. Require the explicit opt-in.
      if (!options?.base64) {
        throw new ExcelFileError(
          "<input>",
          "read",
          "Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ? " +
            "String input requires options.base64 === true (base64-encoded zip archive)."
        );
      }
      buffer = base64ToUint8Array(data);
    } else if (data instanceof Uint8Array) {
      // Covers Buffer (Node) and any typed-array view whose element size is 1.
      buffer = data;
    } else if (data instanceof ArrayBuffer) {
      buffer = new Uint8Array(data);
    } else if (ArrayBuffer.isView(data)) {
      // DataView, Int8Array, Float32Array, … — view onto an underlying buffer.
      buffer = new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    } else {
      throw new ExcelFileError(
        "<input>",
        "read",
        "Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?"
      );
    }

    return this.loadBuffer(buffer, options);
  }

  /**
   * Internal: Load from Uint8Array buffer
   */
  protected async loadBuffer(buffer: Uint8Array, options?: XlsxReadOptions): Promise<Workbook> {
    const parser = new ZipParser(buffer);
    const filesMap = await parser.extractAll();

    // Convert Map to Record for loadFromFiles
    const allFiles: Record<string, Uint8Array> = {};
    for (const [path, content] of filesMap) {
      allFiles[path] = content;
    }

    return this.loadFromFiles(allFiles, options);
  }

  /**
   * Internal: Load workbook from an async stream of ZIP entries.
   *
   * This is the foundation for TRUE streaming reads on platforms that have a
   * streaming ZIP parser (e.g. Node.js `modules/archive` Parse).
   */
  /**
   * Create an empty model for parsing XLSX files.
   * Shared by loadFromZipEntries and loadFromFiles.
   */
  private createEmptyModel(): any {
    return {
      worksheets: [],
      worksheetHash: {},
      worksheetRels: [],
      themes: {},
      media: [],
      mediaIndex: {},
      drawings: {},
      drawingRels: {},
      comments: {},
      tables: {},
      vmlDrawings: {},
      pivotTables: {},
      pivotTableRels: {},
      pivotCacheDefinitions: {},
      pivotCacheRecords: {},
      // Parsed chart entries keyed by chart number
      chartEntries: {} as Record<number, any>,
      // Parsed chart rels keyed by chart number
      chartRels: {} as Record<number, any>,
      // Raw chart style bytes keyed by style number
      chartStyles: {} as Record<number, Uint8Array>,
      // Raw chart colors bytes keyed by colors number
      chartColors: {} as Record<number, Uint8Array>,
      // Raw chartEx entries (Office 2016+ extended charts) keyed by chartEx number
      chartExEntries: {} as Record<number, Uint8Array>,
      // Parsed chartEx rels keyed by chartEx number
      chartExRels: {} as Record<number, any[]>,
      // Structured chartEx entries (built via addChartEx) keyed by chartEx number
      chartExStructuredEntries: {} as Record<number, ChartExEntry>,
      // External workbook links — parsed from xl/externalLinks/externalLinkN.xml
      // during _processDefaultEntry, then reconciled into a dense
      // ExternalLinkModel[] by reconcile() using workbookRels + <externalReferences>.
      externalLinksByIndex: {} as Record<number, ParsedExternalLink>,
      // Raw rels from each externalLinkN.rels file, keyed by index.
      // Contains the actual Target path (e.g. "测试.xlsx", "file:///...")
      // and TargetMode ("External" / "Internal").
      externalLinkRelsByIndex: {} as Record<number, ExternalLinkRelsEntry[]>,
      // Chartsheets keyed by sheet number
      chartsheets: {} as Record<number, any>,
      chartsheetRels: {} as Record<number, any[]>
    };
  }

  /**
   * Collect all data from a stream into a single Uint8Array.
   * Reusable helper for passthrough and drawing processing.
   */
  protected async collectStreamData(stream: IParseStream): Promise<Uint8Array> {
    const chunks: Uint8Array[] = [];
    await new Promise<void>((resolve, reject) => {
      stream.on("data", (chunk: any) => {
        if (typeof chunk === "string") {
          chunks.push(new TextEncoder().encode(chunk));
        } else if (chunk instanceof Uint8Array) {
          chunks.push(chunk);
        } else {
          chunks.push(new Uint8Array(chunk));
        }
      });
      stream.on("end", () => resolve());
      stream.on("error", reject);
    });
    return concatUint8Arrays(chunks);
  }

  /**
   * Process a known OOXML entry (workbook, styles, shared strings, etc.)
   * Returns true if handled, false if should be passed to _processDefaultEntry
   */
  protected async _processKnownEntry(
    stream: IParseStream,
    model: any,
    entryName: string,
    options?: XlsxOptions
  ): Promise<boolean> {
    const sheetNo = getWorksheetNoFromWorksheetPath(entryName);
    if (sheetNo !== undefined) {
      await this._processWorksheetEntry(stream, model, sheetNo, options, entryName);
      return true;
    }

    const chartsheetNo = getChartsheetNoFromPath(entryName);
    if (chartsheetNo !== undefined) {
      await this._processChartsheetEntry(stream, model, chartsheetNo);
      return true;
    }

    switch (entryName) {
      case OOXML_PATHS.rootRels:
        model.globalRels = await this.parseRels(stream);
        return true;
      case OOXML_PATHS.xlWorkbook: {
        const workbook = await this.parseWorkbook(stream);
        model.sheets = workbook.sheets;
        model.definedNames = workbook.definedNames;
        model.views = workbook.views;
        model.properties = workbook.properties;
        model.protection = workbook.protection;
        model.calcProperties = workbook.calcProperties;
        model.pivotCaches = workbook.pivotCaches;
        // Pass-through the ordered list of <externalReference> rIds. These
        // get resolved into a dense externalLinks[] during reconcile().
        model.externalReferences = workbook.externalReferences;
        return true;
      }
      case OOXML_PATHS.xlSharedStrings:
        model.sharedStrings = new SharedStringsXform();
        await model.sharedStrings.parseStream(stream);
        return true;
      case OOXML_PATHS.xlWorkbookRels:
        model.workbookRels = await this.parseRels(stream);
        return true;
      case OOXML_PATHS.docPropsApp: {
        const appXform = new AppXform();
        const appProperties = await appXform.parseStream(stream);
        if (appProperties) {
          model.company = appProperties.company;
          model.manager = appProperties.manager;
        }
        return true;
      }
      case OOXML_PATHS.docPropsCore: {
        const coreXform = new CoreXform();
        const coreProperties = await coreXform.parseStream(stream);
        Object.assign(model, coreProperties);
        return true;
      }
      case OOXML_PATHS.xlStyles:
        model.styles = new StylesXform();
        await model.styles.parseStream(stream);
        return true;
      case OOXML_PATHS.xlMetadata: {
        const metadataXform = new MetadataXform();
        const metadataResult = await metadataXform.parseStream(stream);
        if (metadataResult) {
          model.metadata = metadataResult;
        }
        return true;
      }
      default:
        return false;
    }
  }

  protected async loadFromZipEntries(
    entries: AsyncIterable<ZipEntryLike>,
    options?: XlsxOptions
  ): Promise<any> {
    const model: any = this.createEmptyModel();

    for await (const entry of entries) {
      let drained = false;
      const drainEntry = async () => {
        if (drained) {
          return;
        }
        drained = true;
        await entry.drain();
      };

      if (entry.type === "Directory") {
        await drainEntry();
        continue;
      }

      const entryName = normalizeZipPath(entry.name);
      const stream = entry.stream;

      try {
        const handled = await this._processKnownEntry(stream, model, entryName, options);
        if (!handled) {
          const defaultHandled = await this._processDefaultEntry(stream, model, entryName);
          if (!defaultHandled) {
            // Important for true streaming parsers: always consume unknown entries
            await drainEntry();
          }
        }
      } finally {
        // Make sure we don't leave the entry stream partially consumed.
        // This is critical for true streaming parsers which may otherwise abort
        // the underlying entry stream (showing up as AbortError/ABORT_ERR).
        try {
          await drainEntry();
        } catch {
          // ignore drain errors; the primary parse error (if any) is more useful
        }
      }
    }

    this.reconcile(model, options);
    this.workbook.model = model;
    return this.workbook;
  }

  /**
   * Write workbook to buffer
   */
  async writeBuffer(options?: XlsxWriteOptions): Promise<Uint8Array> {
    options = options || {};

    options.zip = options.zip || {};
    options.zip.modTime ??= this.workbook.modified ?? this.workbook.created;

    const zip = this.createZipWriter(options.zip);
    const stream = this.createStreamBuf();
    zip.pipe(stream);
    await this.writeToZip(zip, options);
    await this._finalize(zip);
    return stream.read() || new Uint8Array(0);
  }

  // ===========================================================================
  // Media handling - base implementation (buffer/base64 only)
  // ===========================================================================

  /**
   * Add media files to ZIP
   * Supports buffer, base64, and filename (if readFileAsync is provided)
   */
  async addMedia(zip: IZipWriter, model: MediaModel): Promise<void> {
    await Promise.all(
      model.media.map(async (medium: WorkbookMediaLike) => {
        if (medium.type !== "image") {
          throw new ImageError("Unsupported media");
        }

        // Preserve legacy behavior: `${undefined}` becomes "undefined" in template strings
        const mediaName = medium.name ?? "undefined";
        const filename = mediaPath(`${mediaName}.${medium.extension}`);

        if (medium.filename) {
          if (this.readFileAsync) {
            const data = await this.readFileAsync(medium.filename);
            return zip.append(data, { name: filename });
          }
          throw new ExcelNotSupportedError(
            "Loading images from filename",
            "not supported in this environment"
          );
        }

        if (medium.buffer) {
          return zip.append(medium.buffer, { name: filename });
        }

        if (medium.base64) {
          const content = medium.base64.substring(medium.base64.indexOf(",") + 1);
          return zip.append(content, { name: filename, base64: true });
        }

        throw new ImageError("Unsupported media");
      })
    );
  }

  /**
   * Optional file reader - can be overridden by subclasses (e.g., Node.js version)
   */
  protected readFileAsync?: (filename: string) => Promise<Uint8Array>;

  // ===========================================================================
  // Parse helpers - shared by all platforms
  // ===========================================================================

  parseRels(stream: IParseStream): Promise<any> {
    const xform = new RelationshipsXform();
    return xform.parseStream(stream);
  }

  parseWorkbook(stream: IParseStream): Promise<any> {
    const xform = new WorkbookXform();
    return xform.parseStream(stream);
  }

  parseSharedStrings(stream: IParseStream): Promise<any> {
    const xform = new SharedStringsXform();
    return xform.parseStream(stream);
  }

  // ===========================================================================
  // Reconcile - shared by all platforms
  // ===========================================================================

  reconcile(model: any, options?: XlsxOptions): void {
    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorkSheetXform(options);
    const drawingXform = new DrawingXform();
    const tableXform = new TableXform();

    workbookXform.reconcile(model);

    // reconcile drawings with their rels
    const drawingOptions: any = {
      media: model.media,
      mediaIndex: model.mediaIndex
    };
    Object.keys(model.drawings).forEach(name => {
      const drawing = model.drawings[name];
      const drawingRel = model.drawingRels[name];
      if (drawingRel) {
        drawingOptions.rels = drawingRel.reduce((o: any, rel: any) => {
          o[rel.Id] = rel;
          return o;
        }, {});
        (drawing.anchors ?? []).forEach((anchor: any) => {
          const hyperlinks = anchor.picture && anchor.picture.hyperlinks;
          if (hyperlinks && drawingOptions.rels[hyperlinks.rId]) {
            hyperlinks.hyperlink = drawingOptions.rels[hyperlinks.rId].Target;
            delete hyperlinks.rId;
          }
        });
        drawingXform.reconcile(drawing, drawingOptions);
      }
    });

    // Reconcile chart references in drawing anchors
    Object.keys(model.drawings).forEach(name => {
      const drawing = model.drawings[name];
      const drawingRel = model.drawingRels[name];
      if (!drawingRel) {
        return;
      }
      const relMap: Record<string, any> = {};
      for (const rel of drawingRel) {
        relMap[rel.Id] = rel;
      }
      for (const anchor of drawing.anchors ?? []) {
        if (anchor.graphicFrame?.rId) {
          const rel = relMap[anchor.graphicFrame.rId];
          if (rel?.Target) {
            // Extract chart number from target like "../charts/chart1.xml"
            const match = /chart(\d+)\.xml/.exec(rel.Target);
            if (match) {
              anchor.chartNumber = parseInt(match[1], 10);
            }
            // Extract chartEx number from target like "../charts/chartEx1.xml"
            const matchEx = /chartEx(\d+)\.xml/.exec(rel.Target);
            if (matchEx) {
              anchor.chartExNumber = parseInt(matchEx[1], 10);
            }
          }
        }
      }
    });

    // reconcile tables with the default styles
    const tableOptions = {
      styles: model.styles
    };
    Object.values(model.tables).forEach((table: any) => {
      tableXform.reconcile(table, tableOptions);
    });

    // Reconcile pivot tables
    this._reconcilePivotTables(model);

    const sheetOptions = {
      styles: model.styles,
      sharedStrings: model.sharedStrings,
      media: model.media,
      mediaIndex: model.mediaIndex,
      date1904: model.properties?.date1904,
      drawings: model.drawings,
      drawingRels: model.drawingRels,
      comments: model.comments,
      tables: model.tables,
      vmlDrawings: model.vmlDrawings,
      pivotTables: model.pivotTablesIndexed,
      hasDynamicArrayMetadata: !!model.metadata?.hasDynamicArrays,
      dynamicArrayCmIndices: model.metadata?.dynamicArrayCmIndices as Set<number> | undefined
    };
    model.worksheets.forEach((worksheet: any) => {
      worksheet.relationships = model.worksheetRels[worksheet.sheetNo];
      worksheetXform.reconcile(worksheet, sheetOptions);
    });

    // Reconcile chartsheets — link their drawing references
    const chartsheetsList = model.chartsheetsList || [];
    for (const cs of chartsheetsList) {
      const csRels = model.chartsheetRels[cs.sheetNo];
      if (cs.drawing && csRels) {
        const drawingRel = csRels.find((r: any) => r.Id === cs.drawing.rId);
        if (drawingRel) {
          const match = drawingRel.Target.match(/\/drawings\/([a-zA-Z0-9]+)[.][a-zA-Z]{3,4}$/);
          if (match) {
            cs.drawingName = match[1];
            // Resolve drawing → chart number from drawing rels
            const drawingRelArr = model.drawingRels[cs.drawingName];
            if (drawingRelArr) {
              for (const dr of drawingRelArr) {
                const chartMatch = /chart(\d+)\.xml/.exec(dr.Target);
                if (chartMatch) {
                  cs.chartNumber = parseInt(chartMatch[1], 10);
                  break;
                }
                const chartExMatch = /chartEx(\d+)\.xml/.exec(dr.Target);
                if (chartExMatch) {
                  cs.chartExNumber = parseInt(chartExMatch[1], 10);
                  break;
                }
              }
            }
          }
        }
      }
    }
    model.chartsheets = chartsheetsList;

    // Reconcile external workbook links before workbookRels / externalReferences
    // are dropped. Joins 3 sources:
    //   1. model.externalReferences  — ordered list of { rId } from workbook.xml
    //   2. model.workbookRels        — maps rId → target path (inside xl/)
    //   3. model.externalLinksByIndex — parsed externalLinkN.xml parts
    //   4. model.externalLinkRelsByIndex — parsed externalLinkN.xml.rels parts
    this._reconcileExternalLinks(model);

    // Preserve parsed chart data through to the workbook model.
    // chartEntries, chartRels, chartStyles, chartColors are kept as-is.

    // delete unnecessary parts
    delete model.worksheetHash;
    delete model.worksheetRels;
    delete model.globalRels;
    delete model.sharedStrings;
    delete model.workbookRels;
    delete model.sheetDefs;
    // Preserve default font before deleting styles
    model.defaultFont = model.styles?.defaultFont;
    delete model.styles;
    delete model.mediaIndex;
    delete model.drawings;
    delete model.drawingRels;
    delete model.vmlDrawings;
    delete model.pivotTableRels;
    delete model.metadata;
    // Internal-only scratch fields consumed by _reconcileExternalLinks.
    delete model.externalReferences;
    delete model.externalLinksByIndex;
    delete model.externalLinkRelsByIndex;
    delete model.chartsheetRels;
    delete model.chartsheetsList;
  }

  /**
   * Join the three on-disk sources that together describe external workbook
   * references into a single dense `model.externalLinks: ExternalLinkModel[]`.
   *
   * Sources:
   *   - `<externalReferences>` list in workbook.xml (declaration order)
   *   - `xl/_rels/workbook.xml.rels` (rId → internal path)
   *   - `xl/externalLinks/externalLink{N}.xml` (sheet names, cached values)
   *   - `xl/externalLinks/_rels/externalLink{N}.xml.rels` (target, TargetMode)
   *
   * The 1-based index of each resulting ExternalLinkModel matches the `[N]`
   * used in formula strings — this is the single source of truth formula
   * code should rely on.
   */
  protected _reconcileExternalLinks(model: any): void {
    const refs = model.externalReferences as Array<{ rId: string }> | undefined;
    if (!refs || refs.length === 0) {
      // Even when workbook.xml has no <externalReferences>, we may still
      // have parsed externalLink parts (e.g. orphaned files); those are
      // dropped silently rather than generating synthesised indices.
      if (!model.externalLinks) {
        model.externalLinks = [];
      }
      return;
    }

    const rels = (model.workbookRels ?? []) as Array<{
      Id: string;
      Type: string;
      Target: string;
    }>;
    const relById = new Map<string, { Target: string }>();
    for (const rel of rels) {
      if (rel.Type === RelType.ExternalLink) {
        relById.set(rel.Id, rel);
      }
    }

    const externalLinks: ExternalLinkModel[] = [];
    for (let i = 0; i < refs.length; i++) {
      const ref = refs[i];
      const rel = relById.get(ref.rId);
      if (!rel) {
        // Broken reference — <externalReference> points at an rId that is
        // not of type ExternalLink. We skip silently; the formula engine
        // will see the now-missing index and fall back to #REF! as before.
        continue;
      }

      // The rel Target is a path inside xl/ like "externalLinks/externalLink1.xml".
      // Extract the trailing index to look up the parsed part.
      const partIndex = externalLinkIndexFromRelTarget(rel.Target);
      if (partIndex === undefined) {
        continue;
      }

      const parsed = model.externalLinksByIndex[partIndex] as ParsedExternalLink | undefined;
      const partRels = (model.externalLinkRelsByIndex[partIndex] ?? []) as ExternalLinkRelsEntry[];

      // Locate the externalLinkPath rel (should be unique within a part).
      const pathRel =
        partRels.find(r => r.Type === RelType.ExternalLinkPath) ??
        partRels.find(r => r.TargetMode === "External");

      externalLinks.push({
        index: i + 1,
        rId: ref.rId,
        target: pathRel?.Target ?? "",
        targetMode: (pathRel?.TargetMode as "External" | "Internal" | undefined) ?? "External",
        sheetNames: parsed?.sheetNames ?? [],
        cachedValues: parsed?.cachedValues ?? {}
      });
    }

    model.externalLinks = externalLinks;
  }

  /**
   * Write-time pass that brings the workbook model into a shape the writer
   * can serialise cleanly. Two concerns:
   *
   *   1. Build the final external-link list for this write, combining
   *      user-declared links (`wb.externalLinks`) with auto-discovered
   *      ones from previous writes (cached on the Workbook). The result
   *      is assigned to `model.externalLinks` and consumed by the writer;
   *      `wb.externalLinks` is **not** modified.
   *
   *   2. Scan every formula cell for `[Book]Sheet!` prefixes. Filename-form
   *      references that don't match an existing link trigger
   *      `_recordAutoExternalLink()` on the Workbook, which adds the target
   *      to the private writer cache (so subsequent writes are fixed-point
   *      stable) but leaves `wb.externalLinks` untouched.
   *
   *   3. Rewrite every external-ref formula so it uses the numeric `[N]`
   *      form, the canonical OOXML storage form. This mutation lands on
   *      the cell's model object — matching the library's existing
   *      write-time pattern for `ssId`, `styleId`, `si`, and `cm`.
   *      Subsequent writes see the `[N]` form directly and resolve it
   *      against `model.externalLinks`, giving idempotent output.
   */
  protected _normaliseExternalLinks(model: any): void {
    // Start from user-declared links, honouring their declaration order.
    const links = this.workbook._collectExternalLinksForWrite();

    // Fast lookup: case-insensitive target → link object in `links`.
    const byTarget = new Map<string, ExternalLinkModel>();
    for (const link of links) {
      if (link.target) {
        byTarget.set(link.target.toLowerCase(), link);
      }
    }

    const scratch: NormaliseScratch = { links, byTarget, workbook: this.workbook };
    for (const worksheet of model.worksheets ?? []) {
      for (const row of worksheet.rows ?? []) {
        for (const cell of row.cells ?? []) {
          if (typeof cell?.formula === "string" && cell.formula.length > 0) {
            cell.formula = this._normaliseFormulaExternalRefs(cell.formula, scratch);
          }
          if (typeof cell?.sharedFormula === "string" && cell.sharedFormula.length > 0) {
            // Shared-formula clones typically carry the master's *address*
            // here, not a formula body — they won't match the ref regex.
            // Masters carry the formula on `.formula` (handled above).
            // We rewrite defensively in case a caller stored an actual
            // formula string here.
            cell.sharedFormula = this._normaliseFormulaExternalRefs(cell.sharedFormula, scratch);
          }
        }
      }
    }

    model.externalLinks = links;
  }

  /**
   * Rewrite a single formula so every external-ref prefix uses the numeric
   * `[N]` form. When an unknown filename-form reference is found we record
   * it on the workbook's private writer cache (so the next write can still
   * resolve it) and append a local link to `scratch.links` so subsequent
   * refs in the same formula see the freshly-assigned index.
   */
  private _normaliseFormulaExternalRefs(formula: string, scratch: NormaliseScratch): string {
    // rewriteExternalRefs internally calls findExternalRefs and returns
    // the original string unchanged when there are no matches — no need
    // for a separate guard scan here.
    return rewriteExternalRefs(formula, ref => {
      // Numeric ref: accept if it resolves, otherwise preserve verbatim so
      // Excel surfaces `#REF!` at load time — same as the old behaviour
      // for truly broken references.
      if (ref.numeric) {
        if (ref.index !== null && ref.index >= 1 && ref.index <= scratch.links.length) {
          upsertSheet(scratch.links[ref.index - 1], ref.sheet);
          return ref.index;
        }
        return null;
      }

      // Filename form — look up or auto-register.
      const key = ref.workbook.toLowerCase();
      let link = scratch.byTarget.get(key);
      if (!link) {
        const index = scratch.workbook._recordAutoExternalLink(ref.workbook, ref.sheet);
        link = {
          index,
          target: ref.workbook,
          targetMode: "External",
          sheetNames: ref.sheet ? [ref.sheet] : [],
          cachedValues: {}
        };
        // Keep the local writer list dense: insert at its future position.
        // `_recordAutoExternalLink` guarantees `index` is user.length + cache.size
        // at the time of insertion, which equals `scratch.links.length + 1`
        // whenever we walk formulas sequentially.
        scratch.links.push(link);
        scratch.byTarget.set(key, link);
      } else {
        upsertSheet(link, ref.sheet);
        // Keep the workbook cache's sheetNames in sync so subsequent
        // writes see the accumulated set.
        if (ref.sheet) {
          scratch.workbook._recordAutoExternalLink(ref.workbook, ref.sheet);
        }
      }
      return link.index;
    });
  }

  /**
   * Reconcile pivot tables by linking them to worksheets and their cache data.
   */
  protected _reconcilePivotTables(model: any): void {
    const rawPivotTables = (model.pivotTables || {}) as Record<string, ParsedPivotTableModel>;
    if (typeof rawPivotTables !== "object" || Object.keys(rawPivotTables).length === 0) {
      model.pivotTables = [];
      model.pivotTablesIndexed = {};
      return;
    }

    const definitionToCacheId = this._buildDefinitionToCacheIdMap(model);

    const cacheMap = new Map<
      number,
      {
        definition: ParsedCacheDefinition;
        records: any;
        definitionName: string;
      }
    >();

    Object.entries(model.pivotCacheDefinitions || {}).forEach(
      ([name, definition]: [string, any]) => {
        const cacheId = definitionToCacheId.get(name);
        if (cacheId !== undefined) {
          const recordsName = name.replace("Definition", "Records");
          cacheMap.set(cacheId, {
            definition,
            records: model.pivotCacheRecords?.[recordsName],
            definitionName: name
          });
        }
      }
    );

    const loadedPivotTables: PivotTable[] = [];
    const pivotTablesIndexed: Record<string, PivotTable> = {};

    Object.entries(rawPivotTables).forEach(([pivotName, pt]) => {
      const tableNumber = this._extractTableNumber(pivotName);
      const cacheData = cacheMap.get(pt.cacheId);

      const defaultMetric = this._determineMetric(pt.dataFields);
      const completePivotTable: PivotTable = {
        ...pt,
        tableNumber,
        cacheId: String(pt.cacheId),
        cacheDefinition: cacheData?.definition,
        cacheRecords: cacheData?.records,
        cacheFields: cacheData?.definition?.cacheFields ?? [],
        rows: pt.rowFields.filter(f => f >= 0),
        columns: pt.colFields.filter(f => f >= 0 && f !== -2),
        values: pt.dataFields.map(df => df.fld),
        pages: pt.pageFields.map(pf => pf.fld),
        metric: defaultMetric,
        valueMetrics: this._determineValueMetrics(pt.dataFields, defaultMetric),
        applyWidthHeightFormats: pt.applyWidthHeightFormats === "1" ? "1" : "0"
      };

      loadedPivotTables.push(completePivotTable);
      // Key by absolute zip path so reconcile can match any rel target layout.
      pivotTablesIndexed[pivotTablePathFromName(pivotName)] = completePivotTable;
    });

    loadedPivotTables.sort((a, b) => a.tableNumber - b.tableNumber);
    model.pivotTables = loadedPivotTables;
    model.pivotTablesIndexed = pivotTablesIndexed;
  }

  protected _extractTableNumber(name: string): number {
    const match = name.match(/pivotTable(\d+)/);
    return match ? parseInt(match[1], 10) : 1;
  }

  protected _buildCacheIdMap(model: any): Map<string, number> {
    const rIdToCacheId = new Map<string, number>();
    const pivotCaches = model.pivotCaches ?? [];
    for (const cache of pivotCaches) {
      if (cache.cacheId && cache.rId) {
        rIdToCacheId.set(cache.rId, parseInt(cache.cacheId, 10));
      }
    }
    return rIdToCacheId;
  }

  protected _buildDefinitionToCacheIdMap(model: any): Map<string, number> {
    const definitionToCacheId = new Map<string, number>();
    const rIdToCacheId = this._buildCacheIdMap(model);
    const workbookRels = model.workbookRels ?? [];

    for (const rel of workbookRels) {
      if (rel.Type === XLSX.RelType.PivotCacheDefinition && rel.Target) {
        const match = rel.Target.match(/pivotCacheDefinition(\d+)\.xml/);
        if (match) {
          const defName = `pivotCacheDefinition${match[1]}`;
          const cacheId = rIdToCacheId.get(rel.Id);
          if (cacheId !== undefined) {
            definitionToCacheId.set(defName, cacheId);
          }
        }
      }
    }

    return definitionToCacheId;
  }

  protected _determineMetric(dataFields: Array<{ subtotal?: string }>): PivotTableSubtotal {
    if (dataFields.length > 0 && dataFields[0].subtotal) {
      return dataFields[0].subtotal as PivotTableSubtotal;
    }
    return "sum";
  }

  protected _determineValueMetrics(
    dataFields: Array<{ subtotal?: string }>,
    defaultMetric: PivotTableSubtotal
  ): PivotTableSubtotal[] {
    return dataFields.map(df => (df.subtotal as PivotTableSubtotal) || defaultMetric);
  }

  // ===========================================================================
  // Process Entry methods - shared by all platforms
  // ===========================================================================

  async _processWorksheetEntry(
    stream: IParseStream,
    model: any,
    sheetNo: number,
    options: XlsxOptions | undefined,
    path: string
  ): Promise<void> {
    const xform = new WorkSheetXform(options);
    const worksheet = await xform.parseStream(stream);
    if (!worksheet) {
      throw new XmlParseError(path, "Failed to parse worksheet");
    }
    worksheet.sheetNo = sheetNo;
    model.worksheetHash[path] = worksheet;
    model.worksheets.push(worksheet);
  }

  async _processChartsheetEntry(stream: IParseStream, model: any, sheetNo: number): Promise<void> {
    const xform = new ChartsheetXform();
    const chartsheet = await xform.parseStream(stream);
    if (chartsheet) {
      chartsheet.sheetNo = sheetNo;
      model.chartsheets[sheetNo] = chartsheet;
    }
  }

  async _processCommentEntry(stream: IParseStream, model: any, zipPath: string): Promise<void> {
    const xform = new CommentsXform();
    const comments = await xform.parseStream(stream);
    // Key by absolute zip path so reconcile can match any rel target layout.
    model.comments[zipPath] = comments;
  }

  async _processTableEntry(stream: IParseStream, model: any, zipPath: string): Promise<void> {
    const xform = new TableXform();
    const table = await xform.parseStream(stream);
    // Key by absolute zip path so reconcile can match any rel target layout.
    model.tables[zipPath] = table;
  }

  async _processWorksheetRelsEntry(
    stream: IParseStream,
    model: any,
    sheetNo: number
  ): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(stream);
    model.worksheetRels[sheetNo] = relationships;
  }

  async _processMediaEntry(stream: IParseStream, model: any, filename: string): Promise<void> {
    const lastDot = filename.lastIndexOf(".");
    if (lastDot >= 1) {
      const extension = filename.substr(lastDot + 1);
      const name = filename.substr(0, lastDot);
      await new Promise<void>((resolve, reject) => {
        const streamBuf = this.createStreamBuf();

        const cleanup = () => {
          stream.off("error", onError);
          streamBuf.off("error", onError);
          streamBuf.off("finish", onFinish);
        };

        const onFinish = () => {
          cleanup();
          model.mediaIndex[filename] = model.media.length;
          model.mediaIndex[name] = model.media.length;
          const medium = {
            type: "image",
            name,
            extension,
            buffer: streamBuf.read()
          };
          model.media.push(medium);
          resolve();
        };

        const onError = (error: Error) => {
          cleanup();
          reject(error);
        };

        streamBuf.once("finish", onFinish);
        stream.on("error", onError);
        streamBuf.on("error", onError);
        stream.pipe(streamBuf);
      });
    }
  }

  /**
   * Process a drawing XML entry.
   *
   * @param stream - Stream to read from (used in loadFromZipEntries path)
   * @param model - Model to populate
   * @param name - Drawing name (e.g., "drawing1")
   * @param rawData - Pre-read raw data (used in loadFromFiles path to avoid re-reading stream)
   */
  async _processDrawingEntry(
    stream: IParseStream,
    model: any,
    name: string,
    rawData?: Uint8Array
  ): Promise<void> {
    // Use provided rawData if available (loadFromFiles path), otherwise collect from stream.
    // In loadFromFiles, the stream is created from already-decoded text, and collecting from
    // it may not work correctly due to PassThrough stream timing issues.
    const data = rawData ?? (await this.collectStreamData(stream));

    // Parse the drawing for normal processing (images, etc.)
    const xform = new DrawingXform();
    const xmlString = this.bufferToString(data);
    const drawing = await xform.parseStream(this.createTextStream(xmlString));
    model.drawings[name] = drawing;
  }

  async _processDrawingRelsEntry(entry: any, model: any, name: string): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(entry);
    model.drawingRels[name] = relationships;
  }

  async _processVmlDrawingEntry(entry: any, model: any, zipPath: string): Promise<void> {
    const xform = new VmlDrawingXform();
    const vmlDrawing = await xform.parseStream(entry);
    // Key by absolute zip path so reconcile can match any rel target layout.
    model.vmlDrawings[zipPath] = vmlDrawing;
  }

  async _processVmlDrawingHFEntry(entry: any, model: any, _name: string): Promise<void> {
    const xform = new VmlDrawingXform();
    const vmlDrawing = await xform.parseStream(entry);
    // Store parsed header image info for reconciliation
    if (vmlDrawing && vmlDrawing.headerImage) {
      if (!model.vmlDrawingHF) {
        model.vmlDrawingHF = {};
      }
      model.vmlDrawingHF[_name] = vmlDrawing.headerImage;
    }
  }

  async _processThemeEntry(stream: IParseStream, model: any, name: string): Promise<void> {
    await new Promise<void>((resolve, reject) => {
      const streamBuf = this.createStreamBuf();

      const cleanup = () => {
        stream.off("error", onError);
        streamBuf.off("error", onError);
        streamBuf.off("finish", onFinish);
      };

      const onFinish = () => {
        cleanup();
        const data = streamBuf.read();
        model.themes[name] = data
          ? typeof data === "string"
            ? data
            : this.bufferToString(data)
          : "";
        resolve();
      };

      const onError = (err: Error) => {
        cleanup();
        reject(err);
      };

      streamBuf.once("finish", onFinish);
      stream.on("error", onError);
      streamBuf.on("error", onError);
      stream.pipe(streamBuf);
    });
  }

  async _processPivotTableEntry(stream: IParseStream, model: any, name: string): Promise<void> {
    const xform = new PivotTableXform();
    const pivotTable = await xform.parseStream(stream);
    if (pivotTable) {
      model.pivotTables[name] = pivotTable;
    }
  }

  async _processPivotTableRelsEntry(stream: IParseStream, model: any, name: string): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(stream);
    model.pivotTableRels[name] = relationships;
  }

  async _processPivotCacheDefinitionEntry(
    stream: IParseStream,
    model: any,
    name: string
  ): Promise<void> {
    const xform = new PivotCacheDefinitionXform();
    const cacheDefinition = await xform.parseStream(stream);
    if (cacheDefinition) {
      model.pivotCacheDefinitions[name] = cacheDefinition;
    }
  }

  async _processPivotCacheRecordsEntry(
    stream: IParseStream,
    model: any,
    name: string
  ): Promise<void> {
    const xform = new PivotCacheRecordsXform();
    const cacheRecords = await xform.parseStream(stream);
    if (cacheRecords) {
      model.pivotCacheRecords[name] = cacheRecords;
    }
  }

  /**
   * Parse `xl/externalLinks/externalLink{N}.xml` into the intermediate
   * ParsedExternalLink shape. Reconciliation (joining with the rels file
   * and the workbook's `<externalReferences>` list) happens later in
   * {@link reconcile}.
   */
  async _processExternalLinkEntry(stream: IParseStream, model: any, index: number): Promise<void> {
    const xform = new ExternalLinkXform();
    const parsed = await xform.parseStream(stream);
    if (parsed) {
      model.externalLinksByIndex[index] = parsed;
    }
  }

  /**
   * Parse `xl/externalLinks/_rels/externalLink{N}.xml.rels`. The Target /
   * TargetMode carried here is what Excel uses to locate the actual external
   * file at open time, so we must preserve it verbatim (including relative
   * paths like `"测试.xlsx"`).
   */
  async _processExternalLinkRelsEntry(
    stream: IParseStream,
    model: any,
    index: number
  ): Promise<void> {
    const relationships = await this.parseRels(stream);
    model.externalLinkRelsByIndex[index] = relationships ?? [];
  }

  async _processChartEntry(
    stream: IParseStream,
    model: any,
    chartNumber: number,
    rawData?: Uint8Array
  ): Promise<void> {
    const data = rawData ?? (await this.collectStreamData(stream));

    // Parse into model for high-level API access
    const xform = new ChartSpaceXform();
    const xmlString = this.bufferToString(data);
    const chart = await xform.parseStream(this.createTextStream(xmlString));
    if (chart) {
      model.chartEntries[chartNumber] = {
        chartNumber,
        model: chart
      };
    }
  }

  async _processChartRelsEntry(
    stream: IParseStream,
    model: any,
    chartNumber: number
  ): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(stream);
    model.chartRels[chartNumber] = relationships;
  }

  async _processChartStyleEntry(
    stream: IParseStream,
    model: any,
    styleNumber: number
  ): Promise<void> {
    const data = await this.collectStreamData(stream);
    model.chartStyles[styleNumber] = data;
  }

  async _processChartColorsEntry(
    stream: IParseStream,
    model: any,
    colorsNumber: number
  ): Promise<void> {
    const data = await this.collectStreamData(stream);
    model.chartColors[colorsNumber] = data;
  }

  // ===========================================================================
  // loadFromFiles - shared logic for loading from pre-extracted ZIP data
  // ===========================================================================

  async loadFromFiles(
    zipData: Record<string, Uint8Array>,
    options?: XlsxReadOptions
  ): Promise<Workbook> {
    const model: any = this.createEmptyModel();

    const entries = Object.keys(zipData).map(name => ({
      name,
      dir: name.endsWith("/"),
      data: zipData[name]
    }));

    for (const entry of entries) {
      if (!entry.dir) {
        const entryName = normalizeZipPath(entry.name);

        // Create appropriate stream based on entry type
        const isBinaryEntry = isBinaryEntryPath(entryName);
        const stream = isBinaryEntry
          ? this.createBinaryStream(entry.data)
          : this.createTextStream(this.bufferToString(entry.data));

        const handled = await this._processKnownEntry(stream, model, entryName, options);
        if (!handled) {
          await this._processDefaultEntry(stream, model, entryName, entry.data);
        }
      }
    }

    this.reconcile(model, options);
    this.workbook.model = model;
    return this.workbook;
  }

  /**
   * Process default entries (drawings, comments, tables, etc.)
   * @param rawData Optional raw entry data for passthrough preservation (used by loadFromFiles)
   */
  protected async _processDefaultEntry(
    stream: IParseStream,
    model: any,
    entryName: string,
    rawData?: Uint8Array
  ): Promise<boolean> {
    const sheetNo = getWorksheetNoFromWorksheetRelsPath(entryName);
    if (sheetNo !== undefined) {
      await this._processWorksheetRelsEntry(stream, model, sheetNo);
      return true;
    }

    const chartsheetRelsNo = getChartsheetNoFromRelsPath(entryName);
    if (chartsheetRelsNo !== undefined) {
      const rels = await this.parseRels(stream);
      model.chartsheetRels[chartsheetRelsNo] = rels;
      return true;
    }

    const mediaFilename = getMediaFilenameFromPath(entryName);
    if (mediaFilename) {
      await this._processMediaEntry(stream, model, mediaFilename);
      return true;
    }

    const drawingName = getDrawingNameFromPath(entryName);
    if (drawingName) {
      await this._processDrawingEntry(stream, model, drawingName, rawData);
      return true;
    }

    const drawingRelsName = getDrawingNameFromRelsPath(entryName);
    if (drawingRelsName) {
      await this._processDrawingRelsEntry(stream, model, drawingRelsName);
      return true;
    }

    const vmlDrawingName = getVmlDrawingNameFromPath(entryName);
    if (vmlDrawingName) {
      await this._processVmlDrawingEntry(stream, model, entryName);
      return true;
    }

    // VML header/footer drawings (watermark in header mode).
    // Parse to extract header image info for round-trip preservation.
    const vmlHFName = getVmlDrawingHFNameFromPath(entryName);
    if (vmlHFName) {
      await this._processVmlDrawingHFEntry(stream, model, vmlHFName);
      return true;
    }

    if (isCommentsPath(entryName)) {
      await this._processCommentEntry(stream, model, entryName);
      return true;
    }

    const tableName = getTableNameFromPath(entryName);
    if (tableName) {
      await this._processTableEntry(stream, model, entryName);
      return true;
    }

    const themeName = getThemeNameFromPath(entryName);
    if (themeName) {
      await this._processThemeEntry(stream, model, themeName);
      return true;
    }

    // Pivot table files
    const pivotTableName = getPivotTableNameFromPath(entryName);
    if (pivotTableName) {
      await this._processPivotTableEntry(stream, model, pivotTableName);
      return true;
    }

    const pivotTableRelsName = getPivotTableNameFromRelsPath(entryName);
    if (pivotTableRelsName) {
      await this._processPivotTableRelsEntry(stream, model, pivotTableRelsName);
      return true;
    }

    // Pivot cache files
    const pivotCacheDefinitionName = getPivotCacheDefinitionNameFromPath(entryName);
    if (pivotCacheDefinitionName) {
      await this._processPivotCacheDefinitionEntry(stream, model, pivotCacheDefinitionName);
      return true;
    }

    // R9-B8: Skip parsing pivotCacheDefinition .rels files — they are never used
    // during reconciliation and were just deleted at cleanup. The cache definition's
    // r:id attribute (preserved in ParsedCacheDefinition.rId) is sufficient.
    const pivotCacheDefinitionRelsName = getPivotCacheDefinitionNameFromRelsPath(entryName);
    if (pivotCacheDefinitionRelsName) {
      return true;
    }

    const pivotCacheRecordsName = getPivotCacheRecordsNameFromPath(entryName);
    if (pivotCacheRecordsName) {
      await this._processPivotCacheRecordsEntry(stream, model, pivotCacheRecordsName);
      return true;
    }

    // External workbook links: xl/externalLinks/externalLinkN.xml and its
    // sibling _rels file. Both parts are required to reconstruct the
    // ExternalLinkModel (the .xml carries sheet names + cached values; the
    // .rels carries the target path and TargetMode).
    const externalLinkIndex = getExternalLinkIndexFromPath(entryName);
    if (externalLinkIndex !== undefined) {
      await this._processExternalLinkEntry(stream, model, externalLinkIndex);
      return true;
    }

    const externalLinkRelsIndex = getExternalLinkIndexFromRelsPath(entryName);
    if (externalLinkRelsIndex !== undefined) {
      await this._processExternalLinkRelsEntry(stream, model, externalLinkRelsIndex);
      return true;
    }

    // Chart files — parse natively before the passthrough catch-all
    const chartNumber = getChartNumberFromPath(entryName);
    if (chartNumber !== undefined) {
      await this._processChartEntry(stream, model, chartNumber, rawData);
      return true;
    }

    const chartRelsNumber = getChartNumberFromRelsPath(entryName);
    if (chartRelsNumber !== undefined) {
      await this._processChartRelsEntry(stream, model, chartRelsNumber);
      return true;
    }

    const chartStyleNumber = getChartStyleNumberFromPath(entryName);
    if (chartStyleNumber !== undefined) {
      if (rawData) {
        model.chartStyles[chartStyleNumber] = rawData;
      } else {
        await this._processChartStyleEntry(stream, model, chartStyleNumber);
      }
      return true;
    }

    const chartColorsNumber = getChartColorsNumberFromPath(entryName);
    if (chartColorsNumber !== undefined) {
      if (rawData) {
        model.chartColors[chartColorsNumber] = rawData;
      } else {
        await this._processChartColorsEntry(stream, model, chartColorsNumber);
      }
      return true;
    }

    // ChartEx files (Office 2016+ extended charts) — stored as raw bytes
    const chartExNumber = getChartExNumberFromPath(entryName);
    if (chartExNumber !== undefined) {
      if (rawData) {
        model.chartExEntries[chartExNumber] = rawData;
      } else {
        model.chartExEntries[chartExNumber] = await this.collectStreamData(stream);
      }
      return true;
    }

    const chartExRelsNumber = getChartExNumberFromRelsPath(entryName);
    if (chartExRelsNumber !== undefined) {
      const relsXform = new RelationshipsXform();
      const relationships = await relsXform.parseStream(stream);
      model.chartExRels[chartExRelsNumber] = relationships;
      return true;
    }

    return false;
  }

  // ===========================================================================
  // Write methods - shared by all platforms
  // ===========================================================================

  /**
   * Helper: render an xform directly to a streaming zip entry.
   * Avoids buffering the entire XML string in memory.
   * Awaits backpressure drain after each entry to respect downstream flow control.
   */
  private async _renderToZip(
    zip: IZipWriter,
    path: string,
    xform: { render(xmlStream: any, model?: any): void },
    model?: any
  ): Promise<void> {
    const entry = zip.createEntry(path);
    const stream = new XmlStreamWriter(entry);
    xform.render(stream, model);
    entry.end();
    // Respect downstream backpressure between entries
    await zip.waitForDrain();
  }

  async addContentTypes(zip: IZipWriter, model: any): Promise<void> {
    await this._renderToZip(zip, OOXML_PATHS.contentTypes, new ContentTypesXform(), model);
  }

  async addApp(zip: IZipWriter, model: any): Promise<void> {
    await this._renderToZip(zip, OOXML_PATHS.docPropsApp, new AppXform(), model);
  }

  async addCore(zip: IZipWriter, model: any): Promise<void> {
    await this._renderToZip(zip, OOXML_PATHS.docPropsCore, new CoreXform(), model);
  }

  async addThemes(zip: IZipWriter, model: any): Promise<void> {
    const themes = model.themes || { theme1: theme1Xml };
    Object.keys(themes).forEach(name => {
      const xml = themes[name];
      zip.append(xml, { name: themePath(name) });
    });
  }

  async addOfficeRels(zip: IZipWriter, _model: any): Promise<void> {
    await this._renderToZip(zip, OOXML_PATHS.rootRels, new RelationshipsXform(), [
      { Id: "rId1", Type: XLSX.RelType.OfficeDocument, Target: OOXML_PATHS.xlWorkbook },
      { Id: "rId2", Type: XLSX.RelType.CoreProperties, Target: OOXML_PATHS.docPropsCore },
      { Id: "rId3", Type: XLSX.RelType.ExtenderProperties, Target: OOXML_PATHS.docPropsApp }
    ]);
  }

  async addWorkbookRels(zip: IZipWriter, model: any): Promise<void> {
    let count = 1;
    const relationships: any[] = [
      { Id: `rId${count++}`, Type: XLSX.RelType.Styles, Target: OOXML_REL_TARGETS.workbookStyles },
      { Id: `rId${count++}`, Type: XLSX.RelType.Theme, Target: OOXML_REL_TARGETS.workbookTheme1 }
    ];
    if (model.sharedStrings.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: XLSX.RelType.SharedStrings,
        Target: OOXML_REL_TARGETS.workbookSharedStrings
      });
    }

    // Add FeaturePropertyBag relationship if checkboxes are used
    if (model.hasCheckboxes) {
      relationships.push({
        Id: `rId${count++}`,
        Type: XLSX.RelType.FeaturePropertyBag,
        Target: OOXML_REL_TARGETS.workbookFeaturePropertyBag
      });
    }
    // Add metadata relationship for dynamic array formulas
    if (model.hasDynamicArrayFormulas) {
      relationships.push({
        Id: `rId${count++}`,
        Type: XLSX.RelType.SheetMetadata,
        Target: OOXML_REL_TARGETS.workbookMetadata
      });
    }
    // R9-B6: Deduplicate pivot cache relationships by cacheId. When multiple pivot
    // tables share the same cache, only one workbook relationship should be created.
    // Also assigns rId to each pivot table (R9-B7: typed on PivotTable interface).
    const seenCacheIds = new Map<string, string>(); // cacheId → rId
    (model.pivotTables ?? []).forEach((pivotTable: PivotTable) => {
      const existing = seenCacheIds.get(pivotTable.cacheId);
      if (existing) {
        // Shared cache: reuse the rId from the first pivot table with this cacheId
        pivotTable.rId = existing;
      } else {
        pivotTable.rId = `rId${count++}`;
        seenCacheIds.set(pivotTable.cacheId, pivotTable.rId);
        relationships.push({
          Id: pivotTable.rId,
          Type: XLSX.RelType.PivotCacheDefinition,
          Target: pivotCacheDefinitionRelTargetFromWorkbook(pivotTable.tableNumber)
        });
      }
    });
    model.worksheets.forEach((worksheet: any) => {
      worksheet.rId = `rId${count++}`;
      // fileIndex is assigned once in prepareModel() — use it directly
      relationships.push({
        Id: worksheet.rId,
        Type: XLSX.RelType.Worksheet,
        Target: worksheetRelTarget(worksheet.fileIndex)
      });
    });

    // Add chartsheet relationships
    (model.chartsheets || []).forEach((cs: any) => {
      cs.rId = `rId${count++}`;
      relationships.push({
        Id: cs.rId,
        Type: RelType.Chartsheet,
        Target: `chartsheets/sheet${cs.sheetNo}.xml`
      });
    });

    // External workbook link rels are written AFTER worksheets on purpose:
    // Excel tolerates either order, but stable ordering (worksheets then
    // externalLinks) keeps the emitted workbook.xml.rels diff-friendly for
    // round-trip tests. Each external link becomes a regular Relationship
    // entry targeting `externalLinks/externalLinkN.xml` inside `xl/`; the
    // actual external file path is pointed at by the nested
    // externalLinkN.xml.rels part written by addExternalLinks().
    //
    // The list items here are the deep-copies produced by
    // `_normaliseExternalLinks` — assigning `link.rId` is safe and does not
    // leak into the user's Workbook.externalLinks.
    const externalLinks = (model.externalLinks ?? []) as ExternalLinkModel[];
    for (const link of externalLinks) {
      link.rId = `rId${count++}`;
      relationships.push({
        Id: link.rId,
        Type: XLSX.RelType.ExternalLink,
        Target: externalLinkRelTargetFromWorkbook(link.index)
      });
    }

    const xform = new RelationshipsXform();
    await this._renderToZip(zip, OOXML_PATHS.xlWorkbookRels, xform, relationships);
  }

  async addFeaturePropertyBag(zip: IZipWriter, model: any): Promise<void> {
    if (!model.hasCheckboxes) {
      return;
    }
    await this._renderToZip(
      zip,
      OOXML_PATHS.xlFeaturePropertyBag,
      new FeaturePropertyBagXform(),
      {}
    );
  }

  async addMetadata(zip: IZipWriter, model: any): Promise<void> {
    if (!model.hasDynamicArrayFormulas) {
      return;
    }
    await this._renderToZip(zip, OOXML_PATHS.xlMetadata, new MetadataXform(), {
      dynamicArrayCount: model.dynamicArrayCount
    });
  }

  async addSharedStrings(zip: IZipWriter, model: any): Promise<void> {
    if (model.sharedStrings && model.sharedStrings.count) {
      await this._renderToZip(
        zip,
        OOXML_PATHS.xlSharedStrings,
        model.sharedStrings,
        model.sharedStrings.model
      );
    }
  }

  async addStyles(zip: IZipWriter, model: any): Promise<void> {
    if (model.styles) {
      await this._renderToZip(zip, OOXML_PATHS.xlStyles, model.styles, model.styles.model);
    }
  }

  async addWorkbook(zip: IZipWriter, model: any): Promise<void> {
    await this._renderToZip(zip, OOXML_PATHS.xlWorkbook, new WorkbookXform(), model);
  }

  async addWorksheets(zip: IZipWriter, model: any): Promise<void> {
    const worksheetXform = new WorkSheetXform();
    const relationshipsXform = new RelationshipsXform();
    const commentsXform = new CommentsXform();
    const vmlDrawingXform = new VmlDrawingXform();
    const ctrlPropXform = new CtrlPropXform();

    for (const worksheet of model.worksheets) {
      const { fileIndex } = worksheet;

      // Worksheet XML: stream directly to zip entry (avoids holding entire XML in memory)
      const wsEntry = zip.createEntry(worksheetPath(fileIndex));
      const wsStream = new XmlStreamWriter(wsEntry);
      worksheetXform.render(wsStream, worksheet);
      wsEntry.end();
      await zip.waitForDrain();

      if (worksheet.rels && worksheet.rels.length) {
        await this._renderToZip(
          zip,
          worksheetRelsPath(fileIndex),
          relationshipsXform,
          worksheet.rels
        );
      }

      // Generate comments XML (separate from VML)
      if (worksheet.comments.length > 0) {
        await this._renderToZip(zip, commentsPath(fileIndex), commentsXform, worksheet);
      }

      // Generate unified VML drawing (contains both notes and form controls)
      const hasComments = worksheet.comments.length > 0;
      const hasFormControls = worksheet.formControls && worksheet.formControls.length > 0;

      if (hasComments || hasFormControls) {
        await this._renderToZip(zip, vmlDrawingPath(fileIndex), vmlDrawingXform, {
          comments: hasComments ? worksheet.comments : [],
          formControls: hasFormControls ? worksheet.formControls : []
        });
      }

      // Generate VML drawing for header/footer images (watermark in header mode)
      if (worksheet.headerImage) {
        const hdrImage = worksheet.headerImage;
        const bookImage = hdrImage.bookImage;
        const imageFileName =
          bookImage.name &&
          bookImage.extension &&
          bookImage.name.endsWith(`.${bookImage.extension}`)
            ? bookImage.name
            : `${bookImage.name}.${bookImage.extension}`;
        const imageRelTarget = `../media/${imageFileName}`;

        // Write the VML file for the header image
        await this._renderToZip(zip, vmlDrawingHFPath(fileIndex), vmlDrawingXform, {
          comments: [],
          formControls: [],
          headerImage: {
            imageRelId: "rId1",
            width: hdrImage.headerWidth,
            height: hdrImage.headerHeight
          }
        });

        // Write the VML rels file referencing the image
        await this._renderToZip(zip, vmlDrawingHFRelsPath(fileIndex), relationshipsXform, [
          {
            Id: "rId1",
            Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            Target: imageRelTarget
          }
        ]);
      }

      // Generate ctrlProp files for form controls
      if (hasFormControls) {
        for (const control of worksheet.formControls) {
          await this._renderToZip(zip, ctrlPropPath(control.ctrlPropId), ctrlPropXform, control);
        }
      }
    }
  }

  async addChartsheets(zip: IZipWriter, model: any): Promise<void> {
    const chartsheetXform = new ChartsheetXform();
    const relsXform = new RelationshipsXform();

    for (const cs of model.chartsheets || []) {
      await this._renderToZip(zip, chartsheetPath(cs.sheetNo), chartsheetXform, cs);

      // Write chartsheet rels if there's a drawing
      if (cs.drawing) {
        const rels = [
          {
            Id: cs.drawing.rId,
            Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            Target: `../drawings/${cs.drawingName}.xml`
          }
        ];
        await this._renderToZip(zip, chartsheetRelsPath(cs.sheetNo), relsXform, rels);
      }
    }
  }

  async addDrawings(zip: IZipWriter, model: any): Promise<void> {
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();

    for (const worksheet of model.worksheets) {
      const { drawing } = worksheet;
      if (drawing) {
        const filteredAnchors = filterDrawingAnchors(drawing.anchors ?? []);
        const drawingForWrite = drawing.anchors
          ? { ...drawing, anchors: filteredAnchors }
          : drawing;
        drawingXform.prepare(drawingForWrite);
        await this._renderToZip(zip, drawingPath(drawing.name), drawingXform, drawingForWrite);

        await this._renderToZip(zip, drawingRelsPath(drawing.name), relsXform, drawing.rels);
      }
    }

    // Chartsheet drawings — each chartsheet references a drawing containing a chart
    for (const cs of model.chartsheets || []) {
      if (cs.drawingName && (cs.chartNumber || cs.chartExNumber)) {
        const chartRId = "rId1";
        const isChartEx = !cs.chartNumber && !!cs.chartExNumber;
        const drawingModel = {
          anchors: [
            {
              range: { tl: { col: 0, row: 0 }, br: { col: 10, row: 15 } },
              graphicFrame: {
                rId: chartRId,
                isChartEx,
                name: isChartEx ? `Chart ${cs.chartExNumber}` : `Chart ${cs.chartNumber}`
              },
              ...(isChartEx ? { alternateContent: { requires: "cx" } } : {})
            }
          ]
        };
        const drawingRels = [
          {
            Id: chartRId,
            Type: isChartEx ? RelType.ChartEx : RelType.Chart,
            Target: isChartEx
              ? chartExRelTargetFromDrawing(cs.chartExNumber)
              : chartRelTargetFromDrawing(cs.chartNumber)
          }
        ];
        drawingXform.prepare(drawingModel);
        await this._renderToZip(zip, drawingPath(cs.drawingName), drawingXform, drawingModel);
        await this._renderToZip(zip, drawingRelsPath(cs.drawingName), relsXform, drawingRels);
      }
    }
  }

  async addCharts(zip: IZipWriter, model: any): Promise<void> {
    const relsXform = new RelationshipsXform();

    for (const [n, chartEntry] of Object.entries(model.chartEntries || {})) {
      // Write chart XML — fully native parse→render
      await this._renderToZip(zip, chartPath(n), new ChartSpaceXform(), (chartEntry as any).model);

      // Write chart style (raw bytes)
      if (model.chartStyles?.[n]) {
        zip.append(model.chartStyles[n], { name: chartStylePath(n) });
      }

      // Write chart colors (raw bytes)
      if (model.chartColors?.[n]) {
        zip.append(model.chartColors[n], { name: chartColorsPath(n) });
      }

      // Build chart rels
      const rels: any[] = [];

      // Collect original rels first (excluding style/colors which we regenerate)
      // We keep their original Ids to avoid breaking r:id references inside chart XML
      const originalRels = model.chartRels?.[n];
      const usedIds = new Set<string>();
      if (Array.isArray(originalRels)) {
        for (const rel of originalRels) {
          if (rel.Type !== RelType.ChartStyle && rel.Type !== RelType.ChartColors) {
            rels.push(rel);
            usedIds.add(rel.Id);
          }
        }
      }

      // Allocate new rIds for style/colors that don't conflict with existing ones
      let rIdCount = 1;
      const nextRId = (): string => {
        let id = `rId${rIdCount++}`;
        while (usedIds.has(id)) {
          id = `rId${rIdCount++}`;
        }
        usedIds.add(id);
        return id;
      };

      // Add style rel if style exists
      if (model.chartStyles?.[n]) {
        rels.push({
          Id: nextRId(),
          Type: RelType.ChartStyle,
          Target: chartStyleRelTarget(n)
        });
      }

      // Add colors rel if colors exist
      if (model.chartColors?.[n]) {
        rels.push({
          Id: nextRId(),
          Type: RelType.ChartColors,
          Target: chartColorsRelTarget(n)
        });
      }

      // Write chart rels if any
      if (rels.length > 0) {
        await this._renderToZip(zip, chartRelsPath(n), relsXform, rels);
      }
    }
  }

  async addChartExEntries(zip: IZipWriter, model: any): Promise<void> {
    const relsXform = new RelationshipsXform();

    // 1. Raw-bytes (round-trip) chartEx entries — byte-preserved from load
    for (const [n, rawBytes] of Object.entries(model.chartExEntries || {})) {
      // Write chartEx XML (raw bytes — passthrough for round-trip)
      zip.append(rawBytes as Uint8Array, { name: chartExPath(n) });

      // Write chartEx rels if present
      const rels = model.chartExRels?.[n];
      if (Array.isArray(rels) && rels.length > 0) {
        await this._renderToZip(zip, chartExRelsPath(n), relsXform, rels);
      }
    }

    // 2. Structured chartEx entries — built programmatically via addChartEx()
    const structured = model.chartExStructuredEntries ?? {};
    for (const [n, entry] of Object.entries(structured) as Array<[string, ChartExEntry]>) {
      // Skip if this number is already in raw bytes (raw takes precedence)
      if (model.chartExEntries?.[n]) {
        continue;
      }
      const xml = renderChartEx(entry.model);
      zip.append(xml, { name: chartExPath(n) });
      if (entry.rels && entry.rels.length > 0) {
        await this._renderToZip(zip, chartExRelsPath(n), relsXform, entry.rels);
      }
    }
  }

  async addTables(zip: IZipWriter, model: any): Promise<void> {
    const tableXform = new TableXform();

    for (const worksheet of model.worksheets) {
      for (const table of worksheet.tables) {
        tableXform.prepare(table, {});
        await this._renderToZip(zip, tablePath(table.target), tableXform, table);
      }
    }
  }

  /**
   * Write every external workbook reference into the archive. For each
   * {@link ExternalLinkModel} in `model.externalLinks` we emit two files:
   *
   *   xl/externalLinks/externalLink{index}.xml          — sheet names + cache
   *   xl/externalLinks/_rels/externalLink{index}.xml.rels — target path
   *
   * The target-path rel carries `TargetMode="External"` with a **bare
   * relative** `Target` whenever the user supplied one. This is the single
   * line that makes Office / WPS resolve the referenced workbook relative
   * to the current file's directory (not the `%USERPROFILE%\Documents`
   * fallback) — the root of the behaviour reported in exceljs#3039.
   */
  async addExternalLinks(zip: IZipWriter, model: any): Promise<void> {
    const externalLinks = (model.externalLinks ?? []) as ExternalLinkModel[];
    if (externalLinks.length === 0) {
      return;
    }

    const externalLinkXform = new ExternalLinkXform();
    const relsXform = new RelationshipsXform();

    for (const link of externalLinks) {
      await this._renderToZip(zip, externalLinkPath(link.index), externalLinkXform, link);

      // Always rId1 — the externalLink part only ever has a single rel.
      // `TargetMode="External"` is what flags Office to look the file up
      // at workbook-open time rather than embed it.
      await this._renderToZip(zip, externalLinkRelsPath(link.index), relsXform, [
        {
          Id: "rId1",
          Type: XLSX.RelType.ExternalLinkPath,
          Target: link.target,
          TargetMode: link.targetMode ?? "External"
        }
      ]);
    }
  }

  async addPivotTables(zip: IZipWriter, model: any): Promise<void> {
    if (!model.pivotTables.length) {
      return;
    }

    const pivotCacheRecordsXform = new PivotCacheRecordsXform();
    const pivotCacheDefinitionXform = new PivotCacheDefinitionXform();
    const pivotTableXform = new PivotTableXform();
    const relsXform = new RelationshipsXform();

    // R9-B6: Track which cacheIds have already been written to avoid duplicating
    // shared caches. Maps cacheId → tableNumber used for the cache file names.
    const writtenCaches = new Map<string, number>();

    for (const pivotTable of model.pivotTables as PivotTable[]) {
      const n = pivotTable.tableNumber;
      const isLoaded = pivotTable.isLoaded;
      const cacheId = pivotTable.cacheId;

      // R9-B6: Only write cache definition/records/rels once per unique cacheId.
      const cacheAlreadyWritten = writtenCaches.has(cacheId);
      if (!cacheAlreadyWritten) {
        writtenCaches.set(cacheId, n);

        if (isLoaded) {
          if (pivotTable.cacheDefinition) {
            await this._renderToZip(
              zip,
              pivotCacheDefinitionPath(n),
              pivotCacheDefinitionXform,
              pivotTable.cacheDefinition
            );
          }
          if (pivotTable.cacheRecords) {
            await this._renderToZip(
              zip,
              pivotCacheRecordsPath(n),
              pivotCacheRecordsXform,
              pivotTable.cacheRecords
            );
          }
        } else {
          await this._renderToZip(
            zip,
            pivotCacheRecordsPath(n),
            pivotCacheRecordsXform,
            pivotTable
          );
          await this._renderToZip(
            zip,
            pivotCacheDefinitionPath(n),
            pivotCacheDefinitionXform,
            pivotTable
          );
        }

        // R9-B4: Only write cache definition rels when cache records exist.
        const hasCacheRecords = isLoaded ? !!pivotTable.cacheRecords : true;
        if (hasCacheRecords) {
          const cacheRecordsRId =
            (isLoaded ? pivotTable.cacheDefinition?.rId : undefined) ?? "rId1";
          await this._renderToZip(zip, pivotCacheDefinitionRelsPath(n), relsXform, [
            {
              Id: cacheRecordsRId,
              Type: XLSX.RelType.PivotCacheRecords,
              Target: pivotCacheRecordsRelTarget(n)
            }
          ]);
        }
      }

      // Pivot table XML is always written (each pivot table has its own file).
      await this._renderToZip(zip, pivotTablePath(n), pivotTableXform, pivotTable);

      // Pivot table rels point to the cache definition file.
      const cacheTableNumber = writtenCaches.get(cacheId)!;
      await this._renderToZip(zip, pivotTableRelsPath(n), relsXform, [
        {
          Id: "rId1",
          Type: XLSX.RelType.PivotCacheDefinition,
          Target: pivotCacheDefinitionRelTargetFromPivotTable(cacheTableNumber)
        }
      ]);
    }
  }

  _finalize(zip: IZipWriter): Promise<this> {
    return new Promise((resolve, reject) => {
      zip.on("finish", () => {
        resolve(this);
      });
      zip.on("error", reject);
      zip.finalize();
    });
  }

  prepareModel(model: any, options: any): void {
    model.creator = model.creator ?? "ExcelTS";
    model.lastModifiedBy = model.lastModifiedBy ?? "ExcelTS";
    model.created = model.created ?? new Date();
    model.modified = model.modified ?? new Date();

    model.useSharedStrings =
      options.useSharedStrings !== undefined ? options.useSharedStrings : true;
    model.useStyles = options.useStyles !== undefined ? options.useStyles : true;

    model.sharedStrings = new SharedStringsXform();

    // Preserve default font from parsed styles if available
    const oldDefaultFont = model.defaultFont;
    model.styles = model.useStyles ? new StylesXform(true) : new (StylesXform as any).Mock();
    if (oldDefaultFont && model.styles.setDefaultFont) {
      model.styles.setDefaultFont(oldDefaultFont);
    }

    const workbookXform = new WorkbookXform();
    const worksheetXform = new WorkSheetXform();

    workbookXform.prepare(model);

    // Normalise external-workbook references before any formula rendering.
    // Two jobs:
    //   1. Scan every formula cell and make sure each referenced workbook
    //      has a matching ExternalLinkModel in `model.externalLinks`, with
    //      a stable 1-based index.
    //   2. Rewrite formula strings from `[filename.xlsx]Sheet!A1` to
    //      `[N]Sheet!A1`, which is the canonical on-disk form Excel
    //      expects inside `<f>` elements.
    //
    // Done once up-front (not per cell) so the index assignment is
    // deterministic and every cell sees the final externalLinks list.
    this._normaliseExternalLinks(model);

    const worksheetOptions: any = {
      sharedStrings: model.sharedStrings,
      styles: model.styles,
      date1904: model.properties?.date1904,
      drawingsCount: 0,
      media: model.media
    };
    worksheetOptions.drawings = model.drawings = [];
    worksheetOptions.commentRefs = model.commentRefs = [];
    worksheetOptions.formControlRefs = model.formControlRefs = [];
    model.hasHeaderWatermark = false;
    let tableCount = 0;
    model.tables = [];
    const tableNameMap = new Map<string, string>(); // name (lowercase) → worksheet name
    model.worksheets.forEach((worksheet: any, index: number) => {
      // Assign fileIndex early so that worksheet-xform.prepare() can use it
      // for comment/VML relationship targets and content type names.
      // This ensures consistency with addWorksheets() which writes ZIP entries
      // using the same fileIndex.
      worksheet.fileIndex = index + 1;

      worksheet.tables.forEach((table: any) => {
        // OOXML requires table names to be unique across the entire workbook
        // (case-insensitive). Detect duplicates early to produce a clear error
        // instead of generating a corrupt file that Excel must repair.
        const nameKey = table.name.toLowerCase();
        const existingSheet = tableNameMap.get(nameKey);
        if (existingSheet !== undefined) {
          throw new TableError(
            `Duplicate table name "${table.name}": already used in worksheet "${existingSheet}". ` +
              `Table names must be unique across the entire workbook (case-insensitive).`
          );
        }
        tableNameMap.set(nameKey, worksheet.name);

        tableCount++;
        table.target = `table${tableCount}.xml`;
        table.id = tableCount;
        model.tables.push(table);
      });

      worksheetXform.prepare(worksheet, worksheetOptions);
    });

    // ContentTypesXform expects this flag
    model.hasCheckboxes = model.styles.hasCheckboxes;

    // Scan all worksheets for dynamic array formulas.
    // cm=1 is assigned later by cell-xform.prepare() during worksheet rendering.
    let dynamicArrayCount = 0;
    model.worksheets.forEach((worksheet: any) => {
      (worksheet.rows ?? []).forEach((row: any) => {
        (row.cells ?? []).forEach((cell: any) => {
          if (cell.isDynamicArray) {
            dynamicArrayCount++;
          }
        });
      });
    });
    model.hasDynamicArrayFormulas = dynamicArrayCount > 0;
    model.dynamicArrayCount = dynamicArrayCount;

    // Propagate header watermark flag from worksheet prepare options
    if (worksheetOptions.hasHeaderWatermark) {
      model.hasHeaderWatermark = true;
    }
  }
}

export { XLSX };
