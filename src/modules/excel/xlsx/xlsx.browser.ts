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
import {
  commentsPath,
  commentsRelTargetFromWorksheetName,
  ctrlPropPath,
  drawingPath,
  drawingRelsPath,
  OOXML_REL_TARGETS,
  pivotTableRelTargetFromWorksheetName,
  pivotCacheDefinitionRelTargetFromWorkbook,
  getCommentsIndexFromPath,
  getDrawingNameFromPath,
  getDrawingNameFromRelsPath,
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
  tableRelTargetFromWorksheetName,
  themePath,
  getThemeNameFromPath,
  getVmlDrawingNameFromPath,
  getVmlDrawingHFNameFromPath,
  getWorksheetNoFromWorksheetPath,
  getWorksheetNoFromWorksheetRelsPath,
  isBinaryEntryPath,
  normalizeZipPath,
  OOXML_PATHS,
  vmlDrawingRelTargetFromWorksheetName,
  vmlDrawingPath,
  vmlDrawingHFPath,
  vmlDrawingHFRelsPath,
  worksheetPath,
  worksheetRelsPath,
  worksheetRelTarget
} from "@excel/utils/ooxml-paths";
import { PassthroughManager } from "@excel/utils/passthrough-manager";
import { StreamBuf } from "@excel/utils/stream-buf";
import type { Workbook } from "@excel/workbook";
import { RelType } from "@excel/xlsx/rel-type";
import { WorkbookXform } from "@excel/xlsx/xform/book/workbook-xform";
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

export interface IStreamBuf extends EmitterLike {
  write(data: any): boolean | void | Promise<boolean>;
  end(): void;
  read(): any;
  toBuffer?(): any;
  pipe?(dest: any): any;
}

export interface IZipWriter extends EmitterLike {
  append(data: any, options: { name: string; base64?: boolean }): void;
  /** Create a streaming entry: write chunks incrementally, then call end(). */
  createEntry(name: string): { write(chunk: string): void; end(): void };
  pipe(stream: any): void;
  finalize(): void;
  /** Wait for downstream backpressure to clear. Resolves immediately if no backpressure. */
  waitForDrain(): Promise<void>;
}

class StreamingZipWriterAdapter implements IZipWriter {
  private static textEncoder = new TextEncoder();

  private readonly zip: StreamingZip;
  private readonly events: Map<string, Set<StreamListener>> = new Map();
  private pipedStream: Pick<IStreamBuf, "write" | "end"> | null = null;
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

  pipe(stream: any): void {
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

export interface XlsxReadOptions {
  base64?: boolean;
  [key: string]: unknown;
}

export interface ZipWriterOptions {
  level?: number;
  /** ZIP entry modification time (optional). If omitted, defaults to current time. */
  modTime?: Date;
  /** Timestamp writing strategy for ZIP entry metadata (optional). */
  timestamps?: ZipTimestampMode;
}

export interface XlsxWriteOptions {
  zip?: ZipWriterOptions;
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
  type: "Directory" | "File";
  stream: IParseStream;
  drain: () => Promise<void>;
}

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
    await this.addSharedStrings(zip, model);
    await this.addDrawings(zip, model);
    await this.addTables(zip, model);
    await this.addPivotTables(zip, model);
    this.addPassthrough(zip, model);
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
  async read(stream: IParseStream, options?: XlsxReadOptions): Promise<any> {
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
  async write(stream: any, options?: XlsxWriteOptions): Promise<XLSX> {
    options = options || {};

    options.zip = options.zip || {};
    options.zip.modTime ??= this.workbook.modified ?? this.workbook.created;

    const zip = this.createZipWriter(options.zip);
    zip.pipe(stream);
    await this.writeToZip(zip, options);
    return this._finalize(zip) as Promise<XLSX>;
  }

  /**
   * Load workbook from buffer/ArrayBuffer/Uint8Array
   */
  async load(data: any, options?: XlsxReadOptions): Promise<any> {
    let buffer: Uint8Array;

    // Validate input
    const isBuffer = typeof Buffer !== "undefined" ? Buffer.isBuffer(data) : false;
    if (
      !data ||
      (typeof data === "object" &&
        !isBuffer &&
        !(data instanceof Uint8Array) &&
        !(data instanceof ArrayBuffer))
    ) {
      throw new ExcelFileError(
        "<input>",
        "read",
        "Can't read the data of 'the loaded zip file'. Is it in a supported JavaScript type (String, Blob, ArrayBuffer, etc) ?"
      );
    }

    // Handle base64 input
    if (options && options.base64) {
      buffer = base64ToUint8Array(data.toString());
    } else if (data instanceof ArrayBuffer) {
      buffer = new Uint8Array(data);
    } else if (data instanceof Uint8Array) {
      buffer = data;
    } else {
      // Node.js Buffer or other array-like
      buffer = new Uint8Array(data);
    }

    return this.loadBuffer(buffer, options);
  }

  /**
   * Internal: Load from Uint8Array buffer
   */
  protected async loadBuffer(buffer: Uint8Array, options?: XlsxReadOptions): Promise<any> {
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
      // Raw drawing XML data for passthrough (when drawing contains chart references)
      rawDrawings: {} as Record<string, Uint8Array>,
      comments: {},
      tables: {},
      vmlDrawings: {},
      pivotTables: {},
      pivotTableRels: {},
      pivotCacheDefinitions: {},
      pivotCacheRecords: {},
      // Passthrough storage for unknown/unsupported files (charts, etc.)
      passthrough: {} as Record<string, Uint8Array>
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
   * Check if a drawing has chart references in its relationships
   */
  private drawingHasChartReference(drawing: any): boolean {
    return (
      drawing.rels && drawing.rels.some((rel: any) => rel.Target && rel.Target.includes("/charts/"))
    );
  }

  /**
   * Check if a drawing rels list references charts.
   * Used to decide whether we need to keep raw drawing XML for passthrough.
   */
  private drawingRelsHasChartReference(drawingRels: any[] | undefined): boolean {
    return (
      Array.isArray(drawingRels) &&
      drawingRels.some(rel => typeof rel?.Target === "string" && rel.Target.includes("/charts/"))
    );
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

    // Trim raw drawings for non-chart drawings to avoid bloating the serialized workbook model.
    if (model.rawDrawings && model.drawingRels) {
      for (const name of Object.keys(model.rawDrawings)) {
        const drawingRel = model.drawingRels[name];
        if (drawingRel && !this.drawingRelsHasChartReference(drawingRel)) {
          delete model.rawDrawings[name];
        }
      }
    }

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
      // Key format (e.g., "../pivotTables/pivotTable1.xml") matches worksheet .rels Target values,
      // allowing worksheet reconciliation to look up pivot tables by their relationship target path.
      pivotTablesIndexed[pivotTableRelTargetFromWorksheetName(pivotName)] = completePivotTable;
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

  async _processCommentEntry(stream: IParseStream, model: any, name: string): Promise<void> {
    const xform = new CommentsXform();
    const comments = await xform.parseStream(stream);
    model.comments[commentsRelTargetFromWorksheetName(name)] = comments;
  }

  async _processTableEntry(stream: IParseStream, model: any, name: string): Promise<void> {
    const xform = new TableXform();
    const table = await xform.parseStream(stream);
    model.tables[tableRelTargetFromWorksheetName(name)] = table;
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

    // Store raw data; reconcile() may later drop it if charts are not referenced.
    model.rawDrawings[name] = data;
  }

  async _processDrawingRelsEntry(entry: any, model: any, name: string): Promise<void> {
    const xform = new RelationshipsXform();
    const relationships = await xform.parseStream(entry);
    model.drawingRels[name] = relationships;
  }

  async _processVmlDrawingEntry(entry: any, model: any, name: string): Promise<void> {
    const xform = new VmlDrawingXform();
    const vmlDrawing = await xform.parseStream(entry);
    model.vmlDrawings[vmlDrawingRelTargetFromWorksheetName(name)] = vmlDrawing;
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

  // ===========================================================================
  // loadFromFiles - shared logic for loading from pre-extracted ZIP data
  // ===========================================================================

  async loadFromFiles(zipData: Record<string, Uint8Array>, options?: any): Promise<any> {
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
          // Pass raw entry data for drawings to enable passthrough
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

    const mediaFilename = getMediaFilenameFromPath(entryName);
    if (mediaFilename) {
      await this._processMediaEntry(stream, model, mediaFilename);
      return true;
    }

    const drawingName = getDrawingNameFromPath(entryName);
    if (drawingName) {
      await this._processDrawingEntry(stream, model, drawingName, rawData);
      // rawData is now stored inside _processDrawingEntry
      return true;
    }

    const drawingRelsName = getDrawingNameFromRelsPath(entryName);
    if (drawingRelsName) {
      await this._processDrawingRelsEntry(stream, model, drawingRelsName);
      return true;
    }

    const vmlDrawingName = getVmlDrawingNameFromPath(entryName);
    if (vmlDrawingName) {
      await this._processVmlDrawingEntry(stream, model, vmlDrawingName);
      return true;
    }

    // VML header/footer drawings (watermark in header mode).
    // Parse to extract header image info for round-trip preservation.
    const vmlHFName = getVmlDrawingHFNameFromPath(entryName);
    if (vmlHFName) {
      await this._processVmlDrawingHFEntry(stream, model, vmlHFName);
      return true;
    }

    const commentsIndex = getCommentsIndexFromPath(entryName);
    if (commentsIndex) {
      await this._processCommentEntry(stream, model, `comments${commentsIndex}`);
      return true;
    }

    const tableName = getTableNameFromPath(entryName);
    if (tableName) {
      await this._processTableEntry(stream, model, tableName);
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

    // Store passthrough files (charts, etc.) for preservation
    if (PassthroughManager.isPassthroughPath(entryName)) {
      // If raw data is available (loadFromFiles path), use it directly
      if (rawData) {
        model.passthrough[entryName] = rawData;
      } else {
        await this._processPassthroughEntry(stream, model, entryName);
      }
      return true;
    }

    return false;
  }

  /**
   * Store a passthrough file for preservation during read/write cycles.
   * These files are not parsed but stored as raw bytes to be written back unchanged.
   */
  async _processPassthroughEntry(
    stream: IParseStream,
    model: any,
    entryName: string
  ): Promise<void> {
    const data = await this.collectStreamData(stream);
    model.passthrough[entryName] = data;
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

  async addDrawings(zip: IZipWriter, model: any): Promise<void> {
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();
    const rawDrawings = model.rawDrawings || {};

    for (const worksheet of model.worksheets) {
      const { drawing } = worksheet;
      if (drawing) {
        // Check if drawing rels contain chart references using helper
        const hasChartReference = this.drawingHasChartReference(drawing);

        if (hasChartReference && rawDrawings[drawing.name]) {
          // Use raw data for drawings with chart references (passthrough)
          zip.append(rawDrawings[drawing.name], { name: drawingPath(drawing.name) });
        } else {
          // Use regenerated XML for normal drawings (images, shapes)
          const filteredAnchors = filterDrawingAnchors(drawing.anchors ?? []);
          const drawingForWrite = drawing.anchors
            ? { ...drawing, anchors: filteredAnchors }
            : drawing;
          drawingXform.prepare(drawingForWrite);
          await this._renderToZip(zip, drawingPath(drawing.name), drawingXform, drawingForWrite);
        }

        await this._renderToZip(zip, drawingRelsPath(drawing.name), relsXform, drawing.rels);
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
   * Write passthrough files (charts, etc.) that were preserved during read.
   * These files are written back unchanged to preserve unsupported features.
   */
  addPassthrough(zip: IZipWriter, model: any): void {
    const passthroughManager = new PassthroughManager();
    passthroughManager.fromRecord(model.passthrough || {});
    passthroughManager.writeToZip(zip);
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

    // Build passthroughContentTypes for ContentTypesXform using PassthroughManager
    const passthrough = model.passthrough || {};
    const passthroughManager = new PassthroughManager();
    passthroughManager.fromRecord(passthrough);
    model.passthroughContentTypes = passthroughManager.getContentTypes();
  }
}

export { XLSX };
