/**
 * WorkbookWriter - Browser Streaming Excel Writer
 *
 * This module contains the full cross-platform implementation for the streaming
 * workbook writer and a browser-compatible `WorkbookWriter` class.
 *
 * Node.js uses `workbook-writer.ts`, which extends the same base implementation
 * with filesystem-specific features (filename output + image loading).
 */

import { Zip, ZipDeflate } from "@archive/zip/stream";
import type { BuiltInCellStyle } from "@excel/core/builtin-cell-styles";
import { BUILTIN_CELL_STYLES } from "@excel/core/builtin-cell-styles";
import type { DefinedNamesData } from "@excel/core/defined-names";
import { createDefinedNames, definedNamesModel } from "@excel/core/defined-names";
import type { NamedStyleEntry } from "@excel/core/workbook-core";
import {
  checkImageData,
  refuseLongSheetName,
  validateCellStyleName,
  validateSheetName
} from "@excel/core/workbook-core";
import type { WorkbookFormat } from "@excel/core/workbook-format";
import { getWorkbookModel } from "@excel/core/workbook-model";
import type { ProtectWorkbookOptions } from "@excel/core/workbook.browser";
import {
  ExcelUnsupportedFeatureError,
  ExcelStreamStateError,
  ExcelImageError
} from "@excel/errors";
import { WorksheetWriter } from "@excel/stream/worksheet-writer";
import type { WorkbookWriterLike } from "@excel/stream/worksheet-writer";
import {
  createStreamedXlsbTables,
  streamedFormulaNames,
  type StreamedXlsbTables
} from "@excel/stream/xlsb-writer";
import type {
  Font,
  ImageData,
  WorkbookView,
  WorkbookProtection,
  AddWorksheetOptions,
  WorksheetProperties,
  WorksheetState,
  PageSetup,
  WorksheetView,
  AutoFilter,
  HeaderFooter,
  NamedStyle
} from "@excel/types";
import {
  argumentChecks,
  checkProtectCall,
  WORKBOOK_PROTECTION_FLAGS
} from "@excel/utils/arguments";
import { unwritableFormulaReason } from "@excel/utils/array-evaluation";
import { filterDrawingAnchors, isExternalImage } from "@excel/utils/drawing-utils";
import type { DrawingAnchor, DrawingRel } from "@excel/utils/drawing-utils";
import { formulaNameSet, withFormulaNames } from "@excel/utils/future-functions";
import {
  drawingPath,
  drawingRelsPath,
  mediaPath,
  OOXML_PATHS,
  OOXML_REL_TARGETS,
  worksheetRelTarget
} from "@excel/utils/ooxml-paths";
import type { PackageSink, PartWriter } from "@excel/utils/package-sink";
import { SharedStrings } from "@excel/utils/shared-strings";
import { splitPasswordOptions } from "@excel/utils/sheet-protection";
import { StreamBuf } from "@excel/utils/stream-buf";
import { checkStyle } from "@excel/utils/style-checks";
import { writableDefinedNames } from "@excel/utils/unwritable-formulas";
import { buildWorkbookProtection } from "@excel/utils/workbook-protection";
import { printNamesFromModel } from "@excel/xlsb/write/model-adapter";
import { withSheetExternSheets } from "@excel/xlsb/write/workbook";
import { RelType } from "@excel/xlsx/rel-type";
import { WorkbookXform } from "@excel/xlsx/xform/book/workbook-xform";
import { AppXform } from "@excel/xlsx/xform/core/app-xform";
import { ContentTypesXform } from "@excel/xlsx/xform/core/content-types-xform";
import { CoreXform } from "@excel/xlsx/xform/core/core-xform";
import { FeaturePropertyBagXform } from "@excel/xlsx/xform/core/feature-property-bag-xform";
import { MetadataXform } from "@excel/xlsx/xform/core/metadata-xform";
import { RelationshipsXform } from "@excel/xlsx/xform/core/relationships-xform";
import { SharedStringsXform } from "@excel/xlsx/xform/strings/shared-strings-xform";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { theme1Xml } from "@excel/xlsx/xml/theme1";
import type { Writable } from "@stream";
import { Streams } from "@stream";
import { stringToUint8Array } from "@utils/binary";
import { excelNameKey, sheetNameKey } from "@utils/name-key";
import { base64ToUint8Array } from "@utils/utils";

const EMPTY_U8 = new Uint8Array(0);
const TEXT_DECODER = new TextDecoder();

/**
 * Drain a resolver list, calling each. Mutates the array to empty.
 *
 * Used by the backpressure machinery: when a sink drains or errors, every
 * parked `_waitForUserSinkDrain()` / pending-async waiter must be woken
 * exactly once, and the array reset so the next backpressure cycle starts
 * clean. Hoisted to a free function so it can be re-used across the three
 * wake sites without per-site duplication of the splice/loop pattern.
 */
function callAllResolvers(resolvers: Array<() => void>): void {
  if (resolvers.length === 0) {
    return;
  }
  // Snapshot then clear, so a resolver that itself triggers a fresh wait
  // (re-pushing into the same array) doesn't get confused with the current
  // batch.
  const snapshot = resolvers.splice(0);
  for (const r of snapshot) {
    r();
  }
}

// ============================================================================
// Types
// ============================================================================

/**
 * An image registered in the streaming writer.
 *
 * Extends the public {@link ImageData} shape with the unique stored name
 * (`name`) assigned by `addImage`, and pins `type` to `"image"`.
 */
export interface Medium extends Omit<ImageData, "extension"> {
  type: "image";
  name: string;
  /**
   * Widened from `ImageData.extension` so an SVG companion medium can carry
   * the `"svg"` extension (the public `addImage` input stays raster-only).
   */
  extension: string;
  /** Media index of an SVG companion (raster blip + svgBlip extension). */
  svgMediaId?: number;
}

interface CommentRef {
  commentName: string;
  vmlDrawing: string;
}

export interface ZlibOptions {
  flush?: number;
  finishFlush?: number;
  chunkSize?: number;
  windowBits?: number;
  level?: number;
  memLevel?: number;
  strategy?: number;
  dictionary?: Uint8Array | ArrayBuffer;
}

export interface WorkbookZipOptions {
  comment?: string;
  forceLocalTime?: boolean;
  forceZip64?: boolean;
  store?: boolean;
  zlib?: Partial<ZlibOptions>;
  compressionOptions?: { level?: number };
}

export interface WorkbookWriterOptions {
  created?: Date;
  modified?: Date;
  creator?: string;
  lastModifiedBy?: string;
  lastPrinted?: Date;
  useSharedStrings?: boolean;
  useStyles?: boolean;
  /**
   * Count dates from 1904 instead of 1900.
   *
   * **Declared here because it could previously only be set by a cast.** A random-access workbook carries it on
   * `properties.date1904`; a `WorkbookWriter` has no `properties` at all, so a caller who needed the 1904 system had to
   * assign an undeclared field — which this library's own test for the setting did. The value reaches both the workbook
   * part and every cell serial through `date1904Flag()`.
   *
   * Read per row rather than captured, so setting it before the first commit is what matters; see `date1904Flag`.
   */
  date1904?: boolean;
  zip?: Partial<WorkbookZipOptions>;
  /**
   * Destination sink. Backpressure is respected, so the sink must already be
   * consumed (or be a terminal sink such as `fs.createWriteStream`, an HTTP
   * response, or an upload body). Handing over an unconsumed intermediate
   * stream — e.g. a bare `PassThrough` whose reader is attached only after
   * `commit()` resolves — deadlocks once its buffers fill.
   *
   * `commit()` ends it. With neither `stream` nor `filename` the writer keeps the package itself and `toBuffer()`
   * returns it.
   */
  stream?: Writable | WritableStream<Uint8Array>;
  filename?: string; // Node.js only
  trueStreaming?: boolean;
  /**
   * Which container to write. Defaults to `"xlsx"`.
   *
   * `"xlsb"` produces the binary form from the same calls, with the same memory bound: rows are encoded and handed
   * to the ZIP as they are committed. See `stream/xlsb-writer` for what is and is not bounded, and for the single
   * record (`BrtWsDim`) a forward pass cannot write.
   */
  format?: WorkbookFormat;
}

interface OutputStreamLike {
  emit(eventName: string | symbol, ...args: unknown[]): boolean;
  write(chunk: Uint8Array | string): boolean | Promise<boolean>;
  end(): void;
  // Node's EventEmitter-style callbacks receive heterogeneous args whose types
  // depend on the event. We keep `any[]` here because `unknown[]` would be too
  // restrictive for callers that declare typed listeners like `(err: Error) =>`.
  once(eventName: string | symbol, listener: (...args: any[]) => void): this;
  removeListener(eventName: string | symbol, listener: (...args: any[]) => void): this;
  // Optional: not all sinks expose `.on` (e.g. internal `StreamBuf` predates
  // the EventEmitter contract). Backpressure listeners are skipped when
  // missing — the runtime guard `typeof stream.on === "function"` is what
  // actually drives the behaviour.
  on?(eventName: string | symbol, listener: (...args: any[]) => void): this;
}

// ============================================================================
// WorksheetWriter interface (to avoid circular dependency)
// ============================================================================

export interface WorksheetWriterLike {
  id: number;
  name: string;
  rId?: string;
  committed?: boolean;
  /** Sequential ZIP entry index, assigned to satisfy the content-types contract. */
  fileIndex?: number;
  stream: InstanceType<typeof StreamBuf>;
  commit(): void;
  /**
   * What the binary encoder could not express in this sheet's *rows*, after `commit()`.
   *
   * Part of the contract because the package writer cannot recover it: by the time it runs the sheet is in the ZIP and
   * its rows are gone. Without this the row-level report was computed and discarded, so a cross-sheet formula that
   * became a blank cell left `WorkbookWriter.xlsbUnsupported` empty.
   *
   * Optional so an XLSX-only worksheet writer need not carry it.
   */
  xlsbUnsupported?: readonly string[];
  /** What this sheet's rows could not express, in either format, as `Sheet!A1: reason`, after `commit()`. */
  unsupported?: readonly string[];
  /** Drawing model — populated after commit if images were added */
  drawing?: { rId: string; name: string; anchors: DrawingAnchor[]; rels: DrawingRel[] };
  /** The sheet's model as the XLSB package reads it; optional so an XLSX-only worksheet writer need not carry it. */
  xlsbModel?(): { readonly media?: unknown };
}

export interface WorksheetWriterConstructor<T extends WorksheetWriterLike> {
  new (options: {
    id: number;
    name: string;
    workbook: WorkbookWriterLike;
    useSharedStrings: boolean;
    properties?: Partial<WorksheetProperties>;
    state?: WorksheetState;
    pageSetup?: Partial<PageSetup>;
    views?: Partial<WorksheetView>[];
    autoFilter?: AutoFilter;
    headerFooter?: Partial<HeaderFooter>;
  }): T;
}

// ============================================================================
// Base Class
// ============================================================================

export abstract class WorkbookWriterBase<TWorksheetWriter extends WorksheetWriterLike> {
  created: Date;
  modified: Date;
  creator: string;
  lastModifiedBy: string;
  lastPrinted?: Date;
  useSharedStrings: boolean;
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  sharedStrings: SharedStrings;
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  styles: StylesXform;
  private _definedNames: DefinedNamesData;
  private _worksheets: TWorksheetWriter[];
  views: WorkbookView[];
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  zipOptions?: Partial<WorkbookZipOptions>;
  /** Which container this writer produces. Read by `WorksheetWriter` to decide what a row becomes. */
  readonly format: WorkbookFormat;
  /**
   * The interning tables and written-path set a streamed XLSB shares across its sheets.
   *
   * Created eagerly rather than on first use so that `WorksheetWriter` can rely on it existing: a `BrtCellIsst`
   * index is workbook-wide, and a table created per sheet would make every sheet after the first name the wrong
   * strings. Undefined for XLSX, where nothing needs it.
   *
   * @deprecated Internal; not part of the public API. Removed in 1.0.
   */
  readonly xlsbTables?: StreamedXlsbTables;
  /**
   * What this writer could not express, once `commit()` has run, in either format. Empty unless something was dropped.
   *
   * A row's loss is `Sheet!A1: reason`, the form `Workbook.toBuffer` reports; a package-level one is as the package
   * writer words it. An XLSX stream has one row loss of its own: an unnamed this-row reference (`[@x]`), which a
   * streamed sheet — holding no table — can only write as the cell's cached value.
   *
   * Reported, not thrown: rows are already in the caller's stream by the time anything is known, so a refusal could
   * take nothing back.
   */
  unsupported: readonly string[] = [];
  private _xlsbUnsupported: readonly string[] = [];
  /** The spelling of each named cell style added, keyed as Excel compares the names (without regard to case). */
  private readonly _cellStyleNames = new Map<string, string>();

  /**
   * @deprecated Use {@link unsupported}. Removed in 1.0.
   *
   * What it always returned: the XLSB writer's losses only — empty for an XLSX stream — with a row's loss as `A1: …`.
   */
  get xlsbUnsupported(): readonly string[] {
    return this._xlsbUnsupported;
  }

  /**
   * The workbook-wide names a formula's token stream resolves against, as they stand *now*.
   *
   * A BIFF12 formula stores `Sheet2!A1` as an `ixti` index rather than a name, so encoding one needs the workbook's
   * sheet list. The streamed sheet writer had no context at all and every cross-sheet formula became a blank cell —
   * reported by the row encoder and then discarded, so the report read empty.
   *
   * **Built on demand rather than held**, because it grows: `addWorksheet` is called between sheet commits, so a sheet
   * committed later resolves against more sheets than an earlier one. That asymmetry is inherent to writing forward and
   * is exactly why it must not be captured once.
   *
   * `externSheets` is the workbook's one table — the object `writeXlsbPackage` writes at the end — so the encoder resolves
   * an `ixti` against the table the file will carry, and the workbook part knows one was resolved. It is mutable on
   * purpose: a reference spanning a range of sheets has no single-sheet entry and the encoder appends one.
   *
   * A forward reference — sheet 1 naming sheet 3 — still cannot resolve, and is now *reported* rather than silent.
   *
   * @deprecated The sheet writer's view of the workbook while it encodes rows; nothing outside the writer needs it.
   * Removed in 1.0.
   */
  xlsbFormulaContext(): {
    readonly sheetNames: readonly string[];
    readonly definedNames: readonly string[];
    readonly definedNameScopes: readonly (number | undefined)[];
    readonly externSheets: { first: number; last: number }[];
  } {
    // **Compacted, as the package's sheet list is.** `_worksheets` is indexed by sheet id, which starts at 1, so slot 0
    // is a hole — and `map` keeps holes, which made every `ixti` one too high: `S!A1` written from sheet `T` read back
    // as `T!A1`.
    const worksheets = this._worksheets.filter(Boolean) as { name?: string; pageSetup?: unknown }[];
    const sheetNames = worksheets.map((sheet, index) => sheet.name ?? `Sheet${index + 1}`);
    // The print names follow the workbook's own, as `writeXlsbPackage` lists them: the streamed future-function stubs
    // come after both, so leaving them out would put every stub index off by their count.
    const names = [
      ...definedNamesModel(this._definedNames),
      ...printNamesFromModel(
        worksheets.map(sheet => ({ name: sheet.name, pageSetup: sheet.pageSetup })) as never
      )
    ];
    return {
      sheetNames,
      definedNames: names.map(defined => defined.name),
      definedNameScopes: names.map(defined => defined.localSheetId),
      // The workbook's one table, shared with every sheet and with `writeXlsbPackage`, so the workbook part carries the
      // table the streamed formulas were encoded against — see `StreamedXlsbTables.externSheets`.
      externSheets: withSheetExternSheets(this.xlsbTables?.externSheets ?? [], sheetNames.length)
    };
  }
  /**
   * Whether this workbook counts days from 1904, as the date system stands *now*.
   *
   * **A serial is meaningless without it, and the streamed row encoders were not asking.** The workbook part is
   * produced at `commit()` from the model, so `BrtWbProp` recorded the real setting while every cell serial had been
   * computed against the other epoch — a package internally inconsistent by 1,462 days. Measured on a workbook with
   * `date1904: true`: `2020-01-15` written through `Workbook.toBuffer` read back as `2020-01-15`, and through the
   * streaming writer as `2024-01-16`.
   *
   * **Named for the setting rather than for a format.** It was `xlsbDate1904` while only the BIFF12 branch consulted
   * it, and that name was load-bearing in the wrong direction: the XLSX branch of `_writeRow` had the identical
   * defect and the identical fix available, and a method whose name said the option did not apply to it is part of
   * why it sat there. `date1904` is not an XLSB concept.
   *
   * Read per row rather than captured, for the reason `xlsbFormulaContext` is: a caller may set the property between
   * `addWorksheet` and the first row. A caller who changes it *after* rows are committed gets the ordinary
   * forward-pass consequence — the rows already encoded keep the epoch they were encoded with — which is the same
   * constraint the columns, panes and views are under.
   *
   * @deprecated The sheet writer's view of the workbook while it encodes rows; nothing outside the writer needs it.
   * Removed in 1.0.
   */
  date1904Flag(): boolean {
    // `date1904` is a declared writer option now. `properties.date1904` is still honoured because a caller may have
    // assigned it — that was the only way to set this before the option existed, and this library's own test did it.
    return this.properties?.date1904 === true;
  }
  /**
   * @internal The workbook's defined names, by `excelNameKey` — see `WorkbookWriterLike.formulaNames`. Read per row
   * like `date1904Flag`: a name defined after rows are committed does not reach them, the forward-pass rule.
   */
  formulaNames(): ReadonlySet<string> {
    return formulaNameSet(Object.values(this._definedNames.nameForKey).map(name => ({ name })));
  }

  /**
   * Workbook properties, which a streaming writer did not have at all.
   *
   * **One field, because `date1904` was being read from two places that could not agree.** The option sets the epoch the
   * *rows* are encoded against, while `writeXlsbPackage` writes `BrtWbProp` from `model.properties?.date1904` — and the
   * model comes from `getWorkbookModel`, which reads `wb.properties`. With no such property the two disagreed by 1,462
   * days in the opposite direction from the original defect: `date1904: true` produced cells encoded for 1904 inside a
   * package declaring 1900, and `2020-01-15` read back as `2016-01-14`.
   *
   * A plain field rather than a getter, so a caller may still assign it — which was the only way to set the date system
   * before the option existed.
   */
  properties: { date1904?: boolean } = {};

  /** Deflate completions for entries opened through `packageSink()`, awaited before the ZIP is finalised. */
  private readonly _sinkCompletions: Promise<void>[] = [];
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  compressionLevel: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  media: Medium[];
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  commentRefs: CommentRef[];
  /**
   * Number of cells with dynamic array formulas, accumulated during worksheet commit.
   * @deprecated Internal; not part of the public API. Removed in 1.0.
   */
  dynamicArrayCount: number;
  /** Workbook-level structure protection */
  protection?: {
    lockStructure?: boolean;
    lockWindows?: boolean;
    lockRevision?: boolean;
    algorithmName?: string;
    hashValue?: string;
    saltValue?: string;
    spinCount?: number;
  };
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  zip: Zip;
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  stream: OutputStreamLike;
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  promise: Promise<void[] | void>;
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  protected _trueStreaming: boolean;
  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  protected WorksheetWriterClass: WorksheetWriterConstructor<TWorksheetWriter>;

  // ---------------------------------------------------------------------------
  // Backpressure tracking for the user-supplied output sink.
  //
  // Set by `_trackBackpressure(ok)` whenever `this.stream.write(data)` returns
  // false (or a Promise that resolves to false). Cleared when the sink emits
  // `'drain'`. Awaited by `_waitForUserSinkDrain()` at async boundaries
  // (between worksheets, before `addWorkbook`, etc) so a slow sink throttles
  // the producer instead of letting bytes accumulate unboundedly inside the
  // sink's internal buffer or in the zip pipeline.
  //
  // Important caveat: this **cannot** block a single tight synchronous
  // `for (...) row.commit()` loop inside one worksheet — JavaScript has no
  // sync wait, and `row.commit()` is sync void. During such a loop, every
  // produced compressed chunk is pushed straight into the sink's internal
  // buffer (Node `Writable` accepts writes after returning false; it just
  // hints "drain"). For very-large single-worksheet workloads with a slow
  // sink, the practical bound on how much can pile up is roughly the total
  // compressed size of one worksheet — only the `wb.commit()` boundary
  // (and any `worksheet.commit()` between sheets) gives the event loop a
  // chance to park here on `_waitForUserSinkDrain()`.
  //
  // Multi-sheet workloads benefit fully because each `worksheet.commit()`
  // hands control back to `_commitWorksheets()` which awaits drain before
  // the next sheet starts.
  private _needsDrain = false;
  private _drainResolvers: Array<() => void> = [];
  private _drainListenerAttached = false;
  private _drainGeneration = 0;
  private _pendingSinkWrites = new Set<Promise<void>>();
  private _pendingZipPushes = new Set<Promise<void>>();
  private _worksheetZipCompletions = new WeakMap<InstanceType<typeof StreamBuf>, Promise<void>>();
  // Captured if the user sink fires 'error' before `_finalize()` attaches its
  // own listener. Replayed by `_finalize()` so the original error is what
  // rejects `commit()`, not a generic timeout.
  private _sinkError: Error | null = null;

  constructor(
    options: WorkbookWriterOptions,
    WorksheetWriterClass: WorksheetWriterConstructor<TWorksheetWriter>
  ) {
    // `null` reached `options.created` and threw a `TypeError`; a number was read as no options at all.
    argumentChecks("WorkbookWriter").record("options", options);
    this.WorksheetWriterClass = WorksheetWriterClass;
    this.created = options.created || new Date();
    this.modified = options.modified || this.created;
    this.creator = options.creator ?? "Documonster";
    this.lastModifiedBy = options.lastModifiedBy ?? "Documonster";
    this.lastPrinted = options.lastPrinted;

    this.useSharedStrings = options.useSharedStrings ?? false;
    // Into `properties`, which is what both readers of this setting consult — the row encoder through `date1904Flag()`
    // and the workbook part through `getWorkbookModel().properties`.
    this.properties = { ...(options.date1904 === undefined ? {} : { date1904: options.date1904 }) };
    this.sharedStrings = new SharedStrings();
    this.styles = options.useStyles ? new StylesXform(true) : new StylesXform.Mock(true);
    this._definedNames = createDefinedNames();
    this._worksheets = [];
    this.views = [];

    this.zipOptions = options.zip;
    this.format = options.format ?? "xlsx";
    this.xlsbTables = this.format === "xlsb" ? createStreamedXlsbTables() : undefined;
    const level = options.zip?.zlib?.level ?? options.zip?.compressionOptions?.level ?? 6;
    this.compressionLevel = Math.max(0, Math.min(9, level)) as
      | 0
      | 1
      | 2
      | 3
      | 4
      | 5
      | 6
      | 7
      | 8
      | 9;

    this.media = [];
    this.commentRefs = [];
    this.dynamicArrayCount = 0;
    this._trueStreaming = options.trueStreaming ?? false;

    // Create Zip instance.
    //
    // Backpressure note: when `this.stream.write(data)` returns false (the
    // user-supplied sink — e.g. fs.WriteStream, PassThrough, HTTP response
    // — has reached its highWaterMark), we cannot synchronously block the
    // zip callback (it's invoked from inside `row.commit()`'s sync chain).
    // Instead we record a `_needsDrain` flag and a Promise that resolves
    // when the sink emits `'drain'`. `commit()` and `_commitWorksheets()`
    // await this promise at their natural async boundaries, so the producer
    // stops generating new zip data until the sink has caught up. This
    // makes `WorkbookWriter` safe against slow sinks (network responses,
    // throttled fs, etc) without changing the public API.
    this.zip = new Zip((err, data, final) => {
      if (err) {
        this.stream.emit("error", err);
      } else {
        // `streaming-zip` already emits `Uint8Array`; avoid copying per chunk.
        const ok = this.stream.write(data);
        this._trackBackpressure(ok);
        if (final) {
          this.stream.end();
        }
      }
    });

    // Setup output stream
    this.stream = this._createOutputStream(options);

    // Eagerly attach error/close listeners on the sink so any backpressure
    // waiters are released the moment the sink fails — without this, a
    // `commit()` parked on `_waitForUserSinkDrain()` would hang forever if
    // the sink errored before emitting 'drain'.
    this._attachSinkLifecycleListeners();

    // Theme and office rels are deferred to commit() so that worksheet files
    // are added to the ZIP first. This ensures StreamingZip sets ondata on
    // the worksheet immediately, allowing pushSync to flow data through
    // without accumulating in _dataQueue.
    this.promise = Promise.resolve();
  }

  /**
   * Create output stream - can be overridden by Node.js to support filename
   *
   * @deprecated Internal; not part of the public API. Removed in 1.0.
   */
  protected _createOutputStream(options: WorkbookWriterOptions): OutputStreamLike {
    if (options.stream) {
      return Streams.toWritable(options.stream);
    }
    const output = new StreamBuf();
    this._ownedOutput = output;
    return output;
  }

  /**
   * Internal: record whether the sink accepted the last write. The
   * `OutputStreamLike.write` type advertises `boolean | Promise<boolean>`
   * for forward compatibility, but in practice every concrete sink we
   * accept (Node `Writable`, browser `Writable` from `@stream`, internal
   * `StreamBuf`, fs.WriteStream, etc) returns a sync `boolean`. We
   * defensively handle the Promise shape but it's never exercised.
   */
  private _trackBackpressure(ok: boolean | void | Promise<boolean>): void {
    if (ok && typeof (ok as PromiseLike<boolean>).then === "function") {
      // Do not let a drain check sample `_needsDrain` before an asynchronous
      // write has reported whether it applied backpressure.
      const generation = this._drainGeneration;
      const pending = Promise.resolve(ok).then(
        result => {
          // If drain fired while the Promise was pending, it already cleared
          // this write's backpressure; do not park waiting for a second event.
          if (!result && generation === this._drainGeneration) {
            this._needsDrain = true;
          }
        },
        err => {
          if (!this._sinkError) {
            this._sinkError = err instanceof Error ? err : new Error(String(err));
          }
          this._wakeAllBackpressureWaiters();
        }
      );
      this._pendingSinkWrites.add(pending);
      pending.finally(() => this._pendingSinkWrites.delete(pending)).catch(() => {});
      this._ensureDrainListener();
      return;
    }
    if (ok === false) {
      this._needsDrain = true;
    }
    this._ensureDrainListener();
  }

  private _ensureDrainListener(): void {
    if (this._drainListenerAttached) {
      return;
    }
    if (typeof this.stream.on !== "function") {
      // StreamBuf and similar sinks that don't follow the Writable contract
      // never emit 'drain'; they also never return false from write(), so
      // they reach this branch only spuriously. Skip listener attach.
      return;
    }
    this._drainListenerAttached = true;
    this.stream.on("drain", () => {
      this._drainGeneration++;
      this._needsDrain = false;
      callAllResolvers(this._drainResolvers);
    });
  }

  /**
   * Attach error/close listeners on the user sink so any parked backpressure
   * waiters are released the moment the sink fails. Without this, a
   * `commit()` parked on `_waitForUserSinkDrain()` would hang forever if
   * the sink errored before emitting 'drain'. Idempotent and a no-op for
   * sinks that don't expose `.on` (e.g. internal `StreamBuf`).
   *
   * Uses a non-consuming listener: if the user has their own 'error' handler
   * it still fires (EventEmitter broadcasts to all listeners). The error is
   * also captured into `_sinkError` so `_finalize()` can replay it — `_finalize`
   * registers its own listener with `once()`, which would miss errors that
   * arrived earlier in the commit pipeline.
   */
  private _lifecycleListenersAttached = false;

  private _attachSinkLifecycleListeners(): void {
    if (this._lifecycleListenersAttached) {
      return;
    }
    if (typeof this.stream.on !== "function") {
      return;
    }
    this._lifecycleListenersAttached = true;
    // Use `.once()` for both events: we only care about the first error
    // (subsequent errors are captured in `_sinkError` only if we haven't
    // recorded one yet). Using `.once()` also avoids leaking the listener
    // if the sink lives longer than the WorkbookWriter — the EventEmitter
    // releases the closure as soon as the event fires.
    if (typeof this.stream.once === "function") {
      this.stream.once("error", (err: Error) => {
        if (!this._sinkError) {
          this._sinkError = err;
        }
        this._wakeAllBackpressureWaiters();
      });
      this.stream.once("close", () => {
        this._wakeAllBackpressureWaiters();
      });
    } else {
      // Fallback: sink only has .on, attach normally.
      this.stream.on("error", (err: Error) => {
        if (!this._sinkError) {
          this._sinkError = err;
        }
        this._wakeAllBackpressureWaiters();
      });
    }
  }

  private _wakeAllBackpressureWaiters(): void {
    this._needsDrain = false;
    callAllResolvers(this._drainResolvers);
  }

  /**
   * Park here until any async writes have settled and the user sink has
   * drained below its high-water mark. Resolves immediately when no
   * backpressure is in flight.
   *
   * Called at async boundaries inside `commit()` so a slow sink throttles
   * the producer instead of letting bytes accumulate unboundedly.
   */
  private async _waitForUserSinkDrain(): Promise<void> {
    await this._waitForZipPushes();
    while (this._pendingSinkWrites.size > 0) {
      await Promise.all(this._pendingSinkWrites);
    }
    // Short-circuit if the sink already errored — no point waiting for a
    // drain that will never come. The error itself surfaces from
    // `_finalize()` later.
    if (this._sinkError) {
      return;
    }
    if (!this._needsDrain) {
      return;
    }
    return new Promise<void>(resolve => this._drainResolvers.push(resolve));
  }

  private _trackZipPush(push: Promise<void>): Promise<void> {
    this._pendingZipPushes.add(push);
    push.finally(() => this._pendingZipPushes.delete(push)).catch(() => {});
    return push;
  }

  private async _waitForZipPushes(): Promise<void> {
    while (this._pendingZipPushes.size > 0) {
      await Promise.all(this._pendingZipPushes);
    }
  }

  get definedNames(): DefinedNamesData {
    return this._definedNames;
  }

  /**
   * The default font for the workbook (fontId=0 / "Normal" style).
   * Must be set before any worksheet rows are committed.
   */
  get defaultFont(): Partial<Font> | undefined {
    return this.styles.defaultFont;
  }

  set defaultFont(font: Partial<Font> | undefined) {
    this._checkNotCommitted("set the default font");
    if (this.styles.setDefaultFont) {
      this.styles.setDefaultFont(font);
    }
  }

  /**
   * Add (or replace) a workbook-level named cell style such as "Heading 1" that streamed cells can
   * reference via their `styleName`, and return the stored entry — as `Workbook.addCellStyle` does.
   * Must be called before any worksheet rows referencing it are committed. Nothing is registered
   * when styles are disabled (`useStyles: false`). Rejects an empty name or the reserved "Normal"
   * style. A name equal to an earlier one ignoring case is that style: it is replaced, and keeps the
   * earlier spelling.
   */
  addCellStyle(name: string, style: NamedStyle): NamedStyleEntry {
    this._checkNotCommitted("add a cell style");
    validateCellStyleName(name);
    // Checked as `Workbook.addCellStyle` checks it: a missing style was spread into an empty one.
    const check = argumentChecks("WorkbookWriter.addCellStyle");
    check.required("style", style);
    check.record("style", style);
    checkStyle("WorkbookWriter.addCellStyle", style);
    const key = excelNameKey(name);
    const stored = this._cellStyleNames.get(key) ?? name;
    this._cellStyleNames.set(key, stored);
    const entry: NamedStyleEntry = { ...structuredClone(style), name: stored };
    if (this.styles.registerNamedStyles) {
      this.styles.registerNamedStyles(new Map([[key, entry]]));
    }
    return entry;
  }

  /**
   * Add one of Excel's built-in cell-style presets (e.g. "Heading1") and return its style name
   * (e.g. "Heading 1") for a cell's `styleName` — as `Workbook.useBuiltInCellStyle` does.
   */
  useBuiltInCellStyle(id: BuiltInCellStyle): string {
    this._checkNotCommitted("add a cell style");
    if (typeof id !== "string" || !Object.hasOwn(BUILTIN_CELL_STYLES, id)) {
      argumentChecks("WorkbookWriter.useBuiltInCellStyle").fail(
        "id",
        `one of ${Object.keys(BUILTIN_CELL_STYLES).join(", ")}`,
        id
      );
    }
    const { name, ...style } = BUILTIN_CELL_STYLES[id];
    return this.addCellStyle(name, style).name;
  }

  /**
   * The registered spelling of the named cell style `name` (compared without regard to case), or `undefined` when
   * {@link addCellStyle} has not registered it — what `Stream.applyCellStyle` checks against.
   * @internal
   */
  _cellStyleName(name: string): string | undefined {
    return this._cellStyleNames.get(excelNameKey(name));
  }

  /**
   * @deprecated Use {@link addCellStyle}, which returns the stored entry. Removed in 1.0.
   */
  defineCellStyle(name: string, style: NamedStyle): void {
    this.addCellStyle(name, style);
  }

  /** @internal */
  _openStream(path: string): InstanceType<typeof StreamBuf> {
    const stream = new StreamBuf({
      bufSize: this._trueStreaming ? 4096 : 65536,
      batch: !this._trueStreaming
    });

    const zipFile = new ZipDeflate(path, { level: this.compressionLevel });
    this.zip.add(zipFile);

    let resolveZipCompletion!: () => void;
    let rejectZipCompletion!: (err: Error) => void;
    const zipCompletion = new Promise<void>((resolve, reject) => {
      resolveZipCompletion = resolve;
      rejectZipCompletion = reject;
    });
    // A user may synchronously call worksheet.commit() long before wb.commit()
    // observes this promise. Keep a rejection handler attached in the interim.
    zipCompletion.catch(() => {});
    this._worksheetZipCompletions.set(stream, zipCompletion);

    const onData = (chunk: Uint8Array) => {
      this._trackZipPush(zipFile.push(chunk));
    };
    stream.on("data", onData);

    stream.once("finish", () => {
      stream.removeListener("data", onData);
      this._trackZipPush(zipFile.push(EMPTY_U8, true)).then(
        () => {
          resolveZipCompletion();
          stream.emit("zipped");
        },
        err => {
          rejectZipCompletion(err instanceof Error ? err : new Error(String(err)));
        }
      );
    });

    return stream;
  }

  /** @deprecated Internal; not part of the public API. Removed in 1.0. */
  protected _addFile(data: string | Uint8Array, name: string, base64?: boolean): void {
    const zipFile = new ZipDeflate(name, { level: this.compressionLevel });
    this.zip.add(zipFile);

    let buffer: Uint8Array;
    if (base64) {
      const base64Data = typeof data === "string" ? data : TEXT_DECODER.decode(data);
      buffer = base64ToUint8Array(base64Data);
    } else if (typeof data === "string") {
      buffer = stringToUint8Array(data);
    } else {
      buffer = data;
    }

    this._trackZipPush(zipFile.push(buffer, true));
  }

  private async _commitWorksheets(): Promise<void> {
    // Commit worksheets sequentially (not in parallel) so we can park on
    // user-sink backpressure between them. Parallel commit was the old
    // behavior; for a single-worksheet workbook the difference is nil, and
    // for multi-sheet workbooks honoring backpressure between them keeps
    // memory bounded against slow sinks. ZIP itself is inherently serial
    // (StreamingZip processes one entry at a time via `activeFile`), so
    // sequential commit imposes no real CPU cost — measured throughput is
    // identical to parallel commit on multi-sheet workbooks.
    for (const worksheet of this._worksheets) {
      if (!worksheet) {
        continue;
      }
      if (this.format === "xlsb") {
        // **`worksheet.stream` is not touched here.** It is a lazy getter that opens `xl/worksheets/sheetN.xml`
        // on first access — so reading it merely to find the completion promise created an XML entry the binary
        // package does not want, registered a promise nothing would ever resolve, and `commit()` hung. The
        // binary sheet's entry is opened by the sink instead, and its completion is awaited below with the rest.
        if (!worksheet.committed) {
          worksheet.commit();
        }
        continue;
      }
      const stream = worksheet.stream;
      const zipCompletion = this._worksheetZipCompletions.get(stream)!;
      if (!worksheet.committed) {
        worksheet.commit();
      }
      // This also covers a worksheet the user committed synchronously before
      // wb.commit(): its async browser deflate work still belongs to the writer.
      await zipCompletion;
      await this._waitForUserSinkDrain();
    }
    if (this.format === "xlsb") {
      // Every entry the sink opened, in one place. They were all started by `_openStream`, so they are all in the
      // same map the XML path reads one at a time — there is just no per-sheet stream object to key them by.
      await Promise.all(this._sinkCompletions);
      await this._waitForUserSinkDrain();
    }
  }

  /**
   * A `PackageSink` over this writer's ZIP.
   *
   * The two operations a sink needs are the two this class already had: `_addFile` writes a whole part and
   * `_openStream` opens an incremental one, both through the same `Zip` with the same compression level and the
   * same backpressure accounting. So the binary writer reaches the destination through the machinery that has
   * been carrying the XML one, rather than through a second ZIP layer beside it.
   *
   * @deprecated The sheet writer's view of the workbook while it encodes rows; nothing outside the writer needs it.
   * Removed in 1.0.
   */
  packageSink(): PackageSink {
    const written: string[] = [];
    return {
      part: (path: string, data: Uint8Array | string): void => {
        written.push(path);
        this._addFile(data, path);
      },
      open: (path: string): PartWriter => {
        written.push(path);
        const stream = this._openStream(path);
        // Kept in an array as well as the `WeakMap` keyed by stream, because a sink caller has no stream object to
        // look one up with — and the map is weak, so it cannot be enumerated.
        const completion = this._worksheetZipCompletions.get(stream);
        if (completion !== undefined) {
          this._sinkCompletions.push(completion);
        }
        return {
          write: (chunk: Uint8Array | string): void => {
            stream.write(chunk as never);
          },
          end: (): void => {
            stream.end();
          }
        };
      },
      get paths(): readonly string[] {
        return written;
      },
      drain: async (): Promise<void> => {
        await this._waitForUserSinkDrain();
      }
    };
  }

  /**
   * Finish the package: commit every sheet still open, write the workbook-level parts and the ZIP central directory,
   * and end the output — the `stream` or `filename` the writer was given, or its own buffer, which {@link toBuffer}
   * then returns.
   *
   * @throws {ExcelStreamStateError} If the workbook was already committed.
   */
  async commit(): Promise<void> {
    this._checkNotCommitted("commit the workbook");
    this._workbookCommitted = true;
    this._commitDone = this._commit();
    await this._commitDone;
  }

  /**
   * The complete file, when the writer was given no `stream` and no `filename`. Commits first if {@link commit} has
   * not been called (and waits for one already running); afterwards it may be called again and returns the same bytes.
   *
   * @throws {ExcelStreamStateError} If the writer was given a `stream` or `filename` — the package went there and the
   * writer holds none — or committing failed.
   */
  async toBuffer(): Promise<Uint8Array> {
    const output = this._ownedOutput;
    if (!output) {
      throw new ExcelStreamStateError(
        "get the workbook bytes",
        "the package was written to the stream or file given in the options"
      );
    }
    await (this._commitDone ?? this.commit());
    const bytes = output.toBuffer();
    if (!bytes) {
      throw new ExcelStreamStateError("get the workbook bytes", "the writer produced no output");
    }
    return bytes;
  }

  /** The output the writer created because it was given none; `toBuffer` reads it. */
  private _ownedOutput: InstanceType<typeof StreamBuf> | undefined;
  /** The running or settled `commit`. */
  private _commitDone: Promise<void> | undefined;

  private async _commit(): Promise<void> {
    await this.promise;
    await this._commitWorksheets();
    if (this.format === "xlsb") {
      // **Everything except the sheet parts, from the same writer the buffered path uses.**
      //
      // The sheets are already in the ZIP — streamed row by row — and `writeXlsbPackage` is told so through
      // `streamed`, along with the interning tables their records index into. It then produces the workbook,
      // styles, shared strings, relationships, content types, drawings and media exactly as it does for
      // `Workbook.toBuffer`, which is why a streamed package and a buffered one differ in one record rather than
      // in a hundred small ways.
      const { writeXlsbPackage } = await import("@excel/xlsb/write/package");
      const model = getWorkbookModel(this as never);
      // A streamed medium's name already carries its extension (`image0.png`, which the XLSX path writes as is),
      // while a model medium's does not — the package adds it, and was writing `image0.png.png`.
      model.media = this.media.map(medium =>
        medium.name.endsWith(`.${medium.extension}`)
          ? { ...medium, name: medium.name.slice(0, -medium.extension.length - 1) }
          : medium
      ) as never;
      // The sheets as the package must see them — with a background among their media, see `xlsbModel`.
      for (const sheetModel of model.worksheets) {
        const sheet = this._worksheets.find(candidate => candidate?.id === sheetModel.id);
        if (sheet?.xlsbModel !== undefined) {
          sheetModel.media = sheet.xlsbModel().media as never;
        }
      }
      const written = await writeXlsbPackage(model, {
        sink: this.packageSink(),
        streamed: {
          sheetPaths: this.xlsbTables!.sheetPaths,
          strings: this.xlsbTables!.strings,
          formats: this.xlsbTables!.formats,
          names: streamedFormulaNames(this.xlsbTables!),
          features: this.xlsbTables!.features,
          externSheets: this.xlsbTables!.externSheets,
          dxfs: this.xlsbTables!.dxfs,
          cfPriorities: this.xlsbTables!.cfPriorities
        }
      });
      // **Both halves of the report, and the row half used to be dropped.**
      //
      // `writeXlsbPackage` sees the workbook after its sheets have been streamed, so it can only report what the
      // *package* could not express. Anything a row could not express was reported by the row encoder at the time and
      // kept on the worksheet — where nothing read it, so a cross-sheet formula that became a blank cell left
      // `xlsbUnsupported` empty. A loss that is computed and then discarded is worse than one that is never computed:
      // it reads as a guarantee.
      //
      // Reported, not thrown: the streaming writer has no `unsupported` option and a refusal here would abort a workbook
      // whose rows are already in the caller's stream. Nothing can be taken back at this point.
      const sheetLosses = this._worksheets.flatMap(sheet => [...(sheet?.xlsbUnsupported ?? [])]);
      const all = [...sheetLosses, ...written.unsupported];
      if (all.length > 0) {
        this._xlsbUnsupported = all;
      }
      const reported = [
        ...this._worksheets.flatMap(sheet => [...(sheet?.unsupported ?? [])]),
        ...written.unsupported
      ];
      if (reported.length > 0) {
        this.unsupported = reported;
      }
      await this._waitForUserSinkDrain();
      await this._finalize();
      return;
    }
    // The rows' losses, which the XLSX path has too: an unnamed this-row reference written as its cached value.
    const rowLosses = this._worksheets.flatMap(sheet => [...(sheet?.unsupported ?? [])]);
    if (rowLosses.length > 0) {
      this.unsupported = rowLosses;
    }
    await this.addMedia();
    await this.addDrawings();
    await this._waitForUserSinkDrain();
    await Promise.all([
      this.addThemes(),
      this.addOfficeRels(),
      this.addContentTypes(),
      this.addApp(),
      this.addCore(),
      this.addSharedStrings(),
      this.addStyles(),
      this.addFeaturePropertyBag(),
      this.addMetadata(),
      this.addWorkbookRels()
    ]);
    await this._waitForUserSinkDrain();
    await this.addWorkbook();
    await this._waitForUserSinkDrain();
    await this._finalize();
  }

  /**
   * @deprecated The id `addWorksheet` gives the next sheet; nothing outside the writer needs it. Removed in 1.0.
   */
  get nextId(): number {
    return this._nextSheetId();
  }

  private _nextSheetId(): number {
    for (let i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  }

  /**
   * Register an image with the workbook and return its numeric id.
   *
   * Supply `buffer`/`base64`/`filename` to **embed** the bytes, or only `link`
   * (a URL or local file path) to reference it **externally** — in which case
   * no bytes are written into the package and the relationship is emitted with
   * `TargetMode="External"`. If both are provided, embedding wins.
   *
   * Linked images work with cell pictures and overlay watermarks; worksheet
   * background images and header/footer (VML) watermarks cannot be linked.
   *
   * @example
   * ```typescript
   * const id = wb.addImage({ extension: "png", link: "https://example.com/logo.png" });
   * ws.addImage(id, "B2:D6");
   * ```
   */
  addImage(image: ImageData): number {
    this._checkNotCommitted("add an image");
    checkImageData("WorkbookWriter.addImage", image);
    const { svg, ...raster } = image;
    if (
      svg &&
      raster.link &&
      raster.buffer == null &&
      raster.base64 == null &&
      raster.filename == null
    ) {
      throw new ExcelImageError(
        "An SVG image requires an embedded raster fallback (buffer/base64/filename); it cannot be combined with an external link."
      );
    }
    const id = this.media.length;
    const medium: Medium = {
      ...raster,
      type: "image" as const,
      name: `image${id}.${raster.extension}`
    };
    this.media.push(medium);

    if (svg) {
      // Register the SVG companion as a second image medium and link it back to
      // the raster blip so the drawing serializer emits the svgBlip extension.
      const svgId = this.media.length;
      this.media.push({
        ...svg,
        type: "image" as const,
        extension: "svg",
        name: `image${svgId}.svg`
      });
      medium.svgMediaId = svgId;
    }

    return id;
  }

  getImage(id: number): Medium | undefined {
    return this.media[id];
  }

  /**
   * Protect the workbook structure with an optional password — the same options object
   * `Workbook.protect(wb, { password, ...flags })` takes. Prevents users from adding, deleting,
   * renaming, moving, or copying worksheets.
   */
  protect(options?: ProtectWorkbookOptions): Promise<void>;
  /**
   * @deprecated Pass the password in the options object: `protect({ password, ...options })`.
   * Removed in 1.0.
   */
  protect(password: string | undefined, options?: Partial<WorkbookProtection>): Promise<void>;
  async protect(
    passwordOrOptions?: string | ProtectWorkbookOptions,
    legacyOptions?: Partial<WorkbookProtection>
  ): Promise<void> {
    this._checkNotCommitted("protect the workbook");
    // Checked as `Workbook.protect` checks it: `{ lockStructure: "yes" }` used to be written as locked.
    checkProtectCall(
      "WorkbookWriter.protect",
      passwordOrOptions,
      legacyOptions,
      WORKBOOK_PROTECTION_FLAGS
    );
    const [password, options] = splitPasswordOptions(passwordOrOptions, legacyOptions);
    this.protection = await buildWorkbookProtection(password, options);
  }

  /**
   * Remove workbook structure protection.
   */
  unprotect(): void {
    this._checkNotCommitted("unprotect the workbook");
    this.protection = undefined;
  }

  /** @internal Set when `commit` starts: the package is being finished, and nothing more can be added to it. */
  _workbookCommitted = false;

  /**
   * Refuse an edit once `commit` has started, with the error a committed sheet raises — a second commit, or a sheet
   * added afterwards, used to fail inside the ZIP writer with an `ArchiveError`.
   */
  private _checkNotCommitted(operation: string): void {
    if (this._workbookCommitted) {
      throw new ExcelStreamStateError(operation, "the workbook has been committed");
    }
  }

  addWorksheet(name?: string, options?: Partial<AddWorksheetOptions>): TWorksheetWriter {
    this._checkNotCommitted("add a worksheet");
    const opts = options || {};
    const useSharedStrings =
      opts.useSharedStrings !== undefined ? opts.useSharedStrings : this.useSharedStrings;

    // `tabColor` was a top-level option in older releases; detect the legacy
    // shape and migrate it into `properties`.
    const legacyTabColor = (opts as { tabColor?: WorksheetProperties["tabColor"] }).tabColor;
    if (legacyTabColor) {
      console.trace("tabColor option has moved to { properties: tabColor: {...} }");
      opts.properties = { tabColor: legacyTabColor, ...opts.properties };
    }

    const id = this._nextSheetId();
    name = name ?? `sheet${id}`;
    // The document model's rules, so a streamed workbook can be read back: a repeated or illegal name was written as
    // given, and the file could not be opened — not even by `Workbook.read`.
    refuseLongSheetName(name);
    validateSheetName(
      {
        _worksheets: [],
        _chartsheets: this._worksheets.filter(Boolean).map(sheet => ({ name: sheet.name }))
      } as unknown as Parameters<typeof validateSheetName>[0],
      name
    );

    const worksheet = new this.WorksheetWriterClass({
      id,
      name,
      workbook: this,
      useSharedStrings,
      properties: opts.properties,
      state: opts.state,
      pageSetup: opts.pageSetup,
      views: opts.views,
      autoFilter: opts.autoFilter,
      headerFooter: opts.headerFooter
    });

    // **`orderNo` is assigned here, and was never assigned at all.**
    //
    // `WorksheetData` declares it as a required `number`, and `getWorksheets` sorts by `a.orderNo - b.orderNo` — so for
    // a streaming writer, whose sheets masquerade as `WorksheetData`, every comparison was `NaN` and `Array.sort` fell
    // back to engine-defined behaviour. The order came out right by accident: `sheetsInTabOrder` falls back to `sheetNo`,
    // which a streaming sheet derives from its `id`. Three layers of fallback standing in for the one field that means
    // tab order.
    //
    // Sequential from the id because a streaming writer has no chartsheets to interleave with — `addChartsheet` is not
    // part of its surface — so the id *is* the tab position, and stating it makes the declared field true.
    (worksheet as { orderNo?: number }).orderNo = id - 1;
    this._worksheets[id] = worksheet;
    return worksheet;
  }

  getWorksheet(id?: string | number): TWorksheetWriter | undefined {
    if (id === undefined) {
      return this._worksheets.find(() => true);
    }
    if (typeof id === "number") {
      return this._worksheets[id];
    }
    if (typeof id === "string") {
      const idKey = sheetNameKey(id);
      return this._worksheets.find(ws => ws?.name !== undefined && sheetNameKey(ws.name) === idKey);
    }
    return undefined;
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addStyles(): Promise<void> {
    return new Promise(resolve => {
      this._addFile(this.styles.xml, OOXML_PATHS.xlStyles);
      resolve();
    });
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addThemes(): Promise<void> {
    return new Promise(resolve => {
      this._addFile(theme1Xml, OOXML_PATHS.xlTheme1);
      resolve();
    });
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addOfficeRels(): Promise<void> {
    return new Promise(resolve => {
      const xform = new RelationshipsXform();
      const xml = xform.toXml([
        { Id: "rId1", Type: RelType.OfficeDocument, Target: OOXML_PATHS.xlWorkbook },
        { Id: "rId2", Type: RelType.CoreProperties, Target: OOXML_PATHS.docPropsCore },
        { Id: "rId3", Type: RelType.ExtenderProperties, Target: OOXML_PATHS.docPropsApp }
      ]);
      this._addFile(xml, OOXML_PATHS.rootRels);
      resolve();
    });
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addContentTypes(): Promise<void> {
    return new Promise(resolve => {
      const worksheets = this._worksheets.filter(Boolean);
      // In the streaming path, ZIP entries use ws.id which is always sequential.
      // Set fileIndex = id to satisfy the ContentTypesXform contract.
      worksheets.forEach(ws => {
        ws.fileIndex = ws.id;
      });

      // Collect drawing models from worksheets that have images
      const drawings = worksheets.filter(ws => ws.drawing).map(ws => ws.drawing);

      const model = {
        worksheets,
        sharedStrings: this.sharedStrings,
        commentRefs: this.commentRefs,
        media: this.media,
        drawings,
        hasCheckboxes: this.styles.hasCheckboxes,
        hasDynamicArrayFormulas: this.dynamicArrayCount > 0
      };
      const xform = new ContentTypesXform();
      this._addFile(xform.toXml(model), OOXML_PATHS.contentTypes);
      resolve();
    });
  }

  /**
   * Add media files - can be overridden by Node.js for file system support
   *
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes the media a second time. Use {@link commit}. Removed in 1.0.
   */
  addMedia(): Promise<void[]> {
    return Promise.all(
      this.media.map(async medium => {
        if (medium.type === "image") {
          // External (linked) images carry only a `link` target — no bytes
          // are written into the package.
          if (isExternalImage(medium)) {
            return;
          }
          const filename = mediaPath(medium.name);
          if (medium.buffer) {
            this._addFile(medium.buffer, filename);
            return;
          }
          if (medium.base64) {
            const content = medium.base64.substring(medium.base64.indexOf(",") + 1);
            this._addFile(content, filename, true);
            return;
          }
          if (medium.filename) {
            throw new ExcelUnsupportedFeatureError(
              "Loading images from filename",
              "not supported in browser. Use buffer or base64."
            );
          }
        }
        throw new ExcelImageError("Unsupported media");
      })
    );
  }

  /**
   * Generate drawing XML and drawing relationship files for worksheets that have images.
   * Must be called after _commitWorksheets() so that each WorksheetWriter has built its
   * drawing model, and after addMedia() so that media files are already in the ZIP.
   *
   * `DrawingXform` is imported on demand, matching how `XLSX` already reaches for
   * it. A static import here defeated that: bundlers hoist a statically imported
   * module into the main chunk, so every `await import()` of it elsewhere became
   * ineffective (`INEFFECTIVE_DYNAMIC_IMPORT`) and ~49 KB of drawing transformers
   * were pulled in even by a consumer that only ever writes cells.
   *
   * The emptiness check comes first so a workbook with no images does not load the
   * module at all — the previous version constructed a `DrawingXform` even when
   * there was nothing to write.
   *
   * @deprecated Internal; not part of the public API. Removed in 1.0.
   */
  protected async addDrawings(): Promise<void> {
    const drawings = this._worksheets
      .map(ws => ws?.drawing)
      .filter((drawing): drawing is NonNullable<typeof drawing> => drawing != null);
    if (drawings.length === 0) {
      return;
    }

    const { DrawingXform } = await import("@excel/xlsx/xform/drawing/drawing-xform");
    const drawingXform = new DrawingXform();
    const relsXform = new RelationshipsXform();

    for (const drawing of drawings) {
      // Filter out invalid anchors using shared utility
      const filteredAnchors = filterDrawingAnchors(drawing.anchors);
      const drawingForWrite = { ...drawing, anchors: filteredAnchors };

      // Prepare and generate drawing XML
      drawingXform.prepare(drawingForWrite);
      const xml = drawingXform.toXml(drawingForWrite);
      this._addFile(xml, drawingPath(drawing.name));

      // Generate drawing relationships
      const relsXml = relsXform.toXml(drawing.rels);
      this._addFile(relsXml, drawingRelsPath(drawing.name));
    }
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addApp(): Promise<void> {
    return new Promise(resolve => {
      const xform = new AppXform();
      this._addFile(
        xform.toXml({ worksheets: this._worksheets.filter(Boolean) }),
        OOXML_PATHS.docPropsApp
      );
      resolve();
    });
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addCore(): Promise<void> {
    return new Promise(resolve => {
      const xform = new CoreXform();
      this._addFile(xform.toXml(this), OOXML_PATHS.docPropsCore);
      resolve();
    });
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addSharedStrings(): Promise<void> {
    if (this.sharedStrings.count) {
      return new Promise(resolve => {
        const xform = new SharedStringsXform();
        this._addFile(xform.toXml(this.sharedStrings), OOXML_PATHS.xlSharedStrings);
        resolve();
      });
    }
    return Promise.resolve();
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addFeaturePropertyBag(): Promise<void> {
    if (this.styles.hasCheckboxes) {
      const xform = new FeaturePropertyBagXform();
      this._addFile(xform.toXml({}), OOXML_PATHS.xlFeaturePropertyBag);
    }
    return Promise.resolve();
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addMetadata(): Promise<void> {
    if (this.dynamicArrayCount <= 0) {
      return Promise.resolve();
    }
    const xform = new MetadataXform();
    this._addFile(
      xform.toXml({ dynamicArrayCount: this.dynamicArrayCount }),
      OOXML_PATHS.xlMetadata
    );
    return Promise.resolve();
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addWorkbookRels(): Promise<void> {
    let count = 1;
    const relationships: Array<{ Id: string; Type: string; Target: string }> = [
      { Id: `rId${count++}`, Type: RelType.Styles, Target: OOXML_REL_TARGETS.workbookStyles },
      { Id: `rId${count++}`, Type: RelType.Theme, Target: OOXML_REL_TARGETS.workbookTheme1 }
    ];
    if (this.sharedStrings.count) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.SharedStrings,
        Target: OOXML_REL_TARGETS.workbookSharedStrings
      });
    }
    // Add FeaturePropertyBag relationship if checkboxes are used
    if (this.styles.hasCheckboxes) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.FeaturePropertyBag,
        Target: OOXML_REL_TARGETS.workbookFeaturePropertyBag
      });
    }
    // Add metadata relationship for dynamic array formulas
    if (this.dynamicArrayCount > 0) {
      relationships.push({
        Id: `rId${count++}`,
        Type: RelType.SheetMetadata,
        Target: OOXML_REL_TARGETS.workbookMetadata
      });
    }
    this._worksheets.forEach(ws => {
      if (ws) {
        ws.rId = `rId${count++}`;
        relationships.push({
          Id: ws.rId,
          Type: RelType.Worksheet,
          Target: worksheetRelTarget(ws.id)
        });
      }
    });

    return new Promise(resolve => {
      const xform = new RelationshipsXform();
      this._addFile(xform.toXml(relationships), OOXML_PATHS.xlWorkbookRels);
      resolve();
    });
  }

  /**
   * @deprecated One of the steps {@link commit} runs to write the package, and only correct there: called on its own it
   * writes its part a second time. Use {@link commit}. Removed in 1.0.
   */
  addWorkbook(): Promise<void> {
    // A name whose definition Excel would not open is left out and reported, as the buffered writer does.
    const lost: string[] = [];
    const definedNames = writableDefinedNames(
      definedNamesModel(this._definedNames),
      unwritableFormulaReason,
      lost
    );
    if (lost.length > 0) {
      this.unsupported = [...this.unsupported, ...lost];
    }
    const model = {
      worksheets: this._worksheets.filter(Boolean),
      definedNames,
      views: this.views,
      // **`this.properties`, not `{}`.** `date1904` is a declared writer option and reaches the XLSB row encoder
      // through `date1904Flag()`, but this — the part that states the workbook's date system to every reader —
      // was built from a literal empty object, so no `<workbookPr date1904="1">` was ever written on the XLSX
      // path and the option was inert in one of the two formats it is offered for.
      properties: this.properties,
      protection: this.protection,
      calcProperties: {}
    };
    return new Promise(resolve => {
      const xform = new WorkbookXform();
      xform.prepare(model);
      this._addFile(
        withFormulaNames(formulaNameSet(model.definedNames), () => xform.toXml(model)),
        OOXML_PATHS.xlWorkbook
      );
      resolve();
    });
  }

  private _finalize(): Promise<this> {
    // If the user sink errored earlier in the commit pipeline (captured by
    // `_attachSinkLifecycleListeners`), surface that error now — `commit()`
    // would otherwise reach `_finalize` and hang waiting for `'close'` from
    // a sink that's already destroyed.
    if (this._sinkError) {
      // End the zip pipeline cleanly so its internal callbacks don't keep
      // firing into a torn-down sink. Best-effort: ignore any error from
      // end() since the original `_sinkError` is what we care about.
      try {
        this.zip.end();
      } catch {
        // Best-effort cleanup.
      }
      return Promise.reject(this._sinkError);
    }

    // Wait for "close" — emitted by all supported output streams (Node Writable,
    // browser Writable, and StreamBuf) after "finish". For file streams this
    // guarantees the fd is released, which is critical on Windows where reading
    // a file before fd close can see truncated content.
    return new Promise((resolve, reject) => {
      const cleanup = () => {
        this.stream.removeListener("close", onDone);
        this.stream.removeListener("error", onError);
      };
      const onError = (err: Error) => {
        cleanup();
        reject(err);
      };
      const onDone = () => {
        cleanup();
        // If an error fired between us checking `_sinkError` and reaching
        // 'close' (rare but possible with concurrent emit), surface it.
        if (this._sinkError) {
          reject(this._sinkError);
          return;
        }
        resolve(this);
      };
      this.stream.once("error", onError);
      this.stream.once("close", onDone);
      this.zip.end();
    });
  }
}

// ============================================================================
// Browser-compatible WorkbookWriter
// ============================================================================

/**
 * Write a workbook a row at a time to a stream, without holding the sheets in memory: `addWorksheet` hands out a
 * worksheet writer, each committed row is written and released, and `commit` finishes the package.
 */
class WorkbookWriter extends WorkbookWriterBase<WorksheetWriter<WorkbookWriter>> {
  constructor(options: WorkbookWriterOptions = {}) {
    super(options, WorksheetWriter);
  }
}

export { WorkbookWriter };
