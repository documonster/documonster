import { ZipParser } from "@archive/unzip/zip-parser";
import type { ChartExEntry } from "@excel/chart/model/chart-ex-types";
import { parseChartEx } from "@excel/chart/serialize/chart-ex-parser";
/**
 * The XLSX reader: ZIP entries in, a populated workbook out.
 *
 * Reached only through `await import()` from `core/xlsx-io`, the same boundary `core/workbook-format.ts`
 * draws for XLSB, so a consumer that only writes never downloads it. Nothing in `write/` may import from
 * here — `scripts/treeshake-verify.ts` asserts that a write-only bundle holds no reader module.
 */
import { loadWorkbookModel } from "@excel/core/model-load";
import type { Workbook } from "@excel/core/workbook.browser";
import { ExcelFileError, XlsxParseError } from "@excel/errors";
import {
  chartsheetPath,
  getChartsheetNoFromPath,
  getChartsheetNoFromRelsPath,
  drawingPath,
  isCommentsPath,
  getChartExNumberFromPath,
  getChartExNumberFromRelsPath,
  getChartNumberFromPath,
  getChartNumberFromRelsPath,
  getChartStyleNumberFromPath,
  getChartColorsNumberFromPath,
  getChartExStyleNumberFromPath,
  getChartExColorsNumberFromPath,
  getDrawingNameFromPath,
  getChartUserShapesNameFromPath,
  getDrawingNameFromRelsPath,
  getExternalLinkIndexFromPath,
  getExternalLinkIndexFromRelsPath,
  getMediaFilenameFromPath,
  getPivotCacheDefinitionNameFromPath,
  getPivotCacheDefinitionNameFromRelsPath,
  getPivotCacheRecordsNameFromPath,
  getPivotTableNameFromPath,
  getPivotTableNameFromRelsPath,
  getTableNameFromPath,
  getThemeNameFromPath,
  getVmlDrawingNameFromPath,
  getVmlDrawingHFNameFromPath,
  getVmlDrawingHFNameFromRelsPath,
  getWorksheetNoFromWorksheetPath,
  getWorksheetNoFromWorksheetRelsPath,
  isBinaryEntryPath,
  normalizeZipPath,
  OOXML_PATHS,
  worksheetPath
} from "@excel/utils/ooxml-paths";
import { snapshotChartModel } from "@excel/xlsx/chart-snapshot";
import type { OpaqueDrop, OpaquePart, OpaqueRelationship } from "@excel/xlsx/opaque-parts";
import { isRelationshipsPart, ownerOfRelationshipsPart } from "@excel/xlsx/opaque-parts";
import type { ExternalLinkRelsEntry } from "@excel/xlsx/read/reconcile";
import { reconcile } from "@excel/xlsx/read/reconcile";
import type { XlsxOptions, XlsxReadOptions, ZipEntryLike } from "@excel/xlsx/types";
import type { ParsedExternalLink } from "@excel/xlsx/xform/book/external-link-xform";
import { ExternalLinkXform } from "@excel/xlsx/xform/book/external-link-xform";
import { WorkbookXform } from "@excel/xlsx/xform/book/workbook-xform";
import { ChartSpaceXform } from "@excel/xlsx/xform/chart/chart-space-xform";
import {
  parsePersonList,
  parseThreadedComments
} from "@excel/xlsx/xform/comment/threaded-comments-parse";
import { AppXform } from "@excel/xlsx/xform/core/app-xform";
import { ContentTypesXform } from "@excel/xlsx/xform/core/content-types-xform";
import { CoreXform } from "@excel/xlsx/xform/core/core-xform";
import { MetadataXform } from "@excel/xlsx/xform/core/metadata-xform";
import type { RelationshipModel } from "@excel/xlsx/xform/core/relationship-xform";
import { RelationshipsXform } from "@excel/xlsx/xform/core/relationships-xform";
import { parseXformStream } from "@excel/xlsx/xform/parse-xform";
import { WorkSheetXform } from "@excel/xlsx/xform/sheet/worksheet-xform";
import { SharedStringsXform } from "@excel/xlsx/xform/strings/shared-strings-xform";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { concatUint8Arrays } from "@utils/binary";
import { bufferToString, base64ToUint8Array } from "@utils/utils";

/**
 * One package entry as every handler below consumes it.
 *
 * `stream` is what an xform parses; `bytes()` is the entry's exact content, for the parts that are kept
 * verbatim — media, chart and chartEx XML, preserved parts. A handler uses one or the other: on the
 * streaming path the two share one ZIP entry stream, which can be read once.
 *
 * Both readers build these and hand them to the same {@link processEntry}; they differ only in where the
 * bytes come from. The buffered reader used to wrap bytes it already held in a `PassThrough` and read them
 * back through stream events, and thread the original bytes alongside as a `rawData` argument because the
 * round trip could not be trusted to reproduce them.
 */
interface XlsxPart {
  readonly stream: AsyncIterable<string | Uint8Array> | Iterable<string | Uint8Array>;
  bytes(): Promise<Uint8Array>;
}

/**
 * A part over bytes already in memory. XML is decoded once with the lenient decoder this reader has
 * always used on this path, so an invalid sequence is replaced rather than rejected; media and themes are
 * passed through as bytes.
 */
function bufferedPart(entryName: string, data: Uint8Array): XlsxPart {
  const chunk = isBinaryEntryPath(entryName) ? data : bufferToString(data);
  return { stream: [chunk], bytes: () => Promise.resolve(data) };
}

/** A part over a ZIP entry stream being read as it arrives. `bytes()` drains it once and keeps the result. */
function streamedPart(stream: AsyncIterable<string | Uint8Array>): XlsxPart {
  let collected: Promise<Uint8Array> | undefined;
  return {
    stream,
    bytes() {
      collected ??= collectBytes(stream);
      return collected;
    }
  };
}

async function collectBytes(stream: AsyncIterable<string | Uint8Array>): Promise<Uint8Array> {
  const encoder = new TextEncoder();
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream) {
    chunks.push(typeof chunk === "string" ? encoder.encode(chunk) : chunk);
  }
  return concatUint8Arrays(chunks);
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
export async function readXlsxInto(
  workbook: Workbook,
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

  return readXlsxBytesInto(workbook, buffer, options);
}

/**
 * Document-level container defaults. The archive library leaves both bounds
 * unbounded because what a process can afford is host-dependent; a workbook is
 * a narrower thing. A part count of 10,000 mirrors the Word reader's
 * `maxPartCount` and is far above real workbooks (each sheet adds ~2–6 parts:
 * sheet XML, rels, drawing, comments, table), which reach the low hundreds.
 * 2 GiB total sits above what Excel itself can open — sheet XML compresses
 * ~10:1, so this admits a ~200 MB `.xlsx` — while rejecting the many-entries
 * bombs a per-entry bound alone lets through. Pass `Infinity` to lift either.
 */
const XLSX_DEFAULT_MAX_ENTRIES = 10_000;
const XLSX_DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE = 2 * 1024 * 1024 * 1024;

/**
 * Internal: Load from Uint8Array buffer
 */
export async function readXlsxBytesInto(
  workbook: Workbook,
  buffer: Uint8Array,
  options?: XlsxReadOptions
): Promise<Workbook> {
  const parser = new ZipParser(buffer, {
    ...options?.zip,
    maxEntries: options?.zip?.maxEntries ?? XLSX_DEFAULT_MAX_ENTRIES,
    maxTotalUncompressedSize:
      options?.zip?.maxTotalUncompressedSize ?? XLSX_DEFAULT_MAX_TOTAL_UNCOMPRESSED_SIZE
  });
  const filesMap = await parser.extractAll();

  // Convert Map to Record for readXlsxFilesInto
  const allFiles: Record<string, Uint8Array> = {};
  for (const [path, content] of filesMap) {
    allFiles[path] = content;
  }

  return readXlsxFilesInto(workbook, allFiles, options);
}

/**
 * Create an empty model for parsing XLSX files.
 * Shared by readXlsxEntriesInto and readXlsxFilesInto.
 */
function createEmptyModel(): any {
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
    vmlDrawingHF: {},
    vmlDrawingHFRels: {},
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
    chartExStyles: {} as Record<number, Uint8Array>,
    chartExColors: {} as Record<number, Uint8Array>,
    // Raw chartEx entries (Office 2016+ extended charts) keyed by chartEx number
    chartExEntries: {} as Record<number, Uint8Array>,
    // Parsed chartEx rels keyed by chartEx number
    chartExRels: {} as Record<number, any[]>,
    // Structured chartEx entries (built via addChartEx) keyed by chartEx number
    chartExStructuredEntries: {} as Record<number, ChartExEntry>,
    // External workbook links — parsed from xl/externalLinks/externalLinkN.xml
    // during processDefaultEntry, then reconciled into a dense
    // ExternalLinkModel[] by reconcile() using workbookRels + <externalReferences>.
    externalLinksByIndex: {} as Record<number, ParsedExternalLink>,
    // Raw rels from each externalLinkN.rels file, keyed by index.
    // Contains the actual Target path (e.g. "测试.xlsx", "file:///...")
    // and TargetMode ("External" / "Internal").
    externalLinkRelsByIndex: {} as Record<number, ExternalLinkRelsEntry[]>,
    // Chartsheets keyed by sheet number
    chartsheets: {} as Record<number, any>,
    chartsheetRels: {} as Record<number, any[]>,
    // Staging for opaque (unmodelled) parts. A ZIP imposes no entry order, so
    // bytes, content types and relationships are gathered independently while
    // walking the package and joined in reconcile() once all three are known.
    opaqueUnknownEntries: new Map<string, Uint8Array>(),
    opaqueContentTypeOverrides: new Map<string, string>(),
    opaqueContentTypeDefaults: {} as Record<string, string>,
    opaqueRelationshipsBySource: new Map<string, OpaqueRelationship[]>(),
    opaqueParts: [] as OpaquePart[],
    opaqueDrops: [] as OpaqueDrop[]
  };
}

/**
 * Process a known OOXML entry (workbook, styles, shared strings, etc.)
 * Returns true if handled, false if should be passed to processDefaultEntry
 */
async function processKnownEntry(
  part: XlsxPart,
  model: any,
  entryName: string,
  options?: XlsxOptions
): Promise<boolean> {
  const sheetNo = getWorksheetNoFromWorksheetPath(entryName);
  if (sheetNo !== undefined) {
    await processWorksheetEntry(part, model, sheetNo, options, entryName);
    return true;
  }

  const chartsheetNo = getChartsheetNoFromPath(entryName);
  if (chartsheetNo !== undefined) {
    await processChartsheetEntry(part, model, chartsheetNo);
    return true;
  }

  switch (entryName) {
    case OOXML_PATHS.rootRels:
      model.globalRels = await parseRelsPart(part, model, "");
      return true;
    case OOXML_PATHS.xlWorkbook: {
      const workbook = await parseWorkbookPart(part);
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
      await parseXformStream(model.sharedStrings, part.stream);
      return true;
    case OOXML_PATHS.xlWorkbookRels:
      model.workbookRels = await parseRelsPart(part, model, OOXML_PATHS.xlWorkbook);
      return true;
    case OOXML_PATHS.contentTypes: {
      // Capture the Override ContentType for /xl/workbook.xml. Templates
      // (.xltx/.xltm) and macro-enabled workbooks (.xlsm) declare a
      // different value than a plain .xlsx; losing it makes Excel refuse
      // the file because the declared content-type no longer matches the
      // (preserved) extension. Parsed structurally via ContentTypesXform
      // rather than by hand so attribute order/quoting variance is a
      // non-issue.
      const parsed = await parseXformStream(new ContentTypesXform(), part.stream);
      if (parsed?.workbookContentType) {
        model.workbookContentType = parsed.workbookContentType;
      }
      for (const [partPath, contentType] of Object.entries(parsed?.overrides ?? {})) {
        model.opaqueContentTypeOverrides.set(partPath, contentType);
      }
      for (const [extension, contentType] of Object.entries(parsed?.defaults ?? {})) {
        model.opaqueContentTypeDefaults[extension] = contentType;
      }
      return true;
    }
    case OOXML_PATHS.docPropsApp: {
      const appXform = new AppXform();
      const appProperties = await parseXformStream(appXform, part.stream);
      if (appProperties) {
        model.company = appProperties.company;
        model.manager = appProperties.manager;
      }
      return true;
    }
    case OOXML_PATHS.docPropsCore: {
      const coreXform = new CoreXform();
      const coreProperties = await parseXformStream(coreXform, part.stream);
      Object.assign(model, coreProperties);
      return true;
    }
    case OOXML_PATHS.xlStyles:
      model.styles = new StylesXform();
      await parseXformStream(model.styles, part.stream);
      return true;
    case OOXML_PATHS.xlMetadata: {
      const metadataXform = new MetadataXform();
      const metadataResult = await parseXformStream(metadataXform, part.stream);
      if (metadataResult) {
        model.metadata = metadataResult;
      }
      return true;
    }
    case "xl/persons/person.xml": {
      // Office 365 threaded-comment person directory. Parsed here so
      // reconcile can attach the list to the workbook. Silently
      // ignored when malformed — threaded comments degrade to
      // "unknown author" rather than breaking the whole load.
      const data = await part.bytes();
      const raw = new TextDecoder().decode(data);
      model.persons = parsePersonList(raw);
      return true;
    }
    default: {
      // Catch threaded-comment per-sheet parts (the path contains a
      // variable sheet index so they can't be matched in the switch).
      const threadedMatch = /^xl\/threadedComments\/threadedComment(\d+)\.xml$/.exec(entryName);
      if (threadedMatch) {
        const sheetIndex = parseInt(threadedMatch[1], 10);
        const data = await part.bytes();
        const raw = new TextDecoder().decode(data);
        model.threadedCommentsByIndex ??= {} as Record<
          number,
          Array<{ ref: string; comment: unknown }>
        >;
        model.threadedCommentsByIndex[sheetIndex] = parseThreadedComments(raw);
        return true;
      }
      // Raw-passthrough capture for slicers and timelines — two
      // coordinated Office dashboard features documonster does not
      // structurally model but must not destroy on round-trip.
      // Each family has two part types (the control itself + its
      // cache); both are captured into maps on the workbook model
      // so the writer can emit them verbatim later.
      if (/^xl\/slicers\/slicer\d+\.xml$/.test(entryName)) {
        model.slicerParts ??= {} as Record<string, Uint8Array>;
        model.slicerParts[entryName] = await part.bytes();
        return true;
      }
      if (/^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(entryName)) {
        model.slicerCacheParts ??= {} as Record<string, Uint8Array>;
        model.slicerCacheParts[entryName] = await part.bytes();
        return true;
      }
      if (/^xl\/timelines\/timeline\d+\.xml$/.test(entryName)) {
        model.timelineParts ??= {} as Record<string, Uint8Array>;
        model.timelineParts[entryName] = await part.bytes();
        return true;
      }
      if (/^xl\/timelineCaches\/timelineCache\d+\.xml$/.test(entryName)) {
        model.timelineCacheParts ??= {} as Record<string, Uint8Array>;
        model.timelineCacheParts[entryName] = await part.bytes();
        return true;
      }
      return false;
    }
  }
}

/**
 * Refuse a binary workbook that reached the XML reader.
 *
 * A `.xlsb` read as XLSX produces an *empty* workbook rather than an error: none of the parts this
 * reader knows are present, so nothing loads and nothing complains. `read()` guards against it by
 * sniffing the bytes before choosing a loader, but that guard is bypassed whenever the format was not
 * inferred from them — `readFile` on a misnamed `.xlsx`, or any caller passing `format: "xlsx"`
 * explicitly. Both loaders below therefore ask this, because they are the two places that have seen
 * the part names.
 */
function assertNotBinaryWorkbook(sawBinaryWorkbook: boolean, sawXmlWorkbook: boolean): void {
  if (!sawBinaryWorkbook || sawXmlWorkbook) {
    return;
  }
  throw new ExcelFileError(
    "<input>",
    "read",
    `the package contains xl/workbook.bin and no ${OOXML_PATHS.xlWorkbook}, so it is XLSB and not ` +
      `XLSX. Reading it here would return an empty workbook rather than fail. Omit the format to ` +
      `detect it, or pass format: "xlsb".`
  );
}

/**
 * Load a workbook from an async stream of ZIP entries.
 *
 * This is the foundation for TRUE streaming reads on platforms that have a
 * streaming ZIP parser (e.g. Node.js `modules/archive` Parse).
 */
export async function readXlsxEntriesInto(
  workbook: Workbook,
  entries: AsyncIterable<ZipEntryLike>,
  options?: XlsxOptions
): Promise<Workbook> {
  const model: any = createEmptyModel();
  let sawBinaryWorkbook = false;
  let sawXmlWorkbook = false;

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
    sawBinaryWorkbook ||= entryName === "xl/workbook.bin";
    sawXmlWorkbook ||= entryName === OOXML_PATHS.xlWorkbook;
    try {
      if (!(await processEntry(streamedPart(entry.stream), model, entryName, options))) {
        // Important for true streaming parsers: always consume unknown entries
        await drainEntry();
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

  assertNotBinaryWorkbook(sawBinaryWorkbook, sawXmlWorkbook);
  await reconcile(model, options);
  loadWorkbookModel(workbook, model);
  return workbook;
}

/**
 * Parse a `.rels` part and register its edges for opaque reachability.
 *
 * `source` is required rather than optional on purpose. Whether a preserved
 * part is still reachable depends on every relationship in the package, not
 * just the ones this writer regenerates, so a `.rels` file parsed without being
 * registered here silently makes an unreachable part look like a part that
 * never had an inbound edge — which is the one case that must be emitted. Making
 * the parameter mandatory means a new `.rels` handler cannot omit it.
 *
 * @param source Part whose relationships these are; `""` for the package root.
 */
async function parseRelsPart(part: XlsxPart, model: any, source: string): Promise<any> {
  const xform = new RelationshipsXform();
  const relationships = await parseXformStream(xform, part.stream);
  recordOpaqueRelationships(model, source, relationships);
  return relationships;
}

function parseWorkbookPart(part: XlsxPart): Promise<any> {
  const xform = new WorkbookXform();
  return parseXformStream(xform, part.stream);
}

async function processWorksheetEntry(
  part: XlsxPart,
  model: any,
  sheetNo: number,
  options: XlsxOptions | undefined,
  path: string
): Promise<void> {
  const xform = new WorkSheetXform(options);
  const worksheet = await parseXformStream(xform, part.stream);
  if (!worksheet) {
    throw new XlsxParseError(path, "Failed to parse worksheet");
  }
  worksheet.sheetNo = sheetNo;
  model.worksheetHash[path] = worksheet;
  model.worksheets.push(worksheet);
}

async function processChartsheetEntry(part: XlsxPart, model: any, sheetNo: number): Promise<void> {
  const { ChartsheetXform } = await import("@excel/xlsx/xform/sheet/chartsheet-xform");
  const xform = new ChartsheetXform();
  const chartsheet = await parseXformStream(xform, part.stream);
  if (chartsheet) {
    chartsheet.sheetNo = sheetNo;
    model.chartsheets[sheetNo] = chartsheet;
  }
}

async function processCommentEntry(part: XlsxPart, model: any, zipPath: string): Promise<void> {
  const { CommentsXform } = await import("@excel/xlsx/xform/comment/comments-xform");
  const xform = new CommentsXform();
  const comments = await parseXformStream(xform, part.stream);
  // Key by absolute zip path so reconcile can match any rel target layout.
  model.comments[zipPath] = comments;
}

async function processTableEntry(part: XlsxPart, model: any, zipPath: string): Promise<void> {
  const { TableXform } = await import("@excel/xlsx/xform/table/table-xform");
  const xform = new TableXform();
  const table = await parseXformStream(xform, part.stream);
  // Key by absolute zip path so reconcile can match any rel target layout.
  model.tables[zipPath] = table;
}

async function processWorksheetRelsEntry(
  part: XlsxPart,
  model: any,
  sheetNo: number
): Promise<void> {
  // A sheet is where Excel looks for printer settings, so a relationship from
  // here may be the only thing that reaches a preserved part.
  model.worksheetRels[sheetNo] = await parseRelsPart(part, model, worksheetPath(sheetNo));
}

async function processMediaEntry(part: XlsxPart, model: any, filename: string): Promise<void> {
  const lastDot = filename.lastIndexOf(".");
  if (lastDot >= 1) {
    const extension = filename.substr(lastDot + 1);
    const name = filename.substr(0, lastDot);
    const buffer = await part.bytes();
    model.mediaIndex[filename] = model.media.length;
    model.mediaIndex[name] = model.media.length;
    model.media.push({ type: "image", name, extension, buffer });
  }
}

/**
 * Process a drawing XML entry: parse it, and keep its exact bytes for chart user-shape drawings.
 */
async function processDrawingEntry(part: XlsxPart, model: any, name: string): Promise<void> {
  const data = await part.bytes();

  // Parse the drawing for normal processing (images, etc.)
  const { DrawingXform } = await import("@excel/xlsx/xform/drawing/drawing-xform");
  const xform = new DrawingXform();
  const xmlString = bufferToString(data);
  const drawing = await parseXformStream(xform, [xmlString]);
  model.drawings[name] = drawing;
  // Also stash the original bytes — chart user-shape drawings use a
  // distinct schema (`c:relSizeAnchor` / `c:userShapes` instead of
  // `xdr:twoCellAnchor`) and are post-reconciled onto their owning
  // ChartEntry so the bytes can be written back verbatim. Regular
  // worksheet drawings don't read this map.
  if (!model.drawingRaw) {
    model.drawingRaw = {} as Record<string, Uint8Array>;
  }
  (model.drawingRaw as Record<string, Uint8Array>)[name] = data;
}

/**
 * Stash raw bytes of a chart-overlay drawing part. `c:userShapes`
 * parts live under `xl/drawings/chartUserShape{N}.xml` in files we
 * write ourselves and can use arbitrary names in foreign files (the
 * rel target is the only authoritative reference). The bytes are
 * keyed by the stem so `reconcileChartUserShapes` can match them
 * against each chart's `ChartUserShapes` rel Target.
 */
async function processChartUserShapesEntry(
  part: XlsxPart,
  model: any,
  name: string
): Promise<void> {
  const data = await part.bytes();
  if (!model.drawingRaw) {
    model.drawingRaw = {} as Record<string, Uint8Array>;
  }
  (model.drawingRaw as Record<string, Uint8Array>)[name] = data;
}

async function processDrawingRelsEntry(part: XlsxPart, model: any, name: string): Promise<void> {
  model.drawingRels[name] = await parseRelsPart(part, model, drawingPath(name));
}

async function processVmlDrawingEntry(part: XlsxPart, model: any, zipPath: string): Promise<void> {
  const { VmlDrawingXform } = await import("@excel/xlsx/xform/drawing/vml-drawing-xform");
  const xform = new VmlDrawingXform();
  const vmlDrawing = await parseXformStream(xform, part.stream);
  // Key by absolute zip path so reconcile can match any rel target layout.
  model.vmlDrawings[zipPath] = vmlDrawing;
}

async function processVmlDrawingHFEntry(part: XlsxPart, model: any, _name: string): Promise<void> {
  const { VmlDrawingXform } = await import("@excel/xlsx/xform/drawing/vml-drawing-xform");
  const xform = new VmlDrawingXform();
  const vmlDrawing = await parseXformStream(xform, part.stream);
  // Store every positioned header/footer image shape (LH/CH/RH/LF/CF/RF)
  // for reconciliation against the VML part's own image relationships.
  const headerImages =
    vmlDrawing?.headerImages ?? (vmlDrawing?.headerImage ? [vmlDrawing.headerImage] : undefined);
  if (headerImages?.length) {
    model.vmlDrawingHF ??= {};
    model.vmlDrawingHF[_name] = headerImages;
  }
}

async function processVmlDrawingHFRelsEntry(
  part: XlsxPart,
  model: any,
  name: string
): Promise<void> {
  model.vmlDrawingHFRels ??= {};
  model.vmlDrawingHFRels[name] = await parseRelsPart(part, model, `xl/drawings/${name}.vml`);
}

async function processThemeEntry(part: XlsxPart, model: any, name: string): Promise<void> {
  model.themes[name] = bufferToString(await part.bytes());
}

async function processPivotTableEntry(part: XlsxPart, model: any, name: string): Promise<void> {
  const { PivotTableXform } = await import("@excel/xlsx/xform/pivot-table/pivot-table-xform");
  const xform = new PivotTableXform();
  const pivotTable = await parseXformStream(xform, part.stream);
  if (pivotTable) {
    model.pivotTables[name] = pivotTable;
  }
}

async function processPivotTableRelsEntry(part: XlsxPart, model: any, name: string): Promise<void> {
  model.pivotTableRels[name] = await parseRelsPart(part, model, `xl/pivotTables/${name}.xml`);
}

async function processPivotCacheDefinitionEntry(
  part: XlsxPart,
  model: any,
  name: string
): Promise<void> {
  const { PivotCacheDefinitionXform } =
    await import("@excel/xlsx/xform/pivot-table/pivot-cache-definition-xform");
  const xform = new PivotCacheDefinitionXform();
  const cacheDefinition = await parseXformStream(xform, part.stream);
  if (cacheDefinition) {
    model.pivotCacheDefinitions[name] = cacheDefinition;
  }
}

async function processPivotCacheRecordsEntry(
  part: XlsxPart,
  model: any,
  name: string
): Promise<void> {
  const { PivotCacheRecordsXform } =
    await import("@excel/xlsx/xform/pivot-table/pivot-cache-records-xform");
  const xform = new PivotCacheRecordsXform();
  const cacheRecords = await parseXformStream(xform, part.stream);
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
async function processExternalLinkEntry(part: XlsxPart, model: any, index: number): Promise<void> {
  const xform = new ExternalLinkXform();
  const parsed = await parseXformStream(xform, part.stream);
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
async function processExternalLinkRelsEntry(
  part: XlsxPart,
  model: any,
  index: number
): Promise<void> {
  const relationships = await parseRelsPart(
    part,
    model,
    `xl/externalLinks/externalLink${index}.xml`
  );
  model.externalLinkRelsByIndex[index] = relationships ?? [];
}

async function processChartEntry(part: XlsxPart, model: any, chartNumber: number): Promise<void> {
  const data = await part.bytes();

  // Parse into model for high-level API access
  const xform = new ChartSpaceXform();
  const xmlString = bufferToString(data);
  const chart = await parseXformStream(xform, [xmlString]);
  if (chart) {
    model.chartEntries[chartNumber] = {
      chartNumber,
      model: chart,
      rawData: data,
      modelSnapshot: snapshotChartModel(chart)
    };
  }
}

async function processChartRelsEntry(
  part: XlsxPart,
  model: any,
  chartNumber: number
): Promise<void> {
  model.chartRels[chartNumber] = await parseRelsPart(
    part,
    model,
    `xl/charts/chart${chartNumber}.xml`
  );
}

async function readXlsxFilesInto(
  workbook: Workbook,
  zipData: Record<string, Uint8Array>,
  options?: XlsxReadOptions
): Promise<Workbook> {
  workbook.sourceFilePath = undefined;
  const model: any = createEmptyModel();

  const entries = Object.keys(zipData).map(name => ({
    name,
    dir: name.endsWith("/"),
    data: zipData[name]
  }));

  assertNotBinaryWorkbook(
    entries.some(entry => normalizeZipPath(entry.name) === "xl/workbook.bin"),
    entries.some(entry => normalizeZipPath(entry.name) === OOXML_PATHS.xlWorkbook)
  );

  for (const entry of entries) {
    if (!entry.dir) {
      const entryName = normalizeZipPath(entry.name);
      await processEntry(bufferedPart(entryName, entry.data), model, entryName, options);
    }
  }

  await reconcile(model, options);
  loadWorkbookModel(workbook, model);
  return workbook;
}

/**
 * Read one entry into the model — the step both readers share. Returns whether anything consumed it, so
 * the streaming reader knows to drain the rest.
 */
async function processEntry(
  part: XlsxPart,
  model: any,
  entryName: string,
  options?: XlsxOptions
): Promise<boolean> {
  return (
    (await processKnownEntry(part, model, entryName, options)) ||
    (await processDefaultEntry(part, model, entryName))
  );
}

/**
 * Process default entries (drawings, comments, tables, etc.)
 */
async function processDefaultEntry(
  part: XlsxPart,
  model: any,
  entryName: string
): Promise<boolean> {
  const sheetNo = getWorksheetNoFromWorksheetRelsPath(entryName);
  if (sheetNo !== undefined) {
    await processWorksheetRelsEntry(part, model, sheetNo);
    return true;
  }

  const chartsheetRelsNo = getChartsheetNoFromRelsPath(entryName);
  if (chartsheetRelsNo !== undefined) {
    model.chartsheetRels[chartsheetRelsNo] = await parseRelsPart(
      part,
      model,
      chartsheetPath(chartsheetRelsNo)
    );
    return true;
  }

  const mediaFilename = getMediaFilenameFromPath(entryName);
  if (mediaFilename) {
    await processMediaEntry(part, model, mediaFilename);
    return true;
  }

  const drawingName = getDrawingNameFromPath(entryName);
  if (drawingName) {
    await processDrawingEntry(part, model, drawingName);
    return true;
  }

  const chartUserShapesName = getChartUserShapesNameFromPath(entryName);
  if (chartUserShapesName) {
    await processChartUserShapesEntry(part, model, chartUserShapesName);
    return true;
  }

  const drawingRelsName = getDrawingNameFromRelsPath(entryName);
  if (drawingRelsName) {
    await processDrawingRelsEntry(part, model, drawingRelsName);
    return true;
  }

  const vmlDrawingName = getVmlDrawingNameFromPath(entryName);
  if (vmlDrawingName) {
    await processVmlDrawingEntry(part, model, entryName);
    return true;
  }

  // VML header/footer drawings (watermark in header mode).
  // Parse to extract header image info for round-trip preservation.
  const vmlHFName = getVmlDrawingHFNameFromPath(entryName);
  if (vmlHFName) {
    await processVmlDrawingHFEntry(part, model, vmlHFName);
    return true;
  }

  // The VML part's own rels map each shape's `o:relid` to a media target.
  const vmlHFRelsName = getVmlDrawingHFNameFromRelsPath(entryName);
  if (vmlHFRelsName) {
    await processVmlDrawingHFRelsEntry(part, model, vmlHFRelsName);
    return true;
  }

  if (isCommentsPath(entryName)) {
    await processCommentEntry(part, model, entryName);
    return true;
  }

  const tableName = getTableNameFromPath(entryName);
  if (tableName) {
    await processTableEntry(part, model, entryName);
    return true;
  }

  const themeName = getThemeNameFromPath(entryName);
  if (themeName) {
    await processThemeEntry(part, model, themeName);
    return true;
  }

  // Pivot table files
  const pivotTableName = getPivotTableNameFromPath(entryName);
  if (pivotTableName) {
    await processPivotTableEntry(part, model, pivotTableName);
    return true;
  }

  const pivotTableRelsName = getPivotTableNameFromRelsPath(entryName);
  if (pivotTableRelsName) {
    await processPivotTableRelsEntry(part, model, pivotTableRelsName);
    return true;
  }

  // Pivot cache files
  const pivotCacheDefinitionName = getPivotCacheDefinitionNameFromPath(entryName);
  if (pivotCacheDefinitionName) {
    await processPivotCacheDefinitionEntry(part, model, pivotCacheDefinitionName);
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
    await processPivotCacheRecordsEntry(part, model, pivotCacheRecordsName);
    return true;
  }

  // External workbook links: xl/externalLinks/externalLinkN.xml and its
  // sibling _rels file. Both parts are required to reconstruct the
  // ExternalLinkModel (the .xml carries sheet names + cached values; the
  // .rels carries the target path and TargetMode).
  const externalLinkIndex = getExternalLinkIndexFromPath(entryName);
  if (externalLinkIndex !== undefined) {
    await processExternalLinkEntry(part, model, externalLinkIndex);
    return true;
  }

  const externalLinkRelsIndex = getExternalLinkIndexFromRelsPath(entryName);
  if (externalLinkRelsIndex !== undefined) {
    await processExternalLinkRelsEntry(part, model, externalLinkRelsIndex);
    return true;
  }

  // Chart files — parse natively before the passthrough catch-all
  const chartNumber = getChartNumberFromPath(entryName);
  if (chartNumber !== undefined) {
    await processChartEntry(part, model, chartNumber);
    return true;
  }

  const chartRelsNumber = getChartNumberFromRelsPath(entryName);
  if (chartRelsNumber !== undefined) {
    await processChartRelsEntry(part, model, chartRelsNumber);
    return true;
  }

  const chartStyleNumber = getChartStyleNumberFromPath(entryName);
  if (chartStyleNumber !== undefined) {
    model.chartStyles[chartStyleNumber] = await part.bytes();
    return true;
  }

  const chartColorsNumber = getChartColorsNumberFromPath(entryName);
  if (chartColorsNumber !== undefined) {
    model.chartColors[chartColorsNumber] = await part.bytes();
    return true;
  }

  const chartExStyleNumber = getChartExStyleNumberFromPath(entryName);
  if (chartExStyleNumber !== undefined) {
    model.chartExStyles[chartExStyleNumber] = await part.bytes();
    return true;
  }

  const chartExColorsNumber = getChartExColorsNumberFromPath(entryName);
  if (chartExColorsNumber !== undefined) {
    model.chartExColors[chartExColorsNumber] = await part.bytes();
    return true;
  }

  // ChartEx files (Office 2016+ extended charts) — raw bytes plus best-effort structured model
  const chartExNumber = getChartExNumberFromPath(entryName);
  if (chartExNumber !== undefined) {
    const data = await part.bytes();
    const rawXml = bufferToString(data);
    model.chartExEntries[chartExNumber] = data;
    try {
      const parsed = parseChartEx(rawXml);
      model.chartExStructuredEntries[chartExNumber] = {
        chartExNumber,
        model: parsed,
        rawData: data,
        modelSnapshot: snapshotChartModel(parsed)
      };
    } catch {
      // Keep legacy-safe passthrough if a third-party chartEx part is not parseable.
    }
    return true;
  }

  const chartExRelsNumber = getChartExNumberFromRelsPath(entryName);
  if (chartExRelsNumber !== undefined) {
    model.chartExRels[chartExRelsNumber] = await parseRelsPart(
      part,
      model,
      `xl/charts/chartEx${chartExRelsNumber}.xml`
    );
    return true;
  }

  // Raw-passthrough catch-all for Office 2010+ slicer/timeline
  // dashboard controls and their associated rels. documonster does not
  // model these structurally yet; capturing the bytes here prevents
  // silent data loss on round-trip when a dashboard workbook comes
  // through. Same idea covers the two-level rels files produced by
  // Excel (the `_rels` subfolder sits next to each part).
  if (
    /^xl\/slicers\/slicer\d+\.xml$/.test(entryName) ||
    /^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(entryName) ||
    /^xl\/timelines\/timeline\d+\.xml$/.test(entryName) ||
    /^xl\/timelineCaches\/timelineCache\d+\.xml$/.test(entryName) ||
    /^xl\/slicers\/_rels\/slicer\d+\.xml\.rels$/.test(entryName) ||
    /^xl\/slicerCaches\/_rels\/slicerCache\d+\.xml\.rels$/.test(entryName) ||
    /^xl\/timelines\/_rels\/timeline\d+\.xml\.rels$/.test(entryName) ||
    /^xl\/timelineCaches\/_rels\/timelineCache\d+\.xml\.rels$/.test(entryName)
  ) {
    const targetMap =
      entryName.startsWith("xl/slicers/") && !entryName.includes("/_rels/")
        ? (model.slicerParts ??= {} as Record<string, Uint8Array>)
        : entryName.startsWith("xl/slicerCaches/") && !entryName.includes("/_rels/")
          ? (model.slicerCacheParts ??= {} as Record<string, Uint8Array>)
          : entryName.startsWith("xl/timelines/") && !entryName.includes("/_rels/")
            ? (model.timelineParts ??= {} as Record<string, Uint8Array>)
            : entryName.startsWith("xl/timelineCaches/") && !entryName.includes("/_rels/")
              ? (model.timelineCacheParts ??= {} as Record<string, Uint8Array>)
              : entryName.startsWith("xl/slicers/_rels/")
                ? (model.slicerParts ??= {} as Record<string, Uint8Array>)
                : entryName.startsWith("xl/slicerCaches/_rels/")
                  ? (model.slicerCacheParts ??= {} as Record<string, Uint8Array>)
                  : entryName.startsWith("xl/timelines/_rels/")
                    ? (model.timelineParts ??= {} as Record<string, Uint8Array>)
                    : (model.timelineCacheParts ??= {} as Record<string, Uint8Array>);
    targetMap[entryName] = await part.bytes();
    return true;
  }

  // Nothing above recognised this entry. Keep it rather than drain it: an
  // unrecognised part is far more likely to be someone's data (a VBA project,
  // custom properties, a data connection) than something this writer is
  // entitled to delete. `collectOpaqueParts` applies the drop policy later,
  // once the whole package has been seen.
  if (isRelationshipsPart(entryName)) {
    const owner = ownerOfRelationshipsPart(entryName);
    if (owner !== undefined) {
      await parseRelsPart(part, model, owner);
    }
    // Handled as relationships, not as bytes: a preserved part's `.rels` is
    // re-emitted from its parsed `relationships` so that a relationship
    // pointing at something we dropped does not survive as a dangling
    // reference. Returning true stops the caller draining it as unknown.
    return true;
  }

  model.opaqueUnknownEntries.set(entryName, await part.bytes());
  return true;
}

/**
 * Record a parsed `.rels` file against the part that declares it.
 *
 * Both the modelled rels (root, workbook) and an opaque part's own rels land
 * here, because resolving "what pointed at this preserved part" needs every
 * relationship in the package, not just the unrecognised ones.
 */
function recordOpaqueRelationships(
  model: any,
  source: string,
  rels: readonly RelationshipModel[] | undefined
): void {
  if (!rels || rels.length === 0) {
    return;
  }
  // `Id`/`Type`/`Target` are optional on `RelationshipModel` because it is the
  // raw attribute bag the SAX parser produced, so a malformed `.rels` can be
  // missing any of them. Such an entry is skipped rather than defaulted: an
  // edge with no target cannot make anything reachable, and inventing one would
  // put a broken relationship into the output.
  const recorded: OpaqueRelationship[] = [];
  for (const rel of rels) {
    if (rel.Id === undefined || rel.Type === undefined || rel.Target === undefined) {
      continue;
    }
    recorded.push({
      id: rel.Id,
      type: rel.Type,
      target: rel.Target,
      targetMode: rel.TargetMode
    });
  }
  if (recorded.length > 0) {
    model.opaqueRelationshipsBySource.set(source, recorded);
  }
}
