/**
 * DOCX Module - Streaming Writer
 *
 * A DOCX generator that serializes body content incrementally and pushes it
 * through a streaming compression pipeline. Uses the same streaming ZIP
 * infrastructure as the Excel module:
 *
 * - `Zip` (StreamingZip) — streams ZIP entries to output
 * - `ZipDeflate` — per-entry deflate compression
 * - `StreamBuf` — event-driven pipe from XML to ZIP
 * - `StringBuf` — efficient XML string builder with buffer reuse
 *
 * Data flow with sink (true end-to-end streaming):
 * ```
 * add(paragraph) → XML → StreamBuf → ZipDeflate → Zip
 *                                                     ↓ (per-chunk callback)
 *                                            await SinkAdapter.write(chunk)
 *                                                     ↓
 *                                           user-supplied WritableStream /
 *                                           Node Writable / duck-typed sink
 * ```
 *
 * Memory profile:
 *   - Body model is never retained: each element is serialised and
 *     compressed as it arrives, so peak per-element memory is
 *     O(largest_single_element).
 *   - When `options.sink` is provided, compressed bytes are pushed
 *     into the sink as soon as they are produced (with backpressure
 *     awaited via {@link SinkAdapter}). Total writer-side memory then
 *     stays O(largest_part) regardless of final DOCX size.
 *   - When `options.sink` is omitted, compressed bytes accumulate in
 *     `_outputChunks` and `toBuffer()` returns the assembled
 *     `Uint8Array`. Total memory is O(compressed_docx_size).
 */

import { Zip, ZipDeflate } from "@archive/zip/stream";
import type { AnySink } from "@stream/core/sink-adapter";
import { SinkAdapter } from "@stream/core/sink-adapter";
import { withStatedDefaults } from "@word/builder/document-finalize";
import {
  ContentType,
  RelType,
  PartPath,
  DOCUMENT_NAMESPACES,
  STD_DOC_ATTRIBUTES
} from "@word/constants";
import {
  checkAppProperties,
  checkArray,
  checkBlock,
  checkCoreProperties,
  checkDocumentSettings,
  checkInteger,
  checkNodeArray,
  checkOneOf,
  checkPackageOptions,
  checkParagraphProperties,
  checkRecord,
  checkRunProperties,
  checkSectionProperties,
  checkString,
  checkStyleDef,
  checkWatermark,
  invalidArgument
} from "@word/core/argument-checks";
import { referencedStyleIds, withReferencedBuiltInStyles } from "@word/core/builtin-styles";
import { collectImageKeys, indexImagesByKey } from "@word/core/image-keys";
import { assertWritableImage, imagePartOverride } from "@word/core/image-media";
import { sanitizeUrl, utf8Encoder } from "@word/core/internal-utils";
import { atNode, checkWritableBlock, within } from "@word/core/node-path";
import { getFileExt, getPartRelsPath } from "@word/core/opc-paths";
import { isRangeMarker, isRun } from "@word/core/text-utils";
import type { DocxVisitor } from "@word/core/walker";
import { walkBlocks } from "@word/core/walker";
import { planWatermark } from "@word/core/watermark-sections";
import type { WatermarkPlan } from "@word/core/watermark-sections";
import { DocxWriteError } from "@word/errors";
import type { DocxSecurityPolicy } from "@word/security/policy";
import type {
  AbstractNumbering,
  AppProperties,
  BodyContent,
  ChartContent,
  ChartExContent,
  CommentDef,
  CoreProperties,
  CustomProperty,
  CustomXmlPart,
  DocDefaults,
  DocumentBackground,
  DocumentSettings,
  DocumentTheme,
  EmbeddedFont,
  EndnoteDef,
  FontDef,
  FooterDef,
  FootnoteDef,
  HeaderDef,
  HeaderFooterRef,
  ImageDef,
  NumberingInstance,
  OpaquePart,
  Paragraph,
  RangeMarker,
  Run,
  SectionProperties,
  StyleDef,
  Watermark
} from "@word/types";
import { renderChartPart } from "@word/writer/chart-writer";
import { renderComments, renderCommentsExtended } from "@word/writer/comment-writer";
import { buildCommonAuxiliaryParts } from "@word/writer/common-parts";
import {
  createContentTypes,
  addContentTypeDefault,
  addContentTypeOverride,
  addImageContentTypeDefaults,
  renderContentTypes,
  typeUntypedParts
} from "@word/writer/content-types";
import { renderBodyContent } from "@word/writer/document-writer";
import { uniqueSanitizedName } from "@word/writer/docx-packager";
import {
  renderHeader,
  renderFooter,
  renderWatermarkHeader
} from "@word/writer/header-footer-writer";
import type { HeaderWatermark } from "@word/writer/header-footer-writer";
import { assertUniqueReferencedIds } from "@word/writer/id-uniqueness";
import { registerPartImages } from "@word/writer/part-images";
import { renderRangeMarkers } from "@word/writer/range-marker-writer";
import {
  collectChartsFromHeaderFooter,
  collectHyperlinksFromHeaderFooter,
  collectHyperlinksFromNotes,
  collectPreservedLinks,
  preservedLinkRegistrar
} from "@word/writer/reference-scanners";
import type { RelationshipsState } from "@word/writer/relationships";
import {
  createRelationships,
  addRelationship,
  addRelationshipWithId,
  getRelationshipCount,
  renderRelationships
} from "@word/writer/relationships";
import type { RevisionIdMap, DocxRenderContext } from "@word/writer/render-context";
import {
  closeStoryPart,
  createRenderContext,
  enterRevisionIds,
  exitRevisionIds,
  leaveRevisionIds,
  renderHelpersFor,
  withRevisionScope
} from "@word/writer/render-context";
import { renderSectionProperties } from "@word/writer/section-writer";
import { StreamBuf } from "@word/writer/stream-buf";
import { StringBuf } from "@word/writer/string-buf";
import { xmlEncodeAttr, xmlValue } from "@xml/encode";
import { XmlWriter } from "@xml/writer";

// Per-instance StringBuf is created in the constructor (see _xmlBuffer field below).

/** How much `document.xml` is collected before it is handed to the compressor — see `_write`. */
const XML_FLUSH_BYTES = 65536;
// Previously this was a module-level singleton which caused data races with concurrent instances.

const EMPTY_U8 = new Uint8Array(0);

/** Stands for the watermark-only header part while the sections are planned. */
const WATERMARK_PART = "\u0000watermark";

/** The streaming writer stores every image under its own file name. */
function streamingImageTarget(img: ImageDef): string {
  return `media/${img.fileName}`;
}

// =============================================================================
// Types
// =============================================================================

/** Options for the streaming DOCX writer. */
export interface StreamingDocxOptions {
  /** Compression level (0-9). Default: 6. */
  readonly compressionLevel?: number;
  /**
   * Timestamp stamped on every ZIP entry. Defaults to `coreProperties.modified`,
   * then `coreProperties.created`, then the current time.
   */
  readonly modTime?: Date;
  /** Progress callback interval: report after every N elements. Default: 1000. */
  readonly chunkSize?: number;
  /** Section properties for the final section. */
  readonly sectionProperties?: SectionProperties;
  /**
   * Range markers after the last element, before the final section properties
   * (see `DocxDocument.bodyRangeMarkersAfter`).
   */
  readonly bodyRangeMarkersAfter?: readonly RangeMarker[];
  /** Document styles. */
  readonly styles?: readonly StyleDef[];
  /** Document defaults. */
  readonly docDefaults?: DocDefaults;
  /** Abstract numbering definitions. */
  readonly abstractNumberings?: readonly AbstractNumbering[];
  /** Numbering instances. */
  readonly numberingInstances?: readonly NumberingInstance[];
  /** Headers. */
  readonly headers?: ReadonlyMap<string, HeaderDef>;
  /** Footers. */
  readonly footers?: ReadonlyMap<string, FooterDef>;
  /** Footnotes. */
  readonly footnotes?: readonly FootnoteDef[];
  /** Endnotes. */
  readonly endnotes?: readonly EndnoteDef[];
  /** Images. */
  readonly images?: readonly ImageDef[];
  /** Fonts. */
  readonly fonts?: readonly FontDef[];
  /** Document settings. */
  readonly settings?: DocumentSettings;
  /** Core properties. */
  readonly coreProperties?: CoreProperties;
  /** App properties. */
  readonly appProperties?: AppProperties;
  /** Comments. */
  readonly comments?: readonly CommentDef[];
  /** Background. */
  readonly background?: DocumentBackground;
  /** Custom properties. */
  readonly customProperties?: readonly CustomProperty[];
  /** Watermark. */
  readonly watermark?: Watermark;
  /** Theme. */
  readonly theme?: DocumentTheme;
  /** Custom XML parts (for SDT data binding). */
  readonly customXmlParts?: readonly CustomXmlPart[];
  /** Embedded font binaries (stored in word/fonts/). */
  readonly embeddedFonts?: readonly EmbeddedFont[];
  /** Opaque (unrecognized) parts preserved for round-trip fidelity. */
  readonly opaqueParts?: readonly OpaquePart[];
  /**
   * How to handle image references whose binary is not in `images`.
   *
   * - `"throw"` (default): throw `DocxWriteError` from `add*` so the caller
   *   notices the broken reference immediately.
   * - `"warn"`: emit a `console.warn` and skip the rId. The output will be
   *   missing this image's relationship — useful only for tooling that
   *   knows it's intentionally producing a partial document.
   *
   * The previous behaviour (silent skip) is gone because it generated
   * invalid DOCX files.
   */
  readonly missingImagePolicy?: "throw" | "warn";
  /**
   * Security policy. Currently used to surface `rawXmlPolicy` to the renderers
   * (preserve / strip / reject) so opaque rawXml fields behave consistently
   * with the buffered `packageDocx` writer.
   */
  readonly securityPolicy?: DocxSecurityPolicy;

  /**
   * Output sink. When provided, compressed bytes flow through it as soon
   * as the underlying ZIP pipeline produces them, with backpressure
   * awaited via {@link SinkAdapter}. Total writer-side memory then
   * stays O(largest_part) regardless of final DOCX size.
   *
   * Accepts:
   *  - Web `WritableStream<Uint8Array>` (browser, Deno, modern Node)
   *  - Node `Writable` (`fs.createWriteStream`, http response, …)
   *  - Any duck-typed object exposing `write(chunk)` + `end()` plus
   *    `once("drain"|"error"|"close"|"finish", …)` listeners
   *
   * When the sink is provided, {@link DocumentWriter.commit} ends the
   * sink once the document is complete, and
   * {@link DocumentWriter.toBuffer} is refused — the bytes are in the
   * sink, not the writer. Use
   * {@link DocumentWriter.addAsync} (instead of
   * {@link DocumentWriter.add}) when you want each `add` call to
   * await actual sink drain — that gives true end-to-end backpressure
   * for tight production loops.
   *
   * When omitted, behaviour is unchanged: compressed bytes accumulate
   * internally and `toBuffer()` returns the assembled `Uint8Array`.
   */
  readonly sink?: AnySink;
}

/** Progress callback for streaming writer. */
export type StreamingProgressCallback = (info: {
  /** Number of body elements written so far. */
  elementsWritten: number;
  /** Current phase: "body" | "finalizing". */
  phase: string;
}) => void;

// =============================================================================
// Streaming DOCX Writer
// =============================================================================

/**
 * A streamed element, checked before anything of it is written: a streamed
 * element cannot be taken back out of the part once it is half-written.
 */
function checkStreamedBlock(element: unknown, field: string, where: string): void {
  checkBlock(element, field, where);
  try {
    checkWritableBlock(element as BodyContent);
  } catch (error) {
    throw atNode(field, error);
  }
}

/** Every element of a batch, checked before any of them is written. */
function checkStreamedBlocks(elements: unknown, where: string): void {
  checkArray(elements, "elements", where);
  (elements as readonly unknown[]).forEach((el, i) =>
    checkStreamedBlock(el, `elements[${i}]`, where)
  );
}

/**
 * Streaming DOCX writer. Body elements are serialized to XML and compressed
 * into the ZIP pipeline as they arrive, so the body **model** is not retained
 * after each `add()`.
 *
 * When constructed with `options.sink`, compressed bytes are pushed into
 * the sink as soon as the ZIP layer produces them, with backpressure
 * awaited via {@link SinkAdapter}; in this mode peak memory is
 * O(largest_part), `commit()` ends the sink and `toBuffer()` is
 * refused (the bytes are already in the sink). Without a sink,
 * compressed bytes accumulate in `_outputChunks` and `toBuffer()`
 * returns the complete file (peak memory O(compressed_docx_size)).
 *
 * Every writer finishes with `commit()`; a writer that owns its output
 * additionally hands it over with `toBuffer()`, which commits first if
 * needed.
 *
 * Use {@link addAsync} (instead of {@link add}) when driving the sink
 * variant to obtain true end-to-end backpressure: each call awaits all
 * pending sink writes before resolving.
 *
 * @example
 * ```ts
 * const writer = new Stream.DocumentWriter({
 *   styles: [{ type: "paragraph", styleId: "Normal", name: "Normal" }]
 * });
 *
 * for (let i = 0; i < 100000; i++) {
 *   writer.addText(`Paragraph ${i}`);
 * }
 *
 * const buffer = await writer.toBuffer();
 * ```
 */
export class DocumentWriter {
  private readonly _options: StreamingDocxOptions;
  private _elementCount = 0;
  private _finalized = false;
  /** The running or settled {@link commit}; `undefined` until the writer is committed. */
  private _finished: Promise<void> | undefined;
  private _onProgress?: StreamingProgressCallback;

  // Per-instance XML buffer (avoids module-level singleton data race)
  private readonly _xmlBuffer = new StringBuf({ size: 65536 });

  // ZIP infrastructure
  private _zip!: InstanceType<typeof Zip>;
  /** Compressed-byte accumulator used when no `sink` is supplied. */
  private _outputChunks: Uint8Array[] = [];
  /** Sink-mode adapter (set when `options.sink` is provided). */
  private _sinkAdapter?: SinkAdapter;
  /**
   * Promise chain serialising every sink write. The `Zip` callback fires
   * synchronously, so we queue chunks via `.then(...)` and let
   * `addAsync` / `finalize` await the chain.
   */
  private _pendingDrain: Promise<void> = Promise.resolve();
  /** Completion of every ZipDeflate push registered so far. */
  /**
   * Compression pushes not yet settled. A set rather than a chain: chaining every push onto the last with
   * `Promise.all(...).then(...)` built two promises per element written — for a 200k-paragraph document most of the
   * writer's own time and a fifth of it in garbage collection — while a push that has already settled, which is every
   * push on the synchronous compression path, has nothing left to wait for.
   */
  private _pendingCompression = new Set<Promise<void>>();
  /**
   * Resolves when the `Zip` archive emits its terminal callback (`final`).
   * In the browser the deflate pipeline (CompressionStream) is asynchronous,
   * so `_zip.end()` returns before the trailing chunks and central directory
   * have been produced; `finalize()` awaits this before reading
   * `_outputChunks` / closing the sink. In Node the deflate is synchronous and
   * this resolves before `_zip.end()` returns, so the await is a no-op.
   */
  private _zipComplete!: Promise<void>;
  private _resolveZipComplete!: () => void;
  private _documentStream!: StreamBuf;
  private _documentZipFile!: InstanceType<typeof ZipDeflate>;
  private _headerWritten = false;
  /**
   * Whether the previously-written body element was a `<w:tbl>`. Tracked
   * so we can insert a separator `<w:p>` between adjacent tables — Word
   * rejects (and silently merges) two `<w:tbl>` blocks that share no
   * paragraph between them per ECMA-376 §17.13.5.34.
   */
  private _prevWasTable = false;
  /**
   * First error reported by the underlying ZIP stream (compression failure,
   * write-after-end, etc.). Stored synchronously by the `Zip` callback and
   * surfaced from `finalize()` so callers receive a rejection instead of an
   * indefinitely-pending promise.
   */
  private _streamError: Error | null = null;

  // Relationships and render context — built up as `add()` is called so that
  // every body element is serialized with the correct `r:id` for embedded
  // images, hyperlinks and charts. Without these, charts crash mid-write
  // ("Chart content was not registered with a relationship id") and
  // hyperlinks/images dangle.
  private _documentRels!: RelationshipsState;
  private _renderCtx!: DocxRenderContext;
  /** Charts encountered in body content; rendered to `word/charts/chartN.xml` at finalize time. */
  private readonly _bodyCharts: ChartContent[] = [];
  /** ChartEx items encountered in body content. */
  private readonly _bodyChartEx: ChartExContent[] = [];
  /**
   * Per-chart sequence numbers fixed at registration time. Both classes
   * use independent monotonic counters; the writer emits
   * `word/charts/chart{n}.xml` for the regular chart family and
   * `word/charts/chartEx{n}.xml` for the chartEx family.
   *
   * The previous scheme used `chartCount + chartExCount + 1` as the
   * sequence number for both classes, which made the rId path encoded in
   * documentRels disagree with the path used at finalize when the writer
   * iterated the two arrays separately. The result was relationships
   * pointing at non-existent chart parts.
   */
  private readonly _chartNum = new WeakMap<object, number>();
  private _nextChartSeq = 0;
  private _nextChartExSeq = 0;
  /** Hyperlink object identities already registered (to keep one rId per object). */
  private readonly _registeredHyperlinks = new WeakSet<object>();
  /** What the walker reports to while registering an element's references; one for every element. */
  private readonly _referenceVisitor: DocxVisitor = {
    enterParagraph: para => {
      this._registerParagraphReferences(para);
    },
    enterRun: run => {
      this._registerRunImages(run);
    },
    enterHyperlink: h => {
      this._registerHyperlink(h);
    }
  };
  /** Every image by each of its keys. */
  private readonly _imageIndex: ReadonlyMap<string, ImageDef>;
  /** Image rIds already registered to documentRels (avoid duplicates). */
  private readonly _registeredImageRIds = new Set<string>();
  /** header map key → newly allocated rId. Populated by `_allocateHeaderFooterRIds`. */
  private readonly _headerKeyToRid = new Map<string, string>();
  /** footer map key → newly allocated rId. */
  private readonly _footerKeyToRid = new Map<string, string>();
  /** rId allocated for the auto-generated watermark header (if any). */
  private _watermarkHeaderRid: string | undefined;
  /** The style ids the body written so far references. */
  private readonly _styleRefs = new Set<string>();
  /**
   * Where the watermark goes, planned section by section as the sections
   * stream past (see `planWatermark`). Its hosts are header map keys.
   */
  private _watermarkPlan: WatermarkPlan | undefined;
  /** Header and footer ids a section break inside the body names. */
  private readonly _innerSectionRIds = new Set<string>();
  /** The header types the final section names the watermark-only header for. */
  private _finalWatermarkTypes: readonly HeaderFooterRef["type"][] = [];

  /**
   * @throws {DocxDuplicateIdError} If two footnotes, endnotes, comments,
   * numbering definitions or builder images in `options` share an id.
   */
  constructor(options: StreamingDocxOptions = {}) {
    checkStreamingOptions(options, "new Stream.DocumentWriter");
    // Sanitize image/font file names up-front. They get embedded into
    // ZIP entry paths and into rels Target attributes; a hostile name
    // (e.g. `../../etc/passwd.png` from a round-tripped untrusted DOCX)
    // would otherwise produce a zipslip-shaped output. Mirrors what
    // `packageDocx` does in `shallowCopyDocForPackaging`.
    assertUniqueReferencedIds(options);
    const sanitized = withHeaderSwitches(sanitizeStreamingOptions(options));
    this._options = sanitized;
    this._imageIndex = indexImagesByKey(sanitized.images);
    this._initZip();
  }

  /** Set a progress callback. */
  onProgress(cb: StreamingProgressCallback): this {
    if (typeof cb !== "function") {
      invalidArgument("Stream.DocumentWriter#onProgress", "cb", "a function", cb);
    }
    this._onProgress = cb;
    return this;
  }

  /**
   * Add a single body element. The element is immediately serialized to XML
   * and pushed into the ZIP compression pipeline. After this call, the element
   * can be garbage collected — it is not retained.
   */
  add(element: BodyContent): this {
    checkStreamedBlock(element, "element", "Stream.DocumentWriter#add");
    return within("element", () => this._add(element));
  }

  /** Write an element {@link checkStreamedBlock} has checked. */
  private _add(element: BodyContent): this {
    if (this._finalized) {
      throw new DocxWriteError("Stream.DocumentWriter: cannot add elements after commit()");
    }

    // Sink-mode early failure: if a previous chunk already failed to
    // reach the sink, surface that immediately rather than letting the
    // caller keep streaming work that will be discarded. Buffered mode
    // keeps the legacy behaviour of deferring all error reporting to
    // `finalize()` because there is no live consumer that could be
    // disrupted by silent queueing.
    if (this._sinkAdapter && this._streamError) {
      throw new DocxWriteError(
        `Stream.DocumentWriter: sink already failed (${this._streamError.message})`,
        { cause: this._streamError }
      );
    }

    // Write document.xml header on first element
    if (!this._headerWritten) {
      this._writeDocumentHeader();
      this._headerWritten = true;
    }

    // Register any chart/hyperlink/image references this element introduces
    // BEFORE serializing it, so the per-element renderBodyContent call has a
    // populated DocxRenderContext (chart rIds, hyperlink rIds, image remap).
    this._registerElementReferences(element);

    // ECMA-376 §17.13.5.34: a `<w:tbl>` must be followed by a paragraph
    // (or section break) before the next `<w:tbl>` may begin. When the
    // caller streams two adjacent tables we synthesise an empty
    // separator paragraph between them so Word does not collapse them
    // into a single malformed tbl.
    if (element.type === "table" && this._prevWasTable) {
      this._writeSeparatorParagraph();
    }

    // Serialize this single element to XML and push to stream
    this._writeBodyElement(this._prepareSectionBreak(element));
    this._elementCount++;
    this._prevWasTable = element.type === "table";

    if (this._onProgress && this._elementCount % (this._options.chunkSize ?? 1000) === 0) {
      this._onProgress({ elementsWritten: this._elementCount, phase: "body" });
    }

    return this;
  }

  /** Add multiple body elements at once. Every one is checked before any is written. */
  addMany(elements: readonly BodyContent[]): this {
    checkStreamedBlocks(elements, "Stream.DocumentWriter#addMany");
    elements.forEach((el, i) => within(`elements[${i}]`, () => this._add(el)));
    return this;
  }

  /**
   * Async variant of {@link add}. After serialising the element, awaits
   * any pending writes to the configured `sink` so callers driving large
   * input get true end-to-end backpressure rather than letting the
   * sink-write queue grow unbounded inside the writer.
   *
   * Without `options.sink` this is equivalent to `add` (resolving
   * synchronously after element serialisation).
   *
   * Throws if the sink reports an error: previous queued writes whose
   * rejection was captured into `_streamError` surface here.
   */
  async addAsync(element: BodyContent): Promise<this> {
    checkStreamedBlock(element, "element", "Stream.DocumentWriter#addAsync");
    return this._addAsync(element);
  }

  private async _addAsync(element: BodyContent, field = "element"): Promise<this> {
    within(field, () => this._add(element));
    if (this._sinkAdapter) {
      this._flushXml();
      // ZipDeflate batches small inputs. A non-final empty push flushes that
      // batch into the compressor without ending document.xml.
      this._trackCompression(this._documentZipFile.push(EMPTY_U8));
      await this._awaitOutputBoundary("sink write failed");
    }
    return this;
  }

  /**
   * Async variant of {@link addMany}. Awaits `addAsync` for each element
   * so backpressure is honoured between every body element.
   */
  async addManyAsync(elements: readonly BodyContent[]): Promise<this> {
    checkStreamedBlocks(elements, "Stream.DocumentWriter#addManyAsync");
    for (let i = 0; i < elements.length; i++) {
      await this._addAsync(elements[i]!, `elements[${i}]`);
    }
    return this;
  }

  /** Add a paragraph with simple text content. */
  addText(content: string, properties?: Paragraph["properties"]): this {
    const where = "Stream.DocumentWriter#addText";
    checkString(content, "content", where);
    checkRecord(properties, "properties", where, { optional: true });
    checkParagraphProperties(properties, "properties", where);
    return this.add({
      type: "paragraph",
      children: [{ content: [{ type: "text", text: content }] }],
      properties
    } as Paragraph);
  }

  /** Get the count of body elements written so far. */
  get elementCount(): number {
    return this._elementCount;
  }

  /**
   * Finish the document.
   *
   * Writes the closing body, every auxiliary part and the ZIP central
   * directory. With `options.sink` it then drains every pending write and
   * ends the sink; without one the bytes stay in the writer for
   * {@link toBuffer}.
   *
   * @throws {DocxWriteError} If the writer was already committed, or the
   * ZIP pipeline or the sink failed (the sink is aborted in that case).
   */
  async commit(): Promise<void> {
    if (this._finalized) {
      throw new DocxWriteError("Stream.DocumentWriter: already committed");
    }
    this._finalized = true;
    this._finished = this._finish();
    await this._finished;
  }

  /**
   * The complete DOCX file. Commits first when {@link commit} has not been
   * called; afterwards it may be called again and resolves to the same bytes.
   *
   * @throws {DocxWriteError} If the writer was given `options.sink` — the
   * bytes went to the sink and the writer holds none — or committing failed.
   */
  async toBuffer(): Promise<Uint8Array> {
    if (this._sinkAdapter) {
      throw new DocxWriteError(
        "Stream.DocumentWriter#toBuffer: the document was written to options.sink; the writer holds no bytes"
      );
    }
    if (this._finished) {
      await this._finished;
    } else {
      await this.commit();
    }
    const bytes = this._assembleOutput();
    this._outputChunks = [bytes];
    return bytes;
  }

  /**
   * Finalize the document.
   *
   * - Without `options.sink`: returns the assembled `Uint8Array`
   *   containing the full DOCX file.
   * - With `options.sink`: drains any pending sink writes, calls
   *   `sink.end()`, and resolves to a zero-length `Uint8Array`.
   *
   * @deprecated Use {@link commit} to finish the document and
   * {@link toBuffer} to obtain its bytes. Removed in 1.0.
   */
  async finalize(): Promise<Uint8Array> {
    if (this._finalized) {
      throw new DocxWriteError("Stream.DocumentWriter: already finalized");
    }
    this._finalized = true;
    this._finished = this._finish();
    await this._finished;
    if (this._sinkAdapter) {
      return EMPTY_U8;
    }
    return this._assembleOutput();
  }

  /** Write the trailing parts and close the archive and, in sink mode, the sink. */
  private async _finish(): Promise<void> {
    if (this._onProgress) {
      this._onProgress({ elementsWritten: this._elementCount, phase: "finalizing" });
    }

    // If no elements were added, still write a minimal document
    if (!this._headerWritten) {
      this._writeDocumentHeader();
    }

    this._planFinalWatermark();

    // Write document.xml footer (close </w:body></w:document>)
    const previousRevisionIds = enterRevisionIds(this._renderCtx.revisionIds);
    try {
      this._writeDocumentFooter();
    } finally {
      exitRevisionIds(previousRevisionIds);
    }

    // End the document.xml stream → finalizes its ZIP entry
    await this._endStream(this._documentStream);

    // Add all auxiliary parts (styles, settings, etc.). They are rendered
    // synchronously before the method first awaits, which is where the scope
    // takes the writer's revision map out of effect again.
    await withRevisionScope(() =>
      this._addAuxiliaryParts(enterRevisionIds(this._renderCtx.revisionIds))
    );

    // Finalize the ZIP archive. Any compression errors during the trailing
    // central-directory write are reported via the `Zip` callback into
    // `_streamError`; surface them as a rejection.
    this._zip.end();

    // Wait for the archive to fully drain. On the browser deflate path the
    // trailing chunks + central directory are produced asynchronously after
    // `end()` returns; without this the buffered output would be assembled
    // from an empty / partial `_outputChunks`. No-op on Node's sync path.
    await this._zipComplete;

    // In sink mode the Zip callback queued every chunk onto _pendingDrain;
    // wait for that promise chain to complete before declaring the writer
    // finished. In buffered mode this is a no-op (resolved promise).
    await this._pendingDrain;

    if (this._streamError) {
      // The document is incomplete: give the sink up rather than leaving it
      // open (and a Web stream locked) behind a failed write.
      await this._sinkAdapter?.abort(this._streamError);
      throw new DocxWriteError(
        `Stream.DocumentWriter: ZIP finalization failed (${this._streamError.message})`,
        { cause: this._streamError }
      );
    }

    if (this._sinkAdapter) {
      // Close the sink so the consumer knows the byte stream is complete.
      // Errors raised during close (e.g. underlying file system) propagate.
      await this._sinkAdapter.end();
    }
  }

  /** Reset the writer for reuse. */
  /**
   * Reset the writer for reuse.
   *
   * Throws when the writer was constructed with an `options.sink`: a
   * sink can only be `end()`ed once, so reusing the same writer would
   * produce an undefined byte stream. Construct a new writer (with a
   * new sink) for each document instead.
   */
  reset(): this {
    if (this._sinkAdapter) {
      throw new DocxWriteError(
        "Stream.DocumentWriter: reset() is not supported in sink mode; create a new writer instance per document."
      );
    }
    this._elementCount = 0;
    this._finalized = false;
    this._finished = undefined;
    this._headerWritten = false;
    this._prevWasTable = false;
    this._outputChunks = [];
    this._bodyCharts.length = 0;
    this._bodyChartEx.length = 0;
    this._nextChartSeq = 0;
    this._nextChartExSeq = 0;
    this._registeredImageRIds.clear();
    this._watermarkHeaderRid = undefined;
    this._innerSectionRIds.clear();
    this._watermarkPlan = undefined;
    this._finalWatermarkTypes = [];
    // _registeredHyperlinks is a WeakSet; old entries become unreachable
    // along with the body model objects they referenced — no manual clear
    // is necessary. _chartNum is a WeakMap with the same property.
    this._initZip();
    return this;
  }

  // ===========================================================================
  // Private: ZIP infrastructure
  // ===========================================================================

  private _initZip(): void {
    this._outputChunks = [];
    this._streamError = null;
    this._pendingDrain = Promise.resolve();
    this._pendingCompression.clear();
    // XML collected for the previous document's `document.xml` belongs to that document: it was never written, and
    // must not be written into the next one.
    this._xmlBuffer.reset();
    if (this._options.sink && !this._sinkAdapter) {
      this._sinkAdapter = new SinkAdapter(this._options.sink);
    }
    this._documentRels = createRelationships();
    this._allocateHeaderFooterRIds();
    this._renderCtx = createRenderContext({
      securityPolicy: this._options.securityPolicy,
      chartRIds: new Map(),
      imageRIdRemap: new Map(),
      hyperlinkRIds: new WeakMap(),
      // Hyperlinks that body markup kept verbatim refers to (a drawing's
      // `a:hlinkClick`) register as the markup is written; document.xml.rels
      // is written after the body.
      preservedLinkRId: preservedLinkRegistrar(this._documentRels)
    });
    this._zipComplete = new Promise<void>(resolve => {
      this._resolveZipComplete = resolve;
    });
    this._zip = new Zip((err, data, _final) => {
      // The ZIP callback reports compression / framing errors out-of-band.
      // Capture only the first error; subsequent callbacks may still be
      // dispatched as the pipeline drains. Surfaced from `finalize()`.
      if (err && !this._streamError) {
        this._streamError = err;
        // Wake up any pending `_endStream` waiter so callers don't hang.
        this._documentStream?.emit("error", err);
      }
      if (data && data.length > 0) {
        if (this._sinkAdapter) {
          // The Zip callback runs synchronously and cannot await, so we
          // chain each sink write onto a single drain promise. Producers
          // either:
          //   (a) call `addAsync()` / `finalize()` which await the chain,
          //   (b) ignore backpressure and let chunks queue in memory until
          //       the next await point, capped by the sink's own queueing.
          // First-error capture: a rejected write is collapsed into
          // `_streamError` so subsequent writes are skipped quickly.
          this._pendingDrain = this._pendingDrain.then(() => {
            if (this._streamError) {
              return;
            }
            return this._sinkAdapter!.write(data).catch((e: unknown) => {
              if (!this._streamError) {
                this._streamError = e instanceof Error ? e : new Error(String(e));
              }
            });
          });
        } else {
          this._outputChunks.push(data);
        }
      }
      // Terminal callback: the archive (incl. trailing central directory) is
      // fully emitted. On the async browser deflate path this fires after
      // `_zip.end()` returns, so `finalize()` must await `_zipComplete`.
      if (_final || err) {
        this._resolveZipComplete?.();
      }
    });

    // Create the document.xml ZIP entry and stream
    const level = this._options.compressionLevel ?? 6;
    this._documentZipFile = new ZipDeflate(PartPath.Document, {
      level,
      modTime: this._entryModTime()
    });
    this._zip.add(this._documentZipFile);

    this._documentStream = new StreamBuf({ bufSize: 65536 });
    this._documentStream.on("data", (chunk: Uint8Array) => {
      this._trackCompression(this._documentZipFile.push(chunk));
    });
    this._documentStream.once("finish", () => {
      void this._trackCompression(this._documentZipFile.push(EMPTY_U8, true)).then(() => {
        this._documentStream.emit("zipped");
      });
    });
  }

  private _trackCompression(completion: Promise<void>): Promise<void> {
    const tracked = completion.catch((err: unknown) => {
      if (!this._streamError) {
        this._streamError = err instanceof Error ? err : new Error(String(err));
      }
    });
    this._pendingCompression.add(tracked);
    void tracked.then(() => this._pendingCompression.delete(tracked));
    return tracked;
  }

  private async _awaitOutputBoundary(message: string): Promise<void> {
    // Pushes made while waiting are waited for too, as the chain did.
    while (this._pendingCompression.size > 0) {
      await Promise.all(this._pendingCompression);
    }
    // Compression callbacks enqueue sink writes synchronously. Capture the
    // drain only after compression completes so browser output is included.
    await this._pendingDrain;
    if (this._streamError) {
      throw new DocxWriteError(`Stream.DocumentWriter: ${message} (${this._streamError.message})`, {
        cause: this._streamError
      });
    }
  }

  /**
   * Append XML to `document.xml`. Collected and handed to the compressor 64 KB at a time rather than once per
   * element: each hand-off copied the element's bytes out and pushed them through the ZIP layer on their own, which
   * for short paragraphs cost more than rendering them. {@link _flushXml} sends what is collected; it runs before
   * `document.xml` is closed and before `addAsync` waits on the sink, so the bytes and their order are unchanged.
   */
  private _write(text: string): void {
    this._xmlBuffer.addText(text);
    if (this._xmlBuffer.length >= XML_FLUSH_BYTES) {
      this._flushXml();
    }
  }

  private _flushXml(): void {
    if (this._xmlBuffer.length > 0) {
      this._documentStream.write(this._xmlBuffer);
      this._xmlBuffer.reset();
    }
  }

  // ===========================================================================
  // Private: Document XML generation
  // ===========================================================================

  private _writeDocumentHeader(): void {
    // XML declaration
    let header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
    header += `<w:document`;
    for (const [key, value] of Object.entries(DOCUMENT_NAMESPACES)) {
      header += ` ${key}="${value}"`;
    }
    header += `>`;

    // Background
    if (this._options.background) {
      const bg = this._options.background;
      header += `<w:background w:color="${xmlEncodeAttr(bg.color ?? "FFFFFF")}"`;
      if (bg.themeColor) {
        header += ` w:themeColor="${xmlEncodeAttr(bg.themeColor)}"`;
      }
      header += `/>`;
    }

    header += `<w:body>`;
    this._write(header);
  }

  private _writeBodyElement(element: BodyContent): void {
    // Serialize a single body element using the shared renderBodyContent
    // function. We pass the writer's accumulated render context so r:embed,
    // chart and hyperlink rIds resolve correctly. Without this, charts throw
    // "Chart content was not registered with a relationship id" and hyperlink
    // / image references would be missing or wrong.
    const writer = new XmlWriter();
    renderBodyContent(writer, element, this._renderCtx);
    this._write(writer.xml);
  }

  /**
   * Emit an empty `<w:p/>` to separate two adjacent tables. Required by
   * ECMA-376 §17.13.5.34 — Word rejects packages where two `<w:tbl>`
   * elements appear without a paragraph between them.
   */
  private _writeSeparatorParagraph(): void {
    this._write("<w:p/>");
  }

  /**
   * Scan a single body element and register any chart / hyperlink / image
   * references it introduces against the writer's accumulated state. This
   * must run before the element is serialized so the render context already
   * carries the relationships the renderer will look up.
   */
  private _registerElementReferences(element: BodyContent): void {
    // Direct top-level chart entries
    if (element.type === "chart") {
      this._registerChart(element);
      return;
    }
    if (element.type === "chartEx") {
      this._registerChartEx(element);
      return;
    }
    // A paragraph of runs and range markers — nearly every streamed element —
    // registers its pictures and styles without the walker.
    if (element.type === "paragraph" && this._registerSimpleParagraph(element)) {
      return;
    }
    for (const id of referencedStyleIds({ body: [element] })) {
      this._styleRefs.add(id);
    }
    // For paragraph-like containers we descend with the shared walker so
    // track-change wrappers (InsertedRun / MovedToRun / hyperlink children)
    // are also covered.
    if (
      element.type === "paragraph" ||
      element.type === "table" ||
      element.type === "sdt" ||
      element.type === "textBox" ||
      element.type === "tableOfContents"
    ) {
      walkBlocks([element], this._referenceVisitor);
      return;
    }
    if (element.type === "floatingImage" && element.rId) {
      this._registerImageRId(element.rId);
      if (element.svgRId) {
        this._registerImageRId(element.svgRId);
      }
    }
  }

  private _registerParagraphReferences(_para: Paragraph): void {
    // Per-paragraph property registration is not needed today — image and
    // hyperlink registration is handled by enterRun / enterHyperlink. This
    // hook exists so future paragraph-level relationships (numPicBullet,
    // tabLeader image, …) can be added without changing call sites.
  }

  /** Register the pictures a run shows. */
  private _registerRunImages(run: Run): void {
    for (const c of run.content) {
      if (c.type === "image" && c.rId) {
        this._registerImageRId(c.rId);
        if (c.svgRId) {
          this._registerImageRId(c.svgRId);
        }
      }
    }
  }

  /**
   * Register the references of a paragraph whose children are all runs and
   * range markers — what the walker would register, without walking — and
   * return `true`; return `false` for any other paragraph, which must be walked.
   */
  private _registerSimpleParagraph(para: Paragraph): boolean {
    const children = para.children;
    for (let i = 0; i < children.length; i++) {
      const child = children[i]!;
      if ("type" in child && !isRangeMarker(child)) {
        return false;
      }
    }
    this._registerParagraphReferences(para);
    this._addStyleRef(para.properties?.style);
    for (let i = 0; i < children.length; i++) {
      const child = children[i]!;
      if (isRun(child)) {
        this._registerRunImages(child);
        this._addStyleRef(child.properties?.style);
      }
    }
    return true;
  }

  private _addStyleRef(id: string | undefined): void {
    if (id) {
      this._styleRefs.add(id);
    }
  }

  private _registerChart(chart: ChartContent): void {
    if (this._renderCtx.chartRIds.has(chart)) {
      return;
    }
    const num = ++this._nextChartSeq;
    const rId = addRelationship(this._documentRels, RelType.Chart, `charts/chart${num}.xml`);
    this._renderCtx.chartRIds.set(chart, rId);
    this._chartNum.set(chart, num);
    this._bodyCharts.push(chart);
  }

  private _registerChartEx(chart: ChartExContent): void {
    if (this._renderCtx.chartRIds.has(chart)) {
      return;
    }
    const num = ++this._nextChartExSeq;
    const rId = addRelationship(this._documentRels, RelType.ChartEx, `charts/chartEx${num}.xml`);
    this._renderCtx.chartRIds.set(chart, rId);
    this._chartNum.set(chart, num);
    this._bodyChartEx.push(chart);
  }

  private _registerHyperlink(h: { readonly url?: string; readonly rId?: string }): void {
    if (!h.url || h.rId || this._registeredHyperlinks.has(h)) {
      return;
    }
    // Drop dangerous schemes (javascript:, vbscript:, data:, file:, …) before
    // they reach document.xml.rels. Mark the link as registered either way so
    // we don't keep re-evaluating it on every flush.
    const safe = sanitizeUrl(h.url);
    if (!safe) {
      this._registeredHyperlinks.add(h);
      return;
    }
    const rId = addRelationship(this._documentRels, RelType.Hyperlink, safe, "External");
    this._renderCtx.hyperlinkRIds.set(h, rId);
    this._registeredHyperlinks.add(h);
  }

  private _registerImageRId(rId: string): void {
    if (this._registeredImageRIds.has(rId)) {
      return;
    }
    const img = this._lookupImage(rId);
    if (!img) {
      // Image reference points at a binary the caller did not provide. The
      // previous behaviour was to silently skip and emit an invalid DOCX —
      // see the policy field on StreamingDocxOptions.
      const policy = this._options.missingImagePolicy ?? "throw";
      if (policy === "warn") {
        console.warn(
          `[Stream.DocumentWriter] image rId "${rId}" referenced by content ` +
            `but not present in options.images. Output will be missing this ` +
            `relationship and may not open in Word.`
        );
        return;
      }
      throw new DocxWriteError(
        `image rId "${rId}" referenced by content but ` +
          `not present in options.images. Add the image to options.images, ` +
          `remove the reference, or set missingImagePolicy: "warn" to ` +
          `accept a broken document.`
      );
    }
    addRelationshipWithId(this._documentRels, rId, RelType.Image, `media/${img.fileName}`);
    this._registeredImageRIds.add(rId);
  }

  /** The image a key (`rId` or one of `aliasRIds`) names. */
  private _lookupImage(rId: string): ImageDef | undefined {
    return this._imageIndex.get(rId);
  }

  /**
   * Register the pictures `content` shows in a part's own `.rels` and return
   * the remap its renderer writes `r:embed` through.
   */
  /** Register an image watermark's picture in the rels of the header drawing it; its id there. */
  private _registerWatermarkImage(rels: RelationshipsState): string | undefined {
    const watermark = this._options.watermark;
    if (watermark?.type !== "image") {
      return undefined;
    }
    const remap = registerPartImages(rels, [watermark.rId], this._imageIndex, streamingImageTarget);
    return remap.get(watermark.rId) ?? watermark.rId;
  }

  private _registerPartImages(rels: RelationshipsState, content: unknown): Map<string, string> {
    const keys = new Set<string>();
    collectImageKeys(content, this._imageIndex, keys);
    return registerPartImages(rels, keys, this._imageIndex, streamingImageTarget);
  }

  private _writeDocumentFooter(): void {
    // Word rejects a `<w:body/>` that contains no `<w:p>`. Synthesise an
    // empty paragraph when the caller streamed zero elements so the
    // package opens cleanly. (Bulk packager handles this implicitly via
    // `Document.create`'s default body element list.)
    if (this._elementCount === 0) {
      this._write("<w:p/>");
    }

    if (this._options.bodyRangeMarkersAfter?.length) {
      const writer = new XmlWriter();
      renderRangeMarkers(
        writer,
        this._options.bodyRangeMarkersAfter,
        renderHelpersFor(this._renderCtx)
      );
      this._write(writer.xml);
    }

    // A range still open here cannot be removed, as the bulk writer removes
    // it: its start has been streamed out already. It is closed at the end of
    // the body instead, so the part still holds whole ranges only.
    const unclosed = new XmlWriter();
    for (const id of this._renderCtx.bookmarkIds?.closePart() ?? []) {
      unclosed.leafNode("w:bookmarkEnd", { "w:id": xmlValue(id, "w:id") });
    }
    for (const id of this._renderCtx.permIds?.closePart() ?? []) {
      unclosed.leafNode("w:permEnd", { "w:id": xmlValue(id, "w:id") });
    }
    if (unclosed.xml) {
      this._write(unclosed.xml);
    }

    // Write final section properties if provided. Header/footer references
    // are rewritten so they refer to the rIds we just allocated in
    // `_allocateHeaderFooterRIds`. References whose target type cannot
    // be resolved (e.g. a custom rId for a header that isn't in the
    // options map) are dropped rather than emitted dangling.
    let sect = this._finalSectionWithHeaders();
    if (
      this._options.footers &&
      this._options.footers.size > 0 &&
      (!sect?.footers || sect.footers.length === 0)
    ) {
      const synth = this._synthesizeFooterRefs();
      if (synth.length > 0) {
        sect = { ...(sect ?? {}), footers: synth };
      }
    }
    // The header types the section shows but neither names nor inherits get
    // the watermark-only header.
    if (this._watermarkHeaderRid && this._finalWatermarkTypes.length > 0) {
      const headers = sect?.headers ? [...sect.headers] : [];
      for (const type of this._finalWatermarkTypes) {
        headers.push({ type, rId: this._watermarkHeaderRid });
      }
      sect = { ...(sect ?? {}), headers };
    }

    if (sect) {
      const writer = new XmlWriter();
      const section = sect;
      within("options.sectionProperties", () => renderSectionProperties(writer, section));
      this._write(writer.xml);
    } else {
      // OOXML CT_Body requires a final <w:sectPr> so Word knows the page
      // geometry. When the caller didn't provide one, fall back to the
      // same default that Document.create() uses (US Letter, 1" margins).
      const writer = new XmlWriter();
      renderSectionProperties(writer, {
        pageSize: { width: 12240, height: 15840 },
        margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      });
      this._write(writer.xml);
    }

    // Close body and document
    this._write(`</w:body></w:document>`);
  }

  /**
   * Allocate header/footer relationship IDs deterministically (in the same
   * order auxiliary parts will be emitted), before anything is streamed: a
   * section break inside the body is written as it streams past, and names
   * its headers and footers by these ids.
   */
  private _allocateHeaderFooterRIds(): void {
    this._headerKeyToRid.clear();
    this._footerKeyToRid.clear();
    if (this._options.headers) {
      let idx = 1;
      for (const [key] of this._options.headers) {
        const rId = addRelationship(this._documentRels, RelType.Header, `header${idx}.xml`);
        this._headerKeyToRid.set(key, rId);
        idx++;
      }
    }
    if (this._options.footers) {
      let idx = 1;
      for (const [key] of this._options.footers) {
        const rId = addRelationship(this._documentRels, RelType.Footer, `footer${idx}.xml`);
        this._footerKeyToRid.set(key, rId);
        idx++;
      }
    }
  }

  /**
   * Plan the watermark for the final section. It joins every header a
   * section shows, as the buffered writer does; a section showing none of a
   * type names a header part of the watermark's own, in the next header
   * slot. The inner sections were planned as they streamed past.
   */
  private _planFinalWatermark(): void {
    const plan = this._watermarkPlanFor();
    if (!plan) {
      return;
    }
    const final = this._finalSectionWithHeaders();
    this._finalWatermarkTypes = plan.section({
      titlePage: final?.titlePage,
      headers: final?.headers ?? []
    });
    if (plan.needsWatermarkHeader) {
      this._watermarkHeaderRid ??= this._allocateWatermarkHeaderRid();
    }
  }

  /**
   * The final section properties with header references resolved to the
   * allocated rIds, synthesised for a caller who registered headers but named
   * none. Footers are not resolved here.
   */
  private _finalSectionWithHeaders(): SectionProperties | undefined {
    const sectIn = this._options.sectionProperties;
    let sect = sectIn ? this._rewireSectionRefs(sectIn) : undefined;
    // Auto-fill: if the caller provided headers but no section references,
    // synthesize one per type. Mirrors what the bulk packager does for
    // builder-style usage.
    if (
      this._options.headers &&
      this._options.headers.size > 0 &&
      (!sect?.headers || sect.headers.length === 0)
    ) {
      const synth = this._synthesizeHeaderRefs();
      if (synth.length > 0) {
        sect = { ...(sect ?? {}), headers: synth };
      }
    }
    return sect;
  }

  /** The watermark plan, created on first use; `undefined` without a watermark. */
  private _watermarkPlanFor(): WatermarkPlan | undefined {
    if (!this._options.watermark) {
      return undefined;
    }
    this._watermarkPlan ??= planWatermark(
      this._options.settings?.evenAndOddHeaders === true,
      WATERMARK_PART
    );
    return this._watermarkPlan;
  }

  private _allocateWatermarkHeaderRid(): string {
    return addRelationship(
      this._documentRels,
      RelType.Header,
      `header${(this._options.headers?.size ?? 0) + 1}.xml`
    );
  }

  /**
   * `element` with its section break's header and footer references resolved
   * to the parts' ids — the break is written now — and naming the
   * watermark-only header for the header types the break shows but neither
   * names nor inherits, whose id is then allocated now.
   */
  private _prepareSectionBreak(element: BodyContent): BodyContent {
    const sect = element.type === "paragraph" ? element.properties?.sectionProperties : undefined;
    if (!sect || element.type !== "paragraph") {
      return element;
    }
    let rewired = this._rewireSectionRefs(sect);
    for (const ref of [...(rewired.headers ?? []), ...(rewired.footers ?? [])]) {
      this._innerSectionRIds.add(ref.rId);
    }
    const plan = this._watermarkPlanFor();
    const missing = plan ? plan.section(rewired) : [];
    if (missing.length > 0) {
      this._watermarkHeaderRid ??= this._allocateWatermarkHeaderRid();
      const rId = this._watermarkHeaderRid;
      rewired = {
        ...rewired,
        headers: [...(rewired.headers ?? []), ...missing.map(type => ({ type, rId }))]
      };
    }
    if (rewired === sect) {
      return element;
    }
    return {
      ...element,
      properties: { ...element.properties, sectionProperties: rewired }
    };
  }

  private _rewireSectionRefs(sect: SectionProperties): SectionProperties {
    const allowedHeader = new Set(this._headerKeyToRid.values());
    const allowedFooter = new Set(this._footerKeyToRid.values());

    const resolveByTypeHeader = (type: string): string | undefined => {
      // Try direct map key match first (e.g. user passed map key === rId).
      // Then fall back to a same-type lookup for the common builder case
      // where keys are "default" / "first" / "even".
      if (this._headerKeyToRid.has(type)) {
        return this._headerKeyToRid.get(type);
      }
      return undefined;
    };
    const resolveByTypeFooter = (type: string): string | undefined => {
      if (this._footerKeyToRid.has(type)) {
        return this._footerKeyToRid.get(type);
      }
      return undefined;
    };

    let out = sect;
    if (sect.headers) {
      const resolved: HeaderFooterRef[] = [];
      for (const ref of sect.headers) {
        // A model key first: it may equal an id just given to another part.
        if (ref.rId && this._headerKeyToRid.has(ref.rId)) {
          resolved.push({ ...ref, rId: this._headerKeyToRid.get(ref.rId)! });
          continue;
        }
        if (ref.rId && allowedHeader.has(ref.rId)) {
          resolved.push(ref);
          continue;
        }
        const byType = resolveByTypeHeader(ref.type);
        if (byType) {
          resolved.push({ ...ref, rId: byType });
        }
        // else drop — no matching part part, do not emit dangling rId.
      }
      if (
        resolved.length !== sect.headers.length ||
        resolved.some((r, i) => r !== sect.headers![i])
      ) {
        out = { ...out, headers: resolved };
      }
    }
    if (sect.footers) {
      const resolved: HeaderFooterRef[] = [];
      for (const ref of sect.footers) {
        // A model key first: it may equal an id just given to another part.
        if (ref.rId && this._footerKeyToRid.has(ref.rId)) {
          resolved.push({ ...ref, rId: this._footerKeyToRid.get(ref.rId)! });
          continue;
        }
        if (ref.rId && allowedFooter.has(ref.rId)) {
          resolved.push(ref);
          continue;
        }
        const byType = resolveByTypeFooter(ref.type);
        if (byType) {
          resolved.push({ ...ref, rId: byType });
        }
      }
      if (
        resolved.length !== sect.footers.length ||
        resolved.some((r, i) => r !== sect.footers![i])
      ) {
        out = { ...out, footers: resolved };
      }
    }
    return out;
  }

  /**
   * Synthesise section-property header references for every header part
   * the caller registered. Recognised type keys (`default`/`first`/`even`)
   * keep their semantics; any other key (round-tripped rId names from
   * readDocx, custom strings) falls back to `"default"` so the header is
   * actually referenced — without this fallback header parts can sit in
   * the package as dangling content. A part a section break inside the
   * body names is that section's, and is left out.
   *
   * If multiple keys map to the same logical type, only the first one is
   * kept so we don't emit two `<w:headerReference w:type="default">`
   * children (Word's behaviour with duplicates is implementation-defined).
   */
  private _synthesizeHeaderRefs(): HeaderFooterRef[] {
    const out: HeaderFooterRef[] = [];
    const seenTypes = new Set<string>();
    for (const [key, rId] of this._headerKeyToRid) {
      // A part a section break names is that section's, not the last's.
      if (this._innerSectionRIds.has(rId)) {
        continue;
      }
      const type: HeaderFooterRef["type"] =
        key === "default" || key === "first" || key === "even" ? key : "default";
      if (seenTypes.has(type)) {
        continue;
      }
      seenTypes.add(type);
      out.push({ type, rId });
    }
    return out;
  }
  private _synthesizeFooterRefs(): HeaderFooterRef[] {
    const out: HeaderFooterRef[] = [];
    const seenTypes = new Set<string>();
    for (const [key, rId] of this._footerKeyToRid) {
      // A part a section break names is that section's, not the last's.
      if (this._innerSectionRIds.has(rId)) {
        continue;
      }
      const type: HeaderFooterRef["type"] =
        key === "default" || key === "first" || key === "even" ? key : "default";
      if (seenTypes.has(type)) {
        continue;
      }
      seenTypes.add(type);
      out.push({ type, rId });
    }
    return out;
  }

  // ===========================================================================
  // Private: Auxiliary parts
  // ===========================================================================

  /** ZIP entry timestamp; see {@link StreamingDocxOptions.modTime}. */
  private _entryModTime(): Date | undefined {
    const { modTime, coreProperties } = this._options;
    return modTime ?? coreProperties?.modified ?? coreProperties?.created;
  }

  /** @param previousRevisionIds - What `enterRevisionIds` returned for this write. */
  private async _addAuxiliaryParts(previousRevisionIds: RevisionIdMap | undefined): Promise<void> {
    const level = this._options.compressionLevel ?? 6;
    const modTime = this._entryModTime();

    // Helper: add a complete XML file to the ZIP
    const writtenParts: string[] = [];
    const addFile = (path: string, data: Uint8Array, compressionLevel = level): void => {
      writtenParts.push(path);
      const file = new ZipDeflate(path, { level: compressionLevel, modTime });
      this._zip.add(file);
      this._trackCompression(file.push(data, true));
    };
    // Every part is closed as a story part: one holding no range is unchanged.
    const addXmlFile = (path: string, renderFn: (xml: XmlWriter) => void): void => {
      const writer = new XmlWriter();
      renderFn(writer);
      addFile(path, utf8Encoder.encode(closeStoryPart(writer.xml, this._renderCtx)));
    };

    // Content types and relationships
    const contentTypes = createContentTypes();
    const packageRels = createRelationships();
    // Reuse the document relationships state we have been populating during
    // add() (charts, hyperlinks, images). Adding the standard parts below
    // augments this state.
    const documentRels = this._documentRels;

    // Package relationships
    addRelationship(packageRels, RelType.OfficeDocument, "word/document.xml");
    addRelationship(packageRels, RelType.CoreProperties, "docProps/core.xml");
    addRelationship(packageRels, RelType.ExtendedProperties, "docProps/app.xml");

    // [Content_Types].xml MUST declare every part. The package
    // relationships file references docProps/core.xml + docProps/app.xml
    // even when the caller didn't supply explicit metadata, so we must
    // register their content types up-front; otherwise Word/LibreOffice
    // refuse to open the file (rejected at the OPC layer before any
    // schema validation).
    addContentTypeOverride(contentTypes, `/${PartPath.CoreProps}`, ContentType.CoreProperties);
    addContentTypeOverride(contentTypes, `/${PartPath.AppProps}`, ContentType.ExtendedProperties);

    // Document relationships
    addRelationship(documentRels, RelType.Styles, "styles.xml");
    addRelationship(documentRels, RelType.Settings, "settings.xml");
    addRelationship(documentRels, RelType.FontTable, "fontTable.xml");
    addRelationship(documentRels, RelType.Theme, "theme/theme1.xml");

    // Numbering
    const hasNumbering =
      (this._options.abstractNumberings && this._options.abstractNumberings.length > 0) ||
      (this._options.numberingInstances && this._options.numberingInstances.length > 0);
    if (hasNumbering) {
      addRelationship(documentRels, RelType.Numbering, "numbering.xml");
    }

    // Each notes part renders r:embed through the remap of its own .rels.
    let footnoteImageRemap = new Map<string, string>();
    let endnoteImageRemap = new Map<string, string>();

    // Footnotes — including their own .rels for in-note hyperlinks/images.
    if (this._options.footnotes && this._options.footnotes.length > 0) {
      addRelationship(documentRels, RelType.Footnotes, "footnotes.xml");
      addContentTypeOverride(contentTypes, `/${PartPath.Footnotes}`, ContentType.Footnotes);
      const fnRels = createRelationships();
      // Pictures first, so an image keeps its own id where it can.
      footnoteImageRemap = this._registerPartImages(fnRels, this._options.footnotes);
      const fnLinks: readonly { readonly url?: string }[] = [
        ...collectHyperlinksFromNotes(this._options.footnotes),
        ...collectPreservedLinks(this._options.footnotes)
      ];
      for (const link of fnLinks) {
        if (link.url) {
          const safe = sanitizeUrl(link.url);
          if (!safe) {
            continue;
          }
          const linkRId = addRelationship(fnRels, RelType.Hyperlink, safe, "External");
          this._renderCtx.hyperlinkRIds.set(link, linkRId);
        }
      }
      if (getRelationshipCount(fnRels) > 0) {
        addXmlFile(`word/_rels/footnotes.xml.rels`, xml => renderRelationships(fnRels, xml));
      }
      // Footnote XML rendering itself happens in buildCommonAuxiliaryParts.
    }

    // Endnotes — same treatment as footnotes.
    if (this._options.endnotes && this._options.endnotes.length > 0) {
      addRelationship(documentRels, RelType.Endnotes, "endnotes.xml");
      addContentTypeOverride(contentTypes, `/${PartPath.Endnotes}`, ContentType.Endnotes);
      const enRels = createRelationships();
      endnoteImageRemap = this._registerPartImages(enRels, this._options.endnotes);
      const enLinks: readonly { readonly url?: string }[] = [
        ...collectHyperlinksFromNotes(this._options.endnotes),
        ...collectPreservedLinks(this._options.endnotes)
      ];
      for (const link of enLinks) {
        if (link.url) {
          const safe = sanitizeUrl(link.url);
          if (!safe) {
            continue;
          }
          const linkRId = addRelationship(enRels, RelType.Hyperlink, safe, "External");
          this._renderCtx.hyperlinkRIds.set(link, linkRId);
        }
      }
      if (getRelationshipCount(enRels) > 0) {
        addXmlFile(`word/_rels/endnotes.xml.rels`, xml => renderRelationships(enRels, xml));
      }
    }

    // Comments — including their own .rels for in-comment hyperlinks/images.
    if (this._options.comments && this._options.comments.length > 0) {
      addRelationship(documentRels, RelType.Comments, "comments.xml");
      addContentTypeOverride(contentTypes, `/${PartPath.Comments}`, ContentType.Comments);

      // Register hyperlink/image rels BEFORE rendering comments.xml so the
      // emitted r:id values match the per-part .rels we are about to write.
      const cmtRels = createRelationships();
      const commentImageRemap = this._registerPartImages(cmtRels, this._options.comments);
      const commentBodies = this._options.comments.map(c => ({ content: c.content }));
      const cmtLinks: readonly { readonly url?: string }[] = [
        ...collectHyperlinksFromNotes(commentBodies),
        ...collectPreservedLinks(commentBodies)
      ];
      for (const link of cmtLinks) {
        if (link.url) {
          const safe = sanitizeUrl(link.url);
          if (!safe) {
            continue;
          }
          const linkRId = addRelationship(cmtRels, RelType.Hyperlink, safe, "External");
          this._renderCtx.hyperlinkRIds.set(link, linkRId);
        }
      }
      if (getRelationshipCount(cmtRels) > 0) {
        addXmlFile(`word/_rels/comments.xml.rels`, xml => renderRelationships(cmtRels, xml));
      }

      addXmlFile(PartPath.Comments, xml =>
        renderComments(xml, this._options.comments!, {
          imageRemap: commentImageRemap,
          hyperlinkRIds: this._renderCtx.hyperlinkRIds,
          nextDocPrId: this._renderCtx.ids.nextDocPrId,
          bookmarkIds: this._renderCtx.bookmarkIds,
          permIds: this._renderCtx.permIds,
          rawXmlPolicy: this._renderCtx.rawXmlPolicy
        })
      );
      // Also write commentsExtended if any have done/parentId
      const hasExtended = this._options.comments.some(c => c.done != null || c.parentId != null);
      if (hasExtended) {
        addRelationship(documentRels, RelType.CommentsExtended, "commentsExtended.xml");
        addContentTypeOverride(
          contentTypes,
          `/${PartPath.CommentsExtended}`,
          ContentType.CommentsExtended
        );
        addXmlFile(PartPath.CommentsExtended, xml =>
          renderCommentsExtended(xml, this._options.comments!)
        );
      }
    }

    // Headers
    //
    // Each header part has its own .rels file. Image / hyperlink / chart
    // references inside header content must register against THAT part's
    // .rels — they are not document.xml.rels relationships, so we mirror
    // the bulk packager's behaviour here to avoid producing dangling
    // r:embed / r:id values inside header XML.
    //
    // The document-level rId for each header was already allocated during
    // `_allocateHeaderFooterRIds` so that section properties and header
    // parts agree.
    let nextHeaderIdx = 1;
    const watermark = this._options.watermark;
    if (this._options.headers) {
      for (const [headerKey, headerDef] of this._options.headers) {
        const headerIdx = nextHeaderIdx++;
        const headerPath = PartPath.header(headerIdx);
        addContentTypeOverride(contentTypes, `/${headerPath}`, ContentType.Header);

        // Register relationships BEFORE rendering the header XML, otherwise
        // the writer cannot resolve the freshly-allocated hyperlink rIds and
        // emits dangling r:id values. addXmlFile() invokes the render
        // callback synchronously, so the order matters.
        const hRels = createRelationships();
        // Images: the header's own pictures, registered in its own .rels; its
        // XML writes r:embed through the remap this returns.
        const headerImageRemap = this._registerPartImages(hRels, headerDef.content);
        // Hyperlinks: same scheme as bulk packager — assign a fresh rId per
        // header for any URL-bearing hyperlink and surface it via
        // hyperlinkRIds so the writer emits matching r:id.
        const hLinks: readonly { readonly url?: string }[] = [
          ...collectHyperlinksFromHeaderFooter(headerDef.content),
          ...collectPreservedLinks(headerDef.content)
        ];
        for (const link of hLinks) {
          if (link.url) {
            const safe = sanitizeUrl(link.url);
            if (!safe) {
              continue;
            }
            const linkRId = addRelationship(hRels, RelType.Hyperlink, safe, "External");
            this._renderCtx.hyperlinkRIds.set(link, linkRId);
          }
        }
        // Charts: collect into _bodyCharts so a chart part is generated, and
        // register the rel against the header's own .rels.
        const headerCharts: ChartContent[] = [];
        collectChartsFromHeaderFooter(headerDef.content, headerCharts);
        for (const chartContent of headerCharts) {
          if (this._renderCtx.chartRIds.has(chartContent)) {
            continue;
          }
          const num = ++this._nextChartSeq;
          const rId = addRelationship(hRels, RelType.Chart, `charts/chart${num}.xml`);
          this._renderCtx.chartRIds.set(chartContent, rId);
          this._chartNum.set(chartContent, num);
          this._bodyCharts.push(chartContent);
        }

        let headerWatermark: HeaderWatermark | undefined;
        if (watermark && this._watermarkPlan?.hosts.has(this._headerKeyToRid.get(headerKey)!)) {
          headerWatermark = { watermark, imageRId: this._registerWatermarkImage(hRels) };
        }

        addXmlFile(headerPath, xml =>
          renderHeader(
            xml,
            headerDef.content,
            {
              imageRemap: headerImageRemap,
              hyperlinkRIds: this._renderCtx.hyperlinkRIds,
              nextDocPrId: this._renderCtx.ids.nextDocPrId,
              bookmarkIds: this._renderCtx.bookmarkIds,
              permIds: this._renderCtx.permIds,
              rawXmlPolicy: this._renderCtx.rawXmlPolicy
            },
            headerWatermark
          )
        );

        if (getRelationshipCount(hRels) > 0) {
          addXmlFile(`word/_rels/header${headerIdx}.xml.rels`, xml =>
            renderRelationships(hRels, xml)
          );
        }
      }
    }

    // Watermark without a default header to join — its own header part,
    // after the caller's. Its rId was allocated during
    // `_allocateHeaderFooterRIds`.
    if (watermark && this._watermarkHeaderRid) {
      const watermarkIdx = nextHeaderIdx++;
      const watermarkPath = PartPath.header(watermarkIdx);
      addContentTypeOverride(contentTypes, `/${watermarkPath}`, ContentType.Header);
      // Image watermarks need a per-header relationship to the image binary.
      const wmRels = createRelationships();
      const watermarkImageRId = this._registerWatermarkImage(wmRels);
      addXmlFile(watermarkPath, xml => renderWatermarkHeader(xml, watermark, watermarkImageRId));
      if (getRelationshipCount(wmRels) > 0) {
        addXmlFile(`word/_rels/header${watermarkIdx}.xml.rels`, xml =>
          renderRelationships(wmRels, xml)
        );
      }
    }

    // Footers — document-level rIds already allocated during
    // `_allocateHeaderFooterRIds`.
    if (this._options.footers) {
      let footerIdx = 1;
      for (const [, footerDef] of this._options.footers) {
        const footerPath = PartPath.footer(footerIdx);
        addContentTypeOverride(contentTypes, `/${footerPath}`, ContentType.Footer);

        // Register relationships BEFORE rendering. addXmlFile() runs the
        // callback synchronously so any hyperlink rIds the renderer needs
        // must already be in this._renderCtx.hyperlinkRIds.
        const fRels = createRelationships();
        const footerImageRemap = this._registerPartImages(fRels, footerDef.content);
        const fLinks: readonly { readonly url?: string }[] = [
          ...collectHyperlinksFromHeaderFooter(footerDef.content),
          ...collectPreservedLinks(footerDef.content)
        ];
        for (const link of fLinks) {
          if (link.url) {
            const safe = sanitizeUrl(link.url);
            if (!safe) {
              continue;
            }
            const linkRId = addRelationship(fRels, RelType.Hyperlink, safe, "External");
            this._renderCtx.hyperlinkRIds.set(link, linkRId);
          }
        }
        const footerCharts: ChartContent[] = [];
        collectChartsFromHeaderFooter(footerDef.content, footerCharts);
        for (const chartContent of footerCharts) {
          if (this._renderCtx.chartRIds.has(chartContent)) {
            continue;
          }
          const num = ++this._nextChartSeq;
          const rId = addRelationship(fRels, RelType.Chart, `charts/chart${num}.xml`);
          this._renderCtx.chartRIds.set(chartContent, rId);
          this._chartNum.set(chartContent, num);
          this._bodyCharts.push(chartContent);
        }

        addXmlFile(footerPath, xml =>
          renderFooter(xml, footerDef.content, {
            imageRemap: footerImageRemap,
            hyperlinkRIds: this._renderCtx.hyperlinkRIds,
            nextDocPrId: this._renderCtx.ids.nextDocPrId,
            bookmarkIds: this._renderCtx.bookmarkIds,
            permIds: this._renderCtx.permIds,
            rawXmlPolicy: this._renderCtx.rawXmlPolicy
          })
        );

        if (getRelationshipCount(fRels) > 0) {
          addXmlFile(`word/_rels/footer${footerIdx}.xml.rels`, xml =>
            renderRelationships(fRels, xml)
          );
        }
        footerIdx++;
      }
    }

    // Custom properties
    if (this._options.customProperties && this._options.customProperties.length > 0) {
      addRelationship(packageRels, RelType.CustomProperties, "docProps/custom.xml");
      addContentTypeOverride(
        contentTypes,
        `/${PartPath.CustomProps}`,
        ContentType.CustomProperties
      );
      // XML rendering handled by buildCommonAuxiliaryParts below
    }

    // Images. Only register images here that were not already registered
    // by `_registerImageRId` during add(). For round-tripped models, the
    // image's `aliasRIds` may have been registered earlier (when a
    // header/footer body referenced the image under one of those alias
    // names) — treat any of those aliases as "already registered" so we
    // don't duplicate the relationship under the canonical rId. Images
    // supplied via options but never referenced in body content also get
    // registered (since the user clearly intended them to be part of the
    // document) but with an anonymous rId.
    if (this._options.images) {
      const extensions = new Set<string>();
      for (const img of this._options.images) {
        const ext = getFileExt(img.fileName);
        if (ext) {
          extensions.add(ext);
        }
        const alreadyRegistered =
          (img.rId && this._registeredImageRIds.has(img.rId)) ||
          (img.aliasRIds && img.aliasRIds.some(a => this._registeredImageRIds.has(a)));
        if (alreadyRegistered) {
          continue;
        }
        if (img.rId) {
          addRelationshipWithId(documentRels, img.rId, RelType.Image, `media/${img.fileName}`);
          this._registeredImageRIds.add(img.rId);
        } else {
          addRelationship(documentRels, RelType.Image, `media/${img.fileName}`);
        }
      }
      addImageContentTypeDefaults(contentTypes, extensions);
      for (const img of this._options.images) {
        const override = imagePartOverride(img);
        if (override) {
          addContentTypeOverride(contentTypes, `/word/media/${img.fileName}`, override);
        }
      }
    }

    // Custom XML parts (for SDT data binding)
    if (this._options.customXmlParts && this._options.customXmlParts.length > 0) {
      this._options.customXmlParts.forEach((part, i) => {
        const num = i + 1;
        const itemPath = `word/customXml/item${num}.xml`;
        const propsPath = `word/customXml/itemProps${num}.xml`;

        // Write the XML content
        const itemData = utf8Encoder.encode(part.xmlContent);
        addFile(itemPath, itemData);

        // Write itemProps*.xml
        const propsWriter = new XmlWriter();
        propsWriter.openXml(STD_DOC_ATTRIBUTES);
        propsWriter.openNode("ds:datastoreItem", {
          "ds:itemID": `{${part.itemId}}`,
          "xmlns:ds": "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
        });
        if (part.schemaReferences && part.schemaReferences.length > 0) {
          propsWriter.openNode("ds:schemaRefs");
          for (const uri of part.schemaReferences) {
            propsWriter.leafNode("ds:schemaRef", { "ds:uri": uri });
          }
          propsWriter.closeNode();
        } else {
          propsWriter.leafNode("ds:schemaRefs");
        }
        propsWriter.closeNode();
        const propsData = utf8Encoder.encode(propsWriter.xml);
        addFile(propsPath, propsData);

        // Write item rels (links itemN.xml → itemPropsN.xml)
        const itemRels = createRelationships();
        addRelationship(itemRels, RelType.CustomXmlProps, `itemProps${num}.xml`);
        addXmlFile(`word/customXml/_rels/item${num}.xml.rels`, xml =>
          renderRelationships(itemRels, xml)
        );

        // Register content types
        addContentTypeOverride(
          contentTypes,
          `/word/customXml/itemProps${num}.xml`,
          "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
        );

        // Add to document rels
        addRelationship(documentRels, RelType.CustomXml, `customXml/item${num}.xml`);
      });
    }

    // Embedded fonts
    if (this._options.embeddedFonts && this._options.embeddedFonts.length > 0) {
      const fontTableRels = createRelationships();

      for (const ef of this._options.embeddedFonts) {
        const partPath = `word/fonts/${ef.fileName}`;
        addFile(partPath, ef.data, 0);

        // Register relationship from fontTable.xml
        addRelationshipWithId(fontTableRels, ef.rId, RelType.Font, `fonts/${ef.fileName}`);

        // Register content type for .odttf / .ttf / .otf
        const ext = getFileExt(ef.fileName);
        if (ext === "odttf") {
          addContentTypeDefault(contentTypes, "odttf", ContentType.ObfuscatedFont);
        } else if (ext === "ttf") {
          addContentTypeDefault(contentTypes, "ttf", "application/x-font-ttf");
        } else if (ext === "otf") {
          addContentTypeDefault(contentTypes, "otf", "application/x-font-otf");
        }
      }

      // Write fontTable.xml.rels
      addXmlFile("word/_rels/fontTable.xml.rels", xml => renderRelationships(fontTableRels, xml));
    }

    // Opaque (unrecognized) parts for round-trip preservation
    if (this._options.opaqueParts) {
      const preserveOle = this._options.securityPolicy?.preserveOleObjects !== false;
      const dropSignatures = this._options.securityPolicy?.dropSignaturesOnModify !== false;
      for (const part of this._options.opaqueParts) {
        // An empty name ending in "/" is a ZIP folder entry, not a part (OPC
        // §9.1.1.1): never written. Anything else invalid is refused below.
        if (part.path.endsWith("/") && part.data.length === 0) {
          continue;
        }
        // Honour `preserveOleObjects`: skip OLE binaries when disabled.
        if (
          !preserveOle &&
          (part.path.startsWith("word/embeddings/") ||
            (part.path.endsWith(".bin") && part.path.includes("embed")))
        ) {
          continue;
        }
        // Honour `dropSignaturesOnModify`: signatures cease to be valid the
        // moment the document is re-serialised, so by default we drop them.
        if (dropSignatures && part.path.startsWith("_xmlsignatures/")) {
          continue;
        }
        addFile(part.path, part.data);

        // Register content type
        if (part.contentType) {
          addContentTypeOverride(contentTypes, `/${part.path}`, part.contentType);
        }

        // Write part relationships if any
        if (part.relationships && part.relationships.length > 0) {
          const partRels = createRelationships();
          for (const rel of part.relationships) {
            addRelationshipWithId(partRels, rel.id, rel.type, rel.target, rel.targetMode);
          }
          const relsPath = getPartRelsPath(part.path);
          addXmlFile(relsPath, xml => renderRelationships(partRels, xml));
        }
      }
    }

    // Content type overrides
    addContentTypeOverride(contentTypes, `/${PartPath.Document}`, ContentType.Document);
    addContentTypeOverride(contentTypes, `/${PartPath.Styles}`, ContentType.Styles);
    addContentTypeOverride(contentTypes, `/${PartPath.Settings}`, ContentType.Settings);
    addContentTypeOverride(contentTypes, `/${PartPath.FontTable}`, ContentType.FontTable);
    addContentTypeOverride(contentTypes, `/${PartPath.Theme}`, ContentType.Theme);

    if (hasNumbering) {
      addContentTypeOverride(contentTypes, `/${PartPath.Numbering}`, ContentType.Numbering);
    }

    // Write common auxiliary parts (styles, settings, fontTable, theme, numbering, properties)
    // using the shared builder to avoid duplicating render logic with docx-packager.
    const commonParts = buildCommonAuxiliaryParts({
      // Everything this writer produces is built by the library: it states the
      // defaults the library lays it out with (see `withStatedDefaults`).
      docDefaults: withStatedDefaults(this._options.docDefaults),
      // A built-in style the content uses is defined as Word defines it on
      // first use (see `withReferencedBuiltInStyles`).
      styles: withReferencedBuiltInStyles(this._options.styles, [
        ...this._styleRefs,
        ...referencedStyleIds({
          body: [],
          headers: this._options.headers,
          footers: this._options.footers,
          footnotes: this._options.footnotes,
          endnotes: this._options.endnotes,
          comments: this._options.comments,
          abstractNumberings: this._options.abstractNumberings,
          watermark: this._options.watermark
        })
      ]),
      settings: this._options.settings,
      fonts: this._options.fonts,
      theme: this._options.theme,
      abstractNumberings: this._options.abstractNumberings,
      numberingInstances: this._options.numberingInstances,
      footnotes: this._options.footnotes,
      endnotes: this._options.endnotes,
      coreProperties: this._options.coreProperties,
      appProperties: this._options.appProperties,
      customProperties: this._options.customProperties,
      rawXmlPolicy: this._renderCtx.rawXmlPolicy,
      // Pass the shared rId tables so in-note hyperlinks (registered earlier
      // against footnotes.xml.rels / endnotes.xml.rels) resolve to the same
      // rIds the renderer is about to emit.
      notesHelpers: {
        imageRemap: footnoteImageRemap,
        hyperlinkRIds: this._renderCtx.hyperlinkRIds,
        nextDocPrId: this._renderCtx.ids.nextDocPrId,
        bookmarkIds: this._renderCtx.bookmarkIds,
        permIds: this._renderCtx.permIds,
        rawXmlPolicy: this._renderCtx.rawXmlPolicy
      },
      endnotesHelpers: {
        imageRemap: endnoteImageRemap,
        hyperlinkRIds: this._renderCtx.hyperlinkRIds,
        nextDocPrId: this._renderCtx.ids.nextDocPrId,
        bookmarkIds: this._renderCtx.bookmarkIds,
        permIds: this._renderCtx.permIds,
        rawXmlPolicy: this._renderCtx.rawXmlPolicy
      }
    });

    for (const part of commonParts) {
      const data = utf8Encoder.encode(part.content);
      addFile(part.path, data);
    }

    // Write images
    if (this._options.images) {
      for (const img of this._options.images) {
        addFile(PartPath.media(img.fileName), img.data, 0);
      }
    }

    // Write chart parts (chartN.xml + content type) for body charts
    // registered during add(). Use the per-chart sequence number captured
    // at registration time so the rId target encoded in documentRels and
    // the actual ZIP entry name agree.
    for (const chartContent of this._bodyCharts) {
      const num = this._chartNum.get(chartContent);
      if (num === undefined) {
        continue; // shouldn't happen; defensive
      }
      const chartPath = `word/charts/chart${num}.xml`;
      const w = new XmlWriter();
      renderChartPart(w, chartContent.chart, { rawXmlPolicy: this._renderCtx.rawXmlPolicy });
      const data = utf8Encoder.encode(w.xml);
      addFile(chartPath, data);
      addContentTypeOverride(contentTypes, `/${chartPath}`, ContentType.Chart);
    }

    // Write ChartEx parts (raw cx:chartSpace XML preserved on the model)
    for (const cxContent of this._bodyChartEx) {
      const num = this._chartNum.get(cxContent);
      if (num === undefined) {
        continue;
      }
      const cxPath = `word/charts/chartEx${num}.xml`;
      const data = utf8Encoder.encode(cxContent.chartExXml);
      addFile(cxPath, data);
      addContentTypeOverride(contentTypes, `/${cxPath}`, ContentType.ChartEx);
    }

    // Write document.xml.rels
    addXmlFile(PartPath.DocumentRels, xml => renderRelationships(documentRels, xml));

    // Write _rels/.rels
    addXmlFile(PartPath.PackageRels, xml => renderRelationships(packageRels, xml));

    // Write [Content_Types].xml
    typeUntypedParts(
      contentTypes,
      writtenParts.filter(path => !path.endsWith(".rels"))
    );
    addXmlFile(PartPath.ContentTypes, xml => renderContentTypes(contentTypes, xml));

    leaveRevisionIds(previousRevisionIds, this._renderCtx.revisionIds);

    // Every final push is asynchronous in browsers. Do not let finalize cross
    // the auxiliary-part boundary until their compression and sink output do.
    await this._awaitOutputBoundary("auxiliary part write failed");
  }

  private _endStream(stream: StreamBuf): Promise<void> {
    return new Promise((resolve, reject) => {
      // If a prior callback already reported an error, surface it
      // synchronously instead of waiting for an event that won't fire.
      if (this._streamError) {
        reject(this._streamError);
        return;
      }
      stream.once("zipped", () => resolve());
      stream.once("error", (err: Error) => reject(err));
      if (stream === this._documentStream) {
        this._flushXml();
      }
      stream.end();
    });
  }

  private _assembleOutput(): Uint8Array {
    if (this._outputChunks.length === 1) {
      return this._outputChunks[0]!;
    }
    const total = this._outputChunks.reduce((sum, c) => sum + c.length, 0);
    const result = new Uint8Array(total);
    let offset = 0;
    for (const chunk of this._outputChunks) {
      result.set(chunk, offset);
      offset += chunk.length;
    }
    return result;
  }
}

/**
 * The writer's options, each checked as the `Document.*` setter for the same
 * field checks it — so a bad value is refused when the writer is created, not
 * at `finalize()` after every element has streamed.
 */
function checkStreamingOptions(options: unknown, where: string): void {
  checkPackageOptions(options, "options", where);
  const o = options as StreamingDocxOptions | undefined;
  if (o === undefined) {
    return;
  }
  checkInteger(o.chunkSize, "options.chunkSize", where, { optional: true, min: 1 });
  if (o.sectionProperties !== undefined) {
    checkSectionProperties(o.sectionProperties, "options.sectionProperties", where);
  }
  for (const kind of ["headers", "footers"] as const) {
    const parts = o[kind];
    if (parts === undefined) {
      continue;
    }
    if (!(parts instanceof Map)) {
      invalidArgument(where, `options.${kind}`, "a Map of key to part", parts);
    }
    for (const [key, part] of parts) {
      const at = `options.${kind}.get(${JSON.stringify(key)})`;
      checkRecord(part, at, where);
      checkRecord(part.content, `${at}.content`, where);
      checkNodeArray(part.content.children, `${at}.content.children`, where);
    }
  }
  if (o.styles !== undefined) {
    checkArray(o.styles, "options.styles", where);
    o.styles.forEach((style, i) => checkStyleDef(style, `options.styles[${i}]`, where));
  }
  if (o.docDefaults !== undefined) {
    checkRecord(o.docDefaults, "options.docDefaults", where);
    checkRunProperties(o.docDefaults.runProperties, "options.docDefaults.runProperties", where);
    checkParagraphProperties(
      o.docDefaults.paragraphProperties,
      "options.docDefaults.paragraphProperties",
      where
    );
  }
  if (o.coreProperties !== undefined) {
    checkCoreProperties(o.coreProperties, "options.coreProperties", where);
  }
  if (o.appProperties !== undefined) {
    checkAppProperties(o.appProperties, "options.appProperties", where);
  }
  if (o.settings !== undefined) {
    checkDocumentSettings(o.settings, "options.settings", where);
  }
  if (o.watermark !== undefined) {
    checkWatermark(o.watermark, "options.watermark", where);
  }
  for (const key of [
    "abstractNumberings",
    "numberingInstances",
    "footnotes",
    "endnotes",
    "images",
    "fonts",
    "comments",
    "customProperties",
    "customXmlParts",
    "embeddedFonts",
    "opaqueParts"
  ] as const) {
    checkNodeArray(o[key], `options.${key}`, where, { optional: true });
  }
  checkRecord(o.background, "options.background", where, { optional: true });
  checkOneOf(o.missingImagePolicy, MISSING_IMAGE_POLICIES, "options.missingImagePolicy", where, {
    optional: true
  });
}

const MISSING_IMAGE_POLICIES: ReadonlySet<string> = new Set(["throw", "warn"]);

/**
 * `options` with the switches a first-page or even-page header needs to be
 * shown, as `Document.setHeader` sets them: the final section's title page
 * for a part stored under `"first"`, and `settings.evenAndOddHeaders` for one
 * under `"even"` — unless the caller stated the switch.
 */
function withHeaderSwitches(options: StreamingDocxOptions): StreamingDocxOptions {
  const has = (type: string): boolean =>
    options.headers?.has(type) === true || options.footers?.has(type) === true;
  let next = options;
  if (has("first") && options.sectionProperties?.titlePage === undefined) {
    next = { ...next, sectionProperties: { ...next.sectionProperties, titlePage: true } };
  }
  if (has("even") && options.settings?.evenAndOddHeaders === undefined) {
    next = { ...next, settings: { ...next.settings, evenAndOddHeaders: true } };
  }
  return next;
}

/**
 * Replace any image/font file names in `options` with a leaf form that's
 * safe to embed into a ZIP entry path. A new options object is returned;
 * the caller's input is not mutated.
 *
 * Names are deduplicated within their respective collection so two
 * different inputs that sanitise to the same string don't silently
 * overwrite each other in the output package.
 */
function sanitizeStreamingOptions(options: StreamingDocxOptions): StreamingDocxOptions {
  let next: StreamingDocxOptions = options;

  if (options.images && options.images.length > 0) {
    const used = new Set<string>();
    let mutated = false;
    const images = options.images.map((img, i) => {
      const safe = uniqueSanitizedName(img.fileName, used, "image.bin");
      assertWritableImage(
        { fileName: safe, mediaType: img.mediaType, contentType: img.contentType },
        "new Stream.DocumentWriter",
        `options.images[${i}]`
      );
      if (safe !== img.fileName) {
        mutated = true;
        return { ...img, fileName: safe };
      }
      return img;
    });
    if (mutated) {
      next = { ...next, images };
    }
  }
  if (options.embeddedFonts && options.embeddedFonts.length > 0) {
    const used = new Set<string>();
    let mutated = false;
    const embeddedFonts = options.embeddedFonts.map(ef => {
      const safe = uniqueSanitizedName(ef.fileName, used, "font.bin");
      if (safe !== ef.fileName) {
        mutated = true;
        return { ...ef, fileName: safe };
      }
      return ef;
    });
    if (mutated) {
      next = { ...next, embeddedFonts };
    }
  }
  return next;
}
