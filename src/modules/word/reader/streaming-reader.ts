/**
 * DOCX Module - Streaming Reader
 *
 * A streaming DOCX reader that yields body content (`BodyContent`) one element
 * at a time without ever materializing the whole `document.xml` DOM. It is the
 * read-side counterpart to {@link DocumentWriter} and mirrors the Excel
 * module's streaming worksheet reader: large sequential content is streamed,
 * while small shared metadata parts (relationships, styles, numbering) are
 * loaded up-front — exactly as the Excel stream reader buffers sharedStrings /
 * styles before streaming rows.
 *
 * Memory profile:
 *   - The body model is never fully retained. As the SAX scanner reaches the
 *     close of each body-level element (`<w:p>`, `<w:tbl>`, `<w:sdt>`, …) it
 *     materializes only that single subtree into a small DOM fragment, parses
 *     it into one (or a few) `BodyContent` entries via the shared
 *     {@link parseBodyChild}, yields, then discards it. Peak per-element memory
 *     is O(largest_single_body_element).
 *   - The raw `document.xml` is *also* never fully resident: it is inflated
 *     incrementally (≈16 KiB chunks) by the unzip layer and fed to the SAX
 *     parser chunk-by-chunk. The first body element is therefore emitted after
 *     only the first inflate chunk — empirically ~0.5 % of a multi-MB part —
 *     rather than after the whole part has been decompressed. This is what
 *     makes the pipeline genuinely O(largest element) end-to-end (inflate
 *     chunk → SAX → one subtree → yield → discard). This holds on both Node
 *     (zlib streaming inflate) and the browser (native `DecompressionStream`);
 *     it is verified in a real Chromium by the streaming-reader browser suite.
 *   - Shared metadata parts (`*.rels`, `styles.xml`, `numbering.xml`) are
 *     small by construction and are buffered before body iteration begins so
 *     hyperlinks / numbering references resolve correctly. This matches the
 *     Excel stream reader, which likewise holds styles + shared strings in
 *     memory while streaming worksheet rows.
 *
 * Source requirement:
 *   The reader makes two passes over the package — one to collect the small
 *   metadata parts, one to stream `document.xml`'s body — so it requires a
 *   re-readable in-memory source (`Uint8Array` / `ArrayBuffer`). One-shot
 *   streams are not accepted because the central directory ordering of a ZIP
 *   does not guarantee that metadata parts precede `document.xml`. (The Excel
 *   stream reader solves the same ordering problem by buffering out-of-order
 *   worksheets to a temp file; here the large part is `document.xml`, so the
 *   simpler and lower-memory choice is to re-read the small parts rather than
 *   buffer the large one.)
 */

import { zip } from "@archive/create-archive";
import { unzip } from "@archive/read-archive";
import { RelType } from "@word/constants";
import { isFlatOpc, parseFlatOpc } from "@word/convert/flat-opc";
import { checkBytes, checkLimit, invalidArgument } from "@word/core/argument-checks";
import { utf8Decoder } from "@word/core/internal-utils";
import { getPartRelsPath } from "@word/core/opc-paths";
import { withRangeMarkersBefore } from "@word/core/range-markers";
import {
  DocxError,
  DocxLimitExceededError,
  DocxMissingPartError,
  DocxParseError
} from "@word/errors";
import {
  checkReadOptions,
  parseBodyChild,
  type ReadDocxOptions,
  watermarkHeadersToDrop,
  withoutHeaderReferences
} from "@word/reader/docx-reader";
import { parseNumberingXml } from "@word/reader/numbering-parser";
import { parseSectionProperties } from "@word/reader/paragraph-section-parsers";
import { attrVal, resolvePartPath, resolveRelTarget, stripWPrefix } from "@word/reader/parse-utils";
import { parseContentPartXml, unwrapCustomXml } from "@word/reader/part-xml";
import { takeRangeMarker } from "@word/reader/range-markers";
import type { ParsedRelationship } from "@word/reader/reader-context";
import { createReaderContext, parseRelationships } from "@word/reader/reader-context";
import { parseSettingsXml } from "@word/reader/settings-parser";
import { parseStyles } from "@word/reader/styles-parser";
import { findWatermark } from "@word/reader/watermark-parser";
import { decryptIfEncrypted } from "@word/security/package-encryption";
import { resolveSecurityPolicy } from "@word/security/policy";
import type { DocxSecurityPolicy } from "@word/security/policy";
import type {
  BodyContent,
  DocDefaults,
  DocumentBackground,
  NumberingInstance,
  RangeMarker,
  SectionProperties,
  StyleDef
} from "@word/types";
import { SaxParser } from "@xml/sax";
import type { XmlElement, XmlNode, XmlSaxTag } from "@xml/types";

// =============================================================================
// Public Types
// =============================================================================

/**
 * Options for {@link DocumentReader}: those of
 * `Io.read` — an encrypted package is decrypted with `password` before its body is
 * streamed, as `Io.read` decrypts it — plus the streaming limits below.
 */
export interface StreamingReadOptions extends ReadDocxOptions {
  /**
   * The maximum number of bytes a single body-level element may occupy while
   * being materialized. Guards against a pathological `document.xml` whose
   * single top-level element is unbounded (which would defeat the O(largest
   * element) memory guarantee). Defaults to the resolved security policy's
   * `maxPartSize`.
   */
  readonly maxElementBytes?: number;

  /**
   * Optional callback invoked as `document.xml` is consumed. Reports the
   * cumulative number of *uncompressed* `document.xml` bytes fed to the parser
   * so far. Useful for progress UIs and for verifying the reader's incremental
   * (true-streaming) behaviour — the first body elements are emitted long
   * before `consumedBytes` reaches the part's total size.
   */
  readonly onProgress?: (consumedBytes: number) => void;
}

/**
 * Document-level metadata loaded before body streaming begins. Exposed as
 * read-only fields on {@link DocumentReader} so consumers can resolve
 * styles / numbering while iterating body content.
 */
export interface StreamingDocxMetadata {
  /** Parsed style definitions (from `styles.xml`), if present. */
  readonly styles: readonly StyleDef[];
  /** Document default run/paragraph properties (from `styles.xml`), if present. */
  readonly docDefaults?: DocDefaults;
  /** Numbering instances (from `numbering.xml`), if present. */
  readonly numberingInstances: readonly NumberingInstance[];
  /** Document background, if present on `<w:background>`. */
  readonly background?: DocumentBackground;
}

// =============================================================================
// DocumentReader
// =============================================================================

/**
 * Streaming reader over a DOCX package. Construct with `new Stream.DocumentReader(bytes)`,
 * then consume body content with `for await (const item of reader)`.
 *
 * @example
 * ```ts
 * const reader = new Stream.DocumentReader(bytes);
 * for await (const item of reader) {
 *   if (item.type === "paragraph") { ... }
 * }
 * // metadata becomes available after the first iteration step has begun
 * const sectPr = reader.sectionProperties;
 * ```
 */
export class DocumentReader implements AsyncIterable<BodyContent> {
  private _source: Uint8Array;
  private readonly _options: StreamingReadOptions;
  private readonly _policy: Required<DocxSecurityPolicy>;

  private _started = false;
  private _consumedBytes = 0;
  private _metadata: StreamingDocxMetadata = {
    styles: [],
    numberingInstances: []
  };
  private _sectionProperties: SectionProperties | undefined;
  private _pendingMarkers: RangeMarker[] = [];
  /** The headers whose references `readDocx` drops (see `_read`). */
  private _droppedHeaders: ReadonlySet<string> = new Set();

  constructor(
    source: Uint8Array | ArrayBuffer | ArrayBufferView,
    options: StreamingReadOptions = {}
  ) {
    const where = "new Stream.DocumentReader";
    const bytes = checkBytes(source, "source", where);
    checkReadOptions(options, where);
    checkLimit(options.maxElementBytes, "options.maxElementBytes", where, { optional: true });
    if (options.onProgress !== undefined && typeof options.onProgress !== "function") {
      invalidArgument(where, "options.onProgress", "a function", options.onProgress);
    }
    this._source = bytes;
    this._options = options;
    this._policy = resolveSecurityPolicy(options.securityPolicy);
  }

  /**
   * Parsed style definitions. Populated once body iteration has begun (the
   * metadata pass runs lazily before the first yielded element).
   */
  get styles(): readonly StyleDef[] {
    return this._metadata.styles;
  }

  /**
   * Cumulative count of uncompressed `document.xml` bytes fed to the parser so
   * far. Grows as iteration proceeds; reaches the part's full size only at the
   * end. (A small value while the first elements are already yielded is direct
   * evidence of incremental, non-buffering streaming.)
   */
  get consumedBytes(): number {
    return this._consumedBytes;
  }

  /** Document default run/paragraph properties, if any. */
  get docDefaults(): DocDefaults | undefined {
    return this._metadata.docDefaults;
  }

  /** Numbering instances, if any. */
  get numberingInstances(): readonly NumberingInstance[] {
    return this._metadata.numberingInstances;
  }

  /** Document background, if any. */
  get background(): DocumentBackground | undefined {
    return this._metadata.background;
  }

  /**
   * Body-level section properties. Only populated once the body has been fully
   * consumed (the trailing `<w:sectPr>` is the last body child).
   */
  get sectionProperties(): SectionProperties | undefined {
    return this._sectionProperties;
  }

  /**
   * Range markers after the last body element (`DocxDocument.bodyRangeMarkersAfter`).
   * Only populated once the body has been fully consumed; markers between
   * elements arrive on the element that follows them.
   */
  get bodyRangeMarkersAfter(): readonly RangeMarker[] | undefined {
    return this._pendingMarkers.length > 0 ? this._pendingMarkers : undefined;
  }

  /** All metadata loaded up-front, as a single immutable record. */
  get metadata(): StreamingDocxMetadata {
    return this._metadata;
  }

  [Symbol.asyncIterator](): AsyncIterator<BodyContent> {
    if (this._started) {
      throw new DocxError("Stream.DocumentReader can only be iterated once");
    }
    this._started = true;
    return this._iterate();
  }

  /**
   * The body, failing as `Io.read` does: a damaged package — a ZIP
   * that ends early, a corrupt entry, malformed XML — is a `DocxParseError`
   * with the archive's or parser's error as its `cause`, never a bare
   * `RangeError` from deep in the reader.
   */
  private async *_iterate(): AsyncGenerator<BodyContent> {
    try {
      yield* this._read();
    } catch (e) {
      if (e instanceof DocxError) {
        throw e;
      }
      const msg = e instanceof Error ? e.message : String(e);
      throw new DocxParseError(`Failed to read DOCX: ${msg}`, { cause: e });
    }
  }

  private async *_read(): AsyncGenerator<BodyContent> {
    // The input as given is held to the package-size limit first, as `Io.read`
    // holds it.
    if (this._source.length > this._policy.maxPackageSize) {
      throw new DocxLimitExceededError(
        "packageSize",
        this._policy.maxPackageSize,
        this._source.length,
        "compressed input larger than maxPackageSize"
      );
    }
    // An encrypted package is decrypted first, as `Io.read` does: the source
    // is already whole in memory, so this costs one copy, and the decrypted
    // package then streams like any other.
    this._source = await decryptIfEncrypted(
      this._source,
      this._options.password,
      this._policy.maxPackageSize
    );
    // A Flat OPC document is the same package as one XML file, read as
    // `Io.read` reads it: as the ZIP of its parts.
    if (isFlatOpc(this._source)) {
      const archive = zip({ level: 0 });
      for (const [path, data] of parseFlatOpc(this._source)) {
        archive.add(path, data);
      }
      this._source = await archive.toBuffer();
    }

    // ---- Phase 1: collect small metadata parts (rels, styles, numbering) ----
    const { documentPartPath, relMap, watermarkHeaders, evenAndOddHeaders } =
      await this._collectMetadata();

    // ---- Phase 2: stream document.xml's body element-by-element ----
    const ctx = createReaderContext(this._policy);
    ctx.relMap = relMap;

    const documentEntryFound = { value: false };
    const maxElementBytes = this._options.maxElementBytes ?? this._policy.maxPartSize;

    // A header holding the document's watermark and nothing else is dropped
    // by `readDocx`, references and all, where the writer would give the
    // section one again. Where that is depends on every section, so the
    // section breaks are read first — only when some header holds a watermark.
    if (watermarkHeaders.size > 0) {
      for await (const entry of unzip(this._source).entries()) {
        if (normalizePath(entry.path) === documentPartPath) {
          const sections = await scanSections(entry.stream());
          this._droppedHeaders = watermarkHeadersToDrop(
            sections,
            evenAndOddHeaders,
            watermarkHeaders,
            createReaderContext(this._policy)
          );
          break;
        }
      }
    }

    const reader = unzip(this._source);
    for await (const entry of reader.entries()) {
      const path = normalizePath(entry.path);
      if (path !== documentPartPath) {
        continue;
      }
      documentEntryFound.value = true;
      yield* this._streamBody(entry.stream(), ctx, maxElementBytes);
      break;
    }

    if (!documentEntryFound.value) {
      throw new DocxMissingPartError(documentPartPath);
    }
  }

  /**
   * Phase 1 — iterate the package once, buffering only the small metadata
   * parts (`*.rels`, `styles.xml`, `numbering.xml`). `document.xml`'s bytes are
   * never buffered here.
   */
  private async _collectMetadata(): Promise<{
    documentPartPath: string;
    relMap: Map<string, ParsedRelationship>;
    /** The header parts that draw a watermark, parsed. */
    watermarkHeaders: Map<string, { readonly root: XmlElement }>;
    evenAndOddHeaders: boolean;
  }> {
    const decoder = utf8Decoder;
    const wanted = new Set<string>(["_rels/.rels"]);
    const collected = new Map<string, Uint8Array>();

    // First pass: grab the package rels to discover the document part path,
    // plus everything small. We do not know the document rels path until we
    // resolve the document part, so we buffer all `*.rels` and small XML parts.
    const reader = unzip(this._source);
    let totalUncompressed = 0;
    let entryCount = 0;
    for await (const entry of reader.entries()) {
      entryCount++;
      if (entryCount > this._policy.maxPartCount) {
        throw new DocxParseError(
          `ZIP contains more entries than maxPartCount (${this._policy.maxPartCount})`
        );
      }
      const path = normalizePath(entry.path);
      // Only buffer the small metadata parts. Skip the (potentially huge)
      // document body and all media — they are not needed to build the
      // relationship / styles / numbering context.
      const isRels = path.endsWith(".rels");
      const isStyles = path.endsWith("/styles.xml") || path === "word/styles.xml";
      const isNumbering = path.endsWith("/numbering.xml") || path === "word/numbering.xml";
      // Headers and settings decide where the watermark lives (see `_read`).
      const isHeaderOrSettings = /(?:^|\/)(?:header[^/]*|settings)\.xml$/.test(path);
      if (!isRels && !isStyles && !isNumbering && !isHeaderOrSettings && !wanted.has(path)) {
        continue;
      }
      const data = await entry.bytes();
      if (data.length > this._policy.maxPartSize) {
        throw new DocxParseError(`entry "${path}" exceeds maxPartSize`);
      }
      totalUncompressed += data.length;
      if (totalUncompressed > this._policy.maxPackageSize) {
        throw new DocxParseError("cumulative metadata size exceeds maxPackageSize");
      }
      collected.set(path, data);
    }

    const getText = (p: string): string | undefined => {
      const data = collected.get(normalizePath(p));
      return data ? decoder.decode(data) : undefined;
    };

    // Discover the document part path via package rels (supports Strict).
    let documentPartPath = "word/document.xml";
    const pkgRelsXml = getText("_rels/.rels");
    if (pkgRelsXml) {
      for (const rel of parseRelationships(pkgRelsXml)) {
        if (rel.type === RelType.OfficeDocument) {
          documentPartPath = rel.target.replace(/^\//, "");
          break;
        }
      }
    }

    const docRelsXml = getText(getPartRelsPath(documentPartPath));
    const docRels = docRelsXml ? parseRelationships(docRelsXml) : [];
    const relMap = new Map(docRels.map(r => [r.id, r]));

    // Parse styles (resolve via relationship, fallback to hardcoded path).
    const stylesPath =
      resolveRelTarget(docRels, RelType.Styles, documentPartPath) ?? "word/styles.xml";
    const stylesXml = getText(stylesPath);
    let styles: StyleDef[] = [];
    let docDefaults: DocDefaults | undefined;
    if (stylesXml) {
      try {
        const parsed = parseStyles(stylesXml);
        styles = parsed.styles;
        docDefaults = parsed.docDefaults;
      } catch {
        // Non-fatal: a broken styles part should not abort body streaming.
      }
    }

    // Parse numbering.
    const numberingPath =
      resolveRelTarget(docRels, RelType.Numbering, documentPartPath) ?? "word/numbering.xml";
    const numberingXml = getText(numberingPath);
    let numberingInstances: NumberingInstance[] = [];
    if (numberingXml) {
      try {
        const parsed = parseNumberingXml(numberingXml);
        numberingInstances = parsed.instances;
      } catch {
        // Non-fatal.
      }
    }

    this._metadata = {
      styles,
      docDefaults,
      numberingInstances
    };

    // The header parts drawing a watermark. A header part named otherwise
    // than `header*.xml` is not buffered, and is taken to draw none.
    const watermarkHeaders = new Map<string, { readonly root: XmlElement }>();
    for (const rel of docRels) {
      if (rel.type !== RelType.Header) {
        continue;
      }
      const xml = getText(resolvePartPath(documentPartPath, rel.target));
      if (!xml) {
        continue;
      }
      try {
        const root = parseContentPartXml(xml, this._policy.maxXmlDepth).root;
        if (findWatermark(root)) {
          watermarkHeaders.set(rel.id, { root });
        }
      } catch {
        // A malformed header is skipped, as `readDocx` skips it.
      }
    }
    const settingsXml = getText(
      resolveRelTarget(docRels, RelType.Settings, documentPartPath) ?? "word/settings.xml"
    );
    let evenAndOddHeaders = false;
    if (settingsXml && watermarkHeaders.size > 0) {
      try {
        evenAndOddHeaders =
          parseSettingsXml(settingsXml, this._policy.maxXmlDepth).evenAndOddHeaders === true;
      } catch {
        // Non-fatal, as in `readDocx`.
      }
    }

    return { documentPartPath, relMap, watermarkHeaders, evenAndOddHeaders };
  }

  /**
   * Phase 2 — feed `document.xml`'s byte chunks to a SAX parser, materializing
   * one body-level subtree at a time and yielding parsed `BodyContent`.
   */
  private async *_streamBody(
    chunks: AsyncIterable<Uint8Array>,
    ctx: ReturnType<typeof createReaderContext>,
    maxElementBytes: number
  ): AsyncGenerator<BodyContent> {
    const builder = new BodySubtreeBuilder(maxElementBytes);
    const parser = new SaxParser({ position: false });

    builder.attach(parser);

    const decoder = new TextDecoder("utf-8");
    let backgroundCaptured = false;

    // The SAX callbacks push completed body-level subtrees into `builder.ready`.
    // We drain that queue after each `write()` and yield parsed content.
    for await (const chunk of chunks) {
      this._consumedBytes += chunk.length;
      this._options.onProgress?.(this._consumedBytes);
      parser.write(decoder.decode(chunk, { stream: true }));

      // Capture the <w:background> element (a direct child of <w:document>,
      // sibling of <w:body>) the first time we see it close.
      if (!backgroundCaptured && builder.background) {
        this._metadata = { ...this._metadata, background: parseBackground(builder.background) };
        backgroundCaptured = true;
      }

      for (const el of builder.drain()) {
        yield* this._emit(el, ctx, builder);
      }
    }
    parser.write(decoder.decode());
    parser.close();

    if (!backgroundCaptured && builder.background) {
      this._metadata = { ...this._metadata, background: parseBackground(builder.background) };
    }

    for (const el of builder.drain()) {
      yield* this._emit(el, ctx, builder);
    }
  }

  private *_emit(
    el: XmlElement,
    ctx: ReturnType<typeof createReaderContext>,
    builder: BodySubtreeBuilder
  ): Generator<BodyContent> {
    const name = stripWPrefix(el.name);
    if (name === "sectPr") {
      this._sectionProperties = this._withHeaderReferences(parseSectionProperties(el));
      return;
    }
    if (takeRangeMarker(el, this._pendingMarkers)) {
      return;
    }
    if (builder.holdsCustomXml(el)) {
      unwrapCustomXml(el);
    }
    for (const parsed of parseBodyChild(el, ctx)) {
      const markers = this._pendingMarkers;
      if (markers.length > 0) {
        this._pendingMarkers = [];
      }
      let item = parsed;
      const sect = item.type === "paragraph" ? item.properties?.sectionProperties : undefined;
      if (item.type === "paragraph" && sect) {
        const next = this._withHeaderReferences(sect);
        if (next !== sect) {
          item = { ...item, properties: { ...item.properties, sectionProperties: next } };
        }
      }
      yield withRangeMarkersBefore(item, markers);
    }
  }

  /** `sect` without its references to the headers `readDocx` drops. */
  private _withHeaderReferences(sect: SectionProperties): SectionProperties {
    const dropped = this._droppedHeaders;
    return dropped.size > 0 && sect.headers?.some(h => dropped.has(h.rId))
      ? withoutHeaderReferences(sect, dropped)
      : sect;
  }
}

/**
 * The section properties of `document.xml` — of each paragraph's section
 * break, in document order, then the body's own — read off the SAX events
 * alone: nothing but the `w:sectPr` subtrees is built.
 */
async function scanSections(chunks: AsyncIterable<Uint8Array>): Promise<SectionProperties[]> {
  const parser = new SaxParser({ position: false });
  const decoder = new TextDecoder("utf-8");
  /** Local names of the open elements, `w:customXml` wrappers left out. */
  const path: string[] = [];
  /** The `w:sectPr` subtree being built, innermost element last. */
  let capture: XmlElement[] = [];
  let capturingFinal = false;
  const inner: SectionProperties[] = [];
  let final: SectionProperties | undefined;
  const done = (el: XmlElement): void => {
    const sect = parseSectionProperties(el);
    if (capturingFinal) {
      final = sect;
    } else {
      inner.push(sect);
    }
  };
  parser.on("opentag", tag => {
    if (capture.length > 0) {
      const el = makeElement(tag);
      capture[capture.length - 1]!.children.push(el);
      if (!tag.isSelfClosing) {
        capture.push(el);
      }
      return;
    }
    const local = stripWPrefix(tag.name);
    const at = path.join("/");
    if (local === "sectPr" && (at === "document/body" || at === "document/body/p/pPr")) {
      capturingFinal = at === "document/body";
      const el = makeElement(tag);
      if (tag.isSelfClosing) {
        done(el);
      } else {
        capture = [el];
      }
      return;
    }
    if (!tag.isSelfClosing && local !== "customXml") {
      path.push(local);
    }
  });
  parser.on("closetag", tag => {
    if (tag.isSelfClosing) {
      return;
    }
    if (capture.length > 0) {
      const el = capture.pop()!;
      if (capture.length === 0) {
        done(el);
      }
      return;
    }
    if (stripWPrefix(tag.name) !== "customXml") {
      path.pop();
    }
  });
  const text = (value: string): void => {
    capture[capture.length - 1]?.children.push({ type: "text", value });
  };
  parser.on("text", text);
  parser.on("cdata", text);
  for await (const chunk of chunks) {
    parser.write(decoder.decode(chunk, { stream: true }));
  }
  parser.write(decoder.decode());
  parser.close();
  return [...inner, final ?? {}];
}

// =============================================================================
// Internal: SAX → body-subtree materializer
// =============================================================================

/**
 * Drives a {@link SaxParser} and, whenever a *body-level* element (a direct
 * child of `<w:body>`) closes, hands the completed subtree to the consumer via
 * {@link drain}. Deeper elements are accumulated into the current top-level
 * subtree; nothing above the body level is retained.
 *
 * This is a focused, body-scoped variant of the generic DOM builder in
 * `@xml/dom` — it never retains more than the single body-level element under
 * construction, which is what gives the streaming reader its O(largest
 * element) memory bound.
 */
class BodySubtreeBuilder {
  /** Completed body-level subtrees awaiting consumption. */
  private readonly _ready: XmlElement[] = [];
  /** The captured `<w:background>` element (sibling of `<w:body>`), if seen. */
  background: XmlElement | undefined;

  /** Element stack for the subtree currently under construction. */
  private _stack: XmlElement[] = [];
  /** Depth within `<w:document>`: 0 = outside, 1 = body's children, etc. */
  private _bodyDepth = -1;
  private _inBody = false;
  private _approxBytes = 0;
  private readonly _maxElementBytes: number;
  /**
   * Open `w:customXml` elements around body blocks. They are not collected as
   * subtrees: their content is read as body children in its own right, as the
   * batch reader does (see `unwrapCustomXml`), which keeps one wrapper around
   * the whole body from becoming one element held in memory.
   */
  private _customXmlDepth = 0;
  /** Subtrees with a `w:customXml` inside, which have to be unwrapped. */
  private readonly _customXmlRoots = new WeakSet<XmlElement>();

  constructor(maxElementBytes: number) {
    this._maxElementBytes = maxElementBytes;
  }

  attach(parser: SaxParser): void {
    parser.on("opentag", tag => this._open(tag));
    parser.on("closetag", tag => this._close(tag));
    parser.on("text", text => this._text(text));
    parser.on("cdata", text => this._text(text));
  }

  /** Whether the completed subtree `root` contains a `w:customXml` element. */
  holdsCustomXml(root: XmlElement): boolean {
    return this._customXmlRoots.has(root);
  }

  /** Take and clear all completed body-level subtrees. */
  drain(): XmlElement[] {
    if (this._ready.length === 0) {
      return [];
    }
    return this._ready.splice(0, this._ready.length);
  }

  private _open(tag: XmlSaxTag): void {
    const localName = stripWPrefix(tag.name);

    if (!this._inBody) {
      // We are still scanning the document prologue (<w:document>, optional
      // <w:background>). Capture <w:background> as a small standalone subtree;
      // begin body-child accumulation once we enter <w:body>.
      if (localName === "background" && tag.name !== "w:body") {
        // Build the background subtree (it is tiny).
        this._beginSubtree(tag);
        return;
      }
      if (localName === "body") {
        this._inBody = true;
        this._bodyDepth = 0;
        return;
      }
      // Inside an already-started small subtree (e.g. <w:background> children).
      if (this._stack.length > 0) {
        this._pushChild(tag);
      }
      return;
    }

    // Inside <w:body>.
    if (this._stack.length === 0 && localName === "customXml") {
      if (!tag.isSelfClosing) {
        this._customXmlDepth++;
      }
      return;
    }
    if (this._stack.length === 0) {
      // Starting a new body-level element.
      this._approxBytes = 0;
      this._beginSubtree(tag);
    } else {
      this._pushChild(tag);
    }
  }

  private _close(tag: XmlSaxTag): void {
    const localName = stripWPrefix(tag.name);

    if (!this._inBody) {
      // Closing the <w:background> subtree.
      if (this._stack.length > 0) {
        if (tag.isSelfClosing) {
          // already handled at open
        } else {
          this._stack.pop();
        }
        if (this._stack.length === 0 && this._completedRoot) {
          this.background = this._completedRoot;
          this._completedRoot = undefined;
        }
      }
      return;
    }

    if (localName === "body") {
      this._inBody = false;
      return;
    }

    if (this._stack.length === 0) {
      if (localName === "customXml" && this._customXmlDepth > 0) {
        this._customXmlDepth--;
      }
      return;
    }

    if (!tag.isSelfClosing) {
      this._stack.pop();
    }

    if (this._stack.length === 0 && this._completedRoot) {
      this._ready.push(this._completedRoot);
      this._completedRoot = undefined;
    }
  }

  private _text(text: string): void {
    if (text.length === 0 || this._stack.length === 0) {
      return;
    }
    this._approxBytes += text.length;
    if (this._approxBytes > this._maxElementBytes) {
      throw new DocxParseError(
        `body element exceeds maxElementBytes (${this._maxElementBytes}); ` +
          "the document may be malformed or too large for streaming"
      );
    }
    const parent = this._stack[this._stack.length - 1];
    const last = parent.children[parent.children.length - 1] as XmlNode | undefined;
    if (last && last.type === "text") {
      last.value += text;
    } else {
      parent.children.push({ type: "text", value: text });
    }
  }

  /** The root element of the subtree most recently completed/under construction. */
  private _completedRoot: XmlElement | undefined;

  private _beginSubtree(tag: XmlSaxTag): void {
    const elem = makeElement(tag);
    this._completedRoot = elem;
    if (!tag.isSelfClosing) {
      this._stack = [elem];
    } else {
      // Self-closing top-level element: it is already complete.
      this._stack = [];
      if (this._inBody) {
        this._ready.push(elem);
        this._completedRoot = undefined;
      } else {
        this.background = elem;
        this._completedRoot = undefined;
      }
    }
  }

  private _pushChild(tag: XmlSaxTag): void {
    const elem = makeElement(tag);
    this._approxBytes += tag.name.length;
    if (this._approxBytes > this._maxElementBytes) {
      throw new DocxParseError(
        `body element exceeds maxElementBytes (${this._maxElementBytes}); ` +
          "the document may be malformed or too large for streaming"
      );
    }
    this._stack[this._stack.length - 1].children.push(elem);
    if (tag.name === "w:customXml" || tag.name === "customXml") {
      this._customXmlRoots.add(this._stack[0]!);
    }
    if (!tag.isSelfClosing) {
      this._stack.push(elem);
    }
  }
}

// =============================================================================
// Internal helpers
// =============================================================================

function makeElement(tag: XmlSaxTag): XmlElement {
  const attributes: Record<string, string> = Object.create(null);
  for (const key in tag.attributes) {
    if (Object.prototype.hasOwnProperty.call(tag.attributes, key)) {
      attributes[key] = tag.attributes[key];
    }
  }
  return { type: "element", name: tag.name, attributes, children: [] };
}

function normalizePath(path: string): string {
  return path.replace(/^\//, "").replace(/\\/g, "/");
}

/** Parse a captured `<w:background>` element into a {@link DocumentBackground}. */
function parseBackground(bgEl: XmlElement): DocumentBackground {
  const bg: {
    color?: string;
    themeColor?: string;
    themeShade?: string;
    themeTint?: string;
  } = {};
  const color = attrVal(bgEl, "color");
  if (color) {
    bg.color = color;
  }
  const themeColor = attrVal(bgEl, "themeColor");
  if (themeColor) {
    bg.themeColor = themeColor;
  }
  const themeShade = attrVal(bgEl, "themeShade");
  if (themeShade) {
    bg.themeShade = themeShade;
  }
  const themeTint = attrVal(bgEl, "themeTint");
  if (themeTint) {
    bg.themeTint = themeTint;
  }
  return bg;
}
