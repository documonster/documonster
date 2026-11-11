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

import { createZip } from "@archive/create-archive";
import { unzipZip } from "@archive/read-archive";
import type { SchemaMatches } from "@utils/schema";
import { createXmlPartLeadTrimmer } from "@utils/xml-part";
import { RelType } from "@word/constants";
import { isFlatOpc, parseFlatOpc } from "@word/convert/flat-opc";
import { checkBytes } from "@word/core/argument-checks";
import { utf8Decoder } from "@word/core/internal-utils";
import { getPartRelsPath } from "@word/core/opc-paths";
import { checkSchema, w } from "@word/core/option-schema";
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
import { WORD_NAMESPACES } from "@word/reader/namespaces";
import { parseNumberingXml } from "@word/reader/numbering-parser";
import { createPartBudget, readOptionsShape } from "@word/reader/package-limits";
import { parseSectionProperties } from "@word/reader/paragraph-section-parsers";
import {
  attrVal,
  isWordDocumentRoot,
  resolvePartPath,
  resolveRelTarget,
  stripWPrefix
} from "@word/reader/parse-utils";
import { asPartLimitError, parseContentPartXml, unwrapCustomXml } from "@word/reader/part-xml";
import { bindSubtreeNamespaces, foreignDeclarations } from "@word/reader/preserved-xml";
import { takeRangeMarker } from "@word/reader/range-markers";
import type { ParsedRelationship } from "@word/reader/reader-context";
import { createReaderContext, parseRelationships } from "@word/reader/reader-context";
import { parseSettingsXml } from "@word/reader/settings-parser";
import { parseStyles } from "@word/reader/styles-parser";
import { findWatermark } from "@word/reader/watermark-parser";
import { decryptIfEncrypted } from "@word/security/package-encryption";
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
import { createPrefixNormalizer } from "@xml/namespace-prefixes";
import { SaxParser } from "@xml/sax";
import type { XmlElement, XmlNode, XmlSaxTag } from "@xml/types";
import { utf8Length } from "@xml/utf8-decoder";

// =============================================================================
// Public Types
// =============================================================================

/**
 * Options for {@link DocumentReader}: those of
 * `Document.read` — an encrypted package is decrypted with `password` before its body is
 * streamed, as `Document.read` decrypts it — plus the streaming limits below.
 */
export interface StreamingReadOptions extends ReadDocxOptions {
  /**
   * The most a single body-level element may hold while being materialized: the UTF-8 bytes of its
   * element names, attribute names and values, and text. Guards against a pathological `document.xml`
   * whose single top-level element is unbounded (which would defeat the O(largest element) memory
   * guarantee). Past it the read fails with `DocxLimitExceededError` (`limit: "elementBytes"`).
   * Defaults to the resolved security policy's `maxPartSize`.
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

const streamingReadOptionsSchema = w.record({
  ...readOptionsShape,
  maxElementBytes: w.limit().optional(),
  onProgress: w.function<(consumedBytes: number) => void>().optional()
});
const _streamingReadOptionsMatch: SchemaMatches<
  typeof streamingReadOptionsSchema,
  StreamingReadOptions
> = true;
void _streamingReadOptionsMatch;

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
  /** The package's parts counted against the policy: see {@link _collectMetadata}. */
  private readonly _budget: ReturnType<typeof createPartBudget>;
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
    checkSchema(streamingReadOptionsSchema, options, where);
    this._policy = checkReadOptions(options, where);
    this._source = bytes;
    this._options = options;
    this._budget = createPartBudget(this._policy);
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
   * The body, failing as `Document.read` does: a damaged package — a ZIP
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
    // The input as given is held to the package-size limit first, as `Document.read`
    // holds it.
    if (this._source.length > this._policy.maxPackageSize) {
      throw new DocxLimitExceededError(
        "packageSize",
        this._policy.maxPackageSize,
        this._source.length,
        "compressed input larger than maxPackageSize"
      );
    }
    // An encrypted package is decrypted first, as `Document.read` does: the source
    // is already whole in memory, so this costs one copy, and the decrypted
    // package then streams like any other.
    this._source = await decryptIfEncrypted(
      this._source,
      this._options.password,
      this._policy.maxPackageSize
    );
    // A Flat OPC document is the same package as one XML file, read as
    // `Document.read` reads it: as the ZIP of its parts.
    if (isFlatOpc(this._source)) {
      const archive = createZip({ level: 0 });
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
      for await (const entry of this._unzip().entries()) {
        if (normalizePath(entry.path) === documentPartPath) {
          const sections = await scanSections(entry.stream(), this._policy, maxElementBytes);
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

    const reader = this._unzip();
    for await (const entry of reader.entries()) {
      const path = normalizePath(entry.path);
      if (path !== documentPartPath) {
        continue;
      }
      documentEntryFound.value = true;
      yield* this._streamBody(path, entry.stream(), ctx, maxElementBytes);
      break;
    }

    if (!documentEntryFound.value) {
      throw new DocxMissingPartError(documentPartPath);
    }
  }

  /**
   * The package's entries, read as `Document.read` reads them: checked against their CRC-32 unless
   * `checkCrc32` is `false`, and none inflated past `maxPartSize`.
   */
  private _unzip(): ReturnType<typeof unzipZip> {
    return unzipZip(this._source, {
      checkCrc32: this._options.checkCrc32,
      maxEntrySize: this._policy.maxPartSize
    });
  }

  /**
   * Phase 1 — iterate the package once, buffering only the small metadata
   * parts (`*.rels`, `styles.xml`, `numbering.xml`). `document.xml`'s bytes are
   * never buffered here.
   *
   * Every part counts against the policy as `Document.read` counts it: one read is charged what it
   * inflates to, and one this reader never inflates — media, the body before it streams — what its
   * entry declares, so a package `Document.read` refuses is refused here too.
   */
  private async _collectMetadata(): Promise<{
    documentPartPath: string;
    relMap: Map<string, ParsedRelationship>;
    /** The header parts that draw a watermark, parsed. */
    watermarkHeaders: Map<string, { readonly root: XmlElement }>;
    evenAndOddHeaders: boolean;
  }> {
    const decoder = utf8Decoder;
    const collected = new Map<string, Uint8Array>();
    /** The parts not read, by their declared uncompressed size. */
    const unread = new Map<string, number>();

    // First pass: grab the package rels to discover the document part path,
    // plus everything small. We do not know the document rels path until we
    // resolve the document part, so we buffer all `*.rels` and small XML parts.
    for await (const entry of this._unzip().entries()) {
      // A folder entry is not a part (OPC §9.1.1.1), as `Document.read` skips it.
      if (entry.type === "directory" || /[/\\]$/.test(entry.path)) {
        continue;
      }
      this._budget.count();
      const path = normalizePath(entry.path);
      // Only buffer the small metadata parts. Skip the (potentially huge)
      // document body and all media — they are not needed to build the
      // relationship / styles / numbering context.
      const isRels = path.endsWith(".rels");
      const isStyles = path.endsWith("/styles.xml") || path === "word/styles.xml";
      const isNumbering = path.endsWith("/numbering.xml") || path === "word/numbering.xml";
      // Headers and settings decide where the watermark lives (see `_read`).
      const isHeaderOrSettings = /(?:^|\/)(?:header[^/]*|settings)\.xml$/.test(path);
      if (!isRels && !isStyles && !isNumbering && !isHeaderOrSettings) {
        unread.set(path, entry.info.uncompressedSize);
        continue;
      }
      collected.set(path, await this._budget.read(path, entry));
    }

    const getText = (p: string): string | undefined => {
      const data = collected.get(normalizePath(p));
      return data ? decoder.decode(data) : undefined;
    };

    // Discover the document part path via package rels (supports Strict).
    const depth = this._policy.maxXmlDepth;
    // Best effort for a damaged auxiliary part, as `Document.read` reads it — but a resource limit is
    // a verdict on the package, not a damaged part.
    const tryParse = <T>(parse: () => T): T | undefined => {
      try {
        return parse();
      } catch (error) {
        if (error instanceof DocxLimitExceededError) {
          throw error;
        }
        return undefined;
      }
    };
    let documentPartPath = "word/document.xml";
    const pkgRelsXml = getText("_rels/.rels");
    if (pkgRelsXml) {
      for (const rel of parseRelationships(pkgRelsXml, depth)) {
        if (rel.type === RelType.OfficeDocument) {
          documentPartPath = rel.target.replace(/^\//, "");
          break;
        }
      }
    }

    // The body is charged as it streams; every other part not read, by what it declares.
    for (const [path, size] of unread) {
      if (path !== documentPartPath) {
        this._budget.charge(path, size);
      }
    }

    const docRelsXml = getText(getPartRelsPath(documentPartPath));
    const docRels = docRelsXml ? parseRelationships(docRelsXml, depth) : [];
    const relMap = new Map(docRels.map(r => [r.id, r]));

    // Parse styles (resolve via relationship, fallback to hardcoded path).
    const stylesPath =
      resolveRelTarget(docRels, RelType.Styles, documentPartPath) ?? "word/styles.xml";
    const stylesXml = getText(stylesPath);
    let styles: StyleDef[] = [];
    let docDefaults: DocDefaults | undefined;
    if (stylesXml) {
      // Non-fatal: a broken styles part should not abort body streaming.
      const parsed = tryParse(() => parseStyles(stylesXml, depth));
      styles = parsed?.styles ?? [];
      docDefaults = parsed?.docDefaults;
    }

    // Parse numbering.
    const numberingPath =
      resolveRelTarget(docRels, RelType.Numbering, documentPartPath) ?? "word/numbering.xml";
    const numberingXml = getText(numberingPath);
    let numberingInstances: NumberingInstance[] = [];
    if (numberingXml) {
      numberingInstances = tryParse(() => parseNumberingXml(numberingXml, depth))?.instances ?? [];
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
      // A malformed header is skipped, as `readDocx` skips it.
      const root = tryParse(() => parseContentPartXml(xml, depth).root);
      if (root !== undefined && findWatermark(root)) {
        watermarkHeaders.set(rel.id, { root });
      }
    }
    const settingsXml = getText(
      resolveRelTarget(docRels, RelType.Settings, documentPartPath) ?? "word/settings.xml"
    );
    let evenAndOddHeaders = false;
    if (settingsXml && watermarkHeaders.size > 0) {
      // Non-fatal, as in `readDocx`.
      evenAndOddHeaders =
        tryParse(() => parseSettingsXml(settingsXml, depth))?.evenAndOddHeaders === true;
    }

    return { documentPartPath, relMap, watermarkHeaders, evenAndOddHeaders };
  }

  /**
   * Phase 2 — feed `document.xml`'s byte chunks to a SAX parser, materializing
   * one body-level subtree at a time and yielding parsed `BodyContent`.
   */
  private async *_streamBody(
    documentPartPath: string,
    chunks: AsyncIterable<Uint8Array>,
    ctx: ReturnType<typeof createReaderContext>,
    maxElementBytes: number
  ): AsyncGenerator<BodyContent> {
    const builder = new BodySubtreeBuilder(maxElementBytes);
    const part = new StreamedPart(this._policy);

    builder.attach(part);

    let backgroundCaptured = false;

    // The SAX callbacks push completed body-level subtrees into `builder.ready`.
    // We drain that queue after each `write()` and yield parsed content.
    for await (const chunk of chunks) {
      this._consumedBytes += chunk.length;
      // The body is held to the limits `Document.read` holds every part to as it is decompressed.
      this._budget.charge(documentPartPath, chunk.length);
      this._options.onProgress?.(this._consumedBytes);
      part.write(chunk);

      // Capture the <w:background> element (a direct child of <w:document>,
      // sibling of <w:body>) the first time we see it close.
      if (!backgroundCaptured && builder.background) {
        this._metadata = {
          ...this._metadata,
          background: parseBackground(builder.background)
        };
        backgroundCaptured = true;
      }

      for (const el of builder.drain()) {
        yield* this._emit(el, ctx, builder);
      }
    }
    part.close();
    if (!builder.sawBody) {
      // As `Document.read` reads it: a main part without a body is not a document — a part whose
      // elements are in another namespace than Word's has none either.
      throw new DocxParseError("Missing w:body element in document.xml");
    }

    if (!backgroundCaptured && builder.background) {
      this._metadata = {
        ...this._metadata,
        background: parseBackground(builder.background)
      };
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
    if (builder.needsBinding(el)) {
      bindSubtreeNamespaces(el, builder.outerNamespaces);
    }
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
          item = {
            ...item,
            properties: { ...item.properties, sectionProperties: next }
          };
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
async function scanSections(
  chunks: AsyncIterable<Uint8Array>,
  policy: Required<DocxSecurityPolicy>,
  maxElementBytes: number
): Promise<SectionProperties[]> {
  const part = new StreamedPart(policy);
  /** Names of the open elements, `w:customXml` wrappers left out. */
  const path: string[] = [];
  /** Whether the root is Word's `w:document`: only its body has sections. */
  let isDocument = false;
  let sawRoot = false;
  /** The `w:sectPr` subtree being built, innermost element last. */
  let capture: XmlElement[] = [];
  /** Its bytes, held to `maxElementBytes` as the body's elements are. */
  let captured = 0;
  const charge = (bytes: number): void => {
    captured += bytes;
    if (captured > maxElementBytes) {
      throw new DocxLimitExceededError(
        "elementBytes",
        maxElementBytes,
        captured,
        "a body element exceeds maxElementBytes"
      );
    }
  };
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
  const onOpen = (tag: XmlSaxTag): void => {
    if (capture.length > 0) {
      charge(tagBytes(tag));
      const el = makeElement(tag);
      capture[capture.length - 1]!.children.push(el);
      if (!tag.isSelfClosing) {
        capture.push(el);
      }
      return;
    }
    if (!sawRoot) {
      sawRoot = true;
      isDocument = isWordDocumentRoot(tag);
    }
    const at = path.join("/");
    if (
      isDocument &&
      tag.name === "w:sectPr" &&
      (at === "w:document/w:body" || at === "w:document/w:body/w:p/w:pPr")
    ) {
      capturingFinal = at === "w:document/w:body";
      captured = 0;
      charge(tagBytes(tag));
      const el = makeElement(tag);
      if (tag.isSelfClosing) {
        done(el);
      } else {
        capture = [el];
      }
      return;
    }
    if (!tag.isSelfClosing && tag.name !== "w:customXml") {
      path.push(tag.name);
    }
  };
  const onClose = (tag: XmlSaxTag): void => {
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
    if (tag.name !== "w:customXml") {
      path.pop();
    }
  };
  const text = (value: string): void => {
    if (capture.length > 0) {
      charge(utf8Length(value, 0, value.length));
      capture[capture.length - 1]!.children.push({ type: "text", value });
    }
  };
  part.on({ opentag: onOpen, closetag: onClose, text });
  for await (const chunk of chunks) {
    part.write(chunk);
  }
  part.close();
  return [...inner, final ?? {}];
}

// =============================================================================
// Internal: a part parsed as it streams
// =============================================================================

/**
 * A package part fed to a SAX parser as its bytes arrive, held to what `Document.read` holds a part
 * to: decoded as it decodes one, read from its first `<`, at most `maxPartSize` bytes, and nested no
 * deeper than `maxXmlDepth`, each limit reported as a `DocxLimitExceededError`.
 */
class StreamedPart {
  private readonly parser: SaxParser;
  private readonly decoder = new TextDecoder("utf-8");
  private readonly trimLead = createXmlPartLeadTrimmer();
  /** Renames each tag to the reader's prefixes, as `parsePartXml` does a part it holds whole. */
  // Lazy: dormant — each tag passed as it is — while the part declares only the reader's prefixes, as
  // every part Word writes does; what a renamed tag is in text is asked for where one is.
  private readonly prefixes = createPrefixNormalizer(WORD_NAMESPACES, {
    lazy: true
  });
  private bytes = 0;

  constructor(private readonly policy: Required<DocxSecurityPolicy>) {
    this.parser = new SaxParser({
      position: false,
      maxDepth: policy.maxXmlDepth
    });
  }

  /**
   * Handle the part's events: its tags named with the reader's prefixes, whatever prefixes the part
   * binds (see `namespaces.ts`).
   */
  on(handlers: {
    opentag: (tag: XmlSaxTag) => void;
    closetag: (tag: XmlSaxTag) => void;
    text: (text: string) => void;
  }): void {
    const { parser, prefixes } = this;
    // The text view of each tag, as `parsePartXml` parses a part it rewrote: what a subtree written
    // back declares is what it uses.
    parser.on("opentag", raw => {
      const named = prefixes.open(raw);
      const tag = prefixes.textOpened() ?? named;
      handlers.opentag(
        tag === raw
          ? raw
          : {
              ...tag,
              attributes: tag.attributes,
              isSelfClosing: raw.isSelfClosing
            }
      );
    });
    parser.on("closetag", raw => {
      const named = prefixes.close(raw.name);
      const name = prefixes.textClosed() ?? named;
      handlers.closetag(name === raw.name ? raw : { ...raw, name });
    });
    parser.on("text", handlers.text);
    parser.on("cdata", handlers.text);
  }

  write(chunk: Uint8Array): void {
    this.bytes += chunk.length;
    if (this.bytes > this.policy.maxPartSize) {
      throw new DocxLimitExceededError(
        "partSize",
        this.policy.maxPartSize,
        this.bytes,
        "the part exceeds maxPartSize"
      );
    }
    this.parse(this.trimLead(this.decoder.decode(chunk, { stream: true })));
  }

  close(): void {
    this.parse(this.trimLead(this.decoder.decode()));
    try {
      this.parser.close();
    } catch (error) {
      throw asPartLimitError(error, this.policy.maxXmlDepth);
    }
  }

  private parse(text: string): void {
    // Shown to the normaliser first, so it wakes before a tag that needs it is reported.
    this.prefixes.scan(text);
    try {
      this.parser.write(text);
    } catch (error) {
      throw asPartLimitError(error, this.policy.maxXmlDepth);
    }
  }
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
  /** Whether the part's `<w:body>` has been read. */
  sawBody = false;

  /** Element stack for the subtree currently under construction. */
  private _stack: XmlElement[] = [];
  /** Depth within `<w:document>`: 0 = outside, 1 = body's children, etc. */
  private _bodyDepth = -1;
  private _inBody = false;
  /** Depth of the elements open outside the body: 0 before the root, 1 inside it, … */
  private _outerDepth = 0;
  /** Whether the root is Word's `w:document`. */
  private _inDocument = false;
  /** The UTF-8 bytes of the names, attributes and text of the body-level element being built. */
  /**
   * The UTF-16 units charged to the element being built, and — once they could be past the limit —
   * its UTF-8 bytes, counted exactly from then on (`-1` before). A unit is one to three bytes, so
   * while three times the units are within the limit the bytes are too, and nothing is counted: a
   * document whose elements are far below the limit, as nearly every one is, pays a length per string.
   */
  private _units = 0;
  private _bytes = -1;
  private readonly _maxElementBytes: number;

  /** Count `bytes` more of the element being built against `maxElementBytes`. */
  /** Start counting a new element. */
  private _resetCharge(): void {
    this._units = 0;
    this._bytes = -1;
  }

  /** Charge `text` — not yet in the element being built — to it. */
  private _chargeText(text: string): void {
    if (this._bytes < 0) {
      this._units += text.length;
      if (this._units * 3 <= this._maxElementBytes) {
        return;
      }
      this._bytes = this._builtBytes();
    }
    this._addBytes(utf8Length(text, 0, text.length));
  }

  /**
   * Charge `tag`'s name and attributes — not yet in the element being built — to it, straight after
   * {@link makeElement} made its element and counted its units.
   */
  private _chargeTag(tag: XmlSaxTag): void {
    if (this._bytes < 0) {
      this._units += madeUnits;
      if (this._units * 3 <= this._maxElementBytes) {
        return;
      }
      this._bytes = this._builtBytes();
    }
    this._addBytes(tagBytes(tag));
  }

  private _addBytes(bytes: number): void {
    this._bytes += bytes;
    if (this._bytes > this._maxElementBytes) {
      throw new DocxLimitExceededError(
        "elementBytes",
        this._maxElementBytes,
        this._bytes,
        "a body element exceeds maxElementBytes"
      );
    }
  }

  /** The UTF-8 bytes of what the element being built holds so far: what has been charged to it. */
  private _builtBytes(): number {
    const root = this._stack[0];
    return root === undefined ? 0 : elementBytes(root);
  }
  /**
   * Open `w:customXml` elements around body blocks. They are not collected as
   * subtrees: their content is read as body children in its own right, as the
   * batch reader does (see `unwrapCustomXml`), which keeps one wrapper around
   * the whole body from becoming one element held in memory.
   */
  private _customXmlDepth = 0;
  /**
   * The declarations of other vocabularies' namespaces in scope around the body's elements — made
   * on `w:document` and `w:body` — which a subtree kept verbatim must carry: see {@link needsBinding}.
   */
  outerNamespaces: Record<string, string> = {};
  /** Subtrees declaring such a namespace inside them. */
  private readonly _declaringRoots = new WeakSet<XmlElement>();

  /** Whether `root`'s elements must be given the declarations they inherit (`bindSubtreeNamespaces`). */
  needsBinding(root: XmlElement): boolean {
    return Object.keys(this.outerNamespaces).length > 0 || this._declaringRoots.has(root);
  }

  private _scopeAround(tag: XmlSaxTag): void {
    const declared = foreignDeclarations(tag.attributes);
    if (declared) {
      Object.assign(this.outerNamespaces, declared);
    }
  }

  /** Subtrees with a `w:customXml` inside, which have to be unwrapped. */
  private readonly _customXmlRoots = new WeakSet<XmlElement>();

  constructor(maxElementBytes: number) {
    this._maxElementBytes = maxElementBytes;
  }

  attach(part: StreamedPart): void {
    part.on({
      opentag: tag => this._open(tag),
      closetag: tag => this._close(tag),
      text: text => this._text(text)
    });
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
      // Inside an already-started small subtree (<w:background>'s children).
      if (this._stack.length > 0) {
        this._pushChild(tag);
        return;
      }
      const depth = this._outerDepth;
      if (!tag.isSelfClosing) {
        this._outerDepth++;
      }
      if (depth === 0) {
        // The root: the body is only read under a `w:document` in Word's namespace.
        this._inDocument = isWordDocumentRoot(tag);
        this._scopeAround(tag);
        return;
      }
      if (depth !== 1 || !this._inDocument) {
        return;
      }
      // A child of <w:document>: the optional <w:background>, captured as a small standalone
      // subtree, and <w:body>, whose children are accumulated from here on.
      if (tag.name === "w:background") {
        this._resetCharge();
        this._beginSubtree(tag);
        return;
      }
      if (tag.name === "w:body" && !tag.isSelfClosing) {
        this._scopeAround(tag);
        this._inBody = true;
        this.sawBody = true;
        this._bodyDepth = 0;
      } else if (tag.name === "w:body") {
        this.sawBody = true;
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
      // Starting a new body-level element: its own tag counts toward it, as its children's do.
      this._resetCharge();
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
          this._outerDepth--;
        }
        return;
      }
      if (!tag.isSelfClosing) {
        this._outerDepth--;
      }
      return;
    }

    // The body's own end tag: at the body level, with no body-level element open.
    if (this._stack.length === 0 && this._customXmlDepth === 0 && localName === "body") {
      this._inBody = false;
      this._outerDepth--;
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
    this._chargeText(text);
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
    // Its own tag counts toward it, as its children's do.
    this._chargeTag(tag);
    if (madeDeclares && foreignDeclarations(tag.attributes)) {
      this._declaringRoots.add(elem);
    }
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
    const declares = madeDeclares;
    this._chargeTag(tag);
    this._stack[this._stack.length - 1].children.push(elem);
    if (declares && this._inBody && foreignDeclarations(tag.attributes)) {
      this._declaringRoots.add(this._stack[0]!);
    }
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

/** The UTF-8 bytes of `element`'s names, attributes and text, as `maxElementBytes` counts them. */
function elementBytes(element: XmlElement): number {
  let bytes = 0;
  const pending: XmlNode[] = [element];
  while (pending.length > 0) {
    const node = pending.pop()!;
    if (node.type === "text") {
      bytes += utf8Length(node.value, 0, node.value.length);
    } else if (node.type === "element") {
      bytes += utf8Length(node.name, 0, node.name.length);
      for (const key in node.attributes) {
        const value = node.attributes[key];
        bytes += utf8Length(key, 0, key.length) + utf8Length(value, 0, value.length);
      }
      for (const child of node.children) {
        pending.push(child);
      }
    }
  }
  return bytes;
}

/** The UTF-8 bytes of a tag's name and attributes, as `maxElementBytes` counts them. */
function tagBytes(tag: XmlSaxTag): number {
  let bytes = utf8Length(tag.name, 0, tag.name.length);
  for (const key in tag.attributes) {
    const value = tag.attributes[key];
    bytes += utf8Length(key, 0, key.length) + utf8Length(value, 0, value.length);
  }
  return bytes;
}

/**
 * What {@link makeElement} saw while copying a tag's attributes, read straight after it: the UTF-16
 * units of the name and attributes (for `maxElementBytes`), and whether one is a namespace
 * declaration — one pass over the attributes, not one for each question.
 */
let madeUnits = 0;
let madeDeclares = false;

function makeElement(tag: XmlSaxTag): XmlElement {
  const attributes: Record<string, string> = Object.create(null);
  let units = tag.name.length;
  let declares = false;
  for (const key in tag.attributes) {
    if (Object.prototype.hasOwnProperty.call(tag.attributes, key)) {
      const value = tag.attributes[key];
      attributes[key] = value;
      units += key.length + value.length;
      if (key.charCodeAt(0) === 0x78 && (key === "xmlns" || key.startsWith("xmlns:"))) {
        declares = true;
      }
    }
  }
  madeUnits = units;
  madeDeclares = declares;
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
