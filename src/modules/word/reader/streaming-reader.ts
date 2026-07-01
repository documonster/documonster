/**
 * DOCX Module - Streaming Reader
 *
 * A streaming DOCX reader that yields body content (`BodyContent`) one element
 * at a time without ever materializing the whole `document.xml` DOM. It is the
 * read-side counterpart to {@link StreamingDocxWriter} and mirrors the Excel
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

import { unzip } from "@archive/read-archive";
import { RelType } from "@word/constants";
import { utf8Decoder } from "@word/core/internal-utils";
import { getPartRelsPath } from "@word/core/opc-paths";
import { DocxError, DocxMissingPartError, DocxParseError } from "@word/errors";
import { parseBodyChild, type ReadDocxOptions } from "@word/reader/docx-reader";
import { parseNumberingXml } from "@word/reader/numbering-parser";
import { parseSectionProperties } from "@word/reader/paragraph-section-parsers";
import { attrVal, resolveRelTarget } from "@word/reader/parse-utils";
import type { ParsedRelationship } from "@word/reader/reader-context";
import { createReaderContext, parseRelationships } from "@word/reader/reader-context";
import { parseStyles } from "@word/reader/styles-parser";
import { resolveSecurityPolicy } from "@word/security/policy";
import type { WordSecurityPolicy } from "@word/security/policy";
import type {
  BodyContent,
  DocDefaults,
  DocumentBackground,
  NumberingInstance,
  SectionProperties,
  StyleDef
} from "@word/types";
import { SaxParser } from "@xml/sax";
import type { XmlElement, XmlNode, SaxTag } from "@xml/types";

// =============================================================================
// Public Types
// =============================================================================

/** Options for {@link createDocxStreamReader} / {@link StreamingDocxReader}. */
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
 * read-only fields on {@link StreamingDocxReader} so consumers can resolve
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
// StreamingDocxReader
// =============================================================================

/**
 * Streaming reader over a DOCX package. Construct via {@link createDocxStreamReader},
 * then consume body content with `for await (const item of reader)`.
 *
 * @example
 * ```ts
 * const reader = createDocxStreamReader(bytes);
 * for await (const item of reader) {
 *   if (item.type === "paragraph") { ... }
 * }
 * // metadata becomes available after the first iteration step has begun
 * const sectPr = reader.sectionProperties;
 * ```
 */
export class StreamingDocxReader implements AsyncIterable<BodyContent> {
  private readonly _source: Uint8Array;
  private readonly _options: StreamingReadOptions;
  private readonly _policy: Required<WordSecurityPolicy>;

  private _started = false;
  private _consumedBytes = 0;
  private _metadata: StreamingDocxMetadata = {
    styles: [],
    numberingInstances: []
  };
  private _sectionProperties: SectionProperties | undefined;

  constructor(source: Uint8Array, options: StreamingReadOptions = {}) {
    this._source = source;
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

  /** All metadata loaded up-front, as a single immutable record. */
  get metadata(): StreamingDocxMetadata {
    return this._metadata;
  }

  [Symbol.asyncIterator](): AsyncIterator<BodyContent> {
    if (this._started) {
      throw new DocxError("StreamingDocxReader can only be iterated once");
    }
    this._started = true;
    return this._iterate();
  }

  private async *_iterate(): AsyncGenerator<BodyContent> {
    // CFB-encrypted packages cannot be streamed (the whole container must be
    // decrypted first). Surface a clear error rather than failing deep in the
    // ZIP parser.
    if (isCfb(this._source)) {
      if (this._options.password == null) {
        throw new DocxError(
          "Encrypted DOCX cannot be read with the streaming reader. Use readDocx() with a password."
        );
      }
      throw new DocxError(
        "Encrypted DOCX is not supported by the streaming reader. Use readDocx() to decrypt first."
      );
    }

    if (this._source.length > this._policy.maxPackageSize) {
      throw new DocxParseError(
        `compressed input larger than maxPackageSize (${this._policy.maxPackageSize})`
      );
    }

    // ---- Phase 1: collect small metadata parts (rels, styles, numbering) ----
    const { documentPartPath, relMap } = await this._collectMetadata();

    // ---- Phase 2: stream document.xml's body element-by-element ----
    const ctx = createReaderContext(this._policy);
    ctx.relMap = relMap;

    const documentEntryFound = { value: false };
    const maxElementBytes = this._options.maxElementBytes ?? this._policy.maxPartSize;

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
      if (!isRels && !isStyles && !isNumbering && !wanted.has(path)) {
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

    return { documentPartPath, relMap };
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
        yield* this._emit(el, ctx);
      }
    }
    parser.write(decoder.decode());
    parser.close();

    if (!backgroundCaptured && builder.background) {
      this._metadata = { ...this._metadata, background: parseBackground(builder.background) };
    }

    for (const el of builder.drain()) {
      yield* this._emit(el, ctx);
    }
  }

  private *_emit(
    el: XmlElement,
    ctx: ReturnType<typeof createReaderContext>
  ): Generator<BodyContent> {
    const name = el.name.replace(/^w:/, "");
    if (name === "sectPr") {
      this._sectionProperties = parseSectionProperties(el);
      return;
    }
    for (const item of parseBodyChild(el, ctx)) {
      yield item;
    }
  }
}

/**
 * Create a streaming DOCX reader over an in-memory package.
 *
 * @param source - The DOCX package bytes (a re-readable in-memory buffer).
 * @param options - Read + streaming options.
 */
export function createDocxStreamReader(
  source: Uint8Array,
  options?: StreamingReadOptions
): StreamingDocxReader {
  return new StreamingDocxReader(source, options);
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

  constructor(maxElementBytes: number) {
    this._maxElementBytes = maxElementBytes;
  }

  attach(parser: SaxParser): void {
    parser.on("opentag", tag => this._open(tag));
    parser.on("closetag", tag => this._close(tag));
    parser.on("text", text => this._text(text));
    parser.on("cdata", text => this._text(text));
  }

  /** Take and clear all completed body-level subtrees. */
  drain(): XmlElement[] {
    if (this._ready.length === 0) {
      return [];
    }
    return this._ready.splice(0, this._ready.length);
  }

  private _open(tag: SaxTag): void {
    const localName = tag.name.replace(/^w:/, "");

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
    if (this._stack.length === 0) {
      // Starting a new body-level element.
      this._approxBytes = 0;
      this._beginSubtree(tag);
    } else {
      this._pushChild(tag);
    }
  }

  private _close(tag: SaxTag): void {
    const localName = tag.name.replace(/^w:/, "");

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

  private _beginSubtree(tag: SaxTag): void {
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

  private _pushChild(tag: SaxTag): void {
    const elem = makeElement(tag);
    this._approxBytes += tag.name.length;
    if (this._approxBytes > this._maxElementBytes) {
      throw new DocxParseError(
        `body element exceeds maxElementBytes (${this._maxElementBytes}); ` +
          "the document may be malformed or too large for streaming"
      );
    }
    this._stack[this._stack.length - 1].children.push(elem);
    if (!tag.isSelfClosing) {
      this._stack.push(elem);
    }
  }
}

// =============================================================================
// Internal helpers
// =============================================================================

function makeElement(tag: SaxTag): XmlElement {
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

function isCfb(buffer: Uint8Array): boolean {
  return (
    buffer.length >= 8 &&
    buffer[0] === 0xd0 &&
    buffer[1] === 0xcf &&
    buffer[2] === 0x11 &&
    buffer[3] === 0xe0 &&
    buffer[4] === 0xa1 &&
    buffer[5] === 0xb1 &&
    buffer[6] === 0x1a &&
    buffer[7] === 0xe1
  );
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
