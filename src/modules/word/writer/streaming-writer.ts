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
 * Data flow:
 * ```
 * add(paragraph) → XML serialization → StreamBuf → ZipDeflate → Zip → _outputChunks
 * ```
 *
 * Memory profile, honestly stated:
 *   The body model itself is NOT retained — each element is serialized and
 *   compressed as it arrives, so the peak per-element XML/compression state
 *   is O(largest_single_element). HOWEVER, the compressed ZIP output bytes
 *   are accumulated in `_outputChunks` and only assembled in `finalize()`,
 *   which means total memory is still O(compressed_docx_size). This class
 *   does NOT yet expose an end-to-end sink (WritableStream / AsyncIterable)
 *   to the caller — callers receive the assembled `Uint8Array` from
 *   `finalize()`. Treat this writer as "streaming on the input side, buffered
 *   on the output side". Adding a true output sink is tracked separately.
 */

import { Zip, ZipDeflate } from "@archive/zip/stream";
import { XmlWriter } from "@xml/writer";

import {
  ContentType,
  RelType,
  PartPath,
  DOCUMENT_NAMESPACES,
  STD_DOC_ATTRIBUTES
} from "../constants";
import { escapeXml, utf8Encoder } from "../core/internal-utils";
import { getFileExt, getPartRelsPath } from "../core/opc-package";
import { DocxWriteError } from "../errors";
import type {
  BodyContent,
  Paragraph,
  SectionProperties,
  StyleDef,
  DocDefaults,
  AbstractNumbering,
  NumberingInstance,
  HeaderDef,
  FooterDef,
  FootnoteDef,
  EndnoteDef,
  ImageDef,
  FontDef,
  DocumentSettings,
  CoreProperties,
  AppProperties,
  CommentDef,
  DocumentBackground,
  CustomProperty,
  Watermark,
  DocumentTheme,
  CustomXmlPart,
  EmbeddedFont,
  OpaquePart
} from "../types";
import { renderComments, renderCommentsExtended } from "./comment-writer";
import { buildCommonAuxiliaryParts } from "./common-parts";
import {
  createContentTypes,
  addContentTypeDefault,
  addContentTypeOverride,
  addImageContentTypeDefaults,
  renderContentTypes
} from "./content-types";
import { renderBodyContent } from "./document-writer";
import { renderHeader, renderFooter, renderWatermarkHeader } from "./header-footer-writer";
import {
  createRelationships,
  addRelationship,
  addRelationshipWithId,
  renderRelationships
} from "./relationships";
import { renderSectionProperties } from "./section-writer";
import { StreamBuf } from "./stream-buf";
import { StringBuf } from "./string-buf";

// Per-instance StringBuf is created in the constructor (see _xmlBuffer field below).
// Previously this was a module-level singleton which caused data races with concurrent instances.

const EMPTY_U8 = new Uint8Array(0);

// =============================================================================
// Types
// =============================================================================

/** Options for the streaming DOCX writer. */
export interface StreamingDocxOptions {
  /** Compression level (0-9). Default: 6. */
  readonly compressionLevel?: number;
  /** Progress callback interval: report after every N elements. Default: 1000. */
  readonly chunkSize?: number;
  /** Section properties for the final section. */
  readonly sectionProperties?: SectionProperties;
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
 * Streaming DOCX writer. Body elements are serialized to XML and compressed
 * into the ZIP pipeline as they arrive, so the body **model** is not retained
 * after each `add()`.
 *
 * Note on memory: the compressed ZIP output is currently accumulated into an
 * internal byte chunk list and assembled into a single `Uint8Array` at
 * `finalize()` time, so peak memory is bounded by the compressed DOCX size,
 * not the input model size. End-to-end output streaming (WritableStream /
 * AsyncIterable sink) is not implemented yet.
 */
export class StreamingDocxWriter {
  private readonly _options: StreamingDocxOptions;
  private _elementCount = 0;
  private _finalized = false;
  private _onProgress?: StreamingProgressCallback;

  // Per-instance XML buffer (avoids module-level singleton data race)
  private readonly _xmlBuffer = new StringBuf({ size: 65536 });

  // ZIP infrastructure
  private _zip!: InstanceType<typeof Zip>;
  private _outputChunks: Uint8Array[] = [];
  private _documentStream!: StreamBuf;
  private _documentZipFile!: InstanceType<typeof ZipDeflate>;
  private _headerWritten = false;

  constructor(options: StreamingDocxOptions = {}) {
    this._options = options;
    this._initZip();
  }

  /** Set a progress callback. */
  onProgress(cb: StreamingProgressCallback): this {
    this._onProgress = cb;
    return this;
  }

  /**
   * Add a single body element. The element is immediately serialized to XML
   * and pushed into the ZIP compression pipeline. After this call, the element
   * can be garbage collected — it is not retained.
   */
  add(element: BodyContent): this {
    if (this._finalized) {
      throw new DocxWriteError("StreamingDocxWriter: cannot add elements after finalize()");
    }

    // Write document.xml header on first element
    if (!this._headerWritten) {
      this._writeDocumentHeader();
      this._headerWritten = true;
    }

    // Serialize this single element to XML and push to stream
    this._writeBodyElement(element);
    this._elementCount++;

    if (this._onProgress && this._elementCount % (this._options.chunkSize ?? 1000) === 0) {
      this._onProgress({ elementsWritten: this._elementCount, phase: "body" });
    }

    return this;
  }

  /** Add multiple body elements at once. */
  addMany(elements: readonly BodyContent[]): this {
    for (const el of elements) {
      this.add(el);
    }
    return this;
  }

  /** Add a paragraph with simple text content. */
  addText(content: string, properties?: Paragraph["properties"]): this {
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
   * Finalize the document and return the DOCX bytes.
   * Writes the document.xml footer, then adds all auxiliary parts
   * (styles, numbering, headers, footers, etc.) and closes the ZIP.
   */
  async finalize(): Promise<Uint8Array> {
    if (this._finalized) {
      throw new DocxWriteError("StreamingDocxWriter: already finalized");
    }
    this._finalized = true;

    if (this._onProgress) {
      this._onProgress({ elementsWritten: this._elementCount, phase: "finalizing" });
    }

    // If no elements were added, still write a minimal document
    if (!this._headerWritten) {
      this._writeDocumentHeader();
    }

    // Write document.xml footer (close </w:body></w:document>)
    this._writeDocumentFooter();

    // End the document.xml stream → finalizes its ZIP entry
    await this._endStream(this._documentStream);

    // Add all auxiliary parts (styles, settings, etc.)
    await this._addAuxiliaryParts();

    // Finalize the ZIP archive
    this._zip.end();

    // Assemble output
    return this._assembleOutput();
  }

  /** Reset the writer for reuse. */
  reset(): this {
    this._elementCount = 0;
    this._finalized = false;
    this._headerWritten = false;
    this._outputChunks = [];
    this._initZip();
    return this;
  }

  // ===========================================================================
  // Private: ZIP infrastructure
  // ===========================================================================

  private _initZip(): void {
    this._outputChunks = [];
    this._zip = new Zip((_err, data, _final) => {
      if (data && data.length > 0) {
        this._outputChunks.push(data);
      }
    });

    // Create the document.xml ZIP entry and stream
    const level = this._options.compressionLevel ?? 6;
    this._documentZipFile = new ZipDeflate(PartPath.Document, { level });
    this._zip.add(this._documentZipFile);

    this._documentStream = new StreamBuf({ bufSize: 65536 });
    this._documentStream.on("data", (chunk: Uint8Array) => {
      this._documentZipFile.push(chunk);
    });
    this._documentStream.once("finish", () => {
      this._documentZipFile.push(EMPTY_U8, true);
      this._documentStream.emit("zipped");
    });
  }

  private _write(text: string): void {
    this._xmlBuffer.reset();
    this._xmlBuffer.addText(text);
    this._documentStream.write(this._xmlBuffer);
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
      header += `<w:background w:color="${escapeXml(bg.color ?? "FFFFFF")}"`;
      if (bg.themeColor) {
        header += ` w:themeColor="${escapeXml(bg.themeColor)}"`;
      }
      header += `/>`;
    }

    header += `<w:body>`;
    this._write(header);
  }

  private _writeBodyElement(element: BodyContent): void {
    // Serialize a single body element using the shared renderBodyContent function.
    // This produces only the element XML (e.g. <w:p>...</w:p>) without wrapper tags.
    const writer = new XmlWriter();
    renderBodyContent(writer, element);
    this._write(writer.xml);
  }

  private _writeDocumentFooter(): void {
    // Write final section properties if provided (using the full section writer)
    if (this._options.sectionProperties) {
      const writer = new XmlWriter();
      renderSectionProperties(writer, this._options.sectionProperties);
      this._write(writer.xml);
    }

    // Close body and document
    this._write(`</w:body></w:document>`);
  }

  // ===========================================================================
  // Private: Auxiliary parts
  // ===========================================================================

  private async _addAuxiliaryParts(): Promise<void> {
    const level = this._options.compressionLevel ?? 6;

    // Helper: add a complete XML file to the ZIP
    const addXmlFile = (path: string, renderFn: (xml: XmlWriter) => void): void => {
      const writer = new XmlWriter();
      renderFn(writer);
      const data = utf8Encoder.encode(writer.xml);
      const file = new ZipDeflate(path, { level });
      this._zip.add(file);
      file.push(data, true);
    };

    // Content types and relationships
    const contentTypes = createContentTypes();
    const packageRels = createRelationships();
    const documentRels = createRelationships();

    // Package relationships
    addRelationship(packageRels, RelType.OfficeDocument, "word/document.xml");
    addRelationship(packageRels, RelType.CoreProperties, "docProps/core.xml");
    addRelationship(packageRels, RelType.ExtendedProperties, "docProps/app.xml");

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

    // Footnotes
    if (this._options.footnotes && this._options.footnotes.length > 0) {
      addRelationship(documentRels, RelType.Footnotes, "footnotes.xml");
      addContentTypeOverride(contentTypes, `/${PartPath.Footnotes}`, ContentType.Footnotes);
      // XML rendering handled by buildCommonAuxiliaryParts below
    }

    // Endnotes
    if (this._options.endnotes && this._options.endnotes.length > 0) {
      addRelationship(documentRels, RelType.Endnotes, "endnotes.xml");
      addContentTypeOverride(contentTypes, `/${PartPath.Endnotes}`, ContentType.Endnotes);
      // XML rendering handled by buildCommonAuxiliaryParts below
    }

    // Comments
    if (this._options.comments && this._options.comments.length > 0) {
      addRelationship(documentRels, RelType.Comments, "comments.xml");
      addContentTypeOverride(contentTypes, `/${PartPath.Comments}`, ContentType.Comments);
      addXmlFile(PartPath.Comments, xml => renderComments(xml, this._options.comments!));
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
    if (this._options.headers) {
      let headerIdx = 1;
      for (const [, headerDef] of this._options.headers) {
        const headerPath = PartPath.header(headerIdx);
        addRelationship(documentRels, RelType.Header, `header${headerIdx}.xml`);
        addContentTypeOverride(contentTypes, `/${headerPath}`, ContentType.Header);
        addXmlFile(headerPath, xml => renderHeader(xml, headerDef.content));
        headerIdx++;
      }
    }

    // Watermark (rendered as a special header if no headers already handle it)
    if (this._options.watermark && !this._options.headers) {
      const watermarkPath = PartPath.header(1);
      addRelationship(documentRels, RelType.Header, "header1.xml");
      addContentTypeOverride(contentTypes, `/${watermarkPath}`, ContentType.Header);
      addXmlFile(watermarkPath, xml => renderWatermarkHeader(xml, this._options.watermark!));
    }

    // Footers
    if (this._options.footers) {
      let footerIdx = 1;
      for (const [, footerDef] of this._options.footers) {
        const footerPath = PartPath.footer(footerIdx);
        addRelationship(documentRels, RelType.Footer, `footer${footerIdx}.xml`);
        addContentTypeOverride(contentTypes, `/${footerPath}`, ContentType.Footer);
        addXmlFile(footerPath, xml => renderFooter(xml, footerDef.content));
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

    // Images
    if (this._options.images) {
      const extensions = new Set<string>();
      for (const img of this._options.images) {
        addRelationship(documentRels, RelType.Image, `media/${img.fileName}`);
        const ext = getFileExt(img.fileName);
        if (ext) {
          extensions.add(ext);
        }
      }
      addImageContentTypeDefaults(contentTypes, extensions);
    }

    // Custom XML parts (for SDT data binding)
    if (this._options.customXmlParts && this._options.customXmlParts.length > 0) {
      this._options.customXmlParts.forEach((part, i) => {
        const num = i + 1;
        const itemPath = `word/customXml/item${num}.xml`;
        const propsPath = `word/customXml/itemProps${num}.xml`;

        // Write the XML content
        const itemData = utf8Encoder.encode(part.xmlContent);
        const itemFile = new ZipDeflate(itemPath, { level });
        this._zip.add(itemFile);
        itemFile.push(itemData, true);

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
        const propsFile = new ZipDeflate(propsPath, { level });
        this._zip.add(propsFile);
        propsFile.push(propsData, true);

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
        const fontFile = new ZipDeflate(partPath, { level: 0 });
        this._zip.add(fontFile);
        fontFile.push(ef.data, true);

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
      for (const part of this._options.opaqueParts) {
        const opaqueFile = new ZipDeflate(part.path, { level });
        this._zip.add(opaqueFile);
        opaqueFile.push(part.data, true);

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
      docDefaults: this._options.docDefaults,
      styles: this._options.styles,
      settings: this._options.settings,
      fonts: this._options.fonts,
      theme: this._options.theme,
      abstractNumberings: this._options.abstractNumberings,
      numberingInstances: this._options.numberingInstances,
      footnotes: this._options.footnotes,
      endnotes: this._options.endnotes,
      coreProperties: this._options.coreProperties,
      appProperties: this._options.appProperties,
      customProperties: this._options.customProperties
    });

    for (const part of commonParts) {
      const data = utf8Encoder.encode(part.content);
      const file = new ZipDeflate(part.path, { level });
      this._zip.add(file);
      file.push(data, true);
    }

    // Write images
    if (this._options.images) {
      for (const img of this._options.images) {
        const file = new ZipDeflate(PartPath.media(img.fileName), { level: 0 });
        this._zip.add(file);
        file.push(img.data, true);
      }
    }

    // Write document.xml.rels
    addXmlFile(PartPath.DocumentRels, xml => renderRelationships(documentRels, xml));

    // Write _rels/.rels
    addXmlFile(PartPath.PackageRels, xml => renderRelationships(packageRels, xml));

    // Write [Content_Types].xml
    addXmlFile(PartPath.ContentTypes, xml => renderContentTypes(contentTypes, xml));
  }

  private _endStream(stream: StreamBuf): Promise<void> {
    return new Promise(resolve => {
      stream.once("zipped", () => resolve());
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
 * Create a new streaming DOCX writer.
 *
 * @example
 * ```ts
 * const writer = createDocxStream({
 *   styles: [{ type: "paragraph", styleId: "Normal", name: "Normal" }]
 * });
 *
 * for (let i = 0; i < 100000; i++) {
 *   writer.addText(`Paragraph ${i}`);
 * }
 *
 * const buffer = await writer.finalize();
 * ```
 */
export function createDocxStream(options?: StreamingDocxOptions): StreamingDocxWriter {
  return new StreamingDocxWriter(options);
}
