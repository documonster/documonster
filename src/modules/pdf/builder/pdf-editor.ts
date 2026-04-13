/**
 * PDF editor — modify existing PDF documents.
 *
 * Supports:
 * - Adding new pages with free-form content
 * - Adding text/shapes/images to existing pages (overlay)
 * - Filling form fields (AcroForm)
 * - Copying pages from other PDFs (merge)
 * - Preserving page properties (Rotate, CropBox, etc.) and metadata
 *
 * Note: save() rebuilds the PDF from scratch rather than using incremental
 * updates. This is simpler and more reliable but means object numbers change
 * and existing digital signatures will be invalidated.
 *
 * @example Edit an existing PDF:
 * ```typescript
 * import { PdfEditor } from "@cj-tech-master/excelts/pdf";
 *
 * const editor = PdfEditor.load(existingPdfBytes);
 * editor.getPage(0).drawText("APPROVED", { x: 200, y: 400, fontSize: 48, color: { r: 0, g: 0.5, b: 0 } });
 * editor.setFormField("name", "John Doe");
 * const result = await editor.save();
 * ```
 */

import { PdfDocument } from "../reader/pdf-document";
import type { PdfContentStream } from "../core/pdf-stream";
import { PdfWriter } from "../core/pdf-writer";
import { PdfDict, pdfRef, pdfString, pdfHexString, pdfNumber } from "../core/pdf-object";
import { FontManager } from "../font/font-manager";
import { parseTtf } from "../font/ttf-parser";
import { initDecryption, isEncrypted } from "../reader/pdf-decrypt";
import { extractFormFields } from "../reader/form-extractor";
import { extractMetadata } from "../reader/metadata-reader";
import type { PdfFormField } from "../reader/form-extractor";
import { isPdfArray, isPdfRef, dictGetName, decodePdfStringBytes } from "../reader/pdf-parser";
import type { PdfDictValue, PdfObject } from "../reader/pdf-parser";
import { PdfPageBuilder } from "./document-builder";
import type {
  DrawTextOptions,
  DrawRectOptions,
  DrawCircleOptions,
  DrawLineOptions,
  DrawImageOptions,
  PageOptions
} from "./document-builder";
import { PdfStructureError } from "../errors";
import { writeImageXObject } from "./image-utils";

// =============================================================================
// Types
// =============================================================================

/** Options for loading a PDF for editing. */
export interface LoadOptions {
  /** Password for encrypted PDFs. */
  password?: string;
}

// =============================================================================
// PdfEditorPage
// =============================================================================

/**
 * Proxy for an existing page that allows overlaying new content.
 * New content is drawn on top of existing content via a separate content stream.
 */
export class PdfEditorPage {
  /** @internal */
  readonly _overlay: PdfPageBuilder;
  /** @internal */
  readonly _pageIndex: number;
  /** @internal */
  readonly _width: number;
  /** @internal */
  readonly _height: number;

  /** @internal */
  constructor(pageIndex: number, width: number, height: number, fontManager: FontManager) {
    this._pageIndex = pageIndex;
    this._width = width;
    this._height = height;
    this._overlay = new PdfPageBuilder(width, height, fontManager);
  }

  /** Page width in points. */
  get width(): number {
    return this._width;
  }

  /** Page height in points. */
  get height(): number {
    return this._height;
  }

  /**
   * Draw text on this existing page (overlaid on top).
   */
  drawText(text: string, options: DrawTextOptions): this {
    this._overlay.drawText(text, options);
    return this;
  }

  /**
   * Draw a rectangle on this existing page.
   */
  drawRect(options: DrawRectOptions): this {
    this._overlay.drawRect(options);
    return this;
  }

  /**
   * Draw a circle on this existing page.
   */
  drawCircle(options: DrawCircleOptions): this {
    this._overlay.drawCircle(options);
    return this;
  }

  /**
   * Draw a line on this existing page.
   */
  drawLine(options: DrawLineOptions): this {
    this._overlay.drawLine(options);
    return this;
  }

  /**
   * Draw an image on this existing page.
   */
  drawImage(options: DrawImageOptions): this {
    this._overlay.drawImage(options);
    return this;
  }

  /**
   * Get the raw overlay content stream.
   */
  getContentStream(): PdfContentStream {
    return this._overlay._stream;
  }

  /** @internal */
  _hasOverlay(): boolean {
    return this._overlay._stream.toString().length > 0 || this._overlay._images.length > 0;
  }
}

// =============================================================================
// PdfEditor
// =============================================================================

/**
 * Editor for modifying existing PDF documents.
 *
 * Load an existing PDF, overlay content on existing pages, fill form fields,
 * add new pages, copy pages from other documents, and save.
 */
export class PdfEditor {
  private _doc: PdfDocument;
  private _pages: PdfEditorPage[] = [];
  private _newPages: PdfPageBuilder[] = [];
  private _fontManager = new FontManager();
  private _formFieldUpdates = new Map<string, string>();
  private _copiedPages: CopiedPage[] = [];
  /** @internal - Writer reference during save(), for deep-clone */
  private _writerForSave: PdfWriter | null = null;
  /** @internal - Cache of cloned indirect refs: "objNum:gen" → new objNum in writer */
  private _clonedRefs = new Map<string, number>();

  private constructor(data: Uint8Array, password: string) {
    this._doc = new PdfDocument(data);

    // Handle encryption
    if (isEncrypted(this._doc)) {
      const success = initDecryption(this._doc, password);
      if (!success) {
        throw new PdfStructureError("Failed to decrypt PDF: incorrect password");
      }
    }

    // Initialize page proxies
    const pagesInfo = this._doc.getPagesWithObjInfo();
    for (let i = 0; i < pagesInfo.length; i++) {
      const { dict } = pagesInfo[i];
      const dims = this._doc.resolvePageBox(dict) ?? { width: 612, height: 792 };
      this._pages.push(new PdfEditorPage(i, dims.width, dims.height, this._fontManager));
    }
  }

  /**
   * Load a PDF for editing.
   *
   * @param data - Raw PDF file bytes
   * @param options - Load options (e.g., password)
   * @returns A PdfEditor instance
   */
  static load(data: Uint8Array, options?: LoadOptions): PdfEditor {
    return new PdfEditor(data, options?.password ?? "");
  }

  /** Number of existing pages. */
  get pageCount(): number {
    return this._pages.length;
  }

  /**
   * Get an existing page for editing (overlaying content).
   *
   * @param index - 0-based page index
   */
  getPage(index: number): PdfEditorPage {
    if (index < 0 || index >= this._pages.length) {
      throw new PdfStructureError(`Page index ${index} out of range (0-${this._pages.length - 1})`);
    }
    return this._pages[index];
  }

  /**
   * Add a new blank page to the end of the document.
   */
  addPage(options?: PageOptions): PdfPageBuilder {
    const width = options?.width ?? 595.28;
    const height = options?.height ?? 841.89;
    const page = new PdfPageBuilder(width, height, this._fontManager);
    this._newPages.push(page);
    return page;
  }

  /**
   * Embed a TrueType font for Unicode/CJK support.
   */
  embedFont(fontBytes: Uint8Array): this {
    const ttfFont = parseTtf(fontBytes);
    this._fontManager.registerEmbeddedFont(ttfFont);
    return this;
  }

  // ===========================================================================
  // Form Fields
  // ===========================================================================

  /**
   * Set the value of a form field.
   * The field is identified by its fully qualified name (e.g., "form.address.city").
   *
   * @param fieldName - Fully qualified field name
   * @param value - New value to set
   */
  setFormField(fieldName: string, value: string): this {
    this._formFieldUpdates.set(fieldName, value);
    return this;
  }

  /**
   * Set multiple form field values at once.
   *
   * @param fields - Object mapping field names to values
   */
  setFormFields(fields: Record<string, string>): this {
    for (const [name, value] of Object.entries(fields)) {
      this._formFieldUpdates.set(name, value);
    }
    return this;
  }

  /**
   * Get current form fields (before any modifications).
   */
  getFormFields(): PdfFormField[] {
    return extractFormFields(this._doc);
  }

  // ===========================================================================
  // Page Copy / Merge
  // ===========================================================================

  /**
   * Copy pages from another PDF document into this document.
   *
   * @param sourcePdf - Raw bytes of the source PDF
   * @param pageIndices - 0-based page indices to copy. Omit to copy all pages.
   * @param options - Load options for the source PDF (e.g., password)
   */
  copyPagesFrom(sourcePdf: Uint8Array, pageIndices?: number[], options?: LoadOptions): this {
    const sourceDoc = new PdfDocument(sourcePdf);

    if (isEncrypted(sourceDoc)) {
      const success = initDecryption(sourceDoc, options?.password ?? "");
      if (!success) {
        throw new PdfStructureError("Failed to decrypt source PDF for page copy");
      }
    }

    const sourcePagesInfo = sourceDoc.getPagesWithObjInfo();
    const indices = pageIndices ?? Array.from({ length: sourcePagesInfo.length }, (_, i) => i);

    for (const idx of indices) {
      if (idx < 0 || idx >= sourcePagesInfo.length) {
        continue;
      }

      const { dict: pageDict } = sourcePagesInfo[idx];
      const dims = sourceDoc.resolvePageBox(pageDict) ?? { width: 612, height: 792 };

      // Collect all content streams from the source page
      const contentStreams = this._collectContentStreams(sourceDoc, pageDict);

      this._copiedPages.push({
        width: dims.width,
        height: dims.height,
        contentStreams,
        sourceDoc,
        sourcePageDict: pageDict
      });
    }

    return this;
  }

  // ===========================================================================
  // Save
  // ===========================================================================

  /**
   * Save the modified PDF.
   *
   * Rebuilds the PDF from scratch — content streams, resources, and page
   * properties are deep-cloned into a new document. Original metadata and
   * XMP streams are preserved. Digital signatures will be invalidated.
   *
   * @returns The modified PDF as Uint8Array
   */
  async save(): Promise<Uint8Array> {
    // Rebuild the PDF (not incremental update — simpler and more reliable)
    const writer = new PdfWriter();
    this._writerForSave = writer;
    this._clonedRefs = new Map();

    // Write font resources for any overlay content
    const fontObjectMap = await this._fontManager.writeFontResources(writer);
    const fontDictStr = this._fontManager.buildFontDictString(fontObjectMap);

    const pagesTreeObjNum = writer.allocObject();
    const pageObjNums: number[] = [];

    // Re-emit existing pages
    const pagesInfo = this._doc.getPagesWithObjInfo();
    for (let i = 0; i < pagesInfo.length; i++) {
      const { dict: pageDict } = pagesInfo[i];
      const dims = this._doc.resolvePageBox(pageDict) ?? { width: 612, height: 792 };
      const editorPage = this._pages[i];

      // Get original content streams
      const originalStreams = this._collectContentStreams(this._doc, pageDict);
      const originalResources = this._serializeResources(this._doc, pageDict);

      // Combine original + overlay content
      const allContentRefs: number[] = [];

      // Write original content streams
      for (const streamData of originalStreams) {
        const objNum = writer.allocObject();
        writer.addStreamObject(objNum, new PdfDict(), streamData);
        allContentRefs.push(objNum);
      }

      // Write overlay content stream if present
      let overlayResourcesStr = "";
      if (editorPage._hasOverlay()) {
        const overlayObjNum = writer.allocObject();
        writer.addStreamObject(overlayObjNum, new PdfDict(), editorPage._overlay._stream);
        allContentRefs.push(overlayObjNum);

        // Build overlay-specific resources (fonts, images)
        const imageXObjectMap = this._writeOverlayImages(writer, editorPage._overlay);
        overlayResourcesStr = this._buildOverlayResourcesStr(fontDictStr, imageXObjectMap);
      }

      // Write resources (merge original + overlay)
      const resourcesObjNum = writer.allocObject();
      const mergedResources = this._mergeResourceStrings(originalResources, overlayResourcesStr);
      writer.addObject(resourcesObjNum, mergedResources || "<< >>");

      // Write page dict
      const contentsStr =
        allContentRefs.length === 1
          ? pdfRef(allContentRefs[0])
          : `[${allContentRefs.map(r => pdfRef(r)).join(" ")}]`;

      // Apply form field updates to annotations
      const annotRefs = this._writeFormFieldUpdates(writer, pageDict, i);

      const pageDict2 = new PdfDict()
        .set("Type", "/Page")
        .set("Parent", pdfRef(pagesTreeObjNum))
        .set("MediaBox", `[0 0 ${pdfNumber(dims.width)} ${pdfNumber(dims.height)}]`)
        .set("Contents", contentsStr)
        .set("Resources", pdfRef(resourcesObjNum));

      // Preserve page-level properties from the original page dict
      this._copyPageProperties(pageDict, pageDict2);

      if (annotRefs.length > 0) {
        pageDict2.set("Annots", `[${annotRefs.map(r => pdfRef(r)).join(" ")}]`);
      }

      const pageObjNum = writer.allocObject();
      writer.addObject(pageObjNum, pageDict2);
      pageObjNums.push(pageObjNum);
    }

    // Write copied pages
    for (const copied of this._copiedPages) {
      const contentRefs: number[] = [];
      for (const streamData of copied.contentStreams) {
        const objNum = writer.allocObject();
        writer.addStreamObject(objNum, new PdfDict(), streamData);
        contentRefs.push(objNum);
      }

      // Deep-clone resources from the source document
      const resourcesStr = this._serializeResources(copied.sourceDoc, copied.sourcePageDict);
      const resourcesObjNum = writer.allocObject();
      writer.addObject(resourcesObjNum, resourcesStr || "<< >>");

      const contentsStr =
        contentRefs.length === 1
          ? pdfRef(contentRefs[0])
          : `[${contentRefs.map(r => pdfRef(r)).join(" ")}]`;

      const pageDict = new PdfDict()
        .set("Type", "/Page")
        .set("Parent", pdfRef(pagesTreeObjNum))
        .set("MediaBox", `[0 0 ${pdfNumber(copied.width)} ${pdfNumber(copied.height)}]`)
        .set("Contents", contentsStr)
        .set("Resources", pdfRef(resourcesObjNum));

      // Preserve page properties from source page
      this._copyPageProperties(copied.sourcePageDict, pageDict, copied.sourceDoc);

      const pageObjNum = writer.allocObject();
      writer.addObject(pageObjNum, pageDict);
      pageObjNums.push(pageObjNum);
    }

    // Write new pages
    for (const page of this._newPages) {
      const imageXObjectMap = this._writeOverlayImages(writer, page);
      let xobjDictStr = "";
      if (imageXObjectMap.size > 0) {
        const entries = [...imageXObjectMap.entries()]
          .map(([name, objNum]) => `/${name} ${pdfRef(objNum)}`)
          .join(" ");
        xobjDictStr = `<< ${entries} >>`;
      }

      const contentObjNum = writer.allocObject();
      writer.addStreamObject(contentObjNum, new PdfDict(), page._stream);

      const resourcesObjNum = writer.allocObject();
      let resStr = "<< ";
      if (fontDictStr) {
        resStr += `/Font ${fontDictStr} `;
      }
      if (xobjDictStr) {
        resStr += `/XObject ${xobjDictStr} `;
      }
      resStr += ">>";
      writer.addObject(resourcesObjNum, resStr);

      const pageObjNum = writer.addPage({
        parentRef: pagesTreeObjNum,
        width: page._width,
        height: page._height,
        contentsRef: contentObjNum,
        resourcesRef: resourcesObjNum
      });
      pageObjNums.push(pageObjNum);
    }

    // Pages tree
    const kidsStr = pageObjNums.map(n => pdfRef(n)).join(" ");
    writer.addObject(
      pagesTreeObjNum,
      new PdfDict()
        .set("Type", "/Pages")
        .set("Kids", `[${kidsStr}]`)
        .set("Count", String(pageObjNums.length))
    );

    // Catalog — with optional AcroForm
    const catalogObjNum = writer.allocObject();
    const catalogDict = new PdfDict().set("Type", "/Catalog").set("Pages", pdfRef(pagesTreeObjNum));

    // If we have form field updates, write an AcroForm dict
    if (this._formFieldUpdates.size > 0) {
      // Add NeedAppearances flag so PDF viewers regenerate field appearances
      catalogDict.set("AcroForm", "<< /NeedAppearances true >>");
    }

    // Preserve XMP metadata stream from original catalog
    try {
      const originalCatalog = this._doc.getCatalog();
      const metadataRef = originalCatalog.get("Metadata");
      if (metadataRef) {
        const clonedRef = this._deepSerialize(this._doc, metadataRef);
        if (clonedRef) {
          catalogDict.set("Metadata", clonedRef);
        }
      }
    } catch {
      // If catalog is inaccessible, skip XMP preservation
    }

    writer.addObject(catalogObjNum, catalogDict);
    writer.setCatalog(catalogObjNum);

    // Info dict — preserve original metadata
    const originalMeta = extractMetadata(this._doc);
    writer.addInfoDict({
      title: originalMeta.title || undefined,
      author: originalMeta.author || undefined,
      subject: originalMeta.subject || undefined,
      creator: originalMeta.creator || "excelts"
    });

    this._writerForSave = null;
    this._clonedRefs.clear();

    return writer.build();
  }

  // ===========================================================================
  // Internal Helpers
  // ===========================================================================

  /** Page-level keys to preserve when rebuilding page dicts. */
  private static readonly _PAGE_PRESERVE_KEYS = [
    "Rotate",
    "CropBox",
    "BleedBox",
    "TrimBox",
    "ArtBox",
    "Group",
    "UserUnit",
    "Tabs"
  ] as const;

  /** @internal - Copy preserved page properties from source to target dict. */
  private _copyPageProperties(source: PdfDictValue, target: PdfDict, doc?: PdfDocument): void {
    const resolveDoc = doc ?? this._doc;
    for (const key of PdfEditor._PAGE_PRESERVE_KEYS) {
      const val = source.get(key);
      if (val !== undefined && val !== null) {
        target.set(key, this._deepSerialize(resolveDoc, val));
      }
    }
  }

  /** @internal - Collect decoded content stream bytes from a page dict. */
  private _collectContentStreams(doc: PdfDocument, pageDict: PdfDictValue): Uint8Array[] {
    const contentsObj = pageDict.get("Contents");
    if (!contentsObj) {
      return [];
    }

    if (isPdfRef(contentsObj)) {
      const result = doc.derefStreamWithObjNum(contentsObj);
      if (result) {
        return [doc.getStreamData(result.stream, result.objNum, result.gen)];
      }
      // Could be a ref to an array
      const resolved = doc.deref(contentsObj);
      if (isPdfArray(resolved)) {
        return this._resolveStreamArray(doc, resolved);
      }
      return [];
    }

    if (isPdfArray(contentsObj)) {
      return this._resolveStreamArray(doc, contentsObj);
    }

    return [];
  }

  /** @internal */
  private _resolveStreamArray(doc: PdfDocument, arr: PdfObject[]): Uint8Array[] {
    const result: Uint8Array[] = [];
    for (const item of arr) {
      const r = doc.derefStreamWithObjNum(item);
      if (r) {
        result.push(doc.getStreamData(r.stream, r.objNum, r.gen));
      }
    }
    return result;
  }

  /** @internal - Serialize a page's Resources dict by deep-cloning objects into the writer. */
  private _serializeResources(doc: PdfDocument, pageDict: PdfDictValue): string {
    const resourcesDict = doc.resolvePageResources(pageDict);
    if (!resourcesDict || resourcesDict.size === 0) {
      return "<< >>";
    }

    // Deep-clone the resources dict into the writer.
    // We need to re-emit Font, XObject, ExtGState sub-dicts with their
    // referenced objects written as new indirect objects in the writer.
    return this._serializeDictToWriter(doc, resourcesDict);
  }

  /**
   * @internal - Recursively serialize a PdfDictValue, writing stream objects
   * as new indirect objects and converting refs to new writer refs.
   */
  private _serializeDictToWriter(doc: PdfDocument, dict: PdfDictValue): string {
    const parts: string[] = ["<<"];

    for (const [key, val] of dict.entries()) {
      const serialized = this._deepSerialize(doc, val);
      if (serialized) {
        parts.push(`/${key} ${serialized}`);
      }
    }

    parts.push(">>");
    return parts.join(" ");
  }

  /**
   * @internal - Deep-serialize a PDF value.
   * For indirect refs: resolve the target, write it as a new object in the writer,
   * and return a ref to the new object.
   * For dicts/arrays: recurse.
   */
  private _deepSerialize(doc: PdfDocument, val: PdfObject): string {
    if (val === null || val === undefined) {
      return "null";
    }
    if (typeof val === "string") {
      return val.startsWith("/") ? val : pdfString(val);
    }
    if (typeof val === "number") {
      return pdfNumber(val);
    }
    if (typeof val === "boolean") {
      return val ? "true" : "false";
    }
    if (val instanceof Uint8Array) {
      return pdfHexString(val);
    }

    if (isPdfRef(val)) {
      // Check if this ref has already been cloned
      const cacheKey = `${val.objNum}:${val.gen}`;
      const cached = this._clonedRefs.get(cacheKey);
      if (cached !== undefined) {
        return pdfRef(cached);
      }

      // Try as stream first
      const streamResult = doc.derefStreamWithObjNum(val);
      if (streamResult) {
        const newObjNum = this._writerForSave!.allocObject();
        this._clonedRefs.set(cacheKey, newObjNum);

        // Get the decoded stream data
        const streamData = doc.getStreamData(
          streamResult.stream,
          streamResult.objNum,
          streamResult.gen
        );

        // Serialize the stream's dict (excluding /Length which will be set automatically)
        const streamDict = new PdfDict();
        for (const [k, v] of streamResult.stream.dict.entries()) {
          if (k === "Length" || k === "Filter" || k === "DecodeParms") {
            // Skip — the writer will re-compress and set these
            continue;
          }
          const sv = this._deepSerialize(doc, v);
          if (sv) {
            streamDict.set(k, sv);
          }
        }

        this._writerForSave!.addStreamObject(newObjNum, streamDict, streamData);
        return pdfRef(newObjNum);
      }

      // Try as regular dict/value
      const resolved = doc.deref(val);
      if (resolved instanceof Map) {
        const newObjNum = this._writerForSave!.allocObject();
        this._clonedRefs.set(cacheKey, newObjNum);
        const dictStr = this._serializeDictToWriter(doc, resolved as PdfDictValue);
        this._writerForSave!.addObject(newObjNum, dictStr);
        return pdfRef(newObjNum);
      }

      // Primitive value behind a ref — just inline it
      return this._deepSerialize(doc, resolved);
    }

    if (isPdfArray(val)) {
      const items = val.map(item => this._deepSerialize(doc, item));
      return `[${items.join(" ")}]`;
    }

    if (val instanceof Map) {
      return this._serializeDictToWriter(doc, val as PdfDictValue);
    }

    return "";
  }

  /** @internal */
  private _writeOverlayImages(writer: PdfWriter, page: PdfPageBuilder): Map<string, number> {
    const map = new Map<string, number>();
    for (let i = 0; i < page._images.length; i++) {
      const img = page._images[i];
      const imgName = `Im${i + 1}`;
      const objNum = writeImageXObject(writer, img.data, img.format);
      map.set(imgName, objNum);
    }
    return map;
  }

  /** @internal */
  private _buildOverlayResourcesStr(
    fontDictStr: string,
    imageXObjectMap: Map<string, number>
  ): string {
    let str = "<< ";
    if (fontDictStr) {
      str += `/Font ${fontDictStr} `;
    }
    if (imageXObjectMap.size > 0) {
      const entries = [...imageXObjectMap.entries()]
        .map(([name, objNum]) => `/${name} ${pdfRef(objNum)}`)
        .join(" ");
      str += `/XObject << ${entries} >> `;
    }
    str += ">>";
    return str;
  }

  /** @internal */
  private _mergeResourceStrings(original: string, overlay: string): string {
    // If no overlay, return original
    if (!overlay || overlay === "<< >>") {
      return original;
    }
    // If no original, return overlay
    if (!original || original === "<< >>") {
      return overlay;
    }

    // Parse top-level keys from both resource dict strings and merge them.
    // For sub-dict keys (/Font, /XObject, /ExtGState) we merge the inner dicts.
    // For other keys the overlay value takes precedence.
    const origEntries = parseResourceDictEntries(original);
    const overlayEntries = parseResourceDictEntries(overlay);

    // Start with all original entries
    const merged = new Map<string, string>(origEntries);

    // Merge overlay entries
    for (const [key, overlayVal] of overlayEntries) {
      const origVal = merged.get(key);
      if (origVal && origVal.startsWith("<<") && overlayVal.startsWith("<<")) {
        // Both are sub-dicts — merge their inner entries
        const origInner = parseResourceDictEntries(origVal);
        const overlayInner = parseResourceDictEntries(overlayVal);
        for (const [ik, iv] of overlayInner) {
          origInner.set(ik, iv); // overlay wins on collision
        }
        const innerParts: string[] = ["<<"];
        for (const [ik, iv] of origInner) {
          innerParts.push(`/${ik} ${iv}`);
        }
        innerParts.push(">>");
        merged.set(key, innerParts.join(" "));
      } else {
        // Overlay value wins
        merged.set(key, overlayVal);
      }
    }

    const parts: string[] = ["<<"];
    for (const [key, val] of merged) {
      parts.push(`/${key} ${val}`);
    }
    parts.push(">>");
    return parts.join(" ");
  }

  /** @internal - Write form field value updates as annotation objects. */
  private _writeFormFieldUpdates(
    writer: PdfWriter,
    pageDict: PdfDictValue,
    _pageIndex: number
  ): number[] {
    if (this._formFieldUpdates.size === 0) {
      return this._copyExistingAnnots(writer, pageDict);
    }

    // Get existing annotations
    const annotRefs = this._copyExistingAnnots(writer, pageDict);

    // We handle form field updates by modifying Widget annotations.
    // This is done during the annotation copy above — if we find a Widget
    // annotation whose field name matches an update, we modify its /V value.
    return annotRefs;
  }

  /** @internal */
  private _copyExistingAnnots(writer: PdfWriter, pageDict: PdfDictValue): number[] {
    const annotsObj = pageDict.get("Annots");
    if (!annotsObj) {
      return [];
    }

    const annotsResolved = this._doc.deref(annotsObj);
    if (!isPdfArray(annotsResolved)) {
      return [];
    }

    const annotRefs: number[] = [];

    for (const annotRef of annotsResolved) {
      const annotDict = this._doc.derefDict(annotRef);
      if (!annotDict) {
        continue;
      }

      const subtype = dictGetName(annotDict, "Subtype");
      const objNum = writer.allocObject();

      if (subtype === "Widget" && this._formFieldUpdates.size > 0) {
        // Check if this widget has a field name that matches an update
        const fieldName = this._resolveFieldName(annotDict);
        const newValue = fieldName ? this._formFieldUpdates.get(fieldName) : undefined;

        if (newValue !== undefined) {
          // Write a modified widget annotation with the new value
          const newDict = this._buildModifiedWidgetDict(annotDict, newValue);
          writer.addObject(objNum, newDict);
          annotRefs.push(objNum);
          continue;
        }
      }

      // Copy annotation as-is (serialize what we can)
      const serialized = this._serializeAnnotDict(annotDict);
      writer.addObject(objNum, serialized);
      annotRefs.push(objNum);
    }

    return annotRefs;
  }

  /** @internal */
  private _resolveFieldName(dict: PdfDictValue): string {
    const parts: string[] = [];
    let current: PdfDictValue | null | undefined = dict;
    const visited = new Set<string>();

    while (current) {
      const tVal = current.get("T");
      if (tVal) {
        const resolved = this._doc.deref(tVal);
        let name = "";
        if (typeof resolved === "string") {
          name = resolved;
        } else if (resolved instanceof Uint8Array) {
          name = decodePdfStringBytes(resolved);
        }
        if (name) {
          parts.unshift(name);
        }
      }

      const parentVal = current.get("Parent");
      if (!parentVal) {
        break;
      }

      // Cycle guard
      const key = String(parentVal);
      if (visited.has(key)) {
        break;
      }
      visited.add(key);

      current = this._doc.derefDict(parentVal);
    }

    return parts.join(".");
  }

  /** @internal */
  private _buildModifiedWidgetDict(originalDict: PdfDictValue, newValue: string): string {
    // Build a minimal widget dict with the updated /V value
    const parts: string[] = ["<<"];

    // Copy known keys from original
    for (const [key, val] of originalDict.entries()) {
      if (key === "V" || key === "AP") {
        // Skip — we'll write our own /V and clear /AP (to force viewer to regenerate appearance)
        continue;
      }

      // Serialize the value
      const serialized = this._serializePdfValue(val);
      if (serialized) {
        parts.push(`/${key} ${serialized}`);
      }
    }

    // Write the new value
    parts.push(`/V ${pdfString(newValue)}`);

    parts.push(">>");
    return parts.join(" ");
  }

  /** @internal */
  private _serializeAnnotDict(dict: PdfDictValue): string {
    const parts: string[] = ["<<"];

    for (const [key, val] of dict.entries()) {
      const serialized = this._serializePdfValue(val);
      if (serialized) {
        parts.push(`/${key} ${serialized}`);
      }
    }

    parts.push(">>");
    return parts.join(" ");
  }

  /** @internal */
  private _serializePdfValue(val: PdfObject): string {
    if (val === null || val === undefined) {
      return "null";
    }
    if (typeof val === "string") {
      // Could be a name (starts with /) or other token
      return val.startsWith("/") ? val : pdfString(val);
    }
    if (typeof val === "number") {
      return pdfNumber(val);
    }
    if (typeof val === "boolean") {
      return val ? "true" : "false";
    }
    if (val instanceof Uint8Array) {
      return pdfHexString(val);
    }
    if (isPdfRef(val)) {
      // Deep-clone the referenced object into the new writer so the ref is valid
      if (this._writerForSave) {
        return this._deepSerialize(this._doc, val);
      }
      // Fallback if no writer context (should not happen in practice)
      return `${val.objNum} ${val.gen} R`;
    }
    if (isPdfArray(val)) {
      const items = val.map(item => this._serializePdfValue(item));
      return `[${items.join(" ")}]`;
    }
    if (val instanceof Map) {
      const parts: string[] = ["<<"];
      for (const [k, v] of val.entries()) {
        const serialized = this._serializePdfValue(v as PdfObject);
        parts.push(`/${k} ${serialized}`);
      }
      parts.push(">>");
      return parts.join(" ");
    }
    return "";
  }
}

// =============================================================================
// Resource Dict String Parser
// =============================================================================

/**
 * Parse a serialized PDF dict string (e.g. `<< /Font << /F1 3 0 R >> /XObject << ... >> >>`)
 * into a Map of top-level key → value string.
 *
 * This is a lightweight parser that handles nested `<< >>` by brace counting.
 * It doesn't need to understand every PDF token — just enough to split the
 * top-level entries so they can be merged.
 */
function parseResourceDictEntries(dictStr: string): Map<string, string> {
  const entries = new Map<string, string>();
  const trimmed = dictStr.trim();
  if (!trimmed.startsWith("<<") || !trimmed.endsWith(">>")) {
    return entries;
  }

  // Strip the outer << >>
  const inner = trimmed.slice(2, -2).trim();
  if (!inner) {
    return entries;
  }

  let i = 0;
  while (i < inner.length) {
    // Skip whitespace
    while (i < inner.length && /\s/.test(inner[i])) {
      i++;
    }
    if (i >= inner.length) {
      break;
    }

    // Expect a name: /KeyName
    if (inner[i] !== "/") {
      break; // malformed — bail
    }
    i++; // skip '/'
    const nameStart = i;
    while (i < inner.length && !/[\s/<>[\]()]/.test(inner[i])) {
      i++;
    }
    const key = inner.slice(nameStart, i);

    // Skip whitespace
    while (i < inner.length && /\s/.test(inner[i])) {
      i++;
    }

    // Read the value — could be a sub-dict <<...>>, array [...], name /..., ref N N R, etc.
    const valueStart = i;
    if (inner[i] === "<" && i + 1 < inner.length && inner[i + 1] === "<") {
      // Sub-dict: count nested << >>
      let depth = 0;
      while (i < inner.length) {
        if (inner[i] === "<" && i + 1 < inner.length && inner[i + 1] === "<") {
          depth++;
          i += 2;
        } else if (inner[i] === ">" && i + 1 < inner.length && inner[i + 1] === ">") {
          depth--;
          i += 2;
          if (depth === 0) {
            break;
          }
        } else {
          i++;
        }
      }
    } else if (inner[i] === "[") {
      // Array: find matching ]
      let depth = 0;
      while (i < inner.length) {
        if (inner[i] === "[") {
          depth++;
        } else if (inner[i] === "]") {
          depth--;
          if (depth === 0) {
            i++;
            break;
          }
        }
        i++;
      }
    } else {
      // Token(s) — read until next '/' at top level or end
      // Could be a ref like "3 0 R", a name "/Something", a number, etc.
      while (i < inner.length) {
        if (inner[i] === "/" && i > valueStart) {
          // Check if this '/' starts a new key (preceded by whitespace)
          const prevChar = inner[i - 1];
          if (/\s/.test(prevChar)) {
            break;
          }
        }
        if (inner[i] === "<" && i + 1 < inner.length && inner[i + 1] === "<") {
          break; // shouldn't happen at top level, but be safe
        }
        i++;
      }
    }

    const value = inner.slice(valueStart, i).trim();
    if (key && value) {
      entries.set(key, value);
    }
  }

  return entries;
}

// =============================================================================
// Internal Types
// =============================================================================

interface CopiedPage {
  width: number;
  height: number;
  contentStreams: Uint8Array[];
  /** Source doc + page dict — for deferred resource deep-clone during save() */
  sourceDoc: PdfDocument;
  sourcePageDict: PdfDictValue;
}
