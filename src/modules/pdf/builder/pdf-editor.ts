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
 * import { Pdf } from "documonster/pdf";
 *
 * const editor = Pdf.Editor.read(existingPdfBytes);
 * editor.getPage(0).drawText("APPROVED", { x: 200, y: 400, fontSize: 48, color: { r: 0, g: 0.5, b: 0 } });
 * editor.setFormField("name", "John Doe");
 * const result = await editor.toBuffer();
 * ```
 */

import {
  buildCatalogIndex,
  keepsStructureTree,
  planThreads,
  rewriteCatalogEntries
} from "@pdf/builder/catalog-rewrite";
import type { CatalogIndex, ThreadPlan } from "@pdf/builder/catalog-rewrite";
import {
  PdfPageBuilder,
  checkPageOptions,
  checkSignatureOptions,
  checkUniqueFieldNames,
  snapshotPage,
  writeBuilderAnnotations,
  writeFormFieldObjects
} from "@pdf/builder/document-builder";
import type {
  PdfPageSnapshot,
  DrawTextOptions,
  DrawRectOptions,
  DrawCircleOptions,
  DrawEllipseOptions,
  DrawLineOptions,
  DrawImageOptions,
  DrawPathOptions,
  DrawSvgOptions,
  PathOp,
  PageOptions,
  AnnotationOptions,
  DocumentMetadata,
  FormFieldOptions,
  PdfSignatureOptions
} from "@pdf/builder/document-builder";
import { discoverDrawnTextFonts, withDrawnTextFontBuild } from "@pdf/builder/drawn-text-fonts";
import type { DrawnTextFontSettings } from "@pdf/builder/drawn-text-fonts";
import { generateTextFieldAppearance, buildAppearanceBBox } from "@pdf/builder/form-appearance";
import { addImageXObject, writeImageXObject } from "@pdf/builder/image-utils";
import type { ImageObjectSink } from "@pdf/builder/image-utils";
import { buildPageIndex, pruneRemovedPages, pruneToPage } from "@pdf/builder/page-index";
import type { PageIndex, Pruning } from "@pdf/builder/page-index";
import {
  parseResourceDict,
  mergeResourceDicts,
  serializeResourceDict
} from "@pdf/builder/resource-merger";
import type { PdfResourceDict } from "@pdf/builder/resource-merger";
import {
  at,
  checkFontOptionNames,
  checkMetadataUpdate,
  checkOptionalString,
  checkString,
  fontOptionsShape,
  pdfBytes,
  readFontSourceArgument
} from "@pdf/core/arguments";
import {
  PdfDict,
  pdfDate,
  pdfName,
  pdfNameBytes,
  pdfRef,
  pdfString,
  pdfHexString,
  pdfNumber
} from "@pdf/core/pdf-object";
import type { PdfContentStream } from "@pdf/core/pdf-stream";
import { PdfWriter, buildIncremental } from "@pdf/core/pdf-writer";
import type { IncrementalObject, PdfWriterSecurityHandler } from "@pdf/core/pdf-writer";
import { decodeXmpPacket, encodeXmpPacket, updateXmpPacket } from "@pdf/core/xmp";
import {
  PdfArgumentError,
  PdfEncryptionError,
  PdfIncrementalUpdateError,
  PdfPasswordError,
  PdfSignatureInvalidationError,
  PdfStructureError,
  PdfXfaFormError
} from "@pdf/errors";
import { compilePdfFontConfig } from "@pdf/font/font-config";
import type { PdfFontConfig, PdfFontSource } from "@pdf/font/font-config";
import { FontManager } from "@pdf/font/font-manager";
import { parseTtf } from "@pdf/font/ttf-parser";
import { extractFormFields } from "@pdf/reader/form-extractor";
import type { PdfFormField } from "@pdf/reader/form-extractor";
import { initDecryption, isEncrypted } from "@pdf/reader/pdf-decrypt";
import { PdfDocument } from "@pdf/reader/pdf-document";
import { isPdfArray, isPdfRef, dictGetName, decodePdfStringBytes } from "@pdf/reader/pdf-parser";
import type { PdfDictValue, PdfObject, PdfRef } from "@pdf/reader/pdf-parser";
import { getDictStringValue } from "@pdf/reader/reader-utils";
import { alphaGsName } from "@pdf/render/page-renderer";
import type { CjkLanguage } from "@utils/cjk";
import { describeValue } from "@utils/describe-value";
import { isRecordArgument } from "@utils/object";
import { check, s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";

// =============================================================================
// Types
// =============================================================================

/**
 * Options for {@link PdfEditor.read}: the password, and the font options of
 * `Pdf.Builder` — the same names, types, checks and behaviour — for the text drawn on the
 * document's pages (`Pdf.EditorPage#drawText`, `Pdf.Editor#addPage`).
 *
 * Without them, text the standard PDF faces cannot encode (Chinese, Japanese, Korean, most
 * scripts beyond Latin-1) is drawn with a system face found on the host, as the builder
 * does; `onWarning` reports that choice and anything no face could draw.
 */
export interface EditorReadOptions {
  /** Password for encrypted PDFs. */
  password?: string;
  /**
   * Embedded typefaces for drawn text, as `Pdf.Builder`'s `fonts`. With this set,
   * {@link PdfEditor.embedFont} is refused: the configuration already names the faces.
   */
  readonly fonts?: PdfFontConfig;
  /**
   * System font families auto-discovery should prefer, in order — see
   * `Pdf.Builder#preferSystemFonts`. No effect in a browser, which has no system fonts to
   * search.
   */
  readonly preferSystemFonts?: readonly string[];
  /**
   * The East Asian written language of the drawn text — see `Pdf.Builder#setTextLanguage`.
   * No effect in a browser, which has no system fonts to search.
   */
  readonly textLanguage?: CjkLanguage;
  /**
   * Turn off the host font scan, for output that does not depend on the fonts installed —
   * see `Pdf.Builder#disableFontAutoDiscovery`. No effect in a browser, which has no such
   * scan.
   */
  readonly disableFontAutoDiscovery?: boolean;
  /**
   * Receive non-fatal font diagnostics raised by each write: a face auto-embedded from the
   * host, characters no face could draw (drawn as `.notdef` boxes), font family names it did
   * not recognise — see `Pdf.Builder#onWarning`.
   */
  readonly onWarning?: (message: string) => void;
}

/** Options every write shares — {@link PdfEditor.toBuffer} and {@link PdfEditor.sign}. */
export interface EditorWriteOptions {
  /**
   * Permit output that breaks a digital signature already in the document.
   * Without it, an operation that would do so throws
   * `PdfSignatureInvalidationError` instead of silently producing a file
   * whose signatures no longer verify.
   * @default false
   */
  invalidateSignatures?: boolean;
  /**
   * What a full rewrite of an **encrypted** document does with its encryption.
   *
   * - `"keep"`: encrypt the output again with the document's own security handler — the
   *   same algorithm, key, `/Encrypt` dictionary and `/ID` — so the user and owner
   *   passwords that open the original open the output, with the same permissions.
   *   Which of the two passwords the document was read with makes no difference.
   * - `"remove"`: write the document decrypted, without `/Encrypt`. An incremental update
   *   cannot remove encryption — the original encrypted bytes stay in the file — so with
   *   `incremental: true` this rewrites the whole file instead, and with
   *   `incremental: "required"` it throws `PdfIncrementalUpdateError`.
   *
   * An incremental update always keeps the encryption: its appended objects are encrypted
   * with the document's handler. Ignored for a document that is not encrypted.
   *
   * The deprecated `save()` — and `saveIncremental()` when it falls back to a full rewrite —
   * default to `"remove"`, the decrypted copy they always wrote.
   * @default "keep"
   */
  encryption?: "keep" | "remove";
}

/** Options for {@link PdfEditor.toBuffer}. */
export interface EditorToBufferOptions extends EditorWriteOptions {
  /**
   * Append an incremental update instead of rewriting the file. An update leaves the
   * original bytes — and so any signature over them — intact.
   *
   * - `true`: append when possible, otherwise fall back to a full rewrite — when the
   *   edits are structural (pages added, removed, copied or rotated), the
   *   cross-reference data had to be recovered, or a page has no object number. On a
   *   signed document that fallback throws `PdfSignatureInvalidationError` unless
   *   `invalidateSignatures` is set. Whether it happened can be read off the result: an
   *   appended file starts with the original bytes.
   * - `"required"`: never fall back; throw `PdfIncrementalUpdateError` instead, naming
   *   the reason.
   *
   * With no edits at all, either form returns the original bytes unchanged.
   * @default false
   */
  incremental?: boolean | "required";
}

/** Options for {@link PdfEditor.copyPagesFrom}. */
export interface EditorCopyPagesOptions {
  /** Password of the source PDF, when it is encrypted. */
  password?: string;
  /**
   * 0-based indices of the source's pages to copy, in the order given. Omit to copy all
   * pages. Each must name a page of the source: an index outside it, or one that is not an
   * integer, is refused with `PdfArgumentError` (field `options.pageIndices[i]`) before anything
   * is copied. (The deprecated positional `copyPagesFrom(source, indices)` keeps skipping an
   * index outside the source.)
   *
   * Named `pageIndices`, like every 0-based page list of the editor, and not `pages`: that is
   * the name of `Pdf.extract`'s list of **1-based** page numbers.
   */
  pageIndices?: readonly number[];
  /**
   * What to do with the form fields on the copied pages:
   *
   * - `"copy"` (default) — copy them: their widgets go onto the copied page and the fields
   *   join this document's form. A field whose fully qualified name this document already
   *   uses (or another copied or added field uses) is refused by the call, since one name
   *   identifies one field.
   * - `"drop"` — copy the pages without their form fields: widget annotations are left out
   *   and no field is added, so no name can clash. Use it to merge the same form twice, or
   *   two forms that share field names. The page content and every other annotation (links,
   *   comments, highlights, …) are still copied — they are what `"drop"` keeps from what
   *   `"copy"` writes, minus the fields. (Up to 0.21, a copied page lost all of its
   *   annotations, fields included.)
   */
  formFields?: "copy" | "drop";
}

/**
 * The argument of {@link PdfEditor.setMetadata}: the fields of `DocumentMetadata`, each of
 * which may also be `null`. It is a JSON merge patch (RFC 7396) of the document properties:
 *
 * - absent or `undefined` — keep what the document has;
 * - a value — set it;
 * - `null` — remove it: the `/Info` entry and the XMP property (in whatever form the packet
 *   holds it — element, attribute, or an `rdf:Alt`/`rdf:Seq` container with every entry).
 *
 * `Pdf.Builder#setMetadata` takes `DocumentMetadata`: a new document has nothing to remove.
 */
export type PdfMetadataUpdate = {
  readonly [K in keyof DocumentMetadata]?: DocumentMetadata[K] | null;
};

/** A checked {@link PdfMetadataUpdate}, copied down to its dates, its `null`s kept. */
function copyMetadataUpdate(update: PdfMetadataUpdate): PdfMetadataUpdate {
  const copy: { -readonly [K in keyof PdfMetadataUpdate]: PdfMetadataUpdate[K] } = {};
  for (const field of ["title", "author", "subject", "keywords", "creator", "producer"] as const) {
    if (update[field] !== undefined) {
      copy[field] = update[field];
    }
  }
  for (const field of ["creationDate", "modDate"] as const) {
    const value = update[field];
    if (value !== undefined) {
      copy[field] = value === null ? null : new Date(value.getTime());
    }
  }
  return copy;
}

/** Options for {@link PdfEditor.splitPages}. */
export interface EditorSplitOptions extends EditorWriteOptions {
  /**
   * 0-based indices of the pages to write, one part each, in the order given — the pages the
   * document was read with, and the pages added by {@link PdfEditor.addPage} or
   * {@link PdfEditor.copyPagesFrom} numbered after them (the first one added is
   * {@link PdfEditor.pageCount}). Omit to split every page that is not removed. Each index must
   * name a page that is not removed: one out of range, naming a removed page, or not an integer
   * is refused with `PdfArgumentError` (field `options.pageIndices[i]`) before any part is
   * written. (The deprecated positional `splitPages(indices)` keeps skipping an index out of
   * range or naming a removed page.)
   */
  pageIndices?: readonly number[];
}

/** Options for {@link PdfEditor.sign}. */
export interface EditorSignOptions extends PdfSignatureOptions, EditorWriteOptions {}

// =============================================================================
// PdfEditorPage
// =============================================================================

/** @internal - What a page proxy asks its editor. */
interface EditorPageOwner {
  isPageRemoved(index: number): boolean;
  /** Throws when `name` is already a field of the document, or of a field added to it. */
  checkFieldName(name: string): void;
}

/**
 * Proxy for an existing page that allows overlaying new content.
 * New content is drawn on top of existing content via a separate content stream.
 *
 * Once its page is removed (`Pdf.Editor#removePage`), every method that would add to the page
 * — the `draw*` calls, `addAnnotation`, `addFormField`, `getContentStream` — throws
 * `PdfArgumentError` (field `page`); content drawn on a removed page used to be accepted and
 * written nowhere. `measureText` and the size getters still answer.
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

  // Private, and not `@internal`, for the reason `PdfPageBuilder`'s is: a page proxy comes from
  // `Pdf.Editor#getPage`, and a stripped constructor published a public `new Pdf.EditorPage()`.
  /** @internal - The editor this page belongs to. */
  private readonly _owner: EditorPageOwner;

  private constructor(
    pageIndex: number,
    width: number,
    height: number,
    fontManager: FontManager,
    resourcePrefix: string,
    owner: EditorPageOwner,
    overlay: PdfPageBuilder | undefined
  ) {
    this._pageIndex = pageIndex;
    this._width = width;
    this._height = height;
    this._owner = owner;
    this._overlay =
      overlay ??
      PdfPageBuilder._create(
        width,
        height,
        fontManager,
        resourcePrefix,
        name => owner.checkFieldName(name),
        "Pdf.EditorPage"
      );
  }

  /**
   * @internal - Made by the editor: one per page of the document as it was read, and one per
   * added page on first request. A page added with `addPage` passes its own builder as
   * `overlay`, so the handle draws onto that page; a page read or copied in gets an overlay of
   * its own, drawn over its content.
   */
  static _create(
    pageIndex: number,
    width: number,
    height: number,
    fontManager: FontManager,
    resourcePrefix: string,
    owner: EditorPageOwner,
    overlay?: PdfPageBuilder
  ): PdfEditorPage {
    return new PdfEditorPage(pageIndex, width, height, fontManager, resourcePrefix, owner, overlay);
  }

  /**
   * @internal - Throw unless the page is still in the document. Drawing on a removed page
   * used to be accepted and then written nowhere.
   */
  private _live(where: string): PdfPageBuilder {
    if (this._owner.isPageRemoved(this._pageIndex)) {
      throw new PdfArgumentError(
        "page",
        `${where}: page ${this._pageIndex} has been removed (removePage), so nothing drawn on it would be written`
      );
    }
    return this._overlay;
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
   * Measure the width of a text string in points.
   */
  measureText(
    text: string,
    options?: {
      fontSize?: number;
      fontFamily?: string;
      bold?: boolean;
      italic?: boolean;
    }
  ): number {
    return this._overlay.measureText(text, options);
  }

  /**
   * Draw text on this existing page (overlaid on top).
   */
  drawText(text: string, options: DrawTextOptions): this {
    this._live("Pdf.EditorPage#drawText").drawText(text, options);
    return this;
  }

  /**
   * Draw a rectangle on this existing page.
   */
  drawRect(options: DrawRectOptions): this {
    this._live("Pdf.EditorPage#drawRect").drawRect(options);
    return this;
  }

  /**
   * Draw a circle on this existing page.
   */
  drawCircle(options: DrawCircleOptions): this {
    this._live("Pdf.EditorPage#drawCircle").drawCircle(options);
    return this;
  }

  /**
   * Draw an ellipse on this existing page.
   */
  drawEllipse(options: DrawEllipseOptions): this {
    this._live("Pdf.EditorPage#drawEllipse").drawEllipse(options);
    return this;
  }

  /**
   * Draw a line on this existing page.
   */
  drawLine(options: DrawLineOptions): this {
    this._live("Pdf.EditorPage#drawLine").drawLine(options);
    return this;
  }

  /**
   * Draw an image on this existing page.
   */
  drawImage(options: DrawImageOptions): this {
    this._live("Pdf.EditorPage#drawImage").drawImage(options);
    return this;
  }

  /**
   * Get the raw overlay content stream.
   */
  getContentStream(): PdfContentStream {
    return this._live("Pdf.EditorPage#getContentStream")._stream;
  }

  /**
   * Add an annotation to this existing page (Highlight, Text, FreeText, Stamp, etc.).
   */
  addAnnotation(options: AnnotationOptions): this {
    this._live("Pdf.EditorPage#addAnnotation").addAnnotation(options);
    return this;
  }

  /**
   * Add a form field to this existing page.
   */
  addFormField(options: FormFieldOptions): this {
    this._live("Pdf.EditorPage#addFormField").addFormField(options);
    return this;
  }

  /**
   * Draw an SVG path on this existing page.
   */
  drawSvgPath(d: string, options?: DrawPathOptions): this {
    this._live("Pdf.EditorPage#drawSvgPath").drawSvgPath(d, options);
    return this;
  }

  /**
   * Draw a complex path from a list of path operations.
   */
  drawPath(ops: readonly PathOp[], options?: DrawPathOptions): this {
    this._live("Pdf.EditorPage#drawPath").drawPath(ops, options);
    return this;
  }

  /**
   * Draw a simple SVG document on this existing page — the same subset as
   * `Pdf.PageBuilder#drawSvg`.
   */
  drawSvg(options: DrawSvgOptions): this {
    this._live("Pdf.EditorPage#drawSvg").drawSvg(options);
    return this;
  }
}

/**
 * @internal - The fully qualified names of the form fields with a widget on `pageDict`: each
 * widget's `/T` and its ancestors', joined with `.` (a widget with no `/T` of its own is a
 * kid of its field and adds nothing).
 */
function widgetFieldNames(doc: PdfDocument, pageDict: PdfDictValue): string[] {
  const annots = doc.deref(pageDict.get("Annots") ?? null);
  const names = new Set<string>();
  for (const annot of isPdfArray(annots) ? annots : []) {
    let dict = doc.derefDict(annot);
    if (!dict || dictGetName(dict, "Subtype") !== "Widget") {
      continue;
    }
    const parts: string[] = [];
    const seen = new Set<PdfDictValue>();
    while (dict && !seen.has(dict)) {
      seen.add(dict);
      const partial = getDictStringValue(dict, "T", doc);
      if (partial) {
        parts.unshift(partial);
      }
      dict = doc.derefDict(dict.get("Parent") ?? null);
    }
    if (parts.length > 0) {
      names.add(parts.join("."));
    }
  }
  return [...names];
}

/** @internal - An inheritable page entry: the page's own, or the nearest page-tree node's. */
function inheritedPageEntry(
  doc: PdfDocument,
  pageDict: PdfDictValue,
  key: string
): PdfObject | undefined {
  const seen = new Set<PdfDictValue>();
  let current: PdfDictValue | null = pageDict;
  while (current && !seen.has(current)) {
    seen.add(current);
    const value = current.get(key);
    if (value !== undefined && value !== null) {
      return value;
    }
    current = doc.derefDict(current.get("Parent") ?? null);
  }
  return undefined;
}

/** @internal - Whether a page snapshot puts anything on the page it overlays. */
function hasOverlay(page: PdfPageSnapshot): boolean {
  return (
    page._stream.hasContent() ||
    page._images.length > 0 ||
    page._builderAnnotations.length > 0 ||
    page._formFields.length > 0
  );
}

// =============================================================================
// PdfEditor
// =============================================================================

/**
 * Editor for modifying existing PDF documents.
 *
 * Load an existing PDF, overlay content on existing pages, fill form fields,
 * add new pages, copy pages from other documents, and save.
 *
 * ## Page indices
 *
 * One 0-based index names each page, and never changes:
 *
 * - `0` to `pageCount - 1` name the pages the document was **read** with;
 * - `pageCount` onwards name the pages **added** with {@link addPage} and
 *   {@link copyPagesFrom}, in the order they were added.
 *
 * Every call that takes an index — {@link getPage}, {@link removePage}, {@link rotatePage},
 * {@link isPageRemoved} and {@link splitPages} — takes the whole range. Removing a page does
 * not renumber the others: it stays a valid index that {@link isPageRemoved} answers `true`
 * for, and that the calls which edit or write a page refuse. {@link pageCount} is the number
 * of pages the document was read with, and does not change; {@link outputPageCount} is the
 * number of pages a write produces now.
 *
 * `getPage` hands out one handle per page: for a page read or copied in, it draws over the
 * page's content; for a page added with `addPage`, it draws onto that page, as the builder
 * `addPage` returned does.
 */
export class PdfEditor {
  private _doc: PdfDocument;
  private _password: string;
  private _pages: PdfEditorPage[] = [];
  /**
   * Pages added with {@link addPage} and {@link copyPagesFrom}, in the order they were added:
   * the document ends with them, after the original pages that were not removed.
   */
  private _appendedPages: AppendedPage[] = [];
  private _fontManager: FontManager;
  /** Prefix that isolates every overlay resource from the original document. */
  private _overlayResourcePrefix: string;
  private _formFieldUpdates = new Map<string, string>();
  /** @internal - Document properties set with {@link setMetadata}; written by the next write. */
  private _metadata: PdfMetadataUpdate | undefined;
  /** @internal - The fields of the document as it was read, by name; read on first need. */
  private _fieldsByName: Map<string, PdfFormField> | undefined;
  /** @internal - Indices of original pages to remove on save */
  private _removedPageIndices = new Set<number>();
  /** @internal - True during an incremental update to preserve original refs */
  private _isIncrementalSave = false;
  /** @internal - Rotation overrides for original pages: index → degrees (0/90/180/270) */
  private _rotationOverrides = new Map<number, number>();
  private _signaturePlaceholder: string | null = null;
  private _hasSignatures: boolean | undefined;
  /** @internal - Writer reference during save(), for deep-clone */
  private _writerForSave: PdfWriter | null = null;
  /** @internal - Cache of cloned indirect refs: "objNum:gen" → new objNum in writer */
  private _clonedRefs = new Map<string, number>();
  /** @internal - Annotation objects already written by the current full rewrite */
  private _writtenAnnotObjNums = new Set<number>();
  /**
   * @internal - While a full rewrite that leaves original pages out runs: what it leaves out —
   * those pages, the annotations only they list, the form fields that lived only on them (see
   * `page-index.ts`). A reference to a dropped object is written as `null`, which the PDF
   * format reads as an absent entry.
   */
  private _pruning: Pruning | null = null;
  /**
   * @internal - While a full rewrite that leaves original pages out or moves them runs: the
   * article threads as it writes them — the beads it leaves out, and the links it rewrites
   * between the beads that are left (see `planThreads` in `catalog-rewrite.ts`).
   */
  private _threadPlan: ThreadPlan | null = null;
  /** @internal - Which objects go with which page, read on the first rewrite that needs it. */
  private _pageIndex: PageIndex | undefined;
  /**
   * @internal - While a full rewrite runs: original objects the catalog rewrite leaves out
   * (by `"objNum:gen"`), consulted beside {@link _pruning} — the outline items it does not
   * write. A predicate rather than a set, so a part of {@link splitPages} need not list them.
   */
  private _catalogDropped: ((key: string) => boolean) | null = null;
  /** @internal - What leads where in the catalog, worked out on the first full rewrite. */
  private _catalogIndex: CatalogIndex | undefined;
  /** @internal - While a write runs: the button fields whose `/V` it changes (see `_buttonParentValues`). */
  private _buttonParents: Map<string, { ref: PdfRef; dict: PdfDictValue; value: string }> | null =
    null;
  /**
   * @internal - While a full rewrite runs: the root fields it writes beyond the document's
   * own — added with `addFormField`, or brought in by `copyPagesFrom` — for its AcroForm.
   */
  private _addedFieldRefs: number[] = [];
  /** @internal - Ids of the source documents objects are cloned from (see `_docId`). */
  private _sourceDocIds = new Map<PdfDocument, string>();
  /**
   * @internal - A suffix `_docId` adds while a copied page's annotations are cloned, so each
   * copy of a page gets annotations of its own (one annotation object belongs to one page).
   */
  private _cloneScope = "";
  /** @internal - How many page copies have cloned their annotations (see `_cloneScope`). */
  private _copyCount = 0;
  /** @internal - Source objects a page copy leaves out: the source's other pages. */
  private _droppedSourceRefs = new Set<string>();
  /** @internal - Objects `_deepSerialize` has numbered but not yet written. */
  private _cloneQueue: PendingClone[] | null = null;
  /**
   * @internal - The tail of this editor's writes. A write keeps its working state on the editor
   * (the writer, clone maps, the font manager's build, the signature placeholder) across awaits, so
   * two writes in flight at once would corrupt each other's output. Each public write therefore
   * waits for the previous one to settle — fulfilled or rejected, so one failure does not fail the
   * writes queued behind it — and runs alone.
   *
   * A write reflects the edits made before it was **called**, not before it started running:
   * `_serialized` freezes them into {@link _writeEdits} synchronously, and with no write in flight
   * the write also starts synchronously, as it always did. Font registration (`embedFont`) is
   * configuration rather than an edit and is not frozen: a write queued behind another sees it.
   */
  private _writeTail: Promise<unknown> = Promise.resolve();
  /** @internal - Writes called and not yet settled. */
  private _writesInFlight = 0;
  /** @internal - The edits the running write serializes, frozen when it was called. */
  private _writeEdits: EditorEdits | null = null;
  /** @internal - The font options the document was read with (`EditorReadOptions`). */
  private readonly _fontSettings: DrawnTextFontSettings;
  /** @internal - Whether it was read with `fonts`, which leaves `embedFont` nothing to do. */
  private readonly _hasFontConfig: boolean;

  private constructor(data: Uint8Array, options: EditorReadOptions) {
    this._doc = new PdfDocument(data);
    const password = options.password ?? "";
    this._password = password;
    this._fontSettings = {
      disableFontAutoDiscovery: options.disableFontAutoDiscovery === true,
      preferSystemFonts: options.preferSystemFonts ? [...options.preferSystemFonts] : [],
      textLanguage: options.textLanguage,
      onWarning: options.onWarning
    };
    this._hasFontConfig = options.fonts !== undefined;

    // Handle encryption before anything resolves an object: resolved objects
    // and parsed object streams are cached, so one read first would stay
    // undecrypted for the editor's lifetime.
    if (isEncrypted(this._doc)) {
      const success = initDecryption(this._doc, password);
      if (!success) {
        throw new PdfPasswordError("Failed to decrypt PDF: incorrect password");
      }
    }

    this._overlayResourcePrefix = this._chooseOverlayResourcePrefix();
    this._fontManager = new FontManager(
      options.fonts === undefined ? undefined : compilePdfFontConfig(options.fonts),
      this._overlayResourcePrefix
    );

    // Initialize page proxies
    const pagesInfo = this._doc.getPagesWithObjInfo();
    for (let i = 0; i < pagesInfo.length; i++) {
      const { dict } = pagesInfo[i];
      const dims = this._doc.resolvePageBox(dict) ?? {
        width: 612,
        height: 792
      };
      this._pages.push(
        PdfEditorPage._create(
          i,
          dims.width,
          dims.height,
          this._fontManager,
          this._overlayResourcePrefix,
          {
            isPageRemoved: index => this._removedPageIndices.has(index),
            checkFieldName: name => this._checkNewFieldName("Pdf.EditorPage#addFormField", name)
          }
        )
      );
    }
    this._doc.throwIfLimitExceeded();
  }

  /**
   * Read a PDF for editing.
   *
   * Named `read` like every other bytes-in entry point (`Pdf.extract`, `Workbook.read`);
   * the result is an editor rather than an extraction result because that is what the
   * class is for.
   *
   * @param data - Raw PDF file bytes — a `Uint8Array`, another `ArrayBufferView` (a `Buffer`,
   *   a `DataView`) or an `ArrayBuffer`, viewed rather than copied
   * @param options - The password, and the font options for text drawn on the document
   *   (see {@link EditorReadOptions})
   * @returns A PdfEditor instance
   * @throws {PdfArgumentError} `data` is neither, or `options` is malformed.
   */
  static read(
    data: Uint8Array | ArrayBuffer | ArrayBufferView,
    options?: EditorReadOptions
  ): PdfEditor {
    const where = "Pdf.Editor.read";
    const bytes = pdfBytes(where, "data", data);
    check(editorReadOptionsSchema.optional(), options, at(where, "options"));
    checkFontOptionNames(where, "options.", options);
    return new PdfEditor(bytes, options ?? {});
  }

  /** @deprecated Use {@link PdfEditor.read}. Removed in 1.0. */
  static load(data: Uint8Array | ArrayBuffer, options?: EditorReadOptions): PdfEditor {
    return PdfEditor.read(data, options);
  }

  /**
   * Whether the loaded document already carries a digital signature: a
   * signature field (/FT /Sig, directly or inherited from a parent field)
   * whose value is a signature dictionary with /ByteRange and /Contents.
   * The /SigFlags SignaturesExist bit alone does not count.
   *
   * A full rewrite ({@link toBuffer}) invalidates such signatures. An
   * incremental update (`toBuffer({ incremental: true })`) leaves the signed byte
   * range intact, so the signature still verifies over the revision it
   * covers — but certification (DocMDP) and field-lock (FieldMDP)
   * permissions are **not** evaluated: an update a certifying signature
   * forbids is still written, and a validator may report it.
   */
  get hasSignatures(): boolean {
    this._hasSignatures ??= this._detectSignatures();
    return this._hasSignatures;
  }

  /**
   * Number of pages the document had when it was read. Indices `0` to `pageCount - 1` name
   * those pages, and the pages added with {@link addPage} and {@link copyPagesFrom} are
   * numbered after them (see the class description). It does not change when a page is
   * removed or added; {@link outputPageCount} is the number of pages a write produces.
   */
  get pageCount(): number {
    return this._pages.length;
  }

  /**
   * Number of pages {@link toBuffer} would write now: the pages the document was read with and
   * the pages added since, less the ones removed.
   */
  get outputPageCount(): number {
    return this._pages.length + this._appendedPages.length - this._removedPageIndices.size;
  }

  /**
   * Get a page for editing — one the document was read with, or one added with
   * {@link addPage} or {@link copyPagesFrom} (numbered after those, in the order added).
   *
   * The handle draws over the page's content: for a page read or copied in, onto an overlay
   * written above what the page already shows; for a page added with `addPage`, onto that
   * page itself — the same page as the builder `addPage` returned. Every call for one index
   * returns the same handle.
   *
   * A removed page cannot be edited: this refuses it, and a page held from before
   * {@link removePage} refuses every drawing call (measuring is still allowed). Ask
   * {@link isPageRemoved} first when it may have been removed.
   *
   * @param index - 0-based index of the page (see the class description)
   * @throws {PdfArgumentError} The index is not an integer, names no page, or names a removed
   *   one (field `index`).
   */
  getPage(index: number): PdfEditorPage {
    this._checkPageIndex("Pdf.Editor#getPage", index);
    this._checkNotRemoved("Pdf.Editor#getPage", index);
    return index < this._pages.length ? this._pages[index] : this._appendedHandle(index);
  }

  /**
   * Add a new blank page to the end of the document: after the original pages that were not
   * removed, and after every page added or copied in before it.
   *
   * Draw on the page through the builder this returns, or through {@link getPage} with its
   * index — {@link pageCount} plus the number of pages added before it, as {@link splitPages},
   * {@link removePage}, {@link rotatePage} and {@link isPageRemoved} number it. Once the page is
   * removed, the builder refuses every drawing call, as a held page handle does.
   */
  addPage(options?: PageOptions): PdfPageBuilder {
    checkPageOptions("Pdf.Editor#addPage", options);
    const width = options?.width ?? 595.28;
    const height = options?.height ?? 841.89;
    const page = PdfPageBuilder._create(
      width,
      height,
      this._fontManager,
      this._overlayResourcePrefix,
      name => this._checkNewFieldName("Pdf.PageBuilder#addFormField", name)
    );
    const index = this._pages.length + this._appendedPages.length;
    page._liveCheck = where => this._checkNotRemoved(where, index, "page");
    this._appendedPages.push({ kind: "new", page });
    return page;
  }

  /**
   * Remove a page from the document — one it was read with, or one added since.
   *
   * Indices do not shift when a page is removed: on a three-page document, `removePage(0)` then
   * `removePage(1)` removes the first two pages, and {@link pageCount} stays 3
   * ({@link outputPageCount} becomes 1). Removal is idempotent: removing a page that is already
   * removed changes nothing. A removed added page is left out of every write, with its form
   * fields, whose names are free again. Removing every page is allowed as long as a page is
   * added or copied in before the document is written; writing a document with no pages throws.
   *
   * @param index - 0-based index of the page (see the class description)
   * @throws {PdfArgumentError} The index is not an integer, or names no page (field `index`).
   */
  removePage(index: number): this {
    this._checkPageIndex("Pdf.Editor#removePage", index);
    this._removedPageIndices.add(index);
    return this;
  }

  /**
   * Whether {@link removePage} has removed a page — one the document was read with, or one
   * added since.
   *
   * @param index - 0-based index of the page (see the class description)
   * @throws {PdfArgumentError} The index is not an integer, or names no page (field `index`).
   */
  isPageRemoved(index: number): boolean {
    this._checkPageIndex("Pdf.Editor#isPageRemoved", index);
    return this._removedPageIndices.has(index);
  }

  /**
   * @internal - Throw unless `index` names a page: one of the document as it was read, or one
   * added since (numbered after them).
   *
   * Every bad index is a `PdfArgumentError` with field `index` — the same class and field as
   * an index naming a removed page ({@link _checkNotRemoved}), so one `catch` covers every
   * index the caller got wrong. These were `PdfStructureError`, which describes a malformed
   * file, not a wrong argument. A fraction or `NaN` inside the range once passed: `getPage`
   * returned `undefined` and `removePage`/`rotatePage` recorded an edit no page matched. A
   * value that is not a number at all is checked before the range: `"1"` passed it by
   * coercion.
   */
  private _checkPageIndex(where: string, index: number): void {
    if (typeof index !== "number" || Number.isNaN(index)) {
      throw new PdfArgumentError(
        "index",
        `${where}: index must be an integer, got ${describeIndex(index)}`
      );
    }
    const count = this._pages.length + this._appendedPages.length;
    if (index < 0 || index >= count) {
      throw new PdfArgumentError(
        "index",
        count === 0
          ? `${where}: index ${index} names no page: the document has no pages`
          : `${where}: index ${index} names no page (0-${count - 1})`
      );
    }
    if (!Number.isInteger(index)) {
      throw new PdfArgumentError(
        "index",
        `${where}: index must be an integer, got ${describeIndex(index)}`
      );
    }
  }

  /**
   * @internal - Throw when page `index` has been removed: an edit of it was accepted and then
   * written nowhere. The index is a valid one; the request is what cannot be honoured.
   */
  private _checkNotRemoved(where: string, index: number, field = "index"): void {
    if (this._removedPageIndices.has(index)) {
      throw new PdfArgumentError(
        field,
        `${where}: page ${index} has been removed (removePage); isPageRemoved() tells`
      );
    }
  }

  /**
   * @internal - The handle {@link getPage} returns for an added page, made on first request.
   * A page added with `addPage` is drawn on through its own builder; a copied page gets an
   * overlay of its own, written above its content.
   */
  private _appendedHandle(index: number): PdfEditorPage {
    const appended = this._appendedPages[index - this._pages.length];
    if (appended.handle) {
      return appended.handle;
    }
    const owner: EditorPageOwner = {
      isPageRemoved: i => this._removedPageIndices.has(i),
      checkFieldName: name => this._checkNewFieldName("Pdf.EditorPage#addFormField", name)
    };
    if (appended.kind === "new") {
      appended.handle = PdfEditorPage._create(
        index,
        appended.page._width,
        appended.page._height,
        this._fontManager,
        this._overlayResourcePrefix,
        owner,
        appended.page
      );
      return appended.handle;
    }
    this._checkOverlayPrefixFree(appended);
    appended.handle = PdfEditorPage._create(
      index,
      appended.width,
      appended.height,
      this._fontManager,
      this._overlayResourcePrefix,
      owner
    );
    appended.overlay = appended.handle._overlay;
    return appended.handle;
  }

  /**
   * @internal - Refuse an overlay on a copied page whose own resources use the names the
   * overlay's resources are given. The prefix was chosen to be absent from this document's
   * pages; a page copied from another document could, in principle, use it, and merging the
   * two resource dictionaries would then repoint the page's own content at the overlay's font.
   */
  private _checkOverlayPrefixFree(copied: CopiedPage): void {
    const resources = copied.sourceDoc.resolvePageResources(copied.sourcePageDict);
    for (const category of ["Font", "XObject", "ExtGState"]) {
      const value = resources.get(category);
      const subDict = value ? copied.sourceDoc.derefDict(value) : null;
      for (const name of subDict?.keys() ?? []) {
        if (name.startsWith(this._overlayResourcePrefix)) {
          throw new PdfStructureError(
            `Pdf.Editor#getPage: the copied page's resource /${name} uses the name prefix this ` +
              `editor gives the resources of what is drawn over a page (${this._overlayResourcePrefix}), ` +
              `so nothing can be drawn over it without changing what the page shows`
          );
        }
      }
    }
  }

  /**
   * Set the rotation of a page — one the document was read with, or one added since.
   *
   * @param index - 0-based index of the page (see the class description)
   * @param degrees - Rotation in degrees (must be 0, 90, 180, or 270)
   * @throws {PdfArgumentError} The index is not an integer, names no page or names a removed
   *   one (field `index`), or `degrees` is not a quarter turn (field `degrees`).
   */
  rotatePage(index: number, degrees: number): this {
    this._checkPageIndex("Pdf.Editor#rotatePage", index);
    this._checkNotRemoved("Pdf.Editor#rotatePage", index);
    if (degrees !== 0 && degrees !== 90 && degrees !== 180 && degrees !== 270) {
      throw new PdfArgumentError(
        "degrees",
        `Pdf.Editor#rotatePage: degrees must be 0, 90, 180 or 270, got ${describeIndex(degrees)}`
      );
    }
    this._rotationOverrides.set(index, degrees);
    return this;
  }

  /**
   * Split the document: save each page (or a subset) as a separate PDF.
   *
   * The document split is the one {@link toBuffer} would write now: a removed page is not a
   * part, every other edit — content drawn on a page, a rotation, a form value — is in the
   * part it concerns, and the pages added with {@link addPage} or {@link copyPagesFrom}
   * follow the original pages in the order they were added. Each part is what `toBuffer()`
   * writes with every other page removed, so it keeps the encryption, the form fields whose
   * widgets are on its page and the fields no page shows, the document-level entries (see
   * {@link toBuffer}), and throws `PdfSignatureInvalidationError` for a signed document. The
   * one difference is the outline: a part keeps only the bookmarks that lead to its page,
   * with their ancestors, where a rewrite also keeps the bookmarks that lead to no page (a
   * link, say). As with every write, the edits are the ones made before the call.
   *
   * @param options - Which pages to write (`pageIndices`, see {@link EditorSplitOptions}), and,
   *   as for {@link toBuffer}, `encryption: "remove"` to write decrypted parts and
   *   `invalidateSignatures` to permit splitting a signed document.
   * @returns Array of Uint8Array, one per requested page, in the order requested.
   * @throws {PdfArgumentError} An index names no page, or a removed one (field
   *   `options.pageIndices[i]`) — before any part is written.
   */
  splitPages(options?: EditorSplitOptions): Promise<Uint8Array[]>;
  /**
   * @deprecated Pass `{ pageIndices }` as one options object. The positional form is removed
   * in 1.0. It skips an index out of range or naming a removed page, as it always has, and
   * throws `PdfStructureError` for one in range that is not an integer.
   */
  splitPages(pageIndices?: readonly number[] | null): Promise<Uint8Array[]>;
  async splitPages(
    optionsOrIndices?: EditorSplitOptions | readonly number[] | null
  ): Promise<Uint8Array[]> {
    // Anything but an options object is the positional list it has always been — `null`
    // included, which meant "every page".
    const where = "Pdf.Editor#splitPages";
    const isOptions = isRecordArgument(optionsOrIndices);
    // Neither form: the positional list was iterated as given, so a value it cannot iterate
    // failed there with a `TypeError`.
    if (
      !isOptions &&
      optionsOrIndices != null &&
      typeof (optionsOrIndices as { [Symbol.iterator]?: unknown })[Symbol.iterator] !== "function"
    ) {
      throw new PdfArgumentError(
        "options",
        `${where}: options must be an options object { pageIndices }, got ${describeValue(optionsOrIndices)}`
      );
    }
    const options: EditorSplitOptions | undefined = isOptions
      ? (optionsOrIndices as EditorSplitOptions)
      : undefined;
    if (options !== undefined) {
      checkNotPagesKey(where, options);
    }
    if (options?.pageIndices !== undefined && !Array.isArray(options.pageIndices)) {
      throw new PdfArgumentError(
        "options.pageIndices",
        `${where}: options.pageIndices must be an array of 0-based page indices`
      );
    }
    const pageIndices = isOptions
      ? options?.pageIndices
      : ((optionsOrIndices ?? undefined) as readonly number[] | undefined);
    checkWriteOptions(where, options);
    const edits = this._snapshotEdits();
    // Every added page, by index, and every removed page, read or added: a part names its page
    // in the one index space the editor uses.
    const appended = this._snapshotAppended();
    const removed = new Set(this._removedPageIndices);
    const pageCount = this._pages.length;
    const total = pageCount + appended.length;
    const requested = pageIndices ?? Array.from({ length: total }, (_, i) => i);
    // The options form names pages, so each index must name one the document has: the result
    // has one part per index, in order, and a part silently missing could not be told apart.
    // Every index is checked before the first part is written. The deprecated positional form
    // keeps skipping an index out of range or naming a removed page, as it always has.
    if (isOptions) {
      requested.forEach((idx, i) => {
        if (typeof idx !== "number" || !Number.isInteger(idx) || idx < 0 || idx >= total) {
          throw new PdfArgumentError(
            `options.pageIndices[${i}]`,
            `${where}: options.pageIndices[${i}] must be the 0-based index of a page of the document ` +
              `(0 to ${total - 1}, the pages added after the ones it was read with), got ${describeIndex(idx)}`
          );
        }
        if (removed.has(idx)) {
          throw new PdfArgumentError(
            `options.pageIndices[${i}]`,
            `${where}: options.pageIndices[${i}] names page ${idx}, which has been removed (removePage)`
          );
        }
      });
    }
    const indices = requested;
    const results: Uint8Array[] = [];

    // Each part names the one page it keeps rather than removing all the others, with what
    // goes with which page read once (`page-index.ts`). Removing the other pages one by one
    // made every part a rewrite of the whole document, and splitting quadratic in the page
    // count.
    for (const idx of indices) {
      if (typeof idx === "number" && (idx < 0 || idx >= total)) {
        continue;
      }
      if (!Number.isInteger(idx)) {
        // Removing every other page from such an index removed them all, which threw this.
        throw new PdfStructureError(
          `${where}: page index must be an integer, got ${describeIndex(idx)}`
        );
      }
      if (removed.has(idx)) {
        continue;
      }
      const original = idx < pageCount;
      const part: EditorEdits = {
        ...edits,
        pages: { kind: "only", index: original ? idx : null },
        appendedPages: original ? [] : [appended[idx - pageCount]]
      };
      results.push(
        await this._serialized(
          () =>
            this._writeFull(options, {
              caller: "splitPages()",
              legacyFormRules: false,
              rewriteEncryption: "keep"
            }),
          part
        )
      );
    }

    return results;
  }

  /**
   * Pick a deterministic resource prefix absent from every original page.
   *
   * Overlay content is merged into each page resource dictionary. Fonts,
   * images, and ExtGState entries therefore need a private namespace; otherwise
   * an overlay `/EF1` or `/Im1` silently replaces the original page resource and
   * corrupts content that was already there.
   */
  private _chooseOverlayResourcePrefix(): string {
    const used = new Set<string>();
    for (const { dict } of this._doc.getPagesWithObjInfo()) {
      const resources = this._doc.resolvePageResources(dict);
      for (const category of ["Font", "XObject", "ExtGState"]) {
        const value = resources.get(category);
        const subDict = value ? this._doc.derefDict(value) : null;
        if (!subDict) {
          continue;
        }
        for (const name of subDict.keys()) {
          used.add(name);
        }
      }
    }
    let id = 1;
    while ([...used].some(name => name.startsWith(`DMO${id}_`))) {
      id++;
    }
    return `DMO${id}_`;
  }

  /**
   * Embed a TrueType font for Unicode/CJK support.
   *
   * @param source - Raw .ttf or .ttc file bytes, or `{ data, collectionIndex }` to pick a
   *   face of a `.ttc` (the index defaults to 0). Overlay text — a watermark, a stamp, a page
   *   number — is drawn by this library rather than copied from the file, so it needs a face
   *   like any other text; without a way to name one inside a collection, a caller holding a
   *   `.ttc` had to embed whichever face happened to be first.
   * @throws {PdfArgumentError} `source` is not font bytes or `{ data, collectionIndex }`, or
   *   the document was read with `options.fonts`.
   */
  embedFont(source: PdfFontSource): this;
  /**
   * @deprecated Pass `{ data: fontBytes, collectionIndex }` as one {@link PdfFontSource}. The
   * positional `collectionIndex` is removed in 1.0.
   */
  embedFont(fontBytes: Uint8Array, collectionIndex: number | undefined): this;
  embedFont(source: PdfFontSource, collectionIndex?: number): this {
    const where = "Pdf.Editor#embedFont";
    const face = readFontSourceArgument(where, source, collectionIndex);
    if (this._hasFontConfig) {
      throw new PdfArgumentError(
        face.field,
        `${where}: the document was read with options.fonts, which already names the faces drawn text uses — add this face to that configuration instead`
      );
    }
    this._fontManager.registerEmbeddedFont(parseTtf(face.data, face.collectionIndex));
    return this;
  }

  // ===========================================================================
  // Metadata
  // ===========================================================================

  /**
   * Set or remove document properties — the fields of `Pdf.Builder#setMetadata`, with the
   * same checks, as a JSON merge patch ({@link PdfMetadataUpdate}): a field given replaces the
   * document's own, a field set to `null` removes it, and a field left out (or `undefined`)
   * keeps what the document has — as does every entry this has no field for (custom
   * entries). Calling it again merges onto the previous call, a later value or `null`
   * replacing an earlier one.
   *
   * ```ts
   * editor.setMetadata({ title: "Q3", author: null }); // set the title, remove the author
   * ``` Nothing is invented:
   * neither write adds a `/Creator` or `/Producer` the document did not have — `producer`
   * and `creator` change only when given — and no date changes unless it is given.
   *
   * The properties are written to the document information dictionary and, when the file has
   * an XMP metadata packet, to the same properties in it (`dc:title`, `dc:creator`,
   * `dc:description`, `pdf:Keywords`, `xmp:CreatorTool`, `pdf:Producer`, `xmp:CreateDate`,
   * `xmp:ModifyDate`), so the two agree. The packet is parsed as XML and edited, not
   * pattern-matched:
   * - only the document's own (top-level) `rdf:Description`s are changed; a property is
   *   changed where it is, in element or attribute form, and one that is absent is added
   *   with its namespace bound;
   * - a language alternative (`dc:title`, `dc:description`) replaces only its `x-default`
   *   entry, so other languages survive;
   * - `dc:creator` is written from `author` split on `;` (each name trimmed, empty names
   *   dropped): `"Ann; Bob"` is two authors;
   * - `xmp:MetadataDate` is set — to `modDate` when given, otherwise the time of the write —
   *   as XMP requires whenever the metadata changes;
   * - the packet is written back uncompressed, in its own encoding (UTF-8, with or without a
   *   byte-order mark, or UTF-16).
   *
   * - a property set to `null` is removed from every top-level description, in every form
   *   (the element with its whole `rdf:Alt`/`rdf:Seq`, or the attribute), and its `/Info`
   *   entry goes too, so the two records still agree.
   *
   * A packet that cannot be decoded, or is not XMP, is left exactly as it is, and only the
   * information dictionary changes. Both a full rewrite and an incremental update write the
   * properties; an incremental update with no other edit appends just the changed objects.
   *
   * The values are copied when the call is made; like every edit, a write reflects the
   * metadata set before it was called.
   *
   * @throws {PdfArgumentError} `metadata` is not an object, a string field is not a string or
   *   contains a character XML cannot carry (a control character other than tab, newline and
   *   carriage return; U+FFFE/U+FFFF; a lone surrogate), or a date is not a valid `Date`.
   */
  setMetadata(metadata: PdfMetadataUpdate): this {
    checkMetadataUpdate("Pdf.Editor#setMetadata", metadata);
    // Merge patch onto the previous calls: a later value or `null` replaces an earlier one.
    this._metadata = { ...this._metadata, ...copyMetadataUpdate(metadata) };
    return this;
  }

  // ===========================================================================
  // Form Fields
  // ===========================================================================

  /**
   * Set the value of a form field.
   * The field is identified by its fully qualified name (e.g., "form.address.city").
   *
   * Values are written into the AcroForm, and saving edited values also adjusts the
   * form so viewers show them: `/NeedAppearances` becomes `true`, and in a **hybrid**
   * form (AcroForm fields plus an XFA packet) the `/XFA` entry is **removed**, because
   * an XFA-aware viewer would otherwise render the stale XFA data. A **dynamic** XFA form
   * (`/NeedsRendering true`, XFA with no AcroForm fields, or an XFA config declaring
   * `dynamicRender` `required`) is never stripped: saving
   * edited values into one throws `PdfXfaFormError` and writes nothing. A save without
   * edited values leaves `/XFA` and `/NeedsRendering` as they were.
   *
   * Those XFA rules are `toBuffer()`'s. `sign()` and the deprecated `save()` /
   * `saveIncremental()` keep their original behaviour: `/XFA` is left in place, a dynamic
   * XFA form is written without complaint, and an incremental update appends no AcroForm
   * revision.
   *
   * @param fieldName - Fully qualified field name, as {@link getFormFields} lists it
   * @param value - New value to set. For a checkbox or radio group, `"Off"` or one of the
   *   states its widgets can show (a checkbox's export value, a radio button's value).
   * @throws {PdfArgumentError} The document has no field named `fieldName`, or `value` is not
   *   a state of the checkbox or radio group it names.
   */
  setFormField(fieldName: string, value: string): this {
    checkString("Pdf.Editor#setFormField", "fieldName", fieldName);
    checkString("Pdf.Editor#setFormField", "value", value);
    this._checkFieldValue("Pdf.Editor#setFormField", "fieldName", fieldName, value);
    this._formFieldUpdates.set(fieldName, value);
    return this;
  }

  /**
   * Set multiple form field values at once. Every entry is checked before any is set, so a
   * call that throws sets none.
   *
   * @param fields - Object mapping field names to values
   * @throws {PdfArgumentError} as for {@link setFormField}, for any entry.
   */
  setFormFields(fields: Record<string, string>): this {
    check(
      s.record(s.string(), { expected: "an object of field names to values" }),
      fields,
      at("Pdf.Editor#setFormFields", "fields")
    );
    for (const [name, value] of Object.entries(fields)) {
      this._checkFieldValue(
        "Pdf.Editor#setFormFields",
        `fields[${JSON.stringify(name)}]`,
        name,
        value
      );
    }
    for (const [name, value] of Object.entries(fields)) {
      this._formFieldUpdates.set(name, value);
    }
    return this;
  }

  /**
   * @internal - Throw unless `name` is a field of the document as it was read, and `value`
   * is a state its buttons can show. A value for a field that does not exist used to be
   * dropped without a word, and a checkbox set to anything but `Off` or its export value
   * was written in a state no viewer draws as checked.
   */
  private _checkFieldValue(where: string, field: string, name: string, value: string): void {
    this._fieldsByName ??= new Map(extractFormFields(this._doc).map(entry => [entry.name, entry]));
    const existing = this._fieldsByName.get(name);
    // A dynamic XFA form has its fields in the XFA packet, not the AcroForm; writing values
    // into one is refused when the document is written, with `PdfXfaFormError`, which says why.
    if (existing === undefined && this._isDynamicXfa()) {
      return;
    }
    if (existing === undefined) {
      throw new PdfArgumentError(
        field,
        `${where}: the document has no form field named ${JSON.stringify(name)}` +
          (this._fieldsByName.size === 0
            ? " (it has no form fields)"
            : " — getFormFields() lists the fields it has")
      );
    }
    if (existing.type !== "checkbox" && existing.type !== "radio") {
      return;
    }
    const states = this._buttonStates(name);
    // A checkbox drawn without appearances has the on-state the field reports (`Yes` by
    // convention, as the reader reads it).
    if (states.size === 0 && existing.type === "checkbox") {
      states.add(existing.exportValue || "Yes");
    }
    if (value !== "Off" && states.size > 0 && !states.has(value)) {
      throw new PdfArgumentError(
        field,
        `${where}: ${JSON.stringify(value)} is not a state of the ${existing.type} ${JSON.stringify(name)}; ` +
          `it can be "Off" or ${[...states].map(state => JSON.stringify(state)).join(", ")}`
      );
    }
  }

  /** @internal - The on-states the widgets of button field `name` can show (their `/AP /N` keys). */
  private _buttonStates(name: string): Set<string> {
    const states = new Set<string>();
    for (const { dict: pageDict } of this._doc.getPagesWithObjInfo()) {
      const annots = this._doc.deref(pageDict.get("Annots"));
      if (!isPdfArray(annots)) {
        continue;
      }
      for (const annotRef of annots) {
        const annotDict = this._doc.derefDict(annotRef);
        if (
          !annotDict ||
          dictGetName(annotDict, "Subtype") !== "Widget" ||
          this._resolveFieldName(annotDict) !== name
        ) {
          continue;
        }
        const normal = this._doc.derefDict(this._doc.derefDict(annotDict.get("AP"))?.get("N"));
        for (const key of normal?.keys() ?? []) {
          if (key !== "Off") {
            states.add(key);
          }
        }
      }
    }
    return states;
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
   * Copy pages from another PDF document into this document. They go to the end, after
   * every page added or copied in before them (see {@link addPage}).
   *
   * `pageIndices` lists 0-based indices of the source's pages, copied in the order given; omit
   * it to copy every page. Each must name a page of the source: an index outside it, or one
   * that is not an integer, is refused (`PdfArgumentError`, field `options.pageIndices[i]`)
   * before anything is copied. (The deprecated positional form keeps skipping an index outside
   * the source, as it always has.) The copied pages are numbered after the pages already in the
   * document, in the order given — the indices {@link getPage}, {@link removePage},
   * {@link rotatePage} and {@link splitPages} take for them.
   *
   * A copied page keeps its geometry (its `/MediaBox` as the source has it, origin included,
   * and an inherited `/Rotate` or `/CropBox`) and its annotations — links, comments and form
   * widgets, each copy with objects of its own. The form fields whose widgets are on the page
   * join this document's form; a reference to another page of the source (a link's
   * destination, a widget of the same field there) is written as `null`. A copied field whose
   * fully qualified name this document already has, or another copied or added field has, is
   * refused by this call; pass `{ formFields: "drop" }` to copy the pages without their form
   * fields (other annotations are still copied) — the way to merge the same form twice.
   *
   * @param sourcePdf - Raw bytes of the source PDF — a `Uint8Array`, another `ArrayBufferView`
   *   or an `ArrayBuffer`
   * @param options - Which pages to copy (`pageIndices`), and the source's password
   * @throws {PdfPasswordError} The source is encrypted and `password` does not open it.
   * @throws {PdfArgumentError} An index names no page of the source, a copied form field's
   *   name is already in use (field `sourcePdf`), or `formFields` is not `"copy"` or `"drop"`
   *   — nothing is copied.
   */
  copyPagesFrom(
    sourcePdf: Uint8Array | ArrayBuffer | ArrayBufferView,
    options?: EditorCopyPagesOptions
  ): this;
  /**
   * @deprecated Pass `{ pageIndices, password }` as one options object. The positional form
   * is removed in 1.0.
   */
  copyPagesFrom(
    sourcePdf: Uint8Array | ArrayBuffer,
    pageIndices?: number[],
    options?: EditorReadOptions
  ): this;
  copyPagesFrom(
    sourcePdf: Uint8Array | ArrayBuffer | ArrayBufferView,
    pagesOrOptions?: number[] | EditorCopyPagesOptions,
    legacyOptions?: EditorReadOptions
  ): this {
    const where = "Pdf.Editor#copyPagesFrom";
    // Anything but an options object is the positional page list — `null`, or any iterable of
    // indices, which the positional form walked with `for…of` — beside its own read options.
    const isOptionsForm = isRecordArgument(pagesOrOptions);
    const options: EditorCopyPagesOptions = isOptionsForm
      ? pagesOrOptions
      : { password: legacyOptions?.password, pageIndices: pagesOrOptions };
    if (isOptionsForm) {
      checkNotPagesKey(where, pagesOrOptions);
    }
    const formFields = options.formFields ?? "copy";
    if (formFields !== "copy" && formFields !== "drop") {
      throw new PdfArgumentError(
        "options.formFields",
        `${where}: options.formFields must be "copy" or "drop", got ${describeIndex(formFields)}`
      );
    }
    if (isOptionsForm && options.pageIndices !== undefined && !Array.isArray(options.pageIndices)) {
      throw new PdfArgumentError(
        "options.pageIndices",
        `${where}: options.pageIndices must be an array of 0-based page indices`
      );
    }
    const sourceBytes = pdfBytes(where, "sourcePdf", sourcePdf);
    checkOptionalString(where, "options.password", options.password);
    const pageIndices = options.pageIndices;
    const sourceDoc = new PdfDocument(sourceBytes);

    if (isEncrypted(sourceDoc)) {
      const success = initDecryption(sourceDoc, options?.password ?? "");
      if (!success) {
        throw new PdfPasswordError("Failed to decrypt source PDF for page copy");
      }
    }

    const sourcePagesInfo = sourceDoc.getPagesWithObjInfo();
    // Spread once: the positional form accepted any iterable, and the list is walked twice.
    const indices = pageIndices
      ? Array.from(pageIndices as Iterable<number>)
      : Array.from({ length: sourcePagesInfo.length }, (_, i) => i);

    // The options form names pages, so an index that names none is refused before any page is
    // queued. The positional form has always skipped an index outside the source, and a caller
    // copying "the first three pages" of a shorter document relies on that; anything else that
    // names no page there — a fraction, `NaN`, a non-number — crashed with a `TypeError`
    // halfway through the list, and is rejected with the `PdfStructureError` it has had since.
    indices.forEach((idx, i) => {
      if (isOptionsForm) {
        if (sourcePagesInfo[idx] === undefined || !Number.isInteger(idx)) {
          throw new PdfArgumentError(
            `options.pageIndices[${i}]`,
            `${where}: options.pageIndices[${i}] must be the 0-based index of a page of the source ` +
              `(0 to ${sourcePagesInfo.length - 1}), got ${describeIndex(idx)}`
          );
        }
      } else if (
        !(idx < 0 || idx >= sourcePagesInfo.length) &&
        sourcePagesInfo[idx] === undefined
      ) {
        throw new PdfStructureError(
          `${where}: page index must be an integer, got ${describeIndex(idx)}`
        );
      }
    });

    // A copied page brings its form fields along. Their names are in the source, so a clash —
    // with this document's fields, with fields added or copied before, or between the pages
    // of this call — is refused here, before any page is queued.
    const fieldNamesByIndex = new Map<number, string[]>();
    const copiedNames = new Set<string>();
    for (const idx of indices) {
      if (idx < 0 || idx >= sourcePagesInfo.length) {
        continue;
      }
      const names =
        formFields === "drop" ? [] : widgetFieldNames(sourceDoc, sourcePagesInfo[idx].dict);
      fieldNamesByIndex.set(idx, names);
      for (const name of names) {
        if (copiedNames.has(name)) {
          throw new PdfArgumentError(
            "sourcePdf",
            `${where}: two copied pages bring a form field named ${JSON.stringify(name)} — a fully qualified name identifies one field; pass { formFields: "drop" } to copy the pages without their form fields`
          );
        }
        copiedNames.add(name);
        this._checkNewFieldName(
          where,
          name,
          "sourcePdf",
          '; pass { formFields: "drop" } to copy the pages without their form fields'
        );
      }
    }

    for (const idx of indices) {
      if (idx < 0 || idx >= sourcePagesInfo.length) {
        continue;
      }

      const { dict: pageDict, objNum, gen } = sourcePagesInfo[idx];
      const dims = sourceDoc.resolvePageBox(pageDict) ?? {
        width: 612,
        height: 792
      };

      // Collect all content streams from the source page
      const contentStreams = this._collectContentStreams(sourceDoc, pageDict);

      this._appendedPages.push({
        kind: "copied",
        width: dims.width,
        height: dims.height,
        contentStreams,
        sourceDoc,
        sourcePageDict: pageDict,
        sourcePageRef: { objNum, gen },
        fieldNames: fieldNamesByIndex.get(idx) ?? [],
        formFields
      });
    }

    sourceDoc.throwIfLimitExceeded();
    return this;
  }

  // ===========================================================================
  // Save
  // ===========================================================================

  /**
   * Serialize the modified PDF.
   *
   * By default the file is rebuilt from scratch, which invalidates any digital signature;
   * pass `{ incremental: true }` to append an update instead (see
   * {@link EditorToBufferOptions.incremental}).
   *
   * A rebuild keeps the document-level entries of the catalog: bookmarks, named
   * destinations, page labels, the open action, viewer preferences, optional content, output
   * intents, attachments, the language and the rest. Those that name pages follow the pages
   * written — a bookmark to a removed page loses its destination, and is dropped unless a
   * bookmark beneath it survives; a named destination or open action to one is dropped; each
   * page keeps its label. The structure tree (`/StructTreeRoot`, with `/MarkInfo`) is kept
   * while no page is removed and dropped when one is, and `/Perms` and `/Legal`, which
   * concern signatures the rebuild invalidates, are always dropped.
   *
   * @returns The modified PDF as Uint8Array
   */
  async toBuffer(options: EditorToBufferOptions = {}): Promise<Uint8Array> {
    checkWriteOptions("Pdf.Editor#toBuffer", options);
    const { incremental } = options;
    if (
      incremental !== undefined &&
      typeof incremental !== "boolean" &&
      incremental !== "required"
    ) {
      throw new PdfArgumentError(
        "options.incremental",
        `Pdf.Editor#toBuffer: options.incremental must be true, false or "required", got ${describeIndex(incremental)}`
      );
    }
    return this._serialized(() => this._toBuffer(options));
  }

  /**
   * @internal - Run one write over the edits made so far, after every write already called on
   * this editor has settled. The edits are frozen here, synchronously, so neither an edit made
   * while the write waits nor one made while it runs reaches its output.
   */
  private _serialized<T>(
    write: () => Promise<T>,
    edits: EditorEdits = this._snapshotEdits()
  ): Promise<T> {
    // Checked before the write is queued, so it rejects without touching the editor's state.
    try {
      this._checkAddedFieldNames(edits);
    } catch (err) {
      return Promise.reject(err);
    }
    const start = async (): Promise<T> => {
      this._writeEdits = edits;
      try {
        return await write();
      } finally {
        this._writeEdits = null;
      }
    };
    const run = this._writesInFlight === 0 ? start() : this._writeTail.then(start, start);
    this._writesInFlight++;
    this._writeTail = run.then(
      () => void this._writesInFlight--,
      () => void this._writesInFlight--
    );
    return run;
  }

  /**
   * @internal - Where `name` is already used as a form field name: by a field the document was
   * read with, one added to a page that is still in the document, or one a copied page
   * brought. `undefined` when it is free. A field on a removed page holds no name — it is not
   * written. A field the document was read with always does, wherever its widget is.
   */
  private _fieldNameUse(name: string): string | undefined {
    this._fieldsByName ??= new Map(extractFormFields(this._doc).map(entry => [entry.name, entry]));
    const quoted = JSON.stringify(name);
    if (this._fieldsByName.has(name)) {
      return `the document already has a form field named ${quoted}`;
    }
    for (const page of this._pages) {
      if (
        !this._removedPageIndices.has(page._pageIndex) &&
        page._overlay._formFields.some(field => field.options.name === name)
      ) {
        return `a form field named ${quoted} was already added to page ${page._pageIndex}`;
      }
    }
    for (const [i, appended] of this._appendedPages.entries()) {
      // A removed page is written nowhere, so the names of its fields are free again.
      if (this._removedPageIndices.has(this._pages.length + i)) {
        continue;
      }
      if (
        appended.kind === "copied" &&
        (appended.fieldNames.includes(name) ||
          appended.overlay?._formFields.some(field => field.options.name === name))
      ) {
        return appended.fieldNames.includes(name)
          ? `a page copied in with copyPagesFrom() already brought a form field named ${quoted}`
          : `a form field named ${quoted} was already added to a page copied in with copyPagesFrom()`;
      }
      if (
        appended.kind === "new" &&
        appended.page._formFields.some(field => field.options.name === name)
      ) {
        return `a form field named ${quoted} was already added to a page added with addPage()`;
      }
    }
    return undefined;
  }

  /** @internal - Throw when `name` is already a form field name (see {@link _fieldNameUse}). */
  private _checkNewFieldName(where: string, name: string, field = "options.name", hint = ""): void {
    const use = this._fieldNameUse(name);
    if (use !== undefined) {
      throw new PdfArgumentError(
        field,
        `${where}: ${use} — a fully qualified name identifies one field${hint}`
      );
    }
  }

  /**
   * @internal - A last check, at write, that the fields a write adds have names of their own.
   * Every way to add a field refuses a clash at its own call — `addFormField` and
   * `copyPagesFrom` both know every name in use then — so no sequence of public calls is
   * known to reach this; it guards the invariant against an internal path that skips them.
   */
  private _checkAddedFieldNames(edits: EditorEdits): void {
    const added = [
      ...edits.overlays.flatMap((page, index) =>
        edits.pages.kind === "removed" && edits.pages.removed.has(index)
          ? []
          : page._formFields.map(field => field.options.name)
      ),
      ...edits.appendedPages.flatMap(page =>
        page.kind === "new"
          ? page.page._formFields.map(field => field.options.name)
          : [...page.fieldNames, ...(page.overlay?._formFields ?? []).map(f => f.options.name)]
      )
    ];
    if (added.length === 0) {
      return;
    }
    this._fieldsByName ??= new Map(extractFormFields(this._doc).map(entry => [entry.name, entry]));
    for (const name of added) {
      if (this._fieldsByName.has(name)) {
        throw new PdfArgumentError(
          "name",
          `Pdf.Editor#toBuffer: the document already has a form field named ${JSON.stringify(name)} — a fully qualified name identifies one field`
        );
      }
    }
    checkUniqueFieldNames("Pdf.Editor#toBuffer", added);
  }

  /**
   * @internal - Freeze the edits a write serializes (see {@link _serialized}). The added pages
   * a write writes are the ones not removed; {@link splitPages} takes the full list, by index,
   * from {@link _snapshotAppended}.
   */
  private _snapshotEdits(): EditorEdits {
    const pageCount = this._pages.length;
    return {
      overlays: this._pages.map(page => snapshotPage(page._overlay)),
      appendedPages: this._snapshotAppended().filter(
        (_, i) => !this._removedPageIndices.has(pageCount + i)
      ),
      formFieldUpdates: new Map(this._formFieldUpdates),
      metadata: this._metadata,
      pages: {
        kind: "removed",
        removed: new Set([...this._removedPageIndices].filter(index => index < pageCount))
      },
      rotationOverrides: new Map(
        [...this._rotationOverrides].filter(([index]) => index < pageCount)
      )
    };
  }

  /** @internal - Every added page as a write sees it, removed ones included, in index order. */
  private _snapshotAppended(): AppendedPageSnapshot[] {
    const pageCount = this._pages.length;
    return this._appendedPages.map((appended, i): AppendedPageSnapshot => {
      const rotation = this._rotationOverrides.get(pageCount + i);
      const turned = rotation === undefined ? {} : { rotation };
      if (appended.kind === "new") {
        return { kind: "new", page: snapshotPage(appended.page), ...turned };
      }
      return {
        kind: "copied",
        width: appended.width,
        height: appended.height,
        contentStreams: appended.contentStreams,
        sourceDoc: appended.sourceDoc,
        sourcePageDict: appended.sourcePageDict,
        sourcePageRef: appended.sourcePageRef,
        fieldNames: appended.fieldNames,
        formFields: appended.formFields,
        ...(appended.overlay === undefined ? {} : { overlay: snapshotPage(appended.overlay) }),
        ...turned
      };
    });
  }

  /** @internal - The running write's edits. */
  private get _edits(): EditorEdits {
    if (!this._writeEdits) {
      throw new PdfStructureError("PdfEditor: no write is running");
    }
    return this._writeEdits;
  }

  private async _toBuffer(options: EditorToBufferOptions): Promise<Uint8Array> {
    const { incremental, ...saveOptions } = options;
    if (incremental === "required") {
      return this._writeIncremental(
        saveOptions,
        {
          caller: 'toBuffer({ incremental: "required" })',
          legacyFormRules: false,
          rewriteEncryption: "keep"
        },
        true
      );
    }
    return incremental
      ? this._writeIncremental(saveOptions, {
          caller: "toBuffer({ incremental: true })",
          legacyFormRules: false,
          rewriteEncryption: "keep"
        })
      : this._writeFull(saveOptions, {
          caller: "toBuffer()",
          legacyFormRules: false,
          rewriteEncryption: "keep"
        });
  }

  /**
   * Rewrite the whole file. An encrypted document is written as a **decrypted** copy without
   * `/Encrypt`, as before, unless `{ encryption: "keep" }` is passed — `toBuffer()` keeps the
   * encryption by default.
   *
   * @deprecated Use {@link PdfEditor.toBuffer} (which keeps the encryption by default; pass
   * `{ encryption: "remove" }` for this method's decrypted copy). Removed in 1.0.
   */
  async save(options?: EditorWriteOptions): Promise<Uint8Array> {
    return this._serialized(() =>
      this._writeFull(options, {
        caller: "save()",
        legacyFormRules: true,
        rewriteEncryption: "remove"
      })
    );
  }

  /**
   * Append an incremental update, or fall back to `save()`. The appended objects of an
   * encrypted document are encrypted with its handler; when this falls back to a full rewrite,
   * the result is `save()`'s **decrypted** copy, as before, unless `{ encryption: "keep" }` is
   * passed.
   *
   * @deprecated Use `toBuffer({ incremental: true })` (whose fallback keeps the encryption by
   * default). Removed in 1.0.
   */
  async saveIncremental(options?: EditorWriteOptions): Promise<Uint8Array> {
    return this._serialized(() =>
      this._writeIncremental(options, {
        caller: "saveIncremental()",
        legacyFormRules: true,
        rewriteEncryption: "remove"
      })
    );
  }

  /**
   * Full rewrite.
   *
   * Rebuilds the PDF from scratch — content streams, resources, and page
   * properties are deep-cloned into a new document. The catalog's document-level entries are
   * carried over (see `catalog-rewrite.ts`). Digital signatures will be invalidated.
   *
   * @returns The modified PDF as Uint8Array
   */
  private async _writeFull(
    options: EditorWriteOptions | undefined,
    call: WriteCall
  ): Promise<Uint8Array> {
    const { caller } = call;
    this._guardDynamicXfa(call);
    this._guardSignatures(options, `${caller} rewrites the whole file`);
    // Rebuild the PDF (not incremental update — simpler and more reliable)
    const writer = new PdfWriter();
    const security = this._rewriteSecurity(options?.encryption ?? call.rewriteEncryption, caller);
    if (security) {
      writer.setSecurityHandler(security);
    }
    this._writerForSave = writer;
    this._clonedRefs = new Map();
    this._addedFieldRefs = [];
    this._droppedSourceRefs = new Set();
    this._pruning = null;
    this._threadPlan = null;
    this._catalogDropped = null;
    this._buttonParents = this._buttonParentValues();

    try {
      return await withDrawnTextFontBuild(this._fontManager, () =>
        this._buildFullSave(writer, call.legacyFormRules)
      );
    } finally {
      this._writerForSave = null;
      this._clonedRefs.clear();
      this._addedFieldRefs = [];
      this._droppedSourceRefs = new Set();
      this._pruning = null;
      this._threadPlan = null;
      this._catalogDropped = null;
      this._buttonParents = null;
    }
  }

  /**
   * @internal - The catalog entries of a full rewrite other than those `_buildFullSave` writes
   * itself (see `catalog-rewrite.ts`). An unreadable catalog has none to carry over.
   */
  private _rewriteCatalogEntries(
    writer: PdfWriter,
    catalogObjNum: number,
    pagesTreeObjNum: number,
    pagesInfo: ReadonlyArray<{ objNum: number; gen: number }>,
    outputPages: ReadonlyArray<number | null>
  ): Array<[string, string]> {
    let catalog: PdfDictValue;
    try {
      catalog = this._doc.getCatalog();
    } catch {
      return [];
    }
    // A reference back to the catalog or to the root of the page tree — from a structure
    // element, say — is to the ones written here, not to a second copy of the originals.
    const rootRef = this._doc.trailer.get("Root");
    if (isPdfRef(rootRef)) {
      this._clonedRefs.set(`main:${rootRef.objNum}:${rootRef.gen}`, catalogObjNum);
    }
    const pagesRef = catalog.get("Pages");
    if (isPdfRef(pagesRef)) {
      this._clonedRefs.set(`main:${pagesRef.objNum}:${pagesRef.gen}`, pagesTreeObjNum);
    }
    this._catalogIndex ??= buildCatalogIndex(this._doc, catalog, pagesInfo);
    return rewriteCatalogEntries(
      {
        doc: this._doc,
        outputPages,
        split: this._edits.pages.kind === "only",
        threads: this._threadPlan,
        serialize: value => this._deepSerialize(this._doc, value),
        allocObject: () => writer.allocObject(),
        addObject: (objNum, body) => writer.addObject(objNum, body),
        mapRef: (ref, objNum) => this._clonedRefs.set(`main:${ref.objNum}:${ref.gen}`, objNum),
        dropRefs: isDropped => {
          this._catalogDropped = isDropped;
        }
      },
      catalog,
      this._catalogIndex
    );
  }

  /**
   * @internal - Write a page copied in with `copyPagesFrom`, with what `getPage` drew over it
   * and the rotation `rotatePage` set; returns its object number.
   */
  private _writeCopiedPage(
    writer: PdfWriter,
    copied: Extract<AppendedPageSnapshot, { kind: "copied" }>,
    pagesTreeObjNum: number,
    fontDictStr: string
  ): number {
    const contentRefs: number[] = [];
    for (const streamData of copied.contentStreams) {
      const objNum = writer.allocObject();
      writer.addStreamObject(objNum, new PdfDict(), streamData);
      contentRefs.push(objNum);
    }

    // Deep-clone resources from the source document, and add the overlay's beside them, as
    // for a page the document was read with.
    const overlay =
      copied.overlay !== undefined && hasOverlay(copied.overlay) ? copied.overlay : undefined;
    let overlayResourcesStr = "";
    if (overlay) {
      const overlayObjNum = writer.allocObject();
      writer.addStreamObject(overlayObjNum, new PdfDict(), overlay._stream);
      contentRefs.push(overlayObjNum);
      const imageXObjectMap = this._writeOverlayImages(writer, overlay);
      overlayResourcesStr = serializeResourceDict(
        this._buildOverlayResourceDict(writer, overlay, fontDictStr, imageXObjectMap)
      );
    }
    const resourcesStr = this._mergeResourceStrings(
      this._serializeResources(copied.sourceDoc, copied.sourcePageDict),
      overlayResourcesStr
    );
    const resourcesObjNum = writer.allocObject();
    writer.addObject(resourcesObjNum, resourcesStr || "<< >>");

    const contentsStr =
      contentRefs.length === 1
        ? pdfRef(contentRefs[0])
        : `[${contentRefs.map(r => pdfRef(r)).join(" ")}]`;

    const pageDict = new PdfDict()
      .set("Type", "/Page")
      .set("Parent", pdfRef(pagesTreeObjNum))
      .set("MediaBox", this._pageMediaBox(copied.sourceDoc, copied.sourcePageDict, copied))
      .set("Contents", contentsStr)
      .set("Resources", pdfRef(resourcesObjNum));

    // Preserve page properties from source page
    this._copyPageProperties(copied.sourcePageDict, pageDict, copied.sourceDoc);

    if (copied.rotation !== undefined) {
      if (copied.rotation === 0) {
        pageDict.delete("Rotate");
      } else {
        pageDict.set("Rotate", String(copied.rotation));
      }
    }

    const pageObjNum = writer.allocObject();
    const annots = this._copyPageAnnotations(copied, pageObjNum);
    if (overlay) {
      const added = writeBuilderAnnotations(writer, overlay._builderAnnotations);
      for (const field of overlay._formFields) {
        const written = writeFormFieldObjects(writer, field.options, pageObjNum);
        added.push(...written.annotRefs);
        this._addedFieldRefs.push(...written.fieldRefs);
      }
      if (added.length > 0) {
        const own = added.map(ref => pdfRef(ref)).join(" ");
        // The copied annotations are an array (`[…]`) or nothing.
        pageDict.set(
          "Annots",
          annots?.startsWith("[") ? `${annots.slice(0, -1)} ${own}]` : `[${own}]`
        );
      } else if (annots !== undefined) {
        pageDict.set("Annots", annots);
      }
    } else if (annots !== undefined) {
      pageDict.set("Annots", annots);
    }
    writer.addObject(pageObjNum, pageDict);
    return pageObjNum;
  }

  /**
   * @internal - The annotations of a copied page — links, comments, widgets — cloned from the
   * source, with the page's form fields added to this document's AcroForm (see
   * `_addedFieldRefs`). They used to be dropped, form fields included. Each copy gets objects
   * of its own; a reference to the page itself (`/P`) is to the copy, and one to another page
   * of the source (a link destination, a field's widget there) is written as `null`, rather
   * than bringing that page along unattached.
   */
  private _copyPageAnnotations(copied: CopiedPage, pageObjNum: number): string | undefined {
    const { sourceDoc, sourcePageDict, sourcePageRef } = copied;
    const annots = sourcePageDict.get("Annots");
    if (annots === undefined || annots === null) {
      return undefined;
    }
    const previousScope = this._cloneScope;
    this._cloneScope = `#${++this._copyCount}`;
    try {
      const docId = this._docId(sourceDoc);
      if (sourcePageRef !== undefined) {
        this._clonedRefs.set(`${docId}:${sourcePageRef.objNum}:${sourcePageRef.gen}`, pageObjNum);
      }
      for (const { objNum, gen } of sourceDoc.getPagesWithObjInfo()) {
        if (objNum !== sourcePageRef?.objNum || gen !== sourcePageRef.gen) {
          this._droppedSourceRefs.add(`${docId}:${objNum}:${gen}`);
        }
      }
      const resolved = sourceDoc.deref(annots);
      if (copied.formFields === "drop") {
        // Every annotation but the widgets: a widget is (part of) a form field.
        const kept = (isPdfArray(resolved) ? resolved : [])
          .filter(annot => {
            const dict = sourceDoc.derefDict(annot);
            return !dict || dictGetName(dict, "Subtype") !== "Widget";
          })
          .map(annot => this._deepSerialize(sourceDoc, annot))
          .filter(value => value !== "" && value !== "null");
        return kept.length === 0 ? undefined : `[${kept.join(" ")}]`;
      }
      // The resolved array rather than `/Annots` as written, so the result is always an inline
      // array — what drawing over the copy (`getPage`) adds its own annotations to.
      const serialized = this._deepSerialize(sourceDoc, resolved);
      // The root field of each widget, as cloned.
      const roots = new Set<number>();
      for (const annot of isPdfArray(resolved) ? resolved : []) {
        let ref: PdfObject = annot;
        let dict = sourceDoc.derefDict(annot);
        if (!dict || dictGetName(dict, "Subtype") !== "Widget") {
          continue;
        }
        const seen = new Set<PdfDictValue>();
        while (dict && !seen.has(dict) && isPdfRef(dict.get("Parent"))) {
          seen.add(dict);
          ref = dict.get("Parent")!;
          dict = sourceDoc.derefDict(ref);
        }
        if (isPdfRef(ref)) {
          const cloned = this._clonedRefs.get(`${docId}:${ref.objNum}:${ref.gen}`);
          if (cloned !== undefined) {
            roots.add(cloned);
          }
        }
      }
      this._addedFieldRefs.push(...roots);
      return serialized;
    } finally {
      this._cloneScope = previousScope;
    }
  }

  /** @internal - Write a page added with `addPage`; returns its object number. */
  private _writeNewPage(
    writer: PdfWriter,
    page: PdfPageSnapshot,
    pagesTreeObjNum: number,
    fontDictStr: string,
    rotation?: number
  ): number {
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

    // Materialise the page's alpha ExtGState entries. The builder
    // captured them via `_applyAlpha` while painting; without this
    // block the `gs` operators in the content stream would reference
    // undefined resources. Matches `document-builder.ts:1417-1430`.
    let gsDictStr = "";
    if (page._alphaValues.size > 0) {
      const gsEntries: string[] = [];
      for (const alpha of page._alphaValues) {
        const gsObjNum = writer.allocObject();
        const gsDict = new PdfDict()
          .set("Type", "/ExtGState")
          .set("ca", pdfNumber(alpha))
          .set("CA", pdfNumber(alpha));
        writer.addObject(gsObjNum, gsDict);
        gsEntries.push(`/${alphaGsName(alpha, this._overlayResourcePrefix)} ${pdfRef(gsObjNum)}`);
      }
      gsDictStr = `<< ${gsEntries.join(" ")} >>`;
    }

    // Annotations and form fields drawn on the page: they used to be dropped.
    const annotRefs = writeBuilderAnnotations(writer, page._builderAnnotations);
    for (const field of page._formFields) {
      const written = writeFormFieldObjects(writer, field.options, undefined);
      annotRefs.push(...written.annotRefs);
      this._addedFieldRefs.push(...written.fieldRefs);
    }

    const resourcesObjNum = writer.allocObject();
    let resStr = "<< ";
    if (fontDictStr) {
      resStr += `/Font ${fontDictStr} `;
    }
    if (xobjDictStr) {
      resStr += `/XObject ${xobjDictStr} `;
    }
    if (gsDictStr) {
      resStr += `/ExtGState ${gsDictStr} `;
    }
    resStr += ">>";
    writer.addObject(resourcesObjNum, resStr);

    return writer.addPage({
      parentRef: pagesTreeObjNum,
      width: page._width,
      height: page._height,
      contentsRef: contentObjNum,
      resourcesRef: resourcesObjNum,
      annotRefs,
      ...(rotation === undefined || rotation === 0 ? {} : { rotate: rotation })
    });
  }

  /** @internal Full rebuild implementation, extracted for try/finally cleanup. */
  private async _buildFullSave(writer: PdfWriter, legacyFormRules: boolean): Promise<Uint8Array> {
    // A page tree with no leaves is a file viewers refuse to open, so writing one is not a
    // faithful rendering of "remove every page" — it is a broken file. Refuse instead; the
    // caller can add or copy a page in first. An input that had no pages to begin with is
    // left alone: a rewrite of it changes nothing about its page count.
    const { pages, appendedPages } = this._edits;
    const keptPages = keptPageIndices(pages, this._pages.length);
    // `this._appendedPages` rather than the written list: a page added and then removed is a
    // page removed, as a read one is.
    const hadPages = this._pages.length > 0 || this._appendedPages.length > 0;
    if (keptPages.length === 0 && hadPages && appendedPages.length === 0) {
      throw new PdfStructureError(
        "Every page was removed and none was added: the document would have no pages. " +
          "Add or copy a page in before writing it."
      );
    }
    // A part of `splitPages` writes one page, so only that page's overlay is looked at; a
    // rewrite prepares every overlay, as it always has.
    const overlays =
      pages.kind === "only" ? keptPages.map(i => this._edits.overlays[i]) : this._edits.overlays;
    for (const page of overlays) {
      page._stream.runPreflight();
    }
    for (const appended of this._edits.appendedPages) {
      if (appended.kind === "new") {
        appended.page._stream.runPreflight();
      } else {
        appended.overlay?._stream.runPreflight();
      }
    }
    discoverDrawnTextFonts(this._fontManager, this._fontSettings);
    // Write font resources for any overlay content
    const fontObjectMap = await this._fontManager.writeFontResources(writer);
    const fontDictStr = this._fontManager.buildFontDictString(fontObjectMap);

    const pagesTreeObjNum = writer.allocObject();
    const pageObjNums: number[] = [];

    // Re-emit existing pages
    const pagesInfo = this._doc.getPagesWithObjInfo();
    const pageObjNumByIndex = this._preassignPageAndAnnotClones(writer, pagesInfo, keptPages);
    const outputPages: Array<number | null> = [...keptPages, ...appendedPages.map(() => null)];
    const keepsStructure = keepsStructureTree(pagesInfo.length, outputPages);
    this._pruning = this._planPruning(pagesInfo);
    this._threadPlan = this._planThreads(pagesInfo, outputPages);
    const prunedFields = this._pruning?.fields();
    this._writtenAnnotObjNums = new Set();
    for (const i of keptPages) {
      const { dict: pageDict } = pagesInfo[i];
      const dims = this._doc.resolvePageBox(pageDict) ?? {
        width: 612,
        height: 792
      };
      const overlay = this._edits.overlays[i];

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
      if (hasOverlay(overlay)) {
        const overlayObjNum = writer.allocObject();
        writer.addStreamObject(overlayObjNum, new PdfDict(), overlay._stream);
        allContentRefs.push(overlayObjNum);

        // Build overlay-specific resources (fonts, images, ExtGState) as
        // structured dict. ExtGState is required whenever the overlay
        // page builder captured non-opaque colours via `_applyAlpha`;
        // without it the emitted `gs` operators would reference
        // undefined resources.
        const imageXObjectMap = this._writeOverlayImages(writer, overlay);
        const overlayDict = this._buildOverlayResourceDict(
          writer,
          overlay,
          fontDictStr,
          imageXObjectMap
        );
        overlayResourcesStr = serializeResourceDict(overlayDict);
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

      // The page's number is needed now: the widgets of a field added to it name it as /P.
      const pageObjNum = pageObjNumByIndex.get(i) ?? writer.allocObject();

      // Write overlay builder annotations (Highlight, Text, FreeText, Stamp, etc.)
      annotRefs.push(...writeBuilderAnnotations(writer, overlay._builderAnnotations));

      // Write overlay form fields — the builder's writer, and into the AcroForm (see
      // `_addedFieldRefs`): a field written only as a widget was no field at all.
      for (const field of overlay._formFields) {
        const written = writeFormFieldObjects(writer, field.options, pageObjNum);
        annotRefs.push(...written.annotRefs);
        this._addedFieldRefs.push(...written.fieldRefs);
      }

      const pageDict2 = new PdfDict()
        .set("Type", "/Page")
        .set("Parent", pdfRef(pagesTreeObjNum))
        .set("MediaBox", this._pageMediaBox(this._doc, pageDict, dims))
        .set("Contents", contentsStr)
        .set("Resources", pdfRef(resourcesObjNum));

      // Preserve page-level properties from the original page dict
      this._copyPageProperties(pageDict, pageDict2);
      const beads = this._pageBeads(pageDict);
      if (beads) {
        pageDict2.set("B", beads);
      }
      // The page's key into the structure tree's /ParentTree, which is written only when every
      // original page is (see `catalog-rewrite.ts`).
      const structParents = this._doc.deref(pageDict.get("StructParents") ?? null);
      if (keepsStructure && typeof structParents === "number") {
        pageDict2.set("StructParents", pdfNumber(structParents));
      }

      // Apply rotation override if set
      const rotationOverride = this._edits.rotationOverrides.get(i);
      if (rotationOverride !== undefined) {
        if (rotationOverride === 0) {
          // Remove rotation (unset the key)
          pageDict2.delete("Rotate");
        } else {
          pageDict2.set("Rotate", String(rotationOverride));
        }
      }

      if (annotRefs.length > 0) {
        pageDict2.set("Annots", `[${annotRefs.map(r => pdfRef(r)).join(" ")}]`);
      }

      writer.addObject(pageObjNum, pageDict2);
      pageObjNums.push(pageObjNum);
    }

    // Write the added and copied pages, in the order they were added
    for (const appended of this._edits.appendedPages) {
      pageObjNums.push(
        appended.kind === "copied"
          ? this._writeCopiedPage(writer, appended, pagesTreeObjNum, fontDictStr)
          : this._writeNewPage(
              writer,
              appended.page,
              pagesTreeObjNum,
              fontDictStr,
              appended.rotation
            )
      );
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

    // Carry the AcroForm over whether or not a value changed: without it the widgets are
    // annotations that belong to no form, and a reader finds no fields at all. Its
    // /Fields resolve to the widgets written above (see `_preassignPageAndAnnotClones`),
    // so a filled value is what the form reports. Inlined rather than referenced so
    // it can be adjusted for edited values (see `_acroFormForEditedValues`).
    // Fields added with `addFormField` or brought by `copyPagesFrom` join the form's /Fields;
    // a document with no form gets one, set up as the builder sets up its own.
    const hasFormUpdates = this._edits.formFieldUpdates.size > 0;
    const addedFields = this._addedFieldRefs.map(ref => pdfRef(ref)).join(" ");
    const newAcroForm =
      addedFields === ""
        ? "<< /NeedAppearances true >>"
        : `<< /Fields [${addedFields}] /NeedAppearances true /DR << /Font << /Helv << /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >> >> >> /DA (/Helv 0 Tf 0 g) >>`;
    let original: PdfDictValue | null = null;
    try {
      original = this._doc.derefDict(this._doc.getCatalog().get("AcroForm") ?? null);
    } catch {
      original = null;
    }
    if (original) {
      const acroForm = hasFormUpdates
        ? this._acroFormForEditedValues(original, legacyFormRules)
        : original;
      const fieldsValue = prunedFields ?? acroForm.get("Fields");
      const raw = new Map<string, string>();
      if (addedFields !== "") {
        const fields = this._doc.deref(fieldsValue ?? null);
        const kept = (isPdfArray(fields) ? fields : [])
          .map(field => this._deepSerialize(this._doc, field))
          .filter(field => field !== "" && field !== "null");
        raw.set("Fields", `[${[...kept, addedFields].join(" ")}]`);
      }
      catalogDict.set(
        "AcroForm",
        this._serializeDictToWriter(
          this._doc,
          acroForm,
          prunedFields ? new Map([["Fields", prunedFields]]) : undefined,
          raw
        )
      );
    } else if (hasFormUpdates || addedFields !== "") {
      catalogDict.set("AcroForm", newAcroForm);
    }

    // The catalog is rebuilt, so /NeedsRendering is carried over by hand: a dynamic XFA
    // form keeps its XFA (it is never edited — see `_guardDynamicXfa`) and needs the flag
    // that tells a viewer to render from it.
    try {
      const needsRendering = this._doc.deref(this._doc.getCatalog().get("NeedsRendering") ?? null);
      if (typeof needsRendering === "boolean") {
        catalogDict.set("NeedsRendering", needsRendering ? "true" : "false");
      }
    } catch {
      // An unreadable catalog has no flag to carry over.
    }

    // Preserve XMP metadata stream from original catalog — updated to agree with /Info when
    // `setMetadata` changed the properties both carry.
    const updatedXmp =
      this._edits.metadata === undefined
        ? undefined
        : this._updatedXmp(this._edits.metadata, value => this._deepSerialize(this._doc, value));
    if (updatedXmp !== undefined) {
      const xmpObjNum = writer.allocObject();
      writer.addStreamObject(xmpObjNum, updatedXmp.dict, updatedXmp.data, { compress: false });
      catalogDict.set("Metadata", pdfRef(xmpObjNum));
    } else {
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
    }

    // Every other document-level entry: bookmarks, named destinations, page labels, optional
    // content, output intents, the language… (see `catalog-rewrite.ts` for which are adjusted
    // to the pages written and which are dropped).
    for (const [key, value] of this._rewriteCatalogEntries(
      writer,
      catalogObjNum,
      pagesTreeObjNum,
      pagesInfo,
      outputPages
    )) {
      catalogDict.set(key, value);
    }

    writer.addObject(catalogObjNum, catalogDict);
    writer.setCatalog(catalogObjNum);

    // Inject signature placeholder if signing is in progress
    if (this._signaturePlaceholder) {
      const sigDictObjNum = writer.allocObject();
      writer.addObject(sigDictObjNum, this._signaturePlaceholder);

      const sigWidgetObjNum = writer.allocObject();
      const sigWidgetDict = new PdfDict()
        .set("Type", "/Annot")
        .set("Subtype", "/Widget")
        .set("FT", "/Sig")
        .set("Rect", "[0 0 0 0]")
        .set("T", pdfString("Signature1"))
        .set("V", pdfRef(sigDictObjNum))
        .set("F", "4");
      writer.addObject(sigWidgetObjNum, sigWidgetDict);

      // Patch catalog to include signature widget in AcroForm with SigFlags
      // Grab existing fields from the catalog dict string representation
      const catalogStr = catalogDict.toString();
      const existingFieldsMatch = catalogStr.match(/\/Fields\s*\[([^\]]*)\]/);
      const existingFields = existingFieldsMatch ? existingFieldsMatch[1].trim() : "";
      const sigFieldRef = pdfRef(sigWidgetObjNum);
      const allFieldRefs = existingFields ? `${existingFields} ${sigFieldRef}` : sigFieldRef;
      catalogDict.set("AcroForm", `<< /Fields [${allFieldRefs}] /SigFlags 3 >>`);
      writer.addObject(catalogObjNum, catalogDict);
    }

    // Info dict — every entry the document has, with what `setMetadata` set. Rebuilding it from
    // the five fields the builder knows lost /Keywords, /ModDate, /Trapped and every custom
    // entry on every rewrite.
    const infoObjNum = writer.allocObject();
    const infoEntries = this._infoEntries(this._edits.metadata, value =>
      this._deepSerialize(this._doc, value)
    );
    writer.addInfoObject(
      infoObjNum,
      `<< ${[...infoEntries].map(([key, value]) => `${pdfNameBytes(key)} ${value}`).join(" ")} >>`
    );

    this._throwIfLimitExceeded();
    return writer.build();
  }

  /**
   * Save the modified PDF using incremental update.
   *
   * Appends new/modified objects after the original PDF bytes, preserving the
   * original data byte-for-byte. This is ideal for overlays and form field
   * updates on existing pages — it preserves digital signatures on unmodified
   * content and produces smaller output.
   *
   * Falls back to a full rebuild if structural changes are
   * present (new pages, copied pages, removed pages, or rotation changes).
   * On a signed document that fallback throws
   * `PdfSignatureInvalidationError` unless `invalidateSignatures` is set.
   *
   * Files whose cross-reference data is a stream are updated with an xref
   * stream, as ISO 32000-2 §7.5.8 expects. Hybrid-reference files (a classic
   * trailer with `/XRefStm`) get a classic section whose `/Prev` names the
   * original trailer; readers reach the compressed objects through that
   * trailer's `/XRefStm` (ISO 32000-1 §7.5.8.4).
   *
   * If the reader had to reconstruct the cross-reference data
   * (`PdfDocument.xrefRecovered` — a wrong or damaged `startxref`, or an entry that does not
   * point at the object it names), there is
   * no reliable section to chain `/Prev` to, so this also falls back to
   * a full rebuild, with the same signature guard.
   *
   * @returns The modified PDF as Uint8Array
   */
  private async _writeIncremental(
    options: EditorWriteOptions | undefined,
    call: WriteCall,
    required = false
  ): Promise<Uint8Array> {
    this._guardDynamicXfa(call);
    // The original bytes stay encrypted, so only a rewrite can drop the encryption.
    if (options?.encryption === "remove" && isEncrypted(this._doc)) {
      return this._fallBackToFull(options, call, "to remove the encryption", required);
    }
    // Fall back to full rebuild if structural changes are present
    if (
      this._edits.appendedPages.length > 0 ||
      removesAnyPage(this._edits.pages, this._pages.length) ||
      this._edits.rotationOverrides.size > 0
    ) {
      return this._fallBackToFull(
        options,
        call,
        "for page additions, removals or rotations",
        required
      );
    }
    // An update writes a page's new content and resources but not new annotations or form
    // fields — those were silently dropped — so a page that gains either is rewritten.
    if (
      this._edits.overlays.some(
        page => page._builderAnnotations.length > 0 || page._formFields.length > 0
      )
    ) {
      return this._fallBackToFull(
        options,
        call,
        "for annotations or form fields added to a page",
        required
      );
    }

    // Check if there are any modifications at all
    const hasOverlays = this._edits.overlays.some(hasOverlay);
    const hasFormUpdates = this._edits.formFieldUpdates.size > 0;

    if (!hasOverlays && !hasFormUpdates && this._edits.metadata === undefined) {
      // No modifications — return the original bytes
      return this._doc.data;
    }

    // A reconstructed xref means the file's own startxref chain is broken:
    // an update appended to it would point /Prev at a damaged section.
    if (this._doc.xrefRecovered || this._doc.startxrefOffset === null) {
      return this._fallBackToFull(
        options,
        call,
        "because the cross-reference data was recovered",
        required
      );
    }

    this._isIncrementalSave = true;
    try {
      return await withDrawnTextFontBuild(this._fontManager, () =>
        this._buildIncrementalUpdate(options, call, required)
      );
    } finally {
      this._isIncrementalSave = false;
      this._writerForSave = null;
      this._clonedRefs.clear();
    }
  }

  /**
   * @internal - The AcroForm a file with edited field values should carry:
   *
   * - `/NeedAppearances true`, set even where the original said `false`. A text field
   *   gets a fresh appearance stream, but other fields have their stale /AP removed, and
   *   only this flag tells a viewer to draw them from the new value.
   * - no `/XFA`. Only reached for a hybrid form — AcroForm fields plus an XFA packet,
   *   `/NeedsRendering` not true (`_guardDynamicXfa` refuses the dynamic kind first). Its
   *   AcroForm fields carry the data being edited, but an XFA-aware viewer renders from
   *   the XFA datasets instead and would show the old values, or overwrite the new ones
   *   when it syncs back. Dropping the packet makes the AcroForm authoritative — what form
   *   tools generally do when they fill only the AcroForm, and XFA is deprecated in
   *   ISO 32000-2 anyway. A file with no edited values keeps its XFA untouched.
   */
  private _acroFormForEditedValues(acroForm: PdfDictValue, keepXfa = false): PdfDictValue {
    const out: PdfDictValue = new Map([["NeedAppearances", true]]);
    for (const [key, value] of acroForm) {
      if (key !== "NeedAppearances" && (keepXfa || key !== "XFA")) {
        out.set(key, value);
      }
    }
    return out;
  }

  /**
   * @internal - Whether the document is a dynamic XFA form, whose XFA is the form's only
   * real definition and so must never be stripped. Any one of these says so:
   *
   * - `/NeedsRendering true` in the catalog — the flag ISO 32000 defines for it;
   * - an XFA packet with no AcroForm field tree beside it;
   * - the XFA config packet's `<dynamicRender>required</dynamicRender>` — what an XFA
   *   designer writes for a dynamic form, and the marker a file that lost its
   *   `/NeedsRendering` but kept placeholder AcroForm fields still carries.
   *
   * A dynamic form that has none of the three — no flag, placeholder fields, and no
   * `dynamicRender` — is indistinguishable from a hybrid one without interpreting the
   * template, and is treated as hybrid.
   */
  private _isDynamicXfa(): boolean {
    let catalog: PdfDictValue;
    try {
      catalog = this._doc.getCatalog();
    } catch {
      return false;
    }
    if (this._doc.deref(catalog.get("NeedsRendering") ?? null) === true) {
      return true;
    }
    const acroForm = this._doc.derefDict(catalog.get("AcroForm") ?? null);
    if (!acroForm?.has("XFA")) {
      return false;
    }
    const fields = this._doc.deref(acroForm.get("Fields") ?? null);
    if (!isPdfArray(fields) || fields.length === 0) {
      return true;
    }
    return this._xfaRequiresDynamicRender(acroForm.get("XFA") ?? null);
  }

  /** @internal - Whether an /XFA value's packets declare `dynamicRender` `required`. */
  private _xfaRequiresDynamicRender(xfa: PdfObject): boolean {
    const resolved = this._doc.deref(xfa);
    // A single stream holds every packet; an array alternates packet names and streams.
    const parts = isPdfArray(resolved) ? resolved.filter((_, i) => i % 2 === 1) : [xfa];
    const latin1 = new TextDecoder("latin1");
    for (const part of parts) {
      try {
        const found = isPdfRef(part) ? this._doc.derefStreamWithObjNum(part) : null;
        if (!found) {
          continue;
        }
        const text = latin1.decode(this._doc.getStreamData(found.stream, found.objNum, found.gen));
        if (/<dynamicRender\b[^>]*>\s*required\s*<\/dynamicRender>/.test(text)) {
          return true;
        }
      } catch {
        // An unreadable packet declares nothing.
      }
    }
    return false;
  }

  /**
   * @internal - Refuse to save edited field values into a dynamic XFA form. Updating the
   * XFA datasets instead was considered and rejected: binding a value to a datasets node
   * needs the template's binding rules (`bind ref`, `match="global"`, repeating subforms),
   * and a name-only guess would write data the form does not read. Nothing is written.
   */
  private _guardDynamicXfa({ caller, legacyFormRules }: WriteCall): void {
    if (!legacyFormRules && this._edits.formFieldUpdates.size > 0 && this._isDynamicXfa()) {
      throw new PdfXfaFormError(
        `${caller}: this is a dynamic XFA form (its XFA packet is the form's only definition), ` +
          "and form values can only be filled into AcroForm fields, which an XFA viewer " +
          "ignores. Nothing was written. Fill it with an XFA-capable tool, or flatten it first."
      );
    }
  }

  /**
   * @internal - Append the AcroForm adjustments of `_acroFormForEditedValues` to an
   * incremental update: a new revision of the AcroForm object when it is indirect,
   * otherwise of the catalog that holds it inline. Nothing is appended when the
   * AcroForm already says `/NeedAppearances true` and has no `/XFA`.
   */
  private _appendAcroFormForEditedValues(modifiedObjects: Map<number, IncrementalObject>): void {
    const rootRef = this._doc.trailer.get("Root");
    if (!isPdfRef(rootRef)) {
      return;
    }
    const catalog = this._doc.derefDict(rootRef);
    const raw = catalog?.get("AcroForm");
    const acroForm = raw === undefined ? null : this._doc.derefDict(raw);
    if (!catalog || !acroForm) {
      return;
    }
    if (acroForm.get("NeedAppearances") === true && !acroForm.has("XFA")) {
      return;
    }
    const serialize = (dict: PdfDictValue): string => {
      const parts = ["<<"];
      for (const [key, value] of dict) {
        parts.push(`${pdfNameBytes(key)} ${this._serializeOriginalValue(value)}`);
      }
      parts.push(">>");
      return parts.join(" ");
    };
    const adjusted = this._acroFormForEditedValues(acroForm);
    if (isPdfRef(raw)) {
      modifiedObjects.set(raw.objNum, { gen: raw.gen, body: serialize(adjusted) });
      return;
    }
    const newCatalog: PdfDictValue = new Map(catalog);
    newCatalog.set("AcroForm", adjusted);
    modifiedObjects.set(rootRef.objNum, { gen: rootRef.gen, body: serialize(newCatalog) });
  }

  /**
   * @internal - Rewrite the whole file because an update cannot be appended — or, when
   * the caller required an update, refuse.
   */
  private async _fallBackToFull(
    options: EditorWriteOptions | undefined,
    call: WriteCall,
    why: string,
    required: boolean
  ): Promise<Uint8Array> {
    const { caller } = call;
    if (required) {
      throw new PdfIncrementalUpdateError(
        `${caller} cannot append an update ${why}; nothing was written. ` +
          "Pass { incremental: true } to accept a full rewrite instead."
      );
    }
    this._guardSignatures(options, `${caller} must fall back to a full rewrite ${why}`);
    this._isIncrementalSave = false;
    return this._writeFull({ ...options, invalidateSignatures: true }, call);
  }

  /** @internal — Core incremental update logic, separated for try/finally cleanup. */
  private async _buildIncrementalUpdate(
    options: EditorWriteOptions | undefined,
    call: WriteCall,
    required: boolean
  ): Promise<Uint8Array> {
    const { caller } = call;
    // Determine the next available object number from the original PDF's /Size
    let nextObjNum = this._doc.objectNumberBound;

    // Collect modified objects: objNum → generation and serialized content
    const modifiedObjects = new Map<number, IncrementalObject>();

    // Check what kinds of modifications exist
    const hasOverlays = this._edits.overlays.some(hasOverlay);

    // We need a temporary PdfWriter for font resources used in overlays
    const writer = new PdfWriter();
    this._writerForSave = writer;
    this._clonedRefs = new Map();

    let fontDictStr = "";
    // Map of writer objNum → actual new objNum we'll use in the incremental update
    const writerFontObjRemap = new Map<number, number>();

    if (hasOverlays) {
      for (const page of this._edits.overlays) {
        page._stream.runPreflight();
      }
      discoverDrawnTextFonts(this._fontManager, this._fontSettings);
      // Write font resources via the writer (to serialize font objects)
      const fontObjectMap = await this._fontManager.writeFontResources(writer);

      // Remap all writer-allocated objects (fonts + their dependencies like
      // CID font descriptors, ToUnicode CMaps, etc.) into the incremental
      // update's object number space.
      const writerObjects = writer.getObjects();
      for (const obj of writerObjects) {
        if (!writerFontObjRemap.has(obj.objectNumber)) {
          writerFontObjRemap.set(obj.objectNumber, nextObjNum++);
        }
      }

      // Write all font objects into modifiedObjects with remapped refs
      for (const obj of writerObjects) {
        const newObjNum = writerFontObjRemap.get(obj.objectNumber)!;
        const remappedContent = this._remapRefsInString(obj.content, writerFontObjRemap);
        if (obj.streamData) {
          // Parse the remapped content back into a PdfDict for stream objects
          modifiedObjects.set(newObjNum, {
            gen: 0,
            body: { dict: PdfDict.fromRawString(remappedContent), data: obj.streamData }
          });
        } else {
          modifiedObjects.set(newObjNum, { gen: 0, body: remappedContent });
        }
      }

      // Build a remapped font dict string
      const rawFontDict = this._fontManager.buildFontDictString(fontObjectMap);
      fontDictStr = this._remapRefsInString(rawFontDict, writerFontObjRemap);
    }

    const pagesInfo = this._doc.getPagesWithObjInfo();

    for (let i = 0; i < pagesInfo.length; i++) {
      const overlay = this._edits.overlays[i];
      const { dict: pageDict, objNum: pageObjNum, gen: pageGen } = pagesInfo[i];

      if (pageObjNum === 0) {
        // Can't do incremental update without knowing the page object number.
        // Fall back to full rebuild (finally block handles cleanup), under
        // the same signature guard as every other fallback.
        return this._fallBackToFull(options, call, "for a page without an object number", required);
      }

      const pageHasOverlay = hasOverlay(overlay);
      const pageHasFormUpdates = this._hasFormUpdatesForPage(pageDict);

      if (!pageHasOverlay && !pageHasFormUpdates) {
        continue; // Page is unchanged
      }

      // Build the updated page dictionary. We start from the original page
      // dict's entries and modify only what's necessary.
      const updatedPageDict = new PdfDict();

      // Copy all existing entries from the original page dict
      for (const [key, val] of pageDict.entries()) {
        if (key === "Contents" && pageHasOverlay) {
          continue; // We'll rewrite /Contents below
        }
        if (key === "Resources" && pageHasOverlay) {
          continue; // We'll rewrite /Resources below
        }
        if (key === "Annots" && pageHasFormUpdates) {
          continue; // We'll rewrite /Annots below
        }
        updatedPageDict.set(key, this._serializeOriginalValue(val));
      }

      if (pageHasOverlay) {
        // Create a new content stream object for the overlay
        const overlayStreamData = overlay._stream.toUint8Array();
        const overlayObjNum = nextObjNum++;
        modifiedObjects.set(overlayObjNum, {
          gen: 0,
          body: { dict: new PdfDict(), data: overlayStreamData }
        });

        // Build the new /Contents array: original refs + overlay ref
        const originalContentsObj = pageDict.get("Contents");
        const contentRefs = this._collectOriginalContentRefs(originalContentsObj);
        contentRefs.push(pdfRef(overlayObjNum));

        const contentsStr =
          contentRefs.length === 1 ? contentRefs[0] : `[${contentRefs.join(" ")}]`;
        updatedPageDict.set("Contents", contentsStr);

        // Build merged resources: original + overlay fonts/images
        const originalResources = this._serializeOriginalResources(this._doc, pageDict);

        // For overlay images, write them as new objects
        const imageObjMap = new Map<string, number>();
        // The same writer as a full rewrite — and so the same soft mask and colour space —
        // over this update's object table.
        const sink: ImageObjectSink = {
          allocate: () => nextObjNum++,
          addStream: (objNum, dict, data) => {
            modifiedObjects.set(objNum, { gen: 0, body: { dict, data } });
          }
        };
        for (let imgIdx = 0; imgIdx < overlay._images.length; imgIdx++) {
          const img = overlay._images[imgIdx];
          const imgName = `${this._overlayResourcePrefix}Im${imgIdx + 1}`;
          imageObjMap.set(imgName, addImageXObject(sink, img.data, img.format));
        }

        // Build overlay resource string
        let overlayStr = "<< ";
        if (fontDictStr) {
          overlayStr += `/Font ${fontDictStr} `;
        }
        if (imageObjMap.size > 0) {
          const entries = [...imageObjMap.entries()]
            .map(([name, objNum]) => `/${name} ${pdfRef(objNum)}`)
            .join(" ");
          overlayStr += `/XObject << ${entries} >> `;
        }
        overlayStr += ">>";

        const mergedResources = this._mergeResourceStrings(originalResources, overlayStr);
        // Write merged resources as a new object
        const resourcesObjNum = nextObjNum++;
        modifiedObjects.set(resourcesObjNum, { gen: 0, body: mergedResources || "<< >>" });
        updatedPageDict.set("Resources", pdfRef(resourcesObjNum));
      }

      if (pageHasFormUpdates) {
        // Build updated annotations list
        const annotsResult = this._buildIncrementalAnnots(pageDict, modifiedObjects, nextObjNum);
        nextObjNum = annotsResult.nextObjNum;
        if (annotsResult.annotRefs.length > 0) {
          updatedPageDict.set(
            "Annots",
            `[${annotsResult.annotRefs.map(r => pdfRef(r.objNum, r.gen)).join(" ")}]`
          );
        }
      }

      // Write the updated page dict as a modified object (same object number
      // and generation as the original, which every /Kids entry refers to)
      modifiedObjects.set(pageObjNum, { gen: pageGen, body: updatedPageDict.toString() });
    }

    for (const { ref, dict, value } of this._buttonParentValues().values()) {
      const updated: PdfDictValue = new Map(dict);
      updated.set("V", value);
      const parts = ["<<"];
      for (const [key, val] of updated) {
        parts.push(`${pdfNameBytes(key)} ${this._serializeOriginalValue(val)}`);
      }
      parts.push(">>");
      modifiedObjects.set(ref.objNum, { gen: ref.gen, body: parts.join(" ") });
    }

    if (
      modifiedObjects.size > 0 &&
      this._edits.formFieldUpdates.size > 0 &&
      !call.legacyFormRules
    ) {
      this._appendAcroFormForEditedValues(modifiedObjects);
    }

    const trailerEntries = this._serializeTrailerEntries();
    const metadata = this._edits.metadata;
    if (metadata !== undefined) {
      // The information dictionary: rewritten at its own number, or added when the file
      // has none (or an inline one, which an update cannot amend).
      const infoRef = this._doc.trailer.get("Info");
      const target = isPdfRef(infoRef) ? infoRef : { objNum: nextObjNum++, gen: 0 };
      const entries = this._infoEntries(metadata, value => this._serializeOriginalValue(value));
      modifiedObjects.set(target.objNum, {
        gen: target.gen,
        body: `<< ${[...entries].map(([key, value]) => `${pdfNameBytes(key)} ${value}`).join(" ")} >>`
      });
      trailerEntries.set("Info", pdfRef(target.objNum, target.gen));
      const xmp = this._updatedXmp(metadata, value => this._serializeOriginalValue(value));
      if (xmp !== undefined && xmp.ref !== undefined) {
        modifiedObjects.set(xmp.ref.objNum, {
          gen: xmp.ref.gen,
          body: { dict: xmp.dict, data: xmp.data }
        });
      }
    }

    this._throwIfLimitExceeded();
    if (modifiedObjects.size === 0) {
      return this._doc.data;
    }

    const prevXrefOffset = this._doc.startxrefOffset;
    if (prevXrefOffset === null) {
      throw new PdfStructureError(`${caller}: the original startxref offset is unknown`);
    }
    return buildIncremental(this._doc.data, modifiedObjects, trailerEntries, {
      xrefStream: this._doc.trailer.get("Type") === "XRef",
      prevXrefOffset,
      // An encrypted file's appended objects are encrypted with its own handler.
      encrypt: this._doc.encryptFn ?? undefined,
      encryptMetadata: this._doc.encryptMetadata
    });
  }

  /**
   * Check if a page has form field updates in its annotations.
   * @internal
   */
  private _hasFormUpdatesForPage(pageDict: PdfDictValue): boolean {
    if (this._edits.formFieldUpdates.size === 0) {
      return false;
    }

    const annotsObj = pageDict.get("Annots");
    if (!annotsObj) {
      return false;
    }

    const annotsResolved = this._doc.deref(annotsObj);
    if (!isPdfArray(annotsResolved)) {
      return false;
    }

    for (const annotRef of annotsResolved) {
      const annotDict = this._doc.derefDict(annotRef);
      if (!annotDict) {
        continue;
      }

      const subtype = dictGetName(annotDict, "Subtype");
      if (subtype === "Widget") {
        const fieldName = this._resolveFieldName(annotDict);
        if (fieldName && this._edits.formFieldUpdates.has(fieldName)) {
          return true;
        }
      }
    }

    return false;
  }

  /**
   * Collect original /Contents refs as serialized ref strings.
   * For incremental update — preserves original object references.
   * @internal
   */
  private _collectOriginalContentRefs(contentsObj: PdfObject | undefined): string[] {
    if (!contentsObj) {
      return [];
    }

    if (isPdfRef(contentsObj)) {
      // Could be a single ref to a stream, or a ref to an array
      const resolved = this._doc.deref(contentsObj);
      if (isPdfArray(resolved)) {
        return resolved
          .filter((item): item is PdfRef => isPdfRef(item))
          .map(item => pdfRef(item.objNum, item.gen));
      }
      return [pdfRef(contentsObj.objNum, contentsObj.gen)];
    }

    if (isPdfArray(contentsObj)) {
      return contentsObj
        .filter((item): item is PdfRef => isPdfRef(item))
        .map(item => pdfRef(item.objNum, item.gen));
    }

    return [];
  }

  /**
   * Serialize a page's Resources dict preserving original object references.
   * Unlike _serializeResources which deep-clones into the writer, this keeps
   * the original object numbers intact for incremental updates.
   * @internal
   */
  private _serializeOriginalResources(doc: PdfDocument, pageDict: PdfDictValue): string {
    const resourcesDict = doc.resolvePageResources(pageDict);
    if (!resourcesDict || resourcesDict.size === 0) {
      return "<< >>";
    }

    const parts: string[] = ["<<"];
    for (const [key, val] of resourcesDict.entries()) {
      const serialized = this._serializeOriginalValue(val);
      if (serialized) {
        parts.push(`${pdfNameBytes(key)} ${serialized}`);
      }
    }
    parts.push(">>");
    return parts.join(" ");
  }

  /**
   * Serialize a PDF value from the original document, preserving original refs.
   * Unlike _deepSerialize which clones refs into a new writer, this keeps
   * the original object numbers intact.
   * @internal
   */
  private _serializeOriginalValue(val: PdfObject): string {
    if (val === null || val === undefined) {
      return "null";
    }
    if (typeof val === "string") {
      // In parsed PDF objects, string type = PDF name (without leading /)
      // Uint8Array = PDF string literal
      return pdfNameBytes(val);
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
      return pdfRef(val.objNum, val.gen);
    }
    if (isPdfArray(val)) {
      const items = val.map(item => this._serializeOriginalValue(item));
      return `[${items.join(" ")}]`;
    }
    if (val instanceof Map) {
      const parts: string[] = ["<<"];
      for (const [k, v] of (val as PdfDictValue).entries()) {
        parts.push(`${pdfNameBytes(k)} ${this._serializeOriginalValue(v)}`);
      }
      parts.push(">>");
      return parts.join(" ");
    }
    return "";
  }

  /**
   * Replace object number references in a serialized string.
   * Maps old obj numbers (from temporary writer) to new obj numbers.
   * @internal
   */
  private _remapRefsInString(str: string, remap: Map<number, number>): string {
    // Replace patterns like "N G R" where N is in the remap map.
    // Linear scan avoids polynomial backtracking on untrusted PDF content.
    const parts: string[] = [];
    let last = 0;
    let i = 0;
    const len = str.length;

    while (i < len) {
      // Look for a digit
      if (str.charCodeAt(i) < 0x30 || str.charCodeAt(i) > 0x39) {
        i++;
        continue;
      }
      // Parse first number (object number)
      const numStart = i;
      while (i < len && str.charCodeAt(i) >= 0x30 && str.charCodeAt(i) <= 0x39) {
        i++;
      }
      // Must be followed by " <digits> R"
      if (i >= len || str.charCodeAt(i) !== 0x20) {
        continue;
      }
      const genStart = i + 1;
      let g = genStart;
      while (g < len && str.charCodeAt(g) >= 0x30 && str.charCodeAt(g) <= 0x39) {
        g++;
      }
      if (g === genStart || g >= len || str.charCodeAt(g) !== 0x20) {
        continue;
      }
      if (g + 1 >= len || str.charCodeAt(g + 1) !== 0x52 /* 'R' */) {
        continue;
      }
      // Ensure R is not followed by another word char (boundary check)
      const afterR = g + 2;
      if (
        afterR < len &&
        str.charCodeAt(afterR) > 0x20 &&
        str.charCodeAt(afterR) !== 0x2f /* '/' */ &&
        str.charCodeAt(afterR) !== 0x3e /* '>' */
      ) {
        continue;
      }

      const objNum = parseInt(str.slice(numStart, i), 10);
      const remapped = remap.get(objNum);
      if (remapped !== undefined) {
        parts.push(str.slice(last, numStart));
        parts.push(`${remapped} ${str.slice(genStart, g)} R`);
        last = afterR;
      }
      i = afterR;
    }

    if (last === 0) {
      return str;
    }
    parts.push(str.slice(last));
    return parts.join("");
  }

  /**
   * Build updated annotations for incremental save.
   * Modified widgets get rewritten at their original object number;
   * unmodified annots keep original refs.
   * @internal
   */
  private _buildIncrementalAnnots(
    pageDict: PdfDictValue,
    modifiedObjects: Map<number, IncrementalObject>,
    nextObjNum: number
  ): { annotRefs: Array<{ objNum: number; gen: number }>; nextObjNum: number } {
    const annotsObj = pageDict.get("Annots");
    if (!annotsObj) {
      return { annotRefs: [], nextObjNum };
    }

    const annotsResolved = this._doc.deref(annotsObj);
    if (!isPdfArray(annotsResolved)) {
      return { annotRefs: [], nextObjNum };
    }

    const annotRefs: Array<{ objNum: number; gen: number }> = [];

    for (const annotRef of annotsResolved) {
      const annotDict = this._doc.derefDict(annotRef);
      if (!annotDict) {
        // Keep original ref if we can't resolve
        if (isPdfRef(annotRef)) {
          annotRefs.push(annotRef);
        }
        continue;
      }

      const subtype = dictGetName(annotDict, "Subtype");

      if (subtype === "Widget" && this._edits.formFieldUpdates.size > 0) {
        const fieldName = this._resolveFieldName(annotDict);
        const newValue = fieldName ? this._edits.formFieldUpdates.get(fieldName) : undefined;

        if (newValue !== undefined) {
          // Rewrite the annotation at its original object number and generation
          const target = isPdfRef(annotRef) ? annotRef : { objNum: nextObjNum++, gen: 0 };
          // The appearance stream is an object of this update, numbered in
          // the original file's space — not in the scratch writer's, whose
          // numbers collide with the original objects.
          const newDict = this._buildModifiedWidgetDict(annotDict, newValue, (dict, data) => {
            const apObjNum = nextObjNum++;
            modifiedObjects.set(apObjNum, { gen: 0, body: { dict, data } });
            return apObjNum;
          });
          modifiedObjects.set(target.objNum, { gen: target.gen, body: newDict });
          annotRefs.push(target);
          continue;
        }
      }

      // Keep original annotation reference
      if (isPdfRef(annotRef)) {
        annotRefs.push(annotRef);
      } else {
        // Inline annotation — write as new object
        const annotObjNum = nextObjNum++;
        const serialized = this._serializeAnnotDict(annotDict);
        modifiedObjects.set(annotObjNum, { gen: 0, body: serialized });
        annotRefs.push({ objNum: annotObjNum, gen: 0 });
      }
    }

    return { annotRefs, nextObjNum };
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

  /**
   * @internal - Copy preserved page properties from source to target dict. `Rotate` and
   * `CropBox` are inheritable (ISO 32000-2 §7.7.3.4): a page that does not carry them takes
   * them from the page tree, and the written page, whose parent is a new page tree, has to
   * carry them itself — or a landscape-by-inheritance document came back portrait.
   */
  private _copyPageProperties(source: PdfDictValue, target: PdfDict, doc?: PdfDocument): void {
    const resolveDoc = doc ?? this._doc;
    for (const key of PdfEditor._PAGE_PRESERVE_KEYS) {
      const val =
        key === "Rotate" || key === "CropBox"
          ? inheritedPageEntry(resolveDoc, source, key)
          : source.get(key);
      if (val !== undefined && val !== null) {
        target.set(key, this._deepSerialize(resolveDoc, val));
      }
    }
  }

  /**
   * @internal - The page's `/MediaBox` as the document has it — inherited through the page
   * tree, origin included. Writing `[0 0 width height]` moved every page whose box does not
   * start at the origin (pdfcrop and LaTeX output, cropped scans): its content kept its
   * coordinates and the box moved under it. `[0 0 width height]` remains the fallback for a
   * page with no readable box.
   */
  private _pageMediaBox(
    doc: PdfDocument,
    pageDict: PdfDictValue,
    dims: { width: number; height: number }
  ): string {
    const box = doc.deref(inheritedPageEntry(doc, pageDict, "MediaBox") ?? null);
    if (
      isPdfArray(box) &&
      box.length === 4 &&
      box.every(value => typeof doc.deref(value) === "number" && Number.isFinite(doc.deref(value)))
    ) {
      return `[${box.map(value => pdfNumber(doc.deref(value) as number)).join(" ")}]`;
    }
    return `[0 0 ${pdfNumber(dims.width)} ${pdfNumber(dims.height)}]`;
  }

  /**
   * Sign this PDF with a digital signature.
   *
   * Performs a full save with an embedded PKCS#7 signature placeholder,
   * then fills in the real CMS SignedData. Because the save is a full
   * rewrite, signing an already-signed document throws
   * `PdfSignatureInvalidationError` unless `invalidateSignatures` is set.
   *
   * This is a write, like {@link toBuffer}: it resolves to the signed bytes and leaves the
   * editor unsigned, so a later `toBuffer()` writes without a signature. `Pdf.Builder#sign`
   * is different in kind: it records the signature, returns the builder for chaining, and
   * signs during that builder's `toBuffer()`.
   *
   * @param options - Certificate, private key, and optional signer metadata
   * @returns The signed PDF as Uint8Array
   *
   * @example
   * ```typescript
   * const editor = Pdf.Editor.read(pdfBytes);
   * const signed = await editor.sign({
   *   certificate: certDerBytes,
   *   privateKey: pkcs8DerBytes,
   *   name: "Jane Doe",
   *   reason: "Approval"
   * });
   * ```
   */
  async sign(options: EditorSignOptions): Promise<Uint8Array> {
    checkSignatureOptions("Pdf.Editor#sign", options);
    checkWriteOptions("Pdf.Editor#sign", options);
    // The signature is applied through a full rewrite, which breaks any
    // signature already present.
    this._guardSignatures(options, "sign() rewrites the whole file");

    // Frozen now, like every write's edits: an edit made while the signing module loads
    // belongs to the next write, not to this signature.
    const { pdfWithPlaceholder, signPdf } = await this._serialized(async () => {
      // **Load-bearing, despite rolldown calling it `INEFFECTIVE_DYNAMIC_IMPORT`.** That warning is
      // about the module not moving into a chunk of its own, which it cannot: `surface/pdf.ts`
      // statically reaches `Signature.fillPlaceholder`, `Signature.verify` and the other primitives
      // as public members. What this import still buys is the thing that matters — it
      // keeps the CMS/ASN.1 machinery out of a consumer who only calls `Pdf.fromTable`, because *this*
      // module is eager for them and the signature module is not. Making it static to silence the
      // warning would hand every `Pdf.fromTable` bundle the whole signing stack.
      const signature = await import("@pdf/core/digital-signature");
      // Injected into the rewrite, which adds the signature field that holds it.
      this._signaturePlaceholder = signature.buildSignatureDictPlaceholder({
        name: options.name,
        reason: options.reason,
        location: options.location,
        contactInfo: options.contactInfo
      }).dictString;
      try {
        // sign() predates toBuffer()'s form rules and keeps the behaviour it had.
        const bytes = await this._writeFull(
          { invalidateSignatures: true, encryption: options.encryption },
          { caller: "sign()", legacyFormRules: true, rewriteEncryption: "keep" }
        );
        return { pdfWithPlaceholder: bytes, signPdf: signature.signPdf };
      } finally {
        this._signaturePlaceholder = null;
      }
    });

    return signPdf(pdfWithPlaceholder, options.certificate, options.privateKey);
  }

  /**
   * @internal - Throw the first resource limit hit while reading this document
   * or a document pages were copied from. Reads during a save tolerate
   * malformed objects (see the `catch` blocks above), so a limit hit is
   * recorded on the document rather than thrown where it happened.
   */
  private _throwIfLimitExceeded(): void {
    this._doc.throwIfLimitExceeded();
    for (const appended of this._edits.appendedPages) {
      if (appended.kind === "copied") {
        appended.sourceDoc.throwIfLimitExceeded();
      }
    }
  }

  /**
   * @internal - The security handler a full rewrite encrypts with: the document's own, so the
   * output opens with the passwords the original does and grants the same permissions.
   *
   * The file key does not depend on which password opened the document, and every object is
   * renumbered by the rewrite, so the handler is the reader's own (`PdfDocument.encryptFn`,
   * which keys each object by the number it is written under) together with the original
   * `/Encrypt` dictionary and `/ID` — the inputs a reader derives the key from. Nothing is
   * re-derived, which is why this works for a document opened with the user password, whose
   * owner password is unknown.
   */
  private _rewriteSecurity(
    encryption: "keep" | "remove",
    caller: string
  ): PdfWriterSecurityHandler | null {
    if (!isEncrypted(this._doc) || encryption === "remove") {
      return null;
    }
    const encryptFn = this._doc.encryptFn;
    const encryptDict = this._doc.derefDict(this._doc.trailer.get("Encrypt") ?? null);
    if (!encryptFn || !encryptDict) {
      throw new PdfEncryptionError(
        `${caller}: the document declares /Encrypt, but its encryption dictionary could not be ` +
          "read, so the rewrite cannot be encrypted with it. Nothing was written. " +
          'Pass { encryption: "remove" } to write it without encryption.'
      );
    }
    const id = this._doc.trailer.get("ID");
    const ids = isPdfArray(id)
      ? id.filter((item): item is Uint8Array => item instanceof Uint8Array)
      : [];
    return {
      encryptDict: this._serializeEncryptDict(encryptDict, caller),
      fileId: ids.length > 0 ? [ids[0], ids[1] ?? ids[0]] : undefined,
      encrypt: encryptFn,
      encryptMetadata: this._doc.encryptMetadata
    };
  }

  /**
   * @internal - The `/Encrypt` dictionary with every reference resolved in place, so it does
   * not point at objects of the original file.
   *
   * Resolving in place copies whatever a reference reaches, so the file decides how much is
   * copied: a reference back into the dictionary never ends, and a value referenced twice
   * from each of a few nested levels doubles per level. A real dictionary is well under a
   * kilobyte, so either is a malformed file, and the encryption cannot be carried over.
   */
  private _serializeEncryptDict(encryptDict: PdfDictValue, caller: string): string {
    const unusable = (why: string): PdfEncryptionError =>
      new PdfEncryptionError(
        `${caller}: the document's encryption dictionary ${why}, so the rewrite cannot be ` +
          'encrypted with it. Nothing was written. Pass { encryption: "remove" } to write it ' +
          "without encryption."
      );
    const open = new Set<string>();
    let size = 0;
    const leaf = (out: string): string => {
      size += out.length;
      if (size > MAX_INLINE_ENCRYPT_DICT_CHARS) {
        throw unusable("is too large to copy");
      }
      return out;
    };
    const inline = (val: PdfObject, depth: number): string => {
      if (depth > 32) {
        return leaf("null");
      }
      const key = isPdfRef(val) ? `${val.objNum}:${val.gen}` : undefined;
      if (key !== undefined && open.has(key)) {
        throw unusable("refers back to itself");
      }
      const resolved = isPdfRef(val) ? this._doc.deref(val) : val;
      if (key !== undefined) {
        open.add(key);
      }
      let out: string;
      if (resolved instanceof Map) {
        const parts = ["<<"];
        for (const [name, item] of resolved) {
          parts.push(`${pdfNameBytes(name)} ${inline(item, depth + 1)}`);
        }
        parts.push(">>");
        out = parts.join(" ");
      } else if (isPdfArray(resolved)) {
        out = `[${resolved.map(item => inline(item, depth + 1)).join(" ")}]`;
      } else {
        out = leaf(this._serializeOriginalValue(resolved));
      }
      if (key !== undefined) {
        open.delete(key);
      }
      return out;
    };
    return inline(encryptDict, 0);
  }

  /** @internal - Throw when the document is signed and the caller has not opted in. */
  private _guardSignatures(options: EditorWriteOptions | undefined, reason: string): void {
    if (this.hasSignatures && !options?.invalidateSignatures) {
      throw new PdfSignatureInvalidationError(
        `The document is digitally signed and ${reason}, which would invalidate its signatures. ` +
          "Pass { invalidateSignatures: true } to proceed anyway."
      );
    }
  }

  /**
   * @internal - A field counts as a signature only when it is (or inherits
   * being) /FT /Sig and its /V is a signature dictionary with /ByteRange and
   * /Contents. /SigFlags alone is a hint a producer may leave set after the
   * signature is removed, so it is not evidence.
   */
  private _detectSignatures(): boolean {
    const root = this._doc.deref(this._doc.trailer.get("Root") ?? null);
    if (!(root instanceof Map)) {
      return false;
    }
    const acroForm = this._doc.deref(root.get("AcroForm") ?? null);
    if (!(acroForm instanceof Map)) {
      return false;
    }
    const visited = new Set<string>();
    // Each entry carries the /FT inherited from its ancestors (ISO 32000-2 §12.7.4.1).
    const stack: Array<{ item: PdfObject; inheritedFt: string | undefined }> = [];
    const fields = this._doc.deref(acroForm.get("Fields") ?? null);
    // Push element by element: spreading an untrusted, unbounded array into
    // an argument list overflows the call stack.
    if (isPdfArray(fields)) {
      for (const f of fields) {
        stack.push({ item: f, inheritedFt: undefined });
      }
    }
    while (stack.length > 0) {
      const { item, inheritedFt } = stack.pop()!;
      if (isPdfRef(item)) {
        const key = `${item.objNum}:${item.gen}`;
        if (visited.has(key)) {
          continue;
        }
        visited.add(key);
      }
      const field = this._doc.deref(item);
      if (!(field instanceof Map)) {
        continue;
      }
      const ft = dictGetName(field, "FT") ?? inheritedFt;
      if (ft === "Sig") {
        const value = this._doc.deref(field.get("V") ?? null);
        if (value instanceof Map && value.has("ByteRange") && value.has("Contents")) {
          return true;
        }
      }
      const kids = this._doc.deref(field.get("Kids") ?? null);
      if (isPdfArray(kids)) {
        for (const kid of kids) {
          stack.push({ item: kid, inheritedFt: ft });
        }
      }
    }
    return false;
  }

  /** @internal - Serialize the trailer entries an incremental update must repeat. */
  private _serializeTrailerEntries(): Map<string, string> {
    const entries = new Map<string, string>();
    // /ID is repeated verbatim: ISO 32000-2 §14.4 requires ID[0] to stay
    // the file's permanent identifier, and keeping ID[1] is permitted.
    for (const key of ["Root", "Info", "Encrypt", "ID"]) {
      const value = this._doc.trailer.get(key);
      if (value !== undefined && value !== null) {
        entries.set(key, this._serializeOriginalValue(value));
      }
    }
    // /Size decides where the update's own xref stream object is numbered;
    // without it that number could collide with an existing object.
    entries.set("Size", String(this._doc.objectNumberBound));
    return entries;
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
  private _serializeDictToWriter(
    doc: PdfDocument,
    dict: PdfDictValue,
    overrides?: ReadonlyMap<string, PdfObject>,
    /** Entries already serialized for the writer, written in place of the dict's own. */
    raw?: ReadonlyMap<string, string>
  ): string {
    const parts: string[] = ["<<"];

    for (const [key, val] of dict.entries()) {
      const serialized =
        raw?.get(key) ?? this._deepSerialize(doc, overrides?.has(key) ? overrides.get(key)! : val);
      if (serialized) {
        parts.push(`${pdfNameBytes(key)} ${serialized}`);
      }
    }
    for (const [key, val] of overrides ?? []) {
      if (!dict.has(key) && !raw?.has(key)) {
        parts.push(`${pdfNameBytes(key)} ${this._deepSerialize(doc, val)}`);
      }
    }
    for (const [key, val] of raw ?? []) {
      if (!dict.has(key)) {
        parts.push(`${pdfNameBytes(key)} ${val}`);
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
    if (this._cloneQueue) {
      return this._serializeValue(doc, val);
    }
    // Outermost call: clone every indirect object reached, one at a time. Recursing into
    // each reference instead made the stack depth equal to the length of the longest
    // reference chain in the file — a /Parent chain, a /Kids tree, a linked list of
    // annotations — which the file is free to make as long as it likes.
    const queue: PendingClone[] = [];
    this._cloneQueue = queue;
    try {
      const out = this._serializeValue(doc, val);
      for (let i = 0; i < queue.length; i++) {
        this._writeClone(queue[i]);
      }
      return out;
    } finally {
      this._cloneQueue = null;
    }
  }

  /** @internal - Write one object `_serializeValue` queued. */
  private _writeClone(clone: PendingClone): void {
    const { doc, objNum } = clone;
    if (clone.kind === "stream") {
      const { stream, objNum: srcNum, gen } = clone.source;
      // The bytes are copied as they are encoded, with their filters. Decoding and writing
      // the result without the filters lost what the decoder passes through — every JPEG
      // came out of a rewrite without `/DCTDecode` — and what it cannot read: a deflate
      // stream that does not inflate was written back as its compressed bytes, unfiltered.
      // A `/Crypt` filter names a crypt filter of the original file's handler, which a
      // rewrite may not keep, so such a stream is still decoded.
      const filter = stream.dict.get("Filter");
      const crypt = filter === "Crypt" || (isPdfArray(filter) && filter.includes("Crypt"));
      const streamData = crypt
        ? doc.getStreamData(stream, srcNum, gen)
        : doc.getEncodedStreamData(stream, srcNum, gen);
      // Serialize the stream's dict (excluding /Length which will be set automatically)
      const streamDict = new PdfDict();
      for (const [k, v] of stream.dict.entries()) {
        if (k === "Length" || k === "DL" || (crypt && (k === "Filter" || k === "DecodeParms"))) {
          // Skip — the writer sets the length; a decoded stream loses its filters
          continue;
        }
        const sv = this._serializeValue(doc, v);
        if (sv) {
          streamDict.set(k, sv);
        }
      }
      this._writerForSave!.addStreamObject(objNum, streamDict, streamData);
      return;
    }
    if (clone.kind === "value") {
      this._writerForSave!.addObject(objNum, this._serializeValue(doc, clone.value));
      return;
    }
    this._writerForSave!.addObject(
      objNum,
      this._serializeDictToWriter(doc, clone.dict, this._cloneOverrides(clone.cacheKey))
    );
  }

  /** @internal - `_deepSerialize` for one value; queues the indirect objects it meets. */
  private _serializeValue(doc: PdfDocument, val: PdfObject): string {
    if (val === null || val === undefined) {
      return "null";
    }
    if (typeof val === "string") {
      // In parsed PDF objects, string type = PDF name (without leading /)
      return pdfNameBytes(val);
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
      const docId = this._docId(doc);
      const cacheKey = `${docId}:${val.objNum}:${val.gen}`;
      const cached = this._clonedRefs.get(cacheKey);
      if (cached !== undefined) {
        return pdfRef(cached);
      }
      if (this._isDroppedRef(cacheKey)) {
        return "null";
      }

      // Try as stream first
      const streamResult = doc.derefStreamWithObjNum(val);
      if (streamResult) {
        const objNum = this._writerForSave!.allocObject();
        this._clonedRefs.set(cacheKey, objNum);
        this._cloneQueue!.push({ kind: "stream", doc, cacheKey, objNum, source: streamResult });
        return pdfRef(objNum);
      }

      // Try as regular dict/value
      const resolved = doc.deref(val);
      if (resolved instanceof Map) {
        const objNum = this._writerForSave!.allocObject();
        this._clonedRefs.set(cacheKey, objNum);
        this._cloneQueue!.push({
          kind: "dict",
          doc,
          cacheKey,
          objNum,
          dict: resolved as PdfDictValue
        });
        return pdfRef(objNum);
      }

      // An array or a reference behind a reference is cloned as an object of its own, like a
      // dictionary: inlining it recursed once per link, so a cycle made only of arrays and
      // references (`4 0 obj [4 0 R]`) never ended, and a long chain overflowed the stack.
      if (isPdfArray(resolved) || isPdfRef(resolved)) {
        const objNum = this._writerForSave!.allocObject();
        this._clonedRefs.set(cacheKey, objNum);
        this._cloneQueue!.push({ kind: "value", doc, cacheKey, objNum, value: resolved });
        return pdfRef(objNum);
      }

      // Any other value behind a ref is a leaf — inline it
      return this._serializeValue(doc, resolved);
    }

    if (isPdfArray(val)) {
      const items = val.map(item => this._serializeValue(doc, item));
      return `[${items.join(" ")}]`;
    }

    if (val instanceof Map) {
      return this._serializeDictToWriter(doc, val as PdfDictValue);
    }

    return "";
  }

  /** @internal */
  private _writeOverlayImages(writer: PdfWriter, page: PdfPageSnapshot): Map<string, number> {
    const map = new Map<string, number>();
    for (let i = 0; i < page._images.length; i++) {
      const img = page._images[i];
      const imgName = `${this._overlayResourcePrefix}Im${i + 1}`;
      const objNum = writeImageXObject(writer, img.data, img.format);
      map.set(imgName, objNum);
    }
    return map;
  }

  /** @internal - Build overlay resources as a structured PdfResourceDict. */
  private _buildOverlayResourceDict(
    writer: PdfWriter,
    overlay: PdfPageSnapshot | undefined,
    fontDictStr: string,
    imageXObjectMap: Map<string, number>
  ): PdfResourceDict {
    const dict: PdfResourceDict = new Map();

    if (fontDictStr) {
      // Parse the font dict string into structured entries
      // fontDictStr is already a `<< /F1 3 0 R ... >>` string
      const fontInner = fontDictStr.trim();
      if (fontInner.startsWith("<<") && fontInner.endsWith(">>")) {
        const parsed = parseResourceDict(`<< /Font ${fontInner} >>`);
        const fontMap = parsed.get("Font");
        if (fontMap) {
          dict.set("Font", fontMap);
        }
      }
    }

    if (imageXObjectMap.size > 0) {
      const xobjMap = new Map<string, string>();
      for (const [name, objNum] of imageXObjectMap) {
        xobjMap.set(name, pdfRef(objNum));
      }
      dict.set("XObject", xobjMap);
    }

    // Serialise `PdfPageBuilder._alphaValues` into an `/ExtGState` sub-dict.
    // Previously this path only wrote `/Font` + `/XObject`, so any alpha
    // emitted by `_applyAlpha` (e.g. a chart with filled radar / translucent
    // area rendered via `drawChartPdf`) left the `gs` operator pointing at
    // an undefined resource. The scheme matches `pdf-exporter.ts:377-389`
    // and `document-builder.ts` so the three pipelines stay name-compatible.
    if (overlay && overlay._alphaValues.size > 0) {
      const gsMap = new Map<string, string>();
      for (const alpha of overlay._alphaValues) {
        const gsObjNum = writer.allocObject();
        const gsDict = new PdfDict()
          .set("Type", "/ExtGState")
          .set("ca", pdfNumber(alpha))
          .set("CA", pdfNumber(alpha));
        writer.addObject(gsObjNum, gsDict);
        gsMap.set(alphaGsName(alpha, this._overlayResourcePrefix), pdfRef(gsObjNum));
      }
      dict.set("ExtGState", gsMap);
    }

    return dict;
  }

  /** @internal - Merge original and overlay resource strings via parsed object graph. */
  private _mergeResourceStrings(original: string, overlay: string): string {
    // If no overlay, return original
    if (!overlay || overlay === "<< >>") {
      return original;
    }
    // If no original, return overlay
    if (!original || original === "<< >>") {
      return overlay;
    }

    const origDict = parseResourceDict(original);
    const overlayDict = parseResourceDict(overlay);
    const merged = mergeResourceDicts(origDict, overlayDict);
    return serializeResourceDict(merged);
  }

  /** @internal - Write form field value updates as annotation objects. */
  private _writeFormFieldUpdates(
    writer: PdfWriter,
    pageDict: PdfDictValue,
    _pageIndex: number
  ): number[] {
    if (this._edits.formFieldUpdates.size === 0) {
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
      let objNum: number;
      let kidsOverride: PdfObject[] | undefined;
      if (isPdfRef(annotRef)) {
        const cacheKey = `main:${annotRef.objNum}:${annotRef.gen}`;
        kidsOverride = this._pruning?.kidsOverride(`${annotRef.objNum}:${annotRef.gen}`);
        const assigned = this._clonedRefs.get(cacheKey);
        if (assigned !== undefined && this._writtenAnnotObjNums.has(assigned)) {
          // The same annotation listed twice — it is one object, written once.
          annotRefs.push(assigned);
          continue;
        }
        objNum = assigned ?? writer.allocObject();
        this._clonedRefs.set(cacheKey, objNum);
        this._writtenAnnotObjNums.add(objNum);
      } else {
        objNum = writer.allocObject();
      }

      if (subtype === "Widget" && this._edits.formFieldUpdates.size > 0) {
        // Check if this widget has a field name that matches an update
        const fieldName = this._resolveFieldName(annotDict);
        const newValue = fieldName ? this._edits.formFieldUpdates.get(fieldName) : undefined;

        if (newValue !== undefined) {
          // Write a modified widget annotation with the new value
          const newDict = this._buildModifiedWidgetDict(annotDict, newValue);
          writer.addObject(objNum, newDict);
          annotRefs.push(objNum);
          continue;
        }
      }

      // Copy annotation as-is (serialize what we can)
      const serialized = this._serializeAnnotDict(annotDict, kidsOverride);
      writer.addObject(objNum, serialized);
      annotRefs.push(objNum);
    }

    return annotRefs;
  }

  /**
   * @internal - The key prefix of `doc`'s objects in `_clonedRefs`: `main` for this document,
   * an id per source document otherwise (two sources of the same byte length used to share
   * one, and so each other's cloned objects), with the clone scope of the page copy running.
   */
  private _docId(doc: PdfDocument): string {
    if (doc === this._doc) {
      return "main";
    }
    let id = this._sourceDocIds.get(doc);
    if (id === undefined) {
      id = `src${this._sourceDocIds.size + 1}`;
      this._sourceDocIds.set(doc, id);
    }
    return id + this._cloneScope;
  }

  /** @internal - Whether a full rewrite leaves an original object out (see `_planPruning`). */
  private _isDroppedRef(cacheKey: string): boolean {
    if (this._droppedSourceRefs.has(cacheKey)) {
      return true;
    }
    if (!cacheKey.startsWith("main:")) {
      return false;
    }
    const key = cacheKey.slice(5);
    return (
      (this._pruning?.isDropped(key) ?? false) ||
      (this._threadPlan?.isDropped(key) ?? false) ||
      (this._catalogDropped !== null && this._catalogDropped(key))
    );
  }

  /**
   * @internal - The entries a full rewrite writes for an original object instead of its own:
   * the /Kids of a field node that loses some, the links of an article bead or thread whose
   * beads on pages left out are removed.
   */
  private _cloneOverrides(cacheKey: string): ReadonlyMap<string, PdfObject> | undefined {
    if (!cacheKey.startsWith("main:")) {
      return undefined;
    }
    const key = cacheKey.slice(5);
    const kids = this._pruning?.kidsOverride(key);
    const base = kids ? new Map([["Kids", kids]]) : this._threadPlan?.overrides(key);
    const button = this._buttonParents?.get(key);
    if (button === undefined) {
      return base;
    }
    // A button's value is a name: a parsed name is a string.
    return new Map<string, PdfObject>([...(base ?? []), ["V", button.value]]);
  }

  /** @internal - What a full rewrite plans for the article threads (see `_threadPlan`). */
  private _planThreads(
    pagesInfo: ReadonlyArray<{ dict: PdfDictValue; objNum: number; gen: number }>,
    outputPages: ReadonlyArray<number | null>
  ): ThreadPlan | null {
    let catalog: PdfDictValue;
    try {
      catalog = this._doc.getCatalog();
    } catch {
      return null;
    }
    this._catalogIndex ??= buildCatalogIndex(this._doc, catalog, pagesInfo);
    return planThreads(
      this._doc,
      catalog,
      this._catalogIndex,
      outputPages,
      this._edits.pages.kind === "only"
    );
  }

  /**
   * @internal - An original page's /B, the article beads on it in drawing order, which ISO
   * 32000-1 §12.4.3 requires of every page with beads. When the rewrite changes the threads,
   * only the beads of a thread in /Threads are listed, and only the ones it writes.
   */
  private _pageBeads(pageDict: PdfDictValue): string | undefined {
    const beads = this._doc.deref(pageDict.get("B") ?? null);
    if (!isPdfArray(beads)) {
      return undefined;
    }
    const plan = this._threadPlan;
    const refs = beads.filter(
      (bead): bead is PdfRef =>
        isPdfRef(bead) &&
        (!plan ||
          (plan.isBead(`${bead.objNum}:${bead.gen}`) &&
            !plan.isDropped(`${bead.objNum}:${bead.gen}`)))
    );
    const written = refs
      .map(bead => this._deepSerialize(this._doc, bead))
      .filter(v => v !== "null");
    return written.length === 0 ? undefined : `[${written.join(" ")}]`;
  }

  /**
   * @internal - Reserve the object numbers a full rewrite gives each kept page and each
   * annotation on it, and record them as those objects' clones.
   *
   * The pages and their annotations are written by hand — a widget with a new value is
   * rebuilt, not copied — but other objects still point at the originals: the AcroForm's
   * /Fields, a parent field's /Kids, a widget's /P. Unless the numbers are known before
   * those are cloned, each reference clones a second, unedited copy; the form then lists
   * widgets that are not the ones on the page and reports the values from before the
   * edit, while the page shows the new ones.
   */
  private _preassignPageAndAnnotClones(
    writer: PdfWriter,
    pagesInfo: ReadonlyArray<{ dict: PdfDictValue; objNum: number; gen: number }>,
    keptPages: readonly number[]
  ): Map<number, number> {
    const pageObjNums = new Map<number, number>();
    for (const i of keptPages) {
      const { dict, objNum, gen } = pagesInfo[i];
      const pageObjNum = writer.allocObject();
      pageObjNums.set(i, pageObjNum);
      if (objNum > 0) {
        this._clonedRefs.set(`main:${objNum}:${gen}`, pageObjNum);
      }
      const annots = this._doc.deref(dict.get("Annots") ?? null);
      if (!isPdfArray(annots)) {
        continue;
      }
      for (const annotRef of annots) {
        if (!isPdfRef(annotRef)) {
          continue;
        }
        const cacheKey = `main:${annotRef.objNum}:${annotRef.gen}`;
        if (!this._clonedRefs.has(cacheKey) && this._doc.derefDict(annotRef)) {
          this._clonedRefs.set(cacheKey, writer.allocObject());
        }
      }
    }
    return pageObjNums;
  }

  /**
   * @internal - What a full rewrite leaves out with the original pages it does not write, or
   * `null` when it writes them all. The rules, and why, are in `page-index.ts`; they are read
   * there once per document and shared by every write.
   *
   * The incremental path needs none of this: removing a page always falls back to a full
   * rewrite, so an update never leaves a field pointing at a page that is gone.
   */
  private _planPruning(
    pagesInfo: ReadonlyArray<{ dict: PdfDictValue; objNum: number; gen: number }>
  ): Pruning | null {
    const { pages } = this._edits;
    if (!removesAnyPage(pages, pagesInfo.length)) {
      return null;
    }
    this._pageIndex ??= buildPageIndex(this._doc, pagesInfo);
    return pages.kind === "only"
      ? pruneToPage(this._pageIndex, pages.index ?? -1)
      : pruneRemovedPages(this._pageIndex, pages.removed);
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
  private _buildModifiedWidgetDict(
    originalDict: PdfDictValue,
    newValue: string,
    addAppearanceStream?: (dict: PdfDict, data: Uint8Array) => number
  ): string {
    // Determine the field type (/FT) — may be directly on the widget or inherited from parent
    const fieldType = this._resolveFieldType(originalDict);

    // For text fields, generate an appearance stream instead of stripping /AP
    if (fieldType === "Tx" && (addAppearanceStream || this._writerForSave)) {
      return this._buildTextFieldWidgetDict(
        originalDict,
        newValue,
        addAppearanceStream ??
          ((dict, data) => {
            const writer = this._writerForSave!;
            const objNum = writer.allocObject();
            writer.addStreamObject(objNum, dict, data, { compress: false });
            return objNum;
          })
      );
    }

    if (fieldType === "Btn") {
      return this._buildButtonWidgetDict(originalDict, newValue);
    }

    // For non-text fields, fall back to stripping /AP (force viewer to regenerate)
    const parts: string[] = ["<<"];

    for (const [key, val] of originalDict.entries()) {
      if (key === "V" || key === "AP") {
        continue;
      }

      const serialized = this._serializePdfValue(val);
      if (serialized) {
        parts.push(`${pdfNameBytes(key)} ${serialized}`);
      }
    }

    parts.push(`/V ${pdfString(newValue)}`);
    parts.push(">>");
    return parts.join(" ");
  }

  /**
   * @internal - A checkbox or radio widget showing `newValue`.
   *
   * A button's value is a *name* and its look is the appearance state `/AS` picks out of
   * `/AP /N`. The generic path wrote the value as a string, left `/AS` as it was and dropped
   * the appearances — so a checkbox set to `Yes` still drew unchecked, and its value was a
   * string no viewer reads as a state. Here the appearances are kept, `/AS` is the new state
   * when this widget has it (each radio button has its own) and `/Off` otherwise, and `/V` is
   * written as a name on the widget when the widget is the field itself; a radio button's
   * value lives on its parent field (see `_buttonParentValues`).
   */
  private _buildButtonWidgetDict(originalDict: PdfDictValue, newValue: string): string {
    const normal = this._doc.derefDict(this._doc.derefDict(originalDict.get("AP"))?.get("N"));
    const ownStates = normal ? [...normal.keys()].filter(key => key !== "Off") : [];
    const state =
      newValue !== "Off" && (ownStates.length === 0 || ownStates.includes(newValue))
        ? newValue
        : "Off";
    const isField = originalDict.has("T");
    const parts: string[] = ["<<"];
    for (const [key, val] of originalDict.entries()) {
      if (key === "AS" || (key === "V" && isField)) {
        continue;
      }
      const serialized = this._serializePdfValue(val);
      if (serialized) {
        parts.push(`${pdfNameBytes(key)} ${serialized}`);
      }
    }
    parts.push(`/AS ${pdfName(state)}`);
    if (isField) {
      parts.push(`/V ${pdfName(newValue)}`);
    }
    parts.push(">>");
    return parts.join(" ");
  }

  /**
   * @internal - The entries of the information dictionary a write puts in the file, serialized:
   * the document's own (each with `serialize`), with the fields of `metadata` written over
   * them. Nothing else: no `/Creator` or `/Producer` is invented — a default here, absent from
   * the XMP packet, made the two records disagree.
   */
  private _infoEntries(
    metadata: PdfMetadataUpdate | undefined,
    serialize: (value: PdfObject) => string | null
  ): Map<string, string> {
    const out = new Map<string, string>();
    const info = this._doc.derefDict(this._doc.trailer.get("Info") ?? null);
    for (const [key, value] of info ?? []) {
      const serialized = serialize(value);
      if (serialized) {
        out.set(key, serialized);
      }
    }
    const strings = {
      title: "Title",
      author: "Author",
      subject: "Subject",
      keywords: "Keywords",
      creator: "Creator",
      producer: "Producer"
    };
    for (const [field, key] of Object.entries(strings) as Array<[keyof typeof strings, string]>) {
      const value = metadata?.[field];
      if (value === null) {
        out.delete(key);
      } else if (value !== undefined) {
        out.set(key, pdfString(value));
      }
    }
    for (const [field, key] of [
      ["creationDate", "CreationDate"],
      ["modDate", "ModDate"]
    ] as const) {
      const value = metadata?.[field];
      if (value === null) {
        out.delete(key);
      } else if (value !== undefined) {
        out.set(key, pdfDate(value));
      }
    }
    return out;
  }

  /**
   * @internal - The document's XMP packet with the properties of `metadata` set (see
   * `updateXmpPacket`), as a stream to write — uncompressed, in the packet's own encoding — and
   * the object it replaces. The stream dictionary's other entries are kept, each with
   * `serialize` (a full rewrite renumbers the references in it).
   *
   * `undefined` when the document has no packet, or one that cannot be decoded or is not XMP:
   * the caller then keeps the original metadata stream untouched and changes only /Info,
   * rather than writing a guess — or stripping the filter of bytes it could not inflate.
   */
  private _updatedXmp(
    metadata: PdfMetadataUpdate,
    serialize: (value: PdfObject) => string | null
  ): { ref: PdfRef | undefined; dict: PdfDict; data: Uint8Array } | undefined {
    let metadataRef: PdfObject | undefined;
    try {
      metadataRef = this._doc.getCatalog().get("Metadata");
    } catch {
      return undefined;
    }
    if (metadataRef === undefined) {
      return undefined;
    }
    const found = this._doc.derefStreamWithObjNum(metadataRef);
    if (!found) {
      return undefined;
    }
    const decoded = decodeXmpPacket(this._doc.getStreamData(found.stream, found.objNum, found.gen));
    if (decoded === undefined) {
      return undefined;
    }
    // XMP requires `xmp:MetadataDate` to change with the metadata: the edit's own modification
    // date when it gives one, else the time of the write.
    const updated = updateXmpPacket(decoded.text, metadata, metadata.modDate ?? new Date());
    if (updated === undefined) {
      return undefined;
    }
    const data = encodeXmpPacket(updated, decoded.encoding);
    const dict = new PdfDict();
    for (const [key, value] of found.stream.dict) {
      // The data is written decoded and uncompressed: its filters and length go.
      if (key === "Length" || key === "Filter" || key === "DecodeParms" || key === "DL") {
        continue;
      }
      const serialized = serialize(value);
      if (serialized) {
        dict.set(key, serialized);
      }
    }
    return { ref: isPdfRef(metadataRef) ? metadataRef : undefined, dict, data };
  }

  /**
   * @internal - The parent fields whose value a write changes because one of their button
   * widgets was set: a radio group, or a checkbox with several widgets, keeps `/V` on the
   * field node above the widgets. Keyed by `"objNum:gen"`.
   */
  private _buttonParentValues(): Map<string, { ref: PdfRef; dict: PdfDictValue; value: string }> {
    const out = new Map<string, { ref: PdfRef; dict: PdfDictValue; value: string }>();
    const updates = this._edits.formFieldUpdates;
    if (updates.size === 0) {
      return out;
    }
    for (const { dict: pageDict } of this._doc.getPagesWithObjInfo()) {
      const annots = this._doc.deref(pageDict.get("Annots"));
      if (!isPdfArray(annots)) {
        continue;
      }
      for (const annotRef of annots) {
        const widget = this._doc.derefDict(annotRef);
        if (!widget || dictGetName(widget, "Subtype") !== "Widget" || widget.has("T")) {
          continue;
        }
        const parentRef = widget.get("Parent");
        if (!isPdfRef(parentRef) || this._resolveFieldType(widget) !== "Btn") {
          continue;
        }
        const value = updates.get(this._resolveFieldName(widget));
        const parent = this._doc.derefDict(parentRef);
        if (value !== undefined && parent) {
          out.set(`${parentRef.objNum}:${parentRef.gen}`, { ref: parentRef, dict: parent, value });
        }
      }
    }
    return out;
  }

  /**
   * @internal - Build a modified widget dict for a text field with an inline
   * appearance stream. The stream renders the field value so it is visible
   * in all viewers, even those that ignore /NeedAppearances.
   */
  private _buildTextFieldWidgetDict(
    originalDict: PdfDictValue,
    newValue: string,
    addAppearanceStream: (dict: PdfDict, data: Uint8Array) => number
  ): string {
    // Extract the widget Rect for sizing the appearance
    const rect = this._resolveWidgetRect(originalDict);

    // Determine alignment from /Q entry (0=left, 1=center, 2=right)
    const qVal = originalDict.get("Q");
    let alignment: "left" | "center" | "right" = "left";
    if (qVal === 1) {
      alignment = "center";
    } else if (qVal === 2) {
      alignment = "right";
    }

    // Generate the appearance stream
    const { stream, resources } = generateTextFieldAppearance({
      value: newValue,
      rect,
      alignment
    });

    // Write the appearance stream as a Form XObject indirect object
    const apDict = new PdfDict()
      .set("Type", "/XObject")
      .set("Subtype", "/Form")
      .set("BBox", buildAppearanceBBox(rect))
      .set("Resources", resources);
    const apObjNum = addAppearanceStream(apDict, stream);

    // Build the widget dict
    const parts: string[] = ["<<"];

    for (const [key, val] of originalDict.entries()) {
      if (key === "V" || key === "AP") {
        continue;
      }

      const serialized = this._serializePdfValue(val);
      if (serialized) {
        parts.push(`${pdfNameBytes(key)} ${serialized}`);
      }
    }

    // Set the new value and appearance
    parts.push(`/V ${pdfString(newValue)}`);
    parts.push(`/AP << /N ${pdfRef(apObjNum)} >>`);

    parts.push(">>");
    return parts.join(" ");
  }

  /**
   * @internal - Resolve the field type (/FT) which may be inherited from a parent dict.
   */
  private _resolveFieldType(dict: PdfDictValue): string | undefined {
    let current: PdfDictValue | null | undefined = dict;
    const visited = new Set<string>();

    while (current) {
      const ft = dictGetName(current, "FT");
      if (ft) {
        return ft;
      }
      const parentVal = current.get("Parent");
      if (!parentVal) {
        break;
      }
      const key = String(parentVal);
      if (visited.has(key)) {
        break;
      }
      visited.add(key);
      current = this._doc.derefDict(parentVal);
    }
    return undefined;
  }

  /**
   * @internal - Extract the /Rect array from a widget annotation dict as [x1, y1, x2, y2].
   */
  private _resolveWidgetRect(dict: PdfDictValue): number[] {
    let rectVal = dict.get("Rect");
    if (rectVal) {
      rectVal = this._doc.deref(rectVal);
    }
    if (rectVal && isPdfArray(rectVal)) {
      return rectVal.map(v => (typeof v === "number" ? v : 0));
    }
    // Fallback: 0-size rect
    return [0, 0, 100, 20];
  }

  /** @internal */
  private _serializeAnnotDict(dict: PdfDictValue, kids?: readonly PdfObject[]): string {
    const parts: string[] = ["<<"];

    for (const [key, val] of dict.entries()) {
      const serialized = this._serializePdfValue(key === "Kids" && kids ? [...kids] : val);
      if (serialized) {
        parts.push(`${pdfNameBytes(key)} ${serialized}`);
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
      // In parsed PDF objects, string type = PDF name (without leading /)
      return pdfNameBytes(val);
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
      // During incremental save, original refs remain valid — preserve them as-is.
      // During full save, deep-clone into the new writer.
      if (this._writerForSave && !this._isIncrementalSave) {
        return this._deepSerialize(this._doc, val);
      }
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
        parts.push(`${pdfNameBytes(k)} ${serialized}`);
      }
      parts.push(">>");
      return parts.join(" ");
    }
    return "";
  }
}

// =============================================================================
// Internal Types
// =============================================================================

type PendingClone = {
  doc: PdfDocument;
  cacheKey: string;
  objNum: number;
} & (
  | { kind: "dict"; dict: PdfDictValue }
  | { kind: "value"; value: PdfObject }
  | {
      kind: "stream";
      source: NonNullable<ReturnType<PdfDocument["derefStreamWithObjNum"]>>;
    }
);

/**
 * One write request, threaded through the write path rather than kept on the editor, so two writes
 * in flight on the same editor cannot see each other's rules.
 */
interface WriteCall {
  /** The public call, for error messages. */
  caller: string;
  /**
   * Write edited form values the way `save()`/`saveIncremental()` always did: no dynamic-XFA
   * refusal, `/XFA` kept, and no AcroForm revision appended to an update. Set only for the
   * deprecated forwards and `sign()`, whose behaviour predates `toBuffer()`'s form rules and must
   * not change under their callers.
   */
  legacyFormRules: boolean;
  /**
   * What a full rewrite of an encrypted document does when the caller passed no `encryption`.
   * `"remove"` only for the deprecated `save()`/`saveIncremental()`: `save()` was documented to
   * write a decrypted copy, and `saveIncremental()` to fall back to `save()`. Everything else
   * keeps the encryption.
   */
  rewriteEncryption: "keep" | "remove";
}

/**
 * Most characters the copied `/Encrypt` dictionary may take (see `_serializeEncryptDict`). The
 * entries a security handler defines come to a few hundred; this leaves room for any number of
 * private ones while ending an exponential expansion long before it costs anything.
 */
const MAX_INLINE_ENCRYPT_DICT_CHARS = 1 << 20;

/** The edits one write serializes, frozen when it was called (see `PdfEditor._serialized`). */
interface EditorEdits {
  /** The overlay of each original page, by page index. */
  readonly overlays: readonly PdfPageSnapshot[];
  /** Pages added with `addPage` / `copyPagesFrom`, in the order they were added. */
  readonly appendedPages: readonly AppendedPageSnapshot[];
  readonly formFieldUpdates: ReadonlyMap<string, string>;
  /** Document properties set with `setMetadata`, if any. */
  readonly metadata: PdfMetadataUpdate | undefined;
  /** Which original pages are written. */
  readonly pages: PageSelection;
  readonly rotationOverrides: ReadonlyMap<number, number>;
}

/**
 * The original pages a write keeps: all but the removed ones, or — for a part of
 * `splitPages` — the one page named.
 * Naming the page rather than listing every other one as removed is what keeps the cost of a
 * part to what is on its page. A part that is an added page keeps no original page: `null`.
 */
type PageSelection =
  | { readonly kind: "removed"; readonly removed: ReadonlySet<number> }
  | { readonly kind: "only"; readonly index: number | null };

const editorReadOptionsSchema = s.object(
  { password: s.string().optional(), ...fontOptionsShape },
  { expected: "an options object" }
);
const _editorReadOptionsMatch: SchemaMatches<
  typeof editorReadOptionsSchema,
  { -readonly [K in keyof EditorReadOptions]: EditorReadOptions[K] }
> = true;
void _editorReadOptionsMatch;

const editorWriteOptionsSchema = s.object(
  {
    invalidateSignatures: s.boolean().optional(),
    encryption: s.oneOf(["keep", "remove"] as const).optional()
  },
  { expected: "an options object" }
);
const _editorWriteOptionsMatch: SchemaMatches<typeof editorWriteOptionsSchema, EditorWriteOptions> =
  true;
void _editorWriteOptionsMatch;

/** @internal - The options every editor write shares. */
function checkWriteOptions(where: string, options: EditorWriteOptions | undefined): void {
  check(editorWriteOptionsSchema.optional(), options, at(where, "options"));
}

/**
 * @internal - Refuse `pages` in an editor options object. The editor's page lists are 0-based
 * `pageIndices`; `pages` is `Pdf.extract`'s 1-based list, and was ignored here — copying or
 * splitting every page.
 */
function checkNotPagesKey(where: string, options: object): void {
  if ("pages" in options) {
    throw new PdfArgumentError(
      "options.pages",
      `${where}: the option is pageIndices (0-based page indices), not pages`
    );
  }
}

/** A page index as an error message shows it: a number as itself, anything else with its type. */
function describeIndex(value: unknown): string {
  if (typeof value === "number" || value === null || value === undefined) {
    return String(value);
  }
  if (typeof value === "string") {
    return `string ${JSON.stringify(value)}`;
  }
  if (Array.isArray(value)) {
    // `removePage([0, 2])` is the likeliest mistake of this kind: say it was an array.
    return `an array (${value.length} items)`;
  }
  const shown =
    typeof value === "object" || typeof value === "function"
      ? Object.prototype.toString.call(value)
      : String(value);
  return `${typeof value} ${shown}`;
}

/** The indices of the original pages `pages` keeps, ascending. */
function keptPageIndices(pages: PageSelection, pageCount: number): number[] {
  if (pages.kind === "only") {
    return pages.index === null ? [] : [pages.index];
  }
  const kept: number[] = [];
  for (let i = 0; i < pageCount; i++) {
    if (!pages.removed.has(i)) {
      kept.push(i);
    }
  }
  return kept;
}

/** Whether `pages` leaves out any of the document's `pageCount` original pages. */
function removesAnyPage(pages: PageSelection, pageCount: number): boolean {
  if (pages.kind === "only") {
    return pageCount > (pages.index === null ? 0 : 1);
  }
  return pages.removed.size > 0;
}

interface CopiedPage {
  kind: "copied";
  width: number;
  height: number;
  contentStreams: Uint8Array[];
  /** Source doc + page dict — for deferred resource deep-clone during save() */
  sourceDoc: PdfDocument;
  sourcePageDict: PdfDictValue;
  /** The source page's own reference, which its annotations' `/P` names. */
  sourcePageRef: { objNum: number; gen: number } | undefined;
  /** The fully qualified names of the form fields with a widget on the page. */
  fieldNames: readonly string[];
  /** `EditorCopyPagesOptions.formFields`: whether the page's widgets and fields are copied. */
  formFields: "copy" | "drop";
}

/**
 * A page added with `addPage`, or one copied in with `copyPagesFrom`, with the handle
 * `getPage` hands out for it (made on first request) and, for a copied page, the overlay that
 * handle draws onto.
 */
type AppendedPage =
  | { kind: "new"; page: PdfPageBuilder; handle?: PdfEditorPage }
  | (CopiedPage & { overlay?: PdfPageBuilder; handle?: PdfEditorPage });

/**
 * An {@link AppendedPage} as a write sees it: a new page's content, and a copied page's
 * overlay, frozen when it was called; `rotation` set with `rotatePage`.
 */
type AppendedPageSnapshot = (
  | { kind: "new"; page: PdfPageSnapshot }
  | (CopiedPage & { overlay?: PdfPageSnapshot })
) & { rotation?: number };
