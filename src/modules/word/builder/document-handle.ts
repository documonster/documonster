/**
 * Document Handle & Namespace for building and editing DOCX documents.
 *
 * Provides the Document namespace with free functions for constructing and
 * editing documents — `DocxDocument` trees.
 *
 * Ownership rule for the Word API: `Document.*` edits the document you pass,
 * in place. Every other namespace function that returns a `DocxDocument`
 * leaves its input unchanged; the result may share unchanged nodes with the
 * input, and `Document.*` copies a shared array or node before it changes it
 * (`@word/core/ownership`), so an edit made through either document never
 * reaches the other.
 */

import { isRecordArgument } from "@utils/object";
import {
  UNRECOGNIZED_KEY,
  discardUnrecognizedValues as discardUnrecognizedValuesUnder,
  findUnrecognizedValues
} from "@utils/unrecognized-values";
import { finalizeDocument } from "@word/builder/document-finalize";
import { checkedHeading, checkedTextParagraph, paragraph } from "@word/builder/paragraph-builders";
import { pageBreak, floatingImage, mathBlock } from "@word/builder/run-builders";
import { checkedSimpleTable } from "@word/builder/table-builders";
import type { SimpleTableOptions } from "@word/builder/table-builders";
import {
  checkAppProperties,
  checkArray,
  checkBlock,
  checkBoolean,
  checkBytes,
  checkCoreProperties,
  checkDateTime,
  checkDocumentSettings,
  checkHandle,
  checkHexColor,
  checkInteger,
  checkListLevel,
  checkNode,
  checkNodeArray,
  checkOneOf,
  checkParagraphProperties,
  checkRecord,
  checkRunProperties,
  checkSectionProperties,
  checkString,
  checkStyleDef,
  checkTableOfContentsOptions,
  checkWatermark,
  invalidArgument,
  under
} from "@word/core/argument-checks";
import { defaultStyleSet } from "@word/core/builtin-styles";
import {
  assertAnchorNumbers,
  assertDrawingNumber,
  assertDrawingNumbers
} from "@word/core/drawing-geometry";
import {
  HEADER_FOOTER_TYPES,
  impliedReferences,
  referencedPartKey
} from "@word/core/header-references";
import { assertImageMediaType } from "@word/core/image-media";
import type { Mutable } from "@word/core/internal-utils";
import { claim, isSharing, markDerived, ownContainer, ownTree } from "@word/core/ownership";
import { removeBlock } from "@word/core/range-markers";
import { parseTocInstruction } from "@word/core/toc-instruction";
import type {
  DocxDiscardUnrecognizedValuesOptions,
  DocxUnrecognizedValue,
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphProperties,
  ParagraphChild,
  RunProperties,
  Table,
  SectionProperties,
  StyleDef,
  DocDefaults,
  HeaderDef,
  FooterDef,
  HeaderFooterContent,
  HeaderFooterType,
  ImageMediaType,
  CoreProperties,
  AppProperties,
  DocumentSettings,
  FontDef,
  Emu,
  Twips,
  DocumentBackground,
  CustomPropertyValue,
  TableOfContents,
  MathContent,
  FloatingImage,
  Watermark,
  AbstractNumbering,
  NumberingInstance,
  FootnoteDef,
  EndnoteDef,
  CommentDef,
  ImageDef
} from "@word/types";

// In-place edits that walk the whole document. They lived in `Query` first,
// which keeps deprecated forwards to them.
export { replaceText } from "@word/query/replace";
export { mailMerge } from "@word/query/mail-merge";
export {
  acceptAllRevisions,
  rejectAllRevisions,
  acceptRevision,
  rejectRevision
} from "@word/query/revisions";
export { setCompatibilityMode } from "@word/query/compat";

/**
 * The values `Document.read` kept because they are outside their field's vocabulary, as they are kept now — a field set
 * since takes its value off the list, `discardUnrecognizedValues` takes them off for good — each with the path
 * `Validation.validate` reports (`body[0].properties.alignment`) and the text the file had. Empty for a document no
 * reader made.
 *
 * ```ts
 * for (const { path, value } of Document.listUnrecognizedValues(doc)) console.log(path, value);
 * ```
 */
export function listUnrecognizedValues(doc: DocxDocument): DocxUnrecognizedValue[] {
  checkHandle(doc, "document", "doc", "Document.listUnrecognizedValues");
  return findUnrecognizedValues(doc);
}

/**
 * Drop the values `Document.read` kept because they are outside their field's vocabulary — what `Validation.validate`
 * reports as `"unrecognized-value"` — so the document is written with each such field as it now reads. Returns how
 * many there were.
 *
 * Until then a kept value is written back as read while its field is absent; assigning the field takes over, but
 * clearing it again brings the kept text back. This is how to remove one instead — all of them, or with `filter` the
 * ones you choose:
 *
 * ```ts
 * Document.discardUnrecognizedValues(doc, { filter: v => v.path.endsWith(".alignment") });
 * ```
 */
export function discardUnrecognizedValues(
  doc: DocxDocument,
  options?: DocxDiscardUnrecognizedValuesOptions
): number {
  checkHandle(doc, "document", "doc", "Document.discardUnrecognizedValues");
  const where = "Document.discardUnrecognizedValues";
  checkRecord(options, "options", where, { optional: true });
  if (options?.filter !== undefined && typeof options.filter !== "function") {
    invalidArgument(where, "options.filter", "a function", options.filter);
  }
  ownTree(doc);
  return discardUnrecognizedValuesUnder(doc, options?.filter);
}
// =============================================================================
// Document Handle & Namespace
// =============================================================================

/** @deprecated Use {@link DocxDocument}. Removed in 1.0. */
export type DocumentHandle = DocxDocument;

type _Writable = { -readonly [K in keyof DocxDocument]: DocxDocument[K] };

/** The document's fields, writable — the helpers below assign through this. */
function _w(doc: DocxDocument): _Writable {
  return doc as _Writable;
}

/** The body, which `doc` may edit in place — see `@word/core/ownership`. */
function _body(doc: DocxDocument): BodyContent[] {
  return ownContainer(doc, "body", doc.body as BodyContent[]);
}

type _ArrayField = {
  [K in keyof DocxDocument]-?: NonNullable<DocxDocument[K]> extends readonly unknown[] ? K : never;
}[keyof DocxDocument];

/**
 * The collection `key`, created on the document when a helper is about to add
 * to it. Only writers call this: reading a collection never creates it, so a
 * document carries no empty collection a helper did not fill.
 */
function _ensure<K extends _ArrayField>(
  doc: DocxDocument,
  key: K
): (NonNullable<DocxDocument[K]> extends readonly (infer T)[] ? T : never)[] {
  const w = _w(doc) as Record<string, unknown>;
  const list = w[key];
  if (list === undefined) {
    const created: unknown[] = [];
    claim(doc, created);
    w[key] = created;
    return created as (NonNullable<DocxDocument[K]> extends readonly (infer T)[] ? T : never)[];
  }
  return ownContainer(
    doc,
    key,
    list as (NonNullable<DocxDocument[K]> extends readonly (infer T)[] ? T : never)[]
  );
}

function _ensureMap(
  doc: DocxDocument,
  key: "headers" | "footers"
): Map<string, HeaderDef | FooterDef> {
  const w = _w(doc) as Record<string, unknown>;
  const map = w[key] as Map<string, HeaderDef | FooterDef> | undefined;
  if (map === undefined) {
    const created = new Map<string, HeaderDef | FooterDef>();
    claim(doc, created);
    w[key] = created;
    return created;
  }
  return ownContainer(doc, key, map);
}

// -----------------------------------------------------------------------------
// Id allocation
// -----------------------------------------------------------------------------
//
// The document is a public tree, so callers may add footnotes, comments,
// numbering or bookmarks themselves between helper calls. A counter derived
// once and cached would then hand out an id the caller has since used, and the
// package would carry two `w:footnote w:id="3"`. So every allocation is
// computed from the document's current content: the next id is one above the
// larger of the highest id currently present and the highest this module has
// handed out for the document (`_issued`). The second term keeps ids monotonic:
// an id given out and then removed by the caller is not reused while the
// caller may still hold it.
//
// The table-backed kinds (footnote, endnote, comment, numbering, image) are
// referenced from elsewhere in the tree — `w:footnoteReference`, `w:numId`,
// `r:embed` — so a clash cannot be repaired at write time without guessing
// which reference meant which entry. Their ids live in one collection each,
// read incrementally (`_highestInUse`), which is exact for the edits the
// `Handle` contract supports — a new array, or entries appended to the end —
// and O(1) amortised. Telling an in-place edit at an already-read position
// apart from no edit would cost a pass over the collection per allocation, so
// that edit is outside the contract, and the writer's uniqueness check
// (`assertUniqueReferencedIds`) turns the duplicate it can cause into a
// `DocxDuplicateIdError` instead of a corrupt package.
//
// Bookmark and drawing ids are different: nothing refers to them. Bookmarks are
// referenced by *name* (fields, hyperlinks, the TOC); the id only pairs a start
// with its end. Drawing ids become `wp:docPr/@id`, which nothing references and
// which the packager already renumbers. So the writer reconciles both as it
// writes (`BookmarkIdMap`, `nextDocPrId`), and the helpers need only avoid the
// ids present when they were first asked: the tree — scattered ids, linear to
// walk — is scanned once per document to seed a high-water mark, and every
// later allocation is O(1).

type _IdKind =
  | "image"
  | "footnote"
  | "endnote"
  | "comment"
  | "abstractNum"
  | "num"
  | "bookmark"
  | "drawing";

const _issued = new WeakMap<DocxDocument, Partial<Record<_IdKind, number>>>();

/**
 * The number in `<prefix><digits><suffix…>` — `__img_7`, `image7.png` — or NaN.
 * Hand-rolled rather than a regex: it runs once per image per allocation.
 */
function _numberAfter(text: string | undefined, prefix: string, terminator: string): number {
  if (text === undefined || !text.startsWith(prefix)) {
    return Number.NaN;
  }
  let n = 0;
  let i = prefix.length;
  for (; i < text.length; i++) {
    const c = text.charCodeAt(i) - 48;
    if (c < 0 || c > 9) {
      break;
    }
    n = n * 10 + c;
  }
  if (i === prefix.length) {
    return Number.NaN;
  }
  return terminator === ""
    ? i === text.length
      ? n
      : Number.NaN
    : text[i] === terminator
      ? n
      : Number.NaN;
}

/**
 * Highest `drawingId` and bookmark id anywhere in the document's content —
 * body, headers, footers, notes and comments.
 *
 * Formatting (`properties` on a run, paragraph, table, row, cell or SDT) is
 * never content in the model, so it is not descended into; that is most of the
 * nodes in a typical document.
 */
function _scanTreeIds(doc: DocxDocument): { drawing: number; bookmark: number } {
  let drawing = 0;
  let bookmark = -1;
  const stack: unknown[] = [
    doc.body,
    doc.bodyRangeMarkersAfter,
    doc.footnotes,
    doc.endnotes,
    doc.comments
  ];
  for (const part of [doc.headers, doc.footers]) {
    for (const def of part?.values() ?? []) {
      stack.push(def);
    }
  }
  while (stack.length > 0) {
    const node = stack.pop();
    if (node === null || typeof node !== "object" || ArrayBuffer.isView(node)) {
      continue;
    }
    if (Array.isArray(node)) {
      for (const item of node) {
        if (item !== null && typeof item === "object") {
          stack.push(item);
        }
      }
      continue;
    }
    const record = node as Record<string, unknown>;
    const drawingId = record.drawingId;
    if (typeof drawingId === "number" && drawingId > drawing) {
      drawing = drawingId;
    }
    if (record.type === "bookmarkStart" && typeof record.id === "number" && record.id > bookmark) {
      bookmark = record.id;
    }
    for (const key in record) {
      if (key === "properties") {
        continue;
      }
      const value = record[key];
      if (value !== null && typeof value === "object") {
        stack.push(value);
      }
    }
  }
  return { drawing, bookmark };
}

type _TableKind = Exclude<_IdKind, "bookmark" | "drawing">;

function _imageNumber(img: ImageDef): number {
  const a = _numberAfter(img.rId, "__img_", "");
  const b = _numberAfter(img.fileName, "image", ".");
  return a > b || Number.isNaN(b) ? a : b;
}

/** The collection that holds `kind`'s ids, the id of one entry, and the floor. */
function _table(
  doc: DocxDocument,
  kind: _TableKind
): [list: readonly unknown[] | undefined, pick: (item: never) => number, floor: number] {
  switch (kind) {
    case "image":
      return [doc.images, _imageNumber, 0];
    case "footnote":
      return [doc.footnotes, (f: FootnoteDef) => f.id, 0];
    case "endnote":
      return [doc.endnotes, (e: EndnoteDef) => e.id, 0];
    case "comment":
      return [doc.comments, (c: CommentDef) => c.id, -1];
    case "abstractNum":
      return [doc.abstractNumberings, (a: AbstractNumbering) => a.abstractNumId, -1];
    case "num":
      return [doc.numberingInstances, (n: NumberingInstance) => n.numId, 0];
  }
}

/**
 * A collection as it was last read: which array, how many of its entries have
 * been read, and the highest id among them.
 */
interface _Scan {
  list: readonly unknown[];
  read: number;
  max: number;
}

const _scans = new WeakMap<DocxDocument, Partial<Record<_TableKind, _Scan>>>();

/**
 * Highest id of `kind` the document currently uses, or the kind's floor.
 *
 * A collection the document holds a new array for is read in full; the same
 * array is read from where the last call stopped. That is exact for every edit
 * the {@link Handle} contract supports — a new array, or entries appended to
 * the end — and costs O(1) amortised per call, so a loop of n additions is
 * linear. An entry replaced or edited in place at a position already read is
 * not seen here: the writer's uniqueness check reports the duplicate it
 * produces.
 */
function _highestInUse(doc: DocxDocument, kind: _TableKind): number {
  const [list, pick, floor] = _table(doc, kind);
  if (!list) {
    return floor;
  }
  let scans = _scans.get(doc);
  if (!scans) {
    scans = {};
    _scans.set(doc, scans);
  }
  let scan = scans[kind];
  if (!scan || scan.list !== list) {
    scan = { list, read: 0, max: floor };
    scans[kind] = scan;
  }
  const read = pick as (item: unknown) => number;
  for (; scan.read < list.length; scan.read++) {
    const v = read(list[scan.read]);
    if (Number.isFinite(v) && v > scan.max) {
      scan.max = v;
    }
  }
  return scan.max;
}

/** A fresh id of `kind`, unused anywhere in the document's current content. */
function _allocate(doc: DocxDocument, kind: _IdKind): number {
  let issued = _issued.get(doc);
  if (!issued) {
    issued = {};
    _issued.set(doc, issued);
  }
  if (kind === "bookmark" || kind === "drawing") {
    if (issued.bookmark === undefined || issued.drawing === undefined) {
      const seed = _scanTreeIds(doc);
      issued.bookmark ??= seed.bookmark;
      issued.drawing ??= seed.drawing;
    }
    const id = (issued[kind] ?? 0) + 1;
    issued[kind] = id;
    return id;
  }
  const id = Math.max(_highestInUse(doc, kind), issued[kind] ?? Number.NEGATIVE_INFINITY) + 1;
  issued[kind] = id;
  return id;
}

/**
 * Ensure a numbering instance exists for the given abstractNumId and return
 * its numId. If multiple instances reference the same abstract, the first
 * one is returned. Used by addBulletList to look up (or create) the runtime
 * numId after it ensures the abstract definition exists.
 */
function _ensureNumberingInstance(doc: DocxDocument, abstractNumId: number): number {
  const existing = doc.numberingInstances?.find(n => n.abstractNumId === abstractNumId);
  if (existing) {
    return existing.numId;
  }
  const numId = _allocate(doc, "num");
  _ensure(doc, "numberingInstances").push({ numId, abstractNumId });
  return numId;
}

/**
 * Namespace of free functions for building DOCX documents.
 *
 * Replaces the former `DocumentBuilder` class with tree-shakeable free functions.
 * Each function operates on a `DocxDocument` — a transparent content tree that
 * `Document.create()` and `Document.read` return and `Document.toBuffer` writes.
 *
 * @example
 * ```ts
 * import { Document } from "documonster/word";
 *
 * const doc = Document.create();
 * Document.addHeading(doc, "Hello World", 1);
 * Document.addParagraph(doc, "This is a paragraph.");
 * Document.addTable(doc, [["Name", "Age"], ["Alice", "30"]]);
 * const bytes = await Document.toBuffer(doc);
 * ```
 */
export function create(): DocxDocument {
  const doc: DocxDocument = {
    body: [],
    sectionProperties: {
      pageSize: { width: 12240, height: 15840 },
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    },
    deriveEastAsianDefaults: true
  };
  claim(doc, doc.body);
  return doc;
}

/**
 * A deep copy of `doc` that shares nothing with it, so editing either in place
 * never reaches the other.
 *
 * `Document.*` needs no clone to keep documents apart: it copies whatever a
 * document shares with another before changing it. Take one to keep a
 * snapshot, or before editing nodes by hand (assigning into a paragraph you
 * read from the tree), which nothing copies for you.
 *
 * The model is plain data — objects, arrays, `Map`s, byte arrays, `Date`s and
 * primitives — so this is a `structuredClone`; a value outside the model (a
 * function stored on a node, say) throws a `DataCloneError`.
 */
export function clone(doc: DocxDocument): DocxDocument {
  checkHandle(doc, "document", "doc", "Document.clone");
  const copy = structuredClone(doc);
  // Ids handed out for `doc` that its content does not show yet (a bookmark id
  // the caller still holds, say) are not handed out again for the copy.
  const issued = _issued.get(doc);
  if (issued) {
    _issued.set(copy, { ...issued });
  }
  return copy;
}

/** A block for the body: any node with a `type`, as `Build.*` makes them. */
function _checkBlock(content: unknown, field: string, where: string): void {
  checkBlock(content, field, where);
}

/**
 * Text or blocks, as the note, comment and text-box helpers take their
 * content: a string becomes one paragraph.
 */
function _paragraphs(content: unknown, field: string, where: string): Paragraph[] {
  if (typeof content === "string") {
    return [checkedTextParagraph(where, content, undefined)];
  }
  if (!Array.isArray(content)) {
    invalidArgument(where, field, "a string or an array of paragraphs", content);
  }
  content.forEach((p, i) => {
    checkRecord(p, `${field}[${i}]`, where);
    if ((p as { type?: unknown }).type !== "paragraph") {
      invalidArgument(where, `${field}[${i}].type`, '"paragraph"', (p as { type?: unknown }).type);
    }
  });
  return content as Paragraph[];
}

/**
 * A position in the body as a lookup or removal takes it: an integer. One
 * outside the body finds nothing, which the caller is told by `undefined`.
 */
function _checkPosition(index: unknown, where: string): void {
  checkInteger(index, "index", where);
}

/** Append a block — a node built with `Build.*` or by hand — to the body. */
export function addContent(doc: DocxDocument, content: BodyContent): void {
  checkHandle(doc, "document", "doc", "Document.addContent");
  _checkBlock(content, "content", "Document.addContent");
  _body(doc).push(content);
}

/**
 * Insert `content` before the block at `index`; `index` equal to the number
 * of blocks appends.
 *
 * @throws {DocxArgumentError} If `index` is not an integer from 0 to
 * the number of blocks. A position outside the body has no block to insert
 * before, and used to be read the way `Array.prototype.splice` reads it —
 * past the end appended, a negative index counted from the end.
 */
export function insertContent(doc: DocxDocument, index: number, content: BodyContent): void {
  checkHandle(doc, "document", "doc", "Document.insertContent");
  const where = "Document.insertContent";
  const body = _body(doc);
  checkInteger(index, "index", where, { min: 0, max: body.length });
  _checkBlock(content, "content", where);
  body.splice(index, 0, content);
}

/** @deprecated Use {@link insertContent}. Removed in 1.0. */
export const insertContentAt = insertContent;

/**
 * Remove the block at `index` and return it, or `undefined` when there is no
 * block there.
 *
 * @throws {DocxArgumentError} If `index` is not an integer.
 */
export function removeContent(doc: DocxDocument, index: number): BodyContent | undefined {
  checkHandle(doc, "document", "doc", "Document.removeContent");
  _checkPosition(index, "Document.removeContent");
  const body = _body(doc);
  if (index < 0 || index >= body.length) {
    return undefined;
  }
  if (isSharing(doc)) {
    // `removeBlock` takes the range markers off the block it removes, which
    // another document may still hold.
    body[index] = { ...body[index]! };
  }
  const { removed, markersAfter } = removeBlock(body, index);
  if (markersAfter.length > 0) {
    _w(doc).bodyRangeMarkersAfter = [...markersAfter, ...(doc.bodyRangeMarkersAfter ?? [])];
  }
  return removed;
}

/** The number of blocks in the body. */
export function getContentCount(doc: DocxDocument): number {
  checkHandle(doc, "document", "doc", "Document.getContentCount");
  return doc.body.length;
}

/**
 * The block at `index`, or `undefined` when there is none.
 *
 * @throws {DocxArgumentError} If `index` is not an integer.
 */
export function getContent(doc: DocxDocument, index: number): BodyContent | undefined {
  checkHandle(doc, "document", "doc", "Document.getContent");
  _checkPosition(index, "Document.getContent");
  return doc.body[index];
}

/**
 * Append a paragraph: either text (built with `properties`) or a ready-made
 * {@link Paragraph} node, e.g. from `Build.paragraph(...)`.
 */
export function addParagraph(
  doc: DocxDocument,
  content: string,
  properties?: ParagraphProperties & { run?: RunProperties }
): void;
export function addParagraph(doc: DocxDocument, para: Paragraph): void;
export function addParagraph(
  doc: DocxDocument,
  content: string | Paragraph,
  properties?: ParagraphProperties & { run?: RunProperties }
): void {
  checkHandle(doc, "document", "doc", "Document.addParagraph");
  const where = "Document.addParagraph";
  // Only a paragraph node is appended as it is. Anything else is text, and must be a string: the
  // text form used to put `null`, a number or `undefined` into a run, which no writer could emit.
  if (isRecordArgument(content)) {
    checkNode(content, "paragraph", "para", where);
    _body(doc).push(content);
  } else {
    _body(doc).push(checkedTextParagraph(where, content, properties));
  }
}

/** @deprecated Use {@link addParagraph}. Removed in 1.0. */
export function addParagraphElement(doc: DocxDocument, para: Paragraph): void {
  checkHandle(doc, "document", "doc", "Document.addParagraphElement");
  checkNode(para, "paragraph", "para", "Document.addParagraphElement");
  _body(doc).push(para);
}

/** Append a heading paragraph (`Heading1`…`Heading9` style) at `level`, 1 by default. */
export function addHeading(
  doc: DocxDocument,
  content: string | readonly ParagraphChild[],
  level: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 = 1
): void {
  checkHandle(doc, "document", "doc", "Document.addHeading");
  _body(doc).push(checkedHeading("Document.addHeading", content, level));
}

/** Append a paragraph holding a page break. */
export function addPageBreak(doc: DocxDocument): void {
  checkHandle(doc, "document", "doc", "Document.addPageBreak");
  _body(doc).push(paragraph([pageBreak()]));
}

/**
 * Append a table: either a grid of strings (built with `options`) or a
 * ready-made {@link Table} node, e.g. from `Build.table(...)`.
 */
export function addTable(
  doc: DocxDocument,
  data: readonly (readonly string[])[],
  options?: SimpleTableOptions
): void;
export function addTable(doc: DocxDocument, table: Table): void;
export function addTable(
  doc: DocxDocument,
  data: readonly (readonly string[])[] | Table,
  options?: SimpleTableOptions
): void {
  checkHandle(doc, "document", "doc", "Document.addTable");
  const where = "Document.addTable";
  // Only a table node is appended as it is; anything else is the rows of a simple table, which
  // reject what is not an array of strings rather than putting it into the body.
  if (isRecordArgument(data)) {
    checkNode(data, "table", "table", where);
    _body(doc).push(data);
  } else {
    _body(doc).push(checkedSimpleTable(where, data, options));
  }
}

/** @deprecated Use {@link addTable}. Removed in 1.0. */
export function addTableElement(doc: DocxDocument, tbl: Table): void {
  checkHandle(doc, "document", "doc", "Document.addTableElement");
  checkNode(tbl, "table", "tbl", "Document.addTableElement");
  _body(doc).push(tbl);
}

/** Image bytes and the options every image helper shares. */
function _checkImage(
  data: unknown,
  options: { altText?: unknown; name?: unknown; fallbackData?: unknown } | undefined,
  where: string
): { data: Uint8Array; fallbackData: Uint8Array | undefined } {
  const bytes = checkBytes(data, "data", where);
  checkRecord(options, "options", where, { optional: true });
  checkString(options?.altText, "options.altText", where, { optional: true });
  checkString(options?.name, "options.name", where, { optional: true });
  const fallbackData =
    options?.fallbackData === undefined
      ? undefined
      : checkBytes(options.fallbackData, "options.fallbackData", where);
  return { data: bytes, fallbackData };
}

/** What {@link addImage} takes beside the image bytes. */
export interface DocxAddImageOptions {
  /** The image format, as a file-extension name (`"png"`, `"jpeg"`, …), not a MIME type. */
  readonly type: ImageMediaType;
  /** Width in EMU — convert with `Units.cmToEmu`, `Units.inchesToEmu` or `Units.pxToEmu`. */
  readonly width: Emu;
  /** Height in EMU — convert with `Units.cmToEmu`, `Units.inchesToEmu` or `Units.pxToEmu`. */
  readonly height: Emu;
  /** Alternative text read by assistive technology. */
  readonly altText?: string;
  /** The drawing's name. Default: `Picture <n>`. */
  readonly name?: string;
  /** A raster fallback for a vector `type` (`svg`), shown by readers that cannot draw it. */
  readonly fallbackData?: Uint8Array;
}

/** What {@link addFloatingImage} takes beside the image bytes: {@link DocxAddImageOptions} plus the anchor. */
export interface DocxAddFloatingImageOptions extends DocxAddImageOptions {
  readonly horizontalPosition?: FloatingImage["horizontalPosition"];
  readonly verticalPosition?: FloatingImage["verticalPosition"];
  readonly wrap?: FloatingImage["wrap"];
  readonly behindDoc?: boolean;
  readonly lockAnchor?: boolean;
  readonly layoutInCell?: boolean;
  readonly allowOverlap?: boolean;
  /** Distance from the text above, in EMU. */
  readonly distT?: Emu;
  /** Distance from the text below, in EMU. */
  readonly distB?: Emu;
  /** Distance from the text to the left, in EMU. */
  readonly distL?: Emu;
  /** Distance from the text to the right, in EMU. */
  readonly distR?: Emu;
  /** Clockwise rotation in degrees. */
  readonly rotation?: number;
  readonly flipHorizontal?: boolean;
  readonly flipVertical?: boolean;
}

/**
 * The options of an image helper called either as `(doc, data, options)` or in the deprecated
 * `(doc, data, mediaType, width, height, options?)` form, told apart by whether the third argument is
 * an options record. In the options form nothing may follow it: a width there would be silently ignored.
 */
function _imageOptions<O extends DocxAddImageOptions>(
  where: string,
  optionsOrType: O | ImageMediaType,
  width: unknown,
  height: unknown,
  rest: Omit<O, "type" | "width" | "height"> | undefined
): O {
  if (isRecordArgument(optionsOrType)) {
    if (width !== undefined || height !== undefined || rest !== undefined) {
      invalidArgument(
        where,
        "options",
        "the last argument after the image bytes",
        width !== undefined ? width : height !== undefined ? height : rest
      );
    }
    return optionsOrType;
  }
  checkRecord(rest, "options", where, { optional: true });
  return { ...rest, type: optionsOrType, width, height } as O;
}

/**
 * Append a paragraph holding an inline picture and register the image bytes with the
 * document: `Document.addImage(doc, png, { type: "png", width: Units.cmToEmu(5), height: Units.cmToEmu(3) })`.
 * Returns the relationship id the picture refers to the image by and its drawing id (`wp:docPr/@id`).
 */
export function addImage(
  doc: DocxDocument,
  data: Uint8Array,
  options: DocxAddImageOptions
): { rId: string; drawingId: number };
/**
 * @deprecated Pass the format and size in the options:
 * `Document.addImage(doc, data, { type, width, height })`. Removed in 1.0.
 */
export function addImage(
  doc: DocxDocument,
  data: Uint8Array,
  mediaType: ImageMediaType,
  width: Emu,
  height: Emu,
  options?: { altText?: string; name?: string; fallbackData?: Uint8Array }
): { rId: string; drawingId: number };
export function addImage(
  doc: DocxDocument,
  data: Uint8Array,
  optionsOrType: DocxAddImageOptions | ImageMediaType,
  widthArg?: Emu,
  heightArg?: Emu,
  rest?: { altText?: string; name?: string; fallbackData?: Uint8Array }
): { rId: string; drawingId: number } {
  checkHandle(doc, "document", "doc", "Document.addImage");
  const options = _imageOptions("Document.addImage", optionsOrType, widthArg, heightArg, rest);
  const record = isRecordArgument(optionsOrType);
  const { type: mediaType, width, height } = options;
  assertImageMediaType(mediaType, "Document.addImage", record ? "options.type" : "mediaType");
  assertDrawingNumbers(
    "Document.addImage",
    { width, height },
    record ? under("options") : undefined
  );
  const bytes = _checkImage(data, options, "Document.addImage");
  const imageId = _allocate(doc, "image");
  const fileName = `image${imageId}.${mediaType}`;
  const rId = `__img_${imageId}`;
  const drawingId = _allocate(doc, "drawing");

  _ensure(doc, "images").push({
    data: bytes.data,
    mediaType,
    fileName,
    rId,
    fallbackData: bytes.fallbackData
  });

  _body(doc).push(
    paragraph([
      {
        content: [
          {
            type: "image",
            rId,
            width,
            height,
            altText: options?.altText,
            name: options?.name ?? `Picture ${imageId}`,
            drawingId
          }
        ]
      }
    ])
  );

  return { rId, drawingId };
}

/**
 * Append a floating (anchored) picture and register the image bytes with the document:
 * `Document.addFloatingImage(doc, png, { type: "png", width, height, wrap: { style: "square" } })`.
 * Returns the relationship id the picture refers to the image by.
 */
export function addFloatingImage(
  doc: DocxDocument,
  data: Uint8Array,
  options: DocxAddFloatingImageOptions
): string;
/**
 * @deprecated Pass the format and size in the options:
 * `Document.addFloatingImage(doc, data, { type, width, height, … })`. Removed in 1.0.
 */
export function addFloatingImage(
  doc: DocxDocument,
  data: Uint8Array,
  mediaType: ImageMediaType,
  width: Emu,
  height: Emu,
  options?: Omit<DocxAddFloatingImageOptions, "type" | "width" | "height">
): string;
export function addFloatingImage(
  doc: DocxDocument,
  data: Uint8Array,
  optionsOrType: DocxAddFloatingImageOptions | ImageMediaType,
  widthArg?: Emu,
  heightArg?: Emu,
  rest?: Omit<DocxAddFloatingImageOptions, "type" | "width" | "height">
): string {
  checkHandle(doc, "document", "doc", "Document.addFloatingImage");
  const options = _imageOptions(
    "Document.addFloatingImage",
    optionsOrType,
    widthArg,
    heightArg,
    rest
  );
  const record = isRecordArgument(optionsOrType);
  const { type: mediaType, width, height } = options;
  assertImageMediaType(
    mediaType,
    "Document.addFloatingImage",
    record ? "options.type" : "mediaType"
  );
  assertDrawingNumbers(
    "Document.addFloatingImage",
    { width, height },
    record ? under("options") : undefined
  );
  assertAnchorNumbers("Document.addFloatingImage", options, under("options"));
  const bytes = _checkImage(data, options, "Document.addFloatingImage");
  const imageId = _allocate(doc, "image");
  const fileName = `image${imageId}.${mediaType}`;
  const rId = `__img_${imageId}`;

  _ensure(doc, "images").push({
    data: bytes.data,
    mediaType,
    fileName,
    rId,
    fallbackData: bytes.fallbackData
  });

  _body(doc).push(
    floatingImage({
      rId,
      width,
      height,
      altText: options?.altText,
      name: options?.name ?? `Picture ${imageId}`,
      horizontalPosition: options?.horizontalPosition,
      verticalPosition: options?.verticalPosition,
      wrap: options?.wrap,
      behindDoc: options?.behindDoc,
      lockAnchor: options?.lockAnchor,
      layoutInCell: options?.layoutInCell,
      allowOverlap: options?.allowOverlap,
      distT: options?.distT,
      distB: options?.distB,
      distL: options?.distL,
      distR: options?.distR,
      rotation: options?.rotation,
      flipHorizontal: options?.flipHorizontal,
      flipVertical: options?.flipVertical
    })
  );

  return rId;
}

/** Add a font table entry (`word/fontTable.xml`). */
export function addFont(doc: DocxDocument, font: FontDef): void {
  checkHandle(doc, "document", "doc", "Document.addFont");
  checkRecord(font, "font", "Document.addFont");
  checkString(font.name, "font.name", "Document.addFont");
  checkString(font.altName, "font.altName", "Document.addFont", { optional: true });
  _ensure(doc, "fonts").push(font);
}

/**
 * Set the document's watermark, replacing any it had. It is drawn in every
 * header a section shows, a default header being added where there is none.
 *
 * `undefined` takes the watermark off — one the document was read with
 * included, since reading lifts a watermark every header shows out of the
 * headers into `doc.watermark`.
 */
export function setWatermark(doc: DocxDocument, watermark: Watermark | undefined): void {
  checkHandle(doc, "document", "doc", "Document.setWatermark");
  if (watermark === undefined) {
    delete _w(doc).watermark;
    return;
  }
  checkWatermark(watermark, "watermark", "Document.setWatermark");
  _w(doc).watermark = watermark;
}

/**
 * Add a footnote and return its id. This adds the note only: reference it
 * from the body with a run whose content is `{ type: "footnoteRef", id }`.
 */
export function addFootnote(doc: DocxDocument, content: string | readonly Paragraph[]): number {
  checkHandle(doc, "document", "doc", "Document.addFootnote");
  const paras = _paragraphs(content, "content", "Document.addFootnote");
  const id = _allocate(doc, "footnote");
  _ensure(doc, "footnotes").push({ id, content: paras });
  return id;
}

/**
 * Add an endnote and return its id. This adds the note only: reference it
 * from the body with a run whose content is `{ type: "endnoteRef", id }`.
 */
export function addEndnote(doc: DocxDocument, content: string | readonly Paragraph[]): number {
  checkHandle(doc, "document", "doc", "Document.addEndnote");
  const paras = _paragraphs(content, "content", "Document.addEndnote");
  const id = _allocate(doc, "endnote");
  _ensure(doc, "endnotes").push({ id, content: paras });
  return id;
}

/**
 * Add a comment and return its id. This adds the comment only: anchor it with
 * `Build.commentRangeStart(id)`, `Build.commentRangeEnd(id)` and
 * `Build.commentReference(id)`. `date` is an `xsd:dateTime`.
 */
export function addComment(
  doc: DocxDocument,
  author: string,
  content: string | readonly Paragraph[],
  options?: { date?: string; initials?: string }
): number {
  checkHandle(doc, "document", "doc", "Document.addComment");
  const where = "Document.addComment";
  checkString(author, "author", where);
  const paras = _paragraphs(content, "content", where);
  checkRecord(options, "options", where, { optional: true });
  checkDateTime(options?.date, "options.date", where, { optional: true });
  checkString(options?.initials, "options.initials", where, { optional: true });
  const id = _allocate(doc, "comment");
  _ensure(doc, "comments").push({
    id,
    author,
    date: options?.date,
    initials: options?.initials,
    content: paras
  });
  return id;
}

/**
 * Append a table of contents field — headings 1–3, as hyperlinks, unless
 * `options` (or its `instruction`) says otherwise. Its entries are filled in
 * by `Transform.updateTableOfContents`, or by Word when the field is updated.
 */
export function addTableOfContents(
  doc: DocxDocument,
  options?: Partial<Omit<TableOfContents, "type">>
): void {
  checkHandle(doc, "document", "doc", "Document.addTableOfContents");
  checkTableOfContentsOptions(options, "options", "Document.addTableOfContents");
  if (options?.instruction !== undefined) {
    // The options default to what the given instruction says, not to `\o "1-3" \h`.
    const read: Mutable<TableOfContents> = { type: "tableOfContents" };
    parseTocInstruction(options.instruction, read);
    _body(doc).push({ ...read, ...options });
    return;
  }
  _body(doc).push({
    type: "tableOfContents",
    headingStyleRange: options?.headingStyleRange ?? "1-3",
    hyperlink: options?.hyperlink ?? true,
    ...options
  });
}

/** Append a display math block (`m:oMathPara`). */
export function addMath(doc: DocxDocument, content: readonly MathContent[]): void {
  checkHandle(doc, "document", "doc", "Document.addMath");
  checkNodeArray(content, "content", "Document.addMath");
  _body(doc).push(mathBlock(content));
}

/** Append a text box holding `content`; `width` and `height` are in twips. */
export function addTextBox(
  doc: DocxDocument,
  content: string | readonly Paragraph[],
  options?: { width?: Twips; height?: Twips; stroke?: boolean; fill?: boolean }
): void {
  checkHandle(doc, "document", "doc", "Document.addTextBox");
  const sized = { sign: "nonNegative", optional: true } as const;
  const paras = _paragraphs(content, "content", "Document.addTextBox");
  checkRecord(options, "options", "Document.addTextBox", { optional: true });
  assertDrawingNumber(options?.width, "options.width", "Document.addTextBox", sized);
  assertDrawingNumber(options?.height, "options.height", "Document.addTextBox", sized);
  checkBoolean(options?.stroke, "options.stroke", "Document.addTextBox", { optional: true });
  checkBoolean(options?.fill, "options.fill", "Document.addTextBox", { optional: true });
  _body(doc).push({
    type: "textBox",
    content: paras,
    width: options?.width,
    height: options?.height,
    stroke: options?.stroke,
    fill: options?.fill
  });
}

/** Kind of list definition the builder writes. */
export type ListKind = "bullet" | "numbered";

/**
 * The builder's abstract definition of `kind` — the first one whose level 0
 * has the matching format — created when the document has none yet.
 */
function _ensureListAbstract(doc: DocxDocument, kind: ListKind): number {
  const format = kind === "bullet" ? "bullet" : "decimal";
  const existing = doc.abstractNumberings?.find(a => a.levels[0]?.format === format);
  if (existing) {
    return existing.abstractNumId;
  }
  const abstractNumId = _allocate(doc, "abstractNum");
  if (kind === "bullet") {
    _ensure(doc, "abstractNumberings").push({
      abstractNumId,
      multiLevelType: "hybridMultilevel",
      levels: [
        {
          level: 0,
          start: 1,
          format: "bullet",
          text: "\uF0B7",
          justification: "left",
          paragraphProperties: { indent: { left: 720, hanging: 360 } },
          runProperties: { font: { ascii: "Symbol", hAnsi: "Symbol" } }
        },
        {
          level: 1,
          start: 1,
          format: "bullet",
          text: "o",
          justification: "left",
          paragraphProperties: { indent: { left: 1440, hanging: 360 } },
          runProperties: { font: { ascii: "Courier New", hAnsi: "Courier New" } }
        },
        {
          level: 2,
          start: 1,
          format: "bullet",
          text: "\uF0A7",
          justification: "left",
          paragraphProperties: { indent: { left: 2160, hanging: 360 } },
          runProperties: { font: { ascii: "Wingdings", hAnsi: "Wingdings" } }
        }
      ]
    });
  } else {
    _ensure(doc, "abstractNumberings").push({
      abstractNumId,
      multiLevelType: "hybridMultilevel",
      levels: [
        {
          level: 0,
          start: 1,
          format: "decimal",
          text: "%1.",
          justification: "left",
          paragraphProperties: { indent: { left: 720, hanging: 360 } }
        },
        {
          level: 1,
          start: 1,
          format: "lowerLetter",
          text: "%2.",
          justification: "left",
          paragraphProperties: { indent: { left: 1440, hanging: 360 } }
        },
        {
          level: 2,
          start: 1,
          format: "lowerRoman",
          text: "%3.",
          justification: "right",
          paragraphProperties: { indent: { left: 2160, hanging: 180 } }
        }
      ]
    });
  }
  return abstractNumId;
}

/**
 * Register a list definition — the builder's bullet or numbered abstract
 * numbering, created if the document has none — and add a numbering instance
 * for it. Returns the instance's `numId`, for paragraphs to reference through
 * `numbering: { numId, level }`.
 *
 * Pass `numId` to pin the instance to an id content already refers to (an
 * importer that emitted list paragraphs before the document existed, say);
 * otherwise a fresh id is allocated.
 *
 * @throws {DocxArgumentError} If `numId` is already used by another
 * instance.
 */
export function addListDefinition(
  doc: DocxDocument,
  kind: ListKind,
  options?: { numId?: number }
): number {
  checkHandle(doc, "document", "doc", "Document.addListDefinition");
  const where = "Document.addListDefinition";
  checkOneOf(kind, LIST_KINDS, "kind", where);
  checkRecord(options, "options", where, { optional: true });
  checkInteger(options?.numId, "options.numId", where, { optional: true, min: 0 });
  const requested = options?.numId;
  if (requested !== undefined && doc.numberingInstances?.some(n => n.numId === requested)) {
    invalidArgument(where, "options.numId", "a numId no numbering instance uses", requested);
  }
  const abstractNumId = _ensureListAbstract(doc, kind);
  const numId = requested ?? _allocate(doc, "num");
  _ensure(doc, "numberingInstances").push({ numId, abstractNumId });
  return numId;
}

const LIST_KINDS: ReadonlySet<string> = new Set<ListKind>(["bullet", "numbered"]);

function _checkListItems(items: unknown, level: unknown, where: string): void {
  checkArray(items, "items", where);
  (items as unknown[]).forEach((item, i) => {
    if (typeof item !== "string") {
      if (!Array.isArray(item)) {
        invalidArgument(where, `items[${i}]`, "a string or an array of runs", item);
      }
      checkNodeArray(item, `items[${i}]`, where);
    }
  });
  checkListLevel(level, "level", where);
}

function _pushListItems(
  doc: DocxDocument,
  items: readonly (string | readonly ParagraphChild[])[],
  numId: number,
  level: number,
  where: string
): void {
  const body = _body(doc);
  for (const item of items) {
    if (typeof item === "string") {
      body.push(checkedTextParagraph(where, item, { numbering: { numId, level } }));
    } else {
      body.push(paragraph(item, { numbering: { numId, level } }));
    }
  }
}

/**
 * Append one paragraph per item as a bullet list at `level` (0–8), defining
 * the document's bullet list the first time.
 */
export function addBulletList(
  doc: DocxDocument,
  items: readonly (string | readonly ParagraphChild[])[],
  level = 0
): void {
  checkHandle(doc, "document", "doc", "Document.addBulletList");
  _checkListItems(items, level, "Document.addBulletList");
  const numId = _ensureNumberingInstance(doc, _ensureListAbstract(doc, "bullet"));
  _pushListItems(doc, items, numId, level, "Document.addBulletList");
}

/**
 * Append one paragraph per item as a numbered list at `level` (0–8). Each call
 * is a list of its own, numbered from 1.
 */
export function addNumberedList(
  doc: DocxDocument,
  items: readonly (string | readonly ParagraphChild[])[],
  level = 0
): void {
  checkHandle(doc, "document", "doc", "Document.addNumberedList");
  _checkListItems(items, level, "Document.addNumberedList");
  const abstractNumId = _ensureListAbstract(doc, "numbered");
  // Each call is a list of its own. Word counts per numbering instance and
  // continues across interruptions — even across instances that share an
  // abstract definition — so reusing one instance numbered a second list on
  // from the first (1, 2, … 3, 4). A fresh instance whose `w:startOverride`
  // restates the level's start makes the restart explicit in the file, so
  // Word, the layout and every converter number it from the start.
  const numId = _allocate(doc, "num");
  const start = doc.abstractNumberings
    ?.find(a => a.abstractNumId === abstractNumId)
    ?.levels.find(l => l.level === level)?.start;
  _ensure(doc, "numberingInstances").push({
    numId,
    abstractNumId,
    overrides: [{ level, startOverride: start ?? 1 }]
  });
  _pushListItems(doc, items, numId, level, "Document.addNumberedList");
}

/**
 * Set the properties of the document's last section (the body's own
 * `sectionProperties`), replacing the ones it had.
 *
 * The section's header and footer references are kept when `props` gives none
 * (`headers` / `footers` absent): they are what wires the document's headers
 * and footers into the section, and a read document's are not expressible
 * without its relationship ids. So is `titlePage` while a kept reference is a
 * first-page one. Pass `headers: []` / `footers: []` to drop them.
 */
export function setSectionProperties(doc: DocxDocument, props: SectionProperties): void {
  checkHandle(doc, "document", "doc", "Document.setSectionProperties");
  checkRecord(props, "props", "Document.setSectionProperties");
  checkSectionProperties(props, "props", "Document.setSectionProperties");
  const previous = doc.sectionProperties;
  const headers = props.headers ?? previous?.headers;
  const footers = props.footers ?? previous?.footers;
  const keptFirst =
    (props.headers === undefined && previous?.headers?.some(r => r.type === "first")) ||
    (props.footers === undefined && previous?.footers?.some(r => r.type === "first"));
  _w(doc).sectionProperties = {
    ...props,
    ...(headers !== undefined && { headers }),
    ...(footers !== undefined && { footers }),
    ...(props.titlePage === undefined &&
      keptFirst &&
      previous?.titlePage !== undefined && { titlePage: previous.titlePage })
  };
}

/**
 * Set the document defaults (`w:docDefaults`: run and paragraph properties
 * every style starts from), replacing the ones it had.
 */
export function setDocDefaults(doc: DocxDocument, defaults: DocDefaults): void {
  checkHandle(doc, "document", "doc", "Document.setDocDefaults");
  const where = "Document.setDocDefaults";
  checkRecord(defaults, "defaults", where);
  // A misspelt key (`{ run: … }` for `runProperties`) used to be stored, then
  // written as an empty `<w:docDefaults/>` — which Word fills from the user's
  // Normal template — with nothing reported.
  const unknown = Object.keys(defaults).filter(
    key => key !== "runProperties" && key !== "paragraphProperties"
  );
  if (unknown.length > 0) {
    invalidArgument(
      where,
      `defaults.${unknown[0]}`,
      "absent: setDocDefaults takes runProperties and paragraphProperties",
      defaults[unknown[0] as keyof DocDefaults]
    );
  }
  checkRunProperties(defaults.runProperties, "defaults.runProperties", where);
  checkParagraphProperties(defaults.paragraphProperties, "defaults.paragraphProperties", where);
  _w(doc).docDefaults = defaults;
}

/**
 * Add a style definition. A style id names one style in a document
 * (`w:styleId` is unique in `styles.xml`), so a style whose id the document
 * already has replaces that style, in its place — the way to customise one of
 * the styles `useDefaultStyles` adds. A default style (`isDefault`) takes the
 * default over from the type's previous one.
 */
export function addStyle(doc: DocxDocument, style: StyleDef): void {
  checkHandle(doc, "document", "doc", "Document.addStyle");
  checkStyleDef(style, "style", "Document.addStyle");
  const styles = _ensure(doc, "styles");
  const same = styles.findIndex(existing => existing.styleId === style.styleId);
  // A type has one default style. Adding another default used to leave two —
  // `useDefaultStyles` then `addStyle({ isDefault: true })` — and the document
  // meant whichever one a reader happened to pick.
  if (style.isDefault) {
    for (let i = 0; i < styles.length; i++) {
      const existing = styles[i]!;
      if (existing.type === style.type && existing.isDefault) {
        styles[i] = { ...existing, isDefault: false };
      }
    }
  }
  if (same >= 0) {
    styles[same] = style;
  } else {
    styles.push(style);
  }
}

/**
 * Give the document the built-in style set — `Normal`, `Heading1`…`Heading9`,
 * `Title`, list, table, header/footer and note styles — and the document
 * defaults they assume (Calibri 11pt, 8pt after each paragraph). A style the
 * document defines under one of those ids is replaced by the built-in one;
 * a document default the document already states is kept. Call it before
 * customising a style with {@link addStyle}.
 */
export function useDefaultStyles(doc: DocxDocument): void {
  checkHandle(doc, "document", "doc", "Document.useDefaultStyles");
  const existing = doc.docDefaults;
  _w(doc).docDefaults = {
    runProperties: {
      // No `eastAsia` face and no `language.eastAsia`: this function runs before any
      // content exists, so it cannot know which East Asian language the document is
      // in, and both were previously answered inconsistently — `SimSun` was stated
      // while the matching `zh-CN` was refused on the grounds that the language was
      // unknowable. Naming a Chinese face is exactly the same claim as naming a
      // Chinese language, so either both are guesses or neither is.
      //
      // Writing the document fills both in from its own text, which is where the
      // answer actually is. A caller who wants to decide for itself sets them here
      // through `setDocDefaults` and writing leaves them alone.
      font: { ascii: "Calibri", hAnsi: "Calibri", cs: "Times New Roman" },
      size: 22,
      sizeCs: 22,
      language: { val: "en-US" },
      // Anything the caller already stated wins. This is "supply the defaults I have
      // not set", not "discard my settings and take mine": overwriting `docDefaults`
      // outright meant a `setDocDefaults` call followed by `useDefaultStyles` silently
      // lost the caller's size, font and spacing.
      ...existing?.runProperties
    },
    paragraphProperties: {
      spacing: { after: 160, line: 259, lineRule: "auto" },
      ...existing?.paragraphProperties
    }
  };

  addDefaultStyles(doc);
}

/**
 * Add the built-in style set, replacing any entry that already uses the same id.
 *
 * Appending unconditionally made a second `useDefaultStyles()` call produce 38 styles
 * with 19 distinct ids — a duplicate definition of every one of them, and a document
 * whose `styles.xml` declares `Normal` twice.
 */
function addDefaultStyles(doc: DocxDocument): void {
  const defaults = defaultStyleSet();

  const styles = _ensure(doc, "styles");
  for (const builtIn of defaults) {
    const at = styles.findIndex(
      existing => existing.styleId === builtIn.styleId && existing.type === builtIn.type
    );
    // A type has one default style; a document that already names another
    // (a localized `Normal`, say) keeps it.
    const style =
      builtIn.isDefault && styles.some((s, i) => i !== at && s.type === builtIn.type && s.isDefault)
        ? { ...builtIn, isDefault: false }
        : builtIn;
    if (at >= 0) {
      styles[at] = style;
    } else {
      styles.push(style);
    }
  }
}

/**
 * Set the header shown on the pages `key` names, in the document's last
 * section.
 *
 * `key` is a header type — `"default"`, `"first"` or `"even"` — or the key
 * under which the document already stores a header (a document from `Document.read`
 * keys its headers by relationship id, `"rId7"`), or a key a section's
 * header references name (`{ type: "default", rId: "cover" }`).
 *
 * - Given a type, it replaces the header the last section references for
 *   that type, whatever key it is stored under, so the new content is what
 *   the section shows. When the section references none of that type, the
 *   header is stored under the type and the section given the reference
 *   `{ type, rId: type }`. The section always ends up with an explicit
 *   reference, so `Layout.*` (and `Pdf.fromDocx`) show what the written file
 *   shows.
 * - `"first"` turns on the section's title page (`titlePage`), without which
 *   Word ignores a first-page header; `"even"` turns on
 *   `settings.evenAndOddHeaders`, likewise.
 * - Given a stored or referenced key, it sets the header under that key.
 *
 * @throws {DocxArgumentError} If `key` is none of these — such a header
 * would be written but shown on no page.
 */
export function setHeader(
  doc: DocxDocument,
  key: HeaderFooterType | (string & {}),
  content: HeaderFooterContent
): void {
  checkHandle(doc, "document", "doc", "Document.setHeader");
  _setHeaderFooter(doc, "headers", key, content, "Document.setHeader");
}

function _checkHeaderFooter(key: unknown, content: unknown, where: string): void {
  checkString(key, "key", where);
  checkRecord(content, "content", where);
  const children = (content as HeaderFooterContent).children;
  checkNodeArray(children, "content.children", where);
  children.forEach((child, i) => {
    const type = (child as { type?: unknown }).type;
    if (type !== "paragraph" && type !== "table") {
      invalidArgument(where, `content.children[${i}].type`, 'one of "paragraph", "table"', type);
    }
  });
}

/** Set the footer shown on the pages `key` names; see {@link setHeader} for the key. */
export function setFooter(
  doc: DocxDocument,
  key: HeaderFooterType | (string & {}),
  content: HeaderFooterContent
): void {
  checkHandle(doc, "document", "doc", "Document.setFooter");
  _setHeaderFooter(doc, "footers", key, content, "Document.setFooter");
}

/** Whether a section of `doc` — the last one, or one a section break ends — references `key`. */
function _sectionsReference(doc: DocxDocument, kind: "headers" | "footers", key: string): boolean {
  if (doc.sectionProperties?.[kind]?.some(r => r.rId === key)) {
    return true;
  }
  return doc.body.some(
    block =>
      block.type === "paragraph" &&
      block.properties?.sectionProperties?.[kind]?.some(r => r.rId === key) === true
  );
}

function _setHeaderFooter(
  doc: DocxDocument,
  kind: "headers" | "footers",
  key: string,
  content: HeaderFooterContent,
  where: string
): void {
  _checkHeaderFooter(key, content, where);
  const stored = doc[kind];
  if (!HEADER_FOOTER_TYPES.has(key)) {
    if (!stored?.has(key) && !_sectionsReference(doc, kind, key)) {
      invalidArgument(
        where,
        "key",
        `"default", "first", "even", the key of one of the document's ${kind} or one a section references`,
        key
      );
    }
    _ensureMap(doc, kind).set(key, { content });
    return;
  }
  const type = key as HeaderFooterType;
  const section = doc.sectionProperties;
  // Always an explicit reference in the section, so the layout and the
  // packager read the same thing. A hand-built model's implied type-keyed
  // parts become explicit first, so setting one type does not hide another.
  const own = section?.[kind];
  const current = own && own.length > 0 ? own : impliedReferences(stored);
  const ref = current.find(r => r.type === type);
  // The part the section shows for this type, when it names one.
  const target = (ref && referencedPartKey(ref, stored)) ?? type;
  const named = { type, rId: target };
  const next = ref
    ? current.map(r => (r === ref ? (ref.rId === target ? ref : named) : r))
    : [...current, named];
  if (next !== own) {
    _w(doc).sectionProperties = { ...section, [kind]: next };
  }
  _ensureMap(doc, kind).set(target, { content });
  if (type === "first" && !doc.sectionProperties?.titlePage) {
    _w(doc).sectionProperties = { ...doc.sectionProperties, titlePage: true };
  } else if (type === "even" && !doc.settings?.evenAndOddHeaders) {
    _w(doc).settings = { ...doc.settings, evenAndOddHeaders: true };
  }
}

/**
 * What `setSettings` keeps from the settings it replaces unless the new ones
 * name it: markup carried verbatim, the compatibility mode and options, the
 * protection `Security.*` sets, and the file's own bookkeeping — none of which
 * a caller setting zoom or tab stops means to drop.
 */
const KEPT_SETTINGS = [
  "preservedXml",
  "mailMergeRawXml",
  // Set by `setHeader` / `setFooter` for an even-page part, which needs it.
  "evenAndOddHeaders",
  "compatibilityMode",
  "compatSettings",
  "compatFlags",
  "documentProtection",
  "writeProtection",
  "rsids",
  "footnoteSeparators",
  "endnoteSeparators",
  UNRECOGNIZED_KEY
] as const;

/**
 * Set the document settings (`word/settings.xml`), replacing the ones it had.
 *
 * What a settings object cannot say, or that other functions manage, is kept
 * unless `settings` has the key (even as `undefined`): verbatim markup the
 * reader kept (`preservedXml`, `mailMergeRawXml`), `evenAndOddHeaders` (which
 * an even-page header set with {@link setHeader} needs), the compatibility mode and
 * options (`compatibilityMode`, `compatSettings`, `compatFlags` — see
 * `Document.setCompatibilityMode`), protection (`documentProtection`,
 * `writeProtection` — see `Security.protect`), revision ids (`rsids`), note
 * separators and values kept as read. Everything else is replaced: to change
 * one setting, pass `{ ...doc.settings, zoom: 120 }`.
 */
export function setSettings(doc: DocxDocument, settings: DocumentSettings): void {
  checkHandle(doc, "document", "doc", "Document.setSettings");
  checkDocumentSettings(settings, "settings", "Document.setSettings");
  const previous = doc.settings as Record<string, unknown> | undefined;
  const kept: Record<string, unknown> = {};
  for (const key of KEPT_SETTINGS) {
    if (previous && Object.hasOwn(previous, key) && !Object.hasOwn(settings, key)) {
      kept[key] = previous[key];
    }
  }
  const next = { ...kept, ...settings } as Record<string, unknown>;
  for (const key of KEPT_SETTINGS) {
    if (Object.hasOwn(settings, key) && settings[key as keyof DocumentSettings] === undefined) {
      delete next[key];
    }
  }
  _w(doc).settings = next as DocumentSettings;
}

/**
 * Set the core properties (`docProps/core.xml`: title, author, dates, …),
 * replacing the ones the document had.
 */
export function setCoreProperties(doc: DocxDocument, props: CoreProperties): void {
  checkHandle(doc, "document", "doc", "Document.setCoreProperties");
  checkCoreProperties(props, "props", "Document.setCoreProperties");
  _w(doc).coreProperties = props;
}

/**
 * Set the application properties (`docProps/app.xml`: company, manager,
 * statistics), replacing the ones the document had.
 */
export function setAppProperties(doc: DocxDocument, props: AppProperties): void {
  checkHandle(doc, "document", "doc", "Document.setAppProperties");
  checkAppProperties(props, "props", "Document.setAppProperties");
  _w(doc).appProperties = props;
}

/** Set the page background colour, replacing any the document had; `undefined` takes it off. */
export function setBackground(doc: DocxDocument, background: DocumentBackground | undefined): void {
  checkHandle(doc, "document", "doc", "Document.setBackground");
  const where = "Document.setBackground";
  if (background === undefined) {
    delete _w(doc).background;
    return;
  }
  checkRecord(background, "background", where);
  checkHexColor(background.color, "background.color", where, { optional: true });
  checkString(background.themeColor, "background.themeColor", where, { optional: true });
  _w(doc).background = background;
}

const CUSTOM_PROPERTY_TYPES: Record<CustomPropertyValue["type"], (v: unknown) => boolean> = {
  string: v => typeof v === "string",
  number: v => typeof v === "number" && Number.isFinite(v),
  boolean: v => typeof v === "boolean",
  date: v => v instanceof Date && !Number.isNaN(v.getTime())
};

/**
 * Add a custom document property (`docProps/custom.xml`). A property of the
 * same name is replaced, in its place: a name identifies one property.
 */
export function addCustomProperty(
  doc: DocxDocument,
  name: string,
  value: CustomPropertyValue
): void {
  checkHandle(doc, "document", "doc", "Document.addCustomProperty");
  const where = "Document.addCustomProperty";
  checkString(name, "name", where);
  if (name === "") {
    invalidArgument(where, "name", "a non-empty string", name);
  }
  checkRecord(value, "value", where);
  const accepts = Object.hasOwn(CUSTOM_PROPERTY_TYPES, value.type)
    ? CUSTOM_PROPERTY_TYPES[value.type]
    : undefined;
  if (accepts === undefined) {
    checkOneOf(value.type, new Set(Object.keys(CUSTOM_PROPERTY_TYPES)), "value.type", where);
  } else if (!accepts(value.value)) {
    invalidArgument(
      where,
      "value.value",
      value.type === "number"
        ? "a finite number"
        : value.type === "date"
          ? "a valid Date"
          : `a ${value.type}`,
      value.value
    );
  }
  // A name identifies one property in `docProps/custom.xml`: a second with
  // the same name would be written beside the first, and a reader takes either.
  const properties = _ensure(doc, "customProperties");
  const at = properties.findIndex(p => p.name === name);
  if (at >= 0) {
    properties[at] = { name, value };
  } else {
    properties.push({ name, value });
  }
}

/**
 * End a section after the last block: `props` become the properties of the
 * section the body so far belongs to (on the last paragraph, or on an empty
 * paragraph added for them). The document's last section keeps its own
 * properties — see {@link setSectionProperties}.
 */
export function addSectionBreak(doc: DocxDocument, props: SectionProperties): void {
  checkHandle(doc, "document", "doc", "Document.addSectionBreak");
  checkRecord(props, "props", "Document.addSectionBreak");
  checkSectionProperties(props, "props", "Document.addSectionBreak");
  const body = _body(doc);
  // Insert as the last paragraph's section properties
  if (body.length > 0) {
    const last = body[body.length - 1];
    if (last.type === "paragraph") {
      const existingProps = last.properties ?? {};
      body[body.length - 1] = {
        ...last,
        properties: { ...existingProps, sectionProperties: props }
      };
      return;
    }
  }
  // If no previous paragraph, add an empty one with section properties
  body.push(paragraph([], { sectionProperties: props }));
}

/**
 * A bookmark id for `Build.bookmarkStart` / `Build.bookmarkEnd`, above every
 * bookmark id in the document when it is first asked and every id handed out
 * since. O(1) after the first call. Bookmarks are referred to by name, so if
 * content added directly later reuses an id, the writer renumbers the clash
 * and keeps each start paired with its end.
 */
export function allocateBookmarkId(doc: DocxDocument): number {
  checkHandle(doc, "document", "doc", "Document.allocateBookmarkId");
  return _allocate(doc, "bookmark");
}

/** @deprecated Use {@link allocateBookmarkId}. Removed in 1.0. */
export const nextBookmarkId = allocateBookmarkId;

/**
 * @deprecated A document needs no build step: pass the document itself to
 * `Document.toBuffer`, `Layout.*`, `WordHtml.fromDocx` or `Pdf.fromDocx`, which apply
 * the same write-time defaults. Removed in 1.0.
 */
export function build(doc: DocxDocument): DocxDocument {
  checkHandle(doc, "document", "doc", "Document.build");
  // As before: always a new top-level object (later edits that replace a
  // top-level field on `doc` must not reach a document already built), no
  // empty collections, default page geometry, and the East Asian defaults
  // already applied.
  const out: Record<string, unknown> = { ...finalizeDocument(doc) };
  for (const [key, value] of Object.entries(out)) {
    if (
      (Array.isArray(value) && value.length === 0) ||
      (value instanceof Map && value.size === 0)
    ) {
      out[key] = undefined;
    }
  }
  out.body = doc.body;
  out.sectionProperties ??= {
    pageSize: { width: 12240, height: 15840 },
    margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
  };
  return markDerived(out as unknown as DocxDocument, doc);
}
