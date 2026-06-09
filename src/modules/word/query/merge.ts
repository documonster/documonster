/**
 * Document Merge API
 *
 * Merge multiple DocxDocument bodies into a single document.
 *
 * The first document is taken as the base (its styles, settings, page layout,
 * etc. become the foundation). Each subsequent document is appended after a
 * section break. Numbering definitions, instances and the body references
 * that point to them are rewritten when the appended document collides with
 * IDs already present in the base. Image rIds, footnote/endnote ids and
 * comment ids are also remapped, and the corresponding parts (footnotes,
 * endnotes, comments) are merged so that references emitted in the body
 * still resolve.
 *
 * What is *not* merged today (callers needing these should compose at a
 * higher level):
 *   - headers / footers (only the base document's are kept; appended
 *     documents' header/footer parts are dropped)
 *   - customXmlParts, embeddedFonts, settings, coreProperties, app
 *     metadata, theme, document protection, signatures, vbaProject
 *
 * Body content from appended documents is deep-cloned so that callers'
 * models are never mutated.
 */

import type {
  DocxDocument,
  AbstractNumbering,
  BodyContent,
  CommentDef,
  EndnoteDef,
  FootnoteDef,
  ImageDef,
  NumberingInstance,
  NumberingRef,
  Paragraph,
  ParagraphProperties,
  StructuredDocumentTag,
  Table,
  TableCell,
  TableRow
} from "../types";

// =============================================================================
// Types
// =============================================================================

/** Options for merging documents. */
export interface MergeOptions {
  /** Break type between merged documents. Default: "nextPage". */
  readonly sectionBreak?: "continuous" | "nextPage" | "evenPage" | "oddPage";
}

interface RemapState {
  /** Map of doc[i].rId → final rId (only set when the source rId collided). */
  readonly imageRIdMap: Map<string, string>;
  /** Map of doc[i].numId → final numId (only set when collided). */
  readonly numIdMap: Map<number, number>;
  /** Map of doc[i] footnote.id → final footnote.id (only set when collided). */
  readonly footnoteIdMap: Map<number, number>;
  /** Map of doc[i] endnote.id → final endnote.id (only set when collided). */
  readonly endnoteIdMap: Map<number, number>;
  /** Map of doc[i] comment.id → final comment.id (only set when collided). */
  readonly commentIdMap: Map<number, number>;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Merge multiple DocxDocument bodies into a single document.
 *
 * The first document is used as the base (preserving its styles, numbering,
 * settings, etc.). Subsequent documents' body content is appended with
 * section breaks between them. IDs that collide with the base are remapped
 * during the deep-clone of each appended body:
 *   - paragraph numbering refs (numId)
 *   - inline / floating image rIds
 *   - footnote, endnote and comment references (and their range markers)
 *
 * The corresponding `footnotes`, `endnotes`, `comments` and `images`
 * collections are merged in tandem so references emitted into the body
 * still resolve in the output package.
 *
 * @param documents - Array of documents to merge (at least 1).
 * @param options - Optional merge settings.
 * @returns A new merged DocxDocument.
 */
export function mergeDocuments(
  documents: readonly DocxDocument[],
  options?: MergeOptions
): DocxDocument {
  if (documents.length === 0) {
    return { body: [] };
  }
  if (documents.length === 1) {
    return documents[0];
  }

  const base = documents[0];
  const breakType = options?.sectionBreak ?? "nextPage";
  const mergedBody: BodyContent[] = [...base.body];

  const mergedImages: ImageDef[] = base.images ? [...base.images] : [];
  const mergedStyles = base.styles ? [...base.styles] : [];
  const existingStyleIds = new Set(mergedStyles.map(s => s.styleId));
  const mergedAbstractNums: AbstractNumbering[] = base.abstractNumberings
    ? [...base.abstractNumberings]
    : [];
  const mergedNumInstances: NumberingInstance[] = base.numberingInstances
    ? [...base.numberingInstances]
    : [];
  const mergedFootnotes: FootnoteDef[] = base.footnotes ? [...base.footnotes] : [];
  const mergedEndnotes: EndnoteDef[] = base.endnotes ? [...base.endnotes] : [];
  const mergedComments: CommentDef[] = base.comments ? [...base.comments] : [];

  for (let i = 1; i < documents.length; i++) {
    const doc = documents[i];

    // Mark a section break BEFORE appending the next document. In OOXML a
    // section break is carried by the `sectPr` of the LAST paragraph of the
    // preceding section — NOT by an extra empty paragraph. Appending an empty
    // <w:p> with only a sectPr makes Word render a stray blank line / blank
    // page. So we attach the break to the last paragraph already in the body;
    // only if the body currently ends with a non-paragraph block (e.g. a
    // table, which cannot carry a sectPr directly) do we fall back to a
    // minimal carrier paragraph.
    const lastBlock = mergedBody[mergedBody.length - 1];
    if (lastBlock && lastBlock.type === "paragraph") {
      const para = lastBlock as Paragraph;
      mergedBody[mergedBody.length - 1] = {
        ...para,
        properties: {
          ...para.properties,
          sectionProperties: { ...para.properties?.sectionProperties, breakType }
        }
      };
    } else {
      const sectionBreakPara: Paragraph = {
        type: "paragraph",
        properties: { sectionProperties: { breakType } },
        children: []
      };
      mergedBody.push(sectionBreakPara);
    }

    // Compute id remappings BEFORE cloning the body so we can rewrite refs
    // during the clone in a single pass.
    const numMaps = mergeNumberingDefinitions(doc, mergedAbstractNums, mergedNumInstances);
    const imageRIdMap = mergeImages(doc, mergedImages);
    const footnoteIdMap = mergeFootnotes(doc, mergedFootnotes);
    const endnoteIdMap = mergeEndnotes(doc, mergedEndnotes);
    const commentIdMap = mergeComments(doc, mergedComments);

    const remap: RemapState = {
      imageRIdMap,
      numIdMap: numMaps.numIdMap,
      footnoteIdMap,
      endnoteIdMap,
      commentIdMap
    };

    // Deep-clone body and rewrite refs as needed.
    const cloned = doc.body.map(b => cloneBlockWithRemap(b, remap));
    mergedBody.push(...cloned);

    // Merge styles (avoid duplicates by styleId)
    if (doc.styles) {
      for (const style of doc.styles) {
        if (!existingStyleIds.has(style.styleId)) {
          mergedStyles.push(style);
          existingStyleIds.add(style.styleId);
        }
      }
    }
  }

  return {
    ...base,
    body: mergedBody,
    images: mergedImages.length > 0 ? mergedImages : undefined,
    styles: mergedStyles.length > 0 ? mergedStyles : undefined,
    abstractNumberings: mergedAbstractNums.length > 0 ? mergedAbstractNums : undefined,
    numberingInstances: mergedNumInstances.length > 0 ? mergedNumInstances : undefined,
    footnotes: mergedFootnotes.length > 0 ? mergedFootnotes : undefined,
    endnotes: mergedEndnotes.length > 0 ? mergedEndnotes : undefined,
    comments: mergedComments.length > 0 ? mergedComments : undefined,
    // Use the final document's section properties (page layout) if available
    sectionProperties: documents[documents.length - 1].sectionProperties ?? base.sectionProperties
  };
}

// =============================================================================
// Image / footnote / endnote / comment merge helpers
// =============================================================================

/**
 * Merge `doc.images` into `mergedImages`, returning a map of original rId →
 * final rId for any image whose rId collided with one already in the base.
 *
 * Images are deduplicated only when both rId AND fileName match — different
 * documents may legitimately use the same fileName for different binaries
 * (`image1.png` is overwhelmingly common), so deduping by fileName alone
 * would silently drop content. Conversely, two ImageDefs that share both
 * rId and fileName almost certainly came from the same source document
 * round-tripped through the same writer, so coalescing is safe.
 */
function mergeImages(doc: DocxDocument, mergedImages: ImageDef[]): Map<string, string> {
  const rIdMap = new Map<string, string>();
  if (!doc.images || doc.images.length === 0) {
    return rIdMap;
  }
  const usedRIds = new Set(mergedImages.map(m => m.rId).filter((x): x is string => !!x));
  const usedFileNames = new Set(mergedImages.map(m => m.fileName));
  // Caches for O(1) collision checks. The previous implementation called
  // `mergedImages.includes(...)` and `mergedImages.find(...)` per image,
  // which is O(N²) overall — unacceptable for documents with thousands of
  // images.
  const seenObjects = new WeakSet<object>();
  for (const m of mergedImages) {
    seenObjects.add(m);
  }
  const exactByRId = new Map<string, Set<string>>(); // rId → fileNames already merged with that rId
  for (const m of mergedImages) {
    if (m.rId) {
      let names = exactByRId.get(m.rId);
      if (!names) {
        names = new Set();
        exactByRId.set(m.rId, names);
      }
      names.add(m.fileName);
    }
  }

  let nextRIdSuffix = 1;

  const allocFreshRId = (oldRId: string): string => {
    let candidate = `${oldRId}_m${nextRIdSuffix++}`;
    while (usedRIds.has(candidate)) {
      candidate = `${oldRId}_m${nextRIdSuffix++}`;
    }
    return candidate;
  };
  const allocFreshFileName = (oldName: string): string => {
    if (!usedFileNames.has(oldName)) {
      return oldName;
    }
    const dot = oldName.lastIndexOf(".");
    const stem = dot >= 0 ? oldName.slice(0, dot) : oldName;
    const ext = dot >= 0 ? oldName.slice(dot) : "";
    let n = 2;
    let candidate = `${stem}_${n}${ext}`;
    while (usedFileNames.has(candidate)) {
      n++;
      candidate = `${stem}_${n}${ext}`;
    }
    return candidate;
  };

  const recordPush = (img: ImageDef): void => {
    seenObjects.add(img);
    if (img.rId) {
      let names = exactByRId.get(img.rId);
      if (!names) {
        names = new Set();
        exactByRId.set(img.rId, names);
      }
      names.add(img.fileName);
    }
  };

  for (const img of doc.images) {
    // Coalesce identical object references in O(1).
    if (seenObjects.has(img)) {
      continue;
    }

    if (!img.rId) {
      // No rId to remap; rename the file conservatively if a different
      // image already claims the name. We can't dedupe by fileName alone
      // because two real documents can both ship `image1.png` with
      // unrelated bytes.
      const fileName = allocFreshFileName(img.fileName);
      const next = fileName === img.fileName ? img : { ...img, fileName };
      mergedImages.push(next);
      usedFileNames.add(fileName);
      recordPush(next);
      continue;
    }

    // Coalesce identical (rId, fileName) entries via O(1) lookup.
    const namesForRId = exactByRId.get(img.rId);
    if (namesForRId && namesForRId.has(img.fileName)) {
      continue;
    }

    const collidesRId = usedRIds.has(img.rId);
    if (!collidesRId) {
      const fileName = allocFreshFileName(img.fileName);
      const next = fileName === img.fileName ? img : { ...img, fileName };
      mergedImages.push(next);
      usedRIds.add(img.rId);
      usedFileNames.add(fileName);
      recordPush(next);
      continue;
    }

    // Both rId is taken AND it points at a different image. Remap.
    const newRId = allocFreshRId(img.rId);
    const fileName = allocFreshFileName(img.fileName);
    rIdMap.set(img.rId, newRId);
    const next: ImageDef = {
      ...img,
      rId: newRId,
      ...(fileName !== img.fileName ? { fileName } : {})
    } as ImageDef;
    mergedImages.push(next);
    usedRIds.add(newRId);
    usedFileNames.add(fileName);
    recordPush(next);
  }
  return rIdMap;
}

/**
 * Iterative max over a Set/iterable. `Math.max(...set)` triggers a stack
 * overflow on JS engines once the spread argument count exceeds ~10⁵
 * (V8: ~125k); a hostile document with 10⁶ collisionful note ids can
 * easily reach that limit.
 */
function maxOf(values: Iterable<number>): number {
  let m = -Infinity;
  for (const v of values) {
    if (v > m) {
      m = v;
    }
  }
  return m;
}

function mergeFootnotes(doc: DocxDocument, merged: FootnoteDef[]): Map<number, number> {
  const idMap = new Map<number, number>();
  if (!doc.footnotes || doc.footnotes.length === 0) {
    return idMap;
  }
  const used = new Set(merged.map(n => n.id));
  let nextId = used.size > 0 ? maxOf(used) + 1 : 1;
  for (const note of doc.footnotes) {
    // ids -1 and 0 are reserved for separators and exist only once per
    // document — assume the base already has them and skip duplicates.
    if (note.id <= 0) {
      if (!used.has(note.id)) {
        merged.push(note);
        used.add(note.id);
      }
      continue;
    }
    let id = note.id;
    if (used.has(id)) {
      id = nextId++;
      idMap.set(note.id, id);
    } else if (id >= nextId) {
      nextId = id + 1;
    }
    used.add(id);
    merged.push(id === note.id ? note : { ...note, id });
  }
  return idMap;
}

function mergeEndnotes(doc: DocxDocument, merged: EndnoteDef[]): Map<number, number> {
  const idMap = new Map<number, number>();
  if (!doc.endnotes || doc.endnotes.length === 0) {
    return idMap;
  }
  const used = new Set(merged.map(n => n.id));
  let nextId = used.size > 0 ? maxOf(used) + 1 : 1;
  for (const note of doc.endnotes) {
    if (note.id <= 0) {
      if (!used.has(note.id)) {
        merged.push(note);
        used.add(note.id);
      }
      continue;
    }
    let id = note.id;
    if (used.has(id)) {
      id = nextId++;
      idMap.set(note.id, id);
    } else if (id >= nextId) {
      nextId = id + 1;
    }
    used.add(id);
    merged.push(id === note.id ? note : { ...note, id });
  }
  return idMap;
}

function mergeComments(doc: DocxDocument, merged: CommentDef[]): Map<number, number> {
  const idMap = new Map<number, number>();
  if (!doc.comments || doc.comments.length === 0) {
    return idMap;
  }
  const used = new Set(merged.map(c => c.id));
  let nextId = used.size > 0 ? maxOf(used) + 1 : 0;
  for (const cmt of doc.comments) {
    let id = cmt.id;
    if (used.has(id)) {
      id = nextId++;
      idMap.set(cmt.id, id);
    } else if (id >= nextId) {
      nextId = id + 1;
    }
    used.add(id);
    merged.push(id === cmt.id ? cmt : { ...cmt, id });
  }
  return idMap;
}

// =============================================================================
// Numbering merge
// =============================================================================

/**
 * Append `doc`'s numbering definitions/instances into the running merge state,
 * remapping abstractNumIds and numIds that collide with what is already there.
 * The returned maps describe the remapping that callers must apply to body refs.
 */
function mergeNumberingDefinitions(
  doc: DocxDocument,
  mergedAbstractNums: AbstractNumbering[],
  mergedNumInstances: NumberingInstance[]
): { absIdMap: Map<number, number>; numIdMap: Map<number, number> } {
  const absIdMap = new Map<number, number>();
  const numIdMap = new Map<number, number>();

  if (doc.abstractNumberings && doc.abstractNumberings.length > 0) {
    const existingAbsIds = new Set(mergedAbstractNums.map(a => a.abstractNumId));
    let nextAbsId = existingAbsIds.size > 0 ? maxOf(existingAbsIds) + 1 : 0;
    for (const abs of doc.abstractNumberings) {
      if (existingAbsIds.has(abs.abstractNumId)) {
        const newId = nextAbsId++;
        absIdMap.set(abs.abstractNumId, newId);
        existingAbsIds.add(newId);
        mergedAbstractNums.push({ ...abs, abstractNumId: newId });
      } else {
        existingAbsIds.add(abs.abstractNumId);
        mergedAbstractNums.push(abs);
      }
    }
  }

  if (doc.numberingInstances && doc.numberingInstances.length > 0) {
    const existingNumIds = new Set(mergedNumInstances.map(n => n.numId));
    let nextNumId = existingNumIds.size > 0 ? maxOf(existingNumIds) + 1 : 1;
    for (const inst of doc.numberingInstances) {
      const remappedAbsId = absIdMap.get(inst.abstractNumId) ?? inst.abstractNumId;
      let numId = inst.numId;
      if (existingNumIds.has(numId)) {
        numId = nextNumId++;
        numIdMap.set(inst.numId, numId);
      }
      existingNumIds.add(numId);
      mergedNumInstances.push({ ...inst, numId, abstractNumId: remappedAbsId });
    }
  }

  return { absIdMap, numIdMap };
}

// =============================================================================
// Body cloning with remap
// =============================================================================

/**
 * Deep-clone a body content block while rewriting any paragraph numbering
 * references, image rIds, and footnote/endnote/comment references through
 * the supplied remap state. The original block is never mutated.
 */
function cloneBlockWithRemap(block: BodyContent, remap: RemapState): BodyContent {
  switch (block.type) {
    case "paragraph":
      return cloneParagraph(block, remap);
    case "table":
      return cloneTable(block, remap);
    case "sdt":
      return cloneSdt(block as StructuredDocumentTag, remap);
    case "floatingImage": {
      const cloned = structuredClone(block) as { rId?: string };
      if (cloned.rId && remap.imageRIdMap.has(cloned.rId)) {
        cloned.rId = remap.imageRIdMap.get(cloned.rId)!;
      }
      return cloned as BodyContent;
    }
    case "textBox": {
      // textBox.content is `Paragraph[]`; clone each paragraph through the
      // full remap so footnoteRef/endnoteRef/comment/numId/image rIds
      // inside are all rewritten.
      const tb = block as { content?: readonly Paragraph[] } & BodyContent;
      const cloned = structuredClone(block) as { content?: Paragraph[] } & BodyContent;
      if (tb.content && tb.content.length > 0) {
        cloned.content = tb.content.map(p => cloneParagraph(p, remap));
      }
      return cloned as BodyContent;
    }
    case "drawingShape": {
      const shape = block as { textContent?: readonly Paragraph[] } & BodyContent;
      const cloned = structuredClone(block) as { textContent?: Paragraph[] } & BodyContent;
      if (shape.textContent && shape.textContent.length > 0) {
        cloned.textContent = shape.textContent.map(p => cloneParagraph(p, remap));
      }
      return cloned as BodyContent;
    }
    case "tableOfContents": {
      const toc = block as { cachedParagraphs?: readonly Paragraph[] } & BodyContent;
      const cloned = structuredClone(block) as { cachedParagraphs?: Paragraph[] } & BodyContent;
      if (toc.cachedParagraphs && toc.cachedParagraphs.length > 0) {
        cloned.cachedParagraphs = toc.cachedParagraphs.map(p => cloneParagraph(p, remap));
      }
      return cloned as BodyContent;
    }
    default:
      // For chart / altChunk / opaqueDrawing / checkBox / math etc.
      // there's no nested paragraph content the merge cares about; only
      // the top-level rId (when present) needs remapping. structuredClone
      // gives an independent copy, then walk for image rIds carried by
      // image-bearing nodes.
      return remapImageRIdsDeep(structuredClone(block), remap.imageRIdMap) as BodyContent;
  }
}

function cloneParagraph(para: Paragraph, remap: RemapState): Paragraph {
  // Walk and clone children in one pass so we can rewrite refs inline.
  const clonedChildren = structuredClone(para.children);
  for (const child of clonedChildren) {
    remapParagraphChild(child, remap);
  }
  const clonedProps = para.properties
    ? cloneParagraphProps(para.properties, remap.numIdMap)
    : undefined;
  const out: Paragraph = {
    ...para,
    properties: clonedProps,
    children: clonedChildren
  };
  return out;
}

function remapParagraphChild(child: unknown, remap: RemapState): void {
  if (!child || typeof child !== "object") {
    return;
  }
  const c = child as Record<string, unknown>;
  if ("type" in c) {
    switch (c.type) {
      case "hyperlink":
        if (Array.isArray(c.children)) {
          for (const sub of c.children as unknown[]) {
            remapParagraphChild(sub, remap);
          }
        }
        return;
      case "insertedRun":
      case "movedToRun":
        remapParagraphChild(c.run, remap);
        return;
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference": {
        const idVal = c.id;
        if (typeof idVal === "number" && remap.commentIdMap.has(idVal)) {
          c.id = remap.commentIdMap.get(idVal);
        }
        return;
      }
    }
  }
  // Plain Run with content.
  if (Array.isArray(c.content)) {
    for (const rc of c.content as unknown[]) {
      remapRunContent(rc, remap);
    }
  }
}

function remapRunContent(rc: unknown, remap: RemapState): void {
  if (!rc || typeof rc !== "object") {
    return;
  }
  const r = rc as Record<string, unknown>;
  switch (r.type) {
    case "image": {
      const rId = r.rId;
      if (typeof rId === "string" && remap.imageRIdMap.has(rId)) {
        r.rId = remap.imageRIdMap.get(rId);
      }
      return;
    }
    case "footnoteRef": {
      const id = r.id;
      if (typeof id === "number" && remap.footnoteIdMap.has(id)) {
        r.id = remap.footnoteIdMap.get(id);
      }
      return;
    }
    case "endnoteRef": {
      const id = r.id;
      if (typeof id === "number" && remap.endnoteIdMap.has(id)) {
        r.id = remap.endnoteIdMap.get(id);
      }
      return;
    }
  }
}

/**
 * Walk an opaque cloned subtree and rewrite `rId` fields that point at an
 * image we just remapped. Used for block types we don't model in detail
 * (textBox, drawingShape, chart container, …) so a body image reference
 * moved across documents still resolves.
 *
 * **Critical**: this is restricted to nodes whose `type` field marks them
 * as image carriers (`image`, `floatingImage`, `drawingShape` carrying a
 * blip rId, etc.). The previous implementation rewrote any string `rId`
 * field whose value matched an entry in `imageRIdMap`, which silently
 * broke chart / hyperlink / customXml relationships whenever the appended
 * document happened to assign an image and a chart the same rId number
 * (very common because each rels space numbers from rId1).
 */
const IMAGE_CARRIER_TYPES = new Set([
  "image",
  "floatingImage",
  // drawingShape can carry a fill image rId on the top-level node — the
  // shape model exposes it as `rId` on the fill descriptor, not the shape
  // itself, so we whitelist it explicitly when the type matches.
  "drawingShape",
  "opaqueDrawing"
]);

function remapImageRIdsDeep(node: unknown, imageRIdMap: Map<string, string>): unknown {
  if (imageRIdMap.size === 0 || !node || typeof node !== "object") {
    return node;
  }
  const stack: unknown[] = [node];
  while (stack.length > 0) {
    const cur = stack.pop();
    if (!cur || typeof cur !== "object") {
      continue;
    }
    if (Array.isArray(cur)) {
      for (const item of cur) {
        stack.push(item);
      }
      continue;
    }
    const obj = cur as Record<string, unknown>;
    const t = typeof obj.type === "string" ? (obj.type as string) : undefined;
    if (
      t !== undefined &&
      IMAGE_CARRIER_TYPES.has(t) &&
      typeof obj.rId === "string" &&
      imageRIdMap.has(obj.rId)
    ) {
      obj.rId = imageRIdMap.get(obj.rId);
    }
    for (const key of Object.keys(obj)) {
      const v = obj[key];
      if (v && typeof v === "object") {
        stack.push(v);
      }
    }
  }
  return node;
}

function cloneParagraphProps(
  props: ParagraphProperties,
  numIdMap: Map<number, number>
): ParagraphProperties {
  const cloned = structuredClone(props) as ParagraphProperties & { numbering?: NumberingRef };
  if (cloned.numbering && numIdMap.has(cloned.numbering.numId)) {
    (cloned.numbering as { numId: number; level: number }).numId = numIdMap.get(
      cloned.numbering.numId
    )!;
  }
  return cloned;
}

function cloneTable(table: Table, remap: RemapState): Table {
  return {
    ...table,
    rows: table.rows.map(r => cloneRow(r, remap))
  } as Table;
}

function cloneRow(row: TableRow, remap: RemapState): TableRow {
  return {
    ...row,
    cells: row.cells.map(c => cloneCell(c, remap))
  } as TableRow;
}

function cloneCell(cell: TableCell, remap: RemapState): TableCell {
  return {
    ...cell,
    content: cell.content.map(item => cloneBlockWithRemap(item as BodyContent, remap))
  } as TableCell;
}

function cloneSdt(sdt: StructuredDocumentTag, remap: RemapState): StructuredDocumentTag {
  return {
    ...sdt,
    content: sdt.content.map(c => {
      if ("type" in c && (c.type === "paragraph" || c.type === "table")) {
        return cloneBlockWithRemap(c as BodyContent, remap) as Paragraph | Table;
      }
      // Run / other in-line children inside an inline SDT: clone and
      // remap refs in place.
      const cloned = structuredClone(c);
      remapParagraphChild(cloned, remap);
      return cloned;
    })
  } as StructuredDocumentTag;
}
