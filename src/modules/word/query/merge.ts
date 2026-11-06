/**
 * Document Merge API
 *
 * Merge multiple DocxDocument bodies into a single document.
 *
 * The first document is taken as the base (its styles, settings, page layout,
 * etc. become the foundation). Each subsequent document is appended after a
 * section break. Numbering definitions (abstract definitions, instances and
 * picture bullets) and the references that point to them are rewritten when
 * the appended document collides with ids already present. Image keys,
 * footnote/endnote ids and comment ids are also remapped, and the
 * corresponding parts are merged so every reference still resolves.
 *
 * Every section keeps the headers and footers of the document it came from:
 * an appended document's header and footer parts are merged in under keys of
 * their own, and its sections reference them explicitly — or reference an
 * empty part, so a document without a header never shows the previous one's.
 *
 * What is *not* merged today (callers needing these should compose at a
 * higher level):
 *   - customXmlParts, embeddedFonts, settings (including
 *     `evenAndOddHeaders`, which is the base's for the whole result),
 *     coreProperties, app metadata, theme, watermark, document protection,
 *     signatures, vbaProject
 *
 * Bookmark names are unique across a document, so a name an appended
 * document shares with an earlier one is renamed (`_Toc1` → `_Toc1_2`),
 * together with every `REF`/`PAGEREF`/`NOTEREF`/`GOTOBUTTON` field, `HYPERLINK
 * \l` switch and internal hyperlink in that document that refers to it.
 *
 * Body content from appended documents is deep-cloned so that callers'
 * models are never mutated.
 */

import type { SchemaMatches } from "@utils/schema";
import { xmlDecode, xmlEncodeAttr } from "@utils/xml-encode";
import { checkArray, checkRecord } from "@word/core/argument-checks";
import {
  imageKeys,
  isSamePicture,
  rewriteImageReferences,
  rewriteMarkupReferences
} from "@word/core/image-keys";
import type { Mutable } from "@word/core/internal-utils";
import { checkSchema, w } from "@word/core/option-schema";
import { markDerived } from "@word/core/ownership";
import { hostRangeMarkers } from "@word/core/range-markers";
import type {
  DocxDocument,
  AbstractNumbering,
  BodyContent,
  CommentDef,
  DrawingShape,
  EndnoteDef,
  FooterDef,
  FootnoteDef,
  HeaderDef,
  HeaderFooterRef,
  ImageDef,
  NumPicBullet,
  NumberingInstance,
  NumberingRef,
  Paragraph,
  ParagraphProperties,
  RangeMarker,
  SectionProperties,
  StructuredDocumentTag,
  Table,
  TableCell,
  TableRow
} from "@word/types";

// =============================================================================
// Types
// =============================================================================

/** Options for merging documents. */
export interface MergeOptions {
  /** Break type between merged documents. Default: "nextPage". */
  readonly sectionBreak?: "continuous" | "nextPage" | "evenPage" | "oddPage";
}

const mergeOptionsSchema = w.record({
  sectionBreak: w.oneOf(["continuous", "nextPage", "evenPage", "oddPage"]).optional()
});
const _mergeOptionsMatch: SchemaMatches<typeof mergeOptionsSchema, MergeOptions> = true;
void _mergeOptionsMatch;

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
  /**
   * Bookmark and permission ids of ranges in doc[i]'s markup kept verbatim,
   * where another document's verbatim markup already uses them. Modelled
   * ranges need no table: the writer renumbers them around verbatim ones.
   */
  readonly markupBookmarkIdMap?: Map<number, number>;
  readonly markupPermIdMap?: Map<number, number>;
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
 *   - paragraph numbering refs (numId), and picture-bullet ids in the
 *     numbering definitions
 *   - image keys, wherever they are used — inline and floating images,
 *     picture fills and other markup kept verbatim, picture bullets
 *   - footnote, endnote and comment references (and their range markers)
 *
 * The corresponding `footnotes`, `endnotes`, `comments`, `images` and
 * `numPicBullets` collections are merged in tandem so references emitted into
 * the body still resolve in the output package. Each section keeps the
 * headers and footers of the document it came from.
 *
 * Does not modify `documents`; the result may share unchanged nodes with them;
 * `Document.*` copies what it edits, so an edit made through one never reaches
 * the other.
 *
 * @param documents - Array of documents to merge (at least 1).
 * @param options - Optional merge settings.
 * @returns A new merged DocxDocument.
 */
export function mergeDocuments(
  documents: readonly DocxDocument[],
  options?: MergeOptions
): DocxDocument {
  const where = "Transform.merge";
  checkArray(documents, "documents", where);
  documents.forEach((d, i) => checkRecord(d, `documents[${i}]`, where));
  checkSchema(mergeOptionsSchema.optional(), options, where);
  if (documents.length === 0) {
    return { body: [] };
  }
  if (documents.length === 1) {
    return documents[0];
  }

  const base = documents[0];
  const breakType = options?.sectionBreak ?? "nextPage";
  /** The final section of the content merged so far, still open. */
  let previousFinal: SectionProperties | undefined = base.sectionProperties;
  const mergedBody: BodyContent[] = [...base.body];
  /**
   * Range markers after the last block of the content merged so far. Appended
   * content follows them, so they move to the paragraph closing the section.
   */
  let trailing: readonly RangeMarker[] | undefined = base.bodyRangeMarkersAfter;

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
  const mergedPicBullets: NumPicBullet[] = base.numPicBullets ? [...base.numPicBullets] : [];

  // Headers and footers: every section is given explicit references, so each
  // shows what it showed in the document it came from (see `SectionParts`).
  const parts = documents.some(d => (d.headers?.size ?? 0) > 0 || (d.footers?.size ?? 0) > 0)
    ? createSectionParts(base)
    : undefined;
  if (parts) {
    for (let b = 0; b < mergedBody.length; b++) {
      const block = mergedBody[b]!;
      if (block.type === "paragraph" && block.properties?.sectionProperties) {
        mergedBody[b] = withSection(
          block,
          parts.section(base, block.properties.sectionProperties, false, false, IDENTITY)
        );
      }
    }
    previousFinal = parts.section(base, base.sectionProperties ?? {}, true, false, IDENTITY);
  }

  const markupRanges = scanMarkupRangeIds(base);

  // Bookmark names: those the merged content has claimed so far, and every
  // name any input uses, which a renamed bookmark must not take either.
  const claimedNames = collectBookmarkNames([
    base.body,
    base.bodyRangeMarkersAfter,
    base.headers,
    base.footers,
    base.footnotes,
    base.endnotes,
    base.comments
  ]);
  const unavailableNames = new Set(claimedNames);
  const namesByDocument = documents.map((doc, i) =>
    i === 0 ? claimedNames : collectBookmarkNames(mergedParts(doc))
  );
  for (const names of namesByDocument) {
    for (const name of names) {
      unavailableNames.add(name);
    }
  }

  for (let i = 1; i < documents.length; i++) {
    const renames = new Map<string, string>();
    for (const name of namesByDocument[i]!) {
      if (claimedNames.has(name)) {
        const renamed = freshBookmarkName(name, unavailableNames);
        unavailableNames.add(renamed);
        renames.set(name, renamed);
      }
    }
    for (const name of namesByDocument[i]!) {
      claimedNames.add(renames.get(name) ?? name);
    }
    const source = documents[i]!;
    const doc: DocxDocument =
      renames.size === 0
        ? source
        : {
            ...source,
            footnotes: renameBookmarks(source.footnotes, renames),
            endnotes: renameBookmarks(source.endnotes, renames),
            comments: renameBookmarks(source.comments, renames)
          };

    // Close the section in progress with the page setup it actually has —
    // the previous document's final section. In OOXML a section's properties
    // ride on the `sectPr` of its LAST paragraph, so they go onto the body's
    // last paragraph; a body that ends with a table, or whose last paragraph
    // already closes an earlier section, gets a minimal carrier paragraph
    // (an empty <w:p> renders as a blank line, so it is the fallback only).
    const closing: SectionProperties = previousFinal ?? {};
    const lastBlock = mergedBody[mergedBody.length - 1];
    if (lastBlock?.type === "paragraph" && !lastBlock.properties?.sectionProperties) {
      const host = (trailing && hostRangeMarkers(lastBlock, trailing, true)) || lastBlock;
      mergedBody[mergedBody.length - 1] = {
        ...host,
        properties: { ...lastBlock.properties, sectionProperties: closing }
      };
    } else {
      mergedBody.push({
        type: "paragraph",
        properties: { sectionProperties: closing },
        children: [],
        ...(trailing ? { rangeMarkersBefore: trailing } : {})
      });
    }

    // Compute id remappings BEFORE cloning the body so we can rewrite refs
    // during the clone in a single pass.
    const imageRIdMap = mergeImages(doc, mergedImages);
    const picBulletIdMap = mergePicBullets(doc, mergedPicBullets, imageRIdMap);
    const numMaps = mergeNumberingDefinitions(
      doc,
      mergedAbstractNums,
      mergedNumInstances,
      picBulletIdMap
    );
    // Notes and comments show pictures too, named by the same keys.
    const notesFrom = mergedFootnotes.length;
    const endnotesFrom = mergedEndnotes.length;
    const commentsFrom = mergedComments.length;
    const footnoteIdMap = mergeFootnotes(doc, mergedFootnotes);
    const endnoteIdMap = mergeEndnotes(doc, mergedEndnotes);
    const commentIdMap = mergeComments(doc, mergedComments);
    const ranges = scanMarkupRangeIds(doc);
    const remap: RemapState = {
      imageRIdMap,
      numIdMap: numMaps.numIdMap,
      footnoteIdMap,
      endnoteIdMap,
      commentIdMap,
      markupBookmarkIdMap: claimIds(markupRanges.bookmark, ranges.bookmark),
      markupPermIdMap: claimIds(markupRanges.perm, ranges.perm)
    };
    /** Image keys, and every id markup kept verbatim refers to, through the same tables. */
    const rewriteAll = <T>(node: T): T =>
      rewriteMarkupReferences(node, imageRIdMap, xml =>
        renameInRawXml(remapMarkupIds(xml, remap), renames)
      );
    rewriteFrom(mergedFootnotes, notesFrom, rewriteAll);
    rewriteFrom(mergedEndnotes, endnotesFrom, rewriteAll);
    rewriteFrom(mergedComments, commentsFrom, rewriteAll);

    // Deep-clone body and rewrite refs as needed.
    const cloned = doc.body.map(b =>
      rewriteAll(renameBookmarks(cloneBlockWithRemap(b, remap), renames))
    );
    trailing = renameBookmarks(
      doc.bodyRangeMarkersAfter?.map(m => cloneMarker(m, remap)),
      renames
    );

    // The document's headers and footers join the merged set under keys of
    // their own, and its sections are pointed at them.
    let finalSection = doc.sectionProperties;
    if (parts) {
      const keys = parts.adopt(doc, part =>
        part.map(block => rewriteAll(renameBookmarks(cloneBlockWithRemap(block, remap), renames)))
      );
      let first = true;
      for (let b = 0; b < cloned.length; b++) {
        const block = cloned[b]!;
        if (block.type === "paragraph" && block.properties?.sectionProperties) {
          cloned[b] = withSection(
            block,
            parts.section(doc, block.properties.sectionProperties, false, first, keys)
          );
          first = false;
        }
      }
      finalSection = parts.section(doc, doc.sectionProperties ?? {}, true, first, keys);
    }

    // The break goes on the section it *opens*: a section's `w:type` says how
    // that section starts (ECMA-376 §17.6.22), which Word's output confirms.
    // That is the appended document's first section — its first paragraph
    // `sectPr`, or its final section when it has no inner breaks.
    const firstInner = cloned.findIndex(
      b => b.type === "paragraph" && b.properties?.sectionProperties !== undefined
    );
    if (firstInner >= 0) {
      const para = cloned[firstInner] as Paragraph;
      cloned[firstInner] = {
        ...para,
        properties: {
          ...para.properties,
          sectionProperties: { ...para.properties!.sectionProperties, breakType }
        }
      };
      previousFinal = finalSection;
    } else {
      previousFinal = { ...finalSection, breakType };
    }
    for (const block of cloned) {
      mergedBody.push(block);
    }

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

  return markDerived(
    {
      ...base,
      body: mergedBody,
      bodyRangeMarkersAfter: trailing,
      images: mergedImages.length > 0 ? mergedImages : undefined,
      styles: mergedStyles.length > 0 ? mergedStyles : undefined,
      abstractNumberings: mergedAbstractNums.length > 0 ? mergedAbstractNums : undefined,
      numberingInstances: mergedNumInstances.length > 0 ? mergedNumInstances : undefined,
      footnotes: mergedFootnotes.length > 0 ? mergedFootnotes : undefined,
      endnotes: mergedEndnotes.length > 0 ? mergedEndnotes : undefined,
      comments: mergedComments.length > 0 ? mergedComments : undefined,
      numPicBullets: mergedPicBullets.length > 0 ? mergedPicBullets : undefined,
      ...(parts
        ? {
            headers: parts.headers.size > 0 ? parts.headers : undefined,
            footers: parts.footers.size > 0 ? parts.footers : undefined
          }
        : {}),
      // The last document's final section, opened by the requested break.
      sectionProperties: previousFinal
    },
    ...documents
  );
}

// =============================================================================
// Headers and footers
// =============================================================================

const HEADER_FOOTER_TYPES = ["default", "first", "even"] as const;
type PartKind = "headers" | "footers";
const PART_KINDS: readonly PartKind[] = ["headers", "footers"];
const IDENTITY: ReadonlyMap<string, string> = new Map();

/**
 * The merged documents' headers and footers, and the section references
 * that select them.
 *
 * A section shows the header of each type it references and otherwise
 * inherits its predecessor's (ECMA-376 §17.10.5). So an appended document's
 * first section would show the previous document's header wherever it named
 * none of its own; it is given an explicit reference to an empty part
 * instead. References a document leaves for the writer to resolve — by type,
 * or synthesised for a final section that names none — are resolved here the
 * way the writer would have, against that document's own parts, before its
 * parts are renamed.
 */
interface SectionParts {
  readonly headers: Map<string, HeaderDef>;
  readonly footers: Map<string, FooterDef>;
  /**
   * Add `doc`'s headers and footers, their content passed through `clone`,
   * under keys free in the merged set. Returns old key → new key.
   */
  adopt(
    doc: DocxDocument,
    clone: (blocks: readonly (Paragraph | Table)[]) => BodyContent[]
  ): ReadonlyMap<string, string>;
  /**
   * `sect` of `doc` with explicit, renamed references. `isFinal` marks the
   * document's final section; `opensDocument` its first.
   */
  section(
    doc: DocxDocument,
    sect: SectionProperties,
    isFinal: boolean,
    opensDocument: boolean,
    keys: ReadonlyMap<string, string>
  ): SectionProperties;
}

function createSectionParts(base: DocxDocument): SectionParts {
  const maps = {
    headers: new Map<string, HeaderDef>(base.headers ?? []),
    footers: new Map<string, FooterDef>(base.footers ?? [])
  };
  /** Per kind, the part each type currently shows, by reference or inheritance. */
  const current = { headers: new Map<string, string>(), footers: new Map<string, string>() };
  const emptyKeys: Partial<Record<PartKind, string>> = {};

  const freeKey = (kind: PartKind, wanted: string): string => {
    let key = wanted;
    for (let n = 1; maps[kind].has(key); n++) {
      key = `${wanted}_m${n}`;
    }
    return key;
  };
  const emptyPart = (kind: PartKind): string => {
    let key = emptyKeys[kind];
    if (key === undefined) {
      key = freeKey(kind, "mergeEmpty");
      maps[kind].set(key, { content: { children: [{ type: "paragraph", children: [] }] } });
      emptyKeys[kind] = key;
    }
    return key;
  };

  /** `doc`'s references of one kind, resolved as the writer resolves them. */
  const resolve = (
    doc: DocxDocument,
    kind: PartKind,
    refs: readonly HeaderFooterRef[] | undefined,
    isFinal: boolean
  ): HeaderFooterRef[] => {
    const own = doc[kind];
    if (!own || own.size === 0) {
      return [];
    }
    if (!refs || refs.length === 0) {
      return isFinal
        ? HEADER_FOOTER_TYPES.filter(t => own.has(t)).map(t => ({ type: t, rId: t }))
        : [];
    }
    return refs.flatMap(ref =>
      own.has(ref.rId) ? [ref] : own.has(ref.type) ? [{ ...ref, rId: ref.type }] : []
    );
  };

  return {
    headers: maps.headers,
    footers: maps.footers,

    adopt(doc, clone) {
      const keys = new Map<string, string>();
      for (const kind of PART_KINDS) {
        for (const [key, def] of doc[kind] ?? []) {
          const next = freeKey(kind, key);
          // The part's own rId is the writer's to assign; a stale one could
          // alias a base part's key.
          const { rId: _rId, ...rest } = def;
          maps[kind].set(next, {
            ...rest,
            content: {
              ...def.content,
              children: clone(def.content.children) as (Paragraph | Table)[]
            }
          });
          if (next !== key) {
            keys.set(`${kind}:${key}`, next);
          }
        }
      }
      return keys;
    },

    section(doc, sect, isFinal, opensDocument, keys) {
      let next = sect;
      for (const kind of PART_KINDS) {
        const refs = resolve(doc, kind, sect[kind], isFinal).map(ref => {
          const key = keys.get(`${kind}:${ref.rId}`);
          return key ? { ...ref, rId: key } : ref;
        });
        const shownBefore = current[kind];
        if (opensDocument) {
          // Nothing of the previous document may show through.
          for (const type of shownBefore.keys()) {
            if (!refs.some(r => r.type === type)) {
              refs.push({ type: type as HeaderFooterRef["type"], rId: emptyPart(kind) });
            }
          }
        }
        if (isFinal && refs.length === 0 && shownBefore.size > 0) {
          // The writer gives a final section without references the parts
          // keyed by type; state the inherited ones instead.
          for (const [type, key] of shownBefore) {
            refs.push({ type: type as HeaderFooterRef["type"], rId: key });
          }
        }
        for (const ref of refs) {
          shownBefore.set(ref.type, ref.rId);
        }
        if (refs.length > 0 || sect[kind] !== undefined) {
          next = { ...next, [kind]: refs };
        }
      }
      return next;
    }
  };
}

function withSection(para: Paragraph, sectionProperties: SectionProperties): Paragraph {
  return sectionProperties === para.properties?.sectionProperties
    ? para
    : { ...para, properties: { ...para.properties, sectionProperties } };
}

// =============================================================================
// Picture bullets
// =============================================================================

/**
 * Append `doc`'s picture bullets, renumbering an id the merged set already
 * uses and renaming the image each shows. Returns old id → new id.
 */
function mergePicBullets(
  doc: DocxDocument,
  merged: NumPicBullet[],
  imageKeyMap: ReadonlyMap<string, string>
): Map<number, number> {
  const idMap = new Map<number, number>();
  if (!doc.numPicBullets || doc.numPicBullets.length === 0) {
    return idMap;
  }
  const used = new Set(merged.map(b => b.id));
  let nextId = used.size > 0 ? maxOf(used) + 1 : 0;
  for (const bullet of doc.numPicBullets) {
    let id = bullet.id;
    if (used.has(id)) {
      id = nextId++;
      idMap.set(bullet.id, id);
    } else if (id >= nextId) {
      nextId = id + 1;
    }
    used.add(id);
    const renamed = bullet.rId !== undefined ? imageKeyMap.get(bullet.rId) : undefined;
    merged.push(
      rewriteImageReferences(
        { ...bullet, id, ...(renamed !== undefined ? { rId: renamed } : {}) },
        imageKeyMap
      )
    );
  }
  return idMap;
}

// =============================================================================
// Bookmark names
// =============================================================================

/** The parts of an appended document the merge keeps. */
function mergedParts(doc: DocxDocument): unknown[] {
  return [
    doc.body,
    doc.bodyRangeMarkersAfter,
    doc.headers,
    doc.footers,
    doc.footnotes,
    doc.endnotes,
    doc.comments
  ];
}

/** Word's limit on the length of a bookmark name. */
const MAX_BOOKMARK_NAME = 40;

const RAW_BOOKMARK_NAME =
  /(<(?:[\w.-]+:)?bookmarkStart\b[^>]*?\s(?:[\w.-]+:)?name\s*=\s*)(["'])(.*?)\2/g;
const RAW_HYPERLINK_ANCHOR =
  /(<(?:[\w.-]+:)?hyperlink\b[^>]*?\s(?:[\w.-]+:)?anchor\s*=\s*)(["'])(.*?)\2/g;

/**
 * Every bookmark name under `roots`, model and markup kept verbatim (any field
 * whose name ends in `Xml`, as {@link rewriteMarkupReferences} reads them) alike.
 */
function collectBookmarkNames(roots: readonly unknown[]): Set<string> {
  const names = new Set<string>();
  const stack: unknown[] = [...roots];
  while (stack.length > 0) {
    const node = stack.pop();
    if (!node || typeof node !== "object" || ArrayBuffer.isView(node)) {
      continue;
    }
    if (node instanceof Map) {
      stack.push(...node.values());
      continue;
    }
    if (Array.isArray(node)) {
      stack.push(...node);
      continue;
    }
    const o = node as Record<string, unknown>;
    if (o.type === "bookmarkStart" && typeof o.name === "string" && o.name !== "") {
      names.add(o.name);
    }
    for (const key in o) {
      const value = o[key];
      if (value && typeof value === "object") {
        stack.push(value);
      } else if (
        typeof value === "string" &&
        key.endsWith("Xml") &&
        value.includes("bookmarkStart")
      ) {
        for (const m of value.matchAll(RAW_BOOKMARK_NAME)) {
          const name = xmlDecode(m[3]!);
          if (name !== "") {
            names.add(name);
          }
        }
      }
    }
  }
  return names;
}

/** `name` with the lowest numeric suffix that makes it a name nobody uses, within Word's length limit. */
function freshBookmarkName(name: string, unavailable: ReadonlySet<string>): string {
  for (let n = 2; ; n++) {
    const suffix = `_${n}`;
    const candidate = name.slice(0, MAX_BOOKMARK_NAME - suffix.length) + suffix;
    if (!unavailable.has(candidate)) {
      return candidate;
    }
  }
}

/**
 * Field instructions whose first argument is a bookmark name: `REF`,
 * `PAGEREF`, `NOTEREF` and `GOTOBUTTON`; and `HYPERLINK`'s `\l` location.
 */
const FIELD_BOOKMARK_ARGUMENT =
  /^(\s*(?:REF|PAGEREF|NOTEREF|GOTOBUTTON)\s+)(?:"([^"]*)"|([^\s"\\]+))/i;
const HYPERLINK_LOCATION = /^(\s*HYPERLINK\b.*?\\l\s+)(?:"([^"]*)"|([^\s"\\]+))/i;

function renameInInstruction(instruction: string, renames: ReadonlyMap<string, string>): string {
  for (const pattern of [FIELD_BOOKMARK_ARGUMENT, HYPERLINK_LOCATION]) {
    const m = pattern.exec(instruction);
    if (!m) {
      continue;
    }
    const quoted = m[2] !== undefined;
    const renamed = renames.get(quoted ? m[2]! : m[3]!);
    if (renamed === undefined) {
      return instruction;
    }
    const argument = quoted ? `"${renamed}"` : renamed;
    return m[1] + argument + instruction.slice(m[0].length);
  }
  return instruction;
}

function renameInRawXml(xml: string, renames: ReadonlyMap<string, string>): string {
  if (renames.size === 0) {
    return xml;
  }
  const rename = (match: string, head: string, quote: string, value: string): string => {
    const renamed = renames.get(xmlDecode(value));
    return renamed === undefined ? match : `${head}${quote}${xmlEncodeAttr(renamed)}${quote}`;
  };
  return xml.replace(RAW_BOOKMARK_NAME, rename).replace(RAW_HYPERLINK_ANCHOR, rename);
}

/**
 * `node` with the bookmarks named in `renames` renamed, and with every modelled
 * reference to them — internal hyperlinks and bookmark-naming fields —
 * following. (Markup kept verbatim is renamed by {@link renameInRawXml}, on the
 * same pass that rewrites its other references.) Copy on write: whatever holds
 * no such name is returned as it is, so the input is never modified.
 */
function renameBookmarks<T>(node: T, renames: ReadonlyMap<string, string>): T {
  if (renames.size === 0 || !node || typeof node !== "object" || ArrayBuffer.isView(node)) {
    return node;
  }
  if (Array.isArray(node)) {
    let changed = false;
    const out = node.map(item => {
      const renamed = renameBookmarks(item, renames);
      changed ||= renamed !== item;
      return renamed;
    });
    return (changed ? out : node) as T;
  }
  const o = node as Record<string, unknown>;
  let copy: Record<string, unknown> | undefined;
  const set = (key: string, value: unknown): void => {
    copy ??= { ...o };
    copy[key] = value;
  };
  if ((o.type === "bookmarkStart" || o.type === "bookmarkEnd") && typeof o.name === "string") {
    const renamed = renames.get(o.name);
    if (renamed !== undefined) {
      set("name", renamed);
    }
  } else if (o.type === "hyperlink" && typeof o.anchor === "string") {
    const renamed = renames.get(o.anchor);
    if (renamed !== undefined) {
      set("anchor", renamed);
    }
  } else if (o.type === "field" && typeof o.instruction === "string") {
    const instruction = renameInInstruction(o.instruction, renames);
    if (instruction !== o.instruction) {
      set("instruction", instruction);
    }
  }
  for (const key of Object.keys(o)) {
    const value = o[key];
    if (value && typeof value === "object" && !(value instanceof Map)) {
      const renamed = renameBookmarks(value, renames);
      if (renamed !== value) {
        set(key, renamed);
      }
    }
  }
  return (copy ?? o) as T;
}

// =============================================================================
// Image / footnote / endnote / comment merge helpers
// =============================================================================

/** Rewrite the entries `list` gained from index `from` on. */
function rewriteFrom<T>(list: T[], from: number, rewriteEntry: (entry: T) => T): void {
  for (let i = from; i < list.length; i++) {
    list[i] = rewriteEntry(list[i]!);
  }
}

/**
 * The start tag of an element in markup kept verbatim that refers to a
 * numbering instance, comment or note by id.
 */
const MARKUP_ID_REFERENCE =
  /<((?:[A-Za-z_][\w.-]*:)?)(numId|commentRangeStart|commentRangeEnd|commentReference|footnoteReference|endnoteReference|bookmarkStart|bookmarkEnd|permStart|permEnd)\b([^>]*)>/g;

/** Bookmark and permission ids used by ranges in a document's markup kept verbatim. */
interface MarkupRangeIds {
  readonly bookmark: Set<number>;
  readonly perm: Set<number>;
}

const MARKUP_RANGE =
  /<((?:[A-Za-z_][\w.-]*:)?)(bookmarkStart|bookmarkEnd|permStart|permEnd)\b([^>]*)>/g;

function scanMarkupRangeIds(doc: DocxDocument): MarkupRangeIds {
  const out: MarkupRangeIds = { bookmark: new Set(), perm: new Set() };
  const stack: unknown[] = [
    doc.body,
    doc.bodyRangeMarkersAfter,
    doc.headers,
    doc.footers,
    doc.footnotes,
    doc.endnotes,
    doc.comments
  ];
  const seen = new Set<object>();
  while (stack.length > 0) {
    const node = stack.pop();
    if (node === null || typeof node !== "object" || ArrayBuffer.isView(node) || seen.has(node)) {
      continue;
    }
    seen.add(node);
    const values = node instanceof Map ? [...node.values()] : Object.entries(node);
    for (const entry of values) {
      const [key, value] = node instanceof Map ? ["", entry] : (entry as [string, unknown]);
      if (typeof value === "string" && key.endsWith("Xml")) {
        for (const m of value.matchAll(MARKUP_RANGE)) {
          const id = new RegExp(`\\s${m[1]}id\\s*=\\s*["'](-?\\d+)["']`).exec(m[3]!)?.[1];
          if (id !== undefined) {
            (m[2]!.startsWith("perm") ? out.perm : out.bookmark).add(Number(id));
          }
        }
      } else if (value !== null && typeof value === "object") {
        stack.push(value);
      }
    }
  }
  return out;
}

/**
 * Rename the ids of `incoming` that `taken` already holds to fresh ones, add
 * them all to `taken`, and return old → new for the renamed ones.
 */
function claimIds(taken: Set<number>, incoming: ReadonlySet<number>): Map<number, number> {
  const renamed = new Map<number, number>();
  let next = Math.max(maxOf(taken), maxOf(incoming), -1) + 1;
  for (const id of incoming) {
    if (taken.has(id)) {
      renamed.set(id, next++);
    }
  }
  for (const id of incoming) {
    taken.add(renamed.get(id) ?? id);
  }
  return renamed;
}

/**
 * `xml` with the numbering, comment and note ids it refers to renamed through
 * the merge's tables — the same tables its modelled content goes through, so
 * markup and model keep naming the same definitions — and its bookmark and
 * permission ranges renumbered where another document's verbatim markup
 * already uses the id. (Modelled ranges are the writer's to renumber; it keeps
 * them clear of verbatim ones.)
 */
function remapMarkupIds(xml: string, remap: RemapState): string {
  if (
    remap.numIdMap.size === 0 &&
    remap.commentIdMap.size === 0 &&
    remap.footnoteIdMap.size === 0 &&
    remap.endnoteIdMap.size === 0 &&
    !remap.markupBookmarkIdMap?.size &&
    !remap.markupPermIdMap?.size
  ) {
    return xml;
  }
  return xml.replace(MARKUP_ID_REFERENCE, (whole, prefix: string, name: string, attrs: string) => {
    const [attr, table]: [string, ReadonlyMap<number, number>] =
      name === "numId"
        ? ["val", remap.numIdMap]
        : name === "footnoteReference"
          ? ["id", remap.footnoteIdMap]
          : name === "endnoteReference"
            ? ["id", remap.endnoteIdMap]
            : name.startsWith("bookmark")
              ? ["id", remap.markupBookmarkIdMap ?? new Map()]
              : name.startsWith("perm")
                ? ["id", remap.markupPermIdMap ?? new Map()]
                : ["id", remap.commentIdMap];
    const pattern = new RegExp(`(\\s${prefix}${attr}\\s*=\\s*["'])(-?\\d+)(["'])`);
    const rewritten = attrs.replace(pattern, (m, open: string, value: string, close: string) => {
      const next = table.get(Number(value));
      return next === undefined ? m : `${open}${next}${close}`;
    });
    return rewritten === attrs ? whole : `<${prefix}${name}${rewritten}>`;
  });
}

/**
 * Merge `doc.images` into `mergedImages`, returning a map of original image
 * key → final key for every key that another picture already holds.
 *
 * Image keys (`rId` and `aliasRIds`) are document-wide, and two documents
 * number theirs independently, so `rId5` in one is routinely a different
 * picture from `rId5` in the other: such a key is renamed, and every
 * reference to it with it. A key that already names the same picture — the
 * same bytes — needs nothing; when all of an image's keys do, the image is
 * already merged and is not added again. File names are deduplicated too:
 * two documents both shipping a different `image1.png` keep both.
 */
function mergeImages(doc: DocxDocument, mergedImages: ImageDef[]): Map<string, string> {
  const keyMap = new Map<string, string>();
  if (!doc.images || doc.images.length === 0) {
    return keyMap;
  }
  const owners = new Map<string, ImageDef>();
  for (const m of mergedImages) {
    for (const key of imageKeys(m)) {
      if (!owners.has(key)) {
        owners.set(key, m);
      }
    }
  }
  const usedFileNames = new Set(mergedImages.map(m => m.fileName));
  const incoming = new Set(doc.images.flatMap(imageKeys));

  let nextSuffix = 1;
  const freshKey = (oldKey: string): string => {
    let candidate: string;
    do {
      candidate = `${oldKey}_m${nextSuffix++}`;
    } while (owners.has(candidate) || incoming.has(candidate));
    return candidate;
  };
  const freshFileName = (oldName: string): string => {
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

  const merged = new Set<ImageDef>(mergedImages);
  for (const img of doc.images) {
    // The very same entry — a document merged with itself, or two sharing
    // an image table — is already there under the same keys.
    if (merged.has(img)) {
      continue;
    }
    const keys = imageKeys(img);
    const kept: string[] = [];
    for (const key of keys) {
      const owner = owners.get(key);
      if (owner === undefined) {
        kept.push(key);
      } else if (!isSamePicture(owner, img)) {
        const renamed = freshKey(key);
        keyMap.set(key, renamed);
        kept.push(renamed);
      }
    }
    if (keys.length > 0 && kept.length === 0) {
      // Every key already names this very picture.
      continue;
    }
    const fileName = freshFileName(img.fileName);
    const unchanged =
      fileName === img.fileName &&
      kept.length === keys.length &&
      kept.every((key, i) => key === keys[i]);
    let next: ImageDef = img;
    if (!unchanged) {
      const { aliasRIds: _aliases, ...rest } = img;
      next = {
        ...rest,
        fileName,
        ...(kept.length > 0 ? { rId: kept[0] } : {}),
        ...(kept.length > 1 ? { aliasRIds: kept.slice(1) } : {})
      } as ImageDef;
    }
    mergedImages.push(next);
    merged.add(next);
    usedFileNames.add(fileName);
    for (const key of kept) {
      owners.set(key, next);
    }
  }
  return keyMap;
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
  mergedNumInstances: NumberingInstance[],
  picBulletIdMap: ReadonlyMap<number, number>
): { absIdMap: Map<number, number>; numIdMap: Map<number, number> } {
  const absIdMap = new Map<number, number>();
  const numIdMap = new Map<number, number>();

  if (doc.abstractNumberings && doc.abstractNumberings.length > 0) {
    const existingAbsIds = new Set(mergedAbstractNums.map(a => a.abstractNumId));
    let nextAbsId = existingAbsIds.size > 0 ? maxOf(existingAbsIds) + 1 : 0;
    for (const source of doc.abstractNumberings) {
      const abs = remapPicBullets(source, picBulletIdMap);
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

/** `abs` with its levels' picture-bullet ids renumbered through `idMap`. */
function remapPicBullets(
  abs: AbstractNumbering,
  idMap: ReadonlyMap<number, number>
): AbstractNumbering {
  if (
    idMap.size === 0 ||
    !abs.levels.some(l => l.picBulletId !== undefined && idMap.has(l.picBulletId))
  ) {
    return abs;
  }
  return {
    ...abs,
    levels: abs.levels.map(l =>
      l.picBulletId !== undefined && idMap.has(l.picBulletId)
        ? { ...l, picBulletId: idMap.get(l.picBulletId)! }
        : l
    )
  };
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
  const cloned = cloneBlockContentWithRemap(block, remap);
  return block.rangeMarkersBefore
    ? { ...cloned, rangeMarkersBefore: block.rangeMarkersBefore.map(m => cloneMarker(m, remap)) }
    : cloned;
}

function cloneBlockContentWithRemap(block: BodyContent, remap: RemapState): BodyContent {
  switch (block.type) {
    case "paragraph":
      return cloneParagraph(block, remap);
    case "table":
      return cloneTable(block, remap);
    case "sdt":
      return cloneSdt(block as StructuredDocumentTag, remap);
    case "floatingImage":
      return structuredClone(block);
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
      const shape = block as DrawingShape;
      const cloned = structuredClone(block) as Mutable<DrawingShape>;
      if (shape.textContent && shape.textContent.length > 0) {
        const text = shape.textContent.map(p => cloneParagraph(p, remap));
        cloned.textContent = text;
        // The markup kept for the text box stands for the text as read; the
        // clone of that text is still the text as read.
        if (shape._preserved?.read.textContent === shape.textContent) {
          cloned._preserved = {
            ...shape._preserved,
            read: { ...shape._preserved.read, textContent: text }
          };
        }
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
      // For chart / altChunk / opaqueDrawing / checkBox / math etc. there's
      // no nested paragraph content the merge cares about. Image references
      // anywhere in a block are renamed afterwards (`rewriteImageReferences`).
      return structuredClone(block);
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
    children: clonedChildren,
    ...(para.rangeMarkersAfter
      ? { rangeMarkersAfter: para.rangeMarkersAfter.map(m => cloneMarker(m, remap)) }
      : {})
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
      case "deletedRun":
      case "movedFromRun":
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

/** A copy of a range marker with its comment id remapped. */
function cloneMarker<T extends RangeMarker>(marker: T, remap: RemapState): T {
  const cloned = structuredClone(marker);
  remapParagraphChild(cloned, remap);
  return cloned;
}

function cloneMarkers(
  markers: readonly RangeMarker[] | undefined,
  remap: RemapState
): { rangeMarkersBefore?: RangeMarker[] } | undefined {
  return markers ? { rangeMarkersBefore: markers.map(m => cloneMarker(m, remap)) } : undefined;
}

function cloneTable(table: Table, remap: RemapState): Table {
  return {
    ...table,
    rows: table.rows.map(r => cloneRow(r, remap)),
    ...(table.rangeMarkersAfter
      ? { rangeMarkersAfter: table.rangeMarkersAfter.map(m => cloneMarker(m, remap)) }
      : {})
  } as Table;
}

function cloneRow(row: TableRow, remap: RemapState): TableRow {
  return {
    ...row,
    cells: row.cells.map(c => cloneCell(c, remap)),
    ...cloneMarkers(row.rangeMarkersBefore, remap),
    ...(row.rangeMarkersAfter
      ? { rangeMarkersAfter: row.rangeMarkersAfter.map(m => cloneMarker(m, remap)) }
      : {})
  } as TableRow;
}

function cloneCell(cell: TableCell, remap: RemapState): TableCell {
  return {
    ...cell,
    content: cell.content.map(item => cloneBlockWithRemap(item as BodyContent, remap)),
    ...cloneMarkers(cell.rangeMarkersBefore, remap)
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
    }),
    ...(sdt.contentRangeMarkersAfter
      ? { contentRangeMarkersAfter: sdt.contentRangeMarkersAfter.map(m => cloneMarker(m, remap)) }
      : {})
  } as StructuredDocumentTag;
}
