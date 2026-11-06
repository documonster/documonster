/**
 * Unified rendering context for DOCX XML serialization.
 * Passed to all writer functions to provide shared state.
 */

import { xmlDecode } from "@utils/xml-encode";
import { rebindRelationshipIds } from "@word/core/relationship-ids";
import type { DocxSecurityPolicy } from "@word/security/policy";
import { DEFAULT_SECURITY_POLICY } from "@word/security/policy";
import type { PreservedLink } from "@word/types";

/** ID generators for various document elements. */
export interface IdGenerators {
  nextDrawingId(): number;
  nextSdtId(): number;
  nextBookmarkId(): number;
  nextDocPrId(): number;
  nextChartId(): number;
  nextImagePartId(): number;
}

/**
 * Maps the bookmark ids a model carries to the ids written, so every
 * `w:bookmarkStart/@w:id` in the package is unique while each `w:bookmarkEnd`
 * still closes the bookmark it closed in the model.
 *
 * Bookmark ids only pair a start with its end — fields, hyperlinks and the
 * TOC refer to bookmarks by *name* — so the writer is free to renumber them,
 * which is what lets `Document.allocateBookmarkId` hand out ids in O(1) without
 * re-reading the tree: a caller who later pushes content with a clashing id is
 * reconciled here, in one pass, as the package is written.
 *
 * - A start keeps its id unless the id was already written or is *reserved*
 *   (used by markup this writer does not write: a part `Package.edit` leaves
 *   untouched); then it gets the next id above every id seen so far.
 *   Markers inside preserved raw XML go through the same map as they are
 *   written ({@link renumberRawRangeIds}), so they need no reservation.
 * - An end closes an open start that carried its model id. When the end names
 *   its bookmark (`BookmarkEnd.name`) it closes the most recent such start of
 *   that name, and otherwise (or when no open start has that name) the most
 *   recent such start. That is exact for ranges
 *   that are sequential or nested, and for ranges that overlap while sharing an
 *   id it is exact only when the ends carry names — without them the model
 *   itself does not say which end closes which start.
 * - An end that closes no open start is kept unchanged when its id is reserved
 *   (it closes a start in an untouched part) and otherwise dropped (`undefined`): written
 *   as-is it could close a different bookmark that was renumbered onto its id.
 * - A range does not span story parts. A start still open when its part ends
 *   ({@link closePart}) is removed from the part — the end that would have
 *   closed it either came before it, and was dropped, or is in another part,
 *   where it closes nothing and is dropped too. A range that cannot be paired
 *   in a part is therefore removed from it, start and end together.
 */
export interface BookmarkIdMap {
  /** Id to write for a `bookmarkStart` carrying `id` (and `name`) in the model. */
  start(id: number, name?: string): number;
  /**
   * Id to write for a `bookmarkEnd` carrying `id` (and optionally `name`) in
   * the model, or `undefined` when the end closes nothing and must be omitted.
   */
  end(id: number, name?: string): number | undefined;
  /**
   * End a story part: the written ids of the starts no end in it closed, which
   * are forgotten, so an end in a later part cannot close them.
   */
  closePart(): number[];
}

/**
 * A set of ids. Documents number bookmarks from 0 upwards, so ids below a
 * bound live in a bitmap — this is consulted for every bookmark of every
 * package written — and any other id in a `Set`.
 */
interface IdSet {
  /** Add `id`; `false` when it was already in the set. */
  add(id: number): boolean;
}

/** Ids below this live in the bitmap (2 MiB at most). */
const BITMAP_IDS = 1 << 24;

/** Whether `id` is an integer in `[0, BITMAP_IDS)`, the range kept in dense tables. */
function isDenseId(id: number): boolean {
  return id >= 0 && id < BITMAP_IDS && (id | 0) === id;
}

function createIdSet(): IdSet {
  let bits = new Uint32Array(64);
  const others = new Set<number>();
  return {
    add(id) {
      if (isDenseId(id)) {
        const word = id >>> 5;
        if (word >= bits.length) {
          const grown = new Uint32Array(Math.max(word + 1, bits.length * 2));
          grown.set(bits);
          bits = grown;
        }
        const bit = 1 << (id & 31);
        const was = bits[word]!;
        if ((was & bit) !== 0) {
          return false;
        }
        bits[word] = was | bit;
        return true;
      }
      if (others.has(id)) {
        return false;
      }
      others.add(id);
      return true;
    }
  };
}

/** A fresh {@link BookmarkIdMap}; one per written package. */
export function createBookmarkIdMap(reserved?: Iterable<number>): BookmarkIdMap {
  interface OpenStart {
    readonly out: number;
    readonly name: string | undefined;
  }
  const written = createIdSet();
  const kept = new Set<number>();
  // Open starts by model id. Almost every id is open once at a time and is
  // small, so that case lives in dense tables indexed by id — `openOut` holds
  // the written id plus one (0: nothing open, MULTI: several, listed in
  // `open`) and `openName` its name — and allocates nothing per bookmark;
  // this runs for every bookmark of every package written. Several starts
  // open under one id, and ids outside the dense range, go in `open`, where a
  // single entry is stored as itself and only a repeated id gets a list.
  const MULTI = -1;
  let openOut = new Float64Array(64);
  const openName: (string | undefined)[] = [];
  const open = new Map<number, OpenStart | OpenStart[]>();
  let highest = -1;
  /** Starts open, so a part that closed them all ends without a scan. */
  let openCount = 0;
  for (const id of reserved ?? []) {
    kept.add(id);
    written.add(id);
    if (id > highest) {
      highest = id;
    }
  }
  /** Close one of several starts open under `id`: the latest named `name`, else the latest. */
  function closeOneOf(list: OpenStart[], name: string | undefined): OpenStart {
    let at = list.length - 1;
    if (name !== undefined) {
      while (at >= 0 && list[at]!.name !== name) {
        at--;
      }
      if (at < 0) {
        at = list.length - 1;
      }
    }
    return list.splice(at, 1)[0]!;
  }
  return {
    start(id, name) {
      let out = id;
      if (!written.add(out)) {
        out = highest + 1;
        written.add(out);
      }
      if (out > highest) {
        highest = out;
      }
      openCount++;
      if (isDenseId(id)) {
        if (id >= openOut.length) {
          const grown = new Float64Array(Math.max(id + 1, openOut.length * 2));
          grown.set(openOut);
          openOut = grown;
        }
        const state = openOut[id]!;
        if (state === 0) {
          openOut[id] = out + 1;
          openName[id] = name;
        } else if (state === MULTI) {
          (open.get(id) as OpenStart[]).push({ out, name });
        } else {
          open.set(id, [
            { out: state - 1, name: openName[id] },
            { out, name }
          ]);
          openOut[id] = MULTI;
          openName[id] = undefined;
        }
        return out;
      }
      const entry: OpenStart = { out, name };
      const prev = open.get(id);
      if (prev === undefined) {
        open.set(id, entry);
      } else if (Array.isArray(prev)) {
        prev.push(entry);
      } else {
        open.set(id, [prev, entry]);
      }
      return out;
    },
    end(id, name) {
      if (isDenseId(id)) {
        const state = id < openOut.length ? openOut[id]! : 0;
        if (state === 0) {
          return kept.has(id) ? id : undefined;
        }
        openCount--;
        if (state !== MULTI) {
          openOut[id] = 0;
          openName[id] = undefined;
          return state - 1;
        }
        const list = open.get(id) as OpenStart[];
        const closed = closeOneOf(list, name);
        if (list.length === 1) {
          open.delete(id);
          openOut[id] = list[0]!.out + 1;
          openName[id] = list[0]!.name;
        }
        return closed.out;
      }
      const prev = open.get(id);
      if (prev === undefined) {
        return kept.has(id) ? id : undefined;
      }
      openCount--;
      if (!Array.isArray(prev)) {
        open.delete(id);
        return prev.out;
      }
      const closed = closeOneOf(prev, name);
      if (prev.length === 1) {
        open.set(id, prev[0]!);
      }
      return closed.out;
    },
    closePart() {
      if (openCount === 0) {
        return [];
      }
      const unclosed: number[] = [];
      for (let id = 0; id < openOut.length; id++) {
        const state = openOut[id]!;
        if (state > 0) {
          unclosed.push(state - 1);
          openOut[id] = 0;
          openName[id] = undefined;
        } else if (state === MULTI) {
          openOut[id] = 0;
        }
      }
      for (const entry of open.values()) {
        for (const start of Array.isArray(entry) ? entry : [entry]) {
          unclosed.push(start.out);
        }
      }
      open.clear();
      openCount = 0;
      return unclosed;
    }
  };
}

/** A `bookmarkStart` / `permStart` element, self-closing or with an empty body. */
const RANGE_START_ELEMENT =
  /<((?:[\w.-]+:)?(bookmarkStart|permStart))\b([^>]*?)(?:\/>|>\s*<\/\1>)/g;

/**
 * `xml`, a story part just written with `helpers`, without the starts no end
 * in it closed (see {@link BookmarkIdMap.closePart}). Their ends are already
 * absent, so the part is left holding whole ranges only. Written start ids are
 * unique in the package, so the id names the element.
 */
export function closeStoryPart(
  xml: string,
  helpers: Pick<RenderHelpers, "bookmarkIds" | "permIds"> | undefined
): string {
  const bookmarks = helpers?.bookmarkIds?.closePart() ?? [];
  const perms = helpers?.permIds?.closePart() ?? [];
  if (bookmarks.length === 0 && perms.length === 0) {
    return xml;
  }
  const remove = { bookmarkStart: new Set(bookmarks), permStart: new Set(perms) };
  return xml.replace(RANGE_START_ELEMENT, (element, _tag, local: string, attrs: string) => {
    const id = RAW_ID_ATTR.exec(attrs)?.[3];
    return id !== undefined && remove[local as keyof typeof remove].has(Number(id)) ? "" : element;
  });
}

/**
 * Hands out the `w:id` written for each revision — the wrappers around
 * revised content (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo`), the marks of
 * a revised paragraph mark, table row or cell, and every property change
 * (`w:pPrChange`, `w:rPrChange`, `w:tblPrChange`, `w:trPrChange`,
 * `w:tcPrChange`, `w:sectPrChange`) — so every one in the package is unique,
 * as ECMA-376 §17.13.1 requires of an annotation identifier. A revision id
 * pairs nothing — accepting and rejecting work on the model — so a revision
 * keeps its model id unless that id was already written or is reserved (used
 * by a part `Package.edit` leaves untouched), and otherwise takes the next id
 * above every id seen so far.
 */
export interface RevisionIdMap {
  /** Id to write for a revision carrying `id` in the model. */
  claim(id: number): number;
}

/** A fresh {@link RevisionIdMap}; one per written package. */
export function createRevisionIdMap(reserved?: Iterable<number>): RevisionIdMap {
  const written = createIdSet();
  let highest = -1;
  for (const id of reserved ?? []) {
    written.add(id);
    if (id > highest) {
      highest = id;
    }
  }
  return {
    claim(id) {
      let out = id;
      if (!written.add(out)) {
        out = highest + 1;
        written.add(out);
      }
      if (out > highest) {
        highest = out;
      }
      return out;
    }
  };
}

/**
 * The revision map of the package being written. Revision ids are written by
 * renderers that receive no context — the run and paragraph property writers
 * serve styles and numbering as well as content — so the writer sets it for
 * the duration of a write rather than passing it down to each of them.
 *
 * Rendering is synchronous, so a map entered and left around a render call
 * cannot be seen by any other write. An asynchronous write enters it inside
 * {@link withRevisionScope}, which leaves it when the write first awaits.
 */
let activeRevisionIds: RevisionIdMap | undefined;

/**
 * The id to write for a revision carrying `id` in the model: claimed from the
 * revision map in effect, or `id` itself when none is (a renderer called on
 * its own).
 */
export function writtenRevisionId(id: number): number {
  return activeRevisionIds === undefined ? id : activeRevisionIds.claim(id);
}

/**
 * Put `map` in effect unless a map already is — the outermost write owns the
 * id space, so content it renders through a context of its own (the glossary
 * document) still draws from it — and return what to hand
 * {@link exitRevisionIds} afterwards.
 */
export function enterRevisionIds(map: RevisionIdMap | undefined): RevisionIdMap | undefined {
  const previous = activeRevisionIds;
  if (previous === undefined) {
    activeRevisionIds = map;
  }
  return previous;
}

/** Restore the revision map in effect before the matching {@link enterRevisionIds}. */
export function exitRevisionIds(previous: RevisionIdMap | undefined): void {
  activeRevisionIds = previous;
}

/**
 * Leave the map an asynchronous write entered with {@link enterRevisionIds}, at
 * the end of the synchronous section that renders its parts.
 *
 * It must still be in effect: {@link withRevisionScope} takes it out of effect
 * at the write's first await, so finding it gone means an await was introduced
 * before the parts were all rendered — and the parts after it were written with
 * the model's ids, which may repeat. That is a library defect, and it is
 * reported rather than written into a package.
 */
export function leaveRevisionIds(
  previous: RevisionIdMap | undefined,
  map: RevisionIdMap | undefined
): void {
  if (activeRevisionIds !== (previous ?? map)) {
    throw new Error(
      "Internal error: the revision id map was left before every part was rendered " +
        "(an await inside the synchronous render section)"
    );
  }
  activeRevisionIds = previous;
}

/**
 * Call `fn`, and restore the revision map in effect before it when it returns.
 * For an asynchronous `fn` that is when it first awaits: a map it entered
 * without leaving is therefore in effect only while it runs synchronously,
 * never while other code runs in the meantime.
 */
export function withRevisionScope<T>(fn: () => T): T {
  const previous = activeRevisionIds;
  try {
    return fn();
  } finally {
    activeRevisionIds = previous;
  }
}

/** The start element of a range whose ids the writer reconciles. */
type RangeStartElement = "bookmarkStart" | "permStart";

const RAW_RANGE_START: Record<RangeStartElement, RegExp> = {
  bookmarkStart: /<(?:[\w.-]+:)?bookmarkStart\b[^>]*?\s(?:[\w.-]+:)?id\s*=\s*["'](-?\d+)["']/g,
  permStart: /<(?:[\w.-]+:)?permStart\b[^>]*?\s(?:[\w.-]+:)?id\s*=\s*["'](-?\d+)["']/g
};

/** Add the id of every `bookmarkStart` in an XML fragment to `out`. */
export function collectBookmarkStartIds(xml: string, out: Set<number>): void {
  collectRangeStartIds(xml, out, "bookmarkStart");
}

/** Add the id of every `element` (a range start) in an XML fragment to `out`. */
export function collectRangeStartIds(
  xml: string,
  out: Set<number>,
  element: RangeStartElement
): void {
  if (!xml.includes(element)) {
    return;
  }
  for (const m of xml.matchAll(RAW_RANGE_START[element])) {
    out.add(Number(m[1]));
  }
}

/**
 * The range markers of preserved raw XML, and the `mc:AlternateContent`
 * structure around them. A closing tag is only matched for a marker element,
 * which is empty by schema but may be written as an open/close pair.
 */
const RAW_MARKUP =
  /<(\/?)(?:[\w.-]+:)?(bookmarkStart|bookmarkEnd|permStart|permEnd|AlternateContent|Choice|Fallback)\b([^>]*?)(\/?)>/g;
const RAW_ID_ATTR = /(\s(?:[\w.-]+:)?id\s*=\s*)(["'])(-?\d+)\2/;
const RAW_NAME_ATTR = /\s(?:[\w.-]+:)?name\s*=\s*(["'])([^"']*)\1/;

/**
 * One `mc:AlternateContent` being rewritten. Its alternatives are renderings
 * of the same content, of which a consumer keeps exactly one, so every
 * alternative after the first repeats the ids the first was given — the
 * markers it closes or opens are the same ranges, not new ones.
 */
interface AlternativesFrame {
  /** 0 before the first alternative, then 1, 2, … */
  alternative: number;
  /** Ids written for the first alternative's markers, in order, by marker kind and model id. */
  readonly written: Map<string, (number | undefined)[]>;
  /** How far each queue in {@link written} has been replayed in the current alternative. */
  replayed: Map<string, number>;
}

const RAW_DOC_PR_ID = /<(?:[\w.-]+:)?docPr\b[^>]*?\sid\s*=\s*["'](\d+)["']/g;

/** Add the id of every `wp:docPr` in an XML fragment to `out`. */
export function collectDocPrIds(xml: string, out: Set<number>): void {
  if (!xml.includes("docPr")) {
    return;
  }
  for (const m of xml.matchAll(RAW_DOC_PR_ID)) {
    out.add(Number(m[1]));
  }
}

/**
 * A revision element's start tag: a wrapper around revised content, the mark
 * of a revised paragraph mark, row or cell, or a property change.
 */
const RAW_REVISION_TAG =
  /<(?:[\w.-]+:)?(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|(?:p|r|tbl|tblPrEx|tr|tc|sect)PrChange|tblGridChange|numberingChange)\b([^>]*)>/g;

/** Whether an XML fragment may hold a revision element (a cheap test before any regex runs). */
function mayHoldRevision(xml: string): boolean {
  return (
    xml.indexOf("ins") !== -1 ||
    xml.indexOf("del") !== -1 ||
    xml.indexOf("move") !== -1 ||
    xml.indexOf("Change") !== -1 ||
    xml.indexOf("cell") !== -1
  );
}

/** Add the id of every revision element (see {@link RevisionIdMap}) in an XML fragment to `out`. */
export function collectRevisionIds(xml: string, out: Set<number>): void {
  if (!mayHoldRevision(xml)) {
    return;
  }
  for (const m of xml.matchAll(RAW_REVISION_TAG)) {
    const id = RAW_ID_ATTR.exec(m[1]!);
    if (id) {
      out.add(Number(id[3]));
    }
  }
}

/**
 * `xml` — preserved raw WordprocessingML about to be copied into a part —
 * with the id of every revision element in it taken through `revisionIds`,
 * as a modelled revision's is, so a revision kept verbatim cannot share its
 * id with one the writer wrote.
 */
export function renumberRawRevisionIds(
  xml: string,
  revisionIds: RevisionIdMap | undefined
): string {
  if (!revisionIds || !mayHoldRevision(xml)) {
    return xml;
  }
  return xml.replace(RAW_REVISION_TAG, (tag: string, attrs: string) => {
    const idMatch = RAW_ID_ATTR.exec(attrs);
    if (!idMatch) {
      return tag;
    }
    const id = Number(idMatch[3]);
    const written = revisionIds.claim(id);
    if (written === id) {
      return tag;
    }
    const idAt = tag.length - 1 - attrs.length + idMatch.index + idMatch[1]!.length + 1;
    return tag.slice(0, idAt) + String(written) + tag.slice(idAt + idMatch[3]!.length);
  });
}

/**
 * `xml` — preserved raw WordprocessingML about to be copied into a story part
 * — with the ids of its bookmark and range-permission markers taken through
 * the package's {@link BookmarkIdMap}s, exactly as a model marker at the same
 * place would be. Raw markers therefore pair with model markers in either
 * direction (a bookmark that opens in the model and closes inside a preserved
 * `w:bdo`, or the reverse), a repeated copy of a raw fragment gets ids of its
 * own, and an end that closes nothing is removed for the reason a model one
 * is. Only the `id` attribute of those elements changes; the rest of the
 * fragment is written as it was.
 *
 * Because the ids are reconciled where the fragment is written, no pass over
 * the model has to look for raw XML before a package is written.
 */
export function renumberRawRangeIds(
  xml: string,
  helpers: Pick<RenderHelpers, "bookmarkIds" | "permIds"> | undefined
): string {
  const bookmarkIds = helpers?.bookmarkIds;
  const permIds = helpers?.permIds;
  if (
    (!bookmarkIds && !permIds) ||
    (xml.indexOf("bookmark") === -1 && xml.indexOf("perm") === -1)
  ) {
    return xml;
  }
  const frames: AlternativesFrame[] = [];
  let out = "";
  let last = 0;
  /** Local name of a marker whose open tag was removed, so its close tag goes too. */
  let dropClose: string | undefined;
  for (const m of xml.matchAll(RAW_MARKUP)) {
    const tag = m[0];
    const closing = m[1] === "/";
    const local = m[2]!;
    const attrs = m[3]!;
    const selfClosing = m[4] === "/";
    const at = m.index;
    if (local === "AlternateContent") {
      if (closing) {
        frames.pop();
      } else if (!selfClosing) {
        frames.push({ alternative: 0, written: new Map(), replayed: new Map() });
      }
      continue;
    }
    if (local === "Choice" || local === "Fallback") {
      const frame = frames[frames.length - 1];
      if (frame && !closing && !selfClosing) {
        frame.alternative++;
        frame.replayed = new Map();
      }
      continue;
    }
    if (closing) {
      if (dropClose === local) {
        out += xml.slice(last, at);
        last = at + tag.length;
        dropClose = undefined;
      }
      continue;
    }
    const map = local.startsWith("bookmark") ? bookmarkIds : permIds;
    const idMatch = map ? RAW_ID_ATTR.exec(attrs) : null;
    if (!map || !idMatch) {
      continue;
    }
    const id = Number(idMatch[3]);
    const key = `${local}:${id}`;
    let written: number | undefined;
    // An alternative after the first replays the ids the first was given.
    const replaying = frames.find(f => f.alternative > 1);
    const queue = replaying?.written.get(key);
    const replayAt = replaying?.replayed.get(key) ?? 0;
    if (replaying && queue && replayAt < queue.length) {
      written = queue[replayAt];
      replaying.replayed.set(key, replayAt + 1);
    } else {
      const rawName = local.startsWith("bookmark") ? RAW_NAME_ATTR.exec(attrs)?.[2] : undefined;
      const name = rawName === undefined ? undefined : xmlDecode(rawName);
      written = local.endsWith("Start") ? map.start(id, name) : map.end(id, name);
    }
    for (const frame of frames) {
      if (frame.alternative === 1) {
        const recorded = frame.written.get(key);
        if (recorded) {
          recorded.push(written);
        } else {
          frame.written.set(key, [written]);
        }
      }
    }
    if (written === id) {
      continue;
    }
    out += xml.slice(last, at);
    last = at + tag.length;
    if (written === undefined) {
      if (!selfClosing) {
        dropClose = local;
      }
      continue;
    }
    const attrsAt = last - 1 - (selfClosing ? 1 : 0) - attrs.length;
    const idAt = attrsAt + idMatch.index + idMatch[1]!.length + 1;
    out += xml.slice(at, idAt) + String(written) + xml.slice(idAt + idMatch[3]!.length, last);
  }
  return last === 0 ? xml : out + xml.slice(last);
}

/** Registers a preserved hyperlink in the part being written; see `DocxRenderContext.preservedLinkRId`. */
export type PreservedLinkRId = (link: PreservedLink) => string | undefined;

/**
 * Lightweight, content-renderer-facing subset of DocxRenderContext.
 *
 * Most leaf renderers (run/paragraph/table/image) only need the rId remap and
 * the hyperlink rId map; passing the full DocxRenderContext would couple them
 * to chartRIds/securityPolicy/etc. They take this `RenderHelpers` instead.
 */
export interface RenderHelpers {
  /** See DocxRenderContext.imageRIdRemap. */
  readonly imageRemap?: ReadonlyMap<string, string>;
  /** See DocxRenderContext.hyperlinkRIds. */
  readonly hyperlinkRIds?: ReadonlyWeakMap<object, string>;
  /** See DocxRenderContext.preservedLinkRId. */
  readonly preservedLinkRId?: PreservedLinkRId;
  /**
   * Allocates the next document-wide unique drawing object id, used for
   * `wp:docPr/@id` (and the matching `pic:cNvPr/@id` / `wps:cNvPr/@id`).
   *
   * Word requires every drawing object id to be a unique positive integer
   * across the entire document — including body, headers, footers, footnotes,
   * endnotes, comments and text boxes. The packager seeds a single counter on
   * the render context and exposes it here so every drawing renderer draws
   * from the same id space, regardless of what (possibly duplicate or unset)
   * `drawingId` the model carries.
   *
   * When undefined (e.g. a standalone renderer call without a context),
   * renderers fall back to the model's `drawingId` or `1`.
   */
  readonly nextDocPrId?: () => number;
  /**
   * Document-wide bookmark id reconciliation (see {@link BookmarkIdMap}).
   * When undefined (a standalone renderer call), model ids are written as-is.
   */
  readonly bookmarkIds?: BookmarkIdMap;
  /**
   * Document-wide range-permission id reconciliation. Permission ranges pair
   * `w:permStart` with `w:permEnd` by id exactly as bookmarks do, so the same
   * {@link BookmarkIdMap} rules keep every written start id unique and every
   * end on its start. When undefined, model ids are written as-is.
   */
  readonly permIds?: BookmarkIdMap;
  /**
   * Raw XML output policy. Controls how preserved/opaque rawXml fragments
   * (opaqueRun, opaqueParagraphChild, opaqueDrawing, _advancedFillXml, …)
   * are emitted. Defaults to `"preserve"` when undefined for backwards
   * compatibility with callers that built helpers without a context.
   */
  readonly rawXmlPolicy?: "preserve" | "strip" | "reject";
}

/**
 * The {@link RenderHelpers} for content rendered into the part `ctx` belongs
 * to. Every paragraph-rendering path takes its helpers from here (or from the
 * part-specific helpers the packager builds), so none of them can bypass the
 * package's bookmark, image and hyperlink maps.
 */
export function renderHelpersFor(ctx: DocxRenderContext): RenderHelpers {
  return {
    imageRemap: ctx.imageRIdRemap,
    hyperlinkRIds: ctx.hyperlinkRIds,
    preservedLinkRId: ctx.preservedLinkRId,
    nextDocPrId: ctx.ids.nextDocPrId,
    bookmarkIds: ctx.bookmarkIds,
    permIds: ctx.permIds,
    rawXmlPolicy: ctx.rawXmlPolicy
  };
}

/**
 * Read-only WeakMap projection.
 * (TS lib doesn't ship one; we mirror the readonly subset of WeakMap.)
 */
interface ReadonlyWeakMap<K extends WeakKey, V> {
  get(key: K): V | undefined;
  has(key: K): boolean;
}

/** Render context passed to all writer functions. */
export interface DocxRenderContext {
  /** Current part being rendered, fully qualified (e.g. `"/word/document.xml"`). */
  readonly partName: string;
  /** Security policy controlling raw XML, external targets, etc. */
  readonly securityPolicy: DocxSecurityPolicy;
  /** ID generators (shared across all parts for a single document render). */
  readonly ids: IdGenerators;
  /** Raw XML output policy. If "reject", opaque content throws instead of being written. */
  readonly rawXmlPolicy: "preserve" | "strip" | "reject";
  /** Chart/ChartEx object → rId mapping. Populated by packager before rendering. */
  readonly chartRIds: Map<object, string>;
  /**
   * Optional remap for image relationship IDs that the packager had to rewrite
   * (e.g. when the model's image rId collided with a non-image relationship in
   * the same .rels file). Writers that emit `r:embed` look up this map first
   * and fall back to the original rId stored on the model. The map is keyed by
   * the original (model-side) rId.
   */
  readonly imageRIdRemap: Map<string, string>;
  /**
   * External hyperlink object → rId mapping. The packager registers a
   * relationship for each hyperlink with a `url` and stores the assigned rId
   * here, keyed by the hyperlink object identity. The paragraph writer reads
   * this when emitting `<w:hyperlink r:id="...">` so the model itself is
   * never mutated.
   */
  readonly hyperlinkRIds: WeakMap<object, string>;
  /**
   * The relationship id of a hyperlink that markup kept verbatim refers to (a
   * drawing's `a:hlinkClick`), registered in the part's `.rels` the first time
   * the markup is written — so no pass over the model has to look for such
   * links before a part is rendered. `undefined` when the link's target is
   * not one the package may carry. Without it, `hyperlinkRIds` is consulted.
   */
  readonly preservedLinkRId?: PreservedLinkRId;
  /**
   * AltChunk object → rId mapping. Populated by the packager when registering
   * the relationship for each `w:altChunk`. Renderers prefer this map over
   * `AltChunk.rId` so the caller's model is never mutated, including when
   * the altChunk is nested inside a table cell or SDT and was therefore not
   * cloned by the shallow body copy.
   */
  readonly altChunkRIds: WeakMap<object, string>;
  /**
   * Bookmark id reconciliation shared by every part of one package (see
   * {@link BookmarkIdMap}). Optional so a hand-built context stays valid; one
   * from `createRenderContext` always has it.
   */
  readonly bookmarkIds?: BookmarkIdMap;
  /** Range-permission id reconciliation, as {@link bookmarkIds} (see `RenderHelpers.permIds`). */
  readonly permIds?: BookmarkIdMap;
  /**
   * Revision id allocation shared by every part of one package (see
   * {@link RevisionIdMap}). Unlike the maps above it does not travel in
   * {@link RenderHelpers}: revision ids are also written by property renderers
   * (styles, numbering) that take no helpers, so the map in effect is set for
   * the duration of a write instead ({@link enterRevisionIds}).
   */
  readonly revisionIds?: RevisionIdMap;
}

/** Create ID generators with given starting values. */
export function createIdGenerators(startValues?: {
  drawingId?: number;
  sdtId?: number;
  bookmarkId?: number;
  docPrId?: number;
  chartId?: number;
  imagePartId?: number;
  /**
   * Drawing ids {@link IdGenerators.nextDocPrId} must not hand out: those of
   * parts the writer passes through untouched (`Package.edit`).
   */
  reservedDocPrIds?: Iterable<number>;
}): IdGenerators {
  let drawingId = startValues?.drawingId ?? 1;
  let sdtId = startValues?.sdtId ?? 1;
  let bookmarkId = startValues?.bookmarkId ?? 0;
  let docPrId = startValues?.docPrId ?? 1;
  let chartId = startValues?.chartId ?? 1;
  let imagePartId = startValues?.imagePartId ?? 1;
  const reservedDocPrIds = new Set(startValues?.reservedDocPrIds);

  return {
    nextDrawingId: () => drawingId++,
    nextSdtId: () => sdtId++,
    nextBookmarkId: () => bookmarkId++,
    nextDocPrId: () => {
      while (reservedDocPrIds.has(docPrId)) {
        docPrId++;
      }
      return docPrId++;
    },
    nextChartId: () => chartId++,
    nextImagePartId: () => imagePartId++
  };
}

/** Create a default render context. */
export function createRenderContext(options?: {
  partName?: string;
  securityPolicy?: DocxSecurityPolicy;
  chartRIds?: Map<object, string>;
  imageRIdRemap?: Map<string, string>;
  hyperlinkRIds?: WeakMap<object, string>;
  preservedLinkRId?: PreservedLinkRId;
  altChunkRIds?: WeakMap<object, string>;
  /**
   * Pre-built ID generators. When provided, the caller is responsible for
   * seeding them (e.g. with the maximum existing SDT id in the document so
   * auto-assigned IDs don't collide with author-supplied ones).
   */
  ids?: IdGenerators;
  /** Bookmark ids used by markup the writer will not renumber (see {@link BookmarkIdMap}). */
  reservedBookmarkIds?: Iterable<number>;
  /** Range-permission ids used by markup the writer will not renumber. */
  reservedPermIds?: Iterable<number>;
  /** Revision ids used by markup the writer will not renumber. */
  reservedRevisionIds?: Iterable<number>;
}): DocxRenderContext {
  const securityPolicy = options?.securityPolicy ?? DEFAULT_SECURITY_POLICY;

  return {
    partName: options?.partName ?? "/word/document.xml",
    securityPolicy,
    ids: options?.ids ?? createIdGenerators(),
    rawXmlPolicy: securityPolicy.rawXmlPolicy ?? "preserve",
    chartRIds: options?.chartRIds ?? new Map(),
    imageRIdRemap: options?.imageRIdRemap ?? new Map(),
    hyperlinkRIds: options?.hyperlinkRIds ?? new WeakMap(),
    ...(options?.preservedLinkRId ? { preservedLinkRId: options.preservedLinkRId } : {}),
    altChunkRIds: options?.altChunkRIds ?? new WeakMap(),
    bookmarkIds: createBookmarkIdMap(options?.reservedBookmarkIds),
    permIds: createBookmarkIdMap(options?.reservedPermIds),
    revisionIds: createRevisionIdMap(options?.reservedRevisionIds)
  };
}

/** A `docPr` element's start tag, whatever its prefix, and the `id` attribute in it. */
const DOC_PR_TAG = /<((?:[A-Za-z_][\w.-]*:)?docPr)\b([^>]*)>/g;
const ID_ATTRIBUTE = /(\sid\s*=\s*)(?:"[^"]*"|'[^']*')/;

/**
 * Markup kept verbatim (`rawXml` of an opaque run, paragraph child, drawing or
 * shape), as it must be written into the part `helpers` belong to:
 *
 * - its image references follow the part's remap (`helpers.imageRemap`), as
 *   modelled ones do;
 * - its bookmark and range-permission markers go through the package's id
 *   maps ({@link renumberRawRangeIds}), so they pair with model markers in
 *   document order and need no reservation pass over the model;
 * - its revision ids are taken from the revision map in effect
 *   ({@link renumberRawRevisionIds}), as modelled ones are;
 * - every `wp:docPr` id is drawn from the document-wide drawing allocator
 *   (`helpers.nextDocPrId`), the one every modelled drawing uses, so a drawing
 *   kept verbatim cannot collide with one the writer numbered.
 *
 * Every writer path that emits such markup goes through here.
 */
export function prepareRawMarkup(
  xml: string,
  helpers: Pick<RenderHelpers, "imageRemap" | "nextDocPrId" | "bookmarkIds" | "permIds"> | undefined
): string {
  let out = renumberRawRevisionIds(
    renumberRawRangeIds(rebindRelationshipIds(xml, helpers?.imageRemap), helpers),
    activeRevisionIds
  );
  const next = helpers?.nextDocPrId;
  if (next && out.includes("docPr")) {
    out = out.replace(DOC_PR_TAG, (_whole, name: string, attrs: string) => {
      const id = String(next());
      const rewritten = ID_ATTRIBUTE.test(attrs)
        ? attrs.replace(ID_ATTRIBUTE, `$1"${id}"`)
        : ` id="${id}"${attrs}`;
      return `<${name}${rewritten}>`;
    });
  }
  return out;
}
