/**
 * Folding block-level range markers into the blocks around them, for a
 * container (or an output format) that cannot hold them between blocks. See
 * {@link RangeMarker}.
 */

import { RANGE_MARKER_TYPES, isRangeMarker, isSpanMarker } from "@word/core/text-utils";
import type { SpanMarker } from "@word/core/text-utils";
import { walkBlocks } from "@word/core/walker";
import type {
  BlockRangeMarkers,
  BodyContent,
  CommentReference,
  FieldRangeMarker,
  Paragraph,
  ParagraphChild,
  PermEnd,
  PermStart,
  RangeMarker,
  Run,
  StructuredDocumentTag,
  Table
} from "@word/types";

type MarkerHost = Paragraph | Table | Run | StructuredDocumentTag;

/**
 * A range marker, or a comment's reference mark. Code that removes content
 * keeps both where the content was: a comment whose reference is dropped is a
 * comment with no anchor, as a range whose end is dropped is no range. A
 * reference is a paragraph child, so only a paragraph can hold one.
 */
export type RangeAnchor = RangeMarker | CommentReference;

/** Whether a model node is a {@link RangeAnchor}. */
export function isRangeAnchor(node: object): node is RangeAnchor {
  return (
    isRangeMarker(node) ||
    ("type" in node && (node as { type: string }).type === "commentReference")
  );
}

/** Whether an anchor closes something, so it belongs after the content before it. */
function isClosingAnchor(anchor: RangeAnchor): boolean {
  return anchor.type === "commentReference" || anchor.type.endsWith("End");
}

/** `block` with `markers` put at its start (`atEnd` false) or end, or `undefined` if it cannot hold them. */
export function hostRangeMarkers<T extends object>(
  host: T,
  markers: readonly RangeAnchor[],
  atEnd: boolean
): T | undefined {
  return hostIn(host as unknown as MarkerHost, markers, atEnd) as T | undefined;
}

function hostIn(
  block: MarkerHost,
  markers: readonly RangeAnchor[],
  atEnd: boolean
): MarkerHost | undefined {
  const references = markers.some(m => m.type === "commentReference");
  if (!("type" in block)) {
    // A run (it has no `type`) can hold markers before it, but not after —
    // and no comment reference, which is a paragraph child.
    return atEnd || references
      ? undefined
      : withRangeMarkersBefore(block, markers as readonly RangeMarker[]);
  }
  switch (block.type) {
    case "paragraph": {
      if (atEnd && block.rangeMarkersAfter && block.rangeMarkersAfter.length > 0) {
        if (!references) {
          return {
            ...block,
            rangeMarkersAfter: [...block.rangeMarkersAfter, ...(markers as readonly RangeMarker[])]
          };
        }
        // A reference cannot sit among the markers after the children, so
        // they all become children again, in order.
        const { rangeMarkersAfter, ...rest } = block;
        return withParagraphChildren(rest as Paragraph, [
          ...block.children,
          ...rangeMarkersAfter,
          ...markers
        ]);
      }
      const arranged = arrangeParagraphChildren(
        atEnd ? [...block.children, ...markers] : [...markers, ...block.children]
      );
      const after = [...(arranged.rangeMarkersAfter ?? []), ...(block.rangeMarkersAfter ?? [])];
      const { rangeMarkersAfter: _after, ...rest } = block;
      return {
        ...rest,
        children: arranged.children,
        ...(after.length > 0 ? { rangeMarkersAfter: after } : {})
      };
    }
    case "table": {
      if (references) {
        return undefined;
      }
      const tableMarkers = markers as readonly RangeMarker[];
      if (atEnd || block.rows.length === 0) {
        return {
          ...block,
          rangeMarkersAfter: [...(block.rangeMarkersAfter ?? []), ...tableMarkers]
        };
      }
      const [first, ...rest] = block.rows;
      return {
        ...block,
        rows: [
          {
            ...first!,
            rangeMarkersBefore: [...tableMarkers, ...(first!.rangeMarkersBefore ?? [])]
          },
          ...rest
        ]
      };
    }
    case "sdt": {
      const content = [...block.content];
      for (let k = 0; k < content.length; k++) {
        const i = atEnd ? content.length - 1 - k : k;
        const hosted = hostIn(content[i]!, markers, atEnd);
        if (hosted) {
          content[i] = hosted as (typeof content)[number];
          return { ...block, content };
        }
      }
      return undefined;
    }
    default:
      return undefined;
  }
}

/**
 * Fold the range markers in a sequence of blocks into the nearest block that
 * can hold them — used by the reader for containers whose model holds no
 * markers of its own (a table cell, a header or footer, a note, a comment, an
 * SDT), and by writers of formats that only allow them inside a paragraph — so
 * they delimit the same text:
 *
 * - an end met while nothing is pending closes after the previous block, so it
 *   is appended to it;
 * - a start, and everything after it, opens before the next block — prepended
 *   to the next paragraph, or put before the first row of a next table;
 * - whatever is still pending when the container ends goes after its last block.
 *
 * A run holds markers before it. Markers left with no block able to hold them
 * — at the end of content made of runs, or in empty content — go to
 * `unhosted` when given, and are dropped otherwise.
 */
export function foldRangeMarkers<T extends object>(
  items: readonly T[],
  unhosted?: RangeAnchor[]
): Exclude<T, RangeAnchor>[] {
  type Block = Exclude<T, RangeAnchor>;
  const out: Block[] = [];
  let pending: RangeAnchor[] = [];
  const hostAtEnd = (markers: readonly RangeAnchor[]): boolean => {
    for (let i = out.length - 1; i >= 0; i--) {
      const hosted = hostRangeMarkers(out[i]!, markers, true);
      if (hosted) {
        out[i] = hosted;
        return true;
      }
    }
    return false;
  };
  for (const item of items) {
    if (isRangeAnchor(item)) {
      if (pending.length === 0 && isClosingAnchor(item) && hostAtEnd([item])) {
        continue;
      }
      pending.push(item);
      continue;
    }
    let block = item as Block;
    if (pending.length > 0) {
      const hosted = hostRangeMarkers(block, pending, false);
      if (hosted) {
        block = hosted;
        pending = [];
      }
    }
    out.push(block);
  }
  if (pending.length > 0 && !hostAtEnd(pending)) {
    unhosted?.push(...pending);
  }
  return out;
}

/**
 * `table` with the range markers between its rows and cells moved into the
 * cells, by the rule of {@link foldRangeMarkers}: an end with nothing pending
 * closes at the end of the previous cell, and a start (with whatever follows
 * it) opens at the start of the next cell. A cell merged into the one above
 * (`verticalMerge: "continue"`) has no content of its own in most formats, so
 * it never receives one. Nested tables are left as they are.
 */
export function foldTableRangeMarkers(table: Table): Table {
  const hasMarkers =
    (table.rangeMarkersAfter?.length ?? 0) > 0 ||
    table.rows.some(
      row =>
        (row.rangeMarkersBefore?.length ?? 0) > 0 ||
        (row.rangeMarkersAfter?.length ?? 0) > 0 ||
        row.cells.some(cell => (cell.rangeMarkersBefore?.length ?? 0) > 0)
    );
  if (!hasMarkers) {
    return table;
  }
  const leading = table.rows.map(row => row.cells.map((): RangeMarker[] => []));
  const trailing = table.rows.map(row => row.cells.map((): RangeMarker[] => []));
  let pending: RangeMarker[] = [];
  let last: { r: number; c: number } | undefined;
  const take = (markers: readonly RangeMarker[] | undefined): void => {
    for (const marker of markers ?? []) {
      if (pending.length === 0 && marker.type.endsWith("End") && last) {
        trailing[last.r]![last.c]!.push(marker);
      } else {
        pending.push(marker);
      }
    }
  };
  table.rows.forEach((row, r) => {
    take(row.rangeMarkersBefore);
    row.cells.forEach((cell, c) => {
      take(cell.rangeMarkersBefore);
      if (cell.properties?.verticalMerge !== "continue") {
        leading[r]![c]!.push(...pending);
        pending = [];
        last = { r, c };
      }
    });
    take(row.rangeMarkersAfter);
  });
  take(table.rangeMarkersAfter);
  if (last) {
    trailing[last.r]![last.c]!.push(...pending);
  }
  const { rangeMarkersAfter: _after, ...rest } = table;
  return {
    ...rest,
    rows: table.rows.map((row, r) => {
      const { rangeMarkersBefore: _before, rangeMarkersAfter: _rowAfter, ...rowRest } = row;
      return {
        ...rowRest,
        cells: row.cells.map((cell, c) => {
          const { rangeMarkersBefore: _cellBefore, ...cellRest } = cell;
          const before = leading[r]![c]!;
          const after = trailing[r]![c]!;
          if (before.length === 0 && after.length === 0) {
            return cellRest;
          }
          return {
            ...cellRest,
            content: foldRangeMarkers([...before, ...cell.content, ...after])
          };
        })
      };
    })
  };
}

/** `block` with `markers` added in front of its own {@link BlockRangeMarkers.rangeMarkersBefore}. */
export function withRangeMarkersBefore<T extends BlockRangeMarkers>(
  block: T,
  markers: readonly RangeMarker[]
): T {
  if (markers.length === 0) {
    return block;
  }
  return { ...block, rangeMarkersBefore: [...markers, ...(block.rangeMarkersBefore ?? [])] };
}

/**
 * The blocks with their {@link BlockRangeMarkers.rangeMarkersBefore} spelled
 * out as items of their own (and the blocks stripped of the field), followed
 * by `after` — the sequence the source had. For a consumer that handles
 * markers by position, such as {@link foldRangeMarkers}.
 */
export function spellOutRangeMarkers<T extends BlockRangeMarkers>(
  blocks: readonly T[],
  after?: readonly RangeMarker[]
): (T | RangeMarker)[] {
  const out: (T | RangeMarker)[] = [];
  for (const block of blocks) {
    if (block.rangeMarkersBefore && block.rangeMarkersBefore.length > 0) {
      const { rangeMarkersBefore, ...rest } = block;
      out.push(...rangeMarkersBefore, rest as T);
    } else {
      out.push(block);
    }
  }
  out.push(...(after ?? []));
  return out;
}

/**
 * Whether a range marker kind is a {@link ParagraphChild} of its own — all
 * but range permissions and span markers ({@link isSpanMarker}).
 */
export function isParagraphChildMarker(
  marker: RangeMarker
): marker is Exclude<RangeMarker, PermStart | PermEnd | SpanMarker> {
  return marker.type !== "permStart" && marker.type !== "permEnd" && !isSpanMarker(marker);
}

/**
 * Paragraph content from children and range markers in source order. A marker
 * that is a paragraph child stays one; any other (a range permission, a span
 * marker) rides on
 * the child that follows it ({@link ParagraphChildRangeMarkers}), or — after
 * the last child — in `rangeMarkersAfter`.
 */
export function arrangeParagraphChildren(items: readonly (ParagraphChild | RangeMarker)[]): {
  children: ParagraphChild[];
  rangeMarkersAfter?: RangeMarker[];
} {
  // Every paragraph read goes through here, and nearly none has a range
  // permission: then the items already are the children.
  if (!items.some(isRidingMarker)) {
    return { children: items.slice() as ParagraphChild[] };
  }
  const children: ParagraphChild[] = [];
  let pending: RangeMarker[] = [];
  for (const item of items) {
    if (isRangeMarker(item) && !isParagraphChildMarker(item)) {
      pending.push(item);
      continue;
    }
    children.push(withRangeMarkersBefore(item as ParagraphChild, pending));
    pending = [];
  }
  return pending.length > 0 ? { children, rangeMarkersAfter: pending } : { children };
}

/**
 * Every range marker inside `content` (paragraphs, tables, SDTs and runs, at
 * any depth) and in `after`, in source order, split at the first run: the
 * markers met before any run are `leading`, the rest `trailing`. For code that
 * replaces content wholesale and must keep the ranges it delimited paired.
 *
 * Span markers ({@link isSpanMarker}) are left out unless `spans` is set: they
 * wrap the content being replaced, which has no counterpart elsewhere.
 */
export function collectRangeMarkers(
  content: readonly (BodyContent | Run)[],
  after?: readonly RangeMarker[],
  options?: { readonly spans?: boolean }
): { leading: RangeMarker[]; trailing: RangeMarker[] } {
  return collectAnchors(content, after, false, options?.spans ?? false) as {
    leading: RangeMarker[];
    trailing: RangeMarker[];
  };
}

/**
 * {@link collectRangeMarkers}, with the comment references among the content
 * too ({@link RangeAnchor}): what code that removes content must keep so no
 * range and no comment loses its anchor.
 */
export function collectRangeAnchors(
  content: readonly (BodyContent | Run)[],
  after?: readonly RangeMarker[]
): { leading: RangeAnchor[]; trailing: RangeAnchor[] } {
  return collectAnchors(content, after, true, false);
}

/** Every {@link RangeAnchor} inside `content` and in `after`, in source order. */
export function listRangeAnchors(
  content: readonly (BodyContent | Run)[],
  after?: readonly RangeMarker[]
): RangeAnchor[] {
  const { leading, trailing } = collectRangeAnchors(content, after);
  return [...leading, ...trailing];
}

function collectAnchors(
  content: readonly (BodyContent | Run)[],
  after: readonly RangeMarker[] | undefined,
  references: boolean,
  spans: boolean
): { leading: RangeAnchor[]; trailing: RangeAnchor[] } {
  const leading: RangeAnchor[] = [];
  const trailing: RangeAnchor[] = [];
  let seenRun = false;
  const holder = {
    type: "sdt",
    content: content as StructuredDocumentTag["content"],
    contentRangeMarkersAfter: after
  } as StructuredDocumentTag;
  const take = (anchor: RangeAnchor): void => {
    (seenRun ? trailing : leading).push(anchor);
  };
  walkBlocks([holder], {
    visitRangeMarker(marker) {
      if (spans || !isSpanMarker(marker)) {
        take(marker);
      }
    },
    ...(references ? { visitCommentReference: take } : {}),
    enterRun() {
      seenRun = true;
    }
  });
  return { leading, trailing };
}

/**
 * `para` with its children replaced by `children`. The range markers among
 * and on the old children, and its comment references, keep their source
 * order: those before the first run go in front of the new children, the rest
 * after them, so every range the paragraph opened or closed stays paired and
 * every comment it referenced keeps its reference. The paragraph's own
 * {@link BlockRangeMarkers.rangeMarkersBefore} is kept as it is.
 */
export function replaceParagraphChildren(
  para: Paragraph,
  children: readonly ParagraphChild[]
): Paragraph {
  const { rangeMarkersBefore: _before, rangeMarkersAfter: _after, ...rest } = para;
  const { leading, trailing } = collectRangeAnchors([{ ...rest, rangeMarkersAfter: _after }]);
  const arranged = arrangeParagraphChildren([...leading, ...children, ...trailing]);
  return {
    ...rest,
    ...(_before ? { rangeMarkersBefore: _before } : {}),
    children: arranged.children,
    ...(arranged.rangeMarkersAfter ? { rangeMarkersAfter: arranged.rangeMarkersAfter } : {})
  };
}

/** Every range marker inside `content` and in `after`, in source order (see {@link collectRangeMarkers}). */
export function listRangeMarkers(
  content: readonly (BodyContent | Run)[],
  after?: readonly RangeMarker[]
): RangeMarker[] {
  const { leading, trailing } = collectRangeMarkers(content, after);
  return [...leading, ...trailing];
}

/**
 * Blocks interleaved with range markers — as left by code that removed blocks
 * and kept their markers in place — with the markers folded into the blocks
 * around them ({@link foldRangeMarkers}). Markers no block can hold end up in
 * a paragraph of their own at the end, so no range loses an end.
 */
export function settleRangeMarkers<T extends object>(
  items: readonly (T | RangeAnchor)[]
): Exclude<T, RangeAnchor>[] {
  if (!items.some(item => isRangeAnchor(item))) {
    return items as Exclude<T, RangeAnchor>[];
  }
  const unhosted: RangeAnchor[] = [];
  const out = foldRangeMarkers(items, unhosted);
  if (unhosted.length > 0) {
    const arranged = arrangeParagraphChildren(unhosted);
    out.push({ type: "paragraph", ...arranged } as unknown as Exclude<T, RangeAnchor>);
  }
  return out;
}

/**
 * `para` with the children `items` — children interleaved with range markers,
 * as left by code that removed children and kept their markers in place —
 * re-arranged so each marker is a child or rides on the child after it
 * ({@link arrangeParagraphChildren}). Markers after the last child go in front
 * of the paragraph's existing {@link Paragraph.rangeMarkersAfter}.
 */
export function withParagraphChildren(
  para: Paragraph,
  items: readonly (ParagraphChild | RangeMarker)[]
): Paragraph {
  const arranged = arrangeParagraphChildren(items);
  const after = [...(arranged.rangeMarkersAfter ?? []), ...(para.rangeMarkersAfter ?? [])];
  const { rangeMarkersAfter: _after, ...rest } = para;
  return {
    ...rest,
    children: arranged.children,
    ...(after.length > 0 ? { rangeMarkersAfter: after } : {})
  };
}

/** Every range marker on and inside a paragraph child, in source order. */
export function listChildRangeMarkers(child: ParagraphChild): RangeMarker[] {
  // The child's siblings stay, so a span it opened or closed still has its
  // other end in the paragraph: keep span markers too.
  const { leading, trailing } = collectRangeMarkers(
    [{ type: "paragraph", children: [child] }],
    undefined,
    { spans: true }
  );
  return [...leading, ...trailing];
}

/**
 * Remove `blocks[index]` in place. The range markers before it sit between
 * blocks and belong to ranges around it, so they move onto the next block;
 * when it was the last block they are returned as `markersAfter`, for the
 * container's list of markers after its last block. `removed` is the removed
 * block itself, with its `rangeMarkersBefore` deleted: those markers now live
 * in the container, and leaving them on the block would duplicate them if it
 * were inserted again.
 */
export function removeBlock<T extends BlockRangeMarkers>(
  blocks: T[],
  index: number
): { removed: T; markersAfter: RangeMarker[] } {
  const [block] = blocks.splice(index, 1) as [T];
  const rangeMarkersBefore = block.rangeMarkersBefore;
  if (!rangeMarkersBefore || rangeMarkersBefore.length === 0) {
    return { removed: block, markersAfter: [] };
  }
  delete (block as { rangeMarkersBefore?: unknown }).rangeMarkersBefore;
  if (index < blocks.length) {
    blocks[index] = withRangeMarkersBefore(blocks[index]!, rangeMarkersBefore);
    return { removed: block, markersAfter: [] };
  }
  return { removed: block, markersAfter: [...rangeMarkersBefore] };
}

/** The range kind of each marker type: a start and its end share it. */
const RANGE_KIND: ReadonlyMap<string, string> = new Map(
  RANGE_MARKER_TYPES.map(type => [type, type.replace(/(Range)?(Start|End)$/, "")])
);

/** The marker types {@link isParagraphChildMarker} rejects: they ride on the next child. */
const RIDING_MARKER_TYPES: ReadonlySet<string> = new Set<RangeMarker["type"]>([
  "permStart",
  "permEnd",
  "bidiStart",
  "bidiEnd",
  "smartTagStart",
  "smartTagEnd"
]);

/**
 * Whether `item` is a range marker that is not a paragraph child of its own
 * but rides on the child after it — what {@link arrangeParagraphChildren}
 * moves. A cheap test, for code that wants to skip arranging when none is there.
 */
export function isRidingMarker(item: object): boolean {
  return "type" in item && RIDING_MARKER_TYPES.has((item as { type: string }).type);
}

/** Identifies a range: a start and its end share it. */
function rangeKey(marker: Exclude<RangeMarker, SpanMarker>): string {
  return `${RANGE_KIND.get(marker.type)}:${marker.id}`;
}

/**
 * What {@link keepClonedRangeMarkers} needs to know about repeated content,
 * the same for every copy: the ranges it opens and closes, and where in it
 * the markers and comment references are — so preparing a copy visits only
 * those places rather than the whole copy. `null` when the content holds no
 * range marker and no comment reference: its copies need no preparing.
 */
export interface RepeatedRanges {
  readonly starts: ReadonlySet<string>;
  readonly ends: ReadonlySet<string>;
  readonly plan: MarkerPlan;
}

/**
 * The part of the content's tree that leads to a marker or a comment
 * reference: at each object, the marker lists it holds, whether it is a
 * paragraph with one among its children, and the keys to descend into.
 */
interface MarkerPlan {
  readonly lists: readonly string[];
  readonly paragraph: boolean;
  readonly next: ReadonlyMap<string, MarkerPlan>;
}

/** {@link RepeatedRanges} of `content`. */
export function repeatedRanges(content: unknown): RepeatedRanges | null {
  const starts = new Set<string>();
  const ends = new Set<string>();
  const planFor = (node: unknown): MarkerPlan | undefined => {
    if (!node || typeof node !== "object" || ArrayBuffer.isView(node)) {
      return undefined;
    }
    let found = false;
    if (isRangeMarker(node)) {
      // A span marker carries no id: every copy keeps its own span.
      if (!isSpanMarker(node)) {
        (node.type.endsWith("End") ? ends : starts).add(rangeKey(node));
      }
      found = true;
    } else if ((node as { type?: unknown }).type === "commentReference") {
      found = true;
    }
    const o = node as Record<string, unknown>;
    const lists: string[] = [];
    const next = new Map<string, MarkerPlan>();
    for (const key in o) {
      const sub = planFor(o[key]);
      if (sub) {
        next.set(key, sub);
      }
      if (MARKER_LIST_FIELDS.has(key) && Array.isArray(o[key])) {
        lists.push(key);
      }
    }
    const paragraph =
      o.type === "paragraph" &&
      Array.isArray(o.children) &&
      (o.children as object[]).some(c => isRangeAnchor(c));
    return found || lists.length > 0 || paragraph || next.size > 0
      ? { lists, paragraph, next }
      : undefined;
  };
  const plan = planFor(content);
  return plan ? { starts, ends, plan } : null;
}

/** Fields that hold a list of range markers outside a paragraph's children. */
const MARKER_LIST_FIELDS: ReadonlySet<string> = new Set([
  "rangeMarkersBefore",
  "rangeMarkersAfter",
  "contentRangeMarkersAfter",
  "bodyRangeMarkersAfter"
]);

/**
 * Prepare copy `index` of `count` of repeated content (a template loop, a
 * repeating section) so the copies together hold each range once, with every
 * pair intact. Mutates `value`, which must be a fresh clone, and returns it.
 *
 * - A range that crosses the boundary of the content — one of its ends lies
 *   outside it — covered all of the repeated content, so it still does: a
 *   start inside is kept on the first copy and an end inside on the last.
 * - A permission range is an editable region, and the region the author
 *   marked covered the repeated content: its start is kept on the first copy
 *   and its end on the last, so one region spans every copy.
 * - Every other range lying wholly inside is kept on the first copy only. A
 *   bookmark name must be unique and a `REF`, `PAGEREF` or internal hyperlink
 *   has one target, which is the first item; a comment range and its
 *   reference belong to one comment; a move or custom-XML revision range
 *   records one edit.
 *
 * A marker is kept on exactly one copy, so every range keeps each of its ends
 * once. A comment reference goes with its comment's end when that end is the
 * only one inside, and with the first copy otherwise. Permission markers
 * riding on a removed child move to the next child.
 *
 * `ranges` describes the repeated content — the same for every copy — so a
 * caller making many copies computes it once, from the template the copies
 * are cloned from, with {@link repeatedRanges}; it is computed from `value`
 * when omitted.
 */
export function keepClonedRangeMarkers<T>(
  value: T,
  index: number,
  count: number,
  ranges: RepeatedRanges | null = repeatedRanges(value)
): T {
  if (count <= 1 || ranges === null) {
    return value;
  }
  const { starts, ends } = ranges;
  const keep = (marker: RangeMarker): boolean => {
    if (isSpanMarker(marker)) {
      // A span wraps content of its own copy: every copy keeps its span.
      return true;
    }
    const key = rangeKey(marker);
    const isEnd = marker.type.endsWith("End");
    const paired = starts.has(key) && ends.has(key);
    return isEnd && (marker.type === "permEnd" || !paired) ? index === count - 1 : index === 0;
  };
  const keepReference = (id: number): boolean => {
    const key = `comment:${id}`;
    return ends.has(key) && !starts.has(key) ? index === count - 1 : index === 0;
  };
  const apply = (node: unknown, plan: MarkerPlan): void => {
    if (!node || typeof node !== "object") {
      return;
    }
    if (Array.isArray(node)) {
      for (const [key, sub] of plan.next) {
        apply(node[Number(key)], sub);
      }
      return;
    }
    const o = node as Record<string, unknown>;
    for (const field of plan.lists) {
      const kept = (o[field] as RangeMarker[]).filter(keep);
      if (kept.length > 0) {
        o[field] = kept;
      } else {
        delete o[field];
      }
    }
    const field = o.type === "field" && Array.isArray(o.rangeMarkers);
    if (field) {
      // Markers inside a complex field (`FieldContent.rangeMarkers`).
      const kept = (o.rangeMarkers as FieldRangeMarker[]).filter(entry => keep(entry.marker));
      if (kept.length > 0) {
        o.rangeMarkers = kept;
      } else {
        delete o.rangeMarkers;
      }
    }
    const childPlans = plan.next.get("children");
    if (plan.paragraph) {
      const children = o.children as ParagraphChild[];
      let out: ParagraphChild[] | undefined;
      let pending: RangeMarker[] = [];
      for (let i = 0; i < children.length; i++) {
        const child = children[i]!;
        const childPlan = childPlans?.next.get(String(i));
        if (childPlan) {
          apply(child, childPlan);
        }
        const removed = isRangeMarker(child)
          ? !keep(child)
          : "type" in child && child.type === "commentReference" && !keepReference(child.id);
        if (removed) {
          out ??= children.slice(0, i);
          if (child.rangeMarkersBefore) {
            pending.push(...child.rangeMarkersBefore);
          }
          continue;
        }
        if (out) {
          out.push(withRangeMarkersBefore(child, pending));
          pending = [];
        }
      }
      if (pending.length > 0) {
        o.rangeMarkersAfter = [...pending, ...((o.rangeMarkersAfter as RangeMarker[]) ?? [])];
      }
      if (out) {
        o.children = out;
      }
    }
    for (const [key, sub] of plan.next) {
      if (
        !MARKER_LIST_FIELDS.has(key) &&
        !(plan.paragraph && key === "children") &&
        !(field && key === "rangeMarkers")
      ) {
        apply(o[key], sub);
      }
    }
  };
  apply(value, ranges.plan);
  return value;
}
