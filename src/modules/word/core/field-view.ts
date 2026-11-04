/**
 * Fields as they read.
 *
 * The model holds a complex field in one of two shapes (see
 * `FieldContent.fieldChar`): whole — one {@link FieldContent} whose result is
 * its `cachedValue` — or in pieces: its begin, separate and end characters each
 * a `FieldContent` of their own, at the places they were, its result the
 * content between the last two, in whichever runs, wrappers and paragraphs
 * that lies. Code that acts on fields — mail merge, field update, the template
 * engine, the converters and the page layout — reads the second shape through
 * this module, so a field in pieces is the one field it is to every consumer:
 *
 * - {@link resolveSplitFields} / {@link forEachField}: the fields a list of
 *   blocks holds, with the instruction and the result they read as;
 * - {@link rewriteSplitFields}: replace a field, or its result, across the
 *   runs and paragraphs it spans, keeping its field characters;
 * - {@link hideFieldCode} / {@link withDisplayFields}: what a document shows —
 *   no content of a field's instruction, page-number fields whole so a layout
 *   can compute them.
 *
 * A field is resolved within one list of blocks (a body, a table cell, a
 * header…): the reader carries a field across paragraphs, not out of the
 * container it began in.
 */

import type { Mutable } from "@word/core/internal-utils";
import { isRangeAnchor, withParagraphChildren } from "@word/core/range-markers";
import {
  extractRunText,
  hyperlinkContent,
  isRangeMarker,
  isRun,
  withHyperlinkContent
} from "@word/core/text-utils";
import { parsePageField } from "@word/query/page-field";
import type {
  BodyContent,
  DeletedRun,
  DocxDocument,
  FieldContent,
  Hyperlink,
  InsertedRun,
  MovedFromRun,
  MovedToRun,
  Paragraph,
  ParagraphChild,
  RangeMarker,
  Run,
  RunContent,
  RunProperties
} from "@word/types";

// =============================================================================
// Types
// =============================================================================

/** A field held in pieces, as it reads. */
export interface SplitField {
  /**
   * Its whole instruction: the instruction text of its pieces, with what the
   * fields nested in it show in between — what Word evaluates.
   */
  readonly instruction: string;
  /**
   * Its instruction as Word's field codes view shows it: the text of its
   * pieces, and each field nested there — whole or in pieces — as itself
   * rather than as what it shows. {@link fieldCodeText} spells it out.
   */
  readonly code: readonly (string | FieldContent | SplitField)[];
  readonly begin: FieldContent;
  readonly separate: FieldContent | undefined;
  /** Its end piece; `undefined` when the list of blocks does not close it. */
  readonly end: FieldContent | undefined;
  /** The text of its result (a paragraph break as `"\n"`). */
  readonly result: string;
  /** The field in pieces it is nested in, if any. */
  readonly parent: SplitField | undefined;
  /** Whether it lies in its parent's instruction rather than its result. */
  readonly inCode: boolean;
  /** Whether its begin is in a tracked deletion (`w:del`, `w:moveFrom`). */
  readonly deleted: boolean;
  /** Index, in the list of blocks, of the paragraph holding its begin. */
  readonly block: number;
  /** Index of the paragraph holding its end, when the list closes it. */
  readonly endBlock: number | undefined;
}

/** A field met by {@link forEachField}. */
export type FieldRef =
  | {
      readonly kind: "whole";
      readonly content: FieldContent;
      readonly instruction: string;
      readonly deleted: boolean;
    }
  | {
      readonly kind: "split";
      readonly field: SplitField;
      readonly instruction: string;
      readonly deleted: boolean;
    };

/**
 * What to put in place of a field in pieces, or of its result. The text (or
 * `content`) becomes one run, with the formatting of the result's first run;
 * `paragraphs` replace a result that is paragraphs — a TOC's entries — the
 * first opening in the paragraph of the separate character and the end
 * closing the last (or keeping the paragraph of its own it had).
 */
export type SplitFieldEdit =
  | {
      readonly replace: "field" | "result";
      readonly text: string;
      /** Content of the replacing run instead of the text. */
      readonly content?: readonly RunContent[];
    }
  | { readonly replace: "result"; readonly paragraphs: readonly Paragraph[] };

type RevisionWrapper = InsertedRun | DeletedRun | MovedFromRun | MovedToRun;
type Wrapper = Hyperlink | RevisionWrapper;

/** One leaf of paragraph content, with the wrappers it sits in (outermost first). */
interface Atom {
  readonly node: ParagraphChild | RangeMarker;
  readonly chain: readonly Wrapper[];
}

interface Pos {
  readonly block: number;
  readonly atom: number;
}

interface Places {
  readonly begin: Pos;
  separate?: Pos;
  end?: Pos;
  /** Whether a block that is not a paragraph lies inside it. */
  barrier: boolean;
}

/** A list of blocks with its fields in pieces resolved. */
export interface ResolvedFields {
  /** The fields in pieces, in the order of their begins. */
  readonly fields: readonly SplitField[];
  /** @internal */
  readonly atoms: readonly (Atom[] | undefined)[];
  /** @internal */
  readonly places: ReadonlyMap<SplitField, Places>;
  /** @internal Positions of the runs inside a field's instruction. */
  readonly code: readonly Pos[];
}

// =============================================================================
// Leaves
// =============================================================================

const REVISIONS: ReadonlySet<string> = new Set([
  "insertedRun",
  "deletedRun",
  "movedFromRun",
  "movedToRun"
]);

function isDeletion(w: Wrapper): boolean {
  return w.type === "deletedRun" || w.type === "movedFromRun";
}

function isParagraph(block: object): block is Paragraph {
  return "type" in block && (block as { type: unknown }).type === "paragraph";
}

/** The field piece a run stands for, when it holds just one. */
function pieceOf(run: Run): FieldContent | undefined {
  const c = run.content.length === 1 ? run.content[0] : undefined;
  return c?.type === "field" && c.fieldChar ? c : undefined;
}

function hasPieceContent(run: Run): boolean {
  for (const c of run.content) {
    if (c.type === "field" && c.fieldChar) {
      return true;
    }
  }
  return false;
}

/** Whether any run among `children` holds a field piece. */
function childrenHoldPiece(children: readonly ParagraphChild[]): boolean {
  for (const child of children) {
    if (isRun(child)) {
      if (hasPieceContent(child)) {
        return true;
      }
    } else if ((child as { type: string }).type === "hyperlink") {
      const link = child as Hyperlink;
      if (
        childrenHoldPiece(link.children) ||
        (link.nestedChildren?.some(n => childrenHoldPiece([n.child])) ?? false)
      ) {
        return true;
      }
    } else if (REVISIONS.has((child as { type: string }).type)) {
      if (hasPieceContent((child as RevisionWrapper).run)) {
        return true;
      }
    }
  }
  return false;
}

/** Whether `paragraph` holds a field piece anywhere in its content. */
export function holdsFieldPiece(paragraph: Paragraph): boolean {
  return childrenHoldPiece(paragraph.children);
}

function withoutMarkersBefore<T extends { readonly rangeMarkersBefore?: unknown }>(node: T): T {
  if (node.rangeMarkersBefore === undefined) {
    return node;
  }
  const { rangeMarkersBefore: _before, ...rest } = node;
  return rest as T;
}

function pushMarkers(
  markers: readonly RangeMarker[] | undefined,
  chain: readonly Wrapper[],
  out: Atom[]
): void {
  if (markers) {
    for (const marker of markers) {
      out.push({ node: marker, chain });
    }
  }
}

/** A run as leaves: each field piece among other content becomes a run of its own. */
function pushRun(run: Run, chain: readonly Wrapper[], out: Atom[]): void {
  if (run.content.length <= 1 || !hasPieceContent(run)) {
    out.push({ node: run, chain });
    return;
  }
  let group: RunContent[] = [];
  const flush = (): void => {
    if (group.length > 0) {
      out.push({ node: { ...run, content: group }, chain });
      group = [];
    }
  };
  for (const c of run.content) {
    if (c.type === "field" && c.fieldChar) {
      flush();
      out.push({ node: { ...run, content: [c] }, chain });
    } else {
      group.push(c);
    }
  }
  flush();
}

function pushChildren(
  children: readonly ParagraphChild[],
  chain: readonly Wrapper[],
  out: Atom[]
): void {
  for (const child of children) {
    pushMarkers(child.rangeMarkersBefore, chain, out);
    if (isRun(child)) {
      pushRun(withoutMarkersBefore(child), chain, out);
      continue;
    }
    const type = (child as { type: string }).type;
    if (type === "hyperlink") {
      const link = child as Hyperlink;
      const inner = [...chain, link];
      pushChildren(hyperlinkContent(link), inner, out);
      pushMarkers(link.rangeMarkersAfter, inner, out);
    } else if (REVISIONS.has(type)) {
      const wrapper = child as RevisionWrapper;
      const inner = [...chain, wrapper];
      pushMarkers(wrapper.run.rangeMarkersBefore, inner, out);
      pushRun(withoutMarkersBefore(wrapper.run), inner, out);
    } else {
      out.push({ node: withoutMarkersBefore(child), chain });
    }
  }
}

/** A paragraph's content as leaves, in document order. */
function atomize(paragraph: Paragraph): Atom[] {
  const out: Atom[] = [];
  pushChildren(paragraph.children, [], out);
  pushMarkers(paragraph.rangeMarkersAfter, [], out);
  return out;
}

/** Leaves back into paragraph content (the inverse of {@link atomize}). */
function rebuildItems(atoms: readonly Atom[], depth: number): (ParagraphChild | RangeMarker)[] {
  const items: (ParagraphChild | RangeMarker)[] = [];
  for (let i = 0; i < atoms.length;) {
    const wrapper = atoms[i]!.chain[depth];
    if (!wrapper) {
      items.push(atoms[i]!.node);
      i++;
      continue;
    }
    let j = i + 1;
    while (j < atoms.length && atoms[j]!.chain[depth] === wrapper) {
      j++;
    }
    const inner = rebuildItems(atoms.slice(i, j), depth + 1);
    if (wrapper.type === "hyperlink") {
      const {
        rangeMarkersBefore: _before,
        rangeMarkersAfter: _after,
        nestedChildren: _nested,
        ...link
      } = wrapper;
      items.push(withHyperlinkContent({ ...link, children: [] }, inner));
    } else {
      const bare = withoutMarkersBefore(wrapper);
      for (const item of inner) {
        items.push(
          isRangeMarker(item) || !isRun(item) ? item : ({ ...bare, run: item } as ParagraphChild)
        );
      }
    }
    i = j;
  }
  return items;
}

function paragraphFrom(base: Paragraph, atoms: readonly Atom[]): Paragraph {
  const { rangeMarkersAfter: _after, ...rest } = base;
  return withParagraphChildren({ ...rest, children: [] }, rebuildItems(atoms, 0));
}

/** The range markers and comment references in removed content, to keep. */
function keptAnchors(atom: Atom, out: Atom[]): void {
  const node = atom.node;
  if (isRangeAnchor(node)) {
    out.push({ node, chain: [] });
    return;
  }
  if (isRun(node)) {
    for (const c of node.content) {
      if (c.type === "field" && c.rangeMarkers) {
        for (const m of c.rangeMarkers) {
          out.push({ node: m.marker, chain: [] });
        }
      }
    }
  }
}

function runText(run: Run): string {
  return extractRunText(run);
}

// =============================================================================
// Resolution
// =============================================================================

/**
 * The fields in pieces `blocks` hold, or `undefined` when no paragraph among
 * them holds a piece (the common case, found without allocating).
 */
export function resolveSplitFields(blocks: readonly object[]): ResolvedFields | undefined {
  let any = false;
  for (const block of blocks) {
    if (isParagraph(block) && holdsFieldPiece(block)) {
      any = true;
      break;
    }
  }
  if (!any) {
    return undefined;
  }
  const atoms: (Atom[] | undefined)[] = [];
  const fields: Mutable<SplitField>[] = [];
  const places = new Map<SplitField, Places>();
  const code: Pos[] = [];
  const stack: Mutable<SplitField>[] = [];

  const attribute = (text: string, code: string | FieldContent = text): void => {
    const top = stack[stack.length - 1];
    if (top && !top.separate) {
      (top.code as (string | FieldContent | SplitField)[]).push(code);
    }
    for (let i = stack.length - 1; i >= 0; i--) {
      const f = stack[i]!;
      if (f.separate) {
        f.result += text;
      } else {
        f.instruction += text;
        return;
      }
    }
  };
  const inCode = (): boolean => stack.some(f => !f.separate);

  for (let b = 0; b < blocks.length; b++) {
    const block = blocks[b]!;
    if (!isParagraph(block)) {
      for (const f of stack) {
        places.get(f)!.barrier = true;
      }
      continue;
    }
    if (stack.length === 0 && !holdsFieldPiece(block)) {
      continue;
    }
    const list = atomize(block);
    atoms[b] = list;
    for (let a = 0; a < list.length; a++) {
      const atom = list[a]!;
      const node = atom.node;
      if (isRangeMarker(node) || !isRun(node)) {
        continue;
      }
      const deleted = atom.chain.some(isDeletion);
      const piece = pieceOf(node);
      if (!piece) {
        if (stack.length === 0) {
          continue;
        }
        if (inCode()) {
          code.push({ block: b, atom: a });
        }
        if (!deleted) {
          for (const c of node.content) {
            if (c.type === "field" && !c.fieldChar) {
              attribute(c.cachedValue ?? "", c);
            } else {
              attribute(runText({ ...node, content: [c] }));
            }
          }
        }
        continue;
      }
      const top = stack[stack.length - 1];
      if (piece.fieldChar === "begin") {
        const field: Mutable<SplitField> = {
          instruction: piece.instruction,
          code: [piece.instruction],
          begin: piece,
          separate: undefined,
          end: undefined,
          result: "",
          parent: top,
          inCode: top !== undefined && !top.separate,
          deleted,
          block: b,
          endBlock: undefined
        };
        places.set(field, { begin: { block: b, atom: a }, barrier: false });
        if (top && !top.separate && !deleted) {
          (top.code as (string | FieldContent | SplitField)[]).push(field);
        }
        fields.push(field);
        stack.push(field);
      } else if (!top) {
        // A separate or end character no field opened.
        continue;
      } else if (piece.fieldChar === "separate") {
        if (!top.separate) {
          top.instruction += piece.instruction;
          (top.code as string[]).push(piece.instruction);
          top.separate = piece;
          places.get(top)!.separate = { block: b, atom: a };
        }
      } else {
        if (!top.separate) {
          top.instruction += piece.instruction;
          (top.code as string[]).push(piece.instruction);
        }
        top.end = piece;
        top.endBlock = b;
        places.get(top)!.end = { block: b, atom: a };
        stack.pop();
      }
    }
    if (stack.length > 0) {
      attribute("\n");
    }
  }
  return { fields, atoms, places, code };
}

/**
 * `field`'s instruction with each field nested in it drawn by `nested`:
 * `{ MERGEFIELD Vip }` for the field codes view, what it shows (or will show)
 * to evaluate the field as Word does.
 */
export function fieldCodeText(
  field: SplitField,
  nested: (inner: FieldContent | SplitField) => string
): string {
  let out = "";
  for (const part of field.code) {
    out += typeof part === "string" ? part : nested(part);
  }
  return out;
}

/**
 * `field`'s instruction as the field codes view shows it: each nested field
 * as `{ its code }`, never by its result.
 */
export function fieldCodeView(field: SplitField): string {
  return fieldCodeText(
    field,
    inner => `{ ${("code" in inner ? fieldCodeView(inner) : inner.instruction).trim()} }`
  );
}

/**
 * What each field in pieces of `resolved` computes to, inner fields first as
 * Word updates them: `evaluate` gets a field's instruction with each field
 * nested in it showing its new result — a field in pieces the one computed
 * here (or the result it has), a whole field `whole(content)` — and returns
 * the new result, or `undefined` to keep it.
 */
export function evaluateSplitFields(
  resolved: ResolvedFields,
  evaluate: (field: SplitField, instruction: string) => string | undefined,
  whole: (content: FieldContent, within: SplitField) => string
): Map<SplitField, string> {
  const values = new Map<SplitField, string>();
  // A nested field begins after the field it is in: backwards, it comes first.
  for (let i = resolved.fields.length - 1; i >= 0; i--) {
    const field = resolved.fields[i]!;
    const instruction = fieldCodeText(field, inner =>
      "code" in inner ? (values.get(inner) ?? inner.result) : whole(inner, field)
    );
    const value = evaluate(field, instruction);
    if (value !== undefined) {
      values.set(field, value);
    }
  }
  return values;
}

/**
 * Every field `blocks` hold — whole or in pieces — in document order (a field
 * in pieces at its begin), in nested lists of blocks too (table cells,
 * content controls, text boxes, shapes, TOC entries), as `walkBlocks` visits them.
 */
export function forEachField(blocks: readonly object[], visit: (field: FieldRef) => void): void {
  const resolved = resolveSplitFields(blocks);
  const begins = new Map<FieldContent, SplitField>();
  for (const f of resolved?.fields ?? []) {
    begins.set(f.begin, f);
  }
  const content = (run: Run, deleted: boolean): void => {
    for (const c of run.content) {
      if (c.type !== "field") {
        continue;
      }
      if (!c.fieldChar) {
        visit({ kind: "whole", content: c, instruction: c.instruction, deleted });
        continue;
      }
      const f = c.fieldChar === "begin" ? begins.get(c) : undefined;
      if (f) {
        visit({ kind: "split", field: f, instruction: f.instruction, deleted: f.deleted });
      }
    }
  };
  const children = (list: readonly ParagraphChild[], deleted: boolean): void => {
    for (const child of list) {
      if (isRun(child)) {
        content(child, deleted);
      } else if ((child as { type: string }).type === "hyperlink") {
        children(hyperlinkContent(child as Hyperlink), deleted);
      } else if (REVISIONS.has((child as { type: string }).type)) {
        const w = child as RevisionWrapper;
        content(w.run, deleted || isDeletion(w));
      }
    }
  };
  for (const block of blocks) {
    if (isParagraph(block)) {
      children(block.children, false);
    } else {
      for (const list of nestedLists(block)) {
        forEachField(list, visit);
      }
    }
  }
}

/** The lists of blocks nested directly in `block`. */
function nestedLists(block: object): (readonly object[])[] {
  const b = block as BodyContent;
  switch (b.type) {
    case "table":
      return b.rows.flatMap(row => row.cells.map(cell => cell.content));
    case "sdt":
      return [b.content];
    case "textBox":
      return b.content ? [b.content] : [];
    case "drawingShape":
      return b.textContent ? [b.textContent] : [];
    case "tableOfContents":
      return b.cachedParagraphs ? [b.cachedParagraphs] : [];
    default:
      return [];
  }
}

// =============================================================================
// Rewriting
// =============================================================================

function compare(a: Pos, b: Pos): number {
  return a.block - b.block || a.atom - b.atom;
}

interface PlannedEdit {
  readonly field: SplitField;
  readonly edit: SplitFieldEdit;
  readonly places: Places;
  /** First and last position removed (inclusive; `to` before `from` when nothing is). */
  readonly from: Pos;
  readonly to: Pos;
}

function plan(field: SplitField, edit: SplitFieldEdit, places: Places): PlannedEdit | undefined {
  const { end, separate } = places;
  if (!end || places.barrier) {
    return undefined;
  }
  if (edit.replace === "field") {
    return { field, edit, places, from: places.begin, to: end };
  }
  if (!separate) {
    if ("paragraphs" in edit) {
      return undefined;
    }
    return { field, edit, places, from: end, to: { block: end.block, atom: end.atom - 1 } };
  }
  return {
    field,
    edit,
    places,
    from: { block: separate.block, atom: separate.atom + 1 },
    to: { block: end.block, atom: end.atom - 1 }
  };
}

/** Whether `[a.from, a.to]` and `[b.from, b.to]` share a position. */
function overlaps(a: PlannedEdit, b: PlannedEdit): boolean {
  const aEmpty = compare(a.from, a.to) > 0;
  const bEmpty = compare(b.from, b.to) > 0;
  if (aEmpty || bEmpty) {
    // An insertion point inside the other range.
    const point = aEmpty ? a : b;
    const range = aEmpty ? b : a;
    return compare(point.from, range.from) > 0 && compare(point.from, range.to) <= 0;
  }
  return compare(a.from, b.to) <= 0 && compare(b.from, a.to) <= 0;
}

/** One block of the list being rewritten: its source, and its leaves once touched. */
interface Slot {
  base: object;
  atoms: Atom[] | undefined;
  dirty: boolean;
}

function runAt(slots: readonly Slot[], pos: Pos): Atom | undefined {
  return slots[pos.block]?.atoms?.[pos.atom];
}

/**
 * `blocks` with the fields in pieces `edit` answers for replaced as it says:
 * the whole field, or its result, across the runs, wrappers and paragraphs it
 * spans. Paragraphs a replaced range joins become one — the first, keeping a
 * section break the last carried. The range markers and comment references
 * inside replaced content are kept. A field nested in another's replaced
 * range, one not closed in `blocks`, or one spanning a block that is not a
 * paragraph is left as it is. Returns `blocks` itself when nothing changes.
 */
export function rewriteSplitFields<T extends object>(
  blocks: readonly T[],
  edit: (field: SplitField) => SplitFieldEdit | undefined,
  resolved: ResolvedFields | undefined = resolveSplitFields(blocks)
): readonly T[] {
  if (!resolved) {
    return blocks;
  }
  const kept: PlannedEdit[] = [];
  for (const field of resolved.fields) {
    const e = edit(field);
    const planned = e && plan(field, e, resolved.places.get(field)!);
    if (planned && !kept.some(k => overlaps(k, planned))) {
      kept.push(planned);
    }
  }
  if (kept.length === 0) {
    return blocks;
  }
  const slots: Slot[] = blocks.map((base, i) => ({
    base,
    atoms: resolved.atoms[i],
    dirty: false
  }));
  kept.sort((a, b) => compare(b.from, a.from));
  for (const planned of kept) {
    if ("paragraphs" in planned.edit) {
      applyParagraphs(slots, planned, planned.edit.paragraphs);
    } else {
      applyInline(slots, planned, planned.edit);
    }
  }
  return slots.map(slot =>
    slot.dirty && slot.atoms ? paragraphFrom(slot.base as Paragraph, slot.atoms) : slot.base
  ) as T[];
}

/** The leaves from `from` to `to` (inclusive), and the anchors among them to keep. */
function takeRange(
  slots: readonly Slot[],
  from: Pos,
  to: Pos,
  keepsLast = false
): { removed: Atom[]; anchors: Atom[] } {
  const removed: Atom[] = [];
  const anchors: Atom[] = [];
  for (let b = from.block; b <= to.block; b++) {
    const atoms = slots[b]!.atoms ?? [];
    if (b > from.block && !(keepsLast && b === to.block)) {
      // A paragraph joined to the first keeps the markers between it and the last.
      pushMarkers((slots[b]!.base as Paragraph).rangeMarkersBefore, [], anchors);
    }
    const start = b === from.block ? from.atom : 0;
    const stop = b === to.block ? to.atom : atoms.length - 1;
    for (let a = start; a <= stop; a++) {
      removed.push(atoms[a]!);
      keptAnchors(atoms[a]!, anchors);
    }
  }
  return { removed, anchors };
}

/** The visible run a replacement takes its formatting and wrappers from. */
function firstResultRun(removed: readonly Atom[]): Atom | undefined {
  return removed.find(
    a =>
      !isRangeMarker(a.node) &&
      isRun(a.node) &&
      !pieceOf(a.node) &&
      !a.chain.some(isDeletion) &&
      a.node.content.length > 0
  );
}

function liveChain(chain: readonly Wrapper[]): readonly Wrapper[] {
  return chain.some(isDeletion) ? chain.filter(w => !isDeletion(w)) : chain;
}

function joinedBase(first: Paragraph, last: Paragraph): Paragraph {
  const section = last.properties?.sectionProperties;
  if (first === last || !section || first.properties?.sectionProperties) {
    return first;
  }
  return { ...first, properties: { ...first.properties, sectionProperties: section } };
}

function applyInline(
  slots: Slot[],
  planned: PlannedEdit,
  edit: {
    readonly replace: "field" | "result";
    readonly text: string;
    readonly content?: readonly RunContent[];
  }
): void {
  const { from, to, places } = planned;
  const { removed, anchors } = takeRange(slots, from, to);
  const begin = runAt(slots, places.begin)!;
  const endAtom = runAt(slots, places.end!)!;
  const source = firstResultRun(removed);
  const anchor =
    source ??
    (edit.replace === "field" ? begin : places.separate ? runAt(slots, places.separate)! : endAtom);
  const properties: RunProperties | undefined = (anchor.node as Run).properties;
  const content =
    edit.content ?? (edit.text === "" ? [] : [{ type: "text", text: edit.text } as RunContent]);
  const insert: Atom[] = [];
  let endReplacement: Atom | undefined;
  if (edit.replace === "result" && !places.separate) {
    // A field with no result yet gets a separate character, which takes the
    // instruction text written before the end.
    const end = pieceOf(endAtom.node as Run)!;
    const run = endAtom.node as Run;
    insert.push({
      node: {
        ...run,
        content: [{ type: "field", fieldChar: "separate", instruction: end.instruction }]
      },
      chain: endAtom.chain
    });
    endReplacement = {
      node: { ...run, content: [{ ...end, instruction: "" }] },
      chain: endAtom.chain
    };
  }
  if (content.length > 0) {
    insert.push({
      node: { ...(properties ? { properties } : {}), content },
      chain: liveChain(anchor.chain)
    });
  }
  const first = slots[from.block]!;
  const last = slots[to.block]!;
  const head = first.atoms!.slice(0, from.atom);
  const tail = last.atoms!.slice(to.atom + 1);
  if (endReplacement) {
    tail[0] = endReplacement;
  }
  const joined: Slot = {
    base: joinedBase(first.base as Paragraph, last.base as Paragraph),
    atoms: [...head, ...insert, ...anchors, ...tail],
    dirty: true
  };
  slots.splice(from.block, to.block - from.block + 1, joined);
}

function applyParagraphs(
  slots: Slot[],
  planned: PlannedEdit,
  paragraphs: readonly Paragraph[]
): void {
  const { from, to, places } = planned;
  const end = places.end!;
  const first = slots[from.block]!;
  const last = slots[end.block]!;
  const prefix = first.atoms!.slice(0, from.atom);
  const tail = last.atoms!.slice(end.atom);
  const own = from.block !== end.block;
  // The end keeps the paragraph of its own it had when nothing of the result
  // was before it there.
  const endAlone = own && last.atoms!.slice(0, end.atom).every(a => isRangeAnchor(a.node));
  const { anchors } = takeRange(slots, from, to, own && (endAlone || paragraphs.length === 0));
  const out: Slot[] = [];
  if (paragraphs.length === 0) {
    out.push({
      base: first.base,
      atoms: own ? [...prefix, ...anchors] : [...prefix, ...anchors, ...tail],
      dirty: true
    });
    if (own) {
      out.push({ base: last.base, atoms: tail, dirty: true });
    }
  } else {
    paragraphs.forEach((p, i) => {
      let atoms = atomize(p);
      let base: Paragraph = p;
      if (i === 0) {
        atoms = [...prefix, ...anchors, ...atoms];
      }
      if (i === paragraphs.length - 1 && !endAlone) {
        atoms = [...atoms, ...tail];
        base = joinedBase(p, last.base as Paragraph);
      }
      out.push({ base, atoms, dirty: true });
    });
    if (endAlone) {
      out.push({ base: last.base, atoms: tail, dirty: true });
    }
  }
  slots.splice(from.block, end.block - from.block + 1, ...out);
}

// =============================================================================
// Display
// =============================================================================

/**
 * `blocks` without the content of any field's instruction — what a field
 * nested in another's instruction shows, say, which Word does not display.
 * The field pieces and the range markers stay. Returns `blocks` itself when
 * no such content is there.
 */
export function hideFieldCode<T extends object>(
  blocks: readonly T[],
  resolved: ResolvedFields | undefined = resolveSplitFields(blocks)
): readonly T[] {
  if (!resolved || resolved.code.length === 0) {
    return blocks;
  }
  const hidden = new Map<number, Set<number>>();
  for (const pos of resolved.code) {
    let set = hidden.get(pos.block);
    if (!set) {
      set = new Set();
      hidden.set(pos.block, set);
    }
    set.add(pos.atom);
  }
  return blocks.map((block, b) => {
    const set = hidden.get(b);
    if (!set) {
      return block;
    }
    const atoms: Atom[] = [];
    resolved.atoms[b]!.forEach((atom, a) => {
      if (set.has(a)) {
        keptAnchors(atom, atoms);
      } else {
        atoms.push(atom);
      }
    });
    return paragraphFrom(block as Paragraph, atoms) as T;
  });
}

const PAGE_FIELDS: ReadonlySet<string> = new Set(["PAGE", "NUMPAGES", "SECTIONPAGES"]);

/**
 * `blocks` as a document shows them: no field instruction content
 * ({@link hideFieldCode}), and every page-number field in pieces that lies in
 * one paragraph made whole, its result its `cachedValue`, so a layout
 * computes it as it does a whole one.
 */
function displayBlocks<T extends object>(blocks: readonly T[]): readonly T[] {
  const resolved = resolveSplitFields(blocks);
  if (!resolved) {
    return blocks;
  }
  const collapsed = rewriteSplitFields(
    blocks,
    field => {
      const name = parsePageField(field.instruction)?.name;
      const places = resolved.places.get(field)!;
      if (
        !name ||
        !PAGE_FIELDS.has(name) ||
        field.deleted ||
        field.parent !== undefined ||
        places.end?.block !== places.begin.block
      ) {
        return undefined;
      }
      return {
        replace: "field",
        text: "",
        content: [{ type: "field", instruction: field.instruction, cachedValue: field.result }]
      };
    },
    resolved
  );
  return hideFieldCode(collapsed, collapsed === blocks ? resolved : undefined);
}

/**
 * `blocks` with `fn` applied to it and to every list of blocks nested in it,
 * rebuilding only what changes.
 */
export function mapBlockLists<T extends object>(
  blocks: readonly T[],
  fn: <U extends object>(list: readonly U[]) => readonly U[]
): readonly T[] {
  const list = fn(blocks);
  let out: T[] | undefined;
  for (let i = 0; i < list.length; i++) {
    const block = list[i]!;
    const mapped = mapNested(block as object, fn) as T;
    if (mapped !== block) {
      out ??= list.slice();
      out[i] = mapped;
    }
  }
  return out ?? list;
}

function mapNested(
  block: object,
  fn: <U extends object>(list: readonly U[]) => readonly U[]
): object {
  if (!("type" in block)) {
    return block;
  }
  const b = block as BodyContent;
  switch (b.type) {
    case "table": {
      let rowsChanged = false;
      const rows = b.rows.map(row => {
        let changed = false;
        const cells = row.cells.map(cell => {
          const content = mapBlockLists(cell.content, fn);
          if (content === cell.content) {
            return cell;
          }
          changed = true;
          return { ...cell, content };
        });
        if (!changed) {
          return row;
        }
        rowsChanged = true;
        return { ...row, cells };
      });
      return rowsChanged ? { ...b, rows } : b;
    }
    case "sdt": {
      const content = mapBlockLists(b.content, fn);
      return content === b.content ? b : { ...b, content };
    }
    case "textBox": {
      const content = b.content && mapBlockLists(b.content, fn);
      return content === b.content ? b : { ...b, content };
    }
    case "drawingShape": {
      const textContent = b.textContent && mapBlockLists(b.textContent, fn);
      return textContent === b.textContent ? b : { ...b, textContent };
    }
    case "tableOfContents": {
      const cachedParagraphs = b.cachedParagraphs && mapBlockLists(b.cachedParagraphs, fn);
      return cachedParagraphs === b.cachedParagraphs ? b : { ...b, cachedParagraphs };
    }
    default:
      return b;
  }
}

/** Whether any list of blocks in `doc` holds a field piece. */
function documentHoldsPiece(doc: DocxDocument): boolean {
  let found = false;
  mapDocumentLists(doc, list => {
    if (!found) {
      found = list.some(b => isParagraph(b) && holdsFieldPiece(b));
    }
    return list;
  });
  return found;
}

/**
 * `doc` with `fn` applied to every list of blocks in it — the body, headers,
 * footers, notes and comments, and the lists nested in them — rebuilding only
 * what changes. Returns `doc` itself when nothing does.
 */
export function mapDocumentLists(
  doc: DocxDocument,
  fn: <U extends object>(list: readonly U[]) => readonly U[]
): DocxDocument {
  let next: Mutable<DocxDocument> | undefined;
  const set = <K extends keyof DocxDocument>(key: K, value: DocxDocument[K]): void => {
    next ??= { ...doc };
    next[key] = value;
  };
  const body = mapBlockLists(doc.body, fn);
  if (body !== doc.body) {
    set("body", body);
  }
  for (const key of ["headers", "footers"] as const) {
    const parts = doc[key];
    if (!parts) {
      continue;
    }
    let changed:
      | Map<string, typeof parts extends ReadonlyMap<string, infer V> ? V : never>
      | undefined;
    for (const [id, part] of parts) {
      const children = mapBlockLists(part.content.children, fn);
      if (children !== part.content.children) {
        changed ??= new Map(parts);
        changed.set(id, { ...part, content: { ...part.content, children } });
      }
    }
    if (changed) {
      set(key, changed);
    }
  }
  for (const key of ["footnotes", "endnotes", "comments"] as const) {
    const notes = doc[key] as readonly { readonly content: readonly Paragraph[] }[] | undefined;
    if (!notes) {
      continue;
    }
    let changed: { readonly content: readonly Paragraph[] }[] | undefined;
    notes.forEach((note, i) => {
      const content = mapBlockLists(note.content, fn);
      if (content !== note.content) {
        changed ??= notes.slice();
        changed[i] = { ...note, content };
      }
    });
    if (changed) {
      set(key, changed as never);
    }
  }
  return next ?? doc;
}

/** Documents {@link withDisplayFields} returned: already as they show. */
const displayed = new WeakSet<DocxDocument>();

/**
 * `doc` as it shows (see {@link displayBlocks}) — what the converters and the
 * page layout read. Returns `doc` itself when it holds no field in pieces.
 */
export function withDisplayFields(doc: DocxDocument): DocxDocument {
  if (displayed.has(doc) || !documentHoldsPiece(doc)) {
    return doc;
  }
  const out = mapDocumentLists(doc, displayBlocks);
  displayed.add(out);
  return out;
}
