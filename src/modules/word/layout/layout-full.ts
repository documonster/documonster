/**
 * Full Layout Engine — produces a complete LayoutDocument with positioned elements.
 *
 * Uses the pagination result from layoutDocument() for page assignments,
 * then computes precise positions (x, y, width, height) for every body
 * element on each page.
 *
 * This is the bridge between the page-number-only LayoutResult and the
 * fully positioned LayoutDocument that renderers (SVG, PDF, Canvas) can consume.
 *
 * Coverage: every variant of `BodyContent` from `../types` produces a
 * `PageContent` variant in the output. The `default:` branch of the
 * dispatch switch in `buildPage()` is a `never`-typed exhaustiveness
 * guard — adding a new body variant without a matching layout function
 * is a build error, never a silent drop.
 */

import { canBreakBetween, countGlyphAdvances, segmentForWrap, splitByScript } from "@utils/cjk";
import {
  getFontAscent,
  getFontDescent,
  measureTextWidth,
  styledFontVariant
} from "@utils/font-metrics";
import { s } from "@utils/schema";
import type { SchemaMatches } from "@utils/schema";
import { ommlToMathML } from "@word/advanced/math-convert";
import { finalizeDocument } from "@word/builder/document-finalize";
import { checkHandle } from "@word/core/argument-checks";
import { withDisplayFields } from "@word/core/field-view";
import { sectionReferences } from "@word/core/header-references";
import { indexImagesByKey } from "@word/core/image-keys";
import { checkSchema, w } from "@word/core/option-schema";
import { extractMathText, isHyperlink, symbolText } from "@word/core/text-utils";
import { walkBlocks } from "@word/core/walker";
import { faceExtents, faceLineExtents, measureFaceText } from "@word/layout/font-face";
import { layoutDocument } from "@word/layout/layout";
import type { LayoutOptions, LayoutResult } from "@word/layout/layout";
import {
  DEFAULT_FONT_SIZE_PT,
  DEFAULT_PAGE_HEIGHT_TWIPS,
  DEFAULT_PAGE_MARGIN_TWIPS,
  DEFAULT_PAGE_WIDTH_TWIPS,
  FIT_EPSILON_PT,
  ROTATED_COLUMN_EXTRA_PT,
  ROTATED_LINE_WIDTH_PT,
  ROTATED_PARAGRAPH_EXTRA_PT,
  isRotatedTextDirection,
  LINE_HEIGHT_FACTOR,
  layoutRunProperties,
  mergeRunProperties,
  minimumRowHeightPt,
  gridColumnCount,
  resolveColumnWidthsTwips,
  SCRIPT_BASELINE_SHIFT_FACTOR,
  SCRIPT_FONT_SIZE_RATIO,
  isCellEndAfterNestedTable,
  nestedTableWidthTwips,
  resolveCellMarginsTwips,
  resolveHeadingScale
} from "@word/layout/layout-constants";
import type {
  LayoutAltChunk,
  LayoutBorderEdge,
  LayoutBorders,
  LayoutChart,
  LayoutCheckBox,
  LayoutDocument,
  LayoutFloat,
  LayoutImage,
  LayoutMath,
  LayoutOpaqueDrawing,
  LayoutPage,
  LayoutParagraph,
  LayoutRect,
  LayoutSdt,
  LayoutShape,
  LayoutTable,
  LayoutTableCell,
  LayoutTableOfContents,
  LayoutTextBox,
  LineBox,
  LineBoxItem,
  PageContent,
  PageGeometry,
  PositionedRun
} from "@word/layout/layout-model";
import { resolveWordLineMetrics } from "@word/layout/line-metrics";
import { finalViewRun, finalViewRuns, isHiddenRun } from "@word/query/final-view";
import { createNumberingCounter, placeholderFormat } from "@word/query/numbering-counter";
import {
  collectSectionProperties,
  formatListCounter,
  formatPageField,
  hasFacingPages,
  nextDisplayedPage,
  openSection,
  type PageSectionSlot,
  parsePageField,
  sectionPageCounts
} from "@word/query/page-field";
import { resolveShadingFill } from "@word/query/style-resolve";
import type { StyleResolveContext } from "@word/query/style-resolve";
import { createStyleResolver } from "@word/query/style-resolver";
import type { StyleResolver } from "@word/query/style-resolver";
import type {
  AltChunk,
  BodyContent,
  FooterDef,
  HeaderDef,
  RangeMarker,
  Border,
  ChartContent,
  ChartExContent,
  CheckBox,
  DocxDocument,
  DrawingShape,
  FloatingImage,
  FootnoteDef,
  ImageDef,
  InlineImageContent,
  LineSpacing,
  MathBlock,
  OpaqueDrawing,
  Paragraph,
  ParagraphFrame,
  ParagraphBorders,
  FieldContent,
  ParagraphProperties,
  PageNumberFormat,
  Run,
  HeaderFooterType,
  SectionProperties,
  StructuredDocumentTag,
  Table,
  TableBorders,
  TableCellProperties,
  TableOfContents,
  TableProperties,
  TabStop,
  TabStopLeader,
  TextBox,
  VerticalCellAlign
} from "@word/types";
import { EMU_PER_POINT, ptToTwips, twipsToPt } from "@word/units";

// =============================================================================
// Public API
// =============================================================================

/**
 * Page geometry overrides for {@link FullLayoutOptions}. All fields are
 * in **points** (72 per inch) — not the twips (1440 per inch) that
 * `doc.sectionProperties` holds them in. Any field not supplied falls back to
 * the corresponding value resolved from `doc.sectionProperties` (or the
 * engine defaults).
 */
export interface PageGeometryOverride {
  readonly pageWidth?: number;
  readonly pageHeight?: number;
  readonly marginTop?: number;
  readonly marginBottom?: number;
  readonly marginLeft?: number;
  readonly marginRight?: number;
  /**
   * Distance of the header band from the top edge of the page, in
   * points. Overrides the section's `pgMar.header`. Header paragraphs
   * are laid out starting at this y-offset from the page top.
   */
  readonly headerMargin?: number;
  /**
   * Distance of the footer band from the bottom edge of the page, in
   * points. Overrides the section's `pgMar.footer`. The footer band's
   * top is placed at `pageHeight - footerMargin`.
   */
  readonly footerMargin?: number;
}

/** Options for the full layout engine. */
export interface FullLayoutOptions extends LayoutOptions {
  /** Font map for font-family resolution (name → actual font). */
  readonly fonts?: ReadonlyMap<string, string>;
  /**
   * Override the page geometry resolved from `doc.sectionProperties`.
   * Used by hosts that drive layout with their own page model (e.g. the
   * PDF bridge translating `DocxToPdfOptions.pageWidth` into a layout
   * geometry override). Any unspecified field falls back to the
   * section properties / engine defaults.
   */
  readonly pageGeometry?: PageGeometryOverride;
}

/**
 * Perform full document layout, producing a LayoutDocument with precise positions.
 *
 * @param doc - The parsed DOCX document.
 * @param options - Layout and font options.
 * @returns A fully positioned LayoutDocument.
 */
const positive = () => w.finite({ positive: true }).optional();

const layoutOptionsShape = {
  defaultFontSize: positive(),
  defaultCharsPerLine: w.integer({ min: 1 }).optional(),
  averageCharWidth: positive(),
  measureText: w.function<NonNullable<LayoutOptions["measureText"]>>().optional(),
  measureTextMetrics: w.function<NonNullable<LayoutOptions["measureTextMetrics"]>>().optional()
};

const layoutOptionsSchema = w.record(layoutOptionsShape);
const _layoutOptionsMatch: SchemaMatches<typeof layoutOptionsSchema, LayoutOptions> = true;
void _layoutOptionsMatch;

/**
 * Layout options: every number is finite and positive where a zero or negative one has no
 * meaning (a font size, a page width), and a callback is a function.
 */
const fullLayoutOptionsSchema = w.record({
  ...layoutOptionsShape,
  fonts: s.custom<ReadonlyMap<string, string>>(v => v instanceof Map, "a Map").optional(),
  pageGeometry: w
    .record({
      pageWidth: positive(),
      pageHeight: positive(),
      marginTop: w.finite().optional(),
      marginBottom: w.finite().optional(),
      marginLeft: w.finite().optional(),
      marginRight: w.finite().optional(),
      headerMargin: w.finite().optional(),
      footerMargin: w.finite().optional()
    })
    .optional()
});
const _fullLayoutOptionsMatch: SchemaMatches<typeof fullLayoutOptionsSchema, FullLayoutOptions> =
  true;
void _fullLayoutOptionsMatch;

/** Check layout options, the parameter `field` of `where`, where they are given. */
export function checkLayoutOptions(options: unknown, field: string, where: string): void {
  checkSchema(fullLayoutOptionsSchema.optional(), options, where, field);
}

export function layoutDocumentFull(doc: DocxDocument, options?: FullLayoutOptions): LayoutDocument {
  checkHandle(doc, "document", "doc", "Layout.compute");
  checkLayoutOptions(options, "options", "Layout.compute");
  return layoutWithEstimate(doc, options).layout;
}

/**
 * The page each body item, bookmark and section of `doc` lands on, as the
 * positioned layout places them — the pages {@link layoutDocumentFull} produces,
 * and so the pages a PDF of `doc` has. Page-number fields are resolved from this:
 * resolving them from the estimating paginator alone gave numbers that disagreed
 * with the document's own rendering.
 */
export function paginateDocument(doc: DocxDocument, options?: FullLayoutOptions): LayoutResult {
  checkHandle(doc, "document", "doc", "Layout.getPageMap");
  checkLayoutOptions(options, "options", "Layout.getPageMap");
  const { layout, estimate } = layoutWithEstimate(doc, options);
  const contentPages = estimate.contentPages.slice();
  const seen = new Set<number>();
  for (const page of layout.pages) {
    for (const block of page.content) {
      if (block.sourceIndex >= 0 && !seen.has(block.sourceIndex)) {
        seen.add(block.sourceIndex);
        contentPages[block.sourceIndex] = page.pageNumber;
      }
    }
  }
  // An item pass 2 placed nothing of (an empty section-break paragraph, say) is
  // on the page of the item before it.
  for (let i = 0; i < contentPages.length; i++) {
    if (!seen.has(i)) {
      contentPages[i] = i > 0 ? contentPages[i - 1]! : 1;
    }
  }
  const breaks = layout.sectionBreaks;
  const sectionPageCounts = breaks.map(
    (start, i) => (i + 1 < breaks.length ? breaks[i + 1]! : layout.totalPages) - start
  );
  return {
    pageCount: layout.totalPages,
    sectionPageCounts,
    contentPages,
    contentSections: estimate.contentSections,
    bookmarkPages: layout.bookmarkPages,
    forcedBreakBefore: estimate.forcedBreakBefore
  };
}

/**
 * Everything about the layout of `doc` that decides where later content goes —
 * every page, block, line and table cell box — and nothing it draws (text,
 * images, charts). Two documents with equal footprints place content that
 * follows them identically.
 */
export function layoutFootprint(doc: DocxDocument, options?: FullLayoutOptions): string {
  const { layout } = layoutWithEstimate(doc, options);
  return JSON.stringify(layout.pages, (key, value: unknown) =>
    key === "runs" || key === "data" || key === "svg" || key === "source" ? undefined : value
  );
}

function layoutWithEstimate(
  doc: DocxDocument,
  options: FullLayoutOptions | undefined
): { layout: LayoutDocument; estimate: LayoutResult } {
  const withNotes = withEndnotesAtEnd(withDisplayFields(finalizeDocument(doc)));
  doc = withNotes.doc;
  // First pass: get page assignments via the existing lightweight layout
  const layoutResult = layoutDocument(doc, options);

  // Resolve list markers once over the whole document so ordered-list
  // counters increment correctly across pages. Stored in a module-level
  // context so that every `layoutParagraph` call — including those reached
  // through tables, text boxes, SDTs, footnotes, etc. — can render markers
  // without threading the map through every container function.
  const styles = createStyleResolver(doc);
  const listMarkers = computeListMarkers(doc, styles);
  // The previous values are restored rather than cleared. Layout is
  // synchronous, but it is not un-reentrant: a caller's `measureText` callback
  // can lay out another document (to size a caption, say), and clearing the
  // slots on the way out of the inner call left the outer one resolving styles,
  // list markers and contextual spacing against nothing for every remaining
  // paragraph. `layout.ts` already saves and restores; this now matches.
  const previousListMarkers = activeListMarkers;
  const previousDoc = activeDoc;
  const previousStyles = activeStyles;
  const previousContextualSpacing = activeContextualSpacing;
  activeListMarkers = listMarkers;
  activeDoc = doc;
  activeStyles = styles;
  // `w:contextualSpacing` needs each paragraph's *siblings*, which no single
  // `layoutParagraph` call can see. Resolved once here for the same reason and
  // through the same mechanism as the list markers.
  activeContextualSpacing = computeContextualSpacing(doc, styles);
  const previousEndnotesStart = activeEndnotesStart;
  activeEndnotesStart = withNotes.notesStart;
  try {
    const layout = layoutDocumentFullInner(doc, options, layoutResult, listMarkers);
    return withNotes.origin
      ? onOriginalBody(layout, layoutResult, withNotes.origin)
      : { layout, estimate: layoutResult };
  } finally {
    activeListMarkers = previousListMarkers;
    activeDoc = previousDoc;
    activeStyles = previousStyles;
    activeContextualSpacing = previousContextualSpacing;
    activeEndnotesStart = previousEndnotesStart;
  }
}

/**
 * `doc` with its endnotes set after the last of its text, where Word puts them
 * by default (`w:endnotePr/w:pos` `docEnd`): the endnote separator, then each
 * note referenced in the body, in the order of the references (pages2 R2h).
 * `origin` gives, for each item of the returned body, the index of the item of
 * `doc.body` it is, or -1 for an inserted note; absent when nothing moved.
 */
function withEndnotesAtEnd(doc: DocxDocument): {
  doc: DocxDocument;
  origin?: number[];
  notesStart?: number;
} {
  const ids: number[] = [];
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const child of node) {
        visit(child);
      }
      return;
    }
    if (!node || typeof node !== "object") {
      return;
    }
    const record = node as { type?: string; id?: number };
    if (record.type === "endnoteRef" && typeof record.id === "number") {
      if (!ids.includes(record.id)) {
        ids.push(record.id);
      }
      return;
    }
    for (const value of Object.values(node)) {
      if (value && typeof value === "object") {
        visit(value);
      }
    }
  };
  if (!doc.endnotes?.some(note => (note.type ?? "normal") === "normal")) {
    return { doc };
  }
  const defs = new Map<number, readonly Paragraph[]>();
  for (const note of doc.endnotes) {
    if ((note.type ?? "normal") === "normal") {
      defs.set(note.id, note.content);
    }
  }
  // The model holds no separator when the package carried the default one,
  // which the writer puts back; Word draws it on a line of its own above the
  // first note (pages2 R11b: the notes open a page under a 13.4pt line).
  const separator = doc.endnotes.find(note => note.type === "separator")?.content;
  const separatorParas = separator?.length ? separator : [DEFAULT_SEPARATOR_PARAGRAPH];
  // `w:pos="sectEnd"` (document settings) ends every section with the notes
  // referenced in it, set before its break (pages2 R3j).
  if (doc.settings?.endnoteProperties?.position === "sectEnd") {
    const body: BodyContent[] = [];
    const origin: number[] = [];
    for (const [index, item] of doc.body.entries()) {
      const sectionProperties =
        item.type === "paragraph" ? item.properties?.sectionProperties : undefined;
      if (!sectionProperties) {
        body.push(item);
        origin.push(index);
        visit(item);
        continue;
      }
      const { sectionProperties: _moved, ...rest } = (item as Paragraph).properties!;
      void _moved;
      body.push({ ...(item as Paragraph), properties: rest });
      origin.push(index);
      visit(item);
      const notes = ids.splice(0).flatMap(id => defs.get(id) ?? []);
      const added = notes.length > 0 ? [...separatorParas, ...notes] : [];
      const last = added.at(-1);
      if (last) {
        body.push(...added.slice(0, -1), {
          ...last,
          properties: { ...last.properties, sectionProperties }
        });
        origin.push(...added.map(() => -1));
      } else {
        body[body.length - 1] = item;
      }
    }
    const tail = ids.flatMap(id => defs.get(id) ?? []);
    const added = tail.length > 0 ? [...separatorParas, ...tail] : [];
    origin.push(...added.map(() => -1));
    return { doc: { ...doc, body: [...body, ...added] }, origin };
  }
  visit(doc.body);
  const notes = ids.flatMap(id => defs.get(id) ?? []);
  if (notes.length === 0) {
    return { doc };
  }
  const added = [...separatorParas, ...notes];
  return {
    doc: { ...doc, body: [...doc.body, ...added] },
    origin: [...doc.body.keys(), ...added.map(() => -1)],
    notesStart: doc.body.length + separatorParas.length
  };
}

/**
 * `layout` and `estimate` with their body indices put back on the caller's
 * body, where {@link withEndnotesAtEnd} inserted notes: a note's blocks are not
 * body items (`sourceIndex` -1), and the per-item arrays drop the notes.
 */
function onOriginalBody(
  layout: LayoutDocument,
  estimate: LayoutResult,
  origin: readonly number[]
): { layout: LayoutDocument; estimate: LayoutResult } {
  const keep = <T>(list: readonly T[]): T[] => list.filter((_, i) => (origin[i] ?? -1) >= 0);
  const remap = <B extends { readonly sourceIndex: number }>(block: B): B =>
    block.sourceIndex < 0 ? block : { ...block, sourceIndex: origin[block.sourceIndex] ?? -1 };
  return {
    layout: {
      ...layout,
      pages: layout.pages.map(page => ({ ...page, content: page.content.map(remap) }))
    },
    estimate: {
      ...estimate,
      contentPages: keep(estimate.contentPages),
      contentSections: keep(estimate.contentSections),
      forcedBreakBefore: keep(estimate.forcedBreakBefore)
    }
  };
}

/**
 * Body index of the first endnote {@link withEndnotesAtEnd} set after the
 * document's text, for the in-flight layout; `undefined` when there are none.
 */
let activeEndnotesStart: number | undefined;

/** Active list-marker map for the in-flight layout (see layoutDocumentFull). */
let activeListMarkers: ReadonlyMap<Paragraph, ListMarker> | undefined;

/** Active contextual-spacing map for the in-flight layout. */
let activeContextualSpacing: ReadonlyMap<Paragraph, ContextualSpacing> | undefined;

/**
 * Active document for the in-flight layout, so `layoutParagraph` can resolve
 * paragraph-style run properties (size/color/font) via `resolveStyle` without
 * threading `doc` through every container function. Layout is synchronous so a
 * single shared slot is safe.
 */
let activeDoc: DocxDocument | undefined;

/** Style resolution bound to `activeDoc`, same lifetime. */
let activeStyles: StyleResolver | undefined;

/**
 * Page facts the page-number fields of a header or footer read while it is laid
 * out, or `undefined` outside that pass (fields then show their cached result).
 * Same single-slot reasoning as `activeDoc`.
 *
 * `numPages` / `sectionPages` are only known once every page exists, so the
 * first pass leaves them unset and headers and footers that need them are laid
 * out again at the end (see `relayoutPageTotals`). Without that, `PAGE` would
 * be live while `NUMPAGES` kept its stale cache — "Page 3 of 1".
 *
 * A slot rather than a parameter because the fields are reached through the
 * body layout's own container functions (`layoutParagraph` →
 * `collectParagraphSegments` → `pushRunSegments`, and `layoutTable` → cells →
 * nested tables, SDTs and text boxes), none of which otherwise know about pages.
 */
interface PageFieldValues {
  /** Number `PAGE` shows; see `displayedPageNumbers`. */
  readonly page: number;
  /** `w:pgNumType w:fmt` of the page's section. */
  readonly format: PageNumberFormat | undefined;
  /** Physical pages in the document (`NUMPAGES`). */
  readonly numPages?: number;
  /** Pages in the page's section (`SECTIONPAGES`). */
  readonly sectionPages?: number;
}

let activePageFields: PageFieldValues | undefined;

/**
 * The line pitch, in points, of the document grid the paragraph being laid out
 * snaps to, or `undefined` off the grid. Body text and notes of a section whose
 * `w:docGrid` has lines snap; Word leaves headers, footers and table cells off
 * it (pages2 P7, P8). Same single-slot reasoning as `activeDoc`.
 */
let activeGridPitch: number | undefined;

/** Run `fn` with lines snapping to a grid of `pitch` points (or none). */
function withGridPitch<T>(pitch: number | undefined, fn: () => T): T {
  const previous = activeGridPitch;
  activeGridPitch = pitch;
  try {
    return fn();
  } finally {
    activeGridPitch = previous;
  }
}

/** The line pitch of a section's document grid, when it has lines. */
function sectionGridPitch(props: SectionProperties | undefined): number | undefined {
  const grid = props?.docGrid;
  if ((grid?.type !== "lines" && grid?.type !== "linesAndChars") || !grid.linePitch) {
    return undefined;
  }
  return twipsToPt(grid.linePitch);
}

/** The page facts for one page's header and footer. */
function pageFieldValues(
  slot: PageSlot,
  totals?: { readonly numPages: number; readonly sectionPages: number }
): PageFieldValues {
  return {
    page: slot.displayedPage,
    format: slot.props?.pageNumbering?.format,
    ...totals
  };
}

/** Run `fn` with page-number fields resolving against `values`. */
function withPageFields<T>(values: PageFieldValues | undefined, fn: () => T): T {
  const previous = activePageFields;
  activePageFields = values;
  try {
    return fn();
  } finally {
    activePageFields = previous;
  }
}

function layoutDocumentFullInner(
  doc: DocxDocument,
  options: FullLayoutOptions | undefined,
  layoutResult: LayoutResult,
  listMarkers: ReadonlyMap<Paragraph, ListMarker>
): LayoutDocument {
  // Second pass: compute precise positions for each page. Footnote
  // ids that don't fit on a given page are carried over to the next
  // (a later page may still have room thanks to less body content
  // or fewer of its own newly-introduced notes).
  const pages: LayoutPage[] = [];
  // Per page, in order: where it sits among the sections and the number `PAGE`
  // shows on it. Headers and footers that need page totals are laid out again
  // from this once every page exists (see `relayoutPageTotals`).
  const slots: PageSlot[] = [];
  let pendingNotes: readonly LaidNote[] = [];
  // Each note laid out once per text width.
  const noteCache = new Map<string, LaidNote>();
  // Blocks the previous page could not hold. Pass 2 measures for real, so it —
  // not the estimating paginator — decides what fits; see `BuildPageResult`.
  let carried: readonly CarriedBlock[] = [];
  // Where the flow has got to in `doc.body`; see `FlowCursor`.
  const flow: FlowCursor = { nextItem: 0 };
  // Image lookup depends only on `doc.images`; build it once for every page.
  const imageMap = buildImageMap(doc.images);

  // Each section brings its own paper size, margins and headers, so every page
  // has to know which one it belongs to and how far into it it is (`titlePage`
  // selects the "first" header on a section's own first page).
  const sectionPropsList = collectSectionProperties(doc);
  const headersInEffect = referencesInEffect(sectionPropsList, doc.headers, "headers");
  const footersInEffect = referencesInEffect(sectionPropsList, doc.footers, "footers");
  const sectionOf = (itemIndex: number): number =>
    layoutResult.contentSections[itemIndex] ?? sectionPropsList.length - 1;
  let sectionIndex = doc.body.length > 0 ? sectionOf(0) : sectionPropsList.length - 1;
  let pageInSection = 1;

  /** The next page's slot, with its displayed number continuing the last one. */
  const slotFor = (index: number, inSection: number): PageSlot => {
    const previous = slots.at(-1);
    const slot = { sectionIndex: index, pageInSection: inSection };
    return {
      ...slot,
      props: sectionPropsList[index],
      headers: headersInEffect[index],
      footers: footersInEffect[index],
      gutterSide: gutterSide(sectionPropsList[index], sectionPropsList[0], doc.settings),
      displayedPage: nextDisplayedPage(
        previous,
        previous?.displayedPage ?? 0,
        slot,
        sectionPropsList
      )
    };
  };

  const maxPages = 20_000;
  /**
   * Pages holding nothing but the notes still queued, continuing the last
   * page's section — its geometry, headers and numbering — until the queue
   * drains. One extra page is not always enough: three notes each nearly a page
   * tall put the first on the body page, the second on the extra page, and
   * dropped the third entirely.
   */
  const emitNoteOnlyPages = (): void => {
    while (pendingNotes.length > 0 && pages.length > 0 && pages.length < maxPages) {
      const last = slots[slots.length - 1];
      const slot = slotFor(last.sectionIndex, last.pageInSection + 1);
      const overflowResult = buildPage(
        doc,
        pages.length + 1,
        layoutResult,
        options,
        pendingNotes,
        listMarkers,
        // No carried content and a cursor past the body: this page hosts only
        // the deferred footnote queue.
        [],
        { nextItem: doc.body.length },
        slot,
        imageMap,
        noteCache
      );
      pages.push(overflowResult.page);
      slots.push(slot);
      const remaining = overflowResult.continuedNotes;
      // The first note of an empty area always places at least a line, so the
      // queue shrinks; stop rather than loop should that ever fail.
      if (totalNoteLines(remaining) >= totalNoteLines(pendingNotes)) {
        break;
      }
      pendingNotes = remaining;
    }
  };

  // Emit pages until the body is exhausted and nothing is left carried. Every
  // iteration either advances `flow.nextItem`, places part of a carried block
  // (a page that starts empty always places at least one line or row — see
  // `splitLayoutParagraph`) or places part of a continued note, so this
  // terminates. `maxPages` only guards a
  // degenerate geometry with no usable content height at all.
  do {
    const slot = slotFor(sectionIndex, pageInSection);
    const pendingNotesIn = pendingNotes;
    const result = composePage(
      doc,
      pages.length + 1,
      layoutResult,
      options,
      pendingNotes,
      listMarkers,
      carried,
      flow,
      slot,
      imageMap,
      noteCache,
      sectionPropsList
    );
    pages.push(result.page);
    slots.push(slot);
    pendingNotes = result.continuedNotes;
    if (
      result.page.content.length === 0 &&
      carried.length > 0 &&
      totalNoteLines(result.continuedNotes) >= totalNoteLines(pendingNotesIn)
    ) {
      // Nothing could be placed at all — stop rather than loop forever. A page
      // that holds only notes continued from the last one has made progress.
      break;
    }
    carried = result.carriedContent;
    // A section that began part-way down this page continues on the next.
    if (result.endSection !== undefined) {
      sectionIndex = result.endSection;
      pageInSection = 1;
    }

    // Which section the next page belongs to. Anything still carried belongs to
    // the section it came from, so the section only advances once the carry is
    // clear and the flow has crossed a boundary.
    const nextSection =
      carried.length === 0 && flow.nextItem < doc.body.length
        ? sectionOf(flow.nextItem)
        : sectionIndex;
    if (nextSection !== sectionIndex) {
      // Notes belong to the section that references them: Word finishes one
      // still continuing on pages of its own before the next section opens
      // (final7 N6d: a landscape page holding only the rest of note 4).
      emitNoteOnlyPages();
      sectionIndex = nextSection;
      pageInSection = 1;
      // An odd/even-page break that lands on the wrong parity gets a blank page
      // when the section continues numbering (see `openSection`).
      const shown = slots.at(-1)?.displayedPage ?? 0;
      if (
        openSection(shown, sectionPropsList[sectionIndex], hasFacingPages(doc)).blankPage &&
        pages.length < maxPages
      ) {
        const blankSlot = slotFor(sectionIndex, 0);
        const blank = buildPage(
          doc,
          pages.length + 1,
          layoutResult,
          options,
          [],
          listMarkers,
          [],
          { nextItem: doc.body.length },
          blankSlot,
          imageMap,
          noteCache
        );
        pages.push(blank.page);
        slots.push(blankSlot);
      }
    } else {
      pageInSection++;
    }
  } while ((flow.nextItem < doc.body.length || carried.length > 0) && pages.length < maxPages);

  // Notes still queued after the body runs out get pages of their own.
  emitNoteOnlyPages();

  relayoutPageTotals(doc, pages, slots, options, imageMap);

  // Bookmark pages and section-break page indices have to describe the pages
  // that were actually produced, not the estimate's. Both are remapped through
  // where pass 2 really put each body item.
  const itemPage = new Map<number, number>();
  for (const page of pages) {
    for (const block of page.content) {
      if (block.sourceIndex >= 0 && !itemPage.has(block.sourceIndex)) {
        itemPage.set(block.sourceIndex, page.pageNumber);
      }
    }
  }

  return {
    pages,
    totalPages: pages.length,
    bookmarkPages: remapBookmarkPages(doc.body, doc.bodyRangeMarkersAfter, pages, layoutResult),
    sectionBreaks: computeSectionBreaks(layoutResult, itemPage)
  };
}

/**
 * Where one page sits among the sections, and the number `PAGE` shows on it.
 * Plain data, so headers and footers can be laid out again from it once page
 * totals are known.
 */
interface PageSlot extends PageSectionSlot {
  /** The section's properties (`undefined` when the document declares none). */
  readonly props: SectionProperties | undefined;
  /** The header of each type in effect: the section's own, else the previous section's (§17.10.5). */
  readonly headers: ReadonlyMap<HeaderFooterType, string>;
  /** The footer of each type in effect, inherited the same way. */
  readonly footers: ReadonlyMap<HeaderFooterType, string>;
  /** The number `PAGE` shows; see `displayedPageNumbers`. */
  readonly displayedPage: number;
  /** The margin the section's gutter widens; see `gutterSide`. */
  readonly gutterSide: GutterSide;
}

type GutterSide = "left" | "right" | "top" | "bottom";

/**
 * The header (or footer) of each type every section shows. A section that
 * names no reference of a type keeps the previous section's; the first section
 * has none of that type (ECMA-376 §17.10.5).
 */
function referencesInEffect(
  sections: readonly (SectionProperties | undefined)[],
  parts: ReadonlyMap<string, HeaderDef | FooterDef> | undefined,
  kind: "headers" | "footers"
): ReadonlyMap<HeaderFooterType, string>[] {
  let inEffect = new Map<HeaderFooterType, string>();
  return sections.map((section, i) => {
    const own = new Map<HeaderFooterType, string>();
    // The packager's rule, shared: references resolved to part keys, and the
    // last section's implied ones when it names none.
    for (const ref of sectionReferences(section?.[kind], parts, i === sections.length - 1)) {
      if (!own.has(ref.type)) {
        own.set(ref.type, ref.rId);
      }
    }
    inEffect = new Map([...inEffect, ...own]);
    return inEffect;
  });
}

function isLandscape(props: SectionProperties | undefined): boolean {
  const size = props?.pageSize;
  return (
    size?.orientation === "landscape" ||
    (size?.orientation === undefined && (size?.width ?? 0) > (size?.height ?? 0))
  );
}

/**
 * Where Word puts a section's gutter. `w:gutterAtTop` names the top, otherwise
 * the left (the right in a right-to-left section) — on pages oriented like the
 * document's first section. A page turned the other way turns the gutter with
 * it: a landscape page in a portrait document takes it at the left with
 * `w:gutterAtTop` and at the bottom without, a portrait page in a landscape
 * document at the left and at the top (measured in Word 16.114).
 */
function gutterSide(
  props: SectionProperties | undefined,
  first: SectionProperties | undefined,
  settings: DocxDocument["settings"]
): GutterSide {
  const atTop = settings?.gutterAtTop === true;
  if (isLandscape(props) === isLandscape(first)) {
    return atTop ? "top" : props?.rtlGutter === true ? "right" : "left";
  }
  if (atTop) {
    return "left";
  }
  return isLandscape(first) ? "top" : "bottom";
}

/**
 * Lay out again the headers and footers of a document whose page-number
 * fields need totals (`NUMPAGES`, `SECTIONPAGES`), now that every page exists.
 * `PAGE` alone is already right after the first pass.
 *
 * A header or footer band has fixed geometry, so this replaces only the bands
 * and cannot change pagination. Documents without such fields skip it.
 */
function relayoutPageTotals(
  doc: DocxDocument,
  pages: LayoutPage[],
  slots: readonly PageSlot[],
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>
): void {
  if (!headerFootersUsePageTotals(doc)) {
    return;
  }
  const sectionPages = sectionPageCounts(slots);
  for (let i = 0; i < pages.length; i++) {
    const slot = slots[i];
    const geometry = computePageGeometry(slot.props, options?.pageGeometry, slot.gutterSide);
    const totals = {
      numPages: pages.length,
      sectionPages: sectionPages.get(slot.sectionIndex) ?? pages.length
    };
    const header = layoutHeader(doc, slot, geometry, options, imageMap, totals);
    const footer = layoutFooter(doc, slot, geometry, options, imageMap, totals);
    const { header: _h, footer: _f, ...rest } = pages[i];
    pages[i] = {
      ...rest,
      ...(header.length > 0 ? { header } : {}),
      ...(footer.length > 0 ? { footer } : {})
    };
  }
}

/** Whether any header or footer holds a field that reads page totals. */
function headerFootersUsePageTotals(doc: DocxDocument): boolean {
  const parts = [...(doc.headers?.values() ?? []), ...(doc.footers?.values() ?? [])];
  return parts.some(part => {
    let found = false;
    walkBlocks(part.content.children, {
      visitRunContent(content) {
        if (content.type === "field" && !content.fieldChar) {
          const name = parsePageField(content.instruction)?.name;
          found ||= name === "NUMPAGES" || name === "SECTIONPAGES";
        }
      },
      enterParagraph: () => (found ? "stop" : undefined)
    });
    return found;
  });
}

/**
 * Bookmark name → page number, keyed to the pages pass 2 produced.
 *
 * The paginator records the page it *estimated* for each bookmark; once pass 2
 * re-flows content to its own measurements those numbers can be off by a page or
 * more, which would send a PDF outline entry to the wrong place.
 */
function remapBookmarkPages(
  body: readonly BodyContent[],
  bodyAfter: readonly RangeMarker[] | undefined,
  pages_: readonly LayoutPage[],
  fallback: LayoutResult
): ReadonlyMap<string, number> {
  const pages = new Map<string, number>();
  resolveBlockBookmarkPages(body, bodyAfter, pages_, pages);
  // Read the placed blocks rather than the source body: a paragraph split across
  // a page boundary reports per-line bookmark names, so a bookmark in the part
  // that spilled over resolves to the page it actually landed on instead of the
  // page the paragraph started on.
  for (const page of pages_) {
    const visit = (blocks: readonly PageContent[] | readonly (LayoutParagraph | LayoutTable)[]) => {
      for (const block of blocks) {
        if (block.type === "paragraph") {
          for (const names of block.lineBookmarks ?? []) {
            for (const name of names) {
              if (!pages.has(name)) {
                pages.set(name, page.pageNumber);
              }
            }
          }
        } else if (block.type === "table") {
          for (const cell of block.cells) {
            visit(cell.content);
          }
        }
      }
    };
    visit(page.content);
  }
  // Anything pass 2 never placed keeps the paginator's estimate.
  for (const [name, page] of fallback.bookmarkPages) {
    if (!pages.has(name)) {
      pages.set(name, page);
    }
  }
  return pages;
}

/**
 * Pages of the bookmarks that start *between* blocks — in the body, or between
 * the rows and cells of a top-level table — where no laid-out line carries
 * them. Such a bookmark is on the page where the content after it starts.
 */
function resolveBlockBookmarkPages(
  body: readonly BodyContent[],
  bodyAfter: readonly RangeMarker[] | undefined,
  pages_: readonly LayoutPage[],
  out: Map<string, number>
): void {
  // First page of each body item, and of each cell of each top-level table.
  const itemPage = new Map<number, number>();
  const cellPage = new Map<string, number>();
  const lastPage = new Map<number, number>();
  for (const page of pages_) {
    for (const block of page.content) {
      if (block.sourceIndex < 0) {
        continue;
      }
      if (!itemPage.has(block.sourceIndex)) {
        itemPage.set(block.sourceIndex, page.pageNumber);
      }
      lastPage.set(block.sourceIndex, page.pageNumber);
      if (block.type === "table") {
        for (const cell of block.cells) {
          const key = `${block.sourceIndex}:${cell.row}:${cell.col}`;
          if (!cellPage.has(key)) {
            cellPage.set(key, page.pageNumber);
          }
        }
      }
    }
  }
  const finalPage = pages_.length > 0 ? pages_[pages_.length - 1]!.pageNumber : 1;
  const set = (markers: readonly RangeMarker[] | undefined, page: number): void => {
    for (const marker of markers ?? []) {
      if (marker.type === "bookmarkStart" && !out.has(marker.name)) {
        out.set(marker.name, page);
      }
    }
  };
  /** First page at or after cell (row, col) of the table at `index`. */
  const cellAt = (index: number, table: Table, row: number, col: number): number | undefined => {
    for (let r = row; r < table.rows.length; r++) {
      const cells = table.rows[r]!.cells.length;
      for (let c = r === row ? col : 0; c < cells; c++) {
        const page = cellPage.get(`${index}:${r}:${c}`);
        if (page !== undefined) {
          return page;
        }
      }
    }
    return undefined;
  };
  let pending: RangeMarker[] = [];
  for (let i = 0; i < body.length; i++) {
    const item = body[i]!;
    pending.push(...(item.rangeMarkersBefore ?? []));
    const start = itemPage.get(i);
    if (start === undefined) {
      // Not placed (a floating item, say): the markers go with the next block.
      continue;
    }
    set(pending, start);
    pending = [];
    if (item.type === "table") {
      const end = lastPage.get(i) ?? start;
      item.rows.forEach((row, r) => {
        set(row.rangeMarkersBefore, cellAt(i, item, r, 0) ?? end);
        row.cells.forEach((cell, c) => set(cell.rangeMarkersBefore, cellAt(i, item, r, c) ?? end));
        set(row.rangeMarkersAfter, cellAt(i, item, r + 1, 0) ?? end);
      });
      set(item.rangeMarkersAfter, end);
    }
  }
  pending.push(...(bodyAfter ?? []));
  set(pending, finalPage);
}

// =============================================================================
// Internal: Page Building
// =============================================================================

/**
 * Per-line wrap exclusion zone (a horizontal band that text must avoid).
 *
 * `xLeft`/`xRight` are content-area-relative coordinates: 0 is the
 * left edge of the content area, `contentWidth` the right edge.
 * `yTop`/`yBottom` are relative to the top of the page's content area.
 * `wrapSide` mirrors ECMA-376's `<wp:wrapSquare wrapText="…">`:
 * `"left"` means text wraps on the float's left side only (i.e. the
 * float blocks the right portion of every line it intersects); `"right"`
 * is the mirror; `"bothSides"` blocks the float's exact horizontal
 * extent and lets text flow both to its left and right; `"largest"`
 * picks whichever side is wider on each line.
 */
/** The half-width of a top-and-bottom float's band: wider than any page, so no text fits beside it. */
const FULL_WIDTH_BAND_PT = 1e6;

interface WrapExclusion {
  readonly xLeft: number;
  readonly xRight: number;
  readonly yTop: number;
  readonly yBottom: number;
  readonly wrapSide: "left" | "right" | "bothSides" | "largest";
}

/**
 * Page-scoped context threaded through `layoutParagraph` so each line
 * can avoid the exclusion zones declared by floats that come earlier
 * on the page. Floats with `wrap.style ∈ { "square" | "tight" |
 * "through" }` populate this; other styles are handled at the
 * cursor-advancement layer in `buildPage`.
 */
interface PageLayoutContext {
  readonly exclusions: readonly WrapExclusion[];
  /** Content-area width — `contentWidth` for the page. */
  readonly contentWidth: number;
}

/**
 * Compute the horizontal slots text may take on the line whose vertical
 * span is `[lineY, lineY + lineHeight)`, left to right. Each has an
 * `xOffset` (relative to the content-area's left edge) and `width`. When
 * no exclusion intersects the line the result is one slot, `{ xOffset: 0,
 * width: contentWidth }` (full width).
 *
 * Algorithm:
 *  1. Collect all exclusions whose y-band intersects the line.
 *  2. For each, derive the "blocked" x-interval on the content axis
 *     based on `wrapSide`:
 *      - `bothSides` blocks `[xLeft, xRight]` only.
 *      - `left` blocks `[xLeft, contentWidth]` (text wraps on the
 *         float's left side only).
 *      - `right` blocks `[0, xRight]`.
 *      - `largest` picks whichever side of the float is wider; the
 *         narrower side is blocked.
 *  3. Subtract every blocked interval from `[0, contentWidth]` and
 *     return the remaining gaps when a `bothSides` float takes part —
 *     Word fills the gap left of it, then the one right of it — and the
 *     longest of them otherwise.
 */
function availableSlotsForLine(
  ctx: PageLayoutContext,
  lineY: number,
  lineHeight: number
): { xOffset: number; width: number }[] {
  const lineBottom = lineY + lineHeight;
  const blocked: { lo: number; hi: number }[] = [];
  // Word sets text on both sides of a `bothSides` float: the line fills the
  // gap left of it, then continues in the gap right of it (pages2 R1c).
  let bothSides = false;
  for (const ex of ctx.exclusions) {
    // Strict overlap check: a line that just touches the float's
    // bottom edge (`lineY === ex.yBottom`) does NOT need to wrap.
    if (lineBottom <= ex.yTop || lineY >= ex.yBottom) {
      continue;
    }
    const exLeft = Math.max(0, ex.xLeft);
    const exRight = Math.min(ctx.contentWidth, ex.xRight);
    if (exLeft >= exRight) {
      continue;
    }
    switch (ex.wrapSide) {
      case "bothSides":
        blocked.push({ lo: exLeft, hi: exRight });
        bothSides = true;
        break;
      case "left":
        // Float blocks the right half of the line.
        blocked.push({ lo: exLeft, hi: ctx.contentWidth });
        break;
      case "right":
        blocked.push({ lo: 0, hi: exRight });
        break;
      case "largest": {
        const leftSpace = exLeft;
        const rightSpace = ctx.contentWidth - exRight;
        if (rightSpace >= leftSpace) {
          // Wrap on the right (block to the left of the float).
          blocked.push({ lo: 0, hi: exRight });
        } else {
          blocked.push({ lo: exLeft, hi: ctx.contentWidth });
        }
        break;
      }
    }
  }

  if (blocked.length === 0) {
    return [{ xOffset: 0, width: ctx.contentWidth }];
  }

  // Merge overlapping blocked intervals.
  blocked.sort((a, b) => a.lo - b.lo);
  const merged: { lo: number; hi: number }[] = [];
  for (const seg of blocked) {
    const last = merged[merged.length - 1];
    if (last && seg.lo <= last.hi) {
      last.hi = Math.max(last.hi, seg.hi);
    } else {
      merged.push({ lo: seg.lo, hi: seg.hi });
    }
  }

  // Build available gaps in [0, contentWidth] minus the merged blocks.
  const gaps: { x: number; width: number }[] = [];
  let cursor = 0;
  for (const seg of merged) {
    if (seg.lo > cursor) {
      gaps.push({ x: cursor, width: seg.lo - cursor });
    }
    cursor = Math.max(cursor, seg.hi);
  }
  if (cursor < ctx.contentWidth) {
    gaps.push({ x: cursor, width: ctx.contentWidth - cursor });
  }

  if (gaps.length === 0) {
    // Line entirely blocked: no room on this row. Callers leave the row empty
    // and try the next, so the text resumes below the float (pages2 R3g: a
    // `wrapText="left"` picture at the left margin pushes the text under it).
    return [{ xOffset: 0, width: 0 }];
  }
  if (bothSides) {
    return gaps.map(gap => ({ xOffset: gap.x, width: gap.width }));
  }
  // Pick the widest gap.
  let best = gaps[0];
  for (let i = 1; i < gaps.length; i++) {
    if (gaps[i].width > best.width) {
      best = gaps[i];
    }
  }
  return [{ xOffset: best.x, width: best.width }];
}

/**
 * Build a single page. Returns the laid `LayoutPage` plus the footnotes, or
 * parts of footnotes, its footnote area had no room for. Callers thread them
 * into the next page as `pendingNotes`, where they open its footnote area.
 */
interface BuildPageResult {
  readonly page: LayoutPage;
  readonly continuedNotes: readonly LaidNote[];
  /**
   * Blocks (or block remainders) that did not fit and must start the next page.
   *
   * The paginator in `layout.ts` estimates heights independently and cannot
   * split a block at all — its item→page map holds exactly one page per body
   * item. Pass 2 has the real measurements, so it owns the final fit decision:
   * anything that would cross the bottom margin is split here and carried
   * forward. Without this, a table or a paragraph taller than the space left
   * simply ran off the page.
   */
  readonly carriedContent: readonly CarriedBlock[];
  /** How far down the body text reached, content-relative. */
  readonly bottom: number;
  /**
   * Why the flow stopped short of the page's end: a column break, a
   * continuous section set in other columns that starts a region of its own,
   * or a page break or a section starting a page. Absent when the page (or
   * column) is simply full or the text has run out.
   */
  readonly stop?: "column" | "region" | "page";
  /** The page's notes, which `composePage` lays out once its columns are done. */
  readonly notes: PageNoteSet;
  /**
   * The section the page ends in, when a continuous section break on it
   * started a later one than the page began with.
   */
  readonly endSection?: number;
}

/** The notes collected for one page; see `createPageNotes`. */
type PageNoteSet = ReturnType<typeof createPageNotes>;

/**
 * Part of a page that body text fills: a column, below whatever regions of
 * earlier sections the page already holds. Content-relative.
 */
interface PageRegion {
  readonly top: number;
  /** The lowest the text may reach, when lower than the page allows (balancing). */
  readonly bottom?: number;
  readonly x: number;
  readonly width: number;
  /** Index of the column within its section, from 0. */
  readonly column: number;
  /** The section whose text fills it. */
  readonly sectionIndex: number;
  /**
   * The least room the page's footnote area takes: the notes of later columns
   * share one area with those of earlier ones, so an earlier column ends above
   * all of them (pages2 R9e).
   */
  readonly noteFloor?: number;
  /**
   * Notes whose reference may not join the page: making room for them moved
   * their reference off it (pages2 R11a), so it goes to the next page instead.
   */
  readonly refusedNotes?: ReadonlySet<number>;
  readonly props: SectionProperties | undefined;
}

/** A laid-out block awaiting placement at the top of the next page. */
type CarriedBlock =
  | {
      readonly kind: "paragraph";
      readonly block: LayoutParagraph;
      /** The rest of a paragraph whose first lines were set earlier. */
      readonly continued?: boolean;
    }
  | {
      readonly kind: "table";
      readonly block: LayoutTable;
      readonly headerRows: number;
      /** Notes referenced in the table, while no part of it has been placed. */
      readonly noteIds: readonly number[];
      readonly splitInfo: TableSplitInfo;
    };

/**
 * How far through `doc.body` page building has got, threaded from page to page.
 *
 * Pass 2 measures for real, so it — not the estimating paginator — decides which
 * page each block lands on. It walks the body once, filling each page until the
 * next block does not fit, and only honours the paginator's *forced* breaks.
 * Selecting items by the paginator's page number instead meant every
 * disagreement between the two showed up in the output: content overflowed the
 * bottom margin where the estimate was too optimistic, and pages stopped early,
 * leaving a blank gap, where it was too pessimistic.
 */
interface FlowCursor {
  /** Index of the next body item to place. */
  nextItem: number;
  /** Last paragraph in a Quote run that must end its page once fully placed. */
  endPageAfterItem?: number;
}

/**
 * Split a laid-out paragraph at a page boundary.
 *
 * `availableHeight` is the space left on the page, measured from the
 * paragraph's own top (`rect.y`) — per line when it depends on the line, as it
 * does when a line's footnotes have to fit below it. Lines that fit stay in `head`; the rest are
 * re-based to the top of the next page and returned as `tail`. `head` drops its
 * space-after (it does not end the paragraph) and `tail` drops its space-before.
 *
 * Word's widow/orphan control forbids leaving a single line stranded, so a
 * paragraph is only split where both sides keep at least two lines — unless the
 * paragraph is taller than a whole page, where refusing to split would put it
 * off the page altogether.
 *
 * `keepSpaceAfter` says the space after the paragraph has to fit with its last
 * line. Word drops it only where the paragraph ends at the bottom margin. In a
 * table cell it moves the lines instead (final6 Qc s04–s10: a three-line cell
 * paragraph goes to the next page while the paragraph above it stays); in a
 * footnote area, or above one, it moves them too (final6 Qb1: a note paragraph
 * that fits only without its space after continues on the next page; Qb2: so
 * does the paragraph referencing a note).
 *
 * `atPageTop` says the paragraph starts a fresh page, so moving it down again
 * would be pointless: at least one line is emitted even if it does not fit,
 * which is both what Word does with an over-tall line and what guarantees the
 * page loop makes progress instead of carrying the same block forever.
 */
function splitLayoutParagraph(
  para: LayoutParagraph,
  availableHeight: number | ((line: number) => number),
  fullPageHeight: number,
  atPageTop: boolean,
  rules: ParagraphBreakRules,
  keepSpaceAfter = false,
  continued = false
): { head: LayoutParagraph | null; tail: LayoutParagraph | null } {
  const availableAt = typeof availableHeight === "number" ? () => availableHeight : availableHeight;
  const lines = para.lines;
  if (lines.length === 0) {
    return atPageTop ? { head: para, tail: null } : { head: null, tail: para };
  }

  // How many leading lines fit in the space left. A line's auto-spacing slack
  // below its text may hang into the bottom margin, but not where the space
  // after has to fit too: above a footnote area it has to fit as well (pages2
  // R6g: a line 0.27pt short of fitting without its 1.07pt slack moves on).
  // The last line fits only with the box's bottom rule under its whole height
  // (pages2 R11d: a bordered paragraph whose last line would end 0.3pt above
  // the foot without its rule breaks 2/2 instead).
  const bottomRule =
    (para.decorationInsets?.bottom ?? 0) > 0 ? borderLeadPt(para.borders?.bottom, false) : 0;
  let fitCount = 0;
  while (fitCount < lines.length) {
    const line = lines[fitCount];
    const ruled = bottomRule > 0 && fitCount === lines.length - 1;
    const slack = keepSpaceAfter || ruled ? 0 : (line.bottomSlack ?? 0);
    if (line.y + line.height - slack + (ruled ? bottomRule : 0) > availableAt(fitCount)) {
      break;
    }
    fitCount++;
  }
  // With `keepSpaceAfter` the last line only fits with the space after it.
  if (
    keepSpaceAfter &&
    fitCount === lines.length &&
    para.rect.height > availableAt(fitCount - 1) + FIT_EPSILON_PT
  ) {
    fitCount--;
  }

  // A `w:br w:type="page"` inside the paragraph overrides fit: everything up to
  // and including its line stays, the rest continues on the next page. Widow
  // control does not apply — the author asked for the break.
  const forced = para.pageBreakAfterLines?.find(index => index + 1 <= fitCount);
  if (forced !== undefined && forced + 1 < lines.length) {
    const cut = sliceParagraphAtLine(para, forced + 1);
    // Nothing after the break: the paragraph mark stays with it (see `placeParagraph`).
    return cut.tail.lines.some(line => line.runs.length > 0) ? cut : { head: cut.head, tail: null };
  }

  if (fitCount >= lines.length) {
    // Every line fits and only the trailing space-after spills. Word does not
    // render space-after across a page break, so trim the box to its last line
    // rather than pushing a paragraph whose text fits onto the next page.
    const last = lines[lines.length - 1];
    // A bottom rule ends the paragraph's box; only the space after is trimmed.
    const rule = para.decorationInsets ? borderLeadPt(para.borders?.bottom, false) : 0;
    const contentBottom = last.y + last.height + rule;
    return {
      head:
        para.rect.height > availableAt(lines.length - 1)
          ? {
              ...para,
              rect: { ...para.rect, height: contentBottom },
              // The trimmed box ends at the text and its rule.
              ...(para.decorationInsets
                ? { decorationInsets: { ...para.decorationInsets, bottom: rule } }
                : {})
            }
          : para,
      tail: null
    };
  }

  if (fitCount === 0) {
    if (!atPageTop) {
      return { head: null, tail: para };
    }
    // An over-tall line still has to go somewhere: at the top of a page there is
    // no emptier page to move it to, so it stays and overflows.
    //
    // When it is the paragraph's *only* line there is nothing left to continue on
    // the next page, and asking for a cut after line 1 of 1 produced an empty
    // tail — whose first line was then read for the shift (`tailLines[0].y`) and
    // threw `Cannot read properties of undefined`. A page barely taller than one
    // line of text is enough to reach it.
    if (lines.length === 1) {
      return { head: para, tail: null };
    }
    fitCount = 1;
  } else if (!atPageTop && para.rect.height <= fullPageHeight) {
    // `w:keepLines` moves the whole paragraph. `w:widowControl` leaves at
    // least two lines on each page, moving a line down when only one would
    // follow; without it Word breaks after whatever fits, one line included.
    if (rules.keepLines) {
      return { head: null, tail: para };
    }
    if (rules.widowControl) {
      if (lines.length - fitCount < 2) {
        fitCount = lines.length - 2;
      }
      if (fitCount < 2) {
        return { head: null, tail: para };
      }
    }
  } else if (rules.widowControl && lines.length - fitCount < 2 && (fitCount > 2 || continued)) {
    // A paragraph that has to break where it starts — at the top of a page or
    // column — still leaves no line on its own after the break (pages2 P12b: a
    // six-line paragraph opening a balanced column breaks four and two). The
    // rest of a paragraph has no first line to keep company, so its last two
    // lines stay together even when that is all of it (pages2 R12j: two lines
    // continued at a page top are not balanced one per column).
    fitCount = lines.length - 2;
    if (fitCount < 1) {
      return { head: null, tail: para };
    }
  }

  return sliceParagraphAtLine(para, fitCount);
}

/** How a paragraph may break across pages (`w:widowControl`, `w:keepLines`). */
interface ParagraphBreakRules {
  readonly widowControl: boolean;
  readonly keepLines: boolean;
}

/**
 * Cut a laid-out paragraph so that `count` lines stay and the rest continue at
 * the top of the next page.
 *
 * The head keeps its space-before but loses its space-after (it does not end the
 * paragraph); the tail is the mirror image. Insets are relative to `rect`, so
 * adjusting them is all the decoration needs. Per-line note ids are sliced
 * alongside the lines so each note follows its own reference.
 */
function sliceParagraphAtLine(
  para: LayoutParagraph,
  count: number
): { head: LayoutParagraph; tail: LayoutParagraph } {
  const headLines = para.lines.slice(0, count);
  const tailLines = para.lines.slice(count);
  const headBottom = headLines[headLines.length - 1].y + headLines[headLines.length - 1].height;
  const shift = tailLines[0].y;
  const insets = para.decorationInsets;
  const noteIds = para.lineNoteIds;
  const bookmarks = para.lineBookmarks;
  const breaks = para.pageBreakAfterLines;
  const columnBreaks = para.columnBreakAfterLines;
  return {
    head: {
      ...para,
      rect: { ...para.rect, height: headBottom },
      lines: headLines,
      ...(noteIds ? { lineNoteIds: noteIds.slice(0, count) } : {}),
      ...(bookmarks ? { lineBookmarks: bookmarks.slice(0, count) } : {}),
      ...(breaks ? { pageBreakAfterLines: breaks.filter(i => i < count) } : {}),
      ...(columnBreaks ? { columnBreakAfterLines: columnBreaks.filter(i => i < count) } : {}),
      ...(insets ? { decorationInsets: { ...insets, bottom: 0 } } : {})
    },
    tail: {
      ...para,
      rect: { ...para.rect, y: 0, height: para.rect.height - shift },
      lines: tailLines.map(line => ({ ...line, y: line.y - shift })),
      ...(noteIds ? { lineNoteIds: noteIds.slice(count) } : {}),
      ...(bookmarks ? { lineBookmarks: bookmarks.slice(count) } : {}),
      ...(breaks
        ? { pageBreakAfterLines: breaks.filter(i => i >= count).map(i => i - count) }
        : {}),
      ...(columnBreaks
        ? {
            columnBreakAfterLines: columnBreaks.filter(i => i >= count).map(i => i - count)
          }
        : {}),
      ...(insets ? { decorationInsets: { ...insets, top: 0 } } : {})
    }
  };
}

/**
 * What splitting a laid-out table needs to know about its source: which rows
 * may break, and for each cell its insets and how each content block breaks.
 */
interface TableSplitInfo {
  /** Rows Word does not break: `w:cantSplit`, or an exact `w:trHeight`. */
  readonly unbreakable: ReadonlySet<number>;
  /** Rows kept on the page of the row below them (see `rowKeepsWithNext`). */
  readonly keptWithNext?: ReadonlySet<number>;
  /** By `row:col` (the layout's cell coordinates). */
  readonly cells: ReadonlyMap<string, CellSplitInfo>;
}

interface CellSplitInfo {
  /** Space above the content: the top margin and the rule above the row. */
  readonly top: number;
  /** Space below the content: the bottom margin and, in the last row, the rule. */
  readonly bottom: number;
  /**
   * Space below the content of the piece a page break cuts off: the bottom
   * margin and the rule that closes the row there — the bottom border in the
   * last row, the inside border in any other. Word keeps a line of a broken row
   * on the page only if that rule fits below it as well (final7 N6e: row k5's
   * second paragraph, which fits only without its 0.5pt rule, moves on).
   */
  readonly cutBottom: number;
  /** Space above the content when the row continues on a new page. */
  readonly carriedTop: number;
  /** Room below the row when it is the last on its page and the table goes on. */
  readonly closing: number;
  /** Parallel to the laid cell's `content`. */
  readonly blocks: readonly CellBlockSplitInfo[];
  /** `w:vAlign`, which a merged cell cut by a page break applies to each piece. */
  readonly verticalAlign?: VerticalCellAlign;
}

type CellBlockSplitInfo =
  | { readonly kind: "paragraph"; readonly rules: ParagraphBreakRules }
  | { readonly kind: "table"; readonly headerRows: number; readonly info: TableSplitInfo };

/**
 * Split a laid-out table at a page boundary.
 *
 * Rows that fit stay; Word then breaks the first row that does not, line by
 * line inside each of its cells, unless the row says it cannot split
 * (`w:cantSplit`, or an exact height) — such a row moves to the next page whole
 * (Word for Mac 16.114, final5 M2-nonest: a two-cell row breaks across pages 4
 * and 5). A cell's paragraphs keep their own widow/orphan and keep-lines rules,
 * and a nested table breaks on its own rows the same way. Rows marked
 * `tableHeader` repeat at the top of the continuation, and never split.
 *
 * `atPageTop` has the same role as in `splitLayoutParagraph`: on a fresh page
 * at least one row is emitted, so a row taller than the page overflows rather
 * than being carried forever.
 */
function splitLayoutTable(
  table: LayoutTable,
  availableHeight: number,
  headerRowCount: number,
  atPageTop: boolean,
  info?: TableSplitInfo,
  fullPageHeight = Infinity
): { head: LayoutTable | null; tail: LayoutTable | null } {
  if (table.cells.length === 0) {
    return { head: table, tail: null };
  }

  // Row boundaries in table-relative coordinates.
  const rowTop = new Map<number, number>();
  const rowBottom = new Map<number, number>();
  for (const cell of table.cells) {
    // A merged cell spans rows; the rows are drawn by the cells of one row.
    if ((cell.rowSpan ?? 1) > 1) {
      continue;
    }
    const top = cell.rect.y;
    const bottom = top + cell.rect.height;
    rowTop.set(cell.row, Math.min(rowTop.get(cell.row) ?? top, top));
    rowBottom.set(cell.row, Math.max(rowBottom.get(cell.row) ?? bottom, bottom));
  }
  const rows = [...rowTop.keys()].sort((a, b) => a - b);

  // First row index (in `rows`) that does not fit in the space left. With cell
  // spacing the piece left on the page closes with s and the table's bottom
  // rule, which have to fit too (pages2 P24a).
  const closing = (row: number): number => {
    return info?.cells.get(`${row}:0`)?.closing ?? 0;
  };
  let fitRows = 0;
  while (
    fitRows < rows.length &&
    (rowBottom.get(rows[fitRows]) ?? 0) +
      (fitRows === rows.length - 1 ? 0 : closing(rows[fitRows]!)) <=
      availableHeight + FIT_EPSILON_PT
  ) {
    fitRows++;
  }

  if (fitRows >= rows.length) {
    return { head: table, tail: null };
  }

  // Break the row that does not fit, when it may break and some of it fits.
  // A row kept with it is then not left without it (pages2 P29f).
  const breaking = rows[fitRows]!;
  if (info && breaking >= headerRowCount && !info.unbreakable.has(breaking)) {
    const split = splitTableRow(
      table,
      breaking,
      availableHeight - (rowTop.get(breaking) ?? 0),
      info,
      fullPageHeight
    );
    if (split) {
      return assembleTableSplit(table, breaking, split, headerRowCount, rowTop, info);
    }
  }

  // A row kept with the next is not left at the foot of a page without it.
  const kept = info?.keptWithNext;
  if (kept && kept.size > 0) {
    while (fitRows > 0 && kept.has(rows[fitRows - 1]!)) {
      fitRows--;
    }
  }

  if (fitRows === 0 || (fitRows <= headerRowCount && !atPageTop)) {
    // Nothing, or nothing but repeated header rows, would land here. Move the
    // table down rather than leave a dangling header — unless this is already a
    // fresh page, where at least one row has to be emitted.
    if (!atPageTop) {
      return { head: null, tail: table };
    }
    fitRows = Math.max(1, Math.min(headerRowCount + 1, rows.length));
  }

  const splitRow = rows[fitRows]!;
  if (splitRow === undefined) {
    return { head: table, tail: null };
  }
  return assembleTableSplit(table, splitRow, undefined, headerRowCount, rowTop, info);
}

/** The two pieces of one row broken across a page boundary. */
interface RowSplit {
  readonly head: readonly LayoutTableCell[];
  readonly headHeight: number;
  readonly tail: readonly LayoutTableCell[];
  readonly tailHeight: number;
}

/**
 * The head and tail of a table cut before row `splitRow` — or through it,
 * when `rowSplit` holds its two pieces. The continuation repeats the header
 * rows, then continues from the cut.
 */
function assembleTableSplit(
  table: LayoutTable,
  splitRow: number,
  rowSplit: RowSplit | undefined,
  headerRowCount: number,
  rowTop: ReadonlyMap<number, number>,
  info: TableSplitInfo | undefined
): { head: LayoutTable; tail: LayoutTable } {
  const top = rowTop.get(splitRow) ?? 0;
  // A merge running across the cut: in a broken row `rowSplit` holds its pieces;
  // at a row edge it is cut here, its text running on below the cut.
  const across = table.cells.filter(c => c.row < splitRow && coversRow(c, splitRow));
  const acrossPieces = rowSplit ? [] : across.map(c => cutMergedCell(c, top, splitRow, info));
  const headCells = [
    ...table.cells.filter(c => c.row < splitRow && !across.includes(c)),
    ...acrossPieces.map(p => p.head)
  ];
  const headHeight = rowSplit ? top + rowSplit.headHeight : top;

  const repeated = table.cells.filter(c => c.row < headerRowCount && c.row < splitRow);
  const repeatedHeight =
    repeated.length > 0 ? Math.max(...repeated.map(c => c.rect.y + c.rect.height)) : 0;
  const rest = table.cells.filter(c => (rowSplit ? c.row > splitRow : c.row >= splitRow));
  const restTop = rowSplit ? repeatedHeight + rowSplit.tailHeight : repeatedHeight;
  const nextTop = rest.length > 0 ? Math.min(...rest.map(c => c.rect.y)) : table.rect.height;
  const shift = nextTop - restTop;
  // A row opening the continuation (no header row repeated above it) takes the
  // room a continued row has above it, not its own (see `rowRulesPt`).
  const splitInfo = info?.cells.get(`${splitRow}:0`);
  const carriedDelta =
    !rowSplit && repeated.length === 0 && splitInfo ? splitInfo.carriedTop - splitInfo.top : 0;

  const tailCells = [
    // Repeated header rows already sit at the top of the table.
    ...repeated,
    ...(rowSplit?.tail ?? []).map(c => ({ ...c, rect: { ...c.rect, y: repeatedHeight } })),
    ...rest.map(c => {
      const y = c.rect.y - shift;
      if (c.row !== splitRow) {
        return { ...c, rect: { ...c.rect, y: y + carriedDelta } };
      }
      return {
        ...c,
        rect: { ...c.rect, y, height: c.rect.height + carriedDelta },
        content: c.content.map(b => ({ ...b, rect: { ...b.rect, y: b.rect.y + carriedDelta } }))
      };
    }),
    ...acrossPieces.map(p => ({ ...p.tail, rect: { ...p.tail.rect, y: restTop } }))
  ];
  const tail = fitMergedCells(tailCells);
  for (const piece of acrossPieces) {
    const laid = tail.cells.find(c => c.content === piece.tail.content);
    if (laid && piece.align) {
      (laid as { content: readonly LayoutBlock[] }).content = alignBlocks(
        [...laid.content],
        piece.align,
        laid.rect.height - piece.bottomInset
      );
    }
  }

  return {
    head: {
      ...table,
      rect: { ...table.rect, height: headHeight },
      cells: headCells
        .filter(c => !rowSplit || !coversRow(c, splitRow))
        .concat(rowSplit?.head ?? [])
    },
    tail: {
      ...table,
      rect: { ...table.rect, y: 0, height: table.rect.height - shift + carriedDelta + tail.grown },
      cells: tail.cells
    }
  };
}

/**
 * A merged cell cut at `cut` (table-relative), which falls on the edge of
 * `row`: what stays above, and the rest as a merged cell from `row` down.
 */
function cutMergedCell(
  cell: LayoutTableCell,
  cut: number,
  row: number,
  info: TableSplitInfo | undefined
): {
  head: LayoutTableCell;
  tail: LayoutTableCell;
  align: VerticalCellAlign | undefined;
  bottomInset: number;
} {
  const cellInfo = info?.cells.get(`${cell.row}:${cell.col}`);
  const limit = cut - cell.rect.y - (cellInfo?.cutBottom ?? 0);
  // Each piece is aligned on its own page (pages2 P9: two centred lines of a
  // merge over four rows stay centred in the two rows above the break).
  const content = cellInfo ? stackBlocks(cell.content, cellInfo.top) : cell.content;
  const pieces =
    cellInfo && cellInfo.blocks.length === content.length
      ? splitCellContent(content, limit, cellInfo.blocks, Infinity)
      : {
          head: content.filter(b => b.rect.y + b.rect.height <= limit + FIT_EPSILON_PT),
          tail: content.filter(b => b.rect.y + b.rect.height > limit + FIT_EPSILON_PT)
        };
  const span = cell.row + (cell.rowSpan ?? 1) - row;
  // The rest starts below the top margin of the row it continues in (pages2
  // R7i: 6pt there, none where the merge began).
  const carriedTop = info?.cells.get(`${row}:${cell.col}`)?.carriedTop ?? cellInfo?.carriedTop ?? 0;
  const tailContent = stackBlocks(pieces.tail, carriedTop);
  return {
    head: {
      ...cell,
      rect: { ...cell.rect, height: cut - cell.rect.y },
      content: alignBlocks(pieces.head, cellInfo?.verticalAlign, limit)
    },
    tail: {
      ...cell,
      row,
      rowSpan: span,
      rect: {
        ...cell.rect,
        height: (tailContent.length > 0 ? blockBottom(tailContent) : 0) + (cellInfo?.bottom ?? 0)
      },
      content: tailContent
    },
    align: cellInfo?.verticalAlign,
    bottomInset: cellInfo?.bottom ?? 0
  };
}

/** `blocks` moved down by the room left above `bottom`, as `w:vAlign` asks. */
function alignBlocks(
  blocks: LayoutBlock[],
  verticalAlign: VerticalCellAlign | undefined,
  bottom: number
): LayoutBlock[] {
  if (blocks.length === 0 || !verticalAlign || verticalAlign === "top") {
    return blocks;
  }
  const slack = bottom - blockBottom(blocks);
  if (slack <= 0) {
    return blocks;
  }
  const shift = verticalAlign === "center" ? slack / 2 : slack;
  return blocks.map(b => ({ ...b, rect: { ...b.rect, y: b.rect.y + shift } }));
}

/**
 * `cells` with the last row of every merge grown, and everything below it moved
 * down, where the merged cell needs more room than its rows give it; merged
 * cells cover their rows. Returns how much the table grew.
 */
function fitMergedCells(cells: readonly LayoutTableCell[]): {
  cells: LayoutTableCell[];
  grown: number;
} {
  const out = cells.map(c => ({ ...c, rect: { ...c.rect } }));
  let grown = 0;
  for (const merged of out.filter(c => c.rowSpan !== undefined)) {
    const last = merged.row + (merged.rowSpan ?? 1) - 1;
    const rowCells = out.filter(c => c.row === last && c.rowSpan === undefined);
    if (rowCells.length === 0) {
      continue;
    }
    const rowBottom = Math.max(...rowCells.map(c => c.rect.y + c.rect.height));
    const need = merged.rect.y + merged.rect.height - rowBottom;
    if (need > 0) {
      for (const c of rowCells) {
        c.rect.height += need;
      }
      for (const c of out) {
        if (c.row > last) {
          c.rect.y += need;
        }
      }
      grown += need;
    }
    merged.rect.height = Math.max(...rowCells.map(c => c.rect.y + c.rect.height)) - merged.rect.y;
  }
  return { cells: out, grown };
}

/**
 * Break row `row` so its first piece is at most `available` tall, or
 * `undefined` when no cell can keep anything on this page.
 */
function splitTableRow(
  table: LayoutTable,
  row: number,
  available: number,
  info: TableSplitInfo,
  fullPageHeight: number
): RowSplit | undefined {
  const rowTop = Math.min(
    ...table.cells.filter(c => c.row === row && (c.rowSpan ?? 1) === 1).map(c => c.rect.y)
  );
  // The row's own cells, and merged cells from above that run through it.
  const cells = table.cells.filter(c => coversRow(c, row));
  const pieces: {
    cell: LayoutTableCell;
    /** How far above the row the cell starts. */
    above: number;
    head: LayoutBlock[];
    tail: LayoutBlock[];
    insets: CellSplitInfo;
  }[] = [];
  let anyHead = false;
  for (const cell of cells) {
    const cellInfo = info.cells.get(`${cell.row}:${cell.col}`);
    if (!cellInfo || cellInfo.blocks.length !== cell.content.length) {
      return undefined;
    }
    const above = Number.isFinite(rowTop) ? rowTop - cell.rect.y : 0;
    const cut = splitCellContent(
      cell.content,
      above + available - cellInfo.cutBottom,
      cellInfo.blocks,
      fullPageHeight
    );
    // Only text of the row's own that stays here makes the cut worth it — not
    // a continuing merged cell's, which sits up where the merge begins.
    anyHead ||= cell.row === row && cut.head.some(b => b.rect.y >= 0);
    pieces.push({ cell, above, head: cut.head, tail: cut.tail, insets: cellInfo });
  }
  // Word breaks a row only where every cell of it keeps something on this page:
  // a cell whose text cannot start here (say a three-line paragraph under widow
  // control with room for one line) takes the whole row to the next page
  // (pages2 R1d: row r49, one line of room, moves whole).
  const stranded = pieces.some(
    p => p.cell.row === row && p.tail.length > 0 && !p.head.some(b => b.rect.y >= 0)
  );
  if (!anyHead || stranded || pieces.every(p => p.tail.length === 0)) {
    return undefined;
  }
  let headHeight = 0;
  let tailHeight = 0;
  const tails: { cell: LayoutTableCell; content: LayoutBlock[]; last: boolean }[] = [];
  for (const { cell, above, head, tail, insets } of pieces) {
    const headBottom = head.length > 0 ? blockBottom(head) : insets.top;
    headHeight = Math.max(headHeight, headBottom + insets.cutBottom - above);
    const restacked = stackBlocks(tail, insets.carriedTop);
    // A merge running on below this row sizes the rows below, not this one.
    const last = cell.row + (cell.rowSpan ?? 1) - 1 === row;
    if (last) {
      tailHeight = Math.max(
        tailHeight,
        (restacked.length > 0 ? blockBottom(restacked) : insets.carriedTop) + insets.bottom
      );
    }
    tails.push({ cell, content: restacked, last });
  }
  return {
    head: pieces.map(({ cell, above, head }) => ({
      ...cell,
      rect: { ...cell.rect, height: above + headHeight },
      content: head
    })),
    headHeight,
    tail: tails.map(({ cell, content, last }) => ({
      ...cell,
      row,
      rect: { ...cell.rect, height: tailHeight },
      content,
      ...(last ? { rowSpan: undefined } : { rowSpan: cell.row + (cell.rowSpan ?? 1) - row })
    })),
    tailHeight
  };
}

/** Whether `cell` occupies `row`: it is one of the row's, or a merge runs through it. */
function coversRow(cell: LayoutTableCell, row: number): boolean {
  return cell.row <= row && row < cell.row + (cell.rowSpan ?? 1);
}

type LayoutBlock = LayoutParagraph | LayoutTable;

/** The lowest edge of `blocks`. */
function blockBottom(blocks: readonly LayoutBlock[]): number {
  return Math.max(...blocks.map(b => b.rect.y + b.rect.height));
}

/** `blocks` stacked one after another from `top`. */
function stackBlocks(blocks: readonly LayoutBlock[], top: number): LayoutBlock[] {
  const out: LayoutBlock[] = [];
  let y = top;
  for (const block of blocks) {
    out.push({ ...block, rect: { ...block.rect, y } });
    y += block.rect.height;
  }
  return out;
}

/**
 * Cut a cell's content so what stays ends by `limit` (cell-relative). Blocks
 * that fit stay whole; the first that does not breaks where its own rules
 * allow, and it and everything after it continue in the next piece.
 */
function splitCellContent(
  content: readonly LayoutBlock[],
  limit: number,
  blocks: readonly CellBlockSplitInfo[],
  fullPageHeight: number
): { head: LayoutBlock[]; tail: LayoutBlock[] } {
  const head: LayoutBlock[] = [];
  const tail: LayoutBlock[] = [];
  for (let k = 0; k < content.length; k++) {
    const block = content[k]!;
    if (tail.length > 0) {
      tail.push(block);
      continue;
    }
    if (block.rect.y + block.rect.height <= limit + FIT_EPSILON_PT) {
      head.push(block);
      continue;
    }
    const blockInfo = blocks[k]!;
    const room = limit - block.rect.y;
    const cut =
      block.type === "paragraph" && blockInfo.kind === "paragraph"
        ? splitLayoutParagraph(block, room, fullPageHeight, false, blockInfo.rules, true)
        : block.type === "table" && blockInfo.kind === "table"
          ? splitLayoutTable(
              block,
              room,
              blockInfo.headerRows,
              false,
              blockInfo.info,
              fullPageHeight
            )
          : { head: null, tail: block };
    if (cut.head) {
      head.push({ ...cut.head, rect: { ...cut.head.rect, y: block.rect.y } } as LayoutBlock);
    }
    if (cut.tail) {
      tail.push(cut.tail);
    }
  }
  // A vertical merge's continuing cells' content sits, taking no room, where
  // the merged cell's text begins, and goes where that text goes (pages2 R7i:
  // the merged cell's two lines move to the next page, and so does the
  // continuing cell's bookmark).
  while (tail.length > 0 && head.length > 0 && head.at(-1)!.rect.height === 0) {
    const moved = head.pop()!;
    tail.unshift(moved);
  }
  return { head, tail };
}

/** {@link TableSplitInfo} for `source`, read the way {@link layoutTable} reads it. */
function tableSplitInfo(source: Table): TableSplitInfo {
  const table = withEffectiveProperties(source);
  const unbreakable = new Set<number>();
  const cells = new Map<string, CellSplitInfo>();
  // Continuing cells' blocks travel with the cell that starts their merge (see
  // `layoutTableOffGrid`), at the front of its content.
  const rowSpans = verticalMergeSpans(table);
  const continuing = (ri: number, ci: number): CellBlockSplitInfo[] => {
    const span = rowSpans.get(`${ri}:${ci}`) ?? 1;
    if (span <= 1) {
      return [];
    }
    let col = table.rows[ri]!.properties?.gridBefore ?? 0;
    for (let k = 0; k < ci; k++) {
      col += Math.max(1, table.rows[ri]!.cells[k]!.properties?.gridSpan ?? 1);
    }
    const out: CellBlockSplitInfo[] = [];
    for (let r = ri + 1; r < ri + span; r++) {
      let at = table.rows[r]!.properties?.gridBefore ?? 0;
      for (const cell of table.rows[r]!.cells) {
        if (at === col) {
          out.push(...cellBlockInfo(cell));
          break;
        }
        at += Math.max(1, cell.properties?.gridSpan ?? 1);
      }
    }
    return out;
  };
  table.rows.forEach((row, ri) => {
    // Rotated text runs along the row and does not break with it (pages2 P20c:
    // a 60pt at-least row with rotated text 57pt from the foot moves whole).
    if (
      row.properties?.cantSplit === true ||
      row.properties?.height?.rule === "exact" ||
      row.cells.some(cell => isRotatedTextDirection(cell.properties?.textDirection))
    ) {
      unbreakable.add(ri);
    }
    const rules = rowRulesPt(table, ri);
    const margins = rowCellMarginsPt(table, row);
    row.cells.forEach((cell, ci) => {
      const blocks: CellBlockSplitInfo[] =
        cell.properties?.verticalMerge === "continue"
          ? []
          : [...continuing(ri, ci), ...cellBlockInfo(cell)];
      cells.set(`${ri}:${ci}`, {
        top: margins.top + rules.above,
        bottom: margins.bottom + rules.below,
        cutBottom: margins.bottom + rules.cut,
        carriedTop: margins.top + rules.carried,
        closing: rules.closing,
        blocks,
        ...(cell.properties?.verticalAlign ? { verticalAlign: cell.properties.verticalAlign } : {})
      });
    });
  });
  const keptWithNext = new Set<number>();
  table.rows.forEach((row, ri) => {
    if (ri < table.rows.length - 1 && rowKeepsWithNext(row)) {
      keptWithNext.add(ri);
    }
  });
  return { unbreakable, cells, keptWithNext };
}

/**
 * Whether a row is kept on the page of the row below it: its first cell's
 * first paragraph has `w:keepNext`. Word reads nothing else of the row (pages2
 * P29: the first cell alone keeping moves the row, the last cell alone or a
 * cell's second paragraph does not; R7b: a heading and a table whose first
 * three rows keep with the next move to the next page together).
 */
function rowKeepsWithNext(row: Table["rows"][number]): boolean {
  const block = row.cells[0]?.content[0];
  if (block?.type !== "paragraph") {
    return false;
  }
  return activeStyles
    ? activeStyles.paragraph(block).paragraphProperties.keepNext === true
    : block.properties?.keepNext === true;
}

/** How each block of a cell may break, in order. */
function cellBlockInfo(cell: Table["rows"][number]["cells"][number]): CellBlockSplitInfo[] {
  const blocks: CellBlockSplitInfo[] = [];
  for (const block of cell.content) {
    if (block.type === "paragraph") {
      blocks.push({ kind: "paragraph", rules: paragraphBreakRules(block) });
    } else if (block.type === "table") {
      blocks.push({
        kind: "table",
        headerRows: countRepeatedHeaderRows(block),
        info: tableSplitInfo(block)
      });
    }
  }
  return blocks;
}

/** Count the leading rows a table repeats on every page it spans. */
function countRepeatedHeaderRows(table: Table): number {
  let count = 0;
  for (const row of table.rows) {
    if (row.properties?.tableHeader !== true) {
      break;
    }
    count++;
  }
  return count;
}

/**
 * Whether the boundary into section `next` starts a new page.
 *
 * The break type lives on the sectPr of the section being *opened* — a
 * section's `w:type` says how that section starts (ECMA-376 §17.6.22), which
 * Word's output confirms; `layout.ts` paginates by the same rule. `continuous`
 * keeps the following content on the same page unless the paper size changes,
 * which forces one regardless. Odd/even parity is decided by `openSection`.
 */
function sectionStartsPage(
  closing: SectionProperties | undefined,
  next: SectionProperties | undefined
): boolean {
  const breakType = next?.breakType ?? "nextPage";
  if (breakType === "continuous") {
    const sameWidth =
      (closing?.pageSize?.width ?? DEFAULT_PAGE_WIDTH_TWIPS) ===
      (next?.pageSize?.width ?? DEFAULT_PAGE_WIDTH_TWIPS);
    const sameHeight =
      (closing?.pageSize?.height ?? DEFAULT_PAGE_HEIGHT_TWIPS) ===
      (next?.pageSize?.height ?? DEFAULT_PAGE_HEIGHT_TWIPS);
    return !(sameWidth && sameHeight);
  }
  return true;
}

/**
 * A section's columns, left to right, content-relative: `w:cols` equal ones
 * share the width left after the gaps between them (`w:space`, half an inch by
 * default); listed ones take their own widths and the space after each.
 */
function sectionColumns(
  props: SectionProperties | undefined,
  contentWidth: number
): { x: number; width: number }[] {
  const cols = props?.columns;
  const listed = cols?.equalWidth === false ? cols.columns : undefined;
  if (listed && listed.length > 1) {
    const out: { x: number; width: number }[] = [];
    let x = 0;
    for (const col of listed) {
      out.push({ x, width: twipsToPt(col.width) });
      x += twipsToPt(col.width) + twipsToPt(col.space ?? 0);
    }
    // A grid wider or narrower than the page is scaled to it.
    const scale = contentWidth / (x - twipsToPt(listed.at(-1)!.space ?? 0));
    return out.map(c => ({ x: c.x * scale, width: c.width * scale }));
  }
  const count = Math.max(1, cols?.count ?? 1);
  if (count === 1) {
    return [{ x: 0, width: contentWidth }];
  }
  const space = twipsToPt(cols?.space ?? 720);
  // Word sizes equal columns in whole twips, rounding down (pages2 R6j: three
  // columns of 4478.67tw are 4478tw wide — a line 0.02pt narrower than the
  // exact third breaks).
  const width = twipsToPt(Math.floor(ptToTwips(contentWidth - space * (count - 1)) / count + 1e-6));
  return Array.from({ length: count }, (_, k) => ({ x: k * (width + space), width }));
}

/** What makes two sections' columns the same, for deciding whether one region holds both. */
function columnsKey(props: SectionProperties | undefined): string {
  const cols = props?.columns;
  const listed = cols?.equalWidth === false ? cols.columns : undefined;
  if (listed && listed.length > 1) {
    return listed.map(c => `${c.width}/${c.space ?? 0}`).join(",");
  }
  const count = Math.max(1, cols?.count ?? 1);
  return count === 1 ? "1" : `${count}x${cols?.space ?? 720}`;
}

/** The properties of section `index`, or the document's when out of range. */
function sectionPropsOf(doc: DocxDocument, index: number): SectionProperties | undefined {
  const sections = collectSectionProperties(doc);
  return sections[index] ?? doc.sectionProperties;
}

/**
 * `geometry` with the body kept clear of the header and footer, as Word does: a
 * header reaching below the top margin starts the body at its bottom, and a
 * footer reaching above the bottom margin ends it at its top (the footer grows
 * up from `pgMar.footer`). Word for Mac 16 starts the body of a page with a
 * four-line header at 36pt at 126pt, not at its 72pt margin.
 */
function roomForHeaderFooter(
  geometry: PageGeometry,
  header: readonly (LayoutParagraph | LayoutTable)[],
  footer: readonly (LayoutParagraph | LayoutTable)[]
): PageGeometry {
  let marginTop = geometry.marginTop;
  for (const block of header) {
    marginTop = Math.max(marginTop, block.rect.y + block.rect.height);
  }
  let marginBottom = geometry.marginBottom;
  if (footer.length > 0) {
    const first = footer[0]!;
    const last = footer[footer.length - 1]!;
    const footerHeight = last.rect.y + last.rect.height - first.rect.y;
    marginBottom = Math.max(marginBottom, geometry.footerOffset + footerHeight);
  }
  if (marginTop === geometry.marginTop && marginBottom === geometry.marginBottom) {
    return geometry;
  }
  return {
    ...geometry,
    marginTop,
    marginBottom,
    contentHeight: geometry.height - marginTop - marginBottom
  };
}

/**
 * `content` with each run of entries of one table of contents — placed as
 * paragraphs — gathered into the block that stands for it on this page.
 */
function groupTableOfContents(doc: DocxDocument, content: PageContent[]): PageContent[] {
  const out: PageContent[] = [];
  for (const block of content) {
    if (block.type !== "paragraph" || doc.body[block.sourceIndex]?.type !== "tableOfContents") {
      out.push(block);
      continue;
    }
    // Entries sit relative to the block, as renderers read them.
    const last = out.at(-1);
    if (last?.type === "tableOfContents" && last.sourceIndex === block.sourceIndex) {
      const entry = { ...block, rect: { ...block.rect, y: block.rect.y - last.rect.y } };
      out[out.length - 1] = {
        ...last,
        rect: { ...last.rect, height: entry.rect.y + entry.rect.height },
        entries: [...last.entries, entry]
      };
    } else {
      out.push({
        type: "tableOfContents",
        rect: { ...block.rect },
        entries: [{ ...block, rect: { ...block.rect, y: 0 } }],
        sourceIndex: block.sourceIndex
      });
    }
  }
  return out;
}

/**
 * The rest of a paragraph broken at a column or page end — or all of one moved
 * there whole — set again at `width` from `y`: the source paragraph's text
 * after what the lines before the break held, without the first line's indent or the space before, which belong to
 * the paragraph's start. `undefined` when the remainder holds more than text
 * and cannot be matched back to its source.
 */
function rewrapRemainder(
  doc: DocxDocument,
  tail: LayoutParagraph,
  y: number,
  width: number,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>
): LayoutParagraph | undefined {
  const source = doc.body[tail.sourceIndex];
  if (!source || source.type !== "paragraph") {
    return undefined;
  }
  const inkOf = (text: string): number => text.replace(/\s+/g, "").length;
  let tailInk = 0;
  for (const line of tail.lines) {
    for (const run of line.runs) {
      if (run.type === "image") {
        return undefined;
      }
      if (!leaderRuns.has(run)) {
        tailInk += inkOf(run.text);
      }
    }
  }
  let sourceInk = 0;
  for (const child of source.children) {
    if ("content" in child) {
      for (const item of child.content) {
        if (item.type === "text") {
          sourceInk += inkOf(item.text);
        } else if (item.type !== "break" && item.type !== "tab") {
          return undefined;
        }
      }
    } else if (child.type !== "bookmarkStart" && child.type !== "bookmarkEnd") {
      return undefined;
    }
  }
  // Characters of ink to drop: those the lines before the break held.
  let skip = sourceInk - tailInk;
  if (skip < 0) {
    return undefined;
  }
  if (skip === 0) {
    // Moved whole: set as it is.
    return {
      ...layoutParagraph(source, y, width, options, undefined, imageMap),
      sourceIndex: tail.sourceIndex
    };
  }
  const children: Paragraph["children"][number][] = [];
  for (const child of source.children) {
    if (!("content" in child)) {
      if (skip <= 0) {
        children.push(child);
      }
      continue;
    }
    const content: (typeof child.content)[number][] = [];
    for (const item of child.content) {
      if (skip <= 0 || item.type !== "text") {
        if (skip <= 0) {
          content.push(item);
        }
        continue;
      }
      let cut = 0;
      while (cut < item.text.length && skip > 0) {
        if (!/\s/.test(item.text[cut]!)) {
          skip--;
        }
        cut++;
      }
      const rest = item.text.slice(cut).replace(/^\s+/, "");
      if (rest.length > 0) {
        content.push({ ...item, text: rest });
      }
    }
    if (content.length > 0) {
      children.push({ ...child, content });
    }
  }
  const props = activeStyles?.paragraph(source).paragraphProperties;
  const remainder: Paragraph = {
    ...source,
    children,
    properties: {
      ...source.properties,
      indent: { ...props?.indent, firstLine: 0, hanging: 0 },
      spacing: { ...props?.spacing, before: 0, beforeAutoSpacing: false }
    }
  };
  const laid = layoutParagraph(remainder, y, width, options, undefined, imageMap);
  return { ...laid, sourceIndex: tail.sourceIndex };
}

/**
 * `laid` without the space above its first line (its space before). A top rule
 * and its own space stay: they belong to the paragraph, not to the gap above it.
 */
function dropSpaceBefore(laid: LayoutParagraph): LayoutParagraph {
  const rule = laid.decorationInsets ? borderLeadPt(laid.borders?.top, false) : 0;
  const lead = (laid.lines[0]?.y ?? 0) - rule;
  if (lead <= 0) {
    return laid;
  }
  return {
    ...laid,
    rect: { ...laid.rect, height: laid.rect.height - lead },
    lines: laid.lines.map(line => ({ ...line, y: line.y - lead })),
    ...(laid.decorationInsets ? { decorationInsets: { ...laid.decorationInsets, top: rule } } : {})
  };
}

/**
 * A shape's anchor as a floating image's, for placing it: the same position
 * and wrap, with the wrap distances the writer gives a shape (0.125in either
 * side, none above or below).
 */
function shapeAsFloat(shape: DrawingShape): FloatingImage {
  return {
    type: "floatingImage",
    rId: "",
    width: shape.width,
    height: shape.height,
    horizontalPosition: shape.horizontalPosition ?? { relativeTo: "column", offset: 0 },
    verticalPosition: shape.verticalPosition ?? { relativeTo: "paragraph", offset: 0 },
    wrap: {
      style: shape.wrap?.style ?? "none",
      ...(shape.wrap?.side ? { side: shape.wrap.side } : {}),
      margins: { top: 0, bottom: 0, left: DEFAULT_SHAPE_WRAP_EMU, right: DEFAULT_SHAPE_WRAP_EMU }
    },
    ...(shape.behindDoc ? { behindDoc: true } : {})
  };
}

/** The wrap distance either side of a shape, as `drawing-frame` writes it: 0.125in. */
const DEFAULT_SHAPE_WRAP_EMU = 114300;

/**
 * The size, in EMU, of a VML shape (`w:pict`) set in the line — one whose style
 * names a width and a height and no absolute position — or `undefined`.
 */
function inlineVmlSize(rawXml: string): { width: number; height: number } | undefined {
  if (!rawXml.startsWith("<w:pict")) {
    return undefined;
  }
  const style = /<v:(?:shape|rect|roundrect|oval)\b[^>]*\bstyle="([^"]*)"/.exec(rawXml)?.[1];
  if (!style || /position\s*:\s*absolute/i.test(style)) {
    return undefined;
  }
  const length = (name: string): number | undefined => {
    const m = new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([\\d.]+)(pt|in|cm|mm|px)?`, "i").exec(style);
    if (!m) {
      return undefined;
    }
    const value = Number(m[1]);
    const unit = (m[2] ?? "px").toLowerCase();
    const pt =
      unit === "pt"
        ? value
        : unit === "in"
          ? value * 72
          : unit === "cm"
            ? (value * 72) / 2.54
            : unit === "mm"
              ? (value * 72) / 25.4
              : value * 0.75;
    return Math.round(pt * 12700);
  };
  const width = length("width");
  const height = length("height");
  return width && height ? { width, height } : undefined;
}

/** The empty paragraph a body-level floating image is anchored in. */
const ANCHOR_CARRIER: Paragraph = { type: "paragraph", children: [] };

/** One page — or one column of one — of `slot`'s section, its lines on the section's document grid. */
function buildPage(...args: Parameters<typeof buildPageOnGrid>): BuildPageResult {
  return withGridPitch(sectionGridPitch((args[11] ?? args[8]).props), () =>
    buildPageOnGrid(...args)
  );
}

/**
 * One page, filled region by region: each section on it gets its columns, one
 * after another, and a continuous section break in the middle of the page
 * starts the next section's columns below the last. The columns of a section
 * ended by such a break are balanced — set as short as they can be while still
 * holding its text — as Word does (pages2 P12b). The page's footnotes are laid
 * out once every region is done.
 */
function composePage(
  doc: DocxDocument,
  pageNumber: number,
  layout: LayoutResult,
  options: FullLayoutOptions | undefined,
  pendingNotes: readonly LaidNote[],
  listMarkers: ReadonlyMap<Paragraph, ListMarker> | undefined,
  carriedIn: readonly CarriedBlock[],
  flow: FlowCursor,
  slot: PageSlot,
  imageMap: ReadonlyMap<string, ImageDef>,
  noteCache: Map<string, LaidNote>,
  sections: readonly (SectionProperties | undefined)[]
): BuildPageResult {
  const contentWidth = computePageGeometry(
    slot.props,
    options?.pageGeometry,
    slot.gutterSide
  ).contentWidth;
  const content: PageContent[] = [];
  let page: LayoutPage | undefined;
  let notes: PageNoteSet | undefined;
  let carried = carriedIn;
  let top = 0;
  let bottom = 0;
  let sectionIndex = slot.sectionIndex;
  let stop: BuildPageResult["stop"];

  /** Fill the columns of `sectionIndex` from `top`, down to `limit` when given. */
  const fillColumns = (
    columns: readonly { x: number; width: number }[],
    limit: number | undefined,
    noteFloor?: number,
    refused?: ReadonlySet<number>
  ): {
    blocks: PageContent[];
    bottom: number;
    stop: BuildPageResult["stop"];
    /** The columns' heights added up. */
    used: number;
  } => {
    const blocks: PageContent[] = [];
    let reached = top;
    let used = 0;
    for (let column = 0; column < columns.length; column++) {
      const { x, width } = columns[column]!;
      const result = buildPage(
        doc,
        pageNumber,
        layout,
        options,
        pendingNotes,
        listMarkers,
        carried,
        flow,
        slot,
        imageMap,
        noteCache,
        {
          top,
          bottom: limit,
          x,
          width,
          column,
          sectionIndex,
          props: sections[sectionIndex],
          ...(noteFloor ? { noteFloor } : {}),
          ...(refused ? { refusedNotes: refused } : {})
        },
        notes
      );
      page ??= result.page;
      notes = result.notes;
      carried = result.carriedContent;
      reached = Math.max(reached, result.bottom);
      used += Math.max(0, result.bottom - top);
      for (const block of result.page.content) {
        blocks.push(
          x === 0 || block.type === "float"
            ? block
            : { ...block, rect: { ...block.rect, x: block.rect.x + x } }
        );
      }
      const done = carried.length === 0 && flow.nextItem >= doc.body.length;
      if (done || result.stop === "region" || result.stop === "page") {
        return { blocks, bottom: reached, stop: done ? undefined : result.stop, used };
      }
    }
    return { blocks, bottom: reached, stop: undefined, used };
  };

  for (;;) {
    const columns = sectionColumns(sections[sectionIndex], contentWidth);
    const start = {
      nextItem: flow.nextItem,
      endPageAfterItem: flow.endPageAfterItem,
      carried,
      notes: notes ? [...notes.notes] : undefined
    };
    const restore = (): void => {
      flow.nextItem = start.nextItem;
      flow.endPageAfterItem = start.endPageAfterItem;
      carried = start.carried;
      if (notes && start.notes) {
        notes.notes.splice(0, notes.notes.length, ...start.notes);
      }
    };
    /**
     * Columns ended by a continuous break, balanced as Word balances them
     * (pages2 P35, P36). Word first tries the even share of the columns'
     * height; when the text does not fit in that, the height goes up in whole
     * lines — multiples of the region's most common line height, counted from
     * its top — never to a height between two such steps: 14 lines of 14.49pt
     * in three columns fit 80.45pt high as 5/5/4, and Word sets them 86.94pt
     * high as 5/6/3. A space after that does not fit hangs below the columns
     * and still pushes the next section down (see `placeParagraph`).
     */
    const balanceColumns = (unbalanced: ReturnType<typeof fillColumns>) => {
      const fits = (bottom: number): boolean => {
        restore();
        const trial = fillColumns(columns, bottom);
        return trial.stop === "region" && carried.length === 0;
      };
      const even = top + unbalanced.used / columns.length;
      const step = mostCommonLineHeight(unbalanced.blocks);
      let bottom = unbalanced.bottom;
      if (fits(even)) {
        bottom = even;
      } else if (step > 0) {
        let lo = Math.ceil((even - top) / step - 1e-6) - 1;
        let hi = Math.ceil((unbalanced.bottom - top) / step);
        while (hi - lo > 1) {
          const mid = Math.floor((lo + hi) / 2);
          if (fits(top + mid * step + FIT_EPSILON_PT)) {
            hi = mid;
          } else {
            lo = mid;
          }
        }
        bottom = Math.min(unbalanced.bottom, top + hi * step + FIT_EPSILON_PT);
      }
      restore();
      return fillColumns(columns, bottom);
    };
    let filled = fillColumns(columns, undefined);
    // Notes referenced in a later column claim room under the earlier ones
    // too: fill again until every column ends above the page's whole area.
    // When that room pushes a reference off the page, Word keeps the first
    // fill and moves the reference on instead (pages2 R11a).
    let refused: Set<number> | undefined;
    for (let round = 0, floor = 0; columns.length > 1 && notes && round < 4; round++) {
      const needed = notes.reserve();
      if (Math.abs(needed - floor) <= 0.01) {
        break;
      }
      floor = needed;
      restore();
      filled = fillColumns(columns, undefined, floor, refused);
      // The refill starts from the notes the previous fill collected, so one
      // whose reference no longer reaches the page still holds its room.
      const referenced = new Set<number>();
      for (const block of filled.blocks) {
        collectLaidNoteIds(block, referenced);
      }
      const earlier = new Set([
        ...pendingNotes.map(note => note.id),
        ...(start.notes ?? []).map(note => note.id)
      ]);
      const lost = notes.notes
        .map(note => note.id)
        .filter(id => !referenced.has(id) && !earlier.has(id));
      if (lost.length > 0) {
        refused = new Set([...(refused ?? []), ...lost]);
        floor = 0;
        restore();
        notes.remove(lost);
        filled = fillColumns(columns, undefined, undefined, refused);
      }
    }
    if (filled.stop === "region" && columns.length > 1) {
      filled = balanceColumns(filled);
    }
    content.push(...filled.blocks);
    bottom = Math.max(bottom, filled.bottom);
    stop = filled.stop;
    if (stop !== "region" || carried.length > 0) {
      break;
    }
    // The next section continues on this page, below this one's columns.
    top = filled.bottom;
    sectionIndex = layout.contentSections[flow.nextItem] ?? sections.length - 1;
  }

  const laidPage = page!;
  const footnotes = layoutFootnoteArea(notes!, laidPage.geometry, bottom);
  const alignment = sections[sectionIndex]?.verticalAlign;
  const aligned =
    sectionIndex === slot.sectionIndex && footnotes.laid.length === 0
      ? alignPageVertically(content, alignment, laidPage.geometry.contentHeight - bottom)
      : content;
  return {
    page: {
      ...laidPage,
      content: groupTableOfContents(doc, aligned),
      ...(footnotes.laid.length > 0 ? { footnoteArea: footnotes.laid } : {}),
      ...(footnotes.laid.length > 0 && footnotes.separator
        ? { footnoteSeparator: footnotes.separator }
        : {})
    },
    continuedNotes: footnotes.continued,
    carriedContent: carried,
    bottom,
    ...(stop ? { stop } : {}),
    notes: notes!,
    ...(sectionIndex !== slot.sectionIndex ? { endSection: sectionIndex } : {})
  };
}

/** The line height most of `blocks`' lines have (0 when they hold no lines). */
function mostCommonLineHeight(blocks: readonly PageContent[]): number {
  const counts = new Map<number, number>();
  for (const block of blocks) {
    if (block.type === "paragraph") {
      for (const line of block.lines) {
        const key = Math.round(line.height * 100) / 100;
        counts.set(key, (counts.get(key) ?? 0) + 1);
      }
    }
  }
  let height = 0;
  let most = 0;
  for (const [key, count] of counts) {
    if (count > most) {
      height = key;
      most = count;
    }
  }
  return height;
}

/**
 * A page's body moved down by the room left below it, as the section's
 * `w:vAlign` asks: half of it for `center`, all of it for `bottom`, and for
 * `both` shared equally between the gaps of consecutive blocks (measured in
 * Word: the k-th of n blocks moves k/(n-1) of the room).
 */
function alignPageVertically(
  content: PageContent[],
  alignment: SectionProperties["verticalAlign"],
  slack: number
): PageContent[] {
  if (!alignment || alignment === "top" || slack <= 0 || content.length === 0) {
    return content;
  }
  const flowing = content.filter(block => block.type !== "float").length;
  let index = -1;
  return content.map(block => {
    if (block.type !== "float") {
      index++;
    }
    const shift =
      alignment === "center"
        ? slack / 2
        : alignment === "bottom"
          ? slack
          : flowing > 1
            ? (slack * Math.max(index, 0)) / (flowing - 1)
            : 0;
    return shift === 0 ? block : { ...block, rect: { ...block.rect, y: block.rect.y + shift } };
  });
}

function buildPageOnGrid(
  doc: DocxDocument,
  pageNumber: number,
  layout: LayoutResult,
  options: FullLayoutOptions | undefined,
  pendingNotes: readonly LaidNote[],
  listMarkers: ReadonlyMap<Paragraph, ListMarker> | undefined,
  carriedIn: readonly CarriedBlock[],
  flow: FlowCursor,
  slot: PageSlot,
  imageMap: ReadonlyMap<string, ImageDef>,
  noteCache: Map<string, LaidNote>,
  region?: PageRegion,
  sharedNotes?: PageNoteSet
): BuildPageResult {
  const sectionProps = slot.props;
  const pageGeometry = computePageGeometry(sectionProps, options?.pageGeometry, slot.gutterSide);
  const header = layoutHeader(doc, slot, pageGeometry, options, imageMap);
  const footer = layoutFooter(doc, slot, pageGeometry, options, imageMap);
  const geometry = roomForHeaderFooter(pageGeometry, header, footer);
  const content: PageContent[] = [];
  /**
   * The page's footnotes, in reading order, after those continued from the
   * previous page. A note joins when the block holding its reference is placed,
   * so a block carried to the next page takes its notes with it.
   */
  const noteColumns = {
    columns: sectionColumns(region?.props, geometry.contentWidth),
    contentHeight: geometry.contentHeight
  };
  const pageNotes =
    sharedNotes ??
    createPageNotes(
      doc,
      pendingNotes,
      geometry.contentWidth,
      options,
      imageMap,
      noteCache,
      noteColumns
    );
  pageNotes.useColumns(noteColumns);
  pageNotes.startColumn();
  /** Where the body text of this page — or of this column of it — starts and how wide it is. */
  const regionTop = region?.top ?? 0;
  const bodyWidth = region?.width ?? geometry.contentWidth;
  /** The lowest the body may reach (content-relative) with `ids` also on the page. */
  const bodyLimit = (ids: readonly number[]): number => {
    const fresh = ids.length > 0 ? pageNotes.newNotes(ids) : [];
    const area = pageNotes.reserve(fresh.length > 0 ? fresh : undefined);
    if (region?.refusedNotes && fresh.some(note => region.refusedNotes!.has(note.id))) {
      return -Infinity;
    }
    return Math.min(
      region?.bottom ?? Infinity,
      geometry.contentHeight - Math.max(region?.noteFloor ?? 0, area)
    );
  };
  /**
   * The lowest each line of `laid` may reach: a line referencing a note only
   * fits if the note does too.
   */
  const lineLimits = (laid: LayoutParagraph): ((line: number) => number) => {
    const ids = laid.lineNoteIds;
    if (!ids || ids.every(list => list.length === 0)) {
      const limit = bodyLimit([]) - laid.rect.y;
      return () => limit;
    }
    const limits: number[] = [];
    const seen: number[] = [];
    for (let i = 0; i < laid.lines.length; i++) {
      seen.push(...(ids[i] ?? []));
      limits.push(bodyLimit(seen) - laid.rect.y);
    }
    return line => limits[Math.min(line, limits.length - 1)]!;
  };

  /**
   * Wrap exclusion zones from floats with `square` / `tight` /
   * `through` wrap, populated as we iterate so subsequent paragraphs
   * (later in document order) avoid them line-by-line. Floats that
   * appear AFTER a paragraph in the source do not push back into
   * preceding lines — this matches Word's behaviour where re-flow on
   * insertion happens at edit time, not render time.
   */
  const pageExclusions: WrapExclusion[] = [];

  let cursorY = regionTop; // relative to content area top

  /**
   * Blocks this page could not hold. Once the first one overflows, everything
   * after it must follow — pushing a later block up past an earlier one would
   * reorder the document.
   */
  const carriedOut: CarriedBlock[] = [];
  /** Why the flow stopped before the page was full, when it is not the page's end. */
  let stop: BuildPageResult["stop"];

  /**
   * A body paragraph's effective `w:widowControl` and `w:keepLines`. Word
   * controls widows unless told not to: a document that never mentions it
   * moves a two-line paragraph whole rather than leave one line behind.
   */
  const breakRules = (index: number): ParagraphBreakRules => {
    const item = doc.body[index];
    if (!item || item.type !== "paragraph" || !activeStyles) {
      return { widowControl: true, keepLines: false };
    }
    const props = activeStyles.paragraph(item).paragraphProperties;
    return { widowControl: props.widowControl !== false, keepLines: props.keepLines === true };
  };

  /**
   * Whether `laid` opens a page with nothing on it, so at least part of it has
   * to be placed here whether or not it fits. A page that opens with notes
   * continued from the previous one is not empty: they are what it holds, and
   * Word gives body text only the room they leave (final6 Qb3: a page of
   * nothing but a continued note). A section's columns that start below other
   * text on the page do not open it: what does not fit there goes to the next
   * page (pages2 R5e: a two-column section starting 7pt above the foot).
   */
  const startsEmptyPage = (laid: LayoutBlock): boolean =>
    content.length === 0 &&
    regionTop <= 0.01 &&
    laid.rect.y <= regionTop + 0.01 &&
    pendingNotes.length === 0;

  /**
   * Place a paragraph, splitting it at the bottom margin when it does not fit.
   * Returns false once the page is closed, so the caller stops placing.
   */
  /**
   * Paragraphs set whole but for the space after that would not fit, with the
   * paragraph as laid out: one moved on to the next page gets its space back
   * (pages2 R4f: a keep-with-next paragraph moved from a page foot).
   */
  const trimmedAfter = new Map<LayoutParagraph, LayoutParagraph>();

  const placeParagraph = (placed: LayoutParagraph, continued = false): boolean => {
    if (carriedOut.length > 0) {
      carriedOut.push({ kind: "paragraph", block: placed, continued });
      return false;
    }
    // Word sets no space before a paragraph that opens a page — after a page
    // break of any kind, or a paragraph moved over for want of room — except
    // on the first page of a section, the document's first page included.
    // Measured in Word for Mac 16: 24pt before lands at the 72pt top margin
    // after a natural break, a page-break run or `w:pageBreakBefore`, and at
    // 96pt on a page a next-page section break opened.
    const laid =
      content.length === 0 &&
      placed.rect.y <= 0.01 &&
      (slot.pageInSection > 1 || (region?.column ?? 0) > 0)
        ? dropSpaceBefore(placed)
        : placed;
    const availableAt = lineLimits(laid);
    const available = availableAt(laid.lines.length - 1);
    // A column break ends the column after its line; the rest of the paragraph
    // (its mark at least) opens the next one.
    const columnBreak = laid.columnBreakAfterLines?.[0];
    if (
      columnBreak !== undefined &&
      columnBreak + 1 < laid.lines.length &&
      (laid.pageBreakAfterLines?.[0] ?? Infinity) >= columnBreak &&
      laid.lines[columnBreak]!.y + laid.lines[columnBreak]!.height <=
        availableAt(columnBreak) + FIT_EPSILON_PT
    ) {
      const cut = sliceParagraphAtLine(laid, columnBreak + 1);
      content.push(cut.head);
      pageNotes.add(cut.head.lineNoteIds?.flat() ?? []);
      cursorY = cut.head.rect.y + cut.head.rect.height;
      carriedOut.push({ kind: "paragraph", block: cut.tail, continued: true });
      stop = "column";
      return false;
    }
    // An internal page break must be honoured even when the whole paragraph
    // would otherwise fit on this page.
    const forcedBreakLine = laid.pageBreakAfterLines?.[0];
    const hasForcedBreak = forcedBreakLine !== undefined && forcedBreakLine + 1 < laid.lines.length;
    if (laid.rect.height <= available + FIT_EPSILON_PT && !hasForcedBreak) {
      content.push(laid);
      pageNotes.add(laid.lineNoteIds?.flat() ?? []);
      cursorY = laid.rect.y + laid.rect.height;
      // A page break on the last line ends the page here, its paragraph mark
      // included: Word starts the next page with the next paragraph, not with
      // an empty line (a heading after a paragraph holding only a page break
      // sits at the top margin in Word for Mac 16).
      if (forcedBreakLine !== undefined) {
        stop = "page";
        return false;
      }
      return true;
    }
    if (
      hasForcedBreak &&
      laid.lines[forcedBreakLine].y + laid.lines[forcedBreakLine].height <=
        availableAt(forcedBreakLine) + FIT_EPSILON_PT
    ) {
      const cut = sliceParagraphAtLine(laid, forcedBreakLine + 1);
      content.push(cut.head);
      pageNotes.add(cut.head.lineNoteIds?.flat() ?? []);
      cursorY = cut.head.rect.y + cut.head.rect.height;
      // A page break that ends its paragraph takes the paragraph mark with it:
      // Word starts the next page with the next paragraph, not with an empty
      // line (Word for Mac 16 puts a heading after such a break at the top
      // margin).
      if (cut.tail.lines.some(line => line.runs.length > 0)) {
        carriedOut.push({ kind: "paragraph", block: cut.tail, continued: true });
      }
      stop = "page";
      return false;
    }
    const atPageTop = startsEmptyPage(laid);
    const { head, tail } = splitLayoutParagraph(
      laid,
      availableAt,
      geometry.contentHeight,
      atPageTop,
      breakRules(laid.sourceIndex),
      // Above a footnote area the space after has to fit too — not at the
      // foot of columns being balanced, which ends above the page's foot
      // (pages2 R11j p1).
      region?.bottom === undefined &&
        (pageNotes.notes.length > 0 || (laid.lineNoteIds?.some(ids => ids.length > 0) ?? false)),
      continued
    );
    if (head) {
      // In columns being balanced a space after that does not fit still
      // pushes down what follows the columns (pages2 P15b, P35g).
      const kept = !tail && head !== laid && region?.bottom !== undefined ? laid : head;
      content.push(kept);
      pageNotes.add(head.lineNoteIds?.flat() ?? []);
      cursorY = kept.rect.y + kept.rect.height;
      if (!tail && kept !== laid) {
        trimmedAfter.set(head, laid);
      }
    }
    if (tail) {
      carriedOut.push({ kind: "paragraph", block: tail, continued: continued || head !== null });
      return false;
    }
    return true;
  };

  /** Place a table, splitting it on a row edge when it does not fit. */
  const placeTable = (
    laid: LayoutTable,
    headerRows: number,
    noteIds: readonly number[],
    splitInfo: TableSplitInfo
  ): boolean => {
    if (carriedOut.length > 0) {
      carriedOut.push({ kind: "table", block: laid, headerRows, noteIds, splitInfo });
      return false;
    }
    // A table's notes are claimed by the first page any of it lands on.
    const available = bodyLimit(noteIds) - laid.rect.y;
    if (laid.rect.height <= available + FIT_EPSILON_PT) {
      content.push(laid);
      pageNotes.add(noteIds);
      cursorY = laid.rect.y + laid.rect.height;
      return true;
    }
    const atPageTop = startsEmptyPage(laid);
    const { head, tail } = splitLayoutTable(
      laid,
      available,
      headerRows,
      atPageTop,
      splitInfo,
      geometry.contentHeight
    );
    if (head) {
      content.push(head);
      pageNotes.add(noteIds);
      cursorY = head.rect.y + head.rect.height;
    }
    if (tail) {
      carriedOut.push({
        kind: "table",
        block: tail,
        headerRows,
        noteIds: head ? [] : noteIds,
        splitInfo
      });
      return false;
    }
    return true;
  };

  // Remainders carried from the previous page go first, at the content top.
  //
  // Remember when a continued table finishes here. Preview's PDF selection
  // analysis extends a table that starts at the top of a continuation page over
  // a callout immediately below it. The callout itself may stay on this page,
  // but it has to be the page's last block: a thematic break below it is another
  // full-width rule, and Preview takes that as the bottom of a phantom final row
  // containing the whole callout.
  let completedContinuedTable = false;
  let completedEndPageParagraph = false;
  // Endnotes running on from the page before open the page under the
  // continuation separator, whether or not a note was split there (pages2
  // R11b: a page starting with a whole note has the separator above it).
  const notesStart = activeEndnotesStart;
  if (
    notesStart !== undefined &&
    notesStart > 0 &&
    (region?.column ?? 0) === 0 &&
    regionTop <= 0.01 &&
    (carriedIn.length > 0
      ? carriedIn[0]!.block.sourceIndex >= notesStart
      : flow.nextItem >= notesStart && flow.nextItem < doc.body.length)
  ) {
    const def = doc.endnotes?.find(note => note.type === "continuationSeparator");
    for (const para of def?.content.length ? def.content : [DEFAULT_SEPARATOR_PARAGRAPH]) {
      const laid = { ...layoutParagraph(para, cursorY, bodyWidth, options), sourceIndex: -1 };
      content.push(laid);
      cursorY = laid.rect.y + laid.rect.height;
    }
  }
  for (const carried of carriedIn) {
    if (carried.kind === "paragraph") {
      // Continuing in a column of another width, the rest of the paragraph is
      // set again at that width (pages2 P12d: two lines of a narrow column
      // are one in the wide column after it).
      const rewrapped =
        Math.abs(carried.block.rect.width - bodyWidth) > 0.01
          ? rewrapRemainder(doc, carried.block, cursorY, bodyWidth, options, imageMap)
          : undefined;
      const rebased: LayoutParagraph = rewrapped ?? {
        ...carried.block,
        rect: { ...carried.block.rect, y: cursorY }
      };
      if (!placeParagraph(rebased, carried.continued)) {
        break;
      }
      completedEndPageParagraph ||= rebased.endPageWhenComplete === true;
    } else {
      // Cell rects are table-relative, so moving the table is enough.
      const rebased: LayoutTable = {
        ...carried.block,
        rect: { ...carried.block.rect, y: cursorY }
      };
      if (!placeTable(rebased, carried.headerRows, carried.noteIds, carried.splitInfo)) {
        break;
      }
      completedContinuedTable = true;
    }
  }

  // Which body items this page draws from: in order from the shared cursor,
  // stopping when one no longer fits. A page of notes only gets a cursor at the
  // end of the body.
  const startItem = flow.nextItem;
  let placedItems = 0;
  if (completedEndPageParagraph) {
    delete flow.endPageAfterItem;
  }
  const nextAfterContinuation = doc.body[flow.nextItem];
  const quoteAfterContinuedTable =
    completedContinuedTable &&
    nextAfterContinuation?.type === "paragraph" &&
    nextAfterContinuation.properties?.style === "Quote";
  let quoteRunEnd = flow.endPageAfterItem ?? -1;
  if (quoteAfterContinuedTable) {
    quoteRunEnd = flow.nextItem;
    let following = doc.body[quoteRunEnd + 1];
    while (following?.type === "paragraph" && following.properties?.style === "Quote") {
      quoteRunEnd++;
      following = doc.body[quoteRunEnd + 1];
    }
    flow.endPageAfterItem = quoteRunEnd;
  }
  let stopFlow = completedEndPageParagraph;
  if (stopFlow) {
    stop = "page";
  }

  /** A body item's effective `w:keepNext` — "do not end a page on me". */
  const keepsWithNext = (index: number): boolean => {
    const item = doc.body[index];
    if (!item || item.type !== "paragraph" || !activeStyles) {
      return false;
    }
    return activeStyles.paragraph(item).paragraphProperties.keepNext === true;
  };

  /** Floating images already placed with the paragraph they are anchored in. */
  const placedFloats = new Set<number>();

  /**
   * How far below `top` a float pushes the text that follows its anchor, or
   * `undefined` when text runs beside it or over it instead.
   */
  const floatClearance = (float: FloatingImage, top: number): number | undefined => {
    const anchored =
      float.simplePos != null || float.horizontalPosition != null || float.verticalPosition != null;
    const style = float.wrap?.style;
    const around = style === "square" || style === "tight" || style === "through";
    if (!((!anchored && !float.behindDoc && !around) || style === "topAndBottom")) {
      return undefined;
    }
    const laid = layoutFloatingImage(
      float,
      top,
      bodyWidth,
      geometry.contentHeight,
      geometry,
      -1,
      imageMap
    );
    const padBottom = float.wrap?.margins?.bottom ? emuToPt(float.wrap.margins.bottom) : 0;
    return laid.rect.y + laid.rect.height + padBottom - top;
  };

  /**
   * How low a float positioned against a paragraph at `top` reaches: Word moves
   * that paragraph to the next page rather than let the float run off this one
   * (pages2 R2e). One placed against the page or margin does not move with it,
   * nor one behind or in front of the text, which the text does not make room for.
   */
  const floatBottom = (float: FloatingImage, top: number): number => {
    const relativeTo = float.verticalPosition?.relativeTo;
    if (
      float.simplePos != null ||
      float.behindDoc === true ||
      (float.wrap?.style ?? "none") === "none" ||
      (relativeTo !== undefined && relativeTo !== "paragraph" && relativeTo !== "line")
    ) {
      return top;
    }
    const laid = layoutFloatingImage(
      float,
      top,
      bodyWidth,
      geometry.contentHeight,
      geometry,
      -1,
      imageMap
    );
    return laid.rect.y + laid.rect.height;
  };

  /**
   * Put a floating image on the page against the cursor: text after a
   * `topAndBottom` float (or an unanchored one) moves below it, text beside a
   * `square`, `tight` or `through` float wraps around it, and a float behind
   * the text leaves it alone.
   */
  const placeFloat = (
    float: FloatingImage,
    index: number,
    draw?: (rect: LayoutRect) => PageContent
  ): void => {
    const laid = layoutFloatingImage(
      float,
      cursorY,
      bodyWidth,
      geometry.contentHeight,
      geometry,
      index,
      imageMap
    );
    content.push(draw ? draw(laid.rect) : laid);
    const style = float.wrap?.style;
    if (style === "topAndBottom" && float.verticalPosition != null && !float.behindDoc) {
      // Text stays above a positioned top-and-bottom float as far as it fits:
      // a line clear of it stays where it is, one it would cross goes below
      // it (pages2 R5f: an empty anchor line 14.5pt high stays above a float
      // 18pt below it; at 0, 6 or 12pt it goes below).
      pageExclusions.push({
        xLeft: -FULL_WIDTH_BAND_PT,
        xRight: FULL_WIDTH_BAND_PT,
        yTop: laid.rect.y - (laid.wrap?.margins?.top ?? 0),
        yBottom: laid.rect.y + laid.rect.height + (laid.wrap?.margins?.bottom ?? 0),
        wrapSide: "bothSides"
      });
      return;
    }
    const clearance = floatClearance(float, cursorY);
    if (clearance !== undefined) {
      cursorY += clearance;
    }
    if ((style === "square" || style === "tight" || style === "through") && !float.behindDoc) {
      // An exclusion band covering the float's rect plus its wrap padding:
      // later paragraphs on the page wrap their lines around it.
      const padL = laid.wrap?.margins?.left ?? 0;
      const padR = laid.wrap?.margins?.right ?? 0;
      const padT = laid.wrap?.margins?.top ?? 0;
      const padB = laid.wrap?.margins?.bottom ?? 0;
      pageExclusions.push({
        xLeft: laid.rect.x - padL,
        xRight: laid.rect.x + laid.rect.width + padR,
        yTop: laid.rect.y - padT,
        yBottom: laid.rect.y + laid.rect.height + padB,
        wrapSide: laid.wrap?.side ?? "bothSides"
      });
    }
  };

  /**
   * Set a framed paragraph (`w:framePr`, drop caps included) at its frame's
   * width and position. A frame text may wrap around takes no room in the flow
   * and later lines wrap around it, `hSpace` / `vSpace` away; a `notBeside`
   * frame stands in the flow at its own width. Word for Mac 16, pages2 P22:
   * - a drop cap is as wide as its letter and as high as its paragraph, at the
   *   cursor; `drop` sits at the column's left edge, `margin` just left of it;
   * - a frame anchored to the text sits at the cursor, `y` below it; one with
   *   no anchor at the margin's top-left; `xAlign` aligns it in the margin
   *   (`page`: the page), else it sits `x` from that edge.
   */
  const placeFrame = (item: Paragraph, frame: ParagraphFrame, index: number): boolean => {
    const drop = frame.dropCap === "drop" || frame.dropCap === "margin";
    const natural = layoutParagraph(item, 0, ROTATED_LINE_WIDTH_PT, options, undefined, imageMap);
    const width =
      frame.width !== undefined && frame.width > 0
        ? twipsToPt(frame.width)
        : Math.min(
            bodyWidth,
            Math.max(0, ...natural.lines.flatMap(l => l.runs.map(r => r.x + r.width)))
          );
    const onPage = frame.hAnchor === "page";
    const origin = onPage ? -geometry.marginLeft : 0;
    const room = onPage ? geometry.width : bodyWidth;
    const x =
      frame.dropCap === "margin"
        ? -width
        : drop
          ? 0
          : frame.xAlign === "right" || frame.xAlign === "outside"
            ? origin + room - width
            : frame.xAlign === "center"
              ? origin + (room - width) / 2
              : origin + (frame.xAlign ? 0 : twipsToPt(frame.x ?? 0));
    const offsetY = twipsToPt(frame.y ?? 0);
    const y =
      drop || frame.vAnchor === "text"
        ? cursorY + offsetY
        : frame.vAnchor === "page"
          ? offsetY - geometry.marginTop
          : offsetY;
    const laid = layoutParagraph(item, y, width, options, undefined, imageMap, listMarkers);
    const placed: LayoutParagraph = {
      ...laid,
      rect: { ...laid.rect, x: laid.rect.x + x },
      sourceIndex: index
    };
    if (frame.wrap === "notBeside" || frame.wrap === "none") {
      return placeParagraph(placed);
    }
    content.push(placed);
    pageNotes.add(placed.lineNoteIds?.flat() ?? []);
    const hSpace = twipsToPt(frame.hSpace ?? 0);
    const vSpace = twipsToPt(frame.vSpace ?? 0);
    pageExclusions.push({
      xLeft: x - hSpace,
      xRight: x + width + hSpace,
      yTop: y - vSpace,
      yBottom: y + laid.rect.height + vSpace,
      wrapSide: "bothSides"
    });
    return true;
  };

  for (let i = startItem; i < doc.body.length; i++) {
    if (stopFlow || carriedOut.length > 0) {
      break;
    }
    // A section boundary ends the page when the section being closed says so,
    // because the next section may use different paper, margins and headers.
    if (
      layout.contentSections[i] !== slot.sectionIndex &&
      (content.length > 0 || placedItems > 0) &&
      sectionStartsPage(slot.props, sectionPropsOf(doc, layout.contentSections[i]))
    ) {
      stop = "page";
      break;
    }
    // A continuous section set in other columns starts a region of its own
    // below this one (see `composePage`).
    const sectionHere = region?.sectionIndex ?? slot.sectionIndex;
    if (
      layout.contentSections[i] !== sectionHere &&
      columnsKey(region?.props ?? slot.props) !==
        columnsKey(sectionPropsOf(doc, layout.contentSections[i]))
    ) {
      stop = "region";
      break;
    }
    // A forced break (explicit page break or section break) must be honoured
    // even when the rest of the page is empty. It comes after the region
    // check: columns ended by a continuous break are balanced even when the
    // next section's first paragraph opens a new page (pages2 P35), and that
    // paragraph, starting a region below them, still breaks the page.
    if (
      layout.forcedBreakBefore[i] &&
      (content.length > 0 ||
        placedItems > 0 ||
        (region !== undefined && (region.top > 0 || region.column > 0)))
    ) {
      stop = "page";
      break;
    }

    const item = doc.body[i];
    /** Set when the item does not fit and starts the next page instead. */
    let deferItem = false;
    // Paragraphs and tables claim their notes as they are placed (see
    // `placeParagraph`, `placeTable`); every other block is placed whole here.
    const itemFootnoteRefs: number[] = [];
    if (item.type !== "paragraph" && item.type !== "tableOfContents") {
      collectFootnoteRefsFromBody(item, itemFootnoteRefs);
    }
    const pageContext: PageLayoutContext = {
      exclusions: pageExclusions,
      contentWidth: bodyWidth
    };
    switch (item.type) {
      case "paragraph": {
        const frame = item.properties?.frame;
        if (frame) {
          if (!placeFrame(item, frame, i)) {
            stopFlow = true;
          }
          break;
        }
        // Floating images anchored among this paragraph's text are positioned
        // against its top, and its lines wrap around them.
        for (let j = i + 1; doc.body[j]?.type === "floatingImage"; j++) {
          const float = doc.body[j] as FloatingImage;
          if (float.inPrecedingParagraph !== true) {
            break;
          }
          placeFloat(float, j);
          placedFloats.add(j);
        }
        const laid = layoutParagraph(
          item,
          cursorY,
          bodyWidth,
          options,
          pageContext,
          imageMap,
          listMarkers
        );
        const positioned = {
          ...laid,
          sourceIndex: i,
          ...(i === quoteRunEnd ? { endPageWhenComplete: true } : {})
        };
        if (!placeParagraph(positioned)) {
          stopFlow = true;
        }
        break;
      }
      case "table": {
        // A table does not wrap: one starting beside a top-and-bottom float
        // starts below it.
        for (const band of pageExclusions) {
          if (
            band.xLeft <= -FULL_WIDTH_BAND_PT &&
            cursorY >= band.yTop - FIT_EPSILON_PT &&
            cursorY < band.yBottom
          ) {
            cursorY = band.yBottom;
          }
        }
        const laid = layoutTable(item, cursorY, bodyWidth, i, options, imageMap);
        if (
          !placeTable(laid, countRepeatedHeaderRows(item), itemFootnoteRefs, tableSplitInfo(item))
        ) {
          stopFlow = true;
        }
        break;
      }
      case "floatingImage": {
        // One read from among a paragraph's text was placed with it.
        if (item.inPrecedingParagraph === true && placedFloats.has(i)) {
          break;
        }
        // A floating image of its own stands for the paragraph holding its
        // anchor, which the writer emits empty and Word lays out like any
        // other: a line of its own, below the image when text may not run
        // beside it (pages2 R1c). The image moves with that paragraph to the
        // next page.
        const carrier = layoutParagraph(
          ANCHOR_CARRIER,
          cursorY,
          bodyWidth,
          options,
          pageContext,
          imageMap,
          listMarkers
        );
        const clearance = floatClearance(item, cursorY);
        if (
          (content.length > 0 || placedItems > 0) &&
          Math.max(cursorY + (clearance ?? 0) + carrier.rect.height, floatBottom(item, cursorY)) >
            bodyLimit([]) + FIT_EPSILON_PT
        ) {
          deferItem = true;
          break;
        }
        placeFloat(item, i);
        placeParagraph({
          ...layoutParagraph(
            ANCHOR_CARRIER,
            cursorY,
            bodyWidth,
            options,
            pageContext,
            imageMap,
            listMarkers
          ),
          sourceIndex: i
        });
        break;
      }
      case "textBox": {
        const laid = layoutTextBox(item, cursorY, bodyWidth, i, options, imageMap);
        // The writer sets a text box in a paragraph of its own, in the line
        // like a picture of its size: the paragraph's line holds it on the
        // baseline, and its spacing comes on top (pages2 R2i).
        const carrier = layoutParagraph(
          {
            type: "paragraph",
            children: [
              {
                content: [
                  {
                    type: "image",
                    rId: "",
                    width: Math.round(laid.rect.width * 12700),
                    height: Math.round(laid.rect.height * 12700)
                  }
                ]
              }
            ]
          },
          cursorY,
          bodyWidth,
          options,
          pageContext,
          imageMap,
          listMarkers
        );
        if (
          (content.length > 0 || placedItems > 0) &&
          carrier.rect.y + carrier.rect.height - (carrier.lines[0]?.bottomSlack ?? 0) >
            bodyLimit([]) + FIT_EPSILON_PT
        ) {
          deferItem = true;
          break;
        }
        const line = carrier.lines[0];
        const top = line ? carrier.rect.y + line.y + line.baseline - laid.rect.height : cursorY;
        content.push({ ...laid, rect: { ...laid.rect, y: top } });
        placeParagraph({ ...carrier, sourceIndex: i });
        break;
      }
      case "drawingShape": {
        // An anchored shape, written in a paragraph of its own like a floating
        // image, is placed like one: the paragraph takes its line, and text
        // runs around, past or over the shape as its wrap says (pages2 R2i).
        const float = shapeAsFloat(item);
        const carrier = layoutParagraph(
          ANCHOR_CARRIER,
          cursorY,
          bodyWidth,
          options,
          pageContext,
          imageMap,
          listMarkers
        );
        const clearance = floatClearance(float, cursorY);
        if (
          (content.length > 0 || placedItems > 0) &&
          Math.max(cursorY + (clearance ?? 0) + carrier.rect.height, floatBottom(float, cursorY)) >
            bodyLimit([]) + FIT_EPSILON_PT
        ) {
          deferItem = true;
          break;
        }
        placeFloat(float, i, rect => {
          const laid = layoutDrawingShape(item, rect.y, bodyWidth, i, options, imageMap);
          return { ...laid, rect: { ...laid.rect, x: rect.x, y: rect.y } };
        });
        placeParagraph({
          ...layoutParagraph(
            ANCHOR_CARRIER,
            cursorY,
            bodyWidth,
            options,
            pageContext,
            imageMap,
            listMarkers
          ),
          sourceIndex: i
        });
        break;
      }
      case "chart":
      case "chartEx": {
        const laid = layoutChart(item, cursorY, bodyWidth, i);
        content.push(laid);
        cursorY = laid.rect.y + laid.rect.height;
        break;
      }
      case "sdt": {
        const laid = layoutSdt(item, cursorY, bodyWidth, i, options, imageMap);
        content.push(laid);
        cursorY = laid.rect.y + laid.rect.height;
        break;
      }
      case "math": {
        const laid = layoutMath(item, cursorY, bodyWidth, i, options);
        content.push(laid);
        cursorY = laid.rect.y + laid.rect.height;
        break;
      }
      case "checkBox": {
        const laid = layoutCheckBox(item, cursorY, i, options);
        content.push(laid);
        cursorY = laid.rect.y + laid.rect.height;
        break;
      }
      case "tableOfContents": {
        // Its entries are paragraphs, placed — and split across pages — one by
        // one like any other; `groupTableOfContents` gathers each page's share
        // back into one block. Placed whole, a table of contents longer than a
        // page ran off its page and everything after it started a page early.
        const laid = layoutTableOfContents(item, 0, bodyWidth, i, options, imageMap);
        for (const entry of laid.entries) {
          if (!placeParagraph({ ...entry, rect: { ...entry.rect, y: cursorY } })) {
            stopFlow = true;
          }
        }
        break;
      }
      case "altChunk": {
        const laid = layoutAltChunk(item, cursorY, bodyWidth, i);
        content.push(laid);
        cursorY = laid.rect.y + laid.rect.height;
        break;
      }
      case "opaqueDrawing": {
        const laid = layoutOpaqueDrawing(item, cursorY, bodyWidth, i);
        content.push(laid);
        cursorY = laid.rect.y + laid.rect.height;
        break;
      }
      default: {
        // Compile-time exhaustiveness check. Adding a new variant to
        // `BodyContent` triggers a TypeScript error here until a
        // corresponding case + layout function are added above. This
        // replaces the previous "Skip unsupported types" silent drop.
        const _exhaustive: never = item;
        throw new Error(
          `layoutDocumentFull: unhandled BodyContent variant ${
            (_exhaustive as { type: string }).type
          }`
        );
      }
    }

    if (deferItem) {
      // It starts the next page, where its own page break is not forced.
      break;
    }
    if (item.type !== "table") {
      pageNotes.add(itemFootnoteRefs);
    }

    // A block that was split leaves its remainder in `carriedOut`; the item
    // itself is done either way, so the cursor always advances and the flow
    // cannot revisit it.
    flow.nextItem = i + 1;
    placedItems++;
    // Keep the quote beside the table it comments on, which is where the
    // source puts it. End the page *after* the quote so the following `---`
    // cannot become the lower ruling line of a table row in Preview's PDF
    // layout analysis.
    if (i === quoteRunEnd && carriedOut.length === 0) {
      stopFlow = true;
      stop = "page";
      delete flow.endPageAfterItem;
    }
  }

  // `w:keepNext` after the fact, rather than by looking ahead.
  //
  // A predictive check has to guess how much of the *next* block counts as
  // "started", and it has to do so for a whole chain of `keepNext` paragraphs
  // at once. Getting that wrong is expensive in both directions: demanding the
  // entire chain fit gave every heading in a run its own page, and the repeated
  // trial layouts were quadratic in the chain length.
  //
  // Deciding afterwards needs no prediction. If the page ends on a `keepNext`
  // paragraph and there is more to come, hand it to the next page — repeatedly,
  // so a run of headings travels together. The page must keep at least one
  // block, since a heading alone on a page is better than an empty one.
  const more = flow.nextItem < doc.body.length || carriedOut.length > 0;
  // A `keepNext` paragraph broken at the foot already ends on the next page,
  // beside what follows it: only a paragraph ending here is moved (pages2
  // P36f: a keepNext paragraph of 9 lines with room for 2 sets 2 here).
  const lastBlock = content[content.length - 1];
  const carriedOn = carriedOut[0];
  const brokenAtFoot =
    lastBlock?.type === "paragraph" &&
    carriedOn?.kind === "paragraph" &&
    carriedOn.continued === true &&
    carriedOn.block.sourceIndex === lastBlock.sourceIndex;
  if (more && !brokenAtFoot && stop !== "region") {
    // Length of the trailing run of `keepNext` paragraphs.
    let run = 0;
    while (run < content.length) {
      const block = content[content.length - 1 - run];
      if (block.type !== "paragraph" || block.sourceIndex < 0) {
        break;
      }
      if (!keepsWithNext(block.sourceIndex)) {
        break;
      }
      run++;
    }
    // Move the whole run down together — but only if something that is not
    // part of it stays behind. When the entire page is one `keepNext` run the
    // request cannot be met (a page has to hold something), and Word likewise
    // lets the break stand rather than emitting a page per paragraph.
    // A region below an earlier section, or a later column, is not the whole
    // page: Word moves a heading alone at its foot (pages2 R12j).
    const pageHoldsMore = regionTop > 0 || (region?.column ?? 0) > 0;
    if (run > 0 && (run < content.length || pageHoldsMore)) {
      // Word keeps the last lines of a `keepNext` paragraph with what follows
      // and leaves the rest of it here, as long as each part can stand on its
      // own; paragraphs of the run after it that cannot be broken go along
      // whole (pages2 P35i: a six-line keepNext paragraph at the foot of a
      // column whose next paragraph cannot start there sets four lines and
      // takes two along; P36e: one line kept with the next goes along whole
      // and the paragraph before it breaks).
      for (let k = content.length - 1; k >= content.length - run; k--) {
        const block = content[k] as LayoutParagraph;
        const rules = breakRules(block.sourceIndex);
        const least = rules.widowControl ? 2 : 1;
        if (rules.keepLines || block.lines.length < 2 * least) {
          continue;
        }
        const cut = sliceParagraphAtLine(
          trimmedAfter.get(block) ?? block,
          block.lines.length - least
        );
        const moved = content.splice(k + 1);
        content[k] = cut.head;
        cursorY = cut.head.rect.y + cut.head.rect.height;
        for (const piece of [...moved].reverse()) {
          const whole = trimmedAfter.get(piece as LayoutParagraph) ?? (piece as LayoutParagraph);
          pageNotes.remove(whole.lineNoteIds?.flat() ?? []);
          carriedOut.unshift({ kind: "paragraph", block: whole });
        }
        pageNotes.remove(cut.tail.lineNoteIds?.flat() ?? []);
        carriedOut.unshift({ kind: "paragraph", block: cut.tail, continued: true });
        run = 0;
        break;
      }
    }
    if (run > 0 && (run < content.length || pageHoldsMore)) {
      const moved = content.splice(content.length - run, run);
      cursorY = moved[0].rect.y;
      for (let k = moved.length - 1; k >= 0; k--) {
        const block =
          trimmedAfter.get(moved[k] as LayoutParagraph) ?? (moved[k] as LayoutParagraph);
        // Their notes go with them.
        pageNotes.remove(block.lineNoteIds?.flat() ?? []);
        carriedOut.unshift({ kind: "paragraph", block });
      }
    }
  }

  // A column leaves its notes to the page it is part of (see `composePage`).
  const footnotes = region
    ? { laid: [], separator: undefined, continued: [] }
    : layoutFootnoteArea(pageNotes, geometry, cursorY);

  return {
    bottom: cursorY,
    ...(stop ? { stop } : {}),
    notes: pageNotes,
    page: {
      pageNumber,
      geometry,
      content: groupTableOfContents(doc, content),
      ...(header.length > 0 ? { header } : {}),
      ...(footer.length > 0 ? { footer } : {}),
      ...(footnotes.laid.length > 0 ? { footnoteArea: footnotes.laid } : {}),
      ...(footnotes.laid.length > 0 && footnotes.separator
        ? { footnoteSeparator: footnotes.separator }
        : {})
    },
    continuedNotes: footnotes.continued,
    carriedContent: carriedOut
  };
}

/**
 * A footnote as the footnote area holds it: its paragraphs stacked from y 0.
 *
 * `continued` marks the remainder of a note split at the foot of an earlier
 * page, or a whole note that page had no room for; either way its reference is
 * not on the page that shows it, so the area opens with the continuation
 * separator.
 */
interface LaidNote {
  readonly id: number;
  readonly paragraphs: readonly LayoutParagraph[];
  /** How each paragraph may break, parallel to `paragraphs`. */
  readonly rules: readonly ParagraphBreakRules[];
  /** Height of the whole note. */
  readonly height: number;
  /**
   * The least of the note that has to share a page with its reference: Word
   * may continue the rest on the next page, but only where the note's own
   * widow/orphan and keep-lines rules allow a break.
   */
  readonly minHeight: number;
  readonly continued: boolean;
}

/** Lines left in a queue of notes — what an overflow page must shrink. */
function totalNoteLines(notes: readonly LaidNote[]): number {
  let total = 0;
  for (const note of notes) {
    for (const p of note.paragraphs) {
      total += p.lines.length;
    }
  }
  return total;
}

/** `paragraphs` stacked from y 0, with their total height. */
function stackParagraphs(paragraphs: readonly LayoutParagraph[]): {
  paragraphs: LayoutParagraph[];
  height: number;
} {
  const out: LayoutParagraph[] = [];
  let cursor = 0;
  for (const p of paragraphs) {
    out.push({ ...p, rect: { ...p.rect, y: cursor } });
    cursor += p.rect.height;
  }
  return { paragraphs: out, height: cursor };
}

/** How a paragraph may break across pages (`w:widowControl`, `w:keepLines`). */
function paragraphBreakRules(para: Paragraph): ParagraphBreakRules {
  if (!activeStyles) {
    return { widowControl: true, keepLines: false };
  }
  const props = activeStyles.paragraph(para).paragraphProperties;
  return { widowControl: props.widowControl !== false, keepLines: props.keepLines === true };
}

/** A note built from laid-out paragraphs; see {@link LaidNote.minHeight}. */
function makeLaidNote(
  id: number,
  paragraphs: readonly LayoutParagraph[],
  rules: readonly ParagraphBreakRules[],
  continued: boolean
): LaidNote {
  const stacked = stackParagraphs(paragraphs);
  const first = stacked.paragraphs[0];
  let minHeight = stacked.height;
  if (first && first.lines.length > 0) {
    const { keepLines, widowControl } = rules[0]!;
    const lines = first.lines.length;
    // Lines that must stay together at the head of the first paragraph.
    const keep = keepLines ? lines : widowControl ? (lines >= 4 ? 2 : lines) : 1;
    if (keep < lines) {
      const line = first.lines[keep - 1]!;
      minHeight = line.y + line.height;
    } else if (stacked.paragraphs.length > 1) {
      // A whole first paragraph counts with its space after: Word moves a
      // reference to the next page when only the text of it would fit (final6
      // Qb2, 31.8pt left below the reference, 35.9 needed with the separator).
      minHeight = first.rect.height;
    }
  }
  return {
    id,
    paragraphs: stacked.paragraphs,
    rules,
    height: stacked.height,
    minHeight,
    continued
  };
}

/** Lay out footnote `def` at `width`. */
function layoutNote(
  def: FootnoteDef,
  width: number,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>
): LaidNote {
  const paragraphs: LayoutParagraph[] = [];
  const rules: ParagraphBreakRules[] = [];
  for (const para of def.content) {
    if (para.type === "paragraph") {
      paragraphs.push(layoutParagraph(para, 0, width, options, undefined, imageMap));
      rules.push(paragraphBreakRules(para));
    }
  }
  return makeLaidNote(def.id, paragraphs, rules, false);
}

/**
 * The paragraph Word draws a footnote separator in: the document's own
 * `separator` / `continuationSeparator` note, or — when the file declares none
 * — the one Word writes, single-spaced with no space after.
 */
/** A note separator's paragraph where the document defines none: one single-spaced line. */
const DEFAULT_SEPARATOR_PARAGRAPH: Paragraph = {
  type: "paragraph",
  properties: { spacing: { after: 0, line: 240, lineRule: "auto" } },
  children: []
};

function layoutSeparator(
  doc: DocxDocument,
  kind: "separator" | "continuationSeparator",
  width: number,
  options: FullLayoutOptions | undefined
): LayoutParagraph[] {
  const def = doc.footnotes?.find(note => note.type === kind);
  const paragraphs: readonly Paragraph[] = def?.content.length
    ? def.content
    : [DEFAULT_SEPARATOR_PARAGRAPH];
  return paragraphs.map(para => layoutParagraph(para, 0, width, options));
}

/**
 * The footnotes of one page, as the page is filled.
 *
 * Word keeps a note on the page of its reference: the line holding a reference
 * only fits if the note fits below the text too, so the room a page has for
 * body text shrinks as notes are referenced. Only the last note on a page may
 * be split, the rest of it continuing at the top of the next page's footnote
 * area, under the continuation separator.
 */
interface PageNotes {
  /** Notes on the page, in order; those continued from earlier pages first. */
  readonly notes: LaidNote[];
  /**
   * The space the footnote area claims if `extra` notes join it. Notes already
   * on the page claim their whole height: Word sets no more body text on a page
   * whose notes have to continue on the next (final6 Qb1: a note longer than a
   * page leaves its page only the paragraph referencing it). Of the notes
   * `extra` adds, the last needs only {@link LaidNote.minHeight}, so the line
   * referencing it may still close the page.
   */
  reserve(extra?: readonly LaidNote[]): number;
  /** The notes `ids` adds to the page — each laid out once, none twice. */
  newNotes(ids: readonly number[]): LaidNote[];
  add(ids: readonly number[]): void;
  remove(ids: readonly number[]): void;
}

function createPageNotes(
  doc: DocxDocument,
  carried: readonly LaidNote[],
  width: number,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>,
  cache: Map<string, LaidNote>,
  noteColumns?: NoteColumns
): PageNotes & {
  separatorHeight(kind: "separator" | "continuationSeparator"): number;
  readonly columns: NoteColumns | undefined;
  useColumns(columns: NoteColumns): void;
  startColumn(): void;
} {
  let columns: NoteColumns | undefined;
  const notes = [...carried];
  /** How many notes the page held when the column being filled began. */
  let columnStart = 0;
  const separators = new Map<string, number>();
  const separatorHeight = (kind: "separator" | "continuationSeparator"): number => {
    let height = separators.get(kind);
    if (height === undefined) {
      height = stackParagraphs(layoutSeparator(doc, kind, width, options)).height;
      separators.set(kind, height);
    }
    return height;
  };
  const defs = new Map<number, FootnoteDef>();
  for (const def of doc.footnotes ?? []) {
    if ((def.type ?? "normal") === "normal") {
      defs.set(def.id, def);
    }
  }
  const newNotes = (ids: readonly number[]): LaidNote[] => {
    const out: LaidNote[] = [];
    for (const id of ids) {
      if (notes.some(n => n.id === id && !n.continued) || out.some(n => n.id === id)) {
        continue;
      }
      const def = defs.get(id);
      if (!def) {
        continue;
      }
      const key = `${id}@${width}@${activeGridPitch ?? 0}`;
      let note = cache.get(key);
      if (!note) {
        note = layoutNote(def, width, options, imageMap);
        cache.set(key, note);
      }
      out.push(note);
    }
    return out;
  };
  /**
   * In a section of several columns Word sets the notes in the columns too, at
   * the column width, flowing from one column to the next (pages2 R9e) — on a
   * page that also holds a section of one column as well (R2a, either order).
   * Unequal columns are not measured; the notes take the first one's width.
   */
  const fullWidth = width;
  const useColumns = (next: NoteColumns): void => {
    // Until a note is referenced the area follows the section being filled:
    // a note referenced in a section of one column spans the page although a
    // section of several came before it on the page (pages2 R11j).
    if (columns && notes.length === 0 && next.columns.length < 2) {
      columns = undefined;
      width = fullWidth;
      separators.clear();
      return;
    }
    if ((columns && notes.length > 0) || next.columns.length < 2 || columns === next) {
      return;
    }
    columns = next;
    width = next.columns[0]!.width;
    separators.clear();
    const fresh = notes.filter(note => !note.continued);
    const kept = notes.filter(note => note.continued);
    notes.splice(0, notes.length, ...kept);
    notes.push(...newNotes(fresh.map(note => note.id)));
  };
  if (noteColumns) {
    useColumns(noteColumns);
  }
  return {
    notes,
    separatorHeight,
    get columns() {
      return columns;
    },
    useColumns,
    startColumn() {
      columnStart = notes.length;
    },
    reserve(extra = []) {
      const all = extra.length > 0 ? [...notes, ...extra] : notes;
      if (all.length === 0) {
        return 0;
      }
      let total = separatorHeight(all[0]!.continued ? "continuationSeparator" : "separator");
      if (columns) {
        // The last note may continue on the next page when the column filling
        // now referenced it: text after its reference in that column still
        // fits (pages2 R2a). One referenced in an earlier column claims its
        // whole height, as on a page of one column (R9e).
        const partial = extra.length > 0 || all.length > columnStart;
        return total + noteColumnsHeight(all, columns, partial);
      }
      for (let i = 0; i < all.length - 1; i++) {
        total += all[i]!.height;
      }
      const last = all[all.length - 1]!;
      return total + (extra.length > 0 ? last.minHeight : last.height);
    },
    newNotes,
    add(ids) {
      notes.push(...newNotes(ids));
    },
    remove(ids) {
      for (const id of ids) {
        const at = notes.findIndex(n => n.id === id && !n.continued);
        if (at >= 0) {
          notes.splice(at, 1);
        }
      }
    }
  };
}

/** The note ids referenced by the lines of `block`, its table cells included. */
function collectLaidNoteIds(block: PageContent, out: Set<number>): void {
  if (block.type === "paragraph") {
    for (const ids of block.lineNoteIds ?? []) {
      for (const id of ids) {
        out.add(id);
      }
    }
  } else if (block.type === "table") {
    for (const cell of block.cells) {
      for (const inner of cell.content) {
        collectLaidNoteIds(inner, out);
      }
    }
  }
}

/** The columns a page's footnotes are set in, content-relative. */
interface NoteColumns {
  readonly columns: readonly { readonly x: number; readonly width: number }[];
  readonly contentHeight: number;
}

/**
 * `notes` flowed through `columns`, each `height` high: a note's paragraphs fill
 * a column and continue at the top of the next, breaking where their widow and
 * keep rules allow. What the last column cannot hold is `continued`. `force`
 * places at least a line in the first column, as the top of an empty area must.
 */
function flowNotesInColumns(
  notes: readonly LaidNote[],
  height: number,
  columns: NoteColumns,
  force: boolean
): { columns: LayoutParagraph[][]; used: number; continued: LaidNote[] } {
  const out: LayoutParagraph[][] = [[]];
  const continued: LaidNote[] = [];
  let used = 0;
  let tallest = 0;
  for (const note of notes) {
    if (continued.length > 0) {
      continued.push({ ...note, continued: true });
      continue;
    }
    const queue = note.paragraphs.map((p, k) => ({ p, rules: note.rules[k]! }));
    let next = 0;
    while (next < queue.length) {
      const { p, rules } = queue[next]!;
      const column = out[out.length - 1]!;
      const x = columns.columns[out.length - 1]!.x;
      if (p.rect.height <= height - used + FIT_EPSILON_PT) {
        column.push({ ...p, rect: { ...p.rect, x: p.rect.x + x, y: used } });
        used += p.rect.height;
        tallest = Math.max(tallest, used);
        next++;
        continue;
      }
      const cut = splitLayoutParagraph(
        { ...p, rect: { ...p.rect, y: 0 } },
        height - used,
        columns.contentHeight,
        force && out.length === 1 && column.length === 0,
        rules,
        true
      );
      if (cut.head) {
        column.push({ ...cut.head, rect: { ...cut.head.rect, x: cut.head.rect.x + x, y: used } });
        tallest = Math.max(tallest, used + cut.head.rect.height);
      }
      if (cut.tail) {
        queue[next] = { p: cut.tail, rules };
      } else {
        next++;
      }
      if (out.length === columns.columns.length) {
        const rest = queue.slice(next);
        if (rest.length > 0) {
          continued.push(
            makeLaidNote(
              note.id,
              rest.map(item => item.p),
              rest.map(item => item.rules),
              true
            )
          );
        }
        break;
      }
      out.push([]);
      used = 0;
    }
  }
  return { columns: out, used: tallest, continued };
}

/**
 * The height a footnote area set in columns needs for `notes`: the least at
 * which they all fit, or with `lastPartial` at which all but the last fit and
 * the last at least starts (the rest of it may continue on the next page).
 */
function noteColumnsHeight(
  notes: readonly LaidNote[],
  columns: NoteColumns,
  lastPartial: boolean
): number {
  const last = notes[notes.length - 1]!;
  const fits = (height: number): boolean => {
    const flow = flowNotesInColumns(notes, height, columns, false);
    if (flow.continued.length === 0) {
      return true;
    }
    return (
      lastPartial &&
      flow.continued.length === 1 &&
      flow.continued[0]!.id === last.id &&
      flow.continued[0]!.height < last.height - 0.01
    );
  };
  let hi = 0;
  for (const note of notes) {
    hi += note.height;
  }
  if (!fits(hi)) {
    return hi;
  }
  let lo = 0;
  while (hi - lo > 0.05) {
    const mid = (lo + hi) / 2;
    if (fits(mid)) {
      hi = mid;
    } else {
      lo = mid;
    }
  }
  return flowNotesInColumns(notes, hi, columns, false).used;
}

/**
 * Lay out a page's footnote area at the foot of its text area.
 *
 * Notes are taken in order while they fit in the space the body left; the one
 * that does not is split where its paragraphs allow, and it and everything
 * after it continue on the next page (`continued`). The first note is placed
 * even when nothing of it fits — at the top of an otherwise empty area there is
 * nowhere better for it — so the queue always shrinks.
 */
function layoutFootnoteArea(
  pageNotes: PageNoteSet,
  geometry: PageGeometry,
  bodyBottom: number
): {
  laid: LayoutParagraph[];
  separator: LayoutPage["footnoteSeparator"] | undefined;
  continued: LaidNote[];
} {
  const notes = pageNotes.notes;
  if (notes.length === 0) {
    return { laid: [], separator: undefined, continued: [] };
  }
  const kind = notes[0]!.continued ? "continuationSeparator" : "separator";
  const separatorHeight = pageNotes.separatorHeight(kind);
  const available = geometry.contentHeight - bodyBottom - separatorHeight;
  if (pageNotes.columns) {
    // The columns start level, at the height the notes need. When the last
    // cannot be held whole it continues on the next page, and the area is only
    // as tall as the rest and its start need (pages2 R2a).
    const whole = noteColumnsHeight(notes, pageNotes.columns, false);
    const height =
      whole <= available + FIT_EPSILON_PT
        ? whole
        : Math.min(available, noteColumnsHeight(notes, pageNotes.columns, true));
    const flow = flowNotesInColumns(notes, height, pageNotes.columns, true);
    const top = geometry.marginTop + geometry.contentHeight - height;
    return {
      laid: flow.columns.flat().map(p => ({ ...p, rect: { ...p.rect, y: top + p.rect.y } })),
      separator: { y: top - separatorHeight / 2, kind },
      continued: flow.continued
    };
  }
  const placed: LayoutParagraph[] = [];
  const continued: LaidNote[] = [];
  let used = 0;
  for (let i = 0; i < notes.length; i++) {
    const note = notes[i]!;
    if (continued.length > 0) {
      continued.push({ ...note, continued: true });
      continue;
    }
    if (used + note.height <= available + FIT_EPSILON_PT) {
      for (const p of note.paragraphs) {
        placed.push({ ...p, rect: { ...p.rect, y: used + p.rect.y } });
      }
      used += note.height;
      continue;
    }
    // Split: whole paragraphs while they fit, then the paragraph that does not.
    // At the top of an empty area at least one line is placed (`atPageTop`).
    const head: LayoutParagraph[] = [];
    const tail: LayoutParagraph[] = [];
    const tailRules: ParagraphBreakRules[] = [];
    let room = available - used;
    // Once a paragraph has been cut — or kept without the space after it — the
    // area is full, and the rest of the note continues on the next page.
    let full = false;
    note.paragraphs.forEach((p, k) => {
      const rules = note.rules[k]!;
      if (tail.length > 0 || full) {
        tail.push(p);
        tailRules.push(rules);
        return;
      }
      if (p.rect.height <= room + FIT_EPSILON_PT) {
        head.push(p);
        room -= p.rect.height;
        return;
      }
      const cut = splitLayoutParagraph(
        { ...p, rect: { ...p.rect, y: 0 } },
        room,
        geometry.contentHeight,
        placed.length === 0 && head.length === 0,
        rules,
        true
      );
      if (cut.head) {
        head.push(cut.head);
      }
      if (cut.tail) {
        tail.push(cut.tail);
        tailRules.push(rules);
      }
      full = true;
    });
    if (head.length > 0) {
      const stacked = stackParagraphs(head);
      for (const p of stacked.paragraphs) {
        placed.push({ ...p, rect: { ...p.rect, y: used + p.rect.y } });
      }
      used += stacked.height;
    }
    if (tail.length > 0) {
      continued.push(makeLaidNote(note.id, tail, tailRules, true));
    }
  }
  // Bottom-aligned: the area ends where the text area does.
  const areaBottom = geometry.marginTop + geometry.contentHeight;
  const top = areaBottom - used;
  return {
    laid: placed.map(p => ({ ...p, rect: { ...p.rect, y: top + p.rect.y } })),
    // The rule sits in the separator paragraph, above the first note.
    separator: { y: top - separatorHeight / 2, kind },
    continued
  };
}

/**
 * Walk a `BodyContent` item's run-level descendants and append every
 * `FootnoteRefContent` id to `out`, in document order. Recurses into
 * the few container variants whose children embed paragraphs (textBox,
 * drawingShape, sdt, tableOfContents, table cells).
 */
function collectFootnoteRefsFromBody(item: BodyContent, out: number[]): void {
  switch (item.type) {
    case "paragraph":
      collectFootnoteRefsFromParagraph(item, out);
      return;
    case "table":
      for (const r of item.rows) {
        for (const c of r.cells) {
          // Word shows nothing a vertical merge's continuing cell holds, its
          // note references included (pages2 R4j: no note on any page).
          if (c.properties?.verticalMerge === "continue") {
            continue;
          }
          for (const inner of c.content) {
            collectFootnoteRefsFromBody(inner, out);
          }
        }
      }
      return;
    case "textBox":
      for (const child of item.content) {
        collectFootnoteRefsFromBody(child, out);
      }
      return;
    case "drawingShape":
      if (item.textContent) {
        for (const child of item.textContent) {
          collectFootnoteRefsFromBody(child, out);
        }
      }
      return;
    case "sdt":
      for (const child of item.content) {
        if (child && typeof child === "object" && "type" in child) {
          collectFootnoteRefsFromBody(child as BodyContent, out);
        }
      }
      return;
    case "tableOfContents":
      if (item.cachedParagraphs) {
        for (const para of item.cachedParagraphs) {
          collectFootnoteRefsFromBody(para, out);
        }
      }
      return;
    case "floatingImage":
    case "math":
    case "checkBox":
    case "chart":
    case "chartEx":
    case "altChunk":
    case "opaqueDrawing":
      return;
    default: {
      const _exhaustive: never = item;
      void _exhaustive;
    }
  }
}

function collectFootnoteRefsFromParagraph(para: Paragraph, out: number[]): void {
  // Only references the final view draws: a note whose reference was deleted
  // (or moved away) has nothing in the text pointing at it.
  for (const run of finalViewRuns(para.children)) {
    collectFootnoteRefsFromRun(run, out);
  }
}

function collectFootnoteRefsFromRun(run: Run, out: number[]): void {
  if (!run || !Array.isArray(run.content)) {
    return;
  }
  for (const c of run.content) {
    if (c.type === "footnoteRef") {
      out.push(c.id);
    }
  }
}

/**
 * The header (or footer) a page shows, per ECMA-376 §17.10: the `"first"` one
 * on a section's first page when it has `titlePage` (a parity blank page,
 * `pageInSection` 0, never takes it), the `"even"` one on even pages when
 * `evenAndOddHeaders` is set, otherwise the `"default"` one. A type that is in
 * effect for none of the sections so far shows nothing — Word does not fall
 * back to the default (a first page without a first footer is bare).
 */
function pickHeaderFooterRef(
  refs: ReadonlyMap<HeaderFooterType, string>,
  displayedPage: number,
  pageInSection: number,
  titlePage: boolean,
  evenAndOdd: boolean
): string | undefined {
  // `titlePage` is per section — its "first" header belongs to the first page of
  // *that* section. Even/odd follows the *displayed* page number, as Word does:
  // a section restarting at 2 opens on an even header.
  if (titlePage && pageInSection === 1) {
    return refs.get("first");
  }
  if (evenAndOdd && displayedPage % 2 === 0) {
    return refs.get("even");
  }
  return refs.get("default");
}

/**
 * Lay out the paragraphs and tables of the resolved header for a page.
 *
 * Resolution order: first/even/default per `pickHeaderFooterRef`.
 * The header band's local y-axis starts at the section's
 * `pgMar.header` (in pt) below the page top, mirroring Word's
 * "Header from top" setting; renderers consume the resulting layout-y
 * directly as a page-y offset.
 *
 * Tables in header content are laid out with the same `layoutTable`
 * that body content uses, and surfaced via the union type on
 * `LayoutPage.header` so renderers can pick them up alongside
 * paragraphs without a special path.
 */
function layoutHeader(
  doc: DocxDocument,
  slot: PageSlot,
  geometry: PageGeometry,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>,
  totals?: { readonly numPages: number; readonly sectionPages: number }
): (LayoutParagraph | LayoutTable)[] {
  const sectionProps = slot.props;
  // Word prints the blank page of an odd/even-page break bare.
  if (slot.pageInSection === 0) {
    return [];
  }
  // `titlePage` picks the "first" reference on the first page *of its section*,
  // not of the document.
  const titlePage = sectionProps?.titlePage === true;
  const evenAndOdd = doc.settings?.evenAndOddHeaders === true;
  const ref = pickHeaderFooterRef(
    slot.headers,
    slot.displayedPage,
    slot.pageInSection,
    titlePage,
    evenAndOdd
  );
  if (!ref) {
    return [];
  }
  const part = doc.headers?.get(ref);
  if (!part) {
    return [];
  }
  const headerOffsetPt = geometry.headerOffset;
  return withPageFields(pageFieldValues(slot, totals), () =>
    withGridPitch(undefined, () =>
      layoutHeaderFooterChildren(part.content.children, headerOffsetPt, geometry, options, imageMap)
    )
  );
}

function layoutFooter(
  doc: DocxDocument,
  slot: PageSlot,
  geometry: PageGeometry,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>,
  totals?: { readonly numPages: number; readonly sectionPages: number }
): (LayoutParagraph | LayoutTable)[] {
  const sectionProps = slot.props;
  // Word prints the blank page of an odd/even-page break bare.
  if (slot.pageInSection === 0) {
    return [];
  }
  const titlePage = sectionProps?.titlePage === true;
  const evenAndOdd = doc.settings?.evenAndOddHeaders === true;
  const ref = pickHeaderFooterRef(
    slot.footers,
    slot.displayedPage,
    slot.pageInSection,
    titlePage,
    evenAndOdd
  );
  if (!ref) {
    return [];
  }
  const part = doc.footers?.get(ref);
  if (!part) {
    return [];
  }
  // Footer band starts at `pageHeight - pgMar.footer` so layout-y is
  // already a page-absolute coordinate (matching the header path,
  // where `pgMar.header` is the absolute offset of the band from the
  // page top). Renderers consume both bands with the same
  // "treat layout-y as page-y" rule.
  const footerOffsetPt = geometry.height - geometry.footerOffset;
  return withPageFields(pageFieldValues(slot, totals), () =>
    layoutHeaderFooterChildren(part.content.children, footerOffsetPt, geometry, options, imageMap)
  );
}

function layoutHeaderFooterChildren(
  children: readonly (Paragraph | Table)[],
  initialCursorY: number,
  geometry: PageGeometry,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>
): (LayoutParagraph | LayoutTable)[] {
  const out: (LayoutParagraph | LayoutTable)[] = [];
  let cursor = initialCursorY;
  for (let idx = 0; idx < children.length; idx++) {
    const child = children[idx];
    if (child.type === "paragraph") {
      const laid = layoutParagraph(
        child,
        cursor,
        geometry.contentWidth,
        options,
        undefined,
        imageMap
      );
      out.push(laid);
      cursor = laid.rect.y + laid.rect.height;
    } else if (child.type === "table") {
      const laid = layoutTable(child, cursor, geometry.contentWidth, idx, options, imageMap);
      out.push(laid);
      cursor = laid.rect.y + laid.rect.height;
    }
  }
  return out;
}

function computePageGeometry(
  sectionProps: DocxDocument["sectionProperties"],
  override: PageGeometryOverride | undefined,
  gutterAt: GutterSide
): PageGeometry {
  const widthTwips = sectionProps?.pageSize?.width ?? DEFAULT_PAGE_WIDTH_TWIPS;
  const heightTwips = sectionProps?.pageSize?.height ?? DEFAULT_PAGE_HEIGHT_TWIPS;
  const sectionWidth = twipsToPt(widthTwips);
  const sectionHeight = twipsToPt(heightTwips);
  const sectionMarginTop = twipsToPt(sectionProps?.margins?.top ?? DEFAULT_PAGE_MARGIN_TWIPS);
  const sectionMarginBottom = twipsToPt(sectionProps?.margins?.bottom ?? DEFAULT_PAGE_MARGIN_TWIPS);
  const sectionMarginLeft = twipsToPt(sectionProps?.margins?.left ?? DEFAULT_PAGE_MARGIN_TWIPS);
  const sectionMarginRight = twipsToPt(sectionProps?.margins?.right ?? DEFAULT_PAGE_MARGIN_TWIPS);
  const gutter = twipsToPt(sectionProps?.margins?.gutter ?? 0);
  const gutterOn = (side: GutterSide): number => (gutterAt === side ? gutter : 0);
  // Header / footer band offsets. Word's default `pgMar.header` /
  // `pgMar.footer` is 720 twips (0.5") — the same default used by the
  // header / footer layout helpers historically.
  const sectionHeaderOffset = twipsToPt(sectionProps?.margins?.header ?? 720);
  const sectionFooterOffset = twipsToPt(sectionProps?.margins?.footer ?? 720);

  // Per-axis override: callers (PDF bridge, custom hosts) may want to
  // pin the page size or margin on one axis without disturbing the
  // others — `pageWidth` doesn't imply overriding margins, etc.
  const width = override?.pageWidth ?? sectionWidth;
  const height = override?.pageHeight ?? sectionHeight;
  const marginTop = override?.marginTop ?? sectionMarginTop + gutterOn("top");
  const marginBottom = override?.marginBottom ?? sectionMarginBottom + gutterOn("bottom");
  const marginLeft = override?.marginLeft ?? sectionMarginLeft + gutterOn("left");
  const marginRight = override?.marginRight ?? sectionMarginRight + gutterOn("right");
  const headerOffset = override?.headerMargin ?? sectionHeaderOffset;
  const footerOffset = override?.footerMargin ?? sectionFooterOffset;

  return {
    width,
    height,
    marginTop,
    marginBottom,
    marginLeft,
    marginRight,
    contentWidth: width - marginLeft - marginRight,
    contentHeight: height - marginTop - marginBottom,
    headerOffset,
    footerOffset
  };
}

function computeSectionBreaks(
  layout: LayoutResult,
  itemPage: ReadonlyMap<number, number>
): number[] {
  const breaks: number[] = [0]; // First section starts at page 0
  let prevSection = 0;
  for (let i = 0; i < layout.contentPages.length; i++) {
    const section = layout.contentSections[i];
    if (section > prevSection) {
      // Prefer the page pass 2 actually placed the item on; fall back to the
      // paginator's estimate for an item pass 2 never emitted.
      const page = itemPage.get(i) ?? layout.contentPages[i];
      breaks.push(page - 1);
      prevSection = section;
    }
  }
  return breaks;
}

// =============================================================================
// Internal: Paragraph Layout
// =============================================================================

/** A resolved list marker for a numbered/bulleted paragraph. */
interface ListMarker {
  /** Marker text including trailing spacing (e.g. "•  ", "1.  ", "a.  "). */
  readonly text: string;
  /** Left indent in points for the list level. */
  readonly indentPt: number;
  /** The level's own run properties (`w:lvl/w:rPr`): the face the marker is set in. */
  readonly runProperties?: Run["properties"];
}

/** What a paragraph's neighbours change about its own spacing and decoration. */
interface ContextualSpacing {
  /** Suppress space-before (`w:contextualSpacing`, same style above). */
  readonly before: boolean;
  /** Suppress space-after (`w:contextualSpacing`, same style below). */
  readonly after: boolean;
  /** The paragraph above shares this one's borders and shading. */
  readonly mergeDecorationTop: boolean;
  /** The paragraph below shares this one's borders and shading. */
  readonly mergeDecorationBottom: boolean;
  /**
   * Points of this paragraph's space before already given by the paragraph
   * above's space after: Word separates two paragraphs by the larger of the two,
   * not their sum (Word for Mac 16: 10pt after and 10pt before leave 10pt).
   */
  readonly sharedBefore?: number;
}

/**
 * Resolve `w:contextualSpacing` — "ignore spacing above and below when using
 * identical styles" — for every paragraph in the document.
 *
 * Word suppresses a paragraph's own space-before when the paragraph directly
 * above uses the same style, and its space-after when the one below does. That
 * is what keeps a bulleted list tight while still separating the list as a whole
 * from surrounding body text; Word's `ListParagraph` style sets the flag. The
 * layout engine ignored it, so every list item carried its full spacing and
 * lists rendered looser than in Word.
 *
 * Adjacency is judged among *siblings*, so each table cell, text box and
 * footnote forms its own run of paragraphs, and a non-paragraph block (a nested
 * table, an image frame) between two paragraphs breaks the run.
 */
/** A paragraph's space before or after in points, as `layoutParagraph` reads it. */
function spacingPt(spacing: LineSpacing | undefined, side: "before" | "after"): number {
  if (side === "before" ? spacing?.beforeAutoSpacing : spacing?.afterAutoSpacing) {
    return 5;
  }
  const value = side === "before" ? spacing?.before : spacing?.after;
  return value != null ? twipsToPt(value) : 0;
}

/**
 * A line's height on a document grid of `pitch` points, from its single-spaced
 * height. The line takes as many whole pitches as its text needs; `auto`
 * spacing multiplies the pitch, not the text, and never shrinks the line below
 * those pitches; `atLeast` is a floor over them. `exact` stays off the grid.
 * Measured in Word 16.114 (pages2 P7, P8: 9–40pt text, multiples 0.83–2, pitches
 * 15.6 and 18pt).
 */
function gridLineHeightPt(single: number, pitch: number, spacing: LineSpacing | undefined): number {
  const pitches = Math.max(1, Math.ceil(single / pitch - 1e-6)) * pitch;
  if (spacing?.lineRule === "atLeast") {
    return Math.max(pitches, twipsToPt(spacing.line ?? 0));
  }
  return Math.max(pitches, spacing?.line ? (pitch * spacing.line) / 240 : pitch);
}

/**
 * Where a line's text sits in its box (see `WordLineMetricsInput.textBand`):
 * `auto` spacing puts its extra below the text and may shrink the box under
 * it, `atLeast` puts its extra above; on a document grid the text is centred in
 * the pitches it needs. Measured against Word 16.114's PDF output (pages2 P7,
 * P8). `exact` keeps the centred text.
 */
function lineTextBand(
  single: number,
  nominal: number,
  pitch: number | undefined,
  spacing: LineSpacing | undefined
): { textBand?: { top: number; height: number }; shrink?: boolean } {
  const rule = spacing?.line ? (spacing.lineRule ?? "auto") : "auto";
  if (rule === "exact") {
    return {};
  }
  if (pitch !== undefined) {
    if (rule !== "atLeast") {
      return {};
    }
    const pitches = Math.max(1, Math.ceil(single / pitch - 1e-6)) * pitch;
    return { textBand: { top: Math.max(0, nominal - pitches), height: pitches } };
  }
  if (rule === "atLeast") {
    return { textBand: { top: Math.max(0, nominal - single), height: single } };
  }
  return { textBand: { top: 0, height: Math.min(nominal, single) }, shrink: nominal < single };
}

function computeContextualSpacing(
  doc: DocxDocument,
  styles: StyleResolver
): Map<Paragraph, ContextualSpacing> {
  const out = new Map<Paragraph, ContextualSpacing>();

  /**
   * A paragraph's decoration reduced to a comparable key, or "" for none.
   *
   * Two adjacent paragraphs sharing a key are one decorated block: a two-
   * paragraph block quote is one bar, a code block is one frame.
   */
  const decorationKey = (item: BodyContent | Paragraph | Table | undefined): string => {
    if (!item || item.type !== "paragraph") {
      return "";
    }
    const effective = styles.paragraph(item).paragraphProperties;
    const borders = resolveParagraphBorders(effective.borders);
    const fill = resolveShadingFill(effective.shading);
    if (!borders && !fill) {
      return "";
    }
    return JSON.stringify({ borders: borders ?? null, fill: fill ?? null });
  };

  const visit = (items: readonly BodyContent[] | readonly (Paragraph | Table)[]): void => {
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type === "table") {
        for (const row of item.rows) {
          for (const cell of row.cells) {
            visit(cell.content);
          }
        }
        continue;
      }
      if (item.type === "tableOfContents") {
        visit(tocParagraphs(item));
        continue;
      }
      if (item.type !== "paragraph") {
        continue;
      }
      // A table of contents ends in its last paragraph.
      const before = items[i - 1];
      const previous = before?.type === "tableOfContents" ? tocParagraphs(before).at(-1) : before;
      const next = items[i + 1];

      // Contextual spacing: the flag is inheritable (that is how
      // `ListParagraph` carries it) and compares by style identity.
      const effective = styles.paragraph(item).paragraphProperties;
      const styleId = item.properties?.style;
      const sameStyle = (other: (typeof items)[number] | undefined): boolean =>
        other?.type === "paragraph" && other.properties?.style === styleId;
      const contextual = effective.contextualSpacing === true;

      // Decoration runs: a shared bar or frame must not be interrupted by the
      // space between the paragraphs it spans, so the shared edge is not inset.
      const own = decorationKey(item);
      const mergeTop = own !== "" && decorationKey(previous) === own;
      const mergeBottom = own !== "" && decorationKey(next) === own;

      const dropsBefore = contextual && sameStyle(previous);
      // The space the paragraph above leaves below it, unless its own
      // contextual spacing drops it.
      let sharedBefore = 0;
      if (!dropsBefore && previous?.type === "paragraph") {
        const above = styles.paragraph(previous).paragraphProperties;
        const aboveDropsAfter =
          above.contextualSpacing === true && previous.properties?.style === styleId;
        if (!aboveDropsAfter) {
          sharedBefore = Math.min(
            spacingPt(effective.spacing, "before"),
            spacingPt(above.spacing, "after")
          );
        }
      }

      if (!contextual && !mergeTop && !mergeBottom && sharedBefore <= 0) {
        continue;
      }
      out.set(item, {
        before: dropsBefore,
        after: contextual && sameStyle(next),
        mergeDecorationTop: mergeTop,
        mergeDecorationBottom: mergeBottom,
        ...(sharedBefore > 0 ? { sharedBefore } : {})
      });
    }
  };

  visit(doc.body);
  return out;
}

/**
 * Resolve list markers for every numbered / bulleted paragraph in the
 * document, in reading order (descending into tables), keyed by paragraph.
 *
 * Which paragraphs are list items, which level definition applies and what
 * number each item shows come from the shared numbering counter — the same
 * rules the HTML, Markdown and semantic converters use — so a list resumed
 * after an interrupting paragraph continues its count here too, as in Word.
 * Only presentation is decided here: the marker text (the level's `%N`
 * template, each placeholder formatted per its own level), the bullet glyph
 * and the indent.
 */
function computeListMarkers(doc: DocxDocument, styles: StyleResolver): Map<Paragraph, ListMarker> {
  const markers = new Map<Paragraph, ListMarker>();
  if (!doc.numberingInstances?.length || !doc.abstractNumberings?.length) {
    return markers;
  }
  const counter = createNumberingCounter(doc);

  const resolveParagraphMarker = (para: Paragraph): void => {
    const numbering = styles.numbering(para);
    if (!numbering) {
      return;
    }
    const { numId, level } = numbering;
    const levelDef = counter.levelDef(numId, level);
    if (!levelDef) {
      return;
    }
    const value = counter.next(numId, level);

    // Indent: prefer the numbering level's own `w:lvl/w:pPr/w:ind`, which is
    // where Word records a list's real geometry (and where any customised
    // list puts it). Only fall back to the conventional half inch per level
    // when the level definition is silent.
    const levelIndent = levelDef.paragraphProperties?.indent?.left;
    const indentPt = levelIndent != null ? twipsToPt(levelIndent) : (level + 1) * 36;

    if (levelDef.format === "bullet") {
      // Bullet symbol. Word authors bullets with Symbol/Wingdings private-use
      // code points (e.g. U+F0B7 ·, U+F0A7 ▪) that PDF standard fonts can't
      // render. Normalize the common ones to WinAnsi-renderable equivalents;
      // fall back to a round bullet when empty or unknown.
      markers.set(para, {
        text: `${normalizeBulletGlyph(levelDef.text)}  `,
        indentPt,
        runProperties: levelDef.runProperties
      });
      return;
    }

    // `%N` names level N-1's current number, formatted per that level, so a
    // multi-level template such as "%1.%2." reads "2.3." rather than "3.3.".
    const text = levelDef.text
      ? levelDef.text.replace(/%([1-9])/g, (_, digit: string) => {
          const ref = Number(digit) - 1;
          const shown = ref === level ? value : counter.current(numId, ref);
          return formatListCounter(shown, placeholderFormat(counter, numId, levelDef, ref));
        })
      : `${formatListCounter(value, placeholderFormat(counter, numId, levelDef, level))}.`;
    markers.set(para, { text: `${text}  `, indentPt, runProperties: levelDef.runProperties });
  };

  const walk = (items: readonly BodyContent[] | readonly (Paragraph | Table)[]): void => {
    for (const item of items) {
      if (item.type === "paragraph") {
        resolveParagraphMarker(item);
      } else if (item.type === "table") {
        for (const row of item.rows) {
          for (const cell of row.cells) {
            walk(cell.content);
          }
        }
      }
    }
  };
  walk(doc.body);
  return markers;
}

/** Normalize a Word bullet glyph to a WinAnsi-renderable equivalent. */
/**
 * Bullet glyphs this engine can substitute for a level definition's own symbol.
 *
 * A host that has to know which faces a document needs *before* laying it out
 * (the PDF bridge, choosing a fallback font) cannot see these in the source
 * model: Word authors bullets as Symbol/Wingdings private-use code points and
 * the layout rewrites them. Exported so that prediction stays derived from the
 * one function that performs the rewrite rather than a second copy of its table.
 */
export function predictBulletGlyphs(doc: DocxDocument): string[] {
  const out: string[] = [];
  for (const abstract of doc.abstractNumberings ?? []) {
    for (const level of abstract.levels ?? []) {
      if (level.format === "bullet") {
        out.push(normalizeBulletGlyph(level.text));
      }
    }
  }
  return out;
}

function normalizeBulletGlyph(text: string | undefined): string {
  if (!text || text.length === 0) {
    return "\u2022"; // round bullet
  }
  const cp = text.codePointAt(0)!;
  switch (cp) {
    // Symbol-font private-use code points Word emits for default bullets.
    case 0xf0b7: // Symbol "·" → round bullet
    case 0x00b7: // middle dot
      return "\u2022";
    case 0xf0a7: // Symbol filled small square
    case 0xf0a8:
      return "\u25aa";
    case 0xf0fc: // Wingdings check
      return "\u2713";
    default:
      // Already a renderable glyph (e.g. "o", "-", "•") — keep it.
      return text;
  }
}

function layoutParagraph(
  para: Paragraph,
  startY: number,
  contentWidth: number,
  options?: FullLayoutOptions,
  pageContext?: PageLayoutContext,
  imageMap?: ReadonlyMap<string, ImageDef>,
  listMarkers?: ReadonlyMap<Paragraph, ListMarker>,
  styleContext?: StyleResolveContext
): LayoutParagraph {
  const props = para.properties;
  // Resolve the effective paragraph *and* run properties from the style chain
  // (own props → Heading1 → Normal → docDefaults). Reading `para.properties`
  // directly would drop everything a named style contributes — which is where
  // heading spacing, quote indents and default body spacing live, so headings
  // ended up with zero space before/after and collided with their neighbours.
  //
  // `styleContext` carries the paragraph's position inside a table when there
  // is one, which is what lets a table style's conditional formats (header row,
  // banded rows, corner cells) reach the cell's text.
  const resolved = activeStyles?.paragraph(para, styleContext);
  const styleRunProps = resolved?.runProperties;
  // `resolveStyle` strips `style` from the merged result (it is the selector,
  // not an inherited value); put it back for consumers that read the selector.
  const effective: ParagraphProperties | undefined = resolved
    ? props?.style
      ? { ...resolved.paragraphProperties, style: props.style }
      : resolved.paragraphProperties
    : props;
  const spacing = effective?.spacing;
  // When the style supplies a concrete font size we honour it; only when it
  // does not do we fall back to the heuristic heading scale so headings stay
  // distinct in documents lacking a styles table.
  const headingScale = activeStyles
    ? resolveHeadingScale(activeStyles, para, styleRunProps?.size)
    : 1;

  // Space before. `w:contextualSpacing` drops it when the paragraph above uses
  // the same style, and the part the paragraph above's space after already
  // gives is not added again (see `computeContextualSpacing`).
  const contextual = activeContextualSpacing?.get(para);
  let spaceBefore = 0;
  if (!contextual?.before) {
    spaceBefore = spacingPt(spacing, "before") - (contextual?.sharedBefore ?? 0);
  }

  const indent = effective?.indent;
  // Prefer an explicitly threaded map; fall back to the active layout's
  // shared map so list markers also render inside tables, text boxes, SDTs,
  // footnotes, etc. (whose layoutParagraph calls don't thread it through).
  const marker = (listMarkers ?? activeListMarkers)?.get(para);
  // List paragraphs are indented by their numbering level; the marker text
  // is injected as a leading run below. An explicit paragraph indent (rare on
  // list items) still wins when larger.
  const markerIndentPt = marker ? marker.indentPt : 0;
  const leftIndentPt = Math.max(indent?.left ? twipsToPt(indent.left) : 0, markerIndentPt);
  const alignment = effective?.alignment ?? "left";

  // Collect runs
  const segments = collectParagraphSegments(para, styleRunProps);
  // Inject the list marker (bullet / number) as a leading text run so it
  // renders inline at the start of the first line. It takes the paragraph
  // mark's formatting, as Word sets it (pages2 P30f: an 11pt marker beside 8pt
  // text), not the first run's.
  // The level's own run properties (a bullet's face) apply over it; the face
  // counts above the baseline only (see `singleLineHeightPt`).
  let markerWidthPt = 0;
  const markerProps = marker
    ? mergeRunProperties(
        mergeRunProperties(styleRunProps, props?.markRunProperties),
        marker.runProperties
      )
    : undefined;
  if (marker) {
    segments.unshift({ text: marker.text, properties: markerProps, listMarker: true });
    markerWidthPt = measureRunText(
      marker.text,
      markerProps,
      getRunFontSizePt(markerProps) * headingScale,
      options,
      markerProps?.bold,
      markerProps?.italic
    );
  }

  /**
   * How far the first line starts *left* of the paragraph's left indent.
   *
   * `w:hanging` states it directly. A list gets it implicitly: the marker is
   * injected inline, so hanging it by exactly its own width makes the text that
   * follows land on the left indent — and every wrapped line, which starts at
   * the left indent, then aligns with the text rather than sitting under the
   * marker. Without this a wrapped list item ran back out to the marker's
   * column, which is the single most obvious way a rendered list looks wrong.
   */
  const hangingIndentPt = indent?.hanging ? twipsToPt(indent.hanging) : marker ? markerWidthPt : 0;
  // A hanging indent is a negative first-line indent; the two are mutually
  // exclusive in OOXML, and `w:hanging` wins where both appear.
  const firstLineIndentPt =
    hangingIndentPt > 0 ? -hangingIndentPt : indent?.firstLine ? twipsToPt(indent.firstLine) : 0;

  // Line height. The natural height comes from the paragraph's largest run,
  // so a 24 pt heading gets a 24 pt-sized line box instead of the document
  // default's — without this every paragraph shared one 14.4 pt line height
  // and large text overlapped the text around it.
  const naturalLineHeightPt =
    naturalParagraphLineHeightPt(segments, props, styleRunProps) * headingScale;
  /** A line's height under `w:spacing/@w:line` from its single-spaced height. */
  const pitch = effective?.snapToGrid === false ? undefined : activeGridPitch;
  const spacedLineHeightPt = (single: number, textSingle = single): number => {
    if (pitch !== undefined && spacing?.lineRule !== "exact") {
      return gridLineHeightPt(single, pitch, spacing);
    }
    if (!spacing?.line) {
      return single;
    }
    switch (spacing.lineRule ?? "auto") {
      case "exact":
        return twipsToPt(spacing.line);
      case "atLeast":
        return Math.max(twipsToPt(spacing.line), single);
      case "auto":
        // The multiple's extra comes from the text alone: a list marker that
        // raises the line adds its height once (pages2 P34: an 11pt Calibri
        // number beside 10pt Arial at 259/240 is 0.0792 × 11.5pt over single).
        return single + (spacing.line / 240 - 1) * textSingle;
      default:
        return single;
    }
  };
  const lineHeightPt = spacedLineHeightPt(naturalLineHeightPt);

  // The text column: the content width less both indents. Ignoring the right
  // indent let a block quote's or code block's right padding reduce nothing, so
  // its text ran to the same right edge as body text.
  const rightIndentForWrapPt = indent?.right ? twipsToPt(indent.right) : 0;
  const fullAvailableWidth = Math.max(1, contentWidth - leftIndentPt - rightIndentForWrapPt);

  // Where borders and shading sit relative to `rect`. Word paints them around
  // the *text*, so the space-before and space-after are inset away, and across
  // the indented text column rather than the full content width. A drawn top or
  // bottom rule takes its own room between the spacing and the text: its
  // `w:space` and its thickness (see `borderLeadPt`).
  let borders = resolveParagraphBorders(effective?.borders);
  const backgroundColor = resolveShadingFill(effective?.shading);
  const decorated = borders !== undefined || backgroundColor !== undefined;

  // Adjacent paragraphs sharing a decoration form one block — a two-paragraph
  // block quote is one bar, a fenced code block one frame. The shared edge is
  // not inset (so the bar and the fill run continuously through the space
  // between them) and its horizontal rule is dropped (so no line is drawn
  // across the middle of the block).
  if (borders && (contextual?.mergeDecorationTop || contextual?.mergeDecorationBottom)) {
    const { top, bottom, ...sides } = borders;
    // At an internal boundary the block's own top/bottom rule would cut across
    // the middle of it. `w:between` is the border OOXML defines *for* that
    // boundary; drawn as the lower paragraph's top edge it appears exactly once
    // per boundary. Absent a `w:between`, the boundary is left open.
    const between = resolveBorderEdge(effective?.borders?.between);
    borders = {
      ...sides,
      ...(contextual.mergeDecorationTop ? (between ? { top: between } : {}) : top ? { top } : {}),
      ...(contextual.mergeDecorationBottom ? {} : bottom ? { bottom } : {})
    };
  }

  const topLead = borderLeadPt(borders?.top, contextual?.mergeDecorationTop === true);
  const bottomLead = borderLeadPt(borders?.bottom, false);

  // When a page has wrap exclusions (square / tight / through floats)
  // we wrap line-by-line, asking the page context for the widest free
  // slot at the line's actual y-position. Otherwise we use the
  // legacy single-width path which never re-evaluates width across
  // lines — this preserves the existing layout output for documents
  // with no wrap (the overwhelming majority).
  // Word 2013 and later (`compatibilityMode` 15) let a justified line take a
  // word more by narrowing its spaces by up to a quarter; earlier modes do not
  // (pages2 P16–P18: 600 measured lines in Times New Roman and Calibri).
  const spaceShrink =
    alignment === "both" && (activeDoc?.settings?.compatibilityMode ?? 0) >= 15 ? 0.25 : 0;
  const tabStops = resolveTabStops(
    effective?.tabs,
    activeDoc?.settings?.defaultTabStop,
    indent?.hanging ? leftIndentPt : undefined
  );
  let lines: ParagraphSegment[][];
  let perLineSlots: LineSlot[] | undefined;
  let emptyDrop = 0;
  if (pageContext && pageContext.exclusions.length > 0) {
    const result = wrapSegmentsToLinesWithExclusions(
      segments,
      leftIndentPt,
      firstLineIndentPt,
      headingScale,
      lineHeightPt,
      startY + spaceBefore + topLead,
      pageContext,
      options,
      spaceShrink,
      tabStops
    );
    lines = result.lines;
    perLineSlots = result.slots;
    emptyDrop = result.emptyDrop ?? 0;
  } else {
    lines = wrapSegmentsToLines(
      segments,
      fullAvailableWidth,
      firstLineIndentPt,
      headingScale,
      options,
      spaceShrink,
      tabStops,
      leftIndentPt
    );
  }

  // Build line boxes
  const lineBoxes: LineBox[] = [];
  const lineNoteIds: number[][] = [];
  const lineBookmarks: string[][] = [];
  const pageBreakAfterLines: number[] = [];
  const columnBreakAfterLines: number[] = [];
  let yOffset = spaceBefore + topLead;

  for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
    const lineSegments = lines[lineIdx];
    const runs: LineBoxItem[] = [];
    // Resolve the line's effective slot. With exclusions each line has
    // its own usable [xOffset, width]; otherwise we keep the legacy
    // single-width behaviour and place the first line indent.
    yOffset += perLineSlots?.[lineIdx]?.drop ?? 0;
    const slot = perLineSlots?.[lineIdx] ?? {
      xOffset: 0,
      width: fullAvailableWidth
    };
    const lineLeftIndent = perLineSlots ? slot.xOffset : leftIndentPt;
    const lineAvailableWidth = perLineSlots ? slot.width : fullAvailableWidth;
    let xPos = lineIdx === 0 ? firstLineIndentPt : 0;

    // Calculate line width for alignment, and measure what the line holds so
    // its box can be sized around it: the ink extents of every face on the
    // line, plus the tallest inline image. An inline image sits *on* the
    // baseline, so it pushes the baseline down rather than only making the box
    // taller.
    let lineWidth = 0;
    let lineAscent = 0;
    let lineDescent = 0;
    let shiftedAscent = 0;
    let shiftedDescent = 0;
    let positioned = false;
    let maxImageHeight = 0;
    // An empty stretch (a bookmark's place) sizes the line only when nothing else does.
    // A picture sizes it too: a line of a picture and a bookmark is the
    // picture's height (pages2 P19).
    const lineHasText = lineSegments.some(
      seg => (isTextSegment(seg) && seg.text !== "") || ("type" in seg && seg.type === "image")
    );
    // A list marker beside text raises the line but never deepens it (pages2
    // P34: an 11pt Calibri number beside 10pt Arial, single spaced, is Calibri's
    // ascent over Arial's descent).
    const markerAboveOnly = lineSegments.some(
      seg => isTextSegment(seg) && seg.text !== "" && seg.listMarker !== true
    );
    for (const seg of lineSegments) {
      if ("type" in seg && seg.type === "break") {
        // A hard break draws nothing; it only ended the previous line.
        continue;
      }
      if ("type" in seg && seg.type === "tab") {
        lineWidth += seg.width ?? 0;
        continue;
      }
      if ("type" in seg && seg.type === "image") {
        const w = emuToPt(seg.content.width);
        const h = emuToPt(seg.content.height);
        lineWidth += w;
        if (h > maxImageHeight) {
          maxImageHeight = h;
        }
      } else if (seg.text !== "" || !lineHasText) {
        const fontSize = getRunFontSizePt(seg.properties) * headingScale;
        lineWidth += measureRunText(
          seg.text,
          seg.properties,
          fontSize,
          options,
          seg.properties?.bold,
          seg.properties?.italic
        );
        const metrics = measureRunMetrics(
          seg.text,
          seg.properties,
          fontSize,
          options,
          seg.properties?.bold,
          seg.properties?.italic
        );
        const scriptShift =
          seg.properties?.vertAlign === "superscript"
            ? fontSize * SCRIPT_BASELINE_SHIFT_FACTOR
            : seg.properties?.vertAlign === "subscript"
              ? -fontSize * SCRIPT_BASELINE_SHIFT_FACTOR
              : 0;
        lineAscent = Math.max(lineAscent, metrics.ascent + scriptShift);
        if (!(markerAboveOnly && isTextSegment(seg) && seg.listMarker === true)) {
          lineDescent = Math.min(lineDescent, metrics.descent + scriptShift);
        }
        const raise = (seg.properties?.position ?? 0) / 2;
        positioned ||= raise !== 0;
        shiftedAscent = Math.max(shiftedAscent, metrics.ascent + scriptShift + raise);
        shiftedDescent = Math.min(shiftedDescent, metrics.descent + scriptShift + raise);
      }
    }
    // `w:position` raises or lowers a run without scaling it. The line's top
    // and bottom move by as much as its runs then reach beyond the unmoved
    // ones, outside the line spacing: text raised beside unraised text makes
    // the line taller, a line of nothing but raised text only moves down
    // (pages2 P26: 11pt text raised 3, 6, 12pt makes its line 3.1, 5.9, 12.0pt
    // taller with 1.08 spacing, lowered 3.2, 6.2, 12.1pt; P27: a raised run's
    // last line, alone on it, keeps its height).
    const raisedAbove = positioned ? shiftedAscent - lineAscent : 0;
    const loweredBelow = positioned ? lineDescent - shiftedDescent : 0;

    // Apply alignment
    if (alignment === "center") {
      xPos = (lineAvailableWidth - lineWidth) / 2;
    } else if (alignment === "right" || alignment === "end") {
      xPos = lineAvailableWidth - lineWidth;
    }

    // `w:jc="both"` stretches every line but the last to the full column width.
    // It reached the layout model as `"justify"` and no renderer consumed it, so
    // justified text — the default for Chinese body copy — drew left-aligned.
    //
    // Latin widens the spaces between words; East Asian text has none to widen,
    // so the slack goes between characters instead. Which of the two applies is
    // decided per line by whether the line actually contains spaces, not by the
    // paragraph's language: a line of Chinese inside an otherwise English
    // paragraph still has to be spaced as Chinese.
    let charSpacing = 0;
    let wordSpacing = 0;
    // Only the text after a line's last tab is justified: what a tab placed stays.
    const lastTab = lineSegments.findLastIndex(seg => "type" in seg && seg.type === "tab");
    const justified = lineSegments.slice(lastTab + 1);
    if (alignment === "both" && !isJustifyExemptLine(lines, lineIdx)) {
      // The width this line actually has, not the paragraph's. A first-line
      // indent shortens it and a hanging indent lengthens it, and using the
      // paragraph width regardless pushed an indented first line 36pt past the
      // right margin while giving it five times the correct spacing.
      const slack = lineAvailableWidth - xPos - lineWidth;
      const { spaces: gapCount } = countJustifyOpportunities(justified);
      if (slack < 0 && spaceShrink > 0 && gapCount > 0) {
        // A line that took a word more by narrowing its spaces is drawn so.
        wordSpacing = slack / gapCount;
      } else if (slack > 0) {
        const { spaces, ideographs } = countJustifyOpportunities(justified);
        // Latin absorbs slack between words; East Asian text absorbs it between
        // characters. A line with both shares it in proportion to how many
        // opportunities each offers.
        //
        // Latin characters are deliberately *not* an opportunity: widening the
        // gaps inside a word is letter-spacing, which Western typesetting treats
        // as a defect outside of display use. Only ideographs are stretched, and
        // only spaces are widened — so a Latin-only line takes word spacing
        // alone, and a Latin word inside a Chinese line stays intact.
        const total = spaces + ideographs;
        if (total > 0) {
          const perOpportunity = slack / total;
          if (spaces > 0) {
            wordSpacing = perOpportunity;
          }
          if (ideographs > 0) {
            charSpacing = perOpportunity;
          }
        }
      }
    }

    xPos += lineLeftIndent;

    for (const [segIndex, seg] of lineSegments.entries()) {
      if ("type" in seg && seg.type === "break") {
        continue;
      }
      if ("type" in seg && seg.type === "tab") {
        const leader = seg.width ? tabLeaderRun(seg, xPos, headingScale, options) : undefined;
        if (leader) {
          runs.push(leader);
        }
        xPos += seg.width ?? 0;
        continue;
      }
      if ("type" in seg && seg.type === "image") {
        const widthPt = emuToPt(seg.content.width);
        const heightPt = emuToPt(seg.content.height);
        const img = seg.content.rId ? imageMap?.get(seg.content.rId) : undefined;
        if (!seg.content.rId) {
          // The room a VML shape takes; nothing of its own to draw here.
          xPos += widthPt;
          continue;
        }
        runs.push({
          type: "image",
          x: xPos,
          width: widthPt,
          height: heightPt,
          data: img?.data ?? new Uint8Array(0),
          mimeType: img?.contentType ?? mediaTypeToMime(img?.mediaType),
          altText: seg.content.altText
        });
        xPos += widthPt;
        continue;
      }
      const fontSize = getRunFontSizePt(seg.properties) * headingScale;
      const fonts = resolveRunFonts(seg.properties);
      // A `PositionedRun` carries a single typeface name and every renderer
      // requests a face by it, so a run naming one typeface for Latin and
      // another for East Asian text has to be emitted as one positioned run per
      // script stretch. Emitting it whole under the Latin name asked the font
      // manager for Chinese glyphs from a Latin face — and measured them there
      // too, so the widths disagreed with what was drawn.
      //
      // Uniform runs stay a single run, which is every Latin document and any
      // run whose typeface covers both scripts — unless justification is active,
      // because `Tc` applies to every glyph in a run and the only way to stretch
      // ideographs without also opening up the inside of a Latin word is to put
      // the two scripts in separate runs.
      const stretches: readonly { text: string; cjk: boolean }[] =
        fonts.ascii === fonts.eastAsia && charSpacing === 0
          ? [{ text: seg.text, cjk: false }]
          : splitByScript(seg.text);

      for (const stretch of stretches) {
        const stretchFont = stretch.cjk ? fonts.eastAsia : fonts.ascii;
        // Ideographs take the character spacing; a Latin stretch takes only the
        // word spacing on its spaces. Widening the gaps inside a Latin word is
        // letter-spacing, which Western typesetting treats as a defect.
        const stretchCharSpacing = stretch.cjk && segIndex > lastTab ? charSpacing : 0;
        const stretchWordSpacing = stretch.cjk || segIndex <= lastTab ? 0 : wordSpacing;
        const measured = measureLayoutText(
          stretch.text,
          stretchFont,
          fontSize,
          options,
          seg.properties?.bold,
          seg.properties?.italic
        );
        // `width` carries the justification slack, so anything measuring or
        // hit-testing the run needs no correction.
        const stretchWidth =
          stretchedRunWidth(measured, stretch.text, seg.properties) +
          justificationWidth(stretch.text, stretchCharSpacing, stretchWordSpacing);
        // `w:spacing` is drawn as character spacing, like justification's.
        const drawnCharSpacing = stretchCharSpacing + twipsToPt(seg.properties?.spacing ?? 0);

        runs.push({
          text: stretch.text,
          x: xPos,
          width: stretchWidth,
          font: stretchFont,
          fontSize,
          charSpacing: drawnCharSpacing !== 0 ? drawnCharSpacing : undefined,
          wordSpacing: stretchWordSpacing > 0 ? stretchWordSpacing : undefined,
          bold: seg.properties?.bold || undefined,
          italic: seg.properties?.italic || undefined,
          color: resolveColorHex(seg.properties?.color),
          underline: seg.properties?.underline !== undefined ? true : undefined,
          strikethrough: seg.properties?.strike || undefined,
          verticalAlign:
            seg.properties?.vertAlign === "superscript" || seg.properties?.vertAlign === "subscript"
              ? seg.properties.vertAlign
              : undefined
        });

        xPos += stretchWidth;
      }
    }

    const mappedAlignment =
      alignment === "both"
        ? "justify"
        : alignment === "end"
          ? "right"
          : alignment === "start"
            ? "left"
            : (alignment as "left" | "center" | "right" | "justify");

    // Each line is as tall as the faces set on it (a bullet in Symbol makes its
    // line taller than the next).
    const singlePt = singleLineHeightPt(lineSegments) * headingScale || naturalLineHeightPt;
    const nominalPt = spacedLineHeightPt(
      singlePt,
      singleLineHeightPt(lineSegments, false) * headingScale || singlePt
    );
    const lineMetrics = resolveWordLineMetrics({
      nominalHeight: nominalPt,
      ascent: lineAscent,
      descent: lineDescent,
      imageAscent: maxImageHeight,
      exact: spacing?.lineRule === "exact",
      ...lineTextBand(singlePt, nominalPt, pitch, spacing)
    });
    const lineBaseline = lineMetrics.baseline + raisedAbove;
    // A picture taller than the text sets the line, and `auto` spacing above
    // single still adds its extra for the text below it: Word for Mac 16 sets a
    // 40pt picture at 259/240 1.06pt (0.079 × 13.43pt of 11pt Calibri) taller,
    // and doubled 13.4pt taller (pages2 P19).
    const autoExtra =
      (spacing?.lineRule ?? "auto") === "auto" && pitch === undefined
        ? Math.max(0, nominalPt - singlePt)
        : 0;
    const lineBoxHeight =
      (autoExtra > 0 && maxImageHeight > 0
        ? Math.max(lineMetrics.height, lineMetrics.baseline - Math.min(0, lineDescent) + autoExtra)
        : lineMetrics.height) +
      raisedAbove +
      loweredBelow;

    const bottomSlack =
      spacing?.line && (spacing.lineRule ?? "auto") === "auto"
        ? Math.min(lineBoxHeight, Math.max(0, nominalPt - singlePt))
        : 0;
    lineBoxes.push({
      y: yOffset,
      height: lineBoxHeight,
      baseline: lineBaseline,
      ...(bottomSlack > 0 ? { bottomSlack } : {}),
      runs,
      alignment: mappedAlignment
    });
    // Which notes are referenced from this line, so a page split can hand each
    // note to the page holding its reference.
    lineNoteIds.push(lineSegments.flatMap(seg => (isTextSegment(seg) ? (seg.noteIds ?? []) : [])));
    lineBookmarks.push(
      lineSegments.flatMap(seg => (isTextSegment(seg) ? (seg.bookmarks ?? []) : []))
    );
    if (lineSegments.some(seg => "type" in seg && seg.type === "break" && seg.pageBreak === true)) {
      pageBreakAfterLines.push(lineIdx);
    }
    if (
      lineSegments.some(seg => "type" in seg && seg.type === "break" && seg.columnBreak === true)
    ) {
      columnBreakAfterLines.push(lineIdx);
    }

    // A line continued right of a float shares its row's top.
    if (perLineSlots?.[lineIdx]?.sameRow !== true) {
      yOffset += lineBoxHeight;
    }
  }

  // An empty paragraph still occupies a line — except a thematic break, whose
  // whole content *is* its border. Reserving a line for it made a Markdown
  // `---` nearly three times as tall as the `hr` it stands for.
  if (lineBoxes.length === 0 && effective?.thematicBreak !== true) {
    yOffset += emptyDrop + lineHeightPt;
  }

  // Space after — suppressed by `w:contextualSpacing` when the paragraph below
  // uses the same style.
  let spaceAfter = 0;
  if (contextual?.after) {
    spaceAfter = 0;
  } else if (spacing?.afterAutoSpacing) {
    spaceAfter = 5;
  } else if (spacing?.after != null) {
    spaceAfter = twipsToPt(spacing.after);
  }

  yOffset += bottomLead;
  const totalHeight = yOffset + spaceAfter;

  return {
    type: "paragraph",
    rect: { x: 0, y: startY, width: contentWidth, height: totalHeight },
    lines: lineBoxes,
    ...(borders ? { borders } : {}),
    ...(backgroundColor ? { backgroundColor } : {}),
    ...(lineNoteIds.some(ids => ids.length > 0) ? { lineNoteIds } : {}),
    ...(lineBookmarks.some(names => names.length > 0) ? { lineBookmarks } : {}),
    ...(pageBreakAfterLines.length > 0 ? { pageBreakAfterLines } : {}),
    ...(columnBreakAfterLines.length > 0 ? { columnBreakAfterLines } : {}),
    ...(decorated
      ? {
          decorationInsets: {
            left: leftIndentPt,
            right: rightIndentForWrapPt,
            top: contextual?.mergeDecorationTop && topLead === 0 ? 0 : spaceBefore + topLead,
            bottom: contextual?.mergeDecorationBottom ? 0 : spaceAfter + bottomLead
          }
        }
      : {}),
    sourceIndex: 0 // overwritten by caller
  };
}

/**
 * Resolve `w:pBdr` to drawable edges, or undefined when nothing is drawn.
 *
 * `w:sz` is in eighths of a point and `w:space` in whole points. A border
 * declared without a size is a hairline, matching how Word renders one.
 */
/** `table` with its effective properties (`StyleResolver.tableProperties`). */
function withEffectiveProperties(table: Table): Table {
  const properties = activeStyles?.tableProperties(table) ?? table.properties;
  return properties === table.properties ? table : { ...table, properties };
}

/**
 * One `w:pBdr` edge as a drawable stroke, or undefined when it draws nothing.
 *
 * `w:sz` is in eighths of a point and `w:space` in whole points. An edge declared
 * without a size is a hairline, which is how Word renders one.
 */
function resolveBorderEdge(b: Border | undefined): LayoutBorderEdge | undefined {
  if (!b || b.style === "none" || b.style === "nil") {
    return undefined;
  }
  const width = Math.max(0.25, b.size != null ? b.size / 8 : 0.5);
  const color = !b.color || b.color === "auto" ? "000000" : b.color.replace(/^#/, "");
  const extent = ruleThicknessPt(b.style, width);
  return { width, color, space: b.space ?? 0, ...(extent !== width ? { extent } : {}) };
}

/**
 * How thick Word draws a rule of `style` whose `w:sz` is `width` points, lines
 * and gaps together. Measured in Word 16.114's PDF output at 0.5, 1.5 and 3pt
 * (pages2 P6); a style not listed is taken as one line.
 */
function ruleThicknessPt(style: string, width: number): number {
  switch (style) {
    case "double":
      return 3 * width;
    case "triple":
      return 5 * width;
    case "thinThickSmallGap":
    case "thickThinSmallGap":
      return width + 1.5;
    case "thinThickMediumGap":
    case "thickThinMediumGap":
      return 2 * width;
    case "thinThickLargeGap":
    case "thickThinLargeGap":
      return width + 2.25;
    case "thinThickThinSmallGap":
      return width + 3;
    case "thinThickThinLargeGap":
      return 2 * width + 3;
    case "wave":
      return width >= 1.5 && width < 3 ? 4 : 3;
    case "doubleWave":
      return 5.25;
    case "dashDotStroked":
      return 3;
    default:
      return width;
  }
}

/**
 * The room a paragraph's top or bottom rule takes between its spacing and its
 * text: the rule's `w:space` and its thickness (`ruleThicknessPt`). A `w:between` rule inside a run of paragraphs sharing
 * borders keeps its space on both sides. Measured in Word 16.114's PDF output
 * (pages2 P5).
 */
function borderLeadPt(edge: LayoutBorderEdge | undefined, between: boolean): number {
  if (!edge) {
    return 0;
  }
  return edge.space * (between ? 2 : 1) + (edge.extent ?? edge.width);
}

function resolveParagraphBorders(borders: ParagraphBorders | undefined): LayoutBorders | undefined {
  if (!borders) {
    return undefined;
  }
  const top = resolveBorderEdge(borders.top);
  const bottom = resolveBorderEdge(borders.bottom);
  // `w:bar` is a vertical bar down the paragraph's leading edge — the same
  // stroke as `w:left` for a left-to-right layout.
  const left = resolveBorderEdge(borders.left ?? borders.bar);
  const right = resolveBorderEdge(borders.right);
  if (!top && !bottom && !left && !right) {
    return undefined;
  }
  return {
    ...(top ? { top } : {}),
    ...(bottom ? { bottom } : {}),
    ...(left ? { left } : {}),
    ...(right ? { right } : {})
  };
}

// =============================================================================
// Internal: Table Layout
// =============================================================================

/** A table; Word sets the lines in its cells off the document grid. */
function layoutTable(...args: Parameters<typeof layoutTableOffGrid>): LayoutTable {
  return withGridPitch(undefined, () => layoutTableOffGrid(...args));
}

function layoutTableOffGrid(
  source: Table,
  startY: number,
  contentWidth: number,
  sourceIndex: number,
  options?: FullLayoutOptions,
  imageMap?: ReadonlyMap<string, ImageDef>,
  nestedInFixed?: boolean
): LayoutTable {
  // The table as Word sees it: its table style's properties beneath its own
  // (the default table style's, for a table naming none).
  const table = withEffectiveProperties(source);
  const numCols = gridColumnCount(table);

  // Resolve per-column widths (in points). Prefer the table's explicit
  // `columnWidths` (twips) — populated e.g. by the Excel→Word bridge —
  // scaled to fit the available content width so a table authored wider
  // than the page still renders proportionally. Fall back to equal
  // division when no column widths are declared. This mirrors the
  // sister layout engine in `layout.ts` (which also honours
  // `columnWidths` + `gridSpan`).
  // A nested table (`nestedInFixed` set) keeps its own width rather than the
  // cell's. Word fits it to its content only inside an auto-fit table (final7
  // nest R3c, an auto-fit nested table with tblW 3000: 61pt and 89pt beside
  // wrapping text; R3, the same in a fixed-layout table: its grid, 75pt each).
  // A table on the page is placed the same way: a fixed-layout table keeps its
  // grid rather than being stretched to the margins (pages2 R6a: a 360pt grid
  // in a 468pt column, its first column 120pt wide).
  const nestedWidth = nestedTableWidthTwips(
    table.properties,
    table.columnWidths,
    ptToTwips(contentWidth),
    nestedInFixed ?? false
  );
  const colWidths =
    (nestedInFixed === true
      ? undefined
      : autoFitColumnWidthsPt(source, table, numCols, contentWidth, options)) ??
    resolveColumnWidthsPt(
      table,
      numCols,
      nestedWidth === undefined ? contentWidth : twipsToPt(nestedWidth)
    );
  // Prefix sums so a cell at column `ci` starts at `colOffsets[ci]` and
  // a `gridSpan` cell can sum the widths it covers.
  const colOffsets: number[] = [0];
  for (let i = 0; i < colWidths.length; i++) {
    colOffsets.push(colOffsets[i] + colWidths[i]);
  }

  const cells: LayoutTableCell[] = [];
  let cursorY = 0;
  const spacing = cellSpacingPt(table.properties);

  // Table-style context, shared by every cell in this table. A table style's
  // conditional formats (`firstRow`, `lastRow`, banding, corner cells) are
  // resolved per cell position, so each cell's paragraphs need to know where
  // they sit. Without this, Word's built-in table styles lost their bold header
  // row and banded shading in the output.
  const totalRows = table.rows.length;
  // How many rows each vertically merged cell covers, keyed `row:cell`.
  const rowSpans = verticalMergeSpans(table);
  /**
   * Cells covering several rows, waiting for the last of them: they take no
   * part in their own rows' heights, and the last row grows if they need more.
   */
  const spanning: {
    cell: LayoutTableCell;
    endRow: number;
    contentBottom: number;
    verticalAlign: VerticalCellAlign | undefined;
    /** The continuing cells below it. */
    continuing: LayoutTableCell[];
    gridCol: number;
  }[] = [];

  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    // Shared with the pagination pass, which had a different answer.
    let maxRowHeight = minimumRowHeightPt(table.properties, DEFAULT_FONT_SIZE_PT);
    const tableBorders = table.properties?.borders;
    const isLastRow = ri === table.rows.length - 1;
    const { above: ruleAbove, below: ruleBelow } = rowRulesPt(table, ri);
    // This row's cells with the `w:vAlign` each asked for. A cell is laid out
    // from its top margin because its final height is not known until every
    // cell in the row has been measured; the alignment is applied once the row
    // height settles, below.
    const rowCells: Array<{
      cell: LayoutTableCell;
      verticalAlign: VerticalCellAlign | undefined;
    }> = [];
    // `w:trHeight`: `exact` fixes the row height outright, `atLeast` raises the
    // floor. Ignoring it let a table with declared row heights render squashed —
    // and, now that this pass owns pagination, shifted everything after it.
    const declaredHeight = row.properties?.height;
    const declaredHeightPt =
      declaredHeight?.value != null ? twipsToPt(declaredHeight.value) : undefined;
    // An `atLeast` height is the room between the rules: the rule above, and the
    // table's bottom rule under its last row, come on top (pages2 P10: ten 45pt
    // rows with 0.5pt rules take 455.5pt, with 3pt rules 483pt).
    if (
      declaredHeightPt !== undefined &&
      declaredHeight?.rule !== "auto" &&
      declaredHeight?.rule !== "exact"
    ) {
      maxRowHeight = Math.max(maxRowHeight, declaredHeightPt + ruleAbove + ruleBelow);
    }

    // Track the grid column each cell occupies, honouring gridSpan so a
    // 2-wide cell pushes the next cell two grid columns to the right.
    let gridCol = row.properties?.gridBefore ?? 0;
    const { top: rowMarginTop, bottom: rowMarginBottom } = rowCellMarginsPt(table, row);
    let hasUnrotatedCell = false;
    let rotatedRowHeight = 0;
    for (let ci = 0; ci < row.cells.length; ci++) {
      const cell = row.cells[ci];
      const span = Math.max(1, cell.properties?.gridSpan ?? 1);
      const startCol = Math.min(gridCol, colWidths.length - 1);
      const endCol = Math.min(gridCol + span, colWidths.length);
      const cellX = colOffsets[startCol] ?? 0;
      const cellWidth = (colOffsets[endCol] ?? contentWidth) - cellX;
      const cellContent: (LayoutParagraph | LayoutTable)[] = [];
      // Word's cell margins (`w:tblCellMar` / `w:tcMar`) inset the content from
      // all four borders. Both the horizontal inset and the reduced wrap width
      // matter: subtracting the margins from the available width alone (as this
      // did) left every cell's text flush against its left border.
      const cellMargins = {
        ...resolveCellMarginsPt(table.properties, cell.properties),
        top: rowMarginTop,
        bottom: rowMarginBottom
      };
      const inset = cellSpacingInsetPt(table, startCol === 0, endCol >= colWidths.length);
      const margins = {
        ...cellMargins,
        left: cellMargins.left + inset.left,
        right: cellMargins.right + inset.right
      };
      const innerWidth = Math.max(1, cellWidth - margins.left - margins.right);
      let cellCursorY = margins.top + ruleAbove;
      const cellStyleContext: StyleResolveContext | undefined = activeStyles?.cellContext(
        source,
        ri,
        startCol,
        Math.max(numCols, colWidths.length)
      );

      cell.content.forEach((block, k) => {
        if (block.type === "paragraph") {
          const laid = layoutParagraph(
            block,
            cellCursorY,
            innerWidth,
            options,
            undefined,
            imageMap,
            undefined,
            cellStyleContext
          );
          // Shift the content right by the left margin. Renderers position a
          // cell's content at `cell.rect.x + inner.rect.x`, so the inset has to
          // live in the block's own origin.
          const height = isCellEndAfterNestedTable(cell.content, k) ? 0 : laid.rect.height;
          cellContent.push({
            ...laid,
            rect: { ...laid.rect, x: laid.rect.x + margins.left, height },
            sourceIndex: -1
          });
          cellCursorY = laid.rect.y + height;
        } else if (block.type === "table") {
          // Nested table: lay it out within the cell's content width and
          // stack it below preceding content. The PDF/SVG renderers
          // already translate nested `LayoutTable` rects by the cell
          // origin, so emitting it here is all that's needed.
          const laidNested = layoutTable(
            block,
            cellCursorY,
            innerWidth,
            -1,
            options,
            imageMap,
            table.properties?.layout === "fixed"
          );
          cellContent.push({
            ...laidNested,
            rect: { ...laidNested.rect, x: laidNested.rect.x + margins.left }
          });
          cellCursorY = laidNested.rect.y + laidNested.rect.height;
        }
      });

      const cellHeight = cellCursorY + margins.bottom + ruleBelow;
      const rowSpan = rowSpans.get(`${ri}:${ci}`) ?? 1;
      // A merged cell's continuation holds nothing of its own; the cell that
      // starts the merge sizes the rows it covers once they are all laid out.
      const merged = rowSpan > 1 || cell.properties?.verticalMerge === "continue";
      // Rotated text runs along the row's height and does not grow it: the row
      // takes its other cells' height, whatever the rotated text's length
      // (pages2 P20/P21: one word to 45 words, all 22.95pt rows). A row of
      // nothing but rotated cells is one line of the first paragraph high.
      if (isRotatedTextDirection(cell.properties?.textDirection)) {
        const first = cell.content[0];
        if (!merged && first?.type === "paragraph") {
          const oneLine = layoutParagraph(
            first,
            0,
            ROTATED_LINE_WIDTH_PT,
            options,
            undefined,
            imageMap,
            undefined,
            cellStyleContext
          ).rect.height;
          rotatedRowHeight = Math.max(
            rotatedRowHeight,
            margins.top + ruleAbove + oneLine + margins.bottom + ruleBelow
          );
        }
      } else if (!merged) {
        maxRowHeight = Math.max(maxRowHeight, cellHeight);
        hasUnrotatedCell = true;
      }

      // Background shading: cell `w:shd` → table `w:shd` → the table style's
      // conditional formats (header band, row stripes, corner cells). Resolved
      // through the shared helper so this and the SVG renderer paint the same
      // colour — the field exists on `LayoutTableCell` and both renderers draw
      // it, but nothing ever filled it in, so every table came out white.
      const backgroundColor = activeStyles
        ? activeStyles.tableCellFill(source, cell, {
            rowIndex: ri,
            colIndex: startCol,
            totalRows,
            totalCols: Math.max(numCols, colWidths.length)
          })
        : undefined;

      cells.push({
        // Table-relative, like `x`. Renderers add the table's origin to both;
        // emitting an absolute `y` here made them add it twice.
        rect: { x: cellX, y: cursorY, width: cellWidth, height: cellHeight },
        row: ri,
        col: ci,
        content: cellContent,
        ...(backgroundColor ? { backgroundColor } : {}),
        borders: mergedCellBorders(
          resolveCellBorders(
            table.properties?.borders,
            cell.properties?.borders,
            ri === 0,
            ri + rowSpan - 1 === table.rows.length - 1,
            startCol === 0,
            endCol >= colWidths.length
          ),
          cell.properties?.verticalMerge === "continue",
          rowSpans.has(`${ri + 1}:continue@${startCol}`)
        ),
        ...(rowSpan > 1 ? { rowSpan } : {})
      });
      if (rowSpan > 1) {
        spanning.push({
          cell: cells[cells.length - 1]!,
          endRow: ri + rowSpan - 1,
          contentBottom: cellCursorY + margins.bottom,
          verticalAlign: cell.properties?.verticalAlign,
          continuing: [],
          gridCol: startCol
        });
      } else {
        if (cell.properties?.verticalMerge === "continue") {
          spanning
            .find(m => m.gridCol === startCol && m.endRow >= ri)
            ?.continuing.push(cells[cells.length - 1]!);
        }
        rowCells.push({
          cell: cells[cells.length - 1]!,
          verticalAlign: merged ? undefined : cell.properties?.verticalAlign
        });
      }

      gridCol += span;
    }

    if (!hasUnrotatedCell) {
      maxRowHeight = Math.max(maxRowHeight, rotatedRowHeight);
    }

    // A merged cell ending here gets the room it needs from this, its last row
    // (pages2 P9: three rows under ten merged lines keep their own heights but
    // the last, which grows to hold the rest).
    for (const merged of spanning) {
      if (merged.endRow === ri) {
        const above = cursorY - merged.cell.rect.y;
        maxRowHeight = Math.max(maxRowHeight, merged.contentBottom + ruleBelow - above);
      }
    }

    // An `exact` rule clamps the row even when its content is taller; Word
    // clips the overflow rather than growing the row.
    // The height includes the rule above; the table's bottom rule under the
    // last row does not (pages2 P10: ten exact 30pt rows take 300.5pt with
    // 0.5pt rules, 303pt with 3pt rules).
    if (declaredHeightPt !== undefined && declaredHeight?.rule === "exact") {
      // With cell spacing an exact height holds the gap above the row but s:
      // the first row takes its height, the others s more (pages2 P23g: three
      // exact 30pt rows with 3pt spacing take 103.9pt, with 10pt 132.1pt; the
      // first row's cells are 2s shorter than the height, the others' s).
      maxRowHeight =
        spacing > 0
          ? declaredHeightPt +
            (ri === 0 ? ruleWidthPt(tableBorders?.top) : spacing) +
            ruleWidthPt(tableBorders?.insideH) +
            (isLastRow ? 2 * spacing + ruleWidthPt(tableBorders?.bottom) : 0)
          : declaredHeightPt + ruleBelow;
    }

    // Normalize cell heights to row max, then honour `w:vAlign`: a cell was
    // measured from its top margin, so anything but `top` has to move its
    // content down into the slack the row's tallest cell created.
    for (const { cell: c, verticalAlign } of rowCells) {
      // Only a cell whose content is genuinely taller than the settled row needs
      // bounding, which happens when an `exact` rule shortened the row. Marking
      // every exact cell would make renderers emit a clip that changes nothing.
      if (c.rect.height > maxRowHeight) {
        (c as { clipToBounds?: boolean }).clipToBounds = true;
      }
      shiftCellContentForVerticalAlign(c, verticalAlign, maxRowHeight);
      (c as { rect: { height: number } }).rect.height = maxRowHeight;
    }

    cursorY += maxRowHeight;
    for (const merged of spanning) {
      if (merged.endRow === ri) {
        const c = merged.cell;
        const height = cursorY - c.rect.y;
        (c as { rect: { height: number } }).rect.height = merged.contentBottom;
        shiftCellContentForVerticalAlign(c, merged.verticalAlign, height);
        (c as { rect: { height: number } }).rect.height = height;
        // What a continuing cell holds is part of the merged cell: Word puts it
        // where the merged cell's text begins, on that page (pages2 R2c). It
        // goes first, taking no room, so a page break never separates it from
        // the start; `tableSplitInfo` lists it the same way.
        const first = c.content[0]?.rect.y ?? 0;
        // Word draws none of it (pages2 R4j, R7i), so only its positions stay.
        const moved = merged.continuing.flatMap(cont => {
          const blocks = cont.content.map(block => ({
            ...hideContent(block),
            rect: { ...block.rect, y: first, height: 0 }
          }));
          (cont as { content: readonly LayoutBlock[] }).content = [];
          return blocks;
        });
        (c as { content: readonly LayoutBlock[] }).content = [...moved, ...c.content];
      }
    }
  }

  return {
    type: "table",
    rect: { x: 0, y: startY, width: contentWidth, height: cursorY },
    cells,
    sourceIndex
  };
}

/** A block with nothing drawn: lines keep their places and bookmarks, not their runs. */
function hideContent(block: LayoutBlock): LayoutBlock {
  if (block.type === "paragraph") {
    const { borders: _b, backgroundColor: _c, ...rest } = block;
    return { ...rest, lines: block.lines.map(line => ({ ...line, runs: [] })) };
  }
  return {
    ...block,
    cells: block.cells.map(({ borders: _b, backgroundColor: _c, ...cell }) => ({
      ...cell,
      content: cell.content.map(hideContent)
    }))
  };
}

/**
 * Move a cell's content down to satisfy `w:vAlign` once the row's height is
 * known.
 *
 * A cell is laid out top-down from its own top margin, because its final height
 * depends on the tallest cell in the row and that is not known until the whole
 * row has been measured. `center` and `bottom` therefore mean "translate the
 * finished content into the slack that appeared". Only the top-level blocks move:
 * a paragraph's lines are positioned relative to the paragraph, and a nested
 * table's cells relative to that table, so both follow their own origin.
 *
 * The cell's height at this point is still its natural content height — the
 * caller normalises it to the row height straight after.
 */
/**
 * How many rows each cell that starts a vertical merge covers, keyed
 * `row:cellIndex`: itself and the rows below whose cell at the same grid column
 * continues it. Each continuing cell is also recorded as `row:continue@column`.
 */
function verticalMergeSpans(table: Table): Map<string, number> {
  const starts = table.rows.map(row => {
    const out = new Map<number, { index: number; merge: string | undefined }>();
    let col = row.properties?.gridBefore ?? 0;
    row.cells.forEach((cell, index) => {
      out.set(col, { index, merge: cell.properties?.verticalMerge });
      col += Math.max(1, cell.properties?.gridSpan ?? 1);
    });
    return out;
  });
  const spans = new Map<string, number>();
  starts.forEach((cols, r) => {
    for (const [col, { index, merge }] of cols) {
      if (merge === "continue") {
        spans.set(`${r}:continue@${col}`, 1);
      }
      if (merge !== "restart") {
        continue;
      }
      let end = r;
      while (starts[end + 1]?.get(col)?.merge === "continue") {
        end++;
      }
      if (end > r) {
        spans.set(`${r}:${index}`, end - r + 1);
      }
    }
  });
  return spans;
}

/**
 * A merged cell's borders: no rule across the merge — a continuing cell draws
 * no top, and none at its bottom while the merge continues below it.
 */
function mergedCellBorders(
  borders: LayoutTableCell["borders"],
  continuing: boolean,
  continuedBelow: boolean
): LayoutTableCell["borders"] {
  if (!borders || (!continuing && !continuedBelow)) {
    return borders;
  }
  const { top, bottom, ...sides } = borders;
  return {
    ...sides,
    ...(top && !continuing ? { top } : {}),
    ...(bottom && !continuedBelow ? { bottom } : {})
  };
}

function shiftCellContentForVerticalAlign(
  cell: LayoutTableCell,
  verticalAlign: VerticalCellAlign | undefined,
  rowHeight: number
): void {
  if (!verticalAlign || verticalAlign === "top") {
    return;
  }
  // A `w:trHeight` of `exact` can leave the row shorter than its content, in
  // which case there is nothing to distribute and Word clips instead.
  const slack = rowHeight - cell.rect.height;
  if (slack <= 0) {
    return;
  }
  const shift = verticalAlign === "center" ? slack / 2 : slack;
  for (const block of cell.content) {
    (block as { rect: { y: number } }).rect.y += shift;
  }
}

/**
 * Resolve the four visible borders of a table cell into layout-model form
 * (`{ width: pt, color: hex }`). A cell's own border wins; otherwise the
 * table-level border applies — outer edges use `top/left/bottom/right`, inner
 * edges use `insideH/insideV`. OOXML border `size` is in eighths of a point.
 */
/**
 * The room a row's rules (and, with cell spacing, gaps) take: `above` its
 * cells, `below` them, `cut` below a piece a page break cuts off, and
 * `carried` above a row that opens a continuation page, and `closing` below
 * the last row on a page when the table goes on.
 *
 * Without spacing Word gives each horizontal rule its own width in the row
 * below it, and the bottom rule to the last row (Word for Mac 16: rows of one
 * 11pt line with 0.5pt rules are 22.95pt apart, not 22.49).
 *
 * With cell spacing (`w:tblCellSpacing`, s) the rules no longer collapse: the
 * table keeps its outer rules, every cell its own top and bottom rule, and 2s
 * separates the table's edge from the cells and the rows from each other. Each
 * gap goes to the row below it, the last one to the last row (pages2 P23: two
 * rows of 22.4pt with 3pt spacing and 0.5pt rules take 65.8pt). A row opening
 * a continuation page has the table's top rule and s above its cells (P24: 4.2pt
 * above the text with 3pt spacing, 10.9pt with 10pt).
 */
function rowRulesPt(
  table: Table,
  ri: number
): { above: number; below: number; cut: number; carried: number; closing: number } {
  const borders = table.properties?.borders;
  const spacing = cellSpacingPt(table.properties);
  const last = ri === table.rows.length - 1;
  const top = ruleWidthPt(borders?.top);
  const inside = ruleWidthPt(borders?.insideH);
  const bottom = ruleWidthPt(borders?.bottom);
  if (spacing > 0) {
    return {
      above: (ri === 0 ? top : 0) + 2 * spacing + inside,
      below: inside + (last ? 2 * spacing + bottom : 0),
      cut: inside + spacing + bottom,
      carried: top + spacing + inside,
      closing: spacing + bottom
    };
  }
  const above = ri === 0 ? top : inside;
  return {
    above,
    below: last ? bottom : 0,
    cut: last ? bottom : inside,
    carried: above,
    closing: 0
  };
}

/**
 * How far cell spacing moves a cell's content in from its grid columns: s on
 * each side, 2s at the table's edges, and the cell's own side rules, which no
 * longer collapse into the grid lines (pages2 P23: text 2s further right in the
 * first column, s in the others; P25: 0.5pt rules narrow the text by 1pt, 1.5pt
 * rules by 3pt). Nothing without spacing.
 */
function cellSpacingInsetPt(
  table: Table,
  firstCol: boolean,
  lastCol: boolean
): { left: number; right: number } {
  const spacing = cellSpacingPt(table.properties);
  if (spacing <= 0) {
    return { left: 0, right: 0 };
  }
  const borders = table.properties?.borders;
  return {
    left: spacing * (firstCol ? 2 : 1) + ruleWidthPt(firstCol ? borders?.left : borders?.insideV),
    right: spacing * (lastCol ? 2 : 1) + ruleWidthPt(lastCol ? borders?.right : borders?.insideV)
  };
}

/** A table's `w:tblCellSpacing` in points, or 0 (only a twips width is meaningful). */
function cellSpacingPt(props: TableProperties | undefined): number {
  const spacing = props?.cellSpacing;
  return spacing?.type === "dxa" && spacing.value > 0 ? twipsToPt(spacing.value) : 0;
}

/** The width a table rule takes, in points, or 0 for none. */
function ruleWidthPt(border: Border | undefined): number {
  if (!border || border.style === "none" || border.style === "nil") {
    return 0;
  }
  return border.size != null ? border.size / 8 : 0.5;
}

function resolveCellBorders(
  tableBorders: TableBorders | undefined,
  cellBorders: TableBorders | undefined,
  isTopRow: boolean,
  isBottomRow: boolean,
  isLeftCol: boolean,
  isRightCol: boolean
): LayoutTableCell["borders"] {
  const edge = (
    cellEdge: Border | undefined,
    outerEdge: Border | undefined,
    innerEdge: Border | undefined,
    isOuter: boolean
  ): { width: number; color: string } | undefined => {
    const b = cellEdge ?? (isOuter ? outerEdge : innerEdge);
    if (!b || b.style === "none" || b.style === "nil") {
      return undefined;
    }
    // `size` is in eighths of a point; default to a hairline (0.5pt) when
    // a border is declared without an explicit size.
    const width = b.size != null ? b.size / 8 : 0.5;
    const color = !b.color || b.color === "auto" ? "000000" : b.color;
    return { width: Math.max(0.25, width), color };
  };

  // `border-collapse`: every interior rule is shared by two cells and must be
  // drawn once. Each cell owns its bottom and right edges (plus the table's
  // outer top and left); the neighbour's coincident top/left is dropped unless
  // that cell declares one itself. Emitting both stroked every interior rule
  // twice, and the second stroke overpainted the first — a table style's darker
  // header rule came out in the lighter body-rule colour.
  const top = isTopRow
    ? edge(cellBorders?.top, tableBorders?.top, tableBorders?.insideH, true)
    : edge(cellBorders?.top, undefined, undefined, false);
  const bottom = edge(
    cellBorders?.bottom,
    tableBorders?.bottom,
    tableBorders?.insideH,
    isBottomRow
  );
  const left = isLeftCol
    ? edge(cellBorders?.left, tableBorders?.left, tableBorders?.insideV, true)
    : edge(cellBorders?.left, undefined, undefined, false);
  const right = edge(cellBorders?.right, tableBorders?.right, tableBorders?.insideV, isRightCol);

  if (!top && !bottom && !left && !right) {
    return undefined;
  }
  return {
    ...(top ? { top } : {}),
    ...(bottom ? { bottom } : {}),
    ...(left ? { left } : {}),
    ...(right ? { right } : {})
  };
}

/**
 * A cell's effective inner margins, converted to points.
 *
 * The cascade itself (cell `w:tcMar` → table `w:tblCellMar` → Word's defaults)
 * lives in `layout-constants` so the paginator derives exactly the same numbers;
 * a hardcoded guess in each was how the two passes came to disagree about how
 * tall every table row is.
 */
/**
 * A row's vertical cell margins: its text starts below the largest top margin
 * of its cells and ends above the largest bottom margin, in every cell (pages2
 * R7i: cells of one row with 0 and 12pt top margins set their first lines on
 * one baseline).
 */
function rowCellMarginsPt(
  table: Table,
  row: Table["rows"][number]
): { top: number; bottom: number } {
  let top = 0;
  let bottom = 0;
  for (const cell of row.cells) {
    const own = resolveCellMarginsPt(table.properties, cell.properties);
    top = Math.max(top, own.top);
    bottom = Math.max(bottom, own.bottom);
  }
  return { top, bottom };
}

function resolveCellMarginsPt(
  tableProps: TableProperties | undefined,
  cellProps: TableCellProperties | undefined
): { top: number; right: number; bottom: number; left: number } {
  const twips = resolveCellMarginsTwips(tableProps, cellProps);
  return {
    top: twipsToPt(twips.top),
    right: twipsToPt(twips.right),
    bottom: twipsToPt(twips.bottom),
    left: twipsToPt(twips.left)
  };
}

/**
 * Column widths of an auto-fit table whose cells ask for no width, from what
 * they hold, or `undefined` for any other table.
 *
 * Word sizes such a table by its content, not by its grid: a three-column
 * table whose middle column holds sentences gives that column most of the
 * page (Word for Mac 16: 30, 401 and 37pt of 468), where an even split wraps
 * every sentence and doubles the table's height. This follows the automatic
 * table layout of CSS 2.1, §17.5.2.2: each column takes its widest unbroken
 * content; leftover width goes to the columns in proportion to that, and
 * a shortfall is taken from what they can give up above their longest word.
 */
function autoFitColumnWidthsPt(
  source: Table,
  table: Table,
  numCols: number,
  contentWidth: number,
  options: FullLayoutOptions | undefined
): number[] | undefined {
  const widths = autoFitContentWidthsPt(source, table, numCols, options);
  if (!widths) {
    return undefined;
  }
  const { minW, maxW } = widths;
  const tableWidth = table.properties?.width;
  const target =
    tableWidth?.type === "pct"
      ? (contentWidth * tableWidth.value) / 5000
      : tableWidth?.type === "dxa" && tableWidth.value > 0
        ? twipsToPt(tableWidth.value)
        : Math.min(
            contentWidth,
            maxW.reduce((a, b) => a + b, 0)
          );
  const sumMax = maxW.reduce((a, b) => a + b, 0);
  const sumMin = minW.reduce((a, b) => a + b, 0);
  if (sumMax <= 0) {
    return undefined;
  }
  if (sumMax <= target) {
    return maxW.map(w => w + ((target - sumMax) * w) / sumMax);
  }
  if (sumMin < target) {
    const give = sumMax - sumMin;
    return minW.map((w, i) => w + ((target - sumMin) * (maxW[i]! - w)) / give);
  }
  return minW.map(w => (w * target) / sumMin);
}

/**
 * The narrowest and widest a table can usefully be, as a cell holding it
 * measures it: by its own width ({@link nestedTableWidthTwips}) where it has
 * one, otherwise an auto-fit table by its columns' content. `undefined` when it
 * cannot be measured.
 */
function nestedTableContentWidthsPt(
  source: Table,
  options: FullLayoutOptions | undefined
): { min: number; max: number } | undefined {
  const table = withEffectiveProperties(source);
  const numCols = gridColumnCount(table);
  // Its own width, where it has one; a percentage of a cell not yet sized is not.
  const own = nestedTableWidthTwips(table.properties, table.columnWidths, 0, false);
  if (own !== undefined && own > 0) {
    return { min: twipsToPt(own), max: twipsToPt(own) };
  }
  if (table.properties?.layout === "fixed") {
    return undefined;
  }
  const widths = autoFitContentWidthsPt(source, table, numCols, options);
  if (!widths) {
    return undefined;
  }
  return {
    min: widths.minW.reduce((a, b) => a + b, 0),
    max: widths.maxW.reduce((a, b) => a + b, 0)
  };
}

/**
 * Per column of an auto-fit table whose cells ask for no width, its longest
 * unbreakable content (`minW`) and its content set on one line (`maxW`), cell
 * margins included. A nested table counts as content of the cell holding it:
 * Word sizes the outer column around it (final6 Qaa–Qac, where the outer
 * table's first column shrinks to 187pt of 468 and the nested one wraps).
 */
function autoFitContentWidthsPt(
  source: Table,
  table: Table,
  numCols: number,
  options: FullLayoutOptions | undefined
): { minW: number[]; maxW: number[] } | undefined {
  if (numCols <= 0 || table.properties?.layout === "fixed") {
    return undefined;
  }
  const minW = new Array<number>(numCols).fill(0);
  const maxW = new Array<number>(numCols).fill(0);
  // Cells spanning several columns, applied once the single columns are known.
  const spanning: { start: number; span: number; min: number; max: number }[] = [];
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r]!;
    if (row.properties?.gridBefore || row.properties?.gridAfter) {
      return undefined;
    }
    let gridCol = 0;
    for (let c = 0; c < row.cells.length; c++) {
      const cell = row.cells[c]!;
      const width = cell.properties?.width;
      const span = Math.max(1, cell.properties?.gridSpan ?? 1);
      if ((width && width.type !== "auto" && width.value > 0) || gridCol + span > numCols) {
        return undefined;
      }
      const start = gridCol;
      gridCol += span;
      // A vertically merged cell's continuation holds nothing of its own.
      if (cell.properties?.verticalMerge === "continue") {
        continue;
      }
      const margins = resolveCellMarginsTwips(table.properties, cell.properties);
      const inset = cellSpacingInsetPt(table, start === 0, gridCol >= numCols);
      const pad = twipsToPt(margins.left + margins.right) + inset.left + inset.right;
      let cellMin = 0;
      let cellMax = 0;
      if (isRotatedTextDirection(cell.properties?.textDirection)) {
        // A rotated column is as wide as its paragraphs are high, each set on
        // one line, whatever the text's length (pages2 P21: one 11pt line with
        // 8pt after → 28.9pt of text room, three → 74.8, five → 120.7, a 20pt
        // line → 40.9, a single-spaced line without spacing → 20.0).
        let height = ROTATED_COLUMN_EXTRA_PT;
        for (const block of cell.content) {
          if (block.type !== "paragraph") {
            return undefined;
          }
          height +=
            layoutParagraph(
              block,
              0,
              ROTATED_LINE_WIDTH_PT,
              options,
              undefined,
              undefined,
              undefined,
              activeStyles?.cellContext(source, r, start, numCols)
            ).rect.height + ROTATED_PARAGRAPH_EXTRA_PT;
        }
        cellMin = cellMax = height + pad;
      } else {
        for (const block of cell.content) {
          if (block.type === "table") {
            const nested = nestedTableContentWidthsPt(block, options);
            if (!nested) {
              return undefined;
            }
            cellMin = Math.max(cellMin, nested.min + pad);
            cellMax = Math.max(cellMax, nested.max + pad);
            continue;
          }
          if (block.type !== "paragraph") {
            return undefined;
          }
          const resolved = activeStyles?.paragraph(
            block,
            activeStyles.cellContext(source, r, start, numCols)
          );
          const styleRunProps = resolved?.runProperties;
          // The paragraph's indents take room in the cell too: a wrapped line
          // starts at the left indent, the first at the hanging indent or the
          // list's text position (pages2 R11f: bullets and a hanging indent).
          const indent = resolved?.paragraphProperties.indent ?? block.properties?.indent;
          const marker = activeListMarkers?.get(block);
          const left = Math.max(indent?.left ? twipsToPt(indent.left) : 0, marker?.indentPt ?? 0);
          const right = indent?.right ? twipsToPt(indent.right) : 0;
          const firstStart = marker
            ? left
            : left -
              (indent?.hanging ? twipsToPt(indent.hanging) : 0) +
              (indent?.firstLine ? twipsToPt(indent.firstLine) : 0);
          const wordPad = pad + Math.max(left, firstStart) + right;
          let line = Math.max(0, firstStart) + right;
          for (const seg of collectParagraphSegments(block, styleRunProps)) {
            // An inline picture is one unbreakable piece of its width (pages2
            // P19: a column holding a 40pt picture is 40pt plus the cell margins).
            if ("type" in seg && seg.type === "image") {
              const w = emuToPt(seg.content.width);
              line += w;
              cellMin = Math.max(cellMin, w + pad + FIT_EPSILON_PT);
              continue;
            }
            // A tab's room depends on where it lands; it is counted as the four
            // spaces it was before tabs were laid out (not measured in Word).
            const text = isTextSegment(seg)
              ? seg.text
              : "type" in seg && seg.type === "tab"
                ? "    "
                : undefined;
            if (text === undefined) {
              return undefined;
            }
            const size = getRunFontSizePt(seg.properties);
            const measure = (text: string): number =>
              measureRunText(
                text,
                seg.properties,
                size,
                options,
                seg.properties?.bold,
                seg.properties?.italic
              );
            line += measure(text);
            for (const word of text.split(/\s+/)) {
              cellMin = Math.max(cellMin, measure(word) + wordPad + FIT_EPSILON_PT);
            }
          }
          // A hair over the content, so the text measured into the column is
          // not broken by rounding when it is measured again to wrap.
          cellMax = Math.max(cellMax, line + pad + FIT_EPSILON_PT);
        }
      }
      if (span === 1) {
        minW[start] = Math.max(minW[start]!, cellMin);
        maxW[start] = Math.max(maxW[start]!, cellMax);
      } else {
        spanning.push({ start, span, min: cellMin, max: cellMax });
      }
    }
  }
  // A spanning cell widens the columns it covers only by what they lack, in
  // proportion to what they already hold (evenly when they hold nothing), as
  // CSS 2.1 automatic layout does.
  const widen = (widths: number[], start: number, span: number, need: number): void => {
    const cols = widths.slice(start, start + span);
    const have = cols.reduce((a, b) => a + b, 0);
    if (need <= have) {
      return;
    }
    for (let k = 0; k < span; k++) {
      widths[start + k] = have > 0 ? (cols[k]! * need) / have : need / span;
    }
  };
  for (const { start, span, min, max } of spanning.sort((x, y) => x.span - y.span)) {
    widen(minW, start, span, min);
    widen(maxW, start, span, Math.max(max, sumOf(minW, start, span)));
  }
  for (let c = 0; c < numCols; c++) {
    maxW[c] = Math.max(maxW[c]!, minW[c]!);
  }
  return { minW, maxW };
}

/** The sum of `widths[start .. start + span)`. */
function sumOf(widths: readonly number[], start: number, span: number): number {
  let total = 0;
  for (let k = start; k < start + span; k++) {
    total += widths[k] ?? 0;
  }
  return total;
}

/**
 * Resolve a table's per-column widths in points.
 *
 * If `table.columnWidths` (twips) is present and covers all columns, it
 * is used and proportionally scaled to fit `contentWidth` (so a table
 * authored wider than the page shrinks to fit rather than overflowing).
 * Otherwise the content width is divided equally among the columns.
 */
function resolveColumnWidthsPt(table: Table, numCols: number, contentWidth: number): number[] {
  // The same scaling the pagination pass applies, expressed in points. Shrinks an
  // overflowing grid and expands an under-wide one, matching how Word distributes a
  // table set to a percentage or auto width.
  return resolveColumnWidthsTwips(table, numCols, ptToTwips(contentWidth)).map(twipsToPt);
}

// =============================================================================
// Internal: Text Helpers
// =============================================================================

interface TextSegment {
  readonly type?: undefined;
  readonly text: string;
  readonly properties: Run["properties"];
  /**
   * Footnote / endnote ids whose reference mark this segment *is*.
   *
   * Carried through wrapping so the paragraph can report which line each note is
   * referenced from. That is what lets a paragraph split across a page boundary
   * hand each note to the page its reference actually landed on.
   */
  readonly noteIds?: readonly number[];
  /**
   * Bookmark names that start at this point in the paragraph.
   *
   * Carried through wrapping for the same reason as `noteIds`: a bookmark in the
   * part of a paragraph that spills onto the next page belongs to that page, and
   * the only way to know is to see which line its marker landed on.
   */
  readonly bookmarks?: readonly string[];
  /** A list's number or bullet, which sizes its line above the baseline only. */
  readonly listMarker?: boolean;
}

/**
 * Inline image segment within a paragraph. Carries the source
 * `InlineImageContent` so the wrap engine can treat it as an
 * unbreakable atom (own width / height in EMU) and the renderer can
 * pull bytes from `imageMap` later.
 */
interface ImageSegment {
  readonly type: "image";
  readonly content: InlineImageContent;
  /** Optional run properties the image inherits (color, etc.). */
  readonly properties?: Run["properties"];
}

/**
 * A hard line break (`<w:br/>`).
 *
 * It has to be its own token rather than a `"\n"` inside a text segment: the
 * wrap engines split text on `/\s+/`, which swallowed the newline as ordinary
 * whitespace, so every hard break vanished — a fenced code block collapsed onto
 * one line and Word's Shift+Enter did nothing.
 */
interface BreakSegment {
  readonly type: "break";
  /** Run properties the break inherits, so it can size an empty line. */
  readonly properties?: Run["properties"];
  /**
   * A `w:br w:type="page"` rather than a line break.
   *
   * It ends the line like any hard break, but it also ends the *page*. Treating
   * one as a plain line break — and separately telling the paginator to move the
   * whole paragraph down — put the text before the break on the wrong page.
   */
  readonly pageBreak?: boolean;
  /**
   * A `w:br w:type="column"`. It ends the line like any hard break — what
   * follows, if only the paragraph mark, takes a line of its own — and ends the
   * column (pages2 P12c).
   */
  readonly columnBreak?: boolean;
}

/**
 * Paragraph-level token: a text segment, an inline image, or a hard break.
 * Returned by `collectParagraphSegments` so wrap algorithms can
 * thread images through without losing them.
 */
/**
 * A `w:tab`. Its width depends on where it falls on its line and on the text
 * after it, so the wrap engine sets `width` (see `tabAdvance`) and the line
 * builder draws `leader` across it.
 */
interface TabSegment {
  readonly type: "tab";
  readonly properties?: Run["properties"];
  readonly width?: number;
  readonly leader?: TabStopLeader;
}

type ParagraphSegment = TextSegment | ImageSegment | BreakSegment | TabSegment;

/** Whether a segment carries drawable text (as opposed to an image or a break). */
function isTextSegment(seg: ParagraphSegment): seg is TextSegment {
  return !("type" in seg) || seg.type === undefined;
}

/**
 * Walk a paragraph's children and emit a flat sequence of paragraph
 * segments — text runs preserve their formatting; inline images become
 * dedicated `ImageSegment` tokens so the wrap engine treats them as
 * unbreakable atoms positioned in document order. Hyperlinks are
 * descended into; bookmark / comment / track-change wrappers are
 * ignored for layout purposes.
 */
function collectParagraphSegments(
  para: Paragraph,
  styleRunProps: Run["properties"] | undefined
): ParagraphSegment[] {
  const segments: ParagraphSegment[] = [];
  const pushVisible = (run: Run) => {
    const properties = activeStyles
      ? layoutRunProperties(activeStyles, run, styleRunProps)
      : mergeRunProperties(styleRunProps, run.properties);
    // Hidden text (`w:vanish`, directly or through a style) takes no space and
    // draws nothing — the converters omit it too.
    if (!isHiddenRun(properties)) {
      pushRunSegments(run, segments, properties);
    }
  };
  for (const child of para.children) {
    const run = finalViewRun(child);
    if (run) {
      // Plain runs and the inserted / moved-to side of a revision; deleted and
      // moved-from text is absent from the final view.
      pushVisible(run);
    } else if (isHyperlink(child)) {
      for (const inner of finalViewRuns(child.children)) {
        pushVisible(inner);
      }
    } else if ("type" in child && child.type === "bookmarkStart" && "name" in child) {
      // An empty marker segment: it draws nothing, but it records the position so
      // a split can tell which page the bookmark ended up on.
      segments.push({ text: "", properties: undefined, bookmarks: [child.name] });
    }
  }
  return segments;
}

/**
 * Emit `ParagraphSegment` tokens for a single run, preserving the
 * relative order of text fragments and inline images. Consecutive
 * text-bearing entries are coalesced into one `TextSegment` so the
 * wrap engine sees fewer atoms.
 *
 * `properties` is the run's *effective* formatting (see `layoutRunProperties`),
 * not `run.properties` — every segment must carry the resolved values so
 * measurement, line height and drawing all agree.
 */
function pushRunSegments(run: Run, out: ParagraphSegment[], properties: Run["properties"]): void {
  let pending = "";
  const flush = () => {
    if (pending.length > 0) {
      out.push(...caseSegments(pending, properties));
      pending = "";
    }
  };
  for (const item of run.content) {
    if (item.type === "text") {
      // A line break inside text is one, as the writer writes it (`<w:br/>`),
      // and so is a tab, as Word sets one kept in `<w:t>` — neither is a
      // character a face has a glyph for.
      if (!/[\t\n\r]/.test(item.text)) {
        pending += item.text;
        continue;
      }
      for (const piece of item.text.split(/(\r\n|\r|\n|\t)/)) {
        if (piece === "\t") {
          flush();
          out.push({ type: "tab", properties });
        } else if (piece === "\n" || piece === "\r" || piece === "\r\n") {
          flush();
          out.push({ type: "break", properties });
        } else {
          pending += piece;
        }
      }
    } else if (item.type === "tab") {
      flush();
      out.push({ type: "tab", properties });
    } else if (item.type === "break") {
      // A hard break ends the current line; see `BreakSegment`.
      flush();
      const breakType = (item as { breakType?: string }).breakType;
      out.push({
        type: "break",
        properties,
        ...(breakType === "page" ? { pageBreak: true } : {}),
        ...(breakType === "column" ? { columnBreak: true } : {})
      });
    } else if (item.type === "image") {
      flush();
      out.push({ type: "image", content: item, properties });
    } else if (item.type === "opaqueRun") {
      // A VML shape set in the line — a text box, as the writer writes one —
      // takes its size there like a picture (pages2 R2i).
      const size = inlineVmlSize(item.rawXml);
      if (size) {
        flush();
        out.push({
          type: "image",
          content: { type: "image", rId: "", width: size.width, height: size.height },
          properties
        });
      }
    } else if (item.type === "symbol") {
      // `w:sym`: the character itself, drawn with the run's font. A symbol font
      // code in the private-use area draws whatever the embedded face maps it to.
      pending += symbolText(item);
    } else if (item.type === "field") {
      pending += fieldDisplayText(item);
    } else if (item.type === "footnoteRef" || item.type === "endnoteRef") {
      // The reference mark itself. Without it a note printed at the foot of the
      // page with nothing in the text pointing at it — and, because the mark was
      // absent from the line boxes, there was no way to tell which page a
      // reference had landed on when its paragraph was split.
      flush();
      const mark = noteMarkFor(item.type, item.id);
      if (mark.length > 0) {
        // Only a footnote is placed on the page of its reference; endnote ids
        // are a separate numbering and would name the wrong footnote.
        out.push({
          text: mark,
          properties,
          ...(item.type === "footnoteRef" ? { noteIds: [item.id] } : {})
        });
      }
    }
  }
  flush();
}

/**
 * A run's text as Word sets it: `w:caps` in capitals; `w:smallCaps` with its
 * lowercase letters as capitals at 80% of the size (pages2 P26: "abcdefghij
 * Mixed Case xyz" in 16pt small caps is 160.0pt wide, as at 12.8pt; at 11pt
 * Word is 0.6pt wider than 8.8pt capitals give). Caps win over small caps.
 */
function caseSegments(text: string, properties: Run["properties"]): TextSegment[] {
  if (properties?.caps === true) {
    return [{ text: text.toUpperCase(), properties }];
  }
  if (properties?.smallCaps !== true) {
    return [{ text, properties }];
  }
  const small = { ...properties, size: getRunFontSizePt(properties) * 2 * SMALL_CAPS_SCALE };
  const out: TextSegment[] = [];
  let stretch = "";
  let lower = false;
  for (const ch of text) {
    const isLower = ch !== ch.toUpperCase();
    if (stretch !== "" && isLower !== lower) {
      out.push(
        lower ? { text: stretch.toUpperCase(), properties: small } : { text: stretch, properties }
      );
      stretch = "";
    }
    stretch += ch;
    lower = isLower;
  }
  out.push(
    lower ? { text: stretch.toUpperCase(), properties: small } : { text: stretch, properties }
  );
  return out;
}

/** The size of `w:smallCaps` capitals standing for lowercase letters. */
const SMALL_CAPS_SCALE = 0.8;

/**
 * A run's advance with `w:w` (horizontal scale) and `w:spacing` (added after
 * every character, the last included, and not scaled) applied (pages2 P26).
 */
function stretchedRunWidth(width: number, text: string, properties: Run["properties"]): number {
  const scale = properties?.scale;
  const spacing = properties?.spacing;
  let out = scale !== undefined && scale > 0 && scale !== 100 ? (width * scale) / 100 : width;
  if (spacing) {
    let count = 0;
    for (const _ of text) {
      count++;
    }
    out += count * twipsToPt(spacing);
  }
  return out;
}

/**
 * The text a field shows. Page-number fields (`PAGE`, `NUMPAGES`,
 * `SECTIONPAGES`) are computed while a header or footer is laid out (see
 * `activePageFields`); every other field — and these outside that pass, or
 * before the totals they need are known — shows its cached result, as the
 * converters do.
 */
function fieldDisplayText(field: FieldContent): string {
  if (field.fieldChar) {
    // One character of a field read in pieces: its result is the content
    // between its pieces, laid out as it is.
    return "";
  }
  const values = activePageFields;
  if (values) {
    const parsed = parsePageField(field.instruction);
    const n =
      parsed?.name === "PAGE"
        ? values.page
        : parsed?.name === "NUMPAGES"
          ? values.numPages
          : values.sectionPages;
    const text = parsed && n !== undefined ? formatPageField(parsed, n, values.format) : undefined;
    if (text !== undefined) {
      return text;
    }
  }
  return field.cachedValue ?? "";
}

/**
 * The text of a note's reference mark: its position among the document's notes,
 * numbered from 1 as Word does by default.
 *
 * Ids at or below zero are the separator and continuation pseudo-notes ECMA-376
 * stores alongside the real ones; they are never referenced from the text.
 */
function noteMarkFor(kind: "footnoteRef" | "endnoteRef", id: number): string {
  const notes = kind === "footnoteRef" ? activeDoc?.footnotes : activeDoc?.endnotes;
  if (!notes) {
    return "";
  }
  const numbered = notes.filter(n => n.id > 0);
  const index = numbered.findIndex(n => n.id === id);
  return index >= 0 ? String(index + 1) : "";
}

/**
 * Wrap a paragraph's text segments into lines whose available widths
 * vary based on per-line wrap exclusions (square / tight / through
 * floats). Word-level break points only — character-level shaping is
 * out of scope for the layout engine.
 *
 * Returns both the per-line `TextSegment[]` and a parallel array of
 * `{ xOffset, width }` describing where each line is placed within
 * the content area. Callers use the slot to set per-line indent /
 * available width for alignment.
 */
/**
 * A unit of wrapping: a word, a run of whitespace, an inline image or a hard
 * break, carrying its measured width and whether a line may break before it.
 */
interface WrapAtom {
  readonly kind: "word" | "space" | "image" | "break" | "tab";
  /** Measured width in points. Zero for a break and for a marker segment. */
  readonly width: number;
  /**
   * There is no break opportunity between this atom and the one before it.
   *
   * A line may break at whitespace, and after a hard break — nowhere else. Two
   * words that meet directly are one unbreakable cluster even when they came
   * from different runs, which is exactly what an inline code span and the
   * punctuation beside it are: `` `sybase.ts`, `` is two runs and one word.
   * Treating every run boundary as a break opportunity instead put that comma
   * alone at the head of the next line, and split `bridge (` from the code span
   * it had just opened.
   */
  readonly glued: boolean;
  /** Text of a word or space atom; empty for an image, a break or a marker. */
  readonly text: string;
  /**
   * The segment this atom came from, so reassembly keeps its run properties and
   * carries its note ids and bookmarks onto the line the atom landed on.
   */
  readonly segment: ParagraphSegment;
  /** Measurer for this atom's formatting, for a mid-word split. */
  readonly measure: (text: string) => number;
  /** A tab's leader, from the stop it reached. */
  readonly leader?: TabStopLeader;
}

/** Measurer for atoms that carry no text. */
const MEASURE_NOTHING = (): number => 0;

/**
 * Split a paragraph's segments into wrap atoms.
 *
 * One tokenizer serves both wrappers below. They differ in how a line's width
 * is chosen — a fixed measure, or a slot between floats — not in what a line is
 * made of, and while each had its own tokenizer the two could disagree about
 * where a break was allowed.
 */
function tokenizeSegments(
  segments: readonly ParagraphSegment[],
  headingScale: number,
  options: FullLayoutOptions | undefined
): WrapAtom[] {
  const atoms: WrapAtom[] = [];
  // One memoised measurer per distinct formatting rather than per segment: a
  // paragraph of many short runs shares a handful of faces between them.
  const measurers = new Map<string, (text: string) => number>();
  const measurerFor = (properties: Run["properties"]): ((text: string) => number) => {
    const fontSize = getRunFontSizePt(properties) * headingScale;
    const fonts = resolveRunFonts(properties);
    const bold = properties?.bold === true;
    const italic = properties?.italic === true;
    const key = `${fonts.ascii}\u0000${fonts.eastAsia}|${fontSize}|${bold ? "b" : ""}${italic ? "i" : ""}|${properties?.scale ?? ""}|${properties?.spacing ?? ""}`;
    let measure = measurers.get(key);
    if (measure === undefined) {
      const latin = memoizedWordMeasure(fonts.ascii, fontSize, options, bold, italic);
      if (fonts.eastAsia === fonts.ascii) {
        measure = latin;
      } else {
        // A run may name a different typeface for each script, so a mixed atom
        // has to be measured in stretches. Atoms are already single-script in
        // practice — `segmentForWrap` breaks at the Latin/ideograph boundary —
        // but a marker or an atom assembled elsewhere need not be, so this does
        // not assume it.
        const eastAsian = memoizedWordMeasure(fonts.eastAsia, fontSize, options, bold, italic);
        measure = (text: string): number => {
          let total = 0;
          for (const run of splitByScript(text)) {
            total += run.cjk ? eastAsian(run.text) : latin(run.text);
          }
          return total;
        };
      }
      if ((properties?.scale !== undefined && properties.scale !== 100) || properties?.spacing) {
        const plain = measure;
        measure = (text: string): number => stretchedRunWidth(plain(text), text, properties);
      }
      measurers.set(key, measure);
    }
    return measure;
  };
  /**
   * Whether an atom pushed now would be glued to the one before it.
   *
   * `nextText` is the text about to be pushed. It matters because an East Asian
   * character opens a break opportunity on both sides regardless of which run it
   * came from: without consulting it, `中文` and `报表` in two different runs
   * became one unbreakable cluster, which is the same defect as a whole Chinese
   * paragraph being one atom, just at a smaller scale. For Latin, `canBreakBetween`
   * returns false between two letters, so the run-boundary gluing this function
   * was written for is preserved exactly.
   */
  const gluedToPrevious = (nextText?: string): boolean => {
    const previous = atoms[atoms.length - 1];
    if (
      previous === undefined ||
      previous.kind === "space" ||
      previous.kind === "break" ||
      previous.kind === "tab"
    ) {
      return false;
    }
    if (nextText !== undefined && nextText.length > 0 && previous.text.length > 0) {
      const prevChars = [...previous.text];
      const prevCp = prevChars[prevChars.length - 1].codePointAt(0)!;
      if (canBreakBetween(prevCp, nextText.codePointAt(0)!)) {
        return false;
      }
    }
    return true;
  };

  for (const segment of segments) {
    if ("type" in segment && segment.type === "break") {
      atoms.push({
        kind: "break",
        width: 0,
        glued: false,
        text: "",
        segment,
        measure: MEASURE_NOTHING
      });
      continue;
    }
    if ("type" in segment && segment.type === "tab") {
      // Its width is set where it lands on a line (`tabAdvance`).
      atoms.push({
        kind: "tab",
        width: 0,
        glued: false,
        text: "",
        segment,
        measure: MEASURE_NOTHING
      });
      continue;
    }
    if ("type" in segment && segment.type === "image") {
      atoms.push({
        kind: "image",
        width: emuToPt(segment.content.width),
        glued: gluedToPrevious(),
        text: "",
        segment,
        measure: MEASURE_NOTHING
      });
      continue;
    }
    const measure = measurerFor(segment.properties);
    if (segment.text.length === 0) {
      // A marker segment draws nothing, but a bookmark or footnote reference it
      // carries has a position — so it travels with the word it sits inside
      // rather than opening a break opportunity of its own.
      atoms.push({ kind: "word", width: 0, glued: gluedToPrevious(), text: "", segment, measure });
      continue;
    }
    for (const token of segment.text.split(/(\s+)/)) {
      if (token.length === 0) {
        continue;
      }
      if (/^\s/.test(token)) {
        atoms.push({
          kind: "space",
          width: measure(token),
          glued: false,
          text: token,
          segment,
          measure
        });
        continue;
      }
      // Subdivide the word at East Asian break opportunities. Splitting on
      // whitespace alone made an entire Chinese paragraph one atom, so it could
      // only be placed whole: a line already holding Latin text flushed early
      // and left a hole, and the paragraph reached a line of its own before the
      // mid-word fallback (`splitTextToFit`) cut it at an arbitrary character
      // with no regard for kinsoku. `segmentForWrap` glues a character that may
      // not begin a line to the piece before it, so the atoms themselves carry
      // the prohibition and the wrapper needs no extra rule.
      let firstPiece = true;
      for (const piece of segmentForWrap(token)) {
        atoms.push({
          kind: "word",
          width: measure(piece),
          // Only the first piece can be glued to what came before; the
          // boundaries between pieces are break opportunities by construction.
          glued: firstPiece ? gluedToPrevious(piece) : false,
          text: piece,
          segment,
          measure
        });
        firstPiece = false;
      }
    }
  }
  return atoms;
}

/** A paragraph's tab stops, in points from the left of its text column. */
interface TabStops {
  readonly stops: readonly {
    readonly position: number;
    readonly type: "left" | "center" | "right" | "decimal";
    readonly leader?: TabStopLeader;
  }[];
  /** `w:defaultTabStop`: Word's stops at every multiple, after the last custom one. */
  readonly step: number;
}

/**
 * Resolve a paragraph's stops. A hanging indent makes the left indent a stop
 * too, so the text after a hanging tab lands on it. A bar tab draws a rule and
 * stops nothing; `start`/`end` are left/right in left-to-right text.
 */
function resolveTabStops(
  tabs: readonly TabStop[] | undefined,
  defaultTabStopTwips: number | undefined,
  hangingStopPt: number | undefined
): TabStops {
  const stops: TabStops["stops"][number][] = [];
  for (const tab of tabs ?? []) {
    if (tab.type === "bar" || tab.type === "clear") {
      continue;
    }
    const type =
      tab.type === "start" || tab.type === "num" ? "left" : tab.type === "end" ? "right" : tab.type;
    stops.push({ position: twipsToPt(tab.position), type, leader: tab.leader });
  }
  if (hangingStopPt !== undefined && !stops.some(t => t.position === hangingStopPt)) {
    stops.push({ position: hangingStopPt, type: "left" });
  }
  stops.sort((a, b) => a.position - b.position);
  return { stops, step: twipsToPt(defaultTabStopTwips || 720) };
}

/**
 * The room a tab at `x` takes on its line, and the leader it draws across it;
 * `undefined` when no stop lies after `x` within `right`, which sends the tab
 * to the next line (pages2 R7j: a tab past a right stop at 468pt of a 481.9pt
 * column wraps, and on the new line reaches the first stop).
 *
 * A right stop ends the text up to the next tab there, a centre stop centres it,
 * a decimal stop puts its first `.` there (none: like a right stop). Text too
 * long to end at the stop starts right after the tab.
 */
function tabAdvance(
  tabStops: TabStops,
  atoms: readonly WrapAtom[],
  index: number,
  x: number,
  right: number
): { width: number; leader?: TabStopLeader } | undefined {
  const EPS = 0.01;
  let stop: TabStops["stops"][number] | undefined = tabStops.stops.find(t => t.position > x + EPS);
  if (stop === undefined) {
    const last = tabStops.stops.at(-1)?.position ?? 0;
    const from = Math.max(x, last);
    const position = (Math.floor(from / tabStops.step + EPS) + 1) * tabStops.step;
    stop = { position, type: "left" };
  }
  if (stop.position > right + EPS) {
    return undefined;
  }
  let following = 0;
  let beforePoint: number | undefined;
  for (let i = index + 1; i < atoms.length; i++) {
    const atom = atoms[i]!;
    if (atom.kind === "tab" || atom.kind === "break") {
      break;
    }
    if (beforePoint === undefined && atom.kind === "word") {
      const point = atom.text.indexOf(".");
      if (point >= 0) {
        beforePoint = following + atom.measure(atom.text.slice(0, point));
      }
    }
    following += atom.width;
  }
  const aligned =
    stop.type === "right"
      ? following
      : stop.type === "center"
        ? following / 2
        : stop.type === "decimal"
          ? (beforePoint ?? following)
          : 0;
  // A hair short of the stop, so the text measured to end there is not pushed
  // off the line by rounding when the wrap adds it back.
  return {
    width: Math.max(0, stop.position - x - aligned - (aligned > 0 ? 1e-6 : 0)),
    leader: stop.leader
  };
}

const LEADER_CHARS: Partial<Record<TabStopLeader, string>> = {
  dot: ".",
  hyphen: "-",
  underscore: "_",
  heavy: "_",
  middleDot: "\u00b7"
};

/** Runs drawn as a tab's leader, which carry no text of the paragraph. */
const leaderRuns = new WeakSet<object>();

/** The leader a tab draws across its room: whole characters, ending at the stop. */
function tabLeaderRun(
  seg: TabSegment,
  x: number,
  headingScale: number,
  options: FullLayoutOptions | undefined
): PositionedRun | undefined {
  const char = seg.leader ? LEADER_CHARS[seg.leader] : undefined;
  if (char === undefined || !seg.width) {
    return undefined;
  }
  const fontSize = getRunFontSizePt(seg.properties) * headingScale;
  const fonts = resolveRunFonts(seg.properties);
  const charWidth = measureLayoutText(char, fonts.ascii, fontSize, options);
  const count = charWidth > 0 ? Math.floor(seg.width / charWidth) : 0;
  if (count === 0) {
    return undefined;
  }
  const run: PositionedRun = {
    text: char.repeat(count),
    x: x + seg.width - count * charWidth,
    width: count * charWidth,
    font: fonts.ascii,
    fontSize,
    color: resolveColorHex(seg.properties?.color)
  };
  leaderRuns.add(run);
  return run;
}

/** Width of the spaces between the words of a line being filled. */
function interiorSpace(line: readonly WrapAtom[]): number {
  let total = 0;
  let seenWord = false;
  for (const atom of line) {
    if (atom.kind === "space") {
      total += seenWord ? atom.width : 0;
    } else {
      seenWord = true;
    }
  }
  return total;
}

/** Whether an atom is a `w:br w:type="page"` (see `BreakSegment.pageBreak`). */
function isPageBreakAtom(atom: WrapAtom): boolean {
  const segment = atom.segment;
  return "type" in segment && segment.type === "break" && segment.pageBreak === true;
}

/**
 * How many atoms at the end of `placed` belong to the same unbreakable cluster
 * as `next`, and therefore have to move down with it.
 *
 * Zero when `next` may legally open a line, and zero when the cluster is the
 * whole line: a cluster wider than the measure has to break somewhere, and the
 * boundary already reached is as good a place as any.
 */
function gluedTailLength(placed: readonly WrapAtom[], next: WrapAtom): number {
  if (!next.glued) {
    return 0;
  }
  let count = 0;
  while (count < placed.length) {
    const atom = placed[placed.length - 1 - count];
    count++;
    if (!atom.glued) {
      break;
    }
  }
  return count >= placed.length ? 0 : count;
}

/**
 * The width `placed` would have left after giving back its last `count` atoms.
 *
 * Both wrappers use this as the test for whether a retreat is worth making: a
 * retreat that empties the line has not moved a cluster down, it has produced a
 * blank line and left the cluster to be broken anyway. The remaining width is
 * only ever read as that predicate — the closing line takes its width from the
 * slot, not from the atoms — so it is computed here rather than written back.
 */
function widthWithoutTail(placed: readonly WrapAtom[], count: number, total: number): number {
  let width = total;
  for (let i = 0; i < count; i++) {
    width -= placed[placed.length - 1 - i].width;
  }
  return width;
}

/**
 * Turn a line's atoms back into paragraph segments.
 *
 * Consecutive atoms from the same source segment merge, so a run that arrived
 * whole leaves whole: one segment per atom would make the renderer emit a
 * separately positioned string per word, re-adding by hand the spacing the wrap
 * had just measured.
 */
function reassembleLine(atoms: readonly WrapAtom[]): ParagraphSegment[] {
  const out: ParagraphSegment[] = [];
  let source: ParagraphSegment | undefined;
  let text = "";
  const flushText = (): void => {
    if (source !== undefined && isTextSegment(source)) {
      out.push(text === source.text ? source : { ...source, text });
    }
    source = undefined;
    text = "";
  };
  for (const atom of atoms) {
    if (atom.kind === "break" || atom.kind === "image") {
      flushText();
      out.push(atom.segment);
      continue;
    }
    if (atom.kind === "tab") {
      flushText();
      out.push({ ...(atom.segment as TabSegment), width: atom.width, leader: atom.leader });
      continue;
    }
    if (atom.segment !== source) {
      flushText();
      source = atom.segment;
    }
    text += atom.text;
  }
  flushText();

  // Trailing spaces are invisible in left-aligned text but they widen the line
  // for centring and right-alignment, and they leave the paragraph's measured
  // width wrong.
  while (out.length > 0) {
    const last = out[out.length - 1];
    if (!isTextSegment(last)) {
      break;
    }
    const trimmed = last.text.replace(/\s+$/, "");
    if (trimmed === last.text) {
      break;
    }
    out.pop();
    if (trimmed.length > 0) {
      out.push({ ...last, text: trimmed });
      break;
    }
    // The segment drew nothing, but a bookmark or note reference it carries
    // still has a position on this line.
    if (last.noteIds !== undefined || last.bookmarks !== undefined) {
      out.push({ ...last, text: "" });
      break;
    }
  }
  return out;
}

function wrapSegmentsToLinesWithExclusions(
  segments: ParagraphSegment[],
  leftIndentPt: number,
  firstLineIndentPt: number,
  headingScale: number,
  lineHeightPt: number,
  paragraphTopPageY: number,
  pageContext: PageLayoutContext,
  options: FullLayoutOptions | undefined,
  spaceShrink = 0,
  tabStops?: TabStops
): { lines: ParagraphSegment[][]; slots: LineSlot[]; emptyDrop?: number } {
  const atoms = tokenizeSegments(segments, headingScale, options);
  const lines: ParagraphSegment[][] = [];
  const slots: LineSlot[] = [];

  // A row a float leaves no room on at all is skipped, and the line moves
  // straight down to the float's bottom — not by whole rows (pages2 R3g: the
  // empty anchor paragraph of a `wrapText="right"` picture at the right margin
  // starts at the picture's bottom, and the text after it follows).
  let y = paragraphTopPageY;
  let drop = 0;
  const blockedUntil = (lineY: number): number | undefined => {
    const gaps = availableSlotsForLine(pageContext, lineY, lineHeightPt);
    if (gaps.length !== 1 || gaps[0]!.width > 0) {
      return undefined;
    }
    let bottom = lineY;
    for (const ex of pageContext.exclusions) {
      if (lineY + lineHeightPt > ex.yTop && lineY < ex.yBottom) {
        bottom = Math.max(bottom, ex.yBottom);
      }
    }
    return bottom > lineY ? bottom : undefined;
  };

  if (atoms.length === 0) {
    const below = blockedUntil(y);
    return below === undefined ? { lines, slots } : { lines, slots, emptyDrop: below - y };
  }

  let cursor = 0;
  let row = 0;
  // The first line of the paragraph is not the result of a wrap, so its leading
  // whitespace is the author's.
  let openedByWrap = false;
  while (cursor < atoms.length) {
    const below = blockedUntil(y);
    if (below !== undefined) {
      drop += below - y;
      y = below;
      continue;
    }
    const lineY = y;
    const gaps = availableSlotsForLine(pageContext, lineY, lineHeightPt);
    const rowStart = lines.length;
    for (let g = 0; g < gaps.length && cursor < atoms.length; g++) {
      const slot = gaps[g]!;
      const lastGap = g === gaps.length - 1;
      // The first line of a paragraph may carry an extra `firstLineIndent`
      // (from `<w:ind firstLine="…"/>`) which subtracts from the usable
      // width on that line only. Subsequent lines use the full slot width
      // (offset by the paragraph's `leftIndentPt`, which is applied by
      // the caller's run x-positioning logic, not here).
      const indentForThisLine = lines.length === 0 ? firstLineIndentPt : 0;
      let usable = slot.width - indentForThisLine;
      let lineXOffset = slot.xOffset + indentForThisLine;

      // Also subtract the paragraph's own leftIndentPt (which the legacy
      // path applies through `availableWidth = contentWidth -
      // leftIndentPt`). We mirror that here so wrap behaviour matches
      // when no exclusion is in play.
      if (slot.xOffset === 0 && leftIndentPt > 0) {
        usable -= leftIndentPt;
        lineXOffset += leftIndentPt;
      }

      if (usable <= 0) {
        if (!lastGap || lines.length > rowStart) {
          continue;
        }
        // Pathological — the line is fully blocked or narrower than the
        // first-line indent. Skip the y position by advancing one line
        // height; placing zero-content lines indefinitely is worse than
        // leaving a small visual gap.
        lines.push([]);
        slots.push({ xOffset: lineXOffset, width: Math.max(0, usable) });
        break;
      }

      const start = cursor;
      const startedByWrap = openedByWrap;
      // Whitespace that opens a *wrapped* line is dropped — it is the separator
      // the break consumed. After a hard break it is the author's (a code block's
      // indentation), so it stays.
      if (openedByWrap && atoms[cursor].kind === "space") {
        cursor++;
        if (cursor >= atoms.length) {
          break;
        }
      }
      openedByWrap = true;

      const lineAtoms: WrapAtom[] = [];
      let lineWidth = 0;
      let hardBreak = false;
      // Every pass consumes an atom, shortens one, or closes a line that already
      // holds content, so this terminates.
      while (cursor < atoms.length) {
        const atom = atoms[cursor];
        if (atom.kind === "break") {
          // A leading page break is a break-before and leaves no blank line —
          // see the sister case in `wrapSegmentsToLines`.
          if (isPageBreakAtom(atom) && lines.length === 0 && lineAtoms.length === 0) {
            cursor++;
            openedByWrap = false;
            continue;
          }
          // A hard break closes this line even though there is room left, and
          // whatever follows keeps its leading whitespace. Kept on the line so a
          // page break's position survives (see `BreakSegment.pageBreak`).
          lineAtoms.push(atom);
          cursor++;
          openedByWrap = false;
          hardBreak = true;
          break;
        }
        if (atom.kind === "tab") {
          const advance = tabStops
            ? tabAdvance(tabStops, atoms, cursor, lineXOffset + lineWidth, lineXOffset + usable)
            : { width: 0 };
          if (advance === undefined && lineAtoms.length > 0) {
            break;
          }
          lineAtoms.push({ ...atom, width: advance?.width ?? 0, leader: advance?.leader });
          lineWidth += advance?.width ?? 0;
          cursor++;
          continue;
        }
        if (lineWidth + atom.width - spaceShrink * interiorSpace(lineAtoms) <= usable) {
          lineAtoms.push(atom);
          lineWidth += atom.width;
          cursor++;
          continue;
        }
        // The atom overflows. Give back the cluster it belongs to, so the whole
        // cluster moves down together rather than breaking at a run boundary.
        const retreat = gluedTailLength(lineAtoms, atom);
        if (retreat > 0 && widthWithoutTail(lineAtoms, retreat, lineWidth) > 0) {
          lineAtoms.length -= retreat;
          cursor -= retreat;
          break;
        }
        if (lineWidth > 0 || !lastGap) {
          // Something is already on the line, or a later gap of the row may
          // take it — try again there.
          break;
        }
        // Alone on the line and still too wide: break inside the word, the way
        // CSS `overflow-wrap: break-word` does. An image is left whole — there is
        // nothing to break — and overflows as Word lets it.
        if (atom.kind === "word" && atom.text.length > 0) {
          const { head, tail } = splitTextToFit(atom.text, atom.measure, usable);
          if (tail.length > 0) {
            // Replace the atom with its remainder so the next line continues
            // from the break point.
            atoms[cursor] = { ...atom, text: tail, width: atom.measure(tail), glued: true };
            const headWidth = atom.measure(head);
            lineAtoms.push({ ...atom, text: head, width: headWidth });
            lineWidth += headWidth;
            break;
          }
        }
        lineAtoms.push(atom);
        lineWidth += atom.width;
        cursor++;
      }

      if (lineAtoms.length === 0 && !lastGap) {
        // Nothing fits in this gap: the next one on the row gets the text.
        cursor = start;
        openedByWrap = startedByWrap;
        continue;
      }
      lines.push(reassembleLine(lineAtoms));
      slots.push({ xOffset: lineXOffset, width: usable });
      if (hardBreak) {
        break;
      }
    }
    // Every line of the row but the last shares its y with the next.
    for (let k = rowStart; k < lines.length - 1; k++) {
      slots[k] = { ...slots[k]!, sameRow: true };
    }
    if (lines.length === rowStart) {
      // A row with no room at all: leave it empty rather than loop.
      lines.push([]);
      slots.push({ xOffset: 0, width: 0 });
    }
    if (drop > 0) {
      slots[rowStart] = { ...slots[rowStart]!, drop };
      drop = 0;
    }
    row++;
    y += lineHeightPt;

    if (row > 100_000) {
      // Defensive — degenerate inputs shouldn't loop the engine.
      break;
    }
  }
  if (endsWithLineBreak(atoms)) {
    lines.push([]);
    const lineY = y;
    const slot = availableSlotsForLine(pageContext, lineY, lineHeightPt)[0]!;
    slots.push({
      xOffset: slot.xOffset + (slot.xOffset === 0 ? leftIndentPt : 0),
      width: slot.width
    });
  }

  return { lines, slots };
}

/** Where one line of a paragraph wrapping around floats is set. */
interface LineSlot {
  readonly xOffset: number;
  readonly width: number;
  /** The next line continues this one's row, right of a float between them. */
  readonly sameRow?: boolean;
  /** How far the line is moved down, below floats that left its rows no room. */
  readonly drop?: number;
}

/**
 * A measuring function for one run's formatting, caching short strings.
 *
 * Character-level breaking measures the same code points repeatedly — a CJK
 * paragraph reuses a few hundred distinct glyphs across thousands of positions —
 * so the cache turns that from a hot path into a lookup.
 */
function memoizedWordMeasure(
  fontName: string,
  fontSize: number,
  options: FullLayoutOptions | undefined,
  bold: boolean | undefined,
  italic: boolean | undefined
): (text: string) => number {
  const cache = new Map<string, number>();
  return (text: string): number => {
    if (text.length > 8) {
      return measureLayoutText(text, fontName, fontSize, options, bold, italic);
    }
    const hit = cache.get(text);
    if (hit !== undefined) {
      return hit;
    }
    const width = measureLayoutText(text, fontName, fontSize, options, bold, italic);
    cache.set(text, width);
    return width;
  };
}

/**
 * The longest prefix of `text` that fits `room`, and the rest.
 *
 * Breaks fall on code-point boundaries, so a surrogate pair is never split into
 * two halves that render as replacement characters. At least one code point is
 * always taken, which guarantees the caller makes progress even in a column too
 * narrow for a single glyph.
 *
 * Widths are accumulated per code point rather than re-measuring every prefix:
 * the engine's own metrics are additive, and quadratic re-measurement of a long
 * CJK paragraph would dominate layout.
 */
function splitTextToFit(
  text: string,
  measure: (s: string) => number,
  room: number
): { head: string; tail: string } {
  let head = "";
  let width = 0;
  let index = 0;
  for (const ch of text) {
    const chWidth = measure(ch);
    if (head.length > 0 && width + chWidth > room) {
      break;
    }
    head += ch;
    width += chWidth;
    index += ch.length;
  }
  return { head, tail: text.slice(index) };
}

function wrapSegmentsToLines(
  segments: ParagraphSegment[],
  availableWidth: number,
  firstLineIndent: number,
  headingScale: number,
  options: FullLayoutOptions | undefined,
  spaceShrink = 0,
  tabStops?: TabStops,
  leftIndentPt = 0
): ParagraphSegment[][] {
  const atoms = tokenizeSegments(segments, headingScale, options);
  const lines: ParagraphSegment[][] = [];
  let line: WrapAtom[] = [];
  let lineWidth = 0;
  let isFirstLine = true;
  let effectiveWidth = availableWidth - firstLineIndent;
  /**
   * Whether the line being built was opened by a *wrap*, in which case
   * whitespace at its head is the separator the break consumed and keeping it
   * indents the line by a stray space. After a hard break — and at the start of
   * the paragraph — leading whitespace is the author's, and in a code block it
   * is the indentation.
   */
  let openedByWrap = false;

  const flushLine = (): void => {
    lines.push(reassembleLine(line));
    line = [];
    lineWidth = 0;
    if (isFirstLine) {
      isFirstLine = false;
      effectiveWidth = availableWidth;
    }
  };

  let index = 0;
  // Every pass consumes an atom, shortens one, or closes a line that already
  // holds content, so this terminates.
  while (index < atoms.length) {
    const atom = atoms[index];

    if (atom.kind === "break") {
      // A *page* break at the very start of a paragraph is a break-before: the
      // block-level machinery has already moved the paragraph, and Word leaves
      // no blank line behind. A leading *line* break does produce an empty line.
      if (isPageBreakAtom(atom) && lines.length === 0 && line.length === 0) {
        openedByWrap = false;
        index++;
        continue;
      }
      // A hard break closes the current line even though there is room left,
      // and whatever follows keeps its leading whitespace. The marker is kept on
      // the line so a page break's position is recoverable; the line-box builder
      // draws nothing for it.
      line.push(atom);
      flushLine();
      openedByWrap = false;
      index++;
      continue;
    }

    if (atom.kind === "tab") {
      const start = leftIndentPt + (isFirstLine ? firstLineIndent : 0);
      const advance = tabStops
        ? tabAdvance(tabStops, atoms, index, start + lineWidth, start + effectiveWidth)
        : { width: 0 };
      if (advance === undefined && line.length > 0) {
        flushLine();
        openedByWrap = true;
        continue;
      }
      line.push({ ...atom, width: advance?.width ?? 0, leader: advance?.leader });
      lineWidth += advance?.width ?? 0;
      index++;
      continue;
    }

    if (atom.kind === "space") {
      if (openedByWrap && lineWidth === 0) {
        index++;
        continue;
      }
      // A space that overflows is trailing whitespace, which `reassembleLine`
      // trims off the line it closes.
      line.push(atom);
      lineWidth += atom.width;
      index++;
      continue;
    }

    // A word or an inline image.
    if (lineWidth + atom.width - spaceShrink * interiorSpace(line) <= effectiveWidth) {
      line.push(atom);
      lineWidth += atom.width;
      index++;
      continue;
    }

    // It does not fit. Give back the cluster it belongs to, so the whole
    // cluster moves down together rather than breaking at a run boundary.
    const retreat = gluedTailLength(line, atom);
    if (retreat > 0 && widthWithoutTail(line, retreat, lineWidth) > 0) {
      line.length -= retreat;
      index -= retreat;
      flushLine();
      openedByWrap = true;
      continue;
    }

    if (lineWidth > 0) {
      // Something is already on the line — try the atom again on a fresh one,
      // intact. Only a cluster that cannot fit a line *by itself* is broken.
      flushLine();
      openedByWrap = true;
      continue;
    }

    // Alone on the line and still too wide. An image is left whole — there is
    // nothing to break — and overflows the way Word lets it. A word breaks
    // inside, the way CSS `overflow-wrap: break-word` does: without that a long
    // URL, a hex digest or any space-less script ran off the right edge.
    if (atom.kind === "image" || atom.text.length === 0) {
      line.push(atom);
      lineWidth += atom.width;
      index++;
      continue;
    }
    const { head, tail } = splitTextToFit(atom.text, atom.measure, effectiveWidth);
    if (tail.length === 0) {
      line.push(atom);
      lineWidth += atom.width;
      index++;
      continue;
    }
    const headWidth = atom.measure(head);
    line.push({ ...atom, text: head, width: headWidth });
    lineWidth += headWidth;
    atoms[index] = { ...atom, text: tail, width: atom.measure(tail), glued: true };
    flushLine();
    openedByWrap = true;
  }

  if (line.length > 0) {
    // Through `flushLine`, so the last line has its trailing whitespace trimmed
    // like every other one.
    flushLine();
  } else if (endsWithLineBreak(atoms)) {
    lines.push([]);
  }

  if (lines.length === 0 && segments.length > 0) {
    lines.push(segments);
  }

  return lines;
}

/**
 * Whether a paragraph's text ends in a line or column break, after which its
 * mark takes a line of its own: Word sets `text<w:br/>` on two lines, and a
 * paragraph of a lone `<w:br/>` on two (pages2 P13). A page break keeps the
 * mark on its own line instead (see `placeParagraph`).
 */
function endsWithLineBreak(atoms: readonly WrapAtom[]): boolean {
  const last = atoms.at(-1);
  return last?.kind === "break" && !isPageBreakAtom(last);
}

/**
 * Resolve the effective font size in points for a run.
 *
 * `<w:sz w:val="…"/>` is in half-points; we halve.
 *
 * Sub/superscript runs are conventionally rendered at ~65 % of the
 * surrounding text's size with a vertical baseline shift. The size
 * scaling lives here so every measurement (line width, line height,
 * wrap) sees the same value; the y-shift is applied at render time
 * via `PositionedRun.verticalAlign`.
 */
function getRunFontSizePt(props: Run["properties"]): number {
  const base = props?.size ? props.size / 2 : DEFAULT_FONT_SIZE_PT;
  if (props?.vertAlign === "superscript" || props?.vertAlign === "subscript") {
    return base * SCRIPT_FONT_SIZE_RATIO;
  }
  return base;
}

/**
 * The natural (single-spaced) line height of a paragraph, in points.
 *
 * Word sizes a line from the tallest run on it. We approximate that with the
 * paragraph's tallest run — one height for the whole paragraph —
 * which keeps wrapping and pagination arithmetic simple while still giving a
 * 24 pt heading a 24 pt-sized line box.
 *
 * Sub/superscript runs are deliberately measured at their shrunken size (they
 * ride the surrounding text's line), and an empty paragraph falls back to its
 * paragraph-mark run properties, then the style's, then the document default.
 */
function naturalParagraphLineHeightPt(
  segments: readonly ParagraphSegment[],
  props: ParagraphProperties | undefined,
  styleRunProps: Run["properties"] | undefined
): number {
  const height = singleLineHeightPt(segments);
  if (height > 0) {
    return height;
  }
  // Only images and/or hard breaks: fall back to the paragraph mark's size,
  // then the style's, then the document default, so an empty line still has
  // the height the author asked for.
  const fromBreak = segments.find(seg => "type" in seg && seg.type === "break")?.properties;
  const mark = runLineExtentsPt(props?.markRunProperties ?? styleRunProps ?? fromBreak);
  return mark.above + mark.below;
}

/**
 * Single-spaced height of a line holding `segments`: the tallest part above the baseline over the tallest below
 * it, among the faces set on the line, as Word sizes it. 0 when the line sets
 * no text — an empty stretch, where a bookmark sits, sizes nothing.
 *
 * A marker's face counts above the baseline only (pages2 P30: a Courier New
 * `o`, which reaches lower than Calibri, leaves the line as tall as its text; a
 * Symbol `•`, which reaches higher, raises it).
 */
function singleLineHeightPt(segments: readonly ParagraphSegment[], withMarker = true): number {
  let above = 0;
  let below = 0;
  const add = (props: Run["properties"]): void => {
    const extents = runLineExtentsPt(props);
    above = Math.max(above, extents.above);
    below = Math.max(below, extents.below);
  };
  for (const seg of segments) {
    if (isTextSegment(seg) && seg.text !== "" && seg.listMarker !== true) {
      add(seg.properties);
    }
  }
  if (withMarker && above + below > 0) {
    for (const seg of segments) {
      if (isTextSegment(seg) && seg.listMarker === true) {
        above = Math.max(above, runLineExtentsPt(seg.properties).above);
      }
    }
  }
  return above + below;
}

/**
 * A run's single-spaced line above and below the baseline, in points: its
 * face's own metrics, as Word sets it (see `font-face`), or
 * {@link LINE_HEIGHT_FACTOR} em split 4:1 for a face without a table.
 */
function runLineExtentsPt(props: Run["properties"]): { above: number; below: number } {
  const sizePt = getRunFontSizePt(props);
  const fonts = resolveRunFonts(props);
  return (
    faceLineExtents(fonts.ascii, sizePt, props?.bold) ?? {
      above: sizePt * LINE_HEIGHT_FACTOR * 0.8,
      below: sizePt * LINE_HEIGHT_FACTOR * 0.2
    }
  );
}

/**
 * The Latin and East Asian typefaces a run draws with.
 *
 * `w:rFonts` names them separately, and this used to read `w:ascii` only — so a
 * run carrying `w:eastAsia="宋体"` was measured with the Latin face, and the
 * typeface the author asked for never reached the layout. Since the PDF bridge
 * measures from this, a Chinese document was laid out against Calibri's metrics
 * and then drawn with whatever face the font manager had discovered.
 *
 * `eastAsia` falls back to the Latin name rather than to `"Calibri"`: a caller
 * that named one typeface for the run meant it for the whole run.
 */
function resolveRunFonts(props: Run["properties"]): { ascii: string; eastAsia: string } {
  const font = props?.font;
  if (!font) {
    return { ascii: "Calibri", eastAsia: "Calibri" };
  }
  if (typeof font === "string") {
    return { ascii: font, eastAsia: font };
  }
  const spec = font as { ascii?: string; hAnsi?: string; eastAsia?: string };
  const ascii = spec.ascii ?? spec.hAnsi ?? "Calibri";
  return { ascii, eastAsia: spec.eastAsia ?? ascii };
}

function measureLayoutText(
  text: string,
  fontName: string,
  fontSize: number,
  options: FullLayoutOptions | undefined,
  bold?: boolean,
  italic?: boolean
): number {
  if (options?.measureText) {
    return options.measureText(text, fontName, fontSize, bold, italic);
  }
  return (
    measureFaceText(text, fontName, fontSize, bold, italic) ??
    measureTextWidth(text, styledFontVariant(fontName, bold, italic), fontSize)
  );
}

/**
 * Whether a line must keep its natural width despite `w:jc="both"`.
 *
 * The last line of a paragraph is never stretched — that is what makes justified
 * text look justified rather than broken. A line ended by a hard break is the
 * last line of its own visual block for the same reason, so it is exempt too.
 */
function isJustifyExemptLine(
  lines: readonly (readonly ParagraphSegment[])[],
  index: number
): boolean {
  if (index >= lines.length - 1) {
    return true;
  }
  const next = lines[index + 1];
  // A `w:br` opens the following line, so the *current* one ended a block.
  const current = lines[index];
  const endsWithBreak = current.some(seg => "type" in seg && seg.type === "break");
  return endsWithBreak || next.length === 0;
}

/**
 * Count the places a line can absorb justification slack.
 *
 * `spaces` counts space characters — the Latin mechanism, and exact, because
 * PDF's `Tw` and SVG's `word-spacing` both add their amount once per space.
 *
 * `ideographs` is the divisor for the East Asian mechanism. Typographically the
 * slack belongs in the gaps *between* characters, one fewer than the count — but
 * `Tc` and `letter-spacing` add after every character including the last, so
 * dividing by the gaps overshoots the column by one unit of spacing. Dividing by
 * the character count instead keeps the line inside its column and leaves the
 * final unit as trailing space, which is invisible (0.18pt on a 468pt column of
 * 100 ideographs) where an overflow would not be.
 */
function countJustifyOpportunities(segments: readonly ParagraphSegment[]): JustifyOpportunities {
  const total: JustifyOpportunities = { spaces: 0, ideographs: 0 };
  for (const seg of segments) {
    if (isTextSegment(seg)) {
      const found = justifyOpportunitiesOf(seg.text);
      total.spaces += found.spaces;
      total.ideographs += found.ideographs;
    }
  }
  return total;
}

/** Places one stretch of text can absorb slack. */
interface JustifyOpportunities {
  spaces: number;
  ideographs: number;
}

/**
 * The two counts for one string.
 *
 * The single definition of the rule. It was written twice — once to divide the
 * slack and once to add the resulting width back onto a run — sixteen lines apart,
 * and the two had to agree exactly or a line's measured width would stop matching
 * what was drawn. Widening {@link isJustifiableSpace} on one side only would have
 * done it.
 */
function justifyOpportunitiesOf(text: string): JustifyOpportunities {
  let spaces = 0;
  let ideographs = 0;
  for (const run of splitByScript(text)) {
    if (run.cjk) {
      // Glyph advances: `Tc` is added after every glyph shown. A combining mark the
      // face can draw is one of them, while a variation selector is folded into the
      // character before it and is not.
      ideographs += countGlyphAdvances(run.text);
      continue;
    }
    for (const char of run.text) {
      if (isJustifiableSpace(char)) {
        spaces++;
      }
    }
  }
  return { spaces, ideographs };
}

/**
 * Whether justification may widen this character as a word space.
 *
 * Only U+0020. PDF expresses word spacing with `Tw`, which the specification
 * applies to **byte 32 alone** — so slack assigned to U+3000 was added to the
 * line's measured width and then never drawn, leaving the layout convinced it had
 * filled a column it had not. An ideographic space is widened as a character
 * instead, through `Tc`, which does apply to it.
 */
function isJustifiableSpace(char: string): boolean {
  return char === " ";
}

/**
 * The width justification adds to one stretch of text.
 *
 * The same opportunities the slack was divided between, priced. Deriving both from
 * {@link justifyOpportunitiesOf} is what keeps the division and the pricing from
 * drifting apart.
 */
function justificationWidth(text: string, charSpacing: number, wordSpacing: number): number {
  if (charSpacing === 0 && wordSpacing === 0) {
    return 0;
  }
  const { spaces, ideographs } = justifyOpportunitiesOf(text);
  return spaces * wordSpacing + ideographs * charSpacing;
}

/**
 * Measure text with the run's own pair of typefaces, one stretch per script.
 *
 * A run naming `w:ascii="Calibri"` and `w:eastAsia="宋体"` draws `报表 Report`
 * with both, so measuring the whole string against either one is wrong. Uniform
 * text takes a single measurement, so Latin-only input costs nothing extra.
 */
function measureRunText(
  text: string,
  props: Run["properties"],
  fontSize: number,
  options: FullLayoutOptions | undefined,
  bold?: boolean,
  italic?: boolean
): number {
  const fonts = resolveRunFonts(props);
  if (fonts.ascii === fonts.eastAsia) {
    return stretchedRunWidth(
      measureLayoutText(text, fonts.ascii, fontSize, options, bold, italic),
      text,
      props
    );
  }
  let total = 0;
  for (const run of splitByScript(text)) {
    total += measureLayoutText(
      run.text,
      run.cjk ? fonts.eastAsia : fonts.ascii,
      fontSize,
      options,
      bold,
      italic
    );
  }
  return stretchedRunWidth(total, text, props);
}

/**
 * The extreme vertical metrics of every face a run draws with.
 *
 * Ascent and descent belong to a face, so a mixed run cannot sum them — but it
 * cannot pick one either. Choosing the East Asian face whenever the text contains
 * any CJK assumed it is always the taller, which is not true: a Latin face with a
 * large ascent inside a mostly-Chinese run was then measured against the CJK face
 * and its glyphs could be clipped by the line box. Each stretch is measured with
 * its own face and the extremes are taken.
 */
function measureRunMetrics(
  text: string,
  props: Run["properties"],
  fontSize: number,
  options: FullLayoutOptions | undefined,
  bold?: boolean,
  italic?: boolean
): { ascent: number; descent: number } {
  const fonts = resolveRunFonts(props);
  if (fonts.ascii === fonts.eastAsia) {
    return measureLayoutFontMetrics(text, fonts.ascii, fontSize, options, bold, italic);
  }
  let ascent = -Infinity;
  let descent = Infinity;
  for (const run of splitByScript(text)) {
    const m = measureLayoutFontMetrics(
      run.text,
      run.cjk ? fonts.eastAsia : fonts.ascii,
      fontSize,
      options,
      bold,
      italic
    );
    ascent = Math.max(ascent, m.ascent);
    descent = Math.min(descent, m.descent);
  }
  return ascent === -Infinity
    ? measureLayoutFontMetrics(text, fonts.ascii, fontSize, options, bold, italic)
    : { ascent, descent };
}

function measureLayoutFontMetrics(
  text: string,
  fontName: string,
  fontSize: number,
  options: FullLayoutOptions | undefined,
  bold?: boolean,
  italic?: boolean
): { ascent: number; descent: number } {
  if (options?.measureTextMetrics) {
    return options.measureTextMetrics(text, fontName, fontSize, bold, italic);
  }
  const extents = faceExtents(fontName, fontSize, bold);
  if (extents) {
    return extents;
  }
  const face = styledFontVariant(fontName, bold, italic);
  return { ascent: getFontAscent(face, fontSize), descent: getFontDescent(face, fontSize) };
}

function resolveColorHex(
  color: Run["properties"] extends { color?: infer C } ? C : unknown
): string | undefined {
  if (!color) {
    return undefined;
  }
  if (typeof color === "string") {
    return color;
  }
  // The `!color` check above already discarded `null`; an additional
  // `color !== null` test was always true and CodeQL flagged it as a
  // comparison between inconvertible types.
  if (typeof color === "object" && "value" in (color as object)) {
    return (color as { value: string }).value;
  }
  return undefined;
}

// =============================================================================
// Internal: Image / Geometry Helpers
// =============================================================================

function emuToPt(emu: number): number {
  return emu / EMU_PER_POINT;
}

/**
 * Convert the docx-internal `ImageMediaType` (`"png"`, `"jpeg"`, …)
 * to the standard MIME string consumers expect on
 * `LayoutImage.mimeType` / `PositionedInlineImage.mimeType`. Unknown
 * tags fall back to `application/octet-stream` so renderers can
 * decide whether to skip or draw a placeholder.
 */
function mediaTypeToMime(mt: string | undefined): string {
  switch (mt) {
    case "png":
      return "image/png";
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "tiff":
      return "image/tiff";
    case "svg":
      return "image/svg+xml";
    case "webp":
      return "image/webp";
    case "emf":
      return "image/x-emf";
    case "wmf":
      return "image/x-wmf";
    default:
      return "application/octet-stream";
  }
}

/**
 * Every image by each of its keys — `rId` and `aliasRIds` — since content in
 * a header, footer or note may name a picture by an alias.
 */
function buildImageMap(images: readonly ImageDef[] | undefined): ReadonlyMap<string, ImageDef> {
  return indexImagesByKey(images);
}

// =============================================================================
// Internal: FloatingImage / TextBox / Shape / Chart / SDT / Math / CheckBox
// =============================================================================

/**
 * Resolve the page-content-area position of a floating image per
 * ECMA-376 §20.4.2.10. Layout coordinates have origin at the top-left
 * of the **content area**; floating-image anchors are normally
 * expressed against the **page** or **margin**, so we translate
 * accordingly.
 *
 * Resolution order:
 *  1. `simplePos="1"` (we currently see only `simplePos.x`/`simplePos.y`
 *     in the model; we treat its presence as the simplePos override)
 *     — page-absolute EMU.
 *  2. `horizontalPosition` / `verticalPosition` with `align` keywords
 *     (left/center/right/inside/outside, top/center/bottom).
 *  3. `horizontalPosition` / `verticalPosition` with `offset` (EMU)
 *     relative to the chosen `relativeTo` reference.
 *  4. Fall back to the cursor (inline-like behaviour).
 *
 * `relativeTo` reference points (subset we resolve):
 *  - `"page"` — page top-left corner
 *  - `"margin"` — margin box top-left corner (same as content area
 *    origin in our coordinate system)
 *  - `"column"` / `"character"` / `"paragraph"` etc. — fall back to the
 *    cursor; reproducing them faithfully would require column/text-flow
 *    info we don't keep at this stage.
 */
function resolveFloatingImageRect(
  fi: FloatingImage,
  cursorY: number,
  contentWidth: number,
  contentHeight: number,
  geometry: PageGeometry,
  widthPt: number,
  heightPt: number
): { x: number; y: number } {
  const usingSimplePos = fi.simplePos !== undefined;

  // 1. simplePos: page-absolute. Translate into content-area coords by
  //    subtracting the page margins.
  if (usingSimplePos) {
    const pageX = emuToPt(fi.simplePos!.x ?? 0);
    const pageY = emuToPt(fi.simplePos!.y ?? 0);
    return {
      x: pageX - geometry.marginLeft,
      y: pageY - geometry.marginTop
    };
  }

  // 2/3. positionH / positionV
  const xPt = resolveHorizontal(fi, contentWidth, geometry, widthPt) ?? 0;
  const yPt = resolveVertical(fi, cursorY, contentHeight, geometry, heightPt) ?? cursorY;

  return { x: xPt, y: yPt };
}

function resolveHorizontal(
  fi: FloatingImage,
  contentWidth: number,
  geometry: PageGeometry,
  widthPt: number
): number | undefined {
  const h = fi.horizontalPosition;
  if (!h) {
    return undefined;
  }
  const relTo = h.relativeTo ?? "column";
  // Reference origin (in content-area coordinates) and width to anchor against.
  let originX = 0;
  let refWidth = contentWidth;
  if (relTo === "page") {
    originX = -geometry.marginLeft;
    refWidth = geometry.width;
  } else if (relTo === "margin" || relTo === "leftMargin" || relTo === "rightMargin") {
    originX = 0;
    refWidth = contentWidth;
  } // else: column/character/insideMargin/outsideMargin — fall back to content area

  if (h.align) {
    switch (h.align) {
      case "left":
      case "inside":
        return originX;
      case "right":
      case "outside":
        return originX + refWidth - widthPt;
      case "center":
        return originX + (refWidth - widthPt) / 2;
    }
  }
  if (h.offset != null) {
    return originX + emuToPt(h.offset);
  }
  return undefined;
}

function resolveVertical(
  fi: FloatingImage,
  cursorY: number,
  contentHeight: number,
  geometry: PageGeometry,
  heightPt: number
): number | undefined {
  const v = fi.verticalPosition;
  if (!v) {
    return undefined;
  }
  const relTo = v.relativeTo ?? "paragraph";
  let originY = cursorY;
  let refHeight = contentHeight;
  if (relTo === "page") {
    originY = -geometry.marginTop;
    refHeight = geometry.height;
  } else if (relTo === "margin" || relTo === "topMargin" || relTo === "bottomMargin") {
    originY = 0;
    refHeight = contentHeight;
  } // else paragraph/line/text-anchored — keep cursor as origin

  if (v.align) {
    switch (v.align) {
      case "top":
      case "inside":
        return originY;
      case "bottom":
      case "outside":
        return originY + refHeight - heightPt;
      case "center":
        return originY + (refHeight - heightPt) / 2;
    }
  }
  if (v.offset != null) {
    return originY + emuToPt(v.offset);
  }
  return undefined;
}

function layoutFloatingImage(
  fi: FloatingImage,
  cursorY: number,
  contentWidth: number,
  contentHeight: number,
  geometry: PageGeometry,
  sourceIndex: number,
  imageMap: ReadonlyMap<string, ImageDef>
): LayoutFloat {
  const widthPt = emuToPt(fi.width);
  const heightPt = emuToPt(fi.height);
  const { x: xPt, y: yPt } = resolveFloatingImageRect(
    fi,
    cursorY,
    contentWidth,
    contentHeight,
    geometry,
    widthPt,
    heightPt
  );

  const img = fi.rId ? imageMap.get(fi.rId) : undefined;
  const imageContent: LayoutImage = img
    ? {
        type: "image",
        rect: { x: xPt, y: yPt, width: widthPt, height: heightPt },
        data: img.data,
        mimeType: img.contentType ?? mediaTypeToMime(img.mediaType),
        altText: fi.altText,
        sourceIndex
      }
    : {
        type: "image",
        rect: { x: xPt, y: yPt, width: widthPt, height: heightPt },
        data: new Uint8Array(0),
        mimeType: "application/octet-stream",
        altText: fi.altText,
        sourceIndex
      };

  return {
    type: "float",
    rect: { x: xPt, y: yPt, width: widthPt, height: heightPt },
    content: imageContent,
    behindText: fi.behindDoc === true,
    ...(fi.wrap ? { wrap: convertWrap(fi.wrap) } : {}),
    sourceIndex
  };
}

/**
 * Translate the source `FloatingImage.wrap` (with `WrapMargins` in EMU)
 * into the layout-side `LayoutFloatWrap` (with margins already in
 * points so renderers don't need to know about EMU).
 */
function convertWrap(wrap: NonNullable<FloatingImage["wrap"]>): NonNullable<LayoutFloat["wrap"]> {
  const out: {
    -readonly [K in keyof NonNullable<LayoutFloat["wrap"]>]: NonNullable<LayoutFloat["wrap"]>[K];
  } = {
    style: wrap.style
  };
  if (wrap.side) {
    out.side = wrap.side;
  }
  if (wrap.margins) {
    const m: {
      -readonly [K in keyof NonNullable<NonNullable<LayoutFloat["wrap"]>["margins"]>]: number;
    } = {};
    if (wrap.margins.top != null) {
      m.top = emuToPt(wrap.margins.top);
    }
    if (wrap.margins.bottom != null) {
      m.bottom = emuToPt(wrap.margins.bottom);
    }
    if (wrap.margins.left != null) {
      m.left = emuToPt(wrap.margins.left);
    }
    if (wrap.margins.right != null) {
      m.right = emuToPt(wrap.margins.right);
    }
    if (Object.keys(m).length > 0) {
      out.margins = m;
    }
  }
  return out;
}

function layoutTextBox(
  tb: TextBox,
  startY: number,
  contentWidth: number,
  sourceIndex: number,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>
): LayoutTextBox {
  const widthPt = tb.width != null ? twipsToPt(tb.width) : contentWidth;
  // Lay out inner paragraphs against the text-box width; their positions
  // are returned relative to the box's top-left so renderers translate
  // by `rect.x`/`rect.y`.
  const inner: PageContent[] = [];
  let innerCursor = 0;
  for (const child of tb.content) {
    const laid = layoutParagraph(child, innerCursor, widthPt, options, undefined, imageMap);
    inner.push({ ...laid, sourceIndex });
    innerCursor = laid.rect.y + laid.rect.height;
  }

  const heightPt = tb.height != null ? twipsToPt(tb.height) : Math.max(innerCursor, 12);

  return {
    type: "textBox",
    rect: { x: 0, y: startY, width: widthPt, height: heightPt },
    content: inner,
    border:
      tb.stroke && tb.strokeColor
        ? { width: 0.75, color: normaliseHex(tb.strokeColor) }
        : undefined,
    background: tb.fill && tb.fillColor ? normaliseHex(tb.fillColor) : undefined,
    sourceIndex
  };
}

function layoutDrawingShape(
  shape: DrawingShape,
  startY: number,
  contentWidth: number,
  sourceIndex: number,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>
): LayoutShape {
  const widthPt = emuToPt(shape.width);
  const heightPt = emuToPt(shape.height);
  const innerWidth = Math.min(widthPt, contentWidth);

  const innerContent: PageContent[] = [];
  if (shape.textContent && shape.textContent.length > 0) {
    let cursor = 0;
    for (const para of shape.textContent) {
      const laid = layoutParagraph(para, cursor, innerWidth, options, undefined, imageMap);
      innerContent.push({ ...laid, sourceIndex });
      cursor = laid.rect.y + laid.rect.height;
    }
  }

  return {
    type: "shape",
    rect: { x: 0, y: startY, width: widthPt, height: heightPt },
    preset: shape.shapeType,
    fillColor: shape.noFill ? undefined : normaliseHexOrUndefined(shape.fillColor),
    strokeColor: shape.noOutline ? undefined : normaliseHexOrUndefined(shape.outlineColor),
    strokeWidth: shape.outlineWidth != null ? emuToPt(shape.outlineWidth) : undefined,
    textContent: innerContent.length > 0 ? innerContent : undefined,
    sourceIndex
  };
}

function layoutChart(
  ch: ChartContent | ChartExContent,
  startY: number,
  contentWidth: number,
  sourceIndex: number
): LayoutChart {
  // Source dimensions:
  //  - ChartContent stores width/height inside the inner Chart model
  //    (writer emits `<wp:extent>` from `chart.width/height`; reader
  //    populates them from the original drawing's `<wp:extent>`).
  //  - ChartExContent carries width/height directly on the content.
  // Both fall back to a 6"×3.5" default that matches Microsoft Word's
  // default insert size when the source supplied none.
  const widthEmu = ch.type === "chart" ? (ch.chart?.width ?? 5_486_400) : (ch.width ?? 5_486_400);
  const heightEmu =
    ch.type === "chart" ? (ch.chart?.height ?? 3_200_400) : (ch.height ?? 3_200_400);
  const widthPt = Math.min(emuToPt(widthEmu), contentWidth);
  const heightPt = emuToPt(heightEmu);
  const title = ch.type === "chart" ? (ch.chart?.title ?? ch.name) : ch.name;

  return {
    type: "chart",
    rect: { x: 0, y: startY, width: widthPt, height: heightPt },
    chartKind: ch.type === "chart" ? "chart" : "chartEx",
    title,
    altText: ch.altText,
    source: ch,
    sourceIndex
  };
}

function layoutSdt(
  sdt: StructuredDocumentTag,
  startY: number,
  contentWidth: number,
  sourceIndex: number,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>
): LayoutSdt {
  // SDT is a transparent flow container in layout terms: lay out its
  // children inline and report a rect that encloses them. Inline-only
  // children (bare runs) are skipped — the SDT-as-block contract is
  // what layout cares about.
  const inner: PageContent[] = [];
  let cursor = 0;
  for (const child of sdt.content) {
    if ("type" in child) {
      if (child.type === "paragraph") {
        const laid = layoutParagraph(child, cursor, contentWidth, options, undefined, imageMap);
        inner.push({ ...laid, sourceIndex });
        cursor = laid.rect.y + laid.rect.height;
      } else if (child.type === "table") {
        const laid = layoutTable(child, cursor, contentWidth, sourceIndex, options, imageMap);
        inner.push(laid);
        cursor = laid.rect.y + laid.rect.height;
      }
      // Run-only SDT children are not flowed at the block level.
    }
  }

  return {
    type: "sdt",
    rect: { x: 0, y: startY, width: contentWidth, height: cursor },
    content: inner,
    tag: sdt.properties?.tag,
    alias: sdt.properties?.alias,
    sourceIndex
  };
}

function layoutMath(
  mb: MathBlock,
  startY: number,
  contentWidth: number,
  sourceIndex: number,
  options: FullLayoutOptions | undefined
): LayoutMath {
  const text = extractMathText(mb.content);
  let mathML: string | undefined;
  try {
    mathML = ommlToMathML(mb.content);
  } catch {
    mathML = undefined;
  }
  const fontSize = DEFAULT_FONT_SIZE_PT;
  const lineHeight = fontSize * 1.2;
  // Width is approximated from the plain-text fallback so renderers that
  // don't handle MathML still see a reasonable bounding box.
  const fontName = options?.fonts?.get("Cambria Math") ?? "Cambria Math";
  const widthPt = Math.min(measureLayoutText(text, fontName, fontSize, options), contentWidth);

  return {
    type: "math",
    rect: { x: 0, y: startY, width: widthPt, height: lineHeight },
    text,
    mathML,
    sourceIndex
  };
}

function layoutCheckBox(
  cb: CheckBox,
  startY: number,
  sourceIndex: number,
  options: FullLayoutOptions | undefined
): LayoutCheckBox {
  const fontSize = DEFAULT_FONT_SIZE_PT;
  const checked = cb.checked === true;
  const glyph = checked
    ? (cb.checkedState?.value ?? "\u2611") // ☑
    : (cb.uncheckedState?.value ?? "\u2610"); // ☐
  const fontName = cb.checkedState?.font ?? options?.fonts?.get("MS Gothic") ?? "MS Gothic";
  const widthPt = measureLayoutText(glyph, fontName, fontSize, options);
  return {
    type: "checkBox",
    rect: { x: 0, y: startY, width: widthPt, height: fontSize * 1.2 },
    checked,
    glyph,
    fontSize,
    sourceIndex
  };
}

const tocParagraphCache = new WeakMap<TableOfContents, readonly Paragraph[]>();

/**
 * The paragraphs a table of contents is written as: its entries, and — unless
 * they hold its field themselves — the empty paragraph its field ends in (see
 * `renderTableOfContents`), which takes a line like any other.
 */
function tocParagraphs(toc: TableOfContents): readonly Paragraph[] {
  let paragraphs = tocParagraphCache.get(toc);
  if (!paragraphs) {
    const entries = toc.cachedParagraphs ?? [];
    paragraphs =
      toc.field === "entries" || toc.field === "none" || entries.length === 0
        ? entries
        : [...entries, { type: "paragraph", children: [] }];
    tocParagraphCache.set(toc, paragraphs);
  }
  return paragraphs;
}

function layoutTableOfContents(
  toc: TableOfContents,
  startY: number,
  contentWidth: number,
  sourceIndex: number,
  options: FullLayoutOptions | undefined,
  imageMap: ReadonlyMap<string, ImageDef>
): LayoutTableOfContents {
  const entries: LayoutParagraph[] = [];
  let cursor = 0;
  const paragraphs = tocParagraphs(toc);
  if (paragraphs.length > 0) {
    for (const para of paragraphs) {
      const laid = layoutParagraph(para, cursor, contentWidth, options, undefined, imageMap);
      entries.push({ ...laid, sourceIndex });
      cursor = laid.rect.y + laid.rect.height;
    }
  } else {
    // Stub: emit a single placeholder paragraph so renderers always have
    // something to render. Consumers wanting a real TOC should run
    // `updateTableOfContents()` before layout.
    const stub: Paragraph = {
      type: "paragraph",
      children: [{ content: [{ type: "text", text: "[Table of Contents]" }] }]
    };
    const laid = layoutParagraph(stub, 0, contentWidth, options, undefined, imageMap);
    entries.push({ ...laid, sourceIndex });
    cursor = laid.rect.height;
  }

  return {
    type: "tableOfContents",
    rect: { x: 0, y: startY, width: contentWidth, height: cursor },
    entries,
    sourceIndex
  };
}

function layoutAltChunk(
  ac: AltChunk,
  startY: number,
  contentWidth: number,
  sourceIndex: number
): LayoutAltChunk {
  // Layout cannot interpret HTML / RTF / MHT payloads; reserve a
  // placeholder rect proportional to a small fixed height so renderers
  // can show a substitution glyph or run their own foreign-content
  // pipeline.
  const heightPt = DEFAULT_FONT_SIZE_PT * 3;
  return {
    type: "altChunk",
    rect: { x: 0, y: startY, width: contentWidth, height: heightPt },
    contentType: ac.contentType ?? "application/octet-stream",
    fileName: ac.fileName,
    sourceIndex
  };
}

function layoutOpaqueDrawing(
  od: OpaqueDrawing,
  startY: number,
  contentWidth: number,
  sourceIndex: number
): LayoutOpaqueDrawing {
  // We have no idea how big the drawing is from XML alone; reserve a
  // square-ish placeholder roughly matching a typical chart slot. High-
  // fidelity renderers can re-parse `rawXml` if they need exact size.
  const heightPt = DEFAULT_FONT_SIZE_PT * 12;
  return {
    type: "opaqueDrawing",
    rect: { x: 0, y: startY, width: contentWidth, height: heightPt },
    rawXml: od.rawXml,
    sourceIndex
  };
}

function normaliseHex(hex: string): string {
  return hex.startsWith("#") ? hex.slice(1) : hex;
}

function normaliseHexOrUndefined(hex: string | undefined): string | undefined {
  return hex ? normaliseHex(hex) : undefined;
}
