/**
 * DOCX Writers - Paragraph Properties & Paragraph
 *
 * Renders w:pPr and w:p elements including numbering, tabs, borders, etc.
 */

import { WRITTEN, writtenToken } from "@word/core/argument-checks";
import {
  PARAGRAPH_ONLY_PROPERTIES,
  type ParagraphPropertiesScope,
  NARROWER_PROPERTIES,
  PARAGRAPH_MARK_REVISIONS
} from "@word/core/change-records";
import { refuseAbsentFields, unwritableNode, within, writableList } from "@word/core/node-path";
import { hyperlinkContent, isRangeMarker, isRun, isSpanMarker } from "@word/core/text-utils";
import { DocxRawXmlPolicyError } from "@word/errors";
import type {
  ParagraphProperties,
  Paragraph,
  ParagraphChild,
  RangeMarker,
  Hyperlink,
  Border,
  ParagraphBorders,
  TabStop,
  CommentReference,
  InsertedRun,
  DeletedRun,
  MovedFromRun,
  MovedToRun,
  RunContent,
  ParagraphFrame,
  Run
} from "@word/types";
import type { OpenSpans } from "@word/writer/range-marker-writer";
import {
  closeOpenSpans,
  renderRangeMarker,
  renderRangeMarkers
} from "@word/writer/range-marker-writer";
import type { RenderHelpers } from "@word/writer/render-context";
import { prepareRawMarkup, revisionAttributes } from "@word/writer/render-context";
import {
  renderBorderElement,
  renderRun,
  renderField,
  renderRunProperties,
  renderRunPropertiesContents,
  renderShading
} from "@word/writer/run-writer";
import { emitInSchemaOrder, type OrderedChild } from "@word/writer/schema-order";
import { renderSectionProperties } from "@word/writer/section-writer";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";
import { XmlWriter } from "@xml/writer";

/** Render a single border element. */
export function renderBorder(xml: XmlSink, tagName: string, border: Border): void {
  renderBorderElement(xml, tagName, border);
}

/** Render paragraph borders. */
function renderParagraphBorders(xml: XmlSink, borders: ParagraphBorders): void {
  xml.openNode("w:pBdr");
  if (borders.top) {
    renderBorder(xml, "w:top", borders.top);
  }
  if (borders.left) {
    renderBorder(xml, "w:left", borders.left);
  }
  if (borders.bottom) {
    renderBorder(xml, "w:bottom", borders.bottom);
  }
  if (borders.right) {
    renderBorder(xml, "w:right", borders.right);
  }
  if (borders.between) {
    renderBorder(xml, "w:between", borders.between);
  }
  if (borders.bar) {
    renderBorder(xml, "w:bar", borders.bar);
  }
  xml.closeNode();
}

/** Render tab stops. */
function renderTabs(xml: XmlSink, tabs: readonly TabStop[]): void {
  xml.openNode("w:tabs");
  tabs.forEach((tab, i) => {
    const attrs: Record<string, string> = {
      "w:val": writtenToken(tab, "type", WRITTEN.tabType, `[${i}].type`) ?? "left",
      "w:pos": xmlValue(tab.position, "w:pos")
    };
    const leader = writtenToken(tab, "leader", WRITTEN.tabLeader, `[${i}].leader`);
    if (leader !== undefined) {
      attrs["w:leader"] = leader;
    }
    xml.leafNode("w:tab", attrs);
  });
  xml.closeNode();
}

/** Render paragraph frame properties (w:framePr). */
function renderFrame(xml: XmlSink, frame: ParagraphFrame): void {
  const attrs: Record<string, string> = {};
  const dropCap = writtenToken(frame, "dropCap", WRITTEN.dropCap, "dropCap");
  if (dropCap !== undefined) {
    attrs["w:dropCap"] = dropCap;
  }
  if (frame.lines !== undefined) {
    attrs["w:lines"] = xmlValue(frame.lines, "w:lines");
  }
  if (frame.width !== undefined) {
    attrs["w:w"] = xmlValue(frame.width, "w:w");
  }
  if (frame.height !== undefined) {
    attrs["w:h"] = xmlValue(frame.height, "w:h");
  }
  if (frame.hSpace !== undefined) {
    attrs["w:hSpace"] = xmlValue(frame.hSpace, "w:hSpace");
  }
  if (frame.vSpace !== undefined) {
    attrs["w:vSpace"] = xmlValue(frame.vSpace, "w:vSpace");
  }
  const wrap = writtenToken(frame, "wrap", WRITTEN.frameWrap, "wrap");
  if (wrap !== undefined) {
    attrs["w:wrap"] = wrap;
  }
  const hAnchor = writtenToken(frame, "hAnchor", WRITTEN.frameAnchor, "hAnchor");
  if (hAnchor !== undefined) {
    attrs["w:hAnchor"] = hAnchor;
  }
  const vAnchor = writtenToken(frame, "vAnchor", WRITTEN.frameAnchor, "vAnchor");
  if (vAnchor !== undefined) {
    attrs["w:vAnchor"] = vAnchor;
  }
  if (frame.x !== undefined) {
    attrs["w:x"] = xmlValue(frame.x, "w:x");
  }
  const xAlign = writtenToken(frame, "xAlign", WRITTEN.frameXAlign, "xAlign");
  if (xAlign !== undefined) {
    attrs["w:xAlign"] = xAlign;
  }
  if (frame.y !== undefined) {
    attrs["w:y"] = xmlValue(frame.y, "w:y");
  }
  const yAlign = writtenToken(frame, "yAlign", WRITTEN.frameYAlign, "yAlign");
  if (yAlign !== undefined) {
    attrs["w:yAlign"] = yAlign;
  }
  xml.leafNode("w:framePr", attrs);
}

/**
 * Render w:pPr (paragraph properties), its children in `CT_PPr` order (`emitInSchemaOrder`). A style's, a default's or a
 * numbering level's (`general`) and a change record's previous ones (`previous`) are narrower types without the mark
 * or the section, and the previous ones without a change record: one set there is refused, not dropped.
 */
export function renderParagraphProperties(
  xml: XmlSink,
  pPr: ParagraphProperties,
  scope: ParagraphPropertiesScope = "paragraph"
): void {
  if (scope === "general") {
    refuseAbsentFields(pPr, PARAGRAPH_ONLY_PROPERTIES, NARROWER_PROPERTIES.generalParagraph);
  } else if (scope === "previous") {
    refuseAbsentFields(
      pPr,
      [...PARAGRAPH_ONLY_PROPERTIES, "propertyChange"],
      NARROWER_PROPERTIES.previousParagraph
    );
  }
  xml.openNode("w:pPr");

  const onOff = (name: string, value: boolean | undefined): OrderedChild =>
    value !== undefined && [name, () => xml.leafNode(name, !value ? { "w:val": "0" } : undefined)];
  const flag = (name: string, value: boolean | undefined): OrderedChild =>
    !!value && [name, () => xml.leafNode(name)];

  const alignment = writtenToken(pPr, "alignment", WRITTEN.alignment, "alignment");
  const textAlignment = writtenToken(pPr, "textAlignment", WRITTEN.textAlignment, "textAlignment");
  const textDirection = writtenToken(pPr, "textDirection", WRITTEN.textDirection, "textDirection");
  const hasMarkRevision = PARAGRAPH_MARK_REVISIONS.some(({ field }) => pPr[field] !== undefined);

  emitInSchemaOrder("w:CT_PPr", [
    !!pPr.style && ["w:pStyle", () => xml.leafNode("w:pStyle", { "w:val": pPr.style! })],
    flag("w:keepNext", pPr.keepNext),
    flag("w:keepLines", pPr.keepLines),
    flag("w:pageBreakBefore", pPr.pageBreakBefore),
    !!pPr.frame && ["w:framePr", () => within("frame", () => renderFrame(xml, pPr.frame!))],
    onOff("w:widowControl", pPr.widowControl),
    !!pPr.numbering && [
      "w:numPr",
      () => {
        xml.openNode("w:numPr");
        xml.leafNode("w:ilvl", { "w:val": xmlValue(pPr.numbering!.level ?? 0, "w:val") });
        xml.leafNode("w:numId", { "w:val": xmlValue(pPr.numbering!.numId, "w:val") });
        xml.closeNode();
      }
    ],
    flag("w:suppressLineNumbers", pPr.suppressLineNumbers),
    !!pPr.borders && [
      "w:pBdr",
      () => within("borders", () => renderParagraphBorders(xml, pPr.borders!))
    ],
    // A thematic break is drawn as a bottom border. Only synthesise one when the
    // paragraph has no borders of its own: `w:pPr` permits a single `w:pBdr`, and
    // emitting a second produced invalid XML whose duplicate element made readers
    // drop the paragraph's real border settings.
    !!pPr.thematicBreak &&
      !pPr.borders && [
        "w:pBdr",
        () => {
          xml.openNode("w:pBdr");
          xml.leafNode("w:bottom", {
            "w:val": "single",
            "w:sz": "6",
            "w:space": "1",
            "w:color": "auto"
          });
          xml.closeNode();
        }
      ],
    !!pPr.shading && ["w:shd", () => within("shading", () => renderShading(xml, pPr.shading!))],
    !!pPr.tabs &&
      pPr.tabs.length > 0 && ["w:tabs", () => within("tabs", () => renderTabs(xml, pPr.tabs!))],
    flag("w:suppressAutoHyphens", pPr.suppressAutoHyphens),
    onOff("w:kinsoku", pPr.kinsoku),
    onOff("w:wordWrap", pPr.wordWrap),
    onOff("w:overflowPunct", pPr.overflowPunctuation),
    onOff("w:topLinePunct", pPr.topLinePunctuation),
    onOff("w:autoSpaceDE", pPr.autoSpaceEastAsianText),
    onOff("w:autoSpaceDN", pPr.autoSpaceEastAsianDigit),
    flag("w:bidi", pPr.bidi),
    onOff("w:snapToGrid", pPr.snapToGrid),
    !!pPr.spacing && ["w:spacing", () => renderSpacing(xml, pPr.spacing!)],
    !!pPr.indent && ["w:ind", () => renderIndent(xml, pPr.indent!)],
    flag("w:contextualSpacing", pPr.contextualSpacing),
    flag("w:mirrorIndents", pPr.mirrorIndents),
    alignment !== undefined && ["w:jc", () => xml.leafNode("w:jc", { "w:val": alignment })],
    textDirection !== undefined && [
      "w:textDirection",
      () => xml.leafNode("w:textDirection", { "w:val": textDirection })
    ],
    textAlignment !== undefined && [
      "w:textAlignment",
      () => xml.leafNode("w:textAlignment", { "w:val": textAlignment })
    ],
    pPr.outlineLevel !== undefined && [
      "w:outlineLvl",
      () => xml.leafNode("w:outlineLvl", { "w:val": xmlValue(pPr.outlineLevel!, "w:val") })
    ],
    !!pPr.cnfStyle && ["w:cnfStyle", () => xml.leafNode("w:cnfStyle", { "w:val": pPr.cnfStyle! })],
    // Paragraph mark properties (w:rPr inside w:pPr).
    //
    // ECMA-376 17.3.1.27 (CT_PPrBase) only allows ONE `w:rPr` child here. We
    // therefore merge the mark's revisions (`paragraphInsertion`, `paragraphDeletion`,
    // `paragraphMoveFrom`, `paragraphMoveTo`) into the same `w:rPr` as
    // `markRunProperties`. Inside this single w:rPr, the schema requires
    // w:ins/w:del/w:moveFrom/w:moveTo to come BEFORE the run-property children
    // (per CT_ParaRPr).
    (hasMarkRevision || !!pPr.markRunProperties) && [
      "w:rPr",
      () => renderParagraphMarkProperties(xml, pPr)
    ],
    !!pPr.sectionProperties && [
      "w:sectPr",
      () => within("sectionProperties", () => renderSectionProperties(xml, pPr.sectionProperties!))
    ],
    // Paragraph property change revision — NOT recursed into when rendering previousProperties
    scope !== "previous" &&
      !!pPr.propertyChange && ["w:pPrChange", () => renderParagraphPropertyChange(xml, pPr)]
  ]);

  xml.closeNode();
}

function renderSpacing(xml: XmlSink, spacing: NonNullable<ParagraphProperties["spacing"]>): void {
  const attrs: Record<string, string> = {};
  if (spacing.before !== undefined) {
    attrs["w:before"] = xmlValue(spacing.before, "w:before");
  }
  if (spacing.after !== undefined) {
    attrs["w:after"] = xmlValue(spacing.after, "w:after");
  }
  if (spacing.line !== undefined) {
    attrs["w:line"] = xmlValue(spacing.line, "w:line");
  }
  const lineRule = writtenToken(spacing, "lineRule", WRITTEN.lineRule, "spacing.lineRule");
  if (lineRule !== undefined) {
    attrs["w:lineRule"] = lineRule;
  }
  if (spacing.beforeAutoSpacing) {
    attrs["w:beforeAutospacing"] = "1";
  }
  if (spacing.afterAutoSpacing) {
    attrs["w:afterAutospacing"] = "1";
  }
  xml.leafNode("w:spacing", attrs);
}

function renderIndent(xml: XmlSink, indent: NonNullable<ParagraphProperties["indent"]>): void {
  const attrs: Record<string, string> = {};
  if (indent.left !== undefined) {
    attrs["w:left"] = xmlValue(indent.left, "w:left");
  }
  if (indent.right !== undefined) {
    attrs["w:right"] = xmlValue(indent.right, "w:right");
  }
  if (indent.hanging !== undefined) {
    attrs["w:hanging"] = xmlValue(indent.hanging, "w:hanging");
  }
  if (indent.firstLine !== undefined) {
    attrs["w:firstLine"] = xmlValue(indent.firstLine, "w:firstLine");
  }
  if (indent.start !== undefined) {
    attrs["w:start"] = xmlValue(indent.start, "w:start");
  }
  if (indent.end !== undefined) {
    attrs["w:end"] = xmlValue(indent.end, "w:end");
  }
  xml.leafNode("w:ind", attrs);
}

function renderParagraphPropertyChange(xml: XmlSink, pPr: ParagraphProperties): void {
  const change = pPr.propertyChange!;
  xml.openNode("w:pPrChange", revisionAttributes(change.revision));
  if (change.previousProperties) {
    const previous = change.previousProperties;
    within("propertyChange.previousProperties", () =>
      renderParagraphProperties(xml, previous, "previous")
    );
  } else {
    // `CT_PPrChange` requires its `w:pPr`: no previous properties is an empty one.
    xml.openNode("w:pPr");
    xml.closeNode();
  }
  xml.closeNode();
}

function renderParagraphMarkProperties(xml: XmlSink, pPr: ParagraphProperties): void {
  xml.openNode("w:rPr");
  for (const { element, field } of PARAGRAPH_MARK_REVISIONS) {
    const revision = pPr[field];
    if (revision) {
      xml.leafNode(`w:${element}`, revisionAttributes(revision));
    }
  }
  if (pPr.markRunProperties) {
    const mark = pPr.markRunProperties;
    within("markRunProperties", () =>
      renderRunPropertiesContents(xml, mark, false, { mark: true })
    );
  }
  xml.closeNode(); // w:rPr
}

/** Render a hyperlink. */
function renderHyperlink(xml: XmlSink, link: Hyperlink, helpers?: RenderHelpers): void {
  const attrs: Record<string, string> = {};
  // Resolve the relationship id: prefer the packager-provided WeakMap (so we
  // never need the model to carry the rId), fall back to whatever the model
  // already had (e.g. round-tripped DOCX where the reader populated rId).
  // Resolve r:id. Prefer the packager-assigned rId from the helpers map
  // (it's the only one guaranteed to exist in the rebuilt .rels file).
  // Fall back to the model's `rId` only when no packager-assigned id is
  // present — this happens for purely model-internal references such as
  // anchor-only hyperlinks.
  const resolvedRId = helpers?.hyperlinkRIds?.get(link) ?? (link.url ? undefined : link.rId);
  if (resolvedRId) {
    attrs["r:id"] = resolvedRId;
  }
  if (link.anchor) {
    attrs["w:anchor"] = link.anchor;
  }
  // CT_Hyperlink (ECMA-376 §17.16.22) requires r:id OR w:anchor — when
  // neither is present, emit the children as bare runs to avoid a schema
  // violation that some validators flag.
  if (!resolvedRId && !link.anchor) {
    renderHyperlinkContent(xml, link, helpers);
    return;
  }
  if (link.tooltip) {
    attrs["w:tooltip"] = link.tooltip;
  }
  if (link.history) {
    attrs["w:history"] = "1";
  }
  if (link.tgtFrame) {
    attrs["w:tgtFrame"] = link.tgtFrame;
  }
  if (link.docLocation) {
    attrs["w:docLocation"] = link.docLocation;
  }
  xml.openNode("w:hyperlink", attrs);
  renderHyperlinkContent(xml, link, helpers);
  xml.closeNode();
}

/** The model path of the item at an index of a rendered child list. */
type ChildPath = (index: number) => string;

/** A paragraph's (or a plain hyperlink's) children: the list is the model's own. */
const childrenPath: ChildPath = i => `children[${i}]`;

/**
 * Where each item of {@link hyperlinkContent} lives in the hyperlink: a run in
 * `children`, a nested child at `nestedChildren[n].child` — the order
 * `hyperlinkContent` merges them in, so an index of the merged list names the
 * item it holds.
 */
function hyperlinkContentPaths(link: Hyperlink): ChildPath {
  const nested = link.nestedChildren;
  if (!nested || nested.length === 0) {
    return childrenPath;
  }
  const paths: string[] = [];
  let next = 0;
  link.children.forEach((_, i) => {
    while (next < nested.length && nested[next]!.before <= i) {
      paths.push(`nestedChildren[${next++}].child`);
    }
    paths.push(`children[${i}]`);
  });
  while (next < nested.length) {
    paths.push(`nestedChildren[${next++}].child`);
  }
  return i => paths[i]!;
}

/** The runs of a hyperlink with its nested children among them, then the markers after them. */
function renderHyperlinkContent(xml: XmlSink, link: Hyperlink, helpers?: RenderHelpers): void {
  renderParagraphChildren(
    xml,
    hyperlinkContent(link),
    link.rangeMarkersAfter,
    helpers,
    hyperlinkContentPaths(link)
  );
}

/** Render a comment reference as a run of its own, with the mark's run formatting. */
function renderCommentReference(xml: XmlSink, cr: CommentReference): void {
  xml.openNode("w:r");
  if (cr.properties) {
    const properties = cr.properties;
    within("properties", () => renderRunProperties(xml, properties));
  }
  xml.leafNode("w:commentReference", { "w:id": xmlValue(cr.id, "w:id") });
  xml.closeNode();
}

/** A run under revision, which is written inside a revision wrapper. */
type TrackedRun = InsertedRun | DeletedRun | MovedFromRun | MovedToRun;

/** The revision wrapper element `child` is written in, if it is a {@link TrackedRun}. */
function revisionWrapperOf(child: ParagraphChild): string | undefined {
  if (!("type" in child)) {
    return undefined;
  }
  switch (child.type) {
    case "insertedRun":
      return "w:ins";
    case "deletedRun":
      return "w:del";
    case "movedFromRun":
      return "w:moveFrom";
    case "movedToRun":
      return "w:moveTo";
    default:
      return undefined;
  }
}

/** Whether `child` is a tracked run of the same kind and revision as `run`. */
function isSameRevision(run: TrackedRun, child: ParagraphChild): boolean {
  if (!("type" in child) || child.type !== run.type) {
    return false;
  }
  const a = run.revision;
  const b = (child as TrackedRun).revision;
  return a === b || (a.id === b.id && a.author === b.author && a.date === b.date);
}

/** Whether markers can be written inside a revision wrapper: none opens a wrapper of its own. */
function canNestInRevision(markers: readonly RangeMarker[] | undefined): boolean {
  return markers === undefined || !markers.some(isSpanMarker);
}

/**
 * Render a deleted run's `w:r`: its text as `w:delText`. A field in it is
 * written as runs of its own, its instruction as `w:delInstrText`.
 */
function renderDeletedRunContent(xml: XmlSink, run: Run, helpers?: RenderHelpers): void {
  // A field writes the run's properties into runs of its own; checked here, a
  // bad one is named at `properties` rather than inside the field.
  if (run.properties) {
    const properties = run.properties;
    within("properties", () => renderRunProperties(new XmlWriter(), properties));
  }
  let start = 0;
  for (let i = 0; i <= run.content.length; i++) {
    const content = run.content[i];
    if (i < run.content.length && content!.type !== "field") {
      continue;
    }
    if (i > start) {
      renderDeletedRunText(xml, run, run.content.slice(start, i));
    }
    if (content?.type === "field") {
      renderField(xml, content, run.properties, helpers, true);
    }
    start = i + 1;
  }
  if (run.content.length === 0) {
    renderDeletedRunText(xml, run, []);
  }
}

function renderDeletedRunText(xml: XmlSink, run: Run, contents: readonly RunContent[]): void {
  xml.openNode("w:r");
  if (run.properties) {
    const properties = run.properties;
    within("properties", () => renderRunProperties(xml, properties));
  }
  for (const content of contents) {
    if (content.type === "text") {
      xml.openNode(content.fieldCode ? "w:delInstrText" : "w:delText", { "xml:space": "preserve" });
      xml.writeText(content.text);
      xml.closeNode();
    } else if (content.type === "break") {
      const brAttrs: Record<string, string> = {};
      if (content.breakType) {
        brAttrs["w:type"] = content.breakType;
      }
      xml.leafNode("w:br", Object.keys(brAttrs).length > 0 ? brAttrs : undefined);
    } else if (content.type === "tab") {
      xml.leafNode("w:tab");
    } else if (content.type === "carriageReturn") {
      xml.leafNode("w:cr");
    } else if (content.type === "noBreakHyphen") {
      xml.leafNode("w:noBreakHyphen");
    } else if (content.type === "softHyphen") {
      xml.leafNode("w:softHyphen");
    }
  }
  xml.closeNode();
}

/**
 * Render the tracked run at `children[start]` in its revision wrapper (track
 * changes), together with every later child the same wrapper holds, and
 * return the index of the first child it does not.
 *
 * The model keeps one tracked run per run, so a revision around several runs
 * is several tracked runs sharing its revision. They are written as the one
 * wrapper they came from: following tracked runs of the same kind and
 * revision join it, as do range markers between two of them — a marker has no
 * revision of its own, and `CT_RunTrackChange` admits range markup. Whatever
 * else separates two pieces of a revision leaves them in wrappers of their
 * own, and every wrapper's id is drawn from the package's revision map
 * ({@link writtenRevisionId}), because a revision id is an annotation
 * identifier, unique within the document (ECMA-376 §17.13.1); author and date
 * are kept.
 */
function renderTrackedRuns(
  xml: XmlSink,
  children: readonly ParagraphChild[],
  start: number,
  wrapper: string,
  helpers: RenderHelpers | undefined,
  pathOf: ChildPath
): number {
  const first = children[start] as TrackedRun;
  let end = start + 1;
  for (let k = start + 1; k < children.length; k++) {
    const child = children[k]!;
    if (!canNestInRevision(child.rangeMarkersBefore)) {
      break;
    }
    if (isSameRevision(first, child)) {
      end = k + 1;
    } else if (!(isRangeMarker(child) && !isSpanMarker(child))) {
      break;
    }
  }
  xml.openNode(wrapper, revisionAttributes(first.revision));
  for (let k = start; k < end; k++) {
    const child = children[k]!;
    within(pathOf(k), () => {
      if (k > start) {
        renderRangeMarkers(xml, child.rangeMarkersBefore, helpers);
      }
      if (isRangeMarker(child)) {
        renderRangeMarker(xml, child, helpers);
      } else if (first.type === "deletedRun") {
        const run = (child as DeletedRun).run;
        within("run", () => renderDeletedRunContent(xml, run, helpers));
      } else {
        const run = (child as TrackedRun).run;
        within("run", () => renderRun(xml, run, helpers));
      }
    });
  }
  xml.closeNode();
  return end;
}

/** Render a paragraph child element other than a tracked run (see {@link renderTrackedRuns}). */
function renderParagraphChild(
  xml: XmlSink,
  child: ParagraphChild,
  helpers?: RenderHelpers,
  spans?: OpenSpans
): void {
  // Markers that are not paragraph children of their own (range permissions,
  // span markers) ride on the child that follows them.
  renderRangeMarkers(xml, child.rangeMarkersBefore, helpers, spans);
  if ("type" in child) {
    switch (child.type) {
      case "bookmarkStart":
      case "bookmarkEnd":
        renderRangeMarker(xml, child, helpers);
        return;
      case "hyperlink":
        renderHyperlink(xml, child, helpers);
        return;
      case "commentReference":
        renderCommentReference(xml, child);
        return;
      case "opaqueParagraphChild": {
        const policy = helpers?.rawXmlPolicy;
        if (policy === "reject") {
          throw new DocxRawXmlPolicyError("opaqueParagraphChild");
        }
        if (policy !== "strip") {
          xml.writeRaw(prepareRawMarkup(child.rawXml, helpers));
        }
        return;
      }
      default:
        if (isRangeMarker(child)) {
          renderRangeMarker(xml, child, helpers);
          return;
        }
        unwritableNode("type", "a paragraph child kind", (child as { type?: unknown }).type);
    }
  }
  if (isRun(child)) {
    renderRun(xml, child, helpers);
    return;
  }
  // A run is the one child without a `type`; its content is a list.
  writableList((child as { content?: readonly unknown[] }).content, "content");
}

/**
 * Render a w:p element. `leading` writes run-level markup the model does not
 * hold ahead of the paragraph's own children — a header's watermark.
 */
export function renderParagraph(
  xml: XmlSink,
  para: Paragraph,
  helpers?: RenderHelpers,
  leading?: (xml: XmlSink) => void
): void {
  renderRangeMarkers(xml, para.rangeMarkersBefore, helpers);
  const pAttrs: Record<string, string> = {};
  if (para.paraId) {
    pAttrs["w14:paraId"] = para.paraId;
  }
  if (para.textId) {
    pAttrs["w14:textId"] = para.textId;
  }
  xml.openNode("w:p", Object.keys(pAttrs).length > 0 ? pAttrs : undefined);

  if (para.properties) {
    const properties = para.properties;
    within("properties", () => renderParagraphProperties(xml, properties));
  }
  leading?.(xml);

  renderParagraphChildren(
    xml,
    writableList(para.children, "children") as readonly ParagraphChild[],
    para.rangeMarkersAfter,
    helpers
  );

  xml.closeNode();
}

/**
 * Render a sequence of paragraph children, then `after`. Span markers among
 * the range markers open and close their wrapper elements where they sit;
 * wrappers still open when the sequence ends are closed there, so the output
 * stays well-formed whatever editing left the pairs in.
 */
function renderParagraphChildren(
  xml: XmlSink,
  children: readonly ParagraphChild[],
  after: readonly RangeMarker[] | undefined,
  helpers?: RenderHelpers,
  pathOf: ChildPath = childrenPath
): void {
  let spans: OpenSpans | undefined;
  for (let i = 0; i < children.length;) {
    const child = children[i]!;
    const at = i;
    const wrapper = within(pathOf(at), () => {
      if (typeof child !== "object" || child === null) {
        unwritableNode("", "a paragraph child", child);
      }
      // Most children carry no markers; only those that do need the span stack.
      if (child.rangeMarkersBefore !== undefined) {
        spans ??= [];
      }
      const kind = revisionWrapperOf(child);
      if (kind === undefined) {
        renderParagraphChild(xml, child, helpers, spans);
      } else {
        renderRangeMarkers(xml, child.rangeMarkersBefore, helpers, spans);
      }
      return kind;
    });
    // A tracked-run wrapper writes several children; each names its own index.
    i =
      wrapper === undefined ? i + 1 : renderTrackedRuns(xml, children, i, wrapper, helpers, pathOf);
  }
  if (after !== undefined) {
    spans ??= [];
    renderRangeMarkers(xml, after, helpers, spans);
  }
  if (spans !== undefined) {
    closeOpenSpans(xml, spans);
  }
}
