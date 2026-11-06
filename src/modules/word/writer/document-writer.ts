/**
 * DOCX Writers - Document Body
 *
 * Renders the main word/document.xml part including all block-level
 * content types: paragraphs, tables, floating images, TOC, math,
 * text boxes, and structured document tags.
 */

import {
  DOCUMENT_NAMESPACES,
  STD_DOC_ATTRIBUTES,
  NS_A,
  NS_C_CHART,
  NS_CX_CHART,
  NS_WPS
} from "@word/constants";
import { atNode, unwritableNode, within } from "@word/core/node-path";
import { DocxRawXmlPolicyError, DocxWriteError } from "@word/errors";
import type {
  DocxDocument,
  BodyContent,
  StructuredDocumentTag,
  CheckBox,
  DrawingFrame,
  DrawingShape,
  OpaqueDrawing,
  ChartContent,
  ChartExContent,
  FloatingImage
} from "@word/types";
import {
  DEFAULT_CHART_HEIGHT_EMU,
  DEFAULT_CHART_EX_HEIGHT_EMU,
  DEFAULT_CHART_WIDTH_EMU
} from "@word/units";
import { checkWritableChart } from "@word/writer/chart-checks";
import { renderCheckBox } from "@word/writer/checkbox-writer";
import { closeDrawingFrame, frameModel, openDrawingFrame } from "@word/writer/drawing-frame";
import { renderFloatingImage, renderFloatingImageRun } from "@word/writer/image-writer";
import { renderMathBlock } from "@word/writer/math-writer";
import { renderParagraph } from "@word/writer/paragraph-writer";
import { renderDocumentBackground } from "@word/writer/parts-writer";
import { renderRangeMarkers } from "@word/writer/range-marker-writer";
import type { RenderHelpers, DocxRenderContext } from "@word/writer/render-context";
import {
  createRenderContext,
  enterRevisionIds,
  exitRevisionIds,
  prepareRawMarkup,
  renderHelpersFor
} from "@word/writer/render-context";
import { renderRun } from "@word/writer/run-writer";
import { renderSdtPr } from "@word/writer/sdt-writer";
import { renderSectionProperties } from "@word/writer/section-writer";
import { renderShapeBody } from "@word/writer/shape-markup";
import { renderTable } from "@word/writer/table-writer";
import { renderTextBox } from "@word/writer/textbox-writer";
import { renderTableOfContents } from "@word/writer/toc-writer";
import type { XmlSink } from "@xml/types";

/**
 * Render a structured document tag.
 *
 * `ctx` may be `undefined` when called from a sub-renderer that only knows
 * about `RenderHelpers` (e.g. table cell, header/footer). In that case we
 * still emit a syntactically valid `<w:sdt>`, but we cannot auto-assign an
 * id from the document-wide id generator — callers that care about
 * uniqueness should provide `properties.id` explicitly.
 */
export function renderSdt(
  xml: XmlSink,
  sdt: StructuredDocumentTag,
  ctx: DocxRenderContext | undefined,
  helpers?: RenderHelpers
): void {
  xml.openNode("w:sdt");
  renderSdtPr(xml, sdt, ctx);

  xml.openNode("w:sdtContent");
  // Use the explicit `helpers` argument when present (it carries the
  // rId/policy state for the surrounding part), and otherwise derive a
  // helpers object from the document-level ctx.
  const sdtHelpers: RenderHelpers | undefined = helpers
    ? helpers
    : ctx
      ? renderHelpersFor(ctx)
      : undefined;
  for (const child of sdt.content) {
    if ("type" in child) {
      if (child.type === "paragraph") {
        renderParagraph(xml, child, sdtHelpers);
      } else if (child.type === "table") {
        renderTable(xml, child, sdtHelpers);
      } else if (child.type === "sdt") {
        // Nested SDT (e.g. items inside a repeating section).
        renderSdt(xml, child, ctx, sdtHelpers);
      }
    } else {
      // No `type` discriminator → it's a bare Run. Some readers (and the
      // parser in this module) preserve a w:r directly inside w:sdtContent
      // for run-level structured document tags. Render the run directly so
      // round-tripping does not silently drop its text.
      renderRangeMarkers(xml, child.rangeMarkersBefore, sdtHelpers);
      renderRun(xml, child, sdtHelpers);
    }
  }
  renderRangeMarkers(xml, sdt.contentRangeMarkersAfter, sdtHelpers);
  xml.closeNode();

  xml.closeNode();
}

/** A floating image whose anchor belongs among the preceding paragraph's text. */
function isSharedFloat(item: BodyContent | undefined): boolean {
  return item?.type === "floatingImage" && item.inPrecedingParagraph === true;
}

/** Render a single body content element (paragraph, table, image, etc.). */
export function renderBodyContent(
  xml: XmlSink,
  content: BodyContent,
  ctx?: DocxRenderContext
): void {
  // Without a context the content is rendered on its own, or as part of
  // whatever write is in effect: its revision ids follow that write's map.
  if (!ctx) {
    renderBodyContentIn(xml, content, createRenderContext());
    return;
  }
  const previous = enterRevisionIds(ctx.revisionIds);
  try {
    renderBodyContentIn(xml, content, ctx);
  } finally {
    exitRevisionIds(previous);
  }
}

function renderBodyContentIn(
  xml: XmlSink,
  content: BodyContent,
  renderCtx: DocxRenderContext
): void {
  if (typeof content !== "object" || content === null) {
    unwritableNode("", "a body block", content);
  }
  const helpers = renderHelpersFor(renderCtx);
  // Range markers between this block and the previous one. A paragraph and a
  // table write their own, as they are also rendered outside the body.
  if (content.type !== "paragraph" && content.type !== "table") {
    renderRangeMarkers(xml, content.rangeMarkersBefore, helpers);
  }
  switch (content.type) {
    case "paragraph":
      renderParagraph(xml, content, helpers);
      break;
    case "table":
      renderTable(xml, content, helpers);
      break;
    case "floatingImage":
      renderFloatingImage(xml, content, renderHelpersFor(renderCtx));
      break;
    case "tableOfContents":
      renderTableOfContents(xml, content, helpers);
      break;
    case "math":
      renderMathBlock(xml, content);
      break;
    case "textBox":
      renderTextBox(xml, content, helpers);
      break;
    case "sdt":
      renderSdt(xml, content, renderCtx);
      break;
    case "checkBox":
      // A checkbox renders as an inline (run-level) SDT whose sdtContent holds
      // a run. At block level that run would be an illegal child of
      // CT_SdtContentBlock, so wrap it in a paragraph — making it a valid
      // run-level SDT inside a block-level paragraph.
      xml.openNode("w:p");
      renderCheckBox(xml, content as CheckBox);
      xml.closeNode();
      break;
    case "drawingShape":
      renderDrawingShape(xml, content as DrawingShape, renderCtx);
      break;
    case "opaqueDrawing":
      renderOpaqueDrawing(xml, content, renderCtx);
      break;
    case "chart":
      renderChartDrawing(xml, content, renderCtx);
      break;
    case "chartEx":
      renderChartExDrawing(xml, content, renderCtx);
      break;
    case "altChunk": {
      // Prefer the rId registered on the render context (so we never have
      // to mutate `content.rId` on the caller's model). Fall back to the
      // model's own rId for callers who pre-assigned one (e.g. when
      // re-rendering an already-packaged document).
      const altRId = renderCtx.altChunkRIds.get(content) ?? content.rId;
      if (altRId) {
        xml.leafNode("w:altChunk", { "r:id": altRId });
      }
      break;
    }
    default:
      unwritableNode("type", "a body block kind", (content as { type?: unknown }).type);
  }
}

/** Render the complete word/document.xml. */
export function renderDocument(xml: XmlSink, doc: DocxDocument, ctx?: DocxRenderContext): void {
  // If no ctx provided (e.g. minimal renders), create a fresh one so id
  // counters are local to this call rather than module-global.
  const renderCtx = ctx ?? createRenderContext();
  const previous = enterRevisionIds(renderCtx.revisionIds);
  try {
    renderDocumentIn(xml, doc, renderCtx);
  } finally {
    exitRevisionIds(previous);
  }
}

function renderDocumentIn(xml: XmlSink, doc: DocxDocument, renderCtx: DocxRenderContext): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:document", DOCUMENT_NAMESPACES);

  // Document background (must come before w:body)
  if (doc.background) {
    renderDocumentBackground(xml, doc.background);
  }

  xml.openNode("w:body");

  // Word / OOXML require a `<w:p>` to separate two adjacent `<w:tbl>` blocks
  // (per ECMA-376 §17.13.5.34: a table must be followed by a paragraph or
  // section break before the next table can begin). When a model contains
  // back-to-back tables we synthesise an empty paragraph between them so
  // the output remains valid in Word.
  // Likewise the last `EG_BlockLevelElts` of a body must be a paragraph
  // (or `sectPr`); if the model body is empty we emit a single `<w:p>` so
  // Word does not see `<w:body/>`.
  let prevWasTable = false;
  let bodyHadAnything = false;
  for (let i = 0; i < doc.body.length; i++) {
    const content = doc.body[i]!;
    try {
      const isTable = content.type === "table";
      if (isTable && prevWasTable) {
        xml.openNode("w:p");
        xml.closeNode();
      }
      // Floating images read from among a paragraph's text go back into it.
      let shared = 0;
      while (content.type === "paragraph" && isSharedFloat(doc.body[i + 1 + shared])) {
        shared++;
      }
      if (shared > 0 && content.type === "paragraph") {
        const floats = doc.body.slice(i + 1, i + 1 + shared) as FloatingImage[];
        const helpers = renderHelpersFor(renderCtx);
        const previous = enterRevisionIds(renderCtx.revisionIds);
        try {
          renderParagraph(xml, content, helpers, sink => {
            for (const float of floats) {
              renderRangeMarkers(sink, float.rangeMarkersBefore, helpers);
              renderFloatingImageRun(sink, float, helpers);
            }
          });
        } finally {
          exitRevisionIds(previous);
        }
        i += shared;
      } else {
        renderBodyContent(xml, content, renderCtx);
      }
      prevWasTable = isTable;
      bodyHadAnything = true;
    } catch (error) {
      throw atNode(`body[${i}]`, error);
    }
  }

  // Range markers after the last block close the body, before its sectPr.
  renderRangeMarkers(xml, doc.bodyRangeMarkersAfter, renderHelpersFor(renderCtx));

  // If the model produced no body content at all, write one empty
  // paragraph so the body satisfies CT_Body's "ends with a paragraph or
  // sectPr" rule.
  if (!bodyHadAnything) {
    xml.openNode("w:p");
    xml.closeNode();
  }

  // Final section properties. If the caller didn't supply any we still
  // write a default one (US Letter, 1" margins) — CT_Body must terminate
  // with a `<w:sectPr>` so Word knows the page geometry. Without this
  // some Word builds open the document but render every paragraph at
  // page-zero size, while LibreOffice silently rejects the package.
  if (doc.sectionProperties) {
    try {
      renderSectionProperties(xml, doc.sectionProperties);
    } catch (error) {
      throw atNode("sectionProperties", error);
    }
  } else {
    renderSectionProperties(xml, {
      pageSize: { width: 12240, height: 15840 },
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
    });
  }

  xml.closeNode(); // w:body
  xml.closeNode(); // w:document
}

// =============================================================================
// DrawingML Shape Writer
// =============================================================================

function renderDrawingShape(xml: XmlSink, shape: DrawingShape, ctx: DocxRenderContext): void {
  // Wrap in w:p > w:r > w:drawing > wp:anchor > a:graphic > a:graphicData > wps:wsp
  // Note: shape.rawXml (if present) carries advanced DrawingML fragments that
  // belong inside wps:spPr (gradient/pattern fills, effect lists, etc.). It is
  // NOT a substitute for the structural wrappers — earlier behaviour wrote it
  // verbatim into w:body which produced invalid OOXML.
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.openNode("w:drawing");

  const drawingId = ctx.ids.nextDocPrId();
  const helpers = renderHelpersFor(ctx);
  const frame = frameModel("anchor", "shape", shape, "Shape");
  const then = shape._frame ? frameModel("anchor", "shape", shape._frame.read, "Shape") : undefined;
  openDrawingFrame(xml, frame, shape._frame, then, drawingId, helpers);

  xml.openNode("a:graphic", { "xmlns:a": NS_A });
  xml.openNode("a:graphicData", { uri: NS_WPS });

  renderShapeBody(xml, shape, drawingId, ctx);
  xml.closeNode(); // a:graphicData
  xml.closeNode(); // a:graphic
  closeDrawingFrame(xml, frame, shape._frame, helpers); // wp:anchor
  xml.closeNode(); // w:drawing
  xml.closeNode(); // w:r
  xml.closeNode(); // w:p
}

// =============================================================================
// Opaque Drawing Writer
// =============================================================================

function renderOpaqueDrawing(xml: XmlSink, drawing: OpaqueDrawing, ctx: DocxRenderContext): void {
  if (ctx.rawXmlPolicy === "reject") {
    throw new DocxRawXmlPolicyError("opaqueDrawing");
  }
  if (ctx.rawXmlPolicy === "strip") {
    // Emit a structurally-valid empty paragraph instead of the opaque drawing.
    xml.openNode("w:p");
    xml.closeNode();
    return;
  }
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.writeRaw(prepareRawMarkup(drawing.rawXml, renderHelpersFor(ctx)));
  xml.closeNode(); // w:r
  xml.closeNode(); // w:p
}

// =============================================================================
// Chart Drawing Writer (inline chart reference)
// =============================================================================

/** Inline drawing options shared between chart (`c:`) and chartEx (`cx:`). */
interface InlineChartOptions {
  /** Width in EMU. */
  readonly cx: number;
  /** Height in EMU. */
  readonly cy: number;
  /** Display name on the wp:docPr element. */
  readonly name: string;
  /** Optional alt text. */
  readonly altText?: string;
  /** GraphicData URI (chart vs chartEx). */
  readonly graphicDataUri: string;
  /** Leaf element local name (`c:chart` or `cx:chart`). */
  readonly leafName: string;
  /** XML namespace declared on the leaf element. */
  readonly leafXmlnsAttr: string;
  /** Namespace value for the leaf's xmlns attribute. */
  readonly leafXmlnsValue: string;
  /** The relationship ID pointing at the chart part. */
  readonly rId: string;
  /** The frame the chart was read with. */
  readonly frame?: DrawingFrame;
}

function renderInlineChartDrawing(
  xml: XmlSink,
  ctx: DocxRenderContext,
  opts: InlineChartOptions
): void {
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.openNode("w:drawing");
  // A chart read anchored stays anchored; one built is inline.
  const kind = opts.frame?.kind ?? "inline";
  const fields = { width: opts.cx, height: opts.cy, name: opts.name, altText: opts.altText };
  const frame = frameModel(kind, "chart", fields, "Chart");
  const then = opts.frame ? frameModel(kind, "chart", opts.frame.read, "Chart") : undefined;
  const helpers = renderHelpersFor(ctx);
  openDrawingFrame(xml, frame, opts.frame, then, ctx.ids.nextDocPrId(), helpers);
  xml.openNode("a:graphic", {
    "xmlns:a": NS_A
  });
  xml.openNode("a:graphicData", { uri: opts.graphicDataUri });
  xml.leafNode(opts.leafName, {
    [opts.leafXmlnsAttr]: opts.leafXmlnsValue,
    "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "r:id": opts.rId
  });
  xml.closeNode(); // a:graphicData
  xml.closeNode(); // a:graphic
  closeDrawingFrame(xml, frame, opts.frame, helpers); // wp:inline / wp:anchor
  xml.closeNode(); // w:drawing
  xml.closeNode(); // w:r
  xml.closeNode(); // w:p
}

function renderChartDrawing(xml: XmlSink, content: ChartContent, ctx: DocxRenderContext): void {
  // The chart part is written apart from the body; its model checks run here,
  // where the chart's place in the document is known.
  within("chart", () => checkWritableChart(content.chart));
  const rId = ctx.chartRIds.get(content);
  if (!rId) {
    throw new DocxWriteError(
      "Chart content was not registered with a relationship id. " +
        "This usually means the chart was rendered outside of packageDocx() — " +
        "use packageDocx() or pre-populate ctx.chartRIds before rendering."
    );
  }
  renderInlineChartDrawing(xml, ctx, {
    cx: content.chart.width ?? DEFAULT_CHART_WIDTH_EMU,
    cy: content.chart.height ?? DEFAULT_CHART_HEIGHT_EMU,
    name: content.name ?? "Chart",
    altText: content.altText,
    graphicDataUri: NS_C_CHART,
    leafName: "c:chart",
    leafXmlnsAttr: "xmlns:c",
    leafXmlnsValue: NS_C_CHART,
    rId,
    frame: content._frame
  });
}

function renderChartExDrawing(xml: XmlSink, content: ChartExContent, ctx: DocxRenderContext): void {
  const rId = ctx.chartRIds.get(content);
  if (!rId) {
    throw new DocxWriteError(
      "ChartEx content was not registered with a relationship id. " +
        "This usually means the chart was rendered outside of packageDocx() — " +
        "use packageDocx() or pre-populate ctx.chartRIds before rendering."
    );
  }
  renderInlineChartDrawing(xml, ctx, {
    cx: content.width ?? DEFAULT_CHART_WIDTH_EMU,
    cy: content.height ?? DEFAULT_CHART_EX_HEIGHT_EMU,
    name: content.name ?? "ChartEx",
    altText: content.altText,
    graphicDataUri: NS_CX_CHART,
    leafName: "cx:chart",
    leafXmlnsAttr: "xmlns:cx",
    leafXmlnsValue: NS_CX_CHART,
    rId,
    frame: content._frame
  });
}
