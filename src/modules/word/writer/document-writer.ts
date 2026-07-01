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
import { DocxRawXmlPolicyError, DocxWriteError } from "@word/errors";
import type {
  DocxDocument,
  BodyContent,
  StructuredDocumentTag,
  CheckBox,
  DrawingShape,
  OpaqueDrawing,
  ChartContent,
  ChartExContent
} from "@word/types";
import {
  DEFAULT_CHART_HEIGHT_EMU,
  DEFAULT_CHART_EX_HEIGHT_EMU,
  DEFAULT_CHART_WIDTH_EMU,
  DEFAULT_RELATIVE_HEIGHT,
  DEFAULT_WRAP_MARGIN_EMU
} from "@word/units";
import { renderCheckBox } from "@word/writer/checkbox-writer";
import { renderFloatingImage } from "@word/writer/image-writer";
import { renderMathBlock } from "@word/writer/math-writer";
import { renderParagraph } from "@word/writer/paragraph-writer";
import { renderDocumentBackground } from "@word/writer/parts-writer";
import type { RenderHelpers, WordRenderContext } from "@word/writer/render-context";
import { createRenderContext } from "@word/writer/render-context";
import { renderRun } from "@word/writer/run-writer";
import { renderSdtPr } from "@word/writer/sdt-writer";
import { renderSectionProperties } from "@word/writer/section-writer";
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
  ctx: WordRenderContext | undefined,
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
      ? {
          imageRemap: ctx.imageRIdRemap,
          hyperlinkRIds: ctx.hyperlinkRIds,
          nextDocPrId: ctx.ids.nextDocPrId,
          rawXmlPolicy: ctx.rawXmlPolicy
        }
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
      renderRun(xml, child, sdtHelpers);
    }
  }
  xml.closeNode();

  xml.closeNode();
}

/** Render a single body content element (paragraph, table, image, etc.). */
export function renderBodyContent(
  xml: XmlSink,
  content: BodyContent,
  ctx?: WordRenderContext
): void {
  // Ensure id counters and helpers are local to this call when no ctx
  // was provided, so callers without a shared context don't accidentally
  // share module-global state across unrelated render passes.
  const renderCtx = ctx ?? createRenderContext();
  const helpers: RenderHelpers = {
    imageRemap: renderCtx.imageRIdRemap,
    hyperlinkRIds: renderCtx.hyperlinkRIds,
    nextDocPrId: renderCtx.ids.nextDocPrId,
    rawXmlPolicy: renderCtx.rawXmlPolicy
  };
  switch (content.type) {
    case "paragraph":
      renderParagraph(xml, content, helpers);
      break;
    case "table":
      renderTable(xml, content, helpers);
      break;
    case "floatingImage":
      renderFloatingImage(xml, content, renderCtx.imageRIdRemap, renderCtx.ids.nextDocPrId);
      break;
    case "tableOfContents":
      renderTableOfContents(xml, content);
      break;
    case "math":
      renderMathBlock(xml, content);
      break;
    case "textBox":
      renderTextBox(xml, content);
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
  }
}

/** Render the complete word/document.xml. */
export function renderDocument(xml: XmlSink, doc: DocxDocument, ctx?: WordRenderContext): void {
  // If no ctx provided (e.g. minimal renders), create a fresh one so id
  // counters are local to this call rather than module-global.
  const renderCtx = ctx ?? createRenderContext();

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
  for (const content of doc.body) {
    const isTable = content.type === "table";
    if (isTable && prevWasTable) {
      xml.openNode("w:p");
      xml.closeNode();
    }
    renderBodyContent(xml, content, renderCtx);
    prevWasTable = isTable;
    bodyHadAnything = true;
  }

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
    renderSectionProperties(xml, doc.sectionProperties);
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

function renderDrawingShape(xml: XmlSink, shape: DrawingShape, ctx: WordRenderContext): void {
  // Wrap in w:p > w:r > w:drawing > wp:anchor > a:graphic > a:graphicData > wps:wsp
  // Note: shape.rawXml (if present) carries advanced DrawingML fragments that
  // belong inside wps:spPr (gradient/pattern fills, effect lists, etc.). It is
  // NOT a substitute for the structural wrappers — earlier behaviour wrote it
  // verbatim into w:body which produced invalid OOXML.
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.openNode("w:drawing");

  const anchorAttrs: Record<string, string> = {
    distT: "0",
    distB: "0",
    distL: String(DEFAULT_WRAP_MARGIN_EMU),
    distR: String(DEFAULT_WRAP_MARGIN_EMU),
    simplePos: "0",
    relativeHeight: String(DEFAULT_RELATIVE_HEIGHT),
    behindDoc: shape.behindDoc ? "1" : "0",
    locked: "0",
    layoutInCell: "1",
    allowOverlap: "1"
  };
  xml.openNode("wp:anchor", anchorAttrs);

  xml.leafNode("wp:simplePos", { x: "0", y: "0" });

  // Horizontal position
  const hp = shape.horizontalPosition;
  xml.openNode("wp:positionH", {
    relativeFrom: hp?.relativeTo ?? "column"
  });
  if (hp?.align) {
    xml.openNode("wp:align");
    xml.writeText(hp.align);
    xml.closeNode();
  } else {
    xml.openNode("wp:posOffset");
    xml.writeText(String(hp?.offset ?? 0));
    xml.closeNode();
  }
  xml.closeNode(); // wp:positionH

  // Vertical position
  const vp = shape.verticalPosition;
  xml.openNode("wp:positionV", {
    relativeFrom: vp?.relativeTo ?? "paragraph"
  });
  if (vp?.align) {
    xml.openNode("wp:align");
    xml.writeText(vp.align);
    xml.closeNode();
  } else {
    xml.openNode("wp:posOffset");
    xml.writeText(String(vp?.offset ?? 0));
    xml.closeNode();
  }
  xml.closeNode(); // wp:positionV

  xml.leafNode("wp:extent", {
    cx: String(shape.width),
    cy: String(shape.height)
  });
  xml.leafNode("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" });

  // Wrap
  const wrap = shape.wrap;
  if (!wrap || wrap.style === "none") {
    xml.leafNode("wp:wrapNone");
  } else if (wrap.style === "square") {
    xml.leafNode("wp:wrapSquare", { wrapText: wrap.side ?? "bothSides" });
  } else if (wrap.style === "tight") {
    xml.leafNode("wp:wrapTight", { wrapText: wrap.side ?? "bothSides" });
  } else if (wrap.style === "topAndBottom") {
    xml.leafNode("wp:wrapTopAndBottom");
  }

  const drawingId = ctx.ids.nextDocPrId();
  const docPrAttrs: Record<string, string> = {
    id: String(drawingId),
    name: shape.name ?? "Shape"
  };
  if (shape.altText) {
    docPrAttrs["descr"] = shape.altText;
  }
  xml.leafNode("wp:docPr", docPrAttrs);

  // wp:cNvGraphicFramePr is optional in the schema but Word and
  // LibreOffice expect it on every wp:anchor; without it some readers
  // refuse to load the drawing. Empty content is sufficient.
  xml.leafNode("wp:cNvGraphicFramePr");

  xml.openNode("a:graphic", { "xmlns:a": NS_A });
  xml.openNode("a:graphicData", { uri: NS_WPS });

  xml.openNode("wps:wsp");

  // Non-visual properties — required by the wordprocessingShape schema
  // (CT_WordprocessingShape.cNvPr + cNvSpPr). The id must match
  // wp:docPr/@id so Word treats them as the same logical object.
  xml.leafNode("wps:cNvPr", {
    id: String(drawingId),
    name: shape.name ?? "Shape"
  });
  xml.leafNode("wps:cNvSpPr");

  // Shape properties
  xml.openNode("wps:spPr");

  // Transform
  const xfrmAttrs: Record<string, string> = {};
  if (shape.rotation) {
    xfrmAttrs["rot"] = String(shape.rotation);
  }
  if (shape.flipHorizontal) {
    xfrmAttrs["flipH"] = "1";
  }
  if (shape.flipVertical) {
    xfrmAttrs["flipV"] = "1";
  }
  xml.openNode("a:xfrm", Object.keys(xfrmAttrs).length > 0 ? xfrmAttrs : {});
  xml.leafNode("a:off", { x: "0", y: "0" });
  xml.leafNode("a:ext", { cx: String(shape.width), cy: String(shape.height) });
  xml.closeNode(); // a:xfrm

  // Preset geometry
  xml.openNode("a:prstGeom", { prst: shape.shapeType });
  xml.leafNode("a:avLst");
  xml.closeNode(); // a:prstGeom

  // Fill (basic + advanced). The OOXML schema requires fill children of
  // spPr to appear before a:ln, hence advanced gradient/pattern fills are
  // injected here rather than alongside effects.
  if (shape.noFill) {
    xml.leafNode("a:noFill");
  } else if (shape.fillColor) {
    xml.openNode("a:solidFill");
    xml.leafNode("a:srgbClr", { val: shape.fillColor });
    xml.closeNode();
  }
  if (shape._advancedFillXml) {
    if (ctx.rawXmlPolicy === "reject") {
      throw new DocxRawXmlPolicyError("drawingShape._advancedFillXml");
    }
    if (ctx.rawXmlPolicy !== "strip") {
      xml.writeRaw(shape._advancedFillXml);
    }
  }

  // Outline
  const lnAttrs: Record<string, string> = {};
  if (shape.outlineWidth) {
    lnAttrs["w"] = String(shape.outlineWidth);
  }
  if (shape.outlineColor || shape.noOutline) {
    xml.openNode("a:ln", lnAttrs);
    if (shape.noOutline) {
      xml.leafNode("a:noFill");
    } else if (shape.outlineColor) {
      xml.openNode("a:solidFill");
      xml.leafNode("a:srgbClr", { val: shape.outlineColor });
      xml.closeNode();
    }
    xml.closeNode(); // a:ln
  }

  // Advanced effects (a:effectLst, a:scene3d, a:sp3d) follow a:ln per the
  // OOXML schema. They were previously interleaved with fill via a single
  // rawXml string which produced documents that violated the schema order.
  if (shape._advancedEffectsXml) {
    if (ctx.rawXmlPolicy === "reject") {
      throw new DocxRawXmlPolicyError("drawingShape._advancedEffectsXml");
    }
    if (ctx.rawXmlPolicy !== "strip") {
      xml.writeRaw(shape._advancedEffectsXml);
    }
  } else if (shape.rawXml && !shape._advancedFillXml) {
    // Backwards-compat: if the caller provided an opaque rawXml without the
    // split fields (e.g. round-tripped from a reader), drop it after a:ln.
    // This preserves the previous "rawXml → spPr tail" behaviour for shapes
    // that don't go through createShape().
    if (ctx.rawXmlPolicy === "reject") {
      throw new DocxRawXmlPolicyError("drawingShape.rawXml");
    }
    if (ctx.rawXmlPolicy !== "strip") {
      xml.writeRaw(shape.rawXml);
    }
  }

  xml.closeNode(); // wps:spPr

  // Text body
  if (shape.textContent && shape.textContent.length > 0) {
    xml.openNode("wps:txbx");
    xml.openNode("w:txbxContent");
    const txbxHelpers: RenderHelpers | undefined = ctx
      ? {
          imageRemap: ctx.imageRIdRemap,
          hyperlinkRIds: ctx.hyperlinkRIds,
          nextDocPrId: ctx.ids.nextDocPrId,
          rawXmlPolicy: ctx.rawXmlPolicy
        }
      : undefined;
    for (const para of shape.textContent) {
      renderParagraph(xml, para, txbxHelpers);
    }
    xml.closeNode(); // w:txbxContent
    xml.closeNode(); // wps:txbx
  }

  // Body properties (required). The vertical text anchor lives on a:bodyPr/@anchor.
  if (shape.textBodyAnchor) {
    xml.leafNode("wps:bodyPr", { anchor: shape.textBodyAnchor });
  } else {
    xml.leafNode("wps:bodyPr");
  }

  xml.closeNode(); // wps:wsp
  xml.closeNode(); // a:graphicData
  xml.closeNode(); // a:graphic
  xml.closeNode(); // wp:anchor
  xml.closeNode(); // w:drawing
  xml.closeNode(); // w:r
  xml.closeNode(); // w:p
}

// =============================================================================
// Opaque Drawing Writer
// =============================================================================

function renderOpaqueDrawing(xml: XmlSink, drawing: OpaqueDrawing, ctx: WordRenderContext): void {
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
  xml.writeRaw(drawing.rawXml);
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
}

function renderInlineChartDrawing(
  xml: XmlSink,
  ctx: WordRenderContext,
  opts: InlineChartOptions
): void {
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.openNode("w:drawing");
  xml.openNode("wp:inline", { distT: "0", distB: "0", distL: "0", distR: "0" });
  xml.leafNode("wp:extent", { cx: String(opts.cx), cy: String(opts.cy) });
  xml.leafNode("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" });
  xml.leafNode("wp:docPr", {
    id: String(ctx.ids.nextDocPrId()),
    name: opts.name,
    ...(opts.altText ? { descr: opts.altText } : {})
  });
  xml.leafNode("wp:cNvGraphicFramePr");
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
  xml.closeNode(); // wp:inline
  xml.closeNode(); // w:drawing
  xml.closeNode(); // w:r
  xml.closeNode(); // w:p
}

function renderChartDrawing(xml: XmlSink, content: ChartContent, ctx: WordRenderContext): void {
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
    rId
  });
}

function renderChartExDrawing(xml: XmlSink, content: ChartExContent, ctx: WordRenderContext): void {
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
    rId
  });
}
