/**
 * DOCX Writers - Document Body
 *
 * Renders the main word/document.xml part including all block-level
 * content types: paragraphs, tables, floating images, TOC, math,
 * text boxes, and structured document tags.
 */

import type { XmlSink } from "@xml/types";

import { DOCUMENT_NAMESPACES, STD_DOC_ATTRIBUTES, NS_A, NS_WPS } from "../constants";
import { DocxWriteError } from "../errors";
import type {
  DocxDocument,
  BodyContent,
  StructuredDocumentTag,
  CheckBox,
  DrawingShape,
  OpaqueDrawing,
  ChartContent,
  ChartExContent
} from "../types";
import {
  DEFAULT_CHART_HEIGHT_EMU,
  DEFAULT_CHART_EX_HEIGHT_EMU,
  DEFAULT_CHART_WIDTH_EMU,
  DEFAULT_RELATIVE_HEIGHT,
  DEFAULT_WRAP_MARGIN_EMU
} from "../units";
import { renderCheckBox } from "./checkbox-writer";
import { renderFloatingImage } from "./image-writer";
import { renderMathBlock } from "./math-writer";
import { renderParagraph } from "./paragraph-writer";
import { renderDocumentBackground } from "./parts-writer";
import { createRenderContext, type RenderHelpers, type WordRenderContext } from "./render-context";
import { renderSectionProperties } from "./section-writer";
import { renderTable } from "./table-writer";
import { renderTextBox } from "./textbox-writer";
import { renderTableOfContents } from "./toc-writer";

/** Render a structured document tag. */
function renderSdt(xml: XmlSink, sdt: StructuredDocumentTag, ctx: WordRenderContext): void {
  xml.openNode("w:sdt");

  if (sdt.properties) {
    xml.openNode("w:sdtPr");
    const p = sdt.properties;
    // Auto-assign ID if not provided (Word strongly expects unique IDs)
    const sdtId = p.id ?? ctx.ids.nextSdtId();
    xml.leafNode("w:id", { "w:val": String(sdtId) });
    if (p.tag) {
      xml.leafNode("w:tag", { "w:val": p.tag });
    }
    if (p.alias) {
      xml.leafNode("w:alias", { "w:val": p.alias });
    }
    // Lock: combine lockContent + lockSdt
    if (p.lockContent && p.lockSdt) {
      xml.leafNode("w:lock", { "w:val": "sdtContentLocked" });
    } else if (p.lockContent) {
      xml.leafNode("w:lock", { "w:val": "contentLocked" });
    } else if (p.lockSdt) {
      xml.leafNode("w:lock", { "w:val": "sdtLocked" });
    }
    // Appearance (w15:appearance for modern SDTs)
    if (p.appearance) {
      xml.leafNode("w15:appearance", { "w15:val": p.appearance });
    } else if (p.showAs) {
      // Back-compat mapping of legacy showAs
      const mapped =
        p.showAs === "boundingBox" ? "boundingBox" : p.showAs === "startEnd" ? "tags" : "hidden";
      xml.leafNode("w15:appearance", { "w15:val": mapped });
    }
    // Show placeholder toggle
    if (p.showingPlaceholder) {
      xml.leafNode("w:showingPlcHdr");
    }
    if (p.placeholder) {
      xml.openNode("w:placeholder");
      xml.leafNode("w:docPart", { "w:val": p.placeholder });
      xml.closeNode();
    }
    if (p.temporary) {
      xml.leafNode("w:temporary");
    }
    if (p.dataBinding) {
      const bindAttrs: Record<string, string> = {
        "w:xpath": p.dataBinding.xpath,
        "w:storeItemID": p.dataBinding.storeItemId
      };
      if (p.dataBinding.prefixMappings) {
        bindAttrs["w:prefixMappings"] = p.dataBinding.prefixMappings;
      }
      xml.leafNode("w:dataBinding", bindAttrs);
    }
    // Type discriminator markers
    if (p.plainText) {
      xml.leafNode("w:text");
    }
    if (p.richText) {
      xml.leafNode("w:richText");
    }
    if (p.picture) {
      xml.leafNode("w:picture");
    }
    if (p.group) {
      xml.leafNode("w:group");
    }
    if (p.equation) {
      xml.leafNode("w:equation");
    }
    if (p.citation) {
      xml.leafNode("w:citation");
    }
    if (p.bibliography) {
      xml.leafNode("w:bibliography");
    }
    if (p.repeatingSectionItem) {
      xml.leafNode("w15:repeatingSectionItem");
    }
    if (p.repeatingSection) {
      // w15:repeatingSection has child elements (not attributes) per the schema
      const rs = p.repeatingSection;
      const hasChildren = rs.sectionTitle !== undefined || rs.allowInsertDelete !== undefined;
      if (hasChildren) {
        xml.openNode("w15:repeatingSection");
        if (rs.sectionTitle !== undefined) {
          xml.leafNode("w15:sectionTitle", { "w15:val": rs.sectionTitle });
        }
        if (rs.allowInsertDelete === false) {
          xml.leafNode("w15:doNotAllowInsertDeleteSection");
        }
        xml.closeNode();
      } else {
        xml.leafNode("w15:repeatingSection");
      }
    }
    // Checkbox (w14 extension)
    if (p.checkbox) {
      xml.openNode("w14:checkbox");
      xml.leafNode("w14:checked", p.checkbox.checked ? { "w14:val": "1" } : { "w14:val": "0" });
      if (p.checkbox.checkedChar || p.checkbox.checkedFont) {
        const cAttrs: Record<string, string> = {};
        if (p.checkbox.checkedChar) {
          cAttrs["w14:val"] = p.checkbox.checkedChar;
        }
        if (p.checkbox.checkedFont) {
          cAttrs["w14:font"] = p.checkbox.checkedFont;
        }
        xml.leafNode("w14:checkedState", cAttrs);
      }
      if (p.checkbox.uncheckedChar || p.checkbox.uncheckedFont) {
        const uAttrs: Record<string, string> = {};
        if (p.checkbox.uncheckedChar) {
          uAttrs["w14:val"] = p.checkbox.uncheckedChar;
        }
        if (p.checkbox.uncheckedFont) {
          uAttrs["w14:font"] = p.checkbox.uncheckedFont;
        }
        xml.leafNode("w14:uncheckedState", uAttrs);
      }
      xml.closeNode();
    }
    // Dropdown list
    if (p.dropdownList) {
      xml.openNode("w:dropDownList");
      for (const item of p.dropdownList) {
        const attrs: Record<string, string> = { "w:value": item.value };
        if (item.displayText) {
          attrs["w:displayText"] = item.displayText;
        }
        xml.leafNode("w:listItem", attrs);
      }
      xml.closeNode();
    }
    // ComboBox
    if (p.comboBox) {
      xml.openNode("w:comboBox");
      for (const item of p.comboBox) {
        const attrs: Record<string, string> = { "w:value": item.value };
        if (item.displayText) {
          attrs["w:displayText"] = item.displayText;
        }
        xml.leafNode("w:listItem", attrs);
      }
      xml.closeNode();
    }
    // Date picker
    if (p.date) {
      const dateAttrs: Record<string, string> = {};
      if (p.date.fullDate) {
        dateAttrs["w:fullDate"] = p.date.fullDate;
      }
      xml.openNode("w:date", Object.keys(dateAttrs).length > 0 ? dateAttrs : undefined);
      if (p.date.dateFormat) {
        xml.leafNode("w:dateFormat", { "w:val": p.date.dateFormat });
      }
      if (p.date.lid) {
        xml.leafNode("w:lid", { "w:val": p.date.lid });
      }
      if (p.date.storeMappedDataAs) {
        xml.leafNode("w:storeMappedDataAs", { "w:val": p.date.storeMappedDataAs });
      }
      xml.closeNode();
    }
    xml.closeNode();
  }

  xml.openNode("w:sdtContent");
  const sdtHelpers: RenderHelpers | undefined = ctx
    ? { imageRemap: ctx.imageRIdRemap, hyperlinkRIds: ctx.hyperlinkRIds }
    : undefined;
  for (const child of sdt.content) {
    if ("type" in child) {
      if (child.type === "paragraph") {
        renderParagraph(xml, child, sdtHelpers);
      } else if (child.type === "table") {
        renderTable(xml, child, sdtHelpers);
      }
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
    hyperlinkRIds: renderCtx.hyperlinkRIds
  };
  switch (content.type) {
    case "paragraph":
      renderParagraph(xml, content, helpers);
      break;
    case "table":
      renderTable(xml, content, helpers);
      break;
    case "floatingImage":
      renderFloatingImage(xml, content, renderCtx.imageRIdRemap);
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
      renderCheckBox(xml, content as CheckBox);
      break;
    case "drawingShape":
      renderDrawingShape(xml, content as DrawingShape, renderCtx);
      break;
    case "opaqueDrawing":
      renderOpaqueDrawing(xml, content);
      break;
    case "chart":
      renderChartDrawing(xml, content, renderCtx);
      break;
    case "chartEx":
      renderChartExDrawing(xml, content, renderCtx);
      break;
    case "altChunk":
      xml.leafNode("w:altChunk", { "r:id": content.rId });
      break;
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

  for (const content of doc.body) {
    renderBodyContent(xml, content, renderCtx);
  }

  // Final section properties (goes directly in w:body)
  if (doc.sectionProperties) {
    renderSectionProperties(xml, doc.sectionProperties);
  }

  xml.closeNode(); // w:body
  xml.closeNode(); // w:document
}

// =============================================================================
// DrawingML Shape Writer
// =============================================================================

function renderDrawingShape(xml: XmlSink, shape: DrawingShape, ctx: WordRenderContext): void {
  // If rawXml is available, use it for round-trip fidelity
  if (shape.rawXml) {
    xml.writeRaw(shape.rawXml);
    return;
  }

  // Wrap in w:p > w:r > w:drawing > wp:anchor > a:graphic > a:graphicData > wps:wsp
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

  const docPrAttrs: Record<string, string> = {
    id: String(ctx.ids.nextDocPrId()),
    name: shape.name ?? "Shape"
  };
  if (shape.altText) {
    docPrAttrs["descr"] = shape.altText;
  }
  xml.leafNode("wp:docPr", docPrAttrs);

  xml.openNode("a:graphic", { "xmlns:a": NS_A });
  xml.openNode("a:graphicData", { uri: NS_WPS });

  xml.openNode("wps:wsp");

  // Shape properties
  xml.openNode("wps:spPr");

  // Transform
  xml.openNode("a:xfrm", shape.rotation ? { rot: String(shape.rotation) } : {});
  xml.leafNode("a:off", { x: "0", y: "0" });
  xml.leafNode("a:ext", { cx: String(shape.width), cy: String(shape.height) });
  xml.closeNode(); // a:xfrm

  // Preset geometry
  xml.openNode("a:prstGeom", { prst: shape.shapeType });
  xml.leafNode("a:avLst");
  xml.closeNode(); // a:prstGeom

  // Fill
  if (shape.noFill) {
    xml.leafNode("a:noFill");
  } else if (shape.fillColor) {
    xml.openNode("a:solidFill");
    xml.leafNode("a:srgbClr", { val: shape.fillColor });
    xml.closeNode();
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

  xml.closeNode(); // wps:spPr

  // Text body
  if (shape.textContent && shape.textContent.length > 0) {
    xml.openNode("wps:txbx");
    xml.openNode("w:txbxContent");
    const txbxHelpers: RenderHelpers | undefined = ctx
      ? { imageRemap: ctx.imageRIdRemap, hyperlinkRIds: ctx.hyperlinkRIds }
      : undefined;
    for (const para of shape.textContent) {
      renderParagraph(xml, para, txbxHelpers);
    }
    xml.closeNode(); // w:txbxContent
    xml.closeNode(); // wps:txbx
  }

  // Body properties (required)
  xml.leafNode("wps:bodyPr");

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

function renderOpaqueDrawing(xml: XmlSink, drawing: OpaqueDrawing): void {
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

const NS_C_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_CX_CHART = "http://schemas.microsoft.com/office/drawing/2014/chartex";

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
