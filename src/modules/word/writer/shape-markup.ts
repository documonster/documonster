/**
 * The `wps:wsp` of a DrawingML shape: its modelled fields, and whatever the
 * reader kept that they do not model (`DrawingShape._preserved`), each piece
 * in its schema position:
 *
 * ```
 * wsp     cNvPr cNvSpPr spPr style? extLst? (txbx | linkedTxbx)? bodyPr
 * spPr    xfrm? geometry? fill? ln? effects/3-D/extLst*
 * ln      fill? dash? join? headEnd? tailEnd? extLst?
 * ```
 *
 * A preserved fragment that stands for a modelled field is written only
 * while the field still holds the value it was read with, so an edited field
 * wins. Fragments go through `prepareRawMarkup` like all markup kept
 * verbatim, so image keys follow the part.
 */

import { drawingInteger } from "@word/core/drawing-geometry";
import { DocxRawXmlPolicyError } from "@word/errors";
import type { DrawingShape, DrawingShapePreserved } from "@word/types";
import { rebindPreservedLinks } from "@word/writer/drawing-frame";
import { renderParagraph } from "@word/writer/paragraph-writer";
import type { DocxRenderContext } from "@word/writer/render-context";
import { prepareRawMarkup, renderHelpersFor } from "@word/writer/render-context";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

export function renderShapeBody(
  xml: XmlSink,
  shape: DrawingShape,
  drawingId: number,
  ctx: DocxRenderContext
): void {
  const preserved = shape._preserved;
  if (preserved && ctx.rawXmlPolicy === "reject" && hasMarkup(preserved)) {
    throw new DocxRawXmlPolicyError("drawingShape._preserved");
  }
  // Under "strip", markup is dropped and only the modelled shape is written.
  const p: DrawingShapePreserved | undefined =
    preserved && ctx.rawXmlPolicy === "strip" ? stripMarkup(preserved) : preserved;
  const helpers = renderHelpersFor(ctx);
  const raw = (fragment: string | undefined): void => {
    if (fragment) {
      xml.writeRaw(prepareRawMarkup(fragment, helpers));
    }
  };
  const read = p?.read;

  xml.openNode("wps:wsp", { ...p?.wspAttrs });

  // Non-visual properties — required by the wordprocessingShape schema
  // (CT_WordprocessingShape.cNvPr + cNvSpPr). The id must match
  // wp:docPr/@id so Word treats them as the same logical object.
  const cNvPrAttrs = {
    id: xmlValue(drawingId, "id"),
    name: shape.name ?? "Shape",
    ...p?.cNvPrAttrs
  };
  if (p?.cNvPrChildrenXml) {
    xml.openNode("wps:cNvPr", cNvPrAttrs);
    xml.writeRaw(rebindPreservedLinks(p.cNvPrChildrenXml, p.links, helpers));
    xml.closeNode();
  } else {
    xml.leafNode("wps:cNvPr", cNvPrAttrs);
  }
  if (p?.cNvSpPrXml) {
    raw(p.cNvSpPrXml);
  } else {
    xml.leafNode("wps:cNvSpPr");
  }

  // Shape properties
  xml.openNode("wps:spPr", { ...p?.spPrAttrs });
  renderTransform(xml, shape, p);

  // Geometry: custom or adjusted geometry while the preset is the one read.
  if (p?.geometryXml && shape.shapeType === read!.shapeType) {
    raw(p.geometryXml);
  } else {
    xml.openNode("a:prstGeom", { prst: shape.shapeType });
    xml.leafNode("a:avLst");
    xml.closeNode();
  }

  // Fill (basic + advanced). The OOXML schema requires fill children of
  // spPr to appear before a:ln, hence advanced gradient/pattern fills are
  // injected here rather than alongside effects.
  const fillAsRead =
    p?.fillXml !== undefined &&
    shape.fillColor === read!.fillColor &&
    Boolean(shape.noFill) === Boolean(read!.noFill);
  if (fillAsRead) {
    raw(p.fillXml);
  } else {
    if (shape.noFill) {
      xml.leafNode("a:noFill");
    } else if (shape.fillColor) {
      xml.openNode("a:solidFill");
      xml.leafNode("a:srgbClr", { val: shape.fillColor });
      xml.closeNode();
    }
    if (shape._advancedFillXml && !shape.noFill && !shape.fillColor) {
      if (ctx.rawXmlPolicy === "reject") {
        throw new DocxRawXmlPolicyError("drawingShape._advancedFillXml");
      }
      if (ctx.rawXmlPolicy !== "strip") {
        raw(shape._advancedFillXml);
      }
    }
  }

  renderLine(xml, shape, p, raw);

  // Effects, 3-D and extensions follow a:ln per the OOXML schema.
  if (p?.spPrTailXml) {
    raw(p.spPrTailXml);
  } else if (shape._advancedEffectsXml) {
    if (ctx.rawXmlPolicy === "reject") {
      throw new DocxRawXmlPolicyError("drawingShape._advancedEffectsXml");
    }
    if (ctx.rawXmlPolicy !== "strip") {
      raw(shape._advancedEffectsXml);
    }
  } else if (shape.rawXml && !shape._advancedFillXml) {
    // Backwards-compat: an opaque rawXml without the split fields goes after
    // a:ln, as it always has.
    if (ctx.rawXmlPolicy === "reject") {
      throw new DocxRawXmlPolicyError("drawingShape.rawXml");
    }
    if (ctx.rawXmlPolicy !== "strip") {
      raw(shape.rawXml);
    }
  }
  xml.closeNode(); // wps:spPr

  raw(p?.styleXml);
  raw(p?.wspExtLstXml);

  // Text body
  if (p?.linkedTxbxXml) {
    raw(p.linkedTxbxXml);
  } else {
    const contentAsRead =
      p?.txbxContentXml !== undefined && shape.textContent === read!.textContent;
    if (contentAsRead || (shape.textContent && shape.textContent.length > 0)) {
      xml.openNode("wps:txbx", { ...p?.txbxAttrs });
      if (contentAsRead) {
        raw(p.txbxContentXml);
      } else {
        xml.openNode("w:txbxContent");
        for (const para of shape.textContent ?? []) {
          renderParagraph(xml, para, helpers);
        }
        xml.closeNode(); // w:txbxContent
      }
      raw(p?.txbxExtLstXml);
      xml.closeNode(); // wps:txbx
    }
  }

  // Body properties (required). The vertical text anchor lives on a:bodyPr/@anchor.
  const bodyPrAttrs: Record<string, string> = { ...p?.bodyPrAttrs };
  if (shape.textBodyAnchor) {
    bodyPrAttrs["anchor"] = shape.textBodyAnchor;
  }
  if (p?.bodyPrChildrenXml) {
    xml.openNode("wps:bodyPr", bodyPrAttrs);
    raw(p.bodyPrChildrenXml);
    xml.closeNode();
  } else {
    xml.leafNode("wps:bodyPr", bodyPrAttrs);
  }

  xml.closeNode(); // wps:wsp
}

function renderTransform(
  xml: XmlSink,
  shape: DrawingShape,
  p: DrawingShapePreserved | undefined
): void {
  const sizeAsRead =
    p !== undefined && shape.width === p.read.width && shape.height === p.read.height;
  const xfrmAttrs: Record<string, string> = {};
  if (shape.rotation !== undefined && shape.rotation !== 0) {
    xfrmAttrs["rot"] = drawingInteger(shape.rotation, "rotation");
  }
  if (shape.flipHorizontal) {
    xfrmAttrs["flipH"] = "1";
  }
  if (shape.flipVertical) {
    xfrmAttrs["flipV"] = "1";
  }
  // A shape read without a transform keeps having none, until something
  // needs one.
  if (p?.xfrm === null && sizeAsRead && Object.keys(xfrmAttrs).length === 0) {
    return;
  }
  const kept = p?.xfrm ?? undefined;
  xml.openNode("a:xfrm", xfrmAttrs);
  if (!kept || kept.off !== undefined) {
    xml.leafNode("a:off", kept?.off ?? { x: "0", y: "0" });
  }
  if (!kept || kept.ext !== undefined || !sizeAsRead) {
    xml.leafNode(
      "a:ext",
      kept?.ext && sizeAsRead
        ? kept.ext
        : { cx: drawingInteger(shape.width, "width"), cy: drawingInteger(shape.height, "height") }
    );
  }
  xml.closeNode(); // a:xfrm
}

function renderLine(
  xml: XmlSink,
  shape: DrawingShape,
  p: DrawingShapePreserved | undefined,
  raw: (fragment: string | undefined) => void
): void {
  const fillAsRead =
    p?.lnFillXml !== undefined &&
    shape.outlineColor === p.read.outlineColor &&
    Boolean(shape.noOutline) === Boolean(p.read.noOutline);
  const hasLine =
    shape.outlineColor !== undefined ||
    shape.noOutline ||
    shape.outlineWidth !== undefined ||
    p?.lnAttrs !== undefined ||
    fillAsRead ||
    p?.lnTailXml !== undefined;
  if (!hasLine) {
    return;
  }
  const lnAttrs: Record<string, string> = {};
  if (shape.outlineWidth !== undefined && shape.outlineWidth !== 0) {
    lnAttrs["w"] = drawingInteger(shape.outlineWidth, "outlineWidth");
  }
  Object.assign(lnAttrs, p?.lnAttrs);
  xml.openNode("a:ln", lnAttrs);
  if (fillAsRead) {
    raw(p.lnFillXml);
  } else if (shape.noOutline) {
    xml.leafNode("a:noFill");
  } else if (shape.outlineColor) {
    xml.openNode("a:solidFill");
    xml.leafNode("a:srgbClr", { val: shape.outlineColor });
    xml.closeNode();
  }
  raw(p?.lnTailXml);
  xml.closeNode(); // a:ln
}

const MARKUP_FIELDS: readonly (keyof DrawingShapePreserved)[] = [
  "cNvPrChildrenXml",
  "cNvSpPrXml",
  "geometryXml",
  "fillXml",
  "lnFillXml",
  "lnTailXml",
  "spPrTailXml",
  "styleXml",
  "wspExtLstXml",
  "txbxContentXml",
  "txbxExtLstXml",
  "linkedTxbxXml",
  "bodyPrChildrenXml"
];

function hasMarkup(p: DrawingShapePreserved): boolean {
  return MARKUP_FIELDS.some(field => p[field] !== undefined);
}

function stripMarkup(p: DrawingShapePreserved): DrawingShapePreserved {
  const out: Record<string, unknown> = { ...p };
  for (const field of MARKUP_FIELDS) {
    delete out[field];
  }
  return out as unknown as DrawingShapePreserved;
}
