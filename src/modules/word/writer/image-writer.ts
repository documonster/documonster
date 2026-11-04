/**
 * DOCX Writers - Floating Image
 *
 * Renders wp:anchor elements for floating (non-inline) images
 * with full support for rotation, flip, lock anchor, outline, and wrap margins.
 */

import { NS_A, NS_PIC, URI_PIC, NS_ASVG, GUID_SVG } from "@word/constants";
import { drawingInteger } from "@word/core/drawing-geometry";
import type { FloatingImage } from "@word/types";
import { closeDrawingFrame, frameModel, openDrawingFrame } from "@word/writer/drawing-frame";
import type { RenderHelpers } from "@word/writer/render-context";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

/** Render a floating image as a standalone paragraph with wp:anchor. */
export function renderFloatingImage(
  xml: XmlSink,
  img: FloatingImage,
  helpers?: RenderHelpers
): void {
  xml.openNode("w:p");
  renderFloatingImageRun(xml, img, helpers);
  xml.closeNode(); // w:p
}

/** Render a floating image's run (`w:r` holding its wp:anchor), for a paragraph to hold. */
export function renderFloatingImageRun(
  xml: XmlSink,
  img: FloatingImage,
  helpers?: RenderHelpers
): void {
  const imageRemap = helpers?.imageRemap;
  const drawingId = helpers?.nextDocPrId?.() ?? img.drawingId ?? 1;
  const name = img.name ?? "Picture";
  // Resolve relationship id used in r:embed via packager-provided remap.
  const embedRId = imageRemap?.get(img.rId) ?? img.rId;
  const svgEmbedRId = img.svgRId ? (imageRemap?.get(img.svgRId) ?? img.svgRId) : undefined;

  xml.openNode("w:r");
  xml.openNode("w:drawing");

  const frame = frameModel("anchor", "picture", img, "Picture");
  const then = img._frame ? frameModel("anchor", "picture", img._frame.read, "Picture") : undefined;
  openDrawingFrame(xml, frame, img._frame, then, drawingId, helpers);

  xml.openNode("a:graphic", { "xmlns:a": NS_A });
  xml.openNode("a:graphicData", { uri: URI_PIC });

  xml.openNode("pic:pic", { "xmlns:pic": NS_PIC });

  xml.openNode("pic:nvPicPr");
  xml.leafNode("pic:cNvPr", { id: xmlValue(drawingId, "id"), name });
  xml.leafNode("pic:cNvPicPr");
  xml.closeNode();

  xml.openNode("pic:blipFill");
  if (img.svgRId) {
    // SVG with raster fallback: a:blip has extLst with asvg:svgBlip
    xml.openNode("a:blip", { "r:embed": embedRId });
    xml.openNode("a:extLst");
    xml.openNode("a:ext", { uri: GUID_SVG });
    xml.leafNode("asvg:svgBlip", {
      "xmlns:asvg": NS_ASVG,
      "r:embed": svgEmbedRId!
    });
    xml.closeNode(); // a:ext
    xml.closeNode(); // a:extLst
    xml.closeNode(); // a:blip
  } else {
    xml.leafNode("a:blip", { "r:embed": embedRId });
  }
  // Source rectangle (crop)
  if (img.srcRect) {
    const sr = img.srcRect;
    const srAttrs: Record<string, string> = {};
    if (sr.l !== undefined) {
      srAttrs["l"] = drawingInteger(sr.l, "srcRect.l");
    }
    if (sr.t !== undefined) {
      srAttrs["t"] = drawingInteger(sr.t, "srcRect.t");
    }
    if (sr.r !== undefined) {
      srAttrs["r"] = drawingInteger(sr.r, "srcRect.r");
    }
    if (sr.b !== undefined) {
      srAttrs["b"] = drawingInteger(sr.b, "srcRect.b");
    }
    xml.leafNode("a:srcRect", srAttrs);
  }
  xml.openNode("a:stretch");
  xml.leafNode("a:fillRect");
  xml.closeNode();
  xml.closeNode();

  xml.openNode("pic:spPr");

  // Transform with rotation/flip
  const xfrmAttrs: Record<string, string> = {};
  if (img.rotation !== undefined && img.rotation !== 0) {
    xfrmAttrs["rot"] = drawingInteger(img.rotation, "rotation");
  }
  if (img.flipHorizontal) {
    xfrmAttrs["flipH"] = "1";
  }
  if (img.flipVertical) {
    xfrmAttrs["flipV"] = "1";
  }
  xml.openNode("a:xfrm", Object.keys(xfrmAttrs).length > 0 ? xfrmAttrs : undefined);
  xml.leafNode("a:off", { x: "0", y: "0" });
  xml.leafNode("a:ext", {
    cx: drawingInteger(img.width, "width"),
    cy: drawingInteger(img.height, "height")
  });
  xml.closeNode();

  xml.openNode("a:prstGeom", { prst: "rect" });
  xml.leafNode("a:avLst");
  xml.closeNode();

  // Outline
  if (img.outline) {
    const lnAttrs: Record<string, string> = {};
    if (img.outline.width !== undefined) {
      lnAttrs["w"] = drawingInteger(img.outline.width, "outline.width");
    }
    xml.openNode("a:ln", Object.keys(lnAttrs).length > 0 ? lnAttrs : undefined);
    if (img.outline.color) {
      xml.openNode("a:solidFill");
      xml.leafNode("a:srgbClr", { val: img.outline.color });
      xml.closeNode();
    }
    xml.closeNode();
  }

  xml.closeNode(); // pic:spPr

  xml.closeNode(); // pic:pic
  xml.closeNode(); // a:graphicData
  xml.closeNode(); // a:graphic
  closeDrawingFrame(xml, frame, img._frame, helpers); // wp:anchor
  xml.closeNode(); // w:drawing
  xml.closeNode(); // w:r
}
