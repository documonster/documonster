/**
 * DOCX Writers - Floating Image
 *
 * Renders wp:anchor elements for floating (non-inline) images
 * with full support for rotation, flip, lock anchor, outline, and wrap margins.
 */

import type { XmlSink } from "@xml/types";

import { NS_A, NS_PIC, URI_PIC, NS_ASVG, GUID_SVG } from "../constants";
import type { FloatingImage } from "../types";

/** Render a floating image as a standalone paragraph with wp:anchor. */
export function renderFloatingImage(xml: XmlSink, img: FloatingImage): void {
  const drawingId = img.drawingId ?? 1;
  const name = img.name ?? "Picture";

  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.openNode("w:drawing");

  // Wrap margins
  const wrapMargins = img.wrap?.margins;

  xml.openNode("wp:anchor", {
    distT: String(img.distT ?? wrapMargins?.top ?? 0),
    distB: String(img.distB ?? wrapMargins?.bottom ?? 0),
    distL: String(img.distL ?? wrapMargins?.left ?? 114300),
    distR: String(img.distR ?? wrapMargins?.right ?? 114300),
    simplePos: img.simplePos ? "1" : "0",
    relativeHeight: String(img.relativeHeight ?? 251658240),
    behindDoc: img.behindDoc ? "1" : "0",
    locked: img.lockAnchor ? "1" : "0",
    layoutInCell: img.layoutInCell === false ? "0" : "1",
    allowOverlap: img.allowOverlap === false ? "0" : "1"
  });

  xml.leafNode("wp:simplePos", {
    x: String(img.simplePos?.x ?? 0),
    y: String(img.simplePos?.y ?? 0)
  });

  // Horizontal position
  {
    const hp = img.horizontalPosition;
    xml.openNode("wp:positionH", { relativeFrom: hp?.relativeTo ?? "column" });
    if (hp?.align) {
      xml.openNode("wp:align");
      xml.writeText(hp.align);
      xml.closeNode();
    } else {
      xml.openNode("wp:posOffset");
      xml.writeText(String(hp?.offset ?? 0));
      xml.closeNode();
    }
    xml.closeNode();
  }

  // Vertical position
  {
    const vp = img.verticalPosition;
    xml.openNode("wp:positionV", { relativeFrom: vp?.relativeTo ?? "paragraph" });
    if (vp?.align) {
      xml.openNode("wp:align");
      xml.writeText(vp.align);
      xml.closeNode();
    } else {
      xml.openNode("wp:posOffset");
      xml.writeText(String(vp?.offset ?? 0));
      xml.closeNode();
    }
    xml.closeNode();
  }

  xml.leafNode("wp:extent", { cx: String(img.width), cy: String(img.height) });
  xml.leafNode("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" });

  // Wrapping
  if (img.wrap) {
    switch (img.wrap.style) {
      case "square":
        xml.leafNode("wp:wrapSquare", { wrapText: img.wrap.side ?? "bothSides" });
        break;
      case "tight":
        xml.leafNode("wp:wrapTight", { wrapText: img.wrap.side ?? "bothSides" });
        break;
      case "through":
        xml.leafNode("wp:wrapThrough", { wrapText: img.wrap.side ?? "bothSides" });
        break;
      case "topAndBottom":
        xml.leafNode("wp:wrapTopAndBottom");
        break;
      case "none":
        xml.leafNode("wp:wrapNone");
        break;
    }
  } else {
    xml.leafNode("wp:wrapNone");
  }

  xml.leafNode("wp:docPr", {
    id: String(drawingId),
    name,
    ...(img.altText ? { descr: img.altText } : {})
  });

  xml.openNode("wp:cNvGraphicFramePr");
  xml.leafNode("a:graphicFrameLocks", { "xmlns:a": NS_A, noChangeAspect: "1" });
  xml.closeNode();

  xml.openNode("a:graphic", { "xmlns:a": NS_A });
  xml.openNode("a:graphicData", { uri: URI_PIC });

  xml.openNode("pic:pic", { "xmlns:pic": NS_PIC });

  xml.openNode("pic:nvPicPr");
  xml.leafNode("pic:cNvPr", { id: String(drawingId), name });
  xml.leafNode("pic:cNvPicPr");
  xml.closeNode();

  xml.openNode("pic:blipFill");
  if (img.svgRId) {
    // SVG with raster fallback: a:blip has extLst with asvg:svgBlip
    xml.openNode("a:blip", { "r:embed": img.rId });
    xml.openNode("a:extLst");
    xml.openNode("a:ext", { uri: GUID_SVG });
    xml.leafNode("asvg:svgBlip", {
      "xmlns:asvg": NS_ASVG,
      "r:embed": img.svgRId
    });
    xml.closeNode(); // a:ext
    xml.closeNode(); // a:extLst
    xml.closeNode(); // a:blip
  } else {
    xml.leafNode("a:blip", { "r:embed": img.rId });
  }
  // Source rectangle (crop)
  if (img.srcRect) {
    const sr = img.srcRect;
    const srAttrs: Record<string, string> = {};
    if (sr.l !== undefined) {
      srAttrs["l"] = String(sr.l);
    }
    if (sr.t !== undefined) {
      srAttrs["t"] = String(sr.t);
    }
    if (sr.r !== undefined) {
      srAttrs["r"] = String(sr.r);
    }
    if (sr.b !== undefined) {
      srAttrs["b"] = String(sr.b);
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
  if (img.rotation) {
    xfrmAttrs["rot"] = String(img.rotation);
  }
  if (img.flipHorizontal) {
    xfrmAttrs["flipH"] = "1";
  }
  if (img.flipVertical) {
    xfrmAttrs["flipV"] = "1";
  }
  xml.openNode("a:xfrm", Object.keys(xfrmAttrs).length > 0 ? xfrmAttrs : undefined);
  xml.leafNode("a:off", { x: "0", y: "0" });
  xml.leafNode("a:ext", { cx: String(img.width), cy: String(img.height) });
  xml.closeNode();

  xml.openNode("a:prstGeom", { prst: "rect" });
  xml.leafNode("a:avLst");
  xml.closeNode();

  // Outline
  if (img.outline) {
    const lnAttrs: Record<string, string> = {};
    if (img.outline.width !== undefined) {
      lnAttrs["w"] = String(img.outline.width);
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
  xml.closeNode(); // wp:anchor
  xml.closeNode(); // w:drawing
  xml.closeNode(); // w:r
  xml.closeNode(); // w:p
}
