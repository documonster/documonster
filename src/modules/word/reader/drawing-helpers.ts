/**
 * DOCX Reader - Drawing/Image Helpers
 *
 * Shared parsing helpers for DrawingML elements (used by both inline image,
 * floating image, and drawing shape parsers).
 */

import { findChild, findChildren } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import { findChildNs } from "./parse-utils";

/** Parse `a:xfrm` element for rotation and flip properties. */
export function parseXfrm(
  spPrEl: XmlElement,
  target: { rotation?: number; flipHorizontal?: boolean; flipVertical?: boolean }
): void {
  const xfrmEl = findChild(spPrEl, "a:xfrm");
  if (xfrmEl) {
    const rot = xfrmEl.attributes["rot"];
    if (rot !== undefined && rot !== "") {
      target.rotation = parseInt(rot, 10);
    }
    if (xfrmEl.attributes["flipH"] === "1") {
      target.flipHorizontal = true;
    }
    if (xfrmEl.attributes["flipV"] === "1") {
      target.flipVertical = true;
    }
  }
}

/** Parse outline (`a:ln`) element. */
export function parseOutline(spPrEl: XmlElement): { width?: number; color?: string } | undefined {
  const lnEl = findChild(spPrEl, "a:ln");
  if (!lnEl) {
    return undefined;
  }
  const outline: { width?: number; color?: string } = {};
  const w = lnEl.attributes["w"];
  if (w) {
    outline.width = parseInt(w, 10);
  }
  const sfEl = findChild(lnEl, "a:solidFill");
  const srgbEl = sfEl ? findChild(sfEl, "a:srgbClr") : undefined;
  if (srgbEl) {
    outline.color = srgbEl.attributes["val"];
  }
  return Object.keys(outline).length > 0 ? outline : undefined;
}

/** Parse source rectangle (crop) from `a:blipFill`. */
export function parseSrcRect(
  blipFillEl: XmlElement
): { l?: number; t?: number; r?: number; b?: number } | undefined {
  const srcRectEl = findChild(blipFillEl, "a:srcRect");
  if (!srcRectEl) {
    return undefined;
  }
  const sr: { l?: number; t?: number; r?: number; b?: number } = {};
  const lAttr = srcRectEl.attributes["l"];
  const tAttr = srcRectEl.attributes["t"];
  const rAttr = srcRectEl.attributes["r"];
  const bAttr = srcRectEl.attributes["b"];
  if (lAttr !== undefined) {
    sr.l = parseInt(lAttr, 10);
  }
  if (tAttr !== undefined) {
    sr.t = parseInt(tAttr, 10);
  }
  if (rAttr !== undefined) {
    sr.r = parseInt(rAttr, 10);
  }
  if (bAttr !== undefined) {
    sr.b = parseInt(bAttr, 10);
  }
  return Object.keys(sr).length > 0 ? sr : undefined;
}

/** Parse SVG blip reference from `a:extLst` inside `a:blip`. */
export function parseSvgBlip(blipEl: XmlElement): string | undefined {
  const extLst = findChild(blipEl, "a:extLst");
  if (!extLst) {
    return undefined;
  }
  for (const ext of findChildren(extLst, "a:ext")) {
    const svgBlip = findChild(ext, "asvg:svgBlip") ?? findChildNs(ext, "svgBlip");
    if (svgBlip) {
      const svgEmbed = svgBlip.attributes["r:embed"];
      if (svgEmbed) {
        return svgEmbed;
      }
    }
  }
  return undefined;
}
