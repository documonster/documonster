/**
 * DOCX Reader - Image Parsers
 *
 * Parsers for inline and floating images (wp:inline / wp:anchor).
 * These are pure parsers that don't depend on ReaderContext.
 */

import { findChild, textContent } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import { type Mutable } from "../core/internal-utils";
import type { FloatingImage, InlineImageContent, OpaqueRunContent, RunContent } from "../types";
import { parseOutline, parseSrcRect, parseSvgBlip, parseXfrm } from "./drawing-helpers";
import { safeParseInt, serializeElement } from "./parse-utils";

function parseDrawingContent(drawingEl: XmlElement, content: RunContent[]): void {
  // Look for wp:inline
  const inlineEl = findChild(drawingEl, "wp:inline");
  if (inlineEl) {
    const extentEl = findChild(inlineEl, "wp:extent");
    const docPrEl = findChild(inlineEl, "wp:docPr");
    const graphicEl = findChild(inlineEl, "a:graphic");
    const graphicDataEl = graphicEl ? findChild(graphicEl, "a:graphicData") : undefined;
    const picEl = graphicDataEl ? findChild(graphicDataEl, "pic:pic") : undefined;
    const blipFillEl = picEl ? findChild(picEl, "pic:blipFill") : undefined;
    const blipEl = blipFillEl ? findChild(blipFillEl, "a:blip") : undefined;

    const rId = blipEl?.attributes["r:embed"];
    if (!rId) {
      // Non-picture inline drawing (chart, SmartArt, OLE, generic shape).
      // Preserve the entire <w:drawing>…</w:drawing> as an opaque run content
      // so the element survives a parse → write round-trip even when the
      // surrounding paragraph lives inside a table cell, header, footer or
      // SDT — locations where the body-level floating-content extractor
      // never runs.
      content.push({
        type: "opaqueRun",
        rawXml: serializeElement(drawingEl)
      } satisfies OpaqueRunContent);
      return;
    }

    const cx = safeParseInt(extentEl?.attributes["cx"], 0);
    const cy = safeParseInt(extentEl?.attributes["cy"], 0);

    const img: Mutable<InlineImageContent> = {
      type: "image",
      rId,
      width: cx,
      height: cy,
      altText: docPrEl?.attributes["descr"],
      name: docPrEl?.attributes["name"],
      drawingId: docPrEl ? safeParseInt(docPrEl.attributes["id"], 1) : undefined
    };

    // Parse xfrm for rotation/flip + outline
    const spPrEl = picEl ? findChild(picEl, "pic:spPr") : undefined;
    if (spPrEl) {
      parseXfrm(spPrEl, img);
      const outline = parseOutline(spPrEl);
      if (outline) {
        img.outline = outline;
      }
    }

    // Source rectangle (crop)
    if (blipFillEl) {
      const srcRect = parseSrcRect(blipFillEl);
      if (srcRect) {
        img.srcRect = srcRect;
      }
    }

    // SVG blip in a:extLst
    if (blipEl) {
      const svgRId = parseSvgBlip(blipEl);
      if (svgRId) {
        img.svgRId = svgRId;
      }
    }

    content.push(img);
  }
}

// =============================================================================
// Floating Image Parser
// =============================================================================

function parseFloatingImage(anchorEl: XmlElement): FloatingImage | undefined {
  const docPrEl = findChild(anchorEl, "wp:docPr");
  const extentEl = findChild(anchorEl, "wp:extent");
  const graphicEl = findChild(anchorEl, "a:graphic");
  const graphicDataEl = graphicEl ? findChild(graphicEl, "a:graphicData") : undefined;
  const picEl = graphicDataEl ? findChild(graphicDataEl, "pic:pic") : undefined;
  const blipFillEl = picEl ? findChild(picEl, "pic:blipFill") : undefined;
  const blipEl = blipFillEl ? findChild(blipFillEl, "a:blip") : undefined;

  const rId = blipEl?.attributes["r:embed"];
  if (!rId) {
    return undefined;
  }

  const cx = safeParseInt(extentEl?.attributes["cx"], 0);
  const cy = safeParseInt(extentEl?.attributes["cy"], 0);

  const img: Mutable<FloatingImage> = {
    type: "floatingImage",
    rId,
    width: cx,
    height: cy,
    altText: docPrEl?.attributes["descr"],
    name: docPrEl?.attributes["name"],
    drawingId: docPrEl ? safeParseInt(docPrEl.attributes["id"], 1) : undefined
  };

  // Attributes
  if (anchorEl.attributes["behindDoc"] === "1") {
    img.behindDoc = true;
  }
  if (anchorEl.attributes["locked"] === "1") {
    img.lockAnchor = true;
  }
  if (anchorEl.attributes["layoutInCell"] === "0") {
    img.layoutInCell = false;
  }
  if (anchorEl.attributes["allowOverlap"] === "0") {
    img.allowOverlap = false;
  }
  const rh = anchorEl.attributes["relativeHeight"];
  if (rh) {
    const n = parseInt(rh, 10);
    if (Number.isFinite(n)) {
      img.relativeHeight = n;
    }
  }
  // Dist* — guard each parseInt against NaN so a hostile attribute
  // never round-trips back into the output XML as the literal "NaN"
  // (Word rejects such files).
  const distT = anchorEl.attributes["distT"];
  if (distT) {
    const n = parseInt(distT, 10);
    if (Number.isFinite(n)) {
      img.distT = n;
    }
  }
  const distB = anchorEl.attributes["distB"];
  if (distB) {
    const n = parseInt(distB, 10);
    if (Number.isFinite(n)) {
      img.distB = n;
    }
  }
  const distL = anchorEl.attributes["distL"];
  if (distL) {
    const n = parseInt(distL, 10);
    if (Number.isFinite(n)) {
      img.distL = n;
    }
  }
  const distR = anchorEl.attributes["distR"];
  if (distR) {
    const n = parseInt(distR, 10);
    if (Number.isFinite(n)) {
      img.distR = n;
    }
  }

  // Simple positioning
  if (anchorEl.attributes["simplePos"] === "1") {
    const sposEl = findChild(anchorEl, "wp:simplePos");
    if (sposEl) {
      const x = safeParseInt(sposEl.attributes["x"], 0);
      const y = safeParseInt(sposEl.attributes["y"], 0);
      img.simplePos = { x, y };
    }
  }

  // Horizontal position
  const hPosEl = findChild(anchorEl, "wp:positionH");
  if (hPosEl) {
    const h: Mutable<NonNullable<FloatingImage["horizontalPosition"]>> = {
      relativeTo: hPosEl.attributes["relativeFrom"] as NonNullable<
        FloatingImage["horizontalPosition"]
      >["relativeTo"]
    };
    const offsetEl = findChild(hPosEl, "wp:posOffset");
    if (offsetEl) {
      const n = parseInt(textContent(offsetEl), 10);
      if (Number.isFinite(n)) {
        h.offset = n;
      }
    }
    const alignEl = findChild(hPosEl, "wp:align");
    if (alignEl) {
      h.align = textContent(alignEl) as NonNullable<FloatingImage["horizontalPosition"]>["align"];
    }
    img.horizontalPosition = h;
  }

  // Vertical position
  const vPosEl = findChild(anchorEl, "wp:positionV");
  if (vPosEl) {
    const v: Mutable<NonNullable<FloatingImage["verticalPosition"]>> = {
      relativeTo: vPosEl.attributes["relativeFrom"] as NonNullable<
        FloatingImage["verticalPosition"]
      >["relativeTo"]
    };
    const offsetEl = findChild(vPosEl, "wp:posOffset");
    if (offsetEl) {
      const n = parseInt(textContent(offsetEl), 10);
      if (Number.isFinite(n)) {
        v.offset = n;
      }
    }
    const alignEl = findChild(vPosEl, "wp:align");
    if (alignEl) {
      v.align = textContent(alignEl) as NonNullable<FloatingImage["verticalPosition"]>["align"];
    }
    img.verticalPosition = v;
  }

  // Wrap
  for (const wrapChild of anchorEl.children) {
    if (wrapChild.type !== "element") {
      continue;
    }
    const wn = wrapChild.name;
    let wrap: Mutable<NonNullable<FloatingImage["wrap"]>> | undefined;
    if (wn === "wp:wrapSquare") {
      wrap = {
        style: "square",
        side: wrapChild.attributes["wrapText"] as NonNullable<FloatingImage["wrap"]>["side"]
      };
    } else if (wn === "wp:wrapTight") {
      wrap = {
        style: "tight",
        side: wrapChild.attributes["wrapText"] as NonNullable<FloatingImage["wrap"]>["side"]
      };
    } else if (wn === "wp:wrapThrough") {
      wrap = {
        style: "through",
        side: wrapChild.attributes["wrapText"] as NonNullable<FloatingImage["wrap"]>["side"]
      };
    } else if (wn === "wp:wrapTopAndBottom") {
      wrap = { style: "topAndBottom" };
    } else if (wn === "wp:wrapNone") {
      wrap = { style: "none" };
    }
    if (wrap) {
      // Parse wrap margins
      const distT = anchorEl.attributes["distT"];
      const distB = anchorEl.attributes["distB"];
      const distL = anchorEl.attributes["distL"];
      const distR = anchorEl.attributes["distR"];
      if (distT || distB || distL || distR) {
        const margins: { top?: number; bottom?: number; left?: number; right?: number } = {};
        if (distT) {
          const n = parseInt(distT, 10);
          if (Number.isFinite(n)) {
            margins.top = n;
          }
        }
        if (distB) {
          const n = parseInt(distB, 10);
          if (Number.isFinite(n)) {
            margins.bottom = n;
          }
        }
        if (distL) {
          const n = parseInt(distL, 10);
          if (Number.isFinite(n)) {
            margins.left = n;
          }
        }
        if (distR) {
          const n = parseInt(distR, 10);
          if (Number.isFinite(n)) {
            margins.right = n;
          }
        }
        wrap.margins = margins;
      }
      img.wrap = wrap;
      break;
    }
  }

  // Rotation/flip from spPr
  const spPrEl = picEl ? findChild(picEl, "pic:spPr") : undefined;
  if (spPrEl) {
    parseXfrm(spPrEl, img);
    const outline = parseOutline(spPrEl);
    if (outline) {
      img.outline = outline;
    }
  }

  // SVG blip in a:extLst
  if (blipEl) {
    const svgRId = parseSvgBlip(blipEl);
    if (svgRId) {
      img.svgRId = svgRId;
    }
  }

  // Source rectangle (crop)
  if (blipFillEl) {
    const srcRect = parseSrcRect(blipFillEl);
    if (srcRect) {
      img.srcRect = srcRect;
    }
  }

  return img;
}

export { parseDrawingContent, parseFloatingImage };
