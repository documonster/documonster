/**
 * DOCX reader — header watermark detection.
 *
 * Watermarks are stored as a VML shape inside the first `<w:p>` of a header
 * part. We detect them by walking the header XML tree looking for a shape
 * whose id contains "WaterMark"; the surrounding header content is parsed
 * by the regular header-content parser.
 *
 * Self-contained — no dependency on the body parser, so it lives in its
 * own file to keep `docx-reader.ts` focused on package orchestration.
 */

import { findChild } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import type { ImageWatermark, TextWatermark, Watermark } from "../types";

/** Detect a watermark shape inside a header XML root, if present. */
export function detectWatermarkFromRoot(root: XmlElement): Watermark | undefined {
  // Look for VML shape with id containing "WaterMark"
  for (const pEl of root.children) {
    if (pEl.type !== "element") {
      continue;
    }
    for (const rEl of pEl.children) {
      if (rEl.type !== "element") {
        continue;
      }
      // Look for w:pict or w:r > w:pict
      const pictEls: XmlElement[] = [];
      const rName = rEl.name.replace(/^w:/, "");
      if (rName === "pict") {
        pictEls.push(rEl);
      } else if (rName === "r") {
        for (const rc of rEl.children) {
          if (rc.type === "element" && rc.name.replace(/^w:/, "") === "pict") {
            pictEls.push(rc);
          }
        }
      }
      for (const pictEl of pictEls) {
        for (const shapeEl of pictEl.children) {
          if (shapeEl.type !== "element") {
            continue;
          }
          const shapeId = shapeEl.attributes["id"] ?? "";
          if (!shapeId.toLowerCase().includes("watermark")) {
            continue;
          }
          // Found watermark shape
          const shapeType = shapeEl.attributes["type"] ?? "";
          if (shapeType.includes("136")) {
            // WordArt text watermark (shapetype 136)
            return parseTextWatermark(shapeEl);
          }
          // Check for image watermark (has v:imagedata)
          const imgData = findChild(shapeEl, "v:imagedata");
          if (imgData) {
            return parseImageWatermark(shapeEl, imgData);
          }
        }
      }
    }
  }
  return undefined;
}

function parseTextWatermark(shapeEl: XmlElement): TextWatermark {
  const fillColor = shapeEl.attributes["fillcolor"] ?? "#C0C0C0";
  const color = fillColor.replace(/^#/, "");

  // Parse rotation from style
  const style = shapeEl.attributes["style"] ?? "";
  let rotation = -45;
  const rotMatch = style.match(/rotation:\s*(-?\d+)/);
  if (rotMatch) {
    rotation = parseInt(rotMatch[1], 10);
  }

  // Get opacity from v:fill
  const fillEl = findChild(shapeEl, "v:fill");
  const opacity = fillEl?.attributes["opacity"] ?? ".5";
  const semiTransparent = opacity !== "1";

  // Get text and font from v:textpath
  const textpathEl = findChild(shapeEl, "v:textpath");
  const text = textpathEl?.attributes["string"] ?? "";
  const tpStyle = textpathEl?.attributes["style"] ?? "";
  let font: string | undefined;
  let fontSize: number | undefined;
  const fontMatch = tpStyle.match(/font-family:\s*"?([^";]+)"?/);
  if (fontMatch) {
    font = fontMatch[1].replace(/&quot;/g, "");
  }
  const sizeMatch = tpStyle.match(/font-size:\s*(\d+(?:\.\d+)?)\s*pt/);
  if (sizeMatch) {
    fontSize = Math.round(parseFloat(sizeMatch[1]) * 2); // convert pt to half-points
  }

  return {
    type: "text",
    text,
    font,
    fontSize,
    color,
    semiTransparent,
    rotation
  };
}

function parseImageWatermark(shapeEl: XmlElement, imgDataEl: XmlElement): ImageWatermark {
  const rId = imgDataEl.attributes["r:id"] ?? "";
  const gain = imgDataEl.attributes["gain"] ?? "";
  const washout = gain.startsWith("19661") || gain === "";

  return {
    type: "image",
    rId,
    washout
  };
}
