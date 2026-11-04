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

import type { ImageWatermark, TextWatermark, Watermark } from "@word/types";
import { findChild } from "@xml/dom";
import type { XmlElement } from "@xml/types";

/** A watermark found in a header, and the markup drawing it. */
export interface FoundWatermark {
  readonly watermark: Watermark;
  /** The header paragraph holding the watermark. */
  readonly paragraph: XmlElement;
  /** The child of `paragraph` — a `w:r`, or a bare `w:pict` — that draws it. */
  readonly carrier: XmlElement;
}

/** Find a watermark shape inside a header XML root, if present. */
export function findWatermark(root: XmlElement): FoundWatermark | undefined {
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
          let watermark: Watermark | undefined;
          if (shapeType.includes("136")) {
            // WordArt text watermark (shapetype 136)
            watermark = parseTextWatermark(shapeEl);
          } else {
            // Check for image watermark (has v:imagedata)
            const imgData = findChild(shapeEl, "v:imagedata");
            if (imgData) {
              watermark = parseImageWatermark(shapeEl, imgData);
            }
          }
          if (watermark) {
            return { watermark, paragraph: pEl, carrier: rEl };
          }
        }
      }
    }
  }
  return undefined;
}

/** `root` without the markup drawing `found`; the input is not modified. */
export function withoutWatermark(root: XmlElement, found: FoundWatermark): XmlElement {
  return {
    ...root,
    children: root.children.map(child =>
      child === found.paragraph
        ? {
            ...found.paragraph,
            children: found.paragraph.children.filter(c => c !== found.carrier)
          }
        : child
    )
  };
}

function parseTextWatermark(shapeEl: XmlElement): TextWatermark {
  const color = vmlColorToHex(shapeEl.attributes["fillcolor"] ?? "#C0C0C0");

  // Parse rotation from style. A shape without one is horizontal: Word's own
  // horizontal watermarks carry no rotation at all.
  const style = shapeEl.attributes["style"] ?? "";
  let rotation = 0;
  const rotMatch = style.match(/rotation:\s*(-?\d+(?:\.\d+)?)/);
  if (rotMatch) {
    rotation = parseFloat(rotMatch[1]);
  }
  const widthPt = styleLengthPt(style, "width");
  const heightPt = styleLengthPt(style, "height");

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
  // 1pt is Word's "auto": the text is stretched over the shape (fitshape), so
  // the box is the size and the text path carries none.
  if (sizeMatch && parseFloat(sizeMatch[1]) !== 1) {
    fontSize = Math.round(parseFloat(sizeMatch[1]) * 2); // convert pt to half-points
  }

  return {
    type: "text",
    text,
    font,
    fontSize,
    color,
    semiTransparent,
    rotation,
    ...(widthPt !== undefined ? { widthPt } : {}),
    ...(heightPt !== undefined ? { heightPt } : {})
  };
}

/**
 * The sixteen HTML 4 colour names, which VML colour attributes accept, as
 * `RRGGBB`. Word writes `fillcolor="silver"` for its own watermarks;
 * `TextWatermark.color` is hex.
 */
const VML_COLOR_NAMES: Readonly<Record<string, string>> = {
  black: "000000",
  silver: "C0C0C0",
  gray: "808080",
  white: "FFFFFF",
  maroon: "800000",
  red: "FF0000",
  purple: "800080",
  fuchsia: "FF00FF",
  green: "008000",
  lime: "00FF00",
  olive: "808000",
  yellow: "FFFF00",
  navy: "000080",
  blue: "0000FF",
  teal: "008080",
  aqua: "00FFFF"
};

/** A VML colour (`#RRGGBB`, `#RGB`, a name, optionally followed by ` [index]`) as `RRGGBB`. */
function vmlColorToHex(value: string): string {
  const color = value.trim().split(/\s+/)[0] ?? "";
  const named = VML_COLOR_NAMES[color.toLowerCase()];
  if (named) {
    return named;
  }
  const hex = color.replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    return [...hex]
      .map(c => c + c)
      .join("")
      .toUpperCase();
  }
  return hex;
}

function parseImageWatermark(shapeEl: XmlElement, imgDataEl: XmlElement): ImageWatermark {
  const rId = imgDataEl.attributes["r:id"] ?? "";
  const gain = imgDataEl.attributes["gain"] ?? "";
  const washout = gain.startsWith("19661") || gain === "";
  const style = shapeEl.attributes["style"] ?? "";
  const widthPt = styleLengthPt(style, "width");
  const heightPt = styleLengthPt(style, "height");

  return {
    type: "image",
    rId,
    washout,
    ...(widthPt !== undefined ? { widthPt } : {}),
    ...(heightPt !== undefined ? { heightPt } : {})
  };
}

/** A length in points from a VML `style` declaration, when it is given in `pt`. */
function styleLengthPt(style: string, property: string): number | undefined {
  const match = style.match(new RegExp(`(?:^|;)\\s*${property}:\\s*(\\d+(?:\\.\\d+)?)pt`));
  return match ? parseFloat(match[1]) : undefined;
}
