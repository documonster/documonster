/**
 * DOCX Writers - Headers & Footers
 *
 * Renders word/header{N}.xml and word/footer{N}.xml parts.
 */

import type { XmlSink } from "@xml/types";

import {
  NS_W,
  NS_R,
  NS_V,
  NS_O,
  NS_W10,
  NS_WP,
  NS_M,
  NS_W14,
  NS_W15,
  NS_WP14,
  NS_MC,
  STD_DOC_ATTRIBUTES
} from "../constants";
import type { HeaderFooterContent, TextWatermark, Watermark } from "../types";
import { renderParagraph } from "./paragraph-writer";
import type { RenderHelpers } from "./render-context";
import { renderTable } from "./table-writer";

/** Render content blocks (shared between header and footer). */
function renderContent(xml: XmlSink, content: HeaderFooterContent, helpers?: RenderHelpers): void {
  for (const child of content.children) {
    if (child.type === "paragraph") {
      renderParagraph(xml, child, helpers);
    } else if (child.type === "table") {
      renderTable(xml, child, helpers);
    }
  }
}

/** Standard namespace attributes for header/footer parts — must include DrawingML namespaces. */
const HEADER_FOOTER_NAMESPACES: Record<string, string> = {
  "xmlns:w": NS_W,
  "xmlns:r": NS_R,
  "xmlns:wp": NS_WP,
  "xmlns:m": NS_M,
  "xmlns:v": NS_V,
  "xmlns:o": NS_O,
  "xmlns:w10": NS_W10,
  "xmlns:w14": NS_W14,
  "xmlns:w15": NS_W15,
  "xmlns:wp14": NS_WP14,
  "xmlns:mc": NS_MC,
  "mc:Ignorable": "w14 w15 wp14"
};

/** Render a header part (word/header{N}.xml). */
export function renderHeader(
  xml: XmlSink,
  content: HeaderFooterContent,
  helpers?: RenderHelpers
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:hdr", HEADER_FOOTER_NAMESPACES);
  renderContent(xml, content, helpers);
  xml.closeNode();
}

/** Render a footer part (word/footer{N}.xml). */
export function renderFooter(
  xml: XmlSink,
  content: HeaderFooterContent,
  helpers?: RenderHelpers
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:ftr", HEADER_FOOTER_NAMESPACES);
  renderContent(xml, content, helpers);
  xml.closeNode();
}

/** Render a header part containing a watermark. */
export function renderWatermarkHeader(xml: XmlSink, watermark: Watermark, imageRId?: string): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:hdr", {
    "xmlns:w": NS_W,
    "xmlns:r": NS_R,
    "xmlns:v": NS_V,
    "xmlns:o": NS_O,
    "xmlns:w10": NS_W10
  });

  // Watermark goes in a paragraph with a single pict run
  xml.openNode("w:p");
  xml.openNode("w:pPr");
  xml.openNode("w:pStyle");
  xml.closeNode();
  xml.closeNode(); // pPr

  xml.openNode("w:r");
  xml.openNode("w:pict");

  if (watermark.type === "text") {
    renderTextWatermarkVml(xml, watermark);
  } else {
    renderImageWatermarkVml(xml, watermark, imageRId);
  }

  xml.closeNode(); // pict
  xml.closeNode(); // r
  xml.closeNode(); // p

  xml.closeNode(); // hdr
}

function renderTextWatermarkVml(xml: XmlSink, wm: TextWatermark): void {
  const color = wm.color ?? "C0C0C0";
  const font = wm.font ?? "Calibri";
  const fontSize = wm.fontSize ?? 1; // half-points; Word uses pt string in style
  const fontPt = fontSize / 2;
  const rotation = wm.rotation ?? -45;
  const opacity = wm.semiTransparent !== false ? ".5" : "1";

  // VML shape for text watermark (PowerWash / WASHOUT style)
  xml.leafNode("v:shapetype", {
    id: "_x0000_t136",
    coordsize: "21600,21600",
    "o:spt": "136",
    path: "m@7,l@8,m@5,21600l@6,21600e"
  });

  const style =
    `position:absolute;margin-left:0;margin-top:0;width:468pt;height:234pt;` +
    `rotation:${rotation};z-index:-251658752;mso-position-horizontal:center;` +
    `mso-position-horizontal-relative:margin;mso-position-vertical:center;` +
    `mso-position-vertical-relative:margin`;

  xml.openNode("v:shape", {
    id: "PowerPlusWaterMarkObject",
    "o:spid": "_x0000_s2049",
    type: "#_x0000_t136",
    style,
    "o:allowincell": "f",
    fillcolor: `#${color}`,
    stroked: "f"
  });

  xml.leafNode("v:fill", { opacity });
  xml.leafNode("v:textpath", {
    style: `font-family:&quot;${font}&quot;;font-size:${fontPt}pt`,
    string: wm.text
  });
  xml.leafNode("w10:wrap", { anchorx: "margin", anchory: "margin" });

  xml.closeNode(); // v:shape
}

function renderImageWatermarkVml(
  xml: XmlSink,
  wm: Watermark & { type: "image" },
  rId?: string
): void {
  const scale = wm.scale ?? 100;
  const rid = rId ?? wm.rId;

  const style =
    `position:absolute;margin-left:0;margin-top:0;width:0;height:0;` +
    `z-index:-251658752;mso-position-horizontal:center;` +
    `mso-position-horizontal-relative:margin;mso-position-vertical:center;` +
    `mso-position-vertical-relative:margin`;

  xml.openNode("v:shape", {
    id: "PowerPlusWaterMarkObject",
    "o:spid": "_x0000_s2050",
    type: "",
    style,
    "o:allowincell": "f"
  });

  const gain = wm.washout !== false ? "19661f" : "65536f";
  const blacklevel = wm.washout !== false ? "22938f" : "0";

  xml.leafNode("v:imagedata", {
    "r:id": rid,
    "o:title": "",
    gain,
    blacklevel,
    ...(scale !== 100 ? { "o:detectmouseclick": "t" } : {})
  });
  xml.leafNode("w10:wrap", { anchorx: "margin", anchory: "margin" });

  xml.closeNode(); // v:shape
}
