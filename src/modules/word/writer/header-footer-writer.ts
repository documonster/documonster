/**
 * DOCX Writers - Headers & Footers
 *
 * Renders word/header{N}.xml and word/footer{N}.xml parts.
 */

import { mapToStandardFont, measureTextWidth } from "@utils/font-metrics";
import {
  DOCUMENT_NAMESPACES,
  NS_W,
  NS_R,
  NS_V,
  NS_O,
  NS_W10,
  STD_DOC_ATTRIBUTES
} from "@word/constants";
import { atNode, unwritableNode, writableList } from "@word/core/node-path";
import type { HeaderFooterContent, TextWatermark, Watermark } from "@word/types";
import { renderParagraph } from "@word/writer/paragraph-writer";
import type { RenderHelpers } from "@word/writer/render-context";
import { renderTable } from "@word/writer/table-writer";
import type { XmlSink } from "@xml/types";

/** A watermark drawn by a header, and the id its picture has in that header's rels. */
export interface HeaderWatermark {
  readonly watermark: Watermark;
  readonly imageRId?: string;
}

/**
 * Render content blocks (shared between header and footer). A watermark is
 * drawn by the first paragraph, as Word places it; a header that starts with
 * a table, or has no content, gets a paragraph of its own for it.
 */
function renderContent(
  xml: XmlSink,
  content: HeaderFooterContent,
  helpers?: RenderHelpers,
  watermark?: HeaderWatermark
): void {
  let emitted = 0;
  const leading = watermark
    ? (sink: XmlSink): void => renderWatermarkRun(sink, watermark.watermark, watermark.imageRId)
    : undefined;
  const children = writableList(content.children, "children");
  const first = children.find(c => c?.type === "paragraph" || c?.type === "table");
  if (leading && first?.type !== "paragraph") {
    renderWatermarkParagraph(xml, watermark!);
    emitted++;
  }
  for (let i = 0; i < children.length; i++) {
    const child = children[i]!;
    try {
      if (child?.type === "paragraph") {
        renderParagraph(xml, child, helpers, child === first ? leading : undefined);
        emitted++;
      } else if (child?.type === "table") {
        renderTable(xml, child, helpers);
        emitted++;
      } else {
        unwritableNode(
          "type",
          "a paragraph or a table",
          (child as { type?: unknown } | null)?.type
        );
      }
    } catch (error) {
      throw atNode(`children[${i}]`, error);
    }
  }
  // OOXML requires every header/footer part to contain at least one
  // <w:p>. When the caller passes an empty children array we still emit
  // a placeholder so Word does not refuse to load the package.
  if (emitted === 0) {
    xml.openNode("w:p");
    xml.closeNode();
  }
}

/** Standard namespace attributes for header/footer parts — must include DrawingML namespaces. */
/**
 * Every prefix the document part declares: drawings and markup kept verbatim
 * in a header use the same ones (`wps:` for a shape, `wp14:` on a frame).
 */
const HEADER_FOOTER_NAMESPACES: Readonly<Record<string, string>> = DOCUMENT_NAMESPACES;

/** Render a header part (word/header{N}.xml). */
export function renderHeader(
  xml: XmlSink,
  content: HeaderFooterContent,
  helpers?: RenderHelpers,
  watermark?: HeaderWatermark
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:hdr", HEADER_FOOTER_NAMESPACES);
  renderContent(xml, content, helpers, watermark);
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

  renderWatermarkParagraph(xml, { watermark, imageRId });

  xml.closeNode(); // hdr
}

/**
 * The paragraph a header without one of its own draws its watermark in:
 * styled as a Header, matching what Microsoft Word emits. The reader
 * recognises exactly this paragraph as carrying nothing but the watermark.
 */
function renderWatermarkParagraph(xml: XmlSink, watermark: HeaderWatermark): void {
  xml.openNode("w:p");
  xml.openNode("w:pPr");
  xml.leafNode("w:pStyle", { "w:val": "Header" });
  xml.closeNode(); // pPr
  renderWatermarkRun(xml, watermark.watermark, watermark.imageRId);
  xml.closeNode(); // p
}

/** The run drawing a watermark. It carries <w:noProof/> so WordArt text is not spell-checked. */
function renderWatermarkRun(xml: XmlSink, watermark: Watermark, imageRId?: string): void {
  xml.openNode("w:r");
  xml.openNode("w:rPr");
  xml.leafNode("w:noProof");
  xml.closeNode(); // rPr
  xml.openNode("w:pict");

  if (watermark.type === "text") {
    renderTextWatermarkVml(xml, watermark);
  } else {
    renderImageWatermarkVml(xml, watermark, imageRId);
  }

  xml.closeNode(); // pict
  xml.closeNode(); // r
}

function renderTextWatermarkVml(xml: XmlSink, wm: TextWatermark): void {
  const color = wm.color ?? "C0C0C0";
  const font = wm.font ?? "Calibri";
  const rotation = wm.rotation ?? -45;

  // Full WordArt shapetype definition (t136) exactly as Microsoft Word
  // emits it. Word for Mac will NOT render the text path unless the
  // shapetype carries the formula/path/textpath/handles/lock children —
  // an empty <v:shapetype/> produces an invisible watermark.
  xml.openNode("v:shapetype", {
    id: "_x0000_t136",
    coordsize: "21600,21600",
    "o:spt": "136",
    adj: "10800",
    path: "m@7,l@8,m@5,21600l@6,21600e"
  });
  xml.openNode("v:formulas");
  for (const eqn of [
    "sum #0 0 10800",
    "prod #0 2 1",
    "sum 21600 0 @1",
    "sum 0 0 @2",
    "sum 21600 0 @3",
    "if @0 @3 0",
    "if @0 21600 @1",
    "if @0 0 @2",
    "if @0 @4 21600",
    "mid @5 @6",
    "mid @8 @5",
    "mid @7 @8",
    "mid @6 @7",
    "sum @6 0 @5"
  ]) {
    xml.leafNode("v:f", { eqn });
  }
  xml.closeNode(); // formulas
  xml.leafNode("v:path", {
    textpathok: "t",
    "o:connecttype": "custom",
    "o:connectlocs": "@9,0;@10,10800;@11,21600;@12,10800",
    "o:connectangles": "270,180,90,0"
  });
  xml.leafNode("v:textpath", { on: "t", fitshape: "t" });
  xml.openNode("v:handles");
  xml.leafNode("v:h", { position: "#0,bottomRight", xrange: "6629,14971" });
  xml.closeNode(); // handles
  xml.leafNode("o:lock", { "v:ext": "edit", text: "t", shapetype: "t" });
  xml.closeNode(); // shapetype

  // The watermark shape itself. Word stretches the text over the shape box
  // (fitshape="t"), so the box is the size: its own watermarks (the
  // Watermarks building blocks it ships) all say font-size:1pt and differ only
  // in width/height. A requested fontSize is therefore expressed as a box one
  // line of that size tall and as wide as the text, and written on the text
  // path as well so it reads back; rotation goes on the shape's style.
  const { widthPt, heightPt } = textWatermarkBox(wm, font);
  const style =
    `position:absolute;margin-left:0;margin-top:0;width:${formatPt(widthPt)}pt;height:${formatPt(heightPt)}pt;` +
    `${rotation ? `rotation:${rotation};` : ""}z-index:-251658752;` +
    `mso-position-horizontal:center;mso-position-horizontal-relative:margin;` +
    `mso-position-vertical:center;mso-position-vertical-relative:margin`;

  xml.openNode("v:shape", {
    id: "PowerPlusWaterMarkObject",
    "o:spid": "_x0000_s2049",
    type: "#_x0000_t136",
    alt: "",
    style,
    "o:allowincell": "f",
    fillcolor: `#${color}`,
    stroked: "f"
  });

  if (wm.semiTransparent !== false) {
    xml.leafNode("v:fill", { opacity: ".5" });
  }
  xml.leafNode("v:textpath", {
    style: `font-family:"${font}";font-size:${wm.fontSize !== undefined ? formatPt(wm.fontSize / 2) : "1"}pt`,
    string: wm.text
  });

  xml.closeNode(); // v:shape
}

/** Word's box for a watermark nobody sized: the width of a Letter page's text column, half as tall. */
const DEFAULT_TEXT_WATERMARK_BOX = { widthPt: 468, heightPt: 234 } as const;

function textWatermarkBox(
  wm: TextWatermark,
  font: string
): { readonly widthPt: number; readonly heightPt: number } {
  if (wm.widthPt !== undefined && wm.heightPt !== undefined) {
    return { widthPt: wm.widthPt, heightPt: wm.heightPt };
  }
  if (wm.fontSize === undefined) {
    return {
      widthPt: wm.widthPt ?? DEFAULT_TEXT_WATERMARK_BOX.widthPt,
      heightPt: wm.heightPt ?? DEFAULT_TEXT_WATERMARK_BOX.heightPt
    };
  }
  const sizePt = wm.fontSize / 2;
  return {
    widthPt: wm.widthPt ?? measureTextWidth(wm.text, mapToStandardFont(font), sizePt),
    heightPt: wm.heightPt ?? sizePt
  };
}

/** A length for a VML style: at most two decimals, no trailing zeros. */
function formatPt(value: number): string {
  return String(Math.round(value * 100) / 100);
}

function renderImageWatermarkVml(
  xml: XmlSink,
  wm: Watermark & { type: "image" },
  rId?: string
): void {
  const rid = rId ?? wm.rId;
  // Default to a large area covering most of the body so the picture is
  // actually visible. width:0;height:0 produces an invisible dot.
  const widthPt = wm.widthPt ?? 415.2;
  const heightPt = wm.heightPt ?? 233.5;

  // Picture-frame shapetype (t75) as Microsoft Word emits it. Without a
  // proper shapetype + non-zero size the image watermark will not render.
  xml.openNode("v:shapetype", {
    id: "_x0000_t75",
    coordsize: "21600,21600",
    "o:spt": "75",
    "o:preferrelative": "t",
    path: "m@4@5l@4@11@9@11@9@5xe",
    filled: "f",
    stroked: "f"
  });
  xml.openNode("v:stroke", { joinstyle: "miter" });
  xml.closeNode();
  xml.openNode("v:formulas");
  for (const eqn of [
    "if lineDrawn pixelLineWidth 0",
    "sum @0 1 0",
    "sum 0 0 @1",
    "prod @2 1 2",
    "prod @3 21600 pixelWidth",
    "prod @3 21600 pixelHeight",
    "sum @0 0 1",
    "prod @6 1 2",
    "prod @7 21600 pixelWidth",
    "sum @8 21600 0",
    "prod @7 21600 pixelHeight",
    "sum @10 21600 0"
  ]) {
    xml.leafNode("v:f", { eqn });
  }
  xml.closeNode(); // formulas
  xml.leafNode("v:path", {
    "o:extrusionok": "f",
    gradientshapeok: "t",
    "o:connecttype": "rect"
  });
  xml.leafNode("o:lock", { "v:ext": "edit", aspectratio: "t" });
  xml.closeNode(); // shapetype

  const style =
    `position:absolute;margin-left:0;margin-top:0;` +
    `width:${widthPt}pt;height:${heightPt}pt;` +
    `z-index:-251658752;mso-position-horizontal:center;` +
    `mso-position-horizontal-relative:margin;mso-position-vertical:center;` +
    `mso-position-vertical-relative:margin`;

  xml.openNode("v:shape", {
    id: "PowerPlusWaterMarkObject",
    "o:spid": "_x0000_s2050",
    type: "#_x0000_t75",
    alt: "",
    style,
    "o:allowincell": "f"
  });

  const gain = wm.washout !== false ? "19661f" : "65536f";
  const blacklevel = wm.washout !== false ? "22938f" : "0";

  xml.leafNode("v:imagedata", {
    "r:id": rid,
    "o:title": "",
    gain,
    blacklevel
  });
  xml.leafNode("w10:wrap", { anchorx: "margin", anchory: "margin" });

  xml.closeNode(); // v:shape
}
