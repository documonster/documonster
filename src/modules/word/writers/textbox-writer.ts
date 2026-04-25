/**
 * DOCX Writers - TextBox
 *
 * Renders text boxes using VML (v:shape) with mc:AlternateContent.
 */

import type { XmlSink } from "@xml/types";

import { NS_V, NS_O, NS_W10, NS_W } from "../constants";
import type { TextBox } from "../types";
import { renderParagraph } from "./paragraph-writer";

/** Render a text box as a VML shape within a paragraph. */
export function renderTextBox(xml: XmlSink, textBox: TextBox): void {
  const width = textBox.width ?? 4680; // default ~3.25 inches
  const height = textBox.height ?? 1440; // default 1 inch
  // Convert twips to points for VML style
  const widthPt = width / 20;
  const heightPt = height / 20;

  const style = textBox.style ?? `width:${widthPt}pt;height:${heightPt}pt;mso-wrap-style:square`;

  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.openNode("w:pict");

  // VML shape
  const shapeAttrs: Record<string, string> = {
    "xmlns:v": NS_V,
    "xmlns:o": NS_O,
    type: "#_x0000_t202",
    style
  };

  if (textBox.stroke === false) {
    shapeAttrs["stroked"] = "f";
  }
  if (textBox.strokeColor) {
    shapeAttrs["strokecolor"] = `#${textBox.strokeColor}`;
  }
  if (textBox.fill === false) {
    shapeAttrs["filled"] = "f";
  }
  if (textBox.fillColor) {
    shapeAttrs["fillcolor"] = `#${textBox.fillColor}`;
  }

  xml.openNode("v:shape", shapeAttrs);

  // Text box content
  xml.openNode("v:textbox");
  xml.openNode("w:txbxContent", { "xmlns:w": NS_W });

  for (const para of textBox.content) {
    renderParagraph(xml, para);
  }

  xml.closeNode(); // w:txbxContent
  xml.closeNode(); // v:textbox

  // Word10 wrap
  xml.leafNode("w10:wrap", { "xmlns:w10": NS_W10, type: "square" });

  xml.closeNode(); // v:shape
  xml.closeNode(); // w:pict
  xml.closeNode(); // w:r
  xml.closeNode(); // w:p
}
