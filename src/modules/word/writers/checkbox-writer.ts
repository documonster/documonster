/**
 * DOCX Writers - CheckBox
 *
 * Renders checkbox as a structured document tag (SDT) with w14:checkbox.
 */

import type { XmlSink } from "@xml/types";

import { NS_W14 } from "../constants";
import type { CheckBox } from "../types";

/** Render a checkbox as a paragraph with SDT. */
export function renderCheckBox(xml: XmlSink, cb: CheckBox): void {
  const checked = cb.checked ?? false;
  const checkedChar = cb.checkedState?.value ?? "2612"; // ☒
  const uncheckedChar = cb.uncheckedState?.value ?? "2610"; // ☐
  const checkedFont = cb.checkedState?.font ?? "MS Gothic";
  const uncheckedFont = cb.uncheckedState?.font ?? "MS Gothic";
  const displayChar = checked ? checkedChar : uncheckedChar;
  const displayFont = checked ? checkedFont : uncheckedFont;

  xml.openNode("w:sdt");

  // SDT properties
  xml.openNode("w:sdtPr");
  xml.openNode("w14:checkbox", { "xmlns:w14": NS_W14 });
  xml.leafNode("w14:checked", { "w14:val": checked ? "1" : "0" });
  xml.leafNode("w14:checkedState", { "w14:val": checkedChar, "w14:font": checkedFont });
  xml.leafNode("w14:uncheckedState", { "w14:val": uncheckedChar, "w14:font": uncheckedFont });
  xml.closeNode();
  xml.closeNode();

  // SDT content - display the checkbox character
  xml.openNode("w:sdtContent");
  xml.openNode("w:r");
  xml.openNode("w:rPr");
  xml.leafNode("w:rFonts", {
    "w:ascii": displayFont,
    "w:hAnsi": displayFont,
    "w:eastAsia": displayFont
  });
  xml.closeNode();
  xml.openNode("w:t");
  // Convert hex code point to character
  xml.writeText(String.fromCodePoint(parseInt(displayChar, 16)));
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();

  xml.closeNode(); // w:sdt
}
