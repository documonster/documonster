/**
 * DOCX Writers - CheckBox
 *
 * Renders checkbox as a structured document tag (SDT) with w14:checkbox.
 *
 * Per ECMA-376 Part 4 §17.5.2.41 the `w14:checkedState/@w14:val` attribute
 * holds a 4-digit hexadecimal Unicode code point. We persist that hex form
 * verbatim into the `w14:val` attribute, but the displayed glyph is rendered
 * by converting to the actual character. To stay forgiving we accept both
 * the canonical hex form (e.g. `"2612"`) and a literal glyph form (e.g.
 * `"☒"`) on the input side, normalising both to a hex string for the
 * attribute and the actual glyph for the displayed run.
 */

import type { XmlSink } from "@xml/types";

import { NS_W14 } from "../constants";
import type { CheckBox } from "../types";

/**
 * Normalise a checked-state value to {hex, glyph}.
 *
 *   - If the input parses as a hexadecimal code point (1-6 hex digits with
 *     no other characters) we treat it as the canonical hex form.
 *   - Otherwise we treat it as a literal glyph and derive its hex from the
 *     first code point of the string.
 *
 * Empty / nullish input falls back to the supplied default.
 */
function normaliseCheckState(
  input: string | undefined,
  fallback: string
): { hex: string; glyph: string } {
  const value = input && input.length > 0 ? input : fallback;
  // Canonical hex form: 1-6 hex digits and nothing else.
  if (/^[0-9a-fA-F]{1,6}$/.test(value)) {
    const cp = parseInt(value, 16);
    if (Number.isFinite(cp) && cp >= 0 && cp <= 0x10ffff) {
      return { hex: value.toUpperCase().padStart(4, "0"), glyph: String.fromCodePoint(cp) };
    }
  }
  // Literal glyph form: take the first Unicode code point.
  const cp = value.codePointAt(0);
  if (cp === undefined) {
    // Should not happen because fallback is always non-empty, but stay safe.
    return { hex: "2610", glyph: "\u2610" };
  }
  return {
    hex: cp.toString(16).toUpperCase().padStart(4, "0"),
    glyph: String.fromCodePoint(cp)
  };
}

/** Render a checkbox as a paragraph with SDT. */
export function renderCheckBox(xml: XmlSink, cb: CheckBox): void {
  const checked = cb.checked ?? false;
  const checkedNorm = normaliseCheckState(cb.checkedState?.value, "2612"); // ☒
  const uncheckedNorm = normaliseCheckState(cb.uncheckedState?.value, "2610"); // ☐
  const checkedFont = cb.checkedState?.font ?? "MS Gothic";
  const uncheckedFont = cb.uncheckedState?.font ?? "MS Gothic";
  const display = checked ? checkedNorm : uncheckedNorm;
  const displayFont = checked ? checkedFont : uncheckedFont;

  xml.openNode("w:sdt");

  // SDT properties
  xml.openNode("w:sdtPr");
  xml.openNode("w14:checkbox", { "xmlns:w14": NS_W14 });
  xml.leafNode("w14:checked", { "w14:val": checked ? "1" : "0" });
  xml.leafNode("w14:checkedState", { "w14:val": checkedNorm.hex, "w14:font": checkedFont });
  xml.leafNode("w14:uncheckedState", { "w14:val": uncheckedNorm.hex, "w14:font": uncheckedFont });
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
  xml.writeText(display.glyph);
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();

  xml.closeNode(); // w:sdt
}
