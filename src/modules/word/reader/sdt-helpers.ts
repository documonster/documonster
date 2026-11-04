/**
 * DOCX Reader - SDT Helper Parsers
 *
 * Standalone parsers for SDT-related elements that don't depend on
 * the main reader context (parseParagraph etc.):
 * - `parseCheckBox` — w14:checkbox SDT specialization
 */

import type { Mutable } from "@word/core/internal-utils";
import type { CheckBox } from "@word/types";
import { findChild } from "@xml/dom";
import type { XmlElement } from "@xml/types";

/**
 * Parse a w14:checkbox SDT element.
 *
 * Handles `w14:checked`, `w14:checkedState`, `w14:uncheckedState`.
 */
export function parseCheckBox(checkBoxEl: XmlElement): CheckBox {
  const cb: Mutable<CheckBox> = { type: "checkBox" };
  const checkedEl = findChild(checkBoxEl, "w14:checked");
  if (checkedEl) {
    const v = checkedEl.attributes["w14:val"] ?? checkedEl.attributes["val"];
    cb.checked = v === "1" || v === "true";
  }
  const checkedStateEl = findChild(checkBoxEl, "w14:checkedState");
  if (checkedStateEl) {
    cb.checkedState = {
      value: checkedStateEl.attributes["w14:val"] ?? checkedStateEl.attributes["val"] ?? "",
      font: checkedStateEl.attributes["w14:font"] ?? checkedStateEl.attributes["font"]
    };
  }
  const uncheckedStateEl = findChild(checkBoxEl, "w14:uncheckedState");
  if (uncheckedStateEl) {
    cb.uncheckedState = {
      value: uncheckedStateEl.attributes["w14:val"] ?? uncheckedStateEl.attributes["val"] ?? "",
      font: uncheckedStateEl.attributes["w14:font"] ?? uncheckedStateEl.attributes["font"]
    };
  }
  return cb;
}
