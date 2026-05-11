/**
 * DOCX Reader - SDT Helper Parsers
 *
 * Standalone parsers for SDT-related elements that don't depend on
 * the main reader context (parseParagraph etc.):
 * - `parseCheckBox` — w14:checkbox SDT specialization
 * - `parseTocInstruction` — TOC field instruction string parser
 */

import { findChild } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import { type Mutable } from "../core/internal-utils";
import type { CheckBox, TableOfContents } from "../types";

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

/**
 * Parse a TOC field instruction string and populate the given TableOfContents object.
 *
 * Recognizes common TOC switches:
 * - `\o "1-3"` — heading style range
 * - `\h` — hyperlinks
 * - `\c "label"` — caption label (table of figures)
 * - `\s "id"` — sequence field identifier
 * - `\p "."`/`"-"`/`"_"` — leader character
 * - `\t "Style,Level;..."` — custom styles with levels
 */
export function parseTocInstruction(instr: string, toc: Mutable<TableOfContents>): void {
  const trimmed = instr.trim();
  if (!/^TOC\b/i.test(trimmed)) {
    return;
  }
  // Match switches: \<letter> followed by either "quoted" or non-quoted non-switch token.
  // The next-switch boundary must be respected: an unquoted value cannot start with \.
  const switchRe = /\\(\w)(?:\s+"([^"]*)"|\s+([^\\\s][^\s]*))?/g;
  let match: RegExpExecArray | null;
  while ((match = switchRe.exec(trimmed)) !== null) {
    const switchName = match[1].toLowerCase();
    const value = match[2] ?? match[3];
    switch (switchName) {
      case "o": // Heading level range e.g. "1-3"
        if (value) {
          toc.headingStyleRange = value;
        }
        break;
      case "h": // Hyperlinks
        toc.hyperlink = true;
        break;
      case "c": // Caption label (table of figures)
        if (value) {
          toc.captionLabel = value;
        }
        break;
      case "s": // Sequence field identifier
        if (value) {
          toc.sequenceFieldIdentifier = value;
        }
        break;
      case "p": // Page-number leader or style separator
        if (value === "." || value === "-" || value === "_") {
          toc.leader = "dot";
          if (value === "-") {
            toc.leader = "hyphen";
          } else if (value === "_") {
            toc.leader = "underscore";
          }
        }
        break;
      case "t": {
        // Styles with levels: "StyleName1,Level1;StyleName2,Level2;..."
        if (!value) {
          break;
        }
        const items: { styleName: string; level: number }[] = [];
        for (const part of value.split(";")) {
          const [styleName, levelStr] = part.split(",");
          if (styleName && levelStr) {
            items.push({ styleName: styleName.trim(), level: parseInt(levelStr, 10) });
          }
        }
        if (items.length > 0) {
          toc.stylesWithLevels = items;
        }
        break;
      }
    }
  }
}
