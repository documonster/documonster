/**
 * DOCX Reader - Form Field Parser
 *
 * Standalone parser for `w:ffData` elements (legacy form fields).
 * Used by the run content parser when encountering form-field-bearing fields.
 */

import { attrInt, attrVal, findChildNs, findChildrenNs } from "@word/reader/parse-utils";
import type { FormField } from "@word/types";
import type { XmlElement } from "@xml/types";

/**
 * Parse a `w:ffData` element into a `FormField`.
 *
 * Supports the three legacy form field types:
 * - `w:textInput` → text field
 * - `w:checkBox` → checkbox field
 * - `w:ddList` → drop-down list field
 *
 * Common properties (name, enabled, helpText, statusText) are extracted.
 *
 * @returns A `FormField` object, or undefined if no recognized field type was found.
 */
export function parseFfData(el: XmlElement): FormField | undefined {
  const nameEl = findChildNs(el, "name");
  const name = nameEl ? attrVal(nameEl, "val") : undefined;
  const enabledEl = findChildNs(el, "enabled");
  const enabled = enabledEl ? attrVal(enabledEl, "val") !== "0" : undefined;
  const helpTextEl = findChildNs(el, "helpText");
  const helpText = helpTextEl ? attrVal(helpTextEl, "val") : undefined;
  const statusTextEl = findChildNs(el, "statusText");
  const statusText = statusTextEl ? attrVal(statusTextEl, "val") : undefined;

  // Text input
  const textInputEl = findChildNs(el, "textInput");
  if (textInputEl) {
    const defEl = findChildNs(textInputEl, "default");
    const maxLenEl = findChildNs(textInputEl, "maxLength");
    const fmtEl = findChildNs(textInputEl, "format");
    return {
      type: "text",
      name,
      default: defEl ? attrVal(defEl, "val") : undefined,
      maxLength: maxLenEl ? attrInt(maxLenEl, "val") : undefined,
      format: fmtEl ? attrVal(fmtEl, "val") : undefined,
      helpText,
      statusText,
      enabled
    };
  }

  // CheckBox
  const cbEl = findChildNs(el, "checkBox");
  if (cbEl) {
    const checkedEl = findChildNs(cbEl, "checked");
    const defEl = findChildNs(cbEl, "default");
    const sizeEl = findChildNs(cbEl, "size");
    return {
      type: "checkBox",
      name,
      checked: checkedEl ? attrVal(checkedEl, "val") !== "0" : undefined,
      default: defEl ? attrVal(defEl, "val") !== "0" : undefined,
      size: sizeEl ? attrInt(sizeEl, "val") : undefined
    };
  }

  // Drop-down list
  const ddlEl = findChildNs(el, "ddList");
  if (ddlEl) {
    const defEl = findChildNs(ddlEl, "default");
    const entries: string[] = [];
    for (const le of findChildrenNs(ddlEl, "listEntry")) {
      const v = attrVal(le, "val");
      if (v !== undefined) {
        entries.push(v);
      }
    }
    return {
      type: "dropDown",
      name,
      entries: entries.length > 0 ? entries : undefined,
      default: defEl ? attrInt(defEl, "val") : undefined,
      helpText,
      statusText,
      enabled
    };
  }

  return undefined;
}
