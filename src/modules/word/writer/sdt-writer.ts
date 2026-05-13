/**
 * DOCX Writers - Structured Document Tag (SDT) Properties
 *
 * This file intentionally only renders `<w:sdtPr>` — the property block — so
 * it has no dependency on the block-level renderers (paragraph/table/run/
 * nested SDT). Callers are responsible for opening `<w:sdt>` and
 * `<w:sdtContent>` themselves and for dispatching child blocks using their
 * own render functions, which keeps the writer module graph acyclic.
 */

import type { XmlSink } from "@xml/types";

import type { StructuredDocumentTag } from "../types";
import type { WordRenderContext } from "./render-context";

/**
 * Render the `<w:sdt>` ↦ `<w:sdtPr>` block for a given SDT.
 *
 * Does not emit `<w:sdt>` itself nor `<w:sdtContent>` — that is the caller's
 * responsibility. Designed so block-level renderers (table cell, header/
 * footer, document body) can share the property emission logic without
 * forming an import cycle through child-block dispatch.
 *
 * `ctx` may be undefined when called from a renderer that only owns a
 * `RenderHelpers`. In that case we cannot draw an auto-id from the document
 * id generator and will simply omit `<w:id>` if the model didn't supply one.
 */
export function renderSdtPr(
  xml: XmlSink,
  sdt: StructuredDocumentTag,
  ctx?: WordRenderContext
): void {
  if (!sdt.properties) {
    return;
  }

  xml.openNode("w:sdtPr");
  const p = sdt.properties;

  // Auto-assign ID if not provided (Word strongly expects unique IDs).
  // When invoked without a ctx we skip the w:id when the model didn't
  // supply one — emitting nothing is safer than reusing a colliding id.
  const sdtId = p.id ?? ctx?.ids.nextSdtId();
  if (sdtId !== undefined) {
    xml.leafNode("w:id", { "w:val": String(sdtId) });
  }
  if (p.tag) {
    xml.leafNode("w:tag", { "w:val": p.tag });
  }
  if (p.alias) {
    xml.leafNode("w:alias", { "w:val": p.alias });
  }
  // Lock: combine lockContent + lockSdt
  if (p.lockContent && p.lockSdt) {
    xml.leafNode("w:lock", { "w:val": "sdtContentLocked" });
  } else if (p.lockContent) {
    xml.leafNode("w:lock", { "w:val": "contentLocked" });
  } else if (p.lockSdt) {
    xml.leafNode("w:lock", { "w:val": "sdtLocked" });
  }
  // Appearance (w15:appearance for modern SDTs)
  if (p.appearance) {
    xml.leafNode("w15:appearance", { "w15:val": p.appearance });
  }
  // Show placeholder toggle
  if (p.showingPlaceholder) {
    xml.leafNode("w:showingPlcHdr");
  }
  if (p.placeholder) {
    xml.openNode("w:placeholder");
    xml.leafNode("w:docPart", { "w:val": p.placeholder });
    xml.closeNode();
  }
  if (p.temporary) {
    xml.leafNode("w:temporary");
  }
  if (p.dataBinding) {
    const bindAttrs: Record<string, string> = {
      "w:xpath": p.dataBinding.xpath,
      "w:storeItemID": p.dataBinding.storeItemId
    };
    if (p.dataBinding.prefixMappings) {
      bindAttrs["w:prefixMappings"] = p.dataBinding.prefixMappings;
    }
    xml.leafNode("w:dataBinding", bindAttrs);
  }
  // Type discriminator markers
  if (p.plainText) {
    xml.leafNode("w:text");
  }
  if (p.richText) {
    xml.leafNode("w:richText");
  }
  if (p.picture) {
    xml.leafNode("w:picture");
  }
  if (p.group) {
    xml.leafNode("w:group");
  }
  if (p.equation) {
    xml.leafNode("w:equation");
  }
  if (p.citation) {
    xml.leafNode("w:citation");
  }
  if (p.bibliography) {
    xml.leafNode("w:bibliography");
  }
  if (p.repeatingSectionItem) {
    xml.leafNode("w15:repeatingSectionItem");
  }
  if (p.repeatingSection) {
    // w15:repeatingSection has child elements (not attributes) per the schema
    const rs = p.repeatingSection;
    const hasChildren = rs.sectionTitle !== undefined || rs.allowInsertDelete !== undefined;
    if (hasChildren) {
      xml.openNode("w15:repeatingSection");
      if (rs.sectionTitle !== undefined) {
        xml.leafNode("w15:sectionTitle", { "w15:val": rs.sectionTitle });
      }
      if (rs.allowInsertDelete === false) {
        xml.leafNode("w15:doNotAllowInsertDeleteSection");
      }
      xml.closeNode();
    } else {
      xml.leafNode("w15:repeatingSection");
    }
  }
  // Checkbox (w14 extension)
  if (p.checkbox) {
    xml.openNode("w14:checkbox");
    xml.leafNode("w14:checked", p.checkbox.checked ? { "w14:val": "1" } : { "w14:val": "0" });
    if (p.checkbox.checkedChar || p.checkbox.checkedFont) {
      const cAttrs: Record<string, string> = {};
      if (p.checkbox.checkedChar) {
        cAttrs["w14:val"] = p.checkbox.checkedChar;
      }
      if (p.checkbox.checkedFont) {
        cAttrs["w14:font"] = p.checkbox.checkedFont;
      }
      xml.leafNode("w14:checkedState", cAttrs);
    }
    if (p.checkbox.uncheckedChar || p.checkbox.uncheckedFont) {
      const uAttrs: Record<string, string> = {};
      if (p.checkbox.uncheckedChar) {
        uAttrs["w14:val"] = p.checkbox.uncheckedChar;
      }
      if (p.checkbox.uncheckedFont) {
        uAttrs["w14:font"] = p.checkbox.uncheckedFont;
      }
      xml.leafNode("w14:uncheckedState", uAttrs);
    }
    xml.closeNode();
  }
  // Dropdown list
  if (p.dropdownList) {
    xml.openNode("w:dropDownList");
    for (const item of p.dropdownList) {
      const attrs: Record<string, string> = { "w:value": item.value };
      if (item.displayText) {
        attrs["w:displayText"] = item.displayText;
      }
      xml.leafNode("w:listItem", attrs);
    }
    xml.closeNode();
  }
  // ComboBox
  if (p.comboBox) {
    xml.openNode("w:comboBox");
    for (const item of p.comboBox) {
      const attrs: Record<string, string> = { "w:value": item.value };
      if (item.displayText) {
        attrs["w:displayText"] = item.displayText;
      }
      xml.leafNode("w:listItem", attrs);
    }
    xml.closeNode();
  }
  // Date picker
  if (p.date) {
    const dateAttrs: Record<string, string> = {};
    if (p.date.fullDate) {
      dateAttrs["w:fullDate"] = p.date.fullDate;
    }
    xml.openNode("w:date", Object.keys(dateAttrs).length > 0 ? dateAttrs : undefined);
    if (p.date.dateFormat) {
      xml.leafNode("w:dateFormat", { "w:val": p.date.dateFormat });
    }
    if (p.date.lid) {
      xml.leafNode("w:lid", { "w:val": p.date.lid });
    }
    if (p.date.storeMappedDataAs) {
      xml.leafNode("w:storeMappedDataAs", { "w:val": p.date.storeMappedDataAs });
    }
    xml.closeNode();
  }

  xml.closeNode(); // w:sdtPr
}
