/**
 * DOCX Writers - Structured Document Tag (SDT) Properties
 *
 * This file intentionally only renders `<w:sdtPr>` — the property block — so
 * it has no dependency on the block-level renderers (paragraph/table/run/
 * nested SDT). Callers are responsible for opening `<w:sdt>` and
 * `<w:sdtContent>` themselves and for dispatching child blocks using their
 * own render functions, which keeps the writer module graph acyclic.
 */

import type { SdtProperties, StructuredDocumentTag } from "@word/types";
import { normaliseCheckState } from "@word/writer/checkbox-writer";
import type { DocxRenderContext } from "@word/writer/render-context";
import { emitInSchemaOrder } from "@word/writer/schema-order";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

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
  ctx?: DocxRenderContext
): void {
  renderSdtProperties(xml, sdt.properties, ctx);
}

/** {@link renderSdtPr} for a control's properties alone (a block-level control in a table cell). */
export function renderSdtProperties(
  xml: XmlSink,
  p: SdtProperties | undefined,
  ctx?: DocxRenderContext
): void {
  if (!p) {
    return;
  }

  xml.openNode("w:sdtPr");

  // Auto-assign ID if not provided (Word strongly expects unique IDs).
  // When invoked without a ctx we skip the w:id when the model didn't
  // supply one — emitting nothing is safer than reusing a colliding id.
  const sdtId = p.id ?? ctx?.ids.nextSdtId();
  const { dataBinding, docPart, repeatingSection, checkbox, dropdownList, comboBox, date } = p;

  // In `CT_SdtPr` order: alias, tag, id, lock, placeholder, temporary, showingPlcHdr, dataBinding, then the type.
  // This wrote the id first and the placeholder flag before the placeholder. The `w14:`/`w15:` extensions stay
  // beside the standard child they follow (`emitInSchemaOrder`).
  emitInSchemaOrder("w:CT_SdtPr", [
    !!p.alias && ["w:alias", () => xml.leafNode("w:alias", { "w:val": p.alias! })],
    !!p.tag && ["w:tag", () => xml.leafNode("w:tag", { "w:val": p.tag! })],
    sdtId !== undefined && [
      "w:id",
      () => xml.leafNode("w:id", { "w:val": xmlValue(sdtId, "w:val") })
    ],
    // Lock: combine lockContent + lockSdt
    !!(p.lockContent || p.lockSdt) && [
      "w:lock",
      () =>
        xml.leafNode("w:lock", {
          "w:val":
            p.lockContent && p.lockSdt
              ? "sdtContentLocked"
              : p.lockContent
                ? "contentLocked"
                : "sdtLocked"
        })
    ],
    // Appearance (w15:appearance for modern SDTs)
    !!p.appearance && [
      "w15:appearance",
      () => xml.leafNode("w15:appearance", { "w15:val": p.appearance! })
    ],
    !!p.placeholder && [
      "w:placeholder",
      () => {
        xml.openNode("w:placeholder");
        xml.leafNode("w:docPart", { "w:val": p.placeholder! });
        xml.closeNode();
      }
    ],
    !!p.temporary && ["w:temporary", () => xml.leafNode("w:temporary")],
    // Show placeholder toggle
    !!p.showingPlaceholder && ["w:showingPlcHdr", () => xml.leafNode("w:showingPlcHdr")],
    !!dataBinding && [
      "w:dataBinding",
      () => {
        const bindAttrs: Record<string, string> = {
          "w:xpath": dataBinding.xpath,
          "w:storeItemID": dataBinding.storeItemId
        };
        if (dataBinding.prefixMappings) {
          bindAttrs["w:prefixMappings"] = dataBinding.prefixMappings;
        }
        xml.leafNode("w:dataBinding", bindAttrs);
      }
    ],
    // Type discriminator markers — one `xsd:choice`, so they share a position and keep this order.
    !!p.plainText && ["w:text", () => xml.leafNode("w:text")],
    !!p.richText && ["w:richText", () => xml.leafNode("w:richText")],
    !!p.picture && ["w:picture", () => xml.leafNode("w:picture")],
    !!p.group && ["w:group", () => xml.leafNode("w:group")],
    !!p.equation && ["w:equation", () => xml.leafNode("w:equation")],
    !!p.citation && ["w:citation", () => xml.leafNode("w:citation")],
    !!p.bibliography && ["w:bibliography", () => xml.leafNode("w:bibliography")],
    !!docPart && [
      docPart.list ? "w:docPartList" : "w:docPartObj",
      () => {
        xml.openNode(docPart.list ? "w:docPartList" : "w:docPartObj");
        if (docPart.gallery !== undefined) {
          xml.leafNode("w:docPartGallery", { "w:val": docPart.gallery });
        }
        if (docPart.category !== undefined) {
          xml.leafNode("w:docPartCategory", { "w:val": docPart.category });
        }
        if (docPart.unique) {
          xml.leafNode("w:docPartUnique");
        }
        xml.closeNode();
      }
    ],
    !!p.repeatingSectionItem && [
      "w15:repeatingSectionItem",
      () => xml.leafNode("w15:repeatingSectionItem")
    ],
    !!repeatingSection && [
      "w15:repeatingSection",
      () => {
        // w15:repeatingSection has child elements (not attributes) per the schema
        const rs = repeatingSection;
        const hasChildren = rs.sectionTitle !== undefined || rs.allowInsertDelete !== undefined;
        if (hasChildren) {
          xml.openNode("w15:repeatingSection");
          emitInSchemaOrder("w15:CT_SdtRepeatedSection", [
            rs.sectionTitle !== undefined && [
              "w15:sectionTitle",
              // `w12:CT_String`, whose `w12` is WordprocessingML's main namespace: `w:val`, not `w15:val`.
              () => xml.leafNode("w15:sectionTitle", { "w:val": rs.sectionTitle! })
            ],
            rs.allowInsertDelete === false && [
              "w15:doNotAllowInsertDeleteSection",
              () => xml.leafNode("w15:doNotAllowInsertDeleteSection")
            ]
          ]);
          xml.closeNode();
        } else {
          xml.leafNode("w15:repeatingSection");
        }
      }
    ],
    // Checkbox (w14 extension)
    !!checkbox && [
      "w14:checkbox",
      () => {
        xml.openNode("w14:checkbox");
        const state = (char: string | undefined, font: string | undefined) => {
          const attrs: Record<string, string> = {};
          if (char) {
            attrs["w14:val"] = checkStateHex(char);
          }
          if (font) {
            attrs["w14:font"] = font;
          }
          return attrs;
        };
        // `CT_SdtCheckbox` ([MS-DOCX] §5.1): checked, checkedState, uncheckedState.
        emitInSchemaOrder("w14:CT_SdtCheckbox", [
          [
            "w14:checked",
            () => xml.leafNode("w14:checked", { "w14:val": checkbox.checked ? "1" : "0" })
          ],
          !!(checkbox.checkedChar || checkbox.checkedFont) && [
            "w14:checkedState",
            () =>
              xml.leafNode("w14:checkedState", state(checkbox.checkedChar, checkbox.checkedFont))
          ],
          !!(checkbox.uncheckedChar || checkbox.uncheckedFont) && [
            "w14:uncheckedState",
            () =>
              xml.leafNode(
                "w14:uncheckedState",
                state(checkbox.uncheckedChar, checkbox.uncheckedFont)
              )
          ]
        ]);
        xml.closeNode();
      }
    ],
    // Dropdown list
    !!dropdownList && [
      "w:dropDownList",
      () => renderListItems(xml, "w:dropDownList", dropdownList)
    ],
    // ComboBox
    !!comboBox && ["w:comboBox", () => renderListItems(xml, "w:comboBox", comboBox)],
    // Date picker
    !!date && [
      "w:date",
      () => {
        const dateAttrs: Record<string, string> = {};
        if (date.fullDate) {
          dateAttrs["w:fullDate"] = date.fullDate;
        }
        xml.openNode("w:date", Object.keys(dateAttrs).length > 0 ? dateAttrs : undefined);
        if (date.dateFormat) {
          xml.leafNode("w:dateFormat", { "w:val": date.dateFormat });
        }
        if (date.lid) {
          xml.leafNode("w:lid", { "w:val": date.lid });
        }
        if (date.storeMappedDataAs) {
          xml.leafNode("w:storeMappedDataAs", { "w:val": date.storeMappedDataAs });
        }
        xml.closeNode();
      }
    ]
  ]);

  xml.closeNode(); // w:sdtPr
}

/** A drop-down list's or a combo box's items. */
function renderListItems(
  xml: XmlSink,
  tagName: "w:dropDownList" | "w:comboBox",
  items: NonNullable<SdtProperties["dropdownList"]>
): void {
  xml.openNode(tagName);
  for (const item of items) {
    const attrs: Record<string, string> = { "w:value": item.value };
    if (item.displayText) {
      attrs["w:displayText"] = item.displayText;
    }
    xml.leafNode("w:listItem", attrs);
  }
  xml.closeNode();
}

/**
 * `w14:val` of a check state is a hex code point (`ST_ShortHexNumber`). One
 * already in that form is written as it is; a glyph (`"☒"`) — which the
 * checkbox builder accepts too — is written as its code point rather than
 * verbatim, which Word refused.
 */
function checkStateHex(value: string): string {
  return /^[0-9A-Fa-f]{4}$/.test(value) ? value : normaliseCheckState(value, value).hex;
}
