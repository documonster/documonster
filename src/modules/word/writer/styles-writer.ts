/**
 * DOCX Writers - Styles
 *
 * Renders word/styles.xml including document defaults, style definitions,
 * and table style conditional formats.
 */

import { NS_W, NS_R, STD_DOC_ATTRIBUTES, TABLE_STYLE_CONDITION_TOKENS } from "@word/constants";
import {
  NARROWER_PROPERTIES,
  type TableStyleProperties,
  tableStyleExcludedFields
} from "@word/core/change-records";
import { refuseAbsentFields, unwritableNode, within } from "@word/core/node-path";
import type { DocDefaults, PreservedXml, StyleDef, TableStyleConditionalFormat } from "@word/types";
import { renderParagraphProperties } from "@word/writer/paragraph-writer";
import {
  createPreservedEmitter,
  preservedElements,
  schemaOrder,
  wordprocessingNames,
  withPreservedNamespaces
} from "@word/writer/preserved-xml";
import { renderRunProperties } from "@word/writer/run-writer";
import { emitInSchemaOrder } from "@word/writer/schema-order";
import {
  renderTableCellProperties,
  renderTableProperties,
  renderTableRowProperties
} from "@word/writer/table-writer";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

/** Render document defaults. */
function renderDocDefaults(xml: XmlSink, defaults: DocDefaults): void {
  xml.openNode("w:docDefaults");

  if (defaults.runProperties) {
    xml.openNode("w:rPrDefault");
    const run = defaults.runProperties;
    within("docDefaults.runProperties", () => renderRunProperties(xml, run));
    xml.closeNode();
  }

  if (defaults.paragraphProperties) {
    xml.openNode("w:pPrDefault");
    const paragraph = defaults.paragraphProperties;
    within("docDefaults.paragraphProperties", () =>
      renderParagraphProperties(xml, paragraph, "general")
    );
    xml.closeNode();
  }

  xml.closeNode();
}

/** Map table style condition type to OOXML value. */
/**
 * Render table style conditional format (w:tblStylePr).
 *
 * Its table, row and cell properties are written by the table writer — they are the same `TableProperties`,
 * `TableRowProperties` and `TableCellProperties` a table carries. This used to have a reduced copy of each: the
 * table's carried only its borders and the row's was not written at all, so a conditional format read from a file
 * lost its shading, margins and row height on the way back; and its borders dropped `themeColor`, `shadow` and
 * `frame`, and invented a size and a colour the model did not have.
 */
function renderTableStyleCondition(xml: XmlSink, condition: TableStyleConditionalFormat): void {
  const typeVal = Object.hasOwn(TABLE_STYLE_CONDITION_TOKENS, condition.type)
    ? TABLE_STYLE_CONDITION_TOKENS[condition.type]
    : unwritableNode(
        "type",
        `one of ${Object.keys(TABLE_STYLE_CONDITION_TOKENS)
          .map(name => JSON.stringify(name))
          .join(", ")}`,
        condition.type
      );
  xml.openNode("w:tblStylePr", { "w:type": typeVal });
  const { paragraphProperties, runProperties, tableProperties, rowProperties, cellProperties } =
    condition;
  emitInSchemaOrder("w:CT_TblStylePr", [
    !!paragraphProperties && [
      "w:pPr",
      () =>
        within("paragraphProperties", () =>
          renderParagraphProperties(xml, paragraphProperties, "general")
        )
    ],
    !!runProperties && [
      "w:rPr",
      () => within("runProperties", () => renderRunProperties(xml, runProperties))
    ],
    // What Word does not allow in a table style is refused (`TABLE_STYLE_EXCLUDED`).
    !!tableProperties && [
      "w:tblPr",
      () =>
        within("tableProperties", () => {
          refuseExcluded(tableProperties, "conditionTable");
          renderTableProperties(xml, tableProperties, true);
        })
    ],
    !!rowProperties && [
      "w:trPr",
      () =>
        within("rowProperties", () => {
          refuseExcluded(rowProperties, "conditionRow");
          renderTableRowProperties(xml, rowProperties);
        })
    ],
    !!cellProperties && [
      "w:tcPr",
      () =>
        within("cellProperties", () => {
          refuseExcluded(cellProperties, "conditionCell");
          renderTableCellProperties(xml, cellProperties, { widthPlaceholder: false });
        })
    ]
  ]);
  xml.closeNode();
}

/** CT_Style's children, from the schema. */
const STYLE_ORDER = schemaOrder(wordprocessingNames("w:CT_Style"));

/** CT_Styles' children, from the schema. */
const STYLES_ORDER = schemaOrder(wordprocessingNames("w:CT_Styles"));

type RawXmlPolicy = "preserve" | "strip" | "reject" | undefined;

/** Render a single style definition. */
function renderStyle(xml: XmlSink, style: StyleDef, rawXmlPolicy: RawXmlPolicy): void {
  const preserved = createPreservedEmitter(
    xml,
    STYLE_ORDER,
    preservedElements(style.preservedXml, rawXmlPolicy, "style.preservedXml")
  );
  const attrs: Record<string, string> = {
    "w:type": style.type,
    "w:styleId": style.styleId
  };
  if (style.isDefault) {
    attrs["w:default"] = "1";
  }
  if (style.customStyle) {
    attrs["w:customStyle"] = "1";
  }
  xml.openNode("w:style", attrs);

  xml.leafNode("w:name", { "w:val": style.name });

  if (style.basedOn) {
    preserved.upTo("w:basedOn");
    xml.leafNode("w:basedOn", { "w:val": style.basedOn });
  }
  if (style.next) {
    preserved.upTo("w:next");
    xml.leafNode("w:next", { "w:val": style.next });
  }
  if (style.link) {
    preserved.upTo("w:link");
    xml.leafNode("w:link", { "w:val": style.link });
  }
  if (style.autoRedefine) {
    preserved.upTo("w:autoRedefine");
    xml.leafNode("w:autoRedefine");
  }
  if (style.hidden) {
    preserved.upTo("w:hidden");
    xml.leafNode("w:hidden");
  }
  if (style.uiPriority !== undefined) {
    preserved.upTo("w:uiPriority");
    xml.leafNode("w:uiPriority", { "w:val": xmlValue(style.uiPriority, "w:val") });
  }
  if (style.semiHidden) {
    preserved.upTo("w:semiHidden");
    xml.leafNode("w:semiHidden");
  }
  if (style.unhideWhenUsed) {
    preserved.upTo("w:unhideWhenUsed");
    xml.leafNode("w:unhideWhenUsed");
  }
  if (style.qFormat) {
    preserved.upTo("w:qFormat");
    xml.leafNode("w:qFormat");
  }
  if (style.locked) {
    preserved.upTo("w:locked");
    xml.leafNode("w:locked");
  }

  // `outlineLevel` is a paragraph property: `w:outlineLvl` belongs in the
  // style's `w:pPr` (CT_PPrBase), never directly under `w:style`.
  const paragraphProperties =
    style.outlineLevel !== undefined && style.paragraphProperties?.outlineLevel === undefined
      ? { ...style.paragraphProperties, outlineLevel: style.outlineLevel }
      : style.paragraphProperties;
  if (paragraphProperties) {
    preserved.upTo("w:pPr");
    within("paragraphProperties", () =>
      renderParagraphProperties(xml, paragraphProperties, "general")
    );
  }

  if (style.runProperties) {
    preserved.upTo("w:rPr");
    const run = style.runProperties;
    within("runProperties", () => renderRunProperties(xml, run));
  }

  // For table styles
  if (style.tableProperties) {
    preserved.upTo("w:tblPr");
    const tableProperties = style.tableProperties;
    within("tableProperties", () => {
      refuseExcluded(tableProperties, "styleTable");
      renderTableProperties(xml, tableProperties, true);
    });
  }

  // Table style conditional formats
  if (style.tableStyleConditions) {
    preserved.upTo("w:tblStylePr");
    style.tableStyleConditions.forEach((cond, i) =>
      within(`tableStyleConditions[${i}]`, () => renderTableStyleCondition(xml, cond))
    );
  }

  preserved.rest();
  xml.closeNode();
}

/**
 * Render word/styles.xml. `preservedXml` is what the model does not represent
 * (`w:latentStyles`, …), written at its schema position.
 */
export function renderStyles(
  xml: XmlSink,
  docDefaults?: DocDefaults,
  styles?: readonly StyleDef[],
  preservedXml?: PreservedXml,
  rawXmlPolicy?: RawXmlPolicy
): void {
  const preserved = createPreservedEmitter(
    xml,
    STYLES_ORDER,
    preservedElements(preservedXml?.elements, rawXmlPolicy, "stylesPreservedXml.elements")
  );
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode(
    "w:styles",
    withPreservedNamespaces({ "xmlns:w": NS_W, "xmlns:r": NS_R }, preservedXml)
  );

  if (docDefaults) {
    preserved.upTo("w:docDefaults");
    renderDocDefaults(xml, docDefaults);
  }

  preserved.upTo("w:style");
  if (styles) {
    styles.forEach((style, i) =>
      within(`styles[${i}]`, () => renderStyle(xml, style, rawXmlPolicy))
    );
  }

  preserved.rest();
  xml.closeNode();
}

/** Refuse what Word does not allow in a table style's properties of `kind`. */
function refuseExcluded(properties: object, kind: TableStyleProperties): void {
  refuseAbsentFields(properties, tableStyleExcludedFields(kind), NARROWER_PROPERTIES[kind]);
}
