/**
 * DOCX Writers - Styles
 *
 * Renders word/styles.xml including document defaults, style definitions,
 * and table style conditional formats.
 */

import { NS_W, NS_R, STD_DOC_ATTRIBUTES } from "@word/constants";
import { within } from "@word/core/node-path";
import type {
  DocDefaults,
  PreservedXml,
  StyleDef,
  TableStyleConditionalFormat,
  TableBorders
} from "@word/types";
import { renderParagraphProperties } from "@word/writer/paragraph-writer";
import {
  createPreservedEmitter,
  preservedElements,
  schemaOrder,
  withPreservedNamespaces
} from "@word/writer/preserved-xml";
import { renderRunProperties, renderShading } from "@word/writer/run-writer";
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
    within("docDefaults.paragraphProperties", () => renderParagraphProperties(xml, paragraph));
    xml.closeNode();
  }

  xml.closeNode();
}

/** Map table style condition type to OOXML value. */
const CONDITION_TYPE_MAP: Record<string, string> = {
  firstRow: "firstRow",
  lastRow: "lastRow",
  firstColumn: "firstCol",
  lastColumn: "lastCol",
  oddRowBanding: "band1Horz",
  evenRowBanding: "band2Horz",
  oddColumnBanding: "band1Vert",
  evenColumnBanding: "band2Vert",
  topLeftCell: "nwCell",
  topRightCell: "neCell",
  bottomLeftCell: "swCell",
  bottomRightCell: "seCell"
};

/** Render table style conditional format (w:tblStylePr). */
function renderTableStyleCondition(xml: XmlSink, condition: TableStyleConditionalFormat): void {
  const typeVal = CONDITION_TYPE_MAP[condition.type] ?? condition.type;
  xml.openNode("w:tblStylePr", { "w:type": typeVal });

  if (condition.paragraphProperties) {
    const paragraph = condition.paragraphProperties;
    within("paragraphProperties", () => renderParagraphProperties(xml, paragraph));
  }
  if (condition.runProperties) {
    const run = condition.runProperties;
    within("runProperties", () => renderRunProperties(xml, run));
  }
  if (condition.tableProperties) {
    xml.openNode("w:tblPr");
    // Simplified table props for conditional format
    if (condition.tableProperties.borders) {
      renderTableBordersInStyle(xml, condition.tableProperties.borders);
    }
    xml.closeNode();
  }
  if (condition.cellProperties) {
    xml.openNode("w:tcPr");
    if (condition.cellProperties.shading) {
      const shading = condition.cellProperties.shading;
      within("cellProperties.shading", () => renderShading(xml, shading));
    }
    if (condition.cellProperties.borders) {
      renderCellBordersInStyle(xml, condition.cellProperties.borders);
    }
    xml.closeNode();
  }

  xml.closeNode();
}

/**
 * Render cell borders inside a style definition (w:tcBorders).
 * Cell-level border container is `w:tcBorders`, not `w:tblBorders` — using
 * the wrong wrapper produces invalid styles.xml that Word silently ignores.
 */
function renderCellBordersInStyle(xml: XmlSink, borders: TableBorders): void {
  xml.openNode("w:tcBorders");
  for (const side of [
    "top",
    "left",
    "bottom",
    "right",
    "insideH",
    "insideV",
    "start",
    "end",
    "tl2br",
    "tr2bl"
  ] as const) {
    const b = borders[side];
    if (b) {
      xml.leafNode(`w:${side}`, {
        "w:val": b.style,
        "w:sz": xmlValue(b.size ?? 4, "w:sz"),
        "w:space": xmlValue(Math.max(0, Math.round(b.space ?? 0)), "w:space"),
        "w:color": b.color ?? "auto"
      });
    }
  }
  xml.closeNode();
}

/** Render table borders inside a style definition. */
function renderTableBordersInStyle(xml: XmlSink, borders: TableBorders): void {
  xml.openNode("w:tblBorders");
  for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"] as const) {
    const b = borders[side];
    if (b) {
      xml.leafNode(`w:${side}`, {
        "w:val": b.style,
        "w:sz": xmlValue(b.size ?? 4, "w:sz"),
        "w:space": xmlValue(Math.max(0, Math.round(b.space ?? 0)), "w:space"),
        "w:color": b.color ?? "auto"
      });
    }
  }
  xml.closeNode();
}

/** CT_Style's children. */
const STYLE_ORDER = schemaOrder([
  "w:name",
  "w:aliases",
  "w:basedOn",
  "w:next",
  "w:link",
  "w:autoRedefine",
  "w:hidden",
  "w:uiPriority",
  "w:semiHidden",
  "w:unhideWhenUsed",
  "w:qFormat",
  "w:locked",
  "w:personal",
  "w:personalCompose",
  "w:personalReply",
  "w:rsid",
  "w:pPr",
  "w:rPr",
  "w:tblPr",
  "w:trPr",
  "w:tcPr",
  "w:tblStylePr"
]);

/** CT_Styles' children. */
const STYLES_ORDER = schemaOrder(["w:docDefaults", "w:latentStyles", "w:style"]);

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
    within("paragraphProperties", () => renderParagraphProperties(xml, paragraphProperties));
  }

  if (style.runProperties) {
    preserved.upTo("w:rPr");
    const run = style.runProperties;
    within("runProperties", () => renderRunProperties(xml, run));
  }

  // For table styles
  if (style.tableProperties) {
    preserved.upTo("w:tblPr");
    xml.openNode("w:tblPr");
    // Per CT_TblPrBase order, band sizes precede w:tblW.
    if (style.tableProperties.rowBandSize !== undefined) {
      xml.leafNode("w:tblStyleRowBandSize", {
        "w:val": xmlValue(style.tableProperties.rowBandSize, "w:val")
      });
    }
    if (style.tableProperties.colBandSize !== undefined) {
      xml.leafNode("w:tblStyleColBandSize", {
        "w:val": xmlValue(style.tableProperties.colBandSize, "w:val")
      });
    }
    if (style.tableProperties.width) {
      xml.leafNode("w:tblW", {
        "w:w": xmlValue(style.tableProperties.width.value, "w:w"),
        "w:type": style.tableProperties.width.type
      });
    }
    if (style.tableProperties.alignment) {
      xml.leafNode("w:jc", { "w:val": style.tableProperties.alignment });
    }
    if (style.tableProperties.indent !== undefined) {
      xml.leafNode("w:tblInd", {
        "w:w": xmlValue(style.tableProperties.indent, "w:w"),
        "w:type": "dxa"
      });
    }
    if (style.tableProperties.layout) {
      xml.leafNode("w:tblLayout", { "w:type": style.tableProperties.layout });
    }
    if (style.tableProperties.borders) {
      renderTableBordersInStyle(xml, style.tableProperties.borders);
    }
    if (style.tableProperties.cellMargins) {
      xml.openNode("w:tblCellMar");
      for (const side of ["top", "left", "bottom", "right"] as const) {
        const m = style.tableProperties.cellMargins[side];
        if (m) {
          xml.leafNode(`w:${side}`, { "w:w": xmlValue(m.value, "w:w"), "w:type": m.type });
        }
      }
      xml.closeNode();
    }
    if (style.tableProperties.cellSpacing) {
      xml.leafNode("w:tblCellSpacing", {
        "w:w": xmlValue(style.tableProperties.cellSpacing.value, "w:w"),
        "w:type": style.tableProperties.cellSpacing.type
      });
    }
    if (style.tableProperties.shading) {
      const shd = style.tableProperties.shading;
      const shdAttrs: Record<string, string> = {
        "w:val": shd.pattern ?? "clear",
        "w:fill": shd.fill ?? "auto"
      };
      if (shd.color) {
        shdAttrs["w:color"] = shd.color;
      }
      xml.leafNode("w:shd", shdAttrs);
    }
    if (style.tableProperties.look) {
      const look = style.tableProperties.look;
      const lookAttrs: Record<string, string> = {};
      if (look.firstRow !== undefined) {
        lookAttrs["w:firstRow"] = look.firstRow ? "1" : "0";
      }
      if (look.lastRow !== undefined) {
        lookAttrs["w:lastRow"] = look.lastRow ? "1" : "0";
      }
      if (look.firstColumn !== undefined) {
        lookAttrs["w:firstColumn"] = look.firstColumn ? "1" : "0";
      }
      if (look.lastColumn !== undefined) {
        lookAttrs["w:lastColumn"] = look.lastColumn ? "1" : "0";
      }
      if (look.noHBand !== undefined) {
        lookAttrs["w:noHBand"] = look.noHBand ? "1" : "0";
      }
      if (look.noVBand !== undefined) {
        lookAttrs["w:noVBand"] = look.noVBand ? "1" : "0";
      }
      xml.leafNode("w:tblLook", lookAttrs);
    }
    if (style.tableProperties.visuallyRightToLeft) {
      xml.leafNode("w:bidiVisual");
    }
    xml.closeNode();
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
