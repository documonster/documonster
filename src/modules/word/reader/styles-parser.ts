/**
 * DOCX Reader - Styles Parser
 *
 * Parses `word/styles.xml` (`DocDefaults` + `StyleDef[]`). Splits out from
 * the legacy `styles-numbering-parsers.ts` to mirror the writer side
 * (`styles-writer.ts` and `numbering-writer.ts` are already separate).
 */

import type { Mutable } from "@word/core/internal-utils";
import { parseParagraphProperties } from "@word/reader/paragraph-section-parsers";
import { attrInt, attrVal, findChildNs, findChildrenNs } from "@word/reader/parse-utils";
import { parsePartXml } from "@word/reader/part-xml";
import {
  namespacesInScope,
  preservedRootXml,
  unmodelledChildren,
  wordLocalName
} from "@word/reader/preserved-xml";
import { parseRowHeight, parseRunProperties, parseShading } from "@word/reader/properties-parsers";
import { parseTableBorders, parseTableProperties } from "@word/reader/table-properties-parsers";
import type {
  DocDefaults,
  PreservedXml,
  StyleDef,
  TableCellProperties,
  TableRowProperties,
  TableStyleConditionalFormat
} from "@word/types";

export function parseStyles(
  xmlStr: string,
  maxDepth?: number
): { docDefaults?: DocDefaults; styles: StyleDef[]; preservedXml?: PreservedXml } {
  const doc = parsePartXml(xmlStr, maxDepth);
  const root = doc.root;
  const rootScope = namespacesInScope({}, root);
  let docDefaults: DocDefaults | undefined;
  const styles: StyleDef[] = [];

  const ddEl = findChildNs(root, "docDefaults");
  if (ddEl) {
    const dd: Mutable<DocDefaults> = {};
    const rPrDefaultEl = findChildNs(ddEl, "rPrDefault");
    if (rPrDefaultEl) {
      const rPrEl = findChildNs(rPrDefaultEl, "rPr");
      if (rPrEl) {
        dd.runProperties = parseRunProperties(rPrEl);
      }
    }
    const pPrDefaultEl = findChildNs(ddEl, "pPrDefault");
    if (pPrDefaultEl) {
      const pPrEl = findChildNs(pPrDefaultEl, "pPr");
      if (pPrEl) {
        dd.paragraphProperties = parseParagraphProperties(pPrEl);
      }
    }
    docDefaults = dd;
  }

  for (const styleEl of findChildrenNs(root, "style")) {
    const s: Mutable<StyleDef> & Record<string, unknown> = {} as Mutable<StyleDef>;
    s.type = attrVal(styleEl, "type") as StyleDef["type"];
    s.styleId = attrVal(styleEl, "styleId") ?? "";
    s.isDefault = attrVal(styleEl, "default") === "1";
    if (attrVal(styleEl, "customStyle") === "1") {
      s.customStyle = true;
    }

    const nameEl = findChildNs(styleEl, "name");
    s.name = nameEl ? (attrVal(nameEl, "val") ?? "") : "";

    const basedOnEl = findChildNs(styleEl, "basedOn");
    if (basedOnEl) {
      s.basedOn = attrVal(basedOnEl, "val");
    }

    const nextEl = findChildNs(styleEl, "next");
    if (nextEl) {
      s.next = attrVal(nextEl, "val");
    }

    const linkEl = findChildNs(styleEl, "link");
    if (linkEl) {
      s.link = attrVal(linkEl, "val");
    }

    const uiPrEl = findChildNs(styleEl, "uiPriority");
    if (uiPrEl) {
      s.uiPriority = attrInt(uiPrEl, "val");
    }

    if (findChildNs(styleEl, "qFormat")) {
      s.qFormat = true;
    }
    if (findChildNs(styleEl, "semiHidden")) {
      s.semiHidden = true;
    }
    if (findChildNs(styleEl, "unhideWhenUsed")) {
      s.unhideWhenUsed = true;
    }
    if (findChildNs(styleEl, "hidden")) {
      s.hidden = true;
    }
    if (findChildNs(styleEl, "locked")) {
      s.locked = true;
    }
    if (findChildNs(styleEl, "autoRedefine")) {
      s.autoRedefine = true;
    }

    const pPrEl = findChildNs(styleEl, "pPr");
    if (pPrEl) {
      s.paragraphProperties = parseParagraphProperties(pPrEl);
    }
    // Earlier versions of this writer put `w:outlineLvl` directly under
    // `w:style`, where the schema has no place for it: read it as the
    // paragraph property it is.
    const misplacedOutline = findChildNs(styleEl, "outlineLvl");
    const misplacedLevel = misplacedOutline ? attrInt(misplacedOutline, "val") : undefined;
    if (misplacedLevel !== undefined && s.paragraphProperties?.outlineLevel === undefined) {
      s.paragraphProperties = { ...s.paragraphProperties, outlineLevel: misplacedLevel };
    }

    const rPrEl = findChildNs(styleEl, "rPr");
    if (rPrEl) {
      s.runProperties = parseRunProperties(rPrEl);
    }

    // Table properties for table styles
    const tblPrEl = findChildNs(styleEl, "tblPr");
    if (tblPrEl) {
      s.tableProperties = parseTableProperties(tblPrEl);
    }

    // Table style conditional formats
    const tblStylePrs = findChildrenNs(styleEl, "tblStylePr");
    if (tblStylePrs.length > 0) {
      const conditions: TableStyleConditionalFormat[] = [];
      for (const tsp of tblStylePrs) {
        const cond: Mutable<TableStyleConditionalFormat> = {
          type: attrVal(tsp, "type") as TableStyleConditionalFormat["type"]
        };
        const cpPr = findChildNs(tsp, "pPr");
        if (cpPr) {
          cond.paragraphProperties = parseParagraphProperties(cpPr);
        }
        const crPr = findChildNs(tsp, "rPr");
        if (crPr) {
          cond.runProperties = parseRunProperties(crPr);
        }
        const ctblPr = findChildNs(tsp, "tblPr");
        if (ctblPr) {
          cond.tableProperties = parseTableProperties(ctblPr);
        }
        const ctrPr = findChildNs(tsp, "trPr");
        if (ctrPr) {
          const rp: Mutable<TableRowProperties> = {};
          const hEl = findChildNs(ctrPr, "trHeight");
          if (hEl) {
            rp.height = parseRowHeight(hEl);
          }
          cond.rowProperties = rp;
        }
        const ctcPr = findChildNs(tsp, "tcPr");
        if (ctcPr) {
          const cp: Mutable<TableCellProperties> = {};
          const bEl = findChildNs(ctcPr, "tcBorders");
          if (bEl) {
            cp.borders = parseTableBorders(bEl);
          }
          const shd = findChildNs(ctcPr, "shd");
          if (shd) {
            cp.shading = parseShading(shd);
          }
          cond.cellProperties = cp;
        }
        conditions.push(cond);
      }
      s.tableStyleConditions = conditions;
    }

    const preservedXml = unmodelledChildren(styleEl, rootScope, child => {
      const local = wordLocalName(child);
      return local !== undefined && MODELLED_STYLE_CHILDREN.has(local);
    });
    if (preservedXml.length > 0) {
      s.preservedXml = preservedXml;
    }

    styles.push(s);
  }

  const preservedXml = preservedRootXml(root, child => {
    const local = wordLocalName(child);
    return local === "docDefaults" || local === "style";
  });
  return { docDefaults, styles, ...(preservedXml ? { preservedXml } : {}) };
}

/** The `w:style` children read into the model; the writer writes them back. */
const MODELLED_STYLE_CHILDREN = new Set([
  "name",
  "basedOn",
  "next",
  "link",
  "autoRedefine",
  "hidden",
  "uiPriority",
  "semiHidden",
  "unhideWhenUsed",
  "qFormat",
  "locked",
  "pPr",
  "rPr",
  "tblPr",
  "tblStylePr",
  // Misplaced by earlier versions of this writer; read into `w:pPr`.
  "outlineLvl"
]);
