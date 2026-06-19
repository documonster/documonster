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
import { parseRunProperties, parseShading } from "@word/reader/properties-parsers";
import { parseTableBorders, parseTableProperties } from "@word/reader/table-properties-parsers";
import type {
  DocDefaults,
  StyleDef,
  TableCellProperties,
  TableRowProperties,
  TableStyleConditionalFormat
} from "@word/types";
import { parseXml } from "@xml/dom";

export function parseStyles(xmlStr: string): { docDefaults?: DocDefaults; styles: StyleDef[] } {
  const doc = parseXml(xmlStr);
  const root = doc.root;
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
            rp.height = {
              value: attrInt(hEl, "val") ?? 0,
              rule: attrVal(hEl, "hRule") as NonNullable<TableRowProperties["height"]>["rule"]
            };
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

    styles.push(s);
  }

  return { docDefaults, styles };
}
