/**
 * DOCX Reader - Styles Parser
 *
 * Parses `word/styles.xml` (`DocDefaults` + `StyleDef[]`). Splits out from
 * the legacy `styles-numbering-parsers.ts` to mirror the writer side
 * (`styles-writer.ts` and `numbering-writer.ts` are already separate).
 */

import { setAsideValue } from "@utils/unrecognized-values";
import { runSteps } from "@utils/yield";
import type { Steps } from "@utils/yield";
import { TABLE_STYLE_CONDITION_TOKENS } from "@word/constants";
import { TABLE_STYLE_EXCLUDED, type TableStyleProperties } from "@word/core/change-records";
import type { Mutable } from "@word/core/internal-utils";
import { parseParagraphProperties } from "@word/reader/paragraph-section-parsers";
import {
  attrInt,
  attrVal,
  findChildNs,
  findChildrenNs,
  serializeElement
} from "@word/reader/parse-utils";
import { parsePartXml } from "@word/reader/part-xml";
import {
  namespacesInScope,
  preservedRootXml,
  unmodelledChildren,
  wordLocalName
} from "@word/reader/preserved-xml";
import { parseRunProperties } from "@word/reader/properties-parsers";
import {
  parseTableCellProperties,
  parseTableProperties,
  parseTableRowProperties
} from "@word/reader/table-properties-parsers";
import type {
  DocDefaults,
  PreservedXml,
  StyleDef,
  TableStyleConditionalFormat,
  TableStyleConditionType
} from "@word/types";
import type { XmlDocument, XmlElement } from "@xml/types";

/** What `word/styles.xml` holds — see {@link parseStyles}. */
export interface ParsedStyles {
  docDefaults?: DocDefaults;
  styles: StyleDef[];
  preservedXml?: PreservedXml;
}

/** Whether a `w:tblStylePr`'s `w:type` is one the model names. */
function isTypedCondition(el: XmlElement): boolean {
  const token = attrVal(el, "type");
  return token !== undefined && CONDITION_BY_TOKEN.has(token);
}

/** A conditional format's model name by the `w:type` token it is written as. */
const CONDITION_BY_TOKEN: ReadonlyMap<string, TableStyleConditionType> = new Map(
  Object.entries(TABLE_STYLE_CONDITION_TOKENS).map(([type, token]) => [
    token,
    type as TableStyleConditionType
  ])
);

export function parseStyles(xmlStr: string, maxDepth?: number): ParsedStyles {
  return runSteps(stylesOfPart(parsePartXml(xmlStr, maxDepth)));
}

/**
 * The styles of a parsed `word/styles.xml`, a style per step — so the streaming reader can give the event loop a turn
 * while it reads a large style part (`runStepsPaced`), with the code {@link parseStyles} runs.
 */
export function* stylesOfPart(doc: XmlDocument): Steps<ParsedStyles> {
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
        dd.paragraphProperties = parseParagraphProperties(pPrEl, { scope: "general" });
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
      s.paragraphProperties = parseParagraphProperties(pPrEl, { scope: "general" });
    }
    // Earlier versions of this writer put `w:outlineLvl` directly under
    // `w:style`, where the schema has no place for it: read it as the
    // paragraph property it is.
    const misplacedOutline = findChildNs(styleEl, "outlineLvl");
    const misplacedLevel = misplacedOutline ? attrInt(misplacedOutline, "val") : undefined;
    if (misplacedLevel !== undefined && s.paragraphProperties?.outlineLevel === undefined) {
      s.paragraphProperties = {
        ...s.paragraphProperties,
        outlineLevel: misplacedLevel
      };
    }

    const rPrEl = findChildNs(styleEl, "rPr");
    if (rPrEl) {
      s.runProperties = parseRunProperties(rPrEl);
    }

    // Table properties for table styles
    const tblPrEl = findChildNs(styleEl, "tblPr");
    if (tblPrEl) {
      s.tableProperties = setAsideExcluded(
        parseTableProperties(tblPrEl, { base: true }),
        tblPrEl,
        "styleTable"
      );
    }

    // Table style conditional formats
    // A conditional format whose `w:type` is not one of `ST_TblStyleOverrideType` — missing, empty or unknown — has
    // no type to model or to resolve by, so it is kept as it was read (`preservedXml`) rather than given one.
    const tblStylePrs = findChildrenNs(styleEl, "tblStylePr").filter(isTypedCondition);
    if (tblStylePrs.length > 0) {
      const conditions: TableStyleConditionalFormat[] = [];
      for (const tsp of tblStylePrs) {
        const cond: Mutable<TableStyleConditionalFormat> = {
          type: CONDITION_BY_TOKEN.get(attrVal(tsp, "type")!)!
        };
        const cpPr = findChildNs(tsp, "pPr");
        if (cpPr) {
          cond.paragraphProperties = parseParagraphProperties(cpPr, { scope: "general" });
        }
        const crPr = findChildNs(tsp, "rPr");
        if (crPr) {
          cond.runProperties = parseRunProperties(crPr);
        }
        const ctblPr = findChildNs(tsp, "tblPr");
        if (ctblPr) {
          cond.tableProperties = setAsideExcluded(
            parseTableProperties(ctblPr, { base: true }),
            ctblPr,
            "conditionTable"
          );
        }
        const ctrPr = findChildNs(tsp, "trPr");
        if (ctrPr) {
          cond.rowProperties = setAsideExcluded(
            parseTableRowProperties(ctrPr),
            ctrPr,
            "conditionRow"
          );
        }
        const ctcPr = findChildNs(tsp, "tcPr");
        if (ctcPr) {
          cond.cellProperties = setAsideExcluded(
            parseTableCellProperties(ctcPr),
            ctcPr,
            "conditionCell"
          );
        }
        conditions.push(cond);
      }
      s.tableStyleConditions = conditions;
    }

    const preservedXml = unmodelledChildren(styleEl, rootScope, child => {
      const local = wordLocalName(child);
      return (
        local !== undefined &&
        MODELLED_STYLE_CHILDREN.has(local) &&
        (local !== "tblStylePr" || isTypedCondition(child))
      );
    });
    if (preservedXml.length > 0) {
      s.preservedXml = preservedXml;
    }

    styles.push(s);
    yield;
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

/**
 * `properties`, read from `source`, without what Word does not allow in a table style's properties of `kind`
 * (`TABLE_STYLE_EXCLUDED`): each such field is cleared and its elements kept as not written back, so the read report
 * names them and writing the document leaves them out — as Word, which drops them, does.
 */
function setAsideExcluded<T extends object>(
  properties: T,
  source: XmlElement,
  kind: TableStyleProperties
): T {
  const excluded: Readonly<Record<string, string>> = TABLE_STYLE_EXCLUDED[kind];
  const byField = new Map<string, string[]>();
  for (const child of source.children) {
    if (child.type !== "element") {
      continue;
    }
    const local = wordLocalName(child);
    if (local !== undefined && Object.hasOwn(excluded, local)) {
      const field = excluded[local]!;
      byField.set(field, [...(byField.get(field) ?? []), serializeElement(child)]);
    }
  }
  for (const [field, elements] of byField) {
    delete (properties as Record<string, unknown>)[field];
    setAsideValue(properties, field, elements.join(""), { writtenBack: false });
  }
  return properties;
}
