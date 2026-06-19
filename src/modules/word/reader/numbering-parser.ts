/**
 * DOCX Reader - Numbering Parser
 *
 * Parses `word/numbering.xml`. Splits out from the legacy
 * `styles-numbering-parsers.ts` to mirror the writer side
 * (`numbering-writer.ts` is already separate from `styles-writer.ts`).
 */

import type { Mutable } from "@word/core/internal-utils";
import { parseParagraphProperties } from "@word/reader/paragraph-section-parsers";
import {
  attrInt,
  attrVal,
  findChildNs,
  findChildrenNs,
  serializeElement
} from "@word/reader/parse-utils";
import { parseRunProperties } from "@word/reader/properties-parsers";
import type {
  AbstractNumbering,
  LevelOverride,
  NumberingInstance,
  NumberingLevel,
  NumPicBullet
} from "@word/types";
import { ptToEmu } from "@word/units";
import { findChild, parseXml } from "@xml/dom";
import type { XmlElement } from "@xml/types";

export function parseNumberingXml(xmlStr: string): {
  abstractNums: AbstractNumbering[];
  instances: NumberingInstance[];
  numPicBullets: NumPicBullet[];
} {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const abstractNums: AbstractNumbering[] = [];
  const instances: NumberingInstance[] = [];
  const numPicBullets: NumPicBullet[] = [];

  // Parse picture bullets
  for (const pbEl of findChildrenNs(root, "numPicBullet")) {
    const id = attrInt(pbEl, "numPicBulletId");
    if (id === undefined) {
      continue;
    }
    const pb: Mutable<NumPicBullet> = { id };
    // Try to extract VML shape info
    const pictEl = findChildNs(pbEl, "pict");
    if (pictEl) {
      // Preserve raw VML for complete fidelity
      let rawVml = "";
      for (const child of pictEl.children) {
        if (child.type === "element") {
          rawVml += serializeElement(child);
        }
      }
      if (rawVml) {
        pb.rawVmlXml = rawVml;
      }
      // Extract rId from v:imagedata
      const shapeEl = findChild(pictEl, "v:shape");
      if (shapeEl) {
        const imgDataEl = findChild(shapeEl, "v:imagedata");
        if (imgDataEl) {
          const rId = imgDataEl.attributes["r:id"] ?? imgDataEl.attributes["r:pict"];
          if (rId) {
            pb.rId = rId;
          }
        }
        // Extract width/height from style
        const style = shapeEl.attributes["style"];
        if (style) {
          const wMatch = /width:([\d.]+)pt/i.exec(style);
          const hMatch = /height:([\d.]+)pt/i.exec(style);
          if (wMatch) {
            pb.width = ptToEmu(parseFloat(wMatch[1]));
          }
          if (hMatch) {
            pb.height = ptToEmu(parseFloat(hMatch[1]));
          }
        }
      }
    }
    numPicBullets.push(pb);
  }

  for (const absEl of findChildrenNs(root, "abstractNum")) {
    const levels: NumberingLevel[] = [];
    for (const lvlEl of findChildrenNs(absEl, "lvl")) {
      levels.push(parseLevel(lvlEl));
    }

    const abs: Mutable<AbstractNumbering> = {
      abstractNumId: attrInt(absEl, "abstractNumId") ?? 0,
      levels
    };
    const mltEl = findChildNs(absEl, "multiLevelType");
    if (mltEl) {
      abs.multiLevelType = attrVal(mltEl, "val") as AbstractNumbering["multiLevelType"];
    }
    const numStyleLinkEl = findChildNs(absEl, "numStyleLink");
    if (numStyleLinkEl) {
      abs.numStyleLink = attrVal(numStyleLinkEl, "val");
    }
    const styleLinkEl = findChildNs(absEl, "styleLink");
    if (styleLinkEl) {
      abs.styleLink = attrVal(styleLinkEl, "val");
    }
    abstractNums.push(abs);
  }

  for (const numEl of findChildrenNs(root, "num")) {
    const absIdEl = findChildNs(numEl, "abstractNumId");
    const overrides: LevelOverride[] = [];
    for (const ovEl of findChildrenNs(numEl, "lvlOverride")) {
      const ov: Mutable<LevelOverride> = { level: attrInt(ovEl, "ilvl") ?? 0 };
      const startOvEl = findChildNs(ovEl, "startOverride");
      if (startOvEl) {
        ov.startOverride = attrInt(startOvEl, "val");
      }
      // Level def override: parse full level definition
      const lvlEl = findChildNs(ovEl, "lvl");
      if (lvlEl) {
        ov.levelDef = parseLevel(lvlEl);
        // Inherit level index from parent if not specified
        if ((ov.levelDef as Mutable<NumberingLevel>).level === undefined) {
          (ov.levelDef as Mutable<NumberingLevel>).level = ov.level;
        }
      }
      overrides.push(ov);
    }
    instances.push({
      numId: attrInt(numEl, "numId") ?? 0,
      abstractNumId: absIdEl ? (attrInt(absIdEl, "val") ?? 0) : 0,
      overrides: overrides.length > 0 ? overrides : undefined
    });
  }

  return { abstractNums, instances, numPicBullets };
}

/** Parse a w:lvl element into a NumberingLevel (shared by abstractNum and lvlOverride). */
function parseLevel(lvlEl: XmlElement): NumberingLevel {
  const level: Partial<Mutable<NumberingLevel>> = {
    level: attrInt(lvlEl, "ilvl") ?? 0
  };

  const startEl = findChildNs(lvlEl, "start");
  if (startEl) {
    level.start = attrInt(startEl, "val");
  }
  const fmtEl = findChildNs(lvlEl, "numFmt");
  if (fmtEl) {
    level.format = attrVal(fmtEl, "val") as NumberingLevel["format"];
  }
  const textEl = findChildNs(lvlEl, "lvlText");
  if (textEl) {
    level.text = attrVal(textEl, "val") ?? "";
  }
  const pStyleEl = findChildNs(lvlEl, "pStyle");
  if (pStyleEl) {
    level.paragraphStyle = attrVal(pStyleEl, "val");
  }
  const jcEl = findChildNs(lvlEl, "lvlJc");
  if (jcEl) {
    level.justification = attrVal(jcEl, "val") as NumberingLevel["justification"];
  }
  const pPrEl = findChildNs(lvlEl, "pPr");
  if (pPrEl) {
    level.paragraphProperties = parseParagraphProperties(pPrEl);
  }
  const rPrEl = findChildNs(lvlEl, "rPr");
  if (rPrEl) {
    level.runProperties = parseRunProperties(rPrEl);
  }
  const suffEl = findChildNs(lvlEl, "suff");
  if (suffEl) {
    level.suffix = attrVal(suffEl, "val") as NumberingLevel["suffix"];
  }
  if (findChildNs(lvlEl, "isLgl")) {
    level.isLegalNumberingStyle = true;
  }
  const lvlRestartEl = findChildNs(lvlEl, "lvlRestart");
  if (lvlRestartEl) {
    level.restartAfterLevel = attrInt(lvlRestartEl, "val");
  }
  const picBulletEl = findChildNs(lvlEl, "lvlPicBulletId");
  if (picBulletEl) {
    level.picBulletId = attrInt(picBulletEl, "val");
  }
  return level as NumberingLevel;
}
