/**
 * DOCX Writers - Numbering
 *
 * Renders word/numbering.xml with abstract numbering definitions and instances.
 * Supports suffix, legal numbering style, restart levels, and level overrides.
 */

import type { XmlSink } from "@xml/types";

import { NS_W, NS_R, STD_DOC_ATTRIBUTES } from "../constants";
import type { AbstractNumbering, NumberingInstance, NumberingLevel, NumPicBullet } from "../types";
import { renderParagraphProperties } from "./paragraph-writer";
import { renderRunProperties } from "./run-writer";

/** Render a single numbering level. */
function renderLevel(xml: XmlSink, level: NumberingLevel): void {
  xml.openNode("w:lvl", { "w:ilvl": String(level.level) });

  // Per ECMA-376 CT_Lvl order:
  // start → numFmt → lvlRestart → pStyle → isLgl → suff → lvlText → lvlPicBulletId → legacy → lvlJc → pPr → rPr
  if (level.start !== undefined) {
    xml.leafNode("w:start", { "w:val": String(level.start) });
  }

  xml.leafNode("w:numFmt", { "w:val": level.format });

  if (level.restartAfterLevel !== undefined) {
    xml.leafNode("w:lvlRestart", { "w:val": String(level.restartAfterLevel) });
  }

  if (level.paragraphStyle) {
    xml.leafNode("w:pStyle", { "w:val": level.paragraphStyle });
  }

  if (level.isLegalNumberingStyle) {
    xml.leafNode("w:isLgl");
  }

  if (level.suffix) {
    xml.leafNode("w:suff", { "w:val": level.suffix });
  }

  xml.leafNode("w:lvlText", { "w:val": level.text });

  if (level.picBulletId !== undefined) {
    xml.leafNode("w:lvlPicBulletId", { "w:val": String(level.picBulletId) });
  }

  if (level.justification) {
    xml.leafNode("w:lvlJc", { "w:val": level.justification });
  }

  if (level.paragraphProperties) {
    renderParagraphProperties(xml, level.paragraphProperties);
  }

  if (level.runProperties) {
    renderRunProperties(xml, level.runProperties);
  }

  xml.closeNode();
}

/** Render word/numbering.xml. */
export function renderNumbering(
  xml: XmlSink,
  abstractNums?: readonly AbstractNumbering[],
  instances?: readonly NumberingInstance[],
  numPicBullets?: readonly NumPicBullet[]
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:numbering", {
    "xmlns:w": NS_W,
    "xmlns:r": NS_R,
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:o": "urn:schemas-microsoft-com:office:office"
  });

  // Picture bullets come first
  if (numPicBullets) {
    for (const pb of numPicBullets) {
      xml.openNode("w:numPicBullet", { "w:numPicBulletId": String(pb.id) });
      if (pb.rawVmlXml) {
        // Use raw VML for full fidelity
        xml.writeRaw(pb.rawVmlXml);
      } else if (pb.rId) {
        // Default VML shape with image fill
        xml.openNode("w:pict");
        xml.leafNode("v:shapetype", {
          id: "_x0000_t75",
          coordsize: "21600,21600",
          "o:spt": "75",
          "o:preferrelative": "t",
          path: "m@4@5l@4@11@9@11@9@5xe",
          filled: "f",
          stroked: "f"
        });
        xml.openNode("v:shape", {
          id: `_x0000_i10${pb.id}`,
          type: "#_x0000_t75",
          style: `width:${pb.width ? Math.round(pb.width / 12700) : 12}pt;height:${pb.height ? Math.round(pb.height / 12700) : 12}pt`
        });
        xml.leafNode("v:imagedata", { "r:id": pb.rId, "o:title": "" });
        xml.closeNode();
        xml.closeNode();
      }
      xml.closeNode();
    }
  }

  if (abstractNums) {
    for (const abs of abstractNums) {
      xml.openNode("w:abstractNum", { "w:abstractNumId": String(abs.abstractNumId) });
      if (abs.multiLevelType) {
        xml.leafNode("w:multiLevelType", { "w:val": abs.multiLevelType });
      }
      if (abs.numStyleLink) {
        xml.leafNode("w:numStyleLink", { "w:val": abs.numStyleLink });
      }
      if (abs.styleLink) {
        xml.leafNode("w:styleLink", { "w:val": abs.styleLink });
      }
      for (const level of abs.levels) {
        renderLevel(xml, level);
      }
      xml.closeNode();
    }
  }

  if (instances) {
    for (const inst of instances) {
      xml.openNode("w:num", { "w:numId": String(inst.numId) });
      xml.leafNode("w:abstractNumId", { "w:val": String(inst.abstractNumId) });
      if (inst.overrides) {
        for (const ov of inst.overrides) {
          xml.openNode("w:lvlOverride", { "w:ilvl": String(ov.level) });
          if (ov.startOverride !== undefined) {
            xml.leafNode("w:startOverride", { "w:val": String(ov.startOverride) });
          }
          if (ov.levelDef) {
            renderLevel(xml, ov.levelDef);
          }
          xml.closeNode();
        }
      }
      xml.closeNode();
    }
  }

  xml.closeNode();
}
