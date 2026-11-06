/**
 * DOCX Writers - Numbering
 *
 * Renders word/numbering.xml with abstract numbering definitions and instances.
 * Supports suffix, legal numbering style, restart levels, and level overrides.
 */

import { NS_W, NS_R, STD_DOC_ATTRIBUTES } from "@word/constants";
import { within } from "@word/core/node-path";
import { rebindRelationshipIds } from "@word/core/relationship-ids";
import { DocxRawXmlPolicyError } from "@word/errors";
import type {
  AbstractNumbering,
  NumberingInstance,
  NumberingLevel,
  NumPicBullet
} from "@word/types";
import { EMU_PER_POINT } from "@word/units";
import { renderParagraphProperties } from "@word/writer/paragraph-writer";
import { renderRunProperties } from "@word/writer/run-writer";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

/** Render a single numbering level. */
function renderLevel(xml: XmlSink, level: NumberingLevel): void {
  xml.openNode("w:lvl", { "w:ilvl": xmlValue(level.level, "w:ilvl") });

  // Per ECMA-376 CT_Lvl order:
  // start → numFmt → lvlRestart → pStyle → isLgl → suff → lvlText → lvlPicBulletId → legacy → lvlJc → pPr → rPr
  if (level.start !== undefined) {
    xml.leafNode("w:start", { "w:val": xmlValue(level.start, "w:val") });
  }

  xml.leafNode("w:numFmt", { "w:val": level.format });

  if (level.restartAfterLevel !== undefined) {
    xml.leafNode("w:lvlRestart", { "w:val": xmlValue(level.restartAfterLevel, "w:val") });
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
    xml.leafNode("w:lvlPicBulletId", { "w:val": xmlValue(level.picBulletId, "w:val") });
  }

  if (level.justification) {
    xml.leafNode("w:lvlJc", { "w:val": level.justification });
  }

  if (level.paragraphProperties) {
    const paragraph = level.paragraphProperties;
    within("paragraphProperties", () => renderParagraphProperties(xml, paragraph));
  }

  if (level.runProperties) {
    const run = level.runProperties;
    within("runProperties", () => renderRunProperties(xml, run));
  }

  xml.closeNode();
}

/** Render word/numbering.xml. */
export function renderNumbering(
  xml: XmlSink,
  abstractNums?: readonly AbstractNumbering[],
  instances?: readonly NumberingInstance[],
  numPicBullets?: readonly NumPicBullet[],
  rawXmlPolicy?: "preserve" | "strip" | "reject",
  imageRemap?: ReadonlyMap<string, string>
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
      xml.openNode("w:numPicBullet", { "w:numPicBulletId": xmlValue(pb.id, "w:numPicBulletId") });
      if (pb.rawVmlXml) {
        if (rawXmlPolicy === "reject") {
          throw new DocxRawXmlPolicyError("numPicBullet.rawVmlXml");
        }
        if (rawXmlPolicy !== "strip") {
          // Use raw VML for full fidelity
          xml.writeRaw(rebindRelationshipIds(pb.rawVmlXml, imageRemap));
        }
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
          style: `width:${pb.width ? Math.round(pb.width / EMU_PER_POINT) : 12}pt;height:${pb.height ? Math.round(pb.height / EMU_PER_POINT) : 12}pt`
        });
        xml.leafNode("v:imagedata", { "r:id": imageRemap?.get(pb.rId) ?? pb.rId, "o:title": "" });
        xml.closeNode();
        xml.closeNode();
      }
      xml.closeNode();
    }
  }

  if (abstractNums) {
    abstractNums.forEach((abs, a) => {
      xml.openNode("w:abstractNum", {
        "w:abstractNumId": xmlValue(abs.abstractNumId, "w:abstractNumId")
      });
      if (abs.multiLevelType) {
        xml.leafNode("w:multiLevelType", { "w:val": abs.multiLevelType });
      }
      if (abs.numStyleLink) {
        xml.leafNode("w:numStyleLink", { "w:val": abs.numStyleLink });
      }
      if (abs.styleLink) {
        xml.leafNode("w:styleLink", { "w:val": abs.styleLink });
      }
      abs.levels.forEach((level, l) =>
        within(`abstractNumberings[${a}].levels[${l}]`, () => renderLevel(xml, level))
      );
      xml.closeNode();
    });
  }

  if (instances) {
    instances.forEach((inst, n) => {
      xml.openNode("w:num", { "w:numId": xmlValue(inst.numId, "w:numId") });
      xml.leafNode("w:abstractNumId", { "w:val": xmlValue(inst.abstractNumId, "w:val") });
      if (inst.overrides) {
        inst.overrides.forEach((ov, o) => {
          xml.openNode("w:lvlOverride", { "w:ilvl": xmlValue(ov.level, "w:ilvl") });
          if (ov.startOverride !== undefined) {
            xml.leafNode("w:startOverride", { "w:val": xmlValue(ov.startOverride, "w:val") });
          }
          if (ov.levelDef) {
            const levelDef = ov.levelDef;
            within(`numberingInstances[${n}].overrides[${o}].levelDef`, () =>
              renderLevel(xml, levelDef)
            );
          }
          xml.closeNode();
        });
      }
      xml.closeNode();
    });
  }

  xml.closeNode();
}
