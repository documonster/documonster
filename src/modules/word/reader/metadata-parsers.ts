/**
 * DOCX Reader - Misc Metadata Parsers
 *
 * Parses small document metadata XML files that don't justify their own files:
 * - webSettings.xml (WebSettings)
 * - people.xml (PersonInfo[])
 *
 * Theme and settings parsing live in `theme-parser.ts` and `settings-parser.ts`.
 * Re-exports them for backward compatibility.
 */

import { parseXml } from "@xml/dom";

import { type Mutable } from "../core/internal-utils";
import type { PersonInfo, WebSettings } from "../types";
import { attrInt, attrVal, findChildNs } from "./parse-utils";

// Backward-compatible re-exports.
export { parseThemeXml } from "./theme-parser";
export { parseSettingsXml } from "./settings-parser";

/** Parse word/webSettings.xml. */
export function parseWebSettings(xmlStr: string): WebSettings {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const ws: Mutable<WebSettings> & Record<string, unknown> = {};

  const ofbEl = findChildNs(root, "optimizeForBrowser");
  if (ofbEl) {
    const ofb: { target?: string; majorVersion?: number } = {};
    const target = attrVal(ofbEl, "target");
    if (target) {
      ofb.target = target;
    }
    const mv = attrInt(ofbEl, "majorVersion");
    if (mv !== undefined) {
      ofb.majorVersion = mv;
    }
    ws.optimizeForBrowser = ofb;
  }
  if (findChildNs(root, "allowPNG")) {
    ws.allowPng = true;
  }
  if (findChildNs(root, "relyOnVML")) {
    ws.relyOnVml = true;
  }
  if (findChildNs(root, "doNotSaveAsSingleFile")) {
    ws.doNotSaveAsSingleFile = true;
  }
  if (findChildNs(root, "doNotOrganizeInFolder")) {
    ws.doNotOrganizeInFolder = true;
  }
  if (findChildNs(root, "useTargetMachineType")) {
    ws.useTargetMachineType = true;
  }
  return ws;
}

/** Parse word/people.xml. */
export function parsePeople(xmlStr: string): PersonInfo[] {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const people: PersonInfo[] = [];
  for (const personEl of root.children) {
    if (personEl.type !== "element") {
      continue;
    }
    const author = personEl.attributes["w15:author"] ?? personEl.attributes["author"];
    if (!author) {
      continue;
    }
    const info: Mutable<PersonInfo> = { author };
    // presenceInfo
    for (const child of personEl.children) {
      if (child.type === "element" && child.name.endsWith("presenceInfo")) {
        const pi: { providerId?: string; userId?: string } = {};
        const providerId = child.attributes["w15:providerId"] ?? child.attributes["providerId"];
        if (providerId) {
          pi.providerId = providerId;
        }
        const userId = child.attributes["w15:userId"] ?? child.attributes["userId"];
        if (userId) {
          pi.userId = userId;
        }
        if (Object.keys(pi).length > 0) {
          info.presenceInfo = pi;
        }
        break;
      }
    }
    people.push(info);
  }
  return people;
}
