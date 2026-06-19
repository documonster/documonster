/**
 * DOCX Reader - Settings Parser
 *
 * Parses `word/settings.xml` (DocumentSettings).
 * Extracted from the original `metadata-parsers.ts`.
 */

import type { Mutable } from "@word/core/internal-utils";
import {
  attrInt,
  attrVal,
  findChildNs,
  findChildrenNs,
  parseNoteProperties,
  serializeElement,
  sidToHashAlgorithm
} from "@word/reader/parse-utils";
import type {
  CompatFlag,
  CompatSetting,
  DocumentSettings,
  HyphenationSettings,
  ProtectionType
} from "@word/types";
import { parseXml } from "@xml/dom";

export function parseSettingsXml(xmlStr: string): DocumentSettings {
  const doc = parseXml(xmlStr);
  const root = doc.root;
  const settings: Mutable<DocumentSettings> = {};

  const zoomEl = findChildNs(root, "zoom");
  if (zoomEl) {
    settings.zoom = attrInt(zoomEl, "percent");
  }
  const tabEl = findChildNs(root, "defaultTabStop");
  if (tabEl) {
    settings.defaultTabStop = attrInt(tabEl, "val");
  }

  const csControlEl = findChildNs(root, "characterSpacingControl");
  if (csControlEl) {
    const v = attrVal(csControlEl, "val");
    if (
      v === "doNotCompress" ||
      v === "compressPunctuation" ||
      v === "compressPunctuationAndJapaneseKana"
    ) {
      settings.characterSpacingControl = v;
    }
  }

  // Extended settings
  if (findChildNs(root, "doNotTrackMoves")) {
    settings.doNotTrackMoves = true;
  }
  if (findChildNs(root, "doNotTrackFormatting")) {
    settings.doNotTrackFormatting = true;
  }
  if (findChildNs(root, "doNotDemoteNonCombiningChars")) {
    settings.doNotDemoteAsianTextFirstLine = true;
  }
  const ssFontsEl = findChildNs(root, "saveSubsetFonts");
  if (ssFontsEl) {
    const v = attrVal(ssFontsEl, "val");
    settings.saveSubsetFonts = v !== "0" && v !== "false";
  }
  if (findChildNs(root, "noPunctuationKerning")) {
    settings.noPunctuationKerning = true;
  }
  if (findChildNs(root, "bordersDoNotSurroundHeader")) {
    settings.bordersDoNotSurroundHeader = true;
  }
  if (findChildNs(root, "bordersDoNotSurroundFooter")) {
    settings.bordersDoNotSurroundFooter = true;
  }
  const clickStyleEl = findChildNs(root, "clickAndTypeStyle");
  if (clickStyleEl) {
    settings.clickAndTypeStyle = attrVal(clickStyleEl, "val");
  }
  const spfEl = findChildNs(root, "stylePaneFormatFilter");
  if (spfEl) {
    settings.stylePaneFormatFilter = attrVal(spfEl, "val");
  }
  const spsEl = findChildNs(root, "stylePaneSortMethod");
  if (spsEl) {
    settings.stylePaneSortMethod = attrVal(spsEl, "val");
  }
  const tflEl = findChildNs(root, "themeFontLang");
  if (tflEl) {
    const tfl: { val?: string; eastAsia?: string; bidi?: string } = {};
    const v = attrVal(tflEl, "val");
    if (v) {
      tfl.val = v;
    }
    const ea = attrVal(tflEl, "eastAsia");
    if (ea) {
      tfl.eastAsia = ea;
    }
    const bd = attrVal(tflEl, "bidi");
    if (bd) {
      tfl.bidi = bd;
    }
    if (Object.keys(tfl).length > 0) {
      settings.themeFontLang = tfl;
    }
  }
  const dsEl = findChildNs(root, "decimalSymbol");
  if (dsEl) {
    settings.decimalSymbol = attrVal(dsEl, "val");
  }
  const lsEl = findChildNs(root, "listSeparator");
  if (lsEl) {
    settings.listSeparator = attrVal(lsEl, "val");
  }

  // RSID list
  const rsidsEl = findChildNs(root, "rsids");
  if (rsidsEl) {
    const rsids: { rsidRoot?: string; rsid?: string[] } = {};
    const rootEl = findChildNs(rsidsEl, "rsidRoot");
    if (rootEl) {
      rsids.rsidRoot = attrVal(rootEl, "val");
    }
    const rsidList: string[] = [];
    for (const rsidEl of findChildrenNs(rsidsEl, "rsid")) {
      const v = attrVal(rsidEl, "val");
      if (v) {
        rsidList.push(v);
      }
    }
    if (rsidList.length > 0) {
      rsids.rsid = rsidList;
    }
    if (Object.keys(rsids).length > 0) {
      settings.rsids = rsids;
    }
  }

  if (findChildNs(root, "evenAndOddHeaders")) {
    settings.evenAndOddHeaders = true;
  }

  if (findChildNs(root, "trackRevisions")) {
    settings.trackRevisions = true;
  }

  if (findChildNs(root, "mirrorMargins")) {
    settings.mirrorMargins = true;
  }

  if (findChildNs(root, "gutterAtTop")) {
    settings.gutterAtTop = true;
  }

  if (findChildNs(root, "displayBackgroundShape")) {
    settings.displayBackgroundShape = true;
  }

  if (findChildNs(root, "updateFields")) {
    settings.updateFieldsOnOpen = true;
  }

  // Hyphenation
  const autoHyphEl = findChildNs(root, "autoHyphenation");
  if (autoHyphEl) {
    settings.autoHyphenation = true;
    const hyph: Mutable<HyphenationSettings> = { autoHyphenation: true };
    const hzEl = findChildNs(root, "hyphenationZone");
    if (hzEl) {
      hyph.hyphenationZone = attrInt(hzEl, "val");
    }
    const chlEl = findChildNs(root, "consecutiveHyphenLimit");
    if (chlEl) {
      hyph.consecutiveHyphenLimit = attrInt(chlEl, "val");
    }
    if (findChildNs(root, "doNotHyphenateCaps")) {
      hyph.doNotHyphenateCaps = true;
    }
    settings.hyphenation = hyph;
  }

  // Document protection
  const protEl = findChildNs(root, "documentProtection");
  if (protEl) {
    settings.documentProtection = {
      type: (attrVal(protEl, "edit") ?? "none") as ProtectionType,
      edit: attrVal(protEl, "edit") ?? undefined,
      enforcement: attrVal(protEl, "enforcement") === "1",
      formatting: attrVal(protEl, "formatting") === "1" ? true : undefined,
      hashAlgorithm: sidToHashAlgorithm(attrVal(protEl, "cryptAlgorithmSid")),
      hashValue: attrVal(protEl, "hash") ?? undefined,
      saltValue: attrVal(protEl, "salt") ?? undefined,
      spinCount: attrVal(protEl, "cryptSpinCount")
        ? parseInt(attrVal(protEl, "cryptSpinCount")!, 10)
        : undefined
    };
  }

  const compatEl = findChildNs(root, "compat");
  if (compatEl) {
    const compatSettings: CompatSetting[] = [];
    const compatFlags: CompatFlag[] = [];
    for (const csEl of compatEl.children) {
      if (csEl.type !== "element") {
        continue;
      }
      const localName = csEl.name.replace(/^w:/, "");
      if (localName === "compatSetting") {
        const name = attrVal(csEl, "name");
        const uri = attrVal(csEl, "uri");
        const val = attrVal(csEl, "val");
        if (name === "compatibilityMode" && val !== undefined) {
          settings.compatibilityMode = parseInt(val, 10);
        } else if (name !== undefined && uri !== undefined && val !== undefined) {
          compatSettings.push({ name, uri, val });
        }
      } else {
        // Legacy compat flags (w:useFELayout, w:balanceSingleByteDoubleByteWidth, etc.)
        compatFlags.push({ name: localName, val: attrVal(csEl, "val") });
      }
    }
    if (compatSettings.length > 0) {
      settings.compatSettings = compatSettings;
    }
    if (compatFlags.length > 0) {
      settings.compatFlags = compatFlags;
    }
  }

  // Mail merge settings (preserve as raw XML)
  const mailMergeEl = findChildNs(root, "mailMerge");
  if (mailMergeEl) {
    settings.mailMergeRawXml = serializeElement(mailMergeEl);
  }

  // Write protection
  const writeProtectionEl = findChildNs(root, "writeProtection");
  if (writeProtectionEl) {
    const wp: NonNullable<DocumentSettings["writeProtection"]> = {};
    const wpMut = wp as Mutable<NonNullable<DocumentSettings["writeProtection"]>>;
    const recommended = attrVal(writeProtectionEl, "recommended");
    if (recommended === "1" || recommended === "true") {
      wpMut.recommended = true;
    }
    const algName = attrVal(writeProtectionEl, "algorithmName");
    if (algName) {
      wpMut.algorithmName = algName;
    }
    const hashValue = attrVal(writeProtectionEl, "hashValue");
    if (hashValue) {
      wpMut.hashValue = hashValue;
    }
    const saltValue = attrVal(writeProtectionEl, "saltValue");
    if (saltValue) {
      wpMut.saltValue = saltValue;
    }
    const spinCount = attrInt(writeProtectionEl, "spinCount");
    if (spinCount !== undefined) {
      wpMut.spinCount = spinCount;
    }
    settings.writeProtection = wp;
  }

  // Document variables
  const docVarsEl = findChildNs(root, "docVars");
  if (docVarsEl) {
    const vars = new Map<string, string>();
    for (const dvEl of findChildrenNs(docVarsEl, "docVar")) {
      const name = attrVal(dvEl, "name");
      const val = attrVal(dvEl, "val");
      if (name !== undefined && val !== undefined) {
        vars.set(name, val);
      }
    }
    if (vars.size > 0) {
      settings.docVars = vars;
    }
  }

  // Footnote/endnote properties at document level
  const fnPrEl = findChildNs(root, "footnotePr");
  if (fnPrEl) {
    const fnProps = parseNoteProperties(fnPrEl);
    if (fnProps) {
      settings.footnoteProperties = fnProps as DocumentSettings["footnoteProperties"];
    }
  }
  const enPrEl = findChildNs(root, "endnotePr");
  if (enPrEl) {
    const enProps = parseNoteProperties(enPrEl);
    if (enProps) {
      settings.endnoteProperties = enProps as DocumentSettings["endnoteProperties"];
    }
  }

  return settings;
}
