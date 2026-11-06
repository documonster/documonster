/**
 * DOCX Reader - Settings Parser
 *
 * Parses `word/settings.xml` (DocumentSettings).
 * Extracted from the original `metadata-parsers.ts`.
 */

import type { Mutable } from "@word/core/internal-utils";
import {
  attrInt,
  attrOnOff,
  attrVal,
  boolToggle,
  findChildNs,
  findChildrenNs,
  parseNoteProperties,
  serializeElement,
  sidToHashAlgorithm
} from "@word/reader/parse-utils";
import { parsePartXml } from "@word/reader/part-xml";
import { preservedRootXml, wordLocalName } from "@word/reader/preserved-xml";
import type {
  CompatFlag,
  CompatSetting,
  DocumentSettings,
  HyphenationSettings,
  NoteSeparatorType,
  ProtectionType
} from "@word/types";
import type { XmlElement } from "@xml/types";

/**
 * Parse `word/settings.xml`. `noteTypes` maps each footnote/endnote id to its
 * `w:type` in the notes parts, so the separator references in `w:footnotePr`
 * and `w:endnotePr` can be kept by type — the writer renumbers those notes.
 */
export function parseSettingsXml(
  xmlStr: string,
  maxDepth?: number,
  noteTypes?: {
    readonly footnote?: ReadonlyMap<number, string>;
    readonly endnote?: ReadonlyMap<number, string>;
  }
): DocumentSettings {
  const doc = parsePartXml(xmlStr, maxDepth);
  const root = doc.root;
  const settings: Mutable<DocumentSettings> = {};
  /** An on/off (`ST_OnOff`) setting element, when present: `w:val` false, 0 or off is false. */
  const onOff = (name: string, set: (value: boolean) => void): void => {
    const value = boolToggle(root, name);
    if (value !== undefined) {
      set(value);
    }
  };

  const zoomEl = findChildNs(root, "zoom");
  if (zoomEl) {
    const percent = attrInt(zoomEl, "percent");
    if (percent !== undefined) {
      settings.zoom = percent;
    }
    const zoomType = attrVal(zoomEl, "val");
    if (
      zoomType === "none" ||
      zoomType === "fullPage" ||
      zoomType === "bestFit" ||
      zoomType === "textFit"
    ) {
      settings.zoomType = zoomType;
    }
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
  onOff("doNotTrackMoves", v => (settings.doNotTrackMoves = v));
  onOff("doNotTrackFormatting", v => (settings.doNotTrackFormatting = v));
  onOff("doNotDemoteNonCombiningChars", v => (settings.doNotDemoteAsianTextFirstLine = v));
  onOff("saveSubsetFonts", v => (settings.saveSubsetFonts = v));
  onOff("noPunctuationKerning", v => (settings.noPunctuationKerning = v));
  onOff("bordersDoNotSurroundHeader", v => (settings.bordersDoNotSurroundHeader = v));
  onOff("bordersDoNotSurroundFooter", v => (settings.bordersDoNotSurroundFooter = v));
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

  onOff("evenAndOddHeaders", v => (settings.evenAndOddHeaders = v));

  onOff("trackRevisions", v => (settings.trackRevisions = v));

  onOff("mirrorMargins", v => (settings.mirrorMargins = v));

  onOff("gutterAtTop", v => (settings.gutterAtTop = v));

  onOff("displayBackgroundShape", v => (settings.displayBackgroundShape = v));

  onOff("updateFields", v => (settings.updateFieldsOnOpen = v));

  // Hyphenation
  const autoHyphenation = boolToggle(root, "autoHyphenation");
  if (autoHyphenation === false) {
    settings.autoHyphenation = false;
  } else if (autoHyphenation) {
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
    const doNotHyphenateCaps = boolToggle(root, "doNotHyphenateCaps");
    if (doNotHyphenateCaps !== undefined) {
      hyph.doNotHyphenateCaps = doNotHyphenateCaps;
    }
    settings.hyphenation = hyph;
  }

  // Document protection
  const protEl = findChildNs(root, "documentProtection");
  if (protEl) {
    settings.documentProtection = {
      type: (attrVal(protEl, "edit") ?? "none") as ProtectionType,
      edit: attrVal(protEl, "edit") ?? undefined,
      enforcement: attrOnOff(protEl, "enforcement") === true,
      formatting: attrOnOff(protEl, "formatting") === true ? true : undefined,
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
    if (attrOnOff(writeProtectionEl, "recommended") === true) {
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

  const footnoteSeparators = fnPrEl
    ? separatorReferences(fnPrEl, "footnote", noteTypes?.footnote)
    : [];
  if (footnoteSeparators.length > 0) {
    settings.footnoteSeparators = footnoteSeparators;
  }
  const endnoteSeparators = enPrEl
    ? separatorReferences(enPrEl, "endnote", noteTypes?.endnote)
    : [];
  if (endnoteSeparators.length > 0) {
    settings.endnoteSeparators = endnoteSeparators;
  }

  const hyphenationModelled = settings.hyphenation !== undefined;
  const preservedXml = preservedRootXml(root, child => {
    const local = wordLocalName(child);
    return (
      local !== undefined &&
      (MODELLED_SETTINGS.has(local) || (hyphenationModelled && HYPHENATION_DETAIL.has(local)))
    );
  });
  if (preservedXml) {
    settings.preservedXml = preservedXml;
  }

  return settings;
}

/** The types of the special notes a `w:footnotePr`/`w:endnotePr` references, in order. */
function separatorReferences(
  el: XmlElement,
  kind: "footnote" | "endnote",
  types: ReadonlyMap<number, string> | undefined
): NoteSeparatorType[] {
  const out: NoteSeparatorType[] = [];
  for (const ref of findChildrenNs(el, kind)) {
    const id = attrInt(ref, "id");
    const type = id === undefined ? undefined : types?.get(id);
    if (type === "separator" || type === "continuationSeparator" || type === "continuationNotice") {
      out.push(type);
    }
  }
  return out;
}

/** The `w:settings` children this parser reads into the model; the writer writes them back. */
const MODELLED_SETTINGS = new Set([
  "zoom",
  "defaultTabStop",
  "characterSpacingControl",
  "doNotTrackMoves",
  "doNotTrackFormatting",
  "doNotDemoteNonCombiningChars",
  "saveSubsetFonts",
  "noPunctuationKerning",
  "bordersDoNotSurroundHeader",
  "bordersDoNotSurroundFooter",
  "clickAndTypeStyle",
  "stylePaneFormatFilter",
  "stylePaneSortMethod",
  "themeFontLang",
  "decimalSymbol",
  "listSeparator",
  "rsids",
  "evenAndOddHeaders",
  "trackRevisions",
  "mirrorMargins",
  "gutterAtTop",
  "displayBackgroundShape",
  "updateFields",
  "autoHyphenation",
  "documentProtection",
  "compat",
  "mailMerge",
  "writeProtection",
  "docVars",
  "footnotePr",
  "endnotePr"
]);

/** Hyphenation details, read only while automatic hyphenation is on — kept verbatim otherwise. */
const HYPHENATION_DETAIL = new Set([
  "consecutiveHyphenLimit",
  "hyphenationZone",
  "doNotHyphenateCaps"
]);
