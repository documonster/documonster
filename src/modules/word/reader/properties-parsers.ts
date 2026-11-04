/**
 * DOCX Reader - Properties Parsers
 *
 * Pure parsers for OOXML property structures that don't depend on the
 * main reader context. These are the building blocks used by the higher-level
 * paragraph/run/table/section parsers.
 *
 * Includes:
 * - `parseRunProperties` (w:rPr)
 * - `parseShading` (w:shd)
 * - `parseBorder` (any border element)
 * - `parseTableWidth` (w:tblW, w:tcW, etc.)
 * - `parseRevisionInfo` (w:ins/w:del/w:rPrChange/etc. revision metadata)
 */

import { setAsideValue } from "@utils/unrecognized-values";
import { WRITTEN, readToken } from "@word/core/argument-checks";
import type { Mutable } from "@word/core/internal-utils";
import {
  attrInt,
  attrOnOff,
  attrVal,
  boolToggle,
  findChildNs,
  safeParseInt
} from "@word/reader/parse-utils";
import type {
  Border,
  ColorSpec,
  EastAsianCombineBrackets,
  EastAsianLayoutSpec,
  FontSpec,
  RevisionInfo,
  RunProperties,
  RunPropertyChange,
  Shading,
  TableRowProperties,
  TableWidth,
  UnderlineStyle
} from "@word/types";
import type { XmlElement } from "@xml/types";

// =============================================================================
// Run Properties
// =============================================================================

export function parseRunProperties(rPrEl: XmlElement): RunProperties {
  const rPr: Mutable<RunProperties> & Record<string, unknown> = {};

  const rStyleEl = findChildNs(rPrEl, "rStyle");
  if (rStyleEl) {
    rPr.style = attrVal(rStyleEl, "val");
  }

  const fontsEl = findChildNs(rPrEl, "rFonts");
  if (fontsEl) {
    const f: Partial<Mutable<FontSpec>> = {};
    const ascii = attrVal(fontsEl, "ascii");
    const hAnsi = attrVal(fontsEl, "hAnsi");
    const eastAsia = attrVal(fontsEl, "eastAsia");
    const cs = attrVal(fontsEl, "cs");
    const hint = attrVal(fontsEl, "hint");
    if (ascii) {
      f.ascii = ascii;
    }
    if (hAnsi) {
      f.hAnsi = hAnsi;
    }
    if (eastAsia) {
      f.eastAsia = eastAsia;
    }
    if (cs) {
      f.cs = cs;
    }
    const knownHint = readToken(hint, WRITTEN.fontHint, f, "hint");
    if (knownHint) {
      f.hint = knownHint as FontSpec["hint"];
    }
    const asciiTheme = attrVal(fontsEl, "asciiTheme");
    if (asciiTheme) {
      f.asciiTheme = asciiTheme;
    }
    const hAnsiTheme = attrVal(fontsEl, "hAnsiTheme");
    if (hAnsiTheme) {
      f.hAnsiTheme = hAnsiTheme;
    }
    const eastAsiaTheme = attrVal(fontsEl, "eastAsiaTheme");
    if (eastAsiaTheme) {
      f.eastAsiaTheme = eastAsiaTheme;
    }
    const cstheme = attrVal(fontsEl, "cstheme");
    if (cstheme) {
      f.cstheme = cstheme;
    }
    rPr.font = f as FontSpec;
  }

  // Boolean toggles: true-only (element presence = true; absence = undefined)
  // Format: [elementName, propertyKey]
  const RUN_ONCE_TOGGLES: ReadonlyArray<[string, keyof RunProperties]> = [
    ["strike", "strike"],
    ["dstrike", "doubleStrike"],
    ["caps", "caps"],
    ["smallCaps", "smallCaps"],
    ["vanish", "vanish"],
    ["emboss", "emboss"],
    ["imprint", "imprint"],
    ["noProof", "noProof"],
    ["specVanish", "specVanish"],
    ["outline", "outline"],
    ["shadow", "shadow"],
    ["cs", "complexScript"],
    ["oMath", "math"],
    ["webHidden", "webHidden"]
  ];
  for (const [tag, key] of RUN_ONCE_TOGGLES) {
    if (findChildNs(rPrEl, tag) && boolToggle(rPrEl, tag) !== false) {
      (rPr as Record<string, unknown>)[key] = true;
    }
  }

  // Boolean toggles: tri-state (can be explicitly true or false)
  const RUN_TRISTATE_TOGGLES: ReadonlyArray<[string, keyof RunProperties]> = [
    ["b", "bold"],
    ["bCs", "boldCs"],
    ["i", "italic"],
    ["iCs", "italicCs"],
    ["snapToGrid", "snapToGrid"],
    ["rtl", "rightToLeft"]
  ];
  for (const [tag, key] of RUN_TRISTATE_TOGGLES) {
    const v = boolToggle(rPrEl, tag);
    if (v !== undefined) {
      (rPr as Record<string, unknown>)[key] = v;
    }
  }

  // fitText
  const fitTextEl = findChildNs(rPrEl, "fitText");
  if (fitTextEl) {
    const val = attrInt(fitTextEl, "val");
    if (val !== undefined) {
      const id = attrInt(fitTextEl, "id");
      rPr.fitText = id !== undefined ? { val, id } : { val };
    }
  }

  const colorEl = findChildNs(rPrEl, "color");
  if (colorEl) {
    const val = attrVal(colorEl, "val");
    const themeColor = attrVal(colorEl, "themeColor");
    if (themeColor) {
      const spec: Partial<Mutable<ColorSpec>> = { val, themeColor };
      const themeTint = attrVal(colorEl, "themeTint");
      const themeShade = attrVal(colorEl, "themeShade");
      if (themeTint) {
        spec.themeTint = themeTint;
      }
      if (themeShade) {
        spec.themeShade = themeShade;
      }
      rPr.color = spec as ColorSpec;
    } else {
      rPr.color = val;
    }
  }

  const szEl = findChildNs(rPrEl, "sz");
  if (szEl) {
    rPr.size = attrInt(szEl, "val");
  }

  const szCsEl = findChildNs(rPrEl, "szCs");
  if (szCsEl) {
    rPr.sizeCs = attrInt(szCsEl, "val");
  }

  const uEl = findChildNs(rPrEl, "u");
  if (uEl) {
    const uVal = attrVal(uEl, "val");
    const uColor = attrVal(uEl, "color");
    if (uVal !== undefined && !WRITTEN.underline.has(uVal)) {
      // A style outside ST_Underline is kept as read, with its colour, and the
      // run reads as not underlined until something sets one.
      setAsideValue(rPr, "underline", uVal, uColor ? { with: { color: uColor } } : {});
    } else if (uColor) {
      rPr.underline = { style: (uVal ?? "single") as UnderlineStyle, color: uColor };
    } else {
      rPr.underline = (uVal ?? "single") as UnderlineStyle;
    }
  }

  const highlightEl = findChildNs(rPrEl, "highlight");
  if (highlightEl) {
    rPr.highlight = readToken(
      attrVal(highlightEl, "val"),
      WRITTEN.highlight,
      rPr,
      "highlight"
    ) as RunProperties["highlight"];
  }

  const vertAlignEl = findChildNs(rPrEl, "vertAlign");
  if (vertAlignEl) {
    rPr.vertAlign = readToken(
      attrVal(vertAlignEl, "val"),
      WRITTEN.vertAlign,
      rPr,
      "vertAlign"
    ) as RunProperties["vertAlign"];
  }

  const spacingEl = findChildNs(rPrEl, "spacing");
  if (spacingEl) {
    rPr.spacing = attrInt(spacingEl, "val");
  }

  const shdEl = findChildNs(rPrEl, "shd");
  if (shdEl) {
    rPr.shading = parseShading(shdEl);
  }

  const langEl = findChildNs(rPrEl, "lang");
  if (langEl) {
    rPr.language = {
      val: attrVal(langEl, "val"),
      eastAsia: attrVal(langEl, "eastAsia"),
      bidi: attrVal(langEl, "bidi")
    };
  }

  // New valued properties
  const kernEl = findChildNs(rPrEl, "kern");
  if (kernEl) {
    rPr.kern = attrInt(kernEl, "val");
  }

  const positionEl = findChildNs(rPrEl, "position");
  if (positionEl) {
    rPr.position = attrInt(positionEl, "val");
  }

  const wEl = findChildNs(rPrEl, "w");
  if (wEl) {
    rPr.scale = attrInt(wEl, "val");
  }

  const effectEl = findChildNs(rPrEl, "effect");
  if (effectEl) {
    rPr.effect = readToken(
      attrVal(effectEl, "val"),
      WRITTEN.effect,
      rPr,
      "effect"
    ) as RunProperties["effect"];
  }

  const emEl = findChildNs(rPrEl, "em");
  if (emEl) {
    rPr.emphasisMark = readToken(
      attrVal(emEl, "val"),
      WRITTEN.emphasisMark,
      rPr,
      "emphasisMark"
    ) as RunProperties["emphasisMark"];
  }

  const eaLayoutEl = findChildNs(rPrEl, "eastAsianLayout");
  if (eaLayoutEl) {
    const ea = parseEastAsianLayout(eaLayoutEl);
    if (ea) {
      rPr.eastAsianLayout = ea;
    }
  }

  const bdrEl = findChildNs(rPrEl, "bdr");
  if (bdrEl) {
    rPr.border = parseBorder(bdrEl);
  }

  // rPrChange (track changes for run properties)
  const rPrChangeEl = findChildNs(rPrEl, "rPrChange");
  if (rPrChangeEl) {
    const rev = parseRevisionInfo(rPrChangeEl);
    if (rev) {
      const prevRPrEl = findChildNs(rPrChangeEl, "rPr");
      const change: RunPropertyChange = {
        revision: rev,
        previousProperties: prevRPrEl ? parseRunProperties(prevRPrEl) : undefined
      };
      rPr.propertyChange = change;
    }
  }

  return rPr;
}

// =============================================================================
// East Asian Layout
// =============================================================================

function parseEastAsianLayout(el: XmlElement): EastAsianLayoutSpec | undefined {
  const spec: Mutable<EastAsianLayoutSpec> = {};

  const id = attrInt(el, "id");
  if (id != null) {
    spec.id = id;
  }
  if (attrOnOff(el, "combine") === true) {
    spec.combine = true;
  }
  const brackets = readToken(
    attrVal(el, "combineBrackets"),
    WRITTEN.combineBrackets,
    spec,
    "combineBrackets"
  );
  if (brackets) {
    spec.combineBrackets = brackets as EastAsianCombineBrackets;
  }
  if (attrOnOff(el, "vert") === true) {
    spec.vert = true;
  }
  if (attrOnOff(el, "vertCompress") === true) {
    spec.vertCompress = true;
  }

  return Object.keys(spec).length > 0 ? (spec as EastAsianLayoutSpec) : undefined;
}

// =============================================================================
// Shading / Border / Width
// =============================================================================

export function parseShading(el: XmlElement): Shading {
  const shading: Mutable<Shading> = {
    color: attrVal(el, "color"),
    fill: attrVal(el, "fill") ?? "auto"
  };
  shading.pattern = readToken(
    attrVal(el, "val"),
    WRITTEN.shading,
    shading,
    "pattern"
  ) as Shading["pattern"];
  return shading;
}

export function parseBorder(el: XmlElement): Border {
  const b: Partial<Mutable<Border>> = {
    style: (attrVal(el, "val") ?? "single") as Border["style"],
    size: attrInt(el, "sz"),
    space: attrInt(el, "space"),
    color: attrVal(el, "color")
  };
  const tc = attrVal(el, "themeColor");
  if (tc) {
    b.themeColor = tc;
  }
  const shadow = attrVal(el, "shadow");
  if (shadow === "1" || shadow === "true") {
    b.shadow = true;
  }
  const frame = attrVal(el, "frame");
  if (frame === "1" || frame === "true") {
    b.frame = true;
  }
  const art = attrVal(el, "art");
  if (art) {
    b.art = art as Border["art"];
  }
  return b as Border;
}

export function parseTableWidth(el: XmlElement): TableWidth {
  const width: Mutable<TableWidth> = {
    // Reject non-numeric values (a hostile DOCX could write
    // `<w:tblW w:w="abc"/>`); falling back to 0 keeps the model field
    // a finite number and prevents `NaN` from being serialised back to
    // the output XML, which Word rejects.
    value: safeParseInt(el.attributes["w:w"] ?? el.attributes["w"], 0),
    type: "dxa"
  };
  width.type = readToken(
    el.attributes["w:type"] ?? el.attributes["type"],
    WRITTEN.tableWidthType,
    width,
    "type",
    "dxa"
  ) as TableWidth["type"];
  return width;
}

type RowHeight = NonNullable<TableRowProperties["height"]>;

/** `w:trHeight`. */
export function parseRowHeight(el: XmlElement): RowHeight {
  const height: Mutable<RowHeight> = { value: attrInt(el, "val") ?? 0 };
  height.rule = readToken(
    attrVal(el, "hRule"),
    WRITTEN.rowHeightRule,
    height,
    "rule"
  ) as RowHeight["rule"];
  return height;
}

// =============================================================================
// Revision Info
// =============================================================================

export function parseRevisionInfo(el: XmlElement): RevisionInfo | undefined {
  const author = attrVal(el, "author");
  const id = attrInt(el, "id");
  // ECMA-376 marks `w:author` and `w:id` as optional on revision elements.
  // Earlier we returned `undefined` if either was missing — that quietly
  // dropped the entire `<w:ins>`/`<w:del>` body in callers that gate on
  // this result. Fall back to a sentinel author/id so the surrounding runs
  // are still preserved (the user-visible text is what matters; the
  // metadata is purely informational).
  return {
    author: author ?? "",
    id: id ?? 0,
    date: attrVal(el, "date")
  };
}
