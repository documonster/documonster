/**
 * DOCX Reader - Theme Parser
 *
 * Parses `word/theme/theme1.xml` (`DocumentTheme`).
 * Extracted from the original `metadata-parsers.ts`.
 */

import type { Mutable } from "@word/core/internal-utils";
import { findChildNs, serializeElement } from "@word/reader/parse-utils";
import type {
  DocumentTheme,
  ThemeColorName,
  ThemeFont,
  ThemeFontScheme,
  ThemeFormatScheme
} from "@word/types";
import { findChild, parseXml } from "@xml/dom";
import type { XmlElement } from "@xml/types";

const THEME_COLOR_NAMES: ThemeColorName[] = [
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink"
];

export function parseThemeXml(xmlStr: string): DocumentTheme {
  const doc = parseXml(xmlStr);
  const root = doc.root;

  // Find a:themeElements
  const themeElements = findChild(root, "a:themeElements") ?? findChildNs(root, "themeElements");

  const defaultScheme = {
    name: "Office",
    colors: {
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "44546A",
      lt2: "E7E6E6",
      accent1: "4472C4",
      accent2: "ED7D31",
      accent3: "A5A5A5",
      accent4: "FFC000",
      accent5: "5B9BD5",
      accent6: "70AD47",
      hlink: "0563C1",
      folHlink: "954F72"
    } as Record<ThemeColorName, string>
  };
  const defaultFontScheme = { name: "Office", majorFont: "Calibri Light", minorFont: "Calibri" };

  if (!themeElements) {
    return {
      name: root.attributes["name"],
      colorScheme: defaultScheme,
      fontScheme: defaultFontScheme
    };
  }

  // Parse color scheme
  const clrSchemeEl =
    findChild(themeElements, "a:clrScheme") ?? findChildNs(themeElements, "clrScheme");
  const colorScheme = { ...defaultScheme };
  if (clrSchemeEl) {
    colorScheme.name = clrSchemeEl.attributes["name"] ?? "Office";
    for (const colorName of THEME_COLOR_NAMES) {
      const colorEl =
        findChild(clrSchemeEl, `a:${colorName}`) ?? findChildNs(clrSchemeEl, colorName);
      if (colorEl) {
        // Color can be sysClr (with lastClr) or srgbClr (with val)
        const srgb = findChild(colorEl, "a:srgbClr") ?? findChildNs(colorEl, "srgbClr");
        if (srgb) {
          const val = srgb.attributes["val"];
          if (val) {
            colorScheme.colors[colorName] = val;
          }
        } else {
          const sys = findChild(colorEl, "a:sysClr") ?? findChildNs(colorEl, "sysClr");
          if (sys) {
            const lastClr = sys.attributes["lastClr"];
            if (lastClr) {
              colorScheme.colors[colorName] = lastClr;
            }
          }
        }
      }
    }
  }

  // Parse font scheme
  const fontSchemeEl =
    findChild(themeElements, "a:fontScheme") ?? findChildNs(themeElements, "fontScheme");
  const fontScheme: Mutable<ThemeFontScheme> = { ...defaultFontScheme };
  if (fontSchemeEl) {
    fontScheme.name = fontSchemeEl.attributes["name"] ?? "Office";
    const majorEl =
      findChild(fontSchemeEl, "a:majorFont") ?? findChildNs(fontSchemeEl, "majorFont");
    if (majorEl) {
      const major = parseThemeFont(majorEl);
      if (major) {
        fontScheme.major = major;
        if (major.latin) {
          fontScheme.majorFont = major.latin;
        }
      }
    }
    const minorEl =
      findChild(fontSchemeEl, "a:minorFont") ?? findChildNs(fontSchemeEl, "minorFont");
    if (minorEl) {
      const minor = parseThemeFont(minorEl);
      if (minor) {
        fontScheme.minor = minor;
        if (minor.latin) {
          fontScheme.minorFont = minor.latin;
        }
      }
    }
  }

  // Parse format scheme (preserve raw XML of its children for round-trip)
  const fmtSchemeEl =
    findChild(themeElements, "a:fmtScheme") ?? findChildNs(themeElements, "fmtScheme");
  let formatScheme: ThemeFormatScheme | undefined;
  if (fmtSchemeEl) {
    let rawXml = "";
    for (const child of fmtSchemeEl.children) {
      if (child.type === "element") {
        rawXml += serializeElement(child);
      }
    }
    formatScheme = {
      name: fmtSchemeEl.attributes["name"] ?? "Office",
      rawXml: rawXml || undefined
    };
  }

  // Preserve extLst (theme extensions) as raw XML
  let extLstXml: string | undefined;
  const extLstEl = findChild(root, "a:extLst") ?? findChildNs(root, "extLst");
  if (extLstEl) {
    extLstXml = serializeElement(extLstEl);
  }

  return {
    name: root.attributes["name"],
    colorScheme,
    fontScheme,
    formatScheme,
    extLstXml
  };
}

/** Parse a theme font (a:majorFont or a:minorFont). */
function parseThemeFont(el: XmlElement): ThemeFont | undefined {
  const font: Partial<Mutable<ThemeFont>> = {};
  const latin = findChild(el, "a:latin") ?? findChildNs(el, "latin");
  if (latin?.attributes["typeface"]) {
    font.latin = latin.attributes["typeface"];
  }
  const ea = findChild(el, "a:ea") ?? findChildNs(el, "ea");
  if (ea?.attributes["typeface"]) {
    font.eastAsia = ea.attributes["typeface"];
  }
  const cs = findChild(el, "a:cs") ?? findChildNs(el, "cs");
  if (cs?.attributes["typeface"]) {
    font.complexScript = cs.attributes["typeface"];
  }
  // Supplemental fonts (a:font script="..." typeface="...")
  const supplementalFonts: Record<string, string> = {};
  for (const child of el.children) {
    if (child.type === "element" && (child.name === "a:font" || child.name === "font")) {
      const script = child.attributes["script"];
      const typeface = child.attributes["typeface"];
      if (script && typeface) {
        supplementalFonts[script] = typeface;
      }
    }
  }
  if (Object.keys(supplementalFonts).length > 0) {
    font.supplementalFonts = supplementalFonts;
  }
  return font.latin !== undefined ? (font as ThemeFont) : undefined;
}
