/**
 * Character width metrics for the 14 PDF standard Type 1 fonts.
 *
 * These metrics are specified in 1/1000 of a text unit. To get the actual width
 * of a character at a given font size:
 *   width_in_points = (charWidth / 1000) * fontSize
 *
 * The 14 standard fonts are guaranteed to be available in every PDF viewer:
 * - Helvetica, Helvetica-Bold, Helvetica-Oblique, Helvetica-BoldOblique
 * - Times-Roman, Times-Bold, Times-Italic, Times-BoldItalic
 * - Courier, Courier-Bold, Courier-Oblique, Courier-BoldOblique
 * - Symbol, ZapfDingbats
 *
 * We provide metrics for the most commonly needed: Helvetica and Times families,
 * plus Courier (fixed-width). For other fonts, we fall back to Helvetica metrics.
 *
 * Character widths are for the Latin-1 (ISO 8859-1) encoding.
 *
 * @see PDF Reference 1.7, Appendix D - Standard Type 1 Fonts
 * @see Adobe Font Metrics files (AFM) for canonical widths
 */

// =============================================================================
// Helvetica Metrics (afm data: Helvetica)
// =============================================================================

/**
 * Average character widths per font. We store only a representative subset:
 * ASCII 32-126. For characters outside this range, we use the average width.
 *
 * Values are in thousandths of a unit of text space.
 */

// prettier-ignore
const HELVETICA_WIDTHS: Record<number, number> = {
  32:278,33:278,34:355,35:556,36:556,37:889,38:667,39:191,40:333,41:333,
  42:389,43:584,44:278,45:333,46:278,47:278,48:556,49:556,50:556,51:556,
  52:556,53:556,54:556,55:556,56:556,57:556,58:278,59:278,60:584,61:584,
  62:584,63:556,64:1015,65:667,66:667,67:722,68:722,69:667,70:611,71:778,
  72:722,73:278,74:500,75:667,76:556,77:833,78:722,79:778,80:667,81:778,
  82:722,83:667,84:611,85:722,86:667,87:944,88:667,89:667,90:611,91:278,
  92:278,93:278,94:469,95:556,96:333,97:556,98:556,99:500,100:556,101:556,
  102:278,103:556,104:556,105:222,106:222,107:500,108:222,109:833,110:556,
  111:556,112:556,113:556,114:333,115:500,116:278,117:556,118:500,119:722,
  120:500,121:500,122:500,123:334,124:260,125:334,126:584
};

// prettier-ignore
const HELVETICA_BOLD_WIDTHS: Record<number, number> = {
  32:278,33:333,34:474,35:556,36:556,37:889,38:722,39:238,40:333,41:333,
  42:389,43:584,44:278,45:333,46:278,47:278,48:556,49:556,50:556,51:556,
  52:556,53:556,54:556,55:556,56:556,57:556,58:333,59:333,60:584,61:584,
  62:584,63:611,64:975,65:722,66:722,67:722,68:722,69:667,70:611,71:778,
  72:722,73:278,74:556,75:722,76:611,77:833,78:722,79:778,80:667,81:778,
  82:722,83:667,84:611,85:722,86:667,87:944,88:667,89:667,90:611,91:333,
  92:278,93:333,94:584,95:556,96:333,97:556,98:611,99:556,100:611,101:556,
  102:333,103:611,104:611,105:278,106:278,107:556,108:278,109:889,110:611,
  111:611,112:611,113:611,114:389,115:556,116:333,117:611,118:556,119:778,
  120:556,121:556,122:500,123:389,124:280,125:389,126:584
};

// prettier-ignore
const TIMES_ROMAN_WIDTHS: Record<number, number> = {
  32:250,33:333,34:408,35:500,36:500,37:833,38:778,39:180,40:333,41:333,
  42:500,43:564,44:250,45:333,46:250,47:278,48:500,49:500,50:500,51:500,
  52:500,53:500,54:500,55:500,56:500,57:500,58:278,59:278,60:564,61:564,
  62:564,63:444,64:921,65:722,66:667,67:667,68:722,69:611,70:556,71:722,
  72:722,73:333,74:389,75:722,76:611,77:889,78:722,79:722,80:556,81:722,
  82:667,83:556,84:611,85:722,86:722,87:944,88:722,89:722,90:611,91:333,
  92:278,93:333,94:469,95:500,96:333,97:444,98:500,99:444,100:500,101:444,
  102:333,103:500,104:500,105:278,106:278,107:500,108:278,109:778,110:500,
  111:500,112:500,113:500,114:333,115:389,116:278,117:500,118:500,119:722,
  120:500,121:500,122:444,123:480,124:200,125:480,126:541
};

// prettier-ignore
const TIMES_BOLD_WIDTHS: Record<number, number> = {
  32:250,33:333,34:555,35:500,36:500,37:1000,38:833,39:278,40:333,41:333,
  42:500,43:570,44:250,45:333,46:250,47:278,48:500,49:500,50:500,51:500,
  52:500,53:500,54:500,55:500,56:500,57:500,58:333,59:333,60:570,61:570,
  62:570,63:500,64:930,65:722,66:667,67:722,68:722,69:667,70:611,71:778,
  72:778,73:389,74:500,75:778,76:667,77:944,78:722,79:778,80:611,81:778,
  82:722,83:556,84:667,85:722,86:722,87:1000,88:722,89:722,90:667,91:333,
  92:278,93:333,94:581,95:500,96:333,97:500,98:556,99:444,100:556,101:444,
  102:333,103:500,104:556,105:278,106:333,107:556,108:278,109:833,110:556,
  111:500,112:556,113:556,114:444,115:389,116:333,117:556,118:500,119:722,
  120:500,121:500,122:444,123:394,124:220,125:394,126:520
};

// Courier is monospaced - every character is 600 units wide
const COURIER_WIDTH = 600;

// =============================================================================
// Font Descriptor Data
// =============================================================================

/**
 * Font descriptor data for ascent, descent, and other metrics.
 * Values are in font units (1/1000 of a text unit).
 */
interface FontDescriptor {
  ascent: number;
  descent: number;
  capHeight: number;
  avgWidth: number;
  widths: Record<number, number> | null; // null for monospaced
  monoWidth?: number;
}

const FONT_DESCRIPTORS: Record<string, FontDescriptor> = {
  Helvetica: {
    ascent: 718,
    descent: -207,
    capHeight: 718,
    avgWidth: 513,
    widths: HELVETICA_WIDTHS
  },
  "Helvetica-Bold": {
    ascent: 718,
    descent: -207,
    capHeight: 718,
    avgWidth: 535,
    widths: HELVETICA_BOLD_WIDTHS
  },
  "Helvetica-Oblique": {
    ascent: 718,
    descent: -207,
    capHeight: 718,
    avgWidth: 513,
    widths: HELVETICA_WIDTHS // same widths as regular
  },
  "Helvetica-BoldOblique": {
    ascent: 718,
    descent: -207,
    capHeight: 718,
    avgWidth: 535,
    widths: HELVETICA_BOLD_WIDTHS
  },
  "Times-Roman": {
    ascent: 683,
    descent: -217,
    capHeight: 662,
    avgWidth: 478,
    widths: TIMES_ROMAN_WIDTHS
  },
  "Times-Bold": {
    ascent: 683,
    descent: -217,
    capHeight: 676,
    avgWidth: 505,
    widths: TIMES_BOLD_WIDTHS
  },
  "Times-Italic": {
    ascent: 683,
    descent: -217,
    capHeight: 653,
    avgWidth: 478,
    widths: TIMES_ROMAN_WIDTHS // approximate
  },
  "Times-BoldItalic": {
    ascent: 683,
    descent: -217,
    capHeight: 669,
    avgWidth: 505,
    widths: TIMES_BOLD_WIDTHS // approximate
  },
  Courier: {
    ascent: 629,
    descent: -157,
    capHeight: 562,
    avgWidth: 600,
    widths: null,
    monoWidth: COURIER_WIDTH
  },
  "Courier-Bold": {
    ascent: 629,
    descent: -157,
    capHeight: 562,
    avgWidth: 600,
    widths: null,
    monoWidth: COURIER_WIDTH
  },
  "Courier-Oblique": {
    ascent: 629,
    descent: -157,
    capHeight: 562,
    avgWidth: 600,
    widths: null,
    monoWidth: COURIER_WIDTH
  },
  "Courier-BoldOblique": {
    ascent: 629,
    descent: -157,
    capHeight: 562,
    avgWidth: 600,
    widths: null,
    monoWidth: COURIER_WIDTH
  }
};

// =============================================================================
// Public API
// =============================================================================

/**
 * Get the width of a character in a given font.
 * @param charCode - Unicode code point (or char code)
 * @param fontName - PDF standard font name
 * @returns Width in thousandths of a unit
 */
export function getCharWidth(charCode: number, fontName: string): number {
  const desc = FONT_DESCRIPTORS[fontName];
  if (!desc) {
    // Fall back to Helvetica
    return getCharWidth(charCode, "Helvetica");
  }
  if (desc.monoWidth !== undefined) {
    return desc.monoWidth;
  }
  return desc.widths?.[charCode] ?? desc.avgWidth;
}

/**
 * Measure the width of a text string in the given font and size.
 * @param text - The string to measure
 * @param fontName - PDF standard font name
 * @param fontSize - Font size in points
 * @returns Width in points
 */
export function measureText(text: string, fontName: string, fontSize: number): number {
  let totalWidth = 0;
  for (let i = 0; i < text.length; i++) {
    totalWidth += getCharWidth(text.charCodeAt(i), fontName);
  }
  return (totalWidth / 1000) * fontSize;
}

/**
 * Get the font ascent for a given font and size.
 * @returns Ascent in points (positive, distance above baseline)
 */
export function getFontAscent(fontName: string, fontSize: number): number {
  const desc = FONT_DESCRIPTORS[fontName] ?? FONT_DESCRIPTORS["Helvetica"];
  return (desc.ascent / 1000) * fontSize;
}

/**
 * Get the font descent for a given font and size.
 * @returns Descent in points (negative, distance below baseline)
 */
export function getFontDescent(fontName: string, fontSize: number): number {
  const desc = FONT_DESCRIPTORS[fontName] ?? FONT_DESCRIPTORS["Helvetica"];
  return (desc.descent / 1000) * fontSize;
}

/**
 * Get the total line height (ascent - descent) for a font.
 * @returns Line height in points
 */
export function getLineHeight(fontName: string, fontSize: number): number {
  const desc = FONT_DESCRIPTORS[fontName] ?? FONT_DESCRIPTORS["Helvetica"];
  return ((desc.ascent - desc.descent) / 1000) * fontSize;
}

/**
 * Check if a font name is a known standard PDF font.
 */
export function isStandardFont(fontName: string): boolean {
  return fontName in FONT_DESCRIPTORS;
}

/**
 * Get all supported standard font names.
 */
export function getStandardFontNames(): string[] {
  return Object.keys(FONT_DESCRIPTORS);
}
