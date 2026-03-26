/**
 * Font metrics for the 14 standard PDF fonts.
 *
 * Width tables are in 1/1000 of a unit for each glyph (cp1252 encoding).
 * These tables allow accurate text width measurement without external dependencies.
 *
 * Reference: PDF 1.7 spec Appendix D, Adobe Font Metrics files.
 */

// =============================================================================
// Types
// =============================================================================

interface FontWidths {
  /** Average character width in 1/1000 units */
  avg: number;
  /** Widths for cp1252 chars 32..126 (printable ASCII) */
  widths: number[];
}

interface FontMetricsEntry {
  normal: FontWidths;
  bold: FontWidths;
  italic: FontWidths;
  boldItalic: FontWidths;
}

// =============================================================================
// Helvetica (sans-serif, used as fallback for Arial/Calibri)
// =============================================================================

const helveticaNormal: FontWidths = {
  avg: 552,
  // Space(32)..tilde(126) widths
  widths: [
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556,
    556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667,
    611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667,
    667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500,
    222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584
  ]
};

const helveticaBold: FontWidths = {
  avg: 579,
  widths: [
    278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556,
    556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667,
    611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667,
    667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556,
    278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584
  ]
};

const helveticaItalic: FontWidths = {
  avg: 552,
  widths: [
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556,
    556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667,
    611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667,
    667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500,
    222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584
  ]
};

const helveticaBoldItalic: FontWidths = {
  avg: 579,
  widths: [
    278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556,
    556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667,
    611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667,
    667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556,
    278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584
  ]
};

// =============================================================================
// Times (serif)
// =============================================================================

const timesNormal: FontWidths = {
  avg: 521,
  widths: [
    250, 333, 408, 500, 500, 833, 778, 180, 333, 333, 500, 564, 250, 333, 250, 278, 500, 500, 500,
    500, 500, 500, 500, 500, 500, 500, 278, 278, 564, 564, 564, 444, 921, 722, 667, 667, 722, 611,
    556, 722, 722, 333, 389, 722, 611, 889, 722, 722, 556, 722, 667, 556, 611, 722, 722, 944, 722,
    722, 611, 333, 278, 333, 469, 500, 333, 444, 500, 444, 500, 444, 333, 500, 500, 278, 278, 500,
    278, 778, 500, 500, 500, 500, 333, 389, 278, 500, 500, 722, 500, 500, 444, 480, 200, 480, 541
  ]
};

const timesBold: FontWidths = {
  avg: 547,
  widths: [
    250, 333, 555, 500, 500, 1000, 833, 278, 333, 333, 500, 570, 250, 333, 250, 278, 500, 500, 500,
    500, 500, 500, 500, 500, 500, 500, 333, 333, 570, 570, 570, 500, 930, 722, 667, 722, 722, 667,
    611, 778, 778, 389, 500, 778, 667, 944, 722, 778, 611, 778, 722, 556, 667, 722, 722, 1000, 722,
    722, 667, 333, 278, 333, 581, 500, 333, 500, 556, 444, 556, 444, 333, 500, 556, 278, 333, 556,
    278, 833, 556, 500, 556, 556, 444, 389, 333, 556, 500, 722, 500, 500, 444, 394, 220, 394, 520
  ]
};

const timesItalic: FontWidths = {
  avg: 521,
  widths: [
    250, 333, 420, 500, 500, 833, 778, 214, 333, 333, 500, 675, 250, 333, 250, 278, 500, 500, 500,
    500, 500, 500, 500, 500, 500, 500, 333, 333, 675, 675, 675, 500, 920, 611, 611, 667, 722, 611,
    611, 722, 722, 333, 444, 667, 556, 833, 667, 722, 611, 722, 611, 500, 556, 722, 611, 833, 611,
    556, 556, 389, 278, 389, 422, 500, 333, 500, 500, 444, 500, 444, 278, 500, 500, 278, 278, 444,
    278, 722, 500, 500, 500, 500, 389, 389, 278, 500, 444, 667, 444, 444, 389, 400, 275, 400, 541
  ]
};

const timesBoldItalic: FontWidths = {
  avg: 547,
  widths: [
    250, 389, 555, 500, 500, 833, 778, 278, 333, 333, 500, 570, 250, 333, 250, 278, 500, 500, 500,
    500, 500, 500, 500, 500, 500, 500, 333, 333, 570, 570, 570, 500, 832, 667, 667, 667, 722, 667,
    667, 722, 778, 389, 500, 667, 611, 889, 722, 722, 611, 722, 667, 556, 611, 722, 667, 889, 667,
    611, 611, 333, 278, 333, 570, 500, 333, 500, 500, 444, 500, 444, 333, 500, 556, 278, 278, 500,
    278, 778, 556, 500, 500, 500, 389, 389, 278, 556, 444, 667, 500, 444, 389, 348, 220, 348, 570
  ]
};

// =============================================================================
// Courier (monospace)
// =============================================================================

const courierNormal: FontWidths = {
  avg: 600,
  widths: new Array(95).fill(600)
};

const courierBold: FontWidths = { ...courierNormal };
const courierItalic: FontWidths = { ...courierNormal };
const courierBoldItalic: FontWidths = { ...courierNormal };

// =============================================================================
// Font Metrics Map
// =============================================================================

const FONT_METRICS: Record<string, FontMetricsEntry> = {
  helvetica: {
    normal: helveticaNormal,
    bold: helveticaBold,
    italic: helveticaItalic,
    boldItalic: helveticaBoldItalic
  },
  times: {
    normal: timesNormal,
    bold: timesBold,
    italic: timesItalic,
    boldItalic: timesBoldItalic
  },
  courier: {
    normal: courierNormal,
    bold: courierBold,
    italic: courierItalic,
    boldItalic: courierBoldItalic
  }
};

// =============================================================================
// Excel font name → PDF base font mapping
// =============================================================================

const FONT_FAMILY_MAP: Record<string, string> = {
  // Sans-serif families → Helvetica
  arial: "helvetica",
  calibri: "helvetica",
  "segoe ui": "helvetica",
  tahoma: "helvetica",
  verdana: "helvetica",
  "trebuchet ms": "helvetica",
  "gill sans": "helvetica",
  "franklin gothic": "helvetica",
  helvetica: "helvetica",
  "helvetica neue": "helvetica",
  "open sans": "helvetica",
  roboto: "helvetica",
  lato: "helvetica",
  // Serif families → Times
  "times new roman": "times",
  times: "times",
  georgia: "times",
  "book antiqua": "times",
  palatino: "times",
  "palatino linotype": "times",
  garamond: "times",
  cambria: "times",
  // Monospace families → Courier
  "courier new": "courier",
  courier: "courier",
  consolas: "courier",
  "lucida console": "courier",
  "source code pro": "courier",
  monaco: "courier",
  menlo: "courier"
};

// =============================================================================
// Public API
// =============================================================================

/**
 * Map an Excel font name to a PDF base font family
 */
export function mapFontFamily(excelFontName?: string): string {
  if (!excelFontName) {
    return "helvetica";
  }
  return FONT_FAMILY_MAP[excelFontName.toLowerCase()] ?? "helvetica";
}

/**
 * Get the PDF base font name for given font properties.
 *
 * Returns one of the 14 standard PDF fonts:
 * - Helvetica, Helvetica-Bold, Helvetica-Oblique, Helvetica-BoldOblique
 * - Times-Roman, Times-Bold, Times-Italic, Times-BoldItalic
 * - Courier, Courier-Bold, Courier-Oblique, Courier-BoldOblique
 */
export function getPdfFontName(family: string, bold: boolean, italic: boolean): string {
  switch (family) {
    case "helvetica":
      if (bold && italic) {
        return "Helvetica-BoldOblique";
      }
      if (bold) {
        return "Helvetica-Bold";
      }
      if (italic) {
        return "Helvetica-Oblique";
      }
      return "Helvetica";
    case "times":
      if (bold && italic) {
        return "Times-BoldItalic";
      }
      if (bold) {
        return "Times-Bold";
      }
      if (italic) {
        return "Times-Italic";
      }
      return "Times-Roman";
    case "courier":
      if (bold && italic) {
        return "Courier-BoldOblique";
      }
      if (bold) {
        return "Courier-Bold";
      }
      if (italic) {
        return "Courier-Oblique";
      }
      return "Courier";
    default:
      return "Helvetica";
  }
}

/**
 * Get the style variant key (normal/bold/italic/boldItalic)
 */
function getStyleKey(bold: boolean, italic: boolean): keyof FontMetricsEntry {
  if (bold && italic) {
    return "boldItalic";
  }
  if (bold) {
    return "bold";
  }
  if (italic) {
    return "italic";
  }
  return "normal";
}

/**
 * Measure text width in PDF points for a given string, font, and size.
 */
export function measureTextWidth(
  text: string,
  family: string,
  bold: boolean,
  italic: boolean,
  fontSize: number
): number {
  const metrics = FONT_METRICS[family] ?? FONT_METRICS.helvetica;
  const styleKey = getStyleKey(bold, italic);
  const fw = metrics[styleKey];
  const { widths, avg } = fw;

  let totalWidth = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.codePointAt(i) ?? 0;
    if (code >= 32 && code <= 126) {
      totalWidth += widths[code - 32];
    } else {
      totalWidth += avg;
    }
  }

  return (totalWidth * fontSize) / 1000;
}

/**
 * Wrap text into lines that fit within a given width.
 * Returns an array of lines.
 */
export function wrapText(
  text: string,
  family: string,
  bold: boolean,
  italic: boolean,
  fontSize: number,
  maxWidth: number
): string[] {
  if (maxWidth <= 0) {
    return [text];
  }

  const lines: string[] = [];
  // Split on explicit line breaks first
  const paragraphs = text.split(/\r?\n/);

  for (const paragraph of paragraphs) {
    if (paragraph === "") {
      lines.push("");
      continue;
    }

    const words = paragraph.split(/\s+/);
    let currentLine = "";

    for (const word of words) {
      if (!currentLine) {
        currentLine = word;
        continue;
      }
      const testLine = currentLine + " " + word;
      const testWidth = measureTextWidth(testLine, family, bold, italic, fontSize);
      if (testWidth > maxWidth) {
        lines.push(currentLine);
        currentLine = word;
      } else {
        currentLine = testLine;
      }
    }

    if (currentLine) {
      lines.push(currentLine);
    }
  }

  return lines.length ? lines : [""];
}
