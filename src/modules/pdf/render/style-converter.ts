/**
 * Converts input styles to PDF rendering parameters.
 *
 * Maps font, color, border, fill, and alignment properties
 * to their PDF equivalents for the layout engine and page renderer.
 *
 * This module is fully independent of the Excel module — it works with
 * the PDF module's own style interfaces (PdfFontStyle, PdfFillData, etc.).
 */

import type {
  PdfColor,
  LayoutBorder,
  LayoutBorders,
  PdfColorData,
  PdfFontStyle,
  PdfFillData,
  PdfBorderSideData,
  PdfBordersData,
  PdfAlignmentData
} from "../types";

// =============================================================================
// Color Conversion
// =============================================================================

/**
 * Convert an ARGB color string to PDF RGB color.
 * Handles "AARRGGBB" (8-char) and "RRGGBB" (6-char) formats.
 * PDF uses 0-1 floats for each component.
 */
export function argbToPdfColor(argb: string | undefined): PdfColor | null {
  if (!argb || argb.length < 6) {
    return null;
  }

  let a: number;
  let r: number;
  let g: number;
  let b: number;

  if (argb.length >= 8) {
    // AARRGGBB format
    a = parseInt(argb.substring(0, 2), 16);
    r = parseInt(argb.substring(2, 4), 16);
    g = parseInt(argb.substring(4, 6), 16);
    b = parseInt(argb.substring(6, 8), 16);
  } else {
    // RRGGBB format
    a = 255;
    r = parseInt(argb.substring(0, 2), 16);
    g = parseInt(argb.substring(2, 4), 16);
    b = parseInt(argb.substring(4, 6), 16);
  }

  if (isNaN(r) || isNaN(g) || isNaN(b)) {
    return null;
  }

  const alpha = a / 255;
  return {
    r: r / 255,
    g: g / 255,
    b: b / 255,
    ...(alpha < 1 ? { a: alpha } : {})
  };
}

/**
 * Convert a color data object to PDF color.
 * Handles both ARGB and theme-based colors.
 */
export function excelColorToPdf(color: Partial<PdfColorData> | undefined): PdfColor | null {
  if (!color) {
    return null;
  }

  // ARGB takes priority
  if (color.argb) {
    return argbToPdfColor(color.argb);
  }

  // Theme colors with optional tint
  if (color.theme !== undefined) {
    const base = themeColorToPdf(color.theme);
    if (!base) {
      return null;
    }
    const tint = color.tint;
    if (tint !== undefined && tint !== 0) {
      return applyTint(base, tint);
    }
    return base;
  }

  return null;
}

/**
 * Map theme color indices to PDF colors.
 * These are the default Office theme colors.
 */
function themeColorToPdf(themeIndex: number): PdfColor | null {
  // Default Office 2019+ theme color palette (hex values verified)
  const themeColors: PdfColor[] = [
    { r: 1, g: 1, b: 1 }, // 0: lt1 — #FFFFFF (white / window background)
    { r: 0, g: 0, b: 0 }, // 1: dk1 — #000000 (black / window text)
    { r: 0.906, g: 0.902, b: 0.902 }, // 2: lt2 — #E7E6E6
    { r: 0.267, g: 0.329, b: 0.416 }, // 3: dk2 — #44546A
    { r: 0.267, g: 0.447, b: 0.769 }, // 4: accent1 — #4472C4 (blue)
    { r: 0.929, g: 0.49, b: 0.192 }, // 5: accent2 — #ED7D31 (orange)
    { r: 0.647, g: 0.647, b: 0.647 }, // 6: accent3 — #A5A5A5 (gray)
    { r: 1, g: 0.753, b: 0 }, // 7: accent4 — #FFC000 (gold)
    { r: 0.357, g: 0.608, b: 0.835 }, // 8: accent5 — #5B9BD5 (light blue)
    { r: 0.439, g: 0.678, b: 0.278 } // 9: accent6 — #70AD47 (green)
  ];

  if (themeIndex >= 0 && themeIndex < themeColors.length) {
    return themeColors[themeIndex];
  }
  return null;
}

/**
 * Apply a tint value to a color.
 * Tint range: -1.0 (fully dark) to +1.0 (fully light).
 * Negative tint darkens, positive tint lightens.
 *
 * @see OOXML §18.8.19 - tint formula
 */
export function applyTint(color: PdfColor, tint: number): PdfColor {
  const apply = (c: number) => {
    if (tint < 0) {
      return c * (1 + tint);
    }
    return c + (1 - c) * tint;
  };
  return {
    r: Math.max(0, Math.min(1, apply(color.r))),
    g: Math.max(0, Math.min(1, apply(color.g))),
    b: Math.max(0, Math.min(1, apply(color.b)))
  };
}

/**
 * Default colors used in PDF rendering.
 */
export const DEFAULT_COLORS = {
  black: { r: 0, g: 0, b: 0 } as PdfColor,
  white: { r: 1, g: 1, b: 1 } as PdfColor,
  gridLine: { r: 0.816, g: 0.816, b: 0.816 } as PdfColor
};

// =============================================================================
// Font Conversion
// =============================================================================

/**
 * Extract font properties for PDF rendering.
 */
export function extractFontProperties(
  font: Partial<PdfFontStyle> | undefined,
  defaultFamily: string,
  defaultSize: number
): {
  fontFamily: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  underline: boolean;
  textColor: PdfColor;
} {
  const fontFamily = font?.name ?? defaultFamily;
  const fontSize = font?.size ?? defaultSize;
  const bold = font?.bold ?? false;
  const italic = font?.italic ?? false;
  const strike = font?.strike ?? false;
  const underline = !!(font?.underline && font.underline !== "none");
  const textColor = excelColorToPdf(font?.color) ?? DEFAULT_COLORS.black;

  return { fontFamily, fontSize, bold, italic, strike, underline, textColor };
}

// =============================================================================
// Fill Conversion
// =============================================================================

/**
 * Convert a fill to a PDF background color.
 * Only pattern fills with "solid" pattern are supported as PDF backgrounds.
 * Other patterns are approximated or ignored.
 */
export function excelFillToPdfColor(fill: PdfFillData | undefined): PdfColor | null {
  if (!fill) {
    return null;
  }

  if (fill.type === "pattern") {
    if (fill.pattern === "solid" && fill.fgColor) {
      return excelColorToPdf(fill.fgColor);
    }
    if (fill.pattern === "none") {
      return null;
    }
    // For other patterns, use fgColor as approximation
    if (fill.fgColor) {
      return excelColorToPdf(fill.fgColor);
    }
  }

  if (fill.type === "gradient") {
    // For gradient fills, use the first stop color as approximation
    if (fill.stops && fill.stops.length > 0) {
      return excelColorToPdf(fill.stops[0].color);
    }
  }

  return null;
}

// =============================================================================
// Border Conversion
// =============================================================================

/**
 * Map border styles to PDF line widths.
 */
export function borderStyleToLineWidth(style: string): number {
  switch (style) {
    case "thin":
      return 0.25;
    case "medium":
      return 0.5;
    case "thick":
      return 1;
    case "double":
      return 0.25;
    case "hair":
      return 0.1;
    case "dotted":
      return 0.25;
    case "dashed":
      return 0.25;
    case "dashDot":
      return 0.25;
    case "dashDotDot":
      return 0.25;
    case "slantDashDot":
      return 0.25;
    case "mediumDashed":
      return 0.5;
    case "mediumDashDot":
      return 0.5;
    case "mediumDashDotDot":
      return 0.5;
    default:
      return 0.25;
  }
}

/**
 * Map border styles to PDF dash patterns.
 * An empty array means a solid line.
 */
function borderStyleToDashPattern(style: string): number[] {
  switch (style) {
    case "dotted":
      return [1, 1];
    case "dashed":
    case "mediumDashed":
      return [3, 2];
    case "dashDot":
    case "mediumDashDot":
      return [3, 1, 1, 1];
    case "dashDotDot":
    case "mediumDashDotDot":
      return [3, 1, 1, 1, 1, 1];
    case "slantDashDot":
      return [4, 2, 1, 2];
    case "hair":
      return [0.5, 0.5];
    default:
      return [];
  }
}

/**
 * Convert a single border side to a PDF LayoutBorder.
 */
function convertBorder(border: Partial<PdfBorderSideData> | undefined): LayoutBorder | null {
  if (!border || !border.style) {
    return null;
  }

  return {
    width: borderStyleToLineWidth(border.style),
    color: excelColorToPdf(border.color) ?? DEFAULT_COLORS.black,
    dashPattern: borderStyleToDashPattern(border.style),
    isDouble: border.style === "double"
  };
}

/**
 * Convert border data to PDF LayoutBorders.
 */
export function excelBordersToPdf(borders: Partial<PdfBordersData> | undefined): LayoutBorders {
  if (!borders) {
    return { top: null, right: null, bottom: null, left: null };
  }

  return {
    top: convertBorder(borders.top),
    right: convertBorder(borders.right),
    bottom: convertBorder(borders.bottom),
    left: convertBorder(borders.left)
  };
}

// =============================================================================
// Alignment Conversion
// =============================================================================

/**
 * Convert horizontal alignment to PDF alignment.
 */
export function excelHAlignToPdf(
  alignment: Partial<PdfAlignmentData> | undefined
): "left" | "center" | "right" {
  if (!alignment?.horizontal) {
    return "left";
  }

  switch (alignment.horizontal) {
    case "center":
    case "centerContinuous":
      return "center";
    case "right":
      return "right";
    case "left":
    case "fill":
    case "justify":
    case "distributed":
    default:
      return "left";
  }
}

/**
 * Convert vertical alignment to PDF alignment.
 */
export function excelVAlignToPdf(
  alignment: Partial<PdfAlignmentData> | undefined
): "top" | "middle" | "bottom" {
  if (!alignment?.vertical) {
    return "bottom"; // Default is bottom
  }

  switch (alignment.vertical) {
    case "top":
      return "top";
    case "middle":
      return "middle";
    case "bottom":
    case "justify":
    case "distributed":
    default:
      return "bottom";
  }
}
