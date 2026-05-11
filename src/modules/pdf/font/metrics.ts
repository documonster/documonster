/**
 * Character width metrics for the 14 PDF standard Type 1 fonts.
 *
 * This module re-exports from the shared @utils/font-metrics to avoid
 * data duplication while maintaining backward compatibility.
 *
 * @see PDF Reference 1.7, Appendix D - Standard Type 1 Fonts
 * @see Adobe Font Metrics files (AFM) for canonical widths
 */

export {
  measureTextWidth as measureText,
  getCharWidth,
  getFontAscent,
  getFontDescent,
  getLineHeight,
  isStandardFont,
  getStandardFontNames
} from "@utils/font-metrics";
