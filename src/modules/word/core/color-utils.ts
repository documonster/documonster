/**
 * DOCX Module - Theme Color Utilities
 *
 * Resolves OOXML theme colors with tint/shade transformations.
 * Extracted to a standalone file so that html-renderer and document.ts
 * can both import it without creating circular heavy dependencies.
 *
 * Core tint/shade/resolve logic lives in @utils/theme-colors; this module
 * re-exports and adapts it for Word's ColorSpec / DocumentTheme types.
 */

import {
  applyTint,
  applyShade,
  resolveOoxmlThemeColor,
  THEME_COLOR_ATTRIBUTE_MAP
} from "@utils/theme-colors";
import type { OoxmlThemeColorName } from "@utils/theme-colors";

import type { DocumentTheme, ColorSpec, HexColor } from "../types";

// Re-export shared utilities so existing consumers that import from here
// continue to work without changes.
export { applyTint, applyShade, THEME_COLOR_ATTRIBUTE_MAP };
export type { OoxmlThemeColorName };

/**
 * Resolve a ColorSpec to an actual hex RGB color using the document theme.
 *
 * Applies theme color lookup + tint/shade transformations per OOXML spec.
 *
 * @param color - The color value (HexColor string or ColorSpec).
 * @param theme - The document theme (from `doc.theme`).
 * @returns Resolved hex color string (6 chars, no #), or undefined if unresolvable.
 */
export function resolveThemeColor(
  color: HexColor | ColorSpec | undefined,
  theme?: DocumentTheme
): HexColor | undefined {
  if (color === undefined) {
    return undefined;
  }
  if (typeof color === "string") {
    return color;
  }
  // ColorSpec with val — use directly
  if (color.val && color.val !== "auto") {
    return color.val;
  }
  // Resolve via theme
  if (!color.themeColor || !theme) {
    return color.val;
  }
  const colors = theme.colorScheme.colors as Record<string, string>;
  const tint = color.themeTint ? parseInt(color.themeTint, 16) / 255 : undefined;
  const shade = color.themeShade ? parseInt(color.themeShade, 16) / 255 : undefined;
  return resolveOoxmlThemeColor(color.themeColor, colors, tint, shade) ?? color.val;
}
