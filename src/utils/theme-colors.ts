/**
 * OOXML Theme Color Resolution — shared between Excel and Word modules.
 *
 * Provides utilities for resolving theme color references with tint/shade
 * transformations per the ECMA-376 color model.
 */

// =============================================================================
// Types
// =============================================================================

/** Standard OOXML theme color names (scheme keys in theme1.xml). */
export type OoxmlThemeColorName =
  | "dk1"
  | "lt1"
  | "dk2"
  | "lt2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hlink"
  | "folHlink";

// =============================================================================
// Constants
// =============================================================================

/**
 * Map from OOXML property attribute values to theme scheme keys.
 *
 * Word/Excel run/paragraph/cell properties use long-form names ("dark1",
 * "light1", "hyperlink") while the theme XML stores short-form keys
 * ("dk1", "lt1", "hlink"). This map normalises both forms.
 */
export const THEME_COLOR_ATTRIBUTE_MAP: Record<string, OoxmlThemeColorName> = {
  dark1: "dk1",
  light1: "lt1",
  dark2: "dk2",
  light2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
  // Also accept direct scheme keys
  dk1: "dk1",
  lt1: "lt1",
  dk2: "dk2",
  lt2: "lt2",
  hlink: "hlink",
  folHlink: "folHlink"
};

// =============================================================================
// Tint / Shade transformations
// =============================================================================

/**
 * Apply a tint (lighten toward white) to a hex color.
 *
 * OOXML tint formula: `newComponent = component + (255 - component) * tint`
 * where tint ∈ [0, 1]. tint=0 → original, tint=1 → white.
 *
 * @param hex - 6-character hex color string (no "#" prefix).
 * @param tint - Tint value in range [0, 1].
 * @returns Tinted 6-character hex color string.
 */
export function applyTint(hex: string, tint: number): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const nr = Math.round(r + (255 - r) * tint);
  const ng = Math.round(g + (255 - g) * tint);
  const nb = Math.round(b + (255 - b) * tint);
  return toHex2(nr) + toHex2(ng) + toHex2(nb);
}

/**
 * Apply a shade (darken toward black) to a hex color.
 *
 * OOXML shade formula: `newComponent = component * shade`
 * where shade ∈ [0, 1]. shade=1 → original, shade=0 → black.
 *
 * @param hex - 6-character hex color string (no "#" prefix).
 * @param shade - Shade value in range [0, 1].
 * @returns Shaded 6-character hex color string.
 */
export function applyShade(hex: string, shade: number): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const nr = Math.round(r * shade);
  const ng = Math.round(g * shade);
  const nb = Math.round(b * shade);
  return toHex2(nr) + toHex2(ng) + toHex2(nb);
}

// =============================================================================
// Theme Color Resolution
// =============================================================================

/**
 * Resolve a theme color reference to a hex color string.
 *
 * Looks up the theme color by attribute name (normalising long-form names
 * like "dark1" to scheme keys like "dk1"), then applies optional tint or
 * shade transformation.
 *
 * @param themeColorName - The theme color attribute value (e.g. "accent1", "dark1", "dk1").
 * @param colors - The theme color scheme (scheme key → 6-char hex mapping).
 * @param tint - Optional tint value in [0, 1]. Values > 1 are treated as
 *               raw bytes (0-255) and normalised by dividing by 255.
 * @param shade - Optional shade value in [0, 1]. Values > 1 are treated as
 *                raw bytes (0-255) and normalised by dividing by 255.
 * @returns Resolved 6-character hex color string, or undefined if the theme
 *          color name cannot be found in the provided scheme.
 */
export function resolveOoxmlThemeColor(
  themeColorName: string,
  colors: Readonly<Record<string, string>>,
  tint?: number,
  shade?: number
): string | undefined {
  const key = THEME_COLOR_ATTRIBUTE_MAP[themeColorName] ?? themeColorName;
  const base = colors[key];
  if (!base) {
    return undefined;
  }
  if (tint !== undefined && tint !== 0) {
    const normalizedTint = tint > 1 ? tint / 255 : tint;
    return applyTint(base, normalizedTint);
  }
  if (shade !== undefined && shade !== 0) {
    const normalizedShade = shade > 1 ? shade / 255 : shade;
    return applyShade(base, normalizedShade);
  }
  return base;
}

// =============================================================================
// Helpers
// =============================================================================

function toHex2(n: number): string {
  const h = Math.max(0, Math.min(255, n)).toString(16);
  return h.length < 2 ? "0" + h : h;
}
