/**
 * DOCX Module - Theme Color Utilities
 *
 * Resolves OOXML theme colors with tint/shade transformations.
 * Extracted to a standalone file so that html-renderer and document.ts
 * can both import it without creating circular heavy dependencies.
 */

import type { DocumentTheme, ColorSpec, HexColor } from "./types";

/**
 * Map OOXML theme color attribute names to theme color scheme keys.
 * Word uses different names in run/paragraph properties vs the theme XML.
 */
const THEME_COLOR_MAP: Record<string, string> = {
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
  // Direct names also work
  dk1: "dk1",
  lt1: "lt1",
  dk2: "dk2",
  lt2: "lt2",
  hlink: "hlink",
  folHlink: "folHlink"
};

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
  const key = THEME_COLOR_MAP[color.themeColor] ?? color.themeColor;
  const base = (theme.colorScheme.colors as Record<string, string>)[key];
  if (!base) {
    return color.val;
  }
  // Apply tint or shade
  if (color.themeTint) {
    return applyTint(base, parseInt(color.themeTint, 16) / 255);
  }
  if (color.themeShade) {
    return applyShade(base, parseInt(color.themeShade, 16) / 255);
  }
  return base;
}

/** Apply tint to a hex color. tint=1 → white, tint=0 → original. */
function applyTint(hex: string, tint: number): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const nr = Math.round(r + (255 - r) * tint);
  const ng = Math.round(g + (255 - g) * tint);
  const nb = Math.round(b + (255 - b) * tint);
  return toHex2(nr) + toHex2(ng) + toHex2(nb);
}

/** Apply shade to a hex color. shade=1 → original, shade=0 → black. */
function applyShade(hex: string, shade: number): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const nr = Math.round(r * shade);
  const ng = Math.round(g * shade);
  const nb = Math.round(b * shade);
  return toHex2(nr) + toHex2(ng) + toHex2(nb);
}

function toHex2(n: number): string {
  const h = Math.max(0, Math.min(255, n)).toString(16);
  return h.length < 2 ? "0" + h : h;
}
