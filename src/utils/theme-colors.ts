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

/**
 * Default Office ("Office" theme) colour scheme, keyed by OOXML scheme slot.
 *
 * Values are the modern Office 2013+ defaults — the same palette Excel ships
 * with today and which {@link module:@excel/xlsx/xml/theme1} embeds in written
 * workbooks. This is the single source of truth: every per-index palette below
 * is derived from it so chart previews, cell-colour rendering (PDF / Word) and
 * the written `theme1.xml` all agree on the same colours.
 */
export const DEFAULT_OFFICE_THEME: Readonly<Record<OoxmlThemeColorName, string>> = Object.freeze({
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
});

/**
 * DrawingML theme-palette order (`CT_ColorMapping`, ECMA-376 §20.1.2.3.29):
 * `dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink`. This is the order a chart
 * / sparkline `<a:schemeClr>` theme index maps to.
 */
export const CHART_THEME_PALETTE: readonly string[] = Object.freeze([
  DEFAULT_OFFICE_THEME.dk1,
  DEFAULT_OFFICE_THEME.lt1,
  DEFAULT_OFFICE_THEME.dk2,
  DEFAULT_OFFICE_THEME.lt2,
  DEFAULT_OFFICE_THEME.accent1,
  DEFAULT_OFFICE_THEME.accent2,
  DEFAULT_OFFICE_THEME.accent3,
  DEFAULT_OFFICE_THEME.accent4,
  DEFAULT_OFFICE_THEME.accent5,
  DEFAULT_OFFICE_THEME.accent6,
  DEFAULT_OFFICE_THEME.hlink,
  DEFAULT_OFFICE_THEME.folHlink
]);

/**
 * SpreadsheetML cell-colour theme order (`<color theme="n"/>` in styles.xml):
 * `lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink`. Note this differs from
 * {@link CHART_THEME_PALETTE}: the first two and the second two slots are
 * swapped (background/text vs the DrawingML dark/light ordering). This is the
 * order an Excel cell's `color.theme` index maps to.
 */
export const CELL_THEME_PALETTE: readonly string[] = Object.freeze([
  DEFAULT_OFFICE_THEME.lt1,
  DEFAULT_OFFICE_THEME.dk1,
  DEFAULT_OFFICE_THEME.lt2,
  DEFAULT_OFFICE_THEME.dk2,
  DEFAULT_OFFICE_THEME.accent1,
  DEFAULT_OFFICE_THEME.accent2,
  DEFAULT_OFFICE_THEME.accent3,
  DEFAULT_OFFICE_THEME.accent4,
  DEFAULT_OFFICE_THEME.accent5,
  DEFAULT_OFFICE_THEME.accent6,
  DEFAULT_OFFICE_THEME.hlink,
  DEFAULT_OFFICE_THEME.folHlink
]);

// =============================================================================
// Hex parsing
// =============================================================================

/** RGB colour with each component in the range [0, 1], plus optional alpha. */
export interface Rgb01 {
  r: number;
  g: number;
  b: number;
  a?: number;
}

/**
 * Parse a hex colour string into 0–1 float RGB components.
 *
 * Accepts `RRGGBB` (6 digits), `AARRGGBB` (8 digits, alpha first — the form
 * Excel uses for cell fills), and `RGB` (3 digits, CSS shorthand). A leading
 * `#` is tolerated. Returns `null` for anything else.
 */
export function hexToRgb01(hex: string | undefined): Rgb01 | null {
  if (!hex) {
    return null;
  }
  let h = hex.startsWith("#") ? hex.slice(1) : hex;
  if (h.length === 3) {
    // CSS shorthand: expand each nibble (e.g. "abc" → "aabbcc").
    h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  }
  let a = 255;
  let rHex: string;
  let gHex: string;
  let bHex: string;
  if (h.length === 8) {
    a = parseInt(h.slice(0, 2), 16);
    rHex = h.slice(2, 4);
    gHex = h.slice(4, 6);
    bHex = h.slice(6, 8);
  } else if (h.length === 6) {
    rHex = h.slice(0, 2);
    gHex = h.slice(2, 4);
    bHex = h.slice(4, 6);
  } else {
    return null;
  }
  const r = parseInt(rHex, 16);
  const g = parseInt(gHex, 16);
  const b = parseInt(bHex, 16);
  if (Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b) || Number.isNaN(a)) {
    return null;
  }
  const alpha = a / 255;
  return alpha < 1
    ? { r: r / 255, g: g / 255, b: b / 255, a: alpha }
    : { r: r / 255, g: g / 255, b: b / 255 };
}

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

/**
 * Apply a tint to a 0–1 float RGB colour, per OOXML §18.8.19.
 *
 * Tint ∈ [-1, 1]: positive lightens toward white
 * (`c + (1 - c) * tint`), negative darkens toward black
 * (`c * (1 + tint)`). `tint = 0` returns the colour unchanged. Result
 * components are clamped to [0, 1]; alpha is preserved.
 */
export function applyTintRgb01(color: Rgb01, tint: number): Rgb01 {
  const apply = (c: number): number => {
    const v = tint < 0 ? c * (1 + tint) : c + (1 - c) * tint;
    return Math.max(0, Math.min(1, v));
  };
  const out: Rgb01 = { r: apply(color.r), g: apply(color.g), b: apply(color.b) };
  if (color.a !== undefined) {
    out.a = color.a;
  }
  return out;
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
  // Emit uppercase hex to match the OOXML ST_HexColorRGB convention and the
  // casing of theme scheme colors (which are stored uppercase), so a tinted /
  // shaded result is consistent with an un-transformed passthrough.
  const h = Math.max(0, Math.min(255, n)).toString(16).toUpperCase();
  return h.length < 2 ? "0" + h : h;
}
