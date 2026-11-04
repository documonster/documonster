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
// Colour transforms — the one implementation
// =============================================================================

/**
 * A DrawingML colour transform (ECMA-376 §20.1.2.3): the child elements of a colour
 * (`<a:lumMod val="15000"/>`, `<a:alpha val="40000"/>`, `<a:comp/>`, …).
 *
 * `val` is in the element's own unit: thousandths of a percent (`100000` is 100 %) for
 * every percentage, and 60 000ths of a degree for `hue` / `hueOff` (`hueMod` is a
 * percentage). The flag transforms (`comp`, `inv`, `gray`, `gamma`, `invGamma`) take none.
 */
export interface ColorTransform {
  readonly kind: ColorTransformKind;
  readonly val?: number;
}

/** The transform elements of a DrawingML colour, as they are named in the markup. */
export type ColorTransformKind =
  | "red"
  | "redMod"
  | "redOff"
  | "green"
  | "greenMod"
  | "greenOff"
  | "blue"
  | "blueMod"
  | "blueOff"
  | "alpha"
  | "alphaMod"
  | "alphaOff"
  | "hue"
  | "hueMod"
  | "hueOff"
  | "sat"
  | "satMod"
  | "satOff"
  | "lum"
  | "lumMod"
  | "lumOff"
  | "shade"
  | "tint"
  | "gray"
  | "comp"
  | "inv"
  | "gamma"
  | "invGamma";

/** A colour with its opacity, each in 0..1. */
export interface Rgba01Color {
  readonly r: number;
  readonly g: number;
  readonly b: number;
  readonly a: number;
}

/**
 * The gamma Office uses between sRGB and the linear ("scRGB") space DrawingML's `tint`,
 * `shade`, `inv`, `gamma` and the per-channel transforms work in — 2.3, as LibreOffice's
 * MSO-compatible `oox::drawingml::Color` has it.
 */
const DECODING_GAMMA = 2.3;
const PERCENT = 100000;
const FULL_TURN = 21600000;

/**
 * Apply DrawingML colour transforms to a colour, in document order — the order matters:
 * `tint` then `lumMod` is a different colour from `lumMod` then `tint`.
 *
 * Semantics are the ones Office implements (and ECMA-376 §20.1.2.3 describes):
 *
 * - `hue*`, `sat*`, `lum*` act on the HSL form; `lum*` reaching black or white drops the
 *   saturation, so a later luminance change stays grey;
 * - `tint` (toward white) and `shade` (toward black), `inv`, `gamma` / `invGamma` and the
 *   `red*` / `green*` / `blue*` channel transforms act on linear RGB;
 * - `comp` turns the hue half a turn; `gray` is the 22 % / 72 % / 6 % weighted grey;
 * - `alpha` sets the opacity, `alphaMod` scales it and `alphaOff` shifts it.
 *
 * Every result is clamped to its range, as the format requires.
 *
 * This is the single colour-transform implementation: chart previews, Excel cell tints
 * ({@link spreadsheetTintTransforms}) and Word theme tints ({@link wordThemeTintTransforms})
 * all go through it.
 */
export function applyColorTransforms(
  color: Rgb01,
  transforms: readonly ColorTransform[]
): Rgba01Color {
  let alpha = color.a ?? 1;
  let space: "rgb" | "linear" | "hsl" = "rgb";
  let c1 = clamp01(color.r);
  let c2 = clamp01(color.g);
  let c3 = clamp01(color.b);

  const toRgb = (): void => {
    if (space === "linear") {
      [c1, c2, c3] = [c1, c2, c3].map(c => Math.pow(c, 1 / DECODING_GAMMA)) as [
        number,
        number,
        number
      ];
    } else if (space === "hsl") {
      [c1, c2, c3] = hslToRgb(c1, c2, c3);
    }
    space = "rgb";
  };
  const toLinear = (): void => {
    if (space !== "linear") {
      toRgb();
      [c1, c2, c3] = [c1, c2, c3].map(c => Math.pow(c, DECODING_GAMMA)) as [number, number, number];
      space = "linear";
    }
  };
  const toHsl = (): void => {
    if (space !== "hsl") {
      toRgb();
      [c1, c2, c3] = rgbToHsl(c1, c2, c3);
      space = "hsl";
    }
  };
  const fraction = (t: ColorTransform): number => (t.val ?? 0) / PERCENT;
  const set = (t: ColorTransform): number => clamp01(fraction(t));
  const mod = (value: number, t: ColorTransform): number => clamp01(value * fraction(t));
  const off = (value: number, t: ColorTransform): number => clamp01(value + fraction(t));
  const lumChanged = (): void => {
    if (c3 <= 0 || c3 >= 1) {
      c2 = 0;
    }
  };

  for (const t of transforms) {
    switch (t.kind) {
      case "red":
        toLinear();
        c1 = set(t);
        break;
      case "redMod":
        toLinear();
        c1 = mod(c1, t);
        break;
      case "redOff":
        toLinear();
        c1 = off(c1, t);
        break;
      case "green":
        toLinear();
        c2 = set(t);
        break;
      case "greenMod":
        toLinear();
        c2 = mod(c2, t);
        break;
      case "greenOff":
        toLinear();
        c2 = off(c2, t);
        break;
      case "blue":
        toLinear();
        c3 = set(t);
        break;
      case "blueMod":
        toLinear();
        c3 = mod(c3, t);
        break;
      case "blueOff":
        toLinear();
        c3 = off(c3, t);
        break;
      case "alpha":
        alpha = set(t);
        break;
      case "alphaMod":
        alpha = mod(alpha, t);
        break;
      case "alphaOff":
        alpha = off(alpha, t);
        break;
      case "hue":
        toHsl();
        c1 = clampTurn((t.val ?? 0) / FULL_TURN);
        break;
      case "hueMod":
        toHsl();
        c1 = clampTurn(c1 * fraction(t));
        break;
      case "hueOff":
        toHsl();
        c1 = clampTurn(c1 + (t.val ?? 0) / FULL_TURN);
        break;
      case "sat":
        toHsl();
        c2 = set(t);
        break;
      case "satMod":
        toHsl();
        c2 = mod(c2, t);
        break;
      case "satOff":
        toHsl();
        c2 = off(c2, t);
        break;
      case "lum":
        toHsl();
        c3 = set(t);
        lumChanged();
        break;
      case "lumMod":
        toHsl();
        c3 = mod(c3, t);
        lumChanged();
        break;
      case "lumOff":
        toHsl();
        c3 = off(c3, t);
        lumChanged();
        break;
      case "shade": {
        toLinear();
        const f = set(t);
        c1 *= f;
        c2 *= f;
        c3 *= f;
        break;
      }
      case "tint": {
        toLinear();
        const f = set(t);
        c1 = 1 - (1 - c1) * f;
        c2 = 1 - (1 - c2) * f;
        c3 = 1 - (1 - c3) * f;
        break;
      }
      case "gray":
        toRgb();
        c1 = c2 = c3 = c1 * 0.22 + c2 * 0.72 + c3 * 0.06;
        break;
      case "comp":
        toHsl();
        c1 = (c1 + 0.5) % 1;
        break;
      case "inv":
        toLinear();
        c1 = 1 - c1;
        c2 = 1 - c2;
        c3 = 1 - c3;
        break;
      case "gamma":
        toLinear();
        [c1, c2, c3] = [c1, c2, c3].map(c => Math.pow(c, 1 / DECODING_GAMMA)) as [
          number,
          number,
          number
        ];
        break;
      case "invGamma":
        toLinear();
        [c1, c2, c3] = [c1, c2, c3].map(c => Math.pow(c, DECODING_GAMMA)) as [
          number,
          number,
          number
        ];
        break;
    }
  }
  toRgb();
  return { r: clamp01(c1), g: clamp01(c2), b: clamp01(c3), a: clamp01(alpha) };
}

/**
 * SpreadsheetML's cell-colour `tint` (ECMA-376 §18.8.19, `-1..1`) as DrawingML transforms:
 * a positive tint lightens the HSL luminance toward white (`L·(1−t) + t`), a negative one
 * darkens it toward black (`L·(1+t)`). That is Excel's own rule — "Lighter 40 %" is
 * `tint 0.4` — and LibreOffice's MSO import of it.
 */
export function spreadsheetTintTransforms(tint: number): ColorTransform[] {
  if (!Number.isFinite(tint) || tint === 0) {
    return [];
  }
  const amount = Math.round(Math.min(1, Math.abs(tint)) * PERCENT);
  return tint > 0
    ? [
        { kind: "lumMod", val: PERCENT - amount },
        { kind: "lumOff", val: amount }
      ]
    : [{ kind: "lumMod", val: PERCENT - amount }];
}

/**
 * WordprocessingML's `w:themeTint` (one hex byte, ECMA-376 §17.3.2.41): the share of the
 * theme colour kept, the rest white, on the HSL luminance — Word's "Lighter 40 %" is
 * `themeTint="99"` (153 / 255 = 60 %), the same colour as DrawingML's
 * `lumMod 60 %, lumOff 40 %`.
 */
export function wordThemeTintTransforms(byte: number): ColorTransform[] {
  return spreadsheetTintTransforms(1 - byte / 255);
}

/**
 * WordprocessingML's `w:themeShade` (one hex byte, ECMA-376 §17.3.2.40): the share of the
 * theme colour's luminance kept — Word's "Darker 25 %" is `themeShade="BF"`.
 */
export function wordThemeShadeTransforms(byte: number): ColorTransform[] {
  return spreadsheetTintTransforms(byte / 255 - 1);
}

/** Six uppercase hex digits for a colour (its opacity is not part of them). */
export function rgb01ToHex(color: Rgb01): string {
  const channel = (c: number): string => toHex2(Math.round(clamp01(c) * 255));
  return channel(color.r) + channel(color.g) + channel(color.b);
}

/**
 * Lighten a hex colour toward white by `tint` (0 = unchanged, 1 = white), the way Office
 * does it for a cell or a theme colour: on the HSL luminance (see
 * {@link spreadsheetTintTransforms}).
 *
 * @param hex - 6-character hex color string (no "#" prefix).
 * @param tint - Tint value in range [0, 1].
 * @returns Tinted 6-character hex color string.
 */
export function applyTint(hex: string, tint: number): string {
  return transformHex(hex, spreadsheetTintTransforms(Math.max(0, tint)));
}

/**
 * Darken a hex colour toward black, keeping `shade` of its luminance (1 = unchanged,
 * 0 = black), as Office does (see {@link spreadsheetTintTransforms}).
 *
 * @param hex - 6-character hex color string (no "#" prefix).
 * @param shade - Shade value in range [0, 1].
 * @returns Shaded 6-character hex color string.
 */
export function applyShade(hex: string, shade: number): string {
  return transformHex(hex, spreadsheetTintTransforms(Math.min(1, Math.max(0, shade)) - 1));
}

/**
 * Apply a SpreadsheetML tint (ECMA-376 §18.8.19, `-1..1`) to a 0–1 float RGB colour:
 * positive lightens and negative darkens the HSL luminance, as Excel does. `tint = 0`
 * returns the colour unchanged; alpha is preserved.
 */
export function applyTintRgb01(color: Rgb01, tint: number): Rgb01 {
  const result = applyColorTransforms(color, spreadsheetTintTransforms(tint));
  const out: Rgb01 = { r: result.r, g: result.g, b: result.b };
  if (color.a !== undefined) {
    out.a = color.a;
  }
  return out;
}

function transformHex(hex: string, transforms: readonly ColorTransform[]): string {
  const base = hexToRgb01(hex);
  if (!base || transforms.length === 0) {
    return hex.slice(0, 6).toUpperCase();
  }
  return rgb01ToHex(applyColorTransforms(base, transforms));
}

function clamp01(value: number): number {
  return Number.isNaN(value) ? 0 : value < 0 ? 0 : value > 1 ? 1 : value;
}

function clampTurn(value: number): number {
  const wrapped = value % 1;
  return wrapped < 0 ? wrapped + 1 : wrapped;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  if (d === 0) {
    return [0, 0, l];
  }
  const s = l <= 0.5 ? d / (max + min) : d / (2 - max - min);
  const h = max === r ? ((g - b) / d + 6) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0 || l >= 1) {
    return [l, l, l];
  }
  // The base colour of the hue, its saturation toward grey, then the luminance toward black
  // or white — the construction Office's own conversion uses.
  const sector = h * 6;
  let r = 0;
  let g = 0;
  let b = 0;
  if (sector <= 1) {
    r = 1;
    g = sector;
  } else if (sector <= 2) {
    r = 2 - sector;
    g = 1;
  } else if (sector <= 3) {
    g = 1;
    b = sector - 2;
  } else if (sector <= 4) {
    g = 4 - sector;
    b = 1;
  } else if (sector <= 5) {
    r = sector - 4;
    b = 1;
  } else {
    r = 1;
    b = 6 - sector;
  }
  [r, g, b] = [r, g, b].map(c => (c - 0.5) * s + 0.5) as [number, number, number];
  const lum = 2 * l - 1;
  if (lum < 0) {
    const f = lum + 1;
    return [r * f, g * f, b * f];
  }
  const f = 1 - lum;
  return [1 - (1 - r) * f, 1 - (1 - g) * f, 1 - (1 - b) * f];
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
