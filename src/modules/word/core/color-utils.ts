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
  applyColorTransforms,
  applyTint,
  applyShade,
  hexToRgb01,
  resolveOoxmlThemeColor,
  rgb01ToHex,
  THEME_COLOR_ATTRIBUTE_MAP,
  wordThemeShadeTransforms,
  wordThemeTintTransforms
} from "@utils/theme-colors";
import type { OoxmlThemeColorName } from "@utils/theme-colors";
import type { DocumentTheme, ColorSpec, HexColor } from "@word/types";

// Re-export shared utilities so existing consumers that import from here
// continue to work without changes.
export { applyTint, applyShade, THEME_COLOR_ATTRIBUTE_MAP };
export type { OoxmlThemeColorName };

/**
 * Word's theme colour names (`ST_ThemeColor`) that are not scheme slots, by
 * the slot each names under Word's default colour mapping
 * (`w:clrSchemeMapping`): text on dark 1, background on light 1. A document
 * that remaps them in its settings is not followed.
 */
const WORD_THEME_COLOR_SLOTS: Readonly<Record<string, string>> = {
  text1: "dk1",
  background1: "lt1",
  text2: "dk2",
  background2: "lt2"
};

const HEX_BYTE = /^[0-9A-Fa-f]{2}$/;

/** A `themeTint` / `themeShade` byte, or `undefined` when absent or not one hex byte. */
function themeByte(value: string | undefined): number | undefined {
  return value !== undefined && HEX_BYTE.test(value) ? parseInt(value, 16) : undefined;
}

/**
 * Resolve a colour to six hex digits, using the document theme for a theme
 * colour reference.
 *
 * - A `HexColor` string `RRGGBB` is returned as it is; `"auto"` — the
 *   application's default colour, which has no RGB of its own — gives
 *   `undefined`.
 * - A `ColorSpec` with a `themeColor` the `theme` defines gives that theme
 *   colour, with `themeTint` / `themeShade` applied to its HSL luminance as Word
 *   does (`@utils/theme-colors`, the one colour-transform implementation): Word ignores `val` then
 *   (ECMA-376 §17.3.2.6). The name is a scheme slot (`accent1`, `dk1`,
 *   `hlink`, …) or a Word name (`dark1`, `hyperlink`, `text1`,
 *   `background1`, …; `text*` / `background*` through Word's default colour
 *   mapping).
 * - Otherwise — no `themeColor`, no theme, or a name the theme does not
 *   define — it gives its `val`, the colour Word last computed, and
 *   `undefined` for `"auto"`.
 *
 * `Theme.resolveColor` refuses a value in none of these forms with
 * `DocxArgumentError`; this internal form, which renderers call on read
 * documents, treats one as unresolvable instead.
 *
 * @param color - The color value (HexColor string or ColorSpec).
 * @param theme - The document theme (from `doc.theme`).
 * @returns Resolved hex color string (6 chars, no #), or undefined if unresolvable.
 */
export function resolveThemeColor(
  color: HexColor | ColorSpec | undefined,
  theme?: DocumentTheme
): HexColor | undefined {
  if (color === undefined || color === null) {
    return undefined;
  }
  if (typeof color === "string") {
    return HEX_COLOR.test(color) ? color : undefined;
  }
  const val = typeof color.val === "string" && HEX_COLOR.test(color.val) ? color.val : undefined;
  // A theme colour, where the theme defines it, wins over `val` (ECMA-376
  // §17.3.2.6: with `themeColor` present `val` is ignored); `val` is the
  // colour last computed from it, the fallback without a theme.
  if (typeof color.themeColor === "string" && theme) {
    const colors = theme.colorScheme.colors as Record<string, string>;
    const name = WORD_THEME_COLOR_SLOTS[color.themeColor] ?? color.themeColor;
    const base = resolveOoxmlThemeColor(name, colors);
    if (base !== undefined) {
      // `themeTint` keeps that share of the colour and lightens the rest toward white;
      // `themeShade` keeps that share of its luminance — both on the HSL luminance, as Word
      // does ("Lighter 40 %" is `themeTint="99"`). They were mixed in RGB, and the tint the
      // wrong way round: `"99"` came out as Word's "Lighter 60 %".
      const tint = themeByte(color.themeTint);
      const shade = themeByte(color.themeShade);
      const transforms = [
        ...(tint === undefined ? [] : wordThemeTintTransforms(tint)),
        ...(shade === undefined ? [] : wordThemeShadeTransforms(shade))
      ];
      const rgb = hexToRgb01(base);
      return transforms.length === 0 || !rgb
        ? base
        : rgb01ToHex(applyColorTransforms(rgb, transforms));
    }
  }
  return val;
}

const HEX_COLOR = /^[0-9A-Fa-f]{6}$/;
