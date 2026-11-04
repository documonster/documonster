/**
 * `Theme` namespace surface — theme color resolution.
 *
 * `import { Theme } from "documonster/word"` →
 *   `Theme.resolveColor(...)` — tree-shaken via `export * as Theme`.
 */
import {
  checkHexColor,
  checkHexByte,
  checkRecord,
  checkString,
  invalidArgument
} from "@word/core/argument-checks";
import { resolveThemeColor } from "@word/core/color-utils";
import type { ColorSpec, DocumentTheme, HexColor } from "@word/types";

/**
 * Resolve a colour to six hex digits (`RRGGBB`), or `undefined` when there is
 * none to give: `"auto"` (the application's default colour), or a theme colour
 * without a theme that defines it and without a `val`.
 *
 * `color` is a `HexColor` — `RRGGBB` (no `#`) or `"auto"` — or a `ColorSpec`.
 * A `ColorSpec`'s `themeColor`, where `theme` defines it, wins over its `val`
 * as it does in Word (ECMA-376 §17.3.2.6), with `themeTint` / `themeShade`
 * applied. A theme colour name is a scheme slot (`accent1`, `dk1`, `hlink`, …)
 * or a Word name (`dark1`, `hyperlink`, `text1`, `background1`, …; `text*` /
 * `background*` resolve through Word's default colour mapping).
 *
 * @param color - The colour: a `HexColor` string or a `ColorSpec`.
 * @param theme - The document theme (`doc.theme`).
 * @returns Six hex digits, or `undefined` if there is no colour.
 * @throws {DocxArgumentError} If `color` is in neither form — `"#FF0000"`,
 * `"red"` — or a `ColorSpec` field is malformed, or `theme` is not an object.
 */
export function resolveColor(
  color: HexColor | ColorSpec | undefined,
  theme?: DocumentTheme
): HexColor | undefined {
  const where = "Theme.resolveColor";
  if (typeof color === "string") {
    checkHexColor(color, "color", where);
  } else if (color !== undefined) {
    if (typeof color !== "object" || color === null) {
      invalidArgument(where, "color", 'a hex color "RRGGBB", "auto" or a ColorSpec', color);
    }
    checkHexColor(color.val, "color.val", where, { optional: true });
    checkString(color.themeColor, "color.themeColor", where, { optional: true });
    checkHexByte(color.themeTint, "color.themeTint", where);
    checkHexByte(color.themeShade, "color.themeShade", where);
  }
  checkRecord(theme, "theme", where, { optional: true });
  return resolveThemeColor(color, theme);
}
