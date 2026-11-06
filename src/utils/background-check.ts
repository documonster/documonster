/**
 * The one check for a renderer's background — `Draw.toSvg` / `toRgba`, the chart renderers
 * (`ChartRenderOptions.background`), Word's page renderer (`RenderOptions.background`).
 *
 * At Layer 0 because those renderers live in three modules and the check needs nothing but the
 * colour lexer: kept in `draw/`, it pulled the drawing module's argument checks and error classes
 * into a bundle of Word's `Layout`, which draws no display list.
 */
import { isNoPaintColor, parseCssColor } from "./svg-lex";

/**
 * How a module reports a refused argument: throw its own `*ArgumentError` for `field`, with
 * the message `${where}: ${field} must be ${expected}, got …`.
 */
export type ArgumentFailure = (
  where: string,
  field: string,
  expected: string,
  value: unknown
) => never;

/** What {@link checkBackgroundWith} says a background must be. */
const BACKGROUND_EXPECTED =
  "a CSS colour (#rgb, #rrggbb, #rrggbbaa, rgb(), hsl(), a colour name, transparent or none) or { r, g, b, a }";

/**
 * A background as every renderer takes it: a CSS colour string (`#rgb`, `#rrggbb`,
 * `#rrggbbaa`, `rgb()`, `hsl()`, a colour name, `transparent` or `none`) or an IR colour
 * `{ r, g, b, a }` with each channel a number from 0 to 1.
 *
 * Each caller passes its own `failWith`, so the error is that module's class while the field
 * and the words are the same everywhere; `undefined` is the caller's to allow.
 */
export function checkBackgroundWith(
  where: string,
  field: string,
  value: unknown,
  failWith: ArgumentFailure
): void {
  if (typeof value === "string") {
    if (!isNoPaintColor(value) && parseCssColor(value) === undefined) {
      failWith(where, field, BACKGROUND_EXPECTED, value);
    }
    return;
  }
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    failWith(where, field, BACKGROUND_EXPECTED, value);
  }
  const record = value as Record<string, unknown>;
  for (const channel of ["r", "g", "b", "a"] as const) {
    const component = record[channel];
    if (typeof component !== "number" || !(component >= 0 && component <= 1)) {
      failWith(where, `${field}.${channel}`, "a number from 0 to 1", component);
    }
  }
}
