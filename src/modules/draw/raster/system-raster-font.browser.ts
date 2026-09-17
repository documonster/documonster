/**
 * Discover system fonts for the rasteriser — browser stub.
 *
 * There is no font directory to walk in a browser, so a chain is exactly what the
 * caller supplied. The point of the split is not the runtime behaviour — the Node
 * module's own `typeof process` checks would return the same nothing — but the
 * **bytes**: that module reaches the per-platform directory tables in
 * `@utils/font-discovery`, tens of kilobytes of `/System/Library/Fonts/Supplemental`,
 * `msyh.ttc` and `PingFang SC` that no browser can act on. A runtime guard cannot
 * remove them, because a bundler has to keep any string the module might reach; only a
 * separate module can.
 *
 * Supplying fonts is **not** stubbed, and it is what makes a browser usable: a caller
 * who fetches font bytes and passes them through `RasterizeOptions.fonts` gets real
 * glyphs. That path is platform-independent, so it lives in
 * `@draw/raster/glyph-outline` and is shared with Node rather than written twice here.
 * Without it the only fallback is the built-in stroke font, which covers ASCII 32–126
 * and draws every other character as `?`.
 *
 * @module
 */

import type { RasterFont } from "@draw/raster/glyph-outline";

/**
 * The faces to try when drawing — exactly what the caller supplied.
 *
 * `text` and `useSystemFonts` are accepted and ignored: there is nothing to search and
 * therefore nothing to opt out of.
 */
export function resolveFontChain(
  _text: string,
  injected: readonly RasterFont[],
  _useSystemFonts: boolean
): readonly RasterFont[] {
  return injected;
}

/** Nothing is discovered in a browser, so there is nothing to forget. */
export function resetDiscoveredFonts(): void {
  // No cache to clear.
}
