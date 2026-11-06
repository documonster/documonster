/**
 * The faces Word lays text out with, measured from their own tables.
 *
 * Word sizes a single-spaced line by the face's line spacing — `hhea` ascent,
 * descent and line gap, or the typo metrics when the face asks for them — and
 * breaks lines by the face's advance widths. Measured in Word for Mac 16:
 * Calibri sets 1.2207 em per line, Arial 1.1499, Times New Roman 1.1499,
 * Cambria 1.1724, Courier New 1.1328 — not the 1.2 em the layout otherwise
 * assumes — and a page of 11 pt Calibri body text holds a different number of
 * lines and of words per line than that assumption gives, so page numbers came
 * out a page early after a few pages.
 *
 * Every function answers `undefined` for a face it holds no table for; the
 * caller keeps its own approximation for those.
 */

import { getCharAdvance, getFontLineMetrics, getLayoutFontMetrics } from "@utils/font-data";
import { isZeroWidthCodePoint } from "@utils/font-metrics";

/** Width of `text` in points, set in `fontName` at `fontSize` points. */
export function measureFaceText(
  text: string,
  fontName: string,
  fontSize: number,
  bold?: boolean,
  italic?: boolean
): number | undefined {
  const metrics = getLayoutFontMetrics(fontName, bold, italic);
  if (!metrics) {
    return undefined;
  }
  let cache = unitCache.get(metrics);
  if (!cache) {
    cache = new Map();
    unitCache.set(metrics, cache);
  }
  let units = cache.get(text);
  if (units === undefined) {
    units = 0;
    for (const ch of text) {
      const codePoint = ch.codePointAt(0)!;
      if (!isZeroWidthCodePoint(codePoint)) {
        units += getCharAdvance(metrics, codePoint);
      }
    }
    if (cache.size >= UNIT_CACHE_LIMIT) {
      cache.clear();
    }
    cache.set(text, units);
  }
  return (units / metrics.header.unitsPerEm) * fontSize;
}

/**
 * Advance of a string in font units, by face. Layout measures the same words
 * over and over — every pass of a table-of-contents update lays the whole
 * document out again — so this is most of what measuring costs.
 */
const unitCache = new Map<object, Map<string, number>>();
const UNIT_CACHE_LIMIT = 50_000;

/** Height in points of a single-spaced line set in `fontName` at `fontSize` points. */
export function faceLineHeight(
  fontName: string,
  fontSize: number,
  bold?: boolean
): number | undefined {
  const metrics = getLayoutFontMetrics(fontName, bold);
  if (!metrics) {
    return undefined;
  }
  const { ascent, descent, lineGap } = getFontLineMetrics(metrics);
  return ((ascent + descent + lineGap) / metrics.header.unitsPerEm) * fontSize;
}

/** Ascent and descent (negative) in points of `fontName` at `fontSize` points. */
export function faceExtents(
  fontName: string,
  fontSize: number,
  bold?: boolean
): { ascent: number; descent: number } | undefined {
  const metrics = getLayoutFontMetrics(fontName, bold);
  if (!metrics) {
    return undefined;
  }
  const { ascent, descent } = getFontLineMetrics(metrics);
  const scale = fontSize / metrics.header.unitsPerEm;
  return { ascent: ascent * scale, descent: -descent * scale };
}

/**
 * The part of a single-spaced line of `fontName` above the baseline — ascent
 * and line gap — and below it, in points. Word sets a line holding several
 * faces by the largest of each: an 11pt line with a Symbol bullet on it is
 * Symbol's ascent over Calibri's descent, taller than either face's own line.
 */
export function faceLineExtents(
  fontName: string,
  fontSize: number,
  bold?: boolean
): { above: number; below: number } | undefined {
  const metrics = getLayoutFontMetrics(fontName, bold);
  if (!metrics) {
    return undefined;
  }
  const { ascent, descent, lineGap } = getFontLineMetrics(metrics);
  const scale = fontSize / metrics.header.unitsPerEm;
  return { above: (ascent + lineGap) * scale, below: descent * scale };
}
