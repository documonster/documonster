/**
 * TrueType `glyf` outlines → contours, and a font handle the rasteriser can use.
 *
 * This is the half of TrueType handling that is specific to *drawing*: locating
 * and decoding a glyph's points. The table parsing beneath it — table directory,
 * `cmap`, `hmtx`, `loca`, the `.ttc` header — is shared with the PDF embedder and
 * lives at Layer 0 in `@utils/font-ttf`.
 *
 * It used to be one file with its own copy of all of that: ~330 lines that
 * re-read the same tables `@pdf/font/ttf-parser` already read, differently. The
 * duplication was not the real cost — see `@utils/font-ttf`'s header for what it
 * actually cost, which was that system-font discovery could only be built once
 * and was built on the other copy.
 *
 * ## Why the decoding stayed here
 *
 * PDF embeds `glyf` bytes verbatim and never needs a point. Only a rasteriser
 * does, so lowering this into Layer 0 would put code there that exactly one
 * consumer can use. `@utils/font-ttf` hands over `glyphOffsets` and the table
 * directory; turning those bytes into contours is this file's job.
 *
 * ## Degrading rather than throwing
 *
 * `parseTtf` rejects a font it cannot use — a missing `glyf`, a CFF face, a
 * truncated header. Callers here draw inside a loop and have nowhere to put an
 * exception, so a rejected font becomes {@link EMPTY_RASTER_FONT}: every
 * `getOutline` answers `undefined` and the text is simply not drawn in this face.
 * `canvas.ts` treats that as "try the next font in the chain", which is what
 * makes a fallback chain possible at all.
 *
 * @module
 */

import { COMP_USE_MY_METRICS, countTtfFaces, glyphComponents, parseTtf } from "@utils/font-ttf";
import type { TtfFont } from "@utils/font-ttf";

// =============================================================================
// Types
// =============================================================================

/** One point of a glyph contour. Off-curve points are quadratic controls. */
export interface GlyphPoint {
  x: number;
  y: number;
  onCurve: boolean;
}

/** A glyph's outline in font units, with the advance the pen should take. */
export interface GlyphOutline {
  contours: GlyphPoint[][];
  advanceWidth: number;
}

/**
 * A parsed font, reduced to what rasterising needs.
 *
 * Deliberately smaller than {@link TtfFont}: a surface that draws glyphs has no
 * use for a PDF font descriptor, and `cmap` is exposed through `hasGlyph` so a
 * fallback chain can ask about coverage without holding the map.
 */
export interface RasterFont {
  unitsPerEm: number;
  ascent: number;
  descent: number;
  getOutline(codePoint: number): GlyphOutline | undefined;
  /**
   * The family this face reports, for diagnostics and cache identity.
   *
   * Optional because this interface is published and a consumer may already
   * implement it: `getOutline` is the only thing a font must provide, and adding a
   * required member would break every existing implementation at compile time for
   * no behavioural gain.
   */
  familyName?: string;
  /**
   * Whether this face can draw the code point at all.
   *
   * Optional for the same reason as {@link familyName}. Use {@link fontHasGlyph}
   * rather than calling this directly — it falls back to `getOutline`, which is the
   * only answer available for a font that does not implement it.
   */
  hasGlyph?(codePoint: number): boolean;
}

/**
 * Whether `font` can draw `codePoint`, for a font that may not implement `hasGlyph`.
 *
 * The fallback is exact rather than approximate: {@link buildRasterFont}'s
 * `getOutline` already refuses a `.notdef` mapping, so "an outline comes back" and
 * "the face really has this glyph" are the same statement. Asking `hasGlyph` first
 * only avoids decoding the outline to find out.
 */
export function fontHasGlyph(font: RasterFont, codePoint: number): boolean {
  return font.hasGlyph !== undefined
    ? font.hasGlyph(codePoint)
    : font.getOutline(codePoint) !== undefined;
}

/**
 * The font that can draw nothing.
 *
 * Returned instead of throwing when a font cannot be parsed. The metrics are
 * plausible defaults so that a caller which reads `unitsPerEm` before finding out
 * there are no glyphs does not divide by zero.
 */
export const EMPTY_RASTER_FONT: RasterFont = {
  unitsPerEm: 1000,
  ascent: 800,
  descent: -200,
  familyName: "",
  hasGlyph: () => false,
  getOutline: () => undefined
};

// =============================================================================
// Byte readers
// =============================================================================

// `glyf` decoding walks points with a bare cursor rather than the bounded
// `BEReader` in `@utils/font-ttf`: it is the hot path (a CJK label is dozens of
// glyphs, each with hundreds of points) and the byte range it walks is already
// clamped into `glyf` by `readLoca`. An offset past the end reads `undefined`,
// which propagates as `NaN` and drops the contour — it cannot escape the buffer.

function u16(data: Uint8Array, offset: number): number {
  return (data[offset] << 8) | data[offset + 1];
}

function i16(data: Uint8Array, offset: number): number {
  const v = (data[offset] << 8) | data[offset + 1];
  return v >= 0x8000 ? v - 0x10000 : v;
}

// =============================================================================
// Glyph Outline Parsing
// =============================================================================

// TrueType simple glyph flags
const ON_CURVE = 0x01;
const X_SHORT = 0x02;
const Y_SHORT = 0x04;
const REPEAT_FLAG = 0x08;
const X_SAME_OR_POS = 0x10;
const Y_SAME_OR_POS = 0x20;

function parseSimpleGlyph(data: Uint8Array, offset: number, numContours: number): GlyphPoint[][] {
  let pos = offset + 10; // skip header (numberOfContours, xMin, yMin, xMax, yMax)
  const endPts: number[] = [];
  for (let i = 0; i < numContours; i++) {
    endPts.push(u16(data, pos));
    pos += 2;
  }
  const numPoints = endPts[endPts.length - 1] + 1;

  // Skip instructions
  const instructionLength = u16(data, pos);
  pos += 2 + instructionLength;

  // Read flags
  const flags: number[] = [];
  while (flags.length < numPoints) {
    const f = data[pos++];
    flags.push(f);
    if (f & REPEAT_FLAG) {
      const repeat = data[pos++];
      for (let r = 0; r < repeat; r++) {
        flags.push(f);
      }
    }
  }

  // Read X coordinates
  const xs: number[] = new Array(numPoints);
  let x = 0;
  for (let i = 0; i < numPoints; i++) {
    const f = flags[i];
    if (f & X_SHORT) {
      const dx = data[pos++];
      x += f & X_SAME_OR_POS ? dx : -dx;
    } else if (!(f & X_SAME_OR_POS)) {
      x += i16(data, pos);
      pos += 2;
    }
    xs[i] = x;
  }

  // Read Y coordinates
  const ys: number[] = new Array(numPoints);
  let y = 0;
  for (let i = 0; i < numPoints; i++) {
    const f = flags[i];
    if (f & Y_SHORT) {
      const dy = data[pos++];
      y += f & Y_SAME_OR_POS ? dy : -dy;
    } else if (!(f & Y_SAME_OR_POS)) {
      y += i16(data, pos);
      pos += 2;
    }
    ys[i] = y;
  }

  // Build contours
  const contours: GlyphPoint[][] = [];
  let start = 0;
  for (let c = 0; c < numContours; c++) {
    const end = endPts[c];
    const contour: GlyphPoint[] = [];
    for (let i = start; i <= end; i++) {
      contour.push({ x: xs[i], y: ys[i], onCurve: !!(flags[i] & ON_CURVE) });
    }
    contours.push(contour);
    start = end + 1;
  }
  return contours;
}

/**
 * How deep a composite may nest before it is treated as empty. The spec says to
 * avoid nesting at all and real fonts stay at one or two levels; the limit is
 * what stops a font whose components reference each other in a cycle from
 * recursing forever.
 */
const MAX_COMPOSITE_DEPTH = 5;

/**
 * Find the glyph whose horizontal metrics a composite adopts.
 *
 * `USE_MY_METRICS` forces the composite's advance width and left side bearing to
 * be those of the flagged component — an i-circumflex takes the metrics of the
 * dotless i it is built from. FreeType implements this by keeping that
 * component's phantom points instead of restoring the parent's, and because the
 * points it keeps are saved per component, the last flagged component is the one
 * that survives. Returns the glyph itself when nothing claims the metrics.
 */
function metricsGlyphId(
  data: Uint8Array,
  glyphId: number,
  glyfBase: number,
  glyphOffsets: Uint32Array,
  depth = 0
): number {
  if (depth > MAX_COMPOSITE_DEPTH) {
    return glyphId;
  }
  let claimed = glyphId;
  for (const component of glyphComponents(
    data,
    glyfBase + glyphOffsets[glyphId],
    glyfBase + glyphOffsets[glyphId + 1]
  )) {
    if (component.flags & COMP_USE_MY_METRICS) {
      claimed = metricsGlyphId(data, component.glyphId, glyfBase, glyphOffsets, depth + 1);
    }
  }
  return claimed;
}

function parseCompositeGlyph(
  data: Uint8Array,
  start: number,
  end: number,
  glyfBase: number,
  glyphOffsets: Uint32Array,
  depth: number
): GlyphPoint[][] {
  const allContours: GlyphPoint[][] = [];
  for (const component of glyphComponents(data, start, end)) {
    const [a, b, c, d] = component.transform;
    const compContours = getGlyphContours(
      data,
      component.glyphId,
      glyfBase,
      glyphOffsets,
      depth + 1
    );
    for (const contour of compContours) {
      allContours.push(
        contour.map(pt => ({
          x: a * pt.x + c * pt.y + component.dx,
          y: b * pt.x + d * pt.y + component.dy,
          onCurve: pt.onCurve
        }))
      );
    }
  }
  return allContours;
}

function getGlyphContours(
  data: Uint8Array,
  glyphId: number,
  glyfBase: number,
  glyphOffsets: Uint32Array,
  depth = 0
): GlyphPoint[][] {
  const start = glyphOffsets[glyphId];
  const end = glyphOffsets[glyphId + 1];
  if (end - start < 10 || depth > MAX_COMPOSITE_DEPTH) {
    return []; // empty glyph (e.g. space), or a composite that nests too deep
  }

  const offset = glyfBase + start;
  const numberOfContours = i16(data, offset);

  if (numberOfContours >= 0) {
    return parseSimpleGlyph(data, offset, numberOfContours);
  }
  return parseCompositeGlyph(data, offset, glyfBase + end, glyfBase, glyphOffsets, depth);
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Parse a font file into a {@link RasterFont} that can render glyphs to pixels.
 *
 * `collectionIndex` selects a face inside a `.ttc`. It is not a detail that can
 * be defaulted away: macOS keeps `Songti SC` at Black in face 0 of `Songti.ttc`,
 * with Bold, Light and Regular after it, so always reading face 0 rasterises a
 * heavier weight than the one that was chosen — and the picture disagrees with
 * the PDF that embeds the same family.
 *
 * An index past the end of the collection falls back to face 0, and a font that
 * cannot be parsed becomes {@link EMPTY_RASTER_FONT}, because this module
 * degrades instead of failing.
 */
export function parseRasterFont(data: Uint8Array, collectionIndex = 0): RasterFont {
  try {
    return buildRasterFont(parseTtf(data, resolveFaceIndex(data, collectionIndex)));
  } catch {
    return EMPTY_RASTER_FONT;
  }
}

/** Clamp a requested face index to one the file actually has. */
function resolveFaceIndex(data: Uint8Array, requested: number): number {
  if (!Number.isInteger(requested) || requested < 0) {
    return 0;
  }
  try {
    return requested < countTtfFaces(data) ? requested : 0;
  } catch {
    return 0;
  }
}

/** Wrap already-parsed tables as a font that can produce outlines. */
export function buildRasterFont(ttf: TtfFont): RasterFont {
  // `parseTtf` rejects a font without `glyf`, so this is present by construction.
  const glyfBase = ttf.tables.get("glyf")!.offset;
  const { data, cmap, advanceWidths, leftSideBearings, glyphOffsets } = ttf;

  function buildOutline(codePoint: number): GlyphOutline | undefined {
    const gid = cmap.get(codePoint);
    if (gid === undefined || gid === 0) {
      return undefined;
    }
    // The glyph a composite borrows its metrics from, which is the glyph
    // itself unless a component claims them with USE_MY_METRICS.
    const metricsGid = metricsGlyphId(data, gid, glyfBase, glyphOffsets);
    const advanceWidth = advanceWidths[metricsGid];

    const contours = getGlyphContours(data, gid, glyfBase, glyphOffsets);
    if (contours.length === 0) {
      // No ink, but a real advance: a space has to move the pen by what the
      // font says, not by a guess.
      return { contours, advanceWidth };
    }
    // Outline coordinates start at the glyph's own `xMin`, but the ink belongs
    // at `pen + lsb`: a rasterizer translates the outline by `lsb - xMin`
    // (FreeType does it unconditionally, so every mainstream renderer agrees).
    // The two values match in a well-formed font and the shift is nothing, yet
    // ~0.4% of glyphs in shipped fonts disagree — by up to 0.35 em in Times
    // New Roman Italic — and those have to land where a PDF viewer would put
    // them, or a chart label drifts away from the text beside it.
    const shift = leftSideBearings[metricsGid] - i16(data, glyfBase + glyphOffsets[metricsGid] + 2);
    return {
      contours:
        shift === 0
          ? contours
          : contours.map(contour => contour.map(pt => ({ ...pt, x: pt.x + shift }))),
      advanceWidth
    };
  }

  // One outline object per code point, so callers that rasterize the same glyph
  // twice — measuring a label and then drawing it — get the same object and can
  // key a cache on it.
  const outlines = new Map<number, GlyphOutline | undefined>();

  return {
    unitsPerEm: ttf.unitsPerEm,
    ascent: ttf.ascent,
    descent: ttf.descent,
    familyName: ttf.familyName,
    // A gid of 0 is `.notdef`, which draws a box rather than the character asked
    // for — so it is not coverage, and a fallback chain must keep looking.
    hasGlyph(codePoint: number): boolean {
      const gid = cmap.get(codePoint);
      return gid !== undefined && gid !== 0;
    },
    getOutline(codePoint: number): GlyphOutline | undefined {
      if (!outlines.has(codePoint)) {
        outlines.set(codePoint, buildOutline(codePoint));
      }
      return outlines.get(codePoint);
    }
  };
}

// =============================================================================
// Supplied fonts
// =============================================================================

/**
 * A font a caller can hand over.
 *
 * Bytes are the common case; a {@link RasterFont} lets a caller parse once and reuse
 * the result across renders, and the object form is the only way to name a face inside
 * a `.ttc` — which matters, because `Songti.ttc` holds `Songti SC` at Black in face 0
 * with Regular further along, so the default is rarely the one wanted.
 */
export type RasterFontSource =
  | Uint8Array
  | RasterFont
  | { readonly data: Uint8Array; readonly collectionIndex?: number };

/**
 * Parse supplied fonts, dropping any that cannot be used.
 *
 * A face that cannot be parsed is skipped rather than throwing: the rasteriser degrades
 * instead of failing, and a caller who passes a CFF font should still get their Latin
 * text drawn.
 *
 * Lives here, beside {@link parseRasterFont}, rather than in the platform-specific
 * font modules — it touches no filesystem, and having a copy in each of them meant the
 * Node and browser builds could disagree about what a caller was allowed to pass.
 */
export function parseInjectedFonts(fonts: readonly RasterFontSource[]): RasterFont[] {
  return (
    fonts
      .map(toRasterFont)
      // Identity, not `familyName`: that member is optional, so filtering on it threw
      // away a caller's own font object for reporting no name. `parseRasterFont`
      // returns this exact singleton for anything it could not use.
      .filter(font => font !== EMPTY_RASTER_FONT)
  );
}

/** Normalise one supplied font to a {@link RasterFont}. */
function toRasterFont(source: RasterFontSource): RasterFont {
  if (isByteSource(source)) {
    return parseRasterFont(normalizeBytes(source));
  }
  if ("data" in source) {
    return isByteSource(source.data)
      ? parseRasterFont(normalizeBytes(source.data), source.collectionIndex ?? 0)
      : EMPTY_RASTER_FONT;
  }
  return source;
}

/**
 * Whether `value` is a byte array, including one from another realm.
 *
 * A type predicate rather than a conversion, so the caller's union narrows. The
 * structural test is the point: a bare `instanceof Uint8Array` is false for a typed
 * array created in another realm — an iframe, a worker — which would then be taken for
 * a `RasterFont`, fail the filter above, and be dropped **silently**. Supplying bytes is
 * the only route to a non-ASCII glyph in a browser, which is exactly where cross-realm
 * values come from.
 */
function isByteSource(value: unknown): value is Uint8Array {
  if (value instanceof Uint8Array) {
    return true;
  }
  return (
    ArrayBuffer.isView(value) && (value as { BYTES_PER_ELEMENT?: number }).BYTES_PER_ELEMENT === 1
  );
}

/** Re-wrap a possibly cross-realm view as a `Uint8Array` this realm can read. */
function normalizeBytes(value: Uint8Array): Uint8Array {
  return value instanceof Uint8Array
    ? value
    : new Uint8Array(
        (value as ArrayBufferView).buffer,
        (value as ArrayBufferView).byteOffset,
        (value as ArrayBufferView).byteLength
      );
}
