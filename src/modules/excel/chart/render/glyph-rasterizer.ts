/**
 * Lightweight TrueType glyph rasterizer for chart PNG text rendering.
 *
 * Parses the minimal set of TTF tables (cmap, hmtx, loca, glyf, head, hhea,
 * maxp) needed to extract glyph outlines, then rasterizes them via scan-line
 * fill. No dependencies outside this module except `@utils/fs`.
 *
 * Supports:
 * - Simple glyphs (positive numberOfContours)
 * - Composite glyphs (negative numberOfContours, recursive component assembly)
 * - Quadratic B-spline outlines (on-curve / off-curve points, implicit
 *   on-curve midpoints between consecutive off-curve points)
 *
 * Does NOT support: hinting, kerning, OpenType features, vertical layout.
 * This is intentional — the goal is readable chart labels, not DTP.
 */

import { fileExistsSync, readFileBytesSync } from "@utils/fs";

// =============================================================================
// Types
// =============================================================================

interface GlyphPoint {
  x: number;
  y: number;
  onCurve: boolean;
}

interface GlyphOutline {
  contours: GlyphPoint[][];
  advanceWidth: number;
}

export interface RasterFont {
  unitsPerEm: number;
  ascent: number;
  descent: number;
  getOutline(codePoint: number): GlyphOutline | undefined;
}

// =============================================================================
// TTF binary reader helpers
// =============================================================================

function u16(data: Uint8Array, offset: number): number {
  return (data[offset] << 8) | data[offset + 1];
}

function i16(data: Uint8Array, offset: number): number {
  const v = (data[offset] << 8) | data[offset + 1];
  return v >= 0x8000 ? v - 0x10000 : v;
}

function u32(data: Uint8Array, offset: number): number {
  return (
    ((data[offset] << 24) |
      (data[offset + 1] << 16) |
      (data[offset + 2] << 8) |
      data[offset + 3]) >>>
    0
  );
}

function tag(data: Uint8Array, offset: number): string {
  return String.fromCharCode(data[offset], data[offset + 1], data[offset + 2], data[offset + 3]);
}

// =============================================================================
// TTF Parser (minimal)
// =============================================================================

interface TableEntry {
  offset: number;
  length: number;
}

function parseTableDirectory(data: Uint8Array): Map<string, TableEntry> {
  const numTables = u16(data, 4);
  const tables = new Map<string, TableEntry>();
  for (let i = 0; i < numTables; i++) {
    const rec = 12 + i * 16;
    const t = tag(data, rec);
    tables.set(t, { offset: u32(data, rec + 8), length: u32(data, rec + 12) });
  }
  return tables;
}

function parseCmap(data: Uint8Array, table: TableEntry): Map<number, number> {
  const base = table.offset;
  const numSubtables = u16(data, base + 2);
  const map = new Map<number, number>();

  for (let i = 0; i < numSubtables; i++) {
    const rec = base + 4 + i * 8;
    const platformID = u16(data, rec);
    const encodingID = u16(data, rec + 2);
    const subtableOffset = base + u32(data, rec + 4);
    const format = u16(data, subtableOffset);

    // Prefer (3,1) format 4 or (3,10) format 12 — standard Windows Unicode mappings
    if (platformID === 3 && (encodingID === 1 || encodingID === 10)) {
      if (format === 4) {
        parseCmapFormat4(data, subtableOffset, map);
      } else if (format === 12) {
        parseCmapFormat12(data, subtableOffset, map);
      }
    }
  }
  return map;
}

function parseCmapFormat4(data: Uint8Array, offset: number, map: Map<number, number>): void {
  const segCount = u16(data, offset + 6) >> 1;
  const endCodesOff = offset + 14;
  const startCodesOff = endCodesOff + segCount * 2 + 2; // +2 for reservedPad
  const idDeltasOff = startCodesOff + segCount * 2;
  const idRangeOffsetsOff = idDeltasOff + segCount * 2;

  for (let i = 0; i < segCount; i++) {
    const endCode = u16(data, endCodesOff + i * 2);
    const startCode = u16(data, startCodesOff + i * 2);
    const idDelta = i16(data, idDeltasOff + i * 2);
    const idRangeOffset = u16(data, idRangeOffsetsOff + i * 2);

    if (startCode === 0xffff) {
      break;
    }

    for (let c = startCode; c <= endCode; c++) {
      let gid: number;
      if (idRangeOffset === 0) {
        gid = (c + idDelta) & 0xffff;
      } else {
        const glyphIndexAddr = idRangeOffsetsOff + i * 2 + idRangeOffset + (c - startCode) * 2;
        gid = u16(data, glyphIndexAddr);
        if (gid !== 0) {
          gid = (gid + idDelta) & 0xffff;
        }
      }
      if (gid !== 0 && !map.has(c)) {
        map.set(c, gid);
      }
    }
  }
}

function parseCmapFormat12(data: Uint8Array, offset: number, map: Map<number, number>): void {
  const numGroups = u32(data, offset + 12);
  let pos = offset + 16;
  for (let i = 0; i < numGroups; i++) {
    const startCharCode = u32(data, pos);
    const endCharCode = u32(data, pos + 4);
    let startGlyphID = u32(data, pos + 8);
    for (let c = startCharCode; c <= endCharCode; c++) {
      if (!map.has(c)) {
        map.set(c, startGlyphID);
      }
      startGlyphID++;
    }
    pos += 12;
  }
}

function parseLoca(
  data: Uint8Array,
  table: TableEntry,
  numGlyphs: number,
  isLong: boolean
): Uint32Array {
  const offsets = new Uint32Array(numGlyphs + 1);
  const base = table.offset;
  if (isLong) {
    for (let i = 0; i <= numGlyphs; i++) {
      offsets[i] = u32(data, base + i * 4);
    }
  } else {
    for (let i = 0; i <= numGlyphs; i++) {
      offsets[i] = u16(data, base + i * 2) * 2;
    }
  }
  return offsets;
}

function parseHmtx(
  data: Uint8Array,
  table: TableEntry,
  numHMetrics: number,
  numGlyphs: number
): Uint16Array {
  const widths = new Uint16Array(numGlyphs);
  const base = table.offset;
  let lastWidth = 0;
  for (let i = 0; i < numHMetrics; i++) {
    lastWidth = u16(data, base + i * 4);
    widths[i] = lastWidth;
  }
  for (let i = numHMetrics; i < numGlyphs; i++) {
    widths[i] = lastWidth;
  }
  return widths;
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

// Composite glyph flags
const COMP_ARG_1_AND_2_ARE_WORDS = 0x0001;
const COMP_ARGS_ARE_XY_VALUES = 0x0002;
const COMP_WE_HAVE_A_SCALE = 0x0008;
const COMP_MORE_COMPONENTS = 0x0020;
const COMP_WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
const COMP_WE_HAVE_A_TWO_BY_TWO = 0x0080;

function parseCompositeGlyph(
  data: Uint8Array,
  offset: number,
  glyfBase: number,
  glyphOffsets: Uint32Array
): GlyphPoint[][] {
  let pos = offset + 10; // skip header
  const allContours: GlyphPoint[][] = [];

  while (true) {
    const flags = u16(data, pos);
    pos += 2;
    const componentGid = u16(data, pos);
    pos += 2;

    // Read translation
    let dx: number;
    let dy: number;
    if (flags & COMP_ARG_1_AND_2_ARE_WORDS) {
      if (flags & COMP_ARGS_ARE_XY_VALUES) {
        dx = i16(data, pos);
        dy = i16(data, pos + 2);
      } else {
        dx = 0;
        dy = 0;
      }
      pos += 4;
    } else {
      if (flags & COMP_ARGS_ARE_XY_VALUES) {
        dx = data[pos] >= 0x80 ? data[pos] - 256 : data[pos];
        dy = data[pos + 1] >= 0x80 ? data[pos + 1] - 256 : data[pos + 1];
      } else {
        dx = 0;
        dy = 0;
      }
      pos += 2;
    }

    // Read scale/transform
    let a = 1;
    let b = 0;
    let c = 0;
    let d = 1;
    if (flags & COMP_WE_HAVE_A_SCALE) {
      a = d = i16(data, pos) / 16384;
      pos += 2;
    } else if (flags & COMP_WE_HAVE_AN_X_AND_Y_SCALE) {
      a = i16(data, pos) / 16384;
      d = i16(data, pos + 2) / 16384;
      pos += 4;
    } else if (flags & COMP_WE_HAVE_A_TWO_BY_TWO) {
      a = i16(data, pos) / 16384;
      b = i16(data, pos + 2) / 16384;
      c = i16(data, pos + 4) / 16384;
      d = i16(data, pos + 6) / 16384;
      pos += 8;
    }

    // Recursively get component outlines
    const compContours = getGlyphContours(data, componentGid, glyfBase, glyphOffsets);
    for (const contour of compContours) {
      const transformed = contour.map(pt => ({
        x: a * pt.x + c * pt.y + dx,
        y: b * pt.x + d * pt.y + dy,
        onCurve: pt.onCurve
      }));
      allContours.push(transformed);
    }

    if (!(flags & COMP_MORE_COMPONENTS)) {
      break;
    }
  }
  return allContours;
}

function getGlyphContours(
  data: Uint8Array,
  glyphId: number,
  glyfBase: number,
  glyphOffsets: Uint32Array
): GlyphPoint[][] {
  const start = glyphOffsets[glyphId];
  const end = glyphOffsets[glyphId + 1];
  if (end - start < 10) {
    return []; // empty glyph (e.g. space)
  }

  const offset = glyfBase + start;
  const numberOfContours = i16(data, offset);

  if (numberOfContours >= 0) {
    return parseSimpleGlyph(data, offset, numberOfContours);
  }
  return parseCompositeGlyph(data, offset, glyfBase, glyphOffsets);
}

// =============================================================================
// Contour to line segments (flattening quadratic B-splines)
// =============================================================================

interface Segment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

/**
 * Flatten a contour (with on-curve and off-curve points) into line segments.
 * TrueType uses quadratic B-splines: between two consecutive off-curve
 * points an implicit on-curve midpoint is inserted.
 */
function flattenContour(contour: GlyphPoint[]): Segment[] {
  if (contour.length < 2) {
    return [];
  }

  const segments: Segment[] = [];
  const n = contour.length;

  // Find first on-curve point (or synthesize one)
  let startIdx = 0;
  let startPt: { x: number; y: number };
  if (contour[0].onCurve) {
    startPt = contour[0];
    startIdx = 1;
  } else if (contour[n - 1].onCurve) {
    startPt = contour[n - 1];
    startIdx = 0;
  } else {
    // Both first and last are off-curve; start at midpoint
    startPt = {
      x: (contour[0].x + contour[n - 1].x) / 2,
      y: (contour[0].y + contour[n - 1].y) / 2
    };
    startIdx = 0;
  }

  let cur = startPt;

  for (let i = startIdx; i < n; i++) {
    const pt = contour[i];
    if (pt.onCurve) {
      segments.push({ x1: cur.x, y1: cur.y, x2: pt.x, y2: pt.y });
      cur = pt;
    } else {
      // Off-curve: find next on-curve (or implicit midpoint)
      let nextOn: { x: number; y: number };
      const nextIdx = (i + 1) % n;
      const next = contour[nextIdx];
      if (next.onCurve) {
        nextOn = next;
        i++; // skip next since we consumed it
        if (nextIdx === 0) {
          // We've wrapped around; use startPt
          nextOn = startPt;
          i = n; // exit loop after this
        }
      } else {
        // Implicit on-curve at midpoint
        nextOn = { x: (pt.x + next.x) / 2, y: (pt.y + next.y) / 2 };
      }
      // Subdivide quadratic bezier: cur, pt(control), nextOn
      subdivideQuadratic(cur.x, cur.y, pt.x, pt.y, nextOn.x, nextOn.y, segments);
      cur = nextOn;
    }
  }

  // Close the contour
  if (cur.x !== startPt.x || cur.y !== startPt.y) {
    segments.push({ x1: cur.x, y1: cur.y, x2: startPt.x, y2: startPt.y });
  }

  return segments;
}

function subdivideQuadratic(
  x0: number,
  y0: number,
  cx: number,
  cy: number,
  x1: number,
  y1: number,
  segments: Segment[]
): void {
  // Adaptive subdivision based on flatness
  const steps = 8; // fixed subdivision — good enough for chart labels
  let prevX = x0;
  let prevY = y0;
  for (let i = 1; i <= steps; i++) {
    const t = i / steps;
    const mt = 1 - t;
    const nx = mt * mt * x0 + 2 * mt * t * cx + t * t * x1;
    const ny = mt * mt * y0 + 2 * mt * t * cy + t * t * y1;
    segments.push({ x1: prevX, y1: prevY, x2: nx, y2: ny });
    prevX = nx;
    prevY = ny;
  }
}

// =============================================================================
// =============================================================================
// Public API
// =============================================================================

/**
 * Parse a TTF font file into a RasterFont that can render glyphs to pixels.
 */
export function parseRasterFont(data: Uint8Array): RasterFont {
  // Handle TTC (TrueType Collection) — use first font
  const sfVersion = u32(data, 0);
  const fontData = data;
  if (sfVersion === 0x74746366) {
    // 'ttcf' header
    const firstOffset = u32(data, 12);
    // Re-read tables relative to the font offset — but table offsets in TTC
    // are absolute, so we keep original data and adjust
    const tables = parseTableDirectoryTTC(data, firstOffset);
    return buildRasterFont(data, tables);
  }

  const tables = parseTableDirectory(fontData);
  return buildRasterFont(fontData, tables);
}

function parseTableDirectoryTTC(data: Uint8Array, fontOffset: number): Map<string, TableEntry> {
  const numTables = u16(data, fontOffset + 4);
  const tables = new Map<string, TableEntry>();
  for (let i = 0; i < numTables; i++) {
    const rec = fontOffset + 12 + i * 16;
    const t = tag(data, rec);
    tables.set(t, { offset: u32(data, rec + 8), length: u32(data, rec + 12) });
  }
  return tables;
}

function buildRasterFont(data: Uint8Array, tables: Map<string, TableEntry>): RasterFont {
  const head = tables.get("head");
  const hhea = tables.get("hhea");
  const maxp = tables.get("maxp");
  const cmapTable = tables.get("cmap");
  const hmtxTable = tables.get("hmtx");
  const locaTable = tables.get("loca");
  const glyfTable = tables.get("glyf");

  if (!head || !hhea || !maxp || !cmapTable || !hmtxTable || !locaTable || !glyfTable) {
    // Return a dummy font if tables are missing
    return { unitsPerEm: 1000, ascent: 800, descent: -200, getOutline: () => undefined };
  }

  const unitsPerEm = u16(data, head.offset + 18);
  const indexToLocFormat = i16(data, head.offset + 50);
  const ascent = i16(data, hhea.offset + 4);
  const descent = i16(data, hhea.offset + 6);
  const numHMetrics = u16(data, hhea.offset + 34);
  const numGlyphs = u16(data, maxp.offset + 4);

  const cmap = parseCmap(data, cmapTable);
  const advanceWidths = parseHmtx(data, hmtxTable, numHMetrics, numGlyphs);
  const glyphOffsets = parseLoca(data, locaTable, numGlyphs, indexToLocFormat !== 0);
  const glyfBase = glyfTable.offset;

  return {
    unitsPerEm,
    ascent,
    descent,
    getOutline(codePoint: number): GlyphOutline | undefined {
      const gid = cmap.get(codePoint);
      if (gid === undefined || gid === 0) {
        return undefined;
      }
      const contours = getGlyphContours(data, gid, glyfBase, glyphOffsets);
      if (contours.length === 0) {
        return undefined;
      }
      return { contours, advanceWidth: advanceWidths[gid] };
    }
  };
}

/**
 * Rasterize a single glyph into an alpha bitmap with 4x supersampled
 * anti-aliasing for smooth edges.
 *
 * @param outline - Glyph outline from RasterFont.getOutline()
 * @param fontSize - Target font size in pixels
 * @param unitsPerEm - Font's unitsPerEm
 * @returns { width, height, offsetX, offsetY, pixels }
 *   offsetX/offsetY are pixel offsets from the pen position (left of baseline)
 *   to the top-left of the bitmap.  pixels values are 0–255 (coverage).
 */
export function rasterizeGlyph(
  outline: GlyphOutline,
  fontSize: number,
  unitsPerEm: number
): { width: number; height: number; offsetX: number; offsetY: number; pixels: Uint8Array } {
  const scale = fontSize / unitsPerEm;

  // Supersample factor — render at Nx resolution then downsample
  const SS = 4;
  const ssScale = scale * SS;

  // Scale all contour points to hi-res pixel space, flipping Y
  const scaledContours: GlyphPoint[][] = outline.contours.map(contour =>
    contour.map(pt => ({
      x: pt.x * ssScale,
      y: -pt.y * ssScale,
      onCurve: pt.onCurve
    }))
  );

  // Find bounding box in hi-res pixel space
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const contour of scaledContours) {
    for (const pt of contour) {
      if (pt.x < minX) {
        minX = pt.x;
      }
      if (pt.x > maxX) {
        maxX = pt.x;
      }
      if (pt.y < minY) {
        minY = pt.y;
      }
      if (pt.y > maxY) {
        maxY = pt.y;
      }
    }
  }

  if (!Number.isFinite(minX)) {
    return {
      width: Math.ceil(outline.advanceWidth * scale),
      height: Math.ceil(fontSize),
      offsetX: 0,
      offsetY: 0,
      pixels: new Uint8Array(0)
    };
  }

  // Output bitmap dimensions (in final 1x pixels)
  const pad = 1;
  const bmpW = Math.ceil((maxX - minX) / SS) + pad * 2;
  const bmpH = Math.ceil((maxY - minY) / SS) + pad * 2;

  if (bmpW <= 0 || bmpH <= 0) {
    return {
      width: Math.ceil(outline.advanceWidth * scale),
      height: Math.ceil(fontSize),
      offsetX: 0,
      offsetY: 0,
      pixels: new Uint8Array(0)
    };
  }

  // Hi-res bitmap dimensions
  const hiW = bmpW * SS;
  const hiH = bmpH * SS;

  // Translate contours so that minX,minY maps to (pad*SS, pad*SS) in hi-res space
  const txOff = -minX + pad * SS;
  const tyOff = -minY + pad * SS;

  // Flatten contours into line segments (in hi-res pixel coordinates)
  const allSegments: Segment[] = [];
  for (const contour of scaledContours) {
    const translated = contour.map(pt => ({
      x: pt.x + txOff,
      y: pt.y + tyOff,
      onCurve: pt.onCurve
    }));
    allSegments.push(...flattenContour(translated));
  }

  // Scan-line fill at hi-res
  const hiBuf = new Uint8Array(hiW * hiH);
  for (let row = 0; row < hiH; row++) {
    const scanY = row + 0.5;

    const intersections: number[] = [];
    for (const seg of allSegments) {
      const y1 = seg.y1;
      const y2 = seg.y2;
      if ((y1 <= scanY && y2 > scanY) || (y2 <= scanY && y1 > scanY)) {
        const t = (scanY - y1) / (y2 - y1);
        intersections.push(seg.x1 + t * (seg.x2 - seg.x1));
      }
    }

    intersections.sort((a, b) => a - b);
    for (let i = 0; i < intersections.length - 1; i += 2) {
      const xStart = Math.max(0, Math.ceil(intersections[i]));
      const xEnd = Math.min(hiW - 1, Math.floor(intersections[i + 1]));
      for (let x = xStart; x <= xEnd; x++) {
        hiBuf[row * hiW + x] = 1;
      }
    }
  }

  // Downsample: average SS×SS blocks → 0–255 coverage
  const pixels = new Uint8Array(bmpW * bmpH);
  const ss2 = SS * SS;
  for (let py = 0; py < bmpH; py++) {
    for (let px = 0; px < bmpW; px++) {
      let count = 0;
      const hiBaseY = py * SS;
      const hiBaseX = px * SS;
      for (let sy = 0; sy < SS; sy++) {
        const hiRow = hiBaseY + sy;
        if (hiRow >= hiH) {
          break;
        }
        for (let sx = 0; sx < SS; sx++) {
          const hiCol = hiBaseX + sx;
          if (hiCol >= hiW) {
            break;
          }
          count += hiBuf[hiRow * hiW + hiCol];
        }
      }
      if (count > 0) {
        pixels[py * bmpW + px] = Math.round((count / ss2) * 255);
      }
    }
  }

  return {
    width: bmpW,
    height: bmpH,
    offsetX: Math.floor(minX / SS) - pad,
    offsetY: Math.floor(minY / SS) - pad,
    pixels
  };
}

// =============================================================================
// System font loading (Node.js only)
// =============================================================================

let _cachedFont: RasterFont | null = null;
let _fontLoadAttempted = false;

/**
 * Load a system font for text rasterization.
 * Returns null in browser environments or if no font is found.
 * Results are cached — only loads once.
 */
export function loadSystemFont(): RasterFont | null {
  if (_fontLoadAttempted) {
    return _cachedFont;
  }
  _fontLoadAttempted = true;

  try {
    const fontPaths = getSystemFontPaths();
    for (const fontPath of fontPaths) {
      try {
        if (fileExistsSync(fontPath)) {
          const data = readFileBytesSync(fontPath);
          _cachedFont = parseRasterFont(data);
          return _cachedFont;
        }
      } catch {
        continue;
      }
    }
  } catch {
    // Not in Node.js or fs not available
  }

  return null;
}

function getSystemFontPaths(): string[] {
  const platform = typeof process !== "undefined" ? process.platform : "";
  const paths: string[] = [];

  if (platform === "darwin") {
    // macOS: prefer Arial, then Helvetica, then SF Pro
    paths.push(
      "/System/Library/Fonts/Supplemental/Arial.ttf",
      "/Library/Fonts/Arial.ttf",
      "/System/Library/Fonts/Helvetica.ttc",
      "/System/Library/Fonts/SFNSText.ttf",
      "/System/Library/Fonts/SFNS.ttf"
    );
  } else if (platform === "win32") {
    const windir = process.env.WINDIR || process.env.windir || "C:\\Windows";
    paths.push(
      `${windir}\\Fonts\\arial.ttf`,
      `${windir}\\Fonts\\calibri.ttf`,
      `${windir}\\Fonts\\segoeui.ttf`,
      `${windir}\\Fonts\\tahoma.ttf`
    );
  } else {
    // Linux
    paths.push(
      "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
      "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
      "/usr/share/fonts/TTF/DejaVuSans.ttf",
      "/usr/share/fonts/noto/NotoSans-Regular.ttf",
      "/usr/share/fonts/truetype/freefont/FreeSans.ttf"
    );
  }

  return paths;
}

/** Reset cached font (for testing). */
export function resetCachedFont(): void {
  _cachedFont = null;
  _fontLoadAttempted = false;
}
