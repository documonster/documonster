/**
 * Glyph contours → an alpha bitmap.
 *
 * The tail end of the text pipeline: it takes a {@link GlyphOutline} decoded by
 * `@draw/raster/glyph-outline`, flattens its quadratic B-splines into line
 * segments, and scan-line fills them at 4x supersampling for anti-aliased edges.
 *
 * Nothing here knows what a font file looks like. It used to: this module carried
 * its own TrueType table parser, a duplicate of the one in `@pdf/font/ttf-parser`.
 * Parsing now lives once at Layer 0 in `@utils/font-ttf`, outline decoding in
 * `glyph-outline.ts`, and this file is the geometry.
 *
 * @module
 */

import type { GlyphOutline, GlyphPoint } from "@draw/raster/glyph-outline";

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
// Rasterization
// =============================================================================

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
