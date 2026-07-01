/**
 * Type3 fallback glyph definitions.
 *
 * Each glyph is a function that writes PDF content-stream operators into
 * a 1000×1000 unit glyph coordinate system (matching a FontMatrix of
 * [0.001 0 0 0.001 0 0]).
 *
 * The drawing callbacks receive a simple `GlyphPen` API so they don't
 * depend on `PdfContentStream` directly — this keeps the module
 * self-contained and easy to test.
 *
 * Coverage strategy:
 *  1. Geometric Shapes          (U+25A0–U+25FF)
 *  2. Miscellaneous Symbols     (U+2600–U+26FF)
 *  3. Dingbats / checkmarks     (U+2700–U+27BF)
 *  4. Arrows                    (U+2190–U+21FF)
 *  5. Mathematical Operators    (U+2200–U+22FF)
 *  6. Box Drawing               (U+2500–U+257F)
 *  7. Block Elements            (U+2580–U+259F)
 *  8. Misc Technical            (U+2300–U+23FF)
 *  9. Enclosed Alphanumerics    (U+2460–U+24FF)
 * 10. Misc Symbols & Arrows     (U+2B00–U+2BFF)
 * 11. Ballot / checkbox symbols (U+2610–U+2612)
 * 12. Supplemental Arrows       (U+27F0–U+27FF)
 * 13. Braille Patterns          (U+2800–U+28FF)
 * 14. Squared symbols etc.      (U+29C0–U+29FF)
 *
 * Characters not in this table get a generic .notdef glyph (open rectangle).
 */

// =============================================================================
// Glyph Pen — abstraction for drawing glyph outlines
// =============================================================================

/**
 * Minimal drawing API used by glyph definitions.
 * Lines are emitted as raw PDF content-stream operators.
 */
export interface GlyphPen {
  /** Move to (x, y). */
  M(x: number, y: number): void;
  /** Line to (x, y). */
  L(x: number, y: number): void;
  /** Cubic Bezier curve. */
  C(x1: number, y1: number, x2: number, y2: number, x3: number, y3: number): void;
  /** Close path. */
  Z(): void;
  /** Rectangle (x, y, w, h). */
  rect(x: number, y: number, w: number, h: number): void;
  /** Circle at (cx, cy) with radius r using 4 Bezier arcs. */
  circle(cx: number, cy: number, r: number): void;
  /** Ellipse at (cx, cy) with radii (rx, ry). */
  ellipse(cx: number, cy: number, rx: number, ry: number): void;
  /** Stroke current path. */
  stroke(): void;
  /** Fill current path (non-zero winding). */
  fill(): void;
  /** Fill then stroke. */
  fillStroke(): void;
  /** Set line width. */
  lineWidth(w: number): void;
}

/** A glyph definition: advance width + drawing callback. */
export interface GlyphDef {
  /** Advance width in glyph units (0–1000). */
  width: number;
  /** Draw the glyph outline using the pen. */
  draw: (p: GlyphPen) => void;
}

// =============================================================================
// Helpers
// =============================================================================

/** Standard advance width for a "normal-width" symbol. */
const W = 600;

/** A filled square glyph (generic, reusable). */
function filledRect(x: number, y: number, w: number, h: number): GlyphDef {
  return {
    width: W,
    draw: p => {
      p.rect(x, y, w, h);
      p.fill();
    }
  };
}

function strokedRect(x: number, y: number, w: number, h: number, lw = 50): GlyphDef {
  return {
    width: W,
    draw: p => {
      p.lineWidth(lw);
      p.rect(x, y, w, h);
      p.stroke();
    }
  };
}

function filledCircle(cx: number, cy: number, r: number): GlyphDef {
  return {
    width: W,
    draw: p => {
      p.circle(cx, cy, r);
      p.fill();
    }
  };
}

function strokedCircle(cx: number, cy: number, r: number, lw = 50): GlyphDef {
  return {
    width: W,
    draw: p => {
      p.lineWidth(lw);
      p.circle(cx, cy, r);
      p.stroke();
    }
  };
}

/** .notdef: open rectangle — drawn for any character we don't have. */
export const NOTDEF_GLYPH: GlyphDef = strokedRect(80, 0, 440, 700, 40);

// =============================================================================
// Geometric Shapes (U+25A0 – U+25FF)
// =============================================================================

const GEO: Record<number, GlyphDef> = {
  // ■ BLACK SQUARE
  [0x25a0]: filledRect(100, 50, 400, 400),
  // □ WHITE SQUARE
  [0x25a1]: strokedRect(100, 50, 400, 400),
  // ▢ WHITE SQUARE WITH ROUNDED CORNERS
  [0x25a2]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(150, 50);
      p.L(450, 50);
      p.C(480, 50, 500, 70, 500, 100);
      p.L(500, 400);
      p.C(500, 430, 480, 450, 450, 450);
      p.L(150, 450);
      p.C(120, 450, 100, 430, 100, 400);
      p.L(100, 100);
      p.C(100, 70, 120, 50, 150, 50);
      p.Z();
      p.stroke();
    }
  },
  // ▣ WHITE SQUARE CONTAINING BLACK SMALL SQUARE
  [0x25a3]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.rect(200, 150, 200, 200);
      p.fill();
    }
  },
  // ▤ SQUARE WITH HORIZONTAL FILL
  [0x25a4]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.lineWidth(30);
      for (let y = 100; y <= 400; y += 80) {
        p.M(100, y);
        p.L(500, y);
      }
      p.stroke();
    }
  },
  // ▥ SQUARE WITH VERTICAL FILL
  [0x25a5]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.lineWidth(30);
      for (let x = 150; x <= 450; x += 80) {
        p.M(x, 50);
        p.L(x, 450);
      }
      p.stroke();
    }
  },
  // ▦ SQUARE WITH ORTHOGONAL CROSSHATCH
  [0x25a6]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.lineWidth(20);
      for (let y = 100; y <= 400; y += 80) {
        p.M(100, y);
        p.L(500, y);
      }
      for (let x = 150; x <= 450; x += 80) {
        p.M(x, 50);
        p.L(x, 450);
      }
      p.stroke();
    }
  },
  // ▨ SQUARE WITH UPPER RIGHT TO LOWER LEFT FILL — simplified as rect
  [0x25a8]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.lineWidth(20);
      for (let i = -400; i <= 400; i += 80) {
        p.M(Math.max(100, 100 + i), Math.min(450, 450 + i));
        p.L(Math.min(500, 500 + i), Math.max(50, 50 + i));
      }
      p.stroke();
    }
  },
  // ▪ BLACK SMALL SQUARE
  [0x25aa]: filledRect(175, 125, 250, 250),
  // ▫ WHITE SMALL SQUARE
  [0x25ab]: strokedRect(175, 125, 250, 250, 40),
  // ▬ BLACK RECTANGLE
  [0x25ac]: filledRect(80, 150, 440, 200),
  // ▭ WHITE RECTANGLE
  [0x25ad]: strokedRect(80, 150, 440, 200),
  // ▮ BLACK VERTICAL RECTANGLE
  [0x25ae]: filledRect(175, 0, 250, 500),
  // ▯ WHITE VERTICAL RECTANGLE
  [0x25af]: strokedRect(175, 0, 250, 500),
  // ▰ BLACK PARALLELOGRAM
  [0x25b0]: {
    width: W,
    draw: p => {
      p.M(150, 100);
      p.L(520, 100);
      p.L(450, 400);
      p.L(80, 400);
      p.Z();
      p.fill();
    }
  },
  // ▲ BLACK UP-POINTING TRIANGLE
  [0x25b2]: {
    width: W,
    draw: p => {
      p.M(300, 500);
      p.L(500, 50);
      p.L(100, 50);
      p.Z();
      p.fill();
    }
  },
  // △ WHITE UP-POINTING TRIANGLE
  [0x25b3]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 500);
      p.L(500, 50);
      p.L(100, 50);
      p.Z();
      p.stroke();
    }
  },
  // ▴ BLACK UP-POINTING SMALL TRIANGLE
  [0x25b4]: {
    width: W,
    draw: p => {
      p.M(300, 450);
      p.L(450, 100);
      p.L(150, 100);
      p.Z();
      p.fill();
    }
  },
  // ▵ WHITE UP-POINTING SMALL TRIANGLE
  [0x25b5]: {
    width: W,
    draw: p => {
      p.lineWidth(35);
      p.M(300, 450);
      p.L(450, 100);
      p.L(150, 100);
      p.Z();
      p.stroke();
    }
  },
  // ▶ BLACK RIGHT-POINTING TRIANGLE
  [0x25b6]: {
    width: W,
    draw: p => {
      p.M(100, 500);
      p.L(500, 250);
      p.L(100, 0);
      p.Z();
      p.fill();
    }
  },
  // ▷ WHITE RIGHT-POINTING TRIANGLE
  [0x25b7]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(100, 500);
      p.L(500, 250);
      p.L(100, 0);
      p.Z();
      p.stroke();
    }
  },
  // ▸ BLACK RIGHT-POINTING SMALL TRIANGLE
  [0x25b8]: {
    width: W,
    draw: p => {
      p.M(150, 450);
      p.L(450, 250);
      p.L(150, 50);
      p.Z();
      p.fill();
    }
  },
  // ▼ BLACK DOWN-POINTING TRIANGLE
  [0x25bc]: {
    width: W,
    draw: p => {
      p.M(100, 500);
      p.L(500, 500);
      p.L(300, 50);
      p.Z();
      p.fill();
    }
  },
  // ▽ WHITE DOWN-POINTING TRIANGLE
  [0x25bd]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(100, 500);
      p.L(500, 500);
      p.L(300, 50);
      p.Z();
      p.stroke();
    }
  },
  // ▾ BLACK DOWN-POINTING SMALL TRIANGLE
  [0x25be]: {
    width: W,
    draw: p => {
      p.M(150, 400);
      p.L(450, 400);
      p.L(300, 100);
      p.Z();
      p.fill();
    }
  },
  // ◀ BLACK LEFT-POINTING TRIANGLE
  [0x25c0]: {
    width: W,
    draw: p => {
      p.M(500, 500);
      p.L(100, 250);
      p.L(500, 0);
      p.Z();
      p.fill();
    }
  },
  // ◁ WHITE LEFT-POINTING TRIANGLE
  [0x25c1]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(500, 500);
      p.L(100, 250);
      p.L(500, 0);
      p.Z();
      p.stroke();
    }
  },
  // ◂ BLACK LEFT-POINTING SMALL TRIANGLE
  [0x25c2]: {
    width: W,
    draw: p => {
      p.M(450, 450);
      p.L(150, 250);
      p.L(450, 50);
      p.Z();
      p.fill();
    }
  },
  // ◆ BLACK DIAMOND
  [0x25c6]: {
    width: W,
    draw: p => {
      p.M(300, 500);
      p.L(530, 250);
      p.L(300, 0);
      p.L(70, 250);
      p.Z();
      p.fill();
    }
  },
  // ◇ WHITE DIAMOND
  [0x25c7]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 500);
      p.L(530, 250);
      p.L(300, 0);
      p.L(70, 250);
      p.Z();
      p.stroke();
    }
  },
  // ◈ WHITE DIAMOND CONTAINING BLACK SMALL DIAMOND
  [0x25c8]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 500);
      p.L(530, 250);
      p.L(300, 0);
      p.L(70, 250);
      p.Z();
      p.stroke();
      p.M(300, 380);
      p.L(410, 250);
      p.L(300, 120);
      p.L(190, 250);
      p.Z();
      p.fill();
    }
  },
  // ◉ FISHEYE (big circle with smaller filled circle)
  [0x25c9]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.circle(300, 250, 120);
      p.fill();
    }
  },
  // ◊ LOZENGE
  [0x25ca]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 550);
      p.L(500, 250);
      p.L(300, -50);
      p.L(100, 250);
      p.Z();
      p.stroke();
    }
  },
  // ○ WHITE CIRCLE
  [0x25cb]: strokedCircle(300, 250, 220),
  // ◌ DOTTED CIRCLE — simplified as dashed circle
  [0x25cc]: strokedCircle(300, 250, 220, 30),
  // ◍ CIRCLE WITH UPPER HALF BLACK
  [0x25cd]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      /* upper half */ p.M(80, 250);
      p.C(80, 371, 179, 470, 300, 470);
      p.C(421, 470, 520, 371, 520, 250);
      p.Z();
      p.fill();
    }
  },
  // ◎ BULLSEYE
  [0x25ce]: {
    width: W,
    draw: p => {
      p.lineWidth(35);
      p.circle(300, 250, 220);
      p.stroke();
      p.circle(300, 250, 140);
      p.stroke();
    }
  },
  // ● BLACK CIRCLE
  [0x25cf]: filledCircle(300, 250, 220),
  // ◐ CIRCLE WITH LEFT HALF BLACK
  [0x25d0]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      /* left half */ p.M(300, 30);
      p.C(179, 30, 80, 129, 80, 250);
      p.C(80, 371, 179, 470, 300, 470);
      p.Z();
      p.fill();
    }
  },
  // ◑ CIRCLE WITH RIGHT HALF BLACK
  [0x25d1]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(300, 470);
      p.C(421, 470, 520, 371, 520, 250);
      p.C(520, 129, 421, 30, 300, 30);
      p.Z();
      p.fill();
    }
  },
  // ◒ CIRCLE WITH LOWER HALF BLACK
  [0x25d2]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(520, 250);
      p.C(520, 129, 421, 30, 300, 30);
      p.C(179, 30, 80, 129, 80, 250);
      p.Z();
      p.fill();
    }
  },
  // ◓ CIRCLE WITH UPPER HALF BLACK
  [0x25d3]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(80, 250);
      p.C(80, 371, 179, 470, 300, 470);
      p.C(421, 470, 520, 371, 520, 250);
      p.Z();
      p.fill();
    }
  },
  // ◔ CIRCLE WITH UPPER RIGHT QUADRANT BLACK
  [0x25d4]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(300, 250);
      p.L(300, 470);
      p.C(421, 470, 520, 371, 520, 250);
      p.Z();
      p.fill();
    }
  },
  // ◕ CIRCLE WITH ALL BUT UPPER LEFT QUADRANT BLACK
  [0x25d5]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(300, 250);
      p.L(300, 470);
      p.C(179, 470, 80, 371, 80, 250);
      p.Z();
      p.fill();
      p.M(300, 250);
      p.L(520, 250);
      p.C(520, 129, 421, 30, 300, 30);
      p.C(179, 30, 80, 129, 80, 250);
      p.Z();
      p.fill();
    }
  },
  // ◖ LEFT HALF BLACK CIRCLE
  [0x25d6]: {
    width: W,
    draw: p => {
      p.M(300, 30);
      p.C(179, 30, 80, 129, 80, 250);
      p.C(80, 371, 179, 470, 300, 470);
      p.Z();
      p.fill();
    }
  },
  // ◗ RIGHT HALF BLACK CIRCLE
  [0x25d7]: {
    width: W,
    draw: p => {
      p.M(300, 470);
      p.C(421, 470, 520, 371, 520, 250);
      p.C(520, 129, 421, 30, 300, 30);
      p.Z();
      p.fill();
    }
  },
  // ◘ INVERSE BULLET (filled square, white circle inside)
  [0x25d8]: {
    width: W,
    draw: p => {
      p.rect(80, 30, 440, 440);
      p.fill();
    }
  },
  // ◙ INVERSE WHITE CIRCLE (filled circle w/ white square)
  [0x25d9]: {
    width: W,
    draw: p => {
      p.circle(300, 250, 240);
      p.fill();
    }
  },
  // ◜◝◞◟  UPPER LEFT / UPPER RIGHT / LOWER RIGHT / LOWER LEFT QUADRANT CIRCULAR ARC
  [0x25dc]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(300, 470);
      p.C(179, 470, 80, 371, 80, 250);
      p.stroke();
    }
  },
  [0x25dd]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(520, 250);
      p.C(520, 371, 421, 470, 300, 470);
      p.stroke();
    }
  },
  [0x25de]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(300, 30);
      p.C(421, 30, 520, 129, 520, 250);
      p.stroke();
    }
  },
  [0x25df]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(80, 250);
      p.C(80, 129, 179, 30, 300, 30);
      p.stroke();
    }
  },
  // ◠ UPPER HALF CIRCLE
  [0x25e0]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(80, 250);
      p.C(80, 371, 179, 470, 300, 470);
      p.C(421, 470, 520, 371, 520, 250);
      p.stroke();
    }
  },
  // ◡ LOWER HALF CIRCLE
  [0x25e1]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(520, 250);
      p.C(520, 129, 421, 30, 300, 30);
      p.C(179, 30, 80, 129, 80, 250);
      p.stroke();
    }
  },
  // ◦ WHITE BULLET
  [0x25e6]: strokedCircle(300, 250, 100, 40),
  // ◯ LARGE CIRCLE
  [0x25ef]: strokedCircle(300, 250, 280, 40)
};

// =============================================================================
// Arrows (U+2190 – U+21FF)
// =============================================================================

const ARR: Record<number, GlyphDef> = {
  // ← LEFTWARDS ARROW
  [0x2190]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(250, 380);
      p.L(250, 120);
      p.Z();
      p.fill();
    }
  },
  // ↑ UPWARDS ARROW
  [0x2191]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(170, 300);
      p.L(430, 300);
      p.Z();
      p.fill();
    }
  },
  // → RIGHTWARDS ARROW
  [0x2192]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(350, 380);
      p.L(350, 120);
      p.Z();
      p.fill();
    }
  },
  // ↓ DOWNWARDS ARROW
  [0x2193]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
      p.M(300, 50);
      p.L(170, 200);
      p.L(430, 200);
      p.Z();
      p.fill();
    }
  },
  // ↔ LEFT RIGHT ARROW
  [0x2194]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(150, 250);
      p.L(450, 250);
      p.stroke();
      p.M(100, 250);
      p.L(220, 360);
      p.L(220, 140);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(380, 360);
      p.L(380, 140);
      p.Z();
      p.fill();
    }
  },
  // ↕ UP DOWN ARROW
  [0x2195]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 100);
      p.L(300, 400);
      p.stroke();
      p.M(300, 450);
      p.L(190, 330);
      p.L(410, 330);
      p.Z();
      p.fill();
      p.M(300, 50);
      p.L(190, 170);
      p.L(410, 170);
      p.Z();
      p.fill();
    }
  },
  // ↖ NORTH WEST ARROW
  [0x2196]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(480, 50);
      p.L(120, 410);
      p.stroke();
      p.M(120, 450);
      p.L(120, 320);
      p.L(250, 450);
      p.Z();
      p.fill();
    }
  },
  // ↗ NORTH EAST ARROW
  [0x2197]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 50);
      p.L(480, 410);
      p.stroke();
      p.M(480, 450);
      p.L(350, 450);
      p.L(480, 320);
      p.Z();
      p.fill();
    }
  },
  // ↘ SOUTH EAST ARROW
  [0x2198]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 450);
      p.L(480, 90);
      p.stroke();
      p.M(480, 50);
      p.L(480, 180);
      p.L(350, 50);
      p.Z();
      p.fill();
    }
  },
  // ↙ SOUTH WEST ARROW
  [0x2199]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(480, 450);
      p.L(120, 90);
      p.stroke();
      p.M(120, 50);
      p.L(250, 50);
      p.L(120, 180);
      p.Z();
      p.fill();
    }
  },
  // ↩ LEFTWARDS ARROW WITH HOOK
  [0x21a9]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(500, 350);
      p.L(200, 350);
      p.C(120, 350, 80, 300, 80, 250);
      p.C(80, 200, 120, 150, 200, 150);
      p.L(350, 150);
      p.stroke();
      p.M(100, 350);
      p.L(230, 430);
      p.L(230, 270);
      p.Z();
      p.fill();
    }
  },
  // ↪ RIGHTWARDS ARROW WITH HOOK
  [0x21aa]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(100, 350);
      p.L(400, 350);
      p.C(480, 350, 520, 300, 520, 250);
      p.C(520, 200, 480, 150, 400, 150);
      p.L(250, 150);
      p.stroke();
      p.M(500, 350);
      p.L(370, 430);
      p.L(370, 270);
      p.Z();
      p.fill();
    }
  },
  // ⇐ LEFTWARDS DOUBLE ARROW
  [0x21d0]: {
    width: W,
    draw: p => {
      p.M(100, 250);
      p.L(280, 420);
      p.L(280, 300);
      p.L(500, 300);
      p.L(500, 200);
      p.L(280, 200);
      p.L(280, 80);
      p.Z();
      p.fill();
    }
  },
  // ⇑ UPWARDS DOUBLE ARROW
  [0x21d1]: {
    width: W,
    draw: p => {
      p.M(300, 480);
      p.L(480, 280);
      p.L(370, 280);
      p.L(370, 30);
      p.L(230, 30);
      p.L(230, 280);
      p.L(120, 280);
      p.Z();
      p.fill();
    }
  },
  // ⇒ RIGHTWARDS DOUBLE ARROW
  [0x21d2]: {
    width: W,
    draw: p => {
      p.M(500, 250);
      p.L(320, 420);
      p.L(320, 300);
      p.L(100, 300);
      p.L(100, 200);
      p.L(320, 200);
      p.L(320, 80);
      p.Z();
      p.fill();
    }
  },
  // ⇓ DOWNWARDS DOUBLE ARROW
  [0x21d3]: {
    width: W,
    draw: p => {
      p.M(300, 20);
      p.L(120, 220);
      p.L(230, 220);
      p.L(230, 470);
      p.L(370, 470);
      p.L(370, 220);
      p.L(480, 220);
      p.Z();
      p.fill();
    }
  },
  // ⇔ LEFT RIGHT DOUBLE ARROW
  [0x21d4]: {
    width: W,
    draw: p => {
      p.M(80, 250);
      p.L(220, 400);
      p.L(220, 300);
      p.L(380, 300);
      p.L(380, 400);
      p.L(520, 250);
      p.L(380, 100);
      p.L(380, 200);
      p.L(220, 200);
      p.L(220, 100);
      p.Z();
      p.fill();
    }
  }
};

// =============================================================================
// Mathematical Operators (U+2200 – U+22FF)
// =============================================================================

const MATH: Record<number, GlyphDef> = {
  // ∀ FOR ALL
  [0x2200]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(100, 500);
      p.L(300, 0);
      p.L(500, 500);
      p.stroke();
      p.M(170, 200);
      p.L(430, 200);
      p.stroke();
    }
  },
  // ∂ PARTIAL DIFFERENTIAL
  [0x2202]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.circle(300, 200, 150);
      p.stroke();
      p.M(450, 200);
      p.L(450, 450);
      p.C(450, 500, 400, 520, 350, 500);
      p.stroke();
    }
  },
  // ∃ THERE EXISTS
  [0x2203]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(450, 500);
      p.L(150, 500);
      p.L(150, 0);
      p.L(450, 0);
      p.stroke();
      p.M(150, 250);
      p.L(400, 250);
      p.stroke();
    }
  },
  // ∅ EMPTY SET
  [0x2205]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.ellipse(300, 250, 180, 220);
      p.stroke();
      p.M(150, 50);
      p.L(450, 450);
      p.stroke();
    }
  },
  // ∆ INCREMENT (triangle)
  [0x2206]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(300, 500);
      p.L(100, 0);
      p.L(500, 0);
      p.Z();
      p.stroke();
    }
  },
  // ∇ NABLA
  [0x2207]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(100, 500);
      p.L(500, 500);
      p.L(300, 0);
      p.Z();
      p.stroke();
    }
  },
  // ∈ ELEMENT OF
  [0x2208]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(480, 450);
      p.C(250, 450, 120, 380, 120, 250);
      p.C(120, 120, 250, 50, 480, 50);
      p.stroke();
      p.M(120, 250);
      p.L(430, 250);
      p.stroke();
    }
  },
  // ∉ NOT AN ELEMENT OF
  [0x2209]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(480, 450);
      p.C(250, 450, 120, 380, 120, 250);
      p.C(120, 120, 250, 50, 480, 50);
      p.stroke();
      p.M(120, 250);
      p.L(430, 250);
      p.stroke();
      p.M(160, 50);
      p.L(440, 450);
      p.stroke();
    }
  },
  // ∋ CONTAINS AS MEMBER
  [0x220b]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 450);
      p.C(350, 450, 480, 380, 480, 250);
      p.C(480, 120, 350, 50, 120, 50);
      p.stroke();
      p.M(170, 250);
      p.L(480, 250);
      p.stroke();
    }
  },
  // ∎ END OF PROOF
  [0x220e]: filledRect(150, 50, 300, 400),
  // ∏ N-ARY PRODUCT
  [0x220f]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 500);
      p.L(480, 500);
      p.stroke();
      p.M(170, 500);
      p.L(170, 0);
      p.stroke();
      p.M(430, 500);
      p.L(430, 0);
      p.stroke();
    }
  },
  // ∑ N-ARY SUMMATION
  [0x2211]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(480, 500);
      p.L(120, 500);
      p.L(300, 250);
      p.L(120, 0);
      p.L(480, 0);
      p.stroke();
    }
  },
  // − MINUS SIGN
  [0x2212]: {
    width: W,
    draw: p => {
      p.lineWidth(60);
      p.M(120, 250);
      p.L(480, 250);
      p.stroke();
    }
  },
  // ∓ MINUS-OR-PLUS SIGN
  [0x2213]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 350);
      p.L(480, 350);
      p.stroke();
      p.M(300, 200);
      p.L(300, 0);
      p.stroke();
      p.M(120, 100);
      p.L(480, 100);
      p.stroke();
    }
  },
  // ∗ ASTERISK OPERATOR
  [0x2217]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(300, 400);
      p.L(300, 100);
      p.stroke();
      p.M(140, 350);
      p.L(460, 150);
      p.stroke();
      p.M(140, 150);
      p.L(460, 350);
      p.stroke();
    }
  },
  // √ SQUARE ROOT
  [0x221a]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(60, 250);
      p.L(150, 200);
      p.L(250, 0);
      p.L(540, 500);
      p.stroke();
    }
  },
  // ∞ INFINITY
  [0x221e]: {
    width: 700,
    draw: p => {
      p.lineWidth(50);
      p.M(350, 250);
      p.C(350, 400, 500, 450, 550, 350);
      p.C(600, 250, 550, 100, 500, 100);
      p.C(450, 100, 350, 250, 350, 250);
      p.C(350, 250, 250, 400, 200, 400);
      p.C(150, 400, 100, 250, 150, 150);
      p.C(200, 50, 350, 100, 350, 250);
      p.stroke();
    }
  },
  // ∠ ANGLE
  [0x2220]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(480, 50);
      p.L(120, 50);
      p.L(380, 450);
      p.stroke();
    }
  },
  // ∧ LOGICAL AND
  [0x2227]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(100, 50);
      p.L(300, 450);
      p.L(500, 50);
      p.stroke();
    }
  },
  // ∨ LOGICAL OR
  [0x2228]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(100, 450);
      p.L(300, 50);
      p.L(500, 450);
      p.stroke();
    }
  },
  // ∩ INTERSECTION
  [0x2229]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 50);
      p.L(120, 300);
      p.C(120, 470, 250, 500, 300, 500);
      p.C(350, 500, 480, 470, 480, 300);
      p.L(480, 50);
      p.stroke();
    }
  },
  // ∪ UNION
  [0x222a]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 450);
      p.L(120, 200);
      p.C(120, 30, 250, 0, 300, 0);
      p.C(350, 0, 480, 30, 480, 200);
      p.L(480, 450);
      p.stroke();
    }
  },
  // ∫ INTEGRAL
  [0x222b]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(380, 530);
      p.C(350, 530, 300, 500, 300, 450);
      p.L(300, 50);
      p.C(300, 0, 250, -30, 220, -30);
      p.stroke();
    }
  },
  // ≈ ALMOST EQUAL TO
  [0x2248]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 350);
      p.C(200, 430, 350, 430, 480, 350);
      p.stroke();
      p.M(120, 200);
      p.C(200, 280, 350, 280, 480, 200);
      p.stroke();
    }
  },
  // ≠ NOT EQUAL TO
  [0x2260]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 330);
      p.L(480, 330);
      p.stroke();
      p.M(120, 170);
      p.L(480, 170);
      p.stroke();
      p.M(380, 450);
      p.L(220, 50);
      p.stroke();
    }
  },
  // ≡ IDENTICAL TO
  [0x2261]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 380);
      p.L(480, 380);
      p.stroke();
      p.M(120, 250);
      p.L(480, 250);
      p.stroke();
      p.M(120, 120);
      p.L(480, 120);
      p.stroke();
    }
  },
  // ≤ LESS-THAN OR EQUAL TO
  [0x2264]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(480, 420);
      p.L(120, 250);
      p.L(480, 80);
      p.stroke();
      p.M(120, 30);
      p.L(480, 30);
      p.stroke();
    }
  },
  // ≥ GREATER-THAN OR EQUAL TO
  [0x2265]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(120, 420);
      p.L(480, 250);
      p.L(120, 80);
      p.stroke();
      p.M(120, 30);
      p.L(480, 30);
      p.stroke();
    }
  },
  // ⊂ SUBSET OF
  [0x2282]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(470, 450);
      p.C(200, 450, 130, 350, 130, 250);
      p.C(130, 150, 200, 50, 470, 50);
      p.stroke();
    }
  },
  // ⊃ SUPERSET OF
  [0x2283]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(130, 450);
      p.C(400, 450, 470, 350, 470, 250);
      p.C(470, 150, 400, 50, 130, 50);
      p.stroke();
    }
  },
  // ⊕ CIRCLED PLUS
  [0x2295]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.circle(300, 250, 200);
      p.stroke();
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
    }
  },
  // ⊖ CIRCLED MINUS
  [0x2296]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.circle(300, 250, 200);
      p.stroke();
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // ⊗ CIRCLED TIMES
  [0x2297]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.circle(300, 250, 200);
      p.stroke();
      p.M(160, 110);
      p.L(440, 390);
      p.stroke();
      p.M(160, 390);
      p.L(440, 110);
      p.stroke();
    }
  },
  // ⊘ CIRCLED DIVISION SLASH
  [0x2298]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.circle(300, 250, 200);
      p.stroke();
      p.M(160, 100);
      p.L(440, 400);
      p.stroke();
    }
  },
  // ⊙ CIRCLED DOT OPERATOR
  [0x2299]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.circle(300, 250, 200);
      p.stroke();
      p.circle(300, 250, 40);
      p.fill();
    }
  }
};

// =============================================================================
// Dingbats / Checkmarks / Misc Symbols (U+2600–U+27BF)
// =============================================================================

const DING: Record<number, GlyphDef> = {
  // ☀ BLACK SUN WITH RAYS — simplified as circle with lines
  [0x2600]: {
    width: W,
    draw: p => {
      p.circle(300, 250, 130);
      p.fill();
      p.lineWidth(40);
      for (let a = 0; a < 360; a += 45) {
        const r1 = 160;
        const r2 = 240;
        const rad = (a * Math.PI) / 180;
        p.M(300 + r1 * Math.cos(rad), 250 + r1 * Math.sin(rad));
        p.L(300 + r2 * Math.cos(rad), 250 + r2 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // ★ BLACK STAR
  [0x2605]: {
    width: W,
    draw: p => {
      const cx = 300;
      const cy = 260;
      const R = 230;
      const r = 90;
      for (let i = 0; i < 5; i++) {
        const a1 = ((i * 72 - 90) * Math.PI) / 180;
        const a2 = ((i * 72 + 36 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        } else {
          p.L(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        }
        p.L(cx + r * Math.cos(a2), cy + r * Math.sin(a2));
      }
      p.Z();
      p.fill();
    }
  },
  // ☆ WHITE STAR
  [0x2606]: {
    width: W,
    draw: p => {
      p.lineWidth(35);
      const cx = 300;
      const cy = 260;
      const R = 230;
      const r = 90;
      for (let i = 0; i < 5; i++) {
        const a1 = ((i * 72 - 90) * Math.PI) / 180;
        const a2 = ((i * 72 + 36 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        } else {
          p.L(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        }
        p.L(cx + r * Math.cos(a2), cy + r * Math.sin(a2));
      }
      p.Z();
      p.stroke();
    }
  },
  // ☐ BALLOT BOX
  [0x2610]: strokedRect(100, 30, 400, 400, 45),
  // ☑ BALLOT BOX WITH CHECK
  [0x2611]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.rect(100, 30, 400, 400);
      p.stroke();
      p.lineWidth(55);
      p.M(180, 230);
      p.L(280, 130);
      p.L(430, 350);
      p.stroke();
    }
  },
  // ☒ BALLOT BOX WITH X
  [0x2612]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.rect(100, 30, 400, 400);
      p.stroke();
      p.lineWidth(50);
      p.M(170, 100);
      p.L(430, 360);
      p.stroke();
      p.M(170, 360);
      p.L(430, 100);
      p.stroke();
    }
  },
  // ☓ SALTIRE (X)
  [0x2613]: {
    width: W,
    draw: p => {
      p.lineWidth(55);
      p.M(120, 50);
      p.L(480, 450);
      p.stroke();
      p.M(120, 450);
      p.L(480, 50);
      p.stroke();
    }
  },
  // ☛ BLACK RIGHT POINTING INDEX — simplified arrow
  [0x261b]: {
    width: W,
    draw: p => {
      p.M(500, 250);
      p.L(300, 400);
      p.L(300, 300);
      p.L(100, 300);
      p.L(100, 200);
      p.L(300, 200);
      p.L(300, 100);
      p.Z();
      p.fill();
    }
  },
  // ☞ WHITE RIGHT POINTING INDEX
  [0x261e]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(300, 400);
      p.L(300, 300);
      p.L(100, 300);
      p.L(100, 200);
      p.L(300, 200);
      p.L(300, 100);
      p.Z();
      p.stroke();
    }
  },
  // ♀ FEMALE SIGN
  [0x2640]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.circle(300, 330, 150);
      p.stroke();
      p.M(300, 180);
      p.L(300, 0);
      p.stroke();
      p.M(220, 90);
      p.L(380, 90);
      p.stroke();
    }
  },
  // ♂ MALE SIGN
  [0x2642]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.circle(250, 220, 150);
      p.stroke();
      p.M(360, 330);
      p.L(500, 470);
      p.stroke();
      p.M(400, 470);
      p.L(500, 470);
      p.L(500, 370);
      p.stroke();
    }
  },
  // ♠ BLACK SPADE SUIT
  [0x2660]: {
    width: W,
    draw: p => {
      p.M(300, 500);
      p.C(300, 500, 80, 350, 80, 200);
      p.C(80, 100, 200, 50, 300, 180);
      p.C(400, 50, 520, 100, 520, 200);
      p.C(520, 350, 300, 500, 300, 500);
      p.Z();
      p.fill();
      p.M(250, 0);
      p.L(350, 0);
      p.L(320, 150);
      p.L(280, 150);
      p.Z();
      p.fill();
    }
  },
  // ♣ BLACK CLUB SUIT
  [0x2663]: {
    width: W,
    draw: p => {
      p.circle(300, 370, 110);
      p.fill();
      p.circle(190, 230, 100);
      p.fill();
      p.circle(410, 230, 100);
      p.fill();
      p.rect(270, 0, 60, 200);
      p.fill();
    }
  },
  // ♥ BLACK HEART SUIT
  [0x2665]: {
    width: W,
    draw: p => {
      p.M(300, 80);
      p.C(300, 80, 80, 200, 80, 340);
      p.C(80, 440, 180, 490, 300, 490);
      p.C(420, 490, 520, 440, 520, 340);
      p.C(520, 200, 300, 80, 300, 80);
      p.Z();
      p.fill();
    }
  },
  // ♦ BLACK DIAMOND SUIT
  [0x2666]: {
    width: W,
    draw: p => {
      p.M(300, 500);
      p.L(500, 250);
      p.L(300, 0);
      p.L(100, 250);
      p.Z();
      p.fill();
    }
  },
  // ♩ QUARTER NOTE
  [0x2669]: {
    width: W,
    draw: p => {
      p.ellipse(230, 80, 100, 70);
      p.fill();
      p.lineWidth(40);
      p.M(330, 80);
      p.L(330, 480);
      p.stroke();
    }
  },
  // ♪ EIGHTH NOTE
  [0x266a]: {
    width: W,
    draw: p => {
      p.ellipse(200, 80, 100, 70);
      p.fill();
      p.lineWidth(35);
      p.M(300, 80);
      p.L(300, 480);
      p.stroke();
      p.M(300, 480);
      p.C(350, 450, 430, 400, 430, 350);
      p.stroke();
    }
  },
  // ♭ MUSIC FLAT SIGN
  [0x266d]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(200, 0);
      p.L(200, 500);
      p.stroke();
      p.M(200, 200);
      p.C(350, 200, 400, 100, 200, 0);
      p.stroke();
    }
  },
  // ♯ MUSIC SHARP SIGN
  [0x266f]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(220, 0);
      p.L(220, 500);
      p.stroke();
      p.M(380, 0);
      p.L(380, 500);
      p.stroke();
      p.lineWidth(50);
      p.M(140, 180);
      p.L(460, 240);
      p.stroke();
      p.M(140, 310);
      p.L(460, 370);
      p.stroke();
    }
  },
  // ✓ CHECK MARK
  [0x2713]: {
    width: W,
    draw: p => {
      p.lineWidth(65);
      p.M(100, 250);
      p.L(240, 80);
      p.L(500, 450);
      p.stroke();
    }
  },
  // ✔ HEAVY CHECK MARK
  [0x2714]: {
    width: W,
    draw: p => {
      p.lineWidth(85);
      p.M(100, 250);
      p.L(240, 80);
      p.L(500, 450);
      p.stroke();
    }
  },
  // ✕ MULTIPLICATION X
  [0x2715]: {
    width: W,
    draw: p => {
      p.lineWidth(55);
      p.M(140, 70);
      p.L(460, 430);
      p.stroke();
      p.M(140, 430);
      p.L(460, 70);
      p.stroke();
    }
  },
  // ✖ HEAVY MULTIPLICATION X
  [0x2716]: {
    width: W,
    draw: p => {
      p.lineWidth(75);
      p.M(140, 70);
      p.L(460, 430);
      p.stroke();
      p.M(140, 430);
      p.L(460, 70);
      p.stroke();
    }
  },
  // ✗ BALLOT X
  [0x2717]: {
    width: W,
    draw: p => {
      p.lineWidth(60);
      p.M(130, 60);
      p.L(470, 440);
      p.stroke();
      p.M(130, 440);
      p.L(470, 60);
      p.stroke();
    }
  },
  // ✘ HEAVY BALLOT X
  [0x2718]: {
    width: W,
    draw: p => {
      p.lineWidth(80);
      p.M(130, 60);
      p.L(470, 440);
      p.stroke();
      p.M(130, 440);
      p.L(470, 60);
      p.stroke();
    }
  },
  // ✚ HEAVY GREEK CROSS
  [0x271a]: {
    width: W,
    draw: p => {
      p.rect(210, 30, 180, 440);
      p.fill();
      p.rect(90, 160, 420, 180);
      p.fill();
    }
  },
  // ✜ HEAVY OPEN CENTRE CROSS — simplified
  [0x271c]: {
    width: W,
    draw: p => {
      p.lineWidth(60);
      p.M(300, 30);
      p.L(300, 470);
      p.stroke();
      p.M(90, 250);
      p.L(510, 250);
      p.stroke();
    }
  },
  // ✠ MALTESE CROSS — simplified
  [0x2720]: {
    width: W,
    draw: p => {
      p.M(300, 480);
      p.L(250, 320);
      p.L(100, 350);
      p.L(200, 250);
      p.L(100, 150);
      p.L(250, 180);
      p.L(300, 20);
      p.L(350, 180);
      p.L(500, 150);
      p.L(400, 250);
      p.L(500, 350);
      p.L(350, 320);
      p.Z();
      p.fill();
    }
  },
  // ❤ HEAVY BLACK HEART
  [0x2764]: {
    width: W,
    draw: p => {
      p.M(300, 80);
      p.C(300, 80, 60, 220, 60, 350);
      p.C(60, 470, 180, 500, 300, 500);
      p.C(420, 500, 540, 470, 540, 350);
      p.C(540, 220, 300, 80, 300, 80);
      p.Z();
      p.fill();
    }
  },
  // ❶–❿ DINGBAT NEGATIVE CIRCLED (1–10)
  ...(() => {
    const r: Record<number, GlyphDef> = {};
    for (let i = 0; i < 10; i++) {
      const cp = 0x2776 + i;
      r[cp] = {
        width: W,
        draw: p => {
          p.circle(300, 250, 240);
          p.fill(); /* digit rendered as glyph outline – too complex for paths; filled circle is best approximation */
        }
      };
    }
    return r;
  })(),
  // ①–⑩ CIRCLED DIGIT (U+2460–U+2469)
  ...(() => {
    const r: Record<number, GlyphDef> = {};
    for (let i = 0; i < 10; i++) {
      r[0x2460 + i] = {
        width: W,
        draw: p => {
          p.lineWidth(40);
          p.circle(300, 250, 220);
          p.stroke();
        }
      };
    }
    return r;
  })()
};

// =============================================================================
// Miscellaneous Technical (U+2300 – U+23FF)
// =============================================================================

const TECH: Record<number, GlyphDef> = {
  // ⌀ DIAMETER SIGN
  [0x2300]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.circle(300, 250, 200);
      p.stroke();
      p.M(140, 90);
      p.L(460, 410);
      p.stroke();
    }
  },
  // ⌂ HOUSE
  [0x2302]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.M(300, 480);
      p.L(520, 250);
      p.L(520, 30);
      p.L(80, 30);
      p.L(80, 250);
      p.Z();
      p.stroke();
    }
  },
  // ⌘ PLACE OF INTEREST SIGN (command key) — simplified as looped square
  [0x2318]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.rect(180, 130, 240, 240);
      p.stroke();
      p.circle(180, 370, 70);
      p.stroke();
      p.circle(420, 370, 70);
      p.stroke();
      p.circle(180, 130, 70);
      p.stroke();
      p.circle(420, 130, 70);
      p.stroke();
    }
  },
  // ⌛ HOURGLASS
  [0x231b]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.M(120, 480);
      p.L(480, 480);
      p.stroke();
      p.M(120, 20);
      p.L(480, 20);
      p.stroke();
      p.M(150, 480);
      p.L(300, 250);
      p.L(450, 480);
      p.stroke();
      p.M(150, 20);
      p.L(300, 250);
      p.L(450, 20);
      p.stroke();
    }
  },
  // ⏎ RETURN SYMBOL
  [0x23ce]: {
    width: W,
    draw: p => {
      p.lineWidth(45);
      p.M(480, 420);
      p.L(480, 200);
      p.L(150, 200);
      p.stroke();
      p.M(150, 200);
      p.L(280, 310);
      p.L(280, 90);
      p.Z();
      p.fill();
    }
  }
};

// =============================================================================
// Box Drawing (U+2500 – U+257F)
// =============================================================================

const BOX: Record<number, GlyphDef> = {
  // ─ BOX DRAWINGS LIGHT HORIZONTAL
  [0x2500]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(0, 250);
      p.L(600, 250);
      p.stroke();
    }
  },
  // ━ BOX DRAWINGS HEAVY HORIZONTAL
  [0x2501]: {
    width: W,
    draw: p => {
      p.lineWidth(80);
      p.M(0, 250);
      p.L(600, 250);
      p.stroke();
    }
  },
  // │ BOX DRAWINGS LIGHT VERTICAL
  [0x2502]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
    }
  },
  // ┃ BOX DRAWINGS HEAVY VERTICAL
  [0x2503]: {
    width: W,
    draw: p => {
      p.lineWidth(80);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
    }
  },
  // ┌ BOX DRAWINGS LIGHT DOWN AND RIGHT
  [0x250c]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 0);
      p.L(300, 250);
      p.L(600, 250);
      p.stroke();
    }
  },
  // ┐ BOX DRAWINGS LIGHT DOWN AND LEFT
  [0x2510]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 0);
      p.L(300, 250);
      p.L(0, 250);
      p.stroke();
    }
  },
  // └ BOX DRAWINGS LIGHT UP AND RIGHT
  [0x2514]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 500);
      p.L(300, 250);
      p.L(600, 250);
      p.stroke();
    }
  },
  // ┘ BOX DRAWINGS LIGHT UP AND LEFT
  [0x2518]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 500);
      p.L(300, 250);
      p.L(0, 250);
      p.stroke();
    }
  },
  // ├ BOX DRAWINGS LIGHT VERTICAL AND RIGHT
  [0x251c]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(300, 250);
      p.L(600, 250);
      p.stroke();
    }
  },
  // ┤ BOX DRAWINGS LIGHT VERTICAL AND LEFT
  [0x2524]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(300, 250);
      p.L(0, 250);
      p.stroke();
    }
  },
  // ┬ BOX DRAWINGS LIGHT DOWN AND HORIZONTAL
  [0x252c]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(0, 250);
      p.L(600, 250);
      p.stroke();
      p.M(300, 250);
      p.L(300, 0);
      p.stroke();
    }
  },
  // ┴ BOX DRAWINGS LIGHT UP AND HORIZONTAL
  [0x2534]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(0, 250);
      p.L(600, 250);
      p.stroke();
      p.M(300, 250);
      p.L(300, 500);
      p.stroke();
    }
  },
  // ┼ BOX DRAWINGS LIGHT VERTICAL AND HORIZONTAL
  [0x253c]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(0, 250);
      p.L(600, 250);
      p.stroke();
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
    }
  },
  // ╔ BOX DRAWINGS DOUBLE DOWN AND RIGHT
  [0x2554]: {
    width: W,
    draw: p => {
      p.lineWidth(30);
      p.M(270, 0);
      p.L(270, 270);
      p.L(600, 270);
      p.stroke();
      p.M(330, 0);
      p.L(330, 230);
      p.L(600, 230);
      p.stroke();
    }
  },
  // ╗ BOX DRAWINGS DOUBLE DOWN AND LEFT
  [0x2557]: {
    width: W,
    draw: p => {
      p.lineWidth(30);
      p.M(330, 0);
      p.L(330, 270);
      p.L(0, 270);
      p.stroke();
      p.M(270, 0);
      p.L(270, 230);
      p.L(0, 230);
      p.stroke();
    }
  },
  // ╚ BOX DRAWINGS DOUBLE UP AND RIGHT
  [0x255a]: {
    width: W,
    draw: p => {
      p.lineWidth(30);
      p.M(270, 500);
      p.L(270, 230);
      p.L(600, 230);
      p.stroke();
      p.M(330, 500);
      p.L(330, 270);
      p.L(600, 270);
      p.stroke();
    }
  },
  // ╝ BOX DRAWINGS DOUBLE UP AND LEFT
  [0x255d]: {
    width: W,
    draw: p => {
      p.lineWidth(30);
      p.M(330, 500);
      p.L(330, 230);
      p.L(0, 230);
      p.stroke();
      p.M(270, 500);
      p.L(270, 270);
      p.L(0, 270);
      p.stroke();
    }
  },
  // ═ BOX DRAWINGS DOUBLE HORIZONTAL
  [0x2550]: {
    width: W,
    draw: p => {
      p.lineWidth(30);
      p.M(0, 270);
      p.L(600, 270);
      p.stroke();
      p.M(0, 230);
      p.L(600, 230);
      p.stroke();
    }
  },
  // ║ BOX DRAWINGS DOUBLE VERTICAL
  [0x2551]: {
    width: W,
    draw: p => {
      p.lineWidth(30);
      p.M(270, 0);
      p.L(270, 500);
      p.stroke();
      p.M(330, 0);
      p.L(330, 500);
      p.stroke();
    }
  }
};

// =============================================================================
// Block Elements (U+2580 – U+259F)
// =============================================================================

const BLOCK: Record<number, GlyphDef> = {
  // ▀ UPPER HALF BLOCK
  [0x2580]: filledRect(0, 250, 600, 250),
  // ▄ LOWER HALF BLOCK
  [0x2584]: filledRect(0, 0, 600, 250),
  // █ FULL BLOCK
  [0x2588]: filledRect(0, 0, 600, 500),
  // ▌ LEFT HALF BLOCK
  [0x258c]: filledRect(0, 0, 300, 500),
  // ▐ RIGHT HALF BLOCK
  [0x2590]: filledRect(300, 0, 300, 500),
  // ░ LIGHT SHADE — simplified as rect with lines
  [0x2591]: {
    width: W,
    draw: p => {
      p.lineWidth(15);
      p.rect(0, 0, 600, 500);
      p.stroke();
      for (let y = 0; y <= 500; y += 60) {
        p.M(0, y);
        p.L(600, y);
      }
      p.stroke();
    }
  },
  // ▒ MEDIUM SHADE
  [0x2592]: {
    width: W,
    draw: p => {
      p.lineWidth(25);
      p.rect(0, 0, 600, 500);
      p.stroke();
      for (let y = 0; y <= 500; y += 50) {
        p.M(0, y);
        p.L(600, y);
      }
      for (let x = 0; x <= 600; x += 60) {
        p.M(x, 0);
        p.L(x, 500);
      }
      p.stroke();
    }
  },
  // ▓ DARK SHADE
  [0x2593]: {
    width: W,
    draw: p => {
      p.lineWidth(35);
      p.rect(0, 0, 600, 500);
      p.fill();
    }
  }
};

// =============================================================================
// Misc Symbols & Arrows (U+2B00 – U+2BFF)
// =============================================================================

const MISC_ARROWS: Record<number, GlyphDef> = {
  // ⬆ UPWARDS BLACK ARROW
  [0x2b06]: {
    width: W,
    draw: p => {
      p.M(300, 500);
      p.L(100, 280);
      p.L(230, 280);
      p.L(230, 0);
      p.L(370, 0);
      p.L(370, 280);
      p.L(500, 280);
      p.Z();
      p.fill();
    }
  },
  // ⬇ DOWNWARDS BLACK ARROW
  [0x2b07]: {
    width: W,
    draw: p => {
      p.M(300, 0);
      p.L(100, 220);
      p.L(230, 220);
      p.L(230, 500);
      p.L(370, 500);
      p.L(370, 220);
      p.L(500, 220);
      p.Z();
      p.fill();
    }
  },
  // ⬅ LEFTWARDS BLACK ARROW
  [0x2b05]: {
    width: W,
    draw: p => {
      p.M(0, 250);
      p.L(220, 430);
      p.L(220, 320);
      p.L(550, 320);
      p.L(550, 180);
      p.L(220, 180);
      p.L(220, 70);
      p.Z();
      p.fill();
    }
  },
  // ➡ BLACK RIGHTWARDS ARROW (U+27A1 not here, but mapped)
  // ⬛ BLACK LARGE SQUARE
  [0x2b1b]: filledRect(60, 10, 480, 480),
  // ⬜ WHITE LARGE SQUARE
  [0x2b1c]: strokedRect(60, 10, 480, 480, 45),
  // ⭐ WHITE MEDIUM STAR — same as ☆
  [0x2b50]: {
    width: W,
    draw: p => {
      p.lineWidth(35);
      const cx = 300;
      const cy = 260;
      const R = 230;
      const r = 90;
      for (let i = 0; i < 5; i++) {
        const a1 = ((i * 72 - 90) * Math.PI) / 180;
        const a2 = ((i * 72 + 36 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        } else {
          p.L(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        }
        p.L(cx + r * Math.cos(a2), cy + r * Math.sin(a2));
      }
      p.Z();
      p.stroke();
    }
  }
};

// =============================================================================
// Squared / circled symbols from Misc Math Symbols-B (U+29C0–U+29FF)
// =============================================================================

const SQUARED: Record<number, GlyphDef> = {
  // ⧀ CIRCLED LESS-THAN
  [0x29c0]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.lineWidth(45);
      p.M(400, 380);
      p.L(200, 250);
      p.L(400, 120);
      p.stroke();
    }
  },
  // ⧁ CIRCLED GREATER-THAN
  [0x29c1]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.lineWidth(45);
      p.M(200, 380);
      p.L(400, 250);
      p.L(200, 120);
      p.stroke();
    }
  },
  // ⧇ SQUARED SMALL CIRCLE
  [0x29c7]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.rect(80, 30, 440, 440);
      p.stroke();
      p.circle(300, 250, 130);
      p.stroke();
    }
  },
  // ⧈ SQUARED SQUARE
  [0x29c8]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.rect(80, 30, 440, 440);
      p.stroke();
      p.rect(190, 140, 220, 220);
      p.stroke();
    }
  },
  // ⧉ TWO JOINED SQUARES
  [0x29c9]: {
    width: W,
    draw: p => {
      p.lineWidth(35);
      p.rect(80, 80, 300, 300);
      p.stroke();
      p.rect(220, 120, 300, 300);
      p.stroke();
    }
  }
};

// =============================================================================
// Currency Symbols (U+20A0 – U+20CF)
// =============================================================================

const CURRENCY: Record<number, GlyphDef> = {
  // ₹ INDIAN RUPEE SIGN
  [0x20b9]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(130, 480);
      p.L(470, 480);
      p.stroke();
      p.M(130, 380);
      p.L(470, 380);
      p.stroke();
      p.M(230, 480);
      p.C(350, 480, 420, 400, 420, 350);
      p.C(420, 280, 350, 200, 230, 200);
      p.stroke();
      p.M(200, 380);
      p.L(400, 0);
      p.stroke();
    }
  },
  // ₺ TURKISH LIRA SIGN
  [0x20ba]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(350, 480);
      p.L(350, 50);
      p.C(350, 0, 200, 0, 150, 50);
      p.stroke();
      p.M(150, 330);
      p.L(450, 250);
      p.stroke();
      p.M(150, 230);
      p.L(450, 150);
      p.stroke();
    }
  },
  // ₽ RUBLE SIGN
  [0x20bd]: {
    width: W,
    draw: p => {
      p.lineWidth(50);
      p.M(180, 0);
      p.L(180, 500);
      p.stroke();
      p.M(180, 500);
      p.L(350, 500);
      p.C(480, 500, 480, 350, 350, 300);
      p.L(180, 300);
      p.stroke();
      p.M(120, 200);
      p.L(400, 200);
      p.stroke();
    }
  }
};

// =============================================================================
// General Punctuation & Misc (assorted useful ones)
// =============================================================================

const PUNCT: Record<number, GlyphDef> = {
  // • BULLET (U+2022)
  [0x2022]: filledCircle(300, 250, 100),
  // … HORIZONTAL ELLIPSIS (U+2026)
  [0x2026]: {
    width: W,
    draw: p => {
      p.circle(120, 100, 40);
      p.fill();
      p.circle(300, 100, 40);
      p.fill();
      p.circle(480, 100, 40);
      p.fill();
    }
  },
  // ‣ TRIANGULAR BULLET (U+2023)
  [0x2023]: {
    width: W,
    draw: p => {
      p.M(150, 400);
      p.L(450, 250);
      p.L(150, 100);
      p.Z();
      p.fill();
    }
  },
  // ⁃ HYPHEN BULLET (U+2043)
  [0x2043]: {
    width: W,
    draw: p => {
      p.lineWidth(60);
      p.M(150, 250);
      p.L(450, 250);
      p.stroke();
    }
  },
  // ⁄ FRACTION SLASH (U+2044)
  [0x2044]: {
    width: W,
    draw: p => {
      p.lineWidth(40);
      p.M(450, 500);
      p.L(150, 0);
      p.stroke();
    }
  }
};

// =============================================================================
// Master lookup table
// =============================================================================

import {
  BOX_FULL,
  BLOCK_FULL,
  BRAILLE,
  LETTERLIKE,
  NUMBER_FORMS,
  ENCLOSED,
  PUNCT_EXT,
  ARROWS_EXT,
  MATH_EXT,
  MISC_EXT,
  DING_EXT,
  TECH_EXT,
  CURRENCY_EXT,
  MATH_SYM_A,
  SUP_ARROWS_A,
  GEOMETRIC_EXT,
  ROMAN_NUMERALS,
  ENCLOSED_EXT
} from "@pdf/font/type3-glyphs-extended";
import {
  ARROWS_HARPOONS,
  DINGBATS_FULL,
  MISC_SYMBOLS_FULL,
  MISC_TECHNICAL_EXT,
  SUP_ARROWS_B,
  MISC_MATH_A_EXT,
  MISC_SYM_ARROWS_EXT
} from "@pdf/font/type3-glyphs-extended2";
import {
  DINGBATS_FILL,
  MISC_SYM_FILL,
  MISC_SYM_ARR_FILL,
  SUP_ARROWS_FILL,
  GEN_PUNCT_FILL,
  NUM_FORMS_FILL,
  FINAL_FILL
} from "@pdf/font/type3-glyphs-fill";
import {
  SPACES,
  CIRCLED_DIGITS,
  CIRCLED_LETTERS,
  CIRCLED_SMALL_LETTERS,
  NEG_CIRCLED_DIGITS,
  REFINED_SYMBOLS,
  SUPP_MATH_OP,
  MISC_MATH_B,
  MATH_OP_FULL,
  LETTERLIKE_FULL,
  CURRENCY_REMAINING,
  ARROWS_REMAINING
} from "@pdf/font/type3-glyphs-quality";

const ALL_TABLES: Array<Record<number, GlyphDef>> = [
  GEO,
  ARR,
  MATH,
  DING,
  TECH,
  BOX,
  BLOCK,
  MISC_ARROWS,
  SQUARED,
  CURRENCY,
  PUNCT,
  // Extended tables
  BOX_FULL,
  BLOCK_FULL,
  BRAILLE,
  LETTERLIKE,
  NUMBER_FORMS,
  ENCLOSED,
  PUNCT_EXT,
  ARROWS_EXT,
  MATH_EXT,
  MISC_EXT,
  DING_EXT,
  TECH_EXT,
  CURRENCY_EXT,
  MATH_SYM_A,
  SUP_ARROWS_A,
  GEOMETRIC_EXT,
  ROMAN_NUMERALS,
  ENCLOSED_EXT,
  // Extended2 tables
  ARROWS_HARPOONS,
  DINGBATS_FULL,
  MISC_SYMBOLS_FULL,
  MISC_TECHNICAL_EXT,
  SUP_ARROWS_B,
  MISC_MATH_A_EXT,
  MISC_SYM_ARROWS_EXT,
  // Fill tables (gap coverage)
  DINGBATS_FILL,
  MISC_SYM_FILL,
  MISC_SYM_ARR_FILL,
  SUP_ARROWS_FILL,
  GEN_PUNCT_FILL,
  NUM_FORMS_FILL,
  FINAL_FILL,
  // Quality overrides + new coverage (must be LAST to override earlier entries)
  SPACES,
  SUPP_MATH_OP,
  MISC_MATH_B,
  MATH_OP_FULL,
  LETTERLIKE_FULL,
  CURRENCY_REMAINING,
  ARROWS_REMAINING,
  CIRCLED_DIGITS,
  CIRCLED_LETTERS,
  CIRCLED_SMALL_LETTERS,
  NEG_CIRCLED_DIGITS,
  REFINED_SYMBOLS
];

/** Merged lookup table (built once on first access). */
let _merged: Map<number, GlyphDef> | null = null;

function getMerged(): Map<number, GlyphDef> {
  if (!_merged) {
    _merged = new Map<number, GlyphDef>();
    for (const table of ALL_TABLES) {
      for (const [cp, def] of Object.entries(table)) {
        _merged.set(Number(cp), def);
      }
    }
  }
  return _merged;
}

/**
 * Look up a glyph definition for a Unicode code point.
 * Returns the glyph def, or `undefined` if no specific drawing is available.
 */
export function lookupGlyph(codePoint: number): GlyphDef | undefined {
  return getMerged().get(codePoint);
}

/**
 * Check whether a specific code point has a dedicated vector glyph.
 */
export function hasGlyph(codePoint: number): boolean {
  return getMerged().has(codePoint);
}

/**
 * Return the total number of defined glyphs (for diagnostics / testing).
 */
export function glyphCount(): number {
  return getMerged().size;
}
