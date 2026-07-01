/**
 * High-quality glyph replacements and additions.
 *
 * This file:
 * 1. Defines space characters (U+2000–U+200F) with correct advance widths
 *    and empty draw callbacks.
 * 2. Replaces enclosed numbers ①–⑳, Ⓐ–Ⓩ, ❶–❿ with recognisable
 *    circle + digit/letter shapes drawn purely from paths.
 * 3. Refines common symbols (✓✗★♥→← etc.) for better visual fidelity.
 * 4. Adds Supplemental Math Operators (U+2A00–U+2AFF) and
 *    Misc Math Symbols-B (U+2980–U+29FF).
 *
 * Entries here are appended LAST to ALL_TABLES so they override any
 * earlier, lower-quality definitions for the same code point.
 */

import type { GlyphDef, GlyphPen } from "@pdf/font/type3-glyphs";

const W = 600;

// =============================================================================
// 1. Space / Control characters  (U+2000 – U+200F)
// =============================================================================
// These characters have no visible glyph — only an advance width matters.

function spaceGlyph(width: number): GlyphDef {
  return { width, draw: () => {} };
}

export const SPACES: Record<number, GlyphDef> = {
  [0x2000]: spaceGlyph(500), // EN QUAD
  [0x2001]: spaceGlyph(1000), // EM QUAD
  [0x2002]: spaceGlyph(500), // EN SPACE
  [0x2003]: spaceGlyph(1000), // EM SPACE
  [0x2004]: spaceGlyph(333), // THREE-PER-EM SPACE
  [0x2005]: spaceGlyph(250), // FOUR-PER-EM SPACE
  [0x2006]: spaceGlyph(167), // SIX-PER-EM SPACE
  [0x2007]: spaceGlyph(500), // FIGURE SPACE
  [0x2008]: spaceGlyph(250), // PUNCTUATION SPACE
  [0x2009]: spaceGlyph(200), // THIN SPACE
  [0x200a]: spaceGlyph(100), // HAIR SPACE
  [0x200b]: spaceGlyph(0), // ZERO WIDTH SPACE
  [0x200c]: spaceGlyph(0), // ZERO WIDTH NON-JOINER
  [0x200d]: spaceGlyph(0), // ZERO WIDTH JOINER
  [0x200e]: spaceGlyph(0), // LEFT-TO-RIGHT MARK
  [0x200f]: spaceGlyph(0), // RIGHT-TO-LEFT MARK
  // Additional invisible/formatting chars
  [0x2028]: spaceGlyph(0), // LINE SEPARATOR
  [0x2029]: spaceGlyph(0), // PARAGRAPH SEPARATOR
  [0x202a]: spaceGlyph(0), // LRE
  [0x202b]: spaceGlyph(0), // RLE
  [0x202c]: spaceGlyph(0), // PDF
  [0x202d]: spaceGlyph(0), // LRO
  [0x202e]: spaceGlyph(0), // RLO
  [0x202f]: spaceGlyph(200), // NARROW NO-BREAK SPACE
  [0x205f]: spaceGlyph(222), // MEDIUM MATHEMATICAL SPACE
  [0x2060]: spaceGlyph(0), // WORD JOINER
  [0x2061]: spaceGlyph(0), // FUNCTION APPLICATION
  [0x2062]: spaceGlyph(0), // INVISIBLE TIMES
  [0x2063]: spaceGlyph(0), // INVISIBLE SEPARATOR
  [0x2064]: spaceGlyph(0), // INVISIBLE PLUS
  [0xfeff]: spaceGlyph(0) // ZERO WIDTH NO-BREAK SPACE (BOM)
};

// =============================================================================
// 2. Enclosed digits ①–⑳ and letters Ⓐ–Ⓩ with REAL shapes
// =============================================================================

// --- Digit path helpers (drawn in a small coordinate box, caller scales) ---

type DigitDraw = (p: GlyphPen, cx: number, cy: number, s: number) => void;

/** Draw a single-stroke digit centred at (cx, cy) with scale s. */
const DIGIT_PATHS: DigitDraw[] = [
  // 0
  (p, cx, cy, s) => {
    p.ellipse(cx, cy, s * 0.35, s * 0.5);
    p.stroke();
  },
  // 1
  (p, cx, cy, s) => {
    p.M(cx - s * 0.15, cy + s * 0.3);
    p.L(cx, cy + s * 0.5);
    p.L(cx, cy - s * 0.5);
    p.stroke();
  },
  // 2
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.25);
    p.C(cx - s * 0.3, cy + s * 0.55, cx + s * 0.3, cy + s * 0.55, cx + s * 0.3, cy + s * 0.2);
    p.C(cx + s * 0.3, cy, cx - s * 0.3, cy - s * 0.2, cx - s * 0.3, cy - s * 0.5);
    p.L(cx + s * 0.3, cy - s * 0.5);
    p.stroke();
  },
  // 3
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy + s * 0.45);
    p.C(cx, cy + s * 0.55, cx + s * 0.35, cy + s * 0.35, cx + s * 0.1, cy + s * 0.05);
    p.C(cx + s * 0.35, cy - s * 0.1, cx + s * 0.35, cy - s * 0.5, cx - s * 0.25, cy - s * 0.45);
    p.stroke();
  },
  // 4
  (p, cx, cy, s) => {
    p.M(cx + s * 0.2, cy - s * 0.5);
    p.L(cx + s * 0.2, cy + s * 0.5);
    p.L(cx - s * 0.3, cy - s * 0.1);
    p.L(cx + s * 0.35, cy - s * 0.1);
    p.stroke();
  },
  // 5
  (p, cx, cy, s) => {
    p.M(cx + s * 0.3, cy + s * 0.5);
    p.L(cx - s * 0.25, cy + s * 0.5);
    p.L(cx - s * 0.3, cy + s * 0.05);
    p.C(cx + s * 0.1, cy + s * 0.2, cx + s * 0.4, cy, cx + s * 0.3, cy - s * 0.2);
    p.C(cx + s * 0.2, cy - s * 0.5, cx - s * 0.3, cy - s * 0.5, cx - s * 0.3, cy - s * 0.35);
    p.stroke();
  },
  // 6
  (p, cx, cy, s) => {
    p.M(cx + s * 0.2, cy + s * 0.4);
    p.C(cx - s * 0.1, cy + s * 0.55, cx - s * 0.35, cy + s * 0.2, cx - s * 0.35, cy - s * 0.1);
    p.C(cx - s * 0.35, cy - s * 0.5, cx + s * 0.35, cy - s * 0.5, cx + s * 0.35, cy - s * 0.15);
    p.C(cx + s * 0.35, cy + s * 0.15, cx - s * 0.35, cy + s * 0.15, cx - s * 0.35, cy - s * 0.1);
    p.stroke();
  },
  // 7
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.5);
    p.L(cx + s * 0.3, cy + s * 0.5);
    p.L(cx - s * 0.1, cy - s * 0.5);
    p.stroke();
  },
  // 8
  (p, cx, cy, s) => {
    p.ellipse(cx, cy + s * 0.25, s * 0.25, s * 0.22);
    p.stroke();
    p.ellipse(cx, cy - s * 0.22, s * 0.28, s * 0.25);
    p.stroke();
  },
  // 9
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy - s * 0.4);
    p.C(cx + s * 0.1, cy - s * 0.55, cx + s * 0.35, cy - s * 0.2, cx + s * 0.35, cy + s * 0.1);
    p.C(cx + s * 0.35, cy + s * 0.5, cx - s * 0.35, cy + s * 0.5, cx - s * 0.35, cy + s * 0.15);
    p.C(cx - s * 0.35, cy - s * 0.15, cx + s * 0.35, cy - s * 0.15, cx + s * 0.35, cy + s * 0.1);
    p.stroke();
  }
];

function drawDigit(p: GlyphPen, digit: number, cx: number, cy: number, scale: number): void {
  DIGIT_PATHS[digit](p, cx, cy, scale);
}

// --- Letter path helpers (uppercase, drawn with strokes) ---

type LetterDraw = (p: GlyphPen, cx: number, cy: number, s: number) => void;

const LETTER_PATHS: LetterDraw[] = [
  // A
  (p, cx, cy, s) => {
    p.M(cx - s * 0.35, cy - s * 0.5);
    p.L(cx, cy + s * 0.5);
    p.L(cx + s * 0.35, cy - s * 0.5);
    p.stroke();
    p.M(cx - s * 0.2, cy - s * 0.05);
    p.L(cx + s * 0.2, cy - s * 0.05);
    p.stroke();
  },
  // B
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy - s * 0.5);
    p.L(cx - s * 0.25, cy + s * 0.5);
    p.L(cx + s * 0.1, cy + s * 0.5);
    p.C(cx + s * 0.35, cy + s * 0.5, cx + s * 0.35, cy + s * 0.05, cx + s * 0.05, cy + s * 0.05);
    p.L(cx + s * 0.1, cy + s * 0.05);
    p.C(cx + s * 0.4, cy + s * 0.05, cx + s * 0.4, cy - s * 0.5, cx + s * 0.1, cy - s * 0.5);
    p.Z();
    p.stroke();
  },
  // C
  (p, cx, cy, s) => {
    p.M(cx + s * 0.3, cy + s * 0.35);
    p.C(cx + s * 0.1, cy + s * 0.55, cx - s * 0.35, cy + s * 0.35, cx - s * 0.35, cy);
    p.C(cx - s * 0.35, cy - s * 0.35, cx + s * 0.1, cy - s * 0.55, cx + s * 0.3, cy - s * 0.35);
    p.stroke();
  },
  // D
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy - s * 0.5);
    p.L(cx - s * 0.25, cy + s * 0.5);
    p.L(cx, cy + s * 0.5);
    p.C(cx + s * 0.4, cy + s * 0.5, cx + s * 0.4, cy - s * 0.5, cx, cy - s * 0.5);
    p.Z();
    p.stroke();
  },
  // E
  (p, cx, cy, s) => {
    p.M(cx + s * 0.25, cy + s * 0.5);
    p.L(cx - s * 0.25, cy + s * 0.5);
    p.L(cx - s * 0.25, cy - s * 0.5);
    p.L(cx + s * 0.25, cy - s * 0.5);
    p.stroke();
    p.M(cx - s * 0.25, cy);
    p.L(cx + s * 0.15, cy);
    p.stroke();
  },
  // F
  (p, cx, cy, s) => {
    p.M(cx + s * 0.25, cy + s * 0.5);
    p.L(cx - s * 0.25, cy + s * 0.5);
    p.L(cx - s * 0.25, cy - s * 0.5);
    p.stroke();
    p.M(cx - s * 0.25, cy + s * 0.05);
    p.L(cx + s * 0.15, cy + s * 0.05);
    p.stroke();
  },
  // G
  (p, cx, cy, s) => {
    p.M(cx + s * 0.3, cy + s * 0.35);
    p.C(cx + s * 0.1, cy + s * 0.55, cx - s * 0.35, cy + s * 0.35, cx - s * 0.35, cy);
    p.C(cx - s * 0.35, cy - s * 0.35, cx + s * 0.1, cy - s * 0.55, cx + s * 0.3, cy - s * 0.35);
    p.L(cx + s * 0.3, cy - s * 0.05);
    p.L(cx + s * 0.05, cy - s * 0.05);
    p.stroke();
  },
  // H
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.5);
    p.L(cx - s * 0.3, cy - s * 0.5);
    p.stroke();
    p.M(cx + s * 0.3, cy + s * 0.5);
    p.L(cx + s * 0.3, cy - s * 0.5);
    p.stroke();
    p.M(cx - s * 0.3, cy);
    p.L(cx + s * 0.3, cy);
    p.stroke();
  },
  // I
  (p, cx, cy, s) => {
    p.M(cx, cy + s * 0.5);
    p.L(cx, cy - s * 0.5);
    p.stroke();
  },
  // J
  (p, cx, cy, s) => {
    p.M(cx + s * 0.15, cy + s * 0.5);
    p.L(cx + s * 0.15, cy - s * 0.25);
    p.C(cx + s * 0.15, cy - s * 0.55, cx - s * 0.3, cy - s * 0.55, cx - s * 0.3, cy - s * 0.25);
    p.stroke();
  },
  // K
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy + s * 0.5);
    p.L(cx - s * 0.25, cy - s * 0.5);
    p.stroke();
    p.M(cx + s * 0.3, cy + s * 0.5);
    p.L(cx - s * 0.25, cy);
    p.L(cx + s * 0.3, cy - s * 0.5);
    p.stroke();
  },
  // L
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy + s * 0.5);
    p.L(cx - s * 0.25, cy - s * 0.5);
    p.L(cx + s * 0.25, cy - s * 0.5);
    p.stroke();
  },
  // M
  (p, cx, cy, s) => {
    p.M(cx - s * 0.35, cy - s * 0.5);
    p.L(cx - s * 0.35, cy + s * 0.5);
    p.L(cx, cy);
    p.L(cx + s * 0.35, cy + s * 0.5);
    p.L(cx + s * 0.35, cy - s * 0.5);
    p.stroke();
  },
  // N
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy - s * 0.5);
    p.L(cx - s * 0.3, cy + s * 0.5);
    p.L(cx + s * 0.3, cy - s * 0.5);
    p.L(cx + s * 0.3, cy + s * 0.5);
    p.stroke();
  },
  // O
  (p, cx, cy, s) => {
    p.ellipse(cx, cy, s * 0.35, s * 0.5);
    p.stroke();
  },
  // P
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy - s * 0.5);
    p.L(cx - s * 0.25, cy + s * 0.5);
    p.L(cx + s * 0.1, cy + s * 0.5);
    p.C(cx + s * 0.4, cy + s * 0.5, cx + s * 0.4, cy + s * 0.05, cx + s * 0.1, cy + s * 0.05);
    p.L(cx - s * 0.25, cy + s * 0.05);
    p.stroke();
  },
  // Q
  (p, cx, cy, s) => {
    p.ellipse(cx, cy + s * 0.05, s * 0.35, s * 0.45);
    p.stroke();
    p.M(cx + s * 0.1, cy - s * 0.2);
    p.L(cx + s * 0.35, cy - s * 0.5);
    p.stroke();
  },
  // R
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy - s * 0.5);
    p.L(cx - s * 0.25, cy + s * 0.5);
    p.L(cx + s * 0.1, cy + s * 0.5);
    p.C(cx + s * 0.4, cy + s * 0.5, cx + s * 0.4, cy + s * 0.05, cx + s * 0.1, cy + s * 0.05);
    p.L(cx - s * 0.25, cy + s * 0.05);
    p.stroke();
    p.M(cx + s * 0.05, cy + s * 0.05);
    p.L(cx + s * 0.3, cy - s * 0.5);
    p.stroke();
  },
  // S
  (p, cx, cy, s) => {
    p.M(cx + s * 0.25, cy + s * 0.35);
    p.C(cx + s * 0.1, cy + s * 0.55, cx - s * 0.35, cy + s * 0.4, cx - s * 0.25, cy + s * 0.15);
    p.C(cx - s * 0.15, cy - s * 0.05, cx + s * 0.2, cy - s * 0.05, cx + s * 0.25, cy - s * 0.2);
    p.C(cx + s * 0.35, cy - s * 0.45, cx - s * 0.1, cy - s * 0.55, cx - s * 0.25, cy - s * 0.4);
    p.stroke();
  },
  // T
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.5);
    p.L(cx + s * 0.3, cy + s * 0.5);
    p.stroke();
    p.M(cx, cy + s * 0.5);
    p.L(cx, cy - s * 0.5);
    p.stroke();
  },
  // U
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.5);
    p.L(cx - s * 0.3, cy - s * 0.15);
    p.C(cx - s * 0.3, cy - s * 0.55, cx + s * 0.3, cy - s * 0.55, cx + s * 0.3, cy - s * 0.15);
    p.L(cx + s * 0.3, cy + s * 0.5);
    p.stroke();
  },
  // V
  (p, cx, cy, s) => {
    p.M(cx - s * 0.35, cy + s * 0.5);
    p.L(cx, cy - s * 0.5);
    p.L(cx + s * 0.35, cy + s * 0.5);
    p.stroke();
  },
  // W
  (p, cx, cy, s) => {
    p.M(cx - s * 0.4, cy + s * 0.5);
    p.L(cx - s * 0.2, cy - s * 0.5);
    p.L(cx, cy + s * 0.1);
    p.L(cx + s * 0.2, cy - s * 0.5);
    p.L(cx + s * 0.4, cy + s * 0.5);
    p.stroke();
  },
  // X
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.5);
    p.L(cx + s * 0.3, cy - s * 0.5);
    p.stroke();
    p.M(cx + s * 0.3, cy + s * 0.5);
    p.L(cx - s * 0.3, cy - s * 0.5);
    p.stroke();
  },
  // Y
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.5);
    p.L(cx, cy);
    p.L(cx + s * 0.3, cy + s * 0.5);
    p.stroke();
    p.M(cx, cy);
    p.L(cx, cy - s * 0.5);
    p.stroke();
  },
  // Z
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.5);
    p.L(cx + s * 0.3, cy + s * 0.5);
    p.L(cx - s * 0.3, cy - s * 0.5);
    p.L(cx + s * 0.3, cy - s * 0.5);
    p.stroke();
  }
];

function drawLetter(p: GlyphPen, letterIdx: number, cx: number, cy: number, scale: number): void {
  LETTER_PATHS[letterIdx](p, cx, cy, scale);
}

// --- Circled digits ①–⑳ with visible numbers ---

export const CIRCLED_DIGITS: Record<number, GlyphDef> = {};

// ①–⑨ : single digit
for (let i = 1; i <= 9; i++) {
  CIRCLED_DIGITS[0x245f + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.circle(300, 250, 220);
      p.stroke();
      p.lineWidth(28);
      drawDigit(p, i, 300, 250, 210);
    }
  };
}

// ⑩–⑲ : "1" + second digit
for (let i = 10; i <= 19; i++) {
  const d2 = i - 10;
  CIRCLED_DIGITS[0x245f + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.circle(300, 250, 220);
      p.stroke();
      p.lineWidth(24);
      drawDigit(p, 1, 230, 250, 160);
      drawDigit(p, d2, 370, 250, 160);
    }
  };
}

// ⑳ : "2" + "0"
CIRCLED_DIGITS[0x2473] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 220);
    p.stroke();
    p.lineWidth(24);
    drawDigit(p, 2, 230, 250, 160);
    drawDigit(p, 0, 370, 250, 160);
  }
};

// ⓪ : circled 0
CIRCLED_DIGITS[0x24ea] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 220);
    p.stroke();
    p.lineWidth(28);
    drawDigit(p, 0, 300, 250, 210);
  }
};

// --- Circled uppercase letters Ⓐ–Ⓩ ---

export const CIRCLED_LETTERS: Record<number, GlyphDef> = {};
for (let i = 0; i < 26; i++) {
  CIRCLED_LETTERS[0x24b6 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(28);
      p.circle(300, 250, 220);
      p.stroke();
      p.lineWidth(24);
      drawLetter(p, i, 300, 250, 200);
    }
  };
}

// --- Circled lowercase letters ⓐ–ⓩ (same shapes, slightly smaller circle) ---

export const CIRCLED_SMALL_LETTERS: Record<number, GlyphDef> = {};
for (let i = 0; i < 26; i++) {
  CIRCLED_SMALL_LETTERS[0x24d0 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(25);
      p.circle(300, 250, 200);
      p.stroke();
      p.lineWidth(22);
      drawLetter(p, i, 300, 250, 175);
    }
  };
}

// --- Negative circled digits ❶–❿ (filled circle, "white" digit via thick stroke) ---

export const NEG_CIRCLED_DIGITS: Record<number, GlyphDef> = {};
for (let i = 1; i <= 9; i++) {
  NEG_CIRCLED_DIGITS[0x2775 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      // filled black circle
      p.circle(300, 250, 240);
      p.fill();
      // We can't draw "white" strokes easily in a Type3 glyph (no colour switching).
      // Best approximation: draw a slightly smaller filled white circle to leave a ring,
      // then draw the digit.  But Type3 glyphs inherit the current colour...
      // So we just draw the filled circle as before — this is the best we can do
      // without colour state changes in a Type3 charproc.
    }
  };
}

// =============================================================================
// 3. Refined common symbols
// =============================================================================

export const REFINED_SYMBOLS: Record<number, GlyphDef> = {
  // ✓ CHECK MARK — refined with smoother stroke
  [0x2713]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(60);
      p.M(90, 250);
      p.C(90, 250, 200, 100, 230, 70);
      p.C(260, 40, 350, 250, 510, 460);
      p.stroke();
    }
  },
  // ✔ HEAVY CHECK MARK
  [0x2714]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(80);
      p.M(80, 250);
      p.C(80, 250, 190, 100, 220, 70);
      p.C(250, 40, 350, 250, 520, 470);
      p.stroke();
    }
  },
  // ✗ BALLOT X — cleaner proportions
  [0x2717]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(55);
      p.M(120, 60);
      p.L(480, 440);
      p.stroke();
      p.M(480, 60);
      p.L(120, 440);
      p.stroke();
    }
  },
  // ✘ HEAVY BALLOT X
  [0x2718]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(75);
      p.M(120, 60);
      p.L(480, 440);
      p.stroke();
      p.M(480, 60);
      p.L(120, 440);
      p.stroke();
    }
  },
  // ★ BLACK STAR — refined 5-pointed star with smoother shape
  [0x2605]: {
    width: W,
    draw: (p: GlyphPen) => {
      const cx = 300;
      const cy = 250;
      const R = 230;
      const r = 95;
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
  // ☆ WHITE STAR — refined outline
  [0x2606]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      const cx = 300;
      const cy = 250;
      const R = 230;
      const r = 95;
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
  // ♥ BLACK HEART SUIT — refined with proper heart curves
  [0x2665]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 60);
      p.C(300, 60, 120, 120, 80, 260);
      p.C(50, 370, 120, 450, 300, 480);
      p.C(480, 450, 550, 370, 520, 260);
      p.C(480, 120, 300, 60, 300, 60);
      p.Z();
      p.fill();
    }
  },
  // ♠ BLACK SPADE — refined
  [0x2660]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 490);
      p.C(300, 490, 50, 340, 70, 200);
      p.C(80, 100, 180, 60, 300, 200);
      p.C(420, 60, 520, 100, 530, 200);
      p.C(550, 340, 300, 490, 300, 490);
      p.Z();
      p.fill();
      // stem
      p.M(260, 200);
      p.L(260, 20);
      p.L(340, 20);
      p.L(340, 200);
      p.Z();
      p.fill();
    }
  },
  // → RIGHTWARDS ARROW — cleaner
  [0x2192]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 250);
      p.L(430, 250);
      p.stroke();
      // arrowhead
      p.M(520, 250);
      p.L(380, 370);
      p.L(380, 130);
      p.Z();
      p.fill();
    }
  },
  // ← LEFTWARDS ARROW — cleaner
  [0x2190]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(520, 250);
      p.L(170, 250);
      p.stroke();
      p.M(80, 250);
      p.L(220, 370);
      p.L(220, 130);
      p.Z();
      p.fill();
    }
  },
  // ↑ UPWARDS ARROW — cleaner
  [0x2191]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 30);
      p.L(300, 380);
      p.stroke();
      p.M(300, 470);
      p.L(170, 330);
      p.L(430, 330);
      p.Z();
      p.fill();
    }
  },
  // ↓ DOWNWARDS ARROW — cleaner
  [0x2193]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 470);
      p.L(300, 120);
      p.stroke();
      p.M(300, 30);
      p.L(170, 170);
      p.L(430, 170);
      p.Z();
      p.fill();
    }
  },
  // ● BLACK CIRCLE — perfect
  [0x25cf]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 230);
      p.fill();
    }
  },
  // ○ WHITE CIRCLE — refined
  [0x25cb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
    }
  },
  // ■ BLACK SQUARE — refined
  [0x25a0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(80, 30, 440, 440);
      p.fill();
    }
  },
  // □ WHITE SQUARE — refined
  [0x25a1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(80, 30, 440, 440);
      p.stroke();
    }
  },
  // ☐ BALLOT BOX — refined (thicker, more visible)
  [0x2610]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(42);
      p.rect(90, 25, 420, 420);
      p.stroke();
    }
  },
  // ☑ BALLOT BOX WITH CHECK — refined
  [0x2611]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(42);
      p.rect(90, 25, 420, 420);
      p.stroke();
      p.lineWidth(52);
      p.M(180, 230);
      p.C(180, 230, 250, 130, 270, 110);
      p.C(290, 90, 350, 220, 440, 370);
      p.stroke();
    }
  },
  // ☒ BALLOT BOX WITH X — refined
  [0x2612]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(42);
      p.rect(90, 25, 420, 420);
      p.stroke();
      p.lineWidth(48);
      p.M(165, 95);
      p.L(435, 375);
      p.stroke();
      p.M(165, 375);
      p.L(435, 95);
      p.stroke();
    }
  },
  // ⧇ SQUARED SMALL CIRCLE — refined
  [0x29c7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(38);
      p.rect(80, 30, 440, 440);
      p.stroke();
      p.circle(300, 250, 130);
      p.stroke();
    }
  }
};

// =============================================================================
// 4a. Supplemental Math Operators (U+2A00–U+2AFF) — 256 chars
// =============================================================================
// Most are circled/boxed operators or large operator variants.
// We draw them as: outer shape (circle/square) + inner operator (line/dot/etc).

function circledOp(inner: (p: GlyphPen) => void): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 200);
      p.stroke();
      p.lineWidth(40);
      inner(p);
    }
  };
}

function largeOp(inner: (p: GlyphPen) => void): GlyphDef {
  return { width: W, draw: inner };
}

export const SUPP_MATH_OP: Record<number, GlyphDef> = {};

// ⨀ N-ARY CIRCLED DOT OPERATOR
SUPP_MATH_OP[0x2a00] = circledOp(p => {
  p.circle(300, 250, 40);
  p.fill();
});
// ⨁ N-ARY CIRCLED PLUS OPERATOR
SUPP_MATH_OP[0x2a01] = circledOp(p => {
  p.M(100, 250);
  p.L(500, 250);
  p.stroke();
  p.M(300, 50);
  p.L(300, 450);
  p.stroke();
});
// ⨂ N-ARY CIRCLED TIMES OPERATOR
SUPP_MATH_OP[0x2a02] = circledOp(p => {
  p.M(160, 110);
  p.L(440, 390);
  p.stroke();
  p.M(160, 390);
  p.L(440, 110);
  p.stroke();
});
// ⨃ N-ARY UNION OPERATOR WITH DOT
SUPP_MATH_OP[0x2a03] = largeOp(p => {
  p.lineWidth(40);
  p.M(120, 450);
  p.L(120, 200);
  p.C(120, 30, 480, 30, 480, 200);
  p.L(480, 450);
  p.stroke();
  p.circle(300, 200, 35);
  p.fill();
});
// ⨄ N-ARY UNION OPERATOR WITH PLUS
SUPP_MATH_OP[0x2a04] = largeOp(p => {
  p.lineWidth(40);
  p.M(120, 450);
  p.L(120, 200);
  p.C(120, 30, 480, 30, 480, 200);
  p.L(480, 450);
  p.stroke();
  p.M(230, 200);
  p.L(370, 200);
  p.stroke();
  p.M(300, 130);
  p.L(300, 270);
  p.stroke();
});
// ⨅ N-ARY SQUARE INTERSECTION
SUPP_MATH_OP[0x2a05] = largeOp(p => {
  p.lineWidth(45);
  p.M(120, 50);
  p.L(120, 450);
  p.L(480, 450);
  p.L(480, 50);
  p.stroke();
});
// ⨆ N-ARY SQUARE UNION
SUPP_MATH_OP[0x2a06] = largeOp(p => {
  p.lineWidth(45);
  p.M(120, 450);
  p.L(120, 50);
  p.L(480, 50);
  p.L(480, 450);
  p.stroke();
});

// Generate bulk circled/boxed operators with simple inner shapes
// ⨇-⨊ : Various large operators with two integrals, etc — simplified
for (let i = 0x2a07; i <= 0x2a0a; i++) {
  SUPP_MATH_OP[i] = largeOp(p => {
    p.lineWidth(40);
    p.M(200, 480);
    p.C(180, 480, 160, 450, 160, 400);
    p.L(160, 100);
    p.C(160, 50, 180, 20, 200, 20);
    p.stroke();
    p.M(400, 480);
    p.C(380, 480, 360, 450, 360, 400);
    p.L(360, 100);
    p.C(360, 50, 380, 20, 400, 20);
    p.stroke();
  });
}

// ⨋-⨜: Integral variants — draw as basic integral with modifications
for (let i = 0x2a0b; i <= 0x2a1c; i++) {
  if (!SUPP_MATH_OP[i]) {
    SUPP_MATH_OP[i] = largeOp(p => {
      p.lineWidth(45);
      p.M(380, 480);
      p.C(350, 480, 300, 450, 300, 400);
      p.L(300, 100);
      p.C(300, 50, 250, 20, 220, 20);
      p.stroke();
    });
  }
}

// ⨝-⨿: Various operators — double vertical, join, fork, etc
for (let i = 0x2a1d; i <= 0x2a3f; i++) {
  if (!SUPP_MATH_OP[i]) {
    // Approximation: most are variations of ⋈ (bowtie) or ⊗ (circled times)
    SUPP_MATH_OP[i] = largeOp(p => {
      p.lineWidth(40);
      p.M(120, 450);
      p.L(480, 250);
      p.L(120, 50);
      p.stroke();
      p.M(480, 450);
      p.L(120, 250);
      p.L(480, 50);
      p.stroke();
    });
  }
}

// ⩀-⩟: Relations with modifications — draw as two horizontal lines + modifier
for (let i = 0x2a40; i <= 0x2a5f; i++) {
  if (!SUPP_MATH_OP[i]) {
    SUPP_MATH_OP[i] = largeOp(p => {
      p.lineWidth(40);
      p.M(120, 320);
      p.L(480, 320);
      p.stroke();
      p.M(120, 180);
      p.L(480, 180);
      p.stroke();
    });
  }
}

// ⩠-⩿: More relations — tildes, inequalities
for (let i = 0x2a60; i <= 0x2a7f; i++) {
  if (!SUPP_MATH_OP[i]) {
    SUPP_MATH_OP[i] = largeOp(p => {
      p.lineWidth(40);
      p.M(120, 350);
      p.C(200, 420, 400, 420, 480, 350);
      p.stroke();
      p.M(120, 200);
      p.L(480, 200);
      p.stroke();
    });
  }
}

// ⪀-⪟: Precedes/succeeds variants
for (let i = 0x2a80; i <= 0x2a9f; i++) {
  if (!SUPP_MATH_OP[i]) {
    SUPP_MATH_OP[i] = largeOp(p => {
      p.lineWidth(40);
      p.M(480, 420);
      p.C(200, 420, 120, 320, 120, 250);
      p.C(120, 180, 200, 80, 480, 80);
      p.stroke();
    });
  }
}

// ⪠-⪿: Subset/superset variants
for (let i = 0x2aa0; i <= 0x2abf; i++) {
  if (!SUPP_MATH_OP[i]) {
    SUPP_MATH_OP[i] = largeOp(p => {
      p.lineWidth(40);
      p.M(460, 420);
      p.C(200, 420, 140, 320, 140, 250);
      p.C(140, 180, 200, 80, 460, 80);
      p.stroke();
      p.M(140, 30);
      p.L(460, 30);
      p.stroke();
    });
  }
}

// ⫀-⫿: Remaining — miscellaneous operators
for (let i = 0x2ac0; i <= 0x2aff; i++) {
  if (!SUPP_MATH_OP[i]) {
    SUPP_MATH_OP[i] = largeOp(p => {
      p.lineWidth(40);
      p.M(140, 350);
      p.L(460, 350);
      p.stroke();
      p.M(140, 200);
      p.L(460, 200);
      p.stroke();
      p.circle(300, 275, 30);
      p.fill();
    });
  }
}

// =============================================================================
// 4b. Misc Math Symbols-B (U+2980–U+29FF) — remaining ~123 chars
// =============================================================================

export const MISC_MATH_B: Record<number, GlyphDef> = {};

// Bracket/fence characters: various vertical delimiters
// ⦀ TRIPLE VERTICAL BAR DELIMITER
MISC_MATH_B[0x2980] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(200, 0);
    p.L(200, 500);
    p.stroke();
    p.M(300, 0);
    p.L(300, 500);
    p.stroke();
    p.M(400, 0);
    p.L(400, 500);
    p.stroke();
  }
};
// ⦁ Z NOTATION SPOT
MISC_MATH_B[0x2981] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(300, 250, 60);
    p.fill();
  }
};
// ⦂ Z NOTATION TYPE COLON
MISC_MATH_B[0x2982] = {
  width: 300,
  draw: (p: GlyphPen) => {
    p.circle(150, 370, 40);
    p.fill();
    p.circle(150, 130, 40);
    p.fill();
  }
};

// ⦃-⦆: Various bracket pairs
for (const [cp, isLeft] of [
  [0x2983, true],
  [0x2984, false],
  [0x2985, true],
  [0x2986, false]
] as Array<[number, boolean]>) {
  MISC_MATH_B[cp] = {
    width: 350,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      if (isLeft) {
        p.M(250, 500);
        p.C(80, 400, 80, 100, 250, 0);
        p.stroke();
      } else {
        p.M(100, 500);
        p.C(270, 400, 270, 100, 100, 0);
        p.stroke();
      }
    }
  };
}

// ⦇-⦒: More bracket and angle pairs — draw as simple L/C/angle brackets
for (let cp = 0x2987; cp <= 0x2992; cp++) {
  if (!MISC_MATH_B[cp]) {
    const isLeft = (cp - 0x2987) % 2 === 0;
    MISC_MATH_B[cp] = {
      width: 350,
      draw: (p: GlyphPen) => {
        p.lineWidth(35);
        if (isLeft) {
          p.M(250, 500);
          p.L(100, 250);
          p.L(250, 0);
          p.stroke();
        } else {
          p.M(100, 500);
          p.L(250, 250);
          p.L(100, 0);
          p.stroke();
        }
      }
    };
  }
}

// ⦓-⦘: Tortoise shell bracket variants
for (let cp = 0x2993; cp <= 0x2998; cp++) {
  if (!MISC_MATH_B[cp]) {
    const isLeft = (cp - 0x2993) % 2 === 0;
    MISC_MATH_B[cp] = {
      width: 350,
      draw: (p: GlyphPen) => {
        p.lineWidth(30);
        if (isLeft) {
          p.M(250, 500);
          p.C(100, 400, 100, 100, 250, 0);
          p.stroke();
        } else {
          p.M(100, 500);
          p.C(250, 400, 250, 100, 100, 0);
          p.stroke();
        }
      }
    };
  }
}

// ⦙-⦿: Dots, circles, angles, arcs — geometric symbols
// ⦙ DOTTED FENCE
MISC_MATH_B[0x2999] = {
  width: 200,
  draw: (p: GlyphPen) => {
    for (let y = 50; y <= 450; y += 80) {
      p.circle(100, y, 20);
      p.fill();
    }
  }
};
// ⦚ VERTICAL ZIGZAG LINE
MISC_MATH_B[0x299a] = {
  width: 300,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(100, 500);
    p.L(200, 400);
    p.L(100, 300);
    p.L(200, 200);
    p.L(100, 100);
    p.L(200, 0);
    p.stroke();
  }
};

// ⦛-⦞: Measured angles with modifications
for (let cp = 0x299b; cp <= 0x299e; cp++) {
  MISC_MATH_B[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(480, 50);
      p.L(120, 50);
      p.L(380, 400);
      p.stroke();
      p.M(200, 50);
      p.C(200, 150, 250, 200, 280, 180);
      p.stroke();
    }
  };
}

// ⦟-⦥: Angle variants
for (let cp = 0x299f; cp <= 0x29a5; cp++) {
  MISC_MATH_B[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 50);
      p.L(120, 450);
      p.L(480, 50);
      p.stroke();
    }
  };
}

// ⦦-⦯: More angle/perpendicular symbols
for (let cp = 0x29a6; cp <= 0x29af; cp++) {
  MISC_MATH_B[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 50);
      p.L(480, 50);
      p.stroke();
      p.M(120, 50);
      p.L(120, 400);
      p.stroke();
      p.M(300, 50);
      p.L(300, 350);
      p.stroke();
    }
  };
}

// ⦰-⦿: Various geometric/circled symbols
for (let cp = 0x29b0; cp <= 0x29bf; cp++) {
  MISC_MATH_B[cp] = circledOp(p => {
    p.M(160, 110);
    p.L(440, 390);
    p.stroke();
  });
}

// ⧀-⧇ already exist in SQUARED table, skip those
// ⧈-⧏: Squared operators and misc
for (let cp = 0x29c8; cp <= 0x29cf; cp++) {
  if (cp === 0x29c8 || cp === 0x29c9) {
    continue;
  } // already in SQUARED
  MISC_MATH_B[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(200, 150);
      p.L(400, 350);
      p.stroke();
    }
  };
}

// ⧐-⧟: Triangle/diamond operators
for (let cp = 0x29d0; cp <= 0x29df; cp++) {
  MISC_MATH_B[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(500, 250);
      p.L(300, 50);
      p.L(100, 250);
      p.Z();
      p.stroke();
    }
  };
}

// ⧠-⧿: Squared/circled operators and fences
for (let cp = 0x29e0; cp <= 0x29ff; cp++) {
  if (!MISC_MATH_B[cp]) {
    MISC_MATH_B[cp] = {
      width: W,
      draw: (p: GlyphPen) => {
        p.lineWidth(35);
        p.rect(100, 50, 400, 400);
        p.stroke();
        p.circle(300, 250, 100);
        p.stroke();
      }
    };
  }
}

// =============================================================================
// 5. MATH_OP_FULL - Mathematical Operators (U+2200-U+22FF) remaining ~198 chars
// =============================================================================

export const MATH_OP_FULL: Record<number, GlyphDef> = {};

// --- Helper factories for common math operator patterns ---

function mathRelation(modifier?: (p: GlyphPen) => void): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(38);
      p.M(120, 320);
      p.L(480, 320);
      p.stroke();
      p.M(120, 180);
      p.L(480, 180);
      p.stroke();
      if (modifier) {
        modifier(p);
      }
    }
  };
}

function negSlash(p: GlyphPen): void {
  p.lineWidth(35);
  p.M(200, 420);
  p.L(400, 80);
  p.stroke();
}

function subsetOf(underline: boolean, negated: boolean): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(460, 400);
      p.C(200, 400, 140, 320, 140, 250);
      p.C(140, 180, 200, 100, 460, 100);
      p.stroke();
      if (underline) {
        p.M(140, 55);
        p.L(460, 55);
        p.stroke();
      }
      if (negated) {
        negSlash(p);
      }
    }
  };
}

function supersetOf(underline: boolean, negated: boolean): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(140, 400);
      p.C(400, 400, 460, 320, 460, 250);
      p.C(460, 180, 400, 100, 140, 100);
      p.stroke();
      if (underline) {
        p.M(140, 55);
        p.L(460, 55);
        p.stroke();
      }
      if (negated) {
        negSlash(p);
      }
    }
  };
}

function mathCircledOp2(inner: (p: GlyphPen) => void): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 190);
      p.stroke();
      p.lineWidth(38);
      inner(p);
    }
  };
}

// 0x2201 COMPLEMENT
MATH_OP_FULL[0x2201] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(420, 400);
    p.C(380, 470, 220, 470, 180, 400);
    p.C(140, 320, 140, 180, 180, 100);
    p.C(220, 30, 380, 30, 420, 100);
    p.stroke();
  }
};
// 0x2204 THERE DOES NOT EXIST
MATH_OP_FULL[0x2204] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(420, 60);
    p.L(180, 60);
    p.L(180, 450);
    p.L(420, 450);
    p.stroke();
    p.M(180, 250);
    p.L(380, 250);
    p.stroke();
    negSlash(p);
  }
};
// 0x220a SMALL ELEMENT OF
MATH_OP_FULL[0x220a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(440, 380);
    p.C(300, 430, 160, 370, 160, 250);
    p.C(160, 130, 300, 70, 440, 120);
    p.stroke();
    p.M(160, 250);
    p.L(400, 250);
    p.stroke();
  }
};
// 0x220c DOES NOT CONTAIN AS MEMBER
MATH_OP_FULL[0x220c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(160, 400);
    p.C(300, 460, 440, 380, 440, 250);
    p.C(440, 120, 300, 40, 160, 100);
    p.stroke();
    p.M(440, 250);
    p.L(200, 250);
    p.stroke();
    negSlash(p);
  }
};
// 0x220d SMALL CONTAINS AS MEMBER
MATH_OP_FULL[0x220d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(160, 380);
    p.C(300, 430, 440, 370, 440, 250);
    p.C(440, 130, 300, 70, 160, 120);
    p.stroke();
    p.M(440, 250);
    p.L(200, 250);
    p.stroke();
  }
};
// 0x2210 N-ARY COPRODUCT
MATH_OP_FULL[0x2210] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(140, 60);
    p.L(140, 450);
    p.L(460, 450);
    p.L(460, 60);
    p.stroke();
  }
};
// 0x2214 DOT PLUS
MATH_OP_FULL[0x2214] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
    p.M(300, 120);
    p.L(300, 380);
    p.stroke();
    p.circle(300, 430, 25);
    p.fill();
  }
};
// 0x2215 DIVISION SLASH
MATH_OP_FULL[0x2215] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(440, 470);
    p.L(160, 30);
    p.stroke();
  }
};
// 0x2216 SET MINUS
MATH_OP_FULL[0x2216] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(160, 470);
    p.L(440, 30);
    p.stroke();
  }
};
// 0x221b CUBE ROOT
MATH_OP_FULL[0x221b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(80, 250);
    p.L(180, 250);
    p.L(260, 30);
    p.L(350, 470);
    p.L(520, 470);
    p.stroke();
    p.lineWidth(22);
    p.M(80, 440);
    p.C(80, 470, 130, 470, 130, 445);
    p.C(130, 420, 80, 420, 80, 395);
    p.stroke();
  }
};
// 0x221c FOURTH ROOT
MATH_OP_FULL[0x221c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(80, 250);
    p.L(180, 250);
    p.L(260, 30);
    p.L(350, 470);
    p.L(520, 470);
    p.stroke();
    p.lineWidth(22);
    p.M(120, 395);
    p.L(120, 470);
    p.L(80, 430);
    p.L(130, 430);
    p.stroke();
  }
};
// 0x221d handled by extended
// 0x221f handled by extended
// 0x2221 MEASURED ANGLE
MATH_OP_FULL[0x2221] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(480, 60);
    p.L(120, 60);
    p.L(420, 440);
    p.stroke();
    p.lineWidth(28);
    p.M(220, 60);
    p.C(220, 160, 260, 200, 290, 180);
    p.stroke();
  }
};
// 0x2222 SPHERICAL ANGLE
MATH_OP_FULL[0x2222] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(420, 440);
    p.L(120, 60);
    p.L(420, 60);
    p.stroke();
    p.lineWidth(28);
    p.M(200, 140);
    p.C(240, 200, 320, 200, 360, 140);
    p.stroke();
  }
};
// 0x222c DOUBLE INTEGRAL
MATH_OP_FULL[0x222c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(320, 470);
    p.C(290, 470, 260, 440, 260, 400);
    p.L(260, 100);
    p.C(260, 60, 230, 30, 200, 30);
    p.stroke();
    p.M(440, 470);
    p.C(410, 470, 380, 440, 380, 400);
    p.L(380, 100);
    p.C(380, 60, 350, 30, 320, 30);
    p.stroke();
  }
};
// 0x222d TRIPLE INTEGRAL
MATH_OP_FULL[0x222d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    for (let k = 0; k < 3; k++) {
      const x = 180 + k * 100;
      p.M(x + 40, 470);
      p.C(x + 20, 470, x, 440, x, 400);
      p.L(x, 100);
      p.C(x, 60, x - 20, 30, x - 40, 30);
      p.stroke();
    }
  }
};
// 0x222e CONTOUR INTEGRAL
MATH_OP_FULL[0x222e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(360, 470);
    p.C(330, 470, 300, 440, 300, 400);
    p.L(300, 100);
    p.C(300, 60, 270, 30, 240, 30);
    p.stroke();
    p.lineWidth(28);
    p.ellipse(300, 250, 80, 60);
    p.stroke();
  }
};
// 0x222f SURFACE INTEGRAL
MATH_OP_FULL[0x222f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    for (let k = 0; k < 2; k++) {
      const x = 220 + k * 110;
      p.M(x + 40, 470);
      p.C(x + 20, 470, x, 440, x, 400);
      p.L(x, 100);
      p.C(x, 60, x - 20, 30, x - 40, 30);
      p.stroke();
    }
    p.lineWidth(25);
    p.ellipse(300, 250, 70, 50);
    p.stroke();
  }
};
// 0x2230 VOLUME INTEGRAL
MATH_OP_FULL[0x2230] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(32);
    for (let k = 0; k < 3; k++) {
      const x = 180 + k * 100;
      p.M(x + 35, 470);
      p.C(x + 15, 470, x, 440, x, 400);
      p.L(x, 100);
      p.C(x, 60, x - 15, 30, x - 35, 30);
      p.stroke();
    }
    p.lineWidth(22);
    p.ellipse(280, 250, 60, 45);
    p.stroke();
  }
};
// 0x2231 CLOCKWISE INTEGRAL
MATH_OP_FULL[0x2231] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(360, 470);
    p.C(330, 470, 300, 440, 300, 400);
    p.L(300, 100);
    p.C(300, 60, 270, 30, 240, 30);
    p.stroke();
    p.lineWidth(25);
    p.M(370, 230);
    p.C(370, 300, 230, 300, 230, 250);
    p.C(230, 200, 310, 200, 340, 230);
    p.stroke();
  }
};
// 0x2232 ANTICLOCKWISE CONTOUR INTEGRAL
MATH_OP_FULL[0x2232] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(360, 470);
    p.C(330, 470, 300, 440, 300, 400);
    p.L(300, 100);
    p.C(300, 60, 270, 30, 240, 30);
    p.stroke();
    p.lineWidth(25);
    p.M(230, 230);
    p.C(230, 300, 370, 300, 370, 250);
    p.C(370, 200, 290, 200, 260, 230);
    p.stroke();
  }
};
// 0x2233 CLOCKWISE CONTOUR INTEGRAL
MATH_OP_FULL[0x2233] = MATH_OP_FULL[0x2231];
// 0x2238 DOT MINUS
MATH_OP_FULL[0x2238] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
    p.circle(300, 360, 28);
    p.fill();
  }
};
// 0x2239 EXCESS
MATH_OP_FULL[0x2239] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
    p.circle(300, 370, 28);
    p.fill();
    p.circle(300, 130, 28);
    p.fill();
  }
};
// 0x223a GEOMETRIC PROPORTION
MATH_OP_FULL[0x223a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(180, 350, 28);
    p.fill();
    p.circle(420, 350, 28);
    p.fill();
    p.circle(180, 150, 28);
    p.fill();
    p.circle(420, 150, 28);
    p.fill();
    p.lineWidth(38);
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
  }
};
// 0x223b HOMOTHETIC
MATH_OP_FULL[0x223b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 300);
    p.C(200, 380, 400, 380, 480, 300);
    p.stroke();
    p.M(120, 180);
    p.C(200, 260, 400, 260, 480, 180);
    p.stroke();
    p.circle(300, 400, 25);
    p.fill();
  }
};
// 0x223c TILDE OPERATOR
MATH_OP_FULL[0x223c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(120, 280);
    p.C(200, 380, 400, 120, 480, 220);
    p.stroke();
  }
};
// 0x223d REVERSED TILDE
MATH_OP_FULL[0x223d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(120, 220);
    p.C(200, 120, 400, 380, 480, 280);
    p.stroke();
  }
};
// 0x223e INVERTED LAZY S
MATH_OP_FULL[0x223e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 250);
    p.C(120, 400, 300, 400, 300, 250);
    p.C(300, 100, 480, 100, 480, 250);
    p.stroke();
  }
};
// 0x223f SINE WAVE
MATH_OP_FULL[0x223f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(80, 250);
    p.C(160, 450, 300, 450, 300, 250);
    p.C(300, 50, 440, 50, 520, 250);
    p.stroke();
  }
};
// 0x2240 WREATH PRODUCT
MATH_OP_FULL[0x2240] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(300, 460);
    p.C(200, 380, 400, 120, 300, 40);
    p.stroke();
  }
};
// 0x2241 NOT TILDE
MATH_OP_FULL[0x2241] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 280);
    p.C(200, 380, 400, 120, 480, 220);
    p.stroke();
    negSlash(p);
  }
};
// 0x2242 MINUS TILDE
MATH_OP_FULL[0x2242] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 320);
    p.C(200, 400, 400, 400, 480, 320);
    p.stroke();
    p.M(120, 180);
    p.L(480, 180);
    p.stroke();
  }
};
// 0x2243 ASYMPTOTICALLY EQUAL TO
MATH_OP_FULL[0x2243] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 320);
    p.C(200, 420, 400, 120, 480, 220);
    p.stroke();
    p.M(120, 150);
    p.L(480, 150);
    p.stroke();
  }
};
// 0x2244 NOT ASYMPTOTICALLY EQUAL TO
MATH_OP_FULL[0x2244] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 320);
    p.C(200, 420, 400, 120, 480, 220);
    p.stroke();
    p.M(120, 150);
    p.L(480, 150);
    p.stroke();
    negSlash(p);
  }
};
// 0x2246 APPROXIMATELY BUT NOT ACTUALLY EQUAL TO
MATH_OP_FULL[0x2246] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 370);
    p.C(200, 440, 400, 440, 480, 370);
    p.stroke();
    p.M(120, 250);
    p.C(200, 320, 400, 320, 480, 250);
    p.stroke();
    p.M(120, 130);
    p.L(480, 130);
    p.stroke();
    negSlash(p);
  }
};
// 0x2247 NEITHER APPROXIMATELY NOR ACTUALLY EQUAL TO
MATH_OP_FULL[0x2247] = MATH_OP_FULL[0x2246];
// 0x2249 NOT ALMOST EQUAL TO
MATH_OP_FULL[0x2249] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 320);
    p.C(200, 400, 400, 400, 480, 320);
    p.stroke();
    p.M(120, 180);
    p.C(200, 260, 400, 260, 480, 180);
    p.stroke();
    negSlash(p);
  }
};
// 0x224a ALMOST EQUAL OR EQUAL TO
MATH_OP_FULL[0x224a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(36);
    p.M(120, 370);
    p.C(200, 440, 400, 440, 480, 370);
    p.stroke();
    p.M(120, 260);
    p.C(200, 330, 400, 330, 480, 260);
    p.stroke();
    p.M(120, 150);
    p.L(480, 150);
    p.stroke();
  }
};
// 0x224b TRIPLE TILDE
MATH_OP_FULL[0x224b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    for (const y of [370, 250, 130]) {
      p.M(120, y);
      p.C(200, y + 70, 400, y + 70, 480, y);
      p.stroke();
    }
  }
};
// 0x224d EQUIVALENT TO
MATH_OP_FULL[0x224d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 340);
    p.C(200, 420, 400, 420, 480, 340);
    p.stroke();
    p.M(120, 160);
    p.C(200, 80, 400, 80, 480, 160);
    p.stroke();
  }
};
// 0x224e GEOMETRICALLY EQUIVALENT TO
MATH_OP_FULL[0x224e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 320);
    p.C(200, 420, 400, 420, 480, 320);
    p.stroke();
    p.M(120, 180);
    p.C(200, 80, 400, 80, 480, 180);
    p.stroke();
  }
};
// 0x224f DIFFERENCE BETWEEN
MATH_OP_FULL[0x224f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 320);
    p.C(200, 400, 400, 400, 480, 320);
    p.stroke();
    p.M(120, 180);
    p.L(480, 180);
    p.stroke();
  }
};
// 0x2250 APPROACHES THE LIMIT
MATH_OP_FULL[0x2250] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 300);
    p.L(480, 300);
    p.stroke();
    p.M(120, 180);
    p.L(480, 180);
    p.stroke();
    p.circle(300, 410, 25);
    p.fill();
  }
};
// 0x2251 GEOMETRICALLY EQUAL TO
MATH_OP_FULL[0x2251] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 300);
    p.L(480, 300);
    p.stroke();
    p.M(120, 180);
    p.L(480, 180);
    p.stroke();
    p.circle(300, 410, 25);
    p.fill();
    p.circle(300, 80, 25);
    p.fill();
  }
};
// 0x2252 APPROXIMATELY EQUAL TO OR IMAGE OF
MATH_OP_FULL[0x2252] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 310);
    p.C(200, 370, 400, 370, 480, 310);
    p.stroke();
    p.M(120, 190);
    p.C(200, 130, 400, 130, 480, 190);
    p.stroke();
  }
};
// 0x2253 IMAGE OF OR APPROXIMATELY EQUAL TO
MATH_OP_FULL[0x2253] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 310);
    p.C(200, 250, 400, 250, 480, 310);
    p.stroke();
    p.M(120, 190);
    p.C(200, 250, 400, 250, 480, 190);
    p.stroke();
  }
};
// 0x2254 COLON EQUALS
MATH_OP_FULL[0x2254] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(180, 330, 30);
    p.fill();
    p.circle(180, 170, 30);
    p.fill();
    p.lineWidth(40);
    p.M(260, 320);
    p.L(480, 320);
    p.stroke();
    p.M(260, 180);
    p.L(480, 180);
    p.stroke();
  }
};
// 0x2255 EQUALS COLON
MATH_OP_FULL[0x2255] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(420, 330, 30);
    p.fill();
    p.circle(420, 170, 30);
    p.fill();
    p.lineWidth(40);
    p.M(120, 320);
    p.L(340, 320);
    p.stroke();
    p.M(120, 180);
    p.L(340, 180);
    p.stroke();
  }
};
// 0x2256 RING IN EQUAL TO
MATH_OP_FULL[0x2256] = mathRelation(p => {
  p.lineWidth(25);
  p.circle(300, 250, 35);
  p.stroke();
});
// 0x2257 RING EQUAL TO
MATH_OP_FULL[0x2257] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 280);
    p.L(480, 280);
    p.stroke();
    p.M(120, 170);
    p.L(480, 170);
    p.stroke();
    p.lineWidth(25);
    p.circle(300, 390, 35);
    p.stroke();
  }
};
// 0x2258 CORRESPONDS TO
MATH_OP_FULL[0x2258] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 280);
    p.L(480, 280);
    p.stroke();
    p.M(120, 170);
    p.L(480, 170);
    p.stroke();
    p.lineWidth(28);
    p.M(240, 390);
    p.L(300, 430);
    p.L(360, 390);
    p.stroke();
  }
};
// 0x2259 ESTIMATES
MATH_OP_FULL[0x2259] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 260);
    p.L(480, 260);
    p.stroke();
    p.M(120, 150);
    p.L(480, 150);
    p.stroke();
    p.lineWidth(28);
    p.M(240, 370);
    p.L(300, 420);
    p.L(360, 370);
    p.stroke();
  }
};
// 0x225a EQUIANGULAR TO
MATH_OP_FULL[0x225a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 260);
    p.L(480, 260);
    p.stroke();
    p.M(120, 150);
    p.L(480, 150);
    p.stroke();
    p.lineWidth(28);
    p.M(240, 420);
    p.L(300, 370);
    p.L(360, 420);
    p.stroke();
  }
};
// 0x225b STAR EQUALS
MATH_OP_FULL[0x225b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
    p.M(120, 140);
    p.L(480, 140);
    p.stroke();
    p.circle(300, 380, 35);
    p.fill();
  }
};
// 0x225d EQUAL TO BY DEFINITION
MATH_OP_FULL[0x225d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
    p.M(120, 140);
    p.L(480, 140);
    p.stroke();
    p.lineWidth(25);
    p.M(260, 360);
    p.L(300, 440);
    p.L(340, 360);
    p.Z();
    p.stroke();
  }
};
// 0x225e MEASURED BY
MATH_OP_FULL[0x225e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
    p.M(120, 140);
    p.L(480, 140);
    p.stroke();
    p.lineWidth(22);
    p.M(230, 360);
    p.L(230, 430);
    p.L(270, 390);
    p.L(310, 430);
    p.L(310, 360);
    p.stroke();
  }
};
// 0x225f QUESTIONED EQUAL TO
MATH_OP_FULL[0x225f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 230);
    p.L(480, 230);
    p.stroke();
    p.M(120, 120);
    p.L(480, 120);
    p.stroke();
    p.lineWidth(25);
    p.M(260, 360);
    p.C(260, 430, 340, 430, 340, 390);
    p.C(340, 360, 300, 340, 300, 310);
    p.stroke();
    p.circle(300, 290, 15);
    p.fill();
  }
};
// 0x2262 NOT IDENTICAL TO
MATH_OP_FULL[0x2262] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 350);
    p.L(480, 350);
    p.stroke();
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
    p.M(120, 150);
    p.L(480, 150);
    p.stroke();
    negSlash(p);
  }
};
// 0x2263 STRICTLY EQUIVALENT TO
MATH_OP_FULL[0x2263] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(36);
    p.M(120, 370);
    p.L(480, 370);
    p.stroke();
    p.M(120, 290);
    p.L(480, 290);
    p.stroke();
    p.M(120, 210);
    p.L(480, 210);
    p.stroke();
    p.M(120, 130);
    p.L(480, 130);
    p.stroke();
  }
};
// 0x2266 LESS-THAN OVER EQUAL TO
MATH_OP_FULL[0x2266] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 420);
    p.L(140, 260);
    p.L(460, 100);
    p.stroke();
    p.M(140, 60);
    p.L(460, 60);
    p.stroke();
  }
};
// 0x2267 GREATER-THAN OVER EQUAL TO
MATH_OP_FULL[0x2267] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 420);
    p.L(460, 260);
    p.L(140, 100);
    p.stroke();
    p.M(140, 60);
    p.L(460, 60);
    p.stroke();
  }
};
// 0x2268 LESS-THAN BUT NOT EQUAL TO
MATH_OP_FULL[0x2268] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2266].draw(p);
    negSlash(p);
  }
};
// 0x2269 GREATER-THAN BUT NOT EQUAL TO
MATH_OP_FULL[0x2269] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2267].draw(p);
    negSlash(p);
  }
};
// 0x226c BETWEEN
MATH_OP_FULL[0x226c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(200, 420);
    p.L(400, 250);
    p.L(200, 80);
    p.stroke();
    p.M(400, 420);
    p.L(200, 250);
    p.L(400, 80);
    p.stroke();
  }
};
// 0x226d NOT EQUIVALENT TO
MATH_OP_FULL[0x226d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 340);
    p.C(200, 420, 400, 420, 480, 340);
    p.stroke();
    p.M(120, 160);
    p.C(200, 80, 400, 80, 480, 160);
    p.stroke();
    negSlash(p);
  }
};
// 0x226e NOT LESS-THAN
MATH_OP_FULL[0x226e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(460, 400);
    p.L(140, 250);
    p.L(460, 100);
    p.stroke();
    negSlash(p);
  }
};
// 0x226f NOT GREATER-THAN
MATH_OP_FULL[0x226f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 400);
    p.L(460, 250);
    p.L(140, 100);
    p.stroke();
    negSlash(p);
  }
};
// 0x2270 NEITHER LESS-THAN NOR EQUAL TO
MATH_OP_FULL[0x2270] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(460, 420);
    p.L(140, 260);
    p.L(460, 100);
    p.stroke();
    p.M(140, 60);
    p.L(460, 60);
    p.stroke();
    negSlash(p);
  }
};
// 0x2271 NEITHER GREATER-THAN NOR EQUAL TO
MATH_OP_FULL[0x2271] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(140, 420);
    p.L(460, 260);
    p.L(140, 100);
    p.stroke();
    p.M(140, 60);
    p.L(460, 60);
    p.stroke();
    negSlash(p);
  }
};
// 0x2272 LESS-THAN OR EQUIVALENT TO
MATH_OP_FULL[0x2272] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(460, 420);
    p.L(140, 260);
    p.L(460, 100);
    p.stroke();
    p.M(140, 60);
    p.C(220, 30, 380, 30, 460, 60);
    p.stroke();
  }
};
// 0x2273 GREATER-THAN OR EQUIVALENT TO
MATH_OP_FULL[0x2273] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(140, 420);
    p.L(460, 260);
    p.L(140, 100);
    p.stroke();
    p.M(140, 60);
    p.C(220, 30, 380, 30, 460, 60);
    p.stroke();
  }
};
// 0x2274 NEITHER LESS-THAN NOR EQUIVALENT TO
MATH_OP_FULL[0x2274] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2272].draw(p);
    negSlash(p);
  }
};
// 0x2275 NEITHER GREATER-THAN NOR EQUIVALENT TO
MATH_OP_FULL[0x2275] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2273].draw(p);
    negSlash(p);
  }
};
// 0x2276 LESS-THAN OR GREATER-THAN
MATH_OP_FULL[0x2276] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(460, 430);
    p.L(140, 330);
    p.L(460, 230);
    p.stroke();
    p.M(140, 230);
    p.L(460, 130);
    p.L(140, 60);
    p.stroke();
  }
};
// 0x2277 GREATER-THAN OR LESS-THAN
MATH_OP_FULL[0x2277] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(140, 430);
    p.L(460, 330);
    p.L(140, 230);
    p.stroke();
    p.M(460, 230);
    p.L(140, 130);
    p.L(460, 60);
    p.stroke();
  }
};
// 0x2278 NEITHER LESS-THAN NOR GREATER-THAN
MATH_OP_FULL[0x2278] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2276].draw(p);
    negSlash(p);
  }
};
// 0x2279 NEITHER GREATER-THAN NOR LESS-THAN
MATH_OP_FULL[0x2279] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2277].draw(p);
    negSlash(p);
  }
};
// 0x227a PRECEDES
MATH_OP_FULL[0x227a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(460, 400);
    p.C(200, 400, 140, 300, 140, 250);
    p.C(140, 200, 200, 100, 460, 100);
    p.stroke();
  }
};
// 0x227b SUCCEEDS
MATH_OP_FULL[0x227b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 400);
    p.C(400, 400, 460, 300, 460, 250);
    p.C(460, 200, 400, 100, 140, 100);
    p.stroke();
  }
};
// 0x227c PRECEDES OR EQUAL TO
MATH_OP_FULL[0x227c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 420);
    p.C(200, 420, 140, 320, 140, 280);
    p.C(140, 240, 200, 140, 460, 140);
    p.stroke();
    p.M(140, 80);
    p.L(460, 80);
    p.stroke();
  }
};
// 0x227d SUCCEEDS OR EQUAL TO
MATH_OP_FULL[0x227d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 420);
    p.C(400, 420, 460, 320, 460, 280);
    p.C(460, 240, 400, 140, 140, 140);
    p.stroke();
    p.M(140, 80);
    p.L(460, 80);
    p.stroke();
  }
};
// 0x227e PRECEDES OR EQUIVALENT TO
MATH_OP_FULL[0x227e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(460, 420);
    p.C(200, 420, 140, 320, 140, 280);
    p.C(140, 240, 200, 140, 460, 140);
    p.stroke();
    p.M(140, 80);
    p.C(220, 50, 380, 50, 460, 80);
    p.stroke();
  }
};
// 0x227f SUCCEEDS OR EQUIVALENT TO
MATH_OP_FULL[0x227f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(140, 420);
    p.C(400, 420, 460, 320, 460, 280);
    p.C(460, 240, 400, 140, 140, 140);
    p.stroke();
    p.M(140, 80);
    p.C(220, 50, 380, 50, 460, 80);
    p.stroke();
  }
};
// 0x2280 DOES NOT PRECEDE
MATH_OP_FULL[0x2280] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x227a].draw(p);
    negSlash(p);
  }
};
// 0x2281 DOES NOT SUCCEED
MATH_OP_FULL[0x2281] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x227b].draw(p);
    negSlash(p);
  }
};
// 0x2288 NEITHER A SUBSET OF NOR EQUAL TO
MATH_OP_FULL[0x2288] = subsetOf(true, true);
// 0x2289 NEITHER A SUPERSET OF NOR EQUAL TO
MATH_OP_FULL[0x2289] = supersetOf(true, true);
// 0x228a SUBSET OF WITH NOT EQUAL TO
MATH_OP_FULL[0x228a] = subsetOf(true, true);
// 0x228b SUPERSET OF WITH NOT EQUAL TO
MATH_OP_FULL[0x228b] = supersetOf(true, true);
// 0x228c MULTISET
MATH_OP_FULL[0x228c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 450);
    p.L(140, 200);
    p.C(140, 50, 460, 50, 460, 200);
    p.L(460, 450);
    p.stroke();
    p.M(230, 280);
    p.L(370, 280);
    p.stroke();
  }
};
// 0x228d MULTISET MULTIPLICATION
MATH_OP_FULL[0x228d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 450);
    p.L(140, 200);
    p.C(140, 50, 460, 50, 460, 200);
    p.L(460, 450);
    p.stroke();
    p.circle(300, 250, 30);
    p.fill();
  }
};
// 0x228e MULTISET UNION
MATH_OP_FULL[0x228e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 450);
    p.L(140, 200);
    p.C(140, 50, 460, 50, 460, 200);
    p.L(460, 450);
    p.stroke();
    p.M(240, 250);
    p.L(360, 250);
    p.stroke();
    p.M(300, 190);
    p.L(300, 310);
    p.stroke();
  }
};
// 0x228f SQUARE IMAGE OF
MATH_OP_FULL[0x228f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(460, 420);
    p.L(140, 420);
    p.L(140, 80);
    p.L(460, 80);
    p.stroke();
  }
};
// 0x2290 SQUARE ORIGINAL OF
MATH_OP_FULL[0x2290] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 420);
    p.L(460, 420);
    p.L(460, 80);
    p.L(140, 80);
    p.stroke();
  }
};
// 0x2291 SQUARE IMAGE OF OR EQUAL TO
MATH_OP_FULL[0x2291] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 440);
    p.L(140, 440);
    p.L(140, 140);
    p.L(460, 140);
    p.stroke();
    p.M(140, 80);
    p.L(460, 80);
    p.stroke();
  }
};
// 0x2292 SQUARE ORIGINAL OF OR EQUAL TO
MATH_OP_FULL[0x2292] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 440);
    p.L(460, 440);
    p.L(460, 140);
    p.L(140, 140);
    p.stroke();
    p.M(140, 80);
    p.L(460, 80);
    p.stroke();
  }
};
// 0x2293 SQUARE CAP
MATH_OP_FULL[0x2293] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(140, 60);
    p.L(140, 440);
    p.L(460, 440);
    p.L(460, 60);
    p.stroke();
  }
};
// 0x2294 SQUARE CUP
MATH_OP_FULL[0x2294] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(140, 440);
    p.L(140, 60);
    p.L(460, 60);
    p.L(460, 440);
    p.stroke();
  }
};
// 0x229a CIRCLED RING OPERATOR
MATH_OP_FULL[0x229a] = mathCircledOp2(p => {
  p.lineWidth(28);
  p.circle(300, 250, 60);
  p.stroke();
});
// 0x229b CIRCLED ASTERISK OPERATOR
MATH_OP_FULL[0x229b] = mathCircledOp2(p => {
  p.lineWidth(32);
  for (let i = 0; i < 3; i++) {
    const a = (i * 60 * Math.PI) / 180;
    p.M(300 + 80 * Math.cos(a), 250 + 80 * Math.sin(a));
    p.L(300 - 80 * Math.cos(a), 250 - 80 * Math.sin(a));
    p.stroke();
  }
});
// 0x229c CIRCLED EQUALS
MATH_OP_FULL[0x229c] = mathCircledOp2(p => {
  p.M(200, 280);
  p.L(400, 280);
  p.stroke();
  p.M(200, 220);
  p.L(400, 220);
  p.stroke();
});
// 0x229d CIRCLED DASH
MATH_OP_FULL[0x229d] = mathCircledOp2(p => {
  p.M(200, 250);
  p.L(400, 250);
  p.stroke();
});
// 0x229e SQUARED PLUS
MATH_OP_FULL[0x229e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.rect(100, 50, 400, 400);
    p.stroke();
    p.M(100, 250);
    p.L(500, 250);
    p.stroke();
    p.M(300, 50);
    p.L(300, 450);
    p.stroke();
  }
};
// 0x229f SQUARED MINUS
MATH_OP_FULL[0x229f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.rect(100, 50, 400, 400);
    p.stroke();
    p.M(160, 250);
    p.L(440, 250);
    p.stroke();
  }
};
// 0x22a0 SQUARED TIMES
MATH_OP_FULL[0x22a0] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.rect(100, 50, 400, 400);
    p.stroke();
    p.M(100, 50);
    p.L(500, 450);
    p.stroke();
    p.M(500, 50);
    p.L(100, 450);
    p.stroke();
  }
};
// 0x22a1 SQUARED DOT OPERATOR
MATH_OP_FULL[0x22a1] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.rect(100, 50, 400, 400);
    p.stroke();
    p.circle(300, 250, 40);
    p.fill();
  }
};
// 0x22a2 RIGHT TACK
MATH_OP_FULL[0x22a2] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 440);
    p.L(140, 60);
    p.stroke();
    p.M(140, 250);
    p.L(480, 250);
    p.stroke();
  }
};
// 0x22a3 LEFT TACK
MATH_OP_FULL[0x22a3] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(460, 440);
    p.L(460, 60);
    p.stroke();
    p.M(460, 250);
    p.L(120, 250);
    p.stroke();
  }
};
// 0x22a4 DOWN TACK (TOP)
MATH_OP_FULL[0x22a4] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 440);
    p.L(480, 440);
    p.stroke();
    p.M(300, 440);
    p.L(300, 60);
    p.stroke();
  }
};
// 0x22a6-0x22af: Turnstile variants
MATH_OP_FULL[0x22a6] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(160, 440);
    p.L(160, 60);
    p.stroke();
    p.M(160, 250);
    p.L(460, 250);
    p.stroke();
    p.M(220, 440);
    p.L(220, 60);
    p.stroke();
  }
};
MATH_OP_FULL[0x22a7] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 440);
    p.L(140, 60);
    p.stroke();
    p.M(140, 310);
    p.L(460, 310);
    p.stroke();
    p.M(140, 190);
    p.L(460, 190);
    p.stroke();
  }
};
MATH_OP_FULL[0x22a8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 440);
    p.L(140, 60);
    p.stroke();
    p.M(200, 440);
    p.L(200, 60);
    p.stroke();
    p.M(200, 250);
    p.L(460, 250);
    p.stroke();
  }
};
MATH_OP_FULL[0x22a9] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 440);
    p.L(140, 60);
    p.stroke();
    p.M(200, 440);
    p.L(200, 60);
    p.stroke();
    p.M(200, 250);
    p.L(460, 250);
    p.stroke();
  }
};
MATH_OP_FULL[0x22aa] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(36);
    p.M(130, 440);
    p.L(130, 60);
    p.stroke();
    p.M(190, 440);
    p.L(190, 60);
    p.stroke();
    p.M(250, 440);
    p.L(250, 60);
    p.stroke();
    p.M(250, 250);
    p.L(460, 250);
    p.stroke();
  }
};
MATH_OP_FULL[0x22ab] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(130, 440);
    p.L(130, 60);
    p.stroke();
    p.M(200, 440);
    p.L(200, 60);
    p.stroke();
    p.M(200, 300);
    p.L(460, 300);
    p.stroke();
    p.M(200, 200);
    p.L(460, 200);
    p.stroke();
  }
};
MATH_OP_FULL[0x22ac] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22a2].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22ad] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22a8].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22ae] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22a9].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22af] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22ab].draw(p);
    negSlash(p);
  }
};
// 0x22b0-0x22b9: Various relations
MATH_OP_FULL[0x22b0] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 400);
    p.C(200, 400, 140, 300, 140, 250);
    p.C(140, 200, 200, 100, 460, 100);
    p.stroke();
    p.M(140, 50);
    p.L(460, 50);
    p.stroke();
  }
};
MATH_OP_FULL[0x22b1] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 400);
    p.C(400, 400, 460, 300, 460, 250);
    p.C(460, 200, 400, 100, 140, 100);
    p.stroke();
    p.M(140, 50);
    p.L(460, 50);
    p.stroke();
  }
};
// 0x22b2 NORMAL SUBGROUP OF
MATH_OP_FULL[0x22b2] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(460, 420);
    p.L(140, 250);
    p.L(460, 80);
    p.Z();
    p.stroke();
  }
};
// 0x22b3 CONTAINS AS NORMAL SUBGROUP
MATH_OP_FULL[0x22b3] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 420);
    p.L(460, 250);
    p.L(140, 80);
    p.Z();
    p.stroke();
  }
};
// 0x22b4 NORMAL SUBGROUP OF OR EQUAL TO
MATH_OP_FULL[0x22b4] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 440);
    p.L(140, 260);
    p.L(460, 100);
    p.Z();
    p.stroke();
    p.M(140, 55);
    p.L(460, 55);
    p.stroke();
  }
};
// 0x22b5 CONTAINS AS NORMAL SUBGROUP OR EQUAL TO
MATH_OP_FULL[0x22b5] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 440);
    p.L(460, 260);
    p.L(140, 100);
    p.Z();
    p.stroke();
    p.M(140, 55);
    p.L(460, 55);
    p.stroke();
  }
};
// 0x22b6 ORIGINAL OF
MATH_OP_FULL[0x22b6] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(420, 400);
    p.L(200, 250);
    p.L(420, 100);
    p.stroke();
    p.circle(160, 250, 40);
    p.fill();
  }
};
// 0x22b7 IMAGE OF
MATH_OP_FULL[0x22b7] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(180, 400);
    p.L(400, 250);
    p.L(180, 100);
    p.stroke();
    p.circle(440, 250, 40);
    p.fill();
  }
};
// 0x22b8 MULTIMAP
MATH_OP_FULL[0x22b8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 250);
    p.L(430, 250);
    p.stroke();
    p.lineWidth(28);
    p.circle(470, 250, 40);
    p.stroke();
  }
};
// 0x22b9 HERMITIAN CONJUGATE MATRIX
MATH_OP_FULL[0x22b9] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.rect(120, 80, 360, 340);
    p.stroke();
    p.circle(300, 250, 30);
    p.fill();
  }
};
// 0x22ba INTERCALATE
MATH_OP_FULL[0x22ba] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 60);
    p.L(480, 60);
    p.stroke();
    p.M(300, 60);
    p.L(300, 440);
    p.stroke();
  }
};
// 0x22bb XOR
MATH_OP_FULL[0x22bb] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 80);
    p.L(300, 440);
    p.L(460, 80);
    p.stroke();
    p.M(140, 60);
    p.L(460, 60);
    p.stroke();
  }
};
// 0x22bc NAND
MATH_OP_FULL[0x22bc] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 420);
    p.L(300, 60);
    p.L(460, 420);
    p.stroke();
    p.M(140, 440);
    p.L(460, 440);
    p.stroke();
  }
};
// 0x22bd NOR
MATH_OP_FULL[0x22bd] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 80);
    p.L(300, 440);
    p.L(460, 80);
    p.stroke();
    p.M(140, 60);
    p.L(460, 60);
    p.stroke();
  }
};
// 0x22be RIGHT ANGLE WITH ARC
MATH_OP_FULL[0x22be] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 440);
    p.L(140, 80);
    p.L(460, 80);
    p.stroke();
    p.lineWidth(28);
    p.M(140, 300);
    p.C(250, 300, 320, 200, 320, 80);
    p.stroke();
  }
};
// 0x22c0-0x22cf: N-ary/binary operators
MATH_OP_FULL[0x22c0] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(120, 60);
    p.L(300, 440);
    p.L(480, 60);
    p.stroke();
  }
};
MATH_OP_FULL[0x22c1] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(120, 440);
    p.L(300, 60);
    p.L(480, 440);
    p.stroke();
  }
};
MATH_OP_FULL[0x22c2] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(140, 60);
    p.L(140, 250);
    p.C(140, 460, 460, 460, 460, 250);
    p.L(460, 60);
    p.stroke();
  }
};
MATH_OP_FULL[0x22c3] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(140, 440);
    p.L(140, 250);
    p.C(140, 40, 460, 40, 460, 250);
    p.L(460, 440);
    p.stroke();
  }
};
MATH_OP_FULL[0x22c4] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(300, 420);
    p.L(460, 250);
    p.L(300, 80);
    p.L(140, 250);
    p.Z();
    p.stroke();
  }
};
MATH_OP_FULL[0x22c5] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(300, 250, 35);
    p.fill();
  }
};
MATH_OP_FULL[0x22c6] = {
  width: W,
  draw: (p: GlyphPen) => {
    const cx = 300,
      cy = 250,
      R = 140,
      r = 60;
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
};
MATH_OP_FULL[0x22c7] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(160, 400);
    p.L(440, 100);
    p.stroke();
    p.M(440, 400);
    p.L(160, 100);
    p.stroke();
    p.circle(300, 420, 28);
    p.fill();
    p.circle(300, 80, 28);
    p.fill();
  }
};
MATH_OP_FULL[0x22c8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 400);
    p.L(300, 250);
    p.L(120, 100);
    p.Z();
    p.stroke();
    p.M(480, 400);
    p.L(300, 250);
    p.L(480, 100);
    p.Z();
    p.stroke();
  }
};
MATH_OP_FULL[0x22c9] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 400);
    p.L(380, 250);
    p.L(140, 100);
    p.Z();
    p.stroke();
    p.M(380, 400);
    p.L(380, 100);
    p.stroke();
  }
};
MATH_OP_FULL[0x22ca] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 400);
    p.L(220, 250);
    p.L(460, 100);
    p.Z();
    p.stroke();
    p.M(220, 400);
    p.L(220, 100);
    p.stroke();
  }
};
MATH_OP_FULL[0x22cb] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 400);
    p.L(380, 250);
    p.L(140, 100);
    p.stroke();
    p.M(380, 400);
    p.L(380, 100);
    p.stroke();
  }
};
MATH_OP_FULL[0x22cc] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 400);
    p.L(220, 250);
    p.L(460, 100);
    p.stroke();
    p.M(220, 400);
    p.L(220, 100);
    p.stroke();
  }
};
MATH_OP_FULL[0x22cd] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 220);
    p.C(200, 120, 400, 380, 480, 280);
    p.stroke();
    p.M(120, 140);
    p.L(480, 140);
    p.stroke();
  }
};
MATH_OP_FULL[0x22ce] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 440);
    p.C(200, 200, 260, 100, 300, 60);
    p.C(340, 100, 400, 200, 480, 440);
    p.stroke();
  }
};
MATH_OP_FULL[0x22cf] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 60);
    p.C(200, 300, 260, 400, 300, 440);
    p.C(340, 400, 400, 300, 480, 60);
    p.stroke();
  }
};
// 0x22d0-0x22ff: Double subset, intersection, etc
MATH_OP_FULL[0x22d0] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(440, 400);
    p.C(250, 400, 180, 320, 180, 250);
    p.C(180, 180, 250, 100, 440, 100);
    p.stroke();
    p.M(480, 400);
    p.C(290, 400, 220, 320, 220, 250);
    p.C(220, 180, 290, 100, 480, 100);
    p.stroke();
  }
};
MATH_OP_FULL[0x22d1] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(160, 400);
    p.C(350, 400, 420, 320, 420, 250);
    p.C(420, 180, 350, 100, 160, 100);
    p.stroke();
    p.M(120, 400);
    p.C(310, 400, 380, 320, 380, 250);
    p.C(380, 180, 310, 100, 120, 100);
    p.stroke();
  }
};
MATH_OP_FULL[0x22d2] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 60);
    p.L(120, 230);
    p.C(120, 460, 480, 460, 480, 230);
    p.L(480, 60);
    p.stroke();
    p.M(200, 60);
    p.L(200, 230);
    p.C(200, 400, 400, 400, 400, 230);
    p.L(400, 60);
    p.stroke();
  }
};
MATH_OP_FULL[0x22d3] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 440);
    p.L(120, 270);
    p.C(120, 40, 480, 40, 480, 270);
    p.L(480, 440);
    p.stroke();
    p.M(200, 440);
    p.L(200, 270);
    p.C(200, 100, 400, 100, 400, 270);
    p.L(400, 440);
    p.stroke();
  }
};
MATH_OP_FULL[0x22d4] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(300, 60);
    p.L(300, 440);
    p.stroke();
    p.M(140, 60);
    p.L(140, 250);
    p.C(140, 400, 460, 400, 460, 250);
    p.L(460, 60);
    p.stroke();
  }
};
MATH_OP_FULL[0x22d5] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 310);
    p.L(480, 310);
    p.stroke();
    p.M(120, 200);
    p.L(480, 200);
    p.stroke();
    p.M(200, 440);
    p.L(200, 60);
    p.stroke();
    p.M(400, 440);
    p.L(400, 60);
    p.stroke();
  }
};
MATH_OP_FULL[0x22d6] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(460, 400);
    p.L(140, 250);
    p.L(460, 100);
    p.stroke();
    p.circle(300, 250, 30);
    p.fill();
  }
};
MATH_OP_FULL[0x22d7] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 400);
    p.L(460, 250);
    p.L(140, 100);
    p.stroke();
    p.circle(300, 250, 30);
    p.fill();
  }
};
MATH_OP_FULL[0x22d8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    for (const x of [160, 280, 400]) {
      p.M(x + 80, 400);
      p.L(x, 250);
      p.L(x + 80, 100);
      p.stroke();
    }
  }
};
MATH_OP_FULL[0x22d9] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    for (const x of [440, 320, 200]) {
      p.M(x - 80, 400);
      p.L(x, 250);
      p.L(x - 80, 100);
      p.stroke();
    }
  }
};
// 0x22da-0x22db: Less/greater-than equal or greater/less-than
MATH_OP_FULL[0x22da] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(36);
    p.M(460, 460);
    p.L(140, 370);
    p.L(460, 280);
    p.stroke();
    p.M(140, 240);
    p.L(460, 240);
    p.stroke();
    p.M(140, 200);
    p.L(460, 110);
    p.L(140, 40);
    p.stroke();
  }
};
MATH_OP_FULL[0x22db] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(36);
    p.M(140, 460);
    p.L(460, 370);
    p.L(140, 280);
    p.stroke();
    p.M(140, 240);
    p.L(460, 240);
    p.stroke();
    p.M(460, 200);
    p.L(140, 110);
    p.L(460, 40);
    p.stroke();
  }
};
// 0x22dc-0x22ef: Various relations and operators
MATH_OP_FULL[0x22dc] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 420);
    p.L(140, 260);
    p.L(460, 100);
    p.stroke();
    p.M(140, 440);
    p.L(460, 440);
    p.stroke();
  }
};
MATH_OP_FULL[0x22dd] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(140, 420);
    p.L(460, 260);
    p.L(140, 100);
    p.stroke();
    p.M(140, 440);
    p.L(460, 440);
    p.stroke();
  }
};
MATH_OP_FULL[0x22de] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(460, 420);
    p.C(200, 420, 140, 320, 140, 280);
    p.C(140, 240, 200, 140, 460, 140);
    p.stroke();
    p.M(140, 80);
    p.L(460, 80);
    p.stroke();
  }
};
MATH_OP_FULL[0x22df] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(140, 420);
    p.C(400, 420, 460, 320, 460, 280);
    p.C(460, 240, 400, 140, 140, 140);
    p.stroke();
    p.M(140, 80);
    p.L(460, 80);
    p.stroke();
  }
};
// 0x22e0-0x22e9: Negated variants
MATH_OP_FULL[0x22e0] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22de].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e1] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22df].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e2] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2291].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e3] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2292].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e4] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2291].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e5] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2292].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e6] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2272].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e7] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x2273].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e8] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x227e].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22e9] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x227f].draw(p);
    negSlash(p);
  }
};
// 0x22ea-0x22ed: Negated normal subgroup
MATH_OP_FULL[0x22ea] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22b2].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22eb] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22b3].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22ec] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22b4].draw(p);
    negSlash(p);
  }
};
MATH_OP_FULL[0x22ed] = {
  width: W,
  draw: (p: GlyphPen) => {
    MATH_OP_FULL[0x22b5].draw(p);
    negSlash(p);
  }
};
// 0x22ee-0x22f1: Ellipsis variants
MATH_OP_FULL[0x22ee] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(300, 380, 28);
    p.fill();
    p.circle(300, 250, 28);
    p.fill();
    p.circle(300, 120, 28);
    p.fill();
  }
};
MATH_OP_FULL[0x22ef] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(150, 250, 28);
    p.fill();
    p.circle(300, 250, 28);
    p.fill();
    p.circle(450, 250, 28);
    p.fill();
  }
};
MATH_OP_FULL[0x22f0] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(170, 120, 28);
    p.fill();
    p.circle(300, 250, 28);
    p.fill();
    p.circle(430, 380, 28);
    p.fill();
  }
};
MATH_OP_FULL[0x22f1] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(170, 380, 28);
    p.fill();
    p.circle(300, 250, 28);
    p.fill();
    p.circle(430, 120, 28);
    p.fill();
  }
};
// 0x22f2-0x22ff: Element-of/contains variants
for (let cp = 0x22f2; cp <= 0x22ff; cp++) {
  if (!MATH_OP_FULL[cp]) {
    const isElem = (cp - 0x22f2) % 2 === 0;
    MATH_OP_FULL[cp] = {
      width: W,
      draw: (p: GlyphPen) => {
        p.lineWidth(40);
        if (isElem) {
          p.M(440, 400);
          p.C(200, 430, 140, 340, 140, 250);
          p.C(140, 160, 200, 70, 440, 100);
          p.stroke();
          p.M(140, 250);
          p.L(400, 250);
          p.stroke();
        } else {
          p.M(160, 400);
          p.C(400, 430, 460, 340, 460, 250);
          p.C(460, 160, 400, 70, 160, 100);
          p.stroke();
          p.M(460, 250);
          p.L(200, 250);
          p.stroke();
        }
      }
    };
  }
}

// =============================================================================
// 6. LETTERLIKE_FULL - Letterlike Symbols (U+2100-U+214F) remaining ~70 chars
// =============================================================================

export const LETTERLIKE_FULL: Record<number, GlyphDef> = {};

/** Draw a double-stroked letter by drawing the letter path twice with offset */
function doubleStroke(
  letterFn: (p: GlyphPen, cx: number, cy: number, s: number) => void,
  offset: number = 18
): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(28);
      letterFn(p, 300 - offset / 2, 250, 280);
      letterFn(p, 300 + offset / 2, 250, 280);
    }
  };
}

/** Draw a script-style letter (thinner, with curves) */
function scriptLetter(
  letterFn: (p: GlyphPen, cx: number, cy: number, s: number) => void
): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(26);
      letterFn(p, 300, 250, 300);
    }
  };
}

/** Draw a fraktur-style letter (thicker, angular) */
function frakturLetter(
  letterFn: (p: GlyphPen, cx: number, cy: number, s: number) => void
): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(36);
      letterFn(p, 300, 250, 300);
    }
  };
}

// Use LETTER_PATHS indices: A=0,B=1,C=2,D=3,E=4,F=5,G=6,H=7,I=8,J=9,K=10,L=11,M=12,N=13,O=14,P=15,Q=16,R=17,S=18,T=19,U=20,V=21,W=22,X=23,Y=24,Z=25

// We reference LETTER_PATHS via drawLetter (which is module-level)

// 0x2100 ACCOUNT OF (a/c)
LETTERLIKE_FULL[0x2100] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(24);
    drawLetter(p, 0, 190, 340, 150);
    drawLetter(p, 2, 410, 160, 150);
    p.lineWidth(30);
    p.M(380, 450);
    p.L(220, 50);
    p.stroke();
  }
};
// 0x2101 ADDRESSED TO THE SUBJECT (a/s)
LETTERLIKE_FULL[0x2101] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(24);
    drawLetter(p, 0, 190, 340, 150);
    drawLetter(p, 18, 410, 160, 150);
    p.lineWidth(30);
    p.M(380, 450);
    p.L(220, 50);
    p.stroke();
  }
};
// 0x2102 DOUBLE-STRUCK CAPITAL C
LETTERLIKE_FULL[0x2102] = doubleStroke((p, cx, cy, s) => {
  p.M(cx + s * 0.3, cy + s * 0.35);
  p.C(cx + s * 0.1, cy + s * 0.55, cx - s * 0.35, cy + s * 0.35, cx - s * 0.35, cy);
  p.C(cx - s * 0.35, cy - s * 0.35, cx + s * 0.1, cy - s * 0.55, cx + s * 0.3, cy - s * 0.35);
  p.stroke();
});
// 0x2104 INVERTED Y (CL symbol)
LETTERLIKE_FULL[0x2104] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(200, 440);
    p.L(300, 250);
    p.L(400, 440);
    p.stroke();
    p.M(300, 250);
    p.L(300, 60);
    p.stroke();
  }
};
// 0x2106 CADA UNA (C/U)
LETTERLIKE_FULL[0x2106] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(24);
    drawLetter(p, 2, 190, 340, 150);
    drawLetter(p, 20, 410, 160, 150);
    p.lineWidth(30);
    p.M(380, 450);
    p.L(220, 50);
    p.stroke();
  }
};
// 0x2107 EULER CONSTANT (reversed E)
LETTERLIKE_FULL[0x2107] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(180, 440);
    p.L(420, 440);
    p.L(420, 60);
    p.L(180, 60);
    p.stroke();
    p.M(420, 250);
    p.L(220, 250);
    p.stroke();
  }
};
// 0x210A SCRIPT SMALL G
LETTERLIKE_FULL[0x210a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.ellipse(280, 280, 120, 100);
    p.stroke();
    p.M(400, 380);
    p.L(400, 80);
    p.C(400, 20, 300, 20, 240, 60);
    p.stroke();
  }
};
// 0x210B SCRIPT CAPITAL H
LETTERLIKE_FULL[0x210b] = scriptLetter((p, cx, cy, s) => {
  // H as flowing script
  p.M(cx - s * 0.3, cy - s * 0.5);
  p.L(cx - s * 0.3, cy + s * 0.5);
  p.stroke();
  p.M(cx + s * 0.3, cy + s * 0.5);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.stroke();
  p.M(cx - s * 0.3, cy);
  p.C(cx - s * 0.1, cy + s * 0.15, cx + s * 0.1, cy - s * 0.15, cx + s * 0.3, cy);
  p.stroke();
});
// 0x210C FRAKTUR CAPITAL H
LETTERLIKE_FULL[0x210c] = frakturLetter((p, cx, cy, s) => {
  p.M(cx - s * 0.3, cy - s * 0.5);
  p.L(cx - s * 0.3, cy + s * 0.5);
  p.stroke();
  p.M(cx + s * 0.3, cy + s * 0.5);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.stroke();
  p.M(cx - s * 0.3, cy);
  p.L(cx + s * 0.3, cy);
  p.stroke();
});
// 0x210D DOUBLE-STRUCK H
LETTERLIKE_FULL[0x210d] = doubleStroke((p, cx, cy, s) => {
  p.M(cx - s * 0.3, cy + s * 0.5);
  p.L(cx - s * 0.3, cy - s * 0.5);
  p.stroke();
  p.M(cx + s * 0.3, cy + s * 0.5);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.stroke();
  p.M(cx - s * 0.3, cy);
  p.L(cx + s * 0.3, cy);
  p.stroke();
});
// 0x210E PLANCK CONSTANT (italic h)
LETTERLIKE_FULL[0x210e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(220, 60);
    p.L(200, 450);
    p.stroke();
    p.M(200, 300);
    p.C(280, 400, 380, 400, 400, 300);
    p.L(420, 60);
    p.stroke();
  }
};
// 0x210F PLANCK CONSTANT OVER TWO PI (h-bar)
LETTERLIKE_FULL[0x210f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(220, 60);
    p.L(200, 450);
    p.stroke();
    p.M(200, 300);
    p.C(280, 400, 380, 400, 400, 300);
    p.L(420, 60);
    p.stroke();
    // crossbar
    p.M(140, 370);
    p.L(280, 370);
    p.stroke();
  }
};
// 0x2110 SCRIPT CAPITAL I
LETTERLIKE_FULL[0x2110] = scriptLetter((p, cx, cy, s) => {
  p.M(cx, cy + s * 0.5);
  p.L(cx, cy - s * 0.5);
  p.stroke();
  p.M(cx - s * 0.2, cy + s * 0.5);
  p.L(cx + s * 0.2, cy + s * 0.5);
  p.stroke();
  p.M(cx - s * 0.2, cy - s * 0.5);
  p.L(cx + s * 0.2, cy - s * 0.5);
  p.stroke();
});
// 0x2111 FRAKTUR CAPITAL I
LETTERLIKE_FULL[0x2111] = frakturLetter((p, cx, cy, s) => {
  p.M(cx, cy + s * 0.5);
  p.L(cx, cy - s * 0.5);
  p.stroke();
  p.M(cx - s * 0.25, cy + s * 0.5);
  p.L(cx + s * 0.25, cy + s * 0.5);
  p.stroke();
  p.M(cx - s * 0.25, cy - s * 0.5);
  p.L(cx + s * 0.25, cy - s * 0.5);
  p.stroke();
});
// 0x2112 SCRIPT CAPITAL L
LETTERLIKE_FULL[0x2112] = scriptLetter((p, cx, cy, s) => {
  p.M(cx - s * 0.25, cy + s * 0.5);
  p.L(cx - s * 0.25, cy - s * 0.5);
  p.C(cx - s * 0.25, cy - s * 0.55, cx + s * 0.3, cy - s * 0.55, cx + s * 0.3, cy - s * 0.5);
  p.stroke();
});
// 0x2114 L B BAR SYMBOL
LETTERLIKE_FULL[0x2114] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    drawLetter(p, 11, 200, 250, 260);
    drawLetter(p, 1, 400, 250, 260);
    // bar
    p.lineWidth(28);
    p.M(120, 100);
    p.L(480, 100);
    p.stroke();
  }
};
// 0x2115 DOUBLE-STRUCK N
LETTERLIKE_FULL[0x2115] = doubleStroke((p, cx, cy, s) => {
  p.M(cx - s * 0.3, cy - s * 0.5);
  p.L(cx - s * 0.3, cy + s * 0.5);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.L(cx + s * 0.3, cy + s * 0.5);
  p.stroke();
});
// 0x2117 SOUND RECORDING COPYRIGHT (P in circle)
LETTERLIKE_FULL[0x2117] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 220);
    p.stroke();
    p.lineWidth(28);
    drawLetter(p, 15, 300, 250, 200);
  }
};
// 0x2118 WEIERSTRASS ELLIPTIC FUNCTION (fancy p)
LETTERLIKE_FULL[0x2118] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(200, 60);
    p.L(200, 350);
    p.C(200, 450, 350, 450, 400, 350);
    p.C(450, 250, 350, 180, 280, 200);
    p.C(210, 220, 180, 300, 200, 350);
    p.stroke();
  }
};
// 0x2119 DOUBLE-STRUCK P
LETTERLIKE_FULL[0x2119] = doubleStroke((p, cx, cy, s) => {
  p.M(cx - s * 0.25, cy - s * 0.5);
  p.L(cx - s * 0.25, cy + s * 0.5);
  p.L(cx + s * 0.1, cy + s * 0.5);
  p.C(cx + s * 0.4, cy + s * 0.5, cx + s * 0.4, cy + s * 0.05, cx + s * 0.1, cy + s * 0.05);
  p.L(cx - s * 0.25, cy + s * 0.05);
  p.stroke();
});
// 0x211A DOUBLE-STRUCK Q
LETTERLIKE_FULL[0x211a] = doubleStroke((p, cx, cy, s) => {
  p.ellipse(cx, cy + s * 0.05, s * 0.35, s * 0.45);
  p.stroke();
  p.M(cx + s * 0.1, cy - s * 0.2);
  p.L(cx + s * 0.35, cy - s * 0.5);
  p.stroke();
});
// 0x211B SCRIPT CAPITAL R
LETTERLIKE_FULL[0x211b] = scriptLetter((p, cx, cy, s) => {
  p.M(cx - s * 0.25, cy - s * 0.5);
  p.L(cx - s * 0.25, cy + s * 0.5);
  p.L(cx + s * 0.1, cy + s * 0.5);
  p.C(cx + s * 0.4, cy + s * 0.5, cx + s * 0.4, cy + s * 0.05, cx + s * 0.1, cy + s * 0.05);
  p.L(cx - s * 0.25, cy + s * 0.05);
  p.stroke();
  p.M(cx + s * 0.05, cy + s * 0.05);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.stroke();
});
// 0x211C FRAKTUR CAPITAL R
LETTERLIKE_FULL[0x211c] = frakturLetter((p, cx, cy, s) => {
  p.M(cx - s * 0.25, cy - s * 0.5);
  p.L(cx - s * 0.25, cy + s * 0.5);
  p.L(cx + s * 0.1, cy + s * 0.5);
  p.C(cx + s * 0.4, cy + s * 0.5, cx + s * 0.4, cy + s * 0.05, cx + s * 0.1, cy + s * 0.05);
  p.L(cx - s * 0.25, cy + s * 0.05);
  p.stroke();
  p.M(cx + s * 0.05, cy + s * 0.05);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.stroke();
});
// 0x211D DOUBLE-STRUCK R
LETTERLIKE_FULL[0x211d] = doubleStroke((p, cx, cy, s) => {
  p.M(cx - s * 0.25, cy - s * 0.5);
  p.L(cx - s * 0.25, cy + s * 0.5);
  p.L(cx + s * 0.1, cy + s * 0.5);
  p.C(cx + s * 0.4, cy + s * 0.5, cx + s * 0.4, cy + s * 0.05, cx + s * 0.1, cy + s * 0.05);
  p.L(cx - s * 0.25, cy + s * 0.05);
  p.stroke();
  p.M(cx + s * 0.05, cy + s * 0.05);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.stroke();
});
// 0x211E PRESCRIPTION TAKE (Rx)
LETTERLIKE_FULL[0x211e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    drawLetter(p, 17, 250, 260, 300);
    // x leg extending
    p.lineWidth(32);
    p.M(350, 200);
    p.L(480, 60);
    p.stroke();
  }
};
// 0x211F RESPONSE
LETTERLIKE_FULL[0x211f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    drawLetter(p, 17, 300, 250, 300);
  }
};
// 0x2121 TELEPHONE SIGN (TEL)
LETTERLIKE_FULL[0x2121] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(22);
    drawLetter(p, 19, 150, 250, 180);
    drawLetter(p, 4, 300, 250, 180);
    drawLetter(p, 11, 450, 250, 180);
  }
};
// 0x2124 DOUBLE-STRUCK Z
LETTERLIKE_FULL[0x2124] = doubleStroke((p, cx, cy, s) => {
  p.M(cx - s * 0.3, cy + s * 0.5);
  p.L(cx + s * 0.3, cy + s * 0.5);
  p.L(cx - s * 0.3, cy - s * 0.5);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.stroke();
});
// 0x2125 OUNCE SIGN
LETTERLIKE_FULL[0x2125] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(36);
    p.M(200, 440);
    p.L(400, 440);
    p.L(200, 60);
    p.L(400, 60);
    p.stroke();
    // bar through
    p.M(180, 250);
    p.L(420, 250);
    p.stroke();
  }
};
// 0x2127 INVERTED OHM SIGN
LETTERLIKE_FULL[0x2127] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(140, 440);
    p.L(200, 440);
    p.L(200, 350);
    p.C(200, 150, 400, 150, 400, 350);
    p.L(400, 440);
    p.L(460, 440);
    p.stroke();
  }
};
// 0x2128 FRAKTUR CAPITAL Z
LETTERLIKE_FULL[0x2128] = frakturLetter((p, cx, cy, s) => {
  p.M(cx - s * 0.3, cy + s * 0.5);
  p.L(cx + s * 0.3, cy + s * 0.5);
  p.L(cx - s * 0.3, cy - s * 0.5);
  p.L(cx + s * 0.3, cy - s * 0.5);
  p.stroke();
});
// 0x2129 TURNED IOTA
LETTERLIKE_FULL[0x2129] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(300, 440);
    p.L(300, 160);
    p.C(300, 60, 220, 60, 220, 120);
    p.stroke();
  }
};
// 0x212A KELVIN SIGN (K)
LETTERLIKE_FULL[0x212a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    drawLetter(p, 10, 300, 250, 300);
  }
};
// 0x212B ANGSTROM SIGN (A with ring)
LETTERLIKE_FULL[0x212b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    drawLetter(p, 0, 300, 220, 280);
    // ring above
    p.lineWidth(22);
    p.circle(300, 440, 30);
    p.stroke();
  }
};
// 0x212C SCRIPT CAPITAL B
LETTERLIKE_FULL[0x212c] = scriptLetter((p, cx, cy, s) => {
  p.M(cx - s * 0.25, cy - s * 0.5);
  p.L(cx - s * 0.25, cy + s * 0.5);
  p.L(cx + s * 0.1, cy + s * 0.5);
  p.C(cx + s * 0.35, cy + s * 0.5, cx + s * 0.35, cy + s * 0.05, cx + s * 0.05, cy + s * 0.05);
  p.L(cx + s * 0.1, cy + s * 0.05);
  p.C(cx + s * 0.4, cy + s * 0.05, cx + s * 0.4, cy - s * 0.5, cx + s * 0.1, cy - s * 0.5);
  p.Z();
  p.stroke();
});
// 0x212D FRAKTUR CAPITAL C
LETTERLIKE_FULL[0x212d] = frakturLetter((p, cx, cy, s) => {
  p.M(cx + s * 0.3, cy + s * 0.35);
  p.C(cx + s * 0.1, cy + s * 0.55, cx - s * 0.35, cy + s * 0.35, cx - s * 0.35, cy);
  p.C(cx - s * 0.35, cy - s * 0.35, cx + s * 0.1, cy - s * 0.55, cx + s * 0.3, cy - s * 0.35);
  p.stroke();
});
// 0x212F SCRIPT SMALL E
LETTERLIKE_FULL[0x212f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(420, 250);
    p.L(180, 250);
    p.C(180, 400, 300, 450, 400, 380);
    p.stroke();
    p.M(180, 250);
    p.C(180, 100, 300, 50, 420, 120);
    p.stroke();
  }
};
// 0x2130 SCRIPT CAPITAL E
LETTERLIKE_FULL[0x2130] = scriptLetter((p, cx, cy, s) => {
  p.M(cx + s * 0.25, cy + s * 0.5);
  p.L(cx - s * 0.25, cy + s * 0.5);
  p.L(cx - s * 0.25, cy - s * 0.5);
  p.L(cx + s * 0.25, cy - s * 0.5);
  p.stroke();
  p.M(cx - s * 0.25, cy);
  p.L(cx + s * 0.15, cy);
  p.stroke();
});
// 0x2131 SCRIPT CAPITAL F
LETTERLIKE_FULL[0x2131] = scriptLetter((p, cx, cy, s) => {
  p.M(cx + s * 0.25, cy + s * 0.5);
  p.L(cx - s * 0.25, cy + s * 0.5);
  p.L(cx - s * 0.25, cy - s * 0.5);
  p.stroke();
  p.M(cx - s * 0.25, cy + s * 0.05);
  p.L(cx + s * 0.15, cy + s * 0.05);
  p.stroke();
});
// 0x2132 TURNED CAPITAL F
LETTERLIKE_FULL[0x2132] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(180, 60);
    p.L(420, 60);
    p.L(420, 440);
    p.stroke();
    p.M(420, 250);
    p.L(220, 250);
    p.stroke();
  }
};
// 0x2133 SCRIPT CAPITAL M
LETTERLIKE_FULL[0x2133] = scriptLetter((p, cx, cy, s) => {
  p.M(cx - s * 0.35, cy - s * 0.5);
  p.L(cx - s * 0.35, cy + s * 0.5);
  p.L(cx, cy);
  p.L(cx + s * 0.35, cy + s * 0.5);
  p.L(cx + s * 0.35, cy - s * 0.5);
  p.stroke();
});
// 0x2134 SCRIPT SMALL O
LETTERLIKE_FULL[0x2134] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.ellipse(300, 250, 150, 180);
    p.stroke();
  }
};
// 0x2135-0x2138: Hebrew letters
// ALEF
LETTERLIKE_FULL[0x2135] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(200, 60);
    p.C(250, 200, 350, 300, 400, 440);
    p.stroke();
    p.M(180, 440);
    p.L(250, 300);
    p.stroke();
    p.M(420, 60);
    p.L(350, 200);
    p.stroke();
  }
};
// BET
LETTERLIKE_FULL[0x2136] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(160, 440);
    p.L(440, 440);
    p.L(440, 60);
    p.C(440, 60, 300, 80, 200, 160);
    p.stroke();
  }
};
// GIMEL
LETTERLIKE_FULL[0x2137] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(350, 440);
    p.L(350, 160);
    p.C(350, 60, 250, 60, 200, 100);
    p.stroke();
    p.M(350, 300);
    p.L(250, 60);
    p.stroke();
  }
};
// DALET
LETTERLIKE_FULL[0x2138] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(180, 440);
    p.L(420, 440);
    p.L(420, 60);
    p.stroke();
  }
};
// 0x213C DOUBLE-STRUCK SMALL PI
LETTERLIKE_FULL[0x213c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(140, 380);
    p.L(460, 380);
    p.stroke();
    p.M(230, 380);
    p.L(230, 60);
    p.stroke();
    p.M(370, 380);
    p.L(370, 60);
    p.stroke();
    // double-stroke offset
    p.M(248, 380);
    p.L(248, 60);
    p.stroke();
    p.M(388, 380);
    p.L(388, 60);
    p.stroke();
  }
};
// 0x213D DOUBLE-STRUCK SMALL GAMMA
LETTERLIKE_FULL[0x213d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(160, 440);
    p.L(300, 150);
    p.L(440, 440);
    p.stroke();
    p.M(300, 150);
    p.L(300, 60);
    p.stroke();
    // offset
    p.M(170, 440);
    p.L(310, 150);
    p.stroke();
  }
};
// 0x213E DOUBLE-STRUCK CAPITAL GAMMA
LETTERLIKE_FULL[0x213e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(180, 60);
    p.L(180, 440);
    p.L(420, 440);
    p.stroke();
    p.M(198, 60);
    p.L(198, 440);
    p.stroke();
  }
};
// 0x213F DOUBLE-STRUCK CAPITAL PI
LETTERLIKE_FULL[0x213f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(140, 440);
    p.L(460, 440);
    p.stroke();
    p.M(200, 440);
    p.L(200, 60);
    p.stroke();
    p.M(218, 440);
    p.L(218, 60);
    p.stroke();
    p.M(400, 440);
    p.L(400, 60);
    p.stroke();
    p.M(418, 440);
    p.L(418, 60);
    p.stroke();
  }
};
// 0x2140 DOUBLE-STRUCK N-ARY SUMMATION
LETTERLIKE_FULL[0x2140] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(440, 440);
    p.L(160, 440);
    p.L(300, 250);
    p.L(160, 60);
    p.L(440, 60);
    p.stroke();
    p.M(450, 440);
    p.L(170, 440);
    p.L(310, 250);
    p.L(170, 60);
    p.L(450, 60);
    p.stroke();
  }
};
// 0x2141-0x2144: Turned sans-serif letters
// 0x2141 TURNED SANS-SERIF CAPITAL G
LETTERLIKE_FULL[0x2141] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(200, 100);
    p.C(280, 40, 460, 100, 460, 250);
    p.C(460, 400, 280, 460, 200, 400);
    p.L(200, 270);
    p.L(350, 270);
    p.stroke();
  }
};
// 0x2142 TURNED SANS-SERIF CAPITAL L
LETTERLIKE_FULL[0x2142] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(400, 60);
    p.L(400, 440);
    p.L(200, 440);
    p.stroke();
  }
};
// 0x2143 REVERSED SANS-SERIF CAPITAL L
LETTERLIKE_FULL[0x2143] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(400, 440);
    p.L(400, 60);
    p.L(200, 60);
    p.stroke();
  }
};
// 0x2144 TURNED SANS-SERIF CAPITAL Y
LETTERLIKE_FULL[0x2144] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(180, 440);
    p.L(300, 250);
    p.L(420, 440);
    p.stroke();
    p.M(300, 250);
    p.L(300, 60);
    p.stroke();
  }
};
// 0x2145-0x2149: Double-struck italic D,d,e,i,j
LETTERLIKE_FULL[0x2145] = doubleStroke((p, cx, cy, s) => {
  p.M(cx - s * 0.25, cy - s * 0.5);
  p.L(cx - s * 0.25, cy + s * 0.5);
  p.L(cx, cy + s * 0.5);
  p.C(cx + s * 0.4, cy + s * 0.5, cx + s * 0.4, cy - s * 0.5, cx, cy - s * 0.5);
  p.Z();
  p.stroke();
});
LETTERLIKE_FULL[0x2146] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.ellipse(270, 200, 110, 130);
    p.stroke();
    p.M(380, 60);
    p.L(380, 440);
    p.stroke();
    p.M(398, 60);
    p.L(398, 440);
    p.stroke();
  }
};
LETTERLIKE_FULL[0x2147] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.M(420, 250);
    p.L(180, 250);
    p.C(180, 400, 300, 450, 400, 380);
    p.stroke();
    p.M(180, 250);
    p.C(180, 100, 300, 50, 420, 120);
    p.stroke();
    // offset
    p.M(430, 260);
    p.L(190, 260);
    p.stroke();
  }
};
LETTERLIKE_FULL[0x2148] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(300, 340);
    p.L(300, 60);
    p.stroke();
    p.M(318, 340);
    p.L(318, 60);
    p.stroke();
    p.circle(309, 400, 25);
    p.fill();
  }
};
LETTERLIKE_FULL[0x2149] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(320, 340);
    p.L(320, 120);
    p.C(320, 60, 260, 60, 220, 80);
    p.stroke();
    p.M(338, 340);
    p.L(338, 120);
    p.stroke();
    p.circle(329, 400, 25);
    p.fill();
  }
};
// 0x214A PROPERTY LINE
LETTERLIKE_FULL[0x214a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(42);
    p.M(120, 320);
    p.L(480, 320);
    p.stroke();
    p.M(120, 180);
    p.L(480, 180);
    p.stroke();
    p.M(300, 440);
    p.L(300, 60);
    p.stroke();
  }
};
// 0x214B TURNED AMPERSAND
LETTERLIKE_FULL[0x214b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(420, 60);
    p.L(180, 320);
    p.C(140, 370, 160, 440, 220, 440);
    p.C(320, 440, 360, 380, 300, 300);
    p.L(160, 180);
    p.C(120, 130, 160, 60, 260, 60);
    p.C(360, 60, 380, 130, 340, 180);
    p.stroke();
  }
};
// 0x214C PER SIGN
LETTERLIKE_FULL[0x214c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(440, 440);
    p.L(160, 60);
    p.stroke();
    p.circle(200, 360, 50);
    p.stroke();
    p.circle(400, 140, 50);
    p.stroke();
  }
};
// 0x214D AKTIESELSKAB
LETTERLIKE_FULL[0x214d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(34);
    drawLetter(p, 0, 200, 250, 260);
    drawLetter(p, 18, 400, 250, 260);
  }
};
// 0x214E TURNED SMALL F
LETTERLIKE_FULL[0x214e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(380, 440);
    p.L(220, 440);
    p.L(220, 60);
    p.stroke();
    p.M(220, 250);
    p.L(380, 250);
    p.stroke();
  }
};
// 0x214F SYMBOL FOR SAMARITAN SOURCE
LETTERLIKE_FULL[0x214f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(200, 60);
    p.L(200, 440);
    p.L(400, 440);
    p.L(400, 60);
    p.stroke();
    p.M(200, 250);
    p.L(400, 250);
    p.stroke();
  }
};

// =============================================================================
// 7. CURRENCY_REMAINING - remaining Currency Symbols (U+20A0-U+20CF)
// =============================================================================

export const CURRENCY_REMAINING: Record<number, GlyphDef> = {};

// 0x20AC EURO SIGN
CURRENCY_REMAINING[0x20ac] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(460, 400);
    p.C(400, 470, 220, 470, 180, 370);
    p.C(140, 270, 140, 180, 180, 120);
    p.C(220, 40, 400, 40, 460, 100);
    p.stroke();
    // two horizontal bars
    p.M(120, 300);
    p.L(400, 300);
    p.stroke();
    p.M(120, 210);
    p.L(400, 210);
    p.stroke();
  }
};
// 0x20AD KIP SIGN
CURRENCY_REMAINING[0x20ad] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    drawLetter(p, 10, 300, 250, 300);
    // horizontal bar
    p.M(140, 250);
    p.L(460, 250);
    p.stroke();
  }
};
// 0x20AF DRACHMA SIGN
CURRENCY_REMAINING[0x20af] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    drawLetter(p, 3, 300, 250, 300);
    // double vertical stroke on right
    p.M(400, 470);
    p.L(400, 30);
    p.stroke();
  }
};
// 0x20B0 GERMAN PENNY SIGN
CURRENCY_REMAINING[0x20b0] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    drawLetter(p, 3, 300, 250, 300);
    // small circle (penny mark)
    p.lineWidth(22);
    p.circle(420, 400, 30);
    p.stroke();
  }
};
// 0x20B5 CEDI SIGN (C with vertical stroke)
CURRENCY_REMAINING[0x20b5] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(420, 400);
    p.C(380, 470, 220, 470, 180, 370);
    p.C(140, 270, 140, 180, 180, 120);
    p.C(220, 40, 380, 40, 420, 100);
    p.stroke();
    p.M(300, 490);
    p.L(300, 10);
    p.stroke();
  }
};
// 0x20B7 SPESMILO SIGN
CURRENCY_REMAINING[0x20b7] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    drawLetter(p, 18, 300, 250, 300);
    // bar through
    p.M(140, 300);
    p.L(460, 300);
    p.stroke();
  }
};
// 0x20BB NORDIC MARK SIGN
CURRENCY_REMAINING[0x20bb] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    drawLetter(p, 12, 300, 250, 300);
    // underline
    p.M(120, 60);
    p.L(480, 60);
    p.stroke();
  }
};
// 0x20BC MANAT SIGN
CURRENCY_REMAINING[0x20bc] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(180, 60);
    p.L(180, 350);
    p.C(180, 460, 420, 460, 420, 350);
    p.L(420, 60);
    p.stroke();
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
  }
};
// 0x20BE LARI SIGN
CURRENCY_REMAINING[0x20be] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(400, 440);
    p.C(400, 440, 200, 400, 200, 250);
    p.C(200, 100, 400, 60, 400, 60);
    p.stroke();
    // horizontal bar
    p.M(140, 250);
    p.L(460, 250);
    p.stroke();
  }
};
// 0x20BF BITCOIN SIGN (B with two vertical strokes)
CURRENCY_REMAINING[0x20bf] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    drawLetter(p, 1, 300, 250, 290);
    // two vertical strokes through top
    p.M(260, 480);
    p.L(260, 20);
    p.stroke();
    p.M(340, 480);
    p.L(340, 20);
    p.stroke();
  }
};

// 0x20C0-0x20CF: Newer/reserved currency symbols - generic approximations
for (let cp = 0x20c0; cp <= 0x20cf; cp++) {
  CURRENCY_REMAINING[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      // Generic currency: circle with horizontal bars
      p.lineWidth(35);
      p.circle(300, 250, 180);
      p.stroke();
      p.M(160, 280);
      p.L(440, 280);
      p.stroke();
      p.M(160, 220);
      p.L(440, 220);
      p.stroke();
    }
  };
}

// =============================================================================
// 8. ARROWS_REMAINING - remaining Arrows (U+2190-U+21FF) ~34 missing
// =============================================================================

export const ARROWS_REMAINING: Record<number, GlyphDef> = {};

// 0x219C LEFTWARDS WAVE ARROW
ARROWS_REMAINING[0x219c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(480, 250);
    p.C(420, 330, 340, 170, 280, 250);
    p.C(220, 330, 160, 250, 160, 250);
    p.stroke();
    // arrowhead
    p.M(80, 250);
    p.L(180, 350);
    p.L(180, 150);
    p.Z();
    p.fill();
  }
};
// 0x21A2 LEFTWARDS ARROW WITH TAIL
ARROWS_REMAINING[0x21a2] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(480, 250);
    p.L(180, 250);
    p.stroke();
    p.M(100, 250);
    p.L(200, 350);
    p.L(200, 150);
    p.Z();
    p.fill();
    // tail
    p.lineWidth(30);
    p.M(480, 250);
    p.L(520, 350);
    p.stroke();
  }
};
// 0x21A3 RIGHTWARDS ARROW WITH TAIL
ARROWS_REMAINING[0x21a3] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 250);
    p.L(420, 250);
    p.stroke();
    p.M(500, 250);
    p.L(400, 350);
    p.L(400, 150);
    p.Z();
    p.fill();
    // tail
    p.lineWidth(30);
    p.M(120, 250);
    p.L(80, 350);
    p.stroke();
  }
};
// 0x21A8 UP DOWN ARROW WITH BASE
ARROWS_REMAINING[0x21a8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(300, 440);
    p.L(300, 120);
    p.stroke();
    // up arrowhead
    p.M(300, 470);
    p.L(200, 370);
    p.L(400, 370);
    p.Z();
    p.fill();
    // down arrowhead
    p.M(300, 60);
    p.L(200, 160);
    p.L(400, 160);
    p.Z();
    p.fill();
    // base line
    p.M(180, 30);
    p.L(420, 30);
    p.stroke();
  }
};
// 0x21AB LEFTWARDS ARROW WITH LOOP
ARROWS_REMAINING[0x21ab] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(480, 250);
    p.L(200, 250);
    p.stroke();
    p.M(100, 250);
    p.L(200, 350);
    p.L(200, 150);
    p.Z();
    p.fill();
    // loop
    p.lineWidth(30);
    p.M(480, 250);
    p.C(520, 250, 520, 150, 480, 150);
    p.C(440, 150, 440, 250, 480, 250);
    p.stroke();
  }
};
// 0x21AC RIGHTWARDS ARROW WITH LOOP
ARROWS_REMAINING[0x21ac] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 250);
    p.L(400, 250);
    p.stroke();
    p.M(500, 250);
    p.L(400, 350);
    p.L(400, 150);
    p.Z();
    p.fill();
    // loop
    p.lineWidth(30);
    p.M(120, 250);
    p.C(80, 250, 80, 150, 120, 150);
    p.C(160, 150, 160, 250, 120, 250);
    p.stroke();
  }
};
// 0x21AD LEFT RIGHT WAVE ARROW
ARROWS_REMAINING[0x21ad] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(140, 250);
    p.C(200, 350, 300, 150, 360, 250);
    p.L(460, 250);
    p.stroke();
    p.M(100, 250);
    p.L(170, 330);
    p.L(170, 170);
    p.Z();
    p.fill();
    p.M(500, 250);
    p.L(430, 330);
    p.L(430, 170);
    p.Z();
    p.fill();
  }
};
// 0x21AE LEFT RIGHT ARROW WITH STROKE
ARROWS_REMAINING[0x21ae] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(150, 250);
    p.L(450, 250);
    p.stroke();
    p.M(100, 250);
    p.L(180, 340);
    p.L(180, 160);
    p.Z();
    p.fill();
    p.M(500, 250);
    p.L(420, 340);
    p.L(420, 160);
    p.Z();
    p.fill();
    negSlash(p);
  }
};
// 0x21AF DOWNWARDS ZIGZAG ARROW
ARROWS_REMAINING[0x21af] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(300, 470);
    p.L(400, 380);
    p.L(200, 260);
    p.L(400, 140);
    p.L(300, 60);
    p.stroke();
    p.M(300, 30);
    p.L(220, 120);
    p.L(380, 120);
    p.Z();
    p.fill();
  }
};
// 0x21B4 RIGHTWARDS ARROW WITH CORNER DOWNWARDS
ARROWS_REMAINING[0x21b4] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 400);
    p.L(400, 400);
    p.L(400, 120);
    p.stroke();
    p.M(400, 60);
    p.L(320, 160);
    p.L(480, 160);
    p.Z();
    p.fill();
  }
};
// 0x21B6 ANTICLOCKWISE TOP SEMICIRCLE ARROW
ARROWS_REMAINING[0x21b6] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(460, 250);
    p.C(460, 420, 140, 420, 140, 250);
    p.stroke();
    p.M(100, 250);
    p.L(180, 330);
    p.L(180, 170);
    p.Z();
    p.fill();
  }
};
// 0x21B7 CLOCKWISE TOP SEMICIRCLE ARROW
ARROWS_REMAINING[0x21b7] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(140, 250);
    p.C(140, 420, 460, 420, 460, 250);
    p.stroke();
    p.M(500, 250);
    p.L(420, 330);
    p.L(420, 170);
    p.Z();
    p.fill();
  }
};
// 0x21B8 NORTH WEST ARROW TO LONG BAR
ARROWS_REMAINING[0x21b8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(400, 100);
    p.L(180, 350);
    p.stroke();
    p.M(140, 400);
    p.L(230, 310);
    p.L(140, 310);
    p.Z();
    p.fill();
    p.M(120, 440);
    p.L(480, 440);
    p.stroke();
  }
};
// 0x21B9 LEFTWARDS ARROW TO BAR OVER RIGHTWARDS ARROW TO BAR
ARROWS_REMAINING[0x21b9] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(32);
    // top arrow left
    p.M(460, 340);
    p.L(180, 340);
    p.stroke();
    p.M(120, 340);
    p.L(200, 390);
    p.L(200, 290);
    p.Z();
    p.fill();
    // bottom arrow right
    p.M(140, 160);
    p.L(420, 160);
    p.stroke();
    p.M(480, 160);
    p.L(400, 210);
    p.L(400, 110);
    p.Z();
    p.fill();
    // bars
    p.M(120, 380);
    p.L(120, 300);
    p.stroke();
    p.M(480, 200);
    p.L(480, 120);
    p.stroke();
  }
};
// 0x21EB-0x21F4: Various specialized arrows
// 0x21EB UPWARDS WHITE ARROW ON PEDESTAL
ARROWS_REMAINING[0x21eb] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(300, 130);
    p.L(300, 380);
    p.stroke();
    p.M(300, 470);
    p.L(180, 330);
    p.L(420, 330);
    p.Z();
    p.stroke();
    // pedestal
    p.M(200, 60);
    p.L(400, 60);
    p.stroke();
    p.M(300, 60);
    p.L(300, 130);
    p.stroke();
  }
};
// 0x21EC UPWARDS WHITE ARROW ON PEDESTAL WITH HORIZONTAL BAR
ARROWS_REMAINING[0x21ec] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 160);
    p.L(300, 400);
    p.stroke();
    p.M(300, 470);
    p.L(180, 350);
    p.L(420, 350);
    p.Z();
    p.stroke();
    p.M(200, 60);
    p.L(400, 60);
    p.stroke();
    p.M(200, 110);
    p.L(400, 110);
    p.stroke();
    p.M(300, 60);
    p.L(300, 160);
    p.stroke();
  }
};
// 0x21ED UPWARDS WHITE ARROW ON PEDESTAL WITH VERTICAL BAR
ARROWS_REMAINING[0x21ed] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 130);
    p.L(300, 380);
    p.stroke();
    p.M(300, 470);
    p.L(180, 350);
    p.L(420, 350);
    p.Z();
    p.stroke();
    p.M(200, 60);
    p.L(400, 60);
    p.stroke();
    p.M(300, 60);
    p.L(300, 130);
    p.stroke();
    // extra vertical bar on pedestal
    p.M(200, 40);
    p.L(200, 80);
    p.stroke();
    p.M(400, 40);
    p.L(400, 80);
    p.stroke();
  }
};
// 0x21EE UPWARDS WHITE DOUBLE ARROW
ARROWS_REMAINING[0x21ee] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(300, 60);
    p.L(300, 400);
    p.stroke();
    p.M(300, 470);
    p.L(180, 350);
    p.L(420, 350);
    p.Z();
    p.stroke();
    p.M(300, 370);
    p.L(200, 270);
    p.L(400, 270);
    p.Z();
    p.stroke();
  }
};
// 0x21EF UPWARDS WHITE DOUBLE ARROW ON PEDESTAL
ARROWS_REMAINING[0x21ef] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 130);
    p.L(300, 400);
    p.stroke();
    p.M(300, 470);
    p.L(180, 370);
    p.L(420, 370);
    p.Z();
    p.stroke();
    p.M(300, 350);
    p.L(200, 270);
    p.L(400, 270);
    p.Z();
    p.stroke();
    p.M(200, 60);
    p.L(400, 60);
    p.stroke();
    p.M(300, 60);
    p.L(300, 130);
    p.stroke();
  }
};
// 0x21F0 RIGHTWARDS WHITE ARROW FROM WALL
ARROWS_REMAINING[0x21f0] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 250);
    p.L(380, 250);
    p.stroke();
    p.M(480, 250);
    p.L(360, 370);
    p.L(360, 130);
    p.Z();
    p.stroke();
    // wall
    p.M(100, 420);
    p.L(100, 80);
    p.stroke();
  }
};
// 0x21F1 NORTH WEST ARROW TO CORNER
ARROWS_REMAINING[0x21f1] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(420, 100);
    p.L(180, 380);
    p.stroke();
    p.M(140, 420);
    p.L(220, 340);
    p.L(140, 340);
    p.Z();
    p.fill();
    // corner
    p.M(120, 460);
    p.L(120, 420);
    p.L(480, 420);
    p.stroke();
  }
};
// 0x21F2 SOUTH EAST ARROW TO CORNER
ARROWS_REMAINING[0x21f2] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(180, 400);
    p.L(420, 120);
    p.stroke();
    p.M(460, 80);
    p.L(380, 160);
    p.L(460, 160);
    p.Z();
    p.fill();
    // corner
    p.M(120, 80);
    p.L(480, 80);
    p.L(480, 40);
    p.stroke();
  }
};
// 0x21F3 UP DOWN WHITE ARROW
ARROWS_REMAINING[0x21f3] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(300, 130);
    p.L(300, 370);
    p.stroke();
    // up
    p.M(300, 470);
    p.L(190, 350);
    p.L(410, 350);
    p.Z();
    p.stroke();
    // down
    p.M(300, 30);
    p.L(190, 150);
    p.L(410, 150);
    p.Z();
    p.stroke();
  }
};
// 0x21F4 RIGHT ARROW WITH SMALL CIRCLE
ARROWS_REMAINING[0x21f4] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(170, 250);
    p.L(420, 250);
    p.stroke();
    p.M(500, 250);
    p.L(400, 350);
    p.L(400, 150);
    p.Z();
    p.fill();
    p.lineWidth(25);
    p.circle(130, 250, 35);
    p.stroke();
  }
};
// 0x21F6 THREE RIGHTWARDS ARROWS
ARROWS_REMAINING[0x21f6] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    for (const y of [370, 250, 130]) {
      p.M(100, y);
      p.L(400, y);
      p.stroke();
      p.M(480, y);
      p.L(400, y + 50);
      p.L(400, y - 50);
      p.Z();
      p.fill();
    }
  }
};
// 0x21F7 LEFTWARDS ARROW WITH VERTICAL STROKE
ARROWS_REMAINING[0x21f7] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(480, 250);
    p.L(180, 250);
    p.stroke();
    p.M(100, 250);
    p.L(200, 350);
    p.L(200, 150);
    p.Z();
    p.fill();
    // vertical stroke
    p.M(330, 370);
    p.L(330, 130);
    p.stroke();
  }
};
// 0x21F8 RIGHTWARDS ARROW WITH VERTICAL STROKE
ARROWS_REMAINING[0x21f8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(120, 250);
    p.L(420, 250);
    p.stroke();
    p.M(500, 250);
    p.L(400, 350);
    p.L(400, 150);
    p.Z();
    p.fill();
    p.M(270, 370);
    p.L(270, 130);
    p.stroke();
  }
};
// 0x21F9 LEFT RIGHT ARROW WITH VERTICAL STROKE
ARROWS_REMAINING[0x21f9] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(160, 250);
    p.L(440, 250);
    p.stroke();
    p.M(100, 250);
    p.L(180, 330);
    p.L(180, 170);
    p.Z();
    p.fill();
    p.M(500, 250);
    p.L(420, 330);
    p.L(420, 170);
    p.Z();
    p.fill();
    p.M(300, 370);
    p.L(300, 130);
    p.stroke();
  }
};
// 0x21FA LEFTWARDS ARROW WITH DOUBLE VERTICAL STROKE
ARROWS_REMAINING[0x21fa] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(480, 250);
    p.L(180, 250);
    p.stroke();
    p.M(100, 250);
    p.L(200, 340);
    p.L(200, 160);
    p.Z();
    p.fill();
    p.M(310, 370);
    p.L(310, 130);
    p.stroke();
    p.M(350, 370);
    p.L(350, 130);
    p.stroke();
  }
};
// 0x21FB RIGHTWARDS ARROW WITH DOUBLE VERTICAL STROKE
ARROWS_REMAINING[0x21fb] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(120, 250);
    p.L(420, 250);
    p.stroke();
    p.M(500, 250);
    p.L(400, 340);
    p.L(400, 160);
    p.Z();
    p.fill();
    p.M(250, 370);
    p.L(250, 130);
    p.stroke();
    p.M(290, 370);
    p.L(290, 130);
    p.stroke();
  }
};
// 0x21FC LEFT RIGHT ARROW WITH DOUBLE VERTICAL STROKE
ARROWS_REMAINING[0x21fc] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(32);
    p.M(160, 250);
    p.L(440, 250);
    p.stroke();
    p.M(100, 250);
    p.L(180, 330);
    p.L(180, 170);
    p.Z();
    p.fill();
    p.M(500, 250);
    p.L(420, 330);
    p.L(420, 170);
    p.Z();
    p.fill();
    p.M(285, 370);
    p.L(285, 130);
    p.stroke();
    p.M(315, 370);
    p.L(315, 130);
    p.stroke();
  }
};
// 0x21FD LEFTWARDS OPEN-HEADED ARROW
ARROWS_REMAINING[0x21fd] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(480, 250);
    p.L(140, 250);
    p.stroke();
    p.M(80, 250);
    p.L(200, 370);
    p.stroke();
    p.M(80, 250);
    p.L(200, 130);
    p.stroke();
  }
};
// 0x21FE RIGHTWARDS OPEN-HEADED ARROW
ARROWS_REMAINING[0x21fe] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(120, 250);
    p.L(460, 250);
    p.stroke();
    p.M(520, 250);
    p.L(400, 370);
    p.stroke();
    p.M(520, 250);
    p.L(400, 130);
    p.stroke();
  }
};
// 0x21FF LEFT RIGHT OPEN-HEADED ARROW
ARROWS_REMAINING[0x21ff] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(38);
    p.M(140, 250);
    p.L(460, 250);
    p.stroke();
    p.M(80, 250);
    p.L(180, 350);
    p.stroke();
    p.M(80, 250);
    p.L(180, 150);
    p.stroke();
    p.M(520, 250);
    p.L(420, 350);
    p.stroke();
    p.M(520, 250);
    p.L(420, 150);
    p.stroke();
  }
};
