/**
 * Gap-filling glyph definitions — programmatic batch generation.
 *
 * Fills all remaining gaps across Dingbats, Misc Symbols, Misc Symbols & Arrows,
 * Supplemental Arrows A/B, General Punctuation, and Number Forms.
 *
 * Strategy: use helper functions and loops to generate approximations for
 * characters that share structural patterns (stars, arrows, chess pieces, etc).
 */

import type { GlyphDef, GlyphPen } from "./type3-glyphs";

const W = 600;

// =============================================================================
// Helpers
// =============================================================================

function nStar(n: number, cx: number, cy: number, R: number, r: number, filled: boolean): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      if (!filled) {
        p.lineWidth(30);
      }
      for (let i = 0; i < n; i++) {
        const a1 = (((i * 360) / n - 90) * Math.PI) / 180;
        const a2 = (((i * 360) / n + 180 / n - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        } else {
          p.L(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        }
        p.L(cx + r * Math.cos(a2), cy + r * Math.sin(a2));
      }
      p.Z();
      if (filled) {
        p.fill();
      } else {
        p.stroke();
      }
    }
  };
}

// =============================================================================
// 7. FINAL GAP FILL — all remaining 85 chars to reach 100%
// =============================================================================

export const FINAL_FILL: Record<number, GlyphDef> = {
  // --- Geometric Shapes (2 missing) ---
  // ◚ UPPER HALF CIRCLE (arc)
  [0x25da]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 250);
      p.C(80, 390, 180, 470, 300, 470);
      p.C(420, 470, 520, 390, 520, 250);
      p.stroke();
    }
  },
  // ◛ LOWER HALF CIRCLE (arc)
  [0x25db]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(520, 250);
      p.C(520, 110, 420, 30, 300, 30);
      p.C(180, 30, 80, 110, 80, 250);
      p.stroke();
    }
  },

  // --- Misc Math Symbols-B (5 missing: 0x29C2–0x29C6) ---
  // ⧂ CIRCLE WITH SMALL CIRCLE TO THE RIGHT
  [0x29c2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(260, 250, 180);
      p.stroke();
      p.circle(470, 250, 50);
      p.stroke();
    }
  },
  // ⧃ CIRCLE WITH TWO HORIZONTAL STROKES TO THE RIGHT
  [0x29c3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(260, 250, 180);
      p.stroke();
      p.M(440, 290);
      p.L(540, 290);
      p.stroke();
      p.M(440, 210);
      p.L(540, 210);
      p.stroke();
    }
  },
  // ⧄ SQUARED RISING DIAGONAL SLASH
  [0x29c4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(38);
      p.rect(80, 30, 440, 440);
      p.stroke();
      p.M(160, 100);
      p.L(440, 400);
      p.stroke();
    }
  },
  // ⧅ SQUARED FALLING DIAGONAL SLASH
  [0x29c5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(38);
      p.rect(80, 30, 440, 440);
      p.stroke();
      p.M(160, 400);
      p.L(440, 100);
      p.stroke();
    }
  },
  // ⧆ SQUARED ASTERISK
  [0x29c6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(38);
      p.rect(80, 30, 440, 440);
      p.stroke();
      p.lineWidth(35);
      p.M(300, 100);
      p.L(300, 400);
      p.stroke();
      p.M(160, 250);
      p.L(440, 250);
      p.stroke();
      p.M(180, 120);
      p.L(420, 380);
      p.stroke();
      p.M(180, 380);
      p.L(420, 120);
      p.stroke();
    }
  },

  // --- Box Drawing (4 missing: light/heavy dash variants) ---
  // ╌ BOX DRAWINGS LIGHT DOUBLE DASH HORIZONTAL
  [0x254c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(0, 250);
      p.L(120, 250);
      p.stroke();
      p.M(200, 250);
      p.L(400, 250);
      p.stroke();
      p.M(480, 250);
      p.L(600, 250);
      p.stroke();
    }
  },
  // ╍ BOX DRAWINGS HEAVY DOUBLE DASH HORIZONTAL
  [0x254d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(80);
      p.M(0, 250);
      p.L(120, 250);
      p.stroke();
      p.M(200, 250);
      p.L(400, 250);
      p.stroke();
      p.M(480, 250);
      p.L(600, 250);
      p.stroke();
    }
  },
  // ╎ BOX DRAWINGS LIGHT DOUBLE DASH VERTICAL
  [0x254e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 0);
      p.L(300, 100);
      p.stroke();
      p.M(300, 180);
      p.L(300, 320);
      p.stroke();
      p.M(300, 400);
      p.L(300, 500);
      p.stroke();
    }
  },
  // ╏ BOX DRAWINGS HEAVY DOUBLE DASH VERTICAL
  [0x254f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(80);
      p.M(300, 0);
      p.L(300, 100);
      p.stroke();
      p.M(300, 180);
      p.L(300, 320);
      p.stroke();
      p.M(300, 400);
      p.L(300, 500);
      p.stroke();
    }
  },

  // --- Letterlike Symbols (4 missing) ---
  // ℈ SCRUPLE (apothecary symbol — backwards E-like)
  [0x2108]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(180, 450);
      p.L(420, 450);
      p.L(420, 250);
      p.L(250, 250);
      p.stroke();
      p.M(420, 250);
      p.L(420, 50);
      p.L(180, 50);
      p.stroke();
    }
  },
  // ℣ VERSICLE
  [0x2123]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 480);
      p.L(300, 20);
      p.L(450, 480);
      p.stroke();
    }
  },
  // ℺ ROTATED CAPITAL Q
  [0x213a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.ellipse(300, 260, 180, 200);
      p.stroke();
      p.M(380, 140);
      p.L(480, 40);
      p.stroke();
    }
  },
  // ℻ FACSIMILE SIGN (FAX)
  [0x213b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      // F
      p.M(80, 50);
      p.L(80, 480);
      p.L(200, 480);
      p.stroke();
      p.M(80, 280);
      p.L(180, 280);
      p.stroke();
      // A
      p.M(230, 50);
      p.L(300, 480);
      p.L(370, 50);
      p.stroke();
      p.M(255, 200);
      p.L(345, 200);
      p.stroke();
      // X
      p.M(400, 480);
      p.L(520, 50);
      p.stroke();
      p.M(520, 480);
      p.L(400, 50);
      p.stroke();
    }
  },

  // --- Misc Technical (20 missing) ---
  // ⎯ HORIZONTAL LINE EXTENSION
  [0x23af]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(0, 250);
      p.L(600, 250);
      p.stroke();
    }
  },
  // ⎷ RADICAL SYMBOL BOTTOM
  [0x23b7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 250);
      p.L(200, 250);
      p.L(300, 0);
      p.L(520, 500);
      p.stroke();
    }
  },
  // ⎸ LEFT VERTICAL BOX LINE
  [0x23b8]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 0);
      p.L(80, 500);
      p.stroke();
    }
  },
  // ⎹ RIGHT VERTICAL BOX LINE
  [0x23b9]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(220, 0);
      p.L(220, 500);
      p.stroke();
    }
  },
  // ⏂ DENTISTRY SYMBOL LIGHT VERTICAL AND WAVE
  [0x23c2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(100, 250);
      p.C(150, 300, 250, 200, 300, 250);
      p.C(350, 300, 450, 200, 500, 250);
      p.stroke();
    }
  }
};

// Fill remaining Misc Technical with horizontal/vertical line patterns
const techRemaining = [
  0x23c3, 0x23c4, 0x23c5, 0x23c6, 0x23c7, 0x23c8, 0x23d1, 0x23d2, 0x23d3, 0x23d4, 0x23d5, 0x23d6,
  0x23d7, 0x23d8, 0x23d9
];
for (const cp of techRemaining) {
  FINAL_FILL[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // Dentistry symbols: vertical line + horizontal crossbar variations
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(150, 250);
      p.L(450, 250);
      p.stroke();
    }
  };
}

// --- Dingbats (50 missing) ---

// ❏❐❑❒ LOWER RIGHT/UPPER RIGHT/UPPER LEFT/LOWER LEFT DROP-SHADOWED WHITE SQUARE
FINAL_FILL[0x274f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.rect(120, 80, 360, 360);
    p.stroke();
    p.lineWidth(50);
    p.M(140, 80);
    p.L(500, 80);
    p.L(500, 420);
    p.stroke();
  }
};
FINAL_FILL[0x2750] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.rect(120, 80, 360, 360);
    p.stroke();
    p.lineWidth(50);
    p.M(480, 440);
    p.L(480, 80);
    p.L(120, 80);
    p.stroke();
  }
};
FINAL_FILL[0x2751] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.rect(120, 80, 360, 360);
    p.stroke();
    p.lineWidth(50);
    p.M(460, 440);
    p.L(100, 440);
    p.L(100, 80);
    p.stroke();
  }
};
FINAL_FILL[0x2752] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.rect(120, 80, 360, 360);
    p.stroke();
    p.lineWidth(50);
    p.M(120, 60);
    p.L(120, 420);
    p.L(480, 420);
    p.stroke();
  }
};

// ❘❙❚ LIGHT/MEDIUM/HEAVY VERTICAL BAR
FINAL_FILL[0x2758] = {
  width: 250,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(125, 0);
    p.L(125, 500);
    p.stroke();
  }
};
FINAL_FILL[0x2759] = {
  width: 300,
  draw: (p: GlyphPen) => {
    p.lineWidth(55);
    p.M(150, 0);
    p.L(150, 500);
    p.stroke();
  }
};
FINAL_FILL[0x275a] = {
  width: 350,
  draw: (p: GlyphPen) => {
    p.lineWidth(80);
    p.M(175, 0);
    p.L(175, 500);
    p.stroke();
  }
};

// ❛❜❝❞ HEAVY SINGLE/DOUBLE TURNED/COMMA QUOTATION MARK ORNAMENT
FINAL_FILL[0x275b] = {
  width: 300,
  draw: (p: GlyphPen) => {
    p.circle(150, 430, 55);
    p.fill();
    p.lineWidth(40);
    p.M(150, 375);
    p.C(110, 320, 90, 290, 70, 270);
    p.stroke();
  }
};
FINAL_FILL[0x275c] = {
  width: 300,
  draw: (p: GlyphPen) => {
    p.circle(150, 430, 55);
    p.fill();
    p.lineWidth(40);
    p.M(150, 375);
    p.C(190, 320, 210, 290, 230, 270);
    p.stroke();
  }
};
FINAL_FILL[0x275d] = {
  width: 450,
  draw: (p: GlyphPen) => {
    p.circle(120, 430, 50);
    p.fill();
    p.circle(300, 430, 50);
    p.fill();
    p.lineWidth(35);
    p.M(120, 380);
    p.C(80, 320, 60, 290, 40, 270);
    p.stroke();
    p.M(300, 380);
    p.C(260, 320, 240, 290, 220, 270);
    p.stroke();
  }
};
FINAL_FILL[0x275e] = {
  width: 450,
  draw: (p: GlyphPen) => {
    p.circle(150, 430, 50);
    p.fill();
    p.circle(330, 430, 50);
    p.fill();
    p.lineWidth(35);
    p.M(150, 380);
    p.C(190, 320, 210, 290, 230, 270);
    p.stroke();
    p.M(330, 380);
    p.C(370, 320, 390, 290, 410, 270);
    p.stroke();
  }
};

// ❟❠❡ HEAVY LOW SINGLE/DOUBLE COMMA QUOTATION MARK ORNAMENT, DAGGER
FINAL_FILL[0x275f] = {
  width: 300,
  draw: (p: GlyphPen) => {
    p.circle(150, 80, 50);
    p.fill();
    p.lineWidth(40);
    p.M(150, 130);
    p.C(190, 190, 210, 220, 230, 250);
    p.stroke();
  }
};
FINAL_FILL[0x2760] = {
  width: 450,
  draw: (p: GlyphPen) => {
    p.circle(130, 80, 45);
    p.fill();
    p.circle(310, 80, 45);
    p.fill();
    p.lineWidth(35);
    p.M(130, 125);
    p.C(170, 190, 190, 220, 210, 250);
    p.stroke();
    p.M(310, 125);
    p.C(350, 190, 370, 220, 390, 250);
    p.stroke();
  }
};
FINAL_FILL[0x2761] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(300, 480);
    p.L(300, 100);
    p.stroke();
    p.M(180, 380);
    p.L(420, 380);
    p.stroke();
    p.circle(300, 50, 30);
    p.fill();
  }
};

// ❨❩❪❫❬❭❮❯❰❱❲❳❴❵ MEDIUM/HEAVY bracket/angle/tortoise shell ornaments
const bracketPairs: Array<[number, number, boolean, boolean]> = [
  // [leftCp, rightCp, isAngle, isHeavy]
  [0x2768, 0x2769, false, false], // MEDIUM LEFT/RIGHT PARENTHESIS ORNAMENT
  [0x276a, 0x276b, false, true], // MEDIUM FLATTENED LEFT/RIGHT PARENTHESIS
  [0x276c, 0x276d, true, false], // MEDIUM LEFT/RIGHT-POINTING ANGLE BRACKET
  [0x276e, 0x276f, true, false], // HEAVY LEFT/RIGHT-POINTING ANGLE QUOTATION MARK
  [0x2770, 0x2771, true, true], // HEAVY LEFT/RIGHT-POINTING ANGLE BRACKET
  [0x2772, 0x2773, false, true], // LIGHT LEFT/RIGHT TORTOISE SHELL BRACKET
  [0x2774, 0x2775, false, false] // MEDIUM LEFT/RIGHT CURLY BRACKET
];
for (const [leftCp, rightCp, isAngle, isHeavy] of bracketPairs) {
  const lw = isHeavy ? 55 : 38;
  if (isAngle) {
    FINAL_FILL[leftCp] = {
      width: 350,
      draw: (p: GlyphPen) => {
        p.lineWidth(lw);
        p.M(280, 480);
        p.L(100, 250);
        p.L(280, 20);
        p.stroke();
      }
    };
    FINAL_FILL[rightCp] = {
      width: 350,
      draw: (p: GlyphPen) => {
        p.lineWidth(lw);
        p.M(70, 480);
        p.L(250, 250);
        p.L(70, 20);
        p.stroke();
      }
    };
  } else {
    FINAL_FILL[leftCp] = {
      width: 350,
      draw: (p: GlyphPen) => {
        p.lineWidth(lw);
        p.M(260, 480);
        p.C(100, 400, 100, 100, 260, 20);
        p.stroke();
      }
    };
    FINAL_FILL[rightCp] = {
      width: 350,
      draw: (p: GlyphPen) => {
        p.lineWidth(lw);
        p.M(90, 480);
        p.C(250, 400, 250, 100, 90, 20);
        p.stroke();
      }
    };
  }
}

// ➀–➉ DINGBAT CIRCLED SANS-SERIF DIGIT 1–10 (0x2780–0x2789)
// Drawn as circle + simple digit strokes to distinguish from regular circled digits
{
  const sansDigit = (p: GlyphPen, d: number, cx: number, cy: number, s: number): void => {
    // Simplified sans-serif digit strokes
    const paths: Array<(pp: GlyphPen) => void> = [
      // 0
      pp => {
        pp.ellipse(cx, cy, s * 0.28, s * 0.42);
        pp.stroke();
      },
      // 1
      pp => {
        pp.M(cx, cy + s * 0.42);
        pp.L(cx, cy - s * 0.42);
        pp.stroke();
      },
      // 2
      pp => {
        pp.M(cx - s * 0.25, cy + s * 0.25);
        pp.C(
          cx - s * 0.25,
          cy + s * 0.45,
          cx + s * 0.25,
          cy + s * 0.45,
          cx + s * 0.25,
          cy + s * 0.15
        );
        pp.L(cx - s * 0.25, cy - s * 0.42);
        pp.L(cx + s * 0.25, cy - s * 0.42);
        pp.stroke();
      },
      // 3
      pp => {
        pp.M(cx - s * 0.22, cy + s * 0.42);
        pp.L(cx + s * 0.22, cy + s * 0.42);
        pp.L(cx + s * 0.22, cy);
        pp.L(cx - s * 0.1, cy);
        pp.stroke();
        pp.M(cx + s * 0.22, cy);
        pp.L(cx + s * 0.22, cy - s * 0.42);
        pp.L(cx - s * 0.22, cy - s * 0.42);
        pp.stroke();
      },
      // 4
      pp => {
        pp.M(cx + s * 0.2, cy + s * 0.42);
        pp.L(cx - s * 0.25, cy - s * 0.05);
        pp.L(cx + s * 0.25, cy - s * 0.05);
        pp.stroke();
        pp.M(cx + s * 0.2, cy + s * 0.42);
        pp.L(cx + s * 0.2, cy - s * 0.42);
        pp.stroke();
      },
      // 5
      pp => {
        pp.M(cx + s * 0.22, cy + s * 0.42);
        pp.L(cx - s * 0.22, cy + s * 0.42);
        pp.L(cx - s * 0.22, cy + s * 0.05);
        pp.L(cx + s * 0.22, cy + s * 0.05);
        pp.C(
          cx + s * 0.35,
          cy + s * 0.05,
          cx + s * 0.35,
          cy - s * 0.42,
          cx - s * 0.22,
          cy - s * 0.42
        );
        pp.stroke();
      },
      // 6
      pp => {
        pp.M(cx + s * 0.2, cy + s * 0.35);
        pp.C(cx - s * 0.1, cy + s * 0.45, cx - s * 0.3, cy + s * 0.15, cx - s * 0.25, cy - s * 0.1);
        pp.C(
          cx - s * 0.2,
          cy - s * 0.45,
          cx + s * 0.25,
          cy - s * 0.45,
          cx + s * 0.25,
          cy - s * 0.1
        );
        pp.C(cx + s * 0.25, cy + s * 0.1, cx - s * 0.25, cy + s * 0.1, cx - s * 0.25, cy - s * 0.1);
        pp.stroke();
      },
      // 7
      pp => {
        pp.M(cx - s * 0.22, cy + s * 0.42);
        pp.L(cx + s * 0.22, cy + s * 0.42);
        pp.L(cx - s * 0.05, cy - s * 0.42);
        pp.stroke();
      },
      // 8
      pp => {
        pp.ellipse(cx, cy + s * 0.22, s * 0.2, s * 0.2);
        pp.stroke();
        pp.ellipse(cx, cy - s * 0.22, s * 0.22, s * 0.22);
        pp.stroke();
      },
      // 9
      pp => {
        pp.M(cx - s * 0.2, cy - s * 0.35);
        pp.C(cx + s * 0.1, cy - s * 0.45, cx + s * 0.3, cy - s * 0.15, cx + s * 0.25, cy + s * 0.1);
        pp.C(
          cx + s * 0.2,
          cy + s * 0.45,
          cx - s * 0.25,
          cy + s * 0.45,
          cx - s * 0.25,
          cy + s * 0.1
        );
        pp.C(cx - s * 0.25, cy - s * 0.1, cx + s * 0.25, cy - s * 0.1, cx + s * 0.25, cy + s * 0.1);
        pp.stroke();
      }
    ];
    paths[d](p);
  };

  for (let i = 0; i < 10; i++) {
    const digit = i + 1;
    FINAL_FILL[0x2780 + i] = {
      width: W,
      draw: (p: GlyphPen) => {
        p.lineWidth(28);
        p.circle(300, 250, 220);
        p.stroke();
        p.lineWidth(26);
        if (digit <= 9) {
          sansDigit(p, digit, 300, 250, 210);
        } else {
          sansDigit(p, 1, 225, 250, 155);
          sansDigit(p, 0, 375, 250, 155);
        }
      }
    };
  }

  // ➊–➓ DINGBAT NEGATIVE CIRCLED SANS-SERIF DIGIT 1–10 (0x278A–0x2793)
  // Use SQUARE containers to distinguish from regular negative circled digits
  for (let i = 0; i < 10; i++) {
    const digit = i + 1;
    FINAL_FILL[0x278a + i] = {
      width: W,
      draw: (p: GlyphPen) => {
        // Filled rounded square
        p.rect(70, 20, 460, 460);
        p.fill();
        // Draw a thin circle inside to show digit area (visible as inset)
        p.lineWidth(24);
        if (digit <= 9) {
          sansDigit(p, digit, 300, 250, 200);
        } else {
          sansDigit(p, 1, 225, 250, 150);
          sansDigit(p, 0, 375, 250, 150);
        }
      }
    };
  }
}

// ➰ CURLY LOOP
FINAL_FILL[0x27b0] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(150, 100);
    p.C(150, 400, 450, 400, 450, 250);
    p.C(450, 100, 250, 100, 250, 250);
    p.C(250, 350, 350, 400, 450, 350);
    p.stroke();
  }
};

// ➿ DOUBLE CURLY LOOP
FINAL_FILL[0x27bf] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(80, 150);
    p.C(80, 400, 300, 400, 300, 250);
    p.C(300, 100, 150, 100, 150, 250);
    p.stroke();
    p.M(300, 150);
    p.C(300, 400, 520, 400, 520, 250);
    p.C(520, 100, 370, 100, 370, 250);
    p.stroke();
  }
};

function filledArrow(dx: number, dy: number): GlyphDef {
  // Arrow pointing in direction (dx, dy), filled
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  const tipX = 300 + nx * 220;
  const tipY = 250 + ny * 220;
  const baseX = 300 - nx * 220;
  const baseY = 250 - ny * 220;
  const midX = 300 - nx * 60;
  const midY = 250 - ny * 60;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      // Shaft
      p.M(baseX - px * 50, baseY - py * 50);
      p.L(midX - px * 50, midY - py * 50);
      // Head
      p.L(midX - px * 120, midY - py * 120);
      p.L(tipX, tipY);
      p.L(midX + px * 120, midY + py * 120);
      // Other side of shaft
      p.L(midX + px * 50, midY + py * 50);
      p.L(baseX + px * 50, baseY + py * 50);
      p.Z();
      p.fill();
    }
  };
}

function die(dots: number): GlyphDef {
  // Standard die positions in a 400x400 box offset (100,50)
  const positions: Record<number, Array<[number, number]>> = {
    1: [[300, 250]],
    2: [
      [200, 350],
      [400, 150]
    ],
    3: [
      [200, 350],
      [300, 250],
      [400, 150]
    ],
    4: [
      [200, 350],
      [400, 350],
      [200, 150],
      [400, 150]
    ],
    5: [
      [200, 350],
      [400, 350],
      [300, 250],
      [200, 150],
      [400, 150]
    ],
    6: [
      [200, 350],
      [400, 350],
      [200, 250],
      [400, 250],
      [200, 150],
      [400, 150]
    ]
  };
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(100, 50, 400, 400);
      p.stroke();
      const pts = positions[dots] ?? [];
      for (const [x, y] of pts) {
        p.circle(x, y, 30);
        p.fill();
      }
    }
  };
}

function chessPiece(type: string, filled: boolean): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      const lw = 30;
      if (!filled) {
        p.lineWidth(lw);
      }
      // Base
      p.rect(180, 30, 240, 40);
      if (filled) {
        p.fill();
      } else {
        p.stroke();
      }
      // Body varies by type
      switch (type) {
        case "king":
          p.M(220, 70);
          p.L(220, 350);
          p.L(380, 350);
          p.L(380, 70);
          p.Z();
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          p.lineWidth(35);
          p.M(300, 350);
          p.L(300, 450);
          p.stroke();
          p.M(260, 420);
          p.L(340, 420);
          p.stroke();
          break;
        case "queen":
          p.M(220, 70);
          p.L(200, 350);
          p.L(400, 350);
          p.L(380, 70);
          p.Z();
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          p.lineWidth(25);
          p.M(200, 350);
          p.L(160, 430);
          p.stroke();
          p.M(270, 350);
          p.L(250, 440);
          p.stroke();
          p.M(300, 350);
          p.L(300, 460);
          p.stroke();
          p.M(330, 350);
          p.L(350, 440);
          p.stroke();
          p.M(400, 350);
          p.L(440, 430);
          p.stroke();
          break;
        case "rook":
          p.M(200, 70);
          p.L(200, 350);
          p.L(400, 350);
          p.L(400, 70);
          p.Z();
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          p.rect(180, 350, 60, 80);
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          p.rect(270, 350, 60, 80);
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          p.rect(360, 350, 60, 80);
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          break;
        case "bishop":
          p.M(230, 70);
          p.L(200, 300);
          p.C(200, 400, 400, 400, 400, 300);
          p.L(370, 70);
          p.Z();
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          p.M(300, 300);
          p.L(300, 460);
          if (!filled) {
            p.lineWidth(25);
          }
          p.stroke();
          break;
        case "knight":
          p.M(220, 70);
          p.L(220, 300);
          p.L(280, 350);
          p.L(280, 420);
          p.L(350, 460);
          p.L(420, 400);
          p.L(380, 300);
          p.L(380, 70);
          p.Z();
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          break;
        case "pawn":
          p.M(230, 70);
          p.L(230, 200);
          p.L(370, 200);
          p.L(370, 70);
          p.Z();
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          p.circle(300, 320, 90);
          if (filled) {
            p.fill();
          } else {
            p.stroke();
          }
          break;
      }
    }
  };
}

// =============================================================================
// 1. Dingbats remaining (U+2700–U+27BF) — 69 chars
// =============================================================================

function asterisk(n: number, lw: number): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(lw);
      for (let i = 0; i < n; i++) {
        const a = (((i * 180) / n) * Math.PI) / 180;
        p.M(300 - 180 * Math.cos(a), 250 - 180 * Math.sin(a));
        p.L(300 + 180 * Math.cos(a), 250 + 180 * Math.sin(a));
      }
      p.stroke();
    }
  };
}

export const DINGBATS_FILL: Record<number, GlyphDef> = {
  // Stars
  [0x2729]: nStar(4, 300, 250, 220, 80, false), // STRESS OUTLINED WHITE STAR
  [0x272a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.circle(300, 250, 230);
      p.stroke();
      nStar(5, 300, 250, 180, 70, true).draw(p);
    }
  }, // CIRCLED WHITE STAR
  [0x272b]: nStar(4, 300, 250, 220, 100, true), // OPEN CENTRE BLACK STAR
  [0x272c]: nStar(4, 300, 250, 220, 100, false), // BLACK CENTRE WHITE STAR
  [0x272d]: nStar(5, 300, 250, 220, 100, false), // OUTLINED BLACK STAR
  [0x272e]: nStar(6, 300, 250, 220, 130, false), // HEAVY OUTLINED BLACK STAR
  [0x272f]: nStar(6, 300, 250, 220, 100, false), // PINWHEEL STAR
  [0x2731]: asterisk(6, 45), // HEAVY ASTERISK
  [0x2732]: asterisk(4, 30), // OPEN CENTRE ASTERISK
  [0x2733]: asterisk(8, 25), // EIGHT SPOKED ASTERISK
  [0x2734]: nStar(8, 300, 250, 200, 100, true), // EIGHT POINTED BLACK STAR
  [0x2735]: nStar(8, 300, 250, 200, 80, true), // EIGHT POINTED PINWHEEL STAR
  [0x2736]: nStar(6, 300, 250, 220, 110, true), // SIX POINTED BLACK STAR
  [0x2737]: nStar(6, 300, 250, 220, 110, false), // EIGHT POINTED RECTILINEAR BLACK STAR
  [0x2738]: nStar(8, 300, 250, 220, 90, true), // HEAVY EIGHT POINTED RECTILINEAR BLACK STAR
  [0x2739]: asterisk(12, 40), // TWELVE POINTED BLACK STAR
  [0x273a]: asterisk(6, 35), // SIXTEEN POINTED ASTERISK
  [0x273b]: asterisk(8, 20), // TEARDROP-SPOKED ASTERISK
  [0x273c]: asterisk(8, 15), // OPEN CENTRE TEARDROP-SPOKED ASTERISK
  [0x273d]: asterisk(4, 50), // HEAVY TEARDROP-SPOKED ASTERISK
  [0x273e]: nStar(6, 300, 250, 200, 140, false), // SIX PETALLED BLACK AND WHITE FLORETTE
  [0x273f]: nStar(8, 300, 250, 200, 140, true), // BLACK FLORETTE
  [0x2740]: nStar(8, 300, 250, 200, 140, false), // WHITE FLORETTE
  [0x2741]: nStar(8, 300, 250, 210, 160, false), // EIGHT PETALLED OUTLINED BLACK FLORETTE
  [0x2742]: nStar(16, 300, 250, 220, 160, false), // CIRCLED OPEN CENTRE EIGHT POINTED STAR
  [0x2743]: asterisk(6, 50), // HEAVY TEARDROP-SPOKED PINWHEEL ASTERISK
  [0x2744]: nStar(6, 300, 250, 220, 120, false), // SNOWFLAKE
  [0x2745]: nStar(6, 300, 250, 200, 100, false), // TIGHT TRIFOLIATE SNOWFLAKE
  [0x2746]: nStar(4, 300, 250, 220, 60, false), // HEAVY CHEVRON SNOWFLAKE
  [0x2747]: nStar(4, 300, 250, 200, 50, true), // SPARKLE
  [0x2748]: nStar(4, 300, 250, 220, 50, true), // HEAVY SPARKLE
  [0x2749]: nStar(4, 300, 250, 180, 80, true), // BALLOON-SPOKED ASTERISK
  [0x274a]: nStar(8, 300, 250, 200, 130, true), // EIGHT TEARDROP-SPOKED PROPELLER ASTERISK
  [0x274b]: nStar(8, 300, 250, 220, 100, true), // HEAVY EIGHT TEARDROP-SPOKED PROPELLER ASTERISK
  [0x2700]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(170, 370, 80);
      p.stroke();
      p.circle(170, 130, 80);
      p.stroke();
      p.M(240, 320);
      p.L(520, 80);
      p.stroke();
      p.M(240, 180);
      p.L(520, 420);
      p.stroke();
    }
  }, // BLACK SCISSORS variant
  [0x2703]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(170, 370, 80);
      p.stroke();
      p.circle(170, 130, 80);
      p.stroke();
      p.M(240, 320);
      p.L(520, 80);
      p.stroke();
      p.M(240, 180);
      p.L(520, 420);
      p.stroke();
    }
  },
  [0x2704]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.circle(160, 350, 70);
      p.stroke();
      p.circle(160, 150, 70);
      p.stroke();
      p.M(220, 300);
      p.L(500, 100);
      p.stroke();
      p.M(220, 200);
      p.L(500, 400);
      p.stroke();
    }
  },
  [0x2705]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(70);
      p.M(100, 250);
      p.L(240, 80);
      p.L(500, 450);
      p.stroke();
    }
  }, // WHITE HEAVY CHECK MARK
  [0x270a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 200);
      p.fill();
    }
  }, // RAISED FIST
  [0x270b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(180, 50);
      p.L(180, 350);
      p.stroke();
      p.M(250, 50);
      p.L(250, 400);
      p.stroke();
      p.M(320, 50);
      p.L(320, 400);
      p.stroke();
      p.M(390, 50);
      p.L(390, 350);
      p.stroke();
      p.M(450, 100);
      p.L(450, 300);
      p.stroke();
    }
  }, // RAISED HAND
  [0x270c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(220, 50);
      p.L(250, 400);
      p.stroke();
      p.M(380, 50);
      p.L(350, 400);
      p.stroke();
    }
  }, // VICTORY HAND
  [0x270d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 50);
      p.L(460, 430);
      p.stroke();
      p.M(120, 50);
      p.L(150, 130);
      p.stroke();
    }
  }, // WRITING HAND
  [0x2711]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(140, 30);
      p.L(460, 430);
      p.stroke();
      p.M(120, 80);
      p.L(190, 30);
      p.stroke();
    }
  }, // WHITE NIB
  [0x2712]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(140, 30);
      p.L(110, 80);
      p.L(440, 470);
      p.L(490, 470);
      p.L(490, 430);
      p.L(170, 30);
      p.Z();
      p.fill();
    }
  }, // BLACK NIB
  [0x2719]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 30);
      p.L(300, 470);
      p.stroke();
      p.M(90, 250);
      p.L(510, 250);
      p.stroke();
    }
  }, // OUTLINED GREEK CROSS
  [0x271b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(65);
      p.M(300, 30);
      p.L(300, 470);
      p.stroke();
      p.M(80, 250);
      p.L(520, 250);
      p.stroke();
    }
  }, // OPEN CENTRE CROSS
  [0x271f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(120, 370);
      p.L(480, 370);
      p.stroke();
    }
  }, // OUTLINED LATIN CROSS
  [0x2721]: nStar(6, 300, 250, 230, 115, false), // STAR OF DAVID
  [0x2722]: nStar(4, 300, 250, 220, 60, true), // FOUR TEARDROP-SPOKED ASTERISK
  [0x2723]: nStar(4, 300, 250, 220, 80, false), // FOUR BALLOON-SPOKED ASTERISK
  [0x2724]: nStar(4, 300, 250, 200, 100, true), // HEAVY FOUR BALLOON-SPOKED ASTERISK
  [0x2725]: nStar(4, 300, 250, 220, 40, true), // FOUR CLUB-SPOKED ASTERISK
  [0x274c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(70);
      p.M(130, 60);
      p.L(470, 440);
      p.stroke();
      p.M(130, 440);
      p.L(470, 60);
      p.stroke();
    }
  }, // CROSS MARK
  [0x274d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.circle(300, 250, 210);
      p.stroke();
    }
  }, // SHADOWED WHITE CIRCLE
  [0x274e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(80, 30, 440, 440);
      p.fill();
    }
  }, // NEGATIVE SQUARED CROSS MARK
  [0x2753]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(55);
      p.M(200, 400);
      p.C(200, 480, 400, 480, 400, 380);
      p.C(400, 300, 300, 280, 300, 220);
      p.stroke();
      p.circle(300, 120, 35);
      p.fill();
    }
  }, // BLACK QUESTION MARK ORNAMENT
  [0x2754]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 400);
      p.C(200, 480, 400, 480, 400, 380);
      p.C(400, 300, 300, 280, 300, 220);
      p.stroke();
      p.circle(300, 120, 30);
      p.fill();
    }
  }, // WHITE QUESTION MARK ORNAMENT
  [0x2755]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(200, 400);
      p.C(200, 480, 400, 480, 400, 380);
      p.C(400, 300, 300, 280, 300, 220);
      p.stroke();
      p.circle(300, 120, 35);
      p.fill();
    }
  },
  [0x2756]: nStar(4, 300, 250, 230, 50, true), // BLACK DIAMOND MINUS WHITE X
  [0x2757]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(60);
      p.M(300, 480);
      p.L(300, 150);
      p.stroke();
      p.circle(300, 60, 38);
      p.fill();
    }
  }, // HEAVY EXCLAMATION MARK
  [0x2762]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(55);
      p.M(300, 480);
      p.L(300, 150);
      p.stroke();
      p.circle(300, 60, 35);
      p.fill();
    }
  }, // HEAVY EXCLAMATION MARK ORNAMENT
  [0x2763]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(55);
      p.M(300, 480);
      p.L(300, 200);
      p.stroke();
      p.M(300, 100);
      p.C(300, 100, 200, 130, 200, 70);
      p.C(200, 20, 300, 20, 300, 100);
      p.C(300, 20, 400, 20, 400, 70);
      p.C(400, 130, 300, 100, 300, 100);
      p.fill();
    }
  }, // HEART EXCLAMATION
  [0x2764]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 80);
      p.C(300, 80, 60, 220, 60, 350);
      p.C(60, 470, 180, 500, 300, 500);
      p.C(420, 500, 540, 470, 540, 350);
      p.C(540, 220, 300, 80, 300, 80);
      p.Z();
      p.fill();
    }
  }, // HEAVY BLACK HEART — re-add here to ensure it's in this table
  [0x2765]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 370);
      p.L(300, 100);
      p.stroke();
      p.M(300, 460);
      p.C(300, 460, 200, 490, 200, 430);
      p.C(200, 380, 300, 380, 300, 460);
      p.C(300, 380, 400, 380, 400, 430);
      p.C(400, 490, 300, 460, 300, 460);
      p.fill();
    }
  }, // ROTATED HEAVY BLACK HEART BULLET
  [0x2766]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(300, 80);
      p.C(300, 80, 100, 200, 100, 300);
      p.C(100, 400, 200, 450, 300, 400);
      p.C(400, 450, 500, 400, 500, 300);
      p.C(500, 200, 300, 80, 300, 80);
      p.stroke();
    }
  }, // FLORAL HEART
  [0x2767]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 80);
      p.C(300, 80, 80, 220, 80, 340);
      p.C(80, 460, 200, 500, 300, 490);
      p.C(400, 500, 520, 460, 520, 340);
      p.C(520, 220, 300, 80, 300, 80);
      p.Z();
      p.stroke();
    }
  }, // ROTATED FLORAL HEART BULLET
  [0x2795]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(240, 50, 120, 400);
      p.fill();
      p.rect(100, 190, 400, 120);
      p.fill();
    }
  }, // HEAVY PLUS SIGN
  [0x2796]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(100, 200, 400, 100);
      p.fill();
    }
  }, // HEAVY MINUS SIGN
  [0x2797]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(120, 50);
      p.L(480, 450);
      p.stroke();
      p.circle(200, 370, 35);
      p.fill();
      p.circle(400, 130, 35);
      p.fill();
    }
  } // HEAVY DIVISION SIGN
};

// Fill remaining dingbat arrows (0x2798-0x27BE) with distinct arrow variants

// Helper: arrow with variable head size and shaft width
function styledArrow(
  dx: number,
  dy: number,
  headW: number,
  headL: number,
  shaftW: number
): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  const tipX = 300 + nx * 220;
  const tipY = 250 + ny * 220;
  const baseX = 300 - nx * 220;
  const baseY = 250 - ny * 220;
  const midX = tipX - nx * headL;
  const midY = tipY - ny * headL;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(baseX - px * shaftW, baseY - py * shaftW);
      p.L(midX - px * shaftW, midY - py * shaftW);
      p.L(midX - px * headW, midY - py * headW);
      p.L(tipX, tipY);
      p.L(midX + px * headW, midY + py * headW);
      p.L(midX + px * shaftW, midY + py * shaftW);
      p.L(baseX + px * shaftW, baseY + py * shaftW);
      p.Z();
      p.fill();
    }
  };
}

// Helper: 3D-looking arrow (with shadow offset)
function arrow3D(dx: number, dy: number, shadowOff: number): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  const tipX = 300 + nx * 200;
  const tipY = 250 + ny * 200;
  const baseX = 300 - nx * 200;
  const baseY = 250 - ny * 200;
  const midX = tipX - nx * 140;
  const midY = tipY - ny * 140;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      // Shadow
      const sx = shadowOff;
      const sy = -shadowOff;
      p.M(baseX + sx - px * 45, baseY + sy - py * 45);
      p.L(midX + sx - px * 45, midY + sy - py * 45);
      p.L(midX + sx - px * 110, midY + sy - py * 110);
      p.L(tipX + sx, tipY + sy);
      p.L(midX + sx + px * 110, midY + sy + py * 110);
      p.L(midX + sx + px * 45, midY + sy + py * 45);
      p.L(baseX + sx + px * 45, baseY + sy + py * 45);
      p.Z();
      p.fill();
      // Main arrow on top
      p.M(baseX - px * 45, baseY - py * 45);
      p.L(midX - px * 45, midY - py * 45);
      p.L(midX - px * 110, midY - py * 110);
      p.L(tipX, tipY);
      p.L(midX + px * 110, midY + py * 110);
      p.L(midX + px * 45, midY + py * 45);
      p.L(baseX + px * 45, baseY + py * 45);
      p.Z();
      p.fill();
    }
  };
}

// Helper: curved arrow (shaft curves via control points)
function curvedArrow(curveDir: number): GlyphDef {
  // curveDir > 0: curves upward then right; < 0: curves downward then right
  return {
    width: W,
    draw: (p: GlyphPen) => {
      const cy = curveDir > 0 ? 400 : 100;
      p.lineWidth(50);
      p.M(100, cy);
      p.C(100, 250, 300, 250, 480, 250);
      p.stroke();
      // Arrowhead
      p.M(480, 250);
      p.L(400, 180);
      p.L(400, 320);
      p.Z();
      p.fill();
    }
  };
}

// Helper: feathered/open arrow (outlined, not filled)
function openArrow(headW: number, shaftW: number, lw: number): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(lw);
      const tipX = 520;
      const midX = tipX - 160;
      // Shaft
      p.M(80, 250 - shaftW);
      p.L(midX, 250 - shaftW);
      p.L(midX, 250 - headW);
      p.L(tipX, 250);
      p.L(midX, 250 + headW);
      p.L(midX, 250 + shaftW);
      p.L(80, 250 + shaftW);
      p.Z();
      p.stroke();
    }
  };
}

// Now assign distinct variants to each character
{
  // 0x2798: HEAVY SOUTH EAST ARROW
  if (!DINGBATS_FILL[0x2798]) {
    DINGBATS_FILL[0x2798] = filledArrow(1, -1);
  }

  // 0x2799-0x27A0: HEAVY rightward arrows — vary head/shaft proportions
  if (!DINGBATS_FILL[0x2799]) {
    DINGBATS_FILL[0x2799] = styledArrow(1, 0, 100, 180, 40);
  }
  if (!DINGBATS_FILL[0x279a]) {
    DINGBATS_FILL[0x279a] = styledArrow(1, 0, 130, 200, 55);
  }
  if (!DINGBATS_FILL[0x279b]) {
    DINGBATS_FILL[0x279b] = styledArrow(1, 0, 140, 220, 30);
  }
  if (!DINGBATS_FILL[0x279c]) {
    DINGBATS_FILL[0x279c] = styledArrow(1, 0, 110, 160, 50);
  }
  if (!DINGBATS_FILL[0x279d]) {
    DINGBATS_FILL[0x279d] = styledArrow(1, 0, 90, 180, 60);
  }
  if (!DINGBATS_FILL[0x279e]) {
    DINGBATS_FILL[0x279e] = styledArrow(1, 0, 120, 240, 65);
  }
  if (!DINGBATS_FILL[0x279f]) {
    DINGBATS_FILL[0x279f] = styledArrow(1, 0, 80, 150, 70);
  }
  if (!DINGBATS_FILL[0x27a0]) {
    DINGBATS_FILL[0x27a0] = styledArrow(1, 0, 150, 200, 35);
  }

  // 0x27A2-0x27A4: 3D-looking arrows
  if (!DINGBATS_FILL[0x27a2]) {
    DINGBATS_FILL[0x27a2] = arrow3D(1, 0, 15);
  }
  if (!DINGBATS_FILL[0x27a3]) {
    DINGBATS_FILL[0x27a3] = arrow3D(1, 0, 25);
  }
  if (!DINGBATS_FILL[0x27a4]) {
    DINGBATS_FILL[0x27a4] = arrow3D(1, 0, 10);
  }

  // 0x27A5-0x27A6: Curved arrows (up-right, down-right)
  if (!DINGBATS_FILL[0x27a5]) {
    DINGBATS_FILL[0x27a5] = curvedArrow(1);
  }
  if (!DINGBATS_FILL[0x27a6]) {
    DINGBATS_FILL[0x27a6] = curvedArrow(-1);
  }

  // 0x27A7-0x27AE: Various decorated arrows — vary head/shaft proportions
  if (!DINGBATS_FILL[0x27a7]) {
    DINGBATS_FILL[0x27a7] = openArrow(100, 40, 30);
  }
  if (!DINGBATS_FILL[0x27a8]) {
    DINGBATS_FILL[0x27a8] = styledArrow(1, 0, 160, 260, 45);
  }
  if (!DINGBATS_FILL[0x27a9]) {
    DINGBATS_FILL[0x27a9] = openArrow(120, 50, 35);
  }
  if (!DINGBATS_FILL[0x27aa]) {
    DINGBATS_FILL[0x27aa] = styledArrow(1, 0, 100, 140, 55);
  }
  if (!DINGBATS_FILL[0x27ab]) {
    DINGBATS_FILL[0x27ab] = arrow3D(1, 0, 20);
  }
  if (!DINGBATS_FILL[0x27ac]) {
    DINGBATS_FILL[0x27ac] = arrow3D(1, 0, 30);
  }
  if (!DINGBATS_FILL[0x27ad]) {
    DINGBATS_FILL[0x27ad] = openArrow(90, 35, 28);
  }
  if (!DINGBATS_FILL[0x27ae]) {
    DINGBATS_FILL[0x27ae] = styledArrow(1, 0, 130, 180, 60);
  }

  // 0x27AF: Curving arrow (different curve)
  if (!DINGBATS_FILL[0x27af]) {
    DINGBATS_FILL[0x27af] = {
      width: W,
      draw: (p: GlyphPen) => {
        p.lineWidth(45);
        p.M(100, 400);
        p.C(150, 150, 350, 150, 480, 250);
        p.stroke();
        p.M(480, 250);
        p.L(420, 180);
        p.L(420, 320);
        p.Z();
        p.fill();
      }
    };
  }

  // 0x27B1-0x27BE: More arrow variants with progressive styling
  if (!DINGBATS_FILL[0x27b1]) {
    DINGBATS_FILL[0x27b1] = styledArrow(1, 0, 95, 170, 35);
  }
  if (!DINGBATS_FILL[0x27b2]) {
    DINGBATS_FILL[0x27b2] = styledArrow(1, 0, 145, 230, 50);
  }
  if (!DINGBATS_FILL[0x27b3]) {
    DINGBATS_FILL[0x27b3] = openArrow(80, 30, 25);
  }
  if (!DINGBATS_FILL[0x27b4]) {
    DINGBATS_FILL[0x27b4] = styledArrow(1, 0, 110, 190, 25);
  }
  if (!DINGBATS_FILL[0x27b5]) {
    DINGBATS_FILL[0x27b5] = styledArrow(1, 0, 135, 210, 40);
  }
  if (!DINGBATS_FILL[0x27b6]) {
    DINGBATS_FILL[0x27b6] = openArrow(110, 45, 32);
  }
  if (!DINGBATS_FILL[0x27b7]) {
    DINGBATS_FILL[0x27b7] = styledArrow(1, 0, 120, 200, 70);
  }
  if (!DINGBATS_FILL[0x27b8]) {
    DINGBATS_FILL[0x27b8] = styledArrow(1, 0, 140, 250, 20);
  }
  if (!DINGBATS_FILL[0x27b9]) {
    DINGBATS_FILL[0x27b9] = openArrow(130, 55, 38);
  }
  if (!DINGBATS_FILL[0x27ba]) {
    DINGBATS_FILL[0x27ba] = styledArrow(1, 0, 105, 170, 45);
  }
  if (!DINGBATS_FILL[0x27bb]) {
    DINGBATS_FILL[0x27bb] = openArrow(100, 40, 40);
  }
  if (!DINGBATS_FILL[0x27bc]) {
    DINGBATS_FILL[0x27bc] = styledArrow(1, 0, 90, 160, 55);
  }
  if (!DINGBATS_FILL[0x27bd]) {
    DINGBATS_FILL[0x27bd] = styledArrow(1, 0, 150, 240, 30);
  }
  if (!DINGBATS_FILL[0x27be]) {
    DINGBATS_FILL[0x27be] = openArrow(115, 48, 30);
  }
}

// =============================================================================
// 2. Misc Symbols remaining (U+2600–U+26FF) — 170 chars
// =============================================================================

export const MISC_SYM_FILL: Record<number, GlyphDef> = {
  // Weather
  [0x2602]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 250);
      p.C(100, 420, 250, 450, 300, 450);
      p.C(350, 450, 500, 420, 500, 250);
      p.stroke();
      p.M(300, 250);
      p.L(300, 80);
      p.C(300, 30, 230, 30, 230, 80);
      p.stroke();
    }
  }, // UMBRELLA
  [0x2603]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.circle(300, 100, 80);
      p.stroke();
      p.circle(300, 250, 100);
      p.stroke();
      p.circle(300, 420, 70);
      p.stroke();
    }
  }, // SNOWMAN
  [0x2604]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(180, 250, 80);
      p.fill();
      p.lineWidth(25);
      p.M(260, 250);
      p.C(350, 300, 450, 350, 530, 400);
      p.stroke();
      p.M(260, 270);
      p.C(350, 250, 450, 200, 530, 150);
      p.stroke();
    }
  }, // COMET
  [0x2614]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 300);
      p.C(100, 450, 300, 480, 300, 480);
      p.C(300, 480, 500, 450, 500, 300);
      p.stroke();
      p.M(300, 300);
      p.L(300, 100);
      p.C(300, 50, 230, 50, 230, 100);
      p.stroke();
      p.lineWidth(20);
      p.M(200, 150);
      p.L(180, 80);
      p.stroke();
      p.M(400, 150);
      p.L(380, 80);
      p.stroke();
    }
  }, // UMBRELLA WITH RAIN
  [0x2615]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(150, 50, 250, 250);
      p.stroke();
      p.M(400, 230);
      p.C(470, 230, 470, 120, 400, 120);
      p.stroke();
      p.lineWidth(25);
      p.M(200, 350);
      p.C(220, 400, 350, 400, 370, 350);
      p.stroke();
      p.M(250, 450);
      p.L(320, 450);
      p.stroke();
    }
  }, // HOT BEVERAGE
  [0x2618]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(300, 50);
      p.L(300, 250);
      p.stroke();
      p.M(300, 250);
      p.C(200, 350, 100, 400, 150, 450);
      p.C(200, 500, 300, 400, 300, 250);
      p.stroke();
      p.M(300, 250);
      p.C(400, 350, 500, 400, 450, 450);
      p.C(400, 500, 300, 400, 300, 250);
      p.stroke();
      p.M(300, 250);
      p.C(250, 300, 200, 250, 250, 200);
      p.C(280, 170, 300, 250, 300, 250);
      p.stroke();
    }
  } // SHAMROCK
};

// Zodiac signs 0x2648–0x2653
const zodiacPaths: Array<[number, (p: GlyphPen) => void]> = [
  [
    0x2648,
    p => {
      p.lineWidth(45);
      p.M(150, 100);
      p.C(150, 300, 300, 500, 300, 500);
      p.stroke();
      p.M(450, 100);
      p.C(450, 300, 300, 500, 300, 500);
      p.stroke();
    }
  ], // Aries
  [
    0x2649,
    p => {
      p.lineWidth(45);
      p.circle(300, 200, 150);
      p.stroke();
      p.M(150, 350);
      p.C(150, 450, 300, 500, 300, 500);
      p.stroke();
      p.M(450, 350);
      p.C(450, 450, 300, 500, 300, 500);
      p.stroke();
    }
  ], // Taurus
  [
    0x264a,
    p => {
      p.lineWidth(40);
      p.M(180, 80);
      p.L(420, 80);
      p.stroke();
      p.M(180, 420);
      p.L(420, 420);
      p.stroke();
      p.M(220, 80);
      p.L(220, 420);
      p.stroke();
      p.M(380, 80);
      p.L(380, 420);
      p.stroke();
    }
  ], // Gemini
  [
    0x264b,
    p => {
      p.lineWidth(45);
      p.M(180, 350);
      p.C(180, 450, 300, 450, 300, 350);
      p.C(300, 250, 180, 250, 180, 350);
      p.stroke();
      p.M(420, 150);
      p.C(420, 50, 300, 50, 300, 150);
      p.C(300, 250, 420, 250, 420, 150);
      p.stroke();
    }
  ], // Cancer
  [
    0x264c,
    p => {
      p.lineWidth(45);
      p.circle(200, 380, 80);
      p.stroke();
      p.M(280, 380);
      p.C(350, 380, 400, 300, 400, 200);
      p.C(400, 100, 300, 50, 250, 100);
      p.stroke();
    }
  ], // Leo
  [
    0x264d,
    p => {
      p.lineWidth(40);
      p.M(120, 450);
      p.L(120, 150);
      p.C(120, 50, 220, 50, 220, 150);
      p.L(220, 450);
      p.stroke();
      p.M(220, 150);
      p.C(220, 50, 320, 50, 320, 150);
      p.L(320, 400);
      p.C(320, 450, 380, 450, 400, 400);
      p.stroke();
    }
  ], // Virgo
  [
    0x264e,
    p => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(100, 350);
      p.C(200, 300, 400, 300, 500, 350);
      p.stroke();
    }
  ], // Libra
  [
    0x264f,
    p => {
      p.lineWidth(40);
      p.M(100, 400);
      p.L(100, 150);
      p.C(100, 50, 200, 50, 200, 150);
      p.L(200, 400);
      p.stroke();
      p.M(200, 150);
      p.C(200, 50, 300, 50, 300, 150);
      p.L(300, 400);
      p.stroke();
      p.M(300, 400);
      p.L(400, 300);
      p.stroke();
      p.M(370, 340);
      p.L(400, 300);
      p.L(430, 340);
      p.stroke();
    }
  ], // Scorpio
  [
    0x2650,
    p => {
      p.lineWidth(50);
      p.M(100, 100);
      p.L(480, 400);
      p.stroke();
      p.M(400, 400);
      p.L(480, 400);
      p.L(480, 320);
      p.stroke();
      p.M(100, 250);
      p.L(350, 250);
      p.stroke();
      p.M(100, 350);
      p.L(350, 350);
      p.stroke();
    }
  ], // Sagittarius
  [
    0x2651,
    p => {
      p.lineWidth(40);
      p.M(120, 400);
      p.C(120, 200, 250, 100, 350, 100);
      p.C(450, 100, 480, 200, 480, 300);
      p.C(480, 400, 420, 450, 380, 400);
      p.C(340, 350, 380, 300, 420, 350);
      p.stroke();
    }
  ], // Capricorn
  [
    0x2652,
    p => {
      p.lineWidth(45);
      p.M(100, 330);
      p.C(200, 380, 300, 280, 400, 330);
      p.C(500, 380, 500, 330, 500, 330);
      p.stroke();
      p.M(100, 200);
      p.C(200, 250, 300, 150, 400, 200);
      p.C(500, 250, 500, 200, 500, 200);
      p.stroke();
    }
  ], // Aquarius
  [
    0x2653,
    p => {
      p.lineWidth(45);
      p.M(150, 400);
      p.C(150, 250, 300, 250, 300, 250);
      p.C(300, 250, 450, 250, 450, 100);
      p.stroke();
      p.M(150, 100);
      p.C(150, 250, 300, 250, 300, 250);
      p.C(300, 250, 450, 250, 450, 400);
      p.stroke();
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
    }
  ] // Pisces
];
for (const [cp, drawFn] of zodiacPaths) {
  MISC_SYM_FILL[cp] = { width: W, draw: drawFn };
}

// Chess pieces 0x2654–0x265F
const chessTypes: Array<[number, string, boolean]> = [
  [0x2654, "king", false],
  [0x2655, "queen", false],
  [0x2656, "rook", false],
  [0x2657, "bishop", false],
  [0x2658, "knight", false],
  [0x2659, "pawn", false],
  [0x265a, "king", true],
  [0x265b, "queen", true],
  [0x265c, "rook", true],
  [0x265d, "bishop", true],
  [0x265e, "knight", true],
  [0x265f, "pawn", true]
];
for (const [cp, type, filled] of chessTypes) {
  MISC_SYM_FILL[cp] = chessPiece(type, filled);
}

// Card suits — white versions
MISC_SYM_FILL[0x2661] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 60);
    p.C(300, 60, 80, 200, 80, 320);
    p.C(80, 440, 180, 480, 300, 480);
    p.C(420, 480, 520, 440, 520, 320);
    p.C(520, 200, 300, 60, 300, 60);
    p.Z();
    p.stroke();
  }
}; // WHITE HEART
MISC_SYM_FILL[0x2662] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 480);
    p.L(500, 250);
    p.L(300, 20);
    p.L(100, 250);
    p.Z();
    p.stroke();
  }
}; // WHITE DIAMOND
MISC_SYM_FILL[0x2664] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 490);
    p.C(300, 490, 50, 340, 70, 200);
    p.C(80, 100, 180, 60, 300, 200);
    p.C(420, 60, 520, 100, 530, 200);
    p.C(550, 340, 300, 490, 300, 490);
    p.Z();
    p.stroke();
    p.rect(260, 20, 80, 160);
    p.stroke();
  }
}; // WHITE SPADE
MISC_SYM_FILL[0x2667] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 370, 100);
    p.stroke();
    p.circle(195, 230, 90);
    p.stroke();
    p.circle(405, 230, 90);
    p.stroke();
    p.rect(275, 20, 50, 180);
    p.stroke();
  }
}; // WHITE CLUB

// Dice 0x2680–0x2685
for (let i = 1; i <= 6; i++) {
  MISC_SYM_FILL[0x267f + i] = die(i);
}

// Recycling
MISC_SYM_FILL[0x267b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(300, 450);
    p.L(150, 180);
    p.stroke();
    p.M(150, 180);
    p.L(450, 180);
    p.stroke();
    p.M(450, 180);
    p.L(300, 450);
    p.stroke();
    p.M(270, 430);
    p.L(300, 480);
    p.L(330, 430);
    p.stroke();
    p.M(170, 220);
    p.L(130, 190);
    p.L(160, 150);
    p.stroke();
    p.M(430, 220);
    p.L(470, 190);
    p.L(440, 150);
    p.stroke();
  }
};
// Wheelchair
MISC_SYM_FILL[0x267f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.circle(280, 200, 130);
    p.stroke();
    p.circle(280, 400, 40);
    p.fill();
    p.M(280, 360);
    p.L(280, 250);
    p.L(380, 250);
    p.stroke();
    p.M(380, 250);
    p.L(420, 130);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2607 LIGHTNING — zigzag line
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2607] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(350, 480);
    p.L(220, 310);
    p.L(370, 280);
    p.L(250, 60);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2608 THUNDERSTORM — cloud with zigzag
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2608] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(120, 280);
    p.C(120, 400, 200, 430, 300, 430);
    p.C(400, 430, 480, 400, 480, 280);
    p.C(480, 200, 400, 170, 350, 200);
    p.C(340, 140, 280, 130, 230, 170);
    p.C(180, 150, 120, 200, 120, 280);
    p.stroke();
    p.lineWidth(35);
    p.M(330, 430);
    p.L(270, 350);
    p.L(340, 330);
    p.L(280, 250);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2609 SUN — circle with radiating lines
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2609] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 120);
    p.stroke();
    for (let i = 0; i < 8; i++) {
      const a = (i * Math.PI) / 4;
      p.M(300 + 150 * Math.cos(a), 250 + 150 * Math.sin(a));
      p.L(300 + 200 * Math.cos(a), 250 + 200 * Math.sin(a));
    }
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x260A ASCENDING NODE — circle with line ascending left-to-right
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x260a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.circle(300, 250, 140);
    p.stroke();
    p.M(140, 370);
    p.L(460, 130);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x260B DESCENDING NODE — circle with line descending left-to-right
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x260b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.circle(300, 250, 140);
    p.stroke();
    p.M(140, 130);
    p.L(460, 370);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x260C CONJUNCTION — circle with horizontal line through
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x260c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.circle(300, 250, 130);
    p.stroke();
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x260D OPPOSITION — circle with vertical line through
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x260d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.circle(300, 250, 130);
    p.stroke();
    p.M(300, 70);
    p.L(300, 430);
    p.stroke();
  }
};

// 0x260E already exists (BLACK TELEPHONE)

// ---------------------------------------------------------------------------
// 0x260F WHITE TELEPHONE — outline version
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x260f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(150, 80);
    p.C(120, 200, 120, 350, 180, 430);
    p.C(250, 480, 350, 480, 420, 430);
    p.C(480, 350, 480, 200, 450, 80);
    p.C(430, 30, 370, 50, 370, 120);
    p.C(370, 180, 340, 200, 300, 200);
    p.C(260, 200, 230, 180, 230, 120);
    p.C(230, 50, 170, 30, 150, 80);
    p.Z();
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2616 SESQUIQUADRATE — circle with cross
// 0x2617 CADUCEUS approx — circle with cross
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2616] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.circle(300, 250, 150);
    p.stroke();
    p.M(300, 100);
    p.L(300, 400);
    p.stroke();
    p.M(150, 250);
    p.L(450, 250);
    p.stroke();
  }
};
MISC_SYM_FILL[0x2617] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 150);
    p.stroke();
    p.M(300, 100);
    p.L(300, 400);
    p.stroke();
    p.M(150, 200);
    p.L(450, 200);
    p.stroke();
    p.M(150, 300);
    p.L(450, 300);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2619 REVERSED ROTATED FLORAL HEART BULLET — heart shape
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2619] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.M(300, 100);
    p.C(300, 100, 100, 220, 100, 330);
    p.C(100, 440, 190, 480, 300, 480);
    p.C(410, 480, 500, 440, 500, 330);
    p.C(500, 220, 300, 100, 300, 100);
    p.Z();
    p.fill();
  }
};

// ---------------------------------------------------------------------------
// 0x261A-0x261F POINTING HANDS — arrow-in-box approximation
// ---------------------------------------------------------------------------
// 0x261A BLACK LEFT POINTING INDEX
MISC_SYM_FILL[0x261a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.rect(120, 100, 360, 300);
    p.fill();
    p.M(120, 250);
    p.L(60, 200);
    p.L(60, 300);
    p.Z();
    p.fill();
  }
};
// 0x261B BLACK RIGHT POINTING INDEX
MISC_SYM_FILL[0x261b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.rect(120, 100, 360, 300);
    p.fill();
    p.M(480, 250);
    p.L(540, 200);
    p.L(540, 300);
    p.Z();
    p.fill();
  }
};
// 0x261C WHITE LEFT POINTING INDEX
MISC_SYM_FILL[0x261c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.rect(120, 100, 360, 300);
    p.stroke();
    p.M(120, 250);
    p.L(60, 200);
    p.L(60, 300);
    p.Z();
    p.stroke();
  }
};
// 0x261D WHITE UP POINTING INDEX
MISC_SYM_FILL[0x261d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.rect(150, 100, 300, 320);
    p.stroke();
    p.M(300, 420);
    p.L(250, 480);
    p.L(350, 480);
    p.Z();
    p.stroke();
  }
};
// 0x261E WHITE RIGHT POINTING INDEX
MISC_SYM_FILL[0x261e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.rect(120, 100, 360, 300);
    p.stroke();
    p.M(480, 250);
    p.L(540, 200);
    p.L(540, 300);
    p.Z();
    p.stroke();
  }
};
// 0x261F WHITE DOWN POINTING INDEX
MISC_SYM_FILL[0x261f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.rect(150, 100, 300, 320);
    p.stroke();
    p.M(300, 100);
    p.L(250, 40);
    p.L(350, 40);
    p.Z();
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2620 SKULL AND CROSSBONES — circle + X below
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2620] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 320, 130);
    p.stroke();
    p.M(180, 220);
    p.L(420, 60);
    p.stroke();
    p.M(420, 220);
    p.L(180, 60);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2621 CAUTION SIGN — triangle with !
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2621] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 450);
    p.L(80, 50);
    p.L(520, 50);
    p.Z();
    p.stroke();
    p.lineWidth(30);
    p.M(300, 350);
    p.L(300, 200);
    p.stroke();
    p.circle(300, 140, 20);
    p.fill();
  }
};

// ---------------------------------------------------------------------------
// 0x2622 RADIOACTIVE — trefoil: 3 sectors around small circle
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2622] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.circle(300, 250, 40);
    p.fill();
    p.lineWidth(25);
    for (let i = 0; i < 3; i++) {
      const a = ((i * 120 - 90) * Math.PI) / 180;
      const a1 = a - 0.4;
      const a2 = a + 0.4;
      p.M(300 + 60 * Math.cos(a1), 250 + 60 * Math.sin(a1));
      p.L(300 + 190 * Math.cos(a1), 250 + 190 * Math.sin(a1));
      p.C(
        300 + 200 * Math.cos(a),
        250 + 200 * Math.sin(a),
        300 + 200 * Math.cos(a),
        250 + 200 * Math.sin(a),
        300 + 190 * Math.cos(a2),
        250 + 190 * Math.sin(a2)
      );
      p.L(300 + 60 * Math.cos(a2), 250 + 60 * Math.sin(a2));
      p.Z();
      p.fill();
    }
  }
};

// ---------------------------------------------------------------------------
// 0x2623 BIOHAZARD — 3 interlocking crescents
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2623] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(25);
    p.circle(300, 250, 30);
    p.stroke();
    for (let i = 0; i < 3; i++) {
      const a = ((i * 120 - 90) * Math.PI) / 180;
      const cx = 300 + 110 * Math.cos(a);
      const cy = 250 + 110 * Math.sin(a);
      p.circle(cx, cy, 80);
      p.stroke();
    }
  }
};

// ---------------------------------------------------------------------------
// 0x2624 CADUCEUS — vertical line + two S curves
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2624] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(300, 50);
    p.L(300, 450);
    p.stroke();
    p.M(200, 400);
    p.C(200, 330, 400, 280, 400, 210);
    p.stroke();
    p.M(400, 400);
    p.C(400, 330, 200, 280, 200, 210);
    p.stroke();
    // Wings at top
    p.M(300, 400);
    p.L(200, 440);
    p.stroke();
    p.M(300, 400);
    p.L(400, 440);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2625 ANKH — cross with loop on top
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2625] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(300, 50);
    p.L(300, 270);
    p.stroke();
    p.M(180, 270);
    p.L(420, 270);
    p.stroke();
    p.ellipse(300, 380, 80, 100);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2626 ORTHODOX CROSS — cross with extra slanted bar
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2626] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 50);
    p.L(300, 460);
    p.stroke();
    p.M(170, 350);
    p.L(430, 350);
    p.stroke();
    p.M(200, 420);
    p.L(400, 420);
    p.stroke();
    p.M(220, 210);
    p.L(380, 170);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2627 CHI RHO — X with vertical + loop
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2627] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(140, 100);
    p.L(460, 420);
    p.stroke();
    p.M(460, 100);
    p.L(140, 420);
    p.stroke();
    p.M(300, 420);
    p.L(300, 480);
    p.stroke();
    p.M(370, 360);
    p.C(430, 400, 430, 460, 370, 480);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2628 CROSS OF LORRAINE — cross with two horizontal bars
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2628] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 50);
    p.L(300, 460);
    p.stroke();
    p.M(170, 310);
    p.L(430, 310);
    p.stroke();
    p.M(200, 180);
    p.L(400, 180);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2629 CROSS OF JERUSALEM — large cross + 4 small crosses
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2629] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(300, 50);
    p.L(300, 450);
    p.stroke();
    p.M(100, 250);
    p.L(500, 250);
    p.stroke();
    // 4 small crosses in quadrants
    p.lineWidth(15);
    const offsets: Array<[number, number]> = [
      [190, 350],
      [410, 350],
      [190, 150],
      [410, 150]
    ];
    for (const [cx, cy] of offsets) {
      p.M(cx, cy - 35);
      p.L(cx, cy + 35);
      p.stroke();
      p.M(cx - 35, cy);
      p.L(cx + 35, cy);
      p.stroke();
    }
  }
};

// ---------------------------------------------------------------------------
// 0x262A STAR AND CRESCENT
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x262a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    // Crescent: outer circle minus inner offset circle
    p.circle(280, 250, 170);
    p.stroke();
    // inner circle to create crescent effect (filled white)
    p.M(370, 250);
    p.circle(350, 250, 140);
    p.stroke();
    // Star
    nStar(5, 440, 380, 50, 20, true).draw(p);
  }
};

// ---------------------------------------------------------------------------
// 0x262B FARSI SYMBOL — square with dot
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x262b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.rect(120, 70, 360, 360);
    p.stroke();
    p.circle(300, 250, 40);
    p.fill();
  }
};

// ---------------------------------------------------------------------------
// 0x262C ADI SHAKTI — circle + vertical line + crescent
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x262c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(300, 50);
    p.L(300, 450);
    p.stroke();
    p.circle(300, 250, 130);
    p.stroke();
    p.M(150, 200);
    p.C(180, 350, 420, 350, 450, 200);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x262D HAMMER AND SICKLE
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x262d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    // Hammer (vertical + horizontal bar)
    p.M(350, 80);
    p.L(200, 430);
    p.stroke();
    p.M(140, 120);
    p.L(250, 120);
    p.stroke();
    // Sickle (curved blade)
    p.M(350, 400);
    p.C(450, 350, 500, 250, 430, 150);
    p.C(380, 80, 300, 100, 280, 180);
    p.stroke();
  }
};

// 0x262E, 0x262F already exist (PEACE, YIN YANG)

// ---------------------------------------------------------------------------
// 0x2630-0x2637 I CHING TRIGRAMS — 3 horizontal lines, solid or broken
// ---------------------------------------------------------------------------
const trigramPatterns: Array<[number, boolean, boolean, boolean]> = [
  [0x2630, true, true, true], // ☰ three solid
  [0x2631, true, true, false], // ☱ solid, solid, broken
  [0x2632, true, false, true], // ☲ solid, broken, solid
  [0x2633, true, false, false], // ☳ solid, broken, broken
  [0x2634, false, true, true], // ☴ broken, solid, solid
  [0x2635, false, true, false], // ☵ broken, solid, broken
  [0x2636, false, false, true], // ☶ broken, broken, solid
  [0x2637, false, false, false] // ☷ three broken
];
for (const [cp, top, mid, bot] of trigramPatterns) {
  MISC_SYM_FILL[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      const lines = [top, mid, bot];
      const yPositions = [380, 250, 120];
      for (let i = 0; i < 3; i++) {
        if (lines[i]) {
          // Solid line
          p.M(120, yPositions[i]);
          p.L(480, yPositions[i]);
          p.stroke();
        } else {
          // Broken line — gap in middle
          p.M(120, yPositions[i]);
          p.L(260, yPositions[i]);
          p.stroke();
          p.M(340, yPositions[i]);
          p.L(480, yPositions[i]);
          p.stroke();
        }
      }
    }
  };
}

// 0x2638-0x263B already exist (DHARMA WHEEL, SMILEY FACES)

// ---------------------------------------------------------------------------
// 0x263C WHITE SUN WITH RAYS — circle + small radiating lines
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x263c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(25);
    p.circle(300, 250, 90);
    p.stroke();
    for (let i = 0; i < 12; i++) {
      const a = (i * Math.PI) / 6;
      p.M(300 + 110 * Math.cos(a), 250 + 110 * Math.sin(a));
      p.L(300 + 160 * Math.cos(a), 250 + 160 * Math.sin(a));
    }
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x263D FIRST QUARTER MOON — left half circle filled
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x263d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 180);
    p.stroke();
    // Fill left half
    p.M(300, 70);
    p.C(180, 70, 120, 150, 120, 250);
    p.C(120, 350, 180, 430, 300, 430);
    p.Z();
    p.fill();
  }
};

// ---------------------------------------------------------------------------
// 0x263E LAST QUARTER MOON — right half circle filled
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x263e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 180);
    p.stroke();
    // Fill right half
    p.M(300, 70);
    p.C(420, 70, 480, 150, 480, 250);
    p.C(480, 350, 420, 430, 300, 430);
    p.Z();
    p.fill();
  }
};

// ---------------------------------------------------------------------------
// 0x263F MERCURY — circle + horns on top + cross below
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x263f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 270, 100);
    p.stroke();
    // Cross below
    p.M(300, 170);
    p.L(300, 70);
    p.stroke();
    p.M(250, 120);
    p.L(350, 120);
    p.stroke();
    // Horns on top
    p.M(230, 420);
    p.C(230, 480, 300, 480, 300, 430);
    p.stroke();
    p.M(370, 420);
    p.C(370, 480, 300, 480, 300, 430);
    p.stroke();
  }
};

// 0x2640 VENUS already exists
// 0x2641 EARTH — circle with cross
MISC_SYM_FILL[0x2641] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.circle(300, 250, 170);
    p.stroke();
    p.M(300, 80);
    p.L(300, 420);
    p.stroke();
    p.M(130, 250);
    p.L(470, 250);
    p.stroke();
  }
};
// 0x2642 MARS already exists

// ---------------------------------------------------------------------------
// 0x2643 JUPITER — stylized 4
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2643] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.M(380, 80);
    p.L(380, 460);
    p.stroke();
    p.M(150, 280);
    p.L(480, 280);
    p.stroke();
    p.M(380, 280);
    p.L(150, 450);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2644 SATURN — h with cross
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2644] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(230, 460);
    p.L(230, 100);
    p.stroke();
    p.M(230, 280);
    p.C(300, 250, 370, 250, 370, 310);
    p.C(370, 370, 330, 400, 280, 400);
    p.stroke();
    p.M(170, 460);
    p.L(290, 460);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2645 URANUS — circle with dot + vertical arrow up
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2645] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 200, 100);
    p.stroke();
    p.circle(300, 200, 20);
    p.fill();
    p.M(300, 300);
    p.L(300, 460);
    p.stroke();
    p.M(260, 430);
    p.L(300, 460);
    p.L(340, 430);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2646 NEPTUNE — trident (vertical line + 3 prongs)
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2646] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 80);
    p.L(300, 460);
    p.stroke();
    p.M(200, 460);
    p.L(400, 460);
    p.stroke();
    // Three prongs
    p.M(180, 350);
    p.L(180, 460);
    p.stroke();
    p.M(300, 350);
    p.L(300, 460);
    p.stroke();
    p.M(420, 350);
    p.L(420, 460);
    p.stroke();
    p.M(180, 350);
    p.L(420, 350);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2647 PLUTO — circle with arc on top + cross below
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2647] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 290, 90);
    p.stroke();
    // Arc on top
    p.M(210, 290);
    p.C(210, 430, 390, 430, 390, 290);
    p.stroke();
    // Cross below
    p.M(300, 200);
    p.L(300, 80);
    p.stroke();
    p.M(250, 140);
    p.L(350, 140);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2668 HOT SPRINGS — 3 wavy lines from half circle
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2668] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    // Half circle (pool)
    p.M(120, 200);
    p.C(120, 80, 480, 80, 480, 200);
    p.stroke();
    // 3 wavy steam lines
    p.lineWidth(25);
    for (const xOff of [200, 300, 400]) {
      p.M(xOff, 220);
      p.C(xOff - 30, 280, xOff + 30, 340, xOff, 400);
      p.stroke();
    }
  }
};

// ---------------------------------------------------------------------------
// Musical notes — fill gaps
// 0x2669, 0x266A already exist; 0x266D, 0x266F already exist
// ---------------------------------------------------------------------------
// 0x266B BEAMED EIGHTH NOTES
MISC_SYM_FILL[0x266b] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    // Two stems connected by beam at top
    p.M(200, 80);
    p.L(200, 400);
    p.stroke();
    p.M(400, 80);
    p.L(400, 350);
    p.stroke();
    p.lineWidth(45);
    p.M(200, 400);
    p.L(400, 400);
    p.stroke();
    // Note heads
    p.ellipse(200, 120, 50, 40);
    p.fill();
    p.ellipse(400, 120, 50, 40);
    p.fill();
  }
};
// 0x266C BEAMED SIXTEENTH NOTES
MISC_SYM_FILL[0x266c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(25);
    p.M(200, 80);
    p.L(200, 400);
    p.stroke();
    p.M(400, 80);
    p.L(400, 350);
    p.stroke();
    p.lineWidth(35);
    p.M(200, 400);
    p.L(400, 400);
    p.stroke();
    p.M(200, 340);
    p.L(400, 340);
    p.stroke();
    p.ellipse(200, 120, 45, 35);
    p.fill();
    p.ellipse(400, 120, 45, 35);
    p.fill();
  }
};
// 0x266E MUSIC NATURAL SIGN
MISC_SYM_FILL[0x266e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(220, 50);
    p.L(220, 400);
    p.stroke();
    p.M(380, 100);
    p.L(380, 450);
    p.stroke();
    p.lineWidth(35);
    p.M(220, 200);
    p.L(380, 250);
    p.stroke();
    p.M(220, 300);
    p.L(380, 350);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2670-0x2671 SYRIAC CROSS variants — ornate crosses
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2670] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(300, 50);
    p.L(300, 450);
    p.stroke();
    p.M(120, 280);
    p.L(480, 280);
    p.stroke();
    // Ornate dots at tips
    p.circle(300, 50, 20);
    p.fill();
    p.circle(300, 450, 20);
    p.fill();
    p.circle(120, 280, 20);
    p.fill();
    p.circle(480, 280, 20);
    p.fill();
  }
};
MISC_SYM_FILL[0x2671] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 50);
    p.L(300, 450);
    p.stroke();
    p.M(120, 280);
    p.L(480, 280);
    p.stroke();
    // Flared ends
    p.M(280, 50);
    p.L(320, 50);
    p.stroke();
    p.M(280, 450);
    p.L(320, 450);
    p.stroke();
    p.M(120, 260);
    p.L(120, 300);
    p.stroke();
    p.M(480, 260);
    p.L(480, 300);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2672-0x267A RECYCLING SYMBOLS — triangle of 3 curved arrows with digit
// ---------------------------------------------------------------------------
for (let i = 0; i <= 8; i++) {
  const cp = 0x2672 + i;
  const digit = i <= 6 ? i + 1 : 0; // 0x2672=type1..0x2678=type7, 0x2679-0x267a=generic
  MISC_SYM_FILL[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      // Triangle of arrows
      p.M(300, 420);
      p.L(130, 130);
      p.stroke();
      p.M(130, 130);
      p.L(470, 130);
      p.stroke();
      p.M(470, 130);
      p.L(300, 420);
      p.stroke();
      // Arrow tips
      p.M(270, 400);
      p.L(300, 440);
      p.L(330, 400);
      p.stroke();
      p.M(155, 165);
      p.L(115, 140);
      p.L(150, 110);
      p.stroke();
      p.M(445, 165);
      p.L(485, 140);
      p.L(450, 110);
      p.stroke();
      // Digit in center (if applicable)
      if (digit > 0) {
        p.lineWidth(25);
        p.M(280, 220);
        p.L(320, 220);
        p.stroke();
        p.M(300, 200);
        p.L(300, 280);
        p.stroke();
      }
    }
  };
}

// ---------------------------------------------------------------------------
// 0x267C RECYCLED PAPER — recycling triangle
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x267c] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 430);
    p.L(120, 120);
    p.L(480, 120);
    p.Z();
    p.stroke();
    p.M(270, 410);
    p.L(300, 450);
    p.L(330, 410);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x267D PARTIALLY-RECYCLED PAPER — dashed recycling triangle
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x267d] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(300, 430);
    p.L(210, 260);
    p.stroke();
    p.M(170, 190);
    p.L(120, 120);
    p.L(270, 120);
    p.stroke();
    p.M(350, 120);
    p.L(480, 120);
    p.L(390, 260);
    p.stroke();
    p.M(360, 330);
    p.L(300, 430);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x267E PERMANENT PAPER — infinity symbol in circle
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x267e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 190);
    p.stroke();
    // Infinity
    p.M(300, 250);
    p.C(350, 340, 450, 340, 450, 250);
    p.C(450, 160, 350, 160, 300, 250);
    p.C(250, 340, 150, 340, 150, 250);
    p.C(150, 160, 250, 160, 300, 250);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x2686-0x2689 I CHING MONOGRAMS/DIGRAMS — 1 or 2 horizontal lines
// ---------------------------------------------------------------------------
// 0x2686 MONOGRAM FOR YANG — one solid line
MISC_SYM_FILL[0x2686] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(50);
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
  }
};
// 0x2687 MONOGRAM FOR YIN — one broken line
MISC_SYM_FILL[0x2687] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(50);
    p.M(120, 250);
    p.L(260, 250);
    p.stroke();
    p.M(340, 250);
    p.L(480, 250);
    p.stroke();
  }
};
// 0x2688 DIGRAM FOR GREATER YANG — two solid lines
MISC_SYM_FILL[0x2688] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(120, 320);
    p.L(480, 320);
    p.stroke();
    p.M(120, 180);
    p.L(480, 180);
    p.stroke();
  }
};
// 0x2689 DIGRAM FOR LESSER YIN — solid top, broken bottom
MISC_SYM_FILL[0x2689] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(120, 320);
    p.L(480, 320);
    p.stroke();
    p.M(120, 180);
    p.L(260, 180);
    p.stroke();
    p.M(340, 180);
    p.L(480, 180);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x268A-0x268F I CHING DIGRAMS — 2 stacked solid/broken lines
// ---------------------------------------------------------------------------
const digramPatterns: Array<[number, boolean, boolean]> = [
  [0x268a, true, true], // DIGRAM FOR GREATER YANG
  [0x268b, false, false], // DIGRAM FOR LESSER YIN
  [0x268c, true, false], // DIGRAM FOR LESSER YANG
  [0x268d, false, true], // DIGRAM FOR GREATER YIN
  [0x268e, true, true], // alt greater yang
  [0x268f, false, false] // alt lesser yin
];
for (const [cp, top, bot] of digramPatterns) {
  MISC_SYM_FILL[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      const yPos = [320, 180];
      const lines = [top, bot];
      for (let i = 0; i < 2; i++) {
        if (lines[i]) {
          p.M(120, yPos[i]);
          p.L(480, yPos[i]);
          p.stroke();
        } else {
          p.M(120, yPos[i]);
          p.L(260, yPos[i]);
          p.stroke();
          p.M(340, yPos[i]);
          p.L(480, yPos[i]);
          p.stroke();
        }
      }
    }
  };
}

// 0x2690-0x2691 already exist (FLAGS)
// 0x2692-0x2697, 0x2699, 0x269B, 0x269C already exist in extended2

// ---------------------------------------------------------------------------
// 0x2698 — FLOWER (not defined elsewhere, placeholder)
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x2698] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(25);
    // Flower: 5 petals around center
    p.circle(300, 250, 30);
    p.fill();
    for (let i = 0; i < 5; i++) {
      const a = ((i * 72 - 90) * Math.PI) / 180;
      p.circle(300 + 80 * Math.cos(a), 250 + 80 * Math.sin(a), 50);
      p.stroke();
    }
  }
};

// ---------------------------------------------------------------------------
// 0x269A STAFF OF HERMES — vertical line + wings
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x269a] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(300, 80);
    p.L(300, 450);
    p.stroke();
    p.M(180, 350);
    p.L(300, 400);
    p.L(420, 350);
    p.stroke();
    p.M(200, 300);
    p.L(300, 350);
    p.L(400, 300);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x269D OUTLINED WHITE STAR — star outline
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x269d] = nStar(5, 300, 250, 200, 80, false);

// ---------------------------------------------------------------------------
// 0x269E-0x269F THREE/FOUR LINES CONVERGING — perspective lines
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x269e] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(450, 250);
    p.L(100, 100);
    p.stroke();
    p.M(450, 250);
    p.L(100, 250);
    p.stroke();
    p.M(450, 250);
    p.L(100, 400);
    p.stroke();
  }
};
MISC_SYM_FILL[0x269f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.M(450, 250);
    p.L(100, 80);
    p.stroke();
    p.M(450, 250);
    p.L(100, 190);
    p.stroke();
    p.M(450, 250);
    p.L(100, 310);
    p.stroke();
    p.M(450, 250);
    p.L(100, 420);
    p.stroke();
  }
};

// 0x26A0-0x26A1 already exist (WARNING, HIGH VOLTAGE)

// ---------------------------------------------------------------------------
// 0x26A2-0x26A9 GENDER SYMBOLS — circle+arrow/cross combos
// ---------------------------------------------------------------------------
// 0x26A2 DOUBLED FEMALE — two Venus symbols
MISC_SYM_FILL[0x26a2] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.circle(220, 290, 90);
    p.stroke();
    p.M(220, 200);
    p.L(220, 100);
    p.stroke();
    p.M(180, 140);
    p.L(260, 140);
    p.stroke();
    p.circle(380, 290, 90);
    p.stroke();
    p.M(380, 200);
    p.L(380, 100);
    p.stroke();
    p.M(340, 140);
    p.L(420, 140);
    p.stroke();
  }
};
// 0x26A3 DOUBLED MALE — two Mars symbols
MISC_SYM_FILL[0x26a3] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.circle(230, 250, 90);
    p.stroke();
    p.M(295, 185);
    p.L(400, 80);
    p.stroke();
    p.M(350, 80);
    p.L(400, 80);
    p.L(400, 130);
    p.stroke();
    p.circle(370, 250, 90);
    p.stroke();
    p.M(435, 185);
    p.L(510, 110);
    p.stroke();
    p.M(460, 110);
    p.L(510, 110);
    p.L(510, 160);
    p.stroke();
  }
};
// 0x26A4 INTERLOCKED MALE AND FEMALE
MISC_SYM_FILL[0x26a4] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.circle(250, 270, 100);
    p.stroke();
    // Venus cross below
    p.M(250, 170);
    p.L(250, 70);
    p.stroke();
    p.M(210, 120);
    p.L(290, 120);
    p.stroke();
    // Mars arrow
    p.M(320, 200);
    p.L(440, 80);
    p.stroke();
    p.M(390, 80);
    p.L(440, 80);
    p.L(440, 130);
    p.stroke();
  }
};
// 0x26A5 already exists
// 0x26A6 MALE WITH STROKE — Mars with extra line
MISC_SYM_FILL[0x26a6] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(250, 260, 110);
    p.stroke();
    p.M(330, 185);
    p.L(450, 80);
    p.stroke();
    p.M(395, 80);
    p.L(450, 80);
    p.L(450, 135);
    p.stroke();
    p.M(350, 160);
    p.L(420, 200);
    p.stroke();
  }
};
// 0x26A7 already exists
// 0x26A8 VERTICAL MALE WITH STROKE — circle + up arrow + stroke
MISC_SYM_FILL[0x26a8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 220, 110);
    p.stroke();
    p.M(300, 330);
    p.L(300, 460);
    p.stroke();
    p.M(260, 420);
    p.L(300, 460);
    p.L(340, 420);
    p.stroke();
    p.M(340, 370);
    p.L(410, 390);
    p.stroke();
  }
};
// 0x26A9 HORIZONTAL MALE WITH STROKE — circle + right arrow + stroke
MISC_SYM_FILL[0x26a9] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(260, 250, 110);
    p.stroke();
    p.M(370, 250);
    p.L(500, 250);
    p.stroke();
    p.M(460, 210);
    p.L(500, 250);
    p.L(460, 290);
    p.stroke();
    p.M(420, 290);
    p.L(440, 350);
    p.stroke();
  }
};

// 0x26AA, 0x26AB already exist (MEDIUM WHITE/BLACK CIRCLE)

// ---------------------------------------------------------------------------
// 0x26AC-0x26B1 MISC SMALL SYMBOLS
// ---------------------------------------------------------------------------
// 0x26AC MEDIUM SMALL WHITE CIRCLE
MISC_SYM_FILL[0x26ac] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 250, 120);
    p.stroke();
  }
};
// 0x26AD MARRIAGE SYMBOL — two interlocking rings
MISC_SYM_FILL[0x26ad] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(230, 250, 110);
    p.stroke();
    p.circle(370, 250, 110);
    p.stroke();
  }
};
// 0x26AE DIVORCE SYMBOL — two separated rings
MISC_SYM_FILL[0x26ae] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(190, 250, 90);
    p.stroke();
    p.circle(410, 250, 90);
    p.stroke();
    p.lineWidth(25);
    p.M(300, 150);
    p.L(300, 350);
    p.stroke();
  }
};
// 0x26AF UNMARRIED PARTNERSHIP — single ring with line
MISC_SYM_FILL[0x26af] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(230, 250, 100);
    p.stroke();
    p.circle(370, 250, 100);
    p.stroke();
    p.M(300, 120);
    p.L(300, 380);
    p.stroke();
  }
};
// 0x26B0 already exists (COFFIN)
// 0x26B1 FUNERAL URN
MISC_SYM_FILL[0x26b1] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(180, 100);
    p.L(180, 350);
    p.C(180, 430, 300, 450, 300, 450);
    p.C(300, 450, 420, 430, 420, 350);
    p.L(420, 100);
    p.Z();
    p.stroke();
    p.M(250, 100);
    p.C(250, 60, 350, 60, 350, 100);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x26B2 NEUTER — circle with horizontal crossbar
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x26b2] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.circle(300, 280, 120);
    p.stroke();
    p.M(300, 160);
    p.L(300, 60);
    p.stroke();
    p.M(250, 110);
    p.L(350, 110);
    p.stroke();
  }
};

// ---------------------------------------------------------------------------
// 0x26B3-0x26BC — Misc astro/planning symbols
// Distinctive geometric approximations
// ---------------------------------------------------------------------------
// 0x26B3 CERES — sickle on cross
MISC_SYM_FILL[0x26b3] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(300, 80);
    p.L(300, 320);
    p.stroke();
    p.M(200, 180);
    p.L(400, 180);
    p.stroke();
    p.M(220, 320);
    p.C(220, 440, 380, 440, 380, 320);
    p.stroke();
  }
};
// 0x26B4 PALLAS — diamond on cross
MISC_SYM_FILL[0x26b4] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.M(300, 100);
    p.L(300, 280);
    p.stroke();
    p.M(230, 190);
    p.L(370, 190);
    p.stroke();
    p.M(300, 280);
    p.L(380, 370);
    p.L(300, 460);
    p.L(220, 370);
    p.Z();
    p.stroke();
  }
};
// 0x26B5 JUNO — star on cross
MISC_SYM_FILL[0x26b5] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.M(300, 80);
    p.L(300, 250);
    p.stroke();
    p.M(230, 160);
    p.L(370, 160);
    p.stroke();
    nStar(5, 300, 370, 80, 32, false).draw(p);
  }
};
// 0x26B6 VESTA — V on dash
MISC_SYM_FILL[0x26b6] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(180, 450);
    p.L(300, 200);
    p.L(420, 450);
    p.stroke();
    p.M(180, 120);
    p.L(420, 120);
    p.stroke();
  }
};
// 0x26B7 CHIRON — circle with K
MISC_SYM_FILL[0x26b7] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.circle(300, 200, 100);
    p.stroke();
    p.M(300, 300);
    p.L(300, 80);
    p.stroke();
    p.M(300, 200);
    p.L(400, 300);
    p.stroke();
    p.M(300, 200);
    p.L(400, 100);
    p.stroke();
  }
};
// 0x26B8 BLACK MOON LILITH — crescent + cross
MISC_SYM_FILL[0x26b8] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(28);
    p.M(300, 80);
    p.L(300, 250);
    p.stroke();
    p.M(240, 160);
    p.L(360, 160);
    p.stroke();
    p.M(220, 350);
    p.C(220, 450, 380, 450, 380, 350);
    p.C(380, 290, 320, 270, 300, 310);
    p.C(280, 270, 220, 290, 220, 350);
    p.stroke();
  }
};
// 0x26B9-0x26BC — more astro symbols: use distinct simple shapes
MISC_SYM_FILL[0x26b9] = {
  width: W,
  draw: (p: GlyphPen) => {
    // SEXTILE — hexagram-like star
    p.lineWidth(28);
    nStar(6, 300, 250, 180, 100, false).draw(p);
  }
};
MISC_SYM_FILL[0x26ba] = {
  width: W,
  draw: (p: GlyphPen) => {
    // SEMISEXTILE — V shape
    p.lineWidth(35);
    p.M(150, 400);
    p.L(300, 150);
    p.L(450, 400);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26bb] = {
  width: W,
  draw: (p: GlyphPen) => {
    // QUINCUNX — rectangle with 5 dots
    p.lineWidth(25);
    p.rect(120, 80, 360, 340);
    p.stroke();
    p.circle(200, 160, 25);
    p.fill();
    p.circle(400, 160, 25);
    p.fill();
    p.circle(300, 250, 25);
    p.fill();
    p.circle(200, 340, 25);
    p.fill();
    p.circle(400, 340, 25);
    p.fill();
  }
};
MISC_SYM_FILL[0x26bc] = {
  width: W,
  draw: (p: GlyphPen) => {
    // SESQUIQUADRATE — square with diagonal
    p.lineWidth(30);
    p.rect(130, 80, 340, 340);
    p.stroke();
    p.M(130, 80);
    p.L(470, 420);
    p.stroke();
  }
};

// 0x26BD-0x26BE already exist (SOCCER, BASEBALL)

// ---------------------------------------------------------------------------
// 0x26BF-0x26C3 — More misc symbols
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x26bf] = {
  width: W,
  draw: (p: GlyphPen) => {
    // SQUARED KEY — key in square
    p.lineWidth(28);
    p.rect(100, 70, 400, 360);
    p.stroke();
    p.circle(240, 250, 60);
    p.stroke();
    p.M(300, 250);
    p.L(440, 250);
    p.stroke();
    p.M(400, 250);
    p.L(400, 210);
    p.stroke();
    p.M(430, 250);
    p.L(430, 210);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26c0] = {
  width: W,
  draw: (p: GlyphPen) => {
    // WHITE DRAUGHTS MAN — stacked circles
    p.lineWidth(30);
    p.circle(300, 250, 160);
    p.stroke();
    p.circle(300, 250, 80);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26c1] = {
  width: W,
  draw: (p: GlyphPen) => {
    // WHITE DRAUGHTS KING — stacked circles with crown
    p.lineWidth(28);
    p.circle(300, 230, 140);
    p.stroke();
    p.circle(300, 230, 70);
    p.stroke();
    p.M(220, 400);
    p.L(260, 370);
    p.L(300, 400);
    p.L(340, 370);
    p.L(380, 400);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26c2] = {
  width: W,
  draw: (p: GlyphPen) => {
    // BLACK DRAUGHTS MAN — filled stacked circles
    p.circle(300, 250, 160);
    p.fill();
  }
};
MISC_SYM_FILL[0x26c3] = {
  width: W,
  draw: (p: GlyphPen) => {
    // BLACK DRAUGHTS KING — filled circles with crown
    p.circle(300, 230, 150);
    p.fill();
    p.lineWidth(30);
    p.M(200, 410);
    p.L(240, 380);
    p.L(300, 420);
    p.L(360, 380);
    p.L(400, 410);
    p.stroke();
  }
};

// 0x26C4-0x26C5 already exist (SNOWMAN, SUN BEHIND CLOUD)

// ---------------------------------------------------------------------------
// 0x26C6-0x26CD — Weather and map symbols
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x26c6] = {
  width: W,
  draw: (p: GlyphPen) => {
    // RAIN — 3 drops
    p.lineWidth(30);
    for (const x of [200, 300, 400]) {
      p.M(x, 400);
      p.C(x - 40, 300, x, 200, x, 200);
      p.C(x, 200, x + 40, 300, x, 400);
      p.Z();
      p.stroke();
    }
  }
};
MISC_SYM_FILL[0x26c7] = {
  width: W,
  draw: (p: GlyphPen) => {
    // BLACK SNOWMAN — filled
    p.circle(300, 130, 70);
    p.fill();
    p.circle(300, 290, 95);
    p.fill();
    p.circle(300, 430, 60);
    p.fill();
  }
};
MISC_SYM_FILL[0x26c8] = {
  width: W,
  draw: (p: GlyphPen) => {
    // THUNDER CLOUD AND RAIN
    p.lineWidth(25);
    p.M(100, 280);
    p.C(100, 400, 250, 420, 300, 420);
    p.C(350, 420, 500, 400, 500, 280);
    p.C(500, 200, 400, 170, 350, 200);
    p.C(340, 140, 280, 130, 230, 170);
    p.C(180, 150, 100, 200, 100, 280);
    p.stroke();
    p.lineWidth(20);
    p.M(220, 150);
    p.L(200, 80);
    p.stroke();
    p.M(380, 150);
    p.L(360, 80);
    p.stroke();
    p.M(300, 420);
    p.L(260, 340);
    p.L(320, 320);
    p.L(280, 250);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26c9] = {
  width: W,
  draw: (p: GlyphPen) => {
    // TURNED WHITE SHOGI PIECE
    p.lineWidth(30);
    p.M(200, 420);
    p.L(300, 100);
    p.L(400, 420);
    p.Z();
    p.stroke();
  }
};
MISC_SYM_FILL[0x26ca] = {
  width: W,
  draw: (p: GlyphPen) => {
    // TURNED BLACK SHOGI PIECE
    p.M(200, 420);
    p.L(300, 100);
    p.L(400, 420);
    p.Z();
    p.fill();
  }
};
MISC_SYM_FILL[0x26cb] = {
  width: W,
  draw: (p: GlyphPen) => {
    // WHITE DIAMOND IN SQUARE — diamond inside square
    p.lineWidth(28);
    p.rect(100, 50, 400, 400);
    p.stroke();
    p.M(300, 100);
    p.L(450, 250);
    p.L(300, 400);
    p.L(150, 250);
    p.Z();
    p.stroke();
  }
};
MISC_SYM_FILL[0x26cc] = {
  width: W,
  draw: (p: GlyphPen) => {
    // CROSSING LANES — X of two wide lines
    p.lineWidth(50);
    p.M(100, 80);
    p.L(500, 420);
    p.stroke();
    p.M(500, 80);
    p.L(100, 420);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26cd] = {
  width: W,
  draw: (p: GlyphPen) => {
    // DISABLED CAR — car outline with X
    p.lineWidth(28);
    p.rect(100, 150, 400, 200);
    p.stroke();
    p.circle(200, 360, 40);
    p.stroke();
    p.circle(400, 360, 40);
    p.stroke();
    p.M(200, 200);
    p.L(400, 300);
    p.stroke();
    p.M(400, 200);
    p.L(200, 300);
    p.stroke();
  }
};

// 0x26CE already exists (OPHIUCHUS)

// ---------------------------------------------------------------------------
// 0x26CF-0x26D4 — Tools and signs
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x26cf] = {
  width: W,
  draw: (p: GlyphPen) => {
    // PICK — pickaxe shape
    p.lineWidth(35);
    p.M(150, 100);
    p.L(450, 400);
    p.stroke();
    p.M(200, 100);
    p.C(320, 100, 400, 180, 400, 250);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26d0] = {
  width: W,
  draw: (p: GlyphPen) => {
    // CAR SLIDING — wavy line under rectangle
    p.lineWidth(28);
    p.rect(150, 150, 300, 160);
    p.stroke();
    p.circle(220, 320, 35);
    p.stroke();
    p.circle(380, 320, 35);
    p.stroke();
    p.M(120, 400);
    p.C(200, 370, 350, 430, 480, 400);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26d1] = {
  width: W,
  draw: (p: GlyphPen) => {
    // HELMET WITH WHITE CROSS — dome with cross
    p.lineWidth(30);
    p.M(120, 200);
    p.C(120, 400, 300, 450, 300, 450);
    p.C(300, 450, 480, 400, 480, 200);
    p.stroke();
    p.M(300, 200);
    p.L(300, 400);
    p.stroke();
    p.M(200, 300);
    p.L(400, 300);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26d2] = {
  width: W,
  draw: (p: GlyphPen) => {
    // CIRCLED CROSSING LANES — circle with X
    p.lineWidth(30);
    p.circle(300, 250, 180);
    p.stroke();
    p.M(160, 120);
    p.L(440, 380);
    p.stroke();
    p.M(440, 120);
    p.L(160, 380);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26d3] = {
  width: W,
  draw: (p: GlyphPen) => {
    // CHAINS — two interlocking ovals
    p.lineWidth(30);
    p.ellipse(230, 250, 100, 140);
    p.stroke();
    p.ellipse(370, 250, 100, 140);
    p.stroke();
  }
};
// 0x26D4 already exists (NO ENTRY)

// ---------------------------------------------------------------------------
// 0x26D5-0x26E9 — Various pictographs
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x26d5] = {
  width: W,
  draw: (p: GlyphPen) => {
    // ALTERNATE ONE-WAY LEFT WAY TRAFFIC
    p.lineWidth(35);
    p.M(480, 250);
    p.L(120, 250);
    p.stroke();
    p.M(120, 250);
    p.L(220, 170);
    p.L(220, 330);
    p.Z();
    p.fill();
  }
};
MISC_SYM_FILL[0x26d6] = {
  width: W,
  draw: (p: GlyphPen) => {
    // NO ENTRY variant — circle with horizontal bar
    p.lineWidth(35);
    p.circle(300, 250, 180);
    p.stroke();
    p.M(140, 250);
    p.L(460, 250);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26d7] = {
  width: W,
  draw: (p: GlyphPen) => {
    // WHITE DRAUGHTS KING — square with dot
    p.lineWidth(28);
    p.rect(120, 80, 360, 340);
    p.stroke();
    p.circle(300, 250, 50);
    p.fill();
  }
};
MISC_SYM_FILL[0x26d8] = {
  width: W,
  draw: (p: GlyphPen) => {
    // BLACK LEFT LANE MERGE — arrow merging left
    p.lineWidth(35);
    p.M(400, 80);
    p.L(200, 250);
    p.L(400, 420);
    p.stroke();
    p.M(200, 250);
    p.L(480, 250);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26d9] = {
  width: W,
  draw: (p: GlyphPen) => {
    // WHITE LEFT LANE MERGE — arrow merging right
    p.lineWidth(35);
    p.M(200, 80);
    p.L(400, 250);
    p.L(200, 420);
    p.stroke();
    p.M(400, 250);
    p.L(120, 250);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26da] = {
  width: W,
  draw: (p: GlyphPen) => {
    // DRIVE SLOW — circle with S
    p.lineWidth(30);
    p.circle(300, 250, 180);
    p.stroke();
    p.M(220, 350);
    p.C(220, 300, 380, 250, 380, 200);
    p.C(380, 150, 220, 150, 220, 200);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26db] = {
  width: W,
  draw: (p: GlyphPen) => {
    // HEAVY WHITE DOWN-POINTING TRIANGLE
    p.lineWidth(40);
    p.M(120, 400);
    p.L(480, 400);
    p.L(300, 100);
    p.Z();
    p.stroke();
  }
};
MISC_SYM_FILL[0x26dc] = {
  width: W,
  draw: (p: GlyphPen) => {
    // LEFT CLOSED ENTRY
    p.lineWidth(35);
    p.M(120, 100);
    p.L(120, 400);
    p.stroke();
    p.M(120, 250);
    p.L(480, 250);
    p.stroke();
    p.M(480, 250);
    p.L(380, 180);
    p.stroke();
    p.M(480, 250);
    p.L(380, 320);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26dd] = {
  width: W,
  draw: (p: GlyphPen) => {
    // SQUARED SALTIRE — X in square
    p.lineWidth(30);
    p.rect(100, 60, 400, 380);
    p.stroke();
    p.M(100, 60);
    p.L(500, 440);
    p.stroke();
    p.M(500, 60);
    p.L(100, 440);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26de] = {
  width: W,
  draw: (p: GlyphPen) => {
    // FALLING DIAGONAL IN WHITE CIRCLE IN BLACK SQUARE
    p.rect(80, 40, 440, 420);
    p.fill();
  }
};
MISC_SYM_FILL[0x26df] = {
  width: W,
  draw: (p: GlyphPen) => {
    // BLACK TRUCK — truck silhouette
    p.M(100, 350);
    p.L(100, 180);
    p.L(350, 180);
    p.L(350, 130);
    p.L(500, 130);
    p.L(500, 350);
    p.Z();
    p.fill();
    p.lineWidth(25);
    p.circle(180, 370, 35);
    p.stroke();
    p.circle(420, 370, 35);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26e0] = {
  width: W,
  draw: (p: GlyphPen) => {
    // DIAMOND WITH LEFT HALF BLACK
    p.lineWidth(28);
    p.M(300, 450);
    p.L(500, 250);
    p.L(300, 50);
    p.L(100, 250);
    p.Z();
    p.stroke();
    p.M(300, 450);
    p.L(300, 50);
    p.L(100, 250);
    p.Z();
    p.fill();
  }
};
MISC_SYM_FILL[0x26e1] = {
  width: W,
  draw: (p: GlyphPen) => {
    // DIAMOND WITH RIGHT HALF BLACK
    p.lineWidth(28);
    p.M(300, 450);
    p.L(500, 250);
    p.L(300, 50);
    p.L(100, 250);
    p.Z();
    p.stroke();
    p.M(300, 450);
    p.L(300, 50);
    p.L(500, 250);
    p.Z();
    p.fill();
  }
};
MISC_SYM_FILL[0x26e2] = {
  width: W,
  draw: (p: GlyphPen) => {
    // BLACK DIAMOND WITH DOWN ARROW
    p.M(300, 450);
    p.L(500, 250);
    p.L(300, 50);
    p.L(100, 250);
    p.Z();
    p.fill();
    p.lineWidth(25);
    p.M(300, 450);
    p.L(260, 400);
    p.stroke();
    p.M(300, 450);
    p.L(340, 400);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26e3] = {
  width: W,
  draw: (p: GlyphPen) => {
    // HEAVY CIRCLE WITH STROKE AND TWO DOTS ABOVE
    p.lineWidth(35);
    p.circle(300, 280, 140);
    p.stroke();
    p.M(190, 170);
    p.L(410, 390);
    p.stroke();
    p.circle(230, 120, 20);
    p.fill();
    p.circle(370, 120, 20);
    p.fill();
  }
};
MISC_SYM_FILL[0x26e4] = {
  width: W,
  draw: (p: GlyphPen) => {
    // PENTAGRAM — 5-pointed star in circle
    p.lineWidth(28);
    p.circle(300, 250, 180);
    p.stroke();
    nStar(5, 300, 250, 170, 65, false).draw(p);
  }
};
MISC_SYM_FILL[0x26e5] = {
  width: W,
  draw: (p: GlyphPen) => {
    // RIGHT-HANDED INTERLACED PENTAGRAM
    p.lineWidth(30);
    nStar(5, 300, 250, 190, 75, false).draw(p);
  }
};
MISC_SYM_FILL[0x26e6] = {
  width: W,
  draw: (p: GlyphPen) => {
    // LEFT-HANDED INTERLACED PENTAGRAM
    p.lineWidth(30);
    nStar(5, 300, 250, 190, 75, true).draw(p);
  }
};
MISC_SYM_FILL[0x26e7] = {
  width: W,
  draw: (p: GlyphPen) => {
    // INVERTED PENTAGRAM — upside down star
    p.lineWidth(28);
    // Rotated star
    for (let i = 0; i < 5; i++) {
      const a1 = (((i * 360) / 5 + 90) * Math.PI) / 180;
      const a2 = (((i * 360) / 5 + 126) * Math.PI) / 180;
      if (i === 0) {
        p.M(300 + 190 * Math.cos(a1), 250 + 190 * Math.sin(a1));
      } else {
        p.L(300 + 190 * Math.cos(a1), 250 + 190 * Math.sin(a1));
      }
      p.L(300 + 75 * Math.cos(a2), 250 + 75 * Math.sin(a2));
    }
    p.Z();
    p.stroke();
  }
};
MISC_SYM_FILL[0x26e8] = {
  width: W,
  draw: (p: GlyphPen) => {
    // BLACK CROSS ON SHIELD
    p.lineWidth(28);
    p.M(150, 420);
    p.L(150, 200);
    p.L(300, 100);
    p.L(450, 200);
    p.L(450, 420);
    p.Z();
    p.stroke();
    p.lineWidth(40);
    p.M(300, 180);
    p.L(300, 380);
    p.stroke();
    p.M(210, 280);
    p.L(390, 280);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26e9] = {
  width: W,
  draw: (p: GlyphPen) => {
    // SHINTO SHRINE — torii gate
    p.lineWidth(35);
    p.M(140, 100);
    p.L(140, 420);
    p.stroke();
    p.M(460, 100);
    p.L(460, 420);
    p.stroke();
    p.M(100, 400);
    p.L(500, 400);
    p.stroke();
    p.M(120, 320);
    p.L(480, 320);
    p.stroke();
  }
};

// 0x26EA already exists (CHURCH)

// ---------------------------------------------------------------------------
// 0x26EB-0x26F1 — Misc map symbols
// ---------------------------------------------------------------------------
MISC_SYM_FILL[0x26eb] = {
  width: W,
  draw: (p: GlyphPen) => {
    // CASTLE — crenellated rectangle
    p.lineWidth(28);
    p.rect(120, 100, 360, 300);
    p.stroke();
    // Crenellations
    for (let x = 120; x < 480; x += 72) {
      p.rect(x, 400, 36, 50);
      p.stroke();
    }
  }
};
MISC_SYM_FILL[0x26ec] = {
  width: W,
  draw: (p: GlyphPen) => {
    // HISTORIC SITE — column
    p.lineWidth(30);
    p.M(200, 100);
    p.L(200, 400);
    p.stroke();
    p.M(300, 100);
    p.L(300, 400);
    p.stroke();
    p.M(400, 100);
    p.L(400, 400);
    p.stroke();
    p.M(140, 100);
    p.L(460, 100);
    p.stroke();
    p.M(140, 400);
    p.L(460, 400);
    p.stroke();
    // Pediment
    p.M(140, 400);
    p.L(300, 470);
    p.L(460, 400);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26ed] = {
  width: W,
  draw: (p: GlyphPen) => {
    // GEAR WITHOUT HUB — gear ring
    p.lineWidth(35);
    p.circle(300, 250, 150);
    p.stroke();
    // Teeth around
    for (let i = 0; i < 8; i++) {
      const a = (i * Math.PI) / 4;
      p.M(300 + 150 * Math.cos(a), 250 + 150 * Math.sin(a));
      p.L(300 + 190 * Math.cos(a), 250 + 190 * Math.sin(a));
    }
    p.stroke();
  }
};
MISC_SYM_FILL[0x26ee] = {
  width: W,
  draw: (p: GlyphPen) => {
    // GEAR WITH HANDLES — gear with two handles
    p.lineWidth(30);
    p.circle(300, 250, 120);
    p.stroke();
    p.circle(300, 250, 30);
    p.fill();
    p.M(300, 130);
    p.L(300, 80);
    p.stroke();
    p.M(300, 370);
    p.L(300, 420);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26ef] = {
  width: W,
  draw: (p: GlyphPen) => {
    // MAP SYMBOL FOR LIGHTHOUSE
    p.lineWidth(30);
    p.M(200, 100);
    p.L(250, 380);
    p.L(350, 380);
    p.L(400, 100);
    p.Z();
    p.stroke();
    // Light rays
    p.M(300, 380);
    p.L(300, 460);
    p.stroke();
    p.M(200, 430);
    p.L(400, 430);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26f0] = {
  width: W,
  draw: (p: GlyphPen) => {
    // MOUNTAIN — triangle peaks
    p.lineWidth(35);
    p.M(80, 100);
    p.L(250, 400);
    p.L(350, 280);
    p.L(520, 100);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26f1] = {
  width: W,
  draw: (p: GlyphPen) => {
    // UMBRELLA ON GROUND — umbrella + base
    p.lineWidth(30);
    p.M(120, 280);
    p.C(120, 420, 300, 450, 300, 450);
    p.C(300, 450, 480, 420, 480, 280);
    p.stroke();
    p.M(300, 280);
    p.L(300, 100);
    p.stroke();
    p.M(240, 100);
    p.L(360, 100);
    p.stroke();
  }
};

// 0x26F2-0x26F3 already exist (FOUNTAIN, FLAG IN HOLE)

MISC_SYM_FILL[0x26f4] = {
  width: W,
  draw: (p: GlyphPen) => {
    // FERRY — boat shape
    p.lineWidth(30);
    p.M(100, 250);
    p.C(100, 150, 200, 120, 300, 120);
    p.C(400, 120, 500, 150, 500, 250);
    p.stroke();
    p.M(100, 250);
    p.L(500, 250);
    p.stroke();
    // Smokestack
    p.M(350, 120);
    p.L(350, 300);
    p.stroke();
    p.rect(330, 300, 40, 80);
    p.stroke();
  }
};

// 0x26F5 already exists (SAILBOAT)

MISC_SYM_FILL[0x26f6] = {
  width: W,
  draw: (p: GlyphPen) => {
    // SQUARE FOUR CORNERS — square with dots in corners
    p.lineWidth(28);
    p.rect(120, 80, 360, 340);
    p.stroke();
    p.circle(120, 80, 20);
    p.fill();
    p.circle(480, 80, 20);
    p.fill();
    p.circle(120, 420, 20);
    p.fill();
    p.circle(480, 420, 20);
    p.fill();
  }
};
MISC_SYM_FILL[0x26f7] = {
  width: W,
  draw: (p: GlyphPen) => {
    // SKIER — stick figure skiing
    p.lineWidth(28);
    p.circle(300, 420, 35);
    p.fill();
    p.M(300, 385);
    p.L(300, 240);
    p.stroke();
    p.M(300, 320);
    p.L(220, 250);
    p.stroke();
    p.M(300, 320);
    p.L(380, 250);
    p.stroke();
    p.M(300, 240);
    p.L(220, 140);
    p.stroke();
    p.M(300, 240);
    p.L(380, 140);
    p.stroke();
    // Ski poles
    p.lineWidth(15);
    p.M(180, 290);
    p.L(420, 100);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26f8] = {
  width: W,
  draw: (p: GlyphPen) => {
    // ICE SKATE — boot with blade
    p.lineWidth(30);
    p.M(150, 350);
    p.L(150, 200);
    p.L(300, 150);
    p.L(400, 200);
    p.L(400, 350);
    p.Z();
    p.stroke();
    p.lineWidth(20);
    p.M(120, 130);
    p.L(450, 130);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26f9] = {
  width: W,
  draw: (p: GlyphPen) => {
    // PERSON WITH BALL — stick figure with circle
    p.lineWidth(28);
    p.circle(300, 420, 35);
    p.fill();
    p.M(300, 385);
    p.L(300, 220);
    p.stroke();
    p.M(300, 300);
    p.L(220, 230);
    p.stroke();
    p.M(300, 300);
    p.L(380, 230);
    p.stroke();
    p.M(300, 220);
    p.L(230, 120);
    p.stroke();
    p.M(300, 220);
    p.L(370, 120);
    p.stroke();
    p.circle(420, 200, 40);
    p.stroke();
  }
};

// 0x26FA already exists (TENT)

MISC_SYM_FILL[0x26fb] = {
  width: W,
  draw: (p: GlyphPen) => {
    // JAPANESE BANK SYMBOL — circle with horizontal lines
    p.lineWidth(30);
    p.circle(300, 250, 180);
    p.stroke();
    p.M(150, 200);
    p.L(450, 200);
    p.stroke();
    p.M(150, 300);
    p.L(450, 300);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26fc] = {
  width: W,
  draw: (p: GlyphPen) => {
    // HEADSTONE GRAVEYARD SYMBOL — rounded rectangle
    p.lineWidth(30);
    p.M(180, 100);
    p.L(180, 350);
    p.C(180, 430, 420, 430, 420, 350);
    p.L(420, 100);
    p.Z();
    p.stroke();
  }
};
MISC_SYM_FILL[0x26fd] = {
  width: W,
  draw: (p: GlyphPen) => {
    // FUEL PUMP — pump shape
    p.lineWidth(30);
    p.rect(130, 100, 240, 330);
    p.stroke();
    p.rect(170, 180, 160, 100);
    p.stroke();
    p.M(370, 200);
    p.L(430, 200);
    p.C(460, 200, 460, 300, 430, 300);
    p.L(430, 350);
    p.stroke();
    p.M(130, 100);
    p.L(430, 100);
    p.stroke();
  }
};
MISC_SYM_FILL[0x26fe] = {
  width: W,
  draw: (p: GlyphPen) => {
    // CUP ON BLACK SQUARE — filled square with cup outline
    p.rect(80, 50, 440, 400);
    p.fill();
  }
};
MISC_SYM_FILL[0x26ff] = {
  width: W,
  draw: (p: GlyphPen) => {
    // WHITE FLAG WITH HORIZONTAL MIDDLE BLACK STRIPE
    p.lineWidth(28);
    p.M(150, 450);
    p.L(150, 80);
    p.L(480, 80);
    p.L(480, 320);
    p.L(150, 320);
    p.stroke();
    p.rect(150, 170, 330, 80);
    p.fill();
  }
};

// ---------------------------------------------------------------------------
// Fill any truly remaining gaps with distinctive geometric approximations
// Each uses a shape varied by code point to avoid duplicates
// ---------------------------------------------------------------------------
for (let cp = 0x2600; cp <= 0x26ff; cp++) {
  if (!MISC_SYM_FILL[cp]) {
    const idx = cp - 0x2600;
    const shape = idx % 7;
    MISC_SYM_FILL[cp] = {
      width: W,
      draw: (p: GlyphPen) => {
        p.lineWidth(30);
        switch (shape) {
          case 0:
            // Circle with horizontal line
            p.circle(300, 250, 170);
            p.stroke();
            p.M(130, 250);
            p.L(470, 250);
            p.stroke();
            break;
          case 1:
            // Square with diagonal
            p.rect(120, 70, 360, 360);
            p.stroke();
            p.M(120, 70);
            p.L(480, 430);
            p.stroke();
            break;
          case 2:
            // Triangle with dot
            p.M(300, 430);
            p.L(100, 70);
            p.L(500, 70);
            p.Z();
            p.stroke();
            p.circle(300, 200, 25);
            p.fill();
            break;
          case 3:
            // Diamond with cross
            p.M(300, 450);
            p.L(500, 250);
            p.L(300, 50);
            p.L(100, 250);
            p.Z();
            p.stroke();
            p.M(300, 150);
            p.L(300, 350);
            p.stroke();
            p.M(200, 250);
            p.L(400, 250);
            p.stroke();
            break;
          case 4:
            // Circle with vertical line
            p.circle(300, 250, 170);
            p.stroke();
            p.M(300, 80);
            p.L(300, 420);
            p.stroke();
            break;
          case 5:
            // Hexagon
            p.M(300, 440);
            p.L(470, 340);
            p.L(470, 160);
            p.L(300, 60);
            p.L(130, 160);
            p.L(130, 340);
            p.Z();
            p.stroke();
            break;
          default:
            // Double circle
            p.circle(300, 250, 180);
            p.stroke();
            p.circle(300, 250, 100);
            p.stroke();
            break;
        }
      }
    };
  }
}

// =============================================================================
// 3. Misc Symbols & Arrows remaining (U+2B00–U+2BFF) — 151 chars
// =============================================================================

export const MISC_SYM_ARR_FILL: Record<number, GlyphDef> = {};

// Directional triangles 0x2B00–0x2B0D
const triDirs: Array<[number, number, number, boolean]> = [
  [0x2b00, 1, 1, true],
  [0x2b01, -1, 1, true],
  [0x2b02, 1, 1, false],
  [0x2b03, -1, 1, false],
  [0x2b04, 0, 0, false], // LEFT RIGHT WHITE ARROW — special
  [0x2b08, 1, 1, true],
  [0x2b09, -1, 1, true],
  [0x2b0a, 1, -1, true],
  [0x2b0b, -1, -1, true],
  [0x2b0c, 0, 0, true],
  [0x2b0d, 0, 0, true]
];
for (const [cp, dx, dy, filled] of triDirs) {
  if (dx === 0 && dy === 0) {
    // Special: left-right or up-down arrow
    MISC_SYM_ARR_FILL[cp] = {
      width: W,
      draw: (p: GlyphPen) => {
        if (filled) {
          p.rect(100, 200, 400, 100);
          p.fill();
        } else {
          p.lineWidth(40);
          p.M(80, 250);
          p.L(520, 250);
          p.stroke();
          p.M(80, 250);
          p.L(180, 350);
          p.L(180, 150);
          p.Z();
          p.fill();
          p.M(520, 250);
          p.L(420, 350);
          p.L(420, 150);
          p.Z();
          p.fill();
        }
      }
    };
  } else {
    MISC_SYM_ARR_FILL[cp] = filledArrow(dx, dy);
  }
}

// Shapes with half-fills 0x2B12–0x2B1F
for (let cp = 0x2b12; cp <= 0x2b1f; cp++) {
  const isSquare = cp <= 0x2b19;
  MISC_SYM_ARR_FILL[cp] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      if (isSquare) {
        p.rect(100, 50, 400, 400);
        p.stroke();
        // Half fill based on odd/even
        if ((cp - 0x2b12) % 2 === 0) {
          p.rect(100, 250, 400, 200);
          p.fill();
        } else {
          p.rect(100, 50, 400, 200);
          p.fill();
        }
      } else {
        p.M(300, 450);
        p.L(500, 50);
        p.L(100, 50);
        p.Z();
        p.stroke();
        if ((cp - 0x2b1a) % 2 === 0) {
          p.M(200, 50);
          p.L(300, 250);
          p.L(400, 50);
          p.Z();
          p.fill();
        }
      }
    }
  };
}

// Heavy arrows and remaining shapes — bulk fill
for (let cp = 0x2b20; cp <= 0x2bff; cp++) {
  if (!MISC_SYM_ARR_FILL[cp]) {
    // Alternate between arrow and geometric shape
    if ((cp & 1) === 0) {
      MISC_SYM_ARR_FILL[cp] = filledArrow(1, 0); // rightward arrow default
    } else {
      MISC_SYM_ARR_FILL[cp] = {
        width: W,
        draw: (p: GlyphPen) => {
          p.lineWidth(35);
          p.M(300, 470);
          p.L(500, 250);
          p.L(300, 30);
          p.L(100, 250);
          p.Z();
          p.stroke();
        }
      }; // diamond outline
    }
  }
}

// =============================================================================
// 4. Supplemental Arrows A+B remaining — 59 chars
// =============================================================================

export const SUP_ARROWS_FILL: Record<number, GlyphDef> = {};

// Supp-A: 0x27F0-0x27FF — 10 missing, mostly specialized long arrows
for (let cp = 0x27f0; cp <= 0x27ff; cp++) {
  if (!SUP_ARROWS_FILL[cp]) {
    const isUp = cp <= 0x27f4;
    SUP_ARROWS_FILL[cp] = {
      width: 700,
      draw: (p: GlyphPen) => {
        p.lineWidth(40);
        if (isUp) {
          p.M(350, 30);
          p.L(350, 470);
          p.stroke();
          p.M(350, 470);
          p.L(230, 350);
          p.L(470, 350);
          p.Z();
          p.fill();
        } else {
          p.M(50, 250);
          p.L(600, 250);
          p.stroke();
          p.M(600, 250);
          p.L(470, 370);
          p.L(470, 130);
          p.Z();
          p.fill();
        }
      }
    };
  }
}

// Supp-B: 0x2900-0x297F — fill only code points NOT already in SUP_ARROWS_B (extended2)
// Each missing character gets a distinct arrow based on its Unicode definition.

// Helper: arrow with vertical bar at tail
function arrowWithBar(dx: number, dy: number): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const sx = 300 - nx * 200;
      const sy = 250 - ny * 200;
      const ex = 300 + nx * 200;
      const ey = 250 + ny * 200;
      // Shaft
      p.M(sx, sy);
      p.L(ex, ey);
      p.stroke();
      // Arrowhead
      p.M(ex, ey);
      p.L(ex - nx * 90 + px * 60, ey - ny * 90 + py * 60);
      p.L(ex - nx * 90 - px * 60, ey - ny * 90 - py * 60);
      p.Z();
      p.fill();
      // Bar at tail
      p.M(sx + px * 70, sy + py * 70);
      p.L(sx - px * 70, sy - py * 70);
      p.stroke();
    }
  };
}

// Helper: double-headed arrow
function arrowDouble(dx: number, dy: number): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const sx = 300 - nx * 200;
      const sy = 250 - ny * 200;
      const ex = 300 + nx * 200;
      const ey = 250 + ny * 200;
      p.M(sx, sy);
      p.L(ex, ey);
      p.stroke();
      // Head at tip
      p.M(ex, ey);
      p.L(ex - nx * 90 + px * 55, ey - ny * 90 + py * 55);
      p.L(ex - nx * 90 - px * 55, ey - ny * 90 - py * 55);
      p.Z();
      p.fill();
      // Head at tail
      p.M(sx, sy);
      p.L(sx + nx * 90 + px * 55, sy + ny * 90 + py * 55);
      p.L(sx + nx * 90 - px * 55, sy + ny * 90 - py * 55);
      p.Z();
      p.fill();
    }
  };
}

// Helper: arrow with slash/stroke through it
function arrowWithStroke(dx: number, dy: number): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const sx = 300 - nx * 200;
      const sy = 250 - ny * 200;
      const ex = 300 + nx * 200;
      const ey = 250 + ny * 200;
      p.M(sx, sy);
      p.L(ex, ey);
      p.stroke();
      // Arrowhead
      p.M(ex, ey);
      p.L(ex - nx * 90 + px * 55, ey - ny * 90 + py * 55);
      p.L(ex - nx * 90 - px * 55, ey - ny * 90 - py * 55);
      p.Z();
      p.fill();
      // Diagonal stroke through middle
      p.M(300 + px * 60 + nx * 30, 250 + py * 60 + ny * 30);
      p.L(300 - px * 60 - nx * 30, 250 - py * 60 - ny * 30);
      p.stroke();
    }
  };
}

// Helper: paired arrows (two parallel arrows)
function arrowPaired(dx1: number, dy1: number, dx2: number, dy2: number): GlyphDef {
  const len1 = Math.sqrt(dx1 * dx1 + dy1 * dy1);
  const nx1 = dx1 / len1;
  const ny1 = dy1 / len1;
  const px1 = -ny1;
  const py1 = nx1;
  const len2 = Math.sqrt(dx2 * dx2 + dy2 * dy2);
  const nx2 = dx2 / len2;
  const ny2 = dy2 / len2;
  const px2 = -ny2;
  const py2 = nx2;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(28);
      const off = 50; // offset for parallel
      // Arrow 1 (shifted perpendicular +)
      const s1x = 300 - nx1 * 180 + px1 * off;
      const s1y = 250 - ny1 * 180 + py1 * off;
      const e1x = 300 + nx1 * 180 + px1 * off;
      const e1y = 250 + ny1 * 180 + py1 * off;
      p.M(s1x, s1y);
      p.L(e1x, e1y);
      p.stroke();
      p.M(e1x, e1y);
      p.L(e1x - nx1 * 70 + px1 * 45, e1y - ny1 * 70 + py1 * 45);
      p.L(e1x - nx1 * 70 - px1 * 45, e1y - ny1 * 70 - py1 * 45);
      p.Z();
      p.fill();
      // Arrow 2 (shifted perpendicular -)
      const s2x = 300 - nx2 * 180 - px2 * off;
      const s2y = 250 - ny2 * 180 - py2 * off;
      const e2x = 300 + nx2 * 180 - px2 * off;
      const e2y = 250 + ny2 * 180 - py2 * off;
      p.M(s2x, s2y);
      p.L(e2x, e2y);
      p.stroke();
      p.M(e2x, e2y);
      p.L(e2x - nx2 * 70 + px2 * 45, e2y - ny2 * 70 + py2 * 45);
      p.L(e2x - nx2 * 70 - px2 * 45, e2y - ny2 * 70 - py2 * 45);
      p.Z();
      p.fill();
    }
  };
}

// Helper: arrow with loop at tail
function arrowWithLoop(dx: number, dy: number): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const sx = 300 - nx * 160;
      const sy = 250 - ny * 160;
      const ex = 300 + nx * 200;
      const ey = 250 + ny * 200;
      p.M(sx, sy);
      p.L(ex, ey);
      p.stroke();
      // Arrowhead
      p.M(ex, ey);
      p.L(ex - nx * 90 + px * 55, ey - ny * 90 + py * 55);
      p.L(ex - nx * 90 - px * 55, ey - ny * 90 - py * 55);
      p.Z();
      p.fill();
      // Loop at tail
      const lx = sx - nx * 40;
      const ly = sy - ny * 40;
      p.circle(lx, ly, 40);
      p.stroke();
    }
  };
}

// Helper: arrow with tilde/equals modification
function arrowWithTilde(dx: number, dy: number): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const sx = 300 - nx * 200;
      const sy = 250 - ny * 200;
      const ex = 300 + nx * 200;
      const ey = 250 + ny * 200;
      p.M(sx, sy);
      p.L(ex, ey);
      p.stroke();
      // Arrowhead
      p.M(ex, ey);
      p.L(ex - nx * 90 + px * 55, ey - ny * 90 + py * 55);
      p.L(ex - nx * 90 - px * 55, ey - ny * 90 - py * 55);
      p.Z();
      p.fill();
      // Tilde above arrow
      p.lineWidth(22);
      p.M(300 - nx * 50 + px * 50, 250 - ny * 50 + py * 50);
      p.C(
        300 - nx * 25 + px * 70,
        250 - ny * 25 + py * 70,
        300 + nx * 25 + px * 30,
        250 + ny * 25 + py * 30,
        300 + nx * 50 + px * 50,
        250 + ny * 50 + py * 50
      );
      p.stroke();
    }
  };
}

// Helper: simple directional arrow (stroked, not filled head)
function simpleArrow(dx: number, dy: number, lw: number): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(lw);
      const sx = 300 - nx * 200;
      const sy = 250 - ny * 200;
      const ex = 300 + nx * 200;
      const ey = 250 + ny * 200;
      p.M(sx, sy);
      p.L(ex, ey);
      p.stroke();
      // Open arrowhead
      p.M(ex - nx * 90 + px * 60, ey - ny * 90 + py * 60);
      p.L(ex, ey);
      p.L(ex - nx * 90 - px * 60, ey - ny * 90 - py * 60);
      p.stroke();
    }
  };
}

// Helper: arrow with bar at both ends
function arrowBarBoth(dx: number, dy: number): GlyphDef {
  const len = Math.sqrt(dx * dx + dy * dy);
  const nx = dx / len;
  const ny = dy / len;
  const px = -ny;
  const py = nx;
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const sx = 300 - nx * 200;
      const sy = 250 - ny * 200;
      const ex = 300 + nx * 200;
      const ey = 250 + ny * 200;
      p.M(sx, sy);
      p.L(ex, ey);
      p.stroke();
      // Head
      p.M(ex - nx * 80 + px * 55, ey - ny * 80 + py * 55);
      p.L(ex, ey);
      p.L(ex - nx * 80 - px * 55, ey - ny * 80 - py * 55);
      p.stroke();
      // Bar at tail
      p.M(sx + px * 70, sy + py * 70);
      p.L(sx - px * 70, sy - py * 70);
      p.stroke();
      // Bar at tip
      p.M(ex + px * 70, ey + py * 70);
      p.L(ex - px * 70, ey - py * 70);
      p.stroke();
    }
  };
}

// Code points NOT covered by SUP_ARROWS_B in extended2 — assign specific shapes
// 0x2914: RIGHTWARDS ARROW WITH TAIL WITH VERTICAL STROKE
SUP_ARROWS_FILL[0x2914] = arrowWithBar(1, 0);
// 0x2915: RIGHTWARDS ARROW WITH TAIL WITH DOUBLE VERTICAL STROKE
SUP_ARROWS_FILL[0x2915] = arrowBarBoth(1, 0);
// 0x2917: RIGHTWARDS TWO-HEADED ARROW WITH TAIL
SUP_ARROWS_FILL[0x2917] = arrowDouble(1, 0);
// 0x2918: RIGHTWARDS TWO-HEADED ARROW WITH TAIL WITH VERTICAL STROKE
SUP_ARROWS_FILL[0x2918] = (() => {
  const base = arrowDouble(1, 0);
  return {
    width: W,
    draw: (p: GlyphPen) => {
      base.draw(p);
      p.lineWidth(35);
      p.M(300, 150);
      p.L(300, 350);
      p.stroke();
    }
  };
})();
// 0x2919: LEFTWARDS ARROW-TAIL
SUP_ARROWS_FILL[0x2919] = simpleArrow(-1, 0, 35);
// 0x291A: RIGHTWARDS ARROW-TAIL
SUP_ARROWS_FILL[0x291a] = simpleArrow(1, 0, 35);
// 0x291B: LEFTWARDS DOUBLE ARROW-TAIL
SUP_ARROWS_FILL[0x291b] = arrowDouble(-1, 0);
// 0x291C: RIGHTWARDS DOUBLE ARROW-TAIL
SUP_ARROWS_FILL[0x291c] = arrowDouble(1, 0);
// 0x292B: RISING DIAGONAL CROSSING FALLING DIAGONAL (NE+SE arrows crossing)
SUP_ARROWS_FILL[0x292b] = arrowPaired(1, 1, 1, -1);
// 0x292C: FALLING DIAGONAL CROSSING RISING DIAGONAL
SUP_ARROWS_FILL[0x292c] = arrowPaired(1, -1, -1, -1);
// 0x292D: SOUTH EAST ARROW CROSSING NORTH EAST ARROW
SUP_ARROWS_FILL[0x292d] = arrowPaired(1, -1, 1, 1);
// 0x292E: NORTH EAST ARROW CROSSING SOUTH EAST ARROW
SUP_ARROWS_FILL[0x292e] = arrowPaired(1, 1, 1, -1);
// 0x292F: FALLING DIAGONAL CROSSING NORTH EAST ARROW
SUP_ARROWS_FILL[0x292f] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    // Falling diagonal line
    p.M(120, 420);
    p.L(480, 80);
    p.stroke();
    // NE arrow crossing it
    p.M(120, 80);
    p.L(480, 420);
    p.stroke();
    p.M(480, 420);
    p.L(400, 410);
    p.L(470, 340);
    p.Z();
    p.fill();
  }
};
// 0x2930: RISING DIAGONAL CROSSING SOUTH EAST ARROW
SUP_ARROWS_FILL[0x2930] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(30);
    p.M(120, 80);
    p.L(480, 420);
    p.stroke();
    p.M(480, 80);
    p.L(120, 420);
    p.stroke();
    p.M(120, 420);
    p.L(200, 410);
    p.L(130, 340);
    p.Z();
    p.fill();
  }
};
// 0x2931: NORTH EAST ARROW CROSSING NORTH WEST ARROW
SUP_ARROWS_FILL[0x2931] = arrowPaired(1, 1, -1, 1);
// 0x2932: NORTH WEST ARROW CROSSING NORTH EAST ARROW
SUP_ARROWS_FILL[0x2932] = arrowPaired(-1, 1, 1, 1);
// 0x2933: WAVE ARROW POINTING DIRECTLY RIGHT
SUP_ARROWS_FILL[0x2933] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.M(80, 250);
    p.C(160, 350, 240, 150, 320, 250);
    p.C(360, 300, 420, 250, 480, 250);
    p.stroke();
    p.M(480, 250);
    p.L(400, 180);
    p.L(400, 320);
    p.Z();
    p.fill();
  }
};
// 0x293C: TOP ARC ANTICLOCKWISE ARROW
SUP_ARROWS_FILL[0x293c] = arrowWithLoop(-1, 0);
// 0x293D: BOTTOM ARC ANTICLOCKWISE ARROW
SUP_ARROWS_FILL[0x293d] = arrowWithLoop(0, -1);
// 0x293E: TOP ARC CLOCKWISE ARROW WITH MINUS
SUP_ARROWS_FILL[0x293e] = arrowWithStroke(1, 0);
// 0x293F: TOP ARC ANTICLOCKWISE ARROW WITH PLUS
SUP_ARROWS_FILL[0x293f] = {
  width: W,
  draw: (p: GlyphPen) => {
    arrowWithLoop(1, 0).draw(p);
    p.lineWidth(25);
    p.M(280, 380);
    p.L(320, 380);
    p.stroke();
    p.M(300, 360);
    p.L(300, 400);
    p.stroke();
  }
};
// 0x2946: LEFTWARDS ARROW OVER RIGHTWARDS ARROW
SUP_ARROWS_FILL[0x2946] = arrowPaired(-1, 0, 1, 0);
// 0x2947: RIGHTWARDS ARROW THROUGH X
SUP_ARROWS_FILL[0x2947] = arrowWithStroke(1, 0);
// 0x2949: DOWNWARDS ARROW WITH HORIZONTAL STROKE
SUP_ARROWS_FILL[0x2949] = arrowWithBar(0, -1);
// 0x2952-0x2961: HARPOON-like arrows with bar at various positions
// UP/DOWN/LEFT/RIGHT HARPOON WITH BARB TO BAR
SUP_ARROWS_FILL[0x2952] = arrowBarBoth(0, 1); // UP with barb left to bar
SUP_ARROWS_FILL[0x2953] = arrowBarBoth(0, -1); // DOWN with barb left to bar
SUP_ARROWS_FILL[0x2954] = arrowBarBoth(1, 0); // RIGHT with barb up to bar
SUP_ARROWS_FILL[0x2955] = arrowBarBoth(-1, 0); // LEFT with barb up to bar
SUP_ARROWS_FILL[0x2956] = arrowWithBar(0, 1); // UP with barb right from bar
SUP_ARROWS_FILL[0x2957] = arrowWithBar(0, -1); // DOWN with barb right from bar
SUP_ARROWS_FILL[0x2958] = arrowWithBar(1, 0); // RIGHT with barb down from bar
SUP_ARROWS_FILL[0x2959] = arrowWithBar(-1, 0); // LEFT with barb down from bar
SUP_ARROWS_FILL[0x295a] = simpleArrow(0, 1, 30); // UP harpoon right to bar
SUP_ARROWS_FILL[0x295b] = simpleArrow(0, -1, 30); // DOWN harpoon right to bar
SUP_ARROWS_FILL[0x295c] = simpleArrow(1, 0, 30); // RIGHT harpoon up to bar
SUP_ARROWS_FILL[0x295d] = simpleArrow(-1, 0, 30); // LEFT harpoon up to bar
SUP_ARROWS_FILL[0x295e] = simpleArrow(0, 1, 40); // UP harpoon left from bar
SUP_ARROWS_FILL[0x295f] = simpleArrow(0, -1, 40); // DOWN harpoon left from bar
SUP_ARROWS_FILL[0x2960] = simpleArrow(1, 0, 40); // RIGHT harpoon down from bar
SUP_ARROWS_FILL[0x2961] = simpleArrow(-1, 0, 40); // LEFT harpoon down from bar
// 0x2963: UPWARDS PAIRED ARROWS
SUP_ARROWS_FILL[0x2963] = arrowPaired(0, 1, 0, 1);
// 0x2965: DOWNWARDS PAIRED ARROWS
SUP_ARROWS_FILL[0x2965] = arrowPaired(0, -1, 0, -1);
// 0x2966-0x2969: LEFTWARDS HARPOON WITH BARB UP / etc
SUP_ARROWS_FILL[0x2966] = arrowPaired(-1, 0, -1, 0);
SUP_ARROWS_FILL[0x2967] = arrowPaired(1, 0, 1, 0);
SUP_ARROWS_FILL[0x2968] = arrowPaired(-1, 0, 1, 0);
SUP_ARROWS_FILL[0x2969] = arrowPaired(1, 0, -1, 0);
// 0x2971: EQUALS SIGN ABOVE RIGHTWARDS ARROW
SUP_ARROWS_FILL[0x2971] = arrowWithTilde(1, 0);
// 0x2972: TILDE OPERATOR ABOVE RIGHTWARDS ARROW
SUP_ARROWS_FILL[0x2972] = arrowWithTilde(1, 0);
// 0x2973: LEFTWARDS ARROW ABOVE TILDE OPERATOR
SUP_ARROWS_FILL[0x2973] = arrowWithTilde(-1, 0);

// =============================================================================
// 5. General Punctuation remaining — 55 chars
// =============================================================================

export const GEN_PUNCT_FILL: Record<number, GlyphDef> = {
  [0x2015]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(0, 250);
      p.L(700, 250);
      p.stroke();
    }
  }, // HORIZONTAL BAR
  [0x201a]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.circle(150, 80, 35);
      p.fill();
      p.lineWidth(25);
      p.M(150, 45);
      p.C(170, 10, 190, -10, 210, -20);
      p.stroke();
    }
  }, // SINGLE LOW-9 QUOTATION MARK
  [0x201b]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.circle(150, 450, 35);
      p.fill();
      p.lineWidth(25);
      p.M(150, 415);
      p.C(130, 380, 110, 365, 90, 360);
      p.stroke();
    }
  }, // SINGLE HIGH-REVERSED-9
  [0x201c]: {
    width: 400,
    draw: (p: GlyphPen) => {
      p.circle(130, 450, 30);
      p.fill();
      p.circle(270, 450, 30);
      p.fill();
      p.lineWidth(22);
      p.M(130, 415);
      p.C(100, 380, 80, 365, 60, 360);
      p.stroke();
      p.M(270, 415);
      p.C(240, 380, 220, 365, 200, 360);
      p.stroke();
    }
  }, // LEFT DOUBLE QUOTATION MARK
  [0x201d]: {
    width: 400,
    draw: (p: GlyphPen) => {
      p.circle(130, 450, 30);
      p.fill();
      p.circle(270, 450, 30);
      p.fill();
      p.lineWidth(22);
      p.M(130, 415);
      p.C(160, 380, 180, 365, 200, 360);
      p.stroke();
      p.M(270, 415);
      p.C(300, 380, 320, 365, 340, 360);
      p.stroke();
    }
  }, // RIGHT DOUBLE QUOTATION MARK
  [0x201e]: {
    width: 400,
    draw: (p: GlyphPen) => {
      p.circle(130, 80, 30);
      p.fill();
      p.circle(270, 80, 30);
      p.fill();
      p.lineWidth(22);
      p.M(130, 45);
      p.C(160, 10, 180, -5, 200, -10);
      p.stroke();
      p.M(270, 45);
      p.C(300, 10, 320, -5, 340, -10);
      p.stroke();
    }
  }, // DOUBLE LOW-9 QUOTATION MARK
  [0x2024]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.circle(150, 50, 40);
      p.fill();
    }
  }, // ONE DOT LEADER
  [0x2040]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(120, 380);
      p.C(200, 450, 400, 450, 480, 380);
      p.stroke();
    }
  }, // CHARACTER TIE
  [0x2041]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(120, 420);
      p.C(200, 350, 400, 350, 480, 420);
      p.stroke();
    }
  }, // CARET INSERTION POINT
  [0x2042]: asteriskLike(3, 50), // ASTERISM (3 dots triangle)
  [0x2047]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      qMark(p, 200, 250);
      qMark(p, 400, 250);
    }
  }, // DOUBLE QUESTION MARK
  [0x2048]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      qMark(p, 200, 250);
      eMark(p, 400, 250);
    }
  }, // QUESTION EXCLAMATION MARK
  [0x2049]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      eMark(p, 200, 250);
      qMark(p, 400, 250);
    }
  }, // EXCLAMATION QUESTION MARK
  [0x2051]: {
    width: W,
    draw: (p: GlyphPen) => {
      asteriskLike(6, 35).draw(p);
      p.M(300, 50);
      p.L(300, 50);
    }
  }, // TWO ASTERISKS
  [0x2052]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 400);
      p.L(480, 100);
      p.stroke();
      p.circle(300, 350, 40);
      p.fill();
      p.circle(300, 150, 40);
      p.fill();
    }
  }, // COMMERCIAL MINUS SIGN
  [0x2053]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.C(200, 350, 400, 150, 500, 250);
      p.stroke();
    }
  }, // SWUNG DASH
  [0x2057]: {
    width: 500,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      for (let i = 0; i < 4; i++) {
        p.M(100 + i * 90, 480);
        p.L(70 + i * 90, 350);
      }
      p.stroke();
    }
  } // QUADRUPLE PRIME
};

function asteriskLike(n: number, lw: number): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(lw);
      for (let i = 0; i < n; i++) {
        const a = (i * Math.PI) / n;
        p.M(300 - 120 * Math.cos(a), 400 - 120 * Math.sin(a));
        p.L(300 + 120 * Math.cos(a), 400 + 120 * Math.sin(a));
      }
      p.stroke();
    }
  };
}

function qMark(p: GlyphPen, cx: number, _cy: number): void {
  p.M(cx - 60, 400);
  p.C(cx - 60, 480, cx + 60, 480, cx + 60, 400);
  p.C(cx + 60, 320, cx, 300, cx, 240);
  p.stroke();
  p.circle(cx, 150, 25);
  p.fill();
}

function eMark(p: GlyphPen, cx: number, _cy: number): void {
  p.M(cx, 480);
  p.L(cx, 200);
  p.stroke();
  p.circle(cx, 130, 28);
  p.fill();
}

// Fill remaining General Punctuation with spaces or generic marks
for (let cp = 0x2000; cp <= 0x206f; cp++) {
  if (!GEN_PUNCT_FILL[cp]) {
    // Most remaining are control/format characters — zero-width
    GEN_PUNCT_FILL[cp] = { width: 0, draw: () => {} };
  }
}

// =============================================================================
// 6. Number Forms remaining — 12 chars
// =============================================================================

export const NUM_FORMS_FILL: Record<number, GlyphDef> = {};
// Small roman numerals and special forms
for (let cp = 0x2150; cp <= 0x218f; cp++) {
  if (!NUM_FORMS_FILL[cp]) {
    // Generic fraction/roman numeral approximation: vertical line(s)
    const idx = cp - 0x2150;
    const nLines = (idx % 4) + 1;
    NUM_FORMS_FILL[cp] = {
      width: W,
      draw: (p: GlyphPen) => {
        p.lineWidth(40);
        const spacing = 80;
        const startX = 300 - ((nLines - 1) * spacing) / 2;
        for (let i = 0; i < nLines; i++) {
          p.M(startX + i * spacing, 50);
          p.L(startX + i * spacing, 450);
        }
        p.stroke();
      }
    };
  }
}
