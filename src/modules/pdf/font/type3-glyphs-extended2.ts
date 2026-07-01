/**
 * Extended Type3 glyph definitions — Part 2
 *
 * Additional hand-crafted glyph blocks:
 *  - Arrows Extended: Harpoons & Paired Arrows (U+21BC–U+21FF)
 *  - Dingbats Full (U+2700–U+27BF)
 *  - Misc Symbols Full (U+2600–U+26FF)
 */

import type { GlyphDef, GlyphPen } from "@pdf/font/type3-glyphs";

const W = 600;

// =============================================================================
// Helper: N-pointed star polygon
// =============================================================================

function starGlyph(
  n: number,
  cx: number,
  cy: number,
  outerR: number,
  innerR: number,
  filled: boolean,
  lw = 35
): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      if (!filled) {
        p.lineWidth(lw);
      }
      for (let i = 0; i < n; i++) {
        const a1 = ((i * 360) / n - 90) * (Math.PI / 180);
        const a2 = ((i * 360) / n + 180 / n - 90) * (Math.PI / 180);
        if (i === 0) {
          p.M(cx + outerR * Math.cos(a1), cy + outerR * Math.sin(a1));
        } else {
          p.L(cx + outerR * Math.cos(a1), cy + outerR * Math.sin(a1));
        }
        p.L(cx + innerR * Math.cos(a2), cy + innerR * Math.sin(a2));
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
// Arrows Extended — HARPOONS & PAIRED ARROWS (U+21BC–U+21FF)
// =============================================================================

export const ARROWS_HARPOONS: Record<number, GlyphDef> = {
  // ↼ LEFTWARDS HARPOON WITH BARB UPWARDS
  [0x21bc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(250, 380);
      p.L(250, 250);
      p.Z();
      p.fill();
    }
  },
  // ↽ LEFTWARDS HARPOON WITH BARB DOWNWARDS
  [0x21bd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(250, 120);
      p.L(250, 250);
      p.Z();
      p.fill();
    }
  },
  // ↾ UPWARDS HARPOON WITH BARB RIGHTWARDS
  [0x21be]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(430, 300);
      p.L(300, 300);
      p.Z();
      p.fill();
    }
  },
  // ↿ UPWARDS HARPOON WITH BARB LEFTWARDS
  [0x21bf]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(170, 300);
      p.L(300, 300);
      p.Z();
      p.fill();
    }
  },
  // ⇀ RIGHTWARDS HARPOON WITH BARB UPWARDS
  [0x21c0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(350, 380);
      p.L(350, 250);
      p.Z();
      p.fill();
    }
  },
  // ⇁ RIGHTWARDS HARPOON WITH BARB DOWNWARDS
  [0x21c1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(350, 120);
      p.L(350, 250);
      p.Z();
      p.fill();
    }
  },
  // ⇂ DOWNWARDS HARPOON WITH BARB RIGHTWARDS
  [0x21c2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
      p.M(300, 50);
      p.L(430, 200);
      p.L(300, 200);
      p.Z();
      p.fill();
    }
  },
  // ⇃ DOWNWARDS HARPOON WITH BARB LEFTWARDS
  [0x21c3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
      p.M(300, 50);
      p.L(170, 200);
      p.L(300, 200);
      p.Z();
      p.fill();
    }
  },
  // ⇄ RIGHTWARDS ARROW OVER LEFTWARDS ARROW
  [0x21c4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // top: rightwards
      p.M(100, 330);
      p.L(500, 330);
      p.stroke();
      p.M(500, 330);
      p.L(380, 400);
      p.L(380, 260);
      p.Z();
      p.fill();
      // bottom: leftwards
      p.M(500, 170);
      p.L(100, 170);
      p.stroke();
      p.M(100, 170);
      p.L(220, 240);
      p.L(220, 100);
      p.Z();
      p.fill();
    }
  },
  // ⇅ UPWARDS ARROW LEFTWARDS OF DOWNWARDS ARROW
  [0x21c5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // left: upwards
      p.M(200, 80);
      p.L(200, 420);
      p.stroke();
      p.M(200, 420);
      p.L(130, 310);
      p.L(270, 310);
      p.Z();
      p.fill();
      // right: downwards
      p.M(400, 420);
      p.L(400, 80);
      p.stroke();
      p.M(400, 80);
      p.L(330, 190);
      p.L(470, 190);
      p.Z();
      p.fill();
    }
  },
  // ⇆ LEFTWARDS ARROW OVER RIGHTWARDS ARROW
  [0x21c6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // top: leftwards
      p.M(500, 330);
      p.L(100, 330);
      p.stroke();
      p.M(100, 330);
      p.L(220, 400);
      p.L(220, 260);
      p.Z();
      p.fill();
      // bottom: rightwards
      p.M(100, 170);
      p.L(500, 170);
      p.stroke();
      p.M(500, 170);
      p.L(380, 240);
      p.L(380, 100);
      p.Z();
      p.fill();
    }
  },
  // ⇇ LEFTWARDS PAIRED ARROWS
  [0x21c7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 330);
      p.L(130, 330);
      p.stroke();
      p.M(130, 330);
      p.L(250, 400);
      p.L(250, 260);
      p.Z();
      p.fill();
      p.M(500, 170);
      p.L(130, 170);
      p.stroke();
      p.M(130, 170);
      p.L(250, 240);
      p.L(250, 100);
      p.Z();
      p.fill();
    }
  },
  // ⇈ UPWARDS PAIRED ARROWS
  [0x21c8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(200, 80);
      p.L(200, 420);
      p.stroke();
      p.M(200, 420);
      p.L(130, 310);
      p.L(270, 310);
      p.Z();
      p.fill();
      p.M(400, 80);
      p.L(400, 420);
      p.stroke();
      p.M(400, 420);
      p.L(330, 310);
      p.L(470, 310);
      p.Z();
      p.fill();
    }
  },
  // ⇉ RIGHTWARDS PAIRED ARROWS
  [0x21c9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 330);
      p.L(470, 330);
      p.stroke();
      p.M(470, 330);
      p.L(350, 400);
      p.L(350, 260);
      p.Z();
      p.fill();
      p.M(100, 170);
      p.L(470, 170);
      p.stroke();
      p.M(470, 170);
      p.L(350, 240);
      p.L(350, 100);
      p.Z();
      p.fill();
    }
  },
  // ⇊ DOWNWARDS PAIRED ARROWS
  [0x21ca]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(200, 420);
      p.L(200, 80);
      p.stroke();
      p.M(200, 80);
      p.L(130, 190);
      p.L(270, 190);
      p.Z();
      p.fill();
      p.M(400, 420);
      p.L(400, 80);
      p.stroke();
      p.M(400, 80);
      p.L(330, 190);
      p.L(470, 190);
      p.Z();
      p.fill();
    }
  },
  // ⇋ LEFTWARDS HARPOON OVER RIGHTWARDS HARPOON
  [0x21cb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // top: leftwards harpoon (barb up)
      p.M(500, 330);
      p.L(100, 330);
      p.stroke();
      p.M(100, 330);
      p.L(220, 410);
      p.L(220, 330);
      p.Z();
      p.fill();
      // bottom: rightwards harpoon (barb down)
      p.M(100, 170);
      p.L(500, 170);
      p.stroke();
      p.M(500, 170);
      p.L(380, 90);
      p.L(380, 170);
      p.Z();
      p.fill();
    }
  },
  // ⇌ RIGHTWARDS HARPOON OVER LEFTWARDS HARPOON
  [0x21cc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // top: rightwards harpoon (barb up)
      p.M(100, 330);
      p.L(500, 330);
      p.stroke();
      p.M(500, 330);
      p.L(380, 410);
      p.L(380, 330);
      p.Z();
      p.fill();
      // bottom: leftwards harpoon (barb down)
      p.M(500, 170);
      p.L(100, 170);
      p.stroke();
      p.M(100, 170);
      p.L(220, 90);
      p.L(220, 170);
      p.Z();
      p.fill();
    }
  },
  // ⇍ LEFTWARDS DOUBLE ARROW WITH STROKE
  [0x21cd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(100, 250);
      p.L(280, 420);
      p.L(280, 300);
      p.L(500, 300);
      p.L(500, 200);
      p.L(280, 200);
      p.L(280, 80);
      p.Z();
      p.fill();
      p.lineWidth(50);
      p.M(350, 120);
      p.L(250, 380);
      p.stroke();
    }
  },
  // ⇎ LEFT RIGHT DOUBLE ARROW WITH STROKE
  [0x21ce]: {
    width: W,
    draw: (p: GlyphPen) => {
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
      p.lineWidth(50);
      p.M(350, 120);
      p.L(250, 380);
      p.stroke();
    }
  },
  // ⇏ RIGHTWARDS DOUBLE ARROW WITH STROKE
  [0x21cf]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 250);
      p.L(320, 420);
      p.L(320, 300);
      p.L(100, 300);
      p.L(100, 200);
      p.L(320, 200);
      p.L(320, 80);
      p.Z();
      p.fill();
      p.lineWidth(50);
      p.M(250, 120);
      p.L(350, 380);
      p.stroke();
    }
  },
  // ⇕ UP DOWN DOUBLE ARROW
  [0x21d5]: {
    width: W,
    draw: (p: GlyphPen) => {
      // up arrowhead
      p.M(300, 480);
      p.L(460, 340);
      p.L(360, 340);
      p.L(360, 160);
      p.L(460, 160);
      p.L(300, 20);
      p.L(140, 160);
      p.L(240, 160);
      p.L(240, 340);
      p.L(140, 340);
      p.Z();
      p.fill();
    }
  },
  // ⇖ NORTH WEST DOUBLE ARROW
  [0x21d6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(100, 460);
      p.L(100, 260);
      p.L(200, 310);
      p.L(380, 130);
      p.L(430, 180);
      p.L(250, 360);
      p.L(300, 460);
      p.Z();
      p.fill();
    }
  },
  // ⇗ NORTH EAST DOUBLE ARROW
  [0x21d7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 460);
      p.L(300, 460);
      p.L(350, 360);
      p.L(170, 180);
      p.L(220, 130);
      p.L(400, 310);
      p.L(500, 260);
      p.Z();
      p.fill();
    }
  },
  // ⇘ SOUTH EAST DOUBLE ARROW
  [0x21d8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 40);
      p.L(500, 240);
      p.L(400, 190);
      p.L(220, 370);
      p.L(170, 320);
      p.L(350, 140);
      p.L(300, 40);
      p.Z();
      p.fill();
    }
  },
  // ⇙ SOUTH WEST DOUBLE ARROW
  [0x21d9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(100, 40);
      p.L(300, 40);
      p.L(250, 140);
      p.L(430, 320);
      p.L(380, 370);
      p.L(200, 190);
      p.L(100, 240);
      p.Z();
      p.fill();
    }
  },
  // ⇚ LEFTWARDS TRIPLE ARROW
  [0x21da]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(500, 310);
      p.L(200, 310);
      p.stroke();
      p.M(500, 250);
      p.L(200, 250);
      p.stroke();
      p.M(500, 190);
      p.L(200, 190);
      p.stroke();
      p.M(100, 250);
      p.L(250, 400);
      p.L(250, 100);
      p.Z();
      p.fill();
    }
  },
  // ⇛ RIGHTWARDS TRIPLE ARROW
  [0x21db]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(100, 310);
      p.L(400, 310);
      p.stroke();
      p.M(100, 250);
      p.L(400, 250);
      p.stroke();
      p.M(100, 190);
      p.L(400, 190);
      p.stroke();
      p.M(500, 250);
      p.L(350, 400);
      p.L(350, 100);
      p.Z();
      p.fill();
    }
  },
  // ⇜ LEFTWARDS SQUIGGLE ARROW
  [0x21dc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.C(440, 330, 380, 170, 300, 250);
      p.C(220, 330, 180, 170, 130, 250);
      p.stroke();
      p.M(130, 250);
      p.L(230, 350);
      p.L(230, 150);
      p.Z();
      p.fill();
    }
  },
  // ⇝ RIGHTWARDS SQUIGGLE ARROW
  [0x21dd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.C(160, 330, 220, 170, 300, 250);
      p.C(380, 330, 420, 170, 470, 250);
      p.stroke();
      p.M(470, 250);
      p.L(370, 350);
      p.L(370, 150);
      p.Z();
      p.fill();
    }
  },
  // ⇞ UPWARDS ARROW WITH DOUBLE STROKE
  [0x21de]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(180, 320);
      p.L(420, 320);
      p.Z();
      p.fill();
      p.lineWidth(40);
      p.M(200, 200);
      p.L(400, 200);
      p.stroke();
      p.M(200, 130);
      p.L(400, 130);
      p.stroke();
    }
  },
  // ⇟ DOWNWARDS ARROW WITH DOUBLE STROKE
  [0x21df]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
      p.M(300, 50);
      p.L(180, 180);
      p.L(420, 180);
      p.Z();
      p.fill();
      p.lineWidth(40);
      p.M(200, 300);
      p.L(400, 300);
      p.stroke();
      p.M(200, 370);
      p.L(400, 370);
      p.stroke();
    }
  },
  // ⇠ LEFTWARDS DASHED ARROW
  [0x21e0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(500, 250);
      p.L(430, 250);
      p.stroke();
      p.M(370, 250);
      p.L(300, 250);
      p.stroke();
      p.M(240, 250);
      p.L(170, 250);
      p.stroke();
      p.M(100, 250);
      p.L(230, 370);
      p.L(230, 130);
      p.Z();
      p.fill();
    }
  },
  // ⇡ UPWARDS DASHED ARROW
  [0x21e1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 50);
      p.L(300, 120);
      p.stroke();
      p.M(300, 180);
      p.L(300, 250);
      p.stroke();
      p.M(300, 310);
      p.L(300, 380);
      p.stroke();
      p.M(300, 450);
      p.L(180, 320);
      p.L(420, 320);
      p.Z();
      p.fill();
    }
  },
  // ⇢ RIGHTWARDS DASHED ARROW
  [0x21e2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(170, 250);
      p.stroke();
      p.M(230, 250);
      p.L(300, 250);
      p.stroke();
      p.M(360, 250);
      p.L(430, 250);
      p.stroke();
      p.M(500, 250);
      p.L(370, 370);
      p.L(370, 130);
      p.Z();
      p.fill();
    }
  },
  // ⇣ DOWNWARDS DASHED ARROW
  [0x21e3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 450);
      p.L(300, 380);
      p.stroke();
      p.M(300, 320);
      p.L(300, 250);
      p.stroke();
      p.M(300, 190);
      p.L(300, 120);
      p.stroke();
      p.M(300, 50);
      p.L(180, 180);
      p.L(420, 180);
      p.Z();
      p.fill();
    }
  },
  // ⇤ LEFTWARDS ARROW TO BAR
  [0x21e4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(500, 250);
      p.L(150, 250);
      p.stroke();
      p.M(150, 250);
      p.L(280, 370);
      p.L(280, 130);
      p.Z();
      p.fill();
      p.M(100, 100);
      p.L(100, 400);
      p.stroke();
    }
  },
  // ⇥ RIGHTWARDS ARROW TO BAR
  [0x21e5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(450, 250);
      p.stroke();
      p.M(450, 250);
      p.L(320, 370);
      p.L(320, 130);
      p.Z();
      p.fill();
      p.M(500, 100);
      p.L(500, 400);
      p.stroke();
    }
  },
  // ⇦ LEFTWARDS WHITE ARROW
  [0x21e6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 250);
      p.L(250, 430);
      p.L(250, 320);
      p.L(520, 320);
      p.L(520, 180);
      p.L(250, 180);
      p.L(250, 70);
      p.Z();
      p.stroke();
    }
  },
  // ⇧ UPWARDS WHITE ARROW
  [0x21e7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(500, 270);
      p.L(390, 270);
      p.L(390, 30);
      p.L(210, 30);
      p.L(210, 270);
      p.L(100, 270);
      p.Z();
      p.stroke();
    }
  },
  // ⇨ RIGHTWARDS WHITE ARROW
  [0x21e8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(520, 250);
      p.L(350, 430);
      p.L(350, 320);
      p.L(80, 320);
      p.L(80, 180);
      p.L(350, 180);
      p.L(350, 70);
      p.Z();
      p.stroke();
    }
  },
  // ⇩ DOWNWARDS WHITE ARROW
  [0x21e9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 30);
      p.L(100, 230);
      p.L(210, 230);
      p.L(210, 470);
      p.L(390, 470);
      p.L(390, 230);
      p.L(500, 230);
      p.Z();
      p.stroke();
    }
  },
  // ⇪ UPWARDS WHITE ARROW FROM BAR
  [0x21ea]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(490, 280);
      p.L(380, 280);
      p.L(380, 100);
      p.L(220, 100);
      p.L(220, 280);
      p.L(110, 280);
      p.Z();
      p.stroke();
      p.M(130, 40);
      p.L(470, 40);
      p.stroke();
    }
  },
  // ⇵ DOWNWARDS ARROW LEFTWARDS OF UPWARDS ARROW
  [0x21f5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // left: downwards
      p.M(200, 420);
      p.L(200, 80);
      p.stroke();
      p.M(200, 80);
      p.L(130, 190);
      p.L(270, 190);
      p.Z();
      p.fill();
      // right: upwards
      p.M(400, 80);
      p.L(400, 420);
      p.stroke();
      p.M(400, 420);
      p.L(330, 310);
      p.L(470, 310);
      p.Z();
      p.fill();
    }
  }
};

// =============================================================================
// Dingbats Full (U+2700–U+27BF)
// =============================================================================

export const DINGBATS_FULL: Record<number, GlyphDef> = {
  // ✨ SPARKLES — three small stars
  [0x2728]: {
    width: W,
    draw: (p: GlyphPen) => {
      // large sparkle center
      p.M(300, 420);
      p.L(270, 300);
      p.L(180, 280);
      p.L(270, 260);
      p.L(300, 140);
      p.L(330, 260);
      p.L(420, 280);
      p.L(330, 300);
      p.Z();
      p.fill();
      // small top-right
      p.M(440, 460);
      p.L(430, 420);
      p.L(400, 410);
      p.L(430, 400);
      p.L(440, 360);
      p.L(450, 400);
      p.L(480, 410);
      p.L(450, 420);
      p.Z();
      p.fill();
      // small bottom-left
      p.M(160, 170);
      p.L(150, 130);
      p.L(120, 120);
      p.L(150, 110);
      p.L(160, 70);
      p.L(170, 110);
      p.L(200, 120);
      p.L(170, 130);
      p.Z();
      p.fill();
    }
  },
  // ✩ STRESS OUTLINED WHITE STAR
  [0x2729]: starGlyph(5, 300, 260, 220, 90, false, 40),
  // ✪ CIRCLED WHITE STAR
  [0x272a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 260, 230);
      p.stroke();
      const cx = 300;
      const cy = 260;
      const R = 180;
      const r = 70;
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
  // ✫ OPEN CENTRE BLACK STAR
  [0x272b]: starGlyph(5, 300, 260, 230, 110, true),
  // ✬ BLACK CENTRE WHITE STAR
  [0x272c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      const cx = 300;
      const cy = 260;
      const R = 220;
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
      p.circle(cx, cy, 60);
      p.fill();
    }
  },
  // ✭ OUTLINED BLACK STAR
  [0x272d]: starGlyph(5, 300, 260, 230, 100, true),
  // ✮ HEAVY OUTLINED BLACK STAR
  [0x272e]: starGlyph(5, 300, 260, 240, 105, true),
  // ✯ PINWHEEL STAR
  [0x272f]: starGlyph(5, 300, 260, 220, 80, false, 30),
  // ✱ HEAVY ASTERISK
  [0x2731]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(60);
      for (let a = 0; a < 360; a += 60) {
        const rad = (a * Math.PI) / 180;
        p.M(300, 250);
        p.L(300 + 200 * Math.cos(rad), 250 + 200 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // ✲ OPEN CENTRE ASTERISK
  [0x2732]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      for (let a = 0; a < 360; a += 60) {
        const rad = (a * Math.PI) / 180;
        p.M(300 + 60 * Math.cos(rad), 250 + 60 * Math.sin(rad));
        p.L(300 + 200 * Math.cos(rad), 250 + 200 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // ✳ EIGHT SPOKED ASTERISK
  [0x2733]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      for (let a = 0; a < 360; a += 45) {
        const rad = (a * Math.PI) / 180;
        p.M(300, 250);
        p.L(300 + 190 * Math.cos(rad), 250 + 190 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // ✴ EIGHT POINTED BLACK STAR
  [0x2734]: starGlyph(8, 300, 260, 220, 100, true),
  // ✵ EIGHT POINTED PINWHEEL STAR
  [0x2735]: starGlyph(8, 300, 260, 220, 80, false, 30),
  // ✶ SIX POINTED BLACK STAR
  [0x2736]: starGlyph(6, 300, 260, 220, 110, true),
  // ✷ EIGHT POINTED RECTILINEAR BLACK STAR
  [0x2737]: starGlyph(8, 300, 260, 220, 90, true),
  // ✸ HEAVY EIGHT POINTED RECTILINEAR BLACK STAR
  [0x2738]: starGlyph(8, 300, 260, 230, 80, true),
  // ✹ TWELVE POINTED BLACK STAR
  [0x2739]: starGlyph(12, 300, 260, 220, 130, true),
  // ✺ SIXTEEN POINTED ASTERISK
  [0x273a]: starGlyph(16, 300, 260, 220, 140, true),
  // ✻ TEARDROP-SPOKED ASTERISK
  [0x273b]: starGlyph(6, 300, 260, 220, 80, true),
  // ✼ OPEN CENTRE TEARDROP-SPOKED ASTERISK
  [0x273c]: starGlyph(6, 300, 260, 220, 80, false, 35),
  // ✽ HEAVY TEARDROP-SPOKED ASTERISK
  [0x273d]: starGlyph(6, 300, 260, 230, 70, true),
  // ✾ SIX PETALLED BLACK AND WHITE FLORETTE
  [0x273e]: starGlyph(6, 300, 260, 210, 120, false, 40),
  // ✿ BLACK FLORETTE
  [0x273f]: starGlyph(6, 300, 260, 210, 120, true),
  // ❀ WHITE FLORETTE
  [0x2740]: starGlyph(6, 300, 260, 210, 120, false, 35),
  // ❁ EIGHT PETALLED OUTLINED BLACK FLORETTE
  [0x2741]: starGlyph(8, 300, 260, 210, 130, false, 40),
  // ❂ CIRCLED OPEN CENTRE EIGHT POINTED STAR
  [0x2742]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.circle(300, 260, 230);
      p.stroke();
      const cx = 300;
      const cy = 260;
      const R = 190;
      const r = 100;
      for (let i = 0; i < 8; i++) {
        const a1 = ((i * 45 - 90) * Math.PI) / 180;
        const a2 = ((i * 45 + 22.5 - 90) * Math.PI) / 180;
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
  // ❃ HEAVY TEARDROP-SPOKED PINWHEEL ASTERISK
  [0x2743]: starGlyph(8, 300, 260, 220, 70, true),
  // ❄ SNOWFLAKE
  [0x2744]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      for (let a = 0; a < 360; a += 60) {
        const rad = (a * Math.PI) / 180;
        p.M(300, 250);
        p.L(300 + 200 * Math.cos(rad), 250 + 200 * Math.sin(rad));
      }
      p.stroke();
      // cross-bars on each arm
      p.lineWidth(25);
      for (let a = 0; a < 360; a += 60) {
        const rad = (a * Math.PI) / 180;
        const mx = 300 + 130 * Math.cos(rad);
        const my = 250 + 130 * Math.sin(rad);
        const perp = rad + Math.PI / 2;
        p.M(mx + 40 * Math.cos(perp), my + 40 * Math.sin(perp));
        p.L(mx - 40 * Math.cos(perp), my - 40 * Math.sin(perp));
      }
      p.stroke();
    }
  },
  // ❅ TIGHT TRIFOLIATE SNOWFLAKE
  [0x2745]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      for (let a = 0; a < 360; a += 60) {
        const rad = (a * Math.PI) / 180;
        p.M(300, 250);
        p.L(300 + 190 * Math.cos(rad), 250 + 190 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // ❆ HEAVY CHEVRON SNOWFLAKE
  [0x2746]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      for (let a = 0; a < 360; a += 60) {
        const rad = (a * Math.PI) / 180;
        p.M(300, 250);
        p.L(300 + 200 * Math.cos(rad), 250 + 200 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // ❇ SPARKLE
  [0x2747]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 480);
      p.L(260, 290);
      p.L(100, 250);
      p.L(260, 210);
      p.L(300, 20);
      p.L(340, 210);
      p.L(500, 250);
      p.L(340, 290);
      p.Z();
      p.fill();
    }
  },
  // ❈ HEAVY SPARKLE
  [0x2748]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 490);
      p.L(250, 300);
      p.L(80, 250);
      p.L(250, 200);
      p.L(300, 10);
      p.L(350, 200);
      p.L(520, 250);
      p.L(350, 300);
      p.Z();
      p.fill();
    }
  },
  // ❉ BALLOON-SPOKED ASTERISK
  [0x2749]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(55);
      for (let a = 0; a < 360; a += 45) {
        const rad = (a * Math.PI) / 180;
        p.M(300, 250);
        p.L(300 + 180 * Math.cos(rad), 250 + 180 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // ❊ TEARDROP-SPOKED PINWHEEL ASTERISK
  [0x274a]: starGlyph(8, 300, 260, 220, 60, true),
  // ❋ HEAVY EIGHT TEARDROP-SPOKED PROPELLER ASTERISK
  [0x274b]: starGlyph(8, 300, 260, 230, 50, true),
  // ❌ CROSS MARK
  [0x274c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(80);
      p.M(120, 70);
      p.L(480, 430);
      p.stroke();
      p.M(120, 430);
      p.L(480, 70);
      p.stroke();
    }
  },
  // ❍ SHADOWED WHITE CIRCLE
  [0x274d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.circle(300, 260, 190);
      p.stroke();
      // shadow offset
      p.lineWidth(25);
      p.M(420, 100);
      p.C(500, 140, 530, 240, 490, 340);
      p.stroke();
    }
  },
  // ❎ NEGATIVE SQUARED CROSS MARK
  [0x274e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(80, 30, 440, 440);
      p.fill();
      // We draw an X in "negative" — approximate by drawing filled square
      // then white X lines on top (not possible with simple fill, so just fill square)
    }
  },
  // ❓ BLACK QUESTION MARK ORNAMENT
  [0x2753]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(65);
      p.M(200, 400);
      p.C(200, 500, 400, 500, 400, 400);
      p.C(400, 320, 300, 300, 300, 220);
      p.stroke();
      p.circle(300, 100, 40);
      p.fill();
    }
  },
  // ❔ WHITE QUESTION MARK ORNAMENT
  [0x2754]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(200, 400);
      p.C(200, 500, 400, 500, 400, 400);
      p.C(400, 320, 300, 300, 300, 220);
      p.stroke();
      p.lineWidth(40);
      p.circle(300, 100, 35);
      p.stroke();
    }
  },
  // ❕ WHITE EXCLAMATION MARK ORNAMENT
  [0x2755]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 480);
      p.L(300, 180);
      p.stroke();
      p.lineWidth(40);
      p.circle(300, 80, 35);
      p.stroke();
    }
  },
  // ❗ HEAVY EXCLAMATION MARK SYMBOL
  [0x2757]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(75);
      p.M(300, 480);
      p.L(300, 180);
      p.stroke();
      p.circle(300, 70, 45);
      p.fill();
    }
  },
  // ❢ EXCLAMATION MARK ORNAMENT
  [0x2762]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(55);
      p.M(300, 480);
      p.L(300, 200);
      p.stroke();
      p.circle(300, 100, 35);
      p.fill();
      // ornament dot above
      p.circle(300, 500, 20);
      p.fill();
    }
  },
  // ❣ HEAVY HEART EXCLAMATION MARK ORNAMENT
  [0x2763]: {
    width: W,
    draw: (p: GlyphPen) => {
      // heart at top
      p.M(300, 320);
      p.C(300, 320, 180, 400, 180, 450);
      p.C(180, 490, 240, 510, 300, 510);
      p.C(360, 510, 420, 490, 420, 450);
      p.C(420, 400, 300, 320, 300, 320);
      p.Z();
      p.fill();
      // stem
      p.lineWidth(55);
      p.M(300, 300);
      p.L(300, 130);
      p.stroke();
      p.circle(300, 50, 35);
      p.fill();
    }
  },
  // ❤ HEAVY BLACK HEART — already exists but repeated for completeness
  // ❥ ROTATED HEAVY BLACK HEART BULLET
  [0x2765]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(250, 80);
      p.C(250, 80, 80, 200, 80, 310);
      p.C(80, 400, 160, 440, 250, 440);
      p.C(340, 440, 400, 400, 400, 310);
      p.C(400, 200, 250, 80, 250, 80);
      p.Z();
      p.fill();
    }
  },
  // ❦ FLORAL HEART
  [0x2766]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 80);
      p.C(300, 80, 100, 220, 100, 340);
      p.C(100, 440, 190, 490, 300, 490);
      p.C(410, 490, 500, 440, 500, 340);
      p.C(500, 220, 300, 80, 300, 80);
      p.Z();
      p.stroke();
      // floral detail: inner curve
      p.M(300, 150);
      p.C(250, 200, 200, 300, 300, 400);
      p.C(400, 300, 350, 200, 300, 150);
      p.stroke();
    }
  },
  // ❧ ROTATED FLORAL HEART BULLET
  [0x2767]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(250, 80);
      p.C(250, 80, 80, 200, 80, 310);
      p.C(80, 400, 160, 440, 250, 440);
      p.C(340, 440, 400, 400, 400, 310);
      p.C(400, 200, 250, 80, 250, 80);
      p.Z();
      p.stroke();
    }
  },
  // ➕ HEAVY PLUS SIGN
  [0x2795]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(220, 50, 160, 400);
      p.fill();
      p.rect(100, 170, 400, 160);
      p.fill();
    }
  },
  // ➖ HEAVY MINUS SIGN
  [0x2796]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(100, 180, 400, 140);
      p.fill();
    }
  },
  // ➗ HEAVY DIVISION SIGN
  [0x2797]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(100, 200, 400, 100);
      p.fill();
      p.circle(300, 400, 50);
      p.fill();
      p.circle(300, 100, 50);
      p.fill();
    }
  },
  // ➘ HEAVY SOUTH EAST ARROW
  [0x2798]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(480, 50);
      p.L(480, 250);
      p.L(320, 200);
      p.L(140, 430);
      p.L(100, 400);
      p.L(280, 170);
      p.L(230, 50);
      p.Z();
      p.fill();
    }
  },
  // ➙ HEAVY RIGHTWARDS ARROW
  [0x2799]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(340, 430);
      p.L(340, 310);
      p.L(80, 310);
      p.L(80, 190);
      p.L(340, 190);
      p.L(340, 70);
      p.Z();
      p.fill();
    }
  },
  // ➚ HEAVY NORTH EAST ARROW
  [0x279a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(480, 450);
      p.L(230, 450);
      p.L(280, 330);
      p.L(100, 100);
      p.L(140, 70);
      p.L(320, 300);
      p.L(480, 250);
      p.Z();
      p.fill();
    }
  },
  // ➛ DRAFTING POINT RIGHTWARDS ARROW
  [0x279b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(350, 400);
      p.L(350, 280);
      p.L(80, 250);
      p.L(350, 220);
      p.L(350, 100);
      p.Z();
      p.fill();
    }
  },
  // ➜ HEAVY ROUND-TIPPED RIGHTWARDS ARROW (exists in DING_EXT, add alias)
  // ➝ TRIANGLE-HEADED RIGHTWARDS ARROW
  [0x279d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 250);
      p.L(400, 250);
      p.stroke();
      p.M(520, 250);
      p.L(370, 380);
      p.L(370, 120);
      p.Z();
      p.fill();
    }
  },
  // ➞ HEAVY TRIANGLE-HEADED RIGHTWARDS ARROW
  [0x279e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(80, 250);
      p.L(380, 250);
      p.stroke();
      p.M(530, 250);
      p.L(350, 420);
      p.L(350, 80);
      p.Z();
      p.fill();
    }
  },
  // ➟ DASHED TRIANGLE-HEADED RIGHTWARDS ARROW
  [0x279f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 250);
      p.L(160, 250);
      p.stroke();
      p.M(210, 250);
      p.L(290, 250);
      p.stroke();
      p.M(340, 250);
      p.L(400, 250);
      p.stroke();
      p.M(520, 250);
      p.L(370, 380);
      p.L(370, 120);
      p.Z();
      p.fill();
    }
  },
  // ➠ HEAVY DASHED TRIANGLE-HEADED RIGHTWARDS ARROW
  [0x27a0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(55);
      p.M(80, 250);
      p.L(170, 250);
      p.stroke();
      p.M(230, 250);
      p.L(350, 250);
      p.stroke();
      p.M(530, 250);
      p.L(350, 420);
      p.L(350, 80);
      p.Z();
      p.fill();
    }
  },
  // ➡ BLACK RIGHTWARDS ARROW (exists in DING_EXT)
  // ➢ THREE-D TOP-LIGHTED RIGHTWARDS ARROWHEAD
  [0x27a2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(200, 430);
      p.L(280, 250);
      p.L(200, 70);
      p.Z();
      p.fill();
    }
  },
  // ➣ THREE-D BOTTOM-LIGHTED RIGHTWARDS ARROWHEAD
  [0x27a3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(520, 250);
      p.L(200, 430);
      p.L(280, 250);
      p.L(200, 70);
      p.Z();
      p.stroke();
    }
  },
  // ➤ BLACK RIGHTWARDS ARROWHEAD
  [0x27a4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(100, 450);
      p.L(100, 50);
      p.Z();
      p.fill();
    }
  },
  // ➥ HEAVY BLACK CURVED DOWNWARDS AND RIGHTWARDS ARROW
  [0x27a5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 200);
      p.L(380, 320);
      p.L(380, 240);
      p.C(300, 240, 200, 280, 150, 380);
      p.L(100, 380);
      p.C(140, 220, 260, 150, 380, 150);
      p.L(380, 80);
      p.Z();
      p.fill();
    }
  },
  // ➦ HEAVY BLACK CURVED UPWARDS AND RIGHTWARDS ARROW
  [0x27a6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 300);
      p.L(380, 420);
      p.L(380, 350);
      p.C(260, 350, 140, 280, 100, 120);
      p.L(150, 120);
      p.C(200, 220, 300, 260, 380, 260);
      p.L(380, 180);
      p.Z();
      p.fill();
    }
  },
  // ➧ SQUAT BLACK RIGHTWARDS ARROW
  [0x27a7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 250);
      p.L(350, 380);
      p.L(350, 310);
      p.L(100, 310);
      p.L(100, 190);
      p.L(350, 190);
      p.L(350, 120);
      p.Z();
      p.fill();
    }
  },
  // ➨ HEAVY CONCAVE-POINTED BLACK RIGHTWARDS ARROW
  [0x27a8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(320, 440);
      p.L(320, 310);
      p.L(80, 310);
      p.C(150, 250, 150, 250, 80, 190);
      p.L(320, 190);
      p.L(320, 60);
      p.Z();
      p.fill();
    }
  },
  // ➩ RIGHT-SHADED WHITE RIGHTWARDS ARROW
  [0x27a9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(520, 250);
      p.L(340, 430);
      p.L(340, 310);
      p.L(80, 310);
      p.L(80, 190);
      p.L(340, 190);
      p.L(340, 70);
      p.Z();
      p.stroke();
    }
  },
  // ➪ LEFT-SHADED WHITE RIGHTWARDS ARROW
  [0x27aa]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(520, 250);
      p.L(340, 430);
      p.L(340, 310);
      p.L(80, 310);
      p.L(80, 190);
      p.L(340, 190);
      p.L(340, 70);
      p.Z();
      p.fillStroke();
    }
  },
  // ➫ BACK-TILTED SHADOWED WHITE RIGHTWARDS ARROW
  [0x27ab]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(340, 440);
      p.L(340, 310);
      p.L(100, 310);
      p.L(100, 190);
      p.L(340, 190);
      p.L(340, 60);
      p.Z();
      p.fill();
    }
  },
  // ➬ FRONT-TILTED SHADOWED WHITE RIGHTWARDS ARROW
  [0x27ac]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(520, 250);
      p.L(340, 440);
      p.L(340, 310);
      p.L(100, 310);
      p.L(100, 190);
      p.L(340, 190);
      p.L(340, 60);
      p.Z();
      p.fillStroke();
    }
  },
  // ➭ HEAVY LOWER RIGHT-SHADOWED WHITE RIGHTWARDS ARROW
  [0x27ad]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(510, 250);
      p.L(340, 420);
      p.L(340, 310);
      p.L(90, 310);
      p.L(90, 190);
      p.L(340, 190);
      p.L(340, 80);
      p.Z();
      p.stroke();
    }
  },
  // ➮ HEAVY UPPER RIGHT-SHADOWED WHITE RIGHTWARDS ARROW
  [0x27ae]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(510, 250);
      p.L(340, 420);
      p.L(340, 310);
      p.L(90, 310);
      p.L(90, 190);
      p.L(340, 190);
      p.L(340, 80);
      p.Z();
      p.stroke();
    }
  },
  // ➯ NOTCHED LOWER RIGHT-SHADOWED WHITE RIGHTWARDS ARROW
  [0x27af]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(350, 430);
      p.L(350, 310);
      p.L(100, 310);
      p.L(140, 250);
      p.L(100, 190);
      p.L(350, 190);
      p.L(350, 70);
      p.Z();
      p.fill();
    }
  },
  // ➱ NOTCHED UPPER RIGHT-SHADOWED WHITE RIGHTWARDS ARROW
  [0x27b1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(520, 250);
      p.L(350, 430);
      p.L(350, 310);
      p.L(100, 310);
      p.L(140, 250);
      p.L(100, 190);
      p.L(350, 190);
      p.L(350, 70);
      p.Z();
      p.stroke();
    }
  },
  // ➲ CIRCLED HEAVY WHITE RIGHTWARDS ARROW
  [0x27b2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 230);
      p.fill();
      // We can only approximate — just a filled circle with arrow shape implied
    }
  },
  // ➳ WHITE-FEATHERED RIGHTWARDS ARROW
  [0x27b3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(80, 250);
      p.L(450, 250);
      p.stroke();
      p.M(520, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
      // feather lines
      p.lineWidth(20);
      p.M(120, 250);
      p.L(80, 320);
      p.stroke();
      p.M(120, 250);
      p.L(80, 180);
      p.stroke();
    }
  },
  // ➴ BLACK-FEATHERED SOUTH EAST ARROW
  [0x27b4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(480, 50);
      p.L(480, 220);
      p.L(200, 400);
      p.L(250, 50);
      p.Z();
      p.fill();
    }
  },
  // ➵ BLACK-FEATHERED RIGHTWARDS ARROW
  [0x27b5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(390, 360);
      p.L(390, 275);
      p.L(80, 275);
      p.L(120, 250);
      p.L(80, 225);
      p.L(390, 225);
      p.L(390, 140);
      p.Z();
      p.fill();
    }
  },
  // ➶ BLACK-FEATHERED NORTH EAST ARROW
  [0x27b6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(480, 450);
      p.L(250, 450);
      p.L(200, 100);
      p.L(480, 280);
      p.Z();
      p.fill();
    }
  },
  // ➷ HEAVY BLACK-FEATHERED SOUTH EAST ARROW
  [0x27b7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 50);
      p.L(500, 240);
      p.L(180, 420);
      p.L(230, 50);
      p.Z();
      p.fill();
    }
  },
  // ➸ HEAVY BLACK-FEATHERED RIGHTWARDS ARROW
  [0x27b8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(530, 250);
      p.L(380, 380);
      p.L(380, 290);
      p.L(70, 290);
      p.L(120, 250);
      p.L(70, 210);
      p.L(380, 210);
      p.L(380, 120);
      p.Z();
      p.fill();
    }
  },
  // ➹ HEAVY BLACK-FEATHERED NORTH EAST ARROW
  [0x27b9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 450);
      p.L(230, 450);
      p.L(180, 80);
      p.L(500, 260);
      p.Z();
      p.fill();
    }
  },
  // ➺ TEARDROP-BARBED RIGHTWARDS ARROW
  [0x27ba]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(370, 380);
      p.L(370, 270);
      p.C(200, 270, 120, 260, 80, 250);
      p.C(120, 240, 200, 230, 370, 230);
      p.L(370, 120);
      p.Z();
      p.fill();
    }
  },
  // ➻ HEAVY TEARDROP-SHANKED RIGHTWARDS ARROW
  [0x27bb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(530, 250);
      p.L(380, 400);
      p.L(380, 290);
      p.C(250, 290, 150, 280, 70, 250);
      p.C(150, 220, 250, 210, 380, 210);
      p.L(380, 100);
      p.Z();
      p.fill();
    }
  },
  // ➼ WEDGE-TAILED RIGHTWARDS ARROW
  [0x27bc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(80, 310);
      p.L(400, 260);
      p.stroke();
      p.M(80, 190);
      p.L(400, 240);
      p.stroke();
      p.M(520, 250);
      p.L(380, 370);
      p.L(380, 130);
      p.Z();
      p.fill();
    }
  },
  // ➽ HEAVY WEDGE-TAILED RIGHTWARDS ARROW
  [0x27bd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(530, 250);
      p.L(380, 400);
      p.L(380, 280);
      p.L(70, 320);
      p.L(70, 180);
      p.L(380, 220);
      p.L(380, 100);
      p.Z();
      p.fill();
    }
  },
  // ➾ OPEN-OUTLINED RIGHTWARDS ARROW
  [0x27be]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(510, 250);
      p.L(340, 420);
      p.L(340, 310);
      p.L(90, 310);
      p.L(90, 190);
      p.L(340, 190);
      p.L(340, 80);
      p.Z();
      p.stroke();
    }
  }
};

// =============================================================================
// Misc Symbols Full (U+2600–U+26FF)
// =============================================================================

export const MISC_SYMBOLS_FULL: Record<number, GlyphDef> = {
  // ☂ UMBRELLA
  [0x2602]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      // dome
      p.M(80, 280);
      p.C(80, 480, 300, 520, 300, 520);
      p.C(300, 520, 520, 480, 520, 280);
      p.L(80, 280);
      p.stroke();
      // handle
      p.M(300, 280);
      p.L(300, 80);
      p.C(300, 30, 260, 20, 240, 50);
      p.stroke();
    }
  },
  // ☃ SNOWMAN
  [0x2603]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 100, 80);
      p.stroke();
      p.circle(300, 260, 110);
      p.stroke();
      p.circle(300, 430, 70);
      p.stroke();
    }
  },
  // ☄ COMET
  [0x2604]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(420, 380, 70);
      p.fill();
      // tail
      p.lineWidth(30);
      p.M(370, 420);
      p.C(250, 470, 100, 480, 80, 460);
      p.stroke();
      p.M(380, 440);
      p.C(280, 490, 120, 500, 80, 490);
      p.stroke();
    }
  },
  // ☔ UMBRELLA WITH RAIN DROPS
  [0x2614]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // dome
      p.M(80, 300);
      p.C(80, 480, 300, 500, 300, 500);
      p.C(300, 500, 520, 480, 520, 300);
      p.L(80, 300);
      p.stroke();
      // handle
      p.M(300, 300);
      p.L(300, 100);
      p.C(300, 50, 260, 40, 240, 70);
      p.stroke();
      // rain drops
      p.lineWidth(25);
      p.M(150, 200);
      p.L(130, 160);
      p.stroke();
      p.M(300, 180);
      p.L(280, 140);
      p.stroke();
      p.M(450, 200);
      p.L(430, 160);
      p.stroke();
    }
  },
  // ☕ HOT BEVERAGE
  [0x2615]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      // cup body
      p.M(120, 50);
      p.L(120, 300);
      p.L(400, 300);
      p.L(400, 50);
      p.stroke();
      // handle
      p.M(400, 250);
      p.C(480, 250, 500, 180, 480, 120);
      p.C(460, 80, 400, 80, 400, 100);
      p.stroke();
      // steam
      p.lineWidth(25);
      p.M(200, 340);
      p.C(200, 380, 220, 400, 200, 440);
      p.stroke();
      p.M(300, 340);
      p.C(300, 380, 320, 400, 300, 440);
      p.stroke();
    }
  },
  // ☘ SHAMROCK
  [0x2618]: {
    width: W,
    draw: (p: GlyphPen) => {
      // three leaves using circles
      p.circle(300, 370, 80);
      p.fill();
      p.circle(220, 280, 80);
      p.fill();
      p.circle(380, 280, 80);
      p.fill();
      // stem
      p.lineWidth(30);
      p.M(300, 290);
      p.C(300, 200, 280, 100, 260, 50);
      p.stroke();
    }
  },
  // ☠ SKULL AND CROSSBONES
  [0x2620]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // skull (circle)
      p.circle(300, 350, 120);
      p.stroke();
      // eyes
      p.circle(260, 370, 25);
      p.fill();
      p.circle(340, 370, 25);
      p.fill();
      // jaw
      p.M(250, 260);
      p.L(350, 260);
      p.stroke();
      // crossbones
      p.lineWidth(30);
      p.M(130, 100);
      p.L(470, 200);
      p.stroke();
      p.M(130, 200);
      p.L(470, 100);
      p.stroke();
    }
  },
  // ☢ RADIOACTIVE SIGN
  [0x2622]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 40);
      p.fill();
      // three sectors (trefoil)
      p.lineWidth(35);
      for (let a = 0; a < 3; a++) {
        const angle = ((a * 120 - 90) * Math.PI) / 180;
        const x1 = 300 + 70 * Math.cos(angle - 0.4);
        const y1 = 250 + 70 * Math.sin(angle - 0.4);
        const x2 = 300 + 200 * Math.cos(angle - 0.4);
        const y2 = 250 + 200 * Math.sin(angle - 0.4);
        const x3 = 300 + 200 * Math.cos(angle + 0.4);
        const y3 = 250 + 200 * Math.sin(angle + 0.4);
        const x4 = 300 + 70 * Math.cos(angle + 0.4);
        const y4 = 250 + 70 * Math.sin(angle + 0.4);
        p.M(x1, y1);
        p.L(x2, y2);
        p.L(x3, y3);
        p.L(x4, y4);
        p.Z();
        p.fill();
      }
    }
  },
  // ☣ BIOHAZARD SIGN
  [0x2623]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // three interlocking circles
      const r = 110;
      const d = 90;
      for (let a = 0; a < 3; a++) {
        const angle = ((a * 120 - 90) * Math.PI) / 180;
        p.circle(300 + d * Math.cos(angle), 250 + d * Math.sin(angle), r);
        p.stroke();
      }
      // center dot
      p.circle(300, 250, 25);
      p.fill();
    }
  },
  // ☦ ORTHODOX CROSS
  [0x2626]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      // vertical
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      // upper cross bar
      p.M(180, 400);
      p.L(420, 400);
      p.stroke();
      // lower slanted bar
      p.M(180, 120);
      p.L(420, 80);
      p.stroke();
      // top small bar
      p.M(220, 460);
      p.L(380, 460);
      p.stroke();
    }
  },
  // ☸ WHEEL OF DHARMA
  [0x2638]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 200);
      p.stroke();
      // spokes
      for (let a = 0; a < 360; a += 45) {
        const rad = (a * Math.PI) / 180;
        p.M(300, 250);
        p.L(300 + 200 * Math.cos(rad), 250 + 200 * Math.sin(rad));
      }
      p.stroke();
      p.circle(300, 250, 40);
      p.fill();
    }
  },
  // ☼ WHITE SUN WITH RAYS
  [0x263c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 110);
      p.stroke();
      for (let a = 0; a < 360; a += 30) {
        const rad = (a * Math.PI) / 180;
        p.M(300 + 130 * Math.cos(rad), 250 + 130 * Math.sin(rad));
        p.L(300 + 210 * Math.cos(rad), 250 + 210 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // ♈ ARIES
  [0x2648]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 50);
      p.C(150, 300, 300, 400, 300, 480);
      p.stroke();
      p.M(450, 50);
      p.C(450, 300, 300, 400, 300, 480);
      p.stroke();
    }
  },
  // ♉ TAURUS
  [0x2649]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.circle(300, 200, 150);
      p.stroke();
      p.M(100, 430);
      p.C(200, 500, 400, 500, 500, 430);
      p.stroke();
    }
  },
  // ♊ GEMINI
  [0x264a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 450);
      p.C(200, 500, 400, 500, 480, 450);
      p.stroke();
      p.M(120, 50);
      p.C(200, 0, 400, 0, 480, 50);
      p.stroke();
      p.M(220, 450);
      p.L(220, 50);
      p.stroke();
      p.M(380, 450);
      p.L(380, 50);
      p.stroke();
    }
  },
  // ♋ CANCER
  [0x264b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(480, 380);
      p.C(300, 500, 120, 380, 120, 300);
      p.stroke();
      p.circle(200, 350, 60);
      p.fill();
      p.M(120, 120);
      p.C(300, 0, 480, 120, 480, 200);
      p.stroke();
      p.circle(400, 150, 60);
      p.fill();
    }
  },
  // ♌ LEO
  [0x264c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.circle(200, 150, 100);
      p.stroke();
      p.M(300, 150);
      p.C(400, 150, 480, 250, 480, 350);
      p.C(480, 450, 420, 480, 380, 430);
      p.stroke();
    }
  },
  // ♍ VIRGO
  [0x264d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 50);
      p.L(100, 400);
      p.C(100, 480, 200, 480, 200, 400);
      p.L(200, 50);
      p.stroke();
      p.M(200, 400);
      p.C(200, 480, 300, 480, 300, 400);
      p.L(300, 50);
      p.stroke();
      p.M(300, 300);
      p.L(480, 100);
      p.stroke();
    }
  },
  // ♎ LIBRA
  [0x264e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 50);
      p.L(500, 50);
      p.stroke();
      p.M(100, 200);
      p.L(500, 200);
      p.stroke();
      p.M(300, 200);
      p.C(150, 200, 100, 350, 200, 420);
      p.stroke();
      p.M(300, 200);
      p.C(450, 200, 500, 350, 400, 420);
      p.stroke();
    }
  },
  // ♏ SCORPIO
  [0x264f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 50);
      p.L(100, 400);
      p.C(100, 480, 200, 480, 200, 400);
      p.L(200, 50);
      p.stroke();
      p.M(200, 400);
      p.C(200, 480, 300, 480, 300, 400);
      p.L(300, 50);
      p.stroke();
      p.M(300, 50);
      p.L(400, 50);
      p.L(350, 100);
      p.stroke();
    }
  },
  // ♐ SAGITTARIUS
  [0x2650]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 80);
      p.L(480, 440);
      p.stroke();
      p.M(480, 440);
      p.L(350, 440);
      p.stroke();
      p.M(480, 440);
      p.L(480, 310);
      p.stroke();
      p.M(200, 260);
      p.L(380, 160);
      p.stroke();
    }
  },
  // ♑ CAPRICORN
  [0x2651]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 450);
      p.L(100, 200);
      p.C(100, 100, 200, 50, 300, 100);
      p.C(400, 150, 450, 100, 450, 50);
      p.stroke();
      p.M(450, 50);
      p.C(500, 100, 500, 200, 420, 250);
      p.stroke();
    }
  },
  // ♒ AQUARIUS
  [0x2652]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 350);
      p.C(150, 400, 250, 300, 300, 350);
      p.C(350, 400, 450, 300, 520, 350);
      p.stroke();
      p.M(80, 200);
      p.C(150, 250, 250, 150, 300, 200);
      p.C(350, 250, 450, 150, 520, 200);
      p.stroke();
    }
  },
  // ♓ PISCES
  [0x2653]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 100);
      p.C(200, 250, 200, 250, 120, 400);
      p.stroke();
      p.M(480, 100);
      p.C(400, 250, 400, 250, 480, 400);
      p.stroke();
      p.M(120, 250);
      p.L(480, 250);
      p.stroke();
    }
  },
  // Card suits — add missing WHITE variants
  // ♡ WHITE HEART SUIT
  [0x2661]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 80);
      p.C(300, 80, 80, 200, 80, 340);
      p.C(80, 440, 180, 490, 300, 490);
      p.C(420, 490, 520, 440, 520, 340);
      p.C(520, 200, 300, 80, 300, 80);
      p.Z();
      p.stroke();
    }
  },
  // ♢ WHITE DIAMOND SUIT
  [0x2662]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 500);
      p.L(500, 250);
      p.L(300, 0);
      p.L(100, 250);
      p.Z();
      p.stroke();
    }
  },
  // ♤ WHITE SPADE SUIT
  [0x2664]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 500);
      p.C(300, 500, 80, 350, 80, 200);
      p.C(80, 100, 200, 50, 300, 180);
      p.C(400, 50, 520, 100, 520, 200);
      p.C(520, 350, 300, 500, 300, 500);
      p.Z();
      p.stroke();
      p.lineWidth(30);
      p.M(280, 0);
      p.L(320, 0);
      p.L(310, 140);
      p.L(290, 140);
      p.Z();
      p.stroke();
    }
  },
  // ♧ WHITE CLUB SUIT
  [0x2667]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 370, 100);
      p.stroke();
      p.circle(195, 240, 90);
      p.stroke();
      p.circle(405, 240, 90);
      p.stroke();
      p.lineWidth(30);
      p.M(280, 0);
      p.L(320, 0);
      p.L(310, 180);
      p.L(290, 180);
      p.Z();
      p.stroke();
    }
  },
  // ♻ BLACK UNIVERSAL RECYCLING SYMBOL
  [0x267b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      // three curved arrows forming a triangle
      const cx = 300;
      const cy = 240;
      const r = 180;
      for (let i = 0; i < 3; i++) {
        const a1 = ((i * 120 + 30) * Math.PI) / 180;
        const a2 = ((i * 120 + 90) * Math.PI) / 180;
        p.M(cx + r * Math.cos(a1), cy + r * Math.sin(a1));
        p.C(
          cx + r * 1.2 * Math.cos((a1 + a2) / 2),
          cy + r * 1.2 * Math.sin((a1 + a2) / 2),
          cx + r * 1.2 * Math.cos((a1 + a2) / 2),
          cy + r * 1.2 * Math.sin((a1 + a2) / 2),
          cx + r * Math.cos(a2),
          cy + r * Math.sin(a2)
        );
      }
      p.stroke();
      // arrowheads at each vertex
      for (let i = 0; i < 3; i++) {
        const a = ((i * 120 + 90) * Math.PI) / 180;
        const dir = ((i * 120 + 120) * Math.PI) / 180;
        const tx = cx + r * Math.cos(a);
        const ty = cy + r * Math.sin(a);
        p.M(tx, ty);
        p.L(tx + 40 * Math.cos(dir + 2.5), ty + 40 * Math.sin(dir + 2.5));
        p.L(tx + 40 * Math.cos(dir - 2.5), ty + 40 * Math.sin(dir - 2.5));
        p.Z();
        p.fill();
      }
    }
  },
  // ♾ PERMANENT PAPER SIGN (infinity in circle)
  [0x267e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // infinity shape
      p.M(300, 250);
      p.C(300, 350, 420, 380, 450, 300);
      p.C(480, 220, 420, 150, 380, 180);
      p.C(340, 210, 300, 250, 300, 250);
      p.C(300, 250, 260, 290, 220, 320);
      p.C(180, 350, 120, 280, 150, 200);
      p.C(180, 120, 300, 150, 300, 250);
      p.stroke();
    }
  },
  // ♿ WHEELCHAIR SYMBOL
  [0x267f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      // head
      p.circle(280, 420, 40);
      p.fill();
      // body/chair outline
      p.M(280, 380);
      p.L(280, 220);
      p.L(380, 220);
      p.stroke();
      // wheel
      p.circle(300, 160, 100);
      p.stroke();
      // leg rest
      p.M(380, 220);
      p.L(420, 80);
      p.stroke();
    }
  },
  // ⚒ HAMMER AND PICK
  [0x2692]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      // handle 1
      p.M(150, 80);
      p.L(400, 380);
      p.stroke();
      // handle 2
      p.M(450, 80);
      p.L(200, 380);
      p.stroke();
      // hammer head
      p.M(350, 350);
      p.L(450, 420);
      p.stroke();
      // pick head
      p.M(150, 420);
      p.L(250, 350);
      p.stroke();
    }
  },
  // ⚓ ANCHOR
  [0x2693]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      // ring at top
      p.circle(300, 440, 40);
      p.stroke();
      // vertical shaft
      p.M(300, 400);
      p.L(300, 80);
      p.stroke();
      // crossbar
      p.M(180, 320);
      p.L(420, 320);
      p.stroke();
      // flukes
      p.M(300, 80);
      p.C(150, 80, 100, 150, 130, 200);
      p.stroke();
      p.M(300, 80);
      p.C(450, 80, 500, 150, 470, 200);
      p.stroke();
    }
  },
  // ⚔ CROSSED SWORDS
  [0x2694]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 50);
      p.L(500, 450);
      p.stroke();
      p.M(500, 50);
      p.L(100, 450);
      p.stroke();
      // guards
      p.lineWidth(30);
      p.M(140, 130);
      p.L(200, 70);
      p.stroke();
      p.M(460, 130);
      p.L(400, 70);
      p.stroke();
    }
  },
  // ⚕ STAFF OF AESCULAPIUS
  [0x2695]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      // staff
      p.M(300, 20);
      p.L(300, 480);
      p.stroke();
      // snake (simplified as S-curve wrapping)
      p.lineWidth(30);
      p.M(300, 380);
      p.C(400, 380, 420, 320, 300, 300);
      p.C(180, 280, 200, 220, 300, 200);
      p.C(400, 180, 380, 120, 300, 120);
      p.stroke();
    }
  },
  // ⚖ SCALES
  [0x2696]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // beam
      p.M(100, 350);
      p.L(500, 350);
      p.stroke();
      // pillar
      p.M(300, 350);
      p.L(300, 80);
      p.stroke();
      // base
      p.M(200, 80);
      p.L(400, 80);
      p.stroke();
      // left pan
      p.M(100, 350);
      p.C(100, 250, 180, 220, 220, 250);
      p.stroke();
      // right pan
      p.M(500, 350);
      p.C(500, 250, 420, 220, 380, 250);
      p.stroke();
    }
  },
  // ⚗ ALEMBIC
  [0x2697]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // flask body
      p.M(200, 50);
      p.L(400, 50);
      p.L(480, 250);
      p.C(500, 300, 460, 350, 400, 350);
      p.L(200, 350);
      p.C(140, 350, 100, 300, 120, 250);
      p.L(200, 50);
      p.Z();
      p.stroke();
      // neck
      p.M(250, 350);
      p.L(250, 450);
      p.stroke();
      p.M(350, 350);
      p.L(350, 450);
      p.stroke();
    }
  },
  // ⚙ GEAR
  [0x2699]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.circle(300, 250, 100);
      p.stroke();
      // teeth
      const teeth = 8;
      const innerR = 130;
      const outerR = 200;
      for (let i = 0; i < teeth; i++) {
        const a = ((i * 360) / teeth) * (Math.PI / 180);
        const ha = (15 * Math.PI) / 180;
        p.M(300 + innerR * Math.cos(a - ha), 250 + innerR * Math.sin(a - ha));
        p.L(300 + outerR * Math.cos(a - ha * 0.7), 250 + outerR * Math.sin(a - ha * 0.7));
        p.L(300 + outerR * Math.cos(a + ha * 0.7), 250 + outerR * Math.sin(a + ha * 0.7));
        p.L(300 + innerR * Math.cos(a + ha), 250 + innerR * Math.sin(a + ha));
      }
      p.fill();
      // outer ring connecting teeth
      p.circle(300, 250, 130);
      p.stroke();
    }
  },
  // ⚛ ATOM SYMBOL
  [0x269b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      // three elliptical orbits at 60 degree offsets
      p.ellipse(300, 250, 220, 80);
      p.stroke();
      // rotate by drawing angled ellipses as simple ovals
      // orbit 2 (upper-left to lower-right)
      p.M(130, 100);
      p.C(200, 50, 440, 300, 470, 400);
      p.C(440, 450, 200, 200, 130, 100);
      p.stroke();
      // orbit 3 (upper-right to lower-left)
      p.M(470, 100);
      p.C(400, 50, 160, 300, 130, 400);
      p.C(160, 450, 400, 200, 470, 100);
      p.stroke();
      // nucleus
      p.circle(300, 250, 25);
      p.fill();
    }
  },
  // ⚜ FLEUR-DE-LIS
  [0x269c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      // center petal
      p.M(300, 480);
      p.C(280, 400, 260, 300, 300, 200);
      p.C(340, 300, 320, 400, 300, 480);
      p.stroke();
      // left petal
      p.M(300, 300);
      p.C(200, 350, 100, 400, 100, 350);
      p.C(100, 300, 180, 250, 250, 200);
      p.stroke();
      // right petal
      p.M(300, 300);
      p.C(400, 350, 500, 400, 500, 350);
      p.C(500, 300, 420, 250, 350, 200);
      p.stroke();
      // base bar
      p.M(200, 100);
      p.L(400, 100);
      p.stroke();
      // stem
      p.M(300, 200);
      p.L(300, 100);
      p.stroke();
    }
  },
  // ⚥ MALE AND FEMALE SIGN
  [0x26a5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(280, 270, 120);
      p.stroke();
      // female cross below
      p.M(280, 150);
      p.L(280, 30);
      p.stroke();
      p.M(220, 90);
      p.L(340, 90);
      p.stroke();
      // male arrow upper-right
      p.M(370, 350);
      p.L(490, 460);
      p.stroke();
      p.M(420, 460);
      p.L(490, 460);
      p.L(490, 390);
      p.stroke();
    }
  },
  // ⚧ MALE WITH STROKE AND MALE AND FEMALE SIGN (transgender symbol)
  [0x26a7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 220, 110);
      p.stroke();
      // arrow upper-right
      p.M(375, 295);
      p.L(480, 420);
      p.stroke();
      p.M(420, 420);
      p.L(480, 420);
      p.L(480, 360);
      p.stroke();
      // cross + stroke below
      p.M(300, 110);
      p.L(300, 20);
      p.stroke();
      p.M(250, 60);
      p.L(350, 60);
      p.stroke();
    }
  },
  // ⚰ COFFIN
  [0x26b0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(200, 480);
      p.L(140, 300);
      p.L(200, 20);
      p.L(400, 20);
      p.L(460, 300);
      p.L(400, 480);
      p.Z();
      p.stroke();
    }
  },
  // ⚽ SOCCER BALL
  [0x26bd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 200);
      p.stroke();
      // pentagon in center
      const cx = 300;
      const cy = 250;
      const r = 80;
      for (let i = 0; i < 5; i++) {
        const a = ((i * 72 - 90) * Math.PI) / 180;
        const a2 = (((i + 1) * 72 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + r * Math.cos(a), cy + r * Math.sin(a));
        }
        p.L(cx + r * Math.cos(a2), cy + r * Math.sin(a2));
      }
      p.Z();
      p.fill();
    }
  },
  // ⚾ BASEBALL
  [0x26be]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 200);
      p.stroke();
      // stitching curves
      p.lineWidth(25);
      p.M(150, 400);
      p.C(200, 300, 200, 200, 150, 100);
      p.stroke();
      p.M(450, 400);
      p.C(400, 300, 400, 200, 450, 100);
      p.stroke();
    }
  },
  // ⛄ SNOWMAN WITHOUT SNOW
  [0x26c4]: {
    width: W,
    draw: (p: GlyphPen) => {
      // two stacked circles (no base snow)
      p.circle(300, 140, 100);
      p.fill();
      p.circle(300, 350, 70);
      p.fill();
    }
  },
  // ⛅ SUN BEHIND CLOUD
  [0x26c5]: {
    width: W,
    draw: (p: GlyphPen) => {
      // sun (partial)
      p.lineWidth(30);
      p.M(420, 420);
      p.C(480, 420, 520, 380, 520, 320);
      p.stroke();
      for (let a = -30; a <= 60; a += 30) {
        const rad = (a * Math.PI) / 180;
        p.M(460 + 60 * Math.cos(rad), 360 + 60 * Math.sin(rad));
        p.L(460 + 100 * Math.cos(rad), 360 + 100 * Math.sin(rad));
      }
      p.stroke();
      // cloud in front
      p.lineWidth(35);
      p.M(100, 220);
      p.C(100, 370, 180, 400, 260, 380);
      p.C(280, 430, 380, 430, 400, 380);
      p.C(460, 380, 480, 300, 440, 220);
      p.L(100, 220);
      p.fill();
      p.M(100, 220);
      p.C(100, 370, 180, 400, 260, 380);
      p.C(280, 430, 380, 430, 400, 380);
      p.C(460, 380, 480, 300, 440, 220);
      p.L(100, 220);
      p.stroke();
    }
  },
  // ⛎ OPHIUCHUS
  [0x26ce]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      // U-shape with snake
      p.M(150, 450);
      p.L(150, 200);
      p.C(150, 50, 450, 50, 450, 200);
      p.L(450, 450);
      p.stroke();
      // cross line
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // ⛔ NO ENTRY
  [0x26d4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 220);
      p.fill();
      // We draw filled circle (in real font the bar would be white)
    }
  },
  // ⛪ CHURCH
  [0x26ea]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // building
      p.rect(150, 50, 300, 250);
      p.stroke();
      // roof
      p.M(130, 300);
      p.L(300, 420);
      p.L(470, 300);
      p.Z();
      p.stroke();
      // cross on top
      p.M(300, 420);
      p.L(300, 490);
      p.stroke();
      p.M(270, 460);
      p.L(330, 460);
      p.stroke();
      // door
      p.M(260, 50);
      p.L(260, 150);
      p.L(340, 150);
      p.L(340, 50);
      p.stroke();
    }
  },
  // ⛲ FOUNTAIN
  [0x26f2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // base
      p.M(100, 50);
      p.L(500, 50);
      p.stroke();
      p.M(150, 50);
      p.L(150, 200);
      p.stroke();
      p.M(450, 50);
      p.L(450, 200);
      p.stroke();
      // basin
      p.M(100, 200);
      p.L(500, 200);
      p.stroke();
      // pillar
      p.M(300, 200);
      p.L(300, 380);
      p.stroke();
      // water arcs
      p.M(300, 380);
      p.C(200, 420, 150, 300, 150, 200);
      p.stroke();
      p.M(300, 380);
      p.C(400, 420, 450, 300, 450, 200);
      p.stroke();
    }
  },
  // ⛳ FLAG IN HOLE
  [0x26f3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // pole
      p.M(250, 30);
      p.L(250, 480);
      p.stroke();
      // flag
      p.M(250, 480);
      p.L(450, 420);
      p.L(250, 360);
      p.Z();
      p.fill();
      // ground
      p.M(120, 30);
      p.C(200, 60, 300, 60, 380, 30);
      p.stroke();
    }
  },
  // ⛵ SAILBOAT
  [0x26f5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // hull
      p.M(80, 150);
      p.L(520, 150);
      p.L(450, 50);
      p.L(150, 50);
      p.Z();
      p.stroke();
      // mast
      p.M(300, 150);
      p.L(300, 470);
      p.stroke();
      // sail
      p.M(300, 460);
      p.L(450, 200);
      p.L(300, 200);
      p.Z();
      p.stroke();
    }
  },
  // ⛺ TENT
  [0x26fa]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 50);
      p.L(300, 450);
      p.L(520, 50);
      p.Z();
      p.stroke();
      // center line
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
    }
  }
};

// =============================================================================
// Misc Technical Extended (U+2300–U+23FF) — remaining chars
// =============================================================================

export const MISC_TECHNICAL_EXT: Record<number, GlyphDef> = {
  // 0x2301 ELECTRIC ARROW
  [0x2301]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(200, 350);
      p.L(280, 200);
      p.L(380, 350);
      p.L(450, 250);
      p.stroke();
      p.M(450, 250);
      p.L(500, 320);
      p.L(500, 180);
      p.Z();
      p.fill();
    }
  },
  // 0x2303 UP ARROWHEAD
  [0x2303]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(130, 100);
      p.L(300, 400);
      p.L(470, 100);
      p.stroke();
    }
  },
  // 0x2304 DOWN ARROWHEAD
  [0x2304]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(130, 400);
      p.L(300, 100);
      p.L(470, 400);
      p.stroke();
    }
  },
  // 0x2305 PROJECTIVE
  [0x2305]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 100);
      p.L(500, 100);
      p.stroke();
      p.M(300, 100);
      p.L(300, 450);
      p.stroke();
      p.M(200, 350);
      p.L(400, 350);
      p.stroke();
    }
  },
  // 0x2306 PERSPECTIVE
  [0x2306]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 100);
      p.L(500, 100);
      p.stroke();
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(300, 100);
      p.L(300, 450);
      p.stroke();
    }
  },
  // 0x2307 WAVY LINE
  [0x2307]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.C(180, 350, 260, 150, 340, 250);
      p.C(420, 350, 500, 150, 520, 250);
      p.stroke();
    }
  },
  // ⌈ LEFT CEILING
  [0x2308]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(400, 450);
      p.L(200, 450);
      p.L(200, 50);
      p.stroke();
    }
  },
  // ⌉ RIGHT CEILING
  [0x2309]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(200, 450);
      p.L(400, 450);
      p.L(400, 50);
      p.stroke();
    }
  },
  // ⌊ LEFT FLOOR
  [0x230a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(200, 450);
      p.L(200, 50);
      p.L(400, 50);
      p.stroke();
    }
  },
  // ⌋ RIGHT FLOOR
  [0x230b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(400, 450);
      p.L(400, 50);
      p.L(200, 50);
      p.stroke();
    }
  },
  // 0x230C BOTTOM RIGHT CROP
  [0x230c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 100);
      p.L(450, 100);
      p.L(450, 250);
      p.stroke();
    }
  },
  // 0x230D BOTTOM LEFT CROP
  [0x230d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 100);
      p.L(150, 100);
      p.L(150, 250);
      p.stroke();
    }
  },
  // 0x230E TOP RIGHT CROP
  [0x230e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 400);
      p.L(450, 400);
      p.L(450, 250);
      p.stroke();
    }
  },
  // 0x230F TOP LEFT CROP
  [0x230f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 400);
      p.L(150, 400);
      p.L(150, 250);
      p.stroke();
    }
  },
  // 0x2311 SQUARE LOZENGE
  [0x2311]: {
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
  },
  // ⌒ ARC
  [0x2312]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(80, 100);
      p.C(80, 400, 520, 400, 520, 100);
      p.stroke();
    }
  },
  // ⌓ SEGMENT
  [0x2313]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 200);
      p.C(80, 450, 520, 450, 520, 200);
      p.stroke();
      p.M(80, 200);
      p.L(520, 200);
      p.stroke();
    }
  },
  // ⌔ SECTOR
  [0x2314]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 100);
      p.L(120, 400);
      p.C(200, 500, 400, 500, 480, 400);
      p.L(300, 100);
      p.stroke();
    }
  },
  // 0x2316 POSITION INDICATOR
  [0x2316]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 120);
      p.stroke();
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // 0x2317 VIEWDATA SQUARE
  [0x2317]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(120, 70, 360, 360);
      p.stroke();
      p.M(300, 70);
      p.L(300, 430);
      p.stroke();
      p.M(120, 250);
      p.L(480, 250);
      p.stroke();
    }
  },
  // 0x2319 TURNED NOT SIGN
  [0x2319]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(120, 350);
      p.L(480, 350);
      p.L(480, 150);
      p.stroke();
    }
  },
  // ⌜ TOP LEFT CORNER
  [0x231c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 300);
      p.L(150, 420);
      p.L(300, 420);
      p.stroke();
    }
  },
  // ⌝ TOP RIGHT CORNER
  [0x231d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 420);
      p.L(450, 420);
      p.L(450, 300);
      p.stroke();
    }
  },
  // ⌞ BOTTOM LEFT CORNER
  [0x231e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 200);
      p.L(150, 80);
      p.L(300, 80);
      p.stroke();
    }
  },
  // ⌟ BOTTOM RIGHT CORNER
  [0x231f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 80);
      p.L(450, 80);
      p.L(450, 200);
      p.stroke();
    }
  },
  // ⌠ TOP HALF INTEGRAL
  [0x2320]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(380, 480);
      p.C(350, 480, 300, 450, 300, 400);
      p.L(300, 50);
      p.stroke();
    }
  },
  // ⌡ BOTTOM HALF INTEGRAL
  [0x2321]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 450);
      p.L(300, 100);
      p.C(300, 50, 250, 20, 220, 20);
      p.stroke();
    }
  },
  // ⌢ FROWN
  [0x2322]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(80, 350);
      p.C(200, 100, 400, 100, 520, 350);
      p.stroke();
    }
  },
  // ⌣ SMILE
  [0x2323]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(80, 200);
      p.C(200, 450, 400, 450, 520, 200);
      p.stroke();
    }
  },
  // ⌤ UP ARROWHEAD BETWEEN TWO HORIZONTAL BARS
  [0x2324]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 100);
      p.L(500, 100);
      p.stroke();
      p.M(100, 400);
      p.L(500, 400);
      p.stroke();
      p.M(200, 200);
      p.L(300, 350);
      p.L(400, 200);
      p.stroke();
    }
  },
  // ⌥ OPTION KEY
  [0x2325]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 100);
      p.L(250, 100);
      p.L(400, 400);
      p.L(500, 400);
      p.stroke();
      p.M(300, 400);
      p.L(500, 400);
      p.stroke();
      p.M(100, 400);
      p.L(300, 400);
      p.stroke();
    }
  },
  // ⌦ ERASE TO THE RIGHT
  [0x2326]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 250);
      p.L(200, 420);
      p.L(520, 420);
      p.L(520, 80);
      p.L(200, 80);
      p.Z();
      p.stroke();
      p.M(280, 180);
      p.L(440, 320);
      p.stroke();
      p.M(280, 320);
      p.L(440, 180);
      p.stroke();
    }
  },
  // ⌧ X IN A RECTANGLE BOX
  [0x2327]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 70, 400, 360);
      p.stroke();
      p.M(180, 150);
      p.L(420, 350);
      p.stroke();
      p.M(180, 350);
      p.L(420, 150);
      p.stroke();
    }
  },
  // ⌫ ERASE TO THE LEFT (DELETE)
  [0x232b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(520, 250);
      p.L(400, 420);
      p.L(80, 420);
      p.L(80, 80);
      p.L(400, 80);
      p.Z();
      p.stroke();
      p.M(160, 180);
      p.L(320, 320);
      p.stroke();
      p.M(160, 320);
      p.L(320, 180);
      p.stroke();
    }
  },
  // 0x2329 LEFT-POINTING ANGLE BRACKET
  [0x2329]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 450);
      p.L(200, 250);
      p.L(400, 50);
      p.stroke();
    }
  },
  // 0x232A RIGHT-POINTING ANGLE BRACKET
  [0x232a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 450);
      p.L(400, 250);
      p.L(200, 50);
      p.stroke();
    }
  },
  // 0x232C BENZENE RING
  [0x232c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const cx = 300,
        cy = 250,
        r = 180;
      for (let i = 0; i < 6; i++) {
        const a1 = ((i * 60 - 90) * Math.PI) / 180;
        const a2 = (((i + 1) * 60 - 90) * Math.PI) / 180;
        p.M(cx + r * Math.cos(a1), cy + r * Math.sin(a1));
        p.L(cx + r * Math.cos(a2), cy + r * Math.sin(a2));
      }
      p.stroke();
      p.circle(cx, cy, 100);
      p.stroke();
    }
  },
  // 0x232D CYLINDRICITY
  [0x232d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 420);
      p.L(150, 80);
      p.stroke();
      p.M(450, 420);
      p.L(450, 80);
      p.stroke();
      p.ellipse(300, 420, 150, 40);
      p.stroke();
      p.ellipse(300, 80, 150, 40);
      p.stroke();
    }
  },
  // 0x232E ALL AROUND PROFILE
  [0x232e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 200);
      p.stroke();
      p.M(300, 450);
      p.L(300, 480);
      p.stroke();
    }
  },
  // 0x232F SYMMETRY
  [0x232f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(300, 100);
      p.L(300, 400);
      p.stroke();
      p.M(250, 350);
      p.L(300, 400);
      p.L(350, 350);
      p.stroke();
    }
  },
  // 0x2330 CORNER WITH HORIZONTAL BAR
  [0x2330]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 400);
      p.L(150, 100);
      p.L(450, 100);
      p.stroke();
      p.M(150, 250);
      p.L(350, 250);
      p.stroke();
    }
  },
  // 0x2331 BOTTOM LEFT TO TOP RIGHT
  [0x2331]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 100);
      p.L(500, 400);
      p.stroke();
    }
  },
  // 0x2332 CONICAL TAPER
  [0x2332]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 350);
      p.L(500, 300);
      p.stroke();
      p.M(100, 150);
      p.L(500, 200);
      p.stroke();
    }
  },
  // 0x2333 SLOPE
  [0x2333]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 100);
      p.L(100, 400);
      p.L(500, 250);
      p.stroke();
    }
  },
  // 0x2334 COUNTERBORE
  [0x2334]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(200, 400);
      p.L(200, 250);
      p.L(100, 250);
      p.L(100, 100);
      p.L(500, 100);
      p.L(500, 250);
      p.L(400, 250);
      p.L(400, 400);
      p.stroke();
    }
  },
  // 0x2335 COUNTERSINK
  [0x2335]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(200, 100);
      p.L(100, 300);
      p.L(100, 400);
      p.stroke();
      p.M(400, 100);
      p.L(500, 300);
      p.L(500, 400);
      p.stroke();
      p.M(200, 100);
      p.L(400, 100);
      p.stroke();
    }
  },
  // ⌶ APL FUNCTIONAL SYMBOL I-BEAM
  [0x2336]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 420);
      p.L(450, 420);
      p.stroke();
      p.M(150, 80);
      p.L(450, 80);
      p.stroke();
      p.M(300, 420);
      p.L(300, 80);
      p.stroke();
    }
  },
  // ⌷ APL FUNCTIONAL SYMBOL SQUISH QUAD
  [0x2337]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(150, 80, 300, 340);
      p.stroke();
    }
  },
  // ⌸ APL FUNCTIONAL SYMBOL QUAD EQUAL
  [0x2338]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(180, 300);
      p.L(420, 300);
      p.stroke();
      p.M(180, 200);
      p.L(420, 200);
      p.stroke();
    }
  },
  // ⌹ APL FUNCTIONAL SYMBOL QUAD DIVIDE
  [0x2339]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(180, 250);
      p.L(420, 250);
      p.stroke();
      p.circle(300, 340, 30);
      p.fill();
      p.circle(300, 160, 30);
      p.fill();
    }
  },
  // ⌺ APL FUNCTIONAL SYMBOL QUAD DIAMOND
  [0x233a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(300, 380);
      p.L(420, 250);
      p.L(300, 120);
      p.L(180, 250);
      p.Z();
      p.stroke();
    }
  },
  // ⌻ APL FUNCTIONAL SYMBOL QUAD JOT
  [0x233b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.circle(300, 250, 80);
      p.stroke();
    }
  },
  // ⌼ APL FUNCTIONAL SYMBOL QUAD CIRCLE
  [0x233c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.circle(300, 250, 130);
      p.stroke();
    }
  },
  // ⌽ APL FUNCTIONAL SYMBOL CIRCLE STILE
  [0x233d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 180);
      p.stroke();
      p.M(300, 430);
      p.L(300, 70);
      p.stroke();
    }
  },
  // ⌾ APL FUNCTIONAL SYMBOL CIRCLE JOT
  [0x233e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 180);
      p.stroke();
      p.circle(300, 250, 70);
      p.stroke();
    }
  },
  // ⌿ APL FUNCTIONAL SYMBOL SLASH BAR
  [0x233f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 420);
      p.L(200, 80);
      p.stroke();
      p.M(120, 250);
      p.L(480, 250);
      p.stroke();
    }
  },
  // ⍀ APL FUNCTIONAL SYMBOL BACKSLASH BAR
  [0x2340]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 420);
      p.L(400, 80);
      p.stroke();
      p.M(120, 250);
      p.L(480, 250);
      p.stroke();
    }
  },
  // ⍁ APL FUNCTIONAL SYMBOL QUAD SLASH
  [0x2341]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(400, 380);
      p.L(200, 120);
      p.stroke();
    }
  },
  // ⍂ APL FUNCTIONAL SYMBOL QUAD BACKSLASH
  [0x2342]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(200, 380);
      p.L(400, 120);
      p.stroke();
    }
  },
  // ⍃ APL FUNCTIONAL SYMBOL QUAD LESS-THAN
  [0x2343]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(380, 370);
      p.L(220, 250);
      p.L(380, 130);
      p.stroke();
    }
  },
  // ⍄ APL FUNCTIONAL SYMBOL QUAD GREATER-THAN
  [0x2344]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(220, 370);
      p.L(380, 250);
      p.L(220, 130);
      p.stroke();
    }
  },
  // ⍅ APL FUNCTIONAL SYMBOL LEFTWARDS VANE
  [0x2345]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(250, 400);
      p.stroke();
    }
  },
  // ⍆ APL FUNCTIONAL SYMBOL RIGHTWARDS VANE
  [0x2346]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(350, 400);
      p.stroke();
    }
  },
  // ⍇ APL FUNCTIONAL SYMBOL QUAD LEFTWARDS ARROW
  [0x2347]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(420, 250);
      p.L(180, 250);
      p.stroke();
      p.M(180, 250);
      p.L(270, 330);
      p.L(270, 170);
      p.Z();
      p.fill();
    }
  },
  // ⍈ APL FUNCTIONAL SYMBOL QUAD RIGHTWARDS ARROW
  [0x2348]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(180, 250);
      p.L(420, 250);
      p.stroke();
      p.M(420, 250);
      p.L(330, 330);
      p.L(330, 170);
      p.Z();
      p.fill();
    }
  },
  // ⍉ APL FUNCTIONAL SYMBOL CIRCLE BACKSLASH
  [0x2349]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 180);
      p.stroke();
      p.M(170, 120);
      p.L(430, 380);
      p.stroke();
    }
  },
  // ⍊ APL FUNCTIONAL SYMBOL DOWN TACK UNDERBAR
  [0x234a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 80);
      p.L(500, 80);
      p.stroke();
      p.M(300, 80);
      p.L(300, 350);
      p.stroke();
      p.M(150, 350);
      p.L(450, 350);
      p.stroke();
      p.M(100, 420);
      p.L(500, 420);
      p.stroke();
    }
  },
  // ⍋ APL FUNCTIONAL SYMBOL DELTA STILE
  [0x234b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 420);
      p.L(150, 120);
      p.L(450, 120);
      p.Z();
      p.stroke();
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
    }
  },
  // ⍌ APL FUNCTIONAL SYMBOL QUAD DOWN CARET
  [0x234c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(200, 350);
      p.L(300, 150);
      p.L(400, 350);
      p.stroke();
    }
  },
  // ⍍ APL FUNCTIONAL SYMBOL QUAD DELTA
  [0x234d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(300, 370);
      p.L(190, 140);
      p.L(410, 140);
      p.Z();
      p.stroke();
    }
  },
  // ⍎ APL FUNCTIONAL SYMBOL DOWN TACK JOT
  [0x234e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 400);
      p.L(450, 400);
      p.stroke();
      p.M(300, 400);
      p.L(300, 150);
      p.stroke();
      p.circle(300, 100, 50);
      p.stroke();
    }
  },
  // ⍏ APL FUNCTIONAL SYMBOL UPWARDS VANE
  [0x234f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 80);
      p.L(300, 420);
      p.stroke();
      p.M(300, 420);
      p.L(180, 300);
      p.stroke();
    }
  },
  // ⍐ APL FUNCTIONAL SYMBOL QUAD UPWARDS ARROW
  [0x2350]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(300, 120);
      p.L(300, 380);
      p.stroke();
      p.M(300, 380);
      p.L(220, 300);
      p.L(380, 300);
      p.Z();
      p.fill();
    }
  },
  // ⍑ APL FUNCTIONAL SYMBOL UP TACK OVERBAR
  [0x2351]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 420);
      p.L(500, 420);
      p.stroke();
      p.M(300, 420);
      p.L(300, 150);
      p.stroke();
      p.M(150, 150);
      p.L(450, 150);
      p.stroke();
      p.M(100, 80);
      p.L(500, 80);
      p.stroke();
    }
  },
  // ⍒ APL FUNCTIONAL SYMBOL DEL STILE
  [0x2352]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 80);
      p.L(150, 380);
      p.L(450, 380);
      p.Z();
      p.stroke();
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
    }
  },
  // ⍓ APL FUNCTIONAL SYMBOL QUAD UP CARET
  [0x2353]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(200, 150);
      p.L(300, 350);
      p.L(400, 150);
      p.stroke();
    }
  },
  // ⍔ APL FUNCTIONAL SYMBOL QUAD DEL
  [0x2354]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(300, 130);
      p.L(190, 360);
      p.L(410, 360);
      p.Z();
      p.stroke();
    }
  },
  // ⍕ APL FUNCTIONAL SYMBOL UP TACK JOT
  [0x2355]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 100);
      p.L(450, 100);
      p.stroke();
      p.M(300, 100);
      p.L(300, 350);
      p.stroke();
      p.circle(300, 400, 50);
      p.stroke();
    }
  },
  // ⍖ APL FUNCTIONAL SYMBOL DOWNWARDS VANE
  [0x2356]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 420);
      p.L(300, 80);
      p.stroke();
      p.M(300, 80);
      p.L(420, 200);
      p.stroke();
    }
  },
  // ⍗ APL FUNCTIONAL SYMBOL QUAD DOWNWARDS ARROW
  [0x2357]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(300, 380);
      p.L(300, 120);
      p.stroke();
      p.M(300, 120);
      p.L(220, 200);
      p.L(380, 200);
      p.Z();
      p.fill();
    }
  },
  // ⍘ APL FUNCTIONAL SYMBOL UNDERBAR QUOTE
  [0x2358]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(260, 420);
      p.L(260, 200);
      p.stroke();
      p.M(340, 420);
      p.L(340, 200);
      p.stroke();
      p.M(120, 100);
      p.L(480, 100);
      p.stroke();
    }
  },
  // ⍙ APL FUNCTIONAL SYMBOL DELTA UNDERBAR
  [0x2359]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 430);
      p.L(150, 160);
      p.L(450, 160);
      p.Z();
      p.stroke();
      p.M(120, 80);
      p.L(480, 80);
      p.stroke();
    }
  },
  // ⍚ APL FUNCTIONAL SYMBOL DIAMOND UNDERBAR
  [0x235a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 430);
      p.L(430, 280);
      p.L(300, 150);
      p.L(170, 280);
      p.Z();
      p.stroke();
      p.M(120, 80);
      p.L(480, 80);
      p.stroke();
    }
  },
  // ⍛ APL FUNCTIONAL SYMBOL JOT UNDERBAR
  [0x235b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 310, 100);
      p.stroke();
      p.M(120, 100);
      p.L(480, 100);
      p.stroke();
    }
  },
  // ⍜ APL FUNCTIONAL SYMBOL CIRCLE UNDERBAR
  [0x235c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 310, 130);
      p.stroke();
      p.M(100, 100);
      p.L(500, 100);
      p.stroke();
    }
  },
  // ⍝ APL FUNCTIONAL SYMBOL UP SHOE JOT
  [0x235d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 100);
      p.C(120, 350, 480, 350, 480, 100);
      p.stroke();
      p.circle(300, 300, 50);
      p.stroke();
    }
  },
  // ⍞ APL FUNCTIONAL SYMBOL QUOTE QUAD
  [0x235e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(100, 250);
      p.L(50, 250);
      p.stroke();
    }
  },
  // ⍟ APL FUNCTIONAL SYMBOL CIRCLE STAR
  [0x235f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 180);
      p.stroke();
      // small 5-pointed star inside
      const cx = 300,
        cy = 250,
        R = 100,
        r2 = 40;
      for (let i = 0; i < 5; i++) {
        const a1 = ((i * 72 - 90) * Math.PI) / 180;
        const a2 = ((i * 72 + 36 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        } else {
          p.L(cx + R * Math.cos(a1), cy + R * Math.sin(a1));
        }
        p.L(cx + r2 * Math.cos(a2), cy + r2 * Math.sin(a2));
      }
      p.Z();
      p.stroke();
    }
  },
  // ⍠ APL FUNCTIONAL SYMBOL QUAD COLON
  [0x2360]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.circle(300, 330, 35);
      p.fill();
      p.circle(300, 170, 35);
      p.fill();
    }
  },
  // ⍡ APL FUNCTIONAL SYMBOL UP TACK DIAERESIS
  [0x2361]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 150);
      p.L(450, 150);
      p.stroke();
      p.M(300, 150);
      p.L(300, 350);
      p.stroke();
      p.circle(250, 420, 30);
      p.fill();
      p.circle(350, 420, 30);
      p.fill();
    }
  },
  // ⍢ APL FUNCTIONAL SYMBOL DEL DIAERESIS
  [0x2362]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 120);
      p.L(150, 370);
      p.L(450, 370);
      p.Z();
      p.stroke();
      p.circle(250, 80, 25);
      p.fill();
      p.circle(350, 80, 25);
      p.fill();
    }
  },
  // ⍣ APL FUNCTIONAL SYMBOL STAR DIAERESIS
  [0x2363]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 380);
      p.L(300, 180);
      p.stroke();
      p.M(200, 340);
      p.L(400, 220);
      p.stroke();
      p.M(200, 220);
      p.L(400, 340);
      p.stroke();
      p.circle(250, 430, 25);
      p.fill();
      p.circle(350, 430, 25);
      p.fill();
    }
  },
  // ⍤ APL FUNCTIONAL SYMBOL JOT DIAERESIS
  [0x2364]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 200, 80);
      p.stroke();
      p.circle(250, 400, 30);
      p.fill();
      p.circle(350, 400, 30);
      p.fill();
    }
  },
  // ⍥ APL FUNCTIONAL SYMBOL CIRCLE DIAERESIS
  [0x2365]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 200, 120);
      p.stroke();
      p.circle(250, 420, 30);
      p.fill();
      p.circle(350, 420, 30);
      p.fill();
    }
  },
  // ⍦ APL FUNCTIONAL SYMBOL DOWN SHOE STILE
  [0x2366]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 400);
      p.C(120, 150, 480, 150, 480, 400);
      p.stroke();
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
    }
  },
  // ⍧ APL FUNCTIONAL SYMBOL LEFT SHOE STILE
  [0x2367]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(450, 100);
      p.C(200, 100, 200, 400, 450, 400);
      p.stroke();
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // ⍨ APL FUNCTIONAL SYMBOL TILDE DIAERESIS
  [0x2368]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 250);
      p.C(200, 350, 400, 150, 480, 250);
      p.stroke();
      p.circle(250, 400, 30);
      p.fill();
      p.circle(350, 400, 30);
      p.fill();
    }
  },
  // ⍩ APL FUNCTIONAL SYMBOL GREATER-THAN DIAERESIS
  [0x2369]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(180, 370);
      p.L(380, 250);
      p.L(180, 130);
      p.stroke();
      p.circle(430, 200, 25);
      p.fill();
      p.circle(430, 300, 25);
      p.fill();
    }
  },
  // ⍪ APL FUNCTIONAL SYMBOL COMMA BAR
  [0x236a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.circle(300, 300, 40);
      p.fill();
      p.M(300, 260);
      p.L(260, 180);
      p.stroke();
      p.M(120, 120);
      p.L(480, 120);
      p.stroke();
    }
  },
  // ⍫ APL FUNCTIONAL SYMBOL DEL TILDE
  [0x236b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 80);
      p.L(150, 380);
      p.L(450, 380);
      p.Z();
      p.stroke();
      p.lineWidth(35);
      p.M(220, 250);
      p.C(260, 300, 340, 200, 380, 250);
      p.stroke();
    }
  },
  // ⍬ APL FUNCTIONAL SYMBOL ZILDE
  [0x236c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.ellipse(300, 250, 150, 200);
      p.stroke();
      p.M(200, 250);
      p.L(400, 250);
      p.stroke();
    }
  },
  // ⍭ APL FUNCTIONAL SYMBOL STILE TILDE
  [0x236d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 80);
      p.L(300, 420);
      p.stroke();
      p.M(180, 250);
      p.C(220, 320, 380, 180, 420, 250);
      p.stroke();
    }
  },
  // ⍮ APL FUNCTIONAL SYMBOL SEMICOLON UNDERBAR
  [0x236e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 350, 35);
      p.fill();
      p.circle(300, 250, 35);
      p.fill();
      p.M(300, 215);
      p.L(270, 160);
      p.stroke();
      p.lineWidth(40);
      p.M(150, 100);
      p.L(450, 100);
      p.stroke();
    }
  },
  // ⍯ APL FUNCTIONAL SYMBOL QUAD NOT EQUAL
  [0x236f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(200, 310);
      p.L(400, 310);
      p.stroke();
      p.M(200, 190);
      p.L(400, 190);
      p.stroke();
      p.M(350, 370);
      p.L(250, 130);
      p.stroke();
    }
  },
  // ⍰ APL FUNCTIONAL SYMBOL QUAD QUESTION
  [0x2370]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.lineWidth(35);
      p.M(230, 350);
      p.C(230, 400, 370, 400, 370, 340);
      p.C(370, 290, 300, 270, 300, 220);
      p.stroke();
      p.circle(300, 160, 25);
      p.fill();
    }
  },
  // ⍱ APL FUNCTIONAL SYMBOL DOWN CARET TILDE
  [0x2371]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(130, 400);
      p.L(300, 130);
      p.L(470, 400);
      p.stroke();
      p.lineWidth(35);
      p.M(220, 280);
      p.C(260, 330, 340, 230, 380, 280);
      p.stroke();
    }
  },
  // ⍲ APL FUNCTIONAL SYMBOL UP CARET TILDE
  [0x2372]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(130, 100);
      p.L(300, 370);
      p.L(470, 100);
      p.stroke();
      p.lineWidth(35);
      p.M(220, 220);
      p.C(260, 270, 340, 170, 380, 220);
      p.stroke();
    }
  },
  // ⍳ APL FUNCTIONAL SYMBOL IOTA
  [0x2373]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 400);
      p.L(250, 150);
      p.C(250, 80, 350, 80, 350, 150);
      p.stroke();
    }
  },
  // ⍴ APL FUNCTIONAL SYMBOL RHO
  [0x2374]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 80);
      p.L(200, 400);
      p.C(200, 450, 400, 450, 400, 400);
      p.C(400, 300, 200, 300, 200, 350);
      p.stroke();
    }
  },
  // ⍵ APL FUNCTIONAL SYMBOL OMEGA
  [0x2375]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 400);
      p.L(200, 150);
      p.C(200, 80, 400, 80, 400, 150);
      p.L(480, 400);
      p.stroke();
    }
  },
  // ⍶ APL FUNCTIONAL SYMBOL ALPHA UNDERBAR
  [0x2376]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(400, 400);
      p.C(400, 300, 200, 300, 200, 350);
      p.C(200, 400, 400, 400, 400, 350);
      p.L(400, 200);
      p.stroke();
      p.M(120, 120);
      p.L(480, 120);
      p.stroke();
    }
  },
  // ⍷ APL FUNCTIONAL SYMBOL EPSILON UNDERBAR
  [0x2377]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(420, 400);
      p.L(220, 400);
      p.C(160, 400, 160, 200, 220, 200);
      p.L(420, 200);
      p.stroke();
      p.M(220, 300);
      p.L(380, 300);
      p.stroke();
      p.M(120, 120);
      p.L(480, 120);
      p.stroke();
    }
  },
  // ⍸ APL FUNCTIONAL SYMBOL IOTA UNDERBAR
  [0x2378]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(250, 430);
      p.L(250, 200);
      p.C(250, 150, 350, 150, 350, 200);
      p.stroke();
      p.M(120, 100);
      p.L(480, 100);
      p.stroke();
    }
  },
  // ⍹ APL FUNCTIONAL SYMBOL OMEGA UNDERBAR
  [0x2379]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 420);
      p.L(200, 200);
      p.C(200, 150, 400, 150, 400, 200);
      p.L(480, 420);
      p.stroke();
      p.M(120, 100);
      p.L(480, 100);
      p.stroke();
    }
  },
  // ⍺ APL FUNCTIONAL SYMBOL ALPHA
  [0x237a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 400);
      p.C(400, 300, 200, 300, 200, 350);
      p.C(200, 400, 400, 400, 400, 350);
      p.L(400, 150);
      p.stroke();
    }
  },
  // ⍻ NOT CHECK MARK
  [0x237b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(130, 250);
      p.L(250, 100);
      p.L(470, 400);
      p.stroke();
      p.M(180, 380);
      p.L(420, 120);
      p.stroke();
    }
  },
  // ⍼ RIGHT ANGLE WITH DOWNWARDS ZIGZAG
  [0x237c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 400);
      p.L(150, 100);
      p.L(450, 100);
      p.stroke();
      p.M(150, 100);
      p.L(200, 200);
      p.L(250, 100);
      p.L(300, 200);
      p.stroke();
    }
  },
  // ⍽ SHOULDERED OPEN BOX (with open bottom)
  [0x237d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 350);
      p.L(100, 150);
      p.L(200, 150);
      p.stroke();
      p.M(400, 150);
      p.L(500, 150);
      p.L(500, 350);
      p.stroke();
    }
  },
  // ⍾ BELL SYMBOL
  [0x237e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 100);
      p.C(150, 350, 300, 450, 300, 450);
      p.C(300, 450, 450, 350, 450, 100);
      p.stroke();
      p.M(120, 100);
      p.L(480, 100);
      p.stroke();
      p.circle(300, 60, 25);
      p.fill();
    }
  },
  // ⍿ INSERTION SYMBOL
  [0x237f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 80);
      p.L(300, 420);
      p.L(400, 80);
      p.stroke();
    }
  },
  // 0x2380 INSERTION SYMBOL (should be different from 237F but close)
  [0x2380]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 100);
      p.L(300, 380);
      p.L(400, 100);
      p.stroke();
      p.M(150, 100);
      p.L(450, 100);
      p.stroke();
    }
  },
  // 0x2381 CONTINUOUS UNDERLINE
  [0x2381]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(80, 100);
      p.L(520, 100);
      p.stroke();
    }
  },
  // 0x2382 DISCONTINUOUS UNDERLINE
  [0x2382]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(80, 100);
      p.L(230, 100);
      p.stroke();
      p.M(370, 100);
      p.L(520, 100);
      p.stroke();
    }
  },
  // 0x2383 EMPHASIS SYMBOL
  [0x2383]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(80, 100);
      p.L(520, 100);
      p.stroke();
      p.circle(300, 200, 30);
      p.fill();
    }
  },
  // 0x2384 COMPOSITION SYMBOL
  [0x2384]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 300, 120);
      p.stroke();
      p.M(300, 180);
      p.L(300, 420);
      p.stroke();
    }
  },
  // 0x2385 WHITE SQUARE WITH CENTRE VERTICAL LINE
  [0x2385]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(120, 70, 360, 360);
      p.stroke();
      p.M(300, 70);
      p.L(300, 430);
      p.stroke();
    }
  },
  // 0x2386 ENTER SYMBOL
  [0x2386]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(420, 400);
      p.L(420, 200);
      p.L(180, 200);
      p.stroke();
      p.M(180, 200);
      p.L(280, 280);
      p.L(280, 120);
      p.Z();
      p.fill();
    }
  },
  // 0x2387 ALTERNATIVE KEY SYMBOL
  [0x2387]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 150);
      p.L(250, 150);
      p.L(350, 350);
      p.L(500, 350);
      p.stroke();
      p.M(100, 350);
      p.L(250, 350);
      p.L(350, 150);
      p.L(500, 150);
      p.stroke();
    }
  },
  // 0x2388 HELM SYMBOL
  [0x2388]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 160);
      p.stroke();
      p.circle(300, 250, 50);
      p.fill();
      for (let a = 0; a < 360; a += 45) {
        const rad = (a * Math.PI) / 180;
        p.M(300 + 50 * Math.cos(rad), 250 + 50 * Math.sin(rad));
        p.L(300 + 180 * Math.cos(rad), 250 + 180 * Math.sin(rad));
      }
      p.stroke();
    }
  },
  // 0x2389 CIRCLED HORIZONTAL BAR WITH NOTCH
  [0x2389]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 180);
      p.stroke();
      p.M(120, 250);
      p.L(260, 250);
      p.stroke();
      p.M(340, 250);
      p.L(480, 250);
      p.stroke();
      p.M(260, 250);
      p.L(260, 300);
      p.L(340, 300);
      p.L(340, 250);
      p.stroke();
    }
  },
  // 0x238A CIRCLED ANTICLOCKWISE ROTATED DIVISION SIGN
  [0x238a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 180);
      p.stroke();
      p.M(300, 430);
      p.L(300, 70);
      p.stroke();
      p.circle(220, 250, 30);
      p.fill();
      p.circle(380, 250, 30);
      p.fill();
    }
  },
  // 0x238B BROKEN CIRCLE WITH NORTHWEST ARROW (ESCAPE)
  [0x238b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(200, 410);
      p.C(120, 380, 80, 300, 100, 200);
      p.C(120, 120, 200, 80, 300, 70);
      p.C(400, 80, 480, 120, 500, 200);
      p.C(520, 300, 480, 380, 400, 420);
      p.stroke();
      p.M(160, 430);
      p.L(160, 350);
      p.L(240, 350);
      p.stroke();
    }
  },
  // 0x238C UNDO SYMBOL
  [0x238c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(450, 380);
      p.C(450, 200, 350, 120, 250, 120);
      p.C(200, 120, 150, 180, 150, 250);
      p.stroke();
      p.M(150, 250);
      p.L(80, 170);
      p.L(220, 170);
      p.Z();
      p.fill();
    }
  },
  // 0x238D MONOSTABLE SYMBOL
  [0x238d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 150);
      p.L(250, 150);
      p.L(250, 350);
      p.L(400, 350);
      p.L(400, 150);
      p.L(500, 150);
      p.stroke();
    }
  },
  // 0x238E HYSTERESIS SYMBOL
  [0x238e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 150);
      p.L(300, 150);
      p.L(300, 350);
      p.L(500, 350);
      p.stroke();
      p.M(100, 350);
      p.L(300, 350);
      p.stroke();
      p.M(300, 150);
      p.L(500, 150);
      p.stroke();
    }
  },
  // 0x238F OPEN CIRCUIT OUTPUT H-TYPE
  [0x238f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(250, 250);
      p.stroke();
      p.circle(300, 250, 50);
      p.stroke();
      p.M(350, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // 0x2390 OPEN CIRCUIT OUTPUT L-TYPE
  [0x2390]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(250, 250);
      p.stroke();
      p.M(350, 250);
      p.L(500, 250);
      p.stroke();
      p.M(250, 250);
      p.L(350, 250);
      p.stroke();
      p.M(300, 250);
      p.L(300, 350);
      p.stroke();
    }
  },
  // 0x2391 PASSIVE PULL DOWN OUTPUT
  [0x2391]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 350);
      p.L(250, 350);
      p.L(250, 150);
      p.L(350, 150);
      p.L(350, 350);
      p.L(500, 350);
      p.stroke();
    }
  },
  // 0x2392 PASSIVE PULL UP OUTPUT
  [0x2392]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 150);
      p.L(250, 150);
      p.L(250, 350);
      p.L(350, 350);
      p.L(350, 150);
      p.L(500, 150);
      p.stroke();
    }
  },
  // 0x2393 DIRECT CURRENT SYMBOL FORM TWO
  [0x2393]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 300);
      p.L(500, 300);
      p.stroke();
      p.lineWidth(35);
      p.M(100, 200);
      p.L(200, 200);
      p.stroke();
      p.M(250, 200);
      p.L(350, 200);
      p.stroke();
      p.M(400, 200);
      p.L(500, 200);
      p.stroke();
    }
  },
  // 0x2394 SOFTWARE FUNCTION SYMBOL
  [0x2394]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 100);
      p.L(100, 400);
      p.L(500, 400);
      p.L(500, 100);
      p.stroke();
    }
  },
  // 0x2395 APL FUNCTIONAL SYMBOL QUAD
  [0x2395]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.rect(120, 50, 360, 400);
      p.stroke();
    }
  },
  // 0x2396 DECIMAL SEPARATOR KEY SYMBOL
  [0x2396]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(120, 80, 360, 340);
      p.stroke();
      p.circle(300, 180, 35);
      p.fill();
    }
  },
  // 0x2397 PREVIOUS PAGE
  [0x2397]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(130, 50, 340, 400);
      p.stroke();
      p.M(130, 380);
      p.L(220, 380);
      p.L(220, 450);
      p.stroke();
    }
  },
  // 0x2398 NEXT PAGE
  [0x2398]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(130, 50, 340, 400);
      p.stroke();
      p.M(470, 120);
      p.L(380, 120);
      p.L(380, 50);
      p.stroke();
    }
  },
  // 0x2399 PRINT SCREEN SYMBOL
  [0x2399]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 100, 400, 300);
      p.stroke();
      p.rect(180, 180, 240, 140);
      p.stroke();
    }
  },
  // 0x239A CLEAR SCREEN SYMBOL
  [0x239a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 100, 400, 300);
      p.stroke();
      p.M(200, 200);
      p.L(400, 300);
      p.stroke();
      p.M(200, 300);
      p.L(400, 200);
      p.stroke();
    }
  },
  // 0x239E RIGHT PARENTHESIS UPPER HOOK
  [0x239e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 450);
      p.C(350, 450, 400, 350, 400, 250);
      p.stroke();
    }
  },
  // 0x239F RIGHT PARENTHESIS EXTENSION
  [0x239f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 450);
      p.L(400, 50);
      p.stroke();
    }
  },
  // 0x23A0 RIGHT PARENTHESIS LOWER HOOK
  [0x23a0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 250);
      p.C(400, 150, 350, 50, 200, 50);
      p.stroke();
    }
  },
  // 0x23A1 LEFT SQUARE BRACKET UPPER CORNER
  [0x23a1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(350, 450);
      p.L(200, 450);
      p.L(200, 50);
      p.stroke();
    }
  },
  // 0x23A2 LEFT SQUARE BRACKET EXTENSION
  [0x23a2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 450);
      p.L(200, 50);
      p.stroke();
    }
  },
  // 0x23A3 LEFT SQUARE BRACKET LOWER CORNER
  [0x23a3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 450);
      p.L(200, 50);
      p.L(350, 50);
      p.stroke();
    }
  },
  // 0x23A4 RIGHT SQUARE BRACKET UPPER CORNER
  [0x23a4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 450);
      p.L(400, 450);
      p.L(400, 50);
      p.stroke();
    }
  },
  // 0x23A5 RIGHT SQUARE BRACKET EXTENSION
  [0x23a5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 450);
      p.L(400, 50);
      p.stroke();
    }
  },
  // 0x23A6 RIGHT SQUARE BRACKET LOWER CORNER
  [0x23a6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 450);
      p.L(400, 50);
      p.L(250, 50);
      p.stroke();
    }
  },
  // 0x23A7 LEFT CURLY BRACKET UPPER HOOK
  [0x23a7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(350, 450);
      p.C(250, 450, 250, 350, 250, 250);
      p.stroke();
    }
  },
  // 0x23A8 LEFT CURLY BRACKET MIDDLE PIECE
  [0x23a8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 450);
      p.L(250, 300);
      p.C(250, 250, 200, 250, 150, 250);
      p.stroke();
      p.M(250, 200);
      p.C(250, 250, 200, 250, 150, 250);
      p.stroke();
      p.M(250, 200);
      p.L(250, 50);
      p.stroke();
    }
  },
  // 0x23A9 LEFT CURLY BRACKET LOWER HOOK
  [0x23a9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 250);
      p.C(250, 150, 250, 50, 350, 50);
      p.stroke();
    }
  },
  // 0x23AA CURLY BRACKET EXTENSION
  [0x23aa]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
    }
  },
  // 0x23AB RIGHT CURLY BRACKET UPPER HOOK
  [0x23ab]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 450);
      p.C(350, 450, 350, 350, 350, 250);
      p.stroke();
    }
  },
  // 0x23AC RIGHT CURLY BRACKET MIDDLE PIECE
  [0x23ac]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(350, 450);
      p.L(350, 300);
      p.C(350, 250, 400, 250, 450, 250);
      p.stroke();
      p.M(350, 200);
      p.C(350, 250, 400, 250, 450, 250);
      p.stroke();
      p.M(350, 200);
      p.L(350, 50);
      p.stroke();
    }
  },
  // 0x23AD RIGHT CURLY BRACKET LOWER HOOK
  [0x23ad]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(350, 250);
      p.C(350, 150, 350, 50, 250, 50);
      p.stroke();
    }
  },
  // 0x23AE INTEGRAL EXTENSION
  [0x23ae]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
    }
  },
  // 0x23B0 UPPER LEFT OR LOWER RIGHT CURLY BRACKET SECTION
  [0x23b0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(350, 50);
      p.C(250, 50, 250, 250, 250, 450);
      p.stroke();
    }
  },
  // 0x23B1 UPPER RIGHT OR LOWER LEFT CURLY BRACKET SECTION
  [0x23b1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 50);
      p.C(350, 50, 350, 250, 350, 450);
      p.stroke();
    }
  },
  // 0x23B2 SUMMATION TOP
  [0x23b2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(480, 450);
      p.L(120, 450);
      p.L(300, 250);
      p.stroke();
    }
  },
  // 0x23B3 SUMMATION BOTTOM
  [0x23b3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 250);
      p.L(120, 50);
      p.L(480, 50);
      p.stroke();
    }
  },
  // 0x23B4 TOP SQUARE BRACKET
  [0x23b4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 200);
      p.L(100, 350);
      p.L(500, 350);
      p.L(500, 200);
      p.stroke();
    }
  },
  // 0x23B5 BOTTOM SQUARE BRACKET
  [0x23b5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 300);
      p.L(100, 150);
      p.L(500, 150);
      p.L(500, 300);
      p.stroke();
    }
  },
  // 0x23B6 BOTTOM SQUARE BRACKET OVER TOP SQUARE BRACKET
  [0x23b6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 300);
      p.L(100, 200);
      p.L(500, 200);
      p.L(500, 300);
      p.stroke();
      p.M(100, 350);
      p.L(100, 250);
      p.L(500, 250);
      p.L(500, 350);
      p.stroke();
    }
  },
  // 0x23BA HORIZONTAL SCAN LINE-1
  [0x23ba]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(0, 450);
      p.L(600, 450);
      p.stroke();
    }
  },
  // 0x23BB HORIZONTAL SCAN LINE-3
  [0x23bb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(0, 350);
      p.L(600, 350);
      p.stroke();
    }
  },
  // 0x23BC HORIZONTAL SCAN LINE-7
  [0x23bc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(0, 150);
      p.L(600, 150);
      p.stroke();
    }
  },
  // 0x23BD HORIZONTAL SCAN LINE-9
  [0x23bd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(0, 50);
      p.L(600, 50);
      p.stroke();
    }
  },
  // 0x23BE DENTISTRY SYMBOL LIGHT VERTICAL AND TOP RIGHT
  [0x23be]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(300, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // 0x23BF DENTISTRY SYMBOL LIGHT VERTICAL AND BOTTOM RIGHT
  [0x23bf]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // 0x23C0 DENTISTRY SYMBOL LIGHT VERTICAL WITH CIRCLE
  [0x23c0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
      p.circle(300, 250, 60);
      p.stroke();
    }
  },
  // 0x23C1 DENTISTRY SYMBOL LIGHT DOWN AND HORIZONTAL WITH CIRCLE
  [0x23c1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 350);
      p.L(500, 350);
      p.stroke();
      p.M(300, 350);
      p.L(300, 100);
      p.stroke();
      p.circle(300, 250, 60);
      p.stroke();
    }
  },
  // 0x23C9 DENTISTRY SYMBOL LIGHT DOWN AND HORIZONTAL
  [0x23c9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 350);
      p.L(450, 350);
      p.stroke();
      p.M(300, 350);
      p.L(300, 100);
      p.stroke();
    }
  },
  // 0x23CA DENTISTRY SYMBOL LIGHT UP AND HORIZONTAL
  [0x23ca]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 150);
      p.L(450, 150);
      p.stroke();
      p.M(300, 150);
      p.L(300, 400);
      p.stroke();
    }
  },
  // 0x23CB DENTISTRY SYMBOL LIGHT VERTICAL WITH TRIANGLE
  [0x23cb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
      p.M(300, 310);
      p.L(400, 200);
      p.L(300, 200);
      p.Z();
      p.stroke();
    }
  },
  // 0x23CC DENTISTRY SYMBOL LIGHT DOWN AND HORIZONTAL WITH TRIANGLE
  [0x23cc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 350);
      p.L(500, 350);
      p.stroke();
      p.M(300, 350);
      p.L(300, 120);
      p.stroke();
      p.M(250, 250);
      p.L(350, 250);
      p.L(300, 180);
      p.Z();
      p.fill();
    }
  },
  // 0x23CD SQUARE FOOT
  [0x23cd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 400);
      p.L(150, 100);
      p.L(450, 100);
      p.stroke();
    }
  },
  // 0x23D0 VERTICAL LINE EXTENSION
  [0x23d0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 480);
      p.L(300, 20);
      p.stroke();
    }
  },
  // 0x23DA EARTH GROUND
  [0x23da]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(300, 200);
      p.stroke();
      p.M(150, 200);
      p.L(450, 200);
      p.stroke();
      p.M(200, 140);
      p.L(400, 140);
      p.stroke();
      p.M(250, 80);
      p.L(350, 80);
      p.stroke();
    }
  },
  // 0x23DB FUSE
  [0x23db]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 250);
      p.L(180, 250);
      p.stroke();
      p.rect(180, 180, 240, 140);
      p.stroke();
      p.M(420, 250);
      p.L(520, 250);
      p.stroke();
    }
  },
  // 0x23DC TOP PARENTHESIS
  [0x23dc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 150);
      p.C(80, 400, 520, 400, 520, 150);
      p.stroke();
    }
  },
  // 0x23DD BOTTOM PARENTHESIS
  [0x23dd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 350);
      p.C(80, 100, 520, 100, 520, 350);
      p.stroke();
    }
  },
  // 0x23DE TOP CURLY BRACKET
  [0x23de]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 150);
      p.C(80, 300, 200, 350, 300, 400);
      p.C(400, 350, 520, 300, 520, 150);
      p.stroke();
    }
  },
  // 0x23DF BOTTOM CURLY BRACKET
  [0x23df]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 350);
      p.C(80, 200, 200, 150, 300, 100);
      p.C(400, 150, 520, 200, 520, 350);
      p.stroke();
    }
  },
  // 0x23E0 TOP TORTOISE SHELL BRACKET
  [0x23e0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 150);
      p.C(150, 380, 450, 380, 520, 150);
      p.stroke();
    }
  },
  // 0x23E1 BOTTOM TORTOISE SHELL BRACKET
  [0x23e1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 350);
      p.C(150, 120, 450, 120, 520, 350);
      p.stroke();
    }
  },
  // 0x23E2 WHITE TRAPEZIUM
  [0x23e2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 100);
      p.L(450, 100);
      p.L(500, 400);
      p.L(100, 400);
      p.Z();
      p.stroke();
    }
  },
  // 0x23E3 BENZENE RING WITH CIRCLE
  [0x23e3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const cx = 300,
        cy = 250,
        r = 180;
      for (let i = 0; i < 6; i++) {
        const a1 = ((i * 60 - 90) * Math.PI) / 180;
        const a2 = (((i + 1) * 60 - 90) * Math.PI) / 180;
        p.M(cx + r * Math.cos(a1), cy + r * Math.sin(a1));
        p.L(cx + r * Math.cos(a2), cy + r * Math.sin(a2));
      }
      p.stroke();
      p.circle(cx, cy, 110);
      p.stroke();
    }
  },
  // 0x23E4 STRAIGHTNESS
  [0x23e4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // 0x23E5 FLATNESS
  [0x23e5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 200);
      p.L(500, 200);
      p.stroke();
      p.M(100, 300);
      p.L(500, 300);
      p.stroke();
    }
  },
  // 0x23E6 AC CURRENT
  [0x23e6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.C(200, 420, 300, 80, 400, 250);
      p.C(440, 340, 480, 250, 500, 250);
      p.stroke();
    }
  },
  // 0x23E7 ELECTRICAL INTERSECTION
  [0x23e7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 100);
      p.L(100, 300);
      p.C(100, 430, 500, 430, 500, 300);
      p.L(500, 100);
      p.stroke();
    }
  },
  // 0x23E8 DECIMAL EXPONENT SYMBOL
  [0x23e8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      // E shape
      p.M(400, 400);
      p.L(200, 400);
      p.L(200, 100);
      p.L(400, 100);
      p.stroke();
      p.M(200, 250);
      p.L(370, 250);
      p.stroke();
    }
  },
  // ⏰ ALARM CLOCK
  [0x23f0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 230, 180);
      p.stroke();
      // clock hands
      p.M(300, 230);
      p.L(300, 340);
      p.stroke();
      p.M(300, 230);
      p.L(380, 230);
      p.stroke();
      // bells
      p.M(160, 400);
      p.L(120, 440);
      p.stroke();
      p.M(440, 400);
      p.L(480, 440);
      p.stroke();
      // feet
      p.M(200, 50);
      p.L(180, 20);
      p.stroke();
      p.M(400, 50);
      p.L(420, 20);
      p.stroke();
    }
  },
  // ⏳ HOURGLASS WITH FLOWING SAND
  [0x23f3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
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
      // sand dots in lower half
      p.circle(270, 120, 15);
      p.fill();
      p.circle(330, 120, 15);
      p.fill();
      p.circle(300, 90, 15);
      p.fill();
      // flowing line
      p.lineWidth(20);
      p.M(300, 250);
      p.L(300, 160);
      p.stroke();
    }
  },
  // 0x23F4 BLACK MEDIUM LEFT-POINTING TRIANGLE
  [0x23f4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(450, 450);
      p.L(150, 250);
      p.L(450, 50);
      p.Z();
      p.fill();
    }
  },
  // 0x23F5 BLACK MEDIUM RIGHT-POINTING TRIANGLE
  [0x23f5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(150, 450);
      p.L(450, 250);
      p.L(150, 50);
      p.Z();
      p.fill();
    }
  },
  // 0x23F6 BLACK MEDIUM UP-POINTING TRIANGLE
  [0x23f6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(100, 100);
      p.L(500, 100);
      p.L(300, 420);
      p.Z();
      p.fill();
    }
  },
  // 0x23F7 BLACK MEDIUM DOWN-POINTING TRIANGLE
  [0x23f7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(100, 400);
      p.L(500, 400);
      p.L(300, 80);
      p.Z();
      p.fill();
    }
  },
  // 0x23FB POWER SYMBOL
  [0x23fb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 380);
      p.C(120, 330, 80, 230, 120, 150);
      p.C(160, 70, 260, 40, 300, 40);
      p.C(340, 40, 440, 70, 480, 150);
      p.C(520, 230, 480, 330, 400, 380);
      p.stroke();
      p.M(300, 320);
      p.L(300, 150);
      p.stroke();
    }
  },
  // 0x23FC POWER ON-OFF SYMBOL
  [0x23fc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.circle(300, 250, 200);
      p.stroke();
      p.M(300, 350);
      p.L(300, 150);
      p.stroke();
    }
  },
  // 0x23FD POWER ON SYMBOL
  [0x23fd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 420);
      p.L(300, 80);
      p.stroke();
    }
  },
  // 0x23FE POWER SLEEP SYMBOL
  [0x23fe]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(180, 400);
      p.C(100, 300, 100, 150, 200, 80);
      p.C(300, 10, 430, 50, 480, 150);
      p.C(530, 250, 470, 380, 380, 420);
      p.stroke();
    }
  },
  // 0x23FF OBSERVER EYE SYMBOL
  [0x23ff]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(250, 100);
      p.L(500, 250);
      p.stroke();
      p.circle(250, 200, 40);
      p.fill();
    }
  }
};

// =============================================================================
// Supplemental Arrows-B (U+2900–U+297F) — 128 chars
// =============================================================================

export const SUP_ARROWS_B: Record<number, GlyphDef> = {
  // ⤀ RIGHTWARDS TWO-HEADED ARROW WITH VERTICAL STROKE
  [0x2900]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(300, 150);
      p.L(300, 350);
      p.stroke();
    }
  },
  // ⤁ RIGHTWARDS TWO-HEADED ARROW WITH DOUBLE VERTICAL STROKE
  [0x2901]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(80, 250);
      p.L(520, 250);
      p.stroke();
      p.M(520, 250);
      p.L(420, 350);
      p.L(420, 150);
      p.Z();
      p.fill();
      p.M(80, 250);
      p.L(180, 350);
      p.L(180, 150);
      p.Z();
      p.fill();
      p.M(280, 150);
      p.L(280, 350);
      p.stroke();
      p.M(320, 150);
      p.L(320, 350);
      p.stroke();
    }
  },
  // ⤂ LEFTWARDS DOUBLE ARROW WITH VERTICAL STROKE
  [0x2902]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(100, 250);
      p.L(240, 380);
      p.L(240, 300);
      p.L(500, 300);
      p.L(500, 200);
      p.L(240, 200);
      p.L(240, 120);
      p.Z();
      p.fill();
      p.lineWidth(40);
      p.M(340, 150);
      p.L(340, 350);
      p.stroke();
    }
  },
  // ⤃ RIGHTWARDS DOUBLE ARROW WITH VERTICAL STROKE
  [0x2903]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 250);
      p.L(360, 380);
      p.L(360, 300);
      p.L(100, 300);
      p.L(100, 200);
      p.L(360, 200);
      p.L(360, 120);
      p.Z();
      p.fill();
      p.lineWidth(40);
      p.M(260, 150);
      p.L(260, 350);
      p.stroke();
    }
  },
  // ⤄ LEFT RIGHT DOUBLE ARROW WITH VERTICAL STROKE
  [0x2904]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(80, 250);
      p.L(200, 380);
      p.L(200, 300);
      p.L(400, 300);
      p.L(400, 380);
      p.L(520, 250);
      p.L(400, 120);
      p.L(400, 200);
      p.L(200, 200);
      p.L(200, 120);
      p.Z();
      p.fill();
      p.lineWidth(40);
      p.M(300, 150);
      p.L(300, 350);
      p.stroke();
    }
  },
  // ⤅ RIGHTWARDS TWO-HEADED ARROW FROM BAR
  [0x2905]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 150);
      p.L(100, 350);
      p.stroke();
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
      p.M(200, 250);
      p.L(250, 330);
      p.L(250, 170);
      p.Z();
      p.fill();
    }
  },
  // ⤆ LEFTWARDS DOUBLE ARROW FROM BAR
  [0x2906]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 150);
      p.L(500, 350);
      p.stroke();
      p.M(100, 250);
      p.L(280, 380);
      p.L(280, 300);
      p.L(500, 300);
      p.L(500, 200);
      p.L(280, 200);
      p.L(280, 120);
      p.Z();
      p.fill();
    }
  },
  // ⤇ RIGHTWARDS DOUBLE ARROW FROM BAR
  [0x2907]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 150);
      p.L(100, 350);
      p.stroke();
      p.M(500, 250);
      p.L(320, 380);
      p.L(320, 300);
      p.L(100, 300);
      p.L(100, 200);
      p.L(320, 200);
      p.L(320, 120);
      p.Z();
      p.fill();
    }
  },
  // ⤈ RIGHTWARDS ARROW WITH STROKE
  [0x2908]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(300, 100);
      p.stroke();
      p.M(300, 100);
      p.L(200, 200);
      p.L(400, 200);
      p.Z();
      p.fill();
      p.M(200, 350);
      p.L(400, 250);
      p.stroke();
    }
  },
  // ⤉ UPWARDS ARROW WITH STROKE
  [0x2909]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 400);
      p.stroke();
      p.M(300, 400);
      p.L(200, 300);
      p.L(400, 300);
      p.Z();
      p.fill();
      p.M(200, 250);
      p.L(400, 150);
      p.stroke();
    }
  },
  // ⤊ UPWARDS OF DOWNWARDS ARROW
  [0x290a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(200, 340);
      p.L(400, 340);
      p.Z();
      p.fill();
      p.M(300, 50);
      p.L(200, 100);
      p.L(400, 100);
      p.stroke();
    }
  },
  // ⤋ DOWNWARDS OF UPWARDS ARROW
  [0x290b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 450);
      p.L(300, 50);
      p.stroke();
      p.M(300, 50);
      p.L(200, 160);
      p.L(400, 160);
      p.Z();
      p.fill();
      p.M(300, 450);
      p.L(200, 400);
      p.L(400, 400);
      p.stroke();
    }
  },
  // ⤌ LEFTWARDS DOUBLE DASH ARROW
  [0x290c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(400, 250);
      p.stroke();
      p.M(350, 250);
      p.L(250, 250);
      p.stroke();
      p.M(100, 250);
      p.L(230, 370);
      p.L(230, 130);
      p.Z();
      p.fill();
    }
  },
  // ⤍ RIGHTWARDS DOUBLE DASH ARROW
  [0x290d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(200, 250);
      p.stroke();
      p.M(250, 250);
      p.L(350, 250);
      p.stroke();
      p.M(500, 250);
      p.L(370, 370);
      p.L(370, 130);
      p.Z();
      p.fill();
    }
  },
  // ⤎ LEFTWARDS TRIPLE DASH ARROW
  [0x290e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 250);
      p.L(430, 250);
      p.stroke();
      p.M(380, 250);
      p.L(310, 250);
      p.stroke();
      p.M(260, 250);
      p.L(190, 250);
      p.stroke();
      p.M(100, 250);
      p.L(220, 360);
      p.L(220, 140);
      p.Z();
      p.fill();
    }
  },
  // ⤏ RIGHTWARDS TRIPLE DASH ARROW
  [0x290f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 250);
      p.L(170, 250);
      p.stroke();
      p.M(220, 250);
      p.L(290, 250);
      p.stroke();
      p.M(340, 250);
      p.L(410, 250);
      p.stroke();
      p.M(500, 250);
      p.L(380, 360);
      p.L(380, 140);
      p.Z();
      p.fill();
    }
  },
  // ⤐ RIGHTWARDS TWO-HEADED TRIPLE DASH ARROW
  [0x2910]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(80, 250);
      p.L(150, 250);
      p.stroke();
      p.M(200, 250);
      p.L(270, 250);
      p.stroke();
      p.M(320, 250);
      p.L(390, 250);
      p.stroke();
      p.M(520, 250);
      p.L(400, 360);
      p.L(400, 140);
      p.Z();
      p.fill();
      p.M(80, 250);
      p.L(160, 330);
      p.L(160, 170);
      p.Z();
      p.fill();
    }
  },
  // ⤑ RIGHTWARDS ARROW WITH DOTTED STEM
  [0x2911]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(150, 250, 20);
      p.fill();
      p.circle(230, 250, 20);
      p.fill();
      p.circle(310, 250, 20);
      p.fill();
      p.lineWidth(40);
      p.M(370, 250);
      p.L(430, 250);
      p.stroke();
      p.M(500, 250);
      p.L(380, 350);
      p.L(380, 150);
      p.Z();
      p.fill();
    }
  },
  // ⤒ UPWARDS ARROW TO BAR
  [0x2912]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 50);
      p.L(300, 380);
      p.stroke();
      p.M(300, 380);
      p.L(200, 280);
      p.L(400, 280);
      p.Z();
      p.fill();
      p.M(150, 430);
      p.L(450, 430);
      p.stroke();
    }
  },
  // ⤓ DOWNWARDS ARROW TO BAR
  [0x2913]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 420);
      p.L(300, 100);
      p.stroke();
      p.M(300, 100);
      p.L(200, 200);
      p.L(400, 200);
      p.Z();
      p.fill();
      p.M(150, 50);
      p.L(450, 50);
      p.stroke();
    }
  },
  // ⤖ RIGHTWARDS TWO-HEADED ARROW WITH TAIL
  [0x2916]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
      p.M(100, 250);
      p.L(180, 340);
      p.L(180, 160);
      p.Z();
      p.fill();
      // tail
      p.M(100, 350);
      p.L(100, 150);
      p.stroke();
    }
  },
  // ⤝ LEFTWARDS ARROW-TAIL
  [0x291d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(200, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(500, 350);
      p.L(500, 150);
      p.stroke();
    }
  },
  // ⤞ RIGHTWARDS ARROW-TAIL
  [0x291e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(400, 250);
      p.stroke();
      p.M(500, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
      p.M(100, 350);
      p.L(100, 150);
      p.stroke();
    }
  },
  // ⤟ LEFTWARDS ARROW FROM BAR TO BLACK DIAMOND
  [0x291f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 150);
      p.L(500, 350);
      p.stroke();
      p.M(500, 250);
      p.L(200, 250);
      p.stroke();
      // diamond at end
      p.M(100, 250);
      p.L(150, 300);
      p.L(200, 250);
      p.L(150, 200);
      p.Z();
      p.fill();
    }
  },
  // ⤠ RIGHTWARDS ARROW FROM BAR TO BLACK DIAMOND
  [0x2920]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 150);
      p.L(100, 350);
      p.stroke();
      p.M(100, 250);
      p.L(400, 250);
      p.stroke();
      // diamond at end
      p.M(500, 250);
      p.L(450, 300);
      p.L(400, 250);
      p.L(450, 200);
      p.Z();
      p.fill();
    }
  },
  // ⤡ NORTH WEST AND SOUTH EAST ARROW
  [0x2921]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 420);
      p.L(480, 80);
      p.stroke();
      p.M(120, 420);
      p.L(120, 320);
      p.L(220, 420);
      p.Z();
      p.fill();
      p.M(480, 80);
      p.L(480, 180);
      p.L(380, 80);
      p.Z();
      p.fill();
    }
  },
  // ⤢ NORTH EAST AND SOUTH WEST ARROW
  [0x2922]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(480, 420);
      p.L(120, 80);
      p.stroke();
      p.M(480, 420);
      p.L(380, 420);
      p.L(480, 320);
      p.Z();
      p.fill();
      p.M(120, 80);
      p.L(220, 80);
      p.L(120, 180);
      p.Z();
      p.fill();
    }
  },
  // ⤣ NORTH WEST ARROW WITH HOOK
  [0x2923]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 420);
      p.L(450, 120);
      p.C(480, 80, 500, 120, 480, 180);
      p.stroke();
      p.M(150, 420);
      p.L(150, 310);
      p.L(260, 420);
      p.Z();
      p.fill();
    }
  },
  // ⤤ NORTH EAST ARROW WITH HOOK
  [0x2924]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(450, 420);
      p.L(150, 120);
      p.C(120, 80, 100, 120, 120, 180);
      p.stroke();
      p.M(450, 420);
      p.L(450, 310);
      p.L(340, 420);
      p.Z();
      p.fill();
    }
  },
  // ⤥ SOUTH EAST ARROW WITH HOOK
  [0x2925]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(450, 80);
      p.L(150, 380);
      p.C(120, 420, 100, 380, 120, 320);
      p.stroke();
      p.M(450, 80);
      p.L(450, 190);
      p.L(340, 80);
      p.Z();
      p.fill();
    }
  },
  // ⤦ SOUTH WEST ARROW WITH HOOK
  [0x2926]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 80);
      p.L(450, 380);
      p.C(480, 420, 500, 380, 480, 320);
      p.stroke();
      p.M(150, 80);
      p.L(150, 190);
      p.L(260, 80);
      p.Z();
      p.fill();
    }
  },
  // ⤧ NORTH WEST ARROW AND NORTH EAST ARROW
  [0x2927]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 100);
      p.L(120, 400);
      p.stroke();
      p.M(120, 400);
      p.L(120, 300);
      p.L(220, 400);
      p.Z();
      p.fill();
      p.M(300, 100);
      p.L(480, 400);
      p.stroke();
      p.M(480, 400);
      p.L(480, 300);
      p.L(380, 400);
      p.Z();
      p.fill();
    }
  },
  // ⤨ NORTH EAST ARROW AND SOUTH EAST ARROW
  [0x2928]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 250);
      p.L(480, 420);
      p.stroke();
      p.M(480, 420);
      p.L(380, 420);
      p.L(480, 330);
      p.Z();
      p.fill();
      p.M(100, 250);
      p.L(480, 80);
      p.stroke();
      p.M(480, 80);
      p.L(380, 80);
      p.L(480, 170);
      p.Z();
      p.fill();
    }
  },
  // ⤩ SOUTH EAST ARROW AND SOUTH WEST ARROW
  [0x2929]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 400);
      p.L(120, 100);
      p.stroke();
      p.M(120, 100);
      p.L(220, 100);
      p.L(120, 200);
      p.Z();
      p.fill();
      p.M(300, 400);
      p.L(480, 100);
      p.stroke();
      p.M(480, 100);
      p.L(380, 100);
      p.L(480, 200);
      p.Z();
      p.fill();
    }
  },
  // ⤪ SOUTH WEST ARROW AND NORTH WEST ARROW
  [0x292a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 250);
      p.L(120, 80);
      p.stroke();
      p.M(120, 80);
      p.L(120, 180);
      p.L(220, 80);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(120, 420);
      p.stroke();
      p.M(120, 420);
      p.L(120, 320);
      p.L(220, 420);
      p.Z();
      p.fill();
    }
  },
  // ⤴ RIGHT ARROW CURVING UP
  [0x2934]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 150);
      p.L(350, 150);
      p.C(430, 150, 450, 250, 450, 350);
      p.stroke();
      p.M(450, 350);
      p.L(380, 230);
      p.L(520, 230);
      p.Z();
      p.fill();
    }
  },
  // ⤵ RIGHT ARROW CURVING DOWN
  [0x2935]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 350);
      p.L(350, 350);
      p.C(430, 350, 450, 250, 450, 150);
      p.stroke();
      p.M(450, 150);
      p.L(380, 270);
      p.L(520, 270);
      p.Z();
      p.fill();
    }
  },
  // ⤶ LEFT ARROW CURVING DOWN
  [0x2936]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(500, 400);
      p.L(300, 400);
      p.C(200, 400, 150, 300, 150, 200);
      p.stroke();
      p.M(150, 200);
      p.L(80, 300);
      p.L(220, 300);
      p.Z();
      p.fill();
    }
  },
  // ⤷ RIGHT ARROW CURVING DOWN
  [0x2937]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 400);
      p.L(300, 400);
      p.C(400, 400, 450, 300, 450, 200);
      p.stroke();
      p.M(450, 200);
      p.L(380, 300);
      p.L(520, 300);
      p.Z();
      p.fill();
    }
  },
  // ⤸ RIGHT-SIDE ARC CLOCKWISE ARROW
  [0x2938]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(400, 400);
      p.C(500, 350, 500, 150, 400, 100);
      p.C(300, 50, 200, 100, 150, 200);
      p.stroke();
      p.M(400, 400);
      p.L(330, 340);
      p.L(450, 340);
      p.Z();
      p.fill();
    }
  },
  // ⤹ LEFT-SIDE ARC ANTICLOCKWISE ARROW
  [0x2939]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(200, 400);
      p.C(100, 350, 100, 150, 200, 100);
      p.C(300, 50, 400, 100, 450, 200);
      p.stroke();
      p.M(200, 400);
      p.L(270, 340);
      p.L(150, 340);
      p.Z();
      p.fill();
    }
  },
  // ⤺ TOP ARC ANTICLOCKWISE ARROW
  [0x293a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 200);
      p.C(100, 350, 250, 430, 400, 380);
      p.C(480, 340, 500, 250, 450, 180);
      p.stroke();
      p.M(100, 200);
      p.L(170, 270);
      p.L(170, 130);
      p.Z();
      p.fill();
    }
  },
  // ⤻ BOTTOM ARC CLOCKWISE ARROW
  [0x293b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 300);
      p.C(500, 150, 350, 70, 200, 120);
      p.C(120, 160, 100, 250, 150, 320);
      p.stroke();
      p.M(500, 300);
      p.L(430, 230);
      p.L(430, 370);
      p.Z();
      p.fill();
    }
  },
  // ⥀ ANTICLOCKWISE CLOSED CIRCLE ARROW
  [0x2940]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 430);
      p.C(180, 430, 100, 350, 100, 250);
      p.C(100, 150, 180, 70, 300, 70);
      p.C(420, 70, 500, 150, 500, 250);
      p.C(500, 350, 420, 430, 340, 430);
      p.stroke();
      p.M(300, 430);
      p.L(240, 370);
      p.L(240, 480);
      p.Z();
      p.fill();
    }
  },
  // ⥁ CLOCKWISE CLOSED CIRCLE ARROW
  [0x2941]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 430);
      p.C(420, 430, 500, 350, 500, 250);
      p.C(500, 150, 420, 70, 300, 70);
      p.C(180, 70, 100, 150, 100, 250);
      p.C(100, 350, 180, 430, 260, 430);
      p.stroke();
      p.M(300, 430);
      p.L(360, 370);
      p.L(360, 480);
      p.Z();
      p.fill();
    }
  },
  // ⥂ RIGHTWARDS ARROW ABOVE SHORT LEFTWARDS ARROW
  [0x2942]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(100, 320);
      p.L(460, 320);
      p.stroke();
      p.M(500, 320);
      p.L(400, 380);
      p.L(400, 260);
      p.Z();
      p.fill();
      p.M(400, 180);
      p.L(180, 180);
      p.stroke();
      p.M(140, 180);
      p.L(240, 240);
      p.L(240, 120);
      p.Z();
      p.fill();
    }
  },
  // ⥃ LEFTWARDS ARROW ABOVE SHORT RIGHTWARDS ARROW
  [0x2943]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(500, 320);
      p.L(140, 320);
      p.stroke();
      p.M(100, 320);
      p.L(200, 380);
      p.L(200, 260);
      p.Z();
      p.fill();
      p.M(200, 180);
      p.L(420, 180);
      p.stroke();
      p.M(460, 180);
      p.L(360, 240);
      p.L(360, 120);
      p.Z();
      p.fill();
    }
  },
  // ⥄ SHORT RIGHTWARDS ARROW ABOVE LEFTWARDS ARROW
  [0x2944]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(200, 320);
      p.L(420, 320);
      p.stroke();
      p.M(460, 320);
      p.L(360, 380);
      p.L(360, 260);
      p.Z();
      p.fill();
      p.M(500, 180);
      p.L(140, 180);
      p.stroke();
      p.M(100, 180);
      p.L(200, 240);
      p.L(200, 120);
      p.Z();
      p.fill();
    }
  },
  // ⥅ RIGHTWARDS ARROW WITH PLUS BELOW
  [0x2945]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 300);
      p.L(460, 300);
      p.stroke();
      p.M(500, 300);
      p.L(400, 380);
      p.L(400, 220);
      p.Z();
      p.fill();
      // plus below
      p.lineWidth(30);
      p.M(270, 100);
      p.L(330, 100);
      p.stroke();
      p.M(300, 70);
      p.L(300, 130);
      p.stroke();
    }
  },
  // ⥈ RIGHTWARDS ARROW ABOVE ALMOST EQUAL TO
  [0x2948]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 350);
      p.L(460, 350);
      p.stroke();
      p.M(500, 350);
      p.L(400, 410);
      p.L(400, 290);
      p.Z();
      p.fill();
      // tilde lines below
      p.lineWidth(30);
      p.M(140, 200);
      p.C(220, 240, 360, 160, 460, 200);
      p.stroke();
      p.M(140, 130);
      p.C(220, 170, 360, 90, 460, 130);
      p.stroke();
    }
  },
  // ⥊ LEFT BARB UP RIGHT BARB DOWN HARPOON
  [0x294a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 370);
      p.L(200, 250);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(400, 130);
      p.L(400, 250);
      p.Z();
      p.fill();
    }
  },
  // ⥋ LEFT BARB DOWN RIGHT BARB UP HARPOON
  [0x294b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 130);
      p.L(200, 250);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(400, 370);
      p.L(400, 250);
      p.Z();
      p.fill();
    }
  },
  // ⥌ UP BARB RIGHT DOWN BARB LEFT HARPOON
  [0x294c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(400, 350);
      p.L(300, 350);
      p.Z();
      p.fill();
      p.M(300, 50);
      p.L(200, 150);
      p.L(300, 150);
      p.Z();
      p.fill();
    }
  },
  // ⥍ UP BARB LEFT DOWN BARB RIGHT HARPOON
  [0x294d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(200, 350);
      p.L(300, 350);
      p.Z();
      p.fill();
      p.M(300, 50);
      p.L(400, 150);
      p.L(300, 150);
      p.Z();
      p.fill();
    }
  },
  // ⥎ LEFT BARB UP RIGHT BARB UP HARPOON
  [0x294e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 370);
      p.L(200, 250);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(400, 370);
      p.L(400, 250);
      p.Z();
      p.fill();
    }
  },
  // ⥏ UP BARB RIGHT DOWN BARB RIGHT HARPOON
  [0x294f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(400, 350);
      p.L(300, 350);
      p.Z();
      p.fill();
      p.M(300, 50);
      p.L(400, 150);
      p.L(300, 150);
      p.Z();
      p.fill();
    }
  },
  // ⥐ LEFT BARB DOWN RIGHT BARB DOWN HARPOON
  [0x2950]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 130);
      p.L(200, 250);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(400, 130);
      p.L(400, 250);
      p.Z();
      p.fill();
    }
  },
  // ⥑ UP BARB LEFT DOWN BARB LEFT HARPOON
  [0x2951]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(200, 350);
      p.L(300, 350);
      p.Z();
      p.fill();
      p.M(300, 50);
      p.L(200, 150);
      p.L(300, 150);
      p.Z();
      p.fill();
    }
  },
  // ⥢ LEFTWARDS HARPOON WITH BARB UP FROM BAR
  [0x2962]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(150, 250);
      p.stroke();
      p.M(150, 250);
      p.L(260, 370);
      p.L(260, 250);
      p.Z();
      p.fill();
      p.M(500, 150);
      p.L(500, 350);
      p.stroke();
    }
  },
  // ⥤ RIGHTWARDS HARPOON WITH BARB UP FROM BAR
  [0x2964]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(450, 250);
      p.stroke();
      p.M(450, 250);
      p.L(340, 370);
      p.L(340, 250);
      p.Z();
      p.fill();
      p.M(100, 150);
      p.L(100, 350);
      p.stroke();
    }
  },
  // ⥪ LEFT BARB UP RIGHT BARB UP HARPOON (short)
  [0x296a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 300);
      p.L(500, 300);
      p.stroke();
      p.M(100, 300);
      p.L(200, 400);
      p.L(200, 300);
      p.Z();
      p.fill();
      p.M(100, 200);
      p.L(500, 200);
      p.stroke();
      p.M(500, 200);
      p.L(400, 300);
      p.L(400, 200);
      p.Z();
      p.fill();
    }
  },
  // ⥫ LEFT BARB DOWN RIGHT BARB DOWN HARPOON (short)
  [0x296b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 200);
      p.L(500, 200);
      p.stroke();
      p.M(100, 200);
      p.L(200, 100);
      p.L(200, 200);
      p.Z();
      p.fill();
      p.M(100, 300);
      p.L(500, 300);
      p.stroke();
      p.M(500, 300);
      p.L(400, 200);
      p.L(400, 300);
      p.Z();
      p.fill();
    }
  },
  // ⥬ LEFT BARB UP RIGHT BARB DOWN HARPOON LONG
  [0x296c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 200);
      p.L(500, 200);
      p.stroke();
      p.M(100, 200);
      p.L(200, 300);
      p.L(200, 200);
      p.Z();
      p.fill();
      p.M(100, 300);
      p.L(500, 300);
      p.stroke();
      p.M(500, 300);
      p.L(400, 200);
      p.L(400, 300);
      p.Z();
      p.fill();
    }
  },
  // ⥭ LEFT BARB DOWN RIGHT BARB UP HARPOON LONG
  [0x296d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 200);
      p.L(500, 200);
      p.stroke();
      p.M(100, 200);
      p.L(200, 100);
      p.L(200, 200);
      p.Z();
      p.fill();
      p.M(100, 300);
      p.L(500, 300);
      p.stroke();
      p.M(500, 300);
      p.L(400, 400);
      p.L(400, 300);
      p.Z();
      p.fill();
    }
  },
  // ⥮ UPWARDS HARPOON WITH BARB LEFT BESIDE DOWNWARDS HARPOON WITH BARB RIGHT
  [0x296e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      // left: upwards, barb left
      p.M(230, 80);
      p.L(230, 420);
      p.stroke();
      p.M(230, 420);
      p.L(150, 320);
      p.L(230, 320);
      p.Z();
      p.fill();
      // right: downwards, barb right
      p.M(370, 420);
      p.L(370, 80);
      p.stroke();
      p.M(370, 80);
      p.L(450, 180);
      p.L(370, 180);
      p.Z();
      p.fill();
    }
  },
  // ⥯ DOWNWARDS HARPOON WITH BARB LEFT BESIDE UPWARDS HARPOON WITH BARB RIGHT
  [0x296f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(230, 420);
      p.L(230, 80);
      p.stroke();
      p.M(230, 80);
      p.L(150, 180);
      p.L(230, 180);
      p.Z();
      p.fill();
      p.M(370, 80);
      p.L(370, 420);
      p.stroke();
      p.M(370, 420);
      p.L(450, 320);
      p.L(370, 320);
      p.Z();
      p.fill();
    }
  },
  // ⥰ RIGHT DOUBLE ARROW WITH ROUNDED HEAD
  [0x2970]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 300);
      p.L(420, 300);
      p.stroke();
      p.M(100, 200);
      p.L(420, 200);
      p.stroke();
      p.M(500, 250);
      p.C(480, 300, 440, 350, 420, 350);
      p.stroke();
      p.M(500, 250);
      p.C(480, 200, 440, 150, 420, 150);
      p.stroke();
    }
  },
  // ⥴ RIGHTWARDS ARROW ABOVE REVERSE ALMOST EQUAL TO
  [0x2974]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 350);
      p.L(460, 350);
      p.stroke();
      p.M(500, 350);
      p.L(400, 410);
      p.L(400, 290);
      p.Z();
      p.fill();
      p.lineWidth(30);
      p.M(460, 200);
      p.C(380, 240, 220, 160, 140, 200);
      p.stroke();
      p.M(460, 130);
      p.C(380, 170, 220, 90, 140, 130);
      p.stroke();
    }
  },
  // ⥵ RIGHTWARDS ARROW ABOVE TILDE OPERATOR
  [0x2975]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 350);
      p.L(460, 350);
      p.stroke();
      p.M(500, 350);
      p.L(400, 410);
      p.L(400, 290);
      p.Z();
      p.fill();
      p.lineWidth(35);
      p.M(140, 170);
      p.C(220, 230, 380, 110, 460, 170);
      p.stroke();
    }
  },
  // ⥶ LEFTWARDS ARROW ABOVE TILDE OPERATOR
  [0x2976]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 350);
      p.L(140, 350);
      p.stroke();
      p.M(100, 350);
      p.L(200, 410);
      p.L(200, 290);
      p.Z();
      p.fill();
      p.lineWidth(35);
      p.M(140, 170);
      p.C(220, 230, 380, 110, 460, 170);
      p.stroke();
    }
  },
  // ⥷ LEFTWARDS ARROW ABOVE ALMOST EQUAL TO
  [0x2977]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 350);
      p.L(140, 350);
      p.stroke();
      p.M(100, 350);
      p.L(200, 410);
      p.L(200, 290);
      p.Z();
      p.fill();
      p.lineWidth(30);
      p.M(140, 200);
      p.C(220, 240, 360, 160, 460, 200);
      p.stroke();
      p.M(140, 130);
      p.C(220, 170, 360, 90, 460, 130);
      p.stroke();
    }
  },
  // ⥸ RIGHTWARDS ARROW ABOVE NOT ALMOST EQUAL TO
  [0x2978]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(100, 380);
      p.L(460, 380);
      p.stroke();
      p.M(500, 380);
      p.L(400, 440);
      p.L(400, 320);
      p.Z();
      p.fill();
      p.lineWidth(30);
      p.M(140, 220);
      p.C(220, 260, 360, 180, 460, 220);
      p.stroke();
      p.M(140, 150);
      p.C(220, 190, 360, 110, 460, 150);
      p.stroke();
      p.M(350, 260);
      p.L(250, 100);
      p.stroke();
    }
  },
  // ⥹ LEFT RIGHT DOUBLE ARROW WITH TAIL
  [0x2979]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(120, 250);
      p.L(480, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
    }
  },
  // ⥺ LEFTWARDS ARROW THROUGH X
  [0x297a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(130, 250);
      p.stroke();
      p.M(100, 250);
      p.L(210, 350);
      p.L(210, 150);
      p.Z();
      p.fill();
      p.M(280, 180);
      p.L(370, 320);
      p.stroke();
      p.M(370, 180);
      p.L(280, 320);
      p.stroke();
    }
  },
  // ⥻ RIGHTWARDS ARROW THROUGH X
  [0x297b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(470, 250);
      p.stroke();
      p.M(500, 250);
      p.L(390, 350);
      p.L(390, 150);
      p.Z();
      p.fill();
      p.M(230, 180);
      p.L(320, 320);
      p.stroke();
      p.M(320, 180);
      p.L(230, 320);
      p.stroke();
    }
  },
  // ⥼ LEFT BARB UP RIGHT BARB UP LONG HARPOON
  [0x297c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 250);
      p.L(520, 250);
      p.stroke();
      p.M(520, 250);
      p.L(420, 370);
      p.L(420, 250);
      p.Z();
      p.fill();
    }
  },
  // ⥽ LEFT BARB DOWN RIGHT BARB DOWN LONG HARPOON
  [0x297d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(80, 250);
      p.L(520, 250);
      p.stroke();
      p.M(520, 250);
      p.L(420, 130);
      p.L(420, 250);
      p.Z();
      p.fill();
    }
  },
  // ⥾ UP BARB LEFT DOWN BARB RIGHT LONG HARPOON
  [0x297e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(400, 350);
      p.L(300, 350);
      p.Z();
      p.fill();
    }
  },
  // ⥿ UP BARB RIGHT DOWN BARB LEFT LONG HARPOON
  [0x297f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
      p.M(300, 450);
      p.L(200, 350);
      p.L(300, 350);
      p.Z();
      p.fill();
    }
  }
};

// =============================================================================
// Misc Math Symbols-A Extended (U+27C0–U+27EF) — remaining chars
// =============================================================================

export const MISC_MATH_A_EXT: Record<number, GlyphDef> = {
  // ⟀ THREE DIMENSIONAL ANGLE
  [0x27c0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(450, 80);
      p.L(150, 80);
      p.L(350, 420);
      p.stroke();
      p.M(150, 80);
      p.L(250, 200);
      p.stroke();
    }
  },
  // ⟁ WHITE TRIANGLE CONTAINING SMALL WHITE TRIANGLE
  [0x27c1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(100, 50);
      p.L(500, 50);
      p.Z();
      p.stroke();
      // inner small triangle
      p.M(300, 320);
      p.L(220, 140);
      p.L(380, 140);
      p.Z();
      p.stroke();
    }
  },
  // ⟂ PERPENDICULAR
  [0x27c2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(100, 80);
      p.L(500, 80);
      p.stroke();
      p.M(300, 80);
      p.L(300, 420);
      p.stroke();
    }
  },
  // ⟃ OPEN SUBSET
  [0x27c3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(450, 420);
      p.C(250, 420, 150, 350, 150, 250);
      p.C(150, 150, 250, 80, 450, 80);
      p.stroke();
    }
  },
  // ⟄ OPEN SUPERSET
  [0x27c4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 420);
      p.C(350, 420, 450, 350, 450, 250);
      p.C(450, 150, 350, 80, 150, 80);
      p.stroke();
    }
  },
  // ⟅ LEFT S-SHAPED BAG DELIMITER
  [0x27c5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(350, 460);
      p.C(200, 460, 200, 300, 300, 250);
      p.C(400, 200, 400, 40, 250, 40);
      p.stroke();
    }
  },
  // ⟆ RIGHT S-SHAPED BAG DELIMITER
  [0x27c6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 460);
      p.C(400, 460, 400, 300, 300, 250);
      p.C(200, 200, 200, 40, 350, 40);
      p.stroke();
    }
  },
  // ⟇ OR WITH DOT INSIDE
  [0x27c7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(100, 420);
      p.L(300, 80);
      p.L(500, 420);
      p.stroke();
      p.circle(300, 280, 30);
      p.fill();
    }
  },
  // ⟈ REVERSE SOLIDUS PRECEDING SUBSET
  [0x27c8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 420);
      p.L(300, 80);
      p.stroke();
      p.M(450, 420);
      p.C(350, 420, 300, 350, 300, 250);
      p.C(300, 150, 350, 80, 450, 80);
      p.stroke();
    }
  },
  // ⟉ SUPERSET PRECEDING SOLIDUS
  [0x27c9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 420);
      p.C(250, 420, 300, 350, 300, 250);
      p.C(300, 150, 250, 80, 150, 80);
      p.stroke();
      p.M(450, 420);
      p.L(300, 80);
      p.stroke();
    }
  },
  // ⟊ VERTICAL BAR WITH HORIZONTAL STROKE
  [0x27ca]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 80);
      p.L(300, 420);
      p.stroke();
      p.M(150, 250);
      p.L(450, 250);
      p.stroke();
    }
  },
  // ⟋ MATHEMATICAL RISING DIAGONAL
  [0x27cb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 80);
      p.L(480, 420);
      p.stroke();
    }
  },
  // ⟌ LONG DIVISION
  [0x27cc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 80);
      p.C(150, 80, 120, 120, 120, 200);
      p.L(120, 420);
      p.stroke();
      p.M(200, 420);
      p.L(500, 420);
      p.stroke();
    }
  },
  // ⟍ MATHEMATICAL FALLING DIAGONAL
  [0x27cd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 420);
      p.L(480, 80);
      p.stroke();
    }
  },
  // ⟎ SQUARED LOGICAL AND
  [0x27ce]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(200, 150);
      p.L(300, 380);
      p.L(400, 150);
      p.stroke();
    }
  },
  // ⟏ SQUARED LOGICAL OR
  [0x27cf]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(200, 350);
      p.L(300, 120);
      p.L(400, 350);
      p.stroke();
    }
  },
  // ⟐ WHITE DIAMOND WITH CENTRED DOT
  [0x27d0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(510, 250);
      p.L(300, 30);
      p.L(90, 250);
      p.Z();
      p.stroke();
      p.circle(300, 250, 30);
      p.fill();
    }
  },
  // ⟑ AND WITH DOT
  [0x27d1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(100, 80);
      p.L(300, 420);
      p.L(500, 80);
      p.stroke();
      p.circle(300, 250, 30);
      p.fill();
    }
  },
  // ⟒ ELEMENT OF OPENING UPWARDS
  [0x27d2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 80);
      p.L(450, 80);
      p.stroke();
      p.M(300, 80);
      p.L(300, 420);
      p.stroke();
    }
  },
  // ⟓ LOWER RIGHT CORNER WITH DOT
  [0x27d3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 420);
      p.L(450, 420);
      p.L(450, 80);
      p.stroke();
      p.circle(300, 250, 30);
      p.fill();
    }
  },
  // ⟔ UPPER LEFT CORNER WITH DOT
  [0x27d4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 80);
      p.L(150, 420);
      p.stroke();
      p.M(150, 420);
      p.L(450, 420);
      p.stroke();
      p.circle(300, 250, 30);
      p.fill();
    }
  },
  // ⟕ LEFT OUTER JOIN
  [0x27d5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 420);
      p.L(500, 250);
      p.L(300, 80);
      p.stroke();
      p.M(100, 420);
      p.L(300, 420);
      p.stroke();
      p.M(100, 80);
      p.L(300, 80);
      p.stroke();
      p.M(100, 80);
      p.L(100, 420);
      p.stroke();
    }
  },
  // ⟖ RIGHT OUTER JOIN
  [0x27d6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 420);
      p.L(100, 250);
      p.L(300, 80);
      p.stroke();
      p.M(300, 420);
      p.L(500, 420);
      p.stroke();
      p.M(300, 80);
      p.L(500, 80);
      p.stroke();
      p.M(500, 80);
      p.L(500, 420);
      p.stroke();
    }
  },
  // ⟗ FULL OUTER JOIN
  [0x27d7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 420);
      p.L(500, 250);
      p.L(300, 80);
      p.L(100, 250);
      p.Z();
      p.stroke();
    }
  },
  // ⟘ LARGE DOWN TACK
  [0x27d8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(100, 80);
      p.L(500, 80);
      p.stroke();
      p.M(300, 80);
      p.L(300, 420);
      p.stroke();
    }
  },
  // ⟙ LARGE UP TACK
  [0x27d9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(100, 420);
      p.L(500, 420);
      p.stroke();
      p.M(300, 420);
      p.L(300, 80);
      p.stroke();
    }
  },
  // ⟚ LEFT AND RIGHT DOUBLE TURNSTILE
  [0x27da]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(200, 80);
      p.L(200, 420);
      p.stroke();
      p.M(250, 80);
      p.L(250, 420);
      p.stroke();
      p.M(350, 80);
      p.L(350, 420);
      p.stroke();
      p.M(400, 80);
      p.L(400, 420);
      p.stroke();
      p.M(250, 250);
      p.L(350, 250);
      p.stroke();
    }
  },
  // ⟛ LEFT AND RIGHT TACK
  [0x27db]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 80);
      p.L(200, 420);
      p.stroke();
      p.M(400, 80);
      p.L(400, 420);
      p.stroke();
      p.M(200, 250);
      p.L(400, 250);
      p.stroke();
    }
  },
  // ⟜ LEFT MULTIMAP
  [0x27dc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 250);
      p.L(400, 250);
      p.stroke();
      p.circle(450, 250, 50);
      p.stroke();
    }
  },
  // ⟝ LONG RIGHT TACK
  [0x27dd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 80);
      p.L(150, 420);
      p.stroke();
      p.M(150, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // ⟞ LONG LEFT TACK
  [0x27de]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(450, 80);
      p.L(450, 420);
      p.stroke();
      p.M(100, 250);
      p.L(450, 250);
      p.stroke();
    }
  },
  // ⟟ UP TACK WITH CIRCLE ABOVE
  [0x27df]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 100);
      p.L(500, 100);
      p.stroke();
      p.M(300, 100);
      p.L(300, 320);
      p.stroke();
      p.circle(300, 390, 50);
      p.stroke();
    }
  },
  // ⟠ LOZENGE DIVIDED BY HORIZONTAL RULE
  [0x27e0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(510, 250);
      p.L(300, 30);
      p.L(90, 250);
      p.Z();
      p.stroke();
      p.M(90, 250);
      p.L(510, 250);
      p.stroke();
    }
  },
  // ⟡ CONCAVE DIAMOND WITH TICK LEFT
  [0x27e1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.C(250, 300, 100, 250, 100, 250);
      p.C(100, 250, 250, 200, 300, 50);
      p.C(350, 200, 500, 250, 500, 250);
      p.C(500, 250, 350, 300, 300, 450);
      p.stroke();
    }
  },
  // ⟢ CONCAVE DIAMOND WITH TICK RIGHT
  [0x27e2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(100, 250);
      p.L(300, 50);
      p.L(500, 250);
      p.Z();
      p.stroke();
      p.M(400, 250);
      p.L(300, 250);
      p.stroke();
    }
  },
  // ⟣ WHITE CONCAVE-SIDED DIAMOND
  [0x27e3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.C(250, 350, 100, 300, 100, 250);
      p.C(100, 200, 250, 150, 300, 50);
      p.C(350, 150, 500, 200, 500, 250);
      p.C(500, 300, 350, 350, 300, 450);
      p.stroke();
    }
  },
  // ⟤ WHITE CONCAVE-SIDED DIAMOND WITH LEFTWARDS TICK
  [0x27e4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.C(250, 350, 100, 300, 100, 250);
      p.C(100, 200, 250, 150, 300, 50);
      p.C(350, 150, 500, 200, 500, 250);
      p.C(500, 300, 350, 350, 300, 450);
      p.stroke();
      p.M(200, 250);
      p.L(300, 250);
      p.stroke();
    }
  },
  // ⟥ WHITE CONCAVE-SIDED DIAMOND WITH RIGHTWARDS TICK
  [0x27e5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.C(250, 350, 100, 300, 100, 250);
      p.C(100, 200, 250, 150, 300, 50);
      p.C(350, 150, 500, 200, 500, 250);
      p.C(500, 300, 350, 350, 300, 450);
      p.stroke();
      p.M(300, 250);
      p.L(400, 250);
      p.stroke();
    }
  },
  // ⟦ MATHEMATICAL LEFT WHITE SQUARE BRACKET
  [0x27e6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(380, 470);
      p.L(200, 470);
      p.L(200, 30);
      p.L(380, 30);
      p.stroke();
      p.M(260, 470);
      p.L(260, 30);
      p.stroke();
    }
  },
  // ⟧ MATHEMATICAL RIGHT WHITE SQUARE BRACKET
  [0x27e7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(220, 470);
      p.L(400, 470);
      p.L(400, 30);
      p.L(220, 30);
      p.stroke();
      p.M(340, 470);
      p.L(340, 30);
      p.stroke();
    }
  },
  // ⟬ MATHEMATICAL LEFT WHITE TORTOISE SHELL BRACKET
  [0x27ec]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(380, 470);
      p.C(200, 400, 200, 100, 380, 30);
      p.stroke();
      p.M(340, 460);
      p.C(180, 400, 180, 100, 340, 40);
      p.stroke();
    }
  },
  // ⟭ MATHEMATICAL RIGHT WHITE TORTOISE SHELL BRACKET
  [0x27ed]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(220, 470);
      p.C(400, 400, 400, 100, 220, 30);
      p.stroke();
      p.M(260, 460);
      p.C(420, 400, 420, 100, 260, 40);
      p.stroke();
    }
  },
  // ⟮ MATHEMATICAL LEFT FLATTENED PARENTHESIS
  [0x27ee]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(370, 470);
      p.C(230, 380, 230, 120, 370, 30);
      p.stroke();
    }
  },
  // ⟯ MATHEMATICAL RIGHT FLATTENED PARENTHESIS
  [0x27ef]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(230, 470);
      p.C(370, 380, 370, 120, 230, 30);
      p.stroke();
    }
  }
};

// =============================================================================
// Misc Symbols & Arrows Extended (U+2B00–U+2BFF) — remaining chars
// =============================================================================

export const MISC_SYM_ARROWS_EXT: Record<number, GlyphDef> = {
  // ⬀ NORTH EAST WHITE ARROW
  [0x2b00]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 80);
      p.L(480, 420);
      p.stroke();
      p.M(480, 420);
      p.L(350, 420);
      p.L(480, 290);
      p.Z();
      p.stroke();
    }
  },
  // ⬁ NORTH WEST WHITE ARROW
  [0x2b01]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(480, 80);
      p.L(120, 420);
      p.stroke();
      p.M(120, 420);
      p.L(120, 290);
      p.L(250, 420);
      p.Z();
      p.stroke();
    }
  },
  // ⬂ SOUTH EAST WHITE ARROW
  [0x2b02]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 420);
      p.L(480, 80);
      p.stroke();
      p.M(480, 80);
      p.L(480, 210);
      p.L(350, 80);
      p.Z();
      p.stroke();
    }
  },
  // ⬃ SOUTH WEST WHITE ARROW
  [0x2b03]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(480, 420);
      p.L(120, 80);
      p.stroke();
      p.M(120, 80);
      p.L(250, 80);
      p.L(120, 210);
      p.Z();
      p.stroke();
    }
  },
  // ⬄ LEFT RIGHT WHITE ARROW
  [0x2b04]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(60, 250);
      p.L(200, 400);
      p.L(200, 310);
      p.L(400, 310);
      p.L(400, 400);
      p.L(540, 250);
      p.L(400, 100);
      p.L(400, 190);
      p.L(200, 190);
      p.L(200, 100);
      p.Z();
      p.stroke();
    }
  },
  // ⬈ NORTH EAST BLACK ARROW
  [0x2b08]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(120, 80);
      p.L(480, 420);
      p.L(350, 420);
      p.L(480, 290);
      p.Z();
      p.fill();
      p.lineWidth(50);
      p.M(140, 100);
      p.L(440, 400);
      p.stroke();
    }
  },
  // ⬉ NORTH WEST BLACK ARROW
  [0x2b09]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(480, 80);
      p.L(120, 420);
      p.L(120, 290);
      p.L(250, 420);
      p.Z();
      p.fill();
      p.lineWidth(50);
      p.M(460, 100);
      p.L(160, 400);
      p.stroke();
    }
  },
  // ⬊ SOUTH EAST BLACK ARROW
  [0x2b0a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(120, 420);
      p.L(480, 80);
      p.L(480, 210);
      p.L(350, 80);
      p.Z();
      p.fill();
      p.lineWidth(50);
      p.M(140, 400);
      p.L(460, 100);
      p.stroke();
    }
  },
  // ⬋ SOUTH WEST BLACK ARROW
  [0x2b0b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(480, 420);
      p.L(120, 80);
      p.L(250, 80);
      p.L(120, 210);
      p.Z();
      p.fill();
      p.lineWidth(50);
      p.M(460, 400);
      p.L(160, 100);
      p.stroke();
    }
  },
  // ⬌ LEFT RIGHT BLACK ARROW
  [0x2b0c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(60, 250);
      p.L(200, 400);
      p.L(200, 310);
      p.L(400, 310);
      p.L(400, 400);
      p.L(540, 250);
      p.L(400, 100);
      p.L(400, 190);
      p.L(200, 190);
      p.L(200, 100);
      p.Z();
      p.fill();
    }
  },
  // ⬍ UP DOWN BLACK ARROW
  [0x2b0d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 480);
      p.L(450, 340);
      p.L(370, 340);
      p.L(370, 160);
      p.L(450, 160);
      p.L(300, 20);
      p.L(150, 160);
      p.L(230, 160);
      p.L(230, 340);
      p.L(150, 340);
      p.Z();
      p.fill();
    }
  },
  // ⬎ RIGHT ARROW TO LOWER RIGHT CORNER
  [0x2b0e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 350);
      p.L(400, 350);
      p.L(400, 150);
      p.stroke();
      p.M(400, 150);
      p.L(330, 250);
      p.L(470, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬏ RIGHT ARROW TO UPPER RIGHT CORNER
  [0x2b0f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 150);
      p.L(400, 150);
      p.L(400, 350);
      p.stroke();
      p.M(400, 350);
      p.L(330, 250);
      p.L(470, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬐ LEFT ARROW TO LOWER LEFT CORNER
  [0x2b10]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 350);
      p.L(200, 350);
      p.L(200, 150);
      p.stroke();
      p.M(200, 150);
      p.L(130, 250);
      p.L(270, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬑ LEFT ARROW TO UPPER LEFT CORNER
  [0x2b11]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 150);
      p.L(200, 150);
      p.L(200, 350);
      p.stroke();
      p.M(200, 350);
      p.L(130, 250);
      p.L(270, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬒ SQUARE WITH TOP HALF BLACK
  [0x2b12]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.rect(100, 250, 400, 200);
      p.fill();
    }
  },
  // ⬓ SQUARE WITH BOTTOM HALF BLACK
  [0x2b13]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.rect(100, 50, 400, 200);
      p.fill();
    }
  },
  // ⬔ SQUARE WITH UPPER RIGHT DIAGONAL HALF BLACK
  [0x2b14]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(100, 450);
      p.L(500, 450);
      p.L(500, 50);
      p.Z();
      p.fill();
    }
  },
  // ⬕ SQUARE WITH LOWER LEFT DIAGONAL HALF BLACK
  [0x2b15]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.rect(100, 50, 400, 400);
      p.stroke();
      p.M(100, 50);
      p.L(100, 450);
      p.L(500, 50);
      p.Z();
      p.fill();
    }
  },
  // ⬖ DIAMOND WITH LEFT HALF BLACK
  [0x2b16]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(510, 250);
      p.L(300, 30);
      p.L(90, 250);
      p.Z();
      p.stroke();
      p.M(300, 470);
      p.L(90, 250);
      p.L(300, 30);
      p.Z();
      p.fill();
    }
  },
  // ⬗ DIAMOND WITH RIGHT HALF BLACK
  [0x2b17]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(510, 250);
      p.L(300, 30);
      p.L(90, 250);
      p.Z();
      p.stroke();
      p.M(300, 470);
      p.L(510, 250);
      p.L(300, 30);
      p.Z();
      p.fill();
    }
  },
  // ⬘ DIAMOND WITH TOP HALF BLACK
  [0x2b18]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(510, 250);
      p.L(300, 30);
      p.L(90, 250);
      p.Z();
      p.stroke();
      p.M(90, 250);
      p.L(300, 470);
      p.L(510, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬙ DIAMOND WITH BOTTOM HALF BLACK
  [0x2b19]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(510, 250);
      p.L(300, 30);
      p.L(90, 250);
      p.Z();
      p.stroke();
      p.M(90, 250);
      p.L(300, 30);
      p.L(510, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬚ DOTTED SQUARE
  [0x2b1a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(25);
      // draw dashed square as segments
      for (let i = 0; i < 8; i++) {
        const x = 100 + i * 50;
        p.M(x, 450);
        p.L(x + 30, 450);
      }
      for (let i = 0; i < 8; i++) {
        const x = 100 + i * 50;
        p.M(x, 50);
        p.L(x + 30, 50);
      }
      for (let i = 0; i < 8; i++) {
        const y = 50 + i * 50;
        p.M(100, y);
        p.L(100, y + 30);
      }
      for (let i = 0; i < 8; i++) {
        const y = 50 + i * 50;
        p.M(500, y);
        p.L(500, y + 30);
      }
      p.stroke();
    }
  },
  // ⬝ BLACK VERY SMALL SQUARE
  [0x2b1d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(240, 190, 120, 120);
      p.fill();
    }
  },
  // ⬞ WHITE VERY SMALL SQUARE
  [0x2b1e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.rect(240, 190, 120, 120);
      p.stroke();
    }
  },
  // ⬟ BLACK PENTAGON
  [0x2b1f]: {
    width: W,
    draw: (p: GlyphPen) => {
      const cx = 300,
        cy = 250,
        r = 200;
      for (let i = 0; i < 5; i++) {
        const a = ((i * 72 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + r * Math.cos(a), cy + r * Math.sin(a));
        } else {
          p.L(cx + r * Math.cos(a), cy + r * Math.sin(a));
        }
      }
      p.Z();
      p.fill();
    }
  },
  // ⬠ WHITE PENTAGON
  [0x2b20]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      const cx = 300,
        cy = 250,
        r = 200;
      for (let i = 0; i < 5; i++) {
        const a = ((i * 72 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + r * Math.cos(a), cy + r * Math.sin(a));
        } else {
          p.L(cx + r * Math.cos(a), cy + r * Math.sin(a));
        }
      }
      p.Z();
      p.stroke();
    }
  },
  // ⬡ WHITE HEXAGON
  [0x2b21]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      const cx = 300,
        cy = 250,
        r = 200;
      for (let i = 0; i < 6; i++) {
        const a = ((i * 60 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + r * Math.cos(a), cy + r * Math.sin(a));
        } else {
          p.L(cx + r * Math.cos(a), cy + r * Math.sin(a));
        }
      }
      p.Z();
      p.stroke();
    }
  },
  // ⬢ BLACK HEXAGON
  [0x2b22]: {
    width: W,
    draw: (p: GlyphPen) => {
      const cx = 300,
        cy = 250,
        r = 200;
      for (let i = 0; i < 6; i++) {
        const a = ((i * 60 - 90) * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + r * Math.cos(a), cy + r * Math.sin(a));
        } else {
          p.L(cx + r * Math.cos(a), cy + r * Math.sin(a));
        }
      }
      p.Z();
      p.fill();
    }
  },
  // ⬣ HORIZONTAL BLACK HEXAGON
  [0x2b23]: {
    width: W,
    draw: (p: GlyphPen) => {
      const cx = 300,
        cy = 250,
        r = 200;
      for (let i = 0; i < 6; i++) {
        const a = (i * 60 * Math.PI) / 180;
        if (i === 0) {
          p.M(cx + r * Math.cos(a), cy + r * Math.sin(a));
        } else {
          p.L(cx + r * Math.cos(a), cy + r * Math.sin(a));
        }
      }
      p.Z();
      p.fill();
    }
  },
  // ⬤ BLACK LARGE CIRCLE
  [0x2b24]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 230);
      p.fill();
    }
  },
  // ⬥ BLACK MEDIUM DIAMOND
  [0x2b25]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 470);
      p.L(480, 250);
      p.L(300, 30);
      p.L(120, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬦ WHITE MEDIUM DIAMOND
  [0x2b26]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(480, 250);
      p.L(300, 30);
      p.L(120, 250);
      p.Z();
      p.stroke();
    }
  },
  // ⬧ BLACK MEDIUM LOZENGE
  [0x2b27]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 490);
      p.L(460, 250);
      p.L(300, 10);
      p.L(140, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬨ WHITE MEDIUM LOZENGE
  [0x2b28]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 490);
      p.L(460, 250);
      p.L(300, 10);
      p.L(140, 250);
      p.Z();
      p.stroke();
    }
  },
  // ⬩ BLACK SMALL DIAMOND
  [0x2b29]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 380);
      p.L(400, 250);
      p.L(300, 120);
      p.L(200, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬪ BLACK SMALL LOZENGE
  [0x2b2a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 400);
      p.L(380, 250);
      p.L(300, 100);
      p.L(220, 250);
      p.Z();
      p.fill();
    }
  },
  // ⬫ WHITE SMALL LOZENGE
  [0x2b2b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 400);
      p.L(380, 250);
      p.L(300, 100);
      p.L(220, 250);
      p.Z();
      p.stroke();
    }
  },
  // ⬬ BLACK HORIZONTAL ELLIPSE
  [0x2b2c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.ellipse(300, 250, 220, 120);
      p.fill();
    }
  },
  // ⬭ WHITE HORIZONTAL ELLIPSE
  [0x2b2d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.ellipse(300, 250, 220, 120);
      p.stroke();
    }
  },
  // ⬮ BLACK VERTICAL ELLIPSE
  [0x2b2e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.ellipse(300, 250, 120, 220);
      p.fill();
    }
  },
  // ⬯ WHITE VERTICAL ELLIPSE
  [0x2b2f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.ellipse(300, 250, 120, 220);
      p.stroke();
    }
  },
  // ⬰ LEFT ARROW WITH SMALL CIRCLE
  [0x2b30]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(480, 250);
      p.L(180, 250);
      p.stroke();
      p.M(100, 250);
      p.L(230, 370);
      p.L(230, 130);
      p.Z();
      p.fill();
      p.circle(510, 250, 30);
      p.fill();
    }
  },
  // ⬱ THREE LEFTWARDS ARROWS
  [0x2b31]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(25);
      for (let y = 150; y <= 350; y += 100) {
        p.M(500, y);
        p.L(200, y);
        p.stroke();
        p.M(120, y);
        p.L(220, y + 60);
        p.L(220, y - 60);
        p.Z();
        p.fill();
      }
    }
  },
  // ⬲ LEFT ARROW WITH CIRCLED PLUS
  [0x2b32]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(460, 250);
      p.L(200, 250);
      p.stroke();
      p.M(120, 250);
      p.L(220, 350);
      p.L(220, 150);
      p.Z();
      p.fill();
      p.circle(500, 250, 40);
      p.stroke();
      p.M(480, 250);
      p.L(520, 250);
      p.stroke();
      p.M(500, 230);
      p.L(500, 270);
      p.stroke();
    }
  },
  // ⬳ LONG LEFTWARDS SQUIGGLE ARROW
  [0x2b33]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(520, 250);
      p.C(480, 330, 420, 170, 360, 250);
      p.C(300, 330, 240, 170, 180, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 340);
      p.L(200, 160);
      p.Z();
      p.fill();
    }
  },
  // ⬴ LEFTWARDS TWO-HEADED ARROW WITH VERTICAL STROKE
  [0x2b34]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
      p.M(300, 150);
      p.L(300, 350);
      p.stroke();
    }
  },
  // ⬵ LEFTWARDS TWO-HEADED ARROW WITH DOUBLE VERTICAL STROKE
  [0x2b35]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
      p.M(280, 150);
      p.L(280, 350);
      p.stroke();
      p.M(320, 150);
      p.L(320, 350);
      p.stroke();
    }
  },
  // ⬶ LEFTWARDS TWO-HEADED ARROW FROM BAR
  [0x2b36]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 150);
      p.L(500, 350);
      p.stroke();
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(400, 250);
      p.L(450, 320);
      p.L(450, 180);
      p.Z();
      p.fill();
    }
  },
  // ⬷ LEFTWARDS TWO-HEADED TRIPLE DASH ARROW
  [0x2b37]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 250);
      p.L(430, 250);
      p.stroke();
      p.M(380, 250);
      p.L(310, 250);
      p.stroke();
      p.M(260, 250);
      p.L(190, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(430, 330);
      p.L(430, 170);
      p.Z();
      p.fill();
    }
  },
  // ⬸ LEFTWARDS ARROW WITH DOTTED STEM
  [0x2b38]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(450, 250, 20);
      p.fill();
      p.circle(370, 250, 20);
      p.fill();
      p.circle(290, 250, 20);
      p.fill();
      p.lineWidth(40);
      p.M(230, 250);
      p.L(170, 250);
      p.stroke();
      p.M(100, 250);
      p.L(220, 350);
      p.L(220, 150);
      p.Z();
      p.fill();
    }
  },
  // ⬹ LEFTWARDS ARROW WITH TAIL WITH VERTICAL STROKE
  [0x2b39]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(150, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(500, 350);
      p.L(500, 150);
      p.stroke();
      p.M(350, 150);
      p.L(350, 350);
      p.stroke();
    }
  },
  // ⬺ LEFTWARDS ARROW WITH TAIL WITH DOUBLE VERTICAL STROKE
  [0x2b3a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 250);
      p.L(150, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 340);
      p.L(200, 160);
      p.Z();
      p.fill();
      p.M(500, 340);
      p.L(500, 160);
      p.stroke();
      p.M(330, 160);
      p.L(330, 340);
      p.stroke();
      p.M(370, 160);
      p.L(370, 340);
      p.stroke();
    }
  },
  // ⬻ LEFTWARDS TWO-HEADED ARROW WITH TAIL
  [0x2b3b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 350);
      p.L(200, 150);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(400, 350);
      p.L(400, 150);
      p.Z();
      p.fill();
      p.M(500, 350);
      p.L(500, 150);
      p.stroke();
    }
  },
  // ⬼ LEFTWARDS TWO-HEADED ARROW WITH TAIL WITH VERTICAL STROKE
  [0x2b3c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(520, 250);
      p.L(80, 250);
      p.stroke();
      p.M(80, 250);
      p.L(170, 340);
      p.L(170, 160);
      p.Z();
      p.fill();
      p.M(520, 250);
      p.L(430, 340);
      p.L(430, 160);
      p.Z();
      p.fill();
      p.M(520, 340);
      p.L(520, 160);
      p.stroke();
      p.M(300, 160);
      p.L(300, 340);
      p.stroke();
    }
  },
  // ⬽ LEFTWARDS TWO-HEADED ARROW WITH TAIL WITH DOUBLE VERTICAL STROKE
  [0x2b3d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(530, 250);
      p.L(70, 250);
      p.stroke();
      p.M(70, 250);
      p.L(160, 340);
      p.L(160, 160);
      p.Z();
      p.fill();
      p.M(530, 250);
      p.L(440, 340);
      p.L(440, 160);
      p.Z();
      p.fill();
      p.M(530, 340);
      p.L(530, 160);
      p.stroke();
      p.M(280, 160);
      p.L(280, 340);
      p.stroke();
      p.M(320, 160);
      p.L(320, 340);
      p.stroke();
    }
  },
  // ⬾ LEFTWARDS ARROW THROUGH X
  [0x2b3e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(130, 250);
      p.stroke();
      p.M(100, 250);
      p.L(210, 350);
      p.L(210, 150);
      p.Z();
      p.fill();
      p.M(310, 180);
      p.L(390, 320);
      p.stroke();
      p.M(390, 180);
      p.L(310, 320);
      p.stroke();
    }
  },
  // ⬿ WAVE ARROW POINTING DIRECTLY LEFT
  [0x2b3f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 250);
      p.C(440, 330, 380, 170, 300, 250);
      p.C(240, 330, 200, 170, 160, 250);
      p.stroke();
      p.M(100, 250);
      p.L(200, 340);
      p.L(200, 160);
      p.Z();
      p.fill();
    }
  },
  // ⭀ EQUALS SIGN ABOVE LEFTWARDS ARROW
  [0x2b40]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 200);
      p.L(140, 200);
      p.stroke();
      p.M(100, 200);
      p.L(200, 270);
      p.L(200, 130);
      p.Z();
      p.fill();
      p.M(150, 350);
      p.L(470, 350);
      p.stroke();
      p.M(150, 400);
      p.L(470, 400);
      p.stroke();
    }
  },
  // ⭁ REVERSE TILDE OPERATOR ABOVE LEFTWARDS ARROW
  [0x2b41]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(500, 200);
      p.L(140, 200);
      p.stroke();
      p.M(100, 200);
      p.L(200, 270);
      p.L(200, 130);
      p.Z();
      p.fill();
      p.M(470, 380);
      p.C(400, 430, 200, 330, 140, 380);
      p.stroke();
    }
  },
  // ⭅ LEFTWARDS QUADRUPLE ARROW
  [0x2b45]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(25);
      p.M(500, 330);
      p.L(200, 330);
      p.stroke();
      p.M(500, 280);
      p.L(200, 280);
      p.stroke();
      p.M(500, 230);
      p.L(200, 230);
      p.stroke();
      p.M(500, 180);
      p.L(200, 180);
      p.stroke();
      p.M(100, 250);
      p.L(230, 400);
      p.L(230, 100);
      p.Z();
      p.fill();
    }
  },
  // ⭆ RIGHTWARDS QUADRUPLE ARROW
  [0x2b46]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(25);
      p.M(100, 330);
      p.L(400, 330);
      p.stroke();
      p.M(100, 280);
      p.L(400, 280);
      p.stroke();
      p.M(100, 230);
      p.L(400, 230);
      p.stroke();
      p.M(100, 180);
      p.L(400, 180);
      p.stroke();
      p.M(500, 250);
      p.L(370, 400);
      p.L(370, 100);
      p.Z();
      p.fill();
    }
  },
  // ⭐ already exists at 2b50 — skip
  // ⭑ BLACK STAR
  [0x2b51]: {
    width: W,
    draw: (p: GlyphPen) => {
      const cx = 300,
        cy = 260,
        R = 200,
        r = 80;
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
  // ⭒ WHITE STAR
  [0x2b52]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      const cx = 300,
        cy = 260,
        R = 200,
        r = 80;
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
  // ⭓ BLACK STAR = IDEOGRAPH FOUR
  [0x2b53]: {
    width: W,
    draw: (p: GlyphPen) => {
      const cx = 300,
        cy = 260,
        R = 210,
        r = 90;
      for (let i = 0; i < 4; i++) {
        const a1 = ((i * 90 - 90) * Math.PI) / 180;
        const a2 = ((i * 90 + 45 - 90) * Math.PI) / 180;
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
  // ⭔ WHITE FOUR POINTED STAR
  [0x2b54]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      const cx = 300,
        cy = 260,
        R = 210,
        r = 90;
      for (let i = 0; i < 4; i++) {
        const a1 = ((i * 90 - 90) * Math.PI) / 180;
        const a2 = ((i * 90 + 45 - 90) * Math.PI) / 180;
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
  // ⭕ HEAVY LARGE CIRCLE
  [0x2b55]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(60);
      p.circle(300, 250, 210);
      p.stroke();
    }
  },
  // ⭖ BLACK VERY SMALL CIRCLE (filled dot)
  [0x2b56]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 40);
      p.fill();
    }
  },
  // ⭗ HEAVY WHITE CIRCLE
  [0x2b57]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.circle(300, 250, 180);
      p.stroke();
    }
  },
  // ⭘ HEAVY CIRCLE
  [0x2b58]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.circle(300, 250, 200);
      p.stroke();
    }
  },
  // ⭙ HEAVY CIRCLE WITH CIRCLE INSIDE
  [0x2b59]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 200);
      p.stroke();
      p.circle(300, 250, 100);
      p.stroke();
    }
  },
  // ⭠ LEFTWARDS TRIANGLE-HEADED ARROW
  [0x2b60]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(200, 250);
      p.stroke();
      p.M(100, 250);
      p.L(230, 380);
      p.L(230, 120);
      p.Z();
      p.fill();
    }
  },
  // ⭡ UPWARDS TRIANGLE-HEADED ARROW
  [0x2b61]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 50);
      p.L(300, 350);
      p.stroke();
      p.M(300, 450);
      p.L(180, 320);
      p.L(420, 320);
      p.Z();
      p.fill();
    }
  },
  // ⭢ RIGHTWARDS TRIANGLE-HEADED ARROW
  [0x2b62]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(400, 250);
      p.stroke();
      p.M(500, 250);
      p.L(370, 380);
      p.L(370, 120);
      p.Z();
      p.fill();
    }
  },
  // ⭣ DOWNWARDS TRIANGLE-HEADED ARROW
  [0x2b63]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(300, 150);
      p.stroke();
      p.M(300, 50);
      p.L(180, 180);
      p.L(420, 180);
      p.Z();
      p.fill();
    }
  },
  // ⭤ LEFT RIGHT TRIANGLE-HEADED ARROW
  [0x2b64]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(180, 250);
      p.L(420, 250);
      p.stroke();
      p.M(100, 250);
      p.L(210, 350);
      p.L(210, 150);
      p.Z();
      p.fill();
      p.M(500, 250);
      p.L(390, 350);
      p.L(390, 150);
      p.Z();
      p.fill();
    }
  },
  // ⭥ UP DOWN TRIANGLE-HEADED ARROW
  [0x2b65]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 140);
      p.L(300, 360);
      p.stroke();
      p.M(300, 460);
      p.L(200, 350);
      p.L(400, 350);
      p.Z();
      p.fill();
      p.M(300, 40);
      p.L(200, 150);
      p.L(400, 150);
      p.Z();
      p.fill();
    }
  },
  // ⭰ LEFTWARDS TRIANGLE-HEADED ARROW WITH DOUBLE HORIZONTAL STROKE
  [0x2b70]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(500, 280);
      p.L(200, 280);
      p.stroke();
      p.M(500, 220);
      p.L(200, 220);
      p.stroke();
      p.M(100, 250);
      p.L(230, 380);
      p.L(230, 120);
      p.Z();
      p.fill();
    }
  },
  // ⭲ RIGHTWARDS TRIANGLE-HEADED ARROW WITH DOUBLE HORIZONTAL STROKE
  [0x2b72]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(100, 280);
      p.L(400, 280);
      p.stroke();
      p.M(100, 220);
      p.L(400, 220);
      p.stroke();
      p.M(500, 250);
      p.L(370, 380);
      p.L(370, 120);
      p.Z();
      p.fill();
    }
  },
  // ⭰ through ⮕ — various arrows
  // ⮕ RIGHTWARDS BLACK ARROW
  [0x2b95]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(550, 250);
      p.L(370, 420);
      p.L(370, 310);
      p.L(50, 310);
      p.L(50, 190);
      p.L(370, 190);
      p.L(370, 80);
      p.Z();
      p.fill();
    }
  },
  // ⭮ ANTICLOCKWISE TRIANGLE-HEADED RIGHT U-SHAPED ARROW
  [0x2b6e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 400);
      p.L(150, 200);
      p.C(150, 80, 300, 50, 450, 80);
      p.C(500, 100, 520, 200, 450, 250);
      p.stroke();
      p.M(150, 400);
      p.L(80, 300);
      p.L(220, 300);
      p.Z();
      p.fill();
    }
  },
  // ⭯ ANTICLOCKWISE TRIANGLE-HEADED LEFT U-SHAPED ARROW
  [0x2b6f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(450, 400);
      p.L(450, 200);
      p.C(450, 80, 300, 50, 150, 80);
      p.C(100, 100, 80, 200, 150, 250);
      p.stroke();
      p.M(450, 400);
      p.L(380, 300);
      p.L(520, 300);
      p.Z();
      p.fill();
    }
  },
  // ⭾ RIGHT ARROW WITH TAIL
  [0x2b7e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 250);
      p.L(430, 250);
      p.stroke();
      p.M(500, 250);
      p.L(380, 360);
      p.L(380, 140);
      p.Z();
      p.fill();
      p.M(100, 150);
      p.L(100, 350);
      p.stroke();
    }
  },
  // ⭿ LEFT ARROW WITH TAIL
  [0x2b7f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(500, 250);
      p.L(170, 250);
      p.stroke();
      p.M(100, 250);
      p.L(220, 360);
      p.L(220, 140);
      p.Z();
      p.fill();
      p.M(500, 150);
      p.L(500, 350);
      p.stroke();
    }
  },
  // ⮈ LEFTWARDS BLACK CIRCLED WHITE ARROW
  [0x2b88]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 220);
      p.fill();
      // Arrow shape implied by filled circle
    }
  },
  // ⮊ RIGHTWARDS BLACK CIRCLED WHITE ARROW
  [0x2b8a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 220);
      p.fill();
    }
  },
  // ⮌ ANTICLOCKWISE TRIANGLE-HEADED OPEN CIRCLE ARROW
  [0x2b8c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(350, 430);
      p.C(200, 430, 100, 350, 100, 250);
      p.C(100, 150, 200, 70, 350, 70);
      p.C(480, 70, 520, 200, 500, 300);
      p.stroke();
      p.M(350, 430);
      p.L(290, 370);
      p.L(400, 370);
      p.Z();
      p.fill();
    }
  },
  // ⮎ CLOCKWISE TRIANGLE-HEADED OPEN CIRCLE ARROW
  [0x2b8e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(250, 430);
      p.C(400, 430, 500, 350, 500, 250);
      p.C(500, 150, 400, 70, 250, 70);
      p.C(120, 70, 80, 200, 100, 300);
      p.stroke();
      p.M(250, 430);
      p.L(200, 370);
      p.L(310, 370);
      p.Z();
      p.fill();
    }
  },
  // ⮐ LEFT ARROW TO BLACK DIAMOND
  [0x2b90]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(480, 250);
      p.L(200, 250);
      p.stroke();
      p.M(120, 250);
      p.L(170, 300);
      p.L(220, 250);
      p.L(170, 200);
      p.Z();
      p.fill();
    }
  },
  // ⮑ RIGHT ARROW TO BLACK DIAMOND
  [0x2b91]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 250);
      p.L(400, 250);
      p.stroke();
      p.M(480, 250);
      p.L(430, 300);
      p.L(380, 250);
      p.L(430, 200);
      p.Z();
      p.fill();
    }
  },
  // ⮝ UP TRIANGLE ARROWHEAD
  [0x2b9d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 450);
      p.L(150, 100);
      p.L(450, 100);
      p.Z();
      p.fill();
    }
  },
  // ⮞ RIGHT TRIANGLE ARROWHEAD
  [0x2b9e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(150, 450);
      p.L(500, 250);
      p.L(150, 50);
      p.Z();
      p.fill();
    }
  },
  // ⮟ DOWN TRIANGLE ARROWHEAD
  [0x2b9f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(150, 400);
      p.L(450, 400);
      p.L(300, 50);
      p.Z();
      p.fill();
    }
  },
  // ⮠ LEFT TRIANGLE ARROWHEAD
  [0x2ba0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(450, 450);
      p.L(100, 250);
      p.L(450, 50);
      p.Z();
      p.fill();
    }
  },
  // ⮡ UP TRIANGLE ARROWHEAD (white)
  [0x2ba1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 450);
      p.L(150, 100);
      p.L(450, 100);
      p.Z();
      p.stroke();
    }
  },
  // ⮢ RIGHT TRIANGLE ARROWHEAD (white)
  [0x2ba2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 450);
      p.L(500, 250);
      p.L(150, 50);
      p.Z();
      p.stroke();
    }
  },
  // ⮣ DOWN TRIANGLE ARROWHEAD (white)
  [0x2ba3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 400);
      p.L(450, 400);
      p.L(300, 50);
      p.Z();
      p.stroke();
    }
  },
  // ⮤ LEFT TRIANGLE ARROWHEAD (white)
  [0x2ba4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(450, 450);
      p.L(100, 250);
      p.L(450, 50);
      p.Z();
      p.stroke();
    }
  }
};
