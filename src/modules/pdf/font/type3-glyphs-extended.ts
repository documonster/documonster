/**
 * Extended Type3 glyph definitions — programmatically generated blocks
 * and additional hand-crafted symbols.
 *
 * This file supplements type3-glyphs.ts with:
 *  - Box Drawing        (U+2500–U+257F)  128 chars — algorithmic
 *  - Block Elements     (U+2580–U+259F)   32 chars — algorithmic
 *  - Braille Patterns   (U+2800–U+28FF)  256 chars — algorithmic
 *  - Letterlike Symbols  (U+2100–U+214F)  hand-crafted
 *  - Number Forms        (U+2150–U+218F)  hand-crafted
 *  - Enclosed Alphanumerics (U+2460–U+24FF) programmatic
 *  - General Punctuation extras (U+2000–U+206F)
 *  - Additional Arrows, Math, Dingbats, Misc Symbols, Currency, Technical
 */

import type { GlyphDef, GlyphPen } from "@pdf/font/type3-glyphs";

const W = 600;

// =============================================================================
// Box Drawing — FULL BLOCK (U+2500–U+257F)  128 characters
// =============================================================================
//
// Each character is a combination of horizontal/vertical line segments
// through the cell centre (300, 250), with light (40) or heavy (80) weight.
// The naming pattern: bits encode which segments are present and their weight.

function boxDraw(left: number, right: number, up: number, down: number): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      if (left) {
        p.lineWidth(left);
        p.M(0, 250);
        p.L(300, 250);
        p.stroke();
      }
      if (right) {
        p.lineWidth(right);
        p.M(300, 250);
        p.L(600, 250);
        p.stroke();
      }
      if (up) {
        p.lineWidth(up);
        p.M(300, 250);
        p.L(300, 500);
        p.stroke();
      }
      if (down) {
        p.lineWidth(down);
        p.M(300, 250);
        p.L(300, 0);
        p.stroke();
      }
    }
  };
}

function boxDouble(leftD: boolean, rightD: boolean, upD: boolean, downD: boolean): GlyphDef {
  const s = 30; // line weight
  const g = 30; // gap between double lines
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(s);
      if (leftD) {
        p.M(0, 250 - g);
        p.L(300, 250 - g);
        p.stroke();
        p.M(0, 250 + g);
        p.L(300, 250 + g);
        p.stroke();
      }
      if (rightD) {
        p.M(300, 250 - g);
        p.L(600, 250 - g);
        p.stroke();
        p.M(300, 250 + g);
        p.L(600, 250 + g);
        p.stroke();
      }
      if (upD) {
        p.M(300 - g, 250);
        p.L(300 - g, 500);
        p.stroke();
        p.M(300 + g, 250);
        p.L(300 + g, 500);
        p.stroke();
      }
      if (downD) {
        p.M(300 - g, 250);
        p.L(300 - g, 0);
        p.stroke();
        p.M(300 + g, 250);
        p.L(300 + g, 0);
        p.stroke();
      }
    }
  };
}

// Light=40, Heavy=80
const L = 40;
const H = 80;

export const BOX_FULL: Record<number, GlyphDef> = {
  // Light lines
  [0x2500]: boxDraw(L, L, 0, 0), // ─
  [0x2501]: boxDraw(H, H, 0, 0), // ━
  [0x2502]: boxDraw(0, 0, L, L), // │
  [0x2503]: boxDraw(0, 0, H, H), // ┃
  [0x2504]: boxDraw(L, L, 0, 0), // ┄ (triple dash — approx as solid)
  [0x2505]: boxDraw(H, H, 0, 0), // ┅
  [0x2506]: boxDraw(0, 0, L, L), // ┆
  [0x2507]: boxDraw(0, 0, H, H), // ┇
  [0x2508]: boxDraw(L, L, 0, 0), // ┈ (quadruple dash)
  [0x2509]: boxDraw(H, H, 0, 0), // ┉
  [0x250a]: boxDraw(0, 0, L, L), // ┊
  [0x250b]: boxDraw(0, 0, H, H), // ┋
  // Corners: light
  [0x250c]: boxDraw(0, L, 0, L), // ┌
  [0x250d]: boxDraw(0, H, 0, L), // ┍
  [0x250e]: boxDraw(0, L, 0, H), // ┎
  [0x250f]: boxDraw(0, H, 0, H), // ┏
  [0x2510]: boxDraw(L, 0, 0, L), // ┐
  [0x2511]: boxDraw(H, 0, 0, L), // ┑
  [0x2512]: boxDraw(L, 0, 0, H), // ┒
  [0x2513]: boxDraw(H, 0, 0, H), // ┓
  [0x2514]: boxDraw(0, L, L, 0), // └
  [0x2515]: boxDraw(0, H, L, 0), // ┕
  [0x2516]: boxDraw(0, L, H, 0), // ┖
  [0x2517]: boxDraw(0, H, H, 0), // ┗
  [0x2518]: boxDraw(L, 0, L, 0), // ┘
  [0x2519]: boxDraw(H, 0, L, 0), // ┙
  [0x251a]: boxDraw(L, 0, H, 0), // ┚
  [0x251b]: boxDraw(H, 0, H, 0), // ┛
  // T-junctions
  [0x251c]: boxDraw(0, L, L, L), // ├
  [0x251d]: boxDraw(0, H, L, L), // ┝
  [0x251e]: boxDraw(0, L, H, L), // ┞
  [0x251f]: boxDraw(0, L, L, H), // ┟
  [0x2520]: boxDraw(0, L, H, H), // ┠
  [0x2521]: boxDraw(0, H, H, L), // ┡
  [0x2522]: boxDraw(0, H, L, H), // ┢
  [0x2523]: boxDraw(0, H, H, H), // ┣
  [0x2524]: boxDraw(L, 0, L, L), // ┤
  [0x2525]: boxDraw(H, 0, L, L), // ┥
  [0x2526]: boxDraw(L, 0, H, L), // ┦
  [0x2527]: boxDraw(L, 0, L, H), // ┧
  [0x2528]: boxDraw(L, 0, H, H), // ┨
  [0x2529]: boxDraw(H, 0, H, L), // ┩
  [0x252a]: boxDraw(H, 0, L, H), // ┪
  [0x252b]: boxDraw(H, 0, H, H), // ┫
  [0x252c]: boxDraw(L, L, 0, L), // ┬
  [0x252d]: boxDraw(H, L, 0, L), // ┭
  [0x252e]: boxDraw(L, H, 0, L), // ┮
  [0x252f]: boxDraw(H, H, 0, L), // ┯
  [0x2530]: boxDraw(L, L, 0, H), // ┰
  [0x2531]: boxDraw(H, L, 0, H), // ┱
  [0x2532]: boxDraw(L, H, 0, H), // ┲
  [0x2533]: boxDraw(H, H, 0, H), // ┳
  [0x2534]: boxDraw(L, L, L, 0), // ┴
  [0x2535]: boxDraw(H, L, L, 0), // ┵
  [0x2536]: boxDraw(L, H, L, 0), // ┶
  [0x2537]: boxDraw(H, H, L, 0), // ┷
  [0x2538]: boxDraw(L, L, H, 0), // ┸
  [0x2539]: boxDraw(H, L, H, 0), // ┹
  [0x253a]: boxDraw(L, H, H, 0), // ┺
  [0x253b]: boxDraw(H, H, H, 0), // ┻
  // Crosses
  [0x253c]: boxDraw(L, L, L, L), // ┼
  [0x253d]: boxDraw(H, L, L, L), // ┽
  [0x253e]: boxDraw(L, H, L, L), // ┾
  [0x253f]: boxDraw(H, H, L, L), // ┿
  [0x2540]: boxDraw(L, L, H, L), // ╀
  [0x2541]: boxDraw(L, L, L, H), // ╁
  [0x2542]: boxDraw(L, L, H, H), // ╂
  [0x2543]: boxDraw(H, L, H, L), // ╃
  [0x2544]: boxDraw(L, H, H, L), // ╄
  [0x2545]: boxDraw(H, L, L, H), // ╅
  [0x2546]: boxDraw(L, H, L, H), // ╆
  [0x2547]: boxDraw(H, H, H, L), // ╇
  [0x2548]: boxDraw(H, H, L, H), // ╈
  [0x2549]: boxDraw(H, L, H, H), // ╉
  [0x254a]: boxDraw(L, H, H, H), // ╊
  [0x254b]: boxDraw(H, H, H, H), // ╋
  // Double lines
  [0x2550]: boxDouble(true, true, false, false), // ═
  [0x2551]: boxDouble(false, false, true, true), // ║
  [0x2552]: boxDouble(false, true, false, true), // ╒ (approx)
  [0x2553]: boxDouble(false, true, false, true), // ╓
  [0x2554]: boxDouble(false, true, false, true), // ╔
  [0x2555]: boxDouble(true, false, false, true), // ╕
  [0x2556]: boxDouble(true, false, false, true), // ╖
  [0x2557]: boxDouble(true, false, false, true), // ╗
  [0x2558]: boxDouble(false, true, true, false), // ╘
  [0x2559]: boxDouble(false, true, true, false), // ╙
  [0x255a]: boxDouble(false, true, true, false), // ╚
  [0x255b]: boxDouble(true, false, true, false), // ╛
  [0x255c]: boxDouble(true, false, true, false), // ╜
  [0x255d]: boxDouble(true, false, true, false), // ╝
  [0x255e]: boxDouble(false, true, true, true), // ╞
  [0x255f]: boxDouble(false, true, true, true), // ╟
  [0x2560]: boxDouble(false, true, true, true), // ╠
  [0x2561]: boxDouble(true, false, true, true), // ╡
  [0x2562]: boxDouble(true, false, true, true), // ╢
  [0x2563]: boxDouble(true, false, true, true), // ╣
  [0x2564]: boxDouble(true, true, false, true), // ╤
  [0x2565]: boxDouble(true, true, false, true), // ╥
  [0x2566]: boxDouble(true, true, false, true), // ╦
  [0x2567]: boxDouble(true, true, true, false), // ╧
  [0x2568]: boxDouble(true, true, true, false), // ╨
  [0x2569]: boxDouble(true, true, true, false), // ╩
  [0x256a]: boxDouble(true, true, true, true), // ╪
  [0x256b]: boxDouble(true, true, true, true), // ╫
  [0x256c]: boxDouble(true, true, true, true), // ╬
  // Rounded corners (light)
  [0x256d]: boxDraw(0, L, 0, L), // ╭
  [0x256e]: boxDraw(L, 0, 0, L), // ╮
  [0x256f]: boxDraw(L, 0, L, 0), // ╯
  [0x2570]: boxDraw(0, L, L, 0), // ╰
  // Diagonals
  [0x2571]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(L);
      p.M(0, 0);
      p.L(600, 500);
      p.stroke();
    }
  }, // ╱
  [0x2572]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(L);
      p.M(0, 500);
      p.L(600, 0);
      p.stroke();
    }
  }, // ╲
  [0x2573]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(L);
      p.M(0, 0);
      p.L(600, 500);
      p.stroke();
      p.M(0, 500);
      p.L(600, 0);
      p.stroke();
    }
  }, // ╳
  // Half lines
  [0x2574]: boxDraw(L, 0, 0, 0), // ╴
  [0x2575]: boxDraw(0, 0, L, 0), // ╵
  [0x2576]: boxDraw(0, L, 0, 0), // ╶
  [0x2577]: boxDraw(0, 0, 0, L), // ╷
  [0x2578]: boxDraw(H, 0, 0, 0), // ╸
  [0x2579]: boxDraw(0, 0, H, 0), // ╹
  [0x257a]: boxDraw(0, H, 0, 0), // ╺
  [0x257b]: boxDraw(0, 0, 0, H), // ╻
  [0x257c]: boxDraw(L, H, 0, 0), // ╼
  [0x257d]: boxDraw(0, 0, L, H), // ╽
  [0x257e]: boxDraw(H, L, 0, 0), // ╾
  [0x257f]: boxDraw(0, 0, H, L) // ╿
};

// =============================================================================
// Block Elements — FULL BLOCK (U+2580–U+259F)  32 characters
// =============================================================================

function blockElement(x: number, y: number, w: number, h: number): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(x, y, w, h);
      p.fill();
    }
  };
}

export const BLOCK_FULL: Record<number, GlyphDef> = {
  [0x2580]: blockElement(0, 250, 600, 250), // ▀ UPPER HALF
  [0x2581]: blockElement(0, 0, 600, 62), // ▁ LOWER ONE EIGHTH
  [0x2582]: blockElement(0, 0, 600, 125), // ▂ LOWER ONE QUARTER
  [0x2583]: blockElement(0, 0, 600, 187), // ▃ LOWER THREE EIGHTHS
  [0x2584]: blockElement(0, 0, 600, 250), // ▄ LOWER HALF
  [0x2585]: blockElement(0, 0, 600, 312), // ▅ LOWER FIVE EIGHTHS
  [0x2586]: blockElement(0, 0, 600, 375), // ▆ LOWER THREE QUARTERS
  [0x2587]: blockElement(0, 0, 600, 437), // ▇ LOWER SEVEN EIGHTHS
  [0x2588]: blockElement(0, 0, 600, 500), // █ FULL BLOCK
  [0x2589]: blockElement(0, 0, 525, 500), // ▉ LEFT SEVEN EIGHTHS
  [0x258a]: blockElement(0, 0, 450, 500), // ▊ LEFT THREE QUARTERS
  [0x258b]: blockElement(0, 0, 375, 500), // ▋ LEFT FIVE EIGHTHS
  [0x258c]: blockElement(0, 0, 300, 500), // ▌ LEFT HALF
  [0x258d]: blockElement(0, 0, 225, 500), // ▍ LEFT THREE EIGHTHS
  [0x258e]: blockElement(0, 0, 150, 500), // ▎ LEFT ONE QUARTER
  [0x258f]: blockElement(0, 0, 75, 500), // ▏ LEFT ONE EIGHTH
  [0x2590]: blockElement(300, 0, 300, 500), // ▐ RIGHT HALF
  [0x2591]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(10);
      for (let y = 10; y < 500; y += 40) {
        for (let x = 10; x < 600; x += 40) {
          p.rect(x, y, 5, 5);
        }
      }
      p.fill();
    }
  }, // ░
  [0x2592]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(10);
      for (let y = 5; y < 500; y += 20) {
        for (let x = 5; x < 600; x += 20) {
          p.rect(x, y, 8, 8);
        }
      }
      p.fill();
    }
  }, // ▒
  [0x2593]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(10);
      for (let y = 3; y < 500; y += 12) {
        for (let x = 3; x < 600; x += 12) {
          p.rect(x, y, 10, 10);
        }
      }
      p.fill();
    }
  }, // ▓
  [0x2594]: blockElement(0, 437, 600, 63), // ▔ UPPER ONE EIGHTH
  [0x2595]: blockElement(525, 0, 75, 500), // ▕ RIGHT ONE EIGHTH
  [0x2596]: blockElement(0, 0, 300, 250), // ▖ QUADRANT LOWER LEFT
  [0x2597]: blockElement(300, 0, 300, 250), // ▗ QUADRANT LOWER RIGHT
  [0x2598]: blockElement(0, 250, 300, 250), // ▘ QUADRANT UPPER LEFT
  [0x2599]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(0, 0, 300, 250);
      p.fill();
      p.rect(0, 250, 300, 250);
      p.fill();
      p.rect(300, 0, 300, 250);
      p.fill();
    }
  }, // ▙
  [0x259a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(0, 250, 300, 250);
      p.fill();
      p.rect(300, 0, 300, 250);
      p.fill();
    }
  }, // ▚
  [0x259b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(0, 0, 300, 250);
      p.fill();
      p.rect(0, 250, 600, 250);
      p.fill();
    }
  }, // ▛
  [0x259c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(300, 0, 300, 250);
      p.fill();
      p.rect(0, 250, 600, 250);
      p.fill();
    }
  }, // ▜
  [0x259d]: blockElement(300, 250, 300, 250), // ▝ QUADRANT UPPER RIGHT
  [0x259e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(300, 250, 300, 250);
      p.fill();
      p.rect(0, 0, 300, 250);
      p.fill();
    }
  }, // ▞
  [0x259f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(300, 0, 300, 500);
      p.fill();
      p.rect(0, 0, 300, 250);
      p.fill();
    }
  } // ▟
};

// =============================================================================
// Braille Patterns — FULL BLOCK (U+2800–U+28FF)  256 characters
// =============================================================================
// Each Braille pattern is a 2×4 grid of dots. The code point encodes which
// dots are raised: bit 0=dot1(top-left), bit 1=dot2(mid-left), bit 2=dot3(bot-left),
// bit 3=dot4(top-right), bit 4=dot5(mid-right), bit 5=dot6(bot-right),
// bit 6=dot7(lower-left), bit 7=dot8(lower-right).

function brailleGlyph(pattern: number): GlyphDef {
  // Dot positions in the 1000-unit glyph space
  const dotR = 40;
  const colX = [200, 400]; // left, right
  const rowY = [400, 300, 200, 100]; // top to bottom (PDF y increases upward)
  // Bit→dot mapping: bits 0-2 = left col rows 0-2, bits 3-5 = right col rows 0-2,
  // bit 6 = left col row 3, bit 7 = right col row 3
  const dots: Array<[number, number]> = [
    [0, 0],
    [0, 1],
    [0, 2], // bits 0,1,2 → left col, rows 0,1,2
    [1, 0],
    [1, 1],
    [1, 2], // bits 3,4,5 → right col, rows 0,1,2
    [0, 3],
    [1, 3] // bits 6,7 → left/right col, row 3
  ];

  return {
    width: W,
    draw: (p: GlyphPen) => {
      for (let bit = 0; bit < 8; bit++) {
        if (pattern & (1 << bit)) {
          const [col, row] = dots[bit];
          p.circle(colX[col], rowY[row], dotR);
          p.fill();
        }
      }
    }
  };
}

export const BRAILLE: Record<number, GlyphDef> = {};
for (let i = 0; i < 256; i++) {
  BRAILLE[0x2800 + i] = brailleGlyph(i);
}

// =============================================================================
// Letterlike Symbols (U+2100–U+214F)
// =============================================================================

export const LETTERLIKE: Record<number, GlyphDef> = {
  // ℃ DEGREE CELSIUS
  [0x2103]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(120, 450, 60);
      p.stroke();
      p.M(550, 450);
      p.C(350, 450, 250, 350, 250, 250);
      p.C(250, 150, 350, 50, 550, 50);
      p.stroke();
    }
  },
  // ℅ CARE OF
  [0x2105]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 350);
      p.C(80, 350, 50, 400, 50, 430);
      p.C(50, 470, 100, 500, 150, 500);
      p.C(200, 500, 250, 470, 250, 430);
      p.C(250, 400, 200, 350, 150, 350);
      p.stroke();
      p.M(450, 0);
      p.L(150, 500);
      p.stroke();
      p.M(450, 150);
      p.C(380, 150, 350, 100, 350, 70);
      p.C(350, 30, 400, 0, 450, 0);
      p.C(500, 0, 550, 30, 550, 70);
      p.C(550, 100, 500, 150, 450, 150);
      p.stroke();
    }
  },
  // ℉ DEGREE FAHRENHEIT
  [0x2109]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(120, 450, 60);
      p.stroke();
      p.M(280, 0);
      p.L(280, 500);
      p.stroke();
      p.M(280, 500);
      p.L(550, 500);
      p.stroke();
      p.M(280, 280);
      p.L(480, 280);
      p.stroke();
    }
  },
  // ℓ SCRIPT SMALL L
  [0x2113]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(350, 500);
      p.C(200, 500, 200, 400, 250, 200);
      p.C(280, 100, 200, 0, 150, 50);
      p.stroke();
    }
  },
  // № NUMERO SIGN
  [0x2116]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 0);
      p.L(80, 500);
      p.stroke();
      p.M(80, 500);
      p.L(350, 0);
      p.stroke();
      p.M(350, 0);
      p.L(350, 500);
      p.stroke();
      p.lineWidth(35);
      p.circle(530, 70, 70);
      p.stroke();
      p.M(380, 170);
      p.L(680, 170);
      p.stroke();
    }
  },
  // ℠ SERVICE MARK
  [0x2120]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(150, 400);
      p.C(150, 500, 350, 500, 350, 450);
      p.C(350, 400, 150, 350, 150, 300);
      p.C(150, 250, 350, 250, 350, 350);
      p.stroke();
      p.M(400, 500);
      p.L(400, 250);
      p.L(475, 400);
      p.L(550, 250);
      p.L(550, 500);
      p.stroke();
    }
  },
  // ™ TRADE MARK SIGN
  [0x2122]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(50, 500);
      p.L(200, 500);
      p.stroke();
      p.M(125, 500);
      p.L(125, 300);
      p.stroke();
      p.M(250, 500);
      p.L(250, 300);
      p.L(350, 450);
      p.L(450, 300);
      p.L(450, 500);
      p.stroke();
    }
  },
  // Ω OHM SIGN
  [0x2126]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 50);
      p.L(200, 50);
      p.L(200, 150);
      p.C(120, 200, 80, 300, 80, 350);
      p.C(80, 450, 180, 520, 300, 520);
      p.C(420, 520, 520, 450, 520, 350);
      p.C(520, 300, 480, 200, 400, 150);
      p.L(400, 50);
      p.L(500, 50);
      p.stroke();
    }
  },
  // ℮ ESTIMATED SIGN
  [0x212e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(480, 220);
      p.L(150, 220);
      p.C(150, 400, 250, 500, 350, 500);
      p.C(450, 500, 530, 400, 500, 300);
      p.C(470, 100, 400, 0, 250, 0);
      p.C(150, 0, 100, 80, 100, 150);
      p.stroke();
    }
  },
  // ℹ INFORMATION SOURCE
  [0x2139]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 450, 40);
      p.fill();
      p.lineWidth(55);
      p.M(300, 350);
      p.L(300, 50);
      p.stroke();
    }
  },
  // ⅟ FRACTION NUMERATOR ONE
  [0x215f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 500);
      p.L(200, 500);
      p.L(200, 350);
      p.stroke();
      p.M(450, 500);
      p.L(150, 0);
      p.stroke();
    }
  }
};

// =============================================================================
// Number Forms (U+2150–U+218F) — Vulgar Fractions
// =============================================================================

function fraction(num: string, denom: string): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      // We can't draw actual text in Type3 glyphs, so draw the fraction bar
      // and approximate numerator/denominator with simple shapes
      p.lineWidth(35);
      p.M(450, 500);
      p.L(150, 0);
      p.stroke();
      // Small circles as digit placeholders
      const numCount = num.length;
      const denCount = denom.length;
      for (let i = 0; i < numCount; i++) {
        p.circle(150 + i * 80, 420, 30);
        p.fill();
      }
      for (let i = 0; i < denCount; i++) {
        p.circle(350 + i * 80, 80, 30);
        p.fill();
      }
    }
  };
}

export const NUMBER_FORMS: Record<number, GlyphDef> = {
  [0x2150]: fraction("1", "7"), // ⅐
  [0x2151]: fraction("1", "9"), // ⅑
  [0x2152]: fraction("1", "10"), // ⅒
  [0x2153]: fraction("1", "3"), // ⅓
  [0x2154]: fraction("2", "3"), // ⅔
  [0x2155]: fraction("1", "5"), // ⅕
  [0x2156]: fraction("2", "5"), // ⅖
  [0x2157]: fraction("3", "5"), // ⅗
  [0x2158]: fraction("4", "5"), // ⅘
  [0x2159]: fraction("1", "6"), // ⅙
  [0x215a]: fraction("5", "6"), // ⅚
  [0x215b]: fraction("1", "8"), // ⅛
  [0x215c]: fraction("3", "8"), // ⅜
  [0x215d]: fraction("5", "8"), // ⅝
  [0x215e]: fraction("7", "8"), // ⅞
  // Roman numerals as simple stroked letters
  [0x2160]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
    }
  }, // Ⅰ
  [0x2161]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(200, 0);
      p.L(200, 500);
      p.stroke();
      p.M(400, 0);
      p.L(400, 500);
      p.stroke();
    }
  }, // Ⅱ
  [0x2162]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 0);
      p.L(120, 500);
      p.stroke();
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(480, 0);
      p.L(480, 500);
      p.stroke();
    }
  }, // Ⅲ
  [0x2163]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(180, 500);
      p.L(300, 0);
      p.L(420, 500);
      p.stroke();
    }
  }, // Ⅳ (V shape inverted — approx)
  [0x2164]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(150, 500);
      p.L(300, 0);
      p.L(450, 500);
      p.stroke();
    }
  } // Ⅴ
};

// =============================================================================
// Enclosed Alphanumerics — circled numbers ①–⑳ (U+2460–U+2473)
// =============================================================================

export const ENCLOSED: Record<number, GlyphDef> = {};
// ① - ⑳ : circled digits (stroke circle only — actual digit not possible w/o font)
for (let i = 0; i < 20; i++) {
  ENCLOSED[0x2460 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 220);
      p.stroke();
    }
  };
}
// ⑴-⑿ Parenthesized digits
for (let i = 0; i < 12; i++) {
  ENCLOSED[0x2474 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(180, 480);
      p.C(120, 480, 80, 380, 80, 250);
      p.C(80, 120, 120, 20, 180, 20);
      p.stroke();
      p.M(420, 20);
      p.C(480, 20, 520, 120, 520, 250);
      p.C(520, 380, 480, 480, 420, 480);
      p.stroke();
    }
  };
}
// Ⓐ-Ⓩ Circled Latin Capital (U+24B6–U+24CF)
for (let i = 0; i < 26; i++) {
  ENCLOSED[0x24b6 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 220);
      p.stroke();
    }
  };
}
// ⓐ-ⓩ Circled Latin Small (U+24D0–U+24E9)
for (let i = 0; i < 26; i++) {
  ENCLOSED[0x24d0 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.circle(300, 250, 200);
      p.stroke();
    }
  };
}
// ⓪ Circled zero (U+24EA)
ENCLOSED[0x24ea] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(35);
    p.circle(300, 250, 220);
    p.stroke();
  }
};

// =============================================================================
// Additional General Punctuation (U+2000–U+206F)
// =============================================================================

export const PUNCT_EXT: Record<number, GlyphDef> = {
  // ‐ HYPHEN (U+2010)
  [0x2010]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(150, 250);
      p.L(450, 250);
      p.stroke();
    }
  },
  // ‑ NON-BREAKING HYPHEN (U+2011)
  [0x2011]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(150, 250);
      p.L(450, 250);
      p.stroke();
    }
  },
  // ‒ FIGURE DASH (U+2012)
  [0x2012]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
    }
  },
  // – EN DASH (U+2013)
  [0x2013]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(80, 250);
      p.L(520, 250);
      p.stroke();
    }
  },
  // — EM DASH (U+2014)
  [0x2014]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(0, 250);
      p.L(700, 250);
      p.stroke();
    }
  },
  // ‖ DOUBLE VERTICAL LINE (U+2016)
  [0x2016]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(230, 0);
      p.L(230, 500);
      p.stroke();
      p.M(370, 0);
      p.L(370, 500);
      p.stroke();
    }
  },
  // ‗ DOUBLE LOW LINE (U+2017)
  [0x2017]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.M(50, 30);
      p.L(550, 30);
      p.stroke();
      p.M(50, 80);
      p.L(550, 80);
      p.stroke();
    }
  },
  // ' LEFT SINGLE QUOTATION MARK (U+2018)
  [0x2018]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.circle(150, 450, 40);
      p.fill();
      p.lineWidth(30);
      p.M(150, 410);
      p.C(120, 370, 100, 350, 80, 340);
      p.stroke();
    }
  },
  // ' RIGHT SINGLE QUOTATION MARK (U+2019)
  [0x2019]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.circle(150, 450, 40);
      p.fill();
      p.lineWidth(30);
      p.M(150, 410);
      p.C(180, 370, 200, 350, 220, 340);
      p.stroke();
    }
  },
  // ‟ DOUBLE HIGH-REVERSED-9 QUOTATION MARK (U+201F)
  [0x201f]: {
    width: 400,
    draw: (p: GlyphPen) => {
      p.circle(120, 450, 35);
      p.fill();
      p.circle(280, 450, 35);
      p.fill();
    }
  },
  // † DAGGER (U+2020)
  [0x2020]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 500);
      p.L(300, 0);
      p.stroke();
      p.M(150, 380);
      p.L(450, 380);
      p.stroke();
    }
  },
  // ‡ DOUBLE DAGGER (U+2021)
  [0x2021]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 500);
      p.L(300, 0);
      p.stroke();
      p.M(150, 380);
      p.L(450, 380);
      p.stroke();
      p.M(150, 180);
      p.L(450, 180);
      p.stroke();
    }
  },
  // ‣ TRIANGULAR BULLET (U+2023)  — already in main file
  // ‥ TWO DOT LEADER (U+2025)
  [0x2025]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(200, 50, 35);
      p.fill();
      p.circle(400, 50, 35);
      p.fill();
    }
  },
  // ‧ HYPHENATION POINT (U+2027)
  [0x2027]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.circle(150, 250, 35);
      p.fill();
    }
  },
  // ‰ PER MILLE SIGN (U+2030)
  [0x2030]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(550, 0);
      p.L(150, 500);
      p.stroke();
      p.circle(180, 400, 70);
      p.stroke();
      p.circle(400, 100, 60);
      p.stroke();
      p.circle(560, 100, 60);
      p.stroke();
    }
  },
  // ‱ PER TEN THOUSAND SIGN (U+2031)
  [0x2031]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(550, 0);
      p.L(150, 500);
      p.stroke();
      p.circle(180, 400, 60);
      p.stroke();
      p.circle(350, 100, 50);
      p.stroke();
      p.circle(470, 100, 50);
      p.stroke();
      p.circle(590, 100, 50);
      p.stroke();
    }
  },
  // ′ PRIME (U+2032)
  [0x2032]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(180, 500);
      p.L(140, 350);
      p.stroke();
    }
  },
  // ″ DOUBLE PRIME (U+2033)
  [0x2033]: {
    width: 400,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 500);
      p.L(110, 350);
      p.stroke();
      p.M(300, 500);
      p.L(260, 350);
      p.stroke();
    }
  },
  // ‹ SINGLE LEFT-POINTING ANGLE QUOTATION MARK (U+2039)
  [0x2039]: {
    width: 350,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(250, 400);
      p.L(100, 250);
      p.L(250, 100);
      p.stroke();
    }
  },
  // › SINGLE RIGHT-POINTING ANGLE QUOTATION MARK (U+203A)
  [0x203a]: {
    width: 350,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 400);
      p.L(250, 250);
      p.L(100, 100);
      p.stroke();
    }
  },
  // ‼ DOUBLE EXCLAMATION MARK (U+203C)
  [0x203c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(200, 500);
      p.L(200, 130);
      p.stroke();
      p.circle(200, 40, 30);
      p.fill();
      p.M(400, 500);
      p.L(400, 130);
      p.stroke();
      p.circle(400, 40, 30);
      p.fill();
    }
  },
  // ⁄ FRACTION SLASH (U+2044)
  [0x2044]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(450, 500);
      p.L(150, 0);
      p.stroke();
    }
  },
  // ⁏ REVERSED SEMICOLON (U+204F)
  [0x204f]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.circle(150, 350, 35);
      p.fill();
      p.circle(150, 130, 35);
      p.fill();
      p.lineWidth(30);
      p.M(150, 95);
      p.C(180, 50, 200, 20, 220, 0);
      p.stroke();
    }
  }
};

// =============================================================================
// Additional Arrows (U+2190–U+21FF)
// =============================================================================

export const ARROWS_EXT: Record<number, GlyphDef> = {
  // ↚ LEFTWARDS ARROW WITH STROKE
  [0x219a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(250, 380);
      p.L(250, 120);
      p.Z();
      p.fill();
      p.lineWidth(50);
      p.M(350, 120);
      p.L(250, 380);
      p.stroke();
    }
  },
  // ↛ RIGHTWARDS ARROW WITH STROKE
  [0x219b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(350, 380);
      p.L(350, 120);
      p.Z();
      p.fill();
      p.lineWidth(50);
      p.M(250, 120);
      p.L(350, 380);
      p.stroke();
    }
  },
  // ↝ RIGHTWARDS WAVE ARROW
  [0x219d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.C(200, 350, 300, 150, 400, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(400, 330);
      p.L(400, 170);
      p.Z();
      p.fill();
    }
  },
  // ↞ LEFTWARDS TWO HEADED ARROW
  [0x219e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(500, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(220, 370);
      p.L(220, 130);
      p.Z();
      p.fill();
      p.M(200, 250);
      p.L(320, 370);
      p.L(320, 130);
      p.Z();
      p.fill();
    }
  },
  // ↟ UPWARDS TWO HEADED ARROW
  [0x219f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 50);
      p.L(300, 470);
      p.stroke();
      p.M(300, 470);
      p.L(180, 350);
      p.L(420, 350);
      p.Z();
      p.fill();
      p.M(300, 370);
      p.L(180, 250);
      p.L(420, 250);
      p.Z();
      p.fill();
    }
  },
  // ↠ RIGHTWARDS TWO HEADED ARROW
  [0x21a0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 250);
      p.L(500, 250);
      p.stroke();
      p.M(500, 250);
      p.L(380, 370);
      p.L(380, 130);
      p.Z();
      p.fill();
      p.M(400, 250);
      p.L(280, 370);
      p.L(280, 130);
      p.Z();
      p.fill();
    }
  },
  // ↡ DOWNWARDS TWO HEADED ARROW
  [0x21a1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 450);
      p.L(300, 30);
      p.stroke();
      p.M(300, 30);
      p.L(180, 150);
      p.L(420, 150);
      p.Z();
      p.fill();
      p.M(300, 130);
      p.L(180, 250);
      p.L(420, 250);
      p.Z();
      p.fill();
    }
  },
  // ↤ LEFTWARDS ARROW FROM BAR
  [0x21a4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(480, 250);
      p.L(130, 250);
      p.stroke();
      p.M(130, 250);
      p.L(260, 370);
      p.L(260, 130);
      p.Z();
      p.fill();
      p.M(500, 100);
      p.L(500, 400);
      p.stroke();
    }
  },
  // ↥ UPWARDS ARROW FROM BAR
  [0x21a5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 30);
      p.L(300, 420);
      p.stroke();
      p.M(300, 420);
      p.L(180, 290);
      p.L(420, 290);
      p.Z();
      p.fill();
      p.M(150, 30);
      p.L(450, 30);
      p.stroke();
    }
  },
  // ↦ RIGHTWARDS ARROW FROM BAR
  [0x21a6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 250);
      p.L(470, 250);
      p.stroke();
      p.M(470, 250);
      p.L(340, 370);
      p.L(340, 130);
      p.Z();
      p.fill();
      p.M(100, 100);
      p.L(100, 400);
      p.stroke();
    }
  },
  // ↧ DOWNWARDS ARROW FROM BAR
  [0x21a7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 470);
      p.L(300, 80);
      p.stroke();
      p.M(300, 80);
      p.L(180, 210);
      p.L(420, 210);
      p.Z();
      p.fill();
      p.M(150, 470);
      p.L(450, 470);
      p.stroke();
    }
  },
  // ↰ UPWARDS ARROW WITH TIP LEFTWARDS
  [0x21b0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 50);
      p.L(400, 350);
      p.L(150, 350);
      p.stroke();
      p.M(400, 450);
      p.L(280, 330);
      p.L(520, 330);
      p.Z();
      p.fill();
    }
  },
  // ↱ UPWARDS ARROW WITH TIP RIGHTWARDS
  [0x21b1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 50);
      p.L(200, 350);
      p.L(450, 350);
      p.stroke();
      p.M(200, 450);
      p.L(80, 330);
      p.L(320, 330);
      p.Z();
      p.fill();
    }
  },
  // ↲ DOWNWARDS ARROW WITH TIP LEFTWARDS
  [0x21b2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(400, 450);
      p.L(400, 150);
      p.L(150, 150);
      p.stroke();
      p.M(400, 50);
      p.L(280, 170);
      p.L(520, 170);
      p.Z();
      p.fill();
    }
  },
  // ↳ DOWNWARDS ARROW WITH TIP RIGHTWARDS
  [0x21b3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(200, 450);
      p.L(200, 150);
      p.L(450, 150);
      p.stroke();
      p.M(200, 50);
      p.L(80, 170);
      p.L(320, 170);
      p.Z();
      p.fill();
    }
  },
  // ↵ DOWNWARDS ARROW WITH CORNER LEFTWARDS (return key symbol)
  [0x21b5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(450, 450);
      p.L(450, 150);
      p.L(150, 150);
      p.stroke();
      p.M(150, 150);
      p.L(280, 260);
      p.L(280, 40);
      p.Z();
      p.fill();
    }
  },
  // ↺ ANTICLOCKWISE OPEN CIRCLE ARROW
  [0x21ba]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(160, 350);
      p.C(160, 450, 300, 500, 400, 450);
      p.C(500, 400, 520, 250, 450, 150);
      p.C(380, 50, 200, 50, 160, 150);
      p.stroke();
      p.M(160, 350);
      p.L(80, 250);
      p.L(250, 270);
      p.Z();
      p.fill();
    }
  },
  // ↻ CLOCKWISE OPEN CIRCLE ARROW
  [0x21bb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(440, 350);
      p.C(440, 450, 300, 500, 200, 450);
      p.C(100, 400, 80, 250, 150, 150);
      p.C(220, 50, 400, 50, 440, 150);
      p.stroke();
      p.M(440, 350);
      p.L(520, 250);
      p.L(350, 270);
      p.Z();
      p.fill();
    }
  }
};

// =============================================================================
// Additional Math Operators (U+2200–U+22FF)
// =============================================================================

export const MATH_EXT: Record<number, GlyphDef> = {
  // ∓ MINUS-OR-PLUS
  [0x2213]: {
    width: W,
    draw: (p: GlyphPen) => {
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
  // ∙ BULLET OPERATOR
  [0x2219]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 50);
      p.fill();
    }
  },
  // ∝ PROPORTIONAL TO
  [0x221d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(80, 250);
      p.C(80, 450, 250, 450, 300, 250);
      p.C(350, 50, 520, 50, 520, 250);
      p.stroke();
    }
  },
  // ∟ RIGHT ANGLE
  [0x221f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(120, 450);
      p.L(120, 50);
      p.L(480, 50);
      p.stroke();
    }
  },
  // ∣ DIVIDES
  [0x2223]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
    }
  },
  // ∤ DOES NOT DIVIDE
  [0x2224]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(200, 100);
      p.L(400, 400);
      p.stroke();
    }
  },
  // ∥ PARALLEL TO
  [0x2225]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(220, 0);
      p.L(220, 500);
      p.stroke();
      p.M(380, 0);
      p.L(380, 500);
      p.stroke();
    }
  },
  // ∦ NOT PARALLEL TO
  [0x2226]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(220, 0);
      p.L(220, 500);
      p.stroke();
      p.M(380, 0);
      p.L(380, 500);
      p.stroke();
      p.M(150, 100);
      p.L(450, 400);
      p.stroke();
    }
  },
  // ∘ RING OPERATOR
  [0x2218]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 70);
      p.stroke();
    }
  },
  // ∴ THEREFORE
  [0x2234]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 400, 40);
      p.fill();
      p.circle(180, 150, 40);
      p.fill();
      p.circle(420, 150, 40);
      p.fill();
    }
  },
  // ∵ BECAUSE
  [0x2235]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 100, 40);
      p.fill();
      p.circle(180, 350, 40);
      p.fill();
      p.circle(420, 350, 40);
      p.fill();
    }
  },
  // ∶ RATIO
  [0x2236]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.circle(150, 350, 40);
      p.fill();
      p.circle(150, 150, 40);
      p.fill();
    }
  },
  // ∷ PROPORTION
  [0x2237]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(200, 350, 35);
      p.fill();
      p.circle(200, 150, 35);
      p.fill();
      p.circle(400, 350, 35);
      p.fill();
      p.circle(400, 150, 35);
      p.fill();
    }
  },
  // ≌ ALL EQUAL TO
  [0x224c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 400);
      p.C(200, 460, 400, 460, 480, 400);
      p.stroke();
      p.M(120, 280);
      p.C(200, 340, 400, 340, 480, 280);
      p.stroke();
      p.M(120, 160);
      p.L(480, 160);
      p.stroke();
    }
  },
  // ≅ APPROXIMATELY EQUAL TO
  [0x2245]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 350);
      p.C(200, 430, 400, 430, 480, 350);
      p.stroke();
      p.M(120, 180);
      p.L(480, 180);
      p.stroke();
    }
  },
  // ≜ DELTA EQUAL TO
  [0x225c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 200);
      p.L(480, 200);
      p.stroke();
      p.M(120, 120);
      p.L(480, 120);
      p.stroke();
      p.M(300, 450);
      p.L(200, 280);
      p.L(400, 280);
      p.Z();
      p.stroke();
    }
  },
  // ≪ MUCH LESS-THAN
  [0x226a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(350, 420);
      p.L(100, 250);
      p.L(350, 80);
      p.stroke();
      p.M(500, 420);
      p.L(250, 250);
      p.L(500, 80);
      p.stroke();
    }
  },
  // ≫ MUCH GREATER-THAN
  [0x226b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 420);
      p.L(350, 250);
      p.L(100, 80);
      p.stroke();
      p.M(250, 420);
      p.L(500, 250);
      p.L(250, 80);
      p.stroke();
    }
  },
  // ⊄ NOT A SUBSET OF
  [0x2284]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(470, 450);
      p.C(200, 450, 130, 350, 130, 250);
      p.C(130, 150, 200, 50, 470, 50);
      p.stroke();
      p.M(200, 50);
      p.L(400, 450);
      p.stroke();
    }
  },
  // ⊅ NOT A SUPERSET OF
  [0x2285]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(130, 450);
      p.C(400, 450, 470, 350, 470, 250);
      p.C(470, 150, 400, 50, 130, 50);
      p.stroke();
      p.M(200, 50);
      p.L(400, 450);
      p.stroke();
    }
  },
  // ⊆ SUBSET OF OR EQUAL TO
  [0x2286]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(470, 420);
      p.C(200, 420, 130, 340, 130, 250);
      p.C(130, 160, 200, 80, 470, 80);
      p.stroke();
      p.M(130, 30);
      p.L(470, 30);
      p.stroke();
    }
  },
  // ⊇ SUPERSET OF OR EQUAL TO
  [0x2287]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(130, 420);
      p.C(400, 420, 470, 340, 470, 250);
      p.C(470, 160, 400, 80, 130, 80);
      p.stroke();
      p.M(130, 30);
      p.L(470, 30);
      p.stroke();
    }
  },
  // ⊥ UP TACK (perpendicular)
  [0x22a5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(120, 50);
      p.L(480, 50);
      p.stroke();
      p.M(300, 50);
      p.L(300, 450);
      p.stroke();
    }
  },
  // ⊿ RIGHT TRIANGLE
  [0x22bf]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 50);
      p.L(480, 50);
      p.L(120, 450);
      p.Z();
      p.stroke();
    }
  }
};

// =============================================================================
// Additional Misc Symbols (U+2600–U+26FF)
// =============================================================================

export const MISC_EXT: Record<number, GlyphDef> = {
  // ☁ CLOUD
  [0x2601]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(120, 200);
      p.C(120, 350, 200, 400, 280, 400);
      p.C(300, 470, 400, 470, 430, 400);
      p.C(500, 400, 520, 300, 480, 200);
      p.L(120, 200);
      p.stroke();
    }
  },
  // ☎ BLACK TELEPHONE — simplified
  [0x260e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(120, 100);
      p.C(120, 50, 480, 50, 480, 100);
      p.stroke();
      p.M(150, 100);
      p.C(150, 350, 200, 450, 250, 450);
      p.stroke();
      p.M(450, 100);
      p.C(450, 350, 400, 450, 350, 450);
      p.stroke();
    }
  },
  // ☮ PEACE SYMBOL
  [0x262e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(300, 470);
      p.L(300, 30);
      p.stroke();
      p.M(300, 250);
      p.L(145, 95);
      p.stroke();
      p.M(300, 250);
      p.L(455, 95);
      p.stroke();
    }
  },
  // ☯ YIN YANG
  [0x262f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.circle(300, 350, 20);
      p.fill();
      p.circle(300, 150, 20);
      p.fill();
    }
  },
  // ☹ WHITE FROWNING FACE
  [0x2639]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.circle(210, 330, 25);
      p.fill();
      p.circle(390, 330, 25);
      p.fill();
      p.M(200, 130);
      p.C(250, 80, 350, 80, 400, 130);
      p.stroke();
    }
  },
  // ☺ WHITE SMILING FACE
  [0x263a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.circle(210, 330, 25);
      p.fill();
      p.circle(390, 330, 25);
      p.fill();
      p.M(190, 170);
      p.C(220, 110, 380, 110, 410, 170);
      p.stroke();
    }
  },
  // ☻ BLACK SMILING FACE
  [0x263b]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 230);
      p.fill();
    }
  },
  // ⚐ WHITE FLAG
  [0x2690]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 0);
      p.L(150, 500);
      p.stroke();
      p.M(150, 500);
      p.L(480, 400);
      p.L(150, 300);
      p.stroke();
    }
  },
  // ⚑ BLACK FLAG
  [0x2691]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 0);
      p.L(150, 500);
      p.stroke();
      p.M(150, 500);
      p.L(480, 400);
      p.L(150, 300);
      p.Z();
      p.fill();
    }
  },
  // ⚠ WARNING SIGN
  [0x26a0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(300, 480);
      p.L(80, 50);
      p.L(520, 50);
      p.Z();
      p.stroke();
      p.lineWidth(40);
      p.M(300, 350);
      p.L(300, 180);
      p.stroke();
      p.circle(300, 110, 25);
      p.fill();
    }
  },
  // ⚡ HIGH VOLTAGE SIGN
  [0x26a1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(350, 500);
      p.L(200, 280);
      p.L(310, 280);
      p.L(250, 0);
      p.L(420, 220);
      p.L(300, 220);
      p.L(350, 500);
      p.Z();
      p.fill();
    }
  },
  // ⚪ MEDIUM WHITE CIRCLE
  [0x26aa]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 180);
      p.stroke();
    }
  },
  // ⚫ MEDIUM BLACK CIRCLE
  [0x26ab]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 180);
      p.fill();
    }
  }
};

// =============================================================================
// Additional Dingbats (U+2700–U+27BF)
// =============================================================================

export const DING_EXT: Record<number, GlyphDef> = {
  // ✁ UPPER BLADE SCISSORS
  [0x2701]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(180, 150, 80);
      p.stroke();
      p.circle(180, 350, 80);
      p.stroke();
      p.M(250, 200);
      p.L(500, 400);
      p.stroke();
      p.M(250, 300);
      p.L(500, 100);
      p.stroke();
    }
  },
  // ✂ BLACK SCISSORS
  [0x2702]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(170, 130, 80);
      p.stroke();
      p.circle(170, 370, 80);
      p.stroke();
      p.M(240, 180);
      p.L(520, 420);
      p.stroke();
      p.M(240, 320);
      p.L(520, 80);
      p.stroke();
    }
  },
  // ✆ TELEPHONE LOCATION SIGN
  [0x2706]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(200, 150);
      p.C(200, 100, 400, 100, 400, 150);
      p.stroke();
      p.M(220, 150);
      p.C(220, 350, 250, 400, 280, 400);
      p.stroke();
      p.M(380, 150);
      p.C(380, 350, 350, 400, 320, 400);
      p.stroke();
    }
  },
  // ✇ TAPE DRIVE — simplified square with circle
  [0x2707]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(80, 50, 440, 400);
      p.stroke();
      p.circle(300, 250, 120);
      p.stroke();
    }
  },
  // ✈ AIRPLANE — simplified
  [0x2708]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 480);
      p.L(260, 300);
      p.L(80, 250);
      p.L(260, 210);
      p.L(260, 80);
      p.L(220, 30);
      p.L(380, 30);
      p.L(340, 80);
      p.L(340, 210);
      p.L(520, 250);
      p.L(340, 300);
      p.Z();
      p.fill();
    }
  },
  // ✉ ENVELOPE
  [0x2709]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(80, 80, 440, 340);
      p.stroke();
      p.M(80, 420);
      p.L(300, 220);
      p.L(520, 420);
      p.stroke();
    }
  },
  // ✎ LOWER RIGHT PENCIL
  [0x270e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 50);
      p.L(480, 430);
      p.stroke();
      p.M(100, 50);
      p.L(130, 130);
      p.stroke();
    }
  },
  // ✏ PENCIL
  [0x270f]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(130, 30);
      p.L(100, 80);
      p.L(420, 470);
      p.L(480, 470);
      p.L(480, 430);
      p.L(160, 30);
      p.Z();
      p.fill();
    }
  },
  // ✐ UPPER RIGHT PENCIL
  [0x2710]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 450);
      p.L(480, 70);
      p.stroke();
      p.M(100, 450);
      p.L(130, 370);
      p.stroke();
    }
  },
  // ✝ LATIN CROSS
  [0x271d]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(55);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(140, 370);
      p.L(460, 370);
      p.stroke();
    }
  },
  // ✞ SHADOWED WHITE LATIN CROSS
  [0x271e]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(65);
      p.M(300, 0);
      p.L(300, 500);
      p.stroke();
      p.M(130, 370);
      p.L(470, 370);
      p.stroke();
    }
  },
  // ✦ BLACK FOUR POINTED STAR
  [0x2726]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(300, 480);
      p.L(230, 280);
      p.L(80, 250);
      p.L(230, 220);
      p.L(300, 20);
      p.L(370, 220);
      p.L(520, 250);
      p.L(370, 280);
      p.Z();
      p.fill();
    }
  },
  // ✧ WHITE FOUR POINTED STAR
  [0x2727]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(300, 480);
      p.L(230, 280);
      p.L(80, 250);
      p.L(230, 220);
      p.L(300, 20);
      p.L(370, 220);
      p.L(520, 250);
      p.L(370, 280);
      p.Z();
      p.stroke();
    }
  },
  // ✰ SHADOWED WHITE STAR
  [0x2730]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      const cx = 300;
      const cy = 260;
      const R = 220;
      const r = 100;
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
  // ➔ HEAVY WIDE-HEADED RIGHTWARDS ARROW
  [0x2794]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(500, 250);
      p.L(300, 430);
      p.L(300, 310);
      p.L(80, 310);
      p.L(80, 190);
      p.L(300, 190);
      p.L(300, 70);
      p.Z();
      p.fill();
    }
  },
  // ➜ HEAVY ROUND-TIPPED RIGHTWARDS ARROW
  [0x279c]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(320, 430);
      p.L(320, 310);
      p.C(200, 310, 80, 290, 80, 250);
      p.C(80, 210, 200, 190, 320, 190);
      p.L(320, 70);
      p.Z();
      p.fill();
    }
  },
  // ➡ BLACK RIGHTWARDS ARROW
  [0x27a1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 250);
      p.L(350, 420);
      p.L(350, 300);
      p.L(80, 300);
      p.L(80, 200);
      p.L(350, 200);
      p.L(350, 80);
      p.Z();
      p.fill();
    }
  }
};

// =============================================================================
// Additional Misc Technical (U+2300–U+23FF)
// =============================================================================

export const TECH_EXT: Record<number, GlyphDef> = {
  // ⌐ REVERSED NOT SIGN
  [0x2310]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(480, 350);
      p.L(120, 350);
      p.L(120, 150);
      p.stroke();
    }
  },
  // ⌕ TELEPHONE RECORDER
  [0x2315]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 250, 200);
      p.stroke();
      p.circle(300, 250, 50);
      p.fill();
    }
  },
  // ⌚ WATCH
  [0x231a]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(150, 50, 300, 400);
      p.stroke();
      p.M(300, 450);
      p.L(300, 500);
      p.stroke();
      p.M(300, 50);
      p.L(300, 0);
      p.stroke();
      p.lineWidth(30);
      p.M(300, 250);
      p.L(300, 350);
      p.stroke();
      p.M(300, 250);
      p.L(380, 250);
      p.stroke();
    }
  },
  // ⌨ KEYBOARD
  [0x2328]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(50, 80, 500, 340);
      p.stroke();
      p.lineWidth(20);
      for (let x = 100; x <= 500; x += 80) {
        p.M(x, 340);
        p.L(x + 40, 340);
        p.L(x + 40, 300);
        p.L(x, 300);
        p.Z();
        p.stroke();
      }
      p.M(150, 200);
      p.L(450, 200);
      p.L(450, 160);
      p.L(150, 160);
      p.Z();
      p.stroke();
    }
  },
  // ⎛⎜⎝ LEFT PARENTHESIS UPPER/EXTENSION/LOWER — simplified
  [0x239b]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(250, 500);
      p.C(100, 400, 100, 300, 150, 250);
      p.stroke();
    }
  },
  [0x239c]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 500);
      p.L(150, 0);
      p.stroke();
    }
  },
  [0x239d]: {
    width: 300,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 250);
      p.C(100, 200, 100, 100, 250, 0);
      p.stroke();
    }
  },
  // ⏏ EJECT SYMBOL
  [0x23cf]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(120, 150);
      p.L(480, 150);
      p.L(300, 420);
      p.Z();
      p.fill();
      p.rect(120, 50, 360, 60);
      p.fill();
    }
  },
  // ⏩ BLACK RIGHT-POINTING DOUBLE TRIANGLE
  [0x23e9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(80, 430);
      p.L(300, 250);
      p.L(80, 70);
      p.Z();
      p.fill();
      p.M(300, 430);
      p.L(520, 250);
      p.L(300, 70);
      p.Z();
      p.fill();
    }
  },
  // ⏪ BLACK LEFT-POINTING DOUBLE TRIANGLE
  [0x23ea]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(520, 430);
      p.L(300, 250);
      p.L(520, 70);
      p.Z();
      p.fill();
      p.M(300, 430);
      p.L(80, 250);
      p.L(300, 70);
      p.Z();
      p.fill();
    }
  },
  // ⏫ BLACK UP-POINTING DOUBLE TRIANGLE
  [0x23eb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(100, 230);
      p.L(300, 450);
      p.L(500, 230);
      p.Z();
      p.fill();
      p.M(100, 30);
      p.L(300, 250);
      p.L(500, 30);
      p.Z();
      p.fill();
    }
  },
  // ⏬ BLACK DOWN-POINTING DOUBLE TRIANGLE
  [0x23ec]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(100, 270);
      p.L(300, 50);
      p.L(500, 270);
      p.Z();
      p.fill();
      p.M(100, 470);
      p.L(300, 250);
      p.L(500, 470);
      p.Z();
      p.fill();
    }
  },
  // ⏭ BLACK RIGHT-POINTING DOUBLE TRIANGLE WITH VERTICAL BAR
  [0x23ed]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(60, 430);
      p.L(250, 250);
      p.L(60, 70);
      p.Z();
      p.fill();
      p.M(250, 430);
      p.L(440, 250);
      p.L(250, 70);
      p.Z();
      p.fill();
      p.rect(460, 70, 50, 360);
      p.fill();
    }
  },
  // ⏮ BLACK LEFT-POINTING DOUBLE TRIANGLE WITH VERTICAL BAR
  [0x23ee]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(540, 430);
      p.L(350, 250);
      p.L(540, 70);
      p.Z();
      p.fill();
      p.M(350, 430);
      p.L(160, 250);
      p.L(350, 70);
      p.Z();
      p.fill();
      p.rect(90, 70, 50, 360);
      p.fill();
    }
  },
  // ⏯ BLACK RIGHT-POINTING TRIANGLE WITH DOUBLE VERTICAL BAR
  [0x23ef]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(80, 430);
      p.L(330, 250);
      p.L(80, 70);
      p.Z();
      p.fill();
      p.rect(370, 70, 50, 360);
      p.fill();
      p.rect(460, 70, 50, 360);
      p.fill();
    }
  },
  // ⏱ STOPWATCH
  [0x23f1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 230, 200);
      p.stroke();
      p.M(260, 470);
      p.L(340, 470);
      p.stroke();
      p.M(300, 430);
      p.L(300, 480);
      p.stroke();
      p.M(300, 230);
      p.L(300, 350);
      p.stroke();
      p.M(300, 230);
      p.L(400, 280);
      p.stroke();
    }
  },
  // ⏲ TIMER CLOCK
  [0x23f2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.circle(300, 230, 200);
      p.stroke();
      p.M(300, 230);
      p.L(300, 370);
      p.stroke();
      p.M(300, 230);
      p.L(420, 230);
      p.stroke();
      p.M(430, 400);
      p.L(470, 440);
      p.stroke();
    }
  },
  // ⏸ DOUBLE VERTICAL BAR (pause)
  [0x23f8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(140, 60, 100, 380);
      p.fill();
      p.rect(360, 60, 100, 380);
      p.fill();
    }
  },
  // ⏹ BLACK SQUARE FOR STOP
  [0x23f9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(100, 50, 400, 400);
      p.fill();
    }
  },
  // ⏺ BLACK CIRCLE FOR RECORD
  [0x23fa]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.circle(300, 250, 200);
      p.fill();
    }
  }
};

// =============================================================================
// Additional Currency Symbols (U+20A0–U+20CF)
// =============================================================================

export const CURRENCY_EXT: Record<number, GlyphDef> = {
  // ₠ EURO-CURRENCY SIGN
  [0x20a0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(480, 430);
      p.C(350, 480, 180, 430, 150, 300);
      p.C(120, 170, 200, 50, 400, 30);
      p.stroke();
      p.M(80, 320);
      p.L(380, 320);
      p.stroke();
      p.M(80, 220);
      p.L(380, 220);
      p.stroke();
    }
  },
  // ₡ COLON SIGN
  [0x20a1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(450, 430);
      p.C(350, 480, 180, 430, 150, 250);
      p.C(150, 70, 350, 20, 450, 70);
      p.stroke();
      p.M(300, 520);
      p.L(300, -20);
      p.stroke();
    }
  },
  // ₢ CRUZEIRO SIGN
  [0x20a2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(450, 430);
      p.C(350, 480, 180, 430, 150, 250);
      p.C(150, 70, 350, 20, 450, 70);
      p.stroke();
      p.M(380, 0);
      p.L(250, 500);
      p.stroke();
    }
  },
  // ₣ FRENCH FRANC SIGN
  [0x20a3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(200, 0);
      p.L(200, 500);
      p.stroke();
      p.M(200, 500);
      p.L(480, 500);
      p.stroke();
      p.M(200, 300);
      p.L(400, 300);
      p.stroke();
      p.M(120, 200);
      p.L(350, 200);
      p.stroke();
    }
  },
  // ₤ LIRA SIGN
  [0x20a4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(200, 0);
      p.C(200, 0, 400, 0, 400, 0);
      p.stroke();
      p.M(200, 0);
      p.L(200, 500);
      p.C(200, 500, 150, 520, 120, 500);
      p.stroke();
      p.M(120, 300);
      p.L(380, 300);
      p.stroke();
      p.M(120, 180);
      p.L(380, 180);
      p.stroke();
    }
  },
  // ₥ MILL SIGN
  [0x20a5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 0);
      p.L(100, 350);
      p.L(200, 200);
      p.L(300, 350);
      p.L(300, 0);
      p.stroke();
      p.M(300, 500);
      p.L(300, -30);
      p.stroke();
    }
  },
  // ₦ NAIRA SIGN
  [0x20a6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(130, 0);
      p.L(130, 500);
      p.stroke();
      p.M(130, 500);
      p.L(470, 0);
      p.stroke();
      p.M(470, 0);
      p.L(470, 500);
      p.stroke();
      p.M(80, 350);
      p.L(520, 350);
      p.stroke();
      p.M(80, 200);
      p.L(520, 200);
      p.stroke();
    }
  },
  // ₧ PESETA SIGN
  [0x20a7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 0);
      p.L(150, 500);
      p.stroke();
      p.M(150, 500);
      p.L(350, 500);
      p.C(450, 500, 480, 400, 350, 350);
      p.L(150, 350);
      p.stroke();
      p.M(350, 0);
      p.L(500, 0);
      p.stroke();
    }
  },
  // ₨ RUPEE SIGN
  [0x20a8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 0);
      p.L(150, 500);
      p.stroke();
      p.M(150, 500);
      p.L(350, 500);
      p.C(450, 500, 480, 400, 350, 350);
      p.L(150, 350);
      p.stroke();
      p.M(280, 350);
      p.L(450, 0);
      p.stroke();
    }
  },
  // ₩ WON SIGN
  [0x20a9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(60, 500);
      p.L(180, 0);
      p.L(300, 350);
      p.L(420, 0);
      p.L(540, 500);
      p.stroke();
      p.M(80, 250);
      p.L(520, 250);
      p.stroke();
      p.M(80, 150);
      p.L(520, 150);
      p.stroke();
    }
  },
  // ₪ NEW SHEQEL SIGN
  [0x20aa]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 500);
      p.L(150, 100);
      p.C(150, 30, 250, 0, 300, 50);
      p.stroke();
      p.M(450, 0);
      p.L(450, 400);
      p.C(450, 470, 350, 500, 300, 450);
      p.stroke();
    }
  },
  // ₫ DONG SIGN
  [0x20ab]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 50);
      p.L(250, 500);
      p.stroke();
      p.M(250, 400);
      p.L(350, 400);
      p.C(470, 400, 470, 200, 350, 200);
      p.L(250, 200);
      p.stroke();
      p.M(150, 50);
      p.L(400, 50);
      p.stroke();
      p.M(170, 120);
      p.L(420, 120);
      p.stroke();
    }
  },
  // ₮ TUGRIK SIGN
  [0x20ae]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(120, 500);
      p.L(480, 500);
      p.stroke();
      p.M(300, 500);
      p.L(300, 0);
      p.stroke();
      p.M(120, 350);
      p.L(480, 350);
      p.stroke();
      p.M(120, 200);
      p.L(480, 200);
      p.stroke();
    }
  },
  // ₱ PESO SIGN
  [0x20b1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(150, 0);
      p.L(150, 500);
      p.stroke();
      p.M(150, 500);
      p.L(350, 500);
      p.C(480, 500, 480, 300, 350, 300);
      p.L(150, 300);
      p.stroke();
      p.M(100, 420);
      p.L(450, 420);
      p.stroke();
    }
  },
  // ₲ GUARANI SIGN
  [0x20b2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(450, 400);
      p.C(400, 480, 200, 480, 150, 350);
      p.C(100, 200, 200, 50, 400, 50);
      p.L(400, 250);
      p.L(250, 250);
      p.stroke();
      p.M(300, 520);
      p.L(300, 0);
      p.stroke();
    }
  },
  // ₳ AUSTRAL SIGN
  [0x20b3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(100, 0);
      p.L(300, 500);
      p.L(500, 0);
      p.stroke();
      p.M(160, 180);
      p.L(440, 180);
      p.stroke();
      p.M(180, 280);
      p.L(420, 280);
      p.stroke();
    }
  },
  // ₴ HRYVNIA SIGN
  [0x20b4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(420, 450);
      p.C(380, 500, 200, 500, 180, 400);
      p.C(160, 300, 400, 250, 420, 150);
      p.C(440, 50, 250, 0, 180, 50);
      p.stroke();
      p.M(100, 330);
      p.L(480, 330);
      p.stroke();
      p.M(100, 220);
      p.L(480, 220);
      p.stroke();
    }
  },
  // ₶ LIVRE TOURNOIS SIGN
  [0x20b6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(250, 0);
      p.L(250, 480);
      p.C(250, 500, 200, 520, 180, 500);
      p.stroke();
      p.M(250, 350);
      p.L(450, 350);
      p.stroke();
    }
  },
  // ₸ TENGE SIGN
  [0x20b8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(50);
      p.M(100, 500);
      p.L(500, 500);
      p.stroke();
      p.M(100, 400);
      p.L(500, 400);
      p.stroke();
      p.M(300, 400);
      p.L(300, 0);
      p.stroke();
    }
  }
};

// =============================================================================
// Misc Math Symbols-A (U+27C0–U+27EF)
// =============================================================================

export const MATH_SYM_A: Record<number, GlyphDef> = {
  // ⟨ MATHEMATICAL LEFT ANGLE BRACKET
  [0x27e8]: {
    width: 350,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(280, 500);
      p.L(100, 250);
      p.L(280, 0);
      p.stroke();
    }
  },
  // ⟩ MATHEMATICAL RIGHT ANGLE BRACKET
  [0x27e9]: {
    width: 350,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(100, 500);
      p.L(280, 250);
      p.L(100, 0);
      p.stroke();
    }
  },
  // ⟪ MATHEMATICAL LEFT DOUBLE ANGLE BRACKET
  [0x27ea]: {
    width: 450,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(250, 500);
      p.L(70, 250);
      p.L(250, 0);
      p.stroke();
      p.M(380, 500);
      p.L(200, 250);
      p.L(380, 0);
      p.stroke();
    }
  },
  // ⟫ MATHEMATICAL RIGHT DOUBLE ANGLE BRACKET
  [0x27eb]: {
    width: 450,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(70, 500);
      p.L(250, 250);
      p.L(70, 0);
      p.stroke();
      p.M(200, 500);
      p.L(380, 250);
      p.L(200, 0);
      p.stroke();
    }
  }
};

// =============================================================================
// Supplemental Arrows-A (U+27F0–U+27FF)
// =============================================================================

export const SUP_ARROWS_A: Record<number, GlyphDef> = {
  // ⟵ LONG LEFTWARDS ARROW
  [0x27f5]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(650, 250);
      p.L(100, 250);
      p.stroke();
      p.M(100, 250);
      p.L(250, 380);
      p.L(250, 120);
      p.Z();
      p.fill();
    }
  },
  // ⟶ LONG RIGHTWARDS ARROW
  [0x27f6]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      p.M(50, 250);
      p.L(600, 250);
      p.stroke();
      p.M(600, 250);
      p.L(450, 380);
      p.L(450, 120);
      p.Z();
      p.fill();
    }
  },
  // ⟷ LONG LEFT RIGHT ARROW
  [0x27f7]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 250);
      p.L(550, 250);
      p.stroke();
      p.M(100, 250);
      p.L(230, 370);
      p.L(230, 130);
      p.Z();
      p.fill();
      p.M(600, 250);
      p.L(470, 370);
      p.L(470, 130);
      p.Z();
      p.fill();
    }
  },
  // ⟸ LONG LEFTWARDS DOUBLE ARROW
  [0x27f8]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.M(80, 250);
      p.L(250, 430);
      p.L(250, 310);
      p.L(620, 310);
      p.L(620, 190);
      p.L(250, 190);
      p.L(250, 70);
      p.Z();
      p.fill();
    }
  },
  // ⟹ LONG RIGHTWARDS DOUBLE ARROW
  [0x27f9]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.M(620, 250);
      p.L(450, 430);
      p.L(450, 310);
      p.L(80, 310);
      p.L(80, 190);
      p.L(450, 190);
      p.L(450, 70);
      p.Z();
      p.fill();
    }
  },
  // ⟺ LONG LEFT RIGHT DOUBLE ARROW
  [0x27fa]: {
    width: 700,
    draw: (p: GlyphPen) => {
      p.M(80, 250);
      p.L(200, 400);
      p.L(200, 300);
      p.L(500, 300);
      p.L(500, 400);
      p.L(620, 250);
      p.L(500, 100);
      p.L(500, 200);
      p.L(200, 200);
      p.L(200, 100);
      p.Z();
      p.fill();
    }
  }
};

// =============================================================================
// Extended Geometric Shapes (U+25A7–U+25FF) — remaining characters
// =============================================================================

// Square bounds used throughout: x=100 y=50 w=400 h=400
const SQ_X = 100;
const SQ_Y = 50;
const SQ_W = 400;
const SQ_H = 400;
const SQ_R = SQ_X + SQ_W; // 500
const SQ_T = SQ_Y + SQ_H; // 450

export const GEOMETRIC_EXT: Record<number, GlyphDef> = {
  // 0x25A7: SQUARE WITH UPPER LEFT TO LOWER RIGHT FILL
  [0x25a7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.lineWidth(20);
      for (let i = -3; i <= 3; i++) {
        const off = i * 100;
        p.M(SQ_X + off, SQ_T);
        p.L(SQ_R + off, SQ_Y);
        p.stroke();
      }
    }
  },
  // 0x25A9: SQUARE WITH DIAGONAL CROSSHATCH
  [0x25a9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.lineWidth(15);
      for (let i = -3; i <= 3; i++) {
        const off = i * 100;
        p.M(SQ_X + off, SQ_T);
        p.L(SQ_R + off, SQ_Y);
        p.stroke();
        p.M(SQ_X + off, SQ_Y);
        p.L(SQ_R + off, SQ_T);
        p.stroke();
      }
    }
  },
  // 0x25B1: WHITE PARALLELOGRAM
  [0x25b1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(180, 50);
      p.L(500, 50);
      p.L(420, 450);
      p.L(100, 450);
      p.Z();
      p.stroke();
    }
  },
  // 0x25B9: WHITE RIGHT-POINTING SMALL TRIANGLE
  [0x25b9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(200, 380);
      p.L(430, 250);
      p.L(200, 120);
      p.Z();
      p.stroke();
    }
  },
  // 0x25BA: BLACK RIGHT-POINTING POINTER
  [0x25ba]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(150, 420);
      p.L(480, 250);
      p.L(150, 80);
      p.Z();
      p.fill();
    }
  },
  // 0x25BB: WHITE RIGHT-POINTING POINTER
  [0x25bb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(150, 420);
      p.L(480, 250);
      p.L(150, 80);
      p.Z();
      p.stroke();
    }
  },
  // 0x25BF: WHITE DOWN-POINTING SMALL TRIANGLE
  [0x25bf]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(180, 380);
      p.L(420, 380);
      p.L(300, 120);
      p.Z();
      p.stroke();
    }
  },
  // 0x25C3: WHITE LEFT-POINTING SMALL TRIANGLE
  [0x25c3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(400, 380);
      p.L(170, 250);
      p.L(400, 120);
      p.Z();
      p.stroke();
    }
  },
  // 0x25C4: BLACK LEFT-POINTING POINTER
  [0x25c4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(450, 420);
      p.L(120, 250);
      p.L(450, 80);
      p.Z();
      p.fill();
    }
  },
  // 0x25C5: WHITE LEFT-POINTING POINTER
  [0x25c5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(450, 420);
      p.L(120, 250);
      p.L(450, 80);
      p.Z();
      p.stroke();
    }
  },
  // 0x25E2: BLACK LOWER RIGHT TRIANGLE
  [0x25e2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(SQ_R, SQ_T);
      p.L(SQ_R, SQ_Y);
      p.L(SQ_X, SQ_Y);
      p.Z();
      p.fill();
    }
  },
  // 0x25E3: BLACK LOWER LEFT TRIANGLE
  [0x25e3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(SQ_X, SQ_T);
      p.L(SQ_X, SQ_Y);
      p.L(SQ_R, SQ_Y);
      p.Z();
      p.fill();
    }
  },
  // 0x25E4: BLACK UPPER LEFT TRIANGLE
  [0x25e4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(SQ_X, SQ_Y);
      p.L(SQ_X, SQ_T);
      p.L(SQ_R, SQ_T);
      p.Z();
      p.fill();
    }
  },
  // 0x25E5: BLACK UPPER RIGHT TRIANGLE
  [0x25e5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.M(SQ_R, SQ_Y);
      p.L(SQ_R, SQ_T);
      p.L(SQ_X, SQ_T);
      p.Z();
      p.fill();
    }
  },
  // 0x25E7: SQUARE WITH LEFT HALF BLACK
  [0x25e7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.rect(SQ_X, SQ_Y, SQ_W / 2, SQ_H);
      p.fill();
    }
  },
  // 0x25E8: SQUARE WITH RIGHT HALF BLACK
  [0x25e8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.rect(SQ_X + SQ_W / 2, SQ_Y, SQ_W / 2, SQ_H);
      p.fill();
    }
  },
  // 0x25E9: SQUARE WITH UPPER LEFT DIAGONAL HALF BLACK
  [0x25e9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.M(SQ_X, SQ_Y);
      p.L(SQ_X, SQ_T);
      p.L(SQ_R, SQ_T);
      p.Z();
      p.fill();
    }
  },
  // 0x25EA: SQUARE WITH LOWER RIGHT DIAGONAL HALF BLACK
  [0x25ea]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.M(SQ_R, SQ_T);
      p.L(SQ_R, SQ_Y);
      p.L(SQ_X, SQ_Y);
      p.Z();
      p.fill();
    }
  },
  // 0x25EB: WHITE SQUARE WITH VERTICAL BISECTING LINE
  [0x25eb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.M(300, SQ_Y);
      p.L(300, SQ_T);
      p.stroke();
    }
  },
  // 0x25EC: WHITE UP-POINTING TRIANGLE WITH DOT
  [0x25ec]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(100, 50);
      p.L(500, 50);
      p.Z();
      p.stroke();
      p.circle(300, 190, 30);
      p.fill();
    }
  },
  // 0x25ED: UP-POINTING TRIANGLE WITH LEFT HALF BLACK
  [0x25ed]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(100, 50);
      p.L(500, 50);
      p.Z();
      p.stroke();
      p.M(300, 470);
      p.L(100, 50);
      p.L(300, 50);
      p.Z();
      p.fill();
    }
  },
  // 0x25EE: UP-POINTING TRIANGLE WITH RIGHT HALF BLACK
  [0x25ee]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      p.M(300, 470);
      p.L(100, 50);
      p.L(500, 50);
      p.Z();
      p.stroke();
      p.M(300, 470);
      p.L(500, 50);
      p.L(300, 50);
      p.Z();
      p.fill();
    }
  },
  // 0x25F0: WHITE SQUARE WITH UPPER LEFT QUADRANT
  [0x25f0]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.rect(SQ_X, SQ_Y + SQ_H / 2, SQ_W / 2, SQ_H / 2);
      p.fill();
    }
  },
  // 0x25F1: WHITE SQUARE WITH LOWER LEFT QUADRANT
  [0x25f1]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.rect(SQ_X, SQ_Y, SQ_W / 2, SQ_H / 2);
      p.fill();
    }
  },
  // 0x25F2: WHITE SQUARE WITH LOWER RIGHT QUADRANT
  [0x25f2]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.rect(SQ_X + SQ_W / 2, SQ_Y, SQ_W / 2, SQ_H / 2);
      p.fill();
    }
  },
  // 0x25F3: WHITE SQUARE WITH UPPER RIGHT QUADRANT
  [0x25f3]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(SQ_X, SQ_Y, SQ_W, SQ_H);
      p.stroke();
      p.rect(SQ_X + SQ_W / 2, SQ_Y + SQ_H / 2, SQ_W / 2, SQ_H / 2);
      p.fill();
    }
  },
  // 0x25F4: WHITE CIRCLE WITH UPPER LEFT QUADRANT
  [0x25f4]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 220);
      p.stroke();
      // Fill upper-left quadrant using a wedge path
      p.M(300, 250);
      p.L(300, 470);
      p.C(179, 470, 80, 371, 80, 250);
      p.L(300, 250);
      p.Z();
      p.fill();
    }
  },
  // 0x25F5: WHITE CIRCLE WITH LOWER LEFT QUADRANT
  [0x25f5]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(300, 250);
      p.L(80, 250);
      p.C(80, 129, 179, 30, 300, 30);
      p.L(300, 250);
      p.Z();
      p.fill();
    }
  },
  // 0x25F6: WHITE CIRCLE WITH LOWER RIGHT QUADRANT
  [0x25f6]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(300, 250);
      p.L(300, 30);
      p.C(421, 30, 520, 129, 520, 250);
      p.L(300, 250);
      p.Z();
      p.fill();
    }
  },
  // 0x25F7: WHITE CIRCLE WITH UPPER RIGHT QUADRANT
  [0x25f7]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.circle(300, 250, 220);
      p.stroke();
      p.M(300, 250);
      p.L(520, 250);
      p.C(520, 371, 421, 470, 300, 470);
      p.L(300, 250);
      p.Z();
      p.fill();
    }
  },
  // 0x25F8: UPPER LEFT TRIANGLE (outline)
  [0x25f8]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(SQ_X, SQ_T);
      p.L(SQ_R, SQ_T);
      p.L(SQ_X, SQ_Y);
      p.Z();
      p.stroke();
    }
  },
  // 0x25F9: UPPER RIGHT TRIANGLE (outline)
  [0x25f9]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(SQ_X, SQ_T);
      p.L(SQ_R, SQ_T);
      p.L(SQ_R, SQ_Y);
      p.Z();
      p.stroke();
    }
  },
  // 0x25FA: LOWER LEFT TRIANGLE (outline)
  [0x25fa]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(SQ_X, SQ_T);
      p.L(SQ_X, SQ_Y);
      p.L(SQ_R, SQ_Y);
      p.Z();
      p.stroke();
    }
  },
  // 0x25FB: WHITE MEDIUM SQUARE
  [0x25fb]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.rect(130, 80, 340, 340);
      p.stroke();
    }
  },
  // 0x25FC: BLACK MEDIUM SQUARE
  [0x25fc]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(130, 80, 340, 340);
      p.fill();
    }
  },
  // 0x25FD: WHITE MEDIUM SMALL SQUARE
  [0x25fd]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(30);
      p.rect(170, 120, 260, 260);
      p.stroke();
    }
  },
  // 0x25FE: BLACK MEDIUM SMALL SQUARE
  [0x25fe]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.rect(170, 120, 260, 260);
      p.fill();
    }
  },
  // 0x25FF: LOWER RIGHT TRIANGLE (outline)
  [0x25ff]: {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(35);
      p.M(SQ_R, SQ_T);
      p.L(SQ_R, SQ_Y);
      p.L(SQ_X, SQ_Y);
      p.Z();
      p.stroke();
    }
  }
};

// =============================================================================
// Roman Numerals (U+2165–U+2183)
// =============================================================================
// Uses simple stroke approximations: vertical lines for I, V-shapes for V,
// X-shapes for X, etc. These are distinctive but not actual letter shapes.

/** Single vertical stroke at given x position. */
function romanI(p: GlyphPen, x: number, bot: number, top: number): void {
  p.M(x, bot);
  p.L(x, top);
  p.stroke();
}

/** V-shape centred at cx. */
function romanV(p: GlyphPen, cx: number, bot: number, top: number): void {
  const hw = 80;
  p.M(cx - hw, top);
  p.L(cx, bot);
  p.L(cx + hw, top);
  p.stroke();
}

/** X-shape centred at cx. */
function romanX(p: GlyphPen, cx: number, bot: number, top: number): void {
  const hw = 70;
  p.M(cx - hw, bot);
  p.L(cx + hw, top);
  p.stroke();
  p.M(cx + hw, bot);
  p.L(cx - hw, top);
  p.stroke();
}

/** Horizontal line (for L and C approximation). */
function romanHL(p: GlyphPen, x1: number, x2: number, y: number): void {
  p.M(x1, y);
  p.L(x2, y);
  p.stroke();
}

// Capital roman numeral helpers — draws the appropriate symbol pattern
function capitalRoman(value: number): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(45);
      const bot = 0;
      const top = 500;
      if (value <= 3) {
        // I, II, III — evenly spaced vertical strokes
        const n = value;
        const gap = 120;
        const startX = 300 - ((n - 1) * gap) / 2;
        for (let i = 0; i < n; i++) {
          romanI(p, startX + i * gap, bot, top);
        }
      } else if (value === 4) {
        // IV
        romanI(p, 200, bot, top);
        romanV(p, 380, bot, top);
      } else if (value === 5) {
        // V
        romanV(p, 300, bot, top);
      } else if (value === 6) {
        // VI
        romanV(p, 200, bot, top);
        romanI(p, 400, bot, top);
      } else if (value === 7) {
        // VII
        romanV(p, 160, bot, top);
        romanI(p, 340, bot, top);
        romanI(p, 460, bot, top);
      } else if (value === 8) {
        // VIII
        p.lineWidth(40);
        romanV(p, 140, bot, top);
        romanI(p, 300, bot, top);
        romanI(p, 400, bot, top);
        romanI(p, 500, bot, top);
      } else if (value === 9) {
        // IX
        romanI(p, 200, bot, top);
        romanX(p, 400, bot, top);
      } else if (value === 10) {
        // X
        romanX(p, 300, bot, top);
      } else if (value === 11) {
        // XI
        romanX(p, 220, bot, top);
        romanI(p, 430, bot, top);
      } else if (value === 12) {
        // XII
        romanX(p, 180, bot, top);
        romanI(p, 370, bot, top);
        romanI(p, 480, bot, top);
      } else if (value === 50) {
        // L — vertical + horizontal base
        romanI(p, 200, bot, top);
        romanHL(p, 200, 450, bot);
      } else if (value === 100) {
        // C — open arc approximation (two horizontal + vertical)
        romanHL(p, 180, 450, top);
        romanI(p, 180, bot, top);
        romanHL(p, 180, 450, bot);
      } else if (value === 500) {
        // D — vertical + curved right side
        romanI(p, 180, bot, top);
        p.M(180, top);
        p.C(480, top, 480, bot, 180, bot);
        p.stroke();
      } else if (value === 1000) {
        // M — two V-shapes pointing down
        romanI(p, 100, bot, top);
        p.M(100, top);
        p.L(230, bot + 200);
        p.L(300, top);
        p.stroke();
        p.M(300, top);
        p.L(370, bot + 200);
        p.L(500, top);
        p.stroke();
        romanI(p, 500, bot, top);
      }
    }
  };
}

// Small roman numeral — same patterns but shorter (70% height, shifted up a bit)
function smallRoman(value: number): GlyphDef {
  return {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(40);
      const bot = 0;
      const top = 350;
      if (value <= 3) {
        const n = value;
        const gap = 100;
        const startX = 300 - ((n - 1) * gap) / 2;
        for (let i = 0; i < n; i++) {
          romanI(p, startX + i * gap, bot, top);
        }
      } else if (value === 4) {
        romanI(p, 200, bot, top);
        romanV(p, 380, bot, top);
      } else if (value === 5) {
        romanV(p, 300, bot, top);
      } else if (value === 6) {
        romanV(p, 200, bot, top);
        romanI(p, 400, bot, top);
      } else if (value === 7) {
        romanV(p, 160, bot, top);
        romanI(p, 340, bot, top);
        romanI(p, 450, bot, top);
      } else if (value === 8) {
        p.lineWidth(35);
        romanV(p, 130, bot, top);
        romanI(p, 280, bot, top);
        romanI(p, 380, bot, top);
        romanI(p, 480, bot, top);
      } else if (value === 9) {
        romanI(p, 200, bot, top);
        romanX(p, 400, bot, top);
      } else if (value === 10) {
        romanX(p, 300, bot, top);
      } else if (value === 11) {
        romanX(p, 220, bot, top);
        romanI(p, 430, bot, top);
      } else if (value === 12) {
        romanX(p, 180, bot, top);
        romanI(p, 370, bot, top);
        romanI(p, 480, bot, top);
      } else if (value === 50) {
        romanI(p, 200, bot, top);
        romanHL(p, 200, 450, bot);
      } else if (value === 100) {
        romanHL(p, 180, 450, top);
        romanI(p, 180, bot, top);
        romanHL(p, 180, 450, bot);
      } else if (value === 500) {
        romanI(p, 180, bot, top);
        p.M(180, top);
        p.C(480, top, 480, bot, 180, bot);
        p.stroke();
      } else if (value === 1000) {
        romanI(p, 100, bot, top);
        p.M(100, top);
        p.L(230, bot + 140);
        p.L(300, top);
        p.stroke();
        p.M(300, top);
        p.L(370, bot + 140);
        p.L(500, top);
        p.stroke();
        romanI(p, 500, bot, top);
      }
    }
  };
}

// Values for U+2165–U+216F (Ⅵ–Ⅿ): index→value
const CAPITAL_VALUES: Array<[number, number]> = [
  [0x2165, 6],
  [0x2166, 7],
  [0x2167, 8],
  [0x2168, 9],
  [0x2169, 10],
  [0x216a, 11],
  [0x216b, 12],
  [0x216c, 50],
  [0x216d, 100],
  [0x216e, 500],
  [0x216f, 1000]
];

// Values for U+2170–U+217F (ⅰ–ⅿ)
const SMALL_VALUES: Array<[number, number]> = [
  [0x2170, 1],
  [0x2171, 2],
  [0x2172, 3],
  [0x2173, 4],
  [0x2174, 5],
  [0x2175, 6],
  [0x2176, 7],
  [0x2177, 8],
  [0x2178, 9],
  [0x2179, 10],
  [0x217a, 11],
  [0x217b, 12],
  [0x217c, 50],
  [0x217d, 100],
  [0x217e, 500],
  [0x217f, 1000]
];

export const ROMAN_NUMERALS: Record<number, GlyphDef> = {};

for (const [cp, val] of CAPITAL_VALUES) {
  ROMAN_NUMERALS[cp] = capitalRoman(val);
}
for (const [cp, val] of SMALL_VALUES) {
  ROMAN_NUMERALS[cp] = smallRoman(val);
}

// U+2180: ROMAN NUMERAL ONE THOUSAND C D — approx as large circle
ROMAN_NUMERALS[0x2180] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.circle(300, 250, 220);
    p.stroke();
    p.M(300, 470);
    p.L(300, 30);
    p.stroke();
  }
};
// U+2181: ROMAN NUMERAL FIVE THOUSAND — approx as large D shape
ROMAN_NUMERALS[0x2181] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(180, 500);
    p.L(180, 0);
    p.stroke();
    p.M(180, 500);
    p.C(520, 500, 520, 0, 180, 0);
    p.stroke();
    p.M(300, 500);
    p.L(300, 0);
    p.stroke();
  }
};
// U+2182: ROMAN NUMERAL TEN THOUSAND — approx as double circle
ROMAN_NUMERALS[0x2182] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(40);
    p.circle(300, 250, 220);
    p.stroke();
    p.circle(300, 250, 150);
    p.stroke();
  }
};
// U+2183: ROMAN NUMERAL REVERSED ONE HUNDRED (Ↄ) — reversed C
ROMAN_NUMERALS[0x2183] = {
  width: W,
  draw: (p: GlyphPen) => {
    p.lineWidth(45);
    p.M(150, 450);
    p.C(450, 450, 450, 50, 150, 50);
    p.stroke();
  }
};

// =============================================================================
// Extended Enclosed Alphanumerics (U+2480–U+24FF)
// =============================================================================

export const ENCLOSED_EXT: Record<number, GlyphDef> = {};

// --- Helpers for digit/letter drawing in enclosed forms ---

type StrokeFn = (p: GlyphPen, cx: number, cy: number, s: number) => void;

const EXT_DIGIT_PATHS: StrokeFn[] = [
  // 0
  (p, cx, cy, s) => {
    p.ellipse(cx, cy, s * 0.28, s * 0.42);
    p.stroke();
  },
  // 1
  (p, cx, cy, s) => {
    p.M(cx, cy + s * 0.42);
    p.L(cx, cy - s * 0.42);
    p.stroke();
  },
  // 2
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy + s * 0.25);
    p.C(cx - s * 0.25, cy + s * 0.45, cx + s * 0.25, cy + s * 0.45, cx + s * 0.25, cy + s * 0.15);
    p.L(cx - s * 0.25, cy - s * 0.42);
    p.L(cx + s * 0.25, cy - s * 0.42);
    p.stroke();
  },
  // 3
  (p, cx, cy, s) => {
    p.M(cx - s * 0.22, cy + s * 0.42);
    p.L(cx + s * 0.22, cy + s * 0.42);
    p.L(cx + s * 0.22, cy);
    p.L(cx - s * 0.1, cy);
    p.stroke();
    p.M(cx + s * 0.22, cy);
    p.L(cx + s * 0.22, cy - s * 0.42);
    p.L(cx - s * 0.22, cy - s * 0.42);
    p.stroke();
  },
  // 4
  (p, cx, cy, s) => {
    p.M(cx + s * 0.2, cy + s * 0.42);
    p.L(cx - s * 0.25, cy - s * 0.05);
    p.L(cx + s * 0.25, cy - s * 0.05);
    p.stroke();
    p.M(cx + s * 0.2, cy + s * 0.42);
    p.L(cx + s * 0.2, cy - s * 0.42);
    p.stroke();
  },
  // 5
  (p, cx, cy, s) => {
    p.M(cx + s * 0.22, cy + s * 0.42);
    p.L(cx - s * 0.22, cy + s * 0.42);
    p.L(cx - s * 0.22, cy + s * 0.05);
    p.L(cx + s * 0.22, cy + s * 0.05);
    p.C(cx + s * 0.35, cy + s * 0.05, cx + s * 0.35, cy - s * 0.42, cx - s * 0.22, cy - s * 0.42);
    p.stroke();
  },
  // 6
  (p, cx, cy, s) => {
    p.M(cx + s * 0.2, cy + s * 0.35);
    p.C(cx - s * 0.1, cy + s * 0.45, cx - s * 0.3, cy + s * 0.15, cx - s * 0.25, cy - s * 0.1);
    p.C(cx - s * 0.2, cy - s * 0.45, cx + s * 0.25, cy - s * 0.45, cx + s * 0.25, cy - s * 0.1);
    p.C(cx + s * 0.25, cy + s * 0.1, cx - s * 0.25, cy + s * 0.1, cx - s * 0.25, cy - s * 0.1);
    p.stroke();
  },
  // 7
  (p, cx, cy, s) => {
    p.M(cx - s * 0.22, cy + s * 0.42);
    p.L(cx + s * 0.22, cy + s * 0.42);
    p.L(cx - s * 0.05, cy - s * 0.42);
    p.stroke();
  },
  // 8
  (p, cx, cy, s) => {
    p.ellipse(cx, cy + s * 0.22, s * 0.2, s * 0.2);
    p.stroke();
    p.ellipse(cx, cy - s * 0.22, s * 0.22, s * 0.22);
    p.stroke();
  },
  // 9
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy - s * 0.35);
    p.C(cx + s * 0.1, cy - s * 0.45, cx + s * 0.3, cy - s * 0.15, cx + s * 0.25, cy + s * 0.1);
    p.C(cx + s * 0.2, cy + s * 0.45, cx - s * 0.25, cy + s * 0.45, cx - s * 0.25, cy + s * 0.1);
    p.C(cx - s * 0.25, cy - s * 0.1, cx + s * 0.25, cy - s * 0.1, cx + s * 0.25, cy + s * 0.1);
    p.stroke();
  }
];

function extDrawDigit(p: GlyphPen, d: number, cx: number, cy: number, s: number): void {
  EXT_DIGIT_PATHS[d](p, cx, cy, s);
}

const EXT_LETTER_PATHS: StrokeFn[] = [
  // a
  (p, cx, cy, s) => {
    p.M(cx + s * 0.25, cy - s * 0.3);
    p.C(cx + s * 0.1, cy - s * 0.45, cx - s * 0.25, cy - s * 0.35, cx - s * 0.25, cy - s * 0.1);
    p.C(cx - s * 0.25, cy + s * 0.1, cx + s * 0.25, cy + s * 0.1, cx + s * 0.25, cy - s * 0.1);
    p.L(cx + s * 0.25, cy - s * 0.35);
    p.stroke();
  },
  // b
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy + s * 0.45);
    p.L(cx - s * 0.2, cy - s * 0.45);
    p.stroke();
    p.M(cx - s * 0.2, cy - s * 0.05);
    p.C(cx - s * 0.2, cy - s * 0.45, cx + s * 0.25, cy - s * 0.45, cx + s * 0.25, cy - s * 0.15);
    p.C(cx + s * 0.25, cy + s * 0.1, cx - s * 0.2, cy + s * 0.1, cx - s * 0.2, cy - s * 0.05);
    p.stroke();
  },
  // c
  (p, cx, cy, s) => {
    p.M(cx + s * 0.2, cy - s * 0.25);
    p.C(cx + s * 0.05, cy - s * 0.45, cx - s * 0.25, cy - s * 0.35, cx - s * 0.25, cy - s * 0.1);
    p.C(cx - s * 0.25, cy + s * 0.15, cx + s * 0.05, cy + s * 0.15, cx + s * 0.2, cy + s * 0.05);
    p.stroke();
  },
  // d
  (p, cx, cy, s) => {
    p.M(cx + s * 0.2, cy + s * 0.45);
    p.L(cx + s * 0.2, cy - s * 0.45);
    p.stroke();
    p.M(cx + s * 0.2, cy - s * 0.05);
    p.C(cx + s * 0.2, cy - s * 0.45, cx - s * 0.25, cy - s * 0.45, cx - s * 0.25, cy - s * 0.15);
    p.C(cx - s * 0.25, cy + s * 0.1, cx + s * 0.2, cy + s * 0.1, cx + s * 0.2, cy - s * 0.05);
    p.stroke();
  },
  // e
  (p, cx, cy, s) => {
    p.M(cx - s * 0.25, cy - s * 0.08);
    p.L(cx + s * 0.25, cy - s * 0.08);
    p.C(cx + s * 0.25, cy + s * 0.15, cx - s * 0.25, cy + s * 0.15, cx - s * 0.25, cy - s * 0.08);
    p.C(cx - s * 0.25, cy - s * 0.4, cx + s * 0.25, cy - s * 0.4, cx + s * 0.25, cy - s * 0.25);
    p.stroke();
  },
  // f
  (p, cx, cy, s) => {
    p.M(cx + s * 0.1, cy + s * 0.4);
    p.C(cx + s * 0.1, cy + s * 0.5, cx - s * 0.1, cy + s * 0.5, cx - s * 0.1, cy + s * 0.4);
    p.L(cx - s * 0.1, cy - s * 0.4);
    p.stroke();
    p.M(cx - s * 0.2, cy + s * 0.15);
    p.L(cx + s * 0.15, cy + s * 0.15);
    p.stroke();
  },
  // g
  (p, cx, cy, s) => {
    p.M(cx + s * 0.22, cy + s * 0.1);
    p.C(cx + s * 0.22, cy + s * 0.4, cx - s * 0.22, cy + s * 0.4, cx - s * 0.22, cy + s * 0.1);
    p.C(cx - s * 0.22, cy - s * 0.2, cx + s * 0.22, cy - s * 0.2, cx + s * 0.22, cy + s * 0.1);
    p.L(cx + s * 0.22, cy - s * 0.3);
    p.C(cx + s * 0.22, cy - s * 0.5, cx - s * 0.22, cy - s * 0.5, cx - s * 0.22, cy - s * 0.35);
    p.stroke();
  },
  // h
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy + s * 0.45);
    p.L(cx - s * 0.2, cy - s * 0.45);
    p.stroke();
    p.M(cx - s * 0.2, cy + s * 0.05);
    p.C(cx - s * 0.2, cy + s * 0.25, cx + s * 0.2, cy + s * 0.25, cx + s * 0.2, cy + s * 0.05);
    p.L(cx + s * 0.2, cy - s * 0.45);
    p.stroke();
  },
  // i
  (p, cx, cy, s) => {
    p.M(cx, cy + s * 0.15);
    p.L(cx, cy - s * 0.35);
    p.stroke();
    p.circle(cx, cy + s * 0.3, s * 0.06);
    p.fill();
  },
  // j
  (p, cx, cy, s) => {
    p.M(cx + s * 0.05, cy + s * 0.15);
    p.L(cx + s * 0.05, cy - s * 0.35);
    p.C(cx + s * 0.05, cy - s * 0.5, cx - s * 0.15, cy - s * 0.5, cx - s * 0.15, cy - s * 0.35);
    p.stroke();
    p.circle(cx + s * 0.05, cy + s * 0.3, s * 0.06);
    p.fill();
  },
  // k
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy + s * 0.45);
    p.L(cx - s * 0.2, cy - s * 0.45);
    p.stroke();
    p.M(cx + s * 0.2, cy + s * 0.15);
    p.L(cx - s * 0.2, cy - s * 0.1);
    p.L(cx + s * 0.2, cy - s * 0.35);
    p.stroke();
  },
  // l
  (p, cx, cy, s) => {
    p.M(cx, cy + s * 0.45);
    p.L(cx, cy - s * 0.45);
    p.stroke();
  },
  // m
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy - s * 0.35);
    p.L(cx - s * 0.3, cy + s * 0.15);
    p.C(cx - s * 0.3, cy + s * 0.25, cx - s * 0.05, cy + s * 0.25, cx - s * 0.05, cy + s * 0.15);
    p.L(cx - s * 0.05, cy - s * 0.35);
    p.stroke();
    p.M(cx - s * 0.05, cy + s * 0.15);
    p.C(cx - s * 0.05, cy + s * 0.25, cx + s * 0.2, cy + s * 0.25, cx + s * 0.2, cy + s * 0.15);
    p.L(cx + s * 0.2, cy - s * 0.35);
    p.stroke();
  },
  // n
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy - s * 0.35);
    p.L(cx - s * 0.2, cy + s * 0.15);
    p.C(cx - s * 0.2, cy + s * 0.25, cx + s * 0.2, cy + s * 0.25, cx + s * 0.2, cy + s * 0.15);
    p.L(cx + s * 0.2, cy - s * 0.35);
    p.stroke();
  },
  // o
  (p, cx, cy, s) => {
    p.ellipse(cx, cy - s * 0.1, s * 0.22, s * 0.28);
    p.stroke();
  },
  // p
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy - s * 0.45);
    p.L(cx - s * 0.2, cy + s * 0.15);
    p.stroke();
    p.M(cx - s * 0.2, cy + s * 0.05);
    p.C(cx - s * 0.2, cy + s * 0.35, cx + s * 0.25, cy + s * 0.35, cx + s * 0.25, cy + s * 0.05);
    p.C(cx + s * 0.25, cy - s * 0.2, cx - s * 0.2, cy - s * 0.2, cx - s * 0.2, cy + s * 0.05);
    p.stroke();
  },
  // q
  (p, cx, cy, s) => {
    p.M(cx + s * 0.2, cy - s * 0.45);
    p.L(cx + s * 0.2, cy + s * 0.15);
    p.stroke();
    p.M(cx + s * 0.2, cy + s * 0.05);
    p.C(cx + s * 0.2, cy + s * 0.35, cx - s * 0.25, cy + s * 0.35, cx - s * 0.25, cy + s * 0.05);
    p.C(cx - s * 0.25, cy - s * 0.2, cx + s * 0.2, cy - s * 0.2, cx + s * 0.2, cy + s * 0.05);
    p.stroke();
  },
  // r
  (p, cx, cy, s) => {
    p.M(cx - s * 0.15, cy - s * 0.35);
    p.L(cx - s * 0.15, cy + s * 0.15);
    p.stroke();
    p.M(cx - s * 0.15, cy + s * 0.05);
    p.C(cx - s * 0.15, cy + s * 0.25, cx + s * 0.15, cy + s * 0.25, cx + s * 0.2, cy + s * 0.1);
    p.stroke();
  },
  // s
  (p, cx, cy, s) => {
    p.M(cx + s * 0.18, cy + s * 0.08);
    p.C(cx + s * 0.18, cy + s * 0.22, cx - s * 0.18, cy + s * 0.22, cx - s * 0.18, cy + s * 0.05);
    p.C(cx - s * 0.18, cy - s * 0.1, cx + s * 0.18, cy - s * 0.1, cx + s * 0.18, cy - s * 0.22);
    p.C(cx + s * 0.18, cy - s * 0.38, cx - s * 0.18, cy - s * 0.38, cx - s * 0.18, cy - s * 0.25);
    p.stroke();
  },
  // t
  (p, cx, cy, s) => {
    p.M(cx, cy + s * 0.35);
    p.L(cx, cy - s * 0.35);
    p.C(cx, cy - s * 0.5, cx + s * 0.15, cy - s * 0.5, cx + s * 0.15, cy - s * 0.4);
    p.stroke();
    p.M(cx - s * 0.15, cy + s * 0.15);
    p.L(cx + s * 0.15, cy + s * 0.15);
    p.stroke();
  },
  // u
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy + s * 0.15);
    p.L(cx - s * 0.2, cy - s * 0.15);
    p.C(cx - s * 0.2, cy - s * 0.4, cx + s * 0.2, cy - s * 0.4, cx + s * 0.2, cy - s * 0.15);
    p.L(cx + s * 0.2, cy + s * 0.15);
    p.L(cx + s * 0.2, cy - s * 0.35);
    p.stroke();
  },
  // v
  (p, cx, cy, s) => {
    p.M(cx - s * 0.22, cy + s * 0.15);
    p.L(cx, cy - s * 0.35);
    p.L(cx + s * 0.22, cy + s * 0.15);
    p.stroke();
  },
  // w
  (p, cx, cy, s) => {
    p.M(cx - s * 0.3, cy + s * 0.15);
    p.L(cx - s * 0.15, cy - s * 0.35);
    p.L(cx, cy);
    p.L(cx + s * 0.15, cy - s * 0.35);
    p.L(cx + s * 0.3, cy + s * 0.15);
    p.stroke();
  },
  // x
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy + s * 0.15);
    p.L(cx + s * 0.2, cy - s * 0.35);
    p.stroke();
    p.M(cx + s * 0.2, cy + s * 0.15);
    p.L(cx - s * 0.2, cy - s * 0.35);
    p.stroke();
  },
  // y
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy + s * 0.15);
    p.L(cx, cy - s * 0.1);
    p.L(cx + s * 0.2, cy + s * 0.15);
    p.stroke();
    p.M(cx, cy - s * 0.1);
    p.L(cx - s * 0.1, cy - s * 0.45);
    p.stroke();
  },
  // z
  (p, cx, cy, s) => {
    p.M(cx - s * 0.2, cy + s * 0.15);
    p.L(cx + s * 0.2, cy + s * 0.15);
    p.L(cx - s * 0.2, cy - s * 0.35);
    p.L(cx + s * 0.2, cy - s * 0.35);
    p.stroke();
  }
];

function extDrawLetter(p: GlyphPen, idx: number, cx: number, cy: number, s: number): void {
  EXT_LETTER_PATHS[idx](p, cx, cy, s);
}

function drawParens(p: GlyphPen): void {
  p.M(160, 480);
  p.C(100, 480, 60, 380, 60, 250);
  p.C(60, 120, 100, 20, 160, 20);
  p.stroke();
  p.M(440, 20);
  p.C(500, 20, 540, 120, 540, 250);
  p.C(540, 380, 500, 480, 440, 480);
  p.stroke();
}

// 0x2474–0x2487: Parenthesized digits ⑴ 1 – ⒇ 20
for (let i = 0; i < 20; i++) {
  const num = i + 1;
  ENCLOSED_EXT[0x2474 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(25);
      drawParens(p);
      p.lineWidth(24);
      if (num <= 9) {
        extDrawDigit(p, num, 300, 250, 200);
      } else {
        const d1 = Math.floor(num / 10);
        const d2 = num % 10;
        extDrawDigit(p, d1, 240, 250, 150);
        extDrawDigit(p, d2, 380, 250, 150);
      }
    }
  };
}

// 0x2480–0x2487: Parenthesized digits ⒀ 13 – ⒇ 20
// (these overlap with the loop above; the loop above covers 0x2474-0x2487 = 1-20,
//  so 0x2480-0x2487 = 13-20 are already covered)

// 0x2488–0x249B: Digit period 1.–20. — digit + period dot
for (let i = 0; i < 20; i++) {
  const num = i + 1;
  ENCLOSED_EXT[0x2488 + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(26);
      if (num <= 9) {
        extDrawDigit(p, num, 260, 250, 210);
      } else {
        const d1 = Math.floor(num / 10);
        const d2 = num % 10;
        extDrawDigit(p, d1, 200, 250, 160);
        extDrawDigit(p, d2, 340, 250, 160);
      }
      // Period dot
      p.circle(num <= 9 ? 400 : 460, 80, 28);
      p.fill();
    }
  };
}

// 0x249C–0x24B5: Parenthesized Latin small letters ⒜–⒵
for (let i = 0; i < 26; i++) {
  ENCLOSED_EXT[0x249c + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      p.lineWidth(25);
      drawParens(p);
      p.lineWidth(24);
      extDrawLetter(p, i, 300, 250, 200);
    }
  };
}

// 0x24EB–0x24FF: Negative circled numbers 11–20 and others
// Use thick circle outline + number (best approximation since we can't do white-on-black)
for (let i = 0; i <= 0x24ff - 0x24eb; i++) {
  const num = i + 11; // 0x24EB = ⓫ (11), ..., 0x24F4 = ⓴ (20), 0x24F5+ are circled 1-10 again, 0x24FF = ⓿ (0)
  ENCLOSED_EXT[0x24eb + i] = {
    width: W,
    draw: (p: GlyphPen) => {
      // Thick circle outline to suggest filled appearance
      p.lineWidth(55);
      p.circle(300, 250, 215);
      p.stroke();
      p.lineWidth(24);
      if (i <= 9) {
        // 11-20: two digits
        const d1 = Math.floor(num / 10);
        const d2 = num % 10;
        extDrawDigit(p, d1, 230, 250, 150);
        extDrawDigit(p, d2, 370, 250, 150);
      } else if (i <= 19) {
        // 0x24F5-0x24FE: double-circled 1-10
        const d = i - 9;
        if (d <= 9) {
          extDrawDigit(p, d, 300, 250, 170);
        } else {
          extDrawDigit(p, 1, 230, 250, 140);
          extDrawDigit(p, 0, 370, 250, 140);
        }
      } else {
        // 0x24FF: ⓿ negative circled zero
        extDrawDigit(p, 0, 300, 250, 170);
      }
    }
  };
}
