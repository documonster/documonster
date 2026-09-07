/**
 * Cyrillic letterforms for the Type3 fallback — U+0400…U+045F.
 *
 * The whole of the modern range: Russian, Bulgarian, Serbian, Macedonian, Ukrainian
 * and Belarusian all read out of these 96 code points. The Cyrillic Supplement and
 * the historic letters above U+045F are absent for the same reason polytonic Greek is
 * — a document that needs them needs a real font, and `requiresEmbeddedFace` still
 * asks for one.
 *
 * Three kinds of reuse do most of the work, and the second is the one that makes this
 * table small:
 *
 *   - **Eleven capitals are Latin letters.** `А В Е К М Н О Р С Т Х` are the same
 *     outlines, taken from the shared stroke font.
 *   - **Cyrillic lowercase is largely its own uppercase at x-height.** `в к м н п т`
 *     and a dozen more are small capitals, not distinct shapes as in Latin, so
 *     {@link asSmall} derives them from the capitals above. Only `б г д е з и й у ф`
 *     and the soft signs need their own drawing or a Latin twin.
 *   - **Three capitals are Greek.** `Г` `П` `Ф` are `Γ` `Π` `Φ`, and are built the
 *     same way here rather than imported, because the Greek table's exports are
 *     glyphs and what this needs is the shape.
 *
 * @module
 */

import type { GlyphDef } from "@pdf/font/type3-glyphs";
import {
  acute,
  ascii,
  asSmall,
  breve,
  diaeresis,
  grave,
  merge,
  shape,
  toGlyphDef,
  withMarkOver,
  withMarkOverSmall,
  type Path,
  type StrokeShape
} from "@pdf/font/type3-letterforms";

// =============================================================================
// Capitals
// =============================================================================

/** Б — `Ь` with a top arm. */
const BE = shape(
  0.5,
  [
    [0.1, 0.15],
    [0.42, 0.15]
  ],
  [
    [0.1, 0.15],
    [0.1, 0.75],
    [0.3, 0.75],
    [0.4, 0.67],
    [0.4, 0.52],
    [0.3, 0.45],
    [0.1, 0.45]
  ]
);

/** Г — a stem and a top arm. */
const GHE = shape(0.44, [
  [0.1, 0.75],
  [0.1, 0.15],
  [0.4, 0.15]
]);

/** Д — a bowl on two feet. */
const DE = shape(
  0.56,
  [
    [0.06, 0.85],
    [0.06, 0.75],
    [0.14, 0.75],
    [0.2, 0.15],
    [0.42, 0.15],
    [0.42, 0.75],
    [0.5, 0.75],
    [0.5, 0.85]
  ],
  [
    [0.14, 0.75],
    [0.42, 0.75]
  ]
);

/** Ж — a stem with four arms. */
const ZHE = shape(
  0.66,
  [
    [0.33, 0.15],
    [0.33, 0.75]
  ],
  [
    [0.06, 0.15],
    [0.31, 0.45]
  ],
  [
    [0.06, 0.75],
    [0.31, 0.45]
  ],
  [
    [0.6, 0.15],
    [0.35, 0.45]
  ],
  [
    [0.6, 0.75],
    [0.35, 0.45]
  ]
);

/** З — two arcs, like a `3`. */
const ZE = shape(0.46, [
  [0.1, 0.22],
  [0.22, 0.15],
  [0.36, 0.2],
  [0.38, 0.33],
  [0.26, 0.44],
  [0.38, 0.55],
  [0.38, 0.68],
  [0.24, 0.75],
  [0.1, 0.68]
]);

/** И — two stems joined by a rising diagonal. */
const I_CYR = shape(
  0.52,
  [
    [0.1, 0.75],
    [0.1, 0.15]
  ],
  [
    [0.42, 0.75],
    [0.42, 0.15]
  ],
  [
    [0.1, 0.75],
    [0.42, 0.15]
  ]
);

/** Л — a bowl-less `П` with a slanted left leg. */
const EL = shape(0.54, [
  [0.06, 0.75],
  [0.16, 0.4],
  [0.2, 0.15],
  [0.44, 0.15],
  [0.44, 0.75]
]);

/** П — two stems under a bridge. */
const PE = shape(0.54, [
  [0.1, 0.75],
  [0.1, 0.15],
  [0.44, 0.15],
  [0.44, 0.75]
]);

/** У — a `Y` whose tail descends and hooks left. */
const U_CYR = shape(
  0.48,
  [
    [0.06, 0.15],
    [0.26, 0.5]
  ],
  [
    [0.44, 0.15],
    [0.16, 0.86],
    [0.06, 0.9]
  ]
);

/** Ф — a full-height stem through a bowl. */
const EF = merge(
  shape(0.58, [
    [0.29, 0.15],
    [0.29, 0.75]
  ]),
  shape(0.58, [
    [0.29, 0.26],
    [0.16, 0.29],
    [0.09, 0.37],
    [0.09, 0.53],
    [0.16, 0.61],
    [0.29, 0.64],
    [0.42, 0.61],
    [0.49, 0.53],
    [0.49, 0.37],
    [0.42, 0.29],
    [0.29, 0.26]
  ])
);

/** Ц — `И` reversed, with a descending tail. */
const TSE = shape(
  0.54,
  [
    [0.1, 0.15],
    [0.1, 0.75],
    [0.42, 0.75],
    [0.42, 0.15]
  ],
  [
    [0.42, 0.75],
    [0.42, 0.88],
    [0.48, 0.88]
  ]
);

/** Ч — a stem with a cup on its left. */
const CHE = shape(0.5, [
  [0.1, 0.15],
  [0.1, 0.38],
  [0.18, 0.46],
  [0.4, 0.46],
  [0.4, 0.15],
  [0.4, 0.75]
]);

/** Ш — three stems on a bar. */
const SHA = shape(0.66, [
  [0.08, 0.15],
  [0.08, 0.75],
  [0.32, 0.75],
  [0.32, 0.15],
  [0.32, 0.75],
  [0.56, 0.75],
  [0.56, 0.15]
]);

/** Щ — `Ш` with a descending tail. */
const SHCHA = merge(
  SHA,
  shape(0.7, [
    [0.56, 0.75],
    [0.56, 0.88],
    [0.62, 0.88]
  ])
);

/** Ъ — a hard sign: a shoulder and a bowl. */
const HARD = shape(
  0.54,
  [
    [0.06, 0.15],
    [0.2, 0.15]
  ],
  [
    [0.2, 0.15],
    [0.2, 0.75],
    [0.38, 0.75],
    [0.47, 0.67],
    [0.47, 0.53],
    [0.38, 0.46],
    [0.2, 0.46]
  ]
);

/** Ь — a soft sign. */
const SOFT = shape(0.46, [
  [0.1, 0.15],
  [0.1, 0.75],
  [0.28, 0.75],
  [0.38, 0.67],
  [0.38, 0.53],
  [0.28, 0.46],
  [0.1, 0.46]
]);

/** Ы — a soft sign with a stem to its right. */
const YERU = merge(
  SOFT,
  shape(0.62, [
    [0.53, 0.15],
    [0.53, 0.75]
  ])
);

/** Э — a reversed `С` with a waist bar. */
const E_CYR = merge(
  shape(0.48, [
    [0.1, 0.22],
    [0.24, 0.15],
    [0.38, 0.24],
    [0.42, 0.45],
    [0.38, 0.66],
    [0.24, 0.75],
    [0.1, 0.68]
  ]),
  shape(0.48, [
    [0.24, 0.45],
    [0.42, 0.45]
  ])
);

/** Ю — a stem tied to a bowl. */
const YU = merge(
  shape(0.66, [
    [0.08, 0.15],
    [0.08, 0.75]
  ]),
  shape(0.66, [
    [0.08, 0.45],
    [0.24, 0.45]
  ]),
  shape(0.66, [
    [0.38, 0.15],
    [0.26, 0.22],
    [0.24, 0.36],
    [0.24, 0.54],
    [0.26, 0.68],
    [0.38, 0.75],
    [0.52, 0.68],
    [0.56, 0.54],
    [0.56, 0.36],
    [0.52, 0.22],
    [0.38, 0.15]
  ])
);

/** Я — a mirrored `R`. */
const YA = shape(
  0.5,
  [
    [0.42, 0.15],
    [0.42, 0.75]
  ],
  [
    [0.42, 0.15],
    [0.2, 0.15],
    [0.1, 0.23],
    [0.1, 0.37],
    [0.2, 0.45],
    [0.42, 0.45]
  ],
  [
    [0.24, 0.45],
    [0.1, 0.75]
  ]
);

// =============================================================================
// Letters outside the Russian alphabet
// =============================================================================

/** Ђ — a barred stem with a bowl on the right. */
const DJE = merge(
  shape(
    0.54,
    [
      [0.06, 0.28],
      [0.34, 0.28]
    ],
    [
      [0.2, 0.15],
      [0.2, 0.75]
    ]
  ),
  shape(0.54, [
    [0.2, 0.45],
    [0.38, 0.45],
    [0.47, 0.53],
    [0.47, 0.67],
    [0.38, 0.75],
    [0.2, 0.75]
  ])
);

/** Є — a reversed `Э` without the descender: an `E` rounded to the right. */
const UKR_IE = merge(
  shape(0.48, [
    [0.4, 0.22],
    [0.26, 0.15],
    [0.12, 0.24],
    [0.08, 0.45],
    [0.12, 0.66],
    [0.26, 0.75],
    [0.4, 0.68]
  ]),
  shape(0.48, [
    [0.08, 0.45],
    [0.3, 0.45]
  ])
);

/** Љ — `Л` tied to a soft sign. */
const LJE = merge(
  shape(0.72, [
    [0.04, 0.75],
    [0.12, 0.4],
    [0.16, 0.15],
    [0.34, 0.15],
    [0.34, 0.75]
  ]),
  shape(0.72, [
    [0.34, 0.45],
    [0.52, 0.45],
    [0.62, 0.53],
    [0.62, 0.67],
    [0.52, 0.75],
    [0.34, 0.75]
  ])
);

/** Њ — `Н` tied to a soft sign. */
const NJE = merge(
  shape(
    0.72,
    [
      [0.08, 0.15],
      [0.08, 0.75]
    ],
    [
      [0.34, 0.15],
      [0.34, 0.75]
    ],
    [
      [0.08, 0.45],
      [0.34, 0.45]
    ]
  ),
  shape(0.72, [
    [0.34, 0.45],
    [0.52, 0.45],
    [0.62, 0.53],
    [0.62, 0.67],
    [0.52, 0.75],
    [0.34, 0.75]
  ])
);

/** Ћ — a barred stem with a bowl, the Serbian tshe. */
const TSHE = merge(
  shape(
    0.54,
    [
      [0.06, 0.28],
      [0.34, 0.28]
    ],
    [
      [0.2, 0.15],
      [0.2, 0.75]
    ]
  ),
  shape(0.54, [
    [0.2, 0.45],
    [0.38, 0.45],
    [0.47, 0.53],
    [0.47, 0.75]
  ])
);

/** Џ — `Ц` with the tail centred. */
const DZHE = shape(
  0.54,
  [
    [0.1, 0.15],
    [0.1, 0.75],
    [0.42, 0.75],
    [0.42, 0.15]
  ],
  [
    [0.26, 0.75],
    [0.26, 0.88]
  ]
);

/** Й — `И` under a breve. */
const SHORT_I = merge(I_CYR, shape(I_CYR.w, ...breve(0.12, 0.04, 0.28)));

/** Ў — `У` under a breve. */
const U_BREVE = merge(U_CYR, shape(U_CYR.w, ...breve(0.1, 0.04, 0.28)));

/** ў — the lowercase of the above, over the Latin `y` that stands in for `у`. */
const U_BREVE_SMALL = merge(ascii("y"), shape(ascii("y").w, ...breve(0.08, 0.27, 0.26)));

// =============================================================================
// Table
// =============================================================================

/**
 * Put a mark over a letter whose advance is much narrower than its mark.
 *
 * `Ї` and `ї` are the cases: centring a two-dot diaeresis on a stem 0.25 em wide by
 * the usual rule puts the right dot outside the advance.
 */
function overNarrow(base: StrokeShape, mark: (x: number, y: number) => Path[]): StrokeShape {
  return merge(base, shape(base.w, ...mark(base.w * 0.5 - 0.07, 0.02)));
}

const SHAPES: Record<number, StrokeShape> = {
  // U+0400–U+040F — the letters outside Russian, and the accented capitals.
  0x400: withMarkOver(ascii("E"), grave),
  0x401: withMarkOver(ascii("E"), diaeresis),
  0x402: DJE,
  0x403: withMarkOver(GHE, acute),
  0x404: UKR_IE,
  0x405: ascii("S"),
  0x406: ascii("I"),
  0x407: overNarrow(ascii("I"), diaeresis),
  0x408: ascii("J"),
  0x409: LJE,
  0x40a: NJE,
  0x40b: TSHE,
  0x40c: withMarkOver(ascii("K"), acute),
  0x40d: withMarkOver(I_CYR, grave),
  0x40e: U_BREVE,
  0x40f: DZHE,

  // U+0410–U+042F — the Russian capitals.
  0x410: ascii("A"),
  0x411: BE,
  0x412: ascii("B"),
  0x413: GHE,
  0x414: DE,
  0x415: ascii("E"),
  0x416: ZHE,
  0x417: ZE,
  0x418: I_CYR,
  0x419: SHORT_I,
  0x41a: ascii("K"),
  0x41b: EL,
  0x41c: ascii("M"),
  0x41d: ascii("H"),
  0x41e: ascii("O"),
  0x41f: PE,
  0x420: ascii("P"),
  0x421: ascii("C"),
  0x422: ascii("T"),
  0x423: U_CYR,
  0x424: EF,
  0x425: ascii("X"),
  0x426: TSE,
  0x427: CHE,
  0x428: SHA,
  0x429: SHCHA,
  0x42a: HARD,
  0x42b: YERU,
  0x42c: SOFT,
  0x42d: E_CYR,
  0x42e: YU,
  0x42f: YA,

  // U+0430–U+044F — the Russian lowercase. Small capitals except where noted.
  0x430: ascii("a"),
  0x431: shape(
    0.46,
    [
      [0.36, 0.28],
      [0.24, 0.24],
      [0.14, 0.3],
      [0.1, 0.42]
    ],
    [
      [0.1, 0.42],
      [0.1, 0.66],
      [0.2, 0.75],
      [0.32, 0.75],
      [0.4, 0.66],
      [0.4, 0.52],
      [0.32, 0.45],
      [0.14, 0.45],
      [0.1, 0.5]
    ]
  ),
  0x432: asSmall(ascii("B")),
  0x433: asSmall(GHE),
  0x434: asSmall(DE),
  0x435: ascii("e"),
  0x436: asSmall(ZHE),
  0x437: asSmall(ZE),
  0x438: asSmall(I_CYR),
  0x439: asSmall(SHORT_I),
  0x43a: asSmall(ascii("K")),
  0x43b: asSmall(EL),
  0x43c: asSmall(ascii("M")),
  0x43d: asSmall(ascii("H")),
  0x43e: ascii("o"),
  0x43f: asSmall(PE),
  0x440: ascii("p"),
  0x441: ascii("c"),
  0x442: asSmall(ascii("T")),
  0x443: ascii("y"),
  0x444: asSmall(EF),
  0x445: ascii("x"),
  0x446: asSmall(TSE),
  0x447: asSmall(CHE),
  0x448: asSmall(SHA),
  0x449: asSmall(SHCHA),
  0x44a: asSmall(HARD),
  0x44b: asSmall(YERU),
  0x44c: asSmall(SOFT),
  0x44d: asSmall(E_CYR),
  0x44e: asSmall(YU),
  0x44f: asSmall(YA),

  // U+0450–U+045F — the lowercase of U+0400–U+040F.
  0x450: withMarkOverSmall(ascii("e"), grave),
  0x451: withMarkOverSmall(ascii("e"), diaeresis),
  0x452: asSmall(DJE),
  0x453: withMarkOverSmall(asSmall(GHE), acute),
  0x454: asSmall(UKR_IE),
  0x455: ascii("s"),
  0x456: ascii("i"),
  0x457: overNarrow(ascii("i"), diaeresis),
  0x458: ascii("j"),
  0x459: asSmall(LJE),
  0x45a: asSmall(NJE),
  0x45b: asSmall(TSHE),
  0x45c: withMarkOverSmall(asSmall(ascii("K")), acute),
  0x45d: withMarkOverSmall(asSmall(I_CYR), grave),
  0x45e: U_BREVE_SMALL,
  0x45f: asSmall(DZHE)
};

/** Cyrillic, as Type3 glyphs. */
export const CYRILLIC: Record<number, GlyphDef> = Object.fromEntries(
  Object.entries(SHAPES).map(([cp, form]) => [Number(cp), toGlyphDef(form)])
);

/** The code points this table defines, for the repertoire to state without importing it. */
export const CYRILLIC_CODE_POINTS: readonly number[] = Object.keys(SHAPES).map(Number);
