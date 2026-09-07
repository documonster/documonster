/**
 * Type3 letterforms for Latin Extended-B — U+0180…U+024F, complete.
 *
 * Excluded from the first pass as "novel shapes rather than compositions, so there is
 * nothing to reuse". That was true of a minority and wrong about the block. Counted
 * properly, 178 of these 208 code points are a letter this directory already draws under
 * a transform it already has:
 *
 *   - **Marks.** `Ǎ ǎ Ǐ ǐ Ǒ ǒ Ǔ ǔ` are carons — the pinyin tone letters, which is the
 *     most-used part of the block by some distance. `Ǖ ǖ Ǘ ǘ Ǚ ǚ Ǜ ǜ` stack a second mark
 *     over the diaeresis, `Ȁ ȁ Ȃ ȃ` take a double grave or an inverted breve, and the
 *     `Ȧ Ȩ Ȯ Ȳ` group takes dots and macrons.
 *   - **Mirrors and turns.** `Ɔ` is a reversed `C`, `Ǝ` a reversed `E`, `Ƨ` a reversed `S`,
 *     `Ɯ` a turned `m`, `Ʌ` a turned `V`, `ǝ`/`Ə` a turned `e`.
 *   - **Bars and strokes.** `Ƀ Ɨ Ɵ Ƶ Ⱦ Ɇ Ɉ Ɍ Ɏ ƚ ƀ` and a dozen more.
 *   - **Hooks.** `Ɓ Ƈ Ɗ Ƒ Ɠ Ƙ Ɲ Ƥ Ƭ Ƴ` — the African orthographies (Hausa, Fula, Ewe,
 *     Serer), which are why this block is not merely a phonetician's alphabet.
 *   - **Digraphs.** `Ǆ ǅ ǆ Ǉ ǈ ǉ Ǌ ǋ ǌ Ǳ ǲ ǳ Ȣ ȣ ȸ ȹ ƕ Ƕ` are two letters set as one.
 *
 * That leaves 30 that genuinely are new shapes — the ezh family `Ʒ ƹ ƺ Ƹ ƻ`, yogh `Ȝ ȝ`,
 * wynn `Ƿ ƿ`, gha `Ƣ ƣ`, the tone letters `Ƽ ƽ ƾ`, the glottal stops `Ɂ ɂ` and a few
 * others. They are drawn here, because the alternative is a block with holes in it: the
 * repertoire is stated as ranges and checked against the glyph tables code point for code
 * point, so 30 gaps would mean 30 more ranges and a reader who cannot tell which absences
 * were decisions.
 *
 * @module
 */

import type { GlyphDef } from "@pdf/font/type3-glyphs";
import {
  acute,
  ascii,
  asSmall,
  caron,
  cedilla,
  diaeresis,
  dotAbove,
  grave,
  ligature,
  macron,
  merge,
  mirrorX,
  ogonek,
  oval,
  ring,
  rotate180,
  shape,
  tilde,
  toGlyphDef,
  withBar,
  withMarkOver,
  withMarkOverSmall,
  withMarkStack,
  withMarkUnder,
  withTopHook,
  type Path,
  type StrokeShape
} from "@pdf/font/type3-letterforms";

// =============================================================================
// Shorthands
// =============================================================================

type Mark = (x: number, y: number) => Path[];

const cap = (ch: string, mark: Mark): StrokeShape => withMarkOver(ascii(ch), mark);
const low = (ch: string, mark: Mark): StrokeShape => withMarkOverSmall(ascii(ch), mark);
/** A lowercase letter with an ascender takes the capital band — see the Extended-A table. */
const tallLow = (ch: string, mark: Mark): StrokeShape => withMarkOver(ascii(ch), mark);
const hook = (ch: string): StrokeShape => withTopHook(ascii(ch));
const turned = (ch: string): StrokeShape => rotate180(ascii(ch));
const reversed = (ch: string): StrokeShape => mirrorX(ascii(ch));

/** A bar across a letter, placed as a fraction of its own box. */
function barred(ch: string, atY: number, from = -0.04, to = 0.04): StrokeShape {
  const base = ascii(ch);
  const points = base.d.flat();
  const midX = (Math.min(...points.map(([x]) => x)) + Math.max(...points.map(([x]) => x))) / 2;
  return withBar(base, midX + from - 0.06, atY, midX + to + 0.06, atY);
}

/** An inverted breve, as on `Ȃ ȃ Ȇ ȇ`. Standing on `y`. */
function invertedBreve(x: number, y: number): Path[] {
  return [
    [
      [x, y - 0.07],
      [x + 0.03, y - 0.02],
      [x + 0.08, y],
      [x + 0.13, y - 0.02],
      [x + 0.16, y - 0.07]
    ]
  ];
}

/** A double grave, as on `Ȁ ȁ`. Standing on `y`. */
function doubleGrave(x: number, y: number): Path[] {
  return [...grave(x - 0.01, y), ...grave(x + 0.09, y)];
}

/** A retroflex or palatal hook curling right off the foot of a letter. */
function withFootHook(base: StrokeShape, atX: number): StrokeShape {
  return merge(
    base,
    shape(base.w + 0.04, [
      [atX, 0.7],
      [atX + 0.05, 0.76],
      [atX + 0.02, 0.82]
    ])
  );
}

// =============================================================================
// The shapes with no letter to derive them from
// =============================================================================

/** Ʒ — ezh: a flat top over a descending tail. */
const EZH = shape(0.46, [
  [0.08, 0.15],
  [0.4, 0.15],
  [0.18, 0.45],
  [0.34, 0.5],
  [0.36, 0.65],
  [0.24, 0.75],
  [0.1, 0.71]
]);

/** ʒ — ezh at x-height, descending. */
const EZH_LOW = shape(0.42, [
  [0.06, 0.38],
  [0.36, 0.38],
  [0.16, 0.62],
  [0.32, 0.68],
  [0.32, 0.86],
  [0.2, 0.94],
  [0.08, 0.9]
]);

/** Ȝ — yogh: like a `3` whose tail descends. */
const YOGH = shape(0.44, [
  [0.1, 0.2],
  [0.24, 0.15],
  [0.38, 0.22],
  [0.34, 0.36],
  [0.2, 0.42],
  [0.36, 0.5],
  [0.36, 0.66],
  [0.22, 0.75],
  [0.08, 0.7]
]);

const YOGH_LOW = asSmall(YOGH);

/** Ƿ — wynn: a narrow bowl on a stem, the runic `w`. */
const WYNN = shape(
  0.46,
  [
    [0.1, 0.75],
    [0.1, 0.15]
  ],
  [
    [0.1, 0.15],
    [0.28, 0.15],
    [0.38, 0.24],
    [0.38, 0.38],
    [0.28, 0.46],
    [0.1, 0.46]
  ]
);

/** Ƣ — gha: an `o` joined to a descending hook. */
const GHA = merge(
  shape(0.5, oval(0.24, 0.45, 0.14, 0.3)),
  shape(0.5, [
    [0.38, 0.45],
    [0.44, 0.6],
    [0.44, 0.75]
  ])
);

/** Ɂ — the glottal stop: a question mark's bowl on a stem. */
const GLOTTAL = shape(0.4, [
  [0.08, 0.24],
  [0.16, 0.15],
  [0.3, 0.18],
  [0.32, 0.3],
  [0.2, 0.42],
  [0.2, 0.75]
]);

/** Ƽ — a tone letter: a `5` with a stroke. */
const TONE_FIVE = withBar(ascii("5"), 0.04, 0.55, 0.24, 0.42);

/** ƍ — a turned delta, drawn as a turned `d` bowl. */
const TURNED_DELTA = rotate180(ascii("d"));

// =============================================================================
// Table
// =============================================================================

const SHAPES: Record<number, StrokeShape> = {
  0x0180: barred("b", 0.3),
  0x0181: hook("B"),
  0x0182: barred("B", 0.15),
  0x0183: barred("b", 0.18),
  0x0184: shape(
    0.44,
    [
      [0.1, 0.15],
      [0.1, 0.75]
    ],
    [
      [0.1, 0.15],
      [0.36, 0.15]
    ],
    [
      [0.1, 0.45],
      [0.3, 0.45]
    ]
  ),
  0x0185: asSmall(
    shape(
      0.44,
      [
        [0.1, 0.15],
        [0.1, 0.75]
      ],
      [
        [0.1, 0.15],
        [0.36, 0.15]
      ],
      [
        [0.1, 0.45],
        [0.3, 0.45]
      ]
    )
  ),
  0x0186: reversed("C"),
  0x0187: hook("C"),
  0x0188: withTopHook(ascii("c")),
  0x0189: withBar(ascii("D"), 0.02, 0.45, 0.22, 0.45),
  0x018a: hook("D"),
  0x018b: barred("d", 0.2),
  0x018c: barred("d", 0.26),
  0x018d: TURNED_DELTA,
  0x018e: reversed("E"),
  0x018f: rotate180(ascii("e")),
  0x0190: mirrorX(EZH),
  0x0191: hook("F"),
  0x0192: withBar(ascii("f"), 0.02, 0.62, 0.24, 0.62),
  0x0193: hook("G"),
  0x0194: shape(
    0.46,
    [
      [0.06, 0.15],
      [0.24, 0.5]
    ],
    [
      [0.44, 0.15],
      [0.16, 0.9]
    ]
  ),
  0x0195: ligature("h", "v", 0.06),
  0x0196: shape(0.26, [
    [0.13, 0.15],
    [0.13, 0.66],
    [0.2, 0.75]
  ]),
  0x0197: barred("I", 0.45, -0.1, 0.1),
  0x0198: hook("K"),
  0x0199: withTopHook(ascii("k")),
  0x019a: barred("l", 0.5, -0.09, 0.09),
  0x019b: withBar(
    shape(
      0.46,
      [
        [0.08, 0.15],
        [0.36, 0.75]
      ],
      [
        [0.3, 0.46],
        [0.06, 0.75]
      ]
    ),
    0.06,
    0.3,
    0.26,
    0.3
  ),
  0x019c: turned("m"),
  0x019d: withTopHook(ascii("N")),
  0x019e: merge(
    ascii("n"),
    shape(0.45, [
      [0.38, 0.75],
      [0.38, 0.95]
    ])
  ),
  0x019f: barred("O", 0.45, -0.12, 0.12),
  0x01a2: GHA,
  0x01a3: asSmall(GHA),
  0x01a4: hook("P"),
  0x01a5: withTopHook(ascii("p")),
  0x01a6: merge(
    ascii("R"),
    shape(0.5, [
      [0.1, 0.15],
      [0.02, 0.1]
    ])
  ),
  0x01a7: reversed("S"),
  0x01a8: reversed("s"),
  0x01a9: shape(0.5, [
    [0.42, 0.15],
    [0.1, 0.15],
    [0.3, 0.45],
    [0.1, 0.75],
    [0.42, 0.75]
  ]),
  0x01aa: merge(
    shape(0.44, [
      [0.34, 0.38],
      [0.14, 0.38],
      [0.3, 0.56],
      [0.12, 0.75]
    ]),
    shape(0.44, [
      [0.12, 0.75],
      [0.3, 0.9]
    ])
  ),
  0x01ab: withFootHook(ascii("t"), 0.24),
  0x01ac: hook("T"),
  0x01ad: withFootHook(ascii("t"), 0.26),
  0x01ae: withFootHook(ascii("T"), 0.24),
  0x01b1: shape(0.5, [
    [0.08, 0.15],
    [0.08, 0.5],
    [0.16, 0.68],
    [0.26, 0.75],
    [0.36, 0.68],
    [0.44, 0.5],
    [0.44, 0.15]
  ]),
  0x01b2: withTopHook(ascii("v")),
  0x01b3: hook("Y"),
  0x01b4: withTopHook(ascii("y")),
  0x01b5: barred("Z", 0.45, -0.1, 0.1),
  0x01b6: barred("z", 0.56, -0.09, 0.09),
  0x01b7: EZH,
  0x01b8: mirrorX(EZH),
  0x01b9: mirrorX(EZH_LOW),
  0x01ba: EZH_LOW,
  0x01bb: withBar(ascii("2"), 0.04, 0.55, 0.24, 0.55),
  0x01bc: TONE_FIVE,
  0x01bd: asSmall(TONE_FIVE),
  0x01be: shape(
    0.36,
    [
      [0.1, 0.38],
      [0.26, 0.75]
    ],
    [
      [0.04, 0.56],
      [0.32, 0.56]
    ]
  ),
  0x01bf: asSmall(WYNN),
  0x01c0: shape(0.24, [
    [0.12, 0.15],
    [0.12, 0.75]
  ]),
  0x01c1: shape(
    0.36,
    [
      [0.1, 0.15],
      [0.1, 0.75]
    ],
    [
      [0.24, 0.15],
      [0.24, 0.75]
    ]
  ),
  0x01c2: merge(
    shape(
      0.36,
      [
        [0.1, 0.15],
        [0.1, 0.75]
      ],
      [
        [0.24, 0.15],
        [0.24, 0.75]
      ]
    ),
    shape(0.36, [
      [0.02, 0.45],
      [0.32, 0.45]
    ])
  ),
  0x01c3: merge(
    shape(0.24, [
      [0.12, 0.15],
      [0.12, 0.55]
    ]),
    shape(0.24, [
      [0.12, 0.7],
      [0.13, 0.7]
    ])
  ),
  0x01c4: merge(ligature("D", "Z", 0.06), shape(1.0, ...caron(0.78, 0.05))),
  0x01c5: merge(ligature("D", "z", 0.06), shape(0.94, ...caron(0.72, 0.28))),
  0x01c6: merge(ligature("d", "z", 0.06), shape(0.92, ...caron(0.7, 0.28))),
  0x01c7: ligature("L", "J", 0.06),
  0x01c8: ligature("L", "j", 0.06),
  0x01c9: ligature("l", "j", 0.06),
  0x01ca: ligature("N", "J", 0.06),
  0x01cb: ligature("N", "j", 0.06),
  0x01cc: ligature("n", "j", 0.06),
  0x01cd: cap("A", caron),
  0x01ce: low("a", caron),
  0x01cf: cap("I", caron),
  0x01d0: tallLow("i", caron),
  0x01d1: cap("O", caron),
  0x01d2: low("o", caron),
  0x01d3: cap("U", caron),
  0x01d4: low("u", caron),
  0x01d5: withMarkStack(cap("U", diaeresis), diaeresis, macron, false),
  0x01d6: withMarkStack(low("u", diaeresis), diaeresis, macron, true),
  0x01d7: withMarkStack(cap("U", diaeresis), diaeresis, acute, false),
  0x01d8: withMarkStack(low("u", diaeresis), diaeresis, acute, true),
  0x01d9: withMarkStack(cap("U", diaeresis), diaeresis, caron, false),
  0x01da: withMarkStack(low("u", diaeresis), diaeresis, caron, true),
  0x01db: withMarkStack(cap("U", diaeresis), diaeresis, grave, false),
  0x01dc: withMarkStack(low("u", diaeresis), diaeresis, grave, true),
  0x01dd: rotate180(ascii("e")),
  0x01de: withMarkStack(cap("A", diaeresis), diaeresis, macron, false),
  0x01df: withMarkStack(low("a", diaeresis), diaeresis, macron, true),
  0x01e0: withMarkStack(cap("A", dotAbove), dotAbove, macron, false),
  0x01e1: withMarkStack(low("a", dotAbove), dotAbove, macron, true),
  0x01e2: withMarkOver(ligature("A", "E", 0.1), macron),
  0x01e3: withMarkOverSmall(ligature("a", "e", 0.08), macron),
  0x01e4: barred("G", 0.45, -0.1, 0.1),
  0x01e5: barred("g", 0.56, -0.09, 0.09),
  0x01e6: cap("G", caron),
  0x01e7: low("g", caron),
  0x01e8: cap("K", caron),
  0x01e9: tallLow("k", caron),
  0x01ea: withMarkUnder(ascii("O"), ogonek),
  0x01eb: withMarkUnder(ascii("o"), ogonek),
  0x01ec: withMarkOver(withMarkUnder(ascii("O"), ogonek), macron),
  0x01ed: withMarkOverSmall(withMarkUnder(ascii("o"), ogonek), macron),
  0x01ee: withMarkOver(EZH, caron),
  0x01ef: withMarkOverSmall(EZH_LOW, caron),
  0x01f0: tallLow("j", caron),
  0x01f1: ligature("D", "Z", 0.06),
  0x01f2: ligature("D", "z", 0.06),
  0x01f3: ligature("d", "z", 0.06),
  0x01f4: cap("G", acute),
  0x01f5: low("g", acute),
  0x01f6: ligature("H", "V", 0.06),
  0x01f7: WYNN,
  0x01f8: cap("N", grave),
  0x01f9: low("n", grave),
  0x01fa: withMarkStack(cap("A", ring), ring, acute, false),
  0x01fb: withMarkStack(low("a", ring), ring, acute, true),
  0x01fc: withMarkOver(ligature("A", "E", 0.1), acute),
  0x01fd: withMarkOverSmall(ligature("a", "e", 0.08), acute),
  0x01fe: withMarkOver(barred("O", 0.45, -0.14, 0.14), acute),
  0x01ff: withMarkOverSmall(barred("o", 0.56, -0.12, 0.12), acute),
  0x0200: cap("A", doubleGrave),
  0x0201: low("a", doubleGrave),
  0x0202: cap("A", invertedBreve),
  0x0203: low("a", invertedBreve),
  0x0204: cap("E", doubleGrave),
  0x0205: low("e", doubleGrave),
  0x0206: cap("E", invertedBreve),
  0x0207: low("e", invertedBreve),
  0x0208: cap("I", doubleGrave),
  0x0209: tallLow("i", doubleGrave),
  0x020a: cap("I", invertedBreve),
  0x020b: tallLow("i", invertedBreve),
  0x020c: cap("O", doubleGrave),
  0x020d: low("o", doubleGrave),
  0x020e: cap("O", invertedBreve),
  0x020f: low("o", invertedBreve),
  0x0210: cap("R", doubleGrave),
  0x0211: low("r", doubleGrave),
  0x0212: cap("R", invertedBreve),
  0x0213: low("r", invertedBreve),
  0x0214: cap("U", doubleGrave),
  0x0215: low("u", doubleGrave),
  0x0216: cap("U", invertedBreve),
  0x0217: low("u", invertedBreve),
  0x021c: YOGH,
  0x021d: YOGH_LOW,
  0x021e: cap("H", caron),
  0x021f: tallLow("h", caron),
  0x0220: merge(
    withTopHook(ascii("N")),
    shape(0.52, [
      [0.42, 0.75],
      [0.42, 0.9]
    ])
  ),
  0x0221: withFootHook(ascii("d"), 0.44),
  0x0222: ligature("O", "U", 0.12),
  0x0223: ligature("o", "u", 0.1),
  0x0224: hook("Z"),
  0x0225: withTopHook(ascii("z")),
  0x0226: cap("A", dotAbove),
  0x0227: low("a", dotAbove),
  0x0228: withMarkUnder(ascii("E"), cedilla),
  0x0229: withMarkUnder(ascii("e"), cedilla),
  0x022a: withMarkStack(cap("O", diaeresis), diaeresis, macron, false),
  0x022b: withMarkStack(low("o", diaeresis), diaeresis, macron, true),
  0x022c: withMarkStack(cap("O", tilde), tilde, macron, false),
  0x022d: withMarkStack(low("o", tilde), tilde, macron, true),
  0x022e: cap("O", dotAbove),
  0x022f: low("o", dotAbove),
  0x0230: withMarkStack(cap("O", dotAbove), dotAbove, macron, false),
  0x0231: withMarkStack(low("o", dotAbove), dotAbove, macron, true),
  0x0232: cap("Y", macron),
  0x0233: low("y", macron),
  0x0234: withFootHook(ascii("l"), 0.22),
  0x0235: withFootHook(ascii("n"), 0.4),
  0x0236: withFootHook(ascii("t"), 0.28),
  0x0237: shape(0.26, [
    [0.14, 0.38],
    [0.14, 0.86],
    [0.06, 0.94]
  ]),
  0x0238: ligature("d", "b", 0.06),
  0x0239: ligature("q", "p", 0.06),
  0x023a: barred("A", 0.5, -0.14, 0.14),
  0x023b: barred("C", 0.45, -0.12, 0.12),
  0x023c: barred("c", 0.56, -0.1, 0.1),
  0x023d: withBar(ascii("L"), 0.04, 0.3, 0.24, 0.3),
  0x023e: withBar(ascii("T"), 0.06, 0.6, 0.42, 0.32),
  0x023f: merge(
    ascii("s"),
    shape(0.4, [
      [0.32, 0.75],
      [0.36, 0.88]
    ])
  ),
  0x0240: merge(
    ascii("z"),
    shape(0.4, [
      [0.34, 0.75],
      [0.38, 0.88]
    ])
  ),
  0x0241: GLOTTAL,
  0x0242: asSmall(GLOTTAL),
  0x0243: barred("B", 0.45, -0.1, 0.1),
  0x0244: barred("U", 0.45, -0.1, 0.1),
  0x0245: turned("V"),
  0x0246: barred("E", 0.45, -0.1, 0.1),
  0x0247: barred("e", 0.56, -0.09, 0.09),
  0x0248: barred("J", 0.45, -0.1, 0.1),
  0x0249: barred("j", 0.56, -0.08, 0.08),
  0x024a: merge(
    ascii("Q"),
    shape(0.55, [
      [0.4, 0.66],
      [0.5, 0.78]
    ])
  ),
  0x024b: merge(
    ascii("q"),
    shape(0.47, [
      [0.4, 0.86],
      [0.48, 0.94]
    ])
  ),
  0x024c: barred("R", 0.45, -0.1, 0.1),
  0x024d: barred("r", 0.56, -0.08, 0.08),
  0x024e: barred("Y", 0.55, -0.1, 0.1),
  0x024f: barred("y", 0.62, -0.09, 0.09)
};

/** Latin Extended-B, as Type3 glyphs. */
export const LATIN_B: Record<number, GlyphDef> = Object.fromEntries(
  Object.entries(SHAPES).map(([cp, form]) => [Number(cp), toGlyphDef(form)])
);

/** The code points this table defines, for the repertoire to state without importing it. */
export const LATIN_B_CODE_POINTS: readonly number[] = Object.keys(SHAPES).map(Number);

// Four code points of this block are defined elsewhere and deliberately not repeated
// here: `Ơ ơ Ư ư` (U+01A0, U+01A1, U+01AF, U+01B0) sit with the Vietnamese tone letters
// they carry, and `Ș ș Ț ț` (U+0218–U+021B) with Romanian in the Extended-A table.
