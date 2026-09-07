/**
 * Greek letterforms for the Type3 fallback — U+0384…U+03CE.
 *
 * Modern Greek, which is what turns up in a document this library is asked to write:
 * the monotonic alphabet with its accented forms, plus the three marks that stand
 * alone. Polytonic Greek and the archaic letters are deliberately absent — they need
 * mark stacking this renderer does not do, and a document that needs them needs a real
 * font, which `requiresEmbeddedFace` still asks for.
 *
 * Fourteen of the twenty-four capitals are Latin letters and are taken from the shared
 * stroke font unchanged: `Α Β Ε Ζ Η Ι Κ Μ Ν Ο Ρ Τ Υ Χ`. The ten that are not, and the
 * lowercase, are authored here in the same stroke-font space so their weight, cap
 * height and sidebearings match the borrowed ones exactly.
 *
 * Every accented vowel is composed rather than drawn: `ά` is `α` plus a tonos. One
 * consequence worth naming — on a capital, Greek sets the tonos to the *left* of the
 * letter (`Ά`), not above it. That is a positioning rule this scalar renderer cannot
 * express per-script, so the mark is centred above instead. It reads correctly as the
 * accented letter it is, which is the point of a fallback, and is one more reason the
 * discovery path still prefers a real face.
 *
 * @module
 */

import type { GlyphDef } from "@pdf/font/type3-glyphs";
import {
  acute,
  ascii,
  diaeresis,
  merge,
  oval,
  shape,
  toGlyphDef,
  withMarkOver,
  withMarkOverSmall,
  type StrokeShape
} from "@pdf/font/type3-letterforms";

// =============================================================================
// Capitals
// =============================================================================

/** Γ — a stem and a top arm. */
const GAMMA = shape(0.46, [
  [0.1, 0.75],
  [0.1, 0.15],
  [0.42, 0.15]
]);

/** Δ — a closed triangle. */
const DELTA = shape(0.56, [
  [0.04, 0.75],
  [0.28, 0.15],
  [0.52, 0.75],
  [0.04, 0.75]
]);

/** Θ — the `O` bowl with a crossbar. */
const THETA = merge(
  ascii("O"),
  shape(0.55, [
    [0.17, 0.45],
    [0.38, 0.45]
  ])
);

/** Λ — `Α` without its crossbar. */
const LAMBDA = shape(0.56, [
  [0.04, 0.75],
  [0.28, 0.15],
  [0.52, 0.75]
]);

/** Ξ — three arms, the middle one inset. */
const XI = shape(
  0.52,
  [
    [0.1, 0.15],
    [0.42, 0.15]
  ],
  [
    [0.14, 0.45],
    [0.38, 0.45]
  ],
  [
    [0.1, 0.75],
    [0.42, 0.75]
  ]
);

/** Π — two stems under a bridge. */
const PI_UPPER = shape(0.56, [
  [0.1, 0.75],
  [0.1, 0.15],
  [0.46, 0.15],
  [0.46, 0.75]
]);

/** Σ — a single folded stroke. */
const SIGMA_UPPER = shape(0.5, [
  [0.42, 0.15],
  [0.1, 0.15],
  [0.3, 0.45],
  [0.1, 0.75],
  [0.42, 0.75]
]);

/** Φ — a full-height stem through a bowl. */
const PHI_UPPER = merge(
  shape(0.58, [
    [0.29, 0.15],
    [0.29, 0.75]
  ]),
  shape(0.58, oval(0.29, 0.45, 0.21, 0.19))
);

/** Ψ — a stem through a cup. */
const PSI_UPPER = merge(
  shape(0.58, [
    [0.29, 0.15],
    [0.29, 0.75]
  ]),
  shape(0.58, [
    [0.08, 0.15],
    [0.08, 0.38],
    [0.16, 0.48],
    [0.29, 0.52],
    [0.42, 0.48],
    [0.5, 0.38],
    [0.5, 0.15]
  ])
);

/** Ω — an open bowl on two feet. */
const OMEGA = shape(0.58, [
  [0.06, 0.75],
  [0.2, 0.75],
  [0.14, 0.62],
  [0.1, 0.46],
  [0.14, 0.28],
  [0.29, 0.18],
  [0.44, 0.28],
  [0.48, 0.46],
  [0.44, 0.62],
  [0.38, 0.75],
  [0.52, 0.75]
]);

// =============================================================================
// Lowercase
// =============================================================================

/** α — a bowl with a stem down its right side. */
const ALPHA_LOWER = merge(
  shape(0.48, oval(0.24, 0.565, 0.14, 0.185)),
  shape(0.48, [
    [0.38, 0.4],
    [0.38, 0.7],
    [0.44, 0.75]
  ])
);

/** β — an ascending stem with two bowls, descending below the baseline. */
const BETA_LOWER = merge(
  shape(0.48, [
    [0.12, 0.95],
    [0.12, 0.28],
    [0.24, 0.2],
    [0.36, 0.26],
    [0.36, 0.4],
    [0.24, 0.47],
    [0.12, 0.47]
  ]),
  shape(0.48, [
    [0.12, 0.47],
    [0.28, 0.47],
    [0.4, 0.55],
    [0.4, 0.67],
    [0.28, 0.75],
    [0.12, 0.75]
  ])
);

/** γ — two strokes, the right one descending. */
const GAMMA_LOWER = shape(
  0.46,
  [
    [0.06, 0.38],
    [0.25, 0.72]
  ],
  [
    [0.46, 0.38],
    [0.14, 0.95]
  ]
);

/** δ — a bowl closed by a stroke curling over its top. */
const DELTA_LOWER = merge(
  shape(0.46, oval(0.25, 0.62, 0.15, 0.13)),
  shape(0.46, [
    [0.4, 0.4],
    [0.28, 0.38],
    [0.18, 0.42],
    [0.26, 0.49]
  ])
);

/** ε — two arcs meeting at a waist. */
const EPSILON_LOWER = shape(
  0.42,
  [
    [0.38, 0.42],
    [0.26, 0.38],
    [0.14, 0.44],
    [0.22, 0.55],
    [0.32, 0.55]
  ],
  [
    [0.22, 0.55],
    [0.13, 0.64],
    [0.24, 0.75],
    [0.38, 0.71]
  ]
);

/** ζ — a folded stroke with a descending tail. */
const ZETA_LOWER = merge(
  shape(0.42, [
    [0.12, 0.3],
    [0.38, 0.3]
  ]),
  shape(0.42, [
    [0.34, 0.3],
    [0.2, 0.5],
    [0.32, 0.62],
    [0.24, 0.86],
    [0.12, 0.93]
  ])
);

/** η — `n` with the right stem carried below the baseline. */
const ETA_LOWER = shape(0.46, [
  [0.1, 0.75],
  [0.1, 0.38],
  [0.24, 0.38],
  [0.36, 0.44],
  [0.38, 0.55],
  [0.38, 0.95]
]);

/** θ — a tall bowl with a crossbar. */
const THETA_LOWER = merge(
  shape(0.46, oval(0.24, 0.45, 0.15, 0.3)),
  shape(0.46, [
    [0.09, 0.45],
    [0.39, 0.45]
  ])
);

/** ι — a stem with a foot. */
const IOTA_LOWER = shape(0.24, [
  [0.12, 0.38],
  [0.12, 0.68],
  [0.2, 0.75]
]);

/** κ — a stem and two arms. */
const KAPPA_LOWER = shape(
  0.44,
  [
    [0.1, 0.38],
    [0.1, 0.75]
  ],
  [
    [0.4, 0.38],
    [0.12, 0.58]
  ],
  [
    [0.18, 0.54],
    [0.42, 0.75]
  ]
);

/** λ — an ascending stroke crossed by a shorter one. */
const LAMBDA_LOWER = shape(
  0.46,
  [
    [0.08, 0.15],
    [0.36, 0.75]
  ],
  [
    [0.3, 0.46],
    [0.06, 0.75]
  ]
);

/** μ — `u` with a descending left stem. */
const MU_LOWER = merge(
  shape(0.48, [
    [0.1, 0.38],
    [0.1, 0.95]
  ]),
  shape(0.48, [
    [0.1, 0.64],
    [0.18, 0.74],
    [0.3, 0.74],
    [0.38, 0.64],
    [0.38, 0.38]
  ])
);

/** ν — a rounded `v`. */
const NU_LOWER = shape(0.44, [
  [0.06, 0.38],
  [0.14, 0.64],
  [0.22, 0.75],
  [0.32, 0.6],
  [0.4, 0.4]
]);

/** ξ — a barred, twice-folded stroke with a tail. */
const XI_LOWER = merge(
  shape(0.42, [
    [0.12, 0.28],
    [0.38, 0.28]
  ]),
  shape(0.42, [
    [0.34, 0.28],
    [0.2, 0.44],
    [0.33, 0.53],
    [0.19, 0.64],
    [0.32, 0.76],
    [0.22, 0.9],
    [0.11, 0.93]
  ])
);

/** π — two stems under a bar. */
const PI_LOWER = merge(
  shape(0.5, [
    [0.06, 0.4],
    [0.46, 0.4]
  ]),
  shape(
    0.5,
    [
      [0.16, 0.4],
      [0.14, 0.75]
    ],
    [
      [0.34, 0.4],
      [0.36, 0.75]
    ]
  )
);

/** ρ — a descending stem with a bowl. */
const RHO_LOWER = merge(
  shape(0.46, [
    [0.12, 0.95],
    [0.12, 0.5]
  ]),
  shape(0.46, oval(0.26, 0.575, 0.14, 0.175))
);

/** ς — final sigma: a bowl opened at the top with a descending tail. */
const FINAL_SIGMA = shape(0.42, [
  [0.38, 0.42],
  [0.26, 0.38],
  [0.15, 0.45],
  [0.14, 0.6],
  [0.26, 0.71],
  [0.33, 0.85],
  [0.24, 0.94]
]);

/** σ — a bowl with a bar off its top right. */
const SIGMA_LOWER = merge(
  shape(0.48, oval(0.24, 0.565, 0.14, 0.185)),
  shape(0.48, [
    [0.28, 0.38],
    [0.46, 0.38]
  ])
);

/** τ — a bar over a stem with a foot. */
const TAU_LOWER = merge(
  shape(0.42, [
    [0.06, 0.4],
    [0.4, 0.4]
  ]),
  shape(0.42, [
    [0.24, 0.4],
    [0.24, 0.68],
    [0.32, 0.75]
  ])
);

/** υ — a cup. */
const UPSILON_LOWER = shape(0.46, [
  [0.1, 0.38],
  [0.1, 0.62],
  [0.2, 0.74],
  [0.32, 0.74],
  [0.4, 0.62],
  [0.4, 0.38]
]);

/** φ — a bowl on a stem that rises and descends. */
const PHI_LOWER = merge(
  shape(0.52, [
    [0.26, 0.3],
    [0.26, 0.95]
  ]),
  shape(0.52, oval(0.26, 0.565, 0.18, 0.175))
);

/** χ — two crossing strokes, both descending. */
const CHI_LOWER = shape(
  0.46,
  [
    [0.06, 0.38],
    [0.4, 0.95]
  ],
  [
    [0.4, 0.38],
    [0.06, 0.95]
  ]
);

/** ψ — a cup on a stem that rises and descends. */
const PSI_LOWER = merge(
  shape(0.52, [
    [0.26, 0.3],
    [0.26, 0.95]
  ]),
  shape(0.52, [
    [0.08, 0.38],
    [0.08, 0.62],
    [0.16, 0.72],
    [0.26, 0.75],
    [0.36, 0.72],
    [0.44, 0.62],
    [0.44, 0.38]
  ])
);

/** ω — two cups sharing a wall. */
const OMEGA_LOWER = shape(0.56, [
  [0.06, 0.38],
  [0.06, 0.66],
  [0.14, 0.75],
  [0.22, 0.72],
  [0.26, 0.6],
  [0.3, 0.72],
  [0.38, 0.75],
  [0.46, 0.66],
  [0.46, 0.38]
]);

// =============================================================================
// Marks that stand alone
// =============================================================================

/** ΄ GREEK TONOS. */
const TONOS = shape(0.3, ...acute(0.1, 0.26));

/** ΅ GREEK DIALYTIKA TONOS. */
const DIALYTIKA_TONOS = merge(
  shape(0.35, ...diaeresis(0.06, 0.24)),
  shape(0.35, ...acute(0.1, 0.1))
);

/** · GREEK ANO TELEIA — a middle dot. */
const ANO_TELEIA = shape(0.26, [
  [0.13, 0.55],
  [0.14, 0.55]
]);

// =============================================================================
// Table
// =============================================================================

/** Add both marks a Greek vowel can carry at once. */
function withDiaeresisAndTonos(base: StrokeShape): StrokeShape {
  return merge(
    base,
    shape(base.w, ...diaeresis(base.w * 0.5 - 0.09, 0.29)),
    shape(base.w, ...acute(base.w * 0.5 - 0.03, 0.17))
  );
}

const SHAPES: Record<number, StrokeShape> = {
  // Marks and accented capitals.
  0x384: TONOS,
  0x385: DIALYTIKA_TONOS,
  0x386: withMarkOver(ascii("A"), acute),
  0x387: ANO_TELEIA,
  0x388: withMarkOver(ascii("E"), acute),
  0x389: withMarkOver(ascii("H"), acute),
  0x38a: withMarkOver(ascii("I"), acute),
  0x38c: withMarkOver(ascii("O"), acute),
  0x38e: withMarkOver(ascii("Y"), acute),
  0x38f: withMarkOver(OMEGA, acute),

  // ΐ, then the capitals.
  0x390: withDiaeresisAndTonos(IOTA_LOWER),
  0x391: ascii("A"),
  0x392: ascii("B"),
  0x393: GAMMA,
  0x394: DELTA,
  0x395: ascii("E"),
  0x396: ascii("Z"),
  0x397: ascii("H"),
  0x398: THETA,
  0x399: ascii("I"),
  0x39a: ascii("K"),
  0x39b: LAMBDA,
  0x39c: ascii("M"),
  0x39d: ascii("N"),
  0x39e: XI,
  0x39f: ascii("O"),
  0x3a0: PI_UPPER,
  0x3a1: ascii("P"),
  0x3a3: SIGMA_UPPER,
  0x3a4: ascii("T"),
  0x3a5: ascii("Y"),
  0x3a6: PHI_UPPER,
  0x3a7: ascii("X"),
  0x3a8: PSI_UPPER,
  0x3a9: OMEGA,

  // Accented lowercase, then the lowercase.
  0x3aa: withMarkOver(ascii("I"), diaeresis),
  0x3ab: withMarkOver(ascii("Y"), diaeresis),
  0x3ac: withMarkOverSmall(ALPHA_LOWER, acute),
  0x3ad: withMarkOverSmall(EPSILON_LOWER, acute),
  0x3ae: withMarkOverSmall(ETA_LOWER, acute),
  0x3af: withMarkOverSmall(IOTA_LOWER, acute),
  0x3b0: withDiaeresisAndTonos(UPSILON_LOWER),
  0x3b1: ALPHA_LOWER,
  0x3b2: BETA_LOWER,
  0x3b3: GAMMA_LOWER,
  0x3b4: DELTA_LOWER,
  0x3b5: EPSILON_LOWER,
  0x3b6: ZETA_LOWER,
  0x3b7: ETA_LOWER,
  0x3b8: THETA_LOWER,
  0x3b9: IOTA_LOWER,
  0x3ba: KAPPA_LOWER,
  0x3bb: LAMBDA_LOWER,
  0x3bc: MU_LOWER,
  0x3bd: NU_LOWER,
  0x3be: XI_LOWER,
  0x3bf: ascii("o"),
  0x3c0: PI_LOWER,
  0x3c1: RHO_LOWER,
  0x3c2: FINAL_SIGMA,
  0x3c3: SIGMA_LOWER,
  0x3c4: TAU_LOWER,
  0x3c5: UPSILON_LOWER,
  0x3c6: PHI_LOWER,
  0x3c7: CHI_LOWER,
  0x3c8: PSI_LOWER,
  0x3c9: OMEGA_LOWER,
  0x3ca: withMarkOverSmall(IOTA_LOWER, diaeresis),
  0x3cb: withMarkOverSmall(UPSILON_LOWER, diaeresis),
  0x3cc: withMarkOverSmall(ascii("o"), acute),
  0x3cd: withMarkOverSmall(UPSILON_LOWER, acute),
  0x3ce: withMarkOverSmall(OMEGA_LOWER, acute)
};

/** Greek, as Type3 glyphs. */
export const GREEK: Record<number, GlyphDef> = Object.fromEntries(
  Object.entries(SHAPES).map(([cp, form]) => [Number(cp), toGlyphDef(form)])
);

/** The code points this table defines, for the repertoire to state without importing it. */
export const GREEK_CODE_POINTS: readonly number[] = Object.keys(SHAPES).map(Number);
