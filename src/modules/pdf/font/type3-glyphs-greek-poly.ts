/**
 * Type3 letterforms for polytonic Greek — U+1F00…U+1FFE.
 *
 * Excluded from the first pass on the same mistaken reason as Vietnamese: that this
 * renderer "cannot position one mark over another". It cannot do so at *draw* time, for a
 * decomposed sequence. Every code point in this block is precomposed, and a precomposed
 * glyph is one outline this directory authors in full — including one carrying a breathing,
 * an accent and an iota subscript at once.
 *
 * ## Generated from the block's own structure, not written out
 *
 * Unicode lays this block out as a product, and hand-writing 233 rows would obscure that
 * while inviting a transcription error in every one. So the table below is built from the
 * grid Unicode used:
 *
 *     U+1F00…U+1F7D   vowel × breathing × accent, in rows of eight
 *     U+1F80…U+1FB4   the same, with an iota subscript
 *     U+1FB6…U+1FFE   the long/short marks, the standalone breathings and accents
 *
 * A row of eight is `smooth, rough` × `none, grave, acute, circumflex`, which is exactly
 * how {@link row} reads. Where Unicode leaves a cell empty — `ε` and `ο` have no
 * circumflex, `ῂ` has no unaccented form — the generator skips it, and
 * `type3-repertoire.ts` states the same gaps as ranges. The two are checked against each
 * other code point for code point, so a mistake in either fails a test rather than
 * shipping a hole.
 *
 * ## What the marks are
 *
 * Three of the four are already drawn: `acute`, `grave` and the `circumflex` here is the
 * Greek one, a tilde-like `perispomeni`. The breathings are new — a smooth breathing is an
 * apostrophe-shaped comma, a rough breathing its mirror — and the iota subscript is a
 * short stroke below. A vowel can carry a breathing *and* an accent, which is the stack
 * the block needs and the reason it is reachable at all.
 *
 * Absent: nothing. The block is complete, which is why it can be stated as three ranges.
 *
 * @module
 */

import type { GlyphDef } from "@pdf/font/type3-glyphs";
import { GREEK_SHAPES } from "@pdf/font/type3-glyphs-greek";
import {
  BASE,
  CAP_TOP,
  X_TOP,
  acute,
  grave,
  macron,
  merge,
  shape,
  toGlyphDef,
  type Path,
  type StrokeShape
} from "@pdf/font/type3-letterforms";

// =============================================================================
// The marks this block adds
// =============================================================================

/** Psili — smooth breathing, a raised comma opening right. Standing on `y`. */
function smooth(x: number, y: number): Path[] {
  return [
    [
      [x + 0.02, y - 0.09],
      [x + 0.07, y - 0.1],
      [x + 0.09, y - 0.06],
      [x + 0.05, y - 0.02],
      [x + 0.02, y]
    ]
  ];
}

/** Dasia — rough breathing, the mirror of the smooth one. Standing on `y`. */
function rough(x: number, y: number): Path[] {
  return [
    [
      [x + 0.08, y - 0.09],
      [x + 0.03, y - 0.1],
      [x + 0.01, y - 0.06],
      [x + 0.05, y - 0.02],
      [x + 0.08, y]
    ]
  ];
}

/** Perispomeni — the Greek circumflex, drawn as a tilde. Standing on `y`. */
function perispomeni(x: number, y: number): Path[] {
  return [
    [
      [x, y - 0.02],
      [x + 0.04, y - 0.07],
      [x + 0.1, y - 0.02],
      [x + 0.14, y - 0.07]
    ]
  ];
}

/** Ypogegrammeni — the iota subscript, a short stroke under the letter. */
function iotaBelow(x: number, y: number): Path[] {
  return [
    [
      [x + 0.06, y],
      [x + 0.06, y + 0.11]
    ]
  ];
}

/** Vrachy — the short mark, a breve standing on `y`. */
function vrachy(x: number, y: number): Path[] {
  return [
    [
      [x, y - 0.07],
      [x + 0.04, y - 0.02],
      [x + 0.08, y],
      [x + 0.12, y - 0.02],
      [x + 0.16, y - 0.07]
    ]
  ];
}

type Mark = (x: number, y: number) => Path[];

/** Heights of the marks that can carry another above them. */
const HEIGHT = new Map<Mark, number>([
  [smooth, 0.1],
  [rough, 0.1],
  [acute, 0.1],
  [grave, 0.1],
  [perispomeni, 0.07]
]);

// =============================================================================
// Composition
// =============================================================================

/** How far a mark stands above the letter, and a second mark above the first. */
const CLEARANCE = 0.1;
const STACK = 0.09;

/** Whether a Greek letter is a capital, which decides the band its marks sit in. */
function isCapital(codePoint: number): boolean {
  return codePoint >= 0x0386 && codePoint <= 0x03ab;
}

/**
 * Put a breathing, an accent, or both, over a Greek letter — and an iota under it.
 *
 * The breathing always sits closer to the letter; the accent stacks above it. Where there
 * is no breathing the accent takes the lower position, which is what makes `ά` and `ἄ`
 * differ by a mark rather than by a height.
 */
function compose(
  codePoint: number,
  breathing: Mark | null,
  accent: Mark | null,
  iota: boolean
): StrokeShape {
  const base = GREEK_SHAPES[codePoint];
  if (base === undefined) {
    throw new Error(`polytonic Greek: no base shape for U+${codePoint.toString(16)}`);
  }
  const top = isCapital(codePoint) ? CAP_TOP : X_TOP;
  const x = base.w * 0.5 - 0.03;
  const parts: StrokeShape[] = [base];

  const lower = breathing ?? accent;
  if (lower !== null) {
    parts.push(shape(base.w, ...lower(x, top - CLEARANCE)));
  }
  if (breathing !== null && accent !== null) {
    const stackAt = top - CLEARANCE - HEIGHT.get(breathing)! - STACK;
    parts.push(shape(base.w, ...accent(x, stackAt)));
  }
  if (iota) {
    parts.push(shape(base.w, ...iotaBelow(x, BASE + 0.02)));
  }
  return merge(...parts);
}

/**
 * One case-pair of a vowel's breathing/accent grid.
 *
 * `cells` is how many the case takes — eight where the vowel has a circumflex, six where
 * it does not — and the uppercase half always begins eight code points after the
 * lowercase, *not* after the last assigned cell. Unicode leaves the two-cell hole in place
 * for `ε` and `ο`, and generating the uppercase from `start + cells` instead of
 * `start + 8` put every capital of those two rows two code points early.
 */
function pair(
  start: number,
  lower: number,
  upper: number,
  cells: 6 | 8,
  iota = false
): Record<number, StrokeShape> {
  const accents: ReadonlyArray<Mark | null> = [null, grave, acute, perispomeni];
  const out: Record<number, StrokeShape> = {};
  for (let index = 0; index < cells; index++) {
    const breathing = index % 2 === 0 ? smooth : rough;
    const accent = accents[Math.floor(index / 2)];
    out[start + index] = compose(lower, breathing, accent, iota);
    out[start + 8 + index] = compose(upper, breathing, accent, iota);
  }
  return out;
}

// =============================================================================
// Table
// =============================================================================

const ALPHA = 0x03b1;
const ALPHA_UPPER = 0x0391;
const EPSILON = 0x03b5;
const EPSILON_UPPER = 0x0395;
const ETA = 0x03b7;
const ETA_UPPER = 0x0397;
const IOTA = 0x03b9;
const IOTA_UPPER = 0x0399;
const OMICRON = 0x03bf;
const OMICRON_UPPER = 0x039f;
const UPSILON = 0x03c5;
const UPSILON_UPPER = 0x03a5;
const OMEGA = 0x03c9;
const OMEGA_UPPER = 0x03a9;
const RHO = 0x03c1;
const RHO_UPPER = 0x03a1;

const SHAPES: Record<number, StrokeShape> = {
  // U+1F00–U+1F6F: vowel × breathing × accent. `ε` and `ο` take no circumflex, so their
  // grids are six wide and Unicode leaves the remaining two cells of each case unassigned.
  ...pair(0x1f00, ALPHA, ALPHA_UPPER, 8),
  ...pair(0x1f10, EPSILON, EPSILON_UPPER, 6),
  ...pair(0x1f20, ETA, ETA_UPPER, 8),
  ...pair(0x1f30, IOTA, IOTA_UPPER, 8),
  ...pair(0x1f40, OMICRON, OMICRON_UPPER, 6),
  ...pair(0x1f60, OMEGA, OMEGA_UPPER, 8),

  // U+1F50–U+1F5F: upsilon. Only the rough-breathing capitals exist.
  0x1f50: compose(UPSILON, smooth, null, false),
  0x1f51: compose(UPSILON, rough, null, false),
  0x1f52: compose(UPSILON, smooth, grave, false),
  0x1f53: compose(UPSILON, rough, grave, false),
  0x1f54: compose(UPSILON, smooth, acute, false),
  0x1f55: compose(UPSILON, rough, acute, false),
  0x1f56: compose(UPSILON, smooth, perispomeni, false),
  0x1f57: compose(UPSILON, rough, perispomeni, false),
  0x1f59: compose(UPSILON_UPPER, rough, null, false),
  0x1f5b: compose(UPSILON_UPPER, rough, grave, false),
  0x1f5d: compose(UPSILON_UPPER, rough, acute, false),
  0x1f5f: compose(UPSILON_UPPER, rough, perispomeni, false),

  // U+1F70–U+1F7D: the plain accented vowels, without a breathing.
  0x1f70: compose(ALPHA, null, grave, false),
  0x1f71: compose(ALPHA, null, acute, false),
  0x1f72: compose(EPSILON, null, grave, false),
  0x1f73: compose(EPSILON, null, acute, false),
  0x1f74: compose(ETA, null, grave, false),
  0x1f75: compose(ETA, null, acute, false),
  0x1f76: compose(IOTA, null, grave, false),
  0x1f77: compose(IOTA, null, acute, false),
  0x1f78: compose(OMICRON, null, grave, false),
  0x1f79: compose(OMICRON, null, acute, false),
  0x1f7a: compose(UPSILON, null, grave, false),
  0x1f7b: compose(UPSILON, null, acute, false),
  0x1f7c: compose(OMEGA, null, grave, false),
  0x1f7d: compose(OMEGA, null, acute, false),

  // U+1F80–U+1FAF: the same again, over an iota subscript.
  ...pair(0x1f80, ALPHA, ALPHA_UPPER, 8, true),
  ...pair(0x1f90, ETA, ETA_UPPER, 8, true),
  ...pair(0x1fa0, OMEGA, OMEGA_UPPER, 8, true),

  // U+1FB0–U+1FBE: alpha's length marks, and the standalone iota.
  0x1fb0: compose(ALPHA, null, null, false),
  0x1fb1: compose(ALPHA, null, macron, false),
  0x1fb2: compose(ALPHA, null, grave, true),
  0x1fb3: compose(ALPHA, null, null, true),
  0x1fb4: compose(ALPHA, null, acute, true),
  0x1fb6: compose(ALPHA, null, perispomeni, false),
  0x1fb7: compose(ALPHA, null, perispomeni, true),
  0x1fb8: compose(ALPHA_UPPER, null, null, false),
  0x1fb9: compose(ALPHA_UPPER, null, macron, false),
  0x1fba: compose(ALPHA_UPPER, null, grave, false),
  0x1fbb: compose(ALPHA_UPPER, null, acute, false),
  0x1fbc: compose(ALPHA_UPPER, null, null, true),
  0x1fbd: shape(0.32, ...smooth(0.08, 0.28)),
  0x1fbe: shape(0.24, ...iotaBelow(0.06, 0.5)),

  // U+1FBF–U+1FCF: the standalone breathings and accents, and eta's iota forms.
  0x1fbf: shape(0.32, ...smooth(0.08, 0.28)),
  0x1fc0: shape(0.34, ...perispomeni(0.08, 0.3)),
  0x1fc1: merge(shape(0.34, ...perispomeni(0.08, 0.18)), shape(0.34, ...smooth(0.1, 0.3))),
  0x1fc2: compose(ETA, null, grave, true),
  0x1fc3: compose(ETA, null, null, true),
  0x1fc4: compose(ETA, null, acute, true),
  0x1fc6: compose(ETA, null, perispomeni, false),
  0x1fc7: compose(ETA, null, perispomeni, true),
  0x1fc8: compose(EPSILON_UPPER, null, grave, false),
  0x1fc9: compose(EPSILON_UPPER, null, acute, false),
  0x1fca: compose(ETA_UPPER, null, grave, false),
  0x1fcb: compose(ETA_UPPER, null, acute, false),
  0x1fcc: compose(ETA_UPPER, null, null, true),
  0x1fcd: merge(shape(0.34, ...smooth(0.08, 0.3)), shape(0.34, ...grave(0.1, 0.16))),
  0x1fce: merge(shape(0.34, ...smooth(0.08, 0.3)), shape(0.34, ...acute(0.1, 0.16))),
  0x1fcf: merge(shape(0.34, ...smooth(0.08, 0.3)), shape(0.34, ...perispomeni(0.08, 0.16))),

  // U+1FD0–U+1FDF: iota's length marks.
  0x1fd0: compose(IOTA, null, vrachy, false),
  0x1fd1: compose(IOTA, null, macron, false),
  0x1fd2: compose(IOTA, null, grave, false),
  0x1fd3: compose(IOTA, null, acute, false),
  0x1fd6: compose(IOTA, null, perispomeni, false),
  0x1fd7: compose(IOTA, null, perispomeni, true),
  0x1fd8: compose(IOTA_UPPER, null, vrachy, false),
  0x1fd9: compose(IOTA_UPPER, null, macron, false),
  0x1fda: compose(IOTA_UPPER, null, grave, false),
  0x1fdb: compose(IOTA_UPPER, null, acute, false),
  0x1fdd: merge(shape(0.34, ...rough(0.08, 0.3)), shape(0.34, ...grave(0.1, 0.16))),
  0x1fde: merge(shape(0.34, ...rough(0.08, 0.3)), shape(0.34, ...acute(0.1, 0.16))),
  0x1fdf: merge(shape(0.34, ...rough(0.08, 0.3)), shape(0.34, ...perispomeni(0.08, 0.16))),

  // U+1FE0–U+1FEF: upsilon's length marks, and rho's breathings.
  0x1fe0: compose(UPSILON, null, vrachy, false),
  0x1fe1: compose(UPSILON, null, macron, false),
  0x1fe2: compose(UPSILON, null, grave, false),
  0x1fe3: compose(UPSILON, null, acute, false),
  0x1fe4: compose(RHO, smooth, null, false),
  0x1fe5: compose(RHO, rough, null, false),
  0x1fe6: compose(UPSILON, null, perispomeni, false),
  0x1fe7: compose(UPSILON, null, perispomeni, true),
  0x1fe8: compose(UPSILON_UPPER, null, vrachy, false),
  0x1fe9: compose(UPSILON_UPPER, null, macron, false),
  0x1fea: compose(UPSILON_UPPER, null, grave, false),
  0x1feb: compose(UPSILON_UPPER, null, acute, false),
  0x1fec: compose(RHO_UPPER, rough, null, false),
  0x1fed: merge(shape(0.34, ...macron(0.06, 0.24)), shape(0.34, ...grave(0.1, 0.16))),
  0x1fee: merge(shape(0.34, ...macron(0.06, 0.24)), shape(0.34, ...acute(0.1, 0.16))),
  0x1fef: shape(0.3, ...grave(0.1, 0.26)),

  // U+1FF2–U+1FFE: omega's iota forms, and the last standalone marks.
  0x1ff2: compose(OMEGA, null, grave, true),
  0x1ff3: compose(OMEGA, null, null, true),
  0x1ff4: compose(OMEGA, null, acute, true),
  0x1ff6: compose(OMEGA, null, perispomeni, false),
  0x1ff7: compose(OMEGA, null, perispomeni, true),
  0x1ff8: compose(OMICRON_UPPER, null, grave, false),
  0x1ff9: compose(OMICRON_UPPER, null, acute, false),
  0x1ffa: compose(OMEGA_UPPER, null, grave, false),
  0x1ffb: compose(OMEGA_UPPER, null, acute, false),
  0x1ffc: compose(OMEGA_UPPER, null, null, true),
  0x1ffd: shape(0.3, ...acute(0.1, 0.26)),
  0x1ffe: shape(0.32, ...rough(0.08, 0.28))
};

/** Polytonic Greek, as Type3 glyphs. */
export const GREEK_POLYTONIC: Record<number, GlyphDef> = Object.fromEntries(
  Object.entries(SHAPES).map(([cp, form]) => [Number(cp), toGlyphDef(form)])
);

/** The code points this table defines, for the repertoire to state without importing it. */
export const GREEK_POLYTONIC_CODE_POINTS: readonly number[] = Object.keys(SHAPES).map(Number);
