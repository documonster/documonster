/**
 * Type3 letterforms for Vietnamese — U+1EA0…U+1EF9, plus the four horned letters.
 *
 * This block was excluded from the first pass on a reason that turned out to be wrong.
 * The note said `ế` is "a base under two stacked marks and this renderer cannot position
 * one mark over another" — which conflates two different things. The renderer does no
 * *runtime* mark positioning, so a decomposed `e` + U+0302 + U+0301 genuinely cannot be
 * drawn correctly. But `ế` is a **precomposed code point**: one glyph, whose outline this
 * directory authors in full. Nothing prevents that outline carrying two marks, and
 * {@link withMarkStack} places the upper one from the lower one's own height.
 *
 * Vietnamese is worth the correction. It is the sixth-largest language by native
 * speakers, its orthography is Latin, and every one of these 94 code points was a box.
 *
 * ## The shape of the block
 *
 * Vietnamese writes twelve vowels — `a ă â e ê i o ô ơ u ư y` — under five tones, and the
 * two that carry a horn (`ơ ư`) live in Latin Extended-B rather than here. So the table is
 * a product, not a list:
 *
 *   - **One mark**, where the vowel is plain: `ạ ả ã ẹ ẻ ẽ ị ọ ỏ ụ ủ ỳ ỵ ỷ ỹ`.
 *   - **Two marks stacked**, where the vowel already carries a breve, circumflex or horn:
 *     `ắ ằ ẳ ẵ ấ ầ ẩ ẫ ế ề ể ễ ố ồ ổ ỗ ớ ờ ở ỡ ứ ừ ử ữ`.
 *   - **A mark above and a dot below**, which do not interact: `ặ ậ ệ ộ ợ ự`.
 *
 * The five tones are `grave`, `acute`, {@link hookAbove}, `tilde` and {@link dotBelow},
 * and the first two were already needed by Greek and by Latin Extended-A.
 *
 * @module
 */

import type { GlyphDef } from "@pdf/font/type3-glyphs";
import {
  acute,
  ascii,
  breveWide,
  circumflex,
  dotBelow,
  grave,
  hookAbove,
  tilde,
  toGlyphDef,
  withHorn,
  withMarkOver,
  withMarkOverSmall,
  withMarkStack,
  withMarkUnder,
  type Path,
  type StrokeShape
} from "@pdf/font/type3-letterforms";

// =============================================================================
// The vowels the tones are written on
// =============================================================================

/** `Ơ ơ Ư ư` — U+01A0, U+01A1, U+01AF, U+01B0, in Latin Extended-B. */
const O_HORN_UPPER = withHorn(ascii("O"));
const O_HORN = withHorn(ascii("o"));
const U_HORN_UPPER = withHorn(ascii("U"));
const U_HORN = withHorn(ascii("u"));

/** The vowels that already carry a mark, and therefore stack. */
const BREVE_A_UPPER = withMarkOver(ascii("A"), breveWide);
const BREVE_A = withMarkOverSmall(ascii("a"), breveWide);
const CIRC_A_UPPER = withMarkOver(ascii("A"), circumflex);
const CIRC_A = withMarkOverSmall(ascii("a"), circumflex);
const CIRC_E_UPPER = withMarkOver(ascii("E"), circumflex);
const CIRC_E = withMarkOverSmall(ascii("e"), circumflex);
const CIRC_O_UPPER = withMarkOver(ascii("O"), circumflex);
const CIRC_O = withMarkOverSmall(ascii("o"), circumflex);

// =============================================================================
// Composition shorthands
// =============================================================================

type Mark = (x: number, y: number) => Path[];

/** A plain vowel under one tone. The letter's case decides the mark's band. */
const one = (ch: string, tone: Mark): StrokeShape =>
  ch === ch.toUpperCase() ? withMarkOver(ascii(ch), tone) : withMarkOverSmall(ascii(ch), tone);

/** A plain vowel over a dot. */
const under = (ch: string, stemX = 0.5): StrokeShape => withMarkUnder(ascii(ch), dotBelow, stemX);

/** A marked vowel under a second, stacked tone. */
const stack = (base: StrokeShape, lower: Mark, tone: Mark, small: boolean): StrokeShape =>
  withMarkStack(base, lower, tone, small);

/** A marked vowel over a dot — the two marks never meet, so they simply merge. */
const dotted = (base: StrokeShape, stemX = 0.5): StrokeShape =>
  withMarkUnder(base, dotBelow, stemX);

// =============================================================================
// Table
// =============================================================================

const SHAPES: Record<number, StrokeShape> = {
  // Latin Extended-B: the horned vowels.
  0x01a0: O_HORN_UPPER,
  0x01a1: O_HORN,
  0x01af: U_HORN_UPPER,
  0x01b0: U_HORN,

  // A
  0x1ea0: under("A", 0.5),
  0x1ea1: under("a", 0.5),
  0x1ea2: one("A", hookAbove),
  0x1ea3: one("a", hookAbove),
  0x1ea4: stack(CIRC_A_UPPER, circumflex, acute, false),
  0x1ea5: stack(CIRC_A, circumflex, acute, true),
  0x1ea6: stack(CIRC_A_UPPER, circumflex, grave, false),
  0x1ea7: stack(CIRC_A, circumflex, grave, true),
  0x1ea8: stack(CIRC_A_UPPER, circumflex, hookAbove, false),
  0x1ea9: stack(CIRC_A, circumflex, hookAbove, true),
  0x1eaa: stack(CIRC_A_UPPER, circumflex, tilde, false),
  0x1eab: stack(CIRC_A, circumflex, tilde, true),
  0x1eac: dotted(CIRC_A_UPPER),
  0x1ead: dotted(CIRC_A),
  0x1eae: stack(BREVE_A_UPPER, breveWide, acute, false),
  0x1eaf: stack(BREVE_A, breveWide, acute, true),
  0x1eb0: stack(BREVE_A_UPPER, breveWide, grave, false),
  0x1eb1: stack(BREVE_A, breveWide, grave, true),
  0x1eb2: stack(BREVE_A_UPPER, breveWide, hookAbove, false),
  0x1eb3: stack(BREVE_A, breveWide, hookAbove, true),
  0x1eb4: stack(BREVE_A_UPPER, breveWide, tilde, false),
  0x1eb5: stack(BREVE_A, breveWide, tilde, true),
  0x1eb6: dotted(BREVE_A_UPPER),
  0x1eb7: dotted(BREVE_A),

  // E
  0x1eb8: under("E"),
  0x1eb9: under("e"),
  0x1eba: one("E", hookAbove),
  0x1ebb: one("e", hookAbove),
  0x1ebc: one("E", tilde),
  0x1ebd: one("e", tilde),
  0x1ebe: stack(CIRC_E_UPPER, circumflex, acute, false),
  0x1ebf: stack(CIRC_E, circumflex, acute, true),
  0x1ec0: stack(CIRC_E_UPPER, circumflex, grave, false),
  0x1ec1: stack(CIRC_E, circumflex, grave, true),
  0x1ec2: stack(CIRC_E_UPPER, circumflex, hookAbove, false),
  0x1ec3: stack(CIRC_E, circumflex, hookAbove, true),
  0x1ec4: stack(CIRC_E_UPPER, circumflex, tilde, false),
  0x1ec5: stack(CIRC_E, circumflex, tilde, true),
  0x1ec6: dotted(CIRC_E_UPPER),
  0x1ec7: dotted(CIRC_E),

  // I
  0x1ec8: one("I", hookAbove),
  0x1ec9: one("i", hookAbove),
  0x1eca: under("I"),
  0x1ecb: under("i"),

  // O
  0x1ecc: under("O"),
  0x1ecd: under("o"),
  0x1ece: one("O", hookAbove),
  0x1ecf: one("o", hookAbove),
  0x1ed0: stack(CIRC_O_UPPER, circumflex, acute, false),
  0x1ed1: stack(CIRC_O, circumflex, acute, true),
  0x1ed2: stack(CIRC_O_UPPER, circumflex, grave, false),
  0x1ed3: stack(CIRC_O, circumflex, grave, true),
  0x1ed4: stack(CIRC_O_UPPER, circumflex, hookAbove, false),
  0x1ed5: stack(CIRC_O, circumflex, hookAbove, true),
  0x1ed6: stack(CIRC_O_UPPER, circumflex, tilde, false),
  0x1ed7: stack(CIRC_O, circumflex, tilde, true),
  0x1ed8: dotted(CIRC_O_UPPER),
  0x1ed9: dotted(CIRC_O),
  0x1eda: withMarkOver(O_HORN_UPPER, acute),
  0x1edb: withMarkOverSmall(O_HORN, acute),
  0x1edc: withMarkOver(O_HORN_UPPER, grave),
  0x1edd: withMarkOverSmall(O_HORN, grave),
  0x1ede: withMarkOver(O_HORN_UPPER, hookAbove),
  0x1edf: withMarkOverSmall(O_HORN, hookAbove),
  0x1ee0: withMarkOver(O_HORN_UPPER, tilde),
  0x1ee1: withMarkOverSmall(O_HORN, tilde),
  0x1ee2: dotted(O_HORN_UPPER),
  0x1ee3: dotted(O_HORN),

  // U
  0x1ee4: under("U"),
  0x1ee5: under("u"),
  0x1ee6: one("U", hookAbove),
  0x1ee7: one("u", hookAbove),
  0x1ee8: withMarkOver(U_HORN_UPPER, acute),
  0x1ee9: withMarkOverSmall(U_HORN, acute),
  0x1eea: withMarkOver(U_HORN_UPPER, grave),
  0x1eeb: withMarkOverSmall(U_HORN, grave),
  0x1eec: withMarkOver(U_HORN_UPPER, hookAbove),
  0x1eed: withMarkOverSmall(U_HORN, hookAbove),
  0x1eee: withMarkOver(U_HORN_UPPER, tilde),
  0x1eef: withMarkOverSmall(U_HORN, tilde),
  0x1ef0: dotted(U_HORN_UPPER),
  0x1ef1: dotted(U_HORN),

  // Y
  0x1ef2: one("Y", grave),
  0x1ef3: one("y", grave),
  0x1ef4: under("Y"),
  0x1ef5: under("y"),
  0x1ef6: one("Y", hookAbove),
  0x1ef7: one("y", hookAbove),
  0x1ef8: one("Y", tilde),
  0x1ef9: one("y", tilde)
};

/** Vietnamese, as Type3 glyphs. */
export const VIETNAMESE: Record<number, GlyphDef> = Object.fromEntries(
  Object.entries(SHAPES).map(([cp, form]) => [Number(cp), toGlyphDef(form)])
);

/** The code points this table defines, for the repertoire to state without importing it. */
export const VIETNAMESE_CODE_POINTS: readonly number[] = Object.keys(SHAPES).map(Number);
