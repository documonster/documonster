/**
 * Type3 letterforms derived from ASCII — Latin Extended-A, superscripts, subscripts.
 *
 * The same gap as Greek and Cyrillic (issue #218) and the same cause: a browser cannot
 * read the host's fonts, so anything outside WinAnsi is drawn by the Type3 fallback or
 * not at all. Measured before this table existed, 121 of the 128 code points in Latin
 * Extended-A had no glyph, which is every accented letter in Polish, Czech, Slovak,
 * Hungarian, Romanian, Croatian, Slovenian, Latvian, Lithuanian, Estonian, Turkish,
 * Maltese and Welsh:
 *
 *     Polish     Łódź zażółć gęślą jaźń   →  Łźżłćęśąźń  as boxes
 *     Czech      Příliš žluťoučký kůň     →  řťčůňě
 *     Hungarian  Árvíztűrő tükörfúrógép   →  űő
 *     Turkish    İstanbul şğıöçü          →  İşğı
 *
 * And 48 more in Superscripts and Subscripts, which is `H₂O`, `CO₂`, `m²`, `x⁴` and
 * every footnote marker past `³`.
 *
 * ## Nothing here is drawn from scratch
 *
 * Every glyph is a transform of a letter the shared stroke font already has:
 *
 *   - **A base plus a mark.** 106 of the 128 are exactly that — `ć` is `c` and an
 *     acute, `ę` is `e` and an ogonek, `ő` is `o` and a double acute. The marks are
 *     the ones the Greek and Cyrillic tables use, plus the seven this block needs.
 *   - **A bar through a letter.** `ł đ ħ ŧ Ł Đ Ħ Ŧ`.
 *   - **A scaled letter.** Every superscript and subscript, through the same
 *     `scaleTo` that derives Cyrillic small capitals.
 *   - **A rotated letter.** `ə`, which is `e` through half a turn.
 *
 * Four letters have no ASCII twin and are authored: `ı` (a dotless `i`, which is the
 * stem of `i` without its dot), `Ŋ`/`ŋ` (eng) and `ſ` (long s).
 *
 * ## What is deliberately absent
 *
 * **Latin Extended-B** (207 code points), because it is not compositions: `Ɔ Ə Ɣ Ƣ ƻ`
 * are novel shapes for African and phonetic orthographies, and drawing them is not
 * reuse of anything. The four Romanian letters that live there — `Ș ș Ț ț`, comma-below
 * rather than cedilla — are included, because leaving them out would mean Romanian
 * renders in the wrong diacritic while every other language in the region is right.
 *
 * **Vietnamese** (Latin Extended Additional), because `ế` is a base under *two* stacked
 * marks and this renderer cannot position one mark over another. A confident wrong
 * shape is worse than a box the caller can see and fix with a font.
 *
 * **Combining marks** (U+0300–U+036F), for the same reason: they are defined to sit on
 * the preceding glyph, and without mark positioning they would be drawn beside it.
 *
 * @module
 */

import type { GlyphDef } from "@pdf/font/type3-glyphs";
import {
  acute,
  ascii,
  breve,
  ligature,
  shiftPaths,
  asSubscript,
  asSuperscript,
  caron,
  cedilla,
  circumflex,
  commaBelow,
  diaeresis,
  dotAbove,
  doubleAcute,
  macron,
  merge,
  ogonek,
  ring,
  rotate180,
  shape,
  tilde,
  toGlyphDef,
  withApostrophe,
  withBar,
  withMarkOver,
  withMarkOverSmall,
  withMarkUnder,
  type Path,
  type StrokeShape
} from "@pdf/font/type3-letterforms";

// =============================================================================
// The four letters with no ASCII twin
// =============================================================================

/**
 * ı — a dotless `i`, Turkish.
 *
 * Taken from `i` by dropping the dot rather than by drawing a stem, so the stem's
 * position and length stay tied to the font it came from.
 */
const DOTLESS_I: StrokeShape = (() => {
  const source = ascii("i");
  // The dot is the short stroke; the stem is the long one.
  const height = (path: readonly (readonly [number, number])[]): number => {
    const ys = path.map(([, y]) => y);
    return Math.max(...ys) - Math.min(...ys);
  };
  const stem = [...source.d].sort((a, b) => height(b) - height(a))[0];
  return shape(source.w, stem);
})();

/** Ŋ — eng: `N` with a descending hook on its right stem. */
const ENG_UPPER = merge(
  ascii("N"),
  shape(0.52, [
    [0.42, 0.75],
    [0.42, 0.88],
    [0.34, 0.93]
  ])
);

/** ŋ — eng: `n` with a descending hook. */
const ENG_LOWER = merge(
  ascii("n"),
  shape(0.45, [
    [0.38, 0.75],
    [0.38, 0.88],
    [0.3, 0.93]
  ])
);

/** ſ — long s: an `f` without the right half of its crossbar. */
const LONG_S = shape(
  0.3,
  [
    [0.24, 0.15],
    [0.16, 0.16],
    [0.12, 0.24],
    [0.12, 0.75]
  ],
  [
    [0.04, 0.38],
    [0.16, 0.38]
  ]
);

/** ĸ — kra, Greenlandic: a `k` with no ascender. */
const KRA = shape(
  0.43,
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

// =============================================================================
// Composition shorthands
// =============================================================================

/** A capital with a mark above it. */
const cap = (ch: string, mark: (x: number, y: number) => Path[]): StrokeShape =>
  withMarkOver(ascii(ch), mark);

/**
 * A lowercase letter with a mark above it.
 *
 * An x-height letter takes this; one with an ascender — `ď` `ľ` `ť` `ĥ` `ĵ` — takes
 * {@link cap} instead, because its stem already reaches the cap height and the
 * x-height placement would put the mark through it.
 */
const low = (ch: string, mark: (x: number, y: number) => Path[]): StrokeShape =>
  withMarkOverSmall(ascii(ch), mark);

/** A letter with a mark hung beneath it. */
const under = (ch: string, mark: (x: number, y: number) => Path[], stemX = 0.5): StrokeShape =>
  withMarkUnder(ascii(ch), mark, stemX);

// =============================================================================
// Latin Extended-A, U+0100–U+017F
// =============================================================================

const LATIN_A: Record<number, StrokeShape> = {
  0x0100: cap("A", macron),
  0x0101: low("a", macron),
  0x0102: cap("A", breveMark),
  0x0103: low("a", breveMark),
  0x0104: under("A", ogonek, 0.72),
  0x0105: under("a", ogonek, 0.78),
  0x0106: cap("C", acute),
  0x0107: low("c", acute),
  0x0108: cap("C", circumflex),
  0x0109: low("c", circumflex),
  0x010a: cap("C", dotAbove),
  0x010b: low("c", dotAbove),
  0x010c: cap("C", caron),
  0x010d: low("c", caron),
  0x010e: cap("D", caron),
  0x010f: withApostrophe(ascii("d")),
  0x0110: withBar(ascii("D"), 0.02, 0.45, 0.22, 0.45),
  0x0111: withBar(ascii("d"), 0.24, 0.28, 0.46, 0.28),
  0x0112: cap("E", macron),
  0x0113: low("e", macron),
  0x0114: cap("E", breveMark),
  0x0115: low("e", breveMark),
  0x0116: cap("E", dotAbove),
  0x0117: low("e", dotAbove),
  0x0118: under("E", ogonek, 0.5),
  0x0119: under("e", ogonek, 0.5),
  0x011a: cap("E", caron),
  0x011b: low("e", caron),
  0x011c: cap("G", circumflex),
  0x011d: low("g", circumflex),
  0x011e: cap("G", breveMark),
  0x011f: low("g", breveMark),
  0x0120: cap("G", dotAbove),
  0x0121: low("g", dotAbove),
  0x0122: under("G", cedilla),
  0x0123: low("g", cedilla),
  0x0124: cap("H", circumflex),
  0x0125: cap("h", circumflex),
  0x0126: withBar(ascii("H"), 0.04, 0.24, 0.48, 0.24),
  0x0127: withBar(ascii("h"), 0.02, 0.26, 0.24, 0.26),
  0x0128: cap("I", tilde),
  0x0129: cap("i", tilde),
  0x012a: cap("I", macron),
  0x012b: cap("i", macron),
  0x012c: cap("I", breveMark),
  0x012d: cap("i", breveMark),
  0x012e: under("I", ogonek, 0.5),
  0x012f: under("i", ogonek, 0.5),
  0x0130: cap("I", dotAbove),
  0x0131: DOTLESS_I,
  0x0132: ligature("I", "J", 0.06),
  0x0133: ligature("i", "j", 0.06),
  0x0134: cap("J", circumflex),
  0x0135: cap("j", circumflex),
  0x0136: under("K", cedilla),
  0x0137: under("k", cedilla),
  0x0138: KRA,
  0x0139: cap("L", acute),
  0x013a: cap("l", acute),
  0x013b: under("L", cedilla, 0.25),
  0x013c: under("l", cedilla, 0.4),
  0x013d: cap("L", caron),
  0x013e: withApostrophe(ascii("l")),
  0x013f: withBar(ascii("L"), 0.24, 0.5, 0.26, 0.5),
  0x0140: withBar(ascii("l"), 0.22, 0.5, 0.24, 0.5),
  0x0141: withBar(ascii("L"), 0.04, 0.55, 0.22, 0.42),
  0x0142: withBar(ascii("l"), 0.04, 0.5, 0.22, 0.37),
  0x0143: cap("N", acute),
  0x0144: low("n", acute),
  0x0145: under("N", cedilla),
  0x0146: under("n", cedilla),
  0x0147: cap("N", caron),
  0x0148: low("n", caron),
  0x0149: merge(
    shape(0.6, [
      [0.04, 0.15],
      [0.08, 0.26]
    ]),
    shape(0.6, ...shiftPaths(ascii("n"), 0.12))
  ),
  0x014a: ENG_UPPER,
  0x014b: ENG_LOWER,
  0x014c: cap("O", macron),
  0x014d: low("o", macron),
  0x014e: cap("O", breveMark),
  0x014f: low("o", breveMark),
  0x0150: cap("O", doubleAcute),
  0x0151: low("o", doubleAcute),
  0x0152: ligature("O", "E", 0.16),
  0x0153: ligature("o", "e", 0.11),
  0x0154: cap("R", acute),
  0x0155: low("r", acute),
  0x0156: under("R", cedilla),
  0x0157: under("r", cedilla, 0.35),
  0x0158: cap("R", caron),
  0x0159: low("r", caron),
  0x015a: cap("S", acute),
  0x015b: low("s", acute),
  0x015c: cap("S", circumflex),
  0x015d: low("s", circumflex),
  0x015e: under("S", cedilla),
  0x015f: under("s", cedilla),
  0x0160: cap("S", caron),
  0x0161: low("s", caron),
  0x0162: under("T", cedilla),
  0x0163: under("t", cedilla, 0.6),
  0x0164: cap("T", caron),
  0x0165: withApostrophe(ascii("t")),
  0x0166: withBar(ascii("T"), 0.1, 0.45, 0.38, 0.45),
  0x0167: withBar(ascii("t"), 0.06, 0.58, 0.28, 0.58),
  0x0168: cap("U", tilde),
  0x0169: low("u", tilde),
  0x016a: cap("U", macron),
  0x016b: low("u", macron),
  0x016c: cap("U", breveMark),
  0x016d: low("u", breveMark),
  0x016e: cap("U", ring),
  0x016f: low("u", ring),
  0x0170: cap("U", doubleAcute),
  0x0171: low("u", doubleAcute),
  0x0172: under("U", ogonek, 0.5),
  0x0173: under("u", ogonek, 0.5),
  0x0174: cap("W", circumflex),
  0x0175: low("w", circumflex),
  0x0176: cap("Y", circumflex),
  0x0177: low("y", circumflex),
  0x0178: cap("Y", diaeresis),
  0x0179: cap("Z", acute),
  0x017a: low("z", acute),
  0x017b: cap("Z", dotAbove),
  0x017c: low("z", dotAbove),
  0x017d: cap("Z", caron),
  0x017e: low("z", caron),
  0x017f: LONG_S
};

// =============================================================================
// The Romanian pair from Latin Extended-B, U+0218–U+021B
// =============================================================================

const ROMANIAN: Record<number, StrokeShape> = {
  0x0218: under("S", commaBelow),
  0x0219: under("s", commaBelow),
  0x021a: under("T", commaBelow),
  0x021b: under("t", commaBelow, 0.6)
};

// =============================================================================
// Spacing diacritics, U+02C7 and U+02D8–U+02DD
// =============================================================================

const SPACING_MARKS: Record<number, StrokeShape> = {
  0x02c7: shape(0.35, ...caron(0.08, 0.3)),
  0x02d8: shape(0.35, ...breveMark(0.08, 0.3)),
  0x02d9: shape(0.35, ...dotAbove(0.1, 0.26)),
  0x02da: shape(0.35, ...ring(0.08, 0.32)),
  0x02db: shape(0.35, ...ogonek(0.1, 0.62)),
  0x02dc: shape(0.35, ...tilde(0.08, 0.3)),
  0x02dd: shape(0.4, ...doubleAcute(0.08, 0.3))
};

// =============================================================================
// Superscripts and subscripts, U+2070–U+209C
// =============================================================================

/** Digits `0`–`9`, raised or lowered. */
function digits(transform: (s: StrokeShape) => StrokeShape): StrokeShape[] {
  return [...Array(10).keys()].map(d => transform(ascii(String(d))));
}

const SUP_DIGITS = digits(asSuperscript);
const SUB_DIGITS = digits(asSubscript);

/** An x-height letter, raised or lowered — its band starts at the x-height, not the cap. */
const supLow = (ch: string): StrokeShape => asSuperscript(ascii(ch), 0.38);
const subLow = (ch: string): StrokeShape => asSubscript(ascii(ch), 0.38);

const SCRIPTS: Record<number, StrokeShape> = {
  0x2070: SUP_DIGITS[0],
  0x2071: supLow("i"),
  0x2074: SUP_DIGITS[4],
  0x2075: SUP_DIGITS[5],
  0x2076: SUP_DIGITS[6],
  0x2077: SUP_DIGITS[7],
  0x2078: SUP_DIGITS[8],
  0x2079: SUP_DIGITS[9],
  0x207a: asSuperscript(ascii("+")),
  0x207b: asSuperscript(ascii("-")),
  0x207c: asSuperscript(ascii("=")),
  0x207d: asSuperscript(ascii("(")),
  0x207e: asSuperscript(ascii(")")),
  0x207f: supLow("n"),
  0x2080: SUB_DIGITS[0],
  0x2081: SUB_DIGITS[1],
  0x2082: SUB_DIGITS[2],
  0x2083: SUB_DIGITS[3],
  0x2084: SUB_DIGITS[4],
  0x2085: SUB_DIGITS[5],
  0x2086: SUB_DIGITS[6],
  0x2087: SUB_DIGITS[7],
  0x2088: SUB_DIGITS[8],
  0x2089: SUB_DIGITS[9],
  0x208a: asSubscript(ascii("+")),
  0x208b: asSubscript(ascii("-")),
  0x208c: asSubscript(ascii("=")),
  0x208d: asSubscript(ascii("(")),
  0x208e: asSubscript(ascii(")")),
  0x2090: subLow("a"),
  0x2091: subLow("e"),
  0x2092: subLow("o"),
  0x2093: subLow("x"),
  0x2094: asSubscript(rotate180(ascii("e")), 0.38),
  0x2095: subLow("h"),
  0x2096: subLow("k"),
  0x2097: subLow("l"),
  0x2098: subLow("m"),
  0x2099: subLow("n"),
  0x209a: subLow("p"),
  0x209b: subLow("s"),
  0x209c: subLow("t")
};

// =============================================================================
// Helpers used above, defined after the marks they wrap
// =============================================================================

/** The breve at the width this block's letters take. */
function breveMark(x: number, y: number): Path[] {
  return breve(x, y, 0.16);
}

// =============================================================================
// Table
// =============================================================================

const SHAPES: Record<number, StrokeShape> = {
  ...SPACING_MARKS,
  ...LATIN_A,
  ...ROMANIAN,
  ...SCRIPTS
};

/** Latin Extended-A, the Romanian comma-below pair, spacing marks and scripts. */
export const LATIN_EXTENDED: Record<number, GlyphDef> = Object.fromEntries(
  Object.entries(SHAPES).map(([cp, form]) => [Number(cp), toGlyphDef(form)])
);

/** The code points this table defines, for the repertoire to state without importing it. */
export const LATIN_EXTENDED_CODE_POINTS: readonly number[] = Object.keys(SHAPES).map(Number);
