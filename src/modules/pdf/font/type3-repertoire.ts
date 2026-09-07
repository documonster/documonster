/**
 * Which code points the Type3 path can draw without any embedded face.
 *
 * `@pdf/font/type3-glyphs` and its companion tables are 27,000 lines, loaded through a
 * dynamic `import()` so a plain-text PDF never bundles them. Two callers need to know
 * the *set* without the outlines, and neither can reach for that import:
 *
 *   - Font discovery decides whether a face that lacks a character is still the right
 *     one. It runs synchronously, and it must not pull the glyph tables into every
 *     bundle that embeds a font.
 *   - Coverage diagnostics decide whether a missing character will be drawn properly
 *     or will show up as a `.notdef` box, which are two different things to tell a
 *     caller.
 *
 * `system-fonts.test.ts` walks every code point up to `U+2FFFF` and asserts that
 * {@link isType3Drawable} agrees with `lookupGlyph`, so the table cannot drift away
 * from the glyphs it describes without a test failing.
 *
 * ## Two questions, not one
 *
 * {@link requiresEmbeddedFace} used to be defined as `!isType3Drawable`, and the two
 * were the same question for as long as Type3 drew nothing but symbols: a checkbox has
 * a built-in drawing, so a face lacking it is still fine, and an ideograph does not, so
 * a face lacking it is not.
 *
 * Adding Greek and Cyrillic letterforms broke that equivalence, because the honest
 * answer for them differs by caller. Type3 *can* draw `Δ`, which is what stops it
 * becoming a box when there is no font at all. But a monoline fallback letter is not
 * as good as the same letter from a real typeface, so discovery should still reject a
 * face that cannot draw it and keep looking. Collapsing those into one predicate would
 * have meant choosing which caller to get wrong: define letters as drawable and a
 * Chinese-only face suddenly "covers" a Bulgarian column header; define them as not
 * drawable and the diagnostics call a letter that renders perfectly well a `.notdef`
 * box.
 *
 * @module
 */

const TYPE3_SYMBOL_RANGES: ReadonlyArray<readonly [number, number]> = [
  [0x2000, 0x206f], // General Punctuation
  [0x20a0, 0x20cf], // Currency Symbols
  [0x2100, 0x23ff], // Letterlike … Miscellaneous Technical
  [0x2460, 0x2bff], // Enclosed Alphanumerics … Miscellaneous Symbols and Arrows
  [0xfeff, 0xfeff] // Zero width no-break space
];

/**
 * Letters the Type3 path draws, and would still rather a real font drew.
 *
 * The scripts that turn up *inside* an otherwise-Latin document, which is the case a
 * fallback exists for: European Latin with diacritics, monotonic Greek, Cyrillic, and
 * the raised and lowered digits a formula needs.
 *
 * What is absent is absent for one reason — this renderer does no mark positioning, so
 * anything needing a mark placed over another mark, or over a base it was not drawn
 * for, would be rendered confidently wrong. That rules out polytonic Greek, Vietnamese,
 * the combining marks at `U+0300–U+036F`, and the Cyrillic Supplement. Latin Extended-B
 * is absent for a different reason: it is novel shapes rather than compositions, so
 * there is nothing to reuse and little in a business document to justify drawing them.
 */
const TYPE3_LETTER_RANGES: ReadonlyArray<readonly [number, number]> = [
  // Latin Extended-A, whole. Every accented letter in Polish, Czech, Slovak,
  // Hungarian, Croatian, Slovenian, Latvian, Lithuanian, Estonian, Turkish, Maltese
  // and Welsh. The seven WinAnsi members (`Œœ Šš Žž Ÿ`) are drawn by a standard-14
  // face and never reach Type3, but they are defined so this range needs no holes.
  [0x0100, 0x017f],
  // Latin Extended-B, whole: the pinyin tone letters, the African hooked consonants, the
  // Serbo-Croatian digraphs, Romanian's comma-below pair and the Vietnamese horns.
  [0x0180, 0x024f],
  // Spacing diacritics: the marks as standalone characters.
  [0x02c7, 0x02c7],
  [0x02d8, 0x02dd],
  // Monotonic Greek, skipping the three code points Unicode leaves unassigned.
  [0x0384, 0x038a],
  [0x038c, 0x038c],
  [0x038e, 0x03a1],
  [0x03a3, 0x03ce],
  // Cyrillic: Russian, Bulgarian, Serbian, Macedonian, Ukrainian, Belarusian.
  [0x0400, 0x045f],
  // Superscripts and subscripts — `m²`, `H₂O`, footnote markers past `³`. Assigned
  // code points only: U+2072, U+2073 and U+208F are unassigned.
  [0x2070, 0x2071],
  [0x2074, 0x208e],
  [0x2090, 0x209c],
  // Vietnamese's tone-marked vowels. Its horned vowels are in Extended-B above.
  [0x1ea0, 0x1ef9],
  // Polytonic Greek. Stated as the runs Unicode assigns, because the block is ragged: it
  // leaves 23 cells empty where a vowel takes no circumflex or a breathing no capital.
  [0x1f00, 0x1f15],
  [0x1f18, 0x1f1d],
  [0x1f20, 0x1f45],
  [0x1f48, 0x1f4d],
  [0x1f50, 0x1f57],
  [0x1f59, 0x1f59],
  [0x1f5b, 0x1f5b],
  [0x1f5d, 0x1f5d],
  [0x1f5f, 0x1f7d],
  [0x1f80, 0x1fb4],
  [0x1fb6, 0x1fc4],
  [0x1fc6, 0x1fd3],
  [0x1fd6, 0x1fdb],
  [0x1fdd, 0x1fef],
  [0x1ff2, 0x1ff4],
  [0x1ff6, 0x1ffe]
];

function inRanges(codePoint: number, ranges: ReadonlyArray<readonly [number, number]>): boolean {
  for (const [start, end] of ranges) {
    if (codePoint >= start && codePoint <= end) {
      return true;
    }
  }
  return false;
}

/** Whether the Type3 path has a real glyph for this code point. */
export function isType3Drawable(codePoint: number): boolean {
  return inRanges(codePoint, TYPE3_SYMBOL_RANGES) || inRanges(codePoint, TYPE3_LETTER_RANGES);
}

/**
 * Whether the Type3 drawing for this code point is a letter rather than a symbol.
 *
 * The distinction a font chooser needs: a symbol's built-in drawing is as good as any
 * font's, so its absence says nothing about a face, while a letter's is a legible
 * stand-in that a real typeface still beats.
 */
export function isType3Letterform(codePoint: number): boolean {
  return inRanges(codePoint, TYPE3_LETTER_RANGES);
}

/**
 * Whether only a real font should be trusted to draw this code point.
 *
 * Named for the question font discovery asks. It is the rule that lets a Chinese face
 * which happens to lack `☐` still win, because the checkbox is drawn by Type3 and its
 * absence says nothing about whether the face suits the text — while a face that
 * cannot draw `Δ` or `П` is still passed over, even though Type3 now has letters for
 * both.
 *
 * It has to be the *exact* Type3 symbol repertoire. Using `isCjkBreakable` for it lost
 * text: that predicate answers "may a line break here", so everything outside East
 * Asian script counted as substitutable — Cyrillic and Greek included. A face covering
 * the Han and nothing else therefore won for `中文报表 Кириллица Ελληνικά`, and the
 * Cyrillic and Greek then reached a Type3 path with no glyph for either: four
 * characters silently became `.notdef` boxes. They would be drawn now, but as a
 * fallback rather than as the typeface the document asked for, so the rule stands.
 */
export function requiresEmbeddedFace(codePoint: number): boolean {
  return !isType3Drawable(codePoint) || isType3Letterform(codePoint);
}
