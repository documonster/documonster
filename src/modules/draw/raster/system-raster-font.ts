/**
 * Acquire system fonts for the rasteriser — Node.js.
 *
 * Separate from the rasteriser because it is the only part of it that touches a
 * filesystem: `glyph-outline.ts` and `glyph-rasterizer.ts` are wanted in every
 * environment, and while the paths sat inside them every browser bundle carried
 * `/System/Library/Fonts/Supplemental/Arial.ttf` and the rest. A
 * `typeof process` guard cannot remove them — a bundler must keep any string the
 * module might reach — so the split is what makes them droppable, and
 * `system-raster-font.browser.ts` is what replaces it.
 *
 * ## What changed, and why the old version drew blanks
 *
 * This used to hold a hardcoded list of five filenames per platform — Arial,
 * Helvetica, SFNS on macOS; DejaVu, Liberation, FreeSans on Linux — take the first
 * that exists, parse it, cache it, done. **Not one of them contains a CJK glyph.**
 * On macOS it resolved to `Arial.ttf`, whose `cmap` has no entry for U+4E2D, and
 * `canvas.ts` responded to a missing outline by advancing the pen and drawing
 * nothing. So a PNG of a Chinese label came out blank while the PDF of the same
 * drawing embedded `Songti SC` correctly — because the PDF pipeline had real font
 * discovery and this had a list.
 *
 * Discovery now comes from `@utils/font-discovery`, the same search the PDF
 * embedder uses, which walks the platform's font directories (including the macOS
 * `AssetsV2` tree where recent releases keep 黑体, 楷体 and 仿宋) and picks a face
 * by coverage and language rather than by filename.
 *
 * ## Why this returns a list
 *
 * PDF can fall back to a Type3 drawing for a code point its chosen face lacks; a
 * rasteriser has nothing to fall back to but another font. So the answer here is a
 * **chain**: the face that covers the text, then a general-purpose Latin face, and
 * `canvas.ts` asks each in turn per character. One face is not enough for
 * `Mixed 混合 ABC` unless it happens to cover both scripts.
 *
 * @module
 */

import type { RasterFont } from "@draw/raster/glyph-outline";
import {
  EMPTY_RASTER_FONT,
  buildRasterFont,
  fontHasGlyph,
  parseRasterFont
} from "@draw/raster/glyph-outline";
import { DEFAULT_TEXT_FAMILY } from "@draw/types";
import type { RasterTextStyle } from "@draw/types";
import type { CjkLanguage } from "@utils/cjk";
import { detectCjkLanguage } from "@utils/cjk";
import { findSystemFontForCodePoints } from "@utils/font-discovery";
import { isNonPrintingControl } from "@utils/font-metrics";
import type { TtfFont } from "@utils/font-ttf";
import { fileExistsSync, readFileBytesSync } from "@utils/fs";

/**
 * The face found for a script, so every label in that script shares one search.
 *
 * Keyed by language rather than by the code points asked for, and that distinction
 * is the whole point. Keyed by code points, a diagram's 18 Chinese labels are 18
 * different keys — they do not use the same characters — so each one ran a full
 * directory walk and parsed `Songti.ttc` again. Measured: 18 labels produced **18
 * separate `Songti SC` objects**, 18 copies of a 43,000-entry `cmap` held at once,
 * and a glyph cache that never hit, because it is keyed on outline identity and
 * every copy has its own.
 *
 * Language is also what makes the *right* face come back. Reusing any already-found
 * face instead would be faster still and wrong: `Songti SC` carries kana, so a
 * Japanese label would be drawn in a Chinese hand rather than prompting a search
 * for a Japanese one. `null` records that nothing on this host can serve the script.
 */
const _byLanguage = new Map<string, RasterFont | null>();

/**
 * Faces found for code points the script's own face could not draw.
 *
 * A Chinese face is chosen for Chinese, but a label may still carry a character it
 * lacks — a rare ideograph, a symbol. Those searches are per code point set and
 * cannot be shared, so they are kept apart from the per-language answer above and
 * bounded by {@link _unresolvable}.
 */
let _extraFaces: RasterFont[] = [];

/**
 * Code points already proven undrawable on this host.
 *
 * Without this, a character nothing can draw costs a full font-directory walk on
 * **every** label that contains it — the search cannot succeed, so nothing caches
 * and nothing stops it repeating.
 */
const _unresolvable = new Set<number>();

/**
 * Faces already built, by which face they are, so one face is parsed once.
 *
 * The per-language cache above is not enough on its own, because two keys legitimately
 * resolve to the same font. `detectCjkLanguage` answers `undefined` for text whose
 * characters are shaped identically in Simplified and Traditional — "开始处理" has no
 * distinguishing evidence, "数据校验" does — so one diagram's labels alternate between
 * the `""` and `"zh-Hans"` keys and each searched separately. Both found `Songti SC`
 * Regular, and built it twice: two 43,000-entry `cmap`s, and a glyph cache that missed
 * on every label because it is keyed on outline identity.
 *
 * Deduplicating on what came back rather than on what was asked for fixes that class of
 * problem once, instead of guessing which keys ought to have been the same. The search
 * still runs per key — cheap, since `@utils/font-discovery` caches the path index and
 * per-face metadata — but the bytes are turned into a usable face only once.
 */
const _facesByIdentity = new Map<string, RasterFont>();

/** The ordinary sans-serif Latin face, found once. */
let _latin: RasterFont | null | undefined;

/**
 * Faces found for a non-default style request, keyed by family, weight and slant.
 *
 * Separate from {@link _latin} because the fast filename path cannot answer "Courier
 * New, bold" — that needs a real search, and repeating it per label would cost a font
 * directory walk each time.
 */
const _styled = new Map<string, RasterFont | null>();

/**
 * Forget every face discovered on this host, and every conclusion drawn from them.
 *
 * The caches hold a font's bytes and its `cmap`, so this is also how a long-lived
 * process releases them. Tests call it to start from a known state.
 */
export function resetDiscoveredFonts(): void {
  _byLanguage.clear();
  _extraFaces = [];
  _unresolvable.clear();
  _facesByIdentity.clear();
  _styled.clear();
  _latin = undefined;
}

/**
 * The ordered faces to try when drawing `text`.
 *
 * `injected` are the caller's own, which outrank anything discovered; then a
 * general-purpose Latin face; then — only if those cannot draw the whole string — a
 * face discovered for what is left.
 *
 * **The Latin face comes before the discovered one deliberately.** A CJK face carries
 * Latin glyphs too, so putting it first meant `Mixed 混合 ABC` drew its ASCII in
 * Songti's serif hand while the SVG of the same drawing used Arial — the two backends
 * disagreed about a string neither had any trouble with. None of the Latin candidates
 * below contains a Han, Kana or Hangul glyph, so this cannot shadow the discovered face
 * for the text that needed it; it only claims the ASCII both faces could draw.
 *
 * **Discovery is skipped when the chain already covers the text**, so a caller who
 * supplied the right font does not pay for a search that can only confirm it.
 *
 * There is no process-wide font registry to consult. Fonts belong to the render or the
 * canvas that supplied them — a module-level setter let one caller change what another
 * caller's canvas drew with, and existed only because the tests used it.
 */
export function resolveFontChain(
  text: string,
  injected: readonly RasterFont[],
  useSystemFonts: boolean,
  style: RasterTextStyle = {}
): readonly RasterFont[] {
  const chain: RasterFont[] = [...injected];
  if (!useSystemFonts) {
    return chain;
  }

  const latin = styleFace(style);
  if (latin) {
    chain.push(latin);
  }

  // One search per script. `detectCjkLanguage` returns undefined for text with no
  // East Asian characters, which is its own key: Cyrillic and Greek need a face too
  // and have no regional hand to protect.
  const language = detectCjkLanguage(text);
  const key = `${language ?? ""}|${style.family ?? ""}|${style.bold ? "b" : ""}${style.italic ? "i" : ""}`;
  if (!_byLanguage.has(key)) {
    const needed = uncoveredBy(chain, text);
    _byLanguage.set(key, needed.size === 0 ? null : discover(needed, language, style));
  }
  const forScript = _byLanguage.get(key);
  if (forScript && !chain.includes(forScript)) {
    chain.push(forScript);
  }
  for (const face of _extraFaces) {
    if (!chain.includes(face)) {
      chain.push(face);
    }
  }

  // Keep asking until the text is covered or a search stops making progress.
  //
  // One search is not enough, because discovery is allowed to return a face that
  // covers only *part* of what was asked for — deliberately, since a face drawing
  // most of the text beats no face at all. Taking the first answer and stopping
  // meant `αא` loaded the Greek face and dropped the Hebrew, even with a Hebrew
  // font installed: the character was lost to a font that existed and was never
  // asked for. Worse, nothing recorded the failure, so every later label repeated
  // the whole search and appended the same face again — an unbounded chain.
  let missing = uncoveredBy(chain, text);
  for (let attempt = 0; missing.size > 0 && attempt < MAX_FALLBACK_FACES; attempt++) {
    const extra = discover(missing, language, style);
    if (extra === null || chain.includes(extra)) {
      break; // nothing installed can help, or it offered a face already in the chain
    }
    const before = missing.size;
    chain.push(extra);
    if (!_extraFaces.includes(extra)) {
      _extraFaces.push(extra);
    }
    missing = uncoveredBy(chain, text);
    if (missing.size >= before) {
      break; // the face covered none of what was missing; stop rather than spin
    }
  }

  // Whatever is still missing cannot be drawn on this host. Remember it, or every
  // later label containing one of these repeats the entire search — a search that
  // cannot succeed, so nothing else would ever cache it.
  for (const cp of missing) {
    _unresolvable.add(cp);
  }
  return chain;
}

/**
 * How many supplemental faces one string may pull in.
 *
 * A label mixing several scripts legitimately needs a few; needing dozens means
 * discovery is returning faces that do not help, and the cost is paid on every
 * character of every later label because the chain is walked per code point.
 */
const MAX_FALLBACK_FACES = 8;

/**
 * Ask the host for a face covering `wanted`, in `language`'s hand and `style`'s design.
 *
 * The rasteriser has no synthetic fallback, so every code point is a requirement —
 * which is `findSystemFontForCodePoints`'s default predicate.
 */
function discover(
  wanted: ReadonlySet<number>,
  language: CjkLanguage | undefined,
  style: RasterTextStyle,
  families: readonly string[] = style.family === undefined ? [] : [style.family]
): RasterFont | null {
  const ttf = findSystemFontForCodePoints(wanted, {
    ...(language === undefined ? {} : { language }),
    ...(families.length === 0 ? {} : { families }),
    ...(style.bold ? { weight: 700 } : {}),
    ...(style.italic ? { italic: true } : {})
  });
  if (!ttf) {
    return null;
  }
  // Identity, not the request: see `_facesByIdentity`. The PostScript name is what
  // actually distinguishes two faces of one collection; family, weight, `unitsPerEm`
  // and glyph count are all legitimately shared between them.
  const identity = [
    ttf.postScriptName,
    ttf.familyName,
    ttf.weightClass,
    ttf.italic ? "i" : "u",
    ttf.unitsPerEm,
    ttf.numGlyphs
  ].join("\u0000");
  const existing = _facesByIdentity.get(identity);
  // A key is a heuristic, so it is checked rather than trusted: the cached face must
  // actually draw everything this search's own font can draw of `wanted`. Without
  // this, one collision hands back a face that cannot draw the text and the caller
  // has no way to tell — the reuse is invisible.
  if (existing && coversSameOf(existing, ttf, wanted)) {
    return existing;
  }
  const face = buildRasterFont(ttf);
  if (!existing) {
    _facesByIdentity.set(identity, face);
  }
  return face;
}

/** Whether `cached` draws every code point of `wanted` that `fresh` can draw. */
function coversSameOf(cached: RasterFont, fresh: TtfFont, wanted: ReadonlySet<number>): boolean {
  for (const cp of wanted) {
    if (fresh.cmap.has(cp) && !fontHasGlyph(cached, cp)) {
      return false;
    }
  }
  return true;
}

/**
 * The code points in `text` that no face in `chain` can draw.
 *
 * ASCII is excluded even when uncovered: it cannot usefully steer a font search —
 * a set containing only ASCII matches the first font on the machine — and a chain
 * that cannot draw ASCII has bigger problems than discovery can solve.
 */
function uncoveredBy(chain: readonly RasterFont[], text: string): Set<number> {
  const missing = new Set<number>();
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    // A formatting control has no glyph in any font, so demanding one would
    // disqualify every candidate — see `isNonPrintingControl`.
    if (cp <= 0x7f || isNonPrintingControl(cp) || _unresolvable.has(cp)) {
      continue;
    }
    if (!chain.some(font => fontHasGlyph(font, cp))) {
      missing.add(cp);
    }
  }
  return missing;
}

/**
 * The face that should draw this style's Latin text.
 *
 * A named family has to be honoured *here*, not only in the CJK slot, because the Latin
 * face is ordered first: with `family: "Courier New"` the chain used to put Arial in
 * front and every ASCII character came out of it, so the request was silently ignored
 * for exactly the text it was most visible on. The same is true of bold and italic.
 *
 * A default request still takes the filename path below — the requirement is then
 * "ordinary sans-serif Latin", which every candidate satisfies, and naming them avoids
 * enumerating a font library to answer a question a path can answer.
 */
function styleFace(style: RasterTextStyle): RasterFont | null {
  if (isDefaultLatinRequest(style)) {
    return defaultLatinFace();
  }
  const key = `${style.family ?? ""}|${style.bold ? "b" : ""}${style.italic ? "i" : ""}`;
  if (!_styled.has(key)) {
    // The families list matters as much as the name: searching for "covers A, a and 0"
    // with no family preference answers with whichever face heads the built-in order,
    // which is a CJK one — so `bold` produced a Chinese face for Latin text, and a
    // family that is not installed did the same. Naming the ordinary Latin families
    // after the requested one makes the fallback a Latin fallback.
    const families = [
      ...(style.family === undefined ? [] : [style.family]),
      ...DEFAULT_LATIN_FAMILIES
    ];
    _styled.set(key, discover(LATIN_PROBE, undefined, { ...style, family: undefined }, families));
  }
  return _styled.get(key) ?? defaultLatinFace();
}

/**
 * Ordinary Latin text families, in preference order.
 *
 * Used as the fallback when a style cannot be answered by {@link defaultLatinFace} — a
 * weight, a slant, or a family that is not installed. These are the families whose
 * regular, bold and italic faces a desktop or a CI container is most likely to have.
 */
const DEFAULT_LATIN_FAMILIES: readonly string[] = [
  "arial",
  "helvetica",
  "helvetica neue",
  "liberation sans",
  "dejavu sans",
  "noto sans",
  "segoe ui",
  "tahoma",
  "verdana",
  "roboto"
];

/**
 * Whether this style wants nothing but the ordinary sans-serif Latin face.
 *
 * `DEFAULT_TEXT_FAMILY` is what a display list carries when its producer named no font,
 * so treating it as "no request" is what keeps the common path off the filesystem scan.
 */
function isDefaultLatinRequest(style: RasterTextStyle): boolean {
  if (style.bold || style.italic) {
    return false;
  }
  const family = style.family?.trim().toLowerCase();
  return family === undefined || family === "" || family === DEFAULT_TEXT_FAMILY;
}

/**
 * Code points a Latin face must cover to be worth having.
 *
 * Upper and lower case plus a digit: enough to reject a symbol or icon font, few enough
 * that no real text face fails it.
 */
const LATIN_PROBE: ReadonlySet<number> = new Set([0x41, 0x61, 0x30]);

function defaultLatinFace(): RasterFont | null {
  if (_latin !== undefined) {
    return _latin;
  }
  _latin = null;
  try {
    for (const fontPath of getLatinFontPaths()) {
      if (!fileExistsSync(fontPath)) {
        continue;
      }
      try {
        const font = parseRasterFont(readFileBytesSync(fontPath));
        if (font !== EMPTY_RASTER_FONT) {
          _latin = font;
          return _latin;
        }
      } catch {
        continue; // unreadable or not a font we can parse
      }
    }
  } catch {
    // No filesystem available.
  }
  return _latin;
}

function getLatinFontPaths(): string[] {
  const platform = typeof process !== "undefined" ? process.platform : "";

  if (platform === "darwin") {
    return [
      "/System/Library/Fonts/Supplemental/Arial.ttf",
      "/Library/Fonts/Arial.ttf",
      "/System/Library/Fonts/Helvetica.ttc",
      "/System/Library/Fonts/SFNS.ttf"
    ];
  }
  if (platform === "win32") {
    const windir = process.env.WINDIR || process.env.windir || "C:\\Windows";
    return [
      `${windir}\\Fonts\\arial.ttf`,
      `${windir}\\Fonts\\calibri.ttf`,
      `${windir}\\Fonts\\segoeui.ttf`,
      `${windir}\\Fonts\\tahoma.ttf`
    ];
  }
  return [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/TTF/DejaVuSans.ttf",
    "/usr/share/fonts/noto/NotoSans-Regular.ttf",
    "/usr/share/fonts/truetype/freefont/FreeSans.ttf"
  ];
}
