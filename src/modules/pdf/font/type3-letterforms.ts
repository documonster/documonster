/**
 * Type3 letterform glyphs — Greek and Cyrillic.
 *
 * Every other Type3 table in this directory draws a *symbol*: an arrow, a box, a
 * dingbat, a circled numeral. This one draws *letters*, and it exists because the
 * absence of them was the one gap a caller could not work around. A browser cannot
 * read the host's fonts — `system-fonts.browser.ts` returns nothing by design — so
 * text outside WinAnsi had exactly two sources: a font the caller supplied, or the
 * Type3 fallback. The fallback covered punctuation, currency, arrows and symbols and
 * no alphabet at all, so `Δ` and `Кириллица` reached it and came out as `.notdef`
 * boxes (issue #218). Measured across v0.9.0–v0.12.0, that outcome was byte-identical:
 * it had never worked, in any release.
 *
 * Greek and Cyrillic are the two scripts that turn up inside otherwise-Latin business
 * documents — a delta in an engineering table, a Bulgarian column header — which is
 * why they are worth 168 glyphs and, say, Devanagari is not: an Indic script needs
 * shaping this renderer does not do, so glyphs alone would not make it correct.
 *
 * **These are a last resort, not a substitute for a real face.** A supplied or
 * discovered font always wins, and `requiresEmbeddedFace` still reports every code
 * point here as one only a real font should draw — see `type3-repertoire.ts` for why
 * that predicate and `isType3Drawable` stopped being each other's inverse. What
 * changed is only what happens when there is no font: a legible monoline letter
 * instead of a box.
 *
 * ## Where the outlines come from
 *
 * Nothing is drawn twice. `@draw/raster/stroke-font` already carries a monoline
 * sans-serif for ASCII 32–126, authored for the chart rasteriser, and three
 * observations make it cover most of this block:
 *
 *   - **Shared shapes.** 14 of the 24 Greek capitals and 11 of the 32 Cyrillic
 *     capitals *are* Latin letters — `Α`/`A`, `Ρ`/`P`, `С`/`C` — so they are the same
 *     outline, not a copy of it.
 *   - **Small capitals.** Cyrillic lowercase is largely its uppercase at x-height:
 *     `в к м н п т` and a dozen more. {@link asSmall} derives them.
 *   - **Composition.** Every accented form in both blocks is a base plus a mark, so
 *     `Ё` is `Е` + diaeresis and `ΐ` is `ι` + two marks.
 *
 * Authoring therefore happens in *stroke-font space* — a 0…1 em, Y-down box with the
 * baseline at 0.75 — so a hand-drawn `Δ` and a borrowed `Α` share one set of metrics
 * and cannot drift apart. {@link toGlyphDef} is the single place that converts to the
 * 1000-unit Y-up glyph space PDF wants.
 *
 * @module
 */

import { STROKE_FONT } from "@draw/raster/stroke-font";
import type { GlyphDef, GlyphPen } from "@pdf/font/type3-glyphs";

// =============================================================================
// Stroke-font space
// =============================================================================

/** Top of a capital, in stroke-font space. */
const CAP_TOP = 0.15;
/**
 * Top of a lowercase x-height letter.
 *
 * Read off the shared stroke font rather than chosen: its `a`, `e`, `n`, `o` and the
 * rest all start at 0.38, and {@link asSmall} derives Cyrillic lowercase from the
 * capitals, so a different value here would make a derived `н` a different height
 * from the borrowed `о` beside it.
 */
const X_TOP = 0.38;
/** The baseline. */
const BASE = 0.75;
/** Bottom of a descender. */
const DESC = 0.95;

/**
 * Stroke-space units to glyph units.
 *
 * Fixed by the cap height rather than by the em, because a Type3 glyph is drawn
 * inline with text set in a standard-14 face and what the eye compares is the
 * height of the capitals. 700/1000 is Helvetica's (718) to within a hair, and is
 * what `NOTDEF_GLYPH` already assumes.
 */
const SCALE = 700 / (BASE - CAP_TOP);

/**
 * Stem weight, in glyph units.
 *
 * The source is a monoline font — one weight for every stroke — so this is the whole
 * of the glyph's colour. 80/1000 against a 700 cap reads as a regular weight; a
 * lighter stem disappears at 9 pt and a heavier one looks bold beside Helvetica.
 */
const STEM = 80;

/** A polyline in stroke-font space. Closed when the last point repeats the first. */
type Path = ReadonlyArray<readonly [number, number]>;

/** An authored glyph before conversion: an advance width and its strokes. */
interface StrokeShape {
  readonly w: number;
  readonly d: readonly Path[];
}

// =============================================================================
// Conversion
// =============================================================================

function emit(p: GlyphPen, paths: readonly Path[]): void {
  for (const path of paths) {
    if (path.length === 0) {
      continue;
    }
    const [first] = path;
    // A closed contour is drawn with `Z` rather than a repeated point, so the join
    // at the start is mitred like every other corner instead of leaving a notch.
    const last = path[path.length - 1];
    const closed = path.length > 2 && first[0] === last[0] && first[1] === last[1];
    const points = closed ? path.slice(0, -1) : path;
    p.M(points[0][0], points[0][1]);
    for (let i = 1; i < points.length; i++) {
      p.L(points[i][0], points[i][1]);
    }
    if (closed) {
      p.Z();
    }
  }
}

/**
 * Convert an authored shape into a glyph.
 *
 * The Y flip lives here and only here: stroke-font space runs downward from a
 * top-left origin, and a glyph runs upward from its baseline.
 */
function toGlyphDef(shape: StrokeShape): GlyphDef {
  const paths = shape.d.map(path => path.map(([x, y]) => [x * SCALE, (BASE - y) * SCALE] as const));
  return {
    width: Math.round(shape.w * SCALE),
    draw: p => {
      p.lineWidth(STEM);
      emit(p, paths);
      p.stroke();
    }
  };
}

// =============================================================================
// Reuse
// =============================================================================

/**
 * The shape of an ASCII character, from the shared stroke font.
 *
 * Throws rather than returning a blank: every call site names a literal, so a miss
 * is a typo in this file and a silently empty glyph would ship as an invisible
 * letter.
 *
 * The points are rebuilt as pairs rather than cast. `StrokeGlyph` types a stroke as
 * `number[][]`, which a cast to a tuple cannot narrow soundly — a three-element point
 * would type-check and then be drawn with its third value ignored.
 */
function ascii(ch: string): StrokeShape {
  const glyph = STROKE_FONT[ch.codePointAt(0)!];
  if (glyph === undefined) {
    throw new Error(`type3 letterforms: stroke font has no '${ch}'`);
  }
  return {
    w: glyph.w,
    d: glyph.d.map(stroke =>
      stroke.map(point => {
        if (point.length < 2) {
          throw new Error(`type3 letterforms: '${ch}' has a point with ${point.length} value(s)`);
        }
        return [point[0], point[1]] as const;
      })
    )
  };
}

/** An authored shape, in stroke-font space. */
function shape(w: number, ...d: Path[]): StrokeShape {
  return { w, d };
}

/** Combine shapes, keeping the widest advance. */
function merge(...parts: StrokeShape[]): StrokeShape {
  return {
    w: Math.max(...parts.map(part => part.w)),
    d: parts.flatMap(part => part.d)
  };
}

/**
 * Remap a shape's vertical band, scaling x by the same ratio so it stays in proportion.
 *
 * The one transform behind small capitals, superscripts and subscripts. Writing three
 * of these separately is how they come to disagree about proportion — a subscript
 * narrower than a superscript of the same digit, for no reason anyone chose.
 */
function scaleTo(
  source: StrokeShape,
  top: number,
  bottom: number,
  fromTop = CAP_TOP,
  fromBottom = BASE
): StrokeShape {
  const ratio = (bottom - top) / (fromBottom - fromTop);
  return {
    w: source.w * ratio,
    d: source.d.map(path => path.map(([x, y]) => [x * ratio, top + (y - fromTop) * ratio] as const))
  };
}

/**
 * Compress a capital onto the x-height, for the Cyrillic lowercase that are small
 * capitals rather than distinct shapes — `в`, `к`, `м`, `н`, `п`, `т` and the rest.
 *
 * The extra sidebearing keeps a narrowed letter off its neighbour, which scaling the
 * advance alone would not.
 */
function asSmall(source: StrokeShape): StrokeShape {
  const ratio = (BASE - X_TOP) / (BASE - CAP_TOP);
  const scaled = scaleTo(source, X_TOP, BASE);
  return { w: scaled.w + (1 - ratio) * 0.06, d: scaled.d };
}

/**
 * Raise a shape to superscript size and position — `m²`, a footnote marker.
 *
 * The band sits inside the cap height rather than above it, which is what keeps a
 * superscript clear of the line above.
 */
function asSuperscript(source: StrokeShape, fromTop = CAP_TOP): StrokeShape {
  return scaleTo(source, 0.13, 0.47, fromTop);
}

/** Drop a shape to subscript size and position — `H₂O`, `CO₂`. */
function asSubscript(source: StrokeShape, fromTop = CAP_TOP): StrokeShape {
  return scaleTo(source, 0.47, 0.81, fromTop);
}

/**
 * Turn a shape through half a turn about its own centre.
 *
 * One letter needs it — `ə`, which is a rotated `e` — and drawing it by hand would
 * mean a second `e` free to drift from the first.
 */
function rotate180(source: StrokeShape): StrokeShape {
  const points = source.d.flat();
  const xs = points.map(([x]) => x);
  const ys = points.map(([, y]) => y);
  const cx = (Math.min(...xs) + Math.max(...xs)) / 2;
  const cy = (Math.min(...ys) + Math.max(...ys)) / 2;
  return {
    w: source.w,
    d: source.d.map(path => path.map(([x, y]) => [2 * cx - x, 2 * cy - y] as const))
  };
}

/** The ink extent of a shape, for a mark that has to attach to the letter itself. */
function bounds(source: StrokeShape): {
  minX: number;
  maxX: number;
  minY: number;
  maxY: number;
} {
  const points = source.d.flat();
  const xs = points.map(([x]) => x);
  const ys = points.map(([, y]) => y);
  return {
    minX: Math.min(...xs),
    maxX: Math.max(...xs),
    minY: Math.min(...ys),
    maxY: Math.max(...ys)
  };
}

/**
 * Reflect a shape about its own vertical centre.
 *
 * What makes the turned and reversed letters of Latin Extended-B reuse rather than
 * redraw: `Ɔ` is a reversed `C`, `Ƨ` a reversed `S`, `ᴎ` a reversed `N`.
 */
function mirrorX(source: StrokeShape): StrokeShape {
  const { minX, maxX } = bounds(source);
  const axis = minX + maxX;
  return {
    w: source.w,
    d: source.d.map(path => path.map(([x, y]) => [axis - x, y] as const))
  };
}

/** Shift a shape, for a mark that sits over a narrow or a wide letter. */
function shift(source: StrokeShape, dx: number, dy = 0): StrokeShape {
  return {
    w: source.w,
    d: source.d.map(path => path.map(([x, y]) => [x + dx, y + dy] as const))
  };
}

/**
 * A closed oval as a polyline, for the bowls Greek lowercase is built from.
 *
 * Twelve segments. The stroke font approximates its own `O` with ten and reads
 * cleanly at text sizes, and the count matters more here than there because these
 * are vector outlines in a PDF rather than pixels: a reader may zoom.
 */
function oval(cx: number, cy: number, rx: number, ry: number, segments = 12): Path {
  const points: Array<readonly [number, number]> = [];
  for (let i = 0; i <= segments; i++) {
    const angle = (i / segments) * Math.PI * 2;
    points.push([
      Number((cx + Math.cos(angle) * rx).toFixed(4)),
      Number((cy + Math.sin(angle) * ry).toFixed(4))
    ]);
  }
  return points;
}

// =============================================================================
// Marks
// =============================================================================

/**
 * Acute / Greek tonos, standing on `y`.
 *
 * **Every mark that goes above a letter is anchored on its bottom edge**, not its top.
 * The marks are different heights — a dot is 0.01 em and a ring 0.10 — so anchoring
 * them on the top put each one a different distance from the letter, and the composer
 * had no single offset that cleared them all. It measurably did not: the caron over `Ť`
 * reached 0.144 em while the `T` bar reached 0.116, so they overlapped by a third of a
 * stem and the letter rendered as `†`. Bottom-anchored, one clearance is right for all
 * of them.
 *
 * The length is not cosmetic either. At 0.07 em over a 0.08 em stem the mark is barely
 * longer than the pen is wide, so it renders as a blob and `Ѐ` reads as `Ė` — a
 * different letter. 0.10 em is where the slant becomes the thing the eye sees.
 */
function acute(x: number, y: number): Path[] {
  return [
    [
      [x, y],
      [x + 0.1, y - 0.1]
    ]
  ];
}

/** Grave — the same stroke, mirrored. See {@link acute} for the anchor and length. */
function grave(x: number, y: number): Path[] {
  return [
    [
      [x, y - 0.1],
      [x + 0.1, y]
    ]
  ];
}

/** Diaeresis / dialytika: two dots, standing on `y`. */
function diaeresis(x: number, y: number): Path[] {
  return [
    [
      [x, y - 0.01],
      [x + 0.01, y - 0.01]
    ],
    [
      [x + 0.13, y - 0.01],
      [x + 0.14, y - 0.01]
    ]
  ];
}

/** Breve, as on `Й` `Ў` `ă` `ğ`. Standing on `y`. */
function breve(x: number, y: number, w: number): Path[] {
  return [
    [
      [x, y - 0.07],
      [x + w * 0.2, y - 0.02],
      [x + w * 0.5, y],
      [x + w * 0.8, y - 0.02],
      [x + w, y - 0.07]
    ]
  ];
}

/** Caron / haček, as on `č` `š` `ž` `ě` `ř`. Standing on `y`. */
function caron(x: number, y: number): Path[] {
  return [
    [
      [x, y - 0.09],
      [x + 0.07, y],
      [x + 0.14, y - 0.09]
    ]
  ];
}

/** Circumflex — the caron inverted. Standing on `y`. */
function circumflex(x: number, y: number): Path[] {
  return [
    [
      [x, y],
      [x + 0.07, y - 0.09],
      [x + 0.14, y]
    ]
  ];
}

/** Tilde, as on `ã` `ñ` `õ`. Standing on `y`. */
function tilde(x: number, y: number): Path[] {
  return [
    [
      [x, y - 0.01],
      [x + 0.04, y - 0.06],
      [x + 0.1, y - 0.01],
      [x + 0.14, y - 0.06]
    ]
  ];
}

/** Macron, as on `ā` `ē` `ī` `ō` `ū`. Standing on `y`. */
function macron(x: number, y: number): Path[] {
  return [
    [
      [x, y],
      [x + 0.16, y]
    ]
  ];
}

/** Dot above, as on `ż` `ė` `ġ` and the Turkish `İ`. Standing on `y`. */
function dotAbove(x: number, y: number): Path[] {
  return [
    [
      [x + 0.07, y],
      [x + 0.08, y]
    ]
  ];
}

/** Ring above, as on the Czech `ů`. Standing on `y`. */
function ring(x: number, y: number): Path[] {
  return [oval(x + 0.07, y - 0.05, 0.05, 0.05, 8)];
}

/** Double acute, as on the Hungarian `ő` `ű`. Standing on `y`. */
function doubleAcute(x: number, y: number): Path[] {
  return [...acute(x - 0.01, y), ...acute(x + 0.09, y)];
}

/** Cedilla, hanging below the baseline — `ç` `ş` `ģ`. */
function cedilla(x: number, y: number): Path[] {
  return [
    [
      [x + 0.07, y],
      [x + 0.07, y + 0.05],
      [x + 0.01, y + 0.09]
    ]
  ];
}

/** Comma below, which is what Romanian `ș` `ț` actually take. */
function commaBelow(x: number, y: number): Path[] {
  return [
    [
      [x + 0.08, y + 0.01],
      [x + 0.03, y + 0.1]
    ]
  ];
}

/**
 * Ogonek, the hook under `ą` `ę` `į` `ų`. Hanging from `y`.
 *
 * A hook opening to the right, not the zigzag this was first: four points alternating
 * left and right read as a small `v` at text size, and `ą` has to stay distinguishable
 * from `a` at 9 pt to be worth drawing at all.
 */
function ogonek(x: number, y: number): Path[] {
  // Kept shallow and turned outward. A hook reaching as deep as a descender makes `ą`
  // read as `q` at text size, which is the one confusion this letter cannot afford.
  return [
    [
      [x + 0.09, y],
      [x + 0.05, y + 0.03],
      [x + 0.04, y + 0.06],
      [x + 0.08, y + 0.09],
      [x + 0.13, y + 0.08]
    ]
  ];
}

/**
 * How far a mark's baseline sits above the letter's topmost outline.
 *
 * A stroke is `STEM` wide and centred on its path, so two paths this far apart clear
 * each other by half a stem. Anything less and the ink merges — which is what turned
 * `Ť` into `†`.
 */
const MARK_CLEARANCE = 0.1;

/** Put a mark over a capital, or over a lowercase letter with an ascender. */
function withMarkOver(base: StrokeShape, mark: (x: number, y: number) => Path[]): StrokeShape {
  return merge(base, shape(base.w, ...mark(base.w * 0.5 - 0.03, CAP_TOP - MARK_CLEARANCE)));
}

/** Put a mark over an x-height letter. */
function withMarkOverSmall(base: StrokeShape, mark: (x: number, y: number) => Path[]): StrokeShape {
  return merge(base, shape(base.w, ...mark(base.w * 0.5 - 0.03, X_TOP - MARK_CLEARANCE)));
}

/**
 * Put a raised apostrophe to the right of a letter — Czech and Slovak `ď` `ľ` `ť`.
 *
 * Not a caron above, which is what these took first. The letters have ascenders, so the
 * mark ended up floating over the stem with a gap wider than the mark itself, and at
 * text size `ť` read as `Í`. The apostrophe is also simply what the orthography uses.
 */
function withApostrophe(base: StrokeShape): StrokeShape {
  return merge(
    base,
    shape(base.w + 0.08, [
      // Leaning right, like the apostrophe it is: up from just above the x-height to the
      // cap height. Drawn the other way it leans like a backslash, which is a mark no
      // orthography uses.
      [base.w - 0.02, CAP_TOP + 0.12],
      [base.w + 0.01, CAP_TOP]
    ])
  );
}

/**
 * Hook above — the Vietnamese `dấu hỏi`, as on `ả` `ể` `ủ`. Standing on `y`.
 *
 * A question mark without its dot: down the left, then a bowl to the right.
 */
function hookAbove(x: number, y: number): Path[] {
  return [
    [
      [x + 0.02, y - 0.1],
      [x + 0.07, y - 0.11],
      [x + 0.1, y - 0.08],
      [x + 0.08, y - 0.04],
      [x + 0.05, y - 0.02],
      [x + 0.05, y]
    ]
  ];
}

/** Dot below — the Vietnamese `dấu nặng`, as on `ạ` `ệ` `ụ`. Hanging from `y`. */
function dotBelow(x: number, y: number): Path[] {
  return [
    [
      [x + 0.07, y + 0.04],
      [x + 0.08, y + 0.04]
    ]
  ];
}

/**
 * The horn of `ơ` `ư` `Ơ` `Ư`, attached to the letter's own top right.
 *
 * Not a floating mark: Vietnamese draws it touching the bowl, so it is placed from the
 * base's ink rather than from its advance — an `o` and a `u` present a different corner.
 */
function withHorn(base: StrokeShape): StrokeShape {
  const { maxX, minY } = bounds(base);
  return merge(
    base,
    shape(base.w + 0.05, [
      [maxX - 0.01, minY + 0.04],
      [maxX + 0.04, minY],
      [maxX + 0.06, minY - 0.05]
    ])
  );
}

/**
 * Stack a second mark over the first — Vietnamese `ế`, polytonic Greek `ᾅ`.
 *
 * This is why those scripts were wrongly excluded. The renderer does no *runtime* mark
 * positioning, so a decomposed `e` + U+0302 + U+0301 cannot be drawn correctly. But a
 * precomposed code point is one glyph whose outline this file authors in full, and
 * nothing stops that outline carrying two marks. The clearance between them is tighter
 * than between a mark and a letter: both are thin strokes, and real type sets them close.
 */
const STACK_CLEARANCE = 0.09;

/** Heights of the marks that can be stacked under another, in stroke-font units. */
const MARK_HEIGHT: ReadonlyMap<(x: number, y: number) => Path[], number> = new Map([
  [acute, 0.1],
  [grave, 0.1],
  [circumflex, 0.09],
  [caron, 0.09],
  [tilde, 0.06],
  [breveWide, 0.07],
  [hookAbove, 0.11],
  [macron, 0.0],
  [dotAbove, 0.0],
  [diaeresis, 0.01],
  [ring, 0.1]
]);

/** The breve at a fixed width, so it can be stacked like the other marks. */
function breveWide(x: number, y: number): Path[] {
  return breve(x, y, 0.16);
}

/**
 * Hang a mark under a letter, on the baseline.
 *
 * Attached to the *stem*, not to the centre of the advance: a cedilla belongs under the
 * upright of a `ş`, and centring it under a round `ç` and a narrow `ţ` alike leaves one
 * of them visibly off its letter.
 */
function withMarkUnder(
  base: StrokeShape,
  mark: (x: number, y: number) => Path[],
  stemX = 0.5
): StrokeShape {
  return merge(base, shape(base.w, ...mark(base.w * stemX - 0.04, BASE + 0.02)));
}

/**
 * Put two marks over a letter, the second above the first.
 *
 * The lower mark's height decides where the upper one stands, which is why
 * {@link MARK_HEIGHT} exists: guessing a single offset would either overlap a tall lower
 * mark or float above a flat one.
 */
function withMarkStack(
  base: StrokeShape,
  lower: (x: number, y: number) => Path[],
  upper: (x: number, y: number) => Path[],
  small = true
): StrokeShape {
  const top = small ? X_TOP : CAP_TOP;
  const x = base.w * 0.5 - 0.03;
  const lowerBottom = top - MARK_CLEARANCE;
  const lowerHeight = MARK_HEIGHT.get(lower);
  if (lowerHeight === undefined) {
    throw new Error("type3 letterforms: stacking needs the lower mark's height");
  }
  return merge(
    base,
    shape(base.w, ...lower(x, lowerBottom)),
    shape(base.w, ...upper(x, lowerBottom - lowerHeight - STACK_CLEARANCE))
  );
}

/**
 * Shift a shape right until none of its ink sits left of the origin, widening the advance
 * to match.
 *
 * A bar or a hook legitimately reaches past the letter on both sides, and on a narrow
 * letter that puts ink at a negative x — outside the box `d1` publishes, so a viewer
 * honouring it clips the left of `Ɨ` and `Ƭ`. Shifting is what a type designer would do:
 * the mark keeps its relation to the letter and the glyph gains a left sidebearing.
 */
function keepInside(source: StrokeShape): StrokeShape {
  const { minX } = bounds(source);
  const margin = 0.02;
  if (minX >= margin) {
    return source;
  }
  const dx = margin - minX;
  return { w: source.w + dx, d: shiftPaths(source, dx) };
}

/**
 * Put a hook on the top of a letter's stem — `Ɓ` `Ƈ` `Ɗ` `Ƙ` of Latin Extended-B.
 *
 * Curving left from the stem's top, which is the shape the block uses throughout.
 */
function withTopHook(base: StrokeShape): StrokeShape {
  const { minX, minY } = bounds(base);
  return keepInside(
    merge(
      base,
      shape(base.w, [
        [minX, minY + 0.02],
        [minX - 0.05, minY - 0.02],
        [minX - 0.09, minY + 0.01]
      ])
    )
  );
}

/**
 * Two letters set as one glyph — `Ĳ` `Œ`, and the digraphs `Ǆ` `ǈ` `ǋ` `Ǳ`.
 *
 * `overlap` is how far the right letter is pulled back over the left, and it is per pair
 * rather than a constant because the kinds of pair differ: `I` and `J` merely tuck
 * together, while the `O` and `E` of `Œ` share a stem. A single overlap gave `Œ` an
 * advance of 1097 units and ink out to 1015 — past the em, and past the box `d1` declares,
 * so a viewer honouring it would have clipped the `E`.
 */
function ligature(left: string, right: string, overlap: number): StrokeShape {
  const a = ascii(left);
  const b = ascii(right);
  const gap = a.w - overlap;
  return {
    w: gap + b.w,
    d: [...a.d, ...b.d.map(path => path.map(([x, y]) => [x + gap, y] as const))]
  };
}

/** A shape's paths moved right, for setting one letter beside another. */
function shiftPaths(source: StrokeShape, dx: number): Path[] {
  return source.d.map(path => path.map(([x, y]) => [x + dx, y] as const));
}

/**
 * Draw a bar through a letter — `ł` `đ` `ħ` `ŧ` `Ł` `Đ`.
 *
 * The bar is placed by the caller rather than derived, because where it crosses is
 * part of the letter: `Ł` takes it low on the stem, `Đ` across the bowl.
 */
function withBar(base: StrokeShape, x1: number, y1: number, x2: number, y2: number): StrokeShape {
  return keepInside(
    merge(
      base,
      shape(base.w, [
        [x1, y1],
        [x2, y2]
      ])
    )
  );
}

export type { StrokeShape, Path };
export { CAP_TOP, X_TOP, BASE, DESC, SCALE, STEM };
export {
  ascii,
  shape,
  merge,
  bounds,
  mirrorX,
  scaleTo,
  asSmall,
  asSuperscript,
  asSubscript,
  rotate180,
  shift,
  oval,
  acute,
  grave,
  diaeresis,
  breve,
  caron,
  circumflex,
  tilde,
  macron,
  dotAbove,
  ring,
  doubleAcute,
  cedilla,
  commaBelow,
  ogonek,
  hookAbove,
  dotBelow,
  breveWide,
  withHorn,
  withMarkStack,
  withTopHook,
  keepInside,
  ligature,
  shiftPaths,
  withMarkOver,
  withMarkOverSmall,
  withMarkUnder,
  withApostrophe,
  withBar,
  toGlyphDef
};
