/**
 * Synthetic TrueType fonts, built in memory.
 *
 * Shared across modules because three of them need the same thing and the host
 * cannot provide it: a face covering an exact, known set of code points. A test
 * asserting "Chinese is drawn" against a discovered system font passes on a
 * developer's macOS and fails on a CI container with no CJK font installed, which
 * says nothing about the code under test.
 *
 * Used by the raster fallback tests (`draw`) and the font subsetter (`word`). It
 * lives here rather than beside either because a fixture only one module can reach
 * is a fixture the next module will rebuild — `pdf/__tests__/ttf-test-utils.ts` and
 * this file already existed separately for the same reason.
 *
 * @module
 */

/** Assemble a font file from tables, padding each to a 4-byte boundary. */
export function assemble(tables: Array<{ tag: string; data: Uint8Array }>): Uint8Array {
  const header = 12 + tables.length * 16;
  let cursor = header;
  const placed = tables.map(({ tag, data }) => {
    const padded = new Uint8Array((data.length + 3) & ~3);
    padded.set(data);
    const entry = { tag, data: padded, offset: cursor };
    cursor += padded.length;
    return entry;
  });

  const file = new Uint8Array(cursor);
  const v = new DataView(file.buffer);
  v.setUint32(0, 0x00010000, false);
  v.setUint16(4, tables.length, false);
  placed.forEach((entry, i) => {
    const rec = 12 + i * 16;
    for (let c = 0; c < 4; c++) {
      file[rec + c] = entry.tag.charCodeAt(c);
    }
    v.setUint32(rec + 8, entry.offset, false);
    v.setUint32(rec + 12, entry.data.length, false);
    file.set(entry.data, entry.offset);
  });
  return file;
}

/**
 * A filled square from 4 on-curve points, so the glyph actually puts ink down.
 *
 * The two-point contour the metrics fixture uses is degenerate — it flattens to a
 * line and the scan-line fill produces nothing — which is fine for asserting an
 * advance but useless for asserting that something was drawn.
 */
function boxGlyph(): Uint8Array {
  const glyph = new Uint8Array(10 + 2 + 2 + 4 + 8 + 8);
  const v = new DataView(glyph.buffer);
  v.setInt16(0, 1, false); // numberOfContours
  v.setInt16(2, 50, false); // xMin
  v.setInt16(4, 0, false); // yMin
  v.setInt16(6, 650, false); // xMax
  v.setInt16(8, 600, false); // yMax
  v.setUint16(10, 3, false); // endPtsOfContours[0] — 4 points
  v.setUint16(12, 0, false); // instructionLength
  let at = 14;
  for (let i = 0; i < 4; i++) {
    glyph[at++] = 0x01; // on-curve, int16 deltas
  }
  // x deltas: 50, +600, 0, -600  → 50, 650, 650, 50
  for (const dx of [50, 600, 0, -600]) {
    v.setInt16(at, dx, false);
    at += 2;
  }
  // y deltas: 0, 0, +600, 0 → 0, 0, 600, 600
  for (const dy of [0, 0, 600, 0]) {
    v.setInt16(at, dy, false);
    at += 2;
  }
  return glyph;
}

/** Group sorted code points into runs of consecutive values. */
function runsOf(sorted: readonly number[]): { start: number; end: number; firstGid: number }[] {
  const runs: { start: number; end: number; firstGid: number }[] = [];
  sorted.forEach((cp, index) => {
    const last = runs[runs.length - 1];
    if (last && cp === last.end + 1) {
      last.end = cp;
      return;
    }
    // Glyph 0 is `.notdef`, so the first mapped code point is glyph 1.
    runs.push({ start: cp, end: cp, firstGid: index + 1 });
  });
  return runs;
}

/**
 * A font that draws exactly `codePoints` and nothing else.
 *
 * Every mapped code point gets the same filled box, with a 700/1000 em advance unless
 * `advanceUnits` says otherwise. Overriding it is how a test tells the two layout modes
 * apart: 700 is within a pixel of what `@utils/text-measure` reports for Latin, so a
 * fixture at the default cannot show whether a line was spaced by the font or by the
 * measurement.
 * `familyName` is written into the `name` table because the fallback chain filters
 * out a face that reports none.
 *
 * There is deliberately no way to map a code point to glyph 0 here: `readCmapFormat4`
 * in `@utils/font-ttf` drops such an entry while parsing, so a fixture offering the
 * option would be describing something it cannot produce. A `.notdef` mapping is
 * tested against `buildRasterFont` directly instead — see `font-fallback.test.ts`.
 */
export function buildCoverageFont(
  codePoints: readonly number[],
  familyName: string,
  style: FaceStyleFixture = {},
  advanceUnits = 700
): Uint8Array {
  const sorted = [...new Set(codePoints)].sort((a, b) => a - b);
  const numGlyphs = sorted.length + 1; // + .notdef

  const head = new Uint8Array(54);
  const headV = new DataView(head.buffer);
  headV.setUint32(0, 0x00010000, false);
  headV.setUint16(18, 1000, false); // unitsPerEm
  // macStyle bit 0 is bold, bit 1 is italic. `parseTtf` reads the slant from here.
  headV.setUint16(44, (style.bold ? 0x01 : 0) | (style.italic ? 0x02 : 0), false);
  headV.setInt16(50, 1, false); // indexToLocFormat = long

  const hhea = new Uint8Array(36);
  const hheaV = new DataView(hhea.buffer);
  hheaV.setUint32(0, 0x00010000, false);
  hheaV.setInt16(4, 800, false); // ascent
  hheaV.setInt16(6, -200, false); // descent
  hheaV.setUint16(34, numGlyphs, false); // numberOfHMetrics

  const maxp = new Uint8Array(6);
  new DataView(maxp.buffer).setUint16(4, numGlyphs, false);

  // --- cmap: format 4, one segment per run, plus the required 0xFFFF sentinel ---
  const runs = runsOf(sorted);
  const segments = [
    ...runs.map(run => ({ start: run.start, end: run.end, delta: run.firstGid - run.start })),
    { start: 0xffff, end: 0xffff, delta: 1 }
  ];
  const cmap = buildCmapFormat4(segments);

  const hmtx = new Uint8Array(numGlyphs * 4);
  const hmtxV = new DataView(hmtx.buffer);
  for (let gid = 0; gid < numGlyphs; gid++) {
    hmtxV.setUint16(gid * 4, advanceUnits, false); // advanceWidth
    hmtxV.setInt16(gid * 4 + 2, 50, false); // leftSideBearing, matching xMin
  }

  // --- glyf / loca: .notdef is empty, every mapped glyph is the same box ---
  const box = boxGlyph();
  const loca = new Uint8Array((numGlyphs + 1) * 4);
  const locaV = new DataView(loca.buffer);
  const glyf = new Uint8Array(box.length * (numGlyphs - 1));
  let writeAt = 0;
  for (let gid = 0; gid < numGlyphs; gid++) {
    locaV.setUint32(gid * 4, writeAt, false);
    if (gid > 0) {
      glyf.set(box, writeAt);
      writeAt += box.length;
    }
  }
  locaV.setUint32(numGlyphs * 4, writeAt, false);

  return assemble([
    { tag: "head", data: head },
    { tag: "hhea", data: hhea },
    { tag: "maxp", data: maxp },
    { tag: "cmap", data: cmap },
    { tag: "hmtx", data: hmtx },
    { tag: "loca", data: loca },
    { tag: "glyf", data: glyf },
    { tag: "name", data: nameTable(familyName) },
    { tag: "OS/2", data: os2Table(style.weight ?? 400) }
  ]);
}

/** The weight and slant a fixture face should report. */
export interface FaceStyleFixture {
  /** OS/2 `usWeightClass`. Defaults to 400. */
  readonly weight?: number;
  /** Sets `head.macStyle` bit 1, which is where `parseTtf` reads the slant from. */
  readonly italic?: boolean;
  /** Sets `head.macStyle` bit 0. Weight is what selection actually ranks on. */
  readonly bold?: boolean;
}

/**
 * A version 1 `OS/2` table carrying a weight class.
 *
 * Without it `parseTtf` reports the default 400 for every face, so a fixture could not
 * express "the bold one" and style selection could only be tested against whatever the
 * host happened to have installed.
 */
function os2Table(weight: number): Uint8Array {
  // Version 1 is 86 bytes; every field but the weight can stay zero.
  const table = new Uint8Array(86);
  const v = new DataView(table.buffer);
  v.setUint16(0, 1, false); // version
  v.setUint16(4, weight, false); // usWeightClass
  v.setInt16(68, 800, false); // sTypoAscender
  v.setInt16(70, -200, false); // sTypoDescender
  return table;
}

/** A `cmap` table with one format 4 subtable for `segments`. */
function buildCmapFormat4(
  segments: readonly { start: number; end: number; delta: number }[]
): Uint8Array {
  const subtableLength = 14 + segments.length * 8 + 2;
  const cmap = new Uint8Array(12 + subtableLength);
  const v = new DataView(cmap.buffer);
  v.setUint16(2, 1, false); // numTables
  v.setUint16(4, 3, false); // platformID: Windows
  v.setUint16(6, 1, false); // encodingID: Unicode BMP
  v.setUint32(8, 12, false); // subtable offset
  v.setUint16(12, 4, false); // format
  v.setUint16(14, subtableLength, false);
  v.setUint16(18, segments.length * 2, false); // segCountX2
  let at = 26;
  for (const s of segments) {
    v.setUint16(at, s.end, false);
    at += 2;
  }
  at += 2; // reservedPad
  for (const s of segments) {
    v.setUint16(at, s.start, false);
    at += 2;
  }
  for (const s of segments) {
    v.setInt16(at, s.delta, false);
    at += 2;
  }
  // idRangeOffset stays 0 for every segment: the deltas above are the mapping.
  return cmap;
}

/** A `name` table carrying one Windows-Unicode family-name record. */
function nameTable(familyName: string): Uint8Array {
  const value = new Uint8Array(familyName.length * 2);
  const valueV = new DataView(value.buffer);
  for (let i = 0; i < familyName.length; i++) {
    valueV.setUint16(i * 2, familyName.charCodeAt(i), false);
  }
  const table = new Uint8Array(6 + 12 + value.length);
  const v = new DataView(table.buffer);
  v.setUint16(0, 0, false); // format
  v.setUint16(2, 1, false); // count
  v.setUint16(4, 6 + 12, false); // stringOffset
  v.setUint16(6, 3, false); // platformID: Windows
  v.setUint16(8, 1, false); // encodingID: Unicode BMP
  v.setUint16(10, 0x0409, false); // languageID: en-US
  v.setUint16(12, 1, false); // nameID: family
  v.setUint16(14, value.length, false);
  v.setUint16(16, 0, false); // offset into the string storage
  table.set(value, 18);
  return table;
}

/** A font file that parses as a container but carries no outlines. */
export function fontWithoutGlyf(): Uint8Array {
  const head = new Uint8Array(54);
  const headV = new DataView(head.buffer);
  headV.setUint32(0, 0x00010000, false);
  headV.setUint16(18, 1000, false);
  return assemble([{ tag: "head", data: head }]);
}

/**
 * A font with many glyphs, so a subsetter has something worth discarding.
 *
 * `subsetFont` refuses when the kept set is most of the font, so a fixture with a
 * handful of glyphs can only ever exercise the refusal path — which is why Word's
 * subsetter had no test that actually subsetted anything.
 *
 * Glyph 1 is a composite referencing glyph 2, so dependency resolution is covered:
 * keeping the character mapped to glyph 1 must keep glyph 2 as well.
 */
export function buildSubsettableFont(
  codePoints: readonly number[],
  familyName: string
): Uint8Array {
  const sorted = [...new Set(codePoints)].sort((a, b) => a - b);
  // Pad the glyph count so the kept subset stays well under the refusal threshold.
  const numGlyphs = Math.max(sorted.length + 2, 40);

  const head = new Uint8Array(54);
  const headV = new DataView(head.buffer);
  headV.setUint32(0, 0x00010000, false);
  headV.setUint32(8, 0x5f0f3cf5, false); // checkSumAdjustment: a stale value to clear
  headV.setUint16(18, 1000, false);
  headV.setInt16(50, 1, false); // indexToLocFormat = long

  const hhea = new Uint8Array(36);
  const hheaV = new DataView(hhea.buffer);
  hheaV.setUint32(0, 0x00010000, false);
  hheaV.setInt16(4, 800, false);
  hheaV.setInt16(6, -200, false);
  hheaV.setUint16(34, numGlyphs, false);

  const maxp = new Uint8Array(6);
  new DataView(maxp.buffer).setUint16(4, numGlyphs, false);

  const runs = runsOf(sorted);
  const segments = [
    ...runs.map(run => ({ start: run.start, end: run.end, delta: run.firstGid - run.start })),
    { start: 0xffff, end: 0xffff, delta: 1 }
  ];
  const cmap = buildCmapFormat4(segments);

  const hmtx = new Uint8Array(numGlyphs * 4);
  const hmtxV = new DataView(hmtx.buffer);
  for (let gid = 0; gid < numGlyphs; gid++) {
    hmtxV.setUint16(gid * 4, 700, false);
    hmtxV.setInt16(gid * 4 + 2, 50, false);
  }

  // glyph 0: empty. glyph 1: composite of glyph 2. glyph 2+: boxes.
  const box = boxGlyph();
  const composite = compositeOf(2);
  const bodies: Uint8Array[] = [new Uint8Array(0), composite];
  for (let gid = 2; gid < numGlyphs; gid++) {
    bodies.push(box);
  }

  const loca = new Uint8Array((numGlyphs + 1) * 4);
  const locaV = new DataView(loca.buffer);
  let total = 0;
  bodies.forEach((body, gid) => {
    locaV.setUint32(gid * 4, total, false);
    total += body.length;
  });
  locaV.setUint32(numGlyphs * 4, total, false);
  const glyf = new Uint8Array(total);
  let at = 0;
  for (const body of bodies) {
    glyf.set(body, at);
    at += body.length;
  }

  return assemble([
    { tag: "head", data: head },
    { tag: "hhea", data: hhea },
    { tag: "maxp", data: maxp },
    { tag: "cmap", data: cmap },
    { tag: "hmtx", data: hmtx },
    { tag: "loca", data: loca },
    { tag: "glyf", data: glyf },
    { tag: "name", data: nameTable(familyName) },
    // Present so a test can assert it is *dropped*: the signature cannot survive the
    // font being modified.
    { tag: "DSIG", data: new Uint8Array([0, 0, 0, 1, 0, 0, 0, 0]) }
  ]);
}

/** A composite glyph made of one component, placed at the origin. */
function compositeOf(component: number): Uint8Array {
  const glyph = new Uint8Array(16);
  const v = new DataView(glyph.buffer);
  v.setInt16(0, -1, false); // composite
  v.setInt16(2, 50, false);
  v.setInt16(4, 0, false);
  v.setInt16(6, 650, false);
  v.setInt16(8, 600, false);
  v.setUint16(10, 0x0002, false); // ARGS_ARE_XY_VALUES, no MORE_COMPONENTS
  v.setUint16(12, component, false);
  glyph[14] = 0;
  glyph[15] = 0;
  return glyph;
}
