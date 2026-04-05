/**
 * Tests for TrueType font parsing, subsetting, and embedding.
 */
import { describe, it, expect } from "vitest";
import { parseTtf } from "@pdf/font/ttf-parser";
import { FontManager } from "@pdf/font/font-manager";
import { PdfFontError } from "@pdf/errors";

/**
 * Build a minimal valid TrueType font for testing.
 * Contains glyphs for .notdef and ASCII 'A' (U+0041) and 'B' (U+0042).
 */
export function buildMinimalTtf(): Uint8Array {
  const tables: Array<{ tag: string; data: Uint8Array }> = [];

  // --- head table (54 bytes) ---
  const head = new Uint8Array(54);
  const headV = new DataView(head.buffer);
  headV.setUint32(0, 0x00010000, false); // majorVersion.minorVersion
  headV.setUint32(4, 0x00010000, false); // fontRevision
  headV.setUint32(8, 0, false); // checksumAdjustment
  headV.setUint32(12, 0x5f0f3cf5, false); // magicNumber
  headV.setUint16(16, 0x000b, false); // flags
  headV.setUint16(18, 1000, false); // unitsPerEm
  // 20-35: created/modified (16 bytes zero)
  headV.setInt16(36, 0, false); // xMin
  headV.setInt16(38, -200, false); // yMin
  headV.setInt16(40, 800, false); // xMax
  headV.setInt16(42, 800, false); // yMax
  headV.setUint16(44, 0, false); // macStyle
  headV.setUint16(46, 8, false); // lowestRecPPEM
  headV.setInt16(48, 2, false); // fontDirectionHint
  headV.setInt16(50, 1, false); // indexToLocFormat = long
  headV.setInt16(52, 0, false); // glyphDataFormat
  tables.push({ tag: "head", data: head });

  // --- hhea table (36 bytes) ---
  const hhea = new Uint8Array(36);
  const hheaV = new DataView(hhea.buffer);
  hheaV.setUint32(0, 0x00010000, false); // majorVersion.minorVersion
  hheaV.setInt16(4, 800, false); // ascent
  hheaV.setInt16(6, -200, false); // descent
  hheaV.setInt16(8, 0, false); // lineGap
  hheaV.setUint16(10, 600, false); // advanceWidthMax
  // 12-33: zeros (various int16 fields)
  hheaV.setUint16(34, 3, false); // numOfLongHorMetrics = 3 glyphs
  tables.push({ tag: "hhea", data: hhea });

  // --- maxp table (6 bytes for version 0.5 or 32 bytes for 1.0) ---
  const maxp = new Uint8Array(6);
  const maxpV = new DataView(maxp.buffer);
  maxpV.setUint32(0, 0x00005000, false); // version 0.5 (TrueType)
  maxpV.setUint16(4, 3, false); // numGlyphs
  tables.push({ tag: "maxp", data: maxp });

  // --- OS/2 table (78 bytes min for version 1; 96 for v2+) ---
  const os2 = new Uint8Array(96);
  const os2V = new DataView(os2.buffer);
  os2V.setUint16(0, 4, false); // version 4
  os2V.setInt16(2, 500, false); // xAvgCharWidth
  os2V.setUint16(4, 400, false); // usWeightClass
  // Skip to byte 68
  os2V.setInt16(68, 800, false); // sTypoAscender
  os2V.setInt16(70, -200, false); // sTypoDescender
  os2V.setInt16(72, 0, false); // sTypoLineGap
  os2V.setUint16(74, 800, false); // usWinAscent
  os2V.setUint16(76, 200, false); // usWinDescent
  // v2+ fields at byte 78
  // ulCodePageRange 1-2 (8 bytes)
  // sxHeight at 86
  os2V.setInt16(88, 700, false); // sCapHeight
  tables.push({ tag: "OS/2", data: os2 });

  // --- post table (32 bytes, version 3.0) ---
  const post = new Uint8Array(32);
  const postV = new DataView(post.buffer);
  postV.setUint32(0, 0x00030000, false); // version 3.0
  // italicAngle = 0 (Fixed 16.16), at offset 4
  // underlinePosition, underlineThickness at 8, 10
  postV.setUint32(12, 0, false); // isFixedPitch = false
  tables.push({ tag: "post", data: post });

  // --- name table ---
  // Minimal: one name record for familyName (nameID=1) and postScriptName (nameID=6)
  const familyStr = encodeUtf16BE("TestFont");
  const psStr = encodeUtf16BE("TestFont-Regular");
  const storageData = concatArrays([familyStr, psStr]);
  const nameHeaderSize = 6 + 2 * 12; // 6 byte header + 2 records
  const nameTable = new Uint8Array(nameHeaderSize + storageData.length);
  const nameV = new DataView(nameTable.buffer);
  let off = 0;
  nameV.setUint16(off, 0, false);
  off += 2; // format
  nameV.setUint16(off, 2, false);
  off += 2; // count = 2
  nameV.setUint16(off, nameHeaderSize, false);
  off += 2; // storageOffset

  // Record 1: familyName
  nameV.setUint16(off, 3, false);
  off += 2; // platformID = Windows
  nameV.setUint16(off, 1, false);
  off += 2; // encodingID = Unicode BMP
  nameV.setUint16(off, 0x0409, false);
  off += 2; // languageID = English US
  nameV.setUint16(off, 1, false);
  off += 2; // nameID = 1 (family)
  nameV.setUint16(off, familyStr.length, false);
  off += 2; // length
  nameV.setUint16(off, 0, false);
  off += 2; // offset = 0

  // Record 2: postScriptName
  nameV.setUint16(off, 3, false);
  off += 2;
  nameV.setUint16(off, 1, false);
  off += 2;
  nameV.setUint16(off, 0x0409, false);
  off += 2;
  nameV.setUint16(off, 6, false);
  off += 2; // nameID = 6 (PostScript)
  nameV.setUint16(off, psStr.length, false);
  off += 2;
  nameV.setUint16(off, familyStr.length, false);

  nameTable.set(storageData, nameHeaderSize);
  tables.push({ tag: "name", data: nameTable });

  // --- cmap table (format 4 for BMP) ---
  // Map: U+0041 ('A') → glyph 1, U+0042 ('B') → glyph 2
  const cmapData = buildCmapFormat4([
    { start: 0x41, end: 0x42, delta: -0x40 } // 0x41-0x40=1, 0x42-0x40=2
  ]);
  tables.push({ tag: "cmap", data: cmapData });

  // --- hmtx table (3 metrics: .notdef, A, B) ---
  const hmtx = new Uint8Array(3 * 4);
  const hmtxV = new DataView(hmtx.buffer);
  hmtxV.setUint16(0, 500, false);
  hmtxV.setInt16(2, 0, false); // glyph 0: width 500
  hmtxV.setUint16(4, 600, false);
  hmtxV.setInt16(6, 0, false); // glyph 1 (A): width 600
  hmtxV.setUint16(8, 550, false);
  hmtxV.setInt16(10, 0, false); // glyph 2 (B): width 550
  tables.push({ tag: "hmtx", data: hmtx });

  // --- loca table (long format, 4 entries for 3 glyphs + end) ---
  const loca = new Uint8Array(4 * 4);
  const locaV = new DataView(loca.buffer);
  locaV.setUint32(0, 0, false); // glyph 0 offset
  locaV.setUint32(4, 0, false); // glyph 1 offset (empty glyph)
  locaV.setUint32(8, 0, false); // glyph 2 offset (empty glyph)
  locaV.setUint32(12, 0, false); // end
  tables.push({ tag: "loca", data: loca });

  // --- glyf table (empty - no actual outlines needed for metrics testing) ---
  tables.push({ tag: "glyf", data: new Uint8Array(0) });

  // --- Assemble ---
  return assembleTtfFromTables(tables);
}

function encodeUtf16BE(str: string): Uint8Array {
  const buf = new Uint8Array(str.length * 2);
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    buf[i * 2] = (code >> 8) & 0xff;
    buf[i * 2 + 1] = code & 0xff;
  }
  return buf;
}

function concatArrays(arrays: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const a of arrays) {
    total += a.length;
  }
  const result = new Uint8Array(total);
  let offset = 0;
  for (const a of arrays) {
    result.set(a, offset);
    offset += a.length;
  }
  return result;
}

function buildCmapFormat4(
  segments: Array<{ start: number; end: number; delta: number }>
): Uint8Array {
  // Add the final 0xFFFF segment
  const allSegs = [...segments, { start: 0xffff, end: 0xffff, delta: 1 }];
  const segCount = allSegs.length;

  // cmap header + encoding record + format 4 subtable
  const headerSize = 4; // version + numTables
  const encodingSize = 8; // one encoding record
  const subtableOffset = headerSize + encodingSize;

  const format4Size = 14 + segCount * 8 + 2; // 14 byte header + 4 arrays × segCount × 2 + reservedPad
  const totalSize = subtableOffset + format4Size;
  const buf = new Uint8Array(totalSize);
  const v = new DataView(buf.buffer);
  let off = 0;

  // cmap header
  v.setUint16(off, 0, false);
  off += 2; // version
  v.setUint16(off, 1, false);
  off += 2; // numTables

  // encoding record
  v.setUint16(off, 3, false);
  off += 2; // platformID = Windows
  v.setUint16(off, 1, false);
  off += 2; // encodingID = Unicode BMP
  v.setUint32(off, subtableOffset, false);
  off += 4;

  // format 4 subtable
  v.setUint16(off, 4, false);
  off += 2; // format
  v.setUint16(off, format4Size, false);
  off += 2; // length
  v.setUint16(off, 0, false);
  off += 2; // language
  v.setUint16(off, segCount * 2, false);
  off += 2; // segCountX2
  v.setUint16(off, 0, false);
  off += 2; // searchRange (simplified)
  v.setUint16(off, 0, false);
  off += 2; // entrySelector
  v.setUint16(off, 0, false);
  off += 2; // rangeShift

  // endCode array
  for (const s of allSegs) {
    v.setUint16(off, s.end, false);
    off += 2;
  }
  v.setUint16(off, 0, false);
  off += 2; // reservedPad

  // startCode array
  for (const s of allSegs) {
    v.setUint16(off, s.start, false);
    off += 2;
  }

  // idDelta array
  for (const s of allSegs) {
    v.setInt16(off, s.delta, false);
    off += 2;
  }

  // idRangeOffset array (all zeros = use delta)
  for (let i = 0; i < segCount; i++) {
    v.setUint16(off, 0, false);
    off += 2;
  }

  return buf;
}

function assembleTtfFromTables(tables: Array<{ tag: string; data: Uint8Array }>): Uint8Array {
  const numTables = tables.length;
  const headerSize = 12 + numTables * 16;

  let dataOffset = headerSize;
  const entries: Array<{ tag: string; data: Uint8Array; offset: number }> = [];

  for (const { tag, data } of tables) {
    const paddedLen = (data.length + 3) & ~3;
    const padded = new Uint8Array(paddedLen);
    padded.set(data);
    entries.push({ tag, data: padded, offset: dataOffset });
    dataOffset += paddedLen;
  }

  const result = new Uint8Array(dataOffset);
  const v = new DataView(result.buffer);
  let off = 0;

  // Offset table
  v.setUint32(off, 0x00010000, false);
  off += 4;
  v.setUint16(off, numTables, false);
  off += 2;
  const pow2 = Math.pow(2, Math.floor(Math.log2(numTables)));
  v.setUint16(off, pow2 * 16, false);
  off += 2;
  v.setUint16(off, Math.floor(Math.log2(numTables)), false);
  off += 2;
  v.setUint16(off, numTables * 16 - pow2 * 16, false);
  off += 2;

  // Table directory
  for (const entry of entries) {
    for (let i = 0; i < 4; i++) {
      result[off++] = entry.tag.charCodeAt(i);
    }
    v.setUint32(off, 0, false);
    off += 4; // checksum (skip)
    v.setUint32(off, entry.offset, false);
    off += 4;
    v.setUint32(off, entry.data.length, false);
    off += 4;
  }

  // Table data
  for (const entry of entries) {
    result.set(entry.data, entry.offset);
  }

  return result;
}

// =============================================================================
// Tests
// =============================================================================

describe("TrueType Font Parser", () => {
  it("should parse a minimal TrueType font", () => {
    const ttfData = buildMinimalTtf();
    const font = parseTtf(ttfData);

    expect(font.familyName).toBe("TestFont");
    expect(font.postScriptName).toBe("TestFont-Regular");
    expect(font.unitsPerEm).toBe(1000);
    expect(font.ascent).toBe(800);
    expect(font.descent).toBe(-200);
    expect(font.numGlyphs).toBe(3);
  });

  it("should read cmap correctly", () => {
    const ttfData = buildMinimalTtf();
    const font = parseTtf(ttfData);

    expect(font.cmap.get(0x41)).toBe(1); // 'A' → glyph 1
    expect(font.cmap.get(0x42)).toBe(2); // 'B' → glyph 2
    expect(font.cmap.get(0x43)).toBeUndefined(); // 'C' not mapped
  });

  it("should read advance widths correctly", () => {
    const ttfData = buildMinimalTtf();
    const font = parseTtf(ttfData);

    expect(font.advanceWidths[0]).toBe(500); // .notdef
    expect(font.advanceWidths[1]).toBe(600); // glyph 1 ('A')
    expect(font.advanceWidths[2]).toBe(550); // glyph 2 ('B')
  });

  it("should read font bounding box", () => {
    const ttfData = buildMinimalTtf();
    const font = parseTtf(ttfData);

    expect(font.bbox).toEqual([0, -200, 800, 800]);
  });

  it("should reject CFF OpenType fonts", () => {
    const data = new Uint8Array(64);
    // 'OTTO' signature
    data[0] = 0x4f;
    data[1] = 0x54;
    data[2] = 0x54;
    data[3] = 0x4f;
    expect(() => parseTtf(data)).toThrow(PdfFontError);
  });

  it("should reject invalid data", () => {
    expect(() => parseTtf(new Uint8Array([0, 0, 0, 0, 0, 0]))).toThrow(PdfFontError);
  });
});

describe("Font Embedding Utilities", () => {
  it("should encode text via FontManager", async () => {
    const ttfData = buildMinimalTtf();
    const font = parseTtf(ttfData);
    const fm = new FontManager();
    fm.registerEmbeddedFont(font);
    fm.trackText("AB");

    const { PdfWriter } = await import("@pdf/core/pdf-writer");
    const writer = new PdfWriter();
    fm.writeFontResources(writer);

    const encoded = fm.encodeText("AB", fm.getEmbeddedResourceName());
    expect(encoded).toBe("<00010002>");
  });

  it("should use .notdef (0) for unmapped characters", async () => {
    const ttfData = buildMinimalTtf();
    const font = parseTtf(ttfData);
    const fm = new FontManager();
    fm.registerEmbeddedFont(font);
    fm.trackText("A");

    const { PdfWriter } = await import("@pdf/core/pdf-writer");
    const writer = new PdfWriter();
    fm.writeFontResources(writer);

    const encoded = fm.encodeText("AC", fm.getEmbeddedResourceName());
    expect(encoded).toBe("<00010000>");
  });

  it("should measure text with embedded font metrics", () => {
    const ttfData = buildMinimalTtf();
    const font = parseTtf(ttfData);
    const fm = new FontManager();
    fm.registerEmbeddedFont(font);

    const resourceName = fm.getEmbeddedResourceName();
    const width = fm.measureText("AB", resourceName, 12);
    // A=600, B=550 in font units, unitsPerEm=1000
    // (600 + 550) / 1000 * 12 = 13.8
    expect(width).toBeCloseTo(13.8, 1);
  });
});

describe("Font Integration with excelToPdf", () => {
  it("should export PDF with embedded font", async () => {
    const { Workbook } = await import("@excel/workbook");
    const { excelToPdf } = await import("@pdf/excel-bridge");

    const ttfData = buildMinimalTtf();

    const wb = new Workbook();
    const ws = wb.addWorksheet("Test");
    ws.getCell("A1").value = "AB";

    const pdf = excelToPdf(wb, { font: ttfData });

    expect(pdf).toBeInstanceOf(Uint8Array);
    expect(pdf.length).toBeGreaterThan(100);

    const text = new TextDecoder().decode(pdf);
    expect(text).toContain("%PDF-2.0");
    expect(text).toContain("%%EOF");
    expect(text).toContain("/Type0");
    expect(text).toContain("/CIDFontType2");
    expect(text).toContain("/Identity-H");
    expect(text).toContain("TestFont-Regular-Subset");
    expect(text).toContain("/FlateDecode");
    expect(text).toContain("/ToUnicode");
    expect(text).toContain("<00010002> Tj"); // 'AB' as subset GIDs 1,2
  });

  it("should correctly remap non-sequential glyph IDs in subset", async () => {
    // This test catches the critical bug where original GIDs are used instead of
    // remapped subset GIDs. The font maps A→GID 5 and B→GID 8. After subsetting
    // to [.notdef, A, B], the new GIDs should be [0, 1, 2].
    // Content stream must contain <00010002>, NOT <00050008>.
    const { Workbook } = await import("@excel/workbook");
    const { excelToPdf } = await import("@pdf/excel-bridge");

    const ttfData = buildSparseGidTtf();

    const wb = new Workbook();
    const ws = wb.addWorksheet("Test");
    ws.getCell("A1").value = "AB";

    const pdf = excelToPdf(wb, { font: ttfData });
    const text = new TextDecoder().decode(pdf);

    expect(text).toContain("%PDF-2.0");
    // Subset GIDs: .notdef=0, A=1, B=2
    expect(text).toContain("<00010002> Tj");
    // Must NOT contain original GIDs
    expect(text).not.toContain("<00050008>");
  });
});

/**
 * Build a TTF where A→GID 5, B→GID 8 (non-sequential, non-starting-from-1).
 * This ensures the subset GID remapping is actually tested.
 */
function buildSparseGidTtf(): Uint8Array {
  const tables: Array<{ tag: string; data: Uint8Array }> = [];

  // head — same as minimal
  const head = new Uint8Array(54);
  const headV = new DataView(head.buffer);
  headV.setUint32(0, 0x00010000, false);
  headV.setUint32(4, 0x00010000, false);
  headV.setUint32(12, 0x5f0f3cf5, false);
  headV.setUint16(16, 0x000b, false);
  headV.setUint16(18, 1000, false);
  headV.setInt16(36, 0, false);
  headV.setInt16(38, -200, false);
  headV.setInt16(40, 800, false);
  headV.setInt16(42, 800, false);
  headV.setInt16(50, 1, false); // long loca
  tables.push({ tag: "head", data: head });

  // hhea — 10 glyphs
  const hhea = new Uint8Array(36);
  const hheaV = new DataView(hhea.buffer);
  hheaV.setUint32(0, 0x00010000, false);
  hheaV.setInt16(4, 800, false);
  hheaV.setInt16(6, -200, false);
  hheaV.setUint16(34, 10, false); // numOfLongHorMetrics = 10
  tables.push({ tag: "hhea", data: hhea });

  // maxp — 10 glyphs
  const maxp = new Uint8Array(6);
  const maxpV = new DataView(maxp.buffer);
  maxpV.setUint32(0, 0x00005000, false);
  maxpV.setUint16(4, 10, false); // numGlyphs = 10
  tables.push({ tag: "maxp", data: maxp });

  // post
  const post = new Uint8Array(32);
  new DataView(post.buffer).setUint32(0, 0x00030000, false);
  tables.push({ tag: "post", data: post });

  // name
  const familyStr = encodeUtf16BE("SparseFont");
  const psStr = encodeUtf16BE("SparseFont-Regular");
  const storageData = concatArrays([familyStr, psStr]);
  const nameHeaderSize = 6 + 2 * 12;
  const nameTable = new Uint8Array(nameHeaderSize + storageData.length);
  const nameV = new DataView(nameTable.buffer);
  let noff = 0;
  nameV.setUint16(noff, 0, false);
  noff += 2;
  nameV.setUint16(noff, 2, false);
  noff += 2;
  nameV.setUint16(noff, nameHeaderSize, false);
  noff += 2;
  // family
  nameV.setUint16(noff, 3, false);
  noff += 2;
  nameV.setUint16(noff, 1, false);
  noff += 2;
  nameV.setUint16(noff, 0x0409, false);
  noff += 2;
  nameV.setUint16(noff, 1, false);
  noff += 2;
  nameV.setUint16(noff, familyStr.length, false);
  noff += 2;
  nameV.setUint16(noff, 0, false);
  noff += 2;
  // postscript
  nameV.setUint16(noff, 3, false);
  noff += 2;
  nameV.setUint16(noff, 1, false);
  noff += 2;
  nameV.setUint16(noff, 0x0409, false);
  noff += 2;
  nameV.setUint16(noff, 6, false);
  noff += 2;
  nameV.setUint16(noff, psStr.length, false);
  noff += 2;
  nameV.setUint16(noff, familyStr.length, false);
  nameTable.set(storageData, nameHeaderSize);
  tables.push({ tag: "name", data: nameTable });

  // cmap — A(0x41)→GID 5, B(0x42)→GID 8
  // Using delta: 0x41 + delta = 5 → delta = 5 - 0x41 = -60
  // But 0x42 + delta would be -60+0x42=6, not 8. So we need separate segments.
  const cmapData = buildCmapFormat4([
    { start: 0x41, end: 0x41, delta: 5 - 0x41 }, // A→5
    { start: 0x42, end: 0x42, delta: 8 - 0x42 } // B→8
  ]);
  tables.push({ tag: "cmap", data: cmapData });

  // hmtx — 10 glyphs
  const hmtx = new Uint8Array(10 * 4);
  const hmtxV = new DataView(hmtx.buffer);
  for (let i = 0; i < 10; i++) {
    hmtxV.setUint16(i * 4, 500 + i * 10, false); // widths: 500, 510, 520...
    hmtxV.setInt16(i * 4 + 2, 0, false);
  }
  tables.push({ tag: "hmtx", data: hmtx });

  // loca — 10 glyphs + end, all zero (empty outlines)
  const loca = new Uint8Array(11 * 4);
  tables.push({ tag: "loca", data: loca });

  // glyf — empty
  tables.push({ tag: "glyf", data: new Uint8Array(0) });

  return assembleTtfFromTables(tables);
}
