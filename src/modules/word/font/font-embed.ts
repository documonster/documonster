/**
 * DOCX Module - Embedded Font Writer
 *
 * Provides the ability to embed font files into DOCX documents.
 * Supports TTF, OTF, and WOFF fonts with optional ODTTF obfuscation
 * as required by the OOXML specification for embedded fonts.
 *
 * ECMA-376 Part 1, §17.8.3 describes how fonts are embedded:
 * - Fonts are stored in word/fonts/ directory
 * - Each embedded font has a relationship in fontTable.xml
 * - Obfuscation uses a GUID-based XOR scheme (ODTTF format)
 *
 * Supports font subsetting to reduce file size by only including
 * glyphs that are actually used in the document.
 */

import type { DocxDocument, FontDef, EmbeddedFont } from "../types";
import { obfuscateFont, generateFontKey } from "./font-obfuscation";

// =============================================================================
// Types
// =============================================================================

/** Font embedding style (which variants to embed). */
export type FontEmbedStyle = "regular" | "bold" | "italic" | "boldItalic";

/** Options for embedding a font. */
export interface EmbedFontOptions {
  /** Font family name as it will appear in the document. */
  readonly name: string;
  /** The raw font file data (TTF/OTF/WOFF). */
  readonly data: Uint8Array;
  /** Which style variant this is. Default: "regular". */
  readonly style?: FontEmbedStyle;
  /** Whether to apply ODTTF obfuscation. Default: true. */
  readonly obfuscate?: boolean;
  /** Font family classification. */
  readonly family?: "roman" | "swiss" | "modern" | "script" | "decorative" | "auto";
  /** Font pitch. */
  readonly pitch?: "default" | "fixed" | "variable";
  /** Panose-1 classification (10-byte hex string). */
  readonly panose1?: string;
  /** Character set (0 = ANSI). */
  readonly charset?: number;
  /**
   * Characters used in the document for subsetting.
   * When provided, the embedded font will only contain glyphs for these characters,
   * significantly reducing file size. Pass all unique characters from the document
   * that use this font.
   */
  readonly usedCharacters?: string;
}

/** Result of embedding a font. */
export interface EmbedFontResult {
  /** The font definition to add to `doc.fonts`. */
  readonly fontDef: FontDef;
  /** The embedded font entry to add to `doc.embeddedFonts`. */
  readonly embeddedFont: EmbeddedFont;
}

// =============================================================================
// Font Subsetting
// =============================================================================

/**
 * Subset a TrueType/OpenType font to include only the specified characters.
 *
 * This performs a minimal subset by:
 * 1. Parsing the cmap table to find glyph IDs for requested characters
 * 2. Building a new glyf/loca table with only needed glyphs (+ composite dependencies)
 * 3. Rebuilding the font with minimal tables
 *
 * For CFF (PostScript-outline) fonts, subsetting is more complex and we fall back
 * to embedding the full font.
 *
 * @param fontData - Raw TTF/OTF font bytes
 * @param characters - String of characters to keep
 * @returns Subsetted font bytes, or original if subsetting fails/not applicable
 */
export function subsetFont(fontData: Uint8Array, characters: string): Uint8Array {
  try {
    return _subsetTtf(fontData, characters);
  } catch {
    // If subsetting fails for any reason, return original data
    return fontData;
  }
}

/** Read a uint16 big-endian from buffer. */
function readU16(data: Uint8Array, offset: number): number {
  return (data[offset] << 8) | data[offset + 1];
}

/** Read a uint32 big-endian from buffer. */
function readU32(data: Uint8Array, offset: number): number {
  return (
    ((data[offset] << 24) |
      (data[offset + 1] << 16) |
      (data[offset + 2] << 8) |
      data[offset + 3]) >>>
    0
  );
}

/** Read an int16 big-endian from buffer. */
function readI16(data: Uint8Array, offset: number): number {
  const v = readU16(data, offset);
  return v > 0x7fff ? v - 0x10000 : v;
}

/** Write uint16 big-endian. */
function writeU16(data: Uint8Array, offset: number, value: number): void {
  data[offset] = (value >> 8) & 0xff;
  data[offset + 1] = value & 0xff;
}

/** Write uint32 big-endian. */
function writeU32(data: Uint8Array, offset: number, value: number): void {
  data[offset] = (value >>> 24) & 0xff;
  data[offset + 1] = (value >>> 16) & 0xff;
  data[offset + 2] = (value >>> 8) & 0xff;
  data[offset + 3] = value & 0xff;
}

interface TableRecord {
  tag: string;
  checksum: number;
  offset: number;
  length: number;
}

function parseTableDirectory(data: Uint8Array): {
  numTables: number;
  tables: Map<string, TableRecord>;
} {
  const numTables = readU16(data, 4);
  const tables = new Map<string, TableRecord>();

  // numTables comes from a u16 → max 65535. Reject directories that
  // claim more entries than the buffer can possibly hold so we don't
  // walk off the end and accumulate garbage records.
  const maxRecords = Math.floor((data.length - 12) / 16);
  if (numTables > maxRecords) {
    return { numTables: 0, tables };
  }

  for (let i = 0; i < numTables; i++) {
    const recOffset = 12 + i * 16;
    const tag = String.fromCharCode(
      data[recOffset],
      data[recOffset + 1],
      data[recOffset + 2],
      data[recOffset + 3]
    );
    tables.set(tag, {
      tag,
      checksum: readU32(data, recOffset + 4),
      offset: readU32(data, recOffset + 8),
      length: readU32(data, recOffset + 12)
    });
  }

  return { numTables, tables };
}

// Hard cap on how many codepoints we materialise from a cmap table.
// Real fonts reach ~150K glyphs at the upper end (CJK pan-Unicode);
// bigger numbers indicate a malformed or hostile table built to drive
// the parser into a multi-gigabyte Map allocation.
const CMAP_MAX_ENTRIES = 200_000;

/**
 * Parse cmap format 4 (BMP characters) to build codepoint → glyph ID map.
 */
function parseCmapFormat4(data: Uint8Array, offset: number): Map<number, number> {
  const map = new Map<number, number>();
  const segCountX2 = readU16(data, offset + 6);
  const segCount = segCountX2 / 2;

  const endCodes = offset + 14;
  const startCodes = endCodes + segCountX2 + 2; // +2 for reservedPad
  const idDeltas = startCodes + segCountX2;
  const idRangeOffsets = idDeltas + segCountX2;

  for (let i = 0; i < segCount; i++) {
    const endCode = readU16(data, endCodes + i * 2);
    const startCode = readU16(data, startCodes + i * 2);
    const idDelta = readI16(data, idDeltas + i * 2);
    const idRangeOffset = readU16(data, idRangeOffsets + i * 2);

    if (startCode === 0xffff) {
      break;
    }

    for (let c = startCode; c <= endCode; c++) {
      if (map.size >= CMAP_MAX_ENTRIES) {
        return map;
      }
      let glyphId: number;
      if (idRangeOffset === 0) {
        glyphId = (c + idDelta) & 0xffff;
      } else {
        const glyphOffset = idRangeOffsets + i * 2 + idRangeOffset + (c - startCode) * 2;
        glyphId = readU16(data, glyphOffset);
        if (glyphId !== 0) {
          glyphId = (glyphId + idDelta) & 0xffff;
        }
      }
      if (glyphId !== 0) {
        map.set(c, glyphId);
      }
    }
  }

  return map;
}

/**
 * Parse cmap format 12 (full Unicode including supplementary planes).
 */
function parseCmapFormat12(data: Uint8Array, offset: number): Map<number, number> {
  const map = new Map<number, number>();
  const numGroups = readU32(data, offset + 12);

  // numGroups is a u32 — without bounds it can claim 4×10⁹ groups and
  // each group can claim a 4×10⁹-codepoint range. Reject crazy values
  // up front. The 4-byte multiplier 12 (group record size) plus the
  // 16-byte header puts an absolute ceiling on legitimate groups.
  const maxGroups = Math.max(0, Math.floor((data.length - offset - 16) / 12));
  if (numGroups > maxGroups) {
    return map;
  }

  for (let i = 0; i < numGroups; i++) {
    const groupOffset = offset + 16 + i * 12;
    const startCharCode = readU32(data, groupOffset);
    const endCharCode = readU32(data, groupOffset + 4);
    const startGlyphID = readU32(data, groupOffset + 8);

    if (
      endCharCode < startCharCode ||
      endCharCode > 0x10ffff /* Unicode max */ ||
      endCharCode - startCharCode > CMAP_MAX_ENTRIES
    ) {
      // Malformed or abusive range — skip rather than expand it.
      continue;
    }

    for (let c = startCharCode; c <= endCharCode; c++) {
      if (map.size >= CMAP_MAX_ENTRIES) {
        return map;
      }
      const gid = startGlyphID + (c - startCharCode);
      map.set(c, gid);
    }
  }

  return map;
}

/**
 * Get the glyph IDs for composite glyphs (compound glyphs referencing other glyphs).
 */
function getCompositeGlyphDeps(
  data: Uint8Array,
  glyfOffset: number,
  locaEntries: number[],
  glyphId: number
): number[] {
  const deps: number[] = [];
  const glyphStart = glyfOffset + locaEntries[glyphId];
  const glyphEnd = glyfOffset + locaEntries[glyphId + 1];

  if (glyphStart >= glyphEnd) {
    return deps; // empty glyph
  }

  const numContours = readI16(data, glyphStart);
  if (numContours >= 0) {
    return deps; // simple glyph
  }

  // Composite glyph
  let pos = glyphStart + 10; // skip header
  const MORE_COMPONENTS = 0x0020;
  const ARG_1_AND_2_ARE_WORDS = 0x0001;
  const WE_HAVE_A_SCALE = 0x0008;
  const WE_HAVE_AN_X_AND_Y_SCALE = 0x0040;
  const WE_HAVE_A_TWO_BY_TWO = 0x0080;

  let flags: number;
  do {
    flags = readU16(data, pos);
    const componentGid = readU16(data, pos + 2);
    deps.push(componentGid);
    pos += 4;

    // Skip arguments
    if (flags & ARG_1_AND_2_ARE_WORDS) {
      pos += 4;
    } else {
      pos += 2;
    }

    // Skip transform data
    if (flags & WE_HAVE_A_TWO_BY_TWO) {
      pos += 8;
    } else if (flags & WE_HAVE_AN_X_AND_Y_SCALE) {
      pos += 4;
    } else if (flags & WE_HAVE_A_SCALE) {
      pos += 2;
    }
  } while (flags & MORE_COMPONENTS);

  return deps;
}

function _subsetTtf(fontData: Uint8Array, characters: string): Uint8Array {
  // Check if it's a TrueType font (not CFF/OTF with PostScript outlines)
  const sfVersion = readU32(fontData, 0);
  const isTtf =
    sfVersion === 0x00010000 || // TrueType
    sfVersion === 0x74727565; // 'true'

  if (!isTtf) {
    // CFF fonts require a full CFF subsetter which is much more complex.
    // Return original for now.
    return fontData;
  }

  const { tables } = parseTableDirectory(fontData);

  const cmapTable = tables.get("cmap");
  const glyfTable = tables.get("glyf");
  const locaTable = tables.get("loca");
  const headTable = tables.get("head");
  const maxpTable = tables.get("maxp");

  if (!cmapTable || !glyfTable || !locaTable || !headTable || !maxpTable) {
    return fontData; // Missing required tables
  }

  // Determine loca format (0 = short, 1 = long)
  const indexToLocFormat = readI16(fontData, headTable.offset + 50);
  const numGlyphs = readU16(fontData, maxpTable.offset + 4);

  // Parse loca table
  const locaEntries: number[] = [];
  for (let i = 0; i <= numGlyphs; i++) {
    if (indexToLocFormat === 0) {
      locaEntries.push(readU16(fontData, locaTable.offset + i * 2) * 2);
    } else {
      locaEntries.push(readU32(fontData, locaTable.offset + i * 4));
    }
  }

  // Parse cmap to get character → glyph mapping
  const cmapOffset = cmapTable.offset;
  const numSubtables = readU16(fontData, cmapOffset + 2);
  let charToGlyph = new Map<number, number>();

  // Prefer format 12 (full Unicode), fall back to format 4 (BMP)
  for (let i = 0; i < numSubtables; i++) {
    const subOffset = cmapOffset + 4 + i * 8;
    const platformId = readU16(fontData, subOffset);
    const encodingId = readU16(fontData, subOffset + 2);
    const subtableOffset = cmapOffset + readU32(fontData, subOffset + 4);
    const format = readU16(fontData, subtableOffset);

    // Prefer platform 3 (Windows), encoding 10 (full Unicode) with format 12
    if (format === 12 && platformId === 3 && encodingId === 10) {
      charToGlyph = parseCmapFormat12(fontData, subtableOffset);
      break;
    }
    // Or platform 3, encoding 1 (BMP) with format 4
    if (format === 4 && platformId === 3 && encodingId === 1 && charToGlyph.size === 0) {
      charToGlyph = parseCmapFormat4(fontData, subtableOffset);
    }
  }

  if (charToGlyph.size === 0) {
    return fontData; // Can't parse cmap
  }

  // Collect glyph IDs we need
  const neededGlyphs = new Set<number>();
  neededGlyphs.add(0); // .notdef is always required

  for (const char of characters) {
    const codePoint = char.codePointAt(0);
    if (codePoint !== undefined) {
      const gid = charToGlyph.get(codePoint);
      if (gid !== undefined) {
        neededGlyphs.add(gid);
      }
    }
  }
  // Handle surrogate pairs
  for (let i = 0; i < characters.length; i++) {
    const cp = characters.codePointAt(i)!;
    if (cp > 0xffff) {
      i++; // skip surrogate pair low
    }
    const gid = charToGlyph.get(cp);
    if (gid !== undefined) {
      neededGlyphs.add(gid);
    }
  }

  // Resolve composite glyph dependencies
  const resolved = new Set<number>();
  const queue = [...neededGlyphs];
  while (queue.length > 0) {
    const gid = queue.pop()!;
    if (resolved.has(gid)) {
      continue;
    }
    resolved.add(gid);
    if (gid < numGlyphs) {
      const deps = getCompositeGlyphDeps(fontData, glyfTable.offset, locaEntries, gid);
      for (const dep of deps) {
        if (!resolved.has(dep)) {
          queue.push(dep);
        }
      }
    }
  }

  // If subset is most of the font, not worth subsetting
  if (resolved.size > numGlyphs * 0.7) {
    return fontData;
  }

  // Build new glyf table: keep glyph data for needed glyphs, zero out others
  const glyfData = fontData.slice(glyfTable.offset, glyfTable.offset + glyfTable.length);
  const newGlyfParts: Uint8Array[] = [];
  const newLocaEntries: number[] = [];
  let currentOffset = 0;

  for (let gid = 0; gid < numGlyphs; gid++) {
    newLocaEntries.push(currentOffset);
    if (resolved.has(gid)) {
      const start = locaEntries[gid];
      const end = locaEntries[gid + 1];
      const glyphBytes = glyfData.slice(start, end);
      newGlyfParts.push(glyphBytes);
      currentOffset += glyphBytes.length;
      // Pad to 4-byte boundary (required for long loca format)
      const padding = (4 - (glyphBytes.length % 4)) % 4;
      if (padding > 0) {
        newGlyfParts.push(new Uint8Array(padding));
        currentOffset += padding;
      }
    }
    // else: empty glyph (offset stays same)
  }
  newLocaEntries.push(currentOffset); // final sentinel

  // Build new loca table (always use long format for simplicity)
  const newLocaData = new Uint8Array((numGlyphs + 1) * 4);
  for (let i = 0; i <= numGlyphs; i++) {
    writeU32(newLocaData, i * 4, newLocaEntries[i]);
  }

  // Build new glyf table
  const newGlyfData = new Uint8Array(currentOffset);
  let writePos = 0;
  for (const part of newGlyfParts) {
    newGlyfData.set(part, writePos);
    writePos += part.length;
  }

  // Rebuild the font, replacing glyf and loca tables
  // Keep all other tables intact, including OpenType Layout tables
  // for kerning, ligatures, and glyph substitution
  const essentialTags = [
    "head",
    "hhea",
    "maxp",
    "OS/2",
    "name",
    "cmap",
    "post",
    "cvt ",
    "fpgm",
    "prep",
    "hmtx",
    "glyf",
    "loca",
    // OpenType Layout tables — preserved intact for kerning/ligatures
    "GPOS", // Glyph Positioning (kerning, mark attachment)
    "GSUB", // Glyph Substitution (ligatures, contextual alternates)
    "GDEF", // Glyph Definition (glyph classes, mark attachment)
    "kern", // Legacy kerning table
    "DSIG" // Digital Signature (preserve for validation)
  ];
  const tablesToInclude: { tag: string; data: Uint8Array }[] = [];

  for (const tag of essentialTags) {
    const tableRec = tables.get(tag);
    if (!tableRec) {
      continue;
    }
    if (tag === "glyf") {
      tablesToInclude.push({ tag, data: newGlyfData });
    } else if (tag === "loca") {
      tablesToInclude.push({ tag, data: newLocaData });
    } else if (tag === "head") {
      // Update indexToLocFormat to long (1)
      const headData = fontData.slice(tableRec.offset, tableRec.offset + tableRec.length);
      const headCopy = new Uint8Array(headData.length);
      headCopy.set(headData);
      writeU16(headCopy, 50, 1); // indexToLocFormat = 1 (long)
      tablesToInclude.push({ tag, data: headCopy });
    } else {
      tablesToInclude.push({
        tag,
        data: fontData.slice(tableRec.offset, tableRec.offset + tableRec.length)
      });
    }
  }

  // Calculate output size
  const numOutputTables = tablesToInclude.length;
  const headerSize = 12 + numOutputTables * 16;
  let totalSize = headerSize;
  for (const t of tablesToInclude) {
    totalSize += (t.data.length + 3) & ~3; // pad each table to 4-byte
  }

  const output = new Uint8Array(totalSize);
  // Write offset table header
  writeU32(output, 0, 0x00010000); // sfVersion
  writeU16(output, 4, numOutputTables);
  // searchRange, entrySelector, rangeShift
  let searchRange = 1;
  let entrySelector = 0;
  while (searchRange * 2 <= numOutputTables) {
    searchRange *= 2;
    entrySelector++;
  }
  searchRange *= 16;
  writeU16(output, 6, searchRange);
  writeU16(output, 8, entrySelector);
  writeU16(output, 10, numOutputTables * 16 - searchRange);

  // Write table records and data
  let dataOffset = headerSize;
  for (let i = 0; i < tablesToInclude.length; i++) {
    const t = tablesToInclude[i];
    const recOffset = 12 + i * 16;
    // Write tag
    for (let j = 0; j < 4; j++) {
      output[recOffset + j] = t.tag.charCodeAt(j);
    }
    // Calculate checksum
    const paddedLen = (t.data.length + 3) & ~3;
    let checksum = 0;
    for (let j = 0; j < paddedLen; j += 4) {
      const b0 = j < t.data.length ? t.data[j] : 0;
      const b1 = j + 1 < t.data.length ? t.data[j + 1] : 0;
      const b2 = j + 2 < t.data.length ? t.data[j + 2] : 0;
      const b3 = j + 3 < t.data.length ? t.data[j + 3] : 0;
      checksum = (checksum + (((b0 << 24) | (b1 << 16) | (b2 << 8) | b3) >>> 0)) >>> 0;
    }
    writeU32(output, recOffset + 4, checksum);
    writeU32(output, recOffset + 8, dataOffset);
    writeU32(output, recOffset + 12, t.data.length);
    // Write table data
    output.set(t.data, dataOffset);
    dataOffset += paddedLen;
  }

  return output;
}

// =============================================================================
// Font Embedding
// =============================================================================

/**
 * Prepare a font for embedding into a DOCX document.
 *
 * Returns both the FontDef (for the font table) and the EmbeddedFont
 * (for the embedded binary). The caller should add these to the document model.
 *
 * @param options - Font embedding options.
 * @returns The font definition and embedded font data.
 *
 * @example
 * ```ts
 * const result = embedFont({
 *   name: "CustomFont",
 *   data: fontFileBytes,
 *   style: "regular",
 *   obfuscate: true
 * });
 *
 * // Add to document
 * const doc = {
 *   ...existingDoc,
 *   fonts: [...(existingDoc.fonts ?? []), result.fontDef],
 *   embeddedFonts: [...(existingDoc.embeddedFonts ?? []), result.embeddedFont]
 * };
 * ```
 */
export function embedFont(options: EmbedFontOptions): EmbedFontResult {
  const style = options.style ?? "regular";
  const shouldObfuscate = options.obfuscate !== false;

  // Generate a font key for obfuscation
  const fontKey = generateFontKey();

  // Optionally subset the font to only include used glyphs
  let fontData = options.data;
  if (options.usedCharacters && options.usedCharacters.length > 0) {
    fontData = subsetFont(new Uint8Array(fontData), options.usedCharacters);
  }

  // Prepare font data (optionally obfuscated)
  let processedData: Uint8Array;
  let fileName: string;

  if (shouldObfuscate) {
    processedData = obfuscateFont(new Uint8Array(fontData), fontKey);
    // ODTTF file naming: {GUID}.odttf (strip braces for filename)
    const keyClean = fontKey.replace(/[{}]/g, "");
    fileName = `${keyClean}.odttf`;
  } else {
    processedData = new Uint8Array(fontData);
    // Use the font name with appropriate extension
    const ext = detectFontFormat(options.data);
    fileName = `${sanitizeFileName(options.name)}_${style}.${ext}`;
  }

  // Allocate a unique rId per call. The packager will register this rId
  // verbatim in `word/_rels/fontTable.xml.rels`, so re-using the same
  // string across calls (the previous behaviour, which keyed off the font
  // name + style) caused `Duplicate relationship ID` errors as soon as a
  // caller embedded two distinct files claiming the same name/style.
  const rId = `rIdFont${nextEmbeddedFontSequence()}`;

  // Build the FontDef with embed key info
  const fontDef: FontDef = {
    name: options.name,
    family: options.family ?? "auto",
    pitch: options.pitch ?? "variable",
    charset: options.charset !== undefined ? String(options.charset) : undefined,
    panose1: options.panose1,
    embedRegular: style === "regular" ? rId : undefined,
    embedRegularKey: style === "regular" ? fontKey : undefined,
    embedBold: style === "bold" ? rId : undefined,
    embedBoldKey: style === "bold" ? fontKey : undefined,
    embedItalic: style === "italic" ? rId : undefined,
    embedItalicKey: style === "italic" ? fontKey : undefined,
    embedBoldItalic: style === "boldItalic" ? rId : undefined,
    embedBoldItalicKey: style === "boldItalic" ? fontKey : undefined
  };

  // Build the EmbeddedFont
  const embeddedFont: EmbeddedFont = {
    rId,
    data: processedData,
    fontKey: shouldObfuscate ? fontKey : undefined,
    fileName
  };

  return { fontDef, embeddedFont };
}

/**
 * Embed multiple font variants (regular, bold, italic, boldItalic) for a font family.
 *
 * @param name - Font family name.
 * @param variants - Map of style to font data.
 * @param options - Shared options for all variants.
 * @returns Array of embed results.
 */
export function embedFontFamily(
  name: string,
  variants: Partial<Record<FontEmbedStyle, Uint8Array>>,
  options?: Omit<EmbedFontOptions, "name" | "data" | "style">
): EmbedFontResult[] {
  const results: EmbedFontResult[] = [];

  for (const [style, data] of Object.entries(variants) as [FontEmbedStyle, Uint8Array][]) {
    // `data` can legitimately be `undefined` when the caller spreads in
    // a partial map. The previous truthy check let an empty Uint8Array
    // through (which is truthy) but was correct for the undefined case.
    if (data === undefined) {
      continue;
    }
    if (data.byteLength === 0) {
      // An empty buffer is never a real font. Skip rather than embed
      // garbage that Word silently rejects.
      continue;
    }
    results.push(
      embedFont({
        name,
        data,
        style,
        ...options
      })
    );
  }

  return results;
}

/**
 * Add embedded fonts to an existing document model.
 * Merges new fonts with existing font definitions.
 *
 * @param doc - The existing document.
 * @param results - Embed results from `embedFont` or `embedFontFamily`.
 * @returns A new document with embedded fonts added.
 */
export function addEmbeddedFonts(
  doc: DocxDocument,
  results: readonly EmbedFontResult[]
): DocxDocument {
  const existingFonts = doc.fonts ? [...doc.fonts] : [];
  const existingEmbedded = doc.embeddedFonts ? [...doc.embeddedFonts] : [];

  for (const result of results) {
    // Check if font already exists and merge embed info
    const existingIdx = existingFonts.findIndex(f => f.name === result.fontDef.name);
    if (existingIdx >= 0) {
      // Merge embed keys into existing font def
      const existing = existingFonts[existingIdx]!;
      existingFonts[existingIdx] = {
        ...existing,
        embedRegular: result.fontDef.embedRegular ?? existing.embedRegular,
        embedRegularKey: result.fontDef.embedRegularKey ?? existing.embedRegularKey,
        embedBold: result.fontDef.embedBold ?? existing.embedBold,
        embedBoldKey: result.fontDef.embedBoldKey ?? existing.embedBoldKey,
        embedItalic: result.fontDef.embedItalic ?? existing.embedItalic,
        embedItalicKey: result.fontDef.embedItalicKey ?? existing.embedItalicKey,
        embedBoldItalic: result.fontDef.embedBoldItalic ?? existing.embedBoldItalic,
        embedBoldItalicKey: result.fontDef.embedBoldItalicKey ?? existing.embedBoldItalicKey
      };
    } else {
      existingFonts.push(result.fontDef);
    }
    existingEmbedded.push(result.embeddedFont);
  }

  return {
    ...doc,
    fonts: existingFonts,
    embeddedFonts: existingEmbedded
  };
}

// =============================================================================
// Internal Helpers
// =============================================================================

function detectFontFormat(data: Uint8Array): string {
  // Check magic numbers
  if (data.length >= 4) {
    // OTF: "OTTO"
    if (data[0] === 0x4f && data[1] === 0x54 && data[2] === 0x54 && data[3] === 0x4f) {
      return "otf";
    }
    // TTF: 0x00010000 or "true"
    if (
      (data[0] === 0x00 && data[1] === 0x01 && data[2] === 0x00 && data[3] === 0x00) ||
      (data[0] === 0x74 && data[1] === 0x72 && data[2] === 0x75 && data[3] === 0x65)
    ) {
      return "ttf";
    }
    // WOFF: "wOFF"
    if (data[0] === 0x77 && data[1] === 0x4f && data[2] === 0x46 && data[3] === 0x46) {
      return "woff";
    }
  }
  return "ttf"; // Default to TTF
}

function sanitizeFileName(name: string): string {
  return name.replace(/[^a-zA-Z0-9_-]/g, "_");
}

/** Module-level monotonic counter for embedded-font rIds. */
let _embeddedFontSeq = 0;
function nextEmbeddedFontSequence(): number {
  _embeddedFontSeq++;
  return _embeddedFontSeq;
}
