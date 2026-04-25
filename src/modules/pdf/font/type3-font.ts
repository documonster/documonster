/**
 * Type3 fallback font builder.
 *
 * Produces a PDF Type3 font object that contains vector-drawn glyphs for
 * Unicode characters that cannot be represented by standard Type1 fonts
 * (WinAnsi encoding).
 *
 * A single Type3 font supports up to 256 glyphs (single-byte encoding).
 * When more than 256 distinct non-WinAnsi characters appear in a document,
 * multiple Type3 fonts are created automatically.
 *
 * @see PDF Reference 1.7, §5.5.4 — Type 3 Fonts
 */

import { PdfDict, pdfName, pdfRef, pdfNumber, pdfArray } from "../core/pdf-object";
import { PdfContentStream } from "../core/pdf-stream";
import type { PdfWriter } from "../core/pdf-writer";
import { lookupGlyph, NOTDEF_GLYPH, type GlyphDef, type GlyphPen } from "./type3-glyphs";

// =============================================================================
// Constants
// =============================================================================

/**
 * Type3 fonts use a FontMatrix to map glyph coordinates to text space.
 * With a 1000-unit em square, FontMatrix = [0.001 0 0 0.001 0 0].
 */
const UNITS_PER_EM = 1000;

/** Maximum glyphs per Type3 font (single-byte encoding limit). */
const MAX_GLYPHS_PER_FONT = 256;

// First usable encoding slot (0x00 = .notdef, start user glyphs at 0x01).
const FIRST_SLOT = 1;

// =============================================================================
// Public API
// =============================================================================

/** Result of writing Type3 font(s) to the PDF. */
export interface Type3FontResult {
  /** Map from Type3 resource name → PDF object number. */
  fontObjects: Map<string, number>;
  /** Map from Unicode code point → { resourceName, charCode }. */
  encoding: Map<number, { resourceName: string; charCode: number }>;
  /** Advance widths: resourceName → Map<charCode, width in 1/1000 em>. */
  widths: Map<string, Map<number, number>>;
}

/**
 * Build and write Type3 fallback fonts to the PDF for the given code points.
 *
 * @param writer - The PdfWriter to add objects to.
 * @param codePoints - Set of Unicode code points that need Type3 rendering.
 * @param resourcePrefix - Prefix for resource names (e.g. "T3F").
 * @returns Encoding and object info for use by FontManager.
 */
export function writeType3Fonts(
  writer: PdfWriter,
  codePoints: Set<number>,
  resourcePrefix = "T3F"
): Type3FontResult {
  const encoding = new Map<number, { resourceName: string; charCode: number }>();
  const fontObjects = new Map<string, number>();
  const widths = new Map<string, Map<number, number>>();

  if (codePoints.size === 0) {
    return { fontObjects, encoding, widths };
  }

  // Sort code points for deterministic output
  const sorted = [...codePoints].sort((a, b) => a - b);

  // Partition into chunks of MAX_GLYPHS_PER_FONT - 1 (slot 0 is .notdef)
  const maxPerFont = MAX_GLYPHS_PER_FONT - FIRST_SLOT;
  let fontIndex = 1;

  for (let offset = 0; offset < sorted.length; offset += maxPerFont) {
    const chunk = sorted.slice(offset, offset + maxPerFont);
    const resourceName = `${resourcePrefix}${fontIndex}`;
    fontIndex++;

    const { objNum, fontWidths } = writeSingleType3Font(writer, chunk, resourceName);
    fontObjects.set(resourceName, objNum);
    widths.set(resourceName, fontWidths);

    // Record encoding for each code point
    for (let i = 0; i < chunk.length; i++) {
      encoding.set(chunk[i], { resourceName, charCode: FIRST_SLOT + i });
    }
  }

  return { fontObjects, encoding, widths };
}

// =============================================================================
// Internal — Single Type3 Font
// =============================================================================

function writeSingleType3Font(
  writer: PdfWriter,
  codePoints: number[],
  resourceName: string
): { objNum: number; fontWidths: Map<number, number> } {
  // Build CharProcs: each glyph is a content stream
  const charProcRefs: Array<{ name: string; objNum: number }> = [];
  const encodingNames: string[] = [];
  const widthValues: number[] = new Array(FIRST_SLOT + codePoints.length).fill(0);
  const fontWidths = new Map<number, number>();

  // Slot 0: .notdef
  const notdefObj = writeGlyphStream(writer, NOTDEF_GLYPH);
  charProcRefs.push({ name: ".notdef", objNum: notdefObj });
  widthValues[0] = NOTDEF_GLYPH.width;

  for (let i = 0; i < codePoints.length; i++) {
    const cp = codePoints[i];
    const glyphDef = lookupGlyph(cp) ?? NOTDEF_GLYPH;
    const glyphName = `uni${cp.toString(16).toUpperCase().padStart(4, "0")}`;

    const objNum = writeGlyphStream(writer, glyphDef);
    charProcRefs.push({ name: glyphName, objNum });
    encodingNames.push(glyphName);

    const slot = FIRST_SLOT + i;
    widthValues[slot] = glyphDef.width;
    fontWidths.set(slot, glyphDef.width);
  }

  // Build CharProcs dictionary
  const charProcsEntries: string[] = [];
  for (const { name, objNum } of charProcRefs) {
    charProcsEntries.push(`${pdfName(name)} ${pdfRef(objNum)}`);
  }
  const charProcsDict = `<<\n${charProcsEntries.join("\n")}\n>>`;

  // Build Encoding dictionary
  const differences: string[] = [];
  differences.push(`${FIRST_SLOT}`);
  for (const name of encodingNames) {
    differences.push(pdfName(name));
  }
  const encodingObjNum = writer.allocObject();
  const encodingDict = new PdfDict()
    .set("Type", "/Encoding")
    .set("Differences", `[${differences.join(" ")}]`);
  writer.addObject(encodingObjNum, encodingDict);

  // Build ToUnicode CMap
  const toUnicodeObjNum = writeToUnicodeCMap(writer, codePoints);

  // Build the Type3 font dictionary
  const fontObjNum = writer.allocObject();
  const fontBBox = `[0 0 ${UNITS_PER_EM} ${UNITS_PER_EM}]`;
  const fontMatrix = "[0.001 0 0 0.001 0 0]";

  const fontDict = new PdfDict()
    .set("Type", "/Font")
    .set("Subtype", "/Type3")
    .set("Name", pdfName(resourceName))
    .set("FontBBox", fontBBox)
    .set("FontMatrix", fontMatrix)
    .set("FirstChar", pdfNumber(0))
    .set("LastChar", pdfNumber(widthValues.length - 1))
    .set("Widths", pdfArray(widthValues.map(w => pdfNumber(w))))
    .set("CharProcs", charProcsDict)
    .set("Encoding", pdfRef(encodingObjNum))
    .set("ToUnicode", pdfRef(toUnicodeObjNum));

  writer.addObject(fontObjNum, fontDict);

  return { objNum: fontObjNum, fontWidths };
}

// =============================================================================
// Glyph Stream Writer
// =============================================================================

function writeGlyphStream(writer: PdfWriter, glyph: GlyphDef): number {
  const stream = new PdfContentStream();

  // d1 operator: wx wy llx lly urx ury — sets glyph width and bounding box
  // This tells the PDF viewer the advance width and clip region
  stream.raw(`${glyph.width} 0 0 0 ${UNITS_PER_EM} ${UNITS_PER_EM} d1`);

  // Draw the glyph using the pen adapter
  const pen = createPen(stream);
  glyph.draw(pen);

  const objNum = writer.allocObject();
  writer.addStreamObject(objNum, new PdfDict(), stream);
  return objNum;
}

/**
 * Create a GlyphPen backed by a PdfContentStream.
 */
function createPen(stream: PdfContentStream): GlyphPen {
  return {
    M: (x, y) => {
      stream.moveTo(x, y);
    },
    L: (x, y) => {
      stream.lineTo(x, y);
    },
    C: (x1, y1, x2, y2, x3, y3) => {
      stream.curveTo(x1, y1, x2, y2, x3, y3);
    },
    Z: () => {
      stream.closePath();
    },
    rect: (x, y, w, h) => {
      stream.rect(x, y, w, h);
    },
    circle: (cx, cy, r) => {
      stream.circle(cx, cy, r);
    },
    ellipse: (cx, cy, rx, ry) => {
      stream.ellipse(cx, cy, rx, ry);
    },
    stroke: () => {
      stream.stroke();
    },
    fill: () => {
      stream.fill();
    },
    fillStroke: () => {
      stream.fillAndStroke();
    },
    lineWidth: w => {
      stream.setLineWidth(w);
    }
  };
}

// =============================================================================
// ToUnicode CMap
// =============================================================================

function writeToUnicodeCMap(writer: PdfWriter, codePoints: number[]): number {
  const lines: string[] = [];
  lines.push("/CIDInit /ProcSet findresource begin");
  lines.push("12 dict begin");
  lines.push("begincmap");
  lines.push("/CIDSystemInfo");
  lines.push("<< /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def");
  lines.push("/CMapName /Adobe-Identity-UCS def");
  lines.push("/CMapType 2 def");
  lines.push("1 begincodespacerange");
  lines.push("<00> <FF>");
  lines.push("endcodespacerange");

  // Write mappings in chunks of 100 (PDF limit per beginbfchar block)
  for (let offset = 0; offset < codePoints.length; offset += 100) {
    const chunk = codePoints.slice(offset, offset + 100);
    lines.push(`${chunk.length} beginbfchar`);
    for (let i = 0; i < chunk.length; i++) {
      const slot = FIRST_SLOT + offset + i;
      const cp = chunk[i];
      const slotHex = slot.toString(16).toUpperCase().padStart(2, "0");

      if (cp > 0xffff) {
        // Supplementary character — encode as UTF-16 surrogate pair
        const hi = Math.floor((cp - 0x10000) / 0x400) + 0xd800;
        const lo = ((cp - 0x10000) % 0x400) + 0xdc00;
        const hiHex = hi.toString(16).toUpperCase().padStart(4, "0");
        const loHex = lo.toString(16).toUpperCase().padStart(4, "0");
        lines.push(`<${slotHex}> <${hiHex}${loHex}>`);
      } else {
        const cpHex = cp.toString(16).toUpperCase().padStart(4, "0");
        lines.push(`<${slotHex}> <${cpHex}>`);
      }
    }
    lines.push("endbfchar");
  }

  lines.push("endcmap");
  lines.push("CMapName currentdict /CMap defineresource pop");
  lines.push("end");
  lines.push("end");

  const cmapStr = lines.join("\n");
  const encoder = new TextEncoder();
  const data = encoder.encode(cmapStr);

  const objNum = writer.allocObject();
  writer.addStreamObject(objNum, new PdfDict(), data);
  return objNum;
}
