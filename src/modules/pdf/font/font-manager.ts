/**
 * Font manager for PDF generation.
 *
 * Manages two kinds of fonts:
 * 1. **Standard Type1 fonts** (Helvetica, Times, Courier) — always available,
 *    used as fallback for Latin text when no embedded font is provided.
 * 2. **Embedded TrueType fonts** — user-provided .ttf files for full
 *    Unicode support (CJK, Arabic, Hindi, etc.)
 *
 * When an embedded font is registered, ALL text uses the embedded font.
 * When no embedded font is provided, the system falls back to standard fonts
 * exactly as before (mapping Calibri→Helvetica, etc.)
 *
 * The manager tracks which Unicode code points are used so the font embedder
 * can create a minimal subset when writing the PDF.
 */

import { PdfDict, pdfName, pdfRef } from "../core/pdf-object";
import type { PdfWriter } from "../core/pdf-writer";
import { PdfFontError } from "../errors";
import { embedTtfFont, type EmbeddedFont } from "./font-embedder";
import {
  measureText as measureType1Text,
  getFontAscent as getType1Ascent,
  getFontDescent as getType1Descent,
  getLineHeight as getType1LineHeight
} from "./metrics";
import type { TtfFont } from "./ttf-parser";

// =============================================================================
// Font Name Mapping (Type1 fallback)
// =============================================================================

const FONT_FAMILY_MAP: Record<string, string> = {
  helvetica: "Helvetica",
  arial: "Helvetica",
  calibri: "Helvetica",
  "segoe ui": "Helvetica",
  "trebuchet ms": "Helvetica",
  verdana: "Helvetica",
  tahoma: "Helvetica",
  "gill sans": "Helvetica",
  "franklin gothic": "Helvetica",
  "lucida sans": "Helvetica",
  aptos: "Helvetica",
  "times new roman": "Times",
  times: "Times",
  georgia: "Times",
  garamond: "Times",
  "book antiqua": "Times",
  palatino: "Times",
  "palatino linotype": "Times",
  cambria: "Times",
  "century schoolbook": "Times",
  "courier new": "Courier",
  courier: "Courier",
  consolas: "Courier",
  "lucida console": "Courier",
  monaco: "Courier",
  "andale mono": "Courier",
  "cascadia code": "Courier",
  "cascadia mono": "Courier",
  menlo: "Courier"
};

function resolveBaseFont(fontFamily: string): string {
  const lower = fontFamily.toLowerCase().trim();
  return FONT_FAMILY_MAP[lower] ?? "Helvetica";
}

/**
 * Get the full PDF standard font name with style variant.
 */
export function resolvePdfFontName(fontFamily: string, bold: boolean, italic: boolean): string {
  const base = resolveBaseFont(fontFamily);

  if (base === "Helvetica") {
    if (bold && italic) {
      return "Helvetica-BoldOblique";
    }
    if (bold) {
      return "Helvetica-Bold";
    }
    if (italic) {
      return "Helvetica-Oblique";
    }
    return "Helvetica";
  }

  if (base === "Times") {
    if (bold && italic) {
      return "Times-BoldItalic";
    }
    if (bold) {
      return "Times-Bold";
    }
    if (italic) {
      return "Times-Italic";
    }
    return "Times-Roman";
  }

  if (base === "Courier") {
    if (bold && italic) {
      return "Courier-BoldOblique";
    }
    if (bold) {
      return "Courier-Bold";
    }
    if (italic) {
      return "Courier-Oblique";
    }
    return "Courier";
  }

  return "Helvetica";
}

// =============================================================================
// Font Manager
// =============================================================================

/**
 * Manages PDF font resources for a document.
 * Supports both standard Type1 fonts and embedded TrueType fonts.
 */
export class FontManager {
  // --- Standard Type1 font tracking ---
  private type1Map = new Map<string, string>(); // pdfFontName → resourceName
  private resourceToType1 = new Map<string, string>(); // resourceName → pdfFontName
  private nextType1Id = 1;

  // --- Embedded TrueType font tracking ---
  private embeddedFont: TtfFont | null = null;
  private embeddedResourceName = "";
  private usedCodePoints = new Set<number>();
  private nextEmbeddedId = 1;

  // ==========================================================================
  // Embedded Font Registration
  // ==========================================================================

  /**
   * Register an embedded TrueType font for use.
   * When set, all text rendering uses this font instead of standard fonts.
   */
  registerEmbeddedFont(font: TtfFont): string {
    this.embeddedFont = font;
    this.embeddedResourceName = `EF${this.nextEmbeddedId++}`;
    return this.embeddedResourceName;
  }

  /**
   * Check if an embedded font is available.
   */
  hasEmbeddedFont(): boolean {
    return this.embeddedFont !== null;
  }

  /**
   * Get the embedded font's resource name (if registered).
   */
  getEmbeddedResourceName(): string {
    return this.embeddedResourceName;
  }

  /**
   * Record that a text string will be rendered, tracking its code points.
   * Must be called for every text string before writing the PDF.
   */
  trackText(text: string): void {
    if (!this.embeddedFont) {
      return;
    }
    for (let i = 0; i < text.length; i++) {
      const cp = text.codePointAt(i)!;
      this.usedCodePoints.add(cp);
      if (cp > 0xffff) {
        i++; // skip low surrogate
      }
    }
  }

  // ==========================================================================
  // Standard Type1 Font Management
  // ==========================================================================

  /**
   * Ensure a standard Type1 font is registered and return its resource name.
   */
  ensureFont(pdfFontName: string): string {
    let resourceName = this.type1Map.get(pdfFontName);
    if (!resourceName) {
      resourceName = `F${this.nextType1Id++}`;
      this.type1Map.set(pdfFontName, resourceName);
      this.resourceToType1.set(resourceName, pdfFontName);
    }
    return resourceName;
  }

  /**
   * Resolve an Excel font specification to a resource name.
   * If an embedded font is registered, returns the embedded font's resource name.
   * Otherwise, falls back to standard Type1 fonts.
   */
  resolveFont(fontFamily: string, bold: boolean, italic: boolean): string {
    if (this.embeddedFont) {
      return this.embeddedResourceName;
    }
    const pdfFontName = resolvePdfFontName(fontFamily, bold, italic);
    return this.ensureFont(pdfFontName);
  }

  /**
   * Get the PDF font name for a given resource name.
   */
  getPdfFontName(resourceName: string): string {
    return this.resourceToType1.get(resourceName) ?? "Helvetica";
  }

  // ==========================================================================
  // Text Measurement
  // ==========================================================================

  /**
   * Measure text width using the correct font metrics.
   */
  measureText(text: string, resourceName: string, fontSize: number): number {
    if (this.embeddedFont && resourceName === this.embeddedResourceName) {
      return measureEmbeddedText(text, this.embeddedFont, fontSize);
    }
    const pdfFontName = this.getPdfFontName(resourceName);
    return measureType1Text(text, pdfFontName, fontSize);
  }

  /**
   * Get the font ascent in points.
   */
  getFontAscent(resourceName: string, fontSize: number): number {
    if (this.embeddedFont && resourceName === this.embeddedResourceName) {
      return (this.embeddedFont.ascent / this.embeddedFont.unitsPerEm) * fontSize;
    }
    return getType1Ascent(this.getPdfFontName(resourceName), fontSize);
  }

  /**
   * Get the font descent in points (negative value).
   */
  getFontDescent(resourceName: string, fontSize: number): number {
    if (this.embeddedFont && resourceName === this.embeddedResourceName) {
      return (this.embeddedFont.descent / this.embeddedFont.unitsPerEm) * fontSize;
    }
    return getType1Descent(this.getPdfFontName(resourceName), fontSize);
  }

  /**
   * Get the line height in points.
   */
  getLineHeight(resourceName: string, fontSize: number): number {
    if (this.embeddedFont && resourceName === this.embeddedResourceName) {
      const f = this.embeddedFont;
      return ((f.ascent - f.descent) / f.unitsPerEm) * fontSize;
    }
    return getType1LineHeight(this.getPdfFontName(resourceName), fontSize);
  }

  // ==========================================================================
  // Text Encoding
  // ==========================================================================

  /**
   * Check if a resource name refers to an embedded font.
   */
  isEmbeddedFont(resourceName: string): boolean {
    return this.embeddedFont !== null && resourceName === this.embeddedResourceName;
  }

  /**
   * Encode text for the given font resource.
   * For embedded fonts, returns a hex string `<0012003A...>`.
   * For Type1 fonts, returns null (caller should use standard string encoding).
   *
   * IMPORTANT: Must be called AFTER writeFontResources(), which builds the
   * subset and produces the unicodeToCid mapping.
   */
  encodeText(text: string, resourceName: string): string | null {
    if (!this.embeddedFont || resourceName !== this.embeddedResourceName) {
      return null;
    }

    // After writeFontResources, use the subset's CID mapping
    // (maps Unicode code points → new sequential glyph IDs in the subset font)
    if (this._embeddedResult) {
      return encodeWithCidMap(text, this._embeddedResult.unicodeToCid);
    }

    // writeFontResources not called yet — this is a programming error
    throw new PdfFontError(
      "encodeText called before writeFontResources — subset mapping not available"
    );
  }

  // ==========================================================================
  // PDF Object Writing
  // ==========================================================================

  /**
   * Write all font resource objects to the PDF.
   * Returns a map from resource name → object number.
   */
  writeFontResources(writer: PdfWriter): Map<string, number> {
    const fontObjectMap = new Map<string, number>();

    // Write standard Type1 fonts
    for (const [pdfFontName, resourceName] of this.type1Map) {
      const objNum = writer.allocObject();
      const dict = new PdfDict()
        .set("Type", "/Font")
        .set("Subtype", "/Type1")
        .set("BaseFont", pdfName(pdfFontName))
        .set("Encoding", "/WinAnsiEncoding");
      writer.addObject(objNum, dict);
      fontObjectMap.set(resourceName, objNum);
    }

    // Write embedded TrueType font
    if (this.embeddedFont && this.embeddedResourceName) {
      const embedded = embedTtfFont(
        writer,
        this.embeddedFont,
        this.usedCodePoints,
        this.embeddedResourceName
      );
      fontObjectMap.set(this.embeddedResourceName, embedded.fontObjNum);
      // Store the embedding result for text re-encoding
      this._embeddedResult = embedded;
    }

    return fontObjectMap;
  }

  /** Stored after writeFontResources is called */
  private _embeddedResult: EmbeddedFont | null = null;

  /**
   * Get the embedded font result (available after writeFontResources).
   */
  getEmbeddedResult(): EmbeddedFont | null {
    return this._embeddedResult;
  }

  /**
   * Build the Font sub-dictionary for a page's Resources dictionary.
   */
  buildFontDictString(fontObjectMap: Map<string, number>): string {
    const parts: string[] = ["<<"];
    for (const [resourceName, objNum] of fontObjectMap) {
      parts.push(`${pdfName(resourceName)} ${pdfRef(objNum)}`);
    }
    parts.push(">>");
    return parts.join("\n");
  }

  /**
   * Get all registered fonts (Type1 only, for backward compat).
   */
  getRegisteredFonts(): Array<{ resourceName: string; pdfFontName: string }> {
    const result: Array<{ resourceName: string; pdfFontName: string }> = [];
    for (const [pdfFontName, resourceName] of this.type1Map) {
      result.push({ resourceName, pdfFontName });
    }
    return result;
  }
}

// =============================================================================
// Internal Helpers
// =============================================================================

/**
 * Encode text as hex string using the font's cmap (original glyph IDs).
 * This is used during content stream generation.
 */
function encodeWithCidMap(text: string, cidMap: Map<number, number>): string {
  let hex = "<";
  for (let i = 0; i < text.length; i++) {
    const cp = text.codePointAt(i)!;
    if (cp > 0xffff) {
      i++;
    }
    const gid = cidMap.get(cp) ?? 0;
    hex += gid.toString(16).toUpperCase().padStart(4, "0");
  }
  hex += ">";
  return hex;
}

/**
 * Measure text width using the embedded font's cmap + advanceWidths.
 */
function measureEmbeddedText(text: string, font: TtfFont, fontSize: number): number {
  let totalWidth = 0;
  for (let i = 0; i < text.length; i++) {
    const cp = text.codePointAt(i)!;
    if (cp > 0xffff) {
      i++;
    }
    const gid = font.cmap.get(cp) ?? 0;
    totalWidth += font.advanceWidths[gid] ?? 0;
  }
  return (totalWidth / font.unitsPerEm) * fontSize;
}
