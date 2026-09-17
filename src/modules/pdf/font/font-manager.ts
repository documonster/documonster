/**
 * Font manager for PDF generation.
 *
 * Manages three kinds of fonts:
 * 1. **Standard Type1 fonts** (Helvetica, Times, Courier) — always available,
 *    used for Latin text (WinAnsi repertoire) when no embedded font is provided.
 * 2. **Embedded TrueType fonts** — user-provided .ttf files for broader
 *    Unicode glyph coverage (for example CJK and Cyrillic). This does not add
 *    OpenType shaping, bidi reordering, or color-emoji rendering.
 * 3. **Type3 fallback fonts** — auto-generated vector-drawn glyphs for
 *    Unicode characters outside WinAnsi when no embedded font is provided.
 *
 * When an embedded font is registered, ALL text uses the embedded font.
 * When no embedded font is provided, the system uses Type1 for WinAnsi
 * characters and Type3 for everything else.
 *
 * The manager tracks which Unicode code points are used so the font embedder
 * and Type3 builder can create minimal subsets when writing the PDF.
 */

import { PdfDict, pdfName, pdfRef } from "@pdf/core/pdf-object";
import { hasNonWinAnsiChars, isWinAnsiCodePoint } from "@pdf/core/pdf-stream";
import type { PdfWriter } from "@pdf/core/pdf-writer";
import { PdfFontError } from "@pdf/errors";
import type { CompiledPdfFontConfig, CompiledPdfFontFace } from "@pdf/font/font-config";
import type { EmbeddedFont, EmbeddedGlyphUse } from "@pdf/font/font-embedder";
import { collectEmbeddedGlyphUses, embedTtfFont } from "@pdf/font/font-embedder";
import { buildFontPlan } from "@pdf/font/font-plan";
import {
  measureText as measureType1Text,
  getFontAscent as getType1Ascent,
  getFontDescent as getType1Descent,
  getLineHeight as getType1LineHeight
} from "@pdf/font/metrics";
import { TextFeatureReport } from "@pdf/font/text-features";
import type { TtfFont } from "@pdf/font/ttf-parser";
import type { Type3FontResult } from "@pdf/font/type3-font";
import { isType3Drawable, isType3Letterform } from "@pdf/font/type3-repertoire";
import {
  addCjkLanguageEvidence,
  concludeCjkLanguage,
  createCjkLanguageEvidence,
  isGlyphlessControl
} from "@utils/cjk";
import type { CjkLanguage } from "@utils/cjk";
import { isSimpleText } from "@utils/complex-text";
import { shapeTextForFace } from "@utils/text-shaping";
import { describeCodePointBlocks } from "@utils/unicode-blocks";

export interface RoutedFontSegment {
  readonly text: string;
  readonly resourceName: string;
  readonly codepoints: readonly number[];
  readonly width: number;
  readonly ascent: number;
  readonly descent: number;
}

export interface RoutedTextMetrics {
  readonly width: number;
  readonly ascent: number;
  readonly descent: number;
  readonly lineHeight: number;
}

interface FontRequest {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
}

interface ConfiguredFaceState {
  readonly source: CompiledPdfFontFace;
  readonly resourceName: string;
  readonly codepoints: Set<number>;
  readonly glyphUses: Map<string, EmbeddedGlyphUse>;
  embedded: EmbeddedFont | null;
}

// =============================================================================
// Font Name Mapping (Type1 fallback)
// =============================================================================

/**
 * Upper bound on cached text-routing results, to keep a large export from
 * retaining every distinct string it ever measured.
 */
const ROUTE_CACHE_LIMIT = 4096;
const ROUTE_CACHE_MAX_TEXT_LENGTH = 4096;

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
 * Supports standard Type1 fonts, embedded TrueType fonts, and auto-generated
 * Type3 fallback fonts for non-WinAnsi Unicode characters.
 */
export class FontManager {
  private readonly config: CompiledPdfFontConfig | null;
  /** Prefix isolates resource names when content is merged into an existing PDF. */
  private readonly resourcePrefix: string;
  private configuredRequests = new Map<string, FontRequest>();
  private configuredRequestKeys = new Map<string, string>();
  private configuredFaces = new Map<CompiledPdfFontFace, ConfiguredFaceState>();
  private configuredResourceFaces = new Map<string, ConfiguredFaceState>();
  private configuredRouteCache = new Map<string, readonly RoutedFontSegment[]>();
  private nextConfiguredRequestId = 1;
  private nextConfiguredFaceId = 1;
  private resourcesWritten = false;
  private missingCodePoints = new Set<number>();
  /** Set by {@link noteAutoDiscoveredFont} when a face was found by scanning. */
  private autoDiscovered: { familyName: string; codePointCount: number } | undefined;
  private readonly textFeatures = new TextFeatureReport();
  /**
   * Set while reserving renderer-generated characters (see `trackText`). Those
   * reservations must not be reported as caller-visible coverage gaps.
   */
  private suppressCoverageDiagnostics = false;

  // --- Standard Type1 font tracking ---
  private type1Map = new Map<string, string>(); // pdfFontName → resourceName
  private resourceToType1 = new Map<string, string>(); // resourceName → pdfFontName
  private nextType1Id = 1;

  // --- Embedded TrueType font tracking ---
  private embeddedFont: TtfFont | null = null;
  /**
   * The runs that shaping might change, kept verbatim until a face is known.
   *
   * Registering a glyph needs the face — a presentation form is only used if that face can
   * draw it — but the face is *discovered from* the code points tracked, so the two cannot
   * happen in one pass. The text is therefore held here during tracking and shaped in
   * {@link prepare}, once discovery has run and before the font resources are frozen.
   *
   * Only text `isSimpleText` rejects is kept, which is why this costs nothing in practice:
   * for a Latin or CJK document it stays empty, so a 100k-cell sheet retains no strings.
   */
  private readonly shapableRuns = new Set<string>();
  private embeddedResourceName = "";
  /**
   * True when `embeddedFont` is a *fallback* face rather than the document's
   * font.
   *
   * `embedFont(bytes)` states "render this document with this font", and every
   * run is routed to it. Build-time system-font auto-discovery means something
   * entirely different: "the standard-14 faces cannot draw these few code
   * points, borrow glyphs for them". Treating the second like the first
   * collapsed every run onto one regular face, so a single `→` in a document
   * silently stripped bold, italic and monospace from all of its text — and
   * mismatched the widths layout had already measured with the Type1 metrics.
   */
  private embeddedFontIsFallback = false;
  private usedCodePoints = new Set<number>();
  private usedGlyphUses = new Map<string, EmbeddedGlyphUse>();
  private nextEmbeddedId = 1;

  // --- Type3 fallback font tracking ---
  private type3CodePoints = new Set<number>();
  /** Running language tally; see {@link getTextLanguage}. */
  private readonly languageEvidence = createCjkLanguageEvidence();
  private _type3Result: Type3FontResult | null = null;

  // --- Diagnostic tracking (consumed by writers that surface warnings) ---
  /**
   * Every distinct unknown font family passed to `resolveFont` since this
   * manager was constructed. A "family" counts as unknown when it isn't
   * in `FONT_FAMILY_MAP` and isn't the canonical "helvetica"/"times"/
   * "courier" identifier. Populated as a set so a document that repeats
   * the same missing family across hundreds of text runs still produces
   * a single diagnostic.
   */
  private _unknownFontFamilies = new Set<string>();

  constructor(config?: CompiledPdfFontConfig, resourcePrefix = "") {
    if (resourcePrefix && !/^[A-Za-z][A-Za-z0-9_]*$/.test(resourcePrefix)) {
      throw new PdfFontError(`Invalid PDF font resource prefix '${resourcePrefix}'`);
    }
    this.config = config ?? null;
    this.resourcePrefix = resourcePrefix;
  }

  // ==========================================================================
  // Embedded Font Registration
  // ==========================================================================

  /**
   * Register an embedded TrueType font for use.
   * When set, all text rendering uses this font instead of standard fonts.
   */
  registerEmbeddedFont(font: TtfFont): string {
    return this.registerTtfFont(font, false);
  }

  /**
   * Register an embedded TrueType font as a *per-code-point fallback*.
   *
   * The standard-14 faces keep drawing everything they can represent, so
   * bold / italic / monospace requests still resolve to `Helvetica-Bold`,
   * `Times-Italic`, `Courier`, … and only the code points WinAnsi cannot encode
   * are routed to this face. Use it for fonts the library found on its own;
   * `registerEmbeddedFont` remains the "this is the document's font" contract.
   */
  registerFallbackFont(font: TtfFont): string {
    return this.registerTtfFont(font, true);
  }

  /**
   * Replace the fallback face with one that covers more, but only when doing so
   * cannot change a width that layout has already used.
   *
   * Discovery runs once, before layout, from the text the layout model contains. A
   * header, a footer, a watermark or a vector chart's labels report their characters
   * only afterwards, so the face chosen for the body may lack one of them — which is
   * why a second, wider face is looked for at all.
   *
   * Swapping the face outright is not safe, though, and the failure is geometric
   * rather than cosmetic. Two CJK faces agree on every ideograph, because a
   * full-width character advances by exactly one em, so the problem hides in
   * everything else: measured across `Kaiti SC` and `Heiti SC`, 345 code points
   * outside Latin-1 that both faces cover have different advances — `Ā` is 0.699 em
   * in one and 0.740 in the other. Replacing the face after those were measured
   * moves a line break by 6% of an em per occurrence.
   *
   * So the swap is allowed only when the new face reports the same advance for every
   * code point already measured through the old one. When it does not, the existing
   * face is kept: a `.notdef` box for a late character is a visible, local defect,
   * while silently re-flowing text against different metrics is neither.
   *
   * Returns `true` when the face was replaced.
   */
  widenFallbackFont(font: TtfFont): boolean {
    if (this.config) {
      throw new PdfFontError("Cannot widen a fallback face on a configured FontManager");
    }
    const incumbent = this.embeddedFont;
    if (incumbent === null || !this.embeddedFontIsFallback) {
      this.registerFallbackFont(font);
      return true;
    }
    for (const codePoint of this.usedCodePoints) {
      if (isWinAnsiCodePoint(codePoint)) {
        continue; // drawn by a standard-14 face, not by either fallback
      }
      const oldGid = incumbent.cmap.get(codePoint);
      if (oldGid === undefined || oldGid === 0) {
        continue; // the incumbent never measured it
      }
      const newGid = font.cmap.get(codePoint);
      if (newGid === undefined || newGid === 0) {
        return false; // the "wider" face is narrower here
      }
      const oldAdvance = advanceEm(incumbent, oldGid);
      const newAdvance = advanceEm(font, newGid);
      if (oldAdvance !== newAdvance) {
        return false;
      }
    }
    this.registerFallbackFont(font);
    return true;
  }

  /**
   * Note that this face came from scanning the host, not from the caller.
   *
   * Reported by {@link reportDiagnostics}, which every pipeline calls — the warning
   * used to be raised at the one discovery site inside `PdfDocumentBuilder.build()`,
   * so `Pdf.create`, `Pdf.fromExcel` and `Pdf.fromDocx` embedded a face off the host
   * and said nothing. Those are the entry points most likely to be producing a
   * golden file, and the output stops being reproducible the moment it happens.
   *
   * Separate from `registerFallbackFont` because a caller naming its own
   * `fallbackFamilies` is not making a guess and must not be warned about.
   */
  noteAutoDiscoveredFont(familyName: string, codePointCount: number): void {
    this.autoDiscovered = { familyName, codePointCount };
  }

  private registerTtfFont(font: TtfFont, fallbackOnly: boolean): string {
    if (this.config) {
      throw new PdfFontError("Cannot register a legacy embedded font on a configured FontManager");
    }
    this.embeddedFont = font;
    this.embeddedFontIsFallback = fallbackOnly;
    this.embeddedResourceName = `${this.resourcePrefix}EF${this.nextEmbeddedId++}`;
    this.textFeatures.noteFontTables(font.familyName, font.tables);
    return this.embeddedResourceName;
  }

  /**
   * Drop a build-local auto-discovered font while retaining collected text.
   *
   * Explicitly configured fonts live for the document's authoring lifetime;
   * system discovery does not. Keeping a discovered face after one build would
   * prevent a later build from choosing a broader face after new text is added.
   */
  clearAutoDiscoveredFont(): void {
    if (this.config) {
      throw new PdfFontError("Configured fonts cannot be cleared as auto-discovered fonts");
    }
    this.embeddedFont = null;
    this.embeddedFontIsFallback = false;
    this.embeddedResourceName = "";
    this._embeddedResult = null;
    this.nextEmbeddedId = 1;
  }

  /**
   * Check if an embedded font is available.
   */
  hasEmbeddedFont(): boolean {
    return this.config !== null || this.embeddedFont !== null;
  }

  /**
   * Whether this manager routes text through a configured font set rather than the
   * legacy single-face path.
   *
   * A configured manager plans every run against a named family, so its `trackText`
   * requires the resource returned by `resolveFont` and it never falls back to Type3.
   * Callers that seed a bare repertoire to load Type3 widths have to skip it.
   */
  hasConfiguredFonts(): boolean {
    return this.config !== null;
  }

  /**
   * Read-only view of the non-WinAnsi code points encountered so far when
   * no font is embedded. Used by callers (`PdfDocumentBuilder.build()`)
   * to decide whether to auto-discover a system font before the Type3
   * fallback kicks in. Returns a defensive copy so consumers cannot
   * mutate the internal set.
   */
  getType3CodePoints(): Set<number> {
    return new Set(this.type3CodePoints);
  }

  /**
   * The East Asian language the tracked text points to, or `undefined`.
   *
   * Accumulated as text is tracked rather than derived afterwards from
   * {@link getType3CodePoints}: that returns a *set*, and a set has no
   * multiplicity, so `国国国國` would look like one Simplified character against
   * one Traditional one and decide nothing. Auto-discovery needs the language to
   * pick a regional face, so the evidence is counted while the text is in hand.
   */
  getTextLanguage(): CjkLanguage | undefined {
    return concludeCjkLanguage(this.languageEvidence);
  }

  /** Code points rendered as .notdef because no configured face covered them. */
  getMissingCodePoints(): Set<number> {
    return new Set(this.missingCodePoints);
  }

  /**
   * Report font problems that degrade the output but must not fail the export.
   *
   * Shared by every entry point so the wording is identical whether the document
   * came from Excel, Word, a chart or the free-form builder. Two kinds are
   * reported: coverage gaps (a configured typeface set that misses a character
   * still produces a readable, extractable PDF, but the page shows `.notdef`,
   * and only the caller can decide which fallback family to add) and text
   * features this writer does not implement — shaping, bidi and color glyphs
   * (see `text-features.ts`).
   */
  reportDiagnostics(warn: (message: string) => void): void {
    this.textFeatures.report(warn, {
      // Recorded as the text was shaped, not inferred from the font list: whether a form was
      // substituted depends on the face publishing a glyph for *that form*, which many do
      // not. A face carrying only the base letters reorders the run and draws it unjoined,
      // and the caller has to be told the second thing without being told the first.
      contextualForms: this.substitutesForms(),
      visualOrder: this.appliesShaping()
    });
    if (this.autoDiscovered) {
      warn(
        `Auto-embedded system font '${this.autoDiscovered.familyName}' to render ` +
          `${this.autoDiscovered.codePointCount} non-WinAnsi character(s). ` +
          `Call embedFont(bytes) explicitly for deterministic output.`
      );
    }
    this.reportCoverage(warn);
    const missing = this.missingCodePoints;
    if (missing.size === 0) {
      return;
    }
    const sample = [...missing]
      .slice(0, 5)
      .map(cp => `U+${cp.toString(16).toUpperCase().padStart(4, "0")}`)
      .join(", ");
    warn(
      `${missing.size} character(s) are not covered by the configured PDF font families ` +
        `and will render with the .notdef glyph (e.g. ${sample}). ` +
        `Add a named fallback family through fallbackFamilies.`
    );
  }

  /**
   * Report non-WinAnsi characters that will render as Type3 NOTDEF boxes.
   *
   * This lives here rather than in `PdfDocumentBuilder.build()` because it is the
   * condition `PdfExportOptions.onWarning` documents — "characters that no
   * configured typeface covers … without this callback the condition is
   * invisible" — and the builder's copy was reachable only from the Word bridge.
   * The spreadsheet exporter is a separate pipeline that never ran it, so
   * `Pdf.create` and `Pdf.fromExcel` produced a PDF with tofu boxes in it and
   * said nothing at all. `missingCodePoints` below does not cover this case: it is
   * recorded while routing text through an *embedded* face's cmap, and here there
   * is no embedded face to route through.
   *
   * Called from {@link reportDiagnostics}, which every pipeline already invokes,
   * so a new entry point cannot forget it.
   *
   * The warning names Unicode blocks rather than listing hex. A list of five code
   * points is a hex dump: it says how many characters are missing but not *what*
   * they are, so it cannot tell the author which font to install. "CJK Unified
   * Ideographs (42, e.g. 中 U+4E2D)" does — the information matplotlib puts inside
   * the tofu box itself via Unicode's Last Resort font, which a warning can carry
   * without needing a font to draw it.
   */
  private reportCoverage(warn: (message: string) => void): void {
    const type3 = this.type3CodePoints;
    if (type3.size === 0) {
      return;
    }
    // A fallback face that covers everything leaves nothing to report; one that is
    // merely incomplete is reported for the part it cannot draw.
    const uncovered = this.hasFallbackFont() ? this.getUncoveredFallbackCodePoints() : type3;
    if (this.hasEmbeddedFont() && uncovered.size === 0) {
      return;
    }
    // Two different outcomes, and reporting them as one sends the reader after the
    // wrong fix. Type3 has a real glyph for every arrow, box-drawing character,
    // dingbat and enclosed numeral — that is exactly why auto-discovery is allowed to
    // pick a regional face that lacks them — and no glyph at all for an ideograph. So
    // `☐` draws correctly while `中` becomes a `.notdef` box, and the old wording
    // promised NOTDEF for both.
    const substituted = [...uncovered].filter(cp => isType3Drawable(cp));
    const tofu = [...uncovered].filter(cp => !isType3Drawable(cp));

    if (tofu.length > 0) {
      warn(
        `${tofu.length} character(s) have no glyph in any available font and will render as ` +
          `.notdef boxes: ${describeCodePointBlocks(new Set(tofu))}. ` +
          `Call embedFont(bytes) with a font that covers these blocks.`
      );
    }
    if (substituted.length > 0) {
      // "Outside the embedded font" is only true when there is one. On the path this
      // message matters most — a browser, which cannot read the host's fonts — there is
      // no embedded face at all, and telling the reader a character fell outside a font
      // they never supplied sends them looking for the wrong thing.
      const where = this.hasEmbeddedFont()
        ? "are outside the embedded font"
        : "are not covered by any font available here";
      // Letters are worth naming apart from symbols. A built-in arrow is as good as any
      // font's, so it is a note; a built-in `Δ` is legible but visibly not the typeface
      // the document asked for, which is a different thing to decide about.
      const letters = substituted.filter(cp => isType3Letterform(cp));
      warn(
        `${substituted.length} character(s) ${where} and will be drawn with built-in ` +
          `Type3 glyphs: ${describeCodePointBlocks(new Set(substituted))}. ` +
          (letters.length > 0
            ? `${letters.length} of them are letters, drawn with a plain built-in letterform ` +
              `rather than the requested typeface. `
            : "") +
          `Pass a font covering them through embedFont(bytes) to keep the whole document in ` +
          `one typeface.`
      );
    }
  }

  /**
   * Read-only view of the font families `resolveFont` saw but could not
   * map to a standard Type1 (Helvetica/Times/Courier) base. Consumers
   * use this to emit one diagnostic per distinct missing family at
   * build time rather than one per text run. The set is deduplicated
   * and preserves the exact casing the caller supplied.
   */
  getUnknownFontFamilies(): Set<string> {
    return new Set(this._unknownFontFamilies);
  }

  /**
   * Get the embedded font's resource name (if registered).
   */
  getEmbeddedResourceName(): string {
    if (this.config) {
      // Return the concrete default face directly. Going through `resolveFont`
      // would register a font *request* as a side effect, and every request is
      // materialised at write time — so an incidental call here would embed an
      // otherwise unused face with an empty subset.
      return this.ensureConfiguredFace(
        selectConfiguredRequestFace(this.config, { family: "", bold: false, italic: false })
      ).resourceName;
    }
    return this.embeddedResourceName;
  }

  /**
   * Resolve the resource name a draw-time-resolved Type1 resource should
   * actually render (and be measured) with, given the font manager's
   * *current* state. If an embedded font exists (possibly auto-discovered
   * at build time, after the text was drawn against a Type1 resource), the
   * embedded resource name is returned so both measurement and encoding go
   * through the CIDFont. Otherwise the original Type1 resource name is kept;
   * `measureText` handles Type3-fallback widths internally from that name.
   *
   * Centralises the routing rule shared by the deferred text renderer and
   * any deferred measurement (anchor alignment, word wrapping) so the two
   * never disagree.
   */
  resolveRenderResourceName(type1ResourceName: string): string {
    if (this.config) {
      return type1ResourceName;
    }
    // A fallback face never replaces the requested family: the Type1 resource
    // that carries the caller's bold / italic / family choice stays in charge
    // and `fallbackResourceFor` lends glyphs per code point.
    if (this.embeddedFont && !this.embeddedFontIsFallback) {
      return this.embeddedResourceName;
    }
    return type1ResourceName;
  }

  /**
   * The resource that must draw `codePoint` instead of the requested Type1
   * face, or null when the Type1 face handles it.
   *
   * Only ever non-null in fallback mode, for a code point WinAnsi cannot encode
   * that the fallback face actually has a glyph for. Code points the fallback
   * cannot draw either are left to the Type3 path.
   */
  fallbackResourceFor(codePoint: number): string | null {
    if (this.config || !this.embeddedFont || !this.embeddedFontIsFallback) {
      return null;
    }
    if (isWinAnsiCodePoint(codePoint)) {
      return null;
    }
    return (this.embeddedFont.cmap.get(codePoint) ?? 0) === 0 ? null : this.embeddedResourceName;
  }

  /** Whether an embedded face is lending glyphs per code point rather than replacing the document font. */
  hasFallbackFont(): boolean {
    return this.embeddedFont !== null && this.embeddedFontIsFallback;
  }

  /**
   * Non-WinAnsi code points the current fallback face cannot draw.
   *
   * Empty when there is no fallback face, because then *every* non-WinAnsi code
   * point is uncovered and {@link getType3CodePoints} already reports them.
   *
   * This exists so a caller can tell "a face is registered" from "every character
   * has a glyph", which are not the same question. A pipeline that measures a
   * document, registers the face it chose, and only later draws more text — a
   * chart's own axis labels are not part of the layout model — needs the second
   * one: the face that covered the body may not cover a character that arrives
   * afterwards, and stopping at "a font is present" left those characters as
   * `.notdef` boxes even though another installed face could draw them.
   */
  getUncoveredFallbackCodePoints(): Set<number> {
    const uncovered = new Set<number>();
    if (!this.embeddedFont || !this.embeddedFontIsFallback) {
      return uncovered;
    }
    for (const cp of this.type3CodePoints) {
      if ((this.embeddedFont.cmap.get(cp) ?? 0) === 0) {
        uncovered.add(cp);
      }
    }
    return uncovered;
  }

  /**
   * Record that a text string will be rendered, tracking its code points.
   * Must be called for every text string before writing the PDF.
   *
   * Two sets are maintained because font selection may be decided *after*
   * drawing (e.g. `PdfDocumentBuilder.build()` auto-discovers and embeds a
   * system font once it sees the accumulated non-WinAnsi code points):
   *
   *   - `usedCodePoints` — every code point seen, always. If an embedded
   *     font ends up being used (whether registered up front or
   *     auto-discovered at build time), the subset must cover all of these,
   *     including plain ASCII, so the CIDFont can encode the full run.
   *   - `type3CodePoints` — non-WinAnsi code points only. Drives the
   *     build-time decision to auto-embed a system font, and the Type3
   *     fallback when none is available.
   */
  trackText(text: string, resourceName?: string): void {
    addCjkLanguageEvidence(this.languageEvidence, text);
    if (this.config) {
      if (!resourceName) {
        throw new PdfFontError(
          "Configured font tracking requires the resource returned by resolveFont"
        );
      }
      this.routeText(text, resourceName);
      // Wrapping normalizes tabs/runs of whitespace to U+0020 when it builds
      // output lines. Reserve that renderer-generated character up front so a
      // wrapped run cannot introduce a glyph after resources are frozen.
      // The reservation is ours, not the caller's text, so it must not surface
      // as a coverage gap the caller is asked to fix.
      if (!text.includes(" ")) {
        this.suppressCoverageDiagnostics = true;
        try {
          this.routeText(" ", resourceName);
        } finally {
          this.suppressCoverageDiagnostics = false;
        }
      }
      return;
    }
    // The configured branch above reaches `noteText` through `routeText`. The
    // legacy branch never calls `routeText` during layout — it only does so at
    // encoding time, which is after diagnostics are reported — so text features
    // must be recorded here or this path would stay silent.
    this.textFeatures.noteText(text);
    collectGlyphUsesInto(this.usedGlyphUses, text);
    if (!isSimpleText(text)) {
      this.shapableRuns.add(text);
    }
    for (let i = 0; i < text.length; i++) {
      const cp = text.codePointAt(i)!;
      if (cp > 0xffff) {
        i++; // skip low surrogate
      }
      this.usedCodePoints.add(cp);
      // `usedCodePoints` keeps the controls, because the embedder registers each
      // one as part of its cluster's sequence and the subset must carry them.
      // `type3CodePoints` must not: it is the set of characters wanting a *drawn*
      // glyph, and a control has none. Counting them here spent a Type3 slot on an
      // invisible character and — worse — reported it through `onWarning` as a
      // character that "will render as a .notdef box", sending the reader after a
      // font to cover U+FE0F.
      if (!isWinAnsiCodePoint(cp) && !isGlyphlessControl(cp)) {
        this.type3CodePoints.add(cp);
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
      resourceName = `${this.resourcePrefix}F${this.nextType1Id++}`;
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
    if (this.config) {
      const family = normalizeFamilyName(fontFamily);
      const key = `${family}\u0000${bold ? 1 : 0}${italic ? 1 : 0}`;
      let resourceName = this.configuredRequestKeys.get(key);
      if (!resourceName) {
        resourceName = `${this.resourcePrefix}FR${this.nextConfiguredRequestId++}`;
        this.configuredRequestKeys.set(key, resourceName);
        this.configuredRequests.set(resourceName, { family: fontFamily, bold, italic });
      }
      return resourceName;
    }
    if (this.embeddedFont && !this.embeddedFontIsFallback) {
      return this.embeddedResourceName;
    }
    // A fallback face does not answer font *requests*: the request keeps
    // resolving to the standard-14 face that carries the caller's family,
    // bold and italic choice, and `fallbackResourceFor` lends glyphs per code
    // point at render time. The spreadsheet exporter registers its discovered
    // font before drawing, so short-circuiting here erased bold and italic
    // from every cell in a workbook containing one CJK character.
    // Record unknown families so writers can emit a single diagnostic
    // at build time instead of spamming one warning per text run.
    // The canonical base-name keys are kept in FONT_FAMILY_MAP; anything
    // not present will `?? "Helvetica"` in `resolveBaseFont` — that's the
    // trigger for a "family not recognised" diagnostic.
    const lowerKey = fontFamily.toLowerCase().trim();
    if (lowerKey && FONT_FAMILY_MAP[lowerKey] === undefined) {
      this._unknownFontFamilies.add(fontFamily);
    }
    const pdfFontName = resolvePdfFontName(fontFamily, bold, italic);
    const resourceName = this.ensureFont(pdfFontName);
    return resourceName;
  }

  /**
   * Route text to concrete font faces without splitting Unicode grapheme clusters.
   * Width and vertical metrics are normalized to em units.
   */
  routeText(text: string, resourceName: string): readonly RoutedFontSegment[] {
    if (!this.suppressCoverageDiagnostics) {
      // Every rendered run passes through here, so this is the one place that
      // sees all of the document's text. Renderer-generated reservations (the
      // wrapping space) are suppressed along with coverage diagnostics.
      this.textFeatures.noteText(text);
    }
    if (!this.config) {
      return this.routeLegacyText(text, resourceName);
    }
    const request = this.configuredRequests.get(resourceName);
    if (!request) {
      const face = this.configuredResourceFaces.get(resourceName);
      if (!face) {
        throw new PdfFontError(`Unknown configured font resource '${resourceName}'`);
      }
      const codepoints = codePointsOf(text);
      for (const codePoint of codepoints) {
        face.codepoints.add(codePoint);
      }
      // Measured and encoded from the shaped text, so the width matches the glyphs.
      const shaped = this.noteSubstitution(text, shapeConfiguredSegment(face, text));
      return [configuredSegment(shaped, codePointsOf(shaped), face)];
    }

    // Long strings carry their full text plus segment/codepoint arrays in a
    // cache entry. Skip them entirely; caching a few thousand multi-megabyte
    // cells would defeat the entry-count bound below.
    const cacheable = text.length <= ROUTE_CACHE_MAX_TEXT_LENGTH;
    const cacheKey = cacheable ? `${resourceName}\u0000${text}` : "";
    const cached = cacheable ? this.configuredRouteCache.get(cacheKey) : undefined;
    if (cached) {
      return cached;
    }

    const planner = buildFontPlan(this.config);
    planner.collect({ text, ...request });
    const plan = planner.finalize();
    const plannedFaces = new Map(plan.faces.map(face => [face.id, face]));
    const routed = Object.freeze(
      plan.segments.map(segment => {
        const plannedFace = plannedFaces.get(segment.faceId)!;
        if (plannedFace.source.kind !== "ttf" || !plannedFace.source.source) {
          throw new PdfFontError("Configured font planning produced a non-TTF face");
        }
        const state = this.ensureConfiguredFace(plannedFace.source.source);
        if (this.resourcesWritten) {
          const missingSequence = state.embedded?.encodeText(segment.text).includes(0) ?? true;
          if (missingSequence) {
            throw new PdfFontError(
              "Text introduced a Unicode sequence after PDF font resources were written. " +
                "Track every rendered text run during layout."
            );
          }
        }
        for (const codePoint of plannedFace.codepoints) {
          if (this.resourcesWritten && !state.codepoints.has(codePoint)) {
            throw new PdfFontError(
              `Text introduced U+${codePoint.toString(16).toUpperCase().padStart(4, "0")} ` +
                "after PDF font resources were written. Track every rendered text run during layout."
            );
          }
          state.codepoints.add(codePoint);
          if (
            !this.suppressCoverageDiagnostics &&
            (state.source.font.cmap.get(codePoint) ?? 0) === 0 &&
            !isGlyphlessControl(codePoint)
          ) {
            this.missingCodePoints.add(codePoint);
          }
        }
        const shapedText = this.noteSubstitution(
          segment.text,
          shapeConfiguredSegment(state, segment.text)
        );
        // A shaped segment is re-measured against the face rather than reusing the
        // planner's width, which was computed from the unshaped characters.
        const shapedSegment =
          shapedText === segment.text
            ? segment
            : configuredSegment(shapedText, codePointsOf(shapedText), state);
        return Object.freeze({
          text: shapedText,
          resourceName: state.resourceName,
          codepoints: shapedSegment.codepoints,
          width: shapedSegment.width,
          ascent: segment.ascent,
          descent: segment.descent
        });
      })
    );
    // Routing the same string repeatedly is the norm: auto row height, wrapping,
    // overflow and rendering all measure the same cell text. Cache it, but keep
    // the cache bounded — a large sheet has millions of distinct strings, and an
    // unbounded map would hold every one of them for the whole export. Repeats
    // are consecutive, so dropping the cache wholesale when it grows past the
    // cap keeps the hot path fast without tracking eviction order.
    if (cacheable) {
      if (this.configuredRouteCache.size >= ROUTE_CACHE_LIMIT) {
        this.configuredRouteCache.clear();
      }
      this.configuredRouteCache.set(cacheKey, routed);
    }
    return routed;
  }

  /** Load lazy Type3 metrics before synchronous layout starts. */
  async prepare(): Promise<void> {
    if (this.config) {
      return;
    }
    // Which code points will actually be drawn by a Type3 glyph. Without an
    // embedded face that is all of them; with one it is the part it cannot draw.
    //
    // The second case used to be skipped, on the assumption that an embedded face
    // covers everything it was chosen for. Auto-discovery no longer requires total
    // coverage — it requires the East Asian text, because a symbol it lacks is
    // drawn by Type3 rather than disqualifying an entire regional face — so an
    // incomplete fallback is now the ordinary case rather than a corner.
    //
    // Skipping it silently measured every such character at the 600 default, and
    // Type3 widths are not uniform: across U+2000–U+27FF they take thirteen
    // distinct values, and 122 of the code points a macOS Chinese face lacks are
    // not 600 — the fixed-width spaces among them (`U+2003` em space is 1000,
    // `U+200B` zero width is 0). Measuring a zero-width space as 0.6 em is a
    // visible layout error, not a rounding one.
    const needingType3 = this.embeddedFont
      ? this.getUncoveredFallbackCodePoints()
      : this.type3CodePoints;
    if (needingType3.size === 0) {
      return;
    }
    const { lookupGlyph } = await import("@pdf/font/type3-glyphs");
    this.type3PlanningWidths = new Map(
      [...needingType3].map(codePoint => [codePoint, lookupGlyph(codePoint)?.width ?? 600])
    );
  }

  /**
   * Register the glyphs shaping will ask for, while the subset can still grow.
   *
   * This is the second half of the two-pass arrangement described on {@link shapableRuns}.
   * Tracking registered the author's characters, which is what font discovery needs; the
   * glyphs actually drawn may be contextual forms, and a CID is only assigned to a
   * sequence that was registered. Skipping this step produced a page of blanks — the
   * encoder asked for a form the subset had never heard of and got CID 0.
   *
   * Called at the top of `writeFontResources`, which is the last point the subset can
   * grow and the one place every pipeline reaches. It is idempotent, so the second pass of
   * a two-pass export costs nothing.
   */
  private registerShapedGlyphs(): void {
    if (!this.embeddedFont || this.shapableRuns.size === 0) {
      return;
    }
    for (const run of this.shapableRuns) {
      for (const use of shapeForFace(this.embeddedFont, run).uses) {
        this.usedGlyphUses.set(use.sequence, use);
        this.usedCodePoints.add(use.codePoint);
      }
    }
  }

  /**
   * Whether this document's text is shaped before it is drawn.
   *
   * True once any face is embedded, whether the caller supplied it, configured it or it was
   * discovered off the host. It is deliberately not "does the face cover every form": that
   * is decided per cluster in {@link shapeForFace}, and a face covering most of a script
   * still means the run is shaped.
   */
  private appliesShaping(): boolean {
    return this.embeddedFont !== null || this.configuredResourceFaces.size > 0;
  }

  /**
   * Whether any run actually had a contextual form substituted into it.
   *
   * Set from the shaping result rather than from the presence of a face, because the two
   * differ for a real and common class of font: measured across the 50 Arabic-capable system
   * faces on one macOS host, eleven publish no presentation forms at all. For those the
   * shaper's output falls back to the original letters, nothing is joined, and claiming
   * otherwise in a warning would be a plain falsehood.
   */
  private substitutedForms = false;

  /**
   * Whether any of this document's text has a contextual form substituted into it.
   *
   * Answered by *predicting* rather than by observing, because the diagnostics are reported
   * before the content streams are serialised: `PdfDocumentBuilder` raises them just above
   * `writeFontResources`, while encoding happens later still, so a flag set during drawing
   * is always false when it is read. Shaping is a pure function of the text and the face, so
   * running it here gives the same answer the encoder will reach.
   *
   * The configured-fonts path routes during tracking and therefore has already measured it,
   * which is what {@link substitutedForms} carries.
   */
  private substitutesForms(): boolean {
    if (this.substitutedForms) {
      return true;
    }
    if (this.embeddedFont) {
      for (const run of this.shapableRuns) {
        if (this.substitutedForms) {
          break;
        }
        this.noteSubstitution(run, shapeForFace(this.embeddedFont, run).visual);
      }
    }
    return this.substitutedForms;
  }

  /** Note that shaping changed the glyphs, not merely their order. */
  private noteSubstitution(original: string, drawn: string): string {
    if (!this.substitutedForms && drawn !== original && !isReorderOnly(original, drawn)) {
      this.substitutedForms = true;
    }
    return drawn;
  }

  /** Alias for build contexts that name the pre-layout step finalize. */
  async finalize(): Promise<void> {
    await this.prepare();
  }

  /**
   * Start one writer-local font materialization session.
   *
   * Collected requests and code points belong to the document model and are
   * retained. Embedded CID maps and Type3 object references belong to one
   * PdfWriter and are discarded, so repeated builds cannot reuse artifacts
   * from a previous PDF.
   */
  beginBuild(): void {
    this.resourcesWritten = false;
    this._embeddedResult = null;
    this._type3Result = null;
    for (const state of this.configuredFaces.values()) {
      state.embedded = null;
    }
  }

  /** End the writer-local session and allow the authoring model to grow again. */
  endBuild(): void {
    this.resourcesWritten = false;
  }

  /**
   * Get the PDF font name for a given resource name.
   */
  getPdfFontName(resourceName: string): string {
    return this.resourceToType1.get(resourceName) ?? "Helvetica";
  }

  // ==========================================================================
  // Type3 Fallback Font
  // ==========================================================================

  /**
   * Check if Type3 fallback fonts are available (after writeFontResources).
   */
  hasType3Fonts(): boolean {
    return this._type3Result !== null && this._type3Result.fontObjects.size > 0;
  }

  /**
   * Resolve the Type3 font resource name and char code for a code point.
   * Returns null if the code point is not in the Type3 encoding.
   */
  resolveType3(codePoint: number): { resourceName: string; charCode: number } | null {
    if (!this._type3Result) {
      return null;
    }
    return this._type3Result.encoding.get(codePoint) ?? null;
  }

  /**
   * Check if a code point needs Type3 rendering (non-WinAnsi, no embedded font).
   */
  needsType3(codePoint: number): boolean {
    if (this.config) {
      // Configured fonts are deterministic by contract: text is drawn only by
      // the faces the caller supplied, and anything they do not cover renders
      // as .notdef (with its Unicode still recoverable through ToUnicode).
      // `writeFontResources` emits no Type3 font in this mode, so claiming a
      // code point needs one would describe a fallback that never exists.
      return false;
    }
    if (isWinAnsiCodePoint(codePoint)) {
      return false;
    }
    // A control draws nothing, so there is no glyph for Type3 to supply. The
    // renderer already strips it from a Type1 run and folds it into an embedded
    // face's cluster; saying it needs a Type3 glyph would ask for a drawing of an
    // invisible character, which resolves to `.notdef`.
    if (isGlyphlessControl(codePoint)) {
      return false;
    }
    if (!this.embeddedFont) {
      return true;
    }
    // A document font covers everything itself. A fallback face only covers the
    // code points it has glyphs for; the rest still need a Type3 glyph.
    return this.embeddedFontIsFallback && this.fallbackResourceFor(codePoint) === null;
  }

  // ==========================================================================
  // Text Measurement
  // ==========================================================================

  /**
   * Measure an unresolved font request without adding it to the document plan.
   *
   * The free-form Builder exposes measurement as a query. Calling it must not
   * embed a face, enlarge a subset, or report missing glyphs for text that was
   * never drawn. Configured fonts can be planned directly and purely; legacy
   * Type1/embedded mode retains the historical resource-based path.
   */
  measureTextRequest(
    text: string,
    family: string,
    bold: boolean,
    italic: boolean,
    fontSize: number
  ): number {
    if (!this.config) {
      return this.measureText(text, this.resolveFont(family, bold, italic), fontSize);
    }
    const planner = buildFontPlan(this.config);
    planner.collect({ text, family, bold, italic });
    return planner.finalize().segments.reduce((sum, segment) => sum + segment.width, 0) * fontSize;
  }

  /**
   * Measure text width using the correct font metrics.
   *
   * Mixed text is measured character by character so every code point is sized
   * by the face that will actually draw it: the requested Type1 face, an
   * embedded fallback face, or a Type3 glyph. Measuring with one face and
   * drawing with another is what produced the stray gaps after non-WinAnsi
   * runs.
   */
  /**
   * The text that will actually be drawn for `text`, shaped against the embedded face.
   *
   * Shaping is confined to this class on purpose. Measurement, glyph registration and
   * encoding all have to agree on the exact string, and every caller that shaped for
   * itself would be another chance for them to diverge — a divergence whose symptom is a
   * CID the subset does not contain, i.e. a blank page. It also means the callers do not
   * change at all: `Pdf`, the Excel renderer, the chart surface and the Word bridge each
   * keep passing the author's text and get shaped output.
   *
   * Only the embedded face is consulted, because it is the only one that can draw a
   * presentation form: the standard 14 fonts have no Arabic at all, and Type3 fallback
   * glyphs cover no form in U+FE70–FEFC.
   */
  private shapedForDrawing(text: string): string {
    if (!this.embeddedFont || isSimpleText(text)) {
      return text;
    }
    return this.noteSubstitution(text, shapeForFace(this.embeddedFont, text).visual);
  }

  measureText(text: string, resourceName: string, fontSize: number): number {
    if (this.config) {
      return (
        this.routeText(text, resourceName).reduce((width, segment) => width + segment.width, 0) *
        fontSize
      );
    }
    // Measure what will be drawn, not what was asked for: a shaped run is around a
    // quarter narrower than the same letters in isolated form, and a caller laying out
    // against the unshaped width would place every following run too far along. The Word
    // bridge injects this method as its measurer, which is how `run.x` and `run.width`
    // stay consistent with the glyphs without the layout engine knowing about shaping.
    text = this.shapedForDrawing(text);
    if (this.embeddedFont && resourceName === this.embeddedResourceName) {
      return measureEmbeddedText(text, this.embeddedFont, fontSize);
    }

    // Type1 only when nothing else can be involved: no Type3 glyphs planned or
    // written, no fallback face, or the text is entirely WinAnsi.
    const mayNeedFallback =
      this._type3Result !== null || this.type3PlanningWidths !== null || this.hasFallbackFont();
    if (!mayNeedFallback || !hasNonWinAnsiChars(text)) {
      const pdfFontName = this.getPdfFontName(resourceName);
      return measureType1Text(text, pdfFontName, fontSize);
    }

    // Mixed text: measure char by char
    let totalWidth = 0;
    const pdfFontName = this.getPdfFontName(resourceName);
    for (let i = 0; i < text.length; i++) {
      const cp = text.codePointAt(i)!;
      if (cp > 0xffff) {
        i++;
      }
      const char = String.fromCodePoint(cp);
      const fallback = this.fallbackResourceFor(cp);
      if (fallback !== null) {
        // Drawn by the embedded fallback face — use its advance width.
        totalWidth += measureEmbeddedText(char, this.embeddedFont!, fontSize);
        continue;
      }
      if (isWinAnsiCodePoint(cp)) {
        totalWidth += measureType1Text(String.fromCodePoint(cp), pdfFontName, fontSize);
      } else {
        // Type3 character width
        const t3 = this._type3Result?.encoding.get(cp);
        if (t3) {
          const widthMap = this._type3Result!.widths.get(t3.resourceName);
          const glyphWidth = widthMap?.get(t3.charCode) ?? 600;
          totalWidth += (glyphWidth / 1000) * fontSize;
        } else {
          // Notdef width
          const glyphWidth = this.type3PlanningWidths?.get(cp) ?? 600;
          totalWidth += (glyphWidth / 1000) * fontSize;
        }
      }
    }
    return totalWidth;
  }

  /**
   * Measure horizontal and vertical extents using every face that actually
   * draws the text. Width-only measurement is insufficient for mixed-family
   * runs: a CJK fallback face can have a taller ascent than the requested Latin
   * face and would otherwise be clipped by top/center alignment.
   */
  measureTextMetrics(text: string, resourceName: string, fontSize: number): RoutedTextMetrics {
    // A fallback face lends glyphs per code point, and `routeLegacyText` reports
    // one segment for the requested face — so its ascent and descent would be
    // the Latin face's even for text a CJK fallback actually draws. Measure
    // those extents from the faces that do the drawing.
    if (!this.config && this.hasFallbackFont() && hasNonWinAnsiChars(text)) {
      const width = this.measureText(text, resourceName, fontSize);
      let ascent = 0;
      let descent = 0;
      let sawLatin = false;
      for (const cp of codePointsOf(text)) {
        if (this.fallbackResourceFor(cp) !== null) {
          const font = this.embeddedFont!;
          ascent = Math.max(ascent, (font.ascent / font.unitsPerEm) * fontSize);
          descent = Math.min(descent, (font.descent / font.unitsPerEm) * fontSize);
        } else {
          sawLatin = true;
        }
      }
      if (sawLatin) {
        ascent = Math.max(ascent, this.getFontAscent(resourceName, fontSize));
        descent = Math.min(descent, this.getFontDescent(resourceName, fontSize));
      }
      return { width, ascent, descent, lineHeight: ascent - descent };
    }
    const segments = this.routeText(text, resourceName);
    if (segments.length === 0) {
      const ascent = this.getFontAscent(resourceName, fontSize);
      const descent = this.getFontDescent(resourceName, fontSize);
      return { width: 0, ascent, descent, lineHeight: ascent - descent };
    }
    // Accumulate unscaled and scale once, exactly as `measureText` does. Scaling
    // per segment and then summing is the same value in exact arithmetic but not in
    // floating point, and the two are used interchangeably — a caller that already
    // has the metrics must be able to read `width` instead of measuring again.
    let unscaledWidth = 0;
    let ascent = 0;
    let descent = 0;
    for (const segment of segments) {
      unscaledWidth += segment.width;
      ascent = Math.max(ascent, segment.ascent * fontSize);
      descent = Math.min(descent, segment.descent * fontSize);
    }
    return { width: unscaledWidth * fontSize, ascent, descent, lineHeight: ascent - descent };
  }

  /**
   * Get the font ascent in points.
   */
  getFontAscent(resourceName: string, fontSize: number): number {
    const configured = this.resolveConfiguredMetricFace(resourceName);
    if (configured) {
      return (configured.source.font.ascent / configured.source.font.unitsPerEm) * fontSize;
    }
    if (this.embeddedFont && resourceName === this.embeddedResourceName) {
      return (this.embeddedFont.ascent / this.embeddedFont.unitsPerEm) * fontSize;
    }
    // Type3 fonts use the same metrics as the base Type1 font
    const base = this.isType3Resource(resourceName)
      ? "Helvetica"
      : this.getPdfFontName(resourceName);
    return getType1Ascent(base, fontSize);
  }

  /**
   * Get the font descent in points (negative value).
   */
  getFontDescent(resourceName: string, fontSize: number): number {
    const configured = this.resolveConfiguredMetricFace(resourceName);
    if (configured) {
      return (configured.source.font.descent / configured.source.font.unitsPerEm) * fontSize;
    }
    if (this.embeddedFont && resourceName === this.embeddedResourceName) {
      return (this.embeddedFont.descent / this.embeddedFont.unitsPerEm) * fontSize;
    }
    const base = this.isType3Resource(resourceName)
      ? "Helvetica"
      : this.getPdfFontName(resourceName);
    return getType1Descent(base, fontSize);
  }

  /**
   * Get the line height in points.
   */
  getLineHeight(resourceName: string, fontSize: number): number {
    const configured = this.resolveConfiguredMetricFace(resourceName);
    if (configured) {
      const font = configured.source.font;
      return ((font.ascent - font.descent) / font.unitsPerEm) * fontSize;
    }
    if (this.embeddedFont && resourceName === this.embeddedResourceName) {
      const f = this.embeddedFont;
      return ((f.ascent - f.descent) / f.unitsPerEm) * fontSize;
    }
    const base = this.isType3Resource(resourceName)
      ? "Helvetica"
      : this.getPdfFontName(resourceName);
    return getType1LineHeight(base, fontSize);
  }

  // ==========================================================================
  // Text Encoding
  // ==========================================================================

  /**
   * Check if a resource name refers to an embedded font.
   */
  private isEmbeddedFont(resourceName: string): boolean {
    if (this.config) {
      return (
        this.configuredRequests.has(resourceName) || this.configuredResourceFaces.has(resourceName)
      );
    }
    return this.embeddedFont !== null && resourceName === this.embeddedResourceName;
  }

  /**
   * Check if a resource name refers to a Type3 fallback font.
   */
  private isType3Resource(resourceName: string): boolean {
    return this._type3Result?.fontObjects.has(resourceName) ?? false;
  }

  /**
   * Encode text for the given font resource.
   * For embedded fonts, returns a hex string `<0012003A...>`.
   * For Type1 fonts, returns null (caller should use standard string encoding).
   *
   * IMPORTANT: Must be called AFTER writeFontResources(), which builds the
   * subset and produces the Unicode-sequence mapping.
   */
  encodeText(text: string, resourceName: string): string | null {
    if (text.length === 0) {
      return this.config || this.isEmbeddedFont(resourceName) ? "<>" : null;
    }
    if (this.config) {
      let state: ConfiguredFaceState | undefined;
      // What gets encoded is the *shaped* text, which is not what was passed in. Routing
      // already decided it, so it is taken from the segment rather than recomputed: shaping
      // a presentation form a second time would put it through the joining tables again.
      let drawn = text;
      if (this.configuredRequests.has(resourceName)) {
        const segments = this.routeText(text, resourceName);
        if (segments.length !== 1) {
          throw new PdfFontError("Text spans multiple font faces; encode each routeText segment");
        }
        state = this.configuredResourceFaces.get(segments[0].resourceName);
        drawn = segments[0].text;
      } else {
        state = this.configuredResourceFaces.get(resourceName);
        // No routing for this resource, so this is the only place the face is known.
        if (state) {
          drawn = this.noteSubstitution(text, shapeConfiguredSegment(state, text));
        }
      }
      if (!state?.embedded) {
        throw new PdfFontError(
          "encodeText called before writeFontResources — subset mapping not available"
        );
      }
      return encodeWithEmbeddedFont(drawn, state.embedded);
    }
    if (!this.embeddedFont || resourceName !== this.embeddedResourceName) {
      return null;
    }

    // After writeFontResources, use the subset's CID mapping
    // (maps Unicode sequences → CIDs, independently from subset glyph IDs)
    if (this._embeddedResult) {
      return encodeWithEmbeddedFont(this.shapedForDrawing(text), this._embeddedResult);
    }

    // writeFontResources not called yet — this is a programming error
    throw new PdfFontError(
      "encodeText called before writeFontResources — subset mapping not available"
    );
  }

  /**
   * Encode a run of code points that all live in the same Type3 font.
   *
   * Type3 fallback fonts use single-byte encoding, so a run of consecutive
   * code points sharing one resource becomes a single multi-byte hex string
   * (`<XXYY…>`) that can be shown with one `Tj`. Emitting the whole run at
   * once keeps the text as one string in the content stream, which is what
   * PDF text extractors (and our own reader) need to recover the original
   * word instead of one fragment per glyph.
   *
   * Returns null when the Type3 fallback has not been built yet, the run is
   * empty, or any code point is missing from `resourceName`'s encoding —
   * callers then fall back to the Type1/WinAnsi path for that run.
   */
  encodeType3Run(codePoints: readonly number[], resourceName: string): string | null {
    if (!this._type3Result || codePoints.length === 0) {
      return null;
    }
    let hex = "";
    for (const cp of codePoints) {
      const entry = this._type3Result.encoding.get(cp);
      if (!entry || entry.resourceName !== resourceName) {
        return null;
      }
      hex += entry.charCode.toString(16).toUpperCase().padStart(2, "0");
    }
    return `<${hex}>`;
  }

  // ==========================================================================
  // PDF Object Writing
  // ==========================================================================

  /**
   * Write all font resource objects to the PDF.
   * Returns a map from resource name → object number.
   *
   * `async` because Type3 fallback fonts (the ~hundreds-of-KB Unicode glyph
   * tables) are loaded lazily via dynamic `import()` — only documents that
   * actually contain non-WinAnsi characters pay for them. A plain text PDF
   * never bundles the glyph tables (verified by scripts/treeshake-verify).
   */
  async writeFontResources(writer: PdfWriter): Promise<Map<string, number>> {
    // Last moment the subset can still grow, and the only one every pipeline passes
    // through. Hooking this to `prepare()` instead left the whole builder path unshaped:
    // `PdfDocumentBuilder` never calls it — only the spreadsheet exporter does — so the
    // glyphs were registered for `Pdf.fromExcel` and not for `Pdf.create`. Doing it here
    // needs no cooperation from the caller and cannot be forgotten by a new one.
    this.registerShapedGlyphs();
    const fontObjectMap = new Map<string, number>();

    if (this.config) {
      for (const [resourceName] of this.configuredRequests) {
        const request = this.configuredRequests.get(resourceName)!;
        const segments = this.routeText("", resourceName);
        if (segments.length === 0) {
          const source = selectConfiguredRequestFace(this.config, request);
          const state = this.ensureConfiguredFace(source);
          this.configuredResourceFaces.set(resourceName, state);
        }
      }
      for (const state of this.configuredFaces.values()) {
        state.embedded = embedTtfFont(
          writer,
          state.source.font,
          state.glyphUses.values(),
          state.resourceName
        );
        fontObjectMap.set(state.resourceName, state.embedded.fontObjNum);
      }
      for (const [requestName, request] of this.configuredRequests) {
        const source = selectConfiguredRequestFace(this.config, request);
        const state = this.ensureConfiguredFace(source);
        fontObjectMap.set(requestName, state.embedded!.fontObjNum);
      }
      this.resourcesWritten = true;
      return fontObjectMap;
    }

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
        this.usedGlyphUses.values(),
        this.embeddedResourceName
      );
      fontObjectMap.set(this.embeddedResourceName, embedded.fontObjNum);
      // Store the embedding result for text re-encoding
      this._embeddedResult = embedded;
    }

    // Write Type3 fallback fonts for the code points nothing else can draw.
    // A *document* font (explicit `embedFont`) covers all of them by contract,
    // so none are needed; without any embedded font every non-WinAnsi code
    // point needs one; with a fallback face only the code points it lacks a
    // glyph for do. The Type3 implementation + Unicode glyph tables are loaded
    // on demand so they stay out of bundles that never render such text.
    const type3Needed = new Set<number>();
    if (!this.embeddedFont || this.embeddedFontIsFallback) {
      for (const cp of this.type3CodePoints) {
        if (this.needsType3(cp)) {
          type3Needed.add(cp);
        }
      }
    }
    if (type3Needed.size > 0) {
      const { writeType3Fonts } = await import("@pdf/font/type3-font");
      this._type3Result = writeType3Fonts(writer, type3Needed, `${this.resourcePrefix}T3F`);
      for (const [resourceName, objNum] of this._type3Result.fontObjects) {
        fontObjectMap.set(resourceName, objNum);
      }
    }

    this.resourcesWritten = true;
    return fontObjectMap;
  }

  /** Stored after writeFontResources is called */
  private _embeddedResult: EmbeddedFont | null = null;
  private type3PlanningWidths: Map<number, number> | null = null;

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

  private ensureConfiguredFace(source: CompiledPdfFontFace): ConfiguredFaceState {
    let state = this.configuredFaces.get(source);
    if (!state) {
      state = {
        source,
        resourceName: `${this.resourcePrefix}EF${this.nextConfiguredFaceId++}`,
        codepoints: new Set<number>(),
        glyphUses: new Map<string, EmbeddedGlyphUse>(),
        embedded: null
      };
      this.configuredFaces.set(source, state);
      this.configuredResourceFaces.set(state.resourceName, state);
      this.textFeatures.noteFontTables(source.font.familyName, source.font.tables);
    }
    return state;
  }

  private resolveConfiguredMetricFace(resourceName: string): ConfiguredFaceState | null {
    if (!this.config) {
      return null;
    }
    const concrete = this.configuredResourceFaces.get(resourceName);
    if (concrete) {
      return concrete;
    }
    const request = this.configuredRequests.get(resourceName);
    return request
      ? this.ensureConfiguredFace(selectConfiguredRequestFace(this.config, request))
      : null;
  }

  private routeLegacyText(text: string, resourceName: string): readonly RoutedFontSegment[] {
    const codepoints = codePointsOf(text);
    const width = this.measureText(text, resourceName, 1);
    return Object.freeze([
      Object.freeze({
        text,
        resourceName,
        codepoints: Object.freeze(codepoints),
        width,
        ascent: this.getFontAscent(resourceName, 1),
        descent: this.getFontDescent(resourceName, 1)
      })
    ]);
  }
}

// =============================================================================
// Internal Helpers
// =============================================================================

/**
 * Encode text as hex string using the font's cmap (original glyph IDs).
 * This is used during content stream generation.
 */
function encodeWithEmbeddedFont(text: string, embedded: EmbeddedFont): string {
  let hex = "<";
  for (const cid of embedded.encodeText(text)) {
    hex += cid.toString(16).toUpperCase().padStart(4, "0");
  }
  hex += ">";
  return hex;
}

/**
 * Measure text width using the embedded font's cmap + advanceWidths.
 */
/**
 * One glyph's advance in 1/1000 em, rounded exactly as the PDF `/W` array is.
 *
 * Comparing raw font units would call two faces different when the widths actually
 * written to the file are identical, which is the only difference that can move a
 * line. Used by {@link FontManager.widenFallbackFont}.
 */
function advanceEm(font: TtfFont, gid: number): number {
  return Math.round(((font.advanceWidths[gid] ?? 0) * 1000) / font.unitsPerEm);
}

function measureEmbeddedText(text: string, font: TtfFont, fontSize: number): number {
  let totalEm = 0;
  for (let i = 0; i < text.length; i++) {
    const cp = text.codePointAt(i)!;
    if (cp > 0xffff) {
      i++;
    }
    // Semantic controls are attached to the neighbouring visible CID by the
    // embedder and have zero advance of their own.
    if (isGlyphlessControl(cp)) {
      continue;
    }
    const gid = font.cmap.get(cp) ?? 0;
    // Match the integer 1/1000-em width written to PDF /W exactly, otherwise a
    // long run can drift from the viewer's placement and cross a wrap boundary.
    totalEm += Math.round(((font.advanceWidths[gid] ?? 0) * 1000) / font.unitsPerEm) / 1000;
  }
  return totalEm * fontSize;
}

function collectGlyphUsesInto(target: Map<string, EmbeddedGlyphUse>, text: string): void {
  for (const use of collectEmbeddedGlyphUses(text)) {
    target.set(use.sequence, use);
  }
}

/**
 * Shape `text` for `state`'s face, register the glyphs it needs, and return what to draw.
 *
 * The configured-fonts path routes a run into per-face segments *before* anything is
 * shaped, so each segment is shaped against the face that will draw it. A word split
 * across two faces therefore loses the join at the seam — the same limitation the legacy
 * path has at a `splitTextRuns` boundary, and preferable to shaping against a face that
 * will not draw the glyph.
 */
/**
 * Whether `drawn` is `original` with the same characters in a different order.
 *
 * Distinguishes the two halves of shaping for the diagnostics: a right-to-left run that was
 * merely reordered has had no glyph substituted, so a warning must not claim its letters are
 * joined. Comparing sorted characters is exact for this purpose — reordering is a
 * permutation, and any substitution changes the multiset.
 */
function isReorderOnly(original: string, drawn: string): boolean {
  if (original.length !== drawn.length) {
    return false;
  }
  return [...original].sort().join("") === [...drawn].sort().join("");
}

function shapeConfiguredSegment(state: ConfiguredFaceState, text: string): string {
  if (isSimpleText(text)) {
    collectGlyphUsesInto(state.glyphUses, text);
    return text;
  }
  const { visual, uses } = shapeForFace(state.source.font, text);
  for (const use of uses) {
    state.glyphUses.set(use.sequence, use);
  }
  for (const use of uses) {
    state.codepoints.add(use.codePoint);
  }
  return visual;
}

/**
 * Shape `text` against `face` and return both what to draw and what it means.
 *
 * The single place shaping is decided, so that measurement, glyph registration and
 * encoding cannot disagree — three callers that each shaped for themselves would be three
 * chances to write a CID the subset does not contain.
 *
 * The `uses` carry the original characters in `text`, so the ToUnicode map the embedder
 * builds from them yields the letters a reader typed rather than the presentation forms
 * drawn for them.
 */
function shapeForFace(face: TtfFont, text: string): { visual: string; uses: EmbeddedGlyphUse[] } {
  const clusters = shapeTextForFace(text, codePoint => face.cmap.has(codePoint));
  let visual = "";
  const uses: EmbeddedGlyphUse[] = [];
  for (const cluster of clusters) {
    visual += cluster.visual;
    const clusterUses = collectEmbeddedGlyphUses(cluster.visual);
    // The source text is attached only when the cluster collapsed to a *single* glyph,
    // which is the case where the glyph drawn and the characters it stands for genuinely
    // differ — a presentation form, or a ligature standing for two letters.
    //
    // When the cluster fell back to its original characters it draws one glyph each, and
    // those glyphs already mean themselves. Attributing the whole cluster to each of them
    // was wrong in a way that only shows up on the clipboard: a rejected lam-alef would
    // have written "لا" twice, once per glyph.
    if (clusterUses.length === 1) {
      uses.push({ ...clusterUses[0], text: cluster.source });
    } else {
      uses.push(...clusterUses);
    }
  }
  return { visual, uses };
}

function normalizeFamilyName(name: string): string {
  return name.trim().replace(/\s+/g, " ").toLocaleLowerCase("en-US");
}

function codePointsOf(text: string): number[] {
  return Array.from(text, char => char.codePointAt(0)!);
}

function configuredSegment(
  text: string,
  codepoints: readonly number[],
  state: ConfiguredFaceState
): RoutedFontSegment {
  const font = state.source.font;
  const width = codepoints.reduce((sum, codePoint) => {
    if (isGlyphlessControl(codePoint)) {
      return sum;
    }
    const glyphId = font.cmap.get(codePoint) ?? 0;
    return sum + Math.round(((font.advanceWidths[glyphId] ?? 0) * 1000) / font.unitsPerEm) / 1000;
  }, 0);
  return Object.freeze({
    text,
    resourceName: state.resourceName,
    codepoints: Object.freeze([...codepoints]),
    width,
    ascent: font.ascent / font.unitsPerEm,
    descent: font.descent / font.unitsPerEm
  });
}

function selectConfiguredRequestFace(
  config: CompiledPdfFontConfig,
  request: FontRequest
): CompiledPdfFontFace {
  const normalized = normalizeFamilyName(request.family);
  const family = config.families.find(
    candidate =>
      candidate.normalizedName === normalized || candidate.normalizedAliases.includes(normalized)
  );
  const faces = family?.faces ?? config.default;
  return request.bold && request.italic
    ? (faces.boldItalic ?? faces.bold ?? faces.italic ?? faces.regular)
    : request.bold
      ? (faces.bold ?? faces.regular)
      : request.italic
        ? (faces.italic ?? faces.regular)
        : faces.regular;
}
