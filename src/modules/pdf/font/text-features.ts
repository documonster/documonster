/**
 * Detection of text features this PDF font pipeline does not implement.
 *
 * The embedder maps each code point to one glyph, in logical order, with no
 * substitution or positioning. That is correct for Latin, Greek, Cyrillic, CJK, and
 * most other scripts, but three families of text need more:
 *
 * - **Complex scripts** need OpenType shaping (GSUB/GPOS): Arabic contextual forms,
 *   Indic reordering and conjuncts, Thai mark stacking.
 * - **Bidirectional text** needs the Unicode Bidi Algorithm to reorder runs.
 * - **Color emoji** live in `COLR`/`CBDT`/`sbix`/`SVG ` tables, which a `glyf`
 *   embedder cannot render; the monochrome outline is used if present.
 *
 * Rather than emit wrong text silently, these are reported through the caller's
 * `onWarning` hook.
 *
 * The first two are not PDF's problem alone — the rasteriser in `draw` draws glyphs
 * one per code point too — so the script tables and the scanning live at Layer 0 in
 * `@utils/complex-text`, which both can reach. What stays here is the part that is
 * genuinely about embedding a font: colour glyph tables, which only a PDF embedder
 * inspects.
 *
 * @module
 */

import type { TextShapingApplied } from "@utils/complex-text";
import { TextFeatureTally } from "@utils/complex-text";

/** Tables that hold color glyph data. */
const COLOR_TABLE_TAGS = ["COLR", "CBDT", "sbix", "SVG "] as const;

/** How the shared warnings name this pipeline. */
const RENDERER = "This PDF writer";

/**
 * Text features seen across a whole document, accumulated as text is routed.
 *
 * Accumulating rather than warning per run keeps a 100k-cell sheet from producing
 * 100k identical warnings.
 */
export class TextFeatureReport {
  private readonly features = new TextFeatureTally();
  private readonly colorFontFamilies = new Set<string>();

  /** Record the features present in one text run. */
  noteText(text: string): void {
    this.features.note(text);
  }

  /** Record that an embedded face carries color glyph tables. */
  noteFontTables(familyName: string, tables: ReadonlyMap<string, unknown>): void {
    if (COLOR_TABLE_TAGS.some(tag => tables.has(tag))) {
      this.colorFontFamilies.add(familyName);
    }
  }

  /**
   * Emit one warning per detected feature.
   *
   * `applies` is a fact about this document, not about the writer. Shaping needs an embedded
   * face — the standard 14 carry no Arabic, and no Type3 fallback glyph covers U+FE70–FEFC —
   * and substituting a *form* additionally needs that face to publish the form. Both are
   * therefore measured as the document is drawn rather than assumed here.
   */
  report(warn: (message: string) => void, applies: TextShapingApplied = {}): void {
    for (const message of this.features.warnings(RENDERER, applies)) {
      warn(message);
    }
    if (this.colorFontFamilies.size > 0) {
      warn(
        `Embedded font(s) ${[...this.colorFontFamilies].sort().join(", ")} carry color ` +
          "glyph tables (COLR/CBDT/sbix/SVG), which this writer does not embed. Affected " +
          "characters render as monochrome outlines, or as .notdef when no outline exists."
      );
    }
  }
}
