/**
 * Text features that need more than one glyph per code point in logical order.
 *
 * Some scripts cannot be rendered by walking code points and drawing a glyph for
 * each: Arabic letters change shape by position and join to their neighbours, Indic
 * vowel signs are stored after the consonant but drawn before it and consonant
 * clusters form conjuncts, Thai marks stack. Getting them right needs OpenType
 * shaping (GSUB/GPOS), and right-to-left text additionally needs the Unicode Bidi
 * Algorithm to be reordered.
 *
 * The point of this module is to *say* which of that a given backend does not do,
 * rather than emit text that looks plausible and is wrong — the failure mode this
 * repository treats as the worst one, because nothing reports it and the output looks
 * fine to anyone who cannot read the script.
 *
 * Both raster backends now apply the part of this that can be done without the font's
 * OpenType tables: `@utils/text-shaping` gives them contextual Arabic forms, right-to-left
 * runs in visual order and Indic vowel signs moved to the side they are drawn on. What
 * is still missing needs GSUB and GPOS themselves — conjunct ligatures and mark
 * positioning — so the warnings name those rather than claiming nothing works.
 *
 * The PDF writer is a weaker case than the rasteriser, and the difference is not a
 * shortcut. A presentation form can only be drawn if the *embedded* face has a glyph for
 * it, and across the 50 Arabic-capable system faces on one macOS host, eleven publish none
 * at all while not one covers the whole block. So the substitution is decided per cluster
 * against that face and falls back to the original letters, which is why a PDF may be
 * reordered without being joined. {@link TextShapingApplied} carries that distinction into
 * the warning, because telling a caller their letters are joined when the face could not
 * join them is as wrong as the message it replaced.
 *
 * ## Which backends this affects
 *
 * Only the ones that decide where each glyph goes:
 *
 * | Backend | Correct? | Why |
 * | --- | --- | --- |
 * | SVG | yes | The markup carries the original text; the viewer shapes it |
 * | DOCX | yes | The document carries the original text; Word shapes it |
 * | PDF | **partly** | Shaped when the embedded face publishes the forms; see below |
 * | Raster | **partly** | Shaped by `@utils/text-shaping`; no GSUB ligatures or GPOS marks |
 *
 * So a caller who needs these scripts has a correct route today, and the warning
 * should point at it.
 *
 * ## Why this is at Layer 0
 *
 * It was written for the PDF embedder and lived in `pdf/font/text-features.ts`, which
 * meant `draw` could not reach it: `pdf` is Layer 5. The rasteriser therefore drew
 * Arabic as disconnected isolated letters and Devanagari with its vowel signs on the
 * wrong side, silently — the same shape of problem as the TrueType parser and system
 * font discovery, which were also built once, in `pdf`, and unusable from `draw`.
 *
 * Script membership is tested with Unicode property escapes rather than hand-written
 * code point ranges: the engine's own Unicode tables are correct by construction and
 * stay current, whereas a hand-rolled range table is a standing source of both misses
 * and false positives.
 *
 * @module
 */

/** A script whose correct rendering requires OpenType shaping. */
interface ScriptPattern {
  readonly name: string;
  readonly pattern: RegExp;
}

const SHAPING_SCRIPTS: readonly ScriptPattern[] = [
  { name: "Arabic", pattern: /\p{Script=Arabic}/u },
  { name: "Syriac", pattern: /\p{Script=Syriac}/u },
  { name: "Thaana", pattern: /\p{Script=Thaana}/u },
  { name: "Mandaic", pattern: /\p{Script=Mandaic}/u },
  { name: "NKo", pattern: /\p{Script=Nko}/u },
  { name: "Adlam", pattern: /\p{Script=Adlam}/u },
  { name: "Devanagari", pattern: /\p{Script=Devanagari}/u },
  { name: "Bengali", pattern: /\p{Script=Bengali}/u },
  { name: "Gurmukhi", pattern: /\p{Script=Gurmukhi}/u },
  { name: "Gujarati", pattern: /\p{Script=Gujarati}/u },
  { name: "Oriya", pattern: /\p{Script=Oriya}/u },
  { name: "Tamil", pattern: /\p{Script=Tamil}/u },
  { name: "Telugu", pattern: /\p{Script=Telugu}/u },
  { name: "Kannada", pattern: /\p{Script=Kannada}/u },
  { name: "Malayalam", pattern: /\p{Script=Malayalam}/u },
  { name: "Sinhala", pattern: /\p{Script=Sinhala}/u },
  { name: "Thai", pattern: /\p{Script=Thai}/u },
  { name: "Lao", pattern: /\p{Script=Lao}/u },
  { name: "Tibetan", pattern: /\p{Script=Tibetan}/u },
  { name: "Myanmar", pattern: /\p{Script=Myanmar}/u },
  { name: "Khmer", pattern: /\p{Script=Khmer}/u },
  { name: "Javanese", pattern: /\p{Script=Javanese}/u },
  { name: "Balinese", pattern: /\p{Script=Balinese}/u },
  { name: "Tifinagh", pattern: /\p{Script=Tifinagh}/u }
];

/**
 * Right-to-left scripts, plus the explicit bidi formatting and isolate controls.
 *
 * `Bidi_Class` is not exposed to property escapes, so the RTL scripts are named
 * individually. Hebrew is here but not in {@link SHAPING_SCRIPTS} — it needs bidi
 * reordering, not shaping.
 */
const RTL_PATTERN =
  /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}\p{Script=Phoenician}\p{Script=Imperial_Aramaic}\p{Script=Kharoshthi}\p{Script=Old_Turkic}\p{Script=Avestan}\p{Script=Hanifi_Rohingya}\u200e\u200f\u202a-\u202e\u2066-\u2069]/u;

/**
 * The explicit bidi formatting and isolate controls.
 *
 * Checked separately because they are `Script=Common`, so {@link SIMPLE_TEXT_PATTERN}
 * accepts them — and it used to: `abc\u202Bdef` was classified as simple text and
 * returned before the right-to-left test ran, so a line whose direction was set
 * explicitly rather than by its letters was never reported. A caller inserting these
 * has said outright that the run needs reordering.
 */
const BIDI_CONTROL_PATTERN = /[\u200e\u200f\u202a-\u202e\u2066-\u2069]/u;

/**
 * Text made only of these needs neither shaping nor reordering. Checking this first
 * means the overwhelmingly common case — Latin, CJK, Cyrillic, Greek and shared
 * punctuation/digits — costs a single scan and skips every script test.
 */
const SIMPLE_TEXT_PATTERN =
  /^[\p{Script=Latin}\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Cyrillic}\p{Script=Greek}\p{Script=Common}\p{Script=Inherited}]*$/u;

/**
 * Whether `text` can be drawn one glyph per code point, in the order given.
 *
 * True for the overwhelming majority of text, and answered by a single scan so that
 * the common case does not pay for the script tests below.
 */
export function isSimpleText(text: string): boolean {
  if (text.length === 0) {
    return true;
  }
  if (BIDI_CONTROL_PATTERN.test(text)) {
    return false;
  }
  return SIMPLE_TEXT_PATTERN.test(text);
}

/**
 * Whether the static advance tables in `@utils/text-measure` describe `text`.
 *
 * A separate question from {@link isSimpleText}, and conflating the two produced a bug
 * in each direction. The measurer's tables cover Latin, CJK, Cyrillic, Greek and shared
 * punctuation; for a complex script they return a width unrelated to the font's real
 * advances (Tamil `வணக்கம் உலகம்` measures 66px at 24pt against roughly twice that on
 * screen). A rasteriser normalising its glyph advances to that measurement therefore
 * has to know which texts it can trust it for.
 *
 * The two predicates differ exactly on the explicit bidi controls. Those are
 * `Script=Common`, so they are well measured — the measurer charges them nothing — but
 * they do request reordering, so {@link isSimpleText} rejects them. Using that rejection
 * to pick the layout mode moved every glyph in `A\u2066B` by a rounding step; using this
 * predicate to gate shaping would skip the reordering the caller asked for.
 */
export function isWellMeasuredText(text: string): boolean {
  return SIMPLE_TEXT_PATTERN.test(text);
}

/**
 * Features of `text` that a one-glyph-per-code-point renderer cannot honour.
 *
 * Accumulate across a document with {@link TextFeatureTally} rather than calling this
 * per run: a 100k-cell sheet would otherwise produce 100k identical findings.
 */
export function textFeaturesOf(text: string): { scripts: string[]; rtl: boolean } {
  if (isSimpleText(text)) {
    return { scripts: [], rtl: false };
  }
  const scripts: string[] = [];
  for (const script of SHAPING_SCRIPTS) {
    if (script.pattern.test(text)) {
      scripts.push(script.name);
    }
  }
  return { scripts, rtl: RTL_PATTERN.test(text) };
}

/**
 * Which halves of shaping a renderer actually carried out, for {@link TextFeatureTally.warnings}.
 *
 * Both default to false, so a backend that does nothing keeps the original message without
 * having to say so.
 */
export interface TextShapingApplied {
  /**
   * Contextual presentation forms were substituted — Arabic letters are joined. Requires
   * the face to have a glyph for the form, so this can be false while `visualOrder` is true.
   */
  readonly contextualForms?: boolean;
  /** Right-to-left runs were put in visual order. Needs no glyph the face lacked. */
  readonly visualOrder?: boolean;
}

/**
 * Features seen across many runs of text, deduplicated.
 *
 * Accumulating rather than reporting per run is what keeps the warning useful: the
 * caller wants to know that a document contains Arabic, not that line 8,412 does.
 */
export class TextFeatureTally {
  private readonly scripts = new Set<string>();
  private rtl = false;

  /** Record the features present in one run of text. */
  note(text: string): void {
    if (isSimpleText(text)) {
      return;
    }
    if (!this.rtl && RTL_PATTERN.test(text)) {
      this.rtl = true;
    }
    for (const script of SHAPING_SCRIPTS) {
      // A script already recorded needs no further scanning.
      if (!this.scripts.has(script.name) && script.pattern.test(text)) {
        this.scripts.add(script.name);
      }
    }
  }

  /** Scripts needing shaping, sorted. Empty when there are none. */
  shapingScripts(): string[] {
    return [...this.scripts].sort();
  }

  /** Whether any right-to-left text was seen. */
  hasRtl(): boolean {
    return this.rtl;
  }

  /** Whether anything at all was recorded. */
  isEmpty(): boolean {
    return this.scripts.size === 0 && !this.rtl;
  }

  /**
   * Human-readable warnings, one per feature, naming a backend that gets it right.
   *
   * `renderer` names what is producing the output, so the message says which stage
   * cannot honour the text rather than blaming the library in general.
   *
   * `applies` says which halves of shaping that stage actually performed, and they are
   * separate because they cost different things. Reordering a right-to-left run needs no
   * glyph the face did not already have, so it either happens or the backend does not do it
   * at all. Substituting a contextual form needs a glyph *for that form*, which many faces
   * do not publish — so a PDF embedding a face that carries only the base letters reorders
   * the run and draws it unjoined, and must be told the second thing and not the first.
   *
   * This changes what the message claims, not merely how it is worded. Telling a caller
   * whose Arabic is now joined that it "will render incorrectly — use SVG instead" sends
   * them away from a backend that works; telling a caller whose face lacks the forms that
   * the letters are joined is the same error in the opposite direction.
   */
  warnings(renderer: string, applies: TextShapingApplied = {}): string[] {
    const messages: string[] = [];
    if (this.scripts.size > 0) {
      const scripts = this.shapingScripts().join(", ");
      messages.push(
        applies.contextualForms
          ? `Text contains ${scripts}. ${renderer} applies contextual forms and vowel ` +
              "reordering, but not the font's GSUB and GPOS tables, so conjunct ligatures " +
              "are drawn as separate glyphs and stacked marks are not positioned. Render " +
              "to SVG or DOCX for full shaping — both carry the original text and leave it " +
              "to the viewer."
          : `Text contains ${scripts}, which requires OpenType ` +
              `shaping (GSUB/GPOS). ${renderer} maps one glyph per code point, so contextual ` +
              "forms, reordering and mark positioning are not applied and the text will " +
              "render incorrectly. Render to SVG or DOCX instead — both carry the original " +
              "text and leave shaping to the viewer — or pre-shape the text yourself."
      );
    }
    if (this.rtl) {
      messages.push(
        applies.visualOrder
          ? `Text contains right-to-left characters. ${renderer} puts them in visual ` +
              "order, but does not implement the full Unicode Bidi Algorithm, so a line " +
              "mixing both directions may still be ordered wrongly at the boundaries."
          : `Text contains right-to-left characters. ${renderer} draws glyphs in logical ` +
              "order and does not run the Unicode Bidi Algorithm, so right-to-left runs will " +
              "appear reversed. Render to SVG or DOCX instead, or reorder the text yourself."
      );
    }
    return messages;
  }
}
