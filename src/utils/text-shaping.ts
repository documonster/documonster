/**
 * Text shaping for complex scripts, without external dependencies.
 *
 * Turns text in *logical* order into clusters in *visual* order, applying the two
 * transformations a renderer cannot skip for these scripts:
 *
 * - **Arabic contextual joining** — a letter's shape depends on its neighbours, and the
 *   Unicode presentation forms (U+FE70–FEFF) are a fixed mapping, so this needs no
 *   `GSUB` table.
 * - **Bidirectional reordering** — a simplified UAX #9, enough to put right-to-left runs
 *   in visual order.
 * - **Indic vowel reordering** — a left-side matra is stored after its consonant and
 *   drawn before it.
 *
 * ## Why this is at Layer 0
 *
 * It was written for the Word module and lived in `word/font/text-shaping.ts`, so the
 * two backends that actually need it could not reach it: the rasteriser is at Layer 1
 * and cannot import `word`, and the PDF writer never did. Both therefore drew Arabic as
 * disconnected isolated letters and put Indic vowel signs on the wrong side of their
 * consonant — while a working shaper sat in the same repository. That is the fifth time
 * a capability was built above the layer that needed it; see `@utils/font-ttf`,
 * `@utils/font-discovery` and `@utils/complex-text` for the others.
 *
 * ## What it does not do
 *
 * No `GSUB`/`GPOS`, so no ligatures and no mark attachment: Devanagari `क् + ष` stays two
 * glyphs instead of becoming the conjunct `क्ष`, and Thai marks are not repositioned.
 * `@utils/complex-text` reports what remains unhandled, so the limits are visible rather
 * than assumed.
 *
 * @stability experimental
 * @module
 */

// =============================================================================
// Public API
// =============================================================================

/** Script classification for a text run. */
export type ScriptType =
  | "latin"
  | "arabic"
  | "hebrew"
  | "devanagari"
  | "bengali"
  | "tamil"
  | "thai"
  | "lao"
  | "cjk"
  | "hangul"
  | "other";

/** BiDi direction for a text segment. */
export type BiDiDirection = "ltr" | "rtl";

/** A shaped glyph cluster — the output of shaping. */
export interface ShapedCluster {
  /** The original characters in this cluster. */
  readonly chars: string;
  /** The visual form (after joining/reordering). */
  readonly visual: string;
  /** Advance width multiplier (1.0 = normal, 0 = zero-width). */
  readonly advanceMultiplier: number;
  /** Script classification. */
  readonly script: ScriptType;
  /** BiDi direction. */
  readonly direction: BiDiDirection;
}

/** Shaping options. */
export interface ShapingOptions {
  /** Base paragraph direction. Default: "ltr". */
  readonly direction?: BiDiDirection;
  /** Enable Arabic joining. Default: true. */
  readonly arabicJoining?: boolean;
  /** Enable BiDi reordering. Default: true. */
  readonly bidiReorder?: boolean;
}

/**
 * Shape a text string for complex script rendering.
 * Returns an array of shaped clusters in visual order.
 *
 * @param text - The input text in logical order.
 * @param options - Shaping options.
 * @returns Array of shaped clusters.
 */
export function shapeText(text: string, options?: ShapingOptions): ShapedCluster[] {
  const dir = options?.direction ?? "ltr";
  const doArabic = options?.arabicJoining !== false;
  const doBidi = options?.bidiReorder !== false;

  // Step 1: Classify characters into script runs
  const runs = classifyRuns(text);

  // Step 2: Apply Arabic joining
  if (doArabic) {
    for (const run of runs) {
      if (run.script === "arabic") {
        applyArabicJoining(run);
      }
    }
  }

  // Step 3: Apply BiDi reordering
  let visualRuns = runs;
  if (doBidi) {
    visualRuns = applyBidiReorder(runs, dir);
  }

  // Step 4: Flatten to clusters
  const clusters: ShapedCluster[] = [];
  for (const run of visualRuns) {
    for (const cluster of run.clusters) {
      clusters.push(cluster);
    }
  }
  return clusters;
}

/** A cluster whose drawn glyph has been chosen against a particular font face. */
export interface FaceShapedCluster {
  /**
   * The characters to draw. A contextual presentation form when the face has a glyph for
   * it, otherwise the original characters — so a face that publishes no presentation
   * forms falls back to today's disconnected letters rather than to `.notdef` boxes.
   */
  readonly visual: string;
  /**
   * What this cluster means, for a ToUnicode map, copy and search. Never a presentation
   * form, and more than one character for a ligature.
   */
  readonly source: string;
}

/**
 * Shape `text`, taking each contextual form only if `hasGlyph` says the face can draw it.
 *
 * A renderer that writes glyphs and positions explicitly — PDF — cannot shape
 * unconditionally the way a viewer can. Measured across the 50 system faces on one macOS
 * host that contain Arabic, **eleven publish no Arabic presentation forms at all** (Noto
 * Nastaliq Urdu and DecoType Nastaleeq among them, which supply the forms through GSUB
 * instead) and **not one covers the whole U+FE70–FEFC block**. Substituting a presentation
 * form regardless would therefore replace today's readable-but-disconnected letters with
 * `.notdef` boxes for those faces — a worse outcome than not shaping at all.
 *
 * So the decision is per cluster, and the fallback is the *unshaped* characters. Note that
 * reordering is kept either way: putting a right-to-left run in visual order needs no
 * glyph the face did not already have, so only the form substitution is conditional.
 *
 * `source` is separate from `visual` because the two must not be conflated downstream: the
 * glyph drawn is a presentation form, while the text a reader copies out has to be the
 * original letter. A ligature makes the difference plain — one glyph, two characters.
 */
export function shapeTextForFace(
  text: string,
  hasGlyph: (codePoint: number) => boolean,
  options?: ShapingOptions
): FaceShapedCluster[] {
  const out: FaceShapedCluster[] = [];
  for (const cluster of shapeText(text, options)) {
    // A cluster consumed by a ligature contributes nothing: the ligature's own cluster
    // carries both characters in `chars`, so dropping this one loses nothing even when the
    // ligature is rejected below and the original characters are drawn instead.
    if (cluster.advanceMultiplier === 0) {
      continue;
    }
    const drawable = [...cluster.visual].every(char => hasGlyph(char.codePointAt(0) as number));
    out.push({
      visual: drawable ? cluster.visual : cluster.chars,
      source: cluster.chars
    });
  }
  return out;
}

/**
 * Detect the dominant script of a text string.
 */
export function detectScript(text: string): ScriptType {
  const counts: Partial<Record<ScriptType, number>> = {};
  for (let i = 0; i < text.length; i++) {
    const code = text.codePointAt(i)!;
    if (code > 0xffff) {
      i++;
    }
    const script = classifyChar(code);
    counts[script] = (counts[script] ?? 0) + 1;
  }

  let maxScript: ScriptType = "latin";
  let maxCount = 0;
  for (const [script, count] of Object.entries(counts) as [ScriptType, number][]) {
    if (count > maxCount && script !== "other") {
      maxScript = script;
      maxCount = count;
    }
  }
  return maxScript;
}

/**
 * Determine the BiDi direction of a text based on its first strong character.
 */
export function detectDirection(text: string): BiDiDirection {
  for (let i = 0; i < text.length; i++) {
    const code = text.codePointAt(i)!;
    if (code > 0xffff) {
      i++;
    }
    const script = classifyChar(code);
    if (script === "arabic" || script === "hebrew") {
      return "rtl";
    }
    if (script === "latin" || script === "cjk" || script === "hangul") {
      return "ltr";
    }
  }
  return "ltr";
}

// =============================================================================
// Character Classification
// =============================================================================

function classifyChar(code: number): ScriptType {
  // Arabic (0600-06FF, 0750-077F, 08A0-08FF, FB50-FDFF, FE70-FEFF)
  if (
    (code >= 0x0600 && code <= 0x06ff) ||
    (code >= 0x0750 && code <= 0x077f) ||
    (code >= 0x08a0 && code <= 0x08ff) ||
    (code >= 0xfb50 && code <= 0xfdff) ||
    (code >= 0xfe70 && code <= 0xfeff)
  ) {
    return "arabic";
  }
  // Hebrew (0590-05FF, FB1D-FB4F)
  if ((code >= 0x0590 && code <= 0x05ff) || (code >= 0xfb1d && code <= 0xfb4f)) {
    return "hebrew";
  }
  // Devanagari (0900-097F)
  if (code >= 0x0900 && code <= 0x097f) {
    return "devanagari";
  }
  // Bengali (0980-09FF)
  if (code >= 0x0980 && code <= 0x09ff) {
    return "bengali";
  }
  // Tamil (0B80-0BFF)
  if (code >= 0x0b80 && code <= 0x0bff) {
    return "tamil";
  }
  // Thai (0E00-0E7F)
  if (code >= 0x0e00 && code <= 0x0e7f) {
    return "thai";
  }
  // Lao (0E80-0EFF)
  if (code >= 0x0e80 && code <= 0x0eff) {
    return "lao";
  }
  // CJK
  if (
    (code >= 0x2e80 && code <= 0x9fff) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xfe30 && code <= 0xfe4f) ||
    (code >= 0xff00 && code <= 0xff60) ||
    (code >= 0x20000 && code <= 0x2fa1f)
  ) {
    return "cjk";
  }
  // Hangul (AC00-D7AF, 1100-11FF, 3130-318F)
  if (
    (code >= 0xac00 && code <= 0xd7af) ||
    (code >= 0x1100 && code <= 0x11ff) ||
    (code >= 0x3130 && code <= 0x318f)
  ) {
    return "hangul";
  }
  // Latin (basic + supplements + extended)
  if (
    (code >= 0x0041 && code <= 0x024f) ||
    (code >= 0x1e00 && code <= 0x1eff) ||
    (code >= 0x2c60 && code <= 0x2c7f)
  ) {
    return "latin";
  }
  return "other";
}

// =============================================================================
// Run Classification
// =============================================================================

interface ScriptRun {
  script: ScriptType;
  direction: BiDiDirection;
  clusters: ShapedCluster[];
}

function classifyRuns(text: string): ScriptRun[] {
  const runs: ScriptRun[] = [];
  let currentScript: ScriptType | null = null;
  let currentChars = "";

  for (let i = 0; i < text.length; i++) {
    const code = text.codePointAt(i)!;
    const char = String.fromCodePoint(code);
    if (code > 0xffff) {
      i++;
    }

    const script = classifyChar(code);

    // Neutral characters (spaces, punctuation) inherit surrounding script
    const isNeutral = script === "other" && (code === 0x20 || (code >= 0x21 && code <= 0x2f));

    if (currentScript === null || (!isNeutral && script !== currentScript)) {
      if (currentChars.length > 0 && currentScript !== null) {
        runs.push(makeRun(currentChars, currentScript));
      }
      currentScript = isNeutral ? (currentScript ?? "latin") : script;
      currentChars = char;
    } else {
      currentChars += char;
    }
  }

  if (currentChars.length > 0) {
    runs.push(makeRun(currentChars, currentScript ?? "latin"));
  }

  return runs;
}

function makeRun(text: string, script: ScriptType): ScriptRun {
  const direction: BiDiDirection = script === "arabic" || script === "hebrew" ? "rtl" : "ltr";
  const clusters: ShapedCluster[] = [];

  // For Devanagari, apply vowel matra reordering before creating clusters
  if (script === "devanagari") {
    const reordered = reorderDevanagariVowels(text);
    for (let i = 0; i < reordered.length; i++) {
      const code = reordered.codePointAt(i)!;
      const char = String.fromCodePoint(code);
      if (code > 0xffff) {
        i++;
      }
      clusters.push({
        chars: char,
        visual: char,
        advanceMultiplier: 1.0,
        script,
        direction
      });
    }
    return { script, direction, clusters };
  }

  // For RTL scripts, each character becomes a cluster
  // For LTR scripts, characters form individual clusters
  for (let i = 0; i < text.length; i++) {
    const code = text.codePointAt(i)!;
    const char = String.fromCodePoint(code);
    if (code > 0xffff) {
      i++;
    }
    clusters.push({
      chars: char,
      visual: char,
      advanceMultiplier: 1.0,
      script,
      direction
    });
  }

  return { script, direction, clusters };
}

// =============================================================================
// Devanagari Vowel Reordering
// =============================================================================

/** Devanagari vowel sign "i" (इ मात्रा) — visually placed before the consonant. */
const DEVANAGARI_VOWEL_I = 0x093f;

/** Devanagari consonant range: Ka (U+0915) to Ha (U+0939) + additional consonants. */
function isDevanagariConsonant(code: number): boolean {
  return (code >= 0x0915 && code <= 0x0939) || (code >= 0x0958 && code <= 0x095f);
}

/**
 * Reorder Devanagari text so that the vowel sign "i" (U+093F) is moved
 * before the consonant it modifies. In Unicode logical order, the vowel sign
 * is stored after the consonant, but visually it appears before it.
 *
 * Example: क + ि (U+0915 U+093F) → ि + क visually (U+093F U+0915 in display order)
 */
function reorderDevanagariVowels(text: string): string {
  const codePoints: number[] = [];
  for (let i = 0; i < text.length; i++) {
    const code = text.codePointAt(i)!;
    codePoints.push(code);
    if (code > 0xffff) {
      i++;
    }
  }

  const result: number[] = [];
  for (let i = 0; i < codePoints.length; i++) {
    if (codePoints[i] === DEVANAGARI_VOWEL_I && i > 0 && isDevanagariConsonant(codePoints[i - 1])) {
      // Move the vowel sign "i" before the preceding consonant.
      // Find the start of the consonant cluster (skip over any preceding halant+consonant sequences)
      let clusterStart = result.length - 1;
      // Walk back over halant+consonant pairs: ...C + halant + C + halant + C is one cluster
      while (
        clusterStart >= 2 &&
        result[clusterStart - 1] === 0x094d && // Halant (virama)
        isDevanagariConsonant(result[clusterStart - 2])
      ) {
        clusterStart -= 2;
      }
      // Insert the vowel sign before the cluster
      result.splice(clusterStart, 0, DEVANAGARI_VOWEL_I);
    } else {
      result.push(codePoints[i]);
    }
  }

  return String.fromCodePoint(...result);
}

// =============================================================================
// Arabic Joining
// =============================================================================

/** Arabic joining types: R=Right-joining, D=Dual-joining, C=Causing, U=Non-joining, T=Transparent */
type JoiningType = "R" | "D" | "C" | "U" | "T";

function getArabicJoiningType(code: number): JoiningType {
  // Simplified joining type classification based on Unicode data
  // Full data would come from UCD ArabicJoining.txt

  // Transparent marks (0610-061A, 064B-065F, 0670, 06D6-06DC, etc.)
  if (
    (code >= 0x0610 && code <= 0x061a) ||
    (code >= 0x064b && code <= 0x065f) ||
    code === 0x0670 ||
    (code >= 0x06d6 && code <= 0x06dc) ||
    (code >= 0x06df && code <= 0x06e4) ||
    (code >= 0x06e7 && code <= 0x06e8) ||
    (code >= 0x06ea && code <= 0x06ed)
  ) {
    return "T";
  }

  // Non-joining (e.g., numbers, some special chars)
  if (
    (code >= 0x0600 && code <= 0x0605) ||
    code === 0x0608 ||
    code === 0x060b ||
    code === 0x0621 // Hamza (isolated)
  ) {
    return "U";
  }

  // Right-joining: Alef variants, Dal, Thal, Ra, Zain, Waw, etc.
  if (
    code === 0x0622 || // Alef with Madda
    code === 0x0623 || // Alef with Hamza Above
    code === 0x0624 || // Waw with Hamza
    code === 0x0625 || // Alef with Hamza Below
    code === 0x0627 || // Alef
    code === 0x0629 || // Teh Marbuta
    code === 0x062f || // Dal
    code === 0x0630 || // Thal
    code === 0x0631 || // Ra
    code === 0x0632 || // Zain
    code === 0x0648 || // Waw
    (code >= 0x0671 && code <= 0x0673) ||
    (code >= 0x0675 && code <= 0x0677) ||
    code === 0x0688 || // Dal with dot
    code === 0x0689 ||
    code === 0x068a ||
    code === 0x068b ||
    code === 0x068c ||
    code === 0x068d ||
    code === 0x068e ||
    code === 0x0691 || // Ra with dot
    code === 0x0692 ||
    code === 0x0693 ||
    code === 0x0694 ||
    code === 0x0695 ||
    code === 0x0696 ||
    code === 0x0697 ||
    code === 0x0698 || // Jeh
    code === 0x0699 ||
    code === 0x06c0 ||
    code === 0x06c3 || // Teh Marbuta Goal
    code === 0x06c5 ||
    code === 0x06c6 ||
    code === 0x06c7 ||
    code === 0x06c8 ||
    code === 0x06c9 ||
    code === 0x06cb ||
    code === 0x06cd ||
    code === 0x06cf ||
    code === 0x06d2 ||
    code === 0x06d3 ||
    code === 0x06d5 ||
    code === 0x06ee ||
    code === 0x06ef
  ) {
    return "R";
  }

  // Most Arabic letters in 0626-0649 range are Dual-joining
  if (code >= 0x0626 && code <= 0x064a) {
    return "D";
  }

  // Dual-joining (default for most Arabic range)
  if (code >= 0x066e && code <= 0x06d1) {
    return "D";
  }

  // Join-causing: ZWJ
  if (code === 0x200d) {
    return "C";
  }

  return "U";
}

/**
 * Apply Arabic contextual joining to a run of Arabic text.
 * Determines if each character should use its Initial, Medial, Final, or Isolated form.
 * Also handles mandatory Lam-Alef ligatures.
 */
function applyArabicJoining(run: ScriptRun): void {
  const clusters = run.clusters;
  const len = clusters.length;

  for (let i = 0; i < len; i++) {
    // Skip zero-width clusters (consumed by ligature)
    if (clusters[i].advanceMultiplier === 0) {
      continue;
    }

    const code = clusters[i].chars.codePointAt(0)!;
    const type = getArabicJoiningType(code);

    if (type === "T" || type === "U") {
      // Transparent marks and non-joining don't change form
      continue;
    }

    // Find previous and next joining characters (skip transparent and zero-width).
    //
    // The two directions ask different questions of the neighbour, and conflating them
    // is what made this wrong. A *preceding* character continues the join if it can
    // attach on its left — D or C. A *following* character continues it if it can attach
    // on its right — D, C or **R**.
    //
    // There is no `L` case: left-joining exists in Unicode (Phags-pa) but not in Arabic,
    // and `JoiningType` here does not model it. TypeScript rejects the test outright,
    // which is the right answer — an unreachable branch would only suggest otherwise.
    //
    // `R` was missing from the second test, and R is not a rare case: it covers alef,
    // dal, reh, waw and their variants, which appear in almost every Arabic word. In
    // `مرحبا` both ر and ا are R, so meem came out isolated instead of initial and beh
    // came out final instead of medial — the word was drawn as disconnected letters
    // even though the joining pass had run.
    let prevJoins = false;
    for (let j = i - 1; j >= 0; j--) {
      if (clusters[j].advanceMultiplier === 0) {
        continue;
      }
      const prevCode = clusters[j].chars.codePointAt(0)!;
      const prevType = getArabicJoiningType(prevCode);
      if (prevType === "T") {
        continue;
      }
      prevJoins = prevType === "D" || prevType === "C";
      break;
    }

    let nextJoins = false;
    for (let j = i + 1; j < len; j++) {
      if (clusters[j].advanceMultiplier === 0) {
        continue;
      }
      const nextCode = clusters[j].chars.codePointAt(0)!;
      const nextType = getArabicJoiningType(nextCode);
      if (nextType === "T") {
        continue;
      }
      nextJoins = nextType === "D" || nextType === "C" || nextType === "R";
      break;
    }

    // Determine contextual form
    let form: "isolated" | "initial" | "medial" | "final";
    if (type === "D") {
      if (prevJoins && nextJoins) {
        form = "medial";
      } else if (prevJoins) {
        form = "final";
      } else if (nextJoins) {
        form = "initial";
      } else {
        form = "isolated";
      }
    } else if (type === "R") {
      if (prevJoins) {
        form = "final";
      } else {
        form = "isolated";
      }
    } else {
      form = "isolated";
    }

    // Get the presentation form character
    const presentationForm = getArabicPresentationForm(code, form);
    if (presentationForm !== code) {
      (clusters[i] as { visual: string }).visual = String.fromCodePoint(presentationForm);
    }
  }

  // Lam-Alef ligatures are applied *after* the contextual forms, and the order is the
  // whole point. Run first, the ligature wrote U+FEFB into the lam's cluster and the loop
  // above then overwrote it: the cluster's `chars` had become "لا", `codePointAt(0)` still
  // answered lam, and the form calculation replaced the ligature with lam's own initial
  // form U+FEDD. Since the alef cluster had already been marked zero-width and dropped,
  // **the alef vanished from the output entirely** — a lost letter, not a wrong shape, and
  // nothing covered it.
  applyLamAlefLigatures(clusters);
}

// Lam-Alef ligature mappings: Alef variant → [isolated ligature, final ligature]
// prettier-ignore
const LAM_ALEF_LIGATURES: Record<number, [number, number]> = {
  0x0627: [0xfefb, 0xfefc],  // Lam + Alef
  0x0622: [0xfef5, 0xfef6],  // Lam + Alef with Madda Above
  0x0623: [0xfef7, 0xfef8],  // Lam + Alef with Hamza Above
  0x0625: [0xfef9, 0xfefa],  // Lam + Alef with Hamza Below
};

/**
 * Detect Lam+Alef sequences and replace them with ligature forms.
 * The Lam cluster gets the ligature glyph, and the Alef cluster is marked as zero-width.
 */
function applyLamAlefLigatures(clusters: ShapedCluster[]): void {
  const len = clusters.length;
  for (let i = 0; i < len - 1; i++) {
    const code = clusters[i].chars.codePointAt(0)!;
    if (code !== 0x0644) {
      // Not Lam
      continue;
    }

    // Find the next non-transparent cluster
    let nextIdx = -1;
    for (let j = i + 1; j < len; j++) {
      const nextCode = clusters[j].chars.codePointAt(0)!;
      const nextType = getArabicJoiningType(nextCode);
      if (nextType === "T") {
        continue;
      }
      nextIdx = j;
      break;
    }

    if (nextIdx === -1) {
      continue;
    }

    const nextCode = clusters[nextIdx].chars.codePointAt(0)!;
    const ligature = LAM_ALEF_LIGATURES[nextCode];
    if (!ligature) {
      continue;
    }

    // Determine if Lam is preceded by a joining character (→ use final ligature form)
    let prevJoins = false;
    for (let j = i - 1; j >= 0; j--) {
      const prevCode = clusters[j].chars.codePointAt(0)!;
      const prevType = getArabicJoiningType(prevCode);
      if (prevType === "T") {
        continue;
      }
      prevJoins = prevType === "D" || prevType === "C";
      break;
    }

    const ligatureForm = prevJoins ? ligature[1] : ligature[0];

    // Replace Lam cluster with the ligature
    (clusters[i] as { chars: string }).chars = clusters[i].chars + clusters[nextIdx].chars;
    (clusters[i] as { visual: string }).visual = String.fromCodePoint(ligatureForm);

    // Mark Alef cluster as consumed (zero-width)
    (clusters[nextIdx] as { visual: string }).visual = "";
    (clusters[nextIdx] as { advanceMultiplier: number }).advanceMultiplier = 0;
  }
}

/**
 * Get the Arabic presentation form for a character and joining context.
 * Uses Unicode Arabic Presentation Forms-B (FE70-FEFF) mappings.
 */
function getArabicPresentationForm(
  code: number,
  form: "isolated" | "initial" | "medial" | "final"
): number {
  // Mapping table for common Arabic letters → presentation forms
  // Each entry: [isolated, final, initial, medial]
  const forms = ARABIC_FORMS[code];
  if (!forms) {
    return code;
  }

  switch (form) {
    case "isolated":
      return forms[0];
    case "final":
      return forms[1] ?? code;
    case "initial":
      return forms[2] ?? code;
    case "medial":
      return forms[3] ?? code;
    default:
      return code;
  }
}

// Arabic Presentation Forms-B mappings (code → [isolated, final, initial, medial])
// prettier-ignore
const ARABIC_FORMS: Record<number, [number, number, number?, number?]> = {
  // --- Basic Arabic letters (U+0621-U+064A) ---
  0x0627: [0xfe8d, 0xfe8e],                         // Alef
  0x0628: [0xfe8f, 0xfe90, 0xfe91, 0xfe92],         // Ba
  0x062a: [0xfe95, 0xfe96, 0xfe97, 0xfe98],         // Ta
  0x062b: [0xfe99, 0xfe9a, 0xfe9b, 0xfe9c],         // Tha
  0x062c: [0xfe9d, 0xfe9e, 0xfe9f, 0xfea0],         // Jeem
  0x062d: [0xfea1, 0xfea2, 0xfea3, 0xfea4],         // Ha
  0x062e: [0xfea5, 0xfea6, 0xfea7, 0xfea8],         // Kha
  0x062f: [0xfea9, 0xfeaa],                          // Dal
  0x0630: [0xfeab, 0xfeac],                          // Thal
  0x0631: [0xfead, 0xfeae],                          // Ra
  0x0632: [0xfeaf, 0xfeb0],                          // Zain
  0x0633: [0xfeb1, 0xfeb2, 0xfeb3, 0xfeb4],         // Seen
  0x0634: [0xfeb5, 0xfeb6, 0xfeb7, 0xfeb8],         // Sheen
  0x0635: [0xfeb9, 0xfeba, 0xfebb, 0xfebc],         // Sad
  0x0636: [0xfebd, 0xfebe, 0xfebf, 0xfec0],         // Dad
  0x0637: [0xfec1, 0xfec2, 0xfec3, 0xfec4],         // Tah
  0x0638: [0xfec5, 0xfec6, 0xfec7, 0xfec8],         // Zah
  0x0639: [0xfec9, 0xfeca, 0xfecb, 0xfecc],         // Ain
  0x063a: [0xfecd, 0xfece, 0xfecf, 0xfed0],         // Ghain
  0x0641: [0xfed1, 0xfed2, 0xfed3, 0xfed4],         // Fa
  0x0642: [0xfed5, 0xfed6, 0xfed7, 0xfed8],         // Qaf
  0x0643: [0xfed9, 0xfeda, 0xfedb, 0xfedc],         // Kaf
  0x0644: [0xfedd, 0xfede, 0xfedf, 0xfee0],         // Lam
  0x0645: [0xfee1, 0xfee2, 0xfee3, 0xfee4],         // Meem
  0x0646: [0xfee5, 0xfee6, 0xfee7, 0xfee8],         // Noon
  0x0647: [0xfee9, 0xfeea, 0xfeeb, 0xfeec],         // Ha
  0x0648: [0xfeed, 0xfeee],                          // Waw
  0x0649: [0xfeef, 0xfef0],                          // Alef Maksura
  0x064a: [0xfef1, 0xfef2, 0xfef3, 0xfef4],         // Ya
  0x0622: [0xfe81, 0xfe82],                          // Alef Madda
  0x0623: [0xfe83, 0xfe84],                          // Alef Hamza Above
  0x0624: [0xfe85, 0xfe86],                          // Waw Hamza
  0x0625: [0xfe87, 0xfe88],                          // Alef Hamza Below
  0x0626: [0xfe89, 0xfe8a, 0xfe8b, 0xfe8c],         // Ya Hamza
  0x0629: [0xfe93, 0xfe94],                          // Teh Marbuta

  // --- Extended Arabic: Persian/Urdu ---
  0x067e: [0xfb56, 0xfb57, 0xfb58, 0xfb59],         // Peh (پ)
  0x0686: [0xfb7a, 0xfb7b, 0xfb7c, 0xfb7d],         // Tcheh (چ)
  0x0698: [0xfb8a, 0xfb8b],                          // Zheh (ژ) — right-joining
  0x06af: [0xfb92, 0xfb93, 0xfb94, 0xfb95],         // Gaf (گ)

  // --- Extended Arabic: Kurdish ---
  0x06a4: [0xfb6a, 0xfb6b, 0xfb6c, 0xfb6d],         // Veh (ڤ)
  0x06a9: [0xfb8e, 0xfb8f, 0xfb90, 0xfb91],         // Kurdish Kef (ک)

  // --- Additional letters ---
  0x0671: [0xfb50, 0xfb51],                          // Alef Wasla (ٱ) — right-joining
  0x0621: [0xfe80, 0xfe80],                          // Hamza (ء) — non-joining (isolated only, same form for all)
  0x06c3: [0xfb84, 0xfb85],                          // Teh Marbuta Goal (ة) — right-joining
};

// =============================================================================
// BiDi Reordering (Simplified UAX #9)
// =============================================================================

/**
 * Apply simplified Unicode BiDi Algorithm to reorder runs for visual display.
 * This handles the most common case: mixing LTR and RTL text.
 */
function applyBidiReorder(runs: ScriptRun[], baseDir: BiDiDirection): ScriptRun[] {
  if (runs.length <= 1) {
    // Single run — just reverse if RTL
    if (runs.length === 1 && runs[0].direction === "rtl") {
      runs[0].clusters.reverse();
    }
    return runs;
  }

  // Assign embedding levels:
  // In base LTR: LTR runs get level 0, RTL runs get level 1
  // In base RTL: RTL runs get level 1, LTR runs get level 2
  const levels: number[] = runs.map(run => {
    if (baseDir === "ltr") {
      return run.direction === "rtl" ? 1 : 0;
    } else {
      return run.direction === "ltr" ? 2 : 1;
    }
  });

  // Reverse clusters within RTL runs
  for (let i = 0; i < runs.length; i++) {
    if (runs[i].direction === "rtl") {
      runs[i].clusters.reverse();
    }
  }

  // Find highest level. Use an iterative max — `Math.max(...levels)`
  // overflows the call stack once `levels.length` exceeds ~125k (V8's
  // spread-argument limit), which is reachable for legitimately long
  // RTL paragraphs.
  let maxLevel = 0;
  for (const lv of levels) {
    if (lv > maxLevel) {
      maxLevel = lv;
    }
  }

  // Reverse sequences of runs at each level from highest to lowest
  const result = [...runs];
  for (let level = maxLevel; level > 0; level--) {
    let i = 0;
    while (i < result.length) {
      if (levels[i] >= level) {
        // Find the extent of this sequence at this level
        let j = i + 1;
        while (j < result.length && levels[j] >= level) {
          j++;
        }
        // Reverse the segment [i, j)
        const segment = result.slice(i, j);
        segment.reverse();
        result.splice(i, j - i, ...segment);
        // Also reverse corresponding levels
        const levelSeg = levels.slice(i, j);
        levelSeg.reverse();
        levels.splice(i, j - i, ...levelSeg);
        i = j;
      } else {
        i++;
      }
    }
  }

  return result;
}
