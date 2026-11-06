/**
 * The font decisions every writer of drawn text makes once its pages are preflighted, and
 * before it writes font resources: auto-discover a system face for what the configured faces
 * cannot draw, then report what is still uncovered and the font families it did not know.
 *
 * Shared by `Pdf.Builder#toBuffer` and every `Pdf.Editor` write, so text drawn on an editor
 * page is embedded, and reported, exactly as the same text on a builder page.
 */

import type { FontManager } from "@pdf/font/font-manager";
import { findSystemFontForCodePoints } from "@pdf/font/system-fonts";
import type { CjkLanguage } from "@utils/cjk";

/** The font settings a writer was given — the Builder's, or the Editor's read options. */
export interface DrawnTextFontSettings {
  readonly disableFontAutoDiscovery: boolean;
  readonly preferSystemFonts: readonly string[];
  readonly textLanguage: CjkLanguage | undefined;
  readonly onWarning: ((message: string) => void) | undefined;
}

/**
 * Run discovery and diagnostics for one build. Call inside {@link withDrawnTextFontBuild},
 * with the pages already preflighted and before `writeFontResources`.
 */
export function discoverDrawnTextFonts(
  fontManager: FontManager,
  settings: DrawnTextFontSettings
): void {
  // Discovery runs when nothing is embedded yet, and *also* when a fallback face
  // is present but cannot draw everything the document ended up containing.
  //
  // The second case arises when a caller shares a font engine it already
  // configured. The Word bridge measures a document against the face it chose and
  // then hands the engine over — but a chart paints its own axis and category
  // labels straight onto the page, and those strings are not in the layout model
  // the bridge inspected. `hasEmbeddedFont()` answers "is a face registered",
  // when the question is "can every character be drawn": a document pinned to
  // Kaiti SC lost `龦` to a `.notdef` box although Heiti SC, equally installed,
  // has it. Re-running discovery over the complete repertoire finds a face that
  // covers all of it.
  //
  // A face the caller embedded explicitly (`embedFont`) is never reconsidered:
  // that is a statement about the document, not a best-effort guess.
  const uncoveredByFallback = fontManager.hasFallbackFont()
    ? fontManager.getUncoveredFallbackCodePoints()
    : null;

  if (!fontManager.hasEmbeddedFont() || (uncoveredByFallback?.size ?? 0) > 0) {
    // Auto-discover a system font when the document contains non-WinAnsi
    // characters (CJK, accented code points beyond WinAnsi, etc.) and
    // the caller did not supply one via `embedFont`. Mirrors the
    // `pdf-exporter` pipeline so `drawChartPdf`-style overlays, ad-hoc
    // `drawText` usage, and the spreadsheet exporter all reach for the
    // same system fonts before falling back to Type3 NOTDEF glyphs.
    // Failures are non-fatal: discovery is best-effort and users who
    // need guaranteed rendering should still call `embedFont` with a
    // font they control.
    //
    // It is registered as a *fallback*, not as the document font: it only
    // lends glyphs for the code points WinAnsi cannot encode. Registering it
    // as the document font would route every run through one regular face,
    // so a single `→` would strip bold, italic and monospace from the whole
    // document and desynchronise the widths already measured during layout.
    const nonWinAnsi = fontManager.getType3CodePoints();
    if (nonWinAnsi.size > 0) {
      // Try auto-discovery unless the caller opted out.
      if (!settings.disableFontAutoDiscovery) {
        const discovered = findSystemFontForCodePoints(
          nonWinAnsi,
          settings.preferSystemFonts,
          // Counted as the text was tracked, not derived from `nonWinAnsi`:
          // that is a Set, and a Set cannot say that `国` outnumbers `國`.
          settings.textLanguage ?? fontManager.getTextLanguage()
        );
        if (discovered) {
          // `widenFallbackFont` rather than `registerFallbackFont`: this runs after
          // the caller has drawn (and measured) text, and replacing the face outright
          // changes the advance of every non-Latin-1 code point the two faces share —
          // 345 of them between `Kaiti SC` and `Heiti SC`. It refuses the swap when
          // that would happen and keeps the measured face.
          if (fontManager.widenFallbackFont(discovered)) {
            // Noted rather than warned here: `reportDiagnostics` below raises it, and
            // it is the one path every pipeline shares. Raising it at each discovery
            // site is what left three of the five entry points silent.
            fontManager.noteAutoDiscoveredFont(discovered.familyName, nonWinAnsi.size);
          }
        }
      }
      // The "nothing covers these" warning is not raised here. It is the same
      // condition `PdfExportOptions.onWarning` documents, and the spreadsheet
      // exporter is a separate pipeline that never reached this block — so
      // `Pdf.create` and `Pdf.fromExcel` emitted tofu boxes silently. It now
      // lives in `FontManager.reportDiagnostics`, called below and by every
      // other pipeline, and is evaluated after the discovery attempt above.
    }
  }

  if (settings.onWarning) {
    fontManager.reportDiagnostics(settings.onWarning);
  }

  // Surface any unknown font families the FontManager saw during
  // resolveFont (one diagnostic per distinct family, independent of
  // how many text runs used it). An embedded font shadows everything
  // anyway, so skip the warning in that case to avoid noise.
  if (!fontManager.hasEmbeddedFont()) {
    const unknown = fontManager.getUnknownFontFamilies();
    if (unknown.size > 0) {
      const list = [...unknown].slice(0, 5).join(", ");
      const more = unknown.size > 5 ? ` (and ${unknown.size - 5} more)` : "";
      settings.onWarning?.(
        `Font family ${unknown.size > 1 ? "names" : "name"} '${list}'${more} not recognised; ` +
          `falling back to Helvetica metrics. Add the typeface to the font-family map or call embedFont(bytes).`
      );
    }
  }
}

/**
 * Run one build's font session: `beginBuild`, the build, `endBuild` — and drop a face that
 * auto-discovery embedded for it. Discovery is a decision for one build, not a change to the
 * document: a later build may hold characters that need a different (broader) system face.
 */
export async function withDrawnTextFontBuild<T>(
  fontManager: FontManager,
  build: () => Promise<T>
): Promise<T> {
  const hadEmbeddedFont = fontManager.hasEmbeddedFont();
  fontManager.beginBuild();
  try {
    return await build();
  } finally {
    if (!hadEmbeddedFont && fontManager.hasEmbeddedFont()) {
      fontManager.clearAutoDiscoveredFont();
    }
    fontManager.endBuild();
  }
}
