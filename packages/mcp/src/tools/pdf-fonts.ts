/**
 * Font diagnostics for the tools that write a PDF.
 *
 * Every PDF this server produces went through `Pdf.fromDocx` / `Pdf.fromExcel`
 * without an `onWarning` handler, so the one degradation a caller cannot see for
 * themselves was the one thing never reported: a character no available typeface
 * covers still carries its Unicode for copy and search, and draws as a `.notdef`
 * box. The tool then said "Converted … (md → pdf, 74.1 kB)" and the boxes were
 * discovered by a human opening the file — which is exactly how a document of
 * Chinese prose came back with every ideograph blank and nobody was told.
 *
 * The model cannot check this either: the result text says to verify by opening
 * the PDF, and this server deliberately cannot read its own PDF output back
 * structurally. So the writer's own warning is the only signal that exists.
 *
 * @module
 */

import { readFileSync } from "node:fs";

import type { RasterFontSource } from "documonster/draw";
import type { CjkLanguage, PdfFontConfig, PdfFontSource } from "documonster/pdf";
import { z } from "zod";

import type { PdfFontRef, ServerConfig } from "../config.js";
import { toolError } from "../errors.js";

/**
 * Font options for one PDF-producing call: the operator's font, if they named
 * one, and a collector for whatever the writer reports.
 *
 * Returned as one object so a call site cannot wire the diagnostics and forget
 * the font, or the other way round — the two exist for the same failure.
 */
export function pdfFontOptions(config: FontConfig): {
  /** Spread into the PDF call. Never carries anything the library does not define. */
  readonly options: {
    readonly fonts?: PdfFontConfig;
    readonly onWarning: (message: string) => void;
  };
  /** Lines to append to the tool result, most serious first. */
  notes(): string[];
  /** Refuse to write a PDF with boxes in it. See {@link FontWarningCollector.assertDrawable}. */
  assertDrawable(allow: boolean): void;
} {
  const collector = collectFontWarnings();
  const fonts =
    config.pdfFont === undefined
      ? undefined
      : loadFont(config.pdfFont, config.pdfFontFallbacks ?? []);
  return {
    options: {
      ...(fonts === undefined ? {} : { fonts }),
      onWarning: collector.onWarning
    },
    notes: collector.notes,
    assertDrawable: collector.assertDrawable
  };
}

/**
 * Read and cache the operator's font.
 *
 * Cached by path because a CJK face is tens of megabytes and a conversion-heavy
 * session would otherwise re-read it per call. `resolveConfig` has already vetted
 * that the file exists and is TrueType, so a failure here is a font deleted while
 * the server was running — worth reporting as itself rather than as a PDF error.
 */
function loadFont(primary: PdfFontRef, fallbacks: readonly PdfFontRef[]): PdfFontConfig {
  const key = [primary, ...fallbacks].map(refKey).join("|");
  const cached = fontCache.get(key);
  if (cached !== undefined) {
    return cached;
  }
  // The fallbacks are named families rather than a second `default`, because that is the
  // only shape the library has for "try this face, then that one". They are numbered rather
  // than named after a typeface: the operator gave paths, and inventing a family name from a
  // filename would be a guess that then has to match what the document asks for.
  const families = fallbacks.map((ref, index) => ({
    name: `fallback-${index + 1}`,
    faces: { regular: readFace(ref) }
  }));
  const config: PdfFontConfig = {
    default: { regular: readFace(primary) },
    ...(families.length === 0
      ? {}
      : { families, fallbackFamilies: families.map(family => family.name) })
  };
  fontCache.set(key, config);
  return config;
}

/** Identify one face for the cache: the same file at two indices is two faces. */
function refKey(ref: PdfFontRef): string {
  return ref.collectionIndex === undefined ? ref.path : `${ref.path}#${ref.collectionIndex}`;
}

/**
 * The parts of the server config this module reads.
 *
 * Named rather than repeated inline: three exported functions take it, and an inline literal
 * in each is three places to update when a font option is added.
 */
export type FontConfig = Pick<ServerConfig, "pdfFont" | "pdfFontFallbacks">;

/**
 * Every face the operator named, primary first.
 *
 * One definition of "which faces are configured", because the three exported functions below
 * each need it and each spelled it out — including `loadFont`'s cache key, which has to agree
 * with them or two different chains collide on one entry.
 */
function configuredRefs(config: FontConfig): readonly PdfFontRef[] {
  return config.pdfFont === undefined ? [] : [config.pdfFont, ...(config.pdfFontFallbacks ?? [])];
}

/**
 * One face's bytes, and which face inside them.
 *
 * The shape both consumers actually want. `readFace` below narrows it to the union the library
 * takes; going the other way needed an `instanceof` check to get the index back out.
 *
 * Cached by path — not by face — because a `.ttc` naming two of its faces would otherwise be
 * read twice, and a CJK collection is tens of megabytes.
 */
function faceBytes(ref: PdfFontRef): { bytes: Uint8Array; collectionIndex?: number } {
  let bytes = byteCache.get(ref.path);
  if (bytes === undefined) {
    try {
      bytes = new Uint8Array(readFileSync(ref.path));
    } catch (cause) {
      throw toolError.unsupported(
        `a font configured with --pdf-font could not be read: ${ref.path}`,
        "It existed at startup, so it has been moved or deleted since.",
        { cause }
      );
    }
    byteCache.set(ref.path, bytes);
  }
  return ref.collectionIndex === undefined
    ? { bytes }
    : { bytes, collectionIndex: ref.collectionIndex };
}

/**
 * Read one face's bytes, keeping `collectionIndex` attached.
 *
 * Cached by path — not by face — because a `.ttc` naming two of its faces would otherwise be
 * read twice, and a CJK collection is tens of megabytes.
 */
function readFace(ref: PdfFontRef): PdfFontSource {
  const face = faceBytes(ref);
  return face.collectionIndex === undefined
    ? face.bytes
    : { data: face.bytes, collectionIndex: face.collectionIndex };
}

const byteCache = new Map<string, Uint8Array>();

/**
 * The operator's faces as bytes plus face index, for an API that takes exactly that.
 *
 * `PdfEditor.embedFont` is one — overlay text is drawn by this library, so it needs a real
 * face — and it cannot take a `PdfFontConfig`. Expressed here rather than at the call site so
 * that the `.ttc` index travels: dropping it embedded whichever face came first, which for
 * `Songti.ttc` is weight 900.
 */
export function overlayFaces(
  config: FontConfig
): Array<{ bytes: Uint8Array; collectionIndex?: number }> {
  return configuredRefs(config).map(faceBytes);
}

/**
 * `textLanguage`, as a tool parameter.
 *
 * Shared so the description is written once: it is the same option for every tool that writes
 * a PDF, and the schema budget is a real limit (52,000 characters for the whole list).
 *
 * Worth a parameter rather than a server flag because it is a property of the *document*, not
 * of the host — one session legitimately converts a Chinese file and then a Japanese one.
 */
export const textLanguageShape = {
  textLanguage: z
    .enum(["zh-Hans", "zh-Hant", "ja", "ko"])
    .optional()
    .describe(
      "East Asian language of the content, so the embedded face is drawn in that regional hand. " +
        "Han characters are shared between Chinese, Japanese and Korean but drawn differently, so " +
        "a font picked purely by coverage can be correct and still look wrong. Inferred from the " +
        "text when omitted, which defaults to Chinese for characters common to all three."
    )
} as const;

/** Spread into a PDF call. Empty when the caller said nothing. */
export function textLanguageOption(args: { readonly textLanguage?: CjkLanguage }): {
  readonly textLanguage?: CjkLanguage;
} {
  return args.textLanguage === undefined ? {} : { textLanguage: args.textLanguage };
}

/**
 * The operator's faces in the shapes the diagram backends take.
 *
 * The rasteriser is offered every configured face, primary and fallbacks alike, because it
 * resolves each character against the chain itself — there is no "default" face to be primary
 * *of*. That is also why this cannot simply reuse `PdfFontConfig`.
 */
export function diagramFontOptions(config: FontConfig): {
  readonly pdf?: PdfFontConfig;
  readonly raster?: readonly RasterFontSource[];
  readonly useSystemFonts?: boolean;
} {
  if (config.pdfFont === undefined) {
    return {};
  }
  return {
    pdf: loadFont(config.pdfFont, config.pdfFontFallbacks ?? []),
    raster: configuredRefs(config).map(readFace),
    // The host is switched off once the operator has named their faces, which is what
    // `--pdf-font` is for: "the same Markdown produces a readable PDF on a laptop and a boxed
    // one in a container" is the problem it exists to remove, and leaving discovery on for the
    // raster backend alone would leave half of it. It also matches what the PDF writer already
    // does — a configured font set disables auto-discovery there — so the two backends now
    // answer the same way instead of one of them quietly consulting the machine.
    useSystemFonts: false
  };
}

const fontCache = new Map<string, PdfFontConfig>();

/** A collector to hand to `onWarning`, plus the notes it accumulated. */
export interface FontWarningCollector {
  /** Pass as `onWarning` to any PDF-producing call. */
  readonly onWarning: (message: string) => void;
  /**
   * Lines to append to the tool result, most serious first.
   *
   * Empty when nothing was raised, so a caller can splat it unconditionally.
   */
  notes(): string[];
  /**
   * The writer's own reports of characters that will draw as boxes.
   *
   * Empty when the page is sound, which is what {@link assertDrawable} turns into a
   * refusal.
   */
  missingGlyphs(): string[];
  /**
   * Refuse to hand back a PDF with boxes in it.
   *
   * A PDF is terminal here — this server cannot read its own PDF output structurally, and
   * the result text asks the caller to verify by opening the file — so a boxed page
   * reported as a success is a defect the caller finds after the fact, if at all. That is
   * how a document of Chinese prose came back with every ideograph blank and nobody was
   * told (issue #218). Failing closed makes the one outcome nobody wants impossible to
   * reach by accident, and `allowMissingGlyphs` is how a caller who has decided the boxes
   * are acceptable says so on purpose.
   *
   * @throws {McpToolError} When any character will render as `.notdef`.
   */
  assertDrawable(allow: boolean): void;
}

/**
 * The substring the library uses for a character that will visibly render as a box.
 *
 * `.notdef` is the discriminator because it appears in both of the writer's wordings —
 * "no glyph in any available font and will render as `.notdef` boxes" on the fallback
 * path, and "not covered by the configured PDF font families and will render with the
 * `.notdef` glyph" when the caller supplied a font that does not reach far enough — and in
 * neither of the ones that are merely notes. Matching only the first meant the failure that
 * follows configuring `--pdf-font`, which is the option's whole point, was filed as an
 * aside.
 *
 * Matched rather than re-derived: the condition is decided inside the font manager, and a
 * second rule here would drift from it. `pdf-fonts.contract.test.ts` runs real exports
 * down both paths and asserts the wording still contains it, so a reword in the library
 * fails a test here instead of quietly turning the check off.
 */
const TOFU_MARKER = ".notdef";

export function collectFontWarnings(): FontWarningCollector {
  const messages: string[] = [];
  const missing = (): string[] => messages.filter(message => message.includes(TOFU_MARKER));
  return {
    onWarning: message => {
      messages.push(message);
    },
    missingGlyphs: missing,
    assertDrawable: allow => {
      const boxes = missing();
      if (allow || boxes.length === 0) {
        return;
      }
      throw toolError.unsupported(
        `the PDF was not written: ${boxes.join(" ")}`,
        "Point the server at a font that covers these characters with --pdf-font, or pass " +
          "allowMissingGlyphs: true to accept a page with boxes on it."
      );
    },
    notes: () => {
      if (messages.length === 0) {
        return [];
      }
      const tofu = messages.filter(message => message.includes(TOFU_MARKER));
      const rest = messages.filter(message => !message.includes(TOFU_MARKER));
      return [
        // Bold, because this one means the page is visibly wrong. The rest are
        // notes about how the output was produced, not about it being broken.
        ...tofu.map(message => `- **font coverage**: ${message}`),
        ...rest.map(message => `- font: ${message}`)
      ];
    }
  };
}
