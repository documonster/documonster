import { hasCjk } from "@utils/cjk";
/**
 * What a document assembled with `Document.create()` still needs when it is
 * written out.
 *
 * Kept apart from `document-handle.ts` so the packager and the PDF bridge can
 * apply it without importing the builder helpers.
 */
import { eastAsianDefaultsFor, withEastAsianDefaults } from "@word/core/east-asian-defaults";
import { extractChildText } from "@word/core/text-utils";
import { walkDocument } from "@word/core/walker";
import { extractText } from "@word/query/search";
import type { DocDefaults, DocxDocument } from "@word/types";

/**
 * The document as it should be written: when it asks for it
 * (`deriveEastAsianDefaults`, set by `Document.create()`), the East Asian
 * proofing language and face are derived from its text unless the caller
 * already stated them, and the defaults of a new Word document are stated
 * wherever it leaves them out ({@link withStatedDefaults}). Any other document
 * is returned unchanged.
 *
 * Doing this at write time rather than on every edit is what lets the answer
 * come from the finished text. Doing it only when asked is what keeps a
 * read → edit → write round trip from acquiring a language its author never
 * chose.
 *
 * Without it, every manually built Chinese document opens in Word with a red
 * underline under every word — the same defect the Markdown importer had,
 * reached by the other construction path. `useDefaultStyles` cannot fix it,
 * because it runs before any content exists; that is why it states only the
 * Latin language and leaves both East Asian slots to be filled from the text
 * here.
 *
 * A stated `language.eastAsia` stops this outright rather than merging field
 * by field: a caller who named the language owns the East Asian declaration,
 * including the choice to leave the face to Word. It also means `extractText`
 * is not walked for a document that has already answered.
 *
 * The input is never modified. A document that asked is returned as a shallow
 * copy with the request answered and the flag cleared, so finalizing the result
 * again — the PDF bridge finalizes, then lays out, which finalizes too — is
 * O(1) instead of another walk of the text.
 */
export function finalizeDocument(doc: DocxDocument): DocxDocument {
  if (doc.deriveEastAsianDefaults !== true) {
    return doc;
  }
  const { deriveEastAsianDefaults: _, ...stated } = doc;
  const rest: DocxDocument = { ...stated, docDefaults: withStatedDefaults(stated.docDefaults) };
  if (rest.docDefaults?.runProperties?.language?.eastAsia !== undefined) {
    return rest;
  }
  // Most documents have no East Asian text at all, and answering that needs
  // neither the joined text nor the language detector: stop at the first
  // paragraph that has some.
  let anyCjk = false;
  walkDocument(rest, {
    enterParagraph(para) {
      // Child by child: the paragraph's text is their concatenation, so it
      // has East Asian text exactly when one of them does.
      for (const child of para.children) {
        if (hasCjk(extractChildText(child))) {
          anyCjk = true;
          return "stop";
        }
      }
      return "skip";
    }
  });
  const derived = anyCjk ? eastAsianDefaultsFor(extractText(rest)) : undefined;
  if (!derived) {
    return rest;
  }
  return {
    ...rest,
    docDefaults: {
      ...rest.docDefaults,
      runProperties: withEastAsianDefaults(rest.docDefaults?.runProperties, derived)
    }
  };
}

/**
 * `defaults` with the defaults of a new document in current Word where it
 * states nothing: Aptos 12pt (`en-US`), 8pt after each paragraph, 1.16 lines
 * (`w:line="278"`) — Word for Mac 16.114 sets a file stating none of them
 * exactly so (`word-pagination/02-nods-plain`).
 *
 * Without `w:docDefaults` — or with an empty one — Word takes the defaults from
 * the user's Normal template, so the same file opened differently in different
 * Words. A document the library builds therefore states them: it is laid out
 * with them and written with them, and every Word sets it as the library laid
 * it out. Only what is missing is filled: a stated value is never replaced.
 * Documents that were read keep their defaults, or the lack of them.
 *
 * Fonts are named rather than referenced through the theme, because the
 * library's theme part names Calibri. Word's `w:kern="2"` and
 * `w14:ligatures` are left out on purpose: kerning moves line breaks, and
 * layout does not kern.
 */
export function withStatedDefaults(defaults: DocDefaults | undefined): DocDefaults {
  const run = defaults?.runProperties;
  const font = run?.font;
  const spacing = defaults?.paragraphProperties?.spacing;
  return {
    ...defaults,
    runProperties: {
      ...run,
      font:
        typeof font === "string"
          ? font
          : {
              ...font,
              // One Latin slot named stands for both, as layout reads it.
              ascii:
                font?.ascii ?? (font?.asciiTheme ? undefined : (font?.hAnsi ?? NEW_DOCUMENT_FONT)),
              hAnsi:
                font?.hAnsi ?? (font?.hAnsiTheme ? undefined : (font?.ascii ?? NEW_DOCUMENT_FONT))
            },
      size: run?.size ?? NEW_DOCUMENT_SIZE,
      sizeCs: run?.sizeCs ?? NEW_DOCUMENT_SIZE,
      language: { ...run?.language, val: run?.language?.val ?? "en-US" }
    },
    paragraphProperties: {
      ...defaults?.paragraphProperties,
      spacing: {
        ...spacing,
        after: spacing?.after ?? 160,
        ...(spacing?.line === undefined ? { line: 278, lineRule: "auto" as const } : {})
      }
    }
  };
}

/** Latin face of a new document in current Word. */
const NEW_DOCUMENT_FONT = "Aptos";
/** Size of a new document in current Word, in half-points (12pt). */
const NEW_DOCUMENT_SIZE = 24;
