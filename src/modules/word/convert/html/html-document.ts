/**
 * HTML → complete DOCX document (`WordHtml.toDocx`), and HTML → a body with
 * every definition it refers to (`WordHtml.toDocxBody`) — the HTML
 * counterparts of `WordMarkdown.toDocx` and `WordMarkdown.toDocxBody`.
 *
 * The internal `htmlToDocx` returns body content and images only, and the body
 * references definitions it does not carry: `Heading1…6` styles and the list
 * numbering instances `HTML_BULLET_LIST_NUM_ID` and `HTML_NUMBERED_LIST_NUM_ID`.
 * Both forms here carry them.
 */
import {
  addContent,
  addListDefinition,
  create,
  useDefaultStyles
} from "@word/builder/document-handle";
import {
  convertHtml,
  HTML_BULLET_LIST_NUM_ID,
  HTML_NUMBERED_LIST_NUM_ID
} from "@word/convert/html/html-import";
import type { HtmlImportOptions } from "@word/convert/html/html-import";
import { referencedStyleIds } from "@word/core/builtin-styles";
import type {
  AbstractNumbering,
  BodyContent,
  DocDefaults,
  DocxDocument,
  ImageDef,
  NumberingInstance,
  StyleDef
} from "@word/types";

/**
 * Convert an HTML string into a complete document, ready for `Io.toBuffer`.
 *
 * The result carries the default style set (including the heading styles the
 * converted `<h1>…<h6>` refer to), bullet and numbered list definitions, and
 * every image decoded from a `data:` URL. A picture it cannot embed — a remote
 * or relative `src`, an undecodable `data:` URL — is written as the italic
 * `[Image: alt]` text `WordMarkdown.toDocx` writes for an image it cannot resolve.
 *
 * @param html - The HTML string to convert.
 * @param options - Conversion settings.
 * @returns A new document.
 *
 * @example
 * ```ts
 * import { WordHtml } from "documonster/word/html";
 * import { Io } from "documonster/word";
 *
 * const doc = WordHtml.toDocx("<h1>Title</h1><ul><li>One</li></ul>");
 * const bytes = await Io.toBuffer(doc);
 * ```
 */
export function htmlToDocument(html: string, options?: HtmlImportOptions): DocxDocument {
  return assemble(html, options, "WordHtml.toDocx");
}

function assemble(
  html: string,
  options: HtmlImportOptions | undefined,
  where: string
): DocxDocument {
  const { body, images } = convertHtml(html, options, where, true);
  const doc = create();
  useDefaultStyles(doc);
  addListDefinition(doc, "bullet", { numId: HTML_BULLET_LIST_NUM_ID });
  addListDefinition(doc, "numbered", { numId: HTML_NUMBERED_LIST_NUM_ID });
  for (const block of body) {
    addContent(doc, block);
  }
  return images.length > 0 ? { ...doc, images } : doc;
}

/**
 * What {@link htmlToDocxFragment} returns: converted body content and the
 * definitions it refers to.
 *
 * The body is not self-contained: headings refer to `Heading1…6`, list
 * paragraphs to the numbering instances `HTML_BULLET_LIST_NUM_ID` and
 * `HTML_NUMBERED_LIST_NUM_ID`, and pictures to their images by rId. Splicing
 * the `body` alone into a host document that lacks them leaves the text
 * unstyled, the lists unnumbered and the pictures dangling.
 *
 * Its ids are fixed — numbering instances, abstract numberings and image rIds
 * start at the same values every time, whatever the host already uses — so
 * copying these arrays into a document that has lists or pictures of its own
 * gives duplicate ids (`DocxDuplicateIdError` at write). To splice HTML into
 * an existing document, convert it whole and merge the two, which renumbers:
 * `Transform.merge([host, WordHtml.toDocx(html)])`. Use this fragment where
 * the host has none of its own, or to inspect the conversion.
 */
export interface HtmlBodyResult {
  readonly body: BodyContent[];
  /** Images decoded from `data:` URLs, each with the rId the body's picture runs reference. */
  readonly images: ImageDef[];
  readonly abstractNumberings: AbstractNumbering[];
  readonly numberingInstances: NumberingInstance[];
  /** The style definitions the body refers to by id. */
  readonly styles: StyleDef[];
  /** Document defaults the body assumes. */
  readonly docDefaults: DocDefaults;
}

/**
 * Convert an HTML string into body content plus every definition it refers
 * to — styles, list numbering, images — for splicing into a document of your
 * own. {@link htmlToDocument} gives the same content as a complete document.
 *
 * @param html - The HTML string to convert.
 * @param options - Conversion settings.
 */
export function htmlToDocxFragment(html: string, options?: HtmlImportOptions): HtmlBodyResult {
  const doc = assemble(html, options, "WordHtml.toDocxBody");
  const referenced = referencedStyleIds(doc);
  return {
    body: [...doc.body],
    images: [...(doc.images ?? [])],
    abstractNumberings: [...(doc.abstractNumberings ?? [])],
    numberingInstances: [...(doc.numberingInstances ?? [])],
    styles: (doc.styles ?? []).filter(st => referenced.has(st.styleId) || st.isDefault),
    docDefaults: doc.docDefaults ?? {}
  };
}
