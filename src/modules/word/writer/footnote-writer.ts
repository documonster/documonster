/**
 * DOCX Writers - Footnotes & Endnotes
 *
 * Renders word/footnotes.xml and word/endnotes.xml parts.
 */

import { DOCUMENT_NAMESPACES, STD_DOC_ATTRIBUTES } from "@word/constants";
import { within } from "@word/core/node-path";
import type { FootnoteDef, EndnoteDef } from "@word/types";
import {
  DEFAULT_CONTINUATION_SEPARATOR_ID,
  DEFAULT_SEPARATOR_ID,
  hasOwnSeparators
} from "@word/writer/note-separators";
import { renderParagraph } from "@word/writer/paragraph-writer";
import type { RenderHelpers } from "@word/writer/render-context";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

/** Render the default separator/continuation separator entries (required by Word). */
function renderDefaultSeparators(xml: XmlSink, elementName: string): void {
  // Separator (id=-1)
  xml.openNode(elementName, {
    "w:type": "separator",
    "w:id": xmlValue(DEFAULT_SEPARATOR_ID, "w:id")
  });
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.leafNode("w:separator");
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();

  // Continuation separator (id=0)
  xml.openNode(elementName, {
    "w:type": "continuationSeparator",
    "w:id": xmlValue(DEFAULT_CONTINUATION_SEPARATOR_ID, "w:id")
  });
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.leafNode("w:continuationSeparator");
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();
}

/** Render a single footnote/endnote entry. */
function renderNote(
  xml: XmlSink,
  elementName: string,
  note: FootnoteDef | EndnoteDef,
  helpers?: RenderHelpers
): void {
  const attrs: Record<string, string> = { "w:id": xmlValue(note.id, "w:id") };
  if (note.type && note.type !== "normal") {
    attrs["w:type"] = note.type;
  }
  xml.openNode(elementName, attrs);
  note.content.forEach((para, i) =>
    within(`content[${i}]`, () => renderParagraph(xml, para, helpers))
  );
  xml.closeNode();
}

/** Render word/footnotes.xml. */
export function renderFootnotes(
  xml: XmlSink,
  footnotes: readonly FootnoteDef[],
  helpers?: RenderHelpers
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  // Every prefix the document part declares: drawings and markup kept
  // verbatim inside a note use the same ones.
  xml.openNode("w:footnotes", DOCUMENT_NAMESPACES);

  // If the caller already provided separator entries (identified by their
  // `type` rather than a magic id range), use those. Otherwise synthesize
  // the default separators, which take ids -1 and 0 — `assertUniqueReferencedIds`
  // has already refused a normal note holding either. Keying off `id <= 0`
  // instead would let a normal note with id 0 suppress the mandatory
  // separators and produce an invalid document.
  if (!hasOwnSeparators(footnotes)) {
    renderDefaultSeparators(xml, "w:footnote");
  }

  footnotes.forEach((fn, i) =>
    within(`footnotes[${i}]`, () => renderNote(xml, "w:footnote", fn, helpers))
  );

  xml.closeNode();
}

/** Render word/endnotes.xml. */
export function renderEndnotes(
  xml: XmlSink,
  endnotes: readonly EndnoteDef[],
  helpers?: RenderHelpers
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  // Every prefix the document part declares: drawings and markup kept
  // verbatim inside a note use the same ones.
  xml.openNode("w:endnotes", DOCUMENT_NAMESPACES);

  if (!hasOwnSeparators(endnotes)) {
    renderDefaultSeparators(xml, "w:endnote");
  }

  endnotes.forEach((en, i) =>
    within(`endnotes[${i}]`, () => renderNote(xml, "w:endnote", en, helpers))
  );

  xml.closeNode();
}
