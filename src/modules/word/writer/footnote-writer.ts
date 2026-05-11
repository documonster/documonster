/**
 * DOCX Writers - Footnotes & Endnotes
 *
 * Renders word/footnotes.xml and word/endnotes.xml parts.
 */

import type { XmlSink } from "@xml/types";

import { NS_W, NS_R, STD_DOC_ATTRIBUTES } from "../constants";
import type { FootnoteDef, EndnoteDef } from "../types";
import { renderParagraph } from "./paragraph-writer";
import type { RenderHelpers } from "./render-context";

/** Render the default separator/continuation separator entries (required by Word). */
function renderDefaultSeparators(xml: XmlSink, elementName: string): void {
  // Separator (id=-1)
  xml.openNode(elementName, { "w:type": "separator", "w:id": "-1" });
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.leafNode("w:separator");
  xml.closeNode();
  xml.closeNode();
  xml.closeNode();

  // Continuation separator (id=0)
  xml.openNode(elementName, { "w:type": "continuationSeparator", "w:id": "0" });
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
  const attrs: Record<string, string> = { "w:id": String(note.id) };
  if (note.type && note.type !== "normal") {
    attrs["w:type"] = note.type;
  }
  xml.openNode(elementName, attrs);
  for (const para of note.content) {
    renderParagraph(xml, para, helpers);
  }
  xml.closeNode();
}

/** Render word/footnotes.xml. */
export function renderFootnotes(
  xml: XmlSink,
  footnotes: readonly FootnoteDef[],
  helpers?: RenderHelpers
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:footnotes", {
    "xmlns:w": NS_W,
    "xmlns:r": NS_R
  });

  // If the caller already provided separator entries (by having any id <= 0), use those.
  // Otherwise, synthesize the default separators.
  const hasSeparators = footnotes.some(fn => fn.id <= 0);
  if (!hasSeparators) {
    renderDefaultSeparators(xml, "w:footnote");
  }

  for (const fn of footnotes) {
    renderNote(xml, "w:footnote", fn, helpers);
  }

  xml.closeNode();
}

/** Render word/endnotes.xml. */
export function renderEndnotes(
  xml: XmlSink,
  endnotes: readonly EndnoteDef[],
  helpers?: RenderHelpers
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:endnotes", {
    "xmlns:w": NS_W,
    "xmlns:r": NS_R
  });

  const hasSeparators = endnotes.some(en => en.id <= 0);
  if (!hasSeparators) {
    renderDefaultSeparators(xml, "w:endnote");
  }

  for (const en of endnotes) {
    renderNote(xml, "w:endnote", en, helpers);
  }

  xml.closeNode();
}
