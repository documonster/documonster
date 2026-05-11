/**
 * DOCX Writers - Table of Contents
 *
 * Renders TOC as a structured document tag (SDT) with field codes.
 */

import type { XmlSink } from "@xml/types";

import type { TableOfContents } from "../types";
import { renderParagraph } from "./paragraph-writer";

/** Build TOC field instruction from properties. */
function buildTocInstruction(toc: TableOfContents): string {
  const parts: string[] = ["TOC"];

  if (toc.headingStyleRange) {
    parts.push(`\\o "${toc.headingStyleRange}"`);
  }
  if (toc.hyperlink) {
    parts.push("\\h");
  }
  if (toc.stylesWithLevels && toc.stylesWithLevels.length > 0) {
    // Format: "Style1,Level1;Style2,Level2;..."
    const entries = toc.stylesWithLevels.map(s => `${s.styleName},${s.level}`).join(";");
    parts.push(`\\t "${entries}"`);
  }
  if (toc.captionLabel) {
    parts.push(`\\c "${toc.captionLabel}"`);
  }
  if (toc.sequenceFieldIdentifier) {
    parts.push(`\\s "${toc.sequenceFieldIdentifier}"`);
  }

  return ` ${parts.join(" ")} `;
}

/** Render a Table of Contents block. */
export function renderTableOfContents(xml: XmlSink, toc: TableOfContents): void {
  const instruction = buildTocInstruction(toc);

  // Wrap in SDT for proper Word handling
  xml.openNode("w:sdt");
  xml.openNode("w:sdtPr");
  xml.openNode("w:docPartObj");
  xml.leafNode("w:docPartGallery", { "w:val": "Table of Contents" });
  xml.leafNode("w:docPartUnique");
  xml.closeNode();
  xml.closeNode();

  xml.openNode("w:sdtContent");

  // Begin field paragraph
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.leafNode("w:fldChar", { "w:fldCharType": "begin" });
  xml.closeNode();
  xml.openNode("w:r");
  xml.openNode("w:instrText", { "xml:space": "preserve" });
  xml.writeText(instruction);
  xml.closeNode();
  xml.closeNode();
  xml.openNode("w:r");
  xml.leafNode("w:fldChar", { "w:fldCharType": "separate" });
  xml.closeNode();
  xml.closeNode();

  // Cached paragraphs (placeholder content until Word updates the field)
  if (toc.cachedParagraphs && toc.cachedParagraphs.length > 0) {
    for (const para of toc.cachedParagraphs) {
      renderParagraph(xml, para);
    }
  } else {
    // Default placeholder
    xml.openNode("w:p");
    xml.openNode("w:r");
    xml.openNode("w:t");
    xml.writeText("Update this table of contents by right-clicking and selecting 'Update Field'");
    xml.closeNode();
    xml.closeNode();
    xml.closeNode();
  }

  // End field paragraph
  xml.openNode("w:p");
  xml.openNode("w:r");
  xml.leafNode("w:fldChar", { "w:fldCharType": "end" });
  xml.closeNode();
  xml.closeNode();

  xml.closeNode(); // w:sdtContent
  xml.closeNode(); // w:sdt
}
