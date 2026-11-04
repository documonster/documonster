/**
 * DOCX Writers - Table of Contents
 *
 * Renders TOC as a structured document tag (SDT) with field codes.
 */

import { isRun } from "@word/core/text-utils";
import { buildTocInstruction, findTocFieldPieces } from "@word/core/toc-instruction";
import type { TocFieldPieces } from "@word/core/toc-instruction";
import type { FieldContent, Paragraph, ParagraphChild, Run, TableOfContents } from "@word/types";
import { renderParagraph } from "@word/writer/paragraph-writer";
import type { RenderHelpers } from "@word/writer/render-context";
import type { XmlSink } from "@xml/types";

/**
 * Render a Table of Contents block. `helpers` are the surrounding part's, so
 * the cached entries' bookmarks and hyperlinks resolve against the same maps
 * as the rest of the part.
 */
export function renderTableOfContents(
  xml: XmlSink,
  toc: TableOfContents,
  helpers?: RenderHelpers
): void {
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

  // The entries holding the field (or a block with none): as they are.
  const inline =
    toc.field === "entries" ? findTocFieldPieces(toc.cachedParagraphs ?? []) : undefined;
  if (inline || toc.field === "none") {
    const paragraphs = inline
      ? withTocInstruction(toc.cachedParagraphs!, inline, instruction)
      : (toc.cachedParagraphs ?? []);
    if (paragraphs.length > 0) {
      for (const para of paragraphs) {
        renderParagraph(xml, para, helpers);
      }
    } else {
      xml.leafNode("w:p");
    }
    xml.closeNode(); // w:sdtContent
    xml.closeNode(); // w:sdt
    return;
  }

  // The field opens in the first entry's paragraph and closes in a paragraph
  // of its own after the last, which is how Word writes a table of contents it
  // generated. Opening it in a paragraph of its own instead left a blank line
  // above the entries that Word removes when it updates the field, so every
  // page number after it moved.
  const begin: Run = { content: [{ type: "field", fieldChar: "begin", instruction }] };
  const separate: Run = { content: [{ type: "field", fieldChar: "separate", instruction: "" }] };
  const [first, ...rest]: readonly Paragraph[] =
    toc.cachedParagraphs && toc.cachedParagraphs.length > 0
      ? toc.cachedParagraphs
      : [
          {
            type: "paragraph",
            children: [
              {
                content: [
                  {
                    type: "text",
                    text: "Update this table of contents by right-clicking and selecting 'Update Field'"
                  }
                ]
              }
            ]
          }
        ];
  renderParagraph(xml, { ...first!, children: [begin, separate, ...first!.children] }, helpers);
  for (const para of rest) {
    renderParagraph(xml, para, helpers);
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

/**
 * `paragraphs` with the TOC field they hold in pieces given `instruction`:
 * unchanged while it is the field's own, otherwise written beside its begin.
 */
function withTocInstruction(
  paragraphs: readonly Paragraph[],
  field: TocFieldPieces,
  instruction: string
): readonly Paragraph[] {
  if (instruction === field.instruction) {
    return paragraphs;
  }
  const replaced = new Map<FieldContent, FieldContent>([
    [field.begin, { ...field.begin, instruction }]
  ]);
  if (field.next) {
    replaced.set(field.next, { ...field.next, instruction: "" });
  }
  const touched = (child: ParagraphChild): boolean =>
    isRun(child) && child.content.some(c => replaced.has(c as FieldContent));
  return paragraphs.map(para => {
    if (!para.children.some(touched)) {
      return para;
    }
    const children = para.children.map((child): ParagraphChild =>
      isRun(child) && touched(child)
        ? { ...child, content: child.content.map(c => replaced.get(c as FieldContent) ?? c) }
        : child
    );
    return { ...para, children };
  });
}
