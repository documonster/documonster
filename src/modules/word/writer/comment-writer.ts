/**
 * DOCX Writers - Comments
 *
 * Renders word/comments.xml and word/commentsExtended.xml parts.
 */

import type { XmlSink } from "@xml/types";

import { NS_W, NS_W15, NS_R, STD_DOC_ATTRIBUTES } from "../constants";
import type { CommentDef } from "../types";
import { renderParagraph } from "./paragraph-writer";
import type { RenderHelpers } from "./render-context";

/** Render word/comments.xml. */
export function renderComments(
  xml: XmlSink,
  comments: readonly CommentDef[],
  helpers?: RenderHelpers
): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:comments", {
    "xmlns:w": NS_W,
    "xmlns:r": NS_R
  });

  for (const comment of comments) {
    const attrs: Record<string, string> = {
      "w:id": String(comment.id),
      "w:author": comment.author
    };
    if (comment.date) {
      attrs["w:date"] = comment.date;
    }
    if (comment.initials) {
      attrs["w:initials"] = comment.initials;
    }
    xml.openNode("w:comment", attrs);
    for (const para of comment.content) {
      renderParagraph(xml, para, helpers);
    }
    xml.closeNode();
  }

  xml.closeNode();
}

/**
 * Render word/commentsExtended.xml — extended comment metadata (done/parentId).
 * Only includes comments that have either `done` or `parentId` set.
 * @returns true if any extended comments were rendered.
 */
export function renderCommentsExtended(xml: XmlSink, comments: readonly CommentDef[]): boolean {
  // Collect comments that need extension
  const extended = comments.filter(c => c.done !== undefined || c.parentId !== undefined);
  if (extended.length === 0) {
    return false;
  }

  // Each extended comment must reference the paraId of its first paragraph
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w15:commentsEx", {
    "xmlns:w": NS_W,
    "xmlns:w15": NS_W15,
    "mc:Ignorable": "w15",
    "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"
  });

  for (const c of extended) {
    const firstPara = c.content[0];
    if (!firstPara?.paraId) {
      continue;
    }
    const attrs: Record<string, string> = { "w15:paraId": firstPara.paraId };
    if (c.done) {
      attrs["w15:done"] = "1";
    }
    if (c.parentId) {
      attrs["w15:paraIdParent"] = c.parentId;
    }
    xml.leafNode("w15:commentEx", attrs);
  }

  xml.closeNode();
  return true;
}
