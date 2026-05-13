/**
 * DOCX reader — comments parsers.
 *
 * Split out from `docx-reader.ts` to keep that file's part-orchestration
 * scope manageable. We accept `parseParagraph` as an injected dependency
 * because the body paragraph parser still lives in `docx-reader.ts` (it
 * has many cross-references that would not be cheap to disentangle); a
 * direct import would form a cycle.
 */

import { parseXml } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import type { Mutable } from "../core/internal-utils";
import type { CommentDef, Paragraph } from "../types";
import { attrInt, attrVal, findChildrenNs } from "./parse-utils";
import { type ReaderContext, createFieldState } from "./reader-context";

/** Parse `word/comments.xml` into a list of CommentDef. */
export function parseCommentsXml(
  xmlStr: string,
  ctx: ReaderContext,
  parseParagraph: (el: XmlElement, ctx: ReaderContext) => Paragraph
): CommentDef[] {
  // Comments are a separate part — reset the field state so a body-level
  // unterminated field does not bleed into comment parsing.
  const savedField = ctx.field;
  ctx.field = createFieldState();
  try {
    const doc = parseXml(xmlStr);
    const root = doc.root;
    const comments: CommentDef[] = [];

    for (const commentEl of findChildrenNs(root, "comment")) {
      const id = attrInt(commentEl, "id");
      const author = attrVal(commentEl, "author");
      if (id === undefined || !author) {
        continue;
      }

      const content: Paragraph[] = [];
      for (const child of commentEl.children) {
        if (child.type === "element" && child.name.replace(/^w:/, "") === "p") {
          content.push(parseParagraph(child, ctx));
        }
      }

      const comment: Mutable<CommentDef> = { id, author, content };
      const date = attrVal(commentEl, "date");
      if (date) {
        comment.date = date;
      }
      const initials = attrVal(commentEl, "initials");
      if (initials) {
        comment.initials = initials;
      }
      comments.push(comment);
    }

    return comments;
  } finally {
    ctx.field = savedField;
  }
}

/**
 * Parse `word/commentsExtended.xml` into a paraId → metadata map.
 *
 * Modern Word stores comment "resolved" state and reply threading in a
 * sidecar part keyed by the paragraph id of each comment's first
 * paragraph. The reader merges this map into `CommentDef.done` /
 * `CommentDef.parentId` once both parts have been parsed.
 */
export function parseCommentsExtendedXml(
  xmlStr: string
): Map<string, { done?: boolean; parentId?: string }> {
  const map = new Map<string, { done?: boolean; parentId?: string }>();
  const doc = parseXml(xmlStr);
  const root = doc.root;
  for (const child of root.children) {
    if (child.type !== "element") {
      continue;
    }
    // w15:commentEx
    const name = child.name;
    if (!name.endsWith("commentEx")) {
      continue;
    }
    const paraId = child.attributes["w15:paraId"] ?? child.attributes["paraId"];
    if (!paraId) {
      continue;
    }
    const entry: { done?: boolean; parentId?: string } = {};
    const done = child.attributes["w15:done"] ?? child.attributes["done"];
    if (done === "1" || done === "true") {
      entry.done = true;
    } else if (done === "0" || done === "false") {
      entry.done = false;
    }
    const pid = child.attributes["w15:paraIdParent"] ?? child.attributes["paraIdParent"];
    if (pid) {
      entry.parentId = pid;
    }
    map.set(paraId, entry);
  }
  return map;
}
