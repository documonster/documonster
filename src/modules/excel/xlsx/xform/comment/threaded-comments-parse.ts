/**
 * Office 365 "threaded comments" (`<ThreadedComments>` XML root) and the workbook-level person directory
 * (`<personList>`).
 *
 * Threaded comments live in a separate part tree from classic VML comments:
 *
 *   - `xl/threadedComments/threadedComment{N}.xml` — one per sheet that has threaded comments; referenced
 *     from the sheet rels
 *   - `xl/persons/person.xml` — workbook-level person directory; referenced from the workbook rels
 *
 * The schema lives in Microsoft's extension namespace
 * `http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments`.
 *
 * **This file reads; `threaded-comments-render.ts` writes** — see there for why they are separate.
 */

import type { ThreadedComment, ThreadedCommentMention, ThreadedCommentPerson } from "@excel/types";
import { synthGuid } from "@excel/utils/guid";
import { findChild, findChildren, parseXml, textContent } from "@xml/dom";

/**
 * Parse a `xl/threadedComments/threadedComment{N}.xml` payload into
 * structured per-cell entries. Returns an empty array on malformed
 * input — the caller can silently drop threaded comments rather than
 * failing the entire workbook load.
 */
export function parseThreadedComments(
  rawXml: string
): Array<{ ref: string; comment: ThreadedComment }> {
  let root;
  try {
    root = parseXml(rawXml).root;
  } catch {
    return [];
  }
  const result: Array<{ ref: string; comment: ThreadedComment }> = [];
  for (const el of findChildren(root, "threadedComment")) {
    const ref = el.attributes.ref;
    if (!ref) {
      continue;
    }
    const personId = el.attributes.personId;
    if (!personId) {
      continue;
    }
    const textEl = findChild(el, "text");
    const mentionsEl = findChild(el, "mentions");
    const mentions: ThreadedCommentMention[] = mentionsEl
      ? findChildren(mentionsEl, "mention").map(m => ({
          mentionId: m.attributes.mentionId,
          mentionPersonId: m.attributes.mentionpersonId ?? "",
          startIndex: parseInt(m.attributes.startIndex ?? "0", 10),
          length: parseInt(m.attributes.length ?? "0", 10)
        }))
      : [];
    const comment: ThreadedComment = {
      personId,
      text: textEl ? textContent(textEl) : "",
      ...(el.attributes.id ? { id: el.attributes.id } : {}),
      ...(el.attributes.parentId ? { parentId: el.attributes.parentId } : {}),
      ...(el.attributes.dT ? { date: el.attributes.dT } : {}),
      ...(el.attributes.done !== undefined ? { done: el.attributes.done === "1" } : {}),
      ...(mentions.length > 0 ? { mentions } : {})
    };
    result.push({ ref, comment });
  }
  return result;
}

/**
 * Parse `xl/persons/person.xml` into a {@link ThreadedCommentPerson}
 * list. Missing ids are auto-generated so downstream parts that
 * reference the list don't accidentally collide.
 */
export function parsePersonList(rawXml: string): ThreadedCommentPerson[] {
  let root;
  try {
    root = parseXml(rawXml).root;
  } catch {
    return [];
  }
  return findChildren(root, "person")
    .map(el => {
      const id = el.attributes.id ?? `{${synthGuid()}}`;
      const displayName = el.attributes.displayName ?? "";
      if (!displayName) {
        return undefined;
      }
      const entry: ThreadedCommentPerson = { id, displayName };
      if (el.attributes.userId !== undefined) {
        entry.userId = el.attributes.userId;
      }
      if (el.attributes.providerId !== undefined) {
        entry.providerId = el.attributes.providerId;
      }
      return entry;
    })
    .filter((x): x is ThreadedCommentPerson => x !== undefined);
}
