/**
 * Reader/writer for Office 365 "threaded comments" (`<ThreadedComments>`
 * XML root) and the workbook-level person directory (`<personList>`).
 *
 * Threaded comments live in a separate part tree from classic VML
 * comments:
 *
 *   - `xl/threadedComments/threadedComment{N}.xml` — one per sheet
 *     that has threaded comments; referenced from the sheet rels
 *   - `xl/persons/person.xml` — workbook-level person directory;
 *     referenced from the workbook rels
 *
 * The schema lives in Microsoft's extension namespace
 * `http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments`
 * (for comments) and the older `…/2018/threadedcomments` for persons
 * (identical URI tag-shared).
 *
 * This module is intentionally free of runtime xform plumbing —
 * threaded comments are a discrete part with only text content, so
 * two small render/parse functions suffice and plug directly into
 * `xlsx.browser.ts` and `workbook.browser.ts`.
 */

import type { ThreadedComment, ThreadedCommentMention, ThreadedCommentPerson } from "@excel/types";
import { synthGuid } from "@excel/utils/guid";
import { findChild, findChildren, parseXml, textContent } from "@xml/dom";

const TC_NAMESPACE = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

/**
 * Render the threaded-comment part for a single worksheet.
 *
 * Input is a list of `(cellRef, comment)` pairs — the cell ref is
 * stored as the `ref` attribute on each `<threadedComment>`, not on
 * the public {@link ThreadedComment} type (which doesn't carry it so
 * the same structure can be shared between the public API and the
 * XForm).
 *
 * Input order is preserved — Excel relies on parent-child ordering
 * when rendering conversation threads, so callers must keep replies
 * after their parents.
 */
export function renderThreadedComments(
  entries: Array<{ ref: string; comment: ThreadedComment }>
): string {
  const parts: string[] = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  parts.push(`<ThreadedComments xmlns="${TC_NAMESPACE}">`);
  for (const { ref, comment: c } of entries) {
    const id = c.id ?? `{${synthGuid()}}`;
    const attrs: string[] = [`ref="${escapeAttr(ref)}"`];
    if (c.date) {
      attrs.push(`dT="${escapeAttr(c.date)}"`);
    }
    attrs.push(`personId="${escapeAttr(c.personId)}"`);
    attrs.push(`id="${escapeAttr(id)}"`);
    if (c.parentId) {
      attrs.push(`parentId="${escapeAttr(c.parentId)}"`);
    }
    if (c.done !== undefined) {
      attrs.push(`done="${c.done ? "1" : "0"}"`);
    }
    parts.push(`<threadedComment ${attrs.join(" ")}>`);
    parts.push(`<text>${escapeXml(c.text)}</text>`);
    if (c.mentions && c.mentions.length > 0) {
      parts.push("<mentions>");
      for (const m of c.mentions) {
        const mAttrs: string[] = [
          `mentionpersonId="${escapeAttr(m.mentionPersonId)}"`,
          `mentionId="${escapeAttr(m.mentionId ?? `{${synthGuid()}}`)}"`,
          `startIndex="${m.startIndex}"`,
          `length="${m.length}"`
        ];
        parts.push(`<mention ${mAttrs.join(" ")}/>`);
      }
      parts.push("</mentions>");
    }
    parts.push(`</threadedComment>`);
  }
  parts.push(`</ThreadedComments>`);
  return parts.join("");
}

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
 * Render the workbook-level `xl/persons/person.xml` part.
 */
export function renderPersonList(persons: ThreadedCommentPerson[]): string {
  const parts: string[] = [];
  parts.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  parts.push(`<personList xmlns="${TC_NAMESPACE}">`);
  for (const p of persons) {
    const attrs: string[] = [
      `displayName="${escapeAttr(p.displayName)}"`,
      `id="${escapeAttr(p.id)}"`
    ];
    if (p.userId !== undefined) {
      attrs.push(`userId="${escapeAttr(p.userId)}"`);
    }
    if (p.providerId !== undefined) {
      attrs.push(`providerId="${escapeAttr(p.providerId)}"`);
    }
    parts.push(`<person ${attrs.join(" ")}/>`);
  }
  parts.push(`</personList>`);
  return parts.join("");
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

function escapeXml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escapeAttr(value: string): string {
  return escapeXml(value).replace(/"/g, "&quot;");
}
