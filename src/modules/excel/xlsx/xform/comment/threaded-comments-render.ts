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
 * **This file renders; `threaded-comments-parse.ts` reads.** They were one module, and a module both
 * the writer and the reader import is kept whole by a bundler that splits chunks per module (rspack): the
 * write-only bundle carried `@xml/dom` and the SAX parser behind it for the two parse functions.
 */

import type { ThreadedComment, ThreadedCommentPerson } from "@excel/types";
import { synthGuid } from "@excel/utils/guid";
import { xmlEncode, xmlEncodeAttr } from "@xml/encode";

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
        // `startIndex` and `length` are `xsd:unsignedInt` per the
        // threaded-comments schema — negative, non-integer, or
        // non-finite values produce XML that strict validators (and
        // Excel's own reader) reject. Fail loud with a descriptive
        // error rather than ship a broken comment part.
        if (!isNonNegativeInt(m.startIndex)) {
          throw new Error(
            `Threaded comment mention.startIndex must be a non-negative integer; got ${m.startIndex}.`
          );
        }
        if (!isNonNegativeInt(m.length)) {
          throw new Error(
            `Threaded comment mention.length must be a non-negative integer; got ${m.length}.`
          );
        }
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

function escapeXml(value: string): string {
  // Route through the canonical encoder so the writer strips
  // XML-illegal control characters / lone surrogates in user-supplied
  // comment text. The previous manual `.replace` chain only handled
  // `& < >` and left every other hazard untouched — e.g. a pasted
  // `\u0008` corrupted the whole `<ThreadedComments>` part on save.
  return xmlEncode(value);
}

function escapeAttr(value: string): string {
  // Attribute values additionally need `\t \n \r` encoded as
  // numeric character references so XML attribute-value normalisation
  // doesn't collapse them to literal spaces; a threaded comment
  // `personId` / `id` GUID shouldn't ever contain whitespace but the
  // `ref` / `displayName` attributes can carry user input.
  return xmlEncodeAttr(value);
}

/**
 * Whether `v` is a finite non-negative integer suitable for the
 * `xsd:unsignedInt` attributes used by the threaded-comments schema.
 * Rejects `NaN`, `±Infinity`, negative numbers, and fractional values.
 */
function isNonNegativeInt(v: unknown): v is number {
  return typeof v === "number" && Number.isFinite(v) && v >= 0 && Math.floor(v) === v;
}
