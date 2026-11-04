/**
 * DOCX Writers - which footnote/endnote ids the separator notes are written with.
 *
 * `settings.xml` names the separator notes by id (`w:footnotePr/w:footnote`),
 * so its references have to follow the ids `footnotes.xml`/`endnotes.xml`
 * actually get — which are not the source's: the reader drops the separators
 * and the writer synthesises them unless the model carries its own.
 */

import type { NoteSeparatorType } from "@word/types";

/** The ids the writer gives synthesised separators. */
export const DEFAULT_SEPARATOR_ID = -1;
export const DEFAULT_CONTINUATION_SEPARATOR_ID = 0;

/** Whether the notes carry their own separators, so none are synthesised. */
export function hasOwnSeparators(notes: readonly { readonly type?: string }[]): boolean {
  return notes.some(n => n.type === "separator" || n.type === "continuationSeparator");
}

/**
 * The ids to reference for `types`, given the notes written to the part. A
 * type no written note has is left out: a reference to a missing note is
 * invalid. No notes means no part, and no references.
 */
export function separatorReferenceIds(
  types: readonly NoteSeparatorType[] | undefined,
  notes: readonly { readonly id: number; readonly type?: string }[] | undefined
): number[] {
  if (!types || !notes || notes.length === 0) {
    return [];
  }
  const synthesised = !hasOwnSeparators(notes);
  const ids: number[] = [];
  for (const type of types) {
    const own = notes.find(n => n.type === type);
    if (own) {
      ids.push(own.id);
    } else if (synthesised && type === "separator") {
      ids.push(DEFAULT_SEPARATOR_ID);
    } else if (synthesised && type === "continuationSeparator") {
      ids.push(DEFAULT_CONTINUATION_SEPARATOR_ID);
    }
  }
  return ids;
}
