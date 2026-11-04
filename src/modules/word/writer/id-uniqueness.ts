/**
 * Write-time guard: no two entries of a referenced collection share an id.
 *
 * Content refers to footnotes, endnotes, comments, numbering definitions and
 * images by id alone, so a duplicate cannot be repaired here without guessing
 * which entry each reference meant. Both writers call this before they emit
 * anything, which is what makes an edit the `Document.*` helpers could not see
 * fail loudly instead of reaching the package.
 */

import { imageKeys, isSamePicture } from "@word/core/image-keys";
import { DocxDuplicateIdError } from "@word/errors";
import type { DocxIdCollection } from "@word/errors";
import type {
  AbstractNumbering,
  CommentDef,
  EndnoteDef,
  FootnoteDef,
  ImageDef,
  NumberingInstance
} from "@word/types";

/** The collections checked; a `DocxDocument` and the streaming options both have this shape. */
export interface ReferencedIdCollections {
  readonly footnotes?: readonly FootnoteDef[];
  readonly endnotes?: readonly EndnoteDef[];
  readonly comments?: readonly CommentDef[];
  readonly abstractNumberings?: readonly AbstractNumbering[];
  readonly numberingInstances?: readonly NumberingInstance[];
  readonly images?: readonly ImageDef[];
}

function assertUnique<T>(
  collection: DocxIdCollection,
  list: readonly T[] | undefined,
  pick: (item: T) => number | string | undefined,
  reserved?: readonly number[]
): void {
  if (!list || list.length === 0) {
    return;
  }
  const seen = new Set<number | string>(reserved);
  for (const item of list) {
    const id = pick(item);
    if (id === undefined) {
      continue;
    }
    if (seen.has(id)) {
      throw new DocxDuplicateIdError(
        collection,
        id,
        typeof id === "number" && reserved?.includes(id)
          ? "the writer adds the separator notes -1 and 0"
          : undefined
      );
    }
    seen.add(id);
  }
}

/**
 * Ids the note writer emits itself: when no entry is a separator it adds the
 * separator (-1) and continuation separator (0).
 */
function reservedNoteIds(notes: readonly (FootnoteDef | EndnoteDef)[] | undefined): number[] {
  const hasSeparators = notes?.some(
    n => n.type === "separator" || n.type === "continuationSeparator"
  );
  return hasSeparators ? [] : [-1, 0];
}

/**
 * The first image key two different pictures share. Keys are `rId` and every
 * `aliasRIds` entry: content in any part names an image by any of them. Two
 * entries for the same bytes under one key are not a conflict — whichever a
 * reference resolves to, it shows the same picture.
 */
function assertUniqueImageKeys(images: readonly ImageDef[] | undefined): void {
  if (!images || images.length < 2) {
    return;
  }
  const owners = new Map<string, ImageDef>();
  for (const img of images) {
    for (const key of imageKeys(img)) {
      const owner = owners.get(key);
      if (owner === undefined) {
        owners.set(key, img);
      } else if (!isSamePicture(owner, img)) {
        throw new DocxDuplicateIdError("images", key, "two different pictures");
      }
    }
  }
}

/**
 * Throw {@link DocxDuplicateIdError} for the first id two entries share.
 *
 * Image keys are document-wide: a relationship id is local to a part, but
 * the reader gives every picture a key no other picture in any part uses, so
 * content anywhere names exactly one image. Duplicate file names need no
 * check: nothing refers to an image by file name, and the packager already
 * renames the second.
 *
 * Linear in the size of the collections.
 */
export function assertUniqueReferencedIds(model: ReferencedIdCollections): void {
  assertUnique("footnotes", model.footnotes, f => f.id, reservedNoteIds(model.footnotes));
  assertUnique("endnotes", model.endnotes, e => e.id, reservedNoteIds(model.endnotes));
  assertUnique("comments", model.comments, c => c.id);
  assertUnique("abstractNumberings", model.abstractNumberings, a => a.abstractNumId);
  assertUnique("numberingInstances", model.numberingInstances, n => n.numId);
  assertUniqueImageKeys(model.images);
}
