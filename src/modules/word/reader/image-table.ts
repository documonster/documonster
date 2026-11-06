/**
 * The document-wide image table, built as each part is read.
 *
 * A relationship id belongs to the part whose `.rels` declares it: two
 * headers may each call their own logo `rId1`, and a header may use `rId5`
 * for a picture while the body uses `rId5` for another. The model has one
 * image table for the whole document, so its keys (`ImageDef.rId` and
 * `aliasRIds`) must be unique across all parts, and every part's content
 * must name images by those keys.
 *
 * The body keeps its own ids, so the common document reads back exactly as
 * it was written. Another part keeps its local id too when that id is free —
 * not the key of a different picture, and not any relationship of the body,
 * whose markup is kept verbatim in places and must not come to mean
 * something else. Otherwise the picture gets a fresh key, and the part's
 * markup is rewritten to it before it is parsed, so structured content and
 * markup kept verbatim agree.
 */

import { RelType } from "@word/constants";
import { mediaTypeForExtension } from "@word/core/image-media";
import { getFileExt, getFileName } from "@word/core/opc-paths";
import { rebindRelationshipIds } from "@word/core/relationship-ids";
import { resolvePartPath } from "@word/reader/parse-utils";
import type { ParsedRelationship } from "@word/reader/reader-context";
import type { ImageDef, ImageMediaType } from "@word/types";

interface MutableImage {
  data: Uint8Array;
  mediaType: ImageMediaType;
  fileName: string;
  rId: string;
  aliasRIds?: string[];
  contentType?: string;
}

export interface ImageTable {
  /** The images collected so far, body first, in the order parts were read. */
  readonly images: readonly ImageDef[];
  /**
   * Register the body's image relationships. Their ids become keys as they
   * are; one image per relationship.
   */
  bindDocument(partPath: string, rels: readonly ParsedRelationship[]): void;
  /**
   * Register another part's image relationships and return its markup with
   * every image reference expressed as a document-wide key.
   */
  bindPart(partPath: string, rels: readonly ParsedRelationship[], xml: string): string;
}

/**
 * @param entries - The package's parts.
 * @param allRelationshipIds - Every relationship id any `.rels` in the
 *   package declares; a fresh key avoids all of them.
 * @param consumed - Receives the path of every media part the table takes.
 */
export function createImageTable(
  entries: ReadonlyMap<string, Uint8Array>,
  allRelationshipIds: ReadonlySet<string>,
  consumed: Set<string>,
  declaredContentType: (path: string) => string | undefined
): ImageTable {
  const images: MutableImage[] = [];
  const byPath = new Map<string, MutableImage>();
  const owner = new Map<string, MutableImage>();
  const pathOf = new Map<MutableImage, string>();
  let documentIds: ReadonlySet<string> = new Set();
  let counter = 0;
  for (const id of allRelationshipIds) {
    const m = /^rId(\d+)$/.exec(id);
    if (m) {
      counter = Math.max(counter, Number(m[1]));
    }
  }

  function freshKey(): string {
    let key: string;
    do {
      key = `rId${++counter}`;
    } while (allRelationshipIds.has(key) || owner.has(key));
    return key;
  }

  function media(partPath: string, rel: ParsedRelationship): [string, Uint8Array] | undefined {
    if (rel.type !== RelType.Image || rel.targetMode === "External") {
      return undefined;
    }
    const path = resolvePartPath(partPath, rel.target);
    const data = entries.get(path);
    return data ? [path, data] : undefined;
  }

  function add(path: string, data: Uint8Array, key: string): MutableImage {
    const fileName = getFileName(path);
    const ext = getFileExt(fileName);
    const known = mediaTypeForExtension(ext);
    // An extension outside `ImageMediaType` keeps the type its package declared,
    // so the part is written back under its own name with that type.
    const contentType = known ? undefined : declaredContentType(path);
    const img: MutableImage = {
      data,
      mediaType: known ?? ((ext || "png") as ImageMediaType),
      fileName,
      rId: key,
      ...(contentType ? { contentType } : {})
    };
    images.push(img);
    owner.set(key, img);
    pathOf.set(img, path);
    if (!byPath.has(path)) {
      byPath.set(path, img);
    }
    consumed.add(path);
    return img;
  }

  function alias(img: MutableImage, key: string): void {
    (img.aliasRIds ??= []).push(key);
    owner.set(key, img);
  }

  return {
    images: images as readonly ImageDef[],

    bindDocument(partPath, rels) {
      documentIds = new Set(rels.map(r => r.id));
      for (const rel of rels) {
        const found = media(partPath, rel);
        if (found && !owner.has(rel.id)) {
          add(found[0], found[1], rel.id);
        }
      }
    },

    bindPart(partPath, rels, xml) {
      const remap = new Map<string, string>();
      for (const rel of rels) {
        const found = media(partPath, rel);
        if (!found) {
          continue;
        }
        const [path, data] = found;
        const local = rel.id;
        const img = byPath.get(path);
        const holder = owner.get(local);
        if (holder && pathOf.get(holder) === path) {
          continue;
        }
        if (holder === undefined && !documentIds.has(local)) {
          if (img) {
            alias(img, local);
          } else {
            add(path, data, local);
          }
          continue;
        }
        const key = freshKey();
        if (img) {
          alias(img, key);
        } else {
          add(path, data, key);
        }
        remap.set(local, key);
      }
      return rebindRelationshipIds(xml, remap);
    }
  };
}
