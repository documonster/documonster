/**
 * Image relationships of one part.
 *
 * The model's image keys (`ImageDef.rId` and `aliasRIds`) are unique across
 * the document, and content in every part names images by key. A package's
 * relationship ids, though, are local to each part. So every part that shows
 * a picture gets its own image relationships, registered here, and a remap
 * from key to the id that part's `.rels` actually uses; the part's renderer
 * writes `r:embed` through that remap.
 *
 * Used by the buffered packager and the streaming writer alike, so the two
 * cannot drift in which references a part registers. The keys a part refers
 * to come from `collectImageKeys`.
 */

import { RelType } from "@word/constants";
import type { ImageDef } from "@word/types";
import type { RelationshipsState } from "@word/writer/relationships";

const CANONICAL_ID = /^rId\d+$/;

/**
 * Register a part's image relationships and return the remap its renderer
 * needs: key → the id this part's `.rels` gives the image, for every key
 * whose id differs. Modelled references and markup kept verbatim are both
 * written through it.
 *
 * A key is registered under its own name when it has the `rId<n>` form Word
 * requires and the id is still free in this part — so a document that was
 * read keeps its ids — and under a freshly allocated id otherwise.
 *
 * @param targetOf - The relationship target for an image, relative to the part.
 * @param secondary - Extra relationships an image brings with it — the SVG of
 *   an SVG picture, whose primary relationship is its raster fallback — keyed
 *   by the image, each with the id content uses for it.
 */
export function registerPartImages(
  rels: RelationshipsState,
  keys: Iterable<string>,
  index: ReadonlyMap<string, ImageDef>,
  targetOf: (img: ImageDef) => string,
  secondary?: ReadonlyMap<ImageDef, { readonly id: string; readonly target: string }>
): Map<string, string> {
  const remap = new Map<string, string>();
  const claim = (key: string, target: string): void => {
    if (rels.hasId(key) && rels.findByTypeAndTarget(RelType.Image, target)?.id === key) {
      return;
    }
    if (!rels.hasId(key) && CANONICAL_ID.test(key)) {
      rels.addWithId(key, RelType.Image, target);
      return;
    }
    const id = rels.add(RelType.Image, target);
    if (id !== key) {
      remap.set(key, id);
    }
  };
  for (const key of keys) {
    const img = index.get(key);
    if (!img) {
      continue;
    }
    claim(key, targetOf(img));
    const extra = secondary?.get(img);
    if (extra) {
      claim(extra.id, extra.target);
    }
  }
  return remap;
}
