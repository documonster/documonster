/**
 * Image keys.
 *
 * The model's image table is shared by every part of a document, and content
 * anywhere names a picture by one of its keys: `ImageDef.rId` or an entry of
 * `ImageDef.aliasRIds`. Keys are unique across the table — two different
 * pictures never share one.
 */

import { findRelationshipIdReferences, rebindRelationshipIds } from "@word/core/relationship-ids";
import type { ImageDef } from "@word/types";

/** Every key an image answers to: its `rId`, then its `aliasRIds`. */
export function imageKeys(img: ImageDef): string[] {
  const keys = img.rId !== undefined ? [img.rId] : [];
  for (const alias of img.aliasRIds ?? []) {
    if (!keys.includes(alias)) {
      keys.push(alias);
    }
  }
  return keys;
}

/**
 * Whether two entries are the same picture — the same bytes in the same
 * format — so a reference resolving to either shows the same thing.
 */
export function isSamePicture(a: ImageDef, b: ImageDef): boolean {
  if (a === b || (a.data === b.data && a.mediaType === b.mediaType)) {
    return true;
  }
  if (a.mediaType !== b.mediaType || a.data.length !== b.data.length) {
    return false;
  }
  for (let i = 0; i < a.data.length; i++) {
    if (a.data[i] !== b.data[i]) {
      return false;
    }
  }
  return true;
}

/** Node types whose `rId` / `svgRId` name an image (an image watermark is `"image"` too). */
const IMAGE_NODE_TYPES = new Set(["image", "floatingImage"]);

/** Every image of the table by each of its keys: `rId` and `aliasRIds`. */
export function indexImagesByKey(
  images: readonly ImageDef[] | undefined
): ReadonlyMap<string, ImageDef> {
  const index = new Map<string, ImageDef>();
  for (const img of images ?? []) {
    if (img.rId && !index.has(img.rId)) {
      index.set(img.rId, img);
    }
  }
  for (const img of images ?? []) {
    for (const alias of img.aliasRIds ?? []) {
      if (!index.has(alias)) {
        index.set(alias, img);
      }
    }
  }
  return index;
}

/**
 * Whether a string field holds markup the writer emits verbatim — `rawXml`,
 * `rawVmlXml`, a shape's `_advancedFillXml`. Its relationship ids are
 * rewritten through the part's remap as it is written, like modelled ones.
 */
function isMarkupField(field: string): boolean {
  return field.endsWith("Xml");
}

/**
 * Add the image keys referenced anywhere under `root` — content blocks, a
 * note or comment list, a header's content, picture bullets — to `out`.
 * Only keys present in `index` are recorded: markup also refers to
 * hyperlinks, charts and other relationships, which are not images.
 */
export function collectImageKeys(
  root: unknown,
  index: ReadonlyMap<string, ImageDef>,
  out: Set<string>
): void {
  visitImageReferences(root, index, key => out.add(key));
}

/**
 * The keys under `root` that a drawing names only as its `svgRId` — the SVG
 * side of a picture whose raster side is the drawing's `rId` — and nothing
 * names as a picture to draw. Such an image is the SVG of a pair the content
 * already holds both halves of, as a read document has it, so it needs no
 * raster fallback of its own.
 */
export function collectSvgPartnerKeys(
  root: unknown,
  index: ReadonlyMap<string, ImageDef>
): Set<string> {
  const svg = new Set<string>();
  const drawn = new Set<string>();
  visitImageReferences(root, index, (key, field) => (field === "svgRId" ? svg : drawn).add(key));
  for (const key of drawn) {
    svg.delete(key);
  }
  return svg;
}

/** How a reference names an image: a drawing's `rId` or `svgRId`, or an id in markup. */
type ImageReferenceField = "rId" | "svgRId" | "markup";

function visitImageReferences(
  root: unknown,
  index: ReadonlyMap<string, ImageDef>,
  visit: (key: string, field: ImageReferenceField) => void
): void {
  if (index.size === 0) {
    return;
  }
  const stack: unknown[] = [root];
  const seen = new Set<object>();
  while (stack.length > 0) {
    const node = stack.pop();
    if (node === null || typeof node !== "object" || ArrayBuffer.isView(node) || seen.has(node)) {
      continue;
    }
    // A model may share or even cycle through objects; visit each once.
    seen.add(node);
    if (Array.isArray(node)) {
      for (const item of node) {
        if (item !== null && typeof item === "object") {
          stack.push(item);
        }
      }
      continue;
    }
    if (node instanceof Map) {
      for (const value of node.values()) {
        stack.push(value);
      }
      continue;
    }
    const record = node as Record<string, unknown>;
    if (typeof record.type === "string" && IMAGE_NODE_TYPES.has(record.type)) {
      for (const field of ["rId", "svgRId"] as const) {
        const key = record[field];
        if (typeof key === "string" && index.has(key)) {
          visit(key, field);
        }
      }
    }
    for (const field in record) {
      const value = record[field];
      if (typeof value === "string") {
        if (isMarkupField(field)) {
          for (const key of findRelationshipIdReferences(value)) {
            if (index.has(key)) {
              visit(key, "markup");
            }
          }
        }
      } else if (value !== null && typeof value === "object") {
        stack.push(value);
      }
    }
  }
}

function isPlainObject(value: object): boolean {
  const proto = Object.getPrototypeOf(value) as unknown;
  return proto === Object.prototype || proto === null;
}

/** What a rewrite renames: image keys, and optionally anything else in markup. */
interface Rewrite {
  readonly images: ReadonlyMap<string, string>;
  readonly markup?: (xml: string) => string;
}

function rewrite(node: unknown, remap: Rewrite, done: Map<object, unknown>): unknown {
  if (node === null || typeof node !== "object" || ArrayBuffer.isView(node)) {
    return node;
  }
  // A subtree reached twice is rewritten once, so references that were
  // shared stay shared.
  if (done.has(node)) {
    return done.get(node);
  }
  // A cycle back to a node being rewritten sees the node as it is.
  done.set(node, node);
  const result = rewriteNode(node, remap, done);
  done.set(node, result);
  return result;
}

function rewriteNode(node: object, remap: Rewrite, done: Map<object, unknown>): unknown {
  if (Array.isArray(node)) {
    let out: unknown[] | undefined;
    for (let i = 0; i < node.length; i++) {
      const next = rewrite(node[i], remap, done);
      if (next !== node[i]) {
        (out ??= node.slice())[i] = next;
      }
    }
    return out ?? node;
  }
  if (node instanceof Map) {
    let out: Map<unknown, unknown> | undefined;
    for (const [key, value] of node) {
      const next = rewrite(value, remap, done);
      if (next !== value) {
        (out ??= new Map(node)).set(key, next);
      }
    }
    return out ?? node;
  }
  if (!isPlainObject(node)) {
    return node;
  }
  const record = node as Record<string, unknown>;
  const isImage = typeof record.type === "string" && IMAGE_NODE_TYPES.has(record.type);
  let out: Record<string, unknown> | undefined;
  for (const field in record) {
    const value = record[field];
    let next = value;
    if (typeof value === "string") {
      if (isImage && (field === "rId" || field === "svgRId")) {
        next = remap.images.get(value) ?? value;
      } else if (field.endsWith("Xml")) {
        next = rebindRelationshipIds(value, remap.images);
        if (remap.markup) {
          next = remap.markup(next as string);
        }
      }
    } else if (value !== null && typeof value === "object") {
      next = rewrite(value, remap, done);
    }
    if (next !== value) {
      (out ??= { ...record })[field] = next;
    }
  }
  return out ?? node;
}

/**
 * `node` with every image reference found in `remap` renamed: the `rId` and
 * `svgRId` of inline and floating images, and the relationship ids of markup
 * kept verbatim (`rawXml`, `rawVmlXml`, a shape's `_advancedFillXml`, …).
 * Copy-on-write: unchanged subtrees, and `node` itself when nothing changes,
 * are returned as they are, and the input is never mutated.
 */
export function rewriteImageReferences<T>(node: T, remap: ReadonlyMap<string, string>): T {
  return remap.size === 0 ? node : (rewrite(node, { images: remap }, new Map()) as T);
}

/**
 * As {@link rewriteImageReferences}, and every string of markup kept verbatim
 * (a field whose name ends in `Xml`) also passed through `markup` — for ids
 * other than image keys that such markup carries.
 */
export function rewriteMarkupReferences<T>(
  node: T,
  images: ReadonlyMap<string, string>,
  markup: (xml: string) => string
): T {
  return rewrite(node, { images, markup }, new Map()) as T;
}
