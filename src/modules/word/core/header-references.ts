/**
 * Which headers and footers a section shows, as both the packager and the
 * layout read it — one rule, so a document cannot look one way on screen and
 * another in the file.
 *
 * A section names its parts by explicit references (`sectionProperties.headers`
 * / `footers`, `{ type, rId }`); `Document.setHeader` / `setFooter` always
 * write one. A model built by hand may instead store a part under its type
 * (`doc.headers.get("default")`) and name none: the document's last section
 * then shows each such type-keyed part.
 */

import type { FooterDef, HeaderDef, HeaderFooterRef, HeaderFooterType } from "@word/types";

/** The header and footer types a section references (`ST_HdrFtr`). */
export const HEADER_FOOTER_TYPES: ReadonlySet<string> = new Set<string>([
  "default",
  "first",
  "even"
] satisfies HeaderFooterType[]);

/**
 * The references a document's last section shows when it names none: one per
 * part stored under a type key, in the parts' order.
 */
export function impliedReferences(
  parts: ReadonlyMap<string, HeaderDef | FooterDef> | undefined
): HeaderFooterRef[] {
  const refs: HeaderFooterRef[] = [];
  for (const key of parts?.keys() ?? []) {
    if (HEADER_FOOTER_TYPES.has(key)) {
      refs.push({ type: key as HeaderFooterType, rId: key });
    }
  }
  return refs;
}

/**
 * The key in `parts` a reference names: the key itself, the part whose `rId`
 * it carries, or — for a reference with no resolvable id — the part stored
 * under its type. `undefined` when it names none, and the writer drops it.
 */
export function referencedPartKey(
  ref: HeaderFooterRef,
  parts: ReadonlyMap<string, HeaderDef | FooterDef> | undefined
): string | undefined {
  if (!parts) {
    return undefined;
  }
  if (ref.rId && parts.has(ref.rId)) {
    return ref.rId;
  }
  if (ref.rId) {
    for (const [key, part] of parts) {
      if (part.rId === ref.rId) {
        return key;
      }
    }
  }
  return parts.has(ref.type) ? ref.type : undefined;
}

/**
 * A section's references with each id replaced by the key of the part it
 * names (unresolvable ones dropped). The document's `last` section, naming
 * none, shows the {@link impliedReferences}.
 */
export function sectionReferences(
  own: readonly HeaderFooterRef[] | undefined,
  parts: ReadonlyMap<string, HeaderDef | FooterDef> | undefined,
  last: boolean
): HeaderFooterRef[] {
  const resolved: HeaderFooterRef[] = [];
  for (const ref of own ?? []) {
    const key = referencedPartKey(ref, parts);
    if (key !== undefined) {
      resolved.push(key === ref.rId ? ref : { ...ref, rId: key });
    }
  }
  if (resolved.length > 0 || !last) {
    return resolved;
  }
  return impliedReferences(parts);
}
