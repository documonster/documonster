/**
 * XML helpers that tolerate namespace prefixes.
 *
 * OOXML serialisers are allowed to write `<drawing>` or `<x:drawing>` —
 * both are legal in the same schema. All validator checks must therefore
 * match on the **local name**, not the prefixed name. These helpers also
 * tolerate namespaced attributes like `r:id` vs `ns0:id`.
 */

import { findChildren, parseXml, walk } from "@xml/dom";
import type { XmlDocument, XmlElement } from "@xml/types";

/**
 * Extract the local name from a tag name, i.e. the part after the last `:`.
 * For `drawing` returns `drawing`; for `x:drawing` returns `drawing`.
 */
export function localName(name: string): string {
  const idx = name.lastIndexOf(":");
  return idx === -1 ? name : name.slice(idx + 1);
}

/** Namespace-insensitive element-name equality. */
export function matchesLocal(name: string, target: string): boolean {
  return localName(name) === target;
}

/** Find first child element with matching **local** name. */
export function findChildLocal(el: XmlElement, local: string): XmlElement | undefined {
  for (const child of el.children) {
    if (child.type === "element" && matchesLocal(child.name, local)) {
      return child;
    }
  }
  return undefined;
}

/** Find all child elements with matching **local** name. */
export function findChildrenLocal(el: XmlElement, local: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const child of el.children) {
    if (child.type === "element" && matchesLocal(child.name, local)) {
      out.push(child);
    }
  }
  return out;
}

/** `true` when any descendant (or self) has matching local name. */
export function hasDescendantLocal(root: XmlElement, local: string): boolean {
  if (matchesLocal(root.name, local)) {
    return true;
  }
  let found = false;
  walk(root, child => {
    if (!found && matchesLocal(child.name, local)) {
      found = true;
    }
  });
  return found;
}

/** Collect all descendants (including self) matching local name. */
export function collectDescendantsLocal(root: XmlElement, local: string): XmlElement[] {
  const out: XmlElement[] = [];
  if (matchesLocal(root.name, local)) {
    out.push(root);
  }
  walk(root, child => {
    if (child !== root && matchesLocal(child.name, local)) {
      out.push(child);
    }
  });
  return out;
}

/**
 * Read an attribute by local name. The OOXML convention is to use
 * namespaced attributes like `r:id` but some serialisers strip or
 * renumber the prefix. We match by suffix `:<local>` or exact `local`.
 */
export function attrByLocalName(el: XmlElement, local: string): string | undefined {
  const direct = el.attributes[local];
  if (direct !== undefined) {
    return direct;
  }
  for (const key in el.attributes) {
    const idx = key.lastIndexOf(":");
    if (idx !== -1 && key.slice(idx + 1) === local) {
      return el.attributes[key];
    }
  }
  return undefined;
}

/**
 * Parse an XML string into a DOM tree. Returns `undefined` and records the
 * failure as a malformed XML problem via `onMalformed`. We don't throw
 * because validators must continue after a parse failure.
 */
export function tryParseXml(
  xml: string,
  onMalformed: (err: Error) => void
): XmlDocument | undefined {
  try {
    return parseXml(xml);
  } catch (err) {
    onMalformed(err instanceof Error ? err : new Error(String(err)));
    return undefined;
  }
}

// Re-export commonly used primitives so checkers only need one import.
export { findChildren };
