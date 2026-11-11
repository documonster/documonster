/**
 * XML → Plain Object — Shared Conversion Logic
 *
 * Internal module used by both `toPlainObject` (DOM path) and
 * `parseXmlToObject` (SAX-direct path). Not part of the public API.
 */

import { checkPlainObjectOptions, snapshotPlainObjectOptions } from "@xml/arguments";
import type { XmlToPlainObjectOptions } from "@xml/types";

// =============================================================================
// Resolved Options
// =============================================================================

/** Options with all defaults resolved — no more `??` checks at hot-path call sites. */
export interface ResolvedOptions {
  readonly ignoreAttributes: boolean;
  readonly attrPrefix: string;
  readonly textKey: string;
  readonly alwaysArray: boolean;
  readonly isArray: ((name: string) => boolean) | null;
  readonly preserveCData: boolean;
  readonly ignoreWS: boolean;
  /**
   * The key of each attribute name, prefixed: built once per name, since a fresh concatenation is
   * a new string to intern every time it is used as a key.
   */
  readonly attrKeys: Map<string, string>;
}

/** The key `name`'s attribute takes in the object: the prefix and the name. */
export function attributeKey(opts: ResolvedOptions, name: string): string {
  if (opts.attrPrefix === "") {
    return name;
  }
  let key = opts.attrKeys.get(name);
  if (key === undefined) {
    key = opts.attrPrefix + name;
    if (opts.attrKeys.size < 4096) {
      opts.attrKeys.set(name, key);
    }
  }
  return key;
}

export function resolveOptions(
  given?: XmlToPlainObjectOptions,
  where = "Xml.toPlainObject"
): ResolvedOptions {
  // Read once, so what is checked is what is used.
  const options = snapshotPlainObjectOptions(given);
  checkPlainObjectOptions(where, options);
  return {
    ignoreAttributes: options?.ignoreAttributes ?? false,
    attrPrefix: options?.attributePrefix ?? "@_",
    textKey: options?.textKey ?? "#text",
    alwaysArray: options?.alwaysArray ?? false,
    isArray: options?.isArray ?? null,
    preserveCData: options?.preserveCData ?? true,
    ignoreWS: options?.ignoreWhitespaceText ?? true,
    attrKeys: new Map()
  };
}

// =============================================================================
// Helpers
// =============================================================================

/** Check if a string contains only whitespace without allocating a trimmed copy. */
export function isWhitespaceOnly(s: string): boolean {
  for (let i = 0; i < s.length; i++) {
    const ch = s.charCodeAt(i);
    // space, tab, newline, carriage return
    if (ch !== 0x20 && ch !== 0x09 && ch !== 0x0a && ch !== 0x0d) {
      return false;
    }
  }
  return true;
}

// =============================================================================
// Value Resolution
// =============================================================================

/**
 * Determine the final plain-object value for an element given its parts.
 *
 * Returns either:
 * - a string (text-only / empty element)
 * - the `obj` record (element with attributes and/or children)
 */
export function resolveValue(
  obj: Record<string, unknown>,
  text: string,
  hasAttributes: boolean,
  hasChildren: boolean,
  opts: ResolvedOptions
): unknown {
  // Discard whitespace-only text when element has children (formatting indentation).
  if (opts.ignoreWS && hasChildren && text.length > 0 && isWhitespaceOnly(text)) {
    text = "";
  }
  const hasText = text.length > 0;

  // Text-only element with no attributes → collapse to string value
  if (hasText && !hasChildren && !hasAttributes) {
    return text;
  }

  // Add text content
  if (hasText) {
    addTextValue(obj, text, opts);
  }

  // Empty element with no attributes → empty string
  if (!hasAttributes && !hasChildren && !hasText) {
    return "";
  }

  return obj;
}

// =============================================================================
// Child Merging
// =============================================================================

/**
 * Add a resolved child value into a parent object, merging repeated names
 * into arrays.
 */
export function addChildValue(
  parent: Record<string, unknown>,
  name: string,
  value: unknown,
  alwaysArray: boolean,
  isArray: ((name: string) => boolean) | null
): void {
  const existing = parent[name];
  if (existing !== undefined) {
    if (Array.isArray(existing)) {
      existing.push(value);
    } else {
      parent[name] = [existing, value];
    }
  } else {
    const wrap = alwaysArray || (isArray !== null && isArray(name));
    parent[name] = wrap ? [value] : value;
  }
}

/**
 * Put an element's text under `textKey`. When an attribute (`attributePrefix: ""`) or a child
 * element already took that key, the text is merged with it as a repeated child name is — into
 * an array, the earlier value first — rather than overwriting it.
 */
export function addTextValue(
  obj: Record<string, unknown>,
  text: string,
  opts: ResolvedOptions
): void {
  const existing = obj[opts.textKey];
  if (existing === undefined) {
    obj[opts.textKey] = text;
  } else if (Array.isArray(existing)) {
    existing.push(text);
  } else {
    obj[opts.textKey] = [existing, text];
  }
}
