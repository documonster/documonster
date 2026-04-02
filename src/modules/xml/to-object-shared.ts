/**
 * XML → Plain Object — Shared Conversion Logic
 *
 * Internal module used by both `toPlainObject` (DOM path) and
 * `parseXmlToObject` (SAX-direct path). Not part of the public API.
 */

import type { ToPlainObjectOptions } from "@xml/types";

// =============================================================================
// Resolved Options
// =============================================================================

/** Options with all defaults resolved — no more `??` checks at hot-path call sites. */
export interface ResolvedOptions {
  readonly attrPrefix: string;
  readonly textKey: string;
  readonly alwaysArray: boolean;
  readonly preserveCData: boolean;
  readonly ignoreWS: boolean;
}

export function resolveOptions(options?: ToPlainObjectOptions): ResolvedOptions {
  return {
    attrPrefix: options?.attributePrefix ?? "@_",
    textKey: options?.textKey ?? "#text",
    alwaysArray: options?.alwaysArray ?? false,
    preserveCData: options?.preserveCData ?? true,
    ignoreWS: options?.ignoreWhitespaceText ?? true
  };
}

// =============================================================================
// Helpers
// =============================================================================

/** Check if a string contains only whitespace without allocating a trimmed copy. */
function isWhitespaceOnly(s: string): boolean {
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
    obj[opts.textKey] = text;
  }

  // Empty element with no attributes → empty string (like fast-xml-parser)
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
  alwaysArray: boolean
): void {
  const existing = parent[name];
  if (existing !== undefined) {
    if (Array.isArray(existing)) {
      existing.push(value);
    } else {
      parent[name] = [existing, value];
    }
  } else {
    parent[name] = alwaysArray ? [value] : value;
  }
}
