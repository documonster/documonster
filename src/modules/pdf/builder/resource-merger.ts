/**
 * PDF resource dictionary merger — structured merging at the parsed object level.
 *
 * Instead of manipulating serialized PDF dict strings directly, this module
 * parses resource dicts into a structured `PdfResourceDict` (nested Maps),
 * merges them, and serializes back to PDF dict strings.
 */

// =============================================================================
// Types
// =============================================================================

/**
 * Structured representation of a PDF resource dictionary.
 *
 * Outer key: category name (e.g. `Font`, `XObject`, `ExtGState`, `ColorSpace`,
 * `Pattern`, `Shading`, `Properties`).
 * Inner key: resource name (e.g. `F1`, `Im1`, `GS0`).
 * Value: serialized ref or value string (e.g. `3 0 R`, `/DeviceRGB`).
 *
 * Categories whose value is not a sub-dict (e.g. `/ProcSet [/PDF /Text]`)
 * are stored with an empty inner key `""` mapping to the raw value string.
 */
export type PdfResourceDict = Map<string, Map<string, string>>;

// Known resource categories that contain sub-dicts
const SUB_DICT_CATEGORIES = new Set([
  "Font",
  "XObject",
  "ExtGState",
  "ColorSpace",
  "Pattern",
  "Shading",
  "Properties"
]);

// =============================================================================
// Merge
// =============================================================================

/**
 * Merge two parsed resource dicts. For sub-dict categories, entries are
 * combined; overlay entries win on name collision. For non-sub-dict categories,
 * the overlay value replaces the original.
 */
export function mergeResourceDicts(
  original: PdfResourceDict,
  overlay: PdfResourceDict
): PdfResourceDict {
  const merged: PdfResourceDict = new Map();

  // Copy all original entries (deep-copy inner maps)
  for (const [category, innerMap] of original) {
    merged.set(category, new Map(innerMap));
  }

  // Merge overlay entries
  for (const [category, overlayInner] of overlay) {
    const existing = merged.get(category);
    if (existing && SUB_DICT_CATEGORIES.has(category)) {
      // Both have this sub-dict category — merge inner entries (overlay wins)
      for (const [name, value] of overlayInner) {
        existing.set(name, value);
      }
    } else if (existing && !SUB_DICT_CATEGORIES.has(category)) {
      // Non-sub-dict category — overlay replaces
      merged.set(category, new Map(overlayInner));
    } else {
      // Category only in overlay — copy it
      merged.set(category, new Map(overlayInner));
    }
  }

  return merged;
}

// =============================================================================
// Parse
// =============================================================================

/**
 * Parse a serialized PDF resource dict string into a structured `PdfResourceDict`.
 *
 * Handles top-level entries like:
 * ```
 * << /Font << /F1 3 0 R /F2 5 0 R >> /XObject << /Im1 7 0 R >> /ProcSet [/PDF /Text] >>
 * ```
 *
 * Sub-dict categories (`Font`, `XObject`, etc.) are parsed into inner name→value maps.
 * Other categories are stored with a single entry keyed by `""`.
 */
export function parseResourceDict(dictStr: string): PdfResourceDict {
  const result: PdfResourceDict = new Map();
  const trimmed = dictStr.trim();

  if (!trimmed.startsWith("<<") || !trimmed.endsWith(">>")) {
    return result;
  }

  // Strip the outer << >>
  const inner = trimmed.slice(2, -2).trim();
  if (!inner) {
    return result;
  }

  // Parse top-level entries
  const entries = parseDictEntries(inner);

  for (const [key, value] of entries) {
    const valueTrimmed = value.trim();
    if (SUB_DICT_CATEGORIES.has(key) && valueTrimmed.startsWith("<<")) {
      // Parse the sub-dict into name→value pairs
      const subInner = valueTrimmed.slice(2, -2).trim();
      const subEntries = parseDictEntries(subInner);
      const innerMap = new Map<string, string>();
      for (const [subKey, subValue] of subEntries) {
        innerMap.set(subKey, subValue.trim());
      }
      result.set(key, innerMap);
    } else {
      // Non-sub-dict category — store with empty key
      const innerMap = new Map<string, string>();
      innerMap.set("", valueTrimmed);
      result.set(key, innerMap);
    }
  }

  return result;
}

// =============================================================================
// Serialize
// =============================================================================

/**
 * Serialize a structured `PdfResourceDict` back to a PDF dict string.
 */
export function serializeResourceDict(dict: PdfResourceDict): string {
  if (dict.size === 0) {
    return "<< >>";
  }

  const parts: string[] = ["<<"];

  for (const [category, innerMap] of dict) {
    if (SUB_DICT_CATEGORIES.has(category) && !(innerMap.size === 1 && innerMap.has(""))) {
      // Sub-dict category — serialize as << /Name value ... >>
      const subParts: string[] = ["<<"];
      for (const [name, value] of innerMap) {
        subParts.push(`/${name} ${value}`);
      }
      subParts.push(">>");
      parts.push(`/${category} ${subParts.join(" ")}`);
    } else {
      // Non-sub-dict or single-value category
      const value = innerMap.get("") ?? "";
      if (value) {
        parts.push(`/${category} ${value}`);
      }
    }
  }

  parts.push(">>");
  return parts.join(" ");
}

// =============================================================================
// Internal: Dict Entry Parser
// =============================================================================

/**
 * Parse the inner content of a `<< ... >>` dict (with outer delimiters already
 * stripped) into key→value pairs. Handles nested `<< >>` and `[ ]` by depth
 * counting.
 */
function parseDictEntries(inner: string): Map<string, string> {
  const entries = new Map<string, string>();
  let i = 0;

  while (i < inner.length) {
    // Skip whitespace
    while (i < inner.length && isWhitespace(inner[i])) {
      i++;
    }
    if (i >= inner.length) {
      break;
    }

    // Expect a name: /KeyName
    if (inner[i] !== "/") {
      break; // malformed — bail
    }
    i++; // skip '/'
    const nameStart = i;
    while (i < inner.length && !isDelimiterOrWhitespace(inner[i])) {
      i++;
    }
    const key = inner.slice(nameStart, i);

    // Skip whitespace
    while (i < inner.length && isWhitespace(inner[i])) {
      i++;
    }

    // Read the value
    const valueStart = i;
    if (i < inner.length && inner[i] === "<" && i + 1 < inner.length && inner[i + 1] === "<") {
      // Sub-dict: count nested << >>
      i = skipNestedDict(inner, i);
    } else if (i < inner.length && inner[i] === "[") {
      // Array: find matching ]
      i = skipNestedArray(inner, i);
    } else {
      // Token(s) — read until next top-level '/' or end
      i = skipTokenValue(inner, i, valueStart);
    }

    const value = inner.slice(valueStart, i).trim();
    if (key && value) {
      entries.set(key, value);
    }
  }

  return entries;
}

function isWhitespace(ch: string): boolean {
  return ch === " " || ch === "\t" || ch === "\n" || ch === "\r" || ch === "\f";
}

function isDelimiterOrWhitespace(ch: string): boolean {
  return (
    isWhitespace(ch) ||
    ch === "/" ||
    ch === "<" ||
    ch === ">" ||
    ch === "[" ||
    ch === "]" ||
    ch === "(" ||
    ch === ")"
  );
}

/** Skip over a `<< ... >>` block, handling nesting. Returns position after `>>`. */
function skipNestedDict(str: string, start: number): number {
  let i = start;
  let depth = 0;
  while (i < str.length) {
    if (str[i] === "<" && i + 1 < str.length && str[i + 1] === "<") {
      depth++;
      i += 2;
    } else if (str[i] === ">" && i + 1 < str.length && str[i + 1] === ">") {
      depth--;
      i += 2;
      if (depth === 0) {
        return i;
      }
    } else {
      i++;
    }
  }
  return i;
}

/** Skip over a `[ ... ]` block, handling nesting. Returns position after `]`. */
function skipNestedArray(str: string, start: number): number {
  let i = start;
  let depth = 0;
  while (i < str.length) {
    if (str[i] === "[") {
      depth++;
    } else if (str[i] === "]") {
      depth--;
      if (depth === 0) {
        i++;
        return i;
      }
    }
    i++;
  }
  return i;
}

/**
 * Skip a token value (ref, name, number, etc.) that is not a dict or array.
 * Stops at the next top-level `/` preceded by whitespace, or at a `<<` pair.
 */
function skipTokenValue(str: string, i: number, valueStart: number): number {
  while (i < str.length) {
    if (str[i] === "/" && i > valueStart) {
      // Check if this '/' starts a new key (preceded by whitespace)
      if (isWhitespace(str[i - 1])) {
        break;
      }
    }
    if (str[i] === "<" && i + 1 < str.length && str[i + 1] === "<") {
      break; // shouldn't happen at this level, but be safe
    }
    i++;
  }
  return i;
}
