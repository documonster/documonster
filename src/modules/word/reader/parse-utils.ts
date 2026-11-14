/**
 * DOCX Reader - Shared XML/parsing utilities
 *
 * Low-level helpers used across the various reader sub-modules.
 * Extracted from docx-reader.ts to keep that file focused on
 * orchestration of the higher-level parsers.
 */

import type { Steps } from "@utils/yield";
import { NS_W, STRICT_TO_TRANSITIONAL_NS } from "@word/constants";
import { resolvePartPath } from "@word/core/opc-paths";
import type { ParsedRelationship } from "@word/reader/reader-context";
import { hashAlgorithmOfSid } from "@word/security/hash-names";
import { findChild, findChildren } from "@xml/dom";
import { xmlEncode, xmlEncodeAttr } from "@xml/encode";
import type { XmlElement } from "@xml/types";

// =============================================================================
// Attribute Readers
// =============================================================================

/** Get an attribute value, trying both `w:name` and bare `name`. */
export function attrVal(el: XmlElement, name: string): string | undefined {
  return el.attributes[`w:${name}`] ?? el.attributes[name];
}

/** Get an attribute as an integer (returns undefined if missing or non-numeric). */
export function attrInt(el: XmlElement, name: string): number | undefined {
  const v = attrVal(el, name);
  if (v === undefined) {
    return undefined;
  }
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * Like `parseInt(value ?? fallback, 10)` but never returns NaN: any
 * non-numeric value falls back to `fallback`. Use this when forwarding
 * raw attribute strings into numeric model fields — round-tripping a
 * `NaN` back into XML produces literal `"NaN"` and corrupts the
 * resulting DOCX.
 */
export function safeParseInt(value: string | undefined, fallback: number): number {
  if (value === undefined) {
    return fallback;
  }
  const n = parseInt(value, 10);
  return Number.isFinite(n) ? n : fallback;
}

// =============================================================================
// Child Element Readers
// =============================================================================

/** Get the local name from a possibly-prefixed tag (`"w:p"` → `"p"`). */
export function localName(name: string): string {
  const idx = name.indexOf(":");
  return idx >= 0 ? name.slice(idx + 1) : name;
}

/**
 * `name` without a leading `w:` prefix — what `name.replace(/^w:/, "")` gives,
 * without running a regex for every element the reader dispatches on.
 */
export function stripWPrefix(name: string): string {
  return name.charCodeAt(0) === 0x77 && name.charCodeAt(1) === 0x3a ? name.slice(2) : name;
}

/**
 * The element a paragraph-content child stands for: itself, or — for `mc:AlternateContent` — the first element of its
 * `mc:Choice`, else of its `mc:Fallback`; `undefined` when it holds neither. The reader reads a paragraph's content
 * through it, and so does the field scan (`splitFieldPieces`): a field character the reader meets must be one the scan
 * met, or the field is neither read in pieces nor ever finished.
 */
export function resolveContentChild(child: XmlElement): XmlElement | undefined {
  if (child.name !== "mc:AlternateContent") {
    return child;
  }
  const chosen = findChild(child, "mc:Choice") ?? findChild(child, "mc:Fallback");
  return chosen?.children.find((c): c is XmlElement => c.type === "element");
}

/**
 * Whether `el` is the root of a main document part: `w:document` in Word's main namespace,
 * Transitional or Strict. The part has been read with its names normalised (`parsePartXml`), so the
 * main namespace is bound to `w` on the root whichever prefix — or the default namespace — the file
 * used; an element of another namespace, or a `w:` prefix bound to nothing, is not Word's.
 */
export function isWordDocumentRoot(el: {
  readonly name: string;
  readonly attributes: Readonly<Record<string, string>>;
}): boolean {
  if (el.name !== "w:document") {
    return false;
  }
  const uri = el.attributes["xmlns:w"];
  return uri === NS_W || (uri !== undefined && STRICT_TO_TRANSITIONAL_NS.get(uri) === NS_W);
}

/** Find a child element by local name (matches both `w:name` and bare `name`). */
export function findChildNs(el: XmlElement, localName: string): XmlElement | undefined {
  return findChild(el, `w:${localName}`) ?? findChild(el, localName);
}

/** Find all child elements with a given local name. */
export function findChildrenNs(el: XmlElement, localName: string): XmlElement[] {
  const a = findChildren(el, `w:${localName}`);
  return a.length > 0 ? a : findChildren(el, localName);
}

/**
 * Namespace-agnostic child finder: matches on the **local name** regardless
 * of the prefix. Useful for namespaces other than `w:` (e.g. `m:`, `c:`,
 * `cx:`) where the prefix may be inconsistent across producers.
 */
export function findChildLocal(el: XmlElement, name: string): XmlElement | undefined {
  for (const child of el.children) {
    if (child.type === "element" && localName(child.name) === name) {
      return child;
    }
  }
  return undefined;
}

/** Namespace-agnostic version of {@link findChildrenNs}. */
export function findChildrenLocal(el: XmlElement, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const child of el.children) {
    if (child.type === "element" && localName(child.name) === name) {
      out.push(child);
    }
  }
  return out;
}

/**
 * Get an attribute value by **local name**, regardless of prefix.
 * Useful when the same attribute may appear as `w:val`, `m:val`, etc.
 */
export function attrLocal(el: XmlElement, name: string): string | undefined {
  for (const [k, v] of Object.entries(el.attributes)) {
    if (localName(k) === name) {
      return v;
    }
  }
  return undefined;
}

/**
 * An OOXML on/off attribute (`ST_OnOff`, ECMA-376 §22.9.2.7): `"0"`,
 * `"false"` and `"off"` are false, any other value (`"1"`, `"true"`, `"on"`)
 * is true; missing yields `undefined`.
 */
export function attrOnOff(el: XmlElement, name: string): boolean | undefined {
  const v = attrVal(el, name);
  if (v === undefined) {
    return undefined;
  }
  return v !== "0" && v !== "false" && v !== "off";
}

/**
 * Check for a boolean toggle element.
 *
 * In OOXML, many properties are represented as toggle elements where:
 * - Element present = true
 * - Element with `w:val` `"0"`, `"false"` or `"off"` = explicit false (`ST_OnOff`)
 * - Element absent = undefined (use parent's default)
 */
export function boolToggle(parent: XmlElement, name: string): boolean | undefined {
  const el = findChildNs(parent, name);
  if (!el) {
    return undefined;
  }
  return attrOnOff(el, "val") ?? true;
}

// =============================================================================
// XML Serialization
// =============================================================================

/** Serialize an XmlElement back to an XML string (used for opaque preservation). */
export function serializeElement(el: XmlElement): string {
  if (el.children.length === 0) {
    return startTag(el) + "/>";
  }
  let s = startTag(el) + ">";
  for (const child of el.children) {
    if (child.type === "element") {
      s += serializeElement(child);
    } else if (child.type === "text") {
      s += xmlEncode(child.value);
    }
  }
  s += `</${el.name}>`;
  return s;
}

/** How many elements {@link serializeElementSteps} serialises in one step. */
const SERIALIZE_STEP = 2048;

/**
 * {@link serializeElement} in steps of about {@link SERIALIZE_STEP} elements, for an element kept as markup however
 * large it is (an inline content control around a whole paragraph's runs). A subtree no larger than a step is
 * serialised by `serializeElement` itself.
 */
export function* serializeElementSteps(el: XmlElement): Steps<string> {
  if (!hasMoreElements(el, SERIALIZE_STEP)) {
    return serializeElement(el);
  }
  // Each step's pieces are joined as the step ends, and the steps' at the end: joining a hundred thousand pieces in
  // one call was a step of its own, as long as the rest together.
  const chunks: string[] = [];
  let parts: string[] = [startTag(el) + ">"];
  for (const child of el.children) {
    if (child.type === "element") {
      parts.push(yield* serializeElementSteps(child));
      if (parts.length >= SERIALIZE_STEP) {
        chunks.push(parts.join(""));
        parts = [];
        yield;
      }
    } else if (child.type === "text") {
      parts.push(xmlEncode(child.value));
    }
  }
  parts.push(`</${el.name}>`);
  chunks.push(parts.join(""));
  return chunks.join("");
}

function startTag(el: XmlElement): string {
  let s = `<${el.name}`;
  for (const [k, v] of Object.entries(el.attributes)) {
    s += ` ${k}="${xmlEncodeAttr(v)}"`;
  }
  return s;
}

/** Whether the subtree of `el` holds more than `limit` elements, counting no further than that. */
function hasMoreElements(el: XmlElement, limit: number): boolean {
  let count = 0;
  const pending: XmlElement[] = [el];
  while (pending.length > 0) {
    for (const child of pending.pop()!.children) {
      if (child.type === "element") {
        if (++count > limit) {
          return true;
        }
        pending.push(child);
      }
    }
  }
  return false;
}

/** Extract all `r:xxx` attribute values (relationship IDs) from an element tree. */
export function collectRIds(el: XmlElement, out: Set<string>): void {
  for (const [k, v] of Object.entries(el.attributes)) {
    if (k.startsWith("r:") || k === "r:id" || k === "r:embed" || k === "r:link") {
      out.add(v);
    }
  }
  for (const child of el.children) {
    if (child.type === "element") {
      collectRIds(child, out);
    }
  }
}

// =============================================================================
// Path / Relationship Helpers
// =============================================================================

export { getPartRelsPath, getFileName, getFileExt, resolvePartPath } from "@word/core/opc-paths";

/**
 * Find the first relationship of a given type and resolve its target path.
 * Returns undefined if no matching relationship exists.
 */
export function resolveRelTarget(
  rels: ParsedRelationship[],
  relType: string,
  sourcePart: string
): string | undefined {
  for (const rel of rels) {
    if (rel.type === relType) {
      return resolvePartPath(sourcePart, rel.target);
    }
  }
  return undefined;
}

// =============================================================================
// Crypto Helpers
// =============================================================================

/** Convert OOXML cryptAlgorithmSid to human-readable algorithm name. */
export function sidToHashAlgorithm(sid: string | undefined): string | undefined {
  return hashAlgorithmOfSid(sid);
}

// =============================================================================
// Note Properties (footnote/endnote shared structure)
// =============================================================================

/**
 * Common shape returned by {@link parseNoteProperties}.
 *
 * `numFmt` is widened to `string` (rather than `NoteNumberFormat`) so the
 * helper can populate either `FootnoteProperties` or `EndnoteProperties`,
 * which only differ in `position`'s narrow union. Callers cast at the
 * assignment site.
 */
export interface ParsedNoteProperties {
  numFmt?: string;
  numStart?: number;
  numRestart?: string;
  position?: string;
}

/** Parse footnote/endnote properties element. */
export function parseNoteProperties(el: XmlElement): ParsedNoteProperties | undefined {
  const props: ParsedNoteProperties = {};
  const numFmtEl = findChildNs(el, "numFmt");
  if (numFmtEl) {
    props.numFmt = attrVal(numFmtEl, "val");
  }
  const numStartEl = findChildNs(el, "numStart");
  if (numStartEl) {
    props.numStart = attrInt(numStartEl, "val");
  }
  const numRestartEl = findChildNs(el, "numRestart");
  if (numRestartEl) {
    props.numRestart = attrVal(numRestartEl, "val");
  }
  const posEl = findChildNs(el, "pos");
  if (posEl) {
    props.position = attrVal(posEl, "val");
  }
  return Object.keys(props).length > 0 ? props : undefined;
}
