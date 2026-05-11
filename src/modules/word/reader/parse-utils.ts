/**
 * DOCX Reader - Shared XML/parsing utilities
 *
 * Low-level helpers used across the various reader sub-modules.
 * Extracted from docx-reader.ts to keep that file focused on
 * orchestration of the higher-level parsers.
 */

import { findChild, findChildren } from "@xml/dom";
import type { XmlElement } from "@xml/types";

import { escapeXml } from "../core/internal-utils";
import type { ParsedRelationship } from "./reader-context";

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

// =============================================================================
// Child Element Readers
// =============================================================================

/** Get the local name from a possibly-prefixed tag (`"w:p"` → `"p"`). */
export function localName(name: string): string {
  const idx = name.indexOf(":");
  return idx >= 0 ? name.slice(idx + 1) : name;
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
 * Check for a boolean toggle element.
 *
 * In OOXML, many properties are represented as toggle elements where:
 * - Element present = true
 * - Element with `w:val="0"` or `"false"` = explicit false
 * - Element absent = undefined (use parent's default)
 */
export function boolToggle(parent: XmlElement, name: string): boolean | undefined {
  const el = findChildNs(parent, name);
  if (!el) {
    return undefined;
  }
  const v = attrVal(el, "val");
  if (v === "0" || v === "false") {
    return false;
  }
  return true;
}

// =============================================================================
// XML Serialization
// =============================================================================

/** Serialize an XmlElement back to an XML string (used for opaque preservation). */
export function serializeElement(el: XmlElement): string {
  let s = `<${el.name}`;
  for (const [k, v] of Object.entries(el.attributes)) {
    s += ` ${k}="${escapeXml(v)}"`;
  }
  if (el.children.length === 0) {
    return s + "/>";
  }
  s += ">";
  for (const child of el.children) {
    if (child.type === "element") {
      s += serializeElement(child);
    } else if (child.type === "text") {
      s += escapeXml(child.value);
    }
  }
  s += `</${el.name}>`;
  return s;
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

export { getPartRelsPath, getFileName, getFileExt } from "../core/opc-package";

/**
 * Resolve a relationship target path to an absolute package-root path.
 *
 * - Leading `/` → package root absolute
 * - `../` / `./` → resolved relative to the source part's directory
 * - Plain paths → resolved relative to the source part's directory
 */
export function resolvePartPath(sourcePart: string, target: string): string {
  if (!target) {
    return "";
  }
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  const lastSlash = sourcePart.lastIndexOf("/");
  const baseDir = lastSlash >= 0 ? sourcePart.substring(0, lastSlash).split("/") : [];
  const segs = target.split("/");
  for (const seg of segs) {
    if (seg === "..") {
      baseDir.pop();
    } else if (seg !== "." && seg !== "") {
      baseDir.push(seg);
    }
  }
  return baseDir.join("/");
}

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
  if (!sid) {
    return undefined;
  }
  switch (sid) {
    case "1":
      return "MD2";
    case "2":
      return "MD4";
    case "3":
      return "MD5";
    case "4":
      return "SHA-1";
    case "12":
      return "SHA-256";
    case "13":
      return "SHA-384";
    case "14":
      return "SHA-512";
    default:
      return `SID-${sid}`;
  }
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
