/**
 * DOCX Module - Content Types Generator
 *
 * Generates [Content_Types].xml for the DOCX package.
 * Uses a plain data record + free functions for tree-shakeability.
 */

import {
  NS_CONTENT_TYPES,
  STD_DOC_ATTRIBUTES,
  ContentType,
  IMAGE_CONTENT_TYPES
} from "@word/constants";
import { DocxInvalidImageError, DocxWriteError } from "@word/errors";
import type { XmlSink } from "@xml/types";

/** Content type override entry. */
export interface ContentTypeOverride {
  readonly partName: string;
  readonly contentType: string;
}

/** Internal state for content types (plain record, not a class). */
export interface ContentTypesState {
  readonly defaults: Map<string, string>;
  readonly overrides: ContentTypeOverride[];
}

/** Create a new ContentTypesState with standard defaults (rels, xml). */
export function createContentTypes(): ContentTypesState {
  const defaults = new Map<string, string>();
  defaults.set("rels", ContentType.Relationships);
  defaults.set("xml", ContentType.Xml);
  return { defaults, overrides: [] };
}

/** Add a default content type for a file extension. */
export function addContentTypeDefault(
  state: ContentTypesState,
  extension: string,
  contentType: string
): void {
  state.defaults.set(extension, contentType);
}

/** Add an override content type for a specific part. Deduplicates by partName. */
export function addContentTypeOverride(
  state: ContentTypesState,
  partName: string,
  contentType: string
): void {
  const normalized = partName.startsWith("/") ? partName : `/${partName}`;
  const overrides = state.overrides as ContentTypeOverride[];

  // Deduplicate: if same partName already exists, update in-place
  const existing = overrides.findIndex(o => o.partName === normalized);
  if (existing !== -1) {
    // Same content type → no-op; different → replace (last wins)
    if (overrides[existing].contentType !== contentType) {
      overrides[existing] = { partName: normalized, contentType };
    }
    return;
  }

  overrides.push({ partName: normalized, contentType });
}

/** Add image extension defaults from a set of used extensions. */
export function addImageContentTypeDefaults(
  state: ContentTypesState,
  extensions: Iterable<string>
): void {
  for (const ext of extensions) {
    const ct = IMAGE_CONTENT_TYPES[ext.toLowerCase()];
    if (ct) {
      state.defaults.set(ext.toLowerCase(), ct);
    }
  }
}

/**
 * Refuse a path that cannot name an OPC part (ECMA-376 Part 2 §9.1.1.1): one
 * ending in "/" (a ZIP folder entry), with an empty segment, a segment ending
 * in ".", or a backslash. Declaring such a name in `[Content_Types].xml` makes
 * the whole package unreadable, so it is an error rather than an output.
 */
export function assertValidPartName(path: string, context: string | undefined): void {
  const segments = path.replace(/^\/+/, "").split("/");
  if (path.includes("\\") || segments.some(s => s === "" || s.endsWith("."))) {
    // The writer shares this check across every writing entry point, so it
    // names no caller (`context` is undefined there).
    const prefix = context === undefined ? "" : `${context}: `;
    throw new DocxWriteError(`${prefix}"${path}" is not a valid part name`);
  }
}

/** The generic content type a part of unknown type is declared with. */
const GENERIC_CONTENT_TYPE = "application/octet-stream";

/**
 * Declare every written part the state does not type yet, because OPC requires a
 * content type for every part (ECMA-376 Part 2 §10.1.2.4) and a package with an
 * untyped part is rejected as unreadable. Such a part is one a caller named — an
 * opaque part without a `contentType`, an embedded font with an unusual extension
 * — and gets the generic `application/octet-stream`: by a `Default` for its
 * extension, or an `Override` when it has none. A media part is never typed
 * generically, since a picture so typed would not render: one is refused.
 */
export function typeUntypedParts(state: ContentTypesState, paths: Iterable<string>): void {
  const overrides = new Set(state.overrides.map(o => o.partName.toLowerCase()));
  for (const path of paths) {
    assertValidPartName(path, undefined);
    const partName = `/${path.replace(/^\/+/, "")}`;
    if (overrides.has(partName.toLowerCase())) {
      continue;
    }
    const leaf = partName.slice(partName.lastIndexOf("/") + 1);
    const dot = leaf.lastIndexOf(".");
    const ext = dot > 0 ? leaf.slice(dot + 1).toLowerCase() : "";
    if (ext && [...state.defaults.keys()].some(k => k.toLowerCase() === ext)) {
      continue;
    }
    if (/\/media\/[^/]+$/.test(partName)) {
      throw new DocxInvalidImageError(
        `parts[${JSON.stringify(path)}]`,
        `media part "${partName}" has no image content type; give it an image extension or a contentType`
      );
    }
    if (ext) {
      state.defaults.set(ext, GENERIC_CONTENT_TYPE);
    } else {
      addContentTypeOverride(state, partName, GENERIC_CONTENT_TYPE);
      overrides.add(partName.toLowerCase());
    }
  }
}

/** Render the [Content_Types].xml to a sink. */
export function renderContentTypes(state: ContentTypesState, xml: XmlSink): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("Types", { xmlns: NS_CONTENT_TYPES });

  // Defaults sorted by extension
  const sortedDefaults = [...state.defaults.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  for (const [ext, ct] of sortedDefaults) {
    xml.leafNode("Default", { Extension: ext, ContentType: ct });
  }

  // Overrides in order
  for (const override of state.overrides) {
    xml.leafNode("Override", { PartName: override.partName, ContentType: override.contentType });
  }

  xml.closeNode();
}
