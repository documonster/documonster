/**
 * DOCX Module - Content Types Generator
 *
 * Generates [Content_Types].xml for the DOCX package.
 * Uses a plain data record + free functions for tree-shakeability.
 */

import type { XmlSink } from "@xml/types";

import {
  NS_CONTENT_TYPES,
  STD_DOC_ATTRIBUTES,
  ContentType,
  IMAGE_CONTENT_TYPES
} from "../constants";

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
