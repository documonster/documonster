/**
 * DOCX Module - Content Types Generator
 *
 * Generates [Content_Types].xml for the DOCX package.
 */

import type { XmlSink } from "@xml/types";

import {
  NS_CONTENT_TYPES,
  STD_DOC_ATTRIBUTES,
  ContentType,
  IMAGE_CONTENT_TYPES
} from "./constants";

/** Content type override entry. */
export interface ContentTypeOverride {
  readonly partName: string;
  readonly contentType: string;
}

/**
 * Generates the [Content_Types].xml part.
 */
export class ContentTypesManager {
  private readonly _defaults = new Map<string, string>();
  private readonly _overrides: ContentTypeOverride[] = [];

  constructor() {
    // Always include rels and xml defaults
    this._defaults.set("rels", ContentType.Relationships);
    this._defaults.set("xml", ContentType.Xml);
  }

  /** Add a default content type for a file extension. */
  addDefault(extension: string, contentType: string): void {
    this._defaults.set(extension, contentType);
  }

  /** Add an override content type for a specific part. */
  addOverride(partName: string, contentType: string): void {
    this._overrides.push({
      partName: partName.startsWith("/") ? partName : `/${partName}`,
      contentType
    });
  }

  /** Add image extension defaults from a set of used extensions. */
  addImageDefaults(extensions: Iterable<string>): void {
    for (const ext of extensions) {
      const ct = IMAGE_CONTENT_TYPES[ext.toLowerCase()];
      if (ct) {
        this._defaults.set(ext.toLowerCase(), ct);
      }
    }
  }

  /** Render the [Content_Types].xml to a sink. */
  render(xml: XmlSink): void {
    xml.openXml(STD_DOC_ATTRIBUTES);
    xml.openNode("Types", { xmlns: NS_CONTENT_TYPES });

    // Defaults sorted by extension
    const sortedDefaults = [...this._defaults.entries()].sort((a, b) => a[0].localeCompare(b[0]));
    for (const [ext, ct] of sortedDefaults) {
      xml.leafNode("Default", { Extension: ext, ContentType: ct });
    }

    // Overrides in order
    for (const override of this._overrides) {
      xml.leafNode("Override", { PartName: override.partName, ContentType: override.contentType });
    }

    xml.closeNode();
  }
}
