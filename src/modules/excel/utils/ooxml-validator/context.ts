/**
 * Shared validation context.
 *
 * Checkers share XML parsing work through this context. Relationships,
 * content types, and parsed XML DOMs are parsed at most once and cached.
 * Every checker consumes from the same cache so they all see identical
 * parse results (this matters for "rels malformed" propagation — the
 * malformed entry is reported once, and every downstream checker simply
 * sees the empty parse result).
 */

import type { ExtractedFile } from "@archive/unzip/extract";
import type { Reporter } from "@excel/utils/ooxml-validator/reporter";
import {
  attrByLocalName,
  findChildrenLocal,
  tryParseXml
} from "@excel/utils/ooxml-validator/xml-utils";
import type { XmlDocument, XmlElement } from "@xml/types";

const textDecoder = new TextDecoder();

// -----------------------------------------------------------------------------
// Parsed representations
// -----------------------------------------------------------------------------

export interface ContentTypesData {
  defaults: Map<string, string>; // extension (lowercase) -> contentType
  overrides: Map<string, string>; // partName (no leading /) -> contentType
  /**
   * `true` if the file was parsed as a `<Types>` root successfully.
   * `false` if the file is missing, malformed or has an unexpected root.
   */
  parseOk: boolean;
  /**
   * Duplicate override part names encountered during parse. Emitted as
   * problems by the content-types checker.
   */
  duplicateOverrides: string[];
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

export interface RelsData {
  rels: Relationship[];
  byId: Map<string, Relationship>;
  parseOk: boolean;
  /**
   * Raw rel entries that lacked Id/Type/Target attributes. Emitted
   * separately so the rels checker can flag them without us dropping
   * them from the data entirely.
   */
  malformedEntries: Array<{
    missingId: boolean;
    missingType: boolean;
    missingTarget: boolean;
    id?: string;
    type?: string;
  }>;
}

// -----------------------------------------------------------------------------
// Context
// -----------------------------------------------------------------------------

export class ValidationContext {
  readonly entries: Map<string, ExtractedFile>;
  readonly reporter: Reporter;

  private readonly textCache = new Map<string, string>();
  private readonly domCache = new Map<string, XmlDocument | undefined>();
  private readonly relsCache = new Map<string, RelsData>();
  private contentTypes?: ContentTypesData;

  constructor(entries: Map<string, ExtractedFile>, reporter: Reporter) {
    this.entries = entries;
    this.reporter = reporter;
  }

  /** `true` if `path` points to a file (not a directory) in the package. */
  has(path: string): boolean {
    const entry = this.entries.get(path);
    return !!entry && entry.type !== "directory";
  }

  /** Iterate every file (skip directories). */
  *files(): IterableIterator<readonly [string, ExtractedFile]> {
    for (const entry of this.entries) {
      if (entry[1].type !== "directory") {
        yield entry;
      }
    }
  }

  /** Decode a part as UTF-8 text. Cached. */
  readText(path: string): string | undefined {
    const cached = this.textCache.get(path);
    if (cached !== undefined) {
      return cached;
    }
    const entry = this.entries.get(path);
    if (!entry || entry.type === "directory") {
      return undefined;
    }
    const text = textDecoder.decode(entry.data);
    this.textCache.set(path, text);
    return text;
  }

  /**
   * Parse a part as XML DOM. Cached. Returns `undefined` if the part is
   * missing or malformed. Callers that need to distinguish those cases
   * should check `has(path)` first.
   */
  readDom(path: string, onMalformed?: (err: Error) => void): XmlDocument | undefined {
    if (this.domCache.has(path)) {
      return this.domCache.get(path);
    }
    const text = this.readText(path);
    if (text === undefined) {
      this.domCache.set(path, undefined);
      return undefined;
    }
    const dom = tryParseXml(text, err => {
      onMalformed?.(err);
    });
    this.domCache.set(path, dom);
    return dom;
  }

  /**
   * Read and cache the content-types table. On first call, parse
   * `[Content_Types].xml`; subsequent calls return the cache.
   * Reports malformed-XML problems via the reporter but never throws.
   */
  readContentTypes(): ContentTypesData {
    if (this.contentTypes) {
      return this.contentTypes;
    }
    const path = "[Content_Types].xml";
    const dom = this.readDom(path, err => {
      this.reporter.error(
        "content-types-malformed",
        `Content types XML parse error: ${err.message}`,
        path
      );
    });
    const data: ContentTypesData = {
      defaults: new Map(),
      overrides: new Map(),
      parseOk: false,
      duplicateOverrides: []
    };
    if (!dom) {
      this.contentTypes = data;
      return data;
    }
    const root = dom.root;
    // Accept namespaced <Types> just in case.
    const looksLikeTypes = root.name === "Types" || root.name.endsWith(":Types");
    if (!looksLikeTypes) {
      this.reporter.error(
        "content-types-malformed",
        `Content types root element is <${root.name}>, expected <Types>`,
        path
      );
      this.contentTypes = data;
      return data;
    }
    for (const el of findChildrenLocal(root, "Default")) {
      const extRaw = attrByLocalName(el, "Extension");
      const ct = attrByLocalName(el, "ContentType");
      if (extRaw && ct) {
        data.defaults.set(extRaw.toLowerCase(), ct);
      }
    }
    for (const el of findChildrenLocal(root, "Override")) {
      const partName = attrByLocalName(el, "PartName");
      const ct = attrByLocalName(el, "ContentType");
      if (!partName || !ct) {
        continue;
      }
      const key = partName.startsWith("/") ? partName.slice(1) : partName;
      if (data.overrides.has(key)) {
        data.duplicateOverrides.push(partName);
      } else {
        data.overrides.set(key, ct);
      }
    }
    data.parseOk = true;
    this.contentTypes = data;
    return data;
  }

  /**
   * Read and cache a relationships file. Every .rels in the package is
   * parsed lazily on first request. Malformed files are reported once
   * and return an empty relationships array.
   */
  readRels(relsPath: string): RelsData {
    const cached = this.relsCache.get(relsPath);
    if (cached) {
      return cached;
    }
    const data: RelsData = {
      rels: [],
      byId: new Map(),
      parseOk: false,
      malformedEntries: []
    };
    if (!this.has(relsPath)) {
      this.relsCache.set(relsPath, data);
      return data;
    }
    const dom = this.readDom(relsPath, err => {
      this.reporter.error(
        "rels-malformed",
        `Relationships XML parse error: ${err.message}`,
        relsPath
      );
    });
    if (!dom) {
      this.relsCache.set(relsPath, data);
      return data;
    }
    const root = dom.root;
    const looksLikeRelationships =
      root.name === "Relationships" || root.name.endsWith(":Relationships");
    if (!looksLikeRelationships) {
      this.reporter.error(
        "rels-malformed",
        `Relationships root element is <${root.name}>, expected <Relationships>`,
        relsPath
      );
      this.relsCache.set(relsPath, data);
      return data;
    }
    for (const el of findChildrenLocal(root, "Relationship")) {
      const id = attrByLocalName(el, "Id");
      const type = attrByLocalName(el, "Type");
      const target = attrByLocalName(el, "Target");
      const targetMode = attrByLocalName(el, "TargetMode");
      if (!id || !type || target === undefined) {
        data.malformedEntries.push({
          missingId: !id,
          missingType: !type,
          missingTarget: target === undefined,
          id,
          type
        });
        continue;
      }
      const rel: Relationship = { id, type, target, targetMode };
      data.rels.push(rel);
      // Note: duplicates are checked elsewhere; we keep the FIRST here.
      if (!data.byId.has(id)) {
        data.byId.set(id, rel);
      }
    }
    data.parseOk = true;
    this.relsCache.set(relsPath, data);
    return data;
  }
}

/** Get the root element of a DOM; returns undefined for missing DOM. */
export function domRoot(dom: XmlDocument | undefined): XmlElement | undefined {
  return dom?.root;
}
