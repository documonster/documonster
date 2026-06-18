/**
 * OpenDoPE Data Binding API
 *
 * Resolve data bindings in a document by evaluating XPath expressions
 * against CustomXML parts and populating the SDT content with values.
 */

import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  Run,
  Table,
  StructuredDocumentTag,
  SdtDataBinding
} from "@word/types";
import { parseXml } from "@xml/dom";
import type { XmlElement, XmlNode } from "@xml/types";

// =============================================================================
// Public API
// =============================================================================

/**
 * Resolve data bindings in a document by evaluating XPath expressions
 * against CustomXML parts and populating the SDT content with values.
 *
 * Supports:
 * - Simple value binding: SDT text is replaced with XPath evaluation result
 * - Repeating sections: SDT with `repeatingSection` is duplicated per data row
 * - Conditional display: SDT is removed if XPath returns empty/false
 *
 * @param doc - The document to process (returns a new modified copy).
 * @param data - Optional override data map (storeItemId → XML string). If not
 *              provided, uses the document's embedded CustomXML parts.
 * @returns A new document with bindings resolved.
 */
export function resolveDataBindings(
  doc: DocxDocument,
  data?: ReadonlyMap<string, string>
): DocxDocument {
  // Build a map of storeItemId → parsed XML content
  const xmlStore = new Map<string, string>();
  if (data) {
    for (const [key, value] of data) {
      xmlStore.set(key.toLowerCase(), value);
    }
  } else if (doc.customXmlParts) {
    for (const part of doc.customXmlParts) {
      xmlStore.set(part.itemId.toLowerCase(), part.xmlContent);
    }
  }

  if (xmlStore.size === 0) {
    return doc;
  }

  const newBody = doc.body.map(block => resolveBlockBinding(block, xmlStore));
  return { ...doc, body: newBody };
}

// =============================================================================
// Internal helpers
// =============================================================================

function resolveBlockBinding(block: BodyContent, xmlStore: Map<string, string>): BodyContent {
  if (block.type === "sdt") {
    return resolveSdtBinding(block, xmlStore);
  }
  if (block.type === "table") {
    return {
      ...block,
      rows: block.rows.map(r => ({
        ...r,
        cells: r.cells.map(c => ({
          ...c,
          content: c.content.map(inner => resolveBlockBinding(inner, xmlStore)) as readonly (
            | Paragraph
            | Table
          )[]
        }))
      }))
    };
  }
  return block;
}

function resolveSdtBinding(sdt: StructuredDocumentTag, xmlStore: Map<string, string>): BodyContent {
  const binding = sdt.properties?.dataBinding;
  if (!binding) {
    // No binding — recurse into content for nested SDTs/tables
    return {
      ...sdt,
      content: sdt.content.map(c => {
        if ("type" in c && c.type === "table") {
          return resolveBlockBinding(c, xmlStore) as Table;
        }
        return c;
      })
    };
  }

  // Resolve the value from CustomXML using simple XPath evaluation
  const value = evaluateSimpleXPath(binding, xmlStore);

  if (sdt.properties?.repeatingSection) {
    // Repeating section: value should be a collection
    // Each child in the data produces a copy of the repeatingSection item template
    const values = evaluateRepeatingXPath(binding, xmlStore);
    if (values.length === 0) {
      // Remove the SDT entirely if no data
      return { type: "paragraph", children: [] };
    }
    // Return the SDT with repeated content, with inner bindings resolved per item
    const repeatedContent: (Paragraph | Run | Table)[] = [];
    const template = sdt.content;

    // Build parent xpath for resolving child element values
    const parentXpath = binding.xpath;
    const storeId = binding.storeItemId.replace(/[{}]/g, "").toLowerCase();
    const xmlContent = xmlStore.get(storeId) ?? xmlStore.get(`{${storeId}}`);

    for (let itemIdx = 0; itemIdx < values.length; itemIdx++) {
      // Clone template content for each item
      for (const item of template) {
        const cloned = structuredClone(item) as Paragraph | Run | Table;
        // Resolve inner SDT bindings within this repeated item
        if (xmlContent) {
          resolveInnerBindings(cloned, parentXpath, itemIdx + 1, xmlStore);
        }
        repeatedContent.push(cloned);
      }
    }
    return { ...sdt, content: repeatedContent };
  }

  if (value === null) {
    // Condition not met — remove SDT (return empty paragraph)
    return { type: "paragraph", children: [] };
  }

  // Simple binding — replace text content of the SDT
  const textRun: Run = {
    content: [{ type: "text", text: value }]
  };
  const textPara: Paragraph = { type: "paragraph", children: [textRun] };
  return { ...sdt, content: [textPara] };
}

/**
 * Evaluate an XPath expression against CustomXML stored as a string.
 *
 * Supports the XPath 1.0 subset that real-world OpenDoPE-style bindings
 * exercise:
 *   - Absolute paths: `/root/child/grandchild`
 *   - Namespace prefixes (matched by local name): `/ns0:root/ns0:child`
 *   - Position predicates: `/root/item[1]`, `/root/item[last()]`
 *   - Attribute access: `/root/element/@attr`
 *   - Text content extraction from leaf nodes (CDATA-aware,
 *     entity-decoded by the XML parser)
 *
 * Does NOT support: axes (ancestor::, following::), functions
 * (count(), string()), arithmetic, union (|), or complex predicates.
 *
 * Implemented on top of the project's XML DOM parser. The previous
 * regex-based scanner was unsafe — it returned entity-encoded text
 * verbatim and could be confused by CDATA, comments, attributes
 * containing `>`, etc.
 *
 * @returns The text content of the matched element/attribute, or null if
 *          not found. Returns the empty string for an element that exists
 *          but has no text — callers that need to distinguish empty from
 *          missing must compare to `null` explicitly.
 */
function evaluateSimpleXPath(
  binding: SdtDataBinding,
  xmlStore: Map<string, string>
): string | null {
  const root = loadStoreRoot(binding, xmlStore);
  if (!root) {
    return null;
  }
  const xpath = binding.xpath;

  // Attribute access: /path/to/element/@attr
  const attrMatch = xpath.match(/^(.+)\/@([a-zA-Z0-9_:-]+)$/);
  if (attrMatch) {
    const elementPath = attrMatch[1]!;
    const attrName = attrMatch[2]!;
    const localAttr = stripPrefix(attrName);
    const el = findElementByPath(root, elementPath);
    if (!el) {
      return null;
    }
    // Match attribute by local name, ignoring any prefix in the source.
    for (const [name, value] of Object.entries(el.attributes)) {
      if (stripPrefix(name) === localAttr) {
        return value;
      }
    }
    return null;
  }

  const el = findElementByPath(root, xpath);
  if (!el) {
    return null;
  }
  return collectTextContent(el);
}

interface PathSegment {
  readonly localName: string;
  /** 1-based position; 0 = first/any; -1 = last(). */
  readonly position: number;
}

function parsePathSegment(raw: string): PathSegment {
  let seg = stripPrefix(raw);
  let position = 0;
  const predMatch = seg.match(/\[(\d+|last\(\))\]$/);
  if (predMatch) {
    seg = seg.substring(0, seg.indexOf("["));
    if (predMatch[1] === "last()") {
      position = -1;
    } else {
      position = parseInt(predMatch[1]!, 10);
    }
  }
  return { localName: seg, position };
}

function stripPrefix(name: string): string {
  const i = name.indexOf(":");
  return i >= 0 ? name.substring(i + 1) : name;
}

function elementLocal(el: XmlElement): string {
  return el.local ?? stripPrefix(el.name);
}

function loadStoreRoot(binding: SdtDataBinding, xmlStore: Map<string, string>): XmlElement | null {
  const storeId = binding.storeItemId.replace(/[{}]/g, "").toLowerCase();
  const xmlContent = xmlStore.get(storeId) ?? xmlStore.get(`{${storeId}}`);
  if (!xmlContent) {
    return null;
  }
  try {
    return parseXml(xmlContent).root;
  } catch {
    // Malformed CustomXML store. Treat as empty rather than throwing —
    // resolveDataBindings is best-effort by design.
    return null;
  }
}

/** Direct-child elements of `el` filtered by local name. */
function directChildren(el: XmlElement, localName: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const child of el.children) {
    if (child.type === "element" && elementLocal(child) === localName) {
      out.push(child);
    }
  }
  return out;
}

/** Resolve a path of segments against `root`. The first segment is matched
 *  against `root` itself (XPath absolute paths begin with the document root). */
function findElementByPath(root: XmlElement, xpath: string): XmlElement | null {
  const segments = xpath
    .split("/")
    .filter(s => s.length > 0)
    .map(parsePathSegment);
  if (segments.length === 0) {
    return root;
  }

  // First segment must match the document root by local name.
  const first = segments[0]!;
  if (elementLocal(root) !== first.localName) {
    return null;
  }
  let current: XmlElement | null = root;
  for (let i = 1; i < segments.length; i++) {
    if (!current) {
      return null;
    }
    const seg = segments[i]!;
    const matches = directChildren(current, seg.localName);
    if (matches.length === 0) {
      return null;
    }
    if (seg.position === -1) {
      current = matches[matches.length - 1]!;
    } else if (seg.position > 0) {
      current = matches[seg.position - 1] ?? null;
    } else {
      current = matches[0]!;
    }
  }
  return current;
}

/** Concatenate descendant text and CDATA content. */
function collectTextContent(el: XmlElement): string {
  let out = "";
  const walk = (node: XmlNode): void => {
    if (node.type === "text" || node.type === "cdata") {
      out += node.value;
      return;
    }
    if (node.type === "element") {
      for (const child of node.children) {
        walk(child);
      }
    }
  };
  for (const child of el.children) {
    walk(child);
  }
  return out;
}

function evaluateRepeatingXPath(binding: SdtDataBinding, xmlStore: Map<string, string>): string[] {
  const root = loadStoreRoot(binding, xmlStore);
  if (!root) {
    return [];
  }
  const segments = binding.xpath
    .split("/")
    .filter(s => s.length > 0)
    .map(parsePathSegment);
  if (segments.length === 0) {
    return [];
  }

  // Walk to the parent of the last segment, then count occurrences of the
  // last segment's local name as direct children.
  const parentSegments = segments.slice(0, -1);
  const last = segments[segments.length - 1]!;

  // Resolve the parent path. If only the root segment was provided we
  // treat root as parent; if the parent has its own segments we walk
  // them through findElementByPath against the root using a synthesised
  // path.
  let parent: XmlElement | null;
  if (parentSegments.length === 0) {
    parent = elementLocal(root) === last.localName ? null : root;
    // Special case: `/root` itself as the repeating xpath — treat as one item.
    if (elementLocal(root) === last.localName) {
      return ["0"];
    }
  } else {
    const parentPath =
      "/" +
      parentSegments
        .map(s =>
          s.position > 0
            ? `${s.localName}[${s.position}]`
            : s.position === -1
              ? `${s.localName}[last()]`
              : s.localName
        )
        .join("/");
    parent = findElementByPath(root, parentPath);
  }

  if (!parent) {
    return [];
  }
  const count = directChildren(parent, last.localName).length;
  return Array.from({ length: count }, (_, i) => String(i));
}

/**
 * Resolve inner SDT data bindings within a repeated section item.
 * Rewrites XPath in inner bindings to reference the specific item index,
 * then evaluates and replaces the SDT text content.
 */
function resolveInnerBindings(
  content: unknown,
  parentXpath: string,
  itemPosition: number,
  xmlStore: Map<string, string>
): void {
  if (!content || typeof content !== "object") {
    return;
  }

  const obj = content as Record<string, unknown>;
  // If this is an SDT with a dataBinding, resolve it in context of the repeated item
  if (
    obj.type === "sdt" &&
    typeof obj.properties === "object" &&
    obj.properties !== null &&
    (obj.properties as Record<string, unknown>).dataBinding
  ) {
    const innerBinding = (obj.properties as Record<string, unknown>).dataBinding as SdtDataBinding;
    // Rewrite the xpath: if its leading path-segment sequence equals
    // `parentXpath`, insert a position predicate at the parent boundary.
    // The previous implementation used a plain string `startsWith` check,
    // which produced false positives when `parentXpath` was a prefix of a
    // sibling's name (e.g. `/root/items` is a string prefix of
    // `/root/items_total/count` even though `/root/items` is not a parent
    // path at the segment level).
    let resolvedXpath = innerBinding.xpath;
    if (xpathStartsWithSegments(resolvedXpath, parentXpath)) {
      resolvedXpath = `${parentXpath}[${itemPosition}]${resolvedXpath.slice(parentXpath.length)}`;
    }

    const modifiedBinding: SdtDataBinding = {
      ...innerBinding,
      xpath: resolvedXpath
    };

    const value = evaluateSimpleXPath(modifiedBinding, xmlStore);
    if (value !== null) {
      const textRun: Run = { content: [{ type: "text", text: value }] };
      const textPara: Paragraph = { type: "paragraph", children: [textRun] };
      obj.content = [textPara];
    }
    return;
  }

  // Recurse into arrays and object properties
  if (Array.isArray(content)) {
    for (const item of content) {
      resolveInnerBindings(item, parentXpath, itemPosition, xmlStore);
    }
  } else {
    for (const key of Object.keys(obj)) {
      const val = obj[key];
      if (val && typeof val === "object") {
        resolveInnerBindings(val, parentXpath, itemPosition, xmlStore);
      }
    }
  }
}

/**
 * Test whether `xpath` starts with `prefix` *as a sequence of path
 * segments* — i.e. the next character after `prefix` is `/`, `[`, or end
 * of string. Plain string `startsWith` mismatches sibling names (e.g.
 * `/root/items` vs. `/root/items_total`).
 */
function xpathStartsWithSegments(xpath: string, prefix: string): boolean {
  if (!xpath.startsWith(prefix)) {
    return false;
  }
  if (xpath.length === prefix.length) {
    return true;
  }
  const next = xpath.charAt(prefix.length);
  return next === "/" || next === "[";
}
