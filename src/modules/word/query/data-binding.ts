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
} from "../types";

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
 * Supports XPath 1.0 subset:
 * - Absolute paths: `/root/child/grandchild`
 * - Namespace prefixes: `/ns0:root/ns0:child` (stripped for matching)
 * - Position predicates: `/root/item[1]`, `/root/item[last()]`
 * - Attribute access: `/root/element/@attr`
 * - Text content extraction from leaf nodes
 *
 * Does NOT support: axes (ancestor::, following::), functions (count(), string()),
 * arithmetic, union (|), or complex predicates.
 *
 * @returns The text content of the matched element, or null if not found.
 */
function evaluateSimpleXPath(
  binding: SdtDataBinding,
  xmlStore: Map<string, string>
): string | null {
  const storeId = binding.storeItemId.replace(/[{}]/g, "").toLowerCase();
  const xmlContent = xmlStore.get(storeId) ?? xmlStore.get(`{${storeId}}`);
  if (!xmlContent) {
    return null;
  }

  const xpath = binding.xpath;

  // Handle attribute access: /root/element/@attrName
  const attrMatch = xpath.match(/^(.+)\/@([a-zA-Z0-9_:-]+)$/);
  if (attrMatch) {
    const elementPath = attrMatch[1];
    const attrName = attrMatch[2];
    const localAttr = attrName.includes(":")
      ? attrName.substring(attrName.indexOf(":") + 1)
      : attrName;
    const element = findElementByPath(xmlContent, elementPath);
    if (!element) {
      return null;
    }
    // Extract attribute value from the element's opening tag
    const attrPattern = new RegExp(
      `(?:^|\\s)(?:[a-zA-Z0-9_-]+:)?${escapeRegex(localAttr)}\\s*=\\s*"([^"]*)"`,
      "i"
    );
    const am = attrPattern.exec(element.openTag);
    return am ? am[1] : null;
  }

  const element = findElementByPath(xmlContent, xpath);
  if (!element) {
    return null;
  }

  // Strip child elements and return text content only
  const textContent = element.innerHTML.replace(/<[^>]*>/g, "").trim();
  return textContent || null;
}

interface FoundElement {
  openTag: string;
  innerHTML: string;
}

/**
 * Find an element in raw XML by following a path like /root/child[2]/grandchild.
 */
function findElementByPath(xml: string, xpath: string): FoundElement | null {
  // Parse path segments
  const segments = xpath
    .split("/")
    .filter(s => s.length > 0)
    .map(parsePathSegment);

  let current = xml;
  let lastOpenTag = "";

  for (const seg of segments) {
    const result = findNthElement(current, seg.localName, seg.position);
    if (!result) {
      return null;
    }
    lastOpenTag = result.openTag;
    current = result.innerHTML;
  }

  return { openTag: lastOpenTag, innerHTML: current };
}

interface PathSegment {
  localName: string;
  position: number; // 1-based, 0 = any/first
}

function parsePathSegment(raw: string): PathSegment {
  // Strip namespace prefix
  let seg = raw;
  const colonIdx = seg.indexOf(":");
  if (colonIdx >= 0) {
    seg = seg.substring(colonIdx + 1);
  }

  // Parse predicate [N] or [last()]
  let position = 0;
  const predMatch = seg.match(/\[(\d+|last\(\))\]$/);
  if (predMatch) {
    seg = seg.substring(0, seg.indexOf("["));
    if (predMatch[1] === "last()") {
      position = -1; // Special marker for "last"
    } else {
      position = parseInt(predMatch[1], 10);
    }
  }

  return { localName: seg, position };
}

/**
 * Find the Nth direct-child occurrence of an element with given local name in
 * raw XML. Only matches elements at depth 0 (i.e. immediate children of the
 * current scope) — descendants are skipped. Same-named nested elements are
 * still tracked correctly to find the matching close tag.
 */
function findNthElement(xml: string, localName: string, position: number): FoundElement | null {
  const escapedName = escapeRegex(localName);
  // Scan all element-like tags so we can track depth across the document.
  const allTags = /<\/?([a-zA-Z_][a-zA-Z0-9_-]*(?::[a-zA-Z_][a-zA-Z0-9_-]*)?)(\s[^>]*?)?(\/?)>/g;

  const matches: FoundElement[] = [];
  let depth = 0;
  let match: RegExpExecArray | null;

  while ((match = allTags.exec(xml)) !== null) {
    const fullTagName = match[1];
    const isClosing = match[0].startsWith("</");
    const isSelfClosing = match[3] === "/";

    if (isClosing) {
      depth = Math.max(0, depth - 1);
      continue;
    }

    // Opening tag (self-closing counts as open then immediately closed).
    const openDepthBefore = depth;
    if (!isSelfClosing) {
      depth++;
    }

    // We only care about direct children of the current scope.
    if (openDepthBefore !== 0) {
      continue;
    }

    // Match by local name — strip namespace prefix.
    const colonIdx = fullTagName.indexOf(":");
    const local = colonIdx >= 0 ? fullTagName.substring(colonIdx + 1) : fullTagName;
    if (!new RegExp(`^${escapedName}$`).test(local)) {
      continue;
    }

    const openTag = match[0];
    const tagEndPos = match.index + openTag.length;

    if (isSelfClosing) {
      matches.push({ openTag, innerHTML: "" });
      continue;
    }

    // Find the matching close tag by scanning further at the current logical
    // depth (we're already +1 inside this element).
    const closeTag = `</${fullTagName}>`;
    const nestedPattern = new RegExp(
      `<${escapeRegex(fullTagName)}(?:\\s[^>]*?)?(/?)>|${escapeRegex(closeTag)}`,
      "g"
    );
    nestedPattern.lastIndex = tagEndPos;
    let inner = 1;
    let foundEnd = -1;
    let innerMatch: RegExpExecArray | null;
    while ((innerMatch = nestedPattern.exec(xml)) !== null) {
      if (innerMatch[0] === closeTag) {
        inner--;
        if (inner === 0) {
          foundEnd = innerMatch.index;
          break;
        }
      } else if (innerMatch[1] !== "/") {
        inner++;
      }
    }

    if (foundEnd >= 0) {
      matches.push({ openTag, innerHTML: xml.substring(tagEndPos, foundEnd) });
      // Resync the outer scanner past the close tag and reset depth bookkeeping
      // (we're back at depth 0 in the outer scope).
      allTags.lastIndex = foundEnd + closeTag.length;
      depth = 0;
    } else {
      matches.push({ openTag, innerHTML: xml.substring(tagEndPos) });
      // No close tag — give up scanning the rest.
      break;
    }
  }

  if (matches.length === 0) {
    return null;
  }

  if (position === -1) {
    return matches[matches.length - 1];
  }
  if (position > 0) {
    return matches[position - 1] ?? null;
  }
  return matches[0];
}

function evaluateRepeatingXPath(binding: SdtDataBinding, xmlStore: Map<string, string>): string[] {
  const storeId = binding.storeItemId.replace(/[{}]/g, "").toLowerCase();
  const xmlContent = xmlStore.get(storeId) ?? xmlStore.get(`{${storeId}}`);
  if (!xmlContent) {
    return [];
  }

  // Find the parent element, then count children matching the last segment
  const xpath = binding.xpath;
  const segments = xpath.split("/").filter(s => s.length > 0);
  if (segments.length === 0) {
    return [];
  }

  const lastSeg = segments[segments.length - 1];
  const localName = lastSeg.includes(":") ? lastSeg.substring(lastSeg.indexOf(":") + 1) : lastSeg;
  const cleanName = localName.replace(/\[.*?\]/g, "");

  // Navigate to parent
  let parentContent = xmlContent;
  if (segments.length > 1) {
    const parentPath = "/" + segments.slice(0, -1).join("/");
    const parentEl = findElementByPath(xmlContent, parentPath);
    if (parentEl) {
      parentContent = parentEl.innerHTML;
    }
  }

  // Count direct-child occurrences of the element in the parent scope. This
  // mirrors the direct-child semantics enforced by `findNthElement`.
  const escapedName = escapeRegex(cleanName);
  const allTags = /<\/?([a-zA-Z_][a-zA-Z0-9_-]*(?::[a-zA-Z_][a-zA-Z0-9_-]*)?)(\s[^>]*?)?(\/?)>/g;
  let depth = 0;
  let count = 0;
  let m: RegExpExecArray | null;
  while ((m = allTags.exec(parentContent)) !== null) {
    const tagName = m[1];
    const isClosing = m[0].startsWith("</");
    const isSelfClosing = m[3] === "/";
    if (isClosing) {
      depth = Math.max(0, depth - 1);
      continue;
    }
    const local = tagName.includes(":") ? tagName.substring(tagName.indexOf(":") + 1) : tagName;
    if (depth === 0 && new RegExp(`^${escapedName}$`).test(local)) {
      count++;
    }
    if (!isSelfClosing) {
      depth++;
    }
  }

  // Return an array with one entry per occurrence (the index)
  return Array.from({ length: count }, (_, i) => String(i));
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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
    // Rewrite the xpath: if it starts with the parent path, add position predicate
    let resolvedXpath = innerBinding.xpath;
    if (resolvedXpath.startsWith(parentXpath)) {
      // Insert position predicate at the parent path level
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
