/**
 * OpenDoPE Data Binding API
 *
 * Resolve data bindings in a document by evaluating XPath expressions
 * against CustomXML parts and populating the SDT content with values.
 */

import { trimXmlPartLead } from "@utils/xml-part";
import {
  checkHandle,
  checkNamedValues,
  checkString,
  namedValueField
} from "@word/core/argument-checks";
import { markDerived } from "@word/core/ownership";
import {
  arrangeParagraphChildren,
  collectRangeMarkers,
  keepClonedRangeMarkers,
  repeatedRanges
} from "@word/core/range-markers";
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
 * Does not modify `doc`; the result may share unchanged nodes with it; `Document.*`
 * copies what it edits, so an edit made through one never reaches the other.
 *
 * @param doc - The document to process (returns a new modified copy).
 * @param data - Optional override data (storeItemId → XML string), as a `Map` or a
 *              plain object. If not provided, uses the document's embedded CustomXML parts.
 * @returns A new document with bindings resolved.
 */
export function resolveDataBindings(
  doc: DocxDocument,
  data?: ReadonlyMap<string, string> | Readonly<Record<string, string>>
): DocxDocument {
  const where = "Transform.resolveDataBindings";
  checkHandle(doc, "document", "doc", where);
  const entries = data === undefined ? undefined : checkNamedValues(data, "data", where);
  for (const [key, value] of entries ?? []) {
    checkString(value, namedValueField("data", key), where);
  }
  // Build a map of storeItemId → parsed XML content
  const xmlStore = new Map<string, string>();
  if (entries) {
    for (const [key, value] of entries) {
      xmlStore.set(key.toLowerCase(), value as string);
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
  return markDerived({ ...doc, body: newBody }, doc);
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
  if (value === UNSUPPORTED) {
    // An expression this evaluator cannot read leaves the SDT as it is.
    return sdt;
  }

  if (sdt.properties?.repeatingSection) {
    // Repeating section: value should be a collection
    // Each child in the data produces a copy of the repeatingSection item template
    const values = evaluateRepeatingXPath(binding, xmlStore);
    if (values === UNSUPPORTED) {
      return sdt;
    }
    if (values.length === 0) {
      // Remove the SDT entirely if no data
      return removedSdtParagraph(sdt);
    }
    // Return the SDT with repeated content, with inner bindings resolved per item
    const repeatedContent: (Paragraph | Run | Table)[] = [];
    const template = sdt.content;

    // Build parent xpath for resolving child element values
    const parentXpath = binding.xpath;
    const storeId = binding.storeItemId.replace(/[{}]/g, "").toLowerCase();
    const xmlContent = xmlStore.get(storeId) ?? xmlStore.get(`{${storeId}}`);

    const ranges = repeatedRanges(template);
    for (let itemIdx = 0; itemIdx < values.length; itemIdx++) {
      // Clone template content for each item
      // One item's copy of the template holds each range once across copies.
      const copy = keepClonedRangeMarkers(
        structuredClone(template),
        itemIdx,
        values.length,
        ranges
      );
      for (const item of copy) {
        const cloned = item as Paragraph | Run | Table;
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
    return removedSdtParagraph(sdt);
  }

  // Simple binding — replace text content of the SDT
  const { contentRangeMarkersAfter: _after, ...rest } = sdt;
  return { ...rest, content: [boundTextParagraph(sdt, value)] };
}

/**
 * The paragraph that replaces an SDT's content with `value`, keeping the
 * range markers of the old content in order around the text so every range
 * it opened or closed stays paired.
 */
function boundTextParagraph(
  sdt: Pick<StructuredDocumentTag, "content" | "contentRangeMarkersAfter">,
  value: string
): Paragraph {
  const { leading, trailing } = collectRangeMarkers(sdt.content, sdt.contentRangeMarkersAfter);
  const textRun: Run = { content: [{ type: "text", text: value }] };
  return { type: "paragraph", ...arrangeParagraphChildren([...leading, textRun, ...trailing]) };
}

/** The empty paragraph that stands for a removed SDT, holding the range markers it held. */
function removedSdtParagraph(sdt: StructuredDocumentTag): Paragraph {
  const { leading, trailing } = collectRangeMarkers([sdt]);
  return { type: "paragraph", ...arrangeParagraphChildren([...leading, ...trailing]) };
}

/**
 * The result of a binding whose XPath, or whose `w:prefixMappings`, is outside the subset this
 * evaluator reads. The SDT keeps its content: evaluating some other expression in its place would
 * put a value in it the binding never named.
 */
const UNSUPPORTED = Symbol("unsupported binding");
type Unsupported = typeof UNSUPPORTED;

/**
 * Evaluate an XPath expression against CustomXML stored as a string.
 *
 * Supports the XPath 1.0 subset that real-world OpenDoPE-style bindings
 * exercise:
 *   - Absolute location paths of child steps: `/root/child/grandchild` — what
 *     Word and OpenDoPE write
 *   - Namespace prefixes, resolved through the binding's `w:prefixMappings`;
 *     a prefix it does not map is matched by local name
 *   - Position predicates: `/root/item[1]`, `/root/item[last()]` (1-based, so
 *     `[0]` matches nothing)
 *   - A final attribute step: `/root/element/@attr`
 *   - Text content extraction from leaf nodes (CDATA-aware,
 *     entity-decoded by the XML parser)
 *
 * Anything else — other axes, functions, arithmetic, unions, other
 * predicates — is reported as {@link UNSUPPORTED}, never approximated.
 *
 * @returns The text content of the first matched element/attribute, or null if
 *          nothing matched. Returns the empty string for an element that exists
 *          but has no text — callers that need to distinguish empty from
 *          missing must compare to `null` explicitly.
 */
function evaluateSimpleXPath(
  binding: SdtDataBinding,
  xmlStore: Map<string, string>
): string | null | Unsupported {
  const evaluated = evaluateBinding(binding, xmlStore);
  if (evaluated === UNSUPPORTED || evaluated === null) {
    return evaluated;
  }
  const { path, chains, mappings } = evaluated;
  if (!path.attribute) {
    return chains.length > 0 ? collectTextContent(chains[0]!.at(-1)!) : null;
  }
  const wanted = path.attribute;
  const uri = wanted.prefix === undefined ? undefined : mappings.get(wanted.prefix);
  for (const chain of chains) {
    for (const [name, value] of Object.entries(chain.at(-1)!.attributes)) {
      const { prefix, local } = splitQName(name);
      if (local !== wanted.local || prefix === "xmlns" || name === "xmlns") {
        continue;
      }
      if (wanted.prefix === undefined) {
        // An unprefixed attribute step names an attribute in no namespace, and an unprefixed
        // attribute is the only one that is.
        if (prefix === undefined) {
          return value;
        }
      } else if (uri === undefined) {
        // A prefix the binding does not map is matched by local name, as it always was.
        return value;
      } else if (prefix !== undefined && namespaceOfPrefix(prefix, chain) === uri) {
        return value;
      }
    }
  }
  return null;
}

interface PathSegment {
  /** The step's prefix, when it has one. */
  readonly prefix?: string;
  readonly localName: string;
  /** 1-based position (`[0]` matches nothing); -1 = last(); absent without a predicate. */
  readonly position?: number;
}

interface ParsedPath {
  readonly segments: readonly PathSegment[];
  /** A final `@name` step. */
  readonly attribute?: { readonly prefix?: string; readonly local: string };
}

const NAME = "[A-Za-z_][\\w.-]*";
const QNAME = `${NAME}(?::${NAME})?`;
const STEP = new RegExp(`^/(${QNAME})(?:\\[(\\d+|last\\(\\))\\])?`);
const ATTRIBUTE_STEP = new RegExp(`^/@(${QNAME})$`);

/**
 * Parse the supported location-path subset. A path without a leading `/` is relative to the
 * document node, which is what a binding evaluates against, so it reads as the absolute path.
 */
function parsePath(xpath: string): ParsedPath | undefined {
  let rest = xpath.trim();
  if (rest.length > 0 && !rest.startsWith("/")) {
    rest = `/${rest}`;
  }
  const segments: PathSegment[] = [];
  while (rest.length > 0) {
    const attr = rest.match(ATTRIBUTE_STEP);
    if (attr) {
      return segments.length > 0 ? { segments, attribute: splitQName(attr[1]!) } : undefined;
    }
    const step = rest.match(STEP);
    if (!step) {
      return undefined;
    }
    const { prefix, local } = splitQName(step[1]!);
    const predicate = step[2];
    segments.push({
      ...(prefix !== undefined && { prefix }),
      localName: local,
      ...(predicate !== undefined && {
        position: predicate === "last()" ? -1 : parseInt(predicate, 10)
      })
    });
    rest = rest.slice(step[0].length);
  }
  return segments.length > 0 ? { segments } : undefined;
}

function splitQName(name: string): { prefix?: string; local: string } {
  const i = name.indexOf(":");
  return i >= 0 ? { prefix: name.substring(0, i), local: name.substring(i + 1) } : { local: name };
}

function elementLocal(el: XmlElement): string {
  return el.local ?? splitQName(el.name).local;
}

/**
 * The binding's `w:prefixMappings` — `xmlns:p="uri" xmlns:q='uri'` — as prefix → namespace: what
 * its XPath's prefixes stand for, whatever prefixes the data itself uses. They are XML namespace
 * declarations, so they are read by the XML parser (references decoded); `undefined` when they
 * are not well-formed declarations.
 */
function parsePrefixMappings(
  mappings: string | undefined
): ReadonlyMap<string, string> | undefined {
  const out = new Map<string, string>();
  if (mappings === undefined || mappings.trim() === "") {
    return out;
  }
  let root: XmlElement;
  try {
    root = parseXml(`<m ${mappings}/>`).root;
  } catch {
    return undefined;
  }
  for (const [name, value] of Object.entries(root.attributes)) {
    const { prefix, local } = splitQName(name);
    if (prefix !== "xmlns") {
      // Only namespace declarations belong in prefixMappings.
      return undefined;
    }
    out.set(local, value);
  }
  return out;
}

/** The namespace `prefix` is bound to on the last element of `chain`, from the root down. */
function namespaceOfPrefix(prefix: string, chain: readonly XmlElement[]): string | undefined {
  for (let i = chain.length - 1; i >= 0; i--) {
    const uri = chain[i]!.ns?.[prefix];
    if (uri !== undefined) {
      return uri;
    }
  }
  return undefined;
}

interface StoreRoot {
  readonly root: XmlElement;
  /** Whether the store's namespaces resolved, so elements carry the URI they are in. */
  readonly namespaced: boolean;
}

/**
 * Whether `el` is what `seg` names: a prefix the binding maps names a namespace, and the element
 * must be in it — the data may bind any prefix to it, or none. A store whose namespaces did not
 * resolve cannot show that, so it matches no mapped step. A step whose prefix the binding does
 * not map, or that has none, is matched by local name, as it always was.
 */
function matchesSegment(
  el: XmlElement,
  seg: PathSegment,
  mappings: ReadonlyMap<string, string>,
  namespaced: boolean
): boolean {
  if (elementLocal(el) !== seg.localName) {
    return false;
  }
  const uri = seg.prefix === undefined ? undefined : mappings.get(seg.prefix);
  return uri === undefined || (namespaced && el.uri === uri);
}

function loadStoreRoot(binding: SdtDataBinding, xmlStore: Map<string, string>): StoreRoot | null {
  const storeId = binding.storeItemId.replace(/[{}]/g, "").toLowerCase();
  const xmlContent = xmlStore.get(storeId) ?? xmlStore.get(`{${storeId}}`);
  if (!xmlContent) {
    return null;
  }
  const text = trimXmlPartLead(xmlContent);
  try {
    // With namespaces, so a mapped prefix is matched by the namespace it stands for.
    return { root: parseXml(text, { xmlns: true }).root, namespaced: true };
  } catch {
    try {
      // Data whose namespaces do not resolve is still read by local name, but no mapped prefix
      // can match in it (see `matchesSegment`).
      return { root: parseXml(text).root, namespaced: false };
    } catch {
      // Malformed CustomXML store. Treat as empty rather than throwing —
      // resolveDataBindings is best-effort by design.
      return null;
    }
  }
}

/**
 * Evaluate the binding's path: every matching element, in document order, as the chain of
 * elements from the root down to it. `null` when there is no store to evaluate against.
 */
function evaluateBinding(
  binding: SdtDataBinding,
  xmlStore: Map<string, string>
):
  | { path: ParsedPath; chains: XmlElement[][]; mappings: ReadonlyMap<string, string> }
  | null
  | Unsupported {
  const path = parsePath(binding.xpath);
  const mappings = parsePrefixMappings(binding.prefixMappings);
  if (!path || !mappings) {
    return UNSUPPORTED;
  }
  const store = loadStoreRoot(binding, xmlStore);
  if (!store) {
    return null;
  }
  // The context starts at the document node, whose only child is the root element.
  let contexts: XmlElement[][] = [[]];
  for (const seg of path.segments) {
    // Each context's children are its own, so no element is reached twice.
    const next: XmlElement[][] = [];
    for (const origin of contexts) {
      const children = origin.length === 0 ? [store.root] : elementChildren(origin.at(-1)!);
      const matches = children.filter(c => matchesSegment(c, seg, mappings, store.namespaced));
      const picked =
        seg.position === undefined
          ? matches
          : [seg.position === -1 ? matches.at(-1) : matches[seg.position - 1]].filter(
              (m): m is XmlElement => m !== undefined
            );
      for (const el of picked) {
        next.push([...origin, el]);
      }
    }
    contexts = next;
  }
  return { path, chains: contexts, mappings };
}

function elementChildren(el: XmlElement): XmlElement[] {
  return el.children.filter((c): c is XmlElement => c.type === "element");
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

/** One entry per element a repeating section's XPath matches. */
function evaluateRepeatingXPath(
  binding: SdtDataBinding,
  xmlStore: Map<string, string>
): string[] | Unsupported {
  const evaluated = evaluateBinding(binding, xmlStore);
  if (evaluated === UNSUPPORTED) {
    return UNSUPPORTED;
  }
  if (evaluated === null || evaluated.path.attribute) {
    return [];
  }
  return evaluated.chains.map((_, i) => String(i));
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
    if (value !== null && value !== UNSUPPORTED) {
      obj.content = [boundTextParagraph(obj as unknown as StructuredDocumentTag, value)];
      delete obj.contentRangeMarkersAfter;
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
