/**
 * Editing the document properties of an existing XMP packet.
 *
 * When an edit changes the document information dictionary, the XMP packet beside it has to
 * say the same thing: ISO 32000-2 §14.3 makes the packet the primary record and a reader is
 * free to believe it over `/Info`, so changing only `/Info` leaves a file that answers "what
 * is the title?" two ways.
 *
 * The packet is parsed with the `xml` module and the tree is edited, then written back. A
 * regex editor came first and was wrong four ways: it wrote prefixes that were not bound
 * where they were used, added a property beside the attribute form of the same property,
 * put a document property inside a nested `xmpMM:History` event, and edited text inside
 * comments. Only the properties being set are touched; every other schema, comment,
 * processing instruction and language alternative is written back as it was read.
 *
 * A packet this cannot decode or recognise — not text in its encoding, not well-formed
 * namespaced XML, no `rdf:RDF` — is reported as `undefined`, and the caller keeps the
 * original metadata stream untouched rather than writing a guess.
 */

import { xmlEncode, xmlEncodeAttr } from "@utils/xml-encode";
import { parseXml } from "@xml/dom";
import type { XmlDocument, XmlElement, XmlNode } from "@xml/types";

/**
 * The properties an edit changes, as `DocumentMetadata` names them: a value sets the
 * property, `null` removes it, absent leaves it as it is.
 */
export interface XmpProperties {
  readonly title?: string | null;
  readonly author?: string | null;
  readonly subject?: string | null;
  readonly keywords?: string | null;
  readonly creator?: string | null;
  readonly producer?: string | null;
  readonly creationDate?: Date | null;
  readonly modDate?: Date | null;
}

const NS = {
  x: "adobe:ns:meta/",
  rdf: "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
  dc: "http://purl.org/dc/elements/1.1/",
  xmp: "http://ns.adobe.com/xap/1.0/",
  pdf: "http://ns.adobe.com/pdf/1.3/"
} as const;

type Shape = "alt" | "seq" | "text";

/** `DocumentMetadata` field → XMP property: namespace, local name, value shape. */
const PROPERTIES: ReadonlyArray<{
  readonly field: keyof XmpProperties;
  readonly ns: "dc" | "xmp" | "pdf";
  readonly name: string;
  readonly shape: Shape;
}> = [
  { field: "title", ns: "dc", name: "title", shape: "alt" },
  { field: "author", ns: "dc", name: "creator", shape: "seq" },
  { field: "subject", ns: "dc", name: "description", shape: "alt" },
  { field: "keywords", ns: "pdf", name: "Keywords", shape: "text" },
  { field: "creator", ns: "xmp", name: "CreatorTool", shape: "text" },
  { field: "producer", ns: "pdf", name: "Producer", shape: "text" },
  { field: "creationDate", ns: "xmp", name: "CreateDate", shape: "text" },
  { field: "modDate", ns: "xmp", name: "ModifyDate", shape: "text" }
];

// =============================================================================
// Encoding
// =============================================================================

/** How a packet's bytes encode its text, so an edit is written back the same way. */
export interface XmpEncoding {
  readonly name: "utf-8" | "utf-16be" | "utf-16le";
  /** Whether the bytes begin with a byte-order mark. */
  readonly bom: boolean;
}

/**
 * Decode a packet by its byte-order mark — or, without one, by the `<` that starts it
 * (`00 3C` is UTF-16BE, `3C 00` UTF-16LE), as XMP §7.3 describes. `undefined` when the
 * bytes are not valid text in that encoding.
 */
export function decodeXmpPacket(
  bytes: Uint8Array
): { text: string; encoding: XmpEncoding } | undefined {
  let encoding: XmpEncoding;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    encoding = { name: "utf-8", bom: true };
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    encoding = { name: "utf-16be", bom: true };
  } else if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    encoding = { name: "utf-16le", bom: true };
  } else if (bytes[0] === 0x00 && bytes[1] === 0x3c) {
    encoding = { name: "utf-16be", bom: false };
  } else if (bytes[0] === 0x3c && bytes[1] === 0x00) {
    encoding = { name: "utf-16le", bom: false };
  } else {
    encoding = { name: "utf-8", bom: false };
  }
  try {
    // `fatal`, so bytes that are not text — compressed data a filter failed to inflate —
    // are refused rather than read as replacement characters. The BOM is dropped.
    const text = new TextDecoder(encoding.name, { fatal: true }).decode(bytes);
    return { text, encoding };
  } catch {
    return undefined;
  }
}

/** Encode `text` as {@link decodeXmpPacket} found the original encoded. */
export function encodeXmpPacket(text: string, encoding: XmpEncoding): Uint8Array {
  if (encoding.name === "utf-8") {
    const body = new TextEncoder().encode(text);
    if (!encoding.bom) {
      return body;
    }
    const out = new Uint8Array(body.length + 3);
    out.set([0xef, 0xbb, 0xbf]);
    out.set(body, 3);
    return out;
  }
  const offset = encoding.bom ? 2 : 0;
  const out = new Uint8Array(offset + text.length * 2);
  const view = new DataView(out.buffer);
  const littleEndian = encoding.name === "utf-16le";
  if (encoding.bom) {
    view.setUint16(0, 0xfeff, littleEndian);
  }
  for (let i = 0; i < text.length; i++) {
    view.setUint16(offset + i * 2, text.charCodeAt(i), littleEndian);
  }
  return out;
}

// =============================================================================
// Tree helpers
// =============================================================================

type Scope = ReadonlyMap<string, string>;

function scopeOf(parent: Scope, element: XmlElement): Scope {
  if (element.ns === undefined || Object.keys(element.ns).length === 0) {
    return parent;
  }
  const scope = new Map(parent);
  for (const [prefix, uri] of Object.entries(element.ns)) {
    scope.set(prefix, uri);
  }
  return scope;
}

function prefixFor(scope: Scope, uri: string): string | undefined {
  for (const [prefix, bound] of scope) {
    if (bound === uri && prefix !== "") {
      // The nearest binding wins; a later shadowing of the same prefix would have replaced it.
      if (scope.get(prefix) === uri) {
        return prefix;
      }
    }
  }
  return undefined;
}

function isElement(node: XmlNode): node is XmlElement {
  return node.type === "element";
}

function is(element: XmlElement, uri: string, local: string): boolean {
  return element.uri === uri && element.local === local;
}

function elementChildren(element: XmlElement): XmlElement[] {
  return element.children.filter(isElement);
}

function textNode(value: string): XmlNode {
  return { type: "text", value } as XmlNode;
}

function newElement(
  name: string,
  attributes: Record<string, string>,
  children: XmlNode[],
  uri: string,
  ns?: Record<string, string>
): XmlElement {
  const [prefix, local] = name.split(":");
  return {
    type: "element",
    name,
    attributes,
    children,
    prefix,
    local,
    uri,
    ...(ns === undefined ? {} : { ns })
  };
}

// =============================================================================
// Serialisation
// =============================================================================

function serializeNode(node: XmlNode, out: string[]): void {
  switch (node.type) {
    case "text":
      out.push(xmlEncode(node.value));
      return;
    case "cdata":
      // `]]>` cannot sit inside a CDATA section; it was not in one when read.
      out.push(`<![CDATA[${node.value}]]>`);
      return;
    case "comment":
      out.push(`<!--${node.value}-->`);
      return;
    case "processing-instruction":
      out.push(`<?${node.target}${node.body === "" ? "" : ` ${node.body}`}?>`);
      return;
    case "element": {
      out.push(`<${node.name}`);
      for (const [key, value] of Object.entries(node.attributes)) {
        out.push(` ${key}="${xmlEncodeAttr(value)}"`);
      }
      if (node.children.length === 0) {
        out.push("/>");
        return;
      }
      out.push(">");
      for (const child of node.children) {
        serializeNode(child, out);
      }
      out.push(`</${node.name}>`);
    }
  }
}

function serializeDocument(doc: XmlDocument): string {
  const out: string[] = [];
  if (doc.declaration !== undefined) {
    out.push(
      `<?xml${Object.entries(doc.declaration)
        .map(([key, value]) => ` ${key}="${xmlEncodeAttr(value)}"`)
        .join("")}?>\n`
    );
  }
  for (const node of doc.prologue) {
    serializeNode(node, out);
    out.push("\n");
  }
  serializeNode(doc.root, out);
  for (const node of doc.epilogue ?? []) {
    out.push("\n");
    serializeNode(node, out);
  }
  return out.join("");
}

// =============================================================================
// Editing
// =============================================================================

/** An XMP date: ISO 8601 to the second, in UTC. */
function xmpDate(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
}

/**
 * `dc:creator` is an ordered list of authors; `/Author` is one string. The list is written
 * from `author` split on `;` — the separator Acrobat uses between the two — with each name
 * trimmed and empty names dropped, so `"Ann; Bob"` becomes two entries and `"Ann"` one.
 */
function authorsOf(author: string): string[] {
  const names = author
    .split(";")
    .map(name => name.trim())
    .filter(name => name !== "");
  return names.length === 0 ? [author] : names;
}

interface Description {
  readonly element: XmlElement;
  readonly scope: Scope;
}

/** Parse a packet and find its `rdf:RDF` — the root, or a child of `x:xmpmeta`. */
function parsePacket(
  xml: string
): { doc: XmlDocument; rdf: XmlElement; rdfScope: Scope } | undefined {
  let doc: XmlDocument;
  try {
    doc = parseXml(xml.startsWith("\uFEFF") ? xml.slice(1) : xml, {
      xmlns: true,
      comments: true,
      processingInstructions: true,
      cdataAsNodes: true
    });
  } catch {
    return undefined;
  }
  const rootScope = scopeOf(new Map(), doc.root);
  if (is(doc.root, NS.rdf, "RDF")) {
    return { doc, rdf: doc.root, rdfScope: rootScope };
  }
  if (is(doc.root, NS.x, "xmpmeta") || is(doc.root, NS.x, "xapmeta")) {
    const rdf = elementChildren(doc.root).find(child => is(child, NS.rdf, "RDF"));
    if (rdf !== undefined) {
      return { doc, rdf, rdfScope: scopeOf(rootScope, rdf) };
    }
  }
  return undefined;
}

/** The text of an element: its text and CDATA, concatenated. */
function textOf(element: XmlElement): string {
  return element.children
    .map(child => (child.type === "text" || child.type === "cdata" ? child.value : ""))
    .join("");
}

/** The XMP properties `Pdf.extract` reports, by `DocumentMetadata`-style name. */
export interface XmpReadProperties {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  creator?: string;
  producer?: string;
  creationDate?: string;
  modDate?: string;
}

/**
 * Read the document properties of a packet: from the top-level `rdf:Description`s only,
 * in element or attribute form, with character references decoded by the parser. A
 * language alternative reads its `x-default` entry (the first entry when it has none); an
 * ordered or unordered list reads its entries joined with `"; "`. `undefined` when the
 * packet is not namespaced XML with an `rdf:RDF`.
 */
export function readXmpProperties(xml: string): XmpReadProperties | undefined {
  const parsed = parsePacket(xml);
  if (parsed === undefined) {
    return undefined;
  }
  const wanted: ReadonlyArray<[keyof XmpReadProperties, string, string]> = [
    ["title", NS.dc, "title"],
    ["author", NS.dc, "creator"],
    ["subject", NS.dc, "description"],
    ["keywords", NS.pdf, "Keywords"],
    ["creator", NS.xmp, "CreatorTool"],
    ["producer", NS.pdf, "Producer"],
    ["creationDate", NS.xmp, "CreateDate"],
    ["modDate", NS.xmp, "ModifyDate"]
  ];
  const out: XmpReadProperties = {};
  for (const description of elementChildren(parsed.rdf).filter(child =>
    is(child, NS.rdf, "Description")
  )) {
    const scope = scopeOf(parsed.rdfScope, description);
    for (const [field, uri, local] of wanted) {
      if (out[field] !== undefined) {
        continue;
      }
      const element = elementChildren(description).find(child => is(child, uri, local));
      if (element !== undefined) {
        const container = elementChildren(element).find(
          child => is(child, NS.rdf, "Alt") || is(child, NS.rdf, "Seq") || is(child, NS.rdf, "Bag")
        );
        if (container === undefined) {
          out[field] = textOf(element).trim();
          continue;
        }
        const items = elementChildren(container).filter(child => is(child, NS.rdf, "li"));
        if (container.local === "Alt") {
          const item = items.find(li => li.attributes["xml:lang"] === "x-default") ?? items[0];
          if (item !== undefined) {
            out[field] = textOf(item).trim();
          }
        } else if (items.length > 0) {
          out[field] = items.map(li => textOf(li).trim()).join("; ");
        }
        continue;
      }
      for (const [name, value] of Object.entries(description.attributes)) {
        const [prefix, attrLocal] = name.includes(":") ? name.split(":") : ["", name];
        if (prefix !== "" && attrLocal === local && scope.get(prefix) === uri) {
          out[field] = value;
        }
      }
    }
  }
  return out;
}

/**
 * The packet with each property in `properties` set — or, for `null`, removed from every
 * top-level description in every form (element, attribute) — and `xmp:MetadataDate` set to
 * `metadataDate` — XMP requires it to change whenever the metadata does. `undefined` when
 * the packet is not well-formed namespaced XML with an `rdf:RDF`.
 *
 * Where a property goes and how it is written:
 * - Only the top-level `rdf:Description`s — the children of `rdf:RDF` — are the document's.
 *   A same-named property inside a structure (a history event, a nested description) is
 *   not the document's and is never touched.
 * - A property already present is changed where it is: an element keeps its start tag (and
 *   so its namespace declarations); an attribute keeps its place. A language alternative
 *   (`dc:title`, `dc:description`) replaces only its `x-default` entry, adding one first
 *   when it has none, so the other languages survive. `dc:creator` is rewritten from the
 *   author string (see {@link authorsOf}).
 * - A property that is absent is added to the top-level description that already has its
 *   namespace in scope — or the first one — declaring the namespace on the new element
 *   when no prefix is bound to it there. A packet with no top-level description gets one.
 */
export function updateXmpPacket(
  xml: string,
  properties: XmpProperties,
  metadataDate: Date
): string | undefined {
  const parsed = parsePacket(xml);
  if (parsed === undefined) {
    return undefined;
  }
  const { doc, rdf, rdfScope } = parsed;

  const descriptions: Description[] = elementChildren(rdf)
    .filter(child => is(child, NS.rdf, "Description"))
    .map(element => ({ element, scope: scopeOf(rdfScope, element) }));
  if (descriptions.length === 0) {
    const rdfPrefix = prefixFor(rdfScope, NS.rdf) ?? "rdf";
    const element = newElement(
      `${rdfPrefix}:Description`,
      { [`${rdfPrefix}:about`]: "" },
      [],
      NS.rdf
    );
    rdf.children.push(element);
    descriptions.push({ element, scope: rdfScope });
  }

  const set = (ns: "dc" | "xmp" | "pdf", local: string, shape: Shape, value: string): void => {
    const uri = NS[ns];
    // Element form, on any top-level description.
    for (const { element, scope } of descriptions) {
      for (const child of elementChildren(element)) {
        if (is(child, uri, local)) {
          writeValue(child, scopeOf(scope, child), shape, value);
          return;
        }
      }
    }
    // Attribute form — a simple value only; a structured one cannot be an attribute, so an
    // attribute of that name is replaced by the element form below.
    for (const { element, scope } of descriptions) {
      for (const name of Object.keys(element.attributes)) {
        const [prefix, attrLocal] = name.includes(":") ? name.split(":") : ["", name];
        if (prefix === "" || attrLocal !== local || scope.get(prefix) !== uri) {
          continue;
        }
        if (shape === "text") {
          element.attributes[name] = value;
          return;
        }
        delete element.attributes[name];
      }
    }
    // Absent: add it where the namespace is already in scope, or to the first description.
    const target =
      descriptions.find(({ scope }) => prefixFor(scope, uri) !== undefined) ?? descriptions[0];
    const bound = prefixFor(target.scope, uri);
    const prefix = bound ?? (target.scope.has(ns) ? `${ns}1` : ns);
    const declaration = bound === undefined ? { [`xmlns:${prefix}`]: uri } : {};
    const added = newElement(
      `${prefix}:${local}`,
      declaration,
      [],
      uri,
      bound === undefined ? { [prefix]: uri } : undefined
    );
    writeValue(added, scopeOf(target.scope, added), shape, value);
    target.element.children.push(added);
  };

  // Remove a property from every top-level description, in every form it takes there: the
  // element (with its container, every language and every entry) and the attribute.
  const remove = (ns: "dc" | "xmp" | "pdf", local: string): void => {
    const uri = NS[ns];
    for (const { element, scope } of descriptions) {
      element.children = element.children.filter(
        child => !(isElement(child) && is(child, uri, local))
      );
      for (const name of Object.keys(element.attributes)) {
        const [prefix, attrLocal] = name.includes(":") ? name.split(":") : ["", name];
        if (prefix !== "" && attrLocal === local && scope.get(prefix) === uri) {
          delete element.attributes[name];
        }
      }
    }
  };

  for (const { field, ns, name, shape } of PROPERTIES) {
    const raw = properties[field];
    if (raw === null) {
      remove(ns, name);
    } else if (raw !== undefined) {
      set(ns, name, shape, raw instanceof Date ? xmpDate(raw) : raw);
    }
  }
  set("xmp", "MetadataDate", "text", xmpDate(metadataDate));
  return serializeDocument(doc);
}

/** Write `value` into property element `element` in `shape`. */
function writeValue(element: XmlElement, scope: Scope, shape: Shape, value: string): void {
  if (shape === "text") {
    element.children = [textNode(value)];
    return;
  }
  const rdfPrefix = prefixFor(scope, NS.rdf);
  // An RDF container element needs the rdf prefix in scope; declare it when it is not.
  const prefix = rdfPrefix ?? "rdf";
  const declare = rdfPrefix === undefined ? { [`xmlns:${prefix}`]: NS.rdf } : {};
  const container = shape === "alt" ? "Alt" : "Seq";
  const existing = elementChildren(element).find(child => is(child, NS.rdf, container));
  if (shape === "alt" && existing !== undefined) {
    const liPrefix = existing.prefix ?? prefix;
    const items = elementChildren(existing).filter(child => is(child, NS.rdf, "li"));
    const fallback = items.find(item => item.attributes["xml:lang"] === "x-default");
    if (fallback !== undefined) {
      fallback.children = [textNode(value)];
      return;
    }
    // XMP puts x-default first.
    existing.children.unshift(
      newElement(`${liPrefix}:li`, { "xml:lang": "x-default" }, [textNode(value)], NS.rdf)
    );
    return;
  }
  const items =
    shape === "alt"
      ? [newElement(`${prefix}:li`, { "xml:lang": "x-default" }, [textNode(value)], NS.rdf)]
      : authorsOf(value).map(name => newElement(`${prefix}:li`, {}, [textNode(name)], NS.rdf));
  if (shape === "seq" && existing !== undefined) {
    existing.children = items;
    return;
  }
  element.children = [newElement(`${prefix}:${container}`, declare, items, NS.rdf)];
}
