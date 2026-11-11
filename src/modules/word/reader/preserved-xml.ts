/**
 * DOCX Reader - markup the model does not represent.
 *
 * A part the writer regenerates from the model (`settings.xml`, `styles.xml`,
 * the font table, the theme, `app.xml`) loses every element the model has no
 * field for unless the reader keeps it. These helpers keep such elements
 * verbatim — each serialized so it declares the namespace prefixes it uses,
 * since it will be written under a root the writer builds — together with the
 * root's namespace declarations and `mc:Ignorable`.
 */

import { WORD_NAMESPACES } from "@word/reader/namespaces";
import { serializeElement } from "@word/reader/parse-utils";
import { parseContentPartXml } from "@word/reader/part-xml";
import type { PreservedXml } from "@word/types";
import type { XmlDocument, XmlElement } from "@xml/types";

/** Namespace declarations made on `el` itself, prefix → URI (`""` = default). */
function declaredNamespaces(el: XmlElement): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, value] of Object.entries(el.attributes)) {
    if (name === "xmlns") {
      out[""] = value;
    } else if (name.startsWith("xmlns:")) {
      out[name.slice(6)] = value;
    }
  }
  return out;
}

/** The namespaces in scope inside `el`, given those in scope around it. */
export function namespacesInScope(
  outer: Readonly<Record<string, string>>,
  el: XmlElement
): Record<string, string> {
  return { ...outer, ...declaredNamespaces(el) };
}

function prefixOf(qualifiedName: string): string {
  const colon = qualifiedName.indexOf(":");
  return colon === -1 ? "" : qualifiedName.slice(0, colon);
}

function collectPrefixes(el: XmlElement, out: Set<string>): void {
  out.add(prefixOf(el.name));
  for (const name of Object.keys(el.attributes)) {
    // An unprefixed attribute is in no namespace, so it needs no declaration.
    if (name !== "xmlns" && !name.startsWith("xmlns:") && name.includes(":")) {
      out.add(prefixOf(name));
    }
  }
  for (const child of el.children) {
    if (child.type === "element") {
      collectPrefixes(child, out);
    }
  }
}

/** Serialize `el` with a declaration for every in-scope prefix it uses. */
export function serializeSelfContained(
  el: XmlElement,
  inScope: Readonly<Record<string, string>>
): string {
  const used = new Set<string>();
  collectPrefixes(el, used);
  const own = declaredNamespaces(el);
  const added: Record<string, string> = {};
  for (const prefix of used) {
    if (prefix === "xml" || prefix in own) {
      continue;
    }
    const uri = inScope[prefix];
    if (uri !== undefined) {
      added[prefix === "" ? "xmlns" : `xmlns:${prefix}`] = uri;
    }
  }
  if (Object.keys(added).length === 0) {
    return serializeElement(el);
  }
  return serializeElement({
    ...el,
    attributes: { ...added, ...el.attributes }
  });
}

/** The child elements of `el` the model does not represent, serialized self-contained. */
export function unmodelledChildren(
  el: XmlElement,
  inScope: Readonly<Record<string, string>>,
  isModelled: (child: XmlElement) => boolean
): string[] {
  const scope = namespacesInScope(inScope, el);
  const out: string[] = [];
  for (const child of el.children) {
    if (child.type === "element" && !isModelled(child)) {
      out.push(serializeSelfContained(child, scope));
    }
  }
  return out;
}

/** A part root's namespace declarations, `mc:Ignorable` and unmodelled children. */
export function preservedRootXml(
  root: XmlElement,
  isModelled: (child: XmlElement) => boolean
): PreservedXml | undefined {
  const namespaces = declaredNamespaces(root);
  const elements = unmodelledChildren(root, {}, isModelled);
  const ignorable = Object.entries(root.attributes).find(
    ([name]) => name.endsWith(":Ignorable") && namespaces[prefixOf(name)] !== undefined
  )?.[1];
  if (elements.length === 0 && ignorable === undefined && Object.keys(namespaces).length === 0) {
    return undefined;
  }
  return {
    ...(Object.keys(namespaces).length > 0 ? { namespaces } : {}),
    ...(ignorable !== undefined ? { ignorable } : {}),
    ...(elements.length > 0 ? { elements } : {})
  };
}

/**
 * A content part (`parseContentPartXml`) whose elements each declare the namespaces of other
 * vocabularies they use ({@link bindInheritedNamespaces}): what the reader keeps of it verbatim is
 * written back bound.
 */
export function parseBoundContentPart(xml: string, maxDepth: number | undefined): XmlDocument {
  const doc = parseContentPartXml(xml, maxDepth);
  bindInheritedNamespaces(doc.root, xml);
  return doc;
}

/**
 * Give each element of `root`'s tree the declarations it relies on from its ancestors for the
 * prefixes the reader does not know — and the default namespace — so that any subtree kept verbatim
 * (an opaque run or paragraph child, preserved markup) still binds them once it is written back
 * under elements the writer builds, which declare only Word's own prefixes. `outer` is what is in
 * scope around `root`. Only the element whose own name or attribute uses such a prefix is given one;
 * its descendants inherit it from there.
 *
 * Skipped unless `text` — the part `root` was read from — declares a prefix the reader does not
 * know or a default namespace: every part Word writes declares only its own.
 */
export function bindInheritedNamespaces(
  root: XmlElement,
  text: string,
  outer: Readonly<Record<string, string>> = {}
): void {
  if (Object.keys(outer).length === 0 && !declaresForeignNamespace(text)) {
    return;
  }
  bindWithin(root, outer);
}

/**
 * {@link bindInheritedNamespaces} for a subtree read on its own — a body element the streaming
 * reader builds — with `outer` the declarations in scope around it ({@link foreignDeclarations}).
 */
export function bindSubtreeNamespaces(
  root: XmlElement,
  outer: Readonly<Record<string, string>>
): void {
  bindWithin(root, outer);
}

/** The declarations `attributes` make of prefixes the reader does not know, or of the default namespace. */
export function foreignDeclarations(
  attributes: Readonly<Record<string, string>>
): Record<string, string> | undefined {
  let out: Record<string, string> | undefined;
  for (const name in attributes) {
    const prefix = name === "xmlns" ? "" : name.startsWith("xmlns:") ? name.slice(6) : undefined;
    if (prefix !== undefined && isForeignPrefix(prefix)) {
      (out ??= {})[prefix] = attributes[name]!;
    }
  }
  return out;
}

const XMLNS_DECLARATION = /\sxmlns(?::([^\s=/>]+))?\s*=/g;

function declaresForeignNamespace(text: string): boolean {
  if (!text.includes("xmlns")) {
    return false;
  }
  XMLNS_DECLARATION.lastIndex = 0;
  for (let m = XMLNS_DECLARATION.exec(text); m !== null; m = XMLNS_DECLARATION.exec(text)) {
    if (isForeignPrefix(m[1] ?? "")) {
      return true;
    }
  }
  return false;
}

function isForeignPrefix(prefix: string): boolean {
  return prefix === "" || (prefix !== "xml" && !WORD_NAMESPACES.prefixes.has(prefix));
}

function bindWithin(el: XmlElement, outer: Readonly<Record<string, string>>): void {
  const own = declaredNamespaces(el);
  let added: Record<string, string> | undefined;
  const need = (prefix: string): void => {
    if (!isForeignPrefix(prefix) || prefix in own) {
      return;
    }
    const uri = outer[prefix];
    if (uri !== undefined) {
      (added ??= {})[prefix === "" ? "xmlns" : `xmlns:${prefix}`] = uri;
    }
  };
  need(prefixOf(el.name));
  for (const name of Object.keys(el.attributes)) {
    if (name !== "xmlns" && !name.startsWith("xmlns:") && name.includes(":")) {
      need(prefixOf(name));
    }
  }
  if (added !== undefined) {
    const attributes = Object.assign(Object.create(null) as Record<string, string>, added);
    (el as { attributes: Record<string, string> }).attributes = Object.assign(
      attributes,
      el.attributes
    );
  }
  let scope: Readonly<Record<string, string>> | undefined;
  for (const child of el.children) {
    if (child.type === "element") {
      scope ??= childScope(el, outer, own);
      bindWithin(child, scope);
    }
  }
}

/**
 * The scope `el`'s children are bound against. Under an element of a namespace the reader does not
 * know, the declarations it carries are left out: its children are only ever written back inside it.
 */
function childScope(
  el: XmlElement,
  outer: Readonly<Record<string, string>>,
  own: Readonly<Record<string, string>>
): Readonly<Record<string, string>> {
  const foreign = isForeignPrefix(prefixOf(el.name));
  if (!foreign && Object.keys(own).length === 0) {
    return outer;
  }
  const scope: Record<string, string> = { ...outer, ...own };
  if (foreign) {
    for (const name of Object.keys(el.attributes)) {
      if (name === "xmlns") {
        delete scope[""];
      } else if (name.startsWith("xmlns:")) {
        delete scope[name.slice(6)];
      }
    }
  }
  return scope;
}

/**
 * The local name of a WordprocessingML element (`w:` prefix or none), or
 * `undefined` for an element of another namespace. Matches how the parsers
 * find children (`findChildNs`).
 */
export function wordLocalName(el: XmlElement): string | undefined {
  if (el.name.startsWith("w:")) {
    return el.name.slice(2);
  }
  return el.name.includes(":") ? undefined : el.name;
}
