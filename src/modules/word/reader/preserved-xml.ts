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

import { serializeElement } from "@word/reader/parse-utils";
import type { PreservedXml } from "@word/types";
import type { XmlElement } from "@xml/types";

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
  return serializeElement({ ...el, attributes: { ...added, ...el.attributes } });
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
