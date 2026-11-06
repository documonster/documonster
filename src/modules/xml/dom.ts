/**
 * DOM XML Parser
 *
 * Builds an in-memory XML tree ({@link XmlDocument} / {@link XmlElement})
 * from an XML string. Built on top of {@link SaxParser} — no duplicate
 * parsing logic.
 *
 * For large documents, prefer SAX-based streaming.
 * This is intended for small-to-medium XML where tree access is convenient.
 */

import { check, s } from "@utils/schema";
import {
  checkElement,
  checkFunctionArgument,
  checkParseOptions,
  checkStringArgument,
  checkXmlInput,
  xmlArgument
} from "@xml/arguments";
import { XmlParseError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import { resolveOptions, resolveValue, addChildValue } from "@xml/to-object-shared";
import type {
  XmlToPlainObjectOptions,
  XmlCData,
  XmlComment,
  XmlDocument,
  XmlElement,
  XmlNode,
  XmlParseOptions,
  XmlProcessingInstruction,
  XmlText
} from "@xml/types";

// =============================================================================
// Security
// =============================================================================

/**
 * Copy an attribute map into one with no prototype. Every name is kept: on a null-prototype
 * object `constructor` or `__proto__` is an ordinary property, so filtering them protected
 * nothing and silently dropped the document's attributes.
 */
function safeAttributes(raw: Record<string, string>): Record<string, string> {
  const attrs: Record<string, string> = Object.create(null);
  for (const key in raw) {
    if (Object.prototype.hasOwnProperty.call(raw, key)) {
      attrs[key] = raw[key];
    }
  }
  return attrs;
}

// =============================================================================
// Element Factory
// =============================================================================

function createElement(name: string, attributes: Record<string, string>): XmlElement {
  return {
    type: "element",
    name,
    attributes: safeAttributes(attributes),
    children: []
  };
}

/** Append text to a parent element, merging with the last child if it is a text node. */
function appendText(parent: XmlElement, text: string): void {
  const lastChild = parent.children[parent.children.length - 1];
  if (lastChild && lastChild.type === "text") {
    (lastChild as XmlText).value += text;
  } else {
    parent.children.push({ type: "text", value: text } as XmlText);
  }
}

// =============================================================================
// parseXml
// =============================================================================

/**
 * Parse an XML string into a DOM tree.
 *
 * @param xml - Complete XML string.
 * @param options - Parse options.
 * @returns The parsed XML document.
 * @throws {XmlParseError} If the XML is malformed.
 *
 * **Fragment mode** (`{ fragment: true }`): suppresses the "multiple root
 * elements" error. The returned `XmlDocument.root` is the first root element,
 * and `XmlDocument.roots` contains all root-level elements.
 *
 * **Lenient about the declaration**: a `<?xml …?>` that is not at the very start (after
 * whitespace, a second one, one inside an element) is dropped rather than refused, because
 * this parser reads files other programs wrote. The writers (`Xml.Writer`,
 * `Xml.StreamWriter`) never produce one.
 *
 * @example
 * ```ts
 * const doc = parseXml('<root><child id="1">text</child></root>');
 * console.log(doc.root.name); // "root"
 * console.log(doc.root.children[0].name); // "child"
 * ```
 */
function parseXml(xml: string, options?: XmlParseOptions): XmlDocument {
  checkXmlInput("Xml.parse", xml);
  checkParseOptions("Xml.parse", options);
  const includeComments = options?.comments ?? false;
  const includePI = options?.processingInstructions ?? false;
  const cdataAsNodes = options?.cdataAsNodes ?? false;
  const isFragment = options?.fragment ?? false;

  const parser = new SaxParser({
    position: true,
    fragment: options?.fragment ?? false,
    xmlns: options?.xmlns ?? false,
    maxDepth: options?.maxDepth,
    maxEntityExpansions: options?.maxEntityExpansions,
    invalidCharHandling: options?.invalidCharHandling
  });

  // Stack of elements being built. The bottom is a synthetic root
  // that collects top-level nodes.
  const syntheticRoot = createElement("__root__", {});
  const stack: XmlElement[] = [syntheticRoot];
  let declaration: Record<string, string> | undefined;
  // The result is discarded on any error, so stop at the first one instead of
  // building the rest of the tree: `fail()` rethrows out of `write()`. This is
  // also what makes `maxDepth` / `maxEntityExpansions` bound the work done.
  parser.on("error", err => {
    throw err instanceof XmlParseError ? err : new XmlParseError(err.message);
  });

  parser.on("opentag", tag => {
    const elem = createElement(tag.name, tag.attributes);
    // Propagate namespace info if available
    if (tag.prefix !== undefined) {
      elem.prefix = tag.prefix;
      elem.local = tag.local;
      elem.uri = tag.uri;
    }
    if (tag.ns) {
      elem.ns = { ...tag.ns };
    }
    stack[stack.length - 1].children.push(elem);
    if (!tag.isSelfClosing) {
      stack.push(elem);
    }
  });

  parser.on("closetag", tag => {
    // Self-closing tags were not pushed onto the stack, so don't pop
    if (!tag.isSelfClosing && stack.length > 1) {
      stack.pop();
    }
  });

  parser.on("text", text => {
    if (text.length === 0) {
      return;
    }
    appendText(stack[stack.length - 1], text);
  });

  parser.on("cdata", text => {
    const parent = stack[stack.length - 1];
    if (cdataAsNodes) {
      const node: XmlCData = { type: "cdata", value: text };
      parent.children.push(node);
    } else {
      appendText(parent, text);
    }
  });

  if (includeComments) {
    parser.on("comment", text => {
      const parent = stack[stack.length - 1];
      const node: XmlComment = { type: "comment", value: text };
      parent.children.push(node);
    });
  }

  parser.on("pi", (target, body) => {
    // Always capture the XML declaration into doc.declaration,
    // regardless of whether PI nodes are included in the tree.
    if (target === "xml") {
      const attrs: Record<string, string> = {};
      const re = /(\w+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(body)) !== null) {
        attrs[m[1]] = m[2] ?? m[3];
      }
      declaration = attrs;
    } else if (includePI) {
      const parent = stack[stack.length - 1];
      parent.children.push({
        type: "processing-instruction",
        target,
        body
      });
    }
  });

  parser.write(xml);
  parser.close();

  // Extract the real root from the synthetic container
  const roots = syntheticRoot.children.filter((n): n is XmlElement => n.type === "element");

  if (roots.length === 0) {
    throw new XmlParseError("document has no root element");
  }

  if (!isFragment && roots.length > 1) {
    throw new XmlParseError("document has multiple root elements");
  }

  // Top-level comments and processing instructions, split at the root: those before it are
  // the prologue, those after it the epilogue (in fragment mode, everything after the first
  // root). They used to be collected together as `prologue`.
  const top = syntheticRoot.children;
  const firstRoot = top.indexOf(roots[0]);
  const isMisc = (n: XmlNode): n is XmlComment | XmlProcessingInstruction =>
    n.type === "comment" || n.type === "processing-instruction";
  const prologue = top.slice(0, firstRoot).filter(isMisc);
  const epilogue = top.slice(firstRoot + 1).filter(isMisc);

  return {
    declaration,
    root: roots[0],
    roots,
    prologue,
    epilogue
  };
}

// =============================================================================
// DOM Query Helpers
// =============================================================================

/**
 * The first child element with the given name, or `undefined`. Published as
 * `Xml.findChild`.
 */
function findChild(element: XmlElement, name: string): XmlElement | undefined {
  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      return child;
    }
  }
  return undefined;
}

/**
 * Every child element with the given name, in document order. Published as
 * `Xml.listChildren`.
 */
function findChildren(element: XmlElement, name: string): XmlElement[] {
  const result: XmlElement[] = [];
  for (const child of element.children) {
    if (child.type === "element" && child.name === name) {
      result.push(child);
    }
  }
  return result;
}

/**
 * Get the concatenated text content of an element (recursive).
 */
function textContent(node: XmlNode): string {
  switch (node.type) {
    case "text":
    case "cdata":
      return node.value;
    case "element": {
      let result = "";
      for (const child of node.children) {
        result += textContent(child);
      }
      return result;
    }
    default:
      return "";
  }
}

/**
 * Get an attribute value, or undefined if not present.
 */
function attr(element: XmlElement, name: string): string | undefined {
  return element.attributes[name];
}

/**
 * Walk all descendant elements depth-first, calling visitor for each.
 */
function walk(element: XmlElement, visitor: (el: XmlElement) => void): void {
  visitor(element);
  for (const child of element.children) {
    if (child.type === "element") {
      walk(child, visitor);
    }
  }
}

// =============================================================================
// DOM → Plain Object Conversion
// =============================================================================

/**
 * Convert an {@link XmlElement} DOM tree to a plain JavaScript object.
 *
 * Produces a plain JavaScript object where element names become object keys,
 * attributes are prefixed (default `@_`), text-only elements collapse to their
 * string value, and repeated sibling names merge into arrays.
 *
 * @param element - The element to convert.
 * @param options - Conversion options.
 * @returns A plain JavaScript object representing the element.
 *
 * @example
 * ```ts
 * const doc = parseXml('<root attr="1"><child>text</child></root>');
 * const obj = toPlainObject(doc.root);
 * // { root: { "@_attr": "1", child: "text" } }
 * ```
 *
 * @example
 * ```ts
 * const doc = parseXml('<list><item>a</item><item>b</item></list>');
 * const obj = toPlainObject(doc.root);
 * // { list: { item: ["a", "b"] } }
 * ```
 */
function toPlainObject(
  element: XmlElement,
  options?: XmlToPlainObjectOptions
): Record<string, unknown> {
  checkElement("Xml.toPlainObject", element);
  const opts = resolveOptions(options);

  function convertElement(el: XmlElement): unknown {
    const obj: Record<string, unknown> = Object.create(null);

    // Add attributes — el.attributes is created via Object.create(null)
    // by safeAttributes(), so no prototype keys to guard against.
    let hasAttributes = false;
    if (!opts.ignoreAttributes) {
      for (const key in el.attributes) {
        obj[opts.attrPrefix + key] = el.attributes[key];
        hasAttributes = true;
      }
    }

    // Collect text and child elements in a single pass.
    let text = "";
    let hasChildren = false;

    for (const child of el.children) {
      switch (child.type) {
        case "text":
          text += child.value;
          break;
        case "cdata":
          if (opts.preserveCData) {
            text += child.value;
          }
          break;
        case "element": {
          const value = convertElement(child);
          addChildValue(obj, child.name, value, opts.alwaysArray, opts.isArray);
          hasChildren = true;
          break;
        }
        // comments and PIs are ignored in plain-object conversion
      }
    }

    return resolveValue(obj, text, hasAttributes, hasChildren, opts);
  }

  // No prototype, like every nested object and `Xml.parseToObject`'s result.
  const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  result[element.name] = convertElement(element);
  return result;
}

export { parseXml, findChild, findChildren, textContent, attr, walk, toPlainObject };

// =============================================================================
// Checked public forms (`Xml.findChild`, …)
// =============================================================================

function checkName(where: string, name: unknown): void {
  checkStringArgument(where, "name", name);
}

/** `Xml.findChild`: {@link findChild}, refusing what is not an element and a name. */
function findChildOf(element: XmlElement, name: string): XmlElement | undefined {
  checkElement("Xml.findChild", element);
  checkName("Xml.findChild", name);
  return findChild(element, name);
}

/** `Xml.listChildren`: {@link findChildren}, refusing what is not an element and a name. */
function listChildrenOf(element: XmlElement, name: string): XmlElement[] {
  checkElement("Xml.listChildren", element);
  checkName("Xml.listChildren", name);
  return findChildren(element, name);
}

const xmlNodeSchema = s.custom(
  v => typeof v === "object" && v !== null && typeof (v as { type?: unknown }).type === "string",
  "an XmlNode (pass doc.root for a document)"
);

/** `Xml.getText`: {@link textContent}, refusing what is not a node. */
function getTextOf(node: XmlNode): string {
  check(xmlNodeSchema, node, xmlArgument("Xml.getText", "node"));
  return textContent(node);
}

/** `Xml.getAttr`: {@link attr}, refusing what is not an element and a name. */
function getAttrOf(element: XmlElement, name: string): string | undefined {
  checkElement("Xml.getAttr", element);
  checkName("Xml.getAttr", name);
  return attr(element, name);
}

/** `Xml.eachElement`: {@link walk}, refusing what is not an element and a function. */
function eachElementOf(element: XmlElement, visitor: (el: XmlElement) => void): void {
  checkElement("Xml.eachElement", element);
  checkFunctionArgument("Xml.eachElement", "visitor", visitor);
  walk(element, visitor);
}

export { findChildOf, listChildrenOf, getTextOf, getAttrOf, eachElementOf };
