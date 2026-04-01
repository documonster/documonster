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

import { SaxParser } from "@xml/sax";
import { XmlParseError } from "@xml/errors";
import type {
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

/** Property names that must never be used as object keys from untrusted input. */
const DANGEROUS_KEYS = new Set([
  "__proto__",
  "constructor",
  "prototype",
  "__defineGetter__",
  "__defineSetter__",
  "__lookupGetter__",
  "__lookupSetter__"
]);

/** Create a clean attribute map, filtering dangerous keys. */
function safeAttributes(raw: Record<string, string>): Record<string, string> {
  const attrs: Record<string, string> = Object.create(null);
  for (const key in raw) {
    if (Object.prototype.hasOwnProperty.call(raw, key) && !DANGEROUS_KEYS.has(key)) {
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
 * @example
 * ```ts
 * const doc = parseXml('<root><child id="1">text</child></root>');
 * console.log(doc.root.name); // "root"
 * console.log(doc.root.children[0].name); // "child"
 * ```
 */
function parseXml(xml: string, options?: XmlParseOptions): XmlDocument {
  const includeComments = options?.comments ?? false;
  const includePI = options?.processingInstructions ?? false;
  const cdataAsNodes = options?.cdataAsNodes ?? false;
  const isFragment = options?.fragment ?? false;

  const parser = new SaxParser({
    position: true,
    fragment: options?.fragment ?? false,
    xmlns: options?.xmlns ?? false,
    maxDepth: options?.maxDepth,
    maxEntityExpansions: options?.maxEntityExpansions
  });

  // Stack of elements being built. The bottom is a synthetic root
  // that collects top-level nodes.
  const syntheticRoot = createElement("__root__", {});
  const stack: XmlElement[] = [syntheticRoot];
  let declaration: Record<string, string> | undefined;
  let error: XmlParseError | undefined;

  parser.on("error", err => {
    if (!error) {
      error = err instanceof XmlParseError ? err : new XmlParseError(err.message);
    }
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

  if (error) {
    throw error;
  }

  // Extract the real root from the synthetic container
  const roots = syntheticRoot.children.filter((n): n is XmlElement => n.type === "element");

  if (roots.length === 0) {
    throw new XmlParseError("document has no root element");
  }

  if (!isFragment && roots.length > 1) {
    throw new XmlParseError("document has multiple root elements");
  }

  // Collect top-level comments and processing instructions (prologue)
  const prologue = syntheticRoot.children.filter(
    (n): n is XmlComment | XmlProcessingInstruction =>
      n.type === "comment" || n.type === "processing-instruction"
  );

  return {
    declaration,
    root: roots[0],
    roots,
    prologue
  };
}

// =============================================================================
// DOM Query Helpers
// =============================================================================

/**
 * Find the first child element with the given name.
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
 * Find all child elements with the given name.
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

export { parseXml, findChild, findChildren, textContent, attr, walk };
