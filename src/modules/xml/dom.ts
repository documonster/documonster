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
  checkAttributeValue,
  checkElement,
  checkFunctionArgument,
  checkParseOptions,
  checkStringArgument,
  checkXmlInput,
  readDomAttributes,
  readDomChildren,
  readDomName,
  readDomText,
  readDomType,
  refuseDom,
  snapshotParseOptions,
  xmlArgument
} from "@xml/arguments";
import type { DomPath } from "@xml/arguments";
import { XmlArgumentError, XmlParseError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import { addChildValue, attributeKey, resolveOptions, resolveValue } from "@xml/to-object-shared";
import type { ResolvedOptions } from "@xml/to-object-shared";
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
 * A `<?xml …?>` that is not at the very start — after white space, a second one, one inside an
 * element — is refused, as XML has it. (This package's document readers parse a package part from
 * its first `<`, which is their policy for parts, not the parser's.)
 *
 * @example
 * ```ts
 * const doc = parseXml('<root><child id="1">text</child></root>');
 * console.log(doc.root.name); // "root"
 * console.log(doc.root.children[0].name); // "child"
 * ```
 */
function parseXml(xml: string, given?: XmlParseOptions): XmlDocument {
  checkXmlInput("Xml.parse", xml);
  const builder = createDomBuilder(given);
  builder.write(xml);
  return builder.close();
}

/**
 * {@link parseXml} fed a piece at a time: `write` each piece of the text in order, then `close` for the document.
 *
 * The same tree, built as the text arrives — for a reader that has to give the event loop a turn while it parses a
 * large part, which one `parseXml` call cannot. A piece may end anywhere, inside a tag or a name included.
 */
export function createDomBuilder(given?: XmlParseOptions): {
  write(text: string): void;
  close(): XmlDocument;
} {
  // Read once, so what is checked is what is used — and the parser need not check it again.
  const options = snapshotParseOptions(given);
  checkParseOptions("Xml.parse", options);
  const includeComments = options?.comments ?? false;
  const includePI = options?.processingInstructions ?? false;
  const cdataAsNodes = options?.cdataAsNodes ?? false;
  const isFragment = options?.fragment ?? false;

  const parser = SaxParser.withCheckedOptions({
    position: true,
    fragment: options?.fragment ?? false,
    xmlns: options?.xmlns ?? false,
    maxDepth: options?.maxDepth,
    maxEntityExpansions: options?.maxEntityExpansions,
    maxEntityExpansionLength: options?.maxEntityExpansionLength,
    invalidCharHandling: options?.invalidCharHandling,
    internalEntities: options?.internalEntities,
    // The whole input is in memory — written in pieces at most, by a caller holding it: its DOCTYPE cannot be larger
    // than it, so the bound that keeps a streamed or serialised prolog small does not apply.
    maxContextBytes: Number.MAX_SAFE_INTEGER
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
    } else if (includePI && !parser.inInternalSubset) {
      // One in the DOCTYPE internal subset belongs to the DTD, not to the document's prologue.
      const parent = stack[stack.length - 1];
      parent.children.push({
        type: "processing-instruction",
        target,
        body
      });
    }
  });

  return {
    write(text: string): void {
      parser.write(text);
    },
    close(): XmlDocument {
      parser.close();
      return finishDocument(syntheticRoot, declaration, isFragment);
    }
  };
}

/** The document a {@link createDomBuilder} built, once its text has all been written. */
function finishDocument(
  syntheticRoot: XmlElement,
  declaration: Record<string, string> | undefined,
  isFragment: boolean
): XmlDocument {
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
 * The depth past which a walk of a tree its caller gave records the elements it is inside, to find one
 * among its own descendants: a cycle descends without end, so it always gets there, and a tree
 * shallower than this — nearly every one — pays nothing for the check.
 */
export const CYCLE_CHECK_DEPTH = 1024;

/**
 * The elements open around the one a walk is at, for a walk that takes a tree from its caller: an
 * element among its own descendants makes the tree a cycle, which a walk would never leave. An
 * element reached twice by different paths — shared, not its own ancestor — is no cycle. The elements
 * are looked up only past {@link CYCLE_CHECK_DEPTH}; above it they are only kept.
 */
export class AncestorPath {
  /** The elements entered, in order: the root first, which is never left. */
  private readonly stack: unknown[];
  /** The same elements as a set, once the walk is deep enough to be looking for a cycle. */
  private open: Set<unknown> | undefined;

  constructor(
    private readonly where: string,
    private readonly field: string,
    root: unknown
  ) {
    this.stack = [root];
  }

  enter(element: unknown): void {
    let open = this.open;
    if (open === undefined && this.stack.length > CYCLE_CHECK_DEPTH) {
      open = this.open = new Set(this.stack);
    }
    if (open !== undefined) {
      if (open.has(element)) {
        throw new XmlArgumentError(
          this.field,
          `${this.where}: ${this.field} contains itself: an element is among its own descendants`
        );
      }
      open.add(element);
    }
    this.stack.push(element);
  }

  leave(): void {
    if (this.stack.length > 1) {
      const element = this.stack.pop();
      this.open?.delete(element);
    }
  }
}

/**
 * Get the concatenated text content of an element (recursive). Walks with an explicit stack, so
 * a tree of any depth is read. Each node's fields are read once and checked as read.
 */
function textContent(node: XmlNode, where = "Xml.getText", guard = false): string {
  return textOf(node, where, guard, "");
}

function textOf(node: unknown, where: string, guard: boolean, at: string): string {
  const type = readDomType(where, "node", node, at);
  switch (type) {
    case "text":
    case "cdata":
      return readDomText(where, "node", node as object, at);
    case "element":
      readDomName(where, "node", node as object, at);
      break;
    default:
      return "";
  }
  // Recursive, then by `textDeep` past `RECURSION_DEPTH`; what is the same for the whole walk is
  // closed over rather than passed at every element.
  const below = (children: readonly unknown[], depth: number): string => {
    let result = "";
    const length = children.length;
    for (let i = 0; i < length; i++) {
      const child = children[i] as
        | { type?: unknown; value?: unknown; children?: unknown }
        | null
        | undefined;
      // The checks of `readDomType` and its siblings, inline: they run for every node of the tree.
      // `null` and a primitive read as no type; told apart only to say why they are refused.
      const childType = child?.type;
      if (childType === "text" || childType === "cdata") {
        const value = child!.value;
        if (typeof value !== "string") {
          refuseDom(where, "node", i, ".value", "a string", value);
        }
        result += value;
      } else if (childType === "element") {
        const grandchildren = child!.children;
        if (!Array.isArray(grandchildren)) {
          refuseDom(where, "node", i, ".children", "an array", grandchildren);
        }
        if (grandchildren.length > 0) {
          result +=
            depth < RECURSION_DEPTH
              ? below(grandchildren, depth + 1)
              : textDeep(child, grandchildren, where, guard);
        }
      } else if (typeof childType !== "string") {
        if (child === null || typeof child !== "object") {
          refuseDom(where, "node", i, "", "an XmlNode", child);
        }
        refuseDom(where, "node", i, ".type", "a node type", childType);
      }
    }
    return result;
  };
  return below(readDomChildren(where, "node", node as object, at), 0);
}

/**
 * How deep a walk of a tree its caller gave recurses before it hands a subtree to a walk with an
 * explicit stack: recursion keeps each element's state on the call stack, which costs nothing to
 * allocate, but the call stack is bounded and a tree is not. A cycle always descends past it, into
 * the explicit walk, which finds it past {@link CYCLE_CHECK_DEPTH}.
 */
export const RECURSION_DEPTH = 512;

/** `textOf`'s recursion with an explicit stack, so a subtree of any depth is read. */
function textDeep(
  element: unknown,
  children: readonly unknown[],
  where: string,
  guard: boolean
): string {
  let result = "";
  // Each frame is an element's children and the next index to visit in it.
  const lists: Array<readonly unknown[]> = [children];
  const indices: number[] = [0];
  const path = guard ? new AncestorPath(where, "node", element) : undefined;
  while (lists.length > 0) {
    const top = lists.length - 1;
    const list = lists[top];
    const i = indices[top];
    if (i >= list.length) {
      lists.pop();
      indices.pop();
      path?.leave();
      continue;
    }
    indices[top] = i + 1;
    const child = list[i];
    if (typeof child !== "object" || child === null) {
      refuseDom(where, "node", i, "", "an XmlNode", child);
    }
    const childType = (child as { type?: unknown }).type;
    if (childType === "text" || childType === "cdata") {
      const value = (child as { value?: unknown }).value;
      if (typeof value !== "string") {
        refuseDom(where, "node", i, ".value", "a string", value);
      }
      result += value;
    } else if (childType === "element") {
      const grandchildren = (child as { children?: unknown }).children;
      if (!Array.isArray(grandchildren)) {
        refuseDom(where, "node", i, ".children", "an array", grandchildren);
      }
      if (grandchildren.length > 0) {
        path?.enter(child);
        lists.push(grandchildren);
        indices.push(0);
      }
    } else if (typeof childType !== "string") {
      refuseDom(where, "node", i, ".type", "a node type", childType);
    }
  }
  return result;
}

/**
 * Get an attribute value, or undefined if not present.
 */
function attr(element: XmlElement, name: string): string | undefined {
  // Its own attributes only: `constructor` is no attribute an element has because objects have it.
  const attributes = element.attributes;
  return Object.hasOwn(attributes, name) ? attributes[name] : undefined;
}

/**
 * Walk all descendant elements depth-first (pre-order), calling visitor for each. Walks with an
 * explicit stack, so a tree of any depth is visited.
 */
function walk(
  element: XmlElement,
  visitor: (el: XmlElement) => void,
  where = "Xml.eachElement",
  guard = false
): void {
  const rootChildren = readDomChildren(where, "element", element, "");
  visitor(element);
  // Recursive, then by `walkDeep` past `RECURSION_DEPTH`; what is the same for the whole walk is
  // closed over rather than passed at every element.
  const below = (children: readonly unknown[], depth: number): void => {
    const length = children.length;
    for (let i = 0; i < length; i++) {
      const child = children[i] as { type?: unknown; children?: unknown } | null | undefined;
      // `null` and a primitive read as no element; told apart from one only to say why it is refused.
      if (child?.type === "element") {
        // `readDomChildren`'s check, inline: it runs for every element of the tree.
        const grandchildren = child.children;
        if (!Array.isArray(grandchildren)) {
          refuseDom(where, "element", i, ".children", "an array", grandchildren);
        }
        visitor(child as XmlElement);
        if (grandchildren.length > 0) {
          if (depth < RECURSION_DEPTH) {
            below(grandchildren, depth + 1);
          } else {
            walkDeep(child, grandchildren, visitor, where, guard);
          }
        }
      } else if (child === null || typeof child !== "object") {
        refuseDom(where, "element", i, "", "an XmlNode", child);
      }
    }
  };
  below(rootChildren, 0);
}

/** `walk`'s recursion with an explicit stack, so a subtree of any depth is walked. */
function walkDeep(
  element: unknown,
  children: readonly unknown[],
  visitor: (el: XmlElement) => void,
  where: string,
  guard: boolean
): void {
  const lists: Array<readonly unknown[]> = [children];
  const indices: number[] = [0];
  const path = guard ? new AncestorPath(where, "element", element) : undefined;
  while (lists.length > 0) {
    const top = lists.length - 1;
    const list = lists[top];
    const i = indices[top];
    if (i >= list.length) {
      lists.pop();
      indices.pop();
      path?.leave();
      continue;
    }
    indices[top] = i + 1;
    const child = list[i];
    if (child === null || typeof child !== "object") {
      refuseDom(where, "element", i, "", "an XmlNode", child);
    }
    if ((child as { type?: unknown }).type === "element") {
      const grandchildren = (child as { children?: unknown }).children;
      if (!Array.isArray(grandchildren)) {
        refuseDom(where, "element", i, ".children", "an array", grandchildren);
      }
      if (grandchildren.length === 0) {
        // A leaf: nothing below it to walk, so nothing to enter.
        visitor(child as XmlElement);
        continue;
      }
      path?.enter(child);
      visitor(child as XmlElement);
      lists.push(grandchildren);
      indices.push(0);
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
 * string value, and repeated sibling names merge into arrays. Text whose `textKey` is already
 * taken — by a child element, or by an attribute when `attributePrefix` is `""` — is merged
 * with it into an array in the same way, never overwriting it.
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
/**
 * {@link convertElement} with an explicit stack rather than recursion, so a subtree of any depth
 * converts. Each frame is finished — its value resolved and added to its parent — once its last child
 * is done, which is the order the recursive form adds values in. Leaves the element's name in
 * `openedName`.
 */
function convertDeep(element: object, at: DomPath, opts: ResolvedOptions): unknown {
  let rootValue: unknown;
  let rootName = "";
  const stack: PlainFrame[] = [openFrame(element, at, opts)];
  const path = new AncestorPath("Xml.toPlainObject", "element", element);
  while (stack.length > 0) {
    const frame = stack[stack.length - 1];
    const children = frame.children;
    if (frame.index >= children.length) {
      stack.pop();
      path.leave();
      const value = resolveValue(
        frame.obj,
        frame.text,
        frame.hasAttributes,
        frame.hasChildren,
        opts
      );
      if (stack.length === 0) {
        rootValue = value;
        rootName = frame.name;
      } else {
        addChildValue(
          stack[stack.length - 1].obj,
          frame.name,
          value,
          opts.alwaysArray,
          opts.isArray
        );
      }
      continue;
    }
    const i = frame.index++;
    const child = children[i];
    // The checks of `readDomType` and `readDomText`, inline: they run for every node of the tree.
    if (typeof child !== "object" || child === null) {
      refuseDom("Xml.toPlainObject", "element", i, "", "an XmlNode", child);
    }
    const type = (child as { type?: unknown }).type;
    if (type === "element") {
      const obj = openElement(child, i, opts);
      const name = openedName;
      const grandchildren = openedChildren;
      const hasAttributes = openedHasAttributes;
      frame.hasChildren = true;
      if (grandchildren.length === 0) {
        // A leaf is finished as it is opened: no frame, and it never enters the stack.
        const value = resolveValue(obj, "", hasAttributes, false, opts);
        addChildValue(frame.obj, name, value, opts.alwaysArray, opts.isArray);
      } else {
        path.enter(child);
        stack.push({
          name,
          children: grandchildren,
          obj,
          hasAttributes,
          index: 0,
          text: "",
          hasChildren: false
        });
      }
    } else if (type === "text" || type === "cdata") {
      const value = (child as { value?: unknown }).value;
      if (typeof value !== "string") {
        refuseDom("Xml.toPlainObject", "element", i, ".value", "a string", value);
      }
      if (type === "text" || opts.preserveCData) {
        frame.text += value;
      }
    } else if (typeof type !== "string") {
      refuseDom("Xml.toPlainObject", "element", i, ".type", "a node type", type);
    }
    // comments and PIs are ignored in plain-object conversion
  }
  openedName = rootName;
  return rootValue;
}

/**
 * The value `el` converts to, and its name in `openedName`. Below {@link RECURSION_DEPTH} the work
 * goes to {@link convertDeep}; a cycle is found there, as it always descends that far.
 */
function convertElement(el: object, at: DomPath, opts: ResolvedOptions, depth: number): unknown {
  if (depth >= RECURSION_DEPTH) {
    return convertDeep(el, at, opts);
  }
  // `openElement`, inline: it runs for every element of the tree.
  const where = "Xml.toPlainObject";
  const name = (el as { name?: unknown }).name;
  if (typeof name !== "string") {
    refuseDom(where, "element", at, ".name", "a string", name);
  }
  const children = (el as { children?: unknown }).children;
  if (!Array.isArray(children)) {
    refuseDom(where, "element", at, ".children", "an array", children);
  }
  const obj: Record<string, unknown> = Object.create(null);
  let hasAttributes = false;
  if (!opts.ignoreAttributes) {
    const attributes = (el as { attributes?: unknown }).attributes;
    if (typeof attributes !== "object" || attributes === null || Array.isArray(attributes)) {
      refuseDom(where, "element", "", ".attributes", "an object", attributes);
    }
    for (const key in attributes) {
      if (Object.prototype.hasOwnProperty.call(attributes, key)) {
        const value: unknown = (attributes as Record<string, unknown>)[key];
        if (typeof value !== "string") {
          checkAttributeValue(where, "element", key, value);
        }
        obj[attributeKey(opts, key)] = value;
        hasAttributes = true;
      }
    }
  }
  let text = "";
  let hasChildren = false;
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    // The checks of `readDomType` and `readDomText`, inline: they run for every node of the tree.
    if (typeof child !== "object" || child === null) {
      refuseDom("Xml.toPlainObject", "element", i, "", "an XmlNode", child);
    }
    const type = (child as { type?: unknown }).type;
    if (type === "element") {
      const value = convertElement(child, i, opts, depth + 1);
      addChildValue(obj, openedName, value, opts.alwaysArray, opts.isArray);
      hasChildren = true;
    } else if (type === "text" || type === "cdata") {
      const value = (child as { value?: unknown }).value;
      if (typeof value !== "string") {
        refuseDom("Xml.toPlainObject", "element", i, ".value", "a string", value);
      }
      if (type === "text" || opts.preserveCData) {
        text += value;
      }
    } else if (typeof type !== "string") {
      refuseDom("Xml.toPlainObject", "element", i, ".type", "a node type", type);
    }
    // comments and PIs are ignored in plain-object conversion
  }
  openedName = name;
  return resolveValue(obj, text, hasAttributes, hasChildren, opts);
}

function toPlainObject(
  element: XmlElement,
  options?: XmlToPlainObjectOptions
): Record<string, unknown> {
  checkElement("Xml.toPlainObject", element);
  const opts = resolveOptions(options);
  // No prototype, like every nested object and `Xml.parseToObject`'s result.
  const result: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
  const value = convertElement(element, "", opts, 0);
  result[openedName] = value;
  return result;
}

/** An element being converted by {@link toPlainObject}. */
interface PlainFrame {
  readonly name: string;
  readonly children: readonly unknown[];
  readonly obj: Record<string, unknown>;
  readonly hasAttributes: boolean;
  index: number;
  text: string;
  hasChildren: boolean;
}

/** What {@link openElement} read besides the object it returns: set by it, read straight after. */
let openedName = "";
let openedChildren: readonly unknown[] = [];
let openedHasAttributes = false;

/**
 * Read `el`'s name, children and attributes, each once and checked, and return the object its
 * attributes start; the rest is left in `opened*`, so a leaf is converted with no frame allocated.
 */
function openElement(el: object, at: DomPath, opts: ResolvedOptions): Record<string, unknown> {
  const where = "Xml.toPlainObject";
  openedName = readDomName(where, "element", el, at);
  openedChildren = readDomChildren(where, "element", el, at);
  const obj: Record<string, unknown> = Object.create(null);
  let hasAttributes = false;
  if (!opts.ignoreAttributes) {
    const attributes = readDomAttributes(where, "element", el);
    for (const key in attributes) {
      if (Object.prototype.hasOwnProperty.call(attributes, key)) {
        const value: unknown = attributes[key];
        if (typeof value !== "string") {
          checkAttributeValue(where, "element", key, value);
        }
        obj[attributeKey(opts, key)] = value;
        hasAttributes = true;
      }
    }
  }
  openedHasAttributes = hasAttributes;
  return obj;
}

function openFrame(el: object, at: DomPath, opts: ResolvedOptions): PlainFrame {
  const obj = openElement(el, at, opts);
  return {
    name: openedName,
    children: openedChildren,
    obj,
    hasAttributes: openedHasAttributes,
    index: 0,
    text: "",
    hasChildren: false
  };
}

export { parseXml, findChild, findChildren, textContent, attr, walk, toPlainObject };

// =============================================================================
// Checked public forms (`Xml.findChild`, …)
// =============================================================================

function checkName(where: string, name: unknown): void {
  checkStringArgument(where, "name", name);
}

/** The first (or every, with `all`) child element of `element` named `name`, each child checked as read. */
function scanChildren(
  where: string,
  element: XmlElement,
  name: string,
  all: boolean
): XmlElement[] {
  const result: XmlElement[] = [];
  const children = readDomChildren(where, "element", element, "");
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    const at = i;
    const type = readDomType(where, "element", child, at);
    if (type === "text" || type === "cdata") {
      readDomText(where, "element", child as object, at);
    } else if (type === "element") {
      const childName = readDomName(where, "element", child as object, at);
      readDomChildren(where, "element", child as object, at);
      if (childName === name) {
        result.push(child as XmlElement);
        if (!all) {
          break;
        }
      }
    }
  }
  return result;
}

/** `Xml.findChild`: {@link findChild}, refusing what is not an element and a name. */
function findChildOf(element: XmlElement, name: string): XmlElement | undefined {
  checkElement("Xml.findChild", element);
  checkName("Xml.findChild", name);
  return scanChildren("Xml.findChild", element, name, false)[0];
}

/** `Xml.listChildren`: {@link findChildren}, refusing what is not an element and a name. */
function listChildrenOf(element: XmlElement, name: string): XmlElement[] {
  checkElement("Xml.listChildren", element);
  checkName("Xml.listChildren", name);
  return scanChildren("Xml.listChildren", element, name, true);
}

const xmlNodeSchema = s.custom(
  v => typeof v === "object" && v !== null && typeof (v as { type?: unknown }).type === "string",
  "an XmlNode (pass doc.root for a document)"
);

/** `Xml.getText`: {@link textContent}, refusing what is not a node. */
function getTextOf(node: XmlNode): string {
  check(xmlNodeSchema, node, xmlArgument("Xml.getText", "node"));
  return textOf(node, "Xml.getText", true, "");
}

/** `Xml.getAttr`: {@link attr}, refusing what is not an element and a name. */
function getAttrOf(element: XmlElement, name: string): string | undefined {
  checkElement("Xml.getAttr", element);
  const attributes = readDomAttributes("Xml.getAttr", "element", element);
  checkName("Xml.getAttr", name);
  // Its own attributes only: `constructor` is no attribute an element has because objects have it.
  if (!Object.hasOwn(attributes, name)) {
    return undefined;
  }
  const value = attributes[name];
  checkAttributeValue("Xml.getAttr", "element", name, value);
  return value;
}

/** `Xml.eachElement`: {@link walk}, refusing what is not an element and a function. */
function eachElementOf(element: XmlElement, visitor: (el: XmlElement) => void): void {
  checkElement("Xml.eachElement", element);
  checkFunctionArgument("Xml.eachElement", "visitor", visitor);
  walk(element, visitor, "Xml.eachElement", true);
}

export { findChildOf, listChildrenOf, getTextOf, getAttrOf, eachElementOf };
