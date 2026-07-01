/**
 * XML → Plain Object (SAX-direct)
 *
 * Parses an XML string directly into a plain JavaScript object, bypassing the
 * DOM tree entirely. This is the fastest path for "XML string → plain object →
 * JSON.stringify" workflows.
 *
 * For converting an already-parsed {@link XmlElement} tree, use
 * {@link toPlainObject} from `./dom` instead.
 */

import { toError } from "@utils/errors";
import { XmlParseError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import type { ResolvedOptions } from "@xml/to-object-shared";
import { resolveOptions, resolveValue, addChildValue } from "@xml/to-object-shared";
import type { ParseXmlToObjectOptions } from "@xml/types";

// =============================================================================
// Stack Frame
// =============================================================================

/**
 * Each open element pushes a frame onto the stack.
 * - `obj`: the plain object being built for this element
 * - `text`: accumulated text content
 * - `hasChildren`: whether any child elements have been seen
 * - `hasAttributes`: whether any attributes exist
 * - `name`: element tag name (needed at close time)
 */
interface Frame {
  obj: Record<string, unknown>;
  text: string;
  hasChildren: boolean;
  hasAttributes: boolean;
  name: string;
}

// =============================================================================
// parseXmlToObject
// =============================================================================

/**
 * Parse an XML string directly into a plain JavaScript object.
 *
 * Unlike `parseXml` + `toPlainObject`, this function builds the plain object
 * in a single SAX pass with **zero intermediate DOM nodes**. Use this when
 * performance matters and you only need the plain-object output.
 *
 * The output format matches {@link toPlainObject}: attributes are prefixed
 * (default `@_`), text-only elements collapse to strings, and repeated
 * sibling names merge into arrays.
 *
 * @param xml - Complete XML string.
 * @param options - Conversion and parser options.
 * @returns A plain JavaScript object representing the root element.
 * @throws {XmlParseError} If the XML is malformed.
 *
 * @example
 * ```ts
 * const obj = parseXmlToObject('<root attr="1"><child>text</child></root>');
 * // { root: { "@_attr": "1", child: "text" } }
 * ```
 */
function parseXmlToObject(xml: string, options?: ParseXmlToObjectOptions): Record<string, unknown> {
  const opts = resolveOptions(options);

  const parser = new SaxParser({
    position: false,
    fragment: options?.fragment ?? false,
    maxDepth: options?.maxDepth,
    maxEntityExpansions: options?.maxEntityExpansions,
    invalidCharHandling: options?.invalidCharHandling
  });

  // Stack: bottom is a synthetic root frame that collects the document root.
  const syntheticObj: Record<string, unknown> = Object.create(null);
  const stack: Frame[] = [
    { obj: syntheticObj, text: "", hasChildren: false, hasAttributes: false, name: "" }
  ];

  let error: XmlParseError | undefined;

  parser.on("error", err => {
    if (!error) {
      error = err instanceof XmlParseError ? err : new XmlParseError(err.message);
    }
  });

  parser.on("opentag", tag => {
    const frame: Frame = {
      obj: Object.create(null),
      text: "",
      hasChildren: false,
      hasAttributes: false,
      name: tag.name
    };

    // Write attributes directly into frame.obj
    if (!opts.ignoreAttributes) {
      for (const key in tag.attributes) {
        frame.obj[opts.attrPrefix + key] = tag.attributes[key];
        frame.hasAttributes = true;
      }
    }

    // Mark parent as having children
    stack[stack.length - 1].hasChildren = true;

    if (tag.isSelfClosing) {
      // Resolve immediately — this element has no children or text.
      finishFrame(frame, stack[stack.length - 1], opts);
    } else {
      stack.push(frame);
    }
  });

  parser.on("text", text => {
    if (text.length > 0) {
      stack[stack.length - 1].text += text;
    }
  });

  parser.on("cdata", text => {
    if (opts.preserveCData && text.length > 0) {
      stack[stack.length - 1].text += text;
    }
  });

  parser.on("closetag", tag => {
    if (tag.isSelfClosing) {
      return;
    }
    if (stack.length <= 1) {
      return;
    }
    const frame = stack.pop()!;
    finishFrame(frame, stack[stack.length - 1], opts);
  });

  // Ignore comments and PIs — not needed for plain-object output

  parser.write(xml);
  parser.close();

  if (error) {
    throw toError(error);
  }

  // The synthetic root's obj should contain exactly one key (the document root).
  return syntheticObj;
}

/**
 * Resolve a completed frame and add it to the parent.
 * The synthetic root (stack depth 1) never gets `alwaysArray` applied to it.
 */
function finishFrame(frame: Frame, parent: Frame, opts: ResolvedOptions): void {
  const value = resolveValue(frame.obj, frame.text, frame.hasAttributes, frame.hasChildren, opts);
  const isDocRoot = parent.name === "";
  addChildValue(
    parent.obj,
    frame.name,
    value,
    isDocRoot ? false : opts.alwaysArray,
    isDocRoot ? null : opts.isArray
  );
}

export { parseXmlToObject };
