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

import {
  checkParseToObjectOptions,
  checkXmlInput,
  snapshotParseToObjectOptions
} from "@xml/arguments";
import { XmlParseError } from "@xml/errors";
import { NO_ATTRIBUTES, SaxParser } from "@xml/sax";
import type { ResolvedOptions } from "@xml/to-object-shared";
import {
  addChildValue,
  addTextValue,
  attributeKey,
  isWhitespaceOnly,
  resolveOptions
} from "@xml/to-object-shared";
import type { XmlParseToObjectOptions, XmlSaxTag } from "@xml/types";

// =============================================================================
// Object Builder
// =============================================================================

/**
 * Each open element pushes a frame onto the stack.
 * - `obj`: the object being built for this element, made when its first attribute or child
 *   element arrives — an element of text alone becomes a string and never needs one
 * - `text`: accumulated text content
 * - `blank`: whether that text is white space alone — kept as it arrives, so the indentation of a
 *   pretty-printed document is never joined into one string just to be found blank and dropped
 * - `name`: element tag name (needed at close time)
 */
interface Frame {
  obj: Record<string, unknown> | undefined;
  text: string;
  blank: boolean;
  hasChildren: boolean;
  name: string;
}

function newObject(): Record<string, unknown> {
  return Object.create(null) as Record<string, unknown>;
}

/**
 * Builds the plain object of one element from its SAX events, as `Xml.parseToObject` does for a
 * document: `Xml.parseToObject` builds the root with one, and `Xml.readRecords` each record, so a
 * record's object is the one `Xml.parseToObject` gives for its text.
 *
 * `result` holds `{ [name]: value }` for each element closed at the outermost level.
 */
export class ObjectBuilder {
  readonly result: Record<string, unknown> = newObject();
  private readonly stack: Frame[];

  /**
   * @param adopt - Whether tags arrive from a parser told to `prefixAttributeNames` with this
   *   prefix, so a tag's `attributes` can become the element's object as they are.
   */
  constructor(
    private readonly opts: ResolvedOptions,
    private readonly adopt = false
  ) {
    this.stack = [{ obj: this.result, text: "", blank: true, hasChildren: false, name: "" }];
  }

  openTag(tag: XmlSaxTag): void {
    const { opts, stack } = this;
    let obj: Record<string, unknown> | undefined;
    if (this.adopt) {
      if (tag.attributes !== NO_ATTRIBUTES) {
        obj = tag.attributes;
      }
    } else if (!opts.ignoreAttributes) {
      const attributes = tag.attributes;
      for (const key in attributes) {
        (obj ??= newObject())[attributeKey(opts, key)] = attributes[key];
      }
    }
    const frame: Frame = { obj, text: "", blank: true, hasChildren: false, name: tag.name };

    const parent = stack[stack.length - 1];
    parent.hasChildren = true;

    if (tag.isSelfClosing) {
      // Resolve immediately — this element has no children or text.
      finishFrame(frame, parent, opts);
    } else {
      stack.push(frame);
    }
  }

  text(text: string): void {
    if (text.length > 0) {
      this.append(text);
    }
  }

  cdata(text: string): void {
    if (this.opts.preserveCData && text.length > 0) {
      this.append(text);
    }
  }

  private append(text: string): void {
    const frame = this.stack[this.stack.length - 1];
    frame.text += text;
    // Most text that is not blank starts with what shows it; indentation starts with a line break.
    if (frame.blank && (text.charCodeAt(0) > 0x20 || !isWhitespaceOnly(text))) {
      frame.blank = false;
    }
  }

  closeTag(tag: XmlSaxTag): void {
    if (tag.isSelfClosing || this.stack.length <= 1) {
      return;
    }
    const frame = this.stack.pop()!;
    finishFrame(frame, this.stack[this.stack.length - 1], this.opts);
  }
}

/**
 * Resolve a completed frame and add it to the parent, as `resolveValue` does for a DOM element:
 * text alone collapses to a string, nothing at all to `""`, anything else is the object with its
 * text under `textKey`. The synthetic root never gets `alwaysArray` applied to it.
 */
function finishFrame(frame: Frame, parent: Frame, opts: ResolvedOptions): void {
  let text = frame.text;
  // Discard whitespace-only text when element has children (formatting indentation).
  if (opts.ignoreWS && frame.hasChildren && frame.blank) {
    text = "";
  }
  const obj = frame.obj;
  let value: unknown = text;
  if (obj !== undefined) {
    if (text.length > 0) {
      addTextValue(obj, text, opts);
    }
    value = obj;
  }
  const isDocRoot = parent.name === "";
  addChildValue(
    (parent.obj ??= newObject()),
    frame.name,
    value,
    isDocRoot ? false : opts.alwaysArray,
    isDocRoot ? null : opts.isArray
  );
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
 * sibling names merge into arrays. Text whose `textKey` is already taken — by a child
 * element, or by an attribute when `attributePrefix` is `""` — is merged with it into an array
 * in the same way, never overwriting it.
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
function parseXmlToObject(xml: string, given?: XmlParseToObjectOptions): Record<string, unknown> {
  checkXmlInput("Xml.parseToObject", xml);
  // Read once, so what is checked is what is used — and the parser need not check it again.
  const options = snapshotParseToObjectOptions(given);
  checkParseToObjectOptions("Xml.parseToObject", options);
  const opts = resolveOptions(options, "Xml.parseToObject");

  const parser = SaxParser.withCheckedOptions({
    position: false,
    fragment: options?.fragment ?? false,
    maxDepth: options?.maxDepth,
    maxEntityExpansions: options?.maxEntityExpansions,
    maxEntityExpansionLength: options?.maxEntityExpansionLength,
    invalidCharHandling: options?.invalidCharHandling,
    internalEntities: options?.internalEntities,
    // The whole input is in memory: its DOCTYPE cannot be larger than it, so the bound that keeps a
    // streamed or serialised prolog small does not apply.
    maxContextBytes: Number.MAX_SAFE_INTEGER
  });
  const builder = new ObjectBuilder(opts);

  // The result is discarded on any error, so stop at the first one instead of
  // building the rest of the tree: `fail()` rethrows out of `write()`. This is
  // also what makes `maxDepth` / `maxEntityExpansions` bound the work done.
  parser.on("error", err => {
    throw err instanceof XmlParseError ? err : new XmlParseError(err.message);
  });
  parser.on("opentag", tag => builder.openTag(tag));
  parser.on("text", text => builder.text(text));
  parser.on("cdata", text => builder.cdata(text));
  parser.on("closetag", tag => builder.closeTag(tag));
  // Comments and PIs are not needed for plain-object output.

  parser.write(xml);
  parser.close();

  // The synthetic root's obj should contain exactly one key (the document root).
  return builder.result;
}

export { parseXmlToObject };
