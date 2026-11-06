/**
 * XmlStreamWriter - True Streaming XML Writer
 *
 * Writes XML directly to a {@link XmlWritableTarget} without buffering
 * the entire document in memory. This is the correct solution for
 * large XML output (e.g. worksheets with hundreds of thousands of rows).
 *
 * Implements the same {@link XmlSink} interface as {@link XmlWriter},
 * so rendering code can target either backend transparently.
 */

import {
  xmlEncode,
  xmlEncodeAttr,
  validateXmlName,
  xmlValue,
  cdataMarkup,
  rawMarkup,
  commentMarkup,
  serializeXmlAttributes,
  checkNewAttributes,
  StdDocAttributes
} from "@xml/encode";
import { XmlArgumentError, XmlWriteError } from "@xml/errors";
import type { XmlWritableTarget, XmlAttributes, XmlSink } from "@xml/types";

// =============================================================================
// XmlStreamWriter
// =============================================================================

/**
 * Streaming XML writer that flushes content directly to a writable target.
 *
 * Unlike {@link XmlWriter}, this class never holds the full XML in memory.
 * Each method call immediately writes its output to the target.
 *
 * **Trade-off**: No rollback support. Once written, content cannot be undone.
 * Use {@link XmlWriter} if you need speculative/transactional writes.
 *
 * @example
 * ```ts
 * const chunks: string[] = [];
 * const target = { write(chunk: string) { chunks.push(chunk); } };
 * const sw = new Xml.StreamWriter(target);
 * sw.openXml();
 * sw.openNode("root");
 * sw.leafNode("item", { id: "1" }, "hello");
 * sw.closeNode();
 * console.log(chunks.join(""));
 * ```
 */
class XmlStreamWriter implements XmlSink {
  private _target: XmlWritableTarget;
  private _stack: string[] = [];
  private _leaf = false;
  private _open = false;
  // Pending start-tag buffer: accumulates tag name + attributes until flushed
  private _pending = "";
  // Whether anything has been written or is pending — the XML declaration must come first
  private _started = false;

  constructor(target: XmlWritableTarget) {
    if (
      typeof target !== "object" ||
      target === null ||
      typeof (target as { write?: unknown }).write !== "function"
    ) {
      throw new XmlArgumentError(
        "target",
        "Xml.StreamWriter: target must be an object with a write(chunk) method"
      );
    }
    this._target = target;
  }

  // ===========================================================================
  // State Queries
  // ===========================================================================

  /** Current nesting depth. */
  get depth(): number {
    return this._stack.length;
  }

  /** Name of the innermost open element, or undefined. */
  get currentElement(): string | undefined {
    return this._stack.length > 0 ? this._stack[this._stack.length - 1] : undefined;
  }

  // ===========================================================================
  // Internal: flush pending start tag
  // ===========================================================================

  private _write(chunk: string): void {
    this._started = true;
    this._target.write(chunk);
  }

  /** Flush the pending start-tag buffer with a closing character. */
  private _flushOpen(suffix: string): void {
    this._write(this._pending + suffix);
    this._pending = "";
    this._open = false;
  }

  // ===========================================================================
  // XmlSink Implementation
  // ===========================================================================

  /**
   * Write the XML declaration. XML allows it only at the very start of a document, so it is
   * refused once anything has been written, including an earlier declaration.
   */
  openXml(attributes?: XmlAttributes): void {
    if (this._started || this._open) {
      throw new XmlWriteError("open xml", "the XML declaration must come first in the document");
    }
    const merged = attributes ? { ...StdDocAttributes, ...attributes } : StdDocAttributes;
    this._write(`<?xml${serializeXmlAttributes(merged)}?>\n`);
  }

  openNode(name: string, attributes?: XmlAttributes): void {
    validateXmlName(name);
    // Build the tag before flushing, so a refused attribute writes nothing.
    const s = `<${name}${serializeXmlAttributes(attributes)}`;
    if (this._open) {
      this._flushOpen(">");
    }
    this._pending = s;
    this._stack.push(name);
    this._leaf = true;
    this._open = true;
  }

  addAttribute(name: string, value: string | number | boolean): void {
    if (!this._open) {
      throw new XmlWriteError("add attribute", "no element is open");
    }
    validateXmlName(name);
    checkNewAttributes(this.currentElement, this._pending, { [name]: value }, () => "name");
    // Append to pending buffer — no write call
    this._pending += ` ${name}="${xmlEncodeAttr(xmlValue(value, name, "value"))}"`;
  }

  addAttributes(attributes: XmlAttributes): void {
    if (!this._open) {
      throw new XmlWriteError("add attributes", "no element is open");
    }
    checkNewAttributes(this.currentElement, this._pending, attributes);
    this._pending += serializeXmlAttributes(attributes);
  }

  writeText(text: string | number): void {
    const encoded = xmlEncode(xmlValue(text, "#text"));
    if (this._open) {
      this._flushOpen(">");
    }
    this._leaf = false;
    this._write(encoded);
  }

  writeRaw(xml: string): void {
    const markup = rawMarkup(xml);
    if (this._open) {
      this._flushOpen(">");
    }
    this._leaf = false;
    this._write(markup);
  }

  writeCData(text: string): void {
    const markup = cdataMarkup(text);
    if (this._open) {
      this._flushOpen(">");
    }
    this._leaf = false;
    this._write(markup);
  }

  writeComment(text: string): void {
    const markup = commentMarkup(text);
    if (this._open) {
      this._flushOpen(">");
    }
    this._leaf = false;
    this._write(markup);
  }

  closeNode(): void {
    const name = this._stack.pop();
    if (name === undefined) {
      throw new XmlWriteError("close node", "no element is open");
    }
    if (this._leaf && this._open) {
      // Self-closing: flush pending tag with "/>", single write
      this._flushOpen("/>");
    } else if (this._open) {
      this._flushOpen(">");
      this._write(`</${name}>`);
    } else {
      this._write(`</${name}>`);
    }
    this._open = false;
    this._leaf = false;
  }

  leafNode(name: string, attributes?: XmlAttributes, text?: string | number): void {
    validateXmlName(name);
    // Build the complete leaf element as one string, before flushing anything.
    let s = `<${name}${serializeXmlAttributes(attributes)}`;
    if (text !== undefined) {
      s += `>${xmlEncode(xmlValue(text, "#text"))}</${name}>`;
    } else {
      s += "/>";
    }
    if (this._open) {
      this._flushOpen(">");
    }
    this._write(s);
    this._leaf = false;
  }

  // ===========================================================================
  // Close All
  // ===========================================================================

  /** Close all open elements. */
  closeAll(): void {
    while (this._stack.length > 0) {
      this.closeNode();
    }
  }
}

export { XmlStreamWriter };
