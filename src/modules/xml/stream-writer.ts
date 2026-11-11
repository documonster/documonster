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
  validateXmlName,
  xmlValue,
  cdataMarkup,
  rawMarkup,
  commentMarkup,
  XmlTagStack,
  BUSY_READING,
  xmlDeclarationMarkup
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
/** {@link XmlStreamWriter.busy}: in the target's `write()`. */
const BUSY_WRITING = 2;
/** {@link XmlStreamWriter.busy}: the target's `write()` threw — what it was given is unknown. */
const BUSY_FAILED = 4;

class XmlStreamWriter implements XmlSink {
  /** Written to as `target.write(chunk)`, a method call, for each chunk. */
  private readonly _target: XmlWritableTarget;
  /**
   * What the writer is in the middle of — `BUSY_READING` the caller's attributes, {@link BUSY_WRITING}
   * to the target — or {@link BUSY_FAILED} once the target threw: `0` when a call may write, so the
   * check every call makes is one comparison. Written by the tag stack too.
   * @internal
   */
  busy = 0;
  /** The open elements, what the open start tag holds and the namespaces declared. */
  private readonly _tags = new XmlTagStack(this);
  private _leaf = false;
  private _open = false;
  // Pending start-tag buffer: accumulates tag name + attributes until flushed
  private _pending = "";
  // Whether anything has been written or is pending — the XML declaration must come first
  private _started = false;
  /** Set once the target's `write` threw: what it was given is unknown, so nothing more is written. */
  private _failure: { error: unknown } | undefined;

  constructor(target: XmlWritableTarget) {
    const write: unknown =
      typeof target === "object" && target !== null
        ? (target as { write?: unknown }).write
        : undefined;
    if (typeof write !== "function") {
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
    return this._tags.depth;
  }

  /** Name of the innermost open element, or undefined. */
  get currentElement(): string | undefined {
    return this._tags.current;
  }

  // ===========================================================================
  // Internal: flush pending start tag
  // ===========================================================================

  private _write(chunk: string): void {
    this._started = true;
    // Every call that writes has found the writer not busy, so it sets and clears the flag outright.
    this.busy = BUSY_WRITING;
    try {
      this._target.write(chunk);
    } catch (error) {
      this.busy = BUSY_FAILED;
      this._failure = { error };
      throw error;
    }
    this.busy = 0;
  }

  /**
   * Refuse `operation`: the writer is busy — in its target's `write()`, reading the caller's
   * attributes — or its target has failed, so its state no longer matches its output. Every call
   * checks `busy` itself before it does anything, so the check costs one comparison.
   */
  private _refuse(operation: string): void {
    if ((this.busy & BUSY_WRITING) !== 0) {
      // Written to from inside its target's `write()`, the writer would write what it has already
      // written: its state changes only once that call returns.
      throw new XmlWriteError(operation, "the target's write() is running; write after it returns");
    }
    if ((this.busy & BUSY_READING) !== 0) {
      // A getter that wrote would put its output inside the start tag being built.
      throw new XmlWriteError(
        operation,
        "the writer is reading an attribute value; write after that call returns"
      );
    }
    if (this._failure !== undefined) {
      throw new XmlWriteError(operation, "the writer failed when its target's write() threw", {
        cause: this._failure.error
      });
    }
  }

  /** Flush the pending start-tag buffer with a closing character. */
  private _flushOpen(suffix: string): void {
    this._write(this._pending + suffix);
    this._pending = "";
    this._open = false;
  }

  /**
   * Write `chunk`, content of the innermost element: after the start tags still pending, closed, in one
   * write — a target called once for a row, not once per tag.
   */
  private _writeContent(chunk: string): void {
    if (this._open) {
      this._flushOpen(">" + chunk);
    } else {
      this._write(chunk);
    }
  }

  // ===========================================================================
  // XmlSink Implementation
  // ===========================================================================

  /**
   * Write the XML declaration. XML allows it only at the very start of a document, so it is
   * refused once anything has been written, including an earlier declaration.
   */
  openXml(attributes?: XmlAttributes): void {
    if (this.busy !== 0) {
      this._refuse("open xml");
    }
    if (this._started || this._open) {
      throw new XmlWriteError("open xml", "the XML declaration must come first in the document");
    }
    this.busy |= BUSY_READING;
    let markup: string;
    try {
      markup = xmlDeclarationMarkup("Xml.StreamWriter#openXml", attributes);
    } finally {
      this.busy &= ~BUSY_READING;
    }
    this._write(markup);
  }

  openNode(name: string, attributes?: XmlAttributes): void {
    if (this.busy !== 0) {
      this._refuse("open node");
    }
    validateXmlName(name);
    // The start tag is held until what follows it is written, the tags it is inside with it: an
    // element opened inside a pending tag closes that tag in the buffer, not with a write.
    let tag: string;
    if (attributes === undefined) {
      this._tags.openPlain(name);
      tag = `<${name}`;
    } else {
      const markup = this._tags.startTag(name, attributes);
      this._tags.open(name, markup);
      tag = `<${name}${markup}`;
    }
    this._pending = this._open ? `${this._pending}>${tag}` : tag;
    this._leaf = true;
    this._open = true;
  }

  addAttribute(name: string, value: string | number | boolean): void {
    if (this.busy !== 0) {
      this._refuse("add attribute");
    }
    if (!this._open) {
      throw new XmlWriteError("add attribute", "no element is open");
    }
    // Append to pending buffer — no write call
    this._pending += this._tags.addOne(name, value);
  }

  addAttributes(attributes: XmlAttributes): void {
    if (this.busy !== 0) {
      this._refuse("add attributes");
    }
    if (!this._open) {
      throw new XmlWriteError("add attributes", "no element is open");
    }
    this._pending += this._tags.addMap(attributes);
  }

  writeText(text: string | number | boolean): void {
    if (this.busy !== 0) {
      this._refuse("write text");
    }
    const encoded = xmlEncode(xmlValue(text, "#text"));
    this._leaf = false;
    this._writeContent(encoded);
  }

  writeRaw(xml: string): void {
    if (this.busy !== 0) {
      this._refuse("write raw");
    }
    const markup = rawMarkup(xml);
    this._leaf = false;
    this._writeContent(markup);
  }

  writeCData(text: string): void {
    if (this.busy !== 0) {
      this._refuse("write CDATA");
    }
    const markup = cdataMarkup(text);
    this._leaf = false;
    this._writeContent(markup);
  }

  writeComment(text: string): void {
    if (this.busy !== 0) {
      this._refuse("write comment");
    }
    const markup = commentMarkup(text);
    this._leaf = false;
    this._writeContent(markup);
  }

  closeNode(): void {
    if (this.busy !== 0) {
      this._refuse("close node");
    }
    const name = this._tags.close();
    if (name === undefined) {
      throw new XmlWriteError("close node", "no element is open");
    }
    if (this._leaf && this._open) {
      // Self-closing: flush pending tag with "/>", single write
      this._flushOpen("/>");
    } else if (this._open) {
      this._flushOpen(`></${name}>`);
    } else {
      this._write(`</${name}>`);
    }
    this._open = false;
    this._leaf = false;
  }

  leafNode(name: string, attributes?: XmlAttributes, text?: string | number | boolean): void {
    if (this.busy !== 0) {
      this._refuse("write leaf node");
    }
    validateXmlName(name);
    // Build the complete leaf element as one string, before flushing anything.
    let s =
      attributes === undefined ? `<${name}` : `<${name}${this._tags.startTag(name, attributes)}`;
    if (text !== undefined) {
      s += `>${xmlEncode(xmlValue(text, "#text"))}</${name}>`;
    } else {
      s += "/>";
    }
    this._writeContent(s);
    this._leaf = false;
  }

  // ===========================================================================
  // Close All
  // ===========================================================================

  /** Close all open elements. */
  closeAll(): void {
    if (this.busy !== 0) {
      this._refuse("close all");
    }
    while (this._tags.depth > 0) {
      this.closeNode();
    }
  }
}

export { XmlStreamWriter };
