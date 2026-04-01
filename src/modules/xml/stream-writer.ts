/**
 * XmlStreamWriter - True Streaming XML Writer
 *
 * Writes XML directly to a {@link WritableTarget} without buffering
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
  encodeCData,
  validateCommentText,
  StdDocAttributes
} from "@xml/encode";
import { XmlWriteError } from "@xml/errors";
import type { WritableTarget, XmlAttributes, XmlSink } from "@xml/types";

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
 * const sw = new XmlStreamWriter(target);
 * sw.openXml();
 * sw.openNode("root");
 * sw.leafNode("item", { id: "1" }, "hello");
 * sw.closeNode();
 * console.log(chunks.join(""));
 * ```
 */
class XmlStreamWriter implements XmlSink {
  private _target: WritableTarget;
  private _stack: string[] = [];
  private _leaf = false;
  private _open = false;
  // Pending start-tag buffer: accumulates tag name + attributes until flushed
  private _pending = "";

  constructor(target: WritableTarget) {
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

  /** Flush the pending start-tag buffer with a closing character. */
  private _flushOpen(suffix: string): void {
    this._target.write(this._pending + suffix);
    this._pending = "";
    this._open = false;
  }

  // ===========================================================================
  // XmlSink Implementation
  // ===========================================================================

  openXml(attributes?: XmlAttributes): void {
    const merged = attributes ? { ...StdDocAttributes, ...attributes } : StdDocAttributes;
    let s = "<?xml";
    for (const key in merged) {
      const value = (merged as any)[key];
      if (value !== undefined) {
        validateXmlName(key);
        s += ` ${key}="${xmlEncodeAttr(String(value))}"`;
      }
    }
    this._target.write(s + "?>\n");
  }

  openNode(name: string, attributes?: XmlAttributes): void {
    validateXmlName(name);
    // Flush any pending open tag first
    if (this._open) {
      this._flushOpen(">");
    }
    // Start building pending tag
    let s = `<${name}`;
    if (attributes) {
      for (const key in attributes) {
        const value = attributes[key];
        if (value !== undefined) {
          validateXmlName(key);
          s += ` ${key}="${xmlEncodeAttr(String(value))}"`;
        }
      }
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
    // Append to pending buffer — no write call
    this._pending += ` ${name}="${xmlEncodeAttr(String(value))}"`;
  }

  addAttributes(attributes: XmlAttributes): void {
    if (!this._open) {
      throw new XmlWriteError("add attributes", "no element is open");
    }
    if (!attributes) {
      return;
    }
    for (const key in attributes) {
      const value = attributes[key];
      if (value !== undefined) {
        validateXmlName(key);
        this._pending += ` ${key}="${xmlEncodeAttr(String(value))}"`;
      }
    }
  }

  writeText(text: string | number): void {
    if (this._open) {
      this._flushOpen(">");
    }
    this._leaf = false;
    this._target.write(xmlEncode(String(text)));
  }

  writeRaw(xml: string): void {
    if (this._open) {
      this._flushOpen(">");
    }
    this._leaf = false;
    this._target.write(xml);
  }

  writeCData(text: string): void {
    if (this._open) {
      this._flushOpen(">");
    }
    this._leaf = false;
    this._target.write(encodeCData(text));
  }

  writeComment(text: string): void {
    if (this._open) {
      this._flushOpen(">");
    }
    this._leaf = false;
    validateCommentText(text);
    this._target.write(`<!--${text}-->`);
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
      this._target.write(`</${name}>`);
    } else {
      this._target.write(`</${name}>`);
    }
    this._open = false;
    this._leaf = false;
  }

  leafNode(name: string, attributes?: XmlAttributes, text?: string | number): void {
    validateXmlName(name);
    // Flush any pending open tag
    if (this._open) {
      this._flushOpen(">");
    }

    // Build complete leaf element as a single string, single write call
    let s = `<${name}`;
    if (attributes) {
      for (const key in attributes) {
        const value = attributes[key];
        if (value !== undefined) {
          validateXmlName(key);
          s += ` ${key}="${xmlEncodeAttr(String(value))}"`;
        }
      }
    }
    if (text !== undefined) {
      s += `>${xmlEncode(String(text))}</${name}>`;
    } else {
      s += "/>";
    }
    this._target.write(s);
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
