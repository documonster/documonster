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

import { xmlEncode, xmlEncodeAttr } from "@xml/encode";
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
  // XmlSink Implementation
  // ===========================================================================

  openXml(attributes?: XmlAttributes): void {
    const defaults: XmlAttributes = { version: "1.0", encoding: "UTF-8", standalone: "yes" };
    const merged = attributes ? { ...defaults, ...attributes } : defaults;
    this._target.write("<?xml");
    this._writeAttributes(merged);
    this._target.write("?>\n");
  }

  openNode(name: string, attributes?: XmlAttributes): void {
    if (this._open) {
      this._target.write(">");
    }
    this._stack.push(name);
    this._target.write(`<${name}`);
    this._writeAttributes(attributes);
    this._leaf = true;
    this._open = true;
  }

  addAttribute(name: string, value: string | number | boolean): void {
    if (!this._open) {
      throw new XmlWriteError("add attribute", "no element is open");
    }
    this._target.write(` ${name}="${xmlEncodeAttr(String(value))}"`);
  }

  addAttributes(attributes: XmlAttributes): void {
    if (!this._open) {
      throw new XmlWriteError("add attributes", "no element is open");
    }
    this._writeAttributes(attributes);
  }

  writeText(text: string | number): void {
    if (this._open) {
      this._target.write(">");
      this._open = false;
    }
    this._leaf = false;
    this._target.write(xmlEncode(String(text)));
  }

  writeRaw(xml: string): void {
    if (this._open) {
      this._target.write(">");
      this._open = false;
    }
    this._leaf = false;
    this._target.write(xml);
  }

  writeCData(text: string): void {
    if (this._open) {
      this._target.write(">");
      this._open = false;
    }
    this._leaf = false;
    // Split on ]]> to produce valid CDATA — the sequence ]]> cannot appear inside CDATA.
    this._target.write("<![CDATA[" + text.split("]]>").join("]]]]><![CDATA[>") + "]]>");
  }

  writeComment(text: string): void {
    if (this._open) {
      this._target.write(">");
      this._open = false;
    }
    this._leaf = false;
    // XML spec: comments must not contain "--" and must not end with "-".
    if (text.includes("--") || text.endsWith("-")) {
      throw new XmlWriteError(
        "write comment",
        'comment text must not contain "--" or end with "-"'
      );
    }
    this._target.write(`<!--${text}-->`);
  }

  closeNode(): void {
    const name = this._stack.pop();
    if (name === undefined) {
      throw new XmlWriteError("close node", "no element is open");
    }
    if (this._leaf) {
      this._target.write("/>");
    } else {
      this._target.write(`</${name}>`);
    }
    this._open = false;
    this._leaf = false;
  }

  leafNode(name: string, attributes?: XmlAttributes, text?: string | number): void {
    this.openNode(name, attributes);
    if (text !== undefined) {
      this.writeText(text);
    }
    this.closeNode();
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

  // ===========================================================================
  // Internal Helpers
  // ===========================================================================

  private _writeAttributes(attributes?: XmlAttributes): void {
    if (!attributes) {
      return;
    }
    for (const key in attributes) {
      if (!Object.prototype.hasOwnProperty.call(attributes, key)) {
        continue;
      }
      const value = attributes[key];
      if (value !== undefined) {
        this._target.write(` ${key}="${xmlEncodeAttr(String(value))}"`);
      }
    }
  }
}

export { XmlStreamWriter };
