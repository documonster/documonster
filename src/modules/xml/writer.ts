/**
 * XmlWriter - Non-Streaming XML Builder
 *
 * Builds XML as an in-memory string. Implements the {@link XmlSink} interface
 * so rendering code can target either this or {@link XmlStreamWriter}.
 *
 * Uses array-based buffering (push + join) for O(n) concatenation performance,
 * with rollback support via snapshot/restore.
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
import type { XmlAttributes, XmlSink } from "@xml/types";

// =============================================================================
// Internal Helpers
// =============================================================================

function pushAttributes(parts: string[], attributes?: XmlAttributes): void {
  if (!attributes) {
    return;
  }
  for (const key in attributes) {
    const value = attributes[key];
    if (value !== undefined) {
      validateXmlName(key);
      parts.push(` ${key}="${xmlEncodeAttr(String(value))}"`);
    }
  }
}

// =============================================================================
// Rollback Snapshot
// =============================================================================

interface Snapshot {
  partsLength: number;
  chunksLength: number;
  stackLength: number;
  leaf: boolean;
  open: boolean;
}

// =============================================================================
// XmlWriter
// =============================================================================

/**
 * Non-streaming XML writer that builds a complete XML string in memory.
 *
 * @example
 * ```ts
 * const w = new XmlWriter();
 * w.openXml();
 * w.openNode("root", { version: "1.0" });
 * w.leafNode("child", { id: "1" }, "text");
 * w.closeNode();
 * console.log(w.toString());
 * // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
 * // <root version="1.0"><child id="1">text</child></root>
 * ```
 */
class XmlWriter implements XmlSink {
  private _parts: string[] = [];
  private _chunks: string[] = [];
  private _stack: string[] = [];
  private _snapshots: Snapshot[] = [];
  private _leaf = false;
  private _open = false;

  /** Periodically consolidate small strings to reduce final join overhead. */
  private _consolidate(): void {
    // Not safe during rollback — joining fragments cannot be undone
    if (this._snapshots.length > 0) {
      return;
    }
    if (this._parts.length >= 10000) {
      this._chunks.push(this._parts.join(""));
      this._parts.length = 0;
    }
  }

  // ===========================================================================
  // State Queries
  // ===========================================================================

  /** Current nesting depth (number of open elements). */
  get depth(): number {
    return this._stack.length;
  }

  /** Name of the innermost open element, or undefined if none. */
  get currentElement(): string | undefined {
    return this._stack.length > 0 ? this._stack[this._stack.length - 1] : undefined;
  }

  /**
   * Monotonic cursor that tracks how much content has been written.
   * Useful for detecting whether a section produced any output.
   */
  get cursor(): number {
    return this._chunks.length * 10000 + this._parts.length;
  }

  // ===========================================================================
  // XmlSink Implementation
  // ===========================================================================

  openXml(attributes?: XmlAttributes): void {
    const merged = attributes ? { ...StdDocAttributes, ...attributes } : StdDocAttributes;
    this._parts.push("<?xml");
    pushAttributes(this._parts, merged);
    this._parts.push("?>\n");
  }

  openNode(name: string, attributes?: XmlAttributes): void {
    validateXmlName(name);
    if (this._open) {
      this._parts.push(">");
    }
    this._stack.push(name);
    // Build complete open tag as single string — reduces array push overhead
    let s = "<" + name;
    if (attributes) {
      for (const key in attributes) {
        const value = (attributes as any)[key];
        if (value !== undefined) {
          validateXmlName(key);
          s += ` ${key}="${xmlEncodeAttr(String(value))}"`;
        }
      }
    }
    this._parts.push(s);
    this._leaf = true;
    this._open = true;
  }

  addAttribute(name: string, value: string | number | boolean): void {
    if (!this._open) {
      throw new XmlWriteError("add attribute", "no element is open");
    }
    validateXmlName(name);
    this._parts.push(` ${name}="${xmlEncodeAttr(String(value))}"`);
  }

  addAttributes(attributes: XmlAttributes): void {
    if (!this._open) {
      throw new XmlWriteError("add attributes", "no element is open");
    }
    pushAttributes(this._parts, attributes);
  }

  writeText(text: string | number): void {
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    this._leaf = false;
    this._parts.push(xmlEncode(String(text)));
  }

  writeRaw(xml: string): void {
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    this._leaf = false;
    this._parts.push(xml);
  }

  writeCData(text: string): void {
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    this._leaf = false;
    this._parts.push(encodeCData(text));
  }

  writeComment(text: string): void {
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    this._leaf = false;
    validateCommentText(text);
    this._parts.push(`<!--${text}-->`);
  }

  closeNode(): void {
    const name = this._stack.pop();
    if (name === undefined) {
      throw new XmlWriteError("close node", "no element is open");
    }
    if (this._leaf) {
      this._parts.push("/>");
    } else {
      this._parts.push(`</${name}>`);
    }
    this._open = false;
    this._leaf = false;
    this._consolidate();
  }

  leafNode(name: string, attributes?: XmlAttributes, text?: string | number): void {
    validateXmlName(name);
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    // Build complete leaf element as single string — reduces 3-7 pushes to 1
    let s = "<" + name;
    if (attributes) {
      for (const key in attributes) {
        const value = (attributes as any)[key];
        if (value !== undefined) {
          validateXmlName(key);
          s += ` ${key}="${xmlEncodeAttr(String(value))}"`;
        }
      }
    }
    if (text !== undefined) {
      s += ">" + xmlEncode(String(text)) + "</" + name + ">";
    } else {
      s += "/>";
    }
    this._parts.push(s);
    this._leaf = false;
  }

  // ===========================================================================
  // Close All & Output
  // ===========================================================================

  /** Close all open elements. */
  closeAll(): void {
    while (this._stack.length > 0) {
      this.closeNode();
    }
  }

  /** Return the built XML string. */
  toString(): string {
    if (this._chunks.length === 0) {
      return this._parts.join("");
    }
    return this._chunks.join("") + this._parts.join("");
  }

  /** Alias for toString(). */
  get xml(): string {
    return this.toString();
  }

  // ===========================================================================
  // Rollback / Transaction Support
  // ===========================================================================

  /**
   * Save a snapshot of the current writer state.
   * Call {@link commit} to discard the snapshot or {@link rollback} to restore it.
   *
   * Snapshots can be nested.
   */
  save(): void {
    this._snapshots.push({
      partsLength: this._parts.length,
      chunksLength: this._chunks.length,
      stackLength: this._stack.length,
      leaf: this._leaf,
      open: this._open
    });
  }

  /** Discard the most recent snapshot (keep current state). */
  commit(): void {
    if (this._snapshots.length === 0) {
      throw new XmlWriteError("commit", "no snapshot to commit");
    }
    this._snapshots.pop();
  }

  /** Restore to the most recent snapshot (discard changes since save). */
  rollback(): void {
    if (this._snapshots.length === 0) {
      throw new XmlWriteError("rollback", "no snapshot to rollback");
    }
    const snap = this._snapshots.pop()!;
    this._parts.length = snap.partsLength;
    this._chunks.length = snap.chunksLength;
    this._stack.length = snap.stackLength;
    this._leaf = snap.leaf;
    this._open = snap.open;
  }

  // ===========================================================================
  // Reset
  // ===========================================================================

  /** Reset the writer to its initial empty state. */
  reset(): void {
    this._parts = [];
    this._chunks = [];
    this._stack = [];
    this._snapshots = [];
    this._leaf = false;
    this._open = false;
  }
}

export { XmlWriter, StdDocAttributes };
