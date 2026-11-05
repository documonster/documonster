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
  xmlValue,
  cdataMarkup,
  rawMarkup,
  commentMarkup,
  serializeXmlAttributes,
  checkNewAttributes,
  StdDocAttributes
} from "@xml/encode";
import { XmlWriteError } from "@xml/errors";
import type { XmlAttributes, XmlSink } from "@xml/types";

// =============================================================================
// Internal Helpers
// =============================================================================

// =============================================================================
// Rollback Snapshot
// =============================================================================

interface Snapshot {
  partsLength: number;
  chunksLength: number;
  /**
   * The open elements, copied: an element closed after `snapshot()` has left the stack, so
   * restoring only its length brought back a hole where the name was.
   */
  stack: string[];
  leaf: boolean;
  open: boolean;
  openTagStart: number;
}

// =============================================================================
// XmlWriter
// =============================================================================

/**
 * Non-streaming XML writer that builds a complete XML string in memory.
 *
 * @example
 * ```ts
 * const w = new Xml.Writer();
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
  /** How many parts were joined into `_chunks`, so `cursor` keeps counting them. */
  private _consolidatedParts = 0;
  private _stack: string[] = [];
  private _snapshots: Snapshot[] = [];
  private _leaf = false;
  private _open = false;
  /** Index in `_parts` where the open element's start tag begins. */
  private _openTagStart = 0;

  /** Periodically consolidate small strings to reduce final join overhead. */
  private _consolidate(): void {
    // Not safe during rollback — joining fragments cannot be undone
    if (this._snapshots.length > 0) {
      return;
    }
    if (this._parts.length >= 10000) {
      this._chunks.push(this._parts.join(""));
      this._consolidatedParts += this._parts.length;
      this._parts.length = 0;
    }
  }

  /** The open element's start tag so far: its name and every attribute written to it. */
  private _openTag(): string {
    return this._parts.slice(this._openTagStart).join("");
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
   * A position that grows with every write and moves back only on {@link rollback}.
   * Compare it before and after a section to learn whether the section wrote anything.
   */
  get cursor(): number {
    return this._consolidatedParts + this._parts.length;
  }

  // ===========================================================================
  // XmlSink Implementation
  // ===========================================================================

  /**
   * Write the XML declaration. XML allows it only at the very start of a document, so it is
   * refused once anything has been written, including an earlier declaration.
   */
  openXml(attributes?: XmlAttributes): void {
    if (this.cursor > 0) {
      throw new XmlWriteError("open xml", "the XML declaration must come first in the document");
    }
    const merged = attributes ? { ...StdDocAttributes, ...attributes } : StdDocAttributes;
    this._parts.push(`<?xml${serializeXmlAttributes(merged)}?>\n`);
  }

  openNode(name: string, attributes?: XmlAttributes): void {
    validateXmlName(name);
    // Build the complete open tag first, so a refused attribute leaves the writer unchanged.
    const s = "<" + name + serializeXmlAttributes(attributes);
    if (this._open) {
      this._parts.push(">");
    }
    this._stack.push(name);
    this._openTagStart = this._parts.length;
    this._parts.push(s);
    this._leaf = true;
    this._open = true;
  }

  addAttribute(name: string, value: string | number | boolean): void {
    if (!this._open) {
      throw new XmlWriteError("add attribute", "no element is open");
    }
    validateXmlName(name);
    checkNewAttributes(this.currentElement, this._openTag(), { [name]: value }, () => "name");
    this._parts.push(` ${name}="${xmlEncodeAttr(xmlValue(value, name, "value"))}"`);
  }

  addAttributes(attributes: XmlAttributes): void {
    if (!this._open) {
      throw new XmlWriteError("add attributes", "no element is open");
    }
    checkNewAttributes(this.currentElement, this._openTag(), attributes);
    this._parts.push(serializeXmlAttributes(attributes));
  }

  writeText(text: string | number): void {
    const encoded = xmlEncode(xmlValue(text, "#text"));
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    this._leaf = false;
    this._parts.push(encoded);
  }

  writeRaw(xml: string): void {
    const markup = rawMarkup(xml);
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    this._leaf = false;
    this._parts.push(markup);
  }

  writeCData(text: string): void {
    const markup = cdataMarkup(text);
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    this._leaf = false;
    this._parts.push(markup);
  }

  writeComment(text: string): void {
    const markup = commentMarkup(text);
    if (this._open) {
      this._parts.push(">");
      this._open = false;
    }
    this._leaf = false;
    this._parts.push(markup);
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
    // Build the complete leaf element as one string before touching any state.
    let s = "<" + name + serializeXmlAttributes(attributes);
    if (text !== undefined) {
      s += ">" + xmlEncode(xmlValue(text, "#text")) + "</" + name + ">";
    } else {
      s += "/>";
    }
    if (this._open) {
      this._parts.push(">");
      this._open = false;
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

  /** @deprecated Use {@link toString}. Removed in 1.0. */
  get xml(): string {
    return this.toString();
  }

  // ===========================================================================
  // Rollback / Transaction Support
  // ===========================================================================

  /**
   * Take a snapshot of the current writer state.
   * Call {@link commit} to keep what was written since and drop the snapshot, or
   * {@link rollback} to return to it — as `Streams.ChunkedBuilder` marks, keeps and undoes.
   *
   * Snapshots can be nested.
   */
  snapshot(): void {
    this._snapshots.push({
      partsLength: this._parts.length,
      chunksLength: this._chunks.length,
      stack: this._stack.slice(),
      leaf: this._leaf,
      open: this._open,
      openTagStart: this._openTagStart
    });
  }

  /** @deprecated Use {@link snapshot}. Removed in 1.0. */
  save(): void {
    this.snapshot();
  }

  /** Discard the most recent snapshot (keep current state). */
  commit(): void {
    if (this._snapshots.length === 0) {
      throw new XmlWriteError("commit", "no snapshot to commit");
    }
    this._snapshots.pop();
  }

  /** Restore to the most recent snapshot (discard changes since {@link snapshot}). */
  rollback(): void {
    if (this._snapshots.length === 0) {
      throw new XmlWriteError("rollback", "no snapshot to rollback");
    }
    const snap = this._snapshots.pop()!;
    this._parts.length = snap.partsLength;
    this._chunks.length = snap.chunksLength;
    this._stack = snap.stack;
    this._leaf = snap.leaf;
    this._open = snap.open;
    this._openTagStart = snap.openTagStart;
  }

  // ===========================================================================
  // Reset
  // ===========================================================================

  /** Reset the writer to its initial empty state. */
  reset(): void {
    this._parts = [];
    this._chunks = [];
    this._consolidatedParts = 0;
    this._stack = [];
    this._snapshots = [];
    this._leaf = false;
    this._open = false;
    this._openTagStart = 0;
  }
}

export { XmlWriter, StdDocAttributes };
