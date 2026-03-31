import { xmlEncode } from "@xml/encode";

/**
 * Reusable utility for collecting raw XML fragments during SAX parsing.
 *
 * Many pivot-table xforms need to capture entire sub-trees (extLst, formats,
 * conditionalFormats, filters, fieldGroup, unknown elements, …) as raw strings
 * for roundtrip preservation.  The pattern — `active` flag, `depth` counter,
 * `buffer` of string fragments, plus identical `feedOpen` / `feedClose` /
 * `feedText` logic — was previously copy-pasted across multiple files.
 *
 * Usage:
 * ```ts
 * const collector = new RawXmlCollector("extLst");
 * // In parseOpen, when you see <extLst>:
 * collector.start();                      // opens with <extLst>
 * // For every child open tag while active:
 * collector.feedOpen(name, attributes);
 * // For every text node while active:
 * collector.feedText(text);
 * // For every close tag while active:
 * if (collector.feedClose(name)) { ... }  // returns true when root tag closed
 * // Retrieve result:
 * collector.result;                       // joined XML string
 * ```
 */
class RawXmlCollector {
  /** Tag name that this collector captures (e.g. "extLst", "formats"). */
  rootTag: string;
  /** Whether the collector is currently capturing. */
  active = false;
  /** Nesting depth *within* the root element (0 = direct children). */
  private depth = 0;
  /** String fragments being accumulated. */
  private buffer: string[] = [];
  /** Index of the last open-tag entry in the buffer (for self-closing collapse). */
  private lastOpenIndex = -1;

  constructor(rootTag: string) {
    this.rootTag = rootTag;
  }

  // ---------------------------------------------------------------------------
  // Lifecycle
  // ---------------------------------------------------------------------------

  /**
   * Begin collecting.  Pushes the opening root tag (with optional attributes)
   * and resets depth.
   *
   * @param attributes - Attributes on the root element (may be undefined/null).
   */
  start(attributes?: Record<string, unknown> | null): void {
    this.active = true;
    this.depth = 0;
    const attrsStr = serializeAttributes(attributes);
    this.buffer = [`<${this.rootTag}${attrsStr ? " " + attrsStr : ""}>`];
  }

  /**
   * Begin collecting with a dynamically determined root tag.
   * Used for catch-all unknown element collectors where the tag name
   * is not known at construction time.
   */
  startAs(rootTag: string, attributes?: Record<string, unknown> | null): void {
    this.rootTag = rootTag;
    this.start(attributes);
  }

  /**
   * Reset the collector to its initial idle state.
   */
  reset(): void {
    this.active = false;
    this.depth = 0;
    this.buffer.length = 0;
    this.lastOpenIndex = -1;
  }

  // ---------------------------------------------------------------------------
  // Feed methods — call these from parseOpen / parseText / parseClose
  // ---------------------------------------------------------------------------

  /**
   * Feed an open-tag event.  Must only be called while `active` is true.
   */
  feedOpen(name: string, attributes?: Record<string, unknown> | null): void {
    const attrsStr = serializeAttributes(attributes);
    this.buffer.push(`<${name}${attrsStr ? " " + attrsStr : ""}>`);
    this.depth++;
    this.lastOpenIndex = this.buffer.length - 1;
  }

  /**
   * Feed a text-node event.  Must only be called while `active` is true.
   */
  feedText(text: string): void {
    this.buffer.push(xmlEncode(text));
    this.lastOpenIndex = -1;
  }

  /**
   * Feed a close-tag event.
   *
   * @returns `true` when the **root** close tag has been received (collector
   *          deactivates itself and the result is ready).  `false` for any
   *          nested close tag.
   */
  feedClose(name: string): boolean {
    if (name === this.rootTag && this.depth === 0) {
      this.buffer.push(`</${name}>`);
      this.active = false;
      this.lastOpenIndex = -1;
      return true;
    }
    // Guard against depth going negative (e.g. mismatched close tags)
    if (this.depth > 0) {
      this.depth--;
    }
    // Collapse self-closing: if the close tag matches the last open tag and
    // nothing was written between them, rewrite `<tag ...>` → `<tag ... />`
    if (this.lastOpenIndex >= 0 && this.lastOpenIndex === this.buffer.length - 1) {
      const openTag = this.buffer[this.lastOpenIndex];
      // openTag ends with ">" — replace with " />"
      this.buffer[this.lastOpenIndex] = openTag.slice(0, -1) + " />";
      this.lastOpenIndex = -1;
      return false;
    }
    this.lastOpenIndex = -1;
    this.buffer.push(`</${name}>`);
    return false;
  }

  // ---------------------------------------------------------------------------
  // Result
  // ---------------------------------------------------------------------------

  /** The collected XML string.  Only meaningful after `feedClose` returns true. */
  get result(): string {
    return this.buffer.join("");
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Serialize an attributes object to an XML attribute string.
 * `null`, `undefined`, and empty objects produce an empty string.
 */
function serializeAttributes(attributes?: Record<string, unknown> | null): string {
  if (!attributes) {
    return "";
  }
  const entries = Object.entries(attributes);
  if (entries.length === 0) {
    return "";
  }
  return entries
    .filter(([, v]) => v != null)
    .map(([k, v]) => `${k}="${xmlEncode(String(v))}"`)
    .join(" ");
}

export { RawXmlCollector, serializeAttributes };
