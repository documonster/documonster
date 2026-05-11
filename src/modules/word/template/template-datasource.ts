/**
 * DOCX Module - Template Data Source
 *
 * An abstraction layer for providing data to the template engine from
 * various sources (JSON, XML, CSV) with composition support.
 *
 * @stability experimental
 */

import { DocxError } from "../errors";
import type { DocxDocument } from "../types";
import { fillTemplate } from "./template-engine";
import type { TemplateOptions } from "./template-engine";

// =============================================================================
// DataSource Interface
// =============================================================================

/**
 * Abstract data provider for the template engine.
 *
 * Implementations transform various data formats into the flat
 * `Record<string, unknown>` structure consumed by `fillTemplate`.
 */
export interface DataSource {
  /** Retrieve all data as a flat record. */
  readonly getData: () => Record<string, unknown>;

  /** Retrieve an array value by key. Returns empty array if not found or not an array. */
  readonly getArray: (key: string) => unknown[];

  /** Retrieve a single value by dot-separated path. Returns undefined if not found. */
  readonly getValue: (path: string) => unknown;
}

// =============================================================================
// Options
// =============================================================================

/** Options for `fillTemplateFromSource`. */
export interface FillFromSourceOptions extends TemplateOptions {
  /** If true, merge arrays with the same key across composite sources (default: false). */
  readonly mergeArrays?: boolean;
}

// =============================================================================
// JsonDataSource
// =============================================================================

/**
 * Data source that reads from a JSON string or object.
 *
 * @stability experimental
 */
export class JsonDataSource implements DataSource {
  private readonly data: Record<string, unknown>;

  /**
   * @param input - A JSON string or an existing object.
   */
  constructor(input: string | Record<string, unknown>) {
    if (typeof input === "string") {
      const parsed = JSON.parse(input) as unknown;
      if (parsed === null || typeof parsed !== "object" || Array.isArray(parsed)) {
        throw new DocxError(
          "JsonDataSource: input must parse to a JSON object (not array or null)"
        );
      }
      this.data = parsed as Record<string, unknown>;
    } else {
      this.data = input;
    }
  }

  getData(): Record<string, unknown> {
    return this.data;
  }

  getArray(key: string): unknown[] {
    const value = this.resolvePath(key);
    if (Array.isArray(value)) {
      return value;
    }
    return [];
  }

  getValue(path: string): unknown {
    return this.resolvePath(path);
  }

  private resolvePath(path: string): unknown {
    const parts = path.split(".");
    let current: unknown = this.data;
    for (const part of parts) {
      if (current === null || current === undefined || typeof current !== "object") {
        return undefined;
      }
      current = (current as Record<string, unknown>)[part];
    }
    return current;
  }
}

// =============================================================================
// XmlDataSource
// =============================================================================

/**
 * Data source that reads from an XML string.
 *
 * Flattens the XML structure into dot-separated key-value pairs.
 * Repeated elements with the same tag are collected into arrays.
 *
 * Example XML:
 * ```xml
 * <root>
 *   <name>John</name>
 *   <address><city>NY</city></address>
 *   <item>A</item>
 *   <item>B</item>
 * </root>
 * ```
 * Produces: `{ name: "John", "address.city": "NY", address: { city: "NY" }, item: ["A", "B"] }`
 *
 * @stability experimental
 */
export class XmlDataSource implements DataSource {
  private readonly data: Record<string, unknown>;

  /**
   * @param xml - An XML string to parse.
   * @param rootTag - Optional root element name to skip. If not provided, the first element is used.
   */
  constructor(xml: string, rootTag?: string) {
    this.data = parseXmlToRecord(xml, rootTag);
  }

  getData(): Record<string, unknown> {
    return this.data;
  }

  getArray(key: string): unknown[] {
    const value = this.resolvePath(key);
    if (Array.isArray(value)) {
      return value;
    }
    return [];
  }

  getValue(path: string): unknown {
    return this.resolvePath(path);
  }

  private resolvePath(path: string): unknown {
    const parts = path.split(".");
    let current: unknown = this.data;
    for (const part of parts) {
      if (current === null || current === undefined || typeof current !== "object") {
        return undefined;
      }
      current = (current as Record<string, unknown>)[part];
    }
    return current;
  }
}

// =============================================================================
// CsvDataSource
// =============================================================================

/**
 * Data source that reads from a CSV string.
 *
 * The first row is treated as headers (column names become keys).
 * Each subsequent row becomes an object. The entire dataset is available
 * as an array under the `rows` key (or a custom key).
 *
 * Individual columns are also available as arrays under their header name.
 *
 * Example CSV:
 * ```
 * name,age
 * Alice,30
 * Bob,25
 * ```
 * Produces:
 * ```
 * {
 *   rows: [{ name: "Alice", age: "30" }, { name: "Bob", age: "25" }],
 *   name: ["Alice", "Bob"],
 *   age: ["30", "25"]
 * }
 * ```
 *
 * @stability experimental
 */
export class CsvDataSource implements DataSource {
  private readonly data: Record<string, unknown>;

  /**
   * @param csv - A CSV string.
   * @param options - Optional parsing configuration.
   */
  constructor(
    csv: string,
    options?: {
      /** Delimiter character (default: ","). */
      readonly delimiter?: string;
      /** Key for the rows array (default: "rows"). */
      readonly rowsKey?: string;
    }
  ) {
    const delimiter = options?.delimiter ?? ",";
    const rowsKey = options?.rowsKey ?? "rows";
    this.data = parseCsvToRecord(csv, delimiter, rowsKey);
  }

  getData(): Record<string, unknown> {
    return this.data;
  }

  getArray(key: string): unknown[] {
    const value = this.data[key];
    if (Array.isArray(value)) {
      return value;
    }
    return [];
  }

  getValue(path: string): unknown {
    const parts = path.split(".");
    let current: unknown = this.data;
    for (const part of parts) {
      if (current === null || current === undefined || typeof current !== "object") {
        return undefined;
      }
      current = (current as Record<string, unknown>)[part];
    }
    return current;
  }
}

// =============================================================================
// CompositeDataSource
// =============================================================================

/**
 * Combines multiple data sources into one.
 *
 * Sources added later take precedence over earlier ones for conflicting keys.
 * Arrays can optionally be merged instead of overwritten.
 *
 * @stability experimental
 */
export class CompositeDataSource implements DataSource {
  private readonly sources: readonly DataSource[];
  private readonly mergeArrays: boolean;

  /**
   * @param sources - Data sources to combine (later sources take precedence).
   * @param options - Configuration options.
   */
  constructor(sources: readonly DataSource[], options?: { readonly mergeArrays?: boolean }) {
    this.sources = sources;
    this.mergeArrays = options?.mergeArrays ?? false;
  }

  getData(): Record<string, unknown> {
    const result: Record<string, unknown> = {};
    for (const source of this.sources) {
      const data = source.getData();
      for (const key of Object.keys(data)) {
        if (this.mergeArrays && Array.isArray(result[key]) && Array.isArray(data[key])) {
          result[key] = [...(result[key] as unknown[]), ...(data[key] as unknown[])];
        } else {
          result[key] = data[key];
        }
      }
    }
    return result;
  }

  getArray(key: string): unknown[] {
    if (this.mergeArrays) {
      const merged: unknown[] = [];
      for (const source of this.sources) {
        const arr = source.getArray(key);
        merged.push(...arr);
      }
      return merged;
    }
    // Last source with a non-empty array wins
    for (let i = this.sources.length - 1; i >= 0; i--) {
      const arr = this.sources[i].getArray(key);
      if (arr.length > 0) {
        return arr;
      }
    }
    return [];
  }

  getValue(path: string): unknown {
    // Last source with a defined value wins
    for (let i = this.sources.length - 1; i >= 0; i--) {
      const value = this.sources[i].getValue(path);
      if (value !== undefined) {
        return value;
      }
    }
    return undefined;
  }
}

// =============================================================================
// fillTemplateFromSource
// =============================================================================

/**
 * Fill a DOCX template using a DataSource instead of a raw data record.
 *
 * This is a convenience function that extracts data from the source and
 * delegates to the core `fillTemplate` function.
 *
 * @param doc - The parsed DocxDocument model.
 * @param source - A DataSource implementation.
 * @param options - Optional template settings.
 * @returns The same DocxDocument with placeholders resolved.
 *
 * @stability experimental
 */
export function fillTemplateFromSource(
  doc: DocxDocument,
  source: DataSource,
  options?: FillFromSourceOptions
): DocxDocument {
  const data = source.getData();
  return fillTemplate(doc, data, options);
}

// =============================================================================
// Internal: XML Parser (minimal, zero-dependency)
// =============================================================================

interface XmlNode {
  readonly tag: string;
  readonly attributes: Record<string, string>;
  readonly children: XmlNode[];
  readonly text: string;
}

/**
 * Parse an XML string into a flat Record.
 * Uses a simple regex-based tokenizer — not a full XML parser, but sufficient
 * for data extraction from well-formed XML documents.
 */
function parseXmlToRecord(xml: string, rootTag?: string): Record<string, unknown> {
  // Strip XML declaration and comments
  const cleaned = xml
    .replace(/<\?xml[^?]*\?>/g, "")
    .replace(/<!--[\s\S]*?-->/g, "")
    .trim();

  // Parse into a tree
  const root = parseXmlElement(cleaned);
  if (!root) {
    return {};
  }

  // If rootTag is specified and matches, use its children scope; otherwise use root directly
  const startNode = root;

  // Convert tree to Record
  return xmlNodeToRecord(startNode);
}

/**
 * Parse a single XML element and its children from a string.
 * Returns the parsed node or null if the string doesn't start with an element.
 */
function parseXmlElement(xml: string): XmlNode | null {
  const trimmed = xml.trim();
  if (!trimmed.startsWith("<")) {
    return null;
  }

  // Match opening tag
  const openTagRegex = /^<([a-zA-Z_][\w:.-]*)((?:\s+[a-zA-Z_][\w:.-]*\s*=\s*"[^"]*")*)\s*(\/?)>/;
  const match = openTagRegex.exec(trimmed);
  if (!match) {
    return null;
  }

  const tag = match[1];
  const attrString = match[2];
  const selfClosing = match[3] === "/";

  // Parse attributes
  const attributes: Record<string, string> = {};
  if (attrString) {
    const attrRegex = /([a-zA-Z_][\w:.-]*)\s*=\s*"([^"]*)"/g;
    let attrMatch: RegExpExecArray | null;
    while ((attrMatch = attrRegex.exec(attrString)) !== null) {
      attributes[attrMatch[1]] = decodeXmlEntities(attrMatch[2]);
    }
  }

  if (selfClosing) {
    return { tag, attributes, children: [], text: "" };
  }

  // Find the content between open and close tags
  const afterOpen = trimmed.slice(match[0].length);
  const closeIdx = findMatchingCloseTag(afterOpen, tag);

  if (closeIdx === -1) {
    // Malformed — treat as self-closing
    return { tag, attributes, children: [], text: "" };
  }

  const innerContent = afterOpen.slice(0, closeIdx);

  // Parse children
  const children: XmlNode[] = [];
  let textContent = "";
  let pos = 0;

  while (pos < innerContent.length) {
    // Skip whitespace
    const nextTag = innerContent.indexOf("<", pos);
    if (nextTag === -1) {
      // Rest is text
      textContent += innerContent.slice(pos);
      break;
    }

    // Collect text before the next tag
    if (nextTag > pos) {
      textContent += innerContent.slice(pos, nextTag);
    }

    // Check if it's a CDATA section
    if (innerContent.startsWith("<![CDATA[", nextTag)) {
      const cdataEnd = innerContent.indexOf("]]>", nextTag + 9);
      if (cdataEnd !== -1) {
        textContent += innerContent.slice(nextTag + 9, cdataEnd);
        pos = cdataEnd + 3;
        continue;
      }
    }

    // Parse child element
    const childXml = innerContent.slice(nextTag);
    const child = parseXmlElement(childXml);
    if (child) {
      children.push(child);
      // Advance past this child element
      const childClose = `</${child.tag}>`;
      const selfCloseCheck = new RegExp(`^<${escapeRegexPattern(child.tag)}(?:\\s+[^>]*)?\\/>`);
      if (selfCloseCheck.test(childXml.trim())) {
        const selfCloseMatch = selfCloseCheck.exec(childXml.trim());
        pos = nextTag + (selfCloseMatch ? selfCloseMatch[0].length : 0);
        // Find actual position in innerContent
        const selfCloseEnd = innerContent.indexOf("/>", nextTag);
        if (selfCloseEnd !== -1) {
          pos = selfCloseEnd + 2;
        } else {
          pos = nextTag + 1;
        }
      } else {
        const childCloseIdx = findMatchingCloseTag(
          innerContent.slice(nextTag + 1 + child.tag.length),
          child.tag
        );
        if (childCloseIdx !== -1) {
          // Account for opening tag length
          const openEnd = innerContent.indexOf(">", nextTag);
          if (openEnd !== -1) {
            const afterChildOpen = openEnd + 1;
            const absCloseIdx = findMatchingCloseTag(innerContent.slice(afterChildOpen), child.tag);
            if (absCloseIdx !== -1) {
              pos = afterChildOpen + absCloseIdx + childClose.length;
            } else {
              pos = nextTag + 1;
            }
          } else {
            pos = nextTag + 1;
          }
        } else {
          pos = nextTag + 1;
        }
      }
    } else {
      // Not a valid element; skip this character
      pos = nextTag + 1;
    }
  }

  return { tag, attributes, children, text: decodeXmlEntities(textContent.trim()) };
}

/**
 * Find the position of the matching close tag, handling nested elements of the same name.
 */
function findMatchingCloseTag(content: string, tag: string): number {
  const openPattern = `<${tag}`;
  const closePattern = `</${tag}>`;
  let depth = 0;
  let pos = 0;

  while (pos < content.length) {
    const nextClose = content.indexOf(closePattern, pos);
    if (nextClose === -1) {
      return -1;
    }

    // Count opens between pos and nextClose
    let searchPos = pos;
    while (searchPos < nextClose) {
      const nextOpen = content.indexOf(openPattern, searchPos);
      if (nextOpen === -1 || nextOpen >= nextClose) {
        break;
      }
      // Verify it's actually an opening tag (followed by space, >, or /)
      const charAfter = content[nextOpen + openPattern.length];
      if (
        charAfter === ">" ||
        charAfter === " " ||
        charAfter === "/" ||
        charAfter === "\t" ||
        charAfter === "\n"
      ) {
        // Check it's not self-closing
        const tagEnd = content.indexOf(">", nextOpen);
        if (tagEnd !== -1 && content[tagEnd - 1] !== "/") {
          depth++;
        }
      }
      searchPos = nextOpen + openPattern.length;
    }

    if (depth === 0) {
      return nextClose;
    }
    depth--;
    pos = nextClose + closePattern.length;
  }

  return -1;
}

/** Convert an XmlNode tree into a flat Record. */
function xmlNodeToRecord(node: XmlNode): Record<string, unknown> {
  const result: Record<string, unknown> = {};

  // Add attributes with @ prefix
  for (const key of Object.keys(node.attributes)) {
    result[`@${key}`] = node.attributes[key];
  }

  // Group children by tag name
  const childGroups = new Map<string, XmlNode[]>();
  for (const child of node.children) {
    const existing = childGroups.get(child.tag);
    if (existing) {
      existing.push(child);
    } else {
      childGroups.set(child.tag, [child]);
    }
  }

  // Convert child groups
  for (const [tag, group] of childGroups) {
    if (group.length === 1) {
      const child = group[0];
      if (child.children.length === 0) {
        // Leaf node — use text value
        result[tag] = child.text || "";
      } else {
        // Nested object
        result[tag] = xmlNodeToRecord(child);
      }
    } else {
      // Multiple elements with same tag — array
      result[tag] = group.map(child => {
        if (child.children.length === 0) {
          return child.text || "";
        }
        return xmlNodeToRecord(child);
      });
    }
  }

  // If node has text and no children, include it
  if (node.children.length === 0 && node.text) {
    result["#text"] = node.text;
  }

  return result;
}

/** Decode basic XML entities and numeric character references. */
function decodeXmlEntities(str: string): string {
  return str
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

/** Escape special regex characters in a string. */
function escapeRegexPattern(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// =============================================================================
// Internal: CSV Parser (minimal, zero-dependency)
// =============================================================================

/**
 * Parse a CSV string into a Record with column arrays and a rows array.
 * Handles quoted fields (RFC 4180 style).
 */
function parseCsvToRecord(
  csv: string,
  delimiter: string,
  rowsKey: string
): Record<string, unknown> {
  const lines = parseCsvLines(csv, delimiter);
  if (lines.length === 0) {
    return { [rowsKey]: [] };
  }

  const headers = lines[0];
  const rows: Record<string, unknown>[] = [];
  const columns: Record<string, unknown[]> = {};

  // Initialize column arrays
  for (const header of headers) {
    columns[header] = [];
  }

  // Parse data rows
  for (let i = 1; i < lines.length; i++) {
    const fields = lines[i];
    const row: Record<string, unknown> = {};
    for (let j = 0; j < headers.length; j++) {
      const value = j < fields.length ? fields[j] : "";
      row[headers[j]] = value;
      columns[headers[j]].push(value);
    }
    rows.push(row);
  }

  const result: Record<string, unknown> = { [rowsKey]: rows };
  for (const header of headers) {
    result[header] = columns[header];
  }
  return result;
}

/**
 * Parse CSV text into an array of rows, where each row is an array of fields.
 * Supports quoted fields with embedded delimiters, newlines, and escaped quotes.
 */
function parseCsvLines(csv: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  const len = csv.length;
  let pos = 0;

  while (pos < len) {
    const { fields, nextPos } = parseCsvRow(csv, pos, delimiter);
    // Skip completely empty trailing lines
    if (nextPos >= len && fields.length === 1 && fields[0] === "") {
      break;
    }
    rows.push(fields);
    pos = nextPos;
  }

  return rows;
}

/**
 * Parse a single CSV row starting at the given position.
 * Returns the parsed fields and the position after the row terminator.
 */
function parseCsvRow(
  csv: string,
  startPos: number,
  delimiter: string
): { fields: string[]; nextPos: number } {
  const fields: string[] = [];
  const len = csv.length;
  let pos = startPos;

  while (pos <= len) {
    if (pos === len) {
      // End of input — add empty field only if we just saw a delimiter
      if (fields.length > 0 && pos > startPos && csv[pos - 1] === delimiter) {
        fields.push("");
      } else if (fields.length === 0) {
        fields.push("");
      }
      break;
    }

    const char = csv[pos];

    // Check for line terminator at start of field (empty row remainder)
    if (char === "\r" || char === "\n") {
      if (fields.length === 0) {
        // Empty line
        fields.push("");
      }
      // Skip line terminator
      if (char === "\r" && pos + 1 < len && csv[pos + 1] === "\n") {
        pos += 2;
      } else {
        pos += 1;
      }
      return { fields, nextPos: pos };
    }

    if (char === '"') {
      // Quoted field
      const { value, nextPos: fieldEnd } = parseQuotedField(csv, pos);
      fields.push(value);
      pos = fieldEnd;

      // After quoted field, expect delimiter or line end
      if (pos < len && csv[pos] === delimiter) {
        pos++; // skip delimiter
      } else if (pos < len && (csv[pos] === "\r" || csv[pos] === "\n")) {
        if (csv[pos] === "\r" && pos + 1 < len && csv[pos + 1] === "\n") {
          pos += 2;
        } else {
          pos += 1;
        }
        return { fields, nextPos: pos };
      }
    } else {
      // Unquoted field
      let fieldEnd = pos;
      while (
        fieldEnd < len &&
        csv[fieldEnd] !== delimiter &&
        csv[fieldEnd] !== "\r" &&
        csv[fieldEnd] !== "\n"
      ) {
        fieldEnd++;
      }
      fields.push(csv.slice(pos, fieldEnd));
      pos = fieldEnd;

      if (pos < len && csv[pos] === delimiter) {
        pos++; // skip delimiter
        // If delimiter is at end of input, add empty trailing field
        if (pos >= len) {
          fields.push("");
        }
      } else if (pos < len && (csv[pos] === "\r" || csv[pos] === "\n")) {
        if (csv[pos] === "\r" && pos + 1 < len && csv[pos + 1] === "\n") {
          pos += 2;
        } else {
          pos += 1;
        }
        return { fields, nextPos: pos };
      }
    }
  }

  return { fields, nextPos: pos };
}

/**
 * Parse a quoted CSV field starting at the given position (which must be a `"`).
 * Handles escaped quotes (`""`) inside the field.
 */
function parseQuotedField(csv: string, startPos: number): { value: string; nextPos: number } {
  const len = csv.length;
  let pos = startPos + 1; // skip opening quote
  let value = "";

  while (pos < len) {
    if (csv[pos] === '"') {
      if (pos + 1 < len && csv[pos + 1] === '"') {
        // Escaped quote
        value += '"';
        pos += 2;
      } else {
        // End of quoted field
        pos++; // skip closing quote
        return { value, nextPos: pos };
      }
    } else {
      value += csv[pos];
      pos++;
    }
  }

  // Unterminated quote — return what we have
  return { value, nextPos: pos };
}
