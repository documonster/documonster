/**
 * DOCX Module - Template Data Source
 *
 * An abstraction layer for providing data to the template engine from
 * various sources (JSON, XML, CSV) with composition support.
 *
 * @stability experimental
 */

import { DocxError } from "@word/errors";
import { fillTemplate } from "@word/template/template-engine";
import type { TemplateOptions } from "@word/template/template-engine";
import type { DocxDocument } from "@word/types";
import { parseXml } from "@xml/dom";
import type { XmlElement } from "@xml/types";

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
      let parsed: unknown;
      try {
        parsed = JSON.parse(input);
      } catch (cause) {
        throw new DocxError(
          `JsonDataSource: failed to parse input as JSON (${(cause as Error).message})`,
          { cause }
        );
      }
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
 * Parse an XML string into a flat Record using the project's hardened
 * SAX-backed DOM parser. The parser enforces the standard `maxDepth` and
 * `maxEntityExpansions` limits, so this helper is safe to call on
 * untrusted input — earlier versions used a hand-rolled regex parser
 * with no depth limit and no protection against billion-laughs / nested
 * element bombs.
 */
function parseXmlToRecord(xml: string, rootTag?: string): Record<string, unknown> {
  let parsed: { root: XmlElement } | undefined;
  try {
    parsed = parseXml(xml);
  } catch {
    return {};
  }
  const root = parsed.root;
  if (!root) {
    return {};
  }
  // The previous regex parser ignored rootTag too; we keep that behaviour
  // for compatibility — converting from `root` directly works regardless
  // of whether the caller named the expected root.
  void rootTag;
  return xmlElementToRecord(root);
}

/** Adapter from XmlElement (DOM) to the legacy XmlNode shape. */
function elementToXmlNode(el: XmlElement): XmlNode {
  const children: XmlNode[] = [];
  let text = "";
  for (const child of el.children) {
    if (child.type === "element") {
      children.push(elementToXmlNode(child));
    } else if (child.type === "text" || child.type === "cdata") {
      text += child.value;
    }
  }
  return {
    tag: el.local ?? el.name.replace(/^.*:/, ""),
    attributes: { ...el.attributes },
    children,
    text
  };
}

function xmlElementToRecord(el: XmlElement): Record<string, unknown> {
  return xmlNodeToRecord(elementToXmlNode(el));
}

/** Stub for the legacy entrypoint kept for tests / external callers. */
function parseXmlElement(xml: string): XmlNode | null {
  try {
    return elementToXmlNode(parseXml(xml).root);
  } catch {
    return null;
  }
}
// Suppress no-unused-vars lint — kept for binary compat with prior internal API.
void parseXmlElement;

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
