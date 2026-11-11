/**
 * DOCX Module - Flat OPC Format Support
 *
 * Implements reading and writing of Flat OPC (Office Open XML Package) format,
 * which represents an entire DOCX package as a single XML file.
 * This is useful for debugging, XML transformations, and enterprise integration scenarios.
 *
 * Flat OPC uses the `pkg:` namespace to wrap all package parts in a single XML document.
 * Each part is represented as a `<pkg:part>` element with its content either as XML or base64.
 *
 * References:
 * - [MS-OFFCRYPTO] §2.1.5 (pkgStart namespace)
 * - ECMA-376 Part 2 (OPC) Annex C (non-normative)
 */

import { xmlEncodeAttr } from "@utils/xml-encode";
import { NS_CONTENT_TYPES, NS_PKG } from "@word/constants";
import { checkBytes, invalidArgument } from "@word/core/argument-checks";
import { utf8Decoder, utf8Encoder } from "@word/core/internal-utils";
import { getFileExt } from "@word/core/opc-paths";
import { DocxParseError } from "@word/errors";
import { XmlParseError } from "@xml/errors";
import { SaxParser } from "@xml/sax";
import type { XmlSaxTag } from "@xml/types";
import { XmlWriter } from "@xml/writer";

// =============================================================================
// Types
// =============================================================================

/** A part extracted from Flat OPC XML. */
export interface FlatOpcPart {
  /** Part name (e.g. "/word/document.xml"). */
  readonly name: string;
  /** Content type of the part. */
  readonly contentType: string;
  /** Raw data (decoded from base64 for binary, or UTF-8 encoded XML). */
  readonly data: Uint8Array;
}

// =============================================================================
// Reading Flat OPC
// =============================================================================

/**
 * Parse a Flat OPC XML document into individual package parts.
 * Returns a Map keyed by part path (without leading slash) to binary data,
 * compatible with the archive entry format used by `readDocx`.
 *
 * @param xmlContent - The Flat OPC XML string or UTF-8 bytes.
 * @returns A map of part paths to their binary content.
 */
export function parseFlatOpc(
  xmlContent: string | Uint8Array | ArrayBuffer | ArrayBufferView
): Map<string, Uint8Array> {
  const given = checkTextOrBytes(xmlContent, "xmlContent", "Convert.flatOpcToParts");
  const bytes = typeof given === "string" ? utf8Encoder.encode(given) : given;
  // Text pasted or saved by an editor often begins with a byte order mark or white space before the
  // XML declaration, where XML allows neither: a Flat OPC file is read from its first `<`.
  let lead = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (lead < bytes.length && isSpaceByte(bytes[lead]!)) {
    lead++;
  }

  const reader = new FlatOpcReader(bytes);
  // With namespaces: a package element is one in the package namespace, whatever its prefix.
  const parser = new SaxParser({ position: false, xmlns: true, byteOffset: lead });
  parser.on("opentag", tag => reader.open(parser, tag));
  parser.on("closetag", tag => reader.close(parser, tag));
  parser.on("text", text => reader.text(parser, text, false));
  parser.on("cdata", text => reader.text(parser, text, true));
  try {
    parser.write(bytes.subarray(lead)).close();
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new DocxParseError(`Flat OPC: not well-formed XML: ${error.message}`, {
        cause: error
      });
    }
    throw error;
  }

  const { parts, contentTypes } = reader;
  // Ensure [Content_Types].xml exists
  if (!parts.has("[Content_Types].xml")) {
    parts.set("[Content_Types].xml", buildContentTypesXml(contentTypes));
  }
  return parts;
}

function isSpaceByte(byte: number): boolean {
  return byte === 0x20 || byte === 0x09 || byte === 0x0a || byte === 0x0d;
}

/** A part being read: where its `xmlData` content lies, or the text of its `binaryData`. */
interface OpenPart {
  readonly path: string;
  readonly contentType: string;
  data?: Uint8Array;
  /** `xmlData`: where its content begins in the input, and the namespaces in force there. */
  xml?: {
    readonly from: number;
    readonly scope: ReadonlyMap<string, string>;
    root?: XmlSaxTag;
    rootStart?: number;
  };
  /** `binaryData`: its base64 text so far. */
  base64?: string;
}

/**
 * Reads a Flat OPC document's parts off SAX events. A part's XML is taken from the input as it is —
 * its comments, processing instructions, CDATA sections and references with it — and its root is given
 * the namespace declarations it used from the elements around it, so the part stands on its own.
 */
class FlatOpcReader {
  readonly parts = new Map<string, Uint8Array>();
  readonly contentTypes = new Map<string, string>();
  /** The namespaces in force at each open element: prefix (`""` the default) → URI. */
  private readonly scopes: Array<ReadonlyMap<string, string>> = [new Map()];
  private part: OpenPart | undefined;
  /** Whether the document's DOCTYPE declares entities or attribute lists. */
  private doctype = false;

  constructor(private readonly bytes: Uint8Array) {}

  open(parser: SaxParser, tag: XmlSaxTag): void {
    const outer = this.scopes[this.scopes.length - 1]!;
    let scope = outer;
    if (tag.ns !== undefined) {
      const inner = new Map(outer);
      for (const prefix of Object.keys(tag.ns)) {
        inner.set(prefix, tag.ns[prefix]!);
      }
      scope = inner;
    }
    const depth = parser.depth;
    if (!tag.isSelfClosing) {
      this.scopes.push(scope);
    }
    const part = this.part;
    if (part?.xml !== undefined && part.data === undefined) {
      // Inside a part's XML: only its root is looked at, to declare on it what it inherits.
      if (depth === 3) {
        if (part.xml.root !== undefined) {
          throw new DocxParseError(
            `Flat OPC: the xmlData of part "${part.path}" has more than one root element`
          );
        }
        part.xml.root = tag;
        part.xml.rootStart = parser.startByte!;
      }
      return;
    }
    if (part?.base64 !== undefined && part.data === undefined) {
      throw new DocxParseError(
        `Flat OPC: the binaryData of part "${part.path}" contains an element <${tag.name}>`
      );
    }
    if (depth === 0) {
      if (!isPackageElement(tag, "package")) {
        throw new DocxParseError(
          `Flat OPC: expected root element <pkg:package> of ${NS_PKG}, got <${tag.name}>`
        );
      }
      this.doctype = scanPrologue(this.bytes, 0, parser.startByte!)?.doctype ?? false;
    } else if (depth === 1 && isPackageElement(tag, "part")) {
      // Get part name (with leading /)
      const name = packageAttribute(tag, scope, "name");
      this.part = {
        // Normalize: remove leading slash for map key
        path: name.startsWith("/") ? name.substring(1) : name,
        contentType: packageAttribute(tag, scope, "contentType")
      };
    } else if (depth === 2 && part !== undefined) {
      const xmlData = isPackageElement(tag, "xmlData");
      if (!xmlData && !isPackageElement(tag, "binaryData")) {
        return;
      }
      if (part.xml !== undefined || part.base64 !== undefined) {
        throw new DocxParseError(
          `Flat OPC: part "${part.path}" has more than one xmlData or binaryData element`
        );
      }
      if (xmlData) {
        if (this.doctype) {
          // Its defaulted attributes and entities would be lost: the part is copied, not re-serialised.
          throw new DocxParseError(
            `Flat OPC: part "${part.path}" has xmlData in a document whose DOCTYPE declares entities or attribute lists, which cannot be carried into the part`
          );
        }
        part.xml = { from: parser.endByte!, scope };
        if (tag.isSelfClosing) {
          this.endXmlData(part, parser.endByte!);
        }
      } else {
        part.base64 = "";
        if (tag.isSelfClosing) {
          this.endBinaryData(part);
        }
      }
    }
  }

  close(parser: SaxParser, tag: XmlSaxTag): void {
    if (!tag.isSelfClosing) {
      this.scopes.pop();
    }
    const depth = parser.depth;
    const part = this.part;
    if (part === undefined || tag.isSelfClosing) {
      return;
    }
    if (depth === 2 && part.xml !== undefined && part.data === undefined) {
      this.endXmlData(part, parser.startByte!);
    } else if (depth === 2 && part.base64 !== undefined && part.data === undefined) {
      this.endBinaryData(part);
    } else if (depth === 1) {
      if (part.path !== "" && part.data !== undefined) {
        this.parts.set(part.path, part.data);
      }
      // Collect content type info for [Content_Types].xml reconstruction
      if (part.path !== "" && part.contentType && part.path !== "[Content_Types].xml") {
        this.contentTypes.set(part.path, part.contentType);
      }
      this.part = undefined;
    }
  }

  text(parser: SaxParser, text: string, cdata: boolean): void {
    const part = this.part;
    if (part === undefined || part.data !== undefined) {
      return;
    }
    if (part.base64 !== undefined) {
      part.base64 += text;
    } else if (part.xml !== undefined && parser.depth === 3 && (cdata || /[^ \t\r\n]/.test(text))) {
      throw new DocxParseError(
        `Flat OPC: the xmlData of part "${part.path}" has character data outside its root element`
      );
    }
  }

  /** The part's XML: its content as the input holds it, from `from` to `to`, made to stand alone. */
  private endXmlData(part: OpenPart, to: number): void {
    const xml = part.xml!;
    const content = this.bytes.subarray(xml.from, to);
    const declaration = utf8Encoder.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    );
    const root = xml.root;
    if (root === undefined) {
      throw new DocxParseError(`Flat OPC: the xmlData of part "${part.path}" has no root element`);
    }
    let declarations = "";
    const own = root.ns ?? {};
    const text = utf8Decoder.decode(content);
    for (const [prefix, uri] of xml.scope) {
      // A binding the root makes itself, or one nothing in the part could use, is not repeated.
      if (Object.hasOwn(own, prefix) || prefix === "xml") {
        continue;
      }
      if (prefix === "" ? uri !== "" : text.includes(`${prefix}:`)) {
        declarations += ` ${prefix === "" ? "xmlns" : `xmlns:${prefix}`}="${xmlEncodeAttr(uri)}"`;
      }
    }
    if (declarations === "") {
      part.data = concatBytes(declaration, trimLeadingSpace(content));
      return;
    }
    // After the root's name: `<name` is its start tag's first bytes.
    const at = utf8Encoder.encode(root.name).length + 1 + (xml.rootStart! - xml.from);
    part.data = concatBytes(
      declaration,
      trimLeadingSpace(content.subarray(0, at)),
      utf8Encoder.encode(declarations),
      content.subarray(at)
    );
  }

  private endBinaryData(part: OpenPart): void {
    part.data = base64Decode(part.base64!.replace(/\s+/g, ""), part.path);
  }
}

/** Whether `tag` is the package element `local`: in the package namespace, or — as some writers leave it — in none. */
function isPackageElement(tag: XmlSaxTag, local: string): boolean {
  return tag.local === local && (tag.uri === NS_PKG || tag.uri === "");
}

/** A package attribute of `tag`: in the package namespace, or unprefixed. */
function packageAttribute(
  tag: XmlSaxTag,
  scope: ReadonlyMap<string, string>,
  local: string
): string {
  for (const name of Object.keys(tag.attributes)) {
    const colon = name.indexOf(":");
    if (
      colon < 0
        ? name === local
        : name.slice(colon + 1) === local && scope.get(name.slice(0, colon)) === NS_PKG
    ) {
      return tag.attributes[name]!;
    }
  }
  return "";
}

function trimLeadingSpace(bytes: Uint8Array): Uint8Array {
  let i = 0;
  while (i < bytes.length && isSpaceByte(bytes[i]!)) {
    i++;
  }
  return bytes.subarray(i);
}

function concatBytes(...parts: Uint8Array[]): Uint8Array {
  let length = 0;
  for (const part of parts) {
    length += part.length;
  }
  const out = new Uint8Array(length);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

/**
 * Check if content is a Flat OPC XML document: its root element is `package` in the package namespace.
 * A ZIP is never one, whatever text it stores; the prologue (declaration, comments, processing
 * instructions, DOCTYPE) is skipped without parsing the rest of the document.
 */
export function isFlatOpc(content: string | Uint8Array | ArrayBuffer | ArrayBufferView): boolean {
  const given = checkTextOrBytes(content, "content", "Convert.isFlatOpc");
  const bytes = typeof given === "string" ? utf8Encoder.encode(given) : given;
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    return false;
  }
  let at = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  const prologue = scanPrologue(bytes, at, bytes.length);
  if (prologue === undefined || prologue.rootStart === undefined) {
    return false;
  }
  at = prologue.rootStart;
  const end = findTagEnd(bytes, at, Math.min(bytes.length, at + MAX_ROOT_TAG_BYTES));
  if (end < 0) {
    return false;
  }
  const tag = utf8Decoder.decode(bytes.subarray(at + 1, end));
  const name = /^[^\s/>]+/.exec(tag)?.[0] ?? "";
  const colon = name.indexOf(":");
  const prefix = colon < 0 ? "" : name.slice(0, colon);
  if (name.slice(colon + 1) !== "package") {
    return false;
  }
  const attr = prefix === "" ? "xmlns" : `xmlns:${prefix}`;
  for (const match of tag.matchAll(/([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/g)) {
    if (match[1] === attr) {
      return match[2]!.slice(1, -1) === NS_PKG;
    }
  }
  return false;
}

/** How far a root start tag is read to find its namespace declarations. */
const MAX_ROOT_TAG_BYTES = 64 * 1024;

function startsWithAscii(bytes: Uint8Array, at: number, text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    if (bytes[at + i] !== text.charCodeAt(i)) {
      return false;
    }
  }
  return true;
}

function indexOfAscii(bytes: Uint8Array, text: string, from: number, to: number): number {
  for (let i = from; i + text.length <= to; i++) {
    if (startsWithAscii(bytes, i, text)) {
      return i;
    }
  }
  return -1;
}

/** The index of the `>` closing the tag that starts at `at`, outside quoted values; -1 before `to`. */
function findTagEnd(bytes: Uint8Array, at: number, to: number): number {
  let quote = 0;
  for (let i = at; i < to; i++) {
    const c = bytes[i]!;
    if (quote !== 0) {
      if (c === quote) {
        quote = 0;
      }
    } else if (c === 0x22 || c === 0x27) {
      quote = c;
    } else if (c === 0x3e) {
      return i;
    }
  }
  return -1;
}

/**
 * Walk the prologue in `[from, to)`: white space, the declaration, comments, processing
 * instructions and a DOCTYPE (internal subset included). Returns where the first start tag
 * begins (`undefined` if `to` comes first) and whether a DOCTYPE was passed, or `undefined`
 * if something else stands there.
 */
function scanPrologue(
  bytes: Uint8Array,
  from: number,
  to: number
): { rootStart: number | undefined; doctype: boolean } | undefined {
  // `doctype`: whether a DOCTYPE declares entities or attribute lists — what a part cut from the
  // document would lose.
  let at = from;
  let doctype = false;
  while (at < to) {
    const c = bytes[at]!;
    if (isSpaceByte(c)) {
      at++;
      continue;
    }
    if (c !== 0x3c) {
      return undefined;
    }
    let end: number;
    if (startsWithAscii(bytes, at, "<?")) {
      end = indexOfAscii(bytes, "?>", at + 2, to);
      end = end < 0 ? -1 : end + 2;
    } else if (startsWithAscii(bytes, at, "<!--")) {
      end = indexOfAscii(bytes, "-->", at + 4, to);
      end = end < 0 ? -1 : end + 3;
    } else if (startsWithAscii(bytes, at, "<!DOCTYPE")) {
      const skipped = skipDoctype(bytes, at + 9, to);
      doctype ||= skipped.declares;
      end = skipped.end;
    } else if (at + 1 < to && bytes[at + 1] !== 0x21 && bytes[at + 1] !== 0x3f) {
      return { rootStart: at, doctype };
    } else {
      return undefined;
    }
    if (end < 0) {
      return { rootStart: undefined, doctype };
    }
    at = end;
  }
  return { rootStart: undefined, doctype };
}

/**
 * The index after a DOCTYPE whose name starts at `at` — past its `>`, skipping quotes, comments,
 * processing instructions and its internal subset, or `-1` before `to` — and whether that subset
 * declares an entity or an attribute list: the declarations that change what an element reads as.
 */
function skipDoctype(
  bytes: Uint8Array,
  at: number,
  to: number
): { end: number; declares: boolean } {
  let quote = 0;
  let subset = false;
  let declares = false;
  for (let i = at; i < to; i++) {
    const c = bytes[i]!;
    if (quote !== 0) {
      if (c === quote) {
        quote = 0;
      }
    } else if (c === 0x22 || c === 0x27) {
      quote = c;
    } else if (subset && (startsWithAscii(bytes, i, "<!--") || startsWithAscii(bytes, i, "<?"))) {
      const comment = bytes[i + 1] === 0x21;
      const end = indexOfAscii(bytes, comment ? "-->" : "?>", i + (comment ? 4 : 2), to);
      if (end < 0) {
        return { end: -1, declares };
      }
      i = end + (comment ? 2 : 1);
    } else if (
      subset &&
      (startsWithAscii(bytes, i, "<!ENTITY") || startsWithAscii(bytes, i, "<!ATTLIST"))
    ) {
      declares = true;
    } else if (c === 0x5b) {
      subset = true;
    } else if (c === 0x5d) {
      subset = false;
    } else if (c === 0x3e && !subset) {
      return { end: i + 1, declares };
    }
  }
  return { end: -1, declares };
}

// =============================================================================
// Writing Flat OPC
// =============================================================================

/**
 * Convert a ZIP archive's entries (as a map of path → data) to Flat OPC XML string.
 *
 * @param entries - Map of part paths to their binary content.
 * @param contentTypes - Map of part paths to content types (from [Content_Types].xml).
 * @returns The Flat OPC XML string.
 */
export function toFlatOpc(
  entries: ReadonlyMap<string, Uint8Array>,
  contentTypes?: ReadonlyMap<string, string>
): string {
  if (!(entries instanceof Map)) {
    invalidArgument("Convert.partsToFlatOpc", "entries", "a Map of part name to bytes", entries);
  }
  if (contentTypes !== undefined && !(contentTypes instanceof Map)) {
    invalidArgument("Convert.partsToFlatOpc", "contentTypes", "a Map", contentTypes);
  }
  const resolvedCT = contentTypes ?? extractContentTypes(entries);
  const xml = new XmlWriter();

  xml.openXml({ version: "1.0", encoding: "UTF-8", standalone: "yes" });
  xml.openNode("pkg:package", {
    "xmlns:pkg": NS_PKG
  });

  // Write [Content_Types].xml first if present
  const ctEntry = entries.get("[Content_Types].xml");
  if (ctEntry) {
    writePartElement(
      xml,
      "/[Content_Types].xml",
      "application/vnd.openxmlformats-package.content-types+xml",
      ctEntry,
      true
    );
  }

  // Write all other parts
  for (const [path, data] of entries) {
    if (path === "[Content_Types].xml") {
      continue;
    }

    const partName = path.startsWith("/") ? path : `/${path}`;
    const ct = resolvedCT.get(path) ?? inferContentTypeFromPath(path);
    const isXml = isXmlContentType(ct) || path.endsWith(".xml") || path.endsWith(".rels");

    writePartElement(xml, partName, ct, data, isXml);
  }

  xml.closeNode(); // </pkg:package>
  return xml.xml;
}

// toFlatOpcFromDoc is exported from document-io.ts (uses dynamic import of this module)

// =============================================================================
// Internal Helpers
// =============================================================================

function writePartElement(
  xml: XmlWriter,
  partName: string,
  contentType: string,
  data: Uint8Array,
  isXml: boolean
): void {
  xml.openNode("pkg:part", {
    "pkg:name": partName,
    "pkg:contentType": contentType
  });

  if (isXml) {
    xml.openNode("pkg:xmlData", {});
    // Write raw XML content (already has XML declaration usually)
    const xmlStr = utf8Decoder.decode(data);
    // Strip XML declaration if present (Flat OPC embeds it inline)
    const stripped = xmlStr.replace(/<\?xml[^?]*\?>\s*/, "");
    xml.writeRaw(stripped);
    xml.closeNode();
  } else {
    xml.openNode("pkg:binaryData", {});
    xml.writeText(base64Encode(data));
    xml.closeNode();
  }

  xml.closeNode(); // </pkg:part>
}

function base64Decode(b64: string, path: string): Uint8Array {
  // Checked here, so the native and the manual decoder refuse the same text, as this module's error.
  if (b64.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(b64)) {
    throw new DocxParseError(`Flat OPC: the binaryData of part "${path}" is not base64`);
  }
  // Use atob in browser, Buffer in Node
  if (typeof globalThis.atob === "function") {
    const binary = globalThis.atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }
  // Fallback manual decode
  return manualBase64Decode(b64);
}

function base64Encode(data: Uint8Array): string {
  if (typeof globalThis.btoa === "function") {
    let binary = "";
    for (let i = 0; i < data.length; i++) {
      binary += String.fromCharCode(data[i]!);
    }
    return globalThis.btoa(binary);
  }
  return manualBase64Encode(data);
}

const B64_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

function manualBase64Encode(data: Uint8Array): string {
  let result = "";
  const len = data.length;
  for (let i = 0; i < len; i += 3) {
    const b0 = data[i]!;
    const b1 = i + 1 < len ? data[i + 1]! : 0;
    const b2 = i + 2 < len ? data[i + 2]! : 0;
    result += B64_CHARS[(b0 >> 2) & 0x3f];
    result += B64_CHARS[((b0 << 4) | (b1 >> 4)) & 0x3f];
    result += i + 1 < len ? B64_CHARS[((b1 << 2) | (b2 >> 6)) & 0x3f]! : "=";
    result += i + 2 < len ? B64_CHARS[b2 & 0x3f]! : "=";
  }
  return result;
}

function manualBase64Decode(b64: string): Uint8Array {
  const lookup = new Uint8Array(128);
  for (let i = 0; i < B64_CHARS.length; i++) {
    lookup[B64_CHARS.charCodeAt(i)] = i;
  }
  const len = b64.length;
  const padLen = b64.endsWith("==") ? 2 : b64.endsWith("=") ? 1 : 0;
  const byteLen = (len * 3) / 4 - padLen;
  const result = new Uint8Array(byteLen);
  let j = 0;
  for (let i = 0; i < len; i += 4) {
    const a = lookup[b64.charCodeAt(i)]!;
    const b = lookup[b64.charCodeAt(i + 1)]!;
    const c = lookup[b64.charCodeAt(i + 2)]!;
    const d = lookup[b64.charCodeAt(i + 3)]!;
    result[j++] = (a << 2) | (b >> 4);
    if (j < byteLen) {
      result[j++] = ((b << 4) | (c >> 2)) & 0xff;
    }
    if (j < byteLen) {
      result[j++] = ((c << 6) | d) & 0xff;
    }
  }
  return result;
}

function isXmlContentType(ct: string): boolean {
  return ct.endsWith("+xml") || ct === "application/xml" || ct === "text/xml";
}

function inferContentTypeFromPath(path: string): string {
  const ext = getFileExt(path);
  return INFERRED_CONTENT_TYPES.get(ext) ?? "application/octet-stream";
}

const INFERRED_CONTENT_TYPES: ReadonlyMap<string, string> = new Map([
  ["xml", "application/xml"],
  ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
  ["png", "image/png"],
  ["jpeg", "image/jpeg"],
  ["jpg", "image/jpeg"],
  ["gif", "image/gif"],
  ["svg", "image/svg+xml"],
  ["emf", "image/x-emf"],
  ["wmf", "image/x-wmf"],
  ["bin", "application/vnd.openxmlformats-officedocument.oleObject"],
  ["odttf", "application/vnd.openxmlformats-officedocument.obfuscatedFont"]
]);

/** Extract content types from [Content_Types].xml if present in entries. */
function extractContentTypes(entries: ReadonlyMap<string, Uint8Array>): Map<string, string> {
  const ctMap = new Map<string, string>();
  const ctXml = entries.get("[Content_Types].xml");
  if (!ctXml) {
    return ctMap;
  }

  const defaults = new Map<string, string>();
  // From the first `<`, as a part is read; elements are matched by namespace and local name.
  let lead = ctXml[0] === 0xef && ctXml[1] === 0xbb && ctXml[2] === 0xbf ? 3 : 0;
  while (lead < ctXml.length && isSpaceByte(ctXml[lead]!)) {
    lead++;
  }
  const parser = new SaxParser({ position: false, xmlns: true, byteOffset: lead });
  parser.on("opentag", tag => {
    if (parser.depth === 0) {
      if (tag.uri !== NS_CONTENT_TYPES || tag.local !== "Types") {
        parser.fail(`expected the root <Types> of ${NS_CONTENT_TYPES}, got <${tag.name}>`);
      }
      return;
    }
    if (parser.depth !== 1 || tag.uri !== NS_CONTENT_TYPES) {
      return;
    }
    const ct = tag.attributes.ContentType ?? "";
    if (tag.local === "Default") {
      const ext = tag.attributes.Extension ?? "";
      if (ext && ct) {
        defaults.set(ext.toLowerCase(), ct);
      }
    } else if (tag.local === "Override") {
      const partName = tag.attributes.PartName ?? "";
      if (partName && ct) {
        ctMap.set(partName.startsWith("/") ? partName.substring(1) : partName, ct);
      }
    }
  });
  try {
    parser.write(ctXml.subarray(lead)).close();
  } catch (error) {
    if (error instanceof XmlParseError) {
      throw new DocxParseError(`Flat OPC: [Content_Types].xml cannot be read: ${error.message}`, {
        cause: error
      });
    }
    throw error;
  }

  // Apply defaults to entries without explicit overrides
  for (const path of entries.keys()) {
    if (path === "[Content_Types].xml") {
      continue;
    }
    if (ctMap.has(path)) {
      continue;
    }
    const ext = getFileExt(path);
    const ct = defaults.get(ext);
    if (ct) {
      ctMap.set(path, ct);
    }
  }

  return ctMap;
}

/** Build [Content_Types].xml from collected content type info. */
function buildContentTypesXml(ctMap: Map<string, string>): Uint8Array {
  // Collect extension defaults
  const defaults = new Map<string, string>();
  defaults.set("rels", "application/vnd.openxmlformats-package.relationships+xml");
  defaults.set("xml", "application/xml");

  for (const [path, ct] of ctMap) {
    const ext = getFileExt(path);
    if (isImageExtension(ext) && !defaults.has(ext)) {
      defaults.set(ext, ct);
    }
  }

  const xml = new XmlWriter();
  xml.openXml({ version: "1.0", encoding: "UTF-8", standalone: "yes" });
  xml.openNode("Types", {
    xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
  });

  // Write Defaults first (sorted)
  const sortedDefaults = [...defaults.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  for (const [ext, ct] of sortedDefaults) {
    xml.leafNode("Default", { Extension: ext, ContentType: ct });
  }

  // Write Overrides for parts that don't match their extension default
  for (const [path, ct] of ctMap) {
    const ext = getFileExt(path);
    const defaultCt = defaults.get(ext);
    if (!defaultCt || defaultCt !== ct) {
      xml.leafNode("Override", { PartName: `/${path}`, ContentType: ct });
    }
  }

  xml.closeNode();
  return utf8Encoder.encode(xml.xml);
}

function isImageExtension(ext: string): boolean {
  return ["png", "jpeg", "jpg", "gif", "bmp", "tiff", "tif", "svg", "webp", "emf", "wmf"].includes(
    ext
  );
}

function checkTextOrBytes(value: unknown, field: string, where: string): string | Uint8Array {
  return typeof value === "string" ? value : checkBytes(value, field, where);
}
