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

import { parseXml } from "@xml/dom";
import type { XmlElement, XmlNode } from "@xml/types";
import { XmlWriter } from "@xml/writer";

import { utf8Decoder, utf8Encoder } from "../core/internal-utils";
import { getFileExt } from "../core/opc-package";
import { DocxParseError } from "../errors";

// =============================================================================
// Constants
// =============================================================================

/** Flat OPC package namespace. */
const NS_PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";

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
export function parseFlatOpc(xmlContent: string | Uint8Array): Map<string, Uint8Array> {
  const xmlStr = typeof xmlContent === "string" ? xmlContent : utf8Decoder.decode(xmlContent);

  const doc = parseXml(xmlStr);
  const root = doc.root;
  if (!root) {
    throw new DocxParseError("Flat OPC: missing root element");
  }

  // Root should be <pkg:package> or <package>
  const rootName = root.name.replace(/^pkg:/, "");
  if (rootName !== "package") {
    throw new DocxParseError(`Flat OPC: expected root element <pkg:package>, got <${root.name}>`);
  }

  const parts = new Map<string, Uint8Array>();
  const collectedContentTypes = new Map<string, string>();

  for (const child of root.children) {
    if (typeof child === "string") {
      continue;
    }
    const el = child as XmlElement;
    const elName = el.name.replace(/^pkg:/, "");
    if (elName !== "part") {
      continue;
    }

    // Get part name (with leading /)
    const partName = el.attributes?.["pkg:name"] ?? el.attributes?.["name"] ?? "";
    if (!partName) {
      continue;
    }

    // Normalize: remove leading slash for map key
    const normalizedPath = partName.startsWith("/") ? partName.substring(1) : partName;

    // Get content type
    const contentType = el.attributes?.["pkg:contentType"] ?? el.attributes?.["contentType"] ?? "";

    // Find the data child: <pkg:xmlData> or <pkg:binaryData>
    let data: Uint8Array | undefined;

    for (const partChild of el.children) {
      if (typeof partChild === "string") {
        continue;
      }
      const pc = partChild as XmlElement;
      const pcName = pc.name.replace(/^pkg:/, "");

      if (pcName === "xmlData") {
        // XML content — serialize the child elements back to XML
        const xmlWriter = new XmlWriter();
        xmlWriter.openXml({ version: "1.0", encoding: "UTF-8", standalone: "yes" });
        serializeXmlChildren(xmlWriter, pc);
        data = utf8Encoder.encode(xmlWriter.xml);
        break;
      } else if (pcName === "binaryData") {
        // Base64 encoded binary content
        const b64Text = getTextContent(pc).replace(/\s+/g, "");
        data = base64Decode(b64Text);
        break;
      }
    }

    if (data) {
      parts.set(normalizedPath, data);
    }
    // Collect content type info for [Content_Types].xml reconstruction
    if (contentType && normalizedPath !== "[Content_Types].xml") {
      collectedContentTypes.set(normalizedPath, contentType);
    }
  }

  // Ensure [Content_Types].xml exists
  if (!parts.has("[Content_Types].xml")) {
    parts.set("[Content_Types].xml", buildContentTypesXml(collectedContentTypes));
  }

  return parts;
}

/**
 * Check if content appears to be a Flat OPC XML document.
 */
export function isFlatOpc(content: string | Uint8Array): boolean {
  const sample =
    typeof content === "string"
      ? content.substring(0, 500)
      : utf8Decoder.decode(content.slice(0, 500));
  // Match the actual opening tag pattern, not just the string presence
  return (
    sample.includes("<pkg:package") ||
    sample.includes('xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"')
  );
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
  entries: Map<string, Uint8Array>,
  contentTypes?: Map<string, string>
): string {
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

function serializeXmlChildren(writer: XmlWriter, el: XmlElement): void {
  for (const child of el.children) {
    const node = child as XmlNode;
    switch (node.type) {
      case "text":
        writer.writeText(node.value);
        break;
      case "cdata":
        writer.writeRaw(`<![CDATA[${node.value}]]>`);
        break;
      case "comment":
        writer.writeRaw(`<!--${node.value}-->`);
        break;
      case "processing-instruction":
        writer.writeRaw(`<?${node.target} ${node.body}?>`);
        break;
      case "element": {
        const childEl = node as XmlElement;
        if (childEl.children && childEl.children.length > 0) {
          writer.openNode(childEl.name, childEl.attributes ?? {});
          serializeXmlChildren(writer, childEl);
          writer.closeNode();
        } else {
          writer.leafNode(childEl.name, childEl.attributes ?? {});
        }
        break;
      }
    }
  }
}

function getTextContent(el: XmlElement): string {
  let result = "";
  for (const child of el.children) {
    const node = child as XmlNode;
    if (node.type === "text" || node.type === "cdata") {
      result += node.value;
    } else if (node.type === "element") {
      result += getTextContent(node as XmlElement);
    }
  }
  return result;
}

function base64Decode(b64: string): Uint8Array {
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
  const map: Record<string, string> = {
    xml: "application/xml",
    rels: "application/vnd.openxmlformats-package.relationships+xml",
    png: "image/png",
    jpeg: "image/jpeg",
    jpg: "image/jpeg",
    gif: "image/gif",
    svg: "image/svg+xml",
    emf: "image/x-emf",
    wmf: "image/x-wmf",
    bin: "application/vnd.openxmlformats-officedocument.oleObject",
    odttf: "application/vnd.openxmlformats-officedocument.obfuscatedFont"
  };
  return map[ext] ?? "application/octet-stream";
}

/** Extract content types from [Content_Types].xml if present in entries. */
function extractContentTypes(entries: Map<string, Uint8Array>): Map<string, string> {
  const ctMap = new Map<string, string>();
  const ctXml = entries.get("[Content_Types].xml");
  if (!ctXml) {
    return ctMap;
  }

  const xmlStr = utf8Decoder.decode(ctXml);
  const doc = parseXml(xmlStr);
  const root = doc.root;
  if (!root) {
    return ctMap;
  }

  const defaults = new Map<string, string>();

  for (const child of root.children) {
    if (typeof child === "string") {
      continue;
    }
    const el = child as XmlElement;
    if (el.name === "Default") {
      const ext = el.attributes?.Extension ?? "";
      const ct = el.attributes?.ContentType ?? "";
      if (ext && ct) {
        defaults.set(ext.toLowerCase(), ct);
      }
    } else if (el.name === "Override") {
      const partName = el.attributes?.PartName ?? "";
      const ct = el.attributes?.ContentType ?? "";
      if (partName && ct) {
        const normalized = partName.startsWith("/") ? partName.substring(1) : partName;
        ctMap.set(normalized, ct);
      }
    }
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
