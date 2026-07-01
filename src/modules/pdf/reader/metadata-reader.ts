/**
 * PDF metadata reader.
 *
 * Extracts document metadata from:
 * 1. Info Dictionary (traditional metadata)
 *    - Title, Author, Subject, Keywords, Creator, Producer
 *    - CreationDate, ModDate
 *
 * 2. XMP Metadata Stream (XML-based, more comprehensive)
 *    - All of the above plus:
 *    - Dublin Core metadata, custom properties
 *
 * @see PDF Reference 1.7, §10.2 - Metadata
 * @see XMP Specification Part 1
 */

import type { PdfDocument } from "@pdf/reader/pdf-document";
import type { PdfDictValue } from "@pdf/reader/pdf-parser";
import { dictGetRef, decodePdfStringBytes } from "@pdf/reader/pdf-parser";

// =============================================================================
// Types
// =============================================================================

/**
 * PDF document metadata.
 */
export interface PdfMetadata {
  /** Document title */
  title: string;
  /** Document author */
  author: string;
  /** Document subject */
  subject: string;
  /** Document keywords */
  keywords: string;
  /** Application that created the original document */
  creator: string;
  /** Application that produced the PDF */
  producer: string;
  /** Date the document was created */
  creationDate: Date | null;
  /** Date the document was last modified */
  modDate: Date | null;
  /** PDF version */
  pdfVersion: string;
  /** Number of pages */
  pageCount: number;
  /** Whether the document is encrypted */
  encrypted: boolean;
  /** Page size of the first page (in points) */
  pageSize: { width: number; height: number } | null;
  /** Raw XMP metadata XML (if available) */
  xmpXml: string | null;
  /** Additional custom metadata from Info dictionary */
  custom: Record<string, string>;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Extract metadata from a PDF document.
 */
export function extractMetadata(doc: PdfDocument): PdfMetadata {
  const metadata: PdfMetadata = {
    title: "",
    author: "",
    subject: "",
    keywords: "",
    creator: "",
    producer: "",
    creationDate: null,
    modDate: null,
    pdfVersion: extractPdfVersion(doc),
    pageCount: 0,
    encrypted: doc.trailer.has("Encrypt"),
    pageSize: null,
    xmpXml: null,
    custom: {}
  };

  // Extract from Info dictionary
  extractInfoDict(doc, metadata);

  // Extract from XMP metadata stream
  extractXmpMetadata(doc, metadata);

  // Get page count and first page size
  try {
    const pages = doc.getPages();
    metadata.pageCount = pages.length;
    if (pages.length > 0) {
      metadata.pageSize = extractPageSize(pages[0], doc);
    }
  } catch {
    // Ignore page tree errors
  }

  return metadata;
}

// =============================================================================
// PDF Version
// =============================================================================

function extractPdfVersion(doc: PdfDocument): string {
  const data = doc.data;
  // First line: %PDF-X.Y
  if (
    data[0] === 0x25 &&
    data[1] === 0x50 &&
    data[2] === 0x44 &&
    data[3] === 0x46 &&
    data[4] === 0x2d
  ) {
    let version = "";
    for (let i = 5; i < Math.min(data.length, 15); i++) {
      const b = data[i];
      if (b === 0x0a || b === 0x0d || b === 0x20) {
        break;
      }
      version += String.fromCharCode(b);
    }
    return version;
  }

  // Check catalog /Version
  try {
    const catalog = doc.getCatalog();
    const version = catalog.get("Version");
    if (typeof version === "string") {
      return version;
    }
  } catch {
    // Ignore
  }

  return "1.0";
}

// =============================================================================
// Info Dictionary
// =============================================================================

function extractInfoDict(doc: PdfDocument, metadata: PdfMetadata): void {
  const infoRef = dictGetRef(doc.trailer, "Info");
  if (!infoRef) {
    return;
  }

  const infoDict = doc.derefDict(infoRef);
  if (!infoDict) {
    return;
  }

  const knownKeys = new Set([
    "Title",
    "Author",
    "Subject",
    "Keywords",
    "Creator",
    "Producer",
    "CreationDate",
    "ModDate"
  ]);

  for (const [key, value] of infoDict) {
    const strValue =
      value instanceof Uint8Array ? decodePdfStringBytes(value) : String(value ?? "");

    switch (key) {
      case "Title":
        metadata.title = strValue;
        break;
      case "Author":
        metadata.author = strValue;
        break;
      case "Subject":
        metadata.subject = strValue;
        break;
      case "Keywords":
        metadata.keywords = strValue;
        break;
      case "Creator":
        metadata.creator = strValue;
        break;
      case "Producer":
        metadata.producer = strValue;
        break;
      case "CreationDate":
        metadata.creationDate = parsePdfDate(strValue);
        break;
      case "ModDate":
        metadata.modDate = parsePdfDate(strValue);
        break;
      default:
        if (!knownKeys.has(key)) {
          metadata.custom[key] = strValue;
        }
        break;
    }
  }
}

// =============================================================================
// XMP Metadata
// =============================================================================

function extractXmpMetadata(doc: PdfDocument, metadata: PdfMetadata): void {
  try {
    const catalog = doc.getCatalog();
    const metadataRef = catalog.get("Metadata");
    if (!metadataRef) {
      return;
    }

    const result = doc.derefStreamWithObjNum(metadataRef);
    if (!result) {
      return;
    }

    const data = doc.getStreamData(result.stream, result.objNum, result.gen);
    const xml = new TextDecoder("utf-8").decode(data);
    metadata.xmpXml = xml;

    // Parse key fields from XMP
    if (!metadata.title) {
      metadata.title = extractXmpField(xml, "dc:title") ?? "";
    }
    if (!metadata.author) {
      metadata.author = extractXmpField(xml, "dc:creator") ?? "";
    }
    if (!metadata.subject) {
      metadata.subject = extractXmpField(xml, "dc:description") ?? "";
    }
    if (!metadata.keywords) {
      metadata.keywords = extractXmpField(xml, "pdf:Keywords") ?? "";
    }
    if (!metadata.creator) {
      metadata.creator = extractXmpField(xml, "xmp:CreatorTool") ?? "";
    }
    if (!metadata.producer) {
      metadata.producer = extractXmpField(xml, "pdf:Producer") ?? "";
    }
    if (!metadata.creationDate) {
      const dateStr = extractXmpField(xml, "xmp:CreateDate");
      if (dateStr) {
        metadata.creationDate = new Date(dateStr);
      }
    }
    if (!metadata.modDate) {
      const dateStr = extractXmpField(xml, "xmp:ModifyDate");
      if (dateStr) {
        metadata.modDate = new Date(dateStr);
      }
    }
  } catch {
    // Ignore XMP errors
  }
}

/**
 * Extract a field value from XMP XML using simple regex.
 * Handles both simple elements and rdf:Alt/rdf:Bag/rdf:Seq containers.
 */
function extractXmpField(xml: string, field: string): string | null {
  // Try simple element: <field>value</field>
  const simpleRegex = new RegExp(`<${field}[^>]*>([^<]+)</${field}>`, "i");
  const simpleMatch = simpleRegex.exec(xml);
  if (simpleMatch) {
    return decodeXmlEntities(simpleMatch[1].trim());
  }

  // Try rdf:Alt/rdf:Bag/rdf:Seq container: <field>...<rdf:li>value</rdf:li>...</field>
  const containerRegex = new RegExp(`<${field}[^>]*>.*?<rdf:li[^>]*>([^<]+)</rdf:li>`, "is");
  const containerMatch = containerRegex.exec(xml);
  if (containerMatch) {
    return decodeXmlEntities(containerMatch[1].trim());
  }

  return null;
}

/** @internal Exported for testing only. */
export function decodeXmlEntities(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

// =============================================================================
// Page Size
// =============================================================================

function extractPageSize(
  pageDict: PdfDictValue,
  doc: PdfDocument
): { width: number; height: number } | null {
  return doc.resolvePageBox(pageDict);
}

// =============================================================================
// PDF Date Parsing
// =============================================================================

/**
 * Parse a PDF date string to a Date object.
 * Format: D:YYYYMMDDHHmmSSOHH'mm
 */
function parsePdfDate(dateStr: string): Date | null {
  if (!dateStr) {
    return null;
  }

  // Remove leading "D:" if present
  let s = dateStr;
  if (s.startsWith("D:")) {
    s = s.substring(2);
  }

  // Parse components
  const year = parseInt(s.substring(0, 4), 10);
  if (isNaN(year)) {
    return null;
  }

  const month = parseInt(s.substring(4, 6), 10) || 1;
  const day = parseInt(s.substring(6, 8), 10) || 1;
  const hour = parseInt(s.substring(8, 10), 10) || 0;
  const minute = parseInt(s.substring(10, 12), 10) || 0;
  const second = parseInt(s.substring(12, 14), 10) || 0;

  // Parse timezone
  const tzChar = s.charAt(14);
  let offsetMinutes = 0;

  if (tzChar === "+" || tzChar === "-") {
    const tzHour = parseInt(s.substring(15, 17), 10) || 0;
    const tzMin = parseInt(s.substring(18, 20), 10) || 0;
    offsetMinutes = (tzHour * 60 + tzMin) * (tzChar === "-" ? -1 : 1);
  }

  // Create Date in UTC
  const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));

  // Apply timezone offset
  if (offsetMinutes !== 0 && tzChar !== "Z") {
    date.setUTCMinutes(date.getUTCMinutes() - offsetMinutes);
  }

  return isNaN(date.getTime()) ? null : date;
}
