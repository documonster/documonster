/**
 * DOCX Module - OLE Embedded Objects
 *
 * Provides support for reading and round-tripping OLE (Object Linking and Embedding)
 * objects within DOCX files. OLE objects allow embedding other documents (Excel
 * spreadsheets, PowerPoint presentations, PDFs, etc.) within a Word document.
 *
 * OLE objects in OOXML are stored as:
 * - Binary data in word/embeddings/*.bin (OLE2 compound document)
 * - Referenced via r:id in w:object/o:OLEObject elements
 * - May have a preview image (EMF/WMF) in word/media/
 *
 * This module focuses on preservation (round-trip) and metadata extraction,
 * not full OLE compound document manipulation.
 */

import { getFileName } from "../core/opc-package";
import type { DocxDocument, OpaquePart } from "../types";

// =============================================================================
// Types
// =============================================================================

/** OLE object type classification. */
export type OleObjectType = "embedded" | "linked";

/** OLE object display style. */
export type OleDisplayAs = "icon" | "content";

/** Metadata about an OLE embedded object. */
export interface OleObject {
  /** Unique relationship ID for the OLE binary. */
  readonly rId: string;
  /** OLE ProgId (e.g. "Excel.Sheet.12", "PowerPoint.Slide.12", "Package"). */
  readonly progId: string;
  /** Whether the object is embedded or linked. */
  readonly objectType: OleObjectType;
  /** Display mode. */
  readonly displayAs: OleDisplayAs;
  /** Relationship ID for the preview image (EMF/WMF). */
  readonly imageRId?: string;
  /** Shape ID for the hosting VML shape. */
  readonly shapeId?: string;
  /** Object width in EMU. */
  readonly width?: number;
  /** Object height in EMU. */
  readonly height?: number;
  /** File name of the embedded data within the archive. */
  readonly fileName?: string;
  /** The raw binary data of the OLE object. */
  readonly data?: Uint8Array;
  /** Link target (for linked objects). */
  readonly linkTarget?: string;
  /** Raw XML of the w:object element for round-trip preservation. */
  readonly rawXml?: string;
}

/** Result of extracting OLE objects from a document. */
export interface OleExtractionResult {
  /** All OLE objects found in the document. */
  readonly objects: readonly OleObject[];
  /** Mapping of progId to count. */
  readonly summary: Readonly<Record<string, number>>;
}

// =============================================================================
// OLE Object Extraction
// =============================================================================

/**
 * Extract metadata about OLE embedded objects from a parsed DOCX document.
 * OLE objects are preserved in opaqueParts during reading; this function
 * inspects them to extract meaningful metadata.
 *
 * @param doc - The parsed DOCX document.
 * @returns Extraction result with OLE object metadata.
 */
export function extractOleObjects(doc: DocxDocument): OleExtractionResult {
  const objects: OleObject[] = [];
  const summary: Record<string, number> = {};

  // Scan opaque parts for OLE embeddings
  if (doc.opaqueParts) {
    for (const part of doc.opaqueParts) {
      if (isOleEmbedding(part.path)) {
        const obj = parseOlePartMetadata(part);
        if (obj) {
          objects.push(obj);
          summary[obj.progId] = (summary[obj.progId] ?? 0) + 1;
        }
      }
    }
  }

  // Also scan body for OpaqueDrawing elements that may contain OLE references
  for (const element of doc.body) {
    if (element.type === "opaqueDrawing") {
      const oleFromDrawing = extractOleFromRawXml(element.rawXml);
      if (oleFromDrawing) {
        // Check if we already have this object from opaque parts
        const exists = objects.some(o => o.rId === oleFromDrawing.rId);
        if (!exists) {
          objects.push(oleFromDrawing);
          summary[oleFromDrawing.progId] = (summary[oleFromDrawing.progId] ?? 0) + 1;
        }
      }
    }
  }

  return { objects, summary };
}

/**
 * Check if a document contains any OLE embedded objects.
 */
export function hasOleObjects(doc: DocxDocument): boolean {
  if (doc.opaqueParts) {
    for (const part of doc.opaqueParts) {
      if (isOleEmbedding(part.path)) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Get the binary data of a specific OLE object by its relationship ID.
 * Returns undefined if not found.
 */
export function getOleObjectData(doc: DocxDocument, rId: string): Uint8Array | undefined {
  if (!doc.opaqueParts) {
    return undefined;
  }

  for (const part of doc.opaqueParts) {
    if (isOleEmbedding(part.path)) {
      // Check relationships for matching rId
      if (part.relationships) {
        for (const rel of part.relationships) {
          if (rel.id === rId) {
            return part.data;
          }
        }
      }
      // Also check if the file path matches expected pattern for this rId
      if (part.path.includes(rId.replace("rId", "oleObject"))) {
        return part.data;
      }
    }
  }

  // Unknown rId — do NOT fall back to "first embedded blob" (that historically
  // returned an arbitrary OLE object and silently leaked the wrong data).
  return undefined;
}

/**
 * Create an OLE object entry for embedding into a document.
 * The object will be stored as an opaque part for round-trip preservation.
 *
 * @param data - The binary data of the OLE object.
 * @param progId - The OLE ProgId (e.g. "Excel.Sheet.12").
 * @param options - Additional options.
 * @returns An OpaquePart that can be added to doc.opaqueParts.
 */
export function createOleEmbedding(
  data: Uint8Array,
  progId: string,
  options?: {
    fileName?: string;
    previewImage?: Uint8Array;
    previewContentType?: string;
  }
): OpaquePart {
  const fileName = options?.fileName ?? `oleObject1.bin`;
  return {
    path: `word/embeddings/${fileName}`,
    data,
    contentType: "application/vnd.openxmlformats-officedocument.oleObject",
    relationships: undefined
  };
}

// =============================================================================
// Internal Helpers
// =============================================================================

function isOleEmbedding(path: string): boolean {
  return path.startsWith("word/embeddings/") || (path.endsWith(".bin") && path.includes("embed"));
}

function parseOlePartMetadata(part: OpaquePart): OleObject | null {
  // Extract what we can from the path and content type
  const fileName = getFileName(part.path);

  // Try to detect progId from file extension or OLE magic
  let progId = "Package"; // Default
  if (fileName.includes("oleObject")) {
    progId = detectProgIdFromData(part.data);
  }

  return {
    rId: "",
    progId,
    objectType: "embedded",
    displayAs: "icon",
    fileName,
    data: part.data
  };
}

function detectProgIdFromData(data: Uint8Array): string {
  // OLE2 compound document magic: D0 CF 11 E0 A1 B1 1A E1
  if (
    data.length >= 8 &&
    data[0] === 0xd0 &&
    data[1] === 0xcf &&
    data[2] === 0x11 &&
    data[3] === 0xe0
  ) {
    // It's an OLE2 compound document — try to detect type
    // Look for common signatures within the embedded data
    const str = tryDecodeAscii(data.slice(0, 2048));
    if (str.includes("Excel") || str.includes("Workbook") || str.includes("Sheet")) {
      return "Excel.Sheet.12";
    }
    if (str.includes("PowerPoint") || str.includes("Presentation")) {
      return "PowerPoint.Show.12";
    }
    if (str.includes("Visio")) {
      return "Visio.Drawing.15";
    }
    return "Package";
  }

  // PDF signature
  if (
    data.length >= 4 &&
    data[0] === 0x25 &&
    data[1] === 0x50 &&
    data[2] === 0x44 &&
    data[3] === 0x46
  ) {
    return "AcroExch.Document";
  }

  return "Package";
}

function tryDecodeAscii(data: Uint8Array): string {
  let str = "";
  for (let i = 0; i < data.length; i++) {
    const b = data[i]!;
    if (b >= 0x20 && b < 0x7f) {
      str += String.fromCharCode(b);
    }
  }
  return str;
}

function extractOleFromRawXml(rawXml: string): OleObject | null {
  // Parse OLE object info from raw XML using regex (lightweight)
  const progIdMatch = rawXml.match(/ProgID="([^"]+)"/i) ?? rawXml.match(/progId="([^"]+)"/i);
  const rIdMatch = rawXml.match(/r:id="([^"]+)"/i);
  const typeMatch = rawXml.match(/Type="([^"]+)"/i);

  if (!progIdMatch) {
    return null;
  }

  const progId = progIdMatch[1]!;
  const rId = rIdMatch ? rIdMatch[1]! : "";
  const objectType: OleObjectType =
    typeMatch && typeMatch[1]!.toLowerCase().includes("link") ? "linked" : "embedded";

  // Extract dimensions
  const widthMatch = rawXml.match(/(?:cx|width)="(\d+)"/i);
  const heightMatch = rawXml.match(/(?:cy|height)="(\d+)"/i);

  return {
    rId,
    progId,
    objectType,
    displayAs: rawXml.includes('DrawAspect="Icon"') ? "icon" : "content",
    width: widthMatch ? parseInt(widthMatch[1]!, 10) : undefined,
    height: heightMatch ? parseInt(heightMatch[1]!, 10) : undefined,
    rawXml
  };
}
