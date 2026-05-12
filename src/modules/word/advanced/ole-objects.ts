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

import { getFileName } from "../core/opc-paths";
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
      // Also check if the file path matches the expected pattern for this
      // rId. We anchor on a non-digit boundary so `rId4` does not also
      // match `oleObject40`, `oleObject41`, … Without the boundary, the
      // first such adjacent file silently leaks back as if it were the
      // requested OLE object.
      const expectedStem = rId.replace("rId", "oleObject");
      const oleStemRe = new RegExp(`(^|/)${escapeRegex(expectedStem)}(?:\\.[^./]+)?$`);
      if (oleStemRe.test(part.path)) {
        return part.data;
      }
    }
  }

  // Unknown rId — do NOT fall back to "first embedded blob" (that historically
  // returned an arbitrary OLE object and silently leaked the wrong data).
  return undefined;
}

/**
 * Result of `createOleEmbedding`. Contains the OLE binary part and,
 * when a preview image was supplied, an additional media part for it
 * plus the relationship rIds the caller should reference from the body
 * model.
 */
export interface OleEmbeddingResult {
  /** OLE binary opaque part (always present). */
  readonly olePart: OpaquePart;
  /** Suggested rId to use for the OLE binary in the document model. */
  readonly oleRId: string;
  /** Preview image media part (only when `options.previewImage` was supplied). */
  readonly previewPart?: OpaquePart;
  /** Suggested rId for the preview image. */
  readonly previewRId?: string;
}

/**
 * Create an OLE embedding plus, optionally, its preview image.
 *
 * Returns an {@link OleEmbeddingResult} with one or two `OpaquePart`s plus
 * suggested relationship IDs. Each call gets a unique counter-based file
 * name so calling this helper repeatedly does not produce path
 * collisions; pass `options.fileName` to force a specific name.
 *
 * The previous signature returned only one `OpaquePart` and silently
 * dropped `options.previewImage` / `options.previewContentType`. That has
 * been removed — the new shape forces callers to wire up preview parts
 * properly when they want one.
 *
 * @param data - The binary data of the OLE object (OLE2 compound document).
 * @param progId - The OLE ProgId (e.g. "Excel.Sheet.12").
 * @param options - Additional options.
 */
export function createOleEmbedding(
  data: Uint8Array,
  progId: string,
  options?: {
    /** Override file name. Defaults to `oleObject<N>.bin` with a process-unique counter. */
    fileName?: string;
    /** Optional preview image bytes (typically EMF or PNG). */
    previewImage?: Uint8Array;
    /** Content type for the preview image. Required when `previewImage` is set. */
    previewContentType?: string;
    /** Override preview file name. Defaults to `image<N>.<ext>` from previewContentType. */
    previewFileName?: string;
  }
): OleEmbeddingResult {
  const oleSeq = nextOleSequence();
  const fileName = options?.fileName ?? `oleObject${oleSeq}.bin`;
  const olePart: OpaquePart = {
    path: `word/embeddings/${fileName}`,
    data,
    contentType: "application/vnd.openxmlformats-officedocument.oleObject",
    relationships: undefined
  };
  const oleRId = `rIdOle${oleSeq}`;
  // progId is metadata for downstream consumers — not stored on
  // OpaquePart but accepted in the signature so callers can pass it
  // alongside without a separate channel. We don't need it here.
  void progId;

  if (!options?.previewImage) {
    return { olePart, oleRId };
  }
  if (!options.previewContentType) {
    throw new Error("createOleEmbedding: options.previewImage requires options.previewContentType");
  }
  const previewExt = previewExtFromContentType(options.previewContentType);
  const previewSeq = nextPreviewSequence();
  const previewFileName = options.previewFileName ?? `oleImage${previewSeq}.${previewExt}`;
  const previewPart: OpaquePart = {
    path: `word/media/${previewFileName}`,
    data: options.previewImage,
    contentType: options.previewContentType,
    relationships: undefined
  };
  const previewRId = `rIdOleImg${previewSeq}`;
  return { olePart, oleRId, previewPart, previewRId };
}

/** Module-level counters used to allocate unique file names per call. */
let _oleSeq = 0;
let _previewSeq = 0;
function nextOleSequence(): number {
  _oleSeq++;
  return _oleSeq;
}
function nextPreviewSequence(): number {
  _previewSeq++;
  return _previewSeq;
}

function previewExtFromContentType(ct: string): string {
  const map: Record<string, string> = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/x-emf": "emf",
    "image/emf": "emf",
    "image/x-wmf": "wmf",
    "image/wmf": "wmf",
    "image/bmp": "bmp",
    "image/svg+xml": "svg"
  };
  return map[ct.toLowerCase()] ?? "bin";
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

/** Escape regex meta-characters in a literal string for use inside RegExp. */
function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
