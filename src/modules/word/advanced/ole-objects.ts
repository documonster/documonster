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

import { asPlainBytes } from "@utils/binary";
import { escapeRegExp } from "@utils/regexp";
import { ContentType } from "@word/constants";
import {
  checkBytes,
  checkFinite,
  checkHandle,
  checkOneOf,
  checkRecord,
  checkString,
  invalidArgument
} from "@word/core/argument-checks";
import { getFileName } from "@word/core/opc-paths";
import { markDerived } from "@word/core/ownership";
import type {
  BodyContent,
  DocxDocument,
  OleObjectPart,
  OpaqueDrawing,
  OpaquePart
} from "@word/types";
import { xmlEncodeAttr } from "@xml/encode";

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
  checkHandle(doc, "document", "doc", "Ole.extract");
  const objects: OleObject[] = [];
  const summary: Record<string, number> = {};

  const pushObject = (obj: OleObject): void => {
    objects.push(obj);
    summary[obj.progId] = (summary[obj.progId] ?? 0) + 1;
  };

  // Structured OLE objects wired on document.xml.rels (preferred form —
  // these carry the real rId and, when available, the round-tripped progId).
  if (doc.oleObjects) {
    for (const ole of doc.oleObjects) {
      pushObject({
        rId: ole.rId,
        progId: ole.progId ?? detectProgIdFromData(ole.data),
        objectType: "embedded",
        displayAs: "icon",
        imageRId: ole.previewRId,
        fileName: getFileName(ole.path),
        data: ole.data
      });
    }
  }

  // Scan opaque parts for OLE embeddings (legacy / hand-built documents that
  // did not go through the structured oleObjects channel).
  if (doc.opaqueParts) {
    for (const part of doc.opaqueParts) {
      if (isOleEmbedding(part.path)) {
        const obj = parseOlePartMetadata(part);
        if (obj) {
          pushObject(obj);
        }
      }
    }
  }

  // Also scan body for OpaqueDrawing elements that may contain OLE references
  for (const element of doc.body) {
    if (element.type === "opaqueDrawing") {
      const oleFromDrawing = extractOleFromRawXml(element.rawXml);
      if (oleFromDrawing) {
        // Check if we already have this object (by rId)
        const exists = objects.some(o => o.rId === oleFromDrawing.rId && o.rId !== "");
        if (!exists) {
          pushObject(oleFromDrawing);
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
  checkHandle(doc, "document", "doc", "Ole.has");
  if (doc.oleObjects && doc.oleObjects.length > 0) {
    return true;
  }
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
export function getOleObjectData(
  doc: DocxDocument,
  rId: string
): Uint8Array<ArrayBuffer> | undefined {
  checkHandle(doc, "document", "doc", "Ole.getData");
  // Structured OLE objects carry the exact rId used on document.xml.rels.
  if (doc.oleObjects) {
    for (const ole of doc.oleObjects) {
      if (ole.rId === rId) {
        return ole.data && asPlainBytes(ole.data);
      }
    }
  }

  if (!doc.opaqueParts) {
    return undefined;
  }

  for (const part of doc.opaqueParts) {
    if (isOleEmbedding(part.path)) {
      // Check relationships for matching rId
      if (part.relationships) {
        for (const rel of part.relationships) {
          if (rel.id === rId) {
            return asPlainBytes(part.data);
          }
        }
      }
      // Also check if the file path matches the expected pattern for this
      // rId. We anchor on a non-digit boundary so `rId4` does not also
      // match `oleObject40`, `oleObject41`, … Without the boundary, the
      // first such adjacent file silently leaks back as if it were the
      // requested OLE object.
      const expectedStem = rId.replace("rId", "oleObject");
      const oleStemRe = new RegExp(`(^|/)${escapeRegExp(expectedStem)}(?:\\.[^./]+)?$`);
      if (oleStemRe.test(part.path)) {
        return asPlainBytes(part.data);
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
  /** OLE ProgId the embedding was created with (e.g. "Excel.Sheet.12"). */
  readonly progId: string;
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
  const where = "Ole.createEmbedding";
  data = checkBytes(data, "data", where);
  checkString(progId, "progId", where);
  if (progId === "") {
    invalidArgument(where, "progId", "a non-empty string", progId);
  }
  checkRecord(options, "options", where, { optional: true });
  checkPartFileName(options?.fileName, "options.fileName", where);
  checkPartFileName(options?.previewFileName, "options.previewFileName", where);
  if (options?.previewImage !== undefined) {
    options = {
      ...options,
      previewImage: checkBytes(options.previewImage, "options.previewImage", where)
    };
    // The preview needs a type for `[Content_Types].xml`.
    checkString(options.previewContentType, "options.previewContentType", where);
  }
  const oleSeq = nextOleSequence();
  const fileName = options?.fileName ?? `oleObject${oleSeq}.bin`;
  const olePart: OpaquePart = {
    path: `word/embeddings/${fileName}`,
    data,
    contentType: ContentType.OleObject,
    relationships: undefined
  };
  const oleRId = `rIdOle${oleSeq}`;
  // progId is carried back on the result so addOleObject() can persist it
  // into the body `<o:OLEObject ProgID="…">` markup for round-trip.

  if (!options?.previewImage) {
    return { olePart, oleRId, progId };
  }
  if (!options.previewContentType) {
    invalidArgument(
      where,
      "options.previewContentType",
      "a content type when options.previewImage is set",
      options.previewContentType
    );
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
  return { olePart, oleRId, progId, previewPart, previewRId };
}

/**
 * Wire an {@link OleEmbeddingResult} into a document so the OLE object is
 * actually rendered and resolvable, returning a new {@link DocxDocument}.
 *
 * Unlike just stuffing the part into `opaqueParts` (which leaves the binary
 * dangling — no relationship, no body reference), this:
 *
 *  - registers the OLE binary (and optional preview) on
 *    `doc.oleObjects` so the packager emits a `word/_rels/document.xml.rels`
 *    relationship with the exact rId and a `[Content_Types].xml` override;
 *  - appends a body paragraph carrying a `<w:object>` / `<o:OLEObject>`
 *    that references the same rId and embeds the ProgId, so the object is
 *    visible in Word and round-trips through `readDocx`.
 *
 * Does not modify `doc`; the result may share unchanged nodes with it; `Document.*`
 * copies what it edits, so an edit made through one never reaches the other.
 *
 * The embedding's relationship id and part paths are kept unless `doc`
 * already uses them — adding the same embedding twice, or adding one to a
 * document read back from an earlier run — in which case the object is given
 * unused ones; read the result's `oleObjects` for the names it ended up with.
 *
 * @param doc - The document to add the OLE object to.
 * @param embedding - Result from {@link createOleEmbedding}.
 * @param options - Display geometry (defaults to a 2"×2" icon box).
 */
export function addOleObject(
  doc: DocxDocument,
  embedding: OleEmbeddingResult,
  options?: {
    /** Display width in points (default 96 = 2 inches at 48pt/in icon). */
    widthPt?: number;
    /** Display height in points (default 96). */
    heightPt?: number;
    /** Display mode. Default "icon". */
    displayAs?: OleDisplayAs;
  }
): DocxDocument {
  checkHandle(doc, "document", "doc", "Ole.add");
  const where = "Ole.add";
  checkRecord(embedding, "embedding", where);
  checkRecord(embedding.olePart, "embedding.olePart", where);
  checkString(embedding.oleRId, "embedding.oleRId", where);
  checkString(embedding.progId, "embedding.progId", where);
  checkRecord(options, "options", where, { optional: true });
  checkFinite(options?.widthPt, "options.widthPt", where, { optional: true, min: 0 });
  checkFinite(options?.heightPt, "options.heightPt", where, { optional: true, min: 0 });
  checkOneOf(options?.displayAs, DISPLAY_AS, "options.displayAs", where, { optional: true });
  const widthPt = options?.widthPt ?? 96;
  const heightPt = options?.heightPt ?? 96;
  const drawAspect = (options?.displayAs ?? "icon") === "icon" ? "Icon" : "Content";
  const placed = withUnusedOleNames(doc, embedding);

  const olePartEntry: OleObjectPart = {
    path: placed.olePart.path,
    data: placed.olePart.data,
    rId: placed.oleRId,
    progId: placed.progId,
    contentType: placed.olePart.contentType,
    ...(placed.previewPart && placed.previewRId
      ? {
          previewPath: placed.previewPart.path,
          previewData: placed.previewPart.data,
          previewRId: placed.previewRId,
          previewContentType: placed.previewPart.contentType
        }
      : {})
  };

  // Build the VML-hosted <w:object>. The o:OLEObject carries ProgID + the
  // r:id of the binary; the v:shape provides geometry and (when present) the
  // preview image fill via v:imagedata. This is the canonical OOXML shape for
  // an embedded OLE object (ECMA-376 §17.3.3.19 + VML).
  const shapeId = `_ole_${placed.oleRId}`;
  const styleWidth = widthPt.toFixed(0);
  const styleHeight = heightPt.toFixed(0);
  const imageData =
    placed.previewRId != null
      ? `<v:imagedata r:id="${xmlEncodeAttr(placed.previewRId)}" o:title=""/>`
      : "";
  const rawXml =
    `<w:object>` +
    `<v:shape id="${xmlEncodeAttr(shapeId)}" type="#_x0000_t75" ` +
    `style="width:${styleWidth}pt;height:${styleHeight}pt">` +
    imageData +
    `</v:shape>` +
    `<o:OLEObject Type="Embed" ProgID="${xmlEncodeAttr(placed.progId)}" ` +
    `ShapeID="${xmlEncodeAttr(shapeId)}" DrawAspect="${drawAspect}" ` +
    `r:id="${xmlEncodeAttr(placed.oleRId)}"/>` +
    `</w:object>`;

  const referencedRIds = [placed.oleRId];
  if (placed.previewRId != null) {
    referencedRIds.push(placed.previewRId);
  }
  const drawing: OpaqueDrawing = {
    type: "opaqueDrawing",
    rawXml,
    referencedRIds
  };

  return markDerived(
    {
      ...doc,
      body: [...doc.body, drawing as BodyContent],
      oleObjects: [...(doc.oleObjects ?? []), olePartEntry]
    },
    doc
  );
}

const DISPLAY_AS: ReadonlySet<string> = new Set<OleDisplayAs>(["icon", "content"]);

/**
 * `embedding` with its relationship ids and part paths moved off any the
 * document already uses. Both are package-internal names: the counter that
 * mints them restarts with every process, so an embedding added to a document
 * read back from an earlier run — or the same embedding added twice — would
 * otherwise collide with an object already there and fail to package.
 */
function withUnusedOleNames(doc: DocxDocument, embedding: OleEmbeddingResult): OleEmbeddingResult {
  const usedIds = new Set<string>();
  const usedPaths = new Set<string>();
  for (const ole of doc.oleObjects ?? []) {
    usedIds.add(ole.rId);
    usedPaths.add(ole.path);
    if (ole.previewRId) {
      usedIds.add(ole.previewRId);
    }
    if (ole.previewPath) {
      usedPaths.add(ole.previewPath);
    }
  }
  for (const part of doc.opaqueParts ?? []) {
    usedPaths.add(part.path);
  }
  const freeId = (id: string): string => {
    let candidate = id;
    for (let n = 2; usedIds.has(candidate); n++) {
      candidate = `${id}_${n}`;
    }
    usedIds.add(candidate);
    return candidate;
  };
  const freePath = (path: string): string => {
    const dot = path.lastIndexOf(".");
    const stem = dot > path.lastIndexOf("/") ? path.slice(0, dot) : path;
    const ext = path.slice(stem.length);
    let candidate = path;
    for (let n = 2; usedPaths.has(candidate); n++) {
      candidate = `${stem}_${n}${ext}`;
    }
    usedPaths.add(candidate);
    return candidate;
  };
  const oleRId = freeId(embedding.oleRId);
  const olePath = freePath(embedding.olePart.path);
  const result: OleEmbeddingResult = {
    ...embedding,
    oleRId,
    olePart:
      olePath === embedding.olePart.path
        ? embedding.olePart
        : { ...embedding.olePart, path: olePath }
  };
  if (!embedding.previewPart || embedding.previewRId == null) {
    return result;
  }
  const previewPath = freePath(embedding.previewPart.path);
  return {
    ...result,
    previewRId: freeId(embedding.previewRId),
    previewPart:
      previewPath === embedding.previewPart.path
        ? embedding.previewPart
        : { ...embedding.previewPart, path: previewPath }
  };
}

/** A part's file name: it is joined to a folder, so it must not name another one. */
function checkPartFileName(value: unknown, field: string, where: string): void {
  if (value === undefined) {
    return;
  }
  if (
    typeof value === "string" &&
    value !== "" &&
    !/[/\\]/.test(value) &&
    value !== "." &&
    value !== ".."
  ) {
    return;
  }
  invalidArgument(where, field, "a file name with no folder", value);
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
  // Pull metadata from the <o:OLEObject> element specifically, so a preview
  // image's <v:imagedata r:id="…"> earlier in the markup is not mistaken for
  // the OLE binary's relationship id.
  const oleTag = rawXml.match(/<o:OLEObject\b[^>]*>/i)?.[0] ?? rawXml;
  const progIdMatch = oleTag.match(/ProgID="([^"]+)"/i) ?? oleTag.match(/progId="([^"]+)"/i);
  const rIdMatch = oleTag.match(/r:id="([^"]+)"/i);
  const typeMatch = oleTag.match(/Type="([^"]+)"/i);

  if (!progIdMatch) {
    return null;
  }

  const progId = progIdMatch[1]!;
  const rId = rIdMatch ? rIdMatch[1]! : "";
  const objectType: OleObjectType =
    typeMatch && typeMatch[1]!.toLowerCase().includes("link") ? "linked" : "embedded";

  // Extract dimensions (from the surrounding shape, hence full rawXml)
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
