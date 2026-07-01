/**
 * PDF annotation extractor.
 *
 * Extracts annotations from a PDF page's `/Annots` array.
 * Supports all standard annotation subtypes defined in PDF Reference 1.7, §12.5.
 *
 * Common annotation types:
 * - **Link** — Hyperlinks (URI, GoTo, GoToR)
 * - **Text** — Sticky notes / comments
 * - **FreeText** — Inline text annotations
 * - **Highlight / Underline / StrikeOut / Squiggly** — Text markup
 * - **Stamp** — Rubber stamp annotations
 * - **Popup** — Associated popup windows
 * - **Widget** — Form field widgets (handled separately by form-extractor)
 *
 * @see PDF Reference 1.7, §12.5 - Annotations
 */

import type { PdfDocument } from "@pdf/reader/pdf-document";
import type { PdfDictValue, PdfObject } from "@pdf/reader/pdf-parser";
import {
  isPdfArray,
  dictGetName,
  dictGetNumber,
  decodePdfStringBytes
} from "@pdf/reader/pdf-parser";
import { getDictStringValue } from "@pdf/reader/reader-utils";

// =============================================================================
// Types
// =============================================================================

/** Rectangle in PDF coordinate space [x1, y1, x2, y2] */
export interface PdfRect {
  /** Left edge (x1) */
  x1: number;
  /** Bottom edge (y1) */
  y1: number;
  /** Right edge (x2) */
  x2: number;
  /** Top edge (y2) */
  y2: number;
}

/** A PDF annotation extracted from a page. */
export interface PdfAnnotation {
  /** Annotation subtype (e.g. "Link", "Text", "Highlight", "FreeText", "Stamp") */
  subtype: string;
  /** Bounding rectangle in page coordinates (points) */
  rect: PdfRect;
  /** Text content (/Contents entry) */
  contents: string;
  /** Author / title (/T entry) */
  author: string;
  /** Subject (/Subj entry) */
  subject: string;
  /** Modification date (/M entry) — raw PDF date string */
  modifiedDate: string;
  /** For Link annotations: the destination URI */
  uri: string;
  /** For Link annotations: named destination */
  destination: string;
  /** Annotation flags (/F entry) */
  flags: number;
  /** Color (/C entry) — array of 0-3 values in [0,1] */
  color: number[];
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Extract annotations from a PDF page.
 *
 * Skips Widget annotations (form fields) — those are handled by the form extractor.
 *
 * @param pageDict - The page dictionary
 * @param doc - The PDF document for resolving references
 * @returns Array of extracted annotations
 */
export function extractAnnotationsFromPage(
  pageDict: PdfDictValue,
  doc: PdfDocument
): PdfAnnotation[] {
  const annotsObj = pageDict.get("Annots");
  if (!annotsObj) {
    return [];
  }

  // Resolve the Annots array (may be an indirect reference)
  const annotsResolved = doc.deref(annotsObj);
  if (!isPdfArray(annotsResolved)) {
    return [];
  }

  const annotations: PdfAnnotation[] = [];

  for (const annotRef of annotsResolved) {
    try {
      const annotDict = doc.derefDict(annotRef);
      if (!annotDict) {
        continue;
      }

      const subtype = dictGetName(annotDict, "Subtype") ?? "";

      // Skip Widget annotations — handled by form-extractor
      if (subtype === "Widget") {
        continue;
      }

      // Skip Popup annotations — they are auxiliary
      if (subtype === "Popup") {
        continue;
      }

      const annotation = parseAnnotation(annotDict, subtype, doc);
      if (annotation) {
        annotations.push(annotation);
      }
    } catch {
      // Skip malformed annotations
    }
  }

  return annotations;
}

// =============================================================================
// Parsing
// =============================================================================

function parseAnnotation(
  dict: PdfDictValue,
  subtype: string,
  doc: PdfDocument
): PdfAnnotation | null {
  const rect = parseRect(dict.get("Rect"), doc);
  if (!rect) {
    return null;
  }

  const contents = getDictStringValue(dict, "Contents", doc);
  const author = getDictStringValue(dict, "T", doc);
  const subject = getDictStringValue(dict, "Subj", doc);
  const modifiedDate = getDictStringValue(dict, "M", doc);
  const flags = dictGetNumber(dict, "F") ?? 0;
  const color = parseColorArray(dict.get("C"), doc);

  // Extract link-specific fields
  let uri = "";
  let destination = "";

  if (subtype === "Link") {
    const actionObj = doc.derefDict(dict.get("A"));
    if (actionObj) {
      const actionType = dictGetName(actionObj, "S");
      if (actionType === "URI") {
        uri = getDictStringValue(actionObj, "URI", doc);
      } else if (actionType === "GoTo") {
        const dest = actionObj.get("D");
        if (typeof dest === "string") {
          destination = dest;
        }
      } else if (actionType === "GoToR") {
        uri = getDictStringValue(actionObj, "F", doc);
      }
    }

    // Check /Dest directly (older PDFs use this instead of /A)
    if (!uri && !destination) {
      const destObj = dict.get("Dest");
      if (destObj) {
        const resolved = doc.deref(destObj);
        if (typeof resolved === "string") {
          destination = resolved;
        } else if (resolved instanceof Uint8Array) {
          destination = decodePdfStringBytes(resolved);
        }
      }
    }
  }

  return {
    subtype,
    rect,
    contents,
    author,
    subject,
    modifiedDate,
    uri,
    destination,
    flags,
    color
  };
}

function parseRect(obj: PdfObject | undefined, doc: PdfDocument): PdfRect | null {
  if (!obj) {
    return null;
  }

  const resolved = doc.deref(obj);
  if (!isPdfArray(resolved) || resolved.length < 4) {
    return null;
  }

  const nums = resolved.map(v => (typeof v === "number" ? v : 0));
  return {
    x1: nums[0],
    y1: nums[1],
    x2: nums[2],
    y2: nums[3]
  };
}

function parseColorArray(obj: PdfObject | undefined, doc: PdfDocument): number[] {
  if (!obj) {
    return [];
  }

  const resolved = doc.deref(obj);
  if (!isPdfArray(resolved)) {
    return [];
  }

  return resolved.map(v => (typeof v === "number" ? v : 0));
}
