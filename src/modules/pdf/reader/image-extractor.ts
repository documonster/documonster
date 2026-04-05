/**
 * PDF image extraction.
 *
 * Extracts images from PDF pages including:
 * - Inline images (BI/ID/EI operators)
 * - XObject images (/Subtype /Image)
 * - Images with various color spaces and filters
 *
 * Supported image formats:
 * - JPEG (DCTDecode) — extracted as-is
 * - JPEG2000 (JPXDecode) — extracted as-is
 * - Raw/Flate-compressed pixel data — extracted with metadata
 * - CCITT fax — extracted as-is
 *
 * @see PDF Reference 1.7, §4.8 - Images
 */

import type { PdfDocument } from "./pdf-document";
import type { PdfDictValue, PdfStream } from "./pdf-parser";
import { isPdfRef, isPdfArray, dictGetName, dictGetNumber } from "./pdf-parser";

// =============================================================================
// Types
// =============================================================================

/**
 * An extracted image from a PDF page.
 */
export interface ExtractedImage {
  /** Image index within the page (0-based) */
  index: number;
  /** Image width in pixels */
  width: number;
  /** Image height in pixels */
  height: number;
  /** Bits per component */
  bitsPerComponent: number;
  /** Color space name */
  colorSpace: string;
  /** Number of color components (1=gray, 3=RGB, 4=CMYK) */
  components: number;
  /**
   * Image data format:
   * - "jpeg" — raw JPEG data (can be written directly as .jpg)
   * - "jpx" — JPEG 2000 data
   * - "raw" — raw pixel data (RGB/CMYK/Gray, decompressed)
   * - "ccitt" — CCITT fax compressed data
   */
  format: "jpeg" | "jpx" | "raw" | "ccitt" | "jbig2";
  /** The image data */
  data: Uint8Array;
  /** Alpha mask data (if present) — same dimensions, 1 component, 8 bits */
  alphaMask: Uint8Array | null;
  /** Filter name from the original stream */
  filter: string;
  /** XObject name (if it was a named XObject) */
  name: string;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Extract all images from a PDF page.
 */
export function extractImagesFromPage(pageDict: PdfDictValue, doc: PdfDocument): ExtractedImage[] {
  const images: ExtractedImage[] = [];

  // Get page resources (centralized with cycle protection)
  const resources = doc.resolvePageResources(pageDict);
  const xobjects = resources.get("XObject");
  if (!xobjects) {
    return images;
  }

  const xobjDict = doc.derefDict(xobjects);
  if (!xobjDict) {
    return images;
  }

  let index = 0;
  for (const [name, ref] of xobjDict) {
    const result = doc.derefStreamWithObjNum(ref);
    if (!result) {
      continue;
    }

    const streamDict = result.stream.dict;
    const subtype = dictGetName(streamDict, "Subtype");
    if (subtype !== "Image") {
      continue;
    }

    const image = extractImage(
      name,
      result.stream,
      streamDict,
      doc,
      index,
      result.objNum,
      result.gen
    );
    if (image) {
      images.push(image);
      index++;
    }
  }

  return images;
}

// =============================================================================
// Image Extraction
// =============================================================================

function extractImage(
  name: string,
  stream: PdfStream,
  dict: PdfDictValue,
  doc: PdfDocument,
  index: number,
  objNum = 0,
  gen = 0
): ExtractedImage | null {
  const width = dictGetNumber(dict, "Width") ?? dictGetNumber(dict, "W") ?? 0;
  const height = dictGetNumber(dict, "Height") ?? dictGetNumber(dict, "H") ?? 0;
  const bpc = dictGetNumber(dict, "BitsPerComponent") ?? dictGetNumber(dict, "BPC") ?? 8;

  if (width === 0 || height === 0) {
    return null;
  }

  // Determine color space
  const { colorSpace, components } = resolveColorSpace(dict, doc);

  // Determine filter to understand the image format
  const filter = getFilterName(dict);

  // Extract image data based on filter
  // For all formats, use getStreamData which handles decryption and filter decoding
  let data: Uint8Array;
  let format: ExtractedImage["format"];

  if (filter === "DCTDecode" || filter === "DCT") {
    // JPEG — use getStreamData which handles decryption properly
    data = doc.getStreamData(stream, objNum, gen);
    format = "jpeg";
  } else if (filter === "JPXDecode") {
    data = doc.getStreamData(stream, objNum, gen);
    format = "jpx";
  } else if (filter === "CCITTFaxDecode" || filter === "CCF") {
    data = doc.getStreamData(stream, objNum, gen);
    format = "ccitt";
  } else if (filter === "JBIG2Decode") {
    data = doc.getStreamData(stream, objNum, gen);
    format = "jbig2";
  } else {
    // Decode all filters to get raw pixel data
    data = doc.getStreamData(stream, objNum, gen);
    format = "raw";
  }

  // Extract soft mask (alpha channel)
  let alphaMask: Uint8Array | null = null;
  const smaskRef = dict.get("SMask");
  if (smaskRef) {
    const smaskResult = doc.derefStreamWithObjNum(smaskRef);
    if (smaskResult) {
      alphaMask = doc.getStreamData(smaskResult.stream, smaskResult.objNum, smaskResult.gen);
    }
  }

  return {
    index,
    width,
    height,
    bitsPerComponent: bpc,
    colorSpace,
    components,
    format,
    data,
    alphaMask,
    filter,
    name
  };
}

// =============================================================================
// Color Space Resolution
// =============================================================================

function resolveColorSpace(
  dict: PdfDictValue,
  doc: PdfDocument
): { colorSpace: string; components: number } {
  const cs = dict.get("ColorSpace") ?? dict.get("CS");

  if (typeof cs === "string") {
    return colorSpaceInfo(cs);
  }

  if (isPdfArray(cs) && cs.length > 0) {
    const csName = cs[0];
    if (typeof csName === "string") {
      if (csName === "ICCBased") {
        // ICC-based color space — get N from the profile stream
        if (cs.length > 1) {
          const profileStream = doc.derefStream(cs[1]);
          if (profileStream) {
            const n = dictGetNumber(profileStream.dict, "N") ?? 3;
            return {
              colorSpace: "ICCBased",
              components: n
            };
          }
        }
        return { colorSpace: "ICCBased", components: 3 };
      }
      if (csName === "Indexed" || csName === "I") {
        return { colorSpace: "Indexed", components: 1 };
      }
      if (csName === "Separation") {
        return { colorSpace: "Separation", components: 1 };
      }
      if (csName === "DeviceN") {
        const numComponents = isPdfArray(cs[1]) ? cs[1].length : 1;
        return { colorSpace: "DeviceN", components: numComponents };
      }
      return colorSpaceInfo(csName);
    }
  }

  if (isPdfRef(cs)) {
    const resolved = doc.deref(cs);
    if (typeof resolved === "string") {
      return colorSpaceInfo(resolved);
    }
    if (isPdfArray(resolved) && resolved.length > 0 && typeof resolved[0] === "string") {
      return resolveColorSpace(new Map([["ColorSpace", resolved]]) as PdfDictValue, doc);
    }
  }

  return { colorSpace: "DeviceRGB", components: 3 };
}

function colorSpaceInfo(name: string): { colorSpace: string; components: number } {
  switch (name) {
    case "DeviceGray":
    case "G":
    case "CalGray":
      return { colorSpace: name, components: 1 };
    case "DeviceRGB":
    case "RGB":
    case "CalRGB":
      return { colorSpace: name, components: 3 };
    case "DeviceCMYK":
    case "CMYK":
      return { colorSpace: name, components: 4 };
    default:
      return { colorSpace: name, components: 3 };
  }
}

// =============================================================================
// Helpers
// =============================================================================

/**
 * Get the image-specific filter name from a stream dictionary.
 * For filter chains, finds the last image-specific filter (DCT, JPX, CCITT, JBIG2).
 * For non-image-specific chains (e.g., just FlateDecode), returns that filter.
 */
function getFilterName(dict: PdfDictValue): string {
  const filter = dict.get("Filter") ?? dict.get("F");
  if (typeof filter === "string") {
    return filter;
  }
  if (isPdfArray(filter) && filter.length > 0) {
    // Look for the last image-specific filter in the chain
    const imageFilters = new Set([
      "DCTDecode",
      "DCT",
      "JPXDecode",
      "CCITTFaxDecode",
      "CCF",
      "JBIG2Decode"
    ]);
    for (let i = filter.length - 1; i >= 0; i--) {
      const f = filter[i];
      if (typeof f === "string" && imageFilters.has(f)) {
        return f;
      }
    }
    // No image-specific filter found — return the first filter
    const first = filter[0];
    if (typeof first === "string") {
      return first;
    }
  }
  return "";
}
