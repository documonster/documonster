/**
 * Shared image utilities for PDF generation.
 *
 * Centralises JPEG/PNG dimension parsing and PDF XObject writing so that
 * both the builder (`document-builder.ts`) and the exporter (`pdf-exporter.ts`)
 * share a single implementation.
 */

import { PdfDict, pdfRef, pdfNumber } from "../core/pdf-object";
import type { PdfWriter } from "../core/pdf-writer";
import { decodePng } from "../render/png-decoder";

// =============================================================================
// Image Dimension Parsing
// =============================================================================

/**
 * Parse image dimensions from raw bytes.
 */
export function parseImageDimensions(
  data: Uint8Array,
  format: "jpeg" | "png"
): { width: number; height: number } {
  if (format === "png") {
    return parsePngDimensions(data);
  }
  return parseJpegDimensions(data);
}

/**
 * Read width/height from a PNG IHDR chunk (bytes 16-23).
 */
export function parsePngDimensions(data: Uint8Array): { width: number; height: number } {
  // PNG header: 8 byte signature, then IHDR chunk: 4 byte length, 4 byte type, 4 byte width, 4 byte height
  if (
    data.length >= 24 &&
    data[12] === 0x49 &&
    data[13] === 0x48 &&
    data[14] === 0x44 &&
    data[15] === 0x52
  ) {
    const width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
    const height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
    return { width, height };
  }
  return { width: 1, height: 1 };
}

/**
 * Read width/height from JPEG SOF marker.
 *
 * Correctly excludes non-SOF markers in the 0xC0-0xCF range:
 * - 0xC4 = DHT (Define Huffman Table)
 * - 0xC8 = JPG (reserved)
 * - 0xCC = DAC (Define Arithmetic Coding)
 */
export function parseJpegDimensions(data: Uint8Array): { width: number; height: number } {
  let offset = 2; // skip SOI marker
  while (offset < data.length - 1) {
    // Skip padding 0xFF bytes
    while (offset < data.length && data[offset] === 0xff && data[offset + 1] === 0xff) {
      offset++;
    }
    if (offset >= data.length - 1 || data[offset] !== 0xff) {
      break;
    }
    const marker = data[offset + 1];
    const isSof =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof && offset + 8 < data.length) {
      return {
        width: (data[offset + 7] << 8) | data[offset + 8],
        height: (data[offset + 5] << 8) | data[offset + 6]
      };
    }
    if (offset + 3 >= data.length) {
      break;
    }
    const segLen = (data[offset + 2] << 8) | data[offset + 3];
    offset += 2 + segLen;
  }
  return { width: 1, height: 1 };
}

// =============================================================================
// PDF Image XObject Writing
// =============================================================================

/**
 * Write an image XObject (JPEG or PNG) to the writer.
 * Returns the allocated object number.
 */
export function writeImageXObject(
  writer: PdfWriter,
  data: Uint8Array,
  format: "jpeg" | "png"
): number {
  if (format === "png") {
    return writePngImageXObject(writer, data);
  }
  return writeJpegImageXObject(writer, data);
}

/**
 * Write a JPEG image using DCTDecode (raw JPEG data embedded directly).
 */
function writeJpegImageXObject(writer: PdfWriter, data: Uint8Array): number {
  const objNum = writer.allocObject();
  const dims = parseJpegDimensions(data);
  const dict = new PdfDict()
    .set("Type", "/XObject")
    .set("Subtype", "/Image")
    .set("Width", pdfNumber(dims.width))
    .set("Height", pdfNumber(dims.height))
    .set("ColorSpace", "/DeviceRGB")
    .set("BitsPerComponent", "8")
    .set("Filter", "/DCTDecode");
  writer.addStreamObject(objNum, dict, data);
  return objNum;
}

/**
 * Write a PNG image: decode to raw RGB, create SMask for alpha if needed.
 */
function writePngImageXObject(writer: PdfWriter, data: Uint8Array): number {
  const png = decodePng(data);
  const objNum = writer.allocObject();

  const dict = new PdfDict()
    .set("Type", "/XObject")
    .set("Subtype", "/Image")
    .set("Width", pdfNumber(png.width))
    .set("Height", pdfNumber(png.height))
    .set("ColorSpace", "/DeviceRGB")
    .set("BitsPerComponent", pdfNumber(png.bitsPerComponent));

  if (png.alpha) {
    const smaskObjNum = writer.allocObject();
    const smaskDict = new PdfDict()
      .set("Type", "/XObject")
      .set("Subtype", "/Image")
      .set("Width", pdfNumber(png.width))
      .set("Height", pdfNumber(png.height))
      .set("ColorSpace", "/DeviceGray")
      .set("BitsPerComponent", "8");
    writer.addStreamObject(smaskObjNum, smaskDict, png.alpha);
    dict.set("SMask", pdfRef(smaskObjNum));
  }

  writer.addStreamObject(objNum, dict, png.pixels);
  return objNum;
}
