/**
 * Shared image utilities for PDF generation.
 *
 * Centralises JPEG/PNG dimension parsing and PDF XObject writing so that
 * both the builder (`document-builder.ts`) and the exporter (`pdf-exporter.ts`)
 * share a single implementation.
 */

import { PdfDict, pdfRef, pdfNumber } from "@pdf/core/pdf-object";
import type { PdfWriter } from "@pdf/core/pdf-writer";
import { PdfArgumentError, PdfRenderError } from "@pdf/errors";
import { decodePng } from "@pdf/render/png-decoder";

// =============================================================================
// Image Dimension Parsing
// =============================================================================

const IMAGE_FORMATS: readonly string[] = ["jpeg", "png"];

/**
 * Throw {@link PdfArgumentError} unless `format` is one an image XObject can be written in. Anything else
 * used to fall through to the JPEG path, so PNG bytes passed as `"image/png"` were embedded as a
 * `DCTDecode` stream no viewer can decode.
 */
export function assertImageFormat(
  format: unknown,
  where: string,
  field = "format"
): asserts format is "jpeg" | "png" {
  if (typeof format !== "string" || !IMAGE_FORMATS.includes(format)) {
    throw new PdfArgumentError(
      field,
      `${where}: invalid image format ${JSON.stringify(format)}; expected one of ${IMAGE_FORMATS.join(", ")}`
    );
  }
}

/**
 * Parse image dimensions from raw bytes.
 */
export function parseImageDimensions(
  data: Uint8Array,
  format: "jpeg" | "png"
): { width: number; height: number } {
  assertImageFormat(format, "Image");
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
  const info = readJpegInfo(data);
  return info === undefined ? { width: 1, height: 1 } : { width: info.width, height: info.height };
}

/** What a JPEG's frame header says about it. */
export interface JpegInfo {
  readonly width: number;
  readonly height: number;
  /** Colour components: 1 (grey), 3 (YCbCr/RGB) or 4 (CMYK/YCCK). */
  readonly components: 1 | 3 | 4;
  /**
   * Whether an Adobe `APP14` segment is present. Adobe writes four-component JPEGs with
   * inverted samples, which a PDF consumer undoes with a `/Decode [1 0 …]` array.
   */
  readonly adobe: boolean;
}

/**
 * Read a JPEG's frame header, or `undefined` when the bytes are not a JPEG this writer can
 * embed: no `SOI` marker, no frame header, a zero size, or a component count PDF's
 * `DCTDecode` does not take. Embedding such bytes used to write a `/Width 1 /Height 1`
 * `DCTDecode` image no viewer can draw.
 */
export function readJpegInfo(data: Uint8Array): JpegInfo | undefined {
  if (data.length < 4 || data[0] !== 0xff || data[1] !== 0xd8) {
    return undefined;
  }
  let adobe = false;
  let offset = 2; // skip SOI marker
  while (offset < data.length - 1) {
    // Skip padding 0xFF bytes
    while (offset < data.length && data[offset] === 0xff && data[offset + 1] === 0xff) {
      offset++;
    }
    if (offset >= data.length - 1 || data[offset] !== 0xff) {
      return undefined;
    }
    const marker = data[offset + 1];
    const isSof =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      if (offset + 9 >= data.length) {
        return undefined;
      }
      const height = (data[offset + 5] << 8) | data[offset + 6];
      const width = (data[offset + 7] << 8) | data[offset + 8];
      const components = data[offset + 9];
      if (width === 0 || height === 0) {
        return undefined;
      }
      if (components !== 1 && components !== 3 && components !== 4) {
        return undefined;
      }
      return { width, height, components, adobe };
    }
    if (offset + 3 >= data.length) {
      return undefined;
    }
    const segLen = (data[offset + 2] << 8) | data[offset + 3];
    // APP14 "Adobe"
    if (
      marker === 0xee &&
      segLen >= 7 &&
      String.fromCharCode(...data.subarray(offset + 4, offset + 9)) === "Adobe"
    ) {
      adobe = true;
    }
    offset += 2 + segLen;
  }
  return undefined;
}

/**
 * Throw {@link PdfArgumentError} unless `data` holds an image of `format` this writer can
 * embed. Checked when the image is given rather than when the file is written, so a
 * mislabelled or truncated image is reported by the call that received it — PNG bytes passed
 * as `"jpeg"` used to be written as a `DCTDecode` stream no viewer can decode, and a PNG that
 * could not be decoded failed `toBuffer` with a plain `Error`.
 *
 * For PNG this reads the header (`IHDR`): the decoder takes 8-bit, non-interlaced images.
 */
export function checkImageData(
  where: string,
  field: string,
  data: Uint8Array,
  format: "jpeg" | "png"
): void {
  if (format === "jpeg") {
    if (readJpegInfo(data) === undefined) {
      throw new PdfArgumentError(
        field,
        `${where}: ${field} is not a JPEG image this writer can embed (it needs an SOI marker and a frame header with 1, 3 or 4 components)`
      );
    }
    return;
  }
  const problem = pngHeaderProblem(data);
  if (problem !== undefined) {
    throw new PdfArgumentError(
      field,
      `${where}: ${field} is not a PNG image this writer can embed: ${problem}`
    );
  }
}

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

/** Why a PNG's header cannot be decoded, or `undefined` when it can. */
function pngHeaderProblem(data: Uint8Array): string | undefined {
  if (data.length < 33 || PNG_SIGNATURE.some((byte, i) => data[i] !== byte)) {
    return "the PNG signature is missing";
  }
  if (String.fromCharCode(data[12], data[13], data[14], data[15]) !== "IHDR") {
    return "the first chunk is not IHDR";
  }
  const width = ((data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19]) >>> 0;
  const height = ((data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23]) >>> 0;
  if (width === 0 || height === 0) {
    return `its size is ${width}x${height}`;
  }
  if (data[24] !== 8) {
    return `bit depth ${data[24]} is not supported (only 8-bit PNGs are)`;
  }
  if (![0, 2, 3, 4, 6].includes(data[25])) {
    return `colour type ${data[25]} does not exist`;
  }
  if (data[28] !== 0) {
    return "interlaced PNGs are not supported";
  }
  return undefined;
}

// =============================================================================
// PDF Image XObject Writing
// =============================================================================

/** Where image objects go: a whole-file writer, or an incremental update's object table. */
export interface ImageObjectSink {
  /** A fresh object number. */
  allocate(): number;
  /** Write a stream object under `objNum`. */
  addStream(objNum: number, dict: PdfDict, data: Uint8Array): void;
}

/** A sink over a {@link PdfWriter}. */
function writerSink(writer: PdfWriter): ImageObjectSink {
  return {
    allocate: () => writer.allocObject(),
    addStream: (objNum, dict, data) => writer.addStreamObject(objNum, dict, data)
  };
}

/**
 * Write an image XObject (JPEG or PNG) to the writer.
 * Returns the allocated object number.
 */
export function writeImageXObject(
  writer: PdfWriter,
  data: Uint8Array,
  format: "jpeg" | "png",
  grayscale = false
): number {
  return addImageXObject(writerSink(writer), data, format, grayscale);
}

/**
 * Write an image XObject — and the soft mask or tint transform it needs — to `sink`.
 * Returns the image's object number.
 *
 * One implementation for every path that embeds an image: the incremental update used to
 * build its own dictionary, which wrote a PNG's compressed file bytes as if they were raw
 * pixels and dropped its alpha.
 */
export function addImageXObject(
  sink: ImageObjectSink,
  data: Uint8Array,
  format: "jpeg" | "png",
  grayscale = false
): number {
  assertImageFormat(format, "Image");
  if (format === "png") {
    return writePngImageXObject(sink, data, grayscale);
  }
  return writeJpegImageXObject(sink, data, grayscale);
}

/**
 * Write the PostScript calculator function that converts a colour to its Rec. 601 luma,
 * and return its object number.
 *
 * Used as the tint transform of a `/DeviceN` space so a JPEG can be recoloured
 * without decoding it: the samples stay untouched and the consumer applies the
 * transform per pixel. Unlike painting a blend on top, this cannot darken
 * transparent areas, because it *is* the color space rather than an overlay.
 */
function writeLumaTintTransform(sink: ImageObjectSink, components: 3 | 4): number {
  const objNum = sink.allocate();
  const dict = new PdfDict()
    .set("FunctionType", "4")
    .set("Domain", components === 3 ? "[0 1 0 1 0 1]" : "[0 1 0 1 0 1 0 1]")
    .set("Range", "[0 1]");
  // RGB — stack on entry is R G B (B on top):
  //   0.114 mul  -> R G 0.114B
  //   exch       -> R 0.114B G
  //   0.587 mul  -> R 0.114B 0.587G
  //   add        -> R (0.114B + 0.587G)
  //   exch       -> (0.114B + 0.587G) R
  //   0.299 mul  -> (0.114B + 0.587G) 0.299R
  //   add        -> luma
  // CMYK — stack on entry is C M Y K: the ink coverage is K plus the luma-weighted CMY, and
  // the grey is what is left of the paper, `1 - coverage`, floored at 0.
  const program =
    components === 3
      ? "{ 0.114 mul exch 0.587 mul add exch 0.299 mul add }"
      : "{ exch 0.114 mul add exch 0.587 mul add exch 0.299 mul add 1 exch sub 0 max }";
  sink.addStream(objNum, dict, new TextEncoder().encode(program));
  return objNum;
}

const JPEG_COLOR_SPACES = { 1: "/DeviceGray", 3: "/DeviceRGB", 4: "/DeviceCMYK" } as const;

/**
 * Write a JPEG image using DCTDecode (raw JPEG data embedded directly).
 *
 * The colour space follows the frame header's component count — a greyscale or CMYK JPEG
 * declared `/DeviceRGB` does not decode. An Adobe four-component JPEG stores inverted
 * samples, which `/Decode [1 0 1 0 1 0 1 0]` undoes, as Acrobat expects.
 */
function writeJpegImageXObject(sink: ImageObjectSink, data: Uint8Array, grayscale = false): number {
  const info = readJpegInfo(data);
  if (info === undefined) {
    throw new PdfArgumentError("data", "Image: data is not a JPEG image this writer can embed");
  }
  // Grayscale without a JPEG decoder: keep the DCTDecode samples and reinterpret
  // the components through a `/DeviceN` space whose tint transform is the
  // luma formula. Component count still matches the JPEG, as PDF requires.
  const tintRef =
    grayscale && info.components !== 1 ? writeLumaTintTransform(sink, info.components) : undefined;
  const objNum = sink.allocate();
  const colorants = info.components === 4 ? "[/C1 /C2 /C3 /C4]" : "[/C1 /C2 /C3]";
  const dict = new PdfDict()
    .set("Type", "/XObject")
    .set("Subtype", "/Image")
    .set("Width", pdfNumber(info.width))
    .set("Height", pdfNumber(info.height))
    .set(
      "ColorSpace",
      tintRef === undefined
        ? JPEG_COLOR_SPACES[info.components]
        : `[/DeviceN ${colorants} /DeviceGray ${pdfRef(tintRef)}]`
    )
    .set("BitsPerComponent", "8")
    .set("Filter", "/DCTDecode");
  if (info.components === 4 && info.adobe) {
    dict.set("Decode", "[1 0 1 0 1 0 1 0]");
  }
  sink.addStream(objNum, dict, data);
  return objNum;
}

/** Rec. 601 luma of an 8-bit RGB triple. */
function lumaByte(r: number, g: number, b: number): number {
  return Math.round(0.299 * r + 0.587 * g + 0.114 * b);
}

/**
 * Write a PNG image: decode to raw RGB, create SMask for alpha if needed.
 */
function writePngImageXObject(sink: ImageObjectSink, data: Uint8Array, grayscale = false): number {
  let png: ReturnType<typeof decodePng>;
  try {
    png = decodePng(data);
  } catch (err) {
    throw new PdfRenderError(
      `Image: the PNG could not be decoded — ${err instanceof Error ? err.message : String(err)}`,
      { cause: err }
    );
  }
  const objNum = sink.allocate();

  // PNG is already decoded to raw samples here, so grayscale is a real pixel
  // conversion: collapse RGB to one luma component and switch to DeviceGray.
  // The alpha channel lives in a separate SMask and is untouched, so transparent
  // regions stay transparent.
  let pixels = png.pixels;
  if (grayscale) {
    const gray = new Uint8Array(png.width * png.height);
    for (let i = 0, p = 0; i < gray.length; i++, p += 3) {
      gray[i] = lumaByte(pixels[p], pixels[p + 1], pixels[p + 2]);
    }
    pixels = gray;
  }

  const dict = new PdfDict()
    .set("Type", "/XObject")
    .set("Subtype", "/Image")
    .set("Width", pdfNumber(png.width))
    .set("Height", pdfNumber(png.height))
    .set("ColorSpace", grayscale ? "/DeviceGray" : "/DeviceRGB")
    .set("BitsPerComponent", pdfNumber(png.bitsPerComponent));

  if (png.alpha) {
    const smaskObjNum = sink.allocate();
    const smaskDict = new PdfDict()
      .set("Type", "/XObject")
      .set("Subtype", "/Image")
      .set("Width", pdfNumber(png.width))
      .set("Height", pdfNumber(png.height))
      .set("ColorSpace", "/DeviceGray")
      .set("BitsPerComponent", "8");
    sink.addStream(smaskObjNum, smaskDict, png.alpha);
    dict.set("SMask", pdfRef(smaskObjNum));
  }

  sink.addStream(objNum, dict, pixels);
  return objNum;
}
