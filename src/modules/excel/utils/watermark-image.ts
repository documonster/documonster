/**
 * Zero-dependency text-to-PNG watermark image generator.
 *
 * Renders text into a semi-transparent PNG suitable for use as an Excel watermark.
 * Uses a built-in bitmap font for ASCII characters — no Canvas or external fonts required.
 * PNG data is deflate-compressed using the archive module's built-in compressor.
 *
 * @example
 * ```typescript
 * const png = createTextWatermarkImage("CONFIDENTIAL", {
 *   fontSize: 48,
 *   color: { r: 128, g: 128, b: 128 },
 *   opacity: 40,
 *   rotation: -45
 * });
 * const imgId = workbook.addImage({ buffer: png, extension: "png" });
 * worksheet.addWatermark({ imageId: imgId });
 * ```
 */

// =============================================================================
// Public API
// =============================================================================

import { deflateRawCompressed } from "@archive/compression/deflate-fallback";

/**
 * Options for text watermark image generation.
 */
export interface TextWatermarkImageOptions {
  /**
   * Approximate font size in pixels (glyph height).
   * The built-in bitmap font is 8px tall; values larger than 8 are achieved by
   * integer scaling. e.g. fontSize 48 → 6x scale.
   * @default 48
   */
  fontSize?: number;
  /**
   * Text color as RGB (0-255 each).
   * @default { r: 128, g: 128, b: 128 }
   */
  color?: { r: number; g: number; b: number };
  /**
   * Opacity as a percentage (0 = fully transparent, 100 = fully opaque).
   *
   * Note: this is a **0–100 percentage** used when rendering the PNG image pixels.
   * It is different from `WatermarkOptions.opacity` (which is 0–1) used by
   * `worksheet.addWatermark()` for DrawingML `alphaModFix`.
   *
   * @default 40
   */
  opacity?: number;
  /**
   * Rotation in degrees (positive = counter-clockwise).
   * @default -45
   */
  rotation?: number;
  /**
   * Padding in pixels around the text (before rotation).
   * @default 20
   */
  padding?: number;
}

/**
 * Generate a PNG image containing watermark text.
 *
 * The image has an alpha channel so the watermark is semi-transparent.
 * Works in both Node.js and browsers with zero dependencies.
 */
export function createTextWatermarkImage(
  text: string,
  options?: TextWatermarkImageOptions
): Uint8Array {
  const fontSize = options?.fontSize ?? 48;
  const color = options?.color ?? { r: 128, g: 128, b: 128 };
  const opacity = Math.max(0, Math.min(100, options?.opacity ?? 40));
  const rotation = options?.rotation ?? -45;
  const padding = options?.padding ?? 20;

  // Scale factor: built-in font is 8px tall
  const scale = Math.max(1, Math.round(fontSize / GLYPH_HEIGHT));

  // Render text to unrotated bitmap
  const { width: textW, height: textH, pixels: textPixels } = renderTextBitmap(text, scale);

  // Add padding
  const paddedW = textW + padding * 2;
  const paddedH = textH + padding * 2;
  const paddedPixels = new Uint8Array(paddedW * paddedH);
  for (let y = 0; y < textH; y++) {
    for (let x = 0; x < textW; x++) {
      paddedPixels[(y + padding) * paddedW + (x + padding)] = textPixels[y * textW + x];
    }
  }

  // Rotate
  const {
    width: rotW,
    height: rotH,
    pixels: rotPixels
  } = rotateBitmap(paddedPixels, paddedW, paddedH, rotation);

  // Convert to RGBA PNG
  const alpha = Math.round((opacity / 100) * 255);
  const rgba = new Uint8Array(rotW * rotH * 4);
  for (let i = 0; i < rotW * rotH; i++) {
    const a = rotPixels[i];
    if (a > 0) {
      rgba[i * 4] = color.r;
      rgba[i * 4 + 1] = color.g;
      rgba[i * 4 + 2] = color.b;
      rgba[i * 4 + 3] = Math.round((a / 255) * alpha);
    }
    // else fully transparent (already 0)
  }

  return encodePng(rgba, rotW, rotH);
}

// =============================================================================
// Bitmap Font — 8px tall monospace ASCII (CP437-style, printable range 32-126)
// =============================================================================

const GLYPH_WIDTH = 6;
const GLYPH_HEIGHT = 8;

/**
 * Compact glyph data: each character is 8 bytes (one byte per row, 6 bits used).
 * Bit 5 = leftmost pixel, bit 0 = rightmost pixel.
 */
const FONT_DATA: Record<number, number[]> = {
  // space
  32: [0, 0, 0, 0, 0, 0, 0, 0],
  // !
  33: [0x04, 0x04, 0x04, 0x04, 0x04, 0x00, 0x04, 0x00],
  // "
  34: [0x0a, 0x0a, 0x0a, 0x00, 0x00, 0x00, 0x00, 0x00],
  // #
  35: [0x0a, 0x0a, 0x1f, 0x0a, 0x1f, 0x0a, 0x0a, 0x00],
  // $
  36: [0x04, 0x0f, 0x14, 0x0e, 0x05, 0x1e, 0x04, 0x00],
  // %
  37: [0x18, 0x19, 0x02, 0x04, 0x08, 0x13, 0x03, 0x00],
  // &
  38: [0x0c, 0x12, 0x14, 0x08, 0x15, 0x12, 0x0d, 0x00],
  // '
  39: [0x04, 0x04, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00],
  // (
  40: [0x02, 0x04, 0x08, 0x08, 0x08, 0x04, 0x02, 0x00],
  // )
  41: [0x08, 0x04, 0x02, 0x02, 0x02, 0x04, 0x08, 0x00],
  // *
  42: [0x00, 0x04, 0x15, 0x0e, 0x15, 0x04, 0x00, 0x00],
  // +
  43: [0x00, 0x04, 0x04, 0x1f, 0x04, 0x04, 0x00, 0x00],
  // ,
  44: [0x00, 0x00, 0x00, 0x00, 0x00, 0x04, 0x04, 0x08],
  // -
  45: [0x00, 0x00, 0x00, 0x1f, 0x00, 0x00, 0x00, 0x00],
  // .
  46: [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x04, 0x00],
  // /
  47: [0x00, 0x01, 0x02, 0x04, 0x08, 0x10, 0x00, 0x00],
  // 0-9
  48: [0x0e, 0x11, 0x13, 0x15, 0x19, 0x11, 0x0e, 0x00],
  49: [0x04, 0x0c, 0x04, 0x04, 0x04, 0x04, 0x0e, 0x00],
  50: [0x0e, 0x11, 0x01, 0x02, 0x04, 0x08, 0x1f, 0x00],
  51: [0x1f, 0x02, 0x04, 0x02, 0x01, 0x11, 0x0e, 0x00],
  52: [0x02, 0x06, 0x0a, 0x12, 0x1f, 0x02, 0x02, 0x00],
  53: [0x1f, 0x10, 0x1e, 0x01, 0x01, 0x11, 0x0e, 0x00],
  54: [0x06, 0x08, 0x10, 0x1e, 0x11, 0x11, 0x0e, 0x00],
  55: [0x1f, 0x01, 0x02, 0x04, 0x08, 0x08, 0x08, 0x00],
  56: [0x0e, 0x11, 0x11, 0x0e, 0x11, 0x11, 0x0e, 0x00],
  57: [0x0e, 0x11, 0x11, 0x0f, 0x01, 0x02, 0x0c, 0x00],
  // :
  58: [0x00, 0x00, 0x04, 0x00, 0x00, 0x04, 0x00, 0x00],
  // ;
  59: [0x00, 0x00, 0x04, 0x00, 0x00, 0x04, 0x04, 0x08],
  // <
  60: [0x02, 0x04, 0x08, 0x10, 0x08, 0x04, 0x02, 0x00],
  // =
  61: [0x00, 0x00, 0x1f, 0x00, 0x1f, 0x00, 0x00, 0x00],
  // >
  62: [0x08, 0x04, 0x02, 0x01, 0x02, 0x04, 0x08, 0x00],
  // ?
  63: [0x0e, 0x11, 0x01, 0x02, 0x04, 0x00, 0x04, 0x00],
  // @
  64: [0x0e, 0x11, 0x17, 0x15, 0x17, 0x10, 0x0e, 0x00],
  // A-Z
  65: [0x0e, 0x11, 0x11, 0x1f, 0x11, 0x11, 0x11, 0x00],
  66: [0x1e, 0x11, 0x11, 0x1e, 0x11, 0x11, 0x1e, 0x00],
  67: [0x0e, 0x11, 0x10, 0x10, 0x10, 0x11, 0x0e, 0x00],
  68: [0x1c, 0x12, 0x11, 0x11, 0x11, 0x12, 0x1c, 0x00],
  69: [0x1f, 0x10, 0x10, 0x1e, 0x10, 0x10, 0x1f, 0x00],
  70: [0x1f, 0x10, 0x10, 0x1e, 0x10, 0x10, 0x10, 0x00],
  71: [0x0e, 0x11, 0x10, 0x17, 0x11, 0x11, 0x0f, 0x00],
  72: [0x11, 0x11, 0x11, 0x1f, 0x11, 0x11, 0x11, 0x00],
  73: [0x0e, 0x04, 0x04, 0x04, 0x04, 0x04, 0x0e, 0x00],
  74: [0x07, 0x02, 0x02, 0x02, 0x02, 0x12, 0x0c, 0x00],
  75: [0x11, 0x12, 0x14, 0x18, 0x14, 0x12, 0x11, 0x00],
  76: [0x10, 0x10, 0x10, 0x10, 0x10, 0x10, 0x1f, 0x00],
  77: [0x11, 0x1b, 0x15, 0x15, 0x11, 0x11, 0x11, 0x00],
  78: [0x11, 0x19, 0x15, 0x13, 0x11, 0x11, 0x11, 0x00],
  79: [0x0e, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0e, 0x00],
  80: [0x1e, 0x11, 0x11, 0x1e, 0x10, 0x10, 0x10, 0x00],
  81: [0x0e, 0x11, 0x11, 0x11, 0x15, 0x12, 0x0d, 0x00],
  82: [0x1e, 0x11, 0x11, 0x1e, 0x14, 0x12, 0x11, 0x00],
  83: [0x0f, 0x10, 0x10, 0x0e, 0x01, 0x01, 0x1e, 0x00],
  84: [0x1f, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x00],
  85: [0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0e, 0x00],
  86: [0x11, 0x11, 0x11, 0x11, 0x11, 0x0a, 0x04, 0x00],
  87: [0x11, 0x11, 0x11, 0x15, 0x15, 0x1b, 0x11, 0x00],
  88: [0x11, 0x11, 0x0a, 0x04, 0x0a, 0x11, 0x11, 0x00],
  89: [0x11, 0x11, 0x0a, 0x04, 0x04, 0x04, 0x04, 0x00],
  90: [0x1f, 0x01, 0x02, 0x04, 0x08, 0x10, 0x1f, 0x00],
  // [ \ ]
  91: [0x0e, 0x08, 0x08, 0x08, 0x08, 0x08, 0x0e, 0x00],
  92: [0x00, 0x10, 0x08, 0x04, 0x02, 0x01, 0x00, 0x00],
  93: [0x0e, 0x02, 0x02, 0x02, 0x02, 0x02, 0x0e, 0x00],
  // ^ _ `
  94: [0x04, 0x0a, 0x11, 0x00, 0x00, 0x00, 0x00, 0x00],
  95: [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x1f, 0x00],
  96: [0x08, 0x04, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00],
  // a-z
  97: [0x00, 0x00, 0x0e, 0x01, 0x0f, 0x11, 0x0f, 0x00],
  98: [0x10, 0x10, 0x16, 0x19, 0x11, 0x11, 0x1e, 0x00],
  99: [0x00, 0x00, 0x0e, 0x10, 0x10, 0x11, 0x0e, 0x00],
  100: [0x01, 0x01, 0x0d, 0x13, 0x11, 0x11, 0x0f, 0x00],
  101: [0x00, 0x00, 0x0e, 0x11, 0x1f, 0x10, 0x0e, 0x00],
  102: [0x06, 0x09, 0x08, 0x1c, 0x08, 0x08, 0x08, 0x00],
  103: [0x00, 0x00, 0x0f, 0x11, 0x0f, 0x01, 0x0e, 0x00],
  104: [0x10, 0x10, 0x16, 0x19, 0x11, 0x11, 0x11, 0x00],
  105: [0x04, 0x00, 0x0c, 0x04, 0x04, 0x04, 0x0e, 0x00],
  106: [0x02, 0x00, 0x06, 0x02, 0x02, 0x12, 0x0c, 0x00],
  107: [0x10, 0x10, 0x12, 0x14, 0x18, 0x14, 0x12, 0x00],
  108: [0x0c, 0x04, 0x04, 0x04, 0x04, 0x04, 0x0e, 0x00],
  109: [0x00, 0x00, 0x1a, 0x15, 0x15, 0x11, 0x11, 0x00],
  110: [0x00, 0x00, 0x16, 0x19, 0x11, 0x11, 0x11, 0x00],
  111: [0x00, 0x00, 0x0e, 0x11, 0x11, 0x11, 0x0e, 0x00],
  112: [0x00, 0x00, 0x1e, 0x11, 0x1e, 0x10, 0x10, 0x00],
  113: [0x00, 0x00, 0x0d, 0x13, 0x0f, 0x01, 0x01, 0x00],
  114: [0x00, 0x00, 0x16, 0x19, 0x10, 0x10, 0x10, 0x00],
  115: [0x00, 0x00, 0x0e, 0x10, 0x0e, 0x01, 0x1e, 0x00],
  116: [0x08, 0x08, 0x1c, 0x08, 0x08, 0x09, 0x06, 0x00],
  117: [0x00, 0x00, 0x11, 0x11, 0x11, 0x13, 0x0d, 0x00],
  118: [0x00, 0x00, 0x11, 0x11, 0x11, 0x0a, 0x04, 0x00],
  119: [0x00, 0x00, 0x11, 0x11, 0x15, 0x15, 0x0a, 0x00],
  120: [0x00, 0x00, 0x11, 0x0a, 0x04, 0x0a, 0x11, 0x00],
  121: [0x00, 0x00, 0x11, 0x11, 0x0f, 0x01, 0x0e, 0x00],
  122: [0x00, 0x00, 0x1f, 0x02, 0x04, 0x08, 0x1f, 0x00],
  // { | } ~
  123: [0x02, 0x04, 0x04, 0x08, 0x04, 0x04, 0x02, 0x00],
  124: [0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x00],
  125: [0x08, 0x04, 0x04, 0x02, 0x04, 0x04, 0x08, 0x00],
  126: [0x00, 0x00, 0x08, 0x15, 0x02, 0x00, 0x00, 0x00]
};

// =============================================================================
// Bitmap Rendering
// =============================================================================

/** Render text string to a grayscale bitmap (0 = transparent, 255 = opaque). */
function renderTextBitmap(
  text: string,
  scale: number
): { width: number; height: number; pixels: Uint8Array } {
  const charW = GLYPH_WIDTH * scale;
  const charH = GLYPH_HEIGHT * scale;
  const width = text.length * charW;
  const height = charH;
  const pixels = new Uint8Array(width * height);

  for (let ci = 0; ci < text.length; ci++) {
    const code = text.charCodeAt(ci);
    const glyph = FONT_DATA[code] ?? FONT_DATA[63]; // fallback to '?'
    const xOff = ci * charW;

    for (let row = 0; row < GLYPH_HEIGHT; row++) {
      const bits = glyph[row];
      for (let col = 0; col < GLYPH_WIDTH; col++) {
        if (bits & (1 << (GLYPH_WIDTH - 1 - col))) {
          // Fill scaled pixel block
          for (let sy = 0; sy < scale; sy++) {
            for (let sx = 0; sx < scale; sx++) {
              const px = xOff + col * scale + sx;
              const py = row * scale + sy;
              if (px < width && py < height) {
                pixels[py * width + px] = 255;
              }
            }
          }
        }
      }
    }
  }

  return { width, height, pixels };
}

/** Rotate a grayscale bitmap by the given angle in degrees. */
function rotateBitmap(
  pixels: Uint8Array,
  srcW: number,
  srcH: number,
  angleDeg: number
): { width: number; height: number; pixels: Uint8Array } {
  if (angleDeg === 0) {
    return { width: srcW, height: srcH, pixels };
  }

  const rad = (angleDeg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);

  // Compute bounding box of rotated rectangle
  const corners = [
    { x: 0, y: 0 },
    { x: srcW, y: 0 },
    { x: srcW, y: srcH },
    { x: 0, y: srcH }
  ];
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const c of corners) {
    const rx = c.x * cos - c.y * sin;
    const ry = c.x * sin + c.y * cos;
    minX = Math.min(minX, rx);
    minY = Math.min(minY, ry);
    maxX = Math.max(maxX, rx);
    maxY = Math.max(maxY, ry);
  }

  const dstW = Math.ceil(maxX - minX);
  const dstH = Math.ceil(maxY - minY);
  const dst = new Uint8Array(dstW * dstH);

  // Inverse rotation: for each dst pixel, find the source pixel
  const invCos = cos; // cos(-θ) = cos(θ)
  const invSin = -sin; // sin(-θ) = -sin(θ)

  for (let dy = 0; dy < dstH; dy++) {
    for (let dx = 0; dx < dstW; dx++) {
      // Map dst to world, then inverse-rotate to source
      const wx = dx + minX;
      const wy = dy + minY;
      const sx = Math.round(wx * invCos - wy * invSin);
      const sy = Math.round(wx * invSin + wy * invCos);

      if (sx >= 0 && sx < srcW && sy >= 0 && sy < srcH) {
        dst[dy * dstW + dx] = pixels[sy * srcW + sx];
      }
    }
  }

  return { width: dstW, height: dstH, pixels: dst };
}

// =============================================================================
// PNG Encoder (RGBA, deflate-compressed, with alpha)
// =============================================================================

/** Encode RGBA pixel data to a PNG file. */
function encodePng(rgba: Uint8Array, width: number, height: number): Uint8Array {
  // Build IDAT data: filter byte (0 = None) + raw RGBA for each row
  const rawRowSize = 1 + width * 4; // filter byte + pixels
  const rawData = new Uint8Array(rawRowSize * height);
  for (let y = 0; y < height; y++) {
    rawData[y * rawRowSize] = 0; // filter: None
    rawData.set(rgba.subarray(y * width * 4, (y + 1) * width * 4), y * rawRowSize + 1);
  }

  // Wrap in zlib stream with deflate compression
  const deflated = zlibCompress(rawData);

  // PNG signature
  const sig = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

  // IHDR chunk
  const ihdr = new Uint8Array(13);
  writeU32BE(ihdr, 0, width);
  writeU32BE(ihdr, 4, height);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type: RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace
  const ihdrChunk = pngChunk(0x49484452, ihdr);

  // IDAT chunk
  const idatChunk = pngChunk(0x49444154, deflated);

  // IEND chunk
  const iendChunk = pngChunk(0x49454e44, new Uint8Array(0));

  // Concatenate
  const result = new Uint8Array(
    sig.length + ihdrChunk.length + idatChunk.length + iendChunk.length
  );
  let offset = 0;
  result.set(sig, offset);
  offset += sig.length;
  result.set(ihdrChunk, offset);
  offset += ihdrChunk.length;
  result.set(idatChunk, offset);
  offset += idatChunk.length;
  result.set(iendChunk, offset);

  return result;
}

/** Build a PNG chunk: length(4) + type(4) + data + crc32(4). */
function pngChunk(type: number, data: Uint8Array): Uint8Array {
  const chunk = new Uint8Array(12 + data.length);
  writeU32BE(chunk, 0, data.length);
  writeU32BE(chunk, 4, type);
  chunk.set(data, 8);
  // CRC32 over type + data
  const crc = crc32(chunk.subarray(4, 8 + data.length));
  writeU32BE(chunk, 8 + data.length, crc);
  return chunk;
}

/** Write a 32-bit big-endian unsigned int. */
function writeU32BE(buf: Uint8Array, offset: number, value: number): void {
  buf[offset] = (value >>> 24) & 0xff;
  buf[offset + 1] = (value >>> 16) & 0xff;
  buf[offset + 2] = (value >>> 8) & 0xff;
  buf[offset + 3] = value & 0xff;
}

/** Wrap raw data in a zlib stream with deflate compression. */
function zlibCompress(data: Uint8Array): Uint8Array {
  // Zlib header: CMF=0x78, FLG=0x01 (deflate, no dict, check bits)
  const deflated = deflateRawCompressed(data, 6);
  const adler = adler32(data);

  const result = new Uint8Array(2 + deflated.length + 4);
  result[0] = 0x78;
  result[1] = 0x01;
  result.set(deflated, 2);
  writeU32BE(result, 2 + deflated.length, adler);
  return result;
}

/** Compute Adler-32 checksum. */
function adler32(data: Uint8Array): number {
  let a = 1;
  let b = 0;
  for (let i = 0; i < data.length; i++) {
    a = (a + data[i]) % 65521;
    b = (b + a) % 65521;
  }
  return (b << 16) | a;
}

/** CRC32 lookup table. */
const CRC_TABLE = /* @__PURE__ */ (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c;
  }
  return table;
})();

/** Compute CRC32 checksum. */
function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    crc = CRC_TABLE[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
