/**
 * Tests for the PNG decoder.
 */
import { describe, it, expect } from "vitest";
import { decodePng } from "@pdf/render/png-decoder";

/**
 * Build a minimal valid PNG file programmatically.
 * Creates a 2x2 image with known pixel values.
 */
function buildTestPng(
  colorType: number,
  pixelData: Uint8Array,
  palette?: Uint8Array,
  trns?: Uint8Array
): Uint8Array {
  const width = 2;
  const height = 2;
  const chunks: Uint8Array[] = [];

  // PNG signature
  chunks.push(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));

  // IHDR chunk
  const ihdr = new Uint8Array(13);
  const ihdrView = new DataView(ihdr.buffer);
  ihdrView.setUint32(0, width, false);
  ihdrView.setUint32(4, height, false);
  ihdr[8] = 8; // bit depth
  ihdr[9] = colorType;
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace (none)
  chunks.push(makeChunk("IHDR", ihdr));

  // PLTE chunk (if palette)
  if (palette) {
    chunks.push(makeChunk("PLTE", palette));
  }

  // tRNS chunk (if transparency)
  if (trns) {
    chunks.push(makeChunk("tRNS", trns));
  }

  // IDAT chunk — compress pixel data with zlib wrapper
  // Add filter byte (0 = None) before each row
  const channels = getChannels(colorType);
  const rowLen = width * channels;
  const filtered = new Uint8Array(height * (1 + rowLen));
  for (let y = 0; y < height; y++) {
    filtered[y * (1 + rowLen)] = 0; // filter type: None
    filtered.set(pixelData.subarray(y * rowLen, (y + 1) * rowLen), y * (1 + rowLen) + 1);
  }
  const compressed = zlibWrap(deflateRaw(filtered));
  chunks.push(makeChunk("IDAT", compressed));

  // IEND chunk
  chunks.push(makeChunk("IEND", new Uint8Array(0)));

  return concatAll(chunks);
}

function getChannels(colorType: number): number {
  switch (colorType) {
    case 0:
      return 1;
    case 2:
      return 3;
    case 3:
      return 1;
    case 4:
      return 2;
    case 6:
      return 4;
    default:
      return 3;
  }
}

function makeChunk(type: string, data: Uint8Array): Uint8Array {
  const chunk = new Uint8Array(4 + 4 + data.length + 4);
  const view = new DataView(chunk.buffer);
  view.setUint32(0, data.length, false);
  chunk[4] = type.charCodeAt(0);
  chunk[5] = type.charCodeAt(1);
  chunk[6] = type.charCodeAt(2);
  chunk[7] = type.charCodeAt(3);
  chunk.set(data, 8);
  // CRC placeholder (PNG readers in the decoder don't verify CRC)
  view.setUint32(8 + data.length, 0, false);
  return chunk;
}

/**
 * Minimal raw deflate using store blocks (no compression).
 * Each block: BFINAL(1) BTYPE=00(2) LEN(16) NLEN(16) data
 */
function deflateRaw(data: Uint8Array): Uint8Array {
  // Use a single stored block (max 65535 bytes)
  const len = data.length;
  if (len > 65535) {
    throw new Error("Test data too large for store-only deflate");
  }
  const result = new Uint8Array(5 + len);
  result[0] = 0x01; // BFINAL=1, BTYPE=00 (stored)
  result[1] = len & 0xff;
  result[2] = (len >> 8) & 0xff;
  result[3] = ~len & 0xff;
  result[4] = (~len >> 8) & 0xff;
  result.set(data, 5);
  return result;
}

function zlibWrap(deflated: Uint8Array): Uint8Array {
  // Zlib header: CMF=0x78 FLG=0x01 (deflate, window=32K, check bits)
  const result = new Uint8Array(2 + deflated.length + 4);
  result[0] = 0x78;
  result[1] = 0x01;
  result.set(deflated, 2);
  // Adler-32 placeholder
  return result;
}

function concatAll(arrays: Uint8Array[]): Uint8Array {
  let total = 0;
  for (const a of arrays) {
    total += a.length;
  }
  const result = new Uint8Array(total);
  let offset = 0;
  for (const a of arrays) {
    result.set(a, offset);
    offset += a.length;
  }
  return result;
}

// =============================================================================
// Tests
// =============================================================================

describe("PNG Decoder", () => {
  it("should decode RGB (colorType=2) PNG", () => {
    // 2x2 RGB image: red, green, blue, white
    const pixels = new Uint8Array([
      255,
      0,
      0,
      0,
      255,
      0, // row 0: red, green
      0,
      0,
      255,
      255,
      255,
      255 // row 1: blue, white
    ]);
    const png = buildTestPng(2, pixels);
    const decoded = decodePng(png);

    expect(decoded.width).toBe(2);
    expect(decoded.height).toBe(2);
    expect(decoded.alpha).toBeNull();
    expect(decoded.pixels[0]).toBe(255); // R of red
    expect(decoded.pixels[1]).toBe(0); // G of red
    expect(decoded.pixels[2]).toBe(0); // B of red
    expect(decoded.pixels[3]).toBe(0); // R of green
    expect(decoded.pixels[4]).toBe(255); // G of green
    expect(decoded.pixels[9]).toBe(255); // R of white
    expect(decoded.pixels[10]).toBe(255); // G of white
    expect(decoded.pixels[11]).toBe(255); // B of white
  });

  it("should decode RGBA (colorType=6) PNG and extract alpha", () => {
    // 2x2 RGBA: red opaque, green half-transparent, blue transparent, white opaque
    const pixels = new Uint8Array([
      255,
      0,
      0,
      255,
      0,
      255,
      0,
      128, // row 0
      0,
      0,
      255,
      0,
      255,
      255,
      255,
      255 // row 1
    ]);
    const png = buildTestPng(6, pixels);
    const decoded = decodePng(png);

    expect(decoded.width).toBe(2);
    expect(decoded.height).toBe(2);
    expect(decoded.alpha).not.toBeNull();
    // RGB values
    expect(decoded.pixels[0]).toBe(255);
    expect(decoded.pixels[3]).toBe(0);
    // Alpha values
    expect(decoded.alpha![0]).toBe(255); // red: opaque
    expect(decoded.alpha![1]).toBe(128); // green: half
    expect(decoded.alpha![2]).toBe(0); // blue: transparent
    expect(decoded.alpha![3]).toBe(255); // white: opaque
  });

  it("should decode RGBA PNG with fully opaque alpha as null alpha", () => {
    const pixels = new Uint8Array([
      255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255, 255, 255, 255, 255
    ]);
    const png = buildTestPng(6, pixels);
    const decoded = decodePng(png);

    expect(decoded.alpha).toBeNull(); // all alpha = 255, should be null
  });

  it("should decode Grayscale (colorType=0) PNG", () => {
    const pixels = new Uint8Array([0, 128, 255, 64]); // 2x2 grayscale
    const png = buildTestPng(0, pixels);
    const decoded = decodePng(png);

    expect(decoded.width).toBe(2);
    expect(decoded.height).toBe(2);
    expect(decoded.alpha).toBeNull();
    // Grayscale expanded to RGB
    expect(decoded.pixels[0]).toBe(0);
    expect(decoded.pixels[1]).toBe(0);
    expect(decoded.pixels[2]).toBe(0);
    expect(decoded.pixels[3]).toBe(128);
    expect(decoded.pixels[4]).toBe(128);
    expect(decoded.pixels[5]).toBe(128);
  });

  it("should decode Grayscale+Alpha (colorType=4) PNG", () => {
    const pixels = new Uint8Array([
      200,
      255,
      100,
      128, // row 0: gray 200 opaque, gray 100 half
      50,
      0,
      255,
      255 // row 1: gray 50 transparent, white opaque
    ]);
    const png = buildTestPng(4, pixels);
    const decoded = decodePng(png);

    expect(decoded.alpha).not.toBeNull();
    expect(decoded.pixels[0]).toBe(200);
    expect(decoded.alpha![0]).toBe(255);
    expect(decoded.alpha![1]).toBe(128);
    expect(decoded.alpha![2]).toBe(0);
  });

  it("should decode Palette (colorType=3) PNG", () => {
    // Palette: index 0=red, index 1=green, index 2=blue
    const palette = new Uint8Array([255, 0, 0, 0, 255, 0, 0, 0, 255]);
    const pixels = new Uint8Array([0, 1, 2, 0]); // 2x2 palette indices
    const png = buildTestPng(3, pixels, palette);
    const decoded = decodePng(png);

    expect(decoded.pixels[0]).toBe(255); // index 0 → red
    expect(decoded.pixels[3]).toBe(0); // index 1 → green R
    expect(decoded.pixels[4]).toBe(255); // index 1 → green G
    expect(decoded.pixels[6]).toBe(0); // index 2 → blue R
    expect(decoded.pixels[8]).toBe(255); // index 2 → blue B
  });

  it("should reject invalid PNG signature", () => {
    expect(() => decodePng(new Uint8Array([0, 0, 0, 0]))).toThrow("Invalid PNG signature");
  });

  it("should reject interlaced PNG", () => {
    // Build a PNG with interlace=1 in IHDR
    const ihdr = new Uint8Array(13);
    const v = new DataView(ihdr.buffer);
    v.setUint32(0, 1, false);
    v.setUint32(4, 1, false);
    ihdr[8] = 8;
    ihdr[9] = 2;
    ihdr[12] = 1; // interlace = 1
    const chunks = [
      new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      makeChunk("IHDR", ihdr)
    ];
    expect(() => decodePng(concatAll(chunks))).toThrow("Interlaced");
  });

  it("should decode Palette with tRNS transparency", () => {
    // Palette: idx 0=red, idx 1=green, idx 2=blue
    // tRNS: idx 0=opaque(255), idx 1=half(128), idx 2=transparent(0)
    const palette = new Uint8Array([255, 0, 0, 0, 255, 0, 0, 0, 255]);
    const trns = new Uint8Array([255, 128, 0]);
    const pixels = new Uint8Array([0, 1, 2, 0]); // 2x2
    const png = buildTestPng(3, pixels, palette, trns);
    const decoded = decodePng(png);

    expect(decoded.alpha).not.toBeNull();
    expect(decoded.alpha![0]).toBe(255); // idx 0 → opaque
    expect(decoded.alpha![1]).toBe(128); // idx 1 → half
    expect(decoded.alpha![2]).toBe(0); // idx 2 → transparent
    expect(decoded.alpha![3]).toBe(255); // idx 0 → opaque
  });

  it("should decode PNG with Sub filter (filter type 1)", () => {
    // 2x2 RGB with filter type 1 (Sub) applied manually
    // Row 0 pixels: [255, 0, 0,  0, 255, 0] → after Sub filter: [255, 0, 0,  1, 255, 0]
    //   (second pixel's R: 0 - 255 = 1 mod 256... actually Sub stores diff from left)
    // Let's just use a known input where Sub(x) = Raw(x) - Raw(x-bpp)
    // For simplicity: row pixels [10, 20, 30, 40, 50, 60]
    // Sub-filtered: [10, 20, 30, 30, 30, 30] (each byte - byte 3 positions left)
    const width = 2;
    const height = 1;
    const channels = 3;
    const rowLen = width * channels;

    // Build filtered data with Sub filter
    const filtered = new Uint8Array(1 + rowLen);
    filtered[0] = 1; // Sub filter
    filtered[1] = 10;
    filtered[2] = 20;
    filtered[3] = 30; // first pixel as-is
    filtered[4] = 30;
    filtered[5] = 30;
    filtered[6] = 30; // diff from left

    const compressed = zlibWrap(deflateRaw(filtered));
    const chunks = [
      new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      makeChunk("IHDR", buildIhdr(width, height, 8, 2)),
      makeChunk("IDAT", compressed),
      makeChunk("IEND", new Uint8Array(0))
    ];
    const decoded = decodePng(concatAll(chunks));

    expect(decoded.width).toBe(2);
    expect(decoded.height).toBe(1);
    // Reconstructed: [10, 20, 30, 10+30=40, 20+30=50, 30+30=60]
    expect(decoded.pixels[0]).toBe(10);
    expect(decoded.pixels[1]).toBe(20);
    expect(decoded.pixels[2]).toBe(30);
    expect(decoded.pixels[3]).toBe(40);
    expect(decoded.pixels[4]).toBe(50);
    expect(decoded.pixels[5]).toBe(60);
  });

  it("should decode PNG with Up filter (filter type 2)", () => {
    const width = 2;
    const height = 2;

    // Row 0: None filter, raw [10, 20, 30, 40, 50, 60]
    // Row 1: Up filter, diff [5, 5, 5, 5, 5, 5] → reconstructed [15, 25, 35, 45, 55, 65]
    const filtered = new Uint8Array([
      0,
      10,
      20,
      30,
      40,
      50,
      60, // row 0: filter=None
      2,
      5,
      5,
      5,
      5,
      5,
      5 // row 1: filter=Up
    ]);

    const compressed = zlibWrap(deflateRaw(filtered));
    const chunks = [
      new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      makeChunk("IHDR", buildIhdr(width, height, 8, 2)),
      makeChunk("IDAT", compressed),
      makeChunk("IEND", new Uint8Array(0))
    ];
    const decoded = decodePng(concatAll(chunks));

    // Row 1 reconstructed: each byte = diff + above
    expect(decoded.pixels[6]).toBe(15); // 10 + 5
    expect(decoded.pixels[7]).toBe(25); // 20 + 5
    expect(decoded.pixels[8]).toBe(35); // 30 + 5
    expect(decoded.pixels[9]).toBe(45); // 40 + 5
    expect(decoded.pixels[10]).toBe(55); // 50 + 5
    expect(decoded.pixels[11]).toBe(65); // 60 + 5
  });

  it("should reject non-8-bit PNG with clear error", () => {
    expect(() => decodePng(buildPngWithBitDepth(4, 0))).toThrow("Unsupported PNG bit depth: 4");
    expect(() => decodePng(buildPngWithBitDepth(16, 2))).toThrow("Unsupported PNG bit depth: 16");
    expect(() => decodePng(buildPngWithBitDepth(1, 0))).toThrow("Unsupported PNG bit depth: 1");
  });

  it("should handle grayscale tRNS transparency", () => {
    // 2x2 grayscale, pixel values: [100, 200, 100, 50]
    // tRNS: gray value 100 should be transparent
    const pixelData = new Uint8Array([100, 200, 100, 50]);
    // tRNS for grayscale: 2 bytes (16-bit value, high byte 0, low byte = sample value)
    const trns = new Uint8Array([0, 100]);
    const png = buildTrnsPng(0, pixelData, trns);
    const decoded = decodePng(png);

    expect(decoded.alpha).not.toBeNull();
    // Pixels with value 100 should be transparent (alpha=0)
    expect(decoded.alpha![0]).toBe(0); // gray=100 matches
    expect(decoded.alpha![1]).toBe(255); // gray=200 no match
    expect(decoded.alpha![2]).toBe(0); // gray=100 matches
    expect(decoded.alpha![3]).toBe(255); // gray=50 no match
  });

  it("should handle truecolor tRNS transparency", () => {
    // 2x2 RGB, transparent color is (255, 0, 0) = red
    const pixelData = new Uint8Array([
      255,
      0,
      0,
      0,
      255,
      0, // row 0: red, green
      0,
      0,
      255,
      255,
      0,
      0 // row 1: blue, red
    ]);
    // tRNS for truecolor: 6 bytes (2 bytes each for R, G, B in 16-bit)
    const trns = new Uint8Array([0, 255, 0, 0, 0, 0]); // R=255, G=0, B=0
    const png = buildTrnsPng(2, pixelData, trns);
    const decoded = decodePng(png);

    expect(decoded.alpha).not.toBeNull();
    expect(decoded.alpha![0]).toBe(0); // red → transparent
    expect(decoded.alpha![1]).toBe(255); // green → opaque
    expect(decoded.alpha![2]).toBe(255); // blue → opaque
    expect(decoded.alpha![3]).toBe(0); // red → transparent
  });
});

function buildIhdr(width: number, height: number, bitDepth: number, colorType: number): Uint8Array {
  const ihdr = new Uint8Array(13);
  const v = new DataView(ihdr.buffer);
  v.setUint32(0, width, false);
  v.setUint32(4, height, false);
  ihdr[8] = bitDepth;
  ihdr[9] = colorType;
  return ihdr;
}

/**
 * Build a PNG with a custom bit depth (for rejection testing).
 */
function buildPngWithBitDepth(bitDepth: number, colorType: number): Uint8Array {
  const chunks: Uint8Array[] = [];
  chunks.push(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));

  const ihdr = buildIhdr(1, 1, bitDepth, colorType);
  chunks.push(makeChunk("IHDR", ihdr));

  // Minimal IDAT: 1 pixel, filter=None
  const rawScanline =
    colorType === 2 ? new Uint8Array([0, 128, 128, 128]) : new Uint8Array([0, 128]);
  const compressed = deflateRaw(rawScanline);
  chunks.push(makeChunk("IDAT", compressed));
  chunks.push(makeChunk("IEND", new Uint8Array(0)));

  let totalLen = 0;
  for (const c of chunks) {
    totalLen += c.length;
  }
  const result = new Uint8Array(totalLen);
  let off = 0;
  for (const c of chunks) {
    result.set(c, off);
    off += c.length;
  }
  return result;
}

/**
 * Build a 2x2 PNG with tRNS chunk for grayscale or truecolor.
 */
function buildTrnsPng(colorType: number, pixelData: Uint8Array, trns: Uint8Array): Uint8Array {
  const width = 2;
  const height = 2;
  const chunks: Uint8Array[] = [];

  chunks.push(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));

  const ihdr = buildIhdr(width, height, 8, colorType);
  chunks.push(makeChunk("IHDR", ihdr));
  chunks.push(makeChunk("tRNS", trns));

  // Add filter byte (None) before each scanline
  const channels = colorType === 0 ? 1 : colorType === 2 ? 3 : 4;
  const scanlineLen = width * channels;
  const filtered = new Uint8Array(height * (1 + scanlineLen));
  for (let y = 0; y < height; y++) {
    filtered[y * (1 + scanlineLen)] = 0; // filter: None
    filtered.set(
      pixelData.subarray(y * scanlineLen, y * scanlineLen + scanlineLen),
      y * (1 + scanlineLen) + 1
    );
  }
  const compressed = deflateRaw(filtered);
  chunks.push(makeChunk("IDAT", compressed));
  chunks.push(makeChunk("IEND", new Uint8Array(0)));

  let totalLen = 0;
  for (const c of chunks) {
    totalLen += c.length;
  }
  const result = new Uint8Array(totalLen);
  let off = 0;
  for (const c of chunks) {
    result.set(c, off);
    off += c.length;
  }
  return result;
}
