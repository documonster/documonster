/**
 * DOCX Font Obfuscation (ODTTF format)
 *
 * Per ECMA-376 Part 3, Section 4, embedded TrueType fonts may be obfuscated by
 * XOR-ing the first 32 bytes of the font file with a 16-byte GUID repeated twice.
 *
 * The GUID in the w:fontKey attribute is in the form "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"
 * but bytes are applied in a specific order derived from the GUID.
 */

/** Parse a GUID string "{8-4-4-4-12}" into 16 bytes in obfuscation-apply order. */
function parseGuid(guid: string): Uint8Array {
  // Strip braces
  const hex = guid.replace(/[{}]/g, "").replace(/-/g, "");
  if (hex.length !== 32) {
    throw new Error(`Invalid GUID: ${guid}`);
  }

  const bytes = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    bytes[i] = parseInt(hex.substring(i * 2, i * 2 + 2), 16);
  }

  // Per spec: apply bytes in reverse order within each GUID segment.
  // GUID layout: 4-2-2-2-6 bytes. Segments 0-2 are reversed, segments 3-4 stay.
  const reordered = new Uint8Array(16);
  // Segment 1 (bytes 0-3, reversed)
  reordered[0] = bytes[3];
  reordered[1] = bytes[2];
  reordered[2] = bytes[1];
  reordered[3] = bytes[0];
  // Segment 2 (bytes 4-5, reversed)
  reordered[4] = bytes[5];
  reordered[5] = bytes[4];
  // Segment 3 (bytes 6-7, reversed)
  reordered[6] = bytes[7];
  reordered[7] = bytes[6];
  // Segments 4-5 (bytes 8-15, kept as-is)
  for (let i = 8; i < 16; i++) {
    reordered[i] = bytes[i];
  }
  return reordered;
}

/**
 * Deobfuscate ODTTF font data to raw TTF/OTF.
 *
 * @param data - The obfuscated font data (first 32 bytes are obfuscated).
 * @param fontKey - The GUID from w:fontKey attribute.
 * @returns The deobfuscated font data (new buffer).
 */
export function deobfuscateFont(data: Uint8Array, fontKey: string): Uint8Array {
  const keyBytes = parseGuid(fontKey);
  const result = new Uint8Array(data.length);
  result.set(data);
  // XOR first 32 bytes with the GUID bytes twice (16 + 16)
  for (let i = 0; i < 32 && i < data.length; i++) {
    result[i] = data[i] ^ keyBytes[i % 16];
  }
  return result;
}

/**
 * Obfuscate a raw TTF/OTF font to ODTTF format.
 *
 * @param data - The raw font data.
 * @param fontKey - The GUID to use for obfuscation (also stored in w:fontKey).
 * @returns The obfuscated font data (new buffer).
 */
export function obfuscateFont(data: Uint8Array, fontKey: string): Uint8Array {
  // Obfuscation is symmetric with deobfuscation (XOR)
  return deobfuscateFont(data, fontKey);
}

/**
 * Generate a new random GUID suitable for w:fontKey.
 *
 * @returns A GUID string in the form "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"
 */
export function generateFontKey(): string {
  // Use crypto.getRandomValues (available in both Node and browser)
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);

  // RFC 4122 v4 UUID: set version and variant bits
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;

  const hex = Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
  return `{${hex.substring(0, 8)}-${hex.substring(8, 12)}-${hex.substring(12, 16)}-${hex.substring(16, 20)}-${hex.substring(20, 32)}}`;
}
