function toHex(byte: number): string {
  return byte.toString(16).padStart(2, "0");
}

function bytesToUuidV4(bytes: Uint8Array): string {
  // RFC 4122, version 4 UUID
  // Set version (4) and variant (10xx)
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;

  const b = bytes;
  return (
    toHex(b[0]) +
    toHex(b[1]) +
    toHex(b[2]) +
    toHex(b[3]) +
    "-" +
    toHex(b[4]) +
    toHex(b[5]) +
    "-" +
    toHex(b[6]) +
    toHex(b[7]) +
    "-" +
    toHex(b[8]) +
    toHex(b[9]) +
    "-" +
    toHex(b[10]) +
    toHex(b[11]) +
    toHex(b[12]) +
    toHex(b[13]) +
    toHex(b[14]) +
    toHex(b[15])
  );
}

/**
 * Generate a UUID v4.
 *
 * This exists to avoid a hard dependency on `crypto.randomUUID()` so that
 * older supported browsers can still generate UUIDs.
 *
 * - Prefers `crypto.randomUUID()` when available.
 * - Falls back to `crypto.getRandomValues()`.
 * - Last-resort fallback uses `Math.random()` (NOT cryptographically secure).
 */
export function uuidV4(): string {
  // `globalThis.crypto` is typed `Crypto` but may be absent on older runtimes,
  // so treat it as possibly-undefined rather than casting through `any`.
  const cryptoObj: Crypto | undefined = globalThis.crypto;

  if (cryptoObj?.randomUUID) {
    return cryptoObj.randomUUID();
  }

  const bytes = new Uint8Array(16);
  if (cryptoObj?.getRandomValues) {
    cryptoObj.getRandomValues(bytes);
    return bytesToUuidV4(bytes);
  }

  for (let i = 0; i < bytes.length; i++) {
    bytes[i] = Math.floor(Math.random() * 256);
  }
  return bytesToUuidV4(bytes);
}
