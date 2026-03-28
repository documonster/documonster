/**
 * PDF encryption support (Standard Security Handler, Revision 3).
 *
 * Implements RC4-128 encryption compatible with PDF 1.4.
 * Supports:
 * - User password (required to open the document)
 * - Owner password (grants full access)
 * - Permission flags (print, copy, modify, etc.)
 *
 * @see PDF Reference 1.7, §3.5 - Encryption
 */

// =============================================================================
// Types
// =============================================================================

/**
 * PDF encryption options.
 */
export interface PdfEncryptionOptions {
  /** User password (required to open the document). Empty string = no open password. */
  userPassword?: string;

  /** Owner password (grants full permissions). Required. */
  ownerPassword: string;

  /** Permissions to grant when opened with user password. */
  permissions?: Partial<PdfPermissions>;
}

/**
 * PDF document permissions (what a user-password holder can do).
 * All default to false.
 */
export interface PdfPermissions {
  /** Allow printing */
  print: boolean;
  /** Allow modifying content */
  modify: boolean;
  /** Allow copying text/images */
  copy: boolean;
  /** Allow adding/modifying annotations */
  annotate: boolean;
  /** Allow filling form fields */
  fillForms: boolean;
  /** Allow extracting content for accessibility */
  accessibility: boolean;
  /** Allow assembling (insert/rotate/delete pages) */
  assemble: boolean;
  /** Allow high-quality printing */
  printHighQuality: boolean;
}

/**
 * Encryption state used during PDF generation.
 */
export interface EncryptionState {
  /** Encryption key (variable length, up to 16 bytes) */
  encryptionKey: Uint8Array;
  /** O value (32 bytes) for the Encrypt dictionary */
  oValue: Uint8Array;
  /** U value (32 bytes) for the Encrypt dictionary */
  uValue: Uint8Array;
  /** Permissions integer (P value) */
  permissions: number;
  /** File identifier (16 bytes) */
  fileId: Uint8Array;
}

// =============================================================================
// Constants
// =============================================================================

/** PDF password padding string (32 bytes) per PDF spec §3.5.2 */
const PASSWORD_PADDING = new Uint8Array([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a
]);

// =============================================================================
// Public API
// =============================================================================

/**
 * Initialize encryption state from the given options.
 */
export function initEncryption(options: PdfEncryptionOptions): EncryptionState {
  const userPwd = options.userPassword ?? "";
  const ownerPwd = options.ownerPassword;
  const perms = computePermissions(options.permissions);
  const fileId = generateFileId();

  // Step 1: Compute O value
  const oValue = computeOValue(ownerPwd, userPwd);

  // Step 2: Compute encryption key
  const encryptionKey = computeEncryptionKey(userPwd, oValue, perms, fileId);

  // Step 3: Compute U value
  const uValue = computeUValue(encryptionKey, fileId);

  return { encryptionKey, oValue, uValue, permissions: perms, fileId };
}

/**
 * Encrypt a string or stream for a specific PDF object.
 * Per-object encryption key = MD5(encryptionKey + objectNumber + generation).
 */
export function encryptData(
  data: Uint8Array,
  objectNumber: number,
  generation: number,
  encryptionKey: Uint8Array
): Uint8Array {
  // Compute per-object key: MD5(encryptionKey + objNum(3LE) + genNum(2LE))
  const keyInput = new Uint8Array(encryptionKey.length + 5);
  keyInput.set(encryptionKey);
  keyInput[encryptionKey.length] = objectNumber & 0xff;
  keyInput[encryptionKey.length + 1] = (objectNumber >> 8) & 0xff;
  keyInput[encryptionKey.length + 2] = (objectNumber >> 16) & 0xff;
  keyInput[encryptionKey.length + 3] = generation & 0xff;
  keyInput[encryptionKey.length + 4] = (generation >> 8) & 0xff;

  const objKey = md5(keyInput);
  // Use min(n+5, 16) bytes of the hash as the RC4 key
  const keyLen = Math.min(encryptionKey.length + 5, 16);
  const rc4Key = objKey.subarray(0, keyLen);

  return rc4(rc4Key, data);
}

// =============================================================================
// RC4 Cipher
// =============================================================================

/**
 * RC4 stream cipher implementation.
 */
export function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
  // Key Scheduling Algorithm (KSA)
  const s = new Uint8Array(256);
  for (let i = 0; i < 256; i++) {
    s[i] = i;
  }
  let j = 0;
  for (let i = 0; i < 256; i++) {
    j = (j + s[i] + key[i % key.length]) & 0xff;
    [s[i], s[j]] = [s[j], s[i]];
  }

  // Pseudo-Random Generation Algorithm (PRGA)
  const result = new Uint8Array(data.length);
  let ii = 0;
  let jj = 0;
  for (let k = 0; k < data.length; k++) {
    ii = (ii + 1) & 0xff;
    jj = (jj + s[ii]) & 0xff;
    [s[ii], s[jj]] = [s[jj], s[ii]];
    result[k] = data[k] ^ s[(s[ii] + s[jj]) & 0xff];
  }
  return result;
}

// =============================================================================
// MD5 Hash
// =============================================================================

/**
 * MD5 hash implementation (RFC 1321).
 * Returns 16-byte digest.
 */
export function md5(input: Uint8Array): Uint8Array {
  // Pre-processing: padding
  const msgLen = input.length;
  const bitLen = msgLen * 8;
  // Pad to 64-byte boundary (56 bytes mod 64, then 8 bytes length)
  const padLen = ((56 - ((msgLen + 1) % 64) + 64) % 64) + 1;
  const padded = new Uint8Array(msgLen + padLen + 8);
  padded.set(input);
  padded[msgLen] = 0x80;
  // Append length in bits as 64-bit little-endian
  const view = new DataView(padded.buffer);
  view.setUint32(padded.length - 8, bitLen >>> 0, true);
  view.setUint32(padded.length - 4, 0, true); // high 32 bits (always 0 for our sizes)

  // Initialize hash values
  let a0 = 0x67452301;
  let b0 = 0xefcdab89;
  let c0 = 0x98badcfe;
  let d0 = 0x10325476;

  // Process each 64-byte block
  for (let i = 0; i < padded.length; i += 64) {
    const M = new Uint32Array(16);
    for (let j = 0; j < 16; j++) {
      M[j] = view.getUint32(i + j * 4, true);
    }

    let A = a0;
    let B = b0;
    let C = c0;
    let D = d0;

    for (let j = 0; j < 64; j++) {
      let F: number;
      let g: number;
      if (j < 16) {
        F = (B & C) | (~B & D);
        g = j;
      } else if (j < 32) {
        F = (D & B) | (~D & C);
        g = (5 * j + 1) % 16;
      } else if (j < 48) {
        F = B ^ C ^ D;
        g = (3 * j + 5) % 16;
      } else {
        F = C ^ (B | ~D);
        g = (7 * j) % 16;
      }
      F = (F + A + K[j] + M[g]) >>> 0;
      A = D;
      D = C;
      C = B;
      B = (B + rotl(F, S[j])) >>> 0;
    }

    a0 = (a0 + A) >>> 0;
    b0 = (b0 + B) >>> 0;
    c0 = (c0 + C) >>> 0;
    d0 = (d0 + D) >>> 0;
  }

  // Produce the 128-bit digest
  const digest = new Uint8Array(16);
  const dv = new DataView(digest.buffer);
  dv.setUint32(0, a0, true);
  dv.setUint32(4, b0, true);
  dv.setUint32(8, c0, true);
  dv.setUint32(12, d0, true);
  return digest;
}

function rotl(x: number, n: number): number {
  return ((x << n) | (x >>> (32 - n))) >>> 0;
}

// MD5 per-round shift amounts
const S = [
  7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14,
  20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6,
  10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21
];

// MD5 per-round constants (floor(2^32 × abs(sin(i+1))))
const K = new Uint32Array([
  0xd76aa478, 0xe8c7b756, 0x242070db, 0xc1bdceee, 0xf57c0faf, 0x4787c62a, 0xa8304613, 0xfd469501,
  0x698098d8, 0x8b44f7af, 0xffff5bb1, 0x895cd7be, 0x6b901122, 0xfd987193, 0xa679438e, 0x49b40821,
  0xf61e2562, 0xc040b340, 0x265e5a51, 0xe9b6c7aa, 0xd62f105d, 0x02441453, 0xd8a1e681, 0xe7d3fbc8,
  0x21e1cde6, 0xc33707d6, 0xf4d50d87, 0x455a14ed, 0xa9e3e905, 0xfcefa3f8, 0x676f02d9, 0x8d2a4c8a,
  0xfffa3942, 0x8771f681, 0x6d9d6122, 0xfde5380c, 0xa4beea44, 0x4bdecfa9, 0xf6bb4b60, 0xbebfbc70,
  0x289b7ec6, 0xeaa127fa, 0xd4ef3085, 0x04881d05, 0xd9d4d039, 0xe6db99e5, 0x1fa27cf8, 0xc4ac5665,
  0xf4292244, 0x432aff97, 0xab9423a7, 0xfc93a039, 0x655b59c3, 0x8f0ccc92, 0xffeff47d, 0x85845dd1,
  0x6fa87e4f, 0xfe2ce6e0, 0xa3014314, 0x4e0811a1, 0xf7537e82, 0xbd3af235, 0x2ad7d2bb, 0xeb86d391
]);

// =============================================================================
// PDF Password / Key Computation
// =============================================================================

/**
 * Pad or truncate a password to 32 bytes using the PDF password padding.
 */
function padPassword(password: string): Uint8Array {
  const result = new Uint8Array(32);
  const bytes = new TextEncoder().encode(password);
  const len = Math.min(bytes.length, 32);
  result.set(bytes.subarray(0, len));
  result.set(PASSWORD_PADDING.subarray(0, 32 - len), len);
  return result;
}

/**
 * Compute the O (owner) value.
 * Algorithm 3 from PDF spec §3.5.2.
 */
function computeOValue(ownerPassword: string, userPassword: string): Uint8Array {
  // Step 1: MD5 hash of padded owner password
  let hash = md5(padPassword(ownerPassword));
  // Step 2: For revision 3, hash 50 more times
  for (let i = 0; i < 50; i++) {
    hash = md5(hash);
  }
  // Use first 16 bytes as RC4 key (128-bit / key length = 16)
  const rc4Key = hash.subarray(0, 16);

  // Step 3: RC4-encrypt the padded user password
  let result = rc4(rc4Key, padPassword(userPassword));

  // Step 4: For revision 3, iterate 1-19 with modified key
  for (let i = 1; i <= 19; i++) {
    const modKey = new Uint8Array(16);
    for (let j = 0; j < 16; j++) {
      modKey[j] = rc4Key[j] ^ i;
    }
    result = rc4(modKey, result);
  }

  return result;
}

/**
 * Compute the encryption key.
 * Algorithm 2 from PDF spec §3.5.2.
 */
function computeEncryptionKey(
  userPassword: string,
  oValue: Uint8Array,
  permissions: number,
  fileId: Uint8Array
): Uint8Array {
  // Concatenate: padded password + O value + P value (4 LE bytes) + file ID
  const paddedPwd = padPassword(userPassword);
  const input = new Uint8Array(32 + 32 + 4 + fileId.length);
  input.set(paddedPwd);
  input.set(oValue, 32);
  const pView = new DataView(input.buffer, input.byteOffset);
  pView.setInt32(64, permissions, true);
  input.set(fileId, 68);

  let hash = md5(input);
  // For revision 3, hash 50 more times
  for (let i = 0; i < 50; i++) {
    hash = md5(hash.subarray(0, 16));
  }

  return hash.subarray(0, 16); // 128-bit key
}

/**
 * Compute the U (user) value.
 * Algorithm 5 from PDF spec §3.5.2 (revision 3).
 */
function computeUValue(encryptionKey: Uint8Array, fileId: Uint8Array): Uint8Array {
  // Step 1: MD5 hash of padding + file ID
  const hashInput = new Uint8Array(32 + fileId.length);
  hashInput.set(PASSWORD_PADDING);
  hashInput.set(fileId, 32);
  const hash = md5(hashInput);

  // Step 2: RC4-encrypt with the encryption key
  let result = rc4(encryptionKey, hash);

  // Step 3: Iterate 1-19 with modified key
  for (let i = 1; i <= 19; i++) {
    const modKey = new Uint8Array(16);
    for (let j = 0; j < 16; j++) {
      modKey[j] = encryptionKey[j] ^ i;
    }
    result = rc4(modKey, result);
  }

  // Pad to 32 bytes with arbitrary padding
  const uValue = new Uint8Array(32);
  uValue.set(result);
  return uValue;
}

/**
 * Compute the permissions integer (P value) from permission flags.
 */
function computePermissions(perms?: Partial<PdfPermissions>): number {
  // Start with all bits set that are "reserved" and must be 1
  // Bits 1-2, 7-8 must be 0; bits 13-32 must be 1 (per spec)
  let p = 0xfffff000 | 0b11000000; // bits 7-8 = reserved 1, high bits = 1

  if (perms?.print) {
    p |= 1 << 2; // bit 3
  }
  if (perms?.modify) {
    p |= 1 << 3; // bit 4
  }
  if (perms?.copy) {
    p |= 1 << 4; // bit 5
  }
  if (perms?.annotate) {
    p |= 1 << 5; // bit 6
  }
  if (perms?.fillForms) {
    p |= 1 << 8; // bit 9
  }
  if (perms?.accessibility) {
    p |= 1 << 9; // bit 10
  }
  if (perms?.assemble) {
    p |= 1 << 10; // bit 11
  }
  if (perms?.printHighQuality) {
    p |= 1 << 11; // bit 12
  }

  // Convert to signed 32-bit
  return p | 0;
}

/**
 * Generate a random file identifier (16 bytes).
 */
function generateFileId(): Uint8Array {
  // Use MD5 of current timestamp + random for determinism in tests
  const seed = new Uint8Array(16);
  const now = Date.now();
  const view = new DataView(seed.buffer);
  view.setFloat64(0, now, true);
  view.setFloat64(8, Math.random() * 1e15, true);
  return md5(seed);
}
