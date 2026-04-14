/**
 * Cryptographic primitives — Browser version.
 *
 * Uses pure JavaScript for synchronous operations (SHA-256, MD5, AES-CBC, RC4)
 * since Web Crypto API is async-only and cannot replace synchronous call sites.
 *
 * Uses Web Crypto API for:
 * - `randomBytes` (crypto.getRandomValues — truly random)
 * - `rsaVerify` / `rsaSign` (SubtleCrypto — hardware-accelerated)
 *
 * Exports the same API as `crypto.ts` (Node.js version).
 */

// =============================================================================
// AES S-Box & Constants
// =============================================================================

const SBOX = new Uint8Array([
  0x63, 0x7c, 0x77, 0x7b, 0xf2, 0x6b, 0x6f, 0xc5, 0x30, 0x01, 0x67, 0x2b, 0xfe, 0xd7, 0xab, 0x76,
  0xca, 0x82, 0xc9, 0x7d, 0xfa, 0x59, 0x47, 0xf0, 0xad, 0xd4, 0xa2, 0xaf, 0x9c, 0xa4, 0x72, 0xc0,
  0xb7, 0xfd, 0x93, 0x26, 0x36, 0x3f, 0xf7, 0xcc, 0x34, 0xa5, 0xe5, 0xf1, 0x71, 0xd8, 0x31, 0x15,
  0x04, 0xc7, 0x23, 0xc3, 0x18, 0x96, 0x05, 0x9a, 0x07, 0x12, 0x80, 0xe2, 0xeb, 0x27, 0xb2, 0x75,
  0x09, 0x83, 0x2c, 0x1a, 0x1b, 0x6e, 0x5a, 0xa0, 0x52, 0x3b, 0xd6, 0xb3, 0x29, 0xe3, 0x2f, 0x84,
  0x53, 0xd1, 0x00, 0xed, 0x20, 0xfc, 0xb1, 0x5b, 0x6a, 0xcb, 0xbe, 0x39, 0x4a, 0x4c, 0x58, 0xcf,
  0xd0, 0xef, 0xaa, 0xfb, 0x43, 0x4d, 0x33, 0x85, 0x45, 0xf9, 0x02, 0x7f, 0x50, 0x3c, 0x9f, 0xa8,
  0x51, 0xa3, 0x40, 0x8f, 0x92, 0x9d, 0x38, 0xf5, 0xbc, 0xb6, 0xda, 0x21, 0x10, 0xff, 0xf3, 0xd2,
  0xcd, 0x0c, 0x13, 0xec, 0x5f, 0x97, 0x44, 0x17, 0xc4, 0xa7, 0x7e, 0x3d, 0x64, 0x5d, 0x19, 0x73,
  0x60, 0x81, 0x4f, 0xdc, 0x22, 0x2a, 0x90, 0x88, 0x46, 0xee, 0xb8, 0x14, 0xde, 0x5e, 0x0b, 0xdb,
  0xe0, 0x32, 0x3a, 0x0a, 0x49, 0x06, 0x24, 0x5c, 0xc2, 0xd3, 0xac, 0x62, 0x91, 0x95, 0xe4, 0x79,
  0xe7, 0xc8, 0x37, 0x6d, 0x8d, 0xd5, 0x4e, 0xa9, 0x6c, 0x56, 0xf4, 0xea, 0x65, 0x7a, 0xae, 0x08,
  0xba, 0x78, 0x25, 0x2e, 0x1c, 0xa6, 0xb4, 0xc6, 0xe8, 0xdd, 0x74, 0x1f, 0x4b, 0xbd, 0x8b, 0x8a,
  0x70, 0x3e, 0xb5, 0x66, 0x48, 0x03, 0xf6, 0x0e, 0x61, 0x35, 0x57, 0xb9, 0x86, 0xc1, 0x1d, 0x9e,
  0xe1, 0xf8, 0x98, 0x11, 0x69, 0xd9, 0x8e, 0x94, 0x9b, 0x1e, 0x87, 0xe9, 0xce, 0x55, 0x28, 0xdf,
  0x8c, 0xa1, 0x89, 0x0d, 0xbf, 0xe6, 0x42, 0x68, 0x41, 0x99, 0x2d, 0x0f, 0xb0, 0x54, 0xbb, 0x16
]);

const INV_SBOX = new Uint8Array(256);
/* @__PURE__ */ (() => {
  for (let i = 0; i < 256; i++) {
    INV_SBOX[SBOX[i]] = i;
  }
})();

const RCON = [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80, 0x1b, 0x36];

// =============================================================================
// AES Helpers
// =============================================================================

function gf2(a: number): number {
  return a < 128 ? a << 1 : (a << 1) ^ 0x11b;
}

function gfMul(a: number, b: number): number {
  let result = 0;
  let aa = a;
  let bb = b;
  while (bb > 0) {
    if (bb & 1) {
      result ^= aa;
    }
    aa = gf2(aa);
    bb >>= 1;
  }
  return result;
}

function aesKeyExpansion(key: Uint8Array): Uint8Array[] {
  const nk = key.length / 4;
  const nr = nk + 6;
  const w: Uint8Array[] = [];

  for (let i = 0; i < nk; i++) {
    w.push(new Uint8Array([key[4 * i], key[4 * i + 1], key[4 * i + 2], key[4 * i + 3]]));
  }

  for (let i = nk; i < 4 * (nr + 1); i++) {
    const temp = new Uint8Array(w[i - 1]);
    if (i % nk === 0) {
      const t0 = temp[0];
      temp[0] = SBOX[temp[1]] ^ RCON[i / nk - 1];
      temp[1] = SBOX[temp[2]];
      temp[2] = SBOX[temp[3]];
      temp[3] = SBOX[t0];
    } else if (nk > 6 && i % nk === 4) {
      temp[0] = SBOX[temp[0]];
      temp[1] = SBOX[temp[1]];
      temp[2] = SBOX[temp[2]];
      temp[3] = SBOX[temp[3]];
    }
    const word = new Uint8Array(4);
    for (let j = 0; j < 4; j++) {
      word[j] = w[i - nk][j] ^ temp[j];
    }
    w.push(word);
  }

  return w;
}

// =============================================================================
// AES Block Encrypt / Decrypt
// =============================================================================

function aesEncryptBlock(block: Uint8Array, roundKeys: Uint8Array[]): Uint8Array {
  const nr = roundKeys.length / 4 - 1;
  const state = new Uint8Array(16);
  state.set(block);

  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < 4; r++) {
      state[4 * c + r] ^= roundKeys[c][r];
    }
  }

  for (let round = 1; round < nr; round++) {
    for (let i = 0; i < 16; i++) {
      state[i] = SBOX[state[i]];
    }

    let tmp: number;
    tmp = state[1];
    state[1] = state[5];
    state[5] = state[9];
    state[9] = state[13];
    state[13] = tmp;
    tmp = state[2];
    state[2] = state[10];
    state[10] = tmp;
    tmp = state[6];
    state[6] = state[14];
    state[14] = tmp;
    tmp = state[15];
    state[15] = state[11];
    state[11] = state[7];
    state[7] = state[3];
    state[3] = tmp;

    for (let c = 0; c < 4; c++) {
      const s0 = state[4 * c];
      const s1 = state[4 * c + 1];
      const s2 = state[4 * c + 2];
      const s3 = state[4 * c + 3];
      state[4 * c] = gf2(s0) ^ gf2(s1) ^ s1 ^ s2 ^ s3;
      state[4 * c + 1] = s0 ^ gf2(s1) ^ gf2(s2) ^ s2 ^ s3;
      state[4 * c + 2] = s0 ^ s1 ^ gf2(s2) ^ gf2(s3) ^ s3;
      state[4 * c + 3] = gf2(s0) ^ s0 ^ s1 ^ s2 ^ gf2(s3);
    }

    const keyOffset = round * 4;
    for (let c = 0; c < 4; c++) {
      for (let r = 0; r < 4; r++) {
        state[4 * c + r] ^= roundKeys[keyOffset + c][r];
      }
    }
  }

  for (let i = 0; i < 16; i++) {
    state[i] = SBOX[state[i]];
  }
  let tmp: number;
  tmp = state[1];
  state[1] = state[5];
  state[5] = state[9];
  state[9] = state[13];
  state[13] = tmp;
  tmp = state[2];
  state[2] = state[10];
  state[10] = tmp;
  tmp = state[6];
  state[6] = state[14];
  state[14] = tmp;
  tmp = state[15];
  state[15] = state[11];
  state[11] = state[7];
  state[7] = state[3];
  state[3] = tmp;

  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < 4; r++) {
      state[4 * c + r] ^= roundKeys[nr * 4 + c][r];
    }
  }

  return state;
}

function aesDecryptBlock(block: Uint8Array, roundKeys: Uint8Array[]): Uint8Array {
  const nr = roundKeys.length / 4 - 1;
  const state = new Uint8Array(16);
  state.set(block);

  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < 4; r++) {
      state[4 * c + r] ^= roundKeys[nr * 4 + c][r];
    }
  }

  for (let round = nr - 1; round >= 1; round--) {
    let tmp: number;
    tmp = state[13];
    state[13] = state[9];
    state[9] = state[5];
    state[5] = state[1];
    state[1] = tmp;
    tmp = state[2];
    state[2] = state[10];
    state[10] = tmp;
    tmp = state[6];
    state[6] = state[14];
    state[14] = tmp;
    tmp = state[3];
    state[3] = state[7];
    state[7] = state[11];
    state[11] = state[15];
    state[15] = tmp;

    for (let i = 0; i < 16; i++) {
      state[i] = INV_SBOX[state[i]];
    }

    const keyOffset = round * 4;
    for (let c = 0; c < 4; c++) {
      for (let r = 0; r < 4; r++) {
        state[4 * c + r] ^= roundKeys[keyOffset + c][r];
      }
    }

    for (let c = 0; c < 4; c++) {
      const s0 = state[4 * c];
      const s1 = state[4 * c + 1];
      const s2 = state[4 * c + 2];
      const s3 = state[4 * c + 3];
      state[4 * c] = gfMul(s0, 14) ^ gfMul(s1, 11) ^ gfMul(s2, 13) ^ gfMul(s3, 9);
      state[4 * c + 1] = gfMul(s0, 9) ^ gfMul(s1, 14) ^ gfMul(s2, 11) ^ gfMul(s3, 13);
      state[4 * c + 2] = gfMul(s0, 13) ^ gfMul(s1, 9) ^ gfMul(s2, 14) ^ gfMul(s3, 11);
      state[4 * c + 3] = gfMul(s0, 11) ^ gfMul(s1, 13) ^ gfMul(s2, 9) ^ gfMul(s3, 14);
    }
  }

  let tmp2: number;
  tmp2 = state[13];
  state[13] = state[9];
  state[9] = state[5];
  state[5] = state[1];
  state[1] = tmp2;
  tmp2 = state[2];
  state[2] = state[10];
  state[10] = tmp2;
  tmp2 = state[6];
  state[6] = state[14];
  state[14] = tmp2;
  tmp2 = state[3];
  state[3] = state[7];
  state[7] = state[11];
  state[11] = state[15];
  state[15] = tmp2;

  for (let i = 0; i < 16; i++) {
    state[i] = INV_SBOX[state[i]];
  }

  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < 4; r++) {
      state[4 * c + r] ^= roundKeys[c][r];
    }
  }

  return state;
}

// =============================================================================
// AES-CBC Public API
// =============================================================================

export function aesCbcEncrypt(plaintext: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
  const padLen = 16 - (plaintext.length % 16);
  const padded = new Uint8Array(plaintext.length + padLen);
  padded.set(plaintext);
  for (let i = plaintext.length; i < padded.length; i++) {
    padded[i] = padLen;
  }

  const roundKeys = aesKeyExpansion(key);
  const numBlocks = padded.length / 16;
  const output = new Uint8Array(padded.length);
  let prevBlock = iv;

  for (let b = 0; b < numBlocks; b++) {
    const block = new Uint8Array(16);
    for (let i = 0; i < 16; i++) {
      block[i] = padded[b * 16 + i] ^ prevBlock[i];
    }
    const encrypted = aesEncryptBlock(block, roundKeys);
    output.set(encrypted, b * 16);
    prevBlock = encrypted;
  }

  return output;
}

export function aesCbcDecrypt(ciphertext: Uint8Array, key: Uint8Array, iv: Uint8Array): Uint8Array {
  const roundKeys = aesKeyExpansion(key);
  const numBlocks = ciphertext.length / 16;
  const output = new Uint8Array(ciphertext.length);
  let prevBlock = iv;

  for (let b = 0; b < numBlocks; b++) {
    const block = ciphertext.subarray(b * 16, (b + 1) * 16);
    const decrypted = aesDecryptBlock(block, roundKeys);

    for (let i = 0; i < 16; i++) {
      output[b * 16 + i] = decrypted[i] ^ prevBlock[i];
    }
    prevBlock = block;
  }

  if (output.length > 0) {
    const padLen = output[output.length - 1];
    if (padLen > 0 && padLen <= 16) {
      let validPadding = true;
      for (let i = 0; i < padLen; i++) {
        if (output[output.length - 1 - i] !== padLen) {
          validPadding = false;
          break;
        }
      }
      if (validPadding) {
        return output.subarray(0, output.length - padLen);
      }
    }
  }

  return output;
}

export function aesCbcDecryptRaw(
  ciphertext: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Uint8Array {
  const roundKeys = aesKeyExpansion(key);
  const numBlocks = ciphertext.length / 16;
  const output = new Uint8Array(ciphertext.length);
  let prevBlock = iv;

  for (let b = 0; b < numBlocks; b++) {
    const block = ciphertext.subarray(b * 16, (b + 1) * 16);
    const decrypted = aesDecryptBlock(block, roundKeys);

    for (let i = 0; i < 16; i++) {
      output[b * 16 + i] = decrypted[i] ^ prevBlock[i];
    }
    prevBlock = block;
  }

  return output;
}

export function aesCbcEncryptRaw(
  plaintext: Uint8Array,
  key: Uint8Array,
  iv: Uint8Array
): Uint8Array {
  if (plaintext.length % 16 !== 0) {
    throw new Error("aesCbcEncryptRaw: plaintext length must be a multiple of 16");
  }

  const roundKeys = aesKeyExpansion(key);
  const numBlocks = plaintext.length / 16;
  const output = new Uint8Array(plaintext.length);
  let prevBlock = iv;

  for (let b = 0; b < numBlocks; b++) {
    const block = new Uint8Array(16);
    for (let i = 0; i < 16; i++) {
      block[i] = plaintext[b * 16 + i] ^ prevBlock[i];
    }
    const encrypted = aesEncryptBlock(block, roundKeys);
    output.set(encrypted, b * 16);
    prevBlock = encrypted;
  }

  return output;
}

export function aesEcbEncrypt(block: Uint8Array, key: Uint8Array): Uint8Array {
  const roundKeys = aesKeyExpansion(key);
  return aesEncryptBlock(block, roundKeys);
}

// =============================================================================
// SHA-256
// =============================================================================

const SHA256_H = new Uint32Array([
  0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19
]);

const SHA256_K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
]);

function rotr32(x: number, n: number): number {
  return ((x >>> n) | (x << (32 - n))) >>> 0;
}

export function sha256(input: Uint8Array): Uint8Array {
  const msgLen = input.length;
  const paddedLen = Math.ceil((msgLen + 9) / 64) * 64;
  const padded = new Uint8Array(paddedLen);
  padded.set(input);
  padded[msgLen] = 0x80;

  const bitLen = msgLen * 8;
  const view = new DataView(padded.buffer, padded.byteOffset, padded.byteLength);
  view.setUint32(paddedLen - 8, 0, false);
  view.setUint32(paddedLen - 4, bitLen, false);

  let h0 = SHA256_H[0];
  let h1 = SHA256_H[1];
  let h2 = SHA256_H[2];
  let h3 = SHA256_H[3];
  let h4 = SHA256_H[4];
  let h5 = SHA256_H[5];
  let h6 = SHA256_H[6];
  let h7 = SHA256_H[7];

  const w = new Uint32Array(64);
  for (let offset = 0; offset < paddedLen; offset += 64) {
    for (let i = 0; i < 16; i++) {
      w[i] = view.getUint32(offset + i * 4, false);
    }
    for (let i = 16; i < 64; i++) {
      const s0 = rotr32(w[i - 15], 7) ^ rotr32(w[i - 15], 18) ^ (w[i - 15] >>> 3);
      const s1 = rotr32(w[i - 2], 17) ^ rotr32(w[i - 2], 19) ^ (w[i - 2] >>> 10);
      w[i] = (w[i - 16] + s0 + w[i - 7] + s1) >>> 0;
    }

    let a = h0;
    let b = h1;
    let c = h2;
    let d = h3;
    let e = h4;
    let f = h5;
    let g = h6;
    let h = h7;

    for (let i = 0; i < 64; i++) {
      const S1 = rotr32(e, 6) ^ rotr32(e, 11) ^ rotr32(e, 25);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + S1 + ch + SHA256_K[i] + w[i]) >>> 0;
      const S0 = rotr32(a, 2) ^ rotr32(a, 13) ^ rotr32(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (S0 + maj) >>> 0;

      h = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }

    h0 = (h0 + a) >>> 0;
    h1 = (h1 + b) >>> 0;
    h2 = (h2 + c) >>> 0;
    h3 = (h3 + d) >>> 0;
    h4 = (h4 + e) >>> 0;
    h5 = (h5 + f) >>> 0;
    h6 = (h6 + g) >>> 0;
    h7 = (h7 + h) >>> 0;
  }

  const result = new Uint8Array(32);
  const resultView = new DataView(result.buffer);
  resultView.setUint32(0, h0, false);
  resultView.setUint32(4, h1, false);
  resultView.setUint32(8, h2, false);
  resultView.setUint32(12, h3, false);
  resultView.setUint32(16, h4, false);
  resultView.setUint32(20, h5, false);
  resultView.setUint32(24, h6, false);
  resultView.setUint32(28, h7, false);

  return result;
}

// =============================================================================
// MD5
// =============================================================================

function rotl(x: number, n: number): number {
  return ((x << n) | (x >>> (32 - n))) >>> 0;
}

const MD5_S = [
  7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 7, 12, 17, 22, 5, 9, 14, 20, 5, 9, 14, 20, 5, 9, 14,
  20, 5, 9, 14, 20, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 4, 11, 16, 23, 6, 10, 15, 21, 6,
  10, 15, 21, 6, 10, 15, 21, 6, 10, 15, 21
];

const MD5_K = new Uint32Array([
  0xd76aa478, 0xe8c7b756, 0x242070db, 0xc1bdceee, 0xf57c0faf, 0x4787c62a, 0xa8304613, 0xfd469501,
  0x698098d8, 0x8b44f7af, 0xffff5bb1, 0x895cd7be, 0x6b901122, 0xfd987193, 0xa679438e, 0x49b40821,
  0xf61e2562, 0xc040b340, 0x265e5a51, 0xe9b6c7aa, 0xd62f105d, 0x02441453, 0xd8a1e681, 0xe7d3fbc8,
  0x21e1cde6, 0xc33707d6, 0xf4d50d87, 0x455a14ed, 0xa9e3e905, 0xfcefa3f8, 0x676f02d9, 0x8d2a4c8a,
  0xfffa3942, 0x8771f681, 0x6d9d6122, 0xfde5380c, 0xa4beea44, 0x4bdecfa9, 0xf6bb4b60, 0xbebfbc70,
  0x289b7ec6, 0xeaa127fa, 0xd4ef3085, 0x04881d05, 0xd9d4d039, 0xe6db99e5, 0x1fa27cf8, 0xc4ac5665,
  0xf4292244, 0x432aff97, 0xab9423a7, 0xfc93a039, 0x655b59c3, 0x8f0ccc92, 0xffeff47d, 0x85845dd1,
  0x6fa87e4f, 0xfe2ce6e0, 0xa3014314, 0x4e0811a1, 0xf7537e82, 0xbd3af235, 0x2ad7d2bb, 0xeb86d391
]);

export function md5(input: Uint8Array): Uint8Array {
  const msgLen = input.length;
  const bitLen = msgLen * 8;
  const padLen = ((56 - ((msgLen + 1) % 64) + 64) % 64) + 1;
  const padded = new Uint8Array(msgLen + padLen + 8);
  padded.set(input);
  padded[msgLen] = 0x80;
  const view = new DataView(padded.buffer);
  view.setUint32(padded.length - 8, bitLen >>> 0, true);
  view.setUint32(padded.length - 4, 0, true);

  let a0 = 0x67452301;
  let b0 = 0xefcdab89;
  let c0 = 0x98badcfe;
  let d0 = 0x10325476;

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
      F = (F + A + MD5_K[j] + M[g]) >>> 0;
      A = D;
      D = C;
      C = B;
      B = (B + rotl(F, MD5_S[j])) >>> 0;
    }

    a0 = (a0 + A) >>> 0;
    b0 = (b0 + B) >>> 0;
    c0 = (c0 + C) >>> 0;
    d0 = (d0 + D) >>> 0;
  }

  const digest = new Uint8Array(16);
  const dv = new DataView(digest.buffer);
  dv.setUint32(0, a0, true);
  dv.setUint32(4, b0, true);
  dv.setUint32(8, c0, true);
  dv.setUint32(12, d0, true);
  return digest;
}

// =============================================================================
// RC4 (legacy)
// =============================================================================

/**
 * RC4 stream cipher.
 * @deprecated Only used for reading legacy encrypted PDFs.
 */
export function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
  const s = new Uint8Array(256);
  for (let i = 0; i < 256; i++) {
    s[i] = i;
  }
  let j = 0;
  for (let i = 0; i < 256; i++) {
    j = (j + s[i] + key[i % key.length]) & 0xff;
    [s[i], s[j]] = [s[j], s[i]];
  }

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
// Random bytes
// =============================================================================

/**
 * Generate cryptographically secure random bytes.
 * Uses crypto.getRandomValues (available in all modern browsers).
 */
export function randomBytes(length: number): Uint8Array {
  const bytes = new Uint8Array(length);
  globalThis.crypto.getRandomValues(bytes);
  return bytes;
}

// =============================================================================
// Generic hash (async in browser — Web Crypto API)
// =============================================================================

/**
 * Compute a hash digest using Web Crypto API.
 *
 * NOTE: In the browser, this is async. The Node.js version is sync.
 * For callers that need sync hashing, use `sha256()` or `md5()` directly.
 *
 * @param algorithm - Hash algorithm name (e.g., "SHA-256", "SHA-512", "SHA-1").
 * @param data - Data to hash
 * @returns The digest bytes
 */
export async function hashAsync(algorithm: string, data: Uint8Array): Promise<Uint8Array> {
  const buf = await globalThis.crypto.subtle.digest(
    normalizeAlgorithmForWebCrypto(algorithm),
    data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer
  );
  return new Uint8Array(buf);
}

/**
 * Normalize a hash algorithm name to the format Web Crypto API expects.
 * Accepts: "sha256", "SHA-256", "sha-256", "SHA256" → "SHA-256"
 */
function normalizeAlgorithmForWebCrypto(algorithm: string): string {
  const lower = algorithm.toLowerCase().replace(/-/g, "");
  switch (lower) {
    case "sha1":
      return "SHA-1";
    case "sha256":
      return "SHA-256";
    case "sha384":
      return "SHA-384";
    case "sha512":
      return "SHA-512";
    default:
      // Pass through for any other algorithm — let Web Crypto validate
      return algorithm;
  }
}

/**
 * Compute a hash digest synchronously (pure JS — SHA-256 and MD5 only).
 *
 * @param algorithm - "SHA-256" or "MD5" (case-insensitive, hyphens optional)
 * @param data - Data to hash
 * @returns The digest bytes
 * @throws If algorithm is not SHA-256 or MD5
 */
export function hash(algorithm: string, data: Uint8Array): Uint8Array {
  const algo = algorithm.toLowerCase().replace(/-/g, "");
  if (algo === "sha256") {
    return sha256(data);
  }
  if (algo === "md5") {
    return md5(data);
  }
  throw new Error(
    `hash: unsupported algorithm "${algorithm}" in browser sync mode. Use hashAsync() for other algorithms.`
  );
}

// =============================================================================
// RSA signature operations (async — Web Crypto API)
// =============================================================================

/**
 * Verify an RSA PKCS#1 v1.5 signature.
 *
 * @param publicKeyDer - DER-encoded SubjectPublicKeyInfo
 * @param signature - The signature bytes
 * @param data - The signed data (will be hashed with SHA-256)
 */
export async function rsaVerify(
  publicKeyDer: Uint8Array,
  signature: Uint8Array,
  data: Uint8Array
): Promise<boolean> {
  const key = await globalThis.crypto.subtle.importKey(
    "spki",
    publicKeyDer.buffer.slice(
      publicKeyDer.byteOffset,
      publicKeyDer.byteOffset + publicKeyDer.byteLength
    ) as ArrayBuffer,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["verify"]
  );
  return globalThis.crypto.subtle.verify(
    "RSASSA-PKCS1-v1_5",
    key,
    signature.buffer.slice(
      signature.byteOffset,
      signature.byteOffset + signature.byteLength
    ) as ArrayBuffer,
    data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer
  );
}

/**
 * Create an RSA PKCS#1 v1.5 signature.
 *
 * @param privateKeyDer - DER-encoded PKCS#8 private key
 * @param data - The data to sign (will be hashed with SHA-256)
 */
export async function rsaSign(privateKeyDer: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const key = await globalThis.crypto.subtle.importKey(
    "pkcs8",
    privateKeyDer.buffer.slice(
      privateKeyDer.byteOffset,
      privateKeyDer.byteOffset + privateKeyDer.byteLength
    ) as ArrayBuffer,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const sig = await globalThis.crypto.subtle.sign(
    "RSASSA-PKCS1-v1_5",
    key,
    data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer
  );
  return new Uint8Array(sig);
}
