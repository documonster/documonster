/**
 * Browser-only Encryptor — uses shared crypto primitives from `@utils/crypto`.
 */

import { hashAsync, randomBytes } from "@utils/crypto";
import { base64ToUint8Array, uint8ArrayToBase64, stringToUtf16Le } from "@utils/utils.base";
import { concatUint8Arrays } from "@utils/binary";

function uint32ToLe(num: number): Uint8Array {
  const arr = new Uint8Array(4);
  arr[0] = num & 0xff;
  arr[1] = (num >> 8) & 0xff;
  arr[2] = (num >> 16) & 0xff;
  arr[3] = (num >> 24) & 0xff;
  return arr;
}

const Encryptor = {
  /**
   * Calculate hash using shared crypto (Web Crypto API in browser).
   */
  async hash(algorithm: string, ...buffers: Uint8Array[]): Promise<Uint8Array> {
    return hashAsync(algorithm, concatUint8Arrays(buffers));
  },

  /**
   * Convert password to hash.
   */
  async convertPasswordToHash(
    password: string,
    hashAlgorithm: string,
    saltValue: string,
    spinCount: number
  ): Promise<string> {
    const passwordBuffer = stringToUtf16Le(password);
    const saltBuffer = base64ToUint8Array(saltValue);

    let key = await this.hash(hashAlgorithm, saltBuffer, passwordBuffer);

    for (let i = 0; i < spinCount; i++) {
      key = await this.hash(hashAlgorithm, key, uint32ToLe(i));
    }

    return uint8ArrayToBase64(key);
  },

  /**
   * Generate cryptographically strong random bytes.
   */
  randomBytes(size: number): Uint8Array {
    return randomBytes(size);
  }
};

export { Encryptor };
