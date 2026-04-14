/**
 * Node.js Encryptor — uses shared crypto primitives from `@utils/crypto`.
 */

import { concatUint8Arrays } from "@utils/binary";
import { hash, randomBytes } from "@utils/crypto";
import { base64ToUint8Array, uint8ArrayToBase64, stringToUtf16Le } from "@utils/utils.base";

function uint32ToLe(num: number): Uint8Array {
  const arr = new Uint8Array(4);
  arr[0] = num & 0xff;
  arr[1] = (num >> 8) & 0xff;
  arr[2] = (num >> 16) & 0xff;
  arr[3] = (num >> 24) & 0xff;
  return arr;
}

const Encryptor = {
  hash(algorithm: string, ...buffers: Uint8Array[]): Uint8Array {
    return hash(algorithm, concatUint8Arrays(buffers));
  },

  async convertPasswordToHash(
    password: string,
    hashAlgorithm: string,
    saltValue: string,
    spinCount: number
  ): Promise<string> {
    const passwordBuffer = stringToUtf16Le(password);
    const saltBuffer = base64ToUint8Array(saltValue);

    let key = this.hash(hashAlgorithm, saltBuffer, passwordBuffer);
    for (let i = 0; i < spinCount; i++) {
      key = this.hash(hashAlgorithm, key, uint32ToLe(i));
    }

    return uint8ArrayToBase64(key);
  },

  randomBytes(size: number): Uint8Array {
    return randomBytes(size);
  }
};

export { Encryptor };
