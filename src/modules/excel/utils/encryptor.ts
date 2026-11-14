/**
 * The password hash of sheet and workbook protection (ECMA-376 Part 1 §18.2.29), with the synchronous `hash` of each
 * platform: SHA-1, SHA-256, SHA-384, SHA-512 and MD5 everywhere — what this library writes is SHA-512 — and whatever
 * else OpenSSL has on Node (RIPEMD-160). An algorithm a file names that the platform does not have is refused as
 * {@link ExcelUnsupportedFeatureError}.
 */

import { ExcelUnsupportedFeatureError } from "@excel/errors";
import { concatUint8Arrays } from "@utils/binary";
import { hash, randomBytes } from "@utils/crypto";
import { base64ToUint8Array, uint8ArrayToBase64, stringToUtf16Le } from "@utils/utils.base";
import { createYieldBudget } from "@utils/yield";

function uint32ToLe(num: number): Uint8Array {
  const arr = new Uint8Array(4);
  arr[0] = num & 0xff;
  arr[1] = (num >> 8) & 0xff;
  arr[2] = (num >> 16) & 0xff;
  arr[3] = (num >> 24) & 0xff;
  return arr;
}

/** Digests the spin runs between two looks at the clock, less one: a mask, as the loop counts in powers of two. */
const SPIN_BATCH_MASK = 1023;

const Encryptor = {
  hash(algorithm: string, ...buffers: Uint8Array[]): Uint8Array {
    return hash(algorithm, concatUint8Arrays(buffers));
  },

  /**
   * The password's hash after `spinCount` rounds — 100,000 by default, which ran as one synchronous loop in Node and
   * as one awaited Web Crypto digest per round in a browser (four seconds in Safari). Synchronous digests on both,
   * yielding once the loop has worked for the budget.
   *
   * @throws {ExcelUnsupportedFeatureError} When the platform has no `hashAlgorithm`.
   */
  async convertPasswordToHash(
    password: string,
    hashAlgorithm: string,
    saltValue: string,
    spinCount: number
  ): Promise<string> {
    const passwordBuffer = stringToUtf16Le(password);
    const saltBuffer = base64ToUint8Array(saltValue);
    const yieldIfDue = createYieldBudget();

    let key: Uint8Array;
    try {
      key = this.hash(hashAlgorithm, saltBuffer, passwordBuffer);
    } catch (error) {
      // The first digest is where an algorithm the platform does not have is found.
      throw new ExcelUnsupportedFeatureError(
        "Protection password hash",
        `the algorithm "${hashAlgorithm}" is not available on this platform`,
        { cause: error }
      );
    }
    for (let i = 0; i < spinCount; i++) {
      key = this.hash(hashAlgorithm, key, uint32ToLe(i));
      if ((i & SPIN_BATCH_MASK) === SPIN_BATCH_MASK) {
        await yieldIfDue();
      }
    }

    return uint8ArrayToBase64(key);
  },

  randomBytes(size: number): Uint8Array {
    return randomBytes(size);
  }
};

export { Encryptor };
