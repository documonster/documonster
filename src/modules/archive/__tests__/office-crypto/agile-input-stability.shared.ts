/**
 * What a decrypt verifies is what it decrypts, and what it refuses it refuses before it pays for it.
 *
 * The read yields between the integrity check and the decryption, so a caller's buffer written to in between was
 * verified as it had been and decrypted as it had become: altered output and no error. The decrypt now works on its own
 * copy of the package and of the parameters it was given. And that copy is made only once the size the package declares
 * has been checked against the bound — it used to be made first, a whole package allocated for a read then refused.
 *
 * Shared by the Node and the browser runs: the platforms hash, copy and sign differently.
 */
import {
  decryptAgilePackage,
  decryptPackage,
  deriveAgileKey,
  encryptPackage,
  OfficeCryptoError,
  readEncryptedPackage,
  verifyAgilePassword
} from "@archive/office-crypto/agile";
import { scheduleMacrotask } from "@utils/yield";
import { afterEach, describe, expect, it, vi } from "vitest";

/** Bytes whose `slice` is a view, as a Node `Buffer`'s is: what a copy made with `slice` would still share. */
class ViewSlicing extends Uint8Array {
  override slice(start?: number, end?: number): Uint8Array<ArrayBuffer> {
    return this.subarray(start, end);
  }
}

/** The tests, for the Node and the browser runs to register. */
export function describeInputStability(): void {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  const zip = new Uint8Array(2 * 1024 * 1024).map((_, i) => (i * 7) & 255);

  /** Whether `bytes` are `zip`'s — compared directly, as `toEqual` takes seconds over two megabytes. */
  function isZip(bytes: Uint8Array): boolean {
    return bytes.length === zip.length && bytes.every((byte, i) => byte === zip[i]);
  }

  describe("decrypting a package whose input changes during the decrypt", () => {
    it("decrypts what it verified, not what the input became", async () => {
      const input = await encryptPackage(zip, "pw", { spinCount: 1000 });

      // The moment the integrity check has been computed, change the end of the input — on the next turn, as a caller
      // writing to its buffer would.
      const subtle = globalThis.crypto.subtle;
      const sign = subtle.sign.bind(subtle);
      vi.spyOn(subtle, "sign").mockImplementation((...args: Parameters<SubtleCrypto["sign"]>) => {
        scheduleMacrotask(() => input.fill(0, input.length - 64 * 1024));
        return sign(...args);
      });

      expect(isZip(await decryptPackage(input, "pw"))).toBe(true);
      expect(subtle.sign).toHaveBeenCalledTimes(1);
    });

    it("decrypts with the parameters it was given, not what the caller's object became", async () => {
      const { info, encryptedPackage } = readEncryptedPackage(
        await encryptPackage(zip, "pw", { spinCount: 50_000 })
      );
      // Change the parameters on every turn while the password is spun: the spin count, and the salt it is spun with.
      let running = true;
      let tampered = 0;
      const tamper = (): void => {
        if (running) {
          (info as { spinCount: number }).spinCount = 1;
          info.keySalt.fill(0);
          tampered++;
          scheduleMacrotask(tamper);
        }
      };
      scheduleMacrotask(tamper);
      try {
        expect(isZip(await decryptAgilePackage(encryptedPackage, info, "pw"))).toBe(true);
      } finally {
        running = false;
      }
      // The parameters were changed while the decrypt ran, not only after it — or nothing was tested.
      expect(tampered).toBeGreaterThan(0);
    });
  });

  describe("the parameters a caller gives", () => {
    const keyInfo = {
      keySalt: new Uint8Array(16).fill(3),
      spinCount: 0,
      hashAlgorithm: "SHA512",
      keyBits: 256
    };
    const blockKey = [1, 2, 3, 4, 5, 6, 7, 8];

    it("are copied whatever kind of view their bytes are", async () => {
      const expected = await deriveAgileKey("pw", keyInfo, new Uint8Array(blockKey));
      const given = ViewSlicing.from(blockKey);
      const derived = deriveAgileKey(
        "pw",
        { ...keyInfo, keySalt: ViewSlicing.from(keyInfo.keySalt) },
        given
      );
      given.fill(9);
      expect(await derived).toEqual(expected);
    });

    it.skipIf(typeof SharedArrayBuffer === "undefined")(
      "are copied out of shared memory",
      async () => {
        // The salt is read again after the password is spun — it is the IV of the package key — so a change to shared
        // memory during the spin reached the decrypt: `structuredClone` copies the view and keeps sharing the memory.
        const { info, encryptedPackage } = readEncryptedPackage(
          await encryptPackage(zip, "pw", { spinCount: 50_000 })
        );
        const salt = new Uint8Array(new SharedArrayBuffer(info.keySalt.length));
        salt.set(info.keySalt);
        let running = true;
        let tampered = 0;
        const tamper = (): void => {
          if (running) {
            salt.fill(0);
            tampered++;
            scheduleMacrotask(tamper);
          }
        };
        const decrypted = decryptAgilePackage(encryptedPackage, { ...info, keySalt: salt }, "pw");
        scheduleMacrotask(tamper);
        try {
          expect(isZip(await decrypted)).toBe(true);
        } finally {
          running = false;
        }
        expect(tampered).toBeGreaterThan(0);
      }
    );

    // The info is a structural type: what else a caller's object holds is neither read nor copied.
    it("may carry fields the operation does not read, of any kind", async () => {
      const { info } = readEncryptedPackage(await encryptPackage(zip, "pw", { spinCount: 1000 }));
      const annotated = { ...info, describe: () => "a function cannot be cloned" };
      expect(await verifyAgilePassword("pw", annotated)).toBe(true);
      const noted = { ...keyInfo, note: () => undefined };
      await expect(deriveAgileKey("pw", noted, new Uint8Array(blockKey))).resolves.toBeInstanceOf(
        Uint8Array
      );
    });
  });

  // A view of a resizable buffer shrinks with it: the package was read as one of another size, and failed with
  // whatever that size made of it.
  it.skipIf(typeof ArrayBuffer.prototype.resize !== "function")(
    "refuses a package whose buffer shrinks while the password is spun",
    async () => {
      const encrypted = await encryptPackage(zip, "pw", { spinCount: 100_000 });
      const buffer = new ArrayBuffer(encrypted.length, { maxByteLength: encrypted.length });
      const bytes = new Uint8Array(buffer);
      bytes.set(encrypted);
      scheduleMacrotask(() => buffer.resize(16));
      await expect(decryptPackage(bytes, "pw")).rejects.toThrow(/changed size while it was read/);
    }
  );

  describe("decrypting a package larger than the read allows", () => {
    it("refuses it before allocating anything for it", async () => {
      const input = await encryptPackage(zip, "pw", { spinCount: 1000 });
      const { info, encryptedPackage } = readEncryptedPackage(input);
      const allocations: number[] = [];
      const Native = globalThis.Uint8Array;
      vi.stubGlobal(
        "Uint8Array",
        new Proxy(Native, {
          construct(target, args, newTarget) {
            if (typeof args[0] === "number" && args[0] >= 1024 * 1024) {
              allocations.push(args[0]);
            }
            return Reflect.construct(target, args, newTarget) as object;
          }
        })
      );
      let refusal: unknown;
      try {
        await decryptAgilePackage(encryptedPackage, info, "pw", 1);
      } catch (error) {
        refusal = error;
      } finally {
        vi.unstubAllGlobals();
      }
      expect(refusal).toBeInstanceOf(OfficeCryptoError);
      expect((refusal as OfficeCryptoError).reason).toBe("limit");
      expect(allocations).toEqual([]);
    });
  });
}
