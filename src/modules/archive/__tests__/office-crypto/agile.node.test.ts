/**
 * The Office file encryption Word and Excel share (`@archive/office-crypto/agile`, over `@archive/cfb/cfb`), against
 * files Office itself encrypted: Excel 16.114 saved `excel-encrypted.xlsx` / `.xlsb` with the password
 * `Excel-pw1`, Word 16.114 saved `word-encrypted.docx` with `Word-pw1` (see `tmp/office-tests/enc/README.md`).
 */
import { readFileSync } from "node:fs";

import { readCfb, writeCfb } from "@archive/cfb/cfb";
import {
  decryptPackage,
  encryptPackage,
  isEncryptedPackage,
  OfficeCryptoError,
  readEncryptedPackage,
  verifyPackagePassword
} from "@archive/office-crypto/agile";
import { unpackZip } from "@archive/zip-one-shot";
import { describe, expect, it } from "vitest";

const fixture = (path: string) => new Uint8Array(readFileSync(new URL(path, import.meta.url)));

const OFFICE_FILES = [
  [
    "Excel .xlsx",
    "../../../excel/__tests__/data/office-encrypted/excel-encrypted.xlsx",
    "Excel-pw1",
    "xl/workbook.xml"
  ],
  [
    "Excel .xlsb",
    "../../../excel/__tests__/data/office-encrypted/excel-encrypted.xlsb",
    "Excel-pw1",
    "xl/workbook.bin"
  ],
  [
    "Word .docx",
    "../../../word/__tests__/data/office-encrypted/word-encrypted.docx",
    "Word-pw1",
    "word/document.xml"
  ]
] as const;

async function failure(promise: Promise<unknown>): Promise<OfficeCryptoError> {
  const error = await promise.then(
    () => undefined,
    (e: unknown) => e
  );
  expect(error).toBeInstanceOf(OfficeCryptoError);
  return error as OfficeCryptoError;
}

describe("a package Office encrypted", () => {
  it.each(OFFICE_FILES)("%s decrypts to its OOXML package", async (_name, path, password, part) => {
    const bytes = fixture(path);
    expect(isEncryptedPackage(bytes)).toBe(true);
    expect(await verifyPackagePassword(bytes, password)).toBe(true);
    const zip = await decryptPackage(bytes, password);
    expect([...zip.subarray(0, 2)]).toEqual([0x50, 0x4b]);
    expect((await unpackZip(zip)).has(part)).toBe(true);
  });

  it("uses separate key-data and key-encryptor salts, and both are needed", () => {
    const { info } = readEncryptedPackage(fixture(OFFICE_FILES[0][1]));
    expect(info.keyData).toBeDefined();
    expect([...info.keyData!.salt]).not.toEqual([...info.keySalt]);
    expect(info.dataIntegritySalt).toBeDefined();
    expect(info.dataIntegrityHmac).toBeDefined();
  });

  it.each(OFFICE_FILES)("%s: a wrong password is reason 'password'", async (_name, path) => {
    const bytes = fixture(path);
    expect(await verifyPackagePassword(bytes, "wrong")).toBe(false);
    expect((await failure(decryptPackage(bytes, "wrong"))).reason).toBe("password");
  });

  it("an altered package fails its data-integrity check", async () => {
    const entries = readCfb(fixture(OFFICE_FILES[0][1]));
    const tampered = entries.map(e => {
      if (e.name !== "EncryptedPackage") {
        return e;
      }
      const data = e.data.slice();
      data[100] ^= 0xff;
      return { ...e, data };
    });
    const error = await failure(decryptPackage(writeCfb(tampered), "Excel-pw1"));
    expect(error.reason).toBe("integrity");
  });
});

describe("a package this library encrypts", () => {
  const zip = new Uint8Array([0x50, 0x4b, 0x05, 0x06, ...new Array(18).fill(0)]);

  it("round-trips, with key data and key encryptor salted apart as Office writes them", async () => {
    const encrypted = await encryptPackage(zip, "pw", { spinCount: 1000 });
    expect(isEncryptedPackage(encrypted)).toBe(true);
    const { info } = readEncryptedPackage(encrypted);
    expect([...info.keyData!.salt]).not.toEqual([...info.keySalt]);
    expect(await decryptPackage(encrypted, "pw")).toEqual(zip);
    expect((await failure(decryptPackage(encrypted, "nope"))).reason).toBe("password");
  });

  it.each([128, 192, 256] as const)("encrypts and decrypts with %i-bit keys", async keyBits => {
    const encrypted = await encryptPackage(zip, "pw", { spinCount: 10, keyBits });
    expect(readEncryptedPackage(encrypted).info.keyBits).toBe(keyBits);
    expect(await decryptPackage(encrypted, "pw")).toEqual(zip);
  });

  it("refuses a declared size past the caller's bound as reason 'limit'", async () => {
    const encrypted = await encryptPackage(zip, "pw", { spinCount: 1000 });
    expect((await failure(decryptPackage(encrypted, "pw", 4))).reason).toBe("limit");
  });

  it("names the encryption it does not implement", async () => {
    const encrypted = await encryptPackage(zip, "pw", { spinCount: 1000 });
    for (const [major, minor, flags, name] of [
      [4, 2, 0x24, "Standard Encryption"],
      [3, 2, 0x04, "RC4 CryptoAPI Encryption"],
      [4, 3, 0x10, "Extensible Encryption"]
    ] as const) {
      const entries = readCfb(encrypted).map(e => {
        if (e.name !== "EncryptionInfo") {
          return e;
        }
        const data = e.data.slice();
        const view = new DataView(data.buffer);
        view.setUint16(0, major, true);
        view.setUint16(2, minor, true);
        view.setUint32(4, flags, true);
        return { ...e, data };
      });
      const error = await failure(decryptPackage(writeCfb(entries), "pw"));
      expect(error.reason).toBe("unsupported");
      expect(error.message).toContain(`${major}.${minor} (${name})`);
    }
  });

  it("writes a package larger than the header DIFAT can index (a DIFAT sector chain)", async () => {
    // 109 FAT sectors index ~7 MB; a larger package needs DIFAT sectors.
    const big = new Uint8Array(12 * 1024 * 1024);
    for (let i = 0; i < big.length; i += 4096) {
      big[i] = i & 0xff;
    }
    big.set([0x50, 0x4b, 0x03, 0x04]);
    const encrypted = await encryptPackage(big, "pw", { spinCount: 10 });
    expect(
      new DataView(encrypted.buffer, encrypted.byteOffset).getUint32(72, true)
    ).toBeGreaterThan(0);
    const decrypted = await decryptPackage(encrypted, "pw");
    // Compared as bytes: a deep equality over twelve million elements takes a minute.
    expect(Buffer.compare(decrypted, big)).toBe(0);
  });
});

describe("a package whose encryption parameters are malformed", () => {
  const zip = new Uint8Array([0x50, 0x4b, 0x05, 0x06, ...new Array(18).fill(0)]);

  /** `encrypted` with its EncryptionInfo XML edited by `xml`, and its EncryptedPackage stream by `data`. */
  function edited(
    encrypted: Uint8Array,
    xml: (text: string) => string,
    data: (bytes: Uint8Array) => Uint8Array = bytes => bytes
  ): Uint8Array {
    return writeCfb(
      readCfb(encrypted).map(entry => {
        if (entry.name === "EncryptionInfo") {
          const text = new TextDecoder().decode(entry.data.subarray(8));
          const body = new TextEncoder().encode(xml(text));
          const info = new Uint8Array(8 + body.length);
          info.set(entry.data.subarray(0, 8));
          info.set(body, 8);
          return { ...entry, data: info };
        }
        return entry.name === "EncryptedPackage" ? { ...entry, data: data(entry.data) } : entry;
      })
    );
  }

  // The size is how many bytes of the verifier and of the integrity HMAC are compared: a file free to give less would
  // have the check compare one byte.
  it.each(["keyData", "p:encryptedKey"])(
    "refuses a %s hashSize other than its hash's",
    async element => {
      const encrypted = edited(await encryptPackage(zip, "pw", { spinCount: 10 }), xml =>
        xml.replace(new RegExp(`(<${element}[^>]*?hashSize=")64"`), '$11"')
      );
      const error = await failure(decryptPackage(encrypted, "pw"));
      expect(error.reason).toBe("malformed");
      expect(error.message).toContain("hashSize");
    }
  );

  // Handed to the platform, a blob that is not whole AES blocks failed with OpenSSL's own error.
  it("refuses a key blob that is not whole AES blocks", async () => {
    const encrypted = edited(await encryptPackage(zip, "pw", { spinCount: 10 }), xml =>
      xml.replace(/encryptedKeyValue="([^"]+)"/, (_all, value: string) => {
        const bytes = Uint8Array.from(atob(value), c => c.charCodeAt(0)).subarray(1);
        return `encryptedKeyValue="${btoa(String.fromCharCode(...bytes))}"`;
      })
    );
    const error = await failure(decryptPackage(encrypted, "pw"));
    expect(error.reason).toBe("malformed");
    expect(error.message).toContain("encryptedKeyValue");
  });

  // A damaged verifier is not a wrong password: the check answered `false` for any failure.
  it("refuses a verifier that is not whole AES blocks, when only the password is checked", async () => {
    const encrypted = edited(await encryptPackage(zip, "pw", { spinCount: 10 }), xml =>
      xml.replace(/encryptedVerifierHashInput="([^"]+)"/, (_all, value: string) => {
        const bytes = Uint8Array.from(atob(value), c => c.charCodeAt(0)).subarray(1);
        return `encryptedVerifierHashInput="${btoa(String.fromCharCode(...bytes))}"`;
      })
    );
    const error = await failure(verifyPackagePassword(encrypted, "pw"));
    expect(error.reason).toBe("malformed");
    expect(error.message).toContain("encryptedVerifierHashInput");
  });

  // Node's decoder skipped the characters it did not know and went on to a wrong-password verdict
  it.each(["saltValue", "encryptedKeyValue", "encryptedHmacKey"])(
    "refuses a %s that is not base64",
    async name => {
      const encrypted = edited(await encryptPackage(zip, "pw", { spinCount: 10 }), xml =>
        xml.replace(new RegExp(`${name}="([^"]+)"`), (_all, value: string) => `${name}="!${value}"`)
      );
      const error = await failure(decryptPackage(encrypted, "pw"));
      expect(error.reason).toBe("malformed");
      expect(error.message).toContain(`${name} is not base64`);
    }
  );

  it("refuses a package cut off inside a cipher block", async () => {
    const encrypted = edited(
      await encryptPackage(new Uint8Array(9000), "pw", { spinCount: 10 }),
      xml => xml.replace(/<dataIntegrity [^>]*\/>/, ""),
      bytes => bytes.subarray(0, bytes.length - 1)
    );
    const error = await failure(decryptPackage(encrypted, "pw"));
    expect(error.reason).toBe("malformed");
    expect(error.message).toContain("not a whole number of AES blocks");
  });
});
