/**
 * One spelling of a hash algorithm works everywhere in Word. `"SHA-256"` (the
 * WebCrypto spelling, and what `Security.getState` reports) is the canonical
 * form; the file-format spelling `"SHA256"` is accepted as well, and each API
 * normalises to what it writes. Before, `Security.protect` refused `"SHA256"`
 * and `WordCrypto.encrypt` refused `"SHA-256"`; a hand-built protection with
 * `"SHA256"` was written as SHA-1 (`cryptAlgorithmSid` 4), and
 * `Agile.deriveKey` read a hyphenated name as SHA-512.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError, WordCrypto } from "../crypto";
import { Document, Io, Package, Security } from "../index";

const SPELLINGS = [
  ["SHA-1", "SHA1"],
  ["SHA-256", "SHA256"],
  ["SHA-384", "SHA384"],
  ["SHA-512", "SHA512"]
] as const;

function doc() {
  const d = Document.create();
  Document.addParagraph(d, "x");
  return d;
}

describe("hash algorithm names", () => {
  it("Security.protect takes both spellings and stores the canonical one", async () => {
    for (const [canonical, compact] of SPELLINGS) {
      for (const name of [canonical, compact]) {
        const locked = await Security.protect(doc(), {
          edit: "readOnly",
          password: "pw",
          hashAlgorithm: name,
          spinCount: 2
        });
        expect(Security.getState(locked)?.hashAlgorithm, name).toBe(canonical);
        expect(await Security.verifyPassword(locked, "pw"), name).toBe(true);
      }
    }
    // Without a password there is no hash, so no algorithm is stored.
    const a = await Security.protect(doc(), { edit: "readOnly" });
    expect(Security.getState(a)?.hashAlgorithm).toBeUndefined();
  });

  it("WordCrypto.encrypt takes both spellings, and Io.read opens the result", async () => {
    const zip = await Io.toBuffer(doc());
    for (const [canonical, compact] of SPELLINGS) {
      for (const name of [canonical, compact]) {
        // SHA-1's 20-byte digest is too short for a 256-bit key.
        const keyBits = canonical === "SHA-1" ? 128 : 256;
        const enc = await WordCrypto.encrypt(zip, "pw", {
          hashAlgorithm: name,
          spinCount: 1,
          keyBits
        });
        expect(await WordCrypto.decrypt(enc, "pw"), name).toEqual(zip);
        // The file holds the format's spelling, whatever the caller wrote.
        const infoXml = new TextDecoder().decode(
          WordCrypto.Cfb.read(enc)
            .find(e => e.name === "EncryptionInfo")!
            .data.subarray(8)
        );
        expect(WordCrypto.Agile.parseEncryptionInfo(infoXml).hashAlgorithm, name).toBe(compact);
      }
    }
  });

  it("WordCrypto.Agile.deriveKey derives the same key from either spelling", async () => {
    const info = { keySalt: new Uint8Array(16).fill(7), spinCount: 3, keyBits: 128 };
    const block = WordCrypto.Agile.BLOCK_KEYS.encryptedKey;
    for (const [canonical, compact] of SPELLINGS) {
      expect(
        await WordCrypto.Agile.deriveKey("pw", { ...info, hashAlgorithm: canonical }, block),
        canonical
      ).toEqual(await WordCrypto.Agile.deriveKey("pw", { ...info, hashAlgorithm: compact }, block));
    }
    // SHA-1 and SHA-512 really differ (a hyphenated name was read as SHA-512).
    expect(
      await WordCrypto.Agile.deriveKey("pw", { ...info, hashAlgorithm: "SHA-1" }, block)
    ).not.toEqual(
      await WordCrypto.Agile.deriveKey("pw", { ...info, hashAlgorithm: "SHA-512" }, block)
    );
  });

  it("a protection set through Document.setSettings writes the algorithm it names", async () => {
    for (const [canonical, compact] of SPELLINGS) {
      const locked = await Security.protect(doc(), {
        edit: "readOnly",
        password: "pw",
        hashAlgorithm: canonical,
        spinCount: 2
      });
      const protection = locked.settings!.documentProtection!;
      const d = doc();
      Document.setSettings(d, { documentProtection: { ...protection, hashAlgorithm: compact } });
      const reread = await Io.read(await Io.toBuffer(d));
      expect(Security.getState(reread)?.hashAlgorithm, compact).toBe(canonical);
      expect(await Security.verifyPassword(reread, "pw"), compact).toBe(true);
      const xml = new TextDecoder().decode(
        (await Package.readPart(await Io.toBuffer(d), "word/settings.xml"))!
      );
      const sid = { "SHA-1": "4", "SHA-256": "12", "SHA-384": "13", "SHA-512": "14" }[canonical];
      expect(xml, compact).toContain(`w:cryptAlgorithmSid="${sid}"`);
    }
  });

  it("writes back the algorithm a file named, MD5 included", async () => {
    const locked = await Security.protect(doc(), {
      edit: "readOnly",
      password: "pw",
      spinCount: 1
    });
    const bytes = await Io.toBuffer(locked);
    const settings = new TextDecoder()
      .decode((await Package.readPart(bytes, "word/settings.xml"))!)
      .replace(/w:cryptAlgorithmSid="\d+"/, 'w:cryptAlgorithmSid="3"');
    const edited = await Package.edit(bytes, [
      { type: "replacePart", path: "word/settings.xml", data: new TextEncoder().encode(settings) }
    ]);
    const read = await Io.read(edited);
    expect(Security.getState(read)?.hashAlgorithm).toBe("MD5");
    const rewritten = new TextDecoder().decode(
      (await Package.readPart(await Io.toBuffer(read), "word/settings.xml"))!
    );
    // It was written as SHA-1 (sid 4), which no longer matches the stored hash.
    expect(rewritten).toContain('w:cryptAlgorithmSid="3"');
  });

  it("refuses a name that is neither spelling, saying both are accepted", async () => {
    const protect = await Security.protect(doc(), {
      edit: "readOnly",
      password: "pw",
      hashAlgorithm: "sha256" as never
    }).catch((e: unknown) => e);
    expect(protect).toBeInstanceOf(DocxArgumentError);
    expect((protect as Error).message).toContain('"SHA-256"');
    const encrypt = await WordCrypto.encrypt(await Io.toBuffer(doc()), "pw", {
      hashAlgorithm: "MD5" as never
    }).catch((e: unknown) => e);
    expect(encrypt).toBeInstanceOf(DocxArgumentError);
    expect((encrypt as DocxArgumentError).field).toBe("options.hashAlgorithm");
  });
});
