/**
 * DOCX Module - Encryption Tests
 *
 * Tests for encrypt/decrypt round-trip and utility functions.
 */

import { describe, it, expect } from "vitest";

import { readDocx } from "../reader/docx-reader";
import { encryptDocx, isEncryptedDocx, decryptDocx } from "../security/encryption";
import type { DocxDocument } from "../types";
import { packageDocx } from "../writer/docx-packager";

// Create a minimal valid DOCX for testing
function createMinimalDoc(): DocxDocument {
  return {
    body: [
      {
        type: "paragraph",
        children: [{ content: [{ type: "text", text: "Encryption test content" }] }]
      }
    ],
    contentTypes: [
      {
        partName: "/word/document.xml",
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
      }
    ]
  } as unknown as DocxDocument;
}

describe("isEncryptedDocx", () => {
  it("returns false for normal ZIP/DOCX", async () => {
    const doc = createMinimalDoc();
    const buffer = await packageDocx(doc);
    expect(isEncryptedDocx(buffer)).toBe(false);
  });

  it("returns true for CFB-signed data", () => {
    // CFB magic number: D0 CF 11 E0 A1 B1 1A E1
    const cfbHeader = new Uint8Array([
      0xd0,
      0xcf,
      0x11,
      0xe0,
      0xa1,
      0xb1,
      0x1a,
      0xe1,
      ...new Array(504).fill(0)
    ]);
    expect(isEncryptedDocx(cfbHeader)).toBe(true);
  });

  it("returns false for empty buffer", () => {
    expect(isEncryptedDocx(new Uint8Array(0))).toBe(false);
  });

  it("returns false for small buffer", () => {
    expect(isEncryptedDocx(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false);
  });
});

describe("encryptDocx / decryptDocx", () => {
  it("encrypts and decrypts round-trip", async () => {
    const doc = createMinimalDoc();
    const original = await packageDocx(doc);
    const password = "testPassword123";

    const encrypted = await encryptDocx(original, password);
    expect(encrypted).toBeInstanceOf(Uint8Array);
    expect(isEncryptedDocx(encrypted)).toBe(true);

    // Encrypted should be different from original
    expect(encrypted.length).not.toBe(original.length);

    const decrypted = await decryptDocx(encrypted, password);
    expect(decrypted).toBeInstanceOf(Uint8Array);
    // Decrypted should be a valid ZIP (DOCX)
    expect(decrypted[0]).toBe(0x50); // P
    expect(decrypted[1]).toBe(0x4b); // K
  }, 30000);

  it("encrypts with custom options (256-bit AES)", async () => {
    const doc = createMinimalDoc();
    const original = await packageDocx(doc);
    const password = "securePass!";

    const encrypted = await encryptDocx(original, password, {
      keyBits: 256,
      hashAlgorithm: "SHA512"
    });
    expect(isEncryptedDocx(encrypted)).toBe(true);

    const decrypted = await decryptDocx(encrypted, password);
    expect(decrypted[0]).toBe(0x50);
    expect(decrypted[1]).toBe(0x4b);
  }, 30000);

  it("decryption with wrong password fails", async () => {
    const doc = createMinimalDoc();
    const original = await packageDocx(doc);
    const password = "correct";

    const encrypted = await encryptDocx(original, password);

    await expect(decryptDocx(encrypted, "wrong")).rejects.toThrow();
  }, 30000);

  it("readDocx can read encrypted document with password", async () => {
    const doc = createMinimalDoc();
    const original = await packageDocx(doc);
    const password = "readTest";

    const encrypted = await encryptDocx(original, password);
    const parsed = await readDocx(encrypted, { password });

    expect(parsed.body.length).toBeGreaterThan(0);
    expect(parsed.body[0].type).toBe("paragraph");
  }, 30000);
});
