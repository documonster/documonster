/**
 * Tests for PDF encryption: crypto primitives and V=5/R=5 AES-256 encryption.
 */
import { describe, it, expect } from "vitest";
import { md5, rc4, sha256, aesCbcEncrypt, aesCbcDecrypt } from "@pdf/core/crypto";
import { initEncryption, encryptData } from "@pdf/core/encryption";
import { readPdf } from "../reader/pdf-reader";
import { pdf } from "../pdf";

describe("MD5", () => {
  it("should hash empty string correctly", () => {
    const result = md5(new Uint8Array(0));
    expect(toHex(result)).toBe("d41d8cd98f00b204e9800998ecf8427e");
  });

  it('should hash "abc" correctly', () => {
    const result = md5(new TextEncoder().encode("abc"));
    expect(toHex(result)).toBe("900150983cd24fb0d6963f7d28e17f72");
  });

  it('should hash "message digest" correctly', () => {
    const result = md5(new TextEncoder().encode("message digest"));
    expect(toHex(result)).toBe("f96b697d7cb7938d525a2f31aaf161d0");
  });

  it("should hash long string correctly", () => {
    const result = md5(new TextEncoder().encode("abcdefghijklmnopqrstuvwxyz"));
    expect(toHex(result)).toBe("c3fcd3d76192e4007dfb496cca67e13b");
  });

  it("should produce 16-byte output", () => {
    const result = md5(new TextEncoder().encode("test"));
    expect(result.length).toBe(16);
  });
});

describe("RC4", () => {
  it("should encrypt and decrypt symmetrically", () => {
    const key = new TextEncoder().encode("secret");
    const plaintext = new TextEncoder().encode("Hello World");
    const encrypted = rc4(key, plaintext);
    const decrypted = rc4(key, encrypted);
    expect(new TextDecoder().decode(decrypted)).toBe("Hello World");
  });

  it("should produce different output from input", () => {
    const key = new TextEncoder().encode("key");
    const data = new TextEncoder().encode("plaintext");
    const encrypted = rc4(key, data);
    expect(encrypted).not.toEqual(data);
  });

  it("should handle empty data", () => {
    const key = new TextEncoder().encode("key");
    const result = rc4(key, new Uint8Array(0));
    expect(result.length).toBe(0);
  });

  it("should match known test vector", () => {
    const key = new TextEncoder().encode("Key");
    const plaintext = new TextEncoder().encode("Plaintext");
    const encrypted = rc4(key, plaintext);
    expect(toHex(encrypted)).toBe("bbf316e8d940af0ad3");
  });
});

describe("AES-256 CBC encrypt/decrypt roundtrip", () => {
  it("should encrypt and decrypt a single block", () => {
    const key = new Uint8Array(32);
    key.fill(0x42);
    const iv = new Uint8Array(16);
    const plaintext = new TextEncoder().encode("Hello AES-256!  "); // exactly 16 bytes

    const encrypted = aesCbcEncrypt(plaintext, key, iv);
    expect(encrypted.length).toBe(32); // 16 data + 16 PKCS#7 padding block

    const decrypted = aesCbcDecrypt(encrypted, key, iv);
    expect(new TextDecoder().decode(decrypted)).toBe("Hello AES-256!  ");
  });

  it("should handle multi-block data", () => {
    const key = new Uint8Array(32);
    for (let i = 0; i < 32; i++) {
      key[i] = i;
    }
    const iv = new Uint8Array(16).fill(0xff);
    const plaintext = new TextEncoder().encode("A".repeat(100));

    const encrypted = aesCbcEncrypt(plaintext, key, iv);
    const decrypted = aesCbcDecrypt(encrypted, key, iv);
    expect(new TextDecoder().decode(decrypted)).toBe("A".repeat(100));
  });

  it("should handle empty data", () => {
    const key = new Uint8Array(32);
    const iv = new Uint8Array(16);
    const encrypted = aesCbcEncrypt(new Uint8Array(0), key, iv);
    const decrypted = aesCbcDecrypt(encrypted, key, iv);
    expect(decrypted.length).toBe(0);
  });
});

describe("initEncryption (V=5, R=5, AES-256)", () => {
  it("should produce 48-byte O and U values", () => {
    const state = initEncryption({
      ownerPassword: "owner",
      userPassword: "user"
    });
    expect(state.oValue.length).toBe(48);
    expect(state.uValue.length).toBe(48);
  });

  it("should produce 32-byte encryption key, OE, UE", () => {
    const state = initEncryption({
      ownerPassword: "test"
    });
    expect(state.encryptionKey.length).toBe(32);
    expect(state.oeValue.length).toBe(32);
    expect(state.ueValue.length).toBe(32);
  });

  it("should produce 16-byte file ID and Perms", () => {
    const state = initEncryption({
      ownerPassword: "test"
    });
    expect(state.fileId.length).toBe(16);
    expect(state.permsValue.length).toBe(16);
  });

  it("should produce different values for different passwords", () => {
    const s1 = initEncryption({ ownerPassword: "pass1" });
    const s2 = initEncryption({ ownerPassword: "pass2" });
    expect(toHex(s1.oValue)).not.toBe(toHex(s2.oValue));
    expect(toHex(s1.uValue)).not.toBe(toHex(s2.uValue));
  });

  it("should set permission bits correctly", () => {
    const state = initEncryption({
      ownerPassword: "owner",
      permissions: { print: true, copy: true }
    });
    expect(state.permissions & (1 << 2)).not.toBe(0); // print
    expect(state.permissions & (1 << 4)).not.toBe(0); // copy
    expect(state.permissions & (1 << 3)).toBe(0); // modify — not set
  });

  it("should handle empty user password", () => {
    const state = initEncryption({
      ownerPassword: "owner",
      userPassword: ""
    });
    expect(state.oValue.length).toBe(48);
    expect(state.uValue.length).toBe(48);
  });
});

describe("encryptData (AES-256 with IV prefix)", () => {
  it("should prepend 16-byte IV to ciphertext", () => {
    const key = new Uint8Array(32).fill(0x42);
    const data = new TextEncoder().encode("Hello");
    const encrypted = encryptData(data, 1, 0, key);
    // IV (16) + ciphertext (at least 16 for one padded block)
    expect(encrypted.length).toBe(16 + 16);
  });

  it("should produce different ciphertexts due to random IV", () => {
    const key = new Uint8Array(32).fill(0x42);
    const data = new TextEncoder().encode("Hello");
    const enc1 = encryptData(data, 1, 0, key);
    const enc2 = encryptData(data, 1, 0, key);
    // Random IVs make each encryption different
    expect(toHex(enc1)).not.toBe(toHex(enc2));
  });

  it("should be decryptable with aesCbcDecrypt", () => {
    const key = new Uint8Array(32).fill(0x42);
    const plaintext = new TextEncoder().encode("Secret Data for AES-256");
    const encrypted = encryptData(plaintext, 5, 0, key);
    // Extract IV and ciphertext
    const iv = encrypted.subarray(0, 16);
    const ciphertext = encrypted.subarray(16);
    const decrypted = aesCbcDecrypt(ciphertext, key, iv);
    expect(new TextDecoder().decode(decrypted)).toBe("Secret Data for AES-256");
  });
});

describe("AES-256 writer → reader roundtrip", () => {
  it("should write AES-256 encrypted PDF and read it back", async () => {
    const pdfBytes = await pdf(
      [
        ["Name", "Value"],
        ["Secret", 42],
        ["Confidential", 99.9]
      ],
      {
        title: "AES-256 Test",
        author: "Test Bot",
        encryption: { ownerPassword: "owner256", userPassword: "user256" }
      }
    );

    // Verify it's PDF 2.0
    const header = new TextDecoder().decode(pdfBytes.subarray(0, 10));
    expect(header).toContain("%PDF-2.0");

    // Read with user password
    const result = await readPdf(pdfBytes, { password: "user256" });
    expect(result.metadata.encrypted).toBe(true);
    expect(result.text).toContain("Secret");
    expect(result.text).toContain("42");
    expect(result.text).toContain("Confidential");
    expect(result.text).toContain("99.9");
    expect(result.metadata.title).toBe("AES-256 Test");
    expect(result.metadata.author).toBe("Test Bot");

    // Read with owner password
    const ownerResult = await readPdf(pdfBytes, { password: "owner256" });
    expect(ownerResult.text).toBe(result.text);
  });

  it("should reject wrong password", async () => {
    const pdfBytes = await pdf([["Test"]], {
      encryption: { ownerPassword: "owner", userPassword: "user" }
    });
    await expect(readPdf(pdfBytes, { password: "wrong" })).rejects.toThrow();
  });

  it("should read owner-only encrypted PDF without password", async () => {
    const pdfBytes = await pdf([["Public Data"]], {
      encryption: { ownerPassword: "admin" }
    });
    // Empty user password → should open without providing a password
    const result = await readPdf(pdfBytes);
    expect(result.text).toContain("Public Data");
  });
});

describe("SHA-256", () => {
  it('should hash "abc" correctly (NIST FIPS 180-4)', () => {
    const input = new TextEncoder().encode("abc");
    expect(toHex(sha256(input))).toBe(
      "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    );
  });

  it("should hash empty string correctly", () => {
    expect(toHex(sha256(new Uint8Array(0)))).toBe(
      "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    );
  });
});

describe("AES-CBC Encrypt — NIST SP 800-38A test vectors", () => {
  it("AES-128-CBC: encrypt should produce correct ciphertext (F.2.1)", () => {
    const key = hex("2b7e151628aed2a6abf7158809cf4f3c");
    const iv = hex("000102030405060708090a0b0c0d0e0f");
    const plaintext = hex("6bc1bee22e409f96e93d7e117393172a");
    const expectedCipher = hex("7649abac8119b246cee98e9b12e9197d");

    // aesCbcEncrypt adds PKCS#7 padding, so output is 2 blocks for 1-block input
    const encrypted = aesCbcEncrypt(plaintext, key, iv);
    // First block should match NIST ciphertext exactly
    expect(toHex(encrypted.subarray(0, 16))).toBe(toHex(expectedCipher));
  });

  it("AES-256-CBC: encrypt should produce correct ciphertext (F.2.5)", () => {
    const key = hex("603deb1015ca71be2b73aef0857d77811f352c073b6108d72d9810a30914dff4");
    const iv = hex("000102030405060708090a0b0c0d0e0f");
    const plaintext = hex("6bc1bee22e409f96e93d7e117393172a");
    const expectedCipher = hex("f58c4c04d6e5f1ba779eabfb5f7bfbd6");

    const encrypted = aesCbcEncrypt(plaintext, key, iv);
    expect(toHex(encrypted.subarray(0, 16))).toBe(toHex(expectedCipher));
  });

  it("AES-256-CBC: encrypt then decrypt should roundtrip with NIST data", () => {
    const key = hex("603deb1015ca71be2b73aef0857d77811f352c073b6108d72d9810a30914dff4");
    const iv = hex("000102030405060708090a0b0c0d0e0f");
    const plaintext = hex(
      "6bc1bee22e409f96e93d7e117393172a" +
        "ae2d8a571e03ac9c9eb76fac45af8e51" +
        "30c81c46a35ce411e5fbc1191a0a52ef"
    );

    const encrypted = aesCbcEncrypt(plaintext, key, iv);
    const decrypted = aesCbcDecrypt(encrypted, key, iv);
    expect(toHex(decrypted)).toBe(toHex(plaintext));
  });
});

describe("PDF 2.0 Encrypt dictionary structure", () => {
  it("should contain V=5, R=5, AESV3 in encrypted output", async () => {
    const pdfBytes = await pdf([["Structure Test"]], {
      encryption: { ownerPassword: "owner", userPassword: "user" }
    });
    const text = new TextDecoder("latin1").decode(pdfBytes);

    expect(text).toContain("/V 5");
    expect(text).toContain("/R 5");
    expect(text).toContain("/Length 256");
    expect(text).toContain("AESV3");
    expect(text).toContain("/StmF /StdCF");
    expect(text).toContain("/StrF /StdCF");
    expect(text).toContain("/OE");
    expect(text).toContain("/UE");
    expect(text).toContain("/Perms");
    expect(text).toContain("/EncryptMetadata true");
  });

  it("should produce PDF 2.0 header", async () => {
    const pdfBytes = await pdf([["Version Test"]], {
      encryption: { ownerPassword: "owner" }
    });
    const header = new TextDecoder().decode(pdfBytes.subarray(0, 10));
    expect(header).toContain("%PDF-2.0");
  });

  it("should produce PDF 2.0 even without encryption", async () => {
    const pdfBytes = await pdf([["No Encryption"]]);
    const header = new TextDecoder().decode(pdfBytes.subarray(0, 10));
    expect(header).toContain("%PDF-2.0");
  });
});

function toHex(bytes: Uint8Array): string {
  return Array.from(bytes)
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

function hex(s: string): Uint8Array {
  const bytes = new Uint8Array(s.length / 2);
  for (let i = 0; i < bytes.length; i++) {
    bytes[i] = parseInt(s.substring(i * 2, i * 2 + 2), 16);
  }
  return bytes;
}
