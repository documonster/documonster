/**
 * Tests for PDF encryption: MD5, RC4, and password/key computation.
 */
import { describe, it, expect } from "vitest";
import { md5, rc4, initEncryption, encryptData } from "@pdf/core/encryption";

describe("MD5", () => {
  it("should hash empty string correctly", () => {
    // MD5("") = d41d8cd98f00b204e9800998ecf8427e
    const result = md5(new Uint8Array(0));
    expect(toHex(result)).toBe("d41d8cd98f00b204e9800998ecf8427e");
  });

  it('should hash "abc" correctly', () => {
    // MD5("abc") = 900150983cd24fb0d6963f7d28e17f72
    const result = md5(new TextEncoder().encode("abc"));
    expect(toHex(result)).toBe("900150983cd24fb0d6963f7d28e17f72");
  });

  it('should hash "message digest" correctly', () => {
    // MD5("message digest") = f96b697d7cb7938d525a2f31aaf161d0
    const result = md5(new TextEncoder().encode("message digest"));
    expect(toHex(result)).toBe("f96b697d7cb7938d525a2f31aaf161d0");
  });

  it("should hash long string correctly", () => {
    // MD5("abcdefghijklmnopqrstuvwxyz") = c3fcd3d76192e4007dfb496cca67e13b
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
    // RC4 is symmetric: encrypting again with same key = decryption
    const decrypted = rc4(key, encrypted);
    expect(new TextDecoder().decode(decrypted)).toBe("Hello World");
  });

  it("should produce different output from input", () => {
    const key = new TextEncoder().encode("key");
    const data = new TextEncoder().encode("plaintext");
    const encrypted = rc4(key, data);
    // Encrypted should differ from plaintext
    expect(encrypted).not.toEqual(data);
  });

  it("should handle empty data", () => {
    const key = new TextEncoder().encode("key");
    const result = rc4(key, new Uint8Array(0));
    expect(result.length).toBe(0);
  });

  // Known RC4 test vector: key="Key", plaintext="Plaintext" → known ciphertext
  it("should match known test vector", () => {
    const key = new TextEncoder().encode("Key");
    const plaintext = new TextEncoder().encode("Plaintext");
    const encrypted = rc4(key, plaintext);
    // RC4("Key", "Plaintext") = BBF316E8D940AF0AD3
    expect(toHex(encrypted)).toBe("bbf316e8d940af0ad3");
  });
});

describe("initEncryption", () => {
  it("should produce 32-byte O and U values", () => {
    const state = initEncryption({
      ownerPassword: "owner",
      userPassword: "user"
    });
    expect(state.oValue.length).toBe(32);
    expect(state.uValue.length).toBe(32);
  });

  it("should produce 16-byte encryption key and file ID", () => {
    const state = initEncryption({
      ownerPassword: "test"
    });
    expect(state.encryptionKey.length).toBe(16);
    expect(state.fileId.length).toBe(16);
  });

  it("should produce different O values for different owner passwords", () => {
    const s1 = initEncryption({ ownerPassword: "pass1" });
    const s2 = initEncryption({ ownerPassword: "pass2" });
    expect(toHex(s1.oValue)).not.toBe(toHex(s2.oValue));
  });

  it("should set permission bits correctly", () => {
    const state = initEncryption({
      ownerPassword: "owner",
      permissions: { print: true, copy: true }
    });
    // Bit 3 (print) and bit 5 (copy) should be set
    expect(state.permissions & (1 << 2)).not.toBe(0); // print
    expect(state.permissions & (1 << 4)).not.toBe(0); // copy
    expect(state.permissions & (1 << 3)).toBe(0); // modify — not set
  });

  it("should handle empty user password", () => {
    const state = initEncryption({
      ownerPassword: "owner",
      userPassword: ""
    });
    expect(state.oValue.length).toBe(32);
    expect(state.uValue.length).toBe(32);
  });
});

describe("encryptData", () => {
  it("should produce different output for different object numbers", () => {
    const key = new Uint8Array(16).fill(0x42);
    const data = new TextEncoder().encode("Hello");
    const enc1 = encryptData(data, 1, 0, key);
    const enc2 = encryptData(data, 2, 0, key);
    expect(toHex(enc1)).not.toBe(toHex(enc2));
  });

  it("should be reversible (encrypt then decrypt)", () => {
    const state = initEncryption({ ownerPassword: "test" });
    const plaintext = new TextEncoder().encode("Secret Data");
    const encrypted = encryptData(plaintext, 5, 0, state.encryptionKey);
    const decrypted = encryptData(encrypted, 5, 0, state.encryptionKey);
    expect(new TextDecoder().decode(decrypted)).toBe("Secret Data");
  });
});

function toHex(bytes: Uint8Array): string {
  return Array.from(bytes)
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}
