/**
 * Unit tests for ZIP encryption/decryption (ZipCrypto and AES)
 */

import { describe, it, expect } from "vitest";

import { crc32 } from "../../compression/crc32";
import {
  aesDerive,
  aesDecrypt,
  aesEncrypt,
  aesCtr,
  aesComputeHmac,
  aesCheckPasswordOnly,
  aesCheckSignature,
  parseAesExtraField,
  AES_KEY_LENGTH,
  AES_SALT_LENGTH,
  AES_AUTH_CODE_LENGTH
} from "../../crypto/aes";
import type { AesKeyStrength } from "../../crypto/aes";
import {
  zipCryptoInitKeys,
  zipCryptoDecryptByte,
  zipCryptoEncryptByte,
  zipCryptoDecrypt,
  zipCryptoEncrypt,
  zipCryptoCheckPassword,
  ZIP_CRYPTO_HEADER_SIZE
} from "../../crypto/zip-crypto";

/**
 * Helper to generate random bytes using crypto.getRandomValues
 * in chunks (max 65536 bytes per call).
 */
function getRandomBytes(length: number): Uint8Array {
  const result = new Uint8Array(length);
  const chunkSize = 65536;
  for (let offset = 0; offset < length; offset += chunkSize) {
    const size = Math.min(chunkSize, length - offset);
    const chunk = new Uint8Array(size);
    crypto.getRandomValues(chunk);
    result.set(chunk, offset);
  }
  return result;
}

/**
 * Create a deterministic PRNG for reproducible tests.
 * Uses a simple LCG (Linear Congruential Generator) algorithm.
 *
 * This is useful for ZipCrypto tests where header verification only checks
 * 1 byte (1/256 false positive rate), so we need deterministic bytes
 * that won't accidentally pass with the wrong password.
 */
function createDeterministicRandom(initialSeed: number = 12345): (length: number) => Uint8Array {
  let seed = initialSeed;
  return (length: number): Uint8Array => {
    const result = new Uint8Array(length);
    for (let i = 0; i < length; i++) {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      result[i] = seed & 0xff;
    }
    return result;
  };
}

describe("ZipCrypto", () => {
  describe("key initialization", () => {
    it("should initialize keys with known password", () => {
      const password = "test";
      const keys = zipCryptoInitKeys(password);

      expect(keys.key0).toBeTypeOf("number");
      expect(keys.key1).toBeTypeOf("number");
      expect(keys.key2).toBeTypeOf("number");
      // Keys should be in 32-bit range
      expect(keys.key0 >>> 0).toBe(keys.key0);
      expect(keys.key1 >>> 0).toBe(keys.key1);
      expect(keys.key2 >>> 0).toBe(keys.key2);
    });

    it("should accept Uint8Array as password", () => {
      const passwordStr = "test";
      const passwordBytes = new TextEncoder().encode(passwordStr);

      const keysStr = zipCryptoInitKeys(passwordStr);
      const keysBytes = zipCryptoInitKeys(passwordBytes);

      expect(keysStr).toEqual(keysBytes);
    });

    it("should produce different keys for different passwords", () => {
      const keys1 = zipCryptoInitKeys("password1");
      const keys2 = zipCryptoInitKeys("password2");

      expect(keys1.key0).not.toBe(keys2.key0);
      expect(keys1.key1).not.toBe(keys2.key1);
      expect(keys1.key2).not.toBe(keys2.key2);
    });
  });

  describe("encryption/decryption roundtrip", () => {
    it("should encrypt and decrypt single bytes", () => {
      const password = "secret";
      const testByte = 0x42;

      const encKeys = zipCryptoInitKeys(password);
      const encrypted = zipCryptoEncryptByte(encKeys, testByte);

      const decKeys = zipCryptoInitKeys(password);
      const decrypted = zipCryptoDecryptByte(decKeys, encrypted);

      expect(decrypted).toBe(testByte);
    });

    it("should encrypt and decrypt buffers", () => {
      const password = "mysecretpassword";
      const plaintext = new TextEncoder().encode("Hello, World! This is a test message.");
      const crcValue = crc32(plaintext);

      // Encrypt
      const encrypted = zipCryptoEncrypt(plaintext, password, crcValue, getRandomBytes);

      // Encrypted data should include 12-byte header
      expect(encrypted.length).toBe(plaintext.length + ZIP_CRYPTO_HEADER_SIZE);

      // Decrypt
      const decrypted = zipCryptoDecrypt(encrypted, password, crcValue);

      expect(decrypted).toEqual(plaintext);
    });

    it("should handle empty data", () => {
      const password = "test";
      const plaintext = new Uint8Array(0);
      const crcValue = crc32(plaintext);

      const encrypted = zipCryptoEncrypt(plaintext, password, crcValue, getRandomBytes);
      const decrypted = zipCryptoDecrypt(encrypted, password, crcValue);

      expect(decrypted).toEqual(plaintext);
    });

    it("should handle binary data with all byte values", () => {
      const password = "binary-test";
      const plaintext = new Uint8Array(256);
      for (let i = 0; i < 256; i++) {
        plaintext[i] = i;
      }
      const crcValue = crc32(plaintext);

      const encrypted = zipCryptoEncrypt(plaintext, password, crcValue, getRandomBytes);
      const decrypted = zipCryptoDecrypt(encrypted, password, crcValue);

      expect(decrypted).toEqual(plaintext);
    });

    it("should handle large data", () => {
      const password = "large-test";
      const plaintext = new Uint8Array(100000);
      for (let i = 0; i < plaintext.length; i++) {
        plaintext[i] = i % 256;
      }
      const crcValue = crc32(plaintext);

      const encrypted = zipCryptoEncrypt(plaintext, password, crcValue, getRandomBytes);
      const decrypted = zipCryptoDecrypt(encrypted, password, crcValue);

      expect(decrypted).toEqual(plaintext);
    });

    it("should return null with wrong password", () => {
      const password = "correct-password";
      const wrongPassword = "wrong-password";
      const plaintext = new TextEncoder().encode("Secret data");
      const crcValue = crc32(plaintext);

      const encrypted = zipCryptoEncrypt(
        plaintext,
        password,
        crcValue,
        createDeterministicRandom()
      );
      const decrypted = zipCryptoDecrypt(encrypted, wrongPassword, crcValue);

      // Should return null because header verification fails
      expect(decrypted).toBeNull();
    });
  });

  describe("header size", () => {
    it("should have correct header size constant", () => {
      expect(ZIP_CRYPTO_HEADER_SIZE).toBe(12);
    });
  });
});

describe("AES Encryption", () => {
  describe("key derivation", () => {
    it("should derive keys for AES-256", async () => {
      const password = "test-password";
      const salt = getRandomBytes(16);

      const keys = await aesDerive(password, salt, 256); // AES-256

      expect(keys.encryptionKey).toHaveLength(32); // 256 bits
      expect(keys.hmacKey).toHaveLength(32);
      expect(keys.passwordVerify).toHaveLength(2);
    });

    it("should derive keys for AES-192", async () => {
      const password = "test-password";
      const salt = getRandomBytes(12);

      const keys = await aesDerive(password, salt, 192); // AES-192

      expect(keys.encryptionKey).toHaveLength(24); // 192 bits
      expect(keys.hmacKey).toHaveLength(24); // WinZip AES: HMAC key length == AES key length
      expect(keys.passwordVerify).toHaveLength(2);
    });

    it("should derive keys for AES-128", async () => {
      const password = "test-password";
      const salt = getRandomBytes(8);

      const keys = await aesDerive(password, salt, 128); // AES-128

      expect(keys.encryptionKey).toHaveLength(16); // 128 bits
      expect(keys.hmacKey).toHaveLength(16); // WinZip AES: HMAC key length == AES key length
      expect(keys.passwordVerify).toHaveLength(2);
    });

    it("should produce same keys for same password and salt", async () => {
      const password = "deterministic-test";
      const salt = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]);

      const keys1 = await aesDerive(password, salt, 256);
      const keys2 = await aesDerive(password, salt, 256);

      expect(keys1.encryptionKey).toEqual(keys2.encryptionKey);
      expect(keys1.hmacKey).toEqual(keys2.hmacKey);
      expect(keys1.passwordVerify).toEqual(keys2.passwordVerify);
    });

    it("should produce different keys for different passwords", async () => {
      const salt = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]);

      const keys1 = await aesDerive("password1", salt, 256);
      const keys2 = await aesDerive("password2", salt, 256);

      expect(keys1.encryptionKey).not.toEqual(keys2.encryptionKey);
      expect(keys1.hmacKey).not.toEqual(keys2.hmacKey);
      expect(keys1.passwordVerify).not.toEqual(keys2.passwordVerify);
    });

    it("should produce different keys for different salts", async () => {
      const password = "same-password";
      const salt1 = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]);
      const salt2 = new Uint8Array([16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]);

      const keys1 = await aesDerive(password, salt1, 256);
      const keys2 = await aesDerive(password, salt2, 256);

      expect(keys1.encryptionKey).not.toEqual(keys2.encryptionKey);
    });

    it("should accept Uint8Array as password", async () => {
      const passwordStr = "test";
      const passwordBytes = new TextEncoder().encode(passwordStr);
      const salt = new Uint8Array(16);

      const keys1 = await aesDerive(passwordStr, salt, 256);
      const keys2 = await aesDerive(passwordBytes, salt, 256);

      expect(keys1.encryptionKey).toEqual(keys2.encryptionKey);
    });
  });

  describe("AES-CTR encryption/decryption", () => {
    it("should encrypt and decrypt data", async () => {
      const key = getRandomBytes(32);
      const plaintext = new TextEncoder().encode("Hello, AES-CTR!");

      const encrypted = await aesCtr(key, plaintext, true);
      const decrypted = await aesCtr(key, encrypted, false);

      expect(decrypted).toEqual(plaintext);
    });

    it("should handle empty data", async () => {
      const key = getRandomBytes(32);
      const plaintext = new Uint8Array(0);

      const encrypted = await aesCtr(key, plaintext, true);
      const decrypted = await aesCtr(key, encrypted, false);

      expect(decrypted).toEqual(plaintext);
    });

    it("should handle large data", async () => {
      const key = getRandomBytes(32);
      const plaintext = getRandomBytes(100000);

      const encrypted = await aesCtr(key, plaintext, true);
      const decrypted = await aesCtr(key, encrypted, false);

      expect(decrypted).toEqual(plaintext);
    });
  });

  describe("HMAC computation", () => {
    it("should compute 10-byte HMAC", async () => {
      const key = getRandomBytes(32);
      const data = new TextEncoder().encode("Data to authenticate");

      const hmac = await aesComputeHmac(key, data);

      expect(hmac).toHaveLength(10); // Truncated to 10 bytes per WinZip spec
    });

    it("should produce same HMAC for same inputs", async () => {
      const key = new Uint8Array([
        1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25,
        26, 27, 28, 29, 30, 31, 32
      ]);
      const data = new TextEncoder().encode("Consistent data");

      const hmac1 = await aesComputeHmac(key, data);
      const hmac2 = await aesComputeHmac(key, data);

      expect(hmac1).toEqual(hmac2);
    });

    it("should produce different HMAC for different data", async () => {
      const key = getRandomBytes(32);
      const data1 = new TextEncoder().encode("Data 1");
      const data2 = new TextEncoder().encode("Data 2");

      const hmac1 = await aesComputeHmac(key, data1);
      const hmac2 = await aesComputeHmac(key, data2);

      expect(hmac1).not.toEqual(hmac2);
    });
  });

  describe("full encryption/decryption roundtrip", () => {
    const testCases: Array<{ name: string; keyStrength: AesKeyStrength }> = [
      { name: "AES-256", keyStrength: 256 },
      { name: "AES-192", keyStrength: 192 },
      { name: "AES-128", keyStrength: 128 }
    ];

    for (const tc of testCases) {
      it(`should encrypt and decrypt with ${tc.name}`, async () => {
        const password = "test-password-123";
        const plaintext = new TextEncoder().encode("This is a test message for AES encryption.");

        const encrypted = await aesEncrypt(plaintext, password, tc.keyStrength);

        // Verify encrypted data has correct structure
        const saltLen = AES_SALT_LENGTH[tc.keyStrength];
        expect(encrypted.length).toBeGreaterThan(saltLen + 2 + AES_AUTH_CODE_LENGTH);

        const decrypted = await aesDecrypt(encrypted, password, tc.keyStrength);

        expect(decrypted).toEqual(plaintext);
      });
    }

    it("should fail with wrong password", async () => {
      const password = "correct-password";
      const wrongPassword = "wrong-password";
      const plaintext = new TextEncoder().encode("Secret data");

      const encrypted = await aesEncrypt(plaintext, password, 256);

      await expect(aesDecrypt(encrypted, wrongPassword, 256)).rejects.toThrow();
    });

    it("should detect data tampering via HMAC", async () => {
      const password = "test-password";
      const plaintext = new TextEncoder().encode("Important data");

      const encrypted = await aesEncrypt(plaintext, password, 256);

      // Tamper with encrypted data (modify a byte in the middle)
      const tamperedIndex = AES_SALT_LENGTH[256] + 2 + Math.floor(encrypted.length / 2);
      encrypted[tamperedIndex] ^= 0xff;

      await expect(aesDecrypt(encrypted, password, 256)).rejects.toThrow(
        "HMAC verification failed"
      );
    });

    it("should handle empty data", async () => {
      const password = "empty-test";
      const plaintext = new Uint8Array(0);

      const encrypted = await aesEncrypt(plaintext, password, 256);
      const decrypted = await aesDecrypt(encrypted, password, 256);

      expect(decrypted).toEqual(plaintext);
    });

    it("should handle binary data with all byte values", async () => {
      const password = "binary-test";
      const plaintext = new Uint8Array(256);
      for (let i = 0; i < 256; i++) {
        plaintext[i] = i;
      }

      const encrypted = await aesEncrypt(plaintext, password, 256);
      const decrypted = await aesDecrypt(encrypted, password, 256);

      expect(decrypted).toEqual(plaintext);
    });

    it("should handle large data (1MB)", async () => {
      const password = "large-test";
      const plaintext = getRandomBytes(1024 * 1024);

      const encrypted = await aesEncrypt(plaintext, password, 256);
      const decrypted = await aesDecrypt(encrypted, password, 256);

      expect(decrypted).toEqual(plaintext);
    });
  });

  describe("AES extra field parsing", () => {
    it("should parse valid AES extra field", () => {
      // AES extra field: version=2, vendor=AE, keyStrength=3, compression=8
      const data = new Uint8Array([
        0x02,
        0x00, // Version 2 (AE-2)
        0x41,
        0x45, // "AE" vendor ID
        0x03, // Key strength 3 (AES-256)
        0x08,
        0x00 // Original compression method (DEFLATE)
      ]);

      const info = parseAesExtraField(data);

      expect(info).not.toBeNull();
      expect(info!.version).toBe(2);
      expect(info!.vendorId).toBe(0x4541); // "AE"
      expect(info!.keyStrength).toBe(256); // Converted from byte 3 to 256
      expect(info!.compressionMethod).toBe(8);
    });

    it("should parse AE-1 version", () => {
      const data = new Uint8Array([
        0x01,
        0x00, // Version 1 (AE-1)
        0x41,
        0x45, // "AE" vendor ID
        0x01, // Key strength 1 (AES-128)
        0x00,
        0x00 // Original compression method (STORE)
      ]);

      const info = parseAesExtraField(data);

      expect(info).not.toBeNull();
      expect(info!.version).toBe(1);
      expect(info!.keyStrength).toBe(128); // Converted from byte 1 to 128
      expect(info!.compressionMethod).toBe(0);
    });

    it("should return null for invalid vendor ID", () => {
      const data = new Uint8Array([
        0x02,
        0x00,
        0x00,
        0x00, // Invalid vendor ID
        0x03,
        0x08,
        0x00
      ]);

      const info = parseAesExtraField(data);

      expect(info).toBeNull();
    });

    it("should return null for too short data", () => {
      const data = new Uint8Array([0x02, 0x00, 0x41, 0x45, 0x03]);

      const info = parseAesExtraField(data);

      expect(info).toBeNull();
    });

    it("should return null for invalid key strength", () => {
      const data = new Uint8Array([
        0x02,
        0x00,
        0x41,
        0x45,
        0x04, // Invalid key strength
        0x08,
        0x00
      ]);

      const info = parseAesExtraField(data);

      expect(info).toBeNull();
    });
  });

  describe("constants", () => {
    it("should have correct key lengths", () => {
      expect(AES_KEY_LENGTH[128]).toBe(16); // AES-128
      expect(AES_KEY_LENGTH[192]).toBe(24); // AES-192
      expect(AES_KEY_LENGTH[256]).toBe(32); // AES-256
    });

    it("should have correct salt lengths", () => {
      expect(AES_SALT_LENGTH[128]).toBe(8); // AES-128
      expect(AES_SALT_LENGTH[192]).toBe(12); // AES-192
      expect(AES_SALT_LENGTH[256]).toBe(16); // AES-256
    });

    it("should have correct auth code length", () => {
      expect(AES_AUTH_CODE_LENGTH).toBe(10);
    });
  });

  describe("checkPasswordOnly", () => {
    it("should return true for correct password", async () => {
      const password = "test-password";
      const plaintext = new TextEncoder().encode("Test data");

      const encrypted = await aesEncrypt(plaintext, password, 256);
      const result = await aesCheckPasswordOnly(encrypted, password, 256);

      expect(result).toBe(true);
    });

    it("should return false for incorrect password", async () => {
      const password = "correct-password";
      const wrongPassword = "wrong-password";
      const plaintext = new TextEncoder().encode("Test data");

      const encrypted = await aesEncrypt(plaintext, password, 256);
      const result = await aesCheckPasswordOnly(encrypted, wrongPassword, 256);

      expect(result).toBe(false);
    });

    it("should return false for too short data", async () => {
      const shortData = new Uint8Array(5);
      const result = await aesCheckPasswordOnly(shortData, "password", 256);

      expect(result).toBe(false);
    });

    it("should work with all key strengths", async () => {
      const password = "test";
      const plaintext = new TextEncoder().encode("Data");

      for (const keyStrength of [128, 192, 256] as const) {
        const encrypted = await aesEncrypt(plaintext, password, keyStrength);
        const result = await aesCheckPasswordOnly(encrypted, password, keyStrength);
        expect(result).toBe(true);
      }
    });
  });

  describe("checkSignature", () => {
    it("should return true for valid signature", async () => {
      const password = "test-password";
      const plaintext = new TextEncoder().encode("Test data");

      const encrypted = await aesEncrypt(plaintext, password, 256);
      const result = await aesCheckSignature(encrypted, password, 256);

      expect(result).toBe(true);
    });

    it("should throw for incorrect password", async () => {
      const password = "correct-password";
      const wrongPassword = "wrong-password";
      const plaintext = new TextEncoder().encode("Test data");

      const encrypted = await aesEncrypt(plaintext, password, 256);

      await expect(aesCheckSignature(encrypted, wrongPassword, 256)).rejects.toThrow(
        "Password verification failed"
      );
    });

    it("should return false for tampered data", async () => {
      const password = "test-password";
      const plaintext = new TextEncoder().encode("Test data");

      const encrypted = await aesEncrypt(plaintext, password, 256);

      // Tamper with the ciphertext (not the HMAC)
      const tamperedIndex = AES_SALT_LENGTH[256] + 2 + 1;
      encrypted[tamperedIndex] ^= 0xff;

      const result = await aesCheckSignature(encrypted, password, 256);
      expect(result).toBe(false);
    });

    it("should throw for too short data", async () => {
      const shortData = new Uint8Array(5);

      await expect(aesCheckSignature(shortData, "password", 256)).rejects.toThrow(
        "Encrypted data too short"
      );
    });
  });
});

describe("ZipCrypto Password Check", () => {
  it("should return true for correct password", () => {
    const password = "test-password";
    const plaintext = new TextEncoder().encode("Test data");
    const crcValue = crc32(plaintext);

    const encrypted = zipCryptoEncrypt(plaintext, password, crcValue, getRandomBytes);
    const result = zipCryptoCheckPassword(encrypted, password, crcValue);

    expect(result).toBe(true);
  });

  it("should return false for incorrect password in most cases", () => {
    // Note: ZipCrypto password verification only checks 1 byte,
    // so there's a 1/256 chance of false positive. We use a specific
    // password combination that is known to fail verification.
    const password = "correct-password";
    const wrongPassword = "different-password-xyz";
    const plaintext = new TextEncoder().encode("Test data for verification");
    const crcValue = crc32(plaintext);

    const encrypted = zipCryptoEncrypt(plaintext, password, crcValue, createDeterministicRandom());
    const result = zipCryptoCheckPassword(encrypted, wrongPassword, crcValue);

    // With deterministic random and these specific passwords, this should be false
    // If this test becomes flaky, adjust the passwords or seed
    expect(result).toBe(false);
  });

  it("should return false for too short data", () => {
    const shortData = new Uint8Array(5);
    const result = zipCryptoCheckPassword(shortData, "password", 0);

    expect(result).toBe(false);
  });
});
